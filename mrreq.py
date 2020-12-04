import openpyxl as pxl

import sys, re
import json
from dataclasses import dataclass, field
from typing import List, Tuple

@dataclass
class Requirement:
    id: str = None
    title: str = None
    level: int = None
    subsystem: str = None
    status: List[str] = None
    priority: str = None
    parent: str = None
    description: str = None
    additional_parents: List[str] = None
    children: List[str] = field(default_factory=list)

@dataclass
class Subsystem:
    keyword: str
    level: int
    sheet_name: str
    sheet: pxl.worksheet.worksheet.Worksheet
    requirements: List[str] = None


class MRReqChecker:
    def __init__(self, filepath, row_start=3, col_start=1):
        self.workbook = pxl.load_workbook(filename=filepath)
        self.level_dict = {'LS': ['MR'],
                           'L0': ['OBJ'],
                           'L1': ['MIS'],
                           'L2': ['SYS', 'MOP'],
                           'L3': ['FAC', 'OPR', 'MCS', 'DPR', 'MEC', 'SDE', 'AVI', 'SOF', 'THR', 'POW']}
        self.sheet_names = self.workbook.sheetnames
        self.row_start = row_start
        self.col_start = col_start
        self.subsystem_dict = {}
        self.requirement_dict = {}

        # add a lead to all requirements
        self.lead_req = Requirement(id="MR", title="MoonRanger Requirement", level=-1)
        self.requirement_dict["MR"] = self.lead_req

        for sheet_name in self.sheet_names:
            for level in self.level_dict:
                for key in self.level_dict[level]:
                    if re.search(f"^{key}", sheet_name, re.IGNORECASE) is not None:
                        subsys = Subsystem(keyword=key,
                                           level=int(level[-1]),
                                           sheet_name=sheet_name,
                                           sheet=self.workbook[sheet_name])
                        reqs = self._get_requirements_per_subsystem(subsys)
                        subsys.requirements = reqs
                        self.subsystem_dict[key] = subsys


    def check_missing_parent(self):
        for req_id in self.requirement_dict:
            if self.requirement_dict[req_id].level != -1:
                parent_id = self.requirement_dict[req_id].parent
                if parent_id not in self.requirement_dict.keys():
                    print(f"<MissingParent> x----- <{req_id}>'s parent missing, current parent <{parent_id}>")
                    # add logging
                else:
                    self.requirement_dict[parent_id].children.append(req_id)


    def check_missing_additional_parent(self):
        for req_id in self.requirement_dict:
            if self.requirement_dict[req_id].level != -1:
                additional_parents = self.requirement_dict[req_id].additional_parents
                if additional_parents is not None:
                    for parent_id in additional_parents:
                        if parent_id not in self.requirement_dict.keys():
                            print(f"<MissingAdditionalParent> x----- <{req_id}>'s additional parent missing, current parent <{parent_id}>")
                            # add logging
                        else:
                            self.requirement_dict[parent_id].children.append(req_id)


    def check_status(self):
        pass


    def convert_to_tree(self):

        def _fetch_children(self, req_id):
            tmp_dict = {}
            req = self.requirement_dict[req_id]
            tmp_dict["name"] = f"[{req_id}] {req.title}"
            children = req.children
            if len(children) == 0:
                tmp_dict["name"] += f" | {req.description}"
                # return tmp_dict
            else:
                tmp_dict["children"] = []
                for child_id in children:
                    tmp_dict["children"].append(_fetch_children(self, child_id))
            return tmp_dict
        
        self.check_missing_parent()
        self.check_missing_additional_parent()

        return _fetch_children(self, self.lead_req.id)


    def cal_statitics(self, sheet_name, status="TBD"):
        pass

    def get_sheet_by_keyword(self, keyword="SYS"):
        for kw in self.subsystem_dict:
            subsys = self.subsystem_dict[kw]
            if (re.search(f"^{keyword}", kw, re.IGNORECASE) is not None) or \
               (re.search(f"^{keyword}", subsys.sheet_name, re.IGNORECASE) is not None):
                return subsys.sheet
        return None


    def _get_col_idx_by_keyword(self, sheet, keyword):
        col_dict = self._get_col_dict(sheet)
        for key in col_dict:
            if re.search(f"^{keyword}", key, re.IGNORECASE) is not None:
                return col_dict[key]
        return None


    def _get_requirements_per_subsystem(self, subsystem):
        sheet = subsystem.sheet
        requirements = []
        attrs = ['id', 'title', 'priority', 'description', 'status', 'parent', 'additional_parents']       # useful attributes to extract from the .xlsx file
        for row in sheet.iter_rows(min_row=self.row_start, max_row=sheet.max_row, values_only=True):
            if row[0] is not None and re.search("^[A-Z]{3}-", row[0], re.IGNORECASE) is not None:
                req = Requirement()
                for attr in attrs:
                    attr_idx = self._get_col_idx_by_keyword(sheet, attr)
                    if attr_idx is not None:
                        value = row[attr_idx]
                        value = self._clean_str(value)
                    else:
                        value = None

                    if attr == 'id' or attr == 'parent':
                        value = self._clean_str(value, handler="upper")
                        value = self._clean_id(value)

                    if attr == 'status':
                        if value is not None:
                            value = list(value.split(','))
                            value = [self._clean_str(v, handler="upper") for v in value]

                    if attr == 'additional_parents':
                        if value is not None:
                            value = list(value.split(','))
                            value = [self._clean_id(self._clean_str(v, handler="upper")) for v in value]

                    req.__setattr__(attr, value)

                req.level = subsystem.level
                req.subsystem = subsystem.keyword

                if req.level == 0:
                    req.parent = "MR"

                if "DELETED" not in req.status and "DELETE" not in req.status:
                    self.requirement_dict[req.id] = req
                    requirements.append(req.id)

        return requirements


    def _get_col_dict(self, sheet):
        col_dict = {}
        counter = 0
        for col in sheet.iter_cols(1, sheet.max_column, values_only=True):
            if col[1] is not None:
                col_dict[col[1]] = counter
                counter += 1
        return col_dict



    @staticmethod
    def _clean_str(text, handler=None):
        # clean leading / trailing space, new line
        if text is None:
            return None

        result = text.strip().strip('\n')

        # capitalize
        if handler is None:
            return result
        if handler == "upper":
            return result.upper()
        elif handler == "lower":
            return result.lower()
        elif handler == "capitalize":
            return result.capitalize()
        else:
            raise ValueError("Wrong handler type to clean string")


    @staticmethod
    def _get_status():
        pass


    @staticmethod
    def _clean_id(req_id):
        if req_id is None or "-" not in req_id:
            return req_id

        parts = req_id.split("-")
        numeric = int(parts[-1])
        return parts[0] + "-" + str(numeric)



if __name__ == "__main__":
    try:
        filepath = sys.argv[1]
    except:
        filepath = "MR-SYS-0001 MoonRanger Requirements.xlsx"

    rc = MRReqChecker(filepath)
    data = rc.convert_to_tree()
    with open("data.json", "w") as outfile:
        json.dump(data, outfile)

    print("JSON file generated. Please open index.html to preview visualization.")