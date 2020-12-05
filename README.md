# MoonRanger Automatated Requirement Analysis & Visualization Pipeline

[[Link to visualization webpage](https://ice-5.github.io/moonranger-reqvis/)]

This repository contains the script for automatic checking & visualizing `MR-SYS-0001 MoonRanger Requirements.xlsx`.

> **Important**: the original requirement sheet (in `.xlsx`) shall not be uploaded to this public repo. Please locally run the script and only commit the newly generated `data.json` for visualization.

**Content**
- [MoonRanger Requirement Visualization](#moonranger-requirement-visualization)
  - [Dependencies](#dependencies)
  - [Usage](#usage)
  - [Guideline for maintaining the pipeline](#guideline-for-maintaining-the-pipeline)

## Dependencies

```
python=3.8.0
openpyxl=3.0.5
```

## Usage
* Step 1, download `MR-SYS-0001 MoonRanger Requirements.xlsx` from Google Sheets, and put / replace it into the folder.
* Step 2, run `mrreq.py`
  * If `MR-SYS-0001 MoonRanger Requirements.xlsx` is in the same directory as `mrreq.py`, just run
    ```bash
    python mrreq.py
    ```
  * If you want to use other path for the requirement sheet, then input the path as first argument.
    ```bash
    python mrreq.py <path-of-xlsx>
    ```
* The script will automatically check for the following errors and print out info.
  * If a requirement's parent is missing or deleted.
  * If a requirement's additional parent is missing or deleted.

* Step 3, if no exception is raised, and `JSON file generated. Please open index.html to preview visualization.` is prompted, then it means `data.json` for visualization is successfully generated and everything is set . Run a server to see the visualization. For example, with `Python 3` you can do, 
  * Run command
    ```bash
    python3 -m http.server
    ```
  * Open `http://localhost:8000/` in browser, and the visualization will be there, viola!


## Guideline for maintaining the pipeline
Since the script depends on structure of the requirement sheet, there are certain rules to follow in order to maintain the pipeline.

1. If a tab contains requirements, make sure to name it with a three-letter keyword that can be recognized by the script. If there is a new keyword, please remember to update the dictionary in the script. Currently, the recognizable keywords are,
  ```
  'L0': ['OBJ'],
  'L1': ['MIS'],
  'L2': ['SYS', 'MOP'],
  'L3': ['FAC', 'OPR', 'MCS', 'DPR', 'MEC', 'SDE', 'AVI', 'SOF', 'THR', 'POW']
  ```

2. If a tab contains requirements, its first row should be title of the tab, its second row **must** be column names. And requirements shall start at the third row.
3. **Avoid** merge cells. Remember one could always choose `overflow` in `text-wrap` options to ensure a conplete display of text.
4. For control of requirement status, here are the list of flags. Upper/lower case doesn't matter, as long as the spelling is correct and words are concatenated. 
```
Deleted
MissingParent
MissingAdditionalParent
TBD
TBR
MissingValue
```
5. Its perfectly fine to create other new flags, as long as they are consistent everywhere in the sheet. Please add new checking function to the `MRReqCheck` class if you see fit.
