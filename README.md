# MoonRanger Automated Requirement Analysis & Visualization Pipeline

[[Link to visualization webpage](https://ice-5.github.io/moonranger-reqvis/)]

This repository contains the script for automatic checking & visualizing `MR-SYS-0001 MoonRanger Requirements.xlsx`.

> **Important**: the original requirement sheet (in `.xlsx`) shall not be uploaded to this public repo. Please locally run the script and only commit the newly generated `data.json` for visualization.

**Content**
- [MoonRanger Automated Requirement Analysis & Visualization Pipeline](#moonranger-automated-requirement-analysis--visualization-pipeline)
  - [Dependencies](#dependencies)
  - [Usage](#usage)
  - [Guideline for maintaining the pipeline](#guideline-for-maintaining-the-pipeline)

## Dependencies

```
python=3.8.0
openpyxl=3.0.5
coloredlogs=14.0
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
* The script will automatically check for the following issues and log messages.
  * **`ERROR`** if a requirement's parent is missing or deleted.
  * **`ERROR`** if a requirement's additional parent is missing or deleted.
  * **`CRITICAL`** if there is `TBD` / `TBR` / missing values such as `XYZ` in the requirement, but there is **NOT** such flag in status column.
  * **`WARNING`** if there is **NOT** `TBD` / `TBR` / missing values such as `XYZ` in the requirement, but there is such flag in status column.

* Step 3, if `data.json for visualization successfully generated, ready to view!` is prompted, it means visualization is all set . Run a server to see it. For example, with `Python 3` you can do, 
  * Run command
    ```bash
    python3 -m http.server
    ```
  * Open `http://localhost:8000/` in browser, and the visualization will be there, viola!

* Step 4, if `statistics.json for calculating TBD/TBR/XYZ successfully generated, ready to view!` is prompted, it means statistics are calculated. Open `statistics.json` for some numbers!


## Guideline for maintaining the pipeline
Since the script depends on structure of the requirement sheet (`.xlsx` file), there are certain rules to follow in order to maintain the pipeline.

1. If a tab (or sheet) contains requirements, then it is associated with a subsystem. Each subsystem is identifiable via a three-letter key. Make sure to name the sheet with something that starts with this key (not case-sensitive). Below shows a hierarchy of currently identifiable keys to the script. For example, `Objective` is a successful naming of a tab, because its first three letters conform with key `OBJ`.
  ```
  'L0': ['OBJ'],
  'L1': ['MIS'],
  'L2': ['SYS', 'MOP'],
  'L3': ['FAC', 'OPR', 'MCS', 'DPR', 'MEC', 'SDE', 'AVI', 'SOF', 'THR', 'POW']
  ```

2. If a tab contains requirements, its first row should be title of the tab, its second row **must** be column names. And requirements **must** start from the third row.
3. **Avoid** merging cells. Remember one could always choose `overflow` in `text-wrap` options to ensure a complete display of text.
4. For control of requirement status, here are the list of flags (not case-sensitive) currently in use. Make sure the spelling is correct and words are concatenated. For example, `missing value` is a **wrong** flag, the correct flag can be `MissingValue`, `missingvalue`, `MISSINGVALUE`, etc.
```
Deleted
MissingParent
MissingAdditionalParent
TBD
TBR
MissingValue
```
5. It's perfectly fine to create other new flags, as long as they are consistent everywhere. 