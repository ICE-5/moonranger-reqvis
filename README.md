# MoonRanger Requirement Visualization

This repository contains the script for automatic checking & visualizing `MR-SYS-0001 MoonRanger Requirements.xlsx`.

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
  * If you want to use other path for the requirement `.xlsx` file, then input the path as first argument.
    ```bash
    python mrreq.py <path-to-xlsx>
    ```
* The script will automatically check for the following errors and print out info.
  * If a requirement's parent is missing or deleted.
  * If a requirement's additional parent is missing or deleted.

* Step 3, if no exception is raised, and `JSON file generated. Please open index.html to preview visualization.` is prompted, then it means `data.json` for visualization is successfully generated and everything is set . Run a server and click on `index.html` to see the visualization. For example, with `Python 3` you can do, 
  * Run command
    ```bash
    python3 -m http.server
    ```
  * Open `http://localhost:8000/` in browser, and the visualization will be there, viola!
