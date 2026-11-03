# Workbench tools (version 1.1)

**Double-click `Setup.bat`.** It opens the EA Workbench Setup window, which does everything below without typing
commands. `Setup (menu).bat` gives the same choices as a numbered menu in a console window.

| Step in the window | What it does |
| --- | --- |
| 1. Check this computer | Python version, the openpyxl and pytest packages, Excel, draw.io; Install Python packages fetches openpyxl and pytest |
| 2. Create a new model | Builds `EA_Model_<Name>` with the ArchiMate 4 metamodel (optionally the sample data), the Power BI project, the draw.io libraries, the code, guides and test sheets |
| 3. Install a model into Excel | Unblocks downloaded files, then opens Excel and runs the installer for you (Excel must allow access to the VBA project while it runs) |
| 4. Maintain a model | Rebuilds the Power BI project after a model folder has moved |
| Run self-tests | Builds a fresh sample model in a temporary folder and runs the 166 checks |

Needs Python 3.9 or later from python.org (tick "Add python.exe to PATH" when installing). Only the person who
creates models needs Python; people who use a model need only Excel.

The scripts behind the window can also be run directly:

- `builder.py <new model folder> [--no-sample]` builds EA_Core.xlsx, EA_Entities.xlsx, model.json and the draw.io files
- `pbip.py <model folder> "<Windows path of the folder>\"` writes the Power BI project into Reports
- `drawio.py` shape styles, library format and diagram layout (used by builder.py; reference for modDiagram)
- `seed.py` and `model.py` the ArchiMate 4 seed and the reference logic the VBA mirrors
- `test_workbench.py` the checks: `python -m pytest -q test_workbench.py` (tests EA_Model_Sample next to this folder, or the folder in the EAWB_MODEL environment variable)
