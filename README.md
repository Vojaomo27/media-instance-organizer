# Media Instance Organizer

A Windows desktop utility for matching image and video filenames against an instance-ID list, previewing the planned organization, and copying files into `WORKED` or `NOT_WORKED` folders. Original files remain untouched.

## Features

- Exact ID matching, including common numbered duplicate suffixes such as `_1`.
- TXT, CSV, XLSX, and XLSM ID lists.
- Preview scan before copying; existing destination files are skipped.
- Optional subfolder scanning and configurable media extensions.
- Timestamped CSV and summary logs, plus locally saved settings.

## Run from source

Python 3.10 or newer is recommended. Tkinter is included with standard Windows Python installations.

```powershell
python -m pip install -r requirements.txt
python move.py
```

Excel list support uses `openpyxl`. TXT and CSV lists do not require it.

## Build a Windows executable

```powershell
.\build_exe.bat
```

The executable is created in `dist\MediaInstanceOrganizer.exe`. Build outputs and local data are excluded from the repository.
