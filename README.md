# Outlook: Move Selected Emails to Year Folders

A VBA macro for **Microsoft Outlook Classic (Windows)** that moves every selected email into a subfolder named for the year it was sent or received (`2025`, `2026`, …). Year folders are created automatically inside the folder the emails are currently in.

Select 100 emails from 2025 and 2026, run the macro, and you get:

```
Inbox
├── 2025   ← emails from 2025
└── 2026   ← emails from 2026
```

## Features

- Creates year subfolders only when needed and reuses existing ones
- Asks for confirmation first and shows a per-year summary at the end
- Uses the received date, then falls back to the sent date, then to the creation date
- Skips emails already in their year folder, so running it twice won't create `2025\2025`
- Works from search results, moving each email into a year folder inside the folder it actually lives in

## Install

1. In Outlook, press **Alt+F11** to open the VBA editor.
2. Go to **File → Import File…** and choose `MoveToYearFolders.bas`.
3. If macros are blocked, go to **File → Options → Trust Center → Trust Center Settings → Macro Settings** and choose *Notifications for all macros*, then restart Outlook.

If you paste the code into a module instead of importing it, leave out the first line (`Attribute VB_Name = ...`).

## Use

Select the emails, press **Alt+F8**, and run `MoveSelectedToYearFolders`.

For one-click access, add it to the Quick Access Toolbar: **File → Options → Quick Access Toolbar → Choose commands from: Macros**.
