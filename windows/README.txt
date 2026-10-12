MLC CLAIMANT DROP — WINDOWS

INSTALL ONCE
1. Extract this ZIP file.
2. Double-click "Install MLC Claimant Drop.cmd".
3. A shortcut named "MLC Claimant Drop" will appear on your desktop.

USE IT
1. Drag one MLC .xlsx report onto the desktop shortcut.
2. Wait for the confirmation box: True Full or True Partial (or Needs review for unrecognized answers).
3. The deduplicated table is already copied. Press Ctrl+V wherever you want it.
4. A CSV copy is also saved beside the original report.

FILTER RULES
- Column E accepts YES, Yes, Y, and (Y), ignoring capitalization and spaces.
- All Yes = True Full; any No (including N and (N)) = True Partial; blanks are ignored.
- One report represents one transfer application and is expected to contain at least one Yes.
- Clipboard and CSV headers: Admin, Admin P#, Publisher, Publisher P#.
- Source columns: R, Q, N, M. Member fields are excluded.
- Duplicate four-column claimant rows are removed.
- Existing Send To shortcuts can continue to call the launcher with the report path.

REQUIREMENTS
- Windows with desktop Excel installed.
- If your employer blocks PowerShell scripts, ask IT to approve the extracted files.
