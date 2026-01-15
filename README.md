# Doc Merger

## Overview
Doc Merger is a Swing desktop application for scanning a directory of `.doc` / `.docx` files and merging them in the exact order shown in the UI. The list supports drag-and-drop reordering, and custom order is persisted between app restarts.

## Drag-and-drop ordering
- Drag rows inside the file table to reorder.
- The merge operation always uses the current table order.
- External drops are ignored.

## Custom order persistence
- The current table order is saved as a list of stable IDs (normalized absolute file paths).
- On restart, the app restores custom order when the same source directory is used.
- Custom order is stored with Java Preferences per user.

### Alignment rules on rescan
When a rescan occurs (or the source directory changes):
1. The app rebuilds the default list using natural sort.
2. If a stored custom order exists for the same directory, the app aligns:
   - Items found in the stored order are restored first.
   - Missing files are ignored (counted in the log).
   - New files not in the stored order are appended using natural sort.
3. The log records whether custom order was applied and how many files were missing/new.

## Reset behavior
Click **Reset to Default Order** to return to natural sort immediately. This clears the persisted custom order so future restarts stay in default mode.

## Acceptance tests (manual)
1. Default natural sort order is correct for: `1xs.docx`, `2f23r23.doc`, `10a.docx`.
2. Drag the 3rd item to position 1, click **Start Merge**:
   - Output file order matches the dragged UI order.
3. Restart the app with the same source directory:
   - Custom order is restored.
4. Delete a source file then rescan:
   - Custom order partially restores; missing files are ignored; new files append; logs explain.
5. Click **Reset to Default Order**:
   - List returns to natural sort immediately.
   - Restart preserves default (custom order cleared).

## Build
```bash
mvn -q -DskipTests package
```
