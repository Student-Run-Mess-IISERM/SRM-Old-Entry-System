# SRM Entry System

SRM Entry System using Custom Tkinter made by Ratnojit Bhattacharya and maintained by Kshitish Kumar Ratha and Gokul P Bharathan.

# TODO
- [ ] Multi-threading for the GUI to prevent freezing.
- [ ] Refactor code into multiple files.
- [ ] Brainstorm a way to implement credentials/config in a better way.
- [ ] Mention the mess number while putting in the entry.

# Changelog

## Current Version

### 1.3.3
- Better colors for meal map.
- Better contrast with text and background for details box.

## Older Versions

### 1.3.2
- Correct key for color map is not used to display student details.

### 1.3.1
- Disable online features when credentials are not provided.

### 1.3.0
- Implement dynamic details box with color-coded meal status, which displays student information and is color coded according to meal status.

### 1.2.2
- Increased indexing to account for `1` based indexing in `openpyxl` and `gspread`'s `update_cell` methods/properties.

### 1.2.1
- Fixed the column numbers in `MEAL_COLUMN_MAPPING` to be `0` based again.

### 1.2.0
- Reworked the logic for opening daily files. Implemented caching for the same. Check `workbook` and `gsheet` function.
- Fixed old broken code that was preventing daily entry.
- Removed Type Hints as they were causing issues.
- Changed formatting of constants to be SCREAMING_SNAKE_CASE.
- Renamed some variables for better clarity.
- Other minor refactoring.

### 1.1.0
Added type hints everywhere which hopefully fixes issues.

### 1.0.0
Working System with Calculations.