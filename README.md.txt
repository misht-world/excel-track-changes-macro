# Excel Track Changes Macro

A sophisticated VBA macro that visually tracks changes in Excel workbooks by coloring modified text in blue. The color automatically reverts to black when values return to their original state.

## Features

- **Visual Change Tracking**: Modified cells are highlighted in blue
- **Multi-Sheet Support**: Tracks changes across all worksheets
- **Formula Awareness**: Detects changes in both values and formulas
- **Smart Revert**: Automatically reverts color when values return to original
- **Cross-Sheet References**: Properly handles linked cells between sheets
- **Initial State Capture**: Remembers the state when file was opened

## Installation

1. Open your Excel workbook
2. Press `Alt + F11` to open VBA Editor
3. Double-click `ThisWorkbook` in Project Explorer
4. Replace all code with the contents of `TrackChangesMacro.vba`
5. Save as `.xlsm` file (Excel Macro-Enabled Workbook)

## Usage Notes

- **Undo Limitations**: Ctrl+Z doesn't affect macro coloring
- **Structural Changes**: Insert/delete rows/columns outside the macro first
- **Performance**: Optimized for medium-sized workbooks
- **Persistence**: Changes are tracked until file is closed and reopened

## File Structure

- `TrackChangesMacro.vba` - Main macro code
- `docs/installation.md` - Detailed installation instructions
- `examples/` - Sample workbooks demonstrating the macro

## Compatibility

- Excel 2010 and later
- Windows OS
- .xlsm files only