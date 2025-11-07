# Excel Track Changes Macro

A sophisticated VBA macro that visually tracks changes in Excel workbooks by coloring modified text in blue. The color automatically reverts to black when values return to their original state.

## 🌟 Features

- **Visual Change Tracking**: Modified cells are highlighted in blue
- **Multi-Sheet Support**: Tracks changes across all worksheets
- **Formula Awareness**: Detects changes in both values and formulas
- **Smart Revert**: Automatically reverts color when values return to original
- **Cross-Sheet References**: Properly handles linked cells between sheets
- **Initial State Capture**: Remembers the state when file was opened

## 🚀 Quick Start

1. Open Excel and press `Alt + F11` to open VBA Editor
2. Double-click `ThisWorkbook` in Project Explorer
3. Replace all code with content from `TrackChangesMacro.vba`
4. Save as `.xlsm` file (Excel Macro-Enabled Workbook)

## 📁 Files

- `TrackChangesMacro.vba` - Main macro code
- `docs/installation.md` - Detailed installation guide
- `docs/installation_ru.md` - Russian installation guide

## ⚠️ Usage Notes

- **Undo Limitations**: Ctrl+Z doesn't affect macro coloring
- **Structural Changes**: Insert/delete rows/columns outside the macro first
- **Performance**: Optimized for medium-sized workbooks
- **Persistence**: Changes are tracked until file is closed and reopened

## 🔧 Compatibility

- Excel 2010 and later
- Windows OS
- .xlsm files only

## 📞 Support

For issues and questions, please open an Issue in this repository.

---

**For Russian version, see [README_ru.md](README_ru.md)**