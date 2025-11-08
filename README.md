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
3. Replace all code with content from [`TrackChangesMacro.vba`](TrackChangesMacro.vba)
4. Save as `.xlsm` file (Excel Macro-Enabled Workbook)

**📖 Full Installation Guide:** [English](docs/installation.md) | [Russian](docs/installation_ru.md)

## 📁 Files

- [`TrackChangesMacro.vba`](TrackChangesMacro.vba) - Main macro code
- [`docs/installation.md`](docs/installation.md) - Detailed installation guide (English)
- [`docs/installation_ru.md`](docs/installation_ru.md) - Russian installation guide

## ⚠️ Important Notes

- **Undo Function Disabled**: Ctrl+Z (undo) does not work while the macro is active. Workflow: make changes → save file.
- **Structural Changes Risk**: Inserting/deleting rows/columns while macro is running may cause file freezing/hanging. Perform structural changes without macro active.
- **Persistence**: Changes are tracked until file is closed and reopened.

## 🔧 Compatibility

- Excel 2010 and later
- Windows OS
- .xlsm files only

## ☕ Support the Project

If this macro has been helpful and you'd like to support its development:

**Cryptocurrency:**
- **Bitcoin:** `bc1qy5la5xsswx2hnukf76x5g2spjpvf6n7ak56rnu`
- **Ethereum:** `0x75F27F55Fa87d13C267c304E7e3a0afeB0f660a9`

**Traditional:**
- **PayPal:** [https://www.paypal.com/paypalme/MikhailZhi](https://www.paypal.com/paypalme/MikhailZhi)

*IBAN (EU) - Coming soon*

Thank you for your support! ❤️

## 📄 License & Disclaimer

This macro is provided as-is for educational and personal use purposes. 

**⚠️ Disclaimer:** The author is not responsible for any data loss, corruption, or other issues that may arise from using this macro. Users should:

- Always backup important files before using macros
- Test the macro on non-critical data first
- Understand the code before implementation
- Use at your own risk

## 📞 Support

For issues and questions, please open an Issue in this repository.

---

**For Russian version, see [README_ru.md](README_ru.md)**