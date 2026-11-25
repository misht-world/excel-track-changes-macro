# Excel Track Changes

An Excel add-in that highlights every cell whose value differs from the original. Undo (Ctrl+Z) keeps working, your own formatting is never touched, and the highlight disappears as soon as a value returns to the original.

## 🌟 Features

- **Undo works**: no code runs while you edit. Highlighting is done by Excel's conditional formatting.
- **Your formatting is safe**: the highlight is drawn on top of the cell's own colors and goes away when the value is restored.
- **Choose the style**: text color, cell fill, or both, in five colors.
- **All sheets**: every worksheet of the workbook is tracked, including cells that were empty.
- **Survives saving**: the original values are stored inside the workbook, so tracking continues after you close and reopen the file. The file can stay a regular `.xlsx`.
- **Accept changes**: accept everything or just the selected cells.
- **One-time install**: the add-in works with any workbook from its own ribbon tab.

## 🚀 Quick Start

1. Download `TrackChanges.xlam` from [Releases](https://github.com/misht-world/excel-track-changes-macro/releases).
2. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**, select the file and make sure **Track Changes** is checked.
3. Open any workbook and go to the **Track Changes** tab → **Start tracking**.

**📖 Full guide:** [docs/installation.md](docs/installation.md)

## 🎛️ Ribbon Buttons

| Button | What it does |
|---|---|
| **Start tracking** | Remembers the current values of all sheets and starts highlighting changes |
| **Accept all** | Current values become the new original, highlighting disappears |
| **Accept selection** | Accepts changes only in the selected cells |
| **Refresh** | Restores the highlighting after pasting cells with formatting and starts tracking newly added sheets |
| **Turn off** | Removes the highlighting and the stored original values. Your data and formatting stay as they are |
| **Style / Color** | How to highlight: text color, cell fill, or both |

## 📁 Files

- [`src/TrackChanges.bas`](src/TrackChanges.bas) - add-in code (VBA module)
- [`src/customUI14.xml`](src/customUI14.xml) - ribbon tab
- [`build/build.ps1`](build/build.ps1) - builds `TrackChanges.xlam` from the sources
- [`legacy/TrackChangesMacro_v1.vba`](legacy/TrackChangesMacro_v1.vba) - previous version (v1), kept for reference
- [`docs/installation.md`](docs/installation.md) - detailed guide

## ⚠️ Important Notes

- **Cells are compared by address.** After inserting, deleting or sorting rows/columns, the shifted cells are highlighted as changed. Do structural changes before starting, or press **Accept all** after them.
- **Values are compared, not formulas.** A new formula that gives the same result is not highlighted. Volatile formulas (`NOW()`, `RAND()`) are highlighted after they recalculate.
- **Pasting with formatting** can remove the highlight in the pasted cells. Press **Refresh** to restore it.
- **The add-in's buttons clear the Undo history once** when pressed (as any macro does). Normal editing is not affected.
- **File size**: the original values are stored on a hidden sheet, so the file grows roughly by the size of its data while tracking is on.
- **Upgrading from v1**: remove the old macro code from `ThisWorkbook` (Alt+F11), otherwise it keeps blocking Undo.

## 🔧 Compatibility

- Excel 2010 and later for Windows (Microsoft 365 included)
- Any workbook format: `.xlsx`, `.xlsm`, `.xlsb`

## ☕ Support the Project

If this macro has been helpful and you'd like to support its development:

**Cryptocurrency:**
- **Bitcoin:** `bc1qy5la5xsswx2hnukf76x5g2spjpvf6n7ak56rnu`
- **Ethereum:** `0x75F27F55Fa87d13C267c304E7e3a0afeB0f660a9`

**Traditional:**
- **PayPal:** [https://www.paypal.com/paypalme/MikhailZhi](https://www.paypal.com/paypalme/MikhailZhi)

*IBAN (EU) - Coming soon*


## 📄 License & Disclaimer

This macro is provided as-is for educational and personal use purposes. 

**⚠️ Disclaimer:** The author is not responsible for any data loss, corruption, or other issues that may arise from using this macro. Users should:

- Always backup important files before using macros
- Test the macro on non-critical data first
- Understand the code before implementation
- Use at your own risk

## 📞 Support

For issues and questions, please open an Issue in this repository.
