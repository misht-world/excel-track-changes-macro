# Detailed Installation Guide

## Quick Setup

1. **Download the macro** from this repository
2. **Open your Excel file** where you want to track changes
3. **Enable Developer Tab** (if not already):
   - File → Options → Customize Ribbon
   - Check "Developer" in right panel

## Step-by-Step Installation

### Method 1: Direct Copy-Paste
1. Press `Alt + F11` to open VBA Editor
2. In Project Explorer, double-click `ThisWorkbook`
3. Delete any existing code
4. Copy contents from `TrackChangesMacro.vba`
5. Paste into the code window
6. Close VBA Editor (`Alt + Q`)
7. Save workbook as `.xlsm` format

### Method 2: Import Module
1. Open VBA Editor (`Alt + F11`)
2. Right-click `ThisWorkbook` → Insert → Module
3. Copy-paste the code
4. Save and clos

## Testing the Macro

1. Open the saved `.xlsm` file
2. Modify any cell value - it should turn blue
3. Change it back to original - it should turn black
4. Test with formulas and across different sheets

## Troubleshooting

**Macro not working?**
- Ensure file is saved as `.xlsm`
- Enable macros when opening file
- Check Trust Center settings

**Performance issues?**
- The macro tracks all used cells
- For very large workbooks, consider limiting range

**Colors not updating?**
- Manual formatting overrides macro colors
- Clear manual formatting if needed

## 📄 License & Disclaimer

This macro is provided as-is for educational and personal use purposes. The author is not responsible for any data loss, corruption, or other issues that may arise from using this macro. Always backup your files and test on non-critical data first.

**Use at your own risk.**

---

**Russian version: [installation_ru.md](installation_ru.md)**