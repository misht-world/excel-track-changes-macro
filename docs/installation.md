# Detailed Guide

## Installing the Add-in

1. Download `TrackChanges.xlam` from [Releases](https://github.com/misht-world/excel-track-changes-macro/releases) and put it in a permanent folder (for example `Documents\Excel`). If you move it later, Excel will not find it.
2. Windows may block downloaded files. Right-click the file → **Properties** → check **Unblock** → **OK**.
3. In Excel: **File → Options → Add-ins**.
4. At the bottom, in **Manage**, choose **Excel Add-ins** → **Go…**
5. Click **Browse…**, select `TrackChanges.xlam`, make sure **Track Changes** is checked and click **OK**.
6. A **Track Changes** tab appears on the ribbon.

The add-in loads every time Excel starts. To disable it, uncheck it in the same dialog.

## How to Use

1. Open a workbook and click **Track Changes → Start tracking**. The current values of all sheets are remembered as the original.
2. Work as usual. Every cell whose value differs from the original is highlighted. If you bring the old value back (by typing or with Ctrl+Z), the highlight disappears.
3. Save the workbook. The original values are stored inside it, so the highlighting is still there after you reopen it.
4. When you have reviewed the changes:
   - **Accept all**: the current values become the original;
   - **Accept selection**: the same, only for the selected cells;
   - **Turn off**: remove the highlighting and the stored original values.

### Highlight Style

In the **Highlight** group choose:
- **Style**: text color, cell fill, or text and fill;
- **Color**: blue, red, green, orange or purple.

The choice is remembered and applied to the open workbook right away.

### The Refresh Button

Press it when:
- cells pasted from the clipboard (with formatting) lost their highlight;
- you added a new sheet: it starts being tracked, with its current values taken as the original;
- you unprotected a sheet that was skipped.

## How It Works

- On start, the values of each sheet are copied to a very hidden sheet `_tc_1`, `_tc_2`, … It does not show up in **Unhide**.
- Each sheet gets a hidden name `_TC_Diff` with the comparison formula and one conditional formatting rule `=_TC_Diff`.
- No code runs while you edit. That is why Ctrl+Z works, the workbook does not slow down and your formatting is not changed: the highlight is drawn on top of it.
- The comparison is type- and case-aware: a blank cell, `0` and an empty string are all different, and so are `abc` and `ABC`.

You can see the rule in **Home → Conditional Formatting → Manage Rules**. Do not delete it by hand; use **Turn off** instead.

## Limitations

- **Cells are compared by address.** After inserting, deleting or sorting rows/columns, the shifted cells are highlighted as changed. Do such changes before starting, or press **Accept all** afterwards.
- **Values are compared, not formulas.** Replacing a formula with another one that gives the same result is not highlighted.
- **Volatile formulas** (`NOW()`, `RAND()`, …) are highlighted after they recalculate, because their value changes.
- **Pasting with formatting** can remove the highlight in the pasted cells. Press **Refresh**.
- **Copying to another workbook.** Copying cells of a tracked workbook to another workbook together with formatting may carry the highlight rule along with a link to the source workbook. Use **Paste Special → Values**, or turn tracking off before copying.
- **The add-in's buttons clear the Undo history once** (as any macro does). Normal editing is not affected.
- **Protected sheets** are skipped. Unprotect them and press **Refresh**.
- **File size** grows roughly by the size of the data while tracking is on.

## Without the Add-in

If add-ins are blocked by company policy, you can add the module to your Personal Macro Workbook:

1. Open the VBA editor (`Alt + F11`).
2. Select the `PERSONAL.XLSB` project. If there is none, record any empty macro and store it in "Personal Macro Workbook".
3. **File → Import File…** and choose [`src/TrackChanges.bas`](../src/TrackChanges.bas). If you copied the code from GitHub as text, create a module instead (**Insert → Module**) and paste the code without the first line `Attribute VB_Name = ...`.
4. Run the commands with `Alt + F8` by typing their names: `TC_Start`, `TC_AcceptAll`, `TC_AcceptSelection`, `TC_Refresh`, `TC_Stop`. There is no ribbon tab in this setup.

## Building the Add-in from Source

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```

The script imports the module through Excel, saves `dist\TrackChanges.xlam` and adds the ribbon tab to it. While building, turn on **File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model**. Turn it off again afterwards.

## Upgrading from Version 1

Remove the old macro from `ThisWorkbook`, otherwise it keeps recoloring cells and blocking Undo:

1. Open the workbook and press `Alt + F11`.
2. Double-click `ThisWorkbook` and delete all the code.
3. If the workbook has no other macros, you can save it as a regular `.xlsx`.
4. Font colors set by the old macro (blue or black) stay. Remove them by hand if needed.

## Troubleshooting

**There is no Track Changes tab.**
Make sure the add-in is checked in **File → Options → Add-ins → Excel Add-ins**. If Excel says macros are disabled, unblock the file (step 2 of the installation) or put it in a trusted location.

**Cells are not highlighted.**
Press **Refresh**. Make sure the sheet is not protected and that no version 1 code is left in the workbook.

**Cells I did not change are highlighted.**
Most likely rows or columns were inserted or deleted, or the data was sorted. Press **Accept all** if the current state is correct.

## 📄 License & Disclaimer

This add-in is provided as-is for educational and personal use. The author is not responsible for any data loss, file corruption or other issues that may arise from using it. Always back up your files and test on non-critical data first.

**Use at your own risk.**

