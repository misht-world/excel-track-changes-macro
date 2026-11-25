"""Generates demo/TrackChanges_Grid.xlsx: two sheets full of numbers for a visual demo.

  * Model  - a grid of 3-digit numbers. A1 (yellow) is the driver.
  * Report - a grid of the same size.
Exactly half of the cells on each sheet are formulas, scattered at random and
formatted like the plain numbers around them. Every formula depends on Model!A1:
on Model directly or through other formulas, on Report through formula cells
of Model. Change A1 and half of both sheets lights up.

Usage:  python demo/make_grid.py        (requires openpyxl)
"""
import os
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'TrackChanges_Grid.xlsx')

ROWS, COLS = 40, 30          # larger than a 1920x1080 screen at 100% zoom
DRIVER = 100

FONT = Font(name='Calibri', size=11, color='2B2F36')
DRIVER_FONT = Font(name='Calibri', size=11, bold=True, color='2B2F36')
DRIVER_FILL = PatternFill('solid', fgColor='FFE699')
CENTER = Alignment(horizontal='center', vertical='center')


def addr(r, c, absolute=False):
    col = get_column_letter(c)
    return f'${col}${r}' if absolute else f'{col}{r}'


def formula_cells(rnd, exclude=()):
    cells = [(r, c) for r in range(1, ROWS + 1) for c in range(1, COLS + 1) if (r, c) not in exclude]
    total = ROWS * COLS
    picked = rnd.sample(cells, total // 2)
    picked.sort()                     # row-major: references always point to earlier cells
    return picked


def build():
    rnd = random.Random(7)
    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    model = wb.active
    model.title = 'Model'
    report = wb.create_sheet('Report')

    for ws in (model, report):
        ws.sheet_view.zoomScale = 100
        for c in range(1, COLS + 1):
            ws.column_dimensions[get_column_letter(c)].width = 8.43

    # ---------------- Model
    m_formulas = formula_cells(rnd, exclude={(1, 1)})
    m_set = set(m_formulas)
    done = []                         # formula cells already written (all depend on A1)
    for r in range(1, ROWS + 1):
        for c in range(1, COLS + 1):
            cell = model.cell(r, c)
            cell.font, cell.alignment, cell.number_format = FONT, CENTER, '0'
            if (r, c) == (1, 1):
                cell.value = DRIVER
                cell.font, cell.fill = DRIVER_FONT, DRIVER_FILL
            elif (r, c) in m_set:
                k = round(rnd.uniform(1.5, 8.5), 2)
                b = rnd.randint(40, 160)
                if done and rnd.random() < 0.6:
                    src = addr(*rnd.choice(done[-60:]))          # chain through another formula
                    k2 = round(rnd.uniform(0.6, 1.4), 2)
                    cell.value = f'=ROUND(MOD({src}*{k2}+$A$1*{k}+{b},900)+100,0)'
                else:
                    cell.value = f'=ROUND(MOD($A$1*{k}+{b},900)+100,0)'
                done.append((r, c))
            else:
                cell.value = rnd.randint(100, 999)

    # ---------------- Report
    r_formulas = set(formula_cells(rnd))
    for r in range(1, ROWS + 1):
        for c in range(1, COLS + 1):
            cell = report.cell(r, c)
            cell.font, cell.alignment, cell.number_format = FONT, CENTER, '0'
            if (r, c) in r_formulas:
                src = addr(*rnd.choice(m_formulas))
                k = round(rnd.uniform(0.7, 1.6), 2)
                b = rnd.randint(10, 300)
                cell.value = f'=ROUND(MOD(Model!{src}*{k}+{b},900)+100,0)'
            else:
                cell.value = rnd.randint(100, 999)

    model.sheet_view.selection[0].activeCell = 'A1'
    model.sheet_view.selection[0].sqref = 'A1'
    wb.save(OUT)
    return OUT


if __name__ == '__main__':
    print(build())
