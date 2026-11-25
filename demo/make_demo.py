"""Generates demo/TrackChanges_Demo.xlsx, the workbook used to show Track Changes in action.

Key technical indicators of a fictional residential project, "Nordic Park Residences":
  * Summary    - design parameters (yellow cells), key indicators, breakdowns and a chart
  * Apartments - 192 apartments in 3 buildings with a room-by-room area schedule
  * Floors     - floor-by-floor area schedule: net internal area, gross floor area, volume

All figures are areas (m²) and volumes (m³). Changing one parameter on the
Summary sheet (e.g. the balcony area factor) recalculates hundreds of cells
across the workbook, which makes the highlighting easy to see.

Usage:  python demo/make_demo.py        (requires openpyxl)
"""
import os
import random

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'TrackChanges_Demo.xlsx')

BUILDINGS = 3
RES_FLOORS = range(2, 18)      # residential floors 2..17
TYPES = ['Studio', '1-bed', '2-bed', '3-bed']
ROOMS = [1, 1, 2, 3]
ORIENT = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW']
FOOTPRINT = [378.0, 392.0, 365.0]   # per building, m²
PLOT_AREA = 8_400

M2 = '#,##0.0" m²"'
M2_INT = '#,##0" m²"'
M3 = '#,##0" m³"'
AREA = '#,##0.0'
INT = '#,##0'
PCT = '0.0%'
INT_NZ = '#,##0;-#,##0;""'      # hide zeros
AREA_NZ = '#,##0.0;-#,##0.0;""'

# ---------------------------------------------------------------- styles
INK, MUTED, LINE = '2B2F36', '7A808A', 'D9DCE1'
HEAD_FILL = PatternFill('solid', fgColor='2F3640')
ZEBRA = PatternFill('solid', fgColor='F5F6F8')
SECTION_FILL = PatternFill('solid', fgColor='ECE6DD')
TOTAL_FILL = PatternFill('solid', fgColor='E4E7EB')
INPUT_FILL = PatternFill('solid', fgColor='FFF2CC')

FONT = 'Segoe UI'
f_title = Font(name=FONT, size=20, bold=True, color=INK)
f_sub = Font(name=FONT, size=10, color=MUTED)
f_h2 = Font(name=FONT, size=12, bold=True, color=INK)
f_head = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_body = Font(name=FONT, size=10, color=INK)
f_bold = Font(name=FONT, size=10, bold=True, color=INK)
f_hint = Font(name=FONT, size=9, italic=True, color='B7791F')
f_kpi = Font(name=FONT, size=14, bold=True, color=INK)

thin = Side(style='thin', color=LINE)
B_BOTTOM = Border(bottom=thin)
B_TOP = Border(top=Side(style='medium', color=INK))
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center', indent=1)
RIGHT = Alignment(horizontal='right', vertical='center')
HEAD = Alignment(horizontal='center', vertical='center', wrap_text=True)


def put(ws, row, col, value=None, font=f_body, fill=None, fmt=None, align=None, border=None):
    c = ws.cell(row, col)
    if value is not None:
        c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def sheet_header(ws, title, subtitle, widths, zoom=110):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.row_dimensions[1].height = 34
    put(ws, 1, 2, title, f_title, align=Alignment(vertical='center'))
    put(ws, 2, 2, subtitle, f_sub)
    ws.column_dimensions['A'].width = 2.5
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def table_head(ws, row, start_col, names, height=32):
    ws.row_dimensions[row].height = height
    for i, name in enumerate(names):
        put(ws, row, start_col + i, name, f_head, HEAD_FILL, align=HEAD)


def build():
    rnd = random.Random(2026)
    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ws_s = wb.active
    ws_s.title = 'Summary'
    ws_a = wb.create_sheet('Apartments')
    ws_f = wb.create_sheet('Floors')

    # Summary parameter cells (referenced by the other sheets)
    P_BALCONY = 'Summary!$D$6'
    P_GROSS = 'Summary!$D$7'
    P_HEIGHT = 'Summary!$D$8'
    P_PLOT = 'Summary!$D$9'

    n_apts = BUILDINGS * len(RES_FLOORS) * 4

    # ================================================================ Apartments
    sheet_header(ws_a, 'Apartment schedule',
                 'Total area = living + kitchen + bath & hall + balcony × balcony factor (see Summary)',
                 {'B': 7, 'C': 13, 'D': 8, 'E': 10, 'F': 10, 'G': 8, 'H': 11, 'I': 11, 'J': 12,
                  'K': 11, 'L': 13, 'M': 13, 'N': 13})
    hr = 4
    table_head(ws_a, hr, 2, ['#', 'Building', 'Floor', 'Apt', 'Type', 'Rooms', 'Living, m²', 'Kitchen, m²',
                             'Bath & hall, m²', 'Balcony, m²', 'Apartment area, m²', 'Total area, m²',
                             'Orientation'])
    layouts = [[0, 1, 2, 3], [1, 2, 2, 3], [0, 1, 1, 2]]
    living_rng = {0: (16.0, 19.5), 1: (16.5, 20.0), 2: (28.0, 34.0), 3: (42.0, 50.0)}
    kitchen_rng = {1: (9.0, 11.5), 2: (10.5, 13.5), 3: (12.0, 16.0)}
    aux_rng = {0: (6.0, 8.0), 1: (8.5, 11.0), 2: (12.0, 15.5), 3: (16.0, 21.0)}

    first = r = hr + 1
    n = 0
    for b in range(1, BUILDINGS + 1):
        apt_no = 1
        for floor in RES_FLOORS:
            for pos in range(4):
                n += 1
                typ = layouts[b - 1][pos]
                living = round(rnd.uniform(*living_rng[typ]), 1)
                kitchen = round(rnd.uniform(*kitchen_rng[typ]), 1) if typ else None
                aux = round(rnd.uniform(*aux_rng[typ]), 1)
                balcony = round(rnd.uniform(3.0, 7.5), 1)
                orient = ORIENT[(pos * 2 + b) % 8]

                row = [n, f'Building {b}', floor, f'{b}-{apt_no:03d}', TYPES[typ], ROOMS[typ],
                       living, kitchen, aux, balcony]
                for i, v in enumerate(row):
                    put(ws_a, r, 2 + i, v,
                        align=CENTER if i in (0, 2, 3, 5) else LEFT if i in (1, 4) else RIGHT,
                        fmt=AREA if i >= 6 else None)
                put(ws_a, r, 12, f'=H{r}+I{r}+J{r}', fmt=AREA, align=RIGHT)
                put(ws_a, r, 13, f'=ROUND(L{r}+K{r}*{P_BALCONY},1)', f_bold, fmt=AREA, align=RIGHT)
                put(ws_a, r, 14, orient, align=CENTER)
                for col in range(2, 15):
                    c = ws_a.cell(r, col)
                    c.border = B_BOTTOM
                    if (r - first) % 2:
                        c.fill = ZEBRA
                apt_no += 1
                r += 1
    last = r - 1
    tot = r
    for col in range(2, 15):
        put(ws_a, tot, col, font=f_bold, fill=TOTAL_FILL, border=B_TOP)
    put(ws_a, tot, 2, 'Total', f_bold, TOTAL_FILL, align=LEFT, border=B_TOP)
    put(ws_a, tot, 7, f'=SUBTOTAL(109,G{first}:G{last})', f_bold, TOTAL_FILL, INT, CENTER, B_TOP)
    for col in 'HIJKLM':
        put(ws_a, tot, ord(col) - 64, f'=SUBTOTAL(109,{col}{first}:{col}{last})', f_bold, TOTAL_FILL, AREA, RIGHT, B_TOP)
    ws_a.freeze_panes = ws_a.cell(hr + 1, 3)
    ws_a.auto_filter.ref = f'B{hr}:N{last}'

    def apt(col):
        return f'Apartments!${col}${first}:${col}${last}'

    # ================================================================ Floors
    sheet_header(ws_f, 'Floor area schedule',
                 'Gross floor area = net internal area × gross-to-net factor; volume = GFA × floor-to-floor height',
                 {'B': 13, 'C': 8, 'D': 20, 'E': 10, 'F': 14, 'G': 15, 'H': 12, 'I': 11, 'J': 11,
                  'K': 11, 'L': 13, 'M': 13, 'N': 13})
    table_head(ws_f, 4, 2, ['Building', 'Floor', 'Use', 'Units', 'Residential, m²', 'Commercial, m²',
                            'Corridors & lobbies, m²', 'Stairs & lifts, m²', 'Technical, m²', 'Parking, m²',
                            'Net internal area, m²', 'Gross floor area, m²', 'Volume, m³'], 44)
    r = 5
    f_first = r
    for b in range(1, BUILDINGS + 1):
        name = f'Building {b}'
        fp = FOOTPRINT[b - 1]
        # section row
        for col in range(2, 15):
            put(ws_f, r, col, fill=SECTION_FILL, font=f_bold)
        put(ws_f, r, 2, name, f_bold, SECTION_FILL, align=LEFT)
        put(ws_f, r, 4, f'Footprint {fp:,.0f} m²'.replace(',', ' '), f_sub, SECTION_FILL, align=LEFT)
        r += 1
        s_first = r
        floors = [(-1, 'Parking & plant')] + [(1, 'Lobby & retail')] + [(f, 'Residential') for f in RES_FLOORS]
        for floor, use in floors:
            stairs = round(rnd.uniform(22, 27), 1)
            if floor == -1:
                vals = dict(com=None, cor=round(rnd.uniform(28, 36), 1), tech=round(rnd.uniform(55, 70), 1),
                            park=round(fp * rnd.uniform(1.55, 1.7), 1))
            elif floor == 1:
                vals = dict(com=round(fp * rnd.uniform(0.5, 0.56), 1), cor=round(rnd.uniform(42, 52), 1),
                            tech=round(rnd.uniform(14, 20), 1), park=None)
            else:
                vals = dict(com=None, cor=round(rnd.uniform(17, 23), 1), tech=round(rnd.uniform(2, 4), 1), park=None)
            put(ws_f, r, 2, name, align=LEFT)
            put(ws_f, r, 3, 'B1' if floor == -1 else floor, align=CENTER)
            put(ws_f, r, 4, use, align=LEFT)
            put(ws_f, r, 5, f'=COUNTIFS({apt("C")},B{r},{apt("D")},C{r})', fmt=INT_NZ, align=CENTER)
            put(ws_f, r, 6, f'=SUMIFS({apt("M")},{apt("C")},B{r},{apt("D")},C{r})', fmt=AREA_NZ, align=RIGHT)
            put(ws_f, r, 7, vals['com'], fmt=AREA, align=RIGHT)
            put(ws_f, r, 8, vals['cor'], fmt=AREA, align=RIGHT)
            put(ws_f, r, 9, stairs, fmt=AREA, align=RIGHT)
            put(ws_f, r, 10, vals['tech'], fmt=AREA, align=RIGHT)
            put(ws_f, r, 11, vals['park'], fmt=AREA, align=RIGHT)
            put(ws_f, r, 12, f'=SUM(F{r}:K{r})', f_bold, fmt=AREA, align=RIGHT)
            put(ws_f, r, 13, f'=ROUND(L{r}*{P_GROSS},1)', fmt=AREA, align=RIGHT)
            put(ws_f, r, 14, f'=ROUND(M{r}*{P_HEIGHT},0)', fmt=INT, align=RIGHT)
            for col in range(2, 15):
                c = ws_f.cell(r, col)
                c.border = B_BOTTOM
                if (r - s_first) % 2:
                    c.fill = ZEBRA
            r += 1
        # building subtotal
        put(ws_f, r, 4, f'{name} total', f_bold, align=RIGHT)
        put(ws_f, r, 5, f'=SUBTOTAL(9,E{s_first}:E{r - 1})', f_bold, fmt=INT, align=CENTER)
        for col in 'FGHIJKLMN':
            put(ws_f, r, ord(col) - 64, f'=SUBTOTAL(9,{col}{s_first}:{col}{r - 1})', f_bold,
                fmt=INT if col == 'N' else AREA, align=RIGHT)
        r += 2
    f_last = r - 1
    f_tot = r
    for col in range(2, 15):
        put(ws_f, f_tot, col, font=f_bold, fill=TOTAL_FILL, border=B_TOP)
    put(ws_f, f_tot, 2, 'Project total', f_bold, TOTAL_FILL, align=LEFT, border=B_TOP)
    put(ws_f, f_tot, 5, f'=SUBTOTAL(9,E{f_first}:E{f_last})', f_bold, TOTAL_FILL, INT, CENTER, B_TOP)
    for col in 'FGHIJKLMN':
        put(ws_f, f_tot, ord(col) - 64, f'=SUBTOTAL(9,{col}{f_first}:{col}{f_last})', f_bold, TOTAL_FILL,
            INT if col == 'N' else AREA, RIGHT, B_TOP)
    ws_f.freeze_panes = 'C5'

    def ftot(col):
        return f'Floors!${col}${f_tot}'

    # ================================================================ Summary
    sheet_header(ws_s, 'Nordic Park Residences',
                 f'Key technical indicators · {BUILDINGS} buildings · 18 storeys · {n_apts} apartments',
                 {'B': 30, 'C': 2, 'D': 13, 'E': 20, 'F': 3, 'G': 30, 'H': 17, 'I': 3}, zoom=115)

    put(ws_s, 4, 2, 'Design parameters', f_h2)
    for i, (label, value, fmt) in enumerate((('Balcony area factor', 0.3, '0.00'),
                                             ('Gross-to-net factor', 1.14, '0.00'),
                                             ('Floor-to-floor height, m', 3.15, '0.00'),
                                             ('Plot area, m²', PLOT_AREA, INT))):
        row = 6 + i
        put(ws_s, row, 2, label, align=LEFT, border=B_BOTTOM)
        put(ws_s, row, 3, border=B_BOTTOM)
        put(ws_s, row, 4, value, f_bold, INPUT_FILL, fmt, RIGHT, B_BOTTOM)
    put(ws_s, 6, 5, '← try changing this', f_hint)

    put(ws_s, 4, 7, 'Key indicators', f_h2)
    fp_total = sum(FOOTPRINT)
    kpis = (
        ('Building footprint', f'={fp_total}', M2_INT, False),
        ('Site coverage', f'=H6/{P_PLOT}', PCT, False),
        ('Gross floor area (GFA)', f'={ftot("M")}', M2_INT, True),
        ('Floor area ratio (FAR)', f'=ROUND(H8/{P_PLOT},2)', '0.00', False),
        ('Net internal area (NIA)', f'={ftot("L")}', M2_INT, False),
        ('Residential sellable area', f'={ftot("F")}', M2_INT, True),
        ('Commercial area', f'={ftot("G")}', M2_INT, False),
        ('Common areas', f'={ftot("H")}+{ftot("I")}+{ftot("J")}', M2_INT, False),
        ('Parking', f'={ftot("K")}', M2_INT, False),
        ('Efficiency (sellable / GFA)', '=H11/H8', PCT, True),
        ('Building volume', f'={ftot("N")}', M3, False),
        ('Apartments', f'=COUNTA({apt("E")})', INT, False),
        ('Average apartment', '=H11/H17', M2, False),
    )
    for i, (label, formula, fmt, big) in enumerate(kpis):
        row = 6 + i
        put(ws_s, row, 7, label, align=LEFT, border=B_BOTTOM)
        put(ws_s, row, 8, formula, f_kpi if big else f_bold, None, fmt, RIGHT, B_BOTTOM)
        if big:
            ws_s.row_dimensions[row].height = 24
    k_last = 6 + len(kpis) - 1

    def breakdown(top, title, first_head, keys, key_col, extra_head, extra_formula, extra_fmt):
        put(ws_s, top, 2, title, f_h2)
        table_head(ws_s, top + 1, 2, [first_head, None, 'Units', 'Living area, m²', None, 'Total area, m²', extra_head])
        for i, key in enumerate(keys):
            row = top + 2 + i
            put(ws_s, row, 2, key, align=LEFT, border=B_BOTTOM)
            put(ws_s, row, 4, f'=COUNTIF({apt(key_col)},B{row})', fmt=INT, align=RIGHT, border=B_BOTTOM)
            put(ws_s, row, 5, f'=SUMIF({apt(key_col)},B{row},{apt("H")})', fmt=AREA, align=RIGHT, border=B_BOTTOM)
            put(ws_s, row, 7, f'=SUMIF({apt(key_col)},B{row},{apt("M")})', f_bold, fmt=AREA, align=RIGHT,
                border=B_BOTTOM)
            put(ws_s, row, 8, extra_formula(row), fmt=extra_fmt, align=RIGHT, border=B_BOTTOM)
            for c in (3, 6):
                put(ws_s, row, c, border=B_BOTTOM)
        return top + 2, top + 1 + len(keys)

    t_first, t_last = breakdown(k_last + 3, 'By apartment type', 'Type', TYPES, 'F', 'Average area, m²',
                                lambda row: f'=ROUND(G{row}/D{row},1)', AREA)
    breakdown(t_last + 3, 'By building', 'Building', [f'Building {b}' for b in range(1, BUILDINGS + 1)], 'C',
              'Share of sellable area', lambda row: f'=G{row}/$H$11', PCT)

    chart = BarChart()
    chart.type = 'col'
    chart.title = 'Total area by apartment type, m²'
    chart.style = 2
    chart.legend = None
    chart.height = 7.5
    chart.width = 14
    chart.y_axis.numFmt = '#,##0'
    chart.y_axis.majorGridlines = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.add_data(Reference(ws_s, min_col=7, min_row=t_first, max_row=t_last), titles_from_data=False)
    chart.set_categories(Reference(ws_s, min_col=2, min_row=t_first, max_row=t_last))
    chart.series[0].graphicalProperties.solidFill = '8C7B6B'
    chart.series[0].graphicalProperties.line.solidFill = '8C7B6B'
    ws_s.add_chart(chart, 'J4')

    wb.save(OUT)
    return OUT


if __name__ == '__main__':
    print(build())
