from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "DDEBF7"
INPUT_FILL = "FFF2CC"   # light yellow = editable input
GREY = "F2F2F2"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(bold=False, size=10, color="000000", italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def title(ws, text, sub=None, width_cols=10):
    ws["A1"] = text
    ws["A1"].font = f(True, 16, NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f(False, 10, "595959", italic=True)
    ws.sheet_view.showGridLines = False


def header_row(ws, row, headers, col=1, color=NAVY):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font = f(True, 10, "FFFFFF")
        c.fill = fill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def style_body(ws, r1, r2, c1, c2, num_formats=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = f()
            cell.border = BORDER
            if num_formats and c in num_formats:
                cell.number_format = num_formats[c]


def chart_fix(ch, axes=True):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.overlay = False
    if ch.title is not None:
        try:
            ch.title.overlay = False
        except Exception:
            pass
    if axes:
        ch.x_axis.delete = False
        ch.y_axis.delete = False
    return ch


def kpi(ws, row, col, label, formula, fmt="0"):
    """2-row KPI tile: label on top, value below, spanning 2 columns."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
    lc = ws.cell(row=row, column=col, value=label)
    lc.font = f(True, 9, "FFFFFF")
    lc.fill = fill(TEAL)
    lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    vc = ws.cell(row=row + 1, column=col, value=formula)
    vc.font = f(True, 18, NAVY)
    vc.fill = fill(LIGHT)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.number_format = fmt
    ws.row_dimensions[row + 1].height = 34
