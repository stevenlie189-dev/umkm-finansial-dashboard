#!/usr/bin/env python3
"""Memutar ulang operasi hasil mock Code.gs (tools/_work/model.json) menjadi file .xlsx.

Dengan cara ini Code.gs adalah satu-satunya sumber formula/format; versi Excel selalu identik.
Pakai:  TZ=Asia/Jakarta node tools/mock_sheets.js && python3 tools/build_xlsx.py
"""
import json, re, datetime, sys
from pathlib import Path
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList

ROOT = Path(__file__).resolve().parent.parent
model = json.loads((ROOT / "tools/_work/model.json").read_text())
args = [a for a in sys.argv[1:] if not a.startswith("--")]
OUT = Path(args[0]) if args else ROOT / "ROTI_RONA_Pembukuan.xlsx"

def hexc(c): return c.lstrip("#").upper()
def val(v):
    if isinstance(v, dict) and "date" in v: return datetime.date.fromisoformat(v["date"])
    return v

wb = Workbook(); wb.remove(wb.active)
for sd in model["sheets"]:
    ws = wb.create_sheet(sd["name"]); meta = sd["meta"]
    style = defaultdict(dict); border = defaultdict(dict)
    for o in sd["ops"]:
        op, a = o["op"], o["args"]
        cells = [(r, c) for r in range(o["r1"], o["r2"] + 1) for c in range(o["c1"], o["c2"] + 1)]
        if op in ("setValue", "setFormula"):
            for r, c in cells: ws.cell(r, c).value = val(a[0])
        elif op in ("setValues", "setFormulas"):
            for i, row in enumerate(a[0]):
                for j, v in enumerate(row): ws.cell(o["r1"] + i, o["c1"] + j).value = val(v)
        elif op == "clearContent":
            for r, c in cells: ws.cell(r, c).value = None
        elif op == "merge":
            ws.merge_cells(start_row=o["r1"], start_column=o["c1"], end_row=o["r2"], end_column=o["c2"])
        elif op == "setDataValidation":
            d = a[0]
            if d["type"] == "list": dv = DataValidation(type="list", formula1='"%s"' % ",".join(d["values"]))
            elif d["type"] == "range":
                s, e = d["range"].split(":"); ab = lambda x: re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", x)
                dv = DataValidation(type="list", formula1=f"{d['sheet']}!{ab(s)}:{ab(e)}")
            elif d["type"] == "gte": dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(d["n"]))
            else: dv = DataValidation(type="date", operator="greaterThan", formula1="1")
            dv.allow_blank = True; dv.showErrorMessage = not d.get("allowInvalid", True)
            dv.errorTitle = "Input tidak valid"; dv.error = d.get("help", "")
            dv.showInputMessage = True; dv.promptTitle = "Petunjuk"; dv.prompt = (d.get("help") or "")[:250]
            dv.add(f"{L(o['c1'])}{o['r1']}:{L(o['c2'])}{o['r2']}"); ws.add_data_validation(dv)
        elif op == "setBorder":
            top, left, bottom, right, vert, horiz, color, st = a
            sty = "thick" if st == "SOLID_THICK" else "thin"; side = Side(style=sty, color=hexc(color or "#000000"))
            for r, c in cells:
                if (top and r == o["r1"]) or (horiz and r > o["r1"]): border[(r, c)]["top"] = side
                if (bottom and r == o["r2"]) or (horiz and r < o["r2"]): border[(r, c)]["bottom"] = side
                if (left and c == o["c1"]) or (vert and c > o["c1"]): border[(r, c)]["left"] = side
                if (right and c == o["c2"]) or (vert and c < o["c2"]): border[(r, c)]["right"] = side
        else:
            key = {"setNumberFormat": "nf", "setBackground": "bg", "setFontColor": "fc", "setFontWeight": "bold",
                   "setFontStyle": "italic", "setFontSize": "size", "setFontFamily": "family",
                   "setHorizontalAlignment": "ha", "setVerticalAlignment": "va", "setWrap": "wrap"}[op]
            v = a[0]
            for r, c in cells: style[(r, c)][key] = v
    for (r, c), s in style.items():
        cell = ws.cell(r, c)
        if "nf" in s: cell.number_format = s["nf"]
        if "bg" in s: cell.fill = PatternFill("solid", start_color=hexc(s["bg"]), end_color=hexc(s["bg"]))
        cell.font = Font(name=s.get("family", "Arial"), size=s.get("size", 10), bold=s.get("bold") == "bold",
                         italic=s.get("italic") == "italic", color=hexc(s["fc"]) if "fc" in s else None)
        cell.alignment = Alignment(horizontal=s.get("ha"), vertical={"middle": "center"}.get(s.get("va"), s.get("va")),
                                   wrap_text=bool(s.get("wrap")))
    for (r, c), sides in border.items(): ws.cell(r, c).border = Border(**sides)
    for c, w in meta["colWidths"].items(): ws.column_dimensions[L(int(c))].width = max(1, w / 7.0)
    for r, h in meta["rowHeights"].items(): ws.row_dimensions[int(r)].height = h * 0.75
    if meta.get("frozen"): ws.freeze_panes = f"A{meta['frozen'] + 1}"
    if meta.get("hideGrid"): ws.sheet_view.showGridLines = False
    if meta.get("tab"): ws.sheet_properties.tabColor = hexc(meta["tab"])
    for rule in meta["cf"]:
        dxf = DifferentialStyle(
            font=Font(color=hexc(rule["fc"]) if rule.get("fc") else None, bold=True if rule.get("bold") else None),
            fill=PatternFill("solid", start_color=hexc(rule["bg"]), end_color=hexc(rule["bg"])) if rule.get("bg") else None)
        rl = Rule(type="expression", dxf=dxf, formula=[rule["formula"].lstrip("=")])
        for rg in rule["ranges"]: ws.conditional_formatting.add(rg, rl)
    # grafik
    for ch in meta["charts"]:
        opt = ch["options"]; w_cm = opt["width"] * 0.02646; h_cm = opt["height"] * 0.02646
        anchor = f"{L(ch['pos']['col'])}{ch['pos']['row']}"
        def ref(rg):
            m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rg)
            from openpyxl.utils import column_index_from_string as ci
            return dict(min_col=ci(m[1]), min_row=int(m[2]), max_col=ci(m[3]), max_row=int(m[4]))
        if ch["type"] == "PIE":
            cat, dat = ref(ch["ranges"][0]), ref(ch["ranges"][1])
            chart = DoughnutChart(); chart.holeSize = 50
            chart.add_data(Reference(ws, **dat), titles_from_data=True)
            chart.set_categories(Reference(ws, **{**cat, "min_row": cat["min_row"] + 1}))
            s = chart.series[0]
            for i, col in enumerate(opt["colors"]):
                pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = hexc(col); s.dPt.append(pt)
            chart.legend.position = "r"
        else:
            if ch["type"] == "COLUMN":
                r0 = ref(ch["ranges"][0]); chart = BarChart(); chart.type = "col"
                chart.add_data(Reference(ws, **r0).__class__(ws, min_col=r0["min_col"] + 1, max_col=r0["max_col"], min_row=r0["min_row"], max_row=r0["max_row"]), titles_from_data=True)
                chart.set_categories(Reference(ws, min_col=r0["min_col"], min_row=r0["min_row"] + 1, max_row=r0["max_row"]))
                chart.legend.position = "b"
            else:
                cat, dat = ref(ch["ranges"][0]), ref(ch["ranges"][1]); chart = BarChart(); chart.type = "bar"
                chart.add_data(Reference(ws, **dat), titles_from_data=True)
                chart.set_categories(Reference(ws, **{**cat, "min_row": cat["min_row"] + 1}))
                chart.legend = None
                chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True
                for f in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(chart.dataLabels, f, False)
            for s, col in zip(chart.series, opt["colors"]):
                s.graphicalProperties.solidFill = hexc(col); s.graphicalProperties.line.solidFill = hexc(col)
            chart.y_axis.numFmt = '#,##0'; chart.y_axis.delete = False; chart.x_axis.delete = False
            chart.gapWidth = 60
        chart.width, chart.height = w_cm, h_cm
        ws.add_chart(chart, anchor)
from openpyxl.worksheet.properties import PageSetupProperties
dash = wb["DASHBOARD"]
dash.page_setup.orientation = "portrait"; dash.page_setup.fitToWidth = 1; dash.page_setup.fitToHeight = 1
dash.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True); dash.print_area = "A1:N69"
if "--preview" in sys.argv:           # untuk tangkapan layar: sembunyikan sheet lain
    for n in ("TRANSAKSI", "KATEGORI", "PETUNJUK"): wb[n].sheet_state = "hidden"
wb.properties.title = model["name"]; wb.properties.creator = "ROTI RONA (simulasi)"
wb.active = 0
wb.save(OUT); print("tersimpan:", OUT)
