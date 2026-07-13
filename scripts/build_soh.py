import csv, datetime, sys
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

import os
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXTRACTED = os.path.join(REPO_ROOT, "data", "masters") + os.sep

N_WEEKS = int(sys.argv[1]) if len(sys.argv) > 1 else 104
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else os.path.join(REPO_ROOT, "SOH_Master.xlsx")
START_MONDAY = datetime.date(2026, 6, 1)  # matches sample data period

def load_csv(name):
    with open(EXTRACTED + name, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

rm_master = load_csv("rm_master.csv")
inter_master = load_csv("intermediate_master.csv")
bom = load_csv("bom.csv")
prodmap = load_csv("product_intermediate_map.csv")
plan_sample = load_csv("production_plan_sample.csv")

# clean numeric fields, drop rows with missing numeric qty
def to_float(x, default=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default

for r in bom:
    r["RM_Qty_Per_Batch"] = to_float(r["RM_Qty_Per_Batch"])
    r["Batch_Size"] = to_float(r["Batch_Size"], 1.0)
    # RM_Qty_Per_Batch in source csv is actually the per-unit ratio; total per batch = ratio * batch size
    r["RM_Total_Per_Batch"] = r["RM_Qty_Per_Batch"] * (r["Batch_Size"] or 1.0)

bom = [r for r in bom if r["RM_Total_Per_Batch"] not in (None, 0) or True]  # keep all, zero rows are fine

# dedupe rm_master by RM_Code (defensive)
seen = set()
rm_master_dedup = []
for r in rm_master:
    if r["RM_Code"] in seen:
        continue
    seen.add(r["RM_Code"])
    rm_master_dedup.append(r)
rm_master = rm_master_dedup

seen = set()
inter_dedup = []
for r in inter_master:
    if r["Intermediate"] in seen:
        continue
    seen.add(r["Intermediate"])
    inter_dedup.append(r)
inter_master = inter_dedup

# production plan lookup by (Intermediate, WeekIndex-approx) -- we don't have true batch counts,
# only historic quantity samples for a few weeks (23-31ish). We'll map sample weeks 23.. to WeekIndex 1..
plan_lookup = {}
for r in plan_sample:
    try:
        wk = int(float(r["Week"]))
    except (TypeError, ValueError):
        continue
    qty = to_float(r["Qty"], 0.0)
    plan_lookup.setdefault(r["Intermediate"], {})[wk] = qty

print(f"RM master: {len(rm_master)}, Intermediates: {len(inter_master)}, BOM rows: {len(bom)}, ProductMap: {len(prodmap)}")

# ---------------------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active)

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")

def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

def add_table(ws, name, ref, style="TableStyleMedium9"):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)

# ============================================================ Cal_Weeks
ws = wb.create_sheet("Cal_Weeks")
ws.append(["WeekIndex", "WeekStart", "WeekEnd", "ISOWeek", "Month", "Year"])
for i in range(1, N_WEEKS + 1):
    wk_start = START_MONDAY + datetime.timedelta(weeks=i - 1)
    wk_end = wk_start + datetime.timedelta(days=6)
    iso_week = wk_start.isocalendar()[1]
    ws.append([i, wk_start, wk_end, iso_week, wk_start.month, wk_start.year])
style_header(ws, 6)
add_table(ws, "Cal_Weeks", f"A1:F{N_WEEKS+1}")
for col, w in zip("ABCDEF", [10, 12, 12, 9, 8, 8]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# ============================================================ M_RawMaterials
ws = wb.create_sheet("M_RawMaterials")
ws.append(["RM_Code", "Description", "Supplier", "Category", "UOM", "SafetyStock_Qty_要入力", "LeadTime_Weeks_要入力"])
for r in rm_master:
    ws.append([r["RM_Code"], r["Description"], r["Supplier"], r["Category"], "kg", 0, 4])
n = len(rm_master) + 1
style_header(ws, 7)
add_table(ws, "M_RawMaterials", f"A1:G{n}")
for col, w in zip("ABCDEFG", [14, 34, 14, 16, 8, 20, 18]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# ============================================================ M_Intermediates
ws = wb.create_sheet("M_Intermediates")
ws.append(["Intermediate", "Type", "PGM", "BatchSize"])
for r in inter_master:
    ws.append([r["Intermediate"], r["Type"], r["PGM"], to_float(r["Batch_Size"], 1.0)])
n = len(inter_master) + 1
style_header(ws, 4)
add_table(ws, "M_Intermediates", f"A1:D{n}")
for col, w in zip("ABCD", [14, 10, 8, 12]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# ============================================================ M_BOM
ws = wb.create_sheet("M_BOM")
ws.append(["Intermediate", "RM_Code", "RM_Qty_Per_Batch"])
for r in bom:
    ws.append([r["Intermediate"], r["RM_Code"], r["RM_Total_Per_Batch"]])
n = len(bom) + 1
style_header(ws, 3)
add_table(ws, "M_BOM", f"A1:C{n}")
for col, w in zip("ABC", [14, 12, 18]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# ============================================================ M_ProductMap
ws = wb.create_sheet("M_ProductMap")
ws.append(["Intermediate", "Product_Code"])
for r in prodmap:
    ws.append([r["Intermediate"], r["Product_Code"]])
n = len(prodmap) + 1
style_header(ws, 2)
add_table(ws, "M_ProductMap", f"A1:B{n}")
for col, w in zip("AB", [14, 16]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

rm_row = {r["RM_Code"]: i + 2 for i, r in enumerate(rm_master)}          # Grid_* row for RM_Code (row1=header)
inter_row = {r["Intermediate"]: i + 2 for i, r in enumerate(inter_master)}  # PP_Grid row for Intermediate

def week_col(week_idx):
    return get_column_letter(1 + week_idx)  # col A = label, col B = week1

# ============================================================ PP_Grid (production plan, INPUT)
ws = wb.create_sheet("PP_Grid")
header = ["Intermediate"] + [f"W{w}" for w in range(1, N_WEEKS + 1)]
ws.append(header)
for i, r in enumerate(inter_master):
    row_num = i + 2
    inter = r["Intermediate"]
    row_vals = [inter]
    weekvals = plan_lookup.get(inter, {})
    for w in range(1, N_WEEKS + 1):
        # seed sample data into first weeks if available (approx mapping from source week numbers 23..)
        src_week = 22 + w
        row_vals.append(to_float(weekvals.get(src_week, 0), 0))
    ws.append(row_vals)
    for w in range(1, N_WEEKS + 1):
        ws.cell(row=row_num, column=1 + w).fill = INPUT_FILL
n = len(inter_master) + 1
style_header(ws, N_WEEKS + 1)
add_table(ws, "PP_Grid", f"A1:{week_col(N_WEEKS)}{n}")
ws.column_dimensions["A"].width = 14
for w in range(1, N_WEEKS + 1):
    ws.column_dimensions[week_col(w)].width = 7
ws.freeze_panes = "B2"

# ============================================================ T_OpeningStock (INPUT)
ws = wb.create_sheet("T_OpeningStock")
ws.append(["RM_Code", "Opening_Qty_要入力", "AsOf"])
for r in rm_master:
    ws.append([r["RM_Code"], 0, START_MONDAY])
n = len(rm_master) + 1
style_header(ws, 3)
add_table(ws, "T_OpeningStock", f"A1:C{n}")
for c in range(2, 4):
    for row in range(2, n + 1):
        ws.cell(row=row, column=c).fill = INPUT_FILL
for col, w in zip("ABC", [14, 16, 12]):
    ws.column_dimensions[col].width = w

# ============================================================ T_Shipments (INPUT, seeded from Shipping Schedule)
raw_shipments = load_csv("shipments_all.csv")
ship_rows = []
for row in raw_shipments:
    rm_code = row["RM_Code"]
    if rm_code is None or rm_code not in rm_row:
        continue
    eta_str = row["Latest_ETA"]
    if not eta_str:
        continue
    eta = datetime.date.fromisoformat(eta_str)
    recv_str = row["Received_Date"]
    recv_date = datetime.date.fromisoformat(recv_str) if recv_str else None
    # Only keep shipments relevant to the forecast horizon: still pending, or
    # received recently (last 14d) / due soon. Older fully-received shipments
    # are assumed already folded into T_OpeningStock and must not be double counted.
    horizon_end = START_MONDAY + datetime.timedelta(weeks=N_WEEKS)
    relevant = (recv_date is None and eta < horizon_end) or \
               (recv_date is not None and recv_date >= START_MONDAY - datetime.timedelta(days=14) and recv_date < horizon_end) or \
               (recv_date is None and eta >= START_MONDAY - datetime.timedelta(days=14))
    if not relevant:
        continue
    ship_rows.append([rm_code, row["PO_No"] or "", to_float(row["Confirmed_Qty"], 0), eta, recv_date, row["Status"] or ""])

ws = wb.create_sheet("T_Shipments")
ws.append(["RM_Code", "PO_No", "Confirmed_Qty", "Latest_ETA", "Received_Date", "Status", "Effective_Week"])
start_row = 2
for i, r in enumerate(ship_rows):
    row_num = start_row + i
    ws.append(r + [None])
    # Effective_Week: use Received_Date if present else Latest_ETA; week index via anchor arithmetic, clamped to sheet horizon
    ws.cell(row=row_num, column=7).value = (
        f'=MAX(1,MIN({N_WEEKS},INT((IF(E{row_num}="",D{row_num},E{row_num})-DATE({START_MONDAY.year},{START_MONDAY.month},{START_MONDAY.day}))/7)+1))'
    )
n = len(ship_rows) + 1
style_header(ws, 7)
add_table(ws, "T_Shipments", f"A1:G{n}")
for col, w in zip("ABCDEFG", [12, 14, 14, 14, 14, 14, 14]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
print("Shipment rows seeded:", len(ship_rows))

# ============================================================ T_StockCount (INPUT, physical count overrides)
ws = wb.create_sheet("T_StockCount")
ws.append(["RM_Code", "WeekIndex", "CountedQty", "Notes"])
ws.append(["(例) CHEM-1010", 1, 0, "棚卸実施時にこの行へ追記"])
style_header(ws, 4)
add_table(ws, "T_StockCount", "A1:D2")
for col, w in zip("ABCD", [16, 10, 12, 30]):
    ws.column_dimensions[col].width = w

jj = 0
def col_letter_ok(ci):
    return get_column_letter(ci)

# ============================================================ Calc_Demand (explosion: BOM row x week)
ws = wb.create_sheet("Calc_Demand")
ws.append(["Intermediate", "RM_Code", "WeekIndex", "RM_Qty_Per_Batch", "Batches", "Demand"])
row_num = 1
for r in bom:
    inter = r["Intermediate"]
    pp_row = inter_row.get(inter)
    for w in range(1, N_WEEKS + 1):
        row_num += 1
        if pp_row:
            batches_formula = f"='PP_Grid'!{week_col(w)}{pp_row}"
        else:
            batches_formula = 0
        ws.append([inter, r["RM_Code"], w, r["RM_Total_Per_Batch"], batches_formula, f"=D{row_num}*E{row_num}"])
n = row_num
style_header(ws, 6)
add_table(ws, "Calc_Demand", f"A1:F{n}", style="TableStyleLight9")
for col, w in zip("ABCDEF", [14, 12, 10, 16, 10, 12]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.sheet_state = "hidden"
print("Calc_Demand rows:", n - 1)

# ============================================================ Grid_Requirement / Grid_Incoming / Grid_Stock
ws_req = wb.create_sheet("Grid_Requirement")
ws_in = wb.create_sheet("Grid_Incoming")
ws_st = wb.create_sheet("Grid_Stock")

header = ["RM_Code"] + [f"W{w}" for w in range(1, N_WEEKS + 1)]
ws_req.append(header)
ws_in.append(header)
ws_st.append(header)

for i, r in enumerate(rm_master):
    rr = i + 2
    rm = r["RM_Code"]
    ws_req.append([rm] + [None] * N_WEEKS)
    ws_in.append([rm] + [None] * N_WEEKS)
    ws_st.append([rm] + [None] * N_WEEKS)
    for w in range(1, N_WEEKS + 1):
        cl = week_col(w)
        ws_req.cell(row=rr, column=1 + w).value = (
            f"=SUMIFS(Calc_Demand[Demand],Calc_Demand[RM_Code],$A{rr},Calc_Demand[WeekIndex],{w})"
        )
        ws_in.cell(row=rr, column=1 + w).value = (
            f"=SUMIFS(T_Shipments[Confirmed_Qty],T_Shipments[RM_Code],$A{rr},T_Shipments[Effective_Week],{w})"
        )
        has_count = f"SUMPRODUCT((T_StockCount[RM_Code]=$A{rr})*(T_StockCount[WeekIndex]={w}))"
        count_val = f"SUMPRODUCT((T_StockCount[RM_Code]=$A{rr})*(T_StockCount[WeekIndex]={w})*T_StockCount[CountedQty])"
        if w == 1:
            prior = f'IFERROR(INDEX(T_OpeningStock[Opening_Qty],MATCH($A{rr},T_OpeningStock[RM_Code],0)),0)'
        else:
            prior = f"{week_col(w-1)}{rr}"
        normal = f"{prior}+'Grid_Incoming'!{cl}{rr}-'Grid_Requirement'!{cl}{rr}"
        ws_st.cell(row=rr, column=1 + w).value = f"=IF({has_count}>0,{count_val},{normal})"

n = len(rm_master) + 1
for ws_ in (ws_req, ws_in, ws_st):
    style_header(ws_, N_WEEKS + 1)
    add_table(ws_, ws_.title, f"A1:{week_col(N_WEEKS)}{n}", style="TableStyleMedium2")
    ws_.column_dimensions["A"].width = 14
    for w in range(1, N_WEEKS + 1):
        ws_.column_dimensions[week_col(w)].width = 9
    ws_.freeze_panes = "B2"

from openpyxl.formatting.rule import CellIsRule

# ============================================================ Dashboard
ws = wb.create_sheet("Dashboard")
ws.append(["RM_Code", "Description", "Category", "Supplier", "SafetyStock_Qty",
           "CurrentStock(W1)", "MinStock_2Y", "MinStock_Week", "Status"])
last_col = week_col(N_WEEKS)
for i, r in enumerate(rm_master):
    rr = i + 2
    rm = r["RM_Code"]
    ws.append([
        rm,
        f'=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A{rr},M_RawMaterials[RM_Code],0)),"")',
        f'=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A{rr},M_RawMaterials[RM_Code],0)),"")',
        f'=IFERROR(INDEX(M_RawMaterials[Supplier],MATCH($A{rr},M_RawMaterials[RM_Code],0)),"")',
        f'=IFERROR(INDEX(M_RawMaterials[SafetyStock_Qty_要入力],MATCH($A{rr},M_RawMaterials[RM_Code],0)),0)',
        f"='Grid_Stock'!B{rr}",
        f"=MIN('Grid_Stock'!B{rr}:{last_col}{rr})",
        f"=MATCH(G{rr},'Grid_Stock'!B{rr}:{last_col}{rr},0)",
        f'=IF(G{rr}<E{rr},"要発注","OK")',
    ])
n = len(rm_master) + 1
style_header(ws, 9)
add_table(ws, "Dashboard", f"A1:I{n}", style="TableStyleMedium4")
for col, w in zip("ABCDEFGHI", [14, 32, 16, 14, 14, 14, 12, 12, 10]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.conditional_formatting.add(
    f"I2:I{n}",
    CellIsRule(operator="equal", formula=['"要発注"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006"))
)

# ============================================================ PO Draft sheets (Chemical Release format)
def build_po_draft(sheet_name, category, title):
    ws = wb.create_sheet(sheet_name)
    ws["B2"] = "TO：（サプライヤー／TTAF担当者名を入力）"
    ws["B3"] = "　　　　　（会社名）"
    ws["B5"] = "FROM：（発行者名）"
    ws["B6"] = "　　　　　（自社名）"
    ws["N2"] = "Order Date:"
    ws["P2"] = datetime.date.today()
    ws["N3"] = "Issue Month:"
    ws["N4"] = "Firm Month:"
    ws["N5"] = "Revision"
    ws["P5"] = "00"
    ws["B8"] = title
    ws["B8"].font = Font(bold=True, size=12)

    hdr_row = 10
    headers = ["Part Name", "TTAF Code", "CSA Code", "UOM", "SafetyStock", "CurrentStock"] + \
              [f"W{w}" for w in range(1, N_WEEKS + 1)] + ["Total"]
    for c, h in enumerate(headers, start=2):
        ws.cell(row=hdr_row, column=c, value=h)
    style_header(ws, len(headers) + 1, row=hdr_row)

    data_row = hdr_row
    first_week_col = 8
    items = [r for r in rm_master if r["Category"] == category]
    for r in items:
        data_row += 1
        rm = r["RM_Code"]
        grow = rm_row[rm]
        ws.cell(row=data_row, column=2, value=f'=IFERROR(INDEX(M_RawMaterials[Description],MATCH("{rm}",M_RawMaterials[RM_Code],0)),"")')
        ws.cell(row=data_row, column=3, value=r.get("TTAF_Code", ""))
        ws.cell(row=data_row, column=4, value=rm)
        ws.cell(row=data_row, column=5, value="kg")
        ws.cell(row=data_row, column=6, value=f'=IFERROR(INDEX(M_RawMaterials[SafetyStock_Qty_要入力],MATCH("{rm}",M_RawMaterials[RM_Code],0)),0)')
        ws.cell(row=data_row, column=7, value=f"='Grid_Stock'!B{grow}")
        for w in range(1, N_WEEKS + 1):
            cl = get_column_letter(first_week_col + w - 1)
            gs_col = week_col(w)
            ws.cell(row=data_row, column=first_week_col + w - 1,
                    value=f"=MAX(0,$F{data_row}-'Grid_Stock'!{gs_col}{grow})")
        total_col = first_week_col + N_WEEKS
        rng_start = get_column_letter(first_week_col)
        rng_end = get_column_letter(first_week_col + N_WEEKS - 1)
        ws.cell(row=data_row, column=total_col, value=f"=SUM({rng_start}{data_row}:{rng_end}{data_row})")

    last_row = data_row
    last_col_idx = 2 + len(headers) - 1
    add_table(ws, sheet_name.replace(" ", "_"), f"B{hdr_row}:{get_column_letter(last_col_idx)}{last_row}", style="TableStyleMedium6")
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    for w in range(1, N_WEEKS + 1):
        ws.column_dimensions[get_column_letter(first_week_col + w - 1)].width = 8
    ws.freeze_panes = f"B{hdr_row+1}"
    print(f"{sheet_name}: {len(items)} items")

build_po_draft("PO_Draft_Chemical", "Chemical", "Chemicals : TTAF Supply")
build_po_draft("PO_Draft_Hazardous", "Hazardous Chemical", "Hazardous Chemicals")
build_po_draft("PO_Draft_Substrate", "Substrate", "Substrates")

# ============================================================ README
ws = wb.create_sheet("README", 0)
readme_lines = [
    "SOH（在庫管理）シート 使い方",
    "",
    "【シート構成】",
    "  M_RawMaterials / M_Intermediates / M_BOM / M_ProductMap : マスタ（原材料・中間体・原単位・完成品紐付け）",
    "  PP_Grid            : 生産計画（中間体×週のバッチ数）。ここを更新すると全体に自動反映されます【入力】",
    "  T_OpeningStock      : 起点となる期首在庫【入力】",
    "  T_Shipments         : 発注〜輸送〜着荷の実績・予定（ETAが変わると自動的に週がズレます）【入力】",
    "  T_StockCount        : 棚卸の実測値。入力するとその週以降の在庫計算がリセットされます【入力】",
    "  Calc_Demand         : 原単位展開の計算過程（非表示・監査用）",
    "  Grid_Requirement    : 原材料の週次所要量（自動計算）",
    "  Grid_Incoming       : 原材料の週次入荷予定（自動計算）",
    "  Grid_Stock          : 原材料の週次在庫（自動計算・2年先まで）",
    "  Dashboard           : 品目ごとの現在庫・最小在庫・要発注アラートを一覧表示",
    "  PO_Draft_*          : 要発注分を注文書ひな形（Chemical Release形式）に自動転記",
    "",
    "【週次の運用】",
    "  1. 生産計画が確定/変更されたら PP_Grid を更新",
    "  2. CSA Reportの最新情報でT_Shipments のETA/着荷日を更新（早着・遅着はここに反映）",
    "  3. 棚卸を実施したらT_StockCountに実測値を追記",
    "  4. Dashboardで「要発注」を確認し、PO_Draft_*から注文書を出力",
    "",
    "【前提・要確認事項】",
    "  - M_RawMaterials の SafetyStock_Qty と LeadTime_Weeks は仮値です。実際の安全在庫水準に置き換えてください。",
    "  - Categoryの割り当て(Chemical/Hazardous Chemical/Substrate)は入手データから機械的に推定した部分があります。要レビュー。",
    "  - 生産計画(PP_Grid)は現状の顧客オーダーからバッチサイズ固定で算出される前提の入力データです。",
]
for i, line in enumerate(readme_lines, start=1):
    ws.cell(row=i, column=1, value=line)
ws["A1"].font = Font(bold=True, size=14)
ws.column_dimensions["A"].width = 100

wb.save(OUT_PATH)
print("Full workbook written to", OUT_PATH)
