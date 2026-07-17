import csv, datetime, sys
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName

import os
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXTRACTED = os.path.join(REPO_ROOT, "data", "masters") + os.sep

N_WEEKS = int(sys.argv[1]) if len(sys.argv) > 1 else 104
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else os.path.join(REPO_ROOT, "SOH_Master.xlsx")
ANCHOR_YEAR = int(sys.argv[3]) if len(sys.argv) > 3 else 2026


def monday_containing_jan1(year):
    jan1 = datetime.date(year, 1, 1)
    return jan1 - datetime.timedelta(days=jan1.weekday())


def week_year_and_number(week_start):
    # Week1 of year Y = the Mon-Sun week that contains Jan 1 of Y.
    for y in (week_start.year - 1, week_start.year, week_start.year + 1):
        b = monday_containing_jan1(y)
        nb = monday_containing_jan1(y + 1)
        if b <= week_start < nb:
            return y, (week_start - b).days // 7 + 1
    raise ValueError(f"could not resolve week-year for {week_start}")


START_MONDAY = monday_containing_jan1(ANCHOR_YEAR)  # Week1 start = Monday of the week containing Jan 1


def date_to_week_index(d, n_weeks):
    idx = (d - START_MONDAY).days // 7 + 1
    return max(1, min(n_weeks, idx))

def load_csv(name):
    with open(EXTRACTED + name, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

rm_master = load_csv("rm_master.csv")
inter_master = load_csv("intermediate_master.csv")
bom = load_csv("bom.csv")
prodmap = load_csv("product_intermediate_map.csv")


def load_csv_optional(name):
    path = os.path.join(EXTRACTED, name)
    if not os.path.exists(path):
        return []
    return load_csv(name)


self_stock_sample = load_csv_optional("self_stock_sample.csv")
ttaf_stock_sample = load_csv_optional("ttaf_stock_sample.csv")
weekly_batches = load_csv("weekly_batches.csv")

# clean numeric fields, drop rows with missing numeric qty
def to_float(x, default=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default

# bom.csv (source: Usage from Production Engineering Slurry/Powder sheets, supplemented by
# Powder & Slurry & Pgm Plan) already stores the direct per-batch kg amount -- no further
# ratio x batch-size multiplication needed.
for r in bom:
    r["RM_Total_Per_Batch"] = to_float(r["RM_Qty_Per_Batch"])

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

# weekly batch counts, source: Powder & Slurry & Pgm Plan ("No. of batches" rows, one entry
# per intermediate per week, extracted from ~36 per-material sheets). Keyed by actual calendar
# date (WeekStart) so it lines up correctly regardless of ANCHOR_YEAR.
plan_lookup = {}
for r in weekly_batches:
    ws_str = r.get("WeekStart")
    if not ws_str:
        continue
    try:
        wk_date = datetime.date.fromisoformat(ws_str)
    except ValueError:
        continue
    qty = to_float(r["Batches"], 0.0)
    plan_lookup.setdefault(r["Intermediate"], {})[wk_date] = qty

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
# WeekIndex(内部の連番, グリッドの列位置に使用) と、年ごとに1へリセットするWeekOfYear/Label(表示用)。
# Week1 = AnchorYearの1月1日を含む月〜日の週。日付・週番号はすべてExcelの数式で計算される
# （B1のAnchorYearを変更すれば全体が再計算される）。
week_labels = {}  # Python側でも列位置の対応付けに使うため、値は事前計算して保持しておく
CAL_HEADER_ROW = 3  # row1=AnchorYear入力, row2=列見出し, row3以降=データ
ws = wb.create_sheet("Cal_Weeks")
ws["A1"] = "AnchorYear"
ws["B1"] = ANCHOR_YEAR
ws["A1"].font = Font(bold=True)
ws["B1"].fill = INPUT_FILL
ws.append([])
ws.append(["WeekIndex", "WeekStart", "Year", "WeekOfYear", "Label", "WeekEnd", "MonthYearLabel"])
for i in range(1, N_WEEKS + 1):
    r = CAL_HEADER_ROW + i
    if i == 1:
        wk_start_formula = "=DATE($B$1,1,1)-WEEKDAY(DATE($B$1,1,1),3)"
    else:
        wk_start_formula = f"=B{r-1}+7"
    year_formula = (
        f"=IF(B{r}>=DATE(YEAR(B{r})+1,1,1)-WEEKDAY(DATE(YEAR(B{r})+1,1,1),3),YEAR(B{r})+1,"
        f"IF(B{r}>=DATE(YEAR(B{r}),1,1)-WEEKDAY(DATE(YEAR(B{r}),1,1),3),YEAR(B{r}),YEAR(B{r})-1))"
    )
    week_of_year_formula = f"=(B{r}-(DATE(C{r},1,1)-WEEKDAY(DATE(C{r},1,1),3)))/7+1"
    label_formula = f'=C{r}&"-W"&TEXT(D{r},"00")'
    week_end_formula = f"=B{r}+6"
    month_year_formula = f'=TEXT(B{r},"mmm-yy")'
    ws.append([i, wk_start_formula, year_formula, week_of_year_formula, label_formula, week_end_formula, month_year_formula])
    ws.cell(row=r, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=6).number_format = "yyyy-mm-dd"
    # Python側の記録用(他シートのビルド時の列位置計算・MATCH対象文字列の生成に使用。
    # 実際のセルの値は上記の通りExcel数式で計算される)
    wk_start = START_MONDAY + datetime.timedelta(weeks=i - 1)
    yr, wn = week_year_and_number(wk_start)
    week_labels[i] = f"{yr}-W{wn:02d}"
style_header(ws, 7, row=CAL_HEADER_ROW)
add_table(ws, "Cal_Weeks", f"A{CAL_HEADER_ROW}:G{CAL_HEADER_ROW+N_WEEKS}")
for col, w in zip("ABCDEFG", [10, 12, 8, 11, 12, 12, 14]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = f"A{CAL_HEADER_ROW+1}"

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
header = ["Intermediate"] + [week_labels[w] for w in range(1, N_WEEKS + 1)]
ws.append(header)
for i, r in enumerate(inter_master):
    row_num = i + 2
    inter = r["Intermediate"]
    row_vals = [inter]
    weekvals = plan_lookup.get(inter, {})
    for w in range(1, N_WEEKS + 1):
        wk_date = START_MONDAY + datetime.timedelta(weeks=w - 1)
        row_vals.append(to_float(weekvals.get(wk_date, 0), 0))
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
    ship_rows.append([rm_code, row["PO_No"] or "", None, to_float(row["Confirmed_Qty"], 0), eta, recv_date, row["Status"] or ""])
    # ship_rows layout: RM_Code, PO_No, Order_Date(unknown from source, left blank for manual entry), Confirmed_Qty, Latest_ETA, Received_Date, Status

ws = wb.create_sheet("T_Shipments")
ws.append(["RM_Code", "PO_No", "Order_Date_発注日", "Confirmed_Qty", "Latest_ETA", "Received_Date", "Status", "Effective_Week"])
if not ship_rows:
    # Excelのテーブル機能は見出し行のみ(データ0行)の範囲を許容しないため、
    # 該当する発注が無い場合はダミー行を1行入れておく（要削除・上書き可）。
    ship_rows = [["(例) CHEM-1010", "", None, 0, START_MONDAY, None, ""]]
start_row = 2
for i, r in enumerate(ship_rows):
    row_num = start_row + i
    ws.append(r + [None])
    ws.cell(row=row_num, column=3).fill = INPUT_FILL  # Order_Date is not in the source file; input by hand
    # Effective_Week: use Received_Date if present else Latest_ETA; week index via anchor arithmetic, clamped to sheet horizon
    ws.cell(row=row_num, column=8).value = (
        f'=MAX(1,MIN({N_WEEKS},INT((IF(F{row_num}="",E{row_num},F{row_num})-DATE({START_MONDAY.year},{START_MONDAY.month},{START_MONDAY.day}))/7)+1))'
    )
n = len(ship_rows) + 1
style_header(ws, 8)
add_table(ws, "T_Shipments", f"A1:H{n}")
for col, w in zip("ABCDEFGH", [12, 14, 16, 14, 14, 14, 14, 14]):
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

# ============================================================ T_SelfStock (自社倉庫の実績, VBA更新)
ws = wb.create_sheet("T_SelfStock")
ws.append(["RM_Code", "WeekIndex", "Date", "Self_Qty"])
self_rows_written = 0
for r in self_stock_sample:
    d = datetime.date.fromisoformat(r["Date"])
    wi = date_to_week_index(d, N_WEEKS)
    ws.append([r["RM_Code"], wi, d, to_float(r["Self_Qty"], 0)])
    self_rows_written += 1
if self_rows_written == 0:
    ws.append(["(例) CHEM-1010", 1, START_MONDAY, 0])
    self_rows_written = 1
style_header(ws, 4)
add_table(ws, "T_SelfStock", f"A1:D{self_rows_written+1}")
for c, w in zip("ABCD", [16, 10, 12, 12]):
    ws.column_dimensions[c].width = w
for row_i in range(2, self_rows_written + 2):
    ws.cell(row=row_i, column=3).number_format = "yyyy-mm-dd"
print("T_SelfStock rows seeded:", self_rows_written)

# ============================================================ T_TTAFStock (TTAF倉庫の実績, VBA更新)
ws = wb.create_sheet("T_TTAFStock")
ws.append(["RM_Code", "WeekIndex", "Date", "TTAF_Qty"])
ttaf_rows_written = 0
for r in ttaf_stock_sample:
    d = datetime.date.fromisoformat(r["Date"])
    wi = date_to_week_index(d, N_WEEKS)
    ws.append([r["RM_Code"], wi, d, to_float(r["TTAF_Qty"], 0)])
    ttaf_rows_written += 1
if ttaf_rows_written == 0:
    ws.append(["(例) CHEM-1010", 1, START_MONDAY, 0])
    ttaf_rows_written = 1
style_header(ws, 4)
add_table(ws, "T_TTAFStock", f"A1:D{ttaf_rows_written+1}")
for c, w in zip("ABCD", [16, 10, 12, 12]):
    ws.column_dimensions[c].width = w
for row_i in range(2, ttaf_rows_written + 2):
    ws.cell(row=row_i, column=3).number_format = "yyyy-mm-dd"
print("T_TTAFStock rows seeded:", ttaf_rows_written)

# ============================================================ Calc_Demand (explosion: BOM row x week)
# Batches と RM_Qty_Per_Batch は PP_Grid / M_BOM への「ライブ参照」(INDEX/MATCH, SUMIFS)。
# VBAでPP_Grid・M_BOMの値やRM_Qty_Per_Batchを更新しても、既存の(中間体,原材料)組み合わせの行は
# 自動的に再計算される（行の追加・並べ替えにも耐える）。
# ただし全く新しい(中間体,原材料)の組み合わせが増えた場合はこの表自体の再生成が必要。
ws = wb.create_sheet("Calc_Demand")
ws.append(["Intermediate", "RM_Code", "WeekIndex", "Label", "RM_Qty_Per_Batch", "Batches", "Demand"])
row_num = 1
seen_pairs = set()
for r in bom:
    inter = r["Intermediate"]
    rm_code = r["RM_Code"]
    if (inter, rm_code) in seen_pairs:
        continue  # avoid duplicate explosion if the same pair appears from multiple source rows
    seen_pairs.add((inter, rm_code))
    for w in range(1, N_WEEKS + 1):
        row_num += 1
        ws.append([
            inter, rm_code, w, week_labels[w],
            f'=SUMIFS(M_BOM[RM_Qty_Per_Batch],M_BOM[Intermediate],A{row_num},M_BOM[RM_Code],B{row_num})',
            f'=IFERROR(INDEX(PP_Grid[#Data],MATCH(A{row_num},PP_Grid[Intermediate],0),MATCH(D{row_num},PP_Grid[#Headers],0)),0)',
            f"=E{row_num}*F{row_num}",
        ])
n = row_num
style_header(ws, 7)
add_table(ws, "Calc_Demand", f"A1:G{n}", style="TableStyleLight9")
for col, w in zip("ABCDEFG", [14, 12, 10, 10, 16, 10, 12]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.sheet_state = "hidden"
print("Calc_Demand rows:", n - 1)

# ============================================================ Grid_Requirement / Grid_Incoming / Grid_Stock
ws_req = wb.create_sheet("Grid_Requirement")
ws_in = wb.create_sheet("Grid_Incoming")
ws_st = wb.create_sheet("Grid_Stock")

header = ["RM_Code"] + [week_labels[w] for w in range(1, N_WEEKS + 1)]
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
        has_self = f"SUMPRODUCT((T_SelfStock[RM_Code]=$A{rr})*(T_SelfStock[WeekIndex]={w}))"
        self_val = f"SUMPRODUCT((T_SelfStock[RM_Code]=$A{rr})*(T_SelfStock[WeekIndex]={w})*T_SelfStock[Self_Qty])"
        has_ttaf = f"SUMPRODUCT((T_TTAFStock[RM_Code]=$A{rr})*(T_TTAFStock[WeekIndex]={w}))"
        ttaf_val = f"SUMPRODUCT((T_TTAFStock[RM_Code]=$A{rr})*(T_TTAFStock[WeekIndex]={w})*T_TTAFStock[TTAF_Qty])"
        if w == 1:
            prior = f'IFERROR(INDEX(T_OpeningStock[Opening_Qty],MATCH($A{rr},T_OpeningStock[RM_Code],0)),0)'
        else:
            prior = f"{week_col(w-1)}{rr}"
        normal = f"{prior}+'Grid_Incoming'!{cl}{rr}-'Grid_Requirement'!{cl}{rr}"
        # 優先順位: 手動棚卸(T_StockCount) > 自社+TTAF実績の合計(両方揃っている週のみ) > 通常のロールフォワード
        ws_st.cell(row=rr, column=1 + w).value = (
            f"=IF({has_count}>0,{count_val},"
            f"IF(({has_self}>0)*({has_ttaf}>0)>0,{self_val}+{ttaf_val},{normal}))"
        )

n = len(rm_master) + 1
for ws_ in (ws_req, ws_in, ws_st):
    style_header(ws_, N_WEEKS + 1)
    add_table(ws_, ws_.title, f"A1:{week_col(N_WEEKS)}{n}", style="TableStyleMedium2")
    ws_.column_dimensions["A"].width = 14
    for w in range(1, N_WEEKS + 1):
        ws_.column_dimensions[week_col(w)].width = 9
    ws_.freeze_panes = "B2"

# ============================================================ Material_Detail
# 「どの材料が何に使われているか」を、材料ごとにブロックで見せるトレーサビリティ表示。
# 90枚のシートに分けず、1枚のシートに材料ブロックを縦に並べる（重さの原因を再現しない）。
bom_by_rm = {}
for r in bom:
    bom_by_rm.setdefault(r["RM_Code"], []).append(r)

ws = wb.create_sheet("Material_Detail")
WEEK_START_COL = 4  # column D
def mdetail_week_col(w):
    return get_column_letter(WEEK_START_COL + w - 1)

MD_MONTHYEAR_ROW, MD_DATE_ROW, MD_WEEKNO_ROW, MD_TABLE_ROW = 1, 2, 3, 4
for w in range(1, N_WEEKS + 1):
    col = WEEK_START_COL + w - 1
    cal_row = CAL_HEADER_ROW + w
    if w == 1:
        my_formula = f"='Cal_Weeks'!G{cal_row}"
    else:
        prev_cal_row = CAL_HEADER_ROW + w - 1
        my_formula = (
            f'=IF(TEXT(\'Cal_Weeks\'!B{cal_row},"mmm-yy")<>TEXT(\'Cal_Weeks\'!B{prev_cal_row},"mmm-yy"),'
            f"'Cal_Weeks'!G{cal_row},\"\")"
        )
    ws.cell(row=MD_MONTHYEAR_ROW, column=col, value=my_formula)
    dcell = ws.cell(row=MD_DATE_ROW, column=col, value=f"='Cal_Weeks'!B{cal_row}")
    dcell.number_format = "m/d"
    ws.cell(row=MD_WEEKNO_ROW, column=col, value=f"='Cal_Weeks'!D{cal_row}")
    ws.cell(row=MD_TABLE_ROW, column=col, value=f"='Cal_Weeks'!E{cal_row}")
for r in (MD_MONTHYEAR_ROW, MD_DATE_ROW, MD_WEEKNO_ROW, MD_TABLE_ROW):
    for c in range(1, WEEK_START_COL + N_WEEKS):
        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="D9E1F2")
        ws.cell(row=r, column=c).font = Font(bold=(r in (MD_MONTHYEAR_ROW, MD_TABLE_ROW)))
ws.cell(row=MD_TABLE_ROW, column=1, value="RM_Code")
ws.cell(row=MD_TABLE_ROW, column=2, value="項目")
ws.cell(row=MD_TABLE_ROW, column=3, value="1バッチ使用量(kg)")

row_num = MD_TABLE_ROW
for rm_code, entries in bom_by_rm.items():
    if rm_code not in rm_row:
        continue
    grow = rm_row[rm_code]
    desc = next((r["Description"] for r in rm_master if r["RM_Code"] == rm_code), "")

    row_num += 1
    mat_header_row = row_num
    ws.cell(row=row_num, column=1, value=rm_code)
    ws.cell(row=row_num, column=2, value=desc)
    for c in range(1, len(header) + 1):
        ws.cell(row=row_num, column=c).fill = PatternFill("solid", fgColor="FFE699")
        ws.cell(row=row_num, column=c).font = Font(bold=True)

    for entry in entries:
        inter = entry["Intermediate"]

        row_num += 1
        batches_row = row_num
        ws.cell(row=row_num, column=2, value=inter)
        ws.cell(row=row_num, column=3, value="No. of batches")
        for w in range(1, N_WEEKS + 1):
            cell = ws.cell(row=row_num, column=WEEK_START_COL + w - 1)
            cell.value = (
                f'=IFERROR(INDEX(PP_Grid[#Data],MATCH($B{row_num},PP_Grid[Intermediate],0),'
                f'MATCH("{week_labels[w]}",PP_Grid[#Headers],0)),0)'
            )

        row_num += 1
        ws.cell(row=row_num, column=2, value="使用量(kg)")
        ws.cell(row=row_num, column=3,
                value=f'=SUMIFS(M_BOM[RM_Qty_Per_Batch],M_BOM[Intermediate],$B{row_num-1},M_BOM[RM_Code],$A{mat_header_row})')
        for w in range(1, N_WEEKS + 1):
            wc = mdetail_week_col(w)
            ws.cell(row=row_num, column=WEEK_START_COL + w - 1,
                    value=f"=$C{row_num}*{wc}{batches_row}")

    row_num += 1
    ws.cell(row=row_num, column=2, value="合計使用量(kg)/週")
    ws.cell(row=row_num, column=2).font = Font(bold=True)
    for w in range(1, N_WEEKS + 1):
        wc = mdetail_week_col(w)
        ws.cell(row=row_num, column=WEEK_START_COL + w - 1, value=f"='Grid_Requirement'!{week_col(w)}{grow}")

    row_num += 1
    ws.cell(row=row_num, column=2, value="（在庫・入荷予定はDashboard参照）")
    ws.cell(row=row_num, column=2).font = Font(italic=True, color="808080")

    row_num += 1  # blank separator row

last_row = row_num
ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 14
for w in range(1, N_WEEKS + 1):
    ws.column_dimensions[mdetail_week_col(w)].width = 9
ws.freeze_panes = f"{get_column_letter(WEEK_START_COL)}{MD_TABLE_ROW+1}"
print("Material_Detail: blocks for", len(bom_by_rm), "materials,", last_row, "rows")

from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ============================================================ Dashboard
# 「最終的にここで在庫を確認する」メイン画面。原材料×週の在庫を2年分横軸で見渡せる。
# A〜E列(RM情報)とヘッダー行(月-年/日付/週No)を固定して、右にスクロールしながら見る設計。
LEFT_COLS = ["RM_Code", "Description", "Category", "SafetyStock_Qty",
             "自社在庫(実績)", "TTAF在庫(実績)", "実績週", "Status"]
WEEK_START_COL_DASH = len(LEFT_COLS) + 1  # F列から週データ
HDR_MONTHYEAR_ROW = 3
HDR_DATE_ROW = 4
HDR_WEEKNO_ROW = 5
HDR_TABLE_ROW = 6  # Excel Table の見出し行としても使う(Label)
DATA_START_ROW = 7

ws = wb.create_sheet("Dashboard")
ws["A1"] = "この週へジャンプ（例: 2026-W23 の形式で入力）"
ws["A1"].font = Font(bold=True)
ws["C1"] = ""
ws["C1"].fill = INPUT_FILL
ws["C1"].font = Font(bold=True, size=12)
_hdr_start_col = get_column_letter(WEEK_START_COL_DASH)
_hdr_end_col = get_column_letter(WEEK_START_COL_DASH + N_WEEKS - 1)
# マクロ不要でジャンプできるHYPERLINKリンク（クリックすると該当週の列へ移動・自動スクロールされる）
ws["D1"] = (
    f'=IFERROR(HYPERLINK("#Dashboard!"&ADDRESS({DATA_START_ROW},'
    f'MATCH($C$1,{_hdr_start_col}${HDR_TABLE_ROW}:{_hdr_end_col}${HDR_TABLE_ROW},0)+{WEEK_START_COL_DASH}-1,4),'
    f'"▶ この週へジャンプ"),"週Noを入力してください（例: 2026-W23）")'
)
ws["D1"].font = Font(bold=True, color="0563C1")

for w in range(1, N_WEEKS + 1):
    col = WEEK_START_COL_DASH + w - 1
    cal_row = CAL_HEADER_ROW + w
    cl = get_column_letter(col)
    if w == 1:
        my_formula = f"='Cal_Weeks'!G{cal_row}"
    else:
        prev_cal_row = CAL_HEADER_ROW + w - 1
        my_formula = (
            f'=IF(TEXT(\'Cal_Weeks\'!B{cal_row},"mmm-yy")<>TEXT(\'Cal_Weeks\'!B{prev_cal_row},"mmm-yy"),'
            f"'Cal_Weeks'!G{cal_row},\"\")"
        )
    ws.cell(row=HDR_MONTHYEAR_ROW, column=col, value=my_formula)
    dcell = ws.cell(row=HDR_DATE_ROW, column=col, value=f"='Cal_Weeks'!B{cal_row}")
    dcell.number_format = "m/d"
    ws.cell(row=HDR_WEEKNO_ROW, column=col, value=f"='Cal_Weeks'!D{cal_row}")
    ws.cell(row=HDR_TABLE_ROW, column=col, value=f"='Cal_Weeks'!E{cal_row}")
    ws.column_dimensions[cl].width = 9

for c, name in enumerate(LEFT_COLS, start=1):
    ws.cell(row=HDR_TABLE_ROW, column=c, value=name)
style_header(ws, WEEK_START_COL_DASH + N_WEEKS - 1, row=HDR_TABLE_ROW)
for row_i in (HDR_MONTHYEAR_ROW, HDR_DATE_ROW, HDR_WEEKNO_ROW):
    for c in range(1, WEEK_START_COL_DASH + N_WEEKS):
        ws.cell(row=row_i, column=c).fill = PatternFill("solid", fgColor="D9E1F2")
        ws.cell(row=row_i, column=c).font = Font(bold=(row_i == HDR_MONTHYEAR_ROW))

for col, w in zip("ABCDEFGH", [14, 32, 16, 14, 14, 12, 12, 10]):
    ws.column_dimensions[col].width = w

last_col_dash = get_column_letter(WEEK_START_COL_DASH + N_WEEKS - 1)
for i, r in enumerate(rm_master):
    rr = DATA_START_ROW + i
    grow = rm_row[r["RM_Code"]]
    rm = r["RM_Code"]
    ws.cell(row=rr, column=1, value=rm)
    ws.cell(row=rr, column=2,
            value=f'=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A{rr},M_RawMaterials[RM_Code],0)),"")')
    ws.cell(row=rr, column=3,
            value=f'=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A{rr},M_RawMaterials[RM_Code],0)),"")')
    ws.cell(row=rr, column=4,
            value=f'=IFERROR(INDEX(M_RawMaterials[SafetyStock_Qty_要入力],MATCH($A{rr},M_RawMaterials[RM_Code],0)),0)')
    ws.cell(row=rr, column=5,
            value=(f'=IFERROR(LOOKUP(2,1/(T_SelfStock!$A$2:$A$2000=$A{rr}),T_SelfStock!$D$2:$D$2000),"")'))
    ws.cell(row=rr, column=6,
            value=(f'=IFERROR(LOOKUP(2,1/(T_TTAFStock!$A$2:$A$2000=$A{rr}),T_TTAFStock!$D$2:$D$2000),"")'))
    ws.cell(row=rr, column=7,
            value=(f'=IFERROR(INDEX(Cal_Weeks[Label],LOOKUP(2,1/(T_SelfStock!$A$2:$A$2000=$A{rr}),'
                   f'T_SelfStock!$B$2:$B$2000)),"")'))
    ws.cell(row=rr, column=8,
            value=f'=IF(MIN({get_column_letter(WEEK_START_COL_DASH)}{rr}:{last_col_dash}{rr})<D{rr},"要発注","OK")')
    for w in range(1, N_WEEKS + 1):
        col = WEEK_START_COL_DASH + w - 1
        gs_col = week_col(w)
        ws.cell(row=rr, column=col, value=f"='Grid_Stock'!{gs_col}{grow}")

n_last_row = DATA_START_ROW + len(rm_master) - 1
# ヘッダー行が数式(Cal_Weeks参照)のため、Excelのテーブル機能(ListObject)ではなく
# 罫線・縞模様の手動書式で「テーブルらしい」見た目にする（テーブル見出しは文字列必須のため）。
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for rr2 in range(HDR_TABLE_ROW, n_last_row + 1):
    stripe = (rr2 - HDR_TABLE_ROW) % 2 == 1
    for c in range(1, WEEK_START_COL_DASH + N_WEEKS):
        cell = ws.cell(row=rr2, column=c)
        cell.border = border
        if rr2 > HDR_TABLE_ROW and stripe:
            cell.fill = PatternFill("solid", fgColor="F2F2F2")
ws.freeze_panes = f"{get_column_letter(WEEK_START_COL_DASH)}{DATA_START_ROW}"

# 安全在庫を下回った週を赤く強調
ws.conditional_formatting.add(
    f"{get_column_letter(WEEK_START_COL_DASH)}{DATA_START_ROW}:{last_col_dash}{n_last_row}",
    FormulaRule(formula=[f"{get_column_letter(WEEK_START_COL_DASH)}{DATA_START_ROW}<$D{DATA_START_ROW}"],
                fill=PatternFill("solid", fgColor="FFC7CE"))
)
# Statusが「要発注」の行を強調
ws.conditional_formatting.add(
    f"H{DATA_START_ROW}:H{n_last_row}",
    CellIsRule(operator="equal", formula=['"要発注"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006"))
)
# 入力した週(C1)と一致する列(週No行)をハイライト
ws.conditional_formatting.add(
    f"{get_column_letter(WEEK_START_COL_DASH)}{HDR_TABLE_ROW}:{last_col_dash}{n_last_row}",
    FormulaRule(
        formula=[f'{get_column_letter(WEEK_START_COL_DASH)}${HDR_TABLE_ROW}=$C$1'],
        fill=PatternFill("solid", fgColor="FFEB9C"),
    )
)
print("Dashboard: week-by-week grid for", len(rm_master), "materials x", N_WEEKS, "weeks")

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
              [week_labels[w] for w in range(1, N_WEEKS + 1)] + ["Total"]
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
    if items:
        # Excelのテーブル機能は見出し行のみ(データ0行)の範囲を許容しない
        # （その状態で保存すると、開いたときに「修復」されテーブルごと削除されてしまう）。
        # 品目が1件もない場合は、テーブル化せず見出し行の書式だけ残す。
        add_table(ws, sheet_name.replace(" ", "_"), f"B{hdr_row}:{get_column_letter(last_col_idx)}{last_row}", style="TableStyleMedium6")
    else:
        ws.cell(row=hdr_row + 1, column=2, value="(該当品目なし)")
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
    "このブックはPower BI等に読み込ませる中間テーブルではなく、これ自体が完成品です。",
    "Dashboard と PO_Draft_* が最終的な閲覧・発注画面、それ以外は計算の土台です。",
    "",
    "【週番号のルール】すべてExcelの数式で計算（Cal_WeeksのB1=AnchorYearを変えると全体が再計算）",
    "  Week1=1月1日を含む月〜日の週。年が変わると再び1にリセットされます（2026-W52の次は2027-W01）。",
    "  各シートの週見出しは3段: 1段目=月-年(Aug-26等) / 2段目=その週月曜の日付 / 3段目=週No。",
    "",
    "【普段見るシート】",
    "  Dashboard           : 原材料×週の在庫を2年分、横軸で見渡せるメイン画面（まずここ）。",
    "                        自社在庫(実績)・TTAF在庫(実績)・実績週の列で内訳も確認できます。",
    "                        C1に'2026-W23'のように入力すると該当週の列が黄色くハイライトされ、",
    "                        D1のリンクをクリックするとその週まで自動移動します（マクロ不要）。",
    "  Material_Detail     : 材料ごとに「どの中間体が・何バッチ・いくら使うか」をブロック表示（トレーサビリティ）",
    "  PO_Draft_*          : 要発注分を注文書ひな形に自動転記",
    "  T_Shipments/T_OpeningStock/T_StockCount/T_SelfStock/T_TTAFStock : 入力用",
    "",
    "  M_RawMaterials・M_BOM・PP_Grid・Grid_Stock・その他非表示シートは内部計算用です。通常は開く必要はありません。",
    "",
    "【重要】原単位・バッチ数・自社/TTAF在庫はPythonを使わず、Excel(VBA)マクロだけで更新できます。",
    "  macros/RefreshData.bas を導入し、RefreshWeeklyBatches / RefreshBOM / RefreshSelfStock /",
    "  RefreshTTAFStock を実行してください（対象ファイルを選ぶだけです。詳細はdocs/SOH_System_Guide.md、",
    "  未検証のため要動作確認）。T_Shipments・T_OpeningStock・T_StockCount等には一切触れません。",
    "",
    "【毎月の運用】",
    "  1. 「Powder & Slurry & Pgm Plan」の新しい月版でRefreshWeeklyBatchesを実行",
    "  2. 「Usage from Production Engineering」が更新されていればRefreshBOMを実行",
    "  3. 自社倉庫の現物確認を実施したらRefreshSelfStockを実行",
    "  4. CSA Reportが届いたらRefreshTTAFStockを実行し、T_Shipments のETA/着荷日/PO番号/発注日も更新",
    "  5. 棚卸を実施したらT_StockCountに実測値を追記",
    "  6. Dashboardで「要発注」を確認し、PO_Draft_*から注文書を出力",
    "  7. 月初は、前月最終週と当月頭のDashboardを見比べて在庫差異を確認（Plan Increase and Decrease /",
    "     Inventory Releasesの報告フォーマットに転記）",
    "",
    "【前提・要確認事項】詳細はdocs/SOH_System_Guide.mdを参照",
    "  - M_RawMaterials の SafetyStock_Qty と LeadTime_Weeks は仮値です。実際の安全在庫水準に置き換えてください。",
    "  - Categoryの割り当て(Chemical/Hazardous Chemical/Substrate)は入手データから機械的に推定した部分があります。要レビュー。",
    "  - 週次バッチ数はPowder & Slurry & Pgm Planの実データ（約36材料シートから抽出）を使用しています。",
]
for i, line in enumerate(readme_lines, start=1):
    ws.cell(row=i, column=1, value=line)
ws["A1"].font = Font(bold=True, size=14)
ws.column_dimensions["A"].width = 100

# ---- ナビゲーション（README上部にジャンプリンクを追加） ----
nav_targets = [
    ("Dashboard", "原材料×週の在庫（2年分・横軸で見渡せるメイン画面。まずここ）"),
    ("Material_Detail", "材料ごとの使用状況（どの材料が何に使われているか）"),
    ("PO_Draft_Chemical", "発注書ドラフト（Chemical）"),
    ("PO_Draft_Hazardous", "発注書ドラフト（Hazardous Chemical）"),
    ("PO_Draft_Substrate", "発注書ドラフト（Substrate）"),
    ("T_Shipments", "発注・着荷の入力"),
    ("T_OpeningStock", "期首在庫の入力"),
    ("T_StockCount", "棚卸実績の入力"),
]
ws.insert_rows(2, amount=len(nav_targets) + 2)
ws["A2"] = "【ジャンプ】クリックで各シートへ移動"
ws["A2"].font = Font(bold=True, size=12)
for i, (target, label) in enumerate(nav_targets, start=3):
    cell = ws.cell(row=i, column=1, value=f"▶ {target} - {label}")
    cell.hyperlink = f"#'{target}'!A1"
    cell.font = Font(color="0563C1", underline="single")

# ---- 内部処理用シートは非表示にして視認性を上げる（Dashboardが週次在庫の表示を兼ねるためGrid_Stockも非表示） ----
for sheet_name in ["Cal_Weeks", "M_Intermediates", "M_ProductMap", "Grid_Requirement", "Grid_Incoming", "Grid_Stock"]:
    if sheet_name in wb.sheetnames:
        wb[sheet_name].sheet_state = "hidden"

# ---- シートの並び順を業務で使う順に ----
order = ["README", "Dashboard", "Material_Detail", "PO_Draft_Chemical", "PO_Draft_Hazardous",
         "PO_Draft_Substrate", "T_Shipments", "T_OpeningStock", "T_StockCount",
         "T_SelfStock", "T_TTAFStock",
         "M_RawMaterials", "M_BOM", "PP_Grid",
         "Cal_Weeks", "M_Intermediates", "M_ProductMap", "Calc_Demand", "Grid_Requirement",
         "Grid_Incoming", "Grid_Stock"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else len(order))

# ---- 保存前チェック: テーブルがヘッダー行のみ(データ0行)になっていないか ----
# (この状態で保存するとExcelで開いたときに「修復」されテーブルが削除されてしまうため)
broken_tables = []
for sheet in wb.worksheets:
    for tbl in sheet.tables.values():
        min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(tbl.ref)
        if max_row <= min_row:
            broken_tables.append(f"{sheet.title}!{tbl.name} (ref={tbl.ref})")
if broken_tables:
    raise RuntimeError(
        "以下のテーブルがヘッダー行のみでデータ行がありません(Excelで開くと修復されます): "
        + ", ".join(broken_tables)
    )

wb.save(OUT_PATH)
print("Full workbook written to", OUT_PATH)
