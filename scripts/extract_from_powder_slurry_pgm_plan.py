"""
Powder & Slurry & Pgm Plan (毎月改版, 例: "...- 2026-06-V02 - Delinked.xlsx") から
週次バッチ数(weekly_batches.csv)と原単位の補完データ(bom_from_plan.csv)を抽出する。

このファイルは~40枚の「材料ごとのシート」を持つ (例: "AS-200")。各シートは:
  row1: 月ラベル, row2: 材料名 + 週初日(日付, col D以降), row3: 材料コード(CHEM-xxxx) + 週番号,
  row4以降: [中間体名, "No. of batches", 週次バッチ数...] と
            [(空欄), "Usage per batch in kg", 1バッチ使用量(kg), 週次使用量...] の繰り返し。

使い方:
    python3 scripts/extract_from_powder_slurry_pgm_plan.py <入力xlsxパス> [出力ディレクトリ]

出力先は既定で data/masters/ (weekly_batches.csv, bom_from_plan.csv を上書き)。
毎月新しいバージョンのファイルが発行されたら、このスクリプトを再実行して
scripts/build_soh.py を再度動かすことで、SOH_Master.xlsx に反映される。
"""
import openpyxl, csv, datetime, sys, os

EXCLUDE_SUBSTR = ["substr", "gpf"]


def normalize_inter(raw, canonical, canonical_upper):
    raw = raw.strip()
    if raw in canonical:
        return raw
    if raw.upper() in canonical_upper:
        return canonical_upper[raw.upper()]
    cand = "SOL-" + raw.upper()
    if cand in canonical_upper:
        return canonical_upper[cand]
    return raw


def extract(src_path, out_dir):
    canonical = set()
    inter_master_path = os.path.join(out_dir, "intermediate_master.csv")
    if os.path.exists(inter_master_path):
        with open(inter_master_path) as f:
            for r in csv.DictReader(f):
                canonical.add(r["Intermediate"])
    canonical_upper = {c.upper(): c for c in canonical}

    wb = openpyxl.load_workbook(src_path, read_only=True, data_only=True)

    batches = {}
    rates = {}
    mat_codes = {}
    unmapped = set()

    for sn in wb.sheetnames:
        if any(x in sn.lower() for x in EXCLUDE_SUBSTR):
            continue
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 5:
            continue
        row2, row3 = rows[1], rows[2]
        date_cols = {ci: v.date() for ci, v in enumerate(row2) if ci >= 3 and isinstance(v, datetime.datetime)}
        if len(date_cols) < 3:
            continue
        mat_code = row3[1] if len(row3) > 1 else None
        if not (isinstance(mat_code, str) and mat_code.startswith("CHEM")):
            continue
        mat_codes[sn] = mat_code

        current_inter = None
        for ri in range(3, len(rows)):
            r = rows[ri]
            if len(r) < 3:
                continue
            label = r[2]
            if isinstance(label, str) and label.strip().lower().startswith("no. of batches"):
                raw_inter = r[1]
                if not raw_inter or not isinstance(raw_inter, str):
                    current_inter = None
                    continue
                current_inter = normalize_inter(raw_inter, canonical, canonical_upper)
                if current_inter not in canonical:
                    unmapped.add(raw_inter)
                for ci, wk_date in date_cols.items():
                    if ci >= len(r):
                        continue
                    val = r[ci] or 0
                    key = (current_inter, wk_date)
                    if key not in batches:
                        batches[key] = float(val)
            elif isinstance(r[1], str) and r[1].strip().lower() == "usage per batch in kg" and current_inter:
                rate = r[2] if len(r) > 2 else None
                if isinstance(rate, (int, float)):
                    rates[(current_inter, mat_code)] = float(rate)

    print("Material sheets used:", len(mat_codes))
    print("Unique (Intermediate,Week) batches:", len(batches))
    print("Unique (Intermediate,RM_Code) rates:", len(rates))
    print("Unmapped intermediate names (kept as-is):", sorted(unmapped))

    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "weekly_batches.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Intermediate", "WeekStart", "Batches"])
        for (inter, wk), qty in sorted(batches.items()):
            if inter in ("-", "", "No of times used"):
                continue
            w.writerow([inter, wk.isoformat(), qty])

    with open(os.path.join(out_dir, "bom_from_plan.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Intermediate", "RM_Code", "RM_Qty_Per_Batch"])
        for (inter, rm), rate in sorted(rates.items()):
            if inter in ("-", "", "No of times used"):
                continue
            w.writerow([inter, rm, rate])


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "masters"
    )
    extract(src, out)
