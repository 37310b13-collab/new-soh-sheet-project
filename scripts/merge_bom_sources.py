"""
bom_from_usage_engineering.csv（主, Usage from Production Engineering由来）と
bom_from_plan.csv（補完, Powder & Slurry & Pgm Plan由来）を統合して bom.csv を作る。
また、新しく見つかった中間体/完成品(Cat)コードを intermediate_master.csv に、
新しく見つかったsubstrateコードを rm_master.csv に追加する。

使い方:
    python3 scripts/merge_bom_sources.py [data/mastersディレクトリ]

extract_bom_from_usage_engineering.py と extract_from_powder_slurry_pgm_plan.py を
両方実行した後に、このスクリプトを実行してください。
"""
import csv, sys, os

BAD = {"No of times used", "-", ""}


def load(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main(out_dir):
    primary = load(os.path.join(out_dir, "bom_from_usage_engineering.csv"))
    supplement = load(os.path.join(out_dir, "bom_from_plan.csv"))

    seen = set((r["Intermediate"], r["RM_Code"]) for r in primary)
    merged = list(primary)
    for r in supplement:
        if r["Intermediate"] in BAD:
            continue
        key = (r["Intermediate"], r["RM_Code"])
        if key not in seen:
            merged.append(r)
            seen.add(key)

    with open(os.path.join(out_dir, "bom.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["Intermediate", "RM_Code", "RM_Qty_Per_Batch"])
        w.writeheader()
        for r in merged:
            w.writerow(r)
    print("bom.csv written:", len(merged), "rows")

    # --- substrate_master.csv を rm_master.csv に統合(Category="Substrate") ---
    # 既存の行(RM_Code一致)はCategoryのみ"Substrate"に補正し、他の項目(TTAF_Code等)は
    # 既存の値(CSA Report等より正確な出所)を優先して残す。未登録のコードのみ新規追加する。
    substrate_master = load(os.path.join(out_dir, "substrate_master.csv"))
    rm_path = os.path.join(out_dir, "rm_master.csv")
    rm_master = load(rm_path)
    rm_codes = set(r["RM_Code"] for r in rm_master)
    substrate_codes = set(r["RM_Code"] for r in substrate_master)

    updated_rm = 0
    for r in rm_master:
        if r["RM_Code"] in substrate_codes and r["Category"] != "Substrate":
            r["Category"] = "Substrate"
            updated_rm += 1

    added_rm = []
    for r in substrate_master:
        if r["RM_Code"] in rm_codes:
            continue
        added_rm.append({
            "RM_Code": r["RM_Code"], "TTAF_Code": "", "Description": r["Description"],
            "Supplier": "", "Category": "Substrate",
        })
        rm_codes.add(r["RM_Code"])

    with open(rm_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["RM_Code", "TTAF_Code", "Description", "Supplier", "Category"])
        w.writeheader()
        for r in rm_master:
            w.writerow(r)
        for r in added_rm:
            w.writerow(r)
    print("rm_master.csv: added", len(added_rm), "new substrate codes,",
          "updated Category for", updated_rm, "existing codes")

    # --- 新しい中間体/完成品(Cat)コードを intermediate_master.csv に追加 ---
    inter_path = os.path.join(out_dir, "intermediate_master.csv")
    inter_master = load(inter_path)
    old_names = set(r["Intermediate"] for r in inter_master)
    old_upper = {n.upper(): n for n in old_names}

    batches = load(os.path.join(out_dir, "weekly_batches.csv"))
    candidates = set(r["Intermediate"] for r in merged) | set(r["Intermediate"] for r in batches)

    # substrateコード(RM_Code)と対になっているIntermediateは完成品(Cat)コードとみなす
    cat_names = set(r["Intermediate"] for r in merged if r["RM_Code"] in substrate_codes) | \
        set(r["Intermediate"] for r in supplement if r["RM_Code"] in substrate_codes)

    added = []
    for name in sorted(candidates):
        if name in BAD or name in old_names or name.upper() in old_upper:
            continue
        added.append(name)

    with open(inter_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["Intermediate", "Type", "PGM", "Batch_Size"])
        w.writeheader()
        for r in inter_master:
            w.writerow(r)
        for name in added:
            if name in cat_names:
                t = "Cat"
            elif name.upper().startswith("SOL"):
                t = "Solution"
            elif name.upper().startswith("TPP"):
                t = "Powder"
            elif name.upper().startswith("TSP"):
                t = "Slurry"
            else:
                t = "Unknown"
            w.writerow({"Intermediate": name, "Type": t, "PGM": "", "Batch_Size": 1})
    print("intermediate_master.csv: added", len(added), "new codes")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "masters"
    )
    main(out)
