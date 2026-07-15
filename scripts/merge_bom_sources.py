"""
bom_from_usage_engineering.csv（主, Usage from Production Engineering由来）と
bom_from_plan.csv（補完, Powder & Slurry & Pgm Plan由来）を統合して bom.csv を作る。
また、新しく見つかった中間体コードを intermediate_master.csv に追加する。

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

    inter_path = os.path.join(out_dir, "intermediate_master.csv")
    inter_master = load(inter_path)
    old_names = set(r["Intermediate"] for r in inter_master)
    old_upper = {n.upper(): n for n in old_names}

    batches = load(os.path.join(out_dir, "weekly_batches.csv"))
    candidates = set(r["Intermediate"] for r in merged) | set(r["Intermediate"] for r in batches)

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
            if name.upper().startswith("SOL"):
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
