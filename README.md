# new-soh-sheet-project

原材料の週次在庫（2年先まで）をロールフォワード計算し、着荷予定（ETA）との突合・発注アラート・
発注書ドラフトまでを1つのExcelブックで扱うSOH（在庫管理）システムです。生産計画・原単位そのものの
再計算はせず、既存ファイルの値を参照するだけに範囲を絞っています。

- **成果物**: [`SOH_Master.xlsx`](./SOH_Master.xlsx)
- **使い方・仕組みの詳細**: [`docs/SOH_System_Guide.md`](./docs/SOH_System_Guide.md)
- **発注書エクスポート用マクロ**: [`macros/PO_Export.bas`](./macros/PO_Export.bas)
- **マスタデータ**: [`data/masters/`](./data/masters/)
- **月次の再抽出・再生成スクリプト**: [`scripts/`](./scripts/)

## 毎月の更新

```
python3 scripts/extract_from_powder_slurry_pgm_plan.py "<最新のPowder & Slurry & Pgm Planファイル>"
python3 scripts/extract_bom_from_usage_engineering.py "<最新のUsage from Production Engineeringファイル>"  # 更新があれば
python3 scripts/merge_bom_sources.py
python3 scripts/build_soh.py
```

詳しい運用手順は [`docs/SOH_System_Guide.md`](./docs/SOH_System_Guide.md) を参照してください。
「Usage from Production Engineering」「Powder & Slurry & Pgm Plan」「CSA Report」の各ファイルは
フォーマットを一切変更していません。「Plan Increase and Decrease」「Inventory June Releases」は
このシステムの計算範囲から切り離しています。
