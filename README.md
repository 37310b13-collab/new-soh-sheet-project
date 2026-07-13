# new-soh-sheet-project

顧客オーダー → 生産計画 → 原単位展開 → 原材料在庫（2年先まで、週次）→ 発注書ドラフト、を
1つのExcelブックで自動計算するSOH（在庫管理）システムです。

- **成果物**: [`SOH_Master.xlsx`](./SOH_Master.xlsx)
- **使い方・仕組みの詳細**: [`docs/SOH_System_Guide.md`](./docs/SOH_System_Guide.md)
- **発注書エクスポート用マクロ**: [`macros/PO_Export.bas`](./macros/PO_Export.bas)
- **マスタデータ／再生成スクリプト**: [`data/masters/`](./data/masters/), [`scripts/build_soh.py`](./scripts/build_soh.py)

`SOH_Master.xlsx` は `python3 scripts/build_soh.py [週数] [出力パス]` で `data/masters/` の
CSVから再生成できます（デフォルトは104週=2年）。既存の `Usage from Production Engineering` /
`CSA Confirmed Order` / `Plan Increase and Decrease` の各ファイルはフォーマット変更していません。
