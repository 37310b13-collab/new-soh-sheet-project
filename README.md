# new-soh-sheet-project

原材料の週次在庫（2年先まで）をロールフォワード計算し、着荷予定（ETA）との突合・発注アラート・
発注書ドラフトまでを1つのExcelブックで扱うSOH（在庫管理）システムです。生産計画・原単位そのものの
再計算はせず、既存ファイルの値を参照するだけに範囲を絞っています。

- **成果物**: [`SOH_Master.xlsx`](./SOH_Master.xlsx)
- **使い方・仕組みの詳細**: [`docs/SOH_System_Guide.md`](./docs/SOH_System_Guide.md)
- **月次更新マクロ（Python不要）**: [`macros/RefreshData.bas`](./macros/RefreshData.bas)
- **発注書エクスポート用マクロ**: [`macros/PO_Export.bas`](./macros/PO_Export.bas)
- **マスタデータ**: [`data/masters/`](./data/masters/)
- **(参考/任意) Python版の再抽出・再生成スクリプト**: [`scripts/`](./scripts/)

## 毎月の更新（Excel + VBAのみ、Python不要）

1. `SOH_Master.xlsx`を「Excel マクロ有効ブック(*.xlsm)」として保存
2. `macros/RefreshData.bas`をVBEに読み込む（Alt+F11 → 挿入 → 標準モジュール）
3. 「Powder & Slurry & Pgm Plan」の新しい月版で `RefreshWeeklyBatches` マクロを実行
4. 「Usage from Production Engineering」が更新されていれば `RefreshBOM` マクロを実行

詳しい運用手順は [`docs/SOH_System_Guide.md`](./docs/SOH_System_Guide.md) を参照してください。
「Usage from Production Engineering」「Powder & Slurry & Pgm Plan」「CSA Report」の各ファイルは
フォーマットを一切変更していません。「Plan Increase and Decrease」「Inventory June Releases」は
このシステムの計算範囲から切り離しています。

**注意**: `macros/RefreshData.bas`と`powerquery/Q_Shipments.pq`は、この開発環境にExcel/VBA/Power
Queryの実行環境がないため未検証です。ご利用のExcelでの動作確認をお願いします。
