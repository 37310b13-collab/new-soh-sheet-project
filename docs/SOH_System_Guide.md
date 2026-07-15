# SOH（在庫管理）シート システムガイド

## 1. これは何か（位置づけ）

これは完成したExcelツールとして使うことを意図しています。`Dashboard`と`PO_Draft_*`が最終的な
閲覧・発注画面、`Material_Detail`が「どの材料が何に使われているか」のトレーサビリティ画面です。

**このブックが担う範囲は「在庫管理」に絞っています。** 生産計画や原単位そのものの再計算はせず、
既存のファイルの値をそのまま参照します。

- **顧客オーダー・生産計画の策定**: このブックの範囲外（既存の運用のまま）
- **原単位（1バッチあたり使用量）**: このブックの範囲外。「Usage from Production Engineering」から
  そのまま参照
- **週次バッチ数**: このブックの範囲外。「Powder & Slurry & Pgm Plan」（毎月改版）からそのまま参照
- **在庫のロールフォワード、着荷予定との突合、発注アラート、PO発行**: このブックの役割

「Plan Increase and Decrease」「Inventory June Releases」の2ファイルは、このブックの計算からは
**完全に切り離しています**。ただし、毎月月初に前月最終週⇔当月頭の在庫差異を報告する運用は
引き続き必要とのことなので、`Grid_Stock`の週次実績がその報告の元データとして使える設計にして
あります（週ごとの在庫が既に手元にあるので、差分を見るだけで済みます）。

## 2. データの出所

| データ | 出所ファイル | 参照方法 |
|---|---|---|
| 原材料マスタ・区分(Chemical/Hazardous/Substrate) | TTAF R-Model系ファイルの Chemical Release / Hazardous Chemical Release シート | 月次で抽出スクリプト実行 |
| 原単位（1バッチ使用量kg） | **Usage from Production Engineering**（Slurry/Powderシート） | 月次で抽出スクリプト実行 |
| 週次バッチ数 | **Powder & Slurry & Pgm Plan**（毎月改版、約36枚の材料別シート） | 月次で抽出スクリプト実行 |
| 完成品⇔中間体の紐付け | Usage from Production Engineeringの`TSP.TPP.CAT`シート | 月次で抽出スクリプト実行 |
| 着荷予定・PO情報 | CSA Reportの`Shipping Schedule`シート | 手入力 or Power Query |

## 3. シート構成と役割

| シート | 種別 | 役割 |
|---|---|---|
| README | - | ブック内の使い方サマリ |
| Cal_Weeks | マスタ | 週の定義（Week1=1月1日を含む週、年ごとにリセット） |
| M_RawMaterials | マスタ | 原材料マスタ |
| M_Intermediates | マスタ | 中間体マスタ |
| M_BOM | マスタ | 原単位（中間体→原材料、1バッチ使用量） |
| M_ProductMap | マスタ | 完成品⇔中間体の紐付け |
| PP_Grid | 入力（月次で再抽出） | 生産計画（中間体×週のバッチ数） |
| T_OpeningStock | 入力 | 起点となる期首在庫 |
| T_Shipments | 入力 | 発注〜輸送〜着荷（PO番号・**発注日**・ETA・着荷日） |
| T_StockCount | 入力 | 棚卸実測値（誤差リセット用） |
| **Material_Detail** | **出力（トレーサビリティ）** | 材料ごとに「使用中間体・バッチ数・使用量・週次合計・入荷予定・在庫」をブロック表示 |
| Calc_Demand | 計算(非表示) | 原単位展開の計算過程（監査用） |
| Grid_Requirement / Grid_Incoming / Grid_Stock | 計算 | 週次所要量／入荷予定／在庫（2年ロールフォワード） |
| **Dashboard** | **出力（最終確認画面）** | 品目ごとの現在庫・最小在庫・要発注アラート |
| **PO_Draft_Chemical / _Hazardous / _Substrate** | **出力（最終発注画面）** | 要発注分を注文書ひな形へ自動転記 |

## 4. 着荷予定(CSA Order)の入力方法

`T_Shipments`に1行追加してください。

| 列 | 内容 |
|---|---|
| RM_Code | 原材料コード |
| PO_No | 発注書番号 |
| Order_Date_発注日 | いつ発注したか（手入力。元データに発注日の情報がないため） |
| Confirmed_Qty | 確定数量 |
| Latest_ETA | 最新の到着予定日 |
| Received_Date | 実際に着荷した日（着荷前は空欄） |
| Status | ステータス |

`Effective_Week`（何週の入荷として計上するか）は自動計算されます（着荷日があれば着荷日、なければ
ETAを使用）。**ETAを書き換えるだけで、その原材料の見込み在庫（`Grid_Stock`）が自動的に更新され、
早着・遅着がそのまま反映されます。**

## 5. 毎月の更新方法

1. 「Powder & Slurry & Pgm Plan」の新しい月版ファイルを受け取ったら:
   ```
   python3 scripts/extract_from_powder_slurry_pgm_plan.py "<新しいファイルのパス>"
   ```
2. 「Usage from Production Engineering」が更新されていれば同様に:
   ```
   python3 scripts/extract_bom_from_usage_engineering.py "<ファイルのパス>"
   ```
3. 上記いずれかを実行した場合は統合:
   ```
   python3 scripts/merge_bom_sources.py
   ```
4. ブックを再生成:
   ```
   python3 scripts/build_soh.py
   ```
5. `T_Shipments`をCSA Reportの最新情報で更新（ETA・着荷日・PO番号・発注日）
6. 棚卸を実施した週は`T_StockCount`に追記
7. `Dashboard`で「要発注」を確認し、`PO_Draft_*`から注文書を発行（`macros/PO_Export.bas`）
8. 月初は、前月最終週と当月頭の`Grid_Stock`を見比べて在庫差異を確認し、従来通り
   Plan Increase and Decrease / Inventory Releasesの報告フォーマットに転記

### Power Query（任意・未検証）
`T_Shipments`（CSA Reportからの取込み）はPower Queryでも自動化できます。`powerquery/Q_Shipments.pq`
を参照してください。**この環境ではPower Queryの実行検証ができないため未検証です**。動作確認の
結果を教えてください。

なお、`PP_Grid`/`M_BOM`（Powder & Slurry & Pgm Plan、Usage from Production Engineering由来）は
シート構造が複雑（材料ごとに約40枚、繰り返しブロック構造）なため、Power QueryよりもPythonスクリプト
（上記の`scripts/extract_*.py`、実データで動作検証済み）での自動化を推奨します。

## 6. 自動反映の仕組み

- **生産計画が変わったとき**: 新しい月版の「Powder & Slurry & Pgm Plan」を上記4章の手順で取り込むと、
  `Calc_Demand` → `Grid_Requirement` → `Grid_Stock` → `Dashboard` → `PO_Draft_*` → `Material_Detail`
  まで自動的に再計算されます。
- **輸入品が早着・遅着したとき**: `T_Shipments`のETA・着荷日を書き換えるだけで、入荷が計上される週が
  自動的にシフトし、在庫予測に反映されます。
- **棚卸で実測とズレがあったとき**: `T_StockCount`に「原材料コード・週番号・実測値」を1行追記すると、
  その週の在庫がその値で上書きされ、以降はそこから積み上げ直されます。

## 7. 動作検証結果

LibreOffice（Excel互換の検証環境）で実際に数式を再計算させ、以下を確認済みです（Power Query部分を
除く）。

- 101品目 × 104週、原単位展開73,000行超を含むフル規模で、**開いて再計算するまで約24〜33秒**
  （現状のTTAF R-Model系ファイルは開くだけで90秒以上、実運用では10分以上とのことでした）
- 全19シートを通じて数式エラー（#REF!等）は0件
- **`Grid_Requirement`の値が、実際の「Powder & Slurry & Pgm Plan」のAS-200シートに記載された
  実際の週次使用量（例: 172.121, 42.812, 86.061 kg等）と完全一致することを確認済み**
  （BOM(Usage from Production Engineeringから独立抽出)×週次バッチ数(Powder & Slurry & Pgm Planから
  独立抽出)を掛け合わせた結果が、元ファイルの計算結果と一致＝抽出・計算ロジックが正しいことの
  裏付けになります）
- 生産計画変更・ETA変更（早着遅着）・安全在庫割れ検知の自動連鎖を実データで確認済み

## 8. 要確認・要入力の項目

- **`M_RawMaterials`の`SafetyStock_Qty` / `LeadTime_Weeks`**: すべて仮値（0 / 4週）です。実際の
  安全在庫水準・リードタイムに置き換えてください。「要発注」判定の基準になります。
- **`M_RawMaterials`の`Category`**: Chemical / Hazardous Chemicalの割り当ては機械的に判定した
  ものです。実際の危険物区分と一致しているか確認してください。
- **Substrates（基材）**: 原材料と異なるコード体系（Lot No等）を使っており、今回は自動統合して
  いません。`PO_Draft_Substrate`は雛形のみで品目が空です。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫（自社＋TTAF合算）を
  入力してください。
- **`T_Shipments`の`Order_Date`（発注日）**: 元データに発注日の情報がないため空欄です。手入力
  してください。
- **原単位・バッチ数の未解決品目**: 抽出スクリプト実行時に「unresolved」「unmapped」として
  表示される品目名は、命名の揺れ等で自動マッチできなかったものです（例: "10H", "20P", "AMMONIA"
  等）。`data/masters/`のCSVを直接確認し、必要に応じて手動で補完してください。
- **Power Queryスクリプト（`powerquery/Q_Shipments.pq`）**: 未検証です。動作確認の結果を教えて
  ください。

## 9. 今後の拡張候補（今回は未実装）

- **基材（Substrates）の統合**: 上記8.参照。
- **Min/Max（週数ベースの安全在庫）モデル**: 現状はSafetyStock_Qtyの単一しきい値のみですが、
  「N週分の使用量」を基準にしたMin/Max運用に拡張することも可能です。
