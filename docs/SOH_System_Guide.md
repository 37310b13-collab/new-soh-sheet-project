# SOH（在庫管理）シート システムガイド

## 1. これは何か

顧客オーダー → 生産計画（中間体バッチ数） → 原単位展開 → 原材料所要量 → 週次在庫（2年先まで）
→ 発注書ドラフト、までを1つのExcelブック（`SOH_Master.xlsx`）内で自動計算する仕組みです。

Google Sheets / Apps Script は使用せず、**Excelの通常の数式とテーブル機能のみ**で構築しています
（VBAマクロは任意で追加可能。Power Query／Power Pivotは未使用ですが、将来的な拡張先として設計上
差し替え可能にしてあります）。既存の `Usage from Production Engineering` / `CSA Confirmed Order` /
`Plan Increase and Decrease` の3ファイルは**フォーマットを一切変更していません**。今回のブックは
これらのファイルから読み取った実データをもとにマスタを構築した独立ファイルです。

## 2. シート構成と役割

| シート | 種別 | 役割 |
|---|---|---|
| README | - | ブック内の使い方サマリ |
| M_RawMaterials | マスタ | 原材料マスタ（コード・品名・仕入先・区分：Chemical/Hazardous Chemical/Substrate） |
| M_Intermediates | マスタ | 中間体マスタ（Slurry/Powder/Solution、バッチサイズ） |
| M_BOM | マスタ | 原単位（中間体 → 原材料 → 1バッチあたり使用量） |
| M_ProductMap | マスタ | 完成品コード ⇔ 中間体コードの紐付け（`TSP.TPP.CAT`シート由来） |
| **PP_Grid** | **入力** | 生産計画（中間体×週のバッチ数）。ここが起点データです |
| **T_OpeningStock** | **入力** | 計算開始週の期首在庫 |
| **T_Shipments** | **入力** | 発注〜輸送〜着荷の実績・予定（PO番号・ETA・着荷日・ステータス） |
| **T_StockCount** | **入力** | 棚卸実測値（誤差リセット用） |
| Calc_Demand | 計算(非表示) | 原単位展開の明細（監査用、570品目×週） |
| Grid_Requirement | 計算 | 原材料の週次所要量 |
| Grid_Incoming | 計算 | 原材料の週次入荷予定 |
| Grid_Stock | 計算 | 原材料の週次在庫（2年=104週ロールフォワード） |
| Dashboard | 出力 | 品目ごとの現在庫・最小在庫・要発注アラート一覧（1画面で全体把握） |
| PO_Draft_Chemical / _Hazardous / _Substrate | 出力 | 要発注分を注文書ひな形(Chemical Release形式)へ自動転記 |

## 3. 自動反映の仕組み

- **生産計画が変わったとき**: `PP_Grid` の該当セルを書き換えるだけで、`Calc_Demand` →
  `Grid_Requirement` → `Grid_Stock` → `Dashboard` → `PO_Draft_*` まで自動的に再計算されます。
  現状の「Plan Increase and Decrease」で行っている差分計算・反映作業は不要になります。
- **輸入品が早着・遅着したとき**: `T_Shipments` の `Latest_ETA`（または`Received_Date`）を書き換える
  だけで、`Effective_Week`（入荷が計上される週）が自動的にシフトし、その週以降の在庫予測に反映され
  ます。
- **棚卸で実測とズレがあったとき**: `T_StockCount` に「原材料コード・週番号・実測値」を1行追記すると、
  その週の在庫がその値で上書きされ、以降はそこから積み上げ直されます（誤差の累積を防ぎます）。

## 4. 週次の運用手順（想定）

1. 生産計画が確定・変更されたら `PP_Grid` を更新
2. CSA Reportの最新情報で `T_Shipments` のETA・着荷日・ステータスを更新
3. 棚卸を実施した週は `T_StockCount` に実測値を追記
4. `Dashboard` で「要発注」になっている品目を確認
5. `PO_Draft_Chemical` / `_Hazardous` / `_Substrate` の内容を確認・必要に応じて数量を手直し
6. マクロ（`macros/PO_Export.bas`）でPO_Draftシートをスナップショットとして書き出し、発注書として送付

## 5. 動作検証結果

LibreOffice（Excel互換の検証環境）で実際に数式を再計算させ、以下を確認済みです。

- 101品目 × 104週（原材料）、570品目 × 104週（原単位展開、59,280行）を含むフル規模の状態で、
  **開いて再計算するまで約19秒**（現状のTTAF R-Modelはファイルを開くだけで90秒以上、実運用では
  10分以上とのことでした）
- 全シートを通じて数式エラー（#REF!等）は0件
- 生産計画変更 → 原材料所要量・在庫への自動反映を実データで確認済み
- ETA変更（早着・遅着）→ 入荷計上週の自動シフトを実データで確認済み
- 安全在庫割れ検知 → Dashboardアラート → PO_Draftへの自動転記を確認済み

## 6. 要確認・要入力の項目（このブックを使う前に）

このブックは実データから自動抽出したマスタで構築していますが、以下は仮値・推定です。実運用前に
必ず見直してください。

- **`M_RawMaterials` の `SafetyStock_Qty` / `LeadTime_Weeks`**: すべて仮値（0 / 4週）です。実際の
  安全在庫水準・リードタイムに置き換えてください。これが「要発注」判定の基準になります。
- **`M_RawMaterials` の `Category`**: Chemical / Hazardous Chemical の割り当ては、TTAF R-Modelの
  `Chemical Release` / `Hazardous Chemical Release` シートから機械的に判定しています。実際の危険物
  区分と一致しているか確認してください。
- **Substrates（基材）**: CSA Reportの`SUBSTRATES`シートは原材料と異なるコード体系（Lot No等）を
  使っており、今回は自動統合していません。`PO_Draft_Substrate`は雛形のみで品目が空です。基材の
  在庫管理を同じ仕組みに統合する場合は、コード体系のすり合わせが別途必要です。
- **`PP_Grid`（生産計画）のシード値**: アップロードいただいた`Plan_Increase_and_decrease`ファイルの
  サンプル週（2026年5〜6月ごろ）のみを暫定的に流し込んでいます。それ以外の週は0です。運用開始時に
  実際の生産計画で全期間を埋めてください。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫（自社＋TTAF合算）を入力
  してください。
- **`T_Shipments`（発注・輸送実績）**: CSA Reportの`Shipping Schedule`から、計算対象期間に関係する
  分（起点週の14日前以降、または未着荷分）のみ111件を取り込んでいます。継続運用では新しいPOを
  この表に追記していく想定です。

## 7. 今後の拡張候補（今回は未実装）

- **生産計画バージョン管理・変更差分の自動タグ付け**: 現状の「Plan Increase and Decrease」のように、
  改定前後の差分を自動集計・変動要因として表示する機能。`PP_Grid`の週次スナップショットを保存する
  仕組みを追加すれば実現できます。
- **Power Query によるデータ取込の自動化**: 現状は`T_Shipments`等を手動更新する前提ですが、
  Power Queryを使えば`CSA Report`ファイルを直接参照して自動更新できます（フォーマット変更は不要）。
  将来的にファイル数・品目数が増えて重くなった場合の対応策として有効です。
- **基材（Substrates）の統合**: 上記6.参照。
