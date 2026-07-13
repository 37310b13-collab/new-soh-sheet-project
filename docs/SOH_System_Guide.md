# SOH（在庫管理）シート システムガイド

## 1. これは何か（位置づけ）

**これは完成したExcelツールとして使うことを意図しています。**Power BIやQueryで後から整理し直すための
「中間テーブル置き場」ではありません。`Dashboard`と`PO_Draft_*`が最終的な閲覧・発注画面で、それ以外
のシートはその計算のための土台です。

顧客オーダー → 生産計画（中間体バッチ数） → 原単位展開 → 原材料所要量 → 週次在庫（2年先まで）
→ 発注書ドラフト、までを1つのExcelブック（`SOH_Master.xlsx`）内で自動計算します。Google Sheets /
Apps Scriptは使用せず、Excelの通常の数式とテーブル機能のみで構築しています。既存の
`Usage from Production Engineering` / `CSA Confirmed Order` / `Plan Increase and Decrease` の
3ファイルは**フォーマットを一切変更していません**。

## 2. シート構成と役割

| シート | 種別 | 役割 |
|---|---|---|
| README | - | ブック内の使い方サマリ |
| Cal_Weeks | マスタ | 週の定義（後述） |
| M_RawMaterials | マスタ | 原材料マスタ（コード・品名・仕入先・区分：Chemical/Hazardous Chemical/Substrate） |
| M_Intermediates | マスタ | 中間体マスタ（Slurry/Powder/Solution、バッチサイズ） |
| M_BOM | マスタ | 原単位（中間体 → 原材料 → 1バッチあたり使用量） |
| M_ProductMap | マスタ | 完成品コード ⇔ 中間体コードの紐付け（`TSP.TPP.CAT`シート由来） |
| **PP_Grid** | **入力**（将来はPower Query自動更新） | 生産計画（中間体×週のバッチ数） |
| **T_OpeningStock** | **入力** | 計算開始週の期首在庫 |
| **T_Shipments** | **入力**（将来はPower Query自動更新） | 発注〜輸送〜着荷の実績・予定 |
| **T_StockCount** | **入力** | 棚卸実測値（誤差リセット用） |
| Calc_Demand | 計算(非表示) | 原単位展開の明細（監査用） |
| Grid_Requirement / Grid_Incoming / Grid_Stock | 計算 | 原材料の週次所要量／入荷予定／在庫（2年ロールフォワード） |
| **Dashboard** | **出力（最終確認画面）** | 品目ごとの現在庫・最小在庫・要発注アラート一覧 |
| **PO_Draft_Chemical / _Hazardous / _Substrate** | **出力（最終発注画面）** | 要発注分を注文書ひな形へ自動転記 |

## 3. 週番号のルール（Cal_Weeks）

- **Week1 = 1月1日を含む月〜日の週**。年をまたぐ場合は、翌年1月1日を含む週で再び **Week1** に
  戻ります（例: `2026-W52` の次は `2027-W01`）。
- 表示・入力は必ず `Label`列（例 `2026-W01`）を使ってください。裏側の`WeekIndex`（1〜104の連番）は
  シート内部でグラフや数式の列位置を揃えるための番号で、ユーザーが直接意識する必要はありません。
- 現在のブックは2026年1月1日の週を起点に104週（≒2026年・2027年の2年分）を用意しています。年が
  変わって表示期間をずらしたい場合は、`scripts/build_soh.py`を新しい起点年で再実行してください
  （下記5.参照）。

## 4. 毎月（毎週）の更新方法 — 現状とこれから

### 現状（このバージョン）
Power Queryの自動接続はまだ組み込んでおらず、以下を手作業で行う必要があります。これは
「他のシートにある情報をSOHにも入れる」二重入力になっており、改善が必要な点として認識しています。

1. `Plan Increase and Decrease`の最新月シート（`All - MM`）の内容を`PP_Grid`に転記
2. `CSA Report`の`Shipping Schedule`の最新分を`T_Shipments`に転記
3. 棚卸実施週は`T_StockCount`に実測値を追記

### これから（Power Queryで自動化）— `powerquery/`フォルダ参照
`powerquery/Q_ProductionPlan.pq` と `powerquery/Q_Shipments.pq` に、①フォルダ内から自動で最新の
月次／週次ファイルを見つけ、②`All - MM`シートや`Shipping Schedule`シートを読み取り、③そのまま
使えるテーブルに変換する、Power QueryのMコードを用意しました。**フォーマット変更は一切不要**で、
今の運用フォルダ構成のまま自動更新できるように設計しています。

> **重要な注意点**: この環境にはPower Queryを実際に動かして検証する手段がないため、このMコードは
> **未検証**です（Excelの数式部分は実データで動作確認済みですが、こちらは違います）。まず貴社の
> Excelで動作確認をお願いします。エラーが出た場合はその内容を教えてください、すぐに修正します。

**導入手順（概要）**
1. `powerquery/Q_ProductionPlan.pq` を開き、`RootFolder`を実際のフォルダパスに書き換える
2. Excelの データ → データの取得 → 空のクエリ → 詳細エディターに貼り付け → 読み込み
3. 同様に `Q_Shipments.pq` も設定する
4. 動作確認できたら、以下の2箇所の数式を切り替える（**この置き換えをして初めて二重入力がなくなります**）

   - `Calc_Demand`の`Batches`列（F列）を全行、次の式に置き換え:
     ```
     =SUMIFS(Q_ProductionPlan[Batches],Q_ProductionPlan[Intermediate],[@Intermediate],Q_ProductionPlan[WeekStart],INDEX(Cal_Weeks[WeekStart],[@WeekIndex]))
     ```
   - `Q_Shipments`テーブルの右隣に`Effective_Week`列を追加し、`T_Shipments`と同じ式を入れる
     （着荷日があれば着荷日、なければETAから週番号を計算）:
     ```
     =MAX(1,MIN(104,INT((IF([@Received_Date]="",[@Latest_ETA],[@Received_Date])-Cal_Weeks[[#Headers],[WeekStart]])/7)+1))
     ```
     ※ 実際のセル参照は環境に合わせて調整してください（`Cal_Weeks`のWeek1開始日を起点にする考え方は
     `T_Shipments`の`Effective_Week`式と同じです）。
   - `Grid_Incoming`の各セルの参照先を`T_Shipments`→`Q_Shipments`に変更

5. 以後は データ → すべて更新（Refresh All） を押すだけで、`PP_Grid`/`T_Shipments`への手入力が
   不要になります

### 完全に自動化した場合の毎月の作業
1. データ → すべて更新（Refresh All）を押す
2. `Dashboard`で「要発注」を確認
3. `PO_Draft_*`を確認・必要なら数量を手直しして発行（マクロ使用）
4. 棚卸を実施した週だけ`T_StockCount`に実測値を追記（これは自動化の対象外です）

## 5. 自動反映の仕組み（数式エンジンの動き）

- **生産計画が変わったとき**: `PP_Grid`（またはPower Query経由の`Q_ProductionPlan`）を更新すると、
  `Calc_Demand` → `Grid_Requirement` → `Grid_Stock` → `Dashboard` → `PO_Draft_*` まで自動的に
  再計算されます。現状の「Plan Increase and Decrease」で行っている差分計算・反映作業は不要です。
- **輸入品が早着・遅着したとき**: `T_Shipments`（または`Q_Shipments`）のETA・着荷日を書き換える
  だけで、入荷が計上される週が自動的にシフトし、在庫予測に反映されます。
- **棚卸で実測とズレがあったとき**: `T_StockCount`に「原材料コード・週番号・実測値」を1行追記すると、
  その週の在庫がその値で上書きされ、以降はそこから積み上げ直されます。

## 6. ブックの再生成方法

`SOH_Master.xlsx`は`data/masters/`のCSVから`scripts/build_soh.py`で生成しています。マスタデータ
（原材料・中間体・原単位・紐付け表）が更新された場合は、CSVを更新した上で以下を再実行してください。

```
python3 scripts/build_soh.py [週数(既定104)] [出力パス] [起点年(既定2026)]
```

## 7. 動作検証結果

LibreOffice（Excel互換の検証環境）で実際に数式を再計算させ、以下を確認済みです（Power Query部分を
除く）。

- 101品目 × 104週、原単位展開59,280行を含むフル規模で、**開いて再計算するまで約19秒**
  （現状のTTAF R-Modelはファイルを開くだけで90秒以上、実運用では10分以上とのことでした）
- 全シートを通じて数式エラー（#REF!等）は0件
- 生産計画変更・ETA変更（早着遅着）・安全在庫割れ検知の自動連鎖を実データで確認済み

## 8. 要確認・要入力の項目（このブックを使う前に）

- **`M_RawMaterials`の`SafetyStock_Qty` / `LeadTime_Weeks`**: すべて仮値（0 / 4週）です。実際の
  安全在庫水準・リードタイムに置き換えてください。「要発注」判定の基準になります。
- **`M_RawMaterials`の`Category`**: Chemical / Hazardous Chemicalの割り当ては、TTAF R-Modelの
  `Chemical Release` / `Hazardous Chemical Release`シートから機械的に判定しています。実際の危険物
  区分と一致しているか確認してください。
- **Substrates（基材）**: `CSA Report`の`SUBSTRATES`シートは原材料と異なるコード体系（Lot No等）を
  使っており、今回は自動統合していません。`PO_Draft_Substrate`は雛形のみで品目が空です。
- **`PP_Grid`のシード値**: サンプルとして提供いただいた期間のみ暫定的に値が入っています。それ以外の
  週は0です。運用開始時に実際の生産計画で埋めてください（またはPower Query化）。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫（自社＋TTAF合算）を
  入力してください。
- **Power Queryスクリプト（`powerquery/*.pq`）**: 未検証です。動作確認の結果を教えてください。

## 9. 今後の拡張候補（今回は未実装）

- **生産計画バージョン管理・変更差分の自動タグ付け**: 現状の「Plan Increase and Decrease」のように、
  改定前後の差分を自動集計・変動要因として表示する機能。`PP_Grid`／`Q_ProductionPlan`の週次
  スナップショットを保存する仕組みを追加すれば実現できます。
- **基材（Substrates）の統合**: 上記8.参照。
