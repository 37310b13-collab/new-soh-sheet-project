# SOH（在庫管理）シート システムガイド

## 1. これは何か（位置づけ）

`Dashboard`が最終的にすべての原材料の在庫を確認するメイン画面、`Material_Detail`がその内訳
（どの材料が何に使われているか）を見る画面、`PO_Draft_*`が発注書の発行画面です。それ以外の
シート（`M_RawMaterials`、`M_BOM`、`PP_Grid`、`Grid_Stock`等）は計算のための土台で、通常は開く
必要はありません（タブは非表示にしてあります）。

**`Dashboard`は、原材料×週の在庫を2年分・横軸で見渡せる設計にしています。** 生産計画から
「Week◯の在庫がいくつ、翌週はいくつ…」と2年先まで追えます。週の見出しは3段（1段目=月-年、
2段目=その週月曜の日付、3段目=週No）で、すべてExcelの数式（`Cal_Weeks`シート参照）で計算
されています。`Dashboard`のC1セルに`2026-W23`のように入力すると、該当する週の列が黄色く
ハイライトされます。

**このブックが担う範囲は「在庫管理」に絞っています。** 生産計画や原単位そのものの再計算はせず、
既存のファイルの値を月次でマクロ経由で取り込みます。

- **顧客オーダー・生産計画の策定**: このブックの範囲外
- **原単位（1バッチあたり使用量）**: 「Usage from Production Engineering」からマクロで取り込み
- **週次バッチ数**: 「Powder & Slurry & Pgm Plan」（毎月改版）からマクロで取り込み
- **在庫のロールフォワード、着荷予定との突合、発注アラート、PO発行**: このブックの役割

「Plan Increase and Decrease」「Inventory June Releases」はこのブックの計算から切り離しています。
月初の在庫差異報告（前月最終週⇔当月頭）には、`Dashboard`の週次実績を元データとしてご利用ください。

## 2. Python不要・Excel(VBA)だけで完結する更新の仕組み

`macros/RefreshData.bas` に2つのマクロを用意しています。どちらも、対象ファイルを選ぶだけで
該当シートの値だけを更新し、**それ以外の入力済みデータ（T_Shipments・T_OpeningStock・
T_StockCount・安全在庫の設定値など）には一切触れません。**

| マクロ | 対象ファイル | 更新するシート |
|---|---|---|
| `RefreshWeeklyBatches` | Powder & Slurry & Pgm Plan（毎月） | PP_Grid |
| `RefreshBOM` | Usage from Production Engineering（改版時） | M_BOM |

**導入方法**
1. `SOH_Master.xlsx`を「名前を付けて保存」→ファイルの種類を「Excel マクロ有効ブック(*.xlsm)」にする
2. Alt+F11 でVBEを開く → 「ファイル」→「ファイルのインポート」→ `macros/RefreshData.bas`を選択
   （`.bas`ファイルを直接選ぶことで、モジュール名も含めて正しく読み込まれます）
   - コードをコピー＆貼り付けする場合は、**1行目の`Attribute VB_Name = "..."`を必ず削除してから**
     貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイルエラーになります
     （標準モジュールを挿入→中身を貼り付け、の手順を使う場合はこの点にご注意ください）
3. `macros/JumpToWeek.bas`・`macros/PO_Export.bas`も同様にインポート
4. Alt+F8 → `RefreshWeeklyBatches`を選択して実行 → ファイル選択ダイアログで最新の
   「Powder & Slurry & Pgm Plan」ファイルを選ぶ
5. 同様に`RefreshBOM`も必要なタイミングで実行
6. お好みで、シート上に図形を配置し「マクロの登録」でボタン化すると次回以降クリックだけで済みます

**重要な注意点**: この環境ではVBAを実際にExcel上で実行して検証することができません（数式エンジンは
実データで動作確認済みですが、このマクロコードは未検証です）。まず貴社のExcelでテスト用のコピーに
対して動作確認をお願いします。エラーが出た場合は、エラーメッセージの内容を教えてください。

実行結果はメッセージボックスで「更新セル数」「新規追加した中間体/組み合わせ数」「未解決の材料名」
が表示されます。全く新しい中間体×原材料の組み合わせが増えた場合、`Calc_Demand`（原単位展開の
明細表）はその場では自動拡張されません。その場合はご連絡いただければ再生成します（通常の月次更新
＝既存の組み合わせのバッチ数・使用量の変動には対応済みです）。

### なぜExcel外部参照やPower Queryではないのか
- **外部参照式**（別ブックのセルを直接参照する数式）は、この規模（原単位・バッチ数あわせて数千件）
  だと参照が多くなりすぎ、動作が重くなる懸念があったため採用しませんでした。
- **Power Query**は対象ファイルの構造が複雑（材料ごとに約40回繰り返すブロック構造）なため、
  Mコードが複雑になり、この環境で検証できないリスクが高いと判断しました。
- VBAは同じ理由で「未検証」ではありますが、通常のプログラムコードとして読みやすく、動作の予測が
  つきやすいため、この方式を採用しています。

`powerquery/Q_Shipments.pq`（CSA ReportのShipping Schedule取込み）だけは、シート構造がシンプルな
ため引き続きPower Query案として残しています（任意・未検証）。

`macros/JumpToWeek.bas`（任意）を導入すると、Dashboard C1に週を入力した状態で`GoToWeek`マクロを
実行することで、該当週の列まで自動スクロールします（黄色ハイライトだけでも十分に見つけやすいため、
必須ではありません）。

### Pythonスクリプトについて（参考・任意）
`scripts/`にPython版の抽出・再生成スクリプトも残しています。Python環境がある場合はこちらでも
更新できますが、**ブックをまるごと再生成する**ため、通常の月次運用にはVBAマクロ（上記）を
お使いください。

## 3. シート構成と役割

| シート | 種別 | 役割 |
|---|---|---|
| README | - | 使い方サマリ・各シートへのジャンプリンク |
| Dashboard | 出力（最終確認画面） | 原材料×週の在庫を2年分・横軸で表示。安全在庫割れは赤で強調、C1入力で該当週を黄色ハイライト |
| Material_Detail | 出力（トレーサビリティ） | 材料ごとに「使用中間体・バッチ数・使用量・週次合計」をブロック表示（在庫・入荷予定はDashboard参照） |
| PO_Draft_Chemical / _Hazardous / _Substrate | 出力（発注書） | 要発注分を注文書ひな形へ自動転記 |
| T_Shipments | 入力 | 発注〜輸送〜着荷（PO番号・発注日・ETA・着荷日） |
| T_OpeningStock | 入力 | 起点となる期首在庫 |
| T_StockCount | 入力 | 棚卸実測値（誤差リセット用） |
| M_RawMaterials | マスタ | 原材料マスタ・安全在庫設定 |
| M_BOM | マスタ（マクロ更新） | 原単位。RefreshBOMで更新 |
| PP_Grid | マスタ（マクロ更新） | 週次バッチ数。RefreshWeeklyBatchesで更新 |
| (非表示) Cal_Weeks / M_Intermediates / M_ProductMap / Calc_Demand / Grid_Requirement / Grid_Incoming | 内部計算 | 通常は開く必要なし |

## 4. 着荷予定(CSA Order)の入力方法

`T_Shipments`に1行追加してください。ETAを入力した週に、その数量が見込み在庫として自動反映
されます。着荷日が確定したら`Received_Date`を入力してください（未確定分はETAが使われます）。
`PO_No`・`Order_Date_発注日`もあわせて記録できます。

## 5. 毎月の運用フロー

1. 「Powder & Slurry & Pgm Plan」の新しい月版を受け取ったら`RefreshWeeklyBatches`を実行
2. 「Usage from Production Engineering」が更新されていれば`RefreshBOM`を実行
3. `T_Shipments`をCSA Reportの最新情報で更新（ETA・着荷日・PO番号・発注日）
4. 棚卸を実施した週は`T_StockCount`に追記
5. `Dashboard`で「要発注」を確認し、`PO_Draft_*`から注文書を発行（`macros/PO_Export.bas`）
6. 月初は、前月最終週と当月頭の`Dashboard`を見比べて在庫差異を確認し、従来通り
   Plan Increase and Decrease / Inventory Releasesの報告フォーマットに転記

## 6. 自動反映の仕組み

- **生産計画が変わったとき**: `RefreshWeeklyBatches`実行後、`Grid_Requirement` → `Grid_Stock` →
  `Dashboard` → `PO_Draft_*` → `Material_Detail`まで自動的に再計算されます。
- **原単位が変わったとき**: `RefreshBOM`実行後、同様に自動反映されます（既存の中間体×原材料の
  組み合わせであれば、`Calc_Demand`の再生成なしで反映されます）。
- **輸入品が早着・遅着したとき**: `T_Shipments`のETA・着荷日を書き換えるだけで、入荷が計上される
  週が自動的にシフトします。
- **棚卸で実測とズレがあったとき**: `T_StockCount`に1行追記すると、その週の在庫が上書きされ、
  以降はそこから積み上げ直されます。

## 7. 動作検証結果

LibreOffice（Excel互換の検証環境）で実際に数式を再計算させ、以下を確認済みです（VBAマクロと
Power Query部分を除く。両者はこの環境で実行できないため未検証です）。

- 101品目 × 104週、原単位展開73,000行超を含むフル規模で、**開いて再計算するまで約25〜70秒**
  （現状のTTAF R-Model系ファイルは開くだけで90秒以上、実運用では10分以上とのことでした）
- 全19シートを通じて数式エラー（#REF!等）は0件
- `Grid_Requirement`の値が、実際の「Powder & Slurry & Pgm Plan」の元シートに記載された週次使用量
  と完全一致することを確認済み
- 生産計画変更・ETA変更（早着遅着）・安全在庫割れ検知の自動連鎖を実データで確認済み
- `M_BOM`・`PP_Grid`にVBAで行を追加しても、`Calc_Demand`のINDEX/MATCH・SUMIFS参照が正しい値を
  拾えるよう設計（数式ロジックはExcelの標準的な参照方式のため動作は確実ですが、VBA自体の実行は
  未検証です）

## 8. 要確認・要入力の項目

- **`M_RawMaterials`の`SafetyStock_Qty` / `LeadTime_Weeks`**: 仮値です。実際の安全在庫水準に
  置き換えてください。「要発注」判定の基準になります。
- **`M_RawMaterials`の`Category`**: 機械的に判定したものです。実際の危険物区分と一致しているか
  確認してください。
- **Substrates（基材）**: 原材料と異なるコード体系（Lot No等）のため今回は未統合です。
  `PO_Draft_Substrate`は雛形のみで品目が空です。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫を入力してください。
- **`RefreshData.bas`・`Q_Shipments.pq`**: 未検証です。動作確認の結果を教えてください。

## 9. 今後の拡張候補

- **基材（Substrates）の統合**: 上記8.参照。
- **Min/Max（週数ベースの安全在庫）モデル**: 現状は単一しきい値のみですが、「N週分の使用量」を
  基準にしたMin/Max運用に拡張することも可能です。
