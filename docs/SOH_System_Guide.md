# SOH（在庫管理）シート システムガイド

## 1. これは何か（位置づけ）

`Dashboard`が最終的にすべての原材料の在庫を確認するメイン画面、`Material_Detail`がその内訳
（どの材料が何に使われているか）を見る画面、`PO_Draft_*`が発注書の発行画面です。それ以外の
シート（`M_RawMaterials`、`M_BOM`、`PP_Grid`、`Grid_Stock`等）は計算のための土台で、通常は開く
必要はありません（タブは非表示にしてあります）。

**`Dashboard`は、原材料×週の在庫を2年分・横軸で見渡せる設計にしています。** 生産計画から
「Week◯の在庫がいくつ、翌週はいくつ…」と2年先まで追えます。週の見出しは3段（1段目=月-年、
2段目=その週月曜の日付、3段目=週No）で、すべてExcelの数式（`Cal_Weeks`シート参照）で計算
されています。`Dashboard`のC1セルに`W23`のように入力すると（現在年の週Noとして検索）、
`Status`列のすぐ右（I列）にその週の在庫が常時ピン留め表示され、該当する週の列も黄色く
ハイライトされます。スクロールや別画面へのジャンプは不要です。

**このブックが担う範囲は「在庫管理」に絞っています。** 生産計画や原単位そのものの再計算はせず、
既存のファイルの値を月次でマクロ経由で取り込みます。

- **顧客オーダー・生産計画の策定**: このブックの範囲外
- **原単位（1バッチあたり使用量）**: 「Usage from Production Engineering」からマクロで取り込み
- **週次バッチ数**: 「Powder & Slurry & Pgm Plan」（毎月改版）からマクロで取り込み
- **自社倉庫の在庫実績**: 「Raw materials daily check」（現物確認シート）からマクロで取り込み
- **TTAF倉庫の在庫実績**: 「CSA Report」の`PIVOT SOH TTAF`シートからマクロで取り込み
- **在庫のロールフォワード、着荷予定との突合、発注アラート、PO発行**: このブックの役割

「Plan Increase and Decrease」「Inventory June Releases」はこのブックの計算から切り離しています。
月初の在庫差異報告（前月最終週⇔当月頭）には、`Dashboard`の週次実績を元データとしてご利用ください。

**自社在庫・TTAF在庫の内訳**: `Dashboard`に「自社在庫(実績)」「TTAF在庫(実績)」「実績週」列があります。
これは各原材料について、実績データが届いている**直近の週**の値を表示しています（自社倉庫は現物確認の
たびに、TTAFは週次のPIVOT SOH TTAFのたびに更新）。合計在庫（週次グリッド本体）は、自社+TTAFの
実績が両方揃っている週があればその実測値で在庫計算をリセットし、それ以外の週は通常のロールフォワード
（入荷予定－使用量）で計算します。**将来週について自社/TTAFの内訳までは予測していません**（内訳は
実績が届いた週のみ分かるものです）。

## 2. Python不要・Excel(VBA)だけで完結する更新の仕組み

`macros/RefreshData.bas` に4つのマクロを用意しています。いずれも、対象ファイルを選ぶだけで
該当シートの値だけを更新し、**それ以外の入力済みデータ（T_Shipments・T_OpeningStock・
T_StockCount・安全在庫の設定値など）には一切触れません。**

| マクロ | 対象ファイル | 更新するシート |
|---|---|---|
| `RefreshWeeklyBatches` | Powder & Slurry & Pgm Plan（毎月） | PP_Grid（+ substrate分のM_BOM） |
| `RefreshBOM` | Usage from Production Engineering（改版時） | M_BOM |
| `RefreshSelfStock` | Raw materials daily check（自社在庫、現物確認のたび） | T_SelfStock |
| `RefreshTTAFStock` | CSA Report（TTAF在庫、週次） | T_TTAFStock |

`RefreshSelfStock`・`RefreshTTAFStock`は、選んだファイルの**ファイル名からDD.MM.YYYY形式の日付を
自動で読み取り**、その日付がどの週に該当するかをCal_Weeksと照合して記録します。ファイル名に日付が
含まれていない場合はエラーになりますので、ファイル名は変更せずそのまま使ってください。

**取り込みは必ず日付が新しい順に行ってください。** `Dashboard`の「直近実績」表示は「その原材料コードに
ついて表の中で一番後ろにある行」を最新として扱う仕組みのため、古いファイルを後から取り込むと表示が
古い値に戻ってしまう可能性があります。

**導入方法**
1. `SOH_Master.xlsx`を「名前を付けて保存」→ファイルの種類を「Excel マクロ有効ブック(*.xlsm)」にする
2. Alt+F11 でVBEを開く → 「ファイル」→「ファイルのインポート」→ `macros/RefreshData.bas`を選択
   （`.bas`ファイルを直接選ぶことで、モジュール名も含めて正しく読み込まれます）
   - コードをコピー＆貼り付けする場合は、**1行目の`Attribute VB_Name = "..."`を必ず削除してから**
     貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイルエラーになります
     （標準モジュールを挿入→中身を貼り付け、の手順を使う場合はこの点にご注意ください）
3. `macros/PO_Export.bas`（完全に任意、下記参照）も使う場合は同様にインポート
4. Alt+F8 → `RefreshWeeklyBatches`を選択して実行 → ファイル選択ダイアログで最新の
   「Powder & Slurry & Pgm Plan」ファイルを選ぶ
5. 同様に`RefreshBOM`も必要なタイミングで実行
6. お好みで、シート上に図形を配置し「マクロの登録」でボタン化すると次回以降クリックだけで済みます

**重要な注意点**: この環境ではVBAを実際にExcel上で実行して検証することができません（数式エンジンは
実データで動作確認済みですが、このマクロコードは未検証です）。まず貴社のExcelでテスト用のコピーに
対して動作確認をお願いします。エラーが出た場合は、エラーメッセージの内容を教えてください。

実行結果はメッセージボックスで「更新セル数」「新規追加した中間体/組み合わせ数」「未解決の材料名」
が表示されます。全く新しい中間体×原材料の組み合わせが増えた場合も、`M_BOM`・`PP_Grid`に行が
追加されるだけで`Grid_Requirement`はその場で自動的に反映されます（`M_BOM`・`PP_Grid`を直接
参照する設計のため、以前のような中間表の再生成は不要です）。

### なぜExcel外部参照やPower Queryではないのか
- **外部参照式**（別ブックのセルを直接参照する数式）は、この規模（原単位・バッチ数あわせて数千件）
  だと参照が多くなりすぎ、動作が重くなる懸念があったため採用しませんでした。
- **Power Query**は対象ファイルの構造が複雑（材料ごとに約40回繰り返すブロック構造）なため、
  Mコードが複雑になり、この環境で検証できないリスクが高いと判断しました。
- VBAは同じ理由で「未検証」ではありますが、通常のプログラムコードとして読みやすく、動作の予測が
  つきやすいため、この方式を採用しています。

`powerquery/Q_Shipments.pq`（CSA ReportのShipping Schedule取込み）だけは、シート構造がシンプルな
ため引き続きPower Query案として残しています（任意・未検証）。

**週の切り替えはマクロ不要です。** `Dashboard`のC1に週No(例:`W23`)を入力すると、`SUMPRODUCT`で
「現在年(`Cal_Weeks`のB1=AnchorYear)×入力した週No」に一致する週を`Cal_Weeks`から検索し、
`Status`列のすぐ右（I列）にその週の在庫を常時ピン留め表示します。列を右へスクロールして探す
必要はありません。該当する週の列（グリッド本体側）も黄色くハイライトされます。すべて通常の
Excel数式のみで実現しており、マクロは使っていません（`TODAY()`等の揮発性関数は使わず、
再計算負荷を抑えています）。

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
| T_StockCount | 入力 | 棚卸実測値（誤差リセット用、手動） |
| T_SelfStock | 入力（マクロ更新） | 自社倉庫の在庫実績。RefreshSelfStockで更新 |
| T_TTAFStock | 入力（マクロ更新） | TTAF倉庫の在庫実績。RefreshTTAFStockで更新 |
| M_RawMaterials | マスタ | 原材料マスタ・安全在庫設定 |
| M_BOM | マスタ（マクロ更新） | 原単位。RefreshBOMで更新 |
| PP_Grid | マスタ（マクロ更新） | 週次バッチ数。RefreshWeeklyBatchesで更新 |
| (非表示) Cal_Weeks / M_Intermediates / M_ProductMap / Grid_Requirement / Grid_Incoming | 内部計算 | 通常は開く必要なし |

## 4. 着荷予定(CSA Order)の入力方法

`T_Shipments`に1行追加してください。ETAを入力した週に、その数量が見込み在庫として自動反映
されます。着荷日が確定したら`Received_Date`を入力してください（未確定分はETAが使われます）。
`PO_No`・`Order_Date_発注日`もあわせて記録できます。

## 4.5 注文書(PO)の発行 — マクロ不要

`PO_Draft_Chemical`/`_Hazardous`/`_Substrate`は、常に最新の在庫予測にもとづいて自動計算される
「生きた」注文書です。TTAF R-Model等の既存の`Chemical Release`シートと同じ考え方で、**このシートを
そのまま印刷する、またはシートを右クリック→「移動またはコピー」→「コピーを作成する」にチェックを
入れて別ブックに複製する**だけで発注書として使えます。マクロは不要です。

表示範囲は`Dashboard`と異なり、**翌月分(Firm)＋翌々月・翌々々月分(Forecast)の13週のみ**です
（2年先までの計画はPOには不要という運用に合わせています）。各シート上部の入力セル
「基準週(WeekIndex)」に週番号を入れると、その週を起点に13週分の表示がスライドします
（既定値は翌月の第1週）。月-年／月曜日の日付／週番号の3段見出しと、Firm/Forecastの区分行は
すべて`Cal_Weeks`参照のExcel関数で自動計算されるため、年をまたいでも手直し不要です。

`macros/PO_Export.bas`は、この複製・PDF化・リビジョン番号の自動採番をボタン1つで行うための**任意の
補助マクロ**です。使わなくても運用に支障はありません。

## 5. 毎月の運用フロー

1. 「Powder & Slurry & Pgm Plan」の新しい月版を受け取ったら`RefreshWeeklyBatches`を実行
2. 「Usage from Production Engineering」が更新されていれば`RefreshBOM`を実行
3. 自社倉庫の現物確認（daily check）を実施したら`RefreshSelfStock`を実行
4. CSA Reportが週次で届いたら`RefreshTTAFStock`を実行し、あわせて`T_Shipments`もETA・着荷日・
   PO番号・発注日で更新（早着・遅着はここに反映）
5. 棚卸を実施した週は`T_StockCount`に追記
6. `Dashboard`で「要発注」を確認し、`PO_Draft_*`から注文書を発行
7. 月初は、前月最終週と当月頭の`Dashboard`を見比べて在庫差異を確認し、従来通り
   Plan Increase and Decrease / Inventory Releasesの報告フォーマットに転記

## 6. 自動反映の仕組み

- **生産計画が変わったとき**: `RefreshWeeklyBatches`実行後、`Grid_Requirement` → `Grid_Stock` →
  `Dashboard` → `PO_Draft_*` → `Material_Detail`まで自動的に再計算されます。
- **原単位が変わったとき**: `RefreshBOM`実行後、同様に自動反映されます。`M_BOM`・`PP_Grid`は
  `Grid_Requirement`から直接参照されているため、全く新しい中間体×原材料の組み合わせが増えた
  場合でも、ブックの再生成は不要です。
- **輸入品が早着・遅着したとき**: `T_Shipments`のETA・着荷日を書き換えるだけで、入荷が計上される
  週が自動的にシフトします。
- **棚卸で実測とズレがあったとき**: `T_StockCount`に1行追記すると、その週の在庫が上書きされ、
  以降はそこから積み上げ直されます。

## 7. 動作検証結果

LibreOffice（Excel互換の検証環境）で実際に数式を再計算させ、以下を確認済みです（VBAマクロと
Power Query部分を除く。両者はこの環境で実行できないため未検証です）。

- 全シートを通じて数式エラー（#REF!等）は0件
- `Grid_Requirement`の値が、実際の「Powder & Slurry & Pgm Plan」の元シートに記載された週次使用量
  と完全一致することを確認済み
- 生産計画変更・ETA変更（早着遅着）・安全在庫割れ検知の自動連鎖を実データで確認済み
- `M_BOM`・`PP_Grid`にVBAで行を追加しても、`Grid_Requirement`のSUMPRODUCT参照が正しい値を
  拾えるよう設計（数式ロジックはExcelの標準的な参照方式のため動作は確実ですが、VBA自体の実行は
  未検証です）

**重量パフォーマンスに関する経緯（重要）**: 初期設計では「BOM行×週」を1行ずつ展開した中間表
（`Calc_Demand`, 73,944行）を`Grid_Requirement`がSUMIFSで週次集計する方式でした。LibreOffice上の
計測では「開いて再計算するまで約25〜70秒」という結果だったため問題ないと判断していましたが、
実際のExcelでは非常に重くなり、開く・編集する・スクロールするたびにフリーズ・強制終了が
頻発する不具合が発生しました。LibreOfficeの計算エンジンはExcelと完全には同じではなく、
今回のように大規模な数式（10,504セル×74,000行のSUMIFS＝15億回超の比較）の実負荷を
正しく再現できていなかったことが原因です。**そのため、このシステムの重さに関する最終的な
判断は、LibreOfficeでの検証結果だけでなく実際のExcelでの動作確認を必ず優先してください。**
この教訓を踏まえ、`Calc_Demand`を廃止し`M_BOM`・`PP_Grid`から直接SUMPRODUCTで集計する方式に
再設計しました（計算量を約200分の1に削減）。実際のExcelでの改善効果については、貴社での
動作確認結果をお知らせください。

## 8. 要確認・要入力の項目

- **`M_RawMaterials`の`SafetyStock_Qty` / `LeadTime_Weeks`**: 仮値です。実際の安全在庫水準に
  置き換えてください。「要発注」判定の基準になります。
- **`M_RawMaterials`の`Category`**: 機械的に判定したものです。実際の危険物区分と一致しているか
  確認してください。
- **Substrates（基材）**: 「Powder & Slurry & Pgm Plan」内の"Japan GPF Substr"/"China Substr"/
  "Poland GPF Substr"（およびEster Film/PP Film等、1シート1品目のフィルム系シート）から
  週次使用量(=完成品Catコードの受注数量×1個あたり使用量)を取り込み、`M_RawMaterials`に
  Category="Substrate"として統合済みです。中間体を経由する化学原料と異なり、substrateは
  完成品コード(Cat)が直接「中間体」の役割を果たします（`M_Intermediates`にType="Cat"として
  登録）。`PO_Draft_Substrate`にも実データが反映されています。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫を入力してください
  （`T_SelfStock`・`T_TTAFStock`に実績があれば、その週以降は自動でリセットされます）。
- **`RefreshData.bas`・`Q_Shipments.pq`**: 未検証です。動作確認の結果を教えてください。
- **`T_SelfStock`/`T_TTAFStock`の取り込み順**: 必ず日付が新しい順に`RefreshSelfStock`/
  `RefreshTTAFStock`を実行してください（詳細は2章参照）。

## 9. 今後の拡張候補

- **Min/Max（週数ベースの安全在庫）モデル**: 現状は単一しきい値のみですが、「N週分の使用量」を
  基準にしたMin/Max運用に拡張することも可能です。
