# SOH（在庫管理）シート システムガイド

## 1. これは何か（位置づけ）

`Dashboard`が最終的にすべての原材料の在庫を確認するメイン画面、`Material_Detail`がその内訳
（どの材料が何に使われているか）を見る画面、`PO_Draft_*`が発注書の発行画面です。それ以外の
シート（`M_RawMaterials`、`M_BOM`、`PP_Grid`、`Grid_Stock`等）は計算のための土台で、通常は開く
必要はありません（タブは非表示にしてあります）。

**`Dashboard`は、原材料×週の在庫を2年分・横軸で見渡せる設計にしています。** 生産計画から
「Week◯の在庫がいくつ、翌週はいくつ…」と2年先まで追えます。週の見出しは3段（1段目=月-年、
2段目=その週月曜の日付、3段目=週No）で、すべてExcelの数式（`Cal_Weeks`シート参照）で計算
されています。材料ごとに「理論在庫」（実績を一切考慮しない純粋なロールフォワード予測）と
「実在庫」（手動棚卸・実績を反映した値）の2行を並べて表示し、両者の乖離も確認できます
（詳細は5.6章）。`Dashboard`のC1セルに`W23`のように入力すると（現在年の週Noとして検索）、該当する週の列が
太枠でハイライトされます。VBAの`JumpToSelectedWeek`（要手動設定・詳細は5.6章）を導入して
いれば、複製データではない本物の該当週列が、固定ペインの直後に来る
ようウィンドウが自動的に横スクロールされます。

**このブックが担う範囲は「在庫管理」に絞っています。** 生産計画や原単位そのものの再計算はせず、
既存のファイルの値を月次でマクロ経由で取り込みます。

- **顧客オーダー・生産計画の策定**: このブックの範囲外
- **原単位（1バッチあたり使用量）**: 「Raw Material - Look Up」からマクロで取り込み
- **週次バッチ数**: 「Powder & Slurry & Pgm Plan」（毎月改版）からマクロで取り込み
- **自社倉庫の在庫実績**: 「Raw materials daily check」（現物確認シート）からマクロで取り込み
- **TTAF倉庫の在庫実績**: 「CSA Report」の`Stock invoiced to CSA`シート（手入力の生データ）からマクロで取り込み
- **在庫のロールフォワード、着荷予定との突合、発注アラート、PO発行**: このブックの役割

「Plan Increase and Decrease」「Inventory June Releases」はこのブックの計算から切り離しています。
月初の在庫差異報告（前月最終週⇔当月頭）には、`Dashboard`の週次実績を元データとしてご利用ください。

**自社在庫・TTAF在庫の内訳**: `Dashboard`に「自社在庫(実績)」「TTAF在庫(実績)」「実績週」列があります。
これは各原材料について、実績データが届いている**直近の週**の値を表示しています（自社倉庫は現物確認の
たびに、TTAFは週次のCSA Report取込みのたびに更新）。合計在庫（週次グリッド本体）は、自社+TTAFの
実績が両方揃っている週があればその実測値で在庫計算をリセットし、それ以外の週は通常のロールフォワード
（入荷予定－使用量）で計算します。**将来週について自社/TTAFの内訳までは予測していません**（内訳は
実績が届いた週のみ分かるものです）。

## 2. Python不要・Excel(VBA)だけで完結する更新の仕組み

`macros/`フォルダに、機能ごとに分割した8つのVBAモジュール(`RefreshData_*.bas`)を
用意しています。以前は1つの巨大な標準モジュールでしたが、メンテナンス性のため
`RefreshData_Utilities`（共通ヘルパー）・`RefreshData_ProductionPlan`・`RefreshData_BOM`・
`RefreshData_StockActuals`・`RefreshData_Shipments`・`RefreshData_Display`・
`RefreshData_MaterialMgmt`・`RefreshData_PODraft`の8モジュールに分割しています
（マクロ名・動作は分割前と同じ）。
更新マクロはいずれも、対象ファイルを選ぶだけで該当シートの値だけを更新し、
**それ以外の入力済みデータ（T_Shipments・T_OpeningStock・T_StockCount・基準在庫の設定値など）
には一切触れません。**

| マクロ | 対象ファイル | 更新するシート |
|---|---|---|
| `RefreshWeeklyBatches` | Powder & Slurry & Pgm Plan（毎月） | PP_Grid |
| `RefreshBOM` | Raw Material - Look Up（改版時） | M_BOM |
| `RefreshSelfStock` | Raw materials daily check（自社在庫、毎週月曜の朝） | T_SelfStock_Log（非表示。目に見えるT_SelfStockは数式で自動反映） |
| `RefreshTTAFStock` | CSA Report（TTAF在庫、毎週月曜） | T_TTAFStock_Log（非表示。目に見えるT_TTAFStockは数式で自動反映） |
| `RefreshShipments` | CSA Report（Shipping Schedule、毎週月曜） | T_Shipments（Shipping Scheduleの全件を材料＋PO番号＋コンテナ＋Original ETDの複合キーで一括反映、分割出荷も欠落なく反映）。あわせてMaterial_DetailのOrder/PO_No行もCSA ReportのStatusに合わせて自動更新（詳細は5.10.1章） |
| `SetupOrderManagementMigration` | （ファイル選択なし・実行するだけ） | Material_DetailへのPO_No行追加と、Grid_Incomingの数式書き換えをまとめて行う、一度だけ実行する移行用マクロ（詳細は5.10.1章） |
| `AddShipmentSplitColumns` | （ファイル選択なし・実行するだけ） | T_ShipmentsにVessel/Container/Original_ETD列を追加し、あわせて全行のEffective_Week数式を復元する移行用マクロ（分割出荷の欠落防止・数式破壊バグの修復、詳細は5.10.2〜5.10.3章。何度実行しても安全） |
| `CleanupOrphanedPreSplitShipmentRows` | （ファイル選択なし・実行するだけ） | `AddShipmentSplitColumns`実行後、最初の`RefreshShipments`で二重計上の原因になり得る移行前の旧形式行を削除する移行用マクロ（詳細は5.10.5章。何度実行しても安全） |
| `HideInactiveIntermediates` | （ファイル選択なし・実行するだけ） | Material_Detailの行の表示/非表示のみ（数値は変更しない） |
| `ShowAllIntermediates` | （ファイル選択なし・実行するだけ） | Material_Detailの非表示行をすべて再表示 |
| `JumpToSelectedWeek` | （ファイル選択なし・C1変更時に自動呼び出し） | Dashboard/Material_Detail/T_SelfStock/T_TTAFStockのウィンドウ表示位置のみ（数値は変更しない） |
| `AddMaterial` | （ファイル選択なし・InputBoxで入力） | 新しい材料を全関連シートの一番下に追加（詳細は5.7章） |
| `FixOpeningStockColumnReference` | （ファイル選択なし・実行するだけ） | Grid_Stock・Grid_TheoreticalStockの週1列の数式にあったT_OpeningStock列名参照の誤りを修正する移行用マクロ（詳細は5.7.1章。何度実行しても安全） |
| `SyncPODraftCategories` | （ファイル選択なし・実行するだけ） | M_RawMaterialsのCategory・Origin_Country列を後から書き換えた際、PO_Draft_*シート側の振り分けを実際の値に合わせて同期し直す（詳細は5.7.2〜5.7.3章。いつでも安全に実行可） |
| `SetupSubstratePODraftByCountry` | （ファイル選択なし・実行するだけ） | M_RawMaterialsへのOrigin_Country列追加、PO_Draft_SubstrateのPO_Draft_Substrate_JPN_CHNへの改名、PO_Draft_Substrate_Polandシートの新規作成をまとめて行う移行用マクロ（Substrateの原産国別PO_Draft分離のため。詳細は5.7.3章。何度実行しても安全） |
| `RemoveMaterial` | （ファイル選択なし・InputBoxで入力） | 指定した材料を全関連シートから削除（詳細は5.7章） |
| `RemoveIntermediate` | （ファイル選択なし・InputBoxで入力） | 生産中止になった中間体をPP_Grid・M_BOM・Material_Detailから削除（詳細は5.7章） |
| `SetupPODraftLetterheadLayout` | （ファイル選択なし・実行するだけ） | PO_Draft_Hazardousに手動で作り込んだレターヘッド形式レイアウト（月/週見出しの不具合修正・基準週参照の名前付き範囲化・Firm/Forecast色分け・SafetyStock/CurrentStockの印刷範囲除外）を修正した上で、PO_Draft_Chemical・PO_Draft_Substrate_JPN_CHN・PO_Draft_Substrate_Polandの3シートにも複製する、一度だけ実行する移行用マクロ（詳細は5.7.4章。何度実行しても安全） |

**TTAF供給材料の在庫予測の考え方（重要）**: TTAFは仕入先であると同時に、原材料を預けている
倉庫でもあります。`T_Shipments`（Status=TTAF Stock）は「TTAFが外部の仕入先から新しく仕入れて
TTAF倉庫に到着する」実績・予定を表します。これはTTAF倉庫内で場所が移っただけの動きではなく、
純粋に合計在庫へ新規に入ってくる量なので、`Grid_Stock`はTTAF供給材料も含めて全材料共通で
「前週＋入庫－消費」のロールフォワード式です（TTAF供給材料だけの特別扱いはありません）。
実績（CSA Reportが毎週月曜に届くたびの`T_TTAFStock`）がある週は、そちらが優先して使われます。

`HideInactiveIntermediates`/`ShowAllIntermediates`/`JumpToSelectedWeek`はいずれもデータを
更新するマクロではなく、見た目（行の表示/非表示、ウィンドウのスクロール位置）だけを
操作するものです。詳細は5.6章を参照してください。

`RefreshTTAFStock`は、CSA Report内の`Stock invoiced to CSA`シート（TTAF側の**手入力の生データ**）
を直接読み込みます。以前使っていた` COUNT SHEET SOH`・`PIVOT SOH TTAF`はいずれもピボット
テーブルで、TTAF側が更新（Refresh）を忘れたまま送ってくると古いキャッシュ値のまま取り込んで
しまう問題がありましたが、手入力データを直接参照する現在の方式ではその心配がありません。
`RefreshShipments`は、CSA Reportファイルを開いた直後に**ブック内の全ピボットテーブルを自動で
更新（RefreshAll）**してからデータを読み込みます（`Shipping Schedule`シート自体がピボットかどうか
未確認のため、念のための対策です）。

`RefreshSelfStock`・`RefreshTTAFStock`は、選んだファイルの**ファイル名からDD.MM.YYYY形式の日付を
自動で読み取り**、その日付がどの週に該当するかをCal_Weeksと照合して記録します。ファイル名に日付が
含まれていない場合はエラーになりますので、ファイル名は変更せずそのまま使ってください。

**`RefreshSelfStock`は、ファイル名の日付から7日引いてから対象週を判定します**（毎週月曜の朝に
自社在庫を確認する運用のため。月曜朝の在庫は前週末時点の状態を表すので、前週の実績として記録
します）。これは`RefreshTTAFStock`（CSA Reportの対象日から7日引く）と同じ考え方です。月曜が祝日で
別の曜日に確認した場合も、7日引くだけなので正しく前週の範囲内に収まります。数週間まとめて休業と
なり、その間まったく確認できなかった場合も特別な対応は不要です。確認できなかった週は空欄
（通常のロールフォワード計算）のままになり、休業明けに次の確認を行えば、その週から7日前＝
休業直前の週の実績として、通常どおり記録されます。

**同じ週内に複数回取り込んでも、自動的に1件にまとめられます。** 取り込んだ実施日から「その週の
月曜日」を計算し、それをキーに記録するため、同じ週の中で日を分けて何度実行しても行が増え続ける
ことはありません（後から実行した方の値で上書きされます）。`Dashboard`の「直近実績」表示は、
週の列を右から見て最後に値がある週を最新として扱う仕組みのため、取り込みの順番が前後しても
（週をまたいで古いファイルを後から取り込んだ場合を除き）表示が古い値に戻ることはありません。

**導入方法**
1. `SOH_Master.xlsx`を「名前を付けて保存」→ファイルの種類を「Excel マクロ有効ブック(*.xlsm)」にする
2. Alt+F11 でVBEを開く → 「ファイル」→「ファイルのインポート」→ `macros/`フォルダの
   `RefreshData_*.bas`8ファイルをすべて選択（複数選択して一括インポート可。1ファイルずつでも
   構いません。インポートする順序は結果に影響しません）
   （`.bas`ファイルを直接選ぶことで、モジュール名も含めて正しく読み込まれます）
   - コードをコピー＆貼り付けする場合は、**各ファイル1行目の`Attribute VB_Name = "..."`を
     必ず削除してから**貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイル
     エラーになります（標準モジュールを挿入→中身を貼り付け、の手順を使う場合はこの点にご注意
     ください。また8つのモジュールはそれぞれ別の標準モジュールとして挿入してください）
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

**週の切り替え自体（該当週の検索・列ハイライト）はマクロ不要です。** `Dashboard`のC1に
週No(例:`W23`)を入力すると、`SUMPRODUCT`で「現在年(`Cal_Weeks`のB1=AnchorYear)×入力した
週No」に一致する週を`Cal_Weeks`から検索し、該当する週の列（グリッド本体側）を太枠で
ハイライトします。すべて通常のExcel数式のみで実現しており、マクロは使っていません
（`TODAY()`等の揮発性関数は使わず、再計算負荷を抑えています）。

**選択週の列を固定ペインのすぐ右に自動で持ってくる部分だけは、VBA(`JumpToSelectedWeek`)を
使います（任意）。** 以前のバージョンでは、選択週の値を別列（ピン留め列）に数式で複製して
常時表示していましたが、複製である以上「本体の週列と数字が食い違って見える」リスクや
複製専用の列がレイアウトを圧迫する欠点がありました。現在は複製列を廃止し、代わりに
C1変更時にウィンドウを横スクロールして、複製ではない本物の週データ列そのものを
ラベル列のすぐ右に表示する方式にしています。`Dashboard`は`乖離(kg)`列の右（K列〜）、
`Material_Detail`は`項目`列の右（D列〜）が、そのまま週1・週2…の実データ列です。
VBA未導入の場合はスクロールが自動では起きませんが、太枠ハイライトを目印に手動で
スクロールしても支障なく使えます（詳細・導入手順は5.6章）。

`Material_Detail`は約1,660行あるため、`Dashboard`のような「該当する週の列を条件付き書式で
ハイライトする」機能はあえて追加していません（週×全行に条件付き書式をかけると対象セル数が
`Dashboard`の約16倍になり、パフォーマンス面で新たなリスクになるため）。

### Pythonスクリプトについて（参考・任意）
`scripts/`にPython版の抽出・再生成スクリプトも残しています。Python環境がある場合はこちらでも
更新できますが、**ブックをまるごと再生成する**ため、通常の月次運用にはVBAマクロ（上記）を
お使いください。

## 3. シート構成と役割

| シート | 種別 | 役割 |
|---|---|---|
| README | - | 使い方サマリ・各シートへのジャンプリンク |
| 操作パネル | - | マクロボタンの設置場所（月次運用手順の順に並べる。詳細は5.8章） |
| Dashboard | 出力（最終確認画面） | 原材料×週の在庫を2年分・横軸で表示。材料ごとに「理論在庫」「実在庫」の2行＋乖離(kg)列。基準在庫の下限/上限を入力すると実在庫行の各週が赤(下限未満)/緑(範囲内)/青(上限超)に自動色分け、C1入力で該当週を太枠ハイライト |
| Material_Detail | 出力（トレーサビリティ）＋発注数量の入力欄 | 材料ごとに「使用中間体・バッチ数・使用量・週次合計・TTAF/自社在庫実績・週末時点の合計在庫」をブロック表示。材料名の右にMOQ手入力欄あり。「Order(発注予定,kg)」行は週ごとの発注数量を直接手入力する欄（5.9章参照） |
| PO_Draft_Chemical / _Hazardous / _Substrate | 出力（発注書） | Material_Detailで手入力したOrderの値を、材料コード・週で突き合わせてそのまま転記（自動計算はしない） |
| T_Shipments | 入力（RefreshShipmentsでも更新可） | 発注〜輸送〜着荷（PO番号・発注日・ETA・着荷日・発注月）。TTAF供給材料についてはTTAF倉庫への到着実績を表す |
| T_OpeningStock | 入力 | 起点となる期首在庫 |
| T_StockCount | 入力 | 棚卸実測値（誤差リセット用、手動） |
| T_SelfStock | 出力（閲覧用、数式のみ） | 自社倉庫の在庫実績を材料×週のグリッドで表示。RefreshSelfStockで裏の`T_SelfStock_Log`が更新されると自動反映。手入力不可 |
| T_TTAFStock | 出力（閲覧用、数式のみ） | TTAF倉庫の在庫実績を材料×週のグリッドで表示。RefreshTTAFStockで裏の`T_TTAFStock_Log`が更新されると自動反映。手入力不可 |
| M_RawMaterials | マスタ | 原材料マスタ・安全在庫設定・TTAF_Code・固定週次消費量_要入力(5.9章参照) |
| M_BOM | マスタ（マクロ更新） | 原単位。RefreshBOMで更新 |
| PP_Grid | マスタ（マクロ更新） | 週次バッチ数。RefreshWeeklyBatchesで更新 |
| (非表示) T_SelfStock_Log / T_TTAFStock_Log | 入力（マクロ更新） | 自社/TTAF在庫実績の生ログ（実施日ベース）。RefreshSelfStock/RefreshTTAFStockが書き込む実体。AnchorYearを何度進めても壊れない安全な保存先 |
| (非表示) Cal_Weeks / M_Intermediates / M_ProductMap / Grid_Requirement / Grid_Incoming / Grid_Stock / Grid_TheoreticalStock | 内部計算 | 通常は開く必要なし。Grid_TheoreticalStockはDashboardの「理論在庫」行の元になる、実績(T_StockCount・自社/TTAF実績)を考慮しないロールフォワード専用シート。ただし月が変わる最初の週だけは前週の実在庫(Grid_Stock)を起点に再同期するため、誤差は最大でも1か月分しか蓄積しない |

## 4. 着荷予定(CSA Order)の入力方法

`T_Shipments`に1行追加してください。ETAを入力した週に、その数量が見込み在庫として自動反映
されます。着荷日が確定したら`Received_Date`を入力してください（未確定分はETAが使われます）。
`PO_No`・`Order_Date_発注日`もあわせて記録できます。

## 4.5 注文書(PO)の発行 — マクロ不要

`PO_Draft_Chemical`/`_Hazardous`/`_Substrate`は、発注する数量を自動計算するシートでは
**ありません**（以前は基準在庫下限とGrid_Stockの在庫予測から自動計算していましたが、その方式は
廃止しました）。発注数量は`Material_Detail`の各材料ブロックにある**「Order(発注予定,kg)」行に
週ごとに直接手入力**してください。`PO_Draft_*`はその手入力値を、材料コード・週を突き合わせて
そのまま転記するだけの「発注書レイアウト」です。TTAF R-Model等の既存の`Chemical Release`シートと
同じ考え方で、**このシートをそのまま印刷する、またはシートを右クリック→「移動またはコピー」→
「コピーを作成する」にチェックを入れて別ブックに複製する**だけで発注書として使えます。
マクロは不要です。

表示範囲は`Dashboard`と異なり、**翌月分(Firm)＋翌々月・翌々々月分(Forecast)の13週のみ**です
（2年先までの計画はPOには不要という運用に合わせています）。各シート上部の入力セル
「基準週(WeekIndex)」に週番号を入れると、その週を起点に13週分の表示がスライドします
（既定値は翌月の第1週）。月-年／月曜日の日付／週番号の3段見出しと、Firm/Forecastの区分行は
すべて`Cal_Weeks`参照のExcel関数で自動計算されるため、年をまたいでも手直し不要です。

`macros/PO_Export.bas`は、この複製・PDF化・リビジョン番号の自動採番をボタン1つで行うための**任意の
補助マクロ**です。使わなくても運用に支障はありません。

## 5. 毎月の運用フロー

1. 「Powder & Slurry & Pgm Plan」の新しい月版を受け取ったら`RefreshWeeklyBatches`を実行
2. 「Raw Material - Look Up」が更新されていれば`RefreshBOM`を実行
3. 自社倉庫の現物確認（daily check）を毎週月曜の朝に実施したら`RefreshSelfStock`を実行
4. CSA Reportが毎週月曜に届いたら`RefreshTTAFStock`と`RefreshShipments`を実行
   （`RefreshShipments`はShipping Scheduleの全PO・全材料を一括で`T_Shipments`に反映）
5. 発注する数量は`Material_Detail`の該当材料ブロックの「Order(発注予定,kg)」行に、発注したい
   週の列へ直接手入力する
6. 棚卸を実施した週は`T_StockCount`に追記
7. `Dashboard`で赤色（基準在庫の下限未満）の週を確認し、`PO_Draft_*`（手入力したOrderの値が
   自動転記されます）から注文書を発行
8. 月初は、前月最終週と当月頭の`Dashboard`を見比べて在庫差異を確認し、従来通り
   Plan Increase and Decrease / Inventory Releasesの報告フォーマットに転記

## 5.5 表示ウィンドウを先の年へ進める（年次作業・Dec-27以降の追加方法）

このブックは「常に2年分(104週)」を表示するローリングウィンドウ方式です。現状は`Cal_Weeks`の
B1セル(AnchorYear)=2026を起点に、2026年〜2027年末あたりまでを表示しています。それより先
(例: 2028年分)を見えるようにしたい場合は、以下の手順で対応してください。**Pythonの再生成も
ブックの作り直しも不要で、同じ1つのファイルを継続して使えます。**

1. `Cal_Weeks`シートのB1セル(AnchorYear)を、次に基準としたい年（例: `2027`）に書き換える
2. これだけで、`Cal_Weeks`のWeekStart/Year/WeekOfYear/Label等が全てExcelの数式で再計算され、
   `Dashboard`・`Material_Detail`・`PO_Draft_*`・`Grid_Requirement`等、日付を扱う全シートの
   週の並びが新しい2年分（2027年〜2028年末あたり）にスライドします
3. `PP_Grid`（生産計画バッチ数）は毎月`RefreshWeeklyBatches`で最新のPlanファイルの内容に
   上書きされるため、AnchorYear切り替え後に最初の`RefreshWeeklyBatches`を実行すれば、新しい
   週番号の並びに合わせて自然に上書きされます（切り替え直後は空欄/0が見えることがありますが、
   次のRefresh実行で解消します）

**過去の実績データの生ログ(`T_StockCount`/非表示の`T_SelfStock_Log`/`T_TTAFStock_Log`)は安全です**:
これらのシートの`WeekIndex`列は、記録した`Date`(実際の暦日)から**毎回ライブ計算する数式**に
なっており、`AnchorYear`が変わっても記録済みのデータが「別の週のデータ」として誤表示される
ことはありません。**AnchorYearを変更する前に事前バックアップを取る必要はなく**、必要なタイミング
でいつでも、どんな頻度でも切り替えて構いません。生ログのセル自体は削除されずに残っているため、
後から見返したい場合はシートの表示を解除して直接開けば確認できます。

一方、**目に見える方の`T_SelfStock`/`T_TTAFStock`（材料×週のグリッド表示）は、`Dashboard`・
`Grid_Stock`と同じく「現在の2年間ウィンドウ」に紐づいています**。これは生ログから毎回数式で
計算し直す「見るための表」なので、AnchorYearを進めるとウィンドウの外に出た週の列がグリッド上
からは見えなくなります（生ログ自体は上記の通り安全に残っているので、データが失われるわけでは
ありません。あくまで「一覧表示の対象期間」が今の2年間に絞られる、というイメージです）。

（本項目は、`T_Shipments`のEffective_Week計算式および`T_StockCount`/`T_SelfStock_Log`/`T_TTAFStock_Log`の
WeekIndex計算式が、いずれもビルド時点のAnchorYearを固定値として埋め込んでいた不具合を修正した
上で有効です。）

## 5.6 基準在庫の色分け／Material_Detailの在庫行・MOQ・中間体の折りたたみ

**Dashboardの基準在庫（下限/上限）**: 従来の単一の`SafetyStock_Qty`（Status="要発注"の判定用）を
廃止し、`M_RawMaterials`に`基準在庫下限_要入力`/`基準在庫上限_要入力`の2列を新設しました。
Dashboardには`Status`列の代わりにこの2つの基準値を入力する列（在庫グリッドのすぐ左）を設け、
各週の在庫セルを次のルールで自動的に色分けします。

- 赤（下限未満）: その週の在庫 < 下限
- 緑（範囲内）: 下限 ≤ その週の在庫 ≤ 上限
- 青（上限超）: その週の在庫 > 上限

C1に選択週（`W23`形式）を入力したときの列ハイライトは、この赤/緑/青の塗りつぶしと視覚的に
競合しないよう、塗りつぶしではなく**太い罫線（左右のボーダー）**に変更しています。

**理論在庫／実在庫の2段表示**: `Dashboard`は材料ごとに2行（上段=グレー文字の「理論在庫」、
下段=通常の黒文字の「実在庫」）を並べています。

- **理論在庫**: `T_StockCount`（手動棚卸）や自社/TTAF在庫実績を一切考慮しない、
  「前週在庫＋入庫－消費」だけで計算するロールフォワード値（非表示の
  `Grid_TheoreticalStock`シートを参照）。ただし月が変わる最初の週だけは、前週の実在庫
  (`Grid_Stock`)を起点に再同期します（毎月リセットされるため、誤差が無期限に積み上がる
  ことはなく、月内の乖離だけを見られます）。
- **実在庫**: 従来どおりの`Grid_Stock`（手動棚卸 > 自社+TTAF実績の合計 > 通常のロールフォワード、
  の優先順位）を参照した値。実際の発注判断にはこちらを使います。
- **乖離(kg)列**: 「実績週」列(H列)と同じ基準週（自社在庫実績が入っている直近の週）時点での
  「実在庫－理論在庫」。理論在庫が月初にリセットされる仕様のため、この値は「今月に入って
  から積み上がった、原単位・生産計画等のシステム上の計算と実際の現場とのズレ」を表します。
  乖離が大きいほど、今月の計画と実績のズレが大きいことを意味します。

基準在庫の赤/緑/青の色分けは**実在庫の行にのみ**適用されます（発注判断に使うのは実際の在庫の
ため）。理論在庫の行は色分けの対象外で、参考表示としてグレー文字にしています。

**Material_Detailの週次在庫行**: 各材料ブロックの「合計使用量(kg)/週」行の下に、以下の3行を
追加しました。

- `TTAF在庫(実績,kg)`: `T_TTAFStock`（材料×週のグリッド）の該当セルをそのまま参照
- `自社在庫(実績,kg)`: `T_SelfStock`（同上）の該当セルをそのまま参照
- `合計在庫(週末時点,kg)`: `Grid_Stock`をそのまま参照（手動棚卸`T_StockCount` > 自社+TTAF実績の
  合計（両方揃っている週のみ） > 通常のロールフォワード、という優先順位で計算される値。
  Dashboardの在庫グリッドと同じ値のため、両シートの数字は必ず一致します）

**MOQ（最小発注量）**: 各材料ブロックのヘッダー行、材料名（B列）の右のC列に、手入力用の
空セル（黄色ではなく水色の入力用セル）を用意しました。数式化はしておらず、手書きの数値を
そのまま入力する運用です。セルにマウスオーバーすると入力を促すコメントが表示されます。

**生産予定の無い中間体の折りたたみ**: Material_Detailは材料数×中間体数で行数が多くなるため、
しばらく（半年・1年など）生産予定の無い中間体の行をボタン一つで折りたためるようにしました。

- `HideInactiveIntermediates`マクロを実行すると、月数を尋ねるダイアログが出ます（例: `6`と
  入力すると、今週から6ヶ月間ずっと「No. of batches」が全週0の中間体の行（No. of batches行＋
  使用量(kg)行の2行）を非表示にします）
- 材料名の行、合計使用量／TTAF在庫実績／自社在庫実績／合計在庫(週末時点)の行は、非表示の対象に
  なりません。**在庫の数量そのものは折りたたみの影響を受けず常に見える**ようになっています
- `ShowAllIntermediates`マクロを実行すると、非表示にした行をすべて再表示します
- 判定は行のラベル文字列（C列="No. of batches"）を基準にしているため、材料や中間体の数が
  増減してもそのまま動作します
- 期間を変えて`HideInactiveIntermediates`を何度実行しても、実行のたびに一度全行を再表示してから
  判定し直すため、常に指定した条件どおりの結果になります

**ボタンの割り当て（一度だけの手動設定）**: openpyxl（このブックの生成に使っているPython
ライブラリ）はクリック可能なボタン図形やそこへのマクロ割り当てを自動生成できないため、以下は
貴社のExcelで一度だけ手動で行ってください。

1. `Material_Detail`シートを開く
2. 「挿入」タブ→「図形」等で、ボタンにしたい図形を1〜2個描く（例:「中間体を折りたたむ」
   「全部表示」）。シート上部の空いている場所（A1付近など）に配置すると便利です
3. 図形を右クリック→「マクロの登録」→片方に`HideInactiveIntermediates`、もう片方に
   `ShowAllIntermediates`を割り当てる

**選択週へのウィンドウ自動スクロール（`JumpToSelectedWeek`、任意・要手動設定）**: 以前の
「選択週の値を別列に複製して常時表示する」方式（ピン留め列）は廃止し、`Dashboard`の`実績週`列
・`Material_Detail`の`項目`列・`T_SelfStock`/`T_TTAFStock`の`Part Name`列のすぐ右に、複製では
ない本物の週データ列（週1・週2…）をそのまま並べる構成にしています。C1に週No(`W23`等)を入力
したときにその該当週列が固定ペインの直後にくるよう自動でウィンドウを横スクロールしたい場合は、
以下を一度だけ設定してください（設定しない場合も、該当週の列は太枠でハイライトされるため、
手動でスクロールして探すことは可能です）。

1. Alt+F11でVBEを開く
2. プロジェクトエクスプローラーで「Dashboard」シートをダブルクリックし、そのシート専用の
   コードモジュールに以下を貼り付ける（**標準モジュールに貼り付けても発火しません**）
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
       Call JumpToSelectedWeek(Me, "F1", 11)   ' 11 = K列(週データ開始列)
   End Sub
   ```
3. 同様に「Material_Detail」シートのコードモジュールにも以下を貼り付ける
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
       Call JumpToSelectedWeek(Me, "F1", 4)   ' 4 = D列(週データ開始列)
   End Sub
   ```
4. 「T_SelfStock」「T_TTAFStock」シートにも、それぞれのコードモジュールに以下を貼り付ける
   （2つとも同じ内容です）
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
       Call JumpToSelectedWeek(Me, "F1", 2)   ' 2 = B列(週データ開始列)
   End Sub
   ```

呼び出し先の`JumpToSelectedWeek`本体は`macros/RefreshData_Display.bas`（標準モジュール）側に
実装済みのため、通常どおり`RefreshData_*.bas`をインポートしていれば追加の作業は不要です。
上記4つの`Worksheet_Change`だけを、対応するシート自身のコードモジュールに貼り付けてください。

## 5.7 材料・中間体の追加・削除（`AddMaterial` / `RemoveMaterial` / `RemoveIntermediate`）— Python不要

駐在員の帰国後もローカル社員だけで運用を続けられるよう、対象材料（TTAF供給品）の追加・削除を
Excel(VBA)だけで完結できるようにしています。ブックの再生成（Pythonスクリプト実行）は不要です。

**`AddMaterial`マクロ**: Alt+F8から実行すると、InputBoxが順番に表示されます。

1. Part Name（RM_Code。既存と重複していないかを自動チェック）
2. Description（品名）
3. Supplier（未入力なら既定で`TTAF`）
4. Category（`Chemical` / `Hazardous Chemical` / `Substrate`のいずれか。それ以外を入力すると
   やり直しを求められます）
5. TTAF_Code（無ければ空欄でOK）

最後に入力内容の確認ダイアログが出るので、内容を確認して「はい」を選ぶと、以下のシートの
一番下に必要な行がまとめて追加されます。

`M_RawMaterials` → `Grid_Requirement` → `Grid_Incoming` → `Grid_Stock` → `Grid_TheoreticalStock` →
`T_OpeningStock` → `T_SelfStock` → `T_TTAFStock` → `Dashboard`（理論在庫・実在庫の2行） →
`Material_Detail` → 該当カテゴリの`PO_Draft_*`

追加直後はこの材料をまだどの中間体も使っていない（`M_BOM`に実績が無い）ため、
`Material_Detail`のブロックは中間体の内訳行が無い簡易版（合計使用量・TTAF在庫・自社在庫・
Order・PO_No・合計在庫の6行のみ）になります。その後**通常どおり`RefreshBOM`を1回実行するだけで**、
実際にこの材料を使う中間体の使用実績が見つかり次第、`Material_Detail`の該当ブロックに
「No. of batches」「使用量(kg)」の内訳行が自動的に追加され、他の材料と同じ見た目になります
（`AddMaterial`を2回実行する必要はありません）。既存の材料が新しい中間体で使われ始めた場合も
同様に、`RefreshBOM`を実行するだけで内訳行が自動的に追加されます。

**`RemoveMaterial`マクロ**: Part Name（RM_Code）を入力すると、`AddMaterial`が追加する全シートから
該当行を削除します。**`T_Shipments`・`T_StockCount`・`T_SelfStock_Log`/
`T_TTAFStock_Log`・`M_BOM`のデータは削除しません**（履歴として残しておき、万一同じPart Nameを
`AddMaterial`で再登録した場合は自動的に再びつながる設計です）。

### 5.7.1 T_OpeningStock列名参照の誤りの修正

`AddMaterial`実行時に「実行時エラー(1004) アプリケーション定義またはオブジェクト定義のエラーです」
というエラーで途中停止することがある不具合が見つかりました。原因は、Grid_Stock・
Grid_TheoreticalStockの週1列（ブック内で一番古い週）の数式が、`T_OpeningStock`テーブルの
列名を`Opening_Qty`と誤って参照していたこと（正しい列名は`Opening_Qty_要入力`）です。
この誤りは`build_soh.py`側にも同じ形であったため、既存の全材料の週1列の数式にも
同じ誤りが書き込まれていました。

実際には、既存材料は週1時点で既に自社在庫・TTAF在庫の実績データが揃っているため、この
壊れた参照を含む分岐は普段は実行されず(数式内のIFの他の分岐が優先される)、これまで
気づかれていませんでした。しかし`AddMaterial`で追加したばかりの材料には週1時点の
実績データがまだ無いため、この壊れた分岐に実際に到達し、エラーになっていました。

`AddMaterial`自体は既に修正済みです。ただし、修正前の状態で既に書き込まれてしまっている
**既存材料**の週1列の数式は、コードを直しただけでは直りません。`FixOpeningStockColumnReference`
マクロを一度実行すると、Grid_Stock・Grid_TheoreticalStockの全材料の週1列の数式が一括で
正しい状態に修正されます（他の週の列には影響しません。誤って複数回実行しても安全です）。

### 5.7.2 M_RawMaterialsのCategory変更をPO_Draft_*に反映する（`SyncPODraftCategories`）

`PO_Draft_Chemical`・`PO_Draft_Hazardous`・`PO_Draft_Substrate_*`の各行は、`AddMaterial`
実行時点（または`build_soh.py`実行時点）の`M_RawMaterials`の`Category`（Substrateの
場合はさらに`Origin_Country`）の値に基づいて、その時1回だけ該当シートに追加されます。
**数式でリアルタイムに参照して自動的に振り分け直す仕組みではありません。**

そのため、材料を登録した後に`M_RawMaterials`側で`Category`や`Origin_Country`を修正しても
（例:「Hazardous Chemical」→「Chemical」に訂正、原産国の追記・訂正）、`PO_Draft_*`シート
側の行は元のシートに残ったままになり、自動的には移動しません（実際にこの不具合が報告され、
発見されました）。

`SyncPODraftCategories`マクロを実行すると、`M_RawMaterials`の現在の`Category`・
`Origin_Country`を正として`PO_Draft_*`の全シートを全材料分スキャンし直し、矛盾している行
（間違ったシートに残っている行・まだどの`PO_Draft_*`にも無い行）を、正しいシートへ
削除→再作成する形で移動させます。発注数量自体はMaterial_Detailへの参照で持っているため、
行を移動しても発注情報が失われることはありません。既に正しい位置にある行には一切触れません。
**CategoryやOrigin_Countryを修正した後は、このマクロを実行する習慣をつけてください。**

### 5.7.3 Substrateを原産国別にPO_Draftを分ける（Poland専用・Japan+China）

TTAF供給のSubstrateは、原産国によって発注書(PO_Draft)のシートを分けています。
Polandだけ専用シート（`PO_Draft_Substrate_Poland`）、Japan・Chinaはまとめて1枚のシート
（`PO_Draft_Substrate_JPN_CHN`）です。仕組みは`Category`（Chemical/Hazardous Chemical/
Substrateの3区分）とは独立した、`M_RawMaterials`の`Origin_Country`列（11列目）で
判定します。**汎用の受け皿シートは無い設計**のため、Origin_Countryが「Japan」「China」
「Poland」のいずれでもない品目（空欄・その他の国・未確認）は、意図的にどの`PO_Draft_*`
にも表示されません（発注していない・まだ原産国が確認できていない品目を、発注書のドラフト
に載せない、という運用判断です）。Origin_Countryは大文字小文字を区別せずに判定します
（「poland」でも「Poland」として扱われます）。

**既存ブックへの導入方法**（`macros/RefreshData_MaterialMgmt.bas`を貼り替えた後、
この順で一度だけ実行）:
1. `SetupSubstratePODraftByCountry` — `M_RawMaterials`への`Origin_Country`列追加、
   既存の`PO_Draft_Substrate`シートの`PO_Draft_Substrate_JPN_CHN`への改名（書式・データ
   をそのまま引き継ぐ）、`PO_Draft_Substrate_Poland`シートの新規作成（データ行は空の
   状態）をまとめて行う
2. `M_RawMaterials`で、各Substrate品目の`Origin_Country`欄に「Japan」「China」「Poland」の
   いずれかを入力する（発注していない・原産国が未確認の品目は空欄のままでよい）
3. `SyncPODraftCategories`を実行する（5.7.2章のマクロ。Origin_Countryも見るように
   拡張済みなので、これを実行すると各品目が正しい`PO_Draft_Substrate_*`へ振り分けられる。
   Japan・Chinaのどちらを入力しても`PO_Draft_Substrate_JPN_CHN`に入る）

`AddMaterial`でSubstrateの新規材料を追加する際は、原産国を尋ねるプロンプトが追加されて
おり、「Japan」「China」「Poland」のいずれかを入力すればそのまま対応する
`PO_Draft_Substrate_*`に登録されます（それ以外・空欄の場合はどの`PO_Draft_*`にも
登録されません。発注する際にOrigin_Countryを入力し直し、`SyncPODraftCategories`を
実行してください）。

発注書の発行（PDFではなくスナップショット用ブックとしての書き出し）は、`PO_Export.bas`の
`ExportSubstrateJPNCHN`・`ExportSubstratePoland`マクロで行います（`ExportChemical`と
同様の仕組みで、`PO_Issued`フォルダにそれぞれ`Substrate_JPN_CHN_Release_yyyymmdd_RevXX.xlsx`
等として保存されます）。

現状はこの2区分（Poland専用・Japan+China）での分離ですが、将来的に区分を変えたくなった
場合は、`POSheetNameForMaterial`関数（VBA）と`build_po_draft`の呼び出し（`build_soh.py`）に
同様の分岐を追加・変更するだけで対応できます。

### 5.7.4 発注書のレターヘッド形式レイアウトとFirm/Forecast色分け（`SetupPODraftLetterheadLayout`）

`PO_Draft_*`シートは元々、`build_soh.py`が生成するシンプルな罫線グリッド（TO/FROM欄は
仮の文字列、見出しは14行目、データは15行目から）でしたが、実運用では`PO_Draft_Hazardous`
シートに、CATALERのレターヘッド・実際のTO/FROM/CC欄・発行日/Issue Month/Firm Month・
Revision番号・基準週(WeekIndex)入力欄・SafetyStock/CurrentStockの参照欄などを手動で
作り込んだ、より実用的なレイアウトへ作り替えられました。以降、このレイアウトを
「レターヘッド形式レイアウト」と呼び、全`PO_Draft_*`シートの標準としています
（`build_soh.py`も新規ブック生成時から最初からこの形式で生成します）。

**レイアウトの行構成**（列は共通: B=Part Name, C=TTAF Code, D=CSA Code, E=UOM/Month/Year,
F=SafetyStock, G=CurrentStock, H〜T=週1〜13, U=Total）:

| 行 | 内容 |
|---|---|
| 8〜9行目 | TO（宛先） |
| 10行目 | N列=Firm Month: ／ P列=`=TEXT(H20,"mmmm")`（見出し1週目の月を表示） |
| 11行目 | CC（必要であれば入力） ／ N列=Revision ／ P列=Revision番号（名前付き範囲`PORevision`） |
| 13行目 | FROM（発行者） ／ N列=基準週(WeekIndex)ラベル ／ P列=基準週の値（名前付き範囲`BaseWeek`） |
| 14行目 | FROM（自社名） |
| 17〜18行目 | タイトル（結合セル） |
| 20行目 | 見出し1段目（Part Name/TTAF Code/CSA Code/Month-Year。20〜26行を縦結合） |
| 21〜24行目 | 非表示の補助行（WeekStart・WeekOfYear・空白スペーサー） |
| 25行目 | 見出し2段目（Week/SafetyStock/CurrentStock/週ラベル/Total） |
| 26行目 | 見出し3段目（UOM／Firm・Forecastの帯ラベル） |
| 27行目〜 | データ行（材料ごとに1行） |

**名前付き範囲`BaseWeek`・`PORevision`**: 基準週セル（P13）とRevisionセル（P11）は、
すべての数式・マクロからシート固有（ローカルスコープ）の名前付き範囲`BaseWeek`・
`PORevision`経由で参照します。実セルの位置がP13・P11から将来動いても、名前の参照先
だけ直せば全ての数式（`AppendPODraftRow`が新規追加する行、月/週見出しの数式）が
自動的に追従します。かつては`$P$7`をVBA側に直接埋め込んでいたため、基準週セルを
手動でP7からP13へ移動した際、既に登録済みだった一部の材料（例: ND TAC/CHEM-1280）の
数式だけが古い`$P$7`参照のまま取り残され、空欄のP7セルを参照し続けて発注数量が
常に0のまま更新されなくなる不具合がありました。

**修正した不具合（`SetupPODraftLetterheadLayout`で修正）**:
1. **月/週見出し(20行目)が消える不具合**: 元は1〜4週目・5〜8週目・9〜13週目という
   固定幅でセル結合し、直前のグループ最終週と月が違う時だけラベルを表示する方式に
   なっていましたが、月によって実際にまたがる週数は変わる（例: 2026年11月は5週に
   またがる）ため、固定幅の結合では境界がずれる月で見出しが消えてしまう不具合が
   必ず発生していました。結合をやめ、「週ごとに1セル、直前の週と月が違う時だけ表示」
   という本来の方式（元の`build_po_draft`と同じ）に統一しました。
2. **`$P$7`の取り残し**: 上記の通り、全データ行の`$P$7`・`$P$13`直接参照を
   名前付き範囲`BaseWeek`へ統一しました。
3. **Firm/Forecastの色分け**: Firm(1〜4週目)は赤系（背景`FFC1C1`・文字`C00000`）、
   Forecast(5〜13週目)は緑系（背景`EBF1DE`・文字`006100`）で、発注数量セルを塗り分けます。
4. **SafetyStock/CurrentStock(F/G列)の印刷対策**: 従来は列を非表示にすることで印刷対象
   から外していましたが、非表示を解除すると印刷にも写ってしまうリスクがありました。
   列の非表示は解除して常に参照できる状態に戻した上で、印刷範囲（`PageSetup.PrintArea`）
   自体からF・G列を除外する方式に変更しました（Excelの印刷範囲は複数の矩形を指定できる
   ため、`$A$1:$E$n,$H$1:$U$n`のように非表示に頼らず狙った列だけ除外できます）。

**`SetupPODraftLetterheadLayout`の動作**: 一度だけ実行する移行用マクロです。
①`PO_Draft_Hazardous`自身の上記の不具合を修正し、②`PO_Draft_Chemical`・
`PO_Draft_Substrate_JPN_CHN`・`PO_Draft_Substrate_Poland`の3シートに同じレイアウトを
複製します。複製時、TO/FROM/CC欄は`PO_Draft_Hazardous`の実際の宛先をそのままコピーせず
仮の文字列に戻します（カテゴリによって担当者・取引先が異なる可能性があるため、複製後に
実際の宛先を入力してください）。Revision・基準週(WeekIndex)は複製前の各シート自身の値を
引き継ぎます。データ行（材料一覧）は複製前の内容を使い回さず、`M_RawMaterials`の現在の
Category・Origin_Countryを基準に作り直します（`SyncPODraftCategories`と同じ判定基準）。
ロゴ画像・バナー等の装飾は、Excelのシートコピー機能により自動的に複製されますが、
バナーに日付文字列が手入力されている場合は、複製後に各シートで内容を確認・修正して
ください。既に移行済みの部分（シートごとに名前付き範囲`BaseWeek`の有無で判定）は
スキップするため、誤って複数回実行しても安全です。

**発注書の発行(`ExportPODraft`)とRevision**: `PO_Export.bas`の`ExportPODraft`は、
名前付き範囲`PORevision`があればそれを、無ければ（未移行の旧レイアウトのシート）
従来通り`P5`セルをRevision番号として読み取ります。レターヘッド形式レイアウトへの
移行後、`PORevision`を見ずに常に`P5`を読んでいると、常に空欄の`P5`を0扱いのまま読み、
発行のたびに`P5`（レターヘッドの空白セル）へ`1`を書き込んで静かに壊してしまうところ
でした。`SetupPODraftLetterheadLayout`実行後は、この心配なく`ExportChemical`等を
実行できます。

**中間体（生産される製品コード）側の追加・削除**: 材料（原材料）だけでなく、中間体（`PP_Grid`の
行、生産計画上の製品コード）の増減にも対応しています。

- **追加**: 何もしなくても自動対応済みです。`RefreshWeeklyBatches`が「Powder & Slurry & Pgm
  Plan」に新しい中間体を見つければ`PP_Grid`に自動で行を追加し、`RefreshBOM`が「Raw Material -
  Look Up」に新しい中間体×材料の組み合わせを見つければ`M_BOM`に追加した上で
  `Material_Detail`の該当材料ブロックにも内訳行を自動追加します（上記参照）。専用マクロを
  実行する必要はありません。
- **削除（`RemoveIntermediate`マクロ）**: 生産中止になった中間体名を入力すると、`PP_Grid`の
  該当行、`M_BOM`のその中間体を使う原単位の行（複数の材料で使われていれば全件）、
  `Material_Detail`側のその中間体の内訳行（この中間体を使っているすべての材料ブロックから）
  を削除します。原材料側のデータ（`T_Shipments`・`T_OpeningStock`・
  `T_StockCount`・実績ログ・`M_RawMaterials`）は削除しません。
  - 中間体の行を明示的に削除しなくても、生産計画から単に消えれば以降の週の「No. of
    batches」は0のまま更新されなくなり、使用量計算（原単位×バッチ数）も自動的に0になるため
    在庫計算が狂うことはありません。また`HideInactiveIntermediates`で表示上も折りたためます。
    `RemoveIntermediate`は、シートが年々肥大化するのを防ぐための**任意のクリーンアップ**です。

**注意点**:
- いずれの操作も**取り消せません**。実行前にファイルのバックアップ（コピー）を取ることを
  強く推奨します。
- `AddMaterial`で追加する行は常に各シートの一番下に追加され、既存の行の途中に挿入することは
  ありません（既存行の数式・参照がずれるのを避けるため）。
- `PO_Draft_*`の週次予測式はGrid_Stock内の行位置を`MATCH`で毎回動的に検索する方式のため、
  材料の追加・削除でGrid_Stockの行位置がずれても数式側が自動的に追従します。
- `PP_Grid`・`M_BOM`から中間体の行を削除しても、他の中間体・他の材料の計算には影響しません。
  `Grid_Requirement`・`Material_Detail`側の数式がすべて`MATCH`／構造化参照（テーブル名[列名]）で
  組まれており、固定の行番号を直接使っていないためです。

## 5.8 操作パネル（マクロボタンを月次手順の順に並べる）

**背景**: `Alt+F8`のマクロ一覧は、VBAモジュールの並びやマクロの記述順に関係なく、**常に
マクロ名のアルファベット順**で表示されます（Excelの仕様のため、設定で変更することは
できません）。月次の運用手順どおりの順番でマクロを実行したい場合は、`Alt+F8`に頼らず、
シート上にボタンを手順順に並べる方式にします。

**`操作パネル`シート**: このブックには、README のすぐ後ろに`操作パネル`シートを用意して
あります。以下の順に、番号・マクロ名・説明が並んだ行(緑色の帯)が並んでいます。

1. `RefreshWeeklyBatches` 〜 5. `RefreshShipments`（毎月・毎週の定型作業。上から順に実行）
6. `HideInactiveIntermediates` 〜 7. `ShowAllIntermediates`（任意の表示調整）
8. `AddMaterial` 〜 10. `RemoveIntermediate`（まれに使う。材料・中間体の追加/削除）

`JumpToSelectedWeek`は、シートのC1セル変更時に自動で呼ばれる想定のマクロ（引数が必須で
Alt+F8の一覧にも出ません）のため、ボタン一覧には含めていません。

**ボタンの割り当て方(手動での一度だけの作業。openpyxlではボタンを自動作成できないため)**:
1. `操作パネル`シートを開く
2. 「挿入」タブ→「図形」で、緑色の帯の行(例: 1行目の`RefreshWeeklyBatches`)に重なるように
   図形を描く（行の高さ・幅に合わせると綺麗に収まります）
3. 図形を右クリック→「マクロの登録」→対応するマクロ名(例: `RefreshWeeklyBatches`)を選択
4. 残り9個の行についても同様に繰り返す
5. お好みで、図形の塗りつぶし色や文字を調整してください（例: マクロ名をそのまま図形内の
   テキストにすると分かりやすくなります）

一度設定すれば、以降は`操作パネル`シートを開いて上から順にボタンをクリックするだけで、
月次の運用手順どおりに作業を進められます。

## 5.9 材料リスト・BOM・生産計画の元データについて（2026年7月改版）

**材料リスト(`M_RawMaterials`)とBOM(`M_BOM`)の元データ**: 「Raw Material - Look Up」ファイル
（`Catalyst Data Base`・`Slurry Data Base`・`Powder Data Base`・`Solution`の4シート）と、
「Raw materials daily check」（自社倉庫の現物確認シート）を突合して作成しています。
製造工程は「Powderを作る→Slurryを作る→Substrateに塗布してCatalyst(完成品)になる」の順で、
BOMも以下の対応関係で作成しています。

- `Powder Data Base`・`Slurry Data Base`・`Solution`: それぞれのシートの1バッチあたり使用量列
  （Powder/Slurry/SolutionいずれもM列）をそのまま採用
- `Catalyst Data Base`: **Substrateの行だけ**をF列（catalyst1個あたりの使用量。PP_Grid側で
  catalystは個数管理のためF列を使う。他の材料はSlurry/Powder/Solution側で既にバッチベースの
  計算式が存在するため、Catalyst Data Base側の直接行(化学品・スラリー参照行)は二重計上を
  避けるため使用しない）

**中間体が別の中間体の原料になっているケース**（例: TSZ-616・TSZ-938はSlurryだが、それ自体が
外部購入品ではなく、TSP-618等の別のSlurryの原料としてのみ使われ、「Powder & Slurry & Pgm Plan」
にも独自の週次バッチ計画を持たない）は、`M_RawMaterials`には一切登録しません（Dashboard等には
表示されません）。代わりに、`PP_Grid`のその中間体の週次セルに、`Grid_Requirement`と全く同じ
`SUMPRODUCT`パターンの数式（それを使う側の実バッチ数から逆算）を自動生成して埋め込みます。
これはビルド時にPythonが数式を1回書き込むだけで、以降はExcelの数式として完結するため、
中間体の増減やレシピの変更があっても再生成なしに追従します。

**固定週次消費量(`M_RawMaterials`の`固定週次消費量_要入力`列)**: Original Towel（梱包資材の
養生紙）のように、生産中間体の構成(BOM)とは無関係に毎週ほぼ一定量を消費する材料向けの入力欄
です。`Grid_Requirement`はBOM経由の計算結果にこの値を単純加算します。通常は0で、値を変える
場合はこのセルを直接書き換えるだけで反映されます(再生成不要)。

**生産計画(`RefreshWeeklyBatches`)の元データ**: 以前は「Powder & Slurry & Pgm Plan」が材料
ごとに別シート(40枚超)に分かれている前提でしたが、現在の運用では**単一シートに全中間体が
まとまった形式**（B列=中間体/完成品(Catalyst)/Solutionの名前、週初日ヘッダー行以降に週次の
生産量）に統一されています。`RefreshWeeklyBatches`もこれに合わせて書き直しています。

行の種類（中間体/完成品/Solution）は行番号ではなく**名前のパターンで自動判定**します。
`TSP-`・`TPP-`・`TSZ-`・`TVS-`・`VSP-`で始まれば中間体(Slurry/Powder)、`操作パネル`シートの
`T_SolutionNames`テーブルに載っている名前(略称。例: `20P`・`SH`・`SCH`)ならSolution(略称は
自動的に正式名`SOL-xxx`へ変換)、どちらでもなければ完成品(Catalyst)として扱います。この方式
により、ファイル側で行が増減してもマクロ側の行番号メンテナンスは一切不要です。新しい
Solutionが増えた場合だけ、`操作パネル`シートの`T_SolutionNames`テーブルに1行追加してください。

なお、Catalyst以外(Powder/Slurry/Solution)の週次生産量は**バッチ数**、Catalystのみ**生産個数**
で記録されています。前者はBOM側もバッチ単位の使用量(M列)、後者はcatalyst1個あたりの使用量
(F列、Substrateのみ)を使うことで、単位を揃えています。

## 5.10 発注数量の入力方法（PO_Draftの自動計算を廃止）

以前は`PO_Draft_Chemical`/`_Hazardous`/`_Substrate`が、基準在庫下限と`Grid_Stock`の在庫予測から
「いつ・いくつ発注すべきか」を自動計算していました（`MAX(0, 基準在庫下限 - その週の予測在庫)`）。
この自動計算は廃止し、**発注数量は`Material_Detail`の「Order(発注予定,kg)」行に、材料ごと・週ごと
直接手入力**する方式に変更しました。

- `Material_Detail`の各材料ブロックの「Order(発注予定,kg)」行（黄色の入力セル）に、発注したい
  週の列へkg数量を入力してください。
- `PO_Draft_*`は、入力された値を材料コード・週で突き合わせてそのまま転記するだけです（自動計算は
  一切しません）。基準在庫下限・現在庫（`SafetyStock`・`CurrentStock`列）はこれまで通り参考情報
  として表示されます。
- 内部的には、Material_Detailの「Order」行にだけ材料コードを複製した見えない列（週データ列の
  2つ右）を使い、`PO_Draft_*`側がその列を`MATCH`で特定して`INDEX`で値を拾っています。ブロックの
  長さ（使用中間体の数）は材料ごとに違いますが、この仕組みにより行番号のズレを気にせず正しく
  対応づけられます。
- 過去に存在した`T_PlannedOrders`シート（発注予定の完全手入力シート）は、この方式に置き換えた
  ことで役割が完全になくなったため、現在はワークブックから削除されています。

### 5.10.1 発注管理（Order行・PO_No行の自動追従）

`Material_Detail`の「Order(発注予定,kg)」行のすぐ下に「PO_No」行があります。運用の流れは:

1. 欲しいタイミングの週セルに、Orderの数量を入力する（この時点でGrid_Incoming・Grid_Stock・
   Dashboard・PO_Draft_*にすぐ反映されます。PO番号はまだ分からなくてOKです）。
2. PO番号が発行されたら、PO_No行の同じ週のセルに追記する（後からでOK）。
3. `RefreshShipments`を実行するたびに、そのPO番号の出荷状況(CSA Reportの`Status`列)に
   合わせて自動的に処理されます。
   - `Unconfirmed`でETAが未定(TBC): `Order_Month`(発注月)＋`M_RawMaterials`の
     `LeadTime_Weeks_要入力`から仮の週を計算し、その週にセルを移動する(あくまで仮の予測)。
   - `Unconfirmed`/`In-transit`でETAが判明: `T_Shipments`の`Effective_Week`
     (Latest ETA+14日＝CSA Reportの「2 week transit to TTAF」列を反映済み)の週に移動する。
     ETAが更新されるたびに追従する。
   - `TTAF Stock`(着荷確定): 最後に分かっている週で数量を確定し、PO_No行に「[済]」を付けて
     Grid_Incomingの計算対象から除外する。**数字自体は削除せず残る**ので、過去の発注履歴・
     発注頻度はいつでもMaterial_Detailを見返せば分かる。
4. 同じPO番号の出荷が複数行に分かれている場合（分割納品）、Order/PO_Noのセルもその週数分に
   自動的に分割される。
5. セルが自動的に移動・分割・確定した場合、変更前の内容をセルコメントに残す。

**過去の`T_PlannedOrders`時代の消込み条件（5.11章「消込み条件について」参照）との違い**:
当時は「TTAF実在庫が納品予定週に追いついたか」という時間経過ベースの判定で問題が起きた経緯が
あります。今回はCSA Reportの`Status`列という、その発注固有の状態を直接見るため、同じ問題は
起きません。分割納品も、当時のFIFO(先入れ先出し)方式ではなくPO番号そのもので突き合わせるため、
早着・延着どちらでも正確に対応できます。

Material_Detailにブロックが無い材料（BOMで使われない梱包資材等）や、PO_Noがどこにも
入力されていない出荷は、この自動追従の対象外です（Grid_Incomingは、そのようなケースでは
従来通り`T_Shipments`を直接参照します）。

**既存ブックへの導入方法**: 新しく`build_soh.py`でブックを作った場合はPO_No行・
Grid_Incomingの数式とも最初からこの形になっていますが、既存のライブブックには入っていません。
`macros/RefreshData_MaterialMgmt.bas`の`SetupOrderManagementMigration`を一度だけ
実行してください（Material_DetailへのPO_No行追加と、Grid_Incomingの数式書き換えを
まとめて行います。ブロック数によっては数十秒〜数分かかりますが、フリーズではありません）。

### 5.10.2 T_Shipmentsの一意キー修正（分割出荷の欠落防止）

以前の`T_Shipments`は「材料名＋PO番号」だけで行を一意に管理していましたが、実際のCSA Report
では同じ材料・同じPO番号の出荷が複数行に分かれる「分割出荷」が頻繁にあります（1つのPOが
5行に分かれているケースも普通にあります）。このキーだと分割出荷の行が同じ行として扱われ、
後から読んだ行が前の行を上書きしてしまい、実際には届いているはずの数量が静かに失われる不具合
がありました。実際のCSA Reportデータを確認したところ、55組の（材料名，PO番号）でこの
上書きが発生しうる状態が見つかっています。

この不具合を修正するため、`T_Shipments`にVessel・Container・Original_ETD列（10〜12列目）を
追加し、一意キーを「材料名＋PO番号＋コンテナ番号＋Original ETD」の複合キーに変更しました。
それでも完全に同じ組み合わせが複数行ある場合（同じコンテナに複数バッチが混載されている等、
ごく稀なケース）は、ファイル内の出現順の連番で最終的に区別します。これによりGrid_Incomingの
計算（Material_Detailにブロックが無い材料のフォールバック時）や`SyncMaterialDetailOrders`の
週別集計も、分割出荷分をすべて正しく合算するようになりました。

**既存ブックへの導入方法**: `macros/RefreshData_Shipments.bas`を貼り替えた後、
`AddShipmentSplitColumns`を一度だけ実行し、T_Shipmentsに上記3列を追加してください。
その後`RefreshShipments`を実行し直すと、以前は上書きされて消えていた分割出荷の行が、
複合キーで正しく区別されて自動的に追加され直されます（＝過去に失われていた数量が復元
されます）。なお、ごく稀な「コンテナ・Original ETDまで完全に同じ」ケースの連番による
区別は、TTAF側のレポート内での行の並び順が週次の再エクスポート間で安定していることに
依存しており、数学的な保証はありませんが、従来の「常に上書きで消える」状態からは大きく
改善しています。

### 5.10.3 Effective_Week数式の破壊バグ修正

上記の複合キー修正の検証中に、さらに別の重大な不具合が見つかりました。以前の
`RefreshShipments`は、既存行を更新する際に行全体を`.Range.Value`で読み込み→一部だけ
書き換え→行全体を`.Range.Value`で書き戻す、という処理をしていました。しかし
`Effective_Week`（8列目、着荷予定週を計算する数式）は数式セルのため、`.Value`で読み込むと
「その時点の計算結果の値」しか取得できません。それを行全体としてそのまま書き戻すと、
数式そのものが計算済みの値で上書きされて壊れてしまいます。

この不具合により、**一度でも`RefreshShipments`で更新された行は、Effective_Weekがその
更新時点の値のまま永久に凍結され**、その後Latest_ETA/Received_Dateがどれだけ変わっても
着荷予定週が一切追従しなくなっていました。「ETAが更新されるたびに追従する」という
発注管理機能の中核的な前提が、実質的に初回更新以降は機能しなくなる状態で、この不具合は
今回の機能追加より前から存在していました。

修正により、既存行の更新は8列目(Effective_Week)を避けて前半(1〜7列)・後半(9〜12列)の
2つに分けて書き戻すようにし、数式に一切触れないようにしました。

**既存ブックへの導入方法**: この修正はコードの貼り替えだけでは、過去に既に壊れてしまった
行までは直りません。`AddShipmentSplitColumns`（5.10.2章）を実行すると、列の追加とあわせて
T_Shipments全行のEffective_Week数式を正しい状態に一括で復元するようにしてあるので、
`RefreshData_Shipments.bas`貼り替え後は`AddShipmentSplitColumns`を一度実行するだけで、
この不具合の影響も含めて解消されます（既に5.10.2章の手順を実施済みでも、もう一度
実行して問題ありません＝安全に再実行できます）。

### 5.10.4 T_Shipmentsの行数について（自動では削除されない）

`T_Shipments`は、行を自動的に削除する仕組みを持っていません。`RefreshShipments`は複合キーが
一致する行があれば上書き更新、無ければ新規追加するだけで、古い行を消す処理は無いためです
（着荷済みのTTAF Stock分も履歴として残す、という5.10.1章の要件どおりの設計です）。

ただし、CSA Reportを取り込むたびに全577行前後がそのまま新規追加されていくわけではありません。
既に取り込み済みの出荷は複合キーで一致するため上書き更新されるだけで、行数が増えるのは
「本当に新しく登場した出荷（新規PO・新しい分割バッチ）」の分だけです。実質的には
「これまでに存在したすべての出荷の履歴ログ」として、新規分だけがゆっくり積み上がっていく
形になります。

一方で、TTAF側のCSA Report自体が、十分古く完了した出荷行を将来的にレポートから除外する
可能性はあります。その場合、その行はCSA Reportに登場しなくなるため`RefreshShipments`側でも
一致する行が無くなり、以後は一切更新されないまま`T_Shipments`に残り続けます（削除されません）。
結果として、`T_Shipments`は長期的には緩やかに増え続けます。現状は特に対応していませんが、
将来的に行数が多くなりすぎてパフォーマンスが気になるようであれば、「[済]が付いてから
一定期間(例: 2年)経過した行を別の履歴シートに退避する」ような、月次程度のアーカイブ用マクロを
追加で用意することは可能です。必要になった時点でご相談ください。

### 5.10.5 移行直後の二重計上リスクとクリーンアップマクロ

`AddShipmentSplitColumns`実行後、最初に`RefreshShipments`を実行すると、移行前から存在していた
行（Container・Original_ETDが空欄）は新しい複合キーと一致しないため、実質すべて「更新」ではなく
「追加」として扱われます（"追加:○件、更新:0件"という結果になるのはこのためで、想定どおりの
動作です）。

この結果、同じ材料・PO番号について「移行前の古い行」と「移行後の新しい行」が両方T_Shipments内に
残ることになります。もしそのPO番号がMaterial_Detail側でまだ「[済]」で確定していなければ、
`SyncMaterialDetailOrders`は古い行・新しい行の両方の数量を合算してしまい、実際より多い数量を
発注予定として二重計上してしまう可能性があります（既に「[済]」で確定済みのPO番号は、
そもそも同期処理の対象から外れるため元々問題になりません）。

これを防ぐため、`CleanupOrphanedPreSplitShipmentRows`マクロを用意しました。同じ材料+PO番号で
新形式の行が既に存在する「移行前の古い行」だけを判別して削除します（実データで確認した限り、
コンテナ番号・Original ETDが両方空欄になることは無いため、「両方空欄」で移行前の行を確実に
判別できます）。まだ新形式の行が無いPO番号（今回のCSA Reportにまだ登場していない出荷）の
行には触れません。

**実行手順**: `AddShipmentSplitColumns`実行後、最初に`RefreshShipments`を実行したら、続けて
`CleanupOrphanedPreSplitShipmentRows`を一度実行し、その後もう一度`RefreshShipments`を実行して
Material_Detailを最新状態に同期し直してください（このクリーンアップは移行直後の一過性の
対応のため、それ以降は実行不要です。誤って複数回実行しても、既にクリーンな状態であれば
「削除対象なし」と表示されるだけで安全です）。

## 6. 自動反映の仕組み

- **生産計画が変わったとき**: `RefreshWeeklyBatches`実行後、`Grid_Requirement` → `Grid_Stock` →
  `Dashboard` → `PO_Draft_*` → `Material_Detail`まで自動的に再計算されます。
- **原単位が変わったとき**: `RefreshBOM`実行後、同様に自動反映されます。`M_BOM`・`PP_Grid`は
  `Grid_Requirement`から直接参照されているため、全く新しい中間体×原材料の組み合わせが増えた
  場合でも、ブックの再生成は不要です。さらに`RefreshBOM`は、`Material_Detail`の該当材料ブロックに
  その中間体の内訳行（No. of batches／使用量(kg)）が無ければ自動的に追加します
  （`SyncMaterialDetailIntermediates`）。
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
再設計しました（計算量を約200分の1に削減）。

**`RefreshWeeklyBatches`実行時に必ず強制終了する不具合について**: この修正後も「Alt+F8 →
RefreshWeeklyBatches実行」のタイミングで確実にExcelが強制終了する不具合が報告されました。
原因は、マクロが「Powder & Slurry & Pgm Plan」の各シート（40枚超）を1セルずつ`.Cells(r,c).Value`
で読み取っていたことです。Excelのシートは、過去の書式設定の残骸等により「使用範囲(UsedRange)」
が実データよりはるかに大きく認識されていることが多く、これがシート数の多さと掛け合わさると
COM通信の呼び出し回数が数十万〜数百万回に達し、Excelが長時間「応答なし」になり強制終了する
（ように見える）動きを引き起こします。対策として、各シートのデータを`sh.Range(...).Value`で
1回だけ配列にまとめて読み込み、以降はメモリ上の配列だけを参照する方式に書き換え、あわせて
使用範囲の異常な肥大化に備えた上限（500行×200列）も設けました。同様の問題を防ぐため、
`RefreshBOM`が使う`ProcessLookupFlatSheet`/`ProcessLookupCatalystSheet`にも同じ対策を適用済みです。

**`RefreshSelfStock`実行時にも強制終了する不具合について**: 同じ理由（1セルずつの読み書き）が
`RefreshSelfStock`にもありました。対象シート「Stock」の読み取りを同様に1回の配列読み込みに
変更し、さらに`T_SelfStock`/`T_TTAFStock`への書き込み(`UpsertStockRow`)も、呼び出すたびに
テーブル全行をセル単位でスキャンしていたのを、(Part Name, WeekIndex)→行番号のDictionaryを
1回だけ作って参照する方式に変更しました（`RefreshTTAFStock`も同様）。これらのテーブルは
運用を重ねるほど行数が増えるため、この修正は将来的な速度低下の予防にもなります。

**同種の負荷がないか全体を見直した結果（追加の予防的修正）**: 上記2件を機に、同じパターン
（表が育つほど遅くなる／週ごとに同じ検索を繰り返す）が他に残っていないか確認し、以下も
あわせて修正しました。いずれも計算結果は変更前後で完全一致することをLibreOfficeで確認済みです。

- `Grid_Stock`の在庫ロールフォワード式: `T_StockCount`/`T_SelfStock`/`T_TTAFStock`の該当有無・
  値の取得に`SUMPRODUCT`のブール配列積を使っていましたが、これは対象表が育つほど遅くなる
  上記と同じ性質を持っていました。`T_SelfStock`/`T_TTAFStock`は(Part Name,WeekIndex)ごとに
  上書き更新される設計のため、定常状態では行数が「原材料数×週数(101×104≈10,504)」程度で
  頭打ちになる見込みですが、それでも将来的な速度低下を避けるため、ネイティブ関数の
  `COUNTIFS`/`SUMIFS`（この環境で構造化参照との組み合わせが正しく動作することを確認済み）に
  置き換えました。
- `Dashboard`の自社在庫/TTAF在庫実績表示（`LOOKUP`の「最後に一致した行」トリック）: 参照範囲を
  `$A$2:$A$2000`に固定していましたが、上記の理論上限（約10,504行）を下回るため、
  `T_SelfStock`/`T_TTAFStock`が育つと数ヶ月程度でこの上限を超え、最新実績が表示されなくなる
  （静かに古いデータのまま止まる）不具合になり得ました。`12,000`行に引き上げて修正しています。
- `Material_Detail`の「No. of batches」行: `PP_Grid`内の該当行を`MATCH`で探す処理を、週ごと
  （104回）に毎回やり直していました（711組×104週＝約7万4千回のMATCH）。`M_BOM`の`PPGridRow`と
  同じ考え方で、行位置を1回だけMATCHしてヘルパー列(最終週列の右隣)にキャッシュし、以降は
  そのキャッシュ値をINDEXで使い回す方式に変更しました（列位置はPP_Gridの列並びが固定のため
  MATCH不要、週番号からそのまま計算できます）。

**それでも`RefreshWeeklyBatches`実行時に強制終了する不具合について（VBA側の直し忘れ）**:
上記のExcel数式の見直しの際、`macros/RefreshData_ProductionPlan.bas`側にも同じパターン（`UpsertBomRow`が
M_BOMへの書き込みのたびに全711行以上をセル単位でスキャンする、`FindOrAddIntermediateRow`が
中間体ごとに毎回`.Find()`を呼ぶ）が残っていると気づいていながら、そちらの修正を実装し忘れて
いました。結果として`RefreshWeeklyBatches`実行時の強制終了が再発しました。`RefreshWeeklyBatches`
の実行開始時に、PP_Grid(中間体名→行番号)とM_BOM(Intermediate|Part Name→行番号)のDictionaryを
1回だけ作っておき（`BuildNameIndex`/`BuildPairIndex`）、以降はそれを参照する方式に修正しました。
あわせて`RefreshBOM`のインデックス構築、`WeekIndexForDate`の週特定処理も同じ考え方で1回の
配列読み込みにまとめています。

なお`BuildNameIndex`は、元の`.Find()`がExcelの既定動作として大文字/小文字を区別しない検索
だったのに対し、単純にDictionaryへ置き換えると区別してしまう(表記ゆれのある中間体名を
別物として扱い、実行のたびに重複行が増えていく)ため、`CompareMode=vbTextCompare`を明示的に
設定し、元と同じ大文字/小文字を区別しない挙動を保っています。

実際のExcelでの動作確認をお願いします。

**`T_SelfStock`/`T_TTAFStock`が縦に積み上がり続ける不具合について**: 上記のAnchorYear対応
（WeekIndexをDateからのライブ計算にした修正）の副作用として、日次で`RefreshSelfStock`を
実行するたびに全材料分の行が新規追加され、Part Nameが何行も重複して積み上がっていく
不具合がありました（以前はWeekIndexをキーにしていたため同じ週内なら自動的に上書きされて
いましたが、Date自体をキーにしたことでこの「週単位でまとめる」効果が失われていました）。
対策として、生データの保存先を「その週の月曜日（実際の暦日から計算。AnchorYearには依存
しない）」をキーにする方式に変更し、同じ週内の複数回の取り込みが1行にまとまるようにしました。

あわせて、目に見える`T_SelfStock`/`T_TTAFStock`シート自体も再設計しました。生データは
非表示の`T_SelfStock_Log`/`T_TTAFStock_Log`（実施日ベース、AnchorYearを何度・どんな頻度で
進めても壊れない）に保存し、目に見える方は`Dashboard`と同じ「材料×週」のグリッド形式に
変更、値はすべて生データからのSUMIFS数式で毎回計算し直す設計にしています。これにより、
`Grid_Stock`・`Dashboard`・`Material_Detail`側の参照も、行数が育ち続ける表へのSUMIFS/COUNTIFS
から、固定104列への直接セル参照に変更でき、以前の「表の成長に伴う速度低下」の懸念自体も
なくなりました（`Grid_Stock`側の該当有無判定も、以前は「記録が無い週」と「0と記録された週」
を区別するために別途COUNTIFSが必要でしたが、グリッド側で「記録が無い週は空欄」にする設計に
したことで、単純な空欄判定だけで済むようになっています）。

**RefreshBOM/RefreshWeeklyBatches実行時のエラーの原因調査について**: 実際のExcelでの検証で、
`ErrHandler`内の`On Error Resume Next`が（VBAの仕様として）`Err`オブジェクトを自動的にクリア
してしまい、エラーメッセージが常に空欄で表示される不具合が見つかりました。エラー番号・内容を
`On Error Resume Next`より前に変数へ退避するよう修正しています。この修正で判明した実際の
エラーは2つありました。1つは`ListObject.DataBodyRange`が`ListRows.Add`直後に不安定に`Nothing`
を返すことがある既知のクセによるもので、既存行の更新をすべて`ListRows(行番号).Range.Cells(...)`
という安定した書き方に統一しました。もう1つは、取込元ファイルが「読み取り専用を推奨」設定の
ファイルだと、`Application.DisplayAlerts`を明示的に`False`にしていない場合、`Workbooks.Open`実行時に
確認ダイアログが処理を妨げ、結果的に`srcWb`が正しく取得できなくなることがあるというもので、
各Refresh系マクロの`Workbooks.Open`前後に`DisplayAlerts`の抑制・復元を追加しています。

**T_SelfStock_Log/T_TTAFStock_LogのWeekIndex列が空欄になる不具合について**: 上記の再設計後、
VBAで新規追加した行のWeekIndex(数式列)が数式ごと空欄のままになり、目に見える
T_SelfStock/T_TTAFStockのグリッドに何も表示されない不具合が実際に報告されました。原因は、
「新しい行を追加すればExcelのテーブル機能が既存行と同じ数式を自動的に複製する」という前提が、
UI上で手動追加した場合の挙動であり、VBAの`ListRows.Add`経由では複製が保証されないためでした。
新規行では直前行のWeekIndexの数式を`FormulaR1C1`で明示的にコピーするよう修正しています
(相対参照はコピー先の行に合わせて自動調整される)。

**TTAF供給材料の在庫予測の紆余曲折について**: 当初、T_Shipments(Status=TTAF Stock)の
Received_Dateを「弊社への入庫日」と誤解していたため、TTAF供給材料についてこれをそのまま
入庫として合算すると、TTAF倉庫側で既にカウント済みの在庫を二重計上してしまう、という問題が
ありました。この理解にもとづき、一時期はTTAFへのCall Off依頼書を取り込む`T_TTAFCallOff`
シートを新設し、TTAF供給材料の自社側残高・TTAF側残高を別々にロールフォワードして合算する
方式を実装し、その後CSA Reportが毎週月曜に届く運用に合わせて「実績が無い週は入庫を見込まず
前週－消費のみ」という簡易版に縮小しました。

しかし、その後T_ShipmentsのStatus=TTAF Stockは実際には「TTAFが外部の仕入先から新しく仕入れて
TTAF倉庫に到着する」実績・予定であり、TTAF倉庫内で場所を移しただけの動きではないと判明した
ため、上記の「入庫を見込まない」という簡易版では、確定済みの入荷予定まで無視してしまい、
将来週（CSA Reportがまだ届いていない先の週）のTTAF供給材料の予測在庫が実際より少なく
（先細りに）算出される問題が新たに発生しました。**最終的に**、TTAF供給材料の特別扱いは
撤廃し、全材料共通の「前週＋入庫－消費」に統一しています。TTAF倉庫内の場所の移動
（Call Off）は合計在庫には影響しないため、そもそも合計値の二重計上の心配はありませんでした。

**CSA Reportの日付が1週間ズレる不具合について**: CSA Reportは毎週月曜（祝日の場合は翌営業日）
に**前週分**の実績が届く運用です。`Stock invoiced to CSA`シートのF4セルの日付はレポートが届いた
月曜日の日付ですが、その数値は実際には前週金曜営業終了後時点の在庫を表すため、そのままF4の
日付で週Noを判定すると1週間ずれて記録されてしまいます。`RefreshTTAFStock`はF4の日付から7日
引いてから週Noを判定するよう修正しています（月曜・金曜どちらも同じ月〜日の週に属するため、
週No自体は7日引いても3日引いても変わりませんが、記録する日付を他の実績と同じ週の月曜に
揃えるため7日引いています）。

**計画中の発注（T_PlannedOrders）とMaterial_Detailの「Order」行について（※この節は過去の設計の
記録です。`T_PlannedOrders`シートは現在ワークブックから削除されており存在しません。5.10章の
通り、Orderの値はMaterial_Detailに直接手入力する方式に変更されています）**: TTAF供給材料は
発注から着荷まで4〜6ヶ月かかり、常に複数件の発注が並行して進みます。当初PO_Draftにボタンを
設けて発注を確定する案も検討しましたが、TTAFが希望通りの数量を用意できるとは限らず後から
数量を手直しする必要があること、また注文書のドラフト自体は編集したくないという要望から、
`T_PlannedOrders`という完全手入力のシートに変更しました（材料名・数量・納品予定日・発注月）。
`Material_Detail`の自社在庫実績と合計在庫の間に「Order」行を追加し、該当週にセル内テキストで
「数量 (m月発注)」の形で表示します。

二重計上の防止は、当初「実データが揃ったら手動で行を削除する」運用も検討しましたが、
消し忘れによる二重計上を懸念する声があったため、完全に数式で自動化しました。同じ材料・
同じ発注月（`Order_Month`）の実データが`T_Shipments`に現れたら、`T_PlannedOrders`側の
`IsReconciled`/`EffectiveQty`列が自動的に0になり、Material_Detailの表示からも自動的に消えます。

**消込み条件について（設計の見直し）**: 当初は「TTAFの実在庫（T_TTAFStock）が納品予定週に
追いついたら消込む」という条件も併用していましたが、これは誤りでした。CSA Reportは毎週、
その注文とは無関係に材料ごとの在庫実績を更新し続けるため、「納品予定週以降に実績データが
存在する」は実質「予定週を過ぎて次のCSA Reportが届いたか」という時間経過の判定にしかならず、
その特定の注文が実際に届いたかどうかとは無関係でした。早着の場合は既に届いているのに
表示が予定週まで残るだけで実害は小さいですが、延着の場合はまだ届いていないのに予定週超過
だけでOrder欄から消えてしまい、見落としにつながる危険がありました。そのため、この条件は
撤廃し、`T_Shipments`にその注文の実データそのものが現れた場合のみ消込む方式に変更しています
（早着・延着どちらでも、実データが確認できるまでOrder欄に表示され続けます）。

**分割納品への対応**: 1つの発注が複数回に分けて納品されるケース（例: 500kg発注が
week1に200kg・week3に200kg・week10に100kg、と3回に分かれて届く）では、単純に「同じ材料・
発注月の実データが1件でもあれば全部消す」だと、week1分の実データが1件届いただけで
week3・week10分もまとめて消えてしまいます。これを避けるため、`IsReconciled`は
「納品予定日が早い順の累積計画数量」と「`T_Shipments`側の累積確定数量」を比較し、
確定数量でカバーされている行から順に（先入れ先出しで）消込む方式にしています。上の例では、
200kg分の実データが届いた時点でweek1行だけが消え、week3・week10行はまだ表示され続けます。
（実際にテストデータで動作確認済みです。）

`T_Shipments`への実データ取込みは、CSA Reportの`Shipping Schedule`シートを丸ごと取り込む
`RefreshShipments`マクロで行います。当初はPO No単位で1件ずつ検索する案でしたが、
発注が常に4〜6件並行するため毎回複数回実行するのは非現実的との指摘を受け、他のRefresh系
マクロと同様「ファイルを選ぶだけで全件まとめて反映」する方式に変更しました。

## 8. 要確認・要入力の項目

- **`M_RawMaterials`の`基準在庫下限_要入力` / `基準在庫上限_要入力` / `LeadTime_Weeks_要入力`**:
  仮値(0)です。実際の基準在庫水準に置き換えてください。Dashboardの週次セルの赤(下限未満)/
  緑(範囲内)/青(上限超)の色分けに使われます。
- **Material_Detailの`MOQ`列**: 全材料とも未入力です。手書きで入力してください（数式化していません）。
- **`M_RawMaterials`の`Category`**: 機械的に判定したものです。実際の危険物区分と一致しているか
  確認してください。
- **Substrates（基材）**: 「Powder & Slurry & Pgm Plan」内の"Japan GPF Substr"/"China Substr"/
  "Poland GPF Substr"（およびEster Film/PP Film等、1シート1品目のフィルム系シート）から
  週次使用量(=完成品Catコードの受注数量×1個あたり使用量)を取り込み、`M_RawMaterials`に
  Category="Substrate"として統合済みです。中間体を経由する化学原料と異なり、substrateは
  完成品コード(Cat)が直接「中間体」の役割を果たします（`M_Intermediates`にType="Cat"として
  登録）。`PO_Draft_Substrate_JPN_CHN`/`_Poland`にも、`Origin_Country`が入力済みの
品目については実データが反映されています（5.7.3章参照）。
- **`T_OpeningStock`（期首在庫）**: 現状すべて0です。運用開始週の実在庫を入力してください
  （`T_SelfStock`・`T_TTAFStock`に実績があれば、その週以降は自動でリセットされます）。
- **`RefreshData_*.bas`・`Q_Shipments.pq`**: 未検証です。動作確認の結果を教えてください。
- **`T_SelfStock`/`T_TTAFStock`シートへの手入力はしないでください**: これらは数式のみのグリッド
  表示です。値は非表示の`T_SelfStock_Log`/`T_TTAFStock_Log`から自動計算されるため、直接
  上書きしても`RefreshSelfStock`/`RefreshTTAFStock`を再実行すると数式に戻ります。

## 9. 今後の拡張候補

- **Min/Max（週数ベースの安全在庫）モデル**: 現状は単一しきい値のみですが、「N週分の使用量」を
  基準にしたMin/Max運用に拡張することも可能です。
