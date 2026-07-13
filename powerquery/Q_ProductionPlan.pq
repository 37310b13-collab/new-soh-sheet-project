// ============================================================================
// Q_ProductionPlan
//
// 目的: 「Plan Increase and decrease - YYYY-MM-Vn.xlsx」を毎月のフォルダから
//       自動的に最新のものを見つけて開き、最新月の "All - MM" シートを
//       中間体×週の縦持ち(long)テーブルに変換する。
//       ※ 元ファイルのフォーマットは一切変更しません。読み取るだけです。
//
// 【注意】この環境ではPower Queryの実行検証ができないため、このMコードは
//        未検証です。まずは貴社のExcelで動作確認をお願いします。
//        うまく動かない場合は、エラー内容を教えてください。
//
// 【導入手順】
//   1. Excelで データ(Data) → データの取得(Get Data) → その他のソースから
//      (From Other Sources) → 空のクエリ(Blank Query)
//   2. クエリの詳細設定(Advanced Editor)を開き、下記コードを丸ごと貼り付け
//   3. 1行目の RootFolder を実際のフォルダパスに書き換える
//      （「Plan Increase and decrease」の月次フォルダが並ぶ「親フォルダ」を指定）
//   4. 完了して読み込む(Close & Load) → テーブルとして新しいシートに読み込み
//   5. クエリ名を "Q_ProductionPlan" にする
//   6. 以後は データ → すべて更新(Refresh All) で最新月のファイルを自動取得
// ============================================================================
let
    RootFolder = "C:\SOH_Inputs\PlanIncreaseDecrease\",  // ← 実際のフォルダパスに書き換えてください

    SourceFiles = Folder.Files(RootFolder),
    ExcelFiles = Table.SelectRows(
        SourceFiles,
        each Text.Lower(Text.End([Name], 5)) = ".xlsx"
            and Text.Contains(Text.Lower([Name]), "plan increase")
    ),
    SortedByDate = Table.Sort(ExcelFiles, {{"Date modified", Order.Descending}}),
    LatestFile = SortedByDate{0},

    Workbook = Excel.Workbook(LatestFile[Content], null, true),
    PlanSheets = Table.SelectRows(
        Workbook,
        each [Kind] = "Sheet" and Text.StartsWith([Name], "All - ")
    ),
    AddMonthNo = Table.AddColumn(
        PlanSheets, "MonthNo",
        each try Number.FromText(Text.End([Name], 2)) otherwise -1
    ),
    SortedSheets = Table.Sort(AddMonthNo, {{"MonthNo", Order.Descending}}),
    LatestSheetData = SortedSheets{0}[Data],

    // シート構造: row1=月ラベル, row2=週初日(日付), row3=週番号(数値/区切り列は"-"), row4以降=データ
    // row3が数値の列だけを実データ列として扱う（区切り列を自動除外）
    Rows = Table.ToRows(LatestSheetData),
    DateHeaderRow = Rows{1},
    WeekNoRow = Rows{2},
    DataRows = List.Skip(Rows, 3),
    ValidCols = List.Select(List.Positions(WeekNoRow), each WeekNoRow{_} is number),

    ExpandedRows = List.Combine(
        List.Transform(
            DataRows,
            (r) => List.Transform(
                ValidCols,
                (c) => [Intermediate = r{0}, WeekStart = Date.From(DateHeaderRow{c}), Batches = r{c}]
            )
        )
    ),
    ResultTable = Table.FromRecords(ExpandedRows),
    Final = Table.TransformColumnTypes(ResultTable, {{"Batches", type number}, {"WeekStart", type date}})
in
    Final
