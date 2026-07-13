// ============================================================================
// Q_Shipments
//
// 目的: 「CSA_Report_DD.MM.YYYY.xlsx」を週次のフォルダから自動的に最新のもの
//       を見つけて開き、"Shipping Schedule" シートから発注・ETA・着荷実績を
//       取り込む。※ 元ファイルのフォーマットは一切変更しません。
//
// 【注意】この環境ではPower Queryの実行検証ができないため、このMコードは
//        未検証です。まずは貴社のExcelで動作確認をお願いします。
//        うまく動かない場合は、エラー内容を教えてください。
//
// 【導入手順】Q_ProductionPlan.pq と同様です。クエリ名は "Q_Shipments" にする。
// ============================================================================
let
    RootFolder = "C:\SOH_Inputs\CSAReport\",  // ← 実際のフォルダパスに書き換えてください

    SourceFiles = Folder.Files(RootFolder),
    ExcelFiles = Table.SelectRows(
        SourceFiles,
        each Text.Lower(Text.End([Name], 5)) = ".xlsx"
            and Text.Contains(Text.Lower([Name]), "csa_report")
    ),
    SortedByDate = Table.Sort(ExcelFiles, {{"Date modified", Order.Descending}}),
    LatestFile = SortedByDate{0},

    Workbook = Excel.Workbook(LatestFile[Content], null, true),
    ShipSheet = Workbook{[Item = "Shipping Schedule", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(ShipSheet, [PromoteAllScalars = true]),

    Selected = Table.SelectColumns(
        Promoted,
        {"CSA Product Code", "CSA PO No.", "Confirmed Order Qty", "Latest ETA", "Received At TTAF", "Status"}
    ),
    Renamed = Table.RenameColumns(Selected, {
        {"CSA Product Code", "RM_Code"},
        {"CSA PO No.", "PO_No"},
        {"Confirmed Order Qty", "Confirmed_Qty"},
        {"Latest ETA", "Latest_ETA"},
        {"Received At TTAF", "Received_Date"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Confirmed_Qty", type number}, {"Latest_ETA", type date}, {"Received_Date", type date}
    })
in
    Typed
