Sub Sheet
    '----------------------------------------------
    ' Constants
    '----------------------------------------------
    Const ColumnInitWidth As Double = 1.57
    Const ColumnSpacerWidth As Double = 3.43
    Const ColumnStoreWidth As Double = 24.14
    Const ColumnValueWidth As Double = 13.14

    Const RowInitSpacerHeight As Double = 12.75
    Const RowSpacerHeight As Double = 38.25
    Const RowTitleHeight As Double = 30.75

    '----------------------------------------------
    ' Member Variables
    '----------------------------------------------
    Private m_Worksheet As Worksheet

    '----------------------------------------------
    ' Constructor / Destructor
    '----------------------------------------------
    Private Sub Class_Initialize()

        Set m_Worksheet = Sheets.Add

        m_Worksheet.name = wsWRStr
        m_Worksheet.PageSetup.Orientation = xlPortrait
        m_Worksheet.PageSetup.Zoom = False
        m_Worksheet.PageSetup.FitToPagesWide = 1
        m_Worksheet.PageSetup.FitToPagesTall = 1
        m_Worksheet.PageSetup.CenterHorizontally = 1
        m_Worksheet.PageSetup.CenterVertically = 1
        m_Worksheet.PageSetup.PrintArea = _
            "A1:M" & (Util.GetTotalCategories / 4) * (Util.GetTotalStores + 4) + 3

        m_Worksheet.Columns("A").ColumnWidth = ColumnInitWidth
        m_Worksheet.Columns("D").ColumnWidth = ColumnSpacerWidth
        m_Worksheet.Columns("G").ColumnWidth = ColumnSpacerWidth
        m_Worksheet.Columns("J").ColumnWidth = ColumnSpacerWidth

        m_Worksheet.Columns("B").ColumnWidth = ColumnStoreWidth
        m_Worksheet.Columns("E").ColumnWidth = ColumnStoreWidth
        m_Worksheet.Columns("H").ColumnWidth = ColumnStoreWidth
        m_Worksheet.Columns("K").ColumnWidth = ColumnStoreWidth

        m_Worksheet.Columns("C").ColumnWidth = ColumnValueWidth
        m_Worksheet.Columns("F").ColumnWidth = ColumnValueWidth
        m_Worksheet.Columns("I").ColumnWidth = ColumnValueWidth
        m_Worksheet.Columns("L").ColumnWidth = ColumnValueWidth

        m_Worksheet.Rows("1").RowHeight = RowInitSpacerHeight
        m_Worksheet.Rows("2").RowHeight = RowTitleHeight
        m_Worksheet.Rows("3").RowHeight = RowSpacerHeight

        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 60

        With Range("B1:L2")
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 28
            .Font.name = "Arial Black"
        End With

        m_Worksheet.Cells(1, 2).Value = "RACK & STACK FOR"
    End Sub

    Private Sub Class_Terminate()

    End Sub
End Sub