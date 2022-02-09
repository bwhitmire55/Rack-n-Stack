Sub Main
    Dim storeMgr As CStoreMgr
    Dim catMgr As CCategoryMgr
    Dim sheetRef As Sheet

    Sub GenerateWorksheet_Click()

        ' Check if the Setup worksheet exists
        If Not WorksheetExists(wsSetupStr) Then
            MsgBox ("Sheet '" & wsSetupStr & "' not found. Please create it to continue.")
            Exit Sub
        End If

        ' Check if the Weekly Rankings worksheet already exists
        If WorksheetExists(wsWRStr) Then
            MsgBox ("Sheet '" & wsWRStr & "' already exists. Delete it if you wish to continue.")
            Exit Sub
        End If

        ' Initialize class managers
        Set storeMgr = New CStoreMgr
        Set catMgr = New CCategoryMgr

        SetTotalStores (storeMgr.GetTotalStores)
        SetTotalCategories (catMgr.GetTotalCategories)

        Set sheetRef = New Sheet

        ' Set the categories ranges, apply formatting, and insert values
        Dim i As Integer
        For i = 1 To catMgr.GetTotalCategories()
            Dim row As Integer
            Dim Column As Integer

            row = 4 + ((4 + storeMgr.GetTotalStores()) * Int((i - 1) / 4))
            Column = (3 * ((i - 1) Mod 4)) + 2

            Call catMgr.SetCategoryRange(i, Range(Cells(row, Column), Cells(row + storeMgr.GetTotalStores() + 2, Column + 1)).Address)
            catMgr.ApplyFormatting (i)
            catMgr.SetTotalValue (i)
            Call catMgr.SetStoreNames(i, storeMgr)
            Call catMgr.ApplyConditionalFormatting(i, storeMgr)
            catMgr.SortData (i)
        Next i
    End Sub

    Sub OpenCategoriesEditor_Click()
        AddCategoryUF.AddCategoryUF_Initialize
        AddCategoryUF.Show
    End Sub
End Sub