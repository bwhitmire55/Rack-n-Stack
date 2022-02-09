Sub CCategory
    '----------------------------------------------
    ' Member Variables
    '----------------------------------------------
    Private m_Name As String
    Private m_LookupValue As String
    Private m_ValueType As String
    Private m_Condition As String
    Private m_TotalType As String
    Private m_SortType As String

    Private m_CategoryRange As Range

    '----------------------------------------------
    ' Constructor / Destructor
    '----------------------------------------------
    Private Sub Class_Initialize()
        Set m_CategoryRange = Range("A1")
    End Sub
 
    Private Sub Class_Terminate()
 
    End Sub

    '----------------------------------------------
    ' Getters / Setters
    '----------------------------------------------
    Public Function GetCategoryName()
        GetCategoryName = m_Name
    End Function

    Public Sub SetCategoryName(name As String)
        m_Name = name
    End Sub

    Public Function GetCategoryLookupValue()
        GetCategoryLookupValue = m_LookupValue
    End Function
 
    Public Sub SetCategoryLookupValue(lValue As String)
        m_LookupValue = lValue
    End Sub

    Public Function GetCategoryValueType()
        GetCategoryValueType = m_ValueType
    End Function
 
    Public Sub SetCategoryValueType(vType As String)
        m_ValueType = vType
    End Sub

    Public Function GetCategoryCondition()
        GetCategoryCondition = m_Condition
    End Function

    Public Sub SetCategoryCondition(condition As String)
        m_Condition = condition
    End Sub

    Public Function GetCategoryTotalType()
        GetCategoryTotalType = m_TotalType
    End Function

    Public Sub SetCategoryTotalType(tType As String)
        m_TotalType = tType
    End Sub

    Public Function GetCategorySortType()
        GetCategorySortType = m_SortType
    End Function

    Public Sub SetCategorySortType(sType As String)
        m_SortType = sType
    End Sub

    Public Function GetCategoryRange() As String
        GetCategoryRange = m_CategoryRange.Address
    End Function

    Public Sub SetCategoryRange(r As String)
        Set m_CategoryRange = Range(r)
    End Sub

    '----------------------------------------------
    ' Public Functions
    '----------------------------------------------
    Public Sub ApplyFormatting()
        With Range(GetCategoryNameAddress())
            .Merge
            .BorderAround
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 14
            .Font.name = "Arial Black"
            .Font.Bold = 1
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(51, 153, 102)
            .RowHeight = 51
            .Value = m_Name
            .WrapText = True
        End With
        
        With Range(GetCategoryInternalSpacerAddress())
            .RowHeight = 21
        End With
        
        With Range(GetCategoryStoreAddress())
            .RowHeight = 23
        
            With .Columns(1)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Font.Size = 14
                .Font.name = "Arial Black"
            End With
            With .Columns(2)
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .Font.Size = 14
                .Font.name = "Arial Black"
                .Font.Bold = 1
                .NumberFormat = m_ValueType
            End With
        End With
        
        With Range(GetCategoryAreaAddress())
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).LineStyle = xlDouble
            .Font.Size = 14
            .Font.name = "Arial Black"
            With .Columns(1)
                .Value = "Area"
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            With .Columns(2)
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
            End With
        End With
    End Sub

    Public Sub ApplyConditionalFormatting(obj As CStoreMgr)

        If m_Condition <> "N/A" And m_LookupValue <> "N/A" Then
            Dim formula As String
            Dim condition As String
            Dim index As Integer
            Dim i As Integer
            
            condition = m_Condition
            index = InStr(condition, "()")
            
            For i = 1 To Util.GetTotalStores() + 1
                If index > 0 Then
                    formula = Replace(condition, "()", "(" & Range(GetCategoryStoreAddress(i)).Columns(2).Address & ")")
                    formula = "IF(" & formula & ", TRUE, FALSE)"
                Else
                    formula = Range(GetCategoryStoreAddress(i)).Columns(2).Address & condition
                    formula = "IF(" & Range(GetCategoryStoreAddress(i)).Columns(2).Value & condition & ", TRUE, FALSE)"
                End If
                
    '            With Range(GetCategoryStoreAddress(i))
    '                .FormatConditions.Add Type:=xlExpression, Formula1:=formula
    '                With .FormatConditions(.FormatConditions.Count)
    '                    .Font.Color = RGB(255, 0, 0)
    '
    '                    If i <> Util.GetTotalStores() + 1 Then
    '                        ' Check if font was formatted
    '                    End If
    '                End With
    '            End With
                
                Debug.Print ("INDEX - " & i)
                Debug.Print ("CONDITION - " & condition)
                Debug.Print ("FORMULA - " & formula)
                
                If Application.Evaluate(formula) Then
                   Range(GetCategoryStoreAddress(i)).Font.Color = RGB(255, 0, 0)
                        
                    If i <> Util.GetTotalStores() + 1 Then
                        Call obj.SetStoreMisses(i, obj.GetStoreMisses(i) + 1)
                    End If
                End If
            Next i
        End If
    End Sub

    Public Sub SortData()
        If m_SortType <> "N/A" And m_LookupValue <> "N/A" Then
            If m_SortType = "HIGH-LOW" Then
                ActiveSheet.Sort.SortFields.Clear
                ActiveSheet.Sort.SortFields.Add _
                    Key:=Range(GetCategoryStoreAddress()).Columns(2), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending, _
                    DataOption:=xlSortNormal
                    
                With ActiveSheet.Sort
                    .SetRange Range(GetCategoryStoreAddress())
                    .Header = xlGuess
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
            ElseIf m_SortType = "LOW-HIGH" Then
                ActiveSheet.Sort.SortFields.Clear
                ActiveSheet.Sort.SortFields.Add _
                    Key:=Range(GetCategoryStoreAddress()).Columns(2), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                    
                With ActiveSheet.Sort
                    .SetRange Range(GetCategoryStoreAddress())
                    .Header = xlGuess
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
            ElseIf m_SortType = "CLOSEST-TO-0" Then
                
                Dim i As Integer
                Dim returnABS(8) As String
                
                ' Store original value, get absolute value, and sort on that
                For i = 0 To Util.GetTotalStores() - 1
                    If Range(GetCategoryStoreAddress(i + 1)).Columns(2).Value < 0 Then
                        Range(GetCategoryStoreAddress(i + 1)).Columns(2).Value = _
                            Range(GetCategoryStoreAddress(i + 1)).Columns(2).Value * -1
                        returnABS(i) = Range(GetCategoryStoreAddress(i + 1)).Columns(1).Value
                    End If
                Next i
                
                ActiveSheet.Sort.SortFields.Clear
                ActiveSheet.Sort.SortFields.Add _
                    Key:=Range(GetCategoryStoreAddress()).Columns(2), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                    
                With ActiveSheet.Sort
                    .SetRange Range(GetCategoryStoreAddress())
                    .Header = xlGuess
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                
                ' Reapply any negative values
                Dim s As Integer
                
                For i = 0 To Util.GetTotalStores() - 1
                    For s = 0 To 8
                        If Range(GetCategoryStoreAddress(i + 1)).Columns(1) = returnABS(s) Then
                            '- Found a match
                            Range(GetCategoryStoreAddress(i + 1)).Columns(2).Value = _
                                Range(GetCategoryStoreAddress(i + 1)).Columns(2).Value * -1
                        End If
                    Next s
                Next i
            Else
                MsgBox ("Error sorting data for category.")
            End If
        End If
    End Sub

    Public Sub SetTotalValue()
        With Range(GetCategoryTotalAddress())
            .Value = "=" & m_TotalType & "(" & GetCategoryTotalRangeAddress() & ")"
        End With
    End Sub

    Public Sub SetStoreNames(obj As CStoreMgr)

        Dim s As Integer
        For s = 1 To Util.GetTotalStores()
            With Range(GetCategoryStoreAddress(s))
                .Columns(1).Value = obj.GetStoreName(s)
                
                If WorksheetExists("Data") Then
                    If m_LookupValue <> "N/A" Then
                    
                        ' Compound operation
                        If InStr(m_LookupValue, "[") Then
                            Dim fLookup As String
                            Dim sLookup As String
                            Dim op As String
                            
                            fLookup = Mid(m_LookupValue, 1, InStr(m_LookupValue, "[") - 2)
                            sLookup = Mid(m_LookupValue, InStr(m_LookupValue, "]") + 2)
                            op = Mid(m_LookupValue, InStr(m_LookupValue, "[") + 1, 1)
                        
                            .Columns(2).Value = _
                                "=IFERROR(HLOOKUP(" _
                                    & Chr(34) & fLookup & Chr(34) & "," _
                                    & "Data!D1:BI11,MATCH(" _
                                    & Chr(34) & obj.GetStoreNumber(s) & " " & obj.GetStoreName(s) & " - " _
                                    & obj.GetStoreType(s) & Chr(34) & "," _
                                    & "Data!A1:A10,0),FALSE),"""")" & " " & op & " " _
                                    & "IFERROR(HLOOKUP(" _
                                    & Chr(34) & sLookup & Chr(34) & "," _
                                    & "Data!D1:BI11,MATCH(" _
                                    & Chr(34) & obj.GetStoreNumber(s) & " " & obj.GetStoreName(s) & " - " _
                                    & obj.GetStoreType(s) & Chr(34) & "," _
                                    & "Data!A1:A10,0),FALSE),"""")"
                                    
                        ' Set the misses
                        ElseIf InStr(m_LookupValue, "MISSES") Then
                            .Columns(2).Value = obj.GetStoreMisses(s)
                            
                        ' Default
                        Else
                            .Columns(2).Value = _
                                "=IFERROR(HLOOKUP(" _
                                    & Chr(34) & m_LookupValue & Chr(34) & "," _
                                    & "Data!D1:BI11,MATCH(" _
                                    & Chr(34) & obj.GetStoreNumber(s) & " " & obj.GetStoreName(s) & " - " _
                                    & obj.GetStoreType(s) & Chr(34) & "," _
                                    & "Data!A1:A10,0),FALSE),"""")"
                        End If
                    End If
                End If
            End With
        Next s
    End Sub
 
    '----------------------------------------------
    ' Private Functions
    '----------------------------------------------
    Private Function GetCategoryNameAddress() As String
        GetCategoryNameAddress = Range(m_CategoryRange.Cells(1, 1), m_CategoryRange.Cells(1, 2)).Address
    End Function

    Private Function GetCategoryInternalSpacerAddress() As String
        GetCategoryInternalSpacerAddress = Range(m_CategoryRange.Cells(2, 1), m_CategoryRange.Cells(2, 2)).Address
    End Function

    Private Function GetCategoryStoreAddress(Optional index As Variant) As String
        If IsMissing(index) Then
            GetCategoryStoreAddress = Range(m_CategoryRange.Cells(3, 1), m_CategoryRange.Cells(3 + Util.GetTotalStores() - 1, 2)).Address
        Else
            GetCategoryStoreAddress = Range(m_CategoryRange.Cells(3 + (index - 1), 1), m_CategoryRange.Cells(3 + (index - 1), 2)).Address
        End If
    End Function

    Private Function GetCategoryAreaAddress() As String
        GetCategoryAreaAddress = Range(m_CategoryRange.Cells(3 + Util.GetTotalStores(), 1), m_CategoryRange.Cells(3 + Util.GetTotalStores(), 2)).Address
    End Function

    Private Function GetCategoryTotalAddress() As String
        GetCategoryTotalAddress = Range(m_CategoryRange.Cells(3 + Util.GetTotalStores(), 2), m_CategoryRange.Cells(3 + Util.GetTotalStores(), 2)).Address
    End Function

    Private Function GetCategoryTotalRangeAddress() As String
        GetCategoryTotalRangeAddress = Range(m_CategoryRange.Cells(3, 2), m_CategoryRange.Cells(3 + Util.GetTotalStores() - 1, 2)).Address
    End Function
End Sub