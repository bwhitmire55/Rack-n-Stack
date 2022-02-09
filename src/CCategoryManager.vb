Sub
    CCategoryMgr
    '----------------------------------------------
    'Member Variables
    '----------------------------------------------
    Private m_Worksheet As Worksheet
    Private m_CategoryCount As Integer
    Private m_CategoryCollection As Collection
    Private m_CategoryTable As Range

    '----------------------------------------------
    ' Constructor / Destructor
    '----------------------------------------------
    Private Sub Class_Initialize()
        
        Set m_CategoryCollection = New Collection
        Set m_Worksheet = Worksheets(wsSetupStr)
        
        Dim ref As Range
        Dim i As Integer
        Dim cat As CCategory
        
        Set ref = m_Worksheet.Cells.Find("Total Categories:")
        
        ' Check if "Total Categories" is set
        If ref Is Nothing Then
            MsgBox ("Failed to intialize Category Manager Class - Could not find 'Total Categories'.")
            Exit Sub
        End If
        
        ' Create extra reference to account for Offset function call
        Dim tRef As Range
        Set tRef = ref
        
        m_CategoryCount = tRef.Offset(0, 1).Value()
        
        ' Check if 'Category Count" is valid
        If m_CategoryCount <= 0 Then
            MsgBox ("Failed to initialize Category Manager Class - 'Total Categories' value is invalid.")
            Exit Sub
        End If
        
        tRef = ref
        Set m_CategoryTable = Range(tRef.Offset(2, -4), ref.Offset(2 + m_CategoryCount, 1))
        
        ' Create a category instance for each found, add to the collection
        For i = 1 To m_CategoryCount
            
            Set cat = New CCategory
            
            cat.SetCategoryName (Application.IfError(Application.HLookup("Text", m_CategoryTable, i + 1, False), "N/A"))
            cat.SetCategoryLookupValue (Application.IfError(Application.HLookup("Lookup Value", m_CategoryTable, i + 1, False), "N/A"))
            cat.SetCategoryValueType (Application.IfError(Application.HLookup("Value Type", m_CategoryTable, i + 1, False), "N/A"))
            cat.SetCategoryCondition (Application.IfError(Application.HLookup("Condition", m_CategoryTable, i + 1, False), "N/A"))
            cat.SetCategoryTotalType (Application.IfError(Application.HLookup("Total Type", m_CategoryTable, i + 1, False), "N/A"))
            cat.SetCategorySortType (Application.IfError(Application.HLookup("Sort Type", m_CategoryTable, i + 1, False), "N/A"))
            
            If cat.GetCategoryName = 0 Or cat.GetCategoryTotalType = 0 Then
                MsgBox ("Error initializing data for category '" & i & "'.")
            Else
                m_CategoryCollection.Add cat
            End If
        Next i
        
    End Sub
     
    Private Sub Class_Terminate()
     
    End Sub
     
    '----------------------------------------------
    ' Public Functions
    '----------------------------------------------
    Public Function GetTotalCategories()
        GetTotalCategories = m_CategoryCount
    End Function
     
    Public Function GetCategoryName(index As Integer) As String
        GetCategoryName = m_CategoryCollection.Item(index).GetCategoryName
    End Function
     
    Public Function GetCategoryLookupValue(index As Integer) As String
        GetCategoryLookupValue = m_CategoryCollection.Item(index).GetCategoryLookupValue
    End Function
     
    Public Function GetCategoryValueType(index As Integer) As String
        GetCategoryValueType = m_CategoryCollection.Item(index).GetCategoryValueType
    End Function
     
    Public Function GetCategoryTotalType(index As Integer) As String
        GetCategoryTotalType = m_CategoryCollection.Item(index).GetCategoryTotalType
    End Function
     
    Public Sub SetCategoryRange(index As Integer, r As String)
        m_CategoryCollection.Item(index).SetCategoryRange (r)
    End Sub
     
    Public Function GetCategoryRange(index As Integer) As String
        GetCategoryRange = m_CategoryCollection.Item(index).GetCategoryRange()
    End Function
     
    Public Sub ApplyFormatting(index As Integer)
        m_CategoryCollection.Item(index).ApplyFormatting
    End Sub
     
    Public Sub ApplyConditionalFormatting(index As Integer, obj As CStogr)
        m_CategoryCollection.Item(index).ApplyConditionalFormatting obj
    End Sub
     
    Public Sub SortData(index As Integer)
        m_CategoryCollection.Item(index).SortData
    End Sub
     
    Public Sub SetTotalValue(index As Integer)
        m_CategoryCollection.Item(index).SetTotalValue
    End Sub
     
    Public Sub SetStoreNames(index As Integer, obj As CStogr)
        m_CategoryCollection.Item(index).SetStoreNames obj
    End Sub
     
    Public Sub PrintCategories()
        Debug.Print ("Total Categories: " & GetTotalCategories())
        
        Dim iCat As CCategory
        For Each iCat In m_CategoryCollection
            Debug.Print ("Category " & iCat.GetCategoryName & "|" & iCat.GetCategoryRange)
        Next iCat
    End Sub
End Sub