
Sub AddCategoryUF
    Public Sub AddCategoryUF_Initialize()
        Debug.Print ("Initialized.")

        CategoryNameTB.Value = ""
        CategoryLookupValueTB.Value = ""
        CategoryFormatConditionTB.Value = ""
        
        With CategoryValueTypeCB
            .AddItem "GENERAL"
            .AddItem "CURRENCY"
            .AddItem "PERCENT"
            .AddItem "NUMBER"
        End With
        
        With CategoryTotalTypeCB
            .AddItem "SUM"
            .AddItem "AVERAGE"
        End With
        
        With CategorySortTypeCB
            .AddItem "LOW-HIGH"
            .AddItem "HIGH-LOW"
        End With
    End Sub

    Private Sub AddCategoryButton_Click()
        If CategoryNameTB.Value = "" Then
            MsgBox ("Must enter a Name for the category.")
            Exit Sub
        End If
        
        If CategoryValueTypeCB.Value = "" Then
            MsgBox ("Must enter a Value Type for the category.")
            Exit Sub
        End If
        
        If CategoryTotalTypeCB.Value = "" Then
            MsgBox ("Must enter a Total Type for the category.")
            Exit Sub
        End If
        
        If Util.GetTotalCategories() = Util.MaxCategories Then
            MsgBox ("Error: The maximum number of categories (" & Util.MaxCategories & ") has already been created.")
            Exit Sub
        End If
    End Sub
End Sub