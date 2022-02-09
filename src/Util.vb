Sub Util
    '----------------------------------------------
    ' Constants
    '----------------------------------------------
    Public Const wsSetupStr As String = "Setup"
    Public Const wsDataStr As String = "Data"
    Public Const wsWRStr As String = "Weekly Rankings"

    Public Const MaxCategories As Integer = 16
    Public Const MaxStores As Integer = 10

    Private g_TotalStores As Integer
    Private g_TotalCategories As Integer

    Public Function GetTotalStores() As Integer
        GetTotalStores = g_TotalStores
    End Function

    Public Sub SetTotalStores(Value As Integer)
        g_TotalStores = Value
    End Sub

    Public Sub SetTotalCategories(Value As Integer)
        g_TotalCategories = Value
    End Sub

    Public Function GetTotalCategories() As Integer
        GetTotalCategories = g_TotalCategories
    End Function

    Function WorksheetExists(sName As String) As Boolean '- STACK OVERFLOW
        WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
    End Function
End Sub