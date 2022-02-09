Sub CStore
    '----------------------------------------------
    ' Member Variables
    '----------------------------------------------
    Private m_Number As Integer
    Private m_Name As String
    Private m_Type As String
    Private m_Misses As Integer

    '----------------------------------------------
    ' Getters / Setters
    '----------------------------------------------
    Public Function GetStoreNumber() As Integer
        GetStoreNumber = m_Number
    End Function

    Public Sub SetStoreNumber(number As Integer)
        m_Number = number
    End Sub

    Public Function GetStoreName()
        GetStoreName = m_Name
    End Function

    Public Sub SetStoreName(name As String)
        m_Name = name
    End Sub

    Public Function GetStoreType()
        GetStoreType = m_Type
    End Function

    Public Sub SetStoreType(uType As String)
        m_Type = uType
    End Sub

    Public Function GetStoreMisses() As Integer
        GetStoreMisses = m_Misses
    End Function

    Public Sub SetStoreMisses(misses As Integer)
        m_Misses = misses
    End Sub
End Sub