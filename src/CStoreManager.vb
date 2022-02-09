Sub CStoreMgr
    '----------------------------------------------
    'Member Variables
    '----------------------------------------------
    Private m_Worksheet As Worksheet
    Private m_StoreCount As Integer
    Private m_StoresCollection As Collection
    Private m_StoreTable As Range

    '----------------------------------------------
    ' Constructor / Destructor
    '----------------------------------------------
    Private Sub Class_Initialize()

        Set m_StoresCollection = New Collection
        Set m_Worksheet = Worksheets(wsSetupStr)

        Dim ref As Range
        Dim i As Integer
        Dim store As CStore

        Set ref = m_Worksheet.Cells.Find("Total Stores:")

        ' Check if "Total Stores" is set
        If ref Is Nothing Then
            MsgBox ("Failed to initialize Store Manager Class - Could not find 'Total Stores' in '" & wsSetupStr & "'")
            Exit Sub
        End If

        ' Create extra reference to account for Offset function call
        Dim tRef As Range
        Set tRef = ref

        m_StoreCount = tRef.Offset(0, 1).Value()

        ' Check if "Store Count" is valid
        If m_StoreCount <= 0 Then
            MsgBox ("Failed to intialize Store Manager Class - 'Total Stores' value is invalid.")
            Exit Sub
        End If

        tRef = ref
        Set m_StoreTable = Range(tRef.Offset(2, -1), ref.Offset(2 + m_StoreCount, 1))

        ' Create a store instance for each found, add to the collection
        For i = 1 To m_StoreCount

            Set store = New CStore

            store.SetStoreNumber (Application.IfError(Application.HLookup("Store Number", m_StoreTable, i + 1, False), 0))
            store.SetStoreName (Application.IfError(Application.HLookup("Name", m_StoreTable, i + 1, False), 0))
            store.SetStoreType (Application.IfError(Application.HLookup("Type", m_StoreTable, i + 1, False), 0))
            store.SetStoreMisses (0)

            If store.GetStoreNumber = 0 Or store.GetStoreName = 0 Or store.GetStoreType = 0 Then
                MsgBox ("Error initializing data for store '" & i & "'.")
            Else
                m_StoresCollection.Add store
            End If
        Next i

    End Sub

    Private Sub Class_Terminate()

    End Sub

    '----------------------------------------------
    ' Public Functions
    '----------------------------------------------
    Public Function GetTotalStores()
        GetTotalStores = m_StoreCount
    End Function

    Public Function GetStoreName(index As Integer) As String
        GetStoreName = m_StoresCollection.Item(index).GetStoreName
    End Function

    Public Function GetStoreNumber(index As Integer) As Integer
        GetStoreNumber = m_StoresCollection.Item(index).GetStoreNumber
    End Function

    Public Function GetStoreType(index As Integer) As String
        GetStoreType = m_StoresCollection.Item(index).GetStoreType
    End Function

    Public Function GetStoreMisses(index As Integer) As Integer
        GetStoreMisses = m_StoresCollection.Item(index).GetStoreMisses
    End Function

    Public Sub SetStoreMisses(index As Integer, misses As Integer)
        m_StoresCollection.Item(index).SetStoreMisses (misses)
    End Sub

    Public Sub PrintStores()
        Debug.Print ("Total Stores: " & GetTotalStores())

        Dim iStore As CStore
        For Each iStore In m_StoresCollection
            Debug.Print ("Store " & iStore.GetStoreNumber & iStore.GetStoreName & iStore.GetStoreType)
        Next iStore
    End Sub
End Sub