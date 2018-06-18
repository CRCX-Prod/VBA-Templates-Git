Attribute VB_Name = "DRAFTSqlUpdate"
Option Explicit

Dim iUpdate As Integer

Sub TestUpdate()
    Dim sSql As String, sTableName As String, sSqlFields As String, sSqlFilters As String
    Dim rs As ADODB.Recordset

    Application.ScreenUpdating = False
    iUpdate = 0
    '___________________________
    'Initialize RecordSet
        Set rs = New ADODB.Recordset
        ConnectProductionServer
        'Get Table Name
        sTableName = Cells(FindLine("Table Name", 1), 2)
        'Get Import fields
        sSqlFields = SqlImportFields(ArrayLine(FindLine("Import Data", 1) + 1))
        'Get Filters statement
        sSqlFilters = SqlImportFilters(ArrayLine(FindLine("Filters", 1) + 1), ArrayLine(FindLine("Filters", 1) + 2))
        'Sql statement'
        sSql = SqlSelectQuery(sTableName, sSqlFields, sSqlFilters)
        rs.Open sSql, oConn
    '____________________________

    'Compare RecordSet with Value
    LookupRs rs

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing
  MsgBox iUpdate & " Record(s) updated in database"
  Application.ScreenUpdating = True


End Sub

Sub LookupRs(rs As Recordset)
    'parcours le Recordset pour trouver le bon enregistrement
    Dim arrayField() As String, arrayUpdate() As String, maxArray As Integer
    Dim iUpdate As Integer

    arrayField = ArrayLine(FindLine("Update Data", 1) + 1)
    maxArray = SizeArray(arrayField())

    rs.MoveFirst
    Do Until rs.EOF
        Cells(FindLine("Update Data", 1) + 3, 1) = rs.Fields(0).Value
        arrayUpdate = ArrayLineDim(FindLine("Update Data", 1) + 3, maxArray)
        CompareRs rs, arrayField, arrayUpdate

        rs.MoveNext
    Loop

End Sub

Sub CompareRs(rs As Recordset, arrayField() As String, arrayUpdate() As String)
    'Parcours les champs de l'enregistrement pour voir les modifications
    Dim maxArray As Integer, iArray As Integer

    maxArray = SizeArray(arrayField())

    For iArray = 2 To maxArray - 1
    
        Select Case rs.Fields(arrayField(iArray)).Value
            Case Is <> arrayUpdate(iArray)
                RunUpdateSQL arrayField(), arrayUpdate()
                iUpdate = iUpdate + 1
                Exit Sub
            'Test other case
            Case Is = Null And arrayUpdate(iArray) = ""
                RunUpdateSQL arrayField(), arrayUpdate()
                iUpdate = iUpdate + 1
                Exit Sub
        End Select
        
''''        If arrayUpdate(iArray) <> rs.Fields(arrayField(iArray)).Value Then
''''            RunUpdateSQL arrayField(), arrayUpdate()
''''            iUpdate = iUpdate + 1
''''            Exit Sub
''''        End If

    Next iArray

End Sub

Sub RunUpdateSQL(arrayField() As String, arrayUpdate() As String)
    Dim sSql As String, sTableName As String, sSqlFields As String

    Dim testLog As String

    sTableName = Cells(FindLine("Table Name", 1), 2)
    sSqlFields = SqlSet(arrayField(), arrayUpdate())

    sSql = "UPDATE " & sTableName & _
            " SET " & sSqlFields & _
            " WHERE " & arrayField(1) & " = " & arrayUpdate(1)
    'MsgBox sSql
    ExecuteUpdateLog sSql, "07preva_log"

End Sub

 Function SqlSet(arrayField() As String, arrayUpdate() As String) As String

    Dim maxArray As Integer, iArray As Integer
    Dim sSqlSet As String, sLog As String

    sSqlSet = ""

    maxArray = SizeArray(arrayField())

    For iArray = 2 To maxArray - 1
        If arrayUpdate(iArray) <> "" Then
            sSqlSet = sSqlSet & arrayField(iArray) & " = '" & arrayUpdate(iArray) & "', "
        Else
           sSqlSet = sSqlSet & arrayField(iArray) & " = NULL, "
        End If
    Next iArray

    sSqlSet = Left(sSqlSet, Len(sSqlSet) - 2)
    SqlSet = sSqlSet

 End Function

 Sub ExecuteUpdateLog(sSql As String, logTable As String)

    Dim sSqlLog As String

    oConn.Execute sSql
    'Add record in LogHistory

    sSqlLog = "INSERT INTO " & logTable & _
              " SET Login = '" & sLogin & _
                    "', sql_query = '" & sSql & _
                    "', timestamp = " & Now()
    'oConn.Execute sSqlLog
    MsgBox sSqlLog

End Sub
