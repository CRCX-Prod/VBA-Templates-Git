Attribute VB_Name = "DRAFTSqlUpdate"
Option Explicit

Sub TestUpdate()
    
    Dim sSql As String, sTableName As String, sSqlFields As String, sSqlFilters As String
    
    Dim rs As ADODB.Recordset

    Application.ScreenUpdating = False
    '___________________________
    'Initialize RecordSet
        Set rs = New ADODB.Recordset
        ConnectProductionServer
        'Get Table Name
        sTableName = Cells(FindLine("Table Name", 1), 2)
        'Get Import fields
        sSqlFields = SqlImportFields(ArrayLine(FindLine("Import Data", 1) + 1))
        'Sql statement'
        sSql = SqlSelectQuery(sTableName, sSqlFields & ", LogHistory", "")
        rs.Open sSql, oConn
    '____________________________
    
    'Compare RecordSet with Value
    RunExcelID rs, 16, 1
   
  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing
  Application.ScreenUpdating = True
        
    
End Sub


Sub RunExcelID(rs As Recordset, firstLine As Integer, firstColumn As Integer)
    'Parcours la table mise a jour
    Dim tableLine As Integer
    Dim arrayField() As String, arrayUpdate() As String
    tableLine = firstLine
    

    arrayField = ArrayLine(FindLine("Update Data", 1) + 1)
    
    'Pracours le tableau Excel
    While Cells(tableLine, firstColumn) <> ""
        Cells(FindLine("Update Data", 1) + 3, 1) = Cells(tableLine, firstColumn)
        arrayUpdate = ArrayLine(FindLine("Update Data", 1) + 3)
        'Teste l'enregistrement
        LookupRs rs, arrayUpdate()
        CompareRs rs, arrayField(), arrayUpdate()
        tableLine = tableLine + 1
    Wend
    
   
''''''' Cells(5, 1) = Cells(firstLine + 1, firstColumn)
'''''''            arrayUpdate = ArrayLine(FindLine("Update Data", 1) + 3)
'''''''    MsgBox arrayUpdate(2)
'''''''    LookupRs rs, arrayUpdate()
'''''''    CompareRs rs, arrayField(), arrayUpdate()
    
End Sub

Sub LookupRs(rs As Recordset, arrayUpdate() As String)
    'parcours le Recordset pour trouver le bon enregistrement
    Dim iId As Integer
    rs.MoveFirst
       
    Do Until rs.EOF
        If arrayUpdate(1) = rs.Fields(0).Value Then
'''''            MsgBox rs.Fields(0).Value
            Exit Sub
        End If
        rs.MoveNext
    Loop

End Sub

Sub CompareRs(rs As Recordset, arrayField() As String, arrayUpdate() As String)
    'Parcours les champs de l'enregistrement pour voir les modifications
    Dim maxArray As Integer, iArray As Integer
        
    maxArray = SizeArray(arrayUpdate())
    
    For iArray = 2 To maxArray - 1
        If arrayUpdate(iArray) <> rs.Fields(arrayField(iArray)).Value Then
            RunUpdateSQL rs, arrayField(), arrayUpdate()
            Exit Sub
        End If

    Next iArray
        
End Sub

'          Sql = "UPDATE " & tableName & " SET " & SqlFields & " WHERE " & Cells(3, 1) & " = " & Cells(5, 1)
'          'MsgBox Sql
'          oConn.Execute Sql

Sub RunUpdateSQL(rs As Recordset, arrayField() As String, arrayUpdate() As String)
    Dim sSql As String, sTableName As String, sSqlFields As String
    
    Dim testLog As String
    
    sTableName = Cells(FindLine("Table Name", 1), 2)
    sSqlFields = SqlSet(rs, arrayField(), arrayUpdate(), True)

    sSql = "UPDATE " & sTableName & _
            " SET " & sSqlFields & _
            " WHERE " & arrayField(1) & " = " & arrayUpdate(1)
    MsgBox sSql
    oConn.Execute sSql
    
End Sub
 
 Function SqlSet(rs As Recordset, arrayField() As String, arrayUpdate() As String, isLogHistory As Boolean) As String
    
    Dim maxArray As Integer, iArray As Integer
    Dim sSqlSet As String, sLog As String
       
    sSqlSet = ""
    sLog = rs.Fields("LogHistory").Value
    sLog = sLog & " " & sLogin & ", "
    maxArray = SizeArray(arrayUpdate())
    
    For iArray = 2 To maxArray - 1

        sSqlSet = sSqlSet & arrayField(iArray) & " = '" & arrayUpdate(iArray) & "', "
        sLog = sLog & arrayField(iArray) & " = " & arrayUpdate(iArray) & ", "
    Next iArray
    
    sLog = sLog & Now() & ";"
    MsgBox sLog
    sSqlSet = sSqlSet & " LogHistory = """ & sLog & """"
    SqlSet = sSqlSet
    
 End Function


Function UpdateLog(sSqlSet As String) As String
    
    UpdateLog = sLogin & ", " & sSqlSet & ", " & Now() & ";"

End Function
