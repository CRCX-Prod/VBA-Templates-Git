Attribute VB_Name = "MySQL_Connection"
Option Explicit

'__________________________________________________________

'Configuration et Connection a la base de donnees
'__________________________________________________________


Public oConn As ADODB.Connection

Sub ConnectProductionServer()
    'Enter here Connection informations'
    
    ConnectDB "192.168.1.153", "01_unlmm", "unlmm", "unlmm"
End Sub

Sub ConnectDB(server_name As String, database_name As String, user_id As String, password As String)
  Set oConn = New ADODB.Connection
  Dim str As String
  str = "DRIVER={MySQL ODBC 5.3 ANSI Driver};" & _
        ";SERVER=" & server_name & _
        ";PORT=3306" & _
        ";DATABASE=" & database_name & _
        ";UID=" & user_id & _
        ";PWD=" & password & _
        ";Option=3"

  oConn.Open str
End Sub

'_________________________________________________________
'AppendData
'input Sheet number
'Table Name is B1
'Take fields values from line 2
'take appened values from line 3
'_________________________________________________________

Sub TestMandatory()

    Dim i As Integer
    
    i = 1
    
    While Cells(2, i) <> ""
        
        If Cells(4, i) = "Mandatory" Then
            
            If Cells(3, i) = "" Then
                MsgBox "Miss mandatory information or Colo already exists"
                Exit Sub
            End If
        End If
        i = i + 1
    Wend
        
        'MsgBox "Append Data"
        AppendData
        MsgBox "Data added in the Database"
    
End Sub

Sub AppendData()

  Dim rs As ADODB.Recordset
  Dim field(), sql, SqlFields As String
  Dim i As Integer

  Set rs = New ADODB.Recordset
  i = 1

  ConnectProductionServer

  While Cells(2, i) <> ""
      If i = 1 Then   'field 1
          SqlFields = Cells(2, i) & "= '" & Cells(3, i) & "'"
      Else            'other fields
          SqlFields = SqlFields & "," & Cells(2, i) & "= '" & Cells(3, i) & "'"
      End If
      i = i + 1
  Wend

  sql = "INSERT INTO " & Cells(1, 2) & " SET " & SqlFields
  'MsgBox Sql
  oConn.Execute sql

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing
End Sub
