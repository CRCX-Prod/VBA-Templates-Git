Attribute VB_Name = "Append_Data"
Option Explicit

'_________________________________________________________
'AppendData
'Table Name is B1
'Take fields values from line 2
'take appened values from line 3
'_________________________________________________________


Sub TestMandatory()

    Dim i As Integer
    
    i = 1
    
    While Cells(12, i) <> ""
        
        If Cells(14, i) = "Mandatory" Then
            
            If Cells(13, i) = "" Then
                MsgBox "Miss mandatory information"
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

  While Cells(12, i) <> ""
      If i = 1 Then   'field 1
          SqlFields = Cells(12, i) & "= '" & Cells(13, i) & "'"
      Else            'other fields
          SqlFields = SqlFields & "," & Cells(12, i) & "= '" & Cells(13, i) & "'"
      End If
      i = i + 1
  Wend

  sql = "INSERT INTO " & Cells(11, 2) & " SET " & SqlFields
  'MsgBox Sql
  oConn.Execute sql

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing
End Sub