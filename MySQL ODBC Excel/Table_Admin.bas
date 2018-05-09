Attribute VB_Name = "Admin"
Option Explicit

Sub HideAllSheets()
  'VeryHide all sheets'

End Sub

Sub ShowSheet(sheetName As String)
    'show sheet according to the sheet name'

End Sub

Sub RunTestPassword(inputLogin As String, inputPassword As String)

    Dim rs As ADODB.Recordset
    Dim sSql As String, adminTable As String
  
    Set rs = New ADODB.Recordset

    adminTable = "06preva_admin"
    sSql = SqlSelectQuery(adminTable, "*", "")
    
    MsgBox sSql
    
    ConnectProductionServer
    rs.Open sSql, oConn
    
    Do Until rs.EOF


    Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

End Sub

Sub TestLogin(inputLogin As String, data)

End Sub

Sub Test()
    
  Dim rs As ADODB.Recordset
  Dim line As Integer, column As Integer

  Set rs = New ADODB.Recordset

  ConnectProductionServer
  rs.Open sqlQuery, oConn

    line = 0
      Do Until rs.EOF

          For column = 0 To rs.Fields.Count - 1
              Cells(line + firstLine, column + firstColumn) = rs.Fields(column).Value

          Next
          line = line + 1
          rs.MoveNext
      Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

    
End Sub
