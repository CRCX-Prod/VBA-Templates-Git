Attribute VB_Name = "Table_Admin"
Option Explicit

Sub HideAllSheets()
  'VeryHide all sheets'
  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVeryHidden
    Next
End Sub

Sub ShowAllSheets()
  'VeryHide all sheets'
  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next
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

    ConnectProductionServer
    rs.Open sSql, oConn

    Do Until rs.EOF
      If inputLogin = rs.Fields("Login") Then
        If inputPassword = rs.Fields("Password") Then
          MsgBox "Good Password"
          LoginForm.Hide
        Else
          MsgBox "Wrong Password"
        End If ' true
      Else
        Exit Sub
      End If
      rs.MoveNext
    Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

End Sub

Sub subTest()


RunTestPassword "Charles", "Charles01"
End Sub

Sub TestLogin(inputLogin As String, data)

    RunTestPassword "Charles", "P@ssw0rd"

End Sub
