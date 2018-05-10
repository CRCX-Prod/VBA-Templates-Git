Attribute VB_Name = "Table_Admin"
Option Explicit

Sub HideAllSheets()
  'VeryHide all sheets appart Main'
  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Main" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next
End Sub

Sub ShowAllSheets()
  'VeryHide all sheets'
  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next
End Sub


Sub ShowSheet(sheetName As String, adminValue As Variant)
    'show sheet according to the sheet name, if adminValue = "x"
    Dim ws As Worksheet

    If adminValue = "x" Then
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = sheetName Then
                ws.Visible = xlSheetVisible
            End If
        Next
    End If

End Sub

Sub TestPassword(inputLogin As String, inputPassword As String)

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
          'MsgBox "Good Password"
          AdminSheets rs
          LoginForm.Hide
        Else
          MsgBox "Wrong Password"
        End If
      Else

      End If
      rs.MoveNext
    Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

End Sub

Sub AdminSheets(rs As Recordset)
    Dim iFields As Integer

    For iFields = 1 To rs.Fields.Count - 1
        ShowSheet rs.Fields(iFields).Name, rs.Fields(iFields).Value
        'MsgBox rs.Fields(iFields).Name
    Next iFields

''    MsgBox rs.Fields(iFields).Name
''''MsgBox rs.Fields(1).Name
''''MsgBox rs.Fields(1).Value

End Sub
