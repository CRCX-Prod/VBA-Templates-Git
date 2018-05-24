Attribute VB_Name = "Table_Admin"
Option Explicit

Public isAdmin As Boolean
Public sLogin As String, sOperator As String, sRegion As String

Sub SetAdmin()
    isAdmin = True
    ShowAllSheets
End Sub

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
       
    HideAllSheets
    isAdmin = False
    
    Do Until rs.EOF
      If inputLogin = rs.Fields("Login") Then
        If inputPassword = rs.Fields("Password") Then
          'MsgBox "Good Password"
          AdminSheets rs
          SetAdminSession rs
          SetAdminFilters rs
          LoginForm.Hide
          sLogin = inputLogin
          Application.CalculateFull
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

Sub SetAdminSession(rs As Recordset)
    
    If rs.Fields("Admin").Value = "x" Then
        SetAdmin
    End If
    
End Sub


Sub AdminSheets(rs As Recordset)
    Dim iFields As Integer
    Application.ScreenUpdating = False
    
    For iFields = 1 To rs.Fields.Count - 1
        
        ShowSheet rs.Fields(iFields).Name, rs.Fields(iFields).Value
        
    Next iFields
    
    Application.ScreenUpdating = True
End Sub

Sub SetAdminFilters(rs As Recordset)

    sOperator = rs.Fields("Operator").Value
    sRegion = rs.Fields("Region").Value

End Sub

'Functions to get Admin config in Excel fields
Function GetLogin() As String
    Application.Volatile True
    GetLogin = sLogin
End Function

'Region Filter
Function GetRegion() As String
    GetRegion = sRegion
End Function

'Operator Filter
Function GetOperator() As String
    Application.Volatile True
    GetOperator = sOperator
End Function
