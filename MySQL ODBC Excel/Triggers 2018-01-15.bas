Attribute VB_Name = "Triggers"
Option Explicit

'Triggers are the codes directly runned by Buttons'

Sub UpdatePrevalidation()
  'UpdateData and Table location'
  Dim isTimeOut As Boolean
  
  isTimeOut = TimeOut(10)
  
  If isTimeOut = True Then
    UpdateData 16, 1
    MsgBox "Data recorded in Database."
  Else
    MsgBox "Data has timed out. Last refresh was more than 10 minutes ago. Please Refresh the data before saving."
  End If

End Sub

Sub RevertPrevalidation()
  'ImportData and Table location'
  TimeIn
  ImportData 16, 1
End Sub

Sub usertest()
  MsgBox Environ("username")
End Sub

'Form Request; need to be generalized'

Sub SaveColoRequest()

  Dim incrTestValue As Integer
  Dim incrDefaultValue As Integer, strMsgError As String

    If MsgBox("Do you want to record this Colocation Request ?", vbYesNo, "New Colocation Request") = 6 Then
        TestMandatory

        Application.ScreenUpdating = False

    End If

End Sub

Sub OpenLoginForm()
    
    If TestVersion = True Then
        LoginForm.Show
    End If
    
End Sub

Function TestVersion() As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    TestVersion = True
    
    ConnectProductionServer
    rs.Open "SELECT Version FROM client_version", oConn
        
    rs.MoveFirst
    If Range("Version").Value < rs.Fields("Version") Then
        MsgBox "Version too old, please open a version more recent than v" & rs.Fields("Version")
        TestVersion = False
    End If
    
    oConn.Close
    Set oConn = Nothing
    Set rs = Nothing
    
End Function
