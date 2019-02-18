Attribute VB_Name = "Triggers"
Option Explicit

'Triggers are the codes directly runned by Buttons'

Sub UpdatePrevalidation()
  'UpdateData and Table location'
  Dim isTimeOut As Boolean
  
  isTimeOut = TimeOut(10)
  
  If isTimeOut = True Then
    UpdateData 25, 1 'Can change the position of the table here
    MsgBox "Data recorded in Database."
  Else
    MsgBox "Data has timed out. Last refresh was more than 10 minutes ago. Please Refresh the data before saving."
  End If

End Sub

Sub RevertPrevalidation()
  'ImportData and Table location'
  TimeIn
  ImportData 25, 1
End Sub

Sub usertest()
  MsgBox Environ("username")
End Sub

'Form Request; need to be generalized'

Sub SaveNewData()

  Dim incrTestValue As Integer
  Dim incrDefaultValue As Integer, strMsgError As String

    If MsgBox("Do you want to add New Data  ?", vbYesNo, "New Data") = 6 Then
        TestMandatory

        Application.ScreenUpdating = False

        RevertPrevalidation

    End If

End Sub