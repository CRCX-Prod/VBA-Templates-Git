Attribute VB_Name = "Triggers"
Option Explicit

'Triggers are the codes directly runned by Buttons'

Sub UpdatePrevalidation()
  'UpdateData and Table location'
  UpdateData 16, 1
End Sub

Sub RevertPrevalidation()
  'ImportData and Table location'
  ImportData 16, 1
End Sub

Sub usertest()
  MsgBox Environ("username")
End Sub

'Form Request; need to be generalized'

Sub SaveColoRequest()

  Dim incrTestValue As Integer
  Dim incrDefaultValue As Integer, strMsgError As String

'______________________________
'Test Form Values
'______________________________

    For incrTestValue = 8 To 20 Step 2
        If Range("D" & incrTestValue).Value = "" Then
            strMsgError = strMsgError & Range("B" & incrTestValue).Value & Chr(13)
        End If
    Next
    If Range("F20").Value = "" Then
       strMsgError = strMsgError & "Upper Space Required" & Chr(13)
    End If

    If strMsgError <> "" Then
        MsgBox "Please enter:" & Chr(10) & Chr(10) & strMsgError
        Exit Sub
    End If

'______________________________
'Update Database
'______________________________

    If MsgBox("Do you want to record this Colocation Request ?", vbYesNo, "New Colocation Request") = 6 Then
        AppendData
        MsgBox "Colocation added in the Database"

'______________________________
'Restore default Values (need to be subbed)
'______________________________

        Application.ScreenUpdating = False

            For incrDefaultValue = 8 To 20 Step 2
                Range("D" & incrDefaultValue).Value = Range("M" & incrDefaultValue).Value
            Next

        Range("F20").Value = Range("N22").Value
        Application.ScreenUpdating = True
    End If

End Sub

Sub OpenLogin()

    LoginForm.Show

End Sub
