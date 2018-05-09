VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3885
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLogin_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblLogin_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Dim sArray() As String
    
    ReDim sArray(2)
    sArray(1) = "Charles"
    
    cmbLogin.AddItem sArray(1)
    
    
    
End Sub


Sub runTest()

    Sql
End Sub
