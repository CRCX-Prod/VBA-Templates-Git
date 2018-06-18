VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3885
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLogin_Change()

End Sub

Private Sub cmdOK_Click()
    
    TestPassword cmbLogin.Value, txtPassword.Value
    
End Sub

Private Sub UserForm_Initialize()
    
    cmbLoginAdditem
       
End Sub


Private Sub cmbLoginAdditem()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ConnectProductionServer
    rs.Open "SELECT Login FROM 06preva_admin ORDER BY Login", oConn
    
    Do Until rs.EOF
      cmbLogin.AddItem rs.Fields("Login")
      rs.MoveNext
    Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing
End Sub

