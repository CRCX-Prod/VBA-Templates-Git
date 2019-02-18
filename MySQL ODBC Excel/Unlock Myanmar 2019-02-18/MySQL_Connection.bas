Attribute VB_Name = "MySQL_Connection"
Option Explicit

'__________________________________________________________

'Configuration et Connection a la base de donnees
'__________________________________________________________


Public oConn As ADODB.Connection

Sub ConnectProductionServer()
    'Enter here Connection informations'
    
    ConnectDB "192.168.1.153", "01_unlmm", "unlmm", "unlmm"
End Sub

Sub ConnectDB(server_name As String, database_name As String, user_id As String, password As String)
  Set oConn = New ADODB.Connection
  Dim str As String
  str = "DRIVER={MySQL ODBC 5.3 ANSI Driver};" & _
        ";SERVER=" & server_name & _
        ";PORT=3306" & _
        ";DATABASE=" & database_name & _
        ";UID=" & user_id & _
        ";PWD=" & password & _
        ";Option=3"

  oConn.Open str
End Sub

