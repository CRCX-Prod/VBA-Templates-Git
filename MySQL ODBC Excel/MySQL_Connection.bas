Attribute VB_Name = "MySQL_Connection"
Option Explicit


Public oConn As ADODB.Connection

'Sub ConnectionTest()
'
''____________________________________________________________________
'
''Connection with variables
'
'Dim server_name As String, database_name As String, user_id As String, password As String
'
'    server_name = "localhost"
'    database_name = "coloprocess"
'    user_id = "root"
'    password = "P@ssw0rd"
'
'ConnectDB server_name, database_name, user_id, password
'
''____________________________________________________________________
'GoTo skipComment
''Connection to localhost:
'    ConnectDB "localhost", "coloprocess", "root", "P@ssw0rd"
'
''Connection to 192.168.1.153:
'    ConnectDB "192.168.1.153", "aavbaconnection", "admin", "P@ssw0rd"
'
'skipComment:
'
'End Sub

Sub ConnectProductionServer()

    ConnectDB "192.168.1.153", "00_colo_production", "colouser", "colouser"

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


'_________________________________________________________

'AppendData

'input Sheet number
'Table Name is B1
'Take fields values from line 2
'take appened values from line 3
'_________________________________________________________

Sub AppendData()


Dim rs As ADODB.Recordset
Dim field(), Sql, SqlFields As String
Dim i As Integer

Set rs = New ADODB.Recordset
i = 1

ConnectProductionServer
            
        While Cells(2, i) <> ""
            
            If i = 1 Then   'field 1
            
                SqlFields = Cells(2, i) & "= '" & Cells(3, i) & "'"
            Else            'other fields
                
                SqlFields = SqlFields & "," & Cells(2, i) & "= '" & Cells(3, i) & "'"
            End If
         
        i = i + 1
        
        Wend
        
        Sql = "INSERT INTO " & Cells(1, 2) & " SET " & SqlFields
        
        'MsgBox Sql
        oConn.Execute Sql

oConn.Close
Set oConn = Nothing
Set rs = Nothing

End Sub
