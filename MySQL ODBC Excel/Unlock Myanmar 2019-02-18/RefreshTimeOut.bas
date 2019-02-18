Attribute VB_Name = "RefreshTimeOut"
Option Explicit

Public dateTimeIn As Date
Public dateTimeOut As Date

Sub TimeIn() 'Add TimeIn while Refresh

    dateTimeIn = Now()
    'MsgBox dateTimeIn
    
End Sub

Function TimeOut(maxTime As Single) As Boolean 'Add TimeOut in Save Data

    Dim dateTest As Date
    
    dateTimeOut = Now() '+ (1 / 144)  ' + 10 Min
    dateTest = dateTimeOut - dateTimeIn
    
    'MsgBox dateTimeIn
    'MsgBox dateTimeOut
    'MsgBox dateTest
    
    If dateTest < (maxTime / 1440) Then '1 minute = 1/1440
        
        TimeOut = True
    Else
        TimeOut = False
    End If
    
    
End Function
