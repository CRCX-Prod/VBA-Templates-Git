Attribute VB_Name = "CheckFile"
Option Explicit

'________________________________________________

'Cette Fonction permet d'afficher differentes strings en fonction de l'existance d'un fichier ou non 
'________________________________________________

Public Function CheckPath(sPath As String, sIfTrue As String, sIfFalse As String) As String
    
    Dim sCheck As String
    
    If Dir(sPath) <> "" Then
        sCheck = sIfTrue
    Else
        sCheck = sIfFalse
    End If
    
    CheckPath = sCheck
    
End Function


