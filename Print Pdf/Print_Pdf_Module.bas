Attribute VB_Name = "Print_Pdf_Module"
Option Explicit

Sub PrintPdf(sFileName As String, sPathFolder As String)
    
    Dim sMessage As String, fileTest As Object
    
    Set fileTest = CreateObject("Scripting.FileSystemObject")
    
    
    If fileTest.fileExists(sPathFolder & sFileName & ".pdf") = True Then
    
        sMessage = sFileName & ".pdf" & Chr(10) & Chr(10) & "Already exists in the folder :" & Chr(10) & Chr(10) & sPathFolder & Chr(10) & Chr(10) & "Please delete the file before generating a new one."
        'MsgBox sMessage

    Else
      
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=sPathFolder & sFileName & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False

        sMessage = sFileName & ".pdf" & Chr(10) & Chr(10) & "is saved in the folder :" & Chr(10) & Chr(10) & sPathFolder
        'MsgBox sMessage
    
    End If

End Sub

Function FolderPathFromFile(sFolderName As String) As String

    Dim sFilePath As String, iPathFolder As String
    Dim iFindFolder As Integer, iLenFolder As Integer
    
    sFilePath = Application.ActiveWorkbook.Path     'Get the path of the current file
    
    iFindFolder = InStr(sFilePath, sFolderName)     'find position of the folder in the string
    iLenFolder = Len(sFolderName)                   'find lenght of folder
    
    iPathFolder = Left(sFilePath, iFindFolder + iLenFolder) 'Get left side of the Path
    
    FolderPathFromFile = iPathFolder
    

End Function

Sub SaveNoticeAsPDF()

    Dim idMin As Integer, idMax As Integer
    
    Application.ScreenUpdating = False

    idMin = Range("idMin").Value
'    idMin = 2
    idMax = Range("idMax").Value
'    idMax = 10
    
    RunID idMin, idMax
    
    Application.ScreenUpdating = True
    
End Sub

Sub RunID(idMin As Integer, idMax As Integer)
    
    Dim sRootFolder As String, sPathFolder As String, sFileName As String
    Dim sColoID As String, lenColoID As Integer, iLen As Integer

    Dim iID As Integer

    For iID = idMin To idMax

        Range("ColoID") = iID

        If Range("toGenerate") = "Generate PDF" Then

            sRootFolder = FolderPathFromFile("Process & Database")
            sPathFolder = sRootFolder & "Documentation\Handover Notice\Not Signed\"
    
            sColoID = Range("ColoID").Value
            lenColoID = Len(sColoID)

            For iLen = lenColoID To 3

                sColoID = "0" & sColoID
                
            Next

            sFileName = "CAR" & sColoID & " Handover Notice"

            PrintPdf sFileName, sPathFolder

        End If

    Next

End Sub