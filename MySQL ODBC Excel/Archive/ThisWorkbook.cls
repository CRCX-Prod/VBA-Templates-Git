VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    
    HideAllSheets
    isAdmin = False
    RefreshReport
        
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    HideAllSheets
    If isAdmin = False Then
        Application.ThisWorkbook.Saved = True 'Informe Excel que le fichier a déjà été enregistré (cela évite d'avoir une demande d'enregistrement à la fermeture)
    End If
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    
    
    If isAdmin = False Then
        HideAllSheets
        Cancel = True 'Annule les demandes d'enregistrement

        MsgBox "Only saving the excel file won't record the data in the database." & Chr(10) _
            & "To save the data, please click on 'Save Change' button"
    End If

End Sub

Sub RefreshReport()
    
    Application.ScreenUpdating = False
    Worksheets("Review List").Visible = xlSheetVisible
    Worksheets("Review List").Activate
    ImportData 16, 1
    ThisWorkbook.RefreshAll
    Worksheets("Main").Activate
    Application.ScreenUpdating = True

End Sub
