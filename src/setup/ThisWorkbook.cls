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

    If Workbooks.Count < 2 Then
        Application.Visible = False
    Else
        Windows(ThisWorkbook.Name).Visible = False
    End If
    frmMain.Show vbModal

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ThisWorkbook.Saved = True

End Sub
