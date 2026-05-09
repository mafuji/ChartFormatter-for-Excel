VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdate 
   Caption         =   "アドインの更新"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5700
   OleObjectBlob   =   "frmUpdate.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    lblMyVersion.Caption = MY_VERSION
    lblLatestVersion.Caption = gLatestVersion
    lblLink.Caption = RELEASE_SITE_URL

End Sub

Private Sub lblLink_Click()

    ThisWorkbook.FollowHyperlink RELEASE_SITE_URL

End Sub

Private Sub btnUpdate_Click()

    If MsgBox("最新版をダウンロードしてアドインを更新します。よろしいですか？" & vbCrLf & vbCrLf & _
              "※「はい」を押すと今開いているExcelは全て保存せずに終了され、" & _
              "最新版のインストーラが起動します。", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Me.MousePointer = fmMousePointerHourGlass
    lblStatus.Visible = True
    DoEvents
    UpdateMe
    Me.MousePointer = fmMousePointerDefault
    lblStatus.Visible = False

End Sub

Private Sub btnClose_Click()

    Unload Me

End Sub
