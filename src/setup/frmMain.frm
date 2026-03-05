VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "ChartFromatterSetup"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7308
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isInstalled As Boolean

Private Sub UserForm_Initialize()
    
    ' インストール状況確認
    isInstalled = IsAddInEnabledInUserAddins(ADDIN_NAME)

    If isInstalled Then
        btnUninstall.Visible = True
        btnInstall.Caption = "Update"
        lblGuid.Caption = _
            ADDIN_NAME & "アドインが以下の場所にインストール済みです。" & vbCrLf & vbCrLf & _
            Application.UserLibraryPath & ADDIN_NAME & vbCrLf & vbCrLf & _
            "＜操作＞" & vbCrLf & _
            "Uninstall：アドイン登録を解除し、上記のファイルを削除します。" & vbCrLf & _
            "Update：アドインファイルを新しいものに上書き保存します。" & vbCrLf & _
            "(コピー元：" & ThisWorkbook.Path & Application.PathSeparator & ADDIN_NAME & "）"
    Else
        btnUninstall.Visible = False
        btnInstall.Caption = "Install"
        lblGuid.Caption = _
            "Installボタンで" & ADDIN_NAME & "アドインをインストールします。" & vbCrLf & vbCrLf & _
            "保存先：" & Application.UserLibraryPath & ADDIN_NAME & vbCrLf & _
            "コピー元：" & ThisWorkbook.Path & Application.PathSeparator & ADDIN_NAME
    End If

End Sub

Private Sub btnInstall_Click()

    Dim msg As String
    
    If isInstalled Then
        ' アップデート
        msg = ADDIN_NAME & "アドインを更新します。よろしいですか？"
        If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DisableAddIn ADDIN_NAME
        InstallAddin
    Else
        ' インストール
        msg = ADDIN_NAME & "アドインをインストールします。よろしいですか？"
        If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
        InstallAddin
    End If
    

    TerminateMe

End Sub

Private Sub btnUninstall_Click()

    Dim msg As String
    
    ' アンインストール処理
    msg = ADDIN_NAME & "アドインをアンインストールします。よろしいですか？"
    If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    UninstallAddin
    TerminateMe

End Sub

Private Sub btnCancel_Click()

    TerminateMe

End Sub

Private Sub UserForm_Terminate()

    TerminateMe

End Sub

Private Sub TerminateMe()

    If Workbooks.Count < 2 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub
