Attribute VB_Name = "modVersionManager"
Option Explicit

'---------------------------------------------------------------
' バージョン管理モジュール
'---------------------------------------------------------------

Private Const myVersion As String = "v1.5.0"
Private Const latestVerUrl As String = "https://"
Private Const releaseUrl As String = "https://"

Public Function IsLatestVersion() As Boolean

    Dim res As Boolean
    res = True

    ' 自身が最新バージョンかどうかチェック
    
    IsLatestVersion = res
    
End Function

Public Sub OpenReleaseSite()

    ' GitHubのリリースサイトを開く

End Sub
