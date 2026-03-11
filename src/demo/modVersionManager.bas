Attribute VB_Name = "modVersionManager"
Option Explicit

'---------------------------------------------------------------
' バージョン管理モジュール
'---------------------------------------------------------------

Private Const myVersion As String = "v1.5.1"
Private Const latestVerUrl As String = "https://raw.githubusercontent.com/mafuji/ChartFormatter-for-Excel/main/version.txt"
Private Const releaseUrl As String = "https://github.com/mafuji/ChartFormatter-for-Excel/releases"

Public Function IsLatestVersion() As Boolean

    Dim latestVersion As String
    Dim result As Boolean
    result = True

    ' 自身が最新バージョンかどうかチェック
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", latestVerUrl, False
    http.SetRequestHeader "User-Agent", "Excel"
    http.Send
    
    If http.Status = 200 Then
        Dim txt As String
        txt = Trim(http.ResponseText)
        txt = Replace(txt, vbCrLf, "")
        txt = Replace(txt, vbCr, "")
        txt = Replace(txt, vbLf, "")
        latestVersion = txt
        
        If myVersion < latestVersion Then
            result = False
        End If
    End If
    
    Set http = Nothing
    
    IsLatestVersion = result
    
End Function

Public Sub OpenReleaseSite()

    ' GitHubのリリースサイトを開く
    ThisWorkbook.FollowHyperlink releaseUrl

End Sub
