Attribute VB_Name = "modVersionManager"
Option Explicit

'---------------------------------------------------------------
' バージョン管理モジュール
'---------------------------------------------------------------

Public gLatestVersion As String
Public Const MY_VERSION As String = "v1.7.0"
Public Const RELEASE_SITE_URL As String = "https://github.com/mafuji/ChartFormatter-for-Excel/releases"

Const LATEST_VER_URL As String = "https://raw.githubusercontent.com/mafuji/ChartFormatter-for-Excel/main/version.txt"

' --- 直接更新用の設定項目 ---
Const OWNER As String = "mafuji" ' GitHubのユーザー名
Const REPO As String = "ChartFormatter-for-Excel" ' リポジトリ名
Const ZIP_NAME As String = "ChartFormatter.zip" ' GitHub上のZip名
Const README_NAME As String = "readme.txt" ' readme名
Const TEMP_FOLDER_NAME As String = "ChartFormatter_0qc34mctq94thm9q" ' zip一時保存用フォルダ名
Const BOMB As String = "bomb_mc4893thug9m8c58v59m4y9mh" ' 自爆キーワード
' ----------------

Public Function IsLatestVersion() As Boolean

    Dim result As Boolean
    result = True

    ' 自身が最新バージョンかどうかチェック
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", LATEST_VER_URL, False
    http.SetRequestHeader "User-Agent", "Excel"
    http.Send
    
    If http.Status = 200 Then
        Dim txt As String
        txt = Trim(http.ResponseText)
        txt = Replace(txt, vbCrLf, "")
        txt = Replace(txt, vbCr, "")
        txt = Replace(txt, vbLf, "")
        gLatestVersion = txt
        
        If MY_VERSION < gLatestVersion Then
            result = False
        End If
    End If
    
    Set http = Nothing
    
    IsLatestVersion = result
    
End Function

Sub ShowUpdateDialog()

    frmUpdate.Show

End Sub

Sub UpdateMe()

    Dim tempDir As String, zipPath As String
    Dim latestUrl As String, exePath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. パスの設定 (%TEMP%内)
    tempDir = fso.GetFolder(Environ("TEMP")).Path & "\" & TEMP_FOLDER_NAME & "\"
    zipPath = tempDir & ZIP_NAME
    
    ' 前回の残骸があれば掃除
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFolder Left(tempDir, Len(tempDir) - 1), True
    On Error GoTo 0
    MkDir tempDir
    
    ' 2. GitHubの最新リリースからZipの直リンクを構築
    ' ※簡易化のため、APIを使わずリダイレクトURLを利用
    latestUrl = "https://github.com/" & OWNER & "/" & REPO & "/releases/latest/download/" & ZIP_NAME
    
    Debug.Print "Downloading: " & latestUrl
    If Not DownloadFile(latestUrl, zipPath) Then
        MsgBox "ダウンロードに失敗しました。URLを確認してください。"
        Exit Sub
    End If
    
    ' 3. ZIPの解凍 (Windows標準機能を利用)
    Debug.Print "Extracting..."
    If Not Unzip(zipPath, tempDir) Then
        MsgBox "解凍に失敗しました。"
        If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, Force:=True ' Force引数をTrueにすると、読み取り専用ファイルも強制削除
        Exit Sub
    End If
    
    ' 3-1. ZIPとreadmeの掃除
    Dim readmePath As String
    
    readmePath = tempDir & README_NAME
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, Force:=True
    If fso.FileExists(readmePath) Then fso.DeleteFile readmePath, Force:=True
    
    ' 4. EXEの特定とMotW解除 (Zone.Identifierの削除)
    ' ※解凍先フォルダ内の .exe を探す
    exePath = Dir(tempDir & "*.exe")
    If exePath <> "" Then
        exePath = tempDir & exePath
        Debug.Print "Unblocking & Launching: " & exePath
        
        ' MotW解除コマンド (PowerShell経由が最も確実)
        Shell "powershell -Command ""Unblock-File -Path '" & exePath & "'""", vbHide
        
        ' 5. "bomb" 引数付きで実行
        Dim cmd As String
        cmd = """" & exePath & """ " & BOMB
        Shell cmd, vbNormalFocus
        
        ' 6. VBA側を終了
        Dim wb As Workbook
        
        For Each wb In Workbooks
            If wb.Name = ThisWorkbook.Name Then
                wb.Close SaveChanges:=False
            End If
        Next
        ThisWorkbook.Saved = True
        Application.Quit
    Else
        MsgBox "EXEが見つかりませんでした。"
    End If
End Sub

' --- ヘルパー関数群 ---

' ファイルダウンロード (WinHttp)
Function DownloadFile(url As String, filePath As String) As Boolean
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    
    If http.Status = 200 Then
        Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1: stream.Open
        stream.Write http.ResponseBody
        stream.SaveToFile filePath, 2
        stream.Close
        DownloadFile = True
    End If
End Function

' ZIP解凍 (Shell.Application)
Function Unzip(zipFile As String, destFolder As String) As Boolean
    On Error Resume Next
    Dim shellApp As Object: Set shellApp = CreateObject("Shell.Application")
    If Dir(destFolder, vbDirectory) = "" Then MkDir destFolder
    
    ' Namespace(dest) に Namespace(zip).Items をコピー
    shellApp.Namespace(CVar(destFolder)).CopyHere shellApp.Namespace(CVar(zipFile)).Items, 4 + 16
    Unzip = (Err.Number = 0)
End Function


