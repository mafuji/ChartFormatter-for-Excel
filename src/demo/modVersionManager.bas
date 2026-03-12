Attribute VB_Name = "modVersionManager"
Option Explicit

'---------------------------------------------------------------
' バージョン管理モジュール
'---------------------------------------------------------------

Private Const myVersion As String = "v1.5.1"
Private Const latestVerUrl As String = "https://raw.githubusercontent.com/mafuji/ChartFormatter-for-Excel/main/version.txt"
Private Const releaseUrl As String = "https://github.com/mafuji/ChartFormatter-for-Excel/releases"

' --- 設定項目 ---
Const OWNER As String = "mafuji"      ' GitHubのユーザー名
Const REPO As String = "ChartFormatter-for-Excel"        ' リポジトリ名
Const ZIP_NAME As String = "ChartFormatter_v1.5.1.zip" ' GitHub上のZip名
' ----------------

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

Sub UpdateMe()

    If MsgBox("最新版をダウンロードしてアドインを更新します。よろしいですか？" & vbCrLf & _
              "※今開いているExcelは保存せずに終了されます。", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Dim tempDir As String, zipPath As String, extractDir As String
    Dim latestUrl As String, exePath As String
    
    ' 1. パスの設定 (%TEMP%\AppUpdate_Test)
    tempDir = Environ("TEMP") & "\AppUpdate_Test\"
    zipPath = tempDir & ZIP_NAME
    extractDir = tempDir & "Extract\"
    
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
    If Not Unzip(zipPath, extractDir) Then
        MsgBox "解凍に失敗しました。"
        Exit Sub
    End If
    
    ' 4. EXEの特定とMotW解除 (Zone.Identifierの削除)
    ' ※解凍先フォルダ内の .exe を探す
    exePath = Dir(extractDir & "*.exe")
    If exePath <> "" Then
        exePath = extractDir & exePath
        Debug.Print "Unblocking & Launching: " & exePath
        
        ' MotW解除コマンド (PowerShell経由が最も確実)
        Shell "powershell -Command ""Unblock-File -Path '" & exePath & "'""", vbHide
        
        ' 5. "bomb" 引数付きで実行
        ' 引数: bomb [TEMPのパス]
        Dim cmd As String
        cmd = """" & exePath & """ bomb """ & extractDir & """"
        Shell cmd, vbNormalFocus
        
        ' 6. VBA側を終了
        ThisWorkbook.Close SaveChanges:=False
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
