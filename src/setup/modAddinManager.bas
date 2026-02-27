Attribute VB_Name = "modAddinManager"
Option Explicit

' アドインファイル名を変えれば汎用的に使用可能
Public Const ADDIN_NAME As String = "ChartFormatter.xlam"

Public Sub InstallAddin()
    
    If CopyXlamToUserAddIns(ADDIN_NAME) = False Then Exit Sub
    EnableAddInFromUserAddins ADDIN_NAME

End Sub

Public Sub UninstallAddin()

    DisableAddIn ADDIN_NAME
    RemoveAddinFile ADDIN_NAME

End Sub

'◆ ThisWorkbook と同じ場所にある .xlam を AddIns 既定フォルダへコピーする
'   - srcName にはコピーしたい .xlam ファイル名（拡張子含む）を指定
Private Function CopyXlamToUserAddIns(ByVal srcName As String) As Boolean

    CopyXlamToUserAddIns = False

    Dim srcFolder As String, srcPath As String
    Dim dstFolder As String, dstPath As String
    Dim fso As Object
    
    ' 元フォルダ：このアドイン/インストーラと同じ場所
    srcFolder = ThisWorkbook.Path
    srcPath = srcFolder & Application.PathSeparator & srcName
    
    ' 先フォルダ：ユーザーの既定アドインフォルダ（AddIns）
    ' 例）C:\Users\<User>\AppData\Roaming\Microsoft\AddIns\
    dstFolder = Application.UserLibraryPath
    dstPath = dstFolder & srcName
    
    ' 事前チェック
    If Len(Dir$(srcPath, vbNormal)) = 0 Then
        MsgBox "コピー元が見つかりません。" & ADDIN_NAME & "をこのファイルと同じフォルダに配置して再度実行してください。" & _
                vbCrLf & srcPath, vbExclamation, "CopyXlamToUserAddIns"
        Exit Function
    End If
    
    ' フォルダが無い環境（稀）に備えて作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dstFolder) Then
        On Error Resume Next
        fso.CreateFolder dstFolder
        If Err.Number <> 0 Then
            MsgBox "既定アドインフォルダを作成できませんでした。" & vbCrLf & dstFolder _
                   & vbCrLf & "権限/プロファイルを確認してください。", vbCritical, "CopyXlamToUserAddIns"
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' 上書きコピー（ファイル使用中だと失敗します）
    On Error Resume Next
    fso.CopyFile srcPath, dstPath, True  ' True=上書き許可
    If Err.Number <> 0 Then
        Dim errMsg As String
        errMsg = "コピーに失敗しました。" & vbCrLf & _
                 "From: " & srcPath & vbCrLf & _
                 "To  : " & dstPath & vbCrLf & _
                 "Error " & Err.Number & " : " & Err.Description
        On Error GoTo 0
        MsgBox errMsg, vbCritical, "CopyXlamToUserAddIns"
        Exit Function
    End If
    On Error GoTo 0
    
    CopyXlamToUserAddIns = True
    
End Function

' 指定ファイル名(.xlam)をユーザーのAddInsフォルダから有効化
' - 既存登録（別パス同名/旧版）がある場合は先に無効化
' - 最終的に UserLibraryPath\<addinFileName> を登録・有効化
Private Sub EnableAddInFromUserAddins(ByVal addinFileName As String)
        
    Dim dstFolder As String, targetFull As String
    Dim ai As AddIn, hit As AddIn, removed As Long

    dstFolder = Application.UserLibraryPath ' 例: C:\Users\<User>\AppData\Roaming\Microsoft\AddIns\
    targetFull = dstFolder & addinFileName

    ' 0) 物理ファイル存在チェック
    If Len(Dir$(targetFull, vbNormal)) = 0 Then
        MsgBox "アドインが見つかりません: " & vbCrLf & targetFull, vbExclamation, "EnableAddIn"
        Exit Sub
    End If

    ' 1) 同名・類似名を事前整理（タイトルや名前一致で無効化）
    '    - 別パスの同名や旧版が有効になっていると取り違えが起きるため
    For Each ai In Application.AddIns
        If ai.Installed Then
            ' 名前かタイトルに対象ファイル名のベースが含まれる場合に候補として扱う
            If LCase$(ai.Name) = LCase$(addinFileName) _
               Or InStr(1, LCase$(ai.Title), LCase$(Left$(addinFileName, InStrRev(addinFileName, ".") - 1)), vbTextCompare) > 0 Then
                ' ただし今回有効化したいファイルそのものは除外（FullName一致で判定）
                If StrComp(ai.FullName, targetFull, vbTextCompare) <> 0 Then
                    ai.Installed = False   ' 旧版/別パスは無効化
                    removed = removed + 1
                End If
            End If
        End If
    Next

    ' 2) 既に登録済みか探す（FullName一致を優先）
    Set hit = Nothing
    For Each ai In Application.AddIns
        If StrComp(ai.FullName, targetFull, vbTextCompare) = 0 Then
            Set hit = ai
            Exit For
        End If
    Next

    ' 3) 未登録なら登録
    If hit Is Nothing Then
        Set hit = Application.AddIns.Add(Filename:=targetFull)    ' 登録
    End If

    ' 4) 有効化（Installed=True）
    On Error Resume Next
    hit.Installed = True
    If Err.Number <> 0 Then
        MsgBox "インストールが失敗しました。", vbExclamation
        UninstallAddin
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 5) 完了通知（ログ用途）
    'Debug.Print "Removed old refs: "; removed; " -> Enabled: "; hit.FullName
    MsgBox "アドインを有効化しました:" & vbCrLf & hit.FullName & _
           IIf(removed > 0, vbCrLf & "(旧版/別パス " & removed & " 件を無効化)", ""), vbInformation, "EnableAddIn"
    
End Sub

' UserLibraryPath にあるアドインを無効化
Public Sub DisableAddIn(ByVal addinFileName As String)

    Dim ai As AddIn, full As String
    full = Application.UserLibraryPath & addinFileName
    
    For Each ai In Application.AddIns
        If StrComp(ai.FullName, full, vbTextCompare) = 0 Then
            If ai.Installed Then ai.Installed = False   ' ← 無効化
            Exit For
        End If
    Next
    
End Sub

' AddIns 既定フォルダから .xlam を削除（退避パスに移動する版）
Private Sub RemoveAddinFile(ByVal addinFileName As String, Optional ByVal moveTo As String = "")

    Dim src As String, fso As Object
    src = Application.UserLibraryPath & addinFileName
    
    If Len(Dir$(src, vbNormal)) = 0 Then
        MsgBox "対象ファイルが見つかりません。" & vbCrLf & src, vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If moveTo <> "" Then
        If Not fso.FolderExists(moveTo) Then fso.CreateFolder moveTo
        fso.MoveFile src, moveTo & Application.PathSeparator & addinFileName  ' 退避
    Else
        fso.DeleteFile src, True                                             ' 削除
    End If
    If Err.Number <> 0 Then
        MsgBox "削除/移動に失敗しました: " & Err.Description, vbCritical
    Else
        MsgBox "アドインファイルを処理しました。", vbInformation
    End If
    
End Sub

' 指定の .xlam（ユーザー既定 AddIns フォルダ内）が有効かどうかを返す
Public Function IsAddInEnabledInUserAddins(ByVal addinFileName As String) As Boolean
    
    Dim targetFull As String, ai As AddIn
    targetFull = Application.UserLibraryPath & addinFileName   ' 例: ...\AddIns\MyTool.xlam

    For Each ai In Application.AddIns
        If StrComp(ai.FullName, targetFull, vbTextCompare) = 0 Then
            IsAddInEnabledInUserAddins = ai.Installed          ' ← True=有効
            Exit Function
        End If
    Next
    
    ' 見つからなければ未登録（＝False相当）を返す
    IsAddInEnabledInUserAddins = False

End Function
