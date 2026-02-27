Attribute VB_Name = "modExporter"
Option Explicit

Public Sub ExportAllXlsmInBin()

    Dim baseDir As String
    baseDir = ThisWorkbook.path  ' bin フォルダ

    Dim parentDir As String
    parentDir = GetParentFolder(baseDir)  ' 既定 = bin の一つ上

    Dim destRoot As String
    destRoot = PickFolder(parentDir)
    If destRoot = "" Then
        MsgBox "キャンセルされました", vbInformation
        Exit Sub
    End If

    Dim newSrc As String
    newSrc = destRoot & "\new-src"
    If IsFolderExists(newSrc) = False Then
        MkDir newSrc
    Else
        If MsgBox("new-srcが既に存在します。上書き保存して良いですか？", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    Dim fileList As Collection
    Set fileList = New Collection

    Call FindXlsmRecursive(baseDir, fileList)

    Dim f As Variant
    For Each f In fileList
        ExportBookModules CStr(f), newSrc
    Next

    MsgBox "new-srcが作成されました。" & vbCrLf & newSrc, vbInformation
End Sub

Private Sub FindXlsmRecursive(dirPath As String, list As Collection)
    Dim f As String

    ' .xlsm ファイル取得
    f = Dir(dirPath & "\*.xlsm")
    Do While f <> ""
        list.Add dirPath & "\" & f
        f = Dir()
    Loop

    ' サブフォルダ再帰
    f = Dir(dirPath & "\*", vbDirectory)
    Do While f <> ""
        If (f <> ".") And (f <> "..") Then
            If (GetAttr(dirPath & "\" & f) And vbDirectory) <> 0 Then
                Call FindXlsmRecursive(dirPath & "\" & f, list)
            End If
        End If
        f = Dir()
    Loop
End Sub

Private Sub ExportBookModules(bookPath As String, destRoot As String)

    Dim presentSetting As Boolean
    
    presentSetting = Application.EnableEvents
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = Workbooks.Open(bookPath)

    Dim outDir As String
    outDir = destRoot & "\" & RemoveExt(wb.Name)
    EnsureFolderExists outDir

    Dim vbc As Object
    For Each vbc In wb.VBProject.VBComponents
        Dim ext As String
        Select Case vbc.Type
            Case 1: ext = ".bas"
            Case 2: ext = ".cls"
            Case 3: ext = ".frm"
            Case 100: ext = ".cls"
            Case Else: ext = ".txt"
        End Select

        vbc.Export outDir & "\" & vbc.Name & ext
    Next

    If bookPath <> ThisWorkbook.FullName Then
        wb.Close False
    End If
    
    Application.EnableEvents = presentSetting
    
End Sub

Private Function PickFolder(initial As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.InitialFileName = initial & "\"
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
End Function

Private Function IsFolderExists(p As String) As Boolean
    If Dir(p, vbDirectory) = "" Then
        IsFolderExists = False
    Else
        IsFolderExists = True
    End If
End Function

Private Sub EnsureFolderExists(p As String)
    If Dir(p, vbDirectory) = "" Then MkDir p
End Sub

Private Function RemoveExt(fn As String) As String
    RemoveExt = Left(fn, InStrRev(fn, ".") - 1)
End Function

Private Function GetParentFolder(path As String) As String
    GetParentFolder = Left(path, InStrRev(path, "\") - 1)
End Function
