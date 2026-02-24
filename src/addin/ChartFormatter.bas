Attribute VB_Name = "ChartFormatter"
Option Explicit

Public colorPalette As New clsColorPalette
Public Const DEFAULT_FONT_NAME As String = "Arial Narrow"
Public Const DEFAULT_FONT_SIZE As Single = 16
Public Const DEFAULT_MARKER_SIZE As Integer = 6

Private Const BAR_NAME As String = "ChartFormatter"

Sub ShowChartFormatForm()

    ' 選択状態チェック
    Dim targetChart As Chart
    
    Set targetChart = GetChart(Selection)
    If targetChart Is Nothing Then
        MsgBox "整形対象のグラフを選択した状態でボタンを押してください。", vbExclamation
        Exit Sub
    End If
    Set frmSetting.targetChart = targetChart
    frmSetting.Show

End Sub

Sub FormatTwoAxisChart()
    
    ' 選択状態チェック
    Dim targetChart As Chart
    
    Set targetChart = GetChart(Selection)
    If Not targetChart Is Nothing Then
        Set targetChart = Selection.Parent
    Else
        MsgBox "整形対象のグラフを選択した状態でボタンを押してください。", vbExclamation
        Exit Sub
    End If
    
    Dim scatterChart As New clsScatterChart
    
    With scatterChart
        .SetDefault targetChart
        .ApplyTo targetChart
    End With

End Sub

' データ範囲を取得する
Function GetRangeFromSeries(s As Series, targetAxis As Integer) As Range

    Dim f As String
    Dim parts As Variant

    f = s.Formula
    f = Replace(f, "=SERIES(", "")
    f = Left(f, Len(f) - 1)

    parts = Split(f, ",")

    'parts(1) が XValues の範囲
    On Error Resume Next
    Set GetRangeFromSeries = Range(parts(targetAxis))
    On Error GoTo 0

End Function

' アドレスから列番号取得
Function GetColumnLetter(rng As Range) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^[A-Z]+"
    re.IgnoreCase = False

    Dim m As Object
    If re.test(rng.Address(False, False)) Then
        Set m = re.Execute(rng.Address(False, False))(0)
        GetColumnLetter = m.Value
    End If
End Function

' オブジェクトの上位にChartがいれば返す。居なければnothingを返す
Function GetChart(ByVal targetObject As Object) As Object

    Dim obj As Object

    Set obj = targetObject
    
    On Error Resume Next
    
    Do Until Err.Number <> 0
        If TypeName(obj) = "Chart" Then
            Set GetChart = obj
            Exit Function
        Else
            Set obj = obj.Parent
        End If
        
        If Err.Number <> 0 Or TypeName(obj) = "Workbook" Then
            Set GetChart = Nothing
            Exit Function
        End If
    Loop
    
    On Error GoTo 0

    Set GetChart = Nothing

End Function

Sub CreateToolbar()

    Dim bar As CommandBar, btn As CommandBarButton
    
    On Error Resume Next
    Application.CommandBars(BAR_NAME).Delete
    On Error GoTo 0

    Set bar = Application.CommandBars.Add(Name:=BAR_NAME, Position:=msoBarTop, Temporary:=True)
'    Set btn = bar.Controls.Add(Type:=msoControlButton)
'    With btn
'        .Caption = "散布図整形_クイック"
'        .OnAction = "FormatTwoAxisChart"
'        .Style = msoButtonIconAndCaption
'        .FaceId = 17 ' 任意アイコン
'    End With
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "散布図整形"
        .OnAction = "ShowChartFormatForm"
        .Style = msoButtonIconAndCaption
        .FaceId = 17 ' 任意アイコン
    End With
    bar.Visible = True
    
End Sub

Sub DeleteToolBar()

    On Error Resume Next
    Application.CommandBars(BAR_NAME).Delete
    On Error GoTo 0

End Sub
