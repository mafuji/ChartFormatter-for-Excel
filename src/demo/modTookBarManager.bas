Attribute VB_Name = "modTookBarManager"
Option Explicit

'---------------------------------------------------------------
' ツールバー管理モジュール
'---------------------------------------------------------------

' コマンドバーの名前のベース
Public Const BAR_NAME As String = "ChartFormatter"

' コマンドバーインデックスの開始番号
Public Const BAR_INDEX_START As Integer = 0

Sub CreateToolbar()

    ' コマンドバー削除
    DeleteToolBar
    
    ' コマンドバー作成
    Dim barNo As Integer: barNo = BAR_INDEX_START
    Dim bar As clsCommandBar
    
    Set bar = New clsCommandBar
    With bar
        .No = barNo
        .Caption = "散布図整形"
        .OnAction = "ShowChartFormatForm"
        .Style = msoButtonIconAndCaption
        .FaceId = 430
    End With
    bar.CreateBar
    
    barNo = barNo + 1
    Set bar = New clsCommandBar
    With bar
        .No = barNo
        .Caption = "充放電グラフ作成"
        .OnAction = "ShowChargeDischargeChartBuilder"
        .Style = msoButtonIconAndCaption
        .FaceId = 422
    End With
    bar.CreateBar
    
End Sub

Sub DeleteToolBar()

    Dim barNo As Integer
    
    barNo = BAR_INDEX_START
    On Error Resume Next
    Do
        Application.CommandBars(BAR_NAME & barNo).Delete
        If Err.Number <> 0 Then Exit Do
        barNo = barNo + 1
    Loop
    On Error GoTo 0

End Sub
