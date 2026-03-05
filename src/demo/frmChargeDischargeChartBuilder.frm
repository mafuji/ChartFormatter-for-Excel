VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChargeDischargeChartBuilder 
   Caption         =   "充放電グラフ作成"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9990
   OleObjectBlob   =   "frmChargeDischargeChartBuilder.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmChargeDischargeChartBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'　充放電グラフ作成フォーム
'---------------------------------------------------------------

Private chargeDischargeChart As New clsChargeDischargeChart
Private chartFormat_ As clsScatterChart

Private Const DEFAULT_CYCLE_COLUMN As String = "B"
Private Const DEFAULT_STEP_COLUMN As String = "C"
Private Const DEFAULT_AH_COLUMN As String = "I"
Private Const DEFAULT_V_COLUMN As String = "F"
Private Const DEFAULT_DATA_START_ROW As Long = 8

Private Sub UserForm_Initialize()

    ' サイクル数既定設定
    With chargeDischargeChart
        chkUseDisplayCycleAdjust = False: .UseDisplayCycleAdjust = False
        txtAlwaysDisplayCycle = 5: .alwaysDisplayCycle = 5
        txtDisplayCycleInterval = 1: .displayCycleInterval = 1
    End With
    SetDisplayCycleAdjust
    
    ' データ範囲既定設定
    txtCycleColumn = DEFAULT_CYCLE_COLUMN
    txtStepColumn = DEFAULT_STEP_COLUMN
    txtAHColumn = DEFAULT_AH_COLUMN
    txtVColumn = DEFAULT_V_COLUMN
    txtDataStartRow = DEFAULT_DATA_START_ROW
    cmbStartStep.List = Array(1, 2, 3, 4)
    cmbStartStep.ListIndex = 0

    ' 開始ステップのChargeTypeを設定
    chargeDischargeChart.StartStep = IIf(optCharge, Charge, Discharge)
    
    ' 線幅既定値
    txtLineWeight = 1
    
    ' 色の既定設定
    txtStartHue = 0
    txtEndHue = 300
    txtSat = 1
    txtVal = 0.85
    lblStartHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtStartHue)(0)
    lblEndHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtEndHue)(0)
    
    ' カラーバーの既定設定
    With chargeDischargeChart
        ' 位置
        optColorBarRight = True: .ColorBar.LabelPosition = lpRight
        
        ' サイズ
        optLengthAuto = True: .ColorBarLengthAuto = True
        SetLengthEnabled
        optWidthAuto = True: .ColorBarWidthAuto = True
        SetWidthEnabled
        
        ' ラベル
        optCardinal = True: .ColorBar.LabelFormat = lfCardinal
        optColorBarUnitCustom = True: .ColorBar.LabelMode = lmCustom
        SetColorBarUnitEnabled
        txtColorBarUnit = 10: .ColorBar.LabelUnit = 10
        chkColorfulLabel = False: .ColorBar.IsColorfulLabel = False
    End With
    
End Sub

Private Sub txtCycleColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If IsValidColumnKeyPress(KeyAscii, txtCycleColumn) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCycleColumn_AfterUpdate()

    Dim rng As range

    On Error Resume Next
    Set rng = range(txtCycleColumn & 1)
    On Error GoTo 0

    If Len(txtCycleColumn) >= 3 Or rng Is Nothing Then
        txtCycleColumn = ""
    End If

End Sub

Private Sub txtStepColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If IsValidColumnKeyPress(KeyAscii, txtStepColumn) = False Then
        KeyAscii = 0
    End If

End Sub

' 対象データ範囲
Private Sub txtStepColumn_AfterUpdate()

    Dim rng As range

    On Error Resume Next
    Set rng = range(txtStepColumn & 1)
    On Error GoTo 0

    If Len(txtStepColumn) >= 3 Or rng Is Nothing Then
        txtStepColumn = ""
    End If

End Sub

Private Sub txtAHColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If IsValidColumnKeyPress(KeyAscii, txtAHColumn) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAHColumn_AfterUpdate()

    Dim rng As range

    On Error Resume Next
    Set rng = range(txtAHColumn & 1)
    On Error GoTo 0

    If Len(txtAHColumn) >= 3 Or rng Is Nothing Then
        txtAHColumn = ""
    End If

End Sub

Private Sub txtVColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If IsValidColumnKeyPress(KeyAscii, txtVColumn) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtVColumn_AfterUpdate()

    Dim rng As range

    On Error Resume Next
    Set rng = range(txtVColumn & 1)
    On Error GoTo 0

    If Len(txtVColumn) >= 3 Or rng Is Nothing Then
        txtVColumn = ""
    End If

End Sub

Private Sub txtDataStartRow_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtDataStartRow, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtDataStartRow_AfterUpdate()

    If (Not IsNumeric(txtDataStartRow)) _
    Or val(txtDataStartRow) <= 0 Then
        txtDataStartRow = ""
    End If

End Sub

Private Sub btnSetStartCells_Click()

    ' 入力チェック
    If txtCycleColumn = "" _
    Or txtStepColumn = "" _
    Or txtVColumn = "" _
    Or txtAHColumn = "" _
    Or txtDataStartRow = "" _
    Or IIf(IsNull(cmbStartStep), "", cmbStartStep) = "" Then
        MsgBox "列番号/データ開始行/開始ステップを全て入力してください。", vbExclamation
        Exit Sub
    End If

    ' 開始セルを探して、フォームとチャートインスタンスにセット
    Dim rng As range
    
    Set rng = range(txtStepColumn & txtDataStartRow)
    Do Until val(rng) = val(cmbStartStep.value) Or rng = ""
        Set rng = rng.Offset(1)
    Loop
    If rng = "" Then
        MsgBox "開始ステップが見つかりませんでした。" & _
               "列番号/データ開始行/開始ステップが、" & _
               "表示中のシートのレイアウトに合致するか確認してください。", vbExclamation
        Exit Sub
    End If

    txtCycleStart = range(txtCycleColumn & rng.Row).Address
    txtStepStart = range(txtStepColumn & rng.Row).Address
    txtAHStart = range(txtAHColumn & rng.Row).Address
    txtVStart = range(txtVColumn & rng.Row).Address

    With chargeDischargeChart
        Set .CycleStartCell = range(txtCycleColumn & rng.Row)
        Set .StepStartCell = range(txtStepColumn & rng.Row)
        Set .AHStartCell = range(txtAHColumn & rng.Row)
        Set .VStartCell = range(txtVColumn & rng.Row)
    End With

End Sub

' 開始セル
Private Sub txtCycleStart_AfterUpdate()

    ValidateRangeText txtCycleStart
    If txtCycleStart.Text <> "" Then
        Set chargeDischargeChart.CycleStartCell = range(txtCycleStart.Text)
    End If

End Sub

Private Sub btnCycleStart_Click()
    
    Me.Hide
    
    OnClickRangeSelecter "CYCLE開始セル", txtCycleStart
    If txtCycleStart.Text <> "" Then
        Set chargeDischargeChart.CycleStartCell = range(txtCycleStart.Text)
    End If
    
    Me.Show

End Sub

Private Sub txtStepStart_AfterUpdate()

    ValidateRangeText txtStepStart
    If txtStepStart.Text <> "" Then
        Set chargeDischargeChart.StepStartCell = range(txtStepStart.Text)
    End If

End Sub

Private Sub btnStepStart_Click()
    
    Me.Hide
    
    OnClickRangeSelecter "STEP開始セル", txtStepStart
    If txtStepStart.Text <> "" Then
        Set chargeDischargeChart.StepStartCell = range(txtStepStart.Text)
    End If
    
    Me.Show

End Sub

Private Sub txtAHStart_AfterUpdate()

    ValidateRangeText txtAHStart
    If txtAHStart.Text <> "" Then
        Set chargeDischargeChart.AHStartCell = range(txtAHStart.Text)
    End If

End Sub

Private Sub btnAHStart_Click()
    
    Me.Hide
    
    OnClickRangeSelecter "AH開始セル", txtAHStart
    If txtAHStart.Text <> "" Then
        Set chargeDischargeChart.AHStartCell = range(txtAHStart.Text)
    End If

    Me.Show

End Sub

Private Sub txtVStart_AfterUpdate()

    ValidateRangeText txtVStart
    If txtVStart.Text <> "" Then
        Set chargeDischargeChart.VStartCell = range(txtVStart.Text)
    End If

End Sub

Private Sub btnVStart_Click()
    
    Me.Hide
    
    OnClickRangeSelecter "V開始セル", txtVStart
    If txtVStart.Text <> "" Then
        Set chargeDischargeChart.VStartCell = range(txtVStart.Text)
    End If

    Me.Show

End Sub

Private Sub OnClickRangeSelecter(ByVal inputboxCaption As String, ByRef ref As MSForms.TextBox)

    Dim rng As range
    
    On Error Resume Next
    Set rng = Application.InputBox(inputboxCaption, Type:=8)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        ref.Text = rng.Address
        ValidateRangeText ref
    End If
    
End Sub

' 選択セルテキストのValidation
Private Sub ValidateRangeText(ByVal ref As MSForms.TextBox, Optional ByVal forceSingleCell As Boolean = True)

    Dim rng As range

    On Error Resume Next
    Set rng = range(ref.Text)
    On Error GoTo 0

    If rng Is Nothing Then
        ' 無効な入力はクリア
        ref.Text = ""
        Exit Sub
    End If

    ' 単一セルに強制
    If forceSingleCell Then
        ref.Text = rng.Cells(1, 1).Address
    End If
    
End Sub

Private Sub optCharge_AfterUpdate()

    chargeDischargeChart.StartStep = Charge

End Sub

Private Sub optDischarge_AfterUpdate()

    chargeDischargeChart.StartStep = Discharge

End Sub

Private Sub chkUseDisplayCycleAdjust_AfterUpdate()

    chargeDischargeChart.UseDisplayCycleAdjust = chkUseDisplayCycleAdjust.value
    SetDisplayCycleAdjust

End Sub

Private Sub SetDisplayCycleAdjust()

    If chkUseDisplayCycleAdjust.value = True Then
        txtAlwaysDisplayCycle.Enabled = True
        txtDisplayCycleInterval.Enabled = True
    Else
        txtAlwaysDisplayCycle.Enabled = False
        txtDisplayCycleInterval.Enabled = False
    End If
    CalculateDisplayCycleCount

End Sub

Private Sub txtCycle_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtCycle, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCycle_AfterUpdate()

    If (Not IsNumeric(txtCycle)) _
    Or Int(val(txtCycle)) <> val(txtCycle) _
    Or val(txtCycle) <= 0 Then
        txtCycle = ""
        chargeDischargeChart.cycleCount = 0
    Else
        chargeDischargeChart.cycleCount = txtCycle.value
    End If
    CalculateDisplayCycleCount
    
End Sub

Private Sub txtAlwaysDisplayCycle_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtAlwaysDisplayCycle, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAlwaysDisplayCycle_AfterUpdate()

    If (Not IsNumeric(txtAlwaysDisplayCycle)) _
    Or Int(val(txtAlwaysDisplayCycle)) <> val(txtAlwaysDisplayCycle) _
    Or val(txtAlwaysDisplayCycle) <= 0 Then
        txtAlwaysDisplayCycle = 5
    End If
    chargeDischargeChart.alwaysDisplayCycle = txtAlwaysDisplayCycle.value
    CalculateDisplayCycleCount

End Sub

Private Sub txtDisplayCycleInterval_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtDisplayCycleInterval, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtDisplayCycleInterval_AfterUpdate()

    If (Not IsNumeric(txtDisplayCycleInterval)) _
    Or Int(val(txtDisplayCycleInterval)) <> val(txtDisplayCycleInterval) _
    Or val(txtDisplayCycleInterval) <= 0 Then
        txtDisplayCycleInterval = 1
    End If
    chargeDischargeChart.displayCycleInterval = txtDisplayCycleInterval.value
    CalculateDisplayCycleCount

End Sub

Private Sub CalculateDisplayCycleCount()

    If chkUseDisplayCycleAdjust = False _
    Or txtCycle = "" Or txtCycle = 0 _
    Or txtAlwaysDisplayCycle = "" _
    Or txtDisplayCycleInterval = "" Or txtDisplayCycleInterval = 0 Then
        txtDisplayCycleCount.value = ""
    Else
        Dim alwaysDisplayCycle As Long
        Dim displayCycleInterval As Long
        Dim cycleCount As Long
        
        alwaysDisplayCycle = txtAlwaysDisplayCycle.value
        displayCycleInterval = txtDisplayCycleInterval.value
        cycleCount = txtCycle.value
        
        txtDisplayCycleCount = _
            alwaysDisplayCycle _
            + Int((cycleCount - alwaysDisplayCycle) / displayCycleInterval)
    End If
    ToggleCycleAlert

End Sub

Private Sub ToggleCycleAlert()

    With lblCycleValidation
        If chkUseDisplayCycleAdjust = False Then
            .Visible = False
        Else
            .Visible = True
            If txtDisplayCycleCount = "" Or txtDisplayCycleCount > 127 Then
                .Caption = "!"
                .ForeColor = vbRed
                .Font.Bold = True
                .ControlTipText = "表示サイクル数が127以下になるようにしてください。" & _
                                  "表示サイクル数=必須表示+(サイクル数-必須表示)/表示間隔" & _
                                  "(小数点以下切捨て)"
            Else
                .Caption = "ok"
                .ForeColor = vbGreen
                .Font.Bold = True
                .ControlTipText = ""
            End If
        End If
    End With

End Sub

' 系列の線幅
Private Sub txtLineWeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtLineWeight, True) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLineWeight_AfterUpdate()

    If (Not IsNumeric(txtLineWeight)) _
    Or val(txtLineWeight) <= 0 Then
        txtLineWeight = 1
    Else
        txtLineWeight = Int((txtLineWeight.value + 0.005) * 100) / 100
    End If
    
End Sub

Private Sub spnLineWeight_SpinUp()
    
    If txtLineWeight = "" Then
        txtLineWeight = 1
    Else
        txtLineWeight = GetQuaterUnit(txtLineWeight.value, True)
    End If

End Sub

Private Sub spnLineWeight_SpinDown()
    
    If txtLineWeight <= 0 Then
        txtLineWeight = 0
    ElseIf txtLineWeight = "" Then
        txtLineWeight = 1
    Else
        txtLineWeight = GetQuaterUnit(txtLineWeight.value, False)
    End If

End Sub

Private Function GetQuaterUnit(ByVal value As Single, ByVal getUpper As Boolean) As Double

    Dim dec As Single
    Dim ret As Single
    
    dec = value - Int(value)
    If getUpper Then
        Select Case True
            Case dec < 0.25
                ret = Int(value) + 0.25
            Case dec < 0.5
                ret = Int(value) + 0.5
            Case dec < 0.75
                ret = Int(value) + 0.75
            Case Else
                ret = Int(value) + 1
        End Select
    Else
        Select Case True
            Case dec = 0
                ret = Int(value) - 0.25
            Case dec <= 0.25
                ret = Int(value)
            Case dec <= 0.5
                ret = Int(value) + 0.25
            Case dec <= 0.75
                ret = Int(value) + 0.5
            Case Else
                ret = Int(value) + 0.75
        End Select
    End If
    GetQuaterUnit = ret

End Function

' 色関係
Private Sub txtStartHue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtStartHue, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtEndHue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtEndHue, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtN_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtN, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSat_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtSat, True) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtVal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtVal, True) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtStartHue_AfterUpdate()

    If (Not IsNumeric(txtStartHue)) _
    Or val(txtStartHue) > 360 _
    Or val(txtStartHue) < 0 Then
        txtStartHue = ""
    Else
        lblStartHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtStartHue)(0)
    End If

End Sub

Private Sub txtEndHue_AfterUpdate()
    
    If (Not IsNumeric(txtEndHue)) _
    Or val(txtEndHue) > 360 _
    Or val(txtEndHue) < 0 Then
        txtEndHue = ""
    Else
        lblEndHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtEndHue)(0)
    End If

End Sub

Private Sub txtN_AfterUpdate()

    If (Not IsNumeric(txtN)) _
    Or Int(val(txtN)) <> val(txtN) _
    Or val(txtN) <= 0 Then
        txtN = ""
    End If

End Sub

Private Sub txtSat_AfterUpdate()

    If (Not IsNumeric(txtSat)) _
    Or val(txtSat) > 1 _
    Or val(txtSat) < 0 Then
        txtSat = ""
    End If
    
End Sub

Private Sub txtVal_AfterUpdate()

    If (Not IsNumeric(txtVal)) _
    Or val(txtVal) > 1 _
    Or val(txtVal) < 0 Then
        txtVal = ""
    End If
    
End Sub

' カラーバー位置
Private Sub optColorBarTop_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelPosition = lpTop

End Sub

Private Sub optColorBarBottom_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelPosition = lpBottom

End Sub

Private Sub optColorBarLeft_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelPosition = lpLeft

End Sub

Private Sub optColorBarRight_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelPosition = lpRight

End Sub

' カラーバーサイズ
Private Sub optLengthAuto_AfterUpdate()

    chargeDischargeChart.ColorBarLengthAuto = True
    SetLengthEnabled

End Sub

Private Sub optLengthCustom_AfterUpdate()

    chargeDischargeChart.ColorBarLengthAuto = False
    SetLengthEnabled

End Sub

Private Sub txtLength_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtLength, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLength_AfterUpdate()

    If (Not IsNumeric(txtLength)) _
    Or Int(val(txtLength)) <> val(txtLength) _
    Or val(txtLength) <= 0 Then
        txtLength = ""
        chargeDischargeChart.ColorBar.Length = 1
    Else
        chargeDischargeChart.ColorBar.Length = txtLength.value
    End If

End Sub

Private Sub SetLengthEnabled()

    If optLengthAuto = True Then
        txtLength.Enabled = False
    Else
        txtLength.Enabled = True
    End If

End Sub

Private Sub optWidthAuto_AfterUpdate()

    chargeDischargeChart.ColorBarWidthAuto = True
    SetWidthEnabled

End Sub

Private Sub optWidthCustom_AfterUpdate()

    chargeDischargeChart.ColorBarWidthAuto = False
    SetWidthEnabled

End Sub

Private Sub txtWidth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtWidth, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtWidth_AfterUpdate()

    If (Not IsNumeric(txtWidth)) _
    Or Int(val(txtWidth)) <> val(txtWidth) _
    Or val(txtWidth) <= 0 Then
        txtWidth = ""
        chargeDischargeChart.ColorBar.Width = 1
    Else
        chargeDischargeChart.ColorBar.Width = txtWidth.value
    End If

End Sub

Private Sub SetWidthEnabled()

    If optWidthAuto = True Then
        txtWidth.Enabled = False
    Else
        txtWidth.Enabled = True
    End If

End Sub

' カラーバーラベル形式
Private Sub optCardinal_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelFormat = lfCardinal

End Sub

Private Sub optOrdinal_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelFormat = lfOrdinal

End Sub

' カラーバーラベル表示間隔
Private Sub optColorBarUnitMaxMin_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelMode = lmMinMax
    SetColorBarUnitEnabled

End Sub

Private Sub optColorBarUnitAll_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelMode = lmAll
    SetColorBarUnitEnabled

End Sub

Private Sub optColorBarUnitCustom_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelMode = lmCustom
    SetColorBarUnitEnabled

End Sub

Private Sub SetColorBarUnitEnabled()

    If optColorBarUnitCustom = True Then
        txtColorBarUnit.Enabled = True
    Else
        txtColorBarUnit.Enabled = False
    End If

End Sub

Private Sub txtColorBarUnit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtColorBarUnit, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtColorBarUnit_AfterUpdate()

    If (Not IsNumeric(txtColorBarUnit)) _
    Or Int(val(txtColorBarUnit)) <> val(txtColorBarUnit) _
    Or val(txtColorBarUnit) <= 0 Then
        txtColorBarUnit = ""
        chargeDischargeChart.ColorBar.LabelUnit = 1
    Else
        chargeDischargeChart.ColorBar.LabelUnit = txtColorBarUnit.value
    End If

End Sub

' カラフルラベル
Private Sub chkColorfulLabel_AfterUpdate()

    chargeDischargeChart.ColorBar.IsColorfulLabel = chkColorfulLabel.value

End Sub

Private Sub btnExecute_Click()
    
    ' 開始セルチェック
    Dim startCells() As Variant
    Dim tmpRow As Integer
    Dim isSameRow As Boolean
    Dim i As Long
    
    tmpRow = 0
    isSameRow = True
    With chargeDischargeChart
        startCells = Array(.CycleStartCell, .StepStartCell, .AHStartCell, .VStartCell)
    End With
    For i = LBound(startCells) To UBound(startCells)
        ' 存在確認
        If startCells(i) Is Nothing Then
            MsgBox "開始セルを全て入力して下さい。", vbExclamation
            Exit Sub
        End If
        ' 行数違い確認
        If tmpRow <> 0 Then
            If tmpRow <> startCells(i).Row Then
                isSameRow = False
            End If
        End If
        tmpRow = startCells(i).Row
    Next
    
    ' サイクル数チェック
    If chargeDischargeChart.cycleCount <= 0 Then
        MsgBox "サイクル数を入力して下さい。", vbExclamation
        Exit Sub
    End If
    
    ' 256系列に収まるように（ダミー1系列、1Cycle2系列）
    If chkUseDisplayCycleAdjust Then
        If txtDisplayCycleCount > 127 Then
            MsgBox "グラフ系列の最大数を超えます。" & _
                   "表示サイクル数が127以下になるようにしてください。" & vbCrLf & vbCrLf & _
                   "表示サイクル数=必須表示+(サイクル数-必須表示)/表示間隔" & vbCrLf & _
                   "(小数点以下切捨て)", vbExclamation
            Exit Sub
        End If
    Else
        If txtCycle > 127 Then
            MsgBox "グラフ系列の最大数を超えます。" & _
                   "サイクル数を127以下にするか、表示サイクル数調整を使って系列を減らしてください。", _
                   vbExclamation
            Exit Sub
        End If
    End If
    
    ' N >= サイクル数を保証
    If IsNumeric(txtN) Then
        If val(txtN) < val(txtCycle) Then
            MsgBox "色の分割数はサイクル数以上にしてください。", vbExclamation
            Exit Sub
        End If
    End If
    
    ' 行数違い確認
    If Not isSameRow Then
        If MsgBox("全ての開始セルが同じ行ではありませんが、よろしいですか？" & vbCrLf & _
                  "※系列のデータ範囲は、STEPの開始セルの行数を基準に決定します。", _
                  vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' Chart作成実行
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Error_Handler
    
    chargeDischargeChart.CreateChart
    If chargeDischargeChart.targetChart.SeriesCollection.Count <= 0 Then
        chargeDischargeChart.DeleteChart
        MsgBox "データが正しくセットされませんでした。グラフ作成が失敗しました。", vbExclamation
        GoTo Exit_Proc
    End If
    
    ' フォーマットクラス生成
    Set chartFormat_ = New clsScatterChart
    chartFormat_.SetDefault chargeDischargeChart.targetChart

    ' グラデーション作成～各系列にセット
    Dim gradColors() As Long
    
    With chartFormat_.Series
        gradColors = colorPalette.MakeHueGradientColors( _
            IIf(txtN = "", chargeDischargeChart.MaxCycle, txtN.value), _
            startHue:=IIf(txtStartHue = "", 300, txtStartHue.value), _
            endHue:=IIf(txtEndHue = "", 0, txtEndHue.value), _
            sat:=IIf(txtSat = "", 1, txtSat.value), _
            val:=IIf(txtVal = "", 0.85, txtVal.value), _
            shortestPath:=chkShortest.value _
        )
        For i = 1 To .Count
            .Item(i).LineColor = gradColors(chargeDischargeChart.ChargeDischargeSeries(.Count - i).CycleIndex - 1)
        Next
    End With

    ' NiceScale設定
    With chartFormat_
        .IsAutoScaleX = False
        .IsAutoScaleY = False
        .marginModeX = ammFitBoth
        .marginModeY = ammFitBoth
    End With
    
    ' 系列書式設定
    With chartFormat_.Series
        .ChartType = Line_
        .IsSmooth = True
        .LineWeight = IIf(txtLineWeight = "", 1, txtLineWeight)
    End With
    
     ' 軸の設定
    With chargeDischargeChart.targetChart
        .Axes(xlCategory, xlPrimary).TickLabels.NumberFormatLocal = NUMBER_FORMAT_ZERO_OR_DECIMAL
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Capacity"
        .Axes(xlValue, xlPrimary).TickLabels.NumberFormatLocal = NUMBER_FORMAT_ZERO_OR_DECIMAL
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Cell voltage"
    End With
    
    ' フォーマット適用
    chartFormat_.ApplyTo chargeDischargeChart.targetChart
    
    ' カラーバー作成
    chargeDischargeChart.DrawColorBar gradColors
    
    ' 完了
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    On Error GoTo 0
    
    MsgBox "完了しました。", vbInformation
    Unload Me
    Exit Sub

Exit_Proc:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

Error_Handler:
    MsgBox "エラー" & Err.Number & ":" & Err.Description
    GoTo Exit_Proc

End Sub

Private Sub btnCancel_Click()

    Unload Me

End Sub
