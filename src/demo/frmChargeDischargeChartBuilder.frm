VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChargeDischargeChartBuilder 
   Caption         =   "充放電グラフ作成"
   ClientHeight    =   8328.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9408.001
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


' カラーバーラベル形式
Private Sub optCardinal_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelFormat = lfCardinal

End Sub

Private Sub optOrdinal_AfterUpdate()

    chargeDischargeChart.ColorBar.LabelFormat = lfOrdinal

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



Private Sub UserForm_Initialize()

    ' データ範囲の既定設定
    txtCycleColumn = DEFAULT_CYCLE_COLUMN
    txtStepColumn = DEFAULT_STEP_COLUMN
    txtAHColumn = DEFAULT_AH_COLUMN
    txtVColumn = DEFAULT_V_COLUMN
    txtDataStartRow = DEFAULT_DATA_START_ROW
    cmbStartStep.List = Array(1, 2, 3, 4)
    cmbStartStep.ListIndex = 1

    ' 開始ステップのChargeTypeを設定
    chargeDischargeChart.StartStep = IIf(optCharge, Charge, Discharge)
    
    ' 色の既定設定
    txtStartHue = 0
    txtEndHue = 300
    txtSat = 1
    txtVal = 0.85
    lblStartHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtStartHue)(0)
    lblEndHue.ForeColor = colorPalette.MakeHueGradientColors(1, txtEndHue)(0)
    
    ' カラーバーの既定設定
    With chargeDischargeChart.ColorBar
        optColorBarRight = True: .LabelPosition = lpRight
        optCardinal = True: .LabelFormat = lfCardinal
        optColorBarUnitCustom = True: .LabelMode = lmCustom
        SetColorBarUnitEnabled
        txtColorBarUnit = 10: .LabelUnit = 10
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
        MsgBox "各列番号,データ開始行を全て入力してください。", vbExclamation
        Exit Sub
    End If

    ' 開始セルを探して、フォームとチャートインスタンスにセット
    Dim rng As range
    
    Set rng = range(txtStepColumn & txtDataStartRow)
    Do Until val(rng) = val(cmbStartStep.value) Or rng = ""
        Set rng = rng.offset(1)
    Loop
    If rng = "" Then
        MsgBox "開始ステップが見つかりませんでした。", vbExclamation
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

Private Sub txtCycle_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtCycle, False) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCycle_AfterUpdate()

    chargeDischargeChart.CycleCount = txtCycle.value
    
End Sub

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
    If chargeDischargeChart.CycleCount <= 0 Then
        MsgBox "サイクル数を入力して下さい。", vbExclamation
        Exit Sub
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
        If MsgBox("全ての開始セルが同じ行ではありませんが、よろしいですか？", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' Chart作成実行
    Application.Cursor = xlWait
    
    chargeDischargeChart.CreateChart
    If chargeDischargeChart.targetChart.SeriesCollection.Count <= 0 Then
        chargeDischargeChart.DeleteChart
        MsgBox "データが正しくセットされませんでした。グラフ作成が失敗しました。", vbExclamation
        GoTo Exit_Proc
    End If
    
    ' フォーマットクラス生成
    Set chartFormat_ = New clsScatterChart
    chartFormat_.SetDefault chargeDischargeChart.targetChart

    ' 書式設定
    With chartFormat_
        .Series.ChartType = Line_
        .Series.IsSmooth = True
    End With

    ' グラデーション作成～各系列にセット
    Dim gradColors() As Long
    
    With chartFormat_.Series
        gradColors = colorPalette.MakeHueGradientColors( _
            IIf(txtN = "", .Count / 2, txtN.value), _
            startHue:=IIf(txtStartHue = "", 300, txtStartHue.value), _
            endHue:=IIf(txtEndHue = "", 0, txtEndHue.value), _
            sat:=IIf(txtSat = "", 1, txtSat.value), _
            val:=IIf(txtVal = "", 0.85, txtVal.value), _
            shortestPath:=chkShortest.value _
        )
        For i = 1 To .Count
            .Item(i).LineColor = gradColors(Int((i + 1) / 2) - 1)
        Next
    End With

    ' 固有設定
    With chartFormat_
        .IsAutoScaleX = False
        .IsAutoScaleY = False
        .marginModeX = ammFitBoth
        .marginModeY = ammFitBoth
    End With
    
    ' フォーマット適用
    chartFormat_.ApplyTo chargeDischargeChart.targetChart
    
    ' カラーバー作成
    chargeDischargeChart.DrawColorBar gradColors
    
    ' 完了
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Unload Me
    
    Application.Cursor = xlDefault
    
    MsgBox "完了しました。", vbInformation

Exit_Proc:
    Application.Cursor = xlDefault

End Sub

Private Sub btnCancel_Click()

    Unload Me

End Sub

