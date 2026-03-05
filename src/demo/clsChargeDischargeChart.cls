VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChargeDischargeChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================
' 充放電グラフ
'===============================================================

'---------------------------------------------------------------
' AH vs. V グラフ
'---------------------------------------------------------------
Private targetChart_ As Chart

' 各項目の開始セル
Private cycleStartCell_ As range
Private stepStartCell_ As range
Private vStartCell_ As range
Private ahStartCell_ As range

Private startStep_ As enChargeStatus ' 開始ステップ:1,2,3,4
Private cycleCount_ As Long ' 最大サイクル数（指定）
Private useDisplayCycleAdjust_ As Boolean ' 表示サイクル調整機能使用フラグ
Private alwaysDisplayCycle_ As Long ' 必須表示のサイクル1～xまで必須表示
Private displayCycleInterval_ As Long ' サイクル数の表示間隔:2なら2,4,6,8,...を系列として表示

Private colorBar_ As New clsColorBar
Private colorBarWidthAuto_ As Boolean
Private colorBarLengthAuto_ As Boolean
Private maxCycle_ As Long ' 最大サイクル数（ループ結果）

Private chargeDischargeSeries_() As New clsChargeDischargeSeries ' 系列情報のバッファ（逆順作成により若いサイクルを前面表示するために使用）

'---------------------------------------------------------------
' CYCLE vs. AH グラフ用
'---------------------------------------------------------------
' 各ステップ最終行データ格納用配列
'   0:サイクル数
'   1:charge[AH]
'   2:discharge[AH]
'   3:(charge[AH]/discharge[AH])*100
Private dataBuffer_() As Variant

Public Property Get targetChart() As Chart
    Set targetChart = targetChart_
End Property

Public Property Set targetChart(ByVal newTargetChart As Chart)
    Set targetChart_ = newTargetChart
End Property


Public Property Get CycleStartCell() As range
    Set CycleStartCell = cycleStartCell_
End Property

Public Property Set CycleStartCell(ByVal newCycleStartCell As range)
    Set cycleStartCell_ = newCycleStartCell
End Property


Public Property Get StepStartCell() As range
    Set StepStartCell = stepStartCell_
End Property

Public Property Set StepStartCell(ByVal newStepStartCell As range)
    Set stepStartCell_ = newStepStartCell
End Property


Public Property Get VStartCell() As range
    Set VStartCell = vStartCell_
End Property

Public Property Set VStartCell(ByVal newVStartCell As range)
    Set vStartCell_ = newVStartCell
End Property


Public Property Get AHStartCell() As range
    Set AHStartCell = ahStartCell_
End Property

Public Property Set AHStartCell(ByVal newAHStartCell As range)
    Set ahStartCell_ = newAHStartCell
End Property

Public Property Get StartStep() As enChargeStatus
    StartStep = startStep_
End Property

Public Property Let StartStep(ByVal newStartStep As enChargeStatus)
    startStep_ = newStartStep
End Property


Public Property Get cycleCount() As Long
    cycleCount = cycleCount_
End Property

Public Property Let cycleCount(ByVal newCycleCount As Long)
    cycleCount_ = newCycleCount
End Property


Public Property Get UseDisplayCycleAdjust() As Boolean
    UseDisplayCycleAdjust = useDisplayCycleAdjust_
End Property

Public Property Let UseDisplayCycleAdjust(ByVal newUseDisplayCycleAdjust As Boolean)
    useDisplayCycleAdjust_ = newUseDisplayCycleAdjust
End Property

Public Property Get alwaysDisplayCycle() As Long
    alwaysDisplayCycle = alwaysDisplayCycle_
End Property

Public Property Let alwaysDisplayCycle(ByVal newAlwaysDisplayCycle As Long)
    alwaysDisplayCycle_ = newAlwaysDisplayCycle
End Property

Public Property Get displayCycleInterval() As Long
    displayCycleInterval = displayCycleInterval_
End Property

Public Property Let displayCycleInterval(ByVal newDisplayCycleInterval As Long)
    displayCycleInterval_ = newDisplayCycleInterval
End Property


Public Property Get ColorBar() As clsColorBar
    Set ColorBar = colorBar_
End Property

Public Property Set ColorBar(ByVal newColorBar As clsColorBar)
    Set colorBar_ = newColorBar
End Property

Public Property Get ColorBarWidthAuto() As Boolean
    ColorBarWidthAuto = colorBarWidthAuto_
End Property

Public Property Let ColorBarWidthAuto(newColorBarWidthAuto As Boolean)
    colorBarWidthAuto_ = newColorBarWidthAuto
End Property

Public Property Get ColorBarLengthAuto() As Boolean
    ColorBarLengthAuto = colorBarLengthAuto_
End Property

Public Property Let ColorBarLengthAuto(newcolorBarLengthAuto As Boolean)
    colorBarLengthAuto_ = newcolorBarLengthAuto
End Property

Public Property Get MaxCycle() As Long
    MaxCycle = maxCycle_
End Property

Public Property Let MaxCycle(ByVal newMaxCycle As Long)
    maxCycle_ = newMaxCycle
End Property

Public Property Get ChargeDischargeSeries(index As Long) As clsChargeDischargeSeries
    Set ChargeDischargeSeries = chargeDischargeSeries_(index)
End Property

Public Property Get DataBuffer() As Variant
    DataBuffer = dataBuffer_
End Property

Public Sub DeleteChart()

    Me.targetChart.Parent.Delete

End Sub

' グラフに系列を追加する
Public Sub CreateChart()

    Set targetChart_ = ActiveSheet.ChartObjects.Add(100, 100, 100, 100).Chart
    targetChart_.ChartType = xlXYScatter

    Dim ws As Worksheet
    Dim rStep As range, rV As range, rAH As range
    Dim currentStep As Long
    Dim targetParity As Long
    Dim CycleIndex As Long
    Dim isCharge As Boolean
    Dim SeriesName As String
    Dim seriesCount As Long
    Dim tmpData(1 To 2) As Variant
    Dim dicDataBuffer As Object
    
    Set ws = Me.StepStartCell.Worksheet
    Set rStep = Me.StepStartCell
    Set rV = Me.VStartCell
    Set rAH = Me.AHStartCell
    Set dicDataBuffer = CreateObject("Scripting.Dictionary")

    ' 開始ステップの奇偶を採用
    currentStep = rStep.value
    targetParity = currentStep Mod 2 ' 0:偶,1:奇

    ' charge/discharge の開始
    isCharge = (Me.StartStep = Charge)

    seriesCount = 0
    CycleIndex = 1

    Do While CycleIndex <= Me.cycleCount And rStep <> ""
            
        '=== 対象ステップの連続ブロックを抽出 ===
        Dim firstRow As Long, lastRow As Long
        firstRow = rStep.Row

        ' 対象 parity のステップが来るまで下へ
        Do While rStep.value Mod 2 <> targetParity
            Set rStep = rStep.Offset(1, 0)
            Set rV = rV.Offset(1, 0)
            Set rAH = rAH.Offset(1, 0)
            If IsEmpty(rStep) Then Exit Sub
        Loop

        ' ブロック開始
        firstRow = rStep.Row

        ' 同じステップが続く限り下へ
        Dim thisStep As Long
        thisStep = rStep.value

        Do While rStep.value = thisStep
            Set rStep = rStep.Offset(1, 0)
            Set rV = rV.Offset(1, 0)
            Set rAH = rAH.Offset(1, 0)
            If IsEmpty(rStep) Then Exit Do
        Loop

        lastRow = rStep.Row - 1
    
        '=== グラフ①の系列として情報登録 ===
        
        ' 表示サイクル調整を使う場合：サイクル1～xは確定、x+1以降は等間隔で飛び地採用
        If UseDisplayCycleAdjust = False _
        Or CycleIndex <= alwaysDisplayCycle _
        Or Int((CycleIndex - alwaysDisplayCycle) / displayCycleInterval) = (CycleIndex - alwaysDisplayCycle) / displayCycleInterval Then
            ReDim Preserve chargeDischargeSeries_(seriesCount)
            
            If isCharge Then
                SeriesName = ToOrdinal(CycleIndex) & " charge"
            Else
                SeriesName = ToOrdinal(CycleIndex) & " discharge"
            End If
            
            With chargeDischargeSeries_(seriesCount)
                .CycleIndex = CycleIndex
                .SeriesName = SeriesName
                .SeriesXValue = ws.range(ws.Cells(firstRow, Me.AHStartCell.Column), ws.Cells(lastRow, Me.AHStartCell.Column))
                .SeriesValue = ws.range(ws.Cells(firstRow, Me.VStartCell.Column), ws.Cells(lastRow, Me.VStartCell.Column))
            End With
            
            seriesCount = seriesCount + 1
        End If
        
        ' プール用データ
        If isCharge Then
            tmpData(1) = ws.Cells(lastRow, AHStartCell.Column).value
        Else
            tmpData(2) = ws.Cells(lastRow, AHStartCell.Column).value
        End If
        
        ' charge/discharge 切り替え
        isCharge = Not isCharge

        ' 2step終わったらcycle++
        If isCharge = (Me.StartStep = Charge) Then
            dicDataBuffer.Add CycleIndex, tmpData ' 各ステップ最終行データをプール(Key:サイクル数)
            CycleIndex = CycleIndex + 1
        End If
    Loop
    maxCycle_ = CycleIndex - 1
    
    ' 各ステップ最終行データプールを出力用に配列に格納
    Dim i As Integer
    ReDim dataBuffer_(0 To maxCycle_, 0 To 3)
    
    ' 見出し
    dataBuffer_(0, 0) = "CYCLE"
    dataBuffer_(0, 1) = "AH(charge)"
    dataBuffer_(0, 2) = "AH(discharge)"
    dataBuffer_(0, 3) = "AH(discharge)/AH(charge)[%]"
    
    ' データ
    For i = 1 To maxCycle_
        dataBuffer_(i, 0) = i ' CYCLE
        dataBuffer_(i, 1) = dicDataBuffer(i)(1) ' AH(charge)
        dataBuffer_(i, 2) = dicDataBuffer(i)(2) ' AH(discharge)
        If dicDataBuffer(i)(1) = 0 Then
            dataBuffer_(i, 3) = ""
        Else
            dataBuffer_(i, 3) = dicDataBuffer(i)(2) / dicDataBuffer(i)(1) * 100 ' AH(discharge)/AH(charge)[%]
        End If
    Next
    
    ' 最終サイクル→1stの順に系列を作成
    Dim s As Series
    
    For i = UBound(chargeDischargeSeries_) To LBound(chargeDischargeSeries_) Step -1
        Set s = Me.targetChart.SeriesCollection.newSeries
        s.Name = chargeDischargeSeries_(i).SeriesName
        s.XValues = chargeDischargeSeries_(i).SeriesXValue
        s.values = chargeDischargeSeries_(i).SeriesValue
    Next

    ' 出力シート作成
    OutputDataSheet

End Sub

Public Sub DrawColorBar(ByRef colors() As Long)

    '-------------------------------------------------
    ' AutoLength用
    '
    ' 1つの色の長さを決定する基準分割数（長さ方向のplotAreaサイズの2/3をこれで割った値を1色の長さにする）
    ' 分割数がこれを超えるときは、バー全体の長さ限界を長さ方向のplotAreaサイズにする
    ' 即ち、
    '  colorsの要素数 <= BaseDevisionCount：全体の長さは1色の長さ×要素数
    '  colorsの要素数 > BaseDevisionCount：全体の長さはplotAreaの長さの2/3に固定、1色の長さが小さくなる
    Const BaseDevisionCount As Long = 15
    '-------------------------------------------------
    
    Dim leftPos As Double
    Dim topPos As Double
    Dim co As ChartObject
    
    Set co = targetChart.Parent
    
    ' サイズ自動設定（自動でない場合は任意数値セットされている前提）
    If colorBarWidthAuto_ Or ColorBar.Width <= 0 Then
        With targetChart.PlotArea
            If ColorBar.LabelPosition = lpLeft Or ColorBar.LabelPosition = lpRight Then
                ' vertical
                ColorBar.Width = .InsideWidth / 30
            Else
                ' horizontal
                ColorBar.Width = .InsideHeight / 30
            End If
        End With
    End If
    If colorBarLengthAuto_ Or ColorBar.Length <= 0 Then
        With targetChart.PlotArea
            If ColorBar.LabelPosition = lpLeft Or ColorBar.LabelPosition = lpRight Then
                ' vertical
                ColorBar.Length = IIf(UBound(colors) + 1 <= BaseDevisionCount, _
                                  .InsideHeight * (2 / 3) / BaseDevisionCount * (UBound(colors) + 1), _
                                  .InsideHeight * (2 / 3))
            Else
                ' horizontal
                ColorBar.Length = IIf(UBound(colors) + 1 <= BaseDevisionCount, _
                                  .InsideWidth * (2 / 3) / BaseDevisionCount * (UBound(colors) + 1), _
                                  .InsideWidth * (2 / 3))
            End If
        End With
    End If
    
    ' 位置
    With targetChart.PlotArea
        Select Case ColorBar.LabelPosition
            Case lpLeft
                leftPos = IIf(co.Left - ColorBar.Width <= 0, 1, co.Left - ColorBar.Width)
                topPos = co.Top + .InsideTop
            Case lpRight
                leftPos = co.Left + co.Width - 50
                topPos = co.Top + .InsideTop
            Case lpTop
                leftPos = co.Left + .InsideLeft
                topPos = IIf(co.Top - ColorBar.Width + 50 <= 0, 1, co.Top - ColorBar.Width + 50)
            Case lpBottom
                leftPos = co.Left + .InsideLeft
                topPos = co.Top + co.Height
        End Select
    End With

    ' draw処理（他のプロパティは適宜設定）
    ColorBar.Initialize colors, ActiveSheet, leftPos, topPos
    ColorBar.Draw

    ' グラフとグループ化
'    Dim grp As Shape
'    Set grp = ActiveSheet.Shapes.range(Array(ColorBar.GroupName, co.ShapeRange(1).Name)).Group

End Sub

' Cycle vs. AH 表/グラフ出力
Private Sub OutputDataSheet()

    Dim aws As Worksheet
    Set aws = ActiveSheet

    ' output用シート作成
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=ActiveSheet)
    ws.Name = UniqueSheetName(ActiveWorkbook, "output")
    
    ' 表出力
    Const startRow As Long = 1
    Const startColumn As Long = 1
    Dim outputRange As range
    Set outputRange = ws.Cells(startRow, startColumn) _
                        .Resize(UBound(dataBuffer_, 1) + 1, UBound(dataBuffer_, 2) + 1)
    outputRange.value = dataBuffer_
    
    ' 表整形
    With outputRange
        .Columns.AutoFit
        .Rows(startRow).Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = vbBlack
    End With
    
    ' グラフ作成
    Dim newChart As Chart
    Set newChart = ws.ChartObjects.Add(100, 100, 100, 100).Chart
    newChart.ChartType = xlXYScatter
    
    ' CYCLE vs. AH(discharge) 系列追加
    Dim s As Series
    Set s = newChart.SeriesCollection.newSeries
    s.Name = "Discharge capacity"
    s.XValues = outputRange.Columns(1).Offset(1).Resize(outputRange.Rows.Count - 1, 1)
    s.values = outputRange.Columns(3).Offset(1).Resize(outputRange.Rows.Count - 1, 1)
    s.AxisGroup = xlPrimary
    
    ' CYCLE vs. charge/discharge 系列追加
    Set s = newChart.SeriesCollection.newSeries
    s.Name = "Coulombic efficiency"
    s.XValues = outputRange.Columns(1).Offset(1).Resize(outputRange.Rows.Count - 1, 1)
    s.values = outputRange.Columns(4).Offset(1).Resize(outputRange.Rows.Count - 1, 1)
    s.AxisGroup = xlSecondary
    
    ' グラフ整形
    Dim chartFormat As New clsScatterChart
    chartFormat.SetDefault newChart
    chartFormat.HasLegend = True
    chartFormat.ApplyTo newChart
    
    With newChart
        .Axes(xlCategory, xlPrimary).TickLabels.NumberFormatLocal = "#,##0"
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Cycle [-]"
        .Axes(xlValue, xlPrimary).TickLabels.NumberFormatLocal = NUMBER_FORMAT_ZERO_OR_DECIMAL
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Discharge capacity"
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormatLocal = "#,##0"
        .Axes(xlValue, xlSecondary).AxisTitle.Text = "Coulombic efficiency [%]"
        .Axes(xlValue, xlSecondary).MaximumScale = 100
        .Axes(xlValue, xlSecondary).MinimumScale = 0
        .Axes(xlValue, xlSecondary).MajorUnit = 10
    End With

    aws.Select

End Sub
