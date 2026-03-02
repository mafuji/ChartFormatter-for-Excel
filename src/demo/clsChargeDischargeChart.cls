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

'---------------------------------------------------------------
' 充放電グラフ
'---------------------------------------------------------------

Private targetChart_ As Chart
Private cycleStartCell_ As range
Private stepStartCell_ As range
Private vStartCell_ As range
Private ahStartCell_ As range
Private startStep_ As enChargeStatus
Private cycleCount_ As Long
Private colorBar_ As New clsColorBar

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


Public Property Get CycleCount() As Long
    CycleCount = cycleCount_
End Property

Public Property Let CycleCount(ByVal newCycleCount As Long)
    cycleCount_ = newCycleCount
End Property


Public Property Get ColorBar() As clsColorBar
    Set ColorBar = colorBar_
End Property

Public Property Set ColorBar(ByVal newColorBar As clsColorBar)
    Set colorBar_ = newColorBar
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
    Dim seriesCount As Long
    Dim cycleIndex As Long
    Dim isCharge As Boolean
    Dim seriesName As String

    Set ws = Me.StepStartCell.Worksheet
    Set rStep = Me.StepStartCell
    Set rV = Me.VStartCell
    Set rAH = Me.AHStartCell

    ' 開始ステップの奇偶を採用
    currentStep = rStep.value
    targetParity = currentStep Mod 2

    ' charge/discharge の開始
    isCharge = (Me.StartStep = Charge)

    seriesCount = 0
    cycleIndex = 1

    Do While cycleIndex <= Me.CycleCount And rStep <> ""

        '=== 対象ステップの連続ブロックを抽出 ===
        Dim firstRow As Long, lastRow As Long
        firstRow = rStep.Row

        ' 対象 parity のステップが来るまで下へ
        Do While rStep.value Mod 2 <> targetParity
            Set rStep = rStep.offset(1, 0)
            Set rV = rV.offset(1, 0)
            Set rAH = rAH.offset(1, 0)
            If IsEmpty(rStep) Then Exit Sub
        Loop

        ' ブロック開始
        firstRow = rStep.Row

        ' 同じステップが続く限り下へ
        Dim thisStep As Long
        thisStep = rStep.value

        Do While rStep.value = thisStep
            Set rStep = rStep.offset(1, 0)
            Set rV = rV.offset(1, 0)
            Set rAH = rAH.offset(1, 0)
            If IsEmpty(rStep) Then Exit Do
        Loop

        lastRow = rStep.Row - 1

        '=== 系列追加 ===
        seriesCount = seriesCount + 1

        If isCharge Then
            seriesName = ToOrdinal(cycleIndex) & " charge"
        Else
            seriesName = ToOrdinal(cycleIndex) & " discharge"
        End If

        Dim s As Series
        Set s = Me.targetChart.SeriesCollection.newSeries
        s.Name = seriesName
        s.XValues = ws.range(ws.Cells(firstRow, Me.AHStartCell.Column), ws.Cells(lastRow, Me.AHStartCell.Column))
        s.values = ws.range(ws.Cells(firstRow, Me.VStartCell.Column), ws.Cells(lastRow, Me.VStartCell.Column))

        ' charge/discharge 切り替え
        isCharge = Not isCharge

        ' 2step終わったらcycle++
        If isCharge = (Me.StartStep = Charge) Then
            cycleIndex = cycleIndex + 1
        End If
    Loop

End Sub

Public Sub DrawColorBar(ByRef colors() As Long)

    '-------------------------------------------------
    ' 1つの色の長さを決定する基準分割数（長さ方向のplotAreaサイズをこれで割った値を1色の長さにする）
    ' 分割数がこれを超えるときは、バー全体の長さ限界を長さ方向のplotAreaサイズにする
    ' 即ち、
    '  colorsの要素数 <= BaseDevisionCount：全体の長さは1色の長さ×要素数
    '  colorsの要素数 > BaseDevisionCount：全体の長さはplotAreaの長さに固定、1色の長さが小さくなる
    Const BaseDevisionCount As Long = 15
    '-------------------------------------------------
    
    Dim leftPos As Double
    Dim topPos As Double
    Dim width As Double
    Dim height As Double
    Dim co As ChartObject
    
    Set co = targetChart.Parent
    
    ' サイズ
    With targetChart.PlotArea
        If ColorBar.LabelPosition = lpLeft Or ColorBar.LabelPosition = lpRight Then
            ' vertical
            width = .InsideWidth / 30
            height = IIf(UBound(colors) + 1 <= BaseDevisionCount, _
                         .InsideHeight / BaseDevisionCount * (UBound(colors) + 1), _
                         .InsideHeight _
                     )
        Else
            ' horizontal
            width = IIf(UBound(colors) + 1 <= BaseDevisionCount, _
                         .InsideWidth / BaseDevisionCount * (UBound(colors) + 1), _
                         .InsideWidth _
                     )
            height = .InsideHeight / 30
        End If
    End With
        
    ' 位置
    With targetChart.PlotArea
        Select Case ColorBar.LabelPosition
            Case lpLeft
                leftPos = IIf(co.Left - width <= 0, 1, co.Left - width)
                topPos = co.Top + .InsideTop
            Case lpRight
                leftPos = co.Left + co.width - 50
                topPos = co.Top + .InsideTop
            Case lpTop
                leftPos = co.Left + .InsideLeft
                topPos = IIf(co.Top - height + 50 <= 0, 1, co.Top - height + 50)
            Case lpBottom
                leftPos = co.Left + .InsideLeft
                topPos = co.Top + co.height
        End Select
    End With

    ' draw処理（他のプロパティは適宜設定）
    ColorBar.Initialize colors, ActiveSheet, leftPos, topPos, width, height
    ColorBar.Draw

    ' グラフとグループ化
'    Dim grp As Shape
'    Set grp = ActiveSheet.Shapes.range(Array(ColorBar.GroupName, co.ShapeRange(1).Name)).Group

End Sub
