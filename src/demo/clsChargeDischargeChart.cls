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

Public Sub DeleteChart()

    Me.targetChart.Parent.Delete

End Sub

' グラフに系列を追加する
Public Sub CreateChart()

    Set targetChart_ = ActiveSheet.ChartObjects.Add(1, 1, 100, 100).Chart
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
    currentStep = rStep.Value
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
        Do While rStep.Value Mod 2 <> targetParity
            Set rStep = rStep.Offset(1, 0)
            Set rV = rV.Offset(1, 0)
            Set rAH = rAH.Offset(1, 0)
            If IsEmpty(rStep) Then Exit Sub
        Loop

        ' ブロック開始
        firstRow = rStep.Row

        ' 同じステップが続く限り下へ
        Dim thisStep As Long
        thisStep = rStep.Value

        Do While rStep.Value = thisStep
            Set rStep = rStep.Offset(1, 0)
            Set rV = rV.Offset(1, 0)
            Set rAH = rAH.Offset(1, 0)
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

' 数字を "1st", "2nd", "3rd", "4th" のような表記に変換
Private Function ToOrdinal(ByVal v As Variant) As String
    Dim n As Long
    Dim suffix As String
    Dim absN As Long
    Dim lastTwo As Long
    Dim lastOne As Long
    Dim sign As String
    
    ' 数値チェック
    If IsNull(v) Or Not IsNumeric(v) Then
        ToOrdinal = ""
        Exit Function
    End If
    
    ' 符号保持（-1 → "-1st" など）
    sign = IIf(CDbl(v) < 0, "-", "")
    
    ' 小数は整数部で判定（例：21.9 → 21st）
    n = Fix(CDbl(v))           ' Fixは0方向に丸め（-1.7→-1）
    absN = Abs(n)
    
    ' 11,12,13 は常に "th"
    lastTwo = absN Mod 100
    lastOne = absN Mod 10
    
    If lastTwo >= 11 And lastTwo <= 13 Then
        suffix = "th"
    Else
        Select Case lastOne
            Case 1: suffix = "st"
            Case 2: suffix = "nd"
            Case 3: suffix = "rd"
            Case Else: suffix = "th"
        End Select
    End If
    
    ToOrdinal = sign & CStr(absN) & suffix
End Function
