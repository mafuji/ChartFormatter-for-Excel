VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChargeDischargeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' 充放電グラフの系列
'   ＜目的＞
'　・系列インデックスとサイクルインデックスの紐づけ
'　・最終サイクル→1stサイクルの順に系列追加することで若いサイクルを手前に出す
'---------------------------------------------------------------

Private cycleIndex_ As Long
Private seriesName_ As String
Private seriesXVlue_ As range
Private seriesVlue_ As range

Public Property Get CycleIndex() As Long
    CycleIndex = cycleIndex_
End Property

Public Property Let CycleIndex(ByVal newCycleIndex As Long)
    cycleIndex_ = newCycleIndex
End Property


Public Property Get SeriesName() As String
    SeriesName = seriesName_
End Property

Public Property Let SeriesName(ByVal newSeriesName As String)
    seriesName_ = newSeriesName
End Property


Public Property Get SeriesXValue() As range
    Set SeriesXValue = seriesXVlue_
End Property

Public Property Set SeriesXValue(ByVal newSeriesXValue As range)
    Set seriesXVlue_ = newSeriesXValue
End Property


Public Property Get SeriesValue() As range
    Set SeriesValue = seriesVlue_
End Property

Public Property Set SeriesValue(ByVal newSeriesValue As range)
    Set seriesVlue_ = newSeriesValue
End Property
