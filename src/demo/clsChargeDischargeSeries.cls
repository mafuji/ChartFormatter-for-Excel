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
Private seriesXVlue_ As Variant
Private seriesVlue_ As Variant

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


Public Property Get SeriesXValue() As Variant
    SeriesXValue = seriesXVlue_
End Property

Public Property Let SeriesXValue(ByVal newSeriesXValue As Variant)
    seriesXVlue_ = newSeriesXValue
End Property


Public Property Get SeriesValue() As Variant
    SeriesValue = seriesVlue_
End Property

Public Property Let SeriesValue(ByVal newSeriesValue As Variant)
    seriesVlue_ = newSeriesValue
End Property
