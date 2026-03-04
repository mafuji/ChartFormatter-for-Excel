VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScatterSeriesCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' 散布図系列コレクション
'---------------------------------------------------------------

Private items_ As Object
Private chartType_ As enChartType
Private isSmooth_ As Boolean

Private Sub Class_Initialize()
    Set items_ = CreateObject("Scripting.Dictionary")
End Sub

' --- chartType ---
Public Property Get ChartType() As enChartType
    ChartType = chartType_
End Property

Public Property Let ChartType(ByVal newChartType As enChartType)
    chartType_ = newChartType
    
    Dim i As Integer
    
    For i = 1 To Me.Count
        Item(i).ChartType = Me.ChartType
    Next
End Property

' --- isSmooth ---
Public Property Get IsSmooth() As Boolean
    IsSmooth = isSmooth_
End Property

Public Property Let IsSmooth(ByVal newIsSmooth As Boolean)
    isSmooth_ = newIsSmooth
    
    Dim i As Integer
    
    For i = 1 To Me.Count
        Item(i).IsSmooth = Me.IsSmooth
    Next
End Property

Public Sub Add(ByVal s As clsScatterSeries)
    items_.Add s.index, s
End Sub

Public Sub Remove(ByVal seriesIdx As Long)
    items_.Remove seriesIdx
End Sub

Public Property Get Item(ByVal seriesIdx As Long) As clsScatterSeries
    Set Item = items_(seriesIdx)
End Property

Public Property Get Count() As Long
    Count = items_.Count
End Property

Public Function Clone() As clsScatterSeriesCollection

    Dim c As New clsScatterSeriesCollection
    Dim i As Long
    For i = 1 To items_.Count
        c.Add items_(i).Clone
    Next i
    Set Clone = c

End Function

' SeriesCollectionから設定をコピー
Public Sub LoadFromSeriesCollection(ByVal targetSC As SeriesCollection)

    items_.RemoveAll
    
    Dim i As Integer
    Dim scatterSeries As clsScatterSeries
    
    For i = 1 To targetSC.Count
        If targetSC(i).Name <> DUMMY_SERIES_NAME Then
            Set scatterSeries = New clsScatterSeries
            scatterSeries.LoadFromSeries i, targetSC(i)
            items_.Add i, scatterSeries
        End If
    Next

End Sub

Public Sub ApplyTo(ByRef targetChart As Chart)

    Dim i As Integer
    
    For i = 1 To Me.Count
        Me.Item(i).ApplyTo targetChart.SeriesCollection(i)
    Next

End Sub
