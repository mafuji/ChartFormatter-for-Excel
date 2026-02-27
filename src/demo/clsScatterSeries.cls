VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScatterSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' 散布図系列の設定
'---------------------------------------------------------------

Private index_ As Long
Private name_ As String
Private dataColumn_ As String ' データ列
Private markerColor_ As Long
Private markerSize_ As Integer
Private lineColor_ As Long
Private lineWeight_ As Single
Private useSecondAxis_ As Boolean
Private chartType_ As enChartType
Private isSmooth_ As Boolean

' --- index ---
Public Property Get Index() As Long
    Index = index_
End Property

Public Property Let Index(ByVal newIndex As Long)
    index_ = newIndex
End Property

' --- name ---
Public Property Get Name() As String
    Name = name_
End Property

Public Property Let Name(ByVal newName As String)
    name_ = newName
End Property

' --- dataColumn ---
Public Property Get DataColumn() As String
    DataColumn = dataColumn_
End Property

Public Property Let DataColumn(ByVal newDataColumn As String)
    dataColumn_ = newDataColumn
End Property

' --- markerColor ---
Public Property Get MarkerColor() As Long
    MarkerColor = markerColor_
End Property

Public Property Let MarkerColor(ByVal newMarkerColor As Long)
    markerColor_ = newMarkerColor
End Property

' --- markerSize ---
Public Property Get MarkerSize() As Integer
    MarkerSize = markerSize_
End Property

Public Property Let MarkerSize(ByVal newMarkerSize As Integer)
    markerSize_ = newMarkerSize
End Property

' --- lineColor ---
Public Property Get LineColor() As Long
    LineColor = lineColor_
End Property

Public Property Let LineColor(ByVal newlineColor As Long)
    lineColor_ = newlineColor
End Property

' --- lineWeight ---
Public Property Get LineWeight() As Single
    LineWeight = lineWeight_
End Property

Public Property Let LineWeight(ByVal newLineWeight As Single)
    lineWeight_ = newLineWeight
End Property

' --- useSecondAxis ---
Public Property Get useSecondAxis() As Boolean
    useSecondAxis = useSecondAxis_
End Property

Public Property Let useSecondAxis(ByVal newUseSecondAxis As Boolean)
    useSecondAxis_ = newUseSecondAxis
End Property

' --- chartType ---
Public Property Get ChartType() As enChartType
    ChartType = chartType_
End Property

Public Property Let ChartType(ByVal newChartType As enChartType)
    chartType_ = newChartType
End Property

' --- isSmooth ---
Public Property Get IsSmooth() As Boolean
    IsSmooth = isSmooth_
End Property

Public Property Let IsSmooth(ByVal newIsSmooth As Boolean)
    isSmooth_ = newIsSmooth
End Property

Public Function Clone() As clsScatterSeries
    
    Dim c As New clsScatterSeries

    c.Name = Me.Name
    c.MarkerColor = Me.MarkerColor
    c.MarkerSize = Me.MarkerSize
    c.useSecondAxis = Me.useSecondAxis
    c.ChartType = Me.ChartType

    Set Clone = c
    
End Function

' Seriesから設定をコピー
Public Sub LoadFromSeries(ByVal seriesIndex As Integer, ByVal targetSeries As Series)

    Me.Index = seriesIndex
    Me.Name = targetSeries.Name
    
    ' マーカー
    Me.MarkerColor = targetSeries.MarkerBackgroundColor
    If targetSeries.MarkerStyle <> xlMarkerStyleNone Then
        Me.MarkerSize = targetSeries.MarkerSize
    Else
        Me.MarkerSize = DEFAULT_MARKER_SIZE
    End If
    
    ' 線
    Me.LineColor = targetSeries.Format.Line.ForeColor.RGB
    If targetSeries.Format.Line.Visible = msoTrue Then
        
        Me.LineWeight = targetSeries.Format.Line.Weight
    Else
        Me.LineWeight = DEFAULT_LINE_WEIGHT
    End If
    
    ' 2軸利用
    Me.useSecondAxis = (targetSeries.AxisGroup = xlSecondary)

    ' データ列
    Dim rng As range
    
    Set rng = GetRangeFromSeries(targetSeries, 2)
    If Not rng Is Nothing Then
        Me.DataColumn = GetColumnLetter(rng)
    End If

End Sub

Public Sub ApplyTo(ByRef targetSeries As Series)

    With targetSeries

        .Name = Me.Name
        Select Case Me.ChartType
            Case Scatter_
                .MarkerStyle = xlMarkerStyleCircle
                .Format.Line.Visible = msoFalse
                .MarkerBackgroundColor = Me.MarkerColor
                .MarkerForegroundColor = Me.MarkerColor
                .MarkerSize = Me.MarkerSize
            Case Line_
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.Visible = msoTrue
                .Format.Line.ForeColor.RGB = Me.LineColor
                .Format.Line.Weight = Me.LineWeight
                .Smooth = Me.IsSmooth
        End Select
        .AxisGroup = IIf(Me.useSecondAxis, xlSecondary, xlPrimary)
    End With

End Sub
