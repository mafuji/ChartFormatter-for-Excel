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

' 散布図系列の設定
Private index_ As Long
Private name_ As String
Private dataColumn_ As String ' データ列
Private markerColor_ As Long
Private markerSize_ As Integer
Private useSecondAxis_ As Boolean

' --- index ---
Public Property Get index() As Long
    index = index_
End Property

Public Property Let index(ByVal newIndex As Long)
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

' --- useSecondAxis ---
Public Property Get UseSecondAxis() As Boolean
    UseSecondAxis = useSecondAxis_
End Property

Public Property Let UseSecondAxis(ByVal newUseSecondAxis As Boolean)
    useSecondAxis_ = newUseSecondAxis
End Property

Public Function Clone() As clsScatterSeries
    
    Dim c As New clsScatterSeries

    c.Name = Me.Name
    c.MarkerColor = Me.MarkerColor
    c.MarkerSize = Me.MarkerSize
    c.UseSecondAxis = Me.UseSecondAxis

    Set Clone = c
    
End Function

' Seriesから設定をコピー
Public Sub LoadFromSeries(ByVal seriesIndex As Integer, ByVal targetSeries As Series)

    Me.index = seriesIndex
    Me.Name = targetSeries.Name
    Me.MarkerColor = targetSeries.MarkerBackgroundColor
    Me.MarkerSize = targetSeries.MarkerSize
    Me.UseSecondAxis = (targetSeries.AxisGroup = xlSecondary)

    Dim rng As Range
    
    Set rng = GetRangeFromSeries(targetSeries, 2)
    If Not rng Is Nothing Then
        Me.DataColumn = GetColumnLetter(rng)
    End If

End Sub

Public Sub ApplyTo(ByRef targetSeries As Series)

    With targetSeries
    
        ' 現状固定設定
        .MarkerStyle = xlMarkerStyleCircle
    
        ' プロパティ設定
        .Name = Me.Name
        .MarkerBackgroundColor = Me.MarkerColor
        .MarkerForegroundColor = Me.MarkerColor
        .MarkerSize = Me.MarkerSize
        .AxisGroup = IIf(Me.UseSecondAxis, xlSecondary, xlPrimary)
    End With

End Sub
