VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScatterChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' 散布図/折れ線グラフの設定
'---------------------------------------------------------------

Private height_ As Double
Private width_ As Double
Private font_ As New clsFont
Private plotArea_ As New clsPlotArea
Private series_ As New clsScatterSeriesCollection
Private hasLegend_ As Boolean
Private isAutoScaleX_ As Boolean
Private isAutoScaleY_ As Boolean
Private marginModeX_ As AxisMarginMode
Private marginModeY_ As AxisMarginMode

Public Property Let Height(ByVal newHeight As Double)
    height_ = newHeight
End Property

Public Property Get Height() As Double
    Height = height_
End Property

Public Property Let Width(ByVal newWidth As Double)
    width_ = newWidth
End Property

Public Property Get Width() As Double
    Width = width_
End Property

Public Property Let HasLegend(ByVal newHasLegend As Boolean)
    hasLegend_ = newHasLegend
End Property

Public Property Get HasLegend() As Boolean
    HasLegend = hasLegend_
End Property

' --- font ---
Public Property Get Font() As clsFont
    Set Font = font_
End Property

Public Property Set Font(ByVal newFont As clsFont)
    Set font_ = newFont.Clone
End Property

' --- plotArea ---
Public Property Get PlotArea() As clsPlotArea
    Set PlotArea = plotArea_
End Property

Public Property Set PlotArea(ByVal newPlotArea As clsPlotArea)
    Set plotArea_ = newPlotArea.Clone
End Property

' --- series (コレクション) ---
Public Property Get Series() As clsScatterSeriesCollection
    Set Series = series_
End Property

Public Property Set Series(ByVal newSeries As clsScatterSeriesCollection)
    Set series_ = newSeries.Clone
End Property

' isAutoScale
Public Property Get IsAutoScaleX() As Boolean
    IsAutoScaleX = isAutoScaleX_
End Property

Public Property Let IsAutoScaleX(ByVal newIsAutoScaleX As Boolean)
    isAutoScaleX_ = newIsAutoScaleX
End Property

Public Property Get IsAutoScaleY() As Boolean
    IsAutoScaleY = isAutoScaleY_
End Property

Public Property Let IsAutoScaleY(ByVal newIsAutoScaleY As Boolean)
    isAutoScaleY_ = newIsAutoScaleY
End Property

' X軸 marginMode
Public Property Get marginModeX() As AxisMarginMode
    marginModeX = marginModeX_
End Property

Public Property Let marginModeX(ByVal newMarginModeX As AxisMarginMode)
    marginModeX_ = newMarginModeX
End Property

' Y軸 marginMode
Public Property Get marginModeY() As AxisMarginMode
    marginModeY = marginModeY_
End Property

Public Property Let marginModeY(ByVal newMarginModeY As AxisMarginMode)
    marginModeY_ = newMarginModeY
End Property

' 既定値の設定
Public Sub SetDefault(ByVal targetChart As Chart)

'    ' フォント
'    With Me.Font
'        .FontName = DEFAULT_FONT_NAME
'        .Size = DEFAULT_FONT_SIZE
'    End With
    
    ' 系列の色/サイズ/種類
    Dim i As Integer
    Dim tmpColor As Long
        
    With Me.Series
        .ChartType = Scatter_
        .LoadFromSeriesCollection targetChart.SeriesCollection
        For i = 1 To .Count
            If .Count = 1 Then
                tmpColor = RGB(0, 0, 0)
            Else
                tmpColor = colorPalette.GetColor(i)
            End If
            .Item(i).MarkerFillType = mftNormal
            .Item(i).MarkerColor = tmpColor
            .Item(i).LineColor = tmpColor
            .Item(i).MarkerSize = DEFAULT_MARKER_SIZE
            .Item(i).LineWeight = DEFAULT_LINE_WEIGHT
        Next
    End With
    
    ' marginMode
    IsAutoScaleX = True
    IsAutoScaleY = True
    marginModeX = ammFitBoth
    marginModeY = ammExpandMax
    
    ' 凡例なし
    Me.HasLegend = False

End Sub

Public Sub ApplyTo(ByRef targetChart As Chart)

    With targetChart
        
        ' 外枠線塗りつぶしなし
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.Visible = msoFalse
        
        ' タイトルを削除
        If .HasTitle Then
            .ChartTitle.Delete
        End If
        
        ' 凡例
        .HasLegend = Me.HasLegend
        If .HasLegend Then
            .Legend.position = xlLegendPositionTop
        End If
    
        ' 系列の設定適用
        Me.Series.ApplyTo targetChart
        
        ' 第2軸利用が無い場合、ダミー系列を追加して2軸表示を実現する
        Dim s As Series
        Dim useSecondaryYAxis As Boolean
        Dim hasDummy As Boolean
        
        useSecondaryYAxis = False
        hasDummy = False
        For Each s In .SeriesCollection
            ' ダミーかどうか
            If s.Name = DUMMY_SERIES_NAME Then
                hasDummy = True
            Else
                ' 正規の第2軸利用があるかどうか
                If s.AxisGroup = xlSecondary Then
                    useSecondaryYAxis = True
                End If
            End If
        Next
        
        ' 第2y軸追加
        .HasAxis(xlValue, xlSecondary) = True
        
        ' 正規の第2軸利用が無ければ、第2軸を使うダミー系列で2軸表示を実現
        If Not useSecondaryYAxis Then
            
            ' ダミー系列追加
            If Not hasDummy Then
                Set s = .SeriesCollection.newSeries
                s.Name = DUMMY_SERIES_NAME
            Else
                Set s = .SeriesCollection(DUMMY_SERIES_NAME)
            End If
            s.MarkerStyle = xlMarkerStyleNone
            s.Format.Line.Visible = False
            s.AxisGroup = xlSecondary
            
            ' ダミーの範囲をxyのmax/minに設定
            Dim chartEx As ChartExtents
            
            chartEx = GetChartExtents(targetChart, aguBoth)
            If chartEx.HasData Then
                s.XValues = Array(chartEx.MinX, chartEx.MaxX)
                s.values = Array(chartEx.MinY, chartEx.MaxY)
            End If
        Else
        
            ' 不要なダミー削除
            If hasDummy Then
                .SeriesCollection(DUMMY_SERIES_NAME).Delete
            End If
        End If
    
        ' 凡例からダミーは削除
        If .HasLegend Then
            DeleteLegendEntryBySeriesName targetChart, DUMMY_SERIES_NAME
        End If
    
        ' 第2x軸追加
        .HasAxis(xlCategory, xlSecondary) = True
        
        ' 軸ラベル
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlSecondary).HasTitle = False
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlSecondary).HasTitle = useSecondaryYAxis
                
        ' 数値表示
        .Axes(xlCategory, xlPrimary).TickLabelPosition = xlNextToAxis
        .Axes(xlCategory, xlSecondary).TickLabelPosition = xlNone
        .Axes(xlValue, xlPrimary).TickLabelPosition = xlNextToAxis
        .Axes(xlValue, xlSecondary).TickLabelPosition = IIf(useSecondaryYAxis, xlNextToAxis, xlNone)
    End With

    ' フォント適用
    Me.Font.ApplyTo targetChart
    
    ' レイアウト
    ' フォントサイズ→軸ラベルの必要スペース
    Dim pxPerPt As Double: pxPerPt = 1.333
    Dim axisLabelSpace As Double
    
    axisLabelSpace = Me.Font.Size * pxPerPt * 3
    Me.PlotArea.Top = axisLabelSpace
    Me.PlotArea.Left = axisLabelSpace
    
    ' 軸ラベルスペース→プロットエリアサイズ
'    Dim plotAreaSize As Double
'
''    plotAreaSize = axisLabelSpace * 4
'    plotAreaSize = 217
'    Me.PlotArea.Height = plotAreaSize
'    Me.PlotArea.Width = plotAreaSize
    
    ' 軸ラベル＋プロットエリア＝全体サイズ 適用
    targetChart.ChartArea.Height = axisLabelSpace * 2 + Me.PlotArea.Height
    targetChart.ChartArea.Width = axisLabelSpace * 2 + Me.PlotArea.Width
    
    ' プロットエリア適用
    Me.PlotArea.ApplyTo targetChart

    ' 軸/目盛りプロパティ
    With targetChart
        ' 全軸共通設定
        Dim ax As Axis
        
        For Each ax In .Axes
            ax.HasMajorGridlines = False
            ax.HasMinorGridlines = False
            ax.MajorTickMark = xlTickMarkInside
            ax.MinorTickMark = xlTickMarkNone
            ax.Format.Line.Visible = msoTrue
            ax.Format.Line.ForeColor.RGB = rgbBlack
            ax.Format.Line.Weight = 1.5
            If ax.HasTitle Then
                ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
            End If
        Next ax
    End With

    ' 各軸スケール設定
    ' オリジナル自動
    ' X
    With targetChart.Axes(xlCategory, xlPrimary)
        .MaximumScaleIsAuto = IsAutoScaleX
        .MinimumScaleIsAuto = IsAutoScaleX
        .MajorUnitIsAuto = IsAutoScaleX
        .MinorUnitIsAuto = IsAutoScaleX
    End With
    With targetChart.Axes(xlCategory, xlSecondary)
        .MaximumScaleIsAuto = IsAutoScaleX
        .MinimumScaleIsAuto = IsAutoScaleX
        .MajorUnitIsAuto = IsAutoScaleX
        .MinorUnitIsAuto = IsAutoScaleX
    End With
    ' Y
    With targetChart.Axes(xlValue, xlPrimary)
        .MaximumScaleIsAuto = IsAutoScaleY
        .MinimumScaleIsAuto = IsAutoScaleY
        .MajorUnitIsAuto = IsAutoScaleY
        .MinorUnitIsAuto = IsAutoScaleY
    End With
    With targetChart.Axes(xlValue, xlSecondary)
        .MaximumScaleIsAuto = IsAutoScaleY
        .MinimumScaleIsAuto = IsAutoScaleY
        .MajorUnitIsAuto = IsAutoScaleY
        .MinorUnitIsAuto = IsAutoScaleY
    End With
    
    ' NiceScale
    Dim atu As AxisTypeUsage
    
    If Not (IsAutoScaleX And IsAutoScaleY) Then
        Select Case True
            ' 両方Nice
            Case (Not IsAutoScaleX) And (Not IsAutoScaleY)
                atu = atuBoth
            ' XだけNice
            Case (Not IsAutoScaleX) And IsAutoScaleY
                atu = atuOnlyX
            ' YだけNice
            Case IsAutoScaleX And (Not IsAutoScaleY)
                atu = atuOnlyY
        End Select
        
        NiceScaling targetChart, atu, useSecondaryYAxis, marginModeX, marginModeY
    End If

End Sub

