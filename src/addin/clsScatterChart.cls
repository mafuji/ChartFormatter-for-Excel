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

' 散布図の設定
Private height_ As Double
Private width_ As Double
Private font_ As New clsFont
Private plotArea_ As New clsPlotArea
Private series_ As New clsScatterSeriesCollection
Private hasLegend_ As Boolean

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
Public Property Get font() As clsFont
    Set font = font_
End Property

Public Property Set font(ByVal newFont As clsFont)
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

Public Sub SetDefault(ByVal targetChart As Chart)

    With Me.font
        .fontName = DEFAULT_FONT_NAME
        .Size = DEFAULT_FONT_SIZE
    End With
    
    Dim i As Integer
    
    With Me.Series
        .LoadFromSeriesCollection targetChart.SeriesCollection
        For i = 1 To .Count
            If .Count = 1 Then
                .Item(i).MarkerColor = RGB(0, 0, 0)
            Else
                .Item(i).MarkerColor = colorPalette.GetColor(i)
            End If
            .Item(i).MarkerSize = DEFAULT_MARKER_SIZE
        Next
    End With

    Me.HasLegend = False

End Sub

Public Sub ApplyTo(ByRef targetChart As Chart)

    ' 現状固定の設定適用
    With targetChart
        
        ' 外枠線塗りつぶしなし
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.Visible = msoFalse
        
        ' タイトルを削除
        If .HasTitle Then
            .ChartTitle.Delete
        End If
        
        ' 凡例削除
        .HasLegend = Me.HasLegend
        If .HasLegend Then
            .Legend.Position = xlLegendPositionTop
        End If
    
        ' 系列の設定適用
        Me.Series.ApplyTo targetChart
        
        ' 2系列以上かつ第2y軸利用なしの場合
        ' 2系列目をダミーの第2y軸利用として設定し
        ' 軸目盛を第1y軸と合わせることで上下左右目盛を実現する
        Dim s As Series
        Dim onlyUseFirstAxis As Boolean
        
        If .SeriesCollection.Count > 1 Then
            onlyUseFirstAxis = True
            For Each s In .SeriesCollection
                If s.AxisGroup = xlSecondary Then
                    onlyUseFirstAxis = False
                    Exit For
                End If
            Next
            If onlyUseFirstAxis Then
                .SeriesCollection(2).AxisGroup = xlSecondary
                .Axes(xlValue, xlSecondary).MinimumScale = .Axes(xlValue, xlPrimary).MinimumScale
                .Axes(xlValue, xlSecondary).MaximumScale = .Axes(xlValue, xlPrimary).MaximumScale
                .Axes(xlValue, xlSecondary).MajorUnit = .Axes(xlValue, xlPrimary).MajorUnit
                .Axes(xlValue, xlSecondary).MinorUnit = .Axes(xlValue, xlPrimary).MinorUnit
                .Axes(xlValue, xlSecondary).CrossesAt = .Axes(xlValue, xlPrimary).CrossesAt
            End If
            
            ' 第2x軸追加
            .HasAxis(xlCategory, xlSecondary) = True
            
            ' 軸ラベル
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlSecondary).HasTitle = False
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlSecondary).HasTitle = IIf(onlyUseFirstAxis, False, True)
        End If
                
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
        Next ax
    End With

    ' フォント適用
    Me.font.ApplyTo targetChart
    
    ' レイアウト
    ' フォントサイズ→軸ラベルの必要スペース
    Dim pxPerPt As Double: pxPerPt = 1.333
    Dim axisLabelSpace As Double
    
    axisLabelSpace = Me.font.Size * pxPerPt * 3
    Me.PlotArea.Top = axisLabelSpace
    Me.PlotArea.Left = axisLabelSpace
    
    ' 軸ラベルスペース→プロットエリアサイズ
    Dim plotAreaSize As Double
    
    plotAreaSize = axisLabelSpace * 5
    Me.PlotArea.Height = plotAreaSize
    Me.PlotArea.Width = plotAreaSize
    
    ' 軸ラベル＋プロットエリア＝全体サイズ 適用
    targetChart.ChartArea.Height = axisLabelSpace * 2 + plotAreaSize
    targetChart.ChartArea.Width = axisLabelSpace * 2 + plotAreaSize
    
    ' プロットエリア適用
    Me.PlotArea.ApplyTo targetChart

End Sub
