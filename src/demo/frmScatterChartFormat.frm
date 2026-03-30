VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScatterChartFormat 
   Caption         =   "散布図整形"
   ClientHeight    =   6435
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10395
   OleObjectBlob   =   "frmScatterChartFormat.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmScatterChartFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'　散布図整形フォーム
'---------------------------------------------------------------

Public Enum enSeriesListCol
    index_ = 0
    name_
    dataColumn_
    color_
    size_
    useSecondAxis_
End Enum

Private targetChart_ As Chart
Private chartFormat_ As New clsScatterChart
Private selectedSeriesIdx As Long
Private selectedListIdx As Long
Private selectedSeries As New clsScatterSeries

' --- 対象グラフ ---
Public Property Get targetChart() As Chart
    Set targetChart = targetChart_
End Property

Public Property Set targetChart(ByVal newTargetChart As Chart)
    Set targetChart_ = newTargetChart
End Property

Private Sub UserForm_Activate()

    ' 既定値設定
    chartFormat_.SetDefault Me.targetChart

    ' フォント一覧
    Dim i As Long

    GetFontList
    If cmbFont.ListCount < 1 Then
        cmbFont.AddItem DEFAULT_FONT_NAME
    End If
    cmbFont.value = DEFAULT_FONT_NAME: chartFormat_.Font.FontName = DEFAULT_FONT_NAME

    ' フォントサイズ一覧
    Dim sizes As Variant
    sizes = Array(6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72)
    For i = LBound(sizes) To UBound(sizes)
        Me.cmbFontSize.AddItem sizes(i)
    Next i
    cmbFontSize.value = DEFAULT_FONT_SIZE: chartFormat_.Font.Size = DEFAULT_FONT_SIZE

    ' plotArea
    txtPlotAreaWidth = DEFAULT_PLOTAREA_WIDTH: chartFormat_.PlotArea.Width = DEFAULT_PLOTAREA_WIDTH
    txtPlotAreaHeight = DEFAULT_PLOTAREA_HEIGHT: chartFormat_.PlotArea.Height = DEFAULT_PLOTAREA_HEIGHT

    ' 凡例
    optHasLegend = chartFormat_.HasLegend

    ' 系列リスト初期設定
    lstSeries.ColumnWidths = _
        "0;" & _
        lblListHeaderName.Width & ";" & _
        lblListHeaderDataColumn.Width & ";" & _
        lblListHeaderColor.Width & ";" & _
        lblListHeaderSize.Width & ";" & _
        lblListHeaderSecond.Width - 2
    ToggleType ' グラフ種類に応じた設定
    
    ' Scale
    optOriginalX = chartFormat_.IsAutoScaleX
    optOriginalY = chartFormat_.IsAutoScaleY
    
    ' MarginModeコンボソース
    Dim ammItems(0 To 3, 0 To 1) As Variant
    
    ammItems(0, 0) = ammFitBoth: ammItems(0, 1) = "Fit Both"
    ammItems(1, 0) = ammExpandMax: ammItems(1, 1) = "Expand Max"
    ammItems(2, 0) = ammExpandMin: ammItems(2, 1) = "Expand Min"
    ammItems(3, 0) = ammExpandBoth: ammItems(3, 1) = "Expand Both"
    
    cmbMarginModeX.List = ammItems
    cmbMarginModeY.List = ammItems
    cmbMarginModeX.ListIndex = chartFormat_.marginModeX
    cmbMarginModeY.ListIndex = chartFormat_.marginModeY
    cmbMarginModeX.Enabled = (optNiceX = True)
    cmbMarginModeY.Enabled = (optNiceY = True)
    
End Sub

Sub GetFontList()
    
    Dim tempBar As CommandBar
    Dim fontList As CommandBarControl
    Dim i As Long
    Dim FontName As String
    
    'シート選択状態が必要
    ActiveSheet.Cells(1, 1).Select
    
    'フォント一覧を取得
    Set tempBar = Application.CommandBars.Add
    Set fontList = tempBar.Controls.Add(ID:=1728)
    
    'フォントの数だけ繰り返し
    For i = 1 To fontList.ListCount - 1
        'フォント名を取得
        FontName = fontList.List(i)
        'フォント名をコンボに設定
        cmbFont.AddItem FontName
    Next
    
    '後片付け
    fontList.Delete
    tempBar.Delete
    Set fontList = Nothing
    Set tempBar = Nothing

End Sub

Private Sub btnCancel_Click()
    
    Unload Me

End Sub

Private Sub btnExecute_Click()
    
    chartFormat_.ApplyTo targetChart_
    Unload Me

End Sub

' --- グラフの種類 ---
Private Sub optLine_AfterUpdate()

    ToggleType

End Sub

Private Sub optScatter_AfterUpdate()

    ToggleType

End Sub

Private Sub optSmoothLine_AfterUpdate()

    ToggleType

End Sub

Private Sub ToggleType()

    ' 設定値
    Select Case True
        Case optScatter
            chartFormat_.Series.ChartType = Scatter_
        Case optLine
            chartFormat_.Series.ChartType = Line_
            chartFormat_.Series.IsSmooth = False
        Case optSmoothLine
            chartFormat_.Series.ChartType = Line_
            chartFormat_.Series.IsSmooth = True
    End Select

    Dim colorText As String
    Dim sizeText As String
    Dim colorValue As Long
    Dim sizeValue As Single
    Dim i As Long
    
    ' UI表示
    Select Case chartFormat_.Series.ChartType
        Case Scatter_
            colorText = "マーカーの色"
            sizeText = "マーカーサイズ"
            fraMarkerFillType.Visible = True
            lblMarkerFillType.Visible = True
            btnMarkerFillNormal.Visible = True
            btnMarkerFillWhite.Visible = True
        Case Line_
            colorText = "線の色"
            sizeText = "線の幅(pt)"
            fraMarkerFillType.Visible = False
            lblMarkerFillType.Visible = False
            btnMarkerFillNormal.Visible = False
            btnMarkerFillWhite.Visible = False
    End Select

    lblColor.Caption = colorText
    lblSize.Caption = sizeText
    lblListHeaderColor.Caption = colorText
    lblListHeaderSize.Caption = sizeText

    lstSeries.Clear
    With chartFormat_.Series
        For i = 1 To .Count
            Select Case chartFormat_.Series.ChartType
                Case Scatter_
                    colorValue = .Item(i).MarkerColor
                    sizeValue = .Item(i).MarkerSize
                Case Line_
                    colorValue = .Item(i).LineColor
                    sizeValue = .Item(i).LineWeight
            End Select
            AddRow lstSeries, Array( _
                .Item(i).index, _
                .Item(i).Name, _
                .Item(i).DataColumn, _
                IIf(.Item(i).MarkerFillType = mftWhite, "(W)", "") & _
                    colorPalette.GetColorName(colorValue), _
                sizeValue, _
                IIf(.Item(i).useSecondAxis, "○", "") _
            )
        Next
    End With
    
    If lstSeries.ListCount > 0 Then lstSeries.ListIndex = 0
    If lstSeries.ListCount = 1 Then
        chkSecondAxis.Enabled = False
    Else
        chkSecondAxis.Enabled = True
    End If
    
    ' 選択系列の詳細表示
    ShowSeriesDetail

End Sub

Private Sub AddRow(lb As MSForms.ListBox, values As Variant)
    
    Dim r As Long, i As Long
    r = lb.ListCount
    lb.AddItem values(0)
    For i = 1 To UBound(values)
        lb.List(r, i) = values(i)
    Next

End Sub

' --- フォント ---
Private Sub cmbFont_AfterUpdate()

    chartFormat_.Font.FontName = cmbFont.value

End Sub

Private Sub cmbFontSize_AfterUpdate()

    chartFormat_.Font.Size = cmbFontSize.value

End Sub

' --- プロットエリア ---
Private Sub txtPlotAreaWidth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtPlotAreaWidth, True) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtPlotAreaWidth_AfterUpdate()

    If (Not IsNumeric(txtPlotAreaWidth)) _
    Or val(txtPlotAreaWidth) <= 0 Then
        txtPlotAreaWidth = ""
        chartFormat_.PlotArea.Width = DEFAULT_PLOTAREA_WIDTH
    Else
        chartFormat_.PlotArea.Width = txtPlotAreaWidth.value
    End If
    
End Sub

Private Sub txtPlotAreaHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If Not IsValidNumericKeyPress(KeyAscii, txtPlotAreaHeight, True) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtPlotAreaHeight_AfterUpdate()

    If (Not IsNumeric(txtPlotAreaHeight)) _
    Or val(txtPlotAreaHeight) <= 0 Then
        txtPlotAreaHeight = ""
        chartFormat_.PlotArea.Height = DEFAULT_PLOTAREA_HEIGHT
    Else
        chartFormat_.PlotArea.Height = txtPlotAreaHeight.value
    End If
    
End Sub

' --- 凡例 ---
Private Sub optHasLegend_AfterUpdate()

    ToggleHasLegend

End Sub

Private Sub optNoLegend_AfterUpdate()

    ToggleHasLegend
    
End Sub

Private Sub ToggleHasLegend()

    If optHasLegend = True Then
        chartFormat_.HasLegend = True
    Else
        chartFormat_.HasLegend = False
    End If

End Sub

' --- 自動スケール設定 ---
Private Sub optOriginalX_AfterUpdate()

    cmbMarginModeX.Enabled = (optNiceX = True)
    chartFormat_.IsAutoScaleX = True

End Sub

Private Sub optOriginalY_AfterUpdate()

    cmbMarginModeY.Enabled = (optNiceY = True)
    chartFormat_.IsAutoScaleY = True
    
End Sub

Private Sub optNiceX_AfterUpdate()

    cmbMarginModeX.Enabled = (optNiceX = True)
    chartFormat_.IsAutoScaleX = False
    chartFormat_.marginModeX = cmbMarginModeX.value

End Sub

Private Sub optNiceY_AfterUpdate()

    cmbMarginModeY.Enabled = (optNiceY = True)
    chartFormat_.IsAutoScaleY = False
    chartFormat_.marginModeY = cmbMarginModeY.value
    
End Sub

Private Sub cmbMarginModeX_AfterUpdate()

    chartFormat_.marginModeX = cmbMarginModeX.value

End Sub

Private Sub cmbMarginModeY_AfterUpdate()

    chartFormat_.marginModeY = cmbMarginModeY.value

End Sub

' --- 系列の色 ---
Private Sub btnColor_Click()

    Dim rgbValue As Long
    Dim r As Long, g As Long, b As Long
    
    Select Case chartFormat_.Series.ChartType
        Case Scatter_
            rgbValue = selectedSeries.MarkerColor
        Case Line_
            rgbValue = selectedSeries.LineColor
    End Select
    
    r = rgbValue And &HFF
    g = (rgbValue \ &H100) And &HFF
    b = (rgbValue \ &H10000) And &HFF

    Dim response As Long '戻り値。RGB値が10進数で返される
    
    If Application.Dialogs(xlDialogEditColor).Show(1, r, g, b) = True Then
        ' RGB値を取得
        response = ActiveWorkbook.colors(1)
        
        ' プロパティ変更
        selectedSeries.MarkerColor = response
        selectedSeries.LineColor = response
        
        ' UI更新
        lblColorDisplay.ForeColor = response
        lstSeries.List(selectedListIdx, enSeriesListCol.color_) = _
            IIf(selectedSeries.MarkerFillType = mftWhite, "(W)", "") & _
            colorPalette.GetColorName(response)
    End If

End Sub

' --- マーカーの塗りつぶし ---
Private Sub optMarkerFillNormal_Change()

    ToggleSelectedMarkerFillType mftNormal

End Sub

Private Sub optMarkerFillWhite_Change()
   
    ToggleSelectedMarkerFillType mftWhite

End Sub

Private Sub ToggleSelectedMarkerFillType(ByVal typeTo As MarkerFillType)

    ' プロパティ変更
    selectedSeries.MarkerFillType = typeTo
    
    ' UI更新
    lstSeries.List(selectedListIdx, enSeriesListCol.color_) = _
        IIf(selectedSeries.MarkerFillType = mftWhite, "(W)", "") & _
        colorPalette.GetColorName(selectedSeries.MarkerColor)

End Sub

Private Sub btnMarkerFillNormal_Click()

    ToggleAllMarkerFillType mftNormal

End Sub

Private Sub btnMarkerFillWhite_Click()

    ToggleAllMarkerFillType mftWhite

End Sub

' 塗りつぶし一括設定
Private Sub ToggleAllMarkerFillType(ByVal typeTo As MarkerFillType)

    ' プロパティ変更
    Dim i As Integer
    
    For i = 1 To chartFormat_.Series.Count
        chartFormat_.Series.Item(i).MarkerFillType = typeTo
    Next
    
    ' 系列リストボックスのUI更新
    Dim s As clsScatterSeries
    
    For i = 0 To lstSeries.ListCount - 1
        Set s = chartFormat_.Series.Item(lstSeries.List(i, 0))
        lstSeries.List(i, enSeriesListCol.color_) = _
            IIf(typeTo = mftWhite, "(W)", "") & _
            colorPalette.GetColorName(s.MarkerColor)
    Next
    
    ' 選択中の系列のUI更新
    SetMarkerFillOptionUI typeTo

End Sub

' --- 系列のサイズ ---
Private Sub spnSize_SpinDown()

    ChangeSize -IIf(chartFormat_.Series.ChartType = Scatter_, 1, 0.25)

End Sub

Private Sub spnSize_SpinUp()

    ChangeSize IIf(chartFormat_.Series.ChartType = Scatter_, 1, 0.25)

End Sub

Private Sub ChangeSize(ByVal deltaValue As Single)

    With selectedSeries
        Select Case chartFormat_.Series.ChartType
            Case Scatter_
                .MarkerSize = .MarkerSize + deltaValue
                txtSize.value = .MarkerSize
                lstSeries.List(selectedListIdx, enSeriesListCol.size_) = .MarkerSize
            Case Line_
                .LineWeight = .LineWeight + deltaValue
                txtSize.value = .LineWeight
                lstSeries.List(selectedListIdx, enSeriesListCol.size_) = .LineWeight
        End Select
    End With

End Sub

Private Sub txtSeriesName_AfterUpdate()
    
    With selectedSeries
        .Name = txtSeriesName.value
        lstSeries.List(selectedListIdx, enSeriesListCol.name_) = txtSeriesName.value
    End With

End Sub

' --- 第2軸利用 ---
Private Sub chkSecondAxis_AfterUpdate()

    With selectedSeries
        .useSecondAxis = chkSecondAxis.value
        lstSeries.List(selectedListIdx, enSeriesListCol.useSecondAxis_) = IIf(chkSecondAxis.value, "○", "")
    End With

End Sub

' --- 系列選択 ---
Private Sub lstSeries_Change()
    
    ShowSeriesDetail

End Sub

' 選択された系列のプロパティを表示
Private Sub ShowSeriesDetail()

    Dim selectedColor As Long
    Dim selectedSize As Single

    selectedListIdx = lstSeries.ListIndex
    If selectedListIdx = -1 Then
        fraTargetSeries.Visible = False
        selectedSeriesIdx = 0
    Else
        fraTargetSeries.Visible = True
        Set selectedSeries = chartFormat_.Series.Item(lstSeries.List(selectedListIdx))
        With selectedSeries
            ' 系列名とデータ列
            selectedSeriesIdx = .index
            txtSeriesName.value = .Name
            
            ' マーカー or 線
            Select Case chartFormat_.Series.ChartType
                Case Scatter_
                    selectedColor = .MarkerColor
                    selectedSize = .MarkerSize
                Case Line_
                    selectedColor = .LineColor
                    selectedSize = .LineWeight
            End Select
            
            ' 色
            If selectedColor >= 0 Then
                lblColorDisplay.ForeColor = selectedColor
            Else
                lblColorDisplay.ForeColor = RGB(0, 0, 0)
            End If
            
            ' 塗りつぶし
            SetMarkerFillOptionUI .MarkerFillType

            ' サイズ
            txtSize.value = selectedSize
            
            ' 第2軸利用
            If .useSecondAxis Then
                chkSecondAxis.value = True
            Else
                chkSecondAxis.value = False
            End If
        End With
    End If

End Sub

' --- ヘルパー ---
' マーカー塗りつぶしオプションのUI更新
Private Sub SetMarkerFillOptionUI(ByVal typeTo As MarkerFillType)

    If typeTo = mftNormal Then
        optMarkerFillNormal.value = True
    ElseIf typeTo = mftWhite Then
        optMarkerFillWhite.value = True
    End If

End Sub
