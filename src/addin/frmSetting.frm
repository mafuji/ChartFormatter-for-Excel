VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetting 
   Caption         =   "散布図整形"
   ClientHeight    =   4272
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9945
   OleObjectBlob   =   "frmSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enSeriesListCol
    index_ = 0
    name_
    dataColumn_
    markerColor_
    markerSize_
    useSecondAxis_
End Enum

Private targetChart_ As Chart
Private setting_ As clsScatterChart
Private selectedSeriesIdx As Long
Private selectedListIdx As Long
Private selectedSeries As New clsScatterSeries

Public Property Get targetChart() As Chart
    Set targetChart = targetChart_
End Property

Public Property Set targetChart(ByVal newTargetChart As Chart)
    Set targetChart_ = newTargetChart
End Property

Private Sub btnCancel_Click()
    
    Unload Me

End Sub

Private Sub btnExecute_Click()
    
    setting_.ApplyTo targetChart_
    Unload Me

End Sub

Private Sub btnMarkerColor_Click()

    Dim rgbValue As Long
    Dim R As Long, G As Long, B As Long
    
    rgbValue = selectedSeries.MarkerColor
    R = rgbValue And &HFF
    G = (rgbValue \ &H100) And &HFF
    B = (rgbValue \ &H10000) And &HFF

    Dim response As Long '戻り値。RGB値が10進数で返される
    
    If Application.Dialogs(xlDialogEditColor).Show(1, R, G, B) = True Then
        response = ActiveWorkbook.colors(1)
        selectedSeries.MarkerColor = response
        lblMarkerColor.ForeColor = response
        lstSeries.List(selectedListIdx, enSeriesListCol.markerColor_) = colorPalette.GetColorName(response)
    End If

End Sub

Private Sub chkSecondAxis_AfterUpdate()

    With selectedSeries
        .UseSecondAxis = chkSecondAxis.Value
        lstSeries.List(selectedListIdx, enSeriesListCol.useSecondAxis_) = IIf(chkSecondAxis.Value, "○", "")
    End With

End Sub

Private Sub cmbFont_AfterUpdate()

    setting_.font.fontName = cmbFont.Value

End Sub

Private Sub cmbFontSize_AfterUpdate()

    setting_.font.Size = cmbFontSize.Value

End Sub

Private Sub optHasLegend_AfterUpdate()

    ToggleHasLegend

End Sub

Private Sub optNoLegend_AfterUpdate()

    ToggleHasLegend
    
End Sub

Private Sub spnMarkerSize_SpinDown()

    ChangeMarkerSize -1

End Sub

Private Sub spnMarkerSize_SpinUp()

    ChangeMarkerSize 1

End Sub

Private Sub ChangeMarkerSize(ByVal deltaValue As Integer)

    With selectedSeries
        .MarkerSize = .MarkerSize + deltaValue
        txtMarkerSize.Value = .MarkerSize
        lstSeries.List(selectedListIdx, enSeriesListCol.markerSize_) = .MarkerSize
    End With

End Sub

Private Sub txtSeriesName_AfterUpdate()
    
    With selectedSeries
        .Name = txtSeriesName.Value
        lstSeries.List(selectedListIdx, enSeriesListCol.name_) = txtSeriesName.Value
    End With

End Sub

Private Sub UserForm_Activate()

    ' 設定生成～既定値設定
    Set setting_ = New clsScatterChart
    setting_.SetDefault Me.targetChart

    ' フォント一覧
    Dim i As Long

    GetFontList
    If cmbFont.ListCount < 1 Then
        cmbFont.AddItem DEFAULT_FONT_NAME
    End If
    cmbFont.Value = DEFAULT_FONT_NAME

    ' フォントサイズ一覧
    Dim sizes As Variant
    sizes = Array(6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72)
    For i = LBound(sizes) To UBound(sizes)
        Me.cmbFontSize.AddItem sizes(i)
    Next i
    cmbFontSize.Value = setting_.font.Size

    ' 凡例
    If setting_.HasLegend Then
        optHasLegend = True
        optNoLegend = False
    Else
        optHasLegend = False
        optNoLegend = True
    End If

    ' 系列リスト初期設定
    lstSeries.ColumnWidths = _
        "0;" & _
        lblListHeaderName.Width & ";" & _
        lblListHeaderDataColumn.Width & ";" & _
        lblListHeaderColor.Width & ";" & _
        lblListHeaderSize.Width & ";" & _
        lblListHeaderSecond.Width
    lstSeries.Clear
    With setting_.Series
        For i = 1 To .Count
            AddRow lstSeries, Array( _
                .Item(i).index, _
                .Item(i).Name, _
                .Item(i).DataColumn, _
                colorPalette.GetColorName(.Item(i).MarkerColor), _
                .Item(i).MarkerSize, _
                IIf(.Item(i).UseSecondAxis, _
                    "○", "" _
                ) _
            )
        Next
    End With
    If lstSeries.ListCount > 0 Then lstSeries.ListIndex = 0
    If lstSeries.ListCount = 1 Then
        chkSecondAxis.Enabled = False
    Else
        chkSecondAxis.Enabled = True
    End If
    ShowSeriesDetail
    
End Sub

Private Sub AddRow(lb As MSForms.ListBox, values As Variant)
    
    Dim R As Long, i As Long
    R = lb.ListCount
    lb.AddItem values(0)
    For i = 1 To UBound(values)
        lb.List(R, i) = values(i)
    Next

End Sub

Private Sub lstSeries_Change()

    ShowSeriesDetail

End Sub

Private Sub ShowSeriesDetail()

    selectedListIdx = lstSeries.ListIndex
    If selectedListIdx = -1 Then
        fraTargetSeries.Visible = False
        selectedSeriesIdx = 0
    Else
        fraTargetSeries.Visible = True
        Set selectedSeries = setting_.Series.Item(lstSeries.List(selectedListIdx))
        With selectedSeries
            selectedSeriesIdx = .index
            txtSeriesName.Value = .Name
            If .MarkerColor >= 0 Then
                lblMarkerColor.ForeColor = .MarkerColor
            Else
                lblMarkerColor.ForeColor = RGB(0, 0, 0)
            End If
            txtMarkerSize.Value = .MarkerSize
            If .UseSecondAxis Then
                chkSecondAxis.Value = True
            Else
                chkSecondAxis.Value = False
            End If
        End With
    End If

End Sub

Sub GetFontList()
    
    Dim tempBar As CommandBar
    Dim fontList As CommandBarControl
    Dim i As Long
    Dim fontName As String
    
    'シート選択状態が必要
    ActiveSheet.Cells(1, 1).Select
    
    'フォント一覧を取得
    Set tempBar = Application.CommandBars.Add
    Set fontList = tempBar.Controls.Add(ID:=1728)
    
    'フォントの数だけ繰り返し
    For i = 1 To fontList.ListCount - 1
        'フォント名を取得
        fontName = fontList.List(i)
        'フォント名をコンボに設定
        cmbFont.AddItem fontName
    Next
    
    '後片付け
    fontList.Delete
    tempBar.Delete
    Set fontList = Nothing
    Set tempBar = Nothing

End Sub

Private Sub ToggleHasLegend()

    If optHasLegend = True Then
        setting_.HasLegend = True
    Else
        setting_.HasLegend = False
    End If

End Sub
