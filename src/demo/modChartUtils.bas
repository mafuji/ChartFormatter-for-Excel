Attribute VB_Name = "modChartUtils"
Option Explicit

'---------------------------------------------------------------
' グラフ系ユーティリティ
'---------------------------------------------------------------

' 第2軸利用のためのダミー系列名
Public Const DUMMY_SERIES_NAME As String = "__DUMMY_SERIES_NAME_106841651__"

' 各種既定値
Public Const DEFAULT_FONT_NAME As String = "Arial Narrow"
Public Const DEFAULT_FONT_NAME_FAR_EAST As String = "MS UI Gothic"
Public Const DEFAULT_FONT_SIZE As Single = 16
Public Const DEFAULT_MARKER_SIZE As Integer = 6
Public Const DEFAULT_LINE_WEIGHT As Single = 1
Public Const DEFAULT_PLOTAREA_WIDTH As Double = 230
Public Const DEFAULT_PLOTAREA_HEIGHT As Double = 230

' カスタム表示形式
Public Const NUMBER_FORMAT_ZERO_OR_DECIMAL As String = "[=0]0;0.0"
Public Const NUMBER_FORMAT_INT As String = "0"

' グラフタイプ
Public Enum enChartType
    Scatter_ = 0
    Line_
End Enum

' 充放電
Public Enum enChargeStatus
    Charge = 0
    Discharge
End Enum

' グラフの端の値セット
Public Type ChartExtents
    HasData As Boolean
    MinX As Double
    MinY As Double
    MaxX As Double
    MaxY As Double
End Type

' 軸の設定セット
Public Type AxisScale
    Min As Double
    Max As Double
    MajorUnit As Double
End Type

' グラフの端の設定
Public Enum AxisMarginMode
    ammFitBoth = 0          ' min,max とも一致OK
    ammExpandMax = 1        ' maxのみ +1目盛
    ammExpandMin = 2        ' minのみ -1目盛
    ammExpandBoth = 3       ' 両方 1目盛拡張
End Enum

' Y軸利用のタイプ（Chart全体として）
Public Enum AxisGroupUsage
    aguBoth = 0 ' 1Y,2Y両方使う
    aguOnly1stY = 1 ' 1Yのみ使う
    aguOnly2ndY = 2 ' 2Yのみ使う（基本ない）
End Enum

' NiceScaleの利用タイプ（Chart全体として）
Public Enum AxisTypeUsage
    atuBoth = 0 ' 両方適用
    atuOnlyX = 1 ' X軸のみ適用
    atuOnlyY = 2 ' Y軸のみ適用
End Enum

' カラーバーラベル位置
Public Enum LabelPosition
    lpRight = 0
    lpLeft = 1
    lpTop = 2
    lpBottom = 3
End Enum

' カラーバーラベル表示間隔
Public Enum LabelMode
    lmAll = 0
    lmMinMax = 1
    lmCustom = 2
End Enum

' カラーバーラベル表示形式
Public Enum LabelFormat
    lfCardinal = 0 ' 1,2,...
    lfOrdinal = 1 ' 1st,2nd,...
End Enum

' マーカー塗りつぶしタイプ
Public Enum MarkerFillType
    mftNormal = 0 ' 任意の色
    mftWhite = 1 ' 白抜き
End Enum

Public Function GetChartExtents(ByVal ch As Chart, ByVal targetAxis As AxisGroupUsage) As ChartExtents
    Dim r As ChartExtents
    Dim s As Series
    Dim vx As Variant, vy As Variant
    Dim i As Long, n As Long
    Dim x As Double, y As Double
    Dim gotAny As Boolean

    ' 初期化（初回比較用に極大/極小を入れておく）
    r.MinX = 1.79E+308
    r.MinY = 1.79E+308
    r.MaxX = -1.79E+308
    r.MaxY = -1.79E+308

    If ch Is Nothing Then
        GetChartExtents = r
        Exit Function
    End If

    On Error GoTo SafeExit

    For Each s In ch.SeriesCollection
        ' ダミー系列はスキップ
        If s.Name = DUMMY_SERIES_NAME Then GoTo NextSeries
        
        ' X / Y を一次元配列に正規化
        vx = SeriesValuesToArray(s.XValues)
        vy = SeriesValuesToArray(s.values)

        If IsEmpty(vy) Then GoTo NextSeries

        ' 長さ決定（X/Y 不一致なら短い方に合わせる）
        n = GetLinearLength(vy)
        If Not IsEmpty(vx) Then
            n = Application.Min(n, GetLinearLength(vx))
        End If
        If n <= 0 Then GoTo NextSeries

        ' X が未設定なら 1..n を X とする
        If IsEmpty(vx) Then
            ReDim vx(1 To n) As Double
            For i = 1 To n
                vx(i) = CDbl(i)
            Next
        End If

        ' 走査
        For i = 1 To n
            If TryGetNumeric(vx(i), x) And TryGetNumeric(vy(i), y) Then
                If x < r.MinX Then r.MinX = x
                If x > r.MaxX Then r.MaxX = x
                
                ' 利用軸指定に応じてスキップ
                If targetAxis = aguBoth _
                Or (targetAxis = aguOnly1stY And s.AxisGroup = xlPrimary) _
                Or (targetAxis = aguOnly2ndY And s.AxisGroup = xlSecondary) Then
                    If y < r.MinY Then r.MinY = y
                    If y > r.MaxY Then r.MaxY = y
                End If
                gotAny = True
            End If
        Next i

NextSeries:
    Next s

SafeExit:
    r.HasData = gotAny
    If Not gotAny Then
        ' データが1件もなければ 0 に寄せる
        r.MinX = 0: r.MinY = 0: r.MaxX = 0: r.MaxY = 0
    End If
    GetChartExtents = r
End Function

Public Function GetChartExtentsFromChartObject(ByVal co As ChartObject) As ChartExtents
    If co Is Nothing Then
        GetChartExtentsFromChartObject = EmptyChartExtents()
    Else
        GetChartExtentsFromChartObject = GetChartExtents(co.Chart, aguBoth)
    End If
End Function

Private Function EmptyChartExtents() As ChartExtents
    Dim r As ChartExtents
    r.HasData = False
    r.MinX = 0: r.MinY = 0: r.MaxX = 0: r.MaxY = 0
    EmptyChartExtents = r
End Function

Private Function SeriesValuesToArray(ByVal v As Variant) As Variant
    Dim resultArr As Variant

    If IsEmpty(v) Then
        SeriesValuesToArray = Empty
        Exit Function
    End If

    If TypeName(v) = "Range" Then
        Dim c As range, cnt As Long
        If v.Cells.CountLarge = 0 Then
            SeriesValuesToArray = Empty
            Exit Function
        End If
        ReDim resultArr(1 To CLng(v.Cells.CountLarge)) As Variant
        For Each c In v.Cells
            cnt = cnt + 1
            resultArr(cnt) = c.value
        Next
        SeriesValuesToArray = resultArr
        Exit Function
    End If

    If IsArray(v) Then
        ' 2次元/1次元どちらでも 1..N にフラット化する
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        Dim i As Long, j As Long, k As Long

        On Error Resume Next
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        If Err.Number <> 0 Then
            ' 1次元配列
            Err.Clear
            ReDim resultArr(1 To UBound(v) - LBound(v) + 1) As Variant
            k = 0
            For i = LBound(v) To UBound(v)
                k = k + 1
                resultArr(k) = v(i)
            Next
            SeriesValuesToArray = resultArr
            Exit Function
        End If
        On Error GoTo 0

        ' 2次元配列（行×列）を 1..N に展開（行→列の順で走査）
        ReDim resultArr(1 To (ub1 - lb1 + 1) * (ub2 - lb2 + 1)) As Variant
        k = 0
        For i = lb1 To ub1
            For j = lb2 To ub2
                k = k + 1
                resultArr(k) = v(i, j)
            Next j
        Next i
        SeriesValuesToArray = resultArr
        Exit Function
    End If

    ' 単一値
    ReDim resultArr(1 To 1) As Variant
    resultArr(1) = v
    SeriesValuesToArray = resultArr
End Function

Private Function GetLinearLength(ByVal arr As Variant) As Long
    If IsEmpty(arr) Then
        GetLinearLength = 0
        Exit Function
    End If
    If IsArray(arr) Then
        GetLinearLength = UBound(arr) - LBound(arr) + 1
    Else
        GetLinearLength = 1
    End If
End Function

Private Function TryGetNumeric(ByVal v As Variant, ByRef d As Double) As Boolean
    On Error GoTo NG
    If IsError(v) Then GoTo NG
    If IsEmpty(v) Or IsNull(v) Then GoTo NG
    If VarType(v) = vbString And Trim$(v) = "" Then GoTo NG

    If IsDate(v) Then
        d = CDbl(CDate(v))   ' 日付 → シリアル値
        TryGetNumeric = True
        Exit Function
    End If

    If IsNumeric(v) Then
        d = CDbl(v)
        TryGetNumeric = True
        Exit Function
    End If
NG:
    TryGetNumeric = False
End Function

' データ範囲を取得する
Function GetRangeFromSeries(s As Series, targetAxis As Integer) As range

    Dim f As String
    Dim parts As Variant

    f = s.Formula
    f = Replace(f, "=SERIES(", "")
    f = Left(f, Len(f) - 1)

    parts = Split(f, ",")

    'parts(1) が XValues の範囲
    On Error Resume Next
    Set GetRangeFromSeries = range(parts(targetAxis))
    On Error GoTo 0

End Function

' アドレスから列番号取得
Function GetColumnLetter(rng As range) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^[A-Z]+"
    re.IgnoreCase = False

    Dim m As Object
    If re.test(rng.Address(False, False)) Then
        Set m = re.Execute(rng.Address(False, False))(0)
        GetColumnLetter = m.value
    End If
End Function

' オブジェクトの上位にChartがいれば返す。居なければnothingを返す
Function GetChart(ByVal targetObject As Object) As Object

    Dim obj As Object

    Set obj = targetObject
    
    On Error Resume Next
    
    Do Until Err.Number <> 0
        If TypeName(obj) = "Chart" Then
            Set GetChart = obj
            Exit Function
        Else
            Set obj = obj.Parent
        End If
        
        If Err.Number <> 0 Or TypeName(obj) = "Workbook" Then
            Set GetChart = Nothing
            Exit Function
        End If
    Loop
    
    On Error GoTo 0

    Set GetChart = Nothing

End Function

' 系列名から対応する凡例を削除する
Public Sub DeleteLegendEntryBySeriesName(ByVal ch As Chart, ByVal SeriesName As String)
    Dim i As Long
    If ch Is Nothing Or Not ch.HasLegend Then Exit Sub
    For i = ch.SeriesCollection.Count To 1 Step -1
        If StrComp(ch.SeriesCollection(i).Name, SeriesName, vbTextCompare) = 0 Then
            If i <= ch.Legend.LegendEntries.Count Then
                ch.Legend.LegendEntries(i).Delete
            End If
            Exit For
        End If
    Next
End Sub

' 散布図ファミリーの ChartType か？
Private Function IsScatterType(ByVal t As XlChartType) As Boolean
    Select Case t
        Case xlXYScatter, _
             xlXYScatterLines, _
             xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, _
             xlXYScatterSmoothNoMarkers
            IsScatterType = True
        Case Else
            IsScatterType = False
    End Select
End Function

' chart が「散布図系（単独 or 複合でも散布図系列のみ）」かどうか
' ・系列が1本以上あること
' ・すべての Series.ChartType が散布図ファミリーであること
Public Function IsScatterOnlyChart(ByVal ch As Chart) As Boolean
    Dim s As Series
    Dim found As Boolean

    If ch Is Nothing Then Exit Function
    If ch.SeriesCollection.Count = 0 Then Exit Function

    For Each s In ch.SeriesCollection
        found = True
        If Not IsScatterType(s.ChartType) Then
            IsScatterOnlyChart = False
            Exit Function
        End If
    Next

    IsScatterOnlyChart = found
End Function

' 便利ラッパ：ChartObject から判定
Public Function IsScatterOnlyChartObject(ByVal co As ChartObject) As Boolean
    If co Is Nothing Then Exit Function
    IsScatterOnlyChartObject = IsScatterOnlyChart(co.Chart)
End Function

' ちょうどいいスケールを設定する
Public Sub NiceScaling(ByRef targetChart As Chart, _
                       ByVal atu As AxisTypeUsage, _
                       Optional ByVal useSecondaryYAxis As Boolean = True, _
                       Optional ByVal marginModeX As AxisMarginMode = ammFitBoth, _
                       Optional ByVal marginModeY As AxisMarginMode = ammFitBoth)

    Dim chtEx1stY As ChartExtents
    Dim chtEx2ndY As ChartExtents
    Dim niceX As AxisScale
    Dim niceY1 As AxisScale
    Dim niceY2 As AxisScale
    
    ' 2Y使う/使わないで分岐
    If useSecondaryYAxis And (atu = atuBoth Or atu = atuOnlyY) Then
        ' 2Y使う場合
        chtEx1stY = GetChartExtents(targetChart, aguOnly1stY)
        chtEx2ndY = GetChartExtents(targetChart, aguOnly2ndY)
    Else
        ' 2Y使わない場合
        chtEx1stY = GetChartExtents(targetChart, aguOnly1stY)
    End If
            
    ' Xのみ/Yのみ/XYで分岐
    If atu = atuBoth Or atu = atuOnlyX Then
        ' X
        niceX = GetNiceAxisScale(chtEx1stY.MinX, chtEx1stY.MaxX, , marginModeX)
    End If
    If atu = atuBoth Or atu = atuOnlyY Then
        ' Y1
        niceY1 = GetNiceAxisScale(chtEx1stY.MinY, chtEx1stY.MaxY, , marginModeY)
        ' Y2
        If useSecondaryYAxis Then
            niceY2 = GetNiceAxisScale(chtEx2ndY.MinY, chtEx2ndY.MaxY, , marginModeY)
        End If
    End If
    
    ' X
    If atu = atuBoth Or atuOnlyX Then
        ' X1
        With targetChart.Axes(xlCategory, xlPrimary)
            .MinimumScale = niceX.Min
            .MaximumScale = niceX.Max
            .MajorUnit = niceX.MajorUnit
        End With
        ' X2
        With targetChart.Axes(xlCategory, xlSecondary)
            .MinimumScale = niceX.Min
            .MaximumScale = niceX.Max
            .MajorUnit = niceX.MajorUnit
        End With
    End If
    
    ' Y
    If atu = atuBoth Or atuOnlyY Then
        ' Y1
        With targetChart.Axes(xlValue, xlPrimary)
            .MinimumScale = niceY1.Min
            .MaximumScale = niceY1.Max
            .MajorUnit = niceY1.MajorUnit
        End With
        ' Y2
        If useSecondaryYAxis Then
            With targetChart.Axes(xlValue, xlSecondary)
                .MinimumScale = niceY2.Min
                .MaximumScale = niceY2.Max
                .MajorUnit = niceY2.MajorUnit
            End With
        Else
            With targetChart.Axes(xlValue, xlSecondary)
                .MinimumScale = niceY1.Min
                .MaximumScale = niceY1.Max
                .MajorUnit = niceY1.MajorUnit
            End With
        End If
    End If
    
End Sub

Public Function GetNiceAxisScale(ByVal dataMin As Double, _
                                 ByVal dataMax As Double, _
                                 Optional ByVal targetTicks As Long = 5, _
                                 Optional ByVal MarginMode As AxisMarginMode = ammFitBoth, _
                                 Optional ByVal marginFactor As Double = 0.25) _
                                 As AxisScale
                                 
    Dim result As AxisScale
    Dim range As Double
    Dim rawStep As Double
    Dim niceStep As Double
    Dim exponent As Double
    Dim fraction As Double
    Dim niceMin As Double
    Dim niceMax As Double
    Dim expandAmount As Double
    
    If dataMin = dataMax Then
        If dataMin = 0 Then
            dataMax = 1
        Else
            dataMin = 0
        End If
    End If
    
    range = dataMax - dataMin
    rawStep = range / targetTicks
    If rawStep = 0 Then rawStep = 1
    
    exponent = 10 ^ Int(Log(rawStep) / Log(10))
    fraction = rawStep / exponent
    
    If fraction <= 1 Then
        niceStep = 1
    ElseIf fraction <= 2 Then
        niceStep = 2
    ElseIf fraction <= 5 Then
        niceStep = 5
    Else
        niceStep = 10
    End If
    
    niceStep = niceStep * exponent
    
    niceMin = WorksheetFunction.Floor(dataMin, niceStep)
    niceMax = WorksheetFunction.Ceiling(dataMax, niceStep)
    
    expandAmount = niceStep * marginFactor
    
    Select Case MarginMode
    
        Case ammExpandMax
            niceMax = niceMax + expandAmount
            
        Case ammExpandMin
            niceMin = niceMin - expandAmount
            
        Case ammExpandBoth
            niceMin = niceMin - expandAmount
            niceMax = niceMax + expandAmount
            
    End Select
    
    result.Min = niceMin
    result.Max = niceMax
    result.MajorUnit = niceStep
    
    GetNiceAxisScale = result
    
End Function
