VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' カラーパレット
'
'  機能1.系列番号に応じて色を返す
'  機能2.色相グラデーションの作成
'---------------------------------------------------------------

Private colors As Variant

Private Sub Class_Initialize()
    ' 必要に応じて色を追加・変更できる
    colors = Array( _
        RGB(255, 0, 0), _
        RGB(0, 0, 255), _
        RGB(0, 176, 80), _
        RGB(255, 192, 0), _
        RGB(112, 48, 160), _
        RGB(0, 176, 240), _
        RGB(255, 0, 255) _
        )
End Sub

Public Function GetColor(ByVal index As Long) As Long
    Dim n As Long
    n = UBound(colors) - LBound(colors) + 1
    GetColor = colors((index - 1) Mod n)
End Function

Public Function GetColorName(ByVal rgbValue As Long) As String
    
    Select Case rgbValue
    
        Case RGB(0, 0, 0)
            GetColorName = "黒"
            
        Case RGB(255, 0, 0)
            GetColorName = "赤"

        Case RGB(0, 0, 255)
            GetColorName = "青"

        Case RGB(0, 176, 80)
            GetColorName = "緑"

        Case RGB(255, 192, 0)
            GetColorName = "黄"

        Case RGB(112, 48, 160)
            GetColorName = "紫"

        Case RGB(0, 176, 240)
            GetColorName = "水色"

        Case RGB(255, 0, 255)
            GetColorName = "マゼンタ"
            
        Case -1, -2
            GetColorName = "不明"
            
        Case Else
            GetColorName = "ユーザー設定"

    End Select

End Function

Public Function Count() As Long
    Count = UBound(colors) - LBound(colors) + 1
End Function

Public Sub SetColor(ByVal index As Long, ByVal colorValue As Long)
    If index >= 1 And index <= Count Then
        colors(index - 1) = colorValue
    End If
End Sub

Public Sub AddColor(ByVal colorValue As Long)
    Dim tmp As Variant
    Dim i As Long

    ReDim tmp(0 To Count)
    For i = 0 To Count - 1
        tmp(i) = colors(i)
    Next i
    tmp(Count) = colorValue
    colors = tmp
End Sub

' =========================================================
' 紫(既定: Hue=300°)→赤(既定: Hue=0°) を色相グラデーションで n 等分
' 返り値は Long()（RGB関数の戻り値）で、両端色を含めて n 個
'    startHue   開始色相（度）
'    endHue     終了色相（度, 0≒360）
'    sat        彩度 [0..1]
'    val        明度 [0..1]
' =========================================================
Public Function MakeHueGradientColors( _
    ByVal n As Long, _
    Optional ByVal startHue As Double = 300, _
    Optional ByVal endHue As Double = 0, _
    Optional ByVal sat As Double = 1, _
    Optional ByVal val As Double = 1, _
    Optional ByVal shortestPath As Boolean = True _
    ) As Long()

    Dim colors() As Long
    Dim i As Long
    Dim h0 As Double, h1 As Double
    Dim dh As Double, t As Double, h As Double
    Dim r As Long, g As Long, b As Long

    If n <= 0 Then
        ReDim colors(0 To 0)
        colors(0) = RGB(0, 0, 0)
        MakeHueGradientColors = colors
        Exit Function
    End If

    ' n=1 は開始色のみ
    If n = 1 Then
        ReDim colors(0 To 0)
        HSVtoRGB NormalizeHue(startHue), Clamp01(sat), Clamp01(val), r, g, b
        colors(0) = RGB(r, g, b)
        MakeHueGradientColors = colors
        Exit Function
    End If

    h0 = NormalizeHue(startHue)
    h1 = NormalizeHue(endHue)

    If shortestPath Then
        ' 最短経路で回す（例: 300→0 は +60° で 360 に回す）
        dh = h1 - h0
        If dh > 180 Then dh = dh - 360
        If dh < -180 Then dh = dh + 360
    Else
        ' そのままの方向（0→300 は -300° など）
        dh = h1 - h0
    End If

    ReDim colors(0 To n - 1)

    For i = 0 To n - 1
        t = i / (n - 1)                  ' 0..1
        h = NormalizeHue(h0 + dh * t)    ' 補間後の Hue
        HSVtoRGB h, Clamp01(sat), Clamp01(val), r, g, b
        colors(i) = RGB(r, g, b)
    Next i

    MakeHueGradientColors = colors
End Function

' ====== HSV(度,0..1,0..1) → RGB(0..255) 変換 ======
Private Sub HSVtoRGB( _
    ByVal hueDeg As Double, _
    ByVal s As Double, _
    ByVal v As Double, _
    ByRef rOut As Long, _
    ByRef gOut As Long, _
    ByRef bOut As Long)

    Dim c As Double, x As Double, m As Double
    Dim h As Double
    Dim r As Double, g As Double, b As Double
    Dim hp As Double, sector As Long
    Dim hpMod2 As Double

    h = NormalizeHue(hueDeg)
    c = v * s
    hp = h / 60#
    ' ★VBAの Mod は整数演算なので使わない（実数の mod 2 を自前で求める）
    '   hpMod2 = hp - 2 * Int(hp / 2) でもOK（hpは0..6なのでどちらでも同じ）
    hpMod2 = hp - 2# * Fix(hp / 2#)
    x = c * (1 - Abs(hpMod2 - 1))
    m = v - c

    sector = Int(hp) ' 0..5（NormalizeHueによりh∈[0,360)なのでhp∈[0,6)でOK）

    Select Case sector
        Case 0: r = c: g = x: b = 0
        Case 1: r = x: g = c: b = 0
        Case 2: r = 0: g = c: b = x
        Case 3: r = 0: g = x: b = c
        Case 4: r = x: g = 0: b = c
        Case Else ' 5
            r = c: g = 0: b = x
    End Select

    rOut = ToByte((r + m) * 255#)
    gOut = ToByte((g + m) * 255#)
    bOut = ToByte((b + m) * 255#)
End Sub

' 0..255 に丸め
Private Function ToByte(ByVal x As Double) As Long
    If x < 0 Then x = 0
    If x > 255 Then x = 255
    ToByte = CLng(x + 0.5)
End Function

' Hue を 0..360 に正規化（360 は 0 扱い）
Private Function NormalizeHue(ByVal h As Double) As Double
    h = h Mod 360#
    If h < 0 Then h = h + 360#
    NormalizeHue = h
End Function

' 0..1 にクランプ
Private Function Clamp01(ByVal x As Double) As Double
    If x < 0 Then x = 0
    If x > 1 Then x = 1
    Clamp01 = x
End Function
