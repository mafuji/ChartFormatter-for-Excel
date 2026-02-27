Attribute VB_Name = "modCommon"
Option Explicit

'---------------------------------------------------------------
' 共通モジュール
'---------------------------------------------------------------

' 色関係の汎用クラス
Public colorPalette As New clsColorPalette

' 第2軸利用のためのダミー系列名
Public Const DUMMY_SERIES_NAME As String = "__DUMMY_SERIES_NAME_106841651__"

' 各種既定値
Public Const DEFAULT_FONT_NAME As String = "Arial Narrow"
Public Const DEFAULT_FONT_SIZE As Single = 16
Public Const DEFAULT_MARKER_SIZE As Integer = 6
Public Const DEFAULT_LINE_WEIGHT As Single = 1.5

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

' 数字オンリーテキストボックスの汎用KeyPress
Public Function IsValidNumericKeyPress( _
    ByVal KeyAscii As Integer, _
    ByVal CurrentText As String, _
    Optional ByVal AllowDecimal As Boolean = False _
) As Boolean

    Select Case KeyAscii
        Case 48 To 57          ' 0?9
            IsValidNumericKeyPress = True

        Case vbKeyBack         ' BackSpace
            IsValidNumericKeyPress = True

        Case 46                ' . (小数点)
            If AllowDecimal Then
                ' 先頭小数点は禁止
                If CurrentText = "" Then
                    IsValidNumericKeyPress = False
                ' 小数点は 1 回だけ
                ElseIf InStr(CurrentText, ".") > 0 Then
                    IsValidNumericKeyPress = False
                Else
                    IsValidNumericKeyPress = True
                End If
            Else
                IsValidNumericKeyPress = False
            End If

        Case Else
            IsValidNumericKeyPress = False
    End Select

End Function

'===========================================================
' Excel列名用 KeyPress 判定（A～Z のみ許可）
'===========================================================
Public Function IsValidColumnKeyPress(ByRef KeyAscii As MSForms.ReturnInteger, ByVal CurrentText As String) As Boolean

    ' 3文字以上は入力不可
    If Len(CurrentText) >= 3 Then
        IsValidColumnKeyPress = False
        Exit Function
    End If

    ' a～z → A～Z に変換
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If

    Select Case KeyAscii
        Case 65 To 90          ' A～Z
            IsValidColumnKeyPress = True

        Case vbKeyBack         ' BackSpace
            IsValidColumnKeyPress = True

        Case Else              ' その他は無効
            IsValidColumnKeyPress = False
    End Select

End Function
