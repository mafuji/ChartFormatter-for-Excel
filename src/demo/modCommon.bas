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
Public Const DEFAULT_FONT_NAME_FAR_EAST As String = "MS UI Gothic"
Public Const DEFAULT_FONT_SIZE As Single = 16
Public Const DEFAULT_MARKER_SIZE As Integer = 6
Public Const DEFAULT_LINE_WEIGHT As Single = 1

' カスタム表示形式
Public Const NUMBER_FORMAT_ZERO_OR_DECIMAL As String = "[=0]0;0.0"

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

' 数字を "1st", "2nd", "3rd", "4th" のような表記に変換
Public Function ToOrdinal(ByVal v As Variant) As String
    Dim n As Long
    Dim suffix As String
    Dim absN As Long
    Dim lastTwo As Long
    Dim lastOne As Long
    Dim sign As String
    
    ' 数値チェック
    If IsNull(v) Or Not IsNumeric(v) Then
        ToOrdinal = ""
        Exit Function
    End If
    
    ' 符号保持（-1 → "-1st" など）
    sign = IIf(CDbl(v) < 0, "-", "")
    
    ' 小数は整数部で判定（例：21.9 → 21st）
    n = Fix(CDbl(v))           ' Fixは0方向に丸め（-1.7→-1）
    absN = Abs(n)
    
    ' 11,12,13 は常に "th"
    lastTwo = absN Mod 100
    lastOne = absN Mod 10
    
    If lastTwo >= 11 And lastTwo <= 13 Then
        suffix = "th"
    Else
        Select Case lastOne
            Case 1: suffix = "st"
            Case 2: suffix = "nd"
            Case 3: suffix = "rd"
            Case Else: suffix = "th"
        End Select
    End If
    
    ToOrdinal = sign & CStr(absN) & suffix
End Function

' シート名被り回避
Public Function UniqueSheetName(ByRef targetBook As Workbook, ByVal orgSheetName As String) As String

    Dim isUnique As Boolean
    Dim duplicateNo As Integer
    Dim tmpName As String

    isUnique = False
    duplicateNo = 0
    Do Until isUnique = True
        tmpName = orgSheetName & IIf(duplicateNo = 0, "", "_" & duplicateNo)
        isUnique = (Not IsSheetExists(targetBook, tmpName))
        duplicateNo = duplicateNo + 1
    Loop
    UniqueSheetName = tmpName

End Function

' シート名存在確認
Public Function IsSheetExists(ByRef targetBook As Workbook, ByVal orgSheetName As String) As Boolean

    Dim ws As Worksheet
    Dim isExists As Boolean
    
    isExists = False
    For Each ws In targetBook.Worksheets
        If ws.Name = orgSheetName Then
            isExists = True
            Exit For
        End If
    Next
    IsSheetExists = isExists

End Function
