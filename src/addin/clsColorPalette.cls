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

' カラーパレット管理クラス
' 系列番号に応じて色を返す
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
