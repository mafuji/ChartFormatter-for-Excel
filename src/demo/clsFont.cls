VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' グラフのフォント設定
'---------------------------------------------------------------

Private fontName_ As String
Private size_ As Single

Public Property Get FontName() As String
    FontName = fontName_
End Property

Public Property Let FontName(ByVal newFontName As String)
    fontName_ = newFontName
End Property

Public Property Get Size() As Single
    Size = size_
End Property

Public Property Let Size(ByVal newSize As Single)
    size_ = newSize
End Property

Public Function Clone() As clsFont

    Dim c As New clsFont

    c.FontName = Me.FontName
    c.Size = Me.Size

    Set Clone = c

End Function

Public Sub ApplyTo(ByRef targetChart As Chart)
    
    With targetChart.ChartArea.Format.TextFrame2.TextRange.Font
        .NameComplexScript = Me.FontName
        .NameFarEast = Me.FontName
        .Name = Me.FontName
        .Size = Me.Size
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With

End Sub
