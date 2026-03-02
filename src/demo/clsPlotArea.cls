VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlotArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' プロットエリア設定
'---------------------------------------------------------------

Private height_ As Double
Private width_ As Double
Private top_ As Double
Private left_ As Double

Public Property Let height(ByVal newHeight As Double)
    height_ = newHeight
End Property

Public Property Get height() As Double
    height = height_
End Property

Public Property Let width(ByVal newWidth As Double)
    width_ = newWidth
End Property

Public Property Get width() As Double
    width = width_
End Property

Public Property Let Left(ByVal newLeft As Double)
    left_ = newLeft
End Property

Public Property Get Left() As Double
    Left = left_
End Property

Public Property Let Top(ByVal newTop As Double)
    top_ = newTop
End Property

Public Property Get Top() As Double
    Top = top_
End Property

Public Function Clone() As clsPlotArea
    
    Dim c As New clsPlotArea
    
    c.width = Me.width
    c.height = Me.height
    c.Top = Me.Top
    c.Left = Me.Left
    
    Set Clone = c

End Function

Public Sub ApplyTo(ByRef targetChart As Chart)

    With targetChart.PlotArea
        
        ' 固定
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = 1.5
        
        WaitUntilStable targetChart
        
        ' リセット
        .InsideTop = 1
        .InsideLeft = 1
        .InsideHeight = 1
        .InsideWidth = 1
        
        WaitUntilStable targetChart
        
        ' プロパティ
        .InsideTop = Me.Top
        .InsideLeft = Me.Left
        .InsideHeight = Me.height
        .InsideWidth = Me.width
    End With

End Sub

Private Sub WaitUntilStable(ByRef targetChart As Chart)

    ' レイアウト確定まで待機(max 3s)
    Dim dbl As Double
    
    dbl = Timer
    Do Until targetChart.PlotArea.InsideTop <> 0 Or Timer - dbl > 3
        Debug.Print "waiting"
    Loop

End Sub
