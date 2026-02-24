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

' プロットエリア設定
Private height_ As Double
Private width_ As Double
Private top_ As Double
Private left_ As Double

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
    
    c.Width = Me.Width
    c.Height = Me.Height
    c.Top = Me.Top
    c.Left = Me.Left
    
    Set Clone = c

End Function

Public Sub ApplyTo(ByRef targetChart As Chart)

    With targetChart.PlotArea
        
        ' 現状固定
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = 1.5
        
        ' リセット
        .Top = 0
        .Left = 0
        .Height = 1
        .Width = 1
        
        ' プロパティ
        .Top = Me.Top
        .Left = Me.Left
        .Height = Me.Height
        .Width = Me.Width
    End With

End Sub
