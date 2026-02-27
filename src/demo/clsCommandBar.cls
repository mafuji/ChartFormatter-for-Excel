VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' コマンドバーのプロパティ
'---------------------------------------------------------------

Private no_ As Integer
Private barName_ As String
Private caption_ As String
Private onAction_ As String
Private style_ As MsoButtonStyle
Private faceId_ As Long

'--- No ---
Public Property Get No() As Integer
    No = no_
End Property
Public Property Let No(ByVal newNo As Integer)
    no_ = newNo
    barName_ = BAR_NAME & newNo
End Property

'--- BarName ---
Public Property Get BarName() As String
    BarName = barName_
End Property

'--- Caption ---
Public Property Get Caption() As String
    Caption = caption_
End Property
Public Property Let Caption(ByVal newCaption As String)
    caption_ = newCaption
End Property

'--- OnAction ---
Public Property Get OnAction() As String
    OnAction = onAction_
End Property
Public Property Let OnAction(ByVal newOnAction As String)
    onAction_ = newOnAction
End Property

'--- Style ---
Public Property Get Style() As MsoButtonStyle
    Style = style_
End Property
Public Property Let Style(ByVal newStyle As MsoButtonStyle)
    style_ = newStyle
End Property

'--- FaceId ---
Public Property Get FaceId() As Long
    FaceId = faceId_
End Property
Public Property Let FaceId(ByVal newFaceId As Long)
    faceId_ = newFaceId
End Property

' コマンドバー作成
Public Sub CreateBar()

    Dim bar As CommandBar, btn As CommandBarButton
    
    With Me
        Set bar = Application.CommandBars.Add(Name:=.BarName, Position:=msoBarTop, Temporary:=True)
        Set btn = bar.Controls.Add(Type:=msoControlButton)
        btn.Caption = .Caption
        btn.OnAction = .OnAction
        btn.Style = .Style
        btn.FaceId = .FaceId
        bar.Visible = True
    End With

End Sub
