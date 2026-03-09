VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColorBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
' カラーバークラス
'---------------------------------------------------------------

' ==========================
' プライベート変数
' ==========================
Private pColors() As Long
Private pParent As Worksheet

Private pLeft As Double
Private pTop As Double
Private pWidth As Double
Private pLength As Double

Private pReverse As Boolean

Private pGroupName As String

' ラベル設定
Private pLabelPosition As LabelPosition
Private pLabelMode As LabelMode
Private pLabelOffset As Double
Private pLabelUnit As Long ' ラベルの表示間隔：lmCustomの時のみ使用
Private pLabelFormat As LabelFormat
Private pIsColorfulLabel As Boolean

' フォントクラス
Private pLabelFont As clsFont

' 図形名保持
Private pShapeNames As Collection

' ==========================
' 初期化
' ==========================
Private Sub Class_Initialize()
    pReverse = False
    pLabelOffset = 4
    
    Set pShapeNames = New Collection
    
    ' デフォルトフォント設定
    Set pLabelFont = New clsFont
    pLabelFont.FontName = DEFAULT_FONT_NAME
    pLabelFont.Size = DEFAULT_FONT_SIZE
End Sub

' ==========================
' 初期化関数
' ==========================
Public Sub Initialize( _
        ByRef colors() As Long, _
        ByVal parentSheet As Worksheet, _
        ByVal leftPos As Double, _
        ByVal topPos As Double)

    pColors = colors
    Set pParent = parentSheet
    
    pLeft = leftPos
    pTop = topPos
End Sub

' ==========================
' プロパティ
' ==========================
Public Property Let Width(ByVal value As Double)
    pWidth = value
End Property

Public Property Get Width() As Double
    Width = pWidth
End Property

Public Property Let Length(ByVal value As Double)
    pLength = value
End Property

Public Property Get Length() As Double
    Length = pLength
End Property

Public Property Let Reverse(ByVal value As Boolean)
    pReverse = value
End Property

Public Property Get Reverse() As Boolean
    Reverse = pReverse
End Property

Public Property Let LabelPosition(ByVal value As LabelPosition)
    pLabelPosition = value
End Property

Public Property Get LabelPosition() As LabelPosition
    LabelPosition = pLabelPosition
End Property

Public Property Let LabelMode(ByVal value As LabelMode)
    pLabelMode = value
End Property

Public Property Get LabelMode() As LabelMode
    LabelMode = pLabelMode
End Property

Public Property Get LabelFont() As clsFont
    Set LabelFont = pLabelFont
End Property

Public Property Set LabelFont(ByVal value As clsFont)
    Set pLabelFont = value
End Property

Public Property Let LabelOffset(ByVal value As Double)
    pLabelOffset = value
End Property

Public Property Get LabelOffset() As Double
    LabelOffset = pLabelOffset
End Property

Public Property Let LabelUnit(ByVal value As Long)
    pLabelUnit = value
End Property

Public Property Get LabelUnit() As Long
    LabelUnit = pLabelUnit
End Property

Public Property Let LabelFormat(ByVal value As LabelFormat)
    pLabelFormat = value
End Property

Public Property Get LabelFormat() As LabelFormat
    LabelFormat = pLabelFormat
End Property

Public Property Let IsColorfulLabel(ByVal value As Boolean)
    pIsColorfulLabel = value
End Property

Public Property Get IsColorfulLabel() As Boolean
    IsColorfulLabel = pIsColorfulLabel
End Property

Public Property Get GroupName() As String
    GroupName = pGroupName
End Property

' ==========================
' 描画（バー＋ラベル） ※グループ化対応
' ==========================
Public Sub Draw()

    Dim baseIndex As Long
    Dim n As Long
    Dim i As Long
    Dim idx As Long
    Dim segSize As Double
    Dim shp As Shape
    Dim lbl As Shape
    Dim ln As Shape
    Dim colorIndex As Long
    Dim labelText As String
    Dim x As Double, y As Double
    Dim arr() As String
    Dim k As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim L As Double: L = Application.CentimetersToPoints(0.1) ' 黒線の長さ

    If pParent Is Nothing Then Exit Sub

    ' --- 既存削除 ---
    Delete
    Set pShapeNames = New Collection

    baseIndex = LBound(pColors)
    n = UBound(pColors) - baseIndex + 1
    If n <= 0 Then Exit Sub

    segSize = pLength / n

    For i = 0 To n - 1

        idx = i

        ' ===== 色インデックス =====
        If pReverse Then
            colorIndex = UBound(pColors) - idx
        Else
            colorIndex = baseIndex + idx
        End If

        ' ===== バー描画 =====
        If pLabelPosition = lpLeft Or pLabelPosition = lpRight Then
            Set shp = pParent.Shapes.AddShape( _
                        msoShapeRectangle, _
                        pLeft, _
                        pTop + idx * segSize, _
                        pWidth, _
                        segSize + 0.5)
        Else
            Set shp = pParent.Shapes.AddShape( _
                        msoShapeRectangle, _
                        pLeft + idx * segSize, _
                        pTop, _
                        segSize + 0.5, _
                        pWidth)
        End If

        shp.Fill.ForeColor.RGB = pColors(colorIndex)
        shp.Line.Visible = msoFalse
        pShapeNames.Add shp.Name

        ' ===== ラベル描画 =====
        If pLabelMode = lmAll _
        Or (pLabelMode = lmMinMax And (i = 0 Or i = n - 1)) _
        Or (pLabelMode = lmCustom And (i = 0 Or i = n - 1 Or Int(CCur(i + 1) / CCur(pLabelUnit)) = CCur(i + 1) / CCur(pLabelUnit))) Then

            Select Case pLabelFormat
                Case lfCardinal
                    labelText = CStr(i + 1)
                Case lfOrdinal
                    labelText = ToOrdinal(i + 1)
            End Select

            If pLabelPosition = lpLeft Or pLabelPosition = lpRight Then
                y = pTop + i * segSize + segSize / 2
                If pLabelPosition = lpLeft Then
                    x = pLeft - pLabelOffset - 30
                Else
                    x = pLeft + pWidth + pLabelOffset
                End If
                Set lbl = pParent.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y - 8, 30, 16)
                If pLabelPosition = lpLeft Then
                    lbl.TextFrame.HorizontalAlignment = xlHAlignRight
                Else
                    lbl.TextFrame.HorizontalAlignment = xlHAlignLeft
                End If
                lbl.TextFrame.VerticalAlignment = xlVAlignCenter
            Else
                x = pLeft + i * segSize + segSize / 2
                If pLabelPosition = lpTop Then
                    y = pTop - 16 - pLabelOffset
                Else
                    y = pTop + pWidth + pLabelOffset
                End If
                Set lbl = pParent.Shapes.AddTextbox(msoTextOrientationHorizontal, x - 15, y, 30, 16)
                lbl.TextFrame.HorizontalAlignment = xlHAlignCenter
                lbl.TextFrame.VerticalAlignment = xlVAlignCenter
            End If

            lbl.Fill.Visible = msoFalse
            lbl.TextFrame.Characters.Text = labelText

            If Not pLabelFont Is Nothing Then
                lbl.TextFrame.Characters.Font.Name = pLabelFont.FontName
                lbl.TextFrame.Characters.Font.Size = pLabelFont.Size
            End If

            lbl.TextFrame2.WordWrap = msoFalse
            lbl.TextFrame2.AutoSize = msoAutoSizeShapeToFitText

            If IsColorfulLabel Then
                lbl.TextFrame.Characters.Font.Color = pColors(i)
            Else
                lbl.TextFrame.Characters.Font.Color = vbBlack
            End If
            lbl.Line.Visible = msoFalse
            pShapeNames.Add lbl.Name

            ' ===== 黒線の追加 =====
            Select Case pLabelPosition
                Case lpLeft
                    ' 右辺中央 → 左端がそこに来る水平線
                    x1 = shp.Left
                    y1 = shp.Top + shp.Height / 2
                    x2 = x1 - L
                    y2 = y1

                Case lpRight
                    ' 左辺中央 → 右端がそこに来る水平線
                    x1 = shp.Left + shp.Width
                    y1 = shp.Top + shp.Height / 2
                    x2 = x1 + L
                    y2 = y1
                    
                Case lpTop
                    ' 上辺中央 → 下端がそこに来る垂直線
                    x1 = shp.Left + shp.Width / 2
                    y1 = shp.Top
                    x2 = x1
                    y2 = y1 - L

                Case lpBottom
                    ' 下辺中央 → 上端がそこに来る垂直線
                    x1 = shp.Left + shp.Width / 2
                    y1 = shp.Top + shp.Height
                    x2 = x1
                    y2 = y1 + L
            End Select

            Set ln = pParent.Shapes.AddLine(x1, y1, x2, y2)
            ln.Line.ForeColor.RGB = vbBlack
            ln.Line.Weight = 1.5
            pShapeNames.Add ln.Name

        End If

    Next i

    ' ===== 最終グループ化 =====
    ReDim arr(1 To pShapeNames.Count)
    For k = 1 To pShapeNames.Count
        arr(k) = pShapeNames(k)
    Next k

    pGroupName = pParent.Shapes.range(arr).Group.Name

End Sub

' ==========================
' 削除
' ==========================
Public Sub Delete()
    If pGroupName <> "" Then
        On Error Resume Next
        pParent.Shapes(pGroupName).Delete
        On Error GoTo 0
        pGroupName = ""
    End If
End Sub

