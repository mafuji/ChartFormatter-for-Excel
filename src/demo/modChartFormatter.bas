Attribute VB_Name = "modChartFormatter"
Option Explicit

'---------------------------------------------------------------
' エントリーポイント
'---------------------------------------------------------------

Sub ShowChartFormatForm()

    ' 選択状態チェック
    Dim targetChart As Chart
    Dim scatterSelected As Boolean
    Dim s As Series
    
    Set targetChart = GetChart(Selection)
    If targetChart Is Nothing Then
        scatterSelected = False
    Else
        scatterSelected = IsScatterOnlyChart(targetChart)
    End If

    If Not scatterSelected Then
        MsgBox "散布図を選択した状態でボタンを押してください。", vbExclamation
        Exit Sub
    End If
    
    Set frmScatterChartFormat.targetChart = targetChart
    frmScatterChartFormat.Show

End Sub

Sub ShowChargeDischargeChartBuilder()

    frmChargeDischargeChartBuilder.Show vbModal

End Sub

