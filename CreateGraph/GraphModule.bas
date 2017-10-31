Attribute VB_Name = "GraphModule"
Option Explicit

Sub create_graph(data_range)
    If ActiveSheet.ChartObjects.Count >= 1 Then
        ActiveSheet.ChartObjects.Delete
    End If

    ' グラフの作成
    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlColumnClustered  ' 棒グラフと設定
        .SetSourceData Source:=data_range  ' データ範囲の指定
        .HasTitle = True ' タイトルを有効
        .ChartTitle.Text = range("B2").Value
        .Axes(xlValue).MaximumScale = 100 '数値軸の変更
        'グラフのX軸(横軸)のタイトルを設定
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "名前"
        ' グラフのY軸(縦軸)のタイトルを設定
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "点数"

    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = range("J2").Top '上端を設定
        .Left = range("J2").Left '左端を設定
        .Height = 200 '高さ
        .Width = 300 '幅
    End With
    
End Sub

