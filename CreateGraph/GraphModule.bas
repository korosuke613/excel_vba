Attribute VB_Name = "GraphModule"
Option Explicit

Sub create_graph(data_range)
    If ActiveSheet.ChartObjects.Count >= 1 Then
        ActiveSheet.ChartObjects.Delete
    End If

    ' �O���t�̍쐬
    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlColumnClustered  ' �_�O���t�Ɛݒ�
        .SetSourceData Source:=data_range  ' �f�[�^�͈͂̎w��
        .HasTitle = True ' �^�C�g����L��
        .ChartTitle.Text = range("B2").Value
        .Axes(xlValue).MaximumScale = 100 '���l���̕ύX
    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = range("J2").Top '��[��ݒ�
        .Left = range("J2").Left '���[��ݒ�
        .Height = 200 '����
        .Width = 300 '��
    End With
    
End Sub

