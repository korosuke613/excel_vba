VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateGraphSanpuzu 
   Caption         =   "�U�z�}�쐬"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "CreateGraphSanpuzu.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CreateGraphSanpuzu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub DisideButton_Click()
    Dim data_range As range
    If RefEdit1.Value <> "" Then
        Set data_range = range(RefEdit1.Value)
        create_graph_sanpuzu data_range
    End If
End Sub

Private Sub RefEdit1_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub
