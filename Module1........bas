Attribute VB_Name = "Module1"
Option Explicit

Sub �ʺA�g��()
Dim deviceName As String
deviceName = InputBox("�п�J�]�ƦW��")
Cells(2, 1).Value = deviceName

Dim modelName As String
modelName = InputBox("�п�J�Ҩ�W��")
Cells(2, 2).Value = modelName

Dim uprice As Integer
uprice = InputBox("�п�J������")
Cells(2, 3).Value = CInt(uprice)

Dim qty As Integer
qty = InputBox("�п�J�ƶq")
Cells(2, 4).Value = CInt(qty)

Dim totalprice As Integer
totalprice = uprice * qty
Cells(2, 5).Value = totalprice



End Sub
