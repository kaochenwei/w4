Attribute VB_Name = "Module1"
Option Explicit

Sub 動態寫值()
Dim deviceName As String
deviceName = InputBox("請輸入設備名稱")
Cells(2, 1).Value = deviceName

Dim modelName As String
modelName = InputBox("請輸入模具名稱")
Cells(2, 2).Value = modelName

Dim uprice As Integer
uprice = InputBox("請輸入單位價格")
Cells(2, 3).Value = CInt(uprice)

Dim qty As Integer
qty = InputBox("請輸入數量")
Cells(2, 4).Value = CInt(qty)

Dim totalprice As Integer
totalprice = uprice * qty
Cells(2, 5).Value = totalprice



End Sub
