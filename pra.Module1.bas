Attribute VB_Name = "Module1"
Sub chatBox()
MsgBox "歡迎來到雲寶寶壽司店"
Dim username As String
username = InputBox("請問你名子")
MsgBox "Hi! " & username & "你好"

Dim username1 As String
username1 = InputBox("請問你想吃什麼?")
MsgBox "好的!馬上為您準備餐點!"

Dim username2 As String
username2 = InputBox("請問還需要其他東西嗎?")
MsgBox "謝謝您!歡迎再次光臨!"

End Sub
