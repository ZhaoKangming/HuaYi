Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

' 设置鼠标的模拟点击位置
Private Sub Command5_Click()
    SetCursorPos 550, 760   
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub


'借用系统画图板程序和截屏贴图功能锁定模拟鼠标点击的位置
Private Sub MouseClick_Click()
    MsgBox "2秒后捕捉位置"
    TimeDelay (2)   
    SetCursorPos 730, 450 '左上角为坐标原点。往右是x，往下是y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub Command1_Click()
    MsgBox "Start!"
    Shell ("D:\2345Pic\2345PicEditor.exe " & "C:\Users\JokeComing\Desktop\迟海燕\BC.JPG")
    TimeDelay (5)
    MsgBox "END"
End Sub


' 延时程序
Public Sub TimeDelay(ByVal PauseSecond As Single)
    Dim Star, PauseTime
    Star = Timer
    PauseTime = PauseSecond
    Do While Timer < Star + PauseTime
        DoEvents
    Loop
End Sub

Sub OpenPic_Click()
    MsgBox "Start!"
    Shell ("D:\2345Pic\2345PicEditor.exe " & "C:\Users\JokeComing\Desktop\迟海燕\33.JPG")
    TimeDelay (2)
    SetCursorPos 553, 760   '左上角为坐标原点。往右是x，往下是y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    TimeDelay (3)

    ' 模拟键盘输入
    SendKeys "^s"
    TimeDelay (1)
    SendKeys "^v"
    SendKeys "%s"
    SendKeys "%y"
    
    SetCursorPos 730, 450   '保存完确认
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    Shell ("D:\2345Pic\2345PicEditor.exe " & "C:\Users\JokeComing\Desktop\迟海燕\一级医生.JPG")
End Sub
