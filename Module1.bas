Attribute VB_Name = "Module1"
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '该声明得到系统开机到现在的时间(单位：毫秒)
Public Function Sleep(T As Long)
Dim Savetime As Long
Savetime = timeGetTime
While timeGetTime < Savetime + T '循环等待
DoEvents '转让控制权
Wend
End Function
