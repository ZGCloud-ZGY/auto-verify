Attribute VB_Name = "Module1"
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '�������õ�ϵͳ���������ڵ�ʱ��(��λ������)
Public Function Sleep(T As Long)
Dim Savetime As Long
Savetime = timeGetTime
While timeGetTime < Savetime + T 'ѭ���ȴ�
DoEvents 'ת�ÿ���Ȩ
Wend
End Function
