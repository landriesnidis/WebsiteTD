Attribute VB_Name = "Wctomb"
Option Explicit

'�ַ�ת��ģ�� Decode the utf-8 text to Chinese


'////////////////////��ȡ��ҳԴ�벢ת��//////////////////

'����Դ��
'Dim Arr_web() As Byte
'Dim Data As String
'Arr_web() = Inet1.OpenURL("http://www.hao123.com/", icByteArray)
'Data = UTF8_Decode(Arr_web())

'///////////////////////////////////////////////////////


'API declartion
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Const CP_UTF8 = 65001

'Decode the utf-8 text to Chinese
Public Function UTF8_Decode(bUTF8() As Byte) As String
    Dim lRet As Long
    Dim lLen As Long
    Dim lBufferSize As Long
    Dim sBuffer As String
    lLen = UBound(bUTF8) + 1
    If lLen = 0 Then Exit Function
    lBufferSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bUTF8(0)), lLen, 0, 0)
    sBuffer = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bUTF8(0)), lLen, StrPtr(sBuffer), lBufferSize)
    UTF8_Decode = sBuffer
End Function


Function g2u(str As String) As String
    Dim i As Long
    Dim arr() As Byte
    arr = StrConv(str, vbFromUnicode)
    For i = LBound(arr) To UBound(arr)
    Next
    'ת��ΪUnicode����
    g2u = UTF8_Decode(arr)
End Function




