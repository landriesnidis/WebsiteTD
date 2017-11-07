Attribute VB_Name = "CutText"


'==============================================
'Դ�����ƣ�CutText(�ı���ȡģ��)
'Դ��汾��v1.0.6
'Դ��˵������Գ��������ı���ȡ����ĺ���ģ�飬���а�����
'
'         CutText_Single    ������������    ��ȡԴ�ַ����������ض��ַ���֮����ַ���
'         CutText_Multi     ������������    ��ȡ����ַ������Դ��зָ������ַ�������
'
'
'
'==============================================
'Դ�����ߣ�Landriesnidis
'����ʱ�䣺2014-6-24
'��ϵ��ʽ��332007893��QQ��
'�������䣺Landriesnidis@yeah.net
'CSDN����: http://blog.csdn.net/lgj123xj/
'==============================================










Public Function CutText_Single(ByVal SourceStr As String, ByVal StartStr As String, ByVal EndStr As String) As String

'===================CutText_Single======================
'
'��ȡԴ�ַ����������ض��ַ���֮����ַ���
'
'=======================================================

On Error GoTo Error

    Dim n1, n2 As Integer
    
    SourceStr = Right(SourceStr, Len(SourceStr) - InStr(SourceStr, StartStr) + 1)
    n1 = InStr(SourceStr, StartStr)
    n2 = InStr(Right(SourceStr, Len(SourceStr) - n1 - Len(StartStr) + 1), EndStr) + n1 + Len(StartStr) - 1
    CutText_Single = Mid(SourceStr, n1 + Len(StartStr), n2 - n1 - Len(StartStr))
    Exit Function
    
Error:
    CutText_Single = ""
    
End Function




Public Function CutText_Multi(ByVal SourceStr As String, ByVal StartStr As String, ByVal EndStr As String, Delimiter As String) As String

'===================CutText_Multi=======================
'
'��ȡ����ַ������Դ��зָ������ַ�������
'
'=======================================================

On Error GoTo Error

    Dim n1, n2 As Integer
    Dim StrReturn As String                 '���Ա��淵��ֵ
    Dim Cache As String                     '����
    
    Do
        SourceStr = Right(SourceStr, Len(SourceStr) - InStr(SourceStr, StartStr) + 1)
        n1 = InStr(SourceStr, StartStr)
        n2 = InStr(SourceStr, EndStr)
        If n1 = 0 Then Exit Do
        Cache = Mid(SourceStr, n1 + Len(StartStr), n2 - n1 - Len(StartStr))
        If StrReturn <> "" Then StrReturn = StrReturn & Delimiter
        StrReturn = StrReturn & Cache
        SourceStr = Right(SourceStr, Len(SourceStr) - n2)
    Loop
    CutText_Multi = StrReturn
    Exit Function
    
Error:
    CutText_Multi = ""
End Function

