Attribute VB_Name = "CutText"


'==============================================
'源码名称：CutText(文本截取模块)
'源码版本：v1.0.6
'源码说明：针对常见几类文本截取问题的函数模块，其中包含：
'
'         CutText_Single    ………………    截取源字符串中两段特定字符串之间的字符串
'         CutText_Multi     ………………    截取多个字符串，以带有分隔符的字符串传回
'
'
'
'==============================================
'源码作者：Landriesnidis
'发布时间：2014-6-24
'联系方式：332007893（QQ）
'电子邮箱：Landriesnidis@yeah.net
'CSDN博客: http://blog.csdn.net/lgj123xj/
'==============================================










Public Function CutText_Single(ByVal SourceStr As String, ByVal StartStr As String, ByVal EndStr As String) As String

'===================CutText_Single======================
'
'截取源字符串中两段特定字符串之间的字符串
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
'截取多个字符串，以带有分隔符的字符串传回
'
'=======================================================

On Error GoTo Error

    Dim n1, n2 As Integer
    Dim StrReturn As String                 '用以保存返回值
    Dim Cache As String                     '缓存
    
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

