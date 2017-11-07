VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReceive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "网址查找器 Website TD"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14520
   Icon            =   "frmReceive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListState 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "frmReceive.frx":58C3A
      Left            =   0
      List            =   "frmReceive.frx":58C3C
      TabIndex        =   24
      Top             =   7800
      Width           =   14655
   End
   Begin VB.Frame FrameEd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "源文本编辑器"
      Height          =   735
      Left            =   13560
      TabIndex        =   47
      Top             =   6840
      Width           =   735
      Begin VB.CommandButton CommandTrim 
         Caption         =   "删除空格"
         Height          =   375
         Left            =   12720
         TabIndex        =   58
         Top             =   720
         Width           =   1695
      End
      Begin VB.Timer TimerEd 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   120
         Top             =   240
      End
      Begin VB.CommandButton CommandExit 
         Caption         =   "退出编辑"
         Height          =   375
         Left            =   12720
         TabIndex        =   55
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CommandCls 
         Caption         =   "清空文本"
         Height          =   375
         Left            =   12720
         TabIndex        =   54
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton CommandCopy 
         Caption         =   "全部复制"
         Height          =   375
         Left            =   12720
         TabIndex        =   53
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox CheckUnder 
         BackColor       =   &H00E0E0E0&
         Caption         =   "下划线"
         Height          =   255
         Left            =   12720
         TabIndex        =   52
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox CheckBold 
         BackColor       =   &H00E0E0E0&
         Caption         =   "文字加粗"
         Height          =   255
         Left            =   12720
         TabIndex        =   50
         Top             =   2760
         Width           =   1695
      End
      Begin VB.HScrollBar HScrollSize 
         Height          =   255
         LargeChange     =   2
         Left            =   12720
         Max             =   24
         Min             =   8
         TabIndex        =   49
         Top             =   3480
         Value           =   10
         Width           =   1695
      End
      Begin RichTextLib.RichTextBox TxtEd 
         Height          =   7695
         Left            =   1800
         TabIndex        =   48
         Top             =   100
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   13573
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   5
         TextRTF         =   $"frmReceive.frx":58C3E
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "文字大小："
         Height          =   255
         Left            =   12720
         TabIndex        =   51
         Top             =   3240
         Width           =   1695
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   14280
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame FrameHigh 
      Caption         =   "高级功能"
      Height          =   3255
      Left            =   7920
      TabIndex        =   38
      Top             =   4320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox CheIntactDel 
         Caption         =   "自动删除不完整网址(不推荐)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   62
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox CheIntact 
         Caption         =   "网址完整状态检测"
         Height          =   255
         Left            =   3120
         TabIndex        =   61
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox CheAuto 
         Caption         =   "自动匹配解析字符串"
         Height          =   255
         Left            =   3120
         TabIndex        =   60
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton CommandADD 
         Caption         =   "添加"
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton CommandDel 
         Caption         =   "移除"
         Height          =   375
         Left            =   1800
         TabIndex        =   45
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ListBox ListItem 
         Height          =   1620
         Left            =   360
         TabIndex        =   44
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton CommandDO 
         Caption         =   "执行删除"
         Height          =   375
         Left            =   4800
         TabIndex        =   43
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox ComboStyle 
         Height          =   315
         ItemData        =   "frmReceive.frx":58CD6
         Left            =   1680
         List            =   "frmReceive.frx":58CE0
         TabIndex        =   40
         Text            =   "包含"
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "从列表中删除"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "以下字符串的项目"
         Height          =   255
         Left            =   2880
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "筛选链接解析列表："
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.ComboBox Texts 
      Height          =   300
      ItemData        =   "frmReceive.frx":58CF2
      Left            =   8280
      List            =   "frmReceive.frx":58CFC
      TabIndex        =   25
      Text            =   ">"
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Cmdcls 
      Caption         =   "清空"
      Height          =   375
      Left            =   13080
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Cmdout 
      Caption         =   "生成文本"
      Height          =   375
      Left            =   7920
      TabIndex        =   21
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Caption         =   "输出"
      Height          =   3255
      Left            =   7920
      TabIndex        =   19
      Top             =   4320
      Width           =   6375
      Begin VB.TextBox Text 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "站名与网址解析"
      Height          =   3495
      Left            =   7920
      TabIndex        =   10
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Textaddn 
         Height          =   300
         ItemData        =   "frmReceive.frx":58D07
         Left            =   240
         List            =   "frmReceive.frx":58D1A
         TabIndex        =   57
         Top             =   3000
         Width           =   5895
      End
      Begin VB.ComboBox Textaddb 
         Height          =   300
         ItemData        =   "frmReceive.frx":58D41
         Left            =   240
         List            =   "frmReceive.frx":58D4B
         TabIndex        =   56
         Top             =   2400
         Width           =   5895
      End
      Begin VB.Frame frame 
         Caption         =   "网址"
         Height          =   1575
         Left            =   3240
         TabIndex        =   12
         Top             =   360
         Width           =   2895
         Begin VB.ComboBox Textss 
            Height          =   300
            ItemData        =   "frmReceive.frx":58D65
            Left            =   120
            List            =   "frmReceive.frx":58D6C
            TabIndex        =   28
            Text            =   "<a href="""
            Top             =   480
            Width           =   2535
         End
         Begin VB.ComboBox Textee 
            Height          =   300
            ItemData        =   "frmReceive.frx":58D7B
            Left            =   120
            List            =   "frmReceive.frx":58D82
            TabIndex        =   27
            Text            =   """>"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "结尾字符串："
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "起始字符串："
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "网站名"
         Height          =   1575
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2895
         Begin VB.ComboBox Texte 
            Height          =   300
            ItemData        =   "frmReceive.frx":58D8A
            Left            =   120
            List            =   "frmReceive.frx":58D91
            TabIndex        =   26
            Text            =   "</a>"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label4 
            Caption         =   "结尾字符串："
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "起始字符串："
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label8 
         Caption         =   "网址尾端填充字符串："
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "网址前端填充字符串："
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "链接解析"
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   7455
      Begin VB.CommandButton CommandHigh 
         Caption         =   "高级"
         Height          =   375
         Left            =   6480
         TabIndex        =   37
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton CmdChange 
         Caption         =   "修改"
         Height          =   375
         Left            =   5640
         TabIndex        =   32
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox TextCache 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Text            =   "信息编辑"
         Top             =   4440
         Width           =   5295
      End
      Begin VB.ComboBox Textue 
         Height          =   300
         ItemData        =   "frmReceive.frx":58D9B
         Left            =   1320
         List            =   "frmReceive.frx":58DA2
         TabIndex        =   30
         Text            =   "</a>"
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox Textus 
         Height          =   300
         ItemData        =   "frmReceive.frx":58DAC
         Left            =   1320
         List            =   "frmReceive.frx":58DB6
         TabIndex        =   29
         Text            =   "<a"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton Cmdlisturl 
         Caption         =   "清空"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox Listurl 
         CausesValidation=   0   'False
         Height          =   2985
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   6975
      End
      Begin VB.CommandButton Cmdfind 
         Caption         =   "解析"
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "是否填充"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "是否填充"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "结尾字符串："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "起始字符串："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "源文本"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton CommandClear 
         Caption         =   "清空文本框"
         Height          =   375
         Left            =   5640
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CommandEd 
         Caption         =   "编辑源文本"
         Height          =   375
         Left            =   2040
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CommandCase 
         Caption         =   "字母转小写"
         Height          =   375
         Left            =   3840
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CommandInter 
         Caption         =   "从网页上导入"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Textr 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   720
         Width           =   6975
      End
   End
   Begin VB.CommandButton CommandFormat 
      Caption         =   "设置格式"
      Height          =   375
      Left            =   10440
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CmdCopyall 
      Caption         =   "复制全部"
      Height          =   375
      Left            =   11760
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim SourceStr As String
    Dim StartStr As String
    Dim EndStr As String
    Dim StrReturn As String
    Dim n1, n2 As Integer
    Dim Cache As String                     '缓存
    Dim XH, XHD As Integer
    Dim ListIndex As Integer                '列表选中项
    Dim FrameTop As Integer
    Public PutoutOrder As Integer
    Public PutoutFill As Integer
    
    
    


Private Sub CheAuto_Click()
    If CheAuto.Value = 1 Then
        Texts.Enabled = False
        Texte.Enabled = False
        Textss.Enabled = False
        Textee.Enabled = False
    Else
        Texts.Enabled = True
        Texte.Enabled = True
        Textss.Enabled = True
        Textee.Enabled = True
    End If
End Sub

Private Sub CheckBold_Click()
    If CheckBold.Value = 1 Then TxtEd.Font.Bold = True Else TxtEd.Font.Bold = False
End Sub

Private Sub CheckUnder_Click()
    If CheckUnder.Value = 1 Then TxtEd.Font.Underline = True Else TxtEd.Font.Underline = False
End Sub

Private Sub CheIntact_Click()
    If CheIntact.Value = 1 Then
        CheIntactDel.Enabled = True
    Else
        CheIntactDel.Enabled = False
        CheIntactDel.Value = 0
    End If
End Sub

Private Sub CmdChange_Click()
    If TextCache.Text = "" Or TextCache.Text = "信息编辑" Then
        TextCache.Text = "信息编辑"
    Else
        Listurl.List(ListIndex) = TextCache.Text
        ListState.Clear
        ListState.AddItem (Time & ":" & vbTab & "数据修改成功！")
    End If
End Sub

Private Sub Cmdcls_Click()
    Text.Text = ""
    ListState.Clear
    ListState.AddItem (Time & ":" & vbTab & "输出文本框清除完毕！")
End Sub

Private Sub CmdCopyall_Click()
    Clipboard.Clear
    Clipboard.SetText Text.Text
    ListState.Clear
    ListState.AddItem (Time & ":" & vbTab & "复制成功！")
End Sub

Private Sub Cmdfind_Click()

On Error GoTo Error

    Dim IntactCount As Integer
    Dim ListCountB As Integer

    If Textr.Text = "" Then Exit Sub

    ListState.Clear
    
    ListCountB = Listurl.ListCount
    IntactCount = 0
    SourceStr = Textr.Text
    StartStr = Textus.Text
    EndStr = Textue.Text
    Do
        SourceStr = Right(SourceStr, Len(SourceStr) - InStr(SourceStr, StartStr) + 1)
        n1 = InStr(SourceStr, StartStr)
        n2 = InStr(Right(SourceStr, Len(SourceStr) - n1 - Len(StartStr) + 1), EndStr) + n1 + Len(StartStr) - 1
        If n1 = 0 Then Exit Do
        Cache = Mid(SourceStr, n1 + Len(StartStr), n2 - n1 - Len(StartStr))
        If Check1.Value = 1 Then Cache = StartStr & Cache
        If Check2.Value = 1 Then Cache = Cache & EndStr
        
        If CheIntact.Value = 1 Then             '网址完整检测
            If InStr(Cache, "http") Then IntactCount = IntactCount + 1
        End If
        
        If CheIntactDel.Value <> 1 And InStr(Cache, "http") Then Listurl.AddItem Cache
        
        SourceStr = Right(SourceStr, Len(SourceStr) - n2)
    Loop
    
    If Listurl.ListCount = ListCountB Then
        ListState.AddItem (Time & ":" & vbTab & "未解析出信息！正在尝试自动调整...")
        Textr.Text = LCase(Textr.Text)
        ListState.AddItem (Time & ":" & vbTab & "字符格式调整完毕，重新检测...")
        If InStr(Textr.Text, Textus.Text) Then
            If InStr(Textr.Text, Textue.Text) Then
                ListState.AddItem (Time & ":" & vbTab & "调试成功，请再次执行链接解析！")
            End If
        Else
            ListState.AddItem (Time & ":" & vbTab & "调试失败，请手动设置起始/结尾字符串！")
        End If
        Exit Sub
    End If
    
    
    ListState.AddItem (Time & ":" & vbTab & "解析完成，共 " & Listurl.ListCount & " 项。")
    If IntactCount <> Listurl.ListCount And CheIntact.Value = 1 Then ListState.AddItem (Time & ":" & vbTab & "经检测，共 " & Listurl.ListCount - IntactCount & " 项网址可能不完整，请审核后填充！")
    
    
    If ListSim(Textus, Textus.Text) = False Then Textus.AddItem Textus.Text
    If ListSim(Textue, Textue.Text) = False Then Textue.AddItem Textue.Text
    
    Exit Sub
    
Error:

    ListState.AddItem (Time & ":" & vbTab & "解析失败，发生未知错误，请检查其实字符串及结尾字符串！")
    
    
End Sub

Private Sub Cmdlisturl_Click()
    Listurl.Clear
    TextCache.Text = "信息编辑"
    ListState.Clear
    ListState.AddItem (Time & ":" & vbTab & "链接解析列表已清空！")
End Sub

Private Sub Cmdout_Click()

    Dim StrWebsiteURL As String
    Dim StrWebsiteName As String
    Dim ErrorCount As Integer
    Dim StrB, StrL As String
    Dim StrFill As String
    Dim AutoError As Integer
    
    AutoError = 0
    ErrorCount = 0
    
    ListState.Clear
    
    For XH = 0 To Listurl.ListCount - 1
    
    
        If CheAuto.Value = 1 Then
            If InStr(Listurl.List(XH), "link=""") Then
                    Textss.Text = "link="""
                    Textee.Text = """"
            Else
                If InStr(Listurl.List(XH), "href=""") Then
                    Textss.Text = "href="""
                    Textee.Text = """"
                Else
                    If InStr(Listurl.List(XH), "href='") Then
                        Textss.Text = "href='"
                        Textee.Text = "'"
                    Else
                        If InStr(Listurl.List(XH), "href=") Then
                            Textss.Text = "href="
                            Textee.Text = " "
                        Else
                            If InStr(Listurl.List(XH), "=""") Then
                                Textss.Text = "="""
                                Textee.Text = " "
                            Else
                                AutoError = AutoError + 1
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
            
            If InStr(Listurl.List(XH), """>") Then
                Texts.Text = """>"
                Texte.Text = "</a>"
            Else
                If InStr(Listurl.List(XH), "'>") Then
                    Texts.Text = "'>"
                    Texte.Text = "</a>"
                Else
                    If InStr(Listurl.List(XH), ">") Then
                        Texts.Text = ">"
                        Texte.Text = "</a>"
                    Else
                        AutoError = AutoError + 1
                        Exit For
                    End If
                End If
            End If
            
        End If
    
    
    
    
    
        StrWebsiteName = CutText_Single(Listurl.List(XH), Texts.Text, Texte.Text)
        StrWebsiteURL = CutText_Single(Listurl.List(XH), Textss.Text, Textee.Text)

        If StrWebsiteName = "" Then
            ListState.AddItem (vbTab & vbTab & "Error:" & "第 " & XH & " 项中无法找到网站名！")
            ErrorCount = ErrorCount + 1
        End If
        
        If StrWebsiteURL = "" Then
            ListState.AddItem (vbTab & vbTab & "Error:" & "第 " & XH & " 项中无法找到网址！")
            ErrorCount = ErrorCount + 1
        End If
        
        Select Case PutoutFill          '间隙填充字符
            Case 0
                StrFill = vbTab
            Case 1
                StrFill = vbCrLf
            Case 2
                StrFill = ""
            Case 3
                StrFill = frmFormating.TextFill.Text
        End Select
        
        StrWebsiteURL = Textaddb.Text & StrWebsiteURL & Textaddn.Text
        
        If PutoutOrder = 0 Then
            StrB = StrWebsiteName
            StrL = StrWebsiteURL
        Else
            StrL = StrWebsiteName
            StrB = StrWebsiteURL
        End If
        
        Text.Text = Text.Text & StrB & StrFill & StrL & vbCrLf
        
    Next XH
    If Listurl.ListCount > 0 Then ListState.AddItem (Time & ":" & vbTab & "输出完毕，共 " & Listurl.ListCount & " 行，错误 " & ErrorCount & " 个！")
    
    '将历史加入列表
    If ListSim(Texts, Texts.Text) = False Then Texts.AddItem Texts.Text
    If ListSim(Texte, Texte.Text) = False Then Texte.AddItem Texte.Text
    If ListSim(Textss, Textss.Text) = False Then Textss.AddItem Textss.Text
    If ListSim(Textee, Textee.Text) = False Then Textee.AddItem Textee.Text
    
    
End Sub


Private Sub ComboStyle_Change()
    If ComboStyle.Text <> "包含" And ComboStyle.Text <> "不包含" Then
        ComboStyle.Text = "包含"
    End If
End Sub





Private Sub CommandCls_Click()
    TxtEd.Text = ""
End Sub

Private Sub CommandCopy_Click()
    Clipboard.Clear
    Clipboard.SetText TxtEd.Text
    ListState.Clear: ListState.AddItem (Time & ":" & vbTab & "复制成功！")
End Sub

Private Sub CommandEd_Click()
    TxtEd.Text = Textr.Text
    TimerEd.Enabled = True
End Sub

Private Sub CommandExit_Click()
    Textr.Text = TxtEd.Text
    TimerEd.Enabled = True
End Sub

Private Sub CommandFormat_Click()
    Load frmFormating
End Sub

Private Sub CommandHigh_Click()
    If FrameHigh.Visible = True Then FrameHigh.Visible = False Else FrameHigh.Visible = True
End Sub

Private Sub CommandADD_Click()
    ListItem.AddItem InputBox("请输入要添加的字符串：")
End Sub

Private Sub CommandClear_Click()
    Textr.Text = ""
End Sub

Private Sub CommandDel_Click()
On Error GoTo Error
    ListItem.RemoveItem (ListItem.ListIndex)
    Exit Sub
Error:
    
End Sub

Private Sub CommandDO_Click()

On Error GoTo Error

    If ComboStyle.Text = "包含" Then
        For XH = 0 To Listurl.ListCount - 1
            For XHD = 0 To ListItem.ListCount - 1
                If InStr(Listurl.List(XH), ListItem.List(XHD)) Then
                    Listurl.RemoveItem (XH)
                    XH = XH - 1
                End If
            Next XHD
        Next XH
    Else
        For XH = 0 To Listurl.ListCount - 1
            For XHD = 0 To ListItem.ListCount - 1
                If InStr(Listurl.List(XH), ListItem.List(XHD)) Then Exit For
                If XHD = ListItem.ListCount - 1 Then
                    Listurl.RemoveItem (XH)
                    XH = XH - 1
                End If
            Next XHD
        Next XH
    End If
    Exit Sub
         
Error:
    Call CommandDO_Click
End Sub

Private Sub CommandInter_Click()
    Dim StrWeb As String
    Dim Arr_web() As Byte
    
    StrWeb = InputBox("请输入目标网站的网址：")
    If Len(StrWeb) = 0 Then Exit Sub
    
    Arr_web() = Inet1.OpenURL(StrWeb, icByteArray)
    Textr.Text = UTF8_Decode(Arr_web())
End Sub

Private Sub CommandCase_Click()
    If CommandCase.Caption = "字母转小写" Then
        Textr.Text = LCase(Textr.Text)
        CommandCase.Caption = "字母转大写"
    ElseIf CommandCase.Caption = "字母转大写" Then
        Textr.Text = UCase(Textr.Text)
        CommandCase.Caption = "字母转小写"
    End If
End Sub

Private Sub CommandTrim_Click()
    TxtEd.Text = Trim(TxtEd.Text)
End Sub

Private Sub Form_Load()
    Me.Caption = App.ProductName & "    v" & App.Major & "." & App.Minor & "." & App.Revision & "   -----   " & App.CompanyName
    FrameEd.Height = ListState.Top
    FrameEd.Width = Me.Width
    FrameEd.Left = 0
    FrameEd.Top = FrameEd.Height * -1
    
    ListState.Top = Frame2.Top + Frame2.Height + 250
    ListState.Height = Me.Height - ListState.Top
    ListState.Left = 0
    ListState.Width = Me.Width
    ListState.AddItem (Date)
    ListState.AddItem (Time & ":" & vbTab & "欢迎使用 " & Me.Caption)
End Sub

Private Sub HScrollSize_Change()
    TxtEd.Font.Size = HScrollSize.Value
End Sub

Private Sub Listurl_Click()
    ListIndex = Listurl.ListIndex
    TextCache.Text = Listurl.Text
End Sub

Public Function ListSim(Obj As Object, Item As String) As Boolean   '产看列表中是否存在指定字符串

    For XH = 0 To Obj.ListCount - 1
        If Item = Obj.List(XH) Then
            ListSim = True
            Exit Function
        End If
    Next XH
    
    ListSim = False

End Function


Private Sub TimerEd_Timer()
    FrameEd.Top = FrameEd.Top - (FrameEd.Top - FrameTop) / 8
    If Abs(FrameEd.Top - FrameTop) < 10 Then
        FrameEd.Top = FrameTop
        If FrameTop = 0 Then FrameTop = FrameEd.Height * -1 Else FrameTop = 0
        TimerEd.Enabled = False
    End If
End Sub
