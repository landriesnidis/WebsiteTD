VERSION 5.00
Begin VB.Form frmFormating 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������ʽ"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   Icon            =   "frmFormatting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   3495
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "��϶���"
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
      Begin VB.TextBox TextFill 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Fill 
         Caption         =   "�Զ���"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Fill 
         Caption         =   "�����"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Fill 
         Caption         =   "���з���Enter��"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Fill 
         Caption         =   "�Ʊ����Tab��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���˳��"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton Order 
         Caption         =   "��ַ��ǰվ���ں�"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Order 
         Caption         =   "վ����ǰ��ַ�ں�"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmFormating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fill_Click(Index As Integer)
    If Fill(3).Value = True Then TextFill.Enabled = True Else TextFill.Enabled = False
    frmReceive.PutoutFill = Index
End Sub

Private Sub Form_Load()
    Me.Height = frmReceive.Height
    Me.Top = frmReceive.Top
    Me.Left = frmReceive.Left + frmReceive.Width
    Me.Show
End Sub

Private Sub Order_Click(Index As Integer)
    frmReceive.PutoutOrder = Index
End Sub
