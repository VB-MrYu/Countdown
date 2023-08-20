VERSION 5.00
Begin VB.Form Frm 
   Caption         =   "倒计时"
   ClientHeight    =   3135
   ClientLeft      =   8505
   ClientTop       =   5490
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "秒数："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Timer1.Enabled = True Then
tcqr
Else
End
End If
Cancel = True
End Sub

Private Sub tcqr()
If MsgBox("倒计时还未结束，您真的要关闭吗？", vbYesNo) = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub Command1_Click()
If Text = "" Then Exit Sub
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
Text = 0
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Text.Text = "0" Then
Timer2.Enabled = True
Timer1.Enabled = False
End If
Text.Locked = False
Text = Text - 1
Text.Locked = True
End Sub

Private Sub Timer2_Timer()
Text = 0
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Text.Locked = False
MsgBox "倒计时已结束！", 64 + 4096
Timer2.Enabled = False
End Sub
