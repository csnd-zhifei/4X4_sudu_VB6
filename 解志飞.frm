VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   11070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   5550
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   450
      TabIndex        =   1
      Top             =   450
      Width           =   9915
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
     Label1.Caption = Space(7) & "苏度库4X4（澳大利亚-单人棋）" & vbCrLf & vbCrLf & _
     "游戏规则：游戏目标是将四套1~4号数字棋子放置于棋盘正中间的4X4的方阵中，并使1~4号数字中的每个数字在棋盘的每一行、每一列以及四个2X2的方阵中都只出现一次。游戏过程中16枚棋子可处于棋盘上的任意位置，通过移动棋块，使各行、列、各2X2的方阵中的数字都达到规则的要求，即为成功。"
End Sub

