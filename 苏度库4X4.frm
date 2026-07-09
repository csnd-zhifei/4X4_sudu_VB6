VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "苏度库4X4"
   ClientHeight    =   9930
   ClientLeft      =   1650
   ClientTop       =   1695
   ClientWidth     =   11325
   DrawWidth       =   2
   FillColor       =   &H008080FF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   26.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11325
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8550
      TabIndex        =   3
      Text            =   "a7-001"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "说明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "布子"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7050
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Label 块 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const 空棋位 = 0
Const 格大小 = 1040         ' 正方形
Const 细间隙 = 6            ' 细线间隙
Const 线粗 = 4              ' 外框/中线粗细
Const 粗间隙 = 16           ' 粗线间隙
Const 格内边距 = 50         ' 格子内边距
Const 选择区颜色 = &H8080FF

Public sf As String         ' 当前选中的数字
Dim 棋位(1 To 4, 1 To 4) As Integer
Dim 题目数(1 To 4) As Integer
Dim 已放数(1 To 4) As Integer
Dim 选中来源 As Integer      ' 0=无 1=选择区 2=棋盘
Dim 选中索引 As Integer      ' 棋盘索引

Private Sub Form_Load()
    Dim i As Integer
    Dim 选择器顶 As Single
    Dim 步长 As Single
    Dim 棋盘左 As Single, 棋盘顶 As Single
    Dim 题库文件 As String
    Dim 选择区宽 As Single, 选择区左 As Single
    Dim 棋盘宽 As Single, 棋盘左对齐 As Single
    Dim 单格左 As Single, 单格顶 As Single
    
    Me.AutoRedraw = True
    Me.Caption = "苏度库4X4（澳大利亚-单人棋）"
    
    ' 初始化控件布局与题库列表
    For i = 1 To 19
        Load 块(i)
    Next i
    
    步长 = 格大小 + 细间隙
    选择区宽 = 4 * 格大小 + 3 * 细间隙
    选择区左 = (Me.ScaleWidth - 选择区宽) / 2
    
    ' 顶部数字选择区
    选择器顶 = Command1.Top + Command1.Height + 360
    For i = 16 To 19
        With 块(i)
            单格左 = 选择区左 + (i - 16) * 步长
            单格顶 = 选择器顶
            .Width = 格大小 - 2 * 格内边距
            .Height = 格大小 - 2 * 格内边距
            .Left = 单格左 + 格内边距
            .Top = 单格顶 + 格内边距
            .Alignment = vbCenter
            .BorderStyle = 1
            .Font.Size = 18
            .Font.Bold = True
            .BackColor = vbWhite
            .ForeColor = vbBlack
            .Caption = i - 15
            .Tag = 3
            .Visible = True
        End With
    Next i
    
    ' 棋盘区域
    棋盘宽 = 4 * 格大小 + 2 * 细间隙 + 粗间隙
    棋盘左对齐 = (Me.ScaleWidth - 棋盘宽) / 2
    棋盘左 = 棋盘左对齐
    棋盘顶 = 选择器顶 + 格大小 + 1250
    
    For i = 0 To 15
        With 块(i)
            单格左 = 棋盘左 + (i Mod 4) * 格大小 + (i Mod 4) * 细间隙 + IIf((i Mod 4) + 1 > 2, 粗间隙 - 细间隙, 0)
            单格顶 = 棋盘顶 + (i \ 4) * 格大小 + (i \ 4) * 细间隙 + IIf((i \ 4) + 1 > 2, 粗间隙 - 细间隙, 0)
            .Width = 格大小 - 2 * 格内边距
            .Height = 格大小 - 2 * 格内边距
            .Left = 单格左 + 格内边距
            .Top = 单格顶 + 格内边距
            .Alignment = vbCenter
            .BorderStyle = 0
            .Font.Size = 18
            .Font.Bold = True
            .BackColor = Me.BackColor
            .ForeColor = vbBlack
            .Caption = ""
            .Tag = 0
            .Visible = True
        End With
    Next i
    
    Combo1.Clear
    题库文件 = Dir(App.Path & "\01-7.9元数独题\*.txt")
    Do While 题库文件 <> ""
        Combo1.AddItem Replace(题库文件, ".txt", "")
        题库文件 = Dir ' 找下一个
    Loop

    Label1.Caption = ""
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    Else
        MsgBox "未找到题库文件，请确保程序同级目录下有 [01-7.9元数独题] 文件夹！", vbExclamation
    End If
    重绘
End Sub

' --------------------------------------------------
' Combo1：切换题目
' --------------------------------------------------
Private Sub Combo1_Click()
    布棋
End Sub

' --------------------------------------------------
' Command1：重布棋 (重新开始本局)
' --------------------------------------------------
Private Sub Command1_Click()
    重布棋
End Sub

Private Sub Command2_Click()
    Form2.Show vbModal
End Sub

Private Sub 块_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Integer, c As Integer
    Dim numVal As Integer
    Dim 冲突提示 As String
    
    ' 选择区点击：拿起数字
    If Index >= 16 And Index <= 19 Then
        If 块(Index).Caption = "" Then Exit Sub
        sf = 块(Index).Caption
        选中来源 = 1
        选中索引 = -1
        Label1.Caption = "已选择: " & sf
        Exit Sub
    End If
    
    ' 棋盘区点击：选择或落子
    If Index >= 0 And Index <= 15 Then
        If 块(Index).Tag = 1 Then Exit Sub
        
        r = Index \ 4 + 1
        c = Index Mod 4 + 1
        
        If 块(Index).Tag = 2 Then
            sf = 块(Index).Caption
            选中来源 = 2
            选中索引 = Index
            Label1.Caption = "已选择: " & sf
            Exit Sub
        End If

        If sf <> "" Then
            numVal = Val(sf)

            If 选中来源 = 2 And 选中索引 >= 0 Then
                If Index <> 选中索引 Then
                    棋位(选中索引 \ 4 + 1, 选中索引 Mod 4 + 1) = 空棋位
                    清空块 选中索引
                End If
            ElseIf numVal >= 1 And numVal <= 4 Then
                已放数(numVal) = 已放数(numVal) + 1
            End If

            设置块 Index, numVal, 2, vbWhite
            棋位(r, c) = numVal

            更新选择器

            冲突提示 = 检查冲突提示(r, c, numVal)
            Label1.Caption = 冲突提示

            清空选择

            If 已填满() Then 判断结果
        End If
    End If
End Sub

Sub 布棋()
    Dim fileName As String, s As String
    Dim ch As String
    Dim i As Integer, j As Integer, k As Integer
    Dim idx As Integer, a As Integer

    ' 读取题目文件并布置固定数字
    重置状态
    For i = 0 To 15
        清空块 i
    Next i
    
    fileName = App.Path & "\01-7.9元数独题\" & Combo1.Text & ".txt"
    If Dir(fileName) = "" Then Exit Sub
    
    On Error GoTo FileErr
    Open fileName For Binary As #1
    s = StrConv(InputB(LOF(1), 1), vbUnicode)
    Close #1
    On Error GoTo 0
    
    k = 1
    For i = 1 To 4
        For j = 1 To 4
            ch = ""
            Do While k <= Len(s)
                ch = Mid(s, k, 1)
                k = k + 1
                If ch >= "0" And ch <= "9" Then Exit Do
            Loop

            a = Val(ch)
            idx = (i - 1) * 4 + (j - 1)
            
            If a <> 0 Then
                设置块 idx, a, 1, &HFFFF&
                棋位(i, j) = a
                If a >= 1 And a <= 4 Then
                    题目数(a) = 题目数(a) + 1
                End If
            End If
        Next j
    Next i

    更新选择器
    重绘
    Exit Sub
    
FileErr:
    MsgBox "无法读取题目文件：" & fileName, vbExclamation, "错误"
End Sub

Sub 重布棋()
    Dim i As Integer, r As Integer, c As Integer
    Dim numVal As Integer
    
    ' 清空玩家落子，保留题目
    For i = 0 To 15
        If 块(i).Tag = 2 Then
            numVal = Val(块(i).Caption)
            清空块 i
            r = i \ 4 + 1
            c = i Mod 4 + 1
            棋位(r, c) = 空棋位
            If numVal >= 1 And numVal <= 4 Then
                已放数(numVal) = 已放数(numVal) - 1
                If 已放数(numVal) < 0 Then 已放数(numVal) = 0
            End If
        End If
    Next i

    清空选择
    更新选择器
    Label1.Caption = ""
    重绘
End Sub

Sub 判断结果()
    Dim i As Integer, j As Integer
    
    For i = 1 To 4
        If 有重复(棋位(i, 1), 棋位(i, 2), 棋位(i, 3), 棋位(i, 4)) Then
            Label1.Caption = "请检查第 " & i & " 行，有重复！"
            Exit Sub
        End If
    Next i

    For j = 1 To 4
        If 有重复(棋位(1, j), 棋位(2, j), 棋位(3, j), 棋位(4, j)) Then
            Label1.Caption = "请检查第 " & j & " 列，有重复！"
            Exit Sub
        End If
    Next j

    If 有重复(棋位(1, 1), 棋位(1, 2), 棋位(2, 1), 棋位(2, 2)) Then
        Label1.Caption = "左上宫格有重复！"
        Exit Sub
    End If
    If 有重复(棋位(1, 3), 棋位(1, 4), 棋位(2, 3), 棋位(2, 4)) Then
        Label1.Caption = "右上宫格有重复！"
        Exit Sub
    End If
    If 有重复(棋位(3, 1), 棋位(3, 2), 棋位(4, 1), 棋位(4, 2)) Then
        Label1.Caption = "左下宫格有重复！"
        Exit Sub
    End If
    If 有重复(棋位(3, 3), 棋位(3, 4), 棋位(4, 3), 棋位(4, 4)) Then
        Label1.Caption = "右下宫格有重复！"
        Exit Sub
    End If

    Label1.Caption = "\^o^/恭喜你答对了！"
End Sub

Function 有重复(a As Integer, b As Integer, c As Integer, d As Integer) As Boolean
    If a = b Or a = c Or a = d Or b = c Or b = d Or c = d Then
        有重复 = True
    Else
        有重复 = False
    End If
End Function

Private Sub Form_Paint()
    Dim 棋盘左 As Single, 棋盘顶 As Single
    Dim 棋盘宽 As Single
    Dim 选择左 As Single, 选择顶 As Single, 选择宽 As Single, 选择高 As Single
    Dim x1 As Single, x2 As Single, x3 As Single
    Dim y1 As Single, y2 As Single, y3 As Single
    Dim tx As Single, ty As Single
    
    ' 绘制选择区背景与棋盘线条
    棋盘左 = 块(0).Left - 格内边距
    棋盘顶 = 块(0).Top - 格内边距
    棋盘宽 = 4 * 格大小 + 2 * 细间隙 + 粗间隙

    选择左 = (块(16).Left - 格内边距) - 200
    选择顶 = (块(16).Top - 格内边距) - 200
    选择宽 = (块(19).Left + 块(19).Width + 格内边距) - (块(16).Left - 格内边距) + 400
    选择高 = 格大小 + 400
    Me.FillStyle = vbSolid
    Me.FillColor = 选择区颜色
    Me.Line (选择左, 选择顶)-(选择左 + 选择宽, 选择顶 + 选择高), 选择区颜色, BF

    Me.FillColor = Me.BackColor
    Me.Line (棋盘左, 棋盘顶)-(棋盘左 + 棋盘宽, 棋盘顶 + 棋盘宽), Me.BackColor, BF

    Me.DrawWidth = 线粗
    Me.Line (棋盘左 - 3, 棋盘顶 - 3)-(棋盘左 + 棋盘宽 + 3, 棋盘顶 + 棋盘宽 + 3), vbBlack, B
    
    x1 = 棋盘左 + 格大小 + 细间隙 / 2
    x2 = 棋盘左 + 2 * 格大小 + 细间隙 + 粗间隙 / 2
    x3 = 棋盘左 + 3 * 格大小 + 2 * 细间隙 + 粗间隙 + 细间隙 / 2
    y1 = 棋盘顶 + 格大小 + 细间隙 / 2
    y2 = 棋盘顶 + 2 * 格大小 + 细间隙 + 粗间隙 / 2
    y3 = 棋盘顶 + 3 * 格大小 + 2 * 细间隙 + 粗间隙 + 细间隙 / 2

    Me.DrawWidth = 2
    Me.Line (x1, 棋盘顶)-(x1, 棋盘顶 + 棋盘宽), vbBlack
    Me.Line (x3, 棋盘顶)-(x3, 棋盘顶 + 棋盘宽), vbBlack
    Me.Line (棋盘左, y1)-(棋盘左 + 棋盘宽, y1), vbBlack
    Me.Line (棋盘左, y3)-(棋盘左 + 棋盘宽, y3), vbBlack

    Me.FillColor = vbBlack
    tx = 线粗 * Screen.TwipsPerPixelX
    ty = 线粗 * Screen.TwipsPerPixelY
    Me.Line (x2 - tx / 2, 棋盘顶)-(x2 + tx / 2, 棋盘顶 + 棋盘宽), vbBlack, BF
    Me.Line (棋盘左, y2 - ty / 2)-(棋盘左 + 棋盘宽, y2 + ty / 2), vbBlack, BF
End Sub

Private Sub 清空选择()
    ' 清除当前选择状态
    sf = ""
    选中来源 = 0
    选中索引 = -1
End Sub

Private Sub 重绘()
    ' 刷新并重画网格与背景
    Me.Refresh
    Form_Paint
End Sub

Private Sub 清空块(ByVal idx As Integer)
    ' 恢复为空白格外观
    With 块(idx)
        .Caption = ""
        .BackColor = Me.BackColor
        .BorderStyle = 0
        .Tag = 0
    End With
End Sub

Private Sub 设置块(ByVal idx As Integer, ByVal val As Integer, ByVal tagVal As Integer, ByVal bgColor As Long)
    ' 设置格子内容与外观
    With 块(idx)
        .Caption = CStr(val)
        .BackColor = bgColor
        .BorderStyle = 1
        .Tag = tagVal
    End With
End Sub

Private Sub 重置状态()
    ' 重置棋盘状态与计数
    Dim i As Integer, j As Integer

    sf = ""
    选中来源 = 0
    选中索引 = -1
    Label1.Caption = ""

    For i = 1 To 4
        For j = 1 To 4
            棋位(i, j) = 空棋位
        Next j
        题目数(i) = 0
        已放数(i) = 0
    Next i
End Sub

Private Function 已填满() As Boolean
    ' 判断棋盘是否已全部填满
    Dim i As Integer, j As Integer

    For i = 1 To 4
        For j = 1 To 4
            If 棋位(i, j) = 空棋位 Then Exit Function
        Next j
    Next i
    已填满 = True
End Function

Private Sub 更新选择器()
    ' 根据题目数与已放数刷新选择区
    Dim i As Integer
    Dim digit As Integer
    Dim 可用数 As Integer

    For i = 16 To 19
        digit = i - 15
        可用数 = 4 - 题目数(digit) - 已放数(digit)
        If 可用数 <= 0 Then
            块(i).Caption = ""
            块(i).Enabled = False
            块(i).BorderStyle = 0
            块(i).BackColor = 选择区颜色
        Else
            块(i).Caption = CStr(digit)
            块(i).Enabled = True
            块(i).BorderStyle = 1
            块(i).BackColor = vbWhite
        End If
    Next i
End Sub

Private Function 检查冲突提示(r As Integer, c As Integer, val As Integer) As String
    ' 检测行/列/宫冲突并返回提示
    Dim i As Integer, j As Integer
    Dim r0 As Integer, c0 As Integer

    For i = 1 To 4
        If i <> c Then
            If 棋位(r, i) = val Then
                检查冲突提示 = "请检查本行落子。"
                Exit Function
            End If
        End If
    Next i

    For i = 1 To 4
        If i <> r Then
            If 棋位(i, c) = val Then
                检查冲突提示 = "请检查本列落子。"
                Exit Function
            End If
        End If
    Next i

    r0 = ((r - 1) \ 2) * 2 + 1
    c0 = ((c - 1) \ 2) * 2 + 1
    For i = r0 To r0 + 1
        For j = c0 To c0 + 1
            If Not (i = r And j = c) Then
                If 棋位(i, j) = val Then
                    检查冲突提示 = "请检查本宫落子。"
                    Exit Function
                End If
            End If
        Next j
    Next i

    检查冲突提示 = ""
End Function
