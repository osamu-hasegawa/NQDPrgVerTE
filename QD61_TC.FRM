VERSION 5.00
Begin VB.Form QD61_TC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "成形オンリー画面"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command2 
      Caption         =   "V エディタ画面"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   58
      Top             =   600
      Width           =   1668
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3000
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "グラフ再描画"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2040
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5500
      Left            =   1800
      ScaleHeight     =   5475
      ScaleWidth      =   8370
      TabIndex        =   10
      Top             =   1860
      Width           =   8400
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   7
         X1              =   6696
         X2              =   6696
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   6
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   5
         X1              =   3348
         X2              =   3348
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   4
         X1              =   1656
         X2              =   1656
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   3
         X1              =   0
         X2              =   8352
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   2
         X1              =   0
         X2              =   8352
         Y1              =   2196
         Y2              =   2196
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   1
         X1              =   0
         X2              =   8352
         Y1              =   3312
         Y2              =   3312
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   0
         X1              =   0
         X2              =   8352
         Y1              =   4392
         Y2              =   4392
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "キャンセル"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   10200
      TabIndex        =   59
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コマンド："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   57
      Top             =   8160
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   1425
      TabIndex        =   56
      Top             =   8160
      Width           =   6000
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   1440
      TabIndex        =   55
      Top             =   8520
      Width           =   5940
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   6360
      TabIndex        =   54
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ショット数："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   9
      Left            =   7920
      TabIndex        =   53
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "サイクルタイム："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   8
      Left            =   7440
      TabIndex        =   52
      Top             =   390
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   9420
      TabIndex        =   51
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   9420
      TabIndex        =   50
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1425
      TabIndex        =   49
      Top             =   7800
      Width           =   8760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コマンド："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   48
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   9480
      TabIndex        =   47
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   7920
      TabIndex        =   46
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3225
      TabIndex        =   45
      Top             =   1320
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "成形状態："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   1920
      TabIndex        =   44
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(分)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   31
      Left            =   9360
      TabIndex        =   42
      Top             =   7560
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "経過時間"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   30
      Left            =   7275
      TabIndex        =   41
      Top             =   7560
      Width           =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   27
      X1              =   10200
      X2              =   10200
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   26
      X1              =   8520
      X2              =   8520
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   25
      X1              =   6840
      X2              =   6840
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   24
      X1              =   5160
      X2              =   5160
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   23
      X1              =   3480
      X2              =   3480
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   29
      Left            =   9930
      TabIndex        =   40
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   28
      Left            =   8355
      TabIndex        =   39
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   27
      Left            =   6660
      TabIndex        =   38
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   26
      Left            =   4965
      TabIndex        =   37
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   3270
      TabIndex        =   36
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   1650
      TabIndex        =   35
      Top             =   7485
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   22
      X1              =   1800
      X2              =   1800
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   21
      X1              =   10200
      X2              =   1800
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "型温度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   23
      Left            =   1230
      TabIndex        =   34
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(℃)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   22
      Left            =   1290
      TabIndex        =   33
      Top             =   1515
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   20
      X1              =   1620
      X2              =   1764
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   19
      X1              =   1620
      X2              =   1764
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   18
      X1              =   1620
      X2              =   1764
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   17
      X1              =   1620
      X2              =   1764
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   16
      X1              =   1620
      X2              =   1764
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   15
      X1              =   1620
      X2              =   1764
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   14
      X1              =   1770
      X2              =   1770
      Y1              =   1875
      Y2              =   7375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   21
      Left            =   1170
      TabIndex        =   32
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   20
      Left            =   1290
      TabIndex        =   31
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   19
      Left            =   1290
      TabIndex        =   30
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   18
      Left            =   1290
      TabIndex        =   29
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   17
      Left            =   1290
      TabIndex        =   28
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   16
      Left            =   1290
      TabIndex        =   27
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "型締圧"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   15
      Left            =   540
      TabIndex        =   26
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(t)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   14
      Left            =   720
      TabIndex        =   25
      Top             =   1515
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   13
      X1              =   1005
      X2              =   1149
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   12
      X1              =   1005
      X2              =   1149
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   11
      X1              =   1005
      X2              =   1149
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   10
      X1              =   1005
      X2              =   1149
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   9
      X1              =   1005
      X2              =   1149
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   8
      X1              =   1005
      X2              =   1149
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   7
      X1              =   1155
      X2              =   1155
      Y1              =   1860
      Y2              =   7360
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   13
      Left            =   645
      TabIndex        =   24
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   12
      Left            =   645
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   11
      Left            =   645
      TabIndex        =   22
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   10
      Left            =   645
      TabIndex        =   21
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   9
      Left            =   645
      TabIndex        =   20
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   8
      Left            =   645
      TabIndex        =   19
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "座標"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   7
      Left            =   30
      TabIndex        =   18
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(mm)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   6
      Left            =   30
      TabIndex        =   17
      Top             =   1515
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   6
      X1              =   390
      X2              =   534
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   5
      X1              =   390
      X2              =   534
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   4
      X1              =   390
      X2              =   534
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   3
      X1              =   390
      X2              =   534
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   2
      X1              =   390
      X2              =   534
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   1
      X1              =   390
      X2              =   534
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   0
      X1              =   540
      X2              =   540
      Y1              =   1860
      Y2              =   7360
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   5
      Left            =   30
      TabIndex        =   16
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   4
      Left            =   30
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   3
      Left            =   30
      TabIndex        =   14
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   11
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コメント："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   1905
      TabIndex        =   9
      Top             =   1005
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3210
      TabIndex        =   8
      Top             =   1005
      Width           =   6945
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   690
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "制御ファイル名："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   1944
      TabIndex        =   4
      Top             =   684
      Width           =   2028
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   7005
      TabIndex        =   3
      Top             =   75
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5385
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5385
      TabIndex        =   1
      Top             =   75
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "測定時間："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   4125
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "QD61_TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lGphNo%
Dim lGphNo0%
Dim EditFlg%
Dim lViewFlg      '前の画面番号
Dim NextView%
Dim lEmgFlg%      '非常停止


Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0  'キャンセル
  lGphNo = 0
  MoniGraph Me.Picture1, 0, lGphNo
Case 1  '終了
  FrmMenuFlg = False
  NextView = 1
  'Unload Me
  'PGM_Menu.Show
Case 2  'グラフ再描画
  lGphNo = lGphNo + 100
  MoniGraph Me.Picture1, 0, lGphNo
Case 3  'edit
  'FrmMenuFlg = False                    'メニューから抜けるときfalse
  'NextView = 8
  EditFlg% = True 'エディタ起動
End Select
End Sub

Private Sub SetData()
  Label2(0) = Format(ptime, "###0")  '測定時間
  Label2(1) = Format(ytemp, "###0")  '予備加熱温度
  Label2(2) = gcoxFlName             '制御ファイル名
  Label2(3) = hcomm(2)               'コメント
' -----------------------------------
  DispGphScale
End Sub

Private Sub GetData()

End Sub

Private Sub Form_Load()
  lViewFlg = ViewFlg      '前の画面番号
  ViewFlg = 2             '画面番号
  FrmMenuFlg = True       'メニューから抜けるときfalse
  EditFlg% = False        'エディタ起動解除
  lEmgFlg = False         '非常停止
  SetData
  Timer2.Enabled = False
  Timer1.Enabled = True
End Sub


Private Sub DispGphScale()
Dim i%, p%, max!, min!, def!, dev%
  '
  GphXSet           '時間軸の時間をセット
  '
  dev = 5
  '
  min = InitDat(1)  'グラフスケール座標 (Min)
  max = InitDat(2)  'グラフスケール座標 (Max)
  def = (max - min) / dev
  p = 0
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(3)  'グラフスケール型締圧 (Min)
  max = InitDat(4)  'グラフスケール型締圧 (Max)
  def = (max - min) / dev
  p = 8
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(5)  'グラフスケール型温度 (Min)
  max = InitDat(6)  'グラフスケール型温度 (Max)
  def = (max - min) / dev
  p = 16
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(7)  'グラフスケール経過時間 (Min)
  max = InitDat(8)  'グラフスケール経過時間 (Max)
  def = (max - min) / dev
  p = 24
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
'
End Sub
Private Sub Timer1_Timer()
  Label2(4).Caption = ""
  Timer1.Enabled = False
  Qd61S_MAIN
End Sub
Private Sub Timer2_Timer()
  'Qd61S_Monitor
End Sub

Public Sub Qd61S_MAIN()
Dim i%, j%, js%, l%, ist0%, ist1%, ndata!, mdata!, ntemp!, mtemp!, iflg%
Dim ied%, ips%, i_s%, irei%, r_ch%, ix%, ix0%, iy%, isp%, stime%
Dim ie02%, ie03%, ie04%
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%
Dim isub As Long, jsub As Long, ksub As Long
Dim m_l%, sdata%, sv%, zch%
Dim c$, com$, tdate$, ttime$
Dim st!, ev!, sev!, fin!          '/* 時間用データ */
Dim btemp!(0 To 4), bposi!, bpre! '/* 温度　位置　圧力 の前データ */
Dim ivd%, id_0%, id_1%, id_2%
'
Dim sdt$, ch%, hdt%, flindex%, sts%, imax%, sts1%, sts2%, ch1%, ch2%
Dim wTm0!, wTm1!, S_StartTime!
Dim stTime!, evTime!, finTime!, sevTime!, mTime!
Dim dt!(0 To 4)
Dim flg As Long, cnt As Long
Dim diTime!, diTime1!, diTime2!, pdt!, pml!, x1dt!, x2dt!, pos!
'
  ips = 1: i_s = -1
  c = "0"
  ivd = 0:   id_0 = 0: id_2 = &H8
  Timer2.Enabled = True       'Debug
'
'----------------------- 成形シングルメインプログラム
Timer2.Enabled = False
  C870Stop
  ServoON       '/* サーボｏｎ */
  CtlDisp       '位置制御
  TrnsReqOFF    '搬送依頼信号OFF
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　　　*************/
'/* SPEC INITIALIZE CMD OUT */
'/* カウンタボードの初期設定 */
  InitDat(10) = 0
'/* 加減速ﾚｰﾄｾｯﾄｺﾏﾝﾄﾞ */
  C870AccRate
'/* 速度設定 */
  C870LSPDSet 300    '/* 300 pps 0.066mm/sec */
'/* ディレータイム設定 */
  C870DelayTime
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　終了  *************/

'/* ＡＴＣ温度リセット */
'/* ロボットデータのフロッピーからの読みとり */
  rozFileLoad
st:
  'clstxt();
  If ied = 2 Then GoTo st2:
'/*  制御ファイルのオープン */
  coxDtRead gcoxFldir & gcoxFlName
  Label2(0).Caption = Format(ptime, "0")
  '/* グラフィック画面の初期化 */
  InitDat(8) = ptime  'グラフスケール経過時間 (Max)
  SetData
  lGphNo0 = 0: lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* 予備加熱温度設定 */
'/* 軸駆動制御コマンドのファイルからの読み取り */
  i = 0
  Do
    sdt = Right("     " & Format(i, "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(seg_num(i), "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(ic(i), "0"), 4)
    sdt = sdt & "  " & Right("         " & Format(z(i), "0.000"), 7)
    sdt = sdt & "  " & Right("         " & Format(vel(i), "0.0"), 7)
    sdt = sdt & "  " & Right("       " & Format(pres(i), "0"), 6)
    sdt = sdt & "  " & Right("     " & Format(t0(i), "0.0"), 4)
    sdt = sdt & "  " & Right("     " & Format(p(i), "0.0"), 4)
    Label2(12).Caption = sdt
    If pres(i) >= 1000 Then ips = 2    '/* ﾌﾟﾚｽ圧が1ton以上で軸変更 */
    i = i + 1
    If ic(i - 1) = 3 Then Exit Do
  Loop

  ic(i) = 4
'/* 表題の表示 */
  Label2(2).Caption = gcoxFlName
'/* 原点出し */
  Label2(4).Caption = "原点出し実行"
  genten
  Ready_Wait
  Counter0
  Label2(4).Caption = "原点出し完了"
'/* カウンタにゼロを書き込む */
  'C870AdrInit       'ＡＤＤＲＥＳＳ ＩＮＩＴＡＬＩＺＥ ＣＯＭＭＡＮＤ
  C870CntPreSet 0   'ＣＯＵＮＴＥＲ ＰＲＥＳＥＴ ＣＯＭＭＡＮＤ
  'InitDat(10) = 0
  pos = r_z()
  GCnt0 = 0: GCnt1 = 0
'/* 自動運転認識 */
  Label2(4).Caption = "自動運転認識中"
  ch1 = 1            'システムレディーと自動
  ch2 = 3            '自動・手動
  Do
    DoEvents
    If FrmMenuFlg = False Then GoTo eend:            'メニューから抜けるときfalse
    '-------------- ピラニ計読み
    Qd61S_Monitor
    '
    DioInput ch1, sts1
    DioInput ch2, sts2
    If sts1 = 1 And sts2 = 1 Then Exit Do
  Loop
  Label2(4).Caption = ""
'/* 成形プロセス開始　連続前コマンド */

  flindex = 0      '制御コマンドファイルの位置
  Do
    DoEvents
    '-------------- ピラニ計読み
    Qd61S_Monitor
    'flindex = flindex + 1
    com = Left(scom(flindex), 1)
    isub = sisub(flindex)
    sdt = Right("    " & scom(flindex), 2)
    sdt = sdt & Right(Space(15) & Format(isub, "0"), 15)
    If com = "S" Then
      jsub = sjsub(flindex)
      ksub = sksub(flindex)
      sdt = sdt & Right(Space(15) & Format(jsub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(ksub, "0"), 15)
    Else
      
    End If
    Label2(7).Caption = sdt
    flindex = flindex + 1
    i = 10
    '
    If ied <> 0 Then GoTo jp0:
    '
    Select Case com
      Case "B"
        'Exit Do
      Case "N"    '/* 窒素ガスの注入 */
        If Mid(scom(flindex), 2, 1) = "S" Then
          If isub = 1 Then
            Label2(4).Caption = "窒素ガス注入"
            N2Open
          End If
          If isub = 0 Then
            Label2(4).Caption = "窒素ガス停止"
            N2Close
          End If
        End If
      Case "J"    '/* 時間待ち */
        Label2(4).Caption = "時間待ち"
        finTime = Timer
        Do
          evTime = Timer
          DoEvents
          'Label2(10).Caption = Format(wTm1 - Timer, "0")
          Label2(10).Caption = Format(diffTime(finTime, evTime), "0")
          If diffTime(finTime, evTime) >= isub Then Exit Do
        Loop
        Label2(10).Caption = ""
      Case "K"    '/* 加熱 */
        Select Case Int(isub)
        Case 1
          Label2(4).Caption = "加熱　ＯＮ"
          HeatON
        Case 0
          Label2(4).Caption = "加熱　ＯＦＦ"
          HeatOFF
        End Select
      Case "S"    '/* ＡＴＣ温度設定 */
        Label2(4).Caption = "ＡＴＣ温度設定"
        finTime = Timer              '待ち初めの時間
        Do
          evTime = Timer
          DoEvents
          diTime = diffTime(finTime, evTime)
          x1dt = diTime / ksub
          ndata = (isub - ntemp) * x1dt + ntemp
          mdata = (jsub - mtemp) * x1dt + mtemp
          TempSet 2, ndata
          TempSet 3, mdata
          If diTime >= ksub Then Exit Do
        Loop
        ntemp = isub
        mtemp = jsub
        TempSet 2, ntemp
        TempSet 3, mtemp
      Case "R"    '/* 冷却 */
        Select Case Int(isub)
        Case 0    '冷却大　ＯＦＦ
          Label2(4).Caption = "冷却大　ＯＦＦ"
          CoolOFF
        Case 1    '冷却大　ＯＮ
          Label2(4).Caption = "冷却大　ＯＮ"
          CoolON
        Case 2    '冷却小　ＯＮ
          Label2(4).Caption = "冷却小　ＯＮ"
          CoolON
        End Select
    End Select
jp0:
    If i < 24 Then
      i = i + 1
    Else
      Label2(4).Caption = ""
    End If
    If com = "B" Then Exit Do
  Loop
'/* 成形プロセス連続運転開始 */
'/* データを読み取る */

'/* ブザーを鳴らす */
  'Label2(4).Caption = ""
  'Label2(10).Caption = ""
st2:
'/* タイトルの表示 */
'/* 型締圧軸の表示 */
'/* 座標値軸の表示 */
'/* 搬送用Ｚ軸位置変更枠表示 */
  Label2(5).Caption = Format(roz(0), "0.000")     '/* ロボット位置表示 */
  Label2(6).Caption = Format(roz(1), "0.000")     '/* ロボット位置表示 */
'/* 成形開始 */
  Do        '----------------- DO LOOP
    DoEvents
    i_s = i_s + 1
    js = 0
    ist0 = -1: ist1 = -1
    ie0 = 0: ie1 = 0: ie2 = 0: ie3 = 0
    S_StartTime = Timer
    sevTime = Timer
'/* カウンタへの出力ｕｐ */

    If i_s <> 0 Then
      InitDat(11) = InitDat(11) + 1  '成形カウンタトウタル
      InitDtSave
      Label2(13).Caption = Str(InitDat(11))
    End If
'/* 成形枠の表示 */
ejs1:
  lGphNo0 = 0: lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* Ｘ軸の表示 */
'/* Ｙ軸の表示 */
'/* ｼｮｯﾄ数ｻｲｸﾙﾀｲﾑ枠表示 */
    sdt = Format(stime / 60, "0") & "分" & Format(stime Mod 60, "0") & "秒"
    Label2(9).Caption = Format(i_s, "0")
    Label2(8).Caption = sdt
    InitDat(10) = i_s               '成形カウンタ
'/* カウンタへの出力ダウン */
    'InitDat(11) = InitDat(11) - 1   '成形カウンタトウタル
    'InitDtSave
    'Label2(13).Caption = Str(InitDat(11))
'/* データの取り込み */
    stTime = Timer
    evTime = Timer
    iflg = 1: ied = 0
    ttime = Time
    mTime = Timer
'----------------------------- For Loop i
    imax = ptime * 60
    For i = 1 To imax
start:
    finTime = Timer
    DoEvents
'/* 成形軸のドライブ*/
      If ist0 > 0 Then
        If ic(ist0 - 1) = 4 Then
          ist0 = ist0 - 1
        End If
      End If
        DispSegm ist0
      If ist0 <> ist1 Then
        gOrgFlg = False                '原点復帰完了=TRUE
        ist1 = ist0
        sevTime = Timer
        Select Case ic(ist0)
        Case 0
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9); サーボON & 速度上限S12
          s_drive z(ist0), vel(ist0)
        Case 1
          m_l = vel(ist0) * 2#
          If m_l > 50 Then m_l = 50
          setcm1 z(ist0)
          Ready_Wait    '
          CtlVelo       'outp(DIO_P+3,5);
          'while((inp(XCN_DT1)&0x01)!=0);
          Do
            DoEvents
            sts = C870Sts(3)
            If (sts And &H1) = 0 Then Exit Do   'PULSE と COMPARE が一致
          Loop
          'Label2(11).Caption = Format(m_l, "0.000") 'printf("%7.3f",m_l);
        Case 2
          Ready_Wait    '
          ServoON     'outp(DIO_P+3,1);
        Case 3
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9);
          genten
          'Ready_Wait
        End Select
      End If
'/* タイムアップ処理 */
      If ist0 < 0 Then GoTo sj1:
      If (ic(ist0) < 3) And (diffTime(Int(sevTime), Int(finTime)) > t0(ist0)) Then
        sdt = "タイムアップ" & Right(Space(11) & Format(diffTime(sevTime, finTime), "0.0"), 11)
        sdt = sdt & Right(Space(11) & Format(t0(ist0), "0.0"), 11)
        Label2(4).Caption = sdt
        hijyou        '非常停止処理
        'getch
        GoTo eend:
      End If
'/* 終了信号の処理 */
      Select Case ic(ist0)
      Case 0    '/* 位置制御の場合 */
        If (C870Sts(1) And 1) = 0 Then ist0 = ist0 + 1
      Case 1    '/* 速度制御の場合 */
        pdt = pres(ist0): pml = m_l
        cal_pid pdt, p(ist0), pml
        sts = C870Sts(3)  'status3 を読む
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* 位置達成で終了 */
          Do
            DoEvents
            sts = C870Sts(3)          'status3 を読む
            If (sts And &H1) <> 0 Then Exit Do
          Loop
        Else
          If r_z() > z(ist0) Then
            ist0 = ist0 + 1             '
            Label2(4).Caption = "位置制御 error" & Str(ist0)
          End If
        End If
      Case 3
        sts = C870Sts(1)
        If (sts And 1) = 0 Then
          ist0 = ist0 + 1     '/* 完了 */
          If Abs(r_z()) > 0.1 Then
            Label2(4).Caption = "原点不良"
            ist0 = ist0 - 1
            genten              '原点出し
          End If
        Else
          '/* カウンタにゼロを書き込む */
          Ready_Wait
          Counter0
        End If
      End Select
'/* エラー表示 */
      '------------------ BITS を読む
      ie0 = 0     '
      ie02 = 0    '
      ie03 = 0
      ErrBitRd ie0, ie02, ie03
  
      ie = (ie02 And &H8) * &H2
      ie04 = (ie02 And &HF0) * &H10 + ie03
  
      If ie <> &H10 Then
        ie2 = ie04
  
          If (ie0 <> ie1) Or (ie2 <> ie3) Then
            ie1 = ie0
            ie3 = ie2
            err_sign ie0 + ie, ie2
          End If
      Else
        ie4 = ie04
        If (ie0 <> ie1) Or (ie4 <> ie5) Then
          ie1 = ie0
          ie5 = ie4
          err_sign ie0 + ie, ie4
        End If
      End If

'/* プロセス実行 */
sj1:
    If iflg = 1 Then
      com = scom(js + flindex)
      isub = sisub(js + flindex)
      jsub = sjsub(js + flindex)
      ksub = sksub(js + flindex)
      js = js + 1
      sdt = com & Right(Space(7) & Format(isub, "0"), 7)
      'locate(60,1);printf("%2s %5ld ",com,isub);
      If Left(com, 1) = "S" Then
        sdt = sdt & Right(Space(7) & Format(jsub, "0"), 7)
        sdt = sdt & Right(Space(7) & Format(ksub, "0"), 7)
        'printf("%5ld %5ld",jsub,ksub);
        'Label2(7).Caption = sdt
      Else
        sdt = sdt
      End If
      Label2(7).Caption = sdt
    End If
        '
        ch1 = 1            'システムレディーと自動
        ch2 = 3
        DioInput ch1, sts1
        DioInput ch2, sts2
        'DioInput 1, sts   'システムレディ/* 非常停止の場合は成形中止 */
        If sts1 = 0 Or sts2 = 0 Then
          FrmMenuFlg = False                    'メニューから抜けるときfalse
          GoTo eend:
        End If
        '
        Select Case Left(com, 1)
        Case "S"    '/* ＡＴＣ温度設定 */
          Label2(4).Caption = "ＡＴＣ温度設定"
          DoEvents
          diTime = diffTime(evTime, finTime)
          x1dt = diTime / ksub
          ndata = (isub - ntemp) * x1dt + ntemp
          mdata = (jsub - mtemp) * x1dt + mtemp
          TempSet 2, ndata
          TempSet 3, mdata
          If diTime >= ksub Then
            iflg = 1
            ntemp = isub    'ndata
            mtemp = jsub    'mdata
            TempSet 2, ntemp
            TempSet 3, mtemp
            evTime = Timer
          Else
            iflg = 0
          End If
        Case "P"    '/* 移動軸制御の駆動 */
          If Mid(com, 2, 1) = "W" Then
            Beep
            ist0 = ist0 + 1
          End If
          If Mid(com, 2, 1) = "R" Then
            iflg = 0
            If ist0 <> ist1 Then iflg = 1
            If isub = 4 And ist0 = 0 Then iflg = 1
          End If
          evTime = Timer
        Case "K"    '/* 加熱 */
          Select Case isub
          Case 1
            HeatON
          Case 0
            HeatOFF
          End Select
        Case "N"
          If Mid(com, 2, 1) = "S" Then
            If isub = 1 Then hdt = hdt
            If isub = 0 Then hdt = hdt
          End If
        Case "R"    '/* 冷却 */
          Select Case isub
          Case 2
            CoolON
          Case 1
            CoolON
          Case 0
            CoolOFF
          End Select
        Case "T"    '/* ＡＴＣ１の温度の読み取り */
          AdRead dt(), flg
          sdata = dt(0) / 10 * 2000   '* 204.8    '上モールド温度
          If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Then
            iflg = 0
          Else
            If iflg = 2 Then iflg = 1 Else iflg = 2
            evTime = Timer
          End If
        Case "J"    '/* 時間待ち */
          diTime1 = diffTime(stTime, finTime)  'diffTime(finTime, stTime)
          diTime2 = diffTime(evTime, finTime)   'diffTime(finTime, evTime)
          If (Mid(com, 2, 1) = "S" And diTime1 >= isub) Or (Mid(com, 2, 1) = "C" And diTime2 >= isub) Then
            iflg = 1
            evTime = Timer
          Else
            iflg = 0
          End If
        Case "C"
          'Label2(11).Caption = Format(aposi(i - 1), "0.000")
        Case "X"    '搬送終了信号（成形開始）
          Select Case Mid(com, 2, 1)
          Case "R"    '成形開始 [搬送終了まで待つ]
            Label2(4).Caption = "成形開始 [搬送終了まで待つ]"
            '
            TrnsReqON  '搬送依頼信号Ch21出力
            '
            Do
              '-------------- ピラニ計読み
              Qd61S_Monitor
              DioInput 13, sts
              If sts = 1 Then
                TrnsReqOFF    '搬送依頼信号ＯＦＦ
                Label2(4).Caption = "搬送依頼信号ＯＦＦ"
                Exit Do
              End If
              DoEvents
            Loop
          Case "W"    '成形終了
          End Select
        Case "E"    '/* 終了　ロボット搬送 */
          If iflg <> 99 Then
            iflg = 0
            If r_z() > 2 Then
              genten
              'Ready_Wait    'while((inp(AX_STS)&1)!=0);
            End If
            TrnsReqON  '搬送依頼信号Ch21出力
            Label2(4).Caption = "搬送依頼信号Ch21出力"
            WaitSec 1     '
            '搬送表示信号Ch15を待つ
            'DioInput 15, sts
            'If sts = 1 Then
              iflg = 99
            'End If
            isp = 0
          Else
            '--------------------- 成形ONLY 終了が押されるまで待つ
            If FrmMenuFlg = False Then  'メニューから抜けるときfalse
              GoTo send:
            End If
            'DioInput 13, sts    '搬送終了信号Ch13を待つ
            'If sts = 1 Then
            '  TrnsReqOFF  '搬送依頼信号OFF
            '  Label2(4).Caption = ""
            '  GoTo send:
            'Else
            '  Label2(4).Caption = "搬送終了信号Ch13を待つ"
            'End If
          End If
        End Select
cjump:
  '-------------- ピラニ計読み
    Qd61S_Monitor
    lEmgFlg = EmgChk()         '非常停止
  If Int(mTime) = Int(Timer) And lEmgFlg = False Then GoTo start:
    mTime = Timer
'
    If FrmMenuFlg = False Then             'メニューから抜けるときfalse
      Select Case NextView
      Case 1
        sdt = "終了を受付"
      Case 8  'edit
        sdt = "editを受付"
      Case Else
      
      End Select
      Label2(10).Caption = sdt
    Else
      Label2(10).Caption = ""
    End If
'                 /* 1秒に1回下に抜ける */
'           /* 圧力　ＰＩＤ制御　Ｐ＞１５　なら速度　ゼロ */
  If ist0 >= 0 Then
    If p(ist0) > 15 Then
      DaVoltOut 1, 0        ' 0V D/A ch=1
    End If
  End If
'/*　温度取り込み */
    DoEvents
    AdRead dt(), flg
    atemp(i, 0) = dt(0) / 10 * 2000   '* 204.8    '上モールド温度
    atemp(i, 1) = dt(1) / 10 * 2000   '* 204.8    '下モールド温度
  
'* 上型位置の取り込み */

    aposi(i) = r_z()
      'Label2(11).Caption = Format(aposi(i), "0.000")
'/* 型圧力の取り込み */

      apre(i) = r_pres()

'      If i = 1 Then GoTo jo0:
'      ix0 = Int(8.3333 / ptime * (i - 1)) + 120
'      ix = Int(8.3333 / ptime * (i)) + 120
'-------------- ピラニ計読み
'      Qd61S_Monitor
'/* 温度分布の表示 */
'/* 型締圧のプロット */
'/* 座標値のプロット */
    lGphNo = i
    GphDataSet lGphNo0, lGphNo
    MoniGraph Me.Picture1, lGphNo0, lGphNo
    lGphNo0 = lGphNo
jo0:
'/* 各種データの画面下表示 */
    sdt = Right(Space(2) & Format(Int(i / 60), "0"), 2) & "分"
    sdt = sdt + Right(Space(3) & Format(i Mod 60, "0"), 3)
    sdt = sdt & "秒  位置" & Right(Space(10) & Format(aposi(i), "0.000"), 10) & "mm"
    sdt = sdt & Right(Space(10) & Format(diffTime(sevTime, finTime), "0.0"), 10)
    Label2(11).Caption = sdt
'/* ﾛﾎﾞｯﾄ位置変更　*/
    'If FrmMenuFlg = False Then GoTo eend:
  Next i   '--------------------------------- For Loop
  js = js - 1
  GoTo ejs1:      '/* 表示終了で元画面へ */
'/* タクトタイムの算出　*/

send:
    stime = i
'/* データの保存　*/
    'ResDtSave i_s, stime
'/* Ｖ入力がされていたら　エディット */
    If FrmMenuFlg = False Then Exit Do            'メニューから抜けるときfalse
    If EditFlg% = True Then 'エディタ起動
      EditFlg% = False      'エディタ起動解除
      MYEdit.Show 1
      ied = 1
      c = 0
    End If
'/* エディットモードであれば　ｓｔにジャンプ */
    If ied <> 0 Then GoTo st:
'/* 自動停止状態であれば停止 */
    ch1 = 1            'システムレディーと自動
    ch2 = 3
    DioInput ch1, sts1
    DioInput ch2, sts2
    If sts1 = 0 Or sts2 = 0 Then
      Label2(4).Caption = "自動運転終了"
      FrmMenuFlg = False              'メニューから抜けるときfalse
      NextView = 1
      Exit Do
    End If
  Loop    '------------------------------------ DO LOOP
'/* 予備加熱をゼロにし、ＯＦＦする */
eend:
  HeatOFF
  CoolOFF
  ServoOFF
'/* ロボットデータのフロッピーへの書き出し */
'/* グラフィック画面の消去 */

'/* Ｖ入力がされていたら　エディット */
  '
  If FrmMenuFlg = False Then             'メニューから抜けるときfalse
    FrmMenuFlg = True                    'メニューから抜けるときfalse
    Select Case NextView
    Case 1
      Unload Me
      PGM_Menu.Show
    Case 2 '成形（シングル）
      QD61_SC.Show
    Case 3  '成形（ダブル）
    Case 4  'I O チェック
      IOChk.Show
    Case 5  'スケール変更
      QD61_GphScale.Show
    Case 6  '読み出し
    Case 7  'メモ帳
    Case 8  'edit
      Unload Me
      MYEdit.Show
    Case Else
      Unload Me
      PGM_Menu.Show
    End Select
  End If
  
End Sub
Private Sub genten()
'--------------
  C870Genten
  gOrgFlg = True                       '原点復帰完了=TRUE
End Sub
Private Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* 設定圧力 */
'         m_p,      /* 設定Ｐ値 */
'         m_lim;    /* 設定リミット値 */
Dim i%, nout%, ch%, v!
Dim pa!, per!
  pa = r_pres()     '/* 圧力 */

  If pa > 1000 Then '/* 指定圧力＋２００Ｋｇ以上で非常停止 */
    hijyou
    Exit Sub
  End If

'/* ＰＩＤ演算 */

  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
  If per > m_lim Then per = m_lim
  If per < (-1 * m_lim) Then per = -1 * m_lim
  nout = Int(40.95 * per) + &H800
  ch = 1
  v = nout
  DaOut ch, Hex(nout)
  'DaVoltOut ch, V
  'outp(ADPORT,(nout%256));
  'outp(ADPORT+1,0x20|(nout/256));
  
End Sub

Private Sub GphXSet()
Dim i%
  For i = 0 To ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub

Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)
    Templd(i) = atemp(i, 1)   '下型温度
    Press(i) = apre(i)
    ZAxis(i) = aposi(i)
  Next i
End Sub

Private Sub DispSegm(ist0%)
Dim sdt$
    If ist0 < 0 Then Exit Sub
    sdt = Right(Space(2) & Format(ist0, "0"), 2)
    sdt = sdt & Right(Space(4) & Format(seg_num(ist0), "0"), 4)
    sdt = sdt & Right(Space(4) & Format(ic(ist0), "0"), 4)
    sdt = sdt & Right(Space(12) & Format(z(ist0), "0.000"), 12)
    sdt = sdt & Right(Space(7) & Format(vel(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(6) & Format(pres(ist0), "0"), 6)
    sdt = sdt & Right(Space(7) & Format(t0(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(7) & Format(p(ist0), "0.0"), 7)
    Label2(12).Caption = sdt
End Sub
Private Function EmgChk%()
Dim ch%, sts%
  ch = 1
  DioInput ch, sts
  If sts = 0 Then
    EmgChk = True
  Else
    EmgChk = False
  End If
End Function

