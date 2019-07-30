VERSION 5.00
Begin VB.Form LS21_TC 
   BackColor       =   &H00C0C0C0&
   Caption         =   "1回成形"
   ClientHeight    =   8532
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   11856
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8532
   ScaleWidth      =   11856
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   720
      Top             =   3240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "強制ソーク"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   5880
      MaskColor       =   &H8000000F&
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   62
      Top             =   655
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.CommandButton Command2 
      Caption         =   "成形開始(指定)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   1920
      TabIndex        =   61
      Top             =   480
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   0
      Left            =   3720
      TabIndex        =   60
      Text            =   "4"
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "成形開始(3回)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1920
      TabIndex        =   59
      Top             =   120
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   5040
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.CommandButton Command2 
      Caption         =   "真空到達"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      Caption         =   "V エディタ画面"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   55
      Top             =   480
      Width           =   1668
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   720
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5500
      Left            =   1800
      ScaleHeight     =   5472
      ScaleWidth      =   8376
      TabIndex        =   9
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "  Z3補正"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   11100
      TabIndex        =   95
      Top             =   4560
      Width           =   612
   End
   Begin VB.Label Label7 
      Caption         =   "  T係数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   10320
      TabIndex        =   94
      Top             =   4560
      Width           =   612
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   10400
      TabIndex        =   83
      Top             =   660
      Width           =   1284
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   9
      Left            =   10320
      TabIndex        =   93
      Top             =   7140
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   8
      Left            =   10320
      TabIndex        =   92
      Top             =   6876
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   7
      Left            =   10320
      TabIndex        =   91
      Top             =   6624
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   6
      Left            =   10320
      TabIndex        =   90
      Top             =   6360
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   5
      Left            =   10320
      TabIndex        =   89
      Top             =   6096
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   4
      Left            =   10320
      TabIndex        =   88
      Top             =   5844
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   3
      Left            =   10320
      TabIndex        =   87
      Top             =   5580
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   2
      Left            =   10320
      TabIndex        =   86
      Top             =   5316
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   10320
      TabIndex        =   85
      Top             =   5064
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   10320
      TabIndex        =   84
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   11160
      TabIndex        =   82
      Top             =   7440
      Width           =   612
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   10320
      TabIndex        =   81
      Top             =   7440
      Width           =   612
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '実線
      Caption         =   "cc3-2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   10320
      TabIndex        =   80
      Top             =   3700
      Width           =   1452
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   9
      Left            =   11100
      TabIndex        =   79
      Top             =   7140
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   8
      Left            =   11100
      TabIndex        =   78
      Top             =   6876
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   7
      Left            =   11100
      TabIndex        =   77
      Top             =   6624
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   6
      Left            =   11100
      TabIndex        =   76
      Top             =   6360
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   5
      Left            =   11100
      TabIndex        =   75
      Top             =   6096
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   4
      Left            =   11100
      TabIndex        =   74
      Top             =   5844
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   3
      Left            =   11100
      TabIndex        =   73
      Top             =   5580
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   2
      Left            =   11100
      TabIndex        =   72
      Top             =   5316
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   11100
      TabIndex        =   71
      Top             =   5064
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   11100
      TabIndex        =   70
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label Label5 
      Caption         =   "cc3"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   10320
      TabIndex        =   69
      Top             =   3320
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   10320
      TabIndex        =   68
      Top             =   2960
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   10320
      TabIndex        =   67
      Top             =   2600
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   10320
      TabIndex        =   66
      Top             =   2200
      Width           =   1452
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '実線
      Caption         =   "ct1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   10320
      TabIndex        =   65
      Top             =   1860
      Width           =   1452
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '実線
      Caption         =   "cp1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   10320
      TabIndex        =   64
      Top             =   4120
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Height          =   252
      Index           =   14
      Left            =   6720
      TabIndex        =   63
      Top             =   7800
      Width           =   4980
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
      Left            =   10750
      TabIndex        =   56
      Top             =   80
      Width           =   950
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
      TabIndex        =   54
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
      Left            =   1428
      TabIndex        =   53
      Top             =   8160
      Width           =   5136
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
      Left            =   6720
      TabIndex        =   52
      Top             =   8160
      Width           =   4980
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   8760
      TabIndex        =   51
      Top             =   960
      Width           =   2952
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
      Left            =   8280
      TabIndex        =   50
      Top             =   120
      Width           =   1548
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
      Left            =   7800
      TabIndex        =   49
      Top             =   396
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   9720
      TabIndex        =   48
      Top             =   80
      Width           =   950
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   9720
      TabIndex        =   47
      Top             =   384
      Width           =   1980
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
      Left            =   1428
      TabIndex        =   46
      Top             =   7800
      Width           =   5136
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
      TabIndex        =   45
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   6720
      TabIndex        =   44
      Top             =   1560
      Width           =   5052
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3240
      TabIndex        =   43
      Top             =   1560
      Width           =   3312
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
      TabIndex        =   42
      Top             =   1560
      Width           =   1296
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(分)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   41
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
         Size            =   10.8
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
      TabIndex        =   40
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
         Size            =   10.8
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
         Size            =   10.8
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
         Size            =   10.8
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
         Size            =   10.8
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
         Size            =   10.8
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
      TabIndex        =   35
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   216
      Index           =   24
      Left            =   1680
      TabIndex        =   34
      Top             =   7488
      Width           =   372
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   23
      Left            =   1230
      TabIndex        =   33
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   22
      Left            =   1290
      TabIndex        =   32
      Top             =   1515
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   20
      X1              =   1620
      X2              =   1764
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   19
      X1              =   1620
      X2              =   1764
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   18
      X1              =   1620
      X2              =   1764
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   17
      X1              =   1620
      X2              =   1764
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   16
      X1              =   1620
      X2              =   1764
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   21
      Left            =   1170
      TabIndex        =   31
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   20
      Left            =   1290
      TabIndex        =   30
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   19
      Left            =   1290
      TabIndex        =   29
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   216
      Index           =   18
      Left            =   1320
      TabIndex        =   28
      Top             =   5076
      Width           =   372
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   17
      Left            =   1290
      TabIndex        =   27
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   16
      Left            =   1290
      TabIndex        =   26
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
         Size            =   10.8
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
      TabIndex        =   25
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(kg)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   216
      Index           =   14
      Left            =   600
      TabIndex        =   24
      Top             =   1512
      Width           =   492
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
         Size            =   10.8
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
      TabIndex        =   23
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   22
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   20
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   19
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
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
      TabIndex        =   18
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   7
      Left            =   30
      TabIndex        =   17
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   6
      Left            =   30
      TabIndex        =   16
      Top             =   1515
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   390
      X2              =   534
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      X1              =   390
      X2              =   534
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   390
      X2              =   534
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   390
      X2              =   534
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   390
      X2              =   534
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      X1              =   390
      X2              =   534
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
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
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   5
      Left            =   30
      TabIndex        =   15
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   4
      Left            =   30
      TabIndex        =   14
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   3
      Left            =   30
      TabIndex        =   13
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   2
      Left            =   30
      TabIndex        =   12
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   10
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
      Left            =   1908
      TabIndex        =   8
      Top             =   1248
      Width           =   1272
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3216
      TabIndex        =   7
      Top             =   1248
      Width           =   8508
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
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
      Top             =   936
      Width           =   4572
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
      Top             =   924
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
      Left            =   7368
      TabIndex        =   3
      Top             =   72
      Width           =   276
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
      Left            =   5748
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
      Left            =   5748
      TabIndex        =   1
      Top             =   72
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
      Left            =   4488
      TabIndex        =   0
      Top             =   108
      Width           =   1272
   End
End
Attribute VB_Name = "LS21_TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    LS21_TC
'
'            update: 2002.6.28 s.f  private sub cal_pid　削除
'            update: 2002.6.29 s.f "DC" 書き換え
'                                  "HC" 新規追加
'　　　　　　　　　　　　　　　　　　difftime　書き換え
'            update: 2002.8.10 s.f roz(0),roz(1)を突当成形のﾊﾟﾗﾒｰﾀへ
'            update: 2002.8.18 s.f DC時　成形回数の戻し　追加　(i_s=i_s-1)
'            update: 2002.8.29 s.f cp,ct,ccデータ表示'
'            update: 2002.9.26 s.f ic(i)=10 で　終了判断　に　訂正
'            update: 2002.10.1 s.f 軸制御モード２へ、　CtlDisp  'DioOut 12,1  位置制御 '  02.10.1 追加
'            update: 2002.10.1 s.f 軸制御　エラー表示　Label2(4)からLabel2(3)へ変更
'            update: 2002.10.2 s.f 軸制御スタート時間表示
'            update: 2002.10.5 s.f タイムアップルーチン見直し（ｾｸﾞﾒﾝﾄ飛び対策）
'            update: 2002.10.5 s.f 時間表示変更
'            update: 2002.10.9 KYOCERA タイマー処理、タイムアップ、コメント表示、時間表示変更
'            update: 2002.10.12 s.f ﾀｲﾑｱｯﾌﾟの成立後　goto文　変更
'            update: 2002.10.16 KYOCERA ﾀｲﾑｱｯﾌﾟ処理 <9 を istend に変更
'            update: 2002.10.16 KYOCERA ﾀｲﾑｱｯﾌﾟで次のｽﾃｯﾌﾟ追加
'            update: 2002.10.17 KYOCERA 原点復帰後に初回原点復帰完了ﾌﾗｸﾞgOrgStartFlgをON
'            update: 2002.10.17 KYOCERA ﾀｲﾑｱｯﾌﾟ処理 <istend を 10 に変更
'            update: 2002.10.26 s.f 軸制御　エラー表示　Label2(3)からLabel2(5)へ変更
'            　　　　　　　　　　　　cc3-cc2表示　追加
'                                   SR　の処理変更　0.1秒に１回ｻﾝﾌﾟﾘﾝｸﾞ
'
'            update: 2002.11.1 s.f iPltMax 初期値　10　->　8　へ変更
'            update: 2002.12.4 s.f 成形データのsave
'            update: 2003.07.10 HND アラーム表示中の　成形プログラム続行
'                                  frmerr_sign, FbiDio, LS21_SC
'            update: 2003.07.19  s.f.  1回成形の回数　６－＞３へ変更　１号機のみ
'            update: 2003.09.11  s.f.  Plt1Jyun()へ　WaitSec 1.5　追加　（成形終了時　非常停止発生　対策）
'                                      'E'の処理に　genten　追加
'            update: 2004. 3. 8 s.f.  変更　成形軸制御モード　’７’追加　（上軸衝突判定付）
'                                    RecEmgDTsave 非常停止メッセージの保存
'            update: 2004. 3.12 s.f.  速度指令電圧　表示
'
'            update: 2004. 4.23 s.f.  timeupで非常停止
'            update: 2004. 4.24 s.f.  カウンタ、ﾀｸﾄﾀｲﾑ、表示　改造
'            update: 2004. 5. 5 s.f   温度係数、肉厚補正ルーチン　追加  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'            update: 2004.5.12  s.f   PGM_KTD　"ｵｰﾊﾞｰﾌﾛｰ"対策　　wTm0!,wTm1!  global化,  LS21_SCと　LS21_TC から　dim削除
'            update: 2004.5.17  s.f   'S'ｺﾏﾝﾄﾞ　バグ対策
'            update: 2004.5.18  s.f   バグ対策 & T係数表示
'            update: 2004.8.17  s.f   ｵｰﾊﾞｰﾌﾛｰ"対策  p(ist0)をppへ  ”：”複数の行を無くす
'            update: 2004.8.27 - 10.30 s.f   T係数関数変更、　　「ＤＣ　０」コマンド　成形前に型在否チェックセンサーのチェック機能追加
'            update: 2004.12.20 s.f   DＣ　０」コマンド　バグ修正
'            update: 2005. 5.25 s.f    Version No表示追加
'            update: 2005. 7.18 s.f    最終成形終了後　２０分の自然冷却
'            update: 2005. 9.28 s.f   T係数　表示色変更
'            update: 2005.11.22 s.f   Melec C-870 counter動作バグ修正　コンペアカウンタ値セット時　符号反転　　setcm1
'                                     C870sts(3) 周り　バグ修正, 右横データ表示順序変更
'            update: 2005.11.23 s.f   11/22 変更のバグ修正　成形軸制御　「C870sts　resetするまで　読み飛ばす」を　復活
'            update: 2005.12.17 s.f   Do-Loop 外の　DoEvent削除 OverFlow 対策 s.f.
'                                     コマンドの　evtime　取り込みを　コマンド開始時へ変更
'　　　　　　　　　　　　　　　　　　DCコマンド　LAコマンド　再チェック修正
'　　　　　　　　　　　　　　　　　　連続前コマンド　evtime　と　fintime　表記入れ替え
'            update: 2006. 3. 3 s.f  edit 使用時　do　loopから抜ける
'　　　　　　　　　　　　　　　　　　DCｺﾏﾝﾄﾞへ　fintime=timer　を　設置
'            update: 2006. 4.14 s.f  on error goto ,  sts as long
'            update: 2006. 4.15 s.f  error 表示
'            update: 2006. 5. 9 s.f  O.F.error 表示　軸制御　end3　追加,  tstime=0#
'            update: 2006. 5.18 s.f 　r_pres()の　DoEvents 　削除、　”J"、１秒に1回　Doevents　追加
'                                    非常停止　表示追加
'       Ver.3.33R_061221 2006.12.21 s.f  LS-33改　対応　　VacuumON、VacuumOFF　を廃止、SeikeiON,SeikeiOFF新設　DO3　割り当て変更
'       Ver.3.33R_070927 2007.09.27 s.f  Z補正　指定したｾｸﾞﾒﾝﾄNo.へ　できるようにする
'           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim lGphNo%
Dim lGphNo0%
Dim EditFlg%
Dim lViewFlg      '前の画面番号
Dim NextView%
Dim lEmgFlg%      '非常停止
Dim lDtSaveFlg%   'データ保存
Dim TCFlg%        'テスト成形中
'Dim iPltMax%      'パレット回転数    '05.7.18 globalへ
Dim l_stime!      'サイクルタイム
Dim lHO_Flg%      'HOコマンド用フラグ
Dim lHO_Time!     'HOコマンドの時間
'Dim lSokuFlg%     '強制ソークタイム
Dim CmndCol!(0 To 1)  'コマンド釦の色
Dim SokuCor!(0 To 1)  '強制ソークタイムのコマンド釦の色
'Dim T_keisuCol!(0 To 1)  '温度係数、肉厚補正表示のbackColor
Dim lCycleTime$       'サイクルタイム
'Dim sdt1$, sdt2$, sdt3$   2006.4.14 global he
'Dim iFlg_hijyou%    '非常停止フラグ  s.f. 2004.3.8   2009.8.17削除
Dim TCi_s%         ' 「１回成形」時の　成形回数
Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0  'キャンセル
  lGphNo = 0
  MoniGraph Me.Picture1, 0, lGphNo
Case 1  '終了
  If TCFlg = True Then         'テスト成形中
    FrmMenuFlg = False
    NextView = 1
  Else
    Unload Me
    PGM_Menu.Show
  End If
Case 2  'グラフ再描画
  lGphNo = lGphNo + 100
  MoniGraph Me.Picture1, 0, lGphNo
Case 3  'edit エディタ起動
  Unload Me
  MYEdit.Show
Case 4      '真空到達
  gVumFlg = 1                       '真空到達=1
Case 5      '"S" ;データセーブ
  lDtSaveFlg = True
Case 6      '成形開始
  iPltMax = 3    'パレット回転数
  Timer1.Enabled = False
  Command2(1).Caption = "中断"
  Command2(3).Enabled = False
  Command2(6).Enabled = False
  Command2(7).Enabled = False
  TC_Main
  Command2(3).Enabled = True
  Command2(6).Enabled = True
  Command2(7).Enabled = True
  Command2(1).Caption = "終了"
  Timer1.Enabled = True
Case 7      '成形開始
  iPltMax = Val(Text1(0))    'パレット回転数
  Timer1.Enabled = False
  Command2(1).Caption = "中断"
  Command2(3).Enabled = False
  Command2(6).Enabled = False
  Command2(7).Enabled = False
  TC_Main
  Command2(3).Enabled = True
  Command2(6).Enabled = True
  Command2(7).Enabled = True
  Command2(1).Caption = "終了"
  Timer1.Enabled = True
Case 8      '強制ソークタイム
  If lSokuFlg = True Then
          lSokuFlg = False          '強制ソークタイム　受付解除
          Command2(8).BackColor = SokuCor(0)
    Else
          lSokuFlg = True           '強制ソークタイム　受付
          Command2(8).BackColor = SokuCor(1)
  End If
End Select

End Sub

Private Sub SetData()

Dim l_sdt$

  Label2(0) = Format(ptime, "###0")  '測定時間
  Label2(1) = Format(ytemp, "###0")  '予備加熱温度
  Label2(2) = gcoxFlName             '制御ファイル名
  Label2(3) = hcomm(2)               'コメント
  '
'  Label2(13).Caption = Str(InitDat(11))   '成形カウンタトウタル      TC_main 内で処理
'/* ｼｮｯﾄ数ｻｲｸﾙﾀｲﾑ枠表示 */
'  l_sdt = Format(l_stime / 60, "0") & "分" & Format(Int(l_stime) Mod 60, "0") & "秒"    '2002.10.09 KYOCERA
'  Label2(9).Caption = Format(InitDat(10), "000")    '成形カウンタ i_s
'  Label2(8).Caption = l_sdt               'タクトタイム
' -----------------------------------
  DispGphScale
End Sub

Private Sub GetData()

End Sub

Private Sub Form_Load()
  DispCenter Me
  LS21_TC.Caption = LS21_TC.Caption + "     " + versionNo
  Me.Top = 0
  SokuCor(0) = &H8000000F     '強制ソークタイムのコマンド釦の色
  SokuCor(1) = &HFF&          '強制ソークタイムのコマンド釦の色 押されたとき
'  T_keisuCol!(0) = &HFFFFFF    '温度係数、肉厚補正　表示backcolor　off
'  T_keisuCol!(1) = &HFFFFC0    '温度係数、肉厚補正　表示backcolor　on
  lDtSaveFlg = False      'データ保存
'  'lSokuFlg = False        '強制ソークタイム
  If lSokuFlg = False Then
          Command2(8).BackColor = SokuCor(0)
    Else
          Command2(8).BackColor = SokuCor(1)
  End If
  lViewFlg = ViewFlg      '前の画面番号
  ViewFlg = 3             '画面番号
  FrmMenuFlg = True       'メニューから抜けるときfalse
  EditFlg% = False        'エディタ起動解除
  lEmgFlg = False         '非常停止
  TCFlg = False           'テスト成形中
  Command2(1).Caption = "終了"
  SetData
  TrnsReqON               '搬送依頼信号ＯＮ
  Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  HeatOFF       '/* 予備加熱をゼロにし、ＯＦＦする */
  CoolOFF
  ServoOFF
  TrnsReqOFF    '搬送依頼信号OFF
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
  'Label2(4).Caption = ""
  'Timer1.Enabled = False
  '-------------- ピラニ計読み
'  LS21S_Monitor        '2006.12.21 削除 s.f
  'LS21T_MAIN
End Sub
Public Sub LS21T_MAIN()   '/* １回成形 メインプログラム 2002.5.28a*/
'/*　　　　　　　　　　　　　　温度表示追加、　時間表示　追加　2002.6.15　*/
Dim i%, j%, js%, l%, ist0%, ist1%, ndata!, mdata!, ntemp!, mtemp!, ntemp0!, mtemp0!, iflg%, isflg%
Dim ied%, ips%, i_s%, irei%, r_ch%, ix%, ix0%, iy%, isp%, stime%, ii%, iii%, istend%
Dim ie02%, ie03%, ie04%, ituflg%, iSRcount
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%
Dim isub As Long, jsub As Long, ksub As Long, lsub As Long
Dim sdata!, sv%, zch%     '  05.11.26 s.s. overflow 対策
'Dim sdata%, sv%, zch%
Dim ct_dummy!, iz3%, itc%
'Dim m_l%, sdata%, sv%, zch%
Dim com$, tdate$, ttime$
Dim m_l!
Dim st!, ev!, sev!, fin!, it!          '/* 時間用データ */
Dim btemp!(0 To 4), bposi!, bpre! '/* 温度　位置　圧力 の前データ */
'
Dim sdt$, ch%, hdt%, flindex%, imax%, sts1%, sts2%, ch1%, ch2%
Dim sts As Long        ' 2006.4.14
Dim S_StartTime!
Dim stTime!, evtime!, sevTime!, mTime!, tsTime!, endTime!
'Dim stTime!, evtime!, fintime!, sevTime!, mTime!, tsTime!, endTime!
Dim dt!(0 To 7)
Dim flg As Long, cnt As Long
Dim diTime!, diTime1!, diTime2!, diTimeSR!, pdt!, pp!, pml!, x1dt!, x2dt!, pos!
Dim dmy$, idmy%, iwt!
Dim r_z_now!, r_z_ave!, r_z_dum!(0 To 180), it_ts%, i_ts%    ' /* 2002.4.9　追加　突当成形　*/
Dim epsilon!
Dim cp_z!, cc_time!(0 To 3), ct_temp!(0 To 2), ct!, ict%   ' CP , CT 用
Dim ct_t!(0 To 10)
'
 On Error GoTo errHandler:
  iFlg_hijyou = 0
  iz3 = 3   '  Z(ist0)　の　　Z3の　index値
  ips = 1
  i_s = -1              '成形回数
  For ii = 1 To 180: r_z_dum(ii) = 0#: Next ii
  tsTime = 0#
'  設定位置下へ移動
'  it_ts = roz(1)       ' 10     '/* 突き当て達成　ﾁｪｯｸ　平均する回数 mzx 180 */
'  epsilon = roz(0)     ' 0.0005 '/* 突当　許容幅　　mm　　*/
'
'----------------------- １回成形メインプログラム
  ppos = "TC"     ' 現在位置= TC
  C870Stop
  ServoON       '/* サーボｏｎ */
  CtlDisp       '位置制御
  'TrnsReqOFF    '搬送依頼信号OFF   SCの時
  TrnsReqON      '搬送依頼信号ＯＮ　TCの時
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　　　*************/
'/* SPEC INITIALIZE CMD OUT */
'/* カウンタボードの初期設定 */
  InitDat(10) = 0
'/* 加減速ﾚｰﾄｾｯﾄｺﾏﾝﾄﾞ */
  C870AccRate
'/* 速度設定 */
'  C870LSPDSet 300    '/* 300 pps 0.066mm/sec */
  C870LSPDSet 800    '/* 300 pps 0.066mm/sec */
'/* ディレータイム設定 */
  C870DelayTime
  rstcm1   '  compareter reset
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　終了  *************/

'/* ＡＴＣ温度リセット */
'/* ロボットデータのフロッピーからの読みとり */
  rozFileLoad
'
  it_ts = Int(roz(1))  ' 10       '/* 突き当て達成　ﾁｪｯｸ　平均する回数 max180*/
  epsilon = roz(0)    ' 0.0005   '/* 突当　許容幅　　mm　　*/
'
st:
  If ied = 2 Then GoTo st2:
'/*  制御ファイルのオープン */
  coxDtRead gcoxFldir & gcoxFlName
  Label2(0).Caption = Format(ptime, "0")
  If T_keisuCont(2) <> 0 Then T_keisuCont(1) = T_keisuCont(2) 'ポインターbackup
  If T_keisuCont(3) <> 0 Then T_keisuCont(0) = T_keisuCont(3) '型個数 backup
'/* グラフィック画面の初期化 */
  InitDat(8) = ptime  'グラフスケール経過時間 (Max)
  SetData
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* 温度係数　肉厚補正の表示 */
  For itc = 0 To 9
    Label4(itc).Caption = Format(T_keisu(itc), "0.000")
    Label6(itc).Caption = Format(Z3_Hosei(itc), "0.000")
    If itc < T_keisuCont(0) Then
       Label4(itc).BackColor = T_keisuCol!(1)
       Label6(itc).BackColor = T_keisuCol!(1)
      Else
        Label4(itc).BackColor = T_keisuCol!(0)
        Label6(itc).BackColor = T_keisuCol!(0)
    End If
  Next itc
'
'/* 軸駆動制御コマンドのファイルからの読み取り */
  i = 0
  Do
    Label2(12).Caption = DispCtrlCommand(i)
    If pres(i) >= 1000 Then ips = 2         '/* ﾌﾟﾚｽ圧が1ton以上で軸変更 */
    i = i + 1                               '/* ipsは軸自動描画時のスケールpara*/
    If ic(i - 1) = 9 Then Exit Do
  Loop

  ic(i) = 10                                ' /* 最終セグメント＋１に　「10」をセット */
  istend = i                           ' 軸制御ｺﾏﾝﾄﾞのend番号
'ic(i) = 4
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
  GCnt0 = 0
  GCnt1 = 0
'/* 自動運転認識 */
  Label2(4).Caption = "自動運転認識中"
  ch1 = 1            'システムレディー
  ch2 = 3            '自動
  Do
    DoEvents
    If FrmMenuFlg = False Then GoTo eend:            'メニューから抜けるときfalse
'    LS21S_Monitor     '-------------- ピラニ計読み 真空なら    '2006.12.21 削除 s.f
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
'    LS21S_Monitor    '2006.12.21 削除 s.f
    'flindex = flindex + 1
    com = Left(scom(flindex), 1)
    isub = sisub(flindex)
    sdt = Right("    " & scom(flindex), 2)
    sdt = sdt & Right(Space(15) & Format(isub, "0"), 15)
    If com = "S" Then
      jsub = sjsub(flindex)
      ksub = sksub(flindex)
      lsub = slsub(flindex)
      sdt = sdt & Right(Space(15) & Format(jsub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(ksub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(lsub, "0"), 15)
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
        Label2(4).Caption = "CASE B DO1"
        'Exit Do
      Case "N"    '/* 窒素ガスの注入 */
        If Mid(scom(flindex), 2, 1) = "S" Then
          If isub = 1 Then
            Label2(4).Caption = "窒素ガス注入 DO1"
            N2Open
          End If
          If isub = 0 Then
            Label2(4).Caption = "窒素ガス停止 DO1"
            N2Close
          End If
        End If
      Case "J"    '/* 時間待ち */
        Label2(4).Caption = "時間待ち DO1"
        evtime = Timer
        Do
          fintime = Timer2func     ' 2009.8.17
'          fintime = Timer
          DoEvents
          Label2(10).Caption = Format(diffTime(fintime, evtime), "0")
          If diffTime(fintime, evtime) >= isub Then Exit Do
        Loop
        Label2(10).Caption = ""
      Case "K"    '/* 加熱 */
        Select Case Int(isub)
        Case 1
          Label2(4).Caption = "加熱　ＯＮ DO1"
          HeatON
        Case 0
          Label2(4).Caption = "加熱　ＯＦＦ DO1"
          HeatOFF
        End Select
      Case "S"    '/* ＡＴＣ温度設定 */
        Label2(4).Caption = "ＡＴＣ温度設定 DO1"
        evtime = Timer              '待ち初めの時間
        ntemp0 = isub
        mtemp0 = jsub
        ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
        mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
        Do
          fintime = Timer2func     ' 2009.8.17
'          fintime = Timer
          DoEvents
          diTime = diffTime(fintime, evtime)
          If ksub <> 0 Then x1dt = diTime / ksub
          ndata = (ntemp0 - ntemp) * x1dt + ntemp
          mdata = (mtemp0 - mtemp) * x1dt + mtemp
          TempSet 2, ndata
          TempSet 3, mdata
          If diTime >= ksub Then Exit Do
        Loop
        ntemp = ntemp0
        mtemp = mtemp0
        TempSet 2, ntemp
        TempSet 3, mtemp
      Case "R"    '/* 冷却 */
        Select Case Int(isub)
        Case 0    '冷却大　ＯＦＦ
          Label2(4).Caption = "冷却大　ＯＦＦ DO1"
          CoolOFF
        Case 1    '冷却大　ＯＮ
          Label2(4).Caption = "冷却大　ＯＮ DO1"
          CoolON
        Case 2    '冷却小　ＯＮ
          Label2(4).Caption = "冷却小　ＯＮ DO1"
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
'  Label2(5).Caption = Format(roz(0), "0.0000")     '/* 突当成形para　幅　 */　02.10.26 s.f 削除
'  Label2(6).Caption = Format(roz(1), "0.0")     '/* 突当成形para　時間 */　02.10.26 s.f 削除
'/* 成形開始 */
  Do        '----------------- DO LOOP
    DoEvents
    i_s = i_s + 1                   ' /* i_s = 成形回数 */
    js = 0
    ist0 = -1
    ist1 = -1           '/* ist0 ist1　(初期値 -1) 軸制御ｺﾏﾝﾄﾞの現在番号 */
    ie0 = 0
    ie1 = 0
    ie2 = 0
    ie3 = 0
    S_StartTime = Timer
    stTime = Timer
    sevTime = Timer
    diTimeSR = -9999.99                        ' 温度設定　ＳＲの初期化
    iSRcount = 1                               ' 温度設定　ＳＲの初期化
    For ii = 0 To 10
      ct_t(ii) = 0
    Next ii    ' 温度設定　ＳＲの初期化
'
    Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  文字　ピンク(ポインター）
    Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  文字　ピンク(ポインター）
    Label4(T_keisuCont(1) - 1).BorderStyle = 1    '  枠付きにする(ポインター）
    Label6(T_keisuCont(1) - 1).BorderStyle = 1  '  枠付きにする(ポインター）
    iz3 = Z3_HoseiCont(2)   ' Z補正　を実施する　ZNo.　　　’07.9.27　追加
    z(iz3) = z(iz3) + Z3_Hosei(T_keisuCont(1) - 1) '  ”Z3"の補正値set
'/* カウンタへの出力ｕｐ */
'                                             TC_main で実施
'    If i_s <> 0 Then
'      InitDat(11) = InitDat(11) + 1  '成形カウンタトウタル
'      InitDtSave
'      Label2(13).Caption = Str(InitDat(11))   '成形カウンタトウタル
'    End If
'/* 成形枠の表示 */
ejs1:
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* Ｘ軸の表示 */
'/* Ｙ軸の表示 */
'/* ｼｮｯﾄ数ｻｲｸﾙﾀｲﾑ枠表示 */
    sdt = Format(stime / 60, "0") & "分" & Format(Int(stime) Mod 60, "0") & "秒"        '2002.10.09 KYOCERA
'    Label2(9).Caption = Format(i_s, "000")
'    Label2(8).Caption = sdt         'サイクルタイム
    lCycleTime = sdt                'サイクルタイム
    InitDat(10) = i_s               '成形カウンタ
'
'    For iii = 0 To 9
'       Label6(iii).Caption = ""
'    Next iii
'
'/* カウンタへの出力ダウン */
    'InitDat(11) = InitDat(11) - 1   '成形カウンタトウタル
    'InitDtSave
    'Label2(13).Caption = Str(InitDat(11))
'/* データの取り込み */
'    stTime = Timer            DO loop 開始直後へ　移動　10/5
    evtime = Timer
    iflg = 1
    ied = 0
    ttime = Time
    mTime = Timer
'----------------------------- For Loop i
    imax = ptime * 60
    For i = 1 To imax
start:
    'finTime = Timer
'    DoEvents           '2005.12.17  s.f.
    ituflg = 0            '　タイムアップflgのリセット10/5
'/* 成形軸のドライブ*/
      If ist0 > 0 Then
       'If ic(ist0 - 1) = 4 Then     '/* 軸制御終了時のﾀﾞﾐｰｾｸﾞﾒﾝﾄ */
        If ic(ist0 - 1) = 10 Then    '/* 「最終ｾｸﾞﾒﾝﾄ+1」は、　「10」*/
          ist0 = ist0 - 1            '/* ループ素通りのためのﾀﾞﾐｰ */
        End If
      End If
        sdt3$ = DispSegm(ist0)
        Label2(12).Caption = sdt3$
      If ist0 <> ist1 Then
        gOrgFlg = False                '原点復帰完了=TRUE
        ist1 = ist0
        sevTime = Timer '            開始時間の取り込み
'
        If (ist0 > 0 And ist0 < 11) Then   '　開始時間の表示　ｄｅｂｕｇ用
           diTime1 = diffTime(sevTime, stTime)          '2002.10.09 KYOCERA
           sdt = Format(ist0, "0") & "=" & Format(Int(diTime1 / 60), "0") & "'" & Format(Int(diTime1) Mod 60, "0") & "秒"       '2002.10.09 KYOCERA
'           Label6(ist0 - 1).Caption = sdt
        End If
'
        Select Case ic(ist0)  '-------- 軸制御モード番号
        Case 0, 8   '-------------------- 位置制御
          ppos = "TC JikuStart 0 8"
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9); サーボON & 速度上限S12
          s_drive z(ist0), vel(ist0)
        Case 1, 3, 7   '-------------------- 速度制御    2004.3.8 「7」追加
          ppos = "TC JikuStart 1 3 7"
          m_l = vel(ist0)
          'm_l = vel(ist0) / 100
          If m_l > 50 Then m_l = 50
          setcm1 z(ist0)
          Ready_Wait    '
          CtlVelo       'outp(DIO_P+3,5);
          'while((inp(XCN_DT1)&0x01)!=0);
          Do    '' 「カウンター一致」状態脱出用
            DoEvents
            sts = C870Sts(3)    'sts=1の時　成立＝＞「-1」　sts=0の時不成立＝＞「0」
            If (sts And &H1) = 0 Then Exit Do   'PULSE と COMPARE が一致
          Loop
          '
        Case 2    '-------------------- ダミー
          ppos = "TC JikuStart 2"
          Ready_Wait
          CtlDisp  'DioOut 12,1  位置制御 '  02.10.1 追加
          Ready_Wait    '
          ServoON     'outp(DIO_P+3,1);
        Case 9    '-------------------- 終了
          ppos = "TC JikuStart 9"
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9);
          genten
          'Ready_Wait
          For ii = 1 To 180         '/* 制御３用のの初期化 */
            r_z_dum(ii) = 0#
          Next ii
          i_ts = 1
          r_z_ave = 0#
        End Select
      End If
'
        fintime = Timer2func     ' 2009.8.17
'        fintime = Timer         '2002.10.09 KYOCERA

'/* タイムアップ処理 */
          '2002.10.09 KYOCERA
      If ist0 < 0 Then GoTo sj1:
      'If ituflg = 0 Then
          If ((ic(ist0) < 10) And (diffTime(fintime, sevTime) > t0(ist0))) Then  '2002.10.16 KYOCERA 2002.10.17 KYOCERA            '10/4
            ituflg = 1
            sdt = "タイムアップ" & Right(Space(11) & Format(diffTime(fintime, sevTime), "0.0"), 11)
            sdt = sdt & Right(Space(11) & Format(t0(ist0), "0.0") & Format(ist0 + 1, "0"), 11)
            Label2(5).Caption = sdt + "TUp=" + Str(gTimeUpCnt) & Str(ist0) & "  時刻;" & Format(Now, "hh:mm:ss")
'
                RecEmgDtSave sdt3, sdt1, sdt2
                gemgmsg = "ﾀｲﾑｱｯﾌﾟ"
                hijyou        '非常停止処理
                iFlg_hijyou = 1        '　ﾀｲﾑｱｯﾌﾟ
                GoTo eend:
'
            ist0 = ist0 + 1             '/タイムアップで次のステップ   '2002.10.16 KYOCERA
            'GoTo TimeUpEnd:
            GoTo jscmdend:              '　終了信号処理を飛び越す    10/12 sf
          End If
      'Else                          ' ダブルチェック　１．２秒後に再確認
          'If ((ic(ist0) < 9) And (diffTime(finTime, sevTime) > (t0(ist0) + 1.2))) Then            '10/4
            'sdt = "タイムアップ" & Right(Space(11) & Format(diffTime(finTime, sevTime), "0.0"), 11)
            'sdt = sdt & Right(Space(11) & Format(t0(ist0), "0.0") & Format(ist0 + 1, "0"), 11)
            'gTimeUpCnt = gTimeUpCnt + 1    'タイムアップのカウンタ
            'label2(5).Caption = sdt + "TUp=" + Str(gTimeUpCnt) & Str(ist0)
            'ist0 = ist0 + 1             '/タイムアップで次のステップ
            'hijyou        '非常停止処理
            'getch
            'GoTo eend:
            'ituflg = 0
            'GoTo jscmdend:              '　終了信号処理を飛び越す    10/4 sf
          'End If
      'End If
TimeUpEnd:
'
'/* 終了信号の処理 */
      Select Case ic(ist0)
      Case 0, 8   '/* 位置制御の場合 */
          ppos = "TC JkE0 8"
        If (C870Sts(1) And 1) = 0 Then
           Label2(5).Caption = "位完sg=" & Str(ist0 + 1)  'ｾｸﾞNo.=ist0+1 10/4  sf
           ist0 = ist0 + 1
        End If
      Case 1    '/* 速度制御の場合 */
          ppos = "TC JkE1"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
          ppos = "TC JkE1 -1cal"
        cal_pid pdt, pp, pml
          ppos = "TC JkE1 cal_pid"
        sts = C870Sts(3)  'status3 を読む
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* 位置達成で終了 */
            Label2(5).Caption = "位置制御 pass CNT " & Str(ist0)    '11/2 s.f
            rstcm1   '  compareter reset
            Ready_Wait    '
         Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "位置 pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "TC JkE1 r_z -1"
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "位置制御 pass PC " & Str(ist0)    '11/2 s.f
'          End If
'          ppos = "TC JkE1 r_z -1"
'        End If
      Case 3    '/* 速度制御　突当成形の場合  2002.4.9 */
          ppos = "TC JkE3"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
          ppos = "TC JkE3 -1cal"
        cal_pid pdt, pp, pml
          ppos = "TC JkE3 cal_pid"
        sts = C870Sts(3)  'status3 を読む
          ppos = "TC JkE3 sts=C870"
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* 位置達成で終了 */
          Label2(5).Caption = "位置制御 pass CNT " & Str(ist0)    '11/2 s.f
            rstcm1   '  compareter reset
            Ready_Wait    '
         Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "位置 pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "TC JkE3 r_z -1"
'
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "位置制御 pass PC " & Str(ist0)    '11/2 s.f
'          End If
'            ppos = "TC JkE3 r_z -1"
'        End If
'        If r_z() < z(ist0) Then
'          r_z_now = r_z()
'          If Int(tsTime) <> Int(mTime) Then
'              ppos = "TC JkE3 r_z -2"
'            tsTime = mTime                  '/* １秒前と、２秒前と */
'            If Abs(r_z_now - r_z_ave) < epsilon Then
'              ist0 = ist0 + 1               '/* it_ts回連続　epsilon以下 */
'            Else                            '/* で　突当達成で終了 */
'              r_z_dum(i_ts) = r_z_now
'              r_z_ave = 0#
'              For ii = 1 To it_ts
'                 r_z_ave = r_z_ave + r_z_dum(ii)
'              Next ii
'              r_z_ave = r_z_ave / it_ts
'              i_ts = i_ts + 1
'              If i_ts > it_ts Then i_ts = 1
'            End If
'          End If
'        End If
'
'
          If r_z() < z(ist0) Then
'            r_z_now = r_z()                    '2008.2.23 移動
              ppos = "TC JkE3 r_z -2"
            If Int(tsTime) <> Int(mTime) Then
              tsTime = mTime                  '/* １秒前と、２秒前と */
              r_z_now = r_z()                    '2008.2.23 移動
              If Abs(r_z_now - r_z_ave) < epsilon Then
                ist0 = ist0 + 1               '/* it_ts回連続　epsilon以下 */
              Else                            '/* で　突当達成で終了 */
                r_z_dum(i_ts) = r_z_now
                r_z_ave = 0#
                For ii = 1 To it_ts
                   r_z_ave = r_z_ave + r_z_dum(ii)
                Next ii
                r_z_ave = r_z_ave / it_ts
                i_ts = i_ts + 1
                If i_ts > it_ts Then i_ts = 1
              End If
            End If
          End If
      Case 7    '/* 速度制御　上軸衝突判定付　　　　　　　　　2004.3.8 s.f. 軸制御「７」追加　　ここから　*/
'　　　　　　　/*　指定圧力より高い圧力が３秒以上続いたら非常停止　　*/
        ppos = "TC JkE7"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
        cal_pid pdt, pp, pml
        sts = C870Sts(3)  'status3 を読む
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* 位置達成で終了 */
          Label2(5).Caption = "位置 pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
            Ready_Wait    '
'          Do            'このdo　loop　意味無い　　2005.11.22  s.f
'            DoEvents
''            sts = C870Sts(3)          'status3 を読む    10/4  sf
''            If (sts And &H1) <> 0 Then Exit Do           10/4 sf
'             If r_z() >= z(ist0) Then Exit Do             '10/4
'          Loop
         Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "位置 pass PC " & Str(ist0)
             End If
           End If
         End If
'
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "位置 pass PC " & Str(ist0)
'          End If
'        End If
'        bpre = r_pres()
'        If bpre > pdt Then              ' 2008.2.18  変更
'        If bpre > pdt * 0.7 Then
'          If Int(tsTime) <> Int(mTime) Then
'            tsTime = mTime                  '/* １秒前と比較 */
'            i_ts = i_ts + 1               '/* i_ts回連続して　圧力が指定値以上 */
'
'
        If Int(tsTime) <> Int(mTime) Then '2008.2.23 変更 1秒に1回チェック
          tsTime = mTime                  '/* １秒前と比較 */
          bpre = r_pres()
          If bpre > pdt Then                ' 2008.2.18 変更
'               If bpre > pdt * 0.7 Then
             i_ts = i_ts + 1               '/* i_ts回連続して　圧力が指定値以上 */
             If i_ts > 3 Then
                gemgmsg = "軸制御7 error"
                hijyou        '非常停止処理
                'getch
                iFlg_hijyou = 2     '軸制御　7　error
                GoTo eend:
             End If
          End If
        End If                                 '/*     '2004.3.8　ここまで　*/
      Case 9    '終了
        ppos = "TC JkE9"
        sts = C870Sts(1)
        If (sts And 1) = 0 Then
          ist0 = ist0 + 1     '/* 完了 */
          If Abs(r_z()) > 0.1 Then
            Label2(5).Caption = "原点不良"
            ist0 = ist0 - 1
            genten              '原点出し
          End If
        Else
          '/* カウンタにゼロを書き込む */
          Ready_Wait
          Counter0
        End If
      End Select
'
'      Select Case ic(ist0)           ' 2004.3.12 s.f  05.11.26 nuku
'        Case 1, 3, 7
'                Label8(0).Caption = nout
'                Label8(1).Caption = v
'      End Select
'
jscmdend:                               '軸制御コマンド　ｅｎｄ  10/4 sf
'
'/* エラー表示 */
    If ArmChk <> 0 Then               'アラームメッセージ
      frmerr_sign.Show   'ALM出力
    Else
      Unload frmerr_sign
    End If
    
'    If ArmChk <> 0 Then   'アラームメッセージ　　'03.7.10上記に変更
'      frmerr_sign.Show 1　　　　　　　　　　　　'03.7.10上記に変更
'    End If　　　　　　　　　　　　　　　　　　　'03.7.10上記に変更

'/* プロセス実行 */
sj1:
    If iflg = 1 Then
      com = scom(js + flindex)
      isub = sisub(js + flindex)
      jsub = sjsub(js + flindex)
      ksub = sksub(js + flindex)
      lsub = slsub(js + flindex)
      js = js + 1
'
      evtime = Timer        ' 2005.12.17  s.f.  開始時間取り込み
'
      sdt = com & Right(Space(7) & Format(isub, "0"), 7)
      'locate(60,1);printf("%2s %5ld ",com,isub);
      If Left(com, 1) = "S" Then
        sdt = sdt & Right(Space(7) & Format(jsub, "0"), 7)
        sdt = sdt & Right(Space(7) & Format(ksub, "0"), 7)
        sdt = sdt & Right(Space(7) & Format(lsub, "0"), 7)
        'printf("%5ld %5ld",jsub,ksub);
        'Label2(7).Caption = sdt
      Else
        sdt = sdt
      End If
      Label2(7).Caption = sdt
    End If
        'システムレディ/* 非常停止の場合は成形中止 */
        sts1 = SystemReadyChk()   'システムレディ or 非常停止
        sts2 = AutoChk()          '自動状態？ (<>0 自動)
        If sts1 = 0 Or sts2 = 0 Then
          Label2(4).Caption = "自動運転終了"
          gemgmsg = ArmEmgMsgChk$()
          iFlg_hijyou = 10              '非常停止ﾒｯｾｰｼﾞのｓａｖｅ
          FrmMenuFlg = False              'メニューから抜けるときfalse
          NextView = 1
          Exit Do
        End If
        '
        Select Case Left(com, 1)
'
        Case "D"    '------------ 成形室の型の有無
            ppos = "TC Proc D"
            If (isub = 0) Then     '在否センサー　チェック
              If (KataChk() > 0) Then                '  2004.10.30  型在否チェック用センサの動作確認用
                 sdt = "DC　在否センサー異常（型有り！！）"
                 Label2(5).Caption = sdt
'
                  RecEmgDtSave sdt3, sdt1, sdt2
                  gemgmsg = "DC error 型有り"
                  hijyou        '非常停止処理
                  iFlg_hijyou = 3     '   DC error  型有り
                  GoTo eend:
              Else
                GoTo scend:
              End If
            End If                           '  2004.10.30  型在否チェック用センサの動作確認用
'
           If KataChk() < 3 Then '型が無い
            Label2(4).Caption = "CASE D 成形室型無し DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '現在時間　　2006.3.3　追加　s.f.
            If (diffTime(fintime, evtime) < isub) Then
               iflg = 0             ' 時間未達の場合
            Else
               idmy = js            '　時間待ち終了の場合
                 Do
                   DoEvents
                   dmy = scom(idmy + flindex)          '　次のコマンドを読み取る
                   If "LA" = dmy Then  '----- コマンドLAまで進める
                     js = idmy          '　　LAが見つかったら　次のコマンドNo.を　LAの　No.にセット
                     '------------- LAが見つかったら次に、セグメントをモード８まで（9の２つ前まで）進める
                     Do
                       DoEvents
                       If ic(ist0) = 8 Then
                         ist0 = ist0 - 1
                         sevTime = Timer        '  2005.12.17 Timeup防止 念のため s.f.
                         Exit Do
                       End If
                       ist0 = ist0 + 1
                       If ist0 > 50 Then   'エラー
'
                         sdt = "DCｺﾏﾝﾄﾞ ist0 > 50 ｴﾗｰ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                         gemgmsg = "DC error 4"
                         hijyou        '非常停止処理
                         iFlg_hijyou = 4     '"DC error 4"
                         GoTo eend:
'
                       End If
                     Loop
                   '
                     Exit Do
                   End If
                   idmy = idmy + 1
                   If idmy > 50 Or "EN" = dmy Then 'エラー
'
                         sdt = "DCｺﾏﾝﾄﾞ ist0 > 50 ｴﾗｰ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                          gemgmsg = "DC error 5"
                        hijyou        '非常停止処理
                         iFlg_hijyou = 5     '"DC error 5"
                         GoTo eend:
'
                   End If
                 Loop
'
              iflg = 1                    '　ｺﾏﾝﾄﾞ終了処理
              i_s = i_s - 1             '成形回数の戻し
'              evtime = Timer          '2005.12.17  s.f.
            End If
          End If                    '　型がある場合はそのまま抜ける
'
        Case "L"    ' 成形室に型が無かった時の飛び先　ﾀﾞﾐｰ
            ppos = "TC Proc L"
            Label2(4).Caption = "CASE L 成形室型無し DO2"
          'iflg = 1　　　　　　常時　iflg=1　の為、必要なし
'
        Case "H"    ' 強制ソーク　　　”ＨＣ”
            ppos = "TC Proc H"
            Label2(4).Caption = "強制ソーク DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '現在時間　　2006.3.3　追加　s.f.
            If (lSokuFlg = True And diffTime(fintime, evtime) < isub) Then
               iflg = 0
             Else
               iflg = 1
               lSokuFlg = False
               Command2(8).BackColor = SokuCor(0)
'               evtime = Timer          '2005.12.17  s.f.
            End If
'
        Case "S"    '/* ＡＴＣ温度設定 */
            ppos = "TC Proc S"
            Label2(4).Caption = "ＡＴＣ温度設定 DO2"
'            DoEvents          '2005.12.17  s.f.削除　　2006.5.18追加　削除
            If Mid(com, 2, 1) = "R" Then             ' SRの場合  注：関連初期化　Do　Loop　Topにあり
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer
               diTime = diffTime(fintime, stTime)    ' 0.1秒に１回温度取り込み（５回実施）
               If ((diTime - diTimeSR) > 0.1) Then
'                   ct_t(0) = ct_t(0) + TempRdMold() '温度読込
                   ct_dummy = TempRdMold(0)    'スリーブ温度読込
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(0) = ct_t(0) + ct_dummy '温度読込
                   iSRcount = iSRcount + 1
                   diTimeSR = diTime
                   iflg = 0
                   If iSRcount > 5 Then
                      ct_t(0) = ct_t(0) / 5
                      ntemp0 = isub
                      ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1)) 'ntemp0
                      mtemp0 = jsub
                      mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1)) 'mtemp0
                      ntemp0 = ct_t(0) + ntemp0
                      mtemp0 = ct_t(0) + mtemp0
                      ntemp = ntemp0
                      mtemp = mtemp0
                      TempSet 2, ntemp
                      TempSet 3, mtemp
                      ct_t(0) = 0
                      Label2(5).Caption = "SR= " & Format(Int(ntemp), "000") & Format(Int(mtemp), "  000")
                      iSRcount = 1
                      iflg = 1
'                      evtime = Timer          '2005.12.17  s.f.
                   End If
               End If
            Else
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer
              diTime = diffTime(fintime, evtime)        'SAの場合
              If ksub <> 0 Then x1dt = diTime / ksub
              ntemp0 = isub
              mtemp0 = jsub
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
              ndata = (ntemp0 - ntemp) * x1dt + ntemp
              mdata = (mtemp0 - mtemp) * x1dt + mtemp
              TempSet 2, ndata
              TempSet 3, mdata
              If diTime >= ksub Then
                iflg = 1
                ntemp = ntemp0
                mtemp = mtemp0
                TempSet 2, ntemp
                TempSet 3, mtemp
'                evtime = Timer          '2005.12.17  s.f.
              Else
                iflg = 0
              End If
            End If
'
        Case "P"    '/* 移動軸制御の駆動 */
            ppos = "TC Proc P"
            Label2(4).Caption = "移動軸制御の駆動 DO2"
          If Mid(com, 2, 1) = "W" Then
            Beep
            ist0 = ist0 + 1
            sevTime = Timer          '2005.12.17  s.f.
'           evtime = Timer          '2002.10.09 KYOCERA          '2005.12.17  s.f.
          End If
          If Mid(com, 2, 1) = "R" Then
            iflg = 0
            If ist0 <> ist1 Then iflg = 1
            If isub = 4 And ist0 = 0 Then iflg = 1
'            If iflg = 1 Then evtime = Timer         '2002.10.09 KYOCERA          '2005.12.17  s.f.
          End If
          'evTime = Timer
        Case "K"    '/* 加熱 */
          ppos = "TC Proc K"
          Select Case isub
          Case 1
            Label2(4).Caption = "加熱 ON DO2"
            HeatON
          Case 0
            HeatOFF
            Label2(4).Caption = "加熱 OFF DO2"
          End Select
        Case "N"
            ppos = "TC Proc N"
            Label2(4).Caption = "CASE N DO2"
          If Mid(com, 2, 1) = "S" Then
            If isub = 1 Then hdt = hdt
            If isub = 0 Then hdt = hdt
          End If
        Case "R"    '/* 冷却 */
          ppos = "TC Proc R"
          Select Case isub
          Case 2
            Label2(4).Caption = "冷却 ON1 DO2"
            CoolON
          Case 1
            Label2(4).Caption = "冷却 ON2 DO2"
            CoolON
          Case 0
            Label2(4).Caption = "冷却 OFF DO2"
            CoolOFF
          End Select
        Case "T"    '/* ＡＴＣ１の温度の読み取り */
            ppos = "TC Proc T"
            Label2(4).Caption = "ＡＴＣ１の温度の読み取り DO2"
          sdata = TempRdMold(0)    'スリーブ温度
          sdata = T_keisu_cread(sdata, T_keisu(T_keisuCont(1) - 1)) 'ndata
          If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Or (Mid(com, 2, 1) = "E" And (sdata > (isub + 20) Or sdata < (isub - 20))) Then
'          If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Then
            iflg = 0
          Else
            If iflg = 2 Then iflg = 1 Else iflg = 2
'            evtime = Timer          '2005.12.17  s.f.
          End If
        Case "J"    '/* 時間待ち */
          ppos = "TC Proc J"
          DoEvents      '2006.5.18  s.f 追加
            Label2(4).Caption = "時間待ち DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '現在時間　　2006.3.3　追加　s.f.
          diTime1 = diffTime(fintime, stTime)
          diTime2 = diffTime(fintime, evtime)
          If (Mid(com, 2, 1) = "S" And diTime1 >= isub) Or (Mid(com, 2, 1) = "C" And diTime2 >= isub) Then
            iflg = 1
'            evtime = Timer          '2005.12.17  s.f.
          Else
            iflg = 0
          End If
'
        Case "C"
          ppos = "TC Proc C"
          Select Case Mid(com, 2, 1)
          Case "P"    '成形終了位置　チェック
            Label2(4).Caption = "成形終了位置　チェック DO2"
            cp_z = r_z()
            Label5(0).Caption = " cp=" & Format(cp_z, "0.000")
'
          Case "C"    '　時間チェック
            Label2(4).Caption = "時間チェック DO2"
            If isub > 3 Then
                ict = 5
              Else
                ict = isub + 2
            End If
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '現在時間　　2006.3.3　追加　s.f.
            cc_time(isub) = diffTime(fintime, stTime)
            sdt = " cc" & Format(isub, "0") & "= " & Format(Int(cc_time(isub) / 60), "0") & "分" & Format(Int(cc_time(isub)) Mod 60, "0") & "秒"        '2002.10.09 KYOCERA
            Label5(ict).Caption = sdt
            If isub = 3 Then
                diTime1 = diffTime(cc_time(isub), cc_time(isub - 1))
                sdt = " cc3-2=  " & Format(Int(diTime1 + 0.5), "0") & "秒"
                Label5(6).Caption = sdt
            End If
'
          Case "T"    '　温度チェック
            Label2(4).Caption = "温度チェック DO2"
            If isub > 2 Then
                ict = 2
              Else
                ict = isub
            End If
            ct_temp(isub - 1) = TempRdMold(0) '温度 0V-300℃ 1V-1300℃
            ct_temp(isub - 1) = T_keisu_cread(ct_temp(isub - 1), T_keisu(T_keisuCont(1) - 1))
            sdt = " ct" & Format(isub, "0") & "=" & Format(ct_temp(isub - 1), "0.0") & "℃"
            Label5(ict).Caption = sdt
          End Select
'
        Case "X"    '搬送終了信号（成形開始）
          ppos = "TC Proc X"
          Select Case Mid(com, 2, 1)
          Case "R"    '成形開始 [搬送終了まで待つ]
            Label2(4).Caption = "成形開始 [搬送終了まで待つ] DO2"
            '--------------------- TCで削除
            'TrnsReqON  '搬送依頼信号Ch21出力
            '
            'Do
              '-------------- ピラニ計読み
            '  LS21S_Monitor
              'DioInput 13, sts        '搬送終了？
            '  sts = TrnsFinChk()      '搬送終了？
            '  If sts = 1 Then
            '    TrnsReqOFF            '搬送依頼信号ＯＦＦ
            '    Label2(4).Caption = "搬送依頼信号ＯＦＦ"
            '    Exit Do
            '  End If
            '  DoEvents
            'Loop
            '--------------------- TCで削除
          Case "W"    '成形終了
            Label2(4).Caption = "成形終了 DO2"
          End Select
        Case "E"    '/* 終了　ロボット搬送 */
            ppos = "TC Proc E"
           If r_z() > 2 Then                                      '03.9.11
              genten                                              '03.9.11
              'Ready_Wait    'while((inp(AX_STS)&1)!=0);          '03.9.11
            End If
             Label2(4).Caption = "終了 ロボット搬送 DO2"
        '--------------------- TCで削除
           iflg = 99
           GoTo send:
        '  Exit Do
scend:
        End Select
cjump:
  '-------------- ピラニ計読み
'  LS21S_Monitor      '2006.12.21 削除 s.f
  lEmgFlg = SystemReadyChk()  '非常停止
  If Int(mTime) = Int(Timer) And lEmgFlg <> 0 Then GoTo start:
  'lEmgFlg = EmgChk()         '非常停止
  'If Int(mTime) = Int(Timer) And lEmgFlg = False Then GoTo start:
'                 /* 1秒に1回下に抜ける */
      mTime = Timer
      ppos = "TC 1 sec Disp 1"
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
'           /* 圧力　ＰＩＤ制御　Ｐ＞１５　なら速度　ゼロ */
  If ist0 >= 0 Then
    If p(ist0) > 15 Then
      DaVoltOut 1, 0        ' 0V D/A ch=1
    End If
  End If
'/*　温度取り込み */
'    DoEvents          '2005.12.17  s.f.
    atemp(i, 0) = TempRdMold(0)   'スリーブ温度 0V-300℃ 1V-1300℃
    atemp(i, 0) = T_keisu_cread(atemp(i, 0), T_keisu(T_keisuCont(1) - 1))
    atemp(i, 1) = 0                 '上モールド温度
    atemp(i, 1) = T_keisu_cread(atemp(i, 1), T_keisu(T_keisuCont(1) - 1))
    atemp(i, 2) = 0                 '下モールド温度
    atemp(i, 2) = T_keisu_cread(atemp(i, 2), T_keisu(T_keisuCont(1) - 1))
  
'* 成形軸位置の取り込み */
      ppos = "TC 1 sec Disp 2"
      aposi(i) = r_z()
      '
'/* 型圧力の取り込み */
      ppos = "TC 1 sec Disp 3"
      apre(i) = r_pres()

'      If i = 1 Then GoTo jo0:
'      ix0 = Int(8.3333 / ptime * (i - 1)) + 120
'      ix = Int(8.3333 / ptime * (i)) + 120
'-------------- ピラニ計読み
'      LS21S_Monitor
'/* 温度分布の表示 */
'/* 型締圧のプロット */
'/* 座標値のプロット */
    lGphNo = i
    GphDataSet lGphNo0, lGphNo
    MoniGraph Me.Picture1, lGphNo0, lGphNo
    lGphNo0 = lGphNo
jo0:
'/* 各種データの画面下表示 １　*/
    DoEvents     '  2006.5.18  追加
    sdt1 = Right(Space(10) & Format(atemp(i, 0), "0.0"), 10) & "℃"
    sdt1 = sdt1 & Right(Space(10) & Format(apre(i), "0.00"), 10) & "kgf"
    sdt1 = sdt1 & Right(Space(10) & Format(aposi(i), "0.000"), 10) & "mm"
    Label2(14).Caption = sdt1
'/* 各種データの画面下表示 ２ */
    it = Timer                                                          ' 10/5
    it = diffTime(it, stTime)
    sdt2 = Right(Space(2) & Format(Int(it / 60), "0"), 2) & "分"
    sdt2 = sdt2 + Right(Space(2) & Format(Int(it) Mod 60, "0"), 3) & "秒　"       '2002.10.09 KYOCERA
    sdt2 = sdt2 + "ct" & Right(Space(7) & Format(diffTime(fintime, evtime), "0.0"), 7) & "  "
    sdt2 = sdt2 + "st" & Right(Space(7) & Format(diffTime(fintime, sevTime), "0.0"), 7) & "  "
    sdt2 = sdt2 + "tt" & Right(Space(7) & Format(diffTime(fintime, stTime), "0.0"), 7)
    Label2(11).Caption = sdt2
'/* 時刻表示 */
    Label10.Caption = Time$
'
'/* ﾛﾎﾞｯﾄ位置変更　*/
    'If FrmMenuFlg = False Then GoTo eend:
  Next i   '--------------------------------- For Loop
  js = js - 1
  GoTo ejs1:      '/* 表示終了で元画面へ */
'/* タクトタイムの算出　*/

send:
      ppos = "TC 1回end"
 
 '   stime = i
'    endTime = Timer
'    stime = diffTime(endTime, stTime)         '  10/5
'    sdt = Format(Int(stime / 60), "0") & "分" & Format(Int(stime) Mod 60, "0") & "秒"   '2002.10.09 KYOCERA
'    lCycleTime = sdt
'    Label2(6).Caption = Format(stime, "000") & Format(i_s, " 000")         '02.10.26 s.f. 削除
'/* データの保存　*/
    If lDtSaveFlg = True Then
      ResDtSave i_s, stime
      lDtSaveFlg = False
    End If
'　/*　成形データのセーブ　*/  2002.12.4 sf
'      Rec_of_Mold = Format(InitDat(11), "000") & "  "
      Rec_of_Mold = "   " & Format(z(iz3), "000.00") & "  "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(ct_temp(0)), "000") & "℃ " & Format(Int(ct_temp(1)), "000") & "℃ "
      Rec_of_Mold = Rec_of_Mold & "     " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00") & " "
      diTime1 = diffTime(cc_time(3), cc_time(2))
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(diTime1 + 0.5), "000") & "s "
      Rec_of_Mold = Rec_of_Mold & "    " & Format(cp_z, "000.000")
      Rec_of_Mold = Rec_of_Mold & "    " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "    " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & "    " & Format(Z3_Hosei(T_keisuCont(1) - 1), "0.000") & "  "

'    Rec_of_Mold = Format(InitDat(11), "000") & "  "　　' TC_MAIN で実施
'
'    Rec_of_Mold = " z " & Format(z(iz3), "000.00") & "  " & Format(z(5), "000.00") & " "
'    Rec_of_Mold = Rec_of_Mold & " :  ct " & Format(Int(ct_temp(0)), "000") & "℃ " & Format(Int(ct_temp(1)), "000") & "℃ "
'    Rec_of_Mold = Rec_of_Mold & " :  cc " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00") & " "
'    Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00") & " "
'    Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00") & " "
'    diTime1 = diffTime(cc_time(3), cc_time(2))
'    Rec_of_Mold = Rec_of_Mold & " :  " & Format(Int(diTime1 + 0.5), "0") & "s "
'    Rec_of_Mold = Rec_of_Mold & " : cp   " & Format(cp_z, "000.000")
'    Rec_of_Mold = Rec_of_Mold & " : t    " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00") & " "
'
'    RecDtSave Rec_of_Mold    ' TC_MAIN で実施
' /* 温度係数、肉厚補正データのカウントアップ
      Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  文字色を元に戻す
      Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  文字色を元に戻す
      Label4(T_keisuCont(1) - 1).BorderStyle = 0  '  枠なしに戻す
      Label6(T_keisuCont(1) - 1).BorderStyle = 0  '  枠なしに戻す
'     *** Z3の値を　戻す
          z(iz3) = z(iz3) - Z3_Hosei(T_keisuCont(1) - 1) '  ”Z3"の補正値reset
'     *** ポインターカウントアップ
      T_keisuCont(1) = T_keisuCont(1) + 1       ' ポインターのカウントアップ 無条件
'      Z3_HoseiCont(1) = Z3_HoseiCont(1) + 1       ' ポインターのカウントアップ
    If T_keisuCont(1) > (T_keisuCont(0)) Then T_keisuCont(1) = 1
'
    T_keisuCont(2) = T_keisuCont(1)       ' ポインターのbuckup
    T_keisuCont(3) = T_keisuCont(0)       '  型個数 backup
'
'    If Z3_HoseiCont(1) > (Z3_HoseiCont(0)) Then Z3_HoseiCont(1) = 1
'
'/* Ｖ入力がされていたら　エディット */
    If FrmMenuFlg = False Then Exit Do            'メニューから抜けるときfalse
    If EditFlg% = True Then 'エディタ起動
       ied = 1
       Exit Do
    End If
'/* 自動停止状態であれば停止 */
'    sts1 = SystemReadyChk()   'システムレディ or 非常停止
'    sts2 = AutoChk()          '自動状態？
'    If sts1 = 0 Or sts2 = 0 Then
      Label2(4).Caption = "自動運転終了"
'      FrmMenuFlg = False              'メニューから抜けるときfalse
'      NextView = 1
      Exit Do
'    End If
  Loop    '------------------------------------ DO LOOP
'/* Ｖ入力がされていたら　エディット */
    If ied = 1 Then 'エディタ起動
       MYEdit.Show 1
       'c = 0
       ied = 0
       GoTo st:     '/* エディットモードであれば　ｓｔにジャンプ */
    End If
'/* エディットモードであれば　ｓｔにジャンプ */
'    If ied <> 0 Then GoTo st:
'/* 予備加熱をゼロにし、ＯＦＦする */
eend:
    If iFlg_hijyou > 0 Then
         RecEmgDtSave sdt3, sdt1, sdt2 & gemgmsg
   End If                 '非常停止メッセージの保存  2004.3.8
  HeatOFF
  CoolOFF
'  ServoOFF
'/* ロボットデータのフロッピーへの書き出し */
'/* グラフィック画面の消去 */

'/* Ｖ入力がされていたら　エディット */
Exit Sub
'
errHandler:
  HeatOFF
  ServoOFF
  CoolOFF
'
  RecEmgDtSave sdt3, sdt1, sdt2
  If Err.Number <> 0 Then
     sdt1 = "エラー番号：" & Err.Number
     sdt2 = "ﾌﾟﾛｼﾞｪｸﾄ名：" & Err.Source & "  " & ppos
     sdt3 = "エラー内容：" & Err.Description
  End If
  RecEmgDtSave sdt1, sdt2, sdt3
  gemgmsg = Err.Number & Err.Description
  hijyou        '非常停止処理
'
End Sub
Private Sub genten()
'--------------
  C870Genten
  gOrgFlg = True                       '原点復帰完了=TRUE
  OrgON
  gOrgStartFlg = True   '2002.10.17 KYOCERA
End Sub

'Private Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* 設定圧力 */
'         m_p,      /* 設定Ｐ値 */
'         m_lim;    /* 設定リミット値 */
'Dim i%, nout%, ch%, v!
'Dim pa!, per!       '/* float（単精度浮動小数点型)*/
'  pa = r_pres()     '/* 圧力 */
'
''  If pa > 1000# Then '/* 1000Ｋｇ以上で非常停止 */
'  If pa > m_sa + 200# Then '/* 指定圧力＋200Ｋｇ以上で非常停止 */
'  hijyou
'    Exit Sub
'  End If
'
''/* ＰＩＤ演算 */
'
'  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
'  If per > m_lim Then per = m_lim
'  If per < (-1 * m_lim) Then per = -1 * m_lim
'  'nout = Int(40.95 * per) + &H800
'  nout = &H800 - Int(4.095 * per / 4#)
'  'nout = &H800 - Int(40.95 * per)
'  ch = 1
'  v = per / 5
'  'v = per / 5
'  DaOut ch, Hex(nout)
'  'DaVoltOut ch, V
'  'outp(ADPORT,(nout%256));
'  'outp(ADPORT+1,0x20|(nout/256));
'
'End Sub

Private Sub GphXSet()
Dim i%
  For i = 0 To ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub

Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)    '温調計
    Templu(i) = atemp(i, 1)   '上型
    Templd(i) = atemp(i, 2)   '下型
    Press(i) = apre(i)
    ZAxis(i) = aposi(i)
  Next i
End Sub

Private Function DispSegm$(ist0%)
Dim sdt$
    If ist0 < 0 Then Exit Function
    sdt = Right(Space(2) & Format(ist0, "0"), 2)
    sdt = sdt & Right(Space(4) & Format(seg_num(ist0), "0"), 4)
    sdt = sdt & Right(Space(4) & Format(ic(ist0), "0"), 4)
    sdt = sdt & Right(Space(12) & Format(z(ist0), "0.000"), 12)
    sdt = sdt & Right(Space(7) & Format(vel(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(6) & Format(pres(ist0), "0"), 6)
    sdt = sdt & Right(Space(7) & Format(t0(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(7) & Format(p(ist0), "0.0"), 7)
    DispSegm = sdt
'    Label2(12).Caption = sdt
End Function
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


Private Sub TC_Main()
Dim i%, flg%
Dim TCstTime!, TCendTime!, TCstime!
'------------ 成形オンリー
  SeikeiON         '成形ON　連続又は１回成形中
  TCFlg = True          'テスト成形中
  lEmgFlg = SystemReadyChk()  '非常停止
  lCycleTime = "------"         'サイクルタイム
  InitDat(11) = InitDat(11) - 1  '成形カウンタトウタル合わせ
  TCstTime = Timer
  TCi_s = -1
'
'/* 成形データ保存ファイルの作成　*/
  RecDtSave0 InitDat(11)
'
'----------
  For i = 1 To iPltMax     'パレット回転数
    If lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
    Label2(4).Caption = "パレット" & Trim(Str(i)) & "順中"
    TCi_s = TCi_s + 1
    Label2(9).Caption = TCi_s   '成形カウンタ
'
   Plt1Jyun
'
    If lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
    If i <> iPltMax Then
'
'  ------ 成形カウンタ管理 -------
          InitDat(11) = InitDat(11) + 1  '成形カウンタトウタル
          InitDtSave
          Label2(13).Caption = Str(InitDat(11))   '成形カウンタトウタル
' /* ---  成形メイン ---
      LS21T_MAIN
'
'    ---  ﾀｸﾄﾀｲﾑ　　計算
        TCendTime = Timer
        TCstime = diffTime(TCendTime, TCstTime)
        lCycleTime$ = Format(Int(TCstime / 60), "0") & "分" & Format(Int(TCstime) Mod 60, "0") & "秒"
        Label2(8).Caption = lCycleTime$           'サイクルタイム
        TCstTime = Timer
'
      Rec_of_Mold = Format(i, "000") & "  " & Rec_of_Mold  '成形データのsave
      RecDtSave Rec_of_Mold
'
      If iFlg_hijyou = 1 Or lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
'
    End If
  Next i
  TCFlg = False         'テスト成形終了
  SeikeiOFF          '成形OFF　待機中
  If lEmgFlg <> 0 Then
    If FrmMenuFlg = False Then
      Label2(4).Caption = "中断"
      FrmMenuFlg = True
    Else
      coolingform.Show
'
      WaitSec (1)
      flg = MsgBox("成形・冷却　終了 " + Time$ + "   ", 48, "1回成形") '終了メッセージ
    End If
  Else
    RecEmgDtSave sdt3, sdt1, sdt2         '非常停止メッセージの保存  2004.3.8
'
    Unload Me
    PGM_Menu.Show
  End If
End Sub

Private Sub Plt1Jyun()
Dim sts%
'---------- パレット1順指令→1順完了まで待つ
  WaitSec 1.5
  PCTrnsReq     ' パレット1順指令
  
  '2002.10.9 KYOCERA
  sts = 0
  Do
    sts = PCTrnsChk()   'BUSYチェック
    lEmgFlg = SystemReadyChk()  '非常停止
    If sts = 1 Or lEmgFlg = 0 Then Exit Do
    DoEvents
  Loop
  
  sts = 0
  Do
    sts = PCTrnsChk()   'PCから搬送中=1
    lEmgFlg = SystemReadyChk()  '非常停止
    If sts = 0 Or lEmgFlg = 0 Then Exit Do
    DoEvents
  Loop
End Sub

Private Sub Timer2_Timer()
If r_z > 0.1 Then
        OrgOFF
    Else
        OrgON
    End If
End Sub
