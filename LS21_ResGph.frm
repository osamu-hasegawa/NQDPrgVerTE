VERSION 5.00
Begin VB.Form LS21_ResGph 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ê¨å`âÊñ "
   ClientHeight    =   7896
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   11856
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7896
   ScaleWidth      =   11856
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command2 
      Caption         =   "åãâ ÉfÅ[É^ì«çûÇ›"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1920
      TabIndex        =   44
      Top             =   120
      Width           =   2172
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ã◊Øƒ
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5500
      Left            =   1800
      ScaleHeight     =   5472
      ScaleWidth      =   8376
      TabIndex        =   6
      Top             =   1860
      Width           =   8400
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   7
         X1              =   6696
         X2              =   6696
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   6
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   5
         X1              =   3348
         X2              =   3348
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   4
         X1              =   1656
         X2              =   1656
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   3
         X1              =   0
         X2              =   8352
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   2
         X1              =   0
         X2              =   8352
         Y1              =   2196
         Y2              =   2196
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   1
         X1              =   0
         X2              =   8352
         Y1              =   3312
         Y2              =   3312
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'ì_ê¸
         Index           =   0
         X1              =   0
         X2              =   8352
         Y1              =   4392
         Y2              =   4392
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉÅÉjÉÖÅ["
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ë™íËéûçèÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Left            =   4200
      TabIndex        =   45
      Top             =   360
      Width           =   1272
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      TabIndex        =   43
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ÉVÉáÉbÉgêîÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   42
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ÉTÉCÉNÉãÉ^ÉCÉÄÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   41
      Top             =   390
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      TabIndex        =   40
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      TabIndex        =   39
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(ï™)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   38
      Top             =   7560
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "åoâﬂéûä‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   37
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
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   36
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   35
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   34
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   33
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   32
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   31
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å^â∑ìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   30
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(Åé)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   29
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   28
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   27
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   26
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   25
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   24
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   23
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å^í˜à≥"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   22
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(t)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   21
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
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   20
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   19
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   18
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   17
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   16
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   15
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ç¿ïW"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   14
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(mm)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   13
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
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   12
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   11
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   10
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   9
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   8
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      BackStyle       =   0  'ìßñæ
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.8
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
      TabIndex        =   7
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Top             =   696
      Width           =   6252
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "êßå‰ÉtÉ@ÉCÉãñºÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Left            =   2304
      TabIndex        =   3
      Top             =   684
      Width           =   2028
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      BackStyle       =   0  'ìßñæ
      Caption         =   "ë™íËì˙ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   108
      Width           =   1020
   End
End
Attribute VB_Name = "LS21_ResGph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lGphNo%
Dim lGphNo0%

Private Sub Command2_Click(Index As Integer)
  Select Case Index
  Case 0  'ÉLÉÉÉìÉZÉã
    lGphNo = 0
    MoniGraph Me.Picture1, 0, lGphNo
  Case 1  'ÉÅÉjÉÖÅ[
    Unload Me
    PGM_Menu.Show
  Case 2  'ÉOÉâÉtçƒï`âÊ
    Picture1.Cls
    ResFlLoad
    SetData
    lGphNo = gGphDtNum
    GphDataSet 0, lGphNo
    MoniGraph Me.Picture1, 0, lGphNo
  
  End Select
End Sub

Private Sub SetData()
  Label2(0) = Format(gDate, "###0")  'ë™íË
  Label2(1) = Format(gTime, "###0")  '
  Label2(2) = gcoxFlName             'êßå‰ÉtÉ@ÉCÉãñº
 '-----------------------------------
  DispGphScale
End Sub

Private Sub GetData()

End Sub

Private Sub Form_Load()
  DispCenter Me
  SetData
End Sub


Private Sub DispGphScale()
Dim i%, p%, max!, min!, def!, dev%
  '
  GphXSet           'éûä‘é≤ÇÃéûä‘ÇÉZÉbÉg
  '
  dev = 5
  '
  min = InitDat(1)  'ÉOÉâÉtÉXÉPÅ[Éãç¿ïW (Min)
  max = InitDat(2)  'ÉOÉâÉtÉXÉPÅ[Éãç¿ïW (Max)
  def = (max - min) / dev
  p = 0
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(3)  'ÉOÉâÉtÉXÉPÅ[Éãå^í˜à≥ (Min)
  max = InitDat(4)  'ÉOÉâÉtÉXÉPÅ[Éãå^í˜à≥ (Max)
  def = (max - min) / dev
  p = 8
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(5)  'ÉOÉâÉtÉXÉPÅ[Éãå^â∑ìx (Min)
  max = InitDat(6)  'ÉOÉâÉtÉXÉPÅ[Éãå^â∑ìx (Max)
  def = (max - min) / dev
  p = 16
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(7)  'ÉOÉâÉtÉXÉPÅ[Éãåoâﬂéûä‘ (Min)
  max = InitDat(8)  'ÉOÉâÉtÉXÉPÅ[Éãåoâﬂéûä‘ (Max)
  def = (max - min) / dev
  p = 24
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
'
End Sub


Private Sub GphXSet()
Dim i%
  For i = 0 To gGphDtNum    'ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub
Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)
    Templu(i) = atemp(i, 1)   'è„å^â∑ìx
    Templd(i) = atemp(i, 2)   'â∫å^â∑ìx
    Press(i) = apre(i)
    ZAxis(i) = aposi(i)
  Next i
End Sub
