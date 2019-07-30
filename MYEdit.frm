VERSION 5.00
Begin VB.Form MYEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "制御データエディタ画面"
   ClientHeight    =   7404
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   10452
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   7.8
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7404
   ScaleWidth      =   10452
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command2 
      Caption         =   "Comment"
      Height          =   300
      Index           =   4
      Left            =   8400
      TabIndex        =   279
      Top             =   1100
      Width           =   900
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   5160
      TabIndex        =   278
      Text            =   "Text16"
      Top             =   5550
      Width           =   612
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   252
      Index           =   9
      Left            =   5040
      TabIndex        =   277
      Text            =   "Text14"
      Top             =   4080
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   276
      Index           =   7
      Left            =   6900
      TabIndex        =   276
      Top             =   600
      Width           =   264
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   7
      Left            =   6200
      TabIndex        =   275
      Text            =   "Text1"
      Top             =   600
      Width           =   700
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   276
      Index           =   6
      Left            =   6900
      TabIndex        =   274
      Top             =   930
      Width           =   264
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   6
      Left            =   6200
      TabIndex        =   273
      Text            =   "Text1"
      Top             =   960
      Width           =   700
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   3
      Left            =   7320
      TabIndex        =   270
      Text            =   "Text10"
      Top             =   600
      Width           =   612
   End
   Begin VB.VScrollBar VScroll10 
      Height          =   240
      Index           =   3
      Left            =   7920
      TabIndex        =   269
      Top             =   600
      Width           =   252
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   7320
      TabIndex        =   268
      Text            =   "Text10"
      Top             =   960
      Width           =   612
   End
   Begin VB.VScrollBar VScroll10 
      Height          =   240
      Index           =   2
      Left            =   7920
      TabIndex        =   267
      Top             =   960
      Width           =   252
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   6
      Left            =   10200
      TabIndex        =   259
      Top             =   6980
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9550
      TabIndex        =   258
      Text            =   "Text15"
      Top             =   6980
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   5
      Left            =   10200
      TabIndex        =   257
      Top             =   6588
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9550
      TabIndex        =   256
      Text            =   "Text15"
      Top             =   6588
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   255
      Top             =   6192
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9550
      TabIndex        =   254
      Text            =   "Text15"
      Top             =   6192
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   253
      Top             =   5796
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9550
      TabIndex        =   252
      Text            =   "Text15"
      Top             =   5796
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   251
      Top             =   5400
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9550
      TabIndex        =   250
      Text            =   "Text15"
      Top             =   5400
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   249
      Top             =   5004
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9550
      TabIndex        =   248
      Text            =   "Text15"
      Top             =   5004
      Width           =   650
   End
   Begin VB.VScrollBar VScroll15 
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   247
      Top             =   4608
      Width           =   200
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9550
      TabIndex        =   246
      Text            =   "Text15"
      Top             =   4608
      Width           =   650
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   10
      Left            =   5400
      TabIndex        =   245
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.VScrollBar VScroll14 
      Height          =   255
      Left            =   5760
      TabIndex        =   244
      Top             =   3840
      Width           =   135
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   8
      Left            =   4840
      TabIndex        =   242
      Text            =   "Text14"
      Top             =   3840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   7
      Left            =   4400
      TabIndex        =   241
      Text            =   "Text14"
      Top             =   3840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   240
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   5
      Left            =   3480
      TabIndex        =   239
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   4
      Left            =   2900
      TabIndex        =   238
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   3
      Left            =   2450
      TabIndex        =   237
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   2
      Left            =   1970
      TabIndex        =   236
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   1
      Left            =   1500
      TabIndex        =   235
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text14 
      Height          =   252
      Index           =   0
      Left            =   1050
      TabIndex        =   234
      Text            =   "Text14"
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   240
      TabIndex        =   233
      Text            =   "Text9"
      Top             =   5280
      Width           =   492
   End
   Begin VB.VScrollBar VScroll9 
      Height          =   252
      Index           =   2
      Left            =   720
      TabIndex        =   232
      Top             =   5280
      Width           =   132
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Notepad"
      Height          =   180
      Index           =   3
      Left            =   9600
      TabIndex        =   230
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6372
      TabIndex        =   228
      Text            =   "Text4"
      Top             =   6980
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   7415
      TabIndex        =   227
      Text            =   "Text5"
      Top             =   6980
      Width           =   700
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   8470
      TabIndex        =   226
      Text            =   "Text6"
      Top             =   6980
      Width           =   700
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   6
      Left            =   7080
      TabIndex        =   225
      Top             =   6980
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   6
      Left            =   8160
      TabIndex        =   224
      Top             =   6980
      Width           =   200
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   6
      Left            =   9200
      TabIndex        =   223
      Top             =   6980
      Width           =   200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "変更確認"
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
      Index           =   2
      Left            =   2760
      TabIndex        =   222
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "増分"
      Height          =   156
      Index           =   9
      Left            =   5160
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   221
      Top             =   4680
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Caption         =   "9"
      Height          =   156
      Index           =   8
      Left            =   4320
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   211
      Top             =   4680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "8"
      Height          =   156
      Index           =   7
      Left            =   3480
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   210
      Top             =   4680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "7"
      Height          =   156
      Index           =   6
      Left            =   2640
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   209
      Top             =   4680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "6"
      Height          =   156
      Index           =   5
      Left            =   1800
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   208
      Top             =   4680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "5"
      Height          =   156
      Index           =   4
      Left            =   5160
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   207
      Top             =   4280
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "4"
      Height          =   156
      Index           =   3
      Left            =   4320
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   206
      Top             =   4280
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   156
      Index           =   2
      Left            =   3480
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   205
      Top             =   4280
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   156
      Index           =   1
      Left            =   2640
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   204
      Top             =   4280
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      Height          =   156
      Index           =   0
      Left            =   1800
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   203
      Top             =   4280
      Width           =   300
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   240
      TabIndex        =   200
      Text            =   "Text13"
      Top             =   5700
      Width           =   300
   End
   Begin VB.VScrollBar VScroll13 
      Height          =   252
      Left            =   550
      TabIndex        =   199
      Top             =   5700
      Width           =   132
   End
   Begin VB.CommandButton command1 
      Caption         =   "加圧時間制御"
      Height          =   240
      Left            =   8400
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   198
      Top             =   72
      UseMaskColor    =   -1  'True
      Width           =   1932
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   4944
      TabIndex        =   196
      Text            =   "Text8"
      Top             =   6876
      Width           =   660
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   336
      Index           =   4
      Left            =   5616
      TabIndex        =   195
      Top             =   6876
      Width           =   220
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   4944
      TabIndex        =   193
      Text            =   "Text7"
      Top             =   6276
      Width           =   660
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   336
      Index           =   4
      Left            =   5616
      TabIndex        =   192
      Top             =   6288
      Width           =   220
   End
   Begin VB.VScrollBar VScroll10 
      Height          =   240
      Index           =   1
      Left            =   10080
      TabIndex        =   187
      Top             =   360
      Width           =   252
   End
   Begin VB.VScrollBar VScroll10 
      Height          =   240
      Index           =   0
      Left            =   10080
      TabIndex        =   186
      Top             =   720
      Width           =   252
   End
   Begin VB.VScrollBar VScroll9 
      Height          =   252
      Index           =   1
      Left            =   720
      TabIndex        =   185
      Top             =   4800
      Width           =   132
   End
   Begin VB.VScrollBar VScroll9 
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   184
      Top             =   4370
      Width           =   132
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   8
      Left            =   4920
      TabIndex        =   181
      Top             =   5520
      Width           =   132
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   7
      Left            =   4080
      TabIndex        =   180
      Top             =   5520
      Width           =   132
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   8
      Left            =   4320
      TabIndex        =   179
      Text            =   "Text12"
      Top             =   5520
      Width           =   612
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   7
      Left            =   3480
      TabIndex        =   178
      Text            =   "Text12"
      Top             =   5520
      Width           =   612
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   6
      Left            =   3240
      TabIndex        =   177
      Top             =   5520
      Width           =   132
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   5
      Left            =   2400
      TabIndex        =   176
      Top             =   5520
      Width           =   132
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   6
      Left            =   2640
      TabIndex        =   175
      Text            =   "Text12"
      Top             =   5520
      Width           =   612
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   5
      Left            =   1800
      TabIndex        =   174
      Text            =   "Text12"
      Top             =   5520
      Width           =   612
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   4
      Left            =   5760
      TabIndex        =   173
      Top             =   5160
      Width           =   132
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   4
      Left            =   5160
      TabIndex        =   172
      Text            =   "Text12"
      Top             =   5160
      Width           =   612
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   3
      Left            =   4920
      TabIndex        =   171
      Top             =   5160
      Width           =   132
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   170
      Top             =   5160
      Width           =   132
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   3
      Left            =   4320
      TabIndex        =   169
      Text            =   "Text12"
      Top             =   5160
      Width           =   612
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   3480
      TabIndex        =   168
      Text            =   "Text12"
      Top             =   5160
      Width           =   612
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   167
      Top             =   5160
      Width           =   132
   End
   Begin VB.VScrollBar VScroll12 
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   166
      Top             =   5160
      Width           =   132
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   2640
      TabIndex        =   165
      Text            =   "Text12"
      Top             =   5160
      Width           =   612
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   1800
      TabIndex        =   164
      Text            =   "Text12"
      Top             =   5160
      Width           =   612
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   9480
      TabIndex        =   163
      Text            =   "Text10"
      Top             =   360
      Width           =   612
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   9480
      TabIndex        =   162
      Text            =   "Text10"
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   240
      TabIndex        =   161
      Text            =   "Text9"
      Top             =   4800
      Width           =   492
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   240
      TabIndex        =   160
      Text            =   "Text9"
      Top             =   4370
      Width           =   492
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   9
      Left            =   5760
      TabIndex        =   159
      Top             =   4800
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   9
      Left            =   5160
      TabIndex        =   158
      Text            =   "Text11"
      Top             =   4800
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   8
      Left            =   4920
      TabIndex        =   157
      Top             =   4800
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   8
      Left            =   4320
      TabIndex        =   156
      Text            =   "Text11"
      Top             =   4800
      Width           =   732
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   7
      Left            =   4080
      TabIndex        =   155
      Top             =   4800
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   7
      Left            =   3480
      TabIndex        =   154
      Text            =   "Text11"
      Top             =   4800
      Width           =   732
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   6
      Left            =   3240
      TabIndex        =   153
      Top             =   4800
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   6
      Left            =   2640
      TabIndex        =   152
      Text            =   "Text11"
      Top             =   4800
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   5
      Left            =   2400
      TabIndex        =   151
      Top             =   4800
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   5
      Left            =   1800
      TabIndex        =   150
      Text            =   "Text11"
      Top             =   4800
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   4
      Left            =   5760
      TabIndex        =   149
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   4
      Left            =   5160
      TabIndex        =   148
      Text            =   "Text11"
      Top             =   4440
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   3
      Left            =   4920
      TabIndex        =   147
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   3
      Left            =   4320
      TabIndex        =   146
      Text            =   "Text11"
      Top             =   4440
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   145
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   3480
      TabIndex        =   144
      Text            =   "Text11"
      Top             =   4440
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   143
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   2640
      TabIndex        =   142
      Text            =   "Text11"
      Top             =   4440
      Width           =   612
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   141
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   1800
      TabIndex        =   140
      Text            =   "Text11"
      Top             =   4440
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
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
      Height          =   2376
      Left            =   360
      ScaleHeight     =   2352
      ScaleWidth      =   5616
      TabIndex        =   139
      Top             =   1440
      Width           =   5640
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   25
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   24
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   23
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   22
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   21
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   20
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   19
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   18
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   17
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   16
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   15
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   14
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   13
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   12
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   11
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   10
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   9
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   8
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   7
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   6
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   5
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   4
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3996
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '点線
         Index           =   0
         X1              =   100
         X2              =   100
         Y1              =   0
         Y2              =   3996
      End
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   336
      Index           =   3
      Left            =   4476
      TabIndex        =   138
      Top             =   6876
      Width           =   220
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   336
      Index           =   3
      Left            =   4476
      TabIndex        =   137
      Top             =   6288
      Width           =   220
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3804
      TabIndex        =   136
      Text            =   "Text8"
      Top             =   6876
      Width           =   660
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3804
      TabIndex        =   135
      Text            =   "Text7"
      Top             =   6276
      Width           =   660
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   336
      Index           =   2
      Left            =   3264
      TabIndex        =   132
      Top             =   6876
      Width           =   220
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   336
      Index           =   2
      Left            =   3264
      TabIndex        =   131
      Top             =   6288
      Width           =   220
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2604
      TabIndex        =   130
      Text            =   "Text8"
      Top             =   6876
      Width           =   660
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2604
      TabIndex        =   129
      Text            =   "Text7"
      Top             =   6276
      Width           =   660
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   336
      Index           =   1
      Left            =   2160
      TabIndex        =   126
      Top             =   6876
      Width           =   220
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   336
      Index           =   1
      Left            =   2160
      TabIndex        =   125
      Top             =   6288
      Width           =   220
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1476
      TabIndex        =   124
      Text            =   "Text8"
      Top             =   6876
      Width           =   660
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1476
      TabIndex        =   123
      Text            =   "Text7"
      Top             =   6276
      Width           =   660
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   5
      Left            =   9200
      TabIndex        =   120
      Top             =   6588
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   5
      Left            =   8160
      TabIndex        =   119
      Top             =   6588
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   5
      Left            =   7080
      TabIndex        =   118
      Top             =   6588
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   8470
      TabIndex        =   117
      Text            =   "Text6"
      Top             =   6588
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7415
      TabIndex        =   116
      Text            =   "Text5"
      Top             =   6588
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6372
      TabIndex        =   115
      Text            =   "Text4"
      Top             =   6588
      Width           =   700
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   4
      Left            =   9200
      TabIndex        =   113
      Top             =   6192
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   4
      Left            =   8160
      TabIndex        =   112
      Top             =   6192
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   4
      Left            =   7080
      TabIndex        =   111
      Top             =   6192
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   8470
      TabIndex        =   110
      Text            =   "Text6"
      Top             =   6192
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7415
      TabIndex        =   109
      Text            =   "Text5"
      Top             =   6192
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6372
      TabIndex        =   108
      Text            =   "Text4"
      Top             =   6192
      Width           =   700
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   3
      Left            =   9200
      TabIndex        =   106
      Top             =   5796
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   3
      Left            =   8160
      TabIndex        =   105
      Top             =   5796
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   3
      Left            =   7080
      TabIndex        =   104
      Top             =   5796
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   8470
      TabIndex        =   103
      Text            =   "Text6"
      Top             =   5796
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7415
      TabIndex        =   102
      Text            =   "Text5"
      Top             =   5796
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6372
      TabIndex        =   101
      Text            =   "Text4"
      Top             =   5796
      Width           =   700
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   2
      Left            =   9200
      TabIndex        =   99
      Top             =   5400
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   2
      Left            =   8160
      TabIndex        =   98
      Top             =   5400
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   2
      Left            =   7080
      TabIndex        =   97
      Top             =   5400
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   8470
      TabIndex        =   96
      Text            =   "Text6"
      Top             =   5400
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7415
      TabIndex        =   95
      Text            =   "Text5"
      Top             =   5400
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6372
      TabIndex        =   94
      Text            =   "Text4"
      Top             =   5400
      Width           =   700
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   1
      Left            =   9200
      TabIndex        =   92
      Top             =   5004
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   1
      Left            =   8160
      TabIndex        =   91
      Top             =   5004
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   1
      Left            =   7080
      TabIndex        =   90
      Top             =   5004
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   8470
      TabIndex        =   89
      Text            =   "Text6"
      Top             =   5004
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7415
      TabIndex        =   88
      Text            =   "Text5"
      Top             =   5004
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6372
      TabIndex        =   87
      Text            =   "Text4"
      Top             =   5004
      Width           =   700
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   5
      Left            =   10152
      TabIndex        =   85
      Top             =   3744
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   5
      Left            =   8676
      TabIndex        =   84
      Top             =   3744
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   5
      Left            =   7200
      TabIndex        =   83
      Top             =   3744
      Width           =   264
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9324
      TabIndex        =   82
      Text            =   "Text3"
      Top             =   3744
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7848
      TabIndex        =   80
      Text            =   "Text2"
      Top             =   3744
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6372
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   3744
      Width           =   804
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   4
      Left            =   10152
      TabIndex        =   76
      Top             =   3348
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   4
      Left            =   8676
      TabIndex        =   75
      Top             =   3348
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   4
      Left            =   7200
      TabIndex        =   74
      Top             =   3348
      Width           =   264
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9324
      TabIndex        =   73
      Text            =   "Text3"
      Top             =   3348
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7848
      TabIndex        =   71
      Text            =   "Text2"
      Top             =   3348
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6372
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   3348
      Width           =   804
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   3
      Left            =   10152
      TabIndex        =   67
      Top             =   2952
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   3
      Left            =   8676
      TabIndex        =   66
      Top             =   2952
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   3
      Left            =   7200
      TabIndex        =   65
      Top             =   2952
      Width           =   264
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9324
      TabIndex        =   64
      Text            =   "Text3"
      Top             =   2952
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7848
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   2952
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6372
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   2952
      Width           =   804
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   2
      Left            =   10152
      TabIndex        =   58
      Top             =   2556
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   2
      Left            =   8676
      TabIndex        =   57
      Top             =   2556
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   2
      Left            =   7200
      TabIndex        =   56
      Top             =   2556
      Width           =   264
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9324
      TabIndex        =   55
      Text            =   "Text3"
      Top             =   2556
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7848
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   2556
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6372
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2556
      Width           =   804
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   1
      Left            =   10152
      TabIndex        =   49
      Top             =   2160
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   1
      Left            =   8676
      TabIndex        =   48
      Top             =   2160
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   1
      Left            =   7200
      TabIndex        =   47
      Top             =   2160
      Width           =   264
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   336
      Index           =   0
      Left            =   1020
      TabIndex        =   46
      Top             =   6876
      Width           =   220
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   336
      Index           =   0
      Left            =   1020
      TabIndex        =   45
      Top             =   6288
      Width           =   220
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   336
      Index           =   0
      Left            =   9200
      TabIndex        =   44
      Top             =   4608
      Width           =   200
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   336
      Index           =   0
      Left            =   8160
      TabIndex        =   43
      Top             =   4608
      Width           =   200
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   336
      Index           =   0
      Left            =   7080
      TabIndex        =   42
      Top             =   4608
      Width           =   200
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   336
      Index           =   0
      Left            =   10152
      TabIndex        =   41
      Top             =   1764
      Width           =   264
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   336
      Index           =   0
      Left            =   8676
      TabIndex        =   40
      Top             =   1764
      Width           =   264
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   336
      Index           =   0
      Left            =   7200
      TabIndex        =   39
      Top             =   1750
      Width           =   264
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9324
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   2160
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7848
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   2160
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6372
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2160
      Width           =   804
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   330
      TabIndex        =   32
      Text            =   "Text8"
      Top             =   6876
      Width           =   660
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   330
      TabIndex        =   31
      Text            =   "Text7"
      Top             =   6276
      Width           =   660
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   8470
      TabIndex        =   30
      Text            =   "Text6"
      Top             =   4608
      Width           =   700
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7415
      TabIndex        =   28
      Text            =   "Text5"
      Top             =   4608
      Width           =   700
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6372
      TabIndex        =   26
      Text            =   "Text4"
      Top             =   4608
      Width           =   700
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9324
      TabIndex        =   24
      Text            =   "Text3"
      Top             =   1764
      Width           =   804
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7848
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   1764
      Width           =   804
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6372
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1780
      Width           =   804
   End
   Begin VB.CommandButton Command2 
      Caption         =   "成形へ戻る"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   252
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ｷｬﾝｾﾙ or　Notepad戻り"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   100
      TabIndex        =   10
      Top             =   90
      Width           =   1150
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ｃp位置"
      Height          =   180
      Index           =   2
      Left            =   6250
      TabIndex        =   272
      Top             =   360
      Width           =   804
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0E0FF&
      Caption         =   "加圧時間"
      Height          =   180
      Index           =   1
      Left            =   7350
      TabIndex        =   271
      Top             =   360
      Width           =   804
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0E0FF&
      Caption         =   "ア　ラ　ー　ム　設　定"
      Height          =   180
      Index           =   0
      Left            =   6240
      TabIndex        =   266
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   " 周目"
      ForeColor       =   &H00C00000&
      Height          =   144
      Index           =   3
      Left            =   240
      TabIndex        =   265
      Top             =   5100
      Width           =   480
   End
   Begin VB.Label Label10 
      Caption         =   "下型"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   8880
      TabIndex        =   264
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "上型"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   1
      Left            =   7920
      TabIndex        =   263
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "IH"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   6960
      TabIndex        =   262
      Top             =   4300
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "温度"
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
      Index           =   49
      Left            =   8400
      TabIndex        =   261
      Top             =   4320
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "温度"
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
      Index           =   48
      Left            =   7440
      TabIndex        =   260
      Top             =   4320
      Width           =   510
   End
   Begin VB.Label Label9 
      Alignment       =   2  '中央揃え
      Caption         =   "型No."
      Height          =   255
      Left            =   360
      TabIndex        =   243
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "Z No."
      Height          =   132
      Index           =   9
      Left            =   5160
      TabIndex        =   231
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "LA"
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
      Index           =   47
      Left            =   6084
      TabIndex        =   229
      Top             =   7080
      Width           =   264
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "9"
      Height          =   132
      Index           =   8
      Left            =   4320
      TabIndex        =   220
      Top             =   5400
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "8"
      Height          =   132
      Index           =   7
      Left            =   3480
      TabIndex        =   219
      Top             =   5400
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "7"
      Height          =   132
      Index           =   6
      Left            =   2640
      TabIndex        =   218
      Top             =   5400
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "6"
      Height          =   132
      Index           =   5
      Left            =   1800
      TabIndex        =   217
      Top             =   5400
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "5"
      Height          =   132
      Index           =   4
      Left            =   5160
      TabIndex        =   216
      Top             =   5040
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "4"
      Height          =   132
      Index           =   3
      Left            =   4320
      TabIndex        =   215
      Top             =   5040
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "3"
      Height          =   132
      Index           =   2
      Left            =   3480
      TabIndex        =   214
      Top             =   5040
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "2"
      Height          =   132
      Index           =   1
      Left            =   2640
      TabIndex        =   213
      Top             =   5040
      Width           =   132
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "1"
      Height          =   132
      Index           =   0
      Left            =   1800
      TabIndex        =   212
      Top             =   5040
      Width           =   132
   End
   Begin VB.Label Label8 
      Alignment       =   2  '中央揃え
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   202
      Top             =   5700
      Width           =   732
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "変更内容"
      ForeColor       =   &H00C00000&
      Height          =   144
      Index           =   2
      Left            =   240
      TabIndex        =   201
      Top             =   5550
      Width           =   640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J5"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   46
      Left            =   4956
      TabIndex        =   197
      Top             =   6684
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "C5"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   45
      Left            =   5000
      TabIndex        =   194
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "加圧時間max"
      ForeColor       =   &H00C00000&
      Height          =   156
      Index           =   1
      Left            =   8436
      TabIndex        =   191
      Top             =   430
      Width           =   936
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "個数"
      ForeColor       =   &H00C00000&
      Height          =   156
      Index           =   0
      Left            =   276
      TabIndex        =   190
      Top             =   4200
      Width           =   336
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "加圧時間min"
      ForeColor       =   &H00C00000&
      Height          =   144
      Index           =   1
      Left            =   8436
      TabIndex        =   189
      Top             =   780
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "ﾎﾟｲﾝﾀｰ"
      ForeColor       =   &H00C00000&
      Height          =   144
      Index           =   0
      Left            =   240
      TabIndex        =   188
      Top             =   4650
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "肉厚補正"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   204
      Left            =   1080
      TabIndex        =   183
      Top             =   5160
      Width           =   684
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "温度係数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   204
      Index           =   0
      Left            =   1080
      TabIndex        =   182
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J4"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   44
      Left            =   3840
      TabIndex        =   134
      Top             =   6680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "C4"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   43
      Left            =   3820
      TabIndex        =   133
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J3"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   42
      Left            =   2640
      TabIndex        =   128
      Top             =   6680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "C3"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   41
      Left            =   2640
      TabIndex        =   127
      Top             =   6060
      Width           =   6060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   40
      Left            =   1440
      TabIndex        =   122
      Top             =   6680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "C2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   39
      Left            =   1560
      TabIndex        =   121
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T6"
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
      Index           =   38
      Left            =   6084
      TabIndex        =   114
      Top             =   6696
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T5"
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
      Index           =   37
      Left            =   6084
      TabIndex        =   107
      Top             =   6300
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T4"
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
      Index           =   36
      Left            =   6084
      TabIndex        =   100
      Top             =   5904
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T3"
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
      Index           =   35
      Left            =   6084
      TabIndex        =   93
      Top             =   5508
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T2"
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
      Index           =   34
      Left            =   6084
      TabIndex        =   86
      Top             =   5112
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P6"
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
      Index           =   33
      Left            =   9036
      TabIndex        =   81
      Top             =   3816
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V6"
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
      Index           =   32
      Left            =   7560
      TabIndex        =   79
      Top             =   3816
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z6"
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
      Index           =   31
      Left            =   6084
      TabIndex        =   77
      Top             =   3816
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P5"
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
      Index           =   30
      Left            =   9036
      TabIndex        =   72
      Top             =   3420
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V5"
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
      Index           =   29
      Left            =   7560
      TabIndex        =   70
      Top             =   3420
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z5"
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
      Index           =   28
      Left            =   6084
      TabIndex        =   68
      Top             =   3420
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P4"
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
      Index           =   27
      Left            =   9036
      TabIndex        =   63
      Top             =   3024
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V4"
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
      Index           =   26
      Left            =   7560
      TabIndex        =   61
      Top             =   3024
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z4"
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
      Index           =   25
      Left            =   6084
      TabIndex        =   59
      Top             =   3024
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P3"
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
      Index           =   24
      Left            =   9036
      TabIndex        =   54
      Top             =   2628
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V3"
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
      Index           =   23
      Left            =   7560
      TabIndex        =   52
      Top             =   2628
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z3"
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
      Index           =   22
      Left            =   6084
      TabIndex        =   50
      Top             =   2628
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P2"
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
      Index           =   21
      Left            =   9036
      TabIndex        =   37
      Top             =   2232
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V2"
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
      Index           =   20
      Left            =   7560
      TabIndex        =   35
      Top             =   2232
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z2"
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
      Index           =   19
      Left            =   6084
      TabIndex        =   33
      Top             =   2232
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J1"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   18
      Left            =   72
      TabIndex        =   29
      Top             =   6876
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "C1"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   17
      Left            =   72
      TabIndex        =   27
      Top             =   6276
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "T1"
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
      Index           =   16
      Left            =   6084
      TabIndex        =   25
      Top             =   4716
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P1"
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
      Index           =   15
      Left            =   9036
      TabIndex        =   23
      Top             =   1836
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "V1"
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
      Index           =   14
      Left            =   7560
      TabIndex        =   21
      Top             =   1836
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Z1"
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
      Index           =   13
      Left            =   6084
      TabIndex        =   19
      Top             =   1836
      Width           =   276
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "待ち時間"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   156
      Index           =   12
      Left            =   360
      TabIndex        =   18
      Top             =   6684
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "チェック温度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   156
      Index           =   11
      Left            =   240
      TabIndex        =   17
      Top             =   6050
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "昇温時間"
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
      Index           =   10
      Left            =   9360
      TabIndex        =   16
      Top             =   4320
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "圧力"
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
      Left            =   9288
      TabIndex        =   15
      Top             =   1476
      Width           =   516
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "速度"
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
      Left            =   7848
      TabIndex        =   14
      Top             =   1476
      Width           =   516
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
      Left            =   180
      TabIndex        =   13
      Top             =   1008
      Width           =   1272
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
      Left            =   1476
      TabIndex        =   12
      Top             =   1008
      Width           =   4608
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
      Left            =   2160
      TabIndex        =   9
      Top             =   684
      Width           =   3852
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
      Left            =   216
      TabIndex        =   8
      Top             =   684
      Width           =   2028
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   6
      Left            =   5850
      TabIndex        =   7
      Top             =   360
      Width           =   228
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   204
      Index           =   5
      Left            =   5850
      TabIndex        =   6
      Top             =   72
      Width           =   228
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
      Index           =   1
      Left            =   5100
      TabIndex        =   5
      Top             =   360
      Width           =   700
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
      Index           =   0
      Left            =   5100
      TabIndex        =   4
      Top             =   72
      Width           =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "予備加熱："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   4
      Left            =   4100
      TabIndex        =   3
      Top             =   396
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "測定時間："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   3
      Left            =   4100
      TabIndex        =   2
      Top             =   108
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "温度"
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
      Left            =   6408
      TabIndex        =   1
      Top             =   4320
      Width           =   516
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "位置"
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
      Left            =   6372
      TabIndex        =   0
      Top             =   1476
      Width           =   516
   End
End
Attribute VB_Name = "MYEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   MYEdit.frm
'   update : 2004.3.20  s.f.   軸制御ｺﾏﾝﾄﾞ　7追加
'                              現在有効コマンド 0,1,2,3,7,8,9
'　　　　　　　　　　　　　　　　SetData(),GetData()　を変更
'   update: 2004.4.25  s.f     VScroll1(j).min = 210 * lK1     "200"を"210"へ変更
'
'   update: 2004.5. 5  s.f     温度係数、肉厚補正追加 、　チェック温度と待ち時間を5個へ
'   update: 2004.9.26  s.f     温度係数計算方法変更に伴う、計数値変更　0.005単位へ（0.005=0.5℃)
'   update: 2005.6. 4  s.f     加圧時間制御追加
'   update: 2005.6.28  s.f     変更型No変数　追加 text13
'   update: 2006.5.18  s.f     LA, DC　追加
'   update: 2006.5.19  s.f   　My_edit内から　メモ帳　呼び出し、　myeditの　DC　削除
'   update: 2007.09.27 s.f  Z補正　指定したｾｸﾞﾒﾝﾄNo.へ　できるようにする
'   update: 2007.11.21 s.f  加圧制御　平均値計算　今回の加圧時間　重み2.0へ
'                                        -a : 型順　表示　調整数追加
'   update: 2008.2.21 s.f  Z3補正の　Ｎｏ．　変更　Disableへ　変更
'   update: 2008.3.15 s.f  NewQDへ移行に伴う変更。　設定温度　温度３　追加。
'   update: 2008.10. 2 s.f  ｅｄｉｔ中で　「メモ帳」可能に。（キャンセルで抜ける）、　？周目の表示
'   update: 2009. 2.25 s.f  ダミー表示色変更
'   update: 2009. 3.14 s.f  ダミー表示色変更　バグ修正
'
Option Explicit
Dim lK1%, lK2%, lK3%, lK4%, lK5%, lK6%, lK7%, lK8%
Dim lK9%, lK10%, lK11%, lK12%, lK13%, lK14%, lK15%
Dim lFlgDisp%, ikii%, iki%
Dim aHenkou$(3)


Private Sub Command1_Click()
  If katCflag = True Then
          katCflag = False          '加圧時間自動制御　OFF
          Command1().BackColor = CmndColoff(3)
          Command1().Caption = "加圧時間制御 OFF"
    Else
          katCflag = True      '加圧時間自動制御　ON
          Command1().BackColor = CmndColon(1)    ' on1=red
          Command1().Caption = "加圧時間制御 ON"
  End If
'
End Sub

Private Sub Command2_Click(Index As Integer)
Dim ddum1!, ddum2!, ddum3!, itk%, flg%, i%
Select Case Index
Case 0  'キャンセル
  lFlgDisp = False
  SetData
  lFlgDisp = True
  Unload Me
  Select Case ViewFlg
    Case 1
      PGM_Menu.Show
    Case 2    'SC
      NQD70_SC.Show
  End Select
Case 1  '戻る
'  flg = MsgBox("本当に戻りますか？", 49, "MyEdit")  '終了メッセージ
'  If (flg = 2) Then GoTo endcase1:
  GetData
  coxDtSet
  coxDtSave gcoxFldir & gcoxFlName
  Unload Me
  Select Case ViewFlg
    Case 1
      PGM_Menu.Show
    Case 2    'SC
      NQD70_SC.Show
  End Select
endcase1:
Case 2  ' 変更確認
'　----　戻るときの　型増減時の変更処理
'  ----  型　変更時の取り扱い　　　加圧時間データ、T係数、Z3補正
'  ----　’　型変更時の取り扱い 0：変更なし　1：型減少　2：型増加　3：型入替え
  GetData
    Select Case Henkou_No
       Case 0                  '型　変更なし
          '  何もしません
       Case 1                  '型　減少
           T_keisuCont(0) = T_keisuCont(0) - 1
           For ikii = T_keisuCont(1) - 1 To T_keisuCont(0) - 1 + 1
             For iki = 0 To 5: kaatsuJ(ikii, iki) = kaatsuJ(ikii + 1, iki): Next iki
           Next ikii
           For iki = 0 To 5: kaatsuJ(T_keisuCont(0) - 1 + 1, iki) = 0: Next iki
           For ikii = T_keisuCont(1) - 1 To 8
               T_keisu(ikii) = T_keisu(ikii + 1)
               Z3_Hosei(ikii) = Z3_Hosei(ikii + 1)
               iflgKataTorF(ikii) = iflgKataTorF(ikii + 1)
           Next ikii
       Case 2                  '型　増加
          T_keisuCont(0) = T_keisuCont(0) + 1
          For ikii = T_keisuCont(0) - 1 To T_keisuCont(1) - 1 + 1 Step -1
            For iki = 0 To 5: kaatsuJ(ikii + 1, iki) = kaatsuJ(ikii, iki): Next iki
          Next ikii
          For iki = 0 To 5: kaatsuJ(T_keisuCont(1) - 1, iki) = 0: Next iki
          ddum1 = T_keisu(9)
          ddum2 = Z3_Hosei(9)
          ddum3 = iflgKataTorF(9)
          For ikii = 8 To T_keisuCont(1) - 1 Step -1
              T_keisu(ikii + 1) = T_keisu(ikii)
              Z3_Hosei(ikii + 1) = Z3_Hosei(ikii)
              iflgKataTorF(ikii + 1) = iflgKataTorF(ikii)
          Next ikii
          T_keisu(T_keisuCont(1) - 1) = ddum1
          Z3_Hosei(T_keisuCont(1) - 1) = ddum2
          iflgKataTorF(T_keisuCont(1) - 1) = ddum3
       Case 3                  '型入替え
          For iki = 0 To 5: kaatsuJ(T_keisuCont(1) - 1, iki) = 0: Next iki
    End Select
    T_keisu(9) = 1#     '  変更用入力を　reset
    Z3_Hosei(9) = 0#
    iflgKataTorF(9) = True
    Henkou_No = 0    ' 変更No.のリセット
'
    For itk = 0 To 9
       If (itk >= T_keisuCont(0)) Then
            T_keisu(itk) = 1#       '温度係数
            Z3_Hosei(itk) = 0#       '肉厚補正
       End If
    Next itk
    For i = 0 To 9
      If (iflgKataTorF(i) = True) Then
           Command3(i).BackColor = CmndColoff(1)  ' col off 1=灰色
      Else
           Command3(i).BackColor = CmndColon(3)  'col on 3=　薄緑
      End If
    Next i
'
    SetData
    Command2(1).Visible = True
'
Case 3     ' notopad
  SetData
  coxDtSet
  coxDtSave gcoxFldir & gcoxFlName
'  MYEdit.Enabled = False
'  MYEdit.Visible = False
  WaitSec 0.2
  ExecMemo gcoxFldir, gcoxFlName
'  MYEdit.Enabled = True          '  マルチで走るため効かない
'  MYEdit.Visible = True
'
Case 4    ' comment 記入
  SetData
  coxDtSet
  coxDtSave gcoxFldir & gcoxFlName
  WaitSec 0.1
  ExecMemo gcoxFldir, gcoxFlName + ".txt"
'
End Select
End Sub

Private Sub SetData()                        'データの表示  var -> form
Dim i%, inp%, itp%, icp%, ijp%, ilp%, idp%, j%, ja%, k%
Dim itkc%, itk%, iz3hc%, iz3h%
Dim dt!
  '
  SetVScroll                        ' VSScrollの量ｓｅｔ
  '
  Label2(0) = Format(ptime, "###0")  '測定時間
  Label2(1) = Format(ytemp, "###0")  '予備加熱温度
  Label2(2) = gcoxFlName             '制御ファイル名
  Label2(3) = hcomm(2)               'コメント
' -----------------------------------
  For i = 0 To 100
    Select Case ic(i)
    Case 0
      inp = inp + 1
      If inp < 7 Then
        dt = z(i): DispText1 inp, dt, True      '位置
        dt = vel(i): DispText2 inp, dt, True    '速度
        dt = pres(i): DispText3 inp, dt, False  '圧力
      End If
    Case 1, 3, 7
      inp = inp + 1
      If inp < 7 Then
        dt = z(i): DispText1 inp, dt, True      '位置
        dt = vel(i): DispText2 inp, dt, True    '速度
        dt = pres(i): DispText3 inp, dt, True   '圧力
      End If
    Case 2
    Case 8
    Case 9
      Exit For
    End Select
  Next i
  If inp < 7 Then
    For i = inp + 1 To 6
      DispText1 i, dt, False    '位置
      DispText2 i, dt, False    '速度
      DispText3 i, dt, False    '圧力
    Next i
  End If
' ---------------------- 制御コマンド
  For i = 0 To 200
    Select Case Left(scom(i), 1)
    Case "S"
      j = j + 1
      itp = itp + 1
      If itp < 7 Then
        dt = sisub(i): DispText4 itp, dt, True     '温度１　　成形室IH
        dt = sjsub(i): DispText5 itp, dt, True     '温度２　　上型
        dt = sksub(i): DispText6 itp, dt, True     '温度３　　下型
        dt = slsub(i): DispText15 itp, dt, True     '昇温時間
      End If
    Case "L"
      ilp = 7
        dt = sisub(i): DispText4 ilp, dt, True     '保温停止温度１
        dt = sjsub(i): DispText5 ilp, dt, True     '5分停止温度２
        dt = sksub(i): DispText6 ilp, dt, True     'ダミー時間
        dt = slsub(i): DispText15 ilp, dt, True     'ダミー
    Case "T"
      icp = icp + 1
      If icp < 6 Then
        dt = sisub(i): DispText7 icp, dt, True     'チェック温度
      End If
    Case "J"
      ijp = ijp + 1
      If ijp < 6 Then
        dt = sisub(i): DispText8 ijp, dt, True     '待ち時間
      End If
    Case "P"
    Case "E"
      Exit For
    End Select
  Next i
  '
  If itp < 7 Then
    dt = 0
    For i = itp + 1 To itp
      DispText4 i, dt, False     '温度１
      DispText5 i, dt, False     '温度２
      DispText6 i, dt, False     '温度３
      DispText15 i, dt, False     '昇温時間
    Next i
  End If
  If icp < 6 Then
    dt = 0
    For i = icp + 1 To icp
      DispText7 i, dt, False     'チェック温度
    Next i
  End If
  If ijp < 6 Then
    dt = 0
    For i = ijp + 1 To ijp
      DispText8 i, dt, False     '待ち時間
    Next i
  End If
'
   dt = T_keisuCont(0): DispText9 0 + 1, dt, True   '温度係数コントロール 型個数
   dt = T_keisuCont(1): DispText9 1 + 1, dt, True   '温度係数コントロール　ポインター
   dt = Z3_HoseiCont(2): DispText16 0 + 1, dt, True   'Z補正コントロール　　Z　No.
   dt = ishu: DispText9 2 + 1, dt, True   ' 成形　？周目
   For itk = 0 To 9
   If itk < T_keisuCont(0) Then
        dt = T_keisu(itk): DispText11 itk + 1, dt, True   '温度係数
      Else
        dt = 1#: DispText11 itk + 1, dt, True
      End If
   Next itk
  
'   ----
   dt = DkatJ(1): DispText10 1 + 1, dt, True   '加圧時間目標値上限
   dt = DkatJ(0): DispText10 0 + 1, dt, True   '加圧時間目標値下限
   For iz3h = 0 To 8
   If iz3h < T_keisuCont(0) Then
        dt = Z3_Hosei(iz3h): DispText12 iz3h + 1, dt, True   '肉厚補正
      Else
        dt = 0: DispText12 iz3h + 1, dt, True
      End If
   Next iz3h
' ----
   dt = AkatJ(1): DispText10 3 + 1, dt, True   '加圧時間ALARM上限
   dt = AkatJ(0): DispText10 2 + 1, dt, True   '加圧時間ALARM下限
'
   dt = Acp(1): DispText1 7 + 1, dt, True   'Ｃｐ位置ALARM上限
   dt = Acp(0): DispText1 6 + 1, dt, True   'Ｃｐ位置ALARM下限
'
   dt = Henkou_No: DispText13 1, dt, True    '変更型内容
' ---- kataNo. set
    For i = 0 To 9
        Text14(i).Text = kataNo(i)
    Next i
   dt = kataNo(10): DispText14 1, dt, True   '変No. 調整数
  '------------------ グラフ
  MyEditGph Me.Picture1
End Sub

Private Sub GetData()                     '  form -> variable
Dim i%, inp%, itp%, icp%, ijp%, ilp%, idp%, j%, ja%, k%
Dim itkc%, itk%, iz3hc%, iz3h%
Dim dt!
' -----------------------------------
  For i = 0 To 100
    Select Case ic(i)
    Case 0
      inp = inp + 1
      If inp < 7 Then
        z(i) = Val(Text1(inp - 1).Text)     '位置
        vel(i) = Val(Text2(inp - 1).Text)   '速度
        pres(i) = Val(Text3(inp - 1).Text)  '圧力
      End If
    Case 1, 3, 7
      inp = inp + 1
      If inp < 7 Then
        z(i) = Val(Text1(inp - 1).Text)     '位置
        vel(i) = Val(Text2(inp - 1).Text)   '速度
        pres(i) = Val(Text3(inp - 1).Text)  '圧力
      End If
    Case Else
      'Exit For
    End Select
  Next i
  
' ---------------------- 制御コマンド
  For i = 0 To 200
    Select Case Left(scom(i), 1)
    Case "S"
      j = j + 1
      itp = itp + 1
      If itp < 7 Then
        sisub(i) = Val(Text4(itp - 1).Text)    '温度１　成形室IH
        sjsub(i) = Val(Text5(itp - 1).Text)    '温度２　上型
        sksub(i) = Val(Text6(itp - 1).Text)    '温度３　下型
        slsub(i) = Val(Text15(itp - 1).Text)    '昇温時間
      End If
    Case "L"
      ilp = 7
        sisub(i) = Val(Text4(ilp - 1).Text)    '保温停止温度１
        sjsub(i) = Val(Text5(ilp - 1).Text)    '5分停止温度２
        sksub(i) = Val(Text6(ilp - 1).Text)    'ダミー時間
        slsub(i) = Val(Text15(ilp - 1).Text)    'ダミー
    Case "T"
      icp = icp + 1
      If icp < 6 Then
        sisub(i) = Val(Text7(icp - 1).Text)    'チェック温度
      End If
    Case "J"
      ijp = ijp + 1
      If ijp < 6 Then
        sisub(i) = Val(Text8(ijp - 1).Text)    '待ち時間
      End If
    Case "P"
    Case "E"
      Exit For
    End Select
  Next i
'
      T_keisuCont(0) = Val(Text9(0).Text)    '温度係数ｃｏｎｔ　型個数
      T_keisuCont(1) = Val(Text9(1).Text)    '温度係数ｃｏｎｔ　ポインター
      Z3_HoseiCont(2) = Val(Text16(0).Text)    'Z補正ｃｏｎｔ　ポインター
      ishu = Val(Text9(2).Text)              ' 成形　何周目
      Henkou_No = Val(Text13.Text)    '変更内容         ----- 05.11.23
'
      For itk = 0 To 9
            T_keisu(itk) = Val(Text11(itk).Text)    '温度係数
      Next itk
      If T_keisuCont(1) > T_keisuCont(0) Then T_keisuCont(1) = T_keisuCont(0)       '　誤入力の訂正
'
'
      DkatJ(1) = Val(Text10(1).Text)    '加圧時間目標値上限
      DkatJ(0) = Val(Text10(0).Text)    '加圧時間目標値下限
      For iz3h = 0 To 8
         Z3_Hosei(iz3h) = Val(Text12(iz3h).Text)    '肉厚補正
      Next iz3h
'
      AkatJ(1) = Val(Text10(3).Text)    '加圧時間ALARM上限
      AkatJ(0) = Val(Text10(2).Text)    '加圧時間ALARM下限
      Acp(1) = Val(Text1(7).Text)    'Ｃｐ位置ALARM上限
      Acp(0) = Val(Text1(6).Text)    'Ｃｐ位置ALARM下限
'
'  --- kataNo. ---
    For i = 0 To 10
        kataNo(i) = Text14(i).Text  '   型No.　の取り込み
    Next i
'
End Sub

Private Sub Command3_Click(Index As Integer)
       If (iflgKataTorF(Index) = True) Then
          iflgKataTorF(Index) = False              '　ﾀﾞﾐｰは薄緑　false
          Command3(Index).BackColor = CmndColon(3)  'col on 3=薄緑
         Else
          iflgKataTorF(Index) = True
          Command3(Index).BackColor = CmndColoff(1)  ' col off 1=　灰色
       End If
End Sub

Private Sub Form_Load()
Dim i%
'
aHenkou$(0) = "変更なし"
aHenkou$(1) = "型減らす"
aHenkou$(2) = "型増やす"
aHenkou$(3) = "型入替え"
'
'
Command2(1).Visible = False
'
DispCenter Me
  lFlgDisp = False
  coxDtRead gcoxFldir & gcoxFlName
  If T_keisuCont(2) <> 0 Then T_keisuCont(1) = T_keisuCont(2)    'ポインターのbackup
  If T_keisuCont(3) <> 0 Then T_keisuCont(0) = T_keisuCont(3)    '型個数のbackup
  If ishu_bkup <> 0 Then ishu = ishu_bkup                        '?週目のbackupを代入
'  Henkou_No = 0    '変更は、NQD70_SC側で
  coxDtSet
  SetData
  lFlgDisp = True
  For i = 0 To 9
    If (iflgKataTorF(i) = True) Then
           Command3(i).BackColor = CmndColoff(1)  '  off 1=灰色
    Else
           Command3(i).BackColor = CmndColon(3)  'on 3=　薄緑
    End If
  Next i
  If katCflag = False Then
          Command1().BackColor = CmndColoff(3)
          Command1().Caption = "加圧時間制御 OFF"
    Else
          Command1().BackColor = CmndColon(1)    ' on1=red
          Command1().Caption = "加圧時間制御 ON"
    End If
End Sub

Private Sub DispText1(i%, dt!, flg%)   '  z　位置
  If flg = False Then
    VScroll1(i - 1).Visible = False
    Text1(i - 1).Visible = False
  Else
    VScroll1(i - 1).Visible = True
    VScroll1(i - 1).Value = dt * lK1
    Text1(i - 1).Visible = True
    Text1(i - 1).Text = Format(dt, "###0.00")
  End If
End Sub
Private Sub DispText2(i%, dt!, flg%)    ' 速度
  If flg = False Then
    VScroll2(i - 1).Visible = False
    Text2(i - 1).Visible = False
  Else
    VScroll2(i - 1).Visible = True
    VScroll2(i - 1).Value = dt * lK2
    Text2(i - 1).Visible = True
    Text2(i - 1).Text = Format(dt, "###0.0")
  End If
End Sub
Private Sub DispText3(i%, dt!, flg%)    '　圧力
  If flg = False Then
    VScroll3(i - 1).Visible = False
    Text3(i - 1).Visible = False
  Else
    VScroll3(i - 1).Visible = True
    VScroll3(i - 1).Value = dt * lK3
    Text3(i - 1).Visible = True
    Text3(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText4(i%, dt!, flg%)    '　温度１
  If flg = False Then
    VScroll4(i - 1).Visible = False
    Text4(i - 1).Visible = False
  Else
    VScroll4(i - 1).Visible = True
    VScroll4(i - 1).Value = dt * lK4
    Text4(i - 1).Visible = True
    Text4(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText5(i%, dt!, flg%)    '　　温度２
  If flg = False Then
    VScroll5(i - 1).Visible = False
    Text5(i - 1).Visible = False
  Else
    VScroll5(i - 1).Visible = True
    VScroll5(i - 1).Value = dt * lK5
    Text5(i - 1).Visible = True
    Text5(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText6(i%, dt!, flg%)    '　　温度3
  If flg = False Then
    VScroll6(i - 1).Visible = False
    Text6(i - 1).Visible = False
  Else
    VScroll6(i - 1).Visible = True
    VScroll6(i - 1).Value = dt * lK6
    Text6(i - 1).Visible = True
    Text6(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText7(i%, dt!, flg%)      '　　チェック温度
  If flg = False Then
    VScroll7(i - 1).Visible = False
    Text7(i - 1).Visible = False
  Else
    VScroll7(i - 1).Visible = True
    VScroll7(i - 1).Value = dt * lK7
    Text7(i - 1).Visible = True
    Text7(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText8(i%, dt!, flg%)    '　　待ち時間
  If flg = False Then
    VScroll8(i - 1).Visible = False
    Text8(i - 1).Visible = False
  Else
    VScroll8(i - 1).Visible = True
    VScroll8(i - 1).Value = dt * lK8
    Text8(i - 1).Visible = True
    Text8(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText9(i%, dt!, flg%)    '　　温度係数ｃｏｎｔ
  If flg = False Then
    VScroll9(i - 1).Visible = False
    Text9(i - 1).Visible = False
  Else
    VScroll9(i - 1).Visible = True
    VScroll9(i - 1).Value = dt * lK9
    Text9(i - 1).Visible = True
    Text9(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText10(i%, dt!, flg%)    '　　加圧時間目標値　上限・下限
  If flg = False Then
    VScroll10(i - 1).Visible = False
    Text10(i - 1).Visible = False
  Else
    VScroll10(i - 1).Visible = True
    VScroll10(i - 1).Value = dt * lK10
    Text10(i - 1).Visible = True
    Text10(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText11(i%, dt!, flg%)   '  温度係数
  If flg = False Then
    VScroll11(i - 1).Visible = False
    Text11(i - 1).Visible = False
  Else
    VScroll11(i - 1).Visible = True
    VScroll11(i - 1).Value = dt * lK11
    Text11(i - 1).Visible = True
    Text11(i - 1).Text = Format(dt, "###0.000")    '  '04.9.26
  End If
End Sub
Private Sub DispText12(i%, dt!, flg%)   '  温度係数
  If flg = False Then
    VScroll12(i - 1).Visible = False
    Text12(i - 1).Visible = False
  Else
    VScroll12(i - 1).Visible = True
    VScroll12(i - 1).Value = dt * lK12
    Text12(i - 1).Visible = True
    Text12(i - 1).Text = Format(dt, "###0.000")
  End If
End Sub
Private Sub DispText13(i%, dt!, flg%)   '  変更内容
  If flg = False Then
    VScroll13.Visible = False
    Text13.Visible = False
  Else
    VScroll13.Visible = True
    VScroll13.Value = dt * lK13
    Text13.Visible = True
    Text13.Text = Format(dt, "##0")
    Label8.Caption = aHenkou$(dt)
  End If
End Sub
Private Sub DispText14(i%, dt!, flg%)   '  型Ｎｏ． 調整数
  If flg = False Then
    VScroll14.Visible = False
    Text14(10).Visible = False
  Else
    VScroll14.Visible = True
    VScroll14.Value = dt * lK14
    Text14(10).Visible = True
    Text14(10).Text = Format(dt, "#0")
  End If
End Sub
Private Sub DispText15(i%, dt!, flg%)    '　　昇温時間
  If flg = False Then
    VScroll15(i - 1).Visible = False
    Text15(i - 1).Visible = False
  Else
    VScroll15(i - 1).Visible = True
    VScroll15(i - 1).Value = dt * lK15
    Text15(i - 1).Visible = True
    Text15(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub DispText16(i%, dt!, flg%)    '　　Z補正
  If flg = False Then
    Text16(i - 1).Visible = False
  Else
    Text16(i - 1).Visible = True
    Text16(i - 1).Text = Format(dt, "###0")
  End If
End Sub
Private Sub SetVScroll()               ' VSScrollの量ｓｅｔ
Dim j%
' ---------------- Z
  For j = 0 To 7                        ' 6,7 追加　090307
    lK1 = 100
    VScroll1(j).min = 210 * lK1       '2004.4.25 S.F 200を210へ変更
    VScroll1(j).max = 0 * lK1
    VScroll1(j).LargeChange = 1 * lK1
    VScroll1(j).SmallChange = 0.01 * lK1
  Next j
' ---------------- V
  For j = 0 To 5
    lK2 = 1
    VScroll2(j).min = 3000 * lK2
    VScroll2(j).max = 0 * lK2
    VScroll2(j).LargeChange = 10 * lK2
    VScroll2(j).SmallChange = 1 * lK2
  Next j
' ---------------- P
  For j = 0 To 5
    lK3 = 1
    VScroll3(j).min = 1000 * lK3
    VScroll3(j).max = 0 * lK3
    VScroll3(j).LargeChange = 10 * lK3
    VScroll3(j).SmallChange = 1 * lK3
  Next j
' ---------------- T　設定温度　スリーブ
  For j = 0 To 6
    lK4 = 1
    VScroll4(j).min = 1000 * lK4
    VScroll4(j).max = -100 * lK4
    VScroll4(j).LargeChange = 10 * lK4
    VScroll4(j).SmallChange = 1 * lK4
  Next j
' ---------------- T 設定温度　　上型
  For j = 0 To 6
    lK5 = 1
    VScroll5(j).min = 1000 * lK5
    VScroll5(j).max = -100 * lK5
    VScroll5(j).LargeChange = 10 * lK5
    VScroll5(j).SmallChange = 1 * lK5
  Next j
' ---------------- T　設定温度　下型
  For j = 0 To 6
    lK6 = 1
    VScroll6(j).min = 1000 * lK6
    VScroll6(j).max = -100 * lK6
    VScroll6(j).LargeChange = 10 * lK6
    VScroll6(j).SmallChange = 1 * lK6
  Next j
' ---------------- チェック温度
  For j = 0 To 4
    lK7 = 1
    VScroll7(j).min = 1000 * lK7
    VScroll7(j).max = 0 * lK7
    VScroll7(j).LargeChange = 10 * lK7
    VScroll7(j).SmallChange = 1 * lK7
  Next j
' ---------------- 待ち時間
  For j = 0 To 4
    lK8 = 1
    VScroll8(j).min = 1000 * lK8
    VScroll8(j).max = 0 * lK8
    VScroll8(j).LargeChange = 10 * lK8
    VScroll8(j).SmallChange = 1 * lK8
  Next j
' ---------------- 温度係数ｃｏｎｔ、　型個数、ポインター
  For j = 0 To 2        ' j=3 は、VISIBLE=falseのため　0,1,2まで
    lK9 = 1
    VScroll9(j).min = 50 * lK9      '  max=10 -> max=50 he  2010.6.23
    VScroll9(j).max = 1 * lK9
    VScroll9(j).LargeChange = 1 * lK9
    VScroll9(j).SmallChange = 1 * lK9
  Next j
' ---------------- 加圧時間目標値　上限
  For j = 0 To 3  'アラーム設定　加圧時間
    lK10 = 1
    VScroll10(j).min = 1000 * lK10
    VScroll10(j).max = 0 * lK10
    VScroll10(j).LargeChange = 1 * lK10
    VScroll10(j).SmallChange = 1 * lK10
  Next j
' ---------------- 温度係数
  For j = 0 To 9
    lK11 = 1000
    VScroll11(j).min = 1.2 * lK11
    VScroll11(j).max = 0.8 * lK11
    VScroll11(j).LargeChange = 0.01 * lK11
    VScroll11(j).SmallChange = 0.001 * lK11
  Next j
' ---------------- 肉厚補正
  For j = 0 To 8
    lK12 = 1000
    VScroll12(j).min = 0.05 * lK12
    VScroll12(j).max = -0.05 * lK12
    VScroll12(j).LargeChange = 0.01 * lK12
    VScroll12(j).SmallChange = 0.005 * lK12
  Next j
' ---------------- 型変更内容
    lK13 = 1
    VScroll13.min = 3 * lK13
    VScroll13.max = 0 * lK13
    VScroll13.LargeChange = 1 * lK13
    VScroll13.SmallChange = 1 * lK13
' ---------------- 型Ｎｏ．　調整数
    lK14 = 1
    VScroll14.min = 5 * lK14
    VScroll14.max = -5 * lK14
    VScroll14.LargeChange = 1 * lK14
    VScroll14.SmallChange = 1 * lK14
' ---------------- 昇温時間
  For j = 0 To 6
    lK15 = 1
    VScroll15(j).min = 1000 * lK15
    VScroll15(j).max = 0 * lK15
    VScroll15(j).LargeChange = 10 * lK15
    VScroll15(j).SmallChange = 1 * lK15
  Next j
End Sub

  Private Sub VScroll1_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll1(Index).Value / lK1
  inp = Index + 1
  DispText1 inp, dt, True      'Ｚ位置
End Sub
Private Sub VScroll2_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll2(Index).Value / lK2
  inp = Index + 1
  DispText2 inp, dt, True      '速度
End Sub
Private Sub VScroll3_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll3(Index).Value / lK3
  inp = Index + 1
  DispText3 inp, dt, True      '圧力
End Sub
Private Sub VScroll4_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll4(Index).Value / lK4
  inp = Index + 1
  DispText4 inp, dt, True      '温度１　　スリーブ（成形室　IH)
End Sub
Private Sub VScroll5_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll5(Index).Value / lK5
  inp = Index + 1
  DispText5 inp, dt, True      '温度２　上型
End Sub
Private Sub VScroll6_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll6(Index).Value / lK6
  inp = Index + 1
  DispText6 inp, dt, True      '温度３　下型
End Sub
Private Sub VScroll7_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll7(Index).Value / lK7
  inp = Index + 1
  DispText7 inp, dt, True      'チェック温度
End Sub
Private Sub VScroll8_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll8(Index).Value / lK8
  inp = Index + 1
  DispText8 inp, dt, True      '待ち時間
End Sub
Private Sub VScroll9_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll9(Index).Value / lK9
  inp = Index + 1
  DispText9 inp, dt, True      '温度係数ｃｏｎｔ
End Sub
Private Sub VScroll10_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll10(Index).Value / lK10
  inp = Index + 1
  DispText10 inp, dt, True      '加圧時間目標値
End Sub
Private Sub VScroll11_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll11(Index).Value / lK11
  inp = Index + 1
  DispText11 inp, dt, True      '温度係数
End Sub
Private Sub VScroll12_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll12(Index).Value / lK12
  inp = Index + 1
  DispText12 inp, dt, True      '肉厚補正
End Sub
Private Sub VScroll13_Change()
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll13.Value / lK13
  inp = 1
  DispText13 inp, dt, True      '変更型内容
End Sub
Private Sub VScroll14_Change()
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll14.Value / lK14
  inp = 1
  DispText14 inp, dt, True      '変更型内容
End Sub
Private Sub VScroll15_Change(Index As Integer)
Dim dt!, inp%
  If lFlgDisp = False Then Exit Sub
  dt = VScroll15(Index).Value / lK15
  inp = Index + 1
  DispText15 inp, dt, True      '昇温時間
End Sub

