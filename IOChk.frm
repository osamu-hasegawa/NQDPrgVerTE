VERSION 5.00
Begin VB.Form IOChk 
   Caption         =   "I/O Check"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command1 
      Caption         =   "â∑ìxâ∫å^ê›íË(Åé)"
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   128
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'âEëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   127
      Text            =   "5000"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'âEëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   120
      Text            =   "0"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'âEëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   119
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "â∑ìxè„å^ê›íË(Åé)"
      Height          =   495
      Index           =   2
      Left            =   5160
      TabIndex        =   118
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'âEëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   117
      Text            =   "0"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "â∑ìxIHê›íË(Åé)"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   116
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉÇÉjÉ^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   6480
      TabIndex        =   115
      Top             =   120
      Width           =   1236
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9000
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   114
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ReSet"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   113
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "å¥ì_"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   112
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Status"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   111
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CCW(â∫)"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   110
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CW(è„)"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   109
      Top             =   5040
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   31
      Left            =   9000
      TabIndex        =   94
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   30
      Left            =   8640
      TabIndex        =   93
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   29
      Left            =   8280
      TabIndex        =   92
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   28
      Left            =   7920
      TabIndex        =   91
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   27
      Left            =   7560
      TabIndex        =   90
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   26
      Left            =   7200
      TabIndex        =   89
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   25
      Left            =   6840
      TabIndex        =   88
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   24
      Left            =   6480
      TabIndex        =   87
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   23
      Left            =   9000
      TabIndex        =   86
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   22
      Left            =   8640
      TabIndex        =   85
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   21
      Left            =   8280
      TabIndex        =   84
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   20
      Left            =   7920
      TabIndex        =   83
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   19
      Left            =   7560
      TabIndex        =   82
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   18
      Left            =   7200
      TabIndex        =   81
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   80
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   16
      Left            =   6480
      TabIndex        =   79
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   15
      Left            =   9000
      TabIndex        =   78
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   14
      Left            =   8640
      TabIndex        =   77
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   13
      Left            =   8280
      TabIndex        =   76
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   12
      Left            =   7920
      TabIndex        =   75
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   11
      Left            =   7560
      TabIndex        =   74
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   73
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   72
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   71
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   70
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   69
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   68
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   67
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   66
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   65
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   64
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   63
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉÇÉjÉ^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   4680
      TabIndex        =   20
      Top             =   120
      Width           =   1236
   End
   Begin VB.Frame Frame1 
      Caption         =   "ê≥ì]Å^ãtì]"
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "OFF"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ãtì]"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ê≥ì]"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉTÅ[É{"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1236
   End
   Begin VB.CheckBox Check1 
      Caption         =   "à íu/ë¨ìxêÿë÷Åië¨ìxëIëÅj"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reset"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÉTÅ[É{ ON/OFF"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ë¨ìxê›íË(V)"
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'âEëµÇ¶
      Appearance      =   0  'Ã◊Øƒ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉLÉÉÉìÉZÉã"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉÅÉjÉÖÅ["
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1410
      TabIndex        =   0
      Top             =   135
      Width           =   1236
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "âÒì]ÉXÉsÅ[Éh"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   4440
      TabIndex        =   131
      Top             =   5280
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   240
      TabIndex        =   130
      Top             =   5400
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÅiÇPâÒì]ÇÃÇ›Åj"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   129
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch8)ÅFãÛÇ´"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   6360
      TabIndex        =   126
      Top             =   5400
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch7)ÅFê¨å`â∫å^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   6360
      TabIndex        =   125
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch6)ÅFê¨å`è„å^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   6360
      TabIndex        =   124
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   8520
      TabIndex        =   123
      Top             =   5400
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   8520
      TabIndex        =   122
      Top             =   5040
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   8520
      TabIndex        =   121
      Top             =   4680
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "INDEX DRIVE"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   108
      Top             =   5040
      Width           =   1365
   End
   Begin VB.Label Message_Label 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Height          =   255
      Left            =   1920
      TabIndex        =   107
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "Çwé≤"
      Height          =   180
      Left            =   1920
      TabIndex        =   106
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Addr_Label 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Height          =   255
      Left            =   1920
      TabIndex        =   105
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÉfÉBÉWÉ^ÉãèoóÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   6480
      TabIndex        =   104
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÉfÉBÉWÉ^Éãì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   6480
      TabIndex        =   103
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   6480
      TabIndex        =   102
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   14
      Left            =   6840
      TabIndex        =   101
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   13
      Left            =   7200
      TabIndex        =   100
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   12
      Left            =   7560
      TabIndex        =   99
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Index           =   11
      Left            =   7920
      TabIndex        =   98
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   10
      Left            =   8280
      TabIndex        =   97
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "7"
      Height          =   180
      Index           =   9
      Left            =   8640
      TabIndex        =   96
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "8"
      Height          =   180
      Index           =   8
      Left            =   9000
      TabIndex        =   95
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   31
      Left            =   8970
      TabIndex        =   62
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   30
      Left            =   8610
      TabIndex        =   61
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   29
      Left            =   8250
      TabIndex        =   60
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   28
      Left            =   7890
      TabIndex        =   59
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   27
      Left            =   7530
      TabIndex        =   58
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   26
      Left            =   7170
      TabIndex        =   57
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   25
      Left            =   6810
      TabIndex        =   56
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   24
      Left            =   6480
      TabIndex        =   55
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   23
      Left            =   8970
      TabIndex        =   54
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   22
      Left            =   8610
      TabIndex        =   53
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   21
      Left            =   8250
      TabIndex        =   52
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   20
      Left            =   7890
      TabIndex        =   51
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   19
      Left            =   7530
      TabIndex        =   50
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   18
      Left            =   7170
      TabIndex        =   49
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   17
      Left            =   6810
      TabIndex        =   48
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   16
      Left            =   6480
      TabIndex        =   47
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   15
      Left            =   8970
      TabIndex        =   46
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   14
      Left            =   8610
      TabIndex        =   45
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   13
      Left            =   8250
      TabIndex        =   44
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   12
      Left            =   7890
      TabIndex        =   43
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   11
      Left            =   7530
      TabIndex        =   42
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   10
      Left            =   7170
      TabIndex        =   41
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   9
      Left            =   6810
      TabIndex        =   40
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   8
      Left            =   6480
      TabIndex        =   39
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "8"
      Height          =   180
      Index           =   7
      Left            =   9000
      TabIndex        =   38
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "7"
      Height          =   180
      Index           =   6
      Left            =   8640
      TabIndex        =   37
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   5
      Left            =   8280
      TabIndex        =   36
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Index           =   4
      Left            =   7920
      TabIndex        =   35
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   3
      Left            =   7560
      TabIndex        =   34
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   2
      Left            =   7200
      TabIndex        =   33
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   1
      Left            =   6840
      TabIndex        =   32
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   6480
      TabIndex        =   31
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   7
      Left            =   9000
      TabIndex        =   30
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   6
      Left            =   8640
      TabIndex        =   29
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   5
      Left            =   8280
      TabIndex        =   28
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   4
      Left            =   7920
      TabIndex        =   27
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   3
      Left            =   7560
      TabIndex        =   26
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   2
      Left            =   7200
      TabIndex        =   25
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   1
      Left            =   6840
      TabIndex        =   24
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Åú"
      Height          =   180
      Index           =   0
      Left            =   6510
      TabIndex        =   23
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch5)ÅFãÛÇ´"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2520
      TabIndex        =   22
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4920
      TabIndex        =   21
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4920
      TabIndex        =   19
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4920
      TabIndex        =   18
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4920
      TabIndex        =   17
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "#######"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch4):ãÛÇ´"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2520
      TabIndex        =   15
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch3):â◊èd"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch2):ãÛÇ´"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AD(Ch1):ê¨å`é∫"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   720
      Width           =   1620
   End
End
Attribute VB_Name = "IOChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
Dim ch%, dt%
  dt = Check1(Index).Value
  Select Case Index
  Case 0      'ÉTÅ[É{ON
    ch = 10         '22
    If Check1(Index).Value = 1 Then
        DioOut ch, 1
    Else
        DioOut ch, 0
    End If
  Case 1      'Reset
    ch = 11         '23
    If Check1(Index).Value = 1 Then
        DioOut ch, 1    '26, 1
    Else
        DioOut ch, 0     '26, 0
    End If
  Case 2      'ë¨ìxÅ^à íu
    ch = 12         '24
    If Check1(Index).Value = 1 Then
'
        DioOut ch, 0        ' ë¨ìxÅ@Å@ìåâhÇÃéû
'        DioOut ch, 0        ' à íuÅ@Å@í÷ÇÃéû
'        DioOut 13, 1        ' à íuÅ@Å@í÷ÇÃéû
    Else
'
        DioOut ch, 1        'à íuÅ@Å@ìåâhÇÃéû
'        DioOut ch, 1        'ë¨ìxÅ@Å@í÷ÇÃéû
'        DioOut 13, 0        'ë¨ìxÅ@Å@í÷ÇÃéû
    End If
  End Select
End Sub

Private Sub Check2_Click(Index As Integer)
Dim ch%, dt%
  dt = Check2(Index).Value
  ch = Index + 1
  DioOut ch, dt
  Select Case ch
  Case 22 'ÉTÅ[É{ON
  Case 23 'Reset
  Case 24 'ë¨ìxÅ^à íu
  End Select
End Sub

Private Sub Command1_Click(Index As Integer)

Dim v!, ch%
  ch = Index + 1
  v = Val(Text1(Index).Text)
  Select Case Index
  Case 0    'ë¨ìxê›íË(V)
    DaVoltOut ch, v
  Case 1    'â∑ìxihê›íË(V)
    'DaVoltOut ch, v
    TempSet ch, v
  Case 2    'â∑ìxè„å^ê›íË(V)
'    DaVoltOut ch, v
    TempSet ch, v
  Case 3    'â∑ìxâ∫å^ê›íË(V)
'    DaVoltOut ch, v
    TempSet ch, v
  End Select
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0  'ÉLÉÉÉìÉZÉã
  lFlgDisp = False
  SetData
  lFlgDisp = True
Case 1  'èIóπ
  GetData
  Unload Me
  PGM_Menu.Show
Case 2
  'adMain.Show
  MplVbSmp.Show
Case 3      'ÉÇÉjÉ^
  Moni
  DispDinput
Case 4      'ÉÇÉjÉ^
  If Command2(Index).Caption = "ÉÇÉjÉ^ON" Then
    Timer1.Enabled = True
    Command2(Index).Caption = "ÉÇÉjÉ^OFF"
  Else
    Timer1.Enabled = False
    Command2(Index).Caption = "ÉÇÉjÉ^ON"
  End If
End Select
End Sub

Private Sub SetData()

End Sub

Private Sub GetData()

End Sub

Private Sub Command3_Click(Index As Integer)
Dim dt!, sdt$
Dim vel As Long
    vel = Val(Text1(4))
  Select Case Index
  Case 0    'CW
    C870HSPDSet vel
    'C870OrgVelSet
    Cw_Index Me
  Case 1    'CCW
    'C870OrgVelSet
    C870HSPDSet vel
    Ccw_Index Me
  Case 2    'Status
    dt = C870Sts(1)
    sdt = Hex(dt)
    dt = C870Sts(2)
    sdt = sdt & "  " & Hex(dt)
    dt = C870Sts(3)
    sdt = sdt & "  " & Hex(dt)
'
    Addr_Label.Caption = sdt
    'Ready_Wait
  Case 3    'å¥ì_
    genten
  Case 4    'ReSet
    C870Reset
  Case 5    'Stop
    C870Stop
  End Select
  Drive_Stop_Disp Me
End Sub

Private Sub Form_Load()
  DispCenter Me
End Sub


Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0      'CW
    DioOut 29, 1
    DioOut 30, 0
Case 1      'CCW
    DioOut 29, 0
    DioOut 30, 1
Case 2      'OFF
    DioOut 29, 0
    DioOut 30, 0
End Select
End Sub

Private Sub Moni()
Dim ch%, hdt!
Dim dt!(0 To 7)
Dim flg As Long
  AdRead dt(), flg
'    For ch = 1 To 5
    For ch = 1 To 8
      'hdt = AdRead1Ch(Ch)
      Select Case ch
      Case 1     'AD(Ch1): ê¨å`é∫Å@IHÉqÅ[É^Å@â∑ìx
        hdt = dt(ch - 1) * 1000 / 10 'TempRdMold(0)
        Label2(ch - 1).Caption = Format(hdt, "0.0")
      Case 2     'AD(2):ñ¢ê⁄ë±
        hdt = 0
        Label2(ch - 1).Caption = Format(hdt, "0.0")
      Case 3    'â◊èd
        hdt = LoadSet(dt(ch - 1))
        Label2(ch - 1).Caption = Format(hdt, "0.0")
      Case 4    'ñ¢ê⁄ë±
        hdt = dt(ch - 1) * 1500 / 10
        Label2(ch - 1).Caption = Format(hdt, "0.000")
      Case 5    'ñ¢ê⁄ë±
        hdt = dt(ch - 1)
        Label2(ch - 1).Caption = Format(hdt, "0.000")
      Case 6     'AD(Ch6): ê¨å`é∫Å@è„å^â∑ìx
        hdt = dt(ch - 1) * 1000 / 10 'TempRdMold(5)
        Label2(ch - 1).Caption = Format(hdt, "0.0")
      Case 7     'AD(Ch7): ê¨å`é∫Å@â∫å^â∑ìx
        hdt = dt(ch - 1) * 1000 / 10 'TempRdMold(6)
        Label2(ch - 1).Caption = Format(hdt, "0.0")
      End Select
    Next ch
  If flg <> AD_ERROR_SUCCESS Then
    DsplyErrMessageDA flg
  End If
End Sub
Private Sub DispDinput()
Dim i%, ch%, hdt%
  For ch = 1 To 32
    DioInput ch, hdt
    Label3(ch - 1).Caption = Format(hdt, "0")
  Next ch
End Sub
Private Sub DispDoutput()

End Sub


Private Sub Timer1_Timer()
  Moni
  DispDinput
End Sub
Private Sub genten()
'--------------
  C870Genten
'/* ÉJÉEÉìÉ^Ç…É[ÉçÇèëÇ´çûÇﬁ */
  'C870CntPreSet 0   'ÇbÇnÇtÇmÇsÇdÇq ÇoÇqÇdÇrÇdÇs ÇbÇnÇlÇlÇ`ÇmÇc
'/* éËìÆópÅ@ë¨ìxÇ÷ÇÃïœçX */
  'C870HSPDSet 36256    '/* 36256 pps  3mm/sec */
End Sub
