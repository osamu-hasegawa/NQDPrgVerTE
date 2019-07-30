VERSION 5.00
Begin VB.Form coolingform 
   Caption         =   "coolingform"
   ClientHeight    =   2268
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4212
   LinkTopic       =   "Form1"
   ScaleHeight     =   2268
   ScaleWidth      =   4212
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command1 
      Caption         =   "Ex"
      Height          =   252
      Left            =   3600
      TabIndex        =   3
      Top             =   1920
      Width           =   492
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   2532
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
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
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "ó‚ãpë“Çøéûä‘ÇÃÇ®ímÇÁÇπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   13.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3732
   End
End
Attribute VB_Name = "coolingform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  DispCenter Me
  Label3.Caption = "ê¨å`èIóπ" + Time$
End Sub

Private Sub Timer1_Timer()
   hyouji
End Sub
Private Sub hyouji()
Dim imachi%
Dim evtime!, difft!
'
 
'  Å@ó‚ãpéûä‘ë“ÇøÅ@-----------------------------
     evtime = Timer
     If (iPltMax <= 3) Then
             imachi = 20 * 60 - 1
       Else: imachi = 15 * 60 - 1
     End If
     Do
       DoEvents
       fintime = Timer2func     ' 2009.8.17
'       fintime = Timer
       difft = diffTime(fintime, evtime)
       If (difft < imachi) Then
          If (Int(difft) <> Int(diffTold)) Then
             Label2.Caption = " ó‚ãpèIóπÇ‹Ç≈ " + Format(Int((imachi - difft) / 60), "0ï™") + Format(Int((imachi - difft)) Mod 60, " 0ïb")
              diffTold = difft
          End If
       Else
          Exit Do              'Å@éûä‘ë“ÇøèIóπ
       End If
     Loop

Unload Me
'
End Sub
