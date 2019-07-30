VERSION 5.00
Begin VB.Form R_Tch61 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4284
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6276
   LinkTopic       =   "Form1"
   ScaleHeight     =   4284
   ScaleWidth      =   6276
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ƒLƒƒƒ“ƒZƒ‹"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   10.8
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
      Top             =   90
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I—¹"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   10.8
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
      Top             =   90
      Width           =   1236
   End
   Begin VB.Label Label2 
      Alignment       =   1  '‰E‘µ‚¦
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3180
      TabIndex        =   5
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  '‰E‘µ‚¦
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3180
      TabIndex        =   4
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '“§–¾
      Caption         =   "ˆÊ’ui‰ºjF"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '“§–¾
      Caption         =   "ˆÊ’uiãjF"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   885
      Width           =   1545
   End
End
Attribute VB_Name = "R_Tch61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' R_Tch61@ÌßÛ¸Ş×Ñ
   
    'update 2002.8.10   ROZ( ) “Ë‚«“–‚Ä¬Œ`—pÊß×Ò°À‚Ö•ÏX



Dim lViewFlg      '‘O‚Ì‰æ–Ê”Ô†

Private Sub Command2_Click(Index As Integer)
  Select Case Index
  Case 0  'ƒLƒƒƒ“ƒZƒ‹
    
  Case 1  'I—¹
    Unload Me
    PGM_Menu.Show
  
  End Select
End Sub

Private Sub Form_Load()
  DispCenter Me
  Timer1.Enabled = False
  lViewFlg = ViewFlg      '‘O‚Ì‰æ–Ê”Ô†
  ViewFlg = 9             '‰æ–Ê”Ô†
  FrmMenuFlg = True       'ƒƒjƒ…[‚©‚ç”²‚¯‚é‚Æ‚«false
  Me.Show
  SetData
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim up%, dw%, md%, ve%
Dim pos!
  DioInput 10, dw
  DioInput 9, up
  DioInput 11, md
  DioInput 12, ve
  pos = r_z()
  If up = 1 Then
    Label2(0).Caption = Format(pos, "0.000")
'    If md = 1 Then roz(0) = pos
  End If
  If dw = 1 Then
    Label2(1).Caption = Format(pos, "0.000")
'    If md = 1 Then roz(1) = pos
  End If
  
  
End Sub

Private Sub SetData()
  ServoON
  C870OrgVelSet   '/* Œ´“_—p‘¬“xİ’è */
  Label2(4).Caption = "Œ´“_o‚µÀs"
  genten
  Ready_Wait
  Label2(4).Caption = "Œ´“_o‚µŠ®—¹"
  C870ManVelSet   '/* è“®—p‘¬“xİ’è */
  '/* ƒJƒEƒ“ƒ^‚Éƒ[ƒ‚ğ‘‚«‚Ş */
  C870AdrInit       '‚`‚c‚c‚q‚d‚r‚r ‚h‚m‚h‚s‚`‚k‚h‚y‚d ‚b‚n‚l‚l‚`‚m‚c
  C870CntPreSet 0   '‚b‚n‚t‚m‚s‚d‚q ‚o‚q‚d‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
  
End Sub
Private Sub genten()
'--------------
  C870Genten
End Sub

