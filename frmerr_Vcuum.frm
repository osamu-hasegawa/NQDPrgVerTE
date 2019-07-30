VERSION 5.00
Begin VB.Form frmerr_Vcuum 
   Caption         =   "アラーム"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8310
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "画面消去"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   576
      Left            =   3240
      TabIndex        =   0
      Top             =   4560
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "アラーム表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   25.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   516
      Index           =   1
      Left            =   2688
      TabIndex        =   2
      Top             =   2040
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "アラーム表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   324
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1836
   End
End
Attribute VB_Name = "frmerr_Vcuum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  DispCenter Me
  SetData
End Sub
Private Sub SetData()
  DispErr
End Sub

Public Sub DispErr()
  'ErrNo = 12: gErrMsg$(EmgArm, ErrNo) = "真空未到達"
  Label1(1).Caption = gErrMsg$(1, 12)   '真空未到達

End Sub
