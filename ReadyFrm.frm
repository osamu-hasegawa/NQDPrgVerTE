VERSION 5.00
Begin VB.Form ReadyFrm 
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10476
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10476
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer Timer1 
      Left            =   576
      Top             =   504
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "事故復旧後、"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   720
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "System not ready"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   25.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   516
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   3828
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異常リセットを押してください"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   720
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   8484
   End
End
Attribute VB_Name = "ReadyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lViewFlg      '前の画面番号
Private Sub Form_Load()
  DispCenter Me
  lViewFlg = ViewFlg      '前の画面番号
  'ViewFlg = 2             '画面番号
  FrmMenuFlg = True                   'メニューから抜けるときfalse
  Timer1.Interval = 100
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Dim ms$
  Timer1.Enabled = False
  Do
    DoEvents
    If SystemReadyChk() = 1 Then Exit Do
    Label2.Caption = MsgChk
  Loop
  '
  Select Case lViewFlg
  Case 1
    PGM_Menu.Show
    Unload Me
  Case 2 '成形（シングル）
    NQD70_SC.Show
    Unload Me
  Case 3  '成形（テスト）
    LS21_TC.Show
    Unload Me
  Case 4  'I O チェック
    IOChk.Show
  Case 5  'スケール変更
    LS21_GphScale.Show
  Case 6  '読み出し
  Case 7  'メモ帳
  Case 8  'edit
    MYEdit.Show
    Unload Me
  Case Else
    PGM_Menu.Show
    Unload Me
  End Select
  
End Sub
