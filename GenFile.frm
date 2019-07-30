VERSION 5.00
Begin VB.Form GenFile 
   Appearance      =   0  'Ì×¯Ä
   BorderStyle     =   3  'ŒÅ’èÀŞ²±Û¸Ş
   Caption         =   "ƒtƒ@ƒCƒ‹"
   ClientHeight    =   5652
   ClientLeft      =   4320
   ClientTop       =   3612
   ClientWidth     =   5712
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ°ÀŞ°
   ScaleHeight     =   5652
   ScaleWidth      =   5712
   Begin VB.CommandButton Command1 
      Caption         =   "‘o"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4968
      TabIndex        =   10
      Top             =   5040
      Width           =   552
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2940
      TabIndex        =   6
      Text            =   "c:\"
      Top             =   1500
      Width           =   2640
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3744
      Left            =   108
      TabIndex        =   2
      Top             =   1008
      Width           =   2670
   End
   Begin VB.CommandButton btnCan 
      Caption         =   "æÁ(&C)"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3324
      TabIndex        =   9
      Top             =   4992
      Width           =   1275
   End
   Begin VB.CommandButton btnExe 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1092
      TabIndex        =   8
      Top             =   5016
      Width           =   1275
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   105
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   630
      Width           =   2664
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2808
      Left            =   2952
      TabIndex        =   7
      Top             =   1944
      Width           =   2640
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2940
      TabIndex        =   4
      Top             =   648
      Width           =   2640
   End
   Begin VB.Label Label3 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "‘I‘ğƒfƒBƒŒƒNƒgƒŠ(&P)"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2952
      TabIndex        =   5
      Top             =   1260
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "‘I‘ğƒhƒ‰ƒCƒu(&D)"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2940
      TabIndex        =   3
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "‘I‘ğƒtƒ@ƒCƒ‹(&F)"
      BeginProperty Font 
         Name            =   "‚l‚r ‚o–¾’©"
         Size            =   10.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   108
      TabIndex        =   0
      Top             =   240
      Width           =   2664
   End
End
Attribute VB_Name = "GenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ControlSetData
  cap   As String
  btn1  As String
  btn2  As String
End Type

Private Type FileSetData
  flag  As Integer
  path  As String
  pttn  As String
  file  As String
  full  As String
  attr  As Byte
  stmp  As Double
  size  As Long
End Type

Private m_ctrls As ControlSetData
Private m_File  As FileSetData

Public Function chkFile(chkDir As String, chkName As String, chkSize As Long, chkFlg As Integer) As Integer
  Dim i As Integer, ck As Integer, fnum As Integer, flg As Integer
  Dim dummy As String, Dt1 As String, dt2 As String
  
  ck = 0
  chkDir = Trim$(chkDir)
  If Right$(chkDir, 1) <> "\" Then chkDir = chkDir & "\"
  
  Dt1 = ""
  For i = 1 To Len(chkName)
    dt2 = Mid$(chkName, i, 1)
    'If dt2 <> " " Then dt1 = dt1 + dt2
    Dt1 = Dt1 + dt2
  Next i
  chkName = Dt1
  
  i = InStr(chkName, ".")
  If i = 0 Then
    Dt1 = chkName
    dt2 = ""
  Else
    Dt1 = Left$(chkName, i - 1)
    dt2 = Right$(chkName, Len(chkName) - i)
  End If
  
  'If Len(dt1) > 8 Then ck = 11
  If Dt1 = "" Then
    chkName = "$$$Dummy.tes"
    If chkFlg <> cChk Then ck = 12
  End If
  
  dummy = ",.*\?$@%"
  For i = 1 To Len(dummy)
    If InStr(Dt1, Mid$(dummy, i, 1)) <> 0 Then ck = 13
    If InStr(dt2, Mid$(dummy, i, 1)) <> 0 Then ck = 13
  Next i
  If InStr(Dt1, Chr$(&H22)) <> 0 Then ck = 13
  If InStr(dt2, Chr$(&H22)) <> 0 Then ck = 13
      
  fnum = FreeFile
  If ck = 0 Then
    dummy = String$(chkSize, "A")

    On Error GoTo chkFileErr
    Open chkDir & "$$$Dummy.tes" For Output As fnum
      Print #fnum, dummy$
    Close fnum
    Kill chkDir & "$$$Dummy.tes"
    Open chkDir & chkName For Input As fnum
    If chkFlg = cSave Then ck = 6
  End If

ChkFileRet:
  On Error GoTo 0
  Close fnum
  If ck <> 0 Then
    i = 48
    Select Case ck
      Case 1: dummy = "‚c‚h‚r‚j‚ª‚r‚d‚s‚³‚ê‚Ä‚Ü‚¹‚ñ"
      Case 2: dummy = "‚c‚h‚r‚j‚ª–¢‚e‚n‚q‚l‚`‚s‚Å‚·"
      Case 3: dummy = "‚c‚h‚r‚j‚Ì—e—Ê‚ª‘«‚è‚Ü‚¹‚ñ"
      Case 4: dummy = "‚c‚h‚r‚j‚ª‘‚İ‹Ö~‚Å‚·"
      Case 5: dummy = "ƒtƒ@ƒCƒ‹‚ª‚ ‚è‚Ü‚¹‚ñ"
      Case 6: dummy = "“¯ˆêƒtƒ@ƒCƒ‹‚ª—L‚è‚Ü‚·" & Chr$(13)
              dummy = dummy & chkDir & chkName & Chr$(13)
              dummy = dummy & "‚ğã‘‚µ‚Ü‚·‚©H"
              i = i + 4
      Case 7: dummy = "ŠY“–ƒfƒBƒŒƒNƒgƒŠ‚ª‚ ‚è‚Ü‚¹‚ñ"
      Case 8: dummy = "ŠY“–ƒhƒ‰ƒCƒu‚ª‚ ‚è‚Ü‚¹‚ñ"
      Case 11: dummy = "ƒtƒ@ƒCƒ‹–¼‚ª’·‚·‚¬‚Ü‚·"
      Case 12: dummy = "ƒtƒ@ƒCƒ‹–¼‚ª NULL‚Å‚·"
      Case 13: dummy = "ƒtƒ@ƒCƒ‹–¼‚É‹Ö~•¶š‚ªŠÜ‚Ü‚ê‚Ä‚¢‚Ü‚·"
    End Select
    flg = MsgBox(dummy, i, "‚c‚‰‚“‚‹ƒGƒ‰[")
    If ck = 6 And flg = 6 Then ck = 0
  End If
  chkFile = ck
Exit Function

chkFileErr:
  Select Case Err
    Case 71: ck = 1
    Case 57: ck = 2
    Case 61: ck = 3
    Case 70: ck = 4
    Case 53
      If (chkFlg = cLoad) Or (chkFlg = cDel) Then ck = 5
    Case 76
      If chkFlg = cSave Then
        'dt1 = Left$(chkDir, 2)
        'dt2 = Right$(chkDir, Len(chkDir) - 3)
        'Do Until drvd2$ = ""
        '  d$ = GenGetToken$(drvd2$, "\")
        '  drvd1$ = drvd1$ + "\" + d$
        '  MkDir drvd1$
        'Loop
        ck = 7
      Else
        ck = 7
      End If
    Case 68: ck = 8
    Case Else: Stop
  End Select
  Resume ChkFileRet
End Function


Private Sub FillPickList()
  On Error GoTo DiskErr
  
  Drive1.Drive = Left(m_File.path, 2)
  Dir1.path = m_File.path
  File1.Pattern = m_File.pttn
  txtDir.Text = Dir1.path
  txtFile.Text = m_File.file

Exit Sub

DiskErr:
  Select Case Err
  Case 76 'DirNotFound
    m_File.path = "\"
    Resume
  Case Else
    Resume Next
  End Select
  
End Sub

Public Sub GetFile(flg, path, file, Optional full, Optional attr, Optional stmp, Optional size)
  
  flg = m_File.flag
  path = m_File.path
  file = m_File.file
  If Not IsMissing(full) Then
    full = m_File.full
  End If
  If Not IsMissing(attr) Then
    attr = m_File.attr
  End If
  If Not IsMissing(stmp) Then
    stmp = m_File.stmp
  End If
  If Not IsMissing(size) Then
    size = m_File.size
  End If
  
End Sub

Public Sub SetFile(flg, path, Optional file, Optional pttn, Optional attr)
  
  m_File.flag = flg
  m_File.path = Trim(path)
  If Right(m_File.path, 1) <> "\" Then
    m_File.path = m_File.path & "\"
  End If
  If Not IsMissing(pttn) Then
    m_File.pttn = pttn
  End If
  If Not IsMissing(file) Then
    m_File.file = file
  End If
  If Not IsMissing(attr) Then
    m_File.attr = attr
  End If

End Sub

Private Sub btnCan_Click()
  m_File.file = ""
  m_File.flag = False
  Unload Me
End Sub

Private Sub btnExe_Click()
  Dim ck As Integer
  Dim DDir As String
  Dim DName As String
  
  DDir = txtDir.Text
  DName = txtFile.Text
  ck = chkFile(DDir, DName, 1, m_File.flag)
  On Error Resume Next
  If ck = 0 Then
    m_File.path = DDir
    m_File.file = DName
    'm_File.full = DDir & DName
    lblFileFullSet
    m_File.attr = GetAttr(m_File.full)
    m_File.stmp = FileDateTime(m_File.full)
    m_File.size = FileLen(m_File.full)
    m_File.flag = True
  Else
    m_File.flag = False
    m_File.file = ""
  End If
  On Error GoTo 0
  Unload Me
End Sub

Private Sub Command1_Click()
'FileListBox ‚Ìƒtƒ@ƒCƒ‹‚ğƒeƒLƒXƒgFile‚É•Û‘¶
Dim fl$, i%, fnum%
fl$ = App.path & "\file.txt"
  fnum = FreeFile
  Open fl For Output As #fnum
    For i = 0 To File1.ListCount - 1
      Write #fnum, File1.List(i)
    Next i
  Close
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
    txtDir.Text = Dir1.path
    lblFileFullSet
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      'Dir1_Change
       Dir1.path = Dir1.List(Dir1.ListIndex)
       Dir1_Change
    End If
End Sub

Private Sub Drive1_Change()
  Dim Msg As String

  On Error GoTo Drv1CngErr
  Dir1.path = Drive1.Drive
  lblFileFullSet
  
Exit Sub
Drv1CngErr:
  Select Case Err
    Case 68
      Msg = "‘¼‚Ìƒhƒ‰ƒCƒu‚ğ‘I‘ğ‚·‚é‚©AƒfƒBƒXƒN‚ğ“ü‚ê‚Ä‚©‚ç‚à‚¤ˆê“x"
      Msg = Msg + Chr$(13) + "ƒhƒ‰ƒCƒu‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢"
      MsgBox Msg, 48, "ƒhƒ‰ƒCƒu‚ª€”õ‚³‚ê‚Ä‚¢‚Ü‚¹‚ñ"
      Resume Next
    Case Else
      Stop
  End Select

End Sub

Private Sub File1_Click()
    txtFile.Text = File1.FileName
    lblFileFullSet
End Sub

Private Sub File1_DblClick()
  File1_Click
  btnExe_Click
End Sub


Private Sub Form_Initialize()
  
  m_ctrls.cap = ""
  m_ctrls.btn1 = "Às"
  m_ctrls.btn2 = "æÁ"
  
  m_File.path = App.path
  m_File.file = "*.*"
  m_File.pttn = "*.*"
  m_File.full = ""
  m_File.attr = 0
  m_File.stmp = 0
  m_File.size = 0

End Sub

Private Sub Form_Load()
  DispCenter Me
  Caption = m_ctrls.cap
  btnExe.Caption = m_ctrls.btn1
  btnCan.Caption = m_ctrls.btn2
  
  txtFile = m_File.file
  
  FillPickList
  Screen.MousePointer = 0  ' ƒ|ƒCƒ“ƒ^‚ğ»Œv‚É•ÏX.
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Unload GenFile

End Sub

Private Sub lblFileFullSet()
  If Right(File1.path, 1) <> "\" Then
    m_File.full = File1.path & "\" & File1.FileName
  Else
    m_File.full = File1.path & File1.FileName
  End If
End Sub

Private Sub txtDir_GotFocus()
    txtDir.SelStart = 0
    txtDir.SelLength = Len(txtDir.Text)
End Sub

Private Sub txtFile_DblClick()
'
  If InStr(txtFile.Text, "*.") Then
    File1.Pattern = txtFile.Text
  End If
End Sub

Private Sub txtFile_GotFocus()
    txtFile.SelStart = 0
    txtFile.SelLength = Len(txtFile.Text)
End Sub


Public Sub SetCtrl(cap, Optional btn1, Optional btn2)
  
  m_ctrls.cap = cap
  If Not IsMissing(btn1) Then
    m_ctrls.btn1 = btn1
  End If
  If Not IsMissing(btn2) Then
    m_ctrls.btn2 = btn2
  End If

End Sub
Private Sub txtFile_KeyPress(KeyAscii As Integer)

  If KeyAscii <> &HD Then Exit Sub
  If InStr(txtFile.Text, "*.") Then
    File1.Pattern = txtFile.Text
  End If
End Sub
