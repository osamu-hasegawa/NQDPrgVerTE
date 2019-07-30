VERSION 5.00
Begin VB.Form MplVbSmp 
   Caption         =   "MPL»İÌßÙÌßÛ¸ß×Ñ"
   ClientHeight    =   4896
   ClientLeft      =   2940
   ClientTop       =   3000
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'Z µ°ÀŞ°
   ScaleHeight     =   4896
   ScaleWidth      =   7320
   Begin VB.CheckBox Check1 
      Caption         =   "ƒT[ƒ{ ON/OFF"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   24
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reset"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   23
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ˆÊ’u/‘¬“x"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   22
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Close_Command 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Open_Command 
      Caption         =   "Open"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton End_Command 
      Caption         =   "I—¹"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Stop_Command 
      Caption         =   "‚r‚”‚‚"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Reset_Command 
      Caption         =   "‚q‚…‚“‚…‚”"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Ccw_Index_Command 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Demo_Command 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2100
      TabIndex        =   12
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Cw_Index_Command 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Rtn_Command 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2100
      TabIndex        =   10
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Org_Command 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2100
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Ccw_Scan_Command 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Cw_Scan_Command 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton DataSet_Command 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Message_Label 
      Alignment       =   2  '’†‰›‘µ‚¦
      BorderStyle     =   1  'Àü
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Addr_Label 
      Alignment       =   2  '’†‰›‘µ‚¦
      BorderStyle     =   1  'Àü
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "‚w²"
      Height          =   180
      Left            =   1080
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "‚c‚…‚‚"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "‚h‚‚„‚…‚˜ ‚c‚’‚‰‚–‚…"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "‚q‚”‚ ‚c‚’‚‰‚–‚…"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "‚n‚’‚‡ ‚c‚’‚‰‚–‚…"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "‚r‚ƒ‚‚ ‚c‚’‚‰‚–‚…"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Data Set"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "MplVbSmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'
' •Ï”’è‹`
'
'************************************************************
'
'
'Dim Ack As Boolean
'Dim MplData As MPL_S_DATA
'Dim MplResult As MPL_S_RESULT
'Dim hDev As Long
'Dim Status1 As Integer
'Dim cmd As Integer
'Dim Data As Integer
'Dim StopFlag As Integer



'*************************************************************
'
' ƒ{ƒ^ƒ“‚Ìó‘Ô‚ğ‚c‚q‚h‚u‚dó‘Ô‚ÉƒZƒbƒg‚·‚éB
'
'*************************************************************
'
'
Private Sub Btn_Drive_Set()
   DataSet_Command.Enabled = False
   Ccw_Scan_Command.Enabled = False
   Cw_Scan_Command.Enabled = False
   Org_Command.Enabled = False
   Rtn_Command.Enabled = False
   Ccw_Index_Command.Enabled = False
   Cw_Index_Command.Enabled = False
   Demo_Command.Enabled = False
   Reset_Command.Enabled = False
   Close_Command.Enabled = False
   End_Command.Enabled = False
   Stop_Command.Enabled = True
End Sub

'*************************************************************
'
' ƒ{ƒ^ƒ“‚Ìó‘Ô‚ğ‚m‚n ‚c‚q‚h‚u‚dó‘Ô‚ÉƒZƒbƒg‚·‚éB
'
'*************************************************************
'
'
Private Sub Btn_No_Drive_Set()
   DataSet_Command.Enabled = True
   Ccw_Scan_Command.Enabled = True
   Cw_Scan_Command.Enabled = True
   Org_Command.Enabled = True
   Rtn_Command.Enabled = True
   Ccw_Index_Command.Enabled = True
   Cw_Index_Command.Enabled = True
   Demo_Command.Enabled = True
   Reset_Command.Enabled = True
   Close_Command.Enabled = True
   End_Command.Enabled = False
   Stop_Command.Enabled = False
End Sub






'*************************************************************
'
' |‚h‚‚„‚…‚˜ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Ccw_Index_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(-3000, MplData)                    '‚h‚m‚b‚q‚d‚l‚d‚m‚s‚`‚k ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' |‚r‚ƒ‚‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Ccw_Scan_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   cmd = &H13                                      '|‚b‚b‚v ‚r‚b‚`‚m ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub


Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    If Check1(Index).Value = 1 Then
        DioOut 22, 1
    Else
        DioOut 22, 0
    End If
Case 1
    If Check1(Index).Value = 1 Then
        DioOut 23, 1
    Else
        DioOut 23, 0
    End If
Case 2
    If Check1(Index).Value = 1 Then
        DioOut 24, 1
    Else
        DioOut 24, 0
    End If
End Select
End Sub

'*************************************************************
'
' ‚b‚Œ‚‚“‚…ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Close_Command_Click()
   Message_Label.Caption = ""
   DataSet_Command.Enabled = False
   Ccw_Scan_Command.Enabled = False
   Cw_Scan_Command.Enabled = False
   Org_Command.Enabled = False
   Rtn_Command.Enabled = False
   Ccw_Index_Command.Enabled = False
   Cw_Index_Command.Enabled = False
   Demo_Command.Enabled = False
   Reset_Command.Enabled = False
   Close_Command.Enabled = False
   End_Command.Enabled = True
   Stop_Command.Enabled = False
   Open_Command.Enabled = True
   Close_Command.Enabled = False
   Addr_Label.Caption = "‚n‚‚…‚‚µ‚Ä‚­‚¾‚³‚¢B"
   Ack = MPL_BClose(hDev, MplResult)
End Sub


'*************************************************************
'
' {‚h‚‚„‚…‚˜ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Cw_Index_Command_Click()
   Cw_Index Me
   Exit Sub
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(3000, MplData)                  '‚h‚m‚b‚q‚d‚l‚d‚m‚s‚`‚k ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' {‚r‚ƒ‚‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Cw_Scan_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   cmd = &H12                                      '|‚b‚b‚v ‚r‚b‚`‚m ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' ‚c‚‚”‚ ‚r‚…‚”ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub DataSet_Command_Click()
   Message_Label.Caption = ""
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = &H28                                        '‚r‚o‚d‚b ‚h‚m‚h‚s‚`‚k‚h‚y‚d‚P ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = 15                                          '‚q‚`‚s‚d ‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveData2(hDev, Data, MplResult)
   Ack = MPL_BWDriveData3(hDev, Data, MplResult)
   cmd = &H6
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(1000, MplData)                     '‚k‚r‚o‚c ‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWDrive(hDev, &H7, MplData, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(5000, MplData)                     ' ‚g‚r‚o‚c ‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWDrive(hDev, &H8, MplData, MplResult)
End Sub

'*************************************************************
'
' ‚c‚…‚‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Demo_Command_Click()
   Dim Count
   Dim hPort As Long
   
   Message_Label.Caption = ""
   StopFlag = 0
   Ack = MPL_BPortOpen(MplVbSmp.hWnd, 0, hPort, MplResult)
   Call Btn_Drive_Set
   
   Data = &H0                                                  '‚n‚t‚s‚O|‚n‚t‚s‚R@‚g‚h‚f‚go—Í
   Ack = MPL_BPortOut(hPort, Data, MplResult)
   Do                                                          '‚h‚m‚O‚ª‚k‚n‚v‚É‚È‚é‚Ü‚Å‘Ò‚Â
      Ack = MPL_BPortIn(hPort, Data, MplResult)
      DoEvents
      If StopFlag <> 0 Then
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
         Exit Sub
      End If
   Loop While (Data And &H10) = 0

   
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = &H3                                                  '‚n‚q‚h‚f‚h‚m ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1E
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   If (Status1 And &HE0) <> 0 Then
      Call Drive_Stop_Disp
      Call Btn_No_Drive_Set
      Ack = MPL_BPortClose(hPort, MplResult)                   '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
      Exit Sub
   End If

   Do                                                          '‚h‚m‚O‚ª‚g‚h‚f‚g‚É‚È‚é‚Ü‚Å‘Ò‚Â
      Ack = MPL_BPortIn(hPort, Data, MplResult)
      DoEvents
      If StopFlag <> 0 Then
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
         Exit Sub
      End If
   Loop While (Data And &H10) <> 0
         
   Ack = MPL_BPortIn(hPort, Data, MplResult)
   Data = Data Or &H1
   Ack = MPL_BPortOut(hPort, Data, MplResult)
   
   Call MplDataSet(0, MplData)                                 '‚`‚c‚c‚q‚d‚r‚r ‚h‚m‚h‚s‚`‚k‚h‚y‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWDrive(hDev, &H3, MplData, MplResult)
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)

   Call MplDataSet(0, MplData)                                 '‚b‚n‚t‚m‚s‚d‚q ‚o‚q‚d‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWCounter(hDev, &H0, MplData, MplResult)

   For Count = 1 To 4
      Call MplDataSet(4000, MplData)                           '‚h‚m‚b‚q‚d‚l‚d‚m‚s‚`‚k ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
      Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
      Call Ready_Wait
      If (Status1 And &HE0) <> 0 Then
         Call Drive_Stop_Disp
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
         Exit Sub
      End If
   Next

   Call MplDataSet(30000, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)           '‚`‚a‚r‚n‚k‚t‚s‚d ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Call Ready_Wait
   If (Status1 And &HE0) <> 0 Then
      Call Drive_Stop_Disp
      Call Btn_No_Drive_Set
      Ack = MPL_BPortClose(hPort, MplResult)                   '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
      Exit Sub
   End If
      
   Call MplDataSet(0, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)           '‚`‚a‚r‚n‚k‚t‚s‚d ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Call Ready_Wait
   Call Drive_Stop_Disp
   Ack = MPL_BPortClose(hPort, MplResult)                      '‚o‚n‚q‚s ‚g‚`‚m‚c‚k‚d@‚b‚k‚n‚r‚d
   Call Btn_No_Drive_Set
End Sub



'*************************************************************
'
' ‚l‚b‚b‚O‚T‚ª‚q‚d‚`‚c‚xó‘Ô‚É‚È‚é‚Ü‚Å‘Ò‚ÂB
'
'*************************************************************
'
'
Private Sub Ready_Wait()
   Do
      DoEvents
      Ack = MPL_IRDrive(hDev, MplData, MplResult)     'Œ»İˆÊ’u‚`‚c‚c‚q‚d‚r‚r‚Ì•\¦
      Addr_Label.Caption = MplDataGet(MplData)
      Ack = MPL_BRStatus1(hDev, Status1, MplResult)
   Loop While (Status1 And &H1) <> 0
   Ack = MPL_IRDrive(hDev, MplData, MplResult)         'Œ»İˆÊ’u‚`‚c‚c‚q‚d‚r‚r‚Ì•\¦
   Addr_Label.Caption = MplDataGet(MplData)
 End Sub



'*****************************************************************************************************
'
' ‚c‚q‚h‚u‚dI—¹^’â~—vˆö‚Ì•\¦
'
'*****************************************************************************************************
'
'
Private Sub Drive_Stop_Disp()
   Ack = MPL_BRStatus1(hDev, Status1, MplResult)
   If (Status1 And &H20) <> 0 Then
      Message_Label.Caption = "LIMIT‚ª“ü—Í‚³‚ê‚Ü‚µ‚½B"
   ElseIf (Status1 And &H80) <> 0 Then
      Message_Label.Caption = "FS STOP‚ª“ü—Í‚³‚ê‚Ü‚µ‚½B"
   ElseIf (Status1 And &H40) <> 0 Then
      Message_Label.Caption = "SL STOP‚ª“ü—Í‚³‚ê‚Ü‚µ‚½B"
   Else
      Message_Label.Caption = "DRIVE‚ªI—¹‚µ‚Ü‚µ‚½B"
   End If
End Sub



'*************************************************************
'
' I—¹ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub End_Command_Click()
   'End
   Unload Me
End Sub

Private Sub Form_Load()
  DispCenter Me
  Addr_Label.Caption = "‚n‚‚…‚‚µ‚Ä‚­‚¾‚³‚¢B"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Ack = MPL_BClose(hDev, MplResult)
End Sub



'*************************************************************
'
' ‚n‚‚…‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Open_Command_Click()
  C870Open
  
   Message_Label.Caption = ""
'   Ack = MPL_BOpen(MplVbSmp.hWnd, 0, MPL_X, hDev, MplResult)  'ƒfƒoƒCƒXƒI[ƒvƒ“
'   Ack = MPL_IRDrive(hDev, MplData, MplResult)           'Œ»İˆÊ’u‚`‚c‚c‚q‚d‚r‚r‚Ì•\¦
   Addr_Label.Caption = MplDataGet(MplData)

   DataSet_Command.Enabled = True
   Ccw_Scan_Command.Enabled = True
   Cw_Scan_Command.Enabled = True
   Org_Command.Enabled = True
   Rtn_Command.Enabled = True
   Ccw_Index_Command.Enabled = True
   Cw_Index_Command.Enabled = True
   Demo_Command.Enabled = True
   Reset_Command.Enabled = True
   Close_Command.Enabled = True
   End_Command.Enabled = False
   Stop_Command.Enabled = False
   Open_Command.Enabled = False
End Sub

'*************************************************************
'
' ‚n‚’‚‡ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Org_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Data = &H4                              '‚n‚q‚h‚f‚h‚m ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1E
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' ‚q‚…‚“‚…‚”ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Reset_Command_Click()
   Message_Label.Caption = ""
   Call MplDataSet(0, MplData)                        '‚`‚c‚c‚q‚d‚r‚r ‚h‚m‚h‚s‚`‚k‚h‚y‚d ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Ack = MPL_IWDrive(hDev, &H3, MplData, MplResult)
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)

   Call MplDataSet(0, MplData)                        '‚b‚n‚t‚m‚s‚d‚q ‚o‚q‚d‚r‚d‚s ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_IWCounter(hDev, &H0, MplData, MplResult)

   Ack = MPL_IRDrive(hDev, MplData, MplResult)        'Œ»İˆÊ’u‚`‚c‚c‚q‚d‚r‚r‚Ì•\¦
   Addr_Label.Caption = MplDataGet(MplData)
End Sub


'*************************************************************
'
' ‚q‚”‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Rtn_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Call MplDataSet(0, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)    '‚`‚a‚r‚n‚k‚t‚s‚d ‚h‚m‚c‚d‚w ‚c‚q‚h‚u‚d ‚b‚n‚l‚l‚`‚m‚c
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' ‚r‚”‚‚ƒ{ƒ^ƒ“‚ª‰Ÿ‚³‚ê‚½‚Ìˆ—
'
'*************************************************************
'
'
Private Sub Stop_Command_Click()
   cmd = &HFF                                      '‚d‚l‚r‚s‚n‚o ‚b‚n‚l‚l‚`‚m‚c
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   StopFlag = 1
End Sub

