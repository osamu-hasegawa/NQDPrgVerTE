VERSION 5.00
Begin VB.Form MplVbSmp 
   Caption         =   "MPLｻﾝﾌﾟﾙﾌﾟﾛｸﾟﾗﾑ"
   ClientHeight    =   4896
   ClientLeft      =   2940
   ClientTop       =   3000
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4896
   ScaleWidth      =   7320
   Begin VB.CheckBox Check1 
      Caption         =   "サーボ ON/OFF"
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
      Caption         =   "位置/速度"
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
      Caption         =   "終了"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Stop_Command 
      Caption         =   "Ｓｔｏｐ"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Reset_Command 
      Caption         =   "Ｒｅｓｅｔ"
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
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Addr_Label 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  '中央揃え
      Caption         =   "Ｘ軸"
      Height          =   180
      Left            =   1080
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Ｄｅｍｏ"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Ｉｎｄｅｘ Ｄｒｉｖｅ"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ｒｔｎ Ｄｒｉｖｅ"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Ｏｒｇ Ｄｒｉｖｅ"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Ｓｃａｎ Ｄｒｉｖｅ"
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
' 変数定義
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
' ボタンの状態をＤＲＩＶＥ状態にセットする。
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
' ボタンの状態をＮＯ ＤＲＩＶＥ状態にセットする。
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
' −Ｉｎｄｅｘボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Ccw_Index_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(-3000, MplData)                    'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' −Ｓｃａｎボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Ccw_Scan_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   cmd = &H13                                      '−ＣＣＷ ＳＣＡＮ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
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
' Ｃｌｏｓｅボタンが押された時の処理
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
   Addr_Label.Caption = "Ｏｐｅｎしてください。"
   Ack = MPL_BClose(hDev, MplResult)
End Sub


'*************************************************************
'
' ＋Ｉｎｄｅｘボタンが押された時の処理
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
   Call MplDataSet(3000, MplData)                  'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' ＋Ｓｃａｎボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Cw_Scan_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   cmd = &H12                                      '−ＣＣＷ ＳＣＡＮ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' Ｄａｔａ Ｓｅｔボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub DataSet_Command_Click()
   Message_Label.Caption = ""
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = &H28                                        'ＳＰＥＣ ＩＮＩＴＡＬＩＺＥ１ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = 15                                          'ＲＡＴＥ ＳＥＴ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveData2(hDev, Data, MplResult)
   Ack = MPL_BWDriveData3(hDev, Data, MplResult)
   cmd = &H6
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(1000, MplData)                     'ＬＳＰＤ ＳＥＴ ＣＯＭＭＡＮＤ
   Ack = MPL_IWDrive(hDev, &H7, MplData, MplResult)

   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Call MplDataSet(5000, MplData)                     ' ＨＳＰＤ ＳＥＴ ＣＯＭＭＡＮＤ
   Ack = MPL_IWDrive(hDev, &H8, MplData, MplResult)
End Sub

'*************************************************************
'
' Ｄｅｍｏボタンが押された時の処理
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
   
   Data = &H0                                                  'ＯＵＴ０−ＯＵＴ３　ＨＩＧＨ出力
   Ack = MPL_BPortOut(hPort, Data, MplResult)
   Do                                                          'ＩＮ０がＬＯＷになるまで待つ
      Ack = MPL_BPortIn(hPort, Data, MplResult)
      DoEvents
      If StopFlag <> 0 Then
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
         Exit Sub
      End If
   Loop While (Data And &H10) = 0

   
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Data = &H3                                                  'ＯＲＩＧＩＮ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1E
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   If (Status1 And &HE0) <> 0 Then
      Call Drive_Stop_Disp
      Call Btn_No_Drive_Set
      Ack = MPL_BPortClose(hPort, MplResult)                   'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
      Exit Sub
   End If

   Do                                                          'ＩＮ０がＨＩＧＨになるまで待つ
      Ack = MPL_BPortIn(hPort, Data, MplResult)
      DoEvents
      If StopFlag <> 0 Then
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
         Exit Sub
      End If
   Loop While (Data And &H10) <> 0
         
   Ack = MPL_BPortIn(hPort, Data, MplResult)
   Data = Data Or &H1
   Ack = MPL_BPortOut(hPort, Data, MplResult)
   
   Call MplDataSet(0, MplData)                                 'ＡＤＤＲＥＳＳ ＩＮＩＴＡＬＩＺＥ ＣＯＭＭＡＮＤ
   Ack = MPL_IWDrive(hDev, &H3, MplData, MplResult)
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)

   Call MplDataSet(0, MplData)                                 'ＣＯＵＮＴＥＲ ＰＲＥＳＥＴ ＣＯＭＭＡＮＤ
   Ack = MPL_IWCounter(hDev, &H0, MplData, MplResult)

   For Count = 1 To 4
      Call MplDataSet(4000, MplData)                           'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
      Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
      Call Ready_Wait
      If (Status1 And &HE0) <> 0 Then
         Call Drive_Stop_Disp
         Call Btn_No_Drive_Set
         Ack = MPL_BPortClose(hPort, MplResult)                'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
         Exit Sub
      End If
   Next

   Call MplDataSet(30000, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)           'ＡＢＳＯＬＵＴＥ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Call Ready_Wait
   If (Status1 And &HE0) <> 0 Then
      Call Drive_Stop_Disp
      Call Btn_No_Drive_Set
      Ack = MPL_BPortClose(hPort, MplResult)                   'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
      Exit Sub
   End If
      
   Call MplDataSet(0, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)           'ＡＢＳＯＬＵＴＥ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Call Ready_Wait
   Call Drive_Stop_Disp
   Ack = MPL_BPortClose(hPort, MplResult)                      'ＰＯＲＴ ＨＡＮＤＬＥ　ＣＬＯＳＥ
   Call Btn_No_Drive_Set
End Sub



'*************************************************************
'
' ＭＣＣ０５がＲＥＡＤＹ状態になるまで待つ。
'
'*************************************************************
'
'
Private Sub Ready_Wait()
   Do
      DoEvents
      Ack = MPL_IRDrive(hDev, MplData, MplResult)     '現在位置ＡＤＤＲＥＳＳの表示
      Addr_Label.Caption = MplDataGet(MplData)
      Ack = MPL_BRStatus1(hDev, Status1, MplResult)
   Loop While (Status1 And &H1) <> 0
   Ack = MPL_IRDrive(hDev, MplData, MplResult)         '現在位置ＡＤＤＲＥＳＳの表示
   Addr_Label.Caption = MplDataGet(MplData)
 End Sub



'*****************************************************************************************************
'
' ＤＲＩＶＥ終了／停止要因の表示
'
'*****************************************************************************************************
'
'
Private Sub Drive_Stop_Disp()
   Ack = MPL_BRStatus1(hDev, Status1, MplResult)
   If (Status1 And &H20) <> 0 Then
      Message_Label.Caption = "LIMITが入力されました。"
   ElseIf (Status1 And &H80) <> 0 Then
      Message_Label.Caption = "FS STOPが入力されました。"
   ElseIf (Status1 And &H40) <> 0 Then
      Message_Label.Caption = "SL STOPが入力されました。"
   Else
      Message_Label.Caption = "DRIVEが終了しました。"
   End If
End Sub



'*************************************************************
'
' 終了ボタンが押された時の処理
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
  Addr_Label.Caption = "Ｏｐｅｎしてください。"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Ack = MPL_BClose(hDev, MplResult)
End Sub



'*************************************************************
'
' Ｏｐｅｎボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Open_Command_Click()
  C870Open
  
   Message_Label.Caption = ""
'   Ack = MPL_BOpen(MplVbSmp.hWnd, 0, MPL_X, hDev, MplResult)  'デバイスオープン
'   Ack = MPL_IRDrive(hDev, MplData, MplResult)           '現在位置ＡＤＤＲＥＳＳの表示
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
' Ｏｒｇボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Org_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Data = &H4                              'ＯＲＩＧＩＮ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveData1(hDev, Data, MplResult)
   cmd = &H1E
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' Ｒｅｓｅｔボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Reset_Command_Click()
   Message_Label.Caption = ""
   Call MplDataSet(0, MplData)                        'ＡＤＤＲＥＳＳ ＩＮＩＴＡＬＩＺＥ ＣＯＭＭＡＮＤ
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)
   Ack = MPL_IWDrive(hDev, &H3, MplData, MplResult)
   Ack = MPL_BWaitDriveCommand(hDev, 0, MplResult)

   Call MplDataSet(0, MplData)                        'ＣＯＵＮＴＥＲ ＰＲＥＳＥＴ ＣＯＭＭＡＮＤ
   Ack = MPL_IWCounter(hDev, &H0, MplData, MplResult)

   Ack = MPL_IRDrive(hDev, MplData, MplResult)        '現在位置ＡＤＤＲＥＳＳの表示
   Addr_Label.Caption = MplDataGet(MplData)
End Sub


'*************************************************************
'
' Ｒｔｎボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Rtn_Command_Click()
   Message_Label.Caption = ""
   Call Btn_Drive_Set
   Call MplDataSet(0, MplData)
   Ack = MPL_IWDrive(hDev, &H15, MplData, MplResult)    'ＡＢＳＯＬＵＴＥ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Call Ready_Wait
   Call Drive_Stop_Disp
   Call Btn_No_Drive_Set
End Sub

'*************************************************************
'
' Ｓｔｏｐボタンが押された時の処理
'
'*************************************************************
'
'
Private Sub Stop_Command_Click()
   cmd = &HFF                                      'ＥＭＳＴＯＰ ＣＯＭＭＡＮＤ
   Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)
   StopFlag = 1
End Sub

