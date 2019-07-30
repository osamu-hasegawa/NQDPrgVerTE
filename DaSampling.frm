VERSION 5.00
Begin VB.Form DaSampling 
   Caption         =   "One sample output"
   ClientHeight    =   3405
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'Z µ°ÀÞ°
   ScaleHeight     =   3405
   ScaleWidth      =   5190
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Ch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   10
      Text            =   "2"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox DeviceHandle 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Ch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Text            =   "1"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Output data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2085
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "HEX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Output one data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Device handle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Channel number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1485
      Width           =   1935
   End
End
Attribute VB_Name = "DaSampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim nRet As Long
    Dim SmplChInf(0 To 15) As DASMPLCHREQ
    Dim wData(0 To 15) As Integer
    Dim ulCh As Long
    
    ' Retrieve a channel number
    If IsNull(Ch(0).Text) Then
        nRet = MsgBox("Invalid channel", (vbOKOnly + vbCritical), "Error code")
        Exit Sub
    End If
    
    ghChannelDA(0) = Val(Ch(0).Text)
    ghChannelDA(1) = Val(Ch(1).Text)
    
    ' Setup the output conditions.
    SmplChInf(0).ulChNo = ghChannelDA(0)
    SmplChInf(0).ulRange = gConfigDA.SmplChReq(0).ulRange

    ' Configure the output data
    wData(0) = Val("&H" + txtData(0).Text)
    
    If ghChannelDA(1) = 0 Then
        ulCh = 1
    Else
        ulCh = 2
        SmplChInf(1).ulChNo = ghChannelDA(1)
        SmplChInf(1).ulRange = gConfigDA.SmplChReq(0).ulRange
        wData(1) = Val("&H" + txtData(1).Text)
    End If
    
    ' Output one sample
    nRet = DaOutputDA(ghDeviceHandleDa, ulCh, SmplChInf(0), wData(0))
    
    If nRet <> DA_ERROR_SUCCESS Then
        Call DsplyErrMessageDA(nRet)
    Else
        nRet = MsgBox("The DA conversion output is completed successfully. [ DaOutputDA ]", vbInformation)
    End If

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub


