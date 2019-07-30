VERSION 5.00
Begin VB.Form ADSampling 
   Caption         =   "Acquire One Sample"
   ClientHeight    =   2775
   ClientLeft      =   1140
   ClientTop       =   1770
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'Z µ°ÀÞ°
   ScaleHeight     =   2775
   ScaleWidth      =   5010
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
      TabIndex        =   7
      Text            =   "2"
      Top             =   1440
      Width           =   615
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
      TabIndex        =   2
      Text            =   "1"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
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
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Channel"
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
      TabIndex        =   6
      Top             =   1485
      Width           =   1815
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
   Begin VB.Label Label1 
      Caption         =   "Acquire one sample."
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
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "ADSampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim nRet As Long
    Dim SmplChInf(0 To 1) As ADSMPLCHREQ
    Dim bSmpData() As Byte
    Dim wSmpData() As Integer
    Dim dwSmpData() As Long
    Dim szDisp As String
    Dim ulCh As Long

    ' Retrieve a channel number
    If IsNull(Ch(0).Text) Then
        nRet = MsgBox("Invalid channel", (vbOKOnly + vbCritical), "Error Code")
        Exit Sub
    End If
    
    ghChannel(0) = Val(Ch(0).Text)
    ghChannel(1) = Val(Ch(1).Text)
    
    SmplChInf(0).ulChNo = ghChannel(0)
    SmplChInf(0).ulRange = gConfig.SmplChReq(0).ulRange

    If ghChannel(1) = 0 Then
        ulCh = 1
    Else
        ulCh = 2
        SmplChInf(1).ulChNo = ghChannel(1)
        SmplChInf(1).ulRange = gConfig.SmplChReq(0).ulRange
    End If

    If gInfo.ulResolution <= 8 Then
        ReDim bSmpData(ulCh)
        
        nRet = AdInputAD(ghDeviceHandle, ulCh, gConfig.ulSingleDiff, SmplChInf(0), bSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(bSmpData(0)) & "h"
            
            If ulCh = 2 Then
                szDisp = szDisp & ", " & Hex(bSmpData(1)) & "h"
            End If
        End If
    ElseIf gInfo.ulResolution > 8 And gInfo.ulResolution <= 16 Then
        ReDim wSmpData(ulCh)
        
        nRet = AdInputAD(ghDeviceHandle, ulCh, gConfig.ulSingleDiff, SmplChInf(0), wSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(wSmpData(0)) & "h"
            
            If ulCh = 2 Then
                szDisp = szDisp & ", " & Hex(wSmpData(1)) & "h"
            End If
        End If
    ElseIf gInfo.ulResolution > 16 Then
        ReDim dwSmpData(ulCh)

        nRet = AdInputAD(ghDeviceHandle, ulCh, gConfig.ulSingleDiff, SmplChInf(0), dwSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(dwSmpData(0)) & "h"
            
            If ulCh = 2 Then
                szDisp = szDisp & ", " & Hex(dwSmpData(1)) & "h"
            End If
        End If
    End If
    
    If nRet = AD_ERROR_SUCCESS Then
        nRet = MsgBox(szDisp, (vbOKOnly + vbInformation), "Sampling Data")
    Else
        Call DsplyErrMessage(nRet)
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


