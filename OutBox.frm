VERSION 5.00
Begin VB.Form OutBox 
   Caption         =   "Specify output contacts"
   ClientHeight    =   2556
   ClientLeft      =   1476
   ClientTop       =   1980
   ClientWidth     =   3744
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'Z µ°ÀÞ°
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   Begin VB.TextBox IDC_BUFFER 
      Height          =   372
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   1200
      Width           =   1572
   End
   Begin VB.CommandButton IDC_CANCEL 
      Caption         =   "CANCEL"
      Height          =   372
      Left            =   2106
      TabIndex        =   3
      Top             =   1920
      Width           =   1092
   End
   Begin VB.CommandButton IDC_OK 
      Caption         =   "OK"
      Height          =   372
      Left            =   546
      TabIndex        =   2
      Top             =   1920
      Width           =   1092
   End
   Begin VB.TextBox IDC_STARTNUM 
      Height          =   372
      Left            =   2040
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label3 
      Caption         =   "(0:OFF 1:ON)"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label2 
      Caption         =   "data to output"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Contact number to output"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "OutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
  DispCenter Me
End Sub

Private Sub IDC_CANCEL_Click()
    Unload Me
End Sub

Private Sub IDC_OK_Click()
    Dim lpszName As String
    Dim dwStartNum As Long
    Dim nRet As Long
    Dim nBuffer As Long
    
    dwStartNum = Val(IDC_STARTNUM.Text)
    nBuffer = Val(IDC_BUFFER.Text)
    lpszName = "FBIDIO1" & Chr(0)
    hDeviceHandle = DioOpen(lpszName, FBIDIO_FLAG_SHARE)

    If hDeviceHandle = &HFFFF Then
        MsgBox ("Opening the board failed.")
        Exit Sub
    End If

    nRet = DioOutputPoint(hDeviceHandle, nBuffer, dwStartNum, 1)
    If nRet <> 0 Then
        MsgBox ("Output the data failed.")
        nRet = DioClose(hDeviceHandle)
        Exit Sub
    End If
        
    nRet = DioClose(hDeviceHandle)
    If nRet <> 0 Then
        MsgBox ("Closing the board failed.")
        Exit Sub
    End If
   

End Sub

