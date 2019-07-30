Attribute VB_Name = "FbiDio"
'  FbiDio .BAS
'
' Declare function prototypes and structures and symbols exported from the FBIDIO.DLL.
'
' Copyright (C) 1998-1999 Interface Corpration

'    update : 2004.2.19   gDebugFlg 追加
'    update : 2008.3.15 水冷却ON,OFF　追加
'
'
'
' -----------------------------------------------------------------------
'       Symbols and/or identifiers
' -----------------------------------------------------------------------
Public hDeviceHandle As Long       ' Device handle
'
Public Const FBIDIO_FLAG_SHARE = &H2            'Flag available to the DioOpen function. This flag shows that the device is opened as shareable.

Public Const FBIDIO_IN1_8 = 0                   'Read data from IN1 to IN8.
Public Const FBIDIO_IN9_16 = 1                  'Read data from IN9 to IN16.
Public Const FBIDIO_IN17_24 = 2                 'Read data from IN17 to IN24.
Public Const FBIDIO_IN25_32 = 3                 'Read data from IN25 to IN32.
Public Const FBIDIO_IN33_40 = 4                 'Read data from IN33 to IN40.
Public Const FBIDIO_IN41_48 = 5                 'Read data from IN41 to IN48.
Public Const FBIDIO_IN49_56 = 6                 'Read data from IN49 to IN56.
Public Const FBIDIO_IN57_64 = 7                 'Read data from IN57 to IN64.


Public Const FBIDIO_IN1_16 = 0                  'Read data from IN1 to IN16.
Public Const FBIDIO_IN17_32 = 2                 'Read data from IN17 to IN32.
Public Const FBIDIO_IN33_48 = 4                 'Read data from IN33 to IN48.
Public Const FBIDIO_IN49_64 = 6                 'Read data from IN49 to IN64.

Public Const FBIDIO_IN1_32 = 0                  'Read data from IN1 to IN32.
Public Const FBIDIO_IN33_64 = 4                 'Read data from IN33 to IN64.

Public Const FBIDIO_OUT1_8 = 0                  'Write data to OUT1 to OUT8
Public Const FBIDIO_OUT9_16 = 1                 'Write data to OUT9 to OUT16
Public Const FBIDIO_OUT17_24 = 2                'Write data to OUT17 to OUT24
Public Const FBIDIO_OUT25_32 = 3                'Write data to OUT25 to OUT32
Public Const FBIDIO_OUT33_40 = 4                'Write data to OUT33 to OUT40
Public Const FBIDIO_OUT41_48 = 5                'Write data to OUT41 to OUT48
Public Const FBIDIO_OUT49_56 = 6                'Write data to OUT49 to OUT56
Public Const FBIDIO_OUT57_64 = 7                'Write data to OUT57 to OUT64


Public Const FBIDIO_OUT1_16 = 0                 'Write data to OUT1 to OUT16
Public Const FBIDIO_OUT17_32 = 2                'Write data to OUT17 to OUT32
Public Const FBIDIO_OUT33_48 = 4                'Write data to OUT33 to OUT48
Public Const FBIDIO_OUT49_64 = 6                'Write data to OUT49 to OUT64

Public Const FBIDIO_OUT1_32 = 0                 'Write data to OUT1 to OUT32
Public Const FBIDIO_OUT33_64 = 4                'Write data to OUT33 to OUT64

Public Const FBIDIO_STB1_ENABLE = &H1           'Enable STB1 event
Public Const FBIDIO_STB1_HIGH_EDGE = &H10       'Enable rising edge for STB1

Public Const FBIDIO_ACK2_ENABLE = &H4           'Enable ACK2 event
Public Const FBIDIO_ACK2_HIGH_EDGE = &H40       'Enable rising edge for ACK2


' -----------------------------------------------------------------------
'       Return value
' -----------------------------------------------------------------------
Public Const FBIDIO_ERROR_SUCCESS = 0                                                           ' Completed successfully
Public Const FBIDIO_ERROR_NOT_DEVICE = &HC0000001                               ' The device is not found.
Public Const FBIDIO_ERROR_NOT_OPEN = &HC0000002                                 ' The system could not open the device.
Public Const FBIDIO_ERROR_INVALID_HANDLE = &HC0000003                           ' The device handle is invalid.
Public Const FBIDIO_ERROR_ALREADY_OPEN = &HC0000004                             ' The device has been already opened.
Public Const FBIDIO_ERROR_HANDLE_EOF = &HC0000005                               ' End of file is reached.
Public Const FBIDIO_ERROR_MORE_DATA = &HC0000006                                ' More available data exists.
Public Const FBIDIO_ERROR_INSUFFICIENT_BUFFER = &HC0000007              ' Data area passed to the system call is too small.
Public Const FBIDIO_ERROR_IO_PENDING = &HC0000008                               ' An asynchronous I/O operation is in progress.
Public Const FBIDIO_ERROR_NOT_SUPPORTED = &HC0000009                            ' The feature is not supported.
Public Const FBIDIO_ERROR_MEMORY_NOTALLOCATED = &HC0001000              ' Allocating work area failed.
Public Const FBIDIO_ERROR_PARAMETER = &HC0001001                                ' Parameters passed to the function are invalid.
Public Const FBIDIO_ERROR_INVALID_CALL = &HC0001002                             ' Invalid function call was occurred.
Public Const FBIDIO_ERROR_DRVCAL = &HC0001003                                           ' The driver could not be called out.
Public Const FBIDIO_ERROR_NULL_POINTER = &HC0001004                             ' NULL pointer is passed between the driver and the DLL.

' -----------------------------------------------------------------------
'       ＤＬＬ
' -----------------------------------------------------------------------
Declare Function DioOpen Lib "FbiDio.DLL" (ByVal lpszName As String, ByVal fdwAttrs As Long) As Long
Declare Function DioClose Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DioInputPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pBuffer As Long, ByVal dwStartNum As Long, ByVal dwNum As Long) As Long
Declare Function DioOutputPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pBuffer As Long, ByVal dwStartNum As Long, ByVal dwNum As Long) As Long
Declare Function DioGetBackGroundUseTimer Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pnUse As Long) As Long
Declare Function DioSetBackGroundUseTimer Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nUse As Long) As Long
Declare Function DioSetBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal dwStartPoint As Long, ByVal dwPointNum As Long, ByVal dwValueNum As Long, ByVal dwCycle As Long, ByVal dwCount As Long, ByVal dwOption As Long) As Long
Declare Function DioFreeBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long) As Long
Declare Function DioStopBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long) As Long
Declare Function DioGetBackGroundStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long, ByRef pnStartPoint As Long, ByRef pnPointNum As Long, ByRef pnValueNum As Long, ByRef pnCycle As Long, ByRef pnCount As Long, ByRef pnOption As Long, ByRef pnExecute As Long, ByRef pnExecCount As Long, ByRef pnBufferOffset As Long, ByRef pnOver As Long) As Long
Declare Function DioInputPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long, ByRef pBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef pOverlapped As OVERLAPPED) As Long
Declare Function DioOutputPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long, ByRef pBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Declare Function DioWatchPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal hBackGroundHandle As Long, ByRef pBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Declare Function DioGetInputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pnInputHandShakeContig As Long, ByRef pdwBitMask1 As Long, ByRef pdwBitMask2 As Long) As Long
Declare Function DioSetInputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nInputHandShakeContig As Long, ByVal dwBitMask1 As Long, ByVal dwBitMask2 As Long) As Long
Declare Function DioGetOutputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pnOutputHandShakeContig As Long, ByRef pdwBitMask1 As Long, ByRef pdwBitMask2 As Long) As Long
Declare Function DioSetOutputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nOutputHandShakeConfig As Long, ByVal dwBitMask1 As Long, ByVal dwBitMask2 As Long) As Long
Declare Function DioInputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Declare Function DioInputHandShakeEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Declare Function DioOutputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumOfBytesWritten As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Declare Function DioOutputHandShakeEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Declare Function DioStopInputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DioStopOutputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DioGetHandShakeStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pdwpDeviceStatus As Long, ByRef pdwpInputedBuffNum As Long, ByRef pdwpOutputedBuffNum As Long) As Long
Declare Function DioInputByte Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByRef pbValue As Byte) As Long
Declare Function DioInputWord Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByRef pwValue As Integer) As Long
Declare Function DioInputDword Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByRef pdwValue As Long) As Long
Declare Function DioOutputByte Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByVal bValue As Byte) As Long
Declare Function DioOutputWord Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByVal wValue As Integer) As Long
Declare Function DioOutputDword Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal nNo As Long, ByVal dwValue As Long) As Long
Declare Function DioGetAckStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbAckStatus As Byte) As Long
Declare Function DioSetAckPulseCommand Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bCommand As Byte) As Long
Declare Function DioGetStbStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbStbStatus As Byte) As Long
Declare Function DioSetStbPulseCommand Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bCommand As Byte) As Long
Declare Function DioInputUniversalPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pdwUniversalPoint As Long) As Long
Declare Function DioOutputUniversalPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal dwUniversalPoint As Long) As Long
Declare Function DioSetTimeOut Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal dwInputTotalTimeout As Long, ByVal dwInputIntervalTimeout As Long, ByVal dwOutputTotalTimeout As Long, ByVal dwOutputIntervalTimeout As Long) As Long
Declare Function DioGetTimeOut Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pdwpInputTotalTimeout As Long, ByRef pdwpInputIntervalTimeout As Long, ByRef pdwpOutputTotalTimeout As Long, ByRef pdwpOutputIntervalTimeout As Long) As Long
Declare Function DioSetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bIrqMask As Byte) As Long
Declare Function DioGetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbIrqMask As Byte) As Long
Declare Function DioSetIrqConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bIrqConfig As Byte) As Long
Declare Function DioGetIrqConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbIrqConfig As Byte) As Long
Declare Function DioGetDeviceConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pdwDeviceConfig As Long) As Long
Declare Function DioSetTimerConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bTimerConfigValue As Byte) As Long
Declare Function DioGetTimerConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbTimerConfigValue As Byte) As Long
Declare Function DioGetTimerCount Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbTimerCount As Byte) As Long
Declare Function DioSetLatchStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal bLatchStatus As Byte) As Long
Declare Function DioGetLatchStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbLatchStatus As Byte) As Long
Declare Function DioGetResetInStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pbResetInStatus As Byte) As Long
Declare Function DioEventRequestPending Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByVal dwEventEnableMask As Long, ByRef pEventBuf As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Declare Function DioCommonGetPciDeviceInfo Lib "FbiDio.DLL" (ByVal hDeviceHandle As Long, ByRef pdwDeviceID As Long, ByRef pdwVenderID As Long, ByRef pdwClassCode As Long, ByRef pdwRevisionID As Long, ByRef pdwBaseAddress0 As Long, ByRef pdwBaseAddress1 As Long, ByRef pdwBaseAddress2 As Long, ByRef pdwBaseAddress3 As Long, ByRef pdwBaseAddress4 As Long, ByRef pdwBaseAddress5 As Long, ByRef pdwSubsystemID As Long, ByRef pdwSubsystemVenderID As Long, ByRef pdwInterruptLine As Long, ByRef pdwBoardID As Long) As Long

' -----------------------------------------------------------------------
'       from WIN32API
' -----------------------------------------------------------------------
Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
'* If a parameter except NULL is specified for lpEventAtrribute argument in the CreatEventA function,
'* an error will occur under Windows 98 and Windows 95.
'* To avoid this error, you should pass a NULL to lpEventAtrribute argument and pass a NULL to lpName argument for the unnamed event object. In case of that,
'* you should change variable type of these arguments. For convenience, we declare an alias of CreateEventA described below.
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Sub DioOut(ch%, dt%)
  Dim lpszName As String
  Dim dwStartNum As Long
  Dim nRet As Long
  Dim nBuffer As Long
  '
  If BrdFlg <> "ON" Then Exit Sub
  '
  dwStartNum = ch     'Val(IDC_STARTNUM.Text)
  nBuffer = dt        'Val(IDC_BUFFER.Text)
  DvcDioOpen
    nRet = DioOutputPoint(hDeviceHandle, nBuffer, dwStartNum, 1)
    If nRet <> 0 Then
        MsgBox ("Output the data failed.")
        nRet = DioClose(hDeviceHandle)
        Exit Sub
    End If
  DvcDioClose

End Sub

Public Sub DvcDioOpen()
Dim lpszName As String
'
    lpszName = "FBIDIO1" & Chr(0)
    hDeviceHandle = DioOpen(lpszName, FBIDIO_FLAG_SHARE)

    If hDeviceHandle = &HFFFF Then
        MsgBox ("Opening the board failed.")
        DoEvents
        Exit Sub
    End If
End Sub
Public Sub DvcDioClose()
Dim nRet As Long
    
    nRet = DioClose(hDeviceHandle)
    If nRet <> 0 Then
        MsgBox ("Closing the board failed.")
        Exit Sub
    End If
End Sub


Public Sub DioInput(ch%, hdt%)
  Dim lpszName As String
  Dim dwStartNum As Long
  Dim nRet As Long
  Dim nBuffer As Long
  
  If BrdFlg <> "ON" Then Exit Sub
  '
  dwStartNum = ch%    'Val(IDC_STARTNUM.Text)
  DvcDioOpen
    nRet = DioInputPoint(hDeviceHandle, nBuffer, dwStartNum, 1)
    If nRet <> 0 Then
        MsgBox ("Input the data failed.")
        'nRet = DioClose(hDeviceHandle)
        Exit Sub
    Else
        'MsgBox ("input data = " + Str(nBuffer))
    End If
    hdt = nBuffer
  DvcDioClose
End Sub

Public Sub ServoON()
'----------------- サーボＯＮ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 10, 1  'DioOut 22, 1    'Servo ON→Driver
  'DioOut 25, 1    'Servo ON→シーケンサ
  
End Sub
Public Sub ServoOFF()
'----------------- サーボＯＦＦ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 10, 0  'DioOut 22, 0    'Servo ON→Driver
  'DioOut 25, 0    'Servo ON→シーケンサ
  
End Sub
Public Sub ResetON()
'----------------- 東栄サーボアンプ　リセット　ＯＮ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 11, 1  '
End Sub
Public Sub ResetOFF()
'----------------- 東栄サーボアンプ　リセット　ＯＦＦ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 11, 0  '
End Sub
Public Sub CtlDisp()
'----------- ServoOn→位置制御→LSショート
  If BrdFlg <> "ON" Then Exit Sub
  ServoON
  DioOut 12, 1  '    '位置制御　　東栄　　080827
'  DioOut 12, 0  '   '位置制御=0 (tsubaki): 080320
'  DioOut 13, 1  '                 速度上限設定（内部設定へ切り替え）tsubaki
End Sub
Public Sub CtlVelo()
Dim disp!
'----------- ServoOn→速度制御→回転方向CW選択
  If BrdFlg <> "ON" Then Exit Sub
  ServoON
  DioOut 12, 0  '    '速度制御　　東栄　080827
'  DioOut 12, 1  '    '速度制御=1:  08.3.20 tsubaki
'  DioOut 13, 0  '                 速度上限設定（外部設定へ切り替え）tsubaki
  disp = r_z()
End Sub
Public Sub N2Open()
'----------- 冷却バルブOPEN
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 2, 1    '冷却
End Sub
Public Sub N2Close()
'----------- 冷却バルブClose
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 2, 0    '冷却
End Sub
Public Sub HeatON()
'----------- 加熱　ON
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 1, 1    '加熱
End Sub
Public Sub HeatOFF()
'----------- 加熱　OFF
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 1, 0    '加熱
End Sub
Public Sub CoolON()
'----------- 冷却大　ON
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 2, 1    '冷却大　ON
End Sub
Public Sub CoolOFF()
'----------- 冷却大　OFF
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 2, 0    '冷却大　OFF
End Sub
Public Sub SuireiON()
'----------- 水冷却　ON
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 4, 1    '水冷却　ON
End Sub
Public Sub SuireiOFF()
'----------- 水冷却　OFF
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 4, 0    '水冷却　OFF
End Sub

Public Function SystemReadyChk%()
Dim sts%, sts1%, sts2%
'----------- システムレディ or 非常停止
  If BrdFlg <> "ON" Then
    sts = 1
  Else
    sts = 0
    DioInput 1, sts1    '
    DioInput 7, sts2    '非常停止
    If sts1 = 1 And sts2 = 0 Then sts = 1
  End If
  If gDebugFlg = 1 Then
    sts = 1
  End If
  SystemReadyChk = sts
End Function
Public Function GentenCmdChk%()
Dim sts%
'----------- 搬送シリンダ原点　確認
  If BrdFlg <> "ON" Then
    sts = 1
  Else
    DioInput 4, sts    '
    
  End If
  GentenCmdChk = sts
End Function
Public Sub SeikeiON()
'----------- 成形ON　　連続成形、１回成形　ｒｕｎ中
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 3, 1    '成形ON　　連続成形、１回成形　ｒｕｎ中
'
'Public Sub VacuumON()
''----------- 真空到達
'  If BrdFlg <> "ON" Then Exit Sub
'  DioOut 3, 1    '真空到達
End Sub
Public Sub SeikeiOFF()
'----------- 成形OFF 待機中
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 3, 0    '成形OFF 待機中
'Public Sub VacuumOFF()
''----------- 真空未到達
'  If BrdFlg <> "ON" Then Exit Sub
'  DioOut 3, 0    '真空未到達
End Sub
Public Sub OrgON()
'----------- 原点位置
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 15, 1    '
  gOrgIL = True
End Sub
Public Sub OrgOFF()
'----------- 原点位置
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 15, 0    '
  gOrgIL = False
End Sub
Public Sub DioAllReset()
  Dim lpszName As String
  Dim dwStartNum As Long
  Dim nRet As Long
  Dim nBuffer As Long
  
  If BrdFlg <> "ON" Then Exit Sub
  '
  nBuffer = 0
  DvcDioOpen
    'nRet = DioOutputDword(hDeviceHandle, nBuffer, dwStartNum, 1)
    If nRet <> 0 Then
        MsgBox ("Input the data failed.")
        Exit Sub
    End If
    hdt = nBuffer
  DvcDioClose
End Sub

Public Sub TrnsReqON()
'----------- 搬送依頼信号ＯＮ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 21, 1    '
  'DioOut 13, 1    '
End Sub
Public Sub TrnsReqOFF()
'----------- 搬送依頼信号ＯＦＦ
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 21, 0    '
  'DioOut 13, 0    '
End Sub
Public Function ArmChk%()
Dim sts%
'-----------  アラームチェック
  If BrdFlg <> "ON" Then
    sts = 0
  Else
    sts = 0
    DioInput 8, sts     'アラーム
  End If
  ArmChk = sts
End Function
Public Function ArmMsgChk$()
Dim b0%, b1%, b2%, b3%, b4%, ErrNo%, EmgArm%
'-----------  アラームメッセージ
  EmgArm = 1
  DioInput 20, b0    'アラームBits-0
  DioInput 21, b1    'アラームBits-1
  DioInput 22, b2    'アラームBits-2
  DioInput 23, b3    'アラームBits-3
  DioInput 17, b4    'アラームBits-4
  ErrNo = b0 * 1 + b1 * 2 + b2 * 4 + b3 * 8 + b4 * 16
  ArmMsgChk = gErrMsg$(EmgArm, ErrNo)
End Function
Public Function EmgMsgChk$()
Dim b0%, b1%, b2%, b3%, b4%, ErrNo%, EmgArm%
'-----------  非常停止メッセージ
  EmgArm = 0
  DioInput 20, b0    'EmgBits-0
  DioInput 21, b1    'EmgBits-1
  DioInput 22, b2    'EmgBits-2
  DioInput 23, b3    'EmgBits-3
  DioInput 17, b4    'EmgBits-4
  ErrNo = b0 * 1 + b1 * 2 + b2 * 4 + b3 * 8 + b4 * 16
  EmgMsgChk = gErrMsg$(EmgArm, ErrNo)
End Function
Public Function ArmEmgMsgChk$()
Dim b0%, b1%, b2%, b3%, b4%, ErrNo%, EmgArm%, b7%, b8%
'-----------  非常停止メッセージ
  DioInput 7, b7     'EMG出力
  DioInput 8, b8     'ARM出力
  If b7 = 1 Then EmgArm = 0
  If b8 = 1 Then EmgArm = 1
  DioInput 20, b0    'EmgBits-0
  DioInput 21, b1    'EmgBits-1
  DioInput 22, b2    'EmgBits-2
  DioInput 23, b3    'EmgBits-3
  DioInput 17, b4    'EmgBits-4
  ErrNo = b0 * 1 + b1 * 2 + b2 * 4 + b3 * 8 + b4 * 16
  ArmEmgMsgChk = gErrMsg$(EmgArm, ErrNo)
End Function
Public Function VacuumTimeOutChk%()
Dim b0%, b1%, b2%, b3%, b4%, ErrNo%, EmgArm%, b7%, b8%
'-----------  真空未到達の確認
  DioInput 7, b7     'EMG出力
  DioInput 8, b8     'ARM出力
  If b7 = 1 Then EmgArm = 0
  If b8 = 1 Then EmgArm = 1
  DioInput 20, b0    'EmgBits-0
  DioInput 21, b1    'EmgBits-1
  DioInput 22, b2    'EmgBits-2
  DioInput 23, b3    'EmgBits-3
  DioInput 17, b4    'EmgBits-4
  ErrNo = b0 * 1 + b1 * 2 + b2 * 4 + b3 * 8 + b4 * 16
  If ErrNo = 12 And EmgArm = 1 Then
    VacuumTimeOutChk = True     '真空未到達
    frmerr_Vcuum.Show
  Else
    VacuumTimeOutChk = False
    Unload frmerr_Vcuum
  End If
End Function
Public Function KataChk%()
Dim sts%, sts1%, sts2%, sts3%
'----------- 成形室 & 予熱室に型が在るか？　'04.9.26変更　s.f
'                                           '08.4.22変更　予備加熱①、②、成形室　３室へ
'                                           '08.3.15変更　予備加熱①、②、成形室　３室へ
If BrdFlg <> "ON" Then
    sts = 0
  Else
    sts = 0                                   ' 型が無い
    DioInput 18, sts1    '成形室     =1；型が在る
    DioInput 19, sts2    '予備加熱２ =1；型が在る
    DioInput 14, sts3    '予備加熱１ =1；型が在る
'                                           '06.12.21 change s.f
'    If sts1 = 0 And sts2 = 1 Then sts = 1   '04.9.26 change　　予備室２のみ型有り
'    If sts1 = 1 And sts2 = 0 Then sts = 2   '04.9.26 change　　成形室のみ型有り
'    If sts1 = 1 And sts2 = 1 Then sts = 3   '04.9.26 change　　成形室、予備室２共に型有り
    sts = sts1 * 4 + sts2 * 2 + sts3
'
End If
  KataChk = sts
End Function

Public Function TrnsFinChk%()
Dim sts%
  DioInput 13, sts        '搬送終了？
  TrnsFinChk = sts
End Function

Public Function AutoChk%()
Dim sts%
  DioInput 3, sts        '自動状態？
  AutoChk = sts
End Function
Public Function PCTrnsChk%()
Dim sts%
  DioInput 6, sts        'PLC　搬送中=1
  PCTrnsChk = sts
End Function
Public Sub PCTrnsReq()
'----------- パレット1順指令
  If BrdFlg <> "ON" Then Exit Sub
  DioOut 9, 1    '
  WaitSec 1
  DioOut 9, 0
End Sub

