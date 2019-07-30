Attribute VB_Name = "FbiAd"
'------------------------------------------------
'  08.07.26  update: s.f.:   TempRdMold!(index)  index 追加
'-------------------------------------------------
'
'   GPC-3100 VisualBasic  Overlapped process identifier
'
'-------------------------------------------------

Public Const FLAG_SYNC = 1                    ' Sampling in a background thread
Public Const FLAG_ASYNC = 2                   ' Sampling in asynchronous operation


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  File format identifier
'
'-----------------------------------------------------------------------------------------------
Public Const FLAG_BIN = 1                     ' Binary format file
Public Const FLAG_CSV = 2                     ' CSV format file


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Sampling status identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_STATUS_STOP_SAMPLING = 1      ' A sampling has been stopped.
Public Const AD_STATUS_WAIT_TRIGGER = 2       ' A sampling is in a waiting state for a trigger.
Public Const AD_STATUS_NOW_SAMPLING = 3       ' A sampling is running.


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Event factor identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_EVENT_SMPLNUM = 1             ' An event that will be signaled when specified a number of samples are acquired.
Public Const AD_EVENT_STOP_TRIGGER = 2        ' A sampling stopped because a trigger asserted.
Public Const AD_EVENT_STOP_FUNCTION = 3       ' A sampling is stopped by software
Public Const AD_EVENT_STOP_TIMEOUT = 4        ' The sampling terminated because a timeout interval elapsed.
Public Const AD_EVENT_STOP_SAMPLING = 5       ' The sampling is completed


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Input configuration identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_INPUT_SINGLE = 1              ' Single-ended
Public Const AD_INPUT_DIFF = 2                ' Differential


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Volume identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_ADJUST_BIOFFSET = 1           ' Bipolar offset adjustment
Public Const AD_ADJUST_UNIOFFSET = 2          ' Unipolar offset adjustment
Public Const AD_ADJUST_BIGAIN = 3             ' Bipolar gain adjustment
Public Const AD_ADJUST_UNIGAIN = 4            ' Unipolar gain adjustment


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Adjustment item identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_ADJUST_UP = 1                 ' Up
Public Const AD_ADJUST_DOWN = 2               ' Down
Public Const AD_ADJUST_STORE = 3              ' Store
Public Const AD_ADJUST_STANDBY = 4            ' Standby
Public Const AD_ADJUST_NOT_STORE = 5          ' Not stored


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Data identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_DATA_PHYSICAL = 1             ' Physical value (voltage [V], current [mA])
Public Const AD_DATA_BIN8 = 2                 ' 8bit binary
Public Const AD_DATA_BIN12 = 3                ' 12bit binary
Public Const AD_DATA_BIN16 = 4                ' 16bit binary
Public Const AD_DATA_BIN24 = 5                ' 24bit binary


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Data conversion identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_CONV_SMOOTH = 1               ' Smoothing is applied to samples.
Public Const AD_CONV_AVERAGE1 = &H100         ' Averaging is applied to samples.
Public Const AD_CONV_AVERAGE2 = &H200         ' Shifted averaging is applied to samples.


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Sampling mode identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_IO_SAMPLING = 1               ' I/O
Public Const AD_FIFO_SAMPLING = 2             ' FIFO
Public Const AD_MEM_SAMPLING = 4              ' Memory


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Trigger point identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_TRIG_START = 1                ' Start-trigger(Default)
Public Const AD_TRIG_STOP = 2                 ' Stop-trigger
Public Const AD_TRIG_START_STOP = 3           ' Start/stop-trigger


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Trigger level identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_FREERUN = 1                   ' Trigger-less(Default)
Public Const AD_EXTTRG = 2                    ' External trigger
Public Const AD_EXTTRG_DI = 3                 ' External trigger with DI masking
Public Const AD_LEVEL_P = 4                   ' Level trigger: level sensitive (low-to-high transition)
Public Const AD_LEVEL_M = 5                   ' Level trigger: level sensitive (high-to-low transition)
Public Const AD_LEVEL_D = 6                   ' Level trigger: level sensitive
Public Const AD_INRANGE = 7                   ' Level trigger: into the range
Public Const AD_OUTRANGE = 8                  ' Level trigger: out of the range
Public Const AD_ETERNITY = 9                  ' Indefinite sampling
Public Const AD_START_P1 = &H10               ' Start-trigger: Level 1: low-to- high transition
Public Const AD_START_M1 = &H20               ' Start-trigger: Level 1: high-to-low transition
Public Const AD_START_D1 = &H40               ' Start-trigger: Level 1: high-to-low or low -to-high transition (direction DON'T CARE)
Public Const AD_START_P2 = &H80               ' Start-trigger: Level 2: low-to- high transition
Public Const AD_START_M2 = &H100              ' Start-trigger: Level 2: high-to-low transition
Public Const AD_START_D2 = &H200              ' Start-trigger: Level 2: high-to-low or low -to-high transition (direction DON'T CARE)
Public Const AD_STOP_P1 = &H400               ' Stop-trigger: Level 1: low-to- high transition
Public Const AD_STOP_M1 = &H800               ' Stop-trigger: Level 1: high-to-low transition
Public Const AD_STOP_D1 = &H1000              ' Stop-trigger: Level 1: high-to-low or low -to-high transition (direction DON'T CARE)
Public Const AD_STOP_P2 = &H2000              ' Stop-trigger: Level 2: low-to- high transition
Public Const AD_STOP_M2 = &H4000              ' Stop-trigger: Level 2: high-to-low transition
Public Const AD_STOP_D2 = &H8000              ' Stop-trigger: Level 2: high-to-low or low -to-high transition (direction DON'T CARE)
Public Const AD_ANALOG_FILTER = &H10000       ' Use an analog trigger filter


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Polarity identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_DOWN_EDGE = 1                 ' Falling edge(Default)
Public Const AD_UP_EDGE = 2                   ' Rising edge


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Pulse polarity identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_LOW_PULSE = 1                 ' Low-pulse(Default)
Public Const AD_HIGH_PULSE = 2                ' High-pulse


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Double-clocked mode identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_NORMAL_MODE = 1               ' Use it in normal clock mode
Public Const AD_FAST_MODE = 2                 ' Use double-clocked mode


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Range identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_0_1V = &H1                    ' Voltage: unipolar 0 to 1 V
Public Const AD_0_2P5V = &H2                  ' Voltage: unipolar 0 to 2.5 V
Public Const AD_0_5V = &H4                    ' Voltage: unipolar 0 to 5 V
Public Const AD_0_10V = &H8                   ' Voltage: unipolar 0 to 10 V
Public Const AD_1_5V = &H10                   ' Voltage: unipolar 1 to 5 V
Public Const AD_0_20mA = &H1000               ' Current: unipolar 0 to 20 mA
Public Const AD_4_20mA = &H2000               ' Current: unipolar 4 to 20 mA
Public Const AD_1V = &H10000                  ' Voltage: bipolar +/- 1 V
Public Const AD_2P5V = &H20000                ' Voltage: bipolar +/- 2.5 V
Public Const AD_5V = &H40000                  ' Voltage: bipolar +/- 5 V
Public Const AD_10V = &H80000                 ' Voltage: bipolar +/- 10 V


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Isolation identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_ISOLATION = 1                 ' Photo-isolated board
Public Const AD_NOT_ISOLATION = 2             ' Not isolated board


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Synchronous mode identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_MASTER_MODE = 1               ' Master mode
Public Const AD_SLAVE_MODE = 2                ' Slave mode


'-------------------------------------------------
'
'   GPC-3100 VisualBasic  Structure declaration
'
'-------------------------------------------------

' -----------------------------------------------------------------------
'  Sampling request condition structure for each channel
' -----------------------------------------------------------------------
Type ADSMPLCHREQ
    ulChNo As Long
    ulRange As Long
End Type

' -----------------------------------------------------------------------
'  Sampling request condition structure
' -----------------------------------------------------------------------
Type ADSMPLREQ
    ulChCount As Long
    SmplChReq(0 To 255) As ADSMPLCHREQ
    ulSamplingMode As Long
    ulSingleDiff As Long
    ulSmplNum As Long
    ulSmplEventNum As Long
    fSmplFreq As Single
    ulTrigPoint As Long
    ulTrigMode As Long
    lTrigDelay As Long
    ulTrigCh As Long
    fTrigLevel1 As Single
    fTrigLevel2 As Single
    ulEClkEdge As Long
    ulATrgPulse As Long
    ulTrigEdge As Long
    ulTrigDI As Long
    ulFastMode As Long
End Type

' -----------------------------------------------------------------------
'  Board specification structure
' -----------------------------------------------------------------------
Type ADBOARDSPEC
    ulBoardType As Long
    ulBoardID As Long
    dwSamplingMode As Long
    ulChCountS As Long
    ulChCountD As Long
    ulResolution As Long
    dwRange As Long
    ulIsolation As Long
    ulDi As Long
    ulDo As Long
End Type


'-----------------------------------------------------------------------------------------------
'
'   GPC-3100 VisualBasic  Error identifier
'
'-----------------------------------------------------------------------------------------------
Public Const AD_ERROR_SUCCESS = 0
Public Const AD_ERROR_NOT_DEVICE = &HC0000001
Public Const AD_ERROR_NOT_OPEN = &HC0000002
Public Const AD_ERROR_INVALID_HANDLE = &HC0000003
Public Const AD_ERROR_ALREADY_OPEN = &HC0000004
Public Const AD_ERROR_NOT_SUPPORTED = &HC0000009
Public Const AD_ERROR_NOW_SAMPLING = &HC0001001
Public Const AD_ERROR_STOP_SAMPLING = &HC0001002
Public Const AD_ERROR_START_SAMPLING = &HC0001003
Public Const AD_ERROR_SAMPLING_TIMEOUT = &HC0001004
Public Const AD_ERROR_INVALID_PARAMETER = &HC0001021
Public Const AD_ERROR_ILLEGAL_PARAMETER = &HC0001022
Public Const AD_ERROR_NULL_POINTER = &HC0001023
Public Const AD_ERROR_GET_DATA = &HC0001024
Public Const AD_ERROR_FILE_OPEN = &HC0001041
Public Const AD_ERROR_FILE_CLOSE = &HC0001042
Public Const AD_ERROR_FILE_READ = &HC0001043
Public Const AD_ERROR_FILE_WRITE = &HC0001044
Public Const AD_ERROR_INVALID_DATA_FORMAT = &HC0001061
Public Const AD_ERROR_INVALID_AVERAGE_OR_SMOOTHING = &HC0001062
Public Const AD_ERROR_INVALID_SOURCE_DATA = &HC0001063
Public Const AD_ERROR_NOT_ALLOCATE_MEMORY = &HC0001081
Public Const AD_ERROR_NOT_LOAD_DLL = &HC0001082
Public Const AD_ERROR_CALL_DLL = &HC0001083



'-------------------------------------------------
'
'   GPC-3100 VisualBasic  Function declaration
'
'-------------------------------------------------

Declare Function AdOpen Lib "FbiAd.DLL" (ByVal lpszName As String) As Long
Declare Function AdClose Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function AdGetDeviceInfo Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef BoardSpec As ADBOARDSPEC) As Long
Declare Function AdSetBoardConfig Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal hEvent As Long, ByVal lpEventProc As Long, ByVal dwUser As Long) As Long
Declare Function AdGetBoardConfig Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef ulAdSmplEventFactor As Long) As Long
Declare Function AdSetSamplingConfig Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef pAdSmplConfig As ADSMPLREQ) As Long
Declare Function AdGetSamplingConfig Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef pAdSmplConfig As ADSMPLREQ) As Long
Declare Function AdGetSamplingData Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef pSmplData As Any, ByRef ulSmplNum As Long) As Long
Declare Function AdClearSamplingData Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function AdStartSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulSyncFlag As Long) As Long
Declare Function AdStartFileSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal pszPathName As String, ByVal ulFileFlag As Long) As Long
Declare Function AdTriggerSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulChNo As Long, ByVal ulRange As Long, ByVal ulSingleDiff As Long, ByVal ulTriggerMode As Long, ByVal ulTrigEdge As Long, ByVal ulSmplNum As Long) As Long
Declare Function AdMemTriggerSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulChCount As Long, ByRef lpSmplChReq As ADSMPLCHREQ, ByVal ulSmplNum As Long, ByVal ulRepeatCount As Long, ByVal ulTrigEdge As Long, ByVal fSmplFreq As Single, ByVal ulEClkEdge As Long, ByVal ulFastMode As Long) As Long
Declare Function AdSyncSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulMode As Long) As Long
Declare Function AdStopSampling Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function AdGetStatus Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef ulAdSmplStatus As Long, ByRef ulAdSmplCount As Long, ByRef ulAdAvailCount As Long) As Long
Declare Function AdInputAD Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulCh As Long, ByVal ulSingleDiff As Long, ByRef lpAdSmplChReq As ADSMPLCHREQ, ByRef lpData As Any) As Long
Declare Function AdInputDI Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByRef dwData As Long) As Long
Declare Function AdOutputDO Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal dwData As Long) As Long
Declare Function AdAdjustVR Lib "FbiAd.DLL" (ByVal hDeviceHandle As Long, ByVal ulAdjustCh As Long, ByVal ulSingleDiff As Long, ByVal ulSelVolume As Long, ByVal ulControl As Long, ByVal ulTap As Long) As Long
Declare Function AdDataConv Lib "FbiAdDC.DLL" (ByVal uSrcFormCode As Long, ByRef pSrcData As Any, ByVal uSrcSmplDataNum As Long, ByRef pSrcSmplReq As ADSMPLREQ, ByVal uDestFormCode As Long, ByRef pDestData As Any, ByRef puDestSmplDataNum As Long, ByRef pDestSmplReq As ADSMPLREQ, ByVal uEffect As Long, ByVal uCount As Long, ByVal lpfnConv As Long) As Long
Declare Function AdReadFile Lib "FbiAdDC.DLL" (ByVal pszPathName As String, ByRef pSmplData As Any, ByVal uFormCode As Long) As Long

Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal ManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long




Public Sub DvcAdOpen()
    Dim i As Long
    Dim nRet As Long
    Dim lpszName As String
    
    'Open "FBIAD**"
    For i = 0 To 256
        lpszName = "FBIAD"
        lpszName = lpszName & (i + 1)
        ghDeviceHandleAD = AdOpen(lpszName)
        If ghDeviceHandleAD <> -1 Then
            AdClose (ghDeviceHandleAD)
            szName = lpszName
            lpszNameAD = lpszName
            'DeviceOpen.DeviceName.Text = lpszName
        End If
    Next i
        
    'Display a device open dialog.
    'DeviceOpen.Show 1
End Sub
Public Sub DvcAdClose()
    Dim nRet As Long
    
    If ghOpenFlag = 1 Then
        nRet = AdClose(ghDeviceHandleAD)
        If nRet <> AD_ERROR_SUCCESS Then
            Call DsplyErrMessage(nRet)
        Else
            'Call DsplyErrMessage(nRet)
            ghOpenFlag = 0
            'Main.IDM_DEVICE_OPEN.Enabled = True
        End If
    End If
    
End Sub
Public Sub DeviceAdName()
    Dim nRet As Long
    Dim lpszName As String

    lpszName = lpszNameAD   'DeviceName.Text
    
    ghDeviceHandleAD = AdOpen(lpszName)
    If ghDeviceHandleAD = -1 Then
        Call DsplyErrMessage(ghDeviceHandleAD)
        'Unload Me
    Else
        'Call DsplyErrMessage(0)
    
        ghChannel(0) = 1
        ghChannel(1) = 2
        ghOpenFlag = 1
        
        'adMain.IDM_DEVICE_OPEN.Enabled = False
        
        ' Read a sampling request condition
        nRet = AdGetSamplingConfig(ghDeviceHandleAD, gConfig)
        If nRet <> AD_ERROR_SUCCESS Then
            Call DsplyErrMessage(nRet)
        End If

        ' Read a device information
        nRet = AdGetDeviceInfo(ghDeviceHandleAD, gInfo)
        If nRet <> AD_ERROR_SUCCESS Then
            'Call DsplyErrMessage(nRet)
        End If

    End If

End Sub
Public Function AdRead1Ch%(ch%)
    Dim nRet As Long
    Dim SmplChInf(0 To 1) As ADSMPLCHREQ
    Dim bSmpData() As Byte
    Dim wSmpData() As Integer
    Dim dwSmpData() As Long
    Dim szDisp As String
    Dim ulCh As Long
    Dim hdt%

    ' Retrieve a channel number
    If IsNull(ch) Then
        nRet = MsgBox("Invalid channel", (vbOKOnly + vbCritical), "Error Code")
        Exit Function
    End If
    
    ghChannel(0) = ch
    ghChannel(1) = ch + 1
    
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
        
        nRet = AdInputAD(ghDeviceHandleAD, ulCh, gConfig.ulSingleDiff, SmplChInf(0), bSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(bSmpData(0)) & "h"
            hdt = wSmpData(0)
        End If
    ElseIf gInfo.ulResolution > 8 And gInfo.ulResolution <= 16 Then
        ReDim wSmpData(ulCh)
        
        nRet = AdInputAD(ghDeviceHandleAD, ulCh, gConfig.ulSingleDiff, SmplChInf(0), wSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(wSmpData(0)) & "h"
            hdt = wSmpData(0)
        End If
    ElseIf gInfo.ulResolution > 16 Then
        ReDim dwSmpData(ulCh)

        nRet = AdInputAD(ghDeviceHandleAD, ulCh, gConfig.ulSingleDiff, SmplChInf(0), dwSmpData(0))
        If nRet = AD_ERROR_SUCCESS Then
            ' Display a Sampling Data
            szDisp = Hex(dwSmpData(0)) & "h"
            hdt = wSmpData(0)
        End If
    End If
    
    If nRet = AD_ERROR_SUCCESS Then
        'nRet = MsgBox(szDisp, (vbOKOnly + vbInformation), "Sampling Data")
    Else
        'Call DsplyErrMessage(nRet)
    End If
  AdRead1Ch = hdt 'dwSmpData(0)
End Function
Public Sub AdRead(dt!(), flg As Long)
Dim Data(0 To 7) As Integer
Dim SmplChReq(0 To 7) As ADSMPLCHREQ
Dim nRet As Long
  ppos = Left(ppos, 22) & " (AdR1)"

  If BrdFlg <> "ON" Then Exit Sub
  SmplChReq(0).ulChNo = 1       '成形室IHヒーター温度
  SmplChReq(0).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(1).ulChNo = 2       '
  SmplChReq(1).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(2).ulChNo = 3       '荷重
  SmplChReq(2).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(3).ulChNo = 4       '
  SmplChReq(3).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(4).ulChNo = 5       '
  SmplChReq(4).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(5).ulChNo = 6       '上モールド温度
  SmplChReq(5).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(6).ulChNo = 7       '下モールド温度
  SmplChReq(6).ulRange = gConfig.SmplChReq(0).ulRange
  SmplChReq(7).ulChNo = 8       '
  SmplChReq(7).ulRange = gConfig.SmplChReq(0).ulRange
  
  
'  nRet = AdInputAD(ghDeviceHandleAD, 5, gConfig.ulSingleDiff, SmplChReq(0), Data(0))
  nRet = AdInputAD(ghDeviceHandleAD, 8, gConfig.ulSingleDiff, SmplChReq(0), Data(0))
  'nRet = AdInputAD(ghDeviceHandleAD, 4, AD_INPUT_SINGLE, SmplChReq(0), Data(0))

 ppos = Left(ppos, 22) & " (AdR2)"
    For i = 0 To 7
      dt(i) = (Data(i) - &H7FF) / 204.8    '2048*10(V)
    Next i

  flg = nRet
End Sub

Public Function LoadSet!(v!)
'----------- 荷重（電圧を荷重値）
    LoadSet = v * 10#
End Function

Public Function VacuSet!(v!)
'----------- 真空計
    VacuSet = v * 1#
End Function
Public Function FTZ6Set!(v!)
'----------- 放射温度計
    FTZ6Set = v * 1#
End Function
Public Function TempRdMold!(index)
'  08.07.26  update: s.f.:   TempRdMold!(index)  index 追加
'
Dim flg As Long, l As Integer
Dim dt!(0 To 7), sum As Double
    sum = 0#
    For l = 1 To 50
        AdRead dt(), flg
        sum = sum + dt(index)
    Next l
    TempRdMold = (sum / 50#) * 1000# / 10#  'スリーブ温度=V*(1000℃/10V)
End Function
