Attribute VB_Name = "FbiDa"
'-----------------------------------------------------------------------------------------------
'
'   Overlapped process identifier
'
'-----------------------------------------------------------------------------------------------
'
Public Const FLAG_SYNC = 1                    ' Output in a background thread.
Public Const FLAG_ASYNC = 2                   ' Output in asynchronous operation.


'-----------------------------------------------------------------------------------------------
'
'   File format identifier
'
'-----------------------------------------------------------------------------------------------
Public Const FLAG_BIN = 1                     ' Binary format file
Public Const FLAG_CSV = 2                     ' CSV format file


'-----------------------------------------------------------------------------------------------
'
'   Analog output status identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_STATUS_STOP_SAMPLING = 1      ' The output is stopped.
Public Const DA_STATUS_WAIT_TRIGGER = 2       ' The output is waiting for an assertion of the trigger.
Public Const DA_STATUS_NOW_SAMPLING = 3       ' The output is running.


'-----------------------------------------------------------------------------------------------
'
'   Event factor identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_EVENT_STOP_TRIGGER = 1        ' The output has been stopped because a trigger asserted.
Public Const DA_EVENT_STOP_FUNCTION = 2       ' The output has been stopped by software.
Public Const DA_EVENT_STOP_SAMPLING = 3       ' The Analog output is completed.
Public Const DA_EVENT_RESET_IN = 4            ' The reset input signal asserted.
Public Const DA_EVENT_CURRENT_OFF = 5         ' The current breaking has been detected.


'-----------------------------------------------------------------------------------------------
'
'   Volume identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_ADJUST_BIOFFSET = 1           ' Bipolar offset adjustment
Public Const DA_ADJUST_UNIOFFSET = 2          ' Unipolar offset adjustment
Public Const DA_ADJUST_BIGAIN = 3             ' Bipolar gain adjustment
Public Const DA_ADJUST_UNIGAIN = 4            ' Unipolar gain adjustment


'-----------------------------------------------------------------------------------------------
'
'   Adjustment item identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_ADJUST_UP = 1                 ' Up
Public Const DA_ADJUST_DOWN = 2               ' Down
Public Const DA_ADJUST_STORE = 3              ' Store
Public Const DA_ADJUST_STANDBY = 4            ' Standby
Public Const DA_ADJUST_NOT_STORE = 5          ' Not stored


'-----------------------------------------------------------------------------------------------
'
'   Data identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_DATA_PHYSICAL = 1             ' Physical value (voltage [V], current [mA])
Public Const DA_DATA_BIN8 = 2                 ' 8bit binary
Public Const DA_DATA_BIN12 = 3                ' 12bit binary
Public Const DA_DATA_BIN16 = 4                ' 16bit binary
Public Const DA_DATA_BIN24 = 5                ' 24bit binary


'-----------------------------------------------------------------------------------------------
'
'   Data conversion identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_CONV_SMOOTH = 1               ' Smoothing is applied to samples.
Public Const DA_CONV_AVERAGE1 = &H100         ' Averaging is applied to samples.
Public Const DA_CONV_AVERAGE2 = &H200         ' Shifted averaging is applied to samples.


'-----------------------------------------------------------------------------------------------
'
'   Data transfer architecture identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_IO_SAMPLING = 1               ' I/O
Public Const DA_FIFO_SAMPLING = 2             ' FIFO
Public Const DA_MEM_SAMPLING = 4              ' Memory


'-----------------------------------------------------------------------------------------------
'
'   Trigger point identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_TRIG_START = 1                ' Start-trigger(Default)
Public Const DA_TRIG_STOP = 2                 ' Stop-trigger
Public Const DA_TRIG_START_STOP = 3           ' Start/stop-trigger


'-----------------------------------------------------------------------------------------------
'
'   Trigger level identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_FREERUN = 1                   ' Trigger-less(Default)
Public Const DA_EXTTRG = 2                    ' External trigger
Public Const DA_EXTTRG_DI = 3                 ' External trigger with DI masking


'-----------------------------------------------------------------------------------------------
'
'   Polarity identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_DOWN_EDGE = 1                 ' Falling edge(Default)
Public Const DA_UP_EDGE = 2                   ' Rising edge


'-----------------------------------------------------------------------------------------------
'
'   Range identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_0_1V = &H1                    ' Voltage: unipolar 0 to 1 V
Public Const DA_0_2P5V = &H2                  ' Voltage: unipolar 0 to 2.5 V
Public Const DA_0_5V = &H4                    ' Voltage: unipolar 0 to 5 V
Public Const DA_0_10V = &H8                   ' Voltage: unipolar 0 to 10 V
Public Const DA_1_5V = &H10                   ' Voltage: unipolar 1 to 5 V
Public Const DA_0_20mA = &H1000               ' Current: unipolar 0 to 20 mA
Public Const DA_4_20mA = &H2000               ' Current: unipolar 4 to 20 mA
Public Const DA_1V = &H10000                  ' Voltage: bipolar +/- 1 V
Public Const DA_2P5V = &H20000                ' Voltage: bipolar +/- 2.5 V
Public Const DA_5V = &H40000                  ' Voltage: bipolar +/- 5 V
Public Const DA_10V = &H80000                 ' Voltage: bipolar +/- 10 V


'-----------------------------------------------------------------------------------------------
'
'   Isolation identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_ISOLATION = 1                 ' Photo-isolated board
Public Const DA_NOT_ISOLATION = 2             ' Not isolated board


'-----------------------------------------------------------------------------------------------
'
'   Range identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_RANGE_UNIPOLAR = 1            ' Unipolar
Public Const DA_RANGE_BIPOLAR = 2             ' Bipolar

'-----------------------------------------------------------------------------------------------
'
'   Waveform generation mode identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_MODE_CUT = 1                  ' Time-based waveform generation
Public Const DA_MODE_SYNTHE = 2               ' Frequency-based waveform generation

'-----------------------------------------------------------------------------------------------
'
'   Repeat mode identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_REPEAT_NONINTERVAL = 1        ' Repeat without the wait state (Default)
Public Const DA_REPEAT_INTERVAL = 2           ' Repeat with the wait state

'-----------------------------------------------------------------------------------------------
'
'   Counter clear identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_COUNTER_CLEAR = 1             ' Cleared (Default)
Public Const DA_COUNTER_NONCLEAR = 2          ' Not cleared

'-----------------------------------------------------------------------------------------------
'
'   DA latch identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_LATCH_CLEAR = 1               ' The voltage is set to the lowest voltage of the range.
Public Const DA_LATCH_NONCLEAR = 2            ' The voltage is sustained.

'-----------------------------------------------------------------------------------------------
'
'   Clock source identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_CLOCK_TIMER = 1               ' Internal programmable timer (8254 compatible)
Public Const DA_CLOCK_FIXED = 2               ' Fixed 5 MHz clock

'-----------------------------------------------------------------------------------------------
'
'   Configurations of the connector CN3 identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_EXTRG_IN = 1                  ' External trigger input (Default)
Public Const DA_EXTRG_OUT = 2                 ' External trigger output

'-----------------------------------------------------------------------------------------------
'
'   Configurations of the connector CN4 identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_EXCLK_IN = 1                  ' External clock input (Default)
Public Const DA_EXCLK_OUT = 2                 ' External clock output

'-----------------------------------------------------------------------------------------------
'
'   Filter identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_FILTER_OFF = 1                ' Not used (Default)
Public Const DA_FILTER_ON = 2                 ' Used

'-----------------------------------------------------------------------------------------------
'
'   Synchronous analog output identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_MASTER_MODE = 1               ' Master Mode
Public Const DA_SLAVE_MODE = 2                ' Slave Mode

'-------------------------------------------------
'
'   Structure declaration
'
'-------------------------------------------------

' -----------------------------------------------------------------------
'  Analog output request condition structure for each channel
' -----------------------------------------------------------------------
Type DASMPLCHREQ
    ulChNo As Long
    ulRange As Long
End Type

' -----------------------------------------------------------------------
'  Analog output request condition structure
' -----------------------------------------------------------------------
Type DASMPLREQ
    ulChCount As Long
    SmplChReq(0 To 255) As DASMPLCHREQ
    ulSamplingMode As Long
    fSmplFreq As Single
    ulSmplRepeat As Long
    ulTrigMode As Long
    ulTrigPoint As Long
    ulTrigDelay As Long
    ulEClkEdge As Long
    ulTrigEdge As Long
    ulTrigDI As Long
End Type

' -----------------------------------------------------------------------
'  Board specification structure
' -----------------------------------------------------------------------
Type DABOARDSPEC
    ulBoardType As Long
    ulBoardID As Long
    ulSamplingMode As Long
    ulChCount As Long
    ulResolution As Long
    ulRange As Long
    ulIsolation As Long
    ulDi As Long
    ulDo As Long
End Type

' -----------------------------------------------------------------------
'  Output range configurations structure for each channel (for PCI-3305)
' -----------------------------------------------------------------------
Type DAMODECHREQ
    ulRange As Long
    fVolt As Single
    ulFilter As Long
End Type

' -----------------------------------------------------------------------
'  Waveform generation mode structure (for PCI-3305)
' -----------------------------------------------------------------------
Type DAMODEREQ
    ModeChReq(0 To 1) As DAMODECHREQ
    ulPulseMode As Long
    ulSyntheOut As Long
    ulInterval As Long
    fIntervalCycle As Single
    ulCounterClear As Long
    ulDaLatch As Long
    ulSamplingClock As Long
    ulExControl As Long
    ulExClock As Long
End Type

'-----------------------------------------------------------------------------------------------
'
'   Error identifier
'
'-----------------------------------------------------------------------------------------------
Public Const DA_ERROR_SUCCESS = 0
Public Const DA_ERROR_NOT_DEVICE = &HC0000001
Public Const DA_ERROR_NOT_OPEN = &HC0000002
Public Const DA_ERROR_INVALID_HANDLE = &HC0000003
Public Const DA_ERROR_ALREADY_OPEN = &HC0000004
Public Const DA_ERROR_NOT_SUPPORTED = &HC0000009
Public Const DA_ERROR_NOW_SAMPLING = &HC0001001
Public Const DA_ERROR_STOP_SAMPLING = &HC0001002
Public Const DA_ERROR_START_SAMPLING = &HC0001003
Public Const DA_ERROR_SAMPLING_TIMEOUT = &HC0001004
Public Const DA_ERROR_INVALID_PARAMETER = &HC0001021
Public Const DA_ERROR_ILLEGAL_PARAMETER = &HC0001022
Public Const DA_ERROR_NULL_POINTER = &HC0001023
Public Const DA_ERROR_SET_DATA = &HC0001024
Public Const DA_ERROR_FILE_OPEN = &HC0001041
Public Const DA_ERROR_FILE_CLOSE = &HC0001042
Public Const DA_ERROR_FILE_READ = &HC0001043
Public Const DA_ERROR_FILE_WRITE = &HC0001044
Public Const DA_ERROR_INVALID_DATA_FORMAT = &HC0001061
Public Const DA_ERROR_INVALID_AVERAGE_OR_SMOOTHING = &HC0001062
Public Const DA_ERROR_INVALID_SOURCE_DATA = &HC0001063
Public Const DA_ERROR_NOT_ALLOCATE_MEMORY = &HC0001081
Public Const DA_ERROR_NOT_LOAD_DLL = &HC0001082
Public Const DA_ERROR_CALL_DLL = &HC0001083



'-------------------------------------------------
'
'   Function declaration
'
'-------------------------------------------------
Declare Function DaOpen Lib "FbiDa.DLL" (ByVal lpszName As String) As Long
Declare Function DaClose Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DaGetDeviceInfo Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pDaBoardSpec As DABOARDSPEC) As Long
Declare Function DaSetBoardConfig Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal ulSmplBufferSize As Long, ByVal hEvent As Long, ByVal lpEventProc As Long, ByVal dwUser As Long) As Long
Declare Function DaGetBoardConfig Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef ulSmplBufferSize As Long, ByRef ulDaSmplEventFactor As Long) As Long
Declare Function DaSetSamplingConfig Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pDaSmplConfig As DASMPLREQ) As Long
Declare Function DaGetSamplingConfig Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pDaSmplConfig As DASMPLREQ) As Long
Declare Function DaSetMode Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pDaMode As DAMODEREQ) As Long
Declare Function DaGetMode Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pDaMode As DAMODEREQ) As Long
Declare Function DaSetSamplingData Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pSmplData As Any, ByVal dwSmplDataNum As Long) As Long
Declare Function DaClearSamplingData Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DaStartSampling Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal fdwSyncFlag As Long) As Long
Declare Function DaStartFileSampling Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal pszPathName As String, ByVal ulFileFlag As Long, ByVal ulSmplNum As Long) As Long
Declare Function DaSyncSampling Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal ulMode As Long) As Long
Declare Function DaStopSampling Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long) As Long
Declare Function DaGetStatus Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef ulDaSmplStatus As Long, ByRef ulDaSmplCount As Long, ByRef ulDaAvailCount As Long, ByRef ulDaAvailRepeat As Long) As Long
Declare Function DaOutputDA Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal nCh As Long, ByRef lpDaSmplChReq As DASMPLCHREQ, ByRef pData As Any) As Long
Declare Function DaInputDI Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByRef pdwData As Long) As Long
Declare Function DaOutputDO Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal dwData As Long) As Long
Declare Function DaAdjustVR Lib "FbiDa.DLL" (ByVal hDeviceHandle As Long, ByVal ulAdjustCh As Long, ByVal ulSelVolume As Long, ByVal ulControl As Long, ByVal v As Long) As Long
Declare Function DaDataConv Lib "FbiDaDc.DLL" (ByVal uSrcFormCode As Long, ByRef pSrcData As Any, ByVal uSrcSmplDataNum As Long, ByRef pSrcSmplReq As DASMPLREQ, ByVal uDestFormCode As Long, ByRef pDestData As Any, ByRef puDestSmplDataNum As Long, ByRef pDestSmplReq As DASMPLREQ, ByVal uEffect As Long, ByVal uCount As Long, ByVal lpfnConv As Long) As Long
Declare Function DaWriteFile Lib "FbiDaDc.DLL" (ByVal pszPathName As String, ByRef pSmplData As Any, ByVal ulFormCode As Long, ByVal ulSmplNum As Long, ByVal ulChCount As Long) As Long

Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal ManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long


Public Sub DvcDaOpen()
    Dim i As Long
    Dim nRet As Long
    Dim lpszName As String
    Dim hHandle As Long
    
    ' Open the device specified by "FBIDA?"
    For i = 0 To 256
        lpszName = "FBIDA"
        lpszName = lpszName & (i + 1)
        hHandle = DaOpen(lpszName)
        If hHandle <> -1 Then
            DaClose (hHandle)
            szName = lpszName
            'DeviceOpen.DeviceName.Text = lpszName
        End If
    Next i
End Sub
Public Sub DvcDaClose()
   Dim nRet As Long
    
    If ghOpenFlagDA = 1 Then
        nRet = DaClose(ghDeviceHandleDa)
        If nRet <> DA_ERROR_SUCCESS Then
            Call DsplyErrMessageDA(nRet)
        Else
            'Call DsplyErrMessageDA(nRet)
            ghOpenFlagDA = 0
            'admain.IDM_DEVICE_OPEN.Enabled = True
        End If
    End If
End Sub

Public Sub DeviceDaName()
Dim nRet As Long
    Dim lpszName As String
    Dim hHnadle As Long

    'lpszName = DeviceName.Text  'FBIDA1
    lpszName = "FBIDA1"
    
    hHandle = DaOpen(lpszName)
    If hHandle = -1 Then
        Call DsplyErrMessageDA(hHandle)
        Exit Sub
    Else
        'Call DsplyErrMessageDA(0)
    
        ghDeviceHandleDa = hHandle
        ghChannelDA(0) = 1
        ghChannelDA(1) = 2
        ghOpenFlagDA = 1
                    
        'admain.IDM_DEVICE_OPEN.Enabled = False
        
        ' Read the DA configuration information.
        nRet = DaGetSamplingConfig(ghDeviceHandleDa, gConfigDA)
        If nRet <> DA_ERROR_SUCCESS Then
            DaClose (ghDeviceHandleDa)
            Call DsplyErrMessageDA(nRet)
            Exit Sub
        End If

        ' Read the device information.
        nRet = DaGetDeviceInfo(ghDeviceHandleDa, gInfoDA)
        If nRet <> DA_ERROR_SUCCESS Then
            DaClose (ghDeviceHandleDa)
            Call DsplyErrMessageDA(nRet)
            Exit Sub
        End If

    End If
End Sub

Public Sub DaOut(ch%, dt$)
    Dim nRet As Long
    Dim SmplChInf(0 To 15) As DASMPLCHREQ
    Dim wData(0 To 15) As Integer
    Dim ulCh As Long
    '
    ' Retrieve a channel number
    If ch = 0 Then
        nRet = MsgBox("Invalid channel", (vbOKOnly + vbCritical), "Error code")
        Exit Sub
    End If
    
    ghChannelDA(0) = ch
    'ghChannelDA(1) = Val(ch(1).Text)
    
    ' Setup the output conditions.
    SmplChInf(0).ulChNo = ghChannelDA(0)
    SmplChInf(0).ulRange = gConfigDA.SmplChReq(0).ulRange

    ' Configure the output data
    'wData(0) = Val("&H" + txtData(0).Text)
    wData(0) = Val("&H" + dt)
    
    'If ghChannelDA(1) = 0 Then
        ulCh = 1
    'Else
    '    ulCh = 2
    '    SmplChInf(1).ulChNo = ghChannelDA(1)
    '    SmplChInf(1).ulRange = gConfigDA.SmplChReq(0).ulRange
    '    wData(1) = Val("&H" + txtData(1).Text)
    'End If
    
    ' Output one sample
    nRet = DaOutputDA(ghDeviceHandleDa, ulCh, SmplChInf(0), wData(0))
    
    'If nRet <> DA_ERROR_SUCCESS Then
    '    Call DsplyErrMessageDA(nRet)
    'Else
    '    nRet = MsgBox("The DA conversion output is completed successfully. [ DaOutputDA ]", vbInformation)
    'End If
    
End Sub

Public Sub DaVoltOut(ch%, V1!)
Dim nRet As Long
Dim SmplChInf(0 To 15) As DASMPLCHREQ
Dim wData(0 To 15) As Integer
Dim ulCh As Long
Dim vdt$
Dim v!
    '
'    If ch = 1 Then             ' ひとつ上の階層のルーチンへ移動
'        v = V1 * gDirect      'S.Mの回転方向 (+1 or -1)
'    Else
'        v = V1
'    End If
    '
  v = V1
  If BrdFlg <> "ON" Then Exit Sub
  If v > 10 Then v = 10
  If v < -10 Then v = -10
  vdt = Hex(v / 10 * &H800 + &H7FF)
  DaOut ch, vdt
  
End Sub
Public Sub TempSet(ch%, temp!)
'　080313　変更　NQD対応　　３，４出力増加
Dim v!, k!, minT!, maxT!
'----------------------- 温度設定
'Ch=2　成形室　IHヒーター
'Ch=3　上軸ヒーター
'Ch=4　下軸ヒーター
  minT = 0: maxT = 1000
  k = 10 / (maxT - minT) * 1.0008 '10V / (MaxTemp-MinTemp)
  v = (temp + minT + 1) * k
  DaVoltOut ch, v
End Sub
Public Sub DaOut1(ch%, dt%)
    Dim nRet As Long
    Dim SmplChInf(0 To 15) As DASMPLCHREQ
    Dim wData(0 To 15) As Integer
    Dim ulCh As Long
    '
    If ch = 0 Then
        nRet = MsgBox("Invalid channel", (vbOKOnly + vbCritical), "Error code")
        Exit Sub
    End If
    
    ghChannelDA(0) = ch
    SmplChInf(0).ulChNo = ghChannelDA(0)
    SmplChInf(0).ulRange = gConfigDA.SmplChReq(0).ulRange

    wData(0) = dt     'Val("&H" + dt)
    
    ulCh = 1

    nRet = DaOutputDA(ghDeviceHandleDa, ulCh, SmplChInf(0), wData(0))
 
End Sub
