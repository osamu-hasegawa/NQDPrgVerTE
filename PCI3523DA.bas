Attribute VB_Name = "PCI3523DA"


' -----------------------------------------------------------------------
'  Global variables
' -----------------------------------------------------------------------
'----------------------------- DA
Public ghDeviceHandleDa As Long       ' Device handle
Public ghChannelDA(0 To 2) As Long    ' Channel number
Public ghOpenFlagDA As Long           ' Open flag
Public gConfigDA As DASMPLREQ         ' Structure to store the analog output configurations and conditions.
Public gInfoDA As DABOARDSPEC         ' Structure to store the device information.

Public szTempDA As String
'------------------------------ AD
Public ghDeviceHandleAD As Long       ' Device handle
Public ghChannelAD(0 To 2) As Long    ' Channel number
Public ghOpenFlagAD As Long           ' Open flag
Public gConfigAD As DASMPLREQ         ' Structure to store the analog output configurations and conditions.
Public gInfoAD As DABOARDSPEC         ' Structure to store the device information.

Public szTempAD As String
Public lpszNameAD As String




Public Sub DsplyErrMessageDA(ByVal uErrCode As Long)
    Select Case uErrCode
        Case DA_ERROR_SUCCESS
            nRet = MsgBox("The process is completed successfully.", (vbOKOnly + vbInformation), "Error code")
        Case DA_ERROR_NOT_DEVICE
            nRet = MsgBox("The device couldn't be found.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NOT_OPEN
            nRet = MsgBox("The system couldn't found the device.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_INVALID_HANDLE
            nRet = MsgBox("Invalid device handle is specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_ALREADY_OPEN
            nRet = MsgBox("The device has been already opened.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NOT_SUPPORTED
            nRet = MsgBox("It is not supported.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NOW_SAMPLING
            nRet = MsgBox("The analog output is running.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_STOP_SAMPLING
            nRet = MsgBox("The analog output is stopped.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_START_SAMPLING
            nRet = MsgBox("The analog output couldn't be performed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_SAMPLING_TIMEOUT
            nRet = MsgBox("The timeout interval elapsed while the analog output is running.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_INVALID_PARAMETER
            nRet = MsgBox("Invalid parameters are specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_ILLEGAL_PARAMETER
            nRet = MsgBox("Invalid analog output conditions are specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NULL_POINTER
            nRet = MsgBox("A NULL pointer is specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_SET_DATA
            nRet = MsgBox("The analog output data couldn't be configured.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_FILE_OPEN
            nRet = MsgBox("Opening the file failed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_FILE_CLOSE
            nRet = MsgBox("Closing the file failed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_FILE_READ
            nRet = MsgBox("Reading the file failed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_FILE_WRITE
            nRet = MsgBox("Writing the file failed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_INVALID_DATA_FORMAT
            nRet = MsgBox("Invalid data format is specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_INVALID_AVERAGE_OR_SMOOTHING
            nRet = MsgBox("Invalid averaging configuration or invalid smoothing configuration is specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_INVALID_SOURCE_DATA
            nRet = MsgBox("Invalid source data are specified.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NOT_ALLOCATE_MEMORY
            nRet = MsgBox("Enough memory couldn't be allocated.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_NOT_LOAD_DLL
            nRet = MsgBox("Loading the DLL failed.", (vbOKOnly + vbCritical), "Error code")
        Case DA_ERROR_CALL_DLL
            nRet = MsgBox("Calling the DLL failed.", (vbOKOnly + vbCritical), "Error code")
        Case Else
            nRet = MsgBox("Unexpected error is occurred.", (vbOKOnly + vbCritical), "Error code")
    End Select
End Sub



'Public Function StrFormat(strData As String, DataLen As Long, FormatData As String)
'    Select Case DataLen
'        Case 1
'            FormatData = "000" & strData
'        Case 2
'            FormatData = "00" & strData
'        Case 3
'            FormatData = "0" & strData
'        Case 4
'            FormatData = strData
'    End Select
'End Function


