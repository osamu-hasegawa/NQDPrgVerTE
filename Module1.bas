Attribute VB_Name = "Module1"

' -----------------------------------------------------------------------
'  Global variable
' -----------------------------------------------------------------------

Public ghDeviceHandle As Long       ' Device handle
Public ghChannel(0 To 1) As Long    ' Channel number
Public ghOpenFlag As Long           ' Open flag
Public gConfig As ADSMPLREQ         ' Sampling request condition structure
Public gInfo As ADBOARDSPEC         ' Board information structure

Public szTemp As String

Public Sub DsplyErrMessage(ByVal uErrCode As Long)
    Select Case uErrCode
        Case AD_ERROR_SUCCESS
            nRet = MsgBox("Completed successfully.", (vbOKOnly + vbInformation), "Error Code")
        Case AD_ERROR_NOT_DEVICE
            nRet = MsgBox("The specified device could not be found.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_NOT_OPEN
            nRet = MsgBox("A device can not be opened.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_INVALID_HANDLE
            nRet = MsgBox("Invalid device handle", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_ALREADY_OPEN
            nRet = MsgBox("The device has been already opened.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_NOT_SUPPORTED
            nRet = MsgBox("It is not supported.", (vbOKOnly + vbCritical), "Error")
        
        Case AD_ERROR_NOW_SAMPLING
            nRet = MsgBox("The sampling is running.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_STOP_SAMPLING
            nRet = MsgBox("The sampling has been stopped.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_START_SAMPLING
            nRet = MsgBox("The sampling could not be done.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_SAMPLING_TIMEOUT
            nRet = MsgBox("The timeout interval elapsed while a sampling is running.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_INVALID_PARAMETER
            nRet = MsgBox("Invalid parameters", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_ILLEGAL_PARAMETER
            nRet = MsgBox("Sampling request condition is invalid.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_NULL_POINTER
            nRet = MsgBox("A null pointer is specified.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_GET_DATA
            nRet = MsgBox("Retrieving Sampling Data failed.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_FILE_OPEN
            nRet = MsgBox("Opening a file failed.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_FILE_CLOSE
            nRet = MsgBox("Closing a file faild.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_FILE_READ
            nRet = MsgBox("Reading from a file failed.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_FILE_WRITE
            nRet = MsgBox("Writing to a file failed.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_INVALID_DATA_FORMAT
            nRet = MsgBox("Invalid data format", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_INVALID_AVERAGE_OR_SMOOTHING
            nRet = MsgBox("Invalid averaging or smoothing is specified. ", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_INVALID_SOURCE_DATA
            nRet = MsgBox("Invalid source data to convert.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_NOT_ALLOCATE_MEMORY
            nRet = MsgBox("Memories could not be allocated.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_NOT_LOAD_DLL
            nRet = MsgBox("Loading a DLL failed.", (vbOKOnly + vbCritical), "Error")
        Case AD_ERROR_CALL_DLL
            nRet = MsgBox("Calling a DLL failed.", (vbOKOnly + vbCritical), "Error")
        Case Else
            nRet = MsgBox("Not-expected error", (vbOKOnly + vbCritical), "Error Code")
    End Select
End Sub


Public Function StrFormat(strData As String, DataLen As Long, FormatData As String)
    Select Case DataLen
        Case 1
            FormatData = "000" & strData
        Case 2
            FormatData = "00" & strData
        Case 3
            FormatData = "0" & strData
        Case 4
            FormatData = strData
    End Select
End Function

