Attribute VB_Name = "MplBDef"
'*************************************************************************************************************************
'　ＭＰＬ 構造体、関数定義
'*************************************************************************************************************************

'**** RESULT 構造体 ****
Type MPL_S_RESULT
   MPL_Result(1 To 4) As Integer
End Type

'**** DATA 構造体 ****
Type MPL_S_DATA
   MPL_Data(1 To 4) As Integer
End Type

'**** 定数定義 ****
Public Const MPL_X As Integer = 0
Public Const MPL_Y As Integer = 1
Public Const MPL_Z As Integer = 2
Public Const MPL_A As Integer = 3
Public Const MPL_B As Integer = 4
Public Const MPL_C As Integer = 5

'**** 関数定義 ****
Declare Function MPL_BOpen Lib "MplB.dll" (ByVal hWnd As Long, ByVal BoardNo As Integer, ByVal Axis As Integer, phDev As Long, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BClose Lib "MplB.dll" (ByVal hDev As Long, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_IWDrive Lib "MplB.dll" (ByVal hDev As Long, ByVal cmd As Integer, psData As MPL_S_DATA, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWDriveCommand Lib "MplB.dll" (ByVal hDev As Long, pCmd As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWDriveData1 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWDriveData2 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWDriveData3 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRStatus1 Lib "MplB.dll" (ByVal hDev As Long, pStatus As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRStatus2 Lib "MplB.dll" (ByVal hDev As Long, pStatus As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRStatus3 Lib "MplB.dll" (ByVal hDev As Long, pStatus As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRStatus4 Lib "MplB.dll" (ByVal hDev As Long, pStatus As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRStatus5 Lib "MplB.dll" (ByVal hDev As Long, pStatus As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_IRDrive Lib "MplB.dll" (ByVal hDev As Long, psData As MPL_S_DATA, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRDriveData1 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRDriveData2 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BRDriveData3 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWaitDriveCommand Lib "MplB.dll" (ByVal hDev As Long, ByVal WaitTime As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BIsWait Lib "MplB.dll" (ByVal hDev As Long, pWait As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BBreakWait Lib "MplB.dll" (ByVal hDev As Long, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_IWCounter Lib "MplB.dll" (ByVal hDev As Long, ByVal cmd As Integer, psData As MPL_S_DATA, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWCounterCommand Lib "MplB.dll" (ByVal hDev As Long, pCmd As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWCounterData1 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWCounterData2 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BWCounterData3 Lib "MplB.dll" (ByVal hDev As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BPortOpen Lib "MplB.dll" (ByVal hWnd As Long, ByVal BoardNo As Integer, phPort As Long, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BPortClose Lib "MplB.dll" (ByVal hPort As Long, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BPortIn Lib "MplB.dll" (ByVal hPort As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_BPortOut Lib "MplB.dll" (ByVal hPort As Long, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_Inp Lib "MplB.dll" (ByVal BoardNo As Integer, pData As Integer, psResult As MPL_S_RESULT) As Boolean
Declare Function MPL_Outp Lib "MplB.dll" (ByVal BoardNo As Integer, pData As Integer, psResult As MPL_S_RESULT) As Boolean

