Attribute VB_Name = "ScreenCopy"
'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�@�ϐ��錾���@�@�i273�j '

Private Declare Sub keybd_event Lib "user32.dll" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_SNAPSHOT = &H2C            'PrintScreen �L�[(P1051)
Private Const VK_LMENU = &HA4               'Alt�L�[
Private Const KEYEVENTF_KEYUP = &H2         '�L�[�̓A�b�v���
Private Const KEYEVENTF_EXTENDEDKEY = &H1   '�X�L�����R�[�h�͊g���R�[�h



'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�{�́@�@�@�@�@�i273�j '

Private Sub SaveWindowPic(Optional ActWind As Boolean = True, _
                                    Optional PrintOn As Boolean = False)
'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�@�@�@�@�@�@�@�@�i273�j '
'�t�H�[����Command�{�^�����Q�\��t���Ă����ĉ������B
'�@ Option Explicit�@�@ 'SampleNo=273�@WindowsXP VB6.0(SP5) 2003.03.30
'�L�[�X�g���[�N���V�~�����[�g����(P1065)

    Dim MyFileName As String, PicData As Picture, OsVer As Single
    Dim sngSt As Single
'
    Clipboard.Clear
    OsVer = CreateObject("SysInfo.SYSINFO").OSVersion

    If ActWind Then
    '�A�N�e�B�u �E�B���h�E�̃X�i�b�v�V���b�g���擾����
    '�ȉ��̂Q���@�ǂ�ł�OK(Win98SE/WinXP/Win95�j
    '�ǂ̕��@�ł���L�m�F�@��͓������삵�܂��̂�MS�̃T���v���̕��@���g�p
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'�@�@�@�@==================== ������ł������悤�ł� ==================
'�@�@�@�@Call keybd_event(VK_LMENU, 0, _
�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or 0, 0)
'�@�@�@�@Call keybd_event(VK_SNAPSHOT, 0, _
�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or 0, 0)
'�@�@�@�@Call keybd_event(VK_SNAPSHOT, 0, _
�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'�@�@�@�@Call keybd_event(VK_LMENU, 0, _
�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    ElseIf ActWind = False And OsVer < 5 Then
    '��ʑS�̂̃X�i�b�v�V���b�g���擾����(Win98SE/Win95)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    Else
    '��ʑS�̂̃X�i�b�v�V���b�g���擾����(WinXP)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    End If

    sngSt = Timer                           ' Windows7 �ɂ́A���̒x��Loop���K�v
    Do While Timer - sngSt < 0.5
       DoEvents
    Loop
    '�N���b�v�{�[�h���Ƀr�b�g�}�b�v�`���̃f�[�^�����邩���ׂ�
    If Clipboard.GetFormat(vbCFBitmap) Then
        '�t�@�C��������������
        MyFileName = App.path & "\" & gcoxFlName$ & Format$(Now, "yymmddhhmmss") & ".BMP"
        '�\���f�[�^�[���r�b�g�}�b�v�`���̃f�[�^�ŕۑ�
        Set PicData = Clipboard.GetData
        Call SavePicture(PicData, MyFileName)
        If PrintOn Then
            '�������ꍇ
            With Printer
                .ScaleMode = vbMillimeters
                .PaperSize = vbPRPSA4
                .Orientation = vbPRORLandscape
                .PaintPicture PicData, 10, 0
                .EndDoc
            End With
        End If
    Else
        MsgBox "�ۑ��o���܂���ł����B"
    End If
End Sub
'
'
'
'Private Sub Command1_Click()
''�A�N�e�B�u�E�C���h�E�݂̂��N���b�v�{�[�h�ɃR�s�[
'    Call SaveWindowPic(True, False)     '�������ꍇ�́@True �ɐݒ�
'End Sub
'
'Private Sub Command2_Click()
''�X�N���[���S�̂��N���b�v�{�[�h�ɃR�s�[
'    Call SaveWindowPic(False, False)
'End Sub




NQD�@Vb�̏ꍇ
��NQD70_SC���֒ǉ�
�t���O�̒ǉ��@�@�`���̐錾��
Dim iflgSCopy As Boolean   ' ScreenCopy �t���O

���R�}���h�{�^���̒ǉ���
Private Sub Command3_Click()
'''�A�N�e�B�u�E�C���h�E���N���b�v�{�[�h�ɃR�s�[�������B�@True �ɐݒ�
  If iflgSCopy = True Then
          iflgSCopy = False          'ScreenCopy�@��t����
          Command3.BackColor = CmndColoff(0)
    Else
          iflgSCopy = True      'ScreenCopy�@��t
          Command3.BackColor = CmndColon(1)    ' on 1=red
  End If
End Sub

<NQD70_SC�̖{�̂ւ́@call���ǉ���
>'/* �f�[�^�̕ۑ��@*/
>      If lDtSaveFlg = True Then
>        ResDtSave i_s, stime
>        lDtSaveFlg = False
>      End If
'
' ScreenCopy iflgSCopy=True �̏ꍇ�AScreenCopy
    If iflgSCopy = True Then
        Call SaveWindowPic(True, False)     'Active Window�̕ۑ�
    End If
    iflgSCopy = False          'ScreenCopy�@��t����
    Command3.BackColor = CmndColoff(0)
'
