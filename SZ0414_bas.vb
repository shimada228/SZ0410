Option Strict Off
Option Explicit On
Module SZ0414BAS
	'******************************************************************
	'*    �V�X�e����    �F  MKK�d���Ǘ��V�X�e��            �@           *
	'*    �v���O������  �F  �i�`�m�}�X�^����
	'*    �v���O�����h�c�F  SZ0414
	'*    ��  ��  ��   �F   SSP@MEGURO
	'******************************************************************
	'*  �R���p�C����        �F2013/04/01
	'*  �ύX�L�[            �F20130401
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.���Y���͔��p�p���̑啶���݂̂Ƃ���
	'*  �C�����e            �F2.�u�����Z�v�̃��x���\����ǉ�
	'******************************************************************'A-20130401-
	'*  �R���p�C����        �F2013/04/24
	'*  �ύX�L�[            �F20130424
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.���i����K17�ɕύX����
	'******************************************************************'A-20130424-
	'*  �R���p�C����        �F2013/05/10
	'*  �ύX�L�[            �F20130510
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.ZAFC_MST��ZAFC_USE�̐ݒ�͌Ăяo������ʂōs��
	'*  �C�����e            �F�@�ďo�����Őݒ肷��֌W�ŉ�ʂɎg��Ȃ����A
	'*  �C�����e            �F�@�t�@���N�V�����{�^����z�u����
	'******************************************************************'A-20130510-
	
	Structure RECT
		'UPGRADE_NOTE: Left �� Left_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right �� Right_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Public InitFlg As Boolean
	Public lpRectSave As RECT
	Public vfaRowWidth() As Integer
	Public tabNo As Short
	
	'Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
	''        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
	''        ByVal nHeight As Long, ByVal bRepaint As Long) As Long
	'
	'Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
	''        lpRect As RECT) As Long
	
	'�q�c�n�֘A���[�N
	'Public RdoEnv  As rdoEnvironment        ' rdo�����
	Public SZ0414SELGE As RDO.rdoQuery
	Public SZ0414SELGT As RDO.rdoQuery
	Public SZ0414SELLT As RDO.rdoQuery
	Public SZ0414RES As RDO.rdoResultset
	Public SZ0415_JAN_BUNRUISEL As RDO.rdoQuery '
	
	Public SZ_ERRSW As Short ' ERROR ����SW=1
	Public SZ_INTSW As Short
	Public SZ_F3SW As Short
	
	Public SZ0414_SELCOD1 As New VB6.FixedLengthString(13) ' �����J�n�R�[�h�^�I���R�[�h�i�Ǝ҃R�[�h�j
	Public SZ0414_KBN As Short ' �L���敪 0:�L�� -1:����
	
	'Public ZAFC_MST(1 To 12) As String
	
	Public SZ0414_OLDCOD3 As New VB6.FixedLengthString(13) ' �O��J�n�R�[�h
	Public SZ0414_HDISP As Short ' ���o���ĕ\��
	
	Public SZ0414_IMTX010 As New VB6.FixedLengthString(13) ' �O��i�`�m�R�[�h�@�J�n
	Public SZ0414_IMTX020 As New VB6.FixedLengthString(13) ' �O��i�`�m�R�[�h�@�I��
	Public SZ0414_IMTX030 As New VB6.FixedLengthString(6) ' �O��i�`�m���i���ށ@�J�n
	Public SZ0414_IMTX040 As New VB6.FixedLengthString(6) ' �O��i�`�m���i���ށ@�I��
	Public SZ0414_IMTX050 As New VB6.FixedLengthString(3) ' �O�񌴎Y��
	Public SZ0414_IMNU060 As Decimal ' �O��d�ʁ@�J�n
	Public SZ0414_IMNU070 As Decimal ' �O��d�ʁ@�I��
	Public SZ0414_IMTX080 As New VB6.FixedLengthString(1) ' �O��ܖ������@�敪
	Public SZ0414_IMNU090 As Decimal ' �O��ܖ������@�J�n
	Public SZ0414_IMNU100 As Decimal ' �O��ܖ������@�I��
	Public SZ0414_IMNU090D As Decimal ' �O��ܖ������@�J�n(�����Z)
	Public SZ0414_IMNU100D As Decimal ' �O��ܖ������@�I��(�����Z)
	
	Public SZ0414_SPRD As Short ' �O���گ�ޱ�è��ROW
	Public SZ0414_LNCNT As Decimal ' �O�s�ԍ�
	
	
	Public SZ0414_PRESW As Short ' PREPARE���f�p����
	
	Public SZ0414_TOPS As Integer ' �e���(TOP)
	Public SZ0414_LEFTS As Integer ' �e���(LEFT)
	Public SZ0414_HEIGHTS As Integer ' �e���(HEIGHT)
	Public SZ0414_WIDTHS As Integer ' �e���(WIDTH)
	Public SZ0414_PS As Short ' �\���ʒu(0.���� 1.���� 2.�E�� 3.���� 4.�E�� )
	Public SZ0414_TIMES As Integer ' RDO��ѱ�ĕb��
	Public SZ0414_KAISYAS As String ' ��ЃR�[�h
	Public SZ0414_HONSITENS As String ' �{�x�X�R�[�h
	''�V�X�e���c�`�s�d
	'Public SYSDATE As String * 8
	''�V�X�e���c�`�s�d�E�����l
	'Public SYSDATES As String * 8
	''�V�X�e���c�`�s�d�E�O���l
	'Public SYSDATEO As String * 6
	'
	''SWITCH ���
	'Public MOUSEFLG As Integer
	
	'�r�v�h�s�b�g�G���A
	Public SZ0414_DSPSW As Short
	'Public DSP1SW As Integer
	
	''�X�C�b�`�@�I�[�v��
	'Public PRNOPNSW As Integer
	
	''129�G���[���o�p
	'Public ERR129_SW As Boolean
	'
	''�K�C�h�\���p
	'Public ZAGD_MST(1 To 5) As String
	
	'�j�a�G���A
	'�������̓G���A
	Structure SZ0414KB_S
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(13),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=13)> Public S010() As Char '�i�`�m�R�[�h�@�J�n
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(13),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=13)> Public S020() As Char '�i�`�m�R�[�h�@�I��
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(6),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=6)> Public S030() As Char '�i�`�m���i���ށ@�J�n
		Dim S030N As String '�i�`�m���i���ޖ��@�J�n
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(6),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=6)> Public S040() As Char '�i�`�m���i���ށ@�I��
		Dim S040N As String '�i�`�m���i���ޖ��@�I��
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(3),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=3)> Public S050() As Char '���Y��
		Dim C060 As Decimal '�d�ʁ@�J�n
		Dim C070 As Decimal '�d�ʁ@�I��
		'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ���N���b�N���Ă��������B
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public S080() As Char '�ܖ������@�敪
		Dim C090 As Decimal '�ܖ������@�J�n
		Dim C100 As Decimal '�ܖ������@�I��
	End Structure
	Public WKBSZ0414 As SZ0414KB_S
	Public KBSZ0414 As SZ0414KB_S
	
	'''�R���g���[���Ǘ��p
	''Type CTRLTBL_S
	''    IGRP    As Integer
	''    ISGRP   As Integer
	''    INEXT   As Integer
	''    IBACK   As Integer
	''    IDOWN   As Integer
	''    CTRL    As Control
	''End Type
	''
	'''�O���[�v�`�F�b�N�p
	''Type GRPTBL_S
	''    CFLG    As Integer
	''    NXTN    As Integer
	''End Type
	''�ő�\���p
	'Public ChMax    As Boolean
	
	Public SZ0414_TOP As Integer ' �e���(TOP)
	Public SZ0414_LEFT As Integer ' �e���(LEFT)
	Public SZ0414_HEIGHT As Integer ' �e���(HIGHT)
	Public SZ0414_WIDTH As Integer ' �e���(WIDTH)
	Public SZ0414_POS As Short ' �\���ʒu(0.���� 1.���� 2.�E�� 3.���� 4.�E�� )
	Public SZ0414_RCN As RDO.rdoConnection ' �ް��ް����
	Public SZ0414_TIME As Integer ' RDO��ѱ�ĕb��
	Public SZ0414_LCODE As String ' �I���R�[�h
	
	Public Function SZ0414_SUB() As Short
		SZ_INTSW = F_OFF ' �����\��RTN���s
		SZ_ERRSW = F_OFF
		
		'    Set ZACN_RCN = SZ0414_RCN
		'    ZACN_DB = SZ0414_DB
		SZ0414_TOPS = SZ0414_TOP ' �e���(TOP)
		SZ0414_LEFTS = SZ0414_LEFT ' �e���(LEFT)
		SZ0414_HEIGHTS = SZ0414_HEIGHT ' �e���(HEIGHT)
		SZ0414_WIDTHS = SZ0414_WIDTH ' �e���(WIDTH)
		SZ0414_PS = SZ0414_POS
		SZ0414_TIMES = SZ0414_TIME
		
		' ��������
		Call INIT_SZ0414_RTN()
		If SZ_ERRSW = F_ERR Then
			SZ0414_SUB = n1
			Exit Function
		End If
		
		SZ0414FRM.ShowDialog()
		SZ0414_LCODE = SZ0414_SELCOD1.Value
		SZ0414_SUB = SZ0414_KBN
		
	End Function
	
	
	Private Sub ENV_RTN()
		Dim IDX As Short
		Dim INI_NAME As String
		Dim WK_FNM As String
		
		'�i�`�m�}�X�^
		MKKCMN.ZAEV_FNO = "JAN"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			CM_ERRSW = F_ERR
			Exit Sub
		Else
			JAN_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		'�i�`�m���ރ}�X�^
		MKKCMN.ZAEV_FNO = "JAN_BUNRUI"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			CM_ERRSW = F_ERR
			Exit Sub
		Else
			JAN_BUNRUI_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		
	End Sub
	Private Function INIGET_ENTRY(ByVal section As String, ByVal entry As String, ByVal def_str As String, ByVal fname As String) As String
		
		'   /*                               */
		'   /* INI�t�@�C���̓��e�擾�i�ʁj */
		'   /*     (Internal Function)       */
		'   /*                               */
		
		Static bUF As New VB6.FixedLengthString(256)
		Dim buftmp As String
		
		bUF.Value = ""
		
		'   INI�t�@�C���̎w��G���g�����擾
		If (ARQCNBAS.GetPrivateProfileString(section, entry, def_str, bUF.Value, 256, fname) > 0) Then
			buftmp = Trim(bUF.Value)
		Else
			buftmp = Trim(def_str)
		End If
		
		'   ������̍Ō�� '\0'���t���Ă����Ȃ珜������
		'UPGRADE_ISSUE: RightB$ �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		If (RightB$(buftmp, 2) = Chr(0)) Then
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LeftB$ �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			INIGET_ENTRY = LeftB$(buftmp, LenB(buftmp) - 2)
		Else
			INIGET_ENTRY = buftmp
		End If
	End Function
	
	Public Sub INIT_SZ0414_RTN()
		
		If SZ0414_PRESW = F_ON Then
			SZ_INTSW = F_ON
			Exit Sub
		End If
		
		' �X�L�[�}���擾
		Call ENV_RTN()
		If SZ_ERRSW = F_ERR Then Exit Sub
		
		' Prepare�R�}���h�ݒ�
		Call PREP_JAN_BUNRUI()
		If SZ_ERRSW = F_ERR Then Exit Sub
		
		' �萔�ɒl��ݒ肷��
		' �t�@���N�V�����\���p
		'ZAFC_MST(1) = "�I  ��"
		'ZAFC_MST(2) = ""
		'ZAFC_MST(3) = "�⍇��"
		'ZAFC_MST(4) = ""
		'ZAFC_MST(5) = "�N���A"
		'''    ZAFC_MST(6) = "�O�ꗗ"       'D-20130510-
		'''    ZAFC_MST(7) = "���ꗗ"       'D-20130510-
		'ZAFC_MST(8) = ""
		'ZAFC_MST(9) = ""
		'ZAFC_MST(10) = ""
		'''    ZAFC_MST(11) = "�I  ��"      'D-20130510-
		'ZAFC_MST(12) = "�I  ��"
		
		'D-20130510-S
		'''    ZAFC_USE(0) = True
		'''    ZAFC_USE(1) = False
		'''    ZAFC_USE(2) = False
		'''    ZAFC_USE(3) = True
		'''    ZAFC_USE(4) = False
		'''    ZAFC_USE(5) = True
		'''    ZAFC_USE(6) = True
		'''    ZAFC_USE(7) = True
		'''    ZAFC_USE(8) = False
		'''    ZAFC_USE(9) = False
		'''    ZAFC_USE(10) = False
		'''    ZAFC_USE(11) = False
		'''    ZAFC_USE(12) = True
		'D-20130510-E
		
		
		' ���[�N�̏�����
		SZ0414_SELCOD1.Value = ""
		SZ0414_OLDCOD3.Value = ""
		SZ0414_IMTX010.Value = ""
		SZ0414_IMTX020.Value = ""
		SZ0414_IMTX030.Value = ""
		SZ0414_IMTX040.Value = ""
		SZ0414_IMTX050.Value = ""
		SZ0414_IMNU060 = 0
		SZ0414_IMNU070 = 0
		SZ0414_IMTX080.Value = ""
		SZ0414_IMNU090 = 0
		SZ0414_IMNU100 = 0
		SZ0414_IMNU090D = 0
		SZ0414_IMNU100D = 0
		
		SZ0414_HDISP = F_ON
		SZ0414_PRESW = F_ON
		SZ0414_DSPSW = F_OFF
		SZ_INTSW = F_OFF
		
	End Sub
	Public Sub PREP_JAN_BUNRUI()
		
		SQL = "Select "
		SQL = SQL & " NVL(BK1, ' ') BK1"
		SQL = SQL & ",NVL(BK2, ' ') BK2"
		SQL = SQL & ",NVL(BK3, ' ') BK3"
		SQL = SQL & ",NVL(BK4, ' ') BK4"
		SQL = SQL & ",NVL(BK5, ' ') BK5"
		SQL = SQL & ",NVL(BK6, ' ') BK6"
		SQL = SQL & ",NVL(BK7, ' ') BK7"
		SQL = SQL & ",NVL(BK8, ' ') BK8"
		SQL = SQL & " From "
		SQL = SQL & RTrim(JAN_BUNRUI_FILE.NAME) & "JAN_BUNRUI "
		SQL = SQL & " WHERE BK1 = ? "
		SQL = SQL & "   AND BK2 = '4' "
		
		On Error Resume Next
		SZ0415_JAN_BUNRUISEL = ZACN_RCN.CreateQuery("SZ0415_JAN_BUNRUISEL", SQL)
		SZ0415_JAN_BUNRUISEL.QueryTimeout = SZ0415_TIMES
		If B_STATUS <> 0 Then
			GoTo PREP_JAN_BUNRUI_ERR
		End If
		On Error GoTo 0
		
		SZ0415_JAN_BUNRUISEL.rdoParameters(0).NAME = "BK1"
		SZ0415_JAN_BUNRUISEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0415_JAN_BUNRUISEL.rdoParameters(0).Size = 6
		'    SZ0415_JAN_BUNRUISEL(1).Name = "BK2"
		'    SZ0415_JAN_BUNRUISEL(1).Type = rdTypeCHAR
		'    SZ0415_JAN_BUNRUISEL(1).Size = 1
		
		Exit Sub
		
PREP_JAN_BUNRUI_ERR: 
		ZAER_KN = 1
		Call ZAER_SUB()
		SZ_ERRSW = F_ERR
		SZ0414_KBN = -1
		On Error GoTo 0
		
	End Sub
	
	Public Sub PREP_JAN()
		Dim IDX As Short
		Dim SQLWHERE As String
		
		SZ0414_DSPSW = F_ON
		
		
		'*** �R�[�h���f�d�p
		SQL = "SELECT "
		SQL = SQL & "NVL(JAN.K4,' ')  AS K4, " 'JAN
		SQL = SQL & "NVL(JAN.K21,' ') AS K21," 'JICFS���i����
		'    SQL = SQL & "NVL(JAN.K20,' ') AS K20,"  '�o�n�r���V�[�g���i�����j
		SQL = SQL & "NVL(JAN.K17,' ') AS K17," '�`�[�p���i����(��)
		SQL = SQL & "NVL(JAN.K44,' ') AS K44," '���Y���R�[�h
		SQL = SQL & "NVL(JAN.K42,0)   AS K42," '�P�i�d��
		SQL = SQL & "NVL(JAN.K57,' ') AS K57," '�L�����ԋ敪�i�ܖ����ԋ敪�j
		SQL = SQL & "NVL(JAN.K58,0)   AS K58," '�L�����ԁi�ܖ����ԁj
		SQL = SQL & "NVL(JAN.K99,0)   AS K99 " '�ܖ������i�����Z�j
		SQL = SQL & " FROM "
		SQL = SQL & RTrim(JAN_FILE.NAME) & "JAN" & " JAN "
		
		SQLWHERE = ""
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 >= ? "
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 <= ? "
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 >= ? "
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 <= ? "
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K44 = ? "
		End If
		If (SZ0414_IMNU060) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 >= ? "
		End If
		If (SZ0414_IMNU070) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 <= ? "
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K57 = ? "
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 >= ? "
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 <= ? "
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 >= ? "
		End If
		If (SZ0414_IMNU100) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 <= ? "
		End If
		SQL = SQL & SQLWHERE
		SQL = SQL & " ORDER BY JAN.K4 "
		
		On Error Resume Next
		SZ0414SELGE = ZACN_RCN.CreateQuery("SZ0414SELGE", SQL)
		SZ0414SELGE.QueryTimeout = SZ0414_TIMES
		If B_STATUS <> 0 Then
			GoTo PREP_JAN_ERR
		End If
		On Error GoTo 0
		
		IDX = -1
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K4F"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGE.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K4T"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGE.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K21F"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGE.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K21T"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGE.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K44"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGE.rdoParameters(IDX).Size = 3
		End If
		If (SZ0414_IMNU060) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K42F"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU070) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K42T"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        IDX = IDX + 1
		'        SZ0414SELGE(IDX).NAME = "K57"
		'        SZ0414SELGE(IDX).Type = rdTypeCHAR
		'        SZ0414SELGE(IDX).Size = 1
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELGE(IDX).NAME = "K58F"
		'        SZ0414SELGE(IDX).Type = rdTypeNUMERIC
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELGE(IDX).NAME = "K58T"
		'        SZ0414SELGE(IDX).Type = rdTypeNUMERIC
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K99F"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU100) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGE.rdoParameters(IDX).NAME = "K99T"
			SZ0414SELGE.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		
		
		
		'*** �R�[�h���f�s�p�i���ꗗ�j
		SQL = "SELECT "
		SQL = SQL & "NVL(JAN.K4,' ')  AS K4, " 'JAN
		SQL = SQL & "NVL(JAN.K21,' ') AS K21," 'JICFS���i����
		'    SQL = SQL & "NVL(JAN.K20,' ') AS K20,"  '�o�n�r���V�[�g���i�����j
		SQL = SQL & "NVL(JAN.K17,' ') AS K17," '�`�[�p���i����(��)
		SQL = SQL & "NVL(JAN.K44,' ') AS K44," '���Y���R�[�h
		SQL = SQL & "NVL(JAN.K42,0)   AS K42," '�P�i�d��
		SQL = SQL & "NVL(JAN.K57,' ') AS K57," '�L�����ԋ敪�i�ܖ����ԋ敪�j
		SQL = SQL & "NVL(JAN.K58,0)   AS K58," '�L�����ԁi�ܖ����ԁj
		SQL = SQL & "NVL(JAN.K99,0)   AS K99 " '�ܖ������i�����Z�j
		SQL = SQL & " FROM "
		SQL = SQL & RTrim(JAN_FILE.NAME) & "JAN" & " JAN "
		
		SQLWHERE = ""
		SQLWHERE = SQLWHERE & " WHERE JAN.K4  > ?" '
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 >= ? "
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 <= ? "
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 >= ? "
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 <= ? "
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K44 = ? "
		End If
		If (SZ0414_IMNU060) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 >= ? "
		End If
		If (SZ0414_IMNU070) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 <= ? "
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K57 = ? "
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 >= ? "
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 <= ? "
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 >= ? "
		End If
		If (SZ0414_IMNU100) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 <= ? "
		End If
		SQL = SQL & SQLWHERE
		SQL = SQL & " ORDER BY JAN.K4 "
		
		On Error Resume Next
		SZ0414SELGT = ZACN_RCN.CreateQuery("SZ0414SELGT", SQL)
		SZ0414SELGT.QueryTimeout = SZ0414_TIMES
		If B_STATUS <> 0 Then
			GoTo PREP_JAN_ERR
		End If
		On Error GoTo 0
		
		SZ0414SELGT.rdoParameters(0).NAME = "K4"
		SZ0414SELGT.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0414SELGT.rdoParameters(0).Size = 13
		IDX = 0
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K4F"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGT.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K4T"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGT.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K21F"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGT.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K21T"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGT.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K44"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELGT.rdoParameters(IDX).Size = 3
		End If
		If (SZ0414_IMNU060) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K42F"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU070) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K42T"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        IDX = IDX + 1
		'        SZ0414SELGT(IDX).NAME = "K57"
		'        SZ0414SELGT(IDX).Type = rdTypeCHAR
		'        SZ0414SELGT(IDX).Size = 1
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELGT(IDX).NAME = "K58F"
		'        SZ0414SELGT(IDX).Type = rdTypeNUMERIC
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELGT(IDX).NAME = "K58T"
		'        SZ0414SELGT(IDX).Type = rdTypeNUMERIC
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K99F"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU100) <> 0 Then
			IDX = IDX + 1
			SZ0414SELGT.rdoParameters(IDX).NAME = "K99T"
			SZ0414SELGT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		
		
		'*** �R�[�h���k�s�p�i�O�ꗗ�j
		SQL = "SELECT "
		SQL = SQL & "NVL(JAN.K4,' ')  AS K4, " 'JAN
		SQL = SQL & "NVL(JAN.K21,' ') AS K21," 'JICFS���i����
		'    SQL = SQL & "NVL(JAN.K20,' ') AS K20,"  '�o�n�r���V�[�g���i�����j
		SQL = SQL & "NVL(JAN.K17,' ') AS K17," '�`�[�p���i����(��)
		SQL = SQL & "NVL(JAN.K44,' ') AS K44," '���Y���R�[�h
		SQL = SQL & "NVL(JAN.K42,0)   AS K42," '�P�i�d��
		SQL = SQL & "NVL(JAN.K57,' ') AS K57," '�L�����ԋ敪�i�ܖ����ԋ敪�j
		SQL = SQL & "NVL(JAN.K58,0)   AS K58," '�L�����ԁi�ܖ����ԁj
		SQL = SQL & "NVL(JAN.K99,0)   AS K99 " '�ܖ������i�����Z�j
		SQL = SQL & " FROM "
		SQL = SQL & RTrim(JAN_FILE.NAME) & "JAN" & " JAN "
		
		SQLWHERE = ""
		SQLWHERE = SQLWHERE & " WHERE JAN.K4  < ?" '
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 >= ? "
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K4 <= ? "
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 >= ? "
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K21 <= ? "
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K44 = ? "
		End If
		If (SZ0414_IMNU060) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 >= ? "
		End If
		If (SZ0414_IMNU070) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K42 <= ? "
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K57 = ? "
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 >= ? "
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        GoSub WHERE_SET
		'        SQLWHERE = SQLWHERE & " JAN.K58 <= ? "
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 >= ? "
		End If
		If (SZ0414_IMNU100) <> 0 Then
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub WHERE_SET
			SQLWHERE = SQLWHERE & " JAN.K99 <= ? "
		End If
		SQL = SQL & SQLWHERE
		SQL = SQL & " ORDER BY JAN.K4 DESC "
		
		On Error Resume Next
		SZ0414SELLT = ZACN_RCN.CreateQuery("SZ0414SELLT", SQL)
		SZ0414SELLT.QueryTimeout = SZ0414_TIMES
		If B_STATUS <> 0 Then
			GoTo PREP_JAN_ERR
		End If
		On Error GoTo 0
		
		SZ0414SELLT.rdoParameters(0).NAME = "K4"
		SZ0414SELLT.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0414SELLT.rdoParameters(0).Size = 13
		IDX = 0
		If RTrim(SZ0414_IMTX010.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K4F"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELLT.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX020.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K4T"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELLT.rdoParameters(IDX).Size = 13
		End If
		If RTrim(SZ0414_IMTX030.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K21F"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELLT.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX040.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K21T"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELLT.rdoParameters(IDX).Size = 6
		End If
		If RTrim(SZ0414_IMTX050.Value) <> "" Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K44"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR
			SZ0414SELLT.rdoParameters(IDX).Size = 3
		End If
		If (SZ0414_IMNU060) <> 0 Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K42F"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU070) <> 0 Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K42T"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		'    If RTrim$(SZ0414_IMTX080) <> "0" Then
		'        IDX = IDX + 1
		'        SZ0414SELLT(IDX).NAME = "K57"
		'        SZ0414SELLT(IDX).Type = rdTypeCHAR
		'        SZ0414SELLT(IDX).Size = 1
		'    End If
		'    If (SZ0414_IMNU090) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELLT(IDX).NAME = "K58F"
		'        SZ0414SELLT(IDX).Type = rdTypeNUMERIC
		'    End If
		'    If (SZ0414_IMNU100) <> 0 Then
		'        IDX = IDX + 1
		'        SZ0414SELLT(IDX).NAME = "K58T"
		'        SZ0414SELLT(IDX).Type = rdTypeNUMERIC
		'    End If
		If (SZ0414_IMNU090) <> 0 Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K99F"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		If (SZ0414_IMNU100) <> 0 Then
			IDX = IDX + 1
			SZ0414SELLT.rdoParameters(IDX).NAME = "K99T"
			SZ0414SELLT.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		End If
		SZ0414_KBN = 0
		
		Exit Sub
		
PREP_JAN_ERR: 
		ZAER_KN = 1
		ZAER_NO.Value = "JAN"
		Call ZAER_SUB()
		SZ_ERRSW = F_ERR
		SZ0414_KBN = -1
		On Error GoTo 0
		Exit Sub
		
WHERE_SET: 
		If Trim(SQLWHERE) = "" Then
			SQLWHERE = " WHERE "
		Else
			SQLWHERE = SQLWHERE & " AND "
		End If
		'UPGRADE_WARNING: Return �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
		Return 
		
	End Sub
	
	Public Sub RD_JANBUNRUI()
		'
		'
		SZ_ERRSW = F_OFF
		
		SZ0415_JAN_BUNRUISEL.rdoParameters("BK1").Value = JAN_BUNRUI_BUF0.BK1
		''SZ0415_JAN_BUNRUISEL!BK2 = "4"
		
		On Error Resume Next
		JAN_BUNRUIRS = SZ0415_JAN_BUNRUISEL.OpenResultset()
		
		Select Case B_STATUS(JAN_BUNRUIRS)
			Case n0
				JAN_BUNRUIINVSW = F_OFF
				JAN_BUNRUI.BK1 = JAN_BUNRUIRS.rdoColumns("BK1").Value
				JAN_BUNRUI.BK2 = JAN_BUNRUIRS.rdoColumns("BK2").Value
				JAN_BUNRUI.BK3 = JAN_BUNRUIRS.rdoColumns("BK3").Value
				JAN_BUNRUI.BK4 = JAN_BUNRUIRS.rdoColumns("BK4").Value
				JAN_BUNRUI.BK5 = JAN_BUNRUIRS.rdoColumns("BK5").Value
				JAN_BUNRUI.BK6 = JAN_BUNRUIRS.rdoColumns("BK6").Value
				JAN_BUNRUI.BK7 = JAN_BUNRUIRS.rdoColumns("BK7").Value
				JAN_BUNRUI.BK8 = JAN_BUNRUIRS.rdoColumns("BK8").Value
			Case 24
				JAN_BUNRUIINVSW = F_INV
			Case Else
				ZAER_KN = n1
				GoTo RD_COM0010_ERR
		End Select
		On Error GoTo 0
		
		JAN_BUNRUIRS.Close()
		
		Exit Sub
		
RD_COM0010_ERR: 
		ZAER_NO.Value = "JAN"
		Call ZAER_SUB()
		SZ_ERRSW = F_ERR
		On Error GoTo 0
		
	End Sub
End Module