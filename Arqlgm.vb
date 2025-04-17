Option Strict Off
Option Explicit On
Module ARQLGMBAS
	Public ZALGM_INC_CODE As New VB6.FixedLengthString(2) '��к���
	Public ZALGM_JG_CODE As New VB6.FixedLengthString(4) '���Ə��R�[�h
	Public ZALGM_SYS_KBN As New VB6.FixedLengthString(1) '�V�X�e���敪
	Public ZALGM_S_DAY As New VB6.FixedLengthString(8) '�������t
	Public ZALGM_S_TIME As New VB6.FixedLengthString(6) '��������
	Public ZALGM_OP_CODE As New VB6.FixedLengthString(6) '�I�y���[�^�R�[�h
	Public ZALGM_PGID As New VB6.FixedLengthString(8) '�v���O�����h�c�i���p�啶���j
	Public ZALGM_SH_KBN As New VB6.FixedLengthString(1) '�����敪
	Public ZALGM_SH_NAIYO As New VB6.FixedLengthString(30) '�������e
	Public ZALGM_GNFLG As New VB6.FixedLengthString(1) '���z�t���O
	
	Public ZALGM_ERR As New VB6.FixedLengthString(1)
	Public ZALGM_KO_NAIYO As New VB6.FixedLengthString(30) '�X�V���e
	Const ZALGM_ERR_POINT As String = "ZALGM_SUB"
	
	'------------------------------------------------------------
	'�y�֐����z ���O�o�̓T�u���[�`��
	'
	'�y�@  �\�z ���O�t�@�C���Ƀ��O��ǉ�����B�����t�@�C���̑��݂�����̂͗����t�@�C�����X�V����B
	'
	'�y�߂�l�z ����
	'
	'------------------------------------------------------------
	Sub ZALGM_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		
		'<< �G���[�t���O�̃N���A >>
		ZALGM_ERR.Value = "0"
		
		'<< �����t�@�C���̍X�V���� >>
		
		'�����t�@�C���X�V�X�g�A�h�v���V�[�W���|�b�`�k�k
		
		'Update 1999/12/15 REP START TOP
		'�����敪���V�K�o�^�̎��́A�����X�V�s��Ȃ��B
		If ZALGM_SH_KBN.Value <> "3" Then
			'Update 1999/12/15 ADD E N D TOP
			Call ZALGM_RKSTRD_SUB(ZALGM_UARCN)
			If ZALGM_ERR.Value = "1" Then
				GoTo ZALGERR
			End If
			'Update 1999/12/15 REP START TOP
		Else
			ZALGM_KO_NAIYO.Value = Space(30)
		End If
		'Update 1999/12/15 REP E N D TOP
		
		'<< �������O�t�@�C���Ƀ��O��ǉ����� >>
		
		'�������O�t�@�C���X�V�X�g�A�h�v���V�[�W���|�b�`�k�k
		Call ZALGM_LGSTRD_SUB(ZALGM_UARCN)
		If ZALGM_ERR.Value = "1" Then
			GoTo ZALGERR
		End If
		
		Exit Sub
		
ZALGERR: 
		'���O�o�͎��s���b�Z�[�W
		MsgBox("���O���o�͂���܂���ł���" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		
	End Sub
	
	Sub ZALGM_RKSTRD_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		Dim CCM9030PRO As RDO.rdoQuery
		Dim MSG_NAME As String
		Dim CSZ_FILE_NAME As String
		Dim RETCD1 As Short
		Dim RETCD2 As String
		Dim RETCD3 As String
		Dim RETCD4 As String
		
		'CCM9030
		MKKCMN.ZAEV_FNO = "CCM9030"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ZALGM_ERR.Value = "1"
			MsgBox("�v���V�[�W���̃X�L�[�}��`�G���[�ł�" & Chr(13) & "CCM9030" & ZALGM_ERR_POINT, 48, "")
			Exit Sub
		Else
			CSZ_FILE_NAME = RTrim(MKKCMN.ZAEV_FNM)
		End If
		
		'�v���V�[�W���̒�`
		On Error GoTo STRD_ERR
		
		'UPDATE 1999/12/21
		'    SQL = "begin "
		'    SQL = SQL & RTrim$(CSZ_FILE_NAME) & "CCM9030"
		'    SQL = SQL & "(?,?,?,?,?,?,?); end;"
		SQL = "{CALL "
		SQL = SQL & RTrim(CSZ_FILE_NAME) & "CCM9030( ?,?,?,?,?,?,? )}"
		'UPDATE 1999/12/21
		
		CCM9030PRO = ZALGM_UARCN.CreateQuery("CCM9030PRO", SQL)
		Select Case B_STATUS
			Case 0
			Case Else
				MsgBox("�v���V�[�W���̒�`�G���[�ł�" & Chr(13) & "CCM9030" & ZALGM_ERR_POINT, 48, "")
				ZALGM_ERR.Value = "1"
				Exit Sub
		End Select
		
		'IN
		CCM9030PRO.rdoParameters(0).NAME = "kosin_key" '�X�V�L�[������
		CCM9030PRO.rdoParameters(1).NAME = "sys_kbn" '�V�X�e���敪
		CCM9030PRO.rdoParameters(2).NAME = "prg_id" 'PRGID
		
		'OUT
		CCM9030PRO.rdoParameters(3).NAME = "RETCD1" '��ԃX�e�[�^�X�i�O�F����A�P�F�G���[�j
		CCM9030PRO.rdoParameters(4).NAME = "RETCD2" '�g���[�X�p
		CCM9030PRO.rdoParameters(5).NAME = "RETCD3" '�G���[���e
		CCM9030PRO.rdoParameters(6).NAME = "RETCD4" '�X�V��
		
		CCM9030PRO.rdoParameters(0).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(1).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(2).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(3).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(4).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(5).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(6).Direction = RDO.DirectionConstants.rdParamOutput
		
		CCM9030PRO.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		
		CCM9030PRO.rdoParameters(0).Value = ZALGM_SH_NAIYO.Value '�X�V�L�[������
		CCM9030PRO.rdoParameters(1).Value = ZALGM_SYS_KBN.Value '�V�X�e���敪
		CCM9030PRO.rdoParameters(2).Value = ZALGM_PGID.Value '�o�q�f�h�c
		
		'�v���V�[�W���̎��s
		CCM9030PRO.QueryTimeout = 0
		CCM9030PRO.Execute()
		
		RETCD1 = CCM9030PRO.rdoParameters(3).Value '��ԃX�e�[�^�X�i�O�F����A�P�F�G���[�j
		RETCD2 = CCM9030PRO.rdoParameters(4).Value '�g���[�X�p
		RETCD3 = CCM9030PRO.rdoParameters(5).Value '�G���[���e
		ZALGM_KO_NAIYO.Value = CCM9030PRO.rdoParameters(6).Value '�o�^��
		
		If RETCD1 = -1 Then '�O�ȊO�G���[
			MsgBox("�����t�@�C���̍X�V�ŃG���[���N����܂���" & Chr(13) & RETCD3 & Chr(13) & ZALGM_ERR_POINT, 48, "")
			ZALGM_ERR.Value = "1"
		End If
		
		'�N�G���[�N���[�Y
		CCM9030PRO.Close()
		'UPGRADE_NOTE: �I�u�W�F�N�g CCM9030PRO ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		CCM9030PRO = Nothing
		Exit Sub
		
STRD_ERR: 
		MsgBox("���̑��̃G���[" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		'UPGRADE_NOTE: �I�u�W�F�N�g CCM9030PRO ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		CCM9030PRO = Nothing
		
	End Sub
	
	Sub ZALGM_LGSTRD_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		Dim CCM9020PRO As RDO.rdoQuery
		Dim MSG_NAME As String
		Dim CSZ_FILE_NAME As String
		Dim RETCD1 As Short
		Dim RETCD2 As String
		Dim RETCD3 As String
		
		'CCM9020
		MKKCMN.ZAEV_FNO = "CCM9020"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ZALGM_ERR.Value = "1"
			MsgBox("�v���V�[�W���̃X�L�[�}��`�G���[�ł�" & Chr(13) & "CCM9020" & ZALGM_ERR_POINT, 48, "")
			Exit Sub
		Else
			CSZ_FILE_NAME = RTrim(MKKCMN.ZAEV_FNM)
		End If
		
		'�v���V�[�W���̒�`
		On Error GoTo STRD_ERR
		'UPDATE 1999/12/21
		'    SQL = "begin "
		'    SQL = SQL & RTrim$(CSZ_FILE_NAME) & "CCM9020"
		'    SQL = SQL & "(?,?,?,?,?,?,?,?,?,?,?,?,?,?); end;"
		SQL = "{CALL "
		SQL = SQL & RTrim(CSZ_FILE_NAME) & "CCM9020( ?,?,?,?,?,?,?,?,?,?,?,?,?,? )}"
		'UPDATE 1999/12/21
		
		CCM9020PRO = ZALGM_UARCN.CreateQuery("CCM9020PRO", SQL)
		Select Case B_STATUS
			Case 0
			Case Else
				MsgBox("�v���V�[�W���̒�`�G���[�ł�" & Chr(13) & "CCM9020" & ZALGM_ERR_POINT, 48, "")
				ZALGM_ERR.Value = "1"
				Exit Sub
		End Select
		
		'IN
		CCM9020PRO.rdoParameters(0).NAME = "Inc_code" '��ЃR�[�h
		CCM9020PRO.rdoParameters(1).NAME = "jg_code" '���Ə��R�[�h
		CCM9020PRO.rdoParameters(2).NAME = "sys_kbn" '�V�X�e���敪
		CCM9020PRO.rdoParameters(3).NAME = "s_day" '�������t
		CCM9020PRO.rdoParameters(4).NAME = "s_time" '��������
		CCM9020PRO.rdoParameters(5).NAME = "op_code" '�I�y���[�^�R�[�h
		CCM9020PRO.rdoParameters(6).NAME = "shori_sikibetu" '��������
		CCM9020PRO.rdoParameters(7).NAME = "shori_kbn" '�����敪
		CCM9020PRO.rdoParameters(8).NAME = "shori_naiyo1" '�������e�P
		CCM9020PRO.rdoParameters(9).NAME = "kosin_naiyo2" '�X�V���e�Q
		CCM9020PRO.rdoParameters(10).NAME = "gn_flg" '���z�t���O
		
		'OUT
		CCM9020PRO.rdoParameters(11).NAME = "RETCD1" '��ԃX�e�[�^�X�i�O�F����A�P�F�G���[�j
		CCM9020PRO.rdoParameters(12).NAME = "RETCD2" '�g���[�X�p
		CCM9020PRO.rdoParameters(13).NAME = "RETCD3" '�G���[���e
		
		CCM9020PRO.rdoParameters(0).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(1).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(2).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(3).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(4).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(5).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(6).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(7).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(8).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(9).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(10).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(11).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9020PRO.rdoParameters(12).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9020PRO.rdoParameters(13).Direction = RDO.DirectionConstants.rdParamOutput
		
		CCM9020PRO.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(7).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(8).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(9).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(10).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(11).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		CCM9020PRO.rdoParameters(12).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(13).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		
		CCM9020PRO.rdoParameters(0).Value = ZALGM_INC_CODE.Value '��ЃR�[�h
		CCM9020PRO.rdoParameters(1).Value = ZALGM_JG_CODE.Value '���Ə��R�[�h
		CCM9020PRO.rdoParameters(2).Value = ZALGM_SYS_KBN.Value '�V�X�e���敪
		CCM9020PRO.rdoParameters(3).Value = ZALGM_S_DAY.Value '�������t
		CCM9020PRO.rdoParameters(4).Value = ZALGM_S_TIME.Value '��������
		CCM9020PRO.rdoParameters(5).Value = ZALGM_OP_CODE.Value '�I�y���[�^�R�[�h
		CCM9020PRO.rdoParameters(6).Value = ZALGM_PGID.Value '��������
		CCM9020PRO.rdoParameters(7).Value = ZALGM_SH_KBN.Value '�����敪
		CCM9020PRO.rdoParameters(8).Value = ZALGM_SH_NAIYO.Value '�������e�P
		CCM9020PRO.rdoParameters(9).Value = ZALGM_KO_NAIYO.Value '�X�V���e�Q
		CCM9020PRO.rdoParameters(10).Value = ZALGM_GNFLG.Value '���z�t���O
		
		'�v���V�[�W���̎��s
		CCM9020PRO.QueryTimeout = 0
		CCM9020PRO.Execute()
		
		RETCD1 = CCM9020PRO.rdoParameters(11).Value '��ԃX�e�[�^�X�i�O�F����A�P�F�G���[�j
		RETCD2 = CCM9020PRO.rdoParameters(12).Value '�g���[�X�p
		RETCD3 = CCM9020PRO.rdoParameters(13).Value '�G���[���e
		If RETCD1 = -1 Then '�O�ȊO�G���[
			MsgBox("�������O�̍X�V�ŃG���[���N����܂���" & Chr(13) & RETCD3 & Chr(13) & ZALGM_ERR_POINT, 48, "")
			ZALGM_ERR.Value = "1"
		End If
		
		'�N�G���[�N���[�Y
		CCM9020PRO.Close()
		'UPGRADE_NOTE: �I�u�W�F�N�g CCM9020PRO ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		CCM9020PRO = Nothing
		Exit Sub
		
STRD_ERR: 
		MsgBox("���̑��̃G���[" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		'UPGRADE_NOTE: �I�u�W�F�N�g CCM9020PRO ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		CCM9020PRO = Nothing
		
	End Sub
End Module