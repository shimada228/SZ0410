Option Strict Off
Option Explicit On
Module SZ0410UBAS
	'
	'
	'   MKK - COMDOC UTILITIES
	'   Written by S. MURAYAMA SAN
	'
	'
	'
	'       ��Ѓ}�X�^
	Private COM0010SELCDU As RDO.rdoQuery
	Private bCOM0010Ready As Boolean
	'       ���Ə��}�X�^
	Private COM0020SELCDU As RDO.rdoQuery
	Private bCOM0020Ready As Boolean
	
	'       ���L����}�X�^
	Private MCM92SELX As RDO.rdoQuery
	Private bMCM92Ready As Boolean
	'       ���L�Ȗڃ}�X�^
	Private MCM93SELX As RDO.rdoQuery
	Private bMCM93Ready As Boolean
	'       ���L����}�X�^
	Private MCM94SELX As RDO.rdoQuery
	Private bMCM94Ready As Boolean
	
	'   ��ЃR�[�h�ɂ���Ж����擾����B
	'   ��Ɂ@CduPrepKaisha()�����s���邱�ƁB
	'   COM0010.bas��Project�ɒǉ����邱��
	Public Function CduDecodeKaisha(ByRef cdKaisha As String, ByRef nmKaisha As String) As Short
		
		If Not bCOM0010Ready Then
			CduDecodeKaisha = F_ERR
			MsgBox("���s�菇�G���[�FCduPrepKaisha()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeKaisha")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeKaisha = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		COM0010SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		
		On Error Resume Next ' (�װ���ׯ��)
		If COM0010RSSW <> "COM0010SELCDU" Or ReQue = False Then
			COM0010RS = COM0010SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			COM0010RSSW = "COM0010SELCDU"
		Else
			COM0010RS.Requery()
		End If
		
		Select Case B_STATUS(COM0010RS) ' (SQL���s�ð���̕]��)
			Case 0
				nmKaisha = COM0010RS.rdoColumns("Inc_name").Value
			Case 24
				CduDecodeKaisha = F_ERR
				nmKaisha = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeKaisha = F_ERR
				nmKaisha = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeKaisha"
				
				''''ZAER_KN = 1
				''''ZAER_NO = "COM0010"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	'
	'   ��ЃR�[�h���f�R�[�h���邽�߂�Query����
	'   COM0010.bas��Project�ɒǉ����邱��
	Public Function CduPrepKaisha() As Object
		
		'   Schema���̎擾  COM0010
		MKKCMN.ZAEV_FNO = "COM0010"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			COM0010_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    COM0010_FILE.NAME = ""
		
		'   ��Ѓ}�X�^��QUERY�쐬
		SQL = "Select Inc_name"
		SQL = SQL & " from "
		SQL = SQL & RTrim(COM0010_FILE.NAME) & "COM0010"
		SQL = SQL & " WHERE Inc_code = ? "
		
		On Error Resume Next
		COM0010SELCDU = ZACN_RCN.CreateQuery("COM0010SELCDU", SQL)
		COM0010SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "COM0010"
			GoTo PREP_COM0010_ERR
		End If
		On Error GoTo 0
		
		COM0010SELCDU.rdoParameters(0).NAME = "Inc_code"
		COM0010SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		COM0010SELCDU.rdoParameters(0).Size = 2
		
		bCOM0010Ready = True
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepKaisha �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepKaisha = F_OFF
		
		Exit Function
		
PREP_COM0010_ERR: 
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepKaisha �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepKaisha = F_ERR
		
		
	End Function
	
	'   ���Ə��R�[�h�ɂ�莖�Ə������擾����B
	'   ��Ɂ@CduPrepJigyo()�����s���邱�ƁB
	'   COM0020.bas��Project�ɒǉ����邱��
	Public Function CduDecodeJigyo(ByRef cdKaisha As String, ByRef cdJigyo As String, ByRef nmJigyo As String) As Short
		
		If Not bCOM0020Ready Then
			CduDecodeJigyo = F_ERR
			MsgBox("���s�菇�G���[�FCduPrepJigyo()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeJigyo")
			Exit Function
		End If
		
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeJigyo = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		COM0020SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		COM0020SELCDU.rdoParameters("jg_code").Value = cdJigyo
		
		On Error Resume Next ' (�װ���ׯ��)
		If COM0020RSSW <> "COM0020SELCDU" Or ReQue = False Then
			COM0020RS = COM0020SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			COM0020RSSW = "COM0020SELCDU"
		Else
			COM0020RS.Requery()
		End If
		
		Select Case B_STATUS(COM0020RS) ' (SQL���s�ð���̕]��)
			Case 0
				nmJigyo = COM0020RS.rdoColumns("jg_name").Value
			Case 24
				CduDecodeJigyo = F_ERR
				nmJigyo = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeJigyo = F_ERR
				nmJigyo = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeJigyo"
				
				''''ZAER_KN = 1
				''''ZAER_NO = "COM0020"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	'
	'   ���Ə��R�[�h���f�R�[�h���邽�߂�Query����
	'   COM0020.bas��Project�ɒǉ����邱��
	Public Function CduPrepJigyo() As Object
		
		'   Schema���̎擾  COM0020
		MKKCMN.ZAEV_FNO = "COM0020"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			COM0020_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    COM0020_FILE.NAME = ""
		
		'   ���Ə��}�X�^��QUERY�쐬
		SQL = "Select jg_name "
		SQL = SQL & " from "
		SQL = SQL & RTrim(COM0020_FILE.NAME) & "COM0020"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		
		On Error Resume Next
		COM0020SELCDU = ZACN_RCN.CreateQuery("COM0020SELCDU", SQL)
		COM0020SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "COM0020"
			GoTo PREP_COM0020_ERR
		End If
		On Error GoTo 0
		
		COM0020SELCDU.rdoParameters(0).NAME = "Inc_code"
		COM0020SELCDU.rdoParameters(1).NAME = "jg_code"
		COM0020SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		COM0020SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		COM0020SELCDU.rdoParameters(0).Size = 2
		COM0020SELCDU.rdoParameters(1).Size = 4
		
		bCOM0020Ready = True
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepJigyo �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepJigyo = F_OFF
		
		Exit Function
		
PREP_COM0020_ERR: 
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepJigyo �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepJigyo = F_ERR
		
	End Function
	
	Public Function DecodeBUSHO(ByRef cdBUSHO As String) As String
		'
		'
		If Not bMCM92Ready Then
			DecodeBUSHO = CStr(F_DUM)
			MsgBox("���s�菇�G���[�FPrepBUSHO()���ɁI", MsgBoxStyle.OKOnly, "DecodeBuSHO")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		DecodeBUSHO = CStr(F_OFF)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		MCM92SELX.rdoParameters("Inc_code").Value = WKB010
		MCM92SELX.rdoParameters("jg_code").Value = WKB020
		MCM92SELX.rdoParameters("bu_code").Value = cdBUSHO
		
		On Error Resume Next ' (�װ���ׯ��)
		If MCM92RSSW <> "MCM92SELX" Or ReQue = False Then
			MCM92RS = MCM92SELX.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			MCM92RSSW = "MCM92SELX"
		Else
			MCM92RS.Requery()
		End If
		
		Select Case B_STATUS(MCM92RS) ' (SQL���s�ð���̕]��)
			Case 0
				DecodeBUSHO = MCM92RS.rdoColumns("CM92004").Value
			Case 24
				DecodeBUSHO = ""
				''''ENDSW = F_END
			Case Else
				DecodeBUSHO = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "MCM92"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'   �����R�[�h���f�R�[�h���邽�߂�Query����
	'   MCM92.bas��Project�ɒǉ����邱��
	Public Sub PrepBUSHO()
		
		'   Schema���̎擾  MCM92
		MKKCMN.ZAEV_FNO = "MCM92"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			MCM92_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    MCM92_FILE.NAME = ""
		
		'   �I�y���[�^�}�X�^��QUERY�쐬
		SQL = "Select CM92004 "
		SQL = SQL & " from "
		SQL = SQL & RTrim(MCM92_FILE.NAME) & "MCM92BUMO"
		SQL = SQL & " WHERE CM92001 = ? "
		SQL = SQL & " AND CM92002 = ? "
		SQL = SQL & " AND CM92003 = ? "
		
		On Error Resume Next
		MCM92SELX = ZACN_RCN.CreateQuery("MCM92SELX", SQL)
		MCM92SELX.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "MCM92"
			
		End If
		On Error GoTo 0
		
		MCM92SELX.rdoParameters(0).NAME = "Inc_code"
		MCM92SELX.rdoParameters(1).NAME = "jg_code"
		MCM92SELX.rdoParameters(2).NAME = "bu_code"
		MCM92SELX.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM92SELX.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM92SELX.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM92SELX.rdoParameters(0).Size = 2
		MCM92SELX.rdoParameters(1).Size = 4
		MCM92SELX.rdoParameters(2).Size = 4
		
		bMCM92Ready = True
		
		
	End Sub
	
	
	
	
	Public Function DecodeKAMOKU(ByRef cdCHU As String, ByRef cdSHO As String) As String
		'
		'
		If Not bMCM94Ready Then
			DecodeKAMOKU = CStr(F_DUM)
			MsgBox("���s�菇�G���[�FPrepKAMOKU()���ɁI", MsgBoxStyle.OKOnly, "DecodeKAMOKU")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		DecodeKAMOKU = CStr(F_OFF)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		MCM94SELX.rdoParameters("Inc_code").Value = WKB010
		MCM94SELX.rdoParameters("jg_code").Value = WKB020
		MCM94SELX.rdoParameters("CHU_code").Value = "0" & cdCHU
		MCM94SELX.rdoParameters("SHO_code").Value = cdSHO
		
		On Error Resume Next ' (�װ���ׯ��)
		If MCM94RSSW <> "MCM94SELX" Or ReQue = False Then
			MCM94RS = MCM94SELX.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			MCM94RSSW = "MCM94SELX"
		Else
			MCM94RS.Requery()
		End If
		
		Select Case B_STATUS(MCM94RS) ' (SQL���s�ð���̕]��)
			Case 0
				DecodeKAMOKU = MCM94RS.rdoColumns("CM94006").Value
			Case 24
				DecodeKAMOKU = ""
				''''ENDSW = F_END
			Case Else
				DecodeKAMOKU = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "MCM94"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'   �Ȗڒ��v�f�A���v�f�R�[�h���f�R�[�h���邽�߂�Query����
	'   MCM94.bas��Project�ɒǉ����邱��
	Public Sub PrepKAMOKU()
		
		'   Schema���̎擾  MCM94
		MKKCMN.ZAEV_FNO = "MCM94"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			MCM94_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    MCM94_FILE.NAME = ""
		
		'   �I�y���[�^�}�X�^��QUERY�쐬
		SQL = "Select CM94006 "
		SQL = SQL & " from "
		SQL = SQL & RTrim(MCM94_FILE.NAME) & "MCM94UCHI"
		SQL = SQL & " WHERE CM94001 = ? "
		SQL = SQL & " AND CM94002 = ? "
		SQL = SQL & " AND CM94003 = ? "
		SQL = SQL & " AND CM94004 = ? "
		
		On Error Resume Next
		MCM94SELX = ZACN_RCN.CreateQuery("MCM94SELX", SQL)
		MCM94SELX.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "MCM94"
			
		End If
		On Error GoTo 0
		
		MCM94SELX.rdoParameters(0).NAME = "Inc_code"
		MCM94SELX.rdoParameters(1).NAME = "jg_code"
		MCM94SELX.rdoParameters(2).NAME = "CHU_code"
		MCM94SELX.rdoParameters(3).NAME = "SHO_code"
		MCM94SELX.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM94SELX.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM94SELX.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM94SELX.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM94SELX.rdoParameters(0).Size = 2
		MCM94SELX.rdoParameters(1).Size = 4
		MCM94SELX.rdoParameters(2).Size = 4
		MCM94SELX.rdoParameters(3).Size = 6
		
		bMCM94Ready = True
		
		
	End Sub
	
	
	Public Function DecodeKAMOCHU(ByRef cdCHU As String) As String
		'
		'
		If Not bMCM93Ready Then
			DecodeKAMOCHU = CStr(F_DUM)
			MsgBox("���s�菇�G���[�FPrepKAMOCHU()���ɁI", MsgBoxStyle.OKOnly, "DecodeKAMOCHU")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		DecodeKAMOCHU = CStr(F_OFF)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		MCM93SELX.rdoParameters("Inc_code").Value = WKB010
		MCM93SELX.rdoParameters("CHU_code").Value = "0" & cdCHU
		
		On Error Resume Next ' (�װ���ׯ��)
		If MCM93RSSW <> "MCM93SELX" Or ReQue = False Then
			MCM93RS = MCM93SELX.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			MCM93RSSW = "MCM93SELX"
		Else
			MCM93RS.Requery()
		End If
		
		Select Case B_STATUS(MCM93RS) ' (SQL���s�ð���̕]��)
			Case 0
				DecodeKAMOCHU = MCM93RS.rdoColumns("CM93004").Value
			Case 24
				DecodeKAMOCHU = ""
				''''ENDSW = F_END
			Case Else
				DecodeKAMOCHU = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "MCM93"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'   �Ȗڒ��v�f�A���v�f�R�[�h���f�R�[�h���邽�߂�Query����
	'   MCM93.bas��Project�ɒǉ����邱��
	Public Sub PrepKAMOCHU()
		
		'   Schema���̎擾  MCM93
		MKKCMN.ZAEV_FNO = "MCM93"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			MCM93_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    MCM93_FILE.NAME = ""
		
		'   ���L�Ȗڃ}�X�^��QUERY�쐬
		'           ���̃e�[�u���ɍ폜�t���O�͂���܂���B
		SQL = "Select CM93004 "
		SQL = SQL & " from "
		SQL = SQL & RTrim(MCM93_FILE.NAME) & "MCM93KAMO"
		SQL = SQL & " WHERE CM93001 = ? "
		SQL = SQL & " AND CM93002 = ? "
		
		On Error Resume Next
		MCM93SELX = ZACN_RCN.CreateQuery("MCM93SELX", SQL)
		MCM93SELX.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "MCM93"
			
		End If
		On Error GoTo 0
		
		MCM93SELX.rdoParameters(0).NAME = "Inc_code"
		MCM93SELX.rdoParameters(1).NAME = "CHU_code"
		MCM93SELX.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM93SELX.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM93SELX.rdoParameters(0).Size = 2
		MCM93SELX.rdoParameters(1).Size = 4
		
		bMCM93Ready = True
		
		
	End Sub
End Module