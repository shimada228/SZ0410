Option Strict Off
Option Explicit On
Module ARQOPBAS
	Dim ARQOPRES As RDO.rdoResultset '���ʃZ�b�g
	Dim ARQOPSV As RDO.rdoResultset '���ʃZ�b�g
	Dim ARQOPNAME As String
	Dim ARQOPDBLINK As String
	
	
	'------------------------------------------------------------
	'�y�֐����z �I�y���[�^���\���T�u���[�`��
	'
	'�y�@  �\�z �I�y���[�^�y�ѻ��ް���t����ʉE���ɕ\��������@�\�ł��B
	'
	'�y�߂�l�z ����
	'
	'------------------------------------------------------------
	'Sub ZAOP_SUB(MC As Form, DB_RCN As rdoConnection, KAISYA As String, OPCODE As String, TABLES As String) '99/12/09 DEL KTT YOSHINO
	Sub ZAOP_SUB(ByRef MC As System.Windows.Forms.Form, ByRef KAISYA As String, ByRef OPCODE As String) '99/12/09 ADD KTT YOSHINO
		
		Dim op_name As New VB6.FixedLengthString(20)
		Dim SYSYMD As String
		Dim INI_NAME As String
		Dim i As Short
		
		'���ϖ�,DB�ݸ���擾
		'    MKKCMN.ZAEV_FNO = Trim(TABLES) & "COM0070"     '99/12/09 DEL KTT YOSHINO
		MKKCMN.ZAEV_FNO = "COM0070" '99/12/09 ADD KTT YOSHINO
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			Exit Sub
		Else
			'        i = InStr(1, MKKCMN.ZAEV_FNM, ".")             '99/12/09 DEL KTT YOSHINO
			'        ARQOPNAME = Mid$(MKKCMN.ZAEV_FNM, 1, i)        '99/12/09 DEL KTT YOSHINO
			'        ARQOPDBLINK = Mid$(MKKCMN.ZAEV_FNM, i + 1)     '99/12/09 DEL KTT YOSHINO
			
			ARQOPNAME = MKKCMN.ZAEV_FNM '99/12/09 ADD KTT YOSHINO
		End If
		
		'    SQL = "Select  NVL(OP_NEME,' ') OP_NEME "      '99/12/20 DEL KTT YOSHINO
		SQL = "Select  NVL(OP_NAME,' ') OP_NEME " '99/12/20 ADD KTT YOSHINO
		
		'    SQL = SQL & "from " & RTrim$(ARQOPNAME) & "COM0070" & RTrim$(ARQOPDBLINK) '99/12/09 DEL KTT YOSHINO
		SQL = SQL & "from " & RTrim(ARQOPNAME) & "COM0070" '99/12/09 ADD KTT YOSHINO
		SQL = SQL & " where INC_CODE = '" & VB6.Format(KAISYA, "00") & "'" '00/01/21 REP IR MEGURO
		SQL = SQL & " AND   OP_CODE  = '" & VB6.Format(OPCODE, "000000") & "'" '00/01/21 REP IR MEGURO
		On Error Resume Next
		'    Set ARQOPRES = DB_RCN.OpenResultset(SQL)       '99/12/09 DEL KTT YOSHINO
		ARQOPRES = ZACN_RCN.OpenResultset(SQL) '99/12/09 ADD KTT YOSHINO
		Select Case B_STATUS(ARQOPRES)
			Case 0
				op_name.Value = ARQOPRES.rdoColumns("op_neme").Value
			Case 24
				op_name.Value = ""
			Case Else
				ERRSW = F_ERR
				'            Set ZAER_RCN = DB_RCN      '99/12/09 DEL KTT YOSHINO
				ZAER_KN = 1
				ZAER_NO.Value = "COM0070"
				ZAER_MS.Value = KAISYA & "-" & OPCODE
				Call ZAER_SUB()
		End Select
		On Error GoTo 0
		
		'�T�[�o�[���t�擾
		SQL = "select sysdate SYMD from dual"
		On Error Resume Next
		'    Set ARQOPSV = DB_RCN.OpenResultset(SQL)        '99/12/09 DEL KTT YOSHINO
		ARQOPSV = ZACN_RCN.OpenResultset(SQL) '99/12/09 ADD KTT YOSHINO
		Select Case B_STATUS(ARQOPSV)
			Case 0
				SYSYMD = VB6.Format(ARQOPSV.rdoColumns("SYMD").Value, "YYYY/MM/DD")
			Case Else
				ERRSW = F_ERR
				'            Set ZAER_RCN = DB_RCN                  '99/12/09 DEL KTT YOSHINO
				ZAER_KN = 1
				ZAER_NO.Value = ""
				ZAER_MS.Value = "ARQOP���G���["
				Call ZAER_SUB()
		End Select
		On Error GoTo 0
		
		'    MC!OPCODE = Format$(OPCODE, "000000") & "�F" & MKKCMN.ZAFIXSTR_SUB(op_name) & Space$(1) & Format$(SYSYMD, "YYYY/MM/DD") '99/12/24 DEL IR MEGURO
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB(op_name) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'UPGRADE_WARNING: �I�u�W�F�N�g MC!OPCODE �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CType(MC.Controls("OPCODE"), Object) = VB6.Format(OPCODE, "000000") & ":" & MKKCMN.ZAFIXSTR_SUB(op_name.Value) & Space(1) & VB6.Format(SYSYMD, "YYYY/MM/DD") '99/12/24 ADD IR MEGURO
		
	End Sub
End Module