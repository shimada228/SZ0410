Option Strict Off
Option Explicit On
Module ARQCNAPBAS
	
	'99/12/09 DEL START KTT YOSHINO
	'Public ZACN_UARCN As rdoConnection        '�ڑ����I�u�W�F�N�g ���㔄�|�c�a
	'Public ZACN_SZRCN As rdoConnection        '�ڑ����I�u�W�F�N�g �d���݌ɂc�a
	'Public ZACN_GCRCN As rdoConnection        '�ڑ����I�u�W�F�N�g �Ɩ��ԋ��ʂc�a
	'99/12/09 DEL END   KTT YOSHINO
	
	'------------------------------------------------------------
	'�y�֐����z �R�l�N�g�T�u���[�`�� �`�o�p
	'
	'�y�@  �\�z �����̃R�l�N�g�T�u���[�`���𗘗p��
	'          ���ڰ�Ͻ��A�T�[�o�[�}�X�^�A���ްϽ��ɂĎw�肳�ꂽ�ް��ް��ւ̐ڑ����s��
	'
	'�y�߂�l�z Boolean�^
	'             True  :�ڑ�����
	'            False  :�ڑ����s
	'
	'�y�֐��d�l�z
	'   Public Function ZACNAP_SUB(Company as string,Operator as string) As Boolean
	'     ���v���V�[�W��������
	'
	'     �����ʈ��n�p�����[�^��
	'        ZACNAP_SUB �Ɠ���̓��e�ł�
	'
	'        Public ZACN_UARCN As rdoConnection   '���ʂc�a�ڑ����I�u�W�F�N�g
	'        Public ZACN_UKRCN As rdoConnection   '���㔄�|�c�a�ڑ����I�u�W�F�N�g
	'        Public ZACN_SZRCN As rdoConnection   '�d���݌ɂc�a�ڑ����I�u�W�F�N�g
	'
	'        Public ZACN_DB As Integer            '�g�p�f�[�^�x�[�X    ORCL:ORACLE
	'                                                               SQLSRV:SQLServer
	'        Public ZACN_TIME As Long           'RDO���߂̳��Ď���
	'        Public ZACN_USERID As String       '�ڑ��������[�U��
	'        Public ZACN_PASSWORD As String     '�ڑ������p�X���[�h
	'        Public ZACN_DBNAME As String       '�ڑ������f�[�^�\�[�X��
	'
	'------------------------------------------------------------
	'Public Function ZACNAP_SUB(Company As String, Operator As String) As Boolean
	'Public Function ZACNAP_SUB() As Boolean                                               '99/12/09 DEL KTT YOSHINO
	Public Function ZACNAP_SUB(ByRef ZAID As String, ByRef ZAPW As String, ByRef ZADSN As String) As Boolean '99/12/09 ADD KTT YOSHINO
		
		Dim StrPos As Short
		Dim StrPos2 As Short
		Dim USERINFOKEY As String
		Dim USERINFO As String
		Dim CONSTR As String
		Dim Ret As Integer
		Dim SQL_STR As String 'SQL������i�[
		
		Dim SMARS As RDO.rdoResultset
		Dim SMBRS As RDO.rdoResultset
		Dim SMCRS As RDO.rdoResultset
		
		Dim Post As String
		Dim Server As String
		
		'99/12/09 ADD START KTT YOSHINO �����w��̂c�a�ɐڑ�
		'UPGRADE_NOTE: �I�u�W�F�N�g ZACN_RCN ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		ZACN_RCN = Nothing
		If Not ZACN_SUB(False, ZAID, ZAPW, ZADSN) Then
			ZACNAP_SUB = False
			Exit Function
		End If
		'99/12/09 END START KTT YOSHINO �����w��̂c�a�ɐڑ�
		
		'------------------------------ 99/12/09 DEL START KTT YOSHINO
		'    '�d���݌ɃT�[�o�[�֐ڑ�
		'    Set ZACN_RCN = Nothing
		'    If Not ZACN_SUB(False, WG_SZID, WG_SZPW, WG_SZDSN) Then
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'
		'    Set ZACN_SZRCN = ZACN_RCN       '�ڑ����I�u�W�F�N�g
		'
		'    'DATABASELINK���g�p����ꍇ(WG_DBLINK=1)�͎d���̃R�l�N�V�����𔄏�A���ʂɐݒ肵�ďI��
		'    If WG_DBLINK = "1" Then
		'
		'        Set ZACN_UARCN = ZACN_SZRCN       '�ڑ����I�u�W�F�N�g
		'        Set ZACN_GCRCN = ZACN_SZRCN       '�ڑ����I�u�W�F�N�g
		'
		'        Set ZACN_RCN = Nothing
		'
		'        On Error GoTo 0
		'        ZACNAP_SUB = True
		'        Exit Function
		'    End If
		'
		'    '���|����T�[�o�[�֐ڑ�
		'    If Not ZACN_SUB(False, WG_UAID, WG_UAPW, WG_UADSN) Then
		'
		'        '�d���T�[�o�[�̐ڑ�������
		'        Ret = ZADISCN_SUB(ZACN_SZRCN)
		'
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'    '�����ް��ް��̐ڑ����Z�[�u
		'    Set ZACN_UARCN = ZACN_RCN       '�ڑ����I�u�W�F�N�g
		'
		'
		'    '�Ɩ��ԋ��ʃT�[�o�[�֐ڑ�
		'    Set ZACN_RCN = Nothing
		'    If Not ZACN_SUB(False, WG_GCID, WG_GCPW, WG_GCDSN) Then
		'        ZACNAP_SUB = False
		'
		'        '�d���݌ɃT�[�o�[�̐ڑ�������
		'        Ret = ZADISCN_SUB(ZACN_SZRCN)
		'
		'        '���㔄�|�T�[�o�[�̐ڑ�������
		'        Ret = ZADISCN_SUB(ZACN_UARCN)
		'
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'    Set ZACN_GCRCN = ZACN_RCN       '�ڑ����I�u�W�F�N�g
		'
		'    '��Ɨp�R�l�N�V�����̏�����
		'    Set ZACN_RCN = Nothing
		'------------------------------ 99/12/09 DEL END KTT YOSHINO
		
		On Error GoTo 0
		
		ZACNAP_SUB = True
		Exit Function
		
ZACNAP_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
		
		'�����Ă����G���[���N���A
		RDOrdoEngine_definst.rdoErrors.Clear()
		
		Err.Clear()
		If Not (ZACN_RCN Is Nothing) Then ZACN_RCN.Close()
		If Not (RdoEnv Is Nothing) Then RdoEnv.Close()
		
		On Error GoTo 0
		
		ZACN_USERID = ""
		ZACN_PASSWORD = ""
		ZACN_DBNAME = ""
		
		ZACNAP_SUB = False
		Exit Function
		
		'ZACNAP_USERINFO:
		'    USERINFO = ""
		'    StrPos = InStr(CONSTR, USERINFOKEY)
		'    If StrPos > 0 Then
		'        StrPos = StrPos + 4
		'        StrPos2 = InStr(StrPos, CONSTR, ";")
		'        If StrPos2 > 0 Then
		'            USERINFO = Mid(CONSTR, StrPos, StrPos2 - StrPos)
		'        End If
		'    End If
		'    Return
		
	End Function
	'Private Function ZACNAP_CONSTR(DSN As String, USER As String, PASSW As String) As String
	'    ZACNAP_CONSTR = "DSN=" & DSN & ";UID=" & USER & ";PWD=" & PASSW & ";"
	'End Function
End Module