Option Strict Off
Option Explicit On
Module ARQCNABAS
	'******************************************************************
	'*    �V�X�e����    �F  �r�l�h�k�d���u�����T����                  *
	'*    �T�u���[�`�����F  �ʃZ�b�V�����R�l�N�g�T�u���[�`��          *
	'*    ��  ��  ��    �F  �r�n�e�s�d�b�|�n��                        *
	'******************************************************************
	
	Public ZACNA_RCN As RDO.rdoConnection '�f�[�^�x�[�X�ڑ����
	
	' **************************************************************
	'   �ڑ�����
	' **************************************************************
	Public Function ZACNA_SUB() As Short
		
		RDOrdoEngine_definst.rdoDefaultCursorDriver = RDO.CursorDriverConstants.rdUseIfNeeded
		
		'�ʃZ�b�V�����ڑ�
		On Error Resume Next
		Err.Clear()
		ZACNA_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACNA_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
		If Err.Number <> n0 Then
			'�ڑ����s
			GoTo ZACNA_EXIT
		End If
		
ZACNA_CONOK: 
		ZACNA_SUB = True
		Exit Function
		
ZACNA_EXIT: 
		Err.Clear()
		If Not (ZACNA_RCN Is Nothing) Then ZACNA_RCN.Close()
		If Not (RdoEnv Is Nothing) Then RdoEnv.Close()
		
		ZACNA_SUB = False
		Exit Function
		
	End Function
	' **************************************************************
	'   �ڑ�����
	' **************************************************************
	Private Function ZACNA_CONSTR(ByRef DSN As String, ByRef USER As String, ByRef PASSW As String) As String
		
		ZACNA_CONSTR = "DSN=" & DSN & ";UID=" & USER & ";PWD=" & PASSW & ";"
		
	End Function
	
	' **************************************************************
	'   �ؒf����
	' **************************************************************
	Public Function ZADISCNA_SUB() As Boolean
		
		On Error GoTo ZADISCNA_ERR
		ZACNA_RCN.Close()
		ZADISCNA_SUB = True
		Exit Function
		
ZADISCNA_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ؒf�G���[")
		
		'�����Ă����G���[���N���A
		RDOrdoEngine_definst.rdoErrors.Clear()
		
		ZADISCNA_SUB = False
		
	End Function
End Module