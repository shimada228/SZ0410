Option Strict Off
Option Explicit On
Module ARQSYSBAS
	
	
	'******************************************************************
	'*      �V�X�e�����t�i�T�[�o�[���t�j�擾      for MKK(�d��)
	'******************************************************************
	'*�y���\�z
	'*      Oracle�T�[�o�[��茻�݂̓��t���捞�݂܂�
	'*�@�@  �܂��A�X�V���̍X�V�����̐ݒ���ȕւɂ��܂�
	'*�y�����z
	'*      1.���t���󂯎�邽�߂̕�����^�̕ϐ�
	'*      2.�g�p�ړI�������t���O: 0�`4 �ȗ��s��
	'*      -----------------------------------------------------------
	'*      ����
	'*      2�̃t���O��0  YYYYMMDD
	'*      2�̃t���O��1  YYYY/MM/DD
	'*      2�̃t���O��2  YYYY-MM-DD HH:MI:SS   (���g�p)
	'*      2�̃t���O��3  YYYYMMDDHHMISS        (�c�a�X�V�p�j
	'*      2�̃t���O��4  YYYY/MM/DD  HH:MI:SS  (���[�o�͗p�j
	'*�y�߂�l�z
	'*      ������:True ���s��:False
	'******************************************************************
	Function ZASYS_SUB(ByRef SYSYMD As String, ByRef FLG As Short) As Boolean
		
		Dim Server As RDO.rdoResultset '���ʃZ�b�g
		
		ZASYS_SUB = False
		
		'UPGRADE_NOTE: IsMissing() �� IsNothing() �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' ���N���b�N���Ă��������B
		If IsNothing(FLG) Then FLG = False
		
		SQL = "select to_char(sysdate,'YYYY-MM-DD HH24:MI:SS') SYMD,to_char(sysdate,'YYYYMMDDHH24MISS') SYMD2,to_char(sysdate,'YYYY/MM/DD  HH24:MI:SS') SYMD3,to_char(sysdate,'YYYYMMDD') AS YMD,to_char(sysdate,'YYYY/MM/DD') AS YMD2 from dual"
		On Error Resume Next
		'    Set Server = ZACN_SZRCN.OpenResultset(SQL)     '99/12/09 DEL KTT YOSHINO
		Server = ZACN_RCN.OpenResultset(SQL) '99/12/09 ADD KTT YOSHINO
		Select Case B_STATUS(Server)
			Case 0
				If FLG = 4 Then
					SYSYMD = Server.rdoColumns("SYMD3").Value
				ElseIf FLG = 3 Then 
					SYSYMD = Server.rdoColumns("SYMD2").Value
				ElseIf FLG = 2 Then 
					SYSYMD = Server.rdoColumns("SYMD").Value
				ElseIf FLG = 1 Then 
					SYSYMD = Server.rdoColumns("YMD2").Value
				Else
					SYSYMD = Server.rdoColumns("YMD").Value
				End If
				ZASYS_SUB = True
			Case Else
				ERRSW = F_ERR
				ZAER_KN = 1
				ZAER_NO.Value = ""
				ZAER_MS.Value = "ARQSYS���G���["
				Call ZAER_SUB()
		End Select
		On Error GoTo 0
		
	End Function
End Module