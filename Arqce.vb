Option Strict Off
Option Explicit On
Module ARQCEBAS
	
	
	'------------------------------------------------------------
	'�y�֐����z �G���[�ϊ��T�u���[�`��
	'
	'�y�@  �\�z Rdo+SQLServer/Rdo+Oracle���甭�������װ�̂����A�]����
	'           Glue+Oracle�ł��ׯ�߂��Ă����װ�������ꍇ�A�]����Glue�Ɠ���
	'           �װ�ð����Ԃ��B����ȊO�̂��̂�99��Ԃ��B
	'
	'�y�߂�l�z Integer�^
	'             0     :�G���[����
	'            -1     :��Ӑ���ᔽ
	'            24     :End Of Fetch
	'           -54     :���b�N��
	'          -955     :���Ɏg�p����Ă���I�u�W�F�N�g�̂��ߍ쐬�ł��Ȃ�
	'          -100     :���ōX�V�ς݁iZACN_DB = SQLSRV �̎��̂݌��o�j
	'            99     :����ȊO�̃G���[
	'
	'�y�֐��d�l�z
	'    Public Function B_STATUS(Optional rKekka As Variant) As Integer
	'        ��������
	'           rKekka�F    OpenResultset,FETCH(MoveNext)��݂̂��̌��ʾ�Ă��w��B
	'                       ����ȊO�̏ꍇ�͏ȗ����邱�ƁB
	'
	'           rKekka���w�肳��Ă�����AEnd Of Fetch���ǂ����̃`�F�b�N��
	'           �ŏ��ɍs���B(End Of Fetch�Ȃ�B_STATUS��24�ŕԂ�)
	'           �ȗ�����Ă�����AEnd Of Fetch�̃`�F�b�N���s��Ȃ��B
	'
	'�y�g�p��z
	'       �P�D���ʾ�Ă��w�肵�Ȃ��ꍇ
	'           AM13INS.Execute
	'           Select Case B_STATUS
	'           ase 0       ����
	'           Case -1     ��Ӑ���ᔽ
	'           Case Else   ����ȊO
	'           End Select
	'
	'       �Q�D���ʾ�Ă��w�肷��ꍇ
	'           Set AM13RS = AM13SEL02.OpenResultset()
	'           Select Case B_STATUS(AM13RS)
	'           Case 0      ����
	'           Case -54    ���b�N��
	'           Case 24    �f�[�^����
	'           Case Else   ����ȊO
	'           End Select
	'
	'----------------------------------------------------------------------------------
	'�y�C�������z
	'
	'   �C�����t�F1998/09/11    �C���ҁFY.Kubo(OSK)
	'   �C�����e�FIntersolv��OracleODBC��ײ�ޑΉ��̂��߁A�uۯ����vStatus�擾���@��ύX�B
	'
	'-----------------------------------------------------------------------------------
	Public Function B_STATUS(Optional ByRef rKekka As Object = Nothing) As Short
		Dim RdoErr As RDO.rdoError
		
		If Err.Number = 0 Then
			'UPGRADE_NOTE: IsMissing() �� IsNothing() �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' ���N���b�N���Ă��������B
			If Not IsNothing(rKekka) Then
				'UPGRADE_WARNING: �I�u�W�F�N�g rKekka.EOF �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				If rKekka.EOF = True Then
					'End Of Fetch
					B_STATUS = 24
					Exit Function
				End If
			End If
			'�G���[����
			B_STATUS = 0
			Exit Function
		End If
		
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			Select Case RdoErr.SQLState
				Case "01S03"
					If ZACN_DB = SQLSRV Then
						'���ōX�V�ς�
						B_STATUS = -100
						Exit Function
					End If
				Case "23000"
					'��Ӑ���ᔽ
					B_STATUS = -1
					Exit Function
				Case "NA000", "S1T00"
					'���b�N��
					B_STATUS = -54
					Exit Function
					'���b�N�����f��ǉ��iIntersolvODBC��ײ��3.0�΍�j
				Case "S1000" '98/09/11�ǉ�
					If ZACN_DB = ORCL And RdoErr.Number = 54 Then '98/09/11�ǉ�
						'���b�N��                                    '98/09/11�ǉ�
						B_STATUS = -54 '98/09/11�ǉ�
						Exit Function '98/09/11�ǉ�
					End If '98/09/11�ǉ�
				Case "S0001"
					'���Ɏg�p����Ă���I�u�W�F�N�g�̂��ߍ쐬�ł��Ȃ�
					B_STATUS = -955
					Exit Function
			End Select
		Next RdoErr
		
		B_STATUS = 99
	End Function
End Module