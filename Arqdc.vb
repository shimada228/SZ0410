Option Strict Off
Option Explicit On
Module ARQDCBAS
	'**************************************************************
	'*      ���t�`�F�b�N�T�u���[�`���@ �@                         *
	'*          �@�@�@�@�@�@�@�@�@�i�`�u�p�c�b�j                  *
	'*                                                            *
	'*      �G���[�Ώۊ�N�@�P�X�O�O�N                          *
	'**************************************************************
	'�G���[�Ώۊ�N
	Const ZADC_KIJUN As Short = 1900
	
	'���n�ݒ�p�����[�^
	Public ZADC_DATE As New VB6.FixedLengthString(8) '������t
	
	'���ʈ��n�p�����[�^
	Public ZADC_STS As New VB6.FixedLengthString(1) '���ʽð���@0:���� 1:�G���[
	Public ZADC_WEEK As New VB6.FixedLengthString(1) '�j���敪   1:���j 2:���j
	'           3:�Ηj 4:���j
	'           5:�ؗj 6:���j
	'           7:�y�j 0:�װ
	'���[�N
	Dim ZADCL_YMD As Object
	
	Sub ZADC_SUB()
		
		'������Ԃ��G���[�ɐݒ�
		ZADC_STS.Value = "1"
		ZADC_WEEK.Value = "0"
		
		'���l�`�F�b�N
		If IsNumeric(ZADC_DATE.Value) Then
			
			'�W�����̓`�F�b�N
			If Len(Trim(ZADC_DATE.Value)) = 8 Then
				
				'�N�Ó����`�F�b�N
				If CDbl(Mid(ZADC_DATE.Value, 1, 4)) >= ZADC_KIJUN Then
					
					'�����Ó����`�F�b�N
					If Mid(ZADC_DATE.Value, 5, 2) >= "01" And Mid(ZADC_DATE.Value, 5, 2) <= "12" Then
						
						'���t�Ó����`�F�b�N
						'UPGRADE_WARNING: �I�u�W�F�N�g ZADCL_YMD �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						ZADCL_YMD = VB6.Format(Val(ZADC_DATE.Value), "0000/00/00")
						If IsDate(ZADCL_YMD) Then
							
							'�j�������߂�
							'UPGRADE_WARNING: �I�u�W�F�N�g ZADCL_YMD �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							ZADC_WEEK.Value = VB6.Format(ZADCL_YMD, "W")
							ZADC_STS.Value = "0"
						End If
					End If
				End If
			End If
		End If
		ZADC_DATE.Value = ""
	End Sub
End Module