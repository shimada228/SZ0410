Option Strict Off
Option Explicit On
Module ARQGDBAS
	
	Public ZAGD_PT As Short ' �Ԋu
	Public ZAGD_NO As New VB6.FixedLengthString(30) ' ���b�Z�[�W��������
	'UPGRADE_ISSUE: �錾�̌^���T�|�[�g����Ă��܂���: �Œ蒷������̔z�� �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"' ���N���b�N���Ă��������B
	Public ZAGD_NOT(10) As String*3 ' ZAGD_NO�̕����������e
	
	Sub ZAGD_SUB(ByRef MC As System.Windows.Forms.Form)
		
		Dim GD_DMS As String
		Dim GD_I As Short
		Dim GD_J As Short
		
		If ZAGD_PT = 0 Then '�����Ԋu�ݒ�
			ZAGD_PT = 2
		End If
		
		'   �K�C�h���b�Z�[�W�\���T�u���[�`��
		GD_DMS = ""
		
		If ZAGD_NO.Value <> Space(30) Then
			For GD_I = 1 To 10
				If Val(ZAGD_NOT(GD_I)) = 0 Then
					If Mid(ZAGD_NO.Value, (GD_I - 1) * 3 + 1, 3) = Space(3) Then
						ZAGD_NOT(GD_I) = "000"
					Else
						ZAGD_NOT(GD_I) = Mid(ZAGD_NO.Value, (GD_I - 1) * 3 + 1, 3)
					End If
				End If
			Next GD_I
		End If
		For GD_I = 1 To 10
			If Val(ZAGD_NOT(GD_I)) = 0 Then
				Exit For
			End If
			GD_J = Val(ZAGD_NOT(GD_I))
			If Len(ZAGD_MST(GD_J)) <> 0 Then
				GD_DMS = GD_DMS & ZAGD_MST(GD_J) & Space(ZAGD_PT)
			End If
		Next GD_I
		
		'99/11/30 MKK �ð���ް�����x���ɕύX�����̂ŏC��
		
		'    MC!StBGUIDE.SimpleText = GD_DMS     ' �K�C�h���b�Z�[�W�\��
		'UPGRADE_WARNING: �I�u�W�F�N�g MC!STBGUIDE �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CType(MC.Controls("STBGUIDE"), Object) = GD_DMS ' �K�C�h���b�Z�[�W�\��
		
		ZAGD_PT = 0
		ZAGD_NO.Value = Space(30)
		'UPGRADE_NOTE: Erase �� System.Array.Clear �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		System.Array.Clear(ZAGD_NOT, 0, ZAGD_NOT.Length)
	End Sub
End Module