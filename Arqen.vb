Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ARQENBAS
	'�v���O�����I���T�u���[�`��   �y�`�d�m�c�Q�r�t�a
	'�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
	'ZAEND_SUB  Command��賨��޳����ق��擾���A
	'�@�@�@�@ �@ �ݒ�ɂ���ƭ���è�ނɂ���
	'�@�@�@      ����۸��т�End����
	'�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
	'�ƭ�����Ă΂�Ă���(�������@HWnd:DB�ڑ�������)�A
	'���A�ƭ������݉�����Ă���Ƃ�
	'mkk.ini��MDIMAX�̐ݒ�ɂ��ő傩�ʏ�̻��ނ�
	'�ƭ���è�ޕ\������
	Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
	Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Public Sub ZAEND_SUB()
		Dim Ret As Integer
		Dim MDIMAX As String
		Dim ZAEN_HWND As Integer '���j���[�E�B���h�E�n���h��
		
		If Len(VB.Command()) <> 0 Then
			If InStr(1, VB.Command(), ":") <> 0 Then
				ZAEN_HWND = CInt(Val(Left(VB.Command(), InStr(1, VB.Command(), ":") - 1)))
				If IsWindow(ZAEN_HWND) <> 0 Then
					'�E�B���h�E�����݂���
					If IsIconic(ZAEN_HWND) <> 0 Then
						'���݉�����Ă���
						Ret = MKKCMN.ZAGI_SUB("����", "MDIMAX", "", MDIMAX, "mkk.ini")
						If Ret = True And Trim(UCase(MDIMAX)) = "TRUE" Then
							'SW_SHOWMAXIMIZED) '�ő剻�\��
							Ret = ShowWindow(ZAEN_HWND, 3)
						Else
							'SW_SHOWNORMAL) '���̂܂ܻ��ނŕ\��
							Ret = ShowWindow(ZAEN_HWND, 1)
						End If
					Else
						'���݉�����ĂȂ�
						Ret = SetForegroundWindow(ZAEN_HWND) 'MENU��O�ʂ�
					End If
				End If
			End If
		End If
		End
	End Sub
End Module