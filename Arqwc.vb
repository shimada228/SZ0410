Option Strict Off
Option Explicit On
Module ARQWCBAS
	
	'
	'  +------------------------------------------------------+
	' �b  ���E�B���h�E�ʒu�T�C�Y�ύX�T�u���[�`���� �@ �@      �b
	'  +------------------------------------------------------+
	'�����@MAXIMIZE_SW�@0:�ő剻���Ȃ��@����ݸނ��Ȃ�
	'                   1:�ő剻���Ȃ��@����ݸނ���
	'                   2:�ő�\������
	'
	Sub ZAWC_SUB(ByRef MC As System.Windows.Forms.Form, ByRef MAXIMIZE_SW As Short)
		Dim GETWORK As String
		On Error Resume Next
		If MKKCMN.ZAGI_SUB("����", "ICONFNAME", "", GETWORK, "mkk.ini") = True Then
			'�A�C�R���̃��[�h
			MC.Icon = New System.Drawing.Icon(Trim(GETWORK))
		End If
		On Error GoTo ERRSIZE
		
		Select Case MAXIMIZE_SW
			Case 0
			Case 1
				'�t�H�[���𒆉��ɕ\��
				If MC.WindowState = System.Windows.Forms.FormWindowState.Normal Then
					MC.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MC.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MC.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				End If
			Case 2
				If MKKCMN.ZAGI_SUB("����", "MDIMAX", "False", GETWORK, "mkk.ini") = True Then
					If UCase(Trim(GETWORK)) = "TRUE" Then
						MC.WindowState = System.Windows.Forms.FormWindowState.Maximized
					End If
				End If
				If MC.WindowState <> System.Windows.Forms.FormWindowState.Maximized Then
					MC.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MC.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MC.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				End If
		End Select
		Exit Sub
		
ERRSIZE: 
		If Err.Number = 387 Or Err.Number = 422 Or Err.Number = 383 Then Resume Next
	End Sub
End Module