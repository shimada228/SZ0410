Option Strict Off
Option Explicit On
Module ARQKBBAS
	'------------------------------------------------'
	'           �d����̓T�u���[�`��   For IMNumer   '
	'------------------------------------------------'
	Public ZAKB_SW As Short ' GotFocus ����  ZAKB_SW = 0�Ƃ��A�L���l�����͂��ꂽ��
	' �\�����N���A�� ZAKB_SW = 1 �Ƃ���
	
	Sub ZAKB_SUB(ByRef KeyAscii As Short)
		If ZAKB_SW = 0 Then
			' ���̃R���g���[�����t�H�[�J�X�������Ă��珉�߂ẴL�[���͂�����
			Select Case KeyAscii
				Case 48 To 57 ' �����O�`�X
					ZAKB_SW = 1
					'UPGRADE_ISSUE: Control Value �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0
				Case 45 ' �}�C�i�X����
					'UPGRADE_ISSUE: Control MinValue �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					If System.Windows.Forms.Form.ActiveForm.ActiveControl.MinValue < 0 Then ' ���͍ŏ��l�����̂Ƃ��̂�
						ZAKB_SW = 1 '
						'UPGRADE_ISSUE: Control Value �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						If Val(System.Windows.Forms.Form.ActiveForm.ActiveControl.Value) <> 0 Then '"0"�\���̂Ƃ�
							'UPGRADE_ISSUE: Control Value �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0 ' �\�����[���ɂ���
						End If
					End If
				Case 46 ' �����_
					'UPGRADE_ISSUE: Control FmtDecDigits �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					If System.Windows.Forms.Form.ActiveForm.ActiveControl.FmtDecDigits <> 0 Then ' �����_�L��̂Ƃ�
						ZAKB_SW = 1
						'UPGRADE_ISSUE: Control Value �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0
					End If
				Case System.Windows.Forms.Keys.Back ' �a�r�L�[
					ZAKB_SW = 1
			End Select
		End If
	End Sub
End Module