Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class SZ0410FRM
	Inherits System.Windows.Forms.Form
	'******************************************************************
	'*  �V�X�e����    �F  �l�j�j  �d���݌ɊǗ��V�X�e��                *
	'*  �v���O������  �F  �d���i�ڊ�{������      �@�@              *
	'*  �v���O�����h�c�F  �r�y�O�S�P�O                                *
	'*  ��  ��  ��   �F               �@�@�@�@�@�@                    *
	'******************************************************************
	
	
	Const N999 As Short = 1
	Const N010 As Short = 2 '   ��ЃR�[�h
	Const N020 As Short = 3 '   ���Ə��R�[�h
	Const N030 As Short = 4 '   �i��
	Const N040 As Short = 5 '   �i��
	Const N050 As Short = 6 '   �K�i
	Const N060 As Short = 7 '   �P��
	'D-CUST-20100610 Start
	'Const N070 = 8                  '   Jan�W���R�[�h
	'Const N080 = 9                  '   Jan�Z�k
	'Const N090 = 10                 '   ���̑��o�[�R�[�h
	'D-CUST-20100610 End
	'A-CUST-20100610 Start
	Const N065 As Integer = N060 + 1 '   ��������
	Const N070 As Object = N065 + 1 '   Jan�W���R�[�h
	Const N080 As Object = N070 + 1 '   Jan�Z�k
	Const N090 As Object = N080 + 1 '   ���̑��o�[�R�[�h
	'A-CUST-20100610 End
	'                   �����E�o���Ȗ�
	Const N100_1 As Object = N090 + 1 '   �K�p���P
	Const N110_1 As Object = N100_1 + 1 '   �����P
	Const N120_1 As Object = N110_1 + 1 '   �_�񉿊i�P
	Const N100_2 As Object = N120_1 + 1
	Const N110_2 As Object = N100_2 + 1
	Const N120_2 As Object = N110_2 + 1
	Const N130 As Object = N120_2 + 1 '   ��p�Ȗ�
	Const N140 As Object = N130 + 1 '   ��p�Ȗ�
	'A-CUST20130212 ��
	Const N150 As Object = N140 + 1 '���Y��
	Const N160 As Object = N150 + 1 '�d��
	'Const N170CMB = N160 + 1            '�ܖ������R���{ 'D-20240115
	'Const N170 = N170CMB + 1            '�ܖ�����       'D-20240115
	'A-CUST20130212��
	
	'A-20240115��
	Const N165 As Object = N160 + 1 '����/�ܖ������敪
	Const N170CMB As Object = N165 + 1 '�ܖ������R���{
	Const N170 As Object = N170CMB + 1 '�ܖ�����
	Const N175 As Object = N170 + 1
	'Const N210 = N175 + 1  'D-CUST-20250201
	'A-20240115��
	'                   �e�핪�ސ���
	Const N210 As Object = N140 + 1 '   �Ȗڕ��ށ@'D-CUST20130212
	'D-20240115��
	'Const N210 = N170 + 1               '   �Ȗڕ��ށ@ 'A-CUST20130212
	'D-20240115��
	'D-20250201��
	'Const N211 = N210 + 1               '   �Ȗڕ���
	'Const N220 = N211 + 1               '   �啪��
	'D-20250201��
	Const N220 As Object = N175 + 1 '   �啪��  'A-20250201
	Const N230 As Object = N220 + 1 '   ������
	Const N240 As Object = N230 + 1 '   ������
	'D-20250201��
	'Const N250 = N240 + 1               '   ����
	'Const N260 = N250 + 1               '   ��������
	'D-20250201��
	Const N260 As Object = N240 + 1 '   ��������    'A-20250201
	Const N270 As Object = N260 + 1 '   CHK������i
	Const N280 As Object = N270 + 1
	Const N290 As Object = N280 + 1
	'A-CUST20130212��
	Const N291 As Object = N290 + 1 'JAN���i����
	'A-CUST20130212��
	'Const N300 = N290 + 1     '   Option�Ǘ��敪'D-CUST20130212
	Const N300 As Object = N291 + 1 '   Option�Ǘ��敪 'A-CUST20130212
	Const N310 As Object = N300 + 1 '   Option�����
	Const N320 As Object = N310 + 1 '   Option�I���P��
	Const N330 As Object = N320 + 1 '   Option�݌ɊǗ�
	Const N340 As Object = N330 + 1 '   Option�e�`�w���M
	Const N350_1 As Object = N340 + 1 '   �����P��
	Const N360_1 As Object = N350_1 + 1 '   ���Z��
	Const N350_2 As Object = N360_1 + 1
	Const N360_2 As Object = N350_2 + 1
	Const N350_3 As Object = N360_2 + 1
	Const N360_3 As Object = N350_3 + 1
	Const N350_4 As Object = N360_3 + 1
	Const N360_4 As Object = N350_4 + 1
	Const N350_5 As Object = N360_4 + 1
	Const N360_5 As Object = N350_5 + 1
	
	Const N370 As Object = N360_5 + 1 'A-20250201
	
	'                   ���̑�
	'Const N410 = N360_5 + 1       '   �ƎҌ��� 'D-20250201
	Const N410 As Object = N370 + 1 '   �ƎҌ���    'A-20250201
	Const N420 As Object = N410 + 1
	'Const N420_1 = N410 + 1       '   ��������
	'Const N420_2 = N420_1 + 1
	'Const N420_3 = N420_2 + 1
	'Const N430 = N420_3 + 1         '   ���ꔭ����
	Const N430 As Object = N420 + 1 '   ���ꔭ����
	
	Const N440 As Object = N430 + 1 '   ����ŗ��敪
	Const N450 As Object = N440 + 1
	Const N460 As Object = N450 + 1
	Const N470 As Object = N460 + 1
	Const N480 As Object = N470 + 1
	Const N490 As Object = N480 + 1
	Const N500 As Object = N490 + 1 '   �����x�~
	Const N510 As Object = N500 + 1 '   �����x�~��
	
	
	Const NF12 As Object = N510 + 1
	Const NEND As Object = NF12 + 1
	
	Const GRP1 As Short = 1
	Const GRP2 As Short = 2
	Const GRP3 As Short = 3
	Const GRP4 As Short = 4
	Const GRP5 As Short = 5
	Const GRP6 As Short = 6
	Const GRP7 As Short = 7
	Const GRP8 As Short = 8
	Const GRP9 As Short = 9
	Const GRP10 As Short = 10
	Const GRP11 As Short = 11
	Const GRP12 As Short = 12
	Const GRP13 As Short = 13
	Const GRP14 As Short = 14
	
	'   GRP1
	'           OptionButton�����敪
	'   GRP2
	'           ��ЁA���Ə��R�[�h
	'   GRP3
	'           �i��
	'   GRP4
	'           �i�����炻�̑��o�[�R�[�h
	'   GRP5
	'           �K�p���A�����A�_�񉿊i�̂P
	'   GRP6
	'           �K�p���A�����A�_�񉿊i�̂Q
	'   GRP7
	'           ��p�Ȗ�
	'   GRP8
	'           �Ȗڕ���
	'   GRP9
	'           �啪�ނ��猟������
	'   GRP10
	'           ������i����e�`�w���M�܂�
	'   GRP11
	'           �����P��
	'   GRP12
	'           �ƎҌ���
	'   GRP13
	'           ��������
	'   GRP14
	'           ���ꔭ�����爵���x�~�܂�
	
	
	Const GEND As Short = 15
	Const MAXNO As Object = NEND + 1
	
	'UPGRADE_WARNING: �z�� CTRLTBL �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Dim CTRLTBL(MAXNO) As CTRLTBL_S
	
	'UPGRADE_WARNING: �z�� GRPTBL �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Dim GRPTBL(GEND) As GRPTBL_S
	
	Dim LST_NO As Short
	Dim CUR_NO As Short
	Dim NXT_NO As Short
	
	Dim SS_KEYCODE As Short '   SpreadOcx KeyDownKey �ۑ�
	
	Dim ByMyself As Boolean '   �C�x���g2�d�N���h�~
	
	Dim lst_row As Integer
	Dim bSPRNotReady As Boolean
	
	
	Dim bBackForm As Boolean
	
	
	Private Sub CHK270_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK270.Enter
		
		If CUR_NO = N270 Then Exit Sub
		
		CUR_NO = N270
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK270_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK270.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK280_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK280.Enter
		
		If CUR_NO = N280 Then Exit Sub
		
		CUR_NO = N280
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK280_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK280.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK290_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK290.Enter
		
		If CUR_NO = N290 Then Exit Sub
		
		CUR_NO = N290
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK290_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK290.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK430_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK430.Enter
		
		If CUR_NO = N430 Then Exit Sub
		
		CUR_NO = N430
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK430_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK430.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK450_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK450.Enter
		
		If CUR_NO = N450 Then Exit Sub
		
		CUR_NO = N450
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK450_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK450.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK460_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK460.Enter
		
		If CUR_NO = N460 Then Exit Sub
		
		CUR_NO = N460
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK460_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK460.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CHK470_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK470.Enter
		
		If CUR_NO = N470 Then Exit Sub
		
		CUR_NO = N470
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK470_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK470.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	'UPGRADE_WARNING: �C�x���g CHK500.CheckStateChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' ���N���b�N���Ă��������B
	Private Sub CHK500_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK500.CheckStateChanged
		
		Dim strToday As String
		
		IMTX510.TabStop = (CHK500.CheckState = 1)
		If CHK500.CheckState <> 1 Then
			KB.tori_kyu_date = ""
			IMTX510.Text = ""
		Else
			''''If Trim(IMTX510.Text) = "" Then
			If Trim(IMTX510.Text) = "" And Trim(KB.tori_kyu_date) = "" Then
				strToday = GETTODAY()
				IMTX510.Text = DateSlashed(strToday)
				KB.tori_kyu_date = strToday
			End If
		End If
		
	End Sub
	
	Private Sub CHK500_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CHK500.Enter
		
		Dim strToday As String
		
		IMTX510.TabStop = (CHK500.CheckState = 1)
		If CHK500.CheckState <> 1 Then
			KB.tori_kyu_date = ""
			IMTX510.Text = ""
		Else
			''''If Trim(IMTX510.Text) = "" Then
			If Trim(IMTX510.Text) = "" And Trim(KB.tori_kyu_date) = "" Then
				strToday = GETTODAY()
				IMTX510.Text = DateSlashed(strToday)
				KB.tori_kyu_date = strToday
			End If
		End If
		
		If CUR_NO = N500 Then Exit Sub
		
		CUR_NO = N500
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CHK500_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CHK500.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CMB060_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB060.Enter
		
		If CUR_NO = N060 Then Exit Sub
		
		CUR_NO = N060
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CMB060_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB060.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Left
				KeyCode = 0
				Call SET_NO(2)
			Case System.Windows.Forms.Keys.Right
				KeyCode = 0
				Call SET_NO(1)
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		End Select
		
	End Sub
	'A-20240115��
	'UPGRADE_WARNING: �C�x���g CMB165.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' ���N���b�N���Ă��������B
	Private Sub CMB165_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB165.SelectedIndexChanged
		If CMB165.SelectedIndex = 0 Or CMB165.SelectedIndex = -1 Then
			CMB170.Enabled = False
			CMB170.BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF80)
			CMB170.SelectedIndex = -1
			KB.k57 = ""
			
			IMNU170(1).Enabled = False
			IMNU170(1).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF80)
			IMNU170(1).Value = 0
			KB.k58 = 0
			KB.k99 = 0
			DSP170(0).Text = CStr(0)
		Else
			CMB170.Enabled = True
			CMB170.BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			
			IMNU170(1).Enabled = True
			IMNU170(1).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
		End If
	End Sub
	
	
	Private Sub CMB165_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB165.Enter
		
		If CUR_NO = N165 Then Exit Sub
		
		CUR_NO = N165
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CMB165_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB165.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Left
				KeyCode = 0
				Call SET_NO(2)
			Case System.Windows.Forms.Keys.Right
				KeyCode = 0
				Call SET_NO(1)
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		End Select
	End Sub
	'A-20240115��
	'A-CUST20130212��
	Private Sub CMB170_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB170.Enter
		
		If CUR_NO = N170CMB Then Exit Sub
		
		CUR_NO = N170CMB
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Private Sub CMB170_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB170.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Left
				KeyCode = 0
				Call SET_NO(2)
			Case System.Windows.Forms.Keys.Right
				KeyCode = 0
				Call SET_NO(1)
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		End Select
		
	End Sub
	'A-CUST20130212��
	Private Sub CMB350_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB350.Enter
		Dim Index As Short = CMB350.GetIndex(eventSender)
		
		Select Case Index
			Case 1
				CTRLTBL(N350_1).CTRL.TabStop = True
				If CUR_NO = N350_1 Then Exit Sub
				CUR_NO = N350_1
			Case 2
				CTRLTBL(N350_2).CTRL.TabStop = True
				If CUR_NO = N350_2 Then Exit Sub
				CUR_NO = N350_2
			Case 3
				CTRLTBL(N350_3).CTRL.TabStop = True
				If CUR_NO = N350_3 Then Exit Sub
				CUR_NO = N350_3
			Case 4
				CTRLTBL(N350_4).CTRL.TabStop = True
				If CUR_NO = N350_4 Then Exit Sub
				CUR_NO = N350_4
			Case 5
				CTRLTBL(N350_5).CTRL.TabStop = True
				If CUR_NO = N350_5 Then Exit Sub
				CUR_NO = N350_5
		End Select
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CMB350_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB350.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = CMB350.GetIndex(eventSender)
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Left
				KeyCode = 0
				Call SET_NO(2)
			Case System.Windows.Forms.Keys.Right
				KeyCode = 0
				Call SET_NO(1)
			Case System.Windows.Forms.Keys.Return
				''''''''''''If CMB350(Index).ListIndex < 0 Then
				If Trim(CMB350(Index).Text) = "" Then
					''''LST_NO = N360_5
					Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
				Else
					Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
				End If
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		End Select
		
		
	End Sub
	
	'A-20250201��
	'UPGRADE_WARNING: �C�x���g CMB370.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' ���N���b�N���Ă��������B
	Private Sub CMB370_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB370.SelectedIndexChanged
		
		If clearActCMB370Click = True Then Exit Sub
		
		CUR_NO = N370
		
		Call IPROCHK_N370()
		
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CMB370_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB370.Enter
		
		If CUR_NO = N370 Then Exit Sub
		
		CUR_NO = N370
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	
	
	Private Sub CMB370_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB370.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Left
				KeyCode = 0
				Call SET_NO(2)
			Case System.Windows.Forms.Keys.Right
				KeyCode = 0
				Call SET_NO(1)
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		End Select
		
	End Sub

	'A-20250201��

	'Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent 'D-20250417
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Click 'A-20250417

		Dim Index As Short = CMDOFNC.GetIndex(eventSender)

		Dim iReturn As Short
		Dim Ret As Short 'A-CUST-20070611

		Select Case Index
			Case 0 '   ESCAPE  �I��

				Call ENDR_RTN()

			Case 2 '   �i�ڌ���    ����͔p�~�ł��B
				'   �i�ڌ����̂c�k�k���Ăяo���B

			Case 3 '   �⍇��
				Call F3QUERY(CUR_NO)
				'        DoEvents
				'        Debug.Print "After F3QUERY CUR_NO="; CUR_NO

			Case 4 '   ����
				Call F4COPY()
				NXT_NO = CUR_NO
				Call FOCUS_SET()

			Case 5 '   �N���A
				'        WKB030 = ""
				'Call SCRCLR_RTN                        'D-CUST-20100610
				Call SCRCLR_RTN(False) 'A-CUST-20100610
				Call DBRollbackTrans()
				Call DBBeginTrans()
				'D-20250303-S
				'NXT_NO = N030
				'D-20250303-E
				'A-20250303-S
				If KBKBN = F_ADD Then
					NXT_NO = N040
				Else
					NXT_NO = N030
				End If
				'A-20250303-E
				Call FOCUS_SET()
				''''    CTRLTBL(N030).CTRL.SetFocus

				'A-CUST-20100610 Start
			Case 6 '�i�ڎ捞
				bBackForm = True
				SZ0411FRM.ShowDialog()
				NXT_NO = LST_NO
				Call FOCUS_SET()

			Case 7 '�i�ڑI��
				bBackForm = True
				SZ0412FRM.ShowDialog()
				NXT_NO = LST_NO
				Call FOCUS_SET()
				'A-CUST-20100610 End

			Case 8 '   �폜
				Call F8DELETE()
				'        NXT_NO = CUR_NO
				'        Call FOCUS_SET

			Case 9 '   �ǉ�        ����͔p�~�ł��B

			Case 12 '   ���s
				If W_KENGEN(1) < 2 Then
					ZAER_KN = n0
					ZAER_CD = 303
					ZAER_MS.Value = ""
					ZAER_NO.Value = "" 'A-CUST-20100610
					Call ZAER_SUB()
					NXT_NO = LST_NO
					Call FOCUS_SET()
					Exit Sub
				End If

				'��A-CUST-20070611
				'�Z�L�����e�B�`�F�b�N�i�R�j���Ə��X�V����
				MKKDBCMN.MKKDBCMN_RCN = ZACN_RCN
				Ret = MKKDBCMN.MKKDBCMN_SQTGET3_SUB(4, "SZ0410", IMTX010.Text, IMTX020.Text, "", WG_OPCODE, W_KENGEN(3))
				If Ret <> n0 Then
					ERRSW = F_ERR
					Exit Sub
				ElseIf W_KENGEN(3) = 0 Then
					ERRSW = F_ERR
					ZAER_KN = n0
					ZAER_CD = 303
					ZAER_NO.Value = ""
					ZAER_MS.Value = ""
					Call ZAER_SUB()
					NXT_NO = LST_NO
					Call FOCUS_SET()
					Exit Sub
				End If
				'��A-CUST-20070611

				iReturn = ALLCHK_RTN()
				If iReturn = 0 Then
					Call GO_EXEC()
					''If ENDSW = F_END Then

					Call SCRCLR_RTN()
					''''        CTRLTBL(N030).CTRL.SetFocus
					'A-CUST-20100610 Start
					If KBKBN = F_ADD Then
						NXT_NO = N040
						CUR_NO = NEND 'A-20250302-
						LST_NO = CUR_NO 'A-20250302-
					Else
						'A-CUST-20100610 End
						NXT_NO = N030
					End If 'A-CUST-20100610
					Call FOCUS_SET()

				End If
		End Select

	End Sub


	Private Sub CMDOFNC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Enter
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		If Index <> 12 Then Exit Sub
		
		If CUR_NO = NF12 Then Exit Sub
		
		If Index = 12 Then
			CUR_NO = NF12
			'    Else
			'        NXT_NO = LST_NO
			'        Call FOCUS_SET
			'        Exit Sub
			
		End If
		
		CUR_NO = NF12
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		ZAKB_SW = 0
		
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
		
	End Sub

	'Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent 'D-20250417
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMDOFNC.KeyDown 'A-20250417
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub SZ0410FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'   Shift,Ctrl,Graph(Alt)�L�[�������A��������
		If Shift <> 0 Then
			Exit Sub
		End If
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape '   �I��
				If CMDOFNC(0).Enabled Then
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
				End If
				
				
			Case System.Windows.Forms.Keys.Return
				Call SET_NO(1) '   ������
			Case System.Windows.Forms.Keys.Up
				Call SET_NO(2) '   �O����
			Case System.Windows.Forms.Keys.Down
				Call SET_NO(3) '   ���O���[�v
				KeyCode = 0
				
				
			Case System.Windows.Forms.Keys.F2
				CMDOFNC_ClickEvent(CMDOFNC.Item(2), New System.EventArgs())
				
			Case System.Windows.Forms.Keys.F3
				If CMDOFNC(3).Text <> "" Then
					CMDOFNC(3).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(3), New System.EventArgs())
				End If
				KeyCode = n0
				
			Case System.Windows.Forms.Keys.F4
				If CMDOFNC(4).Enabled Then
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(4), New System.EventArgs())
				End If
				
			Case System.Windows.Forms.Keys.F5
				If CMDOFNC(5).Enabled Then
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				End If
				
				'A-CUST-20100610 Start
			Case System.Windows.Forms.Keys.F6 'F6
				If CMDOFNC(6).Text <> "" Then
					CMDOFNC(6).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
				KeyCode = n0
				
			Case System.Windows.Forms.Keys.F7 'F7
				If CMDOFNC(7).Text <> "" Then
					CMDOFNC(7).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
                KeyCode = n0
				KeyCode = n0
				'A-CUST-20100610 End

			Case System.Windows.Forms.Keys.F8
				If CMDOFNC(8).Enabled Then
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(8), New System.EventArgs())
				End If
				
			Case System.Windows.Forms.Keys.F9
				
			Case System.Windows.Forms.Keys.F12
				If CMDOFNC(12).Enabled Then
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				End If
				
		End Select
		
	End Sub
	
	Private Sub SZ0410FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'-- �E�B���h�E�ʒu�T�C�Y�ύX�@�T�u���[�`��
		'    Call ZAWC_SUB(Me, 1)
		
		Call TBL_SET()
		
		Call INIT_SPR()
		
		
		KBKBN = F_ADD
		
		Call INIT_RTN()
		If ENDSW = F_END Or ERRSW = F_ERR Then
			Call ENDR_RTN()
			Exit Sub
		End If
		
		If W_KENGEN(1) < 2 Then
			OPTO999(1).Enabled = False
			OPTO999(3).Enabled = False
			KBKBN = 2
		End If
		
		'A-20240115��
		CMB165.Items.Clear() '�R���{�{�b�N�X �N���A
		CMB165.Items.Add(New VB6.ListBoxItem("�����Ȃ�", 0)) '�o�^
		CMB165.Items.Add(New VB6.ListBoxItem("�������", 1)) '�o�^
		CMB165.Items.Add(New VB6.ListBoxItem("�ܖ�����", 2)) '�o�^
		'A-20240115��
		
		'A-20250201��
		CMB370.Items.Clear() '�R���{�{�b�N�X �N���A
		CMB370.Items.Add(New VB6.ListBoxItem("", 0)) '�o�^
		CMB370.Items.Add(New VB6.ListBoxItem("�W��", 1)) '�o�^
		CMB370.Items.Add(New VB6.ListBoxItem("�y��", 2)) '�o�^
		
		'�\���ʒu����
		LBL260.Top = LBL250.Top
		IMTX260.Top = IMTX250.Top
		DSP260.Top = IMTX260.Top
		CHK500.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(LBL490.Top) + 60)
		IMTX510.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(IMTX490.Top) + 15)
		LBL490.Top = LBL480.Top
		IMTX490.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(IMTX480.Top) + 15)
		LBL480.Top = CHK470.Top
		IMTX480.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(CHK470.Top) - 45)
		CHK470.Top = CHK460.Top
		CHK460.Top = CHK450.Top
		CHK450.Top = LBL440.Top
		'A-20250201��
		
		'   ��ʃN���A�������\��
		Call SCRCLR_RTN() '   F5-CLEAR�Ɠ�������
		
		'   ���ʏ�������
		
		'   ��ЁA���Ə��A�I�y���[�^���擾����B
		Call ZAOP_SUB(Me, WG_INCCODE, WG_OPCODE)
		
		'   �N���`�F�b�N�������Ȃ��B
		
		'   ��ʏ����\��
		LST_NO = n0
		
		Me.Show()
		COMBO_INIT(CMB060)
		COMBO_COPY(CMB060, CMB350(1))
		COMBO_COPY(CMB060, CMB350(2))
		COMBO_COPY(CMB060, CMB350(3))
		COMBO_COPY(CMB060, CMB350(4))
		COMBO_COPY(CMB060, CMB350(5))
		
		'A-CUST20130212��
		CMB170.Items.Clear() '�R���{�{�b�N�X �N���A
		CMB170.Items.Add("")
		CMB170.Items.Add(New VB6.ListBoxItem("��", 1)) '�o�^
		CMB170.Items.Add(New VB6.ListBoxItem("��", 2)) '�o�^
		CMB170.Items.Add(New VB6.ListBoxItem("�N", 3)) '�o�^
		'A-CUST20130212��
		
		LST_NO = N999
		CUR_NO = N999
		NXT_NO = N999 '   1/5
		Call FUNCSET_RTN()
		Call FOCUS_SET()
		
		
		TAB010.SelectedIndex = 0
		
	End Sub
	
	Private Sub INIT_SPR()
		
		'''''SPR420.Top = SHA420.Top
		'    SPR420.Top = LIN420(1).Y1 + 30
		'    SPR420.Left = SHA420.Left + 30
		'''''SPR420.Height = SHA420.Height - 360
		'    SPR420.Height = SHA420.Height - (LIN420(1).Y1 - SHA420.Top) - 120
		'    SPR420.Width = SHA420.Width - 30
		
		'   2000/01/23              change  KOKOKARA
		SPR420.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SHA420.Top) + 30 - 120)
		''''SPR420.Top = LIN420(1).Y1 + 30
		SPR420.Left = SHA420.Left
		SPR420.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SHA420.Height) - 60 - 30)
		''''SPR420.Height = SHA420.Height - (LIN420(1).Y1 - SHA420.Top) - 120
		SPR420.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SHA420.Width) + 60) '- 30          '2000/02/15
		SHA420.Visible = False
		LIN420(1).Visible = False
		LBL420.Visible = False
		
		
		
	End Sub
	
	Public Sub SpreadAppend()
		
		Dim saveRow As Integer
		
		saveRow = SPR420.ActiveRow
		
		'   2000/01/23  Add     KOKOKARA
		If SPR420.MaxRows <= SPR420.DataRowCnt Then
			SPR420.MaxRows = SPR420.DataRowCnt + 1
			SPR420.ROW = SPR420.DataRowCnt + 1
			''''SPR420.CellType = SS_CELL_TYPE_FLOAT
			SPR420.set_RowHeight(SPR420.ROW, SPR_HEIGHT)
			Call SpreadProperty((SPR420.ROW))
		End If
		'   2000/01/23  Add     KOKOMADE
		
		'    SPR420.ROW = saveRow
		System.Diagnostics.Debug.Assert(SPR420.ActiveRow = saveRow, "")
		
		
		ROW = SPR420.ActiveRow
		SPR420.ROW = ROW
		SPR420.Col = 1
		SPR420.Text = SEL_FIND
		
		'   2000/01/23  Add     KOKOKARA
		If SPR420.MaxRows <= SPR420.DataRowCnt Then
			SPR420.MaxRows = SPR420.DataRowCnt + 1
			SPR420.ROW = SPR420.DataRowCnt + 1
			''''SPR420.CellType = SS_CELL_TYPE_FLOAT
			SPR420.set_RowHeight(SPR420.ROW, SPR_HEIGHT)
			Call SpreadProperty((SPR420.ROW))
		End If
		'   2000/01/23  Add     KOKOMADE
		
		
		
		SPR420.ROW = ROW + 1
		SPR420.Col = 1
		SPR420.Row2 = ROW + 1
		SPR420.Col2 = 1
		
		SPR420.Action = SS_ACTION_SELECT_BLOCK
		SPR420.Action = SS_ACTION_ACTIVE_CELL
		
		Dim bCancel As Boolean

		'Call SPR420_LeaveCell(SPR420, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(1, ROW, 1, ROW + 1, bCancel)) 'D-20250417
		Call SPR420_LeaveCell(SPR420, New AxFPSpread._DSpreadEvents_LeaveCellEvent(1, ROW, 1, ROW + 1, bCancel)) 'A-20250417

	End Sub
	
	Public Sub SpreadProperty(ByRef IROW As Integer)
		
		'    If True Then Exit Sub
		
		With SPR420
			.ROW = IROW
			.Col = 1
			'        .CellType = SS_CELL_TYPE_EDIT
			'        .TypeEditCharSet = SS_CELL_EDIT_CHAR_SET_NUMERIC    '�����̂�
			'        .TypeEditLen = 5
			'        .TypeTextShadow = False
			'        .TypeTextShadowIn = False
			'        .BackColor = SPRD_BACKCOL_INP
			'        .Value = ""
			.Lock = False
			
			.Col = 2
			'        .TypeTextShadow = False
			'        .TypeTextShadowIn = False
			'        .BackColor = SPRD_BACKCOL_DSP
			'        .Value = ""
			.Lock = True
			
			
			'        .ROW = 0
			'        .Col = 1
			'        .Lock = True
			
		End With
		
	End Sub
	
	Public Sub SpreadInit()
		
		Dim i As Short
		Dim cnt As Integer
		'   2000/01/23  DEL
		'    For i = 1 To MAXSPREAD
		'        SPR420.ROW = i
		'        SPR420.Col = 1
		'        SPR420.Text = ""
		'        SPR420.Col = 2
		'        SPR420.Text = ""
		'        SPR420.Col = 3
		'        SPR420.Text = ""
		'    Next i
		
		lst_row = 0
		
		cnt = SPR420.DataRowCnt
		Do While cnt > 0
			SPR420.ROW = SPR420.ActiveRow
			SPR420.Action = SS_ACTION_DELETE_ROW
			System.Windows.Forms.Application.DoEvents()
			cnt = cnt - 1
		Loop 
		
		SPR420.MaxRows = 0
		
		SPR420.MaxRows = 1
		'    SPR420.ROW = 0
		'    SPR420.RowHidden = True
		SPR420.ROW = 1
		SPR420.Col = 1
		SPR420.Action = SS_ACTION_ACTIVE_CELL
		'''SPR420.CellType = SS_CELL_TYPE_FLOAT
		SPR420.set_RowHeight(1, SPR_HEIGHT)
		Call SpreadProperty(1)
		
		
	End Sub
	
	
	Private Sub SpreadDelete()
		
		'   2000/01/24          Add                     KOKOKARA
		SPR420.ROW = SPR420.ActiveRow
		SPR420.Col = 1
		If Trim(SPR420.Text) = "" And SPR420.ROW > SPR420.DataRowCnt Then
			Exit Sub
		End If
		'   2000/01/24          Add                     KOKOMADE
		
		
		SPR420.ROW = SPR420.ActiveRow
		SPR420.Action = SS_ACTION_DELETE_ROW
		'   2000/01/23              Add             KOKOKARA
		If SPR420.MaxRows > 1 Then
			SPR420.MaxRows = SPR420.MaxRows - 1
		End If
		'   2000/01/23              Add             KOKOMADE
		
		NXT_NO = N420
		Call FOCUS_SET()
		
	End Sub
	
	
	
	Public Function IPROCHK() As Boolean
		
		Dim i As Short
		
		IPROCHK = True
		ERRSW = F_OFF
		ENDSW = F_OFF
		
		'    Debug.Assert LST_NO <> N450 And LST_NO <> N460 And LST_NO <> N470
		
		If CUR_NO = LST_NO Then Exit Function
		
		Select Case LST_NO
			Case N999
				Call IPROCHK_N999()
			Case N010
				Call IPROCHK_N010()
			Case N020
				Call IPROCHK_N020()
			Case N030
				Call IPROCHK_N030()
			Case N040
				Call IPROCHK_N040()
			Case N050
				Call IPROCHK_N050()
			Case N060
				Call IPROCHK_N060()
			Case N065 'A-CUST-20100610
				Call IPROCHK_N065() 'A-CUST-20100610
			Case N070
				Call IPROCHK_N070()
			Case N080
				Call IPROCHK_N080()
			Case N090
				Call IPROCHK_N090()
			Case N100_1, N100_2
				Call IPROCHK_N100(LST_NO)
			Case N110_1, N110_2
				Call IPROCHK_N110(LST_NO)
			Case N120_1, N120_2
				Call IPROCHK_N120(LST_NO)
			Case N130, N140
				Call IPROCHK_N130N140(LST_NO)
				'A-CUST20130212��
			Case N150
				Call IPROCHK_N150()
			Case N160
				Call IPROCHK_N160()
				'A-20240115��
			Case N165
				Call IPROCHK_N165()
				'A-20240115��
			Case N170CMB
				Call IPROCHK_N170CMB()
			Case N170
				Call IPROCHK_N170()
				'A-CUAT20130212��
				'D-CUST-20250201��
				'Case N210, N211
				'Call IPROCHK_N210N211(LST_NO)
				'D-CUST-20250201��
			Case N220, N230, N240
				Call IPROCHK_N220N230N240(LST_NO)
				'D-CUST-20250201��
				'Case N250
				'Call IPROCHK_N250
				'D-CUST-20250201��
			Case N260
				Call IPROCHK_N260()
			Case N270, N280, N290
				Call IPROCHK_CHKBTN(LST_NO)
				'A-CUST20130212��
			Case N291
				Call IPROCHK_N291()
				'A-CUST20130212��
			Case N300 To N340
				Call IPROCHK_OPTO(LST_NO)
			Case N350_1 To N360_5
				Call IPROCHK_N350N360(LST_NO)
				'A-20250201��
			Case N370
				Call IPROCHK_N370()
				'A-20250201��
			Case N410
				Call IPROCHK_N410()
				
			Case N420 '   SPREAD.OCX
				Call IPROCHK_NOCHK(LST_NO)
				
			Case N430, N450 To N470, N500
				Call IPROCHK_CHKBTN(LST_NO)
				'D-20250201��
				'Case N440
				'Call IPROCHK_N440
				'D-20250201��
			Case N480, N490
				Call IPROCHK_N480N490(LST_NO)
			Case N510
				Call IPROCHK_N510()
				
		End Select
		'########## �װ�̔��� ##########
		If ERRSW = F_ERR Then
			IPROCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
		
		'   �G���[�̏ꍇ�A�A�A
		'    If ENDSW = F_END Then
		'        If CUR_NO <= LST_NO Then
		'            ERRSW = F_OFF
		'-------�t����----���O���ڂ̍ĕ\��
		'            Select Case LST_NO
		'                Case N010
		'                    IMTX010.Text = ""
		'                    DSP010.Caption = ""
		'                Case N020
		'                    IMTX020.Text = ""
		'                    DSP020.Caption = ""
		'            End Select
		'        Else
		'            IPROCHK = False
		'            NXT_NO = GetNxtNo(LST_NO, 0)
		'            Call FOCUS_SET
		'        End If
		'    End If
		
	End Function
	
	'   NO CHECK ���ڂ͊m��̂�
	Private Sub IPROCHK_NOCHK(ByRef IDX As Short)
		
	End Sub
	Private Sub IPROCHK_N999()
		
		'A-CUST-20100610 Start
		If KBKBN = F_ADD Then
			If CUR_NO > N040 Then
				ERRSW = F_ERR
				ENDSW = F_END
				Exit Sub
			End If
		Else
			'A-CUST-20100610 End
			If CUR_NO > N030 Then
				ERRSW = F_ERR
				ENDSW = F_END
				Exit Sub
			End If
		End If 'A-CUST-20100610
		
		'    If Trim(WKB010) = "" Or Trim(WKB020) = "" Then
		'        ERRSW = F_ERR
		'        ENDSW = F_END
		'        Exit Sub
		'    End If
		
		'   �폜�̂Ƃ��͉�ЁA���Ə��A�i�ԈȊO��Disable
		Select Case KBKBN
			Case 1 '   �ǉ�
				Call SetMode("A")
				IMTX030.TabStop = False 'A-CUST-20100610
			Case 2 '   �C��
				Call SetMode("C")
			Case 3 '   �폜
				Call SetMode("D")
				
		End Select
		
	End Sub
	
	Private Sub SetMode(ByRef strOpt As String)
		
		Dim i As Short
		
		If strOpt = "D" Then
			For i = N030 + 1 To N510
				CTRLTBL(i).CTRL.TabStop = False
			Next i
			'CTRLTBL(N040).CTRL.TabStop = True
			'        For i = N030 + 1 To N510
			'        Debug.Assert CTRLTBL(i).CTRL.TabStop
			'        Next i
			imtxDummy.TabStop = False
			
			Debug.Print("SetMode(D) " & CMDOFNC(12).TabStop)
			
		Else
			For i = N030 + 1 To N510
				CTRLTBL(i).CTRL.TabStop = True
			Next i
			imtxDummy.TabStop = True
			
		End If
		'Debug.Assert CTRLTBL(N030).CTRL.TabStop                D-CUST-20100610
		
	End Sub
	
	Private Sub IPROCHK_N010()
		
		Dim iReturn As Short
		Dim strN010 As String
		Dim strN010DSP As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX010.Text = WKB010
		'        DSP010.Caption = WKB010DSP
		'        Exit Sub
		'    End If
		
		If RTrim(IMTX010.Text) = "" Then
			DSP010.Text = ""
			ERRSW = F_ERR
			Exit Sub
		End If
		
		IMTX010.Text = VB6.Format(IMTX010.Text, "00")
		
		'   ��ЃR�[�h���݃`�F�b�N
		strN010 = IMTX010.Text '   Fix Length?
		iReturn = CduDecodeKaisha(strN010, strN010DSP)
		
		If iReturn = F_ERR Then
			''''DSP010.Caption = ""
			If CUR_NO < LST_NO Then
				IMTX010.Text = WKB010
				DSP010.Text = WKB010DSP
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   ��к��ނ��ύX���ꂽ�玖�Ə����N���A
		If strN010 <> WKB010 Then
			WKB020 = ""
			WKB030 = ""
			''''Exit Sub
			'   �m��
			WKB010 = strN010
			KB.Inc_code = strN010
			
			'   ��Ж��̕\��
			WKB010DSP = strN010DSP
			DSP010.Text = WKB010DSP
			WKB020DSP = ""
			DSP020.Text = ""
			Call DBRollbackTrans()
			Call DBBeginTrans()
			Call SCR_ADDNEW()
			Call SpreadInit()
			Call SCR_DSPDATA()
			If KBKBN = 3 Then Call SetMode("D")
			
			COMBO_INIT(CMB060)
			COMBO_COPY(CMB060, CMB350(1))
			COMBO_COPY(CMB060, CMB350(2))
			COMBO_COPY(CMB060, CMB350(3))
			COMBO_COPY(CMB060, CMB350(4))
			COMBO_COPY(CMB060, CMB350(5))
			
		End If
		
	End Sub
	
	
	Private Sub IPROCHK_N020()
		
		Dim iReturn As Short
		Dim strN020 As String
		Dim strN020DSP As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX020.Text = WKB020
		'        DSP020.Caption = WKB020DSP
		'        Exit Sub
		'    End If
		
		If RTrim(IMTX020.Text) = "" Then
			If CUR_NO < LST_NO Then
				IMTX020.Text = WKB020
				DSP020.Text = WKB020DSP
				Exit Sub
			End If
			DSP020.Text = ""
			ERRSW = F_ERR
			Exit Sub
		End If
		
		IMTX020.Text = VB6.Format(IMTX020.Text, "0000")
		
		'   ���Ə��R�[�h���݃`�F�b�N
		strN020 = IMTX020.Text '   Fix Length?
		iReturn = CduDecodeJigyo(WKB010, strN020, strN020DSP)
		
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				IMTX020.Text = WKB020
				DSP020.Text = WKB020DSP
				Exit Sub
			End If
			
			DSP020.Text = ""
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   ���Ə����ނ��ς�����i�Ԃ��N���A
		If strN020 = WKB020 Then
			Exit Sub
		Else
			WKB030 = ""
			'   �m��
			WKB020 = strN020
			KB.jg_code = strN020
			'   ���Ə����̕\��
			WKB020DSP = strN020DSP
			DSP020.Text = WKB020DSP
			Call DBRollbackTrans()
			Call DBBeginTrans()
			Call SCR_ADDNEW()
			Call SpreadInit()
			Call SCR_DSPDATA()
			If KBKBN = 3 Then Call SetMode("D")
			
			COMBO_INIT(CMB060)
			COMBO_COPY(CMB060, CMB350(1))
			COMBO_COPY(CMB060, CMB350(2))
			COMBO_COPY(CMB060, CMB350(3))
			COMBO_COPY(CMB060, CMB350(4))
			COMBO_COPY(CMB060, CMB350(5))
			
		End If
		
	End Sub
	
	Private Sub IPROCHK_N030()
		
		Dim iReturn As Short
		Dim strN030 As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX030.Text = WKB030
		'        Exit Sub
		'    End If
		If RTrim(IMTX030.Text) = "" Then
			If CUR_NO < LST_NO Then
				IMTX030.Text = WKB030
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		IMTX030.Text = VB6.Format(IMTX030.Text, "00000")
		
		'   �i�ԑ��݃`�F�b�N
		strN030 = IMTX030.Text '   Fix Length?
		
		'   �i�ڃR�[�h���ς��Ă��Ȃ���΍ĕ\���̕K�v�Ȃ��B2000/1/9 MB-3123
		If strN030 = WKB030 Then
			Exit Sub
		End If
		
		'   ��ЁA���Ə��s���̂Ƃ��̓G���[�ł��B
		
		iReturn = FILGET_SZM0010(WKB010, WKB020, strN030, KB)
		
		If iReturn = F_END Then '   �i�Ԍ�����Ȃ��Ƃ��A
			If KBKBN <> F_ADD Then '   �C���E�폜�̂Ƃ��̓G���[
				If CUR_NO < LST_NO Then '   ��ɖ߂�Ƃ��͌��̒l�ɖ߂��ďI���
					IMTX030.Text = WKB030
					Exit Sub
				End If
				'            IMTX030.Text = KB.hin_code
				ZAER_CD = 1206
				ZAER_NO.Value = "" 'A-CUST-20100610
				Call ZAER_SUB()
				ERRSW = F_ERR '   �J�[�\�����i�ނƂ��̓G���[
				Exit Sub
			Else '   �ǉ��̂Ƃ��͂n�j
				'   �V�K�o�^�̂Ƃ���OK�Ƃ���B
				Call SCR_ADDNEW()
				ENDSW = F_OFF
				ERRSW = F_OFF
			End If
		Else '   �i�Ԃ����������Ƃ��A
			If KBKBN = F_ADD Then '   �ǉ��̂Ƃ��̓G���[
				'D-CUST-20100610 Start
				'If CUR_NO < LST_NO Then         '   ��ɖ߂�Ƃ��͌��̒l�ɖ߂��ďI���
				'    IMTX030.Text = WKB030
				'    Exit Sub
				'End If
				'D-CUST-20100610 End
				'            IMTX030.Text = WKB030
				'            ENDSW = F_END
				'            NXT_NO = N030
				'D-CUST-20100610 Start
				'            ZAER_CD = 1205
				'            Call ZAER_SUB
				'            strN030 = WKB030
				'            Call DBRollbackTrans
				'            Call DBBeginTrans
				'''''            WKB030 = IMTX030.Text
				'            Call SCR_ADDNEW
				'            Call SpreadInit
				'            Call SCR_DSPDATA
				'            WKB030 = strN030
				'            ERRSW = F_ERR                   '   �J�[�\�����i�ނƂ��̓G���[
				'            Exit Sub
				'A-CUST-20100610 End
			Else '   �C���E�폜�̂Ƃ��͂n�j
				If CUR_NO = N999 Then '   �����敪�̂Ƃ��͂��̐�N���A����B
					Exit Sub
				End If
				'A-CUST-20170203 Start
				If KBKBN = F_DEL Then
					Call FILGET_JAN_HENKAN_2(KB.Inc_code, KB.jg_code, KB.hin_code)
					If ENDSW = F_END Then
						Exit Sub
					End If
					If JAN_HENKANINVSW = F_GET Then
						If CUR_NO < LST_NO Then '   ��ɖ߂�Ƃ��͌��̒l�ɖ߂��ďI���
							IMTX030.Text = WKB030
							Exit Sub
						End If
						Call MsgBox("JAN�ϊ��e�[�u���Ƀf�[�^�����݂��܂��B�폜�͂ł��܂���B", MsgBoxStyle.Exclamation)
						ERRSW = F_ERR
						Exit Sub
					End If
				End If
				'A-CUST-20170203e
			End If
		End If
		
		'   OK�����Ǐ�֖߂�Ƃ��A                  '   2000/01/22  �ǉ�
		If CUR_NO < LST_NO Then
			IMTX030.Text = strN030
			IMTX040.Text = KB.hin_name
			KB.hin_code = strN030
			Exit Sub
		End If
		
		'   �m��
		KB.hin_code = strN030
		WKB030 = strN030
		If KBKBN = F_ADD Then
			
			SetLookupMode(False)
			'   FALSE�ɂ����DEL_FLG=1�̂��̂��ǂ߂�B
			
			Call SCR_ADDNEW()
			Call SpreadInit()
			Call SCR_DSPDATA()
			'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N300).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			'Debug.Print("E" & CTRLTBL(N300).CTRL.Name & CTRLTBL(N300).CTRL.Index) 'D-20250417

			Call SCR_BUSHO(False, WKB030)
			
			SetLookupMode(True)
			'   TRUE�ɂ����DEL_FLG=1�̂��͓̂ǂ߂Ȃ��B
			
			Call OptionRefresh()
			'   TAB�ŏ���TAB�ɐݒ�          NR-SZ0410-2
			TAB010.SelectedIndex = 0
			
		Else
			SetLookupMode(False)
			'   FALSE�ɂ����DEL_FLG=1�̂��̂��ǂ߂�B
			Call SpreadInit()
			Call SCR_DSPDATA()
			JANCODESV = RTrim(KB.jan_code) 'A-CUST-20170203
			Call SCR_BUSHO(True, WKB030)
			
			SetLookupMode(True)
			'   TRUE�ɂ����DEL_FLG=1�̂��͓̂ǂ߂Ȃ��B
			
			Call OptionRefresh()
			'   TAB�ŏ���TAB�ɐݒ�          NR-SZ0410-2
			TAB010.SelectedIndex = 0
			
			SentakuFLG = False 'A-CUST-20100610
		End If
		
		If KBKBN = 3 Then Call SetMode("D")
		
		
	End Sub
	
	Private Sub OptionRefresh()
		
		Dim bBef As Boolean
		Dim bAft As Boolean

		'   �Ǘ��敪�|����CTRLTBL
		'bBef = OPTO300(1).Value 'D-20250417
		bBef = OPTO300(1).Checked 'A-20250417
		'CTRLTBL(N300).CTRL = IIf(OPTO300(1).Value, OPTO300(1).Value, OPTO300(2).Value) 'D-20250417
		CTRLTBL(N300).CTRL = IIf(OPTO300(1).Checked, OPTO300(1).Checked, OPTO300(2).Checked) 'A-20250417
		'bAft = OPTO300(1).Value 'D-20250417
		bAft = OPTO300(1).Checked 'A-20250417
		'System.Diagnostics.Debug.Assert(bBef = bAft, "") 'D-20250417

		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N300).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'System.Diagnostics.Debug.Assert(WKB300 = CTRLTBL(N300).CTRL.Index, "") 'D-20250417

		'   ����ŋ敪�|�O��
		'CTRLTBL(N310).CTRL = IIf(OPTO310(1).Value, OPTO310(1).Value, IIf(OPTO310(2).Value, OPTO310(2).Value, OPTO310(3).Value)) 'D-20250417
		CTRLTBL(N310).CTRL = IIf(OPTO310(1).Checked, OPTO310(1).Checked, IIf(OPTO310(2).Checked, OPTO310(2).Checked, OPTO310(3).Checked)) 'A-20250417
		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N310).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'System.Diagnostics.Debug.Assert(WKB310 = CTRLTBL(N310).CTRL.Index, "") 'D-20250417


		'   �I���P���|�d���P��
		'CTRLTBL(N320).CTRL = IIf(OPTO320(1).Value, OPTO320(1).Value, OPTO320(2).Value) 'D-20250417
		CTRLTBL(N320).CTRL = IIf(OPTO320(1).Checked, OPTO320(1).Checked, OPTO320(2).Checked) 'A-20250417
		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N320).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'System.Diagnostics.Debug.Assert(WKB320 = CTRLTBL(N320).CTRL.Index, "") 'D-20250417

		'   �݌ɊǗ��|����
		'CTRLTBL(N330).CTRL = IIf(OPTO330(1).Value, OPTO330(1).Value, OPTO330(2).Value) 'D-20250417
		CTRLTBL(N330).CTRL = IIf(OPTO330(1).Checked, OPTO330(1).Checked, OPTO330(2).Checked) 'D-20250417
		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N330).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'System.Diagnostics.Debug.Assert(WKB330 = CTRLTBL(N330).CTRL.Index, "") 'D-20250417

		'   FAX���M�|����
		'CTRLTBL(N340).CTRL = IIf(OPTO340(1).Value, OPTO340(1).Value, OPTO340(2).Value) 'D-20250417
		CTRLTBL(N340).CTRL = IIf(OPTO340(1).Checked, OPTO340(1).Checked, OPTO340(2).Checked) 'D-20250417
		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N340).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'System.Diagnostics.Debug.Assert(WKB340 = CTRLTBL(N340).CTRL.Index, "") 'D-20250417

	End Sub
	
	Private Sub QUE_KAISHA()
		
		Dim Ret As Integer
		Dim strCode As String
		
		CM9500.CM9500_TOP = VB6.PixelsToTwipsY(Me.Top)
		CM9500.CM9500_LEFT = VB6.PixelsToTwipsX(Me.Left)
		CM9500.CM9500_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
		CM9500.CM9500_WIDTH = VB6.PixelsToTwipsX(Me.Width)
		CM9500.CM9500_POS = 0
		CM9500.CM9500_RCN = ZACN_RCN
		CM9500.CM9500_TIME = 0
		Ret = CM9500.CM9500_SUB
		
		NXT_NO = N010
		Call FOCUS_SET()
		System.Diagnostics.Debug.Assert(LST_NO = 2, "")
		Debug.Print("LST_NO=2:1")
		strCode = CM9500.CM9500_SEL_CODE
		If Ret = n0 Then
			IMTX010.Text = strCode
			System.Diagnostics.Debug.Assert(LST_NO = 2, "")
			Debug.Print("LST_NO=2:2")
			Call SET_NO(1)
			System.Diagnostics.Debug.Assert(LST_NO = 2, "")
			Debug.Print("LST_NO=2:3")
		Else
			NXT_NO = N010
			Call FOCUS_SET()
		End If
		
	End Sub
	
	Private Sub QUE_JIGYO()
		
		Dim Ret As Integer
		Dim strCode As String
		
		CM9510.CM9510_TOP = VB6.PixelsToTwipsY(Me.Top)
		CM9510.CM9510_LEFT = VB6.PixelsToTwipsX(Me.Left)
		CM9510.CM9510_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
		CM9510.CM9510_WIDTH = VB6.PixelsToTwipsX(Me.Width)
		CM9510.CM9510_POS = 0
		CM9510.CM9510_RCN = ZACN_RCN
		CM9510.CM9510_TIME = 0
		CM9510.CM9510_INC_CODE = WKB010
		CM9510.CM9510_INC_NAME = DSP010.Text
		Ret = CM9510.CM9510_SUB
		
		NXT_NO = N020
		Call FOCUS_SET()
		
		strCode = CM9510.CM9510_SEL_CODE
		If Ret = n0 Then
			IMTX020.Text = strCode
			System.Windows.Forms.Application.DoEvents()
			Call SET_NO(1)
			System.Windows.Forms.Application.DoEvents()
		Else
			NXT_NO = N020
			Call FOCUS_SET()
		End If
		Debug.Print("QUE_JIGYO=" & Ret)
		
	End Sub
	
	
	'   �啪�ށA�����ށA�����ނ̖⍇��
	'���ޖ⍇���ǉ�
	Public Function QUE_BUNRUI(ByRef IDX As Short) As Short
		
		Dim Ret As Short
		
		
		Select Case IDX
			Case N220
				SZ0740.SZ0740_TOP = VB6.PixelsToTwipsY(Me.Top)
				SZ0740.SZ0740_LEFT = VB6.PixelsToTwipsX(Me.Left)
				SZ0740.SZ0740_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
				SZ0740.SZ0740_WIDTH = VB6.PixelsToTwipsX(Me.Width)
				SZ0740.SZ0740_POS = 0
				SZ0740.SZ0740_RCN = ZACN_RCN
				SZ0740.SZ0740_TIME = 0
				SZ0740.SZ0740_INC_CODE = WKB010
				SZ0740.SZ0740_INC_NAME = DSP010.Text
				Ret = SZ0740.SZ0740_SUB
				If Ret = 0 Then
					SEL_FIND = SZ0740.SZ0740_SEL_CODE
				Else
					SEL_FIND = ""
				End If
			Case N230
				SZ0750.SZ0750_TOP = VB6.PixelsToTwipsY(Me.Top)
				SZ0750.SZ0750_LEFT = VB6.PixelsToTwipsX(Me.Left)
				SZ0750.SZ0750_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
				SZ0750.SZ0750_WIDTH = VB6.PixelsToTwipsX(Me.Width)
				SZ0750.SZ0750_POS = 0
				SZ0750.SZ0750_RCN = ZACN_RCN
				SZ0750.SZ0750_TIME = 0
				SZ0750.SZ0750_INC_CODE = WKB010
				SZ0750.SZ0750_INC_NAME = DSP010.Text
				SZ0750.SZ0750_D_CODE = IMTX220.Text
				SZ0750.SZ0750_D_NAME = DSP220.Text
				Ret = SZ0750.SZ0750_SUB
				If Ret = 0 Then
					SEL_FIND = SZ0750.SZ0750_SEL_CODE
				Else
					SEL_FIND = ""
				End If
			Case N240
				SZ0760.SZ0760_TOP = VB6.PixelsToTwipsY(Me.Top)
				SZ0760.SZ0760_LEFT = VB6.PixelsToTwipsX(Me.Left)
				SZ0760.SZ0760_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
				SZ0760.SZ0760_WIDTH = VB6.PixelsToTwipsX(Me.Width)
				SZ0760.SZ0760_POS = 0
				SZ0760.SZ0760_RCN = ZACN_RCN
				SZ0760.SZ0760_TIME = 0
				SZ0760.SZ0760_INC_CODE = WKB010
				SZ0760.SZ0760_INC_NAME = DSP010.Text
				SZ0760.SZ0760_D_CODE = IMTX220.Text
				SZ0760.SZ0760_D_NAME = DSP220.Text
				SZ0760.SZ0760_C_CODE = IMTX230.Text
				SZ0760.SZ0760_C_NAME = DSP230.Text
				Ret = SZ0760.SZ0760_SUB
				If Ret = 0 Then
					SEL_FIND = SZ0760.SZ0760_SEL_CODE
				Else
					SEL_FIND = ""
				End If
				'02/05/28 ADD START
				'D-20250201��
				'Case N250
				'CM9600.CM9600_TOP = Me.Top
				'CM9600.CM9600_LEFT = Me.Left
				'CM9600.CM9600_HEIGHT = Me.Height
				'CM9600.CM9600_WIDTH = Me.Width
				'CM9600.CM9600_POS = 0
				'Set CM9600.CM9600_RCN = ZACN_RCN
				'CM9600.CM9600_TIME = 0
				'CM9600.CM9600_INC_CODE = WKB010
				'CM9600.CM9600_INC_NAME = DSP010
				'Ret = CM9600.CM9600_SUB
				'If Ret = 0 Then
				'SEL_FIND = CM9600.CM9600_SEL_CODE
				'Else
				'SEL_FIND = ""
				'End If
				'D-20250201��
				'02/05/28 ADD END
		End Select
		
		
		'    Select Case idx
		'        Case N220
		'            SEL_TYPE = "DAIBUNRUI"
		'        Case N230
		'            SEL_TYPE = "CHUBUNRUI"
		'            SEL_CODE = KB.l_bun_code
		'        Case N240
		'            SEL_TYPE = "SHOBUNRUI"
		'            SEL_CODE = KB.l_bun_code
		'            SEL_CODE2 = KB.m_bun_code
		'        Case Else
		'            Exit Function
		'    End Select
		'
		'    SZ0410GFRM.Show vbModal
		
		If SEL_FIND <> "" Then
			Select Case IDX
				Case N220
					IMTX220.Text = SEL_FIND
				Case N230
					IMTX230.Text = SEL_FIND
				Case N240
					IMTX240.Text = SEL_FIND
					'D-20250201��
					'Case N250                       '02/05/28 ADD
					'IMTX250.Text = SEL_FIND     '02/05/28 ADD
					'D-20250201��
			End Select
			Call SET_NO(1)
			QUE_BUNRUI = F_OFF
		Else
			QUE_BUNRUI = -1
		End If
		
	End Function
	
	Public Function QUE_BUSHO() As Short
		
		Dim Ret As Integer
		
		CM9520.CM9520_TOP = VB6.PixelsToTwipsY(Me.Top)
		CM9520.CM9520_LEFT = VB6.PixelsToTwipsX(Me.Left)
		CM9520.CM9520_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
		CM9520.CM9520_WIDTH = VB6.PixelsToTwipsX(Me.Width)
		CM9520.CM9520_POS = 0
		CM9520.CM9520_RCN = ZACN_RCN
		CM9520.CM9520_TIME = 0
		CM9520.CM9520_INC_CODE = WKB010
		CM9520.CM9520_INC_NAME = DSP010.Text
		CM9520.CM9520_JG_CODE = WKB020
		CM9520.CM9520_JG_NAME = DSP020.Text
		CM9520.CM9520_SKBN = 3
		Ret = CM9520.CM9520_SUB
		If Ret = 0 Then
			SEL_FIND = CM9520.CM9520_SEL_CODE
		Else
			SEL_FIND = ""
		End If
		
		'    SZ0410BFRM.Show vbModal
		
		If SEL_FIND <> "" Then
			SpreadAppend()
		Else
			Call SpreadZeroTrim(-1)
		End If
		
	End Function
	
	
	
	Private Sub IPROCHK_N040()
		
		'    If CUR_NO < LST_NO Then
		'        IMTX040.Text = KB.hin_name
		'        Exit Sub
		'    End If
		
		'   �i��NOTNULL�`�F�b�N
		If RTrim(IMTX040.Text) = "" Then
			If CUR_NO < LST_NO Then
				IMTX040.Text = KB.hin_name
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   �i���̊m��
		KB.hin_name = IMTX040.Text
		'A-CUST-20100610 Start
		If RTrim(IMTX065.Text) = "" Then
			IMTX065.Text = RTrim(KB.hin_name)
			KB.hin_name_seisiki = KB.hin_name
		End If
		'A-CUST-20100610 End
		
	End Sub
	
	Private Sub IPROCHK_N050()
		
		'    If CUR_NO < LST_NO Then
		'        IMTX050.Text = KB.kikaku
		'        Exit Sub
		'    End If
		
		'   �i�ԑ��݃`�F�b�N
		
		'   �K�i�̊m��
		KB.kikaku = IMTX050.Text
		
	End Sub
	
	Private Sub IPROCHK_N060()
		
		'    If CUR_NO < LST_NO Then
		'        Call COMBO_SETLIST(CMB060, KB.tani)
		'        Exit Sub
		'    End If
		
		'   �P�ʃ`�F�b�N
		If Trim(CMB060.Text) = "" Then
			If CUR_NO < LST_NO Then
				Call COMBO_SETLIST(CMB060, KB.tani)
				''''CMB060.Text = KB.tani
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   �P�ʂ̊m��
		If Trim(CMB060.Text) <> "" Then
			KB.tani = CMB060.Text
		Else
			CMB060.SelectedIndex = 0
			'        CMB060.Text = KB.tani
			'        ERRSW = F_ERR
			'        ENDSW = F_END
			
		End If
		
	End Sub
	
	'A-CUST-20100610 Start
	Private Sub IPROCHK_N065()
		KB.hin_name_seisiki = IMTX065.Text
		
	End Sub
	'A-CUST-20100610 End
	
	Private Sub IPROCHK_N070()
		
		If CUR_NO < LST_NO Then
			IMTX070.Text = KB.jan_code
			Exit Sub
		End If
		
		
		'A-CUST20130212��
		If RTrim(KB.jan_code) = RTrim(IMTX070.Text) Then Exit Sub
		Dim chk_jan_hincode As String
		If RTrim(IMTX070.Text) = "" Then
			'A-CUST-20170203 Start
			If KBKBN = F_REP Then
				Call FILGET_JAN_HENKAN_2(KB.Inc_code, KB.jg_code, KB.hin_code)
				If ENDSW = F_END Then
					Exit Sub
				End If
				If JAN_HENKANINVSW = F_GET Then
					Call MsgBox("JAN�ϊ��e�[�u���Ƀf�[�^�����݂��܂��BJAN�W���R�[�h�͏ȗ��ł��܂���B", MsgBoxStyle.Exclamation)
					ERRSW = F_ERR
					Exit Sub
				End If
			End If
			'A-CUST-20170203e
			'JAN�W�����ނ̊m��
			KB.jan_code = IMTX070.Text
			Exit Sub
		Else
			
			'A-20250303��
			'JAN�R�[�h�d���`�F�b�N
			If RTrim(IMTX070.Text) <> "" Then
				chk_jan_hincode = CHK_JANCODE((IMTX070.Text))
				If chk_jan_hincode <> "" Then
					Call MsgBox("���̕i�Ԃœ���JAN�W�����ނ��g�p����Ă��܂��B" & vbCrLf & "�i��[" & chk_jan_hincode & "]", MsgBoxStyle.ApplicationModal + MsgBoxStyle.Exclamation, "�d���i�ڊ�{������")
					ERRSW = F_ERR
					Exit Sub
				End If
			End If
			'A-20250303��
			
			'A-CUST-20170203 Start
			If KBKBN = F_REP Then
				Call FILGET_JAN_HENKAN_1(KB.Inc_code, KB.jg_code, KB.hin_code, IMTX070.Text)
				If ENDSW = F_END Then
					Exit Sub
				End If
				If JAN_HENKANINVSW = F_GET Then
					Call MsgBox("JAN�ϊ��e�[�u���̖��׃f�[�^�ɑ��݂��Ă��܂��B", MsgBoxStyle.Exclamation)
					ERRSW = F_ERR
					Exit Sub
				End If
			End If
			'A-CUST-20170203e
			
			JAN_BUF0.k4 = IMTX070.Text
			If FILGET_JAN() = False Then
				'D-20130501-S
				'            'JAN���ނ����݂��Ȃ������ꍇ
				'             ERRSW = F_ERR
				'             Exit Sub
				'D-20130501-E
				'A-20130501-S
				'JAN�W�����ނ̊m��
				KB.jan_code = IMTX070.Text
				'A-20130501-E
			Else
				
				'JAN�W�����ނ̊m��
				KB.jan_code = IMTX070.Text
				
				If RTrim(KB.BK1) = "" And KB.k42 = 0 And RTrim(KB.k44) = "" And RTrim(KB.k57) = "" And KB.k58 = 0 Then
				Else
					If MsgBox("JANϽ��֘A���ڂ��X�V���Ă���낵���ł��傤���H", MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question, "�d���i�ڊ�{������") = MsgBoxResult.No Then Exit Sub
				End If
				KB.BK1 = JAN.k21
				KB.k44 = JAN.k44
				KB.k42 = JAN.k42
				KB.k57 = JAN.k57
				KB.k58 = JAN.k58
				IMTX150(0).Text = KB.k44
				IMNU160(0).Value = KB.k42
				IMNU170(1).Value = KB.k58
				IMTX291.Text = KB.BK1
				'D-20130424-S
				'            If Trim(JAN.k14) <> "" Then
				'                KB.hin_name_seisiki = JAN.k14
				'                IMTX065.Text = KB.hin_name_seisiki
				'            End If
				'D-20130424-E
				'A-20130424-S
				If Trim(JAN.k17) <> "" Then
					KB.hin_name_seisiki = JAN.k17
					IMTX065.Text = KB.hin_name_seisiki
				End If
				'A-20130424-E
				
				'���t���Z
				KB.k99 = 0 '�v�Z�O�ɃN���A
				DSP170(0).Text = CStr(0) '�v�Z�O�ɃN���A
				If RTrim(KB.k57) = "" Then
					CMB170.SelectedIndex = -1
				Else
					Select Case KB.k57
						Case CStr(1)
							CMB170.Text = "��"
						Case CStr(2)
							CMB170.Text = "��"
						Case CStr(3)
							CMB170.Text = "�N"
						Case Else
							CMB170.SelectedIndex = -1
					End Select
				End If
				Call CNV_DAY() '�����Z����
				'JAN���i���ޖ��擾
				DSP291.Text = "" '�N���A
				JAN_BUNRUI_BUF0.BK1 = KB.BK1
				If FILGET_JAN_BUNRUI() = True Then
					DSP291.Text = JAN_BUNRUI.BK4 '���ޖ�
				End If
				
			End If
		End If
		'A-CUST20130212��
		
	End Sub
	
	Private Sub IPROCHK_N080()
		
		'    If CUR_NO < LST_NO Then
		'        IMTX080.Text = KB.jan_s_code
		'        Exit Sub
		'    End If
		
		'   JAN�Z�k�̊m��
		KB.jan_s_code = IMTX080.Text
		
	End Sub
	
	Private Sub IPROCHK_N090()
		
		'    If CUR_NO < LST_NO Then
		'        IMTX090.Text = KB.bar_code
		'        Exit Sub
		'    End If
		
		'   ���̑����ް���ނ̊m��
		KB.bar_code = IMTX090.Text
		
	End Sub
	
	Private Sub IPROCHK_N100(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim strDate As String
		
		'    If CUR_NO < LST_NO Then
		'        If idx = N100_1 Then
		'            IMTX100(1).Text = DateSlashed(KB.teki_date1)
		'        Else
		'            IMTX100(2).Text = DateSlashed(KB.teki_date2)
		'        End If
		'        Exit Sub
		'    End If
		
		'�K�p���A�́A�󕶂n�j
		If IDX = N100_2 And Trim(IMTX100(2).Text) = "" Then
			IMNU110(2).Value = 0
			IMNU120(2).Value = 0
			KB.baika2 = IMNU110(2).Value
			KB.kei_kin2 = IMNU120(2).Value
			GoTo IPROCHK_N100_L
		End If
		strDate = IIf(IDX = N100_1, IMTX100(1).Text, IMTX100(2).Text)
		iReturn = CHK_DATE(strDate)
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				If IDX = N100_1 Then
					IMTX100(1).Text = DateSlashed(KB.teki_date1)
				Else
					IMTX100(2).Text = DateSlashed(KB.teki_date2)
				End If
				Exit Sub
			End If
			NXT_NO = IDX
			ERRSW = F_ERR
			Exit Sub
			
		End If
		
IPROCHK_N100_L: 
		
		'   �����K�p���P�C�Q�̊m��
		If IDX = N100_1 Then
			KB.teki_date1 = IMTX100(1).Text
			IMTX100(1).Text = DateSlashed(KB.teki_date1)
		Else
			KB.teki_date2 = IMTX100(2).Text
			IMTX100(2).Text = DateSlashed(KB.teki_date2)
		End If
		
	End Sub
	
	Private Sub IPROCHK_N110(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim lBaika As Integer
		
		'    If CUR_NO < LST_NO Then
		'        If idx = N110_1 Then
		'            IMNU110(1).Value = KB.baika1
		'        Else
		'            IMNU110(2).Value = KB.baika2
		'        End If
		'        Exit Sub
		'    End If
		
		lBaika = IIf(IDX = N110_1, IMNU110(1).Value, IMNU110(2).Value)
		iReturn = CHK_AMOUNT(lBaika)
		
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				If IDX = N110_1 Then
					IMNU110(1).Value = KB.baika1
				Else
					IMNU110(2).Value = KB.baika2
				End If
				Exit Sub
			End If
			'        NXT_NO = idx
			ERRSW = F_ERR
			Exit Sub
			
		End If
		
		'   �����P�C�Q�̊m��
		If IDX = N110_1 Then
			KB.baika1 = IMNU110(1).Value
		Else
			KB.baika2 = IMNU110(2).Value
		End If
		
	End Sub
	
	Private Sub IPROCHK_N120(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim curKin As Decimal
		
		'    If CUR_NO < LST_NO Then
		'        If idx = N120_1 Then
		'            IMNU120(1).Value = KB.kei_kin1
		'        Else
		'            IMNU120(2).Value = KB.kei_kin2
		'        End If
		'        Exit Sub
		'    End If
		
		curKin = IIf(IDX = N120_1, IMNU120(1).Value, IMNU120(2).Value)
		iReturn = CHK_CURRENCY(curKin)
		
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				If IDX = N120_1 Then
					IMNU120(1).Value = KB.kei_kin1
				Else
					IMNU120(2).Value = KB.kei_kin2
				End If
				Exit Sub
			End If
			'        NXT_NO = idx
			ERRSW = F_ERR
			Exit Sub
			
		End If
		
		'   �����P�C�Q�̊m��
		If IDX = N120_1 Then
			KB.kei_kin1 = IMNU120(1).Value
		Else
			KB.kei_kin2 = IMNU120(2).Value
		End If
		
	End Sub
	
	Private Sub IPROCHK_N130N140(ByRef IDX As Short)
		'
		'   IMTX1300, IMTX140 �� GRP7 �ł��B
		
		Dim strName As String
		
		'               2000/01/25  DEL     KOKOKARA
		'    If CUR_NO < LST_NO Then
		'        If idx = N130 Then
		'            IMTX130(1).Text = KB.hiyou_k_code1
		'        Else
		'            IMTX140(1).Text = KB.hiyou_k_code2
		'            DSP140(1).Caption = WKB140DSP
		'        End If
		'        Exit Sub
		'    End If
		'               2000/01/25  DEL     KOKOMADE
		
		If IDX = N130 Then
			If Trim(IMTX130(1).Text) = "" Then
				If CUR_NO < LST_NO Then
					IMTX130(1).Text = KB.hiyou_k_code1
					IMTX140(1).Text = KB.hiyou_k_code2
					DSP140(1).Text = WKB140DSP
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		End If
		
		If IDX = N130 Then
			
			IMTX130(1).Text = VB6.Format(IMTX130(1).Text, "000")
			
			strName = DecodeKAMOCHU(IMTX130(1).Text)
			If strName <> "" Then
				WKAMOCHUNM = strName
				KB.hiyou_k_code1 = IMTX130(1).Text
				
				strName = DecodeKAMOKU(IMTX130(1).Text, IMTX140(1).Text)
				If strName <> "" Then
					DSP140(1).Text = WKAMOCHUNM & strName
				Else
					DSP140(1).Text = WKAMOCHUNM
					IMTX140(1).Text = ""
					KB.hiyou_k_code2 = IMTX140(1).Text
				End If
				WKB140DSP = DSP140(1).Text
				
			Else '   Error
				If CUR_NO < LST_NO Then
					IMTX130(1).Text = KB.hiyou_k_code1
					Exit Sub
				End If
				'            NXT_NO = idx
				ERRSW = F_ERR
			End If
		ElseIf IDX = N140 Then 
			
			IMTX140(1).Text = VB6.Format(IMTX140(1).Text, "000000")
			
			strName = DecodeKAMOKU(IMTX130(1).Text, IMTX140(1).Text)
			If strName <> "" Then
				DSP140(1).Text = WKAMOCHUNM & strName
				WKB140DSP = DSP140(1).Text
				KB.hiyou_k_code2 = IMTX140(1).Text
			Else '   Error
				If CUR_NO < LST_NO Then
					IMTX140(1).Text = KB.hiyou_k_code2
					DSP140(1).Text = WKB140DSP
					Exit Sub
				End If
				ERRSW = F_ERR
			End If
		End If
		
	End Sub
	'A-CUST20130212��
	Private Sub IPROCHK_N150()
		'   ���Y��
		'KB.k44 = IMTX150(0).Text   'D-20130401-

		'KB.k44 = StrConv(IMTX150(0).Text, VbStrConv.Uppercase) 'A-20130401-�啶���ɕϊ� 'D-20250417
		KB.k44 = Microsoft.VisualBasic.StrConv(IMTX150(0).Text, VbStrConv.Uppercase) 'A-20130401-�啶���ɕϊ� 'A-20250417
		IMTX150(0).Text = KB.k44 'A-20130401-
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Private Sub IPROCHK_N160()
		'   �d��
		KB.k42 = Val(CStr(IMNU160(0).Value))
		IMNU160(0).Value = KB.k42
		
	End Sub
	'A-CUST20130212��
	
	'A-20240115��
	Private Sub IPROCHK_N165()
		'   ����/�ܖ�����
		KB.Shomi_date_kbn = CStr(VB6.GetItemData(CMB165, CMB165.SelectedIndex))
	End Sub
	'A-20240115��
	
	'A-CUST20130212��
	Private Sub IPROCHK_N170CMB()
		'A-20240115��
		If CUR_NO < LST_NO Then
			If CMB170.SelectedIndex = 0 Or CMB170.SelectedIndex = -1 Then
				CMB170.SelectedIndex = IIf(Trim(KB.k57) = "", -1, KB.k57)
			Else
				KB.k57 = CStr(VB6.GetItemData(CMB170, CMB170.SelectedIndex))
			End If
		Else
			If CDbl(RTrim(CStr(CMB165.SelectedIndex))) <> 0 Then
				If CMB170.SelectedIndex = -1 Or CMB170.SelectedIndex = 0 Then
					ERRSW = F_ERR
					Exit Sub
				End If
			End If
		End If
		'A-20240115��
		
		'   �ܖ������敪
		If CMB170.SelectedIndex = -1 Or CMB170.SelectedIndex = 0 Then
			KB.k57 = " "
			KB.k58 = 0
			KB.k99 = 0
			IMNU170(1).Value = KB.k58
			DSP170(0).Text = CStr(0)
			Exit Sub
		End If
		If Val(CStr(IMNU170(1).Value)) <> 0 Then
			Call CNV_DAY() '�����Z����
		End If
		
		KB.k57 = CStr(VB6.GetItemData(CMB170, CMB170.SelectedIndex))
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Private Sub IPROCHK_N170()
		'   �ܖ�����
		'A-20240115��
		If Not CUR_NO < LST_NO Then
			If Val(CStr(IMNU170(1).Value)) = 0 And CDbl(RTrim(CStr(CMB165.SelectedIndex))) <> 0 Then
				ERRSW = F_ERR
				Exit Sub
			End If
		End If
		'A-20240115��
		
		If Val(CStr(IMNU170(1).Value)) <> 0 Then
			Call CNV_DAY() '�����Z����
		Else
			KB.k99 = 0
			DSP170(0).Text = CStr(KB.k99)
		End If
		
		KB.k58 = Val(CStr(IMNU170(1).Value))
		IMNU170(1).Value = KB.k58
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Private Sub IPROCHK_N291()
		'   JAN���i����
		
		If CUR_NO < LST_NO Then
			IMTX291.Text = JAN_BUNRUI_BUF0.BK1
			Exit Sub
		End If
		
		
		'    If KB.k21 = IMTX291.Text Then Exit Sub
		If RTrim(IMTX291.Text) = "" Then
			DSP291.Text = ""
			Exit Sub
		End If
		
		JAN_BUNRUI_BUF0.BK1 = IMTX291.Text
		If FILGET_JAN_BUNRUI() = True Then
			DSP291.Text = JAN_BUNRUI.BK4 '���ޖ�
		Else
			DSP291.Text = ""
			ERRSW = F_ERR
			Exit Sub
		End If
		
		KB.BK1 = IMTX291.Text
		
	End Sub
	'A-CUST20130212��
	
	'D-20250201��
	'Private Sub IPROCHK_N210N211(IDX As Integer)
	'
	'   IMTX210, IMTX211 �� GRP8 �ł��B
	'Dim strName As String
	'Dim strCode As String
	
	'                               2000/01/25  DEL     KOKOKARA
	'    If CUR_NO < LST_NO Then
	'        If idx = N210 Then
	'            IMTX210.Text = Left(KB.ka_bun_code, 3)
	'        Else
	'            IMTX211.Text = Right(KB.ka_bun_code, 4)
	'            DSP210.Caption = WKB210DSP
	'        End If
	'        Exit Sub
	'    End If
	
	'    If idx = N210 Then
	'        If Trim(IMTX210.Text) = "" Then
	'            ERRSW = F_ERR
	'            Exit Sub
	'        End If
	'    End If
	'                               2000/01/25  DEL     KOKOMADE
	
	'If IDX = N210 Then
	'IMTX210.Text = Format(IMTX210.Text, "000")
	'strCode = IMTX210.Text & IMTX211.Text
	'strName = DecodeKamBunrui(WKB010, WKB020, strCode)
	'If strName <> "" Then
	'DSP210.Caption = strName
	'WKB210DSP = strName
	'KB.ka_bun_code = strCode
	'Else
	'If CUR_NO < LST_NO Then
	'IMTX210.Text = Left(KB.ka_bun_code, 3)
	'Exit Sub
	'End If
	''''ERRSW = F_ERR
	'End If
	
	'ElseIf IDX = N211 Then
	
	'   ����͂��Ȃ��Ă悢�B
	
	'strCode = IMTX210.Text & IMTX211.Text
	'strName = DecodeKamBunrui(WKB010, WKB020, strCode)
	'If strName = "" Then
	'If CUR_NO < LST_NO Then
	'If IDX = N210 Then
	'IMTX210.Text = Left(KB.ka_bun_code, 3)
	'Else
	'IMTX211.Text = Right(KB.ka_bun_code, 4)
	'DSP210.Caption = WKB210DSP
	'End If
	'Exit Sub
	'End If
	'           IMTX210.Text = Left(KB.ka_bun_code, 3)
	'           IMTX211.Text = Right(KB.ka_bun_code, 4)
	'            DSP210.Caption = ""
	'ERRSW = F_ERR
	'Else
	'DSP210.Caption = strName
	'WKB210DSP = strName
	'KB.ka_bun_code = strCode
	'End If
	'End If
	
	'End Sub
	'D-20250201��
	
	Private Sub IPROCHK_N220N230N240(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim strCode As String
		
		'                           2000/01/25      DEL     KOKOKARA
		'    If CUR_NO < LST_NO Then
		'        If idx = N220 Then
		'            IMTX220.Text = KB.l_bun_code
		'            DSP220.Caption = WKB220DSP
		'        ElseIf idx = N230 Then
		'            IMTX230.Text = KB.m_bun_code
		'            DSP230.Caption = WKB230DSP
		'        Else
		'            IMTX240.Text = KB.s_bun_code
		'            DSP240.Caption = WKB240DSP
		'        End If
		'        Exit Sub
		'    End If
		'                           2000/01/25      DEL     KOKOMADE
		
		'   ������͐����ł͂Ȃ��B
		'    Select Case idx
		'        Case N220
		'            IMTX220.Text = Format(IMTX220.Text, "0000")
		'        Case N230
		'            IMTX230.Text = Format(IMTX230.Text, "0000")
		'        Case N240
		'            IMTX240.Text = Format(IMTX240.Text, "0000")
		'    End Select
		
		
		strCode = IIf(IDX = N220, IMTX220.Text, IIf(IDX = N230, IMTX230.Text, IMTX240.Text))
		
		If Trim(strCode) = "" Then
			If CUR_NO < LST_NO Then
				If IDX = N220 Then
					IMTX220.Text = KB.l_bun_code
					DSP220.Text = WKB220DSP
				ElseIf IDX = N230 Then 
					IMTX230.Text = KB.m_bun_code
					DSP230.Text = WKB230DSP
				Else
					IMTX240.Text = KB.s_bun_code
					DSP240.Text = WKB240DSP
				End If
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		Select Case IDX
			Case N220
				iReturn = CHK_BUNRUI(1, strCode, "", "")
			Case N230
				iReturn = CHK_BUNRUI(2, KB.l_bun_code, strCode, "")
			Case N240
				iReturn = CHK_BUNRUI(3, KB.l_bun_code, KB.m_bun_code, strCode)
			Case Else
				iReturn = F_OFF
		End Select
		
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				If IDX = N220 Then
					IMTX220.Text = KB.l_bun_code
					DSP220.Text = WKB220DSP
				ElseIf IDX = N230 Then 
					IMTX230.Text = KB.m_bun_code
					DSP230.Text = WKB230DSP
				Else
					IMTX240.Text = KB.s_bun_code
					DSP240.Text = WKB240DSP
				End If
				Exit Sub
			End If
			Select Case IDX
				Case N220
					DSP220.Text = ""
				Case N230
					DSP230.Text = ""
				Case N240
					DSP240.Text = ""
			End Select
			'        NXT_NO = idx
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   ��ʃ��x�����ς����牺�ʃ��x�����N���A
		Select Case IDX
			Case N220
				If KB.l_bun_code <> strCode Then
					KB.m_bun_code = ""
					WKB230DSP = ""
					IMTX230.Text = ""
					DSP230.Text = ""
					KB.s_bun_code = ""
					WKB240DSP = ""
					IMTX240.Text = ""
					DSP240.Text = ""
				End If
			Case N230
				If KB.m_bun_code <> strCode Then
					KB.s_bun_code = ""
					WKB240DSP = ""
					IMTX240.Text = ""
					DSP240.Text = ""
				End If
		End Select
		
		'   �啪�ށA�����ށA�����ނ̊m��
		Select Case IDX
			Case N220
				KB.l_bun_code = strCode
			Case N230
				KB.m_bun_code = strCode
				Call SCR_DSPTAX() 'A-20190601
			Case N240
				KB.s_bun_code = strCode
		End Select
		
	End Sub
	
	'D-20250201��
	'Private Sub IPROCHK_N250()
	'����
	'02/05/28 ADD START
	'Dim strReturn As String
	'Dim strBunrui As String
	'Dim iReturn As Integer
	
	'strBunrui = RTrim(IMTX250.Text)
	'If strBunrui <> "" Then
	'strBunrui = ZeroFill(strBunrui, 4)
	'iReturn = DecodeBUNRUI(strBunrui, strReturn)
	'If strReturn = "" Then
	'If CUR_NO < LST_NO Then
	'IMTX250.Text = KB.bun_code
	'DSP250.Caption = WKB250DSP
	'Exit Sub
	'End If
	'DSP250.Caption = ""
	'ERRSW = F_ERR
	'Exit Sub
	'End If
	'Else
	'If CUR_NO < LST_NO Then
	'IMTX250.Text = KB.bun_code
	'DSP250.Caption = WKB250DSP
	'Exit Sub
	'End If
	'        ERRSW = F_ERR  '�󔒂̏ꍇ�G���[�Ƃ��Ȃ�
	'WKB250DSP = ""
	'KB.bun_code = ""
	'IMTX250.Text = ""
	'DSP250.Caption = ""
	'Exit Sub
	'End If
	
	'   �m��
	'DSP250.Caption = strReturn
	'WKB250DSP = strReturn
	'IMTX250.Text = strBunrui
	
	'KB.bun_code = strBunrui
	'02/05/28 ADD END
	'02/05/28 DEL START
	''   ����
	''   ���⒆ NOCHECK
	'    KB.bun_code = IMTX250.Text
	'02/05/28 DEL END
	'End Sub
	'D-20250201��
	
	Private Sub IPROCHK_N260()
		'   ��������
		Dim strReturn As String
		Dim strFind As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX260.Text = KB.ken_bun_code
		'        DSP260.Caption = WKB260DSP
		'        Exit Sub
		'    End If
		
		strFind = RTrim(IMTX260.Text)
		If strFind <> "" Then
			strFind = ZeroFill(strFind, 4)
			strReturn = DecodeFIND(strFind)
			If strReturn = "" Then
				If CUR_NO < LST_NO Then
					IMTX260.Text = KB.ken_bun_code
					DSP260.Text = WKB260DSP
					Exit Sub
				End If
				DSP260.Text = ""
				ERRSW = F_ERR
				'            ENDSW = F_END
				Exit Sub
			End If
		Else
			'           2000/01/30  Add
			If CUR_NO < LST_NO Then
				IMTX260.Text = KB.ken_bun_code
				DSP260.Text = WKB260DSP
				Exit Sub
			End If
			'           2000/01/30  Add
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   �m��
		DSP260.Text = strReturn
		WKB260DSP = strReturn
		IMTX260.Text = strFind
		
		KB.ken_bun_code = strFind
		
	End Sub
	
	
	Private Sub IPROCHK_CHKBTN(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim strCode As String
		
		'   �m��
		Select Case IDX
			Case N270 '   ������i
				KB.jutaku = "" & CHK270.CheckState
			Case N280 '   �d�|�敪
				KB.sikakari = "" & CHK280.CheckState
			Case N290 '   ���c����
				KB.zan = "" & CHK290.CheckState
			Case N430 '   ���ꔭ����
				KB.gen_h_ka = "" & CHK430.CheckState
			Case N450 '   �����i
				KB.tyozouhin = "" & CHK450.CheckState
			Case N460 '   ���̋@
				KB.jihan = "" & CHK460.CheckState
			Case N470 '   ����Ώ�
				KB.gensen = "" & CHK470.CheckState
			Case N500 '   �����x�~
				KB.tori_kyu = "" & CHK500.CheckState
				
		End Select
		
	End Sub
	
	Private Sub IPROCHK_OPTO(ByRef IDX As Short)
		
		'   NO OPERATION
		System.Diagnostics.Debug.Assert(IDX >= N300 And IDX <= N340, "")
		
	End Sub
	
	Private Sub IPROCHK_N350N360(ByRef IDX As Short)
		
		Dim iNo As Short
		
		'    If CUR_NO < LST_NO Then
		'        Select Case idx
		'            Case N350_1
		'                Call COMBO_SETLIST(CMB350(1), KB.ha_tanka1)
		'            Case N350_2
		'                Call COMBO_SETLIST(CMB350(2), KB.ha_tanka2)
		'            Case N350_3
		'                Call COMBO_SETLIST(CMB350(3), KB.ha_tanka3)
		'            Case N350_4
		'                Call COMBO_SETLIST(CMB350(4), KB.ha_tanka4)
		'            Case N350_5
		'                Call COMBO_SETLIST(CMB350(5), KB.ha_tanka5)
		'            Case N360_1
		'                IMNU360(1).Value = KB.kansan_num1
		'            Case N360_2
		'                IMNU360(2).Value = KB.kansan_num2
		'            Case N360_3
		'                IMNU360(3).Value = KB.kansan_num3
		'            Case N360_4
		'                IMNU360(4).Value = KB.kansan_num4
		'            Case N360_5
		'                IMNU360(5).Value = KB.kansan_num5
		'        End Select
		'        Exit Sub
		'    End If
		
		Select Case IDX
			Case N350_1, N360_1
				iNo = 1
			Case N350_2, N360_2
				iNo = 2
			Case N350_3, N360_3
				iNo = 3
			Case N350_4, N360_4
				iNo = 4
			Case N350_5, N360_5
				iNo = 5
		End Select
		
		'�ŏ��̂ݕK�{       ���̃`�F�b�N�͎d�l�ύX�ɂ��p�~ 2000/02/23
		'    If idx = N350_1 Then
		'        If Trim(CMB350(1).Text) = "" Then
		'            If LST_NO > CUR_NO Then
		'                Call COMBO_SETLIST(CMB350(1), KB.ha_tanka1)
		'                Exit Sub
		'            End If
		'            ERRSW = F_ERR
		'            Exit Sub
		'        End If
		'    End If
		
		'   ���Z���̃`�F�b�N
		If IDX = N360_1 Then
			If IMNU360(1).Value = 0 Then
				If LST_NO > CUR_NO Then
					IMNU360(1).Value = KB.kansan_num1
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		ElseIf IDX = N360_2 Then 
			If IMNU360(2).Value = 0 Then
				If LST_NO > CUR_NO Then
					IMNU360(2).Value = KB.kansan_num2
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		ElseIf IDX = N360_3 Then 
			If IMNU360(3).Value = 0 Then
				If LST_NO > CUR_NO Then
					IMNU360(3).Value = KB.kansan_num3
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		ElseIf IDX = N360_4 Then 
			If IMNU360(4).Value = 0 Then
				If LST_NO > CUR_NO Then
					IMNU360(4).Value = KB.kansan_num4
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		ElseIf IDX = N360_5 Then 
			If IMNU360(5).Value = 0 Then
				If LST_NO > CUR_NO Then
					IMNU360(5).Value = KB.kansan_num5
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
			
		Else
			If Not CHK_DUPCOMBO(iNo, CMB350(iNo).Text) Then
				If LST_NO > CUR_NO Then
					Select Case IDX
						Case N350_1
							Call COMBO_SETLIST(CMB350(1), KB.ha_tanka1)
						Case N350_2
							Call COMBO_SETLIST(CMB350(2), KB.ha_tanka2)
						Case N350_3
							Call COMBO_SETLIST(CMB350(3), KB.ha_tanka3)
						Case N350_4
							Call COMBO_SETLIST(CMB350(4), KB.ha_tanka4)
						Case N350_5
							Call COMBO_SETLIST(CMB350(5), KB.ha_tanka5)
					End Select
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
			
			'   ���Z�P�ʃR���{�{�b�N�X�̃`�F�b�N
			If Trim(CMB350(iNo).Text) = "" Then
				If LST_NO > CUR_NO Then
					Select Case IDX
						Case N350_1
							Call COMBO_SETLIST(CMB350(1), KB.ha_tanka1)
						Case N350_2
							Call COMBO_SETLIST(CMB350(2), KB.ha_tanka2)
						Case N350_3
							Call COMBO_SETLIST(CMB350(3), KB.ha_tanka3)
						Case N350_4
							Call COMBO_SETLIST(CMB350(4), KB.ha_tanka4)
						Case N350_5
							Call COMBO_SETLIST(CMB350(5), KB.ha_tanka5)
					End Select
					Exit Sub
				End If
				
				Select Case IDX
					Case N350_1 '   ���̃P�[�X�ǉ� 2000/02/23
						KB.ha_tanka1 = ""
						IMNU360(1).Value = 0
						KB.kansan_num1 = 0
						CMB350(2).SelectedIndex = -1
						KB.ha_tanka2 = ""
						IMNU360(2).Value = 0
						KB.kansan_num2 = 0
						CMB350(3).SelectedIndex = -1
						KB.ha_tanka3 = ""
						IMNU360(3).Value = 0
						KB.kansan_num3 = 0
						CMB350(4).SelectedIndex = -1
						KB.ha_tanka4 = ""
						IMNU360(4).Value = 0
						KB.kansan_num4 = 0
						CMB350(5).SelectedIndex = -1
						KB.ha_tanka5 = ""
						IMNU360(5).Value = 0
						KB.kansan_num5 = 0
					Case N350_2
						KB.ha_tanka2 = ""
						IMNU360(2).Value = 0
						KB.kansan_num2 = 0
						CMB350(3).SelectedIndex = -1
						KB.ha_tanka3 = ""
						IMNU360(3).Value = 0
						KB.kansan_num3 = 0
						CMB350(4).SelectedIndex = -1
						KB.ha_tanka4 = ""
						IMNU360(4).Value = 0
						KB.kansan_num4 = 0
						CMB350(5).SelectedIndex = -1
						KB.ha_tanka5 = ""
						IMNU360(5).Value = 0
						KB.kansan_num5 = 0
					Case N350_3
						KB.ha_tanka3 = ""
						IMNU360(3).Value = 0
						KB.kansan_num3 = 0
						CMB350(4).SelectedIndex = -1
						KB.ha_tanka4 = ""
						IMNU360(4).Value = 0
						KB.kansan_num4 = 0
						CMB350(5).SelectedIndex = -1
						KB.ha_tanka5 = ""
						IMNU360(5).Value = 0
						KB.kansan_num5 = 0
					Case N350_4
						KB.ha_tanka4 = ""
						IMNU360(4).Value = 0
						KB.kansan_num4 = 0
						CMB350(5).SelectedIndex = -1
						KB.ha_tanka5 = ""
						IMNU360(5).Value = 0
						KB.kansan_num5 = 0
					Case N350_5
						KB.ha_tanka5 = ""
						IMNU360(5).Value = 0
						KB.kansan_num5 = 0
				End Select
				NXT_NO = N410
				Call FOCUS_SET()
			Else ' 350_2�����
				
			End If
			
		End If
		
		Select Case IDX
			Case N350_1
				KB.ha_tanka1 = CMB350(1).Text
			Case N350_2
				KB.ha_tanka2 = CMB350(2).Text
			Case N350_3
				KB.ha_tanka3 = CMB350(3).Text
			Case N350_4
				KB.ha_tanka4 = CMB350(4).Text
			Case N350_5
				KB.ha_tanka5 = CMB350(5).Text
			Case N360_1
				KB.kansan_num1 = IMNU360(1).Value
			Case N360_2
				KB.kansan_num2 = IMNU360(2).Value
			Case N360_3
				KB.kansan_num3 = IMNU360(3).Value
			Case N360_4
				KB.kansan_num4 = IMNU360(4).Value
			Case N360_5
				KB.kansan_num5 = IMNU360(5).Value
		End Select
		
	End Sub
	
	'A-20250201��
	Private Sub IPROCHK_N370()
		
		Select Case CMB370.Text
			Case "�W��"
				KB.tax_rate_kbn = CStr(1)
			Case "�y��"
				KB.tax_rate_kbn = CStr(5)
			Case Else
				KB.tax_rate_kbn = CStr(3)
		End Select
		
		Call SCR_DSPTAX()

		'If OPTO310(3).Value = False And Trim(CMB370.Text) = "" Then 'D-20250417
		If OPTO310(3).Checked = False And Trim(CMB370.Text) = "" Then 'A-20250417
			If CUR_NO <> N330 Then
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If

	End Sub
	'A-20250201��
	
	Private Sub IPROCHK_N410()
		'
		'   �ƎҌ���̋Ǝ�NO���݃`�F�b�N
		
		Dim strReturn As String
		Dim strCode As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX410.Text = KB.g_gentei_code
		'        DSP410.Caption = WKB410DSP
		'        Exit Sub
		'    End If
		
		'    If Trim(IMTX410.Text) = "" Then
		'        ERRSW = F_ERR
		'        Exit Sub
		'    End If
		
		IMTX410.Text = VB6.Format(IMTX410.Text, "000000")
		
		strCode = Trim(IMTX410.Text)
		If strCode <> "" Then
			strReturn = DecodeGYOSHA(WKB010, WKB020, strCode)
			
			If strReturn = "" Then
				If CUR_NO < LST_NO Then
					IMTX410.Text = KB.g_gentei_code
					DSP410.Text = WKB410DSP
					Exit Sub
				End If
				DSP410.Text = ""
				ERRSW = F_ERR
				Exit Sub
			End If
		End If
		
		'   �m��
		KB.g_gentei_code = strCode
		DSP410.Text = strReturn
		WKB410DSP = strReturn
		
		
	End Sub
	
	
	Private Sub IPROCHK_N440()
		
		'    If CUR_NO < LST_NO Then
		'        IMTX440.Text = KB.tax_kubn
		'        Exit Sub
		'    End If
		
		'   ����ŗ��敪
		If IMTX440.Text >= "1" And IMTX440.Text <= "5" Then
			KB.tax_rate_kbn = IMTX440.Text
		Else
			Call SCR_DSPTAX() 'A-20190601
			If CUR_NO < LST_NO Then
				IMTX440.Text = KB.tax_rate_kbn
				Exit Sub
			End If
			ERRSW = F_ERR
			'        ENDSW = F_END
		End If
		
	End Sub
	
	Private Sub IPROCHK_N480N490(ByRef IDX As Short)
		
		Dim iReturn As Short
		Dim strDate As String
		
		'    If CUR_NO < LST_NO Then
		'        If idx = N480 Then
		'            IMTX480.Text = DateSlashed(KB.nouhin_date)
		'        ElseIf idx = N490 Then
		'            IMTX490.Text = DateSlashed(KB.tekiyo_date)
		'        End If
		'        Exit Sub
		'    End If
		
		Select Case IDX
			Case N480
				strDate = IMTX480.Text
			Case N490
				strDate = IMTX490.Text
		End Select
		
		If Trim(strDate) = "" Then
			Select Case IDX
				Case N480 '   �ŏI�[�i��
					KB.nouhin_date = strDate
				Case N490 '   �K�p�J�n���t
					KB.tekiyo_date = strDate
			End Select
			Exit Sub
		End If
		
		iReturn = CHK_DATE(strDate)
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				If IDX = N480 Then
					IMTX480.Text = DateSlashed(KB.nouhin_date)
				ElseIf IDX = N490 Then 
					IMTX490.Text = DateSlashed(KB.tekiyo_date)
				End If
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   �m��
		Select Case IDX
			Case N480 '   �ŏI�[�i��
				KB.nouhin_date = strDate
				IMTX480.Text = DateSlashed(KB.nouhin_date)
			Case N490 '   �K�p�J�n���t
				KB.tekiyo_date = strDate
				IMTX490.Text = DateSlashed(KB.tekiyo_date)
		End Select
		
	End Sub
	
	
	Public Sub IPROCHK_N510()
		
		Dim iReturn As Short
		Dim strDate As String
		
		'    If CUR_NO < LST_NO Then
		'        IMTX510.Text = DateSlashed(KB.tori_kyu_date)
		'        Exit Sub
		'    End If
		
		strDate = IMTX510.Text
		
		If Trim(strDate) = "" Then
			If KB.tori_kyu <> "1" Then
				KB.tori_kyu_date = strDate
				Exit Sub
			Else
				If CUR_NO < LST_NO Then
					IMTX510.Text = DateSlashed(KB.tori_kyu_date)
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		End If
		
		iReturn = CHK_DATE(strDate)
		If iReturn = F_ERR Then
			If CUR_NO < LST_NO Then
				IMTX510.Text = DateSlashed(KB.tori_kyu_date)
				Exit Sub
			End If
			ERRSW = F_ERR
			Exit Sub
		Else
			If KB.tori_kyu <> "1" Then
				If CUR_NO < LST_NO Then
					IMTX510.Text = DateSlashed(KB.tori_kyu_date)
					Exit Sub
				End If
				ERRSW = F_ERR
				Exit Sub
			End If
		End If
		
		'   �m��
		KB.tori_kyu_date = strDate
		IMTX510.Text = DateSlashed(KB.tori_kyu_date)
		
	End Sub
	
	Public Sub IPROCHK_N500()
		
		If CHK500.CheckState = 1 Then
			KB.tori_kyu_date = ""
			IMTX510.Text = ""
		End If
		
	End Sub
	
	Public Function GPROCHK() As Boolean
		
		GPROCHK = True
		
		ERRSW = F_OFF
		
		'   �Ȗڒ��v�f���v�f�̊m��͐����s�Ȃ��B                    2000/01/25  Add
		If CTRLTBL(CUR_NO).IGRP <= CTRLTBL(LST_NO).IGRP Then
			If CTRLTBL(LST_NO).IGRP <> GRP7 Then
				Exit Function
			End If
		End If
		
		Select Case CTRLTBL(LST_NO).IGRP
			Case GRP1
				Call GPROCHK_GRP1()
				If ERRSW = F_ERR Then
					GPROCHK = False
					'                NXT_NO = LST_NO
					'                If NXT_NO <> 0 Then
					'                    Call FOCUS_SET
					'                End If
					Exit Function
				End If
				'   �����敪����i�Ԉȍ~�֍s���Ƃ���
				If CUR_NO >= N030 Then
					Call GPROCHK_GRP2()
					If ERRSW = F_ERR Then
						GPROCHK = False
						NXT_NO = N010
						Call FOCUS_SET()
						Exit Function
					End If
				End If
				
			Case GRP2
				Call GPROCHK_GRP2()
				If ERRSW = F_ERR Then
					GPROCHK = False
					NXT_NO = LST_NO
					If NXT_NO <> 0 Then
						Call FOCUS_SET()
					End If
					Exit Function
				End If
			Case GRP3
				Call GPROCHK_GRP3()
			Case GRP4
				Call GPROCHK_GRP4()
				If ERRSW = F_ERR Then
					GPROCHK = False
					Exit Function
				End If
				
			Case GRP7 '   ��p�Ȗڒ��v�f�A���v�f�̃`�F�b�N
				Call GPROCHK_GRP7()
				ERRSW = F_OFF '   ���ʂ���̓G���[�Ƃ��Ȃ��B
				
			Case GRP8 '   �Ȗڕ��ނ̃`�F�b�N
				Call GPROCHK_GRP8()
				
		End Select
		
		If ERRSW = F_ERR Then
			GPROCHK = False
			NXT_NO = LST_NO
			If NXT_NO <> 0 Then
				Call FOCUS_SET()
			End If
			'        NXT_NO = CTRLTBL(LST_NO).INEXT
			'        If CTRLTBL(NXT_NO).CTRL.TabStop Then
			'        Debug.Print "GPROCHK ERR"; NXT_NO
			'            Call FOCUS_SET
			'        End If
		End If
		
		If ERRSW = F_ERR Then
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = False
		Else
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = True
		End If
		
		
	End Function
	
	'   ��ЃR�[�h�Ǝ��Ə��R�[�h�̑g�ݍ��킹�`�F�b�N
	Public Sub GPROCHK_GRP1()
		
		Dim cdJigyo As String
		Dim strJIGYO As String
		Dim iReturn As Short
		
		Dim strN010 As String
		Dim strN020 As String
		Dim strN010DSP As String
		Dim strN020DSP As String
		
		'   ��ЃR�[�h���݃`�F�b�N
		If CUR_NO > N010 Then
			strN010 = ZeroFill((IMTX010.Text), 2) '   Fix Length?
			iReturn = CduDecodeKaisha(strN010, strN010DSP)
			If iReturn = F_ERR Then
				ERRSW = F_ERR
				NXT_NO = N010
				Debug.Print(VB6.TabLayout("GRP1-1 " & LST_NO, CUR_NO & NXT_NO))
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		
		'   ���Ə��R�[�h���݃`�F�b�N
		If CUR_NO > N020 Then
			strN020 = ZeroFill((IMTX020.Text), 4) '   Fix Length?
			iReturn = CduDecodeJigyo(WKB010, strN020, strN020DSP)
			If iReturn = F_ERR Then
				ERRSW = F_ERR
				NXT_NO = N020
				Debug.Print(VB6.TabLayout("GRP1-2 " & LST_NO, CUR_NO & NXT_NO))
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		
		
		
		
		'    cdJigyo = Right(("0000" & Trim(IMTX020.Text)), 4)
		'    iReturn = CduDecodeJigyo(WKB010, cdJigyo, strJIGYO)
		'    If iReturn <> F_OFF Then
		'        If CUR_NO < LST_NO Then
		'            ENDSW = F_OFF
		'            ERRSW = F_OFF
		'            IMTX020.Text = WKB020
		'            DSP020.Caption = WKB020DSP
		'            Exit Sub
		'        End If
		'
		'        DSP020.Caption = ""
		'        ENDSW = F_END
		'        ERRSW = F_ERR
		'        Exit Sub
		'    End If
		
	End Sub
	
	'   ��ЃR�[�h�Ǝ��Ə��R�[�h�̑g�ݍ��킹�`�F�b�N
	Public Sub GPROCHK_GRP2()
		
		Dim cdJigyo As String
		Dim strJIGYO As String
		Dim iReturn As Short
		
		cdJigyo = VB.Right("0000" & Trim(IMTX020.Text), 4)
		iReturn = CduDecodeJigyo(WKB010, cdJigyo, strJIGYO)
		If iReturn <> F_OFF Then
			If CUR_NO < LST_NO Then
				ENDSW = F_OFF
				ERRSW = F_OFF
				IMTX020.Text = WKB020
				DSP020.Text = WKB020DSP
				Exit Sub
			End If
			
			DSP020.Text = ""
			ENDSW = F_END
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'�Z�L�����e�B�`�F�b�N�i�Q�j���Ə��Q�ƌ���
		Dim Ret As Short
		
		MKKDBCMN.MKKDBCMN_RCN = ZACN_RCN
		Ret = MKKDBCMN.MKKDBCMN_SQTGET2_SUB(3, "SZ0410", IMTX010.Text, IMTX020.Text, WG_OPCODE, W_KENGEN(2))
		If Ret <> n0 Then
			ERRSW = F_ERR
			ENDSW = F_END
			Exit Sub
		ElseIf W_KENGEN(2) = 0 Then 
			ERRSW = F_ERR
			ZAER_KN = n0
			ZAER_CD = 302
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			ERRSW = F_ERR
			ENDSW = F_END
			bSPRNotReady = True
			Exit Sub
		End If
		
		
		'   �f�[�^����\������B
		If CUR_NO > N030 Then
			
			WKB030 = KB.hin_code
			If KBKBN = F_ADD Then
				SetLookupMode(False)
				'   FALSE�ɂ����DEL_FLG=1�̂��̂��ǂ߂�B
				
				''''Call SCR_ADDNEW
				Call SpreadInit()
				Call SCR_DSPDATA()
				Call SCR_BUSHO(False, WKB030)
				
				SetLookupMode(True)
				'   TRUE�ɂ����DEL_FLG=1�̂��͓̂ǂ߂Ȃ��B
				
				Call OptionRefresh()
				'   TAB�ŏ���TAB�ɐݒ�          NR-SZ0410-2
				TAB010.SelectedIndex = 0
				
			Else
				SetLookupMode(False)
				'   FALSE�ɂ����DEL_FLG=1�̂��̂��ǂ߂�B
				Call SpreadInit()
				Call SCR_DSPDATA()
				Call SCR_BUSHO(True, WKB030)
				
				SetLookupMode(True)
				'   TRUE�ɂ����DEL_FLG=1�̂��͓̂ǂ߂Ȃ��B
				
				Call OptionRefresh()
				'   TAB�ŏ���TAB�ɐݒ�          NR-SZ0410-2
				TAB010.SelectedIndex = 0
			End If
		End If
		
	End Sub
	
	'   GRP3
	'           �i��
	Public Sub GPROCHK_GRP3()
		
	End Sub
	
	'   GRP4
	'           �i�����炻�̑��o�[�R�[�h
	Public Sub GPROCHK_GRP4()
		
		'   �i��NOTNULL�`�F�b�N
		If RTrim(IMTX040.Text) = "" Then
			NXT_NO = N040
			Call FOCUS_SET()
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'   �P�ʃ`�F�b�N
		If Trim(CMB060.Text) = "" Then
			NXT_NO = N060
			Call FOCUS_SET()
			ERRSW = F_ERR
			Exit Sub
		End If
		
		
	End Sub
	
	'   GRP5
	'           �K�p���A�����A�_�񉿊i�̂P
	Public Sub GPROCHK_GRP5()
		
		Dim strDate As String
		Dim iReturn As Short
		
		strDate = IMTX100(1).Text
		strDate = Mid(strDate, 1, 4) & Mid(strDate, 6, 2) & Mid(strDate, 9, 2)
		iReturn = CHK_DATE(strDate)
		If iReturn = F_ERR Then
			NXT_NO = N100_1
			Call FOCUS_SET()
			ERRSW = F_ERR
			Exit Sub
		End If
		
		
	End Sub
	
	Public Sub GPROCHK_GRP7()
		'
		'   IMTX130, IMTX140 �����̃`�F�b�N
		'   �ȖڑΉ��e�[�u���Ƃ̓ˍ���
		
		Dim iReturn As Short
		Dim KamUri As String
		Dim KamSho As String
		Dim KamMat As String
		Dim KamMit As String
		Dim strAcctName As String
		
		iReturn = TaiouKamoku(WKB010, WKB020, IMTX130(1).Text, IMTX140(1).Text, KamUri, KamSho, KamMat, KamMit)
		If iReturn <> F_OFF Then
			'        IMTX130(1).Text = KB.hiyou_k_code1
			'        IMTX140(1).Text = KB.hiyou_k_code2
			'        ERRSW = F_ERR
			'        ENDSW = F_END
			'        NXT_NO = N130       '   ��p�Ȗڒ��v�f�ɂ��ǂ��B
			'        Exit Sub
		End If
		KB.hiyou_k_code1 = IMTX130(1).Text
		KB.hiyou_k_code2 = IMTX140(1).Text
		
		Call AccountName(KamUri, strAcctName)
		IMTX130(2).Text = Mid(KamUri, 1, 3)
		IMTX140(2).Text = Mid(KamUri, 4, 6)
		DSP140(2).Text = strAcctName
		Call AccountName(KamSho, strAcctName)
		IMTX130(3).Text = Mid(KamSho, 1, 3)
		IMTX140(3).Text = Mid(KamSho, 4, 6)
		DSP140(3).Text = strAcctName
		Call AccountName(KamMat, strAcctName)
		IMTX130(4).Text = Mid(KamMat, 1, 3)
		IMTX140(4).Text = Mid(KamMat, 4, 6)
		DSP140(4).Text = strAcctName
		Call AccountName(KamMit, strAcctName)
		IMTX130(5).Text = Mid(KamMit, 1, 3)
		IMTX140(5).Text = Mid(KamMit, 4, 6)
		DSP140(5).Text = strAcctName
		
		
	End Sub
	
	Private Sub unusedAccountName(ByRef cdKAM As String, ByRef nmKAM As String)
		
		Dim nmCHU As String
		Dim nmSHO As String
		
		nmCHU = DecodeKAMOCHU(Mid(cdKAM, 1, 3))
		If nmCHU <> "" Then
			nmSHO = DecodeKAMOKU(Mid(cdKAM, 1, 3), Mid(cdKAM, 4, 6))
			''''nmKAM = DecodeKAMOKU(Mid(cdKAM, 1, 3), Mid(cdKAM, 4, 6))
			
			nmKAM = nmCHU & nmSHO
		Else
			nmKAM = nmCHU
		End If
		
	End Sub
	
	
	Public Sub GPROCHK_GRP8()
		'
		'   IMTX210, IMTX211 �����̃`�F�b�N
		
		'    Dim strName As String
		'    Dim strCode As String
		
		'    strCode = IMTX210.Text & IMTX211.Text
		'    strName = DecodeKamBunrui(WKB010, WKB020, strCode)
		'    If strName = "" Then
		'        IMTX210.Text = Left(KB.ka_bun_code, 3)
		'        IMTX211.Text = Right(KB.ka_bun_code, 4)
		'        DSP210.Caption = ""
		'        ENDSW = F_END
		'        ERRSW = F_ERR
		'    Else
		'        DSP210.Caption = strName
		'        KB.ka_bun_code = strCode
		'    End If
		
	End Sub
	Public Function GVALCHK() As Boolean
		
		GVALCHK = True
		
		'A-CUST-20100610 Start
		Dim nnum As String
		
		If KBKBN <> F_ADD Then Exit Function
		If LST_NO <> n0 Then
			If CTRLTBL(CUR_NO).IGRP = CTRLTBL(LST_NO).IGRP Then Exit Function
		End If
		If CTRLTBL(CUR_NO).IGRP > GRP3 Then
			If CTRLTBL(LST_NO).IGRP < GRP3 Or CTRLTBL(LST_NO).IGRP = GEND Then
				nnum = CStr(New_Number)
				If CDbl(nnum) < 0 Or CDbl(nnum) > 99999 Then
					Call MsgBox("�����̔Ԃ�����ɒB���܂����B�@", MsgBoxStyle.ApplicationModal + MsgBoxStyle.Exclamation, "�d���i�ڊ�{������")
					IMTX030.Text = ""
					NXT_NO = LST_NO
					GVALCHK = False
					Call FOCUS_SET()
				Else
					WKB030 = VB6.Format(nnum, "00000")
					IMTX030.Text = WKB030
					KB.hin_code = WKB030
					LST_NO = N030
				End If
			End If
		End If
		'A-CUST-20100610 End
	End Function
	
	Public Function MVALCHK() As Boolean
		'
		MVALCHK = True
		
		Select Case CUR_NO
			Case N999
				Call MVALCHK_N999()
				If ERRSW = F_ERR Then
					MVALCHK = False
					Exit Function
				End If
				'A-CUST-20100610 Start
			Case N030
				If KBKBN = F_ADD Then
					NXT_NO = LST_NO
					Call FOCUS_SET()
					MVALCHK = False
					Exit Function
				End If
				'A-CUST-20100610 End
			Case N350_1 To N360_5
				Call MVALCHK_N350N360()
				If ERRSW = F_ERR Then
					If LST_NO > CUR_NO Then
						NXT_NO = N350_1
						Call FOCUS_SET()
					ElseIf LST_NO <> CUR_NO - 1 Then 
						NXT_NO = LST_NO
						Call FOCUS_SET()
					End If
					MVALCHK = False
					Exit Function
				ElseIf ERRSW = F_END Then 
					MVALCHK = False
					Exit Function
				End If
				
		End Select
		
		If LST_NO = N999 Then
			If CUR_NO <= N030 Then
				
			Else
				ERRSW = F_ERR
				ENDSW = F_END
				NXT_NO = LST_NO
				Call FOCUS_SET()
				MVALCHK = False
				Exit Function
			End If
		End If
		
		
		
		'   �폜�̂Ƃ��� IMTX030�݂̂ɂ����t�H�[�J�X�Z�b�g�ł��Ȃ��B
		If KBKBN = 3 Then
			If CUR_NO <> N999 And CUR_NO <> N030 And CUR_NO <> N010 And CUR_NO <> N020 And CUR_NO <> NF12 Then
				ERRSW = F_ERR
				ENDSW = F_END
				NXT_NO = LST_NO
				Call FOCUS_SET()
				MVALCHK = False
				Exit Function
			End If
		End If
		
		'   �����x�~����Ȃ��Ƃ���IMTX510�ɂ̓t�H�[�J�X�Z�b�g�ł��Ȃ��B
		If KB.tori_kyu <> "1" Then
			If CUR_NO = N510 And LST_NO <> N500 Then
				If LST_NO < CUR_NO Then
					ERRSW = F_ERR
					NXT_NO = LST_NO
					Call FOCUS_SET()
					MVALCHK = False
					Exit Function
				Else
					NXT_NO = N490
					Call FOCUS_SET()
					MVALCHK = False
					Exit Function
				End If
			End If
		End If
		
		
		'   TAB�̈ړ����R���g���[��
		Dim PrevTab As Short
		
		MVALCHK = True
		
		'    If CUR_NO = N999 And Trim(IMTX030.Text) <> "" Then
		'        If MsgBox("�N���A���܂����H", _
		''        vbYesNo + vbApplicationModal + vbQuestion, _
		''        "�d���i�ڊ�{������") = vbNo Then
		'            ERRSW = F_ERR
		'            ENDSW = F_END
		'            MVALCHK = False
		'            Exit Function
		'        Else
		'    '   �N���A����B
		'            WKB030 = ""
		'            Call SCR_ADDNEW
		'            Call SpreadInit
		'            Call SCR_DSPDATA
		'            Call SCR_BUSHO(False, WKB030)
		'            Call OptionRefresh
		'        End If
		'    End If
		
		If KBKBN = 3 Then Call SetMode("D")
		
		Select Case CUR_NO
			Case N010
				IMTX010.Text = ZeroTrim((IMTX010.Text))
			Case N020
				IMTX020.Text = ZeroTrim((IMTX020.Text))
			Case N030
				IMTX030.Text = ZeroTrim((IMTX030.Text))
			Case N130
				IMTX130(1).Text = ZeroTrim(IMTX130(1).Text)
			Case N140
				IMTX140(1).Text = ZeroTrim9(IMTX140(1).Text)
				'A-CUST20130212��
			Case N140
				IMTX150(0).Text = ZeroTrim9(IMTX150(0).Text)
				'A-CUST20130212��
				'D-20250201��
				'Case N210
				'IMTX210.Text = ZeroTrim(IMTX210.Text)
				'Case N211
				'D-20250201��
				''''IMTX211.Text = ZeroTrim(IMTX211.Text)
			Case N220
				''''IMTX220.Text = ZeroTrim(IMTX220.Text)
			Case N230
				''''IMTX230.Text = ZeroTrim(IMTX230.Text)
			Case N240
				''''IMTX240.Text = ZeroTrim(IMTX240.Text)
				'Case N250  'D-20250201
				''''IMTX250.Text = ZeroTrim(IMTX250.Text)
			Case N260
				IMTX260.Text = ZeroTrim((IMTX260.Text))
			Case N410
				IMTX410.Text = ZeroTrim((IMTX410.Text))
				
		End Select
		
		Select Case CUR_NO
			'        Case N100_1 To N140'D-CUST20130212
			'Case N100_1 To N170 'A-CUST20130212    D-20240115
			Case N100_1 To N175 'A-20240115
				If TAB010.SelectedIndex <> 0 Then
					PrevTab = TAB010.SelectedIndex
					TAB010.SelectedIndex = 0
					''''Debug.Print "TAB Moved from " & PrevTab & " to "; TAB010.Tab
				End If
				
				'Case N210 To N360_5    'D-20250201
			Case N220 To N370 'A-20250201
				If TAB010.SelectedIndex <> 1 Then
					PrevTab = TAB010.SelectedIndex
					TAB010.SelectedIndex = 1
					''''Debug.Print "TAB Moved from " & PrevTab & " to "; TAB010.Tab
				End If
				
			Case N410 To N510
				If TAB010.SelectedIndex <> 2 Then
					PrevTab = TAB010.SelectedIndex
					TAB010.SelectedIndex = 2
					''''Debug.Print "TAB Moved from " & PrevTab & " to "; TAB010.Tab
				End If
				
		End Select
		
	End Function
	
	Private Sub MVALCHK_N350N360()
		'   �����P�ʂ͏ォ�珇�Ԃɋl�܂��Ă��Ȃ���΂Ȃ�Ȃ��B
		
		Select Case CUR_NO
			Case N350_2
				If Trim(CStr(KB.ha_tanka1 = "")) Or Trim(CMB350(1).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N350_3
				If Trim(CStr(KB.ha_tanka2 = "")) Or Trim(CMB350(2).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N350_4
				If Trim(CStr(KB.ha_tanka3 = "")) Or Trim(CMB350(3).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N350_5
				If Trim(CStr(KB.ha_tanka5 = "")) Or Trim(CMB350(4).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N360_1
				If Trim(CStr(KB.ha_tanka1 = "")) Or Trim(CMB350(1).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N360_2
				If Trim(CStr(KB.ha_tanka2 = "")) Or Trim(CMB350(2).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N360_3
				If Trim(CStr(KB.ha_tanka3 = "")) Or Trim(CMB350(3).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N360_4
				If Trim(CStr(KB.ha_tanka4 = "")) Or Trim(CMB350(4).Text) = "" Then
					ERRSW = F_ERR
				End If
			Case N360_5
				If Trim(CStr(KB.ha_tanka5 = "")) Or Trim(CMB350(5).Text) = "" Then
					ERRSW = F_ERR
				End If
				
		End Select
		If ERRSW = F_ERR Then Exit Sub
		
	End Sub
	
	Private Sub MVALCHK_N999()
		
		Dim eCUR_NO As Short
		If Trim(IMTX030.Text) <> "" Then
			If MsgBox("�N���A���܂����H", MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question, "�d���i�ڊ�{������") = MsgBoxResult.No Then
				ERRSW = F_ERR
				ENDSW = F_END
				'OPTO999(KBKBN).Value = True 'D-20250417
				OPTO999(KBKBN).Checked = True 'A-20250417
				CTRLTBL(N999).CTRL = OPTO999(KBKBN)
				NXT_NO = LST_NO
				Call FOCUS_SET()
				Exit Sub
			Else
				'   �N���A����B
				'A-20250201��
				eCUR_NO = CUR_NO
				'A-20250201��
				Call DBRollbackTrans()
				Call DBBeginTrans()
				WKB030 = ""
				Call SCR_ADDNEW()
				Call SpreadInit()
				Call SCR_DSPDATA()
				Call SCR_BUSHO(False, WKB030)
				Call OptionRefresh()
				CUR_NO = eCUR_NO 'A-20250201
			End If
		End If
		
		
	End Sub
	
	Public Sub FUNCSET_RTN()
		
		Dim i As Short
		
		For i = 0 To 12
			ZAFC_N(i) = CShort("00")
		Next i
		
		'--- �t�@���N�V�����E�K�C�h���b�Z�[�W
		'Debug.Assert LST_NO = CUR_NO
		'Debug.Assert CUR_NO <> NF12
		Debug.Print("FUNCSET" & LST_NO)
		Select Case LST_NO
			Case N999
				ZAFC_N(0) = CShort("01")
			Case N010, N020
				ZAFC_N(0) = CShort("01")
				ZAFC_N(3) = CShort("03")
				''''        ZAFC_N(12) = "12"
			Case N030
				ZAFC_N(0) = CShort("01")
				If KBKBN <> 1 Then
					ZAFC_N(3) = CShort("03")
				End If
				If Trim(IMTX030.Text) <> "" Then
					ZAFC_N(5) = CShort("05")
				End If
				'            If KBKBN = 3 Then
				'            ZAFC_N(12) = "12"
				'            End If
				
			Case N130, N140
				ZAFC_N(3) = CShort("03")
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'A-CUST20130212��
				'Case N150, N160, N170, N170CMB         'D-20240115
			Case N150, N160, N165, N170, N170CMB 'A-20240115
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'A-CUST20130212��
				'A-CUST20130212��
			Case N070
				ZAFC_N(3) = CShort("03")
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'A-CUST20130212��
			Case N040 To N090, N100_1 To N120_2
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'Case N210 To N240, N260        'D-CUST20130212
				'Case N210 To N240, N260, N291   'A-CUST20130212    'D-20250201
			Case N220 To N240, N260, N291 'A-20250201
				ZAFC_N(3) = CShort("03")
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'D-20250201��
				'Case N250               '   ���ނ͗\���R�[�h�Ŗ��g�p
				'ZAFC_N(3) = "03"        '02/05/28 ADD
				'ZAFC_N(5) = "05"
				'ZAFC_N(12) = "12"
				'D-20250201��
			Case N270 To N340, N350_1 To N360_5
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
			Case N410
				ZAFC_N(3) = CShort("03")
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
			Case N420
				ZAFC_N(3) = CShort("03")
				ZAFC_N(5) = CShort("05")
				If SPR420.DataRowCnt > 0 Then
					ZAFC_N(8) = CShort("08")
				End If
				
				ZAFC_N(12) = CShort("12")
				
			Case N430 To N510
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				
			Case NF12
				ZAFC_N(12) = 12
				
				'A-20250201��
			Case N370
				ZAFC_N(5) = CShort("05")
				ZAFC_N(12) = CShort("12")
				'A-20250201��
				
		End Select
		'A-CUST-20100610 Start
		Select Case LST_NO
			Case N030
				ZAFC_N(6) = CShort("13")
			Case N040 To N510
				ZAFC_N(6) = CShort("13")
				ZAFC_N(7) = CShort("14")
		End Select
		'A-CUST-20100610 End
		'    If KBKBN = F_ADD Then ZAFC_N(4) = "04"
		If KBKBN = 1 Or KBKBN = 2 Then
			If Trim(WKB030) <> "" Then
				If CUR_NO <> N010 And CUR_NO <> N020 And CUR_NO <> N999 And CUR_NO <> NF12 Then
					ZAFC_N(4) = CShort("04")
				End If
			End If
		End If
		
		
		Select Case LST_NO
			Case N999
				ZAGD_NO.Value = "045"
				'Case N010 To N090                              'D-CUST-20100610
			Case N010 To N060 'A-CUST-20100610
				ZAGD_NO.Value = VB.Right("000" & LST_NO + 2, 3)
				'A-CUST-20100610 Start
			Case N065
				ZAGD_NO.Value = "047"
			Case N070 To N090
				ZAGD_NO.Value = VB.Right("000" & LST_NO + 1, 3)
				'A-CUST-20100610 End
			Case N100_1, N100_2
				ZAGD_NO.Value = "013"
			Case N110_1, N110_2
				ZAGD_NO.Value = "014"
			Case N120_1, N120_2
				ZAGD_NO.Value = "015"
			Case N130, N140
				'ZAGD_NO = Right(("000" & LST_NO - 1), 3)           'D-CUST-20100901
				ZAGD_NO.Value = VB.Right("000" & LST_NO - 2, 3) 'A-CUST-20100901
				'A-CUST20130212��
			Case N150
				ZAGD_NO.Value = "49"
			Case N160
				ZAGD_NO.Value = "50"
				'A-20240115��
			Case N165
				ZAGD_NO.Value = "53"
				'A-20240115��
			Case N170, N170CMB
				ZAGD_NO.Value = "51"
				'A-CUST20130212��
				'D-20250201��
				'Case N210, N211
				'ZAGD_NO = "018"
				'D-20250201��
				'A-20250201��
			Case N230
				ZAGD_NO.Value = "54"
				'A-20250201��
			Case N220 To N340
				'ZAGD_NO = Right(("000" & LST_NO - 2), 3)           'D-CUST-20100901
				'            ZAGD_NO = Right(("000" & LST_NO - 3), 3)            'A-CUST-20100901 'D-CUST20130212
				'A-CUST20130212��
				If LST_NO > N291 Then
					'ZAGD_NO = Right(("000" & LST_NO - 8), 3)       'D-20240115
					'ZAGD_NO = Right(("000" & LST_NO - 10), 3)       'A-20240115    'D-20250201
					ZAGD_NO.Value = VB.Right("000" & LST_NO - 7, 3) 'A-20250201
					'ElseIf LST_NO < N291 Then  'D-20250201
				ElseIf LST_NO < N260 Then  'A-20250201
					'ZAGD_NO = Right(("000" & LST_NO - 7), 3)       'D-20240115
					'ZAGD_NO = Right(("000" & LST_NO - 9), 3)        'A-20240115    'D-20250201
					ZAGD_NO.Value = VB.Right("000" & LST_NO - 7, 3) 'A-20250201
					'A-20250201��
				ElseIf LST_NO < N291 Then 
					ZAGD_NO.Value = VB.Right("000" & LST_NO - 6, 3)
					'A-20250201��
				ElseIf LST_NO = N291 Then 
					ZAGD_NO.Value = "52"
				End If
				'A-CUST20130212��
			Case N350_1, N350_2, N350_3, N350_4, N350_5
				ZAGD_NO.Value = "032"
			Case N360_1, N360_2, N360_3, N360_4, N360_5
				ZAGD_NO.Value = "033"
			Case N410
				ZAGD_NO.Value = "034"
			Case N420
				ZAGD_NO.Value = "035"
				'Case N430 To N510  'D-20250201
				'A-20250201��
			Case N430
				ZAGD_NO.Value = "036"
			Case N450 To N510
				'A-20250201��
				'ZAGD_NO = Right(("000" & LST_NO - 10), 3)          'D-CUST-20100901
				'            ZAGD_NO = Right(("000" & LST_NO - 11), 3)           'A-CUST-20100901 'D-CSUT20130212
				'ZAGD_NO = Right(("000" & LST_NO - 16), 3)           'A-CUST20130212  'D-20240115
				'ZAGD_NO = Right(("000" & LST_NO - 18), 3)            'A-20240115   'D-20250201
				ZAGD_NO.Value = VB.Right("000" & LST_NO - 16, 3) 'A-20250201
				'''''''''Debug.Print LST_NO; ZAGD_NO
			Case NF12
				ZAGD_NO.Value = "046"
				'A-20250201��
			Case N370
				ZAGD_NO.Value = "028"
				'A-20250201��
		End Select
		
		'�t�@���N�V�������b�Z�[�W
		Call ZAFC_SUB(Me)
		''''Debug.Assert LST_NO < 20
		'�K�C�h���b�Z�[�W
		Call ZAGD_SUB(Me)
		
		
		
	End Sub
	
	
	
	
	'   ���̍��ڂ��Z�b�g
	Public Sub SET_NO(ByRef FUNC As Short)
		
		Dim i As Short
		
		If KBKBN = 3 And LST_NO = N030 Then
			If FUNC = 1 Or FUNC = 3 Then
				CMDOFNC(12).Enabled = True
				NXT_NO = NF12
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		'   LST_NO = ZERO �̂Ƃ��͍ŏ��̔ԍ�
		If LST_NO = n0 Then
			
			NXT_NO = N999
			Exit Sub
		End If
		
		i = LST_NO
		Do 
			Select Case FUNC
				Case 1 ' ������
					NXT_NO = GetNxtNo(CTRLTBL(i).INEXT, 1)
				Case 2 ' �O����
					NXT_NO = GetNxtNo(CTRLTBL(i).IBACK, 2)
				Case 3 ' ���O���[�v
					NXT_NO = GetNxtNo(CTRLTBL(i).IDOWN, 3)
			End Select
			
			If NXT_NO = n0 Or NXT_NO = NEND Then Exit Sub
			
			'        If NXT_NO = N010 Then IMTX010.TabStop = True
			
			If CTRLTBL(NXT_NO).CTRL.TabStop = True Then
				System.Diagnostics.Debug.Write("SET_NO from " & LST_NO & " to " & NXT_NO)
				Debug.Print("CTRL=" & CTRLTBL(NXT_NO).CTRL.Name)
				Call FOCUS_SET()
				Exit Sub
			Else
				System.Diagnostics.Debug.Write("NOT TABSTOP" & NXT_NO & " ")
				Debug.Print("CTRL=" & CTRLTBL(NXT_NO).CTRL.Name)
				i = NXT_NO
				
			End If
		Loop 
		
	End Sub
	
	Private Function GetNxtNo(ByVal NxtNo As Short, ByVal kbn As Short) As Short
		'   �����ڂ����͕s�̏ꍇ�ɁA���̎��ɓ��͉\�ȍ��ڔԍ���Ԃ��B
		'   (���ڑI���Ɋ֘A���鏈��)
		'   NxtNo(�����ڔԍ�)�A
		'   Kbn(1:�����ڤ2:�O���ڤ3:���O���[�v�Else:�l�����̂܂ܕԂ�)
		'       ��NXT_NO��LST_NO��Ă���ꍇ�ALST_NO�͓��͉\�Ȃ͂��Ȃ̂�
		'       NxtNo��0��Ă��Ă����΂悢�Ǝv����
		
		If NxtNo = NEND Or NxtNo = n0 Then
			GetNxtNo = NxtNo
			Exit Function
		End If
		
		Do While CTRLTBL(NxtNo).CTRL.Enabled = False
			Select Case kbn
				Case 1 '������
					NxtNo = CTRLTBL(NxtNo).INEXT
				Case 2 '�O����
					NxtNo = CTRLTBL(NxtNo).IBACK
				Case 3 '���O���[�v
					NxtNo = CTRLTBL(NxtNo).IDOWN
				Case Else
					Exit Do
			End Select
			If NxtNo = NEND Or NxtNo = n0 Then Exit Do
		Loop 
		
		GetNxtNo = NxtNo
		
	End Function
	
	Private Sub FOCUS_SET()
		
		'    Debug.Print "FOCUS_SET to:"; NXT_NO
		
		Select Case NXT_NO
			Case N999
				OPTO999(KBKBN).Focus() '�����I��
				
			Case N300
				OPTO300(WKB300).Focus()
				Debug.Print("FoCUS_SET OPTO300(" & WKB300 & ")")
			Case N310
				If WKB310 > 0 Then
					OPTO310(WKB310).Focus()
					Debug.Print("FoCUS_SET OPTO310(" & WKB310 & ")")
				End If
			Case N320
				OPTO320(WKB320).Focus()
				Debug.Print("FoCUS_SET OPTO320(" & WKB320 & ")")
			Case N330
				OPTO330(WKB330).Focus()
				Debug.Print("FoCUS_SET OPTO330(" & WKB330 & ")")
			Case N340
				OPTO340(WKB340).Focus()
				Debug.Print("FoCUS_SET OPTO340(" & WKB340 & ";)")
				
			Case Else
				CTRLTBL(NXT_NO).CTRL.Focus()
				'    Debug.Print "FOCUS_SET:"; CTRLTBL(NXT_NO).CTRL.NAME
				
		End Select
		
	End Sub
	
	Private Sub F8DELETE()
		
		'    Call SPR420_KeyDown(vbKeyF8, 0)
		'    SendKeys "{TAB}"
		Call SpreadDelete()
		
		Call SpreadZeroTrim(-1)
		Call FUNCSET_RTN()
		
		
	End Sub
	
	
	Private Sub F3QUERY(ByRef curno As Short)
		
		Dim iRet As Short
		Dim saveCUR_NO As Short
		
		
		saveCUR_NO = curno
		
		'   �R�[�h���ڂ̖⍇����ʌĂяo���B
		Select Case curno
			Case N010
				bBackForm = True
				Call QUE_KAISHA()
			Case N020
				bBackForm = True
				Call QUE_JIGYO()
				
			Case N030
				bBackForm = True
				iRet = QUE_HINBAN
				If iRet = F_OFF Then
					
					'                IMTX030.Text = WKB030
					Call SET_NO(1)
				Else
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
					
				End If
				'           DoEvents
				
				'A-CUST20130212��
			Case N070 '   �i�`�m�}�X�^����
				bBackForm = True
				SZ0414_TOP = VB6.PixelsToTwipsY(Me.Top)
				SZ0414_LEFT = VB6.PixelsToTwipsX(Me.Left)
				SZ0414_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
				SZ0414_WIDTH = VB6.PixelsToTwipsX(Me.Width)
				SZ0414_POS = 0
				iRet = SZ0414_SUB()
				If iRet = 0 Then
					IMTX070.Text = SZ0414_SELCOD1.Value
					Call SET_NO(1)
				Else
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				'A-CUST20130212��
				
			Case N130 '   ��p�Ȗڒ��v�f
				bBackForm = True
				iRet = QUE_KAMOKU
				If iRet = 0 Then
					'                DoEvents
					
					'               IMTX130(1).Text = KB.hiyou_k_code1
					''''CUR_NO = saveCUR_NO
					Call SET_NO(1)
				Else
					System.Windows.Forms.Application.DoEvents()
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				
				
			Case N140 '   ��p�Ȗڏ��v�f
				bBackForm = True
				iRet = QUE_KAMOKU
				If iRet = 0 Then
					
					'                IMTX140(1).Text = KB.hiyou_k_code2
					Call SET_NO(1)
				Else
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				'            DoEvents
				
				'D-20250201��
				'Case N210, N211     '   �Ȗڕ���
				'bBackForm = True
				'iRet = QUE_KAMBUN()
				'If iRet = 0 Then
				'IMTX210.Text = Mid(KB.ka_bun_code, 1, 3)
				'IMTX211.Text = Mid(KB.ka_bun_code, 4, 4)
				'IMTX210.Text = Mid(SEL_FIND, 1, 3)
				'IMTX211.Text = Mid(SEL_FIND, 4, 4)
				'Call SET_NO(1)
				
				'Else
				'NXT_NO = saveCUR_NO
				'Call FOCUS_SET
				
				'End If
				'D-20250201��
				
			Case N220 To N240 '   �啪�ށA�����ށA������
				bBackForm = True
				iRet = QUE_BUNRUI(curno)
				If iRet = 0 Then
					
				Else
					'           DoEvents
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				'02/05/28 ADD START
				'D-20250201��
				'Case N250
				'bBackForm = True
				'iRet = QUE_BUNRUI(curno)
				'If iRet = 0 Then
				
				'Else
				'NXT_NO = saveCUR_NO
				'Call FOCUS_SET
				'End If
				'D-20250201��
				'02/05/28 ADD END
			Case N260 '   ��������
				bBackForm = True
				iRet = QUE_FIND()
				If iRet = 0 Then
					IMTX260.Text = KB.ken_bun_code
					Call SET_NO(1)
					'                DSP260.Caption = DecodeFIND(KB.ken_bun_code)
					'                WKB260DSP = DSP260.Caption
				Else
					''''DoEvents
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				
				'A-CUST20130212��
			Case N291 '   �i�`�m���i���ރ}�X�^����
				bBackForm = True
				SZ0415_TOP = VB6.PixelsToTwipsY(Me.Top)
				SZ0415_LEFT = VB6.PixelsToTwipsX(Me.Left)
				SZ0415_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
				SZ0415_WIDTH = VB6.PixelsToTwipsX(Me.Width)
				SZ0415_POS = 0
				iRet = SZ0415_SUB()
				If iRet = 0 Then
					IMTX291.Text = SZ0415_SEL_CODE
					Call SET_NO(1)
				Else
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				'A-CUST20130212��
				
			Case N410 '   �ƎҌ���
				bBackForm = True
				iRet = QUE_GYOSHA
				If iRet = 0 Then
					IMTX410.Text = KB.g_gentei_code
					WKB410DSP = DecodeGYOSHA(WKB010, WKB020, KB.g_gentei_code)
					
					DSP410.Text = WKB410DSP
					Call SET_NO(1)
				Else
					''''DoEvents
					NXT_NO = saveCUR_NO
					Call FOCUS_SET()
				End If
				
			Case N420 '   ��������
				bBackForm = True
				iRet = QUE_BUSHO()
				System.Windows.Forms.Application.DoEvents()
				NXT_NO = saveCUR_NO
				Call FOCUS_SET()
				
				
			Case N060
				'Call SCR_BNI001_RTN
				CMB060.Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N070
			Case N350_1
				'Call SCR_BNI001_RTN
				CMB350(1).Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N360_1
			Case N350_2
				'Call SCR_BNI001_RTN
				CMB350(2).Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N360_2
			Case N350_3
				'Call SCR_BNI001_RTN
				CMB350(3).Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N360_3
			Case N350_4
				'Call SCR_BNI001_RTN
				CMB350(4).Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N360_4
			Case N350_5
				'Call SCR_BNI001_RTN
				CMB350(5).Focus()
				System.Windows.Forms.Application.DoEvents()
				System.Windows.Forms.SendKeys.Send("{F4}")
				NXT_NO = N360_5
		End Select
		
	End Sub
	
	Public Function QUE_FIND() As Short
		
		Dim Ret As Short
		
		SZ0720.SZ0720_TOP = VB6.PixelsToTwipsY(Me.Top)
		SZ0720.SZ0720_LEFT = VB6.PixelsToTwipsX(Me.Left)
		SZ0720.SZ0720_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
		SZ0720.SZ0720_WIDTH = VB6.PixelsToTwipsX(Me.Width)
		SZ0720.SZ0720_POS = 0
		SZ0720.SZ0720_RCN = ZACN_RCN
		SZ0720.SZ0720_TIME = 0
		SZ0720.SZ0720_INC_CODE = WKB010
		SZ0720.SZ0720_INC_NAME = DSP010.Text
		SZ0720.SZ0720_JG_CODE = WKB020
		SZ0720.SZ0720_JG_NAME = DSP020.Text
		Ret = SZ0720.SZ0720_SUB
		SEL_FIND = SZ0720.SZ0720_SEL_CODE
		If Ret = 0 Then
			KB.ken_bun_code = SEL_FIND
			IMTX260.Text = KB.ken_bun_code
			QUE_FIND = 0
		Else
			QUE_FIND = -1
		End If
		
		
		'    SZ0410FFRM.Show vbModal
		'
		'    If SEL_FIND <> "" Then
		'        KB.ken_bun_code = SEL_FIND
		''        IMTX260.Text = KB.ken_bun_code
		'        QUE_FIND = 0
		'    Else
		'        QUE_FIND = -1
		'    End If
		
	End Function
	
	''''''''''
	'   �Ȗڕ��ނ̖⍇��
	Public Function QUE_KAMBUN() As Short
		
		Dim Ret As Short
		
		SZ0730.SZ0730_TOP = VB6.PixelsToTwipsY(Me.Top)
		SZ0730.SZ0730_LEFT = VB6.PixelsToTwipsX(Me.Left)
		SZ0730.SZ0730_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
		SZ0730.SZ0730_WIDTH = VB6.PixelsToTwipsX(Me.Width)
		SZ0730.SZ0730_POS = 0
		SZ0730.SZ0730_RCN = ZACN_RCN
		SZ0730.SZ0730_TIME = 0
		SZ0730.SZ0730_INC_CODE = WKB010
		SZ0730.SZ0730_INC_NAME = DSP010.Text
		SZ0730.SZ0730_JG_CODE = WKB020
		SZ0730.SZ0730_JG_NAME = DSP020.Text
		Ret = SZ0730.SZ0730_SUB
		
		SEL_FIND = SZ0730.SZ0730_SEL_CODE1 & SZ0730.SZ0730_SEL_CODE2
		''''KB.ka_bun_code = SEL_FIND
		
		
		If Ret = 0 Then
			QUE_KAMBUN = 0
		Else
			QUE_KAMBUN = -1
		End If
		
		'    SEL_TYPE = "KAMOKUBUNRUI"
		'    SZ0410GFRM.Show vbModal
		'
		'    If SEL_FIND <> "" Then
		'        KB.ka_bun_code = SEL_FIND
		'        QUE_KAMBUN = 0
		'    Else
		'        QUE_KAMBUN = -1
		'    End If
		
	End Function
	
	
	
	
	Private Sub F4COPY()
		
		'   �i������DLL���Ăяo���B
		
		Dim RF As SZM0010_S
		Dim iReturn As Short
		Dim strCopyFrom As String
		Dim strCopyKAISHA As String
		Dim strCopyJIGYO As String
		Dim lRet As Integer
		
		
		bBackForm = True
		
		
		SZ0420.SZ0420_KAISYA = WKB010 '  ��к���
		SZ0420.SZ0420_JGCODE = WKB020 '  ���Ə�����
		SZ0420.SZ0420_BSCODE = "" '  ��������
		SZ0420.SZ0420_CHECK = 0 '  �����׸� �i1.�����L�� �P�ȊO���������j
		SZ0420.SZ0420_TOP = VB6.PixelsToTwipsY(Me.Top) '  �e���(TOP)
		SZ0420.SZ0420_LEFT = VB6.PixelsToTwipsX(Me.Left) '  �e���(LEFT)
		SZ0420.SZ0420_HEIGHT = VB6.PixelsToTwipsY(Me.Height) '  �e���(HEIGHT)
		SZ0420.SZ0420_WIDTH = VB6.PixelsToTwipsX(Me.Width) '  �e���(WIDTH)
		SZ0420.SZ0420_POS = 1 '�@�\���ʒu(0.���� 1.���� 2.�E�� 3.���� 4.�E�� )
		SZ0420.SZ0420_RCN = ZACN_RCN
		SZ0420.SZ0420_TIME = CInt(WG_TIMEOUT) '  RDO��ѱ�ĕb��
		
		lRet = SZ0420.SZ0420_SUB
		
		If lRet = 0 Then
			
			strCopyFrom = SZ0420.SZ0420_LCODE
			strCopyKAISHA = SZ0420.SZ0420_KAISYA
			strCopyJIGYO = SZ0420.SZ0420_JGCODE
			
			iReturn = FILGET_SZM0010(strCopyKAISHA, strCopyJIGYO, strCopyFrom, RF)
			If iReturn = F_OFF Then
				WKB010 = strCopyKAISHA
				WKB020 = strCopyJIGYO
				Call COPYFROM(KB, RF)
				Call SpreadInit()
				Call SCR_DSPDATA()
				If KBKBN = 3 Then Call SetMode("D")
				Call SCR_BUSHO(True, strCopyFrom)
				OptionRefresh()
				SentakuFLG = False 'A-CUST-20100610
				
			End If
		Else
			Call SpreadZeroTrim(-1)
			
		End If
		
		
		
	End Sub
	
	Private Sub COPYFROM(ByRef dst As SZM0010_S, ByRef src As SZM0010_S)
		
		Dim saveInc As String
		Dim saveJGc As String
		Dim saveHin As String
		Dim aOpcode As String
		Dim aOpdate As String
		Dim aOptime As String
		Dim eOpcode As String
		Dim eOpdate As String
		Dim eOptime As String
		
		
		saveInc = dst.Inc_code
		saveJGc = dst.jg_code
		saveHin = dst.hin_code
		aOpcode = dst.Entry_Op_code
		aOpdate = dst.Entry_Op_date
		aOptime = dst.Entry_Op_time
		aOpcode = dst.Edit_Op_code
		aOpdate = dst.Edit_Op_date
		aOptime = dst.Edit_Op_time
		
		'UPGRADE_WARNING: �I�u�W�F�N�g dst �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		dst = src
		
		dst.Inc_code = saveInc
		dst.jg_code = saveJGc
		dst.hin_code = saveHin
		dst.Entry_Op_code = aOpcode
		dst.Entry_Op_date = aOpdate
		dst.Entry_Op_time = aOptime
		dst.Edit_Op_code = aOpcode
		dst.Edit_Op_date = aOpdate
		dst.Edit_Op_time = aOptime
		
	End Sub
	
	
	Private Sub TBL_SET()
		
		'   �O���[�v�̐ݒ�
		'   GRP1
		'           OptionButton�����敪
		CTRLTBL(N999).IGRP = GRP1
		'   GRP2
		'           ��ЁA���Ə��R�[�h
		CTRLTBL(N010).IGRP = GRP2
		CTRLTBL(N020).IGRP = GRP2
		'   GRP3
		'           �i��
		CTRLTBL(N030).IGRP = GRP3
		'   GRP4
		'           �i�����炻�̑��o�[�R�[�h
		CTRLTBL(N040).IGRP = GRP4
		CTRLTBL(N050).IGRP = GRP4
		CTRLTBL(N060).IGRP = GRP4
		CTRLTBL(N065).IGRP = GRP4 'A-CUST-20100610
		CTRLTBL(N070).IGRP = GRP4
		CTRLTBL(N080).IGRP = GRP4
		CTRLTBL(N090).IGRP = GRP4
		'   GRP5
		'           �K�p���A�����A�_�񉿊i�̂P
		CTRLTBL(N100_1).IGRP = GRP5
		CTRLTBL(N110_1).IGRP = GRP5
		CTRLTBL(N120_1).IGRP = GRP5
		'   GRP6
		'           �K�p���A�����A�_�񉿊i�̂Q
		CTRLTBL(N100_2).IGRP = GRP6
		CTRLTBL(N110_2).IGRP = GRP6
		CTRLTBL(N120_2).IGRP = GRP6
		'   GRP7
		'           ��p�Ȗ�
		CTRLTBL(N130).IGRP = GRP7
		CTRLTBL(N140).IGRP = GRP7
		'A-CUST20130212��
		CTRLTBL(N150).IGRP = GRP7
		CTRLTBL(N160).IGRP = GRP7
		CTRLTBL(N165).IGRP = GRP7 'A-20240115
		CTRLTBL(N170CMB).IGRP = GRP7
		CTRLTBL(N170).IGRP = GRP7
		'A-CUST20130212��
		'   GRP8
		'           �Ȗڕ���
		'D-20250201��
		'CTRLTBL(N210).IGRP = GRP8
		'CTRLTBL(N211).IGRP = GRP8
		'D-20250201��
		'   GRP9
		'           �啪�ނ��猟������
		CTRLTBL(N220).IGRP = GRP9
		CTRLTBL(N230).IGRP = GRP9
		CTRLTBL(N240).IGRP = GRP9
		'CTRLTBL(N250).IGRP = GRP9  'D-20250201
		CTRLTBL(N260).IGRP = GRP9
		'   GRP10
		'           ������i����e�`�w���M�܂�
		CTRLTBL(N270).IGRP = GRP10
		CTRLTBL(N280).IGRP = GRP10
		CTRLTBL(N290).IGRP = GRP10
		CTRLTBL(N291).IGRP = GRP10 'A-CUST20130212
		CTRLTBL(N300).IGRP = GRP10
		CTRLTBL(N310).IGRP = GRP10
		CTRLTBL(N320).IGRP = GRP10
		CTRLTBL(N330).IGRP = GRP10
		CTRLTBL(N340).IGRP = GRP10
		CTRLTBL(N370).IGRP = GRP10 'A-20250201
		'   GRP11
		'           �����P��
		CTRLTBL(N350_1).IGRP = GRP11
		CTRLTBL(N360_1).IGRP = GRP11
		CTRLTBL(N350_2).IGRP = GRP11
		CTRLTBL(N360_2).IGRP = GRP11
		CTRLTBL(N350_3).IGRP = GRP11
		CTRLTBL(N360_3).IGRP = GRP11
		CTRLTBL(N350_4).IGRP = GRP11
		CTRLTBL(N360_4).IGRP = GRP11
		CTRLTBL(N350_5).IGRP = GRP11
		CTRLTBL(N360_5).IGRP = GRP11
		'   GRP12
		'           �ƎҌ���
		CTRLTBL(N410).IGRP = GRP12
		'   GRP13
		'           ��������
		CTRLTBL(N420).IGRP = GRP13
		'   GRP14
		'           ���ꔭ�����爵���x�~�܂�
		CTRLTBL(N430).IGRP = GRP13
		'CTRLTBL(N440).IGRP = GRP13 'D-20250201
		CTRLTBL(N450).IGRP = GRP13
		CTRLTBL(N460).IGRP = GRP13
		CTRLTBL(N470).IGRP = GRP13
		CTRLTBL(N480).IGRP = GRP13
		CTRLTBL(N490).IGRP = GRP13
		CTRLTBL(N500).IGRP = GRP13
		CTRLTBL(N510).IGRP = GRP13
		
		CTRLTBL(NF12).IGRP = GEND
		CTRLTBL(NEND).IGRP = GEND
		
		'�����ځA�O���ڂ̐ݒ�
		CTRLTBL(N999).INEXT = N030 '   �����敪
		CTRLTBL(N999).IBACK = N999
		CTRLTBL(N999).IDOWN = N030
		
		CTRLTBL(N010).INEXT = N020 '   ��ЃR�[�h
		CTRLTBL(N010).IBACK = N999
		CTRLTBL(N010).IDOWN = N020
		
		CTRLTBL(N020).INEXT = N030 '   ���Ə��R�[�h
		CTRLTBL(N020).IBACK = N010
		CTRLTBL(N020).IDOWN = N030
		
		CTRLTBL(N030).INEXT = N040 '   �i��
		CTRLTBL(N030).IBACK = N020
		CTRLTBL(N030).IDOWN = N040
		
		CTRLTBL(N040).INEXT = N050 '   �i��
		CTRLTBL(N040).IBACK = N030
		CTRLTBL(N040).IDOWN = N050
		
		CTRLTBL(N050).INEXT = N060 '   �K�i
		CTRLTBL(N050).IBACK = N040
		CTRLTBL(N050).IDOWN = N060
		
		'D-CUST-20100610 Start
		'CTRLTBL(N060).INEXT = N070      '   �P��
		'CTRLTBL(N060).IBACK = N050
		'CTRLTBL(N060).IDOWN = N070
		'D-CUST-20100610 End
		'A-CUST-20100610 Start
		CTRLTBL(N060).INEXT = N065 '   �P��
		CTRLTBL(N060).IBACK = N050
		CTRLTBL(N060).IDOWN = N065
		
		CTRLTBL(N065).INEXT = N070 '   ��������
		CTRLTBL(N065).IBACK = N060
		CTRLTBL(N065).IDOWN = N070
		'A-CUST-20100610 End
		
		CTRLTBL(N070).INEXT = N080 '   Jan�W��
		'CTRLTBL(N070).IBACK = N060                     'D-CUST-20100610
		CTRLTBL(N070).IBACK = N065 'A-CUST-20100610
		CTRLTBL(N070).IDOWN = N080
		
		CTRLTBL(N080).INEXT = N090 '   Jan�Z�k
		CTRLTBL(N080).IBACK = N070
		CTRLTBL(N080).IDOWN = N090
		
		CTRLTBL(N090).INEXT = N100_1 '   ���̑��o�[�R�[�h
		CTRLTBL(N090).IBACK = N080
		CTRLTBL(N090).IDOWN = N100_1
		'                       �����E�o���Ȗ�
		CTRLTBL(N100_1).INEXT = N110_1 '   �K�p���P
		CTRLTBL(N100_1).IBACK = N090
		CTRLTBL(N100_1).IDOWN = N110_1
		
		CTRLTBL(N110_1).INEXT = N120_1 '   �����P
		CTRLTBL(N110_1).IBACK = N100_1
		CTRLTBL(N110_1).IDOWN = N120_1
		
		CTRLTBL(N120_1).INEXT = N100_2 '   �_�񉿊i�P
		CTRLTBL(N120_1).IBACK = N110_1 ' ***
		CTRLTBL(N120_1).IDOWN = N100_2
		
		CTRLTBL(N100_2).INEXT = N110_2 '   �K�p���Q
		CTRLTBL(N100_2).IBACK = N100_1 ' ***
		CTRLTBL(N100_2).IDOWN = N110_2
		
		CTRLTBL(N110_2).INEXT = N120_2 '   �����Q
		CTRLTBL(N110_2).IBACK = N100_2
		CTRLTBL(N110_2).IDOWN = N120_2
		
		CTRLTBL(N120_2).INEXT = N130 '   �_�񉿊i�Q
		CTRLTBL(N120_2).IBACK = N110_2 ' ***
		CTRLTBL(N120_2).IDOWN = N130
		
		CTRLTBL(N130).INEXT = N140 '   ��p�Ȗ�
		CTRLTBL(N130).IBACK = N100_2
		CTRLTBL(N130).IDOWN = N140
		
		'D-CUST20130212��
		'    CTRLTBL(N140).INEXT = N210    '   ��p�Ȗ�
		'    CTRLTBL(N140).IBACK = N130
		'    CTRLTBL(N140).IDOWN = N210
		
		'    CTRLTBL(N210).INEXT = N211    '   �Ȗڕ���
		'    CTRLTBL(N210).IBACK = N130
		'    CTRLTBL(N210).IDOWN = N211
		'D-CUST20130212��
		'A-CUST20130212 ��
		CTRLTBL(N140).INEXT = N150 '   ��p�Ȗ�
		CTRLTBL(N140).IBACK = N130
		CTRLTBL(N140).IDOWN = N150
		
		CTRLTBL(N150).INEXT = N160 '   ���Y��
		CTRLTBL(N150).IBACK = N140
		CTRLTBL(N150).IDOWN = N160
		'D-20240115��
		'CTRLTBL(N160).INEXT = N170CMB    '   �d��
		'CTRLTBL(N160).IBACK = N150
		'CTRLTBL(N160).IDOWN = N170CMB
		
		'CTRLTBL(N170CMB).INEXT = N170    '   �ܖ������R���{
		'CTRLTBL(N170CMB).IBACK = N160
		'CTRLTBL(N170CMB).IDOWN = N170
		
		'CTRLTBL(N170).INEXT = N210    '   �ܖ�����
		'CTRLTBL(N170).IBACK = N170CMB
		'CTRLTBL(N170).IDOWN = N210
		'D-20240115��
		
		'A-20240115��
		CTRLTBL(N160).INEXT = N165 '   �d��
		CTRLTBL(N160).IBACK = N150
		CTRLTBL(N160).IDOWN = N165
		
		CTRLTBL(N165).INEXT = N170CMB '   ����/�ܖ������敪
		CTRLTBL(N165).IBACK = N160
		CTRLTBL(N165).IDOWN = N170CMB
		
		CTRLTBL(N170CMB).INEXT = N170 '   �ܖ������R���{
		CTRLTBL(N170CMB).IBACK = N165
		CTRLTBL(N170CMB).IDOWN = N170
		
		CTRLTBL(N170).INEXT = N175 '   �ܖ�����
		CTRLTBL(N170).IBACK = N170CMB
		CTRLTBL(N170).IDOWN = N175
		
		'D-20250201��
		'CTRLTBL(N175).INEXT = N210
		'CTRLTBL(N175).IBACK = N170
		'CTRLTBL(N175).IDOWN = N210
		'A-20240115��
		
		'CTRLTBL(N210).INEXT = N211    '   �Ȗڕ���
		'CTRLTBL(N210).IBACK = N170       'D-20240115
		'CTRLTBL(N210).IBACK = N175        'A-20240115
		'CTRLTBL(N210).IDOWN = N211
		
		
		'A-CUST20130212��
		'                   �e�핪�ސ���
		
		
		'CTRLTBL(N211).INEXT = N220    '   �Ȗڕ���
		'CTRLTBL(N211).IBACK = N210
		'CTRLTBL(N211).IDOWN = N220
		'D-20250201��
		'A-20250201��
		CTRLTBL(N175).INEXT = N220
		CTRLTBL(N175).IBACK = N170
		CTRLTBL(N175).IDOWN = N220
		'A-20250201��
		
		CTRLTBL(N220).INEXT = N230 '   �啪��
		'CTRLTBL(N220).IBACK = N210 'D-20250201
		CTRLTBL(N220).IBACK = N175 'A-20250201
		CTRLTBL(N220).IDOWN = N230
		
		CTRLTBL(N230).INEXT = N240 '   ������
		CTRLTBL(N230).IBACK = N220
		CTRLTBL(N230).IDOWN = N240
		
		'D-20250201��
		'CTRLTBL(N240).INEXT = N250    '   ������
		'CTRLTBL(N240).IBACK = N230
		'CTRLTBL(N240).IDOWN = N250
		
		'CTRLTBL(N250).INEXT = N260    '   ����
		'CTRLTBL(N250).IBACK = N240
		'CTRLTBL(N250).IDOWN = N260
		'D-20250201��
		'A-20250201��
		CTRLTBL(N240).INEXT = N260 '   ������
		CTRLTBL(N240).IBACK = N230
		CTRLTBL(N240).IDOWN = N260
		'A-20250201��
		
		CTRLTBL(N260).INEXT = N270 '   ��������
		'CTRLTBL(N260).IBACK = N250 'D-20250201
		CTRLTBL(N260).IBACK = N240 'A-20250201
		CTRLTBL(N260).IDOWN = N270
		
		CTRLTBL(N270).INEXT = N280 '   ������i
		CTRLTBL(N270).IBACK = N260
		CTRLTBL(N270).IDOWN = N280
		
		CTRLTBL(N280).INEXT = N290 '   �d�|�敪
		CTRLTBL(N280).IBACK = N270
		CTRLTBL(N280).IDOWN = N290
		'D-CUST20130212��
		'    CTRLTBL(N290).INEXT = N300    '   ���c����
		'    CTRLTBL(N290).IBACK = N280
		'    CTRLTBL(N290).IDOWN = N300
		
		'    CTRLTBL(N300).INEXT = N310    '   �Ǘ��敪
		'    CTRLTBL(N300).IBACK = N290
		'    CTRLTBL(N300).IDOWN = N310
		'D-CUST20130212��
		'A-CUST20130212��
		CTRLTBL(N290).INEXT = N291 '   ���c����
		CTRLTBL(N290).IBACK = N280
		CTRLTBL(N290).IDOWN = N291
		
		'D-20250201��
		'CTRLTBL(N291).INEXT = N300    '   JAN���i����
		'CTRLTBL(N291).IBACK = N290
		'CTRLTBL(N291).IDOWN = N300
		
		'CTRLTBL(N300).INEXT = N310    '   �Ǘ��敪
		'CTRLTBL(N300).IBACK = N291
		'CTRLTBL(N300).IDOWN = N310
		'A-CUST20130212��
		
		
		
		'CTRLTBL(N310).INEXT = N320    '   �����
		'CTRLTBL(N310).IBACK = N300
		'CTRLTBL(N310).IDOWN = N320
		
		'CTRLTBL(N320).INEXT = N330    '   �I���P��
		'CTRLTBL(N320).IBACK = N310
		'CTRLTBL(N320).IDOWN = N330
		
		'CTRLTBL(N330).INEXT = N340    '   �݌ɊǗ�
		'CTRLTBL(N330).IBACK = N320
		'CTRLTBL(N330).IDOWN = N340
		
		'CTRLTBL(N340).INEXT = N350_1    '   �e�`�w���M
		'CTRLTBL(N340).IBACK = N330
		'CTRLTBL(N340).IDOWN = N350_1
		'D-20250201��
		'A-20250201��
		CTRLTBL(N291).INEXT = N310 '   JAN���i����
		CTRLTBL(N291).IBACK = N290
		CTRLTBL(N291).IDOWN = N310
		
		CTRLTBL(N310).INEXT = N370 '   �����
		CTRLTBL(N310).IBACK = N291
		CTRLTBL(N310).IDOWN = N370
		
		CTRLTBL(N370).INEXT = N330 '   �ŗ��敪
		CTRLTBL(N370).IBACK = N310
		CTRLTBL(N370).IDOWN = N330
		
		CTRLTBL(N330).INEXT = N300 '   �݌ɊǗ�
		CTRLTBL(N330).IBACK = N370
		CTRLTBL(N330).IDOWN = N300
		
		CTRLTBL(N300).INEXT = N320 '   �Ǘ��敪
		CTRLTBL(N300).IBACK = N330
		CTRLTBL(N300).IDOWN = N320
		
		CTRLTBL(N320).INEXT = N350_1 '   �I���P��
		CTRLTBL(N320).IBACK = N300
		CTRLTBL(N320).IDOWN = N350_1
		'A-20250201��
		
		CTRLTBL(N350_1).INEXT = N360_1 '   �����P��
		'CTRLTBL(N350_1).IBACK = N340   'D-20250201
		CTRLTBL(N350_1).IBACK = N320 'A-20250201
		CTRLTBL(N350_1).IDOWN = N350_2
		
		CTRLTBL(N360_1).INEXT = N350_2 '   ���Z��
		CTRLTBL(N360_1).IBACK = N350_1
		CTRLTBL(N360_1).IDOWN = N350_2
		
		CTRLTBL(N350_2).INEXT = N360_2 '   �����P��
		CTRLTBL(N350_2).IBACK = N350_1
		CTRLTBL(N350_2).IDOWN = N350_3
		
		CTRLTBL(N360_2).INEXT = N350_3 '   ���Z��
		CTRLTBL(N360_2).IBACK = N350_2
		CTRLTBL(N360_2).IDOWN = N350_3
		
		CTRLTBL(N350_3).INEXT = N360_3 '   �����P��
		CTRLTBL(N350_3).IBACK = N350_2
		CTRLTBL(N350_3).IDOWN = N350_4
		
		CTRLTBL(N360_3).INEXT = N350_4 '   ���Z��
		CTRLTBL(N360_3).IBACK = N350_3
		CTRLTBL(N360_3).IDOWN = N350_4
		
		CTRLTBL(N350_4).INEXT = N360_4 '   �����P��
		CTRLTBL(N350_4).IBACK = N350_3
		CTRLTBL(N350_4).IDOWN = N350_5
		
		CTRLTBL(N360_4).INEXT = N350_5 '   ���Z��
		CTRLTBL(N360_4).IBACK = N350_4
		CTRLTBL(N360_4).IDOWN = N350_5
		
		CTRLTBL(N350_5).INEXT = N360_5 '   �����P��
		CTRLTBL(N350_5).IBACK = N350_4
		CTRLTBL(N350_5).IDOWN = N410
		
		CTRLTBL(N360_5).INEXT = N410 '   ���Z��
		CTRLTBL(N360_5).IBACK = N350_5
		CTRLTBL(N360_5).IDOWN = N410
		'                           ���̑�
		CTRLTBL(N410).INEXT = N420 '   �ƎҌ���
		CTRLTBL(N410).IBACK = N350_1
		CTRLTBL(N410).IDOWN = N420
		
		CTRLTBL(N420).INEXT = N430 '   ��������
		CTRLTBL(N420).IBACK = N410
		CTRLTBL(N420).IDOWN = N430
		
		'D-20250201��
		'CTRLTBL(N430).INEXT = N440          '  ���ꔭ����
		'CTRLTBL(N430).IBACK = N420
		'CTRLTBL(N430).IDOWN = N440
		
		'CTRLTBL(N440).INEXT = N450          '  ����ŗ��敪
		'CTRLTBL(N440).IBACK = N430
		'CTRLTBL(N440).IDOWN = N450
		'D-20250201��
		'A-20250201��
		CTRLTBL(N430).INEXT = N450 '  ���ꔭ����
		CTRLTBL(N430).IBACK = N420
		CTRLTBL(N430).IDOWN = N450
		'A-20250201��
		
		CTRLTBL(N450).INEXT = N460 '  �����i
		'CTRLTBL(N450).IBACK = N440 'D-20250201
		CTRLTBL(N450).IBACK = N430 'A-20250201
		CTRLTBL(N450).IDOWN = N460
		
		CTRLTBL(N460).INEXT = N470 '  ���̋@�̔�
		CTRLTBL(N460).IBACK = N450
		CTRLTBL(N460).IDOWN = N470
		
		CTRLTBL(N470).INEXT = N480 '  ����Ώ�
		CTRLTBL(N470).IBACK = N460
		CTRLTBL(N470).IDOWN = N480
		
		CTRLTBL(N480).INEXT = N490 '  �ŏI�[�i��
		CTRLTBL(N480).IBACK = N470
		CTRLTBL(N480).IDOWN = N490
		
		CTRLTBL(N490).INEXT = N500 '  �K�p�J�n���t
		CTRLTBL(N490).IBACK = N480
		CTRLTBL(N490).IDOWN = N500
		
		CTRLTBL(N500).INEXT = N510 '  �����x�~
		CTRLTBL(N500).IBACK = N490
		CTRLTBL(N500).IDOWN = N510
		
		CTRLTBL(N510).INEXT = NF12 '  �����x�~��
		CTRLTBL(N510).IBACK = N500
		CTRLTBL(N510).IDOWN = NF12
		
		CTRLTBL(NF12).INEXT = NEND '  ���s
		CTRLTBL(NF12).IBACK = N510
		CTRLTBL(NF12).IDOWN = NEND
		'
		'    CTRLTBL(N520).INEXT = N510       '  �_�~�[
		'    CTRLTBL(N520).IBACK = N510
		'    CTRLTBL(N520).IDOWN = N510
		'
		
		CTRLTBL(N999).CTRL = OPTO999(1)
		CTRLTBL(N010).CTRL = IMTX010
		CTRLTBL(N020).CTRL = IMTX020
		CTRLTBL(N030).CTRL = IMTX030
		CTRLTBL(N040).CTRL = IMTX040
		CTRLTBL(N050).CTRL = IMTX050
		CTRLTBL(N060).CTRL = CMB060
		CTRLTBL(N065).CTRL = IMTX065 'A-CUST-20100610
		CTRLTBL(N070).CTRL = IMTX070
		CTRLTBL(N080).CTRL = IMTX080
		CTRLTBL(N090).CTRL = IMTX090
		'                   TAB0-�����E�o���Ȗ�
		CTRLTBL(N100_1).CTRL = IMTX100(1)
		CTRLTBL(N110_1).CTRL = IMNU110(1)
		CTRLTBL(N120_1).CTRL = IMNU120(1)
		CTRLTBL(N100_2).CTRL = IMTX100(2)
		CTRLTBL(N110_2).CTRL = IMNU110(2)
		CTRLTBL(N120_2).CTRL = IMNU120(2)
		CTRLTBL(N130).CTRL = IMTX130(1)
		CTRLTBL(N140).CTRL = IMTX140(1)
		'A-CUST20130212��
		CTRLTBL(N150).CTRL = IMTX150(0)
		CTRLTBL(N160).CTRL = IMNU160(0)
		CTRLTBL(N165).CTRL = CMB165 'A-20240115
		CTRLTBL(N170CMB).CTRL = CMB170
		CTRLTBL(N170).CTRL = IMNU170(1)
		CTRLTBL(N175).CTRL = IMNU175(0) 'A-20240115
		'A-CUST20130212��
		'                   TAB1�e�핪�ސ���
		'D-20250201��
		'Set CTRLTBL(N210).CTRL = IMTX210
		'Set CTRLTBL(N211).CTRL = IMTX211
		'D-20250201��
		CTRLTBL(N220).CTRL = IMTX220
		CTRLTBL(N230).CTRL = IMTX230
		CTRLTBL(N240).CTRL = IMTX240
		'Set CTRLTBL(N250).CTRL = IMTX250   'D-20250201
		CTRLTBL(N260).CTRL = IMTX260
		CTRLTBL(N270).CTRL = CHK270
		CTRLTBL(N280).CTRL = CHK280
		CTRLTBL(N290).CTRL = CHK290
		CTRLTBL(N291).CTRL = IMTX291 'A-CUST20130212
		CTRLTBL(N300).CTRL = OPTO300(1)
		CTRLTBL(N310).CTRL = OPTO310(1)
		CTRLTBL(N320).CTRL = OPTO320(1)
		CTRLTBL(N330).CTRL = OPTO330(1)
		CTRLTBL(N340).CTRL = OPTO340(1)
		CTRLTBL(N350_1).CTRL = CMB350(1)
		CTRLTBL(N360_1).CTRL = IMNU360(1)
		CTRLTBL(N350_2).CTRL = CMB350(2)
		CTRLTBL(N360_2).CTRL = IMNU360(2)
		CTRLTBL(N350_3).CTRL = CMB350(3)
		CTRLTBL(N360_3).CTRL = IMNU360(3)
		CTRLTBL(N350_4).CTRL = CMB350(4)
		CTRLTBL(N360_4).CTRL = IMNU360(4)
		CTRLTBL(N350_5).CTRL = CMB350(5)
		CTRLTBL(N360_5).CTRL = IMNU360(5)
		
		CTRLTBL(N370).CTRL = CMB370 'A-20250201
		
		CTRLTBL(N410).CTRL = IMTX410
		CTRLTBL(N420).CTRL = SPR420
		CTRLTBL(N430).CTRL = CHK430
		CTRLTBL(N440).CTRL = IMTX440
		CTRLTBL(N450).CTRL = CHK450
		CTRLTBL(N460).CTRL = CHK460
		CTRLTBL(N470).CTRL = CHK470
		CTRLTBL(N480).CTRL = IMTX480
		CTRLTBL(N490).CTRL = IMTX490
		CTRLTBL(N500).CTRL = CHK500
		CTRLTBL(N510).CTRL = IMTX510
		CTRLTBL(NF12).CTRL = CMDOFNC(12)
		
		'Const N520 = N510 + 1
		'Const NEND = N520 + 1
		
		CUR_NO = N010
		NXT_NO = n0
		LST_NO = n0
		
		
	End Sub
	
	'Private Sub SCRCLR_RTN()                                           'D-CUST-20100610
	Private Sub SCRCLR_RTN(Optional ByVal CODECLR As Boolean = True) 'A-CUST-20100610
		'A-CUST-20100610 Start
		Dim SvWKB030 As String
		
		If Not CODECLR Then
			If KBKBN <> F_ADD Then
				CODECLR = True
			Else
				SvWKB030 = WKB030
			End If
		End If
		'A-CUST-20100610 Enf
		
		Call SCR_INIT_RTN()
		
		'    WKB010 = WG_INCCODE
		'    WKB020 = WG_JGCODE
		
		KB.Inc_code = WKB010
		KB.jg_code = WKB020
		'A-CUST-20100610 Start
		If Not CODECLR Then
			WKB030 = SvWKB030
			KB.hin_code = SvWKB030
		End If
		'A-CUST-20100610 End
		
		Call SpreadInit()
		Call SCR_DSPDATA()
		If KBKBN = 3 Then Call SetMode("D")
		'   TAB�ŏ���TAB�ɐݒ�          NR-SZ0410-2
		TAB010.SelectedIndex = 0
		SentakuFLG = False 'A-CUST-20100610
		
		
	End Sub
	
	Private Sub SZ0410FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		If Trim(IMTX030.Text) <> "" Then
			If MsgBox("�I�����܂����H", MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question, "�d���i�ڊ�{������") = MsgBoxResult.No Then
				Cancel = True
			End If
		End If
		
		eventArgs.Cancel = Cancel
	End Sub

	'Private Sub SZ0410FRM_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed 'D-20250417
	Private Sub SZ0410FRM_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing 'A-20250417

        'Call ZAEND_SUB                         'D-CUST-20100610
        'UPGRADE_ISSUE: Event �p�����[�^ Cancel �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FB723E3C-1C06-4D2B-B083-E6CD0D334DA8"' ���N���b�N���Ă��������B
        'Cancel = True 'D-20250417
        eventArgs.Cancel = True 'A-20250417
        Call ENDR_RTN() 'A-CUST-20100610

	End Sub





	Private Sub IMNU110_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU110.Enter
		Dim Index As Short = IMNU110.GetIndex(eventSender)
		
		Select Case Index
			Case 1
				If CUR_NO = N110_1 Then Exit Sub
				CUR_NO = N110_1
			Case 2
				If CUR_NO = N110_2 Then Exit Sub
				CUR_NO = N110_2
		End Select
		
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		
		ZAKB_SW = 0
		
		'�m��
		'    If LST_NO = N100_2 Then
		'        NXT_NO = N130
		'        Call SET_NO(0)
		'        Exit Sub
		'    End If
		
		If CUR_NO = N110_2 And Trim(IMTX100(2).Text) = "" Then
			'        NXT_NO = LST_NO
			'        LST_NO = N120_2
			If LST_NO < CUR_NO Then
				NXT_NO = N130
				Call SET_NO(0)
				Exit Sub
			End If
			
		End If
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMNU110_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU110.KeyDownEvent 'D-20250417
	Private Sub IMNU110_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyDownEvent) Handles IMNU110.KeyDownEvent 'A-20250417
		Dim Index As Short = IMNU110.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	'Private Sub IMNU110_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU110.KeyPressEvent 'D-20250417
	Private Sub IMNU110_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyPressEvent) Handles IMNU110.KeyPressEvent 'A-20250417
		Dim Index As Short = IMNU110.GetIndex(eventSender)

		Call ZAKB_SUB(eventArgs.KeyAscii)

	End Sub

	Private Sub IMNU120_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU120.Enter
		Dim Index As Short = IMNU120.GetIndex(eventSender)
		
		Select Case Index
			Case 1
				If CUR_NO = N120_1 Then Exit Sub
				CUR_NO = N120_1
			Case 2
				If CUR_NO = N120_2 Then Exit Sub
				CUR_NO = N120_2
		End Select
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		ZAKB_SW = 0
		
		'�m��
		'    If LST_NO = N100_2 Then
		'        If CUR_NO = N120_1 Then
		'        NXT_NO = N130
		'        Call SET_NO(0)
		'        Exit Sub
		'        End If
		'    End If
		If CUR_NO = N120_2 And Trim(IMTX100(2).Text) = "" Then
			If LST_NO < CUR_NO Then
				NXT_NO = LST_NO
				Call SET_NO(0)
				Exit Sub
			End If
			
		End If
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMNU120_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU120.KeyDownEvent 'D-20250417
	Private Sub IMNU120_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyDownEvent) Handles IMNU120.KeyDownEvent 'A-20250417
		Dim Index As Short = IMNU120.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	'Private Sub IMNU120_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU120.KeyPressEvent 'D-20250417
	Private Sub IMNU120_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyPressEvent) Handles IMNU120.KeyPressEvent 'A-20250417
		Dim Index As Short = IMNU120.GetIndex(eventSender)
		Call ZAKB_SUB(eventArgs.KeyAscii)

	End Sub

	'A-CUST20130212��
	Private Sub IMNU160_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU160.Enter
		Dim Index As Short = IMNU160.GetIndex(eventSender)
		If CUR_NO = N160 Then Exit Sub
		
		CUR_NO = N160
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		ZAKB_SW = 0
	End Sub
    'A-CUST20130212��
    'A-CUST20130212��

    'Private Sub IMNU160_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU160.KeyDownEvent 'D-20250417
    Private Sub IMNU160_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyDownEvent) Handles IMNU160.KeyDownEvent 'A-20250417
        Dim Index As Short = IMNU160.GetIndex(eventSender)
        Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
    End Sub

	'Private Sub IMNU160_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU160.KeyPressEvent 'D-20250417
	Private Sub IMNU160_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyPressEvent) Handles IMNU160.KeyPressEvent 'A-20250417
		Dim Index As Short = IMNU160.GetIndex(eventSender)
		Call ZAKB_SUB(eventArgs.KeyAscii)

	End Sub

	'A-CUST20130212��
	'A-CUST20130212��
	Private Sub IMNU170_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU170.Enter
		Dim Index As Short = IMNU170.GetIndex(eventSender)
		If CUR_NO = N170 Then Exit Sub
		
		CUR_NO = N170
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		ZAKB_SW = 0
	End Sub
	'A-CUST20130212��
	
	
	'A-20240115��
	Private Sub IMNU175_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU175.Enter
		Dim Index As Short = IMNU175.GetIndex(eventSender)
		If LST_NO = N170 Or (LST_NO = N165 And CTRLTBL(N170CMB).CTRL.Enabled = False) Then
			'NXT_NO = N210  'D-20250201
			NXT_NO = N220 'A-20250201
			CTRLTBL(NXT_NO).CTRL.Focus()
			'ElseIf LST_NO = N210 Then  'D-20250201
		ElseIf LST_NO = N220 Then  'A-20250201
			
			If CDbl(RTrim(CStr(CMB165.SelectedIndex))) <> 0 Then
				NXT_NO = N170
			Else
				NXT_NO = N165
			End If
			CTRLTBL(NXT_NO).CTRL.Focus()
		End If
	End Sub
    'A-20240115��

    'A-CUST20130212��
    'Private Sub IMNU170_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU170.KeyDownEvent 'D-20250417
    Private Sub IMNU170_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyDownEvent) Handles IMNU170.KeyDownEvent 'A-20250417
        Dim Index As Short = IMNU170.GetIndex(eventSender)
        Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
    End Sub
	'Private Sub IMNU170_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU170.KeyPressEvent 'D-20250417
	Private Sub IMNU170_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyPressEvent) Handles IMNU170.KeyPressEvent 'A-20250417
		Dim Index As Short = IMNU170.GetIndex(eventSender)
		Call ZAKB_SUB(eventArgs.KeyAscii)
	End Sub
    'A-CUST20130212��

    Private Sub IMNU360_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU360.Enter
        Dim Index As Short = IMNU360.GetIndex(eventSender)

        Dim iCur As Short

        Select Case Index
            Case 1
                iCur = N360_1
            Case 2
                iCur = N360_2
            Case 3
                iCur = N360_3
            Case 4
                iCur = N360_4
            Case 5
                iCur = N360_5
            Case Else
                iCur = 0
        End Select

        If CUR_NO = iCur Then Exit Sub

        CUR_NO = iCur
        System.Diagnostics.Debug.Assert(CUR_NO > 0, "")


        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                Exit Sub
            End If
            If GPROCHK() = False Then
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If
        ZAKB_SW = 0
        '�m��
        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

    'Private Sub IMNU360_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU360.KeyDownEvent 'D-20250417	
    Private Sub IMNU360_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyDownEvent) Handles IMNU360.KeyDownEvent 'A-20250417
        Dim Index As Short = IMNU360.GetIndex(eventSender)

        Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

    End Sub

	'Private Sub IMNU360_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU360.KeyPressEvent 'D-20250417
	Private Sub IMNU360_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximNumber6.INumEvents_KeyPressEvent) Handles IMNU360.KeyPressEvent 'A-20250417
		Dim Index As Short = IMNU360.GetIndex(eventSender)

		Call ZAKB_SUB(eventArgs.KeyAscii)

	End Sub

    Private Sub IMTX010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX010.Enter

        OPTO999(1).TabStop = True
        OPTO999(2).TabStop = True
        OPTO999(3).TabStop = True

        If CUR_NO = N010 Then Exit Sub

        CUR_NO = N010

        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                Exit Sub
            End If
            If GPROCHK() = False Then
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If
        '�m��
        ''''Debug.Print "IMTX010_GotFocus LST_NO before="; LST_NO
        LST_NO = CUR_NO
        ''''Debug.Print "IMTX010_GotFocus LST_NO After ="; LST_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

	'Private Sub IMTX010_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX010.KeyDownEvent 'D-20250417
	Private Sub IMTX010_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX010.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

    Private Sub IMTX020_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX020.Enter

        If CUR_NO = N020 Then Exit Sub

        CUR_NO = N020
        Debug.Print("020 GotFocus:" & LST_NO)
        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                Exit Sub
            End If
            If GPROCHK() = False Then
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If
        '�m��
        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

	'Private Sub IMTX020_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX020.KeyDownEvent 'D-20250417
	Private Sub IMTX020_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX020.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

    Private Sub IMTX030_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX030.Enter
        Dim nnum As Integer

        'Debug.Assert CMDOFNC(0).Enabled
        'Debug.Assert CMDOFNC(12).Enabled

        If CUR_NO = N030 Then Exit Sub

        CUR_NO = N030

        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                bSPRNotReady = True
                Exit Sub
            End If
            If GPROCHK() = False Then
                bSPRNotReady = True
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If

        'D-CUST-20100610 Start
        'If KBKBN = F_ADD And IMTX030.Text = "" Then
        '    nnum = New_Number
        '    If nnum < 0 Or nnum > "99999" Then
        '        Call MsgBox("�����̔Ԃ�����ɒB���܂����B" + Chr(10) + _
        ''        "�̔Ԃ���܂���  �i�ڃR�[�h����͂��Ă��������B�@", _
        ''        vbApplicationModal + vbExclamation, "�d���i�ڊ�{������")
        '        IMTX030.Text = ""
        '    Else
        '        IMTX030.Text = nnum
        '    End If
        'End If
        'D-CUST-20100610 End
        '�m��
        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

	'Private Sub IMTX030_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX030.KeyDownEvent 'D-20250417
	Private Sub IMTX030_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX030.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX040_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX040.Enter

		'UPGRADE_NOTE: IMEMode �� CtlIMEMode �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		'IMTX040.CtlIMEMode = OsktxtLibV5.CIMEMODE.�S�p�Ђ炪�� 'A-20160726- 'D-20250417
		IMTX040.ImeMode = ImeMode.Hiragana                    'A-20250417

		If CUR_NO = N040 Then Exit Sub
		
		CUR_NO = N040
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX040_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX040.KeyDownEvent 'D-0250417
	Private Sub IMTX040_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX040.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX050_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX050.Enter

		'UPGRADE_NOTE: IMEMode �� CtlIMEMode �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		'IMTX050.CtlIMEMode = OsktxtLibV5.CIMEMODE.�S�p�Ђ炪�� 'A-20160726-
		IMTX050.ImeMode = ImeMode.Hiragana                      'A-20250417

		If CUR_NO = N050 Then Exit Sub
		
		CUR_NO = N050
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX050_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX050.KeyDownEvent 'D-20250417
	Private Sub IMTX050_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX050.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

    'A-CUST-20100610 Start
    Private Sub IMTX065_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX065.Enter

        'UPGRADE_NOTE: IMEMode �� CtlIMEMode �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
        'IMTX065.CtlIMEMode = OsktxtLibV5.CIMEMODE.�S�p�Ђ炪�� 'A-20160726-
        IMTX065.ImeMode = ImeMode.Hiragana                      'A-20250417

        If CUR_NO = N065 Then Exit Sub

        CUR_NO = N065

        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                Exit Sub
            End If
            If GPROCHK() = False Then
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If
        '�m��
        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

	'Private Sub IMTX065_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX065.KeyDownEvent 'D-20250417
	Private Sub IMTX065_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX065.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub
	'A-CUST-20100610 End

	Private Sub IMTX070_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX070.Enter
		
		If CUR_NO = N070 Then Exit Sub
		
		CUR_NO = N070
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX070_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX070.KeyDownEvent 'D-20250417
	Private Sub IMTX070_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX070.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX080_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX080.Enter
		
		If CUR_NO = N080 Then Exit Sub
		
		CUR_NO = N080
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX080_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX080.KeyDownEvent 'D-20250417
	Private Sub IMTX080_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX080.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX090_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX090.Enter
		
		If CUR_NO = N090 Then Exit Sub
		
		CUR_NO = N090
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX090_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX090.KeyDownEvent 'D-20240517
	Private Sub IMTX090_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX090.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX100_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX100.Enter
		Dim Index As Short = IMTX100.GetIndex(eventSender)
		
		'���t�f�[�^�Ȃ�΁A/�𔲂��ĕ\������
		If IsDate(IMTX100(Index).Text) Then
			'UPGRADE_WARNING: DateValue �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
			IMTX100(Index).Text = VB6.Format(DateValue(IMTX100(Index).Text), "yyyymmdd")
		End If
		
		Select Case Index
			Case 1
				If CUR_NO = N100_1 Then Exit Sub
				CUR_NO = N100_1
			Case 2
				If CUR_NO = N100_2 Then Exit Sub
				CUR_NO = N100_2
		End Select
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX100_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX100.KeyDownEvent 'D-20250417
	Private Sub IMTX100_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX100.KeyDownEvent 'A-20250417
		Dim Index As Short = IMTX100.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX130_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX130.Enter
		Dim Index As Short = IMTX130.GetIndex(eventSender)
		
		If Index <> 1 Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Sub
		End If
		
		If CUR_NO = N130 Then Exit Sub
		
		CUR_NO = N130
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX130_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX130.KeyDownEvent 'D-20250417
	Private Sub IMTX130_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX130.KeyDownEvent 'A-20250417
		Dim Index As Short = IMTX130.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX140_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX140.Enter
		Dim Index As Short = IMTX140.GetIndex(eventSender)
		
		If Index <> 1 Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Sub
		End If
		
		'   IMTX130,IMTX140��GRP1�ł��B
		If CUR_NO = N140 Then Exit Sub
		
		CUR_NO = N140
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX140_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX140.KeyDownEvent 'D-20250417
	Private Sub IMTX140_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX140.KeyDownEvent 'A-20250417
		Dim Index As Short = IMTX140.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub
	'A-CUST20130212��
	Private Sub IMTX150_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX150.Enter
		Dim Index As Short = IMTX150.GetIndex(eventSender)
		
		If CUR_NO = N150 Then Exit Sub
		
		CUR_NO = N150
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	'Private Sub IMTX150_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX150.KeyDownEvent 'D-20250417
	Private Sub IMTX150_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX150.KeyDownEvent 'A-20250417
		Dim Index As Short = IMTX150.GetIndex(eventSender)

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub
	'A-CUST20130212��

	'D-20250201��
	'Private Sub IMTX210_GotFocus()

	'If CUR_NO = N210 Then Exit Sub

	'CUR_NO = N210

	'�`�F�b�N
	'If LST_NO <> n0 Then
	'If IPROCHK() = False Then
	'Exit Sub
	'End If
	'If GPROCHK() = False Then
	'Exit Sub
	'End If
	'End If
	'If GVALCHK() = False Then
	'Exit Sub
	'End If
	'If MVALCHK() = False Then
	'Exit Sub
	'End If
	'�m��
	'LST_NO = CUR_NO
	'--- �t�@���N�V�������b�Z�[�W
	'Call FUNCSET_RTN

	'End Sub

	'Private Sub IMTX210_KeyDown(KeyCode As Integer, Shift As Integer)

	'Call Form_KeyDown(KeyCode, Shift)

	'End Sub

	'Private Sub IMTX211_GotFocus()

	'If CUR_NO = N211 Then Exit Sub

	'CUR_NO = N211

	'�`�F�b�N
	'If LST_NO <> n0 Then
	'If IPROCHK() = False Then
	'Exit Sub
	'End If
	'If GPROCHK() = False Then
	'Exit Sub
	'End If
	'End If
	'If GVALCHK() = False Then
	'Exit Sub
	'End If
	'If MVALCHK() = False Then
	'Exit Sub
	'End If
	'�m��
	'LST_NO = CUR_NO
	'--- �t�@���N�V�������b�Z�[�W
	'Call FUNCSET_RTN

	'End Sub

	'Private Sub IMTX211_KeyDown(KeyCode As Integer, Shift As Integer)

	'Call Form_KeyDown(KeyCode, Shift)

	'End Sub
	'D-20250201��

	Private Sub IMTX220_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX220.Enter
		
		If CUR_NO = N220 Then Exit Sub
		
		CUR_NO = N220
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX220_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX220.KeyDownEvent 'D-20250417
	Private Sub IMTX220_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX220.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX230_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX230.Enter
		
		If CUR_NO = N230 Then Exit Sub
		
		CUR_NO = N230
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX230_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX230.KeyDownEvent 'D-20250417
	Private Sub IMTX230_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX230.KeyDownEvent 'A-2025417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX240_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX240.Enter
		
		If CUR_NO = N240 Then Exit Sub
		
		CUR_NO = N240
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX240_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX240.KeyDownEvent 'D-20250417
	Private Sub IMTX240_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX240.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	'D-20250201��
	'Private Sub IMTX250_GotFocus()

	'If CUR_NO = N250 Then Exit Sub

	'CUR_NO = N250

	'�`�F�b�N
	'If LST_NO <> n0 Then
	'If IPROCHK() = False Then
	'Exit Sub
	'End If
	'If GPROCHK() = False Then
	'Exit Sub
	'End If
	'End If
	'If GVALCHK() = False Then
	'Exit Sub
	'End If
	'If MVALCHK() = False Then
	'Exit Sub
	'End If
	'�m��
	'LST_NO = CUR_NO
	'--- �t�@���N�V�������b�Z�[�W
	'Call FUNCSET_RTN

	'End Sub

	'Private Sub IMTX250_KeyDown(KeyCode As Integer, Shift As Integer)

	'Call Form_KeyDown(KeyCode, Shift)

	'End Sub
	'D-20250201��

	Private Sub IMTX260_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX260.Enter
		
		If CUR_NO = N260 Then Exit Sub
		
		CUR_NO = N260
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX260_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX260.KeyDownEvent 'D-20250417
	Private Sub IMTX260_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX260.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub


	'A-CUST20130212��
	Private Sub IMTX291_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX291.Enter
		
		If CUR_NO = N291 Then Exit Sub
		
		CUR_NO = N291
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	'Private Sub IMTX291_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX291.KeyDownEvent 'D-20250417
	Private Sub IMTX291_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX291.KeyDownEvent 'A-20250417
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub
	'A-CUST20130212��
	Private Sub IMTX410_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX410.Enter
		
		
		If CUR_NO = N410 Then Exit Sub
		
		CUR_NO = N410
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX410_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX410.KeyDownEvent 'D-20250417
	Private Sub IMTX410_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX410.KeyDownEvent 'A-20250417

		CTRLTBL(N350_1).CTRL.TabStop = True
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub


	Private Sub IMTX440_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX440.Enter
		
		If CUR_NO = N440 Then Exit Sub
		
		CUR_NO = N440
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX440_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX440.KeyDownEvent 'D-20250417
	Private Sub IMTX440_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX440.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX480_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX480.Enter
		
		'���t�f�[�^�Ȃ�΁A/�𔲂��ĕ\������
		If IsDate(IMTX480.Text) Then
			'UPGRADE_WARNING: DateValue �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
			IMTX480.Text = VB6.Format(DateValue(IMTX480.Text), "yyyymmdd")
		End If
		
		If CUR_NO = N480 Then Exit Sub
		
		CUR_NO = N480
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX480_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX480.KeyDownEvent 'D-20250417
	Private Sub IMTX480_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX480.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX490_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX490.Enter
		
		'���t�f�[�^�Ȃ�΁A/�𔲂��ĕ\������
		If IsDate(IMTX490.Text) Then
			'UPGRADE_WARNING: DateValue �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
			IMTX490.Text = VB6.Format(DateValue(IMTX490.Text), "yyyymmdd")
		End If
		
		If CUR_NO = N490 Then Exit Sub
		
		CUR_NO = N490
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX490_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX490.KeyDownEvent 'D-20250417
	Private Sub IMTX490_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX490.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub IMTX510_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX510.Enter
		
		Debug.Print("IMTX510_GotFocus")
		
		If CHK500.CheckState <> 1 And LST_NO = N500 Then
			CMDOFNC(12).Focus()
			Exit Sub
		End If
		
		'���t�f�[�^�Ȃ�΁A/�𔲂��ĕ\������
		If IsDate(IMTX510.Text) Then
			'UPGRADE_WARNING: DateValue �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
			IMTX510.Text = VB6.Format(DateValue(IMTX510.Text), "yyyymmdd")
		End If
		
		If CUR_NO = N510 Then Exit Sub
		
		CUR_NO = N510
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
	End Sub

	'Private Sub IMTX510_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX510.KeyDownEvent 'D-20250417
	Private Sub IMTX510_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As Control.AximText6.ITextEvents_KeyDownEvent) Handles IMTX510.KeyDownEvent 'A-20250417

		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub



	Private Sub imtxDummy_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imtxDummy.Enter

		'UPGRADE_WARNING: �I�u�W�F�N�g CTRLTBL(N300).CTRL.Index �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'Debug.Print("D" & CTRLTBL(N300).CTRL.Name & CTRLTBL(N300).CTRL.Index) 'D-20250417


		'    If CUR_NO <= N290 Then'D-CUST20130212
		If CUR_NO <= N291 Then 'A-cUST20130212
			'CTRLTBL(N300).CTRL.SetFocus    'D-20250201
			CTRLTBL(N310).CTRL.Focus() 'A-20250201
		Else
			'        CTRLTBL(N290).CTRL.SetFocus 'D-CUST20130212
			'A-CUST20130212��
			CTRLTBL(N291).CTRL.Focus()
			'A-CUST20130212��
		End If
		
		
		'   �ɓ�����͈ȉ��̂悤�ɂ�����B
		'�`�F�b�N�����w�p�̃_�~�[
		'    Call Form_KeyDown(vbKeyDown, 0)
		
	End Sub

	'Private Sub OPTO300_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO300.ClickEvent 'D-20250417
	Private Sub OPTO300_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO300.Click 'A-20250417
		Dim Index As Short = OPTO300.GetIndex(eventSender)

		Debug.Print("OPTO300" & Index & "Clicked")
		Call OPTO300_Enter(OPTO300.Item(Index), New System.EventArgs())

	End Sub

	Private Sub OPTO300_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO300.Enter
        Dim Index As Short = OPTO300.GetIndex(eventSender)
        '   �Ǘ��敪OptionButton

        Dim OptBefore As Short

        OptBefore = WKB300
        WKB300 = Index

        System.Diagnostics.Debug.Assert(WKB300 = Index, "")
        'Debug.Print("OPTO300" & OPTO300(Index).Value) 'D-20250417

        CTRLTBL(N300).CTRL.TabStop = True

        KB.kanri_kubn = "" & Index

        '--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
        If CUR_NO = N300 Then GoTo OPTO300_GotFocus_PropartySetting

        '--- ���Ă��������g�ɾ�� ---
        CUR_NO = N300


        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                WKB300 = OptBefore
                'OPTO300(WKB300).Value = True 'D-20250417
                OPTO300(WKB300).Checked = True 'A-20250417

                Exit Sub
            End If
            If GPROCHK() = False Then
                Exit Sub
            End If
        End If
        If GVALCHK() = False Then
            Exit Sub
        End If
        If MVALCHK() = False Then
            Exit Sub
        End If
        '�m��
        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()
OPTO300_GotFocus_PropartySetting:

        '   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
        CTRLTBL(N300).CTRL = Me.OPTO300(Index)
        NXT_NO = N300
        Call FOCUS_SET()

    End Sub

	'Private Sub OPTO300_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO300.KeyDownEvent 'D-20250417
	Private Sub OPTO300_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO300.KeyDown 'A-20250417
		Dim Index As Short = OPTO300.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N300
		WKB300 = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				WKB300 = WKB300 - 1
				If WKB300 < n1 Then WKB300 = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO300_Enter(OPTO300.Item(WKB300), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N300
			Case System.Windows.Forms.Keys.Right
				WKB300 = WKB300 + 1
				If WKB300 > n2 Then WKB300 = n2

				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO300_Enter(OPTO300.Item(WKB300), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N300
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub
	'A-CUST20130212���e�X�g�p
	Private Sub OPTO300_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO300.Leave
		Dim Index As Short = OPTO300.GetIndex(eventSender)
		Dim test As Object
		'UPGRADE_ISSUE: Control NAME �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
		'UPGRADE_WARNING: �I�u�W�F�N�g test �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		test = ActiveControl.Name
	End Sub
	'A-CUST20130212��
	'Private Sub OPTO310_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO310.ClickEvent 'D-20250417
	Private Sub OPTO310_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO310.Click 'A-20250417
		Dim Index As Short = OPTO310.GetIndex(eventSender)
		Call OPTO310_Enter(OPTO310.Item(Index), New System.EventArgs())

	End Sub

	Private Sub OPTO310_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO310.Enter
		Dim Index As Short = OPTO310.GetIndex(eventSender)
		'   �����OptionButton
		
		Dim OptBefore As Short
		
		'A-20190601��
		Dim w_Tax_kubn As New VB6.FixedLengthString(1)
		w_Tax_kubn.Value = KB.Tax_kubn
		'A-20190601��
		
		OptBefore = WKB310
		WKB310 = Index
		CTRLTBL(N310).CTRL.TabStop = True
		
		'--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
		KB.Tax_kubn = "" & Index
		'If CUR_NO = N310 Then GoTo OPTO310_GotFocus_PropartySetting
		
		'--- ���Ă��������g�ɾ�� ---
		CUR_NO = N310
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				WKB310 = OptBefore
				'OPTO310(WKB310).Value = True 'D-20250417
				OPTO310(WKB310).Checked = True 'A-20250417

				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		
		'A-20250201��
		If Index = 3 Then
			CMB370.SelectedIndex = 0
			CMB370.Enabled = False
		Else
			CMB370.Enabled = True
		End If
		'Call IPROCHK_N370
		'A-20250201��
		
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
OPTO310_GotFocus_PropartySetting: 
		
		'A-20190601��
		If w_Tax_kubn.Value <> KB.Tax_kubn Then
			Call SCR_DSPTAX()
		End If
		'A-20190601��
		
		'   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
		CTRLTBL(N310).CTRL = Me.OPTO310(Index)
		NXT_NO = N310
		Call FOCUS_SET()
		
	End Sub

	'Private Sub OPTO310_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO310.KeyDownEvent 'D-20250417
	Private Sub OPTO310_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO310.KeyDown 'A-20250417
		Dim Index As Short = OPTO310.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N310
		WKB310 = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				WKB310 = WKB310 - 1
				If WKB310 < n1 Then WKB310 = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO310_Enter(OPTO310.Item(WKB310), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N310
			Case System.Windows.Forms.Keys.Right
				WKB310 = WKB310 + 1
				If WKB310 > n3 Then WKB310 = n3
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO310_Enter(OPTO310.Item(WKB310), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N310
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub

	'Private Sub OPTO320_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO320.ClickEvent 'D-20250417
	Private Sub OPTO320_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO320.Click 'A-20250417
		Dim Index As Short = OPTO320.GetIndex(eventSender)
		Call OPTO320_Enter(OPTO320.Item(Index), New System.EventArgs())

	End Sub

	Private Sub OPTO320_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO320.Enter
		Dim Index As Short = OPTO320.GetIndex(eventSender)
		'   �I���P��OptionButton
		Dim OptBefore As Short
		
		OptBefore = WKB320
		WKB320 = Index
		CTRLTBL(N320).CTRL.TabStop = True
		
		KB.tana_tanka = "" & Index
		'--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
		If CUR_NO = N320 Then GoTo OPTO320_GotFocus_PropartySetting
		
		'--- ���Ă��������g�ɾ�� ---
		CUR_NO = N320
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				WKB320 = OptBefore
				'OPTO320(WKB320).Value = True 'D-0250417
				OPTO320(WKB320).Checked = True 'A-20250417
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
OPTO320_GotFocus_PropartySetting: 
		
		'   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
		CTRLTBL(N320).CTRL = Me.OPTO320(Index)
		NXT_NO = N320
		Call FOCUS_SET()
		
	End Sub

	'Private Sub OPTO320_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO320.KeyDownEvent 'D-20250417
	Private Sub OPTO320_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO320.KeyDown 'A-20250417
		Dim Index As Short = OPTO320.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N320
		WKB320 = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				WKB320 = WKB320 - 1
				If WKB320 < n1 Then WKB320 = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO320_Enter(OPTO320.Item(WKB320), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N320
			Case System.Windows.Forms.Keys.Right
				WKB320 = WKB320 + 1
				If WKB320 > n2 Then WKB320 = n2

				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO320_Enter(OPTO320.Item(WKB320), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N320
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	'Private Sub OPTO330_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO330.ClickEvent 'D-20250417
	Private Sub OPTO330_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO330.Click 'A-20250417
		Dim Index As Short = OPTO330.GetIndex(eventSender)
		Call OPTO330_Enter(OPTO330.Item(Index), New System.EventArgs())

	End Sub

	Private Sub OPTO330_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO330.Enter
		Dim Index As Short = OPTO330.GetIndex(eventSender)
		'   �݌ɊǗ�OptionButton
		
		Dim OptBefore As Short
		
		OptBefore = WKB330
		WKB330 = Index
		CTRLTBL(N330).CTRL.TabStop = True
		
		KB.zaiko = "" & Index
		'--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
		If CUR_NO = N330 Then GoTo OPTO330_GotFocus_PropartySetting
		
		'--- ���Ă��������g�ɾ�� ---
		CUR_NO = N330
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				WKB330 = OptBefore
				'OPTO330(WKB330).Value = True 'D-20250417
				OPTO330(WKB330).Checked = True 'A-20250417

				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
OPTO330_GotFocus_PropartySetting: 
		
		'   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
		CTRLTBL(N330).CTRL = Me.OPTO330(Index)
		NXT_NO = N330
		Call FOCUS_SET()
		
	End Sub

	'Private Sub OPTO330_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO330.KeyDownEvent 'D-20250417
	Private Sub OPTO330_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO330.KeyDown 'A-20250417
		Dim Index As Short = OPTO330.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N330
		WKB330 = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				WKB330 = WKB330 - 1
				If WKB330 < n1 Then WKB330 = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO330_Enter(OPTO330.Item(WKB330), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N330
			Case System.Windows.Forms.Keys.Right
				WKB330 = WKB330 + 1
				If WKB330 > n2 Then WKB330 = n2

				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO330_Enter(OPTO330.Item(WKB330), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N330
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	'Private Sub OPTO340_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO340.ClickEvent 'D-20250417
	Private Sub OPTO340_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO340.Click 'A-20250417
		Dim Index As Short = OPTO340.GetIndex(eventSender)
		Call OPTO340_Enter(OPTO340.Item(Index), New System.EventArgs())

	End Sub

	Private Sub OPTO340_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO340.Enter
		Dim Index As Short = OPTO340.GetIndex(eventSender)
		'   FAX���MOptionButton
		Dim OptBefore As Short
		
		OptBefore = WKB340
		WKB340 = Index
		CTRLTBL(N340).CTRL.TabStop = True
		
		KB.Fax_yn = "" & (Index - 1)
		Debug.Print("KB.Fax_tn=[" & KB.Fax_yn & "]")
		'--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
		If CUR_NO = N340 Then GoTo OPTO340_GotFocus_PropartySetting
		
		'--- ���Ă��������g�ɾ�� ---
		CUR_NO = N340
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				WKB340 = OptBefore
				'OPTO340(WKB340).Value = True 'D-20250417
				OPTO340(WKB340).Checked = True 'A-20250417
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
OPTO340_GotFocus_PropartySetting: 
		
		'   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
		CTRLTBL(N340).CTRL = Me.OPTO340(Index)
		NXT_NO = N340
		Call FOCUS_SET()
		
	End Sub

	'Private Sub OPTO340_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO340.KeyDownEvent 'D-20250417
	Private Sub OPTO340_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO340.KeyDown 'A-20250417
		Dim Index As Short = OPTO340.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N340
		WKB340 = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				WKB340 = WKB340 - 1
				If WKB340 < n1 Then WKB340 = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO340_Enter(OPTO340.Item(WKB340), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N340
			Case System.Windows.Forms.Keys.Right
				WKB340 = WKB340 + 1
				If WKB340 > n2 Then WKB340 = n2

				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call OPTO340_Enter(OPTO340.Item(WKB340), New System.EventArgs())
				'Call FOCUS_SET
				CUR_NO = N340
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

	End Sub

	Private Sub OPTO999_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OPTO999.Enter
		Dim Index As Short = OPTO999.GetIndex(eventSender)
		
		If bBackForm Then
			bBackForm = False
			Exit Sub
		End If
		
		'--- �����̏ꍇ�͏����𔲂���(ϳ��œ������ڂ�I�������ꍇ�Ȃ�) ---
		If CUR_NO = N999 Then GoTo OPTO999_SELF
		
		'--- ���Ă��������g�ɾ�� ---
		CUR_NO = N999
		'    If True Then
		'        LST_NO = CUR_NO
		'        Call FUNCSET_RTN
		'        Exit Sub
		'    End If
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				'OPTO999(KBKBN).Value = True
				OPTO999(KBKBN).Checked = True
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		'�m��
		
		KBKBN = Index
		CTRLTBL(N999).CTRL = Me.OPTO999(KBKBN)
		
		LST_NO = CUR_NO
		'--- �t�@���N�V�������b�Z�[�W
		Call FUNCSET_RTN()
		
OPTO999_SELF: 
		'   --- ���޸����߼�����݂̊m��l�Ƃ��Ď擾���� ---
		KBKBN = Index
		CTRLTBL(N999).CTRL = Me.OPTO999(KBKBN)
		NXT_NO = N999
		Call FOCUS_SET()
		'A-CUST-20100610 Start
		If Index = n1 Then
			IMTX030.TabStop = False
		Else
			IMTX030.TabStop = True
		End If
		'A-CUST-20100610 End
		
		'A-20250305-S
		'�폜�̂Ƃ��͉�ЁA���Ə��A�i�ԈȊO��Disable
		Select Case KBKBN
			Case 1 '   �ǉ�
				Call SetMode("A")
			Case 2 '   �C��
				Call SetMode("C")
			Case 3 '   �폜
				Call SetMode("D")
		End Select
		'A-20250305-E
		
	End Sub

	'Private Sub OPTO999_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskoptLibV5.__OSKOptBtn_KeyDownEvent) Handles OPTO999.KeyDownEvent 'D-20250417
	Private Sub OPTO999_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles OPTO999.KeyDown 'A-20250417
		Dim Index As Short = OPTO999.GetIndex(eventSender)
		'########## ���Ă��������g�ɾ�� ##########
		CUR_NO = N999
		KBKBN = Index

		'########## ���ق̍��E���݂ɂ���߼�����ݓ���̫������ړ����� ##########
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Left
				KBKBN = KBKBN - 1
				If KBKBN < n1 Then KBKBN = n1
				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call FOCUS_SET()
				CUR_NO = N999
			Case System.Windows.Forms.Keys.Right
				KBKBN = KBKBN + 1
				If KBKBN > n3 Then KBKBN = n3

				'========== ��޼ު�Ă�ݒ肵̫����ړ����� ==========
				Call FOCUS_SET()
				CUR_NO = N999
		End Select
		Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub



	'UPGRADE_ISSUE: PictureBox �C�x���g picDummy.KeyDown �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub picDummy_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		
		'    Call Form_KeyDown(KeyCode, Shift)
		
	End Sub


	'Private Sub SPR420_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SPR420.ClickEvent 'D-20250417
	Private Sub SPR420_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_ClickEvent) Handles SPR420.ClickEvent 'A-20250417

		Dim IROW As Integer

		If SPR420.MaxRows <= 0 Then Exit Sub
		If bSPRNotReady Then Exit Sub

		'�l�����͂���Ă���Ō�̃Z���̈ʒu+�P���擾����
		'    iRow = SPR420.DataRowCnt + 1
		''''Debug.Print "DataRowCnt = "; SPR420.DataRowCnt

		'�N���b�N�����Z�����Ō�̃Z���̈ʒu+�P���傫���ꍇ�A�Ō�̃Z���̈ʒu+�P���A�N�e�B�u�ɂ���
		'    If ROW > iRow Then
		'        SPR420.Col = 1
		'        SPR420.ROW = iRow
		'        SPR420.Col2 = 1
		'        SPR420.Row2 = iRow
		'        SPR420.Action = SS_ACTION_SELECT_BLOCK
		'        SPR420.Action = SS_ACTION_ACTIVE_CELL
		'    End If


	End Sub

    Private Sub SPR420_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPR420.Enter

        If CUR_NO = N420 Then Exit Sub

        CUR_NO = N420

        '�`�F�b�N
        If LST_NO <> n0 Then
            If IPROCHK() = False Then
                bSPRNotReady = True
                Exit Sub
            End If
            If GPROCHK() = False Then
                bSPRNotReady = True
                Exit Sub
            End If
        Else
            bSPRNotReady = True
            Exit Sub
        End If
        If GVALCHK() = False Then
            bSPRNotReady = True
            Exit Sub
        End If
        If MVALCHK() = False Then
            bSPRNotReady = True
            Exit Sub
        End If

        '�m��
        bSPRNotReady = False

        If SPR420.MaxRows <= 0 Then
            SPR420.MaxRows = 1
            SPR420.set_RowHeight(1, SPR_HEIGHT)
            Call SpreadProperty(1)
        End If


        ''''Call SpreadZeroTrim(1)
        If lst_row = 0 Then
            Call SpreadZeroTrim(1)
        ElseIf lst_row < 0 Then
            Call SpreadZeroTrim(-1)
        Else
            Call SpreadZeroTrim(lst_row)
        End If


        LST_NO = CUR_NO
        '--- �t�@���N�V�������b�Z�[�W
        Call FUNCSET_RTN()

    End Sub

	'Private Sub SPR420_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SPR420.KeyDownEvent 'D-20250417
	Private Sub SPR420_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_KeyDownEvent) Handles SPR420.KeyDownEvent 'A-20250417

		Dim ROW As Integer
		Dim Col As Integer

		Dim IROW As Short

		'F2,F3,F4�̏ꍇ�́A��������
		'    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF4 Then
		'        Call Form_KeyDown(KeyCode, Shift)
		'    '    KeyCode = 0
		''        Exit Sub
		'    End If

		SS_KEYCODE = eventArgs.KeyCode

		Dim iPrev As Short
		Select Case eventArgs.KeyCode

			'        Case vbKeyEscape    '   �I��
			'            Call CMDOFNC_Click(0)
			'            KeyCode = 0
			'            Exit Sub

			Case System.Windows.Forms.Keys.F3
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(3), New System.EventArgs())
				Col = 1
				ROW = SPR420.ActiveRow
				SPR420.Col = Col
				SPR420.Row = ROW
				SPR420.Col2 = Col
				SPR420.Row2 = ROW
				SPR420.Action = SS_ACTION_SELECT_BLOCK
				SPR420.Action = SS_ACTION_ACTIVE_CELL
				''''DoEvents
				eventArgs.KeyCode = 0
				Exit Sub

			Case System.Windows.Forms.Keys.F5
				''''        Call CMDOFNC_Click(5)
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
				IMTX030.Focus()
				System.Windows.Forms.Application.DoEvents()
				eventArgs.KeyCode = 0
				Exit Sub

			Case System.Windows.Forms.Keys.F12
				''''        Call CMDOFNC_Click(12)
				''''        KeyCode = 0
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
				System.Windows.Forms.Application.DoEvents()
				Exit Sub


				'Enter Key,��
			Case System.Windows.Forms.Keys.Return, System.Windows.Forms.Keys.Down '13, 40
				'Active Cell����̏ꍇ�́A���̍��ڂֈړ�����
				SPR420.Row = SPR420.ActiveRow
				SPR420.Col = 1

				If SPR420.Text = "" And SPR420.ActiveRow > SPR420.DataRowCnt Then
					eventArgs.KeyCode = 0
					NXT_NO = IIf(eventArgs.KeyCode = System.Windows.Forms.Keys.Return, CTRLTBL(N420).INEXT, CTRLTBL(N420).IDOWN)
					Call FOCUS_SET()
					Exit Sub
				Else
					'   2000/01/23  Add KOKOKARA
					If SPR420.MaxRows <= SPR420.DataRowCnt Then
						SPR420.MaxRows = SPR420.DataRowCnt + 1
						SPR420.Row = SPR420.DataRowCnt + 1
						''''SPR420.CellType = SS_CELL_TYPE_FLOAT
						SPR420.set_RowHeight(SPR420.Row, SPR_HEIGHT)
						Call SpreadProperty((SPR420.Row))

					End If
					'   2000/01/23  Add KOKOMADE
				End If

				'��
			Case System.Windows.Forms.Keys.Up '38
				'Active Cell���擪�s�Ŗ��m��̏ꍇ�́A�s���N���A���A�O�̍��ڂֈړ�����
				SPR420.Row = SPR420.ActiveRow
				SPR420.Col = 3
				If SPR420.Text <> "1" Then '2000/01/23 "1"->1
					IROW = SPR420.Row
					'�s���N���A����
					SPR420.Col = -1
					SPR420.Action = SS_ACTION_CLEAR_TEXT
				End If
				If SPR420.Row = 1 Then
					'�O�̍��ڂֈړ�����

					eventArgs.KeyCode = 0
					NXT_NO = CTRLTBL(N420).IBACK
					Call FOCUS_SET()
					''''CTRLTBL(iPrev).CTRL.SetFocus
				End If
				'��
			Case System.Windows.Forms.Keys.Right '39
				'��
			Case System.Windows.Forms.Keys.Left '37
				'F8�i�폜�j
			Case System.Windows.Forms.Keys.F8 '119

				'   2000/01/24          Add                     KOKOKARA
				If Trim(CMDOFNC(8).Text) = "" Then
					Exit Sub
				End If
				SPR420.Row = SPR420.ActiveRow
				SPR420.Col = 1
				If Trim(SPR420.Text) = "" And SPR420.Row > SPR420.DataRowCnt Then
					Exit Sub
				End If
				'   2000/01/24          Add                     KOKOMADE

				'�s�̍폜
				SPR420.Row = SPR420.ActiveRow
				SPR420.Action = SS_ACTION_DELETE_ROW

				IROW = SPR420.ActiveRow
				SPR420.Col = 1
				SPR420.Row = IROW
				SPR420.Col2 = 1
				SPR420.Row2 = IROW
				SPR420.Action = SS_ACTION_SELECT_BLOCK
				SPR420.Action = SS_ACTION_ACTIVE_CELL

				'   2000/01/23              Add             KOKOKARA
				If SPR420.MaxRows > 1 Then
					SPR420.MaxRows = SPR420.MaxRows - 1
				End If
				Call SpreadZeroTrim((SPR420.ActiveRow))
				'   2000/01/23              Add             KOKOMADE
				Call FUNCSET_RTN()

				'        SendKeys "{TAB}"
				'        SendKeys "{ESC}"
				'   Base1�z��SUB�����s
				'            NXT_NO = N420
				'            Call FOCUS_SET
				'           �}�E�X�ŃN���b�N���ꂽ�ꍇ�͂���ł悢���A
				'           F8KeyDown���痈���Ƃ���


				'���̑�
			Case Else
				Call SZ0410FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))

		End Select

		Call FUNCSET_RTN()

	End Sub

	'Private Sub SPR420_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SPR420.LeaveCell 'D-20240517
	Private Sub SPR420_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_LeaveCellEvent) Handles SPR420.LeaveCell 'A-20250417

		Dim strCode As String
		Dim strName As String
		Dim iReturn As Short

		'           2000/01/24      Add     KOKOKARA
		If bSPRNotReady Then
			Exit Sub
		End If
		'           2000/01/24      Add     KOKOMADE
		Debug.Print("NewROW=" & eventArgs.NewRow & "NewCol=" & eventArgs.NewCol)

		If eventArgs.NewRow = 0 Then
			eventArgs.Cancel = True
			Exit Sub
		End If


        '���͒l���擾����
        'SPR420.Col = Col 'D-20250417
        SPR420.Col = eventArgs.Col 'A-20250417
        SPR420.Row = ROW
		SPR420.Text = VB6.Format(SPR420.Text, "0000")
		strCode = SPR420.Text

		'    If Trim(strCode) = "" Then
		'
		'        Exit Sub
		'    End If

		'   �������݃`�F�b�N
		''''strName = DecodeBUSHO(strCode)
		strName = CduDecodeBUSHO(strCode)
		'   DUP CHECK�d���`�F�b�N
		iReturn = CHK_DUPFIND(strCode, ROW)
		If iReturn <> F_OFF Then
			strName = ""
		End If

		Dim strUnchanged As String
		Dim UnchangedName As String
		If strName = "" Or strName = "-" Then
			'   �G���[�̂Ƃ�

			ERRSW = F_ERR
			ENDSW = F_END

			''''''''If NewRow < SPR420.DataRowCnt Then      'ROW Then
			If eventArgs.NewRow < ROW Then 'ROW Then
				'   �������ރR�[�h�����Ƃ̒l�ɖ߂��B
				SPR420.Col = 4
				strUnchanged = SPR420.Text
				SPR420.Col = 1
				SPR420.Text = strUnchanged
				UnchangedName = CduDecodeBUSHO(strUnchanged)
				SPR420.Col = 2
				SPR420.Text = UnchangedName

				'   �m��t���O
				SPR420.Col = 3
				SPR420.Row = ROW
				SPR420.Text = IIf(Len(strUnchanged) > 0, "1", "")

				'   �O���ւ̈ړ��Ȃ狖���P�[�X
				If eventArgs.NewRow > 0 Then
					SPR420.Col = 1
					SPR420.Row = eventArgs.NewRow
					SPR420.Col2 = 1
					SPR420.Row2 = eventArgs.NewRow
					SPR420.Action = SS_ACTION_SELECT_BLOCK
					SPR420.Action = SS_ACTION_ACTIVE_CELL
					Call SpreadZeroTrim(eventArgs.NewRow)
				End If

			Else

				If strName = "-" Then
					ZAER_KN = n0
					ZAER_CD = 314
					ZAER_NO.Value = ""
					ZAER_MS.Value = ""
					strName = ""
					Call ZAER_SUB()
				End If

				SPR420.Col = 2
				SPR420.Text = strName
				SPR420.Col = 1
				SPR420.Text = strCode
				'   �t�H�[�J�X�����Ƃ̃Z���ɖ߂��B�ړ��������Ȃ��P�[�X
				SPR420.Col = 1 'Col
				SPR420.Row = ROW
				SPR420.Col2 = 1 'Col
				SPR420.Row2 = ROW
				SPR420.Action = SS_ACTION_SELECT_BLOCK
				SPR420.Action = SS_ACTION_ACTIVE_CELL
				Call SpreadZeroTrim(ROW)

			End If

			If SPR420.MaxRows > SPR420.DataRowCnt + 1 Then
				Debug.Print("ERR:" & SPR420.MaxRows & SPR420.DataRowCnt)
				'            SPR420.ROW = SPR420.MaxRows
				'            SPR420.Action = SS_ACTION_DELETE_ROW
				SPR420.MaxRows = SPR420.DataRowCnt + 1
			End If

			Exit Sub

		End If

		'�f�[�^���e���ڂɕ\������
		SPR420.Row = ROW

		SPR420.Col = 1
		SPR420.Text = strCode
		SPR420.Col = 2
		SPR420.Text = strName
		SPR420.Col = 4
		SPR420.Text = strCode

		'�m��t���O
		SPR420.Col = 3
		SPR420.Text = "1"

		lst_row = eventArgs.NewRow '   ���ꂪ�����ł���B          2000/01/18

		If eventArgs.NewCol <> 1 Then
			eventArgs.Cancel = True
		Else
			SpreadZeroTrim((eventArgs.NewRow))
		End If
		Call FUNCSET_RTN()


	End Sub



	Private Sub SpreadZeroTrim(ByRef lRow As Integer)
		
		Dim strCut As String
		
		If lRow = 0 Then Exit Sub
		
		If lRow < 0 Then
			Call SpreadZeroTrim((SPR420.ActiveRow))
			Exit Sub
		End If
		
		SPR420.ROW = lRow
		SPR420.Col = 1
		strCut = SPR420.Text
		strCut = ZeroTrim(strCut)
		SPR420.Text = strCut
		
	End Sub

	'Private Sub SPR420_TopLeftChange(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_TopLeftChangeEvent) Handles SPR420.TopLeftChange 'D-20250417
	Private Sub SPR420_TopChange(ByVal eventSender As System.Object, ByVal eventArgs As FarPoint.Win.Spread.TopChangeEventArgs) Handles SPR420.TopChange 'A-20250417

		Dim nRow As Integer '   �X�N���[���������Row
		Dim IROW As Integer '   �X�N���[��TopRow
		Dim aRow As Integer '   ����Row


		If ByMyself Then Exit Sub


		aRow = SPR420.ActiveRow
		nRow = SPR420.DataRowCnt + 1
		''''nRow = nRow - 4
		nRow = nRow - 2 '   2000/01/26 Fix
		nRow = IIf(nRow < 0, 0, nRow)

		IROW = SPR420.TopRow

		''''If iRow > (nRow - 4) Then
		If IROW > nRow Then
			ByMyself = True

			SPR420.Col = 1
			SPR420.Row = nRow
			SPR420.Col2 = 1
			SPR420.Row2 = nRow
			SPR420.Action = SS_ACTION_SELECT_BLOCK
			SPR420.Action = SS_ACTION_ACTIVE_CELL


			If aRow > nRow Then
				SPR420.Col = 1
				SPR420.Row = aRow
				SPR420.Col2 = 1
				SPR420.Row2 = aRow
				SPR420.Action = SS_ACTION_SELECT_BLOCK
				SPR420.Action = SS_ACTION_ACTIVE_CELL
			End If

			ByMyself = False
		End If

		'    Dim nRow As Long
		'    Dim iRow As Long
		'
		'    If ByMyself Then Exit Sub
		'
		'    nRow = SPR420.DataRowCnt + 1
		''    nRow = nRow - 2
		'    nRow = IIf(nRow < 0, 0, nRow)
		'
		'    iRow = SPR420.TopRow
		'
		'    If iRow > (nRow - 2) Then
		'        ByMyself = True
		'
		'        SPR420.Col = 1
		'        SPR420.ROW = nRow
		'        SPR420.Col2 = 1
		'        SPR420.Row2 = nRow
		'        SPR420.Action = SS_ACTION_SELECT_BLOCK
		'        SPR420.Action = SS_ACTION_ACTIVE_CELL
		'        ByMyself = False
		'    End If
		'

	End Sub

	Private Sub SPR420_LeftChange(ByVal eventSender As System.Object, ByVal eventArgs As FarPoint.Win.Spread.LeftChangeEventArgs) Handles SPR420.LeftChange 'A-20250417

		Dim nRow As Integer '   �X�N���[���������Row
		Dim IROW As Integer '   �X�N���[��TopRow
		Dim aRow As Integer '   ����Row


		If ByMyself Then Exit Sub


		aRow = SPR420.ActiveRow
		nRow = SPR420.DataRowCnt + 1
		''''nRow = nRow - 4
		nRow = nRow - 2 '   2000/01/26 Fix
		nRow = IIf(nRow < 0, 0, nRow)

		IROW = SPR420.TopRow

		''''If iRow > (nRow - 4) Then
		If IROW > nRow Then
			ByMyself = True

			SPR420.Col = 1
			SPR420.Row = nRow
			SPR420.Col2 = 1
			SPR420.Row2 = nRow
			SPR420.Action = SS_ACTION_SELECT_BLOCK
			SPR420.Action = SS_ACTION_ACTIVE_CELL


			If aRow > nRow Then
				SPR420.Col = 1
				SPR420.Row = aRow
				SPR420.Col2 = 1
				SPR420.Row2 = aRow
				SPR420.Action = SS_ACTION_SELECT_BLOCK
				SPR420.Action = SS_ACTION_ACTIVE_CELL
			End If

			ByMyself = False
		End If

		'    Dim nRow As Long
		'    Dim iRow As Long
		'
		'    If ByMyself Then Exit Sub
		'
		'    nRow = SPR420.DataRowCnt + 1
		''    nRow = nRow - 2
		'    nRow = IIf(nRow < 0, 0, nRow)
		'
		'    iRow = SPR420.TopRow
		'
		'    If iRow > (nRow - 2) Then
		'        ByMyself = True
		'
		'        SPR420.Col = 1
		'        SPR420.ROW = nRow
		'        SPR420.Col2 = 1
		'        SPR420.Row2 = nRow
		'        SPR420.Action = SS_ACTION_SELECT_BLOCK
		'        SPR420.Action = SS_ACTION_ACTIVE_CELL
		'        ByMyself = False
		'    End If
		'

	End Sub

	Private Function CHK_DUPFIND(ByRef strFind As String, ByRef lRow As Integer) As Short
		
		Dim lEnd As Integer
		Dim lx As Integer
		Dim iReturn As Short
		Dim saveRow, saveCol As Integer
		
		saveRow = SPR420.ROW
		saveCol = SPR420.Col
		iReturn = F_OFF
		
		lEnd = SPR420.DataRowCnt
		For lx = 1 To lEnd
			SPR420.ROW = lx
			SPR420.Col = 1
			If lx <> lRow And strFind = SPR420.Text Then
				iReturn = F_ERR
				Exit For
			End If
		Next lx
		
		SPR420.ROW = saveRow
		SPR420.Col = saveCol
		
		CHK_DUPFIND = iReturn
		
	End Function
	'
	'   2000/02/23  �d�l�ύX�ɂ�����
	Private Function CHK_DUPCOMBO(ByRef ix As Short, ByRef strTani As String) As Boolean
		
		CHK_DUPCOMBO = True
		If Trim(strTani) = "" Then
			Exit Function
		End If
		
		'   ��{�P�ʂƂ��Ȃ��Ȃ�G���[
		If Trim(strTani) = Trim(KB.tani) Then
			CHK_DUPCOMBO = False
			Exit Function
		End If
		
		'   ���Z�P�ʏd���`�F�b�N
		Dim i As Short
		
		For i = 1 To ix - 1
			If Trim(strTani) = Trim(CMB350(i).Text) Then
				CHK_DUPCOMBO = False
				Exit Function
			End If
			
		Next i
		
	End Function
	
	Private Sub SCR_INIT_RTN()
		
		WKB030 = "" '   �i��        12/26
		Call DBRollbackTrans()
		Call DBBeginTrans()
		Call SCR_ADDNEW()
		
		WKB140DSP = "" '   ���ޖ���
		WKB210DSP = "" '   ���ޖ���
		WKB220DSP = "" '   ���ޖ���
		WKB230DSP = "" '   ���ޖ���
		WKB240DSP = "" '   ���ޖ���
		WKB250DSP = "" '   ���ޖ���
		WKB260DSP = "" '   ���ޖ���
		WKB410DSP = "" '   �ƎҖ���
		WKB291DSP = "" '   JAN���i���ޖ��@A-CUST20130212
		WKAMOCHUNM = ""
		
		'A-20250201��
		CMB370.Items.Clear() '�R���{�{�b�N�X �N���A
		CMB370.Items.Add(New VB6.ListBoxItem("", 0)) '�o�^
		CMB370.Items.Add(New VB6.ListBoxItem("�W��", 1)) '�o�^
		CMB370.Items.Add(New VB6.ListBoxItem("�y��", 2)) '�o�^
		'A-20250201��
		
	End Sub
	'##################################################
	'##################################################
	'#####        ���s�����O �� �S���ڌQ������
	'##################################################
	'##################################################
	Private Function ALLCHK_RTN() As Short
		'���̓f�[�^�̍ă`�F�b�N������B
		Dim strCode1 As String
		Dim strCode2 As String
		Dim strCode3 As String
		Dim strName As String
		Dim WKB As SZM0010_S
		Dim iReturn As Short
		
		ALLCHK_RTN = -1
		ERRSW = F_OFF
		ENDSW = F_OFF
		
		ZAER_NO.Value = ""
		ZAER_KN = 0
		
		'   ��ЃR�[�h���݃`�F�b�N
		strCode1 = IMTX010.Text 'kb.Inc_code          '
		iReturn = CduDecodeKaisha(strCode1, strName)
		If iReturn = F_ERR Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N010
			Call FOCUS_SET()
			Exit Function
		End If
		
		'   ���Ə��R�[�h���݃`�F�b�N
		strCode2 = IMTX020.Text 'KB.jg_code          '
		iReturn = CduDecodeJigyo(strCode1, strCode2, strName)
		If iReturn = F_ERR Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N020
			Call FOCUS_SET()
			Exit Function
		End If
		
		'   �i�ԑ��݃`�F�b�N
		If KBKBN <> F_ADD Then 'A-CUST-20100610
			strCode3 = IMTX030.Text 'KB.hin_code          '
			iReturn = FILGET_SZM0010(strCode1, strCode2, strCode3, WKB)
			If iReturn = F_END Then
				'If KBKBN <> F_ADD Then                                     'D-CUST-20100610
				ZAER_CD = 120
				ZAER_NO.Value = "" 'A-CUST-20100610
				Call ZAER_SUB()
				NXT_NO = N030
				Call FOCUS_SET()
				Exit Function
				'End If                                                     'D-CUST-20100610
			Else
				'D-CUST-20100610 Start
				'If KBKBN = F_ADD Then
				'    ZAER_CD = 120
				'    Call ZAER_SUB
				'    NXT_NO = N030
				'    Call FOCUS_SET
				'    Exit Function
				'End If
				'D-CUST-20100610 End
			End If
		End If 'A-CUST-20100610
		
		'��ADD-2001/01/23 =========================================
		Dim Jisseki As Short
		
		If KBKBN = F_DEL Then '�폜Ӱ��
			'�i�ں��ނ̎��є���ı�ގ��s
			If PSZ0410SP_CALL_RTN(Jisseki) = False Then
				'�ı�ޓ��Ŵװ����
				NXT_NO = N030
				Call FOCUS_SET()
				Exit Function
			End If
			If Jisseki <> 0 Then '�g�p���т�����
				ZAER_CD = 3900 '���т�����̂ō폜�ł��܂���
				ZAER_NO.Value = "" 'A-CUST-20100610
				Call ZAER_SUB()
				NXT_NO = N030
				Call FOCUS_SET()
				Exit Function
			End If
			
		End If
		'��ADD-2001/01/23 =========================================
		
		'�P�ʃ`�F�b�N
		strCode1 = Trim(CMB060.Text) 'Trim(KB.tani)          '
		If strCode1 = "" Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N060
			Call FOCUS_SET()
			Exit Function
		End If
		
		'�K�p���@�`�F�b�N
		strCode1 = IMTX100(1).Text 'DateSlashed(KB.teki_date1)
		If Not IsDate(strCode1) Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N100_1
			Call FOCUS_SET()
			Exit Function
		End If
		
		'��p�ȖڂP�`�F�b�N
		strCode1 = IMTX130(1).Text 'KB.hiyou_k_code1
		strName = DecodeKAMOCHU(strCode1)
		If strName = "" Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N130
			Call FOCUS_SET()
			Exit Function
		End If
		
		'��p�ȖڂQ�`�F�b�N
		strCode2 = IMTX140(1).Text 'kb.hiyou_k_code2
		strName = DecodeKAMOKU(strCode1, strCode2)
		If strName = "" Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N140
			Call FOCUS_SET()
			Exit Function
		End If
		
		
		'A 050909 TOP NAGANO---------------------------------------START
		'��p�Ή��Ȗڃ`�F�b�N
		strName = FILGET_SZM0170(strCode1, strCode2)
		If strName = "" Then
			MsgBox("��p�Ή��Ȗ�Ͻ��ɑ��݂��܂���B", 48, "")
			NXT_NO = N130
			Call FOCUS_SET()
			Exit Function
		End If
		'A 050909 TOP NAGANO---------------------------------------END
		
		'�Ȗڕ��ރ`�F�b�N
		'D-20250201��
		'strCode1 = IMTX210.Text & IMTX211.Text  'KB.ka_bun_code
		'strName = DecodeKamBunrui(WKB010, WKB020, strCode1)
		'    strCode1 = Left(KB.ka_bun_code, 3)
		'    strCode2 = Left(KB.ka_bun_code, 4)
		'    strCode2 = Mid(KB.ka_bun_code, 4, 4)
		'    strName = DecodeKAMOKU(strCode1, strCode2)
		'If strName = "" Then
		'ZAER_CD = 120
		'ZAER_NO = ""                                                'A-CUST-20100610
		'Call ZAER_SUB
		'NXT_NO = N210
		'Call FOCUS_SET
		'Exit Function
		'End If
		'D-20250201��
		
		'�啪�ރ`�F�b�N
		strCode1 = IMTX220.Text 'KB.l_bun_code
		iReturn = CHK_BUNRUI(1, strCode1, "", "")
		If iReturn = F_ERR Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N220
			Call FOCUS_SET()
			Exit Function
		End If
		
		'�����ރ`�F�b�N
		strCode1 = IMTX230.Text 'KB.m_bun_code
		iReturn = CHK_BUNRUI(2, KB.l_bun_code, strCode1, "")
		If iReturn = F_ERR Then
			ZAER_NO.Value = "" 'A-CUST-20100610
			ZAER_CD = 120
			Call ZAER_SUB()
			NXT_NO = N230
			Call FOCUS_SET()
			Exit Function
		End If
		
		'�����ރ`�F�b�N
		strCode1 = IMTX240.Text 'KB.s_bun_code
		iReturn = CHK_BUNRUI(3, KB.l_bun_code, KB.m_bun_code, strCode1)
		If iReturn = F_ERR Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N240
			Call FOCUS_SET()
			Exit Function
		End If
		
		'02/05/28 ADD START
		' ���ރ`�F�b�N
		'D-20250201��
		'strCode1 = IMTX250.Text
		'iReturn = CHK_BUNRUI(4, strCode1, "", "")
		'If iReturn = F_ERR Then
		'ZAER_CD = 120
		'ZAER_NO = ""                                                'A-CUST-20100610
		'Call ZAER_SUB
		'NXT_NO = N250
		'Call FOCUS_SET
		'Exit Function
		'End If
		'D-20250201��
		
		'02/05/28 ADD END
		'�����P�ʇ@�`�F�b�N         ���̃`�F�b�N�p�~ 2000/02/23
		'    strCode1 = Trim(CMB350(1).Text)    'Trim(KB.ha_tanka1)
		'    If strCode1 = "" Then
		'        ZAER_CD = 120
		'        Call ZAER_SUB
		'        NXT_NO = N350_1
		'        Call FOCUS_SET
		'        Exit Function
		'    End If
		
		Dim i As Short
		
		'�����P�ʇ@�`�D�`�F�b�N
		For i = 1 To 5
			strCode1 = Trim(CMB350(i).Text)
			If strCode1 <> "" Then
				'   ���Z�P��DUP�`�F�b�N         2000/02/23  Add
				If Not CHK_DUPCOMBO(i, strCode1) Then
					ZAER_NO.Value = "" 'A-CUST-20100610
					ZAER_CD = 120
					Call ZAER_SUB()
					NXT_NO = IIf(i = 1, N350_1, IIf(i = 2, N350_2, IIf(i = 3, N350_3, IIf(i = 4, N350_4, N350_5))))
					Call FOCUS_SET()
					Exit Function
				End If '   2000/02/23  Add
				'   ���Z�P�ʎw�肠��Ƃ��͊��Z�����K�{
				If IMNU360(i).Value <= 0 Then
					ZAER_NO.Value = "" 'A-CUST-20100610
					ZAER_CD = 120
					Call ZAER_SUB()
					NXT_NO = IIf(i = 1, N360_1, IIf(i = 2, N360_2, IIf(i = 3, N360_3, IIf(i = 4, N360_4, N360_5))))
					Call FOCUS_SET()
					Exit Function
				End If
			End If
		Next i
        '����ŋ敪�`�F�b�N
        'D-20250201��
        'strCode1 = Trim(KB.Tax_kubn)
        'If strCode1 = "" And strCode1 >= "1" And strCode1 <= "5" Then
        'D-20250201��
        'A-20250201��
        strCode1 = Trim(KB.tax_rate_kbn)
		'If strCode1 = "3" And OPTO310(3).Value = False Then 'D-20250201
		If strCode1 = "3" And OPTO310(3).Checked = False Then 'A-20250417
			'A-20250201��
			ZAER_NO.Value = "" 'A-CUST-20100610
			ZAER_CD = 120
			Call ZAER_SUB()
			'NXT_NO = N440  'D-20250201
			NXT_NO = N370 'A-20250201
			Call FOCUS_SET()
			Exit Function
		End If

		If CHK500.CheckState = 1 And Trim(IMTX510.Text) = "" Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N500
			Call FOCUS_SET()
			Exit Function
		End If
		If CHK500.CheckState <> 1 And Trim(IMTX510.Text) <> "" Then
			ZAER_CD = 120
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			NXT_NO = N500
			Call FOCUS_SET()
			Exit Function
		End If
		
		'    Call IPROCHK_N510
		'    If ERRSW = F_ERR Then
		'        NXT_NO = LST_NO
		'        Call FOCUS_SET
		'        Exit Function
		'    End If
		
		
		'   �ȖڑΉ��e�[�u���Ƃ̓ˍ���
		Call GPROCHK_GRP7()
		If ERRSW = F_ERR And False Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Function
		End If
		If ERRSW = F_ERR Then
			ERRSW = F_OFF
		End If
		If ENDSW = F_END Then
			ENDSW = F_OFF
		End If
		
		'A-CUST-20100610 Start
		Dim nnum As Integer
		If KBKBN = F_ADD Then
			nnum = New_Number
			If nnum < 0 Or nnum > 99999 Then
				Call MsgBox("�����̔Ԃ�����ɒB���܂����B�@", MsgBoxStyle.ApplicationModal + MsgBoxStyle.Exclamation, "�d���i�ڊ�{������")
				IMTX030.Text = ""
				NXT_NO = LST_NO
				Call FOCUS_SET()
			Else
				KB.hin_code = VB6.Format(nnum, "00000")
				WKB030 = KB.hin_code
				IMTX030.Text = KB.hin_code
			End If
		End If
		'A-CUST-20100610 End
		
		'    Call GPROCHK_GRP8
		'    If ERRSW = F_ERR Then
		'        NXT_NO = LST_NO
		'        Call FOCUS_SET
		'        Exit Function
		'    End If
		
		'A-20240115��
		'�L�������敪�`�F�b�N
		If CMB165.SelectedIndex <> 0 Then
			If CMB170.SelectedIndex = 0 Or CMB170.SelectedIndex = -1 Then
				ZAER_CD = 120
				ZAER_NO.Value = ""
				Call ZAER_SUB()
				NXT_NO = N170CMB
				Call FOCUS_SET()
				Exit Function
			End If
		End If
		'�L�������`�F�b�N
		If Val(CStr(IMNU170(1).Value)) = 0 And CDbl(RTrim(CStr(CMB165.SelectedIndex))) <> 0 Then
			ZAER_CD = 120
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			NXT_NO = N170
			Call FOCUS_SET()
			Exit Function
		End If
		'A-20240115��
		
		'A-20250303��
		'JAN�R�[�h�d���`�F�b�N
		Dim chk_jan_hincode As String
		If RTrim(KB.jan_code) <> "" Then
			chk_jan_hincode = CHK_JANCODE(KB.jan_code)
			If chk_jan_hincode <> "" Then
				Call MsgBox("���̕i�Ԃœ���JAN�W�����ނ��g�p����Ă��܂��B" & vbCrLf & "�i��[" & chk_jan_hincode & "]", MsgBoxStyle.ApplicationModal + MsgBoxStyle.Exclamation, "�d���i�ڊ�{������")
				NXT_NO = N070
				Call FOCUS_SET()
				Exit Function
			End If
		End If
		'A-20250303��
		
		ALLCHK_RTN = 0
		
	End Function
	
	Private Sub TAB010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TAB010.Enter
		
		If KBKBN = 1 Or KBKBN = 2 Then
			If TAB010.SelectedIndex = 0 Then
				NXT_NO = N100_1
			ElseIf TAB010.SelectedIndex = 1 Then 
				'NXT_NO = N210  'D-20250201
				NXT_NO = N220 'A-20250201
			Else
				NXT_NO = N410
			End If
		Else
			NXT_NO = N030
		End If
		Call FOCUS_SET()
		
	End Sub
	
	Function New_Number() As Integer
		Dim stSql As String
		Dim nnum As String
		
		If KBKBN <> F_ADD Then Exit Function
		
		New_Number = -1
		
		qSZM0010_NSEL.rdoParameters("Inc_code").Value = IMTX010.Text
		qSZM0010_NSEL.rdoParameters("jg_code").Value = IMTX020.Text
		qSZM0010_NSEL.rdoParameters("Inc_code2").Value = IMTX010.Text 'A-CUST-20100610
		qSZM0010_NSEL.rdoParameters("jg_code2").Value = IMTX020.Text 'A-CUST-20100610
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0010_NmyRSSW <> "qSZM0010_NSEL" Or ReQue = False Then
			SZM0010_NmyRS = qSZM0010_NSEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0010_NmyRSSW = "qSZM0010_NSEL"
		Else
			SZM0010_NmyRS.Requery()
		End If
		
		Select Case B_STATUS(SZM0010_NmyRS) ' (SQL���s�ð���̕]��)
			Case 0
				
				nnum = SZM0010_NmyRS.rdoColumns("maxnum").Value
				New_Number = Val(nnum) + 1
			Case 24
				ENDSW = F_END
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
			Case -54 '   ���b�N
				ZAER_CD = 201
				ZAER_NO.Value = "" 'A-CUST-20100610
				Call ZAER_SUB()
				ENDSW = F_END
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
				
			Case Else
				ENDSW = F_END
				ERRSW = F_ERR
				ZAER_KN = 1
				ZAER_NO.Value = "RSZM0010"
				Call ZAER_SUB()
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'A-CUST-20100610 Start
	Public Sub DSP_SENTAKU()
		
		'A-CUST-20100823 Start
		IMTX070.Text = RTrim(KB.jan_code)
		IMTX080.Text = RTrim(KB.jan_s_code)
		IMTX090.Text = RTrim(KB.bar_code)
		'A-CUST-20100823 End
		IMTX040.Text = RTrim(KB.hin_name)
		IMTX050.Text = RTrim(KB.kikaku)
		'UPGRADE_ISSUE: ComboBox �v���p�e�B CMB060.DataField �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
		'CMB060.DataField = KB.tani 'D-20250417
		CMB060.DataSource = KB.tani 'A-20250417
		Call COMBO_SETLIST(CMB060, KB.tani)
		IMTX065.Text = RTrim(KB.hin_name_seisiki)
		IMNU120(1).Value = KB.kei_kin1
		'A-CUST-20100823 Start
		IMTX100(1).Text = DateSlashed(KB.teki_date1)
		Call COMBO_SETLIST(CMB350(1), KB.ha_tanka1)
		IMNU360(1).Value = KB.kansan_num1
		'A-CUST-20100823 End
		'A-CUST20130222��
		KB.jan_code = IMTX070.Text
		JAN_BUF0.k4 = IMTX070.Text
		If FILGET_JAN() = False Then
		Else
			KB.BK1 = JAN.k21
			KB.k44 = JAN.k44
			KB.k42 = JAN.k42
			KB.k57 = JAN.k57
			KB.k58 = JAN.k58
			IMTX150(0).Text = KB.k44
			IMNU160(0).Value = KB.k42
			IMNU170(1).Value = KB.k58
			IMTX291.Text = KB.BK1
			'D-20130424-S
			'        If Trim(JAN.k14) <> "" Then
			'            KB.hin_name_seisiki = JAN.k14
			'            IMTX065.Text = KB.hin_name_seisiki
			'        End If
			'D-20130424-E
			'A-20130424-S
			If Trim(JAN.k17) <> "" Then
				KB.hin_name_seisiki = JAN.k17
				IMTX065.Text = KB.hin_name_seisiki
			End If
			'A-20130424-E
			
			'A-20240115��
			Select Case KB.Shomi_date_kbn
				Case CStr(0)
					CMB165.Text = "�����Ȃ�"
				Case CStr(1)
					CMB165.Text = "�������"
				Case CStr(2)
					CMB165.Text = "�ܖ�����"
				Case Else
					CMB165.SelectedIndex = -1
			End Select
			'A-20240115��
			
			'���t���Z
			KB.k99 = 0 '�v�Z�O�ɃN���A
			DSP170(0).Text = CStr(0) '�v�Z�O�ɃN���A
			If RTrim(KB.k57) = "" Then
				CMB170.SelectedIndex = -1
			Else
				Select Case KB.k57
					Case CStr(1)
						CMB170.Text = "��"
					Case CStr(2)
						CMB170.Text = "��"
					Case CStr(3)
						CMB170.Text = "�N"
					Case Else
						CMB170.SelectedIndex = -1
				End Select
			End If
			Call CNV_DAY() '�����Z����
			'JAN���i���ޖ��擾
			DSP291.Text = "" '�N���A
			JAN_BUNRUI_BUF0.BK1 = KB.BK1
			If FILGET_JAN_BUNRUI() = True Then
				DSP291.Text = JAN_BUNRUI.BK4 '���ޖ�
			End If
		End If
		'A-CUST20130222��
	End Sub
	'A-CUST-20100610 End
	'A-CUST20130212��
	'�������Z����
	Public Sub CNV_DAY()
		If CMB170.SelectedIndex = -1 Then Exit Sub
		'�����Z
		Select Case VB6.GetItemData(CMB170, CMB170.SelectedIndex)
			Case 1 '���̏ꍇ
				DSP170(0).Text = CStr(IMNU170(1).Value)
				KB.k99 = CDec(DSP170(0).Text)
			Case 2 '���̏ꍇ
				'DSP170(0).Caption = Fix(Val(IMNU170(1).Value) * 30.5)  'D-20130227-
				DSP170(0).Text = CStr(Fix((Val(CStr(IMNU170(1).Value)) * 30.416) + 0.5)) 'A-20130227-
				KB.k99 = CDec(DSP170(0).Text)
			Case 3 '�N�̏ꍇ
				DSP170(0).Text = CStr(Val(CStr(IMNU170(1).Value)) * 365)
				KB.k99 = CDec(DSP170(0).Text)
		End Select
	End Sub
	'A-CUST20130212��
End Class