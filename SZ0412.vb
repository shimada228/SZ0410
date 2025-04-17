Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class SZ0412FRM
	Inherits System.Windows.Forms.Form
	'A-CUST-20100610 �t�H�[���ǉ� '�C���f�b�N�X�̍ŏ��l���P�ɐݒ�
	
	Dim CUR_NO As Short '�����͈ʒu���۰ه�
	Dim LST_NO As Short '�O���͈ʒu���۰ه�
	Dim NXT_NO As Short '�����͈ʒu���۰ه�
	
	'����(�X�v���b�h)�p���
	Dim PRE_VALUE As Object '���בO�l
	Dim ViewCol As Integer '���׌���
	Dim SPRD_ERR As Short '�X�v���b�h���͂n�j�t���O
	
	'------------------------------------------------------------------
	'         ��ʺ��۰ٍ��ڐݒ�
	'------------------------------------------------------------------
	Const N005 As Short = 1 '����
	Const NEND As Short = 2 '
	'UPGRADE_WARNING: �z�� CTRLTBL �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Dim CTRLTBL(NEND) As CTRLTBL_S '��ʺ��۰ٔz��
	
	Const GRP1 As Short = 1
	Const GEND As Short = 2
	'UPGRADE_WARNING: �z�� GRPTBL �̉����� 0 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Dim GRPTBL(GEND) As GRPTBL_S '��ʸ�ٰ�ߔz��
	
	Enum SPRD_COL
		col_SEN = 1
		col_HIN_NAME
		col_KIKAKU
		col_GYO_NAME
		col_TANI
		col_TANKA
		'A-CUST-20100823 Start
		col_TEKI_DATE
		col_HA_TANI
		col_KANSANSU
		col_JAN_CODE
		col_JAN_S_CODE
		col_BAR_CODE
		'A-CUST-20100823 End
		col_RENBAN '��\��
	End Enum
	
	Private Sentakurow As Integer
	Private ButtnFlg As Boolean
	
	Private Function ALLCHK_UPD() As Boolean
		Dim ii As Integer
		Dim SvRow As Integer
		
		SvRow = SPRD050.ActiveRow
		With SPRD050
			
			Sentakurow = 0
			'�S�����F�̃`�F�b�N
			For ii = 1 To .MaxRows
				.ROW = ii
				.Col = 1
				
				If CBool(.Value) = True Then
					If Sentakurow = 0 Then
						Sentakurow = ii
					Else
						Sentakurow = 0
						Exit For
					End If
				End If
			Next ii
			
			If Sentakurow = 0 Then
				ALLCHK_UPD = False
				Exit Function
			End If
			
			SPRD050.ROW = SvRow
			
		End With
		
		ALLCHK_UPD = True
		
	End Function
	
	Public Sub INITIAL_RTN()
		'��ʍ��ڏ����l�ݒ�
		DSP010.Text = RTrim(WKB010)
		DSP020.Text = RTrim(SZ0410FRM.DSP010.Text)
		DSP030.Text = RTrim(WKB020)
		DSP040.Text = RTrim(SZ0410FRM.DSP020.Text)
		
		SPRD050.MaxRows = 0
		
	End Sub
	
	
	'******************************************************************
	'*      ��ʺ��۰ُ����ݒ�                                (TBL_SET)
	'******************************************************************
	Sub TBL_SET()
		CTRLTBL(N005).IGRP = GRP1 '���׃O���[�v
		CTRLTBL(NEND).IGRP = GEND
		
SET_NO: 
		'------------------------------------------------------------
		'   �����ځA�O���ڂ̐ݒ�
		'------------------------------------------------------------
		
		CTRLTBL(N005).INEXT = NEND '����
		CTRLTBL(N005).IBACK = n0
		CTRLTBL(N005).IDOWN = NEND
		
		CTRLTBL(NEND).INEXT = n0
		CTRLTBL(NEND).IBACK = N005
		CTRLTBL(NEND).IDOWN = n0
		
		'------------------------------------------------------------
		'   ���۰ٕۑ�
		'------------------------------------------------------------
		CTRLTBL(N005).CTRL = SPRD050 '����
		CTRLTBL(NEND).CTRL = CMDOFNC(12) '���s�{�^��
		
	End Sub
	
	'******************************************************************
	'*      �t�@���N�V�����E�{�^���iGotFocus�j
	'******************************************************************
	Private Sub CMDOFNC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Enter
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		If MOUSEFLG = 0 Then MOUSEFLG = VB6.MouseButtonConstants.LeftButton
		
		If Index <> 12 Then Exit Sub
		
		If CUR_NO = NEND Then Exit Sub '  �����Ȃ牽�����Ȃ�
		CUR_NO = NEND '  ���̌��݂̈ʒu��ݒ肷��
		
		If LST_NO <> n0 Then '�y�`�F�b�N�z
			If IPROCHK() = False Then '  LostFocus��������
				Exit Sub
			End If
			If GPROCHK() = False Then '  LostFocus���ڌQ����
			End If
		End If
		If GVALCHK() = False Then '  GotFocus���ڌQ����
			Exit Sub
		End If
		If MVALCHK() = False Then '  GotFocus��������
			Exit Sub
		End If
		
		LST_NO = CUR_NO '�y��ʺ��۰ه��m��z
		
		Call FUNCSET_RTN() '�y�t�@���N�V�����K�C�h�\���z
		
	End Sub
	
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		'���s�ȊO�͖���
		If Index < 12 Then
			Exit Sub
		End If
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				
				NXT_NO = N005
				Call FOCUS_SET()
				Exit Sub
				
		End Select
		
		Call SZ0412FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	'******************************************************************
	'*      �t�@���N�V�����E�{�^���iMouseDown�j
	'******************************************************************
	Private Sub CMDOFNC_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDOFNC.MouseDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		MOUSEFLG = eventArgs.Button
	End Sub
	
	'******************************************************************
	'*      �t�@���N�V�����E�{�^���iClick�j
	'******************************************************************
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		Dim i As Integer
		
		If MOUSEFLG <> VB6.MouseButtonConstants.LeftButton And MOUSEFLG <> 0 Then
			MOUSEFLG = 0
			Exit Sub
		End If
		
		MOUSEFLG = 0
		
		If CMDOFNC(Index).Text = "" Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Sub
		End If
		
		Dim IROWSelect As Short
		Select Case Index
			Case 0 '�y�I���z
				'UPGRADE_NOTE: EditMode �� CtlEditMode �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
				SPRD050.CtlEditMode = False
				IMTXDUM.Focus()
				
			Case 5 '�y�N���A�z
				With SPRD050
					If .MaxRows > 0 Then
						For IROWSelect = 1 To .MaxRows
							.ROW = IROWSelect
							.Col = 1
							.Value = CStr(False)
						Next 
					End If
					
					.Focus()
				End With
				
				'A-20110621-S
			Case 9 '�y�폜�z
				If ALLCHK_UPD() = False Then
					ZAER_KN = n0
					ZAER_CD = 120 '"���͓��e�Ɍ�肪����܂��E�E�E"
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					DUMMY.Focus()
					Exit Sub
				End If
				
				If MsgBox("�폜���s���܂��B��낵���ł����H", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then
					NXT_NO = N005
					Call FOCUS_SET()
					Exit Sub
				End If
				
				Call BEGIN_RTN()
				SPRD050.ROW = Sentakurow
				SPRD050.Col = SPRD_COL.col_RENBAN
				RENBAN_SEN = CInt(SPRD050.Value)
				Call GO_WKDELETE() '�폜
				If ERRSW = F_ERR Then
					Call ROLLBACK_RTN()
				Else
					Call COMMIT_RTN()
				End If
				
				ButtnFlg = False
				Call SET_SPRD050_UPD() '�ĕ\��
				ButtnFlg = True
				'A-20110621-E
				
			Case 12 '�y���s�z
				If ALLCHK_UPD() = False Then
					ZAER_KN = n0
					ZAER_CD = 120 '"���͓��e�Ɍ�肪����܂��E�E�E"
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					DUMMY.Focus()
					Exit Sub
				End If
				
				'�`�F�b�N�n�j
				If MsgBox("�i�ڃf�[�^�̎捞�݂��s���܂��B��낵���ł����H", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then
					NXT_NO = N005
					Call FOCUS_SET()
					Exit Sub
				End If
				
				SPRD050.ROW = Sentakurow
				SPRD050.Col = SPRD_COL.col_HIN_NAME
				KB.hin_name_seisiki = SPRD050.Value
				'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				KB.hin_name = MKKCMN.ZACHGSTR_SUB(KB.hin_name_seisiki, Len(KB.hin_name))
				SPRD050.Col = SPRD_COL.col_KIKAKU
				KB.kikaku = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_TANI
				KB.tani = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_TANKA
				KB.kei_kin1 = CDec(SPRD050.Value)
				'A-CUST-20100823 Start
				SPRD050.Col = SPRD_COL.col_TEKI_DATE
				If RTrim(SPRD050.Value) = "" Then
					KB.teki_date1 = ""
				Else
					KB.teki_date1 = VB.Left(SPRD050.Value, 4) & Mid(SPRD050.Value, 6, 2) & Mid(SPRD050.Value, 9, 2)
				End If
				SPRD050.Col = SPRD_COL.col_HA_TANI
				KB.ha_tanka1 = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_KANSANSU
				KB.kansan_num1 = CDec(SPRD050.Value)
				SPRD050.Col = SPRD_COL.col_JAN_CODE
				KB.jan_code = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_JAN_S_CODE
				KB.jan_s_code = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_BAR_CODE
				KB.bar_code = SPRD050.Value
				'A-CUST-20100823 End
				
				SPRD050.Col = SPRD_COL.col_RENBAN
				RENBAN_SEN = CInt(SPRD050.Value)
				
				SentakuFLG = True
				
				Call SZ0410FRM.DSP_SENTAKU()
				
				IMTXDUM.Focus()
				
		End Select
	End Sub
	
	'UPGRADE_ISSUE: PictureBox �C�x���g DUMMY.GotFocus �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub DUMMY_GotFocus()
		SPRD050.Focus()
		
	End Sub
	
	'UPGRADE_WARNING: Form �C�x���g SZ0412FRM.Activate �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
	Private Sub SZ0412FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'�}�E�X�J�[�\����߂�
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		'�E�C���h�E�\���ʒu�ݒ�T�u���[�`��
		Call ZAWC_SUB(Me, 0)
		Me.Top = 0
		Me.Left = 0
		
		'�I�y���[�^���\���T�u���[�`��
		Call ZAOP_SUB(Me, WKB010, WG_OPCODE)
		If ERRSW = F_ERR Then
			Call ENDR_RTN()
		End If
		
		Call INITIAL_RTN() '������ʕ\��
		
		'�N�������`�F�b�N
		Dim lRet As Integer
		Dim OP_KENGEN As Integer
		
		'lRet = MKKDBCMN.MKKDBCMN_SQTGET1_SUB(3, "SZ0410", WKB010, WG_OPCODE, OP_KENGEN)        'D-CUST-20100901
		lRet = MKKDBCMN.MKKDBCMN_SQTGET1_SUB(3, "SZ0412", WKB010, WG_OPCODE, OP_KENGEN) 'A-CUST-20100901
		If lRet <> n0 Then
			IMTXDUM.Focus()
			Exit Sub
		End If
		If OP_KENGEN = 0 Then
			ZAER_KN = n0
			ZAER_CD = 301
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			IMTXDUM.Focus()
		End If
		
		'    '�X�V�����`�F�b�N
		'    '�X�V����
		'    lRet = MKKDBCMN.MKKDBCMN_SQTGET3_SUB(4, "SZ0410", WKB010, WKB020, "", WG_OPCODE, OP_KENGEN)
		'    If lRet <> n0 Then
		'        IMTXDUM.SetFocus
		'        Exit Sub
		'    End If
		'    '�X�V�����Ȃ�
		'    If OP_KENGEN = 0 Then
		'        ZAER_KN = 0
		'        ZAER_CD = 303
		'        ZAER_NO = ""
		'        Call ZAER_SUB
		'        IMTXDUM.SetFocus
		'        Exit Sub
		'    End If
		
		ButtnFlg = False
		Call SET_SPRD050_UPD()
		ButtnFlg = True
		
		CUR_NO = n0
		LST_NO = N005
		NXT_NO = N005
		Call FOCUS_SET()
		
	End Sub
	
	'******************************************************************
	'*      �e�n�q�l�i�j�����c�������j
	'******************************************************************
	Private Sub SZ0412FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'�e�R���g���[���̋��ʂ̃L�[������s��
		'�ŗL�̃L�[����͊e�R���g���[����KeyDown�C�x���g�ōs��
		
		If Me.Enabled = False Then
			KeyCode = n0
			Exit Sub
		End If
		If Shift <> n0 Then 'Shift,Ctrl,Graph(Alt)�L�[�������A��������
			Exit Sub
		End If
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape '�yESC�z
				KeyCode = n0
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
				End If
			Case System.Windows.Forms.Keys.Return '�y���ݷ��z
				Call SET_NO(1)
				KeyCode = n0
			Case System.Windows.Forms.Keys.Up '�y�����z
				Call SET_NO(2)
				KeyCode = n0
			Case System.Windows.Forms.Keys.Down '�y�����z
				Call SET_NO(3)
				KeyCode = n0
			Case System.Windows.Forms.Keys.F1 '�y�e�P�z
			Case System.Windows.Forms.Keys.F2 '�y�e�Q�z
				If CMDOFNC(2).Text <> "" Then
					CMDOFNC(2).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(2), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F3 '�y�e�R�z
				If CMDOFNC(3).Text <> "" Then
					CMDOFNC(3).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(3), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F4 '�y�e�S�z
				If CMDOFNC(4).Text <> "" Then
					CMDOFNC(4).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(4), New System.EventArgs())
					CTRLTBL(CUR_NO).CTRL.Focus()
				End If
				'KeyCode = n0
			Case System.Windows.Forms.Keys.F5 '�y�e�T�z
				If CMDOFNC(5).Text <> "" Then
					CMDOFNC(5).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F6 '�y�e�U�z
				If CMDOFNC(6).Text <> "" Then
					CMDOFNC(6).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F7 '�y�e�V�z
				If CMDOFNC(7).Text <> "" Then
					CMDOFNC(7).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F8 '�y�e�W�z
				If CMDOFNC(8).Text <> "" Then
					CMDOFNC(8).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(8), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F9 '�y�e�X�z
				If CMDOFNC(9).Text <> "" Then
					CMDOFNC(9).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(9), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F10 '�y�e10�z
				KeyCode = n0
			Case System.Windows.Forms.Keys.F11 '�y�e11�z
				If CMDOFNC(11).Text <> "" Then
					CMDOFNC(11).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(11), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F12 '�y�e12�z
				KeyCode = n0
				If CMDOFNC(12).Text <> "" Then
					CMDOFNC(12).Focus()
					System.Windows.Forms.Application.DoEvents()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				End If
		End Select
		
	End Sub
	
	'******************************************************************
	'*      �e�n�q�l�i�k�n�`�c�j
	'******************************************************************
	Private Sub SZ0412FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: Form �v���p�e�B SZ0412FRM.HelpContextID �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
		Me.HelpContextID = SM_HelpContextID
		
		Call TBL_SET() '��ʺ��۰ُ����ݒ�
		
	End Sub
	
	'******************************************************************
	'*      �e�n�q�l�i�p���������t�����������j
	'******************************************************************
	Private Sub SZ0412FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'******************************************************************
	'*      �t�H�[�J�X�Z�b�g                                (FOCUS_SET)
	'******************************************************************
	Private Sub FOCUS_SET()
		
		If NXT_NO <= 0 Then Exit Sub
		
		Select Case NXT_NO
			Case N005
				SPRD050.Focus()
				SPRD050.Col = 1
				SPRD050.Action = 0
				
			Case NEND
				CType(Me.Controls("CMDOFNC"), Object)(12).Enabled = True
				CType(Me.Controls("LBLFNC"), Object)(12).Enabled = True
				CMDOFNC(12).Focus()
				
			Case Else
				CTRLTBL(NXT_NO).CTRL.Focus()
				
		End Select
		
	End Sub
	
	'******************************************************************
	'*      �e�t�m�b�s�h�n�m�Z�b�g                            (FUNCSET)
	'******************************************************************
	Private Sub FUNCSET_RTN()
		
		'--- �t�@���N�V�����E�K�C�h���b�Z�[�W
		Select Case LST_NO
			
			Case N005 '����
				With SPRD050
					Select Case ViewCol
						Case 1
							ZAFC_N(0) = 1
							ZAFC_N(5) = 5
							ZAFC_N(9) = 8 'A-20110621-
							ZAFC_N(12) = 12
					End Select
				End With
				
			Case NEND
				ZAFC_N(12) = 12
				
		End Select
		
		'�t�@���N�V�������b�Z�[�W
		Call ZAFC_SUB(Me)
		
		'�K�C�h���b�Z�[�W
		Call ZAGD_SUB(Me)
	End Sub
	
	'******************************************************************
	'*      ��ʺ��۰ه��Z�b�g                                 (SET_NO)
	'******************************************************************
	Sub SET_NO(ByRef FUNC As Short)
		
		Select Case FUNC
			Case 1 ' ������
				
				NXT_NO = CTRLTBL(LST_NO).INEXT
				Call FOCUS_SET()
			Case 2 ' �O����
				If CTRLTBL(LST_NO).IBACK <> 0 Then
					NXT_NO = CTRLTBL(LST_NO).IBACK
					Call FOCUS_SET()
				End If
			Case 3 ' ���O���[�v
				NXT_NO = CTRLTBL(LST_NO).INEXT
				Call FOCUS_SET()
		End Select
		
	End Sub
	
	'******************************************************************
	'*      �O���[�v�`�F�b�N�iLostFocus���ڌQ�`�F�b�N�j       (GPROCHK)
	'******************************************************************
	Function GPROCHK() As Short
		GPROCHK = True
		
		ERRSW = F_OFF
		ENDSW = F_OFF
		
		If CTRLTBL(CUR_NO).IGRP <= CTRLTBL(LST_NO).IGRP Then Exit Function
		
		Select Case CTRLTBL(LST_NO).IGRP
			Case GRP1
				Call GPROCHK_GRP1()
		End Select
		
		
		If ERRSW = F_ERR Then '�G���[
			GPROCHK = False
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = False
			Select Case CTRLTBL(LST_NO).IGRP
				Case GRP1
					NXT_NO = GRPTBL(GRP1).NXTN
			End Select
			Call FOCUS_SET()
		Else
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = True
		End If
		
	End Function
	
	'******************************************************************
	'*      �O���[�v�`�F�b�N�i�P�j                       (GPROCHK_GRP1)
	'******************************************************************
	Sub GPROCHK_GRP1()
		
		GRPTBL(GRP1).CFLG = True
		
	End Sub
	
	'******************************************************************
	'*      ���͉ۃ`�F�b�N�i�O���[�v�j                      (GVALCHK)
	'******************************************************************
	Function GVALCHK() As Short
		GVALCHK = True
		ERRSW = F_OFF
		
		If LST_NO <> n0 Then
			If CTRLTBL(CUR_NO).IGRP = CTRLTBL(LST_NO).IGRP Then Exit Function
		End If
		
		Select Case CTRLTBL(CUR_NO).IGRP
			Case GRP1
				Call GVALCHK_GRP2()
		End Select
		If ERRSW = F_ERR Then
			GVALCHK = False
			Call FOCUS_SET()
		End If
		
	End Function
	
	'******************************************************************
	'*      ���͉ۃ`�F�b�N�i�O���[�v�j�A               (GVALCHK_GRP2)
	'******************************************************************
	Sub GVALCHK_GRP2()
	End Sub
	
	'******************************************************************
	'*      ���͓��e�`�F�b�N�iLoasFocus���ڂ������j           (IPROCHK)
	'******************************************************************
	Function IPROCHK() As Short
		Dim i As Short
		
		IPROCHK = True
		ERRSW = F_OFF '�@�װ�ر
		ENDSW = F_OFF
		
		If CUR_NO = LST_NO Then Exit Function '�@���ڊԂ̈ړ����Ȃ��ꍇ�͉������Ȃ�
		
		Select Case LST_NO '�@�ړ��O���ڂ̃`�F�b�N
		End Select
		
		If ENDSW = F_END Then
			IPROCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Function
		End If
		
		'�G���[��
		If ERRSW = F_ERR Then
			If CUR_NO < LST_NO Then
				ERRSW = F_OFF
				'�t�����̂Ƃ��͒��O���ڒl�̍ĕ\��
				Select Case LST_NO
				End Select
			Else
				IPROCHK = False
				NXT_NO = LST_NO
				Call FOCUS_SET()
			End If
		End If
		
	End Function
	
	'******************************************************************
	'*      �X�v���b�h�Z�b�g����  (SET_SPRD050_UPD)
	'******************************************************************
	Private Function SET_SPRD050_UPD() As Short
		Dim i As Integer
		Dim FLG As Boolean
		Dim wROW As Integer
		
		SET_SPRD050_UPD = 0
		SPRD050.Col = SPRD_COL.col_RENBAN 'A-CUST-20100823
		SPRD050.ColHidden = True 'A-CUST-20100823
		
		i = 0
		
		WSZ0410SEL01.rdoParameters("Inc_code").Value = WKB010 '��ЃR�[�h
		WSZ0410SEL01.rdoParameters("jg_code").Value = WKB020 '���Ə��R�[�h
		
		On Error Resume Next
		WSZ0410RS = WSZ0410SEL01.OpenResultset()
		Select Case B_STATUS(WSZ0410RS)
			Case 0
				SPRD050.ReDraw = False
				Do 
					i = i + 1
					With SPRD050
						.MaxRows = i
						.ROW = i
						
						.Col = SPRD_COL.col_SEN
						If SentakuFLG Then
							If WSZ0410RS.rdoColumns("y_code").Value = RENBAN_SEN Then
								.Value = CStr(True)
								FLG = True
								wROW = i
							Else
								.Value = CStr(False)
							End If
						Else
							.Value = CStr(False)
						End If
						
						.Col = SPRD_COL.col_HIN_NAME
						.Value = Trim(WSZ0410RS.rdoColumns("hin_name_seisiki").Value)
						
						.Col = SPRD_COL.col_KIKAKU
						.Value = Trim(WSZ0410RS.rdoColumns("kikaku").Value)
						
						.Col = SPRD_COL.col_GYO_NAME
						.Value = Trim(WSZ0410RS.rdoColumns("gyo_name").Value)
						
						.Col = SPRD_COL.col_TANI
						.Value = Trim(WSZ0410RS.rdoColumns("tani").Value)
						
						.Col = SPRD_COL.col_TANKA
						.Value = Trim(WSZ0410RS.rdoColumns("tanka").Value)
						
						.Col = SPRD_COL.col_RENBAN
						.Value = Trim(WSZ0410RS.rdoColumns("y_code").Value)
						
						'A-CUST-20100823 Start
						.Col = SPRD_COL.col_TEKI_DATE
						If Trim(WSZ0410RS.rdoColumns("teki_date").Value) = "" Then
							.Value = ""
						Else
							.Value = VB6.Format(WSZ0410RS.rdoColumns("teki_date").Value, "@@@@/@@/@@")
						End If
						
						.Col = SPRD_COL.col_HA_TANI
						.Value = Trim(WSZ0410RS.rdoColumns("ha_tani").Value)
						
						.Col = SPRD_COL.col_KANSANSU
						.Value = Trim(WSZ0410RS.rdoColumns("kansansu").Value)
						
						.Col = SPRD_COL.col_JAN_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("jan_code").Value)
						
						.Col = SPRD_COL.col_JAN_S_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("jan_s_code").Value)
						
						.Col = SPRD_COL.col_BAR_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("bar_code").Value)
						'A-CUST-20100823 End
						SPRD050.set_RowHeight(i, 12)
						
						WSZ0410RS.MoveNext()
					End With
				Loop Until WSZ0410RS.EOF = True
				
				SentakuFLG = FLG
				SPRD050.Col = 1
				If FLG Then
					SPRD050.ROW = wROW
				Else
					SPRD050.ROW = 1
				End If
				SPRD050.Action = SS_ACTION_ACTIVE_CELL
				
				SPRD050.ReDraw = True
				
			Case 24
				SentakuFLG = FLG
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(RENBAN_SEN, "000000")
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
				On Error GoTo 0
				Exit Function
		End Select
		
	End Function
	
	'******************************************************************
	'*      ���ړ��͉ۃ`�F�b�N                              (MVALCHK)
	'******************************************************************
	Function MVALCHK() As Short
		MVALCHK = True
		ERRSW = F_OFF
		
		Select Case CUR_NO '�@�ړ��ۂ̃`�F�b�N
		End Select
		If ERRSW = F_ERR Then
			MVALCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
	End Function
	
	'******************************************************************
	'*      ���ړ��͉ۃ`�F�b�N�C                       (MVALCHK_N007)
	'******************************************************************
	Sub MVALCHK_N007()
		
	End Sub
	
	'�A�N�e�B�u�Z�����ړ������ꍇ�̏����F
	'    �A�N�e�B�u��p�ϐ��̋L��
	'    �Z�����l�̋L��
	'    �t�@���N�V�����ݒ�
	'    �n�����\��
	Private Sub MyProcOfCell(ByVal Col As Integer, ByVal ROW As Integer, ByVal NewCol As Integer, ByVal NewRow As Integer)
		
		With SPRD050
			'
			ViewCol = NewCol '���׌���
			Call FUNCSET_RTN() '�t�@���N�V�����\��
			
			.ROW = NewRow
			.Col = NewCol
			
			
			'����OK�t���O(�����ڂւ̈ړ�����)�̉ی��肨��сA
			'���̏��̋L��
			Select Case .Col
				Case 1
					'���l �L��
					If SPRD_ERR <> 1 Then
						'UPGRADE_WARNING: �I�u�W�F�N�g PRE_VALUE �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						PRE_VALUE = SPRD050.Value
					End If
			End Select
			
			'Col��߂�
			.Col = NewCol
			
		End With
		
	End Sub
	
	Private Sub IMTXDUM_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTXDUM.Enter
		Me.Close()
		
	End Sub
	
	'******************************************************************
	'*      ���ׁiGotFocus�j
	'******************************************************************
	Private Sub SPRD050_ButtonClicked(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SPRD050.ButtonClicked
		Dim IROWSelect As Short
		If ButtnFlg Then
			If eventArgs.ButtonDown <> 0 Then
				
				With SPRD050
					If .MaxRows > 0 Then
						For IROWSelect = 1 To .MaxRows
							If IROWSelect <> eventArgs.ROW Then
								.eventArgs.ROW = IROWSelect
								.eventArgs.Col = 1
								.Value = CStr(False)
							End If
						Next 
					End If
				End With
			End If
		End If
		
	End Sub
	
	'******************************************************************
	'*      ���ׁiGotFocus�j
	'******************************************************************
	Private Sub SPRD050_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD050.Enter
		
		If CUR_NO = N005 Then Exit Sub
		CUR_NO = N005
		
		'�`�F�b�N�J�n
		If LST_NO <> n0 Then
			'LostFocus���ڂ̃`�F�b�N
			If IPROCHK() = False Then
				Exit Sub
			End If
			'LostFocus���ڌQ�̃`�F�b�N
			If GPROCHK() = False Then
			End If
		End If
		'GotFocus���ڌQ�̃`�F�b�N
		If GVALCHK() = False Then
			Exit Sub
		End If
		'GotFocus���ڂ̃`�F�b�N
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		LST_NO = CUR_NO
		
		'�Z���ړ�������
		With SPRD050
			.Col = .ActiveCol
			.ROW = .ActiveRow
			Call MyProcOfCell(.Col, .ROW, .Col, .ROW)
		End With
		
		SPRD_ERR = 0
		
	End Sub
	
	'******************************************************************
	'*      ���ׁiKeyDown�j
	'******************************************************************
	Private Sub SPRD050_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SPRD050.KeyDownEvent
		
		With SPRD050
			'ROW COL �ݒ�
			.ROW = .ActiveRow
			.Col = .ActiveCol
			
			Select Case eventArgs.KeyCode
				
				Case System.Windows.Forms.Keys.Return 'Enter��
					
					Select Case .ActiveCol
						Case 1
							eventArgs.KeyCode = 0
							'���s�P��ڂɈړ�
							If (.ROW < .MaxRows) Then
							Else
								Call SET_NO(1)
								Exit Sub
							End If
							
					End Select
					
				Case System.Windows.Forms.Keys.Up
					
				Case System.Windows.Forms.Keys.Down
					eventArgs.KeyCode = 0
					'���͍��ڊԂ̈ړ�
					Select Case .ActiveCol
						Case 1
							If .ROW < .MaxRows Then
								'�ŏI�s����̍s�̏ꍇ�F
								'���͉\���ڂ������ꍇ�F
								.Col = 1
								.ROW = .ROW + 1
								.Action = SS_ACTION_ACTIVE_CELL
								Call MyProcOfCell(.Col, .ROW, .Col, .ROW)
							Else
								Call SET_NO(1)
								Exit Sub
							End If
					End Select
					
				Case Else
					Call SZ0412FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
					
			End Select
		End With
		
	End Sub
	
	Private Sub SPRD050_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SPRD050.LeaveCell
		
		If (eventArgs.NewRow < 0) Or (eventArgs.NewCol < 0) Then Exit Sub
		
		With SPRD050
			'�د��ŕ\�����ڂɈړ����鎖�̖h�~
			If (eventArgs.NewCol > 1) Then
				.eventArgs.Col = eventArgs.Col
				.eventArgs.ROW = eventArgs.ROW
				eventArgs.Cancel = True
				Exit Sub
			End If
			
			'�}�E�X�N���b�N�ɂ�鑼�Z���ւ̈ړ��̏ꍇ�F
			'    �E�E���ւ̈ړ��F���͒l�n�j�̏ꍇ�̂݋���
			'    ���E��ւ̈ړ��F���͒l�m�f�̏ꍇ�A���̒l�ɖ߂�
			.eventArgs.Col = eventArgs.Col
			.eventArgs.ROW = eventArgs.ROW
			Select Case eventArgs.Col
				
			End Select
			
			'�ړ���̍s�ɂ����āA�ړ���̗�������̗񂪖����͂�������A���̗�Ɉړ�
			.eventArgs.ROW = eventArgs.NewRow
			
		End With
		
LC_EXIT: 
		'�Z���ړ�������
		Call MyProcOfCell(eventArgs.Col, eventArgs.ROW, eventArgs.NewCol, eventArgs.NewRow)
		
	End Sub
	
	Private Sub SPRD050_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD050.Leave
		
		If (SPRD050.MaxRows < 1) Then Exit Sub
		If (SPRD050.ActiveRow < 0) Then Exit Sub
		If (SPRD050.ActiveCol < 0) Then Exit Sub
		
		With SPRD050
			.ROW = .ActiveRow
			.Col = .ActiveCol
			SPRD_ERR = 0
			
			Select Case .ActiveCol
				Case 1
					'UPGRADE_ISSUE: Control NAME �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					If (Me.ActiveControl.Name = "CMDOFNC") Then
						'If Me.ActiveControl.Index <> 12 Then       'D-20110621-
						'UPGRADE_ISSUE: Control Index �́A�ėp���O��� ActiveControl ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						If Me.ActiveControl.Index = 12 Or Me.ActiveControl.Index = 9 Then 'A-20110621-
						Else 'A-20110621-
							'UPGRADE_WARNING: �I�u�W�F�N�g PRE_VALUE �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							.Value = PRE_VALUE
							Exit Sub
						End If
					End If
					
			End Select
		End With
		
	End Sub
End Class