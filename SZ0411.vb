Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class SZ0411FRM
	Inherits System.Windows.Forms.Form
	'A-CUST-20100610 �t�H�[���ǉ�
	
	Dim LST_NO As Short '�O���͈ʒu���۰ه�
	Dim NXT_NO As Short '�����͈ʒu���۰ه�
	Dim CUR_NO As Short '�����͈ʒu���۰ه�
	Dim MAXNO As Short
	Dim CTRL As System.Windows.Forms.Control
	
	Dim SETSW As Short '�f�[�^�Z�b�g���F�n�m
	
	Const N200 As Short = 1 'CSV̧�ٖ�
	Const N912 As Short = 2 '�� �s
	Const NEND As Short = 3
	
	Const GRP1 As Short = 1
	Const GEND As Short = 2
	
	'UPGRADE_WARNING: �z�� CTRLTBL �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Private CTRLTBL(NEND) As CTRLTBL_S '��ʺ��۰ٔz��
	'UPGRADE_WARNING: �z�� GRPTBL �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Private GRPTBL(GEND) As GRPTBL_S '��ʸ�ٰ�ߔz��
	
	Private Structure EXPROT_PATH
		Dim EP_FPATH As String '�o�̓t�@�C���̃p�X�i�[�p
		Dim EP_FNAME As String '�o�̓t�@�C�����i�[�p
	End Structure
	Private EPF As EXPROT_PATH
	
	Private sPath As String '�J�����g�f�B���N�g���̏����ʒu
	Private sDrive As String '�J�����g�h���C�u�̏����ʒu
	Private FILECHKFLG As Short '�t�@�C���`�F�b�N��������TRUE
	Private DEL_INC_CODE As String
	Private DEL_JG_CODE As String
	
	Private blnCheckPass As Boolean
	
	Private Sub TBL_SET() '��ʺ��۰ُ����ݒ�
		
		'�O���[�v�̐ݒ�
		CTRLTBL(N200).IGRP = GRP1
		
		CTRLTBL(N912).IGRP = GEND
		CTRLTBL(NEND).IGRP = GEND
		
		'�����ځA�O���ڂ̐ݒ�
		CTRLTBL(N200).INEXT = N912
		CTRLTBL(N200).IBACK = 0
		CTRLTBL(N200).IDOWN = N912
		
		CTRLTBL(N912).INEXT = n0
		CTRLTBL(N912).IBACK = N200
		CTRLTBL(N912).IDOWN = n0
		
		CTRLTBL(N200).CTRL = IMTX200
		
		CTRLTBL(N912).CTRL = CMDOFNC(12)
		
		MAXNO = NEND
		
		NXT_NO = N200
	End Sub
	
	Private Sub FUNCSET_RTN()
		
		'--- �t�@���N�V�����E�K�C�h���b�Z�[�W
		Select Case LST_NO
			Case N200 'CSV̧�ٖ�
				CMDOFNC(5).Text = "�N���A"
				CMDOFNC(5).Enabled = True
				LBLFNC(5).Enabled = True
				CMDOFNC(8).Text = "�t�@�C��"
				CMDOFNC(8).Enabled = True
				LBLFNC(8).Enabled = True
				ZAGD_NO.Value = "048"
			Case Else
				CMDOFNC(5).Text = ""
				CMDOFNC(5).Enabled = False
				LBLFNC(5).Enabled = False
				CMDOFNC(8).Text = ""
				CMDOFNC(8).Enabled = False
				LBLFNC(8).Enabled = False
				ZAGD_NO.Value = ""
		End Select
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call ZAFC_SUB(Me)
		
		'--- �K�C�h���b�Z�[�W�\��
		Call ZAGD_SUB(Me)
	End Sub
	
	Public Sub ENABLED_RTN(ByRef TF As Short)
		'��ʂ̗L���A�����A�\�����e�ݒ�
		
		'�����w����
		IMTX200.Enabled = TF '�����ݒ��
		
		'��������݂̐���
		'�o��
		If TF = False Then
			CMDOFNC(12).Text = "��  �f"
			CMDOFNC(12).MousePointer = System.Windows.Forms.Cursors.Arrow
		Else
			CMDOFNC(12).Text = "��  �s"
			CMDOFNC(12).MousePointer = System.Windows.Forms.Cursors.Default
		End If
		
		If TF = False Then
			CMDOFNC(0).Text = ""
			CMDOFNC(8).Text = ""
		Else
			CMDOFNC(0).Text = ZAFC_MST(1)
			CMDOFNC(8).Text = ZAFC_MST(8)
		End If
		
		CMDOFNC(0).Enabled = TF '�I��
		LBLFNC(0).Enabled = TF
		CMDOFNC(8).Enabled = TF '�t�@�C��
		LBLFNC(8).Enabled = TF
		
		'�K�C�h���b�Z�[�W�\��
		If TF = False Then 'TRUE �� FALSE
			ZAGD_NO.Value = ""
			Call ZAGD_SUB(Me)
		Else
			Call ZAGD_SUB(Me) '�޲��ү���޸ر
		End If
		Me.Refresh()
	End Sub
	
	Private Sub ENB_SET_RTN(ByRef TGNO As Short)
		'��޽į�߂̐ݒ�
		
		Select Case TGNO
			Case N200
				IMTX200.TabStop = True
			Case Else
				IMTX200.TabStop = True
		End Select
	End Sub
	
	Private Sub CDL_INIT()
		'CommonDialog�̏����ݒ�����܂��B
		
		'�L�����Z�����G���[�Ƃ��Ĉ���
		'UPGRADE_WARNING: The CommonDialog CancelError �v���p�e�B�� Visual Basic .NET �ŃT�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"' ���N���b�N���Ă��������B
		CDL010.CancelError = True
		'[�t�@�C�����J��]�޲�۸��ޯ���ɐݒ�
		'UPGRADE_WARNING: MSComDlg.CommonDialog �v���p�e�B CDL010.Flags �́A�V������������� CDL010Open.ShowReadOnly �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' ���N���b�N���Ă��������B
		'UPGRADE_WARNING: FileOpenConstants �萔 FileOpenConstants.cdlOFNHideReadOnly �́A�V������������� OpenFileDialog.ShowReadOnly �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' ���N���b�N���Ă��������B
		CDL010Open.ShowReadOnly = False
		'���X�g �{�b�N�X�ɕ\�������t�B���^��ݒ�
		'UPGRADE_WARNING: Filter �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
		CDL010Open.Filter = "CSV(TAB��؂�)|*.CSV|�S�Ă�̧��|*.*|"
		'CSV ������̃t�B���^�Ƃ��Ďw��
		CDL010Open.FilterIndex = 1
		'�f�t�H���g�̊g���q��ݒ�
		CDL010Open.DefaultExt = ".txt"
		'�f�t�H���g�̃f�B���N�g����ݒ�(ini����擾)
		CDL010Open.InitialDirectory = WG_EXCELPATH
	End Sub
	
	Private Sub INITIAL_RTN()
		'��ʍ��ڏ����l�ݒ�
		WKBCSVFILE = WG_EXCELPATH
		
		DSP010.Text = RTrim(WKB010)
		DSP020.Text = RTrim(SZ0410FRM.DSP010.Text)
		DSP030.Text = RTrim(WKB020)
		DSP040.Text = RTrim(SZ0410FRM.DSP020.Text)
		
		'�\��
		IMTX200.Text = RTrim(WKBCSVFILE)
		
		EPF.EP_FNAME = ""
		EPF.EP_FPATH = WG_EXCELPATH
		
		Call CDL_INIT()
	End Sub
	
	Private Sub FOCUS_SET() '̫������
		Select Case NXT_NO
			Case N200 'CSV̧�ٖ�
				IMTX200.Focus()
			Case N912 '���s
				CMDOFNC(12).Focus()
		End Select
	End Sub
	
	Private Sub SET_NO(ByRef FUNC As Short) '��ʺ��۰ه��Z�b�g
		Dim i As Short
		
		i = LST_NO
		
		Do 
			Select Case FUNC
				Case 1 '������
					NXT_NO = CTRLTBL(i).INEXT
				Case 2 '�O����
					NXT_NO = CTRLTBL(i).IBACK
				Case 3 '���O���[�v
					NXT_NO = CTRLTBL(i).IDOWN
			End Select
			
			If NXT_NO = n0 Then Exit Sub
			
			If CTRLTBL(NXT_NO).CTRL.TabStop = True And CTRLTBL(NXT_NO).CTRL.Enabled = True And CTRLTBL(NXT_NO).CTRL.Visible = True Then
				Call FOCUS_SET()
				Exit Sub
			Else
				i = NXT_NO
			End If
		Loop 
	End Sub
	
	Private Sub ALLCHK_RTN()
		Dim IDX As Short
		Dim CHKFLG As Short
		'�S�`�F�b�N & ���s
		
		CUR_NO = NEND
		
		'���O���ڂ̃`�F�b�N
		If IPROCHK() = False Then
			Exit Sub
		End If
		
		'�S�O���[�v�̃`�F�b�N
		If GPROCHK() = False Then
			Exit Sub
		End If
		
		'If MsgBox("�b�r�u�̎捞�݂��s���܂��B��낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2, Me.Caption) = vbNo Then    'D-CUST-20100901
		If MsgBox("�b�r�u�̎捞�݂��s���܂��B��낵���ł����H", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then 'A-CUST-20100901
			NXT_NO = N200
			Call FOCUS_SET()
			Exit Sub
		End If
		
		'�������O�̏o�́i����ޭ��A��������s�������_�ŏ������ށj
		'�T�[�o�[�̓��t�E�������擾
		Dim strSvrDate As String
		
		SYSDATE = CduServerDate
		strSvrDate = VB6.Format(SYSDATE, "YYYYMMDDHHNNSS")
		
		'���O�o�̓T�u���[�`���ďo
		ZALGM_INC_CODE.Value = WKB010 '���
		ZALGM_JG_CODE.Value = WKB020 '���Ə�
		ZALGM_SYS_KBN.Value = "3"
		ZALGM_S_DAY.Value = VB.Left(strSvrDate, 8) '���t
		ZALGM_S_TIME.Value = VB.Right(strSvrDate, 6) '����
		ZALGM_OP_CODE.Value = WG_OPCODE '�I�y���[�^�R�[�h
		ZALGM_PGID.Value = "SZ0410" '�V�X�e��
		ZALGM_SH_KBN.Value = "3"
		ZALGM_SH_NAIYO.Value = WKB010 & "-" & WKB020
		ZALGM_GNFLG.Value = "0"
		Call ZALGM_SUB(ZACNA_RCN)
		
		'�}�E�X�J�[�\���������v�ɐݒ�
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		Call ENABLED_RTN(False)
		
		PRNSW = F_ON
		
		Call BEGIN_RTN()
		
		FSTSW = F_FST
		ENDSW = F_OFF
		ERRSW = F_OFF
		'*** �捞���� ***
		Call TORIKOMI_RTN()
		If CSV_CNT = 0 And ERRSW <> F_ERR Then
			'CSV�t�@�C�����ɑΏۂƂȂ�f�[�^���Ȃ�����
			ZAER_CD = 129 '�Ώۃf�[�^�Ȃ�
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			ERRSW = F_ERR
		End If
		
		If ERRSW <> F_ERR Then
			'�f�[�^�捞�����I
			Call COMMIT_RTN()
			
			MsgBox("����I�����܂����B", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
		Else
			'�f�[�^�捞���s�I
			Call ROLLBACK_RTN()
		End If
		
		PRNSW = F_OFF
		ENDSW = F_OFF
		
		Call ENABLED_RTN(True)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		If ERRSW = F_OFF Then
			Call HIDE_RTN()
		Else
			ERRSW = F_OFF
			NXT_NO = N200
			Call FOCUS_SET()
		End If
	End Sub
	
	Private Sub HIDE_RTN()
		picDummy.Focus()
		Me.Hide()
		
	End Sub
	
	Private Function IPROCHK() As Short 'LostFocus���ڃ`�F�b�N
		IPROCHK = True
		ERRSW = F_OFF
		
		If CUR_NO = LST_NO Then Exit Function
		Select Case LST_NO
			Case N200 'CSV̧�ٖ�
				'Call IPROCHK_N200
		End Select
		
		If ERRSW = F_ERR Or CUR_NO < LST_NO Then
			Select Case LST_NO
				Case N200
					IMTX200.Text = RTrim(WKBCSVFILE)
			End Select
			
			If ERRSW = F_ERR Then
				IPROCHK = False
				NXT_NO = LST_NO
				Call FOCUS_SET()
			End If
		End If
	End Function
	
	Private Function CHK_FNAME(ByRef fname As String) As Boolean
		'�t�@�C�����Ɏg���Ȃ��������g�p���Ă��邩�ǂ����̃`�F�b�N
		Dim i As Object
		Dim l As Short
		Dim wname As String
		Dim dnum As Short
		
		CHK_FNAME = False
		
		wname = Trim(fname)
		
		l = Len(wname)
		dnum = 0
		
		'�t�@�C�����Ɏg���Ȃ����������͂���Ă�H
		For i = 3 To l '3�����ڂ��猩��
			'UPGRADE_WARNING: �I�u�W�F�N�g i �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			Select Case Mid(wname, i, 1)
				Case "/", ":", ",", ";", "*", "?", "<", ">", "|", """"
					Exit Function
				Case "."
					If dnum >= 1 Then Exit Function
					dnum = dnum + 1
			End Select
		Next 
		
		CHK_FNAME = True
	End Function
	
	Private Sub IPROCHK_N200()
		Dim PCRET As Short '�p�X�`�F�b�N�T�u���[�`���̖߂�l
		Dim FULLNAME As String '���E�̃X�y�[�X���J�b�g�����t�@�C�����i�[�G���A
		Dim i As Object
		Dim j As Integer
		Dim wStr As String
		
		If CUR_NO < LST_NO Then
			Exit Sub
		End If
		
		'�����͂̃`�F�b�N
		If VB.Right(Trim(IMTX200.Text), 1) = "\" Or Trim(IMTX200.Text) = "" Then
			ERRSW = F_ERR
			Exit Sub
		End If
		
#If 0 Then
		'UPGRADE_NOTE: �� 0 �� True �ɕ]������Ȃ��������A�܂��͂܂������]������Ȃ��������߁A#If #EndIf �u���b�N�̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"' ���N���b�N���Ă��������B
		'�p�X���w�肳��Ă��Ȃ���΁A�f�t�H���g�̃p�X���w�肳����
		If InStr(1, Trim$(IMTX200.Text), "\", vbTextCompare) = 0 Then
		wStr = ""
		'�t�@�C�����ƃp�X�̊Ԃ�"\"���K�v���ǂ����̔���
		If Right$(Trim$(WG_EXCELPATH), 1) <> "\" And Left$(Trim$(IMTX200.Text), 1) <> "\" Then
		wStr = "\"
		End If
		
		IMTX200.Text = WG_EXCELPATH & wStr & Trim$(IMTX200.Text)
		End If
#End If
		
		'   '�g���q�̃`�F�b�N
		'    Select Case StrConv(Right$(Trim$(IMTX200.Text), 4), vbUpperCase)
		'        Case ".CSV"
		'        Case Else
		'            IMTX200.Text = Trim$(IMTX200.Text) & ".CSV"
		'    End Select
		
		'�t�@�C�����Ɏg���Ȃ����������͂���Ă邩�ǂ����̃`�F�b�N
		If CHK_FNAME((IMTX200.Text)) = False Then
			ZAER_CD = 11 '̧�ٖ��s��
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'�t�@�C�����`�F�b�N���܂��������Ƃ������A�`�F�b�N���s��
		'(���x���㏑���m�F���o���Ȃ�����)
		If FILECHKFLG = False Then
			'�t�@�C�����`�F�b�N
			FULLNAME = Trim(IMTX200.Text)
			PCRET = MKKCMN.ZAPC_SUB(FULLNAME)
			Select Case PCRET
				Case 0 '̧�ٖ���
					ZAER_CD = 12
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
				Case -1 '�t�@�C�����������FOK�I
				Case 11 '�t�@�C�����s��
					ZAER_CD = PCRET
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
				Case 190 '�h���C�u�̏������ł��Ă��Ȃ�
					ZAER_CD = PCRET
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
				Case Else
					ZAER_CD = 11
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
			End Select
			
			'�`�F�b�N�����킾������A�`�F�b�N�ς݃t���O�����Ă�
			FILECHKFLG = True
			WKBCSVFILE = Trim(IMTX200.Text)
		End If
	End Sub
	
	Private Function GPROCHK() As Short 'LostFocus���ڌQ�`�F�b�N
		GPROCHK = True
		ERRSW = F_OFF
		
		If CTRLTBL(CUR_NO).IGRP <= CTRLTBL(LST_NO).IGRP Then
			Exit Function
		End If
		Select Case CTRLTBL(LST_NO).IGRP
			Case GRP1
				Call GPROCHK_GRP1()
		End Select
		If ERRSW = F_ERR Then
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = False
			GPROCHK = False
			NXT_NO = GRPTBL(CTRLTBL(LST_NO).IGRP).NXTN
			Call FOCUS_SET()
		Else
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = True
		End If
	End Function
	
	Private Sub GPROCHK_GRP1()
		Call IPROCHK_N200()
		If ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N200
			ZAER_CD = 120
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			Exit Sub
		End If
	End Sub
	
	Private Function GVALCHK() As Short '���ڌQ���͉ۃ`�F�b�N
		GVALCHK = True
		ERRSW = F_OFF
		
		If LST_NO <> n0 Then
			If CTRLTBL(CUR_NO).IGRP = CTRLTBL(LST_NO).IGRP Then Exit Function
		End If
		Select Case CTRLTBL(CUR_NO).IGRP
			Case GRP1
				Call GVALCHK_GRP1()
			Case GEND
				Call GVALCHK_GEND()
		End Select
		If ERRSW = F_ERR Then
			GRPTBL(CTRLTBL(CUR_NO).IGRP).CFLG = False
			GVALCHK = False
		Else
			GRPTBL(CTRLTBL(CUR_NO).IGRP).CFLG = True
		End If
	End Function
	
	Private Sub GVALCHK_GRP1()
		
	End Sub
	
	Private Sub GVALCHK_GEND()
		If GRPTBL(CTRLTBL(GRP1).IGRP).CFLG = False Then
			ERRSW = F_ERR
		End If
		
	End Sub
	
	Private Function MVALCHK() As Short '���ړ��͉ۃ`�F�b�N
		MVALCHK = True
		ERRSW = F_OFF
		
		Select Case CUR_NO
			Case N200 'CSV̧�ٖ�
				Call MVALCHK_N200()
		End Select
		
		If ERRSW = F_ERR Then
			MVALCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
	End Function
	
	Sub MVALCHK_N200() 'CSV̧�ٖ�
		
	End Sub
	
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		If CMDOFNC(Index).Text = "" Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Sub
		End If
		
		Select Case Index
			Case 0 '�I��
				Call HIDE_RTN()
			Case 5 '�N���A
				WKBCSVFILE = RTrim(WG_EXCELPATH)
				IMTX200.Text = WKBCSVFILE
				NXT_NO = N200
				Call FOCUS_SET()
				
			Case 8 '�t�@�C��
				Call GETDIR_RTN() '�޲�۸��ޯ�����A�t�@�C�����y�уp�X���擾����
				ChDir((sPath)) '�J�����g�f�B���N�g����߂�
				ChDrive((sDrive)) '�J�����g�h���C�u��߂�
				NXT_NO = CUR_NO
				Call FOCUS_SET()
				
			Case 12 '���s
				'�K�C�h���b�Z�[�W�\��
				If PRNSW = F_OFF Then
					'�}�E�X�J�[�\���������v�ɐݒ�
					Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
					'�S�`�F�b�N & ���s
					Call ALLCHK_RTN()
					'�}�E�X�J�[�\����ʏ탂�[�h�ɐݒ�
					Me.Cursor = System.Windows.Forms.Cursors.Default
				Else '������f
					If MsgBox("���f���܂����H", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.Yes Then
						CANSW = F_CAN
						ENDSW = F_END
					Else
						ActiveControl.Focus()
					End If
				End If
				blnCheckPass = False
		End Select
	End Sub
	
	Private Sub GETDIR_RTN()
		Dim StrLen As Integer
		Dim i As Integer
		
		On Error GoTo ErrHandler
		'�޲�۸��ޯ���\���I
		CDL010Open.ShowDialog()
		
		' ���[�U�[�� [�J��] ���N���b�N�����B
		
		'�t�@�C�����A�p�X�����o���B
		'UPGRADE_WARNING: CommonDialog �v���p�e�B CDL010.FileTitle �́A�V������������� CDL010.FileName �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' ���N���b�N���Ă��������B
		EPF.EP_FNAME = CDL010Open.FileName '�t�@�C����
		StrLen = InStr(CDL010Open.FileName, EPF.EP_FNAME)
		EPF.EP_FPATH = Mid(CDL010Open.FileName, 1, StrLen - 1) '�p�X
		
		'�m��
		WKBCSVFILE = CDL010Open.FileName '�t�@�C����(�p�X���܂�)
		IMTX200.Text = WKBCSVFILE
		
		'�޲�۸��ޯ���Đݒ�
		CDL010Open.InitialDirectory = EPF.EP_FPATH
		'UPGRADE_WARNING: CommonDialog �v���p�e�B CDL010.FileTitle �́A�V������������� CDL010.FileName �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' ���N���b�N���Ă��������B
		CDL010Open.FileName = CDL010Open.FileName
		
ErrHandler: 
		' ���[�U�[�� [�L�����Z��] ���N���b�N�����B
		Exit Sub
	End Sub
	
	Private Sub CMDOFNC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Enter
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		'����ޭ��A���s�{�^���ȊO�͖���
		If Index < 12 Then
			Exit Sub
		End If
		If blnCheckPass Then
			Exit Sub
		End If
		
		If Index = 12 Then
			If CUR_NO = N912 Then GoTo CMDOFNC_END
			CUR_NO = N912
		End If
		
		If IPROCHK() = False Then
			Exit Sub
		End If
		If GPROCHK() = False Then
		End If
		If GVALCHK() = False Then
			Exit Sub
		End If
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		LST_NO = CUR_NO
		
CMDOFNC_END: 
		' �t�@���N�V�����K�C�h
		Call FUNCSET_RTN()
	End Sub
	
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		'����ޭ��A����ȊO�͖���
		If Index < 12 Then
			Exit Sub
		End If
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				Select Case Index
					Case 12 '���s
						NXT_NO = N200 'CSV̧�ٖ���
						Call FOCUS_SET()
						Exit Sub
				End Select
		End Select
		
		Call SZ0411FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub
	
	Private Sub CMDOFNC_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDOFNC.MouseDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		MOUSEFLG = eventArgs.Button
	End Sub
	
	'UPGRADE_ISSUE: PictureBox �C�x���g DUMMYDEL.GotFocus �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub DUMMYDEL_GotFocus()
		'�ߋ��f�[�^�@�폜
		'�������O�̏o�́i����ޭ��A��������s�������_�ŏ������ށj
		'�T�[�o�[�̓��t�E�������擾
		Dim strSvrDate As String
		
		SYSDATE = CduServerDate
		strSvrDate = VB6.Format(SYSDATE, "YYYYMMDDHHNNSS")
		
		'���O�o�̓T�u���[�`���ďo
		ZALGM_INC_CODE.Value = WKB010 '���
		ZALGM_JG_CODE.Value = WKB020 '���Ə�
		ZALGM_SYS_KBN.Value = "3"
		ZALGM_S_DAY.Value = VB.Left(strSvrDate, 8) '���t
		ZALGM_S_TIME.Value = VB.Right(strSvrDate, 6) '����
		ZALGM_OP_CODE.Value = WG_OPCODE '�I�y���[�^�R�[�h
		ZALGM_PGID.Value = "SZ0410" '�V�X�e��
		ZALGM_SH_KBN.Value = "5"
		ZALGM_SH_NAIYO.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Month, -1, SYSDATE), "YYYYMMDD")
		ZALGM_GNFLG.Value = "0"
		Call ZALGM_SUB(ZACNA_RCN)
		
		'�}�E�X�J�[�\���������v�ɐݒ�
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Call BEGIN_RTN()
		
		ERRSW = F_OFF
		'*** �ʕi�ڃ}�X�^�捞���� ***
		Call TORIKOMI_DEL()
		Me.Cursor = System.Windows.Forms.Cursors.Default
		If ERRSW = F_ERR Then
			Call ROLLBACK_RTN()
			Call HIDE_RTN()
		Else
			Call COMMIT_RTN()
			DEL_INC_CODE = WKB010
			DEL_JG_CODE = WKB020
			LST_NO = N200
			NXT_NO = N200
			Call FOCUS_SET()
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Form �C�x���g SZ0411FRM.Activate �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
	Private Sub SZ0411FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'�}�E�X�J�[�\����߂�
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		Me.Text = "�i�ڏ����͂b�r�u�捞"
		
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
		lRet = MKKDBCMN.MKKDBCMN_SQTGET1_SUB(3, "SZ0411", WKB010, WG_OPCODE, OP_KENGEN) 'A-CUST-20100901
		If lRet <> n0 Then
			Call HIDE_RTN()
			Exit Sub
		End If
		If OP_KENGEN = 0 Then
			ZAER_KN = n0
			ZAER_CD = 301
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			Call HIDE_RTN()
			Exit Sub
		End If
		
		''�X�V�����`�F�b�N
		''�X�V����
		'lRet = MKKDBCMN.MKKDBCMN_SQTGET3_SUB(4, "SZ0410", WKB010, WKB020, "", WG_OPCODE, OP_KENGEN)
		'If lRet <> n0 Then
		'    Call HIDE_RTN
		'    Exit Sub
		'End If
		''�X�V�����Ȃ�
		'If OP_KENGEN = 0 Then
		'    ZAER_KN = 0
		'    ZAER_CD = 303
		'    ZAER_NO = ""
		'    Call ZAER_SUB
		'    Call HIDE_RTN
		'    Exit Sub
		'End If
		
		LST_NO = N200
		NXT_NO = N200
		If DEL_INC_CODE = WKB010 And DEL_JG_CODE = WKB020 Then
		Else
			DUMMYDEL.Focus()
			Exit Sub
		End If
		
		Call FOCUS_SET()
	End Sub
	
	Private Sub SZ0411FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If Me.Enabled = False Then Exit Sub
		If Shift <> n0 Then Exit Sub
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
			Case System.Windows.Forms.Keys.Return
				Call SET_NO(1) ' ������
				KeyCode = n0
			Case System.Windows.Forms.Keys.Up
				Call SET_NO(2) ' �O����
				KeyCode = n0
			Case System.Windows.Forms.Keys.Down
				Call SET_NO(3) ' ���O���[�v
				KeyCode = 0
			Case System.Windows.Forms.Keys.F5 'F5
				If CMDOFNC(5).Text <> "" Then
					CMDOFNC(5).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F8 'F8
				If CMDOFNC(8).Text <> "" Then
					CMDOFNC(8).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(8), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F12 'F12
				If CMDOFNC(12).Text <> "" Then
					ERRSW = F_OFF
					blnCheckPass = True
					CMDOFNC(12).Focus()
					System.Windows.Forms.Application.DoEvents()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				End If
				KeyCode = n0
		End Select
	End Sub
	
	Private Sub SZ0411FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: Form �v���p�e�B SZ0411FRM.HelpContextID �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
		Me.HelpContextID = SM_HelpContextID
		
		Call TBL_SET() '��ʺ��۰ُ����ݒ�
		LST_NO = NEND
		
		sPath = CurDir() '�J�����g�f�B���N�g���̏����ʒu�擾
		sDrive = VB.Left(CurDir(), 1) '�J�����g�h���C�u�̏����ʒu�擾
	End Sub
	
	Private Sub SZ0411FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		'�I�� or ���f
		'UPGRADE_ISSUE: �萔 vbFormCode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
		If UnloadMode <> vbFormCode Then
			If SETSW = F_ON Or PRNSW = F_ON Then
				If MsgBox("���f���܂����H", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question, Me.Text) = MsgBoxResult.Yes Then
					CANSW = F_CAN
					ENDSW = F_END
				Else
					Cancel = True
					If SETSW <> F_ON Then
						ActiveControl.Focus()
					End If
					Exit Sub
				End If
			End If
		End If
		Call HIDE_RTN()
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub IMTX200_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX200.Change
		' �l���ύX���ꂽ�̂ŁA�`�F�b�N�ς݃t���O�𖢃`�F�b�N�ɂ���
		FILECHKFLG = False
	End Sub
	
	Private Sub IMTX200_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX200.Enter
		'CSV̧�ٖ�
		
		CUR_NO = N200
		
		'�`�F�b�N
		If LST_NO <> CUR_NO Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
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
		
		'�t�@���N�V�����K�C�h
		Call FUNCSET_RTN()
	End Sub
	
	Private Sub IMTX200_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX200.KeyDownEvent
		'CSV̧�ٖ�
		Call SZ0411FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub
End Class