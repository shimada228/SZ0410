Option Strict Off
Option Explicit On
Friend Class SZ0415FRM
	Inherits System.Windows.Forms.Form
	
	Const N010 As Short = 1 '�啪��
	Const N020 As Short = 2 '������
	Const N030 As Short = 3 '������
	Const N040 As Short = 4 '�\���{�^��
	Const N050 As Short = 5 '�ڍו���
	Dim LST_NO As Short
	Dim CUR_NO As Short
	Dim NXT_NO As Short
	
	
	Dim RDO_STATUS As Short
	Dim CM_IDX As Short 'INDEX ���׈ړ��pINDEX
	
	Dim CM_LNCNT As Short '�s�ԍ�
	Dim CM_REP As Short '�J�E���^�[
	Dim CM_I As Short '���ו\���p�J�E���^�[
	Dim CM_IMAX As Short '�\���s��(1�`10)
	Dim CM_ENDSW As Short 'AT END ����SW=1
	
	'�����������n�p
	Dim SZ0415_JANBUNCD As New VB6.FixedLengthString(6) '�i�`�m���i���ރR�[�h
	Dim MOUSEFLG As Short
	
	Private Sub ENDRR_RTN(ByRef MyForm As System.Windows.Forms.Form)
		'�R�[�h�⍇���t�H�[���I��������
		Dim Ret As Integer
		
		'    '�E�B���h�E�\���ʒu�Z�[�u
		'    Ret = GetWindowRect(MyForm.hwnd, lpRectSave)
		
		'�t�H�[���A�����[�h
		MyForm.Close()
		
	End Sub
	
	Private Sub INITIAL_RTN()
		'��ʍ��ڏ����l�ݒ�
		
		Call COMBO_INIT_SZ0415(CMB010, 1)
		Call COMBO_INIT_SZ0415(CMB020, 2)
		Call COMBO_INIT_SZ0415(CMB030, 3)
		
		Call COMBO_SETLIST_SZ0415(CMB010, SZ0415_DAI_CODES.Value)
		Call COMBO_SETLIST_SZ0415(CMB020, SZ0415_CHU_CODES.Value)
		Call COMBO_SETLIST_SZ0415(CMB030, SZ0415_SHO_CODES.Value)
		
		'�X�C�b�`���ڃN���A
		CM_ENDSW = F_OFF
		
		SPRD.MaxRows = 0
		
	End Sub
	
	Private Sub COMBO_INIT_SZ0415(ByRef cBox As System.Windows.Forms.ComboBox, ByRef InitType As Short)
		Dim sCODE As Short
		Dim wStr As String
		
		CM_ENDSW = 9
		Call SZ0415_GET_SUB(InitType)
		
		cBox.Items.Clear() '�R���{�{�b�N�X �N���A
		cBox.Items.Add("")
		If CM_ENDSW = 1 Then Exit Sub
		
		Do Until SZ0415RS.EOF
			Select Case InitType
				Case 1
					sCODE = CShort(VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "0"))
					wStr = VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "0") & Space(5)
				Case 2
					sCODE = CShort(VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "00"))
					wStr = VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "00") & Space(4)
				Case 3
					sCODE = CShort(VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "0000"))
					wStr = VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "0000") & Space(2)
			End Select
			cBox.Items.Add(New VB6.ListBoxItem(wStr & RTrim(SZ0415RS.rdoColumns("BK4").Value), sCODE))
			SZ0415RS.MoveNext()
		Loop 
		
	End Sub
	
	Private Sub COMBO_SETLIST_SZ0415(ByRef cBox As System.Windows.Forms.ComboBox, ByRef Txt As String)
		
		Dim lx As Integer
		For lx = 0 To cBox.Items.Count - 1
			If Trim(CStr(VB6.GetItemData(cBox, lx))) = Trim(Txt) Then
				cBox.SelectedIndex = lx
				Exit Sub
			End If
		Next lx
		cBox.SelectedIndex = -1
		
	End Sub
	
	Private Sub CMB010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB010.Enter
		
		If CUR_NO = N010 Then Exit Sub
		CUR_NO = N010
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call FUNCSET_RTN
		
	End Sub
	
	Private Sub CMB010_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB010.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0415FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	Private Sub CMB020_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB020.Enter
		
		If CUR_NO = N020 Then Exit Sub
		CUR_NO = N020
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call FUNCSET_RTN
		
	End Sub
	
	Private Sub CMB020_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB020.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0415FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	
	Private Sub CMB030_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB030.Enter
		
		If CUR_NO = N030 Then Exit Sub
		CUR_NO = N030
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call FUNCSET_RTN
		
	End Sub
	
	Private Sub CMB030_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB030.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0415FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	
	Private Sub CMDODSP_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDODSP.ClickEvent
		'�����J�n
		Call KENSAKU_START_RTN()
		
	End Sub
	
	Private Sub CMDODSP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDODSP.Enter
		
		If CUR_NO = N040 Then Exit Sub
		CUR_NO = N040
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call FUNCSET_RTN
		
	End Sub
	
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		Dim WK_DATA As Object
		Dim WVAL As Object
		
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		
		Select Case Index
			Case 0 '�I��
				'��ʓ��͏������Ȃ��̂ŕ\���Ϗ��͕ێ����Ȃ�
				SZ0415_KBN = -1
				Call ENDRR_RTN(Me)
				
			Case 5 '�N���A
				SZ0415_DAI_CODES.Value = ""
				SZ0415_CHU_CODES.Value = ""
				SZ0415_SHO_CODES.Value = ""
				Call COMBO_INIT_SZ0415(CMB010, 1)
				CMB020.Items.Clear()
				CMB030.Items.Clear()
				SPRD.MaxRows = 0
				
			Case 12 '�I��
				SZ0415_KBN = 0
				Call SPRD.GetText(1, SPRD.ActiveRow, WVAL)
				'UPGRADE_WARNING: �I�u�W�F�N�g WVAL �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				SZ0415_SEL_CODES.Value = WVAL
				'D-20130328-S
				''''            Call ENDRR_RTN(Me)
				'D-20130328-E
				'A-20130328-S
				SZ0415_DAI_CODE.Value = SZ0415_DAI_CODES.Value
				SZ0415_CHU_CODE.Value = SZ0415_CHU_CODES.Value
				SZ0415_SHO_CODE.Value = SZ0415_SHO_CODES.Value
				
				SZ0415_SPRD = SPRD.ActiveRow
				Call SPRD.GetText(1, 1, WVAL)
				Me.Visible = False
				'A-20130328-E
				
		End Select
		
	End Sub
	
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		If Me.Enabled = False Then Exit Sub
		
		If eventArgs.Shift <> n0 Then Exit Sub
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Escape
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
		End Select
		
	End Sub
	
	Private Sub CMDOFNC_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDOFNC.MouseDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		MOUSEFLG = eventArgs.Button
		
	End Sub
	
	'UPGRADE_WARNING: Form �C�x���g SZ0415FRM.Activate �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
	Private Sub SZ0415FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Me.Cursor = System.Windows.Forms.Cursors.Default '�}�E�X�J�[�\����߂�
		
		SZ0415_JANBUNCD.Value = Space(6)
		
		'�����J�n
		Call KENSAKU_START_RTN()
		
		SPRD.ReDraw = False
		
		If (CM_IMAX <> 0) Then
			SPRD.ROW = SZ0415_SPRD
			SPRD.Col = 1
			SPRD.Action = n0
			SPRD.Focus()
		End If
		
		SPRD.ReDraw = True
		
		If CM_DSP1SW <> n0 Then
			Exit Sub
		End If
		
		CM_DSP1SW = n1
		
	End Sub
	
	Private Sub SZ0415FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
				Select Case LST_NO
					Case N010 : CMB020.Focus()
					Case N020 : CMB030.Focus()
					Case N030 : CMDODSP.Focus()
					Case N040 : SPRD.Focus()
					Case N050 : Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				End Select
				KeyCode = 0
				
			Case System.Windows.Forms.Keys.Up
			Case System.Windows.Forms.Keys.Down
			Case System.Windows.Forms.Keys.End
			Case System.Windows.Forms.Keys.F5 '�N���A
				If CMDOFNC(5).Text <> "" Then
					CMDOFNC(5).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				KeyCode = 0
				
			Case System.Windows.Forms.Keys.F12 '�I��
				If CMDOFNC(12).Text <> "" Then
					CMDOFNC(12).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				KeyCode = 0
				
		End Select
		
	End Sub
	
	Private Sub SZ0415FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: Form �v���p�e�B SZ0415FRM.HelpContextID �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
		Me.HelpContextID = SM_HelpContextID
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor '�}�E�X�J�[�\���������v�ɐݒ�
		
		'���ʃA�C�R���̕\��
		Call ZAWC_SUB(Me, 3)
		Call INITIAL_RTN() '������ʕ\��
		
		'����\���̏ꍇ
		
		Select Case SZ0415_PS
			Case 0 '����
				Me.Top = VB6.TwipsToPixelsY(SZ0415_TOPS + ((SZ0415_HEIGHTS - VB6.PixelsToTwipsY(Me.Height)) \ 2))
				Me.Left = VB6.TwipsToPixelsX(SZ0415_LEFTS + ((SZ0415_WIDTHS - VB6.PixelsToTwipsX(Me.Width)) \ 2))
			Case 1 '����
				Me.Top = VB6.TwipsToPixelsY(SZ0415_TOPS + 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0415_LEFTS + 200)
			Case 2 '�E��
				Me.Top = VB6.TwipsToPixelsY(SZ0415_TOPS + 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0415_LEFTS + SZ0415_WIDTHS - VB6.PixelsToTwipsX(Me.Width) - 200)
			Case 3 '����
				Me.Top = VB6.TwipsToPixelsY(SZ0415_TOPS + SZ0415_HEIGHTS - VB6.PixelsToTwipsY(Me.Height) - 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0415_LEFTS + 200)
			Case 4 '�E��
				Me.Top = VB6.TwipsToPixelsY(SZ0415_TOPS + SZ0415_HEIGHTS - VB6.PixelsToTwipsY(Me.Height) - 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0415_LEFTS + SZ0415_WIDTHS - VB6.PixelsToTwipsX(Me.Width) - 200)
		End Select
		
	End Sub
	
	Private Sub SZ0415FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		Dim WVAL As Object
		
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Then
			SZ0415_KBN = -1
		End If
		
		SZ0415_DAI_CODE.Value = SZ0415_DAI_CODES.Value
		SZ0415_CHU_CODE.Value = SZ0415_CHU_CODES.Value
		SZ0415_SHO_CODE.Value = SZ0415_SHO_CODES.Value
		
		SZ0415_SPRD = SPRD.ActiveRow
		Call SPRD.GetText(1, 1, WVAL)
		
		Call ENDRR_RTN(Me)
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub FUNCSET_RTN()
		
		ZAFC_N(0) = 1
		
		'�����f�[�^�L��̏ꍇ
		If SPRD.MaxRows > 1 Then
			ZAFC_N(5) = 5
			ZAFC_N(12) = 12
		End If
		
		Call ZAFC_SUB(Me)
		
	End Sub
	
	Private Sub KENSAKU_START_RTN()
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		Me.Refresh()
		
		SPRD.ReDraw = False
		
		CM_LNCNT = 0
		
		Call KENSAKU_RTN()
		
		If (CM_IMAX <> 0) Then
			SPRD.ROW = 1
			SPRD.Focus()
		End If
		
		SPRD.ReDraw = True
		
	End Sub
	
	Private Sub KENSAKU_RTN()
		
		Call SZ0415_STA_SUB(0)
		
		If CM_ERRSW = 1 Then
			CMDOFNC(0).Focus()
			Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
			Exit Sub
		End If
		
		'�Y���Ȃ�
		If CM_IMAX = n0 Then
			SPRD.MaxRows = 0
			SPRD.Enabled = False
		Else
			'�Y������
			SPRD.Enabled = True
		End If
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Refresh()
		
	End Sub
	
	Private Sub SPRD_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SPRD.ClickEvent
		
		CM_IDX = SPRD.ActiveRow
		
	End Sub
	
	Private Sub SPRD_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SPRD.DblClick
		
		Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
		
	End Sub
	
	Private Sub SPRD_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD.Enter
		
		If CUR_NO = N050 Then Exit Sub
		CUR_NO = N050
		
		'�`�F�b�N
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		'�m��
		LST_NO = CUR_NO
		
		'--- �t�@���N�V�������b�Z�[�W
		'    Call FUNCSET_RTN
		
		CM_IDX = SPRD.ActiveRow
		
	End Sub
	
	Private Sub SPRD_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SPRD.KeyDownEvent
		'�X�v���b�h���ł̃L�[����
		
		If eventArgs.Shift <> n0 Then Exit Sub
		
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Escape
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
				
			Case System.Windows.Forms.Keys.F5
				CMDOFNC(5).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				
			Case System.Windows.Forms.Keys.F12, System.Windows.Forms.Keys.Return '�I���i���s���I���j
				CMDOFNC(12).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
		End Select
		
	End Sub
	
	Private Sub SZ0415_STA_SUB(ByRef OP As Short)
		'�����J�n
		Dim WK_VAL As Object
		
		'����\���̏ꍇ
		CM_I = 0
		CM_IMAX = 0
		
		'�����A�P���ڎ�o��
		Call SZ0415_GET_SUB(4)
		If CM_ERRSW = 1 Then '�G���[
			Exit Sub
		End If
		If CM_ENDSW = 1 Then '�Y���Ȃ�
			CM_IMAX = 0
			CM_ENDSW = 0
			Exit Sub
		End If
		
STA_DISP: 
		
		CM_IMAX = 0
		SPRD.MaxRows = 0
		
		Do Until SZ0415RS.EOF
			CM_IMAX = CM_IMAX + 1
			SPRD.MaxRows = CM_IMAX
			
			Call SPRD.SetText(1, CM_IMAX, VB6.Format(SZ0415RS.rdoColumns("BK1").Value, "000000"))
			'UPGRADE_WARNING: �I�u�W�F�N�g WK_VAL �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			WK_VAL = SZ0415RS.rdoColumns("BK4").Value
			Call SPRD.SetText(2, CM_IMAX, WK_VAL)
			
			SZ0415RS.MoveNext()
		Loop 
		
	End Sub
	
	Private Sub SZ0415_ERR_SUB()
		
		ZAER_FID = "RAZ99"
		ZAER_KN = 1
		ZAER_NO.Value = "COM0050"
		Call ZAER_SUB()
		CM_ERRSW = 1
		
	End Sub
	
	Private Sub SZ0415_GET_SUB(ByRef iType As Short)
		
		Select Case iType
			Case 1
				SZ0415SELGE.rdoParameters("BK1_F").Value = "0"
				SZ0415SELGE.rdoParameters("BK1_T").Value = "9"
				SZ0415SELGE.rdoParameters("BK2").Value = "1" '�啪��
			Case 2
				SZ0415SELGE.rdoParameters("BK1_F").Value = SZ0415_DAI_CODES.Value & "0"
				SZ0415SELGE.rdoParameters("BK1_T").Value = SZ0415_DAI_CODES.Value & "9"
				SZ0415SELGE.rdoParameters("BK2").Value = "2" '������
			Case 3
				SZ0415SELGE.rdoParameters("BK1_F").Value = SZ0415_CHU_CODES.Value & "00"
				SZ0415SELGE.rdoParameters("BK1_T").Value = SZ0415_CHU_CODES.Value & "99"
				SZ0415SELGE.rdoParameters("BK2").Value = "3" '������
			Case 4
				SZ0415SELGE.rdoParameters("BK1_F").Value = SZ0415_SHO_CODES.Value & "00"
				SZ0415SELGE.rdoParameters("BK1_T").Value = SZ0415_SHO_CODES.Value & "99"
				SZ0415SELGE.rdoParameters("BK2").Value = "4" '�ו���
		End Select
		
		On Error Resume Next
		SZ0415RSSW = "SZ0415SELGE"
		SZ0415RS = SZ0415SELGE.OpenResultset()
		RDO_STATUS = B_STATUS(SZ0415RS)
		Select Case RDO_STATUS
			Case 0
			Case 24
				'�Y���Ȃ�
				CM_ENDSW = 1
				Exit Sub
			Case Else
				'�G���[
				Call SZ0415_ERR_SUB()
				Exit Sub
		End Select
		
		On Error GoTo 0
		
	End Sub
	
	
	Private Function IPROCHK() As Boolean
		
		Dim i As Short
		
		IPROCHK = True
		CM_ERRSW = F_OFF
		CM_ERRSW = F_OFF
		
		If CUR_NO = LST_NO Then Exit Function
		
		Select Case LST_NO
			Case N010
				Call IPROCHK_N010()
			Case N020
				Call IPROCHK_N020()
			Case N030
				Call IPROCHK_N030()
			Case N040
				'Call IPROCHK_N040
			Case N050
				'Call IPROCHK_N050
				
		End Select
		'########## �װ�̔��� ##########
		If CM_ERRSW = F_ERR Then
			IPROCHK = False
			NXT_NO = LST_NO
			'Call FOCUS_SET
		End If
		
	End Function
	
	Private Sub IPROCHK_N010()
		
		If CMB010.SelectedIndex = -1 Then
			SZ0415_DAI_CODES.Value = ""
			SZ0415_CHU_CODES.Value = ""
			SZ0415_SHO_CODES.Value = ""
			CMB020.Items.Clear()
			CMB030.Items.Clear()
			SPRD.MaxRows = 0
			Exit Sub
		End If
		
		If SZ0415_DAI_CODES.Value = CStr(VB6.GetItemData(CMB010, CMB010.SelectedIndex)) Then
		Else
			SZ0415_DAI_CODES.Value = CStr(VB6.GetItemData(CMB010, CMB010.SelectedIndex))
			SZ0415_CHU_CODES.Value = ""
			SZ0415_SHO_CODES.Value = ""
			CMB020.Items.Clear()
			CMB030.Items.Clear()
			Call COMBO_INIT_SZ0415(CMB020, 2)
			'Call COMBO_INIT_SZ0415(CMB030, 3)
			SPRD.MaxRows = 0
		End If
		
	End Sub
	
	Private Sub IPROCHK_N020()
		
		If CMB020.SelectedIndex = -1 Then
			SZ0415_CHU_CODES.Value = ""
			SZ0415_SHO_CODES.Value = ""
			CMB030.Items.Clear()
			SPRD.MaxRows = 0
			Exit Sub
		End If
		
		If SZ0415_CHU_CODES.Value = CStr(VB6.GetItemData(CMB020, CMB020.SelectedIndex)) Then
		Else
			SZ0415_CHU_CODES.Value = CStr(VB6.GetItemData(CMB020, CMB020.SelectedIndex))
			SZ0415_SHO_CODES.Value = ""
			CMB030.Items.Clear()
			Call COMBO_INIT_SZ0415(CMB030, 3)
			SPRD.MaxRows = 0
		End If
		
	End Sub
	
	Private Sub IPROCHK_N030()
		
		If CMB030.SelectedIndex = -1 Then
			SZ0415_SHO_CODES.Value = ""
			SPRD.MaxRows = 0
			Exit Sub
		End If
		
		If SZ0415_SHO_CODES.Value = CStr(VB6.GetItemData(CMB030, CMB030.SelectedIndex)) Then
		Else
			SZ0415_SHO_CODES.Value = CStr(VB6.GetItemData(CMB030, CMB030.SelectedIndex))
			SPRD.MaxRows = 0
		End If
		
	End Sub
End Class