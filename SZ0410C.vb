Option Strict Off
Option Explicit On
Module SZ0410CBAS
	
	'       �P�ʃ}�X�^
	Private SZM0110SELCDU As RDO.rdoQuery
	Private bSZM0110Ready As Boolean
	'       �n�o�}�X�^
	Private COM0070SELCDU As RDO.rdoQuery
	Private bCOM0070Ready As Boolean
	
	Private SYSDATERS As RDO.rdoResultset
	Private CDUSYSDATE As Date
	
	'   �������ޖ⍇������          QUE_FIND�Ŏg�p
	Public SEL_FIND As String '   �⍇���߂�l
	Public SEL_TYPE As String '   �⍇�����
	Public SEL_CODE As String '   �⏕����
	Public SEL_CODE2 As String '   �⏕���ނ��̂Q
	
	
	
	
	
	
	
	
	Public Function CduLoadUNIT(ByRef cdKaisha As String, ByRef cdJigyo As String, ByRef cBox As System.Windows.Forms.ComboBox) As Short
		
		Dim nUnit As Short '   �ǂ݂��܂ꂽ�P�ʐ�
		Dim strUNIT As String '   �P�ʓǂ݂��ݍ�ƈ�
		Dim bFirst As Boolean
		
		Call CduPrepSZM0110()
		
		cBox.Items.Clear()
		nUnit = 0
		bFirst = True
		Erase Tani_T 'A-CUST-20100823
		
		Do While True
			On Error Resume Next ' (�װ���ׯ��)
			If bFirst Then
				
				SZM0110SELCDU.rdoParameters("Inc_code").Value = cdKaisha
				SZM0110SELCDU.rdoParameters("jg_code").Value = cdJigyo
				
				If SZM0110RSSW <> "SZM0110SELCDU" Or ReQue = False Then
					SZM0110RS = SZM0110SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
					SZM0110RSSW = "SZM0110SELCDU"
				Else
					SZM0110RS.Requery()
				End If
				bFirst = False
				
			Else
				SZM0110RS.MoveNext()
			End If
			On Error GoTo 0
			
			
			Select Case B_STATUS(SZM0110RS) ' (SQL���s�ð���̕]��)
				Case 0 '   ����
					strUNIT = SZM0110RS.rdoColumns("t_name").Value
					cBox.Items.Add(strUNIT)
					nUnit = nUnit + 1
					'A-CUST-20100823 Start
					'UPGRADE_WARNING: �z�� Tani_T �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
					ReDim Preserve Tani_T(nUnit)
					Tani_T(nUnit) = strUNIT
					'A-CUST-20100823 End
					
				Case 24 '   EOF
					Exit Do
					
				Case Else
					ERRSW = F_ERR
					Exit Do
					
			End Select
			
		Loop 
		
		cBox.Items.Add(Space(4))
		TaniCnt = nUnit 'A-CUST-20100823
		nUnit = nUnit + 1
		CduLoadUNIT = nUnit
		
	End Function
	
	Private Sub CduPrepSZM0110()
		
		'   Schema���̎擾  SZM0110
		MKKCMN.ZAEV_FNO = "SZM0110"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0110_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0110_FILE.NAME = ""
		
		'   ���Ə��}�X�^��QUERY�쐬
		SQL = "Select t_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0110_FILE.NAME) & "SZM0110"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? AND Del_Flg <> '1' "
		
		On Error Resume Next
		SZM0110SELCDU = ZACN_RCN.CreateQuery("SZM0110SELCDU", SQL)
		SZM0110SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0110"
			GoTo PREP_SZM0110_ERR
		End If
		On Error GoTo 0
		
		SZM0110SELCDU.rdoParameters(0).NAME = "Inc_code"
		SZM0110SELCDU.rdoParameters(1).NAME = "jg_code"
		SZM0110SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0110SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0110SELCDU.rdoParameters(0).Size = 2
		SZM0110SELCDU.rdoParameters(1).Size = 4
		
		bSZM0110Ready = True
		Exit Sub
		
PREP_SZM0110_ERR: 
		CduPrepJigyo = F_ERR
		
	End Sub
	
	Public Function DateSlashed(ByRef sdate As String) As String
		
		If Trim(sdate) = "" Then
			DateSlashed = Space(10)
		Else
			DateSlashed = Mid(sdate, 1, 4) & "/" & Mid(sdate, 5, 2) & "/" & Mid(sdate, 7, 2)
		End If
		
	End Function
	
	Public Function CduServerDate() As Date
		
		'�T�[�o�[���t�擾
		SQL = "select sysdate SYMD from dual"
		On Error Resume Next
		SYSDATERS = ZACN_RCN.OpenResultset(SQL) '99/12/09 ADD KTT YOSHINO
		Select Case B_STATUS(SYSDATERS)
			Case 0
				CDUSYSDATE = SYSDATERS.rdoColumns("SYMD").Value
			Case Else
				ERRSW = F_ERR
				ZAER_KN = 1
				ZAER_NO.Value = ""
				ZAER_MS.Value = ""
				Call ZAER_SUB()
		End Select
		
		CduServerDate = CDUSYSDATE
		
	End Function
	
	
	
	'   �n�o�R�[�h�ɂ��I�y���[�^�����擾����B
	'   ��Ɂ@CduPrepOper()�����s���邱�ƁB
	'   COM0070.bas��Project�ɒǉ����邱��
	Public Function CduDecodeOper(ByRef cdKaisha As String, ByRef cdOper As String, ByRef nmOper As String) As Short
		
		If Not bCOM0070Ready Then
			CduDecodeOper = F_DUM
			MsgBox("���s�菇�G���[�FCduPrepOper()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeOper")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeOper = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		COM0070SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		COM0070SELCDU.rdoParameters("op_code").Value = cdOper
		
		On Error Resume Next ' (�װ���ׯ��)
		If COM0070RSSW <> "COM0070SELCDU" Or ReQue = False Then
			COM0070RS = COM0070SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			COM0070RSSW = "COM0070SELCDU"
		Else
			COM0070RS.Requery()
		End If
		
		Select Case B_STATUS(COM0070RS) ' (SQL���s�ð���̕]��)
			Case 0
				nmOper = COM0070RS.rdoColumns("op_name").Value
			Case 24
				CduDecodeOper = F_END
				nmOper = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeOper = F_END
				nmOper = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeOper"
				
				ZAER_KN = 1
				ZAER_NO.Value = "COM0070"
				Call ZAER_SUB()
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	'
	'   �n�o�R�[�h���f�R�[�h���邽�߂�Query����
	'   COM0070.bas��Project�ɒǉ����邱��
	Public Function CduPrepOper() As Object
		
		'   Schema���̎擾  COM0070
		MKKCMN.ZAEV_FNO = "COM0070"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			COM0070_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    COM0070_FILE.NAME = ""
		
		'   �I�y���[�^�}�X�^��QUERY�쐬
		SQL = "Select op_name "
		SQL = SQL & " from "
		SQL = SQL & RTrim(COM0070_FILE.NAME) & "COM0070"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND op_code = ? "
		
		On Error Resume Next
		COM0070SELCDU = ZACN_RCN.CreateQuery("COM0070SELCDU", SQL)
		COM0070SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "COM0070"
			GoTo PREP_COM0070_ERR
		End If
		On Error GoTo 0
		
		COM0070SELCDU.rdoParameters(0).NAME = "Inc_code"
		COM0070SELCDU.rdoParameters(1).NAME = "op_code"
		COM0070SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		COM0070SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		COM0070SELCDU.rdoParameters(0).Size = 2
		COM0070SELCDU.rdoParameters(1).Size = 6
		
		bCOM0070Ready = True
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepOper �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepOper = F_OFF
		
		Exit Function
		
PREP_COM0070_ERR: 
		'UPGRADE_WARNING: �I�u�W�F�N�g CduPrepOper �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		CduPrepOper = F_ERR
		
	End Function
	
	
	Public Function QUE_HINBAN() As Short
		
		Dim lRet As Integer
		
		SZ0420.SZ0420_KAISYA = WKB010 '  ��к���
		SZ0420.SZ0420_JGCODE = WKB020 '  ���Ə�����
		SZ0420.SZ0420_BSCODE = "" '  ��������
		SZ0420.SZ0420_CHECK = 0 '  �����׸� �i1.�����L�� �P�ȊO���������j
		SZ0420.SZ0420_TOP = VB6.PixelsToTwipsY(SZ0410FRM.Top) '  �e���(TOP)
		SZ0420.SZ0420_LEFT = VB6.PixelsToTwipsX(SZ0410FRM.Left) '  �e���(LEFT)
		SZ0420.SZ0420_HEIGHT = VB6.PixelsToTwipsY(SZ0410FRM.Height) '  �e���(HEIGHT)
		SZ0420.SZ0420_WIDTH = VB6.PixelsToTwipsX(SZ0410FRM.Width) '  �e���(WIDTH)
		SZ0420.SZ0420_POS = 1 '�@�\���ʒu(0.���� 1.���� 2.�E�� 3.���� 4.�E�� )
		SZ0420.SZ0420_RCN = ZACN_RCN '  �ڑ����̈��n��
		SZ0420.SZ0420_TIME = CInt(WG_TIMEOUT) '  RDO��ѱ�ĕb��
		
		lRet = SZ0420.SZ0420_SUB
		If lRet = 0 Then
			'        WKB030 = SZ0420.SZ0420_LCODE
			SZ0410FRM.IMTX030.Text = SZ0420.SZ0420_LCODE
			QUE_HINBAN = 0
		Else
			QUE_HINBAN = -1
		End If
		
	End Function
	
	
	Public Function QUE_KAMOKU() As Short
		
		'    Select Case iOpt
		'        Case 1
		'            SEL_TYPE = "KAMOKUCHU"
		'        Case 2
		'            SEL_TYPE = "KAMOKUSHO"
		'            SEL_CODE = CHU
		'        Case Else
		'            QUE_KAMOKU = -2
		'            Exit Function
		'    End Select
		'
		'    SZ0410GFRM.Show vbModal
		'
		'    If SEL_FIND <> "" Then
		'        Select Case iOpt
		'            Case 1
		'                KB.hiyou_k_code1 = SEL_FIND
		'
		'            Case 2
		'                KB.hiyou_k_code2 = SEL_FIND
		'        End Select
		'        QUE_KAMOKU = 0
		'
		'    Else
		'        QUE_KAMOKU = -1
		'    End If
		Dim iRet As Short
		
		CM9550.CM9550_LEFT = VB6.PixelsToTwipsX(SZ0410FRM.Left)
		CM9550.CM9550_TOP = VB6.PixelsToTwipsY(SZ0410FRM.Top)
		CM9550.CM9550_HEIGHT = VB6.PixelsToTwipsY(SZ0410FRM.Height)
		CM9550.CM9550_WIDTH = VB6.PixelsToTwipsX(SZ0410FRM.Width)
		CM9550.CM9550_RCN = ZACN_RCN
		CM9550.CM9550_TIME = CInt(WG_TIMEOUT)
		CM9550.CM9550_POS = 1
		CM9550.CM9550_INC_CODE = WKB010
		CM9550.CM9550_INC_NAME = WKB010DSP
		CM9550.CM9550_JG_CODE = WKB020
		CM9550.CM9550_JG_NAME = WKB020DSP
		iRet = CM9550.CM9550_SUB
		If iRet = 0 Then
			'    Debug.Assert Len(CM9550.CM9550_KMCODE) > 0
			SZ0410FRM.IMTX130(1).Text = Right(CM9550.CM9550_KMCODE, 3)
			SZ0410FRM.IMTX140(1).Text = CM9550.CM9550_KSCODE
			QUE_KAMOKU = 0
		Else
			QUE_KAMOKU = -1
		End If
		
	End Function
	
	'
	
	
	Public Function QUE_GYOSHA() As Short
		
		Dim iRet As Short
		
		SZ0310.SZ0310_KAISYA = WKB010
		SZ0310.SZ0310_HONSITEN = WKB020
		SZ0310.SZ0310_LEFT = VB6.PixelsToTwipsX(SZ0410FRM.Left)
		SZ0310.SZ0310_TOP = VB6.PixelsToTwipsY(SZ0410FRM.Top)
		SZ0310.SZ0310_HEIGHT = VB6.PixelsToTwipsY(SZ0410FRM.Height)
		SZ0310.SZ0310_WIDTH = VB6.PixelsToTwipsX(SZ0410FRM.Width)
		SZ0310.SZ0310_RCN = ZACN_RCN
		SZ0310.SZ0310_TIME = CInt(WG_TIMEOUT)
		SZ0310.SZ0310_POS = 1
		iRet = SZ0310.SZ0310_SUB
		'    MsgBox "iRet=" & iRet
		'    MsgBox "Lcode= " & SZ0310.SZ0310_LCODE
		If iRet = 0 Then
			System.Diagnostics.Debug.Assert(Len(SZ0310.SZ0310_LCODE) > 0, "")
			KB.g_gentei_code = SZ0310.SZ0310_LCODE
			QUE_GYOSHA = 0
		Else
			QUE_GYOSHA = -1
		End If
		
		
	End Function
	
	'UPGRADE_NOTE: str �� str_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Public Function ZeroFill(ByRef str_Renamed As String, ByRef lengt As Short) As String
		
		Dim strTrim As String
		Dim strZero As String
		
		strTrim = Trim(str_Renamed)
		strZero = New String("0", lengt)
		ZeroFill = Right(strZero & strTrim, lengt)
		
	End Function
	
	'UPGRADE_NOTE: str �� str_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Public Function ZeroTrim(ByRef str_Renamed As String) As String
		
		Dim strTrim As String
		Dim lenOrg As Short
		Dim IDX As Short
		
		lenOrg = Len(str_Renamed)
		strTrim = Trim(str_Renamed)
		lenOrg = Len(strTrim)
		
		IDX = 1
		Do While IDX <= lenOrg
			If Mid(strTrim, IDX, 1) <> "0" Then Exit Do
			IDX = IDX + 1
		Loop 
		
		ZeroTrim = Mid(strTrim, IDX, lenOrg - IDX + 1)
		
	End Function
	
	'UPGRADE_NOTE: str �� str_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Public Function ZeroTrim9(ByRef str_Renamed As String) As String
		
		'   Zero Surpress like ZZZZZ9
		
		Dim strTrim As String
		Dim lenOrg As Short
		Dim IDX As Short
		
		lenOrg = Len(str_Renamed)
		strTrim = Trim(str_Renamed)
		lenOrg = Len(strTrim)
		
		IDX = 1
		Do While IDX <= (lenOrg - 1)
			If Mid(strTrim, IDX, 1) <> "0" Then Exit Do
			IDX = IDX + 1
		Loop 
		
		ZeroTrim9 = Mid(strTrim, IDX, lenOrg - IDX + 1)
		
	End Function
End Module