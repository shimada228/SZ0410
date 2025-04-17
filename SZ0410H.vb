Option Strict Off
Option Explicit On
Module SZ0410HBAS
	
	'       �Ȗڕ��ރ}�X�^
	Private SZM0040SELH As RDO.rdoQuery
	Private bSZM0040Ready As Boolean
	
	'       �������ރ}�X�^
	Private SZM0050SELX As RDO.rdoQuery
	Private bSZM0050Ready As Boolean
	
	'       �ȖڑΉ��}�X�^
	Private SZM0170SELH As RDO.rdoQuery
	Private bSZM0170Ready As Boolean
	
	'       ���L�}�X�^
	Private MCM97SELH As RDO.rdoQuery
	Private bMCM97Ready As Boolean
	
	Private SZM0060SELCDU As RDO.rdoQuery
	Private bSZM0060ready As Boolean
	Private SZM0070SELCDU As RDO.rdoQuery
	Private bSZM0070ready As Boolean
	Private SZM0080SELCDU As RDO.rdoQuery
	Private bSZM0080ready As Boolean
	
	Private qCOM0050SEL As RDO.rdoQuery
	Private bCOM0050SEL As Boolean
	Private qCOM0050RS As RDO.rdoResultset
	Private qCOM0050RSSW As String
	
	Private LKPFlag As Boolean
	'���ރ}�X�^                                 '02/05/28 ADD
	Private SZM0055SEL As RDO.rdoQuery '02/05/28 ADD
	Private bSZM0055ready As Boolean '02/05/28 ADD
	
	Public Sub SetLookupMode(ByRef bFlag As Boolean)
		LKPFlag = bFlag
	End Sub
	
	Public Function DecodeKamBunrui(ByRef cdKaisha As String, ByRef cdJigyo As String, ByRef codePlus As String) As String
		'   codePlus��code1(3)+code2(4)
		
		Dim cd3 As String
		Dim cd4 As String
		Dim strReturn As String
		
		DecodeKamBunrui = ""
		
		If Not bSZM0040Ready Then
			MsgBox("���s�菇�G���[�FPrepKamBunrui()���ɁI", MsgBoxStyle.OKOnly, "DecodeKamBunrui")
			Exit Function
		End If
		
		cd3 = Mid(codePlus, 1, 3)
		cd4 = Mid(codePlus, 4, 4)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0040SELH.rdoParameters("Inc_code").Value = cdKaisha
		SZM0040SELH.rdoParameters("jg_code").Value = cdJigyo
		SZM0040SELH.rdoParameters("code1").Value = cd3
		SZM0040SELH.rdoParameters("code2").Value = cd4
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0040RSSW <> "SZM0040SELH" Or ReQue = False Then
			SZM0040RS = SZM0040SELH.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0040RSSW = "SZM0040SELH"
		Else
			SZM0040RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0040RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0040RS.rdoColumns("del_flg").Value >= "1" Then
					strReturn = ""
				Else
					strReturn = SZM0040RS.rdoColumns("kamoku_name").Value
				End If
			Case 24
				strReturn = ""
				
			Case Else
				strReturn = ""
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "RSZM0040"
				''''Call ZAER_SUB
				Exit Function
		End Select
		
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		DecodeKamBunrui = strReturn
		
	End Function
	
	Public Sub PrepKamBunrui()
		
		'   Schema���̎擾  SZM0040
		MKKCMN.ZAEV_FNO = "SZM0040"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0040_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0040_FILE.NAME = ""
		
		'   �I�y���[�^�}�X�^��QUERY�쐬
		SQL = "Select kamoku_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0040_FILE.NAME) & "SZM0040"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & "   AND jg_code = ? "
		SQL = SQL & "   AND code1 = ? "
		SQL = SQL & "   AND code2 = ? "
		
		On Error Resume Next
		SZM0040SELH = ZACN_RCN.CreateQuery("SZM0040SELH", SQL)
		SZM0040SELH.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0040"
			
		End If
		On Error GoTo 0
		
		SZM0040SELH.rdoParameters(0).NAME = "Inc_code"
		SZM0040SELH.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0040SELH.rdoParameters(0).Size = 2
		SZM0040SELH.rdoParameters(1).NAME = "jg_code"
		SZM0040SELH.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0040SELH.rdoParameters(1).Size = 4
		SZM0040SELH.rdoParameters(2).NAME = "code1"
		SZM0040SELH.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0040SELH.rdoParameters(2).Size = 3
		SZM0040SELH.rdoParameters(3).NAME = "code2"
		SZM0040SELH.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0040SELH.rdoParameters(3).Size = 4
		
		bSZM0040Ready = True
		
		
	End Sub
	
	Public Function DecodeGYOSHA(ByRef cdKaisha As String, ByRef cdJigyo As String, ByRef cdGYOSHA As String) As String
		
		If Not bMCM97Ready Then
			MsgBox("���s�菇�G���[�FPrepGYOSHA()���ɁI", MsgBoxStyle.OKOnly, "DecodeGYOSHA")
			Exit Function
		End If
		
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		MCM97SELH.rdoParameters("Inc_code").Value = cdKaisha
		MCM97SELH.rdoParameters("jg_code").Value = cdJigyo
		MCM97SELH.rdoParameters("gyo_code").Value = cdGYOSHA
		
		On Error Resume Next ' (�װ���ׯ��)
		If MCM97RSSW <> "MCM97SELH" Or ReQue = False Then
			MCM97RS = MCM97SELH.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			MCM97RSSW = "MCM97SELH"
		Else
			MCM97RS.Requery()
		End If
		
		Select Case B_STATUS(MCM97RS) ' (SQL���s�ð���̕]��)
			Case 0
				DecodeGYOSHA = MCM97RS.rdoColumns("cm97008").Value
				
			Case 24
				DecodeGYOSHA = ""
				
			Case Else
				DecodeGYOSHA = ""
				''''ERRSW = F_ERR
				''''ZAER_KN = 1
				''''ZAER_NO = "RMCM97"
				''''Call ZAER_SUB
				
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Public Sub PrepGYOSHA()
		
		'   Schema���̎擾  MCM97
		MKKCMN.ZAEV_FNO = "MCM97"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			MCM97_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    MCM97_FILE.NAME = "SHIIRE."
		
		'   ���L�Ǝ҃}�X�^��QUERY�쐬
		SQL = ""
		SQL = SQL & "Select CM97008 from "
		SQL = SQL & RTrim(MCM97_FILE.NAME) & "MCM97GYOS"
		SQL = SQL & " WHERE CM97001 = ? "
		SQL = SQL & " AND CM97002 = ? "
		SQL = SQL & " AND CM97003 = ? "
		On Error Resume Next
		MCM97SELH = ZACN_RCN.CreateQuery("MCM97SELH", SQL)
		MCM97SELH.QueryTimeout = ZACN_TIME
		On Error GoTo 0
		
		
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "MCM97"
			Exit Sub
			
		End If
		MCM97SELH.rdoParameters(0).NAME = "Inc_code"
		MCM97SELH.rdoParameters(1).NAME = "jg_code"
		MCM97SELH.rdoParameters(2).NAME = "gyo_code"
		MCM97SELH.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM97SELH.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM97SELH.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		MCM97SELH.rdoParameters(0).Size = 2
		MCM97SELH.rdoParameters(1).Size = 4
		MCM97SELH.rdoParameters(2).Size = 6
		
		bMCM97Ready = True
		
	End Sub
	
	Public Function TaiouKamoku(ByRef cdKaisha As String, ByRef cdJigyo As String, ByRef cdHIYO3byte As String, ByRef cdHIYO6byte As String, ByRef cdKAMURI As String, ByRef cdKAMSHO As String, ByRef cdKAMMAT As String, ByRef cdKAMMIT As String) As Short
		
		Dim strC3 As String
		Dim strC6 As String
		
		If Not bSZM0170Ready Then
			TaiouKamoku = F_DUM
			MsgBox("���s�菇�G���[�FTaiouKamokkuPrep()���ɁI", MsgBoxStyle.OKOnly, "TaiouKamoku")
			Exit Function
		End If
		
		TaiouKamoku = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0170SELH.rdoParameters("Inc_code").Value = cdKaisha
		SZM0170SELH.rdoParameters("jg_code").Value = cdJigyo
		SZM0170SELH.rdoParameters("hiCHU").Value = cdHIYO3byte
		SZM0170SELH.rdoParameters("hiSHO").Value = cdHIYO6byte
		
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0170RSSW <> "SZM0170SELH" Or ReQue = False Then
			SZM0170RS = SZM0170SELH.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0170RSSW = "SZM0170SELH"
		Else
			SZM0170RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0170RS) ' (SQL���s�ð���̕]��)
			Case 0
				strC3 = SZM0170RS.rdoColumns("uri_code1").Value
				strC6 = SZM0170RS.rdoColumns("uri_code2").Value
				cdKAMURI = strC3 & strC6
				strC3 = SZM0170RS.rdoColumns("sho_code1").Value
				strC6 = SZM0170RS.rdoColumns("sho_code2").Value
				cdKAMSHO = strC3 & strC6
				strC3 = SZM0170RS.rdoColumns("matu_code1").Value
				strC6 = SZM0170RS.rdoColumns("matu_code2").Value
				cdKAMMAT = strC3 & strC6
				strC3 = SZM0170RS.rdoColumns("mi_code1").Value
				strC6 = SZM0170RS.rdoColumns("mi_code2").Value
				cdKAMMIT = strC3 & strC6
			Case 24
				TaiouKamoku = F_END
				cdKAMURI = Space(9)
				cdKAMSHO = Space(9)
				cdKAMMAT = Space(9)
				cdKAMMIT = Space(9)
				
			Case Else
				TaiouKamoku = F_ERR
				cdKAMURI = Space(9)
				cdKAMSHO = Space(9)
				cdKAMMAT = Space(9)
				cdKAMMIT = Space(9)
				ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "RSZM0170"
				''''Call ZAER_SUB
				Exit Function
		End Select
		
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Public Sub TaiouKamokuPrep()
		
		'   Schema���̎擾  SZM0170
		MKKCMN.ZAEV_FNO = "SZM0170"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0170_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0170_FILE.NAME = ""
		
		
		SQL = "Select "
		
		SQL = SQL & "uri_code1,"
		SQL = SQL & "uri_code2,"
		SQL = SQL & "sho_code1,"
		SQL = SQL & "sho_code2,"
		SQL = SQL & "matu_code1,"
		SQL = SQL & "matu_code2,"
		SQL = SQL & "mi_code1,"
		SQL = SQL & "mi_code2, del_flg "
		
		'   �I�y���[�^�}�X�^��QUERY�쐬
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0170_FILE.NAME) & "SZM0170"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & "   AND jg_code = ? "
		SQL = SQL & "   AND hi_code1 = ? "
		SQL = SQL & "   AND hi_code2 = ? "
		
		On Error Resume Next
		SZM0170SELH = ZACN_RCN.CreateQuery("SZM0170SELH", SQL)
		SZM0170SELH.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0170"
			
		End If
		On Error GoTo 0
		
		SZM0170SELH.rdoParameters(0).NAME = "Inc_code"
		SZM0170SELH.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170SELH.rdoParameters(0).Size = 2
		SZM0170SELH.rdoParameters(1).NAME = "jg_code"
		SZM0170SELH.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170SELH.rdoParameters(1).Size = 4
		SZM0170SELH.rdoParameters(2).NAME = "hiCHU"
		SZM0170SELH.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170SELH.rdoParameters(2).Size = 3
		SZM0170SELH.rdoParameters(3).NAME = "hiSHO"
		SZM0170SELH.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170SELH.rdoParameters(3).Size = 6
		
		bSZM0170Ready = True
		
	End Sub
	
	
	
	'   �啪�ރR�[�h�ɂ��啪�ޖ����擾����B
	'   ��Ɂ@CduPrepDAIBunrui()�����s���邱�ƁB
	'   SZM0060.bas��Project�ɒǉ����邱��
	Public Function CduDecodeDAIBunrui(ByRef cdKaisha As String, ByRef DAI As String, ByRef Dname As String) As Short
		
		If Not bSZM0060ready Then
			CduDecodeDAIBunrui = F_DUM
			MsgBox("���s�菇�G���[�FCduPrepDAIBunrui()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeDAIBunrui")
			Exit Function
		End If
		
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeDAIBunrui = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0060SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		SZM0060SELCDU.rdoParameters("d_code").Value = DAI
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0060RSSW <> "SZM0060SELCDU" Or ReQue = False Then
			SZM0060RS = SZM0060SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0060RSSW = "SZM0060SELCDU"
		Else
			SZM0060RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0060RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0060RS.rdoColumns("del_flg").Value >= "1" Then
					Dname = ""
					CduDecodeDAIBunrui = F_END
				Else
					Dname = SZM0060RS.rdoColumns("d_name").Value
				End If
				
			Case 24
				Dname = ""
				CduDecodeDAIBunrui = F_END
				''''ENDSW = F_END
			Case Else
				CduDecodeDAIBunrui = F_END
				Dname = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeDAIBunrui"
				
				''''ZAER_KN = 1
				''''ZAER_NO = "SZM0060"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Public Function CduPrepDAIBunrui() As Short
		
		'   Schema���̎擾  SZM0060
		MKKCMN.ZAEV_FNO = "SZM0060"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			SZM0060_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0060_FILE.NAME = ""
		
		'   ���Ə��}�X�^��QUERY�쐬
		SQL = "Select d_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0060_FILE.NAME) & "SZM0060"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND d_code = ? "
		
		On Error Resume Next
		SZM0060SELCDU = ZACN_RCN.CreateQuery("SZM0060SELCDU", SQL)
		SZM0060SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0060"
			GoTo PREP_SZM0060_ERR
		End If
		On Error GoTo 0
		
		SZM0060SELCDU.rdoParameters(0).NAME = "Inc_code"
		SZM0060SELCDU.rdoParameters(1).NAME = "d_code"
		SZM0060SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0060SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0060SELCDU.rdoParameters(0).Size = 2
		SZM0060SELCDU.rdoParameters(1).Size = 4
		
		bSZM0060ready = True
		CduPrepDAIBunrui = F_OFF
		
		Exit Function
		
PREP_SZM0060_ERR: 
		CduPrepDAIBunrui = F_ERR
		
	End Function
	
	'   �啪�ރR�[�h�A�����ރR�[�h�ɂ�蒆���ޖ����擾����B
	'   ��Ɂ@CduPrepCHUBunrui()�����s���邱�ƁB
	'   SZM0070.bas��Project�ɒǉ����邱��
	Public Function CduDecodeCHUBunrui(ByRef cdKaisha As String, ByRef DAI As String, ByRef CHU As String, ByRef Cname As String) As Short
		
		If Not bSZM0070ready Then
			CduDecodeCHUBunrui = F_DUM
			MsgBox("���s�菇�G���[�FCduPrepCHUBunrui()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeCHUBunrui")
			Exit Function
		End If
		
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeCHUBunrui = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0070SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		SZM0070SELCDU.rdoParameters("d_code").Value = DAI
		SZM0070SELCDU.rdoParameters("c_code").Value = CHU
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0070RSSW <> "SZM0070SELCDU" Or ReQue = False Then
			SZM0070RS = SZM0070SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0070RSSW = "SZM0070SELCDU"
		Else
			SZM0070RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0070RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0070RS.rdoColumns("del_flg").Value >= "1" Then
					CduDecodeCHUBunrui = F_END
					Cname = ""
				Else
					Cname = SZM0070RS.rdoColumns("c_name").Value
				End If
				
			Case 24
				CduDecodeCHUBunrui = F_END
				Cname = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeCHUBunrui = F_END
				Cname = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeCHUBunrui"
				
				''''ZAER_KN = 1
				''''ZAER_NO = "SZM0070"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Public Function CduPrepCHUBunrui() As Short
		
		'   Schema���̎擾  SZM0070
		MKKCMN.ZAEV_FNO = "SZM0070"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			SZM0070_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0070_FILE.NAME = ""
		
		'   ���Ə��}�X�^��QUERY�쐬
		SQL = "Select c_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0070_FILE.NAME) & "SZM0070"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND d_code = ? "
		SQL = SQL & " AND c_code = ? "
		
		On Error Resume Next
		SZM0070SELCDU = ZACN_RCN.CreateQuery("SZM0070SELCDU", SQL)
		SZM0070SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0070"
			GoTo PREP_SZM0070_ERR
		End If
		On Error GoTo 0
		
		SZM0070SELCDU.rdoParameters(0).NAME = "Inc_code"
		SZM0070SELCDU.rdoParameters(1).NAME = "d_code"
		SZM0070SELCDU.rdoParameters(2).NAME = "c_code"
		SZM0070SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0070SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0070SELCDU.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0070SELCDU.rdoParameters(0).Size = 2
		SZM0070SELCDU.rdoParameters(1).Size = 4
		SZM0070SELCDU.rdoParameters(2).Size = 4
		
		bSZM0070ready = True
		CduPrepCHUBunrui = F_OFF
		
		Exit Function
		
PREP_SZM0070_ERR: 
		CduPrepCHUBunrui = F_ERR
		
	End Function
	
	'   �啪�ރR�[�h�A�����ރR�[�h�A�����ރR�[�h�ɂ�菬���ޖ����擾����B
	'   ��Ɂ@CduPrepSHOBunrui()�����s���邱�ƁB
	'   SZM0080.bas��Project�ɒǉ����邱��
	Public Function CduDecodeSHOBunrui(ByRef cdKaisha As String, ByRef DAI As String, ByRef CHU As String, ByRef SHO As String, ByRef SName As String) As Short
		
		If Not bSZM0080ready Then
			CduDecodeSHOBunrui = F_DUM
			MsgBox("���s�菇�G���[�FCduPrepSHOBunrui()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeSHOBunrui")
			Exit Function
		End If
		
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeSHOBunrui = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0080SELCDU.rdoParameters("Inc_code").Value = cdKaisha
		SZM0080SELCDU.rdoParameters("d_code").Value = DAI
		SZM0080SELCDU.rdoParameters("c_code").Value = CHU
		SZM0080SELCDU.rdoParameters("s_code").Value = SHO
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0080RSSW <> "SZM0080SELCDU" Or ReQue = False Then
			SZM0080RS = SZM0080SELCDU.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0080RSSW = "SZM0080SELCDU"
		Else
			SZM0080RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0080RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0080RS.rdoColumns("del_flg").Value >= "1" Then
					CduDecodeSHOBunrui = F_END
					SName = ""
				Else
					SName = SZM0080RS.rdoColumns("s_name").Value
				End If
				
			Case 24
				CduDecodeSHOBunrui = F_END
				SName = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeSHOBunrui = F_END
				SName = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "ERR", vbOKOnly, "CduDecodeSHOBunrui"
				
				ZAER_KN = 1
				ZAER_NO.Value = "SZM0080"
				Call ZAER_SUB()
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	Public Function CduPrepSHOBunrui() As Short
		
		'   Schema���̎擾  SZM0080
		MKKCMN.ZAEV_FNO = "SZM0080"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			SZM0080_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0080_FILE.NAME = ""
		
		'   ���Ə��}�X�^��QUERY�쐬
		SQL = "Select s_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0080_FILE.NAME) & "SZM0080"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND d_code = ? "
		SQL = SQL & " AND c_code = ? "
		SQL = SQL & " AND s_code = ? "
		
		On Error Resume Next
		SZM0080SELCDU = ZACN_RCN.CreateQuery("SZM0080SELCDU", SQL)
		SZM0080SELCDU.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0080"
			GoTo PREP_SZM0080_ERR
		End If
		On Error GoTo 0
		
		SZM0080SELCDU.rdoParameters(0).NAME = "Inc_code"
		SZM0080SELCDU.rdoParameters(1).NAME = "d_code"
		SZM0080SELCDU.rdoParameters(2).NAME = "c_code"
		SZM0080SELCDU.rdoParameters(3).NAME = "s_code"
		SZM0080SELCDU.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0080SELCDU.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0080SELCDU.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0080SELCDU.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0080SELCDU.rdoParameters(0).Size = 2
		SZM0080SELCDU.rdoParameters(1).Size = 4
		SZM0080SELCDU.rdoParameters(2).Size = 4
		SZM0080SELCDU.rdoParameters(3).Size = 4
		
		bSZM0080ready = True
		CduPrepSHOBunrui = F_OFF
		
		Exit Function
		
PREP_SZM0080_ERR: 
		CduPrepSHOBunrui = F_ERR
		
	End Function
	
	'
	'   ���ʕ����}�X�^ COM0050 ��蕔�������f�R�[�h����
	Public Function CduDecodeBUSHO(ByRef cdBUSHO As String) As String
		'
		'
		If Not bCOM0050SEL Then
			CduDecodeBUSHO = CStr(F_DUM)
			MsgBox("���s�菇�G���[�FCduPrepBUSHO()���ɁI", MsgBoxStyle.OKOnly, "CduDecodeBuSHO")
			Exit Function
		End If
		
		'   �ŏ���OK�߂�l�Z�b�g
		CduDecodeBUSHO = CStr(F_OFF)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		qCOM0050SEL.rdoParameters("Inc_code").Value = WKB010
		qCOM0050SEL.rdoParameters("jg_code").Value = WKB020
		qCOM0050SEL.rdoParameters("bu_code").Value = cdBUSHO
		
		On Error Resume Next ' (�װ���ׯ��)
		If qCOM0050RSSW <> "qCOM0050SEL" Or ReQue = False Then
			qCOM0050RS = qCOM0050SEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			qCOM0050RSSW = "qCOM0050SEL"
		Else
			qCOM0050RS.Requery()
		End If
		
		Dim sy As String
		
		Select Case B_STATUS(qCOM0050RS) ' (SQL���s�ð���̕]��)
			Case 0
				sy = qCOM0050RS.rdoColumns("sy_bumon").Value
				If sy <> "0" Then
					CduDecodeBUSHO = "-"
				Else
					CduDecodeBUSHO = qCOM0050RS.rdoColumns("bu_name").Value
				End If
			Case 24
				CduDecodeBUSHO = ""
				''''ENDSW = F_END
			Case Else
				CduDecodeBUSHO = ""
				''''ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "MCM92"
				''''Call ZAER_SUB
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'   �����R�[�h���f�R�[�h���邽�߂�Query����
	'   COM0050.bas��Project�ɒǉ����邱��
	Public Sub CduPrepBUSHO()
		
		'   Schema���̎擾  MCM92
		MKKCMN.ZAEV_FNO = "COM0050"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			COM0050_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		'   ���ʕ����}�X�^��QUERY�쐬
		SQL = "Select bu_name, sy_bumon "
		SQL = SQL & " from "
		SQL = SQL & RTrim(COM0050_FILE.NAME) & "COM0050"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		SQL = SQL & " AND bu_code = ? "
		SQL = SQL & " AND si_shiyo_flg = '1' "
		'    SQL = SQL & " AND sy_bumon = '0' "
		'   �d���g�p����Ɣ�W�v����̏����ǉ�      2000/01/27  SZ0410-10
		'   �W�v����̏����͂����ł͏��O            2000/02/01
		
		On Error Resume Next
		qCOM0050SEL = ZACN_RCN.CreateQuery("qCOM0050SEL", SQL)
		qCOM0050SEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "COM0050"
			
		End If
		On Error GoTo 0
		
		qCOM0050SEL.rdoParameters(0).NAME = "Inc_code"
		qCOM0050SEL.rdoParameters(1).NAME = "jg_code"
		qCOM0050SEL.rdoParameters(2).NAME = "bu_code"
		qCOM0050SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qCOM0050SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		qCOM0050SEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		qCOM0050SEL.rdoParameters(0).Size = 2
		qCOM0050SEL.rdoParameters(1).Size = 4
		qCOM0050SEL.rdoParameters(2).Size = 4
		
		bCOM0050SEL = True
		
		
	End Sub
	
	
	'   �������ރ}�X�^���猟�����ޖ��̂��擾����B
	'   SZM0050.bas���v���W�F�N�g�Ɏ�肱�ނ���
	'
	Public Function DecodeFIND(ByRef cdFIND As String) As String
		
		Dim strName As String
		
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0050SELX.rdoParameters("Inc_code_WP").Value = WKB010
		SZM0050SELX.rdoParameters("jg_code_WP").Value = WKB020
		SZM0050SELX.rdoParameters("find_code_WP").Value = cdFIND
		
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0050RSSW <> "SZM0050SELX" Or ReQue = False Then
			SZM0050RS = SZM0050SELX.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0050RSSW = "SZM0050SELX"
		Else
			SZM0050RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0050RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0050RS.rdoColumns("del_flg").Value >= "1" Then
					strName = ""
				Else
					strName = SZM0050RS.rdoColumns("find_name").Value
				End If
				
			Case 24
				strName = ""
				
			Case Else
				strName = ""
				ENDSW = F_END
				ERRSW = F_ERR
				
				''''ZAER_KN = 1
				''''ZAER_NO = "RSZM0050"
				''''Call ZAER_SUB
				Exit Function
		End Select
		DecodeFIND = strName
		
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Public Sub PrepFind()
		
		'   Schema���̎擾  SZM0050
		MKKCMN.ZAEV_FNO = "SZM0050"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0050_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0050_FILE.NAME = ""
		
		'   �������ރ}�X�^��QUERY�쐬
		SQL = ""
		SQL = SQL & "Select Inc_code, jg_code, find_code, find_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0050_FILE.NAME) & "SZM0050"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		SQL = SQL & " AND find_code = ? "
		
		On Error Resume Next
		SZM0050SELX = ZACN_RCN.CreateQuery("SZM0050SELX", SQL)
		SZM0050SELX.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0050"
			
		End If
		On Error GoTo 0
		
		SZM0050SELX.rdoParameters(0).NAME = "Inc_code_WP"
		SZM0050SELX.rdoParameters(1).NAME = "jg_code_WP"
		SZM0050SELX.rdoParameters(2).NAME = "find_code_WP"
		SZM0050SELX.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0050SELX.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0050SELX.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0050SELX.rdoParameters(0).Size = 2
		SZM0050SELX.rdoParameters(1).Size = 4
		SZM0050SELX.rdoParameters(2).Size = 4
		
		bSZM0050Ready = True
		
	End Sub
	'02/05/28 ADD START
	
	Public Function PrepBunrui() As Short
		
		'   Schema���̎擾  SZM0055
		MKKCMN.ZAEV_FNO = "SZM0055"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Function
		Else
			SZM0055_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		'   ���ރ}�X�^��QUERY�쐬
		SQL = "select bun_name, del_flg "
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0055_FILE.NAME) & "SZM0055"
		SQL = SQL & " where Inc_code = ? "
		SQL = SQL & " and bun_code = ? "
		
		On Error Resume Next
		SZM0055SEL = ZACN_RCN.CreateQuery("SZM0055SEL", SQL)
		SZM0055SEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "SZM0055"
			GoTo PREP_SZM0055_ERR
		End If
		On Error GoTo 0
		
		SZM0055SEL.rdoParameters(0).NAME = "Inc_code"
		SZM0055SEL.rdoParameters(1).NAME = "bun_code"
		SZM0055SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0055SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0055SEL.rdoParameters(0).Size = 2
		SZM0055SEL.rdoParameters(1).Size = 4
		
		bSZM0055ready = True
		PrepBunrui = F_OFF
		
		Exit Function
		
PREP_SZM0055_ERR: 
		PrepBunrui = F_ERR
		
	End Function
	
	Public Function DecodeBUNRUI(ByRef cdBUNRUI As String, ByRef rBUNRUI_NAME As String) As Short
		'���ރ}�X�^���番�ޖ����擾����B
		Dim strName As String
		
		DecodeBUNRUI = F_OFF
		'���ރR�[�h�����͂���Ă��Ȃ����͌������Ȃ�
		If RTrim(cdBUNRUI) = "" Then
			rBUNRUI_NAME = ""
			Exit Function
		End If
		
		'�����ݒ�
		SZM0055SEL.rdoParameters("Inc_code").Value = WKB010 '��ЃR�[�h
		SZM0055SEL.rdoParameters("bun_code").Value = cdBUNRUI '���ރR�[�h
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0055RSSW <> "SZM0055SEL" Or ReQue = False Then
			SZM0055RS = SZM0055SEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0055RSSW = "SZM0055SEL"
		Else
			SZM0055RS.Requery()
		End If
		
		Select Case B_STATUS(SZM0055RS) ' (SQL���s�ð���̕]��)
			Case 0
				If LKPFlag And SZM0055RS.rdoColumns("del_flg").Value >= "1" Then
					strName = ""
				Else
					strName = SZM0055RS.rdoColumns("bun_name").Value
				End If
			Case 24
				strName = ""
				DecodeBUNRUI = F_END
			Case Else
				strName = ""
				DecodeBUNRUI = F_END
				ENDSW = F_END
				ERRSW = F_ERR
				Exit Function
		End Select
		rBUNRUI_NAME = strName
		
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	'02/05/28 ADD END
End Module