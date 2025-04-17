Option Strict Off
Option Explicit On
Module SZ0415BAS
	'******************************************************************
	'*  システム名     ： 三井観光開発株式会社  仕入管理システム
	'*  プログラム名   ： ＪＡＮ商品分類検索
	'*  プログラムＩＤ ： SZ0415
	'*  作  成  者     ： SSP@目黒
	'******************************************************************
	
	Structure RECT
		'UPGRADE_NOTE: Left は Left_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right は Right_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	'Public InitFlg      As Boolean
	'Public lpRectSave   As RECT
	'Public vfaRowWidth()    As Long
	'Public tabNo        As Integer
	
	Public SZ0415SELGE As RDO.rdoQuery '初回
	
	Public SZ0415RS As RDO.rdoResultset
	Public SZ0415RSSW As String
	Public SZ0415INVSW As String
	
	Public CM_ERRSW As Short 'ERROR 時にSW=1
	Public CM_INTSW As Short
	
	Public SZ0415_TOPS As Integer '親画面(TOP)
	Public SZ0415_LEFTS As Integer '親画面(LEFT)
	Public SZ0415_HEIGHTS As Integer '親画面(HEIGHT)
	Public SZ0415_WIDTHS As Integer '親画面(WIDTH)
	Public SZ0415_PS As Short '表示位置(0.中央 1.左上 2.右上 3.左下 4.右下 )
	Public SZ0415_TIMES As Integer 'RDOﾀｲﾑｱｳﾄ秒数
	
	Public SZ0415_DAI_CODES As New VB6.FixedLengthString(1) '大分類
	Public SZ0415_CHU_CODES As New VB6.FixedLengthString(2) '中分類
	Public SZ0415_SHO_CODES As New VB6.FixedLengthString(4) '小分類
	
	Public SZ0415_OLDCOD1 As New VB6.FixedLengthString(2) '前回開始コード
	Public SZ0415_SPRD As Short '前スプレッドアクティブＲＯＷ
	
	'選択結果
	Public SZ0415_SEL_CODES As New VB6.FixedLengthString(6) '選択
	Public SZ0415_KBN As Short '有効区分 0:有効 -1:無効
	Public SZ0415_PRESW As Short 'PREPARE判断用ｽｲｯﾁ
	
	' ＳＷＩＴＣＨエリア
	Public CM_DSP1SW As Short
	Public SZ0415_TOP As Integer '親画面(TOP)
	Public SZ0415_LEFT As Integer '親画面(LEFT)
	Public SZ0415_HEIGHT As Integer '親画面(HIGHT)
	Public SZ0415_WIDTH As Integer '親画面(WIDTH)
	Public SZ0415_POS As Short '表示位置(0.中央 1.左上 2.右上 3.左下 4.右下 )
	Public SZ0415_RCN As RDO.rdoConnection 'ＤＢ情報
	Public SZ0415_DB As Short 'ＤＢ名
	Public SZ0415_TIME As Integer 'RDOタイムアウト秒数
	
	'引数
	Public SZ0415_DAI_CODE As New VB6.FixedLengthString(1) '大分類
	Public SZ0415_CHU_CODE As New VB6.FixedLengthString(2) '中分類
	Public SZ0415_SHO_CODE As New VB6.FixedLengthString(4) '小分類
	
	'戻り値
	Public SZ0415_SEL_CODE As String '選択 細分類
	
	Public Function SZ0415_SUB() As Short
		
		CM_INTSW = F_OFF ' 初期表示RTN実行
		CM_ERRSW = F_OFF
		
		SZ0415_TOPS = SZ0415_TOP ' 親画面(TOP)
		SZ0415_LEFTS = SZ0415_LEFT ' 親画面(LEFT)
		SZ0415_HEIGHTS = SZ0415_HEIGHT ' 親画面(HEIGHT)
		SZ0415_WIDTHS = SZ0415_WIDTH ' 親画面(WIDTH)
		SZ0415_PS = SZ0415_POS
		SZ0415_TIMES = SZ0415_TIME
		
		SZ0415_DAI_CODES.Value = SZ0415_DAI_CODE.Value
		SZ0415_CHU_CODES.Value = SZ0415_CHU_CODE.Value
		SZ0415_SHO_CODES.Value = SZ0415_SHO_CODE.Value
		
		'初期処理
		Call INIT_SZ0415_RTN()
		If CM_ERRSW = F_ERR Then
			SZ0415_SUB = n1
			Exit Function
		End If
		
		SZ0415FRM.ShowDialog()
		SZ0415_SEL_CODE = SZ0415_SEL_CODES.Value
		
		SZ0415_SUB = SZ0415_KBN
		
	End Function
	
	Private Sub ENV_RTN()
		
		'ＪＡＮ分類マスタ
		MKKCMN.ZAEV_FNO = "JAN_BUNRUI"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			CM_ERRSW = F_ERR
			Exit Sub
		Else
			JAN_BUNRUI_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
	End Sub
	
	Public Sub INIT_SZ0415_RTN()
		Dim Ret As Short
		Dim Rq As String
		
		If SZ0415_PRESW = F_ON Then
			CM_INTSW = F_ON
			Exit Sub
		End If
		
		
		'スキーマ名取得
		Call ENV_RTN()
		If CM_ERRSW = F_ERR Then Exit Sub
		
		' Prepareコマンド設定
		Call PREP_SZ0415()
		If CM_ERRSW = F_ERR Then
			Exit Sub
		End If
		
		'ワークの初期化
		SZ0415_PRESW = F_ON
		CM_INTSW = F_OFF
		SZ0415_SPRD = n1
		SZ0415_DAI_CODES.Value = ""
		SZ0415_CHU_CODES.Value = ""
		SZ0415_SHO_CODES.Value = ""
		
		
	End Sub
	
	Private Sub PREP_SZ0415()
		
		SQL = "Select "
		SQL = SQL & " NVL(BK1, ' ') BK1"
		SQL = SQL & ",NVL(BK2, ' ') BK2"
		SQL = SQL & ",NVL(BK3, ' ') BK3"
		SQL = SQL & ",NVL(BK4, ' ') BK4"
		SQL = SQL & ",NVL(BK5, ' ') BK5"
		SQL = SQL & ",NVL(BK6, ' ') BK6"
		SQL = SQL & ",NVL(BK7, ' ') BK7"
		SQL = SQL & ",NVL(BK8, ' ') BK8"
		SQL = SQL & " From "
		SQL = SQL & RTrim(JAN_BUNRUI_FILE.NAME) & "JAN_BUNRUI "
		SQL = SQL & " WHERE BK1 >= ? "
		SQL = SQL & "   AND BK1 <= ? "
		SQL = SQL & "   AND BK2  = ? "
		SQL = SQL & "Order By BK1 "
		
		On Error Resume Next
		SZ0415SELGE = ZACN_RCN.CreateQuery("SZ0415SELGE", SQL)
		SZ0415SELGE.QueryTimeout = SZ0415_TIMES
		If B_STATUS <> 0 Then
			GoTo PREP_ERR
		End If
		On Error GoTo 0
		
		SZ0415SELGE.rdoParameters(0).NAME = "BK1_F"
		SZ0415SELGE.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0415SELGE.rdoParameters(0).Size = 6
		SZ0415SELGE.rdoParameters(1).NAME = "BK1_T"
		SZ0415SELGE.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0415SELGE.rdoParameters(1).Size = 6
		SZ0415SELGE.rdoParameters(2).NAME = "BK2"
		SZ0415SELGE.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZ0415SELGE.rdoParameters(2).Size = 1
		
		
		Exit Sub
		
PREP_ERR: 
		
		ZAER_FID = "JAN_BUNRUI"
		ZAER_KN = 1
		ZAER_NO.Value = "JAN_BUNRUI"
		Call ZAER_SUB()
		CM_ERRSW = F_ERR
		SZ0415_KBN = -1
		On Error GoTo 0
		
	End Sub
End Module