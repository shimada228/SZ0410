Option Strict Off
Option Explicit On
Module ARQLGMBAS
	Public ZALGM_INC_CODE As New VB6.FixedLengthString(2) '会社ｺｰﾄﾞ
	Public ZALGM_JG_CODE As New VB6.FixedLengthString(4) '事業所コード
	Public ZALGM_SYS_KBN As New VB6.FixedLengthString(1) 'システム区分
	Public ZALGM_S_DAY As New VB6.FixedLengthString(8) '処理日付
	Public ZALGM_S_TIME As New VB6.FixedLengthString(6) '処理時刻
	Public ZALGM_OP_CODE As New VB6.FixedLengthString(6) 'オペレータコード
	Public ZALGM_PGID As New VB6.FixedLengthString(8) 'プログラムＩＤ（半角大文字）
	Public ZALGM_SH_KBN As New VB6.FixedLengthString(1) '処理区分
	Public ZALGM_SH_NAIYO As New VB6.FixedLengthString(30) '処理内容
	Public ZALGM_GNFLG As New VB6.FixedLengthString(1) '減額フラグ
	
	Public ZALGM_ERR As New VB6.FixedLengthString(1)
	Public ZALGM_KO_NAIYO As New VB6.FixedLengthString(30) '更新内容
	Const ZALGM_ERR_POINT As String = "ZALGM_SUB"
	
	'------------------------------------------------------------
	'【関数名】 ログ出力サブルーチン
	'
	'【機  能】 ログファイルにログを追加する。履歴ファイルの存在するものは履歴ファイルを更新する。
	'
	'【戻り値】 無し
	'
	'------------------------------------------------------------
	Sub ZALGM_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		
		'<< エラーフラグのクリア >>
		ZALGM_ERR.Value = "0"
		
		'<< 履歴ファイルの更新処理 >>
		
		'履歴ファイル更新ストアドプロシージャ－ＣＡＬＬ
		
		'Update 1999/12/15 REP START TOP
		'処理区分が新規登録の時は、履歴更新行わない。
		If ZALGM_SH_KBN.Value <> "3" Then
			'Update 1999/12/15 ADD E N D TOP
			Call ZALGM_RKSTRD_SUB(ZALGM_UARCN)
			If ZALGM_ERR.Value = "1" Then
				GoTo ZALGERR
			End If
			'Update 1999/12/15 REP START TOP
		Else
			ZALGM_KO_NAIYO.Value = Space(30)
		End If
		'Update 1999/12/15 REP E N D TOP
		
		'<< 処理ログファイルにログを追加する >>
		
		'処理ログファイル更新ストアドプロシージャ－ＣＡＬＬ
		Call ZALGM_LGSTRD_SUB(ZALGM_UARCN)
		If ZALGM_ERR.Value = "1" Then
			GoTo ZALGERR
		End If
		
		Exit Sub
		
ZALGERR: 
		'ログ出力失敗メッセージ
		MsgBox("ログが出力されませんでした" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		
	End Sub
	
	Sub ZALGM_RKSTRD_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		Dim CCM9030PRO As RDO.rdoQuery
		Dim MSG_NAME As String
		Dim CSZ_FILE_NAME As String
		Dim RETCD1 As Short
		Dim RETCD2 As String
		Dim RETCD3 As String
		Dim RETCD4 As String
		
		'CCM9030
		MKKCMN.ZAEV_FNO = "CCM9030"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ZALGM_ERR.Value = "1"
			MsgBox("プロシージャのスキーマ定義エラーです" & Chr(13) & "CCM9030" & ZALGM_ERR_POINT, 48, "")
			Exit Sub
		Else
			CSZ_FILE_NAME = RTrim(MKKCMN.ZAEV_FNM)
		End If
		
		'プロシージャの定義
		On Error GoTo STRD_ERR
		
		'UPDATE 1999/12/21
		'    SQL = "begin "
		'    SQL = SQL & RTrim$(CSZ_FILE_NAME) & "CCM9030"
		'    SQL = SQL & "(?,?,?,?,?,?,?); end;"
		SQL = "{CALL "
		SQL = SQL & RTrim(CSZ_FILE_NAME) & "CCM9030( ?,?,?,?,?,?,? )}"
		'UPDATE 1999/12/21
		
		CCM9030PRO = ZALGM_UARCN.CreateQuery("CCM9030PRO", SQL)
		Select Case B_STATUS
			Case 0
			Case Else
				MsgBox("プロシージャの定義エラーです" & Chr(13) & "CCM9030" & ZALGM_ERR_POINT, 48, "")
				ZALGM_ERR.Value = "1"
				Exit Sub
		End Select
		
		'IN
		CCM9030PRO.rdoParameters(0).NAME = "kosin_key" '更新キー文字列
		CCM9030PRO.rdoParameters(1).NAME = "sys_kbn" 'システム区分
		CCM9030PRO.rdoParameters(2).NAME = "prg_id" 'PRGID
		
		'OUT
		CCM9030PRO.rdoParameters(3).NAME = "RETCD1" '状態ステータス（０：正常、１：エラー）
		CCM9030PRO.rdoParameters(4).NAME = "RETCD2" 'トレース用
		CCM9030PRO.rdoParameters(5).NAME = "RETCD3" 'エラー内容
		CCM9030PRO.rdoParameters(6).NAME = "RETCD4" '更新ｷｰ
		
		CCM9030PRO.rdoParameters(0).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(1).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(2).Direction = RDO.DirectionConstants.rdParamInput
		CCM9030PRO.rdoParameters(3).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(4).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(5).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9030PRO.rdoParameters(6).Direction = RDO.DirectionConstants.rdParamOutput
		
		CCM9030PRO.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9030PRO.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		
		CCM9030PRO.rdoParameters(0).Value = ZALGM_SH_NAIYO.Value '更新キー文字列
		CCM9030PRO.rdoParameters(1).Value = ZALGM_SYS_KBN.Value 'システム区分
		CCM9030PRO.rdoParameters(2).Value = ZALGM_PGID.Value 'ＰＲＧＩＤ
		
		'プロシージャの実行
		CCM9030PRO.QueryTimeout = 0
		CCM9030PRO.Execute()
		
		RETCD1 = CCM9030PRO.rdoParameters(3).Value '状態ステータス（０：正常、１：エラー）
		RETCD2 = CCM9030PRO.rdoParameters(4).Value 'トレース用
		RETCD3 = CCM9030PRO.rdoParameters(5).Value 'エラー内容
		ZALGM_KO_NAIYO.Value = CCM9030PRO.rdoParameters(6).Value '登録ｷｰ
		
		If RETCD1 = -1 Then '０以外エラー
			MsgBox("履歴ファイルの更新でエラーが起こりました" & Chr(13) & RETCD3 & Chr(13) & ZALGM_ERR_POINT, 48, "")
			ZALGM_ERR.Value = "1"
		End If
		
		'クエリークローズ
		CCM9030PRO.Close()
		'UPGRADE_NOTE: オブジェクト CCM9030PRO をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		CCM9030PRO = Nothing
		Exit Sub
		
STRD_ERR: 
		MsgBox("その他のエラー" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		'UPGRADE_NOTE: オブジェクト CCM9030PRO をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		CCM9030PRO = Nothing
		
	End Sub
	
	Sub ZALGM_LGSTRD_SUB(ByRef ZALGM_UARCN As RDO.rdoConnection)
		Dim CCM9020PRO As RDO.rdoQuery
		Dim MSG_NAME As String
		Dim CSZ_FILE_NAME As String
		Dim RETCD1 As Short
		Dim RETCD2 As String
		Dim RETCD3 As String
		
		'CCM9020
		MKKCMN.ZAEV_FNO = "CCM9020"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ZALGM_ERR.Value = "1"
			MsgBox("プロシージャのスキーマ定義エラーです" & Chr(13) & "CCM9020" & ZALGM_ERR_POINT, 48, "")
			Exit Sub
		Else
			CSZ_FILE_NAME = RTrim(MKKCMN.ZAEV_FNM)
		End If
		
		'プロシージャの定義
		On Error GoTo STRD_ERR
		'UPDATE 1999/12/21
		'    SQL = "begin "
		'    SQL = SQL & RTrim$(CSZ_FILE_NAME) & "CCM9020"
		'    SQL = SQL & "(?,?,?,?,?,?,?,?,?,?,?,?,?,?); end;"
		SQL = "{CALL "
		SQL = SQL & RTrim(CSZ_FILE_NAME) & "CCM9020( ?,?,?,?,?,?,?,?,?,?,?,?,?,? )}"
		'UPDATE 1999/12/21
		
		CCM9020PRO = ZALGM_UARCN.CreateQuery("CCM9020PRO", SQL)
		Select Case B_STATUS
			Case 0
			Case Else
				MsgBox("プロシージャの定義エラーです" & Chr(13) & "CCM9020" & ZALGM_ERR_POINT, 48, "")
				ZALGM_ERR.Value = "1"
				Exit Sub
		End Select
		
		'IN
		CCM9020PRO.rdoParameters(0).NAME = "Inc_code" '会社コード
		CCM9020PRO.rdoParameters(1).NAME = "jg_code" '事業所コード
		CCM9020PRO.rdoParameters(2).NAME = "sys_kbn" 'システム区分
		CCM9020PRO.rdoParameters(3).NAME = "s_day" '処理日付
		CCM9020PRO.rdoParameters(4).NAME = "s_time" '処理時刻
		CCM9020PRO.rdoParameters(5).NAME = "op_code" 'オペレータコード
		CCM9020PRO.rdoParameters(6).NAME = "shori_sikibetu" '処理識別
		CCM9020PRO.rdoParameters(7).NAME = "shori_kbn" '処理区分
		CCM9020PRO.rdoParameters(8).NAME = "shori_naiyo1" '処理内容１
		CCM9020PRO.rdoParameters(9).NAME = "kosin_naiyo2" '更新内容２
		CCM9020PRO.rdoParameters(10).NAME = "gn_flg" '減額フラグ
		
		'OUT
		CCM9020PRO.rdoParameters(11).NAME = "RETCD1" '状態ステータス（０：正常、１：エラー）
		CCM9020PRO.rdoParameters(12).NAME = "RETCD2" 'トレース用
		CCM9020PRO.rdoParameters(13).NAME = "RETCD3" 'エラー内容
		
		CCM9020PRO.rdoParameters(0).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(1).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(2).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(3).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(4).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(5).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(6).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(7).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(8).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(9).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(10).Direction = RDO.DirectionConstants.rdParamInput
		CCM9020PRO.rdoParameters(11).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9020PRO.rdoParameters(12).Direction = RDO.DirectionConstants.rdParamOutput
		CCM9020PRO.rdoParameters(13).Direction = RDO.DirectionConstants.rdParamOutput
		
		CCM9020PRO.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(7).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(8).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(9).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(10).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(11).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		CCM9020PRO.rdoParameters(12).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		CCM9020PRO.rdoParameters(13).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		
		CCM9020PRO.rdoParameters(0).Value = ZALGM_INC_CODE.Value '会社コード
		CCM9020PRO.rdoParameters(1).Value = ZALGM_JG_CODE.Value '事業所コード
		CCM9020PRO.rdoParameters(2).Value = ZALGM_SYS_KBN.Value 'システム区分
		CCM9020PRO.rdoParameters(3).Value = ZALGM_S_DAY.Value '処理日付
		CCM9020PRO.rdoParameters(4).Value = ZALGM_S_TIME.Value '処理時刻
		CCM9020PRO.rdoParameters(5).Value = ZALGM_OP_CODE.Value 'オペレータコード
		CCM9020PRO.rdoParameters(6).Value = ZALGM_PGID.Value '処理識別
		CCM9020PRO.rdoParameters(7).Value = ZALGM_SH_KBN.Value '処理区分
		CCM9020PRO.rdoParameters(8).Value = ZALGM_SH_NAIYO.Value '処理内容１
		CCM9020PRO.rdoParameters(9).Value = ZALGM_KO_NAIYO.Value '更新内容２
		CCM9020PRO.rdoParameters(10).Value = ZALGM_GNFLG.Value '減額フラグ
		
		'プロシージャの実行
		CCM9020PRO.QueryTimeout = 0
		CCM9020PRO.Execute()
		
		RETCD1 = CCM9020PRO.rdoParameters(11).Value '状態ステータス（０：正常、１：エラー）
		RETCD2 = CCM9020PRO.rdoParameters(12).Value 'トレース用
		RETCD3 = CCM9020PRO.rdoParameters(13).Value 'エラー内容
		If RETCD1 = -1 Then '０以外エラー
			MsgBox("処理ログの更新でエラーが起こりました" & Chr(13) & RETCD3 & Chr(13) & ZALGM_ERR_POINT, 48, "")
			ZALGM_ERR.Value = "1"
		End If
		
		'クエリークローズ
		CCM9020PRO.Close()
		'UPGRADE_NOTE: オブジェクト CCM9020PRO をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		CCM9020PRO = Nothing
		Exit Sub
		
STRD_ERR: 
		MsgBox("その他のエラー" & Chr(13) & ZALGM_ERR_POINT, 48, "")
		ZALGM_ERR.Value = "1"
		'UPGRADE_NOTE: オブジェクト CCM9020PRO をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		CCM9020PRO = Nothing
		
	End Sub
End Module