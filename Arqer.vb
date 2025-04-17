Option Strict Off
Option Explicit On
Module ARQERBAS
	'*
	'*  MKK用 エラーﾒｯｾｰｼﾞ表示サブルーチン
	'*
	'*  1999/11/24 MOD KTT YOSHINO
	'*  1999/12/09 MOD KTT YOSHINO
	'*             ZACN_RCN対応に戻す、ZAER0_SUBを元に戻す
	
	'Global ZAER_RCN As rdoConnection    'ADD 99/12/06 -> 99/12/09 DEL KTT YOSHINO
	Public ZAER_FID As String
	Public ZAER_CD As Short
	Public ZAER_KN As Short '0:SMILE, 1:ORACLE
	Public ZAER_NO As New VB6.FixedLengthString(7) 'KTT YOSHINO 3->7
	Public ZAER_MS As New VB6.FixedLengthString(24)
	
	' ファイル名
	Public ZAER_FILE As New VB6.FixedLengthString(40) '99/12/02 DEL KTT YOSHINO    '99/12/09 復活
	'Global ZAER_FILE As String          '99/12/02 ADD KTT YOSHINO  '99/12/09 DEL
	
	Public ZAER_ERR As New VB6.FixedLengthString(1)
	
	Const ZAER_MSG As String = "ﾒｯｾｰｼﾞ ﾅｼ"
	
	'ORACLE用ｴﾗｰ表示で追加
	Public ERR1 As Short
	Public ERR2 As Short
	
	'ＲＤＯ関連ワーク
	Public RdoErr As RDO.rdoError '97/07/03Add
	
	Public Sub ZAER_SUB()
		Dim SQLRet As Short
		Dim ERR_WMS1 As New VB6.FixedLengthString(48)
		Dim ERR_WMS2 As New VB6.FixedLengthString(48)
		Dim ERR_WMS3 As New VB6.FixedLengthString(48)
		Dim ERR_WFL As New VB6.FixedLengthString(24)
		Dim ERR_STS As New VB6.FixedLengthString(3)
		Dim ERR_WMSG As String
		Dim ERR_WORK1 As String
		Dim ERR_WORK2 As String
		Dim ERR_OLDFID As String
		
		'99/11/24 MOD START KTT-YOSHINO
		'    If ZAER_NO = String$(3, 0) Then    ' Asciiｺｰﾄﾞ All"0"のとき
		'        ZAER_NO = Space$(3)            ' スペースに変換
		'    End If                             ' （MsgBoxで改行表示が
		If ZAER_NO.Value = New String(Chr(0), 7) Then ' Asciiｺｰﾄﾞ All"0"のとき
			ZAER_NO.Value = Space(7) ' スペースに変換
		End If
		'99/11/24 MOD END KTT-YOSHINO
		
		If ZAER_MS.Value = New String(Chr(0), 24) Then '      おかしくなるため
			ZAER_MS.Value = Space(24) '         スペースに変換）
		End If '
		
		'   エラーメッセージ表示
		ZAER_ERR.Value = "0"
		ERR_WMS1.Value = Space(48)
		ERR_WMS2.Value = Space(48)
		ERR_WMS3.Value = Space(48)
		ERR_WFL.Value = Space(24)
		ERR_STS.Value = "0"
		
		'   先頭の"Ｎ"をとったﾌｧｲﾙIDを取得
		ERR_OLDFID = Mid(ZAER_FID, 2, 4)
		
		'   << ＯＲＡＣＬＥエラーセット>>
		If ZAER_KN = 1 Then
			'97/07/03Del ERR2 = GlueGetNumber("ERR2", 0)
			'97/07/03Del ERR1 = GlueGetNumber("ERR1", 0)
			'97/07/03Del SQLRet = execsql("get message for :ERR1: into :ERR3:")
			'97/07/03Del ERR_WORK1 = GlueGetString("ERR3", 0)
			ERR_WORK1 = "" '97/07/03Add
			For	Each RdoErr In RDOrdoEngine_definst.rdoErrors '97/07/03Add
				ERR_WORK1 = ERR_WORK1 & RDOrdoEngine_definst.rdoErrors(RdoErr).Description & vbCr '97/07/03Add
			Next RdoErr '97/07/03Add
			
		End If
		
		'   << ファイルデータ ＲＥＡＤ >>
		'97/07/03Del SQL = "FOR 1 SELECT * from " & ZAER_FILE & " WHERE "
		'SQL = "SELECT * from " & ZAER_FILE & "(NOLOCK) WHERE "                      '97/07/03Add
		'97/07/03Del SQL = SQL & ERR_OLDFID & "001 = '1' AND " & ERR_OLDFID & "002 = '" & ZAER_NO & "'"
		If ZACN_DB = ORCL Then '97/07/03Add
			SQL = "SELECT * from " & ZAER_FILE.Value & " WHERE " '97/07/03Add
		ElseIf ZACN_DB = SQLSRV Then  '97/07/03Add
			SQL = "SELECT * from " & ZAER_FILE.Value & "(NOLOCK) WHERE " '97/07/03Add
		End If
		SQL = SQL & ERR_OLDFID & "001 = 1 AND " & ERR_OLDFID & "002 = '" & ZAER_NO.Value & "'" '97/07/03Add
		'97/07/03Del SQLRet = execsql(SQL)
		'97/07/03Del If SQLRet = 0 Then
		'97/07/03Del     ERR_WFL = GlueGetString(ERR_OLDFID & "003", 0)
		'97/07/03Del Else
		'97/07/03Del     If SQLRet <> 24 Then
		' 「レコード無し」以外のエラーでエラーメッセージファイルが読めなかった
		'97/07/03Del ERR2 = GlueGetNumber("ERR2", 0)
		'97/07/03Del ERR1 = GlueGetNumber("ERR1", 0)
		'97/07/03Del SQLRet = execsql("get message for :ERR1: into :ERR3:")
		'97/07/03Del ERR_WORK2 = GlueGetString("ERR3", 0)
		'97/07/03Del ERR_WMSG = "ｴﾗｰﾒｯｾｰｼﾞﾌｧｲﾙ READｴﾗｰ STS = " & ERR_WORK2
		'97/07/03Del ERR_WMSG = ERR_WMSG & "  KEY = " & "1" & ZAER_NO
		'97/07/03Del MsgBox ERR_WMSG, 48, ""
		'97/07/03Del End If
		'97/07/03Del ERR_WFL = Space$(24)
		'97/07/03Del End If
		On Error Resume Next '97/07/03Add
		AZ99RS = ZACN_RCN.OpenResultset(SQL) '99/11/26 DEL YOSHINO    99/12/09 復活
		'    Set AZ99RS = ZAER_RCN.OpenResultset(SQL)      '99/12/06 ADD YOSHINO  99/12/09 DEL
		
		
		If Err.Number <> 0 Then '97/07/03Add
			ERR_WORK2 = "" '97/07/03Add
			For	Each RdoErr In RDOrdoEngine_definst.rdoErrors '97/07/03Add
				ERR_WORK2 = ERR_WORK2 & RDOrdoEngine_definst.rdoErrors(RdoErr).Description & vbCr '97/07/03Add
			Next RdoErr '97/07/03Add
			ERR_WMSG = "ｴﾗｰﾒｯｾｰｼﾞﾌｧｲﾙ READｴﾗｰ STS = " & ERR_WORK2 '97/07/03Add
			ERR_WMSG = ERR_WMSG & "  KEY = 1" & ZAER_NO.Value '97/07/03Add
			MsgBox(ERR_WMSG, 48, "") '97/07/03Add
		End If '97/07/03Add
		If AZ99RS.EOF = False Then '97/07/03Add
			ERR_WFL.Value = AZ99RS.rdoColumns(2).Value '97/07/03Add
		Else '97/07/03Add
			ERR_WFL.Value = Space(24) '97/07/03Add
		End If '97/07/03Add
		On Error GoTo 0 '97/07/03Add
		
		If ZAER_KN = 1 Then
			' ORACLEのエラー
			ERR_WMSG = ERR_WORK1 & Chr(10) & ZAER_NO.Value & RTrim(ERR_WFL.Value) & Chr(10) & RTrim(ZAER_MS.Value)
			MsgBox(ERR_WMSG, 48, "")
			
			' ログ出力
			ZALG_KBN.Value = "9"
			ZALG_NAIYO = ERR_WORK1 & RTrim(" " & ZAER_NO.Value & ERR_WFL.Value)
			Call ZALG_SUB()
			ZAER_ERR.Value = ZALG_ERR.Value
		Else
			' それ以外のエラー
			'   << メッセージデータ  ＲＥＡＤ >>
			'99/11/26 MOD START YOSHINO
			'        If ZAER_CD > 999 Then
			'           ZAER_CD = 8 & RightB(ZAER_CD, 2)
			'        End If
			If ZAER_CD > 9999 Then
				'UPGRADE_ISSUE: RightB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
				ZAER_CD = CShort(8 & RightB(ZAER_CD, 2))
			End If
			'99/11/26 MOD END YOSHINO
			
			'97/07/03Del SQL = "FOR 1 SELECT * from " & ZAER_FILE & " WHERE "
			'97/07/03Add SQL = SQL & ERR_OLDFID & "001 = '0' AND " & ERR_OLDFID & "002 = '" & Format(ZAER_CD, "000") & "'"
			If ZACN_DB = ORCL Then '97/07/03Add
				SQL = "SELECT * from " & ZAER_FILE.Value & " WHERE " '97/07/03Add
			ElseIf ZACN_DB = SQLSRV Then  '97/07/03Add
				SQL = "SELECT * from " & ZAER_FILE.Value & "(NOLOCK) WHERE " '97/07/03Add
			End If
			'99/11/26 MOD START YOSHINO
			'       SQL = SQL & ERR_OLDFID & "001 = 0 AND " & ERR_OLDFID & "002 = '" & Format(ZAER_CD, "000") & "'"  '97/07/03Add
			SQL = SQL & ERR_OLDFID & "001 = 0 AND " & ERR_OLDFID & "002 = '" & VB6.Format(ZAER_CD, "0000") & "'"
			'99/11/26 MOD END YOSHINO
			
			'97/07/03Del SQLRet = execsql(SQL)
			'97/07/03Del If SQLRet = 0 Then
			' エラーメッセージ取得成功
			'97/07/03Del     ERR_WMS = GlueGetString(ERR_OLDFID & "003", 0)
			'97/07/03Del     If GlueGetString(ERR_OLDFID & "006", 0) = "1" Then
			'97/07/03Del         ZAER_MS = GlueGetString(ERR_OLDFID & "007", 0)
			'97/07/03Del     End If
			'97/07/03Del     If GlueGetString(ERR_OLDFID & "005", 0) <> "0" Then
			'97/07/03Del         ZAER_CD = 0
			'97/07/03Del     End If
			'97/07/03Del     ERR_WMSG = Format(ZAER_CD, "000") + " " & RTrim$(ERR_WMS) & Chr$(10)
			'97/07/03Del     ERR_WMSG = ERR_WMSG & ZAER_NO & RTrim$(ERR_WFL) & Chr$(10) & RTrim$(ZAER_MS)
			'97/07/03Del ElseIf SQLRet = 24 Then
			' レコード無し
			'97/07/03Del     ZAER_CD = 0
			'97/07/03Del     ERR_WMS = ZAER_MSG
			'97/07/03Del     ERR_WMSG = Format(ZAER_CD, "000") + " " & RTrim$(ERR_WMS) & Chr$(10)
			'97/07/03Del     ERR_WMSG = ERR_WMSG & ZAER_NO & RTrim$(ERR_WFL) & Chr$(10) & RTrim$(ZAER_MS)
			'97/07/03Del Else
			' その他のエラー発生
			'97/07/03Del     ERR2 = GlueGetNumber("ERR2", 0)
			'97/07/03Del     ERR1 = GlueGetNumber("ERR1", 0)
			'97/07/03Del     SQLRet = execsql("get message for :ERR1: into :ERR3:")
			'97/07/03Del     ERR_WORK2 = GlueGetString("ERR3", 0)
			'97/07/03Del     ERR_WMSG = "ｴﾗｰﾒｯｾｰｼﾞﾌｧｲﾙ READｴﾗｰ STS = " & ERR_WORK2
			'97/07/03Del     ERR_WMSG = ERR_WMSG & "  KEY = " & "0" & Format(ZAER_CD, "000")
			'97/07/03Del End If
			On Error Resume Next '97/07/03Add
			AZ99RS = ZACN_RCN.OpenResultset(SQL) '99/11/26 DEL YOSHINO 97/07/03Add    '99/12/09 復活
			'        Set AZ99RS = ZAER_RCN.OpenResultset(SQL)    '99/12/06 ADD YOSHINO 97/07/03Add  '99/12/09 ADD
			
			If Err.Number = 0 Then '97/07/03Add
				If AZ99RS.EOF = False Then '97/07/03Add
					' エラーメッセージ取得成功
					ERR_WMS1.Value = AZ99RS.rdoColumns(2).Value '97/07/03Add
					ERR_WMS2.Value = AZ99RS.rdoColumns(3).Value
					ERR_WMS3.Value = AZ99RS.rdoColumns(4).Value
					If AZ99RS.rdoColumns(7).Value = "1" Then '97/07/03Add
						ZAER_MS.Value = AZ99RS.rdoColumns(8).Value '97/07/03Add
					End If '97/07/03Add
					If AZ99RS.rdoColumns(6).Value <> "0" Then '97/07/03Add
						ZAER_CD = 0 '97/07/03Add
					End If
					
					'99/11/26 MOD START YOSHINO
					'               ERR_WMSG = Format(ZAER_CD, "000") + " " & RTrim$(ERR_WMS1) & Chr$(10)    '97/07/03Add
					ERR_WMSG = VB6.Format(ZAER_CD, "0000") & " " & RTrim(ERR_WMS1.Value) & Chr(10)
					'99/11/26 MOD END   YOSHINO
					
					If Trim(ERR_WMS2.Value) <> "" Then '97/09/30ADD
						ERR_WMSG = ERR_WMSG & New String(" ", 6) & RTrim(ERR_WMS2.Value) & Chr(10) '97/09/30ADD
					End If
					If Trim(ERR_WMS3.Value) <> "" Then '97/09/30ADD
						ERR_WMSG = ERR_WMSG & Space(6) & RTrim(ERR_WMS3.Value) & Chr(10) '97/09/30ADD
					End If
					ERR_WMSG = ERR_WMSG & ZAER_NO.Value & RTrim(ERR_WFL.Value) & Chr(10) & RTrim(ZAER_MS.Value)
				Else
					' レコード無し
					ZAER_CD = 0
					ERR_WMS1.Value = ZAER_MSG
					'99/11/26 MOD START YOSHINO
					'               ERR_WMSG = Format(ZAER_CD, "000") + " " & RTrim$(ERR_WMS1) & Chr$(10)
					ERR_WMSG = VB6.Format(ZAER_CD, "0000") & " " & RTrim(ERR_WMS1.Value) & Chr(10)
					'99/11/26 MOD END   YOSHINO
					ERR_WMSG = ERR_WMSG & ZAER_NO.Value & RTrim(ERR_WFL.Value) & Chr(10) & RTrim(ZAER_MS.Value)
				End If
			Else
				' その他のエラー発生
				ERR_WORK2 = "" '97/07/03Add
				For	Each RdoErr In RDOrdoEngine_definst.rdoErrors '97/07/03Add
					ERR_WORK2 = ERR_WORK2 & RDOrdoEngine_definst.rdoErrors(RdoErr).Description & vbCr '97/07/03Add
				Next RdoErr '97/07/03Add
				ERR_WMSG = "ｴﾗｰﾒｯｾｰｼﾞﾌｧｲﾙ READｴﾗｰ STS = " & ERR_WORK2
				'99/11/26 MOD START YOSHINO
				'           ERR_WMSG = ERR_WMSG & "  KEY = " & "0" & Format(ZAER_CD, "000")
				ERR_WMSG = ERR_WMSG & "  KEY = " & "0" & VB6.Format(ZAER_CD, "0000")
				'99/11/26 MOD END YOSHINO
				
			End If
			MsgBox(ERR_WMSG, 48, "")
		End If
		
		ZAER_CD = 0
		ZAER_KN = 0
		ZAER_NO.Value = ""
		ZAER_MS.Value = ""
	End Sub
	
	
	Public Sub ZAERO_SUB()
		Dim SQLRet As Short
		
		'   エラーフラグの初期化
		ZAER_ERR.Value = "0"
		
		'   << エラーインディケータのセット >>
		'97/07/03Del SQLRet = execsql("set Errorindicator :ERR1:")
		'97/07/03Del SQLRet = execsql("set Errordetail :ERR2:")
		
		'   << テーブル環境設定サブルーチン/テーブル名取得 >>
		MKKCMN.ZAEV_FNO = ZAER_FID
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ZAER_ERR.Value = "1"
		Else
			
			ZAER_FILE.Value = RTrim(MKKCMN.ZAEV_FNM) & ZAER_FID & "ERRM" '99/12/02 DEL YOSHINO 99/12/09 ADD YOSHINO
			
			'99/12/02 ADD START KTT YOSHINO
			'        '引数 "[SECTION名]TABLE名" を分割する
			'Dim wTABLE As String
			'Dim i      As Integer
			'
			'        i = InStr(1, ZAER_FID, "]")
			'        If i > 0 Then
			'            wTABLE = RTrim$(Mid$(ZAER_FID, i + 1))
			'        Else
			'            ZAER_ERR = "1"
			'            Exit Sub
			'        End If
			'
			'        i = InStr(1, MKKCMN.ZAEV_FNM, ".")
			'        ZAER_FILE = Mid$(MKKCMN.ZAEV_FNM, 1, i) & wTABLE & "ERRM"
			'        ZAER_FILE = ZAER_FILE & RTrim$(Mid$(MKKCMN.ZAEV_FNM, i + 1))
			'        ZAER_FILE = ZAER_FILE & Space$(1)
			'        ZAER_FID = wTABLE       '99/12/08 ADD
			''99/12/02 ADD END  KTT YOSHINO
			'------------------------------------------ 99/12/09 DEL END KTT YOSHINO ↑
			
		End If
	End Sub
End Module