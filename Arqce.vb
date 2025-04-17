Option Strict Off
Option Explicit On
Module ARQCEBAS
	
	
	'------------------------------------------------------------
	'【関数名】 エラー変換サブルーチン
	'
	'【機  能】 Rdo+SQLServer/Rdo+Oracleから発生したｴﾗｰのうち、従来の
	'           Glue+Oracle版でﾄﾗｯﾌﾟしていたｴﾗｰだった場合、従来のGlueと同じ
	'           ｴﾗｰｽﾃｰﾀｽを返す。それ以外のものは99を返す。
	'
	'【戻り値】 Integer型
	'             0     :エラー無し
	'            -1     :一意制約違反
	'            24     :End Of Fetch
	'           -54     :ロック中
	'          -955     :既に使用されているオブジェクトのため作成できない
	'          -100     :他で更新済み（ZACN_DB = SQLSRV の時のみ検出）
	'            99     :それ以外のエラー
	'
	'【関数仕様】
	'    Public Function B_STATUS(Optional rKekka As Variant) As Integer
	'        ＜引数＞
	'           rKekka：    OpenResultset,FETCH(MoveNext)後のみその結果ｾｯﾄを指定。
	'                       それ以外の場合は省略すること。
	'
	'           rKekkaが指定されていたら、End Of Fetchかどうかのチェックを
	'           最初に行う。(End Of FetchならB_STATUSは24で返す)
	'           省略されていたら、End Of Fetchのチェックを行わない。
	'
	'【使用例】
	'       １．結果ｾｯﾄを指定しない場合
	'           AM13INS.Execute
	'           Select Case B_STATUS
	'           ase 0       正常
	'           Case -1     一意制約違反
	'           Case Else   それ以外
	'           End Select
	'
	'       ２．結果ｾｯﾄを指定する場合
	'           Set AM13RS = AM13SEL02.OpenResultset()
	'           Select Case B_STATUS(AM13RS)
	'           Case 0      正常
	'           Case -54    ロック中
	'           Case 24    データ無し
	'           Case Else   それ以外
	'           End Select
	'
	'----------------------------------------------------------------------------------
	'【修正履歴】
	'
	'   修正日付：1998/09/11    修正者：Y.Kubo(OSK)
	'   修正内容：Intersolv製OracleODBCﾄﾞﾗｲﾊﾞ対応のため、「ﾛｯｸ中」Status取得方法を変更。
	'
	'-----------------------------------------------------------------------------------
	Public Function B_STATUS(Optional ByRef rKekka As Object = Nothing) As Short
		Dim RdoErr As RDO.rdoError
		
		If Err.Number = 0 Then
			'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
			If Not IsNothing(rKekka) Then
				'UPGRADE_WARNING: オブジェクト rKekka.EOF の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				If rKekka.EOF = True Then
					'End Of Fetch
					B_STATUS = 24
					Exit Function
				End If
			End If
			'エラー無し
			B_STATUS = 0
			Exit Function
		End If
		
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			Select Case RdoErr.SQLState
				Case "01S03"
					If ZACN_DB = SQLSRV Then
						'他で更新済み
						B_STATUS = -100
						Exit Function
					End If
				Case "23000"
					'一意制約違反
					B_STATUS = -1
					Exit Function
				Case "NA000", "S1T00"
					'ロック中
					B_STATUS = -54
					Exit Function
					'ロック中判断を追加（IntersolvODBCﾄﾞﾗｲﾊﾞ3.0対策）
				Case "S1000" '98/09/11追加
					If ZACN_DB = ORCL And RdoErr.Number = 54 Then '98/09/11追加
						'ロック中                                    '98/09/11追加
						B_STATUS = -54 '98/09/11追加
						Exit Function '98/09/11追加
					End If '98/09/11追加
				Case "S0001"
					'既に使用されているオブジェクトのため作成できない
					B_STATUS = -955
					Exit Function
			End Select
		Next RdoErr
		
		B_STATUS = 99
	End Function
End Module