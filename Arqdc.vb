Option Strict Off
Option Explicit On
Module ARQDCBAS
	'**************************************************************
	'*      日付チェックサブルーチン　 　                         *
	'*          　　　　　　　　　（ＡＶＱＤＣ）                  *
	'*                                                            *
	'*      エラー対象基準年　１９００年                          *
	'**************************************************************
	'エラー対象基準年
	Const ZADC_KIJUN As Short = 1900
	
	'引渡設定パラメータ
	Public ZADC_DATE As New VB6.FixedLengthString(8) '西暦日付
	
	'結果引渡パラメータ
	Public ZADC_STS As New VB6.FixedLengthString(1) '結果ｽﾃｰﾀｽ　0:正常 1:エラー
	Public ZADC_WEEK As New VB6.FixedLengthString(1) '曜日区分   1:日曜 2:月曜
	'           3:火曜 4:水曜
	'           5:木曜 6:金曜
	'           7:土曜 0:ｴﾗｰ
	'ワーク
	Dim ZADCL_YMD As Object
	
	Sub ZADC_SUB()
		
		'初期状態をエラーに設定
		ZADC_STS.Value = "1"
		ZADC_WEEK.Value = "0"
		
		'数値チェック
		If IsNumeric(ZADC_DATE.Value) Then
			
			'８桁入力チェック
			If Len(Trim(ZADC_DATE.Value)) = 8 Then
				
				'年妥当性チェック
				If CDbl(Mid(ZADC_DATE.Value, 1, 4)) >= ZADC_KIJUN Then
					
					'月日妥当性チェック
					If Mid(ZADC_DATE.Value, 5, 2) >= "01" And Mid(ZADC_DATE.Value, 5, 2) <= "12" Then
						
						'日付妥当性チェック
						'UPGRADE_WARNING: オブジェクト ZADCL_YMD の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
						ZADCL_YMD = VB6.Format(Val(ZADC_DATE.Value), "0000/00/00")
						If IsDate(ZADCL_YMD) Then
							
							'曜日を求める
							'UPGRADE_WARNING: オブジェクト ZADCL_YMD の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							ZADC_WEEK.Value = VB6.Format(ZADCL_YMD, "W")
							ZADC_STS.Value = "0"
						End If
					End If
				End If
			End If
		End If
		ZADC_DATE.Value = ""
	End Sub
End Module