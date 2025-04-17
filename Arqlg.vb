Option Strict Off
Option Explicit On
Module ARQLGBAS
	
	Public ZALG_KBN As New VB6.FixedLengthString(1)
	Public ZALG_NAIYO As String
	
	
	Public ZALG_ERR As New VB6.FixedLengthString(1)
	
	
	Public Sub ZALG_SUB()
		Dim Ret As Short
		
		Dim LOGDIRNAME As String 'ログ出力ディレクトリ名
		Dim LOGFNAME As String 'パスを含むログファイル名
		'
		Dim SYSDATE As New VB6.FixedLengthString(8) 'システム日付
		Dim SYSTIME As New VB6.FixedLengthString(6) 'システム時間
		Dim PRGID As String 'プログラムＩＤ
		
		Dim OUTFNum As Short '出力ファイルのファイル番号
		Dim OUT_REC As String '出力ファイルレイアウト
		
		'   << エラーフラグのクリア >>
		ZALG_ERR.Value = "0"
		
		'   << システム日付、時間、プログラムＩＤ取り込み >>
		SYSDATE.Value = VB6.Format(Now, "YYYYMMDD")
		SYSTIME.Value = VB6.Format(Now, "HHMMSS")
		'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		PRGID = My.Application.Info.AssemblyName
		
		'   << Smile.iniより、ログ出力ディレクトリの取得 >>
		Ret = MKKCMN.ZAGI_SUB("LOG", "LOGFNAME", "", LOGDIRNAME, "SMILE.INI")
		If Ret = False Then
			GoTo ZALG_END
		End If
		
		'   << 出力ファイル名生成（パス含み） >>
		If Len(LOGDIRNAME) <> 0 Then
			If Mid(LOGDIRNAME, Len(LOGDIRNAME) - 1, 1) = "\" Then
				LOGFNAME = LOGDIRNAME & "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
			Else
				LOGFNAME = LOGDIRNAME & "\" & "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
			End If
		Else
			LOGFNAME = "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
		End If
		
		'   << 出力ファイルのファイル名チェック >>
		Ret = MKKCMN.ZAPC_SUB(LOGFNAME)
		If Ret <> 0 And Ret <> -1 Then
			GoTo ZALG_END
		End If
		
		'****** << ログファイル出力処理 >> *****
		
		'   <<  オープン >>
ZALG_0010: 
		OUTFNum = FreeFile
		On Error Resume Next
		Err.Clear()
		FileOpen(OUTFNum, LOGFNAME, OpenMode.Append, , OpenShare.LockReadWrite)
		Select Case Err.Number
			Case 0 ' 正常.
				
			Case 53, 75, 76 ' パスが不正、または見つからなかった
				GoTo ZALG_END
			Case 52, 64 ' ファイル名が無効
				GoTo ZALG_END
			Case 70 ' 書き込み不能→ファイル使用中
				GoTo ZALG_0010
			Case 68, 71 ' デバイス/ドライブの準備ができていない
				GoTo ZALG_END
			Case Else ' アクセス不可能
				GoTo ZALG_END
		End Select
		On Error GoTo 0
		
		'   << 書き込み処理 >>
		OUT_REC = SYSDATE.Value & "," & SYSTIME.Value & "," & PRGID & ","
		'
		Select Case ZALG_KBN.Value
			Case "0"
				OUT_REC = OUT_REC & "開始," & ZALG_NAIYO
			Case "1"
				OUT_REC = OUT_REC & "終了," & ZALG_NAIYO
			Case "9"
				OUT_REC = OUT_REC & "エラー," & ZALG_NAIYO
		End Select
		
		On Error Resume Next
		PrintLine(OUTFNum, OUT_REC)
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("ログファイルの書き込みに失敗しました。" & Err.Number, 48, "")
				ZALG_ERR.Value = "1"
		End Select
		On Error GoTo 0
		
		'   <<  ＣＬＯＳＥ >>
		On Error Resume Next
		FileClose(OUTFNum)
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("ログファイルの書き込みに失敗しました。" & Err.Number, 48, "")
				ZALG_ERR.Value = "1"
		End Select
		On Error GoTo 0
		
		
ZALG_END: 
		
	End Sub
End Module