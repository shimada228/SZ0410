Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class SZ0411FRM
	Inherits System.Windows.Forms.Form
	'A-CUST-20100610 フォーム追加
	
	Dim LST_NO As Short '前入力位置ｺﾝﾄﾛｰﾙ№
	Dim NXT_NO As Short '次入力位置ｺﾝﾄﾛｰﾙ№
	Dim CUR_NO As Short '現入力位置ｺﾝﾄﾛｰﾙ№
	Dim MAXNO As Short
	Dim CTRL As System.Windows.Forms.Control
	
	Dim SETSW As Short 'データセット中：ＯＮ
	
	Const N200 As Short = 1 'CSVﾌｧｲﾙ名
	Const N912 As Short = 2 '実 行
	Const NEND As Short = 3
	
	Const GRP1 As Short = 1
	Const GEND As Short = 2
	
	'UPGRADE_WARNING: 配列 CTRLTBL の下限が 1 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Private CTRLTBL(NEND) As CTRLTBL_S '画面ｺﾝﾄﾛｰﾙ配列
	'UPGRADE_WARNING: 配列 GRPTBL の下限が 1 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Private GRPTBL(GEND) As GRPTBL_S '画面ｸﾞﾙｰﾌﾟ配列
	
	Private Structure EXPROT_PATH
		Dim EP_FPATH As String '出力ファイルのパス格納用
		Dim EP_FNAME As String '出力ファイル名格納用
	End Structure
	Private EPF As EXPROT_PATH
	
	Private sPath As String 'カレントディレクトリの初期位置
	Private sDrive As String 'カレントドライブの初期位置
	Private FILECHKFLG As Short 'ファイルチェックをしたらTRUE
	Private DEL_INC_CODE As String
	Private DEL_JG_CODE As String
	
	Private blnCheckPass As Boolean
	
	Private Sub TBL_SET() '画面ｺﾝﾄﾛｰﾙ初期設定
		
		'グループの設定
		CTRLTBL(N200).IGRP = GRP1
		
		CTRLTBL(N912).IGRP = GEND
		CTRLTBL(NEND).IGRP = GEND
		
		'次項目、前項目の設定
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
		
		'--- ファンクション・ガイドメッセージ
		Select Case LST_NO
			Case N200 'CSVﾌｧｲﾙ名
				CMDOFNC(5).Text = "クリア"
				CMDOFNC(5).Enabled = True
				LBLFNC(5).Enabled = True
				CMDOFNC(8).Text = "ファイル"
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
		
		'--- ファンクションメッセージ
		'    Call ZAFC_SUB(Me)
		
		'--- ガイドメッセージ表示
		Call ZAGD_SUB(Me)
	End Sub
	
	Public Sub ENABLED_RTN(ByRef TF As Short)
		'画面の有効、無効、表示内容設定
		
		'条件指定画面
		IMTX200.Enabled = TF '条件設定域
		
		'ｺﾏﾝﾄﾞﾎﾞﾀﾝの制御
		'出力
		If TF = False Then
			CMDOFNC(12).Text = "中  断"
			CMDOFNC(12).MousePointer = System.Windows.Forms.Cursors.Arrow
		Else
			CMDOFNC(12).Text = "実  行"
			CMDOFNC(12).MousePointer = System.Windows.Forms.Cursors.Default
		End If
		
		If TF = False Then
			CMDOFNC(0).Text = ""
			CMDOFNC(8).Text = ""
		Else
			CMDOFNC(0).Text = ZAFC_MST(1)
			CMDOFNC(8).Text = ZAFC_MST(8)
		End If
		
		CMDOFNC(0).Enabled = TF '終了
		LBLFNC(0).Enabled = TF
		CMDOFNC(8).Enabled = TF 'ファイル
		LBLFNC(8).Enabled = TF
		
		'ガイドメッセージ表示
		If TF = False Then 'TRUE → FALSE
			ZAGD_NO.Value = ""
			Call ZAGD_SUB(Me)
		Else
			Call ZAGD_SUB(Me) 'ｶﾞｲﾄﾞﾒｯｾｰｼﾞｸﾘｱ
		End If
		Me.Refresh()
	End Sub
	
	Private Sub ENB_SET_RTN(ByRef TGNO As Short)
		'ﾀﾌﾞｽﾄｯﾌﾟの設定
		
		Select Case TGNO
			Case N200
				IMTX200.TabStop = True
			Case Else
				IMTX200.TabStop = True
		End Select
	End Sub
	
	Private Sub CDL_INIT()
		'CommonDialogの初期設定をします。
		
		'キャンセルをエラーとして扱う
		'UPGRADE_WARNING: The CommonDialog CancelError プロパティは Visual Basic .NET でサポートされていません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"' をクリックしてください。
		CDL010.CancelError = True
		'[ファイルを開く]ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽに設定
		'UPGRADE_WARNING: MSComDlg.CommonDialog プロパティ CDL010.Flags は、新しい動作をもつ CDL010Open.ShowReadOnly にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' をクリックしてください。
		'UPGRADE_WARNING: FileOpenConstants 定数 FileOpenConstants.cdlOFNHideReadOnly は、新しい動作をもつ OpenFileDialog.ShowReadOnly にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' をクリックしてください。
		CDL010Open.ShowReadOnly = False
		'リスト ボックスに表示されるフィルタを設定
		'UPGRADE_WARNING: Filter に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		CDL010Open.Filter = "CSV(TAB区切り)|*.CSV|全てのﾌｧｨﾙ|*.*|"
		'CSV を既定のフィルタとして指定
		CDL010Open.FilterIndex = 1
		'デフォルトの拡張子を設定
		CDL010Open.DefaultExt = ".txt"
		'デフォルトのディレクトリを設定(iniから取得)
		CDL010Open.InitialDirectory = WG_EXCELPATH
	End Sub
	
	Private Sub INITIAL_RTN()
		'画面項目初期値設定
		WKBCSVFILE = WG_EXCELPATH
		
		DSP010.Text = RTrim(WKB010)
		DSP020.Text = RTrim(SZ0410FRM.DSP010.Text)
		DSP030.Text = RTrim(WKB020)
		DSP040.Text = RTrim(SZ0410FRM.DSP020.Text)
		
		'表示
		IMTX200.Text = RTrim(WKBCSVFILE)
		
		EPF.EP_FNAME = ""
		EPF.EP_FPATH = WG_EXCELPATH
		
		Call CDL_INIT()
	End Sub
	
	Private Sub FOCUS_SET() 'ﾌｫｰｶｽｾｯﾄ
		Select Case NXT_NO
			Case N200 'CSVﾌｧｲﾙ名
				IMTX200.Focus()
			Case N912 '実行
				CMDOFNC(12).Focus()
		End Select
	End Sub
	
	Private Sub SET_NO(ByRef FUNC As Short) '画面ｺﾝﾄﾛｰﾙ№セット
		Dim i As Short
		
		i = LST_NO
		
		Do 
			Select Case FUNC
				Case 1 '次項目
					NXT_NO = CTRLTBL(i).INEXT
				Case 2 '前項目
					NXT_NO = CTRLTBL(i).IBACK
				Case 3 '次グループ
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
		'全チェック & 実行
		
		CUR_NO = NEND
		
		'直前項目のチェック
		If IPROCHK() = False Then
			Exit Sub
		End If
		
		'全グループのチェック
		If GPROCHK() = False Then
			Exit Sub
		End If
		
		'If MsgBox("ＣＳＶの取込みを行います。よろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2, Me.Caption) = vbNo Then    'D-CUST-20100901
		If MsgBox("ＣＳＶの取込みを行います。よろしいですか？", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then 'A-CUST-20100901
			NXT_NO = N200
			Call FOCUS_SET()
			Exit Sub
		End If
		
		'処理ログの出力（ﾌﾟﾚﾋﾞｭｰ、印刷を実行した時点で書き込む）
		'サーバーの日付・時刻を取得
		Dim strSvrDate As String
		
		SYSDATE = CduServerDate
		strSvrDate = VB6.Format(SYSDATE, "YYYYMMDDHHNNSS")
		
		'ログ出力サブルーチン呼出
		ZALGM_INC_CODE.Value = WKB010 '会社
		ZALGM_JG_CODE.Value = WKB020 '事業所
		ZALGM_SYS_KBN.Value = "3"
		ZALGM_S_DAY.Value = VB.Left(strSvrDate, 8) '日付
		ZALGM_S_TIME.Value = VB.Right(strSvrDate, 6) '時刻
		ZALGM_OP_CODE.Value = WG_OPCODE 'オペレータコード
		ZALGM_PGID.Value = "SZ0410" 'システム
		ZALGM_SH_KBN.Value = "3"
		ZALGM_SH_NAIYO.Value = WKB010 & "-" & WKB020
		ZALGM_GNFLG.Value = "0"
		Call ZALGM_SUB(ZACNA_RCN)
		
		'マウスカーソルを砂時計に設定
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		Call ENABLED_RTN(False)
		
		PRNSW = F_ON
		
		Call BEGIN_RTN()
		
		FSTSW = F_FST
		ENDSW = F_OFF
		ERRSW = F_OFF
		'*** 取込処理 ***
		Call TORIKOMI_RTN()
		If CSV_CNT = 0 And ERRSW <> F_ERR Then
			'CSVファイル中に対象となるデータがなかった
			ZAER_CD = 129 '対象データなし
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			ERRSW = F_ERR
		End If
		
		If ERRSW <> F_ERR Then
			'データ取込成功！
			Call COMMIT_RTN()
			
			MsgBox("正常終了しました。", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
		Else
			'データ取込失敗！
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
	
	Private Function IPROCHK() As Short 'LostFocus項目チェック
		IPROCHK = True
		ERRSW = F_OFF
		
		If CUR_NO = LST_NO Then Exit Function
		Select Case LST_NO
			Case N200 'CSVﾌｧｲﾙ名
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
		'ファイル名に使えない文字を使用しているかどうかのチェック
		Dim i As Object
		Dim l As Short
		Dim wname As String
		Dim dnum As Short
		
		CHK_FNAME = False
		
		wname = Trim(fname)
		
		l = Len(wname)
		dnum = 0
		
		'ファイル名に使えない文字が入力されてる？
		For i = 3 To l '3文字目から見る
			'UPGRADE_WARNING: オブジェクト i の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
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
		Dim PCRET As Short 'パスチェックサブルーチンの戻り値
		Dim FULLNAME As String '左右のスペースをカットしたファイル名格納エリア
		Dim i As Object
		Dim j As Integer
		Dim wStr As String
		
		If CUR_NO < LST_NO Then
			Exit Sub
		End If
		
		'未入力のチェック
		If VB.Right(Trim(IMTX200.Text), 1) = "\" Or Trim(IMTX200.Text) = "" Then
			ERRSW = F_ERR
			Exit Sub
		End If
		
#If 0 Then
		'UPGRADE_NOTE: 式 0 が True に評価されなかったか、またはまったく評価されなかったため、#If #EndIf ブロックはアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"' をクリックしてください。
		'パスが指定されていなければ、デフォルトのパスを指定させる
		If InStr(1, Trim$(IMTX200.Text), "\", vbTextCompare) = 0 Then
		wStr = ""
		'ファイル名とパスの間に"\"が必要かどうかの判定
		If Right$(Trim$(WG_EXCELPATH), 1) <> "\" And Left$(Trim$(IMTX200.Text), 1) <> "\" Then
		wStr = "\"
		End If
		
		IMTX200.Text = WG_EXCELPATH & wStr & Trim$(IMTX200.Text)
		End If
#End If
		
		'   '拡張子のチェック
		'    Select Case StrConv(Right$(Trim$(IMTX200.Text), 4), vbUpperCase)
		'        Case ".CSV"
		'        Case Else
		'            IMTX200.Text = Trim$(IMTX200.Text) & ".CSV"
		'    End Select
		
		'ファイル名に使えない文字が入力されてるかどうかのチェック
		If CHK_FNAME((IMTX200.Text)) = False Then
			ZAER_CD = 11 'ﾌｧｲﾙ名不正
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			ERRSW = F_ERR
			Exit Sub
		End If
		
		'ファイル名チェックがまだだったときだけ、チェックを行う
		'(何度も上書き確認を出さないため)
		If FILECHKFLG = False Then
			'ファイル名チェック
			FULLNAME = Trim(IMTX200.Text)
			PCRET = MKKCMN.ZAPC_SUB(FULLNAME)
			Select Case PCRET
				Case 0 'ﾌｧｲﾙ無し
					ZAER_CD = 12
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
				Case -1 'ファイルがあった：OK！
				Case 11 'ファイル名不正
					ZAER_CD = PCRET
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					ERRSW = F_ERR
					Exit Sub
				Case 190 'ドライブの準備ができていない
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
			
			'チェックが正常だったら、チェック済みフラグをたてる
			FILECHKFLG = True
			WKBCSVFILE = Trim(IMTX200.Text)
		End If
	End Sub
	
	Private Function GPROCHK() As Short 'LostFocus項目群チェック
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
	
	Private Function GVALCHK() As Short '項目群入力可否チェック
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
	
	Private Function MVALCHK() As Short '項目入力可否チェック
		MVALCHK = True
		ERRSW = F_OFF
		
		Select Case CUR_NO
			Case N200 'CSVﾌｧｲﾙ名
				Call MVALCHK_N200()
		End Select
		
		If ERRSW = F_ERR Then
			MVALCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
	End Function
	
	Sub MVALCHK_N200() 'CSVﾌｧｲﾙ名
		
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
			Case 0 '終了
				Call HIDE_RTN()
			Case 5 'クリア
				WKBCSVFILE = RTrim(WG_EXCELPATH)
				IMTX200.Text = WKBCSVFILE
				NXT_NO = N200
				Call FOCUS_SET()
				
			Case 8 'ファイル
				Call GETDIR_RTN() 'ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽより、ファイル名及びパスを取得する
				ChDir((sPath)) 'カレントディレクトリを戻す
				ChDrive((sDrive)) 'カレントドライブを戻す
				NXT_NO = CUR_NO
				Call FOCUS_SET()
				
			Case 12 '実行
				'ガイドメッセージ表示
				If PRNSW = F_OFF Then
					'マウスカーソルを砂時計に設定
					Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
					'全チェック & 実行
					Call ALLCHK_RTN()
					'マウスカーソルを通常モードに設定
					Me.Cursor = System.Windows.Forms.Cursors.Default
				Else '印刷中断
					If MsgBox("中断しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.Yes Then
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
		'ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽ表示！
		CDL010Open.ShowDialog()
		
		' ユーザーが [開く] をクリックした。
		
		'ファイル名、パスを取り出す。
		'UPGRADE_WARNING: CommonDialog プロパティ CDL010.FileTitle は、新しい動作をもつ CDL010.FileName にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' をクリックしてください。
		EPF.EP_FNAME = CDL010Open.FileName 'ファイル名
		StrLen = InStr(CDL010Open.FileName, EPF.EP_FNAME)
		EPF.EP_FPATH = Mid(CDL010Open.FileName, 1, StrLen - 1) 'パス
		
		'確定
		WKBCSVFILE = CDL010Open.FileName 'ファイル名(パスを含む)
		IMTX200.Text = WKBCSVFILE
		
		'ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽ再設定
		CDL010Open.InitialDirectory = EPF.EP_FPATH
		'UPGRADE_WARNING: CommonDialog プロパティ CDL010.FileTitle は、新しい動作をもつ CDL010.FileName にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"' をクリックしてください。
		CDL010Open.FileName = CDL010Open.FileName
		
ErrHandler: 
		' ユーザーが [キャンセル] をクリックした。
		Exit Sub
	End Sub
	
	Private Sub CMDOFNC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Enter
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		'ﾌﾟﾚﾋﾞｭｰ、実行ボタン以外は無視
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
		' ファンクションガイド
		Call FUNCSET_RTN()
	End Sub
	
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		'ﾌﾟﾚﾋﾞｭｰ、印刷以外は無視
		If Index < 12 Then
			Exit Sub
		End If
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				Select Case Index
					Case 12 '実行
						NXT_NO = N200 'CSVﾌｧｲﾙ名へ
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
	
	'UPGRADE_ISSUE: PictureBox イベント DUMMYDEL.GotFocus はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' をクリックしてください。
	Private Sub DUMMYDEL_GotFocus()
		'過去データ　削除
		'処理ログの出力（ﾌﾟﾚﾋﾞｭｰ、印刷を実行した時点で書き込む）
		'サーバーの日付・時刻を取得
		Dim strSvrDate As String
		
		SYSDATE = CduServerDate
		strSvrDate = VB6.Format(SYSDATE, "YYYYMMDDHHNNSS")
		
		'ログ出力サブルーチン呼出
		ZALGM_INC_CODE.Value = WKB010 '会社
		ZALGM_JG_CODE.Value = WKB020 '事業所
		ZALGM_SYS_KBN.Value = "3"
		ZALGM_S_DAY.Value = VB.Left(strSvrDate, 8) '日付
		ZALGM_S_TIME.Value = VB.Right(strSvrDate, 6) '時刻
		ZALGM_OP_CODE.Value = WG_OPCODE 'オペレータコード
		ZALGM_PGID.Value = "SZ0410" 'システム
		ZALGM_SH_KBN.Value = "5"
		ZALGM_SH_NAIYO.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Month, -1, SYSDATE), "YYYYMMDD")
		ZALGM_GNFLG.Value = "0"
		Call ZALGM_SUB(ZACNA_RCN)
		
		'マウスカーソルを砂時計に設定
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Call BEGIN_RTN()
		
		ERRSW = F_OFF
		'*** 別品目マスタ取込処理 ***
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
	
	'UPGRADE_WARNING: Form イベント SZ0411FRM.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub SZ0411FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'マウスカーソルを戻す
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		Me.Text = "品目情報入力ＣＳＶ取込"
		
		'ウインドウ表示位置設定サブルーチン
		Call ZAWC_SUB(Me, 0)
		Me.Top = 0
		Me.Left = 0
		
		'オペレータ名表示サブルーチン
		Call ZAOP_SUB(Me, WKB010, WG_OPCODE)
		If ERRSW = F_ERR Then
			Call ENDR_RTN()
		End If
		
		Call INITIAL_RTN() '初期画面表示
		
		'起動権限チェック
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
		
		''更新権限チェック
		''更新権限
		'lRet = MKKDBCMN.MKKDBCMN_SQTGET3_SUB(4, "SZ0410", WKB010, WKB020, "", WG_OPCODE, OP_KENGEN)
		'If lRet <> n0 Then
		'    Call HIDE_RTN
		'    Exit Sub
		'End If
		''更新権限なし
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
				Call SET_NO(1) ' 次項目
				KeyCode = n0
			Case System.Windows.Forms.Keys.Up
				Call SET_NO(2) ' 前項目
				KeyCode = n0
			Case System.Windows.Forms.Keys.Down
				Call SET_NO(3) ' 次グループ
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
		
		'UPGRADE_ISSUE: Form プロパティ SZ0411FRM.HelpContextID はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
		Me.HelpContextID = SM_HelpContextID
		
		Call TBL_SET() '画面ｺﾝﾄﾛｰﾙ初期設定
		LST_NO = NEND
		
		sPath = CurDir() 'カレントディレクトリの初期位置取得
		sDrive = VB.Left(CurDir(), 1) 'カレントドライブの初期位置取得
	End Sub
	
	Private Sub SZ0411FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		'終了 or 中断
		'UPGRADE_ISSUE: 定数 vbFormCode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		If UnloadMode <> vbFormCode Then
			If SETSW = F_ON Or PRNSW = F_ON Then
				If MsgBox("中断しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question, Me.Text) = MsgBoxResult.Yes Then
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
		' 値が変更されたので、チェック済みフラグを未チェックにする
		FILECHKFLG = False
	End Sub
	
	Private Sub IMTX200_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX200.Enter
		'CSVﾌｧｲﾙ名
		
		CUR_NO = N200
		
		'チェック
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
		
		'確定
		LST_NO = CUR_NO
		
		'ファンクションガイド
		Call FUNCSET_RTN()
	End Sub
	
	Private Sub IMTX200_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX200.KeyDownEvent
		'CSVﾌｧｲﾙ名
		Call SZ0411FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
	End Sub
End Class