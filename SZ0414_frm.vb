Option Strict Off
Option Explicit On
Friend Class SZ0414FRM
	Inherits System.Windows.Forms.Form
	
	
	Dim LST_NO As Short '前入力位置ｺﾝﾄﾛｰﾙ№
	Dim NXT_NO As Short '次入力位置ｺﾝﾄﾛｰﾙ№
	Dim CUR_NO As Short '現入力位置ｺﾝﾄﾛｰﾙ№
	Dim MAXNO As Short
	Dim CTRL As System.Windows.Forms.Control
	
	Const N010 As Short = 1 '
	Const N020 As Short = 2 '
	Const N030 As Short = 3 '
	Const N040 As Short = 4 '
	Const N050 As Short = 5 '
	Const N060 As Short = 6 '
	Const N070 As Short = 7 '
	Const N080 As Short = 8 '
	Const N090 As Short = 9 '
	Const N100 As Short = 10 '
	Const NEND As Short = 11
	
	Const GRP1 As Short = 1
	Const GEND As Short = 2
	
	'UPGRADE_WARNING: 配列 CTRLTBL の下限が 1 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Dim CTRLTBL(NEND) As CTRLTBL_S '画面ｺﾝﾄﾛｰﾙ配列
	'UPGRADE_WARNING: 配列 GRPTBL の下限が 1 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Dim GRPTBL(GEND) As GRPTBL_S '画面ｸﾞﾙｰﾌﾟ配列
	
	Dim RDO_STATUS As Short
	
	
	Dim SZ_IDOSW As Short ' 選択中 LOSTで1をｾｯﾄしCLICKでSW=0なら取込 1なら0に
	Dim SZ_UPSW As Short ' １行目で↑を押しているときSW=1
	Dim SZ_DOWNSW As Short ' 10行目で↓を押しているときSW=1
	Dim SZ_IDX As Short ' INDEX 明細移動用INDEX
	
	Dim SZ_LNCNT As Decimal ' 行番号
	Dim SZ_REP As Decimal ' カウンター
	Dim SZ_I As Short ' 明細表示用カウンター
	Dim SZ_IMAX As Short ' 表示行数(1～10)
	Dim SZ_ENDSW As Short ' AT END 時にSW=1
	Const SZ0414_MAX_ROW As Short = 10 ' １画面表示行数
	
	
	'UPGRADE_ISSUE: 宣言の型がサポートされていません: 固定長文字列の配列 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"' をクリックしてください。
	Dim SJAN_K4(SZ0414_MAX_ROW) As String*13 ' 業者コード
	
	Dim MOUSEFLG As Short
	
	Dim HOUKOU As Short ' 読み込み方法（0:順方向 1:逆方向）
	Dim FETCH_MODE As String ' SELECTの状態
	
	Dim SZ0414SEL_SW As Short ' 索引順SELECT文の判断
	
	Dim B_OP As Short
	Const B_GET_EQ As Short = 5
	Const B_GET_NEXT As Short = 6
	Const B_GET_PRE As Short = 7
	Const B_GET_GT As Short = 8
	Const B_GET_GE As Short = 9
	Const B_GET_LT As Short = 10
	Const B_GET_LE As Short = 11
	
	Private Sub ENDRR_RTN(ByRef MyForm As System.Windows.Forms.Form)
		'
		' ｺｰﾄﾞ問合せ ﾌｫｰﾑ終了時処理
		'
		'*** ｳｲﾝﾄﾞｳ表示位置ｾｰﾌﾞ
		'    Dim Ret As Long
		'    Ret = GetWindowRect(MyForm.hwnd, lpRectSave)
		
		'*** ﾌｫｰﾑ ｱﾝﾛｰﾄﾞ
		MyForm.Close()
		
	End Sub
	
	Private Sub SPRD_ROWHT_RTN()
		
		Dim IDX As Short
		
		For IDX = 1 To SPRD.MaxRows
			SPRD.set_RowHeight(IDX, 19.6)
		Next IDX
		
	End Sub
	
	Private Sub TBL_SET() '画面ｺﾝﾄﾛｰﾙ初期設定
		
		'グループの設定
		CTRLTBL(N010).IGRP = GRP1
		CTRLTBL(N020).IGRP = GRP1
		CTRLTBL(N030).IGRP = GRP1
		CTRLTBL(N040).IGRP = GRP1
		CTRLTBL(N050).IGRP = GRP1
		CTRLTBL(N060).IGRP = GRP1
		CTRLTBL(N070).IGRP = GRP1
		CTRLTBL(N080).IGRP = GRP1
		CTRLTBL(N090).IGRP = GRP1
		CTRLTBL(N100).IGRP = GRP1
		CTRLTBL(NEND).IGRP = GEND
		
		'次項目、前項目の設定
		CTRLTBL(N010).INEXT = N020
		CTRLTBL(N010).IBACK = n0
		CTRLTBL(N010).IDOWN = N020
		
		CTRLTBL(N020).INEXT = N030
		CTRLTBL(N020).IBACK = N010
		CTRLTBL(N020).IDOWN = N030
		
		CTRLTBL(N030).INEXT = N040
		CTRLTBL(N030).IBACK = N010
		CTRLTBL(N030).IDOWN = N040
		
		CTRLTBL(N040).INEXT = N050
		CTRLTBL(N040).IBACK = N030
		CTRLTBL(N040).IDOWN = N050
		
		CTRLTBL(N050).INEXT = N060
		CTRLTBL(N050).IBACK = N030
		CTRLTBL(N050).IDOWN = N060
		
		CTRLTBL(N060).INEXT = N070
		CTRLTBL(N060).IBACK = N050
		CTRLTBL(N060).IDOWN = N070
		
		CTRLTBL(N070).INEXT = N080
		CTRLTBL(N070).IBACK = N060
		CTRLTBL(N070).IDOWN = N080
		
		CTRLTBL(N080).INEXT = N090
		CTRLTBL(N080).IBACK = N060
		CTRLTBL(N080).IDOWN = N090
		
		CTRLTBL(N090).INEXT = N100
		CTRLTBL(N090).IBACK = N080
		CTRLTBL(N090).IDOWN = N100
		
		CTRLTBL(N100).INEXT = NEND
		CTRLTBL(N100).IBACK = N090
		CTRLTBL(N100).IDOWN = NEND
		
		CTRLTBL(N010).CTRL = IMTX010
		CTRLTBL(N020).CTRL = IMTX020
		CTRLTBL(N030).CTRL = IMTX030
		CTRLTBL(N040).CTRL = IMTX040
		CTRLTBL(N050).CTRL = IMTX050
		CTRLTBL(N060).CTRL = IMNU060
		CTRLTBL(N070).CTRL = IMNU070
		CTRLTBL(N080).CTRL = CMB080
		CTRLTBL(N090).CTRL = IMNU090
		CTRLTBL(N100).CTRL = IMNU100
		
		MAXNO = NEND
		NXT_NO = N010
		
	End Sub
	
	Private Sub INITIAL_RTN()
		
		
		'画面項目初期値設定
		WKBSZ0414.S010 = "" 'ＪＡＮコード　開始
		WKBSZ0414.S020 = "" 'ＪＡＮコード　終了
		WKBSZ0414.S030 = "" 'ＪＡＮ商品部類　開始
		WKBSZ0414.S030N = "" 'ＪＡＮ商品部類名　開始
		WKBSZ0414.S040 = "" 'ＪＡＮ商品部類　終了
		WKBSZ0414.S040N = "" 'ＪＡＮ商品部類名　終了
		WKBSZ0414.S050 = "" '原産国
		WKBSZ0414.C060 = 0 '重量　開始
		WKBSZ0414.C070 = 0 '重量　終了
		WKBSZ0414.S080 = "" '賞味期限　区分
		WKBSZ0414.C090 = 0 '賞味期限　開始
		WKBSZ0414.C100 = 0 '賞味期限　終了
		
		
		'表示
		IMTX010.Text = RTrim(WKBSZ0414.S010)
		IMTX020.Text = RTrim(WKBSZ0414.S020)
		IMTX030.Text = RTrim(WKBSZ0414.S030)
		DSP030.Text = RTrim(WKBSZ0414.S030N)
		IMTX040.Text = RTrim(WKBSZ0414.S040)
		DSP040.Text = RTrim(WKBSZ0414.S040N)
		IMTX050.Text = RTrim(WKBSZ0414.S050)
		IMNU060.Value = WKBSZ0414.C060
		IMNU070.Value = WKBSZ0414.C070
		CMB080.SelectedIndex = -1
		IMNU090.Value = WKBSZ0414.C090
		IMNU100.Value = WKBSZ0414.C100
		
		CMB080.Items.Clear()
		CMB080.Items.Add("")
		CMB080.Items.Add(New VB6.ListBoxItem("日", 1))
		CMB080.Items.Add(New VB6.ListBoxItem("月", 2))
		CMB080.Items.Add(New VB6.ListBoxItem("年", 3))
		
		SPRD.MaxRows = 0
		JAN.k4 = ""
		
	End Sub
	
	'
	'表示ボタン
	'
	Private Sub CMDO010_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO010.ClickEvent
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		
		Call ALLCHK_RTN()
		If SZ_ERRSW = F_ERR Then Exit Sub
		
		If Trim(WKBSZ0414.S010) <> "" And Trim(WKBSZ0414.S020) <> "" Then
			If Trim(WKBSZ0414.S010) > Trim(WKBSZ0414.S020) Then
				ZAER_CD = 120
				ZAER_NO.Value = ""
				ZAER_MS.Value = ""
				Call ZAER_SUB()
				NXT_NO = N010
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		If Trim(WKBSZ0414.S030) <> "" And Trim(WKBSZ0414.S040) <> "" Then
			If Trim(WKBSZ0414.S030) > Trim(WKBSZ0414.S040) Then
				ZAER_CD = 120
				ZAER_NO.Value = ""
				ZAER_MS.Value = ""
				Call ZAER_SUB()
				NXT_NO = N030
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		If WKBSZ0414.C060 <> 0 And WKBSZ0414.C070 <> 0 Then
			If WKBSZ0414.C060 > WKBSZ0414.C070 Then
				ZAER_CD = 120
				ZAER_NO.Value = ""
				ZAER_MS.Value = ""
				Call ZAER_SUB()
				NXT_NO = N060
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		If (WKBSZ0414.C090 <> 0 Or WKBSZ0414.C100 <> 0) And WKBSZ0414.S080 = "0" Then
			ZAER_CD = 120
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			NXT_NO = N080
			Call FOCUS_SET()
			Exit Sub
		End If
		
		If WKBSZ0414.C090 <> 0 And WKBSZ0414.C100 <> 0 Then
			If WKBSZ0414.C090 > WKBSZ0414.C100 Then
				ZAER_CD = 120
				ZAER_NO.Value = ""
				ZAER_MS.Value = ""
				Call ZAER_SUB()
				NXT_NO = N090
				Call FOCUS_SET()
				Exit Sub
			End If
		End If
		
		
		
		If SZ0414_DSPSW = F_ON Then
			SZ0414SELGE.Close()
			SZ0414SELGT.Close()
			SZ0414SELLT.Close()
		End If
		
		JAN_BUF0.k4 = Space(13) 'JAN
		
		SZ0414_IMTX010.Value = IMTX010.Text ' 前回ＪＡＮコード　開始
		SZ0414_IMTX020.Value = IMTX020.Text ' 前回ＪＡＮコード　終了
		SZ0414_IMTX030.Value = IMTX030.Text ' 前回ＪＡＮ商品部類　開始
		SZ0414_IMTX040.Value = IMTX040.Text ' 前回ＪＡＮ商品部類　終了
		SZ0414_IMTX050.Value = IMTX050.Text ' 前回原産国
		SZ0414_IMNU060 = IMNU060.Value ' 前回重量　開始
		SZ0414_IMNU070 = IMNU070.Value ' 前回重量　終了
		If CMB080.SelectedIndex < 0 Then
			SZ0414_IMTX080.Value = CStr(0)
		Else
			SZ0414_IMTX080.Value = CStr(VB6.GetItemData(CMB080, CMB080.SelectedIndex)) ' 前回賞味期限　区分
		End If
		SZ0414_IMNU090 = IMNU090.Value ' 前回賞味期限　開始
		SZ0414_IMNU100 = IMNU100.Value ' 前回賞味期限　終了
		If CMB080.SelectedIndex = -1 Or CMB080.SelectedIndex = 0 Then
			SZ0414_IMNU090D = 0
			SZ0414_IMNU100D = 0
		Else
			'日換算
			Select Case VB6.GetItemData(CMB080, CMB080.SelectedIndex)
				Case 1 '日の場合
					SZ0414_IMNU090D = SZ0414_IMNU090
					SZ0414_IMNU100D = SZ0414_IMNU100
				Case 2 '月の場合
					''SZ0414_IMNU090D = Fix(SZ0414_IMNU090 * 30.5)          'D-20130227
					''SZ0414_IMNU100D = Fix(SZ0414_IMNU100 * 30.5)          'D-20130227
					SZ0414_IMNU090D = Fix(SZ0414_IMNU090 * 30.416 + 0.5) 'A-20130227
					SZ0414_IMNU100D = Fix(SZ0414_IMNU100 * 30.416 + 0.5) 'A-20130227
				Case 3 '年の場合
					SZ0414_IMNU090D = SZ0414_IMNU090 * 365
					SZ0414_IMNU100D = SZ0414_IMNU100 * 365
			End Select
		End If
		
		'ＪＡＮマスタ
		Call PREP_JAN()
		If SZ_ERRSW = F_ERR Then Call ENDRR_RTN(Me)
		
		SPRD.ReDraw = False
		SZ_LNCNT = 0
		Call KENSAKU_RTN()
		If (SZ_IMAX <> 0) Then
			SPRD.ROW = 1
			SPRD.Focus()
		End If
		SPRD.ReDraw = True
		
	End Sub
	
	Private Sub KENSAKU_RTN()
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		Me.Refresh()
		
		Call SZ0414_STA_SUB(0)
		If SZ_ERRSW = 1 Then
			CMDOFNC(0).Focus()
			Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
			Exit Sub
		End If
		
		If SZ_IMAX = n0 Then
			SPRD.MaxRows = 0
			SPRD.Enabled = False
			CMDO020.Enabled = False
			CMDO030.Enabled = False
		Else
			SPRD.Enabled = True
			CMDO020.Enabled = True
			CMDO030.Enabled = True
		End If
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Refresh()
	End Sub
	
	Private Sub CMDO010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO010.Enter
		'    If CUR_NO = NEND Then Exit Sub
		
		CUR_NO = NEND
		
		'チェック
		If LST_NO <> n0 Then
			If IPROCHK() = False Then
				Exit Sub
			End If
			If GPROCHK() = False Then
				Exit Sub
			End If
		End If
		
		If GVALCHK() = False Then
			Exit Sub
		End If
		
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		'確定
		LST_NO = NEND
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub CMDO010_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDO010.KeyDownEvent
		If eventArgs.Shift <> n0 Then Exit Sub
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Return
				Call CMDO010_ClickEvent(CMDO010, New System.EventArgs())
			Case System.Windows.Forms.Keys.Up
				IMNU100.Focus()
		End Select
		
	End Sub
	
	Private Sub CMDO010_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDO010.MouseDownEvent
		MOUSEFLG = eventArgs.Button
		
	End Sub
	
	Private Sub CMDO020_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO020.ClickEvent
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		
		SZ_UPSW = 0
		
	End Sub
	
	Private Sub CMDO020_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO020.Enter
		SPRD.ROW = SZ_IDX
		SPRD.Focus()
		
	End Sub
	
	Private Sub CMDO020_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDO020.MouseDownEvent
		MOUSEFLG = eventArgs.Button
		
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			Exit Sub
		End If
		
		SZ_UPSW = 1
		
		Do 
			If SZ_IDX = 1 Then
				Do 
					If SZ_IDX = 1 Then
						Call SZ0414_PRE_SUB()
					End If
					If SZ_ERRSW = 1 Then
						Exit Sub
					End If
					SPRD.ROW = SZ_IDX
					SPRD.Col = 1
					SPRD.Action = SS_ACTION_ACTIVE_CELL
					SPRD.Focus()
					System.Windows.Forms.Application.DoEvents()
				Loop Until (SZ_UPSW = 0) Or (SZ_ENDSW = 1)
			Else
				Do 
					If (SZ_IDX > 1) Then
						SZ_IDX = SZ_IDX - 1
					End If
					SPRD.ROW = SZ_IDX
					SPRD.Col = 1
					SPRD.Action = SS_ACTION_ACTIVE_CELL
					SPRD.Focus()
					System.Windows.Forms.Application.DoEvents()
				Loop Until (SZ_UPSW = 0) Or (SZ_IDX = 1)
			End If
			System.Windows.Forms.Application.DoEvents()
		Loop Until (SZ_UPSW = 0) Or (SZ_ENDSW = 1)
		
	End Sub
	
	Private Sub CMDO020_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseUpEvent) Handles CMDO020.MouseUpEvent
		SZ_UPSW = 0
		
	End Sub
	
	Private Sub CMDO030_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO030.ClickEvent
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		
		SZ_DOWNSW = 0
		
	End Sub
	
	Private Sub CMDO030_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO030.Enter
		SPRD.ROW = SZ_IDX
		SPRD.Focus()
	End Sub
	
	Private Sub CMDO030_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDO030.MouseDownEvent
		MOUSEFLG = eventArgs.Button
		
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			Exit Sub
		End If
		
		SZ_DOWNSW = 1
		
		Do 
			If SZ_IDX = SPRD.MaxRows Then
				Do 
					If SZ_IDX = SPRD.MaxRows Then
						Call SZ0414_NXT_SUB()
					End If
					If SZ_ERRSW = 1 Then
						Exit Sub
					End If
					SPRD.ROW = SZ_IDX
					SPRD.Col = 1
					SPRD.Action = SS_ACTION_ACTIVE_CELL
					SPRD.Focus()
					System.Windows.Forms.Application.DoEvents()
				Loop Until (SZ_DOWNSW = 0) Or (SZ_ENDSW = 1)
			Else
				Do 
					If (SZ_IDX < SPRD.MaxRows) Then
						SZ_IDX = SZ_IDX + 1
					End If
					SPRD.ROW = SZ_IDX
					SPRD.Col = 1
					SPRD.Action = SS_ACTION_ACTIVE_CELL
					SPRD.Focus()
					System.Windows.Forms.Application.DoEvents()
				Loop Until (SZ_DOWNSW = 0) Or (SZ_IDX = SZ0414_MAX_ROW)
			End If
			System.Windows.Forms.Application.DoEvents()
		Loop Until (SZ_DOWNSW = 0) Or (SZ_ENDSW = 1)
		
	End Sub
	
	Private Sub CMDO030_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseUpEvent) Handles CMDO030.MouseUpEvent
		SZ_DOWNSW = 0
		
	End Sub
	
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		Dim WK_DATA As Object
		Dim Ret As Short
		
		
		If MOUSEFLG = VB6.MouseButtonConstants.RightButton Then
			MOUSEFLG = VB6.MouseButtonConstants.LeftButton
			Exit Sub
		End If
		If SZ_UPSW <> n0 Or SZ_DOWNSW <> n0 Then Exit Sub
		Select Case Index
			Case 0 '終了
				SZ0414_KBN = -1
				If SPRD.MaxRows >= 1 Then
					SZ0414_OLDCOD3.Value = SJAN_K4(1)
					SZ0414_SPRD = SPRD.ActiveRow
					Call SPRD.GetText(1, 1, WK_DATA)
					'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					SZ0414_LNCNT = Val(WK_DATA)
				End If
				Call ENDRR_RTN(Me)
			Case 3 '問合せ
				Select Case LST_NO
					Case N030
						SZ_F3SW = 1
						SZ0415_TOP = VB6.PixelsToTwipsY(Me.Top)
						SZ0415_LEFT = VB6.PixelsToTwipsX(Me.Left)
						SZ0415_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
						SZ0415_WIDTH = VB6.PixelsToTwipsX(Me.Width)
						SZ0415_POS = 0
						Ret = SZ0415_SUB()
						If Ret = 0 Then
							''LST_NO = N030   'A-20130328-
							IMTX030.Text = SZ0415_SEL_CODE
							''Call SET_NO(1)
							
							NXT_NO = N040
							Call FOCUS_SET()
							
						Else
							NXT_NO = LST_NO
							Call FOCUS_SET()
						End If
						
					Case N040
						SZ_F3SW = 1
						SZ0415_TOP = VB6.PixelsToTwipsY(Me.Top)
						SZ0415_LEFT = VB6.PixelsToTwipsX(Me.Left)
						SZ0415_HEIGHT = VB6.PixelsToTwipsY(Me.Height)
						SZ0415_WIDTH = VB6.PixelsToTwipsX(Me.Width)
						SZ0415_POS = 0
						Ret = SZ0415_SUB()
						If Ret = 0 Then
							LST_NO = N040 'A-20130328-
							IMTX040.Text = SZ0415_SEL_CODE
							Call SET_NO(1)
						Else
							NXT_NO = LST_NO
							Call FOCUS_SET()
						End If
						
				End Select
			Case 5 '条件に戻る
				Call INITIAL_RTN()
				NXT_NO = N010
				Call FOCUS_SET()
			Case 6 '前一覧
				Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
				Me.Refresh()
				Call SZ0414_RDW_SUB()
				SPRD.ROW = 1
				SPRD.Focus()
				Me.Cursor = System.Windows.Forms.Cursors.Default
				Me.Refresh()
			Case 7 '次一覧
				Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
				Me.Refresh()
				Call SZ0414_RUP_SUB()
				SPRD.ROW = 1
				SPRD.Focus()
				Me.Cursor = System.Windows.Forms.Cursors.Default
				Me.Refresh()
			Case 12 '選択
				SZ0414_KBN = 0
				SZ0414_OLDCOD3.Value = SJAN_K4(1)
				SZ0414_SELCOD1.Value = SJAN_K4(SPRD.ActiveRow)
				SZ0414_SPRD = SPRD.ActiveRow
				Call SPRD.GetText(1, 1, WK_DATA)
				'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				SZ0414_LNCNT = Val(WK_DATA)
				'Call ZAHLP_Quit(Me.hwnd)
				
				''Call ENDRR_RTN(Me)    'D-20130328-
				Me.Visible = False 'A-20130328-
				
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
	
	'UPGRADE_WARNING: Form イベント SZ0414FRM.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub SZ0414FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		If SZ_F3SW = 1 Then Exit Sub
		
		'マウスカーソルを戻す
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		SPRD.Enabled = False
		CMDO020.Enabled = False
		CMDO030.Enabled = False
		
		If SZ_INTSW <> F_ON Then
			LST_NO = N010
			NXT_NO = N010
			Call FOCUS_SET()
		Else
			LST_NO = N010
			
			'        '会社コード・本支店コードが前回と違う場合、初期表示
			'        If SZ0414_OLDCOD1 <> SZ0414_KAISYAS Or SZ0414_OLDCOD2 <> SZ0414_HONSITENS Then
			'            SZ0414_OLDCOD3 = ""
			'            NXT_NO = N030
			'            Call FOCUS_SET
			'            Exit Sub
			'        End If
			
			If Val(SZ0414_OLDCOD3.Value) = 0 Then
				JAN_BUF0.k4 = ""
				JAN.k4 = ""
			Else
				JAN_BUF0.k4 = VB6.Format(Val(SZ0414_OLDCOD3.Value) - 1, "0000000000000") 'JAN
				JAN.k4 = VB6.Format(Val(SZ0414_OLDCOD3.Value) - 1, "0000000000000") 'JAN
			End If
			
			IMTX010.Text = RTrim(SZ0414_IMTX010.Value)
			IMTX020.Text = RTrim(SZ0414_IMTX020.Value)
			IMTX030.Text = RTrim(SZ0414_IMTX030.Value)
			IMTX040.Text = RTrim(SZ0414_IMTX040.Value)
			IMTX050.Text = RTrim(SZ0414_IMTX050.Value)
			IMNU060.Value = CDbl(RTrim(CStr(SZ0414_IMNU060)))
			IMNU070.Value = CDbl(RTrim(CStr(SZ0414_IMNU070)))
			Call COMBO_SETLIST_SZ0414(CMB080, RTrim(SZ0414_IMTX080.Value))
			IMNU090.Value = CDbl(RTrim(CStr(SZ0414_IMNU090)))
			IMNU100.Value = CDbl(RTrim(CStr(SZ0414_IMNU100)))
			
			Call ALLCHK_RTN()
			If SZ_ERRSW = F_ERR Then Exit Sub
			
			If SZ0414_DSPSW = F_ON Then
				SZ0414SELGE.Close()
				SZ0414SELGT.Close()
				SZ0414SELLT.Close()
			End If
			
			'ＪＡＮマスタ
			Call PREP_JAN()
			If SZ_ERRSW = F_ERR Then Call ENDRR_RTN(Me)
			
			SPRD.ReDraw = False
			SZ_LNCNT = SZ0414_LNCNT - 1
			'''B_OP% = B_GET_GE%
			'''Call SZ0414_STA_SUB(0)
			
			B_OP = B_GET_GT
			Call SZ0414_GET_SUB()
			
			If SZ_ENDSW = 1 Then
				SPRD.MaxRows = 0
				SPRD.Enabled = False
				CMDO020.Enabled = False
				CMDO030.Enabled = False
			Else
				B_OP = B_GET_NEXT
				SZ0414SEL_SW = 0
				Call SZ0414_STA_SUB(1)
				
				If SZ_ERRSW = 1 Then
					CMDOFNC(0).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
					Exit Sub
				End If
				
				If SZ_IMAX = n0 Then
					SPRD.MaxRows = 0
					SPRD.Enabled = False
					CMDO020.Enabled = False
					CMDO030.Enabled = False
				Else
					SPRD.Enabled = True
					CMDO020.Enabled = True
					CMDO030.Enabled = True
					SPRD.ROW = SZ0414_SPRD
					SPRD.Col = 1
					SPRD.Action = SS_ACTION_ACTIVE_CELL
					SPRD.Focus()
				End If
				SPRD.ReDraw = True
			End If
		End If
		
	End Sub
	
	Private Sub SZ0414FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		' ﾌｫｰﾑのｺﾝﾄﾛｰﾙﾒﾆｭｰから「閉じる」ｺﾏﾝﾄﾞが選択された場合。
		Dim WK_DATA As Object
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Then
			SZ0414_KBN = 1
			If SPRD.MaxRows >= 1 Then
				SZ0414_OLDCOD3.Value = SJAN_K4(1)
				SZ0414_SPRD = SPRD.ActiveRow
				Call SPRD.GetText(1, 1, WK_DATA)
				'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				SZ0414_LNCNT = Val(WK_DATA)
			End If
			'Call ZAHLP_Quit(Me.hwnd)
			Call ENDRR_RTN(Me)
		End If
		
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub FOCUS_SET() 'ﾌｫｰｶｽｾｯﾄ
		
		Select Case NXT_NO
			Case N010 '
				IMTX010.Focus()
			Case N020 '
				IMTX020.Focus()
			Case N030 '
				IMTX030.Focus()
			Case N040 '
				IMTX040.Focus()
			Case N050 '
				IMTX050.Focus()
			Case N060 '
				IMNU060.Focus()
			Case N070 '
				IMNU070.Focus()
			Case N080 '
				CMB080.Focus()
			Case N090 '
				IMNU090.Focus()
			Case N100 '
				IMNU100.Focus()
			Case NEND '表示ボタン
				CMDO010.Focus()
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
			
			If NXT_NO = NEND Then
				Call FOCUS_SET()
				Exit Sub
			End If
			
			If CTRLTBL(NXT_NO).CTRL.TabStop = True And CTRLTBL(NXT_NO).CTRL.Enabled = True And CTRLTBL(NXT_NO).CTRL.Visible = True Then
				Call FOCUS_SET()
				Exit Sub
			Else
				i = NXT_NO
			End If
		Loop 
		
	End Sub
	
	Private Sub FUNCSET_RTN()
		
		'ガイドメッセージ表示
		Select Case LST_NO
			Case N010, N020, N050, N060, N070, N080, N090, N100
				ZAFC_N(0) = 1
				ZAFC_N(5) = 5
			Case N030, N040
				ZAFC_N(0) = 1
				ZAFC_N(3) = 3
				ZAFC_N(5) = 5
			Case Else
				ZAFC_N(0) = 1
		End Select
		Call ZAFC_SUB(Me)
		
		
	End Sub
	
	Private Function IPROCHK() As Short 'LostFocus項目チェック
		
		IPROCHK = True
		SZ_ERRSW = F_OFF
		
		Select Case LST_NO
			Case N010 '
				Call IPROCHK_N010()
			Case N020 '
				Call IPROCHK_N020()
			Case N030 '
				Call IPROCHK_N030()
			Case N040 '
				Call IPROCHK_N040()
			Case N050 '
				Call IPROCHK_N050()
			Case N060 '
				Call IPROCHK_N060()
			Case N070 '
				Call IPROCHK_N070()
			Case N080 '
				Call IPROCHK_N080()
			Case N090 '
				Call IPROCHK_N090()
			Case N100 '
				Call IPROCHK_N100()
		End Select
		
		If SZ_ERRSW = F_ERR Then
			If CUR_NO < LST_NO Then
				Select Case LST_NO
					Case N010
						IMTX010.Text = RTrim(WKBSZ0414.S010)
					Case N020
						IMTX020.Text = RTrim(WKBSZ0414.S020)
					Case N030
						IMTX030.Text = RTrim(WKBSZ0414.S030)
						DSP030.Text = RTrim(WKBSZ0414.S030N)
					Case N040
						IMTX040.Text = RTrim(WKBSZ0414.S040)
						DSP040.Text = RTrim(WKBSZ0414.S040N)
					Case N050
						IMTX050.Text = RTrim(WKBSZ0414.S050)
					Case N060
						'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						IMNU060.CtlText = RTrim(CStr(WKBSZ0414.C060))
					Case N070
						'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						IMNU070.CtlText = RTrim(CStr(WKBSZ0414.C070))
					Case N080
						
					Case N090
						'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						IMNU090.CtlText = RTrim(CStr(WKBSZ0414.C090))
					Case N070
						'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						IMNU100.CtlText = RTrim(CStr(WKBSZ0414.C100))
						
				End Select
			Else
				IPROCHK = False
				NXT_NO = LST_NO
				Call FOCUS_SET()
			End If
		End If
		
	End Function
	
	Private Sub IPROCHK_N010() '
		
		WKBSZ0414.S010 = IMTX010.Text
		
		If Trim(IMTX020.Text) = "" Then
			IMTX020.Text = IMTX010.Text
			WKBSZ0414.S020 = IMTX020.Text
		End If
		
	End Sub
	
	Private Sub IPROCHK_N020() '
		
		WKBSZ0414.S020 = IMTX020.Text
		
	End Sub
	
	Private Sub IPROCHK_N030() '
		'商品部類
		SZ_F3SW = 0
		
		If Trim(IMTX030.Text) <> "" Then
			JAN_BUNRUI_BUF0.BK1 = IMTX030.Text
			Call RD_JANBUNRUI()
			If SZ_ERRSW = F_ERR Then Call ENDRR_RTN(Me)
			If JAN_BUNRUIINVSW = F_INV Then
				WKBSZ0414.S030N = " "
				DSP030.Text = WKBSZ0414.S030N
				SZ_ERRSW = F_ERR
				Exit Sub
			Else
				WKBSZ0414.S030N = RTrim(JAN_BUNRUI.BK4)
			End If
		Else
			WKBSZ0414.S030N = ""
		End If
		
		WKBSZ0414.S030 = IMTX030.Text
		DSP030.Text = WKBSZ0414.S030N
		
		If Trim(IMTX040.Text) = "" Then
			IMTX040.Text = IMTX030.Text
			WKBSZ0414.S040 = IMTX040.Text
			WKBSZ0414.S040N = WKBSZ0414.S030N
			DSP040.Text = WKBSZ0414.S040N
		End If
		
	End Sub
	
	Private Sub IPROCHK_N040() '
		SZ_F3SW = 0
		
		If Trim(IMTX040.Text) <> "" Then
			JAN_BUNRUI_BUF0.BK1 = IMTX040.Text
			Call RD_JANBUNRUI()
			If SZ_ERRSW = F_ERR Then Call ENDRR_RTN(Me)
			If JAN_BUNRUIINVSW = F_INV Then
				WKBSZ0414.S040N = " "
				DSP040.Text = WKBSZ0414.S040N
				SZ_ERRSW = F_ERR
				Exit Sub
			Else
				WKBSZ0414.S040N = RTrim(JAN_BUNRUI.BK4)
			End If
		Else
			WKBSZ0414.S040N = ""
		End If
		
		WKBSZ0414.S040 = IMTX040.Text
		DSP040.Text = WKBSZ0414.S040N
		
	End Sub
	
	Private Sub IPROCHK_N050() '
		
		'    WKBSZ0414.S050 = IMTX050.Text      'D-20130401-
		
		WKBSZ0414.S050 = StrConv(IMTX050.Text, VbStrConv.UpperCase) 'A-20130401-
		IMTX050.Text = WKBSZ0414.S050 'A-20130401-
		
	End Sub
	
	Private Sub IPROCHK_N060() '
		
		WKBSZ0414.C060 = IMNU060.Value
		
		If IMNU070.Value = 0 Then
			IMNU070.Value = IMNU060.Value
			WKBSZ0414.C070 = IMNU070.Value
		End If
		
	End Sub
	
	Private Sub IPROCHK_N070() '
		
		WKBSZ0414.C070 = IMNU070.Value
		
	End Sub
	
	Private Sub IPROCHK_N080() '
		
		If CMB080.SelectedIndex < 0 Then
			WKBSZ0414.S080 = "0"
		Else
			WKBSZ0414.S080 = CStr(VB6.GetItemData(CMB080, CMB080.SelectedIndex))
		End If
		
	End Sub
	
	Private Sub IPROCHK_N090() '
		
		'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		WKBSZ0414.C090 = CDec(IMNU090.CtlText)
		
		If IMNU100.Value = 0 Then
			IMNU100.Value = IMNU090.Value
			WKBSZ0414.C100 = IMNU100.Value
		End If
		
		
	End Sub
	
	Private Sub IPROCHK_N100() '
		
		WKBSZ0414.C100 = IMNU100.Value
		
	End Sub
	
	
	Private Function GPROCHK() As Short 'LostFocus項目群チェック
		
		GPROCHK = True
		SZ_ERRSW = F_OFF
		
		If CTRLTBL(CUR_NO).IGRP <= CTRLTBL(LST_NO).IGRP Then
			Exit Function
		End If
		Select Case CTRLTBL(LST_NO).IGRP
			Case GRP1
				Call GPROCHK_GRP1()
		End Select
		If SZ_ERRSW = F_ERR Then
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = False
			GPROCHK = False
		Else
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = True
		End If
		
	End Function
	
	Private Sub GPROCHK_GRP1()
		
		
		Call IPROCHK_N010()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N010
			Exit Sub
		End If
		
		Call IPROCHK_N020()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N020
			Exit Sub
		End If
		
		Call IPROCHK_N030()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N030
			Exit Sub
		End If
		
		Call IPROCHK_N040()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N040
			Exit Sub
		End If
		
		Call IPROCHK_N050()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N050
			Exit Sub
		End If
		
		Call IPROCHK_N060()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N060
			Exit Sub
		End If
		
		Call IPROCHK_N070()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N070
			Exit Sub
		End If
		
		Call IPROCHK_N080()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N080
			Exit Sub
		End If
		
		Call IPROCHK_N090()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N090
			Exit Sub
		End If
		
		Call IPROCHK_N100()
		If SZ_ERRSW = F_ERR Then
			GRPTBL(GRP1).NXTN = N100
			Exit Sub
		End If
		
	End Sub
	
	Private Function GVALCHK() As Short '項目群入力可否チェック
		
		GVALCHK = True
		SZ_ERRSW = F_OFF
		
		If LST_NO <> n0 Then
			If CTRLTBL(CUR_NO).IGRP = CTRLTBL(LST_NO).IGRP Then Exit Function
		End If
		Select Case CTRLTBL(CUR_NO).IGRP
			Case GRP1
				Call GVALCHK_GRP1()
		End Select
		If SZ_ERRSW = F_ERR Then
			GRPTBL(CTRLTBL(CUR_NO).IGRP).CFLG = False
			GVALCHK = False
		Else
			GRPTBL(CTRLTBL(CUR_NO).IGRP).CFLG = True
		End If
		
	End Function
	
	Private Sub GVALCHK_GRP1()
	End Sub
	
	Private Function MVALCHK() As Short '項目入力可否チェック
		
		MVALCHK = True
		SZ_ERRSW = F_OFF
		
		Select Case CUR_NO
			Case N010 '
				Call MVALCHK_N010()
			Case N020 '
				Call MVALCHK_N020()
			Case N030 '
				Call MVALCHK_N030()
			Case N040 '
				Call MVALCHK_N040()
			Case N050 '
				Call MVALCHK_N050()
			Case N060 '
				Call MVALCHK_N060()
			Case N070 '
				Call MVALCHK_N070()
			Case N080 '
				Call MVALCHK_N080()
			Case N090 '
				Call MVALCHK_N090()
			Case N100 '
				Call MVALCHK_N100()
		End Select
		
		If SZ_ERRSW = F_ERR Then
			MVALCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
		
	End Function
	
	Sub MVALCHK_N010() '
	End Sub
	Sub MVALCHK_N020() '
	End Sub
	Sub MVALCHK_N030() '
	End Sub
	Sub MVALCHK_N040() '
	End Sub
	Sub MVALCHK_N050() '
	End Sub
	Sub MVALCHK_N060() '
	End Sub
	Sub MVALCHK_N070() '
	End Sub
	Sub MVALCHK_N080() '
	End Sub
	Sub MVALCHK_N090() '
		'    If CMB080.ListIndex < 0 Then
		'        SZ_ERRSW = F_ERR
		'    End If
	End Sub
	Sub MVALCHK_N100() '
		'    If CMB080.ListIndex < 0 Then
		'        SZ_ERRSW = F_ERR
		'    End If
	End Sub
	
	'
	'全チェック & 実行
	'
	Private Sub ALLCHK_RTN()
		
		CUR_NO = NEND
		
		'直前項目のチェック
		If GPROCHK() = False Then
		End If
		
		'全グループのチェック
		If GRPTBL(GRP1).CFLG = False Then
			SZ_ERRSW = F_ERR
			ZAER_CD = 120
			ZAER_KN = 0
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			NXT_NO = GRPTBL(GRP1).NXTN
			Call FOCUS_SET()
			Exit Sub
		End If
		
		SZ_ERRSW = F_ERR
		
		'UPGRADE_WARNING: オブジェクト KBSZ0414 の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		KBSZ0414 = WKBSZ0414
		
		SZ_ERRSW = F_OFF
		
	End Sub
	
	
	Private Sub SZ0414FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If Me.Enabled = False Then Exit Sub ' A-98.11
		If Shift <> n0 Then Exit Sub
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
			Case System.Windows.Forms.Keys.Return
				Call SET_NO(1) ' 次項目
			Case System.Windows.Forms.Keys.Up
				Call SET_NO(2) ' 前項目
			Case System.Windows.Forms.Keys.Down
				Call SET_NO(3) ' 次グループ
				KeyCode = 0
			Case System.Windows.Forms.Keys.PageDown
				If CMDOFNC(7).Text <> "" Then
					CMDOFNC(7).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.PageUp
				If CMDOFNC(6).Text <> "" Then
					CMDOFNC(6).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.End
				'        If OSKCMNV5.ZADH_HLP = True Then
				'            If CMDOFNC(3).Caption <> "" Then
				'                CMDOFNC(3).SetFocus
				'            End If
				'            Call CMDOFNC_Click(3)
				'            KeyCode = n0
				'        End If
			Case System.Windows.Forms.Keys.F3
				If CMDOFNC(3).Text <> "" Then
					CMDOFNC(3).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(3), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.F5
				If CMDOFNC(5).Text <> "" Then
					CMDOFNC(5).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.F6
				If CMDOFNC(6).Text <> "" Then
					CMDOFNC(6).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.F7
				If CMDOFNC(7).Text <> "" Then
					CMDOFNC(7).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
				KeyCode = 0
			Case System.Windows.Forms.Keys.F12
				If CMDOFNC(12).Text <> "" Then
					CMDOFNC(12).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				KeyCode = 0
		End Select
		
		
		
	End Sub
	
	Private Sub SZ0414FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'マウスカーソルを砂時計に設定
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		'共通アイコンの表示
		Call ZAWC_SUB(Me, 3)
		Call TBL_SET() '画面ｺﾝﾄﾛｰﾙ初期設定
		Call INITIAL_RTN() '初期画面表示
		
		
		If InitFlg Then
			InitFlg = Not (InitFlg)
			GoTo Form_Load_START
		End If
		Call STARTRR_RTN(Me)
		Exit Sub
		
Form_Load_START: 
		Select Case SZ0414_PS
			Case 0 '中央
				Me.Top = VB6.TwipsToPixelsY(SZ0414_TOPS + ((SZ0414_HEIGHTS - VB6.PixelsToTwipsY(Me.Height)) \ 2))
				Me.Left = VB6.TwipsToPixelsX(SZ0414_LEFTS + ((SZ0414_WIDTHS - VB6.PixelsToTwipsX(Me.Width)) \ 2))
			Case 1 '左上
				Me.Top = VB6.TwipsToPixelsY(SZ0414_TOPS + 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0414_LEFTS + 200)
			Case 2 '右上
				Me.Top = VB6.TwipsToPixelsY(SZ0414_TOPS + 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0414_LEFTS + SZ0414_WIDTHS - VB6.PixelsToTwipsX(Me.Width) - 200)
			Case 3 '左下
				Me.Top = VB6.TwipsToPixelsY(SZ0414_TOPS + SZ0414_HEIGHTS - VB6.PixelsToTwipsY(Me.Height) - 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0414_LEFTS + 200)
			Case 4 '右下
				Me.Top = VB6.TwipsToPixelsY(SZ0414_TOPS + SZ0414_HEIGHTS - VB6.PixelsToTwipsY(Me.Height) - 300)
				Me.Left = VB6.TwipsToPixelsX(SZ0414_LEFTS + SZ0414_WIDTHS - VB6.PixelsToTwipsX(Me.Width) - 200)
		End Select
		
	End Sub
	Private Sub STARTRR_RTN(ByRef MyForm As System.Windows.Forms.Form)
		'
		' ｺｰﾄﾞ問合せ ﾌｫｰﾑ開始時処理
		'
		'*** ｳｲﾝﾄﾞｳ表示位置ｾｰﾌﾞ
		Dim Ret As Integer
		Dim nWidth As Integer
		Dim nHeight As Integer
		
		'*** ｱｲｺﾝ(ﾊﾟｽ+ﾌｧｲﾙ名)
		Dim GETWORK As String
		
		
		On Error GoTo STARTRR_ERROR
		
		'*** 前回表示位置、ｻｲｽﾞで再表示
		'    nWidth = lpRectSave.Right - lpRectSave.Left
		'    nHeight = lpRectSave.Bottom - lpRectSave.Top
		'    Ret = MoveWindow(MyForm.hwnd, lpRectSave.Left, lpRectSave.Top, _
		''          nWidth, nHeight, True)
		
		Exit Sub
		
STARTRR_ERROR: 
		
	End Sub
	
	Private Sub IMNU060_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU060.KeyPressEvent
		Call ZAKB_SUB(eventArgs.KeyAscii)
		
	End Sub
	
	Private Sub IMNU070_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU070.KeyPressEvent
		Call ZAKB_SUB(eventArgs.KeyAscii)
		
	End Sub
	
	Private Sub IMNU090_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU090.KeyPressEvent
		Call ZAKB_SUB(eventArgs.KeyAscii)
		
	End Sub
	
	Private Sub IMNU100_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyPressEvent) Handles IMNU100.KeyPressEvent
		Call ZAKB_SUB(eventArgs.KeyAscii)
		
	End Sub
	
	Private Sub IMTX010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX010.Enter
		
		If CUR_NO = N010 Then Exit Sub
		CUR_NO = N010
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N010
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub IMTX010_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX010.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	Private Sub IMTX020_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX020.Enter
		
		If CUR_NO = N020 Then Exit Sub
		CUR_NO = N020
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N020
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub IMTX020_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX020.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	Private Sub IMTX030_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX030.Enter
		
		If CUR_NO = N030 Then Exit Sub
		CUR_NO = N030
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N030
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub IMTX030_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX030.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	Private Sub IMTX040_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX040.Enter
		
		If CUR_NO = N040 Then Exit Sub
		CUR_NO = N040
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N040
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub IMTX040_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX040.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	Private Sub IMTX050_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTX050.Enter
		
		If CUR_NO = N050 Then Exit Sub
		CUR_NO = N050
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N050
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	Private Sub IMTX050_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsktxtLibV5.__ImText_KeyDownEvent) Handles IMTX050.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	Private Sub IMNU060_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU060.Enter
		'
		If CUR_NO = N060 Then Exit Sub
		CUR_NO = N060
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N060
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		ZAKB_SW = 0
		
	End Sub
	
	Private Sub IMNU060_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU060.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	
	Private Sub IMNU070_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU070.Enter
		'
		
		If CUR_NO = N070 Then Exit Sub
		CUR_NO = N070
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N070
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		ZAKB_SW = 0
		
	End Sub
	
	Private Sub IMNU070_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU070.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	
	Private Sub CMB080_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMB080.Enter
		'
		
		If CUR_NO = N080 Then Exit Sub
		CUR_NO = N080
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N080
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		
	End Sub
	
	
	Private Sub CMB080_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CMB080.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
		
	End Sub
	
	
	Private Sub IMNU090_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU090.Enter
		'
		
		If CUR_NO = N090 Then Exit Sub
		CUR_NO = N090
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N090
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		ZAKB_SW = 0
		
	End Sub
	
	Private Sub IMNU090_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU090.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	
	Private Sub IMNU100_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMNU100.Enter
		'
		
		If CUR_NO = N100 Then Exit Sub
		CUR_NO = N100
		
		'チェック
		If LST_NO <> n0 Then
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
		LST_NO = N100
		
		'ファンクションガイド
		Call FUNCSET_RTN()
		ZAKB_SW = 0
		
	End Sub
	
	Private Sub IMNU100_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOsknumLibV5.__ImNumber_KeyDownEvent) Handles IMNU100.KeyDownEvent
		
		Call SZ0414FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	
	Private Sub SPRD_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SPRD.ClickEvent
		SZ_IDX = SPRD.ActiveRow
		
	End Sub
	
	
	Private Sub SPRD_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SPRD.DblClick
		If eventArgs.ROW <> 0 Then
			Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
		End If
		
	End Sub
	
	
	Private Sub SPRD_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD.Enter
		'ファンクションガイド
		ZAFC_N(0) = 1
		ZAFC_N(5) = 5
		ZAFC_N(6) = 6
		ZAFC_N(7) = 7
		ZAFC_N(12) = 11
		Call ZAFC_SUB(Me)
		
		SZ_IDX = SPRD.ActiveRow
		
	End Sub
	
	
	Private Sub SPRD_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SPRD.KeyDownEvent
		
		If eventArgs.Shift <> n0 Then Exit Sub
		If SZ_UPSW <> n0 Or SZ_DOWNSW <> n0 Then Exit Sub
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Escape
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
				End If
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
			Case System.Windows.Forms.Keys.Down
				If SZ_IDX = SZ0414_MAX_ROW Then
					Call SZ0414_NXT_SUB()
					SPRD.ROW = SZ0414_MAX_ROW
					SPRD.Focus()
				Else
					If SZ_IDX < SPRD.MaxRows Then
						SZ_IDX = SZ_IDX + 1
					End If
				End If
			Case System.Windows.Forms.Keys.Up
				If SZ_IDX = 1 Then
					Call SZ0414_PRE_SUB()
					SPRD.ROW = 1
					SPRD.Focus()
				Else
					SZ_IDX = SZ_IDX - 1
				End If
			Case System.Windows.Forms.Keys.F5
				CMDOFNC(5).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
			Case System.Windows.Forms.Keys.F6, System.Windows.Forms.Keys.PageUp
				CMDOFNC(6).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
			Case System.Windows.Forms.Keys.F7, System.Windows.Forms.Keys.PageDown
				CMDOFNC(7).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
			Case System.Windows.Forms.Keys.F12, System.Windows.Forms.Keys.Return
				CMDOFNC(12).Focus()
				Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
		End Select
		
	End Sub
	
	
	Private Sub SZ0414_DISP_SUB()
		Dim WK_VAL As Object
		
		'ＪＡＮ
		'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		WK_VAL = JAN.k4
		Call SPRD.SetText(1, SZ_I, WK_VAL)
		
		'商品部類
		JAN_BUNRUI_BUF0.BK1 = JAN.k21
		Call RD_JANBUNRUI()
		If SZ_ERRSW = F_ERR Then Call ENDRR_RTN(Me)
		If JAN_BUNRUIINVSW = F_INV Then
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = " "
		Else
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = RTrim(JAN_BUNRUI.BK4)
		End If
		Call SPRD.SetText(2, SZ_I, WK_VAL)
		
		'商品名
		'    WK_VAL = RTrim$(JAN.k20)   'D-20130424-
		'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		WK_VAL = RTrim(JAN.k17) 'A-20130424-
		Call SPRD.SetText(3, SZ_I, WK_VAL)
		
		'原産国
		'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		WK_VAL = RTrim(JAN.k44)
		Call SPRD.SetText(4, SZ_I, WK_VAL)
		
		'重量
		'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		WK_VAL = VB6.Format(RTrim(CStr(JAN.k42)), "#,###")
		Call SPRD.SetText(5, SZ_I, WK_VAL)
		
		'有効期限
		If RTrim(JAN.k57) = "1" Then
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = "日"
		ElseIf RTrim(JAN.k57) = "2" Then 
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = "月"
		ElseIf RTrim(JAN.k57) = "3" Then 
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = "年"
		Else
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = "　"
		End If
		If Val(CStr(JAN.k58)) = 0 Then
		Else
			'UPGRADE_WARNING: オブジェクト WK_VAL の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			WK_VAL = WK_VAL & Space(2) & VB6.Format(RTrim(CStr(JAN.k58)), "@@@")
		End If
		Call SPRD.SetText(6, SZ_I, WK_VAL)
		
		'キーの保存
		SJAN_K4(SZ_I) = JAN.k4 '
		
	End Sub
	
	Private Sub SZ0414_STA_SUB(ByRef OP As Short)
		
		If OP = 1 Then
			GoTo STA_DISP
		End If
		
		'キーの保存
		
		SZ_I = 0
		SZ_IMAX = 0
		SZ0414SEL_SW = 0
		B_OP = B_GET_GE
		Call SZ0414_GET_SUB()
		If SZ_ERRSW = 1 Then
			Exit Sub
		End If
		If SZ_ENDSW = 1 Then
			SZ_IMAX = 0
			SZ_ENDSW = 0
			Exit Sub
		End If
		
STA_DISP: 
		SZ_I = 1
		SZ_IMAX = 1
		SZ_LNCNT = SZ_LNCNT + 1
		SPRD.MaxRows = 1
		Call SPRD_ROWHT_RTN()
		Call SZ0414_DISP_SUB()
		
		Do 
			
			SZ_I = SZ_I + 1
			SZ_LNCNT = SZ_LNCNT + 1
			If HOUKOU = 0 Then
				B_OP = B_GET_NEXT
			Else
				JAN.k4 = SJAN_K4(SZ_IMAX)
				B_OP = B_GET_GT
				SZ0414SEL_SW = 0
			End If
			Call SZ0414_GET_SUB()
			If SZ_ERRSW = 1 Then
				Exit Sub
			End If
			If SZ_ENDSW = 1 Then
				SZ_ENDSW = 0
				Exit Sub
			End If
			SZ_IMAX = SZ_I
			SPRD.MaxRows = SZ_IMAX
			Call SPRD_ROWHT_RTN()
			Call SZ0414_DISP_SUB()
			
		Loop Until SZ_I = SZ0414_MAX_ROW
		
		
		
	End Sub
	Private Sub SZ0414_ERR_SUB()
		
		ZAER_KN = 1
		ZAER_NO.Value = "JAN"
		Call ZAER_SUB()
		SZ_ERRSW = 1
		
	End Sub
	
	Private Sub SZ0414_GET_SUB()
		Dim WK_STATUS As Short
		
		
		'*** コード順
		
		Select Case B_OP
			Case B_GET_GE
				HOUKOU = 0
				FETCH_MODE = "SZ0414SELGE"
				If RTrim(SZ0414_IMTX010.Value) <> "" Then
					SZ0414SELGE.rdoParameters("k4F").Value = RTrim(SZ0414_IMTX010.Value)
				End If
				If RTrim(SZ0414_IMTX020.Value) <> "" Then
					SZ0414SELGE.rdoParameters("k4T").Value = RTrim(SZ0414_IMTX020.Value)
				End If
				If RTrim(SZ0414_IMTX030.Value) <> "" Then
					SZ0414SELGE.rdoParameters("k21F").Value = RTrim(SZ0414_IMTX030.Value)
				End If
				If RTrim(SZ0414_IMTX040.Value) <> "" Then
					SZ0414SELGE.rdoParameters("k21T").Value = RTrim(SZ0414_IMTX040.Value)
				End If
				If RTrim(SZ0414_IMTX050.Value) <> "" Then
					SZ0414SELGE.rdoParameters("k44").Value = RTrim(SZ0414_IMTX050.Value)
				End If
				If (SZ0414_IMNU060) <> 0 Then
					SZ0414SELGE.rdoParameters("K42F").Value = RTrim(CStr(SZ0414_IMNU060))
				End If
				If (SZ0414_IMNU070) <> 0 Then
					SZ0414SELGE.rdoParameters("K42T").Value = RTrim(CStr(SZ0414_IMNU070))
				End If
				'            If RTrim$(SZ0414_IMTX080) <> "0" Then
				'                SZ0414SELGE!k57 = RTrim$(SZ0414_IMTX080)
				'            End If
				'            If (SZ0414_IMNU090) <> 0 Then
				'                SZ0414SELGE!K58F = RTrim$(SZ0414_IMNU090)
				'            End If
				'            If (SZ0414_IMNU100) <> 0 Then
				'                SZ0414SELGE!K58T = RTrim$(SZ0414_IMNU100)
				'            End If
				If (SZ0414_IMNU090D) <> 0 Then
					SZ0414SELGE.rdoParameters("K99F").Value = RTrim(CStr(SZ0414_IMNU090D))
				End If
				If (SZ0414_IMNU100D) <> 0 Then
					SZ0414SELGE.rdoParameters("K99T").Value = RTrim(CStr(SZ0414_IMNU100D))
				End If
				
				On Error Resume Next
				JANRSSW = "SZ0414SELGE"
				SZ0414RES = SZ0414SELGE.OpenResultset()
				
			Case B_GET_GT
				HOUKOU = 0
				FETCH_MODE = "SZ0414SELGT"
				
				SZ0414SELGT.rdoParameters("k4").Value = JAN.k4
				If RTrim(SZ0414_IMTX010.Value) <> "" Then
					SZ0414SELGT.rdoParameters("k4F").Value = RTrim(SZ0414_IMTX010.Value)
				End If
				If RTrim(SZ0414_IMTX020.Value) <> "" Then
					SZ0414SELGT.rdoParameters("k4T").Value = RTrim(SZ0414_IMTX020.Value)
				End If
				If RTrim(SZ0414_IMTX030.Value) <> "" Then
					SZ0414SELGT.rdoParameters("k21F").Value = RTrim(SZ0414_IMTX030.Value)
				End If
				If RTrim(SZ0414_IMTX040.Value) <> "" Then
					SZ0414SELGT.rdoParameters("k21T").Value = RTrim(SZ0414_IMTX040.Value)
				End If
				If RTrim(SZ0414_IMTX050.Value) <> "" Then
					SZ0414SELGT.rdoParameters("k44").Value = RTrim(SZ0414_IMTX050.Value)
				End If
				If (SZ0414_IMNU060) <> 0 Then
					SZ0414SELGT.rdoParameters("K42F").Value = RTrim(CStr(SZ0414_IMNU060))
				End If
				If (SZ0414_IMNU070) <> 0 Then
					SZ0414SELGT.rdoParameters("K42T").Value = RTrim(CStr(SZ0414_IMNU070))
				End If
				'            If RTrim$(SZ0414_IMTX080) <> "0" Then
				'                SZ0414SELGT!k57 = RTrim$(SZ0414_IMTX080)
				'            End If
				'            If (SZ0414_IMNU090) <> 0 Then
				'                SZ0414SELGT!K58F = RTrim$(SZ0414_IMNU090)
				'            End If
				'            If (SZ0414_IMNU100) <> 0 Then
				'                SZ0414SELGT!K58T = RTrim$(SZ0414_IMNU100)
				'            End If
				If (SZ0414_IMNU090D) <> 0 Then
					SZ0414SELGT.rdoParameters("K99F").Value = RTrim(CStr(SZ0414_IMNU090D))
				End If
				If (SZ0414_IMNU100D) <> 0 Then
					SZ0414SELGT.rdoParameters("K99T").Value = RTrim(CStr(SZ0414_IMNU100D))
				End If
				On Error Resume Next
				JANRSSW = "SZ0414SELGT"
				SZ0414RES = SZ0414SELGT.OpenResultset()
				
			Case B_GET_LT
				HOUKOU = 1
				FETCH_MODE = "SZ0414SELLT"
				
				SZ0414SELLT.rdoParameters("k4").Value = JAN.k4
				If RTrim(SZ0414_IMTX010.Value) <> "" Then
					SZ0414SELLT.rdoParameters("k4F").Value = RTrim(SZ0414_IMTX010.Value)
				End If
				If RTrim(SZ0414_IMTX020.Value) <> "" Then
					SZ0414SELLT.rdoParameters("k4T").Value = RTrim(SZ0414_IMTX020.Value)
				End If
				If RTrim(SZ0414_IMTX030.Value) <> "" Then
					SZ0414SELLT.rdoParameters("k21F").Value = RTrim(SZ0414_IMTX030.Value)
				End If
				If RTrim(SZ0414_IMTX040.Value) <> "" Then
					SZ0414SELLT.rdoParameters("k21T").Value = RTrim(SZ0414_IMTX040.Value)
				End If
				If RTrim(SZ0414_IMTX050.Value) <> "" Then
					SZ0414SELLT.rdoParameters("k44").Value = RTrim(SZ0414_IMTX050.Value)
				End If
				If (SZ0414_IMNU060) <> 0 Then
					SZ0414SELLT.rdoParameters("K42F").Value = RTrim(CStr(SZ0414_IMNU060))
				End If
				If (SZ0414_IMNU070) <> 0 Then
					SZ0414SELLT.rdoParameters("K42T").Value = RTrim(CStr(SZ0414_IMNU070))
				End If
				'            If RTrim$(SZ0414_IMTX080) <> "0" Then
				'                SZ0414SELLT!k57 = RTrim$(SZ0414_IMTX080)
				'            End If
				'            If (SZ0414_IMNU090) <> 0 Then
				'                SZ0414SELLT!K58F = RTrim$(SZ0414_IMNU090)
				'            End If
				'            If (SZ0414_IMNU100) <> 0 Then
				'                SZ0414SELLT!K58T = RTrim$(SZ0414_IMNU100)
				'            End If
				If (SZ0414_IMNU090D) <> 0 Then
					SZ0414SELLT.rdoParameters("K99F").Value = RTrim(CStr(SZ0414_IMNU090D))
				End If
				If (SZ0414_IMNU100D) <> 0 Then
					SZ0414SELLT.rdoParameters("K99T").Value = RTrim(CStr(SZ0414_IMNU100D))
				End If
				On Error Resume Next
				JANRSSW = "SZ0414SELLT"
				SZ0414RES = SZ0414SELLT.OpenResultset()
				
			Case B_GET_NEXT, B_GET_PRE
				On Error Resume Next
				SZ0414RES.MoveNext()
				
		End Select
		
		RDO_STATUS = B_STATUS(SZ0414RES)
		Select Case RDO_STATUS
			Case 0
				' 終了区分
				SZ_ENDSW = 0
				JAN.k4 = ""
				JAN.k24 = ""
				'        JAN.k20 = ""       'D-20130424-
				JAN.k17 = "" 'A-20130424-
				JAN.k44 = ""
				JAN.k42 = CDec("")
				JAN.k57 = ""
				JAN.k58 = CDec("")
				JAN.k4 = SZ0414RES.rdoColumns("k4").Value
				JAN.k21 = SZ0414RES.rdoColumns("k21").Value
				'        JAN.k20 = SZ0414RES!k20    'D-20130424-
				JAN.k17 = SZ0414RES.rdoColumns("k17").Value 'A-20130424-
				JAN.k44 = SZ0414RES.rdoColumns("k44").Value
				JAN.k42 = SZ0414RES.rdoColumns("k42").Value
				JAN.k57 = SZ0414RES.rdoColumns("k57").Value
				JAN.k58 = SZ0414RES.rdoColumns("k58").Value
				JAN.k99 = SZ0414RES.rdoColumns("k99").Value
				
			Case 24
				SZ_ENDSW = 1
			Case Else
				Call SZ0414_ERR_SUB()
		End Select
		On Error GoTo 0
		
		
	End Sub
	
	Private Sub SZ0414_NXT_SUB()
		'
		' １０行目で下矢印を押したときの処理
		'
		Dim SZ_COL As Short
		Dim WK_DATA As Object
		
		SPRD.ReDraw = False
		
		If RDO_STATUS <> 0 And HOUKOU = 0 Then
			' 前回 at end 状態なら読み込まないようにする
			Exit Sub
		End If
		If HOUKOU = 0 Then
			B_OP = B_GET_NEXT
		Else
			JAN.k4 = SJAN_K4(SZ_IMAX)
			B_OP = B_GET_GT
			SZ0414SEL_SW = 0
		End If
		
		Call SZ0414_GET_SUB()
		If (SZ_ENDSW = 1) Or (SZ_ERRSW = 1) Then
			SZ_ENDSW = 0
			Exit Sub
		End If
		B_OP = B_GET_NEXT
		
		SZ_I = 1
		Do 
			' １～１０行までスライド
			For SZ_COL = 1 To SPRD.MaxCols
				If SZ_COL = 1 Then
					Call SPRD.GetText(SZ_COL, SZ_I + 1, WK_DATA)
					'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					SZ_LNCNT = Val(WK_DATA)
				End If
				Call SPRD.GetText(SZ_COL, SZ_I + 1, WK_DATA)
				Call SPRD.SetText(SZ_COL, SZ_I, WK_DATA)
			Next SZ_COL
			SJAN_K4(SZ_I) = SJAN_K4(SZ_I + 1)
			SZ_I = SZ_I + 1
		Loop Until SZ_I = SZ0414_MAX_ROW
		SZ_LNCNT = SZ_LNCNT + 1
		Call SZ0414_DISP_SUB()
		
		SPRD.ReDraw = True
	End Sub
	
	Private Sub SZ0414_PRE_SUB()
		'
		' １行目で上矢印を押したときの処理
		'
		Dim SZ_COL As Short
		Dim WK_DATA As Object
		
		SPRD.ReDraw = False
		
		
		If RDO_STATUS <> 0 And HOUKOU = 1 Then
			' 前回 at end 状態なら読み込まないようにする
			Exit Sub
		End If
		If HOUKOU = 1 Then
			B_OP = B_GET_PRE
		Else
			JAN.k4 = SJAN_K4(1)
			B_OP = B_GET_LT
			SZ0414SEL_SW = 0
		End If
		
		Call SZ0414_GET_SUB()
		If (SZ_ENDSW = 1) Or (SZ_ERRSW = 1) Then
			SZ_ENDSW = 0
			Exit Sub
		End If
		B_OP = B_GET_PRE
		
		If SPRD.MaxRows <> SZ0414_MAX_ROW Then
			SPRD.MaxRows = SPRD.MaxRows + 1
		End If
		Call SPRD_ROWHT_RTN()
		SZ_I = SPRD.MaxRows
		
		Do 
			' １～１０行までスライド
			For SZ_COL = 1 To SPRD.MaxCols
				If SZ_COL = 1 Then
					Call SPRD.GetText(SZ_COL, SZ_I - 1, WK_DATA)
					'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					SZ_LNCNT = Val(WK_DATA)
				End If
				Call SPRD.GetText(SZ_COL, SZ_I - 1, WK_DATA)
				Call SPRD.SetText(SZ_COL, SZ_I, WK_DATA)
			Next SZ_COL
			SJAN_K4(SZ_I) = SJAN_K4(SZ_I - 1)
			SZ_I = SZ_I - 1
		Loop Until SZ_I = 1
		SZ_I = 1
		If SZ_LNCNT <> 1 Then
			SZ_LNCNT = SZ_LNCNT - 1
		End If
		Call SZ0414_DISP_SUB()
		SZ_IMAX = SPRD.MaxRows
		
		SPRD.ReDraw = True
		
	End Sub
	
	Private Sub SZ0414_RUP_SUB()
		'
		' 次一覧
		'
		Dim WK_DATA As Object
		
		
		SPRD.ReDraw = False
		
		If SZ_IMAX = SZ0414_MAX_ROW Then
			If RDO_STATUS <> 0 And HOUKOU = 0 Then
				' 前回 at end 状態なら読み込まないようにする
				Exit Sub
			End If
			If HOUKOU = 0 Then
				B_OP = B_GET_NEXT
			Else
				'If Trim$(SMCM97_001) = "" And Trim$(SMCM97_002) = "" And Trim$(SMCM97_003(SZ_IMAX)) = "" Then
				If Trim(SJAN_K4(SZ_IMAX)) = "" Then
					Exit Sub
				Else
					JAN.k4 = SJAN_K4(SZ_IMAX)
					B_OP = B_GET_GT
					SZ0414SEL_SW = 0
					'行番号の取得
					Call SPRD.GetText(1, SZ_IMAX, WK_DATA)
					'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					SZ_LNCNT = Val(WK_DATA)
				End If
			End If
			Call SZ0414_GET_SUB()
			If (SZ_ERRSW = 1) Or (SZ_ENDSW = 1) Then
				Exit Sub
			End If
			Call SZ0414_STA_SUB(1)
		End If
		
		SPRD.ReDraw = True
		
	End Sub
	
	Private Sub SZ0414_RDW_SUB()
		'
		' 前一覧
		'
		Dim WK_DATA As Object
		
		
		SPRD.ReDraw = False
		For SZ_REP = 1 To (SZ0414_MAX_ROW) Step 1
			If RDO_STATUS <> 0 And FETCH_MODE = "SZ0414SELLT" Then
				' 前回 at end 状態なら読み込まないようにする
				Exit Sub
			End If
			If HOUKOU = 1 Then
				B_OP = B_GET_PRE
			Else
				JAN.k4 = SJAN_K4(1)
				B_OP = B_GET_LT
				SZ0414SEL_SW = 0
			End If
			'行番号の取得
			Call SPRD.GetText(1, 1, WK_DATA)
			'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			If Val(WK_DATA) > SZ0414_MAX_ROW Then
				'UPGRADE_WARNING: オブジェクト WK_DATA の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				SZ_LNCNT = Val(WK_DATA) - SZ0414_MAX_ROW - 1
			Else
				SZ_LNCNT = 0
			End If
			Call SZ0414_GET_SUB()
			If SZ_ERRSW = 1 Then
				Exit Sub
			End If
			If SZ_ENDSW = 1 Then
				SZ_ENDSW = 0
				If RDO_STATUS <> 24 Then
					Exit Sub
				Else
					JAN.k4 = ""
					JAN_BUF0.k4 = ""
					B_OP = B_GET_GT
					SZ0414SEL_SW = 0
					Call SZ0414_GET_SUB()
					If SZ_ERRSW = 1 Then
						Exit Sub
					End If
					Exit For
				End If
			End If
		Next SZ_REP
		Call SZ0414_STA_SUB(1)
		SPRD.ReDraw = True
		
	End Sub
	
	Private Sub COMBO_SETLIST_SZ0414(ByRef cBox As System.Windows.Forms.ComboBox, ByRef Txt As String)
		
		Dim lx As Integer
		For lx = 0 To cBox.Items.Count - 1
			If Trim(CStr(VB6.GetItemData(cBox, lx))) = Trim(Txt) Then
				cBox.SelectedIndex = lx
				Exit Sub
			End If
		Next lx
		cBox.SelectedIndex = -1
		
	End Sub
End Class