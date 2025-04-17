Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class SZ0412FRM
	Inherits System.Windows.Forms.Form
	'A-CUST-20100610 フォーム追加 'インデックスの最小値を１に設定
	
	Dim CUR_NO As Short '現入力位置ｺﾝﾄﾛｰﾙ№
	Dim LST_NO As Short '前入力位置ｺﾝﾄﾛｰﾙ№
	Dim NXT_NO As Short '次入力位置ｺﾝﾄﾛｰﾙ№
	
	'明細(スプレッド)用情報
	Dim PRE_VALUE As Object '明細前値
	Dim ViewCol As Integer '明細現列
	Dim SPRD_ERR As Short 'スプレッド入力ＯＫフラグ
	
	'------------------------------------------------------------------
	'         画面ｺﾝﾄﾛｰﾙ項目設定
	'------------------------------------------------------------------
	Const N005 As Short = 1 '明細
	Const NEND As Short = 2 '
	'UPGRADE_WARNING: 配列 CTRLTBL の下限が 1 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Dim CTRLTBL(NEND) As CTRLTBL_S '画面ｺﾝﾄﾛｰﾙ配列
	
	Const GRP1 As Short = 1
	Const GEND As Short = 2
	'UPGRADE_WARNING: 配列 GRPTBL の下限が 0 から 0 に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' をクリックしてください。
	Dim GRPTBL(GEND) As GRPTBL_S '画面ｸﾞﾙｰﾌﾟ配列
	
	Enum SPRD_COL
		col_SEN = 1
		col_HIN_NAME
		col_KIKAKU
		col_GYO_NAME
		col_TANI
		col_TANKA
		'A-CUST-20100823 Start
		col_TEKI_DATE
		col_HA_TANI
		col_KANSANSU
		col_JAN_CODE
		col_JAN_S_CODE
		col_BAR_CODE
		'A-CUST-20100823 End
		col_RENBAN '非表示
	End Enum
	
	Private Sentakurow As Integer
	Private ButtnFlg As Boolean
	
	Private Function ALLCHK_UPD() As Boolean
		Dim ii As Integer
		Dim SvRow As Integer
		
		SvRow = SPRD050.ActiveRow
		With SPRD050
			
			Sentakurow = 0
			'全未承認のチェック
			For ii = 1 To .MaxRows
				.ROW = ii
				.Col = 1
				
				If CBool(.Value) = True Then
					If Sentakurow = 0 Then
						Sentakurow = ii
					Else
						Sentakurow = 0
						Exit For
					End If
				End If
			Next ii
			
			If Sentakurow = 0 Then
				ALLCHK_UPD = False
				Exit Function
			End If
			
			SPRD050.ROW = SvRow
			
		End With
		
		ALLCHK_UPD = True
		
	End Function
	
	Public Sub INITIAL_RTN()
		'画面項目初期値設定
		DSP010.Text = RTrim(WKB010)
		DSP020.Text = RTrim(SZ0410FRM.DSP010.Text)
		DSP030.Text = RTrim(WKB020)
		DSP040.Text = RTrim(SZ0410FRM.DSP020.Text)
		
		SPRD050.MaxRows = 0
		
	End Sub
	
	
	'******************************************************************
	'*      画面ｺﾝﾄﾛｰﾙ初期設定                                (TBL_SET)
	'******************************************************************
	Sub TBL_SET()
		CTRLTBL(N005).IGRP = GRP1 '明細グループ
		CTRLTBL(NEND).IGRP = GEND
		
SET_NO: 
		'------------------------------------------------------------
		'   次項目、前項目の設定
		'------------------------------------------------------------
		
		CTRLTBL(N005).INEXT = NEND '明細
		CTRLTBL(N005).IBACK = n0
		CTRLTBL(N005).IDOWN = NEND
		
		CTRLTBL(NEND).INEXT = n0
		CTRLTBL(NEND).IBACK = N005
		CTRLTBL(NEND).IDOWN = n0
		
		'------------------------------------------------------------
		'   ｺﾝﾄﾛｰﾙ保存
		'------------------------------------------------------------
		CTRLTBL(N005).CTRL = SPRD050 '明細
		CTRLTBL(NEND).CTRL = CMDOFNC(12) '実行ボタン
		
	End Sub
	
	'******************************************************************
	'*      ファンクション・ボタン（GotFocus）
	'******************************************************************
	Private Sub CMDOFNC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.Enter
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		If MOUSEFLG = 0 Then MOUSEFLG = VB6.MouseButtonConstants.LeftButton
		
		If Index <> 12 Then Exit Sub
		
		If CUR_NO = NEND Then Exit Sub '  自分なら何もしない
		CUR_NO = NEND '  仮の現在の位置を設定する
		
		If LST_NO <> n0 Then '【チェック】
			If IPROCHK() = False Then '  LostFocus項目ﾁｪｯｸ
				Exit Sub
			End If
			If GPROCHK() = False Then '  LostFocus項目群ﾁｪｯｸ
			End If
		End If
		If GVALCHK() = False Then '  GotFocus項目群ﾁｪｯｸ
			Exit Sub
		End If
		If MVALCHK() = False Then '  GotFocus項目ﾁｪｯｸ
			Exit Sub
		End If
		
		LST_NO = CUR_NO '【画面ｺﾝﾄﾛｰﾙ№確定】
		
		Call FUNCSET_RTN() '【ファンクションガイド表示】
		
	End Sub
	
	Private Sub CMDOFNC_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDOFNC.KeyDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		
		'実行以外は無視
		If Index < 12 Then
			Exit Sub
		End If
		
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				
				NXT_NO = N005
				Call FOCUS_SET()
				Exit Sub
				
		End Select
		
		Call SZ0412FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
		
	End Sub
	
	'******************************************************************
	'*      ファンクション・ボタン（MouseDown）
	'******************************************************************
	Private Sub CMDOFNC_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_MouseDownEvent) Handles CMDOFNC.MouseDownEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		MOUSEFLG = eventArgs.Button
	End Sub
	
	'******************************************************************
	'*      ファンクション・ボタン（Click）
	'******************************************************************
	Private Sub CMDOFNC_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDOFNC.ClickEvent
		Dim Index As Short = CMDOFNC.GetIndex(eventSender)
		Dim i As Integer
		
		If MOUSEFLG <> VB6.MouseButtonConstants.LeftButton And MOUSEFLG <> 0 Then
			MOUSEFLG = 0
			Exit Sub
		End If
		
		MOUSEFLG = 0
		
		If CMDOFNC(Index).Text = "" Then
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Sub
		End If
		
		Dim IROWSelect As Short
		Select Case Index
			Case 0 '【終了】
				'UPGRADE_NOTE: EditMode は CtlEditMode にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
				SPRD050.CtlEditMode = False
				IMTXDUM.Focus()
				
			Case 5 '【クリア】
				With SPRD050
					If .MaxRows > 0 Then
						For IROWSelect = 1 To .MaxRows
							.ROW = IROWSelect
							.Col = 1
							.Value = CStr(False)
						Next 
					End If
					
					.Focus()
				End With
				
				'A-20110621-S
			Case 9 '【削除】
				If ALLCHK_UPD() = False Then
					ZAER_KN = n0
					ZAER_CD = 120 '"入力内容に誤りがあります・・・"
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					DUMMY.Focus()
					Exit Sub
				End If
				
				If MsgBox("削除を行います。よろしいですか？", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then
					NXT_NO = N005
					Call FOCUS_SET()
					Exit Sub
				End If
				
				Call BEGIN_RTN()
				SPRD050.ROW = Sentakurow
				SPRD050.Col = SPRD_COL.col_RENBAN
				RENBAN_SEN = CInt(SPRD050.Value)
				Call GO_WKDELETE() '削除
				If ERRSW = F_ERR Then
					Call ROLLBACK_RTN()
				Else
					Call COMMIT_RTN()
				End If
				
				ButtnFlg = False
				Call SET_SPRD050_UPD() '再表示
				ButtnFlg = True
				'A-20110621-E
				
			Case 12 '【実行】
				If ALLCHK_UPD() = False Then
					ZAER_KN = n0
					ZAER_CD = 120 '"入力内容に誤りがあります・・・"
					ZAER_NO.Value = ""
					Call ZAER_SUB()
					DUMMY.Focus()
					Exit Sub
				End If
				
				'チェックＯＫ
				If MsgBox("品目データの取込みを行います。よろしいですか？", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, Me.Text) = MsgBoxResult.No Then
					NXT_NO = N005
					Call FOCUS_SET()
					Exit Sub
				End If
				
				SPRD050.ROW = Sentakurow
				SPRD050.Col = SPRD_COL.col_HIN_NAME
				KB.hin_name_seisiki = SPRD050.Value
				'UPGRADE_WARNING: オブジェクト MKKCMN.ZACHGSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				KB.hin_name = MKKCMN.ZACHGSTR_SUB(KB.hin_name_seisiki, Len(KB.hin_name))
				SPRD050.Col = SPRD_COL.col_KIKAKU
				KB.kikaku = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_TANI
				KB.tani = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_TANKA
				KB.kei_kin1 = CDec(SPRD050.Value)
				'A-CUST-20100823 Start
				SPRD050.Col = SPRD_COL.col_TEKI_DATE
				If RTrim(SPRD050.Value) = "" Then
					KB.teki_date1 = ""
				Else
					KB.teki_date1 = VB.Left(SPRD050.Value, 4) & Mid(SPRD050.Value, 6, 2) & Mid(SPRD050.Value, 9, 2)
				End If
				SPRD050.Col = SPRD_COL.col_HA_TANI
				KB.ha_tanka1 = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_KANSANSU
				KB.kansan_num1 = CDec(SPRD050.Value)
				SPRD050.Col = SPRD_COL.col_JAN_CODE
				KB.jan_code = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_JAN_S_CODE
				KB.jan_s_code = SPRD050.Value
				SPRD050.Col = SPRD_COL.col_BAR_CODE
				KB.bar_code = SPRD050.Value
				'A-CUST-20100823 End
				
				SPRD050.Col = SPRD_COL.col_RENBAN
				RENBAN_SEN = CInt(SPRD050.Value)
				
				SentakuFLG = True
				
				Call SZ0410FRM.DSP_SENTAKU()
				
				IMTXDUM.Focus()
				
		End Select
	End Sub
	
	'UPGRADE_ISSUE: PictureBox イベント DUMMY.GotFocus はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' をクリックしてください。
	Private Sub DUMMY_GotFocus()
		SPRD050.Focus()
		
	End Sub
	
	'UPGRADE_WARNING: Form イベント SZ0412FRM.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub SZ0412FRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'マウスカーソルを戻す
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
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
		lRet = MKKDBCMN.MKKDBCMN_SQTGET1_SUB(3, "SZ0412", WKB010, WG_OPCODE, OP_KENGEN) 'A-CUST-20100901
		If lRet <> n0 Then
			IMTXDUM.Focus()
			Exit Sub
		End If
		If OP_KENGEN = 0 Then
			ZAER_KN = n0
			ZAER_CD = 301
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			IMTXDUM.Focus()
		End If
		
		'    '更新権限チェック
		'    '更新権限
		'    lRet = MKKDBCMN.MKKDBCMN_SQTGET3_SUB(4, "SZ0410", WKB010, WKB020, "", WG_OPCODE, OP_KENGEN)
		'    If lRet <> n0 Then
		'        IMTXDUM.SetFocus
		'        Exit Sub
		'    End If
		'    '更新権限なし
		'    If OP_KENGEN = 0 Then
		'        ZAER_KN = 0
		'        ZAER_CD = 303
		'        ZAER_NO = ""
		'        Call ZAER_SUB
		'        IMTXDUM.SetFocus
		'        Exit Sub
		'    End If
		
		ButtnFlg = False
		Call SET_SPRD050_UPD()
		ButtnFlg = True
		
		CUR_NO = n0
		LST_NO = N005
		NXT_NO = N005
		Call FOCUS_SET()
		
	End Sub
	
	'******************************************************************
	'*      ＦＯＲＭ（ＫｅｙＤｏｗｎ）
	'******************************************************************
	Private Sub SZ0412FRM_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'各コントロールの共通のキー制御を行う
		'固有のキー制御は各コントロールのKeyDownイベントで行う
		
		If Me.Enabled = False Then
			KeyCode = n0
			Exit Sub
		End If
		If Shift <> n0 Then 'Shift,Ctrl,Graph(Alt)キー押下時、処理無効
			Exit Sub
		End If
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape '【ESC】
				KeyCode = n0
				If CMDOFNC(0).Text <> "" Then
					CMDOFNC(0).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(0), New System.EventArgs())
				End If
			Case System.Windows.Forms.Keys.Return '【ﾘﾀｰﾝｷｰ】
				Call SET_NO(1)
				KeyCode = n0
			Case System.Windows.Forms.Keys.Up '【↑ｷｰ】
				Call SET_NO(2)
				KeyCode = n0
			Case System.Windows.Forms.Keys.Down '【↓ｷｰ】
				Call SET_NO(3)
				KeyCode = n0
			Case System.Windows.Forms.Keys.F1 '【Ｆ１】
			Case System.Windows.Forms.Keys.F2 '【Ｆ２】
				If CMDOFNC(2).Text <> "" Then
					CMDOFNC(2).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(2), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F3 '【Ｆ３】
				If CMDOFNC(3).Text <> "" Then
					CMDOFNC(3).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(3), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F4 '【Ｆ４】
				If CMDOFNC(4).Text <> "" Then
					CMDOFNC(4).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(4), New System.EventArgs())
					CTRLTBL(CUR_NO).CTRL.Focus()
				End If
				'KeyCode = n0
			Case System.Windows.Forms.Keys.F5 '【Ｆ５】
				If CMDOFNC(5).Text <> "" Then
					CMDOFNC(5).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(5), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F6 '【Ｆ６】
				If CMDOFNC(6).Text <> "" Then
					CMDOFNC(6).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(6), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F7 '【Ｆ７】
				If CMDOFNC(7).Text <> "" Then
					CMDOFNC(7).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(7), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F8 '【Ｆ８】
				If CMDOFNC(8).Text <> "" Then
					CMDOFNC(8).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(8), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F9 '【Ｆ９】
				If CMDOFNC(9).Text <> "" Then
					CMDOFNC(9).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(9), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F10 '【Ｆ10】
				KeyCode = n0
			Case System.Windows.Forms.Keys.F11 '【Ｆ11】
				If CMDOFNC(11).Text <> "" Then
					CMDOFNC(11).Focus()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(11), New System.EventArgs())
				End If
				KeyCode = n0
			Case System.Windows.Forms.Keys.F12 '【Ｆ12】
				KeyCode = n0
				If CMDOFNC(12).Text <> "" Then
					CMDOFNC(12).Focus()
					System.Windows.Forms.Application.DoEvents()
					Call CMDOFNC_ClickEvent(CMDOFNC.Item(12), New System.EventArgs())
				End If
		End Select
		
	End Sub
	
	'******************************************************************
	'*      ＦＯＲＭ（ＬＯＡＤ）
	'******************************************************************
	Private Sub SZ0412FRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: Form プロパティ SZ0412FRM.HelpContextID はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
		Me.HelpContextID = SM_HelpContextID
		
		Call TBL_SET() '画面ｺﾝﾄﾛｰﾙ初期設定
		
	End Sub
	
	'******************************************************************
	'*      ＦＯＲＭ（ＱｕｅｒｙＵｎｌｏａｄ）
	'******************************************************************
	Private Sub SZ0412FRM_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'******************************************************************
	'*      フォーカスセット                                (FOCUS_SET)
	'******************************************************************
	Private Sub FOCUS_SET()
		
		If NXT_NO <= 0 Then Exit Sub
		
		Select Case NXT_NO
			Case N005
				SPRD050.Focus()
				SPRD050.Col = 1
				SPRD050.Action = 0
				
			Case NEND
				CType(Me.Controls("CMDOFNC"), Object)(12).Enabled = True
				CType(Me.Controls("LBLFNC"), Object)(12).Enabled = True
				CMDOFNC(12).Focus()
				
			Case Else
				CTRLTBL(NXT_NO).CTRL.Focus()
				
		End Select
		
	End Sub
	
	'******************************************************************
	'*      ＦＵＮＣＴＩＯＮセット                            (FUNCSET)
	'******************************************************************
	Private Sub FUNCSET_RTN()
		
		'--- ファンクション・ガイドメッセージ
		Select Case LST_NO
			
			Case N005 '明細
				With SPRD050
					Select Case ViewCol
						Case 1
							ZAFC_N(0) = 1
							ZAFC_N(5) = 5
							ZAFC_N(9) = 8 'A-20110621-
							ZAFC_N(12) = 12
					End Select
				End With
				
			Case NEND
				ZAFC_N(12) = 12
				
		End Select
		
		'ファンクションメッセージ
		Call ZAFC_SUB(Me)
		
		'ガイドメッセージ
		Call ZAGD_SUB(Me)
	End Sub
	
	'******************************************************************
	'*      画面ｺﾝﾄﾛｰﾙ№セット                                 (SET_NO)
	'******************************************************************
	Sub SET_NO(ByRef FUNC As Short)
		
		Select Case FUNC
			Case 1 ' 次項目
				
				NXT_NO = CTRLTBL(LST_NO).INEXT
				Call FOCUS_SET()
			Case 2 ' 前項目
				If CTRLTBL(LST_NO).IBACK <> 0 Then
					NXT_NO = CTRLTBL(LST_NO).IBACK
					Call FOCUS_SET()
				End If
			Case 3 ' 次グループ
				NXT_NO = CTRLTBL(LST_NO).INEXT
				Call FOCUS_SET()
		End Select
		
	End Sub
	
	'******************************************************************
	'*      グループチェック（LostFocus項目群チェック）       (GPROCHK)
	'******************************************************************
	Function GPROCHK() As Short
		GPROCHK = True
		
		ERRSW = F_OFF
		ENDSW = F_OFF
		
		If CTRLTBL(CUR_NO).IGRP <= CTRLTBL(LST_NO).IGRP Then Exit Function
		
		Select Case CTRLTBL(LST_NO).IGRP
			Case GRP1
				Call GPROCHK_GRP1()
		End Select
		
		
		If ERRSW = F_ERR Then 'エラー
			GPROCHK = False
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = False
			Select Case CTRLTBL(LST_NO).IGRP
				Case GRP1
					NXT_NO = GRPTBL(GRP1).NXTN
			End Select
			Call FOCUS_SET()
		Else
			GRPTBL(CTRLTBL(LST_NO).IGRP).CFLG = True
		End If
		
	End Function
	
	'******************************************************************
	'*      グループチェック（１）                       (GPROCHK_GRP1)
	'******************************************************************
	Sub GPROCHK_GRP1()
		
		GRPTBL(GRP1).CFLG = True
		
	End Sub
	
	'******************************************************************
	'*      入力可否チェック（グループ）                      (GVALCHK)
	'******************************************************************
	Function GVALCHK() As Short
		GVALCHK = True
		ERRSW = F_OFF
		
		If LST_NO <> n0 Then
			If CTRLTBL(CUR_NO).IGRP = CTRLTBL(LST_NO).IGRP Then Exit Function
		End If
		
		Select Case CTRLTBL(CUR_NO).IGRP
			Case GRP1
				Call GVALCHK_GRP2()
		End Select
		If ERRSW = F_ERR Then
			GVALCHK = False
			Call FOCUS_SET()
		End If
		
	End Function
	
	'******************************************************************
	'*      入力可否チェック（グループ）②               (GVALCHK_GRP2)
	'******************************************************************
	Sub GVALCHK_GRP2()
	End Sub
	
	'******************************************************************
	'*      入力内容チェック（LoasFocus項目のﾁｪｯｸ）           (IPROCHK)
	'******************************************************************
	Function IPROCHK() As Short
		Dim i As Short
		
		IPROCHK = True
		ERRSW = F_OFF '　ｴﾗｰｸﾘｱ
		ENDSW = F_OFF
		
		If CUR_NO = LST_NO Then Exit Function '　項目間の移動がない場合は何もしない
		
		Select Case LST_NO '　移動前項目のチェック
		End Select
		
		If ENDSW = F_END Then
			IPROCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
			Exit Function
		End If
		
		'エラー時
		If ERRSW = F_ERR Then
			If CUR_NO < LST_NO Then
				ERRSW = F_OFF
				'逆方向のときは直前項目値の再表示
				Select Case LST_NO
				End Select
			Else
				IPROCHK = False
				NXT_NO = LST_NO
				Call FOCUS_SET()
			End If
		End If
		
	End Function
	
	'******************************************************************
	'*      スプレッドセット処理  (SET_SPRD050_UPD)
	'******************************************************************
	Private Function SET_SPRD050_UPD() As Short
		Dim i As Integer
		Dim FLG As Boolean
		Dim wROW As Integer
		
		SET_SPRD050_UPD = 0
		SPRD050.Col = SPRD_COL.col_RENBAN 'A-CUST-20100823
		SPRD050.ColHidden = True 'A-CUST-20100823
		
		i = 0
		
		WSZ0410SEL01.rdoParameters("Inc_code").Value = WKB010 '会社コード
		WSZ0410SEL01.rdoParameters("jg_code").Value = WKB020 '事業所コード
		
		On Error Resume Next
		WSZ0410RS = WSZ0410SEL01.OpenResultset()
		Select Case B_STATUS(WSZ0410RS)
			Case 0
				SPRD050.ReDraw = False
				Do 
					i = i + 1
					With SPRD050
						.MaxRows = i
						.ROW = i
						
						.Col = SPRD_COL.col_SEN
						If SentakuFLG Then
							If WSZ0410RS.rdoColumns("y_code").Value = RENBAN_SEN Then
								.Value = CStr(True)
								FLG = True
								wROW = i
							Else
								.Value = CStr(False)
							End If
						Else
							.Value = CStr(False)
						End If
						
						.Col = SPRD_COL.col_HIN_NAME
						.Value = Trim(WSZ0410RS.rdoColumns("hin_name_seisiki").Value)
						
						.Col = SPRD_COL.col_KIKAKU
						.Value = Trim(WSZ0410RS.rdoColumns("kikaku").Value)
						
						.Col = SPRD_COL.col_GYO_NAME
						.Value = Trim(WSZ0410RS.rdoColumns("gyo_name").Value)
						
						.Col = SPRD_COL.col_TANI
						.Value = Trim(WSZ0410RS.rdoColumns("tani").Value)
						
						.Col = SPRD_COL.col_TANKA
						.Value = Trim(WSZ0410RS.rdoColumns("tanka").Value)
						
						.Col = SPRD_COL.col_RENBAN
						.Value = Trim(WSZ0410RS.rdoColumns("y_code").Value)
						
						'A-CUST-20100823 Start
						.Col = SPRD_COL.col_TEKI_DATE
						If Trim(WSZ0410RS.rdoColumns("teki_date").Value) = "" Then
							.Value = ""
						Else
							.Value = VB6.Format(WSZ0410RS.rdoColumns("teki_date").Value, "@@@@/@@/@@")
						End If
						
						.Col = SPRD_COL.col_HA_TANI
						.Value = Trim(WSZ0410RS.rdoColumns("ha_tani").Value)
						
						.Col = SPRD_COL.col_KANSANSU
						.Value = Trim(WSZ0410RS.rdoColumns("kansansu").Value)
						
						.Col = SPRD_COL.col_JAN_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("jan_code").Value)
						
						.Col = SPRD_COL.col_JAN_S_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("jan_s_code").Value)
						
						.Col = SPRD_COL.col_BAR_CODE
						.Value = Trim(WSZ0410RS.rdoColumns("bar_code").Value)
						'A-CUST-20100823 End
						SPRD050.set_RowHeight(i, 12)
						
						WSZ0410RS.MoveNext()
					End With
				Loop Until WSZ0410RS.EOF = True
				
				SentakuFLG = FLG
				SPRD050.Col = 1
				If FLG Then
					SPRD050.ROW = wROW
				Else
					SPRD050.ROW = 1
				End If
				SPRD050.Action = SS_ACTION_ACTIVE_CELL
				
				SPRD050.ReDraw = True
				
			Case 24
				SentakuFLG = FLG
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(RENBAN_SEN, "000000")
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
				On Error GoTo 0
				Exit Function
		End Select
		
	End Function
	
	'******************************************************************
	'*      項目入力可否チェック                              (MVALCHK)
	'******************************************************************
	Function MVALCHK() As Short
		MVALCHK = True
		ERRSW = F_OFF
		
		Select Case CUR_NO '　移動可否のチェック
		End Select
		If ERRSW = F_ERR Then
			MVALCHK = False
			NXT_NO = LST_NO
			Call FOCUS_SET()
		End If
	End Function
	
	'******************************************************************
	'*      項目入力可否チェック④                       (MVALCHK_N007)
	'******************************************************************
	Sub MVALCHK_N007()
		
	End Sub
	
	'アクティブセルが移動した場合の処理：
	'    アクティブ列用変数の記憶
	'    セル元値の記憶
	'    ファンクション設定
	'    Ｏｐ情報表示
	Private Sub MyProcOfCell(ByVal Col As Integer, ByVal ROW As Integer, ByVal NewCol As Integer, ByVal NewRow As Integer)
		
		With SPRD050
			'
			ViewCol = NewCol '明細現列
			Call FUNCSET_RTN() 'ファンクション表示
			
			.ROW = NewRow
			.Col = NewCol
			
			
			'入力OKフラグ(次項目への移動許可)の可否決定および、
			'名称情報の記憶
			Select Case .Col
				Case 1
					'元値 記憶
					If SPRD_ERR <> 1 Then
						'UPGRADE_WARNING: オブジェクト PRE_VALUE の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
						PRE_VALUE = SPRD050.Value
					End If
			End Select
			
			'Colを戻す
			.Col = NewCol
			
		End With
		
	End Sub
	
	Private Sub IMTXDUM_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IMTXDUM.Enter
		Me.Close()
		
	End Sub
	
	'******************************************************************
	'*      明細（GotFocus）
	'******************************************************************
	Private Sub SPRD050_ButtonClicked(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SPRD050.ButtonClicked
		Dim IROWSelect As Short
		If ButtnFlg Then
			If eventArgs.ButtonDown <> 0 Then
				
				With SPRD050
					If .MaxRows > 0 Then
						For IROWSelect = 1 To .MaxRows
							If IROWSelect <> eventArgs.ROW Then
								.eventArgs.ROW = IROWSelect
								.eventArgs.Col = 1
								.Value = CStr(False)
							End If
						Next 
					End If
				End With
			End If
		End If
		
	End Sub
	
	'******************************************************************
	'*      明細（GotFocus）
	'******************************************************************
	Private Sub SPRD050_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD050.Enter
		
		If CUR_NO = N005 Then Exit Sub
		CUR_NO = N005
		
		'チェック開始
		If LST_NO <> n0 Then
			'LostFocus項目のチェック
			If IPROCHK() = False Then
				Exit Sub
			End If
			'LostFocus項目群のチェック
			If GPROCHK() = False Then
			End If
		End If
		'GotFocus項目群のチェック
		If GVALCHK() = False Then
			Exit Sub
		End If
		'GotFocus項目のチェック
		If MVALCHK() = False Then
			Exit Sub
		End If
		
		LST_NO = CUR_NO
		
		'セル移動時処理
		With SPRD050
			.Col = .ActiveCol
			.ROW = .ActiveRow
			Call MyProcOfCell(.Col, .ROW, .Col, .ROW)
		End With
		
		SPRD_ERR = 0
		
	End Sub
	
	'******************************************************************
	'*      明細（KeyDown）
	'******************************************************************
	Private Sub SPRD050_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SPRD050.KeyDownEvent
		
		With SPRD050
			'ROW COL 設定
			.ROW = .ActiveRow
			.Col = .ActiveCol
			
			Select Case eventArgs.KeyCode
				
				Case System.Windows.Forms.Keys.Return 'Enterｷｰ
					
					Select Case .ActiveCol
						Case 1
							eventArgs.KeyCode = 0
							'次行１列目に移動
							If (.ROW < .MaxRows) Then
							Else
								Call SET_NO(1)
								Exit Sub
							End If
							
					End Select
					
				Case System.Windows.Forms.Keys.Up
					
				Case System.Windows.Forms.Keys.Down
					eventArgs.KeyCode = 0
					'入力項目間の移動
					Select Case .ActiveCol
						Case 1
							If .ROW < .MaxRows Then
								'最終行より上の行の場合：
								'入力可能項目だった場合：
								.Col = 1
								.ROW = .ROW + 1
								.Action = SS_ACTION_ACTIVE_CELL
								Call MyProcOfCell(.Col, .ROW, .Col, .ROW)
							Else
								Call SET_NO(1)
								Exit Sub
							End If
					End Select
					
				Case Else
					Call SZ0412FRM_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(eventArgs.KeyCode Or eventArgs.Shift * &H10000))
					
			End Select
		End With
		
	End Sub
	
	Private Sub SPRD050_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SPRD050.LeaveCell
		
		If (eventArgs.NewRow < 0) Or (eventArgs.NewCol < 0) Then Exit Sub
		
		With SPRD050
			'ｸﾘｯｸで表示項目に移動する事の防止
			If (eventArgs.NewCol > 1) Then
				.eventArgs.Col = eventArgs.Col
				.eventArgs.ROW = eventArgs.ROW
				eventArgs.Cancel = True
				Exit Sub
			End If
			
			'マウスクリックによる他セルへの移動の場合：
			'    右・下への移動：入力値ＯＫの場合のみ許す
			'    左・上への移動：入力値ＮＧの場合、元の値に戻す
			.eventArgs.Col = eventArgs.Col
			.eventArgs.ROW = eventArgs.ROW
			Select Case eventArgs.Col
				
			End Select
			
			'移動先の行において、移動先の列よりも左の列が未入力だったら、その列に移動
			.eventArgs.ROW = eventArgs.NewRow
			
		End With
		
LC_EXIT: 
		'セル移動時処理
		Call MyProcOfCell(eventArgs.Col, eventArgs.ROW, eventArgs.NewCol, eventArgs.NewRow)
		
	End Sub
	
	Private Sub SPRD050_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SPRD050.Leave
		
		If (SPRD050.MaxRows < 1) Then Exit Sub
		If (SPRD050.ActiveRow < 0) Then Exit Sub
		If (SPRD050.ActiveCol < 0) Then Exit Sub
		
		With SPRD050
			.ROW = .ActiveRow
			.Col = .ActiveCol
			SPRD_ERR = 0
			
			Select Case .ActiveCol
				Case 1
					'UPGRADE_ISSUE: Control NAME は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
					If (Me.ActiveControl.Name = "CMDOFNC") Then
						'If Me.ActiveControl.Index <> 12 Then       'D-20110621-
						'UPGRADE_ISSUE: Control Index は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
						If Me.ActiveControl.Index = 12 Or Me.ActiveControl.Index = 9 Then 'A-20110621-
						Else 'A-20110621-
							'UPGRADE_WARNING: オブジェクト PRE_VALUE の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							.Value = PRE_VALUE
							Exit Sub
						End If
					End If
					
			End Select
		End With
		
	End Sub
End Class