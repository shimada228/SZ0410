Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ARQENBAS
	'プログラム終了サブルーチン   ＺＡＥＮＤ＿ＳＵＢ
	'−−−−−−−−−−−−−−−−−−−−−−−−
	'ZAEND_SUB  Commandよりｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙを取得し、
	'　　　　 　 設定によりﾒﾆｭｰをｱｸﾃｨﾌﾞにして
	'　　　      自ﾌﾟﾛｸﾞﾗﾑをEndする
	'−−−−−−−−−−−−−−−−−−−−−−−−
	'ﾒﾆｭｰから呼ばれていて(引数が　HWnd:DB接続文字列)、
	'かつ、ﾒﾆｭｰがｱｲｺﾝ化されているとき
	'mkk.iniのMDIMAXの設定により最大か通常のｻｲｽﾞで
	'ﾒﾆｭｰをｱｸﾃｨﾌﾞ表示する
	Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
	Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Public Sub ZAEND_SUB()
		Dim Ret As Integer
		Dim MDIMAX As String
		Dim ZAEN_HWND As Integer 'メニューウィンドウハンドル
		
		If Len(VB.Command()) <> 0 Then
			If InStr(1, VB.Command(), ":") <> 0 Then
				ZAEN_HWND = CInt(Val(Left(VB.Command(), InStr(1, VB.Command(), ":") - 1)))
				If IsWindow(ZAEN_HWND) <> 0 Then
					'ウィンドウが存在する
					If IsIconic(ZAEN_HWND) <> 0 Then
						'ｱｲｺﾝ化されている
						Ret = MKKCMN.ZAGI_SUB("操作", "MDIMAX", "", MDIMAX, "mkk.ini")
						If Ret = True And Trim(UCase(MDIMAX)) = "TRUE" Then
							'SW_SHOWMAXIMIZED) '最大化表示
							Ret = ShowWindow(ZAEN_HWND, 3)
						Else
							'SW_SHOWNORMAL) 'そのままｻｲｽﾞで表示
							Ret = ShowWindow(ZAEN_HWND, 1)
						End If
					Else
						'ｱｲｺﾝ化されてない
						Ret = SetForegroundWindow(ZAEN_HWND) 'MENUを前面に
					End If
				End If
			End If
		End If
		End
	End Sub
End Module