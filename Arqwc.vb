Option Strict Off
Option Explicit On
Module ARQWCBAS
	
	'
	'  +------------------------------------------------------+
	' ｜  ＜ウィンドウ位置サイズ変更サブルーチン＞ 　 　      ｜
	'  +------------------------------------------------------+
	'引数　MAXIMIZE_SW　0:最大化しない　ｾﾝﾀﾘﾝｸﾞしない
	'                   1:最大化しない　ｾﾝﾀﾘﾝｸﾞする
	'                   2:最大表示する
	'
	Sub ZAWC_SUB(ByRef MC As System.Windows.Forms.Form, ByRef MAXIMIZE_SW As Short)
		Dim GETWORK As String
		On Error Resume Next
		If MKKCMN.ZAGI_SUB("操作", "ICONFNAME", "", GETWORK, "mkk.ini") = True Then
			'アイコンのロード
			MC.Icon = New System.Drawing.Icon(Trim(GETWORK))
		End If
		On Error GoTo ERRSIZE
		
		Select Case MAXIMIZE_SW
			Case 0
			Case 1
				'フォームを中央に表示
				If MC.WindowState = System.Windows.Forms.FormWindowState.Normal Then
					MC.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MC.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MC.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				End If
			Case 2
				If MKKCMN.ZAGI_SUB("操作", "MDIMAX", "False", GETWORK, "mkk.ini") = True Then
					If UCase(Trim(GETWORK)) = "TRUE" Then
						MC.WindowState = System.Windows.Forms.FormWindowState.Maximized
					End If
				End If
				If MC.WindowState <> System.Windows.Forms.FormWindowState.Maximized Then
					MC.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MC.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MC.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				End If
		End Select
		Exit Sub
		
ERRSIZE: 
		If Err.Number = 387 Or Err.Number = 422 Or Err.Number = 383 Then Resume Next
	End Sub
End Module