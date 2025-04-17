Option Strict Off
Option Explicit On
Module ARQKBBAS
	'------------------------------------------------'
	'           電卓入力サブルーチン   For IMNumer   '
	'------------------------------------------------'
	Public ZAKB_SW As Short ' GotFocus 時に  ZAKB_SW = 0とし、有効値が入力されたら
	' 表示をクリアし ZAKB_SW = 1 とする
	
	Sub ZAKB_SUB(ByRef KeyAscii As Short)
		If ZAKB_SW = 0 Then
			' そのコントロールがフォーカスを持ってから初めてのキー入力だった
			Select Case KeyAscii
				Case 48 To 57 ' 数字０～９
					ZAKB_SW = 1
					'UPGRADE_ISSUE: Control Value は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
					System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0
				Case 45 ' マイナス符号
					'UPGRADE_ISSUE: Control MinValue は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
					If System.Windows.Forms.Form.ActiveForm.ActiveControl.MinValue < 0 Then ' 入力最小値＝負のときのみ
						ZAKB_SW = 1 '
						'UPGRADE_ISSUE: Control Value は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
						If Val(System.Windows.Forms.Form.ActiveForm.ActiveControl.Value) <> 0 Then '"0"表示のとき
							'UPGRADE_ISSUE: Control Value は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
							System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0 ' 表示をゼロにする
						End If
					End If
				Case 46 ' 小数点
					'UPGRADE_ISSUE: Control FmtDecDigits は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
					If System.Windows.Forms.Form.ActiveForm.ActiveControl.FmtDecDigits <> 0 Then ' 小数点有りのとき
						ZAKB_SW = 1
						'UPGRADE_ISSUE: Control Value は、汎用名前空間 ActiveControl 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
						System.Windows.Forms.Form.ActiveForm.ActiveControl.Value = 0
					End If
				Case System.Windows.Forms.Keys.Back ' ＢＳキー
					ZAKB_SW = 1
			End Select
		End If
	End Sub
End Module