Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ARQENBAS
	'ƒvƒƒOƒ‰ƒ€I—¹ƒTƒuƒ‹[ƒ`ƒ“   ‚y‚`‚d‚m‚cQ‚r‚t‚a
	'||||||||||||||||||||||||
	'ZAEND_SUB  Command‚æ‚è³¨İÄŞ³ÊİÄŞÙ‚ğæ“¾‚µA
	'@@@@ @ İ’è‚É‚æ‚èÒÆ­°‚ğ±¸Ã¨ÌŞ‚É‚µ‚Ä
	'@@@      ©ÌßÛ¸Ş×Ñ‚ğEnd‚·‚é
	'||||||||||||||||||||||||
	'ÒÆ­°‚©‚çŒÄ‚Î‚ê‚Ä‚¢‚Ä(ˆø”‚ª@HWnd:DBÚ‘±•¶š—ñ)A
	'‚©‚ÂAÒÆ­°‚ª±²ºİ‰»‚³‚ê‚Ä‚¢‚é‚Æ‚«
	'mkk.ini‚ÌMDIMAX‚Ìİ’è‚É‚æ‚èÅ‘å‚©’Êí‚Ì»²½Ş‚Å
	'ÒÆ­°‚ğ±¸Ã¨ÌŞ•\¦‚·‚é
	Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
	Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Public Sub ZAEND_SUB()
		Dim Ret As Integer
		Dim MDIMAX As String
		Dim ZAEN_HWND As Integer 'ƒƒjƒ…[ƒEƒBƒ“ƒhƒEƒnƒ“ƒhƒ‹
		
		If Len(VB.Command()) <> 0 Then
			If InStr(1, VB.Command(), ":") <> 0 Then
				ZAEN_HWND = CInt(Val(Left(VB.Command(), InStr(1, VB.Command(), ":") - 1)))
				If IsWindow(ZAEN_HWND) <> 0 Then
					'ƒEƒBƒ“ƒhƒE‚ª‘¶İ‚·‚é
					If IsIconic(ZAEN_HWND) <> 0 Then
						'±²ºİ‰»‚³‚ê‚Ä‚¢‚é
						Ret = MKKCMN.ZAGI_SUB("‘€ì", "MDIMAX", "", MDIMAX, "mkk.ini")
						If Ret = True And Trim(UCase(MDIMAX)) = "TRUE" Then
							'SW_SHOWMAXIMIZED) 'Å‘å‰»•\¦
							Ret = ShowWindow(ZAEN_HWND, 3)
						Else
							'SW_SHOWNORMAL) '‚»‚Ì‚Ü‚Ü»²½Ş‚Å•\¦
							Ret = ShowWindow(ZAEN_HWND, 1)
						End If
					Else
						'±²ºİ‰»‚³‚ê‚Ä‚È‚¢
						Ret = SetForegroundWindow(ZAEN_HWND) 'MENU‚ğ‘O–Ê‚É
					End If
				End If
			End If
		End If
		End
	End Sub
End Module