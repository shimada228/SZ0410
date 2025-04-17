Option Strict Off
Option Explicit On
Friend Class ARQCNFRM
	Inherits System.Windows.Forms.Form
	
	Private Sub CMDO010_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO010.ClickEvent
		ZACN_USERID = Trim(TXT010.Text)
		ZACN_PASSWORD = Trim(TXT020.Text)
		ZACN_DBNAME = Trim(TXT030.Text) '99/11/24 ADD FOR MKK
		ZACN_DOCNCT = True
		Me.Close()
	End Sub
	
	Private Sub CMDO010_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDO010.KeyDownEvent
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				'        TXT020.SetFocus            '99/11/24 DEL FOR MKK
				TXT030.Focus() '99/11/24 ADD FOR MKK
			Case System.Windows.Forms.Keys.Right
				CMDO020.Focus()
		End Select
	End Sub
	
	Private Sub CMDO020_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CMDO020.ClickEvent
		ZACN_DOCNCT = False
		Me.Close()
	End Sub
	
	Private Sub CMDO020_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxOskcmdLibV5.__OSKButton_KeyDownEvent) Handles CMDO020.KeyDownEvent
		Select Case eventArgs.KeyCode
			Case System.Windows.Forms.Keys.Up
				'        TXT020.SetFocus            '99/11/24 DEL FOR MKK
				TXT030.Focus() '99/11/24 ADD FOR MKK
			Case System.Windows.Forms.Keys.Left
				CMDO010.Focus()
		End Select
	End Sub
	
	
	'UPGRADE_WARNING: Form イベント ARQCNFRM.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub ARQCNFRM_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'未入力があればそこにFocusｾｯﾄ
		If ZACN_USERID = "" Then
			TXT010.Focus()
			Exit Sub
		End If
		If ZACN_PASSWORD = "" Then
			TXT020.Focus()
			Exit Sub
		End If
		If ZACN_DBNAME = "" Then '99/11/24 ADD FOR MKK
			TXT030.Focus() '99/11/24 ADD FOR MKK
			Exit Sub '99/11/24 ADD FOR MKK
		End If '99/11/24 ADD FOR MKK
		TXT010.Focus()
	End Sub
	
	Private Sub ARQCNFRM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call ZAWC_SUB(Me, 1)
		TXT010.Text = ZACN_USERID
		TXT020.Text = ZACN_PASSWORD
		TXT030.Text = ZACN_DBNAME '99/11/24 ADD FOR MKK
	End Sub
	
	
	Private Sub TXT010_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TXT010.Enter
		TXT010.SelectionStart = 0
		TXT010.SelectionLength = Len(TXT010.Text)
	End Sub
	
	Private Sub TXT010_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TXT010.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Return
				TXT020.Focus()
		End Select
	End Sub
	
	Private Sub TXT020_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TXT020.Enter
		TXT020.SelectionStart = 0
		TXT020.SelectionLength = Len(TXT020.Text)
	End Sub
	
	Private Sub TXT020_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TXT020.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Up
				TXT010.Focus()
			Case System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Return
				'        CMDO010.SetFocus                   '99/11/24 DEL FOR MKK
				TXT030.Focus() '99/11/24 ADD FOR MKK
		End Select
	End Sub
	
	Private Sub TXT030_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TXT030.Enter '99/11/24 ADD FOR MKK
		TXT030.SelectionStart = 0
		TXT030.SelectionLength = Len(TXT030.Text)
	End Sub
	
	Private Sub TXT030_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TXT030.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000 '99/11/24 ADD FOR MKK
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Up
				TXT020.Focus()
			Case System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Return
				CMDO010.Focus()
		End Select
	End Sub
End Class