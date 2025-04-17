Option Strict Off
Option Explicit On
Module ARQFCBAS
	Public ZAFC_N(12) As Short
	Public ZAFC_USE(12) As Short
	
	
	Sub ZAFC_SUB(ByRef MC As System.Windows.Forms.Form)
		Dim i As Short
		Dim cf As Short
		For i = 0 To 12
			If ZAFC_USE(i) = True Then
				If ZAFC_N(i) <> 0 Then
					CType(MC.Controls("CMDOFNC"), Object)(i).Text = Trim(ZAFC_MST(ZAFC_N(i)))
					CType(MC.Controls("CMDOFNC"), Object)(i).Enabled = True
					CType(MC.Controls("LBLFNC"), Object)(i).Enabled = True
				Else
					CType(MC.Controls("CMDOFNC"), Object)(i).Text = ""
					CType(MC.Controls("CMDOFNC"), Object)(i).Enabled = False
					CType(MC.Controls("LBLFNC"), Object)(i).Enabled = False
				End If
			End If
		Next i
		'UPGRADE_NOTE: Erase は System.Array.Clear にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		System.Array.Clear(ZAFC_N, 0, ZAFC_N.Length)
		
	End Sub
End Module