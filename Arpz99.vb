Option Strict Off
Option Explicit On
Module ARPZ99BAS
	' エラーメッセージファイル　　ＲＡＺ９９ＥＲＲＭ
	' ファイルレイアウト
	Structure AZ99_S
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public S001() As Char
		'    S002 As String * 3
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(4),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=4)> Public S002() As Char '99/12/09 ADD KTT YOSHINO
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(48),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=48)> Public S003() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(48),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=48)> Public S004() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(48),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=48)> Public S005() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(2),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=2)> Public S006() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public S007() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public S008() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public S009() As Char
	End Structure
	
	Structure AZ99_KEY0_S
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public S001() As Char
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(3),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=3)> Public S002() As Char
	End Structure
	
	' データバッファ
	Public AZ99 As AZ99_S
	' キーナンバー
	Public AZ99_KEYNO As Short
	' キーバッファ
	Public AZ99_BUF0 As AZ99_KEY0_S
	
	Public AZ99OPENSW As Short
	Public AZ99INVSW As Short
	
	'RDO関連オブジェクト
	Public AZ99RS As RDO.rdoResultset
	Public AZ99INS As RDO.rdoQuery
	Public AZ99UPD As RDO.rdoQuery
	Public AZ99DEL As RDO.rdoQuery
	Public AZ99RSSW As String
	
	Sub AZ99CNV_SUB(ByRef ACT As Object)
		
		'UPGRADE_WARNING: オブジェクト ACT の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		If ACT = "GET" Then
			AZ99.S001 = VB6.Format(AZ99RS.rdoColumns("AZ99001").Value, "0")
			AZ99.S002 = AZ99RS.rdoColumns("AZ99002").Value
			AZ99.S003 = AZ99RS.rdoColumns("AZ99003").Value
			AZ99.S004 = AZ99RS.rdoColumns("AZ99004").Value
			AZ99.S005 = AZ99RS.rdoColumns("AZ99005").Value
			AZ99.S006 = VB6.Format(AZ99RS.rdoColumns("AZ99006").Value, "00")
			AZ99.S007 = VB6.Format(AZ99RS.rdoColumns("AZ99007").Value, "0")
			AZ99.S008 = VB6.Format(AZ99RS.rdoColumns("AZ99008").Value, "0")
			AZ99.S009 = AZ99RS.rdoColumns("AZ99009").Value
			'UPGRADE_WARNING: オブジェクト ACT の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		ElseIf ACT = "SET" Then 
			AZ99RS.rdoColumns("AZ99001").Value = Val(AZ99.S001)
			AZ99RS.rdoColumns("AZ99002").Value = AZ99.S002
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99RS.rdoColumns("AZ99003").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S003)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99RS.rdoColumns("AZ99004").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S004)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99RS.rdoColumns("AZ99005").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S005)
			AZ99RS.rdoColumns("AZ99006").Value = Val(AZ99.S006)
			AZ99RS.rdoColumns("AZ99007").Value = Val(AZ99.S007)
			AZ99RS.rdoColumns("AZ99008").Value = Val(AZ99.S008)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99RS.rdoColumns("AZ99009").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S009)
			'UPGRADE_WARNING: オブジェクト ACT の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		ElseIf ACT = "INS" Then 
			AZ99INS.rdoParameters("AZ99001").Value = Val(AZ99.S001)
			AZ99INS.rdoParameters("AZ99002").Value = AZ99.S002
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99INS.rdoParameters("AZ99003").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S003)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99INS.rdoParameters("AZ99004").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S004)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99INS.rdoParameters("AZ99005").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S005)
			AZ99INS.rdoParameters("AZ99006").Value = Val(AZ99.S006)
			AZ99INS.rdoParameters("AZ99007").Value = Val(AZ99.S007)
			AZ99INS.rdoParameters("AZ99008").Value = Val(AZ99.S008)
			'UPGRADE_WARNING: オブジェクト MKKCMN.ZAFIXSTR_SUB() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			AZ99INS.rdoParameters("AZ99009").Value = MKKCMN.ZAFIXSTR_SUB(AZ99.S009)
		End If
		
	End Sub
	Sub AZ99NSET_SUB()
		AZ99INS.rdoParameters(0).Name = "AZ99001"
		AZ99INS.rdoParameters(1).Name = "AZ99002"
		AZ99INS.rdoParameters(2).Name = "AZ99003"
		AZ99INS.rdoParameters(3).Name = "AZ99004"
		AZ99INS.rdoParameters(4).Name = "AZ99005"
		AZ99INS.rdoParameters(5).Name = "AZ99006"
		AZ99INS.rdoParameters(6).Name = "AZ99007"
		AZ99INS.rdoParameters(7).Name = "AZ99008"
		AZ99INS.rdoParameters(8).Name = "AZ99009"
		
	End Sub
End Module