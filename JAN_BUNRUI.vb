Option Strict Off
Option Explicit On
Module JAN_BUNRUIBAS
	'-------------------------------------------------------------
	'   三井観光開発（株）
	'   仕入在庫管理システム
	'   JAN_BUNRUIマスタ(JAN_BUNRUI)
	'-------------------------------------------------------------
	Public Structure JAN_BUNRUI_S
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(6),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=6)> Public BK1() As Char
		BK1   		As String *6  	’ JAN_BUNRUI商品分類コード                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public BK2() As Char
		BK2   		As String *1  	’ 制御フラグ１                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public BK3() As Char
		BK3   		As String *1  	’ 制御フラグ２                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=64)> Public BK4() As Char
		BK4   		As String *64  	’ 分類名
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(14),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=14)> Public BK5() As Char
		BK5   		As String *14  	’ 分類名（カナ）                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(28),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=28)> Public BK6() As Char
		BK6   		As String *28  	’ 分類名（漢字）                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(8),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=8)> Public BK7() As Char
		BK7   		As String *8  	’ 登録日付                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(8),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=8)> Public BK8() As Char
		BK8   		As String *8  	’ 更新日付                        
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
	End Structure
	
	Structure JAN_BUNRUI_KEY0_S
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(6),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=6)> Public BK1() As Char
		BK1   		As String *6  	’ JAN_BUNRUI商品分類コード
		'UPGRADE_ISSUE: 前の行を解析できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"' をクリックしてください。
	End Structure
	Public JAN_BUNRUI_KEY0 As Short
	Public JAN_BUNRUI_BUF0 As JAN_BUNRUI_KEY0_S
	
	
	'　データバッファー
	Public JAN_BUNRUI As JAN_BUNRUI_S
	Public JAN_BUNRUIOPENSW As Short
	Public JAN_BUNRUIINVSW As Short
	
	'　ＲＤＯ関連オブジェクト
	Public JAN_BUNRUIRS As RDO.rdoResultset
	Public JAN_BUNRUIINS As RDO.rdoQuery
	Public JAN_BUNRUIUPD As RDO.rdoQuery
	Public JAN_BUNRUIDEL As RDO.rdoQuery
	Public JAN_BUNRUIRSSW As String
	Public JAN_BUNRUI_FILE As FILNAME_S
End Module