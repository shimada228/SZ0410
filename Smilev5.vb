Option Strict Off
Option Explicit On
Module SMILEV5BAS
	'−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−
	'　　　　コンスタントワーク
	'−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−−
	' スペースクリア用
	Public Const SPS As String = ""
	
	' 定数
	Public Const n0 As Short = 0
	Public Const n1 As Short = 1
	Public Const n2 As Short = 2
	Public Const n3 As Short = 3
	Public Const n4 As Short = 4
	Public Const n5 As Short = 5
	Public Const n6 As Short = 6
	Public Const n7 As Short = 7
	Public Const n8 As Short = 8
	Public Const n9 As Short = 9
	
	' スイッチ用定数
	' 共通
	Public Const F_OFF As Short = 0
	Public Const F_ON As Short = 1
	' ＥＮＤＳＷ
	Public Const F_ADD As Short = 1
	Public Const F_REP As Short = 2
	Public Const F_DEL As Short = 3
	Public Const F_DUM As Short = 4
	Public Const F_NXT As Short = 8
	Public Const F_END As Short = 9
	' ＸＸＸＸＢＡＫＳＷ
	Public Const F_YES As Short = 1
	' ＥＲＲＳＷ
	Public Const F_ERR As Short = 1
	' ＩＮＴＳＷ
	Public Const F_INT As Short = 1
	Public Const F_SLT As Short = 2
	' ＸＸＸＸＩＳＷ
	Public Const F_INV As Short = 1
	Public Const F_SKP As Short = 2
	Public Const F_GET As Short = 3
	' ＦＳＴＳＷ／ＳＦＳＳＷ
	Public Const F_FST As Short = 1
	' ＣＡＮＳＷ
	Public Const F_CAN As Short = 1
	' ＸＸＸＸＯＰＮＳＷ
	Public Const F_CLS As Short = 0
	Public Const F_OPN As Short = 1
	' ＹＭＤＳＷ
	Public Const F_YM As Short = 1
	Public Const F_YMD As Short = 2
	' ＯＲＡＣＬＥ用     95/11追加
	Public SQL As String 'SQL文格納用ﾜｰｸ
	Public GINITGLUE As String 'CONNECT処理判断用
	Public ROW As Short '行ｲﾝﾃﾞｨｹｰﾀ取得用
	' ＲＤＯ用     97/08/12 追加
	Public Const ORCL As Short = 0 'ZACN_DB変数の値(ORACLE)
	Public Const SQLSRV As Short = 1 'ZACN_DB変数の値(SQL Server)
	Public ZACN_RCN As RDO.rdoConnection '接続情報オブジェクト
	Public ZACN_DB As Short '使用データベース (0:ORACLE / 1:SQL Server)
	Public ZACN_TIME As Integer 'RDO命令のｳｪｲﾄ時間
	
	Structure FILNAME_S
		'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' をクリックしてください。
		<VBFixedString(40),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=40)> Public NAME() As Char
		'   DBLINK As String * 10               '99/12/09 DEL KTT YOSHINO
	End Structure
	
	'97/07/03Del Global B_STATUS As Integer               ' ｽﾃｰﾀｽｺｰﾄﾞ
	
	' HelpContextID 受け渡し用     97/09/30 追加
	Public SM_HelpContextID As Integer
	
	
	' Intersolv ODBCドライバー使用区分 98/09/14追加
	Public ReQue As Boolean 'True：使用 ／False：使用しない
	
	'DataBaseVersion
	Public DBVersion As Double '98/11/30 追加
End Module