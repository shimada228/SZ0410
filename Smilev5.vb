Option Strict Off
Option Explicit On
Module SMILEV5BAS
	'|||||||||||||||||||||||||||||||||
	'@@@@ƒRƒ“ƒXƒ^ƒ“ƒgƒ[ƒN
	'|||||||||||||||||||||||||||||||||
	' ƒXƒy[ƒXƒNƒŠƒA—p
	Public Const SPS As String = ""
	
	' ’è”
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
	
	' ƒXƒCƒbƒ`—p’è”
	' ‹¤’Ê
	Public Const F_OFF As Short = 0
	Public Const F_ON As Short = 1
	' ‚d‚m‚c‚r‚v
	Public Const F_ADD As Short = 1
	Public Const F_REP As Short = 2
	Public Const F_DEL As Short = 3
	Public Const F_DUM As Short = 4
	Public Const F_NXT As Short = 8
	Public Const F_END As Short = 9
	' ‚w‚w‚w‚w‚a‚`‚j‚r‚v
	Public Const F_YES As Short = 1
	' ‚d‚q‚q‚r‚v
	Public Const F_ERR As Short = 1
	' ‚h‚m‚s‚r‚v
	Public Const F_INT As Short = 1
	Public Const F_SLT As Short = 2
	' ‚w‚w‚w‚w‚h‚r‚v
	Public Const F_INV As Short = 1
	Public Const F_SKP As Short = 2
	Public Const F_GET As Short = 3
	' ‚e‚r‚s‚r‚v^‚r‚e‚r‚r‚v
	Public Const F_FST As Short = 1
	' ‚b‚`‚m‚r‚v
	Public Const F_CAN As Short = 1
	' ‚w‚w‚w‚w‚n‚o‚m‚r‚v
	Public Const F_CLS As Short = 0
	Public Const F_OPN As Short = 1
	' ‚x‚l‚c‚r‚v
	Public Const F_YM As Short = 1
	Public Const F_YMD As Short = 2
	' ‚n‚q‚`‚b‚k‚d—p     95/11’Ç‰Á
	Public SQL As String 'SQL•¶Ši”[—pÜ°¸
	Public GINITGLUE As String 'CONNECTˆ—”»’f—p
	Public ROW As Short 's²İÃŞ¨¹°Àæ“¾—p
	' ‚q‚c‚n—p     97/08/12 ’Ç‰Á
	Public Const ORCL As Short = 0 'ZACN_DB•Ï”‚Ì’l(ORACLE)
	Public Const SQLSRV As Short = 1 'ZACN_DB•Ï”‚Ì’l(SQL Server)
	Public ZACN_RCN As RDO.rdoConnection 'Ú‘±î•ñƒIƒuƒWƒFƒNƒg
	Public ZACN_DB As Short 'g—pƒf[ƒ^ƒx[ƒX (0:ORACLE / 1:SQL Server)
	Public ZACN_TIME As Integer 'RDO–½—ß‚Ì³ª²ÄŠÔ
	
	Structure FILNAME_S
		'UPGRADE_WARNING: ŒÅ’è’·•¶š—ñ‚ÌƒTƒCƒY‚Íƒoƒbƒtƒ@‚É‡‚í‚¹‚é•K—v‚ª‚ ‚è‚Ü‚·B Ú×‚É‚Â‚¢‚Ä‚ÍA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"' ‚ğƒNƒŠƒbƒN‚µ‚Ä‚­‚¾‚³‚¢B
		<VBFixedString(40),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=40)> Public NAME() As Char
		'   DBLINK As String * 10               '99/12/09 DEL KTT YOSHINO
	End Structure
	
	'97/07/03Del Global B_STATUS As Integer               ' ½Ã°À½º°ÄŞ
	
	' HelpContextID ó‚¯“n‚µ—p     97/09/30 ’Ç‰Á
	Public SM_HelpContextID As Integer
	
	
	' Intersolv ODBCƒhƒ‰ƒCƒo[g—p‹æ•ª 98/09/14’Ç‰Á
	Public ReQue As Boolean 'TrueFg—p ^FalseFg—p‚µ‚È‚¢
	
	'DataBaseVersion
	Public DBVersion As Double '98/11/30 ’Ç‰Á
End Module