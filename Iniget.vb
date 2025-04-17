Option Strict Off
Option Explicit On
Module INIGETBAS
	'*
	'* MKK 仕入在庫管理システム用  INI取得サブルーチン
	'*
	'* 1999/11/24 KTT-YOSHINO
	'*
	
	
	'Windows 95 VB Ver4.0 API
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	
	'*-------- グローバル変数 --------*
	
	'データベース接続情報
	'Public WG_DBLINK          As String     '0:使用する  1:使用しない  '99/12/09 DEL KTT YOSHINO
	
	Public WG_UAID As String '[DATABASE] 売上売掛DBﾕｰｻﾞ名
	Public WG_UAPW As String '[DATABASE] 売上売掛DBﾊﾟｽﾜｰﾄﾞ
	Public WG_UADSN As String '[DATABASE] 売上売掛DBDSN名
	
	Public WG_SZID As String '[DATABASE] 仕入在庫DBﾕｰｻﾞ名
	Public WG_SZPW As String '[DATABASE] 仕入在庫DBﾊﾟｽﾜｰﾄﾞ
	Public WG_SZDSN As String '[DATABASE] 仕入在庫DBDSN名
	
	Public WG_GCID As String '[DATABASE] 業務間共通DBﾕｰｻﾞ名
	Public WG_GCPW As String '[DATABASE] 業務間共通DBﾊﾟｽﾜｰﾄﾞ
	Public WG_GCDSN As String '[DATABASE] 業務間共通DBDSN名
	
	'共通情報
	Public WG_TIMEOUT As String '[RDO] ﾀｲﾑｱｳﾄ時間
	Public WG_REQUERY As String '[RDO] REQUERY有無
	
	Public WG_OPCODE As String '[OPERATOR] ｵﾍﾟﾚｰﾀｺｰﾄﾞ
	Public WG_INCCODE As String '[OPERATOR] 会社ｺｰﾄﾞ
	Public WG_JGCODE As String '[OPERATOR] 事業所ｺｰﾄﾞ
	Public WG_BUSYOCODE As String '[OPERATOR] 部所ｺｰﾄﾞ
	
	Public WG_DEBUG As String '[MAIN] DEBUG  "1"の場合はOPCODEを有効とする
	
	Public WG_EXCELPATH As String '[SZPRG] EXCELﾌｧｲﾙ PATH
	
	Public WG_FAXID As String '[FAX] ﾛｸﾞｵﾝID
	Public WG_FAXPW As String '[FAX] ﾊﾟｽﾜｰﾄﾞ
	
	
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const KEY_QUERY_VALUE As Short = &H1s
	Public Const ERROR_SUCCESS As Short = 0
	Public Const ERROR_FILE_NOT_FOUND As Short = 2
	
	'+------------------------------------------------------------------+
	'| オリジナル初期設定ファイル情報                                     |
	'+------------------------------------------------------------------+
	Public OrgIniPathName As String 'パス
	Public OrgIniFileName As String 'ファイル名
	'
	Public Sub REGGET()
		
		'レジストリからＩＮＩファイルのパス、ファイル名を取得する
		
		Dim Ret As Integer 'Return Code
		Dim hKey As Integer 'Key Handle
		Dim lpSubKey As String 'Sub Key
		Dim phkResult As Integer 'Open Key Handle
		Dim lpValueName As String '値の名前
		Dim lpData As String 'データ
		Dim lpcbData As Integer 'データの長さ
		Dim lpType As Integer 'データのタイプ
		
		Const SubKey As String = "SOFTWARE\MKK"
		Const ValuePathName As String = "IniPath"
		Const ValueIniName As String = "IniFile"
		
		OrgIniPathName = ""
		OrgIniFileName = ""
		
		hKey = HKEY_LOCAL_MACHINE
		lpSubKey = SubKey & Chr(0)
		
		Ret = RegOpenKeyEx(hKey, lpSubKey, 0, KEY_QUERY_VALUE, phkResult)
		If Ret <> ERROR_SUCCESS Then
			'        MsgBox "RegOpenKeyEx Error!!  Code = " & ret
			Exit Sub
		End If
		
		'オリジナル初期設定ファイルのパスを取得
		lpValueName = ValuePathName & Chr(0)
		lpData = Space(256)
		lpcbData = 256
		Ret = RegQueryValueEx(phkResult, lpValueName, 0, lpType, lpData, lpcbData)
		If Ret <> ERROR_SUCCESS Then
			'        MsgBox "RegQueryValueEx Error!!  Code = " & ret
		Else
			OrgIniPathName = Left(lpData, InStr(lpData, Chr(0)) - 1)
		End If
		
		'オリジナル初期設定ファイル名を取得
		lpValueName = ValueIniName & Chr(0)
		lpData = Space(256)
		lpcbData = 256
		Ret = RegQueryValueEx(phkResult, lpValueName, 0, lpType, lpData, lpcbData)
		If Ret <> ERROR_SUCCESS Then
			'        MsgBox "RegQueryValueEx Error!!  Code = " & ret
		Else
			OrgIniFileName = Left(lpData, InStr(lpData, Chr(0)) - 1)
		End If
		
		Ret = RegCloseKey(phkResult)
		If Ret <> ERROR_SUCCESS Then
			'        MsgBox "RegCloseKey Error!!  Code = " & ret
			Exit Sub
		End If
		
	End Sub
	
	Private Function INIGET_ENTRY(ByVal section As String, ByVal entry As String, ByVal def_str As String, ByVal fname As String) As String
		
		'   /*                               */
		'   /* INIファイルの内容取得（個別） */
		'   /*     (Internal Function)       */
		'   /*                               */
		
		Static bUF As New VB6.FixedLengthString(256)
		Dim buftmp As String
		
		bUF.Value = ""
		
		'   INIファイルの指定エントリを取得
		If (GetPrivateProfileString(section, entry, def_str, bUF.Value, 256, fname) > 0) Then
			buftmp = Trim(bUF.Value)
		Else
			buftmp = Trim(def_str)
		End If
		
		'   文字列の最後に '\0'が付いていたなら除去する
		'UPGRADE_ISSUE: RightB$ 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
		If (RightB$(buftmp, 2) = Chr(0)) Then
			'UPGRADE_ISSUE: LenB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
			'UPGRADE_ISSUE: LeftB$ 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
			INIGET_ENTRY = LeftB$(buftmp, LenB(buftmp) - 2)
		Else
			INIGET_ENTRY = buftmp
		End If
	End Function
	
	Sub INIGET_SUB(ByVal fname As String)
		
		'   用  例 : INIGET_L_SUB("MKK.INI")
		'
		'*******************************************************************************
		'*     ＲＥＧＥＤＩＴ．ＥＸＥにて下記情報を追加しておくこと。                      *
		'*     -HKEY_LOCAL_MACHINE\SOFTWAREMKKを作成し、その下に                        *
		'*     IniPath   :INIﾌｧｲﾙのﾊﾟｽ                                                 *
		'*     IniFile   :INIﾌｧｲﾙ名　　　　　　　　　　　　　　　　　　　　　　　　　　　　*
		'*******************************************************************************
		
		Dim INI_NAME As String
		'  If App.PrevInstance Then End
		
		'レジストリからファイル名を取得する
		Call REGGET()
		
		If OrgIniPathName = "" Then
			OrgIniPathName = CurDir() & "\"
		End If
		
		If fname = "" Then
			INI_NAME = OrgIniFileName
		Else
			INI_NAME = fname
		End If
		
		INI_NAME = OrgIniPathName & INI_NAME
		
		'ＩＮＩファイルから項目取得
		
		'データベース接続情報
		'    WG_DBLINK = INIGET_ENTRY("DATABASE", "DBLINK", "", INI_NAME)  '[DATABASE] DBLINK  99/12/09 DEL KTT YOSHINO
		
		WG_UAID = INIGET_ENTRY("DATABASE", "UAID", "", INI_NAME) '[DATABASE] 売上売掛DBﾕｰｻﾞ名
		WG_UAPW = INIGET_ENTRY("DATABASE", "UAPW", "", INI_NAME) '[DATABASE] 売上売掛DBﾊﾟｽﾜｰﾄﾞ
		WG_UADSN = INIGET_ENTRY("DATABASE", "UADSN", "", INI_NAME) '[DATABASE] 売上売掛DBDSN名
		
		WG_SZID = INIGET_ENTRY("DATABASE", "SZID", "", INI_NAME) '[DATABASE] 仕入在庫DBﾕｰｻﾞ名
		WG_SZPW = INIGET_ENTRY("DATABASE", "SZPW", "", INI_NAME) '[DATABASE] 仕入在庫DBﾊﾟｽﾜｰﾄﾞ
		WG_SZDSN = INIGET_ENTRY("DATABASE", "SZDSN", "", INI_NAME) '[DATABASE] 仕入在庫DBDSN名
		
		WG_GCID = INIGET_ENTRY("DATABASE", "GCID", "", INI_NAME) '[DATABASE] 業務間共通DBﾕｰｻﾞ名
		WG_GCPW = INIGET_ENTRY("DATABASE", "GCPW", "", INI_NAME) '[DATABASE] 業務間共通DBﾊﾟｽﾜｰﾄﾞ
		WG_GCDSN = INIGET_ENTRY("DATABASE", "GCDSN", "", INI_NAME) '[DATABASE] 業務間共通DBDSN名
		
		'共通情報
		WG_TIMEOUT = INIGET_ENTRY("RDO", "TIMEOUT", "", INI_NAME) '[RDO] ﾀｲﾑｱｳﾄ時間
		WG_REQUERY = INIGET_ENTRY("RDO", "REQUERY", "", INI_NAME) '[RDO] REQUERY有無
		
		WG_OPCODE = INIGET_ENTRY("OPERATOR", "OPCODE", "", INI_NAME) '[OPERATOR] ｵﾍﾟﾚｰﾀｺｰﾄﾞ
		WG_INCCODE = INIGET_ENTRY("OPERATOR", "INCCODE", "", INI_NAME) '[OPERATOR] 会社ｺｰﾄﾞ
		WG_JGCODE = INIGET_ENTRY("OPERATOR", "JGCODE", "", INI_NAME) '[OPERATOR] 事業所ｺｰﾄﾞ
		WG_BUSYOCODE = INIGET_ENTRY("OPERATOR", "BUSYOCODE", "", INI_NAME) '[OPERATOR] 部所ｺｰﾄﾞ
		
		WG_DEBUG = INIGET_ENTRY("MAIN", "DEBUG", "", INI_NAME) '[MAIN] DEBUG  "1"の場合はOPCODEを有効とする
		
		WG_EXCELPATH = INIGET_ENTRY("SZPRG", "EXCELPATH", "", INI_NAME) '[SZPRG] EXCELﾌｧｲﾙ PATH
		
		WG_FAXID = INIGET_ENTRY("FAX", "FAXID", "", INI_NAME) '[FAX] ﾛｸﾞｵﾝID
		WG_FAXPW = INIGET_ENTRY("FAX", "FAXPW", "", INI_NAME) '[FAX] ﾊﾟｽﾜｰﾄﾞ
		
	End Sub
End Module