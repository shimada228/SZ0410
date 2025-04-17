Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ARQCNBAS
	'-------------------------------------------
	' <Create>
	'  Date 97.07.29
	'  K.Tsubata
	
	' <MODIFY>
	'  DATE 99.11.24  MKK 仕入在庫管理システム用
	'  K.YOSHINO
	'
	'  DATE 99.12.09  MKK ZADISCN_SUB を元に戻す
	'  K.YOSHINO
	'-------------------------------------------
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetPrivateProfileInt Lib "kernel32"  Alias "GetPrivateProfileIntA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
	
	
	'ODBC API用関数宣言
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function SQLGetInfo Lib "ODBC32.DLL" (ByVal hdbc As Integer, ByVal fInfoType As Short, ByRef rgbInfoValue As Any, ByVal cbInfoMax As Short, ByRef cbInfoOut As Short) As Short
	Declare Function SQLGetInfoString Lib "ODBC32.DLL"  Alias "SQLGetInfo"(ByVal hdbc As Integer, ByVal fInfoType As Short, ByVal rgbInfoValue As String, ByVal cbInfoMax As Short, ByRef cbInfoOut As Short) As Short
	Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Integer, ByVal fDirection As Short, ByVal szDSN As String, ByVal cbDSNMax As Short, ByRef pcbDSN As Short, ByVal szDescription As String, ByVal cbDescriptionMax As Short, ByRef pcbDescription As Short) As Short
	Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (ByRef env As Integer) As Short
	Declare Function GetComputerName Lib "Kernel32.dll"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	'ODBC API用定数宣言
	Private Const SQL_DBMS_NAME As Integer = 17
	Private Const SQL_SERVER_NAME As Integer = 13
	Private Const SQL_ERROR As Integer = -1
	Private Const SQL_INVALID_HANDLE As Integer = -2
	Private Const SQL_NO_DATA_FOUND As Integer = 100
	Private Const SQL_SUCCESS As Integer = 0
	Private Const SQL_SUCCESS_WITH_INFO As Integer = 1
	Private Const SQL_FETCH_NEXT As Integer = 1
	
	'レジストリ取得用API関数宣言
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	'レジストリ取得用API定数宣言
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const KEY_QUERY_VALUE As Short = &H1s
	Public Const ERROR_SUCCESS As Short = 0
	
	
	'定数宣言
	'Private Const ININAME = "Smile.ini"     'SMILE情報格納ファイル名
	Private Const ININAME As String = "MKK.ini" 'SMILE情報格納ファイル名
	
	
	
	
	'サブルーチン内部変数
	Public ZACN_DOCNCT As Boolean '接続ダイアログの終了モード（True:接続／False:中止）
	
	'結果引渡しパラメータ変数
	Public ZACN_USERID As String '接続したユーザ名
	Public ZACN_PASSWORD As String '接続したパスワード
	Public ZACN_DBNAME As String '接続したデータソース名
	'------------------------------------------------------------
	'【関数名】 コネクトサブルーチン
	'
	'【機  能】 ODBCによりﾃﾞｰﾀﾍﾞｰｽに接続する。
	'          使用ﾃﾞｰﾀﾍﾞｰｽがSQLServerかOracleかを判断し結果引き渡しﾊﾟﾗﾒｰﾀにｾｯﾄする。
	'           SQLServerの場合はSmile.iniからﾃﾞｰﾀﾍﾞｰｽ名を取得してﾃﾞｰﾀﾍﾞｰｽを移動する。
	'          RDO命令のｳｪｲﾄ時間をSmile.iniより取得し、結果引き渡しﾊﾟﾗﾒｰﾀにｾｯﾄする。
	'          (Smile.iniに指定が無ければﾃﾞﾌｫﾙﾄ値として「3」をｾｯﾄする)
	'
	'
	'【戻り値】 Boolean型
	'             True  :接続成功
	'            False  :接続失敗
	'
	'【関数仕様】
	'   Public Function ZACN_SUB(Optional USR As String, Optional PASSW As String, Optional DLGFLG As Boolean) As Boolean
	'     ＜プロシージャ引数＞
	'        DLGFLG As Integer  省略可
	'                            True:接続情報入力のダイアログを必ず表示する。
	'                           False:接続情報入力のダイアログを必要な時のみ表示する。
	'        USR As String      省略可。OLE Serverで起動された時のみ使用。
	'                           OLE Clientで接続したﾕｰｻﾞ名。
	'        PASSW As String    省略可。OLE Serverで起動された時のみ使用。
	'                           OLE Clientで接続したﾊﾟｽﾜｰﾄﾞ。
	'
	'     ＜結果引渡パラメータ＞
	'        Public ZACN_RCN As rdoConnection   '接続情報オブジェクト
	'        Public ZACN_DB As Integer          '使用データベース    ORCL:ORACLE
	'                                                               SQLSRV:SQLServer
	'        Public ZACN_TIME As Long           'RDO命令のｳｪｲﾄ時間
	'        Public ZACN_USERID As String       '接続したユーザ名
	'        Public ZACN_PASSWORD As String     '接続したパスワード
	'        Public ZACN_DBNAME As String       '接続したデータソース名
	'
	'【使用例】
	'   Public Sub Main()
	'       If ZACN_SUB() = False Then Exit Sub
	'       PRGTESTFRM.Show
	'   End Sub
	'
	'------------------------------------------------------------
	'Public Function ZACN_SUB(Optional DLGFLG As Boolean, Optional USR As String, Optional PASSW As String) As Boolean
	Public Function ZACN_SUB(Optional ByRef DLGFLG As Boolean = False, Optional ByRef USR As String = "", Optional ByRef PASSW As String = "", Optional ByRef DSN As String = "") As Boolean
		
		
		Dim StrPos As Short
		Dim StrPos2 As Short
		Dim USERINFOKEY As String
		Dim USERINFO As String
		Dim CONSTR As String
		Dim Ret As Integer
		Dim GETSTRWORK As New VB6.FixedLengthString(1024)
		Dim ININAMESTR As String 'Smile.iniのフルパスファイル名格納
		Dim SQL_STR As String 'SQL文字列格納
		Dim i As Integer
		Dim UNIFIEDLOGIN As Boolean
		Dim USEDB As String
		Dim UDB As RDO.rdoQuery
		Dim SName As String 'DMO接続のためのサーバ名
		Dim CompName As New VB6.FixedLengthString(32)
		Dim OServer As Object 'DMO用 ADD-1998/10/27 for SQLServer7
		'Dim OServer As New SQLOLE.SQLServer  'SQLServer用オブジェクト
		Dim DRVNAME As New VB6.FixedLengthString(128)
		Dim DRVSTRNUM As Short
		Dim DBTYPE As String
		Dim SS_SEC_MOD As Integer 'SQLServer7のセキュリティモード 98/11/4
		
		'ダイアログ表示の引数が省略されていたら、Falseとみなす。
		'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
		If IsNothing(DLGFLG) Then DLGFLG = False
		
		'MKK.iniフルパスファイル名格納      (参照しない）
		If Right(CurDir(), 1) = "\" Then
			ININAMESTR = CurDir() & ININAME & Chr(0)
		Else
			ININAMESTR = CurDir() & "\" & ININAME & Chr(0)
		End If
		
		On Error GoTo ZACN_ERR
		
		RDOrdoEngine_definst.rdoDefaultCursorDriver = RDO.CursorDriverConstants.rdUseIfNeeded
		RdoEnv = RDOrdoEngine_definst.rdoEnvironments(0)
		
		On Error Resume Next
		
		'UPGRADE_ISSUE: 定数 vbSModeStandalone はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		'UPGRADE_ISSUE: App プロパティ App.StartMode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"' をクリックしてください。
		If App.StartMode = vbSModeStandalone Then
			'独立型で起動されている時のみＤＳＮ接続用の引数があるかチェック(EXEの引数から接続文字列を取得してみる)
			StrPos = InStr(VB.Command(), ":")
			If StrPos > 0 And StrPos < Len(VB.Command()) And DLGFLG = False Then
				'DSN接続文字列引数があって、ダイアログを強制表示しないモードなら接続してみる
				Err.Clear()
				ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, Mid(VB.Command(), StrPos + 1))
				If Err.Number = 0 Then
					'データベースタイプをセット(実際にコネクトした情報から取得)
					Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_DBMS_NAME, DRVNAME.Value, 128, DRVSTRNUM)
					If Ret <> SQL_SUCCESS Then
						'ドライバ名取得失敗
						ZACN_SUB = False
						'UPGRADE_NOTE: オブジェクト ZACN_RCN をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
						ZACN_RCN = Nothing
						Exit Function
					End If
					If InStr(UCase(Left(DRVNAME.Value, DRVSTRNUM)), "SQL SERVER") > 0 Then
						'セキュリティモードはどちらだろうとＤＭＯ接続はしないので標準セキュリティのフラグとする
						'（正しいユーザ名はZACN_USERIDに入っているのでわざわざＤＭＯで再度取得する必要はない）
						ZACN_DB = SQLSRV
						UNIFIEDLOGIN = False
					Else
						ZACN_DB = ORCL
						UNIFIEDLOGIN = False
					End If
					GoTo ZACN_CONOK
				End If
			End If
		End If
		
		' iniからDSN接続文字列を取得
		'99/11/24 MOD START FOR MKK
		'    Ret = GetPrivateProfileString("CONNECT", "DBNAME", "", GETSTRWORK, Len(GETSTRWORK), ININAMESTR)
		'    ZACN_DBNAME = StrConv(LeftB(StrConv(GETSTRWORK, vbFromUnicode), Ret), vbUnicode)
		
		ZACN_DBNAME = DSN
		'99/11/24 MOD END FOR MKK
		
		If Trim(ZACN_DBNAME) = "" Then
			'        MsgBox "Smile.iniにデータソース名が記述されてないので接続できません。", vbCritical, "データベース接続エラー"
			MsgBox("Mkk.iniにデータソース名が記述されてないので接続できません。", MsgBoxStyle.Critical, "データベース接続エラー")
			GoTo ZACN_EXIT
		End If
		ZACN_USERID = ""
		ZACN_PASSWORD = ""
		
		' iniからデータベースタイプを取得
		'99/11/24 MOD START FOR MKK
		'    Ret = GetPrivateProfileString("CONNECT", "DBTYPE", "", GETSTRWORK, Len(GETSTRWORK), ININAMESTR)
		'    DBTYPE = UCase(StrConv(LeftB(StrConv(GETSTRWORK, vbFromUnicode), Ret), vbUnicode))
		
		DBTYPE = "ORA"
		'99/11/24 MOD END FOR MKK
		
		Select Case DBTYPE
			Case "ORA"
				ZACN_DB = ORCL
				UNIFIEDLOGIN = False
			Case "SS_H"
				ZACN_DB = SQLSRV
				UNIFIEDLOGIN = False
			Case Else
				If DBTYPE = "SS_T" Then
					ZACN_DB = SQLSRV
				Else
					'DSNからデータベースのタイプを取得
					If ZACN_GETDBTYPE() = False Then
						MsgBox("データソース名が不正です。接続できません。", MsgBoxStyle.Critical, "データベース接続エラー")
						GoTo ZACN_EXIT
					End If
				End If
				
				'統合セキュリティフラグはデフォルトＯＦＦに。
				UNIFIEDLOGIN = False
				If ZACN_DB = SQLSRV Then
					'SQL Serverの統合セキュリティ用に、DSNだけで接続してみる
					Err.Clear()
					UNIFIEDLOGIN = False
					ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, "", ""))
					If Err.Number = 0 Then
						'接続に成功した
						UNIFIEDLOGIN = True
						'SQLServerが6.5と7のときには動作が異なるためバージョンを取得
						Ret = SQLGetInfoString(ZACN_RCN.hdbc, 18, DRVNAME.Value, 128, DRVSTRNUM)
						If Ret <> SQL_SUCCESS Then
							DBVersion = 0
						Else
							DBVersion = Val(Left(DRVNAME.Value, InStr(DRVNAME.Value, Chr(0)) - 1))
						End If
						
						'ＤＭＯで再接続するため、データソースより接続サーバを取得
						Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_SERVER_NAME, DRVNAME.Value, 128, DRVSTRNUM)
						If Ret <> SQL_SUCCESS Then '取得に失敗
							SName = ""
						Else
							SName = Left(DRVNAME.Value, DRVSTRNUM)
						End If
						
						Ret = GetComputerName(CompName.Value, 32)
						If Left(DRVNAME.Value, DRVSTRNUM) = Left(CompName.Value, 32) Then
							SName = ""
						End If
						DRVNAME.Value = ""
						
						If DBVersion = 7 Then 'SQLServer7
							OServer = CreateObject("SQLDMO.SQLServer")
							'ＤＭＯでの接続
							'UPGRADE_WARNING: オブジェクト OServer.LoginTimeout の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.LoginTimeout = ZACN_TIME
							'UPGRADE_WARNING: オブジェクト OServer.LoginSecure の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.LoginSecure = True 'ｾｷｭﾘﾃｨｵﾌﾟｼｮﾝ
							'UPGRADE_WARNING: オブジェクト OServer.Connect の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.Connect(ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD)
							If Err.Number = 0 Then
								SS_SEC_MOD = 100
								'UPGRADE_WARNING: オブジェクト OServer.Integratedsecurity の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								SS_SEC_MOD = OServer.Integratedsecurity.securitymode 'SecurityMode
								'UPGRADE_WARNING: オブジェクト OServer.TrueLogin の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								ZACN_USERID = OServer.TrueLogin
								'UPGRADE_WARNING: オブジェクト OServer.Password の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								ZACN_PASSWORD = OServer.Password
								'UPGRADE_WARNING: オブジェクト OServer.Disconnect の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								OServer.Disconnect() ' disconnect method
								Select Case SS_SEC_MOD
									Case 1 'Integrated Security
									Case Else 'Mixed Security or Normal Security
										If sGetSec(ZACN_DBNAME) = False Then
											UNIFIEDLOGIN = False
											GoTo NOT_UNIFIED
										End If
										'Case Else 'Normal Security
										'    UNIFIEDLOGIN = False
										'    GoTo NOT_UNIFIED
								End Select
								If Err.Number <> 0 Then
									MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "データベース接続エラー")
									GoTo ZACN_EXIT
								End If
							Else
								MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "データベース接続エラー")
								GoTo ZACN_EXIT
							End If
						Else 'SQLServer6.5
							OServer = CreateObject("SQLOLE.SQLServer")
							'ＤＭＯでの接続
							'UPGRADE_WARNING: オブジェクト OServer.LoginTimeout の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.LoginTimeout = ZACN_TIME
							'UPGRADE_WARNING: オブジェクト OServer.LoginSecure の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.LoginSecure = True 'ｾｷｭﾘﾃｨｵﾌﾟｼｮﾝ
							'UPGRADE_WARNING: オブジェクト OServer.Connect の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
							OServer.Connect(ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD)
							If Err.Number = 0 Then
								'UPGRADE_WARNING: オブジェクト OServer.TrueLogin の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								ZACN_USERID = OServer.TrueLogin
								'UPGRADE_WARNING: オブジェクト OServer.Password の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								ZACN_PASSWORD = OServer.Password
								
								'UPGRADE_WARNING: オブジェクト OServer.Disconnect の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
								OServer.Disconnect() ' disconnect method
								If Err.Number <> 0 Then
									MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "データベース接続エラー")
									GoTo ZACN_EXIT
								End If
							Else
								MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "データベース接続エラー")
								GoTo ZACN_EXIT
							End If
						End If
						
						GoTo ZACN_CONOK
					Else
						For i = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
							If RDOrdoEngine_definst.rdoErrors(i).SQLState = "28000" And DBTYPE <> "SS_T" Then
								If RDOrdoEngine_definst.rdoErrors(i).Number = 4002 Then 'ADD 98/11
									'統合セキュリティではないので処理継続
									GoTo NOT_UNIFIED
								ElseIf RDOrdoEngine_definst.rdoErrors(i).Number = 18456 Then  'ADD98/11
									'統合セキュリティでSQL Serverの使用権限が無い
									MsgBox("データベースを使用する権限のあるユーザでＮＴにログオンしなおしてプログラムを再実行して下さい。", MsgBoxStyle.Critical, "データベース接続エラー")
									
									'入っていたエラーをクリア
									RDOrdoEngine_definst.rdoErrors.Clear()
									GoTo ZACN_EXIT
								End If
							ElseIf RDOrdoEngine_definst.rdoErrors(i).SQLState = "08004" Then 
								'統合セキュリティでSQL Serverの使用権限が無い
								MsgBox("データベースを使用する権限のあるユーザでＮＴにログオンしなおしてプログラムを再実行して下さい。", MsgBoxStyle.Critical, "データベース接続エラー")
								
								'入っていたエラーをクリア
								RDOrdoEngine_definst.rdoErrors.Clear()
								GoTo ZACN_EXIT
							ElseIf RDOrdoEngine_definst.rdoErrors(i).SQLState = "37000" And RDOrdoEngine_definst.rdoErrors(i).Number = 18452 Then  'change 98/11
								'SQLServer７の場合にはSQLStateが37000でNumberが18452で統合セキュリティではないので処理継続
								GoTo NOT_UNIFIED
								
							End If
						Next i
						'統合セキュリティで、それ以外のエラー
						GoTo ZACN_ERR
					End If
				End If
		End Select
		
NOT_UNIFIED: 
		'ﾕｰｻﾞ名･ﾊﾟｽﾜｰﾄﾞの引数が省略されていたか、引数のﾕｰｻﾞ名が空白だったら接続情報をSmile.iniから取得
		'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
		If IsNothing(USR) Or USR = "" Then
			Ret = GetPrivateProfileString("CONNECT", "USERID", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
			'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
			'UPGRADE_ISSUE: LeftB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
			ZACN_USERID = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
		Else
			ZACN_USERID = USR
		End If
		'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
		If IsNothing(PASSW) Or USR = "" Then
			Ret = GetPrivateProfileString("CONNECT", "PASSWORD", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
			'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
			'UPGRADE_ISSUE: LeftB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
			ZACN_PASSWORD = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
		Else
			ZACN_PASSWORD = PASSW
		End If
		
		'SQLServer7のときにはSmile.IniのUserIDとPasswordのいずれかが空白のときにはダイアログ強制表示 98/11/06
		If ZACN_DB = SQLSRV And DBVersion = 7 Then
			If ZACN_USERID = "" Or ZACN_PASSWORD = "" Then
				DLGFLG = True
			End If
		End If
		
		'ダイアログ強制表示モードでなければ、この条件で接続してみる
		If DLGFLG = False Then
			Err.Clear()
			ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
			If Err.Number = 0 Then
				'接続に成功したので終了
				GoTo ZACN_CONOK
			End If
		End If
		
		'接続に失敗したか、ダイアログ強制表示だったので、接続情報セットのダイアログ表示
		Do 
			On Error GoTo 0
			ARQCNFRM.ShowDialog()
			If ZACN_DOCNCT = False Then GoTo ZACN_EXIT
			On Error Resume Next
			
			Err.Clear()
			ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
			If Err.Number = 0 Then
				'接続成功
				GoTo ZACN_CONOK
			Else
				'接続失敗
				If ZACN_ERR_SUB() = False Then GoTo ZACN_EXIT
			End If
		Loop 
		
		On Error GoTo ZACN_ERR
		
ZACN_CONOK: 
		' iniからTIMEOUT秒数を取得
		'99/11/24 MOD START FOR MKK
		'    ZACN_TIME = GetPrivateProfileInt("RDO", "TIMEOUT", 3, ININAMESTR)
		
		ZACN_TIME = CInt(WG_TIMEOUT)
		'99/11/24 MOD END FOR MKK
		
		'データベースバージョンを取得
		Ret = SQLGetInfoString(ZACN_RCN.hdbc, 18, DRVNAME.Value, 128, DRVSTRNUM)
		If Ret <> SQL_SUCCESS Then
			DBVersion = 0
		Else
			DBVersion = Val(Left(DRVNAME.Value, InStr(DRVNAME.Value, Chr(0)) - 1))
		End If
		
		'接続情報をZACN_USERID,ZACN_PASSWORD,ZACN_DBNAMEにセット
		CONSTR = ZACN_RCN.Connect
		If UNIFIEDLOGIN = False Then
			USERINFOKEY = "UID=" 'ユーザ名
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub ZACN_USERINFO
			ZACN_USERID = USERINFO
			USERINFOKEY = "PWD=" 'パスワード
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub ZACN_USERINFO
			ZACN_PASSWORD = USERINFO
		End If
		USERINFOKEY = "DSN=" 'データソース名
		'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
		GoSub ZACN_USERINFO
		ZACN_DBNAME = USERINFO
		
		If ZACN_DB = SQLSRV Then 'SQLServerを使用
			'データベースを移動
			Ret = GetPrivateProfileString("CONNECT", "USEDB", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			If Ret > 0 Then
				'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
				'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
				'UPGRADE_ISSUE: LeftB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
				USEDB = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
				UDB = ZACN_RCN.CreateQuery("UDB", "USE " & USEDB)
				UDB.Execute()
				If Err.Number <> 0 Then
					UDB.Close()
					GoTo ZACN_ERR
				End If
				UDB.Close()
			End If
			
			'統合セキュリティの場合､DMOでユーザの接続情報を取得
			'If UNIFIEDLOGIN Then
			'ＤＭＯで再接続するため、データソースより接続サーバを取得
			'    Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_SERVER_NAME, DRVNAME, 128, DRVSTRNUM)
			'    If Ret <> SQL_SUCCESS Then      '取得に失敗
			'        SName = ""
			'    Else
			'        SName = Left(DRVNAME, DRVSTRNUM)
			'    End If
			'
			'    Ret = GetComputerName(CompName, 32)
			'    If Left(DRVNAME, DRVSTRNUM) = Left(CompName, 32) Then
			'        SName = ""
			'    End If
			'    DRVNAME = ""
			'
			'    'ＤＭＯでの接続
			'    OServer.LoginTimeout = ZACN_TIME
			'    OServer.LoginSecure = True    'ｾｷｭﾘﾃｨｵﾌﾟｼｮﾝ
			'    OServer.Connect ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD
			'    If Err.Number = 0 Then
			'        ZACN_USERID = OServer.TrueLogin
			'        ZACN_PASSWORD = OServer.Password
			'
			'        OServer.Disconnect              ' disconnect method
			'        If Err.Number <> 0 Then
			'            MsgBox Err.Number - vbObjectError & vbCr & Err.Description, vbCritical, "データベース接続エラー"
			'            GoTo ZACN_EXIT
			'        End If
			'    Else
			'        MsgBox Err.Number - vbObjectError & vbCr & Err.Description, vbCritical, "データベース接続エラー"
			'        GoTo ZACN_EXIT
			'    End If
			'End If
		End If
		
		ZACN_SUB = True
		Exit Function
		
ZACN_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "データベース接続エラー")
		
		'入っていたエラーをクリア
		RDOrdoEngine_definst.rdoErrors.Clear()
		
ZACN_EXIT: 
		Err.Clear()
		If Not (ZACN_RCN Is Nothing) Then ZACN_RCN.Close()
		If Not (RdoEnv Is Nothing) Then RdoEnv.Close()
		ZACN_USERID = ""
		ZACN_PASSWORD = ""
		ZACN_DBNAME = ""
		
		ZACN_SUB = False
		Exit Function
		
ZACN_USERINFO: 
		USERINFO = ""
		StrPos = InStr(CONSTR, USERINFOKEY)
		If StrPos > 0 Then
			StrPos = StrPos + 4
			StrPos2 = InStr(StrPos, CONSTR, ";")
			If StrPos2 > 0 Then
				USERINFO = Mid(CONSTR, StrPos, StrPos2 - StrPos)
			End If
		End If
		'UPGRADE_WARNING: Return に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		Return 
	End Function
	
	
	'------------------------------------------------------------
	'【関数名】 ディスコネクトサブルーチン
	'
	'【機  能】 ODBC接続したﾃﾞｰﾀﾍﾞｰｽを切断する。
	'
	'【仕様】
	'        ODBC接続したﾃﾞｰﾀﾍﾞｰｽを切断する。
	'        切断に成功すればTrueを返す。
	'        失敗した場合はエラーメッセージ表示後、Falseを返す。
	'
	'
	'【戻り値】 Boolean型
	'             True  :切断成功
	'            False  :切断失敗
	'
	'【関数仕様】
	'   Public Function ZADISCN_SUB() As Boolean
	'     ＜引渡パラメータ＞
	'        Public RCN As rdoConnection   '接続情報オブジェクト
	'
	'【使用例】
	'       Dim Ret As Boolean
	'       Ret = ZADISCN_SUB()
	'       End
	'------------------------------------------------------------
	Public Function ZADISCN_SUB() As Boolean
		'Public Function ZADISCN_SUB(Optional zRCN As Variant) As Boolean  '99/11/24 FOR MKK -> 99/12/09 DEL KTT YOSHINO
		
		On Error GoTo ZADISCN_ERR
		
		'99/12/09 復活 KTT YOSHINO ↓
		ZACN_RCN.Close()
		ZADISCN_SUB = True
		'99/12/09 復活 KTT YOSHINO ↑
		
		'------------------------------- 99/12/09 DEL START KTT YOSHINO ↓
		''99/11/24 ADD START FOR MKK
		'    If Not IsMissing(zRCN) Then
		'        zRCN.Close
		'    Else
		'        If WG_DBLINK = 1 Then
		'            ZACN_SZRCN.Close
		'        Else
		'            ZACN_UARCN.Close
		'            ZACN_SZRCN.Close
		'            ZACN_GCRCN.Close
		'        End If
		'    End If
		'    ZADISCN_SUB = True
		''99/11/24 ADD END FOR MKK
		'------------------------------- 99/12/09 DEL END KTT YOSHINO ↑
		
		Exit Function
		
ZADISCN_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "データベース切断エラー")
		
		'入っていたエラーをクリア
		RDOrdoEngine_definst.rdoErrors.Clear()
		
		ZADISCN_SUB = False
	End Function
	
	
	Private Function ZACN_CONSTR(ByRef DSN As String, ByRef USER As String, ByRef PASSW As String) As String
		ZACN_CONSTR = "DSN=" & DSN & ";UID=" & USER & ";PWD=" & PASSW & ";"
	End Function
	
	Private Function ZACN_ERR_SUB() As Boolean
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		Dim MSGRet As Short
		
		Select Case Err.Number
		End Select
		
		ERR_MSG = "データベースへの接続に失敗しました。"
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & vbCr & RdoErr.Description & ":" & RdoErr.Number
		Next RdoErr
		
		If MsgBox(ERR_MSG, MsgBoxStyle.RetryCancel + MsgBoxStyle.Exclamation, "データベース接続エラー") = MsgBoxResult.Retry Then
			ZACN_ERR_SUB = True
		Else
			ZACN_ERR_SUB = False
		End If
		
		'入っていたエラーをクリア
		RDOrdoEngine_definst.rdoErrors.Clear()
	End Function
	
	
	'ZACN_DBNAMEのデータソースのデータベース種別をZACN_DBにセットする
	Public Function ZACN_GETDBTYPE() As Boolean
		Dim Ret As Short
		Dim sDSNItem As New VB6.FixedLengthString(1024)
		Dim sDRVItem As New VB6.FixedLengthString(1024)
		Dim sDSN As String
		Dim sDRV As String
		Dim iDSNLen As Short
		Dim iDRVLen As Short
		Dim lHenv As Integer '環境ﾊﾝﾄﾞﾙ
		
		ZACN_GETDBTYPE = False
		
		'ﾃﾞｰﾀｿｰｽ名を取得します。
		If SQLAllocEnv(lHenv) <> -1 Then
			Ret = SQL_SUCCESS
			Do 
				sDSNItem.Value = Space(1024)
				sDRVItem.Value = Space(1024)
				Ret = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem.Value, 1024, iDSNLen, sDRVItem.Value, 1024, iDRVLen)
				If Ret <> SQL_SUCCESS Then Exit Do
				
				'データソース名取得
				If UCase(Left(sDSNItem.Value, iDSNLen)) = UCase(ZACN_DBNAME) Then
					'ターゲットのデータソースだったので、ドライバの記述子を取得
					If InStr(UCase(Left(sDRVItem.Value, iDRVLen)), "SQL SERVER") > 0 Then
						ZACN_DB = SQLSRV
					Else
						ZACN_DB = ORCL
					End If
					ZACN_GETDBTYPE = True
					Exit Do
				End If
			Loop 
		End If
	End Function
	
	Private Function sGetSec(ByVal DSN_N As String) As Boolean
		
		Dim Ret As Integer 'Return Code
		Dim hKey As Integer 'Key Handle
		Dim lpSubKey As String 'Sub Key
		Dim phkResult As Integer 'Open Key Handle
		Dim lpValueName As String '値の名前
		Dim lpData As String 'データ
		Dim lpcbData As Integer 'データの長さ
		Dim lpType As Integer 'データのタイプ
		Dim N As Short 'Counter
		
		Const SubKey As String = "SOFTWARE\ODBC\ODBC.INI\"
		
		sGetSec = False
		
		For N = 1 To 2
			Select Case N
				Case 1
					hKey = HKEY_CURRENT_USER
				Case 2
					hKey = HKEY_LOCAL_MACHINE
			End Select
			lpSubKey = SubKey & Trim(DSN_N) & Chr(0)
			
			Ret = RegOpenKeyEx(hKey, lpSubKey, 0, KEY_QUERY_VALUE, phkResult)
			If Ret <> ERROR_SUCCESS Then
				'        MsgBox "RegOpenKeyEx Error!!  Code = " & ret
				If N = 1 Then GoTo Next_N
				Exit Function
			End If
			
			'Trusted_Connectionを取得
			lpValueName = "Trusted_Connection" & Chr(0)
			lpData = Space(256)
			lpcbData = 256
			Ret = RegQueryValueEx(phkResult, lpValueName, 0, lpType, lpData, lpcbData)
			If Ret <> ERROR_SUCCESS Then
				sGetSec = False
			Else
				sGetSec = True
				Ret = RegCloseKey(phkResult)
				Exit Function
			End If
			
			Ret = RegCloseKey(phkResult)
			'If Ret = ERROR_SUCCESS Then
			'    Exit Function
			'End If
Next_N: 
		Next N
		
	End Function
End Module