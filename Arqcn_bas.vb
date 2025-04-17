Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module ARQCNBAS
	'-------------------------------------------
	' <Create>
	'  Date 97.07.29
	'  K.Tsubata
	
	' <MODIFY>
	'  DATE 99.11.24  MKK �d���݌ɊǗ��V�X�e���p
	'  K.YOSHINO
	'
	'  DATE 99.12.09  MKK ZADISCN_SUB �����ɖ߂�
	'  K.YOSHINO
	'-------------------------------------------
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetPrivateProfileInt Lib "kernel32"  Alias "GetPrivateProfileIntA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
	
	
	'ODBC API�p�֐��錾
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function SQLGetInfo Lib "ODBC32.DLL" (ByVal hdbc As Integer, ByVal fInfoType As Short, ByRef rgbInfoValue As Any, ByVal cbInfoMax As Short, ByRef cbInfoOut As Short) As Short
	Declare Function SQLGetInfoString Lib "ODBC32.DLL"  Alias "SQLGetInfo"(ByVal hdbc As Integer, ByVal fInfoType As Short, ByVal rgbInfoValue As String, ByVal cbInfoMax As Short, ByRef cbInfoOut As Short) As Short
	Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Integer, ByVal fDirection As Short, ByVal szDSN As String, ByVal cbDSNMax As Short, ByRef pcbDSN As Short, ByVal szDescription As String, ByVal cbDescriptionMax As Short, ByRef pcbDescription As Short) As Short
	Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (ByRef env As Integer) As Short
	Declare Function GetComputerName Lib "Kernel32.dll"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	'ODBC API�p�萔�錾
	Private Const SQL_DBMS_NAME As Integer = 17
	Private Const SQL_SERVER_NAME As Integer = 13
	Private Const SQL_ERROR As Integer = -1
	Private Const SQL_INVALID_HANDLE As Integer = -2
	Private Const SQL_NO_DATA_FOUND As Integer = 100
	Private Const SQL_SUCCESS As Integer = 0
	Private Const SQL_SUCCESS_WITH_INFO As Integer = 1
	Private Const SQL_FETCH_NEXT As Integer = 1
	
	'���W�X�g���擾�pAPI�֐��錾
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	'���W�X�g���擾�pAPI�萔�錾
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const KEY_QUERY_VALUE As Short = &H1s
	Public Const ERROR_SUCCESS As Short = 0
	
	
	'�萔�錾
	'Private Const ININAME = "Smile.ini"     'SMILE���i�[�t�@�C����
	Private Const ININAME As String = "MKK.ini" 'SMILE���i�[�t�@�C����
	
	
	
	
	'�T�u���[�`�������ϐ�
	Public ZACN_DOCNCT As Boolean '�ڑ��_�C�A���O�̏I�����[�h�iTrue:�ڑ��^False:���~�j
	
	'���ʈ��n���p�����[�^�ϐ�
	Public ZACN_USERID As String '�ڑ��������[�U��
	Public ZACN_PASSWORD As String '�ڑ������p�X���[�h
	Public ZACN_DBNAME As String '�ڑ������f�[�^�\�[�X��
	'------------------------------------------------------------
	'�y�֐����z �R�l�N�g�T�u���[�`��
	'
	'�y�@  �\�z ODBC�ɂ���ް��ް��ɐڑ�����B
	'          �g�p�ް��ް���SQLServer��Oracle���𔻒f�����ʈ����n�����Ұ��ɾ�Ă���B
	'           SQLServer�̏ꍇ��Smile.ini�����ް��ް������擾�����ް��ް����ړ�����B
	'          RDO���߂̳��Ď��Ԃ�Smile.ini���擾���A���ʈ����n�����Ұ��ɾ�Ă���B
	'          (Smile.ini�Ɏw�肪���������̫�Ēl�Ƃ��āu3�v��Ă���)
	'
	'
	'�y�߂�l�z Boolean�^
	'             True  :�ڑ�����
	'            False  :�ڑ����s
	'
	'�y�֐��d�l�z
	'   Public Function ZACN_SUB(Optional USR As String, Optional PASSW As String, Optional DLGFLG As Boolean) As Boolean
	'     ���v���V�[�W��������
	'        DLGFLG As Integer  �ȗ���
	'                            True:�ڑ������͂̃_�C�A���O��K���\������B
	'                           False:�ڑ������͂̃_�C�A���O��K�v�Ȏ��̂ݕ\������B
	'        USR As String      �ȗ��BOLE Server�ŋN�����ꂽ���̂ݎg�p�B
	'                           OLE Client�Őڑ�����հ�ޖ��B
	'        PASSW As String    �ȗ��BOLE Server�ŋN�����ꂽ���̂ݎg�p�B
	'                           OLE Client�Őڑ������߽ܰ�ށB
	'
	'     �����ʈ��n�p�����[�^��
	'        Public ZACN_RCN As rdoConnection   '�ڑ����I�u�W�F�N�g
	'        Public ZACN_DB As Integer          '�g�p�f�[�^�x�[�X    ORCL:ORACLE
	'                                                               SQLSRV:SQLServer
	'        Public ZACN_TIME As Long           'RDO���߂̳��Ď���
	'        Public ZACN_USERID As String       '�ڑ��������[�U��
	'        Public ZACN_PASSWORD As String     '�ڑ������p�X���[�h
	'        Public ZACN_DBNAME As String       '�ڑ������f�[�^�\�[�X��
	'
	'�y�g�p��z
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
		Dim ININAMESTR As String 'Smile.ini�̃t���p�X�t�@�C�����i�[
		Dim SQL_STR As String 'SQL������i�[
		Dim i As Integer
		Dim UNIFIEDLOGIN As Boolean
		Dim USEDB As String
		Dim UDB As RDO.rdoQuery
		Dim SName As String 'DMO�ڑ��̂��߂̃T�[�o��
		Dim CompName As New VB6.FixedLengthString(32)
		Dim OServer As Object 'DMO�p ADD-1998/10/27 for SQLServer7
		'Dim OServer As New SQLOLE.SQLServer  'SQLServer�p�I�u�W�F�N�g
		Dim DRVNAME As New VB6.FixedLengthString(128)
		Dim DRVSTRNUM As Short
		Dim DBTYPE As String
		Dim SS_SEC_MOD As Integer 'SQLServer7�̃Z�L�����e�B���[�h 98/11/4
		
		'�_�C�A���O�\���̈������ȗ�����Ă�����AFalse�Ƃ݂Ȃ��B
		'UPGRADE_NOTE: IsMissing() �� IsNothing() �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' ���N���b�N���Ă��������B
		If IsNothing(DLGFLG) Then DLGFLG = False
		
		'MKK.ini�t���p�X�t�@�C�����i�[      (�Q�Ƃ��Ȃ��j
		If Right(CurDir(), 1) = "\" Then
			ININAMESTR = CurDir() & ININAME & Chr(0)
		Else
			ININAMESTR = CurDir() & "\" & ININAME & Chr(0)
		End If
		
		On Error GoTo ZACN_ERR
		
		RDOrdoEngine_definst.rdoDefaultCursorDriver = RDO.CursorDriverConstants.rdUseIfNeeded
		RdoEnv = RDOrdoEngine_definst.rdoEnvironments(0)
		
		On Error Resume Next
		
		'UPGRADE_ISSUE: �萔 vbSModeStandalone �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
		'UPGRADE_ISSUE: App �v���p�e�B App.StartMode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"' ���N���b�N���Ă��������B
		If App.StartMode = vbSModeStandalone Then
			'�Ɨ��^�ŋN������Ă��鎞�݂̂c�r�m�ڑ��p�̈��������邩�`�F�b�N(EXE�̈�������ڑ���������擾���Ă݂�)
			StrPos = InStr(VB.Command(), ":")
			If StrPos > 0 And StrPos < Len(VB.Command()) And DLGFLG = False Then
				'DSN�ڑ�����������������āA�_�C�A���O�������\�����Ȃ����[�h�Ȃ�ڑ����Ă݂�
				Err.Clear()
				ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, Mid(VB.Command(), StrPos + 1))
				If Err.Number = 0 Then
					'�f�[�^�x�[�X�^�C�v���Z�b�g(���ۂɃR�l�N�g������񂩂�擾)
					Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_DBMS_NAME, DRVNAME.Value, 128, DRVSTRNUM)
					If Ret <> SQL_SUCCESS Then
						'�h���C�o���擾���s
						ZACN_SUB = False
						'UPGRADE_NOTE: �I�u�W�F�N�g ZACN_RCN ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
						ZACN_RCN = Nothing
						Exit Function
					End If
					If InStr(UCase(Left(DRVNAME.Value, DRVSTRNUM)), "SQL SERVER") > 0 Then
						'�Z�L�����e�B���[�h�͂ǂ��炾�낤�Ƃc�l�n�ڑ��͂��Ȃ��̂ŕW���Z�L�����e�B�̃t���O�Ƃ���
						'�i���������[�U����ZACN_USERID�ɓ����Ă���̂ł킴�킴�c�l�n�ōēx�擾����K�v�͂Ȃ��j
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
		
		' ini����DSN�ڑ���������擾
		'99/11/24 MOD START FOR MKK
		'    Ret = GetPrivateProfileString("CONNECT", "DBNAME", "", GETSTRWORK, Len(GETSTRWORK), ININAMESTR)
		'    ZACN_DBNAME = StrConv(LeftB(StrConv(GETSTRWORK, vbFromUnicode), Ret), vbUnicode)
		
		ZACN_DBNAME = DSN
		'99/11/24 MOD END FOR MKK
		
		If Trim(ZACN_DBNAME) = "" Then
			'        MsgBox "Smile.ini�Ƀf�[�^�\�[�X�����L�q����ĂȂ��̂Őڑ��ł��܂���B", vbCritical, "�f�[�^�x�[�X�ڑ��G���["
			MsgBox("Mkk.ini�Ƀf�[�^�\�[�X�����L�q����ĂȂ��̂Őڑ��ł��܂���B", MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
			GoTo ZACN_EXIT
		End If
		ZACN_USERID = ""
		ZACN_PASSWORD = ""
		
		' ini����f�[�^�x�[�X�^�C�v���擾
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
					'DSN����f�[�^�x�[�X�̃^�C�v���擾
					If ZACN_GETDBTYPE() = False Then
						MsgBox("�f�[�^�\�[�X�����s���ł��B�ڑ��ł��܂���B", MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
						GoTo ZACN_EXIT
					End If
				End If
				
				'�����Z�L�����e�B�t���O�̓f�t�H���g�n�e�e�ɁB
				UNIFIEDLOGIN = False
				If ZACN_DB = SQLSRV Then
					'SQL Server�̓����Z�L�����e�B�p�ɁADSN�����Őڑ����Ă݂�
					Err.Clear()
					UNIFIEDLOGIN = False
					ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, "", ""))
					If Err.Number = 0 Then
						'�ڑ��ɐ�������
						UNIFIEDLOGIN = True
						'SQLServer��6.5��7�̂Ƃ��ɂ͓��삪�قȂ邽�߃o�[�W�������擾
						Ret = SQLGetInfoString(ZACN_RCN.hdbc, 18, DRVNAME.Value, 128, DRVSTRNUM)
						If Ret <> SQL_SUCCESS Then
							DBVersion = 0
						Else
							DBVersion = Val(Left(DRVNAME.Value, InStr(DRVNAME.Value, Chr(0)) - 1))
						End If
						
						'�c�l�n�ōĐڑ����邽�߁A�f�[�^�\�[�X���ڑ��T�[�o���擾
						Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_SERVER_NAME, DRVNAME.Value, 128, DRVSTRNUM)
						If Ret <> SQL_SUCCESS Then '�擾�Ɏ��s
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
							'�c�l�n�ł̐ڑ�
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.LoginTimeout �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.LoginTimeout = ZACN_TIME
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.LoginSecure �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.LoginSecure = True '����è��߼��
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Connect �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.Connect(ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD)
							If Err.Number = 0 Then
								SS_SEC_MOD = 100
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Integratedsecurity �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								SS_SEC_MOD = OServer.Integratedsecurity.securitymode 'SecurityMode
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.TrueLogin �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								ZACN_USERID = OServer.TrueLogin
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Password �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								ZACN_PASSWORD = OServer.Password
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Disconnect �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
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
									MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
									GoTo ZACN_EXIT
								End If
							Else
								MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
								GoTo ZACN_EXIT
							End If
						Else 'SQLServer6.5
							OServer = CreateObject("SQLOLE.SQLServer")
							'�c�l�n�ł̐ڑ�
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.LoginTimeout �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.LoginTimeout = ZACN_TIME
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.LoginSecure �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.LoginSecure = True '����è��߼��
							'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Connect �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							OServer.Connect(ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD)
							If Err.Number = 0 Then
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.TrueLogin �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								ZACN_USERID = OServer.TrueLogin
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Password �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								ZACN_PASSWORD = OServer.Password
								
								'UPGRADE_WARNING: �I�u�W�F�N�g OServer.Disconnect �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								OServer.Disconnect() ' disconnect method
								If Err.Number <> 0 Then
									MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
									GoTo ZACN_EXIT
								End If
							Else
								MsgBox(Err.Number - vbObjectError & vbCr & Err.Description, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
								GoTo ZACN_EXIT
							End If
						End If
						
						GoTo ZACN_CONOK
					Else
						For i = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
							If RDOrdoEngine_definst.rdoErrors(i).SQLState = "28000" And DBTYPE <> "SS_T" Then
								If RDOrdoEngine_definst.rdoErrors(i).Number = 4002 Then 'ADD 98/11
									'�����Z�L�����e�B�ł͂Ȃ��̂ŏ����p��
									GoTo NOT_UNIFIED
								ElseIf RDOrdoEngine_definst.rdoErrors(i).Number = 18456 Then  'ADD98/11
									'�����Z�L�����e�B��SQL Server�̎g�p����������
									MsgBox("�f�[�^�x�[�X���g�p���錠���̂��郆�[�U�łm�s�Ƀ��O�I�����Ȃ����ăv���O�������Ď��s���ĉ������B", MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
									
									'�����Ă����G���[���N���A
									RDOrdoEngine_definst.rdoErrors.Clear()
									GoTo ZACN_EXIT
								End If
							ElseIf RDOrdoEngine_definst.rdoErrors(i).SQLState = "08004" Then 
								'�����Z�L�����e�B��SQL Server�̎g�p����������
								MsgBox("�f�[�^�x�[�X���g�p���錠���̂��郆�[�U�łm�s�Ƀ��O�I�����Ȃ����ăv���O�������Ď��s���ĉ������B", MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
								
								'�����Ă����G���[���N���A
								RDOrdoEngine_definst.rdoErrors.Clear()
								GoTo ZACN_EXIT
							ElseIf RDOrdoEngine_definst.rdoErrors(i).SQLState = "37000" And RDOrdoEngine_definst.rdoErrors(i).Number = 18452 Then  'change 98/11
								'SQLServer�V�̏ꍇ�ɂ�SQLState��37000��Number��18452�œ����Z�L�����e�B�ł͂Ȃ��̂ŏ����p��
								GoTo NOT_UNIFIED
								
							End If
						Next i
						'�����Z�L�����e�B�ŁA����ȊO�̃G���[
						GoTo ZACN_ERR
					End If
				End If
		End Select
		
NOT_UNIFIED: 
		'հ�ޖ���߽ܰ�ނ̈������ȗ�����Ă������A������հ�ޖ����󔒂�������ڑ�����Smile.ini����擾
		'UPGRADE_NOTE: IsMissing() �� IsNothing() �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' ���N���b�N���Ă��������B
		If IsNothing(USR) Or USR = "" Then
			Ret = GetPrivateProfileString("CONNECT", "USERID", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LeftB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			ZACN_USERID = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
		Else
			ZACN_USERID = USR
		End If
		'UPGRADE_NOTE: IsMissing() �� IsNothing() �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' ���N���b�N���Ă��������B
		If IsNothing(PASSW) Or USR = "" Then
			Ret = GetPrivateProfileString("CONNECT", "PASSWORD", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LeftB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			ZACN_PASSWORD = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
		Else
			ZACN_PASSWORD = PASSW
		End If
		
		'SQLServer7�̂Ƃ��ɂ�Smile.Ini��UserID��Password�̂����ꂩ���󔒂̂Ƃ��ɂ̓_�C�A���O�����\�� 98/11/06
		If ZACN_DB = SQLSRV And DBVersion = 7 Then
			If ZACN_USERID = "" Or ZACN_PASSWORD = "" Then
				DLGFLG = True
			End If
		End If
		
		'�_�C�A���O�����\�����[�h�łȂ���΁A���̏����Őڑ����Ă݂�
		If DLGFLG = False Then
			Err.Clear()
			ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
			If Err.Number = 0 Then
				'�ڑ��ɐ��������̂ŏI��
				GoTo ZACN_CONOK
			End If
		End If
		
		'�ڑ��Ɏ��s�������A�_�C�A���O�����\���������̂ŁA�ڑ����Z�b�g�̃_�C�A���O�\��
		Do 
			On Error GoTo 0
			ARQCNFRM.ShowDialog()
			If ZACN_DOCNCT = False Then GoTo ZACN_EXIT
			On Error Resume Next
			
			Err.Clear()
			ZACN_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACN_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
			If Err.Number = 0 Then
				'�ڑ�����
				GoTo ZACN_CONOK
			Else
				'�ڑ����s
				If ZACN_ERR_SUB() = False Then GoTo ZACN_EXIT
			End If
		Loop 
		
		On Error GoTo ZACN_ERR
		
ZACN_CONOK: 
		' ini����TIMEOUT�b�����擾
		'99/11/24 MOD START FOR MKK
		'    ZACN_TIME = GetPrivateProfileInt("RDO", "TIMEOUT", 3, ININAMESTR)
		
		ZACN_TIME = CInt(WG_TIMEOUT)
		'99/11/24 MOD END FOR MKK
		
		'�f�[�^�x�[�X�o�[�W�������擾
		Ret = SQLGetInfoString(ZACN_RCN.hdbc, 18, DRVNAME.Value, 128, DRVSTRNUM)
		If Ret <> SQL_SUCCESS Then
			DBVersion = 0
		Else
			DBVersion = Val(Left(DRVNAME.Value, InStr(DRVNAME.Value, Chr(0)) - 1))
		End If
		
		'�ڑ�����ZACN_USERID,ZACN_PASSWORD,ZACN_DBNAME�ɃZ�b�g
		CONSTR = ZACN_RCN.Connect
		If UNIFIEDLOGIN = False Then
			USERINFOKEY = "UID=" '���[�U��
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub ZACN_USERINFO
			ZACN_USERID = USERINFO
			USERINFOKEY = "PWD=" '�p�X���[�h
			'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
			GoSub ZACN_USERINFO
			ZACN_PASSWORD = USERINFO
		End If
		USERINFOKEY = "DSN=" '�f�[�^�\�[�X��
		'UPGRADE_ISSUE: GoSub �X�e�[�g�����g�̓T�|�[�g����Ă��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' ���N���b�N���Ă��������B
		GoSub ZACN_USERINFO
		ZACN_DBNAME = USERINFO
		
		If ZACN_DB = SQLSRV Then 'SQLServer���g�p
			'�f�[�^�x�[�X���ړ�
			Ret = GetPrivateProfileString("CONNECT", "USEDB", "", GETSTRWORK.Value, Len(GETSTRWORK.Value), ININAMESTR)
			If Ret > 0 Then
				'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
				'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
				'UPGRADE_ISSUE: LeftB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
				USEDB = StrConv(LeftB(StrConv(GETSTRWORK.Value, vbFromUnicode), Ret), vbUnicode)
				UDB = ZACN_RCN.CreateQuery("UDB", "USE " & USEDB)
				UDB.Execute()
				If Err.Number <> 0 Then
					UDB.Close()
					GoTo ZACN_ERR
				End If
				UDB.Close()
			End If
			
			'�����Z�L�����e�B�̏ꍇ�DMO�Ń��[�U�̐ڑ������擾
			'If UNIFIEDLOGIN Then
			'�c�l�n�ōĐڑ����邽�߁A�f�[�^�\�[�X���ڑ��T�[�o���擾
			'    Ret = SQLGetInfoString(ZACN_RCN.hdbc, SQL_SERVER_NAME, DRVNAME, 128, DRVSTRNUM)
			'    If Ret <> SQL_SUCCESS Then      '�擾�Ɏ��s
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
			'    '�c�l�n�ł̐ڑ�
			'    OServer.LoginTimeout = ZACN_TIME
			'    OServer.LoginSecure = True    '����è��߼��
			'    OServer.Connect ServerName:=SName, Login:=ZACN_USERID, Password:=ZACN_PASSWORD
			'    If Err.Number = 0 Then
			'        ZACN_USERID = OServer.TrueLogin
			'        ZACN_PASSWORD = OServer.Password
			'
			'        OServer.Disconnect              ' disconnect method
			'        If Err.Number <> 0 Then
			'            MsgBox Err.Number - vbObjectError & vbCr & Err.Description, vbCritical, "�f�[�^�x�[�X�ڑ��G���["
			'            GoTo ZACN_EXIT
			'        End If
			'    Else
			'        MsgBox Err.Number - vbObjectError & vbCr & Err.Description, vbCritical, "�f�[�^�x�[�X�ڑ��G���["
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
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ڑ��G���[")
		
		'�����Ă����G���[���N���A
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
		'UPGRADE_WARNING: Return �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
		Return 
	End Function
	
	
	'------------------------------------------------------------
	'�y�֐����z �f�B�X�R�l�N�g�T�u���[�`��
	'
	'�y�@  �\�z ODBC�ڑ������ް��ް���ؒf����B
	'
	'�y�d�l�z
	'        ODBC�ڑ������ް��ް���ؒf����B
	'        �ؒf�ɐ��������True��Ԃ��B
	'        ���s�����ꍇ�̓G���[���b�Z�[�W�\����AFalse��Ԃ��B
	'
	'
	'�y�߂�l�z Boolean�^
	'             True  :�ؒf����
	'            False  :�ؒf���s
	'
	'�y�֐��d�l�z
	'   Public Function ZADISCN_SUB() As Boolean
	'     �����n�p�����[�^��
	'        Public RCN As rdoConnection   '�ڑ����I�u�W�F�N�g
	'
	'�y�g�p��z
	'       Dim Ret As Boolean
	'       Ret = ZADISCN_SUB()
	'       End
	'------------------------------------------------------------
	Public Function ZADISCN_SUB() As Boolean
		'Public Function ZADISCN_SUB(Optional zRCN As Variant) As Boolean  '99/11/24 FOR MKK -> 99/12/09 DEL KTT YOSHINO
		
		On Error GoTo ZADISCN_ERR
		
		'99/12/09 ���� KTT YOSHINO ��
		ZACN_RCN.Close()
		ZADISCN_SUB = True
		'99/12/09 ���� KTT YOSHINO ��
		
		'------------------------------- 99/12/09 DEL START KTT YOSHINO ��
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
		'------------------------------- 99/12/09 DEL END KTT YOSHINO ��
		
		Exit Function
		
ZADISCN_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "�f�[�^�x�[�X�ؒf�G���[")
		
		'�����Ă����G���[���N���A
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
		
		ERR_MSG = "�f�[�^�x�[�X�ւ̐ڑ��Ɏ��s���܂����B"
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & vbCr & RdoErr.Description & ":" & RdoErr.Number
		Next RdoErr
		
		If MsgBox(ERR_MSG, MsgBoxStyle.RetryCancel + MsgBoxStyle.Exclamation, "�f�[�^�x�[�X�ڑ��G���[") = MsgBoxResult.Retry Then
			ZACN_ERR_SUB = True
		Else
			ZACN_ERR_SUB = False
		End If
		
		'�����Ă����G���[���N���A
		RDOrdoEngine_definst.rdoErrors.Clear()
	End Function
	
	
	'ZACN_DBNAME�̃f�[�^�\�[�X�̃f�[�^�x�[�X��ʂ�ZACN_DB�ɃZ�b�g����
	Public Function ZACN_GETDBTYPE() As Boolean
		Dim Ret As Short
		Dim sDSNItem As New VB6.FixedLengthString(1024)
		Dim sDRVItem As New VB6.FixedLengthString(1024)
		Dim sDSN As String
		Dim sDRV As String
		Dim iDSNLen As Short
		Dim iDRVLen As Short
		Dim lHenv As Integer '�������
		
		ZACN_GETDBTYPE = False
		
		'�ް���������擾���܂��B
		If SQLAllocEnv(lHenv) <> -1 Then
			Ret = SQL_SUCCESS
			Do 
				sDSNItem.Value = Space(1024)
				sDRVItem.Value = Space(1024)
				Ret = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem.Value, 1024, iDSNLen, sDRVItem.Value, 1024, iDRVLen)
				If Ret <> SQL_SUCCESS Then Exit Do
				
				'�f�[�^�\�[�X���擾
				If UCase(Left(sDSNItem.Value, iDSNLen)) = UCase(ZACN_DBNAME) Then
					'�^�[�Q�b�g�̃f�[�^�\�[�X�������̂ŁA�h���C�o�̋L�q�q���擾
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
		Dim lpValueName As String '�l�̖��O
		Dim lpData As String '�f�[�^
		Dim lpcbData As Integer '�f�[�^�̒���
		Dim lpType As Integer '�f�[�^�̃^�C�v
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
			
			'Trusted_Connection���擾
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