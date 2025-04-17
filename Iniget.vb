Option Strict Off
Option Explicit On
Module INIGETBAS
	'*
	'* MKK �d���݌ɊǗ��V�X�e���p  INI�擾�T�u���[�`��
	'*
	'* 1999/11/24 KTT-YOSHINO
	'*
	
	
	'Windows 95 VB Ver4.0 API
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	
	'*-------- �O���[�o���ϐ� --------*
	
	'�f�[�^�x�[�X�ڑ����
	'Public WG_DBLINK          As String     '0:�g�p����  1:�g�p���Ȃ�  '99/12/09 DEL KTT YOSHINO
	
	Public WG_UAID As String '[DATABASE] ���㔄�|DBհ�ޖ�
	Public WG_UAPW As String '[DATABASE] ���㔄�|DB�߽ܰ��
	Public WG_UADSN As String '[DATABASE] ���㔄�|DBDSN��
	
	Public WG_SZID As String '[DATABASE] �d���݌�DBհ�ޖ�
	Public WG_SZPW As String '[DATABASE] �d���݌�DB�߽ܰ��
	Public WG_SZDSN As String '[DATABASE] �d���݌�DBDSN��
	
	Public WG_GCID As String '[DATABASE] �Ɩ��ԋ���DBհ�ޖ�
	Public WG_GCPW As String '[DATABASE] �Ɩ��ԋ���DB�߽ܰ��
	Public WG_GCDSN As String '[DATABASE] �Ɩ��ԋ���DBDSN��
	
	'���ʏ��
	Public WG_TIMEOUT As String '[RDO] ��ѱ�Ď���
	Public WG_REQUERY As String '[RDO] REQUERY�L��
	
	Public WG_OPCODE As String '[OPERATOR] ���ڰ�����
	Public WG_INCCODE As String '[OPERATOR] ��к���
	Public WG_JGCODE As String '[OPERATOR] ���Ə�����
	Public WG_BUSYOCODE As String '[OPERATOR] ��������
	
	Public WG_DEBUG As String '[MAIN] DEBUG  "1"�̏ꍇ��OPCODE��L���Ƃ���
	
	Public WG_EXCELPATH As String '[SZPRG] EXCEĻ�� PATH
	
	Public WG_FAXID As String '[FAX] ۸޵�ID
	Public WG_FAXPW As String '[FAX] �߽ܰ��
	
	
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const KEY_QUERY_VALUE As Short = &H1s
	Public Const ERROR_SUCCESS As Short = 0
	Public Const ERROR_FILE_NOT_FOUND As Short = 2
	
	'+------------------------------------------------------------------+
	'| �I���W�i�������ݒ�t�@�C�����                                     |
	'+------------------------------------------------------------------+
	Public OrgIniPathName As String '�p�X
	Public OrgIniFileName As String '�t�@�C����
	'
	Public Sub REGGET()
		
		'���W�X�g������h�m�h�t�@�C���̃p�X�A�t�@�C�������擾����
		
		Dim Ret As Integer 'Return Code
		Dim hKey As Integer 'Key Handle
		Dim lpSubKey As String 'Sub Key
		Dim phkResult As Integer 'Open Key Handle
		Dim lpValueName As String '�l�̖��O
		Dim lpData As String '�f�[�^
		Dim lpcbData As Integer '�f�[�^�̒���
		Dim lpType As Integer '�f�[�^�̃^�C�v
		
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
		
		'�I���W�i�������ݒ�t�@�C���̃p�X���擾
		lpValueName = ValuePathName & Chr(0)
		lpData = Space(256)
		lpcbData = 256
		Ret = RegQueryValueEx(phkResult, lpValueName, 0, lpType, lpData, lpcbData)
		If Ret <> ERROR_SUCCESS Then
			'        MsgBox "RegQueryValueEx Error!!  Code = " & ret
		Else
			OrgIniPathName = Left(lpData, InStr(lpData, Chr(0)) - 1)
		End If
		
		'�I���W�i�������ݒ�t�@�C�������擾
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
		'   /* INI�t�@�C���̓��e�擾�i�ʁj */
		'   /*     (Internal Function)       */
		'   /*                               */
		
		Static bUF As New VB6.FixedLengthString(256)
		Dim buftmp As String
		
		bUF.Value = ""
		
		'   INI�t�@�C���̎w��G���g�����擾
		If (GetPrivateProfileString(section, entry, def_str, bUF.Value, 256, fname) > 0) Then
			buftmp = Trim(bUF.Value)
		Else
			buftmp = Trim(def_str)
		End If
		
		'   ������̍Ō�� '\0'���t���Ă����Ȃ珜������
		'UPGRADE_ISSUE: RightB$ �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		If (RightB$(buftmp, 2) = Chr(0)) Then
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LeftB$ �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
			INIGET_ENTRY = LeftB$(buftmp, LenB(buftmp) - 2)
		Else
			INIGET_ENTRY = buftmp
		End If
	End Function
	
	Sub INIGET_SUB(ByVal fname As String)
		
		'   �p  �� : INIGET_L_SUB("MKK.INI")
		'
		'*******************************************************************************
		'*     �q�d�f�d�c�h�s�D�d�w�d�ɂĉ��L����ǉ����Ă������ƁB                      *
		'*     -HKEY_LOCAL_MACHINE\SOFTWAREMKK���쐬���A���̉���                        *
		'*     IniPath   :INI̧�ق��߽                                                 *
		'*     IniFile   :INI̧�ٖ��@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*
		'*******************************************************************************
		
		Dim INI_NAME As String
		'  If App.PrevInstance Then End
		
		'���W�X�g������t�@�C�������擾����
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
		
		'�h�m�h�t�@�C�����獀�ڎ擾
		
		'�f�[�^�x�[�X�ڑ����
		'    WG_DBLINK = INIGET_ENTRY("DATABASE", "DBLINK", "", INI_NAME)  '[DATABASE] DBLINK  99/12/09 DEL KTT YOSHINO
		
		WG_UAID = INIGET_ENTRY("DATABASE", "UAID", "", INI_NAME) '[DATABASE] ���㔄�|DBհ�ޖ�
		WG_UAPW = INIGET_ENTRY("DATABASE", "UAPW", "", INI_NAME) '[DATABASE] ���㔄�|DB�߽ܰ��
		WG_UADSN = INIGET_ENTRY("DATABASE", "UADSN", "", INI_NAME) '[DATABASE] ���㔄�|DBDSN��
		
		WG_SZID = INIGET_ENTRY("DATABASE", "SZID", "", INI_NAME) '[DATABASE] �d���݌�DBհ�ޖ�
		WG_SZPW = INIGET_ENTRY("DATABASE", "SZPW", "", INI_NAME) '[DATABASE] �d���݌�DB�߽ܰ��
		WG_SZDSN = INIGET_ENTRY("DATABASE", "SZDSN", "", INI_NAME) '[DATABASE] �d���݌�DBDSN��
		
		WG_GCID = INIGET_ENTRY("DATABASE", "GCID", "", INI_NAME) '[DATABASE] �Ɩ��ԋ���DBհ�ޖ�
		WG_GCPW = INIGET_ENTRY("DATABASE", "GCPW", "", INI_NAME) '[DATABASE] �Ɩ��ԋ���DB�߽ܰ��
		WG_GCDSN = INIGET_ENTRY("DATABASE", "GCDSN", "", INI_NAME) '[DATABASE] �Ɩ��ԋ���DBDSN��
		
		'���ʏ��
		WG_TIMEOUT = INIGET_ENTRY("RDO", "TIMEOUT", "", INI_NAME) '[RDO] ��ѱ�Ď���
		WG_REQUERY = INIGET_ENTRY("RDO", "REQUERY", "", INI_NAME) '[RDO] REQUERY�L��
		
		WG_OPCODE = INIGET_ENTRY("OPERATOR", "OPCODE", "", INI_NAME) '[OPERATOR] ���ڰ�����
		WG_INCCODE = INIGET_ENTRY("OPERATOR", "INCCODE", "", INI_NAME) '[OPERATOR] ��к���
		WG_JGCODE = INIGET_ENTRY("OPERATOR", "JGCODE", "", INI_NAME) '[OPERATOR] ���Ə�����
		WG_BUSYOCODE = INIGET_ENTRY("OPERATOR", "BUSYOCODE", "", INI_NAME) '[OPERATOR] ��������
		
		WG_DEBUG = INIGET_ENTRY("MAIN", "DEBUG", "", INI_NAME) '[MAIN] DEBUG  "1"�̏ꍇ��OPCODE��L���Ƃ���
		
		WG_EXCELPATH = INIGET_ENTRY("SZPRG", "EXCELPATH", "", INI_NAME) '[SZPRG] EXCEĻ�� PATH
		
		WG_FAXID = INIGET_ENTRY("FAX", "FAXID", "", INI_NAME) '[FAX] ۸޵�ID
		WG_FAXPW = INIGET_ENTRY("FAX", "FAXPW", "", INI_NAME) '[FAX] �߽ܰ��
		
	End Sub
End Module