Option Strict Off
Option Explicit On
Module STOREDV5BAS
	Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc As Integer, ByRef phstmt As Integer) As Short
	Declare Function SQLError Lib "odbc32.dll" (ByVal henv As Integer, ByVal hdbc As Integer, ByVal hstmt As Integer, ByVal szSqlState As String, ByRef pfNativeError As Integer, ByVal szErrorMsg As String, ByVal cbErrorMsgMax As Short, ByRef pcbErrorMsg As Short) As Short
	Declare Function SQLExecDirect Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal szSqlStr As String, ByVal cbSqlStr As Integer) As Short
	Declare Function SQLExecute Lib "odbc32.dll" (ByVal hstmt As Integer) As Short
	Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal fOption As Short) As Short
	Declare Function SQLPrepare Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal szSqlStr As String, ByVal cbSqlStr As Integer) As Short
	Declare Function SQLRowCount Lib "odbc32.dll" (ByVal hstmt As Integer, ByRef pcrow As Integer) As Short
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function SQLSetParam Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal ipar As Short, ByVal fCType As Short, ByVal fSqlType As Short, ByVal cbColDef As Integer, ByVal ibScale As Short, ByRef rgbValue As Any, ByRef pcbValue As Integer) As Short
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function SQLBindParameter Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal ipar As Short, ByVal fParamType As Short, ByVal fCType As Short, ByVal fSqlType As Short, ByVal cbColDef As Integer, ByVal ibScale As Short, ByRef rgbValue As Any, ByVal cbValueMax As Integer, ByRef pcbValue As Integer) As Short
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt As Integer, ByRef szTblQualifier As Any, ByVal cbTblQualifier As Short, ByRef szTblOwner As Any, ByVal cbTblOwner As Short, ByRef szTblName As Any, ByVal cbTblName As Short, ByRef szColName As Any, ByVal cbColName As Short) As Short
	Declare Function SQLSetStmtOption Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal fOption As Short, ByVal vParam As Integer) As Short
	Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hdbc As Integer, ByVal szConnStrIn As String, ByVal cbConnStrIn As Short, ByVal szConnStrOut As String, ByVal cbConnStrOutMax As Short, ByRef pcbConnStrOut As Short) As Short
	Declare Function SQLParamOptions Lib "odbc32.dll" (ByVal hstmt As Integer, ByVal crow As Short, ByRef pirow As Integer) As Short
	Public Const SQL_NTS As Integer = -3 '  NTS = Null Terminated String
	Public Const SQL_ERROR As Integer = -1
	Public Const SQL_INVALID_HANDLE As Integer = -2
	Public Const SQL_NO_DATA_FOUND As Integer = 100
	Public Const SQL_SUCCESS As Integer = 0
	Public Const SQL_SUCCESS_WITH_INFO As Integer = 1
	Public Const SQL_C_DEFAULT As Integer = 99
	Public Const SQL_COMMIT As Integer = 0
	Public Const SQL_ROLLBACK As Integer = 1
	Public Const SQL_CHAR As Integer = 1
	Public Const SQL_NUMERIC As Integer = 2
	Public Const SQL_DECIMAL As Integer = 3
	Public Const SQL_INTEGER As Integer = 4
	Public Const SQL_DOUBLE As Integer = 8
	Public Const SQL_VARCHAR As Integer = 12
	Public Const SQL_C_CHAR As Integer = SQL_CHAR '  CHAR, VARCHAR, DECIMAL, NUMERIC
	Public Const SQL_C_DOUBLE As Integer = SQL_DOUBLE '  FLOAT, DOUBLE
	Public Const SQL_PARAM_INPUT As Integer = 1
	Public Const SQL_PARAM_OUTPUT As Integer = 4
	Public Const SQL_DROP As Integer = 1
	
	
	Public Function bytestostring(ByRef byte_array() As Byte) As String
		Dim data As String
		Dim StrLen As String
		'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		data = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(byte_array), vbUnicode)
		StrLen = CStr(InStr(data, Chr(0)) - 1)
		bytestostring = Left(data, CInt(StrLen))
	End Function
	
	Public Sub DescribeError(ByVal phenv As Integer, ByVal hdbc As Integer, ByVal hstmt As Integer)
		' Print an error message for the given connection handle
		' and statement handle
		Dim rgbValue1 As New VB6.FixedLengthString(16)
		Dim rgbValue3 As New VB6.FixedLengthString(256)
		Dim Outlen As Short
		Dim Native As Integer
		Dim rc As Short
		Dim retcd As Short
		
		rgbValue1.Value = New String(Chr(0), 16)
		rgbValue3.Value = New String(Chr(0), 256)
		Do 
			rc = SQLError(phenv, hdbc, hstmt, rgbValue1.Value, Native, rgbValue3.Value, 256, Outlen)
			'Screen.MousePointer = Normal
			If rc = SQL_SUCCESS Or rc = SQL_SUCCESS_WITH_INFO Then
				If Outlen = 0 Then
					MsgBox("Error -- No error information available")
				Else
					If rc = SQL_ERROR Then
						MsgBox(Left(rgbValue3.Value, Outlen))
					Else
						MsgBox(Left(rgbValue3.Value, Outlen))
					End If
				End If
			End If
		Loop Until rc <> SQL_SUCCESS
	End Sub
	
	Public Sub strtobyte(ByRef data As String, ByRef bytelen As Short, ByRef return_buff() As Byte)
		Dim N As Object
		' convert string to byte array.
		Dim StrLen, Count As Short
		Dim tmpbyte() As Byte
		Dim tmpstr As String
		
		'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		tmpstr = StrConv(data, vbFromUnicode)
		'UPGRADE_ISSUE: LenB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
		bytelen = LenB(tmpstr)
		'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		'UPGRADE_TODO: System.Text.UnicodeEncoding.Unicode.GetBytes() を使うためにコードがアップグレードされましたが、動作が異なる可能性があります。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"' をクリックしてください。
		tmpbyte = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(data, vbFromUnicode))
		For N = 0 To bytelen - 1
			'UPGRADE_WARNING: オブジェクト N の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			return_buff(N) = tmpbyte(N)
		Next N
	End Sub
End Module