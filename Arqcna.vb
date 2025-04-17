Option Strict Off
Option Explicit On
Module ARQCNABAS
	'******************************************************************
	'*    システム名    ：  ＳＭＩＬＥαＶｅｒ５共通                  *
	'*    サブルーチン名：  別セッションコネクトサブルーチン          *
	'*    作  成  者    ：  ＳＯＦＴＥＣ－渡部                        *
	'******************************************************************
	
	Public ZACNA_RCN As RDO.rdoConnection 'データベース接続情報
	
	' **************************************************************
	'   接続処理
	' **************************************************************
	Public Function ZACNA_SUB() As Short
		
		RDOrdoEngine_definst.rdoDefaultCursorDriver = RDO.CursorDriverConstants.rdUseIfNeeded
		
		'別セッション接続
		On Error Resume Next
		Err.Clear()
		ZACNA_RCN = RdoEnv.OpenConnection("", RDO.PromptConstants.rdDriverNoPrompt, False, ZACNA_CONSTR(ZACN_DBNAME, ZACN_USERID, ZACN_PASSWORD))
		If Err.Number <> n0 Then
			'接続失敗
			GoTo ZACNA_EXIT
		End If
		
ZACNA_CONOK: 
		ZACNA_SUB = True
		Exit Function
		
ZACNA_EXIT: 
		Err.Clear()
		If Not (ZACNA_RCN Is Nothing) Then ZACNA_RCN.Close()
		If Not (RdoEnv Is Nothing) Then RdoEnv.Close()
		
		ZACNA_SUB = False
		Exit Function
		
	End Function
	' **************************************************************
	'   接続処理
	' **************************************************************
	Private Function ZACNA_CONSTR(ByRef DSN As String, ByRef USER As String, ByRef PASSW As String) As String
		
		ZACNA_CONSTR = "DSN=" & DSN & ";UID=" & USER & ";PWD=" & PASSW & ";"
		
	End Function
	
	' **************************************************************
	'   切断処理
	' **************************************************************
	Public Function ZADISCNA_SUB() As Boolean
		
		On Error GoTo ZADISCNA_ERR
		ZACNA_RCN.Close()
		ZADISCNA_SUB = True
		Exit Function
		
ZADISCNA_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "データベース切断エラー")
		
		'入っていたエラーをクリア
		RDOrdoEngine_definst.rdoErrors.Clear()
		
		ZADISCNA_SUB = False
		
	End Function
End Module