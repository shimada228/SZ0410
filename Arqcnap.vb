Option Strict Off
Option Explicit On
Module ARQCNAPBAS
	
	'99/12/09 DEL START KTT YOSHINO
	'Public ZACN_UARCN As rdoConnection        '接続情報オブジェクト 売上売掛ＤＢ
	'Public ZACN_SZRCN As rdoConnection        '接続情報オブジェクト 仕入在庫ＤＢ
	'Public ZACN_GCRCN As rdoConnection        '接続情報オブジェクト 業務間共通ＤＢ
	'99/12/09 DEL END   KTT YOSHINO
	
	'------------------------------------------------------------
	'【関数名】 コネクトサブルーチン ＡＰ用
	'
	'【機  能】 既存のコネクトサブルーチンを利用し
	'          ｵﾍﾟﾚｰﾀﾏｽﾀ、サーバーマスタ、ｻｰﾊﾞｰﾏｽﾀにて指定されたﾃﾞｰﾀﾍﾞｰｽへの接続を行う
	'
	'【戻り値】 Boolean型
	'             True  :接続成功
	'            False  :接続失敗
	'
	'【関数仕様】
	'   Public Function ZACNAP_SUB(Company as string,Operator as string) As Boolean
	'     ＜プロシージャ引数＞
	'
	'     ＜結果引渡パラメータ＞
	'        ZACNAP_SUB と同一の内容です
	'
	'        Public ZACN_UARCN As rdoConnection   '共通ＤＢ接続情報オブジェクト
	'        Public ZACN_UKRCN As rdoConnection   '売上売掛ＤＢ接続情報オブジェクト
	'        Public ZACN_SZRCN As rdoConnection   '仕入在庫ＤＢ接続情報オブジェクト
	'
	'        Public ZACN_DB As Integer            '使用データベース    ORCL:ORACLE
	'                                                               SQLSRV:SQLServer
	'        Public ZACN_TIME As Long           'RDO命令のｳｪｲﾄ時間
	'        Public ZACN_USERID As String       '接続したユーザ名
	'        Public ZACN_PASSWORD As String     '接続したパスワード
	'        Public ZACN_DBNAME As String       '接続したデータソース名
	'
	'------------------------------------------------------------
	'Public Function ZACNAP_SUB(Company As String, Operator As String) As Boolean
	'Public Function ZACNAP_SUB() As Boolean                                               '99/12/09 DEL KTT YOSHINO
	Public Function ZACNAP_SUB(ByRef ZAID As String, ByRef ZAPW As String, ByRef ZADSN As String) As Boolean '99/12/09 ADD KTT YOSHINO
		
		Dim StrPos As Short
		Dim StrPos2 As Short
		Dim USERINFOKEY As String
		Dim USERINFO As String
		Dim CONSTR As String
		Dim Ret As Integer
		Dim SQL_STR As String 'SQL文字列格納
		
		Dim SMARS As RDO.rdoResultset
		Dim SMBRS As RDO.rdoResultset
		Dim SMCRS As RDO.rdoResultset
		
		Dim Post As String
		Dim Server As String
		
		'99/12/09 ADD START KTT YOSHINO 引数指定のＤＢに接続
		'UPGRADE_NOTE: オブジェクト ZACN_RCN をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		ZACN_RCN = Nothing
		If Not ZACN_SUB(False, ZAID, ZAPW, ZADSN) Then
			ZACNAP_SUB = False
			Exit Function
		End If
		'99/12/09 END START KTT YOSHINO 引数指定のＤＢに接続
		
		'------------------------------ 99/12/09 DEL START KTT YOSHINO
		'    '仕入在庫サーバーへ接続
		'    Set ZACN_RCN = Nothing
		'    If Not ZACN_SUB(False, WG_SZID, WG_SZPW, WG_SZDSN) Then
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'
		'    Set ZACN_SZRCN = ZACN_RCN       '接続情報オブジェクト
		'
		'    'DATABASELINKを使用する場合(WG_DBLINK=1)は仕入のコネクションを売上、共通に設定して終了
		'    If WG_DBLINK = "1" Then
		'
		'        Set ZACN_UARCN = ZACN_SZRCN       '接続情報オブジェクト
		'        Set ZACN_GCRCN = ZACN_SZRCN       '接続情報オブジェクト
		'
		'        Set ZACN_RCN = Nothing
		'
		'        On Error GoTo 0
		'        ZACNAP_SUB = True
		'        Exit Function
		'    End If
		'
		'    '売掛売上サーバーへ接続
		'    If Not ZACN_SUB(False, WG_UAID, WG_UAPW, WG_UADSN) Then
		'
		'        '仕入サーバーの接続を解除
		'        Ret = ZADISCN_SUB(ZACN_SZRCN)
		'
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'    '売上ﾃﾞｰﾀﾍﾞｰｽの接続情報セーブ
		'    Set ZACN_UARCN = ZACN_RCN       '接続情報オブジェクト
		'
		'
		'    '業務間共通サーバーへ接続
		'    Set ZACN_RCN = Nothing
		'    If Not ZACN_SUB(False, WG_GCID, WG_GCPW, WG_GCDSN) Then
		'        ZACNAP_SUB = False
		'
		'        '仕入在庫サーバーの接続を解除
		'        Ret = ZADISCN_SUB(ZACN_SZRCN)
		'
		'        '売上売掛サーバーの接続を解除
		'        Ret = ZADISCN_SUB(ZACN_UARCN)
		'
		'        ZACNAP_SUB = False
		'        Exit Function
		'    End If
		'    Set ZACN_GCRCN = ZACN_RCN       '接続情報オブジェクト
		'
		'    '作業用コネクションの初期化
		'    Set ZACN_RCN = Nothing
		'------------------------------ 99/12/09 DEL END KTT YOSHINO
		
		On Error GoTo 0
		
		ZACNAP_SUB = True
		Exit Function
		
ZACNAP_ERR: 
		Dim ERR_MSG As String
		Dim RdoErr As RDO.rdoError
		
		For	Each RdoErr In RDOrdoEngine_definst.rdoErrors
			ERR_MSG = ERR_MSG & RdoErr.Description & ":" & RdoErr.Number & vbCr
		Next RdoErr
		
		MsgBox(ERR_MSG, MsgBoxStyle.Critical, "データベース接続エラー")
		
		'入っていたエラーをクリア
		RDOrdoEngine_definst.rdoErrors.Clear()
		
		Err.Clear()
		If Not (ZACN_RCN Is Nothing) Then ZACN_RCN.Close()
		If Not (RdoEnv Is Nothing) Then RdoEnv.Close()
		
		On Error GoTo 0
		
		ZACN_USERID = ""
		ZACN_PASSWORD = ""
		ZACN_DBNAME = ""
		
		ZACNAP_SUB = False
		Exit Function
		
		'ZACNAP_USERINFO:
		'    USERINFO = ""
		'    StrPos = InStr(CONSTR, USERINFOKEY)
		'    If StrPos > 0 Then
		'        StrPos = StrPos + 4
		'        StrPos2 = InStr(StrPos, CONSTR, ";")
		'        If StrPos2 > 0 Then
		'            USERINFO = Mid(CONSTR, StrPos, StrPos2 - StrPos)
		'        End If
		'    End If
		'    Return
		
	End Function
	'Private Function ZACNAP_CONSTR(DSN As String, USER As String, PASSW As String) As String
	'    ZACNAP_CONSTR = "DSN=" & DSN & ";UID=" & USER & ";PWD=" & PASSW & ";"
	'End Function
End Module