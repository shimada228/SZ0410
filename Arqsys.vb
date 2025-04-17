Option Strict Off
Option Explicit On
Module ARQSYSBAS
	
	
	'******************************************************************
	'*      システム日付（サーバー日付）取得      for MKK(仕入)
	'******************************************************************
	'*【効能】
	'*      Oracleサーバーより現在の日付を取込みます
	'*　　  また、更新時の更新日時の設定を簡便にします
	'*【引数】
	'*      1.日付を受け取るための文字列型の変数
	'*      2.使用目的を示すフラグ: 0～4 省略不可
	'*      -----------------------------------------------------------
	'*      結果
	'*      2のフラグが0  YYYYMMDD
	'*      2のフラグが1  YYYY/MM/DD
	'*      2のフラグが2  YYYY-MM-DD HH:MI:SS   (未使用)
	'*      2のフラグが3  YYYYMMDDHHMISS        (ＤＢ更新用）
	'*      2のフラグが4  YYYY/MM/DD  HH:MI:SS  (帳票出力用）
	'*【戻り値】
	'*      成功時:True 失敗時:False
	'******************************************************************
	Function ZASYS_SUB(ByRef SYSYMD As String, ByRef FLG As Short) As Boolean
		
		Dim Server As RDO.rdoResultset '結果セット
		
		ZASYS_SUB = False
		
		'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
		If IsNothing(FLG) Then FLG = False
		
		SQL = "select to_char(sysdate,'YYYY-MM-DD HH24:MI:SS') SYMD,to_char(sysdate,'YYYYMMDDHH24MISS') SYMD2,to_char(sysdate,'YYYY/MM/DD  HH24:MI:SS') SYMD3,to_char(sysdate,'YYYYMMDD') AS YMD,to_char(sysdate,'YYYY/MM/DD') AS YMD2 from dual"
		On Error Resume Next
		'    Set Server = ZACN_SZRCN.OpenResultset(SQL)     '99/12/09 DEL KTT YOSHINO
		Server = ZACN_RCN.OpenResultset(SQL) '99/12/09 ADD KTT YOSHINO
		Select Case B_STATUS(Server)
			Case 0
				If FLG = 4 Then
					SYSYMD = Server.rdoColumns("SYMD3").Value
				ElseIf FLG = 3 Then 
					SYSYMD = Server.rdoColumns("SYMD2").Value
				ElseIf FLG = 2 Then 
					SYSYMD = Server.rdoColumns("SYMD").Value
				ElseIf FLG = 1 Then 
					SYSYMD = Server.rdoColumns("YMD2").Value
				Else
					SYSYMD = Server.rdoColumns("YMD").Value
				End If
				ZASYS_SUB = True
			Case Else
				ERRSW = F_ERR
				ZAER_KN = 1
				ZAER_NO.Value = ""
				ZAER_MS.Value = "ARQSYS内エラー"
				Call ZAER_SUB()
		End Select
		On Error GoTo 0
		
	End Function
End Module