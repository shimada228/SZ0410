Option Strict Off
Option Explicit On
Module ARQLGBAS
	
	Public ZALG_KBN As New VB6.FixedLengthString(1)
	Public ZALG_NAIYO As String
	
	
	Public ZALG_ERR As New VB6.FixedLengthString(1)
	
	
	Public Sub ZALG_SUB()
		Dim Ret As Short
		
		Dim LOGDIRNAME As String '���O�o�̓f�B���N�g����
		Dim LOGFNAME As String '�p�X���܂ރ��O�t�@�C����
		'
		Dim SYSDATE As New VB6.FixedLengthString(8) '�V�X�e�����t
		Dim SYSTIME As New VB6.FixedLengthString(6) '�V�X�e������
		Dim PRGID As String '�v���O�����h�c
		
		Dim OUTFNum As Short '�o�̓t�@�C���̃t�@�C���ԍ�
		Dim OUT_REC As String '�o�̓t�@�C�����C�A�E�g
		
		'   << �G���[�t���O�̃N���A >>
		ZALG_ERR.Value = "0"
		
		'   << �V�X�e�����t�A���ԁA�v���O�����h�c��荞�� >>
		SYSDATE.Value = VB6.Format(Now, "YYYYMMDD")
		SYSTIME.Value = VB6.Format(Now, "HHMMSS")
		'UPGRADE_WARNING: App �v���p�e�B App.EXEName �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
		PRGID = My.Application.Info.AssemblyName
		
		'   << Smile.ini���A���O�o�̓f�B���N�g���̎擾 >>
		Ret = MKKCMN.ZAGI_SUB("LOG", "LOGFNAME", "", LOGDIRNAME, "SMILE.INI")
		If Ret = False Then
			GoTo ZALG_END
		End If
		
		'   << �o�̓t�@�C���������i�p�X�܂݁j >>
		If Len(LOGDIRNAME) <> 0 Then
			If Mid(LOGDIRNAME, Len(LOGDIRNAME) - 1, 1) = "\" Then
				LOGFNAME = LOGDIRNAME & "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
			Else
				LOGFNAME = LOGDIRNAME & "\" & "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
			End If
		Else
			LOGFNAME = "SMIL" & Mid(SYSDATE.Value, 1, 4) & ".LOG"
		End If
		
		'   << �o�̓t�@�C���̃t�@�C�����`�F�b�N >>
		Ret = MKKCMN.ZAPC_SUB(LOGFNAME)
		If Ret <> 0 And Ret <> -1 Then
			GoTo ZALG_END
		End If
		
		'****** << ���O�t�@�C���o�͏��� >> *****
		
		'   <<  �I�[�v�� >>
ZALG_0010: 
		OUTFNum = FreeFile
		On Error Resume Next
		Err.Clear()
		FileOpen(OUTFNum, LOGFNAME, OpenMode.Append, , OpenShare.LockReadWrite)
		Select Case Err.Number
			Case 0 ' ����.
				
			Case 53, 75, 76 ' �p�X���s���A�܂��͌�����Ȃ�����
				GoTo ZALG_END
			Case 52, 64 ' �t�@�C����������
				GoTo ZALG_END
			Case 70 ' �������ݕs�\���t�@�C���g�p��
				GoTo ZALG_0010
			Case 68, 71 ' �f�o�C�X/�h���C�u�̏������ł��Ă��Ȃ�
				GoTo ZALG_END
			Case Else ' �A�N�Z�X�s�\
				GoTo ZALG_END
		End Select
		On Error GoTo 0
		
		'   << �������ݏ��� >>
		OUT_REC = SYSDATE.Value & "," & SYSTIME.Value & "," & PRGID & ","
		'
		Select Case ZALG_KBN.Value
			Case "0"
				OUT_REC = OUT_REC & "�J�n," & ZALG_NAIYO
			Case "1"
				OUT_REC = OUT_REC & "�I��," & ZALG_NAIYO
			Case "9"
				OUT_REC = OUT_REC & "�G���[," & ZALG_NAIYO
		End Select
		
		On Error Resume Next
		PrintLine(OUTFNum, OUT_REC)
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("���O�t�@�C���̏������݂Ɏ��s���܂����B" & Err.Number, 48, "")
				ZALG_ERR.Value = "1"
		End Select
		On Error GoTo 0
		
		'   <<  �b�k�n�r�d >>
		On Error Resume Next
		FileClose(OUTFNum)
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("���O�t�@�C���̏������݂Ɏ��s���܂����B" & Err.Number, 48, "")
				ZALG_ERR.Value = "1"
		End Select
		On Error GoTo 0
		
		
ZALG_END: 
		
	End Sub
End Module