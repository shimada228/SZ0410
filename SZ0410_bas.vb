Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module SZ0410BAS
	'******************************************************************
	'*  �V�X�e����    �F  �l�j�j  �d���݌ɊǗ��V�X�e��                *
	'*  �v���O������  �F  �d���i�ڊ�{������      �@�@              *
	'*  �v���O�����h�c�F  �r�y�O�S�P�O                                *
	'*  ��  ��  ��   �F               �@�@�@�@�@�@                    *
	'******************************************************************
	'*  ���{�C������
	'*
	'* 2001/01/23 SOFIX M.MASUYA
	'* �@ �폜���s���ɊY���i�ڂ̎��т𒲂ׂĎ��т�������̂̍폜��s�Ƃ���
	'******************************************************************
	'*  �v���O������        �F�d���i�ڊ�{������                    *
	'*  �v���O����ID        �FSZ0410.EXE                              *
	'*  �R���p�C����        �F2007/06/11                              *
	'*  �ύX�L�[            �FCUST-20070611                           *
	'*  �ύX�S����          �F�r�r�o�|���                            *
	'*  �C�����e            �F�������X�V�����Ή�                      *
	'******************************************************************'A-CUST-20070611
	'*  �R���p�C����        �F2010/06/24                              *
	'*  �ύX�L�[            �FCUST-20100610                           *
	'*  �ύX�S����          �F�r�r�o�|��R��                          *
	'*  �C�����e            �F�@�������̂̒ǉ��B                      *
	'*                        �A�i�Ԃ̍̔ԕ��@��ύX�B                *
	'*                        �B�o�^���͕i�Ԃ̓��͂�s�Ƃ���B      *
	'*                        �C�i�ڏ��b�r�u�捞�@�\�̒ǉ��B        *
	'*                        �D�捞�f�[�^�̑I���@�\�̒ǉ��B          *
	'*                        �E�G���[���b�Z�[�W�\���ɂ���Ȃ��t�@�C��*
	'*                          �����\�������̂��C���B              *
	'*                        �F�e�`�w���M�˃��[�����M�Ɍ��o����ύX  *
	'******************************************************************'A-CUST-20100610
	'*  �R���p�C����        �F2010/08/23                              *
	'*  �ύX�L�[            �FCUST-20100823                           *
	'*  �ύX�S����          �F�r�r�o�|��R��                          *
	'*  �C�����e            �F�@�b�r�u�捞�ɓK�p���A�����P�ʁA���Z���A*
	'*                          JAN�W���R�[�h�AJAN�Z�k�R�[�h�A���̑��o*
	'*                          �[�R�[�h��ǉ�����B                  *
	'*                        �A�P�ʂɒP�ʃ}�X�^�̃`�F�b�N���s���B    *
	'*                        �B�捞�ǉ����ڂ��ꗗ�ɒǉ�����B        *
	'******************************************************************'A-CUST-20100823
	'*  �R���p�C����        �F2010/09/01                              *
	'*  �ύX�L�[            �FCUST-20100901                           *
	'*  �ύX�S����          �F�r�r�o�|��R��                          *
	'*  �C�����e            �F�@�K�C�h�̂�����C���B                  *
	'*                        �A�b�r�u�捞�̊m�F���b�Z�[�W�̃A�C�R����*
	'*                          �C���B                                *
	'*                        �B�捞�E�I����ʂ̋N���`�F�b�N��ύX�B  *
	'******************************************************************'A-CUST-20100901
	'*  �R���p�C����        �F2010/06/21
	'*  �ύX�L�[            �F20110621
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F�i�ڑI���ɍ폜�@�\��ǉ��B�t�@���N�V�����@�\
	'******************************************************************'A-20110621-
	'*  �R���p�C����        �F2013/02/20
	'*  �ύX�L�[            �F20130212
	'*  �ύX�S����          �F�r�r�o�|�N��
	'*  �C�����e            �FJAN�֘A���ڂ�ǉ�
	'******************************************************************'A-20130212-
	'*  �R���p�C����        �F2013/02/22
	'*  �ύX�L�[            �F20130222
	'*  �ύX�S����          �F�r�r�o�|�N��
	'*  �C�����e            �FF7:�i�ڑI����JAN�֘A���ڂɑΉ�������
	'******************************************************************'A-20130212-
	'*  �R���p�C����        �F2013/02/27
	'*  �ύX�L�[            �F20130227
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F�ܖ������̌�������ւ̊��Z�����~30.416�̏����_�ȉ��l�̌ܓ��ɕύX
	'******************************************************************'A-20130227-
	'*  �R���p�C����        �F2013/04/01
	'*  �ύX�L�[            �F20130401
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.���Y���͔��p�p���̑啶���݂̂Ƃ���
	'*  �C�����e            �F2.�u�����Z�v�̃��x���\����ǉ�
	'******************************************************************'A-20130401-
	'*  �R���p�C����        �F2013/04/24
	'*  �ύX�L�[            �F20130424
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.JAN�}�X�^����̐������̂�K17�ɕύX
	'******************************************************************'A-20130424-
	'*  �R���p�C����        �F2013/05/01
	'*  �ύX�L�[            �F20130501
	'*  �ύX�S����          �F�r�r�o�|�ڍ�
	'*  �C�����e            �F1.JAN�R�[�h���͎���JAN�}�X�^�ɑ��݂��Ȃ��Ă��n�j�Ƃ���
	'******************************************************************'A-20130501-
	'*  �R���p�C�����F2016/09/23
	'*  �ύX�L�[    �F20160726
	'*  �ύX�S����  �FSSP.MEGURO
	'*  �C�����e    �F��2016�N�V�X�e���ڍs���
	'*  �C�����e    �FVB6�ɕϊ�
	'*  �C�����e    �F�������ޖ��̘̂g���L����
	'********************************************************************
	'*  �R���p�C����        �F2017/02/03
	'*  �ύX�L�[            �F20170203
	'*  �ύX�S����          �F�r�r�o�|��R��
	'*  �C�����e            �F1.JAN�ϊ��e�[�u���Ή�
	'*                        2.JAN�����̃t�@���N�V�����\���p�̏C��
	'******************************************************************
	'*  �R���p�C����        �F2019/08/27
	'*  �ύX�L�[            �F20190601
	'*  �ύX�S����          �F�r�r�o�|�n��
	'*  �C�����e            �F1.�y���ŗ��Ή�
	'******************************************************************
	'******************************************************************
	'*  �R���p�C����        �F2024/02/22
	'*  �ύX�L�[            �F20240115
	'*  �ύX�S����          �F�r�r�o�|��k
	'*  �C�����e            �F1.�����Ȃ��A��������A�ܖ������̂��Âꂩ��I������I�v�V�������͂�ǉ�
	'*                            ����������ܖ�������I�����́A�����i�N�E���E���j�Ɛ��l�̓��͂�K�{�Ƃ���
	'******************************************************************
	'*  �R���p�C����        �F2025/02/03
	'*  �ύX�L�[            �F20250201
	'*  �ύX�S����          �F�r�r�o�|�A��
	'*  �C�����e            �F�@�Ȗڕ��ށA���ށA���[�����M����ʕ\�����ڂ���p�~����B
	'*                        �A���̑���ʂ́u����ŗ��敪�v���e�핪�ސ����ʂɈړ����A
	'*                          ����Ŋ֘A�̓��͍��ڂ��ꂩ���ɂ܂Ƃߑ��쐫�����シ��B
	'******************************************************************
	'*  �R���p�C����        �F2025/03/14
	'*  �ύX�L�[            �F20250303
	'*  �ύX�S����          �F�r�r�o�|���c
	'*  �C�����e            �F�@JAN�R�[�h�̏d���`�F�b�N�̒ǉ�
	'*                        �A�V�K�o�^�X�V��ɕi�Ԃ��̔Ԃ���Ȃ��s��̏C��
	'******************************************************************
	
	Public Const MAXSPREAD As Short = 500
	
	'   �R���g���[���Ǘ��p
	Structure CTRLTBL_S
		Dim IGRP As Short
		Dim INEXT As Short
		Dim IBACK As Short
		Dim IDOWN As Short
		Dim CTRL As System.Windows.Forms.Control
	End Structure
	'   �O���[�v�`�F�b�N�p
	Structure GRPTBL_S
		Dim CFLG As Short
		Dim NXTN As Short
	End Structure
	
	
	'��d�Ăяo���}�~�p
	'Public CallBack As Boolean                  'Client�ʒm�ς݃t���O
	'   �����n������
	'Public DBNAME   As String
	Public MOUSEFLG As Short '�}�E�X
	
	'�q�c�n�֘A�I�u�W�F�N�g
	Public RdoEnv As RDO.rdoEnvironment 'rdo�����
	Private qSZM0010SEL As RDO.rdoQuery
	Public qSZM0010_NSEL As RDO.rdoQuery
	Private qSZM0011SEL As RDO.rdoQuery
	Public PSZ0410SP As RDO.rdoQuery 'ADD-2001/01/23 ���є���ı�ޗp
	Private SZM0170_SEL As RDO.rdoQuery 'A 050909 TOP NAGANO
	Private SZM0170RS2SW As String 'A 050909 TOP NAGANO
	Private SZM0170RS2 As RDO.rdoResultset 'A 050909 TOP NAGANO
	Public WSZ0410SEL01 As RDO.rdoQuery 'A-CUST-20100610
	Public WSZ0410SEL02 As RDO.rdoQuery 'A-CUST-20100610
	Private qJANSEL As RDO.rdoQuery 'A-CUST20130212 JANϽ�SELECT
	Private qJAN_BUNRUISEL As RDO.rdoQuery 'A-CUST20130212 JAN����Ͻ�SELECT
	Private JAN_HENKANSEL1 As RDO.rdoQuery 'A-CUST-20170203
	Private JAN_HENKANSEL2 As RDO.rdoQuery 'A-CUST-20170203
	Private JAN_CHK_SEL As RDO.rdoQuery 'A-20250303
	
	Public MKKCMN As New MKKCMNPRJ.MKKCMNCLS
	Public MKKDBCMN As New MKKDBCMNPRJ.MKKDBCMNCLS
	Public CMTAX As New CMTAXPRJ.CMTAXCLS '����ŗ��擾���i           'A-20190601
	
	Public SZ0310 As New SZ0310PRJ.SZ0310CLS
	Public SZ0420 As New SZ0420PRJ.SZ0420CLS
	
	Public CM9500 As New CM9500.CM9500CLS
	Public CM9510 As New CM9510.CM9510CLS
	Public CM9520 As New CM9520.CM9520CLS
	Public CM9550 As New CM9550PRJ.CM9550CLS
	Public CM9600 As New CM9600.CM9600CLS '02/05/28 ADD
	
	'�Z�L�����e�B�A����
	'UPGRADE_WARNING: �z�� W_KENGEN �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Public W_KENGEN(3) As Integer
	
	Public SZ0720 As New SZ0720.SZ0720CLS
	Public SZ0730 As New SZ0730.SZ0730CLS
	Public SZ0740 As New SZ0740.SZ0740CLS
	Public SZ0750 As New SZ0750.SZ0750CLS
	Public SZ0760 As New SZ0760.SZ0760CLS
	'�X�C�b�`�@�G���A
	Public ENDSW As Short
	
	Public ERRSW As Short
	Public REDSW As Short
	Public CLRSW As Short
	Public BCHKSW As Short
	Public FCSMVSW As Short
	'�V�X�e���c�`�s�d
	Public SYSDATE As Date
	
	
	'�t�@���N�V�����@�G���A
	'Public ZAFC_MST(0 To 12) As String             'D-CUST-20100610
	'UPGRADE_WARNING: �z�� ZAFC_MST �̉����� 0 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Public ZAFC_MST(14) As String 'A-CUST-20100610
	'�K�C�h���b�Z�[�W�@�G���A
	'UPGRADE_WARNING: �z�� ZAGD_MST �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Public ZAGD_MST(100) As String
	
	
	Public KBKBN As Short '   �����敪
	Public WKB010 As String '   ��ЃR�[�h
	Public WKB020 As String '   ���Ə��R�[�h
	Public WKB010DSP As String '   ��Ж���
	Public WKB020DSP As String '   ���Ə�����
	Public WKB030 As String '   �i��
	Public WKB030DSP As String '   �i��
	Public WKB140DSP As String '   ���ޖ���
	Public WKB210DSP As String '   ���ޖ���
	Public WKB220DSP As String '   ���ޖ���
	Public WKB230DSP As String '   ���ޖ���
	Public WKB240DSP As String '   ���ޖ���
	Public WKB250DSP As String '   ���ޖ���
	Public WKB260DSP As String '   ���ޖ���
	Public WKB410DSP As String '   �ƎҖ���
	Public WKB291DSP As String '   JAN���i���ޖ���'A-CUST20130212
	
	Public WKB300 As Short '   �Ǘ��敪
	Public WKB310 As Short '   �����
	Public WKB320 As Short '   �I���P��
	Public WKB330 As Short '   �݌ɊǗ�
	Public WKB340 As Short '   FAX���M
	
	Public WKAMOCHUNM As String
	
	'BeginTrans�p����
	Public TRANSW As Boolean
	
	Public FKB010 As String
	Public FKB020 As String
	
	'   �d���i�ڃ}�X�^�f�[�^
	Public KB As SZM0010_S
	
	Dim SZM0010myRS As RDO.rdoResultset
	Dim SZM0010myRSSW As String
	Public SZM0010_NmyRS As RDO.rdoResultset
	Public SZM0010_NmyRSSW As String
	Dim SZM0011myRS As RDO.rdoResultset
	Dim SZM0011myRSSW As String
	
	'A-20250303��
	Public JAN_CHKRS As RDO.rdoResultset
	Public JAN_CHKRSSW As String
	'A-20250303��
	
	Public Const SPR_HEIGHT As Short = 330
	
	'A-CUST-20100610 Start
	'�b�r�u�捞�p�ϐ�
	Public WKBCSVFILE As String
	Public PRNSW As Short
	Public CSV_CNT As Short
	Public CSVERR_CNT As Short '�X�V�Ɏ��s��������
	Public FOPENSW As Boolean ' ���̓t�@�C�����n�o�d�m����
	Public INPFNum As Short ' ����̧�ٔԍ��i�[�ر
	Public INPFENDSW As Short '�@�e�L�X�g�ǂݍ��݂̏I��SW
	Public FSTSW As Short
	Public CANSW As Short
	
	Public Structure CSV_DATA
		Dim hin_name As String
		Dim kikaku As String
		Dim tani As String
		Dim gyo_name As String
		'tanka       As String          'D-CUST-20100823
		Dim tanka As Decimal 'A-CUST-20100823
		'A-CUST-20100823 Start
		Dim teki_date As String
		Dim ha_tani As String
		Dim kansansu As Decimal
		Dim jan_code As String
		Dim jan_s_code As String
		Dim bar_code As String
		'A-CUST-20100823 End
	End Structure
	Public WCSV_DATA As CSV_DATA 'CSV1�s��ǂݍ��񂾃f�[�^
	'Public IN_ITEM_MAX  As Integer     'D-CUST-20100823
	'UPGRADE_WARNING: �z�� IN_ITEM �̉����� 1 ���� 0 �ɕύX����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"' ���N���b�N���Ă��������B
	Public IN_ITEM(99) As String ' TXT�ϊ�ܰ�
	Private IN_ITEM_CNT As Short
	
	'A-CUST-20100823 Start
	Public Enum CsvPos
		DUMMY = 5
		teki_date
		ha_tani
		kansansu
		jan_code
		jan_s_code
		bar_code
		EndPos
	End Enum
	'A-CUST-20100823 End
	Public Const IN_ITEM_MAX As Double = CsvPos.EndPos - 1 'A-CUST-20100823
	
	Public SentakuFLG As Boolean
	Public RENBAN_SEN As Integer
	
	Public SETUZOKU As Boolean
	'A-CUST-20100610 End
	
	Public Tani_T() As String 'A-CUST-20100823
	Public TaniCnt As Short 'A-CUST-20100823
	
	Public JANCODESV As String 'A-CUST-20170203
	
	Public clearActCMB370Click As Boolean 'A-20250201
	
	'UPGRADE_WARNING: Sub Main() �����������Ƃ��ɃA�v���P�[�V�����͏I�����܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"' ���N���b�N���Ă��������B
	Public Sub Main()
		
		Dim wOPCODE As String
		
		'�p�����[�^�̃I�y���[�^�R�[�h��ޔ����Ă���
		If VB.Command() <> "" Then
			wOPCODE = Mid(VB.Command(), 1, 6)
		End If
		
		'INI�t�@�C������̎�o��
		Call INIGET_SUB("MKK.INI")
		
		'���s���[�h�̏ꍇ�͂ŃI�y���[�^�R�[�h���n����Ă�����I�y���[�^�ϐ�����������
		If CDbl(WG_DEBUG) = 0 And Trim(wOPCODE) <> "" Then
			WG_OPCODE = wOPCODE
		End If
		
		
		'   DB �w Connect ����
		If ZACNAP_SUB(WG_SZID, WG_SZPW, WG_SZDSN) = False Then
			'   �ڑ��Ɏ��s�����炷���I��
			Call ZAEND_SUB()
			Exit Sub
		End If
		'A-CUST-20100610 Start
		If ZACNA_SUB = False Then
			Call ZADISCN_SUB()
			Call ZAEND_SUB()
			Exit Sub
		End If
		SETUZOKU = True
		'A-CUST-20100610 End
		
		'   ��ʕ\��
		SZ0410FRM.Show()
		
	End Sub
	
	
	Public Sub INIT_RTN()
		
		'�G���[���b�Z�[�W�t�@�C���I�[�v��
		ZAER_FID = "RAZ99"
		Call ZAERO_SUB()
		If ZAER_ERR.Value <> "0" Then
			ERRSW = F_ERR
			Exit Sub
		End If
		
		
		'--- �t�@���N�V��������
		ZAFC_MST(1) = "�I  ��"
		ZAFC_MST(2) = ""
		ZAFC_MST(3) = "�⍇��"
		ZAFC_MST(4) = "��  ��"
		ZAFC_MST(5) = "�N���A"
		ZAFC_MST(6) = ""
		ZAFC_MST(7) = ""
		ZAFC_MST(6) = "�O�ꗗ" 'SZ0414�p   'A-CUST-20170203
		ZAFC_MST(7) = "���ꗗ" 'A-CUST-20170203
		ZAFC_MST(8) = "��  ��"
		ZAFC_MST(9) = ""
		ZAFC_MST(10) = ""
		ZAFC_MST(11) = ""
		ZAFC_MST(11) = "�I�@��" 'SZ0414�p   'A-CUST-20170203
		ZAFC_MST(12) = "��  �s"
		ZAFC_MST(13) = "�i�ڎ捞" 'A-CUST-20100610
		ZAFC_MST(14) = "�i�ڑI��" 'A-CUST-20100610
		
		ZAGD_MST(1) = "�I��..ESC"
		ZAGD_MST(2) = "�����..."
		ZAGD_MST(3) = ""
		ZAGD_MST(4) = "��к��ނ���͂��Ă��������B"
		ZAGD_MST(5) = "���Ə����ނ���͂��Ă��������B"
		ZAGD_MST(6) = "�i�Ԃ���͂��Ă��������B"
		ZAGD_MST(7) = "�i������͂��Ă��������B"
		ZAGD_MST(8) = "�K�i����͂��Ă��������B"
		ZAGD_MST(9) = "�P�ʂ�I�����Ă��������B"
		ZAGD_MST(10) = "JAN�W�����ނ���͂��Ă��������B"
		ZAGD_MST(11) = "JAN�Z�k����͂��Ă��������B"
		ZAGD_MST(12) = "���̑����ް���ނ���͂��Ă��������B"
		ZAGD_MST(13) = "�K�p�J�n������͂��Ă��������B"
		ZAGD_MST(14) = "��������͂��Ă��������B"
		ZAGD_MST(15) = "�_�񉿊i����͂��Ă��������B"
		
		ZAGD_MST(16) = "��p�Ȗځ|���v�f����͂��Ă��������B"
		ZAGD_MST(17) = "��p�Ȗځ|���v�f����͂��Ă��������B"
		ZAGD_MST(18) = "�Ȗڕ��ނ���͂��Ă��������B"
		
		ZAGD_MST(19) = "�啪�ނ���͂��Ă��������B"
		'    ZAGD_MST(20) = "�����ނ���͂��Ă��������B"                    'D-20190601
		ZAGD_MST(20) = "*�t���̐ŗ��\���͌y���ŗ��ł��B" 'A-20190601
		ZAGD_MST(21) = "�����ނ���͂��Ă��������B"
		ZAGD_MST(22) = "���ނ���͂��Ă��������B"
		ZAGD_MST(23) = "�������ނ���͂��Ă��������B"
		ZAGD_MST(24) = "������i�̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(25) = "�d�|�敪�̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(26) = "���c���̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(27) = "�Ǘ��敪��I�����Ă��������B"
		'    ZAGD_MST(28) = "����ł�I�����Ă��������B"                    'D-20190601
		ZAGD_MST(28) = "*�t���̐ŗ��\���͌y���ŗ��ł��B" 'A-20190601
		ZAGD_MST(29) = "�I���P����I�����Ă��������B"
		ZAGD_MST(30) = "�݌ɊǗ���I�����Ă��������B"
		'ZAGD_MST(31) = "FAX���M��I�����Ă��������B"                   'D-CUST-20100901
		ZAGD_MST(31) = "���[�����M��I�����Ă��������B" 'A-CUST-20100901
		ZAGD_MST(32) = "�����P�ʂ�I�����Ă��������B"
		ZAGD_MST(33) = "���Z����I�����Ă��������B"
		ZAGD_MST(34) = "�ƎҌ��躰�ނ���͂��Ă��������B"
		ZAGD_MST(35) = "�������躰�ނ���͂��Ă��������B"
		ZAGD_MST(36) = "���ꔭ���̏ꍇ�`�F�b�N���Ă��������B"
		'    ZAGD_MST(37) = "����ŋ敪�i1�`�T�j����͂��Ă��������B"       'D-20190601
		ZAGD_MST(37) = "*�t���̐ŗ��\���͌y���ŗ��ł��B" 'A-20190601
		ZAGD_MST(38) = "�����i�̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(39) = "���̋@�̔��̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(40) = "����Ώۂ̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(41) = "�ŏI�[�i������͂��Ă��������B"
		ZAGD_MST(42) = "�K�p�J�n������͂��Ă��������B"
		ZAGD_MST(43) = "�����x�~�̏ꍇ�`�F�b�N���Ă��������B"
		ZAGD_MST(44) = "�����x�~������͂��Ă��������B"
		ZAGD_MST(45) = "�����敪��I�����Ă��������B"
		ZAGD_MST(46) = "���s�{�^���������Ă��������B"
		'A-CUST-20100610 Start
		ZAGD_MST(47) = "����������͂��Ă��������B"
		ZAGD_MST(48) = "�t�@�C��������͂��ĉ������B"
		'A-CUST-20100610 End
		'A-CUST20130212��
		ZAGD_MST(49) = "���Y������͂��ĉ������B"
		ZAGD_MST(50) = "�d�ʂ���͂��ĉ������B"
		ZAGD_MST(51) = "�ܖ���������͂��ĉ������B"
		ZAGD_MST(52) = "JAN���i���ނ���͂��ĉ������B"
		'A-CUST20130212��
		ZAGD_MST(53) = "����/�ܖ���������͂��Ă��������B" 'A-20240115
		ZAGD_MST(54) = "�����ނ���͂��Ă��������B" 'A-20250201
		'UPGRADE_NOTE: Erase �� System.Array.Clear �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		System.Array.Clear(ZAFC_USE, 0, ZAFC_USE.Length)
		ZAFC_USE(0) = True
		ZAFC_USE(1) = False
		ZAFC_USE(2) = False
		ZAFC_USE(3) = True
		ZAFC_USE(4) = True
		ZAFC_USE(5) = True
		'ZAFC_USE(6) = False                    'D-CUST-20100610
		'ZAFC_USE(7) = False                    'D-CUST-20100610
		ZAFC_USE(6) = True 'A-CUST-20100610
		ZAFC_USE(7) = True 'A-CUST-20100610
		ZAFC_USE(8) = True
		'ZAFC_USE(9) = False        'D-20110621-
		ZAFC_USE(9) = True 'A-20110621-
		ZAFC_USE(10) = False
		ZAFC_USE(11) = False
		ZAFC_USE(12) = True
		
		
		'�Z�L�����e�B�`�F�b�N�i�P�j�N������
		Dim Ret As Short
		
		MKKDBCMN.MKKDBCMN_RCN = ZACN_RCN
		Ret = MKKDBCMN.MKKDBCMN_SQTGET1_SUB(3, "SZ0410", VB6.Format(Val(WG_INCCODE), "00"), WG_OPCODE, W_KENGEN(1))
		If Ret <> n0 Then
			Call ENDR_RTN()
		ElseIf W_KENGEN(1) = 0 Then 
			ZAER_KN = n0
			ZAER_CD = 301
			ZAER_MS.Value = ""
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			Call ENDR_RTN()
		End If
		
		
		'�Z�L�����e�B�`�F�b�N�i�Q�j���Ə��Q�ƌ���
		MKKDBCMN.MKKDBCMN_RCN = ZACN_RCN
		Ret = MKKDBCMN.MKKDBCMN_SQTGET2_SUB(3, "SZ0410", WG_INCCODE, WG_JGCODE, WG_OPCODE, W_KENGEN(2))
		If Ret <> n0 Then
			ERRSW = F_ERR
			Exit Sub
		ElseIf W_KENGEN(2) = 0 Then 
			ERRSW = F_ERR
			ZAER_KN = n0
			ZAER_CD = 302
			ZAER_NO.Value = ""
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			Exit Sub
		End If
		
		WKB300 = 1
		WKB310 = 1
		WKB320 = 1
		WKB330 = 1
		WKB340 = 1
		
		WKB010 = WG_INCCODE
		WKB020 = WG_JGCODE
		
		Call PREP_RTN()
		
	End Sub
	
	Public Sub ENDR_RTN()
		
		'A-CUST-20100610 Start
		If SETUZOKU Then
			Call ZADISCN_SUB()
			Call ZADISCNA_SUB()
		End If
		'A-CUST-20100610 End
		Call ZAEND_SUB()
		
	End Sub
	
	Public Sub DBBeginTrans()
		
		' DB �ɖ₢���킹��...
		On Error Resume Next
		'RdoEnv.BeginTrans                  'D-CUST-20100610
		ZACN_RCN.BeginTrans() 'A-CUST-20100610
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		Else
			TRANSW = True
		End If
		On Error GoTo 0
		
	End Sub
	
	Public Sub DBCommitTrans()
		
		On Error Resume Next
		'RdoEnv.CommitTrans                 'D-CUST-20100610
		ZACN_RCN.CommitTrans() 'A-CUST-20100610
		If B_STATUS = 0 Then
			TRANSW = False
		Else
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		End If
		On Error GoTo 0
		
	End Sub
	
	Public Sub DBRollbackTrans()
		
		On Error Resume Next
		'RdoEnv.RollbackTrans               'D-CUST-20100610
		ZACN_RCN.RollbackTrans() 'A-CUST-20100610
		If B_STATUS = 0 Then
			TRANSW = False
		Else
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		End If
		On Error GoTo 0
		
	End Sub
	
	Private Sub LOG_PUT_RTN(ByRef strKAISHA As String, ByRef strJIGYO As String, ByRef SvrDate As String, ByRef strPgm As String, ByRef strKbn As String, ByRef strPrc As String)
		'���O�o��
		
		''''Call BEGIN_RTN
		If ERRSW = F_ERR Then
			Exit Sub
		End If
		
		ZALGM_INC_CODE.Value = strKAISHA '��к���
		ZALGM_JG_CODE.Value = strJIGYO '���Ə��R�[�h
		ZALGM_SYS_KBN.Value = CStr(3) '�V�X�e���敪
		ZALGM_S_DAY.Value = Mid(SvrDate, 1, 8) '�������t
		ZALGM_S_TIME.Value = Mid(SvrDate, 9, 6) '��������
		ZALGM_OP_CODE.Value = WG_OPCODE '�I�y���[�^�R�[�h
		ZALGM_PGID.Value = strPgm '�v���O�����h�c�i���p�啶���j
		ZALGM_SH_KBN.Value = strKbn '�����敪
		ZALGM_SH_NAIYO.Value = strPrc '�������e
		ZALGM_GNFLG.Value = "0" '���z�t���O
		
		Call ZALGM_SUB(ZACN_RCN)
		If ZALGM_ERR.Value <> "0" Then
			'   2000/02/01  LOG_PUT_ERROR�͖�������B
			''''ERRSW = F_ERR
			''''Call ROLLBACK_RTN
			Exit Sub
		End If
		
		''''Call COMMIT_RTN
		'    If ERRSW = F_ERR Then
		'    ''''Call ROLLBACK_RTN
		'        Exit Sub
		'    End If
		
	End Sub
	
	Public Sub PREP_RTN()
		
		Call CduPrepKaisha()
		Call CduPrepJigyo()
		Call PREP_SZM0010()
		Call PREP_SZM0011()
		
		Call CduPrepOper()
		
		Call PrepFind()
		Call CduPrepDAIBunrui()
		Call CduPrepCHUBunrui()
		Call CduPrepSHOBunrui()
		Call PrepBunrui() '02/05/28 ADD
		
		Call PrepKAMOCHU()
		Call PrepKAMOKU()
		Call PrepGYOSHA()
		''''Call PrepBUSHO
		Call CduPrepBUSHO()
		
		Call PrepKamBunrui()
		
		Call TaiouKamokuPrep()
		
		Call PSZ0410_PREP_RTN() 'ADD-2001/01/23 �i�ڂ̎��є���ı������߱
		
		Call PREP_WSZ0410_RTN() 'A-CUST-20100610
		
		Call PREP_JAN_RTN() 'A-CSUT20130212 JANϽ��擾
		
		Call PREP_JAN_BUNRUI_RTN() 'A-CUST20130212 JAN����Ͻ��擾
		
		Call PREP_JAN_HENKAN_RTN() 'A-CUST-20170203
		
		Call PREP_JAN_CHK_RTN() 'A-20250303
		
	End Sub
	
	
	Public Sub COMBO_INIT(ByRef cBox As System.Windows.Forms.ComboBox)
		
		Dim nUnit As Short
		
		nUnit = CduLoadUNIT(WKB010, WKB020, cBox)
		
		
	End Sub
	
	Public Sub COMBO_SETLIST(ByRef cBox As System.Windows.Forms.ComboBox, ByRef Txt As String)
		
		Dim lx As Integer
		For lx = 0 To cBox.Items.Count - 1
			If Trim(VB6.GetItemString(cBox, lx)) = Trim(Txt) Then
				cBox.SelectedIndex = lx
				Exit Sub
			End If
		Next lx
		cBox.SelectedIndex = -1
		
	End Sub
	Public Sub GO_EXEC()
		
		Dim iReturn As Short
		
		Dim strSvrDate As String
		
		Call ZASYS_SUB(strSvrDate, 3)
		
		Call LOG_PUT_RTN(WKB010, WKB020, strSvrDate, "SZ0410", "" & (KBKBN + 2), "SZ0410")
		
		KB.Inc_code = WKB010
		KB.jg_code = WKB020
		KB.hin_code = WKB030
		
		'UPGRADE_WARNING: Screen �v���p�e�B Screen.MousePointer �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		'A-CUST-20100610 Start
		If SentakuFLG Then
			SentakuFLG = False
			Call GO_WKDELETE()
			If ENDSW = F_END Then Exit Sub
		End If
		'A-CUST-20100610 End
		
		Select Case KBKBN
			Case 1
				iReturn = GO_INSERT()
			Case 2
				iReturn = GO_UPDATE()
			Case 3
				iReturn = GO_DELETE()
				
		End Select
		
		'UPGRADE_WARNING: Screen �v���p�e�B Screen.MousePointer �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Public Function GO_INSERT() As Short
		
		Dim strToday As String
		Dim SYSDATE_YMD As String
		Dim SYSDATE_HMS As String
		
		GO_INSERT = F_OFF
		
		
		SYSDATE = CduServerDate
		strToday = VB6.Format(SYSDATE, "YYYYMMDDHHMMSS")
		SYSDATE_YMD = Mid(strToday, 1, 8)
		SYSDATE_HMS = Mid(strToday, 9, 6)
		
		KB.Entry_Op_code = WG_OPCODE
		KB.Entry_Op_date = SYSDATE_YMD
		KB.Entry_Op_time = SYSDATE_HMS
		
		Call DBBeginTrans()
		
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Inc_code, 2) '��к���
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jg_code, 4) '���Ə�����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5) '�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("hin_name").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_name, 20) '�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("kikaku").Value = MKKCMN.ZACHGSTR_SUB(KB.kikaku, 20) '�K�i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tani").Value = MKKCMN.ZACHGSTR_SUB(KB.tani, 4) '�P��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("jan_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jan_code, 13) 'JAN�W������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("jan_s_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jan_s_code, 7) 'JAN�Z�k
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("bar_code").Value = MKKCMN.ZACHGSTR_SUB(KB.bar_code, 30) '���̑����ް����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("teki_date1").Value = MKKCMN.ZACHGSTR_SUB(KB.teki_date1, 8) '�����E�v���P
		SZM0010INS.rdoParameters("baika1").Value = KB.baika1 '�����P
		SZM0010INS.rdoParameters("kei_kin1").Value = KB.kei_kin1 '�_�񉿊i1
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("teki_date2").Value = MKKCMN.ZACHGSTR_SUB(KB.teki_date2, 8) '�����E�v���Q
		SZM0010INS.rdoParameters("baika2").Value = KB.baika2 '�����Q
		SZM0010INS.rdoParameters("kei_kin2").Value = KB.kei_kin2 '�_�񉿊i�Q
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("hiyou_k_code1").Value = MKKCMN.ZACHGSTR_SUB(KB.hiyou_k_code1, 3) '��p�Ȗځi���v�f�j
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("hiyou_k_code2").Value = MKKCMN.ZACHGSTR_SUB(KB.hiyou_k_code2, 6) '��p�Ȗځi���v�f�j
		'SZM0010INS!ka_bun_code = MKKCMN.ZACHGSTR_SUB(KB.ka_bun_code, 7)       '�Ȗڕ���    'D-20250201
		SZM0010INS.rdoParameters("ka_bun_code").Value = " " '�Ȗڕ���                                        'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("l_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.l_bun_code, 4) '�啪��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("m_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.m_bun_code, 4) '������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("s_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.s_bun_code, 4) '������
		'SZM0010INS!bun_code = MKKCMN.ZACHGSTR_SUB(KB.bun_code, 4)             '����    'D-20250201
		SZM0010INS.rdoParameters("bun_code").Value = " " '����                                     'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ken_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.ken_bun_code, 4) '��������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("jutaku").Value = MKKCMN.ZACHGSTR_SUB(KB.jutaku, 1) '������i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("sikakari").Value = MKKCMN.ZACHGSTR_SUB(KB.sikakari, 1) '�d�|�敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("zan").Value = MKKCMN.ZACHGSTR_SUB(KB.zan, 1) 'ϲŽ�c����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("zaiko").Value = MKKCMN.ZACHGSTR_SUB(KB.zaiko, 1) '�݌ɊǗ�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("kanri_kubn").Value = MKKCMN.ZACHGSTR_SUB(KB.kanri_kubn, 1) '�Ǘ��敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Tax_kubn").Value = MKKCMN.ZACHGSTR_SUB(KB.Tax_kubn, 1) '����ŋ敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tana_tanka").Value = MKKCMN.ZACHGSTR_SUB(KB.tana_tanka, 1) '�I���P���敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ha_tanka1").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka1, 4) '�����P��1
		SZM0010INS.rdoParameters("kansan_num1").Value = KB.kansan_num1 '���Z��1
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ha_tanka2").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka2, 4) '�����P��2
		SZM0010INS.rdoParameters("kansan_num2").Value = KB.kansan_num2 '���Z��2
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ha_tanka3").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka3, 4) '�����P��3
		SZM0010INS.rdoParameters("kansan_num3").Value = KB.kansan_num3 '���Z��3
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ha_tanka4").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka4, 4) '�����P��4
		SZM0010INS.rdoParameters("kansan_num4").Value = KB.kansan_num4 '���Z��4
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("ha_tanka5").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka5, 4) '�����P��5
		SZM0010INS.rdoParameters("kansan_num5").Value = KB.kansan_num5 '���Z��5
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("g_gentei_code").Value = MKKCMN.ZACHGSTR_SUB(KB.g_gentei_code, 6) '�ƎҌ��躰��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("gen_h_ka").Value = MKKCMN.ZACHGSTR_SUB(KB.gen_h_ka, 1) '���ꔭ����
		'SZM0010INS!Fax_yn = MKKCMN.ZACHGSTR_SUB(KB.Fax_yn, 1)             'Fax���M�s�� 'D-20250201
		SZM0010INS.rdoParameters("Fax_yn").Value = "0" 'Fax���M�s��                                'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tax_rate_kbn").Value = MKKCMN.ZACHGSTR_SUB(KB.tax_rate_kbn, 1) '�ŗ��敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tyozouhin").Value = MKKCMN.ZACHGSTR_SUB(KB.tyozouhin, 1) '�����i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("jihan").Value = MKKCMN.ZACHGSTR_SUB(KB.jihan, 1) '���̋@�̔�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("gensen").Value = MKKCMN.ZACHGSTR_SUB(KB.gensen, 1) '����Ώ�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("nouhin_date").Value = MKKCMN.ZACHGSTR_SUB(KB.nouhin_date, 8) '�ŏI�[�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tekiyo_date").Value = MKKCMN.ZACHGSTR_SUB(KB.tekiyo_date, 8) '�K�p�J�n���t
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tori_kyu").Value = MKKCMN.ZACHGSTR_SUB(KB.tori_kyu, 1) '���x�~
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("tori_kyu_date").Value = MKKCMN.ZACHGSTR_SUB(KB.tori_kyu_date, 8) '���x�~���t
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Entry_Op_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_code, 6) '�o�^�I�y���[�^
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Entry_Op_date").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_date, 8) '�o�^�n��_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Entry_Op_time").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_time, 6) '�o�^�n��_time
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Edit_Op_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_code, 6) '�C���I�y���[�^
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Edit_Op_date").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_date, 8) '�C���n��_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("Edit_Op_time").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_time, 6) '�C���n��_time
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010INS.rdoParameters("del_flg").Value = MKKCMN.ZACHGSTR_SUB(KB.del_flg, 1) '�폜�t���O
		'A-CUST-20100610 Start
		If RTrim(KB.hin_name_seisiki) = "" Then
			SZM0010INS.rdoParameters("hin_name_seisiki").Value = " "
		Else
			SZM0010INS.rdoParameters("hin_name_seisiki").Value = RTrim(KB.hin_name_seisiki) '��������
		End If
		'A-CUST-20100610 End
		'A-CUST20130212��
		SZM0010INS.rdoParameters("BK1").Value = KB.BK1
		SZM0010INS.rdoParameters("k42").Value = KB.k42
		SZM0010INS.rdoParameters("k44").Value = KB.k44
		SZM0010INS.rdoParameters("k57").Value = KB.k57
		SZM0010INS.rdoParameters("k58").Value = KB.k58
		SZM0010INS.rdoParameters("k99").Value = KB.k99
		'A-CUST20130212��
		SZM0010INS.rdoParameters("Shomi_date_kbn").Value = KB.Shomi_date_kbn 'A-20240115
		On Error Resume Next
		
		SZM0010INS.Execute()
		
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ERRSW = F_ERR
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010INS"
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			GoTo GO_INSERT_ERR
		End If
		On Error GoTo 0
		
		Call GO_BUSHO(1)
		If ERRSW = F_ERR Then
			GoTo GO_INSERT_ERR
		End If
		
		Call DBCommitTrans()
		Exit Function
		
GO_INSERT_ERR: 
		Call DBRollbackTrans()
		
		
	End Function
	
	
	Public Function GO_UPDATE() As Short
		
		Dim strToday As String
		Dim SYSDATE_YMD As String
		Dim SYSDATE_HMS As String
		
		GO_UPDATE = F_OFF
		
		SYSDATE = CduServerDate
		strToday = VB6.Format(SYSDATE, "YYYYMMDDHHMMSS")
		SYSDATE_YMD = Mid(strToday, 1, 8)
		SYSDATE_HMS = Mid(strToday, 9, 6)
		
		KB.Edit_Op_code = WG_OPCODE
		KB.Edit_Op_date = SYSDATE_YMD
		KB.Edit_Op_time = SYSDATE_HMS
		
		Debug.Print("GO_UPDATE" & Mid(SZM0010UPD.SQL, 1, 100))
		Debug.Print("GO_UPDATE" & Mid(SZM0010UPD.SQL, 101, 100))
		Debug.Print("GO_UPDATE" & Mid(SZM0010UPD.SQL, 201, 100))
		Debug.Print("GO_UPDATE" & Mid(SZM0010UPD.SQL, 301, 100))
		Debug.Print("GO_UPDATE" & Mid(SZM0010UPD.SQL, 401, 100))
		
		
		On Error Resume Next
		
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Inc_code, 2) '��к���
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jg_code, 4) '���Ə�����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5) '�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("hin_name").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_name, 20) '�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("kikaku").Value = MKKCMN.ZACHGSTR_SUB(KB.kikaku, 20) '�K�i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tani").Value = MKKCMN.ZACHGSTR_SUB(KB.tani, 4) '�P��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("jan_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jan_code, 13) 'JAN�W������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("jan_s_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jan_s_code, 7) 'JAN�Z�k
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("bar_code").Value = MKKCMN.ZACHGSTR_SUB(KB.bar_code, 30) '���̑����ް����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("teki_date1").Value = MKKCMN.ZACHGSTR_SUB(KB.teki_date1, 8) '�����E�v���P
		SZM0010UPD.rdoParameters("baika1").Value = KB.baika1 '�����P
		SZM0010UPD.rdoParameters("kei_kin1").Value = KB.kei_kin1 '�_�񉿊i1
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("teki_date2").Value = MKKCMN.ZACHGSTR_SUB(KB.teki_date2, 8) '�����E�v���Q
		SZM0010UPD.rdoParameters("baika2").Value = KB.baika2 '�����Q
		SZM0010UPD.rdoParameters("kei_kin2").Value = KB.kei_kin2 '�_�񉿊i�Q
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("hiyou_k_code1").Value = MKKCMN.ZACHGSTR_SUB(KB.hiyou_k_code1, 3) '��p�Ȗځi���v�f�j
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("hiyou_k_code2").Value = MKKCMN.ZACHGSTR_SUB(KB.hiyou_k_code2, 6) '��p�Ȗځi���v�f�j
		'SZM0010UPD!ka_bun_code = MKKCMN.ZACHGSTR_SUB(KB.ka_bun_code, 7)       '�Ȗڕ���    'D-20250201
		SZM0010UPD.rdoParameters("ka_bun_code").Value = " " '�Ȗڕ���                                        'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("l_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.l_bun_code, 4) '�啪��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("m_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.m_bun_code, 4) '������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("s_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.s_bun_code, 4) '������
		'SZM0010UPD!bun_code = MKKCMN.ZACHGSTR_SUB(KB.bun_code, 4)             '����    'D-20250201
		SZM0010UPD.rdoParameters("bun_code").Value = " " '����                                     'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ken_bun_code").Value = MKKCMN.ZACHGSTR_SUB(KB.ken_bun_code, 4) '��������
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("jutaku").Value = MKKCMN.ZACHGSTR_SUB(KB.jutaku, 1) '������i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("sikakari").Value = MKKCMN.ZACHGSTR_SUB(KB.sikakari, 1) '�d�|�敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("zan").Value = MKKCMN.ZACHGSTR_SUB(KB.zan, 1) 'ϲŽ�c����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("zaiko").Value = MKKCMN.ZACHGSTR_SUB(KB.zaiko, 1) '�݌ɊǗ�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("kanri_kubn").Value = MKKCMN.ZACHGSTR_SUB(KB.kanri_kubn, 1) '�Ǘ��敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Tax_kubn").Value = MKKCMN.ZACHGSTR_SUB(KB.Tax_kubn, 1) '����ŋ敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tana_tanka").Value = MKKCMN.ZACHGSTR_SUB(KB.tana_tanka, 1) '�I���P���敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ha_tanka1").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka1, 4) '�����P��1
		SZM0010UPD.rdoParameters("kansan_num1").Value = KB.kansan_num1 '���Z��1
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ha_tanka2").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka2, 4) '�����P��2
		SZM0010UPD.rdoParameters("kansan_num2").Value = KB.kansan_num2 '���Z��2
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ha_tanka3").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka3, 4) '�����P��3
		SZM0010UPD.rdoParameters("kansan_num3").Value = KB.kansan_num3 '���Z��3
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ha_tanka4").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka4, 4) '�����P��4
		SZM0010UPD.rdoParameters("kansan_num4").Value = KB.kansan_num4 '���Z��4
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("ha_tanka5").Value = MKKCMN.ZACHGSTR_SUB(KB.ha_tanka5, 4) '�����P��5
		SZM0010UPD.rdoParameters("kansan_num5").Value = KB.kansan_num5 '���Z��5
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("g_gentei_code").Value = MKKCMN.ZACHGSTR_SUB(KB.g_gentei_code, 6) '�ƎҌ��躰��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("gen_h_ka").Value = MKKCMN.ZACHGSTR_SUB(KB.gen_h_ka, 1) '���ꔭ����
		'SZM0010UPD!Fax_yn = MKKCMN.ZACHGSTR_SUB(KB.Fax_yn, 1)             'Fax���M�s�� 'D-20250201
		SZM0010UPD.rdoParameters("Fax_yn").Value = "0" 'Fax���M�s��                                'A-20250201
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tax_rate_kbn").Value = MKKCMN.ZACHGSTR_SUB(KB.tax_rate_kbn, 1) '�ŗ��敪
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tyozouhin").Value = MKKCMN.ZACHGSTR_SUB(KB.tyozouhin, 1) '�����i
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("jihan").Value = MKKCMN.ZACHGSTR_SUB(KB.jihan, 1) '���̋@�̔�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("gensen").Value = MKKCMN.ZACHGSTR_SUB(KB.gensen, 1) '����Ώ�
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("nouhin_date").Value = MKKCMN.ZACHGSTR_SUB(KB.nouhin_date, 8) '�ŏI�[�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tekiyo_date").Value = MKKCMN.ZACHGSTR_SUB(KB.tekiyo_date, 8) '�K�p�J�n���t
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tori_kyu").Value = MKKCMN.ZACHGSTR_SUB(KB.tori_kyu, 1) '���x�~
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("tori_kyu_date").Value = MKKCMN.ZACHGSTR_SUB(KB.tori_kyu_date, 8) '���x�~���t
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Entry_Op_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_code, 6) '�o�^�I�y���[�^
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Entry_Op_date").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_date, 8) '�o�^�n��_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Entry_Op_time").Value = MKKCMN.ZACHGSTR_SUB(KB.Entry_Op_time, 6) '�o�^�n��_time
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Edit_Op_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_code, 6) '�C���I�y���[�^
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Edit_Op_date").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_date, 8) '�C���n��_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("Edit_Op_time").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_time, 6) '�C���n��_time
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010UPD.rdoParameters("del_flg").Value = MKKCMN.ZACHGSTR_SUB(KB.del_flg, 1) '�폜�t���O
		'A-CUST-20100610 Start
		If RTrim(KB.hin_name_seisiki) = "" Then
			SZM0010UPD.rdoParameters("hin_name_seisiki").Value = " "
		Else
			SZM0010UPD.rdoParameters("hin_name_seisiki").Value = RTrim(KB.hin_name_seisiki) '��������
		End If
		'A-CUST-20100610 End
		'A-CUST20130212��
		SZM0010UPD.rdoParameters("BK1").Value = KB.BK1
		SZM0010UPD.rdoParameters("k42").Value = KB.k42
		SZM0010UPD.rdoParameters("k44").Value = KB.k44
		SZM0010UPD.rdoParameters("k57").Value = KB.k57
		SZM0010UPD.rdoParameters("k58").Value = KB.k58
		SZM0010UPD.rdoParameters("k99").Value = KB.k99
		'A-CUST20130212��
		SZM0010UPD.rdoParameters("Shomi_date_kbn").Value = KB.Shomi_date_kbn 'A-20240115
		SZM0010UPD.Execute()
		
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ERRSW = F_ERR
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010"
			ZAER_MS.Value = "" 'SM02_KEY0.S001
			Call ZAER_SUB()
			GoTo GO_UPDATE_ERR
		End If
		On Error GoTo 0
		
		Call GO_BUSHO(2)
		If ERRSW = F_ERR Then
			GoTo GO_UPDATE_ERR
		End If
		
		'A-CUST-20170203 Start
		'�i�ڃ}�X�^��JAN�ϊ��e�[�u����JAN�W���R�[�h�͓������Ă���̂��O��
		If RTrim(KB.jan_code) = JANCODESV Then
		ElseIf RTrim(KB.jan_code) = "" Then  '�󔒂�JAN�ϊ����Ȃ��ꍇ�݂̂Ȃ̂Ŏ��ۂɂ͕K�v�Ȃ����f
		Else
			Call UPD_JAN_HENKAN() 'JAN�ϊ��e�[�u�������݂��Ȃ���Δ����邾��
			If ERRSW = F_ERR Then
				GoTo GO_UPDATE_ERR
			End If
		End If
		'A-CUST-20170203e
		
		Call DBCommitTrans()
		Exit Function
		
GO_UPDATE_ERR: 
		Call DBRollbackTrans()
		
		
	End Function
	
	Public Function GO_DELETE() As Short
		
		Dim strToday As String
		Dim SYSDATE_YMD As String
		Dim SYSDATE_HMS As String
		
		
		GO_DELETE = F_OFF
		
		SYSDATE = CduServerDate
		strToday = VB6.Format(SYSDATE, "YYYYMMDDHHMMSS")
		SYSDATE_YMD = Mid(strToday, 1, 8)
		SYSDATE_HMS = Mid(strToday, 9, 6)
		
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		KB.Edit_Op_code = MKKCMN.ZACHGSTR_SUB(WG_OPCODE, 6)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		KB.Edit_Op_date = MKKCMN.ZACHGSTR_SUB(SYSDATE_YMD, 8)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		KB.Edit_Op_time = MKKCMN.ZACHGSTR_SUB(SYSDATE_HMS, 6)
		
		SZM0010DEL.rdoParameters("DelOpCode").Value = WG_OPCODE
		SZM0010DEL.rdoParameters("DelOpDate").Value = SYSDATE_YMD
		SZM0010DEL.rdoParameters("DelOpTime").Value = SYSDATE_HMS
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010DEL.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Inc_code, 2) '��к���
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010DEL.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jg_code, 4) '���Ə�����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0010DEL.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5) '�i��
		
		SZM0010DEL.Execute()
		
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ERRSW = F_ERR
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010DEL"
			ZAER_MS.Value = ""
			Call ZAER_SUB()
			GoTo GO_DELETE_ERR
		End If
		
		Call GO_BUSHO(3)
		If ERRSW = F_ERR Then
			GoTo GO_DELETE_ERR
		End If
		
		Call DBCommitTrans()
		Exit Function
		
		
GO_DELETE_ERR: 
		Call DBRollbackTrans()
		
		
	End Function
	
	Public Sub GO_BUSHO(ByRef iKBN As Short)
		
		Dim iReturn As Short
		
		If iKBN = 2 Or iKBN = 3 Then ' UPDATE or DELETE
			iReturn = FILSZM0011_DELETE()
		End If
		
		If iKBN = 1 Or iKBN = 2 Then ' INSERT or UPDATE
			iReturn = FILSZM0011_INSERT()
		End If
		
	End Sub
	
	Public Function GETTODAY() As String
		
		Dim strToday As String
		Dim SYSDATE_YMD As String
		Dim SYSDATE_HMS As String
		
		SYSDATE = CduServerDate
		strToday = VB6.Format(SYSDATE, "YYYYMMDDHHMMSS")
		GETTODAY = Mid(strToday, 1, 8)
		SYSDATE_HMS = Mid(strToday, 9, 6)
		
	End Function
	Public Function FILSZM0011_DELETE() As Short
		
		FILSZM0011_DELETE = F_OFF
		
		SZM0011DEL.rdoParameters("Inc_code").Value = WKB010
		SZM0011DEL.rdoParameters("jg_code").Value = WKB020
		SZM0011DEL.rdoParameters("hin_code").Value = WKB030
		SZM0011DEL.Execute()
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0021"
			ZAER_MS.Value = "" 'SM02_KEY0.S001
			Call ZAER_SUB()
			FILSZM0011_DELETE = F_ERR
		End If
		
	End Function
	
	Public Function FILSZM0011_INSERT() As Short
		
		Dim nCnt As Integer
		Dim cdBUSHO As String
		Dim i As Short
		
		FILSZM0011_INSERT = F_OFF
		
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0011INS.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(WKB010, 2)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0011INS.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(WKB020, 4)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		SZM0011INS.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(WKB030, 5)
		
		With SZ0410FRM
			nCnt = .SPR420.DataRowCnt
			For i = 1 To nCnt
				.SPR420.ROW = i
				.SPR420.Col = 1
				cdBUSHO = ZeroFill((.SPR420.Text), 4) '   2000/02/20
				If Trim(cdBUSHO) <> "" Then
					SZM0011INS.rdoParameters("y_code").Value = i
					'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
					SZM0011INS.rdoParameters("bu_code").Value = MKKCMN.ZACHGSTR_SUB(cdBUSHO, 4)
					
					SZM0011INS.Execute()
					If B_STATUS <> 0 Then
						ENDSW = F_END
						ZAER_KN = 1
						ZAER_NO.Value = "SZM0011INS"
						ZAER_MS.Value = "" 'SM02_KEY0.S001
						Call ZAER_SUB()
						FILSZM0011_INSERT = F_ERR
						Exit Function
					End If
				End If
				
			Next i
		End With
		
	End Function
	
	
	Public Sub COMBO_COPY(ByRef cboxSRC As System.Windows.Forms.ComboBox, ByRef cboxDST As System.Windows.Forms.ComboBox)
		'   COMBOBOOX��LIST���𕡎ʂ���B
		
		
		Dim l As Integer
		
		cboxDST.Items.Clear()
		
		For l = 0 To cboxSRC.Items.Count - 1
			cboxDST.Items.Add(VB6.GetItemString(cboxSRC, l))
		Next l
		
		
	End Sub
	
	Public Sub PREP_SZM0010()
		
		
		'   Schema���̎擾  SZM0010
		MKKCMN.ZAEV_FNO = "SZM0010"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0010_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		''''SZM0010_FILE.NAME = ""
		
		
		'SELECT LOCK
		SQL = "Select  "
		
		SQL = SQL & "Inc_code," '��к���
		SQL = SQL & "jg_code," '���Ə�����"
		SQL = SQL & "hin_code," '�i��"
		SQL = SQL & "hin_name," '�i��"
		SQL = SQL & "kikaku," '�K�i"
		SQL = SQL & "tani," '�P��"
		SQL = SQL & "jan_code," 'JAN�W������"
		SQL = SQL & "jan_s_code," 'JAN�Z�k"
		SQL = SQL & "bar_code," '���̑����ް����"
		SQL = SQL & "teki_date1," '�����E�v���P"
		SQL = SQL & "baika1," '�����P"
		SQL = SQL & "kei_kin1," '�_�񉿊i1"
		SQL = SQL & "teki_date2," '�����E�v���Q"
		SQL = SQL & "baika2," '�����Q"
		SQL = SQL & "kei_kin2," '�_�񉿊i�Q"
		SQL = SQL & "hiyou_k_code1," '��p�Ȗځi���v�f�j"
		SQL = SQL & "hiyou_k_code2," '��p�Ȗځi���v�f�j"
		SQL = SQL & "ka_bun_code," '�Ȗڕ���"
		SQL = SQL & "l_bun_code," '�啪��"
		SQL = SQL & "m_bun_code," '������"
		SQL = SQL & "s_bun_code," '������"
		SQL = SQL & "bun_code," '����"
		SQL = SQL & "ken_bun_code," '��������"
		SQL = SQL & "jutaku," '������i"
		SQL = SQL & "sikakari," '�d�|�敪
		SQL = SQL & "zan," 'ϲŽ�c����"
		SQL = SQL & "zaiko," '�݌ɊǗ�"
		SQL = SQL & "kanri_kubn," '�Ǘ��敪"
		SQL = SQL & "tax_kubn," '����ŋ敪"
		SQL = SQL & "tana_tanka," '�I���P���敪"
		SQL = SQL & "ha_tanka1," '�����P��1"
		SQL = SQL & "kansan_num1," '���Z��1"
		SQL = SQL & "ha_tanka2," '�����P��2"
		SQL = SQL & "kansan_num2," '���Z��2"
		SQL = SQL & "ha_tanka3," '�����P��3"
		SQL = SQL & "kansan_num3," '���Z��3"
		SQL = SQL & "ha_tanka4," '�����P��4"
		SQL = SQL & "kansan_num4," '���Z��4"
		SQL = SQL & "ha_tanka5," '�����P��5"
		SQL = SQL & "kansan_num5," '���Z��5"
		SQL = SQL & "g_gentei_code," '�ƎҌ��躰��"
		SQL = SQL & "gen_h_ka," '���ꔭ����"
		SQL = SQL & "Fax_yn," 'Fax���M�s��"
		SQL = SQL & "tax_rate_kbn," '�ŗ��敪"
		SQL = SQL & "tyozouhin," '�����i"
		SQL = SQL & "jihan," '���̋@�̔�"
		SQL = SQL & "gensen," '����Ώ�"
		SQL = SQL & "nouhin_date," '�ŏI�[�i��"
		SQL = SQL & "tekiyo_date," '�K�p�J�n���t"
		SQL = SQL & "tori_kyu," '���x�~"
		SQL = SQL & "tori_kyu_date," '���x�~���t"
		SQL = SQL & "Entry_Op_code," '�o�^�I�y���[�^"
		SQL = SQL & "Entry_Op_date," '�o�^�n��_date"
		SQL = SQL & "Entry_Op_time," '�o�^�n��_time"
		SQL = SQL & "Edit_Op_code," '�C���I�y���[�^"
		SQL = SQL & "Edit_Op_date," '�C���n��_date"
		SQL = SQL & "Edit_Op_time," '�C���n��_time"
		SQL = SQL & "del_flg " '�폜�t���O
		SQL = SQL & ",hin_name_seisiki" '��������           A-CUST-20100610
		'A-CUST20130212��
		SQL = SQL & ",BK1" 'JAN���i���ރR�[�h
		SQL = SQL & ",K42" '�P�i�d��
		SQL = SQL & ",K44" '���Y���R�[�h
		SQL = SQL & ",K57" '���Y���R�[�h
		SQL = SQL & ",K58" '�L������
		SQL = SQL & ",K99" '�L������ �����Z
		'A-CUST20130212��
		SQL = SQL & ",Shomi_date_kbn" '����/�ܖ������敪   A-20240115
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0010_FILE.NAME) & "SZM0010 "
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		SQL = SQL & " AND hin_code = ? "
		SQL = SQL & " for UPDATE NOWAIT"
		
		On Error Resume Next
		qSZM0010SEL = ZACN_RCN.CreateQuery("qSZM0010SEL", SQL)
		qSZM0010SEL.QueryTimeout = ZACN_TIME
		''''qSZM0010SEL.LockType = rdConcurLock '   '���b�N�^�C�v���u�r���v�ɐݒ� rdConcurReadOnly
		''''qSZM0010SEL.LockType = rdConcurRowver
		If B_STATUS <> 0 Then
			'   ZAER_NO = "SZM0010"
			MsgBox("qSZM0010SEL CreateQuery Error")
			
			GoTo PREP_SZM0010_ERR
		End If
		On Error GoTo 0
		
		qSZM0010SEL.rdoParameters(0).NAME = "Inc_code"
		qSZM0010SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010SEL.rdoParameters(0).Size = 2
		qSZM0010SEL.rdoParameters(1).NAME = "jg_code"
		qSZM0010SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010SEL.rdoParameters(1).Size = 4
		qSZM0010SEL.rdoParameters(2).NAME = "hin_code"
		qSZM0010SEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010SEL.rdoParameters(2).Size = 5
		
		
		'   �i�ڃ}�X�^��QUERY�쐬(INSERT)
		Call PREP_SZM0010INS()
		'
		'   �i�ڃ}�X�^��QUERY�쐬(UPDATE)
		Call PREP_SZM0010UPD()
		
		'   �i�ڃ}�X�^��QUERY�쐬(DELETE)
		Call PREP_SZM0010DEL()
		
		'   �i�ں��ގ擾
		'D-CUST-20100610 Start
		'SQL = ""
		'SQL = SQL + "select inc_code,jg_code,max(hin_code) as maxnum "
		'SQL = SQL + " from szm0010 "
		'SQL = SQL + " where inc_code = ?"
		'SQL = SQL + "   and jg_code  = ?"
		'SQL = SQL + " group by inc_code,jg_code"
		'D-CUST-20100610 End
		'A-CUST-20100610 Start
		SQL = "SELECT MIN(HIN_CODE) AS maxnum "
		SQL = SQL & "FROM("
		SQL = SQL & "SELECT MIN(HIN_CODE) AS HIN_CODE "
		SQL = SQL & "FROM SZM0010 A "
		SQL = SQL & "WHERE A.INC_CODE = ?"
		SQL = SQL & "  AND A.JG_CODE = ?"
		SQL = SQL & "  AND NOT EXISTS("
		SQL = SQL & "SELECT * FROM SZM0010 B "
		SQL = SQL & "Where B.INC_CODE = A.INC_CODE"
		SQL = SQL & "  AND B.JG_CODE = A.JG_CODE"
		SQL = SQL & "  AND B.HIN_CODE = TO_CHAR(TO_NUMBER(A.HIN_CODE) + 1, 'FM00000')) "
		SQL = SQL & "Union "
		SQL = SQL & "SELECT DECODE(NVL(MIN(HIN_CODE), '00000'), '00000', '00000', '00001','99999', '00000') AS HIN_CODE "
		SQL = SQL & "FROM SZM0010 C "
		SQL = SQL & "WHERE C.INC_CODE = ?"
		SQL = SQL & "  AND C.JG_CODE = ?"
		SQL = SQL & ") DAT"
		'A-CUST-20100610 End
		On Error Resume Next
		qSZM0010_NSEL = ZACN_RCN.CreateQuery("qSZM0010_NSEL", SQL)
		qSZM0010_NSEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			MsgBox("qSZM0010_NSEL CreateQuery Error")
			
			GoTo PREP_SZM0010_ERR
		End If
		On Error GoTo 0
		
		qSZM0010_NSEL.rdoParameters(0).NAME = "Inc_code"
		qSZM0010_NSEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010_NSEL.rdoParameters(0).Size = 2
		qSZM0010_NSEL.rdoParameters(1).NAME = "jg_code"
		qSZM0010_NSEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010_NSEL.rdoParameters(1).Size = 4
		'A-CUST-20100610 Start
		qSZM0010_NSEL.rdoParameters(2).NAME = "Inc_code2"
		qSZM0010_NSEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010_NSEL.rdoParameters(2).Size = 2
		qSZM0010_NSEL.rdoParameters(3).NAME = "jg_code2"
		qSZM0010_NSEL.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0010_NSEL.rdoParameters(3).Size = 4
		'A-CUST-20100610 End
		
		'A 050909 TIO NAGANO----------------------------------------------START
		'��p�Ή��Ȗڃ`�F�b�N
		SQL = ""
		SQL = SQL & "select hi_code1,hi_code2 "
		SQL = SQL & " from szm0170 "
		SQL = SQL & " where inc_code = ?"
		SQL = SQL & "   and jg_code  = ?"
		SQL = SQL & "   and hi_code1  = ?"
		SQL = SQL & "   and hi_code2  = ?"
		On Error Resume Next
		SZM0170_SEL = ZACN_RCN.CreateQuery("SZM0170_SEL", SQL)
		SZM0170_SEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			MsgBox("SZM0170_SEL CreateQuery Error")
			GoTo PREP_SZM0010_ERR
		End If
		On Error GoTo 0
		SZM0170_SEL.rdoParameters(0).NAME = "Inc_code"
		SZM0170_SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170_SEL.rdoParameters(0).Size = 2
		SZM0170_SEL.rdoParameters(1).NAME = "jg_code"
		SZM0170_SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170_SEL.rdoParameters(1).Size = 4
		SZM0170_SEL.rdoParameters(2).NAME = "hi_code1"
		SZM0170_SEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170_SEL.rdoParameters(2).Size = 3
		SZM0170_SEL.rdoParameters(3).NAME = "hi_code2"
		SZM0170_SEL.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0170_SEL.rdoParameters(3).Size = 6
		'A 050909 TOP NAGANO----------------------------------------------END
		
		Exit Sub
		
PREP_SZM0010_ERR: 
		'    ZAER_FID = "RAZ99"
		'    ZAER_KN = 1
		'    Call ZAER_SUB
		ERRSW = F_ERR
		On Error GoTo 0
		
	End Sub
	
	Private Sub PREP_SZM0010INS()
		
		'   �i�ڃ}�X�^��QUERY�쐬(INSERT)
		SQL = ""
		SQL = SQL & "INSERT INTO "
		SQL = SQL & RTrim(SZM0010_FILE.NAME) & "SZM0010("
		SQL = SQL & "Inc_code," ' * 2    '��к���
		SQL = SQL & "jg_code," ' * 4    '���Ə�����
		SQL = SQL & "hin_code," ' * 5    '�i��
		SQL = SQL & "hin_name," ' * 20   '�i��
		SQL = SQL & "kikaku," ' * 20   '�K�i
		SQL = SQL & "tani," ' * 4    '�P��
		SQL = SQL & "jan_code," ' * 13   'JAN�W������
		SQL = SQL & "jan_s_code," ' * 7    'JAN�Z�k
		SQL = SQL & "bar_code," ' * 30   '���̑����ް����
		SQL = SQL & "teki_date1," ' * 8    '�����E�v���P
		SQL = SQL & "baika1," '          '�����P
		SQL = SQL & "kei_kin1," '      '�_�񉿊i1
		SQL = SQL & "teki_date2," ' * 8    '�����E�v���Q
		SQL = SQL & "baika2," '          '�����Q
		SQL = SQL & "kei_kin2," '      '�_�񉿊i�Q
		SQL = SQL & "hiyou_k_code1," ' * 3    '��p�Ȗځi���v�f�j
		SQL = SQL & "hiyou_k_code2," ' * 6    '��p�Ȗځi���v�f�j
		SQL = SQL & "ka_bun_code," ' * 7    '�Ȗڕ���
		SQL = SQL & "l_bun_code," ' * 4    '�啪��
		SQL = SQL & "m_bun_code," ' * 4    '������
		SQL = SQL & "s_bun_code," ' * 4    '������
		SQL = SQL & "bun_code," ' * 4    '����
		SQL = SQL & "ken_bun_code," ' * 4    '��������
		SQL = SQL & "jutaku," ' * 1    '������i
		SQL = SQL & "sikakari," ' * 1    '�d�|�敪
		SQL = SQL & "zan," ' * 1    'ϲŽ�c����
		SQL = SQL & "zaiko," ' * 1    '�݌ɊǗ�
		SQL = SQL & "kanri_kubn," ' * 1    '�Ǘ��敪
		SQL = SQL & "tax_kubn," ' * 1    '����ŋ敪
		SQL = SQL & "tana_tanka," ' * 1    '�I���P���敪
		SQL = SQL & "ha_tanka1," ' * 4    '�����P��1
		SQL = SQL & "kansan_num1," '      '���Z��1
		SQL = SQL & "ha_tanka2," ' * 4    '�����P��2
		SQL = SQL & "kansan_num2," '      '���Z��2
		SQL = SQL & "ha_tanka3," ' * 4    '�����P��3
		SQL = SQL & "kansan_num3," '      '���Z��3
		SQL = SQL & "ha_tanka4," ' * 4    '�����P��4
		SQL = SQL & "kansan_num4," '      '���Z��4
		SQL = SQL & "ha_tanka5," ' * 4    '�����P��5
		SQL = SQL & "kansan_num5," '      '���Z��5
		SQL = SQL & "g_gentei_code," ' * 6    '�ƎҌ��躰��
		SQL = SQL & "gen_h_ka," ' * 1    '���ꔭ����
		SQL = SQL & "Fax_yn," ' * 1    'Fax���M�s��
		SQL = SQL & "tax_rate_kbn," ' * 1    '�ŗ��敪
		SQL = SQL & "tyozouhin," ' * 1    '�����i
		SQL = SQL & "jihan," ' * 1    '���̋@�̔�
		SQL = SQL & "gensen," ' * 1    '����Ώ�
		SQL = SQL & "nouhin_date," ' * 8    '�ŏI�[�i��
		SQL = SQL & "tekiyo_date," ' * 8    '�E�v�J�n���t
		SQL = SQL & "tori_kyu," ' * 1    '���x�~
		SQL = SQL & "tori_kyu_date," ' * 8    '���x�~���t
		SQL = SQL & "Entry_Op_code," ' * 6    '�o�^�I�y���[�^
		SQL = SQL & "Entry_Op_date," ' * 8    '�o�^�n��_date
		SQL = SQL & "Entry_Op_time," ' * 6    '�o�^�n��_time
		SQL = SQL & "Edit_Op_code," ' * 6    '�C���I�y���[�^
		SQL = SQL & "Edit_Op_date," ' * 8    '�C���n��_date
		SQL = SQL & "Edit_Op_time," ' * 6    '�C���n��_time
		'SQL = SQL & "del_flg) "                 ' * 1    '�폜�t���O           D-CUST-20100610
		SQL = SQL & "del_flg," ' * 1    '�폜�t���O
		'    SQL = SQL & "hin_name_seisiki)"         ' * 40   '��������              A-CUST-20100610 'D-CUST20130212
		SQL = SQL & "hin_name_seisiki," ' * 40   '��������              'A-CUST20130212
		'A-CUST20130212��
		SQL = SQL & "BK1, " ' * 6    'JAN���i���ރR�[�h
		SQL = SQL & "K42 " ' �@     '�P�i�d��
		SQL = SQL & ",K44 " ' * 3    '���Y���R�[�h
		SQL = SQL & ",K57 " ' * 1 '���Y���R�[�h
		SQL = SQL & ",K58 " '        '�L������
		'SQL = SQL & ",K99 )"                    '        '�L������ �����Z       'D-20240115
		'A-CUST20130212��
		SQL = SQL & ",K99 " '�L������ �����Z    A-20240115
		SQL = SQL & ",Shomi_date_kbn)" '����/�ܖ������敪  A-20240115
		SQL = SQL & "Values("
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?,?,?"
		SQL = SQL & ",?" 'A-CUST-20100610
		SQL = SQL & ",?,?,?,?,?,?" 'A-CUST20130212
		SQL = SQL & ",?" 'A-20240115
		SQL = SQL & ") "
		
		On Error Resume Next
		SZM0010INS = ZACN_RCN.CreateQuery("SZM0010INS", SQL)
		SZM0010INS.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		
		SZM0010INS.rdoParameters(0).NAME = "Inc_code"
		SZM0010INS.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(0).Size = 2
		SZM0010INS.rdoParameters(1).NAME = "jg_code"
		SZM0010INS.rdoParameters(1).Size = 4
		SZM0010INS.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		
		SZM0010INS.rdoParameters(2).NAME = "hin_code" '�i��"
		SZM0010INS.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(2).Size = 5
		SZM0010INS.rdoParameters(3).NAME = "hin_name" '�i��"
		SZM0010INS.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(3).Size = 20
		SZM0010INS.rdoParameters(4).NAME = "kikaku" '�K�i"
		SZM0010INS.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(4).Size = 20
		SZM0010INS.rdoParameters(5).NAME = "tani" '�P��"
		SZM0010INS.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(5).Size = 4
		SZM0010INS.rdoParameters(6).NAME = "jan_code" 'JAN�W������"
		SZM0010INS.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(6).Size = 13
		SZM0010INS.rdoParameters(7).NAME = "jan_s_code" 'JAN�Z�k"
		SZM0010INS.rdoParameters(7).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(7).Size = 7
		SZM0010INS.rdoParameters(8).NAME = "bar_code" '���̑����ް����"
		SZM0010INS.rdoParameters(8).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(8).Size = 30
		SZM0010INS.rdoParameters(9).NAME = "teki_date1" '�����E�v���P"
		SZM0010INS.rdoParameters(9).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(9).Size = 8
		SZM0010INS.rdoParameters(10).NAME = "baika1" '�����P"
		SZM0010INS.rdoParameters(10).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(11).NAME = "kei_kin1" '�_�񉿊i1"
		SZM0010INS.rdoParameters(11).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(12).NAME = "teki_date2" '�����E�v���Q"
		SZM0010INS.rdoParameters(12).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(12).Size = 8
		SZM0010INS.rdoParameters(13).NAME = "baika2" '�����Q"
		SZM0010INS.rdoParameters(13).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(14).NAME = "kei_kin2" '�_�񉿊i�Q"
		SZM0010INS.rdoParameters(14).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(15).NAME = "hiyou_k_code1" '��p�Ȗځi���v�f�j"
		SZM0010INS.rdoParameters(15).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(15).Size = 3
		SZM0010INS.rdoParameters(16).NAME = "hiyou_k_code2" '��p�Ȗځi���v�f�j"
		SZM0010INS.rdoParameters(16).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(16).Size = 6
		SZM0010INS.rdoParameters(17).NAME = "ka_bun_code" '�Ȗڕ���"
		SZM0010INS.rdoParameters(17).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(17).Size = 7
		SZM0010INS.rdoParameters(18).NAME = "l_bun_code" '�啪��"
		SZM0010INS.rdoParameters(18).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(18).Size = 4
		SZM0010INS.rdoParameters(19).NAME = "m_bun_code" '������"
		SZM0010INS.rdoParameters(19).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(19).Size = 4
		SZM0010INS.rdoParameters(20).NAME = "s_bun_code" '������"
		SZM0010INS.rdoParameters(20).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(20).Size = 4
		SZM0010INS.rdoParameters(21).NAME = "bun_code" '����"
		SZM0010INS.rdoParameters(21).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(21).Size = 4
		SZM0010INS.rdoParameters(22).NAME = "ken_bun_code" '��������"
		SZM0010INS.rdoParameters(22).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(22).Size = 4
		SZM0010INS.rdoParameters(23).NAME = "jutaku" '������i"
		SZM0010INS.rdoParameters(23).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(23).Size = 1
		SZM0010INS.rdoParameters(24).NAME = "sikakari" '�d�|�敪"
		SZM0010INS.rdoParameters(24).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(24).Size = 1
		SZM0010INS.rdoParameters(25).NAME = "zan" 'ϲŽ�c����"
		SZM0010INS.rdoParameters(25).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(25).Size = 1
		SZM0010INS.rdoParameters(26).NAME = "zaiko" '�݌ɊǗ�"
		SZM0010INS.rdoParameters(26).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(26).Size = 1
		SZM0010INS.rdoParameters(27).NAME = "kanri_kubn" '�Ǘ��敪"
		SZM0010INS.rdoParameters(27).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(27).Size = 1
		SZM0010INS.rdoParameters(28).NAME = "tax_kubn" '����ŋ敪"
		SZM0010INS.rdoParameters(28).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(28).Size = 1
		SZM0010INS.rdoParameters(29).NAME = "tana_tanka" '�I���P���敪"
		SZM0010INS.rdoParameters(29).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(29).Size = 1
		SZM0010INS.rdoParameters(30).NAME = "ha_tanka1" '�����P��1"
		SZM0010INS.rdoParameters(30).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(30).Size = 4
		SZM0010INS.rdoParameters(31).NAME = "kansan_num1" '���Z��1"
		SZM0010INS.rdoParameters(31).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(32).NAME = "ha_tanka2" '�����P��2"
		SZM0010INS.rdoParameters(32).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(32).Size = 4
		SZM0010INS.rdoParameters(33).NAME = "kansan_num2" '���Z��2"
		SZM0010INS.rdoParameters(33).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(34).NAME = "ha_tanka3" '�����P��3"
		SZM0010INS.rdoParameters(34).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(34).Size = 4
		SZM0010INS.rdoParameters(35).NAME = "kansan_num3" '���Z��3"
		SZM0010INS.rdoParameters(35).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(36).NAME = "ha_tanka4" '�����P��4"
		SZM0010INS.rdoParameters(36).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(36).Size = 4
		SZM0010INS.rdoParameters(37).NAME = "kansan_num4" '���Z��4"
		SZM0010INS.rdoParameters(37).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(38).NAME = "ha_tanka5" '�����P��5"
		SZM0010INS.rdoParameters(38).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(38).Size = 4
		SZM0010INS.rdoParameters(39).NAME = "kansan_num5" '���Z��5"
		SZM0010INS.rdoParameters(39).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010INS.rdoParameters(40).NAME = "g_gentei_code" '�ƎҌ��躰��"
		SZM0010INS.rdoParameters(40).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(40).Size = 6
		SZM0010INS.rdoParameters(41).NAME = "gen_h_ka" '���ꔭ����"
		SZM0010INS.rdoParameters(41).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(41).Size = 1
		SZM0010INS.rdoParameters(42).NAME = "Fax_yn" 'Fax���M�s��"
		SZM0010INS.rdoParameters(42).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(42).Size = 1
		SZM0010INS.rdoParameters(43).NAME = "tax_rate_kbn" '�ŗ��敪"
		SZM0010INS.rdoParameters(43).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(43).Size = 1
		SZM0010INS.rdoParameters(44).NAME = "tyozouhin" '�����i"
		SZM0010INS.rdoParameters(44).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(44).Size = 1
		SZM0010INS.rdoParameters(45).NAME = "jihan" '���̋@�̔�"
		SZM0010INS.rdoParameters(45).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(45).Size = 1
		SZM0010INS.rdoParameters(46).NAME = "gensen" '����Ώ�"
		SZM0010INS.rdoParameters(46).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(46).Size = 1
		SZM0010INS.rdoParameters(47).NAME = "nouhin_date" '�ŏI�[�i��"
		SZM0010INS.rdoParameters(47).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(47).Size = 8
		SZM0010INS.rdoParameters(48).NAME = "tekiyo_date" '�E�v�J�n���t"
		SZM0010INS.rdoParameters(48).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(48).Size = 8
		SZM0010INS.rdoParameters(49).NAME = "tori_kyu" '���x�~"
		SZM0010INS.rdoParameters(49).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(49).Size = 1
		SZM0010INS.rdoParameters(50).NAME = "tori_kyu_date" '���x�~���t"
		SZM0010INS.rdoParameters(50).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(50).Size = 8
		SZM0010INS.rdoParameters(51).NAME = "Entry_Op_code" '�o�^�I�y���[�^"
		SZM0010INS.rdoParameters(51).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(51).Size = 6
		SZM0010INS.rdoParameters(52).NAME = "Entry_Op_date" '�o�^�n��_date"
		SZM0010INS.rdoParameters(52).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(52).Size = 8
		SZM0010INS.rdoParameters(53).NAME = "Entry_Op_time" '�o�^�n��_time"
		SZM0010INS.rdoParameters(53).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(53).Size = 6
		SZM0010INS.rdoParameters(54).NAME = "Edit_Op_code" '�C���I�y���[�^"
		SZM0010INS.rdoParameters(54).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(54).Size = 6
		SZM0010INS.rdoParameters(55).NAME = "Edit_Op_date" '�C���n��_date"
		SZM0010INS.rdoParameters(55).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(55).Size = 8
		SZM0010INS.rdoParameters(56).NAME = "Edit_Op_time" '�C���n��_time"
		SZM0010INS.rdoParameters(56).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(56).Size = 6
		SZM0010INS.rdoParameters(57).NAME = "del_flg" '�폜�t���O"
		SZM0010INS.rdoParameters(57).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(57).Size = 1
		'A-CUST-20100610 Start
		SZM0010INS.rdoParameters(58).NAME = "hin_name_seisiki" '��������
		SZM0010INS.rdoParameters(58).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		SZM0010INS.rdoParameters(58).Size = 40
		'A-CUST-20100610 End
		'A-CUST20130212��
		SZM0010INS.rdoParameters(59).NAME = "BK1" 'JAN���i���ރR�[�h
		SZM0010INS.rdoParameters(59).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(59).Size = 6
		
		SZM0010INS.rdoParameters(60).NAME = "K42" '�P�i�d��
		SZM0010INS.rdoParameters(60).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		
		SZM0010INS.rdoParameters(61).NAME = "K44" '���Y���R�[�h
		SZM0010INS.rdoParameters(61).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(61).Size = 3
		
		SZM0010INS.rdoParameters(62).NAME = "K57" '�L������ �敪
		SZM0010INS.rdoParameters(62).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(62).Size = 1
		
		SZM0010INS.rdoParameters(63).NAME = "K58" '�L������
		SZM0010INS.rdoParameters(63).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		
		SZM0010INS.rdoParameters(64).NAME = "K99" '�L������ �����Z
		SZM0010INS.rdoParameters(64).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		'A-CUST20130212��
		
		'A-20240115��
		SZM0010INS.rdoParameters(65).NAME = "Shomi_date_kbn" '����/�ܖ������敪
		SZM0010INS.rdoParameters(65).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010INS.rdoParameters(65).Size = 1
		'A-20240115��
		
	End Sub
	Private Sub PREP_SZM0010UPD()
		
		'   �i�ڃ}�X�^��QUERY�쐬(UPDATE)
		SQL = ""
		SQL = SQL & "UPDATE  "
		SQL = SQL & RTrim(SZM0010_FILE.NAME) & "SZM0010 SET "
		SQL = SQL & "hin_name = ?," ' * 20   '�i��
		SQL = SQL & "kikaku = ?," ' * 20   '�K�i
		SQL = SQL & "tani = ?," ' * 4    '�P��
		SQL = SQL & "jan_code = ?," ' * 13   'JAN�W������
		SQL = SQL & "jan_s_code = ?," ' * 7    'JAN�Z�k
		SQL = SQL & "bar_code = ?," ' * 30   '���̑����ް����
		SQL = SQL & "teki_date1 = ?," ' * 8    '�����E�v���P
		SQL = SQL & "baika1 = ?," '          '�����P
		SQL = SQL & "kei_kin1 = ?," '      '�_�񉿊i1
		SQL = SQL & "teki_date2 = ?," ' * 8    '�����E�v���Q
		SQL = SQL & "baika2 = ?," '          '�����Q
		SQL = SQL & "kei_kin2 = ?," '      '�_�񉿊i�Q
		SQL = SQL & "hiyou_k_code1 = ?," ' * 3    '��p�Ȗځi���v�f�j
		SQL = SQL & "hiyou_k_code2 = ?," ' * 6    '��p�Ȗځi���v�f�j
		SQL = SQL & "ka_bun_code = ?," ' * 7    '�Ȗڕ���
		SQL = SQL & "l_bun_code = ?," ' * 4    '�啪��
		SQL = SQL & "m_bun_code = ?," ' * 4    '������
		SQL = SQL & "s_bun_code = ?," ' * 4    '������
		SQL = SQL & "bun_code = ?," ' * 4    '����
		SQL = SQL & "ken_bun_code = ?," ' * 4    '��������
		SQL = SQL & "jutaku = ?," ' * 1    '������i
		SQL = SQL & "sikakari = ?," ' * 1    '�d�|�敪
		SQL = SQL & "zan = ?," ' * 1    'ϲŽ�c����
		SQL = SQL & "zaiko = ?," ' * 1    '�݌ɊǗ�
		SQL = SQL & "kanri_kubn = ?," ' * 1    '�Ǘ��敪
		SQL = SQL & "tax_kubn = ?," ' * 1    '����ŋ敪
		SQL = SQL & "tana_tanka = ?," ' * 1    '�I���P���敪
		SQL = SQL & "ha_tanka1 = ?," ' * 4    '�����P��1
		SQL = SQL & "kansan_num1 = ?," '      '���Z��1
		SQL = SQL & "ha_tanka2 = ?," ' * 4    '�����P��2
		SQL = SQL & "kansan_num2 = ?," '      '���Z��2
		SQL = SQL & "ha_tanka3 = ?," ' * 4    '�����P��3
		SQL = SQL & "kansan_num3 = ?," '      '���Z��3
		SQL = SQL & "ha_tanka4 = ?," ' * 4    '�����P��4
		SQL = SQL & "kansan_num4 = ?," '      '���Z��4
		SQL = SQL & "ha_tanka5 = ?," ' * 4    '�����P��5
		SQL = SQL & "kansan_num5 = ?," '      '���Z��5
		SQL = SQL & "g_gentei_code = ?," ' * 6    '�ƎҌ��躰��
		SQL = SQL & "gen_h_ka = ?," ' * 1    '���ꔭ����
		SQL = SQL & "Fax_yn = ?," ' * 1    'Fax���M�s��
		SQL = SQL & "tax_rate_kbn = ?," ' * 1    '�ŗ��敪
		SQL = SQL & "tyozouhin = ?," ' * 1    '�����i
		SQL = SQL & "jihan = ?," ' * 1    '���̋@�̔�
		SQL = SQL & "gensen = ?," ' * 1    '����Ώ�
		SQL = SQL & "nouhin_date = ?," ' * 8    '�ŏI�[�i��
		SQL = SQL & "tekiyo_date = ?," ' * 8    '�E�v�J�n���t
		SQL = SQL & "tori_kyu = ?," ' * 1    '���x�~
		SQL = SQL & "tori_kyu_date = ?," ' * 8    '���x�~���t
		SQL = SQL & "Entry_Op_code = ?," ' * 6    '�o�^�I�y���[�^
		SQL = SQL & "Entry_Op_date = ?," ' * 8    '�o�^�n��_date
		SQL = SQL & "Entry_Op_time = ?," ' * 6    '�o�^�n��_time
		SQL = SQL & "Edit_Op_code = ?," ' * 6    '�C���I�y���[�^
		SQL = SQL & "Edit_Op_date = ?," ' * 8    '�C���n��_date
		SQL = SQL & "Edit_Op_time = ?," ' * 6    '�C���n��_time
		SQL = SQL & "del_flg = ? " ' * 1    '�폜�t���O
		SQL = SQL & ",hin_name_seisiki = ? " ' * 40   '��������
		'A-CUST20130212��
		SQL = SQL & ",BK1 = ? " ' * 6    'JAN���i���ރR�[�h
		SQL = SQL & ",K42 = ? " ' �@     '�P�i�d��
		SQL = SQL & ",K44 = ? " ' * 3    '���Y���R�[�h
		SQL = SQL & ",K57 = ? " ' * 1 '���Y���R�[�h
		SQL = SQL & ",K58 = ? " '        '�L������
		SQL = SQL & ",K99 = ? " '        '�L������ �����Z
		'A-CUST20130212��
		SQL = SQL & ",Shomi_date_kbn = ? " '����/�ܖ������敪   A-20240115
		SQL = SQL & "WHERE Inc_code  = ? "
		SQL = SQL & "  AND jg_code  = ? "
		SQL = SQL & "  AND hin_code  = ? "
		
		On Error Resume Next
		SZM0010UPD = ZACN_RCN.CreateQuery("SZM0010UPD", SQL)
		SZM0010UPD.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010UPD"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		
		SZM0010UPD.rdoParameters(0).NAME = "hin_name" '�i��"
		SZM0010UPD.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(0).Size = 20
		SZM0010UPD.rdoParameters(1).NAME = "kikaku" '�K�i"
		SZM0010UPD.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(1).Size = 20
		SZM0010UPD.rdoParameters(2).NAME = "tani" '�P��"
		SZM0010UPD.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(2).Size = 4
		SZM0010UPD.rdoParameters(3).NAME = "jan_code" 'JAN�W������"
		SZM0010UPD.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(3).Size = 13
		SZM0010UPD.rdoParameters(4).NAME = "jan_s_code" 'JAN�Z�k"
		SZM0010UPD.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(4).Size = 7
		SZM0010UPD.rdoParameters(5).NAME = "bar_code" '���̑����ް����"
		SZM0010UPD.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(5).Size = 30
		SZM0010UPD.rdoParameters(6).NAME = "teki_date1" '�����E�v���P"
		SZM0010UPD.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(6).Size = 8
		SZM0010UPD.rdoParameters(7).NAME = "baika1" '�����P"
		SZM0010UPD.rdoParameters(7).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(8).NAME = "kei_kin1" '�_�񉿊i1"
		SZM0010UPD.rdoParameters(8).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(9).NAME = "teki_date2" '�����E�v���Q"
		SZM0010UPD.rdoParameters(9).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(9).Size = 8
		SZM0010UPD.rdoParameters(10).NAME = "baika2" '�����Q"
		SZM0010UPD.rdoParameters(10).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(11).NAME = "kei_kin2" '�_�񉿊i�Q"
		SZM0010UPD.rdoParameters(11).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(12).NAME = "hiyou_k_code1" '��p�Ȗځi���v�f�j"
		SZM0010UPD.rdoParameters(12).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(12).Size = 3
		SZM0010UPD.rdoParameters(13).NAME = "hiyou_k_code2" '��p�Ȗځi���v�f�j"
		SZM0010UPD.rdoParameters(13).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(13).Size = 6
		SZM0010UPD.rdoParameters(14).NAME = "ka_bun_code" '�Ȗڕ���"
		SZM0010UPD.rdoParameters(14).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(14).Size = 7
		SZM0010UPD.rdoParameters(15).NAME = "l_bun_code" '�啪��"
		SZM0010UPD.rdoParameters(15).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(15).Size = 4
		SZM0010UPD.rdoParameters(16).NAME = "m_bun_code" '������"
		SZM0010UPD.rdoParameters(16).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(16).Size = 4
		SZM0010UPD.rdoParameters(17).NAME = "s_bun_code" '������"
		SZM0010UPD.rdoParameters(17).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(17).Size = 4
		SZM0010UPD.rdoParameters(18).NAME = "bun_code" '����"
		SZM0010UPD.rdoParameters(18).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(18).Size = 4
		SZM0010UPD.rdoParameters(19).NAME = "ken_bun_code" '��������"
		SZM0010UPD.rdoParameters(19).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(19).Size = 4
		SZM0010UPD.rdoParameters(20).NAME = "jutaku" '������i"
		SZM0010UPD.rdoParameters(20).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(20).Size = 1
		SZM0010UPD.rdoParameters(21).NAME = "sikakari" '�d�|�敪"
		SZM0010UPD.rdoParameters(21).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(21).Size = 1
		SZM0010UPD.rdoParameters(22).NAME = "zan" 'ϲŽ�c����"
		SZM0010UPD.rdoParameters(22).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(22).Size = 1
		SZM0010UPD.rdoParameters(23).NAME = "zaiko" '�݌ɊǗ�"
		SZM0010UPD.rdoParameters(23).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(23).Size = 1
		SZM0010UPD.rdoParameters(24).NAME = "kanri_kubn" '�Ǘ��敪"
		SZM0010UPD.rdoParameters(24).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(24).Size = 1
		SZM0010UPD.rdoParameters(25).NAME = "tax_kubn" '����ŋ敪"
		SZM0010UPD.rdoParameters(25).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(25).Size = 1
		SZM0010UPD.rdoParameters(26).NAME = "tana_tanka" '�I���P���敪"
		SZM0010UPD.rdoParameters(26).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(26).Size = 1
		SZM0010UPD.rdoParameters(27).NAME = "ha_tanka1" '�����P��1"
		SZM0010UPD.rdoParameters(27).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(27).Size = 4
		SZM0010UPD.rdoParameters(28).NAME = "kansan_num1" '���Z��1"
		SZM0010UPD.rdoParameters(28).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(29).NAME = "ha_tanka2" '�����P��2"
		SZM0010UPD.rdoParameters(29).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(29).Size = 4
		SZM0010UPD.rdoParameters(30).NAME = "kansan_num2" '���Z��2"
		SZM0010UPD.rdoParameters(30).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(31).NAME = "ha_tanka3" '�����P��3"
		SZM0010UPD.rdoParameters(31).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(31).Size = 4
		SZM0010UPD.rdoParameters(32).NAME = "kansan_num3" '���Z��3"
		SZM0010UPD.rdoParameters(32).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(33).NAME = "ha_tanka4" '�����P��4"
		SZM0010UPD.rdoParameters(33).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(33).Size = 4
		SZM0010UPD.rdoParameters(34).NAME = "kansan_num4" '���Z��4"
		SZM0010UPD.rdoParameters(34).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(35).NAME = "ha_tanka5" '�����P��5"
		SZM0010UPD.rdoParameters(35).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(35).Size = 4
		SZM0010UPD.rdoParameters(36).NAME = "kansan_num5" '���Z��5"
		SZM0010UPD.rdoParameters(36).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(37).NAME = "g_gentei_code" '�ƎҌ��躰��"
		SZM0010UPD.rdoParameters(37).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(37).Size = 6
		SZM0010UPD.rdoParameters(38).NAME = "gen_h_ka" '���ꔭ����"
		SZM0010UPD.rdoParameters(38).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(38).Size = 1
		SZM0010UPD.rdoParameters(39).NAME = "Fax_yn" 'Fax���M�s��"
		SZM0010UPD.rdoParameters(39).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(39).Size = 1
		SZM0010UPD.rdoParameters(40).NAME = "tax_rate_kbn" '�ŗ��敪"
		SZM0010UPD.rdoParameters(40).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(40).Size = 1
		SZM0010UPD.rdoParameters(41).NAME = "tyozouhin" '�����i"
		SZM0010UPD.rdoParameters(41).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(41).Size = 1
		SZM0010UPD.rdoParameters(42).NAME = "jihan" '���̋@�̔�"
		SZM0010UPD.rdoParameters(42).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(42).Size = 1
		SZM0010UPD.rdoParameters(43).NAME = "gensen" '����Ώ�"
		SZM0010UPD.rdoParameters(43).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(43).Size = 1
		SZM0010UPD.rdoParameters(44).NAME = "nouhin_date" '�ŏI�[�i��"
		SZM0010UPD.rdoParameters(44).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(44).Size = 8
		SZM0010UPD.rdoParameters(45).NAME = "tekiyo_date" '�E�v�J�n���t"
		SZM0010UPD.rdoParameters(45).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(45).Size = 8
		SZM0010UPD.rdoParameters(46).NAME = "tori_kyu" '���x�~"
		SZM0010UPD.rdoParameters(46).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(46).Size = 1
		SZM0010UPD.rdoParameters(47).NAME = "tori_kyu_date" '���x�~���t"
		SZM0010UPD.rdoParameters(47).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(47).Size = 8
		SZM0010UPD.rdoParameters(48).NAME = "Entry_Op_code" '�o�^�I�y���[�^"
		SZM0010UPD.rdoParameters(48).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(48).Size = 6
		SZM0010UPD.rdoParameters(49).NAME = "Entry_Op_date" '�o�^�n��_date"
		SZM0010UPD.rdoParameters(49).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(49).Size = 8
		SZM0010UPD.rdoParameters(50).NAME = "Entry_Op_time" '�o�^�n��_time"
		SZM0010UPD.rdoParameters(50).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(50).Size = 6
		SZM0010UPD.rdoParameters(51).NAME = "Edit_Op_code" '�C���I�y���[�^"
		SZM0010UPD.rdoParameters(51).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(51).Size = 6
		SZM0010UPD.rdoParameters(52).NAME = "Edit_Op_date" '�C���n��_date"
		SZM0010UPD.rdoParameters(52).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(52).Size = 8
		SZM0010UPD.rdoParameters(53).NAME = "Edit_Op_time" '�C���n��_time"
		SZM0010UPD.rdoParameters(53).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(53).Size = 6
		SZM0010UPD.rdoParameters(54).NAME = "del_flg" '�폜�t���O"
		SZM0010UPD.rdoParameters(54).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(54).Size = 1
		'D-CUST-20100610 Start
		'SZM0010UPD(55).NAME = "Inc_code"
		'SZM0010UPD(55).Type = rdTypeCHAR
		'SZM0010UPD(55).Size = 2
		'SZM0010UPD(56).NAME = "jg_code"
		'SZM0010UPD(56).Size = 4
		'SZM0010UPD(56).Type = rdTypeCHAR
		'SZM0010UPD(57).NAME = "hin_code"              '�i��"
		'SZM0010UPD(57).Type = rdTypeCHAR
		'SZM0010UPD(57).Size = 5
		'D-CUST-20100610 End
		'D-CUST-20100610 Start
		SZM0010UPD.rdoParameters(55).NAME = "hin_name_seisiki" '��������
		SZM0010UPD.rdoParameters(55).Type = RDO.DataTypeConstants.rdTypeVARCHAR
		SZM0010UPD.rdoParameters(55).Size = 40
		'D-CUST20130212��
		'    SZM0010UPD(56).NAME = "Inc_code"
		'    SZM0010UPD(56).Type = rdTypeCHAR
		'    SZM0010UPD(56).Size = 2
		'    SZM0010UPD(57).NAME = "jg_code"
		'    SZM0010UPD(57).Size = 4
		'    SZM0010UPD(57).Type = rdTypeCHAR
		'    SZM0010UPD(58).NAME = "hin_code"              '�i��"
		'    SZM0010UPD(58).Type = rdTypeCHAR
		'    SZM0010UPD(58).Size = 5
		'D-CUST20130212��
		'A-CUST20130212��
		SZM0010UPD.rdoParameters(56).NAME = "BK1" 'JAN���i���ރR�[�h
		SZM0010UPD.rdoParameters(56).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(56).Size = 6
		SZM0010UPD.rdoParameters(57).NAME = "K42" '�P�i�d��
		SZM0010UPD.rdoParameters(57).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(58).NAME = "K44" '���Y���R�[�h
		SZM0010UPD.rdoParameters(58).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(58).Size = 3
		SZM0010UPD.rdoParameters(59).NAME = "K57" '�L������ �敪
		SZM0010UPD.rdoParameters(59).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(59).Size = 1
		SZM0010UPD.rdoParameters(60).NAME = "K58" '�L������
		SZM0010UPD.rdoParameters(60).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0010UPD.rdoParameters(61).NAME = "K99" '�L������ �����Z
		SZM0010UPD.rdoParameters(61).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		
		'D-20240115��
		'SZM0010UPD(62).NAME = "Inc_code"
		'SZM0010UPD(62).Type = rdTypeCHAR
		'SZM0010UPD(62).Size = 2
		'SZM0010UPD(63).NAME = "jg_code"
		'SZM0010UPD(63).Size = 4
		'SZM0010UPD(63).Type = rdTypeCHAR
		'SZM0010UPD(64).NAME = "hin_code"              '�i��"
		'SZM0010UPD(64).Type = rdTypeCHAR
		'SZM0010UPD(64).Size = 5
		'D-20240115��
		'A-CUST20130212��
		'D-CUST-20100610 End
		
		'A-20240115��
		SZM0010UPD.rdoParameters(62).NAME = "Shomi_date_kbn" '����/�ܖ������敪
		SZM0010UPD.rdoParameters(62).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(62).Size = 1
		
		SZM0010UPD.rdoParameters(63).NAME = "Inc_code"
		SZM0010UPD.rdoParameters(63).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(63).Size = 2
		SZM0010UPD.rdoParameters(64).NAME = "jg_code"
		SZM0010UPD.rdoParameters(64).Size = 4
		SZM0010UPD.rdoParameters(64).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(65).NAME = "hin_code" '�i��"
		SZM0010UPD.rdoParameters(65).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010UPD.rdoParameters(65).Size = 5
		'A-20240115��
		
		'    Debug.Print Mid(SQL, 1, 100)
		'    Debug.Print Mid(SQL, 101, 100)
		'    Debug.Print Mid(SQL, 201, 100)
		'    Debug.Print Mid(SQL, 301, 100)
		'    Debug.Print Mid(SQL, 401, 100)
		'    Debug.Print Mid(SQL, 501, 100)
		'    Debug.Print Mid(SQL, 601, 100)
		'    Debug.Print Mid(SQL, 701, 100)
		
		
	End Sub
	
	Private Sub PREP_SZM0010DEL()
		
		'   �ƎҌ������ރ}�X�^��QUERY�쐬(DELETE)
		SQL = ""
		SQL = SQL & "UPDATE  "
		SQL = SQL & RTrim(SZM0010_FILE.NAME) & "SZM0010 "
		SQL = SQL & " SET del_flg = '1', "
		SQL = SQL & " Edit_Op_Code = ?, Edit_Op_Date = ?, Edit_Op_Time = ? "
		SQL = SQL & "WHERE Inc_code  = ? "
		SQL = SQL & "  AND jg_code  = ? "
		SQL = SQL & "  AND hin_code  = ? "
		
		On Error Resume Next
		SZM0010DEL = ZACN_RCN.CreateQuery("SZM0010DEL", SQL)
		SZM0010DEL.QueryTimeout = ZACN_TIME '�^�C���A�E�g���Ԃ��u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0010DEL"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		SZM0010DEL.rdoParameters(0).NAME = "DelOpCode"
		SZM0010DEL.rdoParameters(1).NAME = "DelOpDate"
		SZM0010DEL.rdoParameters(2).NAME = "DelOpTime"
		SZM0010DEL.rdoParameters(3).NAME = "Inc_code"
		SZM0010DEL.rdoParameters(4).NAME = "jg_code"
		SZM0010DEL.rdoParameters(5).NAME = "hin_code"
		SZM0010DEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0010DEL.rdoParameters(0).Size = 6
		SZM0010DEL.rdoParameters(1).Size = 8
		SZM0010DEL.rdoParameters(2).Size = 6
		SZM0010DEL.rdoParameters(3).Size = 2
		SZM0010DEL.rdoParameters(4).Size = 4
		SZM0010DEL.rdoParameters(5).Size = 5
		On Error GoTo 0
		
		
	End Sub
	
	
	Public Function FILGET_SZM0010(ByRef strKAISHA As String, ByRef strJIGYO As String, ByRef HINMOKU As String, ByRef bUF As SZM0010_S) As Short
		
		'   �i�ڃ}�X�^�ǂݍ���
		
		FILGET_SZM0010 = F_OFF
		
		Call DBRollbackTrans()
		Call DBBeginTrans()
		
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		qSZM0010SEL.rdoParameters("Inc_code").Value = strKAISHA
		qSZM0010SEL.rdoParameters("jg_code").Value = strJIGYO
		qSZM0010SEL.rdoParameters("hin_code").Value = HINMOKU
		
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 1, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 101, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 201, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 301, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 401, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 501, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 601, 100)
		'Debug.Print "FILGET"; Mid(qSZM0010SEL.SQL, 701, 100)
		'
		
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0010myRSSW <> "qSZM0010SEL" Or ReQue = False Then
			SZM0010myRS = qSZM0010SEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0010myRSSW = "qSZM0010SEL"
		Else
			SZM0010myRS.Requery()
		End If
		
		Select Case B_STATUS(SZM0010myRS) ' (SQL���s�ð���̕]��)
			Case 0
				If KBKBN = 2 Or KBKBN = 3 Then
					If SZM0010myRS.rdoColumns("del_flg").Value <> "0" Then
						FILGET_SZM0010 = F_END
						ENDSW = F_END
						Exit Function
					End If
				End If
				
				Call SZM0010CNV_SUB(bUF)
				''''strName = SZM0010myRS!find_name
			Case 24
				FILGET_SZM0010 = F_END
				ENDSW = F_END
				''''MsgBox "EOF", vbOKOnly, "FILSZM0010GET"
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
			Case -54 '   ���b�N
				FILGET_SZM0010 = F_END
				ZAER_CD = 201
				ZAER_NO.Value = "" 'A-CUST-20100610
				Call ZAER_SUB()
				ENDSW = F_END
				''''MsgBox "EOF", vbOKOnly, "FILSZM0010GET"
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
				
			Case Else
				FILGET_SZM0010 = F_END
				ENDSW = F_END
				ERRSW = F_ERR
				''''MsgBox "FILSZM0010_GET ERR"
				''''MsgBox "ERR", vbOKOnly, "FILSZM0010_GET"
				
				ZAER_KN = 1
				ZAER_NO.Value = "RSZM0010"
				Call ZAER_SUB()
				On Error GoTo 0 ' (�װ�ׯ�߉���)
				Exit Function
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	
	Private Sub SZM0010CNV_SUB(ByRef bUF As SZM0010_S)
		
		SZM0010.Inc_code = SZM0010myRS.rdoColumns("Inc_code").Value '��к���
		SZM0010.jg_code = SZM0010myRS.rdoColumns("jg_code").Value '���Ə�����
		SZM0010.hin_code = SZM0010myRS.rdoColumns("hin_code").Value '�i��
		SZM0010.hin_name = SZM0010myRS.rdoColumns("hin_name").Value '�i��
		SZM0010.kikaku = SZM0010myRS.rdoColumns("kikaku").Value '�K�i
		SZM0010.tani = SZM0010myRS.rdoColumns("tani").Value '�P��
		SZM0010.jan_code = SZM0010myRS.rdoColumns("jan_code").Value 'JAN�W������
		SZM0010.jan_s_code = SZM0010myRS.rdoColumns("jan_s_code").Value 'JAN�Z�k
		SZM0010.bar_code = SZM0010myRS.rdoColumns("bar_code").Value '���̑����ް����
		SZM0010.teki_date1 = SZM0010myRS.rdoColumns("teki_date1").Value '�����E�v���P
		SZM0010.baika1 = SZM0010myRS.rdoColumns("baika1").Value '�����P
		SZM0010.kei_kin1 = SZM0010myRS.rdoColumns("kei_kin1").Value '�_�񉿊i1
		SZM0010.teki_date2 = SZM0010myRS.rdoColumns("teki_date2").Value '�����E�v���Q
		SZM0010.baika2 = SZM0010myRS.rdoColumns("baika2").Value '�����Q
		SZM0010.kei_kin2 = SZM0010myRS.rdoColumns("kei_kin2").Value '�_�񉿊i�Q
		SZM0010.hiyou_k_code1 = SZM0010myRS.rdoColumns("hiyou_k_code1").Value '��p�Ȗځi���v�f�j
		SZM0010.hiyou_k_code2 = SZM0010myRS.rdoColumns("hiyou_k_code2").Value '��p�Ȗځi���v�f�j
		SZM0010.ka_bun_code = SZM0010myRS.rdoColumns("ka_bun_code").Value '�Ȗڕ���
		SZM0010.l_bun_code = SZM0010myRS.rdoColumns("l_bun_code").Value '�啪��
		SZM0010.m_bun_code = SZM0010myRS.rdoColumns("m_bun_code").Value '������
		SZM0010.s_bun_code = SZM0010myRS.rdoColumns("s_bun_code").Value '������
		SZM0010.bun_code = SZM0010myRS.rdoColumns("bun_code").Value '����
		SZM0010.ken_bun_code = SZM0010myRS.rdoColumns("ken_bun_code").Value '��������
		SZM0010.jutaku = SZM0010myRS.rdoColumns("jutaku").Value '������i
		SZM0010.sikakari = SZM0010myRS.rdoColumns("sikakari").Value '�d�|�敪
		SZM0010.zan = SZM0010myRS.rdoColumns("zan").Value 'ϲŽ�c����
		SZM0010.zaiko = SZM0010myRS.rdoColumns("zaiko").Value '�݌ɊǗ�
		SZM0010.kanri_kubn = SZM0010myRS.rdoColumns("kanri_kubn").Value '�Ǘ��敪
		SZM0010.Tax_kubn = SZM0010myRS.rdoColumns("Tax_kubn").Value '����ŋ敪
		SZM0010.tana_tanka = SZM0010myRS.rdoColumns("tana_tanka").Value '�I���P���敪
		SZM0010.ha_tanka1 = SZM0010myRS.rdoColumns("ha_tanka1").Value '�����P��1
		SZM0010.kansan_num1 = SZM0010myRS.rdoColumns("kansan_num1").Value '���Z��1
		SZM0010.ha_tanka2 = SZM0010myRS.rdoColumns("ha_tanka2").Value '�����P��2
		SZM0010.kansan_num2 = SZM0010myRS.rdoColumns("kansan_num2").Value '���Z��2
		SZM0010.ha_tanka3 = SZM0010myRS.rdoColumns("ha_tanka3").Value '�����P��3
		SZM0010.kansan_num3 = SZM0010myRS.rdoColumns("kansan_num3").Value '���Z��3
		SZM0010.ha_tanka4 = SZM0010myRS.rdoColumns("ha_tanka4").Value '�����P��4
		SZM0010.kansan_num4 = SZM0010myRS.rdoColumns("kansan_num4").Value '���Z��4
		SZM0010.ha_tanka5 = SZM0010myRS.rdoColumns("ha_tanka5").Value '�����P��5
		SZM0010.kansan_num5 = SZM0010myRS.rdoColumns("kansan_num5").Value '���Z��5
		SZM0010.g_gentei_code = SZM0010myRS.rdoColumns("g_gentei_code").Value '�ƎҌ��躰��
		SZM0010.gen_h_ka = SZM0010myRS.rdoColumns("gen_h_ka").Value '���ꔭ����
		SZM0010.Fax_yn = SZM0010myRS.rdoColumns("Fax_yn").Value 'Fax���M�s��
		SZM0010.tax_rate_kbn = SZM0010myRS.rdoColumns("tax_rate_kbn").Value '�ŗ��敪
		SZM0010.tyozouhin = SZM0010myRS.rdoColumns("tyozouhin").Value '�����i
		SZM0010.jihan = SZM0010myRS.rdoColumns("jihan").Value '���̋@�̔�
		SZM0010.gensen = SZM0010myRS.rdoColumns("gensen").Value '����Ώ�
		SZM0010.nouhin_date = SZM0010myRS.rdoColumns("nouhin_date").Value '�ŏI�[�i��
		SZM0010.tekiyo_date = SZM0010myRS.rdoColumns("tekiyo_date").Value '�K�p�J�n���t
		SZM0010.tori_kyu = SZM0010myRS.rdoColumns("tori_kyu").Value '���x�~
		SZM0010.tori_kyu_date = SZM0010myRS.rdoColumns("tori_kyu_date").Value '���x�~���t
		SZM0010.Entry_Op_code = SZM0010myRS.rdoColumns("Entry_Op_code").Value '�o�^�I�y���[�^
		SZM0010.Entry_Op_date = SZM0010myRS.rdoColumns("Entry_Op_date").Value '�o�^�n��_date
		SZM0010.Entry_Op_time = SZM0010myRS.rdoColumns("Entry_Op_time").Value '�o�^�n��_time
		SZM0010.Edit_Op_code = SZM0010myRS.rdoColumns("Edit_Op_code").Value '�C���I�y���[�^
		SZM0010.Edit_Op_date = SZM0010myRS.rdoColumns("Edit_Op_date").Value '�C���n��_date
		SZM0010.Edit_Op_time = SZM0010myRS.rdoColumns("Edit_Op_time").Value '�C���n��_time
		SZM0010.del_flg = SZM0010myRS.rdoColumns("del_flg").Value '�폜�t���O
		SZM0010.hin_name_seisiki = SZM0010myRS.rdoColumns("hin_name_seisiki").Value '��������       A-CUST-20100610
		'A-CUST20130212��
		SZM0010.BK1 = SZM0010myRS.rdoColumns("BK1").Value
		SZM0010.k42 = SZM0010myRS.rdoColumns("k42").Value
		SZM0010.k44 = SZM0010myRS.rdoColumns("k44").Value
		SZM0010.k57 = SZM0010myRS.rdoColumns("k57").Value
		SZM0010.k58 = SZM0010myRS.rdoColumns("k58").Value
		SZM0010.k99 = SZM0010myRS.rdoColumns("k99").Value
		'A-CUST20130212��
		SZM0010.Shomi_date_kbn = SZM0010myRS.rdoColumns("Shomi_date_kbn").Value '����/�ܖ������敪   A-20240115
		'UPGRADE_WARNING: �I�u�W�F�N�g bUF �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		bUF = SZM0010
		
	End Sub
	
	Public Sub SCR_ADDNEW()
		
		
		
		KB.Inc_code = WKB010 '   ��к���
		KB.jg_code = WKB020 '   ���Ə�����
		KB.hin_code = WKB030 '   �i��
		KB.hin_name = "" '   �i��
		KB.kikaku = "" '   �K�i
		KB.tani = "" '   �P��
		KB.jan_code = "" '   JAN�W������
		KB.jan_s_code = "" '   JAN�Z�k
		KB.bar_code = "" '   ���̑����ް����
		KB.teki_date1 = "" '   �����E�v���P
		KB.baika1 = 0 '   �����P
		KB.kei_kin1 = 0 '   �_�񉿊i1
		KB.teki_date2 = "" '   �����E�v���Q
		KB.baika2 = 0 '   �����Q
		KB.kei_kin2 = 0 '   �_�񉿊i�Q
		KB.hiyou_k_code1 = "" '   ��p�Ȗځi���v�f�j
		KB.hiyou_k_code2 = "" '   ��p�Ȗځi���v�f�j
		KB.ka_bun_code = "" '   �Ȗڕ���
		KB.l_bun_code = "" '   �啪��
		KB.m_bun_code = "" '   ������
		KB.s_bun_code = "" '   ������
		KB.bun_code = "" '   ����
		KB.ken_bun_code = "" '   ��������
		KB.jutaku = "0" '   ������i
		KB.sikakari = "0" '   �d�|�敪
		KB.zan = "0" '   ϲŽ�c����
		KB.zaiko = "1" '   �݌ɊǗ�
		WKB330 = 1 '   �݌ɊǗ�WKB330
		KB.kanri_kubn = "1" '   �Ǘ��敪
		WKB300 = 1 '   �Ǘ��敪WKB300
		KB.Tax_kubn = "1" '   ����ŋ敪
		WKB310 = 1 '   ����ŋ敪WKB310
		KB.tana_tanka = "1" '   �I���P���敪
		WKB320 = 1 '   �I���P���敪WKB320
		KB.ha_tanka1 = "" '   �����P��1
		KB.kansan_num1 = 0 '   ���Z��1
		KB.ha_tanka2 = "" '   �����P��2
		KB.kansan_num2 = 0 '   ���Z��2
		KB.ha_tanka3 = "" '   �����P��3
		KB.kansan_num3 = 0 '   ���Z��3
		KB.ha_tanka4 = "" '   �����P��4
		KB.kansan_num4 = 0 '   ���Z��4
		KB.ha_tanka5 = "" '   �����P��5
		KB.kansan_num5 = 0 '   ���Z��5
		KB.g_gentei_code = "" '   �ƎҌ��躰��
		KB.gen_h_ka = "0" '   ���ꔭ����
		KB.Fax_yn = "0" '   Fax���M�s��
		WKB340 = 1 '   Fax���M�s��WKB340
		KB.tax_rate_kbn = CStr(1) '   �ŗ��敪
		KB.tyozouhin = "0" '   �����i
		KB.jihan = "0" '   ���̋@�̔�
		KB.gensen = "0" '   ����Ώ�
		KB.nouhin_date = "" '   �ŏI�[�i��
		KB.tekiyo_date = "" '   �K�p�J�n���t
		KB.tori_kyu = "0" '   ���x�~
		KB.tori_kyu_date = "" '   ���x�~���t
		''''KB.Entry_Op_code = WG_OPCODE        '   �o�^�I�y���[�^
		KB.Entry_Op_code = Space(6) '   �o�^�I�y���[�^
		KB.Entry_Op_date = Space(8) '   �o�^�n��_date
		KB.Entry_Op_time = Space(6) '   �o�^�n��_time
		KB.Edit_Op_code = Space(6) '   �C���I�y���[�^
		KB.Edit_Op_date = Space(8) '   �C���n��_date
		KB.Edit_Op_time = Space(6) '   �C���n��_time
		KB.del_flg = CStr(0) '   �폜�t���O
		KB.hin_name_seisiki = "" '   �����i��            A-CUST-20100610
		'A-CUST20130212��
		KB.BK1 = ""
		KB.k42 = 0
		KB.k44 = ""
		KB.k57 = ""
		KB.k58 = 0
		KB.k99 = 0
		'A-CUST20130212��
		KB.Shomi_date_kbn = "0" '����/�ܖ������敪   A-20240115
		SentakuFLG = False
		
	End Sub
	
	Public Sub SCR_DSPDATA()
		
		
		Dim strKAISHA As String
		Dim strJIGYO As String
		
		Dim nmOper As String
		
		Dim nmCHU, nmDAI, nmSHO As String
		Dim rBUNRUI_NAME As String '02/05/28 ADD
		Dim iReturn As Short
		
		Dim KamUri As String
		Dim KamSho As String
		Dim KamMat As String
		Dim KamMit As String
		Dim strAcctName As String
		
		
		
		'   ��Ж��Ǝ��Ə���
		iReturn = CduDecodeKaisha(WKB010, strKAISHA)
		iReturn = CduDecodeJigyo(WKB010, WKB020, strJIGYO)
		
		'    Call SpreadInit
		With SZ0410FRM
			
			.IMTX010.Text = WKB010 'KB.Inc_code         '   ��к���
			.DSP010.Text = strKAISHA
			WKB010DSP = strKAISHA
			.IMTX020.Text = WKB020 'KB.jg_code          '   ���Ə�����
			.DSP020.Text = strJIGYO
			WKB020DSP = strJIGYO
			
			'   �o�^�A�C���I�y���[�^
			.DSP_OP0_CD.Text = KB.Entry_Op_code
			.DSP_OP1_CD.Text = KB.Edit_Op_code
			iReturn = CduDecodeOper(WKB010, KB.Entry_Op_code, nmOper)
			.DSP_OP0_NM.Text = nmOper
			iReturn = CduDecodeOper(WKB010, KB.Edit_Op_code, nmOper)
			.DSP_OP1_NM.Text = nmOper
			.DSP_OP0_DATE.Text = DateSlashed(KB.Entry_Op_date)
			.DSP_OP1_DATE.Text = DateSlashed(KB.Edit_Op_date)
			
			.IMTX030.Text = KB.hin_code '   �i��
			.IMTX040.Text = RTrim(KB.hin_name) '   �i��
			.IMTX050.Text = RTrim(KB.kikaku)
			'UPGRADE_ISSUE: ComboBox �v���p�e�B CMB060.DataField �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
			'.CMB060.DataField = KB.tani 'D-20250417
			.CMB060.DataSource = KB.tani 'A-20250417
			Call COMBO_SETLIST(.CMB060, KB.tani)
			.IMTX065.Text = RTrim(KB.hin_name_seisiki) 'A-CUST-20100610
			.IMTX070.Text = RTrim(KB.jan_code)
			.IMTX080.Text = RTrim(KB.jan_s_code)
			.IMTX090.Text = RTrim(KB.bar_code)
			
			.IMTX100(1).Text = DateSlashed(KB.teki_date1)
			.IMNU110(1).Value = KB.baika1
			.IMNU120(1).Value = KB.kei_kin1
			.IMTX100(2).Text = DateSlashed(KB.teki_date2)
			.IMNU110(2).Value = KB.baika2
			.IMNU120(2).Value = KB.kei_kin2
			.IMTX130(1).Text = KB.hiyou_k_code1
			.IMTX140(1).Text = KB.hiyou_k_code2
			'    .DSP140(1).Caption = DecodeKAMOKU(KB.hiyou_k_code1, _
			''                                      KB.hiyou_k_code2)
			'           2000/01/26  FIXED   KOKOKARA
			Call AccountName(KB.hiyou_k_code1 & KB.hiyou_k_code2, strAcctName)
			.DSP140(1).Text = strAcctName
			WKB140DSP = .DSP140(1).Text
			'           2000/01/26  FIXED   KOKOKARA
			
			'       �ȖڑΉ��e�[�u�����Q��
			
			iReturn = TaiouKamoku(WKB010, WKB020, KB.hiyou_k_code1, KB.hiyou_k_code2, KamUri, KamSho, KamMat, KamMit)
			
			Call AccountName(KamUri, strAcctName) '   ����Ȗږ��̎擾
			.IMTX130(2).Text = Mid(KamUri, 1, 3)
			.IMTX140(2).Text = Mid(KamUri, 4, 6)
			.DSP140(2).Text = strAcctName
			Call AccountName(KamSho, strAcctName)
			.IMTX130(3).Text = Mid(KamSho, 1, 3)
			.IMTX140(3).Text = Mid(KamSho, 4, 6)
			.DSP140(3).Text = strAcctName
			Call AccountName(KamMat, strAcctName)
			.IMTX130(4).Text = Mid(KamMat, 1, 3)
			.IMTX140(4).Text = Mid(KamMat, 4, 6)
			.DSP140(4).Text = strAcctName
			Call AccountName(KamMit, strAcctName)
			.IMTX130(5).Text = Mid(KamMit, 1, 3)
			.IMTX140(5).Text = Mid(KamMit, 4, 6)
			.DSP140(5).Text = strAcctName
			
			'A-20240115��
			If RTrim(KB.Shomi_date_kbn) = "" Then
				.CMB165.Text = "�����Ȃ�"
				KB.Shomi_date_kbn = "0"
			Else
				Select Case KB.Shomi_date_kbn
					Case CStr(0)
						.CMB165.Text = "�����Ȃ�"
					Case CStr(1)
						.CMB165.Text = "�������"
					Case CStr(2)
						.CMB165.Text = "�ܖ�����"
				End Select
			End If
			'A-20240115��
			
			'A-CUST20130212��
			'JAN�֘A����
			.IMTX150(0).Text = KB.k44
			.IMNU160(0).Value = KB.k42
			.IMNU170(1).Value = KB.k58
			If RTrim(KB.k57) = "" Then
				.CMB170.SelectedIndex = -1
			Else
				Select Case KB.k57
					Case CStr(1)
						.CMB170.Text = "��"
					Case CStr(2)
						.CMB170.Text = "��"
					Case CStr(3)
						.CMB170.Text = "�N"
				End Select
			End If
			.DSP170(0).Text = CStr(KB.k99)
			.IMTX291.Text = KB.BK1
			
			JAN_BUNRUI_BUF0.BK1 = KB.BK1
			If FILGET_JAN_BUNRUI() = True Then
				.DSP291.Text = JAN_BUNRUI.BK4 '���ޖ�
			Else
				.DSP291.Text = ""
			End If
			'A-CUST20130212��
			
			'   �Ȗڕ���
			.IMTX210.Text = Left(KB.ka_bun_code, 3)
			.IMTX211.Text = Right(KB.ka_bun_code, 4)
			.DSP210.Text = DecodeKamBunrui(WKB010, WKB020, KB.ka_bun_code)
			WKB210DSP = .DSP210.Text
			
			'   �啪�ށA�����ށA������
			.IMTX220.Text = KB.l_bun_code
			.IMTX230.Text = KB.m_bun_code
			.IMTX240.Text = KB.s_bun_code
			iReturn = CduDecodeDAIBunrui(WKB010, KB.l_bun_code, nmDAI)
			iReturn = CduDecodeCHUBunrui(WKB010, KB.l_bun_code, KB.m_bun_code, nmCHU)
			iReturn = CduDecodeSHOBunrui(WKB010, KB.l_bun_code, KB.m_bun_code, KB.s_bun_code, nmSHO)
			.DSP220.Text = nmDAI
			.DSP230.Text = nmCHU
			.DSP240.Text = nmSHO
			WKB220DSP = .DSP220.Text
			WKB230DSP = .DSP230.Text
			WKB240DSP = .DSP240.Text
			
			.IMTX250.Text = KB.bun_code
			'    .DSP250.Caption = ""                '   ���g�p�R�[�h   '02/05/28 DEL
			iReturn = DecodeBUNRUI(KB.bun_code, rBUNRUI_NAME) '02/05/28 ADD
			.DSP250.Text = rBUNRUI_NAME '02/05/28 ADD
			WKB250DSP = .DSP250.Text '02/05/28 ADD
			'   ��������
			.IMTX260.Text = KB.ken_bun_code
			.DSP260.Text = DecodeFIND(KB.ken_bun_code)
			WKB260DSP = .DSP260.Text
			
			.CHK270.CheckState = IIf(KB.jutaku = "1", 1, 0)
			.CHK280.CheckState = IIf(KB.sikakari = "1", 1, 0)
			.CHK290.CheckState = IIf(KB.zan = "1", 1, 0)

			''''.OPTO300(1).Value = (KB.kanri_kubn = "1")   '   �Ǘ��敪�|����
			.OPTO300(1).Checked = (KB.kanri_kubn <> "2") '   �Ǘ��敪�|���ʃt�H���[
			.OPTO300(2).Checked = (KB.kanri_kubn = "2") '   �Ǘ��敪�|���z
			WKB300 = IIf(.OPTO300(1).Checked, 1, 2)

			.OPTO310(1).Checked = (KB.Tax_kubn = "1") '   ����ŋ敪�|�O��
			.OPTO310(2).Checked = (KB.Tax_kubn = "2") '   ����ŋ敪�|�ō�
			.OPTO310(3).Checked = (KB.Tax_kubn = "3") '   ����ŋ敪�|��ې�
			.OPTO310(1).Checked = (Not .OPTO310(2).Checked) And (Not .OPTO310(3).Checked)
			'   ����ŋ敪�|�O�Ńt�H���[
			WKB310 = IIf(.OPTO310(1).Checked, 1, IIf(.OPTO310(2).Checked, 2, 3))

			''''.OPTO320(1).Value = (KB.tana_tanka = "1")   '   �I���P���|�d���P��
			.OPTO320(1).Checked = (KB.tana_tanka <> "2") '   �I���P���|�d���P��
			.OPTO320(2).Checked = (KB.tana_tanka = "2") '   �I���P���|���͒P��
			WKB320 = IIf(.OPTO320(1).Checked, 1, 2)

			''''.OPTO330(1).Value = (KB.tana_tanka = "1")   '   �݌ɊǗ��|����
			.OPTO330(1).Checked = (KB.zaiko <> "2") '   �݌ɊǗ��|����
			.OPTO330(2).Checked = (KB.zaiko = "2") '   �݌ɊǗ��|���Ȃ�
			WKB330 = IIf(.OPTO330(1).Checked, 1, 2)

			''''.OPTO340(1).Value = (KB.tana_tanka = "0")   '   FAX���M�|����
			.OPTO340(1).Checked = (KB.Fax_yn <> "1") '   FAX���M�|����
			.OPTO340(2).Checked = (KB.Fax_yn = "1") '   FAX���M�|���Ȃ�
			WKB340 = IIf(.OPTO340(1).Checked, 1, 2)

			'    .CMB060.DataField = KB.tani
			Call COMBO_SETLIST(.CMB350(1), KB.ha_tanka1)
			Call COMBO_SETLIST(.CMB350(2), KB.ha_tanka2)
			Call COMBO_SETLIST(.CMB350(3), KB.ha_tanka3)
			Call COMBO_SETLIST(.CMB350(4), KB.ha_tanka4)
			Call COMBO_SETLIST(.CMB350(5), KB.ha_tanka5)
			
			'    .CMB350(1).Text = KB.ha_tanka1
			'    .CMB350(2).Text = KB.ha_tanka2
			'    .CMB350(3).Text = KB.ha_tanka3
			'    .CMB350(4).Text = KB.ha_tanka4
			'    .CMB350(5).Text = KB.ha_tanka5
			.IMNU360(1).Value = KB.kansan_num1
			.IMNU360(2).Value = KB.kansan_num2
			.IMNU360(3).Value = KB.kansan_num3
			.IMNU360(4).Value = KB.kansan_num4
			.IMNU360(5).Value = KB.kansan_num5
			
			'A-20250201��
			clearActCMB370Click = True
			Select Case KB.tax_rate_kbn
				Case CStr(1)
					.CMB370.SelectedIndex = 1
				Case CStr(3)
					.CMB370.SelectedIndex = 0
				Case CStr(5)
					.CMB370.SelectedIndex = 2
				Case Else
					.CMB370.SelectedIndex = 0
			End Select
			
			Select Case KB.Tax_kubn
				Case CStr(3)
					.CMB370.Enabled = False
				Case Else
					.CMB370.Enabled = True
			End Select
			
			clearActCMB370Click = False
			'A-20250201��
			
			'           ���̑�
			'   �ƎҌ���
			.IMTX410.Text = KB.g_gentei_code
			.DSP410.Text = DecodeGYOSHA(WKB010, WKB020, KB.g_gentei_code)
			WKB410DSP = .DSP410.Text
			
			'   �i�ڌ����SPREADOCX�ɂ���̂Ō�܂킵
			
			.CHK430.CheckState = IIf(KB.gen_h_ka = "1", 1, 0)
			'.IMTX440.Text = KB.tax_rate_kbn    'D-20250201
			.CHK450.CheckState = IIf(KB.tyozouhin = "1", 1, 0)
			.CHK460.CheckState = IIf(KB.jihan = "1", 1, 0)
			.CHK470.CheckState = IIf(KB.gensen = "1", 1, 0)
			.IMTX480.Text = DateSlashed(KB.nouhin_date)
			.IMTX490.Text = DateSlashed(KB.tekiyo_date)
			.CHK500.CheckState = IIf(KB.tori_kyu = "1", 1, 0)
			.IMTX510.Text = DateSlashed(KB.tori_kyu_date)
			
			Call SCR_DSPTAX() 'A-20190601
			
		End With
		
		
	End Sub
	
	'   �i�ڕ�������}�X�^����ݕ�������X�v���b�h�������\������B
	Public Sub SCR_BUSHO(ByRef bOpt As Boolean, ByRef cdHIN As String)
		'
		'   �i�ڌ��蕔��
		Dim bFirst As Boolean
		Dim nCnt As Integer
		Dim strBusho As String
		
		'    Call SpreadInit
		
		If bOpt Then
			bFirst = True
			nCnt = 0
			
			With SZ0410FRM
				.SPR420.ReDraw = False
				Do While True
					
					On Error Resume Next ' (�װ���ׯ��)
					If bFirst Then
						qSZM0011SEL.rdoParameters("Inc_code").Value = WKB010
						qSZM0011SEL.rdoParameters("jg_code").Value = WKB020
						qSZM0011SEL.rdoParameters("hin_code").Value = cdHIN
						If SZM0011myRSSW <> "qSZM0011SEL" Or ReQue = False Then
							SZM0011myRS = qSZM0011SEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
							SZM0011myRSSW = "qSZM0011SEL"
						Else
							SZM0011myRS.Requery()
						End If
						
						bFirst = False
					Else
						SZM0011myRS.MoveNext()
					End If
					
					On Error GoTo 0
					
					If B_STATUS(SZM0011myRS) Then
						'   EOF or NotFound
						Exit Do
					Else
						nCnt = nCnt + 1
						'           2000/01/23  Add     KOKOKARA
						If .SPR420.MaxRows <= nCnt Then
							.SPR420.MaxRows = nCnt + 1
							.SPR420.set_RowHeight(nCnt + 1, SPR_HEIGHT)
							Call SZ0410FRM.SpreadProperty(nCnt + 1)
							
						End If
						Call SZ0410FRM.SpreadProperty(1)
						If .SPR420.get_RowHeight(1) <> SPR_HEIGHT Then
							.SPR420.set_RowHeight(1, SPR_HEIGHT)
						End If
						'           2000/01/23  Add     KOKOMADE
						
						strBusho = SZM0011myRS.rdoColumns("bu_code").Value
						.SPR420.ROW = nCnt
						.SPR420.Col = 1
						.SPR420.Text = strBusho
						.SPR420.Col = 2
						''''            .SPR420.Text = DecodeBUSHO(strBusho)
						.SPR420.Text = CduDecodeBUSHO(strBusho)
						.SPR420.Col = 3 '2000/1/7 Add
						.SPR420.Text = "1" '2000/1/7 Add
						.SPR420.Col = 4 '2000/1/7 Add
						.SPR420.Text = strBusho '2000/1/7 Add
					End If
					
				Loop 
				
				.SPR420.ReDraw = True
			End With
		End If
		
	End Sub
	
	Public Sub PREP_SZM0011()
		
		'   Schema���̎擾  SZM0011
		MKKCMN.ZAEV_FNO = "SZM0011"
		
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			SZM0011_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    SZM0011_FILE.NAME = ""
		
		'   �ƎҌ������ރ}�X�^��QUERY�쐬
		SQL = "Select bu_code"
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0011_FILE.NAME) & "SZM0011"
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		SQL = SQL & " AND hin_code = ? "
		SQL = SQL & " ORDER by y_code "
		
		On Error Resume Next
		qSZM0011SEL = ZACN_RCN.CreateQuery("qSZM0011SEL", SQL)
		qSZM0011SEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "qSZM0011SEL"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		On Error GoTo 0
		
		qSZM0011SEL.rdoParameters(0).NAME = "Inc_code"
		qSZM0011SEL.rdoParameters(1).NAME = "jg_code"
		qSZM0011SEL.rdoParameters(2).NAME = "hin_code"
		qSZM0011SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0011SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0011SEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		qSZM0011SEL.rdoParameters(0).Size = 2
		qSZM0011SEL.rdoParameters(1).Size = 4
		qSZM0011SEL.rdoParameters(2).Size = 5
		
		'   �ƎҌ������ރ}�X�^��QUERY�쐬(INSERT)
		SQL = ""
		SQL = SQL & "INSERT INTO "
		SQL = SQL & RTrim(SZM0011_FILE.NAME) & "SZM0011("
		SQL = SQL & "Inc_code,jg_code,hin_code,y_code,bu_code) "
		SQL = SQL & "Values(?,?,?,?,? ) "
		
		On Error Resume Next
		SZM0011INS = ZACN_RCN.CreateQuery("SZM0011INS", SQL)
		SZM0011INS.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0011INS"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		
		SZM0011INS.rdoParameters(0).NAME = "Inc_code"
		SZM0011INS.rdoParameters(1).NAME = "jg_code"
		SZM0011INS.rdoParameters(2).NAME = "hin_code"
		SZM0011INS.rdoParameters(3).NAME = "y_code"
		SZM0011INS.rdoParameters(4).NAME = "bu_code"
		SZM0011INS.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011INS.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011INS.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011INS.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		SZM0011INS.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011INS.rdoParameters(0).Size = 2
		SZM0011INS.rdoParameters(1).Size = 4
		SZM0011INS.rdoParameters(2).Size = 5
		SZM0011INS.rdoParameters(4).Size = 4
		
		'   �ƎҌ������ރ}�X�^��QUERY�쐬(DELETE)
		SQL = ""
		SQL = SQL & "DELETE From "
		SQL = SQL & RTrim(SZM0011_FILE.NAME) & "SZM0011 "
		SQL = SQL & "WHERE Inc_code  = ? "
		SQL = SQL & "  AND jg_code  = ? "
		SQL = SQL & "  AND hin_code  = ? "
		SZM0011DEL = ZACN_RCN.CreateQuery("SZM0011DEL", SQL)
		SZM0011DEL.QueryTimeout = ZACN_TIME '�^�C���A�E�g���Ԃ��u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "SZM0011DEL"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		SZM0011DEL.rdoParameters(0).NAME = "Inc_code"
		SZM0011DEL.rdoParameters(1).NAME = "jg_code"
		SZM0011DEL.rdoParameters(2).NAME = "hin_code"
		SZM0011DEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011DEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011DEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		SZM0011DEL.rdoParameters(0).Size = 2
		SZM0011DEL.rdoParameters(1).Size = 4
		SZM0011DEL.rdoParameters(2).Size = 5
		
		On Error GoTo 0
		
	End Sub
	
	Public Sub AccountName(ByRef cdKAM As String, ByRef nmKAM As String)
		
		Dim nmCHU As String
		Dim nmSHO As String
		
		nmCHU = DecodeKAMOCHU(Mid(cdKAM, 1, 3))
		If nmCHU <> "" Then
			nmSHO = DecodeKAMOKU(Mid(cdKAM, 1, 3), Mid(cdKAM, 4, 6))
			''''nmKAM = DecodeKAMOKU(Mid(cdKAM, 1, 3), Mid(cdKAM, 4, 6))
			
			nmKAM = nmCHU & nmSHO
		Else
			nmKAM = nmCHU
		End If
		''''nmKAM = DecodeKAMOKU(Mid(cdKAM, 1, 3), Mid(cdKAM, 4, 6))
		
	End Sub
	
	'
	'   ���t�`�F�b�N
	Public Function CHK_DATE(ByRef vstrDate As String) As Short
		
		Dim strDate As String
		
		strDate = Trim(vstrDate)
		If Len(strDate) <> 8 Then
			CHK_DATE = F_ERR
			Exit Function
		End If
		
		ZADC_DATE.Value = strDate
		Call ZADC_SUB()
		If ZADC_STS.Value = "0" Then
			CHK_DATE = n0
		Else
			CHK_DATE = F_ERR
		End If
		
	End Function
	
	Public Function CHK_AMOUNT(ByRef lAmount As Integer) As Short
		
	End Function
	
	Public Function CHK_CURRENCY(ByRef cur As Decimal) As Short
		
	End Function
	
	Public Function CHK_BUNRUI(ByRef iOpt As Short, ByRef lBun As String, ByRef mBun As String, ByRef sBun As String) As Short
		'
		'   �啪�ށA�����ޤ�����ނ̃`�F�b�N�Ɩ��̃f�R�[�h
		Dim BunruiName As String
		Dim iReturn As Short
		
		Select Case iOpt
			Case 1
				iReturn = CduDecodeDAIBunrui(WKB010, lBun, BunruiName)
				If iReturn = F_OFF Then
					SZ0410FRM.DSP220.Text = BunruiName
					WKB220DSP = BunruiName
				End If
				
			Case 2
				iReturn = CduDecodeCHUBunrui(WKB010, lBun, mBun, BunruiName)
				If iReturn = F_OFF Then
					SZ0410FRM.DSP230.Text = BunruiName
					WKB230DSP = BunruiName
				End If
				
			Case 3
				iReturn = CduDecodeSHOBunrui(WKB010, lBun, mBun, sBun, BunruiName)
				If iReturn = F_OFF Then
					SZ0410FRM.DSP240.Text = BunruiName
					WKB240DSP = BunruiName
				End If
				'02/05/28 ADD START
			Case 4
				iReturn = DecodeBUNRUI(lBun, BunruiName)
				If iReturn = F_OFF Then
					SZ0410FRM.DSP250.Text = BunruiName
					WKB250DSP = BunruiName
				End If
				'02/05/28 ADD END
		End Select
		
		CHK_BUNRUI = IIf(iReturn = F_OFF, F_OFF, F_ERR)
		
	End Function
	
	'��ADD-2001/01/23 ==================================================================
	Private Function PSZ0410_PREP_RTN() As Boolean
		'�i�ڂ̎��є���ı������߱
		Dim SQL As String
		
		SQL = "{ CALL PSZ0410( ?,?,?,?,?,?)}"
		
		On Error Resume Next
		PSZ0410SP = ZACN_RCN.CreateQuery("", SQL)
		
		PSZ0410SP.QueryTimeout = 0
		
		If B_STATUS = 0 Then
			
			On Error GoTo 0
			
			PSZ0410SP.rdoParameters(0).NAME = "INC_CODE" : PSZ0410SP.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR : PSZ0410SP.rdoParameters(0).Size = 2
			PSZ0410SP.rdoParameters(1).NAME = "JG_CODE" : PSZ0410SP.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR : PSZ0410SP.rdoParameters(1).Size = 4
			PSZ0410SP.rdoParameters(2).NAME = "HIN_CODE" : PSZ0410SP.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR : PSZ0410SP.rdoParameters(2).Size = 5
			PSZ0410SP.rdoParameters(3).NAME = "ERRCD" : PSZ0410SP.rdoParameters(3).Direction = RDO.DirectionConstants.rdParamOutput : PSZ0410SP.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeNUMERIC
			PSZ0410SP.rdoParameters(4).NAME = "ERRMSG" : PSZ0410SP.rdoParameters(4).Direction = RDO.DirectionConstants.rdParamOutput : PSZ0410SP.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeVARCHAR
			PSZ0410SP.rdoParameters(5).NAME = "RETCODE" : PSZ0410SP.rdoParameters(5).Direction = RDO.DirectionConstants.rdParamOutput : PSZ0410SP.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeNUMERIC
			PSZ0410_PREP_RTN = True
		Else
			ZAER_NO.Value = "PSZ0410"
			On Error GoTo 0
		End If
		
	End Function
	
	Public Function PSZ0410SP_CALL_RTN(ByRef res As Short) As Boolean
		'�i�ڂ̎��є���ı�ގ��s
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		PSZ0410SP.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Inc_code, 2) '��к���
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		PSZ0410SP.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jg_code, 4) '���Ə�����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		PSZ0410SP.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5) '�i��
		
		On Error Resume Next
		Call PSZ0410SP.Execute()
		If B_STATUS = 0 And PSZ0410SP.rdoParameters(3).Value = 0 Then
			res = PSZ0410SP.rdoParameters(5).Value '���茋�� (0:���тȂ� or else:���т���)
			PSZ0410SP_CALL_RTN = True
		Else
			ZAER_CD = B_STATUS
			ZAER_KN = 1
			ZAER_NO.Value = "PSZ0410"
			ZAER_MS.Value = PSZ0410SP.rdoParameters(4).Value
			ZAER_NO.Value = "" 'A-CUST-20100610
			Call ZAER_SUB()
			ENDSW = F_END
			ERRSW = F_ERR
		End If
		On Error GoTo 0
		
	End Function
	
	'A 050909 TOP NAGANO    �ǉ��T�u���[�`��
	Public Function FILGET_SZM0170(ByRef CD1 As String, ByRef CD2 As String) As String
		
		'   �ŏ���OK�߂�l�Z�b�g
		FILGET_SZM0170 = CStr(F_OFF)
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		SZM0170_SEL.rdoParameters("Inc_code").Value = WKB010
		SZM0170_SEL.rdoParameters("jg_code").Value = WKB020
		SZM0170_SEL.rdoParameters("hi_code1").Value = CD1
		SZM0170_SEL.rdoParameters("hi_code2").Value = CD2
		On Error Resume Next ' (�װ���ׯ��)
		If SZM0170RS2SW <> "SZM0170_SEL" Or ReQue = False Then
			SZM0170RS2 = SZM0170_SEL.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			SZM0170RS2SW = "SZM0170_SEL"
		Else
			SZM0170RS2.Requery()
		End If
		
		Select Case B_STATUS(SZM0170RS2) ' (SQL���s�ð���̕]��)
			Case 0
				FILGET_SZM0170 = SZM0170RS2.rdoColumns("hi_code2").Value
			Case 24
				FILGET_SZM0170 = ""
			Case Else
				FILGET_SZM0170 = ""
				ERRSW = F_ERR
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
		
	End Function
	
	'A-CUST-20100610 Start
	'******************************************************************
	'�b�r�u�o�́E�i�ڑI��p���[�`��
	'******************************************************************
	
	Private Sub PREP_WSZ0410_RTN()
		
		'   Schema���̎擾  WSZ0410
		MKKCMN.ZAEV_FNO = "WSZ0410"
		
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			WSZ0410_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		'�i�ڎ捞�v�j
		'SQL = "Select  Y_CODE, hin_name_seisiki, KIKAKU, TANI, GYO_NAME, TANKA from "          'D-CUST-20100823
		'A-CUST-20100823 Start
		SQL = "Select  Y_CODE, hin_name_seisiki, KIKAKU, TANI, GYO_NAME, TANKA"
		SQL = SQL & ", TEKI_DATE, HA_TANI, KANSANSU, JAN_CODE, JAN_S_CODE, BAR_CODE"
		SQL = SQL & " from "
		'A-CUST-20100823 End
		SQL = SQL & RTrim(WSZ0410_FILE.NAME) & "WSZ0410 WHERE INC_CODE = ? AND JG_CODE = ? "
		SQL = SQL & "ORDER BY INC_CODE, JG_CODE, Y_CODE"
		On Error Resume Next
		WSZ0410SEL01 = ZACN_RCN.CreateQuery("WSZ0410SEL01", SQL)
		WSZ0410SEL01.QueryTimeout = 0 '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "WSZ0410"
			GoTo PREPSZM_ERR
		End If
		WSZ0410SEL01.rdoParameters(0).NAME = "Inc_code" : WSZ0410SEL01.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR : WSZ0410SEL01.rdoParameters(0).Size = 2
		WSZ0410SEL01.rdoParameters(1).NAME = "jg_code" : WSZ0410SEL01.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR : WSZ0410SEL01.rdoParameters(1).Size = 4
		On Error GoTo 0
		
		'�i�ڎ捞�v�j
		SQL = "Select INC_CODE "
		SQL = SQL & "from " & RTrim(WSZ0410_FILE.NAME) & "WSZ0410 "
		SQL = SQL & "WHERE INC_CODE = ? " '��ЃR�[�h
		SQL = SQL & "AND JG_CODE = ? " '���Ə��R�[�h
		SQL = SQL & "AND Y_CODE = ? " '�����R�[�h
		SQL = SQL & "for update nowait"
		On Error Resume Next
		WSZ0410SEL02 = ZACNA_RCN.CreateQuery("WSZ0410SEL02", SQL)
		WSZ0410SEL02.QueryTimeout = 0 '�^�C���A�E�g���u�����v�ɐݒ�
		WSZ0410SEL02.LockType = RDO.LockTypeConstants.rdConcurRowVer
		WSZ0410SEL02.CursorType = RDO.ResultsetTypeConstants.rdOpenKeyset
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "WSZ0410"
			GoTo PREPSZM_ERR
		End If
		WSZ0410SEL02.rdoParameters(0).NAME = "Inc_code" : WSZ0410SEL02.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR : WSZ0410SEL02.rdoParameters(0).Size = 2
		WSZ0410SEL02.rdoParameters(1).NAME = "jg_code" : WSZ0410SEL02.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR : WSZ0410SEL02.rdoParameters(1).Size = 4
		WSZ0410SEL02.rdoParameters(2).NAME = "y_code" : WSZ0410SEL02.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeNUMERIC ': WSZ0410SEL02(2).Size = 6
		On Error GoTo 0
		
		'�i�ڎ捞�v�j��INSERT�pQUERY
		SQL = ""
		SQL = SQL & "INSERT INTO "
		SQL = SQL & RTrim(WSZ0410_FILE.NAME) & "WSZ0410 ("
		SQL = SQL & "INC_CODE,JG_CODE,Y_CODE,hin_name_seisiki,KIKAKU,TANI,"
		'SQL = SQL & "GYO_NAME,TANKA,ENTRY_OP_CODE,ENTRY_OP_DATE,ENTRY_OP_TIME) "   'D-CUST-20100823
		'SQL = SQL & "Values(?,?,?,?,?,?,?,?,?,?,?) "                               'D-CUST-20100823
		'A-CUST-20100823 Start
		SQL = SQL & "GYO_NAME,TANKA,ENTRY_OP_CODE,ENTRY_OP_DATE,ENTRY_OP_TIME,"
		SQL = SQL & "TEKI_DATE,HA_TANI,KANSANSU,JAN_CODE,JAN_S_CODE,BAR_CODE"
		SQL = SQL & ")Values("
		SQL = SQL & "?,?,?,?,?,?,?,?,?,?,?,"
		SQL = SQL & "?,?,?,?,?,?"
		SQL = SQL & ")"
		'A-CUST-20100823 End
		On Error Resume Next
		WSZ0410INS = ZACNA_RCN.CreateQuery("WSZ0410INS", SQL)
		WSZ0410INS.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "WSZ0410"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		WSZ0410INS.rdoParameters(0).NAME = "INC_CODE"
		WSZ0410INS.rdoParameters(1).NAME = "JG_CODE"
		WSZ0410INS.rdoParameters(2).NAME = "Y_CODE"
		WSZ0410INS.rdoParameters(3).NAME = "hin_name_seisiki"
		WSZ0410INS.rdoParameters(4).NAME = "KIKAKU"
		WSZ0410INS.rdoParameters(5).NAME = "TANI"
		WSZ0410INS.rdoParameters(6).NAME = "GYO_NAME"
		WSZ0410INS.rdoParameters(7).NAME = "TANKA"
		WSZ0410INS.rdoParameters(8).NAME = "ENTRY_OP_CODE"
		WSZ0410INS.rdoParameters(9).NAME = "ENTRY_OP_DATE"
		WSZ0410INS.rdoParameters(10).NAME = "ENTRY_OP_TIME"
		'A-CUST-20100823 Start
		WSZ0410INS.rdoParameters(11).NAME = "TEKI_DATE"
		WSZ0410INS.rdoParameters(12).NAME = "HA_TANI"
		WSZ0410INS.rdoParameters(13).NAME = "KANSANSU"
		WSZ0410INS.rdoParameters(14).NAME = "JAN_CODE"
		WSZ0410INS.rdoParameters(15).NAME = "JAN_S_CODE"
		WSZ0410INS.rdoParameters(16).NAME = "BAR_CODE"
		'A-CUST-20100823 End
		WSZ0410INS.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeDECIMAL
		WSZ0410INS.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(4).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(5).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(6).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(7).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		WSZ0410INS.rdoParameters(8).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(9).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(10).Type = RDO.DataTypeConstants.rdTypeCHAR
		'A-CUST-20100823 Start
		WSZ0410INS.rdoParameters(11).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(12).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(13).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		WSZ0410INS.rdoParameters(14).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(15).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410INS.rdoParameters(16).Type = RDO.DataTypeConstants.rdTypeCHAR
		'A-CUST-20100823 End
		WSZ0410INS.rdoParameters(0).Size = 2
		WSZ0410INS.rdoParameters(1).Size = 4
		'WSZ0410INS(2).Size = 6
		WSZ0410INS.rdoParameters(3).Size = 40
		WSZ0410INS.rdoParameters(4).Size = 20
		WSZ0410INS.rdoParameters(5).Size = 4
		WSZ0410INS.rdoParameters(6).Size = 30
		'WSZ0410INS(7).Size = 8.2
		WSZ0410INS.rdoParameters(8).Size = 6
		WSZ0410INS.rdoParameters(9).Size = 8
		WSZ0410INS.rdoParameters(10).Size = 6
		'A-CUST-20100823 Start
		WSZ0410INS.rdoParameters(11).Size = 8
		WSZ0410INS.rdoParameters(12).Size = 4
		'WSZ0410INS(13).Size =
		WSZ0410INS.rdoParameters(14).Size = 13
		WSZ0410INS.rdoParameters(15).Size = 7
		WSZ0410INS.rdoParameters(16).Size = 30
		'A-CUST-20100823 End
		
		'�i�ڎ捞�v�j��DELETE�pQUERY
		SQL = ""
		SQL = SQL & "DELETE "
		SQL = SQL & "FROM "
		SQL = SQL & RTrim(WSZ0410_FILE.NAME) & "WSZ0410 "
		SQL = SQL & "WHERE Inc_code=? and jg_code=?  and y_code=? "
		On Error Resume Next
		WSZ0410DEL = ZACNA_RCN.CreateQuery("WSZ0410DEL", SQL)
		WSZ0410DEL.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = "WSZ0410"
			Call ZAER_SUB()
			On Error GoTo 0
			Exit Sub
		End If
		WSZ0410DEL.rdoParameters(0).NAME = "Inc_code"
		WSZ0410DEL.rdoParameters(1).NAME = "jg_code"
		WSZ0410DEL.rdoParameters(2).NAME = "y_code"
		WSZ0410DEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410DEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		WSZ0410DEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeNUMERIC
		WSZ0410DEL.rdoParameters(0).Size = 2
		WSZ0410DEL.rdoParameters(1).Size = 4
		'WSZ0410DEL(2).Size = 6
		
		
		Exit Sub
		
PREPSZM_ERR: 
		ZAER_FID = "RAZ99"
		ZAER_KN = 1
		Call ZAER_SUB()
		ERRSW = F_ERR
		On Error GoTo 0
		
	End Sub
	
	Public Sub BEGIN_RTN()
		
		' DB �ɖ₢���킹��...
		On Error Resume Next
		ZACNA_RCN.BeginTrans()
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		End If
		On Error GoTo 0
		
	End Sub
	
	Public Sub COMMIT_RTN()
		
		' DB �ɖ₢���킹��...
		On Error Resume Next
		ZACNA_RCN.CommitTrans()
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		End If
		On Error GoTo 0
		
	End Sub
	
	Public Sub ROLLBACK_RTN()
		
		' DB �ɖ₢���킹��...
		On Error Resume Next
		ZACNA_RCN.RollbackTrans()
		If B_STATUS <> 0 Then
			ENDSW = F_END
			ZAER_KN = 1
			ZAER_NO.Value = ""
			Call ZAER_SUB()
		End If
		On Error GoTo 0
		
	End Sub
	
	Public Function TORIKOMI_DEL() As Boolean
		Dim strDate As String
		
		'--- �����Ώۃf�[�^�����b�N
		strDate = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Month, -1, SYSDATE), "YYYYMMDD")
		If GET_WSZ0410_LOCK_RTN(strDate) = False Then
			Exit Function '�G���[�����I
		End If
		If WSZ0410INVSW <> F_INV Then
			'--- �����Ώۃf�[�^���폜
			If DEL_WSZ0410_RTN(strDate) = False Then
				Exit Function '�G���[�����I
			End If
		End If
		
	End Function
	
	Public Function TORIKOMI_RTN() As Boolean
		'**************************************************
		'*  �i�ڎ捞����  �T�u���[�`��                    *
		'*      CSV����Ώۃf�[�^���捞                   *
		'**************************************************
		CSV_CNT = 0
		CSVERR_CNT = 0 '�X�V�Ɏ��s��������
		FOPENSW = False
		ERRSW = F_OFF
		
		WSZ0410.Inc_code = WKB010 '��к���
		WSZ0410.jg_code = WKB020 '���Ə�����
		WSZ0410.Entry_Op_date = VB6.Format(SYSDATE, "YYYYMMDD")
		WSZ0410.Entry_Op_time = VB6.Format(SYSDATE, "HHNNDD")
		
		If GET_WSZ0410_RENBAN_RTN(WSZ0410.y_code) = False Then
			Exit Function '�G���[�����I
		End If
		
		'IN_ITEM_MAX = 5                        'D-CUST-20100823
		
		'--- �i�ڎ捞�����v�j  �X�V����
		Do Until ENDSW = F_END
TORIKOMI_NXT01: 
			Call READMEISAI_RTN() 'CSV�f�[�^READ�I
			If ENDSW = F_END Then GoTo TORIKOMI_END01
			If ERRSW = F_ERR Then GoTo TORIKOMI_END01 '�G���[�����I
			
			'�X�V�p�f�[�^�Z�b�g
			If WSZ0410.y_code < 999999 Then
				WSZ0410.y_code = WSZ0410.y_code + 1
			Else
				WSZ0410.y_code = 1
			End If
			WSZ0410.hin_name_seisiki = WCSV_DATA.hin_name
			WSZ0410.kikaku = WCSV_DATA.kikaku
			WSZ0410.tani = WCSV_DATA.tani
			WSZ0410.gyo_name = WCSV_DATA.gyo_name
			WSZ0410.tanka = WCSV_DATA.tanka
			'A-CUST-20100823 Start
			WSZ0410.teki_date = WCSV_DATA.teki_date
			WSZ0410.ha_tani = WCSV_DATA.ha_tani
			WSZ0410.kansansu = WCSV_DATA.kansansu
			WSZ0410.jan_code = WCSV_DATA.jan_code
			WSZ0410.jan_s_code = WCSV_DATA.jan_s_code
			WSZ0410.bar_code = WCSV_DATA.bar_code
			'A-CUST-20100823 End
			
			If INS_WSZ0410_RTN = False Then
				GoTo TORIKOMI_END01 '�G���[�����I
			End If
		Loop 
		
TORIKOMI_END01: 
		If FOPENSW = True Then
			Call TextFile_Close(INPFNum)
		End If
	End Function
	
	' **************************************************************
	'   ��荞�������ł���f�[�^��1���ǂݍ���
	' **************************************************************
	Private Sub READMEISAI_RTN()
		
		Call READMEISAI_RTN_CORE() '1��ذ��
		
		If ENDSW = F_OFF Then
			If CHK_CSVDATA() Then '��荞�߂�f�[�^���H
				Call CONVMEISAI_RTN() '�捞OK
			Else
				ERRSW = F_ERR
			End If
		End If
	End Sub
	' **************************************************************
	'   ���ׂ̓ǂݍ���
	' **************************************************************
	Private Sub READMEISAI_RTN_CORE()
		INPFENDSW = F_OFF
		If FSTSW = F_FST Then
			FSTSW = F_OFF
			INPFNum = TextFile_Read_Open(Trim(WKBCSVFILE))
			If ERRSW = F_ERR Then
				Exit Sub
			End If
			FOPENSW = True
		End If
		Call TextF_Read(INPFNum, IN_ITEM_MAX)
		
		If INPFENDSW = F_END Then
			ENDSW = F_END
			Call TextFile_Close(INPFNum)
			FOPENSW = False
		End If
	End Sub
	
	' **************************************************************
	'   ÷��̧�� Read Open
	' **************************************************************
	Private Function TextFile_Read_Open(ByRef TextF_Name As String) As Short 'A-04/08/23
		On Error Resume Next
		TextFile_Read_Open = FreeFile
		FileOpen(TextFile_Read_Open, TextF_Name, OpenMode.Input, OpenAccess.Read)
		If Err.Number <> n0 Then
			ZAER_CD = 904 '����ݴװ
			ZAER_NO.Value = ""
			Call ZAER_SUB()
			
			ERRSW = F_ERR
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
	End Function
	
	' **************************************************************
	'   ÷��̧�� Read
	' **************************************************************
	Private Sub TextF_Read(ByRef TextFNum As Short, ByRef ITEM_MAX As Short)
		Dim RD_LEN As Integer
		Dim SET_LEN As Integer ' �擾���ڂ� Length
		Dim SET_CNT As Short '
		Dim HD_CNT As Integer
		Dim IWRD As String ' TXT �S�̎擾
		Dim IWLEN As Integer ' TXT �S�̂� Length
		
		RD_LEN = n1
		SET_LEN = n1
		HD_CNT = n1
		'UPGRADE_NOTE: Erase �� System.Array.Clear �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		System.Array.Clear(IN_ITEM, 0, IN_ITEM.Length)
		IN_ITEM_CNT = n0
		
		On Error Resume Next
		IWRD = LineInput(TextFNum)
		Select Case Err.Number
			Case n0
				On Error GoTo 0
				IWLEN = Len(IWRD)
				For SET_CNT = n1 To IN_ITEM_MAX + 1
					Do Until Mid(IWRD, RD_LEN, n1) = vbTab Or RD_LEN > IWLEN
						RD_LEN = RD_LEN + n1
						SET_LEN = SET_LEN + n1
					Loop 
					' txt �� �ϊ�ܰ�
					IN_ITEM(SET_CNT) = Mid(IWRD, HD_CNT, SET_LEN - n1)
					IN_ITEM_CNT = SET_CNT
					' �O�������ٺ�ð��݂��폜
					If Left(IN_ITEM(SET_CNT), n1) = Chr(34) And Right(IN_ITEM(SET_CNT), n1) = Chr(34) Then
						IN_ITEM(SET_CNT) = Mid(IN_ITEM(SET_CNT), n2, Len(IN_ITEM(SET_CNT)) - n2)
					End If
					IN_ITEM(SET_CNT) = RTrim(IN_ITEM(SET_CNT)) 'A-CUST-20100823
					If RD_LEN > IWLEN Then
						Exit For
					End If
					HD_CNT = RD_LEN + n1
					RD_LEN = RD_LEN + n1
					SET_LEN = n1
				Next SET_CNT
				CSV_CNT = CSV_CNT + n1
				
			Case 62 '�d�n�e
				On Error GoTo 0
				INPFENDSW = F_END
				Exit Sub
				
			Case Else
				ZAER_CD = 906 '�Ǎ��ݴװ
				ZAER_NO.Value = ""
				Call ZAER_SUB()
				
				On Error GoTo 0
				ERRSW = F_ERR
				INPFENDSW = F_END
				Exit Sub
		End Select
	End Sub
	
	' **************************************************************
	'   ÷��̧�� Close
	' **************************************************************
	Private Sub TextFile_Close(ByRef TextFNum As Short) 'A-04/08/23
		On Error Resume Next
		FileClose(TextFNum)
		On Error GoTo 0
	End Sub
	
	' **************************************************************
	'   �ǂݍ���CSV�f�[�^�� �捞�f�[�^���ǂ����𔻒肷��
	' **************************************************************
	Private Function CHK_CSVDATA() As Boolean
		Dim FLG As Short
		Dim strMsg As String
		Dim i As Integer
		Dim sTanka As String
		
		sTanka = Trim(IN_ITEM(5))
		IN_ITEM(5) = ""
		For i = 1 To Len(sTanka)
			If Mid(sTanka, i, 1) <> "," Then
				IN_ITEM(5) = IN_ITEM(5) & Mid(sTanka, i, 1)
			End If
		Next i
		'A-CUST-20100823 Start
		sTanka = IN_ITEM(CsvPos.kansansu)
		IN_ITEM(CsvPos.kansansu) = ""
		For i = 1 To Len(sTanka)
			If Mid(sTanka, i, 1) <> "," Then
				IN_ITEM(CsvPos.kansansu) = IN_ITEM(CsvPos.kansansu) & Mid(sTanka, i, 1)
			End If
		Next i
		'A-CUST-20100823 End

		If IN_ITEM_CNT <> IN_ITEM_MAX Then
			strMsg = "���C�A�E�g���Ⴂ�܂��B"
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(1), vbFromUnicode)) > 40 Then
			strMsg = "�����i���̕����������߂��܂��B"
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(2), vbFromUnicode)) > 20 Then
			strMsg = "�K�i�̕����������߂��܂��B"
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(3), vbFromUnicode)) > 4 Then
			strMsg = "�P�ʂ̕����������߂��܂��B"
			'A-CUST-20100823 Start
		ElseIf Not CHK_tani(IN_ITEM(3)) Then
			strMsg = "�P�ʂ��o�^����Ă��܂���B"
			'A-CUST-20100823 End
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(4), vbFromUnicode)) > 30 Then
			strMsg = "���Ǝ҂̕����������߂��܂��B"
		ElseIf Not IsNumeric(IN_ITEM(5)) Then
			strMsg = "�P�������l�ł͂���܂���B"
		ElseIf Val(IN_ITEM(5)) > 99999999.99 Or Val(IN_ITEM(5)) < 0 Then
			strMsg = "�P���̒l���Ԉ���Ă��܂��B"
			'ElseIf Val(Format$(Val(IN_ITEM(5)), "0.00")) <> Val(IN_ITEM(5)) < 0 Then      '�������`�F�b�N  D-CUST-20100823
		ElseIf Val(VB6.Format(Val(IN_ITEM(5)), "0.00")) <> Val(IN_ITEM(5)) Then  '�������`�F�b�N       A-CUST-20100823
			strMsg = "�P���̒l���Ԉ���Ă��܂��B"
			'A-CUST-20100823 Start
		ElseIf Not CHK_Tekiyobi(IN_ITEM(CsvPos.teki_date), strMsg) Then
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(CsvPos.ha_tani), vbFromUnicode)) > 4 Then
			strMsg = "�����P�ʂ̕����������߂��܂��B"
		ElseIf Not CHK_tani(IN_ITEM(CsvPos.ha_tani)) Then
			strMsg = "�����P�ʂ��o�^����Ă��܂���B"
		ElseIf Not CHK_Kansansu(IN_ITEM(CsvPos.ha_tani), IN_ITEM(CsvPos.kansansu), strMsg) Then
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(CsvPos.jan_code), vbFromUnicode)) > 13 Then
			strMsg = "JAN�W���R�[�h�̕����������߂��܂��B"
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(CsvPos.jan_s_code), vbFromUnicode)) > 7 Then
			strMsg = "JAN�Z�k�R�[�h�̕����������߂��܂��B"
			'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
			'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		ElseIf LenB(Microsoft.VisualBasic.StrConv(IN_ITEM(CsvPos.bar_code), vbFromUnicode)) > 30 Then
			strMsg = "���̑��o�[�R�[�h�̕����������߂��܂��B"
			'A-CUST-20100823 End
		Else
			CHK_CSVDATA = True
			Exit Function
		End If
		
		CHK_CSVDATA = False
		
		Call MsgBox(strMsg & VB6.Format(CSV_CNT, "@@") & "�s��", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, SZ0411FRM.Text)
		
	End Function
	
	'A-CUST-20100823 Start
	Private Function CHK_tani(ByVal sTani As String) As Boolean
		Dim i As Integer
		
		If sTani = "" Then
			CHK_tani = True
			Exit Function
		End If
		
		CHK_tani = False
		For i = 1 To TaniCnt
			If RTrim(Tani_T(i)) = sTani Then
				CHK_tani = True
				Exit For
			End If
		Next i
		
	End Function
	
	Private Function CHK_Tekiyobi(ByVal sTekiyobi As String, ByRef strMsg As String) As Boolean
		If sTekiyobi = "" Then
			CHK_Tekiyobi = True
			Exit Function
		End If
		
		If Not IsNumeric(sTekiyobi) Then
			strMsg = "�K�p�������l�ł͂���܂���B"
			CHK_Tekiyobi = False
			Exit Function
		End If
		'UPGRADE_ISSUE: �萔 vbFromUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' ���N���b�N���Ă��������B
		'UPGRADE_ISSUE: LenB �֐��̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' ���N���b�N���Ă��������B
		If LenB(Microsoft.VisualBasic.StrConv(sTekiyobi, vbFromUnicode)) > 8 Then
			strMsg = "�K�p���̕����������߂��܂��B"
			CHK_Tekiyobi = False
			Exit Function
		End If

		ZADC_DATE.Value = sTekiyobi
		Call ZADC_SUB()
		If ZADC_STS.Value = "0" Then
			CHK_Tekiyobi = True
		Else
			CHK_Tekiyobi = False
			strMsg = "�K�p���̒l���Ԉ���Ă��܂��B"
		End If
		
	End Function
	
	Private Function CHK_Kansansu(ByVal ha_tani As String, ByVal kansansu As String, ByRef strMsg As String) As Boolean
		CHK_Kansansu = False
		
		If kansansu = "" Then
			kansansu = "0"
		End If
		If Not IsNumeric(kansansu) Then
			strMsg = "���Z�������l�ł͂���܂���B"
			Exit Function
		End If
		
		If ha_tani = "" Then
			If Val(kansansu) = 0 Then
				CHK_Kansansu = True
			Else
				strMsg = "�����P�ʂȂ��Ŋ��Z��������܂��B"
			End If
			Exit Function
		End If
		
		If Val(kansansu) > 99999.99 Or Val(kansansu) < 0 Then
			strMsg = "���Z���̒l���Ԉ���Ă��܂��B"
		ElseIf Val(VB6.Format(Val(kansansu), "0.00")) <> Val(kansansu) Then  '�������`�F�b�N
			strMsg = "���Z���̒l���Ԉ���Ă��܂��B"
		Else
			CHK_Kansansu = True
		End If
		
	End Function
	'A-CUST-20100823 End
	
	' **************************************************************
	'   �ǂݍ���CSV�f�[�^�� �\����WCSV_DATA�ɐݒ肷��
	' **************************************************************
	Private Sub CONVMEISAI_RTN()
		Dim i As Short
		
		For i = 1 To IN_ITEM_MAX
			Select Case i
				Case 1 '��������
					WCSV_DATA.hin_name = RTrim(IN_ITEM(i))
				Case 2 '�K�i
					WCSV_DATA.kikaku = RTrim(IN_ITEM(i))
				Case 3 '�P��
					WCSV_DATA.tani = RTrim(IN_ITEM(i))
				Case 4 '�ƎҖ�
					WCSV_DATA.gyo_name = RTrim(IN_ITEM(i))
				Case 5 '�P��
					WCSV_DATA.tanka = CDec(Val(Trim(IN_ITEM(i))))
					'A-CUST-20100823 Start
				Case CsvPos.teki_date '�K�p��
					WCSV_DATA.teki_date = IN_ITEM(i)
				Case CsvPos.ha_tani '�����P��
					WCSV_DATA.ha_tani = IN_ITEM(i)
				Case CsvPos.kansansu '���Z��
					WCSV_DATA.kansansu = CDec(Val(IN_ITEM(i)))
				Case CsvPos.jan_code 'JAN�W���R�[�h
					WCSV_DATA.jan_code = IN_ITEM(i)
				Case CsvPos.jan_s_code 'JAN�Z�k�R�[�h
					WCSV_DATA.jan_s_code = IN_ITEM(i)
				Case CsvPos.bar_code '���̑��o�[�R�[�h
					WCSV_DATA.bar_code = IN_ITEM(i)
					'A-CUST-20100823 End
			End Select
		Next 
		
	End Sub
	
	Private Function GET_WSZ0410_LOCK_RTN(ByVal strDate As String) As Boolean
		'******************************************************************************************
		'   �i�ڎ捞�v�j
		'       �Ώۃf�[�^�̃��b�N
		'       �߂�l  True : ����     False : ���s
		'******************************************************************************************
		Dim SQL As String
		
		WSZ0410INVSW = F_OFF
		
		SQL = "SELECT * "
		SQL = SQL & " FROM " & RTrim(WSZ0410_FILE.NAME) & "WSZ0410"
		SQL = SQL & " WHERE inc_code = '" & WKB010 & "'"
		SQL = SQL & "   AND jg_code = '" & WKB020 & "'"
		SQL = SQL & "   AND ENTRY_OP_DATE < '" & strDate & "'"
		SQL = SQL & " FOR UPDATE NOWAIT"
		
		On Error Resume Next
		WSZ0410RS = ZACNA_RCN.OpenResultset(SQL)
		Select Case B_STATUS(WSZ0410RS)
			Case n0
				GET_WSZ0410_LOCK_RTN = True
				
			Case 24
				WSZ0410INVSW = F_INV
				GET_WSZ0410_LOCK_RTN = True
				
			Case -54
				GET_WSZ0410_LOCK_RTN = True
				ZAER_CD = 201
				ZAER_KN = 0
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020
				ERRSW = F_ERR
				Call ZAER_SUB()
				WSZ0410RS.Close()
				
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
				WSZ0410RS.Close()
		End Select
		
		On Error GoTo 0
	End Function
	
	Private Function GET_WSZ0410_RENBAN_RTN(ByRef lngRenban As Integer) As Boolean
		'******************************************************************************************
		'   �i�ڎ捞�v�j
		'       �A�Ԏ擾
		'       �߂�l  True : ����     False : ���s
		'******************************************************************************************
		Dim SQL As String
		
		SQL = "SELECT NVL(MAX(Y_CODE),0) Y_CODE "
		SQL = SQL & " FROM " & RTrim(WSZ0410_FILE.NAME) & "WSZ0410"
		SQL = SQL & " WHERE inc_code = '" & WKB010 & "'"
		SQL = SQL & "   AND jg_code = '" & WKB020 & "'"
		
		On Error Resume Next
		WSZ0410RS = ZACN_RCN.OpenResultset(SQL)
		Select Case B_STATUS(WSZ0410RS)
			Case n0
				GET_WSZ0410_RENBAN_RTN = True
				lngRenban = WSZ0410RS.rdoColumns("y_code").Value
			Case 24
				lngRenban = 0
				GET_WSZ0410_RENBAN_RTN = True
				
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
				Exit Function
		End Select
		WSZ0410RS.Close()
		On Error GoTo 0
		
	End Function
	
	Private Function DEL_WSZ0410_RTN(ByVal strDate As String) As Boolean
		'******************************************************************************************
		'   �i�ڎ捞�v�j
		'       �폜�p
		'       �߂�l  True : ����     False : ���s
		'******************************************************************************************
		Dim SQL As String
		
		SQL = "DELETE FROM " & RTrim(WSZ0410_FILE.NAME) & "WSZ0410 "
		SQL = SQL & " WHERE inc_code = '" & WKB010 & "'"
		SQL = SQL & "   AND jg_code = '" & WKB020 & "'"
		SQL = SQL & "   AND ENTRY_OP_DATE < '" & strDate & "'"
		
		On Error Resume Next
		Call ZACNA_RCN.Execute(SQL)
		Select Case B_STATUS
			Case n0
				DEL_WSZ0410_RTN = True
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
		End Select
		
		On Error GoTo 0
	End Function
	
	Private Function INS_WSZ0410_RTN() As Boolean
		'******************************************************************************************
		'   �i�ڎ捞�v�j
		'       CSV�̃f�[�^����ɃC���T�[�g����
		'       �߂�l  True : ����     False : ���s
		'******************************************************************************************
		WSZ0410INS.rdoParameters("Inc_code").Value = WSZ0410.Inc_code
		WSZ0410INS.rdoParameters("jg_code").Value = WSZ0410.jg_code
		WSZ0410INS.rdoParameters("y_code").Value = WSZ0410.y_code
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("hin_name_seisiki").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.hin_name_seisiki)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("kikaku").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.kikaku)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("tani").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.tani)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("gyo_name").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.gyo_name)
		WSZ0410INS.rdoParameters("tanka").Value = WSZ0410.tanka
		WSZ0410INS.rdoParameters("Entry_Op_code").Value = WG_OPCODE
		WSZ0410INS.rdoParameters("Entry_Op_date").Value = WSZ0410.Entry_Op_date
		WSZ0410INS.rdoParameters("Entry_Op_time").Value = WSZ0410.Entry_Op_time
		'A-CUST-20100823 Start
		WSZ0410INS.rdoParameters("teki_date").Value = WSZ0410.teki_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("ha_tani").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.ha_tani)
		WSZ0410INS.rdoParameters("kansansu").Value = WSZ0410.kansansu
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("jan_code").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.jan_code)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("jan_s_code").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.jan_s_code)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZAFIXSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		WSZ0410INS.rdoParameters("bar_code").Value = MKKCMN.ZAFIXSTR_SUB(WSZ0410.bar_code)
		'A-CUST-20100823 End
		
		On Error Resume Next
		Call WSZ0410INS.Execute()
		Select Case B_STATUS
			Case n0
				INS_WSZ0410_RTN = True
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WSZ0410.Inc_code & "-" & WSZ0410.jg_code
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
		End Select
		
		On Error GoTo 0
	End Function
	'A-CUST-20100610 End
	
	Public Sub GO_WKDELETE()
		'**************************************************
		'*  �i�ڎ捞����  �T�u���[�`��                    *
		'**************************************************
		ERRSW = F_OFF
		ENDSW = F_OFF
		
		'--- �����Ώۃf�[�^�����b�N
		WSZ0410SEL02.rdoParameters("Inc_code").Value = WKB010
		WSZ0410SEL02.rdoParameters("jg_code").Value = WKB020
		WSZ0410SEL02.rdoParameters("y_code").Value = RENBAN_SEN
		
		On Error Resume Next
		WSZ0410RS = WSZ0410SEL02.OpenResultset(SQL)
		Select Case B_STATUS(WSZ0410RS)
			Case n0
			Case 24, -54
				Exit Sub
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(RENBAN_SEN, "000000")
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
				Exit Sub
		End Select
		
		WSZ0410DEL.rdoParameters("Inc_code").Value = WKB010
		WSZ0410DEL.rdoParameters("jg_code").Value = WKB020
		WSZ0410DEL.rdoParameters("y_code").Value = RENBAN_SEN
		
		On Error Resume Next
		Call WSZ0410DEL.Execute(SQL)
		Select Case B_STATUS
			Case n0
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "WSZ0410"
				ZAER_MS.Value = WKB010 & "-" & WKB020 & "-" & VB6.Format(RENBAN_SEN, "000000")
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
		End Select
		
		On Error GoTo 0
	End Sub
	
	'A-CUST20130212��
	Public Sub PREP_JAN_RTN()
		
		'   Schema���̎擾
		MKKCMN.ZAEV_FNO = "JAN"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			JAN_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    JAN_FILE.NAME = ""
		'D-20130424-S
		'    SQL = "Select k21,k42,k44,k57,k58,SUBSTRB(K14,1,40) K14 "
		'D-20130424-E
		'A-20130424-S
		SQL = "Select "
		SQL = SQL & " NVL(k21,' ') K21"
		SQL = SQL & ",NVL(k42,  0) K42"
		SQL = SQL & ",NVL(k44,' ') K44"
		SQL = SQL & ",NVL(k57,' ') K57"
		SQL = SQL & ",NVL(k58,  0) K58"
		SQL = SQL & ",NVL(K17,' ') K17"
		'A-20130424-E
		SQL = SQL & " from "
		SQL = SQL & RTrim(JAN_FILE.NAME) & "JAN"
		SQL = SQL & " WHERE k4 = ? "
		
		On Error Resume Next
		qJANSEL = ZACN_RCN.CreateQuery("qJANSEL", SQL)
		qJANSEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "JAN"
			
		End If
		On Error GoTo 0
		
		qJANSEL.rdoParameters(0).NAME = "k4"
		qJANSEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qJANSEL.rdoParameters(0).Size = 13
		
		'    bJANReady = True
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Public Function FILGET_JAN() As Boolean
		
		'   �ŏ���OK�߂�l�Z�b�g
		FILGET_JAN = F_OFF
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		qJANSEL.rdoParameters("k4").Value = JAN_BUF0.k4
		On Error Resume Next
		JANRS = qJANSEL.OpenResultset()
		
		Select Case B_STATUS(JANRS) ' (SQL���s�ð���̕]��)
			Case 0
				FILGET_JAN = True
				JAN.k21 = JANRS.rdoColumns("k21").Value
				JAN.k44 = JANRS.rdoColumns("k44").Value
				JAN.k42 = JANRS.rdoColumns("k42").Value
				JAN.k57 = JANRS.rdoColumns("k57").Value
				JAN.k58 = JANRS.rdoColumns("k58").Value
				'JAN.k14 = JANRS!k14    'D-20130424-
				JAN.k17 = JANRS.rdoColumns("k17").Value 'A-20130424-
			Case 24
				FILGET_JAN = False
				JAN.k21 = ""
				JAN.k44 = ""
				JAN.k42 = 0
				JAN.k57 = ""
				JAN.k58 = 0
				'JAN.k14 = ""       'D-20130424-
				JAN.k17 = "" 'A-20130424-
			Case Else
				FILGET_JAN = False
				ERRSW = F_ERR
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	'A-CUST20130212��
	'A -CUST20130212��
	Public Sub PREP_JAN_BUNRUI_RTN()
		
		'   Schema���̎擾
		MKKCMN.ZAEV_FNO = "JAN_BUNRUI"
		Call MKKCMN.ZAEV_SUB()
		If CDbl(MKKCMN.ZAEV_ERR) <> 0 Then
			ERRSW = F_ERR
			Exit Sub
		Else
			JAN_BUNRUI_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		'    JAN_BUNRUI_FILE.NAME = ""
		
		SQL = "Select Bk4  "
		SQL = SQL & " from "
		SQL = SQL & RTrim(JAN_BUNRUI_FILE.NAME) & "JAN_BUNRUI"
		SQL = SQL & " WHERE Bk1 = ? "
		SQL = SQL & "   AND Bk2 = '4' " '�ڍו��ނ̂�
		
		On Error Resume Next
		qJAN_BUNRUISEL = ZACN_RCN.CreateQuery("qJAN_BUNRUISEL", SQL)
		qJAN_BUNRUISEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "JAN_BUNRUI"
			
		End If
		On Error GoTo 0
		
		qJAN_BUNRUISEL.rdoParameters(0).NAME = "Bk1"
		qJAN_BUNRUISEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		qJAN_BUNRUISEL.rdoParameters(0).Size = 6
		
		'    bJAN_BUNRUIReady = True
	End Sub
	'A-CUST20130212��
	'A-CUST20130212��
	Public Function FILGET_JAN_BUNRUI() As Boolean
		
		'   �ŏ���OK�߂�l�Z�b�g
		FILGET_JAN_BUNRUI = False
		
		If RTrim(JAN_BUNRUI_BUF0.BK1) = "" Then Exit Function
		
		'   WHERE�̌��������ɋƎ�NO��ݒ�
		qJAN_BUNRUISEL.rdoParameters("BK1").Value = JAN_BUNRUI_BUF0.BK1
		On Error Resume Next
		JAN_BUNRUIRS = qJAN_BUNRUISEL.OpenResultset()
		
		Select Case B_STATUS(JAN_BUNRUIRS) ' (SQL���s�ð���̕]��)
			Case 0
				FILGET_JAN_BUNRUI = True
				JAN_BUNRUI.BK4 = JAN_BUNRUIRS.rdoColumns("BK4").Value
			Case 24
				FILGET_JAN_BUNRUI = False
				JAN_BUNRUI.BK4 = ""
			Case Else
				FILGET_JAN_BUNRUI = False
				ERRSW = F_ERR
		End Select
		On Error GoTo 0 ' (�װ�ׯ�߉���)
		
	End Function
	'A-CUST20130212��
	
	'A-CUST-20170203 Start
	'�i�`�m�ϊ��e�[�u��
	Private Sub PREP_JAN_HENKAN_RTN()
		
		'�i�`�m�ϊ��e�[�u��
		MKKCMN.ZAEV_FNO = "JAN_HENKAN"
		Call MKKCMN.ZAEV_SUB()
		If MKKCMN.ZAEV_ERR <> "0" Then
			ERRSW = F_ERR
			Exit Sub
		Else
			JAN_HENKAN_FILE.NAME = MKKCMN.ZAEV_FNM
		End If
		
		'SELECT
		SQL = "Select"
		SQL = SQL & " Inc_code"
		SQL = SQL & " from "
		SQL = SQL & RTrim(JAN_HENKAN_FILE.NAME) & "JAN_HENKAN"
		SQL = SQL & " Where Inc_code = ? "
		SQL = SQL & " and jg_code = ? "
		SQL = SQL & " and hin_code = ? "
		SQL = SQL & " and renban <> 0 "
		SQL = SQL & " and jan_code = ? "
		On Error Resume Next
		JAN_HENKANSEL1 = ZACN_RCN.CreateQuery("JAN_HENKANSEL1", SQL)
		JAN_HENKANSEL1.QueryTimeout = 0
		JAN_HENKANSEL1.rdoParameters(0).NAME = "Inc_code" : JAN_HENKANSEL1.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL1.rdoParameters(0).Size = 2
		JAN_HENKANSEL1.rdoParameters(1).NAME = "jg_code" : JAN_HENKANSEL1.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL1.rdoParameters(1).Size = 4
		JAN_HENKANSEL1.rdoParameters(2).NAME = "hin_code" : JAN_HENKANSEL1.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL1.rdoParameters(2).Size = 5
		JAN_HENKANSEL1.rdoParameters(3).NAME = "jan_code" : JAN_HENKANSEL1.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL1.rdoParameters(3).Size = 13
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "JAN_HENKAN"
			GoTo PREPJAN_HENKAN_ERR
		End If
		On Error GoTo 0
		
		'SELECT
		SQL = "Select"
		SQL = SQL & " Inc_code"
		SQL = SQL & " from "
		SQL = SQL & RTrim(JAN_HENKAN_FILE.NAME) & "JAN_HENKAN"
		SQL = SQL & " Where Inc_code = ? "
		SQL = SQL & " and jg_code = ? "
		SQL = SQL & " and hin_code = ? "
		On Error Resume Next
		JAN_HENKANSEL2 = ZACN_RCN.CreateQuery("JAN_HENKANsel2", SQL)
		JAN_HENKANSEL2.QueryTimeout = 0
		JAN_HENKANSEL2.rdoParameters(0).NAME = "Inc_code" : JAN_HENKANSEL2.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL2.rdoParameters(0).Size = 2
		JAN_HENKANSEL2.rdoParameters(1).NAME = "jg_code" : JAN_HENKANSEL2.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL2.rdoParameters(1).Size = 4
		JAN_HENKANSEL2.rdoParameters(2).NAME = "hin_code" : JAN_HENKANSEL2.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANSEL2.rdoParameters(2).Size = 5
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "JAN_HENKAN"
			GoTo PREPJAN_HENKAN_ERR
		End If
		On Error GoTo 0
		
		Dim IDX As Short
		
		'UPDATE
		SQL = "UPDATE " & RTrim(JAN_HENKAN_FILE.NAME) & "JAN_HENKAN SET "
		SQL = SQL & "jan_code = ?, "
		SQL = SQL & "jan_hinname = ?, "
		SQL = SQL & "edit_op_code = ?, "
		SQL = SQL & "edit_op_date = ?, "
		SQL = SQL & "edit_op_time = ? "
		SQL = SQL & " where Inc_code = ? "
		SQL = SQL & " and jg_code = ? "
		SQL = SQL & " and hin_code = ? "
		SQL = SQL & " and renban = 0 "
		On Error Resume Next
		JAN_HENKANUPD = ZACN_RCN.CreateQuery("JAN_HENKANUPD", SQL)
		JAN_HENKANUPD.QueryTimeout = ZACN_TIME '�^�C���A�E�g���u�����v�ɐݒ�
		
		IDX = -1
		
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "jan_code" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 13
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "jan_hinname" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 20
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "Edit_Op_code" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 6
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "Edit_Op_date" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 8
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "Edit_Op_time" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 6
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "Inc_code" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 2
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "jg_code" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 4
		IDX = IDX + 1 : JAN_HENKANUPD.rdoParameters(IDX).NAME = "hin_code" : JAN_HENKANUPD.rdoParameters(IDX).Type = RDO.DataTypeConstants.rdTypeCHAR : JAN_HENKANUPD.rdoParameters(IDX).Size = 5
		If B_STATUS <> 0 Then
			ZAER_NO.Value = "JAN_HENKAN"
			GoTo PREPJAN_HENKAN_ERR
		End If
		On Error GoTo 0
		
		Exit Sub
		
PREPJAN_HENKAN_ERR: 
		ZAER_FID = "RAZ99"
		ZAER_KN = 1
		Call ZAER_SUB()
		ERRSW = F_ERR
		ENDSW = F_END
		On Error GoTo 0
		
	End Sub
	
	Public Sub FILGET_JAN_HENKAN_1(ByVal strInc_code As String, ByVal strJg_code As String, ByVal strHin_code As String, ByVal strJan_code As String)
		
		JAN_HENKANSEL1.rdoParameters("Inc_code").Value = strInc_code
		JAN_HENKANSEL1.rdoParameters("jg_code").Value = strJg_code
		JAN_HENKANSEL1.rdoParameters("hin_code").Value = strHin_code
		JAN_HENKANSEL1.rdoParameters("jan_code").Value = strJan_code
		
		On Error Resume Next ' (�װ���ׯ��)
		If JAN_HENKANRSSW <> "JAN_HENKANSEL1" Or ReQue = False Then
			JAN_HENKANRS = JAN_HENKANSEL1.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			JAN_HENKANRSSW = "JAN_HENKANSEL1"
		Else
			JAN_HENKANRS.Requery()
		End If
		
		Select Case B_STATUS(JAN_HENKANRS)
			Case 0
				JAN_HENKANINVSW = F_GET
			Case 24
				JAN_HENKANINVSW = F_INV
			Case Else
				JAN_HENKANINVSW = F_OFF
				ZAER_KN = 1
				Call ZAER_SUB()
				ERRSW = F_ERR
				ENDSW = F_END
		End Select
		On Error GoTo 0
		
	End Sub
	
	Public Sub FILGET_JAN_HENKAN_2(ByVal strInc_code As String, ByVal strJg_code As String, ByVal strHin_code As String)
		
		JAN_HENKANSEL2.rdoParameters("Inc_code").Value = strInc_code
		JAN_HENKANSEL2.rdoParameters("jg_code").Value = strJg_code
		JAN_HENKANSEL2.rdoParameters("hin_code").Value = strHin_code
		
		On Error Resume Next ' (�װ���ׯ��)
		If JAN_HENKANRSSW <> "JAN_HENKANSEL2" Or ReQue = False Then
			JAN_HENKANRS = JAN_HENKANSEL2.OpenResultset() '�iSQL�����s���A�₢�������ʂ����ʾ�ĂɊi�[����)
			JAN_HENKANRSSW = "JAN_HENKANSEL2"
		Else
			JAN_HENKANRS.Requery()
		End If
		
		Select Case B_STATUS(JAN_HENKANRS)
			Case 0
				JAN_HENKANINVSW = F_GET
			Case 24
				JAN_HENKANINVSW = F_INV
			Case Else
				JAN_HENKANINVSW = F_OFF
				ZAER_KN = 1
				Call ZAER_SUB()
				ERRSW = F_ERR
				ENDSW = F_END
		End Select
		On Error GoTo 0
		
	End Sub
	
	Public Sub UPD_JAN_HENKAN()
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("Inc_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Inc_code, 2) '��к���
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("jg_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jg_code, 4) '���Ə�����
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5) '�i��
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("jan_code").Value = MKKCMN.ZACHGSTR_SUB(KB.jan_code, 13) 'JAN�W������
		If RTrim(KB.hin_name_seisiki) = "" Then
			JAN_HENKANUPD.rdoParameters("jan_hinname").Value = " "
		Else
			'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			JAN_HENKANUPD.rdoParameters("jan_hinname").Value = RTrim(MKKCMN.ZACHGSTR_SUB(KB.hin_name_seisiki, 20)) '�i�`�m���i��
		End If
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("Edit_Op_code").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_code, 6) '�C���I�y���[�^
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("Edit_Op_date").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_date, 8) '�C���n��_date
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_HENKANUPD.rdoParameters("Edit_Op_time").Value = MKKCMN.ZACHGSTR_SUB(KB.Edit_Op_time, 6) '�C���n��_time
		
		On Error Resume Next
		Call JAN_HENKANUPD.Execute()
		Select Case B_STATUS
			Case n0
			Case Else
				ZAER_CD = B_STATUS
				ZAER_KN = 1
				ZAER_NO.Value = "JAN_HENKAN"
				ZAER_MS.Value = KB.Inc_code & "-" & KB.jg_code & "-" & KB.hin_code
				ENDSW = F_END
				ERRSW = F_ERR
				Call ZAER_SUB()
		End Select
		
		On Error GoTo 0
	End Sub
	'A-CUST-20170203e
	
	
	Public Sub SCR_DSPTAX() 'A-20190601
		'����ŗ��擾
		
		Dim strTax As String
		
		If RTrim(KB.hin_code) = "" And RTrim(KB.l_bun_code) = "" And RTrim(KB.m_bun_code) = "" Then
			strTax = ""
			GoTo SCR_DSPTAX_END
		End If
		
		CMTAX.CMTAX_RCN = ZACN_RCN
		CMTAX.CMTAX_TIME = CInt(WG_TIMEOUT)
		
		CMTAX.CMTAX_INC_CODE = WKB010
		CMTAX.CMTAX_JG_CODE = WKB020
		CMTAX.CMTAX_HIN_CODE = KB.hin_code
		CMTAX.CMTAX_DATE = GETTODAY()
		CMTAX.CMTAX_D_CODE = KB.l_bun_code
		CMTAX.CMTAX_C_CODE = KB.m_bun_code
		CMTAX.CMTAX_KAZEI_KBN = KB.Tax_kubn
		CMTAX.CMTAX_TAX_KBN = KB.tax_rate_kbn
		
		Call CMTAX.CMTAX_SUB()
		
		Select Case Val(CMTAX.CMTAX_TAX_CODE)
			Case 2
				strTax = VB6.Format("*" & VB6.Format(CMTAX.CMTAX_TAX * 100, "#0.00"), "@@@@@@") & "%"
			Case Else
				strTax = VB6.Format(" " & VB6.Format(CMTAX.CMTAX_TAX * 100, "#0.00"), "@@@@@@") & "%"
		End Select
		
SCR_DSPTAX_END: 
		With SZ0410FRM
			.DSP230A.Text = strTax
			.DSP440A.Text = strTax
		End With
		
	End Sub
	
	
	'A-20250303��
	'�i�`�m�`�F�b�N����
	Private Sub PREP_JAN_CHK_RTN()
		
		'SELECT LOCK
		SQL = "Select  "
		SQL = SQL & "hin_code" '�i��"
		SQL = SQL & " from "
		SQL = SQL & RTrim(SZM0010_FILE.NAME) & "SZM0010 "
		SQL = SQL & " WHERE Inc_code = ? "
		SQL = SQL & " AND jg_code = ? "
		SQL = SQL & " AND jan_code = ? "
		SQL = SQL & " AND hin_code <> ? "
		SQL = SQL & " AND DEL_FLG ='0' "
		
		On Error Resume Next
		JAN_CHK_SEL = ZACN_RCN.CreateQuery("JAN_CHK_SEL", SQL)
		JAN_CHK_SEL.QueryTimeout = ZACN_TIME
		If B_STATUS <> 0 Then
			MsgBox("JAN_CHK_SEL CreateQuery Error")
			GoTo PREP_JAN_CHK_RTN_ERR
		End If
		On Error GoTo 0
		
		JAN_CHK_SEL.rdoParameters(0).NAME = "Inc_code"
		JAN_CHK_SEL.rdoParameters(0).Type = RDO.DataTypeConstants.rdTypeCHAR
		JAN_CHK_SEL.rdoParameters(0).Size = 2
		JAN_CHK_SEL.rdoParameters(1).NAME = "jg_code"
		JAN_CHK_SEL.rdoParameters(1).Type = RDO.DataTypeConstants.rdTypeCHAR
		JAN_CHK_SEL.rdoParameters(1).Size = 4
		JAN_CHK_SEL.rdoParameters(2).NAME = "jan_code"
		JAN_CHK_SEL.rdoParameters(2).Type = RDO.DataTypeConstants.rdTypeCHAR
		JAN_CHK_SEL.rdoParameters(2).Size = 13
		JAN_CHK_SEL.rdoParameters(3).NAME = "hin_code"
		JAN_CHK_SEL.rdoParameters(3).Type = RDO.DataTypeConstants.rdTypeCHAR
		JAN_CHK_SEL.rdoParameters(3).Size = 5
		
		Exit Sub
		
PREP_JAN_CHK_RTN_ERR: 
		ERRSW = F_ERR
		On Error GoTo 0
		
	End Sub
	
	Public Function CHK_JANCODE(ByRef SJANCODE As String) As String
		'*************************************************
		'�I�y���[�^�}�X�^�����`�F�b�N
		'*************************************************
		
		' ���l���i�[����...
		On Error Resume Next
		JAN_CHK_SEL.rdoParameters("Inc_code").Value = WKB010
		JAN_CHK_SEL.rdoParameters("jg_code").Value = WKB020
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_CHK_SEL.rdoParameters("jan_code").Value = MKKCMN.ZACHGSTR_SUB(SJANCODE, 13)
		'UPGRADE_WARNING: �I�u�W�F�N�g MKKCMN.ZACHGSTR_SUB() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		JAN_CHK_SEL.rdoParameters("hin_code").Value = MKKCMN.ZACHGSTR_SUB(KB.hin_code, 5)
		' DB �ɖ₢���킹��...
		If JAN_CHKRSSW <> "JAN_CHK_SEL" Or (ZACN_DB = ORCL And ReQue = False) Then
			JAN_CHKRS = JAN_CHK_SEL.OpenResultset()
			JAN_CHKRSSW = "JAN_CHK_SEL"
		Else
			JAN_CHKRS.Requery()
		End If
		'�₢�������ʂ̔���
		If B_STATUS(JAN_CHKRS) = 24 Then
			CHK_JANCODE = ""
		Else
			If B_STATUS <> n0 Then
				CHK_JANCODE = ""
			Else
				CHK_JANCODE = JAN_CHKRS.rdoColumns("hin_code").Value
				Exit Function
			End If
		End If
		On Error GoTo 0
		
	End Function
	'A-20250303��
End Module