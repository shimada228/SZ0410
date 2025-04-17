Option Strict Off
Option Explicit On
Module SpreadConst
	'----------------------------------------------------------
	' Program Name: SSOCX.BAS
	' Description : Spread v2.5J 定数定義ﾌｧｲﾙ
	' Make        : 1996-08-01
	' Version     : 1.0
	' Copyright (C) 1996 FarPoint Technologies.
	' All rights reserved.
	'----------------------------------------------------------
	
	
	'ｽﾌﾟﾚｯﾄﾞｼｰﾄの操作機能 (Action ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_ACTION_ACTIVE_CELL As Short = 0
	Public Const SS_ACTION_GOTO_CELL As Short = 1
	Public Const SS_ACTION_SELECT_BLOCK As Short = 2
	Public Const SS_ACTION_CLEAR As Short = 3
	Public Const SS_ACTION_DELETE_COL As Short = 4
	Public Const SS_ACTION_DELETE_ROW As Short = 5
	Public Const SS_ACTION_INSERT_COL As Short = 6
	Public Const SS_ACTION_INSERT_ROW As Short = 7
	Public Const SS_ACTION_RECALC As Short = 11
	Public Const SS_ACTION_CLEAR_TEXT As Short = 12
	Public Const SS_ACTION_PRINT As Short = 13
	Public Const SS_ACTION_DESELECT_BLOCK As Short = 14
	Public Const SS_ACTION_DSAVE As Short = 15
	Public Const SS_ACTION_SET_CELL_BORDER As Short = 16
	Public Const SS_ACTION_ADD_MULTISELBLOCK As Short = 17
	Public Const SS_ACTION_GET_MULTI_SELECTION As Short = 18
	Public Const SS_ACTION_COPY_RANGE As Short = 19
	Public Const SS_ACTION_MOVE_RANGE As Short = 20
	Public Const SS_ACTION_SWAP_RANGE As Short = 21
	Public Const SS_ACTION_CLIPBOARD_COPY As Short = 22
	Public Const SS_ACTION_CLIPBOARD_CUT As Short = 23
	Public Const SS_ACTION_CLIPBOARD_PASTE As Short = 24
	Public Const SS_ACTION_SORT As Short = 25
	Public Const SS_ACTION_COMBO_CLEAR As Short = 26
	Public Const SS_ACTION_COMBO_REMOVE As Short = 27
	Public Const SS_ACTION_RESET As Short = 28
	Public Const SS_ACTION_SEL_MODE_CLEAR As Short = 29
	Public Const SS_ACTION_VMODE_REFRESH As Short = 30
	Public Const SS_ACTION_SMARTPRINT As Short = 32
	
	'ﾌﾞﾛｯｸ選択範囲の設定 (SelectBlockOptions ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_SELBLOCKOPT_COLS As Short = 1
	Public Const SS_SELBLOCKOPT_ROWS As Short = 2
	Public Const SS_SELBLOCKOPT_BLOCKS As Short = 4
	Public Const SS_SELBLOCKOPT_ALL As Short = 8
	
	'ﾌｨｰﾙﾄﾞに対する列幅の調整(DAutoSizeCols ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_AUTOSIZE_NO As Short = 0
	Public Const SS_AUTOSIZE_MAX_COL_WIDTH As Short = 1
	Public Const SS_AUTOSIZE_BEST_GUESS As Short = 2
	
	'罫線と背景色の表示 (BackColorStyle ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_BACKCOLORSTYLE_OVERGRID As Short = 0
	Public Const SS_BACKCOLORSTYLE_UNDERGRID As Short = 1
	Public Const SS_BACKCOLORSTYLE_HORZGRIDONLY As Short = 2
	Public Const SS_BACKCOLORSTYLE_VERTGRIDONLY As Short = 3
	
	'ｾﾙ型の設定 (CellType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_TYPE_DATE As Short = 0
	Public Const SS_CELL_TYPE_EDIT As Short = 1
	Public Const SS_CELL_TYPE_FLOAT As Short = 2
	Public Const SS_CELL_TYPE_INTEGER As Short = 3
	Public Const SS_CELL_TYPE_PIC As Short = 4
	Public Const SS_CELL_TYPE_STATIC_TEXT As Short = 5
	Public Const SS_CELL_TYPE_TIME As Short = 6
	Public Const SS_CELL_TYPE_BUTTON As Short = 7
	Public Const SS_CELL_TYPE_COMBOBOX As Short = 8
	Public Const SS_CELL_TYPE_PICTURE As Short = 9
	Public Const SS_CELL_TYPE_CHECKBOX As Short = 10
	Public Const SS_CELL_TYPE_OWNER_DRAWN As Short = 11
	
	'ｾﾙの罫線の描画範囲 (CellBorderType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_BORDER_TYPE_NONE As Short = 0
	Public Const SS_BORDER_TYPE_OUTLINE As Short = 16
	Public Const SS_BORDER_TYPE_LEFT As Short = 1
	Public Const SS_BORDER_TYPE_RIGHT As Short = 2
	Public Const SS_BORDER_TYPE_TOP As Short = 4
	Public Const SS_BORDER_TYPE_BOTTOM As Short = 8
	
	'ｾﾙの罫線ｽﾀｲﾙ (CellBorderStyle ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_BORDER_STYLE_DEFAULT As Short = 0
	Public Const SS_BORDER_STYLE_SOLID As Short = 1
	Public Const SS_BORDER_STYLE_DASH As Short = 2
	Public Const SS_BORDER_STYLE_DOT As Short = 3
	Public Const SS_BORDER_STYLE_DASH_DOT As Short = 4
	Public Const SS_BORDER_STYLE_DASH_DOT_DOT As Short = 5
	Public Const SS_BORDER_STYLE_BLANK As Short = 6
	Public Const SS_BORDER_STYLE_FINE_SOLID As Short = 11
	Public Const SS_BORDER_STYLE_FINE_DASH As Short = 12
	Public Const SS_BORDER_STYLE_FINE_DOT As Short = 13
	Public Const SS_BORDER_STYLE_FINE_DASH_DOT As Short = 14
	Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT As Short = 15
	
	'列/行のﾀｲﾄﾙの設定 (ColHeaderDisplay/RowHeaderDisplay ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_HEADER_BLANK As Short = 0
	Public Const SS_HEADER_NUMBERS As Short = 1
	Public Const SS_HEADER_LETTERS As Short = 2
	
	'ﾁｪｯｸﾎﾞｯｸｽ型ｾﾙのﾃｷｽﾄの配置 (TypeCheckTextAlign ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CHECKBOX_TEXT_LEFT As Short = 0
	Public Const SS_CHECKBOX_TEXT_RIGHT As Short = 1
	
	'ﾏｳｽｶｰｿﾙの形状 (CursorStyle ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CURSOR_STYLE_USER_DEFINED As Short = 0
	Public Const SS_CURSOR_STYLE_DEFAULT As Short = 1
	Public Const SS_CURSOR_STYLE_ARROW As Short = 2
	Public Const SS_CURSOR_STYLE_DEFCOLRESIZE As Short = 3
	Public Const SS_CURSOR_STYLE_DEFROWRESIZE As Short = 4
	
	'ﾏｳｽﾎﾟｲﾝﾀの位置 (CursorType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CURSOR_TYPE_DEFAULT As Short = 0
	Public Const SS_CURSOR_TYPE_COLRESIZE As Short = 1
	Public Const SS_CURSOR_TYPE_ROWRESIZE As Short = 2
	Public Const SS_CURSOR_TYPE_BUTTON As Short = 3
	Public Const SS_CURSOR_TYPE_GRAYAREA As Short = 4
	Public Const SS_CURSOR_TYPE_LOCKEDCELL As Short = 5
	Public Const SS_CURSOR_TYPE_COLHEADER As Short = 6
	Public Const SS_CURSOR_TYPE_ROWHEADER As Short = 7
	
	'ｵﾍﾟﾚｰｼｮﾝﾓｰﾄﾞの設定 (OperationMode ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_OP_MODE_NORMAL As Short = 0
	Public Const SS_OP_MODE_READONLY As Short = 1
	Public Const SS_OP_MODE_ROWMODE As Short = 2
	Public Const SS_OP_MODE_SINGLE_SELECT As Short = 3
	Public Const SS_OP_MODE_MULTI_SELECT As Short = 4
	Public Const SS_OP_MODE_EXT_SELECT As Short = 5
	
	'ｿｰﾄ順序[昇順/降順] (SortKeyOrder ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_SORT_ORDER_NONE As Short = 0
	Public Const SS_SORT_ORDER_ASCENDING As Short = 1
	Public Const SS_SORT_ORDER_DESCENDING As Short = 2
	
	'ｿｰﾄ対象[列/行] (SortBy ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_SORT_BY_ROW As Short = 0
	Public Const SS_SORT_BY_COL As Short = 1
	
	'列幅/行の高さの変更の対象 (UserResize ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_USER_RESIZE_NONE As Short = 0
	Public Const SS_USER_RESIZE_COL As Short = 1
	Public Const SS_USER_RESIZE_ROW As Short = 2
	
	'列幅/行の高さの変更の可/不可 (UserResizeCol / UserResizeRow ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_USER_RESIZE_DEFAULT As Short = 0
	Public Const SS_USER_RESIZE_ON As Short = 1
	Public Const SS_USER_RESIZE_OFF As Short = 2
	
	'拡張ｽｸﾛｰﾙﾊﾞｰの表示 (VScrollSpecialType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_VSCROLLSPECIAL_NO_HOME_END As Short = 1
	Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN As Short = 2
	Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN As Short = 4
	
	'ｱｸﾃｨﾌﾞｾﾙのｼｰﾄ上の配置 (Position ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_POSITION_UPPER_LEFT As Short = 0
	Public Const SS_POSITION_UPPER_CENTER As Short = 1
	Public Const SS_POSITION_UPPER_RIGHT As Short = 2
	Public Const SS_POSITION_CENTER_LEFT As Short = 3
	Public Const SS_POSITION_CENTER_CENTER As Short = 4
	Public Const SS_POSITION_CENTER_RIGHT As Short = 5
	Public Const SS_POSITION_BOTTOM_LEFT As Short = 6
	Public Const SS_POSITION_BOTTOM_CENTER As Short = 7
	Public Const SS_POSITION_BOTTOM_RIGHT As Short = 8
	
	'ｽｸﾛｰﾙﾊﾞｰの設定 (ScrollBars ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_SCROLLBAR_NONE As Short = 0
	Public Const SS_SCROLLBAR_H_ONLY As Short = 1
	Public Const SS_SCROLLBAR_V_ONLY As Short = 2
	Public Const SS_SCROLLBAR_BOTH As Short = 3
	
	'印刷時の用紙の向き (PrintOrientation ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_PRINTORIENT_DEFAULT As Short = 0
	Public Const SS_PRINTORIENT_PORTRAIT As Short = 1
	Public Const SS_PRINTORIENT_LANDSCAPE As Short = 2
	
	'印刷範囲 (PrintType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_PRINT_ALL As Short = 0
	Public Const SS_PRINT_CELL_RANGE As Short = 1
	Public Const SS_PRINT_CURRENT_PAGE As Short = 2
	Public Const SS_PRINT_PAGE_RANGE As Short = 3
	
	'ﾎﾞﾀﾝ型ｾﾙのﾎﾞﾀﾝの種類 (TypeButtonType ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_BUTTON_NORMAL As Short = 0
	Public Const SS_CELL_BUTTON_TWO_STATE As Short = 1
	
	'ﾎﾞﾀﾝ型ｾﾙのﾋﾟｸﾁｬｰの配置 (TypeButtonAlign ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_BUTTON_ALIGN_BOTTOM As Short = 0
	Public Const SS_CELL_BUTTON_ALIGN_TOP As Short = 1
	Public Const SS_CELL_BUTTON_ALIGN_LEFT As Short = 2
	Public Const SS_CELL_BUTTON_ALIGN_RIGHT As Short = 3
	
	'ｺﾏﾝﾄﾞﾎﾞﾀﾝ/ｺﾝﾎﾞﾎﾞｯｸｽ型ｾﾙの表示 (ButtonDrawMode ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_BDM_ALWAYS As Short = 0
	Public Const SS_BDM_CURRENT_CELL As Short = 1
	Public Const SS_BDM_CURRENT_COLUMN As Short = 2
	Public Const SS_BDM_CURRENT_ROW As Short = 4
	
	'日付の表示形式 (TypeDateFormat ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_DATE_FORMAT_DDMONYY As Short = 0
	Public Const SS_CELL_DATE_FORMAT_DDMMYY As Short = 1
	Public Const SS_CELL_DATE_FORMAT_MMDDYY As Short = 2
	Public Const SS_CELL_DATE_FORMAT_YYMMDD As Short = 3
	Public Const SS_CELL_DATE_FORMAT_YYMM As Short = 4
	Public Const SS_CELL_DATE_FORMAT_MMDD As Short = 5
	Public Const SS_CELL_DATE_FORMAT_NYYMMDD As Short = 6
	Public Const SS_CELL_DATE_FORMAT_NNYYMMDD As Short = 7
	Public Const SS_CELL_DATE_FORMAT_NNNNYYMMDD As Short = 8
	
	'文字型ｾﾙの入力文字の変換 (TypeEditCharCase ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_EDIT_CASE_LOWER_CASE As Short = 0
	Public Const SS_CELL_EDIT_CASE_NO_CASE As Short = 1
	Public Const SS_CELL_EDIT_CASE_UPPER_CASE As Short = 2
	
	'文字型ｾﾙの入力文字種 (TypeEditCharSet ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_EDIT_CHAR_SET_ASCII As Short = 0
	Public Const SS_CELL_EDIT_CHAR_SET_ALPHA As Short = 1
	Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC As Short = 2
	Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC As Short = 3
	Public Const SS_CELL_EDIT_CHAR_SET_KANJI_ONLY As Short = 4
	Public Const SS_CELL_EDIT_CHAR_SET_KANJI_ONLY_IME As Short = 5
	Public Const SS_CELL_EDIT_CHAR_SET_ALL_IME As Short = 6
	
	'ﾗﾍﾞﾙ型ｾﾙの文字列の縦方向の配置 (TypeTextAlignVert ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM As Short = 0
	Public Const SS_CELL_STATIC_V_ALIGN_CENTER As Short = 1
	Public Const SS_CELL_STATIC_V_ALIGN_TOP As Short = 2
	
	'時刻の表示形式 (TypeTime24Hour ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_TIME_12_HOUR_CLOCK As Short = 0
	Public Const SS_CELL_TIME_24_HOUR_CLOCK As Short = 1
	Public Const SS_CELL_TIME_12_HOUR_CLOCK_AM As Short = 2
	Public Const SS_CELL_TIME_12_AM_HOUR_CLOCK As Short = 3
	
	'列/行の表示位置の設定単位 (Unittype ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_UNIT_NORMAL As Short = 0
	Public Const SS_CELL_UNIT_VGA As Short = 1
	Public Const SS_CELL_UNIT_TWIPS As Short = 2
	
	'ｾﾙ内でのﾃｷｽﾄの配置 (TypeHAlign ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_H_ALIGN_LEFT As Short = 0
	Public Const SS_CELL_H_ALIGN_RIGHT As Short = 1
	Public Const SS_CELL_H_ALIGN_CENTER As Short = 2
	
	'改行ｷｰ押下時の移動先ｾﾙ (EditEnterAction ﾌﾟﾛﾊﾟﾃｨ)
	Public Const SS_CELL_EDITMODE_EXIT_NONE As Short = 0
	Public Const SS_CELL_EDITMODE_EXIT_UP As Short = 1
	Public Const SS_CELL_EDITMODE_EXIT_DOWN As Short = 2
	Public Const SS_CELL_EDITMODE_EXIT_LEFT As Short = 3
	Public Const SS_CELL_EDITMODE_EXIT_RIGHT As Short = 4
	Public Const SS_CELL_EDITMODE_EXIT_NEXT As Short = 5
	Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS As Short = 6
	Public Const SS_CELL_EDITMODE_EXIT_SAME As Short = 7
	Public Const SS_CELL_EDITMODE_EXIT_NEXTROW As Short = 8
	
	'ﾕｰｻﾞ定義関数のﾊﾟﾗﾒｰﾀ型 (CFGetParamInfo ﾒｿｯﾄﾞ)
	Public Const SS_VALUE_TYPE_LONG As Short = 0
	Public Const SS_VALUE_TYPE_DOUBLE As Short = 1
	Public Const SS_VALUE_TYPE_STR As Short = 2
	Public Const SS_VALUE_TYPE_CELL As Short = 3
	Public Const SS_VALUE_TYPE_RANGE As Short = 4
	
	'ﾕｰｻﾞ定義関数の戻り値 (CFGetParamInfo ﾒｿｯﾄﾞ)
	Public Const SS_VALUE_STATUS_OK As Short = 0
	Public Const SS_VALUE_STATUS_ERROR As Short = 1
	Public Const SS_VALUE_STATUS_EMPTY As Short = 2
	
	'数式のｾﾙ参照形式 (GetRefStyle/SetRefStyle ﾒｿｯﾄﾞ)
	Public Const SS_REFSTYLE_DEFAULT As Short = 0
	Public Const SS_REFSTYLE_A1 As Short = 1
	Public Const SS_REFSTYLE_R1C1 As Short = 2
	
	'引数の数が可変のﾕｰｻﾞ定義関数の登録 (AddCustomFunctionExt ﾒｿｯﾄﾞ)
	Public Const SS_CUSTFUNC_WANTCELLREF As Short = 1
	Public Const SS_CUSTFUNC_WANTRANGEREF As Short = 2
End Module