�
 TFRMLD4Q01 0P  TPF0�
TfrmLD4Q01	frmLD4Q01Left/Top� HeightCaptionfrmLD4Q01-���װ����ScaledPixelsPerInch`
TextHeight �TLabellabRelationEnabledVisible  �TPanelpnlCondLeft	TopWidth	Height
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TSpeedButton
btnSt_dateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TSpeedButton
btnEd_dateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel10Left� TopWidth
HeightCaption-Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.StylefsBold 
ParentFont  TLabelLabel11LeftTopWidth4HeightCaption��������Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TGaugeSPrctLeft�TopWidthHeightAlignalRightColor �� 	ForeColor �� ParentColorProgress   TLabelLabel1Left�TopWidth4HeightCaption�������Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  	TDateEdit
mskST_dateLeft<TopWidthMHeightColor��� EditMask9999/99/99;0;_ImeName�ѱ���(�ѱ�)	MaxLength
TabOrder OnKeyUpUP_KeyUpYear Month Day   	TDateEdit
mskEd_dateLeft� TopWidthNHeightColor��� EditMask9999/99/99;0;_ImeName�ѱ���(�ѱ�)	MaxLength
TabOrderOnKeyUpUP_KeyUpYear Month Day    �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TToolButtonToolButton3EnabledVisible  �TSpeedButtonbtnSaveEnabledVisible  �TSpeedButton	btnCancelEnabledVisible  �TSpeedButtonbtnFindEnabled  �TSpeedButtonbtnPrintTagEnabled  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Caption���װ��  ��(��������)  �	TComboBoxcmbRelationVisible  TtsGrid	grdMasterLeftTop:WidthHeight�CheckBoxStylestCheckColsDefaultRowHeightFont.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.Style GridMode	gmListBoxHeadingHorzAlignment	htaCenterHeadingButtonhbCellHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWindowTextHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeight
ParentFontParentShowHintRowsRowSelectModersSingleShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedColPropertiesDataCol	Col.Width2    TQueryqryBunjuDatabaseNamedatabaseSQL.StringsSelect * From BunjuWhere Cod_bjjs = '15'  And Dat_Bunju >= :In_St_Date  And Dat_Bunju <= :In_Ed_Date LefttTopx	ParamDataDataType	ftUnknownName
In_St_Date	ParamType	ptUnknown DataType	ftUnknownName
In_Ed_Date	ParamType	ptUnknown    TQuery	qryGumginDatabaseNamedatabaseSQL.Strings-Select Num_Jumin, Num_Id, Desc_Name, Dat_GmgnFrom GumginWhere Cod_Jisa = :In_Cod_jisa    And Num_id = :In_Num_id    And (Cod_Jangbi Like 'SS%'    Or    Cod_Hul <> ''    OR   Cod_Chuga <> '')Order By Dat_gmgn Left� Topz	ParamDataDataType	ftUnknownNameIn_Cod_jisa	ParamType	ptUnknown DataType	ftUnknownName	In_Num_id	ParamType	ptUnknown    TQuery	qryGulkwaDatabaseNamedatabaseSQL.StringsSelect * From GulkwaWhere Cod_bjjs = '15'   And NUm_id = :In_Num_id    And Gubn_Part = :In_Gubn_PartOrder By Dat_Gmgn Desc Left� Top~	ParamDataDataType	ftUnknownName	In_Num_id	ParamType	ptUnknown DataType	ftUnknownNameIn_Gubn_Part	ParamType	ptUnknown     