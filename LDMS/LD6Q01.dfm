�
 TFRMLD6Q01 0�  TPF0�
TfrmLD6Q01	frmLD6Q01LeftkTop� CaptionfrmLD6Q01-�̼Ұ��� ListPixelsPerInch`
TextHeight �TLabellabRelationEnabledVisible  �TPanelpnlCondLeft	TopWidth	Height#
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TLabelLabel2LeftTopWidth<HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclTealFont.Height�	Font.Name����
Font.Style 
ParentFont  TSpeedButtonbtnDateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TSpeedButtonbtnDATE1LeftTopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel1Left� TopWidthHeightCaption~Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.StylefsBold 
ParentFont  	TDateEdit
mskST_dateLeftKTop	WidthMHeightColor��� EditMask9999/99/99;0;_ImeName�ѱ���(�ѱ�)	MaxLength
TabOrder OnKeyUpUP_KeyUpYear Month Day   	TDateEdit
mskED_dateLeft� Top	WidthMHeightColor��� EditMask9999/99/99;0;_ImeName�ѱ���(�ѱ�)	MaxLength
TabOrderOnKeyUpUP_KeyUpYear Month Day   TPanelpnlJisaLeft?TopWidth� Height"
BevelOuter	bvLoweredCaptionpnlJisaColor��� TabOrder TLabelLabel3Left#Top
WidthHeightCaption����Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  	TComboBoxcmbJisaLeftOTopWidthmHeightColor��� DropDownCountImeName�ѱ���(�ѱ�)
ItemHeightTabOrder     �TtsGrid	grdMasterLeft	TopGWidth	Height�Hint����ListCheckBoxStylestCheckColsDefaultRowHeightGridMode	gmListBoxHeadingHorzAlignment	htaCenterHeadingButtonhbCellHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWindowTextHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightRowsRowSelectModersSingleTabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedOnHeadingClickgrdMasterHeadingClickOnRowChangedgrdMasterRowChangedColPropertiesDataCol	Col.Width5 DataCol	Col.WidthM DataCol	Col.Widthf DataCol	Col.Width�  DataCol	Col.Width_ DataCol	Col.Widthl DataCol	Col.Width�     �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TToolButtonToolButton3EnabledVisible  �TSpeedButtonbtnSaveEnabledVisible  �TSpeedButton	btnCancelEnabledVisible  �TSpeedButtonbtnFindEnabledVisible  �TSpeedButtonbtnPrintTagOnClickbtnPrintClick  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Caption�� �� �� �� L I S T  �	TComboBoxcmbRelationVisible  TQuery	qrygumginDatabaseNamedatabaseSQL.Strings7select  a.cod_bjjs,  a.dat_bunju, a.cod_jisa, a.num_id,=           a.dat_gmgn, a.num_bunju, b.desc_name,  b.num_jumin)from   bunju a  left  outer join gumgin b    on   a.num_id   =  b.num_id     and  a.dat_gmgn = b.dat_gmgn#    and  a.cod_jisa   =  b.cod_jisawhere  a.dat_bunju >= :st_date   and  a.dat_bunju <= :ed_date"   and  a.cod_jisa   =   :cod_jisa   and  b.ysno_sokun = 'N' 
UpdateModeupWhereKeyOnlyLeftiTopr	ParamDataDataType	ftUnknownNamest_date	ParamType	ptUnknown DataType	ftUnknownNameed_date	ParamType	ptUnknown DataType	ftUnknownNamecod_jisa	ParamType	ptUnknown     