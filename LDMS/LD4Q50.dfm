�
 TFRMLD4Q50 0�  TPF0�
TfrmLD4Q50	frmLD4Q50Left>Top� Width�HeightICaption$frmLD4Q50-URIN ���ֺ� ��� ���� ��ȸPixelsPerInch`
TextHeight �TBevelBevel1Width�  �TLabellabRelationLeft�EnabledVisible  �TGaugeGrideLeft
TopWidth�HeightColorclAqua	ForeColorclRedParentColorProgress   �TPanelpnlCondLeft	TopWidth�Height*
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TLabelLabel7Left'Top
Width<HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel2LeftTopWidth7HeightCaption
�� �� �� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TSpeedButtonbtnBdateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel4Left� Top	Width<HeightCaption
���ֹ�ȣ :Font.CharsetHANGEUL_CHARSET
Font.ColorclBlackFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel6Left,TopWidthHeightCaption��Font.CharsetHANGEUL_CHARSET
Font.ColorclBlackFont.Height�	Font.Name����
Font.StylefsBold 
ParentFont  	TComboBoxcmbJisaTagLefteTopWidthuHeightColor��� DropDownCountImeName�ѱ���(�ѱ�)
ItemHeightTabOrder   	TDateEdit	msk_BgmgnLeftETopWidthQHeightColor��� EditMask9999/99/99;0;_Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength

ParentFontTabOrderYear Month Day   TRadioButtonRBtn_preveiwLeft<TopWidthIHeightCaption�̸�����Checked	TabOrderTabStop	  TRadioButtonRadioButton2Left�TopWidth6HeightCaption���TabOrder  	TMaskEditmskFromLeft� TopWidth%HeightEditMask9999;1;_ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLengthTabOrderText      	TMaskEditmskToLeftBTopWidth%HeightEditMask9999;1;_ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLengthTabOrderText       �TtsGrid	grdMasterLeftTopHWidth�Height�CheckBoxStylestCheckColsDefaultRowHeight-HeadingHorzAlignment	htaCenterHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclNavyHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightHeadingParentFontHeadingVertAlignment	vtaCenterParentShowHintRows RowSelectModersSingleShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedColPropertiesDataColCol.Heading���ֹ�ȣCol.HorzAlignment	htaCenterCol.VertAlignment	vtaCenter	Col.Width? DataColCol.Heading$��  ��  ��  ��  (  ��  ��  ��  ��  )	Col.WidthK    �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TSpeedButtonbtnSaveEnabled  �TSpeedButtonbtnPrintTagOnClickbtnPrintClick  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Left}Width,CaptionURIN ���ֺ� ��� ���� ��ȸ  �	TComboBoxcmbRelationLeft/Visible  TQuery
qry_gumginDatabaseNamedatabaseSQL.Strings  Left� Topv  TQueryqry_HangmokDatabaseNamedatabaseSQL.Stringsselect * from hangmokwhere cod_hm = :cod_hm Left*Topx	ParamDataDataType	ftUnknownNamecod_hm	ParamType	ptUnknown    TQuery
qry_gulkwaDatabaseNamedatabaseSQL.StringsSELECT *FROM  gulkwaWHERE num_id   = :num_idAND   dat_gmgn = :dat_gmgn Left� Topz	ParamDataDataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_Gmgn	ParamType	ptUnknown    TQueryqryNo_hangmokDatabaseNamedatabaseSQL.StringsSELECT * FROM no_hangmokWHERE  dat_apply  <= :dat_applyAND      gubn_code = :gubn_codeAND      gubn_yh    = :gubn_yhORDER BY dat_apply DESC LeftFTopx	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamegubn_yh	ParamType	ptUnknown    TQueryqryTkgumDatabaseNamedatabaseSQL.Stringsselect * from tkgumwhere cod_jisa = :cod_jisa   and num_jumin = :num_jumin   and dat_gmgn = :dat_gmgn Left� Top� 	ParamDataDataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName	num_jumin	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown    TQueryqryPf_hmDatabaseNamedatabaseSQL.Strings%select cod_pf, cod_hm from profile_Hmwhere cod_pf = :cod_pf Left� Top� 	ParamDataDataType	ftUnknownNamecod_pf	ParamType	ptUnknown    TQuery
qryProfileDatabaseNamedatabaseSQL.StringsSelect Cod_hm from Profile_hmWhere Cod_pf = :In_Cod_pf Left,Top� 	ParamDataDataType	ftUnknownName	In_Cod_pf	ParamType	ptUnknown     