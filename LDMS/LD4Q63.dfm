�
 TFRMLD4Q63 0`  TPF0�
TfrmLD4Q63	frmLD4Q63LeftETop� Width$HeightCaption#frmLD4Q63-���� ���� B�� ��ȸ ����ƮPixelsPerInch`
TextHeight �TLabellabRelationEnabledVisible  �TGaugeGrideLeft
Top�WidthHeight!ColorclAqua	ForeColorclRedParentColorProgress   �TPanelpnlCondLeft	TopWidth	Height5
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TLabelLabel7Left2TopWidth"HeightCaption���� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel2LeftTopWidth<HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel3Left� TopWidth<HeightCaption
���ù�ȣ :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TSpeedButtonbtnBdateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel5Left.Top
WidthHeightCaption��Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.StylefsBold 
ParentFont  TLabelLabel1LeftxTopWidth/HeightCaption������ :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  	TComboBoxcmbJisaTagLeftYTopWidth[HeightColor��� DropDownCountEnabledImeName�ѱ���(�ѱ�)
ItemHeightTabOrder   	TDateEditedtDateLeftETopWidthQHeightColor��� EditMask9999/99/99;0;_Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength

ParentFontTabOrderYear Month Day   TEditSampNo1Left� TopWidth1HeightFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength
ParentFontTabOrder  TEditSampNo2Left>TopWidth1HeightFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength
ParentFontTabOrder  TRadioButtonRBtn_preveiwLeft�TopWidthIHeightCaption�̸�����Checked	TabOrderTabStop	  TRadioButtonRadioButton2Left�TopWidth� HeightCaption���TabOrder  	TComboBoxComB_ShFloorLeft�TopWidth|HeightFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)
ItemHeightItems.Strings  
ParentFontTabOrder  TPanelPanel2LeftTopWidth� HeightColor��� TabOrder TLabelLabel4LeftTopWidth<HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TRadioButtonrb_neaLeftGTopWidth1HeightCaption����TabOrder   TRadioButtonrb_chulLeft� TopWidth1HeightCaption����TabOrder  TRadioButtonrb_AllLeft� TopWidth1HeightCaption��üChecked	TabOrderTabStop	    �TtsGrid	grdMasterLeft
TopXWidthHeightpHint����������TableCheckBoxStylestCheckColsDefaultRowHeightHeadingHorzAlignment	htaCenterHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclNavyHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightHeadingParentFontHeadingVertAlignment	vtaCenterParentShowHintRows RowSelectModersSingleShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedOnRowChangedgrdMasterRowChangedColPropertiesDataColCol.HeadingNoCol.HorzAlignment	htaCenterCol.VertAlignment	vtaCenter	Col.Width( DataColCol.Heading���ù�ȣ	Col.WidthA DataColCol.Heading����	Col.WidthP DataColCol.Heading�������	Col.Widthx DataColCol.Heading����	Col.WidthX DataColCol.Heading������	Col.Widthn DataColCol.Heading���	Col.WidthO    �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TSpeedButtonbtnSaveEnabled  �TSpeedButtonbtnPrintTagOnClickbtnPrintClick  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Width� Caption���� ���� B�� ��ȸ ����Ʈ  �	TComboBoxcmbRelationVisible  TQuery	qryGumginDatabaseNamedatabaseLeft� Top�   TQueryqryPf_hmDatabaseNamedatabaseSQL.StringsSELECT P.cod_hm, H.gubn_partBFROM profile_hm P left outer join Hangmok H ON P.cod_hm = H.cod_hmWHERE P.cod_pf = :cod_pf  
UpdateModeupWhereKeyOnlyLeftFTop� 	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQueryqryNo_hangmokDatabaseNamedatabaseSQL.StringsSELECT * FROM no_hangmokWHERE  dat_apply  <= :dat_applyAND      gubn_code = :gubn_codeAND      gubn_yh    = :gubn_yhORDER BY dat_apply DESC Left� Top� 	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamegubn_yh	ParamType	ptUnknown    TQuery	qryPf_hm2DatabaseNamedatabaseSQL.StringsSELECT A.cod_hm, B.gubn_part+FROM   profile_hm A INNER JOIN hangmok B ON           A.cod_hm = B.cod_hmWHERE A.cod_pf = :cod_pf 
UpdateModeupWhereKeyOnlyLeft@Top� 	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQueryqryPart0507_01DatabaseNamedatabaseSQL.Strings0select distinct cod_hm from hangmok with(nolock)where gubn_part in('05','07')Yand cod_hm not in ('S016','S021','E068','T006','T007','SE02','TT02','T009','TT01','TT02')  Left� Top�   TQuery
qryProfileDatabaseNamedatabaseSQL.Strings2select cod_pf, Gubn_Gmsa from profile with(nolock)where cod_pf = :cod_Pf Left$Top� 	ParamDataDataType	ftUnknownNamecod_Pf	ParamType	ptUnknown    TQueryqryPart0507_02DatabaseNamedatabaseSQL.Strings0select distinct cod_hm from hangmok with(nolock)where gubn_part in('05','07')2and cod_hm not in ('S016', 'T006', 'TT01', 'TT02') LeftTop�    