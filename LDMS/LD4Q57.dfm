�
 TFRMLD4Q57 0  TPF0�
TfrmLD4Q57	frmLD4Q57LeftITop� WidthjHeightCaption"frmLD4Q57_(��ȭ��)����˻��׸���ȸPixelsPerInch`
TextHeight �TBevelBevel1Width�  �TLabellabRelationLeft�EnabledVisible  �TGaugeGrideLeft	Top�WidthKHeight!ColorclAqua	ForeColorclRedParentColorProgress   �TPanelpnlCondLeft	TopWidthHHeight"
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TLabelLabel7Left^TopWidth<HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel2LeftTopWidth7HeightCaption
�� �� �� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TLabelLabel3Left� TopWidth<HeightCaption
���ֹ�ȣ :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFont  TSpeedButtonbtnBdateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel5Left!Top
WidthHeightCaption��Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.StylefsBold 
ParentFont  	TComboBoxcmbJisaTagLeft�TopWidth[HeightColor��� DropDownCountImeName�ѱ���(�ѱ�)
ItemHeightTabOrder  	TDateEditedtBDateLeft@TopWidthQHeightColor��� EditMask9999/99/99;0;_Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength

ParentFontTabOrder Year Month Day   TEditbunjuno1Left� TopWidth)HeightFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength
ParentFontTabOrder  TEditbunjuno2Left7TopWidth&HeightFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)	MaxLength
ParentFontTabOrder  TRadioButtonRBtn_preveiwLeft�TopWidthIHeightCaption�̸�����Checked	TabOrderTabStop	  TRadioButtonRadioButton2Left�TopWidthFHeightCaption���TabOrder   �TtsGrid	grdMasterLeft
Top@WidthGHeightzHint����������TableCheckBoxStylestCheckColsDefaultRowHeightHeadingHorzAlignment	htaCenterHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclNavyHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeight"HeadingParentFontHeadingVertAlignment	vtaCenterParentShowHintRows RowSelectModersSingleShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedColPropertiesDataColCol.HeadingNoCol.HorzAlignment	htaCenterCol.VertAlignment	vtaCenter	Col.Width( DataColCol.Heading���ù�ȣ DataColCol.Heading���ֹ�ȣ DataColCol.Heading�� ��	Col.WidthW DataColCol.Heading�ֹι�ȣ	Col.WidthW DataColCol.HeadingAMP	Col.Width9 DataColCol.HeadingCannabinoids	Col.Width] DataColCol.HeadingMorphine	Col.WidthH DataCol	Col.HeadingCocain DataCol
Col.HeadingMET DataColCol.HeadingHeroin DataColCol.HeadingOPI    �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TSpeedButtonbtnSaveEnabled  �TSpeedButtonbtnPrintTagOnClickbtnPrintClick  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Left@TopWidth9Caption��ȭ�� ����˻� �׸���ȸ  �	TComboBoxcmbRelationLeft�Visible  TQueryqryBunjuDatabaseNamedatabaseLeft� Top�   TQueryqryPf_hmDatabaseNamedatabaseSQL.StringsSELECT P.cod_hm, H.gubn_partBFROM profile_hm P left outer join Hangmok H ON P.cod_hm = H.cod_hmWHERE P.cod_pf = :cod_pf  
UpdateModeupWhereKeyOnlyLeftFTop� 	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQRCompositeReportQRCompositeReportOnAddReportsQRCompositeReportAddReportsOptions PrinterSettings.CopiesPrinterSettings.DuplexPrinterSettings.FirstPage PrinterSettings.LastPage PrinterSettings.OutputBinAutoPrinterSettings.Orientation
poPortraitPrinterSettings.PaperSizeqr11X17Left� Top�   TQueryqryTkgumDatabaseNamedatabaseSQL.StringsSELECT cod_prf, cod_chugaFROM   tkgumWHERE cod_jisa = :cod_jisaAND      num_jumin = :num_juminAND      dat_gmgn = :dat_gmgn 
UpdateModeupWhereKeyOnlyLeft� Top� 	ParamDataDataTypeftStringNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName	num_jumin	ParamType	ptUnknown DataTypeftStringNamedat_gmgn	ParamType	ptUnknown    TQueryqryNo_hangmokDatabaseNamedatabaseFiltered	SQL.StringsSELECT * FROM no_hangmokWHERE  dat_apply  <= :dat_applyAND      gubn_code = :gubn_codeAND      gubn_yh    = :gubn_yhORDER BY dat_apply DESC Left>Top� 	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamegubn_yh	ParamType	ptUnknown    TQuery	qryGulkwaDatabaseNamedatabaseSQL.Strings�SELECT K.gubn_part , G.desc_name, G.num_id, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, K.desc_glkwa1, K.desc_glkwa2, K.desc_glkwa3,�Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp,  K.num_bunju FROM Gulkwa K with (nolock)gleft outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn/left outer join Danche D ON G.cod_dc = D.cod_dc�left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_chawhere K.num_id = :num_id  and K.dat_bunju = :dat_bunju  and K.dat_gmgn = :dat_gmgn  and K.gubn_part =:gubn_partorder by K.dat_bunju desc    LeftTop� 	ParamDataDataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown DataType	ftUnknownName	gubn_part	ParamType	ptUnknown    TQuery
qryProfileDatabaseNamedatabaseSQL.StringsSelect Cod_hm from Profile_hmWhere Cod_pf = :In_Cod_pf Left� Top� 	ParamDataDataType	ftUnknownName	In_Cod_pf	ParamType	ptUnknown     