�
 TFRMLD8Q05 0�  TPF0�
TfrmLD8Q05	frmLD8Q05Left�ToptWidth�Height�ActiveControlmskFromCaption	frmLD8Q05PixelsPerInch`
TextHeight �TBevelBevel1LeftWidth�  �TLabellabRelationLeft�  �TPanelpnlCondLeftTopWidth�Height"
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder  TSpeedButtonbtnsDateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TLabelLabel2LeftTop	Width(HeightCaption������:Font.CharsetHANGEUL_CHARSET
Font.ColorclNavyFont.Height�	Font.Name����
Font.Style 
ParentFont  	TDateEditmskFromLeftETopWidthDHeightColor��� EditMask9999/99/99;0;_Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLength

ParentFontTabOrderYear Month Day   TPanelpnlJisaLeft�TopWidth� Height
BevelInnerbvRaised
BevelOuter	bvLoweredColor��� TabOrder  TLabelLabel1LeftTopWidth0HeightCaption��������Color��� Font.CharsetHANGEUL_CHARSET
Font.ColorclNavyFont.Height�	Font.Name����
Font.Style ParentColor
ParentFont  	TComboBoxcmbJisaTagLeft@TopWidthqHeightColor��� DropDownCountFont.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style ImeName�ѱ���(�ѱ�)
ItemHeightItems.Strings  
ParentFontTabOrder    	TCheckBoxGubn_hm1Left� TopWidthYHeightCaption
����ɰ˻�Font.CharsetHANGEUL_CHARSET
Font.ColorclNavyFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder  	TCheckBoxGubn_hm2LeftTopWidthIHeightCaption������Font.CharsetHANGEUL_CHARSET
Font.ColorclNavyFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder  	TCheckBoxGubn_hm3LefthTopWidthqHeightCaption�ڱð�ξϰ˻�Font.CharsetHANGEUL_CHARSET
Font.ColorclNavyFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder   �TToolBarToolBar1HeightButtonHeightTabOrder �TSpeedButtonbtnQueryHeightOnClickbtnQueryClick  �TSpeedButton	btnInsertHeightEnabled  �TSpeedButton	btnDeleteHeightEnabled  �TSpeedButtonbtnSaveHeightEnabled  �TSpeedButton	btnCancelHeightEnabled  �TSpeedButtonbtnFindHeight  �TSpeedButtonbtnPrintTagHeightEnabled  �TSpeedButtonbtnOpenOfficeHeight  �TSpeedButtonbtnExcelTag	Height  �TSpeedButtonbtnQuitHeight   �TPanelPanel1Left�Width� Caption�� �� �� �� �� ȸ  �	TComboBoxcmbRelationLeft�EnabledItems.Strings    TtsGrid	grdMasterLeftTopFWidth�HeightkHint������CheckBoxStylestCheckColsDefaultRowHeightFont.CharsetHANGEUL_CHARSET
Font.ColorclBlackFont.Height�	Font.Name����
Font.Style GridMode	gmListBoxHeadingHorzAlignment	htaCenterHeadingColor �� HeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWhiteHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightHeadingParentFont
ParentFontParentShowHintRows RowSelectModersSingleSelectionColorclBtnHighlightShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedOnHeadingClickgrdMasterHeadingClickOnRowChangedgrdMasterRowChangedColPropertiesDataColCol.Heading��������Col.HeadingButtoncbCellCol.SortPicturespDown	Col.WidthC DataColCol.Heading��������Col.HeadingButtoncbCellCol.SortPicturespDown	Col.WidthD DataColCol.Heading��    ��Col.HeadingButtoncbCellCol.SortPicturespDown	Col.WidthB DataColCol.Heading�ֹι�ȣCol.HeadingButtoncbCell	Col.Width] DataColCol.Heading��������	Col.WidthE DataColCol.Heading������Col.HeadingButtoncbCellCol.SortPicturespDown	Col.WidthQ DataColCol.Heading���ֹ�ȣCol.HeadingButtoncbCellCol.SortPicturespDown DataColCol.Heading������	Col.Width_ DataCol	Col.Heading�μ���	Col.Width[ DataCol
Col.Heading�̻��׸�Col.HeadingButtoncbCellCol.SortPicturespDown	Col.Widthh DataColCol.Heading��ġ	Col.WidthY    TQuery
qry_gulkwaDatabaseNamedatabaseSessionNameDefaultSQL.Strings8select G.cod_jisa, G.dat_gmgn, G.num_jumin, G.desc_name,:          K.cod_bjjs, K.dat_bunju, K.num_bunju, D.desc_dc,!          G.desc_dept, G.num_id, <          G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga!from bunju K  inner join gumgin G	on K.cod_jisa = G.cod_jisa	and K.num_id = G.num_id	and K.dat_gmgn = G.dat_gmgn	inner join Danche D	on G.cod_dc = D.cod_dc   
UpdateModeupWhereKeyOnlyLeftATop�   TQuery
qryPGulkwaDatabaseNamedatabaseFiltered	SQL.Stringsselect * from gulkwawhere cod_bjjs = :cod_bjjsand num_id = :num_idand dat_gmgn = :dat_gmgnorder by dat_bunju desc Left� Top� 	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown    TQueryqryPf_hmDatabaseNamedatabaseSQL.StringsSELECT P.cod_hm, H.gubn_partBFROM profile_hm P left outer join Hangmok H ON P.cod_hm = H.cod_hmWHERE P.cod_pf = :cod_pf  
UpdateModeupWhereKeyOnlyLeftNTop� 	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQueryqryCellDatabaseNamedatabaseSQL.Stringsselect * from cellwhere cod_bjjs = :cod_bjjsand num_id = :num_idand dat_gmgn = :dat_gmgn-and( (cod_hm = 'P003') or (cod_hm = 'P001') )and gubn_class = '3' Left� Top� 	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown    TQueryqry_gulkwa2DatabaseNamedatabaseFiltered	SessionNameDefaultSQL.Strings9select G.cod_jisa, G.dat_gmgn, G.num_jumin, G.desc_name, ;          K.cod_bjjs, K.dat_bunju, K.num_bunju, D.desc_dc, !          G.desc_dept, G.num_id, >          G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, "          C.gubn_class, C.cod_hm  &from bunju K  left outer join gumgin G	on K.cod_jisa = G.cod_jisa	and K.num_id = G.num_id	and K.dat_gmgn = G.dat_gmgn	left outer join Danche D	on G.cod_dc = D.cod_dc	left outer join Cell C	on K.cod_jisa = G.cod_jisa	and G.num_id = C.num_id	and G.dat_gmgn = C.dat_gmgn 
UpdateModeupWhereKeyOnlyLeft� Top�    