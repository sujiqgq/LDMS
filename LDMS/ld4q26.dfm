�
 TFRMLD4Q26 0  TPF0�
TfrmLD4Q26	frmLD4Q26LeftYTop� Width=Height�CaptionfrmLD4Q26- ä�� ��� ���� ��ȸPixelsPerInch`
TextHeight �TLabellabRelationEnabledVisible  �TPanelpnlCondLeft	TopWidthHeight(
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TPanel	pnlgumginLeft TopWidthHeight"
BevelInner	bvLoweredColor��� TabOrder  TLabelLabel1Left
Top	Width7HeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name	���� ����
Font.Style 
ParentFont  TSpeedButtonbtnsDateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TSpeedButton	btnCOD_DCLeftsTopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 33<��333?��ww�����̀����www��ww ����wwwwwwwp��ws�37�33sp������s7�37�37�p������s7�3w�37�0������37?3w33737����333?�337����333w�33;���33s�w�s3?� �� �337w7ww33;������333?�w�33?������333w�w�333�����3333w�w3333?����33337ws33333���33333333333	NumGlyphsOnClickUP_Click  TLabelLabel4LeftTop	Width3HeightCaption��     ü :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name	���� ����
Font.Style 
ParentFont  TLabelLabel7LefthTop	Width'HeightCaption��  �� :Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name	���� ����
Font.Style 
ParentFont  	TDateEdit	msk_SGmgnLeftETopWidthPHeightColor��� EditMask9999/99/99;0;_Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name	���� ����
Font.Style ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLength

ParentFontTabOrder Year Month Day   TEdit
edtDESC_DCLeft�TopWidth{HeightColor	clBtnFaceImeName�ѱ���(�ѱ�)ReadOnly	TabOrder  TEdit	edtCOD_DCLeft;TopWidth8HeightCharCaseecUpperCaseFont.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name	���� ����
Font.Style ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLength
ParentFontTabOrderOnChangeedtCOD_DCChange  	TComboBoxcmbJisaTagLeft�TopWidthiHeightColor��� DropDownCountFont.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name	���� ����
Font.Style ImeName�ѱ���(�ѱ�)
ItemHeight
ParentFontTabOrder    �TtsGrid	grdMasterLeftTopHWidth	Height�CheckBoxStylestCheckDefaultRowHeightFont.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name	���� ����
Font.Style HeadingHorzAlignment	htaCenterHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWindowTextHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeight
ParentFontParentShowHintRowsRowSelectModersSingleShowHint	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedColPropertiesDataCol	Col.Widthn DataCol	Col.WidthH DataCol	Col.Width�  DataCol	Col.Width�   CellPropertiesDataColDataRowCell.ControlTypectText    �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TSpeedButtonbtnSaveEnabled  �TSpeedButtonbtnPrintTagOnClickbtnPrintClick  �TSpeedButtonbtnExcelTag	   �TPanelPanel1Captionä�� ��� ���� ��ȸ  �	TComboBoxcmbRelationVisible  TQuery	qryGumginDatabaseNamedatabaseSQL.Strings  select*from gumgin with(nolock)" where cod_chuga like '%JJ1I%' and(         (cod_hul in (select cod_pf from profile with(nolock) where cod_pf in (select cod_pf from profile_hm with(nolock) where cod_hm in ('H001','H002','H003','H004','H005','H006','H007','H008','H009','H010','H011','H012','H013','H014','H015','H016','H017','H018','H019','H020','H021','H022'))) or+         (cod_jangbi in (select cod_pf from profile with(nolock) where cod_pf in (select cod_pf from profile_hm with(nolock) where cod_hm in ('H001','H002','H003','H004','H005','H006','H007','H008','H009','H010','H011','H012','H013','H014','H015','H016','H017','H018','H019','H020','H021','H022'))) or,         (cod_cancer in (select cod_pf from profile with(nolock) where cod_pf in (select cod_pf from profile_hm with(nolock) where cod_hm in ('H001','H002','H003','H004','H005','H006','H007','H008','H009','H010','H011','H012','H013','H014','H015','H016','H017','H018','H019','H020','H021','H022'))) or )         (cod_chuga like '%H001%') or (cod_chuga like '%H002%') or (cod_chuga like '%H003%') or (cod_chuga like '%H004%') or (cod_chuga like '%H005%') or (cod_chuga like '%H006%') or (cod_chuga like '%H007%') or (cod_chuga like '%H008%') or (cod_chuga like '%H009%') or (cod_chuga like '%H010%') or )         (cod_chuga like '%H011%') or (cod_chuga like '%H012%') or (cod_chuga like '%H013%') or (cod_chuga like '%H014%') or (cod_chuga like '%H015%') or (cod_chuga like '%H016%') or (cod_chuga like '%H017%') or (cod_chuga like '%H018%') or (cod_chuga like '%H019%') or (cod_chuga like '%H020%') or @       (cod_chuga like '%H021%') or (cod_chuga like '%H022%')))) LeftTop�    