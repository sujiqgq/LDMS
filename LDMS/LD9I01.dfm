�
 TFRMLD9I01 0  TPF0�
TfrmLD9I01	frmLD9I01Left�TopZWidthVHeightbCaption$LD9I01-���α׷��� ��ȸ�ڵ� �߰�/����PixelsPerInch`
TextHeight �TLabellabRelationLeft=EnabledVisible  �TPanelpnlCondLeftTopWidth.Height?
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthColor��� TabOrder  TLabelLabel1LeftTopWidth^HeightCaption���α׷� ���� :   TLabelLabel2LeftTop$WidthDHeightCaption��ȸ ���� :   TLabelLabel3LeftwTopWidth|HeightCaption��ȸ ���α׷� ���� :   	TComboBox
CB_ProgramLefteTopWidth� HeightImeNameMicrosoft Office IME 2007
ItemHeightItems.Strings1. �˻����2. ������Ȳ3. ���/���Է� ��ȸ4. �����˻��׸� ������Ȳ TabOrder OnChange	UP_Change  	TComboBoxCB_GroupLeftKTop Width� HeightImeNameMicrosoft Office IME 2007
ItemHeightTabOrder  TEdit
Edt_CreateLeft�TopWidthHeightColor	clBtnFaceCtl3D	EnabledImeNameMicrosoft Office IME 2007ParentCtl3DTabOrder   �TToolBarToolBar1TabOrder �TSpeedButtonbtnQueryOnClickbtnQueryClick  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertOnClickbtnInsertClick  �TSpeedButton	btnDeleteEnabledVisible  �TSpeedButtonbtnSaveEnabledVisible  �TSpeedButton	btnCancelOnClickbtnCancelClick  �TSpeedButtonbtnFindEnabledVisible  �TSpeedButtonbtnPrintTagEnabled  �TSpeedButtonbtnOpenOfficeVisible  �TSpeedButtonbtnExcelTag	Enabled   �TPanelPanel1LeftECaption��ȸ �˻��ڵ� �߰�/����  �	TComboBoxcmbRelationLeftwWidth� EnabledItems.Strings��ü���� �������������ü���� �˻���� Visible  TtsGrid	grdMasterLeftTopWidthCHeightjHint
���ָ���ƮCheckBoxStylestCheckDefaultRowHeightHeadingHorzAlignment	htaCenterHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWindowTextHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightParentShowHint	RowMovingRows ShowHint		StoreData	TabOrderThumbTracking	Version2.0OnCellLoadedgrdMasterCellLoadedOnRowChangedgrdMasterRowChangedColPropertiesDataColCol.Heading��Ʈ	Col.Width1 DataColCol.Heading�˻��ڵ�	Col.Width: DataColCol.Heading�˻��	Col.Widthq DataColCol.Heading��������  Data
             TPanelPanel2LeftTopfWidth� HeightCaption�˻��ڵ� ����ƮColorclPurpleFont.CharsetHANGEUL_CHARSET
Font.ColorclWhiteFont.Height�	Font.Name����
Font.StylefsBold 
ParentFontTabOrder  TPanelPanel3Left�TopfWidth� HeightCaption���α׷� ��ȸ �ڵ� ����ƮColorclPurpleFont.CharsetHANGEUL_CHARSET
Font.ColorclWhiteFont.Height�	Font.Name����
Font.StylefsBold 
ParentFontTabOrder  TButton	btn_chugaLeftOTopWidthKHeightCaption�߰� >>TabOrderOnClickbtn_chugaClick  TButton
btn_deleteLeftOTop<WidthKHeightCaption<< ����TabOrder	OnClickbtn_deleteClick  TtsGrid
grdProgramLeft�TopWidth�HeightjCheckBoxStylestCheckColsDefaultRowHeightHeadingFont.CharsetHANGEUL_CHARSETHeadingFont.ColorclWindowTextHeadingFont.Height�HeadingFont.Name����HeadingFont.Style HeadingHeightParentShowHintRows ShowHint	TabOrderVersion2.0OnCellLoadedgrdProgramCellLoadedColPropertiesDataColCol.Heading��ƮCol.HeadingHorzAlignment	htaCenter	Col.Width0 DataColCol.Heading�˻��ڵ�Col.HeadingHorzAlignment	htaCenterCol.HeadingVertAlignment	vtaCenter	Col.Width; DataColCol.Heading�˻��Col.HeadingHorzAlignment	htaCenter	Col.Widthw DataColCol.Heading�߰�����Col.HeadingHorzAlignment	htaCenterCol.HeadingVertAlignment	vtaCenter	Col.WidthI DataColCol.Heading�߰����Col.HeadingHorzAlignment	htaCenterCol.HeadingVertAlignment	vtaCenter    TBitBtnbtnPSaveTagLeft�Top�WidthqHeightHint�ڷḦ �����մϴ�.Caption����(&S)Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.Style 
ParentFontParentShowHintShowHint	TabOrder
OnClickbtnPSaveClick
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3333 pw 3333wwww3333 �� 3333w�3w3333 �� 3333w��w3333    ?���wwww        wwwwwwww������333337������?����� �̜�w7swwww����9�3?����w�  �𙙓ww77ww�������?���www �  9�3w7ww7w����9�3?��s7w���3y�3w7�?ww3�����3��swws3   33333www33333	NumGlyphs  TQuery
qryHangmokDatabaseNamedatabaseSQL.Strings5select H.gubn_part, H.cod_hm, H.desc_kor, H.dat_applyTfrom  (select ROW_NUMBER() OVER(PARTITION BY cod_hm ORDER BY dat_apply DESC) As SEQ,W       cod_hm, desc_kor, dat_apply, gubn_part, ysno_hidden from hangmok with(nolock)) HEwhere H.gubn_part in ('01', '02', '03', '04', '05', '07', '08', '09')and H.ysno_hidden='N'and H.SEQ='1' 
UpdateModeupWhereKeyOnlyLeftTop�  TQuery
qryHM_ListDatabaseNamedatabaseSQL.Strings$SELECT H.desc_kor, H.pos_start, L.* �FROM LDMS_HM  L inner join (select SEQ=ROW_NUMBER() OVER(PARTITION BY cod_hm ORDER BY dat_apply DESC), desc_kor, cod_hm, pos_start from hangmok) H on L.cod_hm = H.cod_hmWHERE SEQ='1'and  L.PROGRAM_ID = :PROGRAM_IDAND L.YSNO_HIDDEN='N'AND L.GROUP_HM = :GROUP_HMORDER BY L.POS_HM  
UpdateModeupWhereKeyOnlyLeft�Top�	ParamDataDataType	ftUnknownName
Program_ID	ParamType	ptUnknown DataType	ftUnknownNameGroup_HM	ParamType	ptUnknown    TQueryqryGroupDatabaseNamedatabaseSQL.Strings5SELECT PROGRAM_ID, GROUP_HM FROM LDMS_HM with(nolock)WHERE PROGRAM_ID = :PROGRAM_IDGROUP BY PROGRAM_ID, GROUP_HM Left8Top(	ParamDataDataType	ftUnknownName
PROGRAM_ID	ParamType	ptUnknown    TQueryqrySawonDatabaseNamedatabaseSQL.Strings(select desc_name from sawon with(nolock)where cod_sawon = :cod_sawon 
UpdateModeupWhereKeyOnlyLeft�Top�	ParamDataDataType	ftUnknownName	cod_sawon	ParamType	ptUnknown    TQueryqryIHM_ListDatabaseNamedatabaseSQL.Stringsinsert into LDMS_HMN(PROGRAM_ID, COD_HM, YSNO_HIDDEN, GUBN_PART, GROUP_HM, COD_INSERT, DAT_INSERT)values (J:program_id, :cod_hm, 'N', :gubn_part, :group_hm, :cod_insert, :dat_insert) 
UpdateModeupWhereKeyOnlyLefthTop� 	ParamDataDataType	ftUnknownName
Program_ID	ParamType	ptUnknown DataType	ftUnknownNamecod_hm	ParamType	ptUnknown DataType	ftUnknownName	gubn_part	ParamType	ptUnknown DataType	ftUnknownNamegroup_hm	ParamType	ptUnknown DataType	ftUnknownName
cod_insert	ParamType	ptUnknown DataType	ftUnknownName
dat_insert	ParamType	ptUnknown    TQueryqryUHM_ListDatabaseNamedatabaseSQL.Strings#update LDMS_HM set YSNO_HIDDEN='Y',0COD_UPDATE= :cod_update, DAT_UPDATE= :dat_updatewhere PROGRAM_ID= :program_id and GROUP_HM= :group_hm and COD_HM= :cod_hmand YSNO_HIDDEN='N' 
UpdateModeupWhereKeyOnlyLefthTopY	ParamDataDataType	ftUnknownName
cod_update	ParamType	ptUnknown DataType	ftUnknownName
dat_update	ParamType	ptUnknown DataType	ftUnknownName
program_id	ParamType	ptUnknown DataType	ftUnknownNamegroup_hm	ParamType	ptUnknown DataType	ftUnknownNamecod_hm	ParamType	ptUnknown    TQuery
qryU_PosHMDatabaseNamedatabaseSQL.Strings#update LDMS_HM set POS_HM = :pos_hmwhere PROGRAM_ID= :program_id and GROUP_HM= :group_hm and COD_HM= :cod_hmand YSNO_HIDDEN='N' 
UpdateModeupWhereKeyOnlyLeft�Top�	ParamDataDataType	ftUnknownNamepos_hm	ParamType	ptUnknown DataType	ftUnknownName
program_id	ParamType	ptUnknown DataType	ftUnknownNamegroup_hm	ParamType	ptUnknown DataType	ftUnknownNamecod_hm	ParamType	ptUnknown     