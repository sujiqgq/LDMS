�
 TFRMLD2I04 0�:  TPF0�
TfrmLD2I04	frmLD2I04Left� Top� WidthZHeight9ActiveControlmskDateCaption frmLD2I04_���� ���װ˻� �߰�����PixelsPerInch`
TextHeight �TBevelBevel1LeftWidth  �TLabellabRelationEnabledVisible  �TPanel	pnlMasterLeft� Top� WidthYHeight� 
BevelInner	bvLoweredColor��� Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder TSpeedButtonbtnDateLeft� TopWidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 0      37wwwwww�0������37�3?�37�0�� ��37�3ws37�0������37�37�37�0�� ���37�3w�37�0������37�37337�0������37�������0 �����37w������0�� �p�37��w�w��0����p�37��77w��0�����37�3s77��0����  37�333ww�0�����37����730����37wwwws30���� 337����w330    337wwwws33	NumGlyphsOnClickUP_Click  TGaugeGaugeLeft&Top\WidthHeight	ForeColor��� Progress   TLabelLabel1Left'TopGWidth<HeightCaption
�۾����� :  TLabel	labStatusLeftlTopGWidth4HeightCaption�����Ȳ  TSpeedButtonbtnDcLeft� Top WidthHeight
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 33<��333?��ww�����̀����www��ww ����wwwwwwwp��ws�37�33sp������s7�37�37�p������s7�3w�37�0������37?3w33737����333?�337����333w�33;���33s�w�s3?� �� �337w7ww33;������333?�w�33?������333w�w�333�����3333w�w3333?����33337ws33333���33333333333	NumGlyphsOnClickUP_Click  TPanelPanel2LeftTopWidth^HeightCaption�۾�����Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder  	TDateEditmskDateLefthTopWidthPHeightColor��� EditMask9999/99/99;0;_ImeName�ѱ���(�ѱ�)	MaxLength
TabOrder OnKeyUpUP_KeyUpYear Month Day   TBitBtnbtnStartTagLeft� TopWidth[HeightHint0������� ����� �� �ֵ��� �ڷḦ �������ݴϴ�.Caption���ֽ���Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����
Font.Style 
ParentFontParentShowHintShowHint	TabOrderOnClickbtnStartClick
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 33   3333wwws33330��33333�ww3�3303033373ws733333033333�37333�033333037333337333333333?333333?333333s333333s33333333?333333?333333s33333�s333330 3?��337w�  33�ww33ww	�33�ww33ww	�33�ww33wws	�330 3ww337w3  33033wws���333300033333777333	NumGlyphs  TAnimateAnimate1Left&Top� WidthHeight<Active	CommonAVIaviCopyFile	StopFrame  TAnimateAnimate2Left� Top� Width0Height2Active	CommonAVIaviFindFile	StopFrame  TPanelPanel3Left Top8WidthXHeightTabOrder  TPanelPanel4LeftTop Width^HeightCaption��ü�ڵ�Font.CharsetHANGEUL_CHARSET
Font.ColorclMaroonFont.Height�	Font.Name����
Font.Style 
ParentFontTabOrder  TEditedtDcLeftiTop Width@HeightCharCaseecUpperCaseColor��� ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLengthTabOrderOnChange	UP_ChangeOnKeyUpUP_KeyUp  TEdit
edtDESC_DCLeft� Top Width� HeightColor	clBtnFaceImeName�ѱ���(�ѱ�)ReadOnly	TabOrder   �TToolBarToolBar1WidthTabOrder �TToolButtonToolButton1Visible  �TSpeedButtonbtnQueryEnabledVisible  �TToolButtonToolButton2EnabledVisible  �TSpeedButton	btnInsertEnabledVisible  �TSpeedButton	btnDeleteEnabledVisible  �TToolButtonToolButton3EnabledVisible  �TSpeedButtonbtnSaveEnabledVisible  �TSpeedButton	btnCancelEnabledVisible  �TToolButtonToolButton4EnabledVisible  �TSpeedButtonbtnFindEnabledVisible  �TSpeedButtonbtnPrintTagEnabledVisible  �TToolButtonToolButton5Visible  �TSpeedButtonbtnExcelTag	EnabledVisible   �TPanelPanel1LeftWidth� Height!Caption���� ���װ˻� �߰����� �۾�TabOrder   �	TComboBoxcmbRelationTabOrderVisible  TQuery	qryGumginDatabaseNamedatabaseSQL.StringsOSELECT cod_jisa, num_id, num_jumin, desc_name, cod_hul, cod_jangbi, cod_cancer,C            cod_chuga, gubn_hulgum, dat_gmgn, num_samp, gubn_nosin,D            gubn_tkgm, gubn_bogun, gubn_adult, gubn_agab, gubn_gong,B            gubn_nsyh, gubn_tkyh, gubn_bgyh, gubn_adyh, gubn_agyh,+            gubn_goyh, gubn_life, gubn_lfyhFROM   gumgin 
UpdateModeupWhereKeyOnlyLeftYTop*  TQuery
qryProfileDatabaseNamedatabaseSQL.StringsSELECT A.gubn_gmsa, B.cod_hm +FROM   profile A INNER JOIN profile_hm B ON           A.cod_pf = B.cod_pfWHERE A.cod_pf = :cod_pf 
UpdateModeupWhereKeyOnlyLeft9Topb	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQueryqryU_GumginDatabaseNamedatabaseSQL.StringsUPDATE gumgin"SET       ysno_bunju = :ysno_bunjuWHERE  cod_jisa = :cod_jisaAND      num_id = :num_idAND      dat_gmgn = :dat_gmgn 
UpdateModeupWhereKeyOnlyLeftaTop2	ParamDataDataType	ftUnknownName
ysno_bunju	ParamType	ptUnknown DataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown    TQueryqryBunjuDatabaseNamedatabaseFiltered	SQL.StringsSELECT * FROM   bunjuWHERE cod_bjjs  = :cod_bjjsAND   dat_bunju = :dat_bunju 
UpdateModeupWhereKeyOnlyLeftTopj	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown    TQuery
qryI_BunjuDatabaseNamedatabaseSQL.StringsINSERT INTO bunju=           (cod_bjjs, num_id, dat_bunju, num_bunju, dat_gmgn,9            gubn_gmsa, cod_jisa,  dat_insert, cod_insert)VALUES B           (:cod_bjjs, :num_id, :dat_bunju, :num_bunju, :dat_gmgn,=            :gubn_gmsa, :cod_jisa,  :dat_insert, :cod_insert) 
UpdateModeupWhereKeyOnlyLeftiTop� 	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown DataType	ftUnknownName	num_bunju	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown DataType	ftUnknownName	gubn_gmsa	ParamType	ptUnknown DataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName
dat_insert	ParamType	ptUnknown DataType	ftUnknownName
cod_insert	ParamType	ptUnknown    TQueryqryI_GulkwaDatabaseNamedatabaseSQL.StringsINSERT INTO gulkwa>           (cod_bjjs, num_id, dat_bunju, num_bunju, gubn_part,7            dat_gmgn, cod_jisa, dat_insert, cod_insert)VALUES C           (:cod_bjjs, :num_id, :dat_bunju, :num_bunju, :gubn_part,;            :dat_gmgn, :cod_jisa, :dat_insert, :cod_insert) 
UpdateModeupWhereKeyOnlyLeftYTopB	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown DataType	ftUnknownName	num_bunju	ParamType	ptUnknown DataType	ftUnknownName	gubn_part	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown DataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName
dat_insert	ParamType	ptUnknown DataType	ftUnknownName
cod_insert	ParamType	ptUnknown    TQuery	qryI_CellDatabaseNamedatabaseSQL.StringsINSERT INTO cellC           (cod_bjjs, num_id, cod_cell, cod_jisa, dat_gmgn, cod_hm,/            ysno_sokun, dat_insert, cod_insert)VALUES I           (:cod_bjjs, :num_id, :cod_cell, :cod_jisa, :dat_gmgn, :cod_hm,2            :ysno_sokun, :dat_insert, :cod_insert) 
UpdateModeupWhereKeyOnlyLeftYTop	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamecod_cell	ParamType	ptUnknown DataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown DataType	ftUnknownNamecod_hm	ParamType	ptUnknown DataType	ftUnknownName
ysno_sokun	ParamType	ptUnknown DataType	ftUnknownName
dat_insert	ParamType	ptUnknown DataType	ftUnknownName
cod_insert	ParamType	ptUnknown    TQuery
qryHangmokDatabaseNamedatabaseFiltered	SQL.Strings$SELECT cod_hm, gubn_part, dat_apply FROM   hangmokWHERE dat_apply <= :dat_applyORDER BY dat_apply DESC 
UpdateModeupWhereKeyOnlyLeftaTop� 	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown    TQueryqrySeqDatabaseNamedatabaseSQL.StringsSELECT MAX(cod_cell) resultFROM   cellWHERE cod_cell > :cod_cellsAND      cod_cell < :cod_celle 
UpdateModeupWhereKeyOnlyLeft�Top	ParamDataDataType	ftUnknownName	cod_cells	ParamType	ptUnknown DataType	ftUnknownName	cod_celle	ParamType	ptUnknown    TQueryqryProfileGDatabaseNamedatabaseSQL.StringsSELECT B.cod_hm +FROM   profile A INNER JOIN profile_hm B ON           A.cod_pf = B.cod_pfWHERE A.cod_pf = :cod_pf1 OR       A.cod_pf = :cod_pf2OR       A.cod_pf = :cod_pf3OR       A.cod_pf = :cod_pf4OR       A.cod_pf = :cod_pf5OR       A.cod_pf = :cod_pf6OR       A.cod_pf = :cod_pf7OR       A.cod_pf = :cod_pf8OR       A.cod_pf = :cod_pf9OR       A.cod_pf = :cod_pf10OR       A.cod_pf = :cod_pf11OR       A.cod_pf = :cod_pf12OR       A.cod_pf = :cod_pf13OR       A.cod_pf = :cod_pf14OR       A.cod_pf = :cod_pf15OR       A.cod_pf = :cod_pf16OR       A.cod_pf = :cod_pf17OR       A.cod_pf = :cod_pf18OR       A.cod_pf = :cod_pf19OR       A.cod_pf = :cod_pf20OR       A.cod_pf = :cod_pf21OR       A.cod_pf = :cod_pf22OR       A.cod_pf = :cod_pf23OR       A.cod_pf = :cod_pf24GROUP BY B.cod_hm 
UpdateModeupWhereKeyOnlyLeft�Top� 	ParamDataDataTypeftStringNamecod_pf1	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf2	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf3	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf4	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf5	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf6	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf7	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf8	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf9	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf10	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf11	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf12	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf13	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf14	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf15	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf16	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf17	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf18	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf19	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf20	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf21	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf22	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf23	ParamType	ptUnknownValue  DataTypeftStringNamecod_pf24	ParamType	ptUnknownValue     TQueryqryNo_hangmokDatabaseNamedatabaseFiltered	SQL.StringsSELECT * FROM no_hangmokWHERE  dat_apply  <= :dat_applyAND      gubn_code = :gubn_codeAND      gubn_yh    = :gubn_yhORDER BY dat_apply DESC Left�Topx	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamegubn_yh	ParamType	ptUnknown    TQueryqryJaegumhmDatabaseNamedatabaseSQL.StringsSELECT * FROM jaegumhmWHERE  cod_jisa = :cod_jisaAND      num_jumin = :num_juminAND      gubn_code = :gubn_codeAND      num_sil = :num_silORDER BY dat_gmgn desc Left�Top�	ParamDataDataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName	num_jumin	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamenum_sil	ParamType	ptUnknown    TQuery
qryBjValueDatabaseNamedatabaseSQL.Strings(SELECT MAX(num_bunju) bjValue FROM bunjuWHERE cod_bjjs   = :cod_bjjsAND   dat_bunju  = :dat_bunjuAND   num_bunju >= :snumAND   num_bunju <= :enum 
UpdateModeupWhereKeyOnlyLeftTop*	ParamDataDataTypeftStringNamecod_bjjs	ParamType	ptUnknown DataTypeftStringName	dat_bunju	ParamType	ptUnknown DataTypeftStringNamesnum	ParamType	ptUnknown DataTypeftStringNameenum	ParamType	ptUnknown    TQueryqryTkgumDatabaseNamedatabaseSQL.StringsSELECT cod_prf, cod_chugaFROM   tkgumWHERE cod_jisa = :cod_jisaAND      num_jumin = :num_juminAND      dat_gmgn = :dat_gmgn 
UpdateModeupWhereKeyOnlyLeftYTop�	ParamDataDataTypeftStringNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownName	num_jumin	ParamType	ptUnknown DataTypeftStringNamedat_gmgn	ParamType	ptUnknown    TQueryqryHconditionDatabaseNamedatabaseSQL.Stringsselect * from hul_conditionwhere cod_jisa = :cod_jisa   and num_id = :num_id   and dat_gmgn = :dat_gmgn LeftpTopP	ParamDataDataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown    TQueryqry_cellDatabaseNamedatabaseSQL.Stringsselect * from cellwhere num_id   = :num_id  and dat_gmgn = :dat_gmgn  and cod_hm   = :cod_hm Left�Topx	ParamDataDataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownNamedat_gmgn	ParamType	ptUnknown DataType	ftUnknownNamecod_hm	ParamType	ptUnknown    TQueryqryHgulkwa_chkDatabaseNamedatabaseSQL.Stringsselect * from Hgulkwa_chkwhere cod_bjjs = :cod_bjjsand dat_bunju = :dat_bunju Left� Top� 	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown    TQueryqryIHgulkwa_chkDatabaseNamedatabaseSQL.Strings,insert into Hgulkwa_chk(cod_bjjs, dat_bunju)5                        values(:cod_bjjs, :dat_bunju) Left� Top0	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown    TQueryqryDacheDatabaseNamedatabaseSQL.Stringsselect desc_dept from gumginwhere cod_dc = :cod_dc  and dat_gmgn >= :dat_sgmgn  and dat_gmgn <= :dat_egmgngroup by desc_dept Left� Top� 	ParamDataDataType	ftUnknownNamecod_dc	ParamType	ptUnknown DataType	ftUnknownName	dat_sgmgn	ParamType	ptUnknown DataType	ftUnknownName	dat_egmgn	ParamType	ptUnknown     