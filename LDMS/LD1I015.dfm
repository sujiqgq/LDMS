�
 TFRMLD1I015 0l  TPF0�TfrmLD1I015
frmLD1I015Left^Top� CaptionfrmLD1I014-��û���ϰ�����ClientHeight� ClientWidth�OldCreateOrder	PixelsPerInch`
TextHeight TGaugeGauge1LeftTop� Width�Height)	BackColor��� Colorww� KindgkNeedleParentColorProgress   TLabelLabel9Left� TopJWidthLHeightCaption
���ֹ�ȣ :Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����ü
Font.StylefsBold 
ParentFont  TLabelLabel1LeftTopJWidthLHeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����ü
Font.StylefsBold 
ParentFont  TLabelLabel2Left>TopJWidth
HeightCaption~  TLabelLabel3LeftTop/WidthLHeightCaption
�������� :Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����ü
Font.StylefsBold 
ParentFont  TPanelPanel36Left Top Width�Height 
BevelInner	bvLowered
BevelOuterbvNone
BevelWidthCaption���ױ�Ÿ1  �� �� �� ��Color��� Font.CharsetHANGEUL_CHARSET
Font.ColorclWindowTextFont.Height�	Font.Name����ü
Font.StylefsBold 
ParentFontTabOrder  TBitBtnbtnOkLeft_Top� WidthKHeightCaptionȮ��Default	TabOrderOnClickUpClick
Glyph.Data
j  f  BMf      v   (               �                       �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� wwwwwwwwww  wwwwwwwwww  wwwwwwwwww  wwwwwwwwww  wwp �wwwww  ww�"wwwww  wz"" �wwww  w�"*"wwww  w�"�� �www  w�*w�"www  w��wz� �ww  w��ww�"ww  w�wwwz� �w  wwwwww�"w  wwwwwwz� �  wwwwwwwz"  wwwwwwww�'  wwwwwwwwz�  wwwwwwwwww  wwwwwwwwww    TBitBtn	btnCancelLeftTop� WidthKHeightCancel	Caption���ModalResultTabOrderOnClickbtnCancelClick
Glyph.Data
j  f  BMf      v   (               �                       �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� wwwwwwwwww  wwwwwwwwww  wwwwwwwwww  wwwwwwwwww  wywww���w  w���w���w  w��w����w  w��w��ww  w������ww  ww��	�www  ww����www  www���wwww  www����www  ww�	��www  ww����www  w�w���ww  w���w�ww  w�www��ww  wwwwwwwwww  wwwwwwwwww    TPanelPanel1Left�Top� Width�HeightTabOrder  	TDateEditmskDateLeft]TopFWidthMHeightColor��� EditMask9999/99/99;0;_ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLength
TabOrder Year Month Day   	TMaskEditmskFromLeftTopFWidth%HeightColor��� EditMask9999;0;_ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLengthTabOrder  	TMaskEditmskToLeftMTopFWidth%HeightColor��� EditMask9999;0;_ImeModeimSAlphaImeName�ѱ���(�ѱ�)	MaxLengthTabOrder  	TCheckBoxckbS014Left	TopgWidthxHeightCaptionHAV-IgM(S014)Checked	State	cbCheckedTabOrder  	TCheckBoxckbS015Left� TopgWidthqHeightCaptionHAV-IgG(S015)Checked	State	cbCheckedTabOrder  	TComboBoxcmbJisaLeft]Top-WidthHeightColor��� ImeName�ѱ���(�ѱ�)
ItemHeightTabOrder	  TQueryqryGlkwaDatabaseNamedatabaseSQL.Strings6SELECT A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju,D           A.desc_glkwa1, A.desc_glkwa2, A.desc_glkwa3, A.gubn_part,0           B.num_jumin, B.desc_name, B.dat_gmgn,>           B.cod_hul, B.cod_cancer, B.cod_chuga, B.cod_jangbi,@           B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh,L           B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, B.Gubn_tkgm,=           B.gubn_gong, B.gubn_goyh, B.cod_jisa, B.gubn_tkgm,#           B.gubn_life, B.gubn_lfyh$FROM   gulkwa A INNER JOIN gumgin B )           ON A.cod_jisa = B.cod_jisa AND'                A.num_id = B.num_id AND'                A.dat_gmgn = B.dat_gmgn WHERE    A.cod_bjjs  = :cod_bjjs!AND      A.dat_bunju = :dat_bunjuAND      A.num_bunju >= :sbunjuAND      A.num_bunju <= :ebunjuAND      A.gubn_part = '08'  
UpdateModeupWhereKeyOnlyLeft� Top� 	ParamDataDataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown DataType	ftUnknownNamesbunju	ParamType	ptUnknown DataType	ftUnknownNameebunju	ParamType	ptUnknown    TQueryqryPf_hmDatabaseNamedatabaseSQL.StringsSELECT cod_hmFROM   profile_hmWHERE cod_pf = :cod_pf 
UpdateModeupWhereKeyOnlyLeftTop� 	ParamDataDataTypeftStringNamecod_pf	ParamType	ptUnknown    TQueryqryNo_hangmokDatabaseNamedatabaseFiltered	SQL.StringsSELECT * FROM no_hangmokWHERE  dat_apply  <= :dat_applyAND      gubn_code = :gubn_codeAND      gubn_yh    = :gubn_yhORDER BY dat_apply DESC LeftVTop� 	ParamDataDataType	ftUnknownName	dat_apply	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamegubn_yh	ParamType	ptUnknown    TQueryqryJaegumhmDatabaseNamedatabaseSQL.StringsSELECT * FROM jaegumhmWHERE  cod_jisa = :cod_jisaAND      num_id = :num_idAND      gubn_code = :gubn_codeAND      num_sil = :num_silORDER BY dat_apply Left� Top� 	ParamDataDataType	ftUnknownNamecod_jisa	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownName	gubn_code	ParamType	ptUnknown DataType	ftUnknownNamenum_sil	ParamType	ptUnknown    TQuery
qryU_GlkwaDatabaseNamedatabaseSQL.StringsUPDATE gulkwa&SET        desc_glkwa1 = :desc_glkwa1,(             desc_glkwa2 = :desc_glkwa2,'             desc_glkwa3 = :desc_glkwa3WHERE   cod_bjjs = :cod_bjjsAND        num_id = :num_id!AND        dat_bunju = :dat_bunju!AND        num_bunju = :num_bunju!AND        gubn_part = :gubn_part LeftNTop� 	ParamDataDataType	ftUnknownNamedesc_glkwa1	ParamType	ptUnknown DataType	ftUnknownNamedesc_glkwa2	ParamType	ptUnknown DataType	ftUnknownNamedesc_glkwa3	ParamType	ptUnknown DataType	ftUnknownNamecod_bjjs	ParamType	ptUnknown DataType	ftUnknownNamenum_id	ParamType	ptUnknown DataType	ftUnknownName	dat_bunju	ParamType	ptUnknown DataType	ftUnknownName	num_bunju	ParamType	ptUnknown DataType	ftUnknownName	gubn_part	ParamType	ptUnknown    TQueryqryTkgumDatabaseNamedatabaseSQL.Strings$Select Cod_prf, Cod_chuga from tkgumWhere Num_Jumin = :In_Num_Jumin  And Dat_Gmgn = :In_Dat_gmgn  And Gubn_Cha = :In_Gubn_Cha Left�Top� 	ParamDataDataType	ftUnknownNameIn_Num_Jumin	ParamType	ptUnknown DataType	ftUnknownNameIn_Dat_gmgn	ParamType	ptUnknown DataType	ftUnknownNameIn_Gubn_Cha	ParamType	ptUnknown     