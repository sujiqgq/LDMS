//==============================================================================
// ���α׷� ���� : �����������
// �ý���        :
// ��������      : 2012.12.07
// ������        : ���ö
// ��������      :
// ��������      :
//==============================================================================
// ���α׷� ���� : �����������
// �ý���        :
// ��������      : 2016.02.03
// ������        : �ڴ��
// ��������      : 1. �������� Ȱ�뵿�� �ֱ����� ����
//                 2. ������� �߰� - 5.�ȳ����ڹ߼�(DM���ڹ߼�)
//                 3. ����â�� ���÷��� ��ư �ش� ������������ ���ٰ���
//                 4. VIP������ Ȯ�� ���� ����
// ��������      : ���� ����CRM�� 1500087 - ����� : ������
//==============================================================================
// ��������      : 2016.08.06
// ������        : ������
// ��������      : CTI�ű԰���(��Ʈã�Ƽ�����:�ű�CTI�۾�...)
//==============================================================================
// ��������      : 2017.05.01
// ������        : ������
// ��������      : [Ư��]�����ڷ� �߰�...
//==============================================================================
// ��������      : 2017.05.11
// ������        : ������
// ��������      : �������� ǥ��(��������->��������)
//==============================================================================
unit CK5I17;


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, ToolWin, ComCtrls, ImgList, ExtCtrls, Stdctrls, mask, Buttons,
  Menus, DateEdit, Grids_ts, TSGrid, Grids, DBGrids, TeEngine, Series,
  Db, DBTables, OleCtrls, ShellAPI, FileCtrl, ORMaster,
  Gauges, ComObj, CAMODULELib_TLB;

type
  TfrmCK5I17 = class(TfrmMaster)
    ToolButton1: TToolButton;
    btnInsert: TSpeedButton;
    btnDelete: TSpeedButton;
    ToolButton3: TToolButton;
    btnSave: TSpeedButton;
    btnCancel: TSpeedButton;
    ToolButton4: TToolButton;
    btnFind: TSpeedButton;
    btnQuit: TSpeedButton;
    ToolButton5: TToolButton;
    btnPrint: TSpeedButton;
    ToolButton6: TToolButton;
    btnQuery: TSpeedButton;
    ToolButton2: TToolButton;
    btnExcel: TSpeedButton;
    Panel1: TPanel;
    Bevel1: TBevel;
    btnOpenOffice: TSpeedButton;
    Panel2: TPanel;
    Panel3: TPanel;
    edtName: TEdit;
    edtJumin: TMaskEdit;
    Panel5: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    consult_date: TEdit;
    consult_time: TEdit;
    Panel12: TPanel;
    consult_note: TMemo;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    con_type: TComboBox;
    con_kind: TComboBox;
    next_gn_date: TDateEdit;
    btnNext_gn_date: TSpeedButton;
    GroupBox1: TGroupBox;
    Ck_hospital: TCheckBox;
    h_code: TEdit;
    h_name: TEdit;
    qryGumgin: TQuery;
    qryJc_cohospital: TQuery;
    qryJc_consults_list: TQuery;
    qryJc_consults: TQuery;
    btn_New: TSpeedButton;
    btn_save: TSpeedButton;
    btn_delete: TSpeedButton;
    qryI_Jc_consults: TQuery;
    qryKicho: TQuery;
    qryD_Jc_consults: TQuery;
    Panel17: TPanel;
    Cmb_sokun_name: TComboBox;
    qryjc_cohospital_sokun_code: TQuery;   
    qryU_Jc_consults: TQuery;
    qryhospital_Sokun_code: TQuery;
    qryI_jc_cohospital_sokun: TQuery;
    qryD_jc_cohospital_sokun: TQuery;
    btn_search: TSpeedButton;
    qryJc_consults_idx: TQuery;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    grdMaster_Detail: TtsGrid;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    Edt_Company: TEdit;
    Panel4: TPanel;
    Panel6: TPanel;
    Edt_Dept: TEdit;
    Panel7: TPanel;
    Edt_Tel: TEdit;
    Edt_Address: TEdit;
    Panel8: TPanel;
    grd_Kicho: TtsGrid;
    GroupBox: TGroupBox;
    SBtn_CTI_Login: TSpeedButton;
    SBtn_CTI_Monitor: TSpeedButton;
    SBtn_CTI_Callback: TSpeedButton;
    SBtn_CTI_Setting: TSpeedButton;
    SBtn_CTI_Answer: TSpeedButton;
    SBtn_CTI_Break: TSpeedButton;
    SBtn_CTI_Disconnect: TSpeedButton;
    SBtn_CTI_Hu: TSpeedButton;
    SBtn_CTI_Etc: TSpeedButton;
    grd_Hul: TtsGrid;
    grd_Jangbi: TtsGrid;
    grd_sokun: TtsGrid;
    qryHul_gumgin: TQuery;
    qryInjouk: TQuery;
    qryTot_sokun: TQuery;
    qryJangbi: TQuery;
    qrySawon_CTI: TQuery;
    qryHul: TQuery;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    labRelation: TLabel;
    cmbRelation: TComboBox;
    Panel9: TPanel;
    grd_file: TtsGrid;
    qryFile_directory: TQuery;
    qryHm: TQuery;
    qryGlkwa: TQuery;
    qryHm_name: TQuery;
    qryGmgn: TQuery;
    TabSheet6: TTabSheet;
    grdMaster: TtsGrid;
    Panel16: TPanel;
    cmbCOD_JISA: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Label15: TLabel;
    qryI_SMS: TQuery;
    qryI_MMS: TQuery;
    qryI_MMS_MSG: TQuery;
    qryI_SC_TRAN: TQuery;
    meMess: TMemo;
    Label17: TLabel;
    Label18: TLabel;
    btnYeyak: TSpeedButton;
    Label19: TLabel;
    Label20: TLabel;
    callback_tel: TEdit;
    Num_tel: TEdit;
    date_yeyak: TDateEdit;
    Chk_yeyak: TCheckBox;
    Cmb_HH: TComboBox;
    Cmb_MM: TComboBox;
    qryJc_sms_message: TQuery;
    Btn_SMS: TButton;
    qryGmgn2: TQuery;
    SBtn_Call: TSpeedButton;
    pnlYsno_useinfo: TPanel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryTkgum: TQuery;
    qry_Get_CTI_ID: TQuery;
    Timer_hu: TTimer;
    Timer1: TTimer;
    Chk_ReCall: TCheckBox;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    TabSheet9: TTabSheet;
    Label77: TLabel;
    Label78: TLabel;
    Label79: TLabel;
    Label80: TLabel;
    Label81: TLabel;
    Label82: TLabel;
    Label83: TLabel;
    Label84: TLabel;
    Label85: TLabel;
    Label86: TLabel;
    Label87: TLabel;
    Label88: TLabel;
    Label89: TLabel;
    Label90: TLabel;
    Label91: TLabel;
    Label92: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    Label95: TLabel;
    Label96: TLabel;
    Label97: TLabel;
    Label98: TLabel;
    Label99: TLabel;
    Label100: TLabel;
    Label101: TLabel;
    Label102: TLabel;
    Label103: TLabel;
    Label104: TLabel;
    Label105: TLabel;
    Label106: TLabel;
    Label107: TLabel;
    Label108: TLabel;
    Label109: TLabel;
    Label110: TLabel;
    Label111: TLabel;
    Label112: TLabel;
    Label113: TLabel;
    Label114: TLabel;
    Label115: TLabel;
    Label116: TLabel;
    Label117: TLabel;
    Label118: TLabel;
    Label119: TLabel;
    Label120: TLabel;
    Label121: TLabel;
    Label122: TLabel;
    Bevel4: TBevel;
    Ck_Mun1_J1: TCheckBox;
    Ck_Mun1_J2: TCheckBox;
    Ck_Mun1_J3: TCheckBox;
    Ck_Mun1_J4: TCheckBox;
    Ck_Mun1_J5: TCheckBox;
    Ck_Mun1_J8: TCheckBox;
    Ck_Mun1_D1: TCheckBox;
    Ck_Mun1_D2: TCheckBox;
    Ck_Mun1_D3: TCheckBox;
    Ck_Mun1_D4: TCheckBox;
    Ck_Mun1_D5: TCheckBox;
    Ck_Mun1_D8: TCheckBox;
    Mun1_J_YN: TComboBox;
    Mun1_D_YN: TComboBox;
    Ck_Mun1_J7: TCheckBox;
    Ck_Mun1_D7: TCheckBox;
    Ck_Mun1_J6: TCheckBox;
    Ck_Mun1_D6: TCheckBox;
    Ck_Mun2_1: TCheckBox;
    Ck_Mun2_2: TCheckBox;
    Ck_Mun2_3: TCheckBox;
    Ck_Mun2_4: TCheckBox;
    Ck_Mun2_5: TCheckBox;
    Mun2_YN: TComboBox;
    Ck_Mun2_6: TCheckBox;
    Mun3: TComboBox;
    Mun4_1: TComboBox;
    Mun4_2_Year: TEdit;
    Mun4_2_Day: TEdit;
    Mun4_3_Year: TEdit;
    Mun4_3_Day: TEdit;
    Mun5_2: TComboBox;
    Mun5_3_2: TEdit;
    Mun5_1: TComboBox;
    Mun5_4: TEdit;
    Mun5_3_1: TComboBox;
    Mun6_2: TComboBox;
    Mun6_3: TComboBox;
    Mun6_4: TComboBox;
    Mun6_1: TComboBox;
    Mun8_1: TComboBox;
    Mun8_2: TComboBox;
    Mun8_3: TComboBox;
    Mun8_4: TComboBox;
    CMun1: TComboBox;
    StaticText40: TStaticText;
    CMun1_Bung: TEdit;
    StaticText26: TStaticText;
    CMun2: TComboBox;
    CMun2_kg: TEdit;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText47: TStaticText;
    StaticText48: TStaticText;
    CMun3_Can6_Bung: TEdit;
    StaticText37: TStaticText;
    CMun3_Can1_Yn: TComboBox;
    CMun3_Can2_Yn: TComboBox;
    CMun3_Can3_Yn: TComboBox;
    CMun3_Can6_Yn: TComboBox;
    CMun3_Can4_Yn: TComboBox;
    CMun3_Can5_Yn: TComboBox;
    CMun3_Can1_F1: TCheckBox;
    CMun3_Can1_F2: TCheckBox;
    CMun3_Can1_F3: TCheckBox;
    CMun3_Can1_F4: TCheckBox;
    CMun3_Can1_F5: TCheckBox;
    StaticText38: TStaticText;
    StaticText39: TStaticText;
    StaticText46: TStaticText;
    StaticText49: TStaticText;
    CMun3_Can2_F1: TCheckBox;
    CMun3_Can2_F2: TCheckBox;
    CMun3_Can2_F3: TCheckBox;
    CMun3_Can2_F4: TCheckBox;
    CMun3_Can2_F5: TCheckBox;
    CMun3_Can3_F1: TCheckBox;
    CMun3_Can3_F2: TCheckBox;
    CMun3_Can3_F3: TCheckBox;
    CMun3_Can3_F4: TCheckBox;
    CMun3_Can3_F5: TCheckBox;
    CMun3_Can4_F1: TCheckBox;
    CMun3_Can4_F2: TCheckBox;
    CMun3_Can4_F3: TCheckBox;
    CMun3_Can4_F4: TCheckBox;
    CMun3_Can4_F5: TCheckBox;
    CMun3_Can5_F1: TCheckBox;
    CMun3_Can5_F2: TCheckBox;
    CMun3_Can5_F3: TCheckBox;
    CMun3_Can5_F4: TCheckBox;
    CMun3_Can5_F5: TCheckBox;
    CMun3_Can6_F1: TCheckBox;
    CMun3_Can6_F2: TCheckBox;
    CMun3_Can6_F3: TCheckBox;
    CMun3_Can6_F4: TCheckBox;
    CMun3_Can6_F5: TCheckBox;
    StaticText41: TStaticText;
    CMun4_Can1: TComboBox;
    CMun4_Can2: TComboBox;
    CMun4_Can3: TComboBox;
    CMun4_Can4: TComboBox;
    CMun4_Can5: TComboBox;
    CMun4_Can6: TComboBox;
    CMun4_Can7: TComboBox;
    CMun4_Can8: TComboBox;
    StaticText44: TStaticText;
    CMun5_1: TCheckBox;
    CMun5_2: TCheckBox;
    CMun5_3: TCheckBox;
    CMun5_4: TCheckBox;
    CMun5_5: TCheckBox;
    CMun5_YN: TComboBox;
    StaticText45: TStaticText;
    CMun6_1: TCheckBox;
    CMun6_2: TCheckBox;
    CMun6_3: TCheckBox;
    CMun6_4: TCheckBox;
    CMun6_5: TCheckBox;
    StaticText50: TStaticText;
    CMun7_1: TCheckBox;
    CMun7_2: TCheckBox;
    CMun7_4: TCheckBox;
    CMun7_3: TCheckBox;
    CMun7_5: TCheckBox;
    CMun6_YN: TComboBox;
    CMun7_YN: TComboBox;
    Woman1: TPanel;
    Label123: TLabel;
    Label124: TLabel;
    Label125: TLabel;
    Label126: TLabel;
    Label127: TLabel;
    StaticText34: TStaticText;
    CMun8: TComboBox;
    CMun8_Year: TEdit;
    StaticText55: TStaticText;
    CMun9: TComboBox;
    CMun9_Year: TEdit;
    StaticText52: TStaticText;
    CMun10: TComboBox;
    StaticText53: TStaticText;
    CMun11: TComboBox;
    StaticText58: TStaticText;
    CMun12: TComboBox;
    StaticText59: TStaticText;
    CMun13: TComboBox;
    StaticText33: TStaticText;
    CMun14: TComboBox;
    Panel35: TPanel;
    Label128: TLabel;
    Label129: TLabel;
    Label266: TLabel;
    Label130: TLabel;
    Label131: TLabel;
    Label132: TLabel;
    Life_Haji_time: TEdit;
    Life_blance_state: TComboBox;
    Life_work: TComboBox;
    Life_blance_time: TEdit;
    StaticText19: TStaticText;
    StaticText90: TStaticText;
    cmbYsno_agree: TComboBox;
    GroupBox13: TGroupBox;
    edtSang_Form1: TCheckBox;
    edtSang_Form2: TCheckBox;
    edtSang_Form4: TCheckBox;
    edtSang_Form3: TCheckBox;
    CmbSang_Form1_gubn: TComboBox;
    retDesc_Sokun1: TMemo;
    retDesc_Sokun: TMemo;
    GroupBox10: TGroupBox;
    edtGum_Form1: TCheckBox;
    edtGum_Form2: TCheckBox;
    edtGum_Form3: TCheckBox;
    edtGum_Form3_Etc: TEdit;
    GroupBox11: TGroupBox;
    Label74: TLabel;
    Label75: TLabel;
    edtDesc_Gum1: TCheckBox;
    edtDesc_Gum2: TCheckBox;
    edtGum_Remark: TEdit;
    edtGum_Chk1: TCheckBox;
    edtGum_Chk2: TCheckBox;
    GroupBox9: TGroupBox;
    Label73: TLabel;
    ckYSNO_VIRUS1: TCheckBox;
    ckYSNO_VIRUS2: TCheckBox;
    ckYSNO_VIRUS3: TCheckBox;
    ckYSNO_VIRUS4: TCheckBox;
    ckYSNO_VIRUS5: TCheckBox;
    edtDESC_VIRUS: TEdit;
    ckYsno_Virus7: TCheckBox;
    ckYsno_Virus6: TCheckBox;
    ckYsno_Virus8: TCheckBox;
    Panel38: TPanel;
    cmbCell_Hangmok: TComboBox;
    Panel39: TPanel;
    retDESC_KGBL: TRichEdit;
    Panel22: TPanel;
    edtDESC_BUWI: TEdit;
    Panel37: TPanel;
    Panel21: TPanel;
    cmbGUBN_CLASS: TComboBox;
    cmbCOD_DOCT: TComboBox;
    Panel36: TPanel;
    mskDAT_RESULT: TDateEdit;
    Panel40: TPanel;
    Panel60: TPanel;
    edtCod_Sokun: TEdit;
    edtGUBN_PAN: TEdit;
    GroupBox12: TGroupBox;
    Label76: TLabel;
    edtSun_Form1: TCheckBox;
    edtSun_Form2: TCheckBox;
    edtSun_Form4: TEdit;
    edtSun_Form3: TCheckBox;
    Panel18: TPanel;
    Panel19: TPanel;
    Label134: TLabel;
    lblYYtot9: TLabel;
    lblYYtot5: TLabel;
    Label135: TLabel;
    lblYYtot8: TLabel;
    lblYYtot6: TLabel;
    lblYYtot7: TLabel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    lblYYtot10: TLabel;
    Bevel7: TBevel;
    Label136: TLabel;
    Label137: TLabel;
    Bevel8: TBevel;
    Label138: TLabel;
    Label139: TLabel;
    Panel44: TPanel;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    StaticText6: TStaticText;
    StaticText7: TStaticText;
    StaticText8: TStaticText;
    StaticText9: TStaticText;
    StaticText10: TStaticText;
    StaticText11: TStaticText;
    StaticText12: TStaticText;
    StaticText13: TStaticText;
    StaticText14: TStaticText;
    StaticText15: TStaticText;
    HPan1: TStaticText;
    HPan2: TStaticText;
    HPan3: TStaticText;
    HPan4: TStaticText;
    HPan5: TStaticText;
    HPan6: TStaticText;
    HPan7: TStaticText;
    HPan8: TStaticText;
    HPan9: TStaticText;
    HPan10: TStaticText;
    HPan11: TStaticText;
    HPan12: TStaticText;
    HPan13: TStaticText;
    Panel45: TPanel;
    StaticText31: TStaticText;
    StaticText32: TStaticText;
    StaticText35: TStaticText;
    StaticText36: TStaticText;
    StaticText42: TStaticText;
    StaticText43: TStaticText;
    StaticText51: TStaticText;
    StaticText54: TStaticText;
    StaticText56: TStaticText;
    StaticText57: TStaticText;
    StaticText60: TStaticText;
    StaticText61: TStaticText;
    StaticText62: TStaticText;
    JPan1: TStaticText;
    JPan2: TStaticText;
    JPan3: TStaticText;
    JPan4: TStaticText;
    JPan5: TStaticText;
    JPan6: TStaticText;
    JPan7: TStaticText;
    JPan8: TStaticText;
    JPan9: TStaticText;
    JPan10: TStaticText;
    JPan11: TStaticText;
    JPan12: TStaticText;
    JPan13: TStaticText;
    StaticText76: TStaticText;
    JPan14: TStaticText;
    StaticText78: TStaticText;
    JPan15: TStaticText;
    StaticText80: TStaticText;
    JPan16: TStaticText;
    StaticText82: TStaticText;
    StaticText83: TStaticText;
    StaticText84: TStaticText;
    StaticText85: TStaticText;
    qryCell: TQuery;
    qryCell_sokun: TQuery;
    ComboBox1: TComboBox;
    Panel20: TPanel;
    cmb_gmgn: TComboBox;
    Panel_PAP: TPanel;
    QryTot_Munjin2010: TQuery;
    qryKicho2: TQuery;
    qryHsok: TQuery;
    qryHul_gumgin2: TQuery;
    qryHm_sp: TQuery;
    qryGlkwa2: TQuery;
    qryJangbi2: TQuery;
    Btn_Complain: TBitBtn;
    Ysno_Label: TPanel;
    qryGmgn3: TQuery;
    qryKicho3: TQuery;
    Panel54: TPanel;
    pnl_WaitCnt: TPanel;
    Panel_CTI: TPanel;
    Pnl_CTISvr: TPanel;
    Pnl_Mod: TPanel;
    Pnl_CTIUser: TPanel;
    Panel55: TPanel;
    Panel56: TPanel;
    Panel_CTINo: TPanel;
    Pnl_PopupType: TPanel;
    SBtn_Call2: TSpeedButton;
    Panel_CTI_Connec: TPanel;
    Panel41: TPanel;
    Panel57: TPanel;
    Panel59: TPanel;
    Cmb_CTI_status: TComboBox;
    Cmb_CTI_status_all: TComboBox;
    Panel_CTI_Custom: TEdit;
    Pnl_tel: TEdit;
    QryTemp: TQuery;
    TabSheet10: TTabSheet;
    Panel46: TPanel;
    pnlTkgum: TPanel;
    Panel50: TPanel;
    Panel23: TPanel;
    EdtTNum_seq: TEdit;
    Panel24: TPanel;
    EdtTCod_HMatter: TEdit;
    EdtDesc_HMatter: TEdit;
    Panel26: TPanel;
    edtTGubn_LR: TEdit;
    Panel27: TPanel;
    EdtTCod_pan: TEdit;
    Panel28: TPanel;
    EdtTcod_sokun: TEdit;
    EdtTDesc_sokun: TEdit;
    Panel29: TPanel;
    EdtTCod_jochi: TEdit;
    EdtTDesc_jochi: TEdit;
    Panel30: TPanel;
    EdtTTot_Jilhan: TEdit;
    Panel89: TPanel;
    Panel31: TPanel;
    EdtTCod_jilhan: TEdit;
    EdtTCod_sahu1: TEdit;
    Panel32: TPanel;
    EdtTDesc_jilhan: TEdit;
    EdtTDesc_sahu1: TEdit;
    Panel33: TPanel;
    EdtTCod_ujs: TEdit;
    EdtTDesc_ujs: TEdit;
    EdtTDesc_sahu2: TEdit;
    EdtTCod_sahu2: TEdit;
    Panel34: TPanel;
    Panel42: TPanel;
    MM_Tetcsokun: TMemo;
    Panel58: TPanel;
    MM_Tetcbigo: TMemo;
    Panel74: TPanel;
    MM_TKMI_bigo: TMemo;
    Panel43: TPanel;
    MM_TetcJochi: TMemo;
    grdTkSokun: TtsGrid;
    Panel48: TPanel;
    Panel49: TPanel;
    Panel51: TPanel;
    EdtTCod_janggi: TEdit;
    EdtTDesc_janggi: TEdit;
    EdtCod_Doctor: TEdit;
    EdtDesc_Doctor: TEdit;
    grdResult: TtsGrid;
    grdResult2: TtsGrid;
    Panel25: TPanel;
    EdtGubn_Cha: TEdit;
    qryTkgumCheck: TQuery;
    qryTk_sokun2009: TQuery;
    qryTK_Jangbi: TQuery;
    qryTk_Comm12: TQuery;
    qryTK_Hul: TQuery;
    qryTk_gulkwa: TQuery;
    qryTKicho: TQuery;
    Pnl_Complete: TPanel;
    Mm_munjin2: TMemo;
    StaticText16: TStaticText;
    TabSheet11: TTabSheet;
    qry_gmgn: TQuery;
    grdGmgn: TtsGrid;
    qrySokun: TQuery;
    qryComm06: TQuery;
    qryNs_sokun: TQuery;
    qrySawon: TQuery;
    btn_Cancel: TSpeedButton;
    Panel47: TPanel;
    Panel52: TPanel;
    TabSheet12: TTabSheet;
    Label175: TLabel;
    Label176: TLabel;
    Label178: TLabel;
    Label179: TLabel;
    Label180: TLabel;
    Label181: TLabel;
    Label182: TLabel;
    Label183: TLabel;
    Label184: TLabel;
    Label185: TLabel;
    Label186: TLabel;
    Label187: TLabel;
    Label188: TLabel;
    Label189: TLabel;
    Label191: TLabel;
    Label192: TLabel;
    Label193: TLabel;
    Label194: TLabel;
    Label195: TLabel;
    Label177: TLabel;
    Ck_Mun1_J1_2018: TCheckBox;
    Ck_Mun1_J2_2018: TCheckBox;
    Ck_Mun1_J3_2018: TCheckBox;
    Ck_Mun1_J4_2018: TCheckBox;
    Ck_Mun1_J5_2018: TCheckBox;
    Ck_Mun1_J7_2018: TCheckBox;
    Ck_Mun1_D1_2018: TCheckBox;
    Ck_Mun1_D2_2018: TCheckBox;
    Ck_Mun1_D3_2018: TCheckBox;
    Ck_Mun1_D4_2018: TCheckBox;
    Ck_Mun1_D5_2018: TCheckBox;
    Ck_Mun1_D7_2018: TCheckBox;
    Ck_Mun2_1_2018: TCheckBox;
    Ck_Mun2_2_2018: TCheckBox;
    Ck_Mun2_3_2018: TCheckBox;
    Ck_Mun2_4_2018: TCheckBox;
    Mun3_2018: TComboBox;
    Mun4_2018: TComboBox;
    Mun4_1_Year_2018: TEdit;
    Mun4_1_Day_2018: TEdit;
    Mun4_2_Year_2018: TEdit;
    Mun4_2_Day_2018: TEdit;
    Mun6_num_2018: TComboBox;
    Mun2_YN_2018: TComboBox;
    Ck_Mun1_J6_2018: TCheckBox;
    Ck_Mun1_D6_2018: TCheckBox;
    Ck_Mun2_5_2018: TCheckBox;
    Mun5_2018: TComboBox;
    Mun5_1_2018: TComboBox;
    Mun6_count_2018: TEdit;
    Mun1_J_YN_2018: TComboBox;
    Mun1_D_YN_2018: TComboBox;
    Mun6_1_2018: TEdit;
    Mun6_2_2018: TEdit;
    Label190: TLabel;
    Label196: TLabel;
    Mun7_2_H_2018: TEdit;
    Mun7_2_M_2018: TEdit;
    Mun7_1_2018: TComboBox;
    Label197: TLabel;
    Label198: TLabel;
    Mun8_2_M_2018: TEdit;
    Mun8_2_H_2018: TEdit;
    Mun8_1_2018: TComboBox;
    Label199: TLabel;
    Mun9_2018: TComboBox;
    Bevel9: TBevel;
    Label200: TLabel;
    Label201: TLabel;
    Label202: TLabel;
    Label203: TLabel;
    Label204: TLabel;
    Label205: TLabel;
    Label206: TLabel;
    Label207: TLabel;
    Label208: TLabel;
    Label209: TLabel;
    Label210: TLabel;
    Label211: TLabel;
    Label212: TLabel;
    Label213: TLabel;
    Label214: TLabel;
    CMun1_2018: TComboBox;
    StaticText17: TStaticText;
    CMun1_Bung_2018: TEdit;
    StaticText18: TStaticText;
    CMun2_2018: TComboBox;
    CMun2_kg_2018: TEdit;
    StaticText20: TStaticText;
    StaticText21: TStaticText;
    StaticText22: TStaticText;
    StaticText23: TStaticText;
    CMun3_Can6_Bung_2018: TEdit;
    StaticText24: TStaticText;
    CMun3_Can1_Yn_2018: TComboBox;
    CMun3_Can2_Yn_2018: TComboBox;
    CMun3_Can3_Yn_2018: TComboBox;
    CMun3_Can6_Yn_2018: TComboBox;
    CMun3_Can4_Yn_2018: TComboBox;
    CMun3_Can5_Yn_2018: TComboBox;
    CMun3_Can1_F1_2018: TCheckBox;
    CMun3_Can1_F2_2018: TCheckBox;
    CMun3_Can1_F3_2018: TCheckBox;
    CMun3_Can1_F4_2018: TCheckBox;
    CMun3_Can1_F5_2018: TCheckBox;
    StaticText25: TStaticText;
    StaticText27: TStaticText;
    StaticText28: TStaticText;
    StaticText29: TStaticText;
    CMun3_Can2_F1_2018: TCheckBox;
    CMun3_Can2_F2_2018: TCheckBox;
    CMun3_Can2_F3_2018: TCheckBox;
    CMun3_Can2_F4_2018: TCheckBox;
    CMun3_Can2_F5_2018: TCheckBox;
    CMun3_Can3_F1_2018: TCheckBox;
    CMun3_Can3_F2_2018: TCheckBox;
    CMun3_Can3_F3_2018: TCheckBox;
    CMun3_Can3_F4_2018: TCheckBox;
    CMun3_Can3_F5_2018: TCheckBox;
    CMun3_Can4_F1_2018: TCheckBox;
    CMun3_Can4_F2_2018: TCheckBox;
    CMun3_Can4_F3_2018: TCheckBox;
    CMun3_Can4_F4_2018: TCheckBox;
    CMun3_Can4_F5_2018: TCheckBox;
    CMun3_Can5_F1_2018: TCheckBox;
    CMun3_Can5_F2_2018: TCheckBox;
    CMun3_Can5_F3_2018: TCheckBox;
    CMun3_Can5_F4_2018: TCheckBox;
    CMun3_Can5_F5_2018: TCheckBox;
    CMun3_Can6_F1_2018: TCheckBox;
    CMun3_Can6_F2_2018: TCheckBox;
    CMun3_Can6_F3_2018: TCheckBox;
    CMun3_Can6_F4_2018: TCheckBox;
    CMun3_Can6_F5_2018: TCheckBox;
    StaticText30: TStaticText;
    CMun4_Can1_2018: TComboBox;
    CMun4_Can2_2018: TComboBox;
    CMun4_Can3_2018: TComboBox;
    CMun4_Can4_2018: TComboBox;
    CMun4_Can5_2018: TComboBox;
    CMun4_Can6_2018: TComboBox;
    CMun4_Can7_2018: TComboBox;
    CMun4_Can8_2018: TComboBox;
    StaticText63: TStaticText;
    CMun5_1_2018: TCheckBox;
    CMun5_2_2018: TCheckBox;
    CMun5_3_2018: TCheckBox;
    CMun5_4_2018: TCheckBox;
    CMun5_5_2018: TCheckBox;
    CMun5_YN_2018: TComboBox;
    StaticText64: TStaticText;
    CMun6_1_2018: TCheckBox;
    CMun6_2_2018: TCheckBox;
    CMun6_3_2018: TCheckBox;
    CMun6_4_2018: TCheckBox;
    CMun6_5_2018: TCheckBox;
    StaticText65: TStaticText;
    CMun7_1_2018: TCheckBox;
    CMun7_2_2018: TCheckBox;
    CMun7_4_2018: TCheckBox;
    CMun7_3_2018: TCheckBox;
    CMun7_5_2018: TCheckBox;
    CMun6_YN_2018: TComboBox;
    CMun7_YN_2018: TComboBox;
    Woman1_2018: TPanel;
    Label215: TLabel;
    Label216: TLabel;
    Label217: TLabel;
    Label218: TLabel;
    Label219: TLabel;
    StaticText66: TStaticText;
    CMun8_2018: TComboBox;
    CMun8_Year_2018: TEdit;
    StaticText67: TStaticText;
    CMun9_2018: TComboBox;
    CMun9_Year_2018: TEdit;
    StaticText68: TStaticText;
    CMun10_2018: TComboBox;
    StaticText69: TStaticText;
    CMun11_2018: TComboBox;
    StaticText70: TStaticText;
    CMun12_2018: TComboBox;
    StaticText71: TStaticText;
    CMun13_2018: TComboBox;
    StaticText72: TStaticText;
    CMun14_2018: TComboBox;
    StaticText73: TStaticText;
    cmbYsno_agree_2018: TComboBox;
    Mm_munjin2_2018: TMemo;
    StaticText74: TStaticText;
    QryTot_Munjin2018: TQuery;
    PnlYsno_Privacy_Forbid: TPanel;
    Label1: TLabel;
    Label4: TLabel;
    Button1: TButton;
    Button2: TButton;
    Label5: TLabel;
    qryD_File_directory: TQuery;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Pnl_Mun_2019: TPanel;
    Label247: TLabel;
    Label280: TLabel;
    Label281: TLabel;
    Label282: TLabel;
    Label283: TLabel;
    Label284: TLabel;
    Label285: TLabel;
    Label354: TLabel;
    Label286: TLabel;
    Label287: TLabel;
    Label359: TLabel;
    Label360: TLabel;
    Label361: TLabel;
    Label362: TLabel;
    Label363: TLabel;
    Label364: TLabel;
    Label248: TLabel;
    Label250: TLabel;
    Mun4_2019: TComboBox;
    Mun4_1_Day_2019: TEdit;
    Mun4_1_PYear_2019: TEdit;
    Mun4_2_2019: TComboBox;
    Mun4_2_1_2019: TComboBox;
    Mun4_2_Year_2019: TEdit;
    Mun4_2_Day_2019: TEdit;
    Mun4_2_PYear_2019: TEdit;
    Mun5_2019: TComboBox;
    Mun5_1_2019: TComboBox;
    Mun4_1_2019: TComboBox;
    Mun4_1_Year_2019: TEdit;



    procedure btnQuitClick(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure btnOpenOfficeClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_VarMemSet(iLength : Integer; sGubun : String);
    procedure UP_DisplayDetail;
    procedure grdMaster_DetailCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure grdMaster_DetailRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btn_NewClick(Sender: TObject);
    procedure btn_saveClick(Sender: TObject);
    procedure btn_deleteClick(Sender: TObject);
    procedure Ck_hospitalClick(Sender: TObject);
    procedure btn_searchClick(Sender: TObject);
    procedure edtJuminChange(Sender: TObject);
    procedure edtNameChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PageControl1Change(Sender: TObject);
    procedure CreateTextArray;
    procedure CreateTextArray_Hul;
    procedure CreateTextArray_Jangbi;
    procedure CreateTextArray_Sokun;
    procedure CreateTextArray_Cell;
    procedure CreateTextArray_Pan;
    procedure grd_KichoCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grd_HulCellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);
    procedure grd_sokunCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grd_sokunRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grd_JangbiCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grd_JangbiRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure SBtn_CTI_LoginClick(Sender: TObject);
    procedure SBtn_CTI_AnswerClick(Sender: TObject);
    procedure SBtn_CTI_DisconnectClick(Sender: TObject);
    procedure SBtn_CTI_BreakClick(Sender: TObject);
    procedure SBtn_CTI_HuClick(Sender: TObject);
    procedure SBtn_CTI_EtcClick(Sender: TObject);
    procedure SBtn_CTI_SettingClick(Sender: TObject);
    procedure grd_HulRowChanged(Sender: TObject; OldRow, NewRow: Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure cmbRelationChange(Sender: TObject);
    procedure grd_fileCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grd_fileButtonClick(Sender: TObject; DataCol,
      DataRow: Integer);
    procedure btnQueryClick(Sender: TObject);
    procedure grd_KichoRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grd_KichoCellChanged(Sender: TObject; OldCol, NewCol, OldRow,
      NewRow: Integer);
    procedure cmbCOD_JISAChange(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdMasterClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure Btn_SMSClick(Sender: TObject);
    procedure btnYeyakClick(Sender: TObject);
    procedure meMessChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Chk_Edit(Sender: TObject);
    procedure SBtn_CallClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure Timer_huTimer(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure UP_PAPQuery(sCod_hm : String);
    procedure UP_Change(Sender: TObject);
    procedure CreateTextArrayTot_Munjin;  
    procedure CreateTextArrayTot_Munjin2018;
    procedure UP_ResultQuery;
    procedure UP_Separate(sData : String; var sValue : String);
    procedure Btn_ComplainClick(Sender: TObject);
    procedure CAModuleCTIConnect(Sender: TObject);
    procedure CAModuleCTIConnectError(Sender: TObject);
    procedure CAModuleCTIEvent(Sender: TObject; const strMsg: WideString);
    procedure SBtn_Call2Click(Sender: TObject);
    procedure SBtn_CTI_CallbackClick(Sender: TObject);
    procedure SBtn_CTI_MonitorClick(Sender: TObject);
    procedure grdTkSokunCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdTkSokunRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdResultCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdResult2CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdGmgnCellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);
    procedure btn_CancelClick(Sender: TObject);
    procedure grdMaster_DetailClickCell(Sender: TObject; DataColDown,
      DataRowDown, DataColUp, DataRowUp: Integer; DownPos,
      UpPos: TtsClickPosition);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);




  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vcod_jisa, UV_vDesc_name, UV_vNum_jumin, UV_vNum_id, UV_vDesc_dc, UV_vDesc_Dept, UV_vConsult_date, UV_vConsult_time, UV_vFile_name,
    UV_vCod_sawon, UV_vSawon_name, UV_vConsult_note, UV_vNext_gn_date, UV_nH_name, UV_nsokun_name, UV_vIdx,
    UV_vDconsult_date, UV_vDconsult_time, UV_vDSawon_name, UV_vDConsult_note, UV_vDNext_gn_date, UV_vDH_Code, UV_vDH_name, UV_vDsokun_name,
    UV_vHul_h, UV_vHul_l, UV_vBiman, UV_vSinjang, UV_vChejung, UV_vear_left5, UV_vear_left10, UV_vear_lef20, UV_vear_left30, UV_vear_left40,
    UV_vear_left60, UV_vear_right5, UV_vear_right10, UV_vear_right20, UV_vear_right30, UV_vear_right40, UV_vear_right60, UV_vdesc_ear,
    UV_veye_lman, UV_veye_rman, UV_veye_lkyo, UV_veye_rkyo, UV_vanap_left, UV_vanap_right, UV_vgubn_seksin, UV_vhyungwi, UV_vbookwi, UV_vDat_gmgn,
    UV_vInjouk_num_id, UV_vCod_hm, UV_vDCon_type  : Variant;

    UV_vDat_gmgn_G, UV_vCod_janbi_G, UV_vCod_hul_G, UV_vCod_cancer_G, UV_vCod_chuga_G, UV_vNum_tel_G, UV_vGubn_nosin_G, UV_vNum_id_G,
    UV_vGubn_nscg_G, UV_vGubn_adult_G, UV_vGubn_adcg_G, UV_vGubn_life_G, UV_vGubn_lfcg_G, UV_vDesc_yhname_G, UV_vDesc_saler_G, UV_vGubn_jinch_G : Variant;

    //���װ���
    UV_HPan1_check, UV_HPan2_check, UV_HPan3_check, UV_HPan4_check, UV_HPan5_check, UV_HPan6_check,
    UV_HPan7_check, UV_HPan8_check, UV_HPan9_check, UV_HPan10_check, UV_HPan11_check, UV_HPan12_check : Boolean;

    //���װ˻�
    UV_vCod_chuga, UV_vCod_totyh, UV_vCod_chuga2, UV_vCod_totyh2,UV_vTot_cod_hm, UV_vTot_cod_hm_name, UV_vHul_Gulkwa, UV_vHul_dat_gmgn, UV_vHul_Low_High, UV_vFont_color, UV_vHul_num_id : Variant;
    //���˻�
    UV_vDat_gmgn_Jangbi, UV_vDesc_kor, UV_vCod_sokun_Jangbi, UV_vGubn_pan, UV_vDesc_sokun_Jangbi, UV_vDesc_sbsg_Jangbi : Variant;
    //���μҰ�
    UV_vDat_gmgn_sokun, UV_vgubn_sokun, UV_vCod_sokun, UV_vDesc_sokun : Variant;

    //Ư�������ڷ�(������ �Ұ�)
    UV_vTCod_matter, UV_vTDesc_matter, UV_vTGubn_LR, UV_vTNum_seq, UV_vTCod_pan, UV_vTcod_sokun, UV_vTcod_jochi, UV_vTcod_ujs, UV_vTCod_sahu1, UV_vTCod_sahu2, UV_vTCod_jilhwan, UV_vTCod_janggi,
    UV_vTDesc_etcSokun, UV_vTDesc_etcJochi, UV_vTDesc_bigo, UV_vTDesc_KMI_bigo, UV_vTTot_jilhan : Variant;

    //Ư���������(��ġ��)
    UV_vCod_Mhm, UV_vCod_Hhm, UV_vTDesc_kor, UV_vIms_low, UV_vIms_high, UV_vDesc_unit, UV_vGulkwa : Variant;

    //Ư���������(������)
    UV_vCod_Mhm2, UV_vCod_Hhm2, UV_vDesc_kor2, UV_vCKMSGulkwa2, UV_vHEMSGulkwa2, UV_Gubun2, UV_vCKMSCODE : Variant;

    //���� �������ΰ�� üũ
    bCK_Num_id : Boolean;
    index3     : integer;
    //SMS����
    UV_CkByte  : Boolean;
    UV_vSms_content : Variant;

    //CTI ����..
    UV_sCallNo : Extended;               // CID ��ȣ
    UV_sCustNo, UV_vCallNo : String;     // �������� ����

    FUNCTION  UF_CONNEC(sTemp : String) : String;
    procedure UP_CTI_Login(sLogin : string);

    procedure UP_VarMemSet_Tkgum(iLength : Integer; sGubun : String);
    procedure CreateTextArrayTkgum;
    procedure UP_TkSokun_Display;
    procedure UP_TkGulkwa_Display(sCod_Prf, sCod_chuga, sDan_sokun, sDesc_EK, sDesc_UBio, sDesc_PBio : String; iSex : integer);
    procedure UP_Grid_Clear(iRow : Integer; sGubun : String);

    function  UF_CodeEdit(iMode : integer; sCode : string) : string;
    procedure CreateTextArray_Gmgn;
    procedure UP_VarMemSet_Gmgn(iLength: Integer);
    procedure Gmgn_retest;
    procedure UP_GumsaCheck(gubun, sID_1, sJisa_1, sNsGubn_1, sNsCg_1,
      sDate1, sDate2: String);
    function UF_Sokun_cnvt(sCod_sokun: string): String;
    function UF_Hmatter_cnvt(sCod_Hmatter: string): String;
  public
    { Public declarations }
    UV_sHCode_Check, UV_sNum_id : String;

    //CTI����
    huTime : Integer; //��ó�� �ð�

    chk_gmgn, UV_sCod_dc : String;

    UV_bEdit : Boolean;
  end;

  const UC_mcij = 20;

var
  frmCK5I17: TfrmCK5I17;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;
  
implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     CK5I17A, CK5I17B, CK5I171, CK3I015, CK5I161, CK5I162, CK5I163, CK5I164;

{$R *.DFM}

procedure TfrmCK5I17.UP_Grid_Clear(iRow : Integer; sGubun : String);
begin
   if sGubun = 'grdResult' then
   begin
      UV_vCod_Mhm[iRow]   := '';
      UV_vCod_Hhm[iRow]   := '';
      UV_vTDesc_kor[iRow] := '';
      UV_vIms_low[iRow]   := '';
      UV_vIms_high[iRow]  := '';
      UV_vDesc_unit[iRow] := '';
      UV_vGulkwa[iRow]    := '';
   end
   else if sGubun = 'grdResult2' then
   begin
      UV_vCod_Mhm2[iRow]    := '';
      UV_vCod_Hhm2[iRow]    := '';
      UV_vDesc_kor2[iRow]   := '';
      UV_vCKMSGulkwa2[iRow] := '';
      UV_vHEMSGulkwa2[iRow] := '';
      UV_Gubun2[iRow]       := '';
      UV_vCKMSCODE[iRow]    := '';
   end;
end;

function TfrmCK5I17.UF_CodeEdit(iMode : integer; sCode : string) : string;
var sSqlText : String;
begin
   //sMode : [1]�Ұ� [2]��ġ [3]ǥ�����
   Result := ''; sSqlText := '';

   case iMode of
      1 : sSqlText := ' SELECT �Ұ߸�  Desc_Code FROM HEMS..Comm20  WHERE �Ұ��ڵ� = ''' + sCode + ''' ';
      2 : sSqlText := ' SELECT ��ġ��  Desc_Code FROM HEMS..Comm21  WHERE ��ġ�ڵ� = ''' + sCode + ''' ';
      3 : sSqlText := ' SELECT CODE_NM Desc_Code FROM HEMS..HM_Comm WHERE CODE_ID  = ''HM05D'' AND CODE = ''' + sCode + ''' ';
   end;

   with dmComm.qryShare do
   begin
      Active := False;
      SQL.Clear;
      SQL.Text := sSqlText;
      Active := True;

      if RecordCount > 0 then Result := FieldByName('Desc_Code').AsString
      else                    Result := '';

      Active := False;
   end;

end;
procedure TfrmCK5I17.UP_VarMemSet_Tkgum(iLength : Integer; sGubun : String);
begin
   // Variant Memory Allocation
   if sGubun = 'grdTkSokun' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vTCod_matter    := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_matter   := VarArrayCreate([0, 0], varOleStr);
         UV_vTGubn_LR       := VarArrayCreate([0, 0], varOleStr);
         UV_vTNum_seq       := VarArrayCreate([0, 0], varOleStr);
         UV_vTCod_pan       := VarArrayCreate([0, 0], varOleStr);
         UV_vTcod_sokun     := VarArrayCreate([0, 0], varOleStr);
         UV_vTcod_jochi     := VarArrayCreate([0, 0], varOleStr);
         UV_vTcod_ujs       := VarArrayCreate([0, 0], varOleStr);
         UV_vTCod_sahu1     := VarArrayCreate([0, 0], varOleStr);
         UV_vTCod_sahu2     := VarArrayCreate([0, 0], varOleStr);
         UV_vTCod_jilhwan   := VarArrayCreate([0, 0], varOleStr);
         UV_vTCod_janggi    := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_etcSokun := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_etcJochi := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_bigo     := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_KMI_bigo := VarArrayCreate([0, 0], varOleStr);
         UV_vTTot_jilhan    := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vTCod_matter,    iLength);
         VarArrayRedim(UV_vTDesc_matter,   iLength);
         VarArrayRedim(UV_vTGubn_LR,       iLength);
         VarArrayRedim(UV_vTNum_seq,       iLength);
         VarArrayRedim(UV_vTCod_pan,       iLength);
         VarArrayRedim(UV_vTcod_sokun,     iLength);
         VarArrayRedim(UV_vTcod_jochi,     iLength);
         VarArrayRedim(UV_vTcod_ujs,       iLength);
         VarArrayRedim(UV_vTCod_sahu1,     iLength);
         VarArrayRedim(UV_vTCod_sahu2,     iLength);
         VarArrayRedim(UV_vTCod_jilhwan,   iLength);
         VarArrayRedim(UV_vTCod_janggi,    iLength);
         VarArrayRedim(UV_vTDesc_etcSokun, iLength);
         VarArrayRedim(UV_vTDesc_etcJochi, iLength);
         VarArrayRedim(UV_vTDesc_bigo,     iLength);
         VarArrayRedim(UV_vTDesc_KMI_bigo, iLength);
         VarArrayRedim(UV_vTTot_jilhan,    iLength);
      end;
   end
   else if sGubun = 'grdResult' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vCod_Mhm   := VarArrayCreate([0, 0], varOleStr);
         UV_vCod_Hhm   := VarArrayCreate([0, 0], varOleStr);
         UV_vTDesc_kor := VarArrayCreate([0, 0], varOleStr);
         UV_vIms_low   := VarArrayCreate([0, 0], varOleStr);
         UV_vIms_high  := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_unit := VarArrayCreate([0, 0], varOleStr);
         UV_vGulkwa    := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vCod_Mhm,   iLength);
         VarArrayRedim(UV_vCod_Hhm,   iLength);
         VarArrayRedim(UV_vTDesc_kor, iLength);
         VarArrayRedim(UV_vIms_low,   iLength);
         VarArrayRedim(UV_vIms_high,  iLength);
         VarArrayRedim(UV_vDesc_unit, iLength);
         VarArrayRedim(UV_vGulkwa,    iLength);
      end;
   end
   else if sGubun = 'grdResult2' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vCod_Mhm2     := VarArrayCreate([0, 0], varOleStr);
         UV_vCod_Hhm2     := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_kor2    := VarArrayCreate([0, 0], varOleStr);
         UV_vCKMSGulkwa2  := VarArrayCreate([0, 0], varOleStr);
         UV_vHEMSGulkwa2  := VarArrayCreate([0, 0], varOleStr);
         UV_Gubun2        := VarArrayCreate([0, 0], varOleStr);
         UV_vCKMSCODE     := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vCod_Mhm2,    iLength);
         VarArrayRedim(UV_vCod_Hhm2,    iLength);
         VarArrayRedim(UV_vDesc_kor2,   iLength);
         VarArrayRedim(UV_vCKMSGulkwa2, iLength);
         VarArrayRedim(UV_vHEMSGulkwa2, iLength);
         VarArrayRedim(UV_Gubun2,       iLength);
         VarArrayRedim(UV_vCKMSCODE,    iLength);
      end;
   end;
end;

procedure TfrmCK5I17.CreateTextArrayTkgum;
var
   sSqlText : string;
   iSex     : integer;
begin
   //Ư�������ڷ� �ʱ�ȭ...
   grdTkSokun.Rows := 0; UP_VarMemSet_Tkgum(0, 'grdTkSokun');
   grdResult.Rows  := 0; UP_VarMemSet_Tkgum(0, 'grdResult');
   grdResult2.Rows := 0; UP_VarMemSet_Tkgum(0, 'grdResult2');
   GP_FieldClear(pnlTkgum);
   MM_Tetcsokun.Lines.Clear;
   MM_Tetcbigo.Lines.Clear;
   MM_TetcJochi.Lines.Clear;
   MM_TKMI_bigo.Lines.Clear;
   Pnl_Complete.Caption := '';

   with qryTkgumCheck do
   begin
      Active := False;
      Sql.Clear;
      Sql.Text := ' SELECT G.num_id, dbo.F_GET_SAWON_FIND(T.cod_Doctor) Desc_Doctor, T.* ' + #13 +
                  ' FROM GUMGIN G WITH (NOLOCK) INNER JOIN tkgum T WITH (NOLOCK)ON G.cod_jisa = T.cod_jisa AND G.num_jumin = T.num_jumin AND G.dat_gmgn = T.dat_gmgn AND G.gubn_tkgm = T.gubn_cha ' + #13 +
                  ' WHERE G.cod_jisa  = ''' + copy(cmb_gmgn.text,12,2) + ''' ' + #13 +
                  '   and G.num_jumin = ''' + edtJumin.Text            + ''' ' + #13 +
                  '   and G.dat_gmgn  = ''' + copy(cmb_gmgn.text,1,4) + copy(cmb_gmgn.text,6,2) + copy(cmb_gmgn.text,9,2) + ''' ';
      Active := True;

      if RecordCount > 0 then
      begin
         //�ǵ��ǻ��ڵ�
         EdtCod_Doctor.Text  := FieldByName('cod_Doctor').AsString;
         //�ǵ��ǻ��
         EdtDesc_Doctor.Text := FieldByName('desc_Doctor').AsString;
         //Ư������
         EdtGubn_Cha.Text    := FieldByName('gubn_cha').AsString;

         //�����Ϸ� üũ...
         if FieldByName('Ysno_Complete').AsString = 'Y' then Pnl_Complete.Caption := '�����Ϸ�'
         else                                                Pnl_Complete.Caption := '������(�̿Ϸ�)';

         iSex := 0;
         Case StrToInt(copy(FieldByName('num_jumin').AsString,7,1)) of
            1, 3, 5, 7, 9 : iSex := 1;
            2, 4, 6, 8    : iSex := 2;
         end;

         with qryTk_sokun2009 do
         begin
            Active := False;
            ParamByName('cod_jisa').AsString   := qryTkgumCheck.FieldByName('cod_jisa').AsString;
            ParamByName('num_jumin').AsString  := qryTkgumCheck.FieldByName('num_jumin').AsString;
            ParamByName('dat_gmgn').AsString   := qryTkgumCheck.FieldByName('dat_gmgn').AsString;
            ParamByName('gubn_cha').AsString   := qryTkgumCheck.FieldByName('gubn_cha').AsString;
            Active := True;
         end;

         with qryTK_Hul do
         begin
            Active := False;
            ParamByName('cod_jisa').AsString    := qryTkgumCheck.FieldByName('cod_jisa').AsString;
            ParamByName('num_id').AsString      := qryTkgumCheck.FieldByName('num_id').AsString;
            ParamByName('dat_gmgn').AsString    := qryTkgumCheck.FieldByName('dat_gmgn').AsString;
            Active := True;
         end;

         with qryTK_Jangbi do
         begin
            Active := False;
            ParamByName('cod_jisa').AsString    := qryTkgumCheck.FieldByName('cod_jisa').AsString;
            ParamByName('num_id').AsString      := qryTkgumCheck.FieldByName('num_id').AsString;
            ParamByName('dat_gmgn').AsString    := qryTkgumCheck.FieldByName('dat_gmgn').AsString;
            Active := True;
         end;

         with qryTk_Gulkwa do
         begin
            Active := False;
            ParamByName('cod_jisa').AsString   := qryTkgumCheck.FieldByName('cod_jisa').AsString;
            ParamByName('num_jumin').AsString  := qryTkgumCheck.FieldByName('num_jumin').AsString;
            ParamByName('dat_gmgn').AsString   := qryTkgumCheck.FieldByName('dat_gmgn').AsString;
            ParamByName('gubn_cha').AsString   := qryTkgumCheck.FieldByName('gubn_cha').AsString;
            Active := True;
         end;

         with qryTKicho do
         begin
            Active := False;
            ParamByName('cod_jisa').AsString := qryTkgumCheck.FieldByName('cod_jisa').AsString;
            ParamByName('num_id').AsString   := qryTkgumCheck.FieldByName('num_id').AsString;
            ParamByName('dat_gmgn').AsString := qryTkgumCheck.FieldByName('dat_gmgn').AsString;
            Active := True;
         end;

         if Not qryTk_Comm12.Active then qryTk_Comm12.Active := True;

         //�Ұ� Display...
         UP_TkSokun_Display;
         //��� Display...
         UP_TkGulkwa_Display(qryTkgumCheck.FieldByName('cod_prf').AsString,     //[Ư��]��������
                             qryTkgumCheck.FieldByName('cod_chuga').AsString,   //[Ư��]�˻��߰�
                             FieldByName('dan_sokun').AsString,                 //ġ���Ұ�
                             FieldByName('Desc_EK').AsString,                   //�̰�Ұ�
                             FieldByName('Desc_UBio').AsString,                 //�Һ����� �Ұ�
                             FieldByName('Desc_PBio').AsString,                 //���㼼�� �Ұ�
                             iSex);                                             //����(1:��, 2:��)
      end;
   end;

   qryTkgumCheck.Active   := False;
   qryTK_Hul.Active       := False;
   qryTK_Jangbi.Active    := False;
   qryTk_sokun2009.Active := False;
   qryTk_gulkwa.Active    := False;
   qryTk_Comm12.Active    := False;
   qryTKicho.Active    := False;
end;

procedure TfrmCK5I17.UP_TkSokun_Display;
var
   iCnt : integer;
begin
   iCnt := 0;

   with qryTk_sokun2009 do
   begin
      while Not Eof do
      begin
         UP_VarMemSet_Tkgum(iCnt, 'grdTkSokun');

         UV_vTCod_matter[iCnt]    := qryTk_sokun2009.FieldByName('�����ڵ�').AsString;
         UV_vTDesc_matter[iCnt]   := qryTk_sokun2009.FieldByName('���ع�����').AsString;
         UV_vTGubn_LR[iCnt]       := qryTk_sokun2009.FieldByName('gubn_LR').AsString;
         UV_vTNum_seq[iCnt]       := qryTk_sokun2009.FieldByName('num_seq').AsString;

         UV_vTCod_pan[iCnt]       := qryTk_sokun2009.FieldByName('cod_pan').AsString;
         UV_vTcod_sokun[iCnt]     := qryTk_sokun2009.FieldByName('cod_sokun').AsString;
         UV_vTcod_jochi[iCnt]     := qryTk_sokun2009.FieldByName('cod_jochi').AsString;
         UV_vTcod_ujs[iCnt]       := qryTk_sokun2009.FieldByName('cod_ujs').AsString;
         UV_vTCod_sahu1[iCnt]     := qryTk_sokun2009.FieldByName('cod_sahu1').AsString;
         UV_vTCod_sahu2[iCnt]     := qryTk_sokun2009.FieldByName('cod_sahu2').AsString;
         UV_vTCod_jilhwan[iCnt]   := Trim(qryTk_sokun2009.FieldByName('cod_jilhan').AsString);
         UV_vTCod_janggi[iCnt]    := Trim(qryTk_sokun2009.FieldByName('cod_janggi').AsString);

         UV_vTDesc_etcSokun[iCnt] := Trim(qryTk_sokun2009.FieldByName('desc_etcSokun').AsString);
         UV_vTDesc_etcJochi[iCnt] := Trim(qryTk_sokun2009.FieldByName('desc_etcJochi').AsString);
         UV_vTDesc_bigo[iCnt]     := Trim(qryTk_sokun2009.FieldByName('desc_bigo').AsString);
         UV_vTDesc_KMI_bigo[iCnt] := Trim(qryTk_sokun2009.FieldByName('desc_KMI_bigo').AsString);
         UV_vTTot_jilhan[iCnt]    := qryTk_sokun2009.FieldByName('Tot_jilhan').AsString;
         inc(iCnt);

         Next;
      end;
   end;

   grdTkSokun.Rows := iCnt;
end;

procedure TfrmCK5I17.UP_TkGulkwa_Display(sCod_Prf, sCod_chuga, sDan_sokun, sDesc_EK, sDesc_UBio, sDesc_PBio : String; iSex : integer);
var
   sSqlText, sHmCode, sHmCodeDump, sDat_gmgn, sTemp : string;
   iTemp, iRet, iCountX, ihangmok, iHamgmokCnt, iHamgmokCnt2, i : integer;
   vCod_Hangmok : Variant;

   eSinjang, eChejung, eHul_h, eHul_l, eBookwi,
   eEAR_LEFT5,   eEAR_RIGHT5,   eEAR_LEFT10,  eEAR_RIGHT10,  eEAR_LEFT20,  eEAR_RIGHT20,
   eEAR_LEFT30,  eEAR_RIGHT30,  eEAR_LEFT40,  eEAR_RIGHT40,  eEAR_LEFT60,  eEAR_RIGHT60,
   eEAR_LEFT5k,  eEAR_RIGHT5k,  eEAR_LEFT10k, eEAR_RIGHT10k, eEAR_LEFT20k, eEAR_RIGHT20k,
   eEAR_LEFT30k, eEAR_RIGHT30k, eEAR_LEFT40k, eEAR_RIGHT40k, eEAR_LEFT60k, eEAR_RIGHT60k,
   eFEV1,        eFEV1_p,       eFVC,         eFVC_p,        eFEVFVC,      ePEF,  ePEF_p,
   eU003,        eC025,         eC026,        eC028,         eC032 : Extended;

   sEye_LEft, sEye_Right, sU004, sU005, sU009, sJJD7, sPekiSokun, sC083, sUrine_Char : string;
begin
   sEye_LEft := ''; sEye_Right := ''; sU004 := ''; sU005 := ''; sU009 := ''; sJJD7 := ''; sPekiSokun := ''; sDat_gmgn := '';  sC083 := ''; sTemp := ''; sUrine_Char := '';

   eSinjang     := 0; eChejung      := 0; eHul_h       := 0; eHul_l        := 0; eBookwi      := 0;
   eEAR_LEFT5   := 0; eEAR_RIGHT5   := 0; eEAR_LEFT10  := 0; eEAR_RIGHT10  := 0; eEAR_LEFT20  := 0; eEAR_RIGHT20  := 0;
   eEAR_LEFT30  := 0; eEAR_RIGHT30  := 0; eEAR_LEFT40  := 0; eEAR_RIGHT40  := 0; eEAR_LEFT60  := 0; eEAR_RIGHT60  := 0;
   eEAR_LEFT5k  := 0; eEAR_RIGHT5k  := 0; eEAR_LEFT10k := 0; eEAR_RIGHT10k := 0; eEAR_LEFT20k := 0; eEAR_RIGHT20k := 0;
   eEAR_LEFT30k := 0; eEAR_RIGHT30k := 0; eEAR_LEFT40k := 0; eEAR_RIGHT40k := 0; eEAR_LEFT60k := 0; eEAR_RIGHT60k := 0;
   eFEV1        := 0; eFEV1_p       := 0; eFVC         := 0; eFVC_p        := 0; eFEVFVC      := 0; ePEF          := 0; ePEF_p := 0;
   eU003        := 0; eC025         := 0; eC026        := 0; eC028         := 0; eC032        := 0;

   //�˻��׸� ���...
   //------------------------------------------------------------------------------
   if Trim(sCod_Prf) <> '' then
   begin
      sSqlText := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';

      for iTemp := 1 to length(sCod_Prf) div 4 do
      begin
         if iTemp = length(sCod_Prf) div 4 then sSqlText := sSqlText + ' A.cod_pf = ''' + Copy(sCod_Prf, (iTemp*4)-3, 4) + ''''
         else                                   sSqlText := sSqlText + ' A.cod_pf = ''' + Copy(sCod_Prf, (iTemp*4)-3, 4) + ''' or ';
      end;

      if (pos('TC02',Trim(sCod_Prf)) > 0) or (pos('TC08',Trim(sCod_Prf)) > 0) or
         (pos('TC11',Trim(sCod_Prf)) > 0) or (pos('TC17',Trim(sCod_Prf)) > 0) or
         (pos('TC97',Trim(sCod_Prf)) > 0) or (pos('TCA8',Trim(sCod_Prf)) > 0) or
         (pos('TCA9',Trim(sCod_Prf)) > 0) or (pos('TCB1',Trim(sCod_Prf)) > 0) or
         (pos('TCB5',Trim(sCod_Prf)) > 0) or (pos('TCB7',Trim(sCod_Prf)) > 0) or
         (pos('TC98',Trim(sCod_Prf)) > 0) then  sSqlText := sSqlText + ' ) and B.cod_pf <> ''SE81'' and B.cod_hm <> ''JJ01'' GROUP BY B.cod_hm '
      else                                      sSqlText := sSqlText + ' ) and B.cod_pf <> ''SE81'' and B.cod_hm <> ''JJ01'' and B.cod_hm <> ''JJ08'' GROUP BY B.cod_hm ';

      QryTemp.Active := False;
      QryTemp.SQL.Clear;
      QryTemp.SQL.Text := sSqlText;
      QryTemp.Active := True;

      if QryTemp.RecordCount > 0 then
      begin
         while not QryTemp.Eof do
         begin
            sHmCode := sHmCode + QryTemp.FieldByName('cod_hm').AsString;
            QryTemp.Next;
         end;
      end;
      QryTemp.Active := False;
   end;
   sHmCode := sHmCode + sCod_chuga;
   //===========================================================================

   //�ߺ��˻��׸� ����...
   iRet := GF_MulToSingle(sHmCode, 4, vCod_Hangmok);
   sHmCodeDump := '';
   for iCountX := 0 to iRet - 1 do
   begin
      if pos(vCod_Hangmok[iCountX],sHmCodeDump) <= 0 then
         sHmCodeDump := sHmCodeDump + vCod_Hangmok[iCountX]
   end;
   sHmCode := sHmCodeDump;

   //�˻��׸� ��������..
   //UV_sTkHangmokList := sHmCode;
   //UV_sTkPrfList     := UV_vCod_prf[NewRow-1];

   //�˻��׸� Display
   iRet := GF_MulToSingle(sHmCode, 4, vCod_Hangmok);

   //���� ����...
   //---------------------------------------------------------------------------
   with qryTKicho do
   begin
      if RecordCount > 0 then
      begin
         //��������
         sDat_gmgn := FieldByName('Dat_gmgn').AsString;
         //����
         if FieldByName('Sinjang').AsString <> '' Then eSinjang := FieldByName('Sinjang').AsFloat;

         //ü��
         if FieldByName('Chejung').AsString <> '' Then eChejung := FieldByName('Chejung').AsFloat;

         //�÷�
         if (FieldByName('Eye_lman').AsString > '0') or (FieldByName('Eye_lkyo').AsString > '0') then
         begin
            if (FieldByName('Eye_lman').AsString = '9.9')  Or
               (FieldByName('Eye_lkyo').AsString = '9.9')  Then sEye_Left := '(��)�Ǹ�'
            else if FieldByName('Eye_lman').AsString > '0' Then sEye_Left := '(��)' + FieldByName('Eye_lman').AsString
            else if FieldByName('Eye_lkyo').AsString > '0' Then sEye_Left := '(��)' + FieldByName('Eye_lkyo').AsString;
         End;

         if (FieldByName('Eye_rman').AsString > '0') or
            (FieldByName('Eye_rkyo').AsString > '0') then
         begin
            if (FieldByName('Eye_rman').AsString = '9.9')  or
               (FieldByName('Eye_rkyo').AsString = '9.9')  then sEye_Right := '(��)�Ǹ�'
            else if FieldByName('Eye_lman').AsString > '0' then sEye_Right := '(��)' + FieldByName('Eye_rman').AsString
            else if FieldByName('Eye_lkyo').AsString > '0' then sEye_Right := '(��)' + FieldByName('Eye_rkyo').AsString;
         end;

         //����
         if (pos('JJ05', sHmCode) > 0) then eBookwi   := FieldByName('Bookwi').AsFloat;

         //����
         if (pos('JJ08', sHmCode) > 0) then
         begin
            eHul_h   := FieldByName('hul_h').AsFloat;
            eHul_l   := FieldByName('hul_l').AsFloat;
         end;

         //û��(�⺻)
         if pos('JJ06', sHmCode) > 0 then
         begin
            eEAR_LEFT20   := FieldByName('EAR_LEFT20').AsFloat;
            eEAR_RIGHT20  := FieldByName('EAR_RIGHT20').AsFloat;
            eEAR_LEFT30   := FieldByName('EAR_LEFT30').AsFloat;
            eEAR_RIGHT30  := FieldByName('EAR_RIGHT30').AsFloat;
            eEAR_LEFT40   := FieldByName('EAR_LEFT40').AsFloat;
            eEAR_RIGHT40  := FieldByName('EAR_RIGHT40').AsFloat;
         end;

         //û��(����)
         if pos('JJ76', sHmCode) > 0 then
         begin
            eEAR_LEFT5    := FieldByName('EAR_LEFT5').AsFloat;
            eEAR_RIGHT5   := FieldByName('EAR_RIGHT5').AsFloat;
            eEAR_LEFT10   := FieldByName('EAR_LEFT10').AsFloat;
            eEAR_RIGHT10  := FieldByName('EAR_RIGHT10').AsFloat;
            eEAR_LEFT20   := FieldByName('EAR_LEFT20').AsFloat;
            eEAR_RIGHT20  := FieldByName('EAR_RIGHT20').AsFloat;
            eEAR_LEFT30   := FieldByName('EAR_LEFT30').AsFloat;
            eEAR_RIGHT30  := FieldByName('EAR_RIGHT30').AsFloat;
            eEAR_LEFT40   := FieldByName('EAR_LEFT40').AsFloat;
            eEAR_RIGHT40  := FieldByName('EAR_RIGHT40').AsFloat;
            eEAR_LEFT60   := FieldByName('EAR_LEFT60').AsFloat;
            eEAR_RIGHT60  := FieldByName('EAR_RIGHT60').AsFloat;

            eEAR_LEFT5k   := FieldByName('EAR_LEFT5k').AsFloat;
            eEAR_RIGHT5k  := FieldByName('EAR_RIGHT5k').AsFloat;
            eEAR_LEFT10k  := FieldByName('EAR_LEFT10k').AsFloat;
            eEAR_RIGHT10k := FieldByName('EAR_RIGHT10k').AsFloat;
            eEAR_LEFT20k  := FieldByName('EAR_LEFT20k').AsFloat;
            eEAR_RIGHT20k := FieldByName('EAR_RIGHT20k').AsFloat;
            eEAR_LEFT30k  := FieldByName('EAR_LEFT30k').AsFloat;
            eEAR_RIGHT30k := FieldByName('EAR_RIGHT30k').AsFloat;
            eEAR_LEFT40k  := FieldByName('EAR_LEFT40k').AsFloat;
            eEAR_RIGHT40k := FieldByName('EAR_RIGHT40k').AsFloat;
            eEAR_LEFT60k  := FieldByName('EAR_LEFT60k').AsFloat;
            eEAR_RIGHT60k := FieldByName('EAR_RIGHT60k').AsFloat;
         end;

         //����
         if (pos('JJ07', sHmCode) > 0) then
         begin
            eFEV1   := FieldByName('YUL_PEKI2').AsFloat;
            eFEV1_p := FieldByName('RSLT_PEKI2').AsFloat;
            eFVC    := FieldByName('YUL_PEKI1').AsFloat;
            eFVC_p  := FieldByName('RSLT_PEKI1').AsFloat;
            eFEVFVC := FieldByName('YUL_PEKI4').AsFloat;
            ePEF    := FieldByName('YUL_PEKI3').AsFloat;
            ePEF_p  := FieldByName('RSLT_PEKI3').AsFloat;
            sPekiSokun := FieldByName('DESC_PEKI').AsString;
         end;

         //��˻�(PH)
         if (pos('U003', sHmCode) > 0) then eU003 := FieldByName('gubn_u003').AsFloat;

         //��˻�(�ܹ�)
         if (pos('U004', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u004').AsString = '1' then sU004 := '����'
            else if  FieldByName('gubn_u004').AsString = '2' then sU004 := '��'
            else if  FieldByName('gubn_u004').AsString = '3' then sU004 := '+1'
            else if  FieldByName('gubn_u004').AsString = '4' then sU004 := '+2'
            else if  FieldByName('gubn_u004').AsString = '5' then sU004 := '+3'
            else if  FieldByName('gubn_u004').AsString = '6' then sU004 := '+4';
         end;

         //��˻�(��)
         if (pos('U005', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u005').AsString = '1' then sU005 := '����'
            else if  FieldByName('gubn_u005').AsString = '2' then sU005 := '��'
            else if  FieldByName('gubn_u005').AsString = '3' then sU005 := '+1'
            else if  FieldByName('gubn_u005').AsString = '4' then sU005 := '+2'
            else if  FieldByName('gubn_u005').AsString = '5' then sU005 := '+3'
            else if  FieldByName('gubn_u005').AsString = '6' then sU005 := '+4';
         end;

         //��˻�(������)
         if (pos('U009', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u009').AsString = '1' then sU009 := '����'
            else if  FieldByName('gubn_u009').AsString = '2' then sU009 := '��'
            else if  FieldByName('gubn_u009').AsString = '3' then sU009 := '+1'
            else if  FieldByName('gubn_u009').AsString = '4' then sU009 := '+2'
            else if  FieldByName('gubn_u009').AsString = '5' then sU009 := '+3'
            else if  FieldByName('gubn_u009').AsString = '6' then sU009 := '+4';
         end;

      end;
      Active := False;
   end;
   //===========================================================================

   ihangmok := 0;      //�˻� ����...
   iHamgmokCnt := 0;   //��ġ�˻� Cnt...
   iHamgmokCnt2 := 0;  //���ڰ˻� Cnt...

   //[��ġ]����
   UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
   UP_Grid_Clear(iHamgmokCnt,'grdResult');

   UV_vCod_Mhm[iHamgmokCnt]  := '----';
   UV_vCod_Hhm[iHamgmokCnt]  := '****';
   UV_vTDesc_kor[iHamgmokCnt] := '����';
   UV_vGulkwa[iHamgmokCnt]   := eSinjang;
   UV_vIms_low[iHamgmokCnt]  := '';
   UV_vIms_high[iHamgmokCnt] := '';
   UV_vDesc_unit[iHamgmokCnt] := 'Cm';
   Inc(iHamgmokCnt);

   //[��ġ]ü��
   UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
   UP_Grid_Clear(iHamgmokCnt,'grdResult');

   UV_vCod_Mhm[iHamgmokCnt]  := '----';
   UV_vCod_Hhm[iHamgmokCnt]  := '****';
   UV_vTDesc_kor[iHamgmokCnt] := 'ü��';
   UV_vGulkwa[iHamgmokCnt]   := eChejung;
   UV_vIms_low[iHamgmokCnt]  := '';
   UV_vIms_high[iHamgmokCnt] := '';
   UV_vDesc_unit[iHamgmokCnt] := 'Kg';
   Inc(iHamgmokCnt);

   while ihangmok <= iRet - 1 do
   begin
      if vCod_Hangmok[ihangmok] = 'JJD7' then
      begin
         //[����]��ι�缱(Lateral_Lt) ���...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJD7';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0433';
         UV_vDesc_kor2[iHamgmokCnt2] := '��ι�缱(Lateral_Lt)';

         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0433''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sJJD7;

         qryTk_gulkwa.Filter := 'cod_Hhm = ''0433''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'SE81' then
      begin
         //[����]�����˻� ���...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'SE81';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0373';
         UV_vDesc_kor2[iHamgmokCnt2] := '�����˻�(ġ����,ġ�ֿ���)';

         //[2017.02.06-������]��������� �׸񿡼� ��������...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0373''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDan_sokun;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0373''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJF4' then
      begin
         for i := 1 to 2 do
         begin
            case i of
               1 : begin
                      //[����]�̰�˻� ���...
                      UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
                      UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

                      UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJF4';
                      UV_vCod_Hhm2[iHamgmokCnt2]  := '0507';
                      UV_vDesc_kor2[iHamgmokCnt2] := '���� ����(�̰�)-��';

                      //[2017.02.06-������]��������� �׸񿡼� ��������...
                      //---------------------------------------------------------------------
                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0507''';
                      if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
                      //=====================================================================

                      UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_EK;
                      qryTk_gulkwa.Filter := 'cod_Hhm = ''0507''';
                      if qryTk_gulkwa.RecordCount > 0 then
                      begin
                         UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
                         UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
                      end;
                   end;
               2 : begin
                      //[����]�̰�˻� ���...
                      UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
                      UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

                      UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJF4';
                      UV_vCod_Hhm2[iHamgmokCnt2]  := '0532';
                      UV_vDesc_kor2[iHamgmokCnt2] := '���� ����(�̰�)-��';

                      //[2017.02.06-������]��������� �׸񿡼� ��������...
                      //---------------------------------------------------------------------
                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0532''';
                      if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
                      //=====================================================================

                      UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_EK;
                      qryTk_gulkwa.Filter := 'cod_Hhm = ''0532''';
                      if qryTk_gulkwa.RecordCount > 0 then
                      begin
                         UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
                         UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
                      end;
                   end;
            end;
            Inc(iHamgmokCnt2);
         end;
      end
      else if vCod_Hangmok[ihangmok] = 'JJFG' then
      begin
         //[����]�Һ����������˻� ���...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJFG';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0399';
         UV_vDesc_kor2[iHamgmokCnt2] := '���Ĵ��ݶ��˻�(�似���˻�)';

         //[2017.02.06-������]��������� �׸񿡼� ��������...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0399''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_UBio;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0399''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJFH' then
      begin
         //[����]���㼼���˻� ���...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJFH';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0270';
         UV_vDesc_kor2[iHamgmokCnt2] := '���㼼�������˻�';

         //[2017.02.06-������]��������� �׸񿡼� ��������...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0270''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_PBio;

         qryTk_gulkwa.Filter := 'cod_Hhm = ''0270''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJ05' then
      begin
         //[��ġ]���εѷ�...
         UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
         UP_Grid_Clear(iHamgmokCnt,'grdResult');

         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ05';
         UV_vCod_Hhm[iHamgmokCnt]  := '0522';
         UV_vTDesc_kor[iHamgmokCnt] := '���εѷ�';
         UV_vGulkwa[iHamgmokCnt]   := eBookwi;

         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0522''';
         if qryTk_Comm12.RecordCount > 0 then
         begin
            if iSex = 1 then
            begin
               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
            end
            else
            begin
               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
            end;
            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
         end;

         Inc(iHamgmokCnt);
      end
      else if vCod_Hangmok[ihangmok] = 'JJ08' then
      begin
         //���� ��� LSIT...
         for i := 1 to 2 do
         begin
            UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
            UP_Grid_Clear(iHamgmokCnt,'grdResult');

            case i of
               1 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ08';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0181';
                      UV_vTDesc_kor[iHamgmokCnt] := '����(�ְ�)';
                      UV_vGulkwa[iHamgmokCnt]   := eHul_h;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0181''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               2 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ08';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0182';
                      UV_vTDesc_kor[iHamgmokCnt] := '����(����)';
                      UV_vGulkwa[iHamgmokCnt]   := eHul_l;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0182''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
            end;

            Inc(iHamgmokCnt);
         end;
      end
      else if vCod_Hangmok[ihangmok] = 'JJ07' then
      begin
         //���� ��� LSIT...
         for i := 1 to 7 do
         begin
            UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
            UP_Grid_Clear(iHamgmokCnt,'grdResult');

            case i of
               1 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0515';
                      UV_vTDesc_kor[iHamgmokCnt] := 'FEV1';
                      UV_vGulkwa[iHamgmokCnt]   := eFEV1;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0515''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               2 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0516';
                      UV_vTDesc_kor[iHamgmokCnt] := 'FEV1(%)';
                      UV_vGulkwa[iHamgmokCnt]   := eFEV1_p;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0516''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               3 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0517';
                      UV_vTDesc_kor[iHamgmokCnt] := 'FVC';
                      UV_vGulkwa[iHamgmokCnt]   := eFVC;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0517''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               4 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0518';
                      UV_vTDesc_kor[iHamgmokCnt] := 'FVC(%)';
                      UV_vGulkwa[iHamgmokCnt]   := eFVC_p;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0518''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               5 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0519';
                      UV_vTDesc_kor[iHamgmokCnt] := 'FEV/FVC(%)';
                      UV_vGulkwa[iHamgmokCnt]   := eFEVFVC;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0519''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               6 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0521';
                      UV_vTDesc_kor[iHamgmokCnt] := 'PEF';
                      UV_vGulkwa[iHamgmokCnt]   := ePEF;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0521''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               7 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ07';
                      UV_vCod_Hhm[iHamgmokCnt]  := '****';
                      UV_vTDesc_kor[iHamgmokCnt] := 'PEF(%)';
                      UV_vGulkwa[iHamgmokCnt]   := ePEF_p;

                      UV_vIms_low[iHamgmokCnt]   := '';
                      UV_vIms_high[iHamgmokCnt]  := '';
                      UV_vDesc_unit[iHamgmokCnt] := '';
                   end;
            end;

            Inc(iHamgmokCnt);
         end;

         //���� �Ұ�...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJ07';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0437';
         UV_vDesc_kor2[iHamgmokCnt2] := '��Ȱ���˻�';

         //[2017.02.06-������]��������� �׸񿡼� ��������...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0437''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //UV_Gubun2[iHamgmokCnt2]     := '018';
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sPekiSokun;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0437''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
         end;
         
         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJ06' then
      begin
         if pos('JJ76',sHmCode) <= 0 then
         begin
            //1�� û�� �׸� LSIT...
            for i := 1 to 6 do
            begin
               UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
               UP_Grid_Clear(iHamgmokCnt,'grdResult');

               case i of
                  1 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0160';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 2000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0160''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
                  2 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0162';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 2000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0162''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
                  3 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0363';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 3000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0363''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
                  4 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0364';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 3000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0364''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
                  5 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0164';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 4000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0164''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
                  6 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0166';
                         UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 4000Hz(��)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40;

                         qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0166''';
                         if qryTk_Comm12.RecordCount > 0 then
                         begin
                            if iSex = 1 then
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                            end
                            else
                            begin
                               UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                               UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                            end;
                            UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                         end;
                      end;
               end;

               Inc(iHamgmokCnt);
            end;
         end;
      end
      else if vCod_Hangmok[ihangmok] = 'JJ73' then
      begin
         for i := 1 to 2 do
         begin
            case i of
               1 : begin
                      UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
                      UP_Grid_Clear(iHamgmokCnt,'grdResult');

                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ73';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0106';
                      UV_vTDesc_kor[iHamgmokCnt] := '�Ƿ°˻�(��)';

                      //�����..
                      qryJangbi.Filter := 'cod_jangbi = ''' + vCod_Hangmok[ihangmok] + '''';
                      if qryJangbi.RecordCount > 0 then
                         UV_vGulkwa[iHamgmokCnt] := copy(Trim(qryJangbi.FieldByName('desc_sokun').AsString),
                                                    1,
                                                    pos('/',Trim(qryJangbi.FieldByName('desc_sokun').AsString))-1);

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0106''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               2 : begin
                      UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
                      UP_Grid_Clear(iHamgmokCnt,'grdResult');

                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ73';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0107';
                      UV_vTDesc_kor[iHamgmokCnt] := '�Ƿ°˻�(��)';

                      //�����..
                      qryJangbi.Filter := 'cod_jangbi = ''' + vCod_Hangmok[ihangmok] + '''';
                      if qryJangbi.RecordCount > 0 then
                         UV_vGulkwa[iHamgmokCnt] := copy(Trim(qryJangbi.FieldByName('desc_sokun').AsString),
                                                    pos('/',Trim(qryJangbi.FieldByName('desc_sokun').AsString))+1,
                                                    length(Trim(qryJangbi.FieldByName('desc_sokun').AsString)));

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0107''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
            end;
            Inc(iHamgmokCnt);
         end;
      end
      else if vCod_Hangmok[ihangmok] = 'JJ76' then
      begin
         //2�� ����û�� �׸� LSIT...
         for i := 1 to 22 do
         begin
            UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
            UP_Grid_Clear(iHamgmokCnt,'grdResult');

            case i of
               1 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0152';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 500Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT5;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0152''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               2 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0154';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 500Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT5;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0154''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               3 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0156';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 1000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT10;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0156''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               4 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0158';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 1000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT10;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0158''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               5 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0160';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 2000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0160''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               6 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0162';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 2000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0162''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               7 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0363';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 3000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0363''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               8 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0364';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 3000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0364''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
               9 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0164';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 4000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0164''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              10 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0166';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 4000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0166''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              11 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0367';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 6000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT60;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0367''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              12 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0368';
                      UV_vTDesc_kor[iHamgmokCnt] := '�⵵û�� 6000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT60;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0368''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              //----------------------------------------------------------------
              13 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0153';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 500Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT5k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0153''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              14 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0155';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 500Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT5k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0155''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              15 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0157';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 1000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT10k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0157''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              16 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0159';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 1000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT10k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0159''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              17 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0161';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 2000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0161''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              18 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0163';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 2000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0163''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              19 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0365';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 3000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0365''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              20 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0366';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 3000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0366''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              21 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0165';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 4000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0165''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
              22 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0167';
                      UV_vTDesc_kor[iHamgmokCnt] := '������û�� 4000Hz(��)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40k;

                      qryTk_Comm12.Filter := ' �˻��׸��ڵ� = ''0167''';
                      if qryTk_Comm12.RecordCount > 0 then
                      begin
                         if iSex = 1 then
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
                         end
                         else
                         begin
                            UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                            UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
                         end;
                         UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;
                      end;
                   end;
            end;

            Inc(iHamgmokCnt);
         end;
      end
      else if (vCod_Hangmok[ihangmok] = 'JJF8') or (vCod_Hangmok[ihangmok] = 'JJF9') or
              (vCod_Hangmok[ihangmok] = 'JJFA') or (vCod_Hangmok[ihangmok] = 'JJFB') or
              (vCod_Hangmok[ihangmok] = 'JJFC') or (vCod_Hangmok[ihangmok] = 'JJFD') or
              (vCod_Hangmok[ihangmok] = 'JJFE') or (vCod_Hangmok[ihangmok] = 'JJFF') or
              (vCod_Hangmok[ihangmok] = 'JJF3') then
      begin
         //�����ڵ�� ȭ�鿡 Display ����...
      end
      else
      begin
         if vCod_Hangmok[ihangmok] = 'JJB9' then vCod_Hangmok[ihangmok] := 'JJ34';

         qryTk_Comm12.Filter := ' MDMS�׸��ڵ� = ''' + vCod_Hangmok[ihangmok] + '''';
         if qryTk_Comm12.RecordCount > 0 then
         begin
            if qryTk_Comm12.FieldByName('HM_GUM_INFO').AsString = '1' then
            begin
               UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
               UP_Grid_Clear(iHamgmokCnt,'grdResult');

               UV_vCod_Mhm[iHamgmokCnt]  := vCod_Hangmok[ihangmok];
               UV_vCod_Hhm[iHamgmokCnt]  := qryTk_Comm12.FieldByName('�˻��׸��ڵ�').AsString;
               UV_vTDesc_kor[iHamgmokCnt] := qryTk_Comm12.FieldByName('�˻��׸��_��').AsString;
               if iSex = 1 then
               begin
                  UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_LOW').AsString;
                  UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_M_HIGH').AsString;
               end
               else
               begin
                  UV_vIms_low[iHamgmokCnt]  := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_LOW').AsString;
                  UV_vIms_high[iHamgmokCnt] := qryTk_Comm12.FieldByName('HM_GUM_ABN_W_HIGH').AsString;
               end;
               UV_vDesc_unit[iHamgmokCnt]   := qryTk_Comm12.FieldByName('DESC_UNIT').AsString;

               //--------------------------------------------------------------------------
               // ����� �Է�...
               dmComm.qryHangmok.Filter := 'COD_HM = ''' + vCod_Hangmok[ihangmok] + ''' AND ' + 'DAT_APPLY <= ''' + sDat_gmgn + '''';
               if dmComm.qryHangmok.RecordCount > 0 then
               begin
                  if dmComm.qryHangmok.FieldByName('gubn_part').AsString < '10' then
                  begin
                     //���װ��..
                     qryTK_Hul.Filter := 'gubn_part = ''' + dmComm.qryHangmok.FieldByName('gubn_part').AsString + '''';

                     if qryTK_Hul.RecordCount > 0 then
                     begin
                        UV_fGulkwa := '';
                        UV_fGulkwa1 := qryTK_Hul.FieldByName('DESC_GLKWA1').AsString;
                        UV_fGulkwa2 := qryTK_Hul.FieldByName('DESC_GLKWA2').AsString;
                        UV_fGulkwa3 := qryTK_Hul.FieldByName('DESC_GLKWA3').AsString;
                        GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                        UV_vGulkwa[iHamgmokCnt] := Trim(Copy(UV_fGulkwa, dmComm.qryHangmok.FieldByName('pos_start').AsInteger, 6));

                        if Trim(Copy(UV_fGulkwa, dmComm.qryHangmok.FieldByName('pos_start').AsInteger, 6)) = '' then
                        begin
                           if UV_vCod_Mhm[iHamgmokCnt] = 'U003' then UV_vGulkwa[iHamgmokCnt] := eU003;
                           if UV_vCod_Mhm[iHamgmokCnt] = 'U004' then UV_vGulkwa[iHamgmokCnt] := sU004;
                           if UV_vCod_Mhm[iHamgmokCnt] = 'U005' then UV_vGulkwa[iHamgmokCnt] := sU005;
                           if UV_vCod_Mhm[iHamgmokCnt] = 'U009' then UV_vGulkwa[iHamgmokCnt] := sU009;
                        end;

                        //������ı� Check - C026,C028, C032 �˻��ڷ� ����;
                        If (Trim(UV_vGulkwa[iHamgmokCnt]) <> '') And
                           ((vCod_Hangmok[ihangmok] = 'C026') Or
                            (vCod_Hangmok[ihangmok] = 'C028') Or
                            (vCod_Hangmok[ihangmok] = 'C032')) Then
                        Begin
                           If vCod_Hangmok[ihangmok] = 'C026' Then eC026 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C028' Then eC028 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C032' Then eC032 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                        End;

                        //C025,C026, C028 �˻��ڷ� ����;
                        If (Trim(UV_vGulkwa[iHamgmokCnt]) <> '') And
                           ((vCod_Hangmok[ihangmok] = 'C025') Or
                            (vCod_Hangmok[ihangmok] = 'C026') Or
                            (vCod_Hangmok[ihangmok] = 'C028')) Then
                        Begin
                           If vCod_Hangmok[ihangmok] = 'C025' Then eC025 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C026' Then eC026 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C028' Then
                           Begin
                              eC028 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                              //���� �Ҵ�
                              Inc(iHamgmokCnt);
                              UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');

                              UV_vCod_Mhm[iHamgmokCnt]   := 'C027';
                              UV_vCod_Hhm[iHamgmokCnt]   := '****';
                              UV_vTDesc_kor[iHamgmokCnt]  := '**LDL�ݷ����׷�';
                              UV_vIms_low[iHamgmokCnt]   := '50.0';
                              UV_vIms_high[iHamgmokCnt]  := '129.9';
                              UV_vDesc_unit[iHamgmokCnt] := 'mg/dL';
                              UV_vGulkwa[iHamgmokCnt]    := '---';
                              If (eC025 >0) And (eC026 > 0) And (eC028 > 0) Then
                              begin
                                  If eC028 >= 400 Then
                                  begin
                                     UV_vGulkwa[iHamgmokCnt]  := '???';

                                     //[2017.02.09-������]���� ���Ư��������Ʈ1700412
                                     //TG(C028) ���� 400 �� ������ LDL(CO27)�� �����˻��� ��ġ ����...
                                     //�������� ���ϵǾ� ����..
                                     if Trim(Copy(UV_fGulkwa, 133, 6)) <> '' then
                                     begin
                                        Inc(iHamgmokCnt);
                                        UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');

                                        UV_vCod_Mhm[iHamgmokCnt]   := 'C027';
                                        UV_vCod_Hhm[iHamgmokCnt]   := '****';
                                        UV_vTDesc_kor[iHamgmokCnt]  := '**LDL-�ݷ����׷�(����)';
                                        UV_vIms_low[iHamgmokCnt]   := '50.0';
                                        UV_vIms_high[iHamgmokCnt]  := '129.9';
                                        UV_vDesc_unit[iHamgmokCnt] := 'mg/dL';
                                        UV_vGulkwa[iHamgmokCnt]    := Trim(Copy(UV_fGulkwa, 133, 6));
                                     end;
                                  end
                                  Else UV_vGulkwa[iHamgmokCnt] := eC025 - eC026 - (eC028 /5);
                              end;

                              // Hb A1C ��� Ȯ�� �� 2015.04.23 ���� ��ܺ��ǰ���������Ʈ1500252
                              sC083 := Trim(Copy(UV_fGulkwa, 385, 6));
                              If Trim(sC083) <> '' Then
                              Begin
                                 //���� �Ҵ�
                                 Inc(iHamgmokCnt);
                                 UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');

                                 UV_vCod_Mhm[iHamgmokCnt]   := 'C083';
                                 UV_vCod_Hhm[iHamgmokCnt]   := '****';
                                 UV_vTDesc_kor[iHamgmokCnt]  := '**Hb A1C';
                                 UV_vIms_low[iHamgmokCnt]   := '4.4';
                                 UV_vIms_high[iHamgmokCnt]  := '6.4';
                                 UV_vDesc_unit[iHamgmokCnt] := '%';
                                 UV_vGulkwa[iHamgmokCnt]    := sC083;
                              End;
                           End;
                        End;
                     end;
                  end
                  else
                  begin
                     //�����..
                     qryJangbi.Filter := 'cod_jangbi = ''' + dmComm.qryHangmok.FieldByName('cod_hm').AsString + '''';
                     if qryJangbi.RecordCount > 0 then UV_vGulkwa[iHamgmokCnt] := Trim(qryJangbi.FieldByName('desc_sokun').AsString);
                  end;
               end;

               Inc(iHamgmokCnt);
            end
            else
            begin
               // �׿� �˻�....
               UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
               UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

               UV_vCod_Mhm2[iHamgmokCnt2]  := vCod_Hangmok[ihangmok];
               UV_vCod_Hhm2[iHamgmokCnt2]  := qryTk_Comm12.FieldByName('�˻��׸��ڵ�').AsString;
               UV_vDesc_kor2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('�˻��׸��_��').AsString;
               UV_Gubun2[iHamgmokCnt2]     := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;

               //--------------------------------------------------------------------------
               // ����� �Է�...
               dmComm.qryHangmok.Filter := 'COD_HM = ''' + vCod_Hangmok[ihangmok] + ''' AND ' + 'DAT_APPLY <= ''' + sDat_gmgn + '''';
               if dmComm.qryHangmok.RecordCount > 0 then
               begin
                  if dmComm.qryHangmok.FieldByName('gubn_part').AsString < '10' then
                  begin
                     //���װ��..
                     qryTK_Hul.Filter := 'gubn_part = ''' + dmComm.qryHangmok.FieldByName('gubn_part').AsString + '''';

                     if qryTK_Hul.RecordCount > 0 then
                     begin
                        UV_fGulkwa := '';
                        UV_fGulkwa1 := qryTK_Hul.FieldByName('DESC_GLKWA1').AsString;
                        UV_fGulkwa2 := qryTK_Hul.FieldByName('DESC_GLKWA2').AsString;
                        UV_fGulkwa3 := qryTK_Hul.FieldByName('DESC_GLKWA3').AsString;
                        GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                        sTemp := Trim(Copy(UV_fGulkwa, dmComm.qryHangmok.FieldByName('pos_start').AsInteger, 6));
                        if  (UV_vCod_Mhm2[iHamgmokCnt2] = 'U001') and (sTemp <> '')  then
                        begin
                             GF_UrineChar('U001', sTemp, sUrine_Char, sDat_gmgn );
                             UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U004') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U004', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U005') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U005', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U006') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U006', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U007') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U007', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U008') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U008', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else if (UV_vCod_Mhm2[iHamgmokCnt2] = 'U009') and (sTemp <> '')  then
                        begin
                          GF_UrineChar('U009', sTemp, sUrine_Char, sDat_gmgn );
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sUrine_Char;
                        end
                        else
                        begin
                          UV_vCKMSGulkwa2[iHamgmokCnt2] := sTemp;
                        end;

                        if Trim(Copy(UV_fGulkwa, dmComm.qryHangmok.FieldByName('pos_start').AsInteger, 6)) = '' then
                        begin
                           if UV_vCod_Mhm2[iHamgmokCnt2] = 'U003' then UV_vCKMSGulkwa2[iHamgmokCnt2] := eU003;
                           if UV_vCod_Mhm2[iHamgmokCnt2] = 'U004' then UV_vCKMSGulkwa2[iHamgmokCnt2] := sU004;
                           if UV_vCod_Mhm2[iHamgmokCnt2] = 'U005' then UV_vCKMSGulkwa2[iHamgmokCnt2] := sU005;
                           if UV_vCod_Mhm2[iHamgmokCnt2] = 'U009' then UV_vCKMSGulkwa2[iHamgmokCnt2] := sU009;
                        end;
                     end;
                  end
                  else
                  begin
                     //�÷°��..
                     if dmComm.qryHangmok.FieldByName('cod_hm').AsString = 'JJ02' then
                     begin
                        UV_vCKMSGulkwa2[iHamgmokCnt2] := sEye_Left +'/' + sEye_Right;
                     end
                     //�����..
                     Else if dmComm.qryHangmok.FieldByName('cod_hm').AsString = 'JJ34' then
                     begin
                        // �� ���ð��� �ƴ϶� ���� ���ð��� ���
                        qryTK_Jangbi.Filter := 'cod_jangbi = ''' + dmComm.qryHangmok.FieldByName('cod_hm').AsString + '''';
                        if (qryTK_Jangbi.RecordCount > 0) And (qryTK_Jangbi.FieldByName('Cod_sokun').AsString <> '****') then
                           UV_vCKMSGulkwa2[iHamgmokCnt2] := Trim(qryTK_Jangbi.FieldByName('desc_sokun').AsString)
                        else
                        begin
                           qryTK_Jangbi.Filter := 'cod_jangbi = ''JJB9''';
                           if qryTK_Jangbi.RecordCount > 0 then UV_vCKMSGulkwa2[iHamgmokCnt2] := Trim(qryTK_Jangbi.FieldByName('desc_sokun').AsString)
                        end;
                     end
                     else
                     begin
                        qryTK_Jangbi.Filter := 'cod_jangbi = ''' + dmComm.qryHangmok.FieldByName('cod_hm').AsString + '''';
                        if qryTK_Jangbi.RecordCount > 0 then UV_vCKMSGulkwa2[iHamgmokCnt2] := Trim(qryTK_Jangbi.FieldByName('desc_sokun').AsString);
                        if dmComm.qryHangmok.FieldByName('cod_hm').AsString = 'JJ12' then
                        begin
                           sJJD7 := Trim(qryTK_Jangbi.FieldByName('desc_sokun').AsString);
                           UV_vCKMSCODE[iHamgmokCnt2] := Trim(qryTK_Jangbi.FieldByName('Cod_Sokun').AsString);
                           If Trim(qryTK_Jangbi.FieldByName('Cod_Sokun').AsString) = '****' Then
                              UV_vCKMSGulkwa2[iHamgmokCnt2] := '���Կ�';   //2014.05.08 ���� �����Ƿ�����������1400018
                        end;
                     end;
                  end;
               end;

               qryTk_gulkwa.Filter := 'cod_Hhm = ''' + qryTk_Comm12.FieldByName('�˻��׸��ڵ�').AsString + '''';
               if qryTk_gulkwa.RecordCount > 0 then
               begin
                  UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('�˻�����').AsString;
               end;
               Inc(iHamgmokCnt2);
            end;
         end;
      end;

      Inc(ihangmok);
   end;

   grdResult.Rows  := iHamgmokCnt;
   grdResult2.Rows := iHamgmokCnt2;
end;

procedure TfrmCK5I17.UP_CTI_Login(sLogin : string);
var
  sStrid, sStrPwd, sStrExtension : string;
begin
  inherited;
  sStrid  := '';       //���� ID
  sStrPwd := '';       //���� ��й�ȣ
  sStrExtension := ''; //���� ������ȣ

  if sLogin = 'IN' then
  begin
     qrySawon_CTI.Active := False;
     qrySawon_CTI.ParamByName('cod_sawon').AsString := GV_sUserId;
     qrySawon_CTI.Active := True;

     if qrySawon_CTI.RecordCount > 0 then
     begin
       sStrid        := qrySawon_CTI.FieldByName('CTI_ID').AsString;
       sStrPwd       := qrySawon_CTI.FieldByName('CTI_ID').AsString;
       sStrExtension := qrySawon_CTI.FieldByName('CTI_ID').AsString;

       Panel_CTINo.Caption := sStrid;
       CAModule.Login(sStrid,sStrPwd,sStrExtension);
     end
     else
     begin
        showmessage('CTI ����ڰ� �ƴմϴ�.' + #13#10 + '���� ID�� Ȯ�� �ٶ��ϴ�.');
        exit;
     end;
  end
  else if sLogin = 'OUT' then
  begin
     CAModule.Logout();
     Panel_CTINo.Caption := '����';
     Panel_CTI.Caption := '�α׾ƿ�';
     Pnl_CTIUser.Caption := '';
     Pnl_Mod.Caption := '';
     pnl_WaitCnt.Caption := '';
  end;
end;

procedure TfrmCK5I17.MouseWheelHandler(var Message: TMessage);  //����
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //����
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster_Detail.Refresh; //����
      end;
   end;
end;

procedure TfrmCK5I17.Chk_Edit(Sender: TObject);
begin
   // Edit Flag Check
   UV_bEdit := True;
end;

function  TfrmCK5I17.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryNo_hangmok do
   begin
      Active := False;
      ParamByName('dat_apply').AsString   := sDate;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('gubn_yh').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmCK5I17.btnQuitClick(Sender: TObject);
begin
   inherited;

   Close;
end;

procedure TfrmCK5I17.btnExcelClick(Sender: TObject);
begin
   inherited;

   // Excel ��ȯ Dialog Box�� Open�Ѵ�.
   GP_Excel(Self);
end;

procedure TfrmCK5I17.btnOpenOfficeClick(Sender: TObject);
begin
   inherited;

   // Excel ��ȯ Dialog Box�� Open�Ѵ�.
   GP_OpenOffice(Self);
end;

procedure TfrmCK5I17.FormKeyPress(Sender: TObject; var Key: Char);
begin
   inherited;

   // ESC Key�� ������ ��ҹ�ư Click
   if Key = GC_ESC then btnCancel.Click;

end;

procedure TfrmCK5I17.UP_Click(Sender: TObject);
begin
   inherited;

   if Sender = btnNext_gn_date then GF_CalendarClick(next_gn_date);
end;

procedure TfrmCK5I17.UP_VarMemSet_Gmgn(iLength : Integer);
begin
   // Variant Memory Allocation
    if iLength = 0 then
    begin
       // Variant������ ����ϱ� ���ؼ� ����
       UV_vDat_gmgn_G    := VarArrayCreate([0, 0], varOleStr);
       UV_vCod_janbi_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vCod_hul_G     := VarArrayCreate([0, 0], varOleStr);
       UV_vCod_cancer_G  := VarArrayCreate([0, 0], varOleStr);
       UV_vCod_chuga_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vNum_tel_G     := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_nosin_G  := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_nscg_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_adult_G  := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_adcg_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_life_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_lfcg_G   := VarArrayCreate([0, 0], varOleStr);
       UV_vDesc_yhname_G := VarArrayCreate([0, 0], varOleStr);
       UV_vDesc_saler_G  := VarArrayCreate([0, 0], varOleStr);
       UV_vGubn_jinch_G  := VarArrayCreate([0, 0], varOleStr);
       UV_vNum_id_G      := VarArrayCreate([0, 0], varOleStr);
    end
    else
    begin
       // �̹� ������ �������� ũ�⸦ ����
       VarArrayRedim(UV_vDat_gmgn_G,    iLength);
       VarArrayRedim(UV_vCod_janbi_G,   iLength);
       VarArrayRedim(UV_vCod_hul_G,     iLength);
       VarArrayRedim(UV_vCod_cancer_G,  iLength);
       VarArrayRedim(UV_vCod_chuga_G,   iLength);
       VarArrayRedim(UV_vNum_tel_G,     iLength);
       VarArrayRedim(UV_vGubn_nosin_G,  iLength);
       VarArrayRedim(UV_vGubn_nscg_G,   iLength);
       VarArrayRedim(UV_vGubn_adult_G,  iLength);
       VarArrayRedim(UV_vGubn_adcg_G,   iLength);
       VarArrayRedim(UV_vGubn_life_G,   iLength);
       VarArrayRedim(UV_vGubn_lfcg_G,   iLength);
       VarArrayRedim(UV_vDesc_yhname_G, iLength);
       VarArrayRedim(UV_vDesc_saler_G,  iLength);
       VarArrayRedim(UV_vGubn_jinch_G,  iLength);
       VarArrayRedim(UV_vNum_id_G,      iLength);
    end;
end;

procedure TfrmCK5I17.UP_VarMemSet(iLength : Integer; sGubun : String);
begin
   // Variant Memory Allocation
   if sGubun = 'G' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vcod_jisa      := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_name     := VarArrayCreate([0, 0], varOleStr);
         UV_vNum_jumin     := VarArrayCreate([0, 0], varOleStr);
         UV_vNum_id        := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_dc       := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_Dept     := VarArrayCreate([0, 0], varOleStr);
         UV_vConsult_date  := VarArrayCreate([0, 0], varOleStr);
         UV_vConsult_time  := VarArrayCreate([0, 0], varOleStr);
         UV_vCod_sawon     := VarArrayCreate([0, 0], varOleStr);
         UV_vSawon_name    := VarArrayCreate([0, 0], varOleStr);
         UV_vConsult_note  := VarArrayCreate([0, 0], varOleStr);
         UV_vNext_gn_date  := VarArrayCreate([0, 0], varOleStr);
         UV_nH_name        := VarArrayCreate([0, 0], varOleStr);
         UV_nsokun_name    := VarArrayCreate([0, 0], varOleStr);
         UV_vFile_name     := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vcod_jisa,       iLength);
         VarArrayRedim(UV_vDesc_name,      iLength);
         VarArrayRedim(UV_vNum_jumin,      iLength);
         VarArrayRedim(UV_vNum_id,         iLength);
         VarArrayRedim(UV_vDesc_dc,        iLength);
         VarArrayRedim(UV_vDesc_Dept,      iLength);
         VarArrayRedim(UV_vConsult_date,   iLength);
         VarArrayRedim(UV_vConsult_time,   iLength);
         VarArrayRedim(UV_vCod_sawon,      iLength);
         VarArrayRedim(UV_vSawon_name,     iLength);
         VarArrayRedim(UV_vConsult_note,   iLength);
         VarArrayRedim(UV_vNext_gn_date,   iLength);
         VarArrayRedim(UV_nH_name,         iLength);
         VarArrayRedim(UV_nsokun_name,     iLength);
         VarArrayRedim(UV_vFile_name,      iLength);
      end;
   end
   else if sGubun = 'I' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vInjouk_num_id := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vInjouk_num_id,      iLength);
      end;
   end
   else if sGubun = 'F' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vFile_name    := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vFile_name,      iLength);
      end;
   end
   else if sGubun = 'H' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vHul_dat_gmgn    := VarArrayCreate([0, 0], varOleStr);
         UV_vHul_num_id      := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vHul_dat_gmgn,      iLength);
         VarArrayRedim(UV_vHul_num_id,        iLength);
      end;
   end
   else if sGubun = 'HG' then
   begin
      if iLength = 0 then
      begin
         // Variant������ ����ϱ� ���ؼ� ����
         UV_vHul_Gulkwa      := VarArrayCreate([0, 50, 0, 0], varOleStr);
         UV_vTot_cod_hm_name := VarArrayCreate([0, 0, 0, 0], varOleStr);
         UV_vHul_Low_High    := VarArrayCreate([0, 0, 0, 0], varOleStr);
         UV_vFont_color      := VarArrayCreate([0, 50, 0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vHul_Gulkwa,      iLength);
         VarArrayRedim(UV_vTot_cod_hm_name, iLength);
         VarArrayRedim(UV_vHul_Low_High,    iLength);
         VarArrayRedim(UV_vFont_color,      iLength);
      end;
   end
   else if sGubun = 'D' then
   begin
      if iLength = 0 then
      begin
         UV_vIdx           := VarArrayCreate([0, 0], varOleStr);
         UV_vDconsult_date := VarArrayCreate([0, 0], varOleStr);
         UV_vDconsult_time := VarArrayCreate([0, 0], varOleStr);
         UV_vDSawon_name   := VarArrayCreate([0, 0], varOleStr);
         UV_vDConsult_note := VarArrayCreate([0, 0], varOleStr);
         UV_vDNext_gn_date := VarArrayCreate([0, 0], varOleStr);
         UV_vDH_code       := VarArrayCreate([0, 0], varOleStr);
         UV_vDH_name       := VarArrayCreate([0, 0], varOleStr);
         UV_vDsokun_name   := VarArrayCreate([0, 0], varOleStr);
         UV_vDCon_type     := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vIdx,             iLength);
         VarArrayRedim(UV_vDconsult_date,   iLength);
         VarArrayRedim(UV_vDconsult_time,   iLength);
         VarArrayRedim(UV_vDSawon_name,     iLength);
         VarArrayRedim(UV_vDConsult_note,   iLength);
         VarArrayRedim(UV_vDNext_gn_date,   iLength);
         VarArrayRedim(UV_vDH_code,         iLength);
         VarArrayRedim(UV_vDH_name,         iLength);
         VarArrayRedim(UV_vDsokun_name,     iLength);
         VarArrayRedim(UV_vDCon_type,       iLength);
      end;
   end
   else if sGubun = 'K' then
   begin
      if iLength = 0 then
      begin
          UV_vHul_h       := VarArrayCreate([0, 0], varOleStr);
          UV_vHul_l       := VarArrayCreate([0, 0], varOleStr);
          UV_vBiman       := VarArrayCreate([0, 0], varOleStr);
          UV_vSinjang     := VarArrayCreate([0, 0], varOleStr);
          UV_vChejung     := VarArrayCreate([0, 0], varOleStr);
          UV_vear_left5   := VarArrayCreate([0, 0], varOleStr);
          UV_vear_left10  := VarArrayCreate([0, 0], varOleStr);
          UV_vear_lef20   := VarArrayCreate([0, 0], varOleStr);
          UV_vear_left30  := VarArrayCreate([0, 0], varOleStr);
          UV_vear_left40  := VarArrayCreate([0, 0], varOleStr);
          UV_vear_left60  := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right5  := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right10 := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right20 := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right30 := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right40 := VarArrayCreate([0, 0], varOleStr);
          UV_vear_right60 := VarArrayCreate([0, 0], varOleStr);
          UV_vdesc_ear    := VarArrayCreate([0, 0], varOleStr);
          UV_veye_lman    := VarArrayCreate([0, 0], varOleStr);
          UV_veye_rman    := VarArrayCreate([0, 0], varOleStr);
          UV_veye_lkyo    := VarArrayCreate([0, 0], varOleStr);
          UV_veye_rkyo    := VarArrayCreate([0, 0], varOleStr);
          UV_vanap_left   := VarArrayCreate([0, 0], varOleStr);
          UV_vanap_right  := VarArrayCreate([0, 0], varOleStr);
          UV_vgubn_seksin := VarArrayCreate([0, 0], varOleStr);
          UV_vhyungwi     := VarArrayCreate([0, 0], varOleStr);
          UV_vbookwi      := VarArrayCreate([0, 0], varOleStr);
          UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // �̹� ������ �������� ũ�⸦ ����
         VarArrayRedim(UV_vHul_h,         iLength);
         VarArrayRedim(UV_vHul_l,         iLength);
         VarArrayRedim(UV_vBiman,         iLength);
         VarArrayRedim(UV_vSinjang,       iLength);
         VarArrayRedim(UV_vChejung,       iLength);
         VarArrayRedim(UV_vear_left5,     iLength);
         VarArrayRedim(UV_vear_left10,    iLength);
         VarArrayRedim(UV_vear_lef20,     iLength);
         VarArrayRedim(UV_vear_left30,    iLength);
         VarArrayRedim(UV_vear_left40,    iLength);
         VarArrayRedim(UV_vear_left60,    iLength);
         VarArrayRedim(UV_vear_right5,    iLength);
         VarArrayRedim(UV_vear_right10,   iLength);
         VarArrayRedim(UV_vear_right20,   iLength);
         VarArrayRedim(UV_vear_right30,   iLength);
         VarArrayRedim(UV_vear_right40,   iLength);
         VarArrayRedim(UV_vear_right60,   iLength);
         VarArrayRedim(UV_vdesc_ear,      iLength);
         VarArrayRedim(UV_veye_lman,      iLength);
         VarArrayRedim(UV_veye_rman,      iLength);
         VarArrayRedim(UV_veye_lkyo,      iLength);
         VarArrayRedim(UV_veye_rkyo,      iLength);
         VarArrayRedim(UV_vanap_left,     iLength);
         VarArrayRedim(UV_vanap_right,    iLength);
         VarArrayRedim(UV_vgubn_seksin,   iLength);
         VarArrayRedim(UV_vhyungwi,       iLength);
         VarArrayRedim(UV_vbookwi,        iLength);
         VarArrayRedim(UV_vDat_gmgn,      iLength);
      end;
   end
   else if sGubun = 'J' then
   begin
      if iLength = 0 then
      begin
         UV_vDat_gmgn_Jangbi   := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_kor          := VarArrayCreate([0, 0], varOleStr);
         UV_vCod_sokun_Jangbi  := VarArrayCreate([0, 0], varOleStr);
         UV_vGubn_pan          := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_sokun_Jangbi := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_sbsg_Jangbi  := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         VarArrayRedim(UV_vDat_gmgn_Jangbi,   iLength);
         VarArrayRedim(UV_vDesc_kor,          iLength);
         VarArrayRedim(UV_vCod_sokun_Jangbi,  iLength);
         VarArrayRedim(UV_vGubn_pan,          iLength);
         VarArrayRedim(UV_vDesc_sokun_Jangbi, iLength);
         VarArrayRedim(UV_vDesc_sbsg_Jangbi,  iLength);
      end;
   end
   else if sGubun = 'S' then
   begin
      if iLength = 0 then
      begin
         UV_vDat_gmgn_sokun := VarArrayCreate([0, 0], varOleStr);
         UV_vgubn_sokun     := VarArrayCreate([0, 0], varOleStr);
         UV_vCod_sokun      := VarArrayCreate([0, 0], varOleStr);
         UV_vDesc_sokun     := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         VarArrayRedim(UV_vDat_gmgn_sokun, iLength);
         VarArrayRedim(UV_vgubn_sokun,     iLength);
         VarArrayRedim(UV_vCod_sokun,      iLength);
         VarArrayRedim(UV_vDesc_sokun,     iLength);
      end;
   end
   else if sGubun = 'SMS' then
   begin
      if iLength = 0 then
      begin
         UV_vSms_content := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         VarArrayRedim(UV_vSms_content, iLength);
      end;
   end;
end;

procedure TfrmCK5I17.FormCreate(Sender: TObject);
begin
   inherited;

   UP_VarMemSet(0, 'G');
   UP_VarMemSet(0, 'D');
   UP_VarMemSet(0, 'K');

   UP_VarMemSet(0, 'J');
   UP_VarMemSet(0, 'S');
   UP_VarMemSet(0, 'I');

   // Grid�� ȯ�� ����
   grd_Kicho.Rows  := 0;
   grd_Hul.Rows    := 0;
   grd_Jangbi.Rows := 0;
   grd_sokun.Rows  := 0;

   chk_gmgn := '';

   UP_VarMemSet(0, 'SMS');
   UV_CkByte  := False;
   GP_ComboJisa(cmbCOD_JISA);
   GP_ComboDisplay(cmbCOD_JISA, copy(GV_sUserId,1,2), 2);
   GP_ComboSawonAll(cmbCOD_DOCT, 1);
   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);


//   GP_ComboJisa(cmbJisa);
//   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   with qryjc_cohospital_sokun_code do
   begin
      Active := False;
      Active := True;

      if Recordcount > 0 then
      begin
         while Not Eof do
         begin
            Cmb_sokun_name.Items.Add(FieldByName('sokun_name').AsString);
            next;
         end;
      end;
   end;

   pnlYsno_useinfo.Visible := False;
   PnlYsno_Privacy_Forbid.Visible := False;

   // �׸��ڷ� Query
   if not qryHm.Active then qryHm.Active := True;

   Timer1.Enabled := True;

   PageControl1.ActivePage := TabSheet1;

   if qrySokun.Active  = False then qrySokun.Active  := True;
   if qryComm06.Active = False then qryComm06.Active := True;
end;

procedure TfrmCK5I17.UP_DisplayDetail;
var index : Integer;
    sSelect, sWhere, sOrder : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrder := '';
   index   := 0;
   // Grid �ʱ�ȭ
   grdMaster_Detail.Rows := 0;

   with qryJc_consults_list do
   begin
      Active := False;

      sSelect :=  ' select J.*, I.num_id as Injouk_num_id, I.num_jumin, I.desc_name, I.desc_dept, D.desc_dc, S.desc_name As Sawon_name, S.cod_jisa as Sawon_cod_jisa, C.sokun_name from jc_consults J left outer join injouk I ';
      sSelect :=  sSelect + ' On J.num_id = I.num_id ';
      sSelect :=  sSelect + ' inner join Danche D On I.cod_dc = D.cod_dc ';
      sSelect :=  sSelect + ' inner join Sawon S On J.cod_sawon = S.cod_sawon ';
      sSelect :=  sSelect + ' left outer join jc_cohospital_sokun K On J.idx = K.idx ';
      sSelect :=  sSelect + ' left outer join jc_cohospital_sokun_code C On K.sokun_code = C.sokun_code ';
      sWhere  := ' WHERE  I.num_jumin = ''' + edtJumin.Text + ''' ';

      sOrder :=  ' order by J.consult_date Desc , J.consult_time Desc';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      if qryJc_consults_list.IsEmpty = False then
      begin
         index := 0;
         while Not Eof do
         begin
            UP_VarMemSet(index, 'D');
            UV_vIdx[index]           := FieldByName('idx').AsString;
            UV_vDconsult_date[index] := FieldByName('consult_date').AsString;
            UV_vDconsult_time[index] := FieldByName('consult_time').AsString;
            UV_vDSawon_name[index]   := FieldByName('Sawon_name').AsString;
            UV_vDCon_type[index]     := FieldByName('con_type').AsString;
            if      FieldByName('Sawon_cod_jisa').AsString = '11' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (��������)'
            else if FieldByName('Sawon_cod_jisa').AsString = '12' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (���ǵ�����)'
            else if FieldByName('Sawon_cod_jisa').AsString = '15' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (��������)'
            else if FieldByName('Sawon_cod_jisa').AsString = '43' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (��������)'
            else if FieldByName('Sawon_cod_jisa').AsString = '51' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (���ּ���)'
            else if FieldByName('Sawon_cod_jisa').AsString = '61' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (�λ꼾��)'
            else if FieldByName('Sawon_cod_jisa').AsString = '71' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (�뱸����)';

            UV_vDConsult_note[index] := FieldByName('Consult_note').AsString;
            UV_vDNext_gn_date[index] := FieldByName('Next_gn_date').AsString;
            UV_vDsokun_name[index]   := FieldByName('sokun_name').AsString;
            if FieldByName('H_code').AsString <> '' then
            begin
               with qryJc_cohospital do
               begin
                  Active := False;
                  ParamByName('h_code').AsString := qryJc_consults_list.FieldByName('h_code').AsString;
                  Active := True;
                  if Recordcount > 0 then
                  begin
                     UV_vDH_code[index] := FieldByName('h_code').AsString;
                     UV_vDH_name[index] := FieldByName('h_name').AsString;
                  end;
               end;
            end;

            Next;
            Inc(index);
         end;

      end;
      // Table�� Disconnect
      Active := False;
   end;

   // Grid�� �ڷḦ �Ҵ�
   if index > 0 then
      grdMaster_Detail.Rows := index;

end;

procedure TfrmCK5I17.grdMaster_DetailCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := intToStr(DataRow);
       2 : Value := UV_vDconsult_date[DataRow-1];
       3 : Value := UV_vDconsult_time[DataRow-1];
       4 : Value := UV_vDSawon_name[DataRow-1];
       5 : Value := UV_vDConsult_note[DataRow-1];
       6 : Value := UV_vDNext_gn_date[DataRow-1];
       7 : Value := UV_vDH_name[DataRow-1];
       8 : Value := UV_vDsokun_name[DataRow-1];
   end;

   if trim(UV_vDCon_type[DataRow - 1]) = 'C' then
   begin
        with grdMaster_Detail do
        begin
            AssignCellFont(DataCol, DataRow);
            CellFont[DataCol, DataRow].Color := clMaroon;
            CellFont[DataCol, DataRow].Style := [fsBold];
        end;
   end;

end;

procedure TfrmCK5I17.grdMaster_DetailRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grdMaster_Detail.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grdMaster_Detail.Rows Do
      begin
         grdMaster_Detail.CellColor[1, i]  := clWindow;
         grdMaster_Detail.CellColor[2, i]  := clWindow;
         grdMaster_Detail.CellColor[3, i]  := clWindow;
         grdMaster_Detail.CellColor[4, i]  := clWindow;
         grdMaster_Detail.CellColor[5, i]  := clWindow;
         grdMaster_Detail.CellColor[6, i]  := clWindow;
         grdMaster_Detail.CellColor[7, i]  := clWindow;
         grdMaster_Detail.CellColor[8, i]  := clWindow;
         Next;
      end;
   end;

   if (grdMaster_Detail.Rows > 0) and (NewRow > 0) then
   begin
      grdMaster_Detail.CellColor[1, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[2, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[3, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[4, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[5, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[6, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[7, NewRow]  := $001AD7C9;
      grdMaster_Detail.CellColor[8, NewRow]  := $001AD7C9;
   end;

   consult_date.Text := ''; consult_time.Text := ''; h_code.Text := ''; h_name.Text := ''; next_gn_date.Text := '';
   con_type.ItemIndex := 0; con_kind.ItemIndex := 0;
   Cmb_sokun_name.Text := '';
   Chk_Recall.Checked  := False;

   with qryJc_consults do
   begin
      Active := False;

      sSelect :=  ' select J.*, I.num_jumin, I.desc_name, I.desc_dept, D.desc_dc, S.desc_name As Sawon_name, C.sokun_name from jc_consults J inner join injouk I ';
      sSelect :=  sSelect + ' On J.num_id = I.num_id ';
      sSelect :=  sSelect + ' inner join Danche D On I.cod_dc = D.cod_dc ';
      sSelect :=  sSelect + ' inner join Sawon S On J.cod_sawon = S.cod_sawon ';
      sSelect :=  sSelect + ' left outer join jc_cohospital_sokun K On J.idx = K.idx ';
      sSelect :=  sSelect + ' left outer join jc_cohospital_sokun_code C On K.sokun_code = C.sokun_code ';
      sWhere  := ' WHERE  I.num_jumin = ''' + edtJumin.Text + ''' ';
      sWhere  := sWhere + ' And  J.consult_date = ''' + UV_vDconsult_date[grdMaster_Detail.CurrentDataRow-1] + ''' ';
      sWhere  := sWhere + ' And  J.consult_time = ''' + UV_vDconsult_time[grdMaster_Detail.CurrentDataRow-1] + ''' ';
      sOrder  :=  ' order by J.consult_date Desc , J.consult_time Desc';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      if qryJc_consults.IsEmpty = False then
      begin
         while Not Eof do
         begin
            consult_date.Text := qryJc_consults.FieldByName('consult_date').AsString;
            consult_time.Text := qryJc_consults.FieldByName('consult_time').AsString;
            consult_note.Text := qryJc_consults.FieldByName('Consult_note').AsString;
            next_gn_date.Text := Copy(qryJc_consults.FieldByName('Next_gn_date').AsString,1,4) + Copy(qryJc_consults.FieldByName('Next_gn_date').AsString,6,2) + Copy(qryJc_consults.FieldByName('Next_gn_date').AsString,9,2);
            //�Ұ�
            Cmb_sokun_name.Text := qryJc_consults.FieldByName('sokun_name').AsString;

            //������
            if trim(FieldByName('con_type').AsString) = 'N' then
               con_type.ItemIndex := 0
            else if trim(FieldByName('con_type').AsString) = 'G' then
               con_type.ItemIndex := 1
            else if trim(FieldByName('con_type').AsString) = 'C' then
               con_type.ItemIndex := 2
            else if trim(FieldByName('con_type').AsString) = 'E' then
               con_type.ItemIndex := 3
            else con_type.ItemIndex := -1;

            //������
            if trim(FieldByName('Gubn_Recall').AsString) = 'Y' then
               chk_ReCall.Checked := True
            Else
               chk_ReCall.Checked := False;

            //�������
            if trim(FieldByName('con_kind').AsString) = '1' then
               con_kind.ItemIndex := 0
            else if trim(FieldByName('con_kind').AsString) = '2' then
               con_kind.ItemIndex := 1
            else if trim(FieldByName('con_kind').AsString) = '3' then
               con_kind.ItemIndex := 2
            else if trim(FieldByName('con_kind').AsString) = '4' then
               con_kind.ItemIndex := 3
            else if trim(FieldByName('con_kind').AsString) = '5' then
               con_kind.ItemIndex := 4
            else con_kind.ItemIndex := -1;


            if FieldByName('H_code').AsString <> '' then
            begin
               with qryJc_cohospital do
               begin
                  Active := False;
                  ParamByName('h_code').AsString := qryJc_consults.FieldByName('h_code').AsString;
                  Active := True;
                  if Recordcount > 0 then
                  begin
                     h_code.Text  := FieldByName('h_code').AsString;
                     h_name.Text := FieldByName('h_name').AsString;
                  end;
               end;
            end;

            Next;
         end;
      end;
      // Table�� Disconnect
      Active := False;
   end;
end;

procedure TfrmCK5I17.btn_NewClick(Sender: TObject);
begin
  inherited;
   consult_date.Text := ''; consult_time.Text := ''; h_code.Text := ''; h_name.Text := ''; next_gn_date.Text := '';
   con_type.ItemIndex := 0; con_kind.ItemIndex := 0;
   Chk_ReCall.Checked := False;
   consult_note.Text := '';
   Cmb_sokun_name.Text := '';
   Cmb_sokun_name.Enabled := False;
   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmCK5I17.btn_saveClick(Sender: TObject);
var sSokun_code : String;
begin
  inherited;
   // Confirm message
   if not GF_MsgBox('Confirmation', '���� �Ͻðڽ��ϱ�.?') then exit;
   if consult_date.Text = '' then
   begin
      with qryI_Jc_consults do
      begin
         ParamByName('cod_jisa').AsString     := Copy(GV_sUserId,1,2);
         ParamByName('num_id').AsString       := UV_sNum_id;
         ParamByName('consult_note').AsMemo   := consult_note.Text;
         ParamByName('cod_sawon').AsString    := GV_sUserId;
         ParamByName('consult_date').AsString := Copy(GV_sToday,1,4) + '-' + Copy(GV_sToday,5,2) + '-' + Copy(GV_sToday,7,2);

         if next_gn_date.Text <> '' then
            ParamByName('next_gn_date').AsString := Copy(next_gn_date.Text,1,4) + '-' + Copy(next_gn_date.Text,5,2) + '-' + Copy(next_gn_date.Text,7,2)
         else ParamByName('next_gn_date').AsString := '';

         ParamByName('isValidate').AsString   := 'y';
         ParamByName('consult_type').AsString := 'T';
         ParamByName('h_code').AsString       := h_code.Text;
         ParamByName('con_type').AsString     := Copy(con_type.Text,1,1);
         If Chk_ReCall.Checked Then
            ParamByName('Gubn_ReCall').AsString := 'Y'
         Else
            ParamByName('Gubn_ReCall').AsString := 'N';
         ParamByName('con_kind').AsString     := Copy(con_kind.Text,1,1);
         GP_query_log(qryI_Jc_consults, 'CK5I17', 'I', 'Y');
         Execsql;
      end;

      with qryJc_consults_idx do
      begin
         Close;
         ParamByName('num_id').AsString := UV_sNum_id;
         Open;
      end;

      if Cmb_sokun_name.Text <> '' then
      begin
         with qryhospital_Sokun_code do
         begin
            Close;
            ParamByName('sokun_name').AsString := trim(Cmb_sokun_name.Text);
            Open;
            sSokun_code := qryhospital_Sokun_code.FieldByName('sokun_code').AsString;
         end;

         with qryI_jc_cohospital_sokun do
         begin
            ParamByName('idx').AsString   := qryJc_consults_idx.FieldByName('idx').AsString;

            ParamByName('Sokun_code').AsString := sSokun_code;
            GP_query_log(qryI_jc_cohospital_sokun, 'CK5I17', 'I', 'Y');
            Execsql;
         end;
      end
      else if Cmb_sokun_name.Text = '' then
      begin
         with qryD_jc_cohospital_sokun do
         begin
            ParamByName('idx').AsString   := qryJc_consults_idx.FieldByName('idx').AsString;
            GP_query_log(qryD_jc_cohospital_sokun, 'CK5I17', 'D', 'Y');
            Execsql;
         end;
      end;
   end
   else if consult_date.Text <> '' then
   begin
      with qryU_Jc_consults do
      begin
         ParamByName('num_id').AsString       := UV_sNum_id;
         ParamByName('consult_date').AsString := consult_date.Text;
         ParamByName('consult_time').AsString := consult_time.Text;

         ParamByName('consult_note').AsMemo   := consult_note.Text;
         if next_gn_date.Text <> '' then
            ParamByName('next_gn_date').AsString := Copy(next_gn_date.Text,1,4) + '-' + Copy(next_gn_date.Text,5,2) + '-' + Copy(next_gn_date.Text,7,2)
         else ParamByName('next_gn_date').AsString := '';
         ParamByName('h_code').AsString       := h_code.Text;
         ParamByName('con_type').AsString     := Copy(con_type.Text,1,1);
         If Chk_ReCall.Checked Then
            ParamByName('Gubn_ReCall').AsString := 'Y'
         Else
            ParamByName('Gubn_ReCall').AsString := 'N';
         ParamByName('con_kind').AsString     := Copy(con_kind.Text,1,1);
         GP_query_log(qryU_Jc_consults, 'CK5I17', 'U', 'Y');
         Execsql;
      end;

      if Cmb_sokun_name.Text <> '' then
      begin
         with qryhospital_Sokun_code do
         begin
            Close;
            ParamByName('sokun_name').AsString := trim(Cmb_sokun_name.Text);
            open;
            sSokun_code := qryhospital_Sokun_code.FieldByName('sokun_code').AsString;
         end;
         with qryD_jc_cohospital_sokun do
         begin
            ParamByName('idx').AsString        := UV_vIdx[grdMaster_Detail.CurrentDataRow-1];
            GP_query_log(qryD_jc_cohospital_sokun, 'CK5I17', 'D', 'Y');
            Execsql;
         end;         
         with qryI_jc_cohospital_sokun do
         begin
            ParamByName('idx').AsString        := UV_vIdx[grdMaster_Detail.CurrentDataRow-1];
            ParamByName('Sokun_code').AsString := sSokun_code;
            GP_query_log(qryI_jc_cohospital_sokun, 'CK5I17', 'I', 'Y');
            Execsql;
         end;
      end
      else if Cmb_sokun_name.Text = '' then
      begin
         with qryD_jc_cohospital_sokun do
         begin
            ParamByName('idx').AsString        := UV_vIdx[grdMaster_Detail.CurrentDataRow-1];
            GP_query_log(qryD_jc_cohospital_sokun, 'CK5I17', 'D', 'Y');
            Execsql;
         end;
      end;
   end;

   UV_bEdit := False;
   //edtJuminChange(edtJumin);
   grdMaster_Detail.repaint;
   UP_DisplayDetail;
end;

procedure TfrmCK5I17.btn_deleteClick(Sender: TObject);
begin
  inherited;
   // Confirm message
   if not GF_MsgBox('Confirmation', '���� �Ͻðڽ��ϱ�.?') then exit;
   with qryD_Jc_consults do
   begin
      ParamByName('num_id').AsString       := UV_sNum_id;
      ParamByName('consult_date').AsString := consult_date.Text;
      ParamByName('consult_time').AsString := consult_time.Text;
      GP_query_log(qryD_Jc_consults, 'CK5I17', 'D', 'Y');
      Execsql;
   end;
   grdMaster_Detail.Repaint;
   UP_DisplayDetail;
end;

procedure TfrmCK5I17.Ck_hospitalClick(Sender: TObject);
begin
  inherited;
   UV_sHCode_Check := 'insert';
   if Ck_hospital.Checked then
   begin
      frmCK5I17A := TFrmCK5I17A.Create(self);
      frmCK5I17A.ShowModal;
      Cmb_sokun_name.Enabled := True;
   end
   else if Ck_hospital.Checked = False then
   begin
      h_code.Text := '';
      h_name.Text := '';
      Cmb_sokun_name.Enabled := False;
   end;
end;

procedure TfrmCK5I17.edtJuminChange(Sender: TObject);
var index, index2 : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   UV_sNum_id := '';

   if length(trim(edtJumin.Text)) = 13 then
   begin
      Btn_Complain.Enabled := True;

      // Grid �ʱ�ȭ
      grdMaster_Detail.Rows := 0;
      grd_Kicho.Rows        := 0;
      grd_Hul.Rows          := 0;
      grd_Jangbi.Rows       := 0;
      grd_sokun.Rows        := 0;

      grd_file.Rows := 0;
      index  := 0;
      index2 := 0;
      index3 := 0;
      consult_date.Text := '';
      consult_time.Text := '';
      Cmb_sokun_name.Text := '';
      Ck_hospital.Checked := False;
      h_code.Text := '';
      h_name.Text := '';
      consult_note.Text := '';
      Chk_Recall.Checked := False;
//      con_type.Text := '';
      next_gn_date.Text := '';
//      con_kind.Text := '';
      bCK_Num_id := False;
      Num_tel.Text  := '';
      cmb_gmgn.items.Clear;
      TabSheet1.Show;

      With  qryGumgin  do
      begin
         Active := False;
         sSelect :=  ' select J.*, I.num_id as Injouk_num_id, I.num_jumin, I.desc_name, I.desc_dept, I.num_tel, I.desc_addr, I.Ysno_Privacy_Forbid, D.desc_dc, S.desc_name As Sawon_name, C.sokun_name from injouk I left outer join jc_consults J ';
         sSelect :=  sSelect + ' On J.num_id = I.num_id ';
         sSelect :=  sSelect + ' inner join Danche D On I.cod_dc = D.cod_dc ';
         sSelect :=  sSelect + ' left outer join Sawon S On J.cod_sawon = S.cod_sawon ';
         sSelect :=  sSelect + ' left outer join jc_cohospital_sokun K On J.idx = K.idx ';
         sSelect :=  sSelect + ' left outer join jc_cohospital_sokun_code C On K.sokun_code = C.sokun_code ';

         sWhere := 'WHERE ';
      
         if edtJumin.Text <> '' then
         begin
            sWhere := sWhere + ' I.num_jumin = ''' + edtJumin.Text + ''' ';
            sOrder := ' order by J.consult_date, J.consult_time ';
         end;

         if edtJumin.Text = '' then
            sOrder := ' order by J.consult_date, J.consult_time ';
      
         SQL.Clear;
         SQL.Add(sSelect + sWhere + sOrder );
         Active := True;

         if qryGumgin.IsEmpty = False then
         begin
            //query_log�����...
            GP_query_log(qryGumgin,'CK5I17', 'Q', 'Y');

            //[2018.08.09-������]���� �λ�CRM����1800029 �ǰ�]
            //Ÿ�ο��� �������� ���� ����(������������ ���� ��û��)
            //------------------------------------------------------------
            if FieldByName('Ysno_Privacy_Forbid').AsString = 'Y' then
            begin
               PnlYsno_Privacy_Forbid.Visible := True;
               GF_MsgBox('Warning', '[�ڰ����������� ��������]' + #13#10#13#10 +
                                    'Ÿ�ο��� �������� ���� ���� ��û�� �����Դϴ�.' + #13#10 +
                                    '(���� �� ����Ȯ�κҰ� �� ��� ���� �߼�(����) ����)');
            end
            else
            begin
               PnlYsno_Privacy_Forbid.Visible := False;
            end;
            //==================================================================

            // DataSet�� ���� Variant������ �̵�
            while Not Eof do
            begin
               UP_VarMemSet(index, 'G');
               UV_vDesc_name[index]    := FieldByName('Desc_name').AsString;
               UV_vNum_jumin[index]    := FieldByName('num_jumin').AsString;
               UV_vNum_id[index]       := FieldByName('Injouk_num_id').AsString;
               UV_sNum_id              := FieldByName('Injouk_num_id').AsString;
               UV_vDesc_dc[index]      := FieldByName('desc_dc').AsString;
               UV_vDesc_Dept[index]    := FieldByName('desc_dept').AsString;
               UV_vConsult_date[index] := FieldByName('Consult_date').AsString;
               UV_vSawon_name[index]   := FieldByName('Sawon_name').AsString;
               UV_vConsult_note[index] := FieldByName('Consult_note').AsString;
               UV_nsokun_name[index]   := FieldByName('sokun_name').AsString;

               //�������� ǥ��..
               with qryGmgn2 do
               begin
                  Active := False;
                  ParamByName('num_jumin').AsString  := qryGumgin.FieldByName('num_jumin').AsString;
                  Active := True;
                  GP_query_log(qryGmgn2, 'CK5I17', 'Q', 'Y');

                  edtName.Text            := FieldByName('Desc_name').AsString;
                  Edt_Company.Text        := FieldByName('desc_dc').AsString;
                  UV_sCod_dc              := FieldByName('cod_dc').AsString;
                  Edt_Dept.Text           := FieldByName('desc_dept').AsString;
                  Edt_Tel.Text            := FieldByName('num_tel').AsString;
                  Edt_Address.Text        := FieldByName('desc_addr').AsString;
                  Num_tel.Text            := FieldByName('num_tel').AsString;
               end;

               if FieldByName('H_code').AsString <> '' then
               begin
                  with qryJc_cohospital do
                  begin
                     Active := False;
                     ParamByName('h_code').AsString := qryGumgin.FieldByName('h_code').AsString;
                     Active := True;
                     if Recordcount > 0 then
                     begin
                        UV_nH_name[index] := FieldByName('h_name').AsString;
                     end;
                  end;
               end;
               Next;
               Inc(index);
            end;
         end
         else
         begin
            // �ڷᰡ ������ ǥ��
            GF_MsgBox('Information', 'NODATA');
            exit;
         end;
         Active := False;
      end;

      with qryFile_directory do
      begin
         qryFile_directory.Active := false;
         qryFile_directory.ParamByName('num_jumin').AsString      := edtJumin.Text;
         qryFile_directory.ParamByName('gubn_use').AsString       := 'Hospital';
         qryFile_directory.ParamByName('file_directory').AsString := '\\222.222.1.6\hospital\';
         qryFile_directory.Active := True;
         if qryFile_directory.RecordCount > 0 then
         begin
            while Not Eof do
            begin
               UP_VarMemSet(index2, 'F');
               UV_vFile_name[index2] := FieldByName('file_name').AsString;
               Next;
               Inc(index2);
            end;
         end;
         // Grid�� �ڷḦ �Ҵ�
         if index2 > 0 then
            grd_file.Rows := index2;
      end;

      with qryInjouk do
      begin
         qryInjouk.Active := false;
         qryInjouk.ParamByName('num_jumin').AsString := edtJumin.Text;
         qryInjouk.Active := True;
         if qryInjouk.RecordCount > 0 then
         begin
            if qryInjouk.RecordCount > 1 then bCK_Num_id := True    //ID �������� ��� True...
            else                              bCK_Num_id := False;

            while Not Eof do
            begin
               UP_VarMemSet(index3, 'I');
               UV_vInjouk_num_id[index3] := FieldByName('num_id').AsString;


               with qryGmgn3 do //2016.02.03 - �ڴ�� - �������̺��� �����������ǿ��� �������̺� �����������Ƿ� ����
               begin
                  Active := False;
                  ParamByName('num_jumin').AsString  := edtJumin.Text;
                  Active := True;

                  if FieldByName('ysno_useinfo').AsString = 'Y' then pnlYsno_useinfo.Visible := False //�������� Ȱ�� ����
                  else                                               pnlYsno_useinfo.Visible := True;

                  with qryKicho3 do //2016.02.03 - �ڴ�� - �������̺��� ���� �ֱ��� VIP��� ���
                  begin
                       Active := False;
                       ParamByName('num_id').AsString  := qrygmgn3.FieldByName('num_id').AsString;
                       Active := True;

                       if FieldByName('ysno_label').AsString = 'Y' then Ysno_Label.Visible := True       //VIP
                       else                                             Ysno_Label.Visible := False;

                       Active := False;
                  end;
                  Active := False;
               end;



               Next;
               Inc(index3);
            end;
         end;
      end;

      UP_DisplayDetail;

      CreateTextArray_Gmgn;
      CreateTextArray_Hul;
      cmb_gmgn.ItemIndex :=0;

      grd_file.Repaint;
      grd_Hul.Repaint;

      //Sheet ���� �� �� ���.. ó������ ǥ�� �ʿ� ����...
      //CreateTextArray;
      //CreateTextArray_Jangbi;
      //CreateTextArray_Sokun;
      //CreateTextArray_cell;
      //CreateTextArrayTot_Munjin;
      //CreateTextArray_Pan;

      //grd_Kicho.Repaint;
      //grd_Jangbi.Repaint;
      //grd_sokun.Repaint;

      Gmgn_retest;
      UP_Change(cmb_gmgn);
      
      //[2017.08.17-������]���� �λ�CRM����1700014�ǰ� �۾�
      //������ ��㳻���� �ƴ� �ű�â ��û...
      btn_NewClick(Self);
   end;
end;

procedure TfrmCK5I17.grd_fileCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vFile_name[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grd_fileButtonClick(Sender: TObject; DataCol,
  DataRow: Integer);
var j : Integer;
    sFull_FileName : String;

    AppHandle: HWND;
begin
  inherited;
   j := DataRow - 1;

   sFull_FileName := '\\222.222.1.6\hospital\' + UV_vFile_name[j];
   if DirectoryExists('\\222.222.1.6\hospital\') then
   begin
      AppHandle := ShellExecute(Handle,'open',PChar(sFull_FileName),'','',SW_SHOWNORMAL);
   end
   else ShowMessage('\\222.222.1.6\chart\ ��Ʈ��ũ ����̺� ������ �Ǿ����� �ʽ��ϴ�. ' + #13#10 +
                    '��Ʈ��ũ ����̺� ������ ���� Ȯ���� �����մϴ�.');
end;

procedure TfrmCK5I17.btn_searchClick(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
  if GF_InjoukClick(Self, sCode, sName) then
  begin
     edtJumin.Text := sCode;
     edtName.Text  := sName;
  end;
end;

procedure TfrmCK5I17.edtNameChange(Sender: TObject);
begin
  inherited;
  if edtName.Text = '' then
     edtJumin.Text := '';
end;

procedure TfrmCK5I17.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if (ssAlt in Shift) and (chr(Key) in ['N', 'n']) then
      btn_NewClick(btn_New);
   if (ssAlt in Shift) and (chr(Key) in ['S', 's']) then
      btn_saveClick(btn_save);
   if (ssAlt in Shift) and (chr(Key) in ['D', 'd']) then
      btn_deleteClick(btn_delete);
end;

procedure TfrmCK5I17.PageControl1Change(Sender: TObject);
begin
  inherited;
  if PageControl1.ActivePage = TabSheet2 then
  begin
      CreateTextArray;
      grd_Kicho.Repaint;
  end
  else if PageControl1.ActivePage = TabSheet3 then
  begin
      CreateTextArray_Hul;
      grd_Hul.Repaint;
  end
  else if PageControl1.ActivePage = TabSheet4 then
  begin
      CreateTextArray_Jangbi;
      grd_Jangbi.Repaint;
  end
  else if PageControl1.ActivePage = TabSheet5 then
  begin
      CreateTextArray_Sokun;
      grd_sokun.Repaint;
  end
  else if PageControl1.ActivePage = TabSheet6 then
  begin
     Num_tel.Text := Edt_Tel.Text;
  end
  else if (PageControl1.ActivePage = TabSheet7)and (cmb_gmgn.Text <>'') then
  begin
     CreateTextArray_cell;
  end
  else if ((PageControl1.ActivePage = TabSheet8) or (PageControl1.ActivePage = TabSheet12))and (cmb_gmgn.Text <>'') then
  begin
     if copy(cmb_gmgn.text,1,4) < '2018' then
     begin
        CreateTextArrayTot_Munjin;
        TabSheet8.TabVisible := True;
        TabSheet12.TabVisible := False;
     end
     else
     begin
        CreateTextArrayTot_Munjin2018;  
        TabSheet8.TabVisible := False;
        TabSheet12.TabVisible := True;
     end;
  end
  else if (PageControl1.ActivePage = TabSheet9)and (cmb_gmgn.Text <>'') and (chk_gmgn <> cmb_gmgn.Text) then
  begin
     CreateTextArray_Pan;
  end
  else if (PageControl1.ActivePage = TabSheet10)and (cmb_gmgn.Text <>'') then
  begin
     CreateTextArrayTkgum;
  end
end;

function TfrmCK5I17.UF_Sokun_cnvt(sCod_sokun : string): String;
begin
   Result := '';
   qrySokun.Filter := ' cod_sokun = ''' + sCod_sokun + '''';

   if qrySokun.RecordCount > 0 then Result := qrySokun.FieldByName('Desc_sokun').AsString;
end;

function TfrmCK5I17.UF_Hmatter_cnvt(sCod_Hmatter : string): String;
begin
   Result := '';
   qryComm06.Filter := ' �����ڵ� = ''' + sCod_Hmatter + '''';

   if qryComm06.RecordCount > 0 then Result := qryComm06.FieldByName('���ع�����').AsString;
end;

procedure TfrmCK5I17.UP_GumsaCheck(gubun, sID_1, sJisa_1, sNsGubn_1, sNsCg_1, sDate1, sDate2  : String);
var
   sSelect, sWhere, sOrderBy,
   sTemp, sTemp_NS, sTemp_Tk, sTk_YHGM, sTkCg : String;

begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sTemp := ''; sTemp_NS := ''; sTemp_Tk  := '';  sTk_YHGM:= '';{��������} sTkCg := '';{Ư��û������}

   if gubun = '1' then
   begin
      {���}
      qryTemp.Active := False;

      sSelect := ' select N.*, T.medi_care1, T.medi_care2, T.medi_care3, T.medi_care4 from Ns_sokun N left outer join Tot_Munjin2010 T';
      sSelect := sSelect + ' On N.cod_jisa = T.cod_jisa and N.num_id = T.num_id and N.dat_gmgn = T.dat_gmgn ';
      sWhere  := ' where N.cod_jisa = ''' + sJisa_1 + '''';
      sWhere  := sWhere + ' and N.num_id   = ''' + sID_1   + '''';
      sWhere  := sWhere + ' and N.dat_gmgn = ''' + sDate1  + '''';

      qryTemp.SQL.Clear;
      qryTemp.SQL.Add(sSelect + sWhere + sOrderBy);
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then
      begin
         if (qryTemp.FieldByName('cod_panjg1').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun1').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun1').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun1').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg2').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun2').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun2').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun2').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg3').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun3').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun3').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun3').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg4').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun4').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun4').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun4').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg5').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun5').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun5').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun5').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg6').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun6').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun6').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun6').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg7').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun7').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun7').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun7').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg8').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun8').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun8').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun8').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg9').AsString  = 'R2') and
            (((qryTemp.FieldByName('cod_sokun9').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun9').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun9').AsString) + ',';
         if (qryTemp.FieldByName('cod_panjg10').AsString = 'R2') and
            (((qryTemp.FieldByName('cod_sokun10').AsString  = 'F002') and (qryTemp.FieldByName('medi_care1').AsString <> 'Y'))
          or ((qryTemp.FieldByName('cod_sokun10').AsString  = 'C001') and (qryTemp.FieldByName('medi_care2').AsString <> 'Y'))) then
            sTemp_NS := sTemp_NS + UF_Sokun_cnvt(qryTemp.FieldByName('cod_sokun10').AsString) + ',';
      end;

      qryTemp.Active := False;
   end
   else if gubun = '2' then
   begin
      {Ư��}
      //[20160905-������]Ư�� �������� ã��...
      //------------------------------------------------------------------------
      //[20170419-������]1�� ���� û�� ���� ��������...
      sSelect := ' select G.gubn_tkcg, T.* from tkgum T with (Nolock) left outer join gumgin G with (Nolock) ON T.cod_jisa = G.cod_jisa AND T.num_jumin = G.num_jumin AND T.dat_gmgn = G.dat_gmgn AND T.gubn_cha = G.gubn_tkgm ';
      sWhere  := ' where T.cod_jisa = ''' + sJisa_1 + '''';
      sWhere  := sWhere + ' and T.num_jumin = ''' + edtJumin.Text + '''';
      sWhere  := sWhere + ' and T.dat_gmgn = ''' + sDate1 + '''';
      sWhere  := sWhere + ' and T.gubn_cha = ''1''';
{
      sSelect := ' select * from tkgum with (Nolock) ';
      sWhere  := ' where cod_jisa = ''' + sJisa_1 + '''';
      sWhere  := sWhere + ' and num_jumin = ''' + mskNUM_JUMIN.Text + '''';
      sWhere  := sWhere + ' and dat_gmgn = ''' + sDate1 + '''';
      sWhere  := sWhere + ' and gubn_cha = ''1''';
}
      qryTemp.Active := False;
      qryTemp.SQL.Clear;
      qryTemp.SQL.Add(sSelect + sWhere + sOrderBy);
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then
      begin
         if      pos('TC77', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := 'Ư��[TC77-��ġ��]'  //��ġ��
         else if pos('TC93', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := 'Ư��[TC93-��ġ��]'  //��ġ��
         else if pos('TC79', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := 'Ư��[TC79-�ӽ�]'    //�ӽ�
         else if pos('TC78', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := 'Ư��[TC78-����]'    //����
         else if pos('TCA7', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := 'Ư��[TCA7-����]'    //����
         else if Trim(sTk_YHGM) = ''                                      then sTk_YHGM := 'Ư��';              //Ư��

         if      qryTemp.FieldByName('gubn_Tkcg').AsString = '1' then begin sTk_YHGM := sTk_YHGM + '(ȸû)'; sTkCg := '1'; end
         else if qryTemp.FieldByName('gubn_Tkcg').AsString = '2' then begin sTk_YHGM := sTk_YHGM + '(��û)'; sTkCg := '2'; end;
      end;
      //========================================================================

      //��� �׸� ã��..
      //------------------------------------------------------------------------
      sSelect := ' select * from tk_sokun2009 ';
      sWhere  := ' where cod_jisa = ''' + sJisa_1 + '''';
      sWhere  := sWhere + ' and num_jumin = ''' + edtJumin.Text + '''';
      sWhere  := sWhere + ' and dat_gmgn = ''' + sDate1 + '''';
      sWhere  := sWhere + ' and gubn_cha = ''1''';

      qryTemp.Active := False;
      qryTemp.SQL.Clear;
      qryTemp.SQL.Add(sSelect + sWhere + sOrderBy);
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then
      begin
         while not qryTemp.Eof do
         begin
            if qryTemp.FieldByName('cod_pan').AsString  = 'R' then sTemp_Tk := sTemp_Tk + UF_Hmatter_cnvt(qryTemp.FieldByName('cod_Hmatter').AsString) + ',';

            qryTemp.Next;
         end;
      end;
      //========================================================================

      qryTemp.Active := False;
   end;

   if gubun = '1' then
      begin
      {�Ϲ�}
      if (sTemp_NS <> '') and (sNsGubn_1 <> 'L') and (sNsCg_1 = '2') then
      begin
         if sNsGubn_1 = 'N' then sTemp := '�Ϲݰ���';
         if sNsGubn_1 = 'A' then sTemp := '���κ�����';
//         if sNsGubn_1 = 'L' then sTemp := '������ȯ�����';
      
         {if      sNsCg_1 = '1' then sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(ȸ��)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else if sNsCg_1 = '2' then sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(����)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else                       sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(' + sNsCg_1 + ')]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')' ;
         }
         sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(����)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')';
      end
      else if (sNsGubn_1 = 'L') and (sNsCg_1 = '2') then
      begin
         if sNsGubn_1 = 'L' then sTemp := '������ȯ�����';
         {if      sNsCg_1 = '1' then sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(ȸ��)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else if sNsCg_1 = '2' then sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(����)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else                       sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(' + sNsCg_1 + ')]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')' ;
         }
         sTemp := sTemp + '[����(' + sJisa_1 + ')|������(' + sDate1 + ')|û��(����)]' + '���(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')';

         With qryNs_sokun Do
         Begin
            Active := False;
            ParamByName('Cod_Jisa').AsString  := sJisa_1;
            ParamByName('Dat_Gmgn').AsString  := sDate1;
            ParamByName('num_id').AsString    := sID_1;
            Active := True;
            If RecordCount > 0 Then
            begin
               if (FieldByName('cod_panjg1').AsString = '') and
                  (FieldByName('cod_panjg2').AsString = '') and
                  (FieldByName('cod_panjg3').AsString = '') and
                  (FieldByName('cod_panjg4').AsString = '') and
                  (FieldByName('cod_panjg5').AsString = '') then
               begin
                  sTemp := sTemp + #10#13#13 + '������� ���, 1�� �Ұ����� �ȵ�';
               end;
            end
            else
            begin
                  sTemp := sTemp + #10#13#13 + '������� ���, 1�� �Ұ����� �ȵ�';
            end;
            Active := False;
         End;
      end;
   end;

   {if gubun = '2' then
   begin
      if sTemp_Tk <> '' then GP_ComboDisplay(cmbGUBN_TKCG, sTkCg, 1);
      //Ư��
      if (sTemp_Tk <> '') and (sTemp_NS <> '') then sTemp := sTemp + #10#13 + 'Ư��[����(' + sJisa_1 + ')|������(' + sDate1 + ')]' + '���(' + copy(sTemp_Tk,1,length(sTemp_Tk)-1) + ')'
      else if sTemp_Tk <> '' then sTemp := 'Ư��[����(' + sJisa_1 + ')|������(' + sDate1 + ')]' + '���(' + copy(sTemp_Tk,1,length(sTemp_Tk)-1) + ')';
   end;}

   if sTemp <> '' then ShowMessage(sTemp + #10#13#13 + ' ## ' + sTk_YHGM + ' ��˴���� �Դϴ�.');
end;

procedure TfrmCK5I17.Gmgn_retest;
var sDate_NS1, sDate_NS2, sDate_TK1, sDate_TK2, sID, sJisa, sNsGubn, sNsCg, sSelect, sWhere, sOrder : String;
begin
    with qryGmgn do
    begin
     Active := False;
     SQL.Clear;

     sSelect := ' SELECT cod_jisa, num_id, dat_gmgn, desc_name, cod_jangbi, cod_hul, cod_chuga, cod_dc, gubn_nosin, gubn_adult, gubn_life, gubn_tkgm, gubn_nscg, gubn_adcg, gubn_lfcg  FROM gumgin with(nolock)';
     sWhere := ' WHERE num_jumin = ''' + edtJumin.Text + '''';
     sOrder := ' ORDER BY dat_gmgn';

     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrder);
     Active := True;

     if RecordCount > 0 then
     begin
        sDate_NS1 := ''; sDate_NS2 := ''; sDate_TK1 := ''; sDate_TK2 := ''; sID := ''; sJisa := '';
        sNsGubn   := ''; sNsCg     := '';

        while not Eof do
        begin
           if (FieldByName('gubn_nosin').AsString = '1') or
              (FieldByName('gubn_adult').AsString = '1') or
              (FieldByName('gubn_life').AsString  = '1') then
           begin
              sDate_NS1 := FieldByName('dat_gmgn').AsString;
              sID       := FieldByName('num_id').AsString;
              sJisa     := FieldByName('cod_jisa').AsString;

              if FieldByName('gubn_nosin').AsString = '1' then sNsGubn := 'N';
              if FieldByName('gubn_adult').AsString = '1' then sNsGubn := 'A';
              if FieldByName('gubn_life').AsString  = '1' then sNsGubn := 'L';

              if FieldByName('gubn_nosin').AsString = '1' then
                 sNsCg := FieldByName('gubn_nscg').AsString;
              if FieldByName('gubn_adult').AsString = '1' then
                 sNsCg := FieldByName('gubn_adcg').AsString;
              if FieldByName('gubn_life').AsString  = '1' then
                 sNsCg := FieldByName('gubn_lfcg').AsString;
           end;

           if (FieldByName('gubn_nosin').AsString = '2') or
              (FieldByName('gubn_adult').AsString = '2') or
              (FieldByName('gubn_life').AsString  = '2') then sDate_NS2 := FieldByName('dat_gmgn').AsString;

           if (FieldByName('gubn_tkgm').AsString = '1') then
           begin
              sDate_TK1 := FieldByName('dat_gmgn').AsString;
              sID       := FieldByName('num_id').AsString;
              sJisa     := FieldByName('cod_jisa').AsString;
           end;

           if (FieldByName('gubn_tkgm').AsString = '2') then sDate_TK2 := FieldByName('dat_gmgn').AsString;

           //---------------------------------------------------------------
           Next;
        end;
     end;
     // Table�� Disconnect
     Active := False;
    end;

    if (copy(sDate_NS1,1,4) = copy(StringReplace(Trim(cmb_gmgn.text),'-','',[rfReplaceAll]),1,4)) or
     ((copy(StringReplace(Trim(cmb_gmgn.text),'-','',[rfReplaceAll]),5,2) = '01') and (copy(sDate_NS1,1,4) = IntTostr(StrToInt(copy(StringReplace(Trim(cmb_gmgn.text),'-','',[rfReplaceAll]),1,4))-1))) then
    begin
     if      sDate_NS2 = ''        then UP_GumsaCheck('1', sID, sJisa, sNsGubn, sNsCg, sDate_NS1, sDate_NS2)
     else if sDate_NS1 > sDate_NS2 then UP_GumsaCheck('1', sID, sJisa, sNsGubn, sNsCg, sDate_NS1, sDate_NS2);
    end;
end;

procedure TfrmCK5I17.CreateTextArray_Gmgn;
var i, index : Integer;
begin
  inherited;

  grdGmgn.Rows := 0;

  With  qry_gmgn  do
  begin
     Active := False;
     ParamByName('num_jumin').AsString := edtJumin.Text;
     Active := True;

     GP_query_log(qry_gmgn, 'CK5I17', 'Q', 'N');

     index := 0;
     if qry_gmgn.RecordCount > 0 then
     begin
        while Not qry_gmgn.Eof do
        begin
           UP_VarMemSet_Gmgn(index);
           UV_vNum_id_G[index]      := FieldByName('num_id').AsString;
           UV_vDat_gmgn_G[index]    := FieldByName('dat_gmgn').AsString;
           UV_vCod_janbi_G[index]   := FieldByName('cod_jangbi').AsString;
           UV_vCod_hul_G[index]     := FieldByName('cod_hul').AsString;
           UV_vCod_cancer_G[index]  := FieldByName('cod_cancer').AsString;
           UV_vCod_chuga_G[index]   := FieldByName('cod_chuga').AsString;
           UV_vNum_tel_G[index]     := FieldByName('num_tel').AsString;
           UV_vGubn_nosin_G[index]  := FieldByName('gubn_nosin').AsString;
           UV_vGubn_nscg_G[index]   := FieldByName('gubn_nscg').AsString;
           UV_vGubn_adult_G[index]  := FieldByName('gubn_adult').AsString;
           UV_vGubn_adcg_G[index]   := FieldByName('gubn_adcg').AsString;
           UV_vGubn_life_G[index]   := FieldByName('gubn_life').AsString;
           UV_vGubn_lfcg_G[index]   := FieldByName('gubn_lfcg').AsString;
           UV_vDesc_yhname_G[index] := FieldByName('desc_yhname').AsString;
           UV_vDesc_saler_G[index]  := FieldByName('desc_name').AsString;
           UV_vGubn_jinch_G[index]  := FieldByName('desc_code').AsString;

           Next;
           Inc(index);
        end;
     end
     else
     begin
        // �ڷᰡ ������ ǥ��
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

  grdGmgn.Rows := index;

end;


procedure TfrmCK5I17.CreateTextArray;
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
  sSelect := '';
  sWhere  := '';
  sOrder  := '';

  With  qryKicho  do
  begin
     Active := False;
     sSelect := ' select K.* from gumgin G with (Nolock) inner join kicho K with (Nolock) ON G.Cod_jisa = K.cod_jisa and G.num_id = K.num_id and G.dat_gmgn = K.dat_gmgn ';
     sWhere  := ' where G.num_jumin = ''' + edtJumin.Text + ''' ';
     sOrder  := ' order by K.dat_gmgn DESC ';

     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrder );
     Active := True;

     GP_query_log(qryKicho, 'CK5I17', 'Q', 'N');

     index := 0;
     UP_VarMemSet(index, 'K');
     UV_vSinjang[0]     := '��        ��   (cm)';
     UV_vChejung[0]     := 'ü        ��   (kg)';
     UV_vBiman[0]       := '��    ��   ��   (%)';
     UV_vHul_h[0]       := '�ְ� ����    (mmHg)';
     UV_vHul_l[0]       := '���� ����    (mmHg)';
     UV_vear_left5[0]   := 'û��(�� 500Hz dB)';
     UV_vear_left10[0]  := 'û��(��1000Hz dB)';
     UV_vear_lef20[0]   := 'û��(��2000Hz dB)';
     UV_vear_left30[0]  := 'û��(��3000Hz dB)';
     UV_vear_left40[0]  := 'û��(��4000Hz dB)';
     UV_vear_left60[0]  := 'û��(��6000Hz dB)';
     UV_vear_right5[0]  := 'û��(�� 500Hz dB)';
     UV_vear_right10[0] := 'û��(��1000Hz dB)';
     UV_vear_right20[0] := 'û��(��2000Hz dB)';
     UV_vear_right30[0] := 'û��(��3000Hz dB)';
     UV_vear_right40[0] := 'û��(��4000Hz dB)';
     UV_vear_right60[0] := 'û��(��6000Hz dB)';
     UV_vdesc_ear[0]    := 'û�¼Ұ�';
     UV_veye_lman[0]    := '��         ��(��)';
     UV_veye_rman[0]    := '��         ��(��)';
     UV_veye_lkyo[0]    := '�� ��   �� ��(��)';
     UV_veye_rkyo[0]    := '�� ��   �� ��(��)';
     UV_vanap_left[0]   := '��   ��(��)  (mmHg)';
     UV_vanap_right[0]  := '��   ��(��)  (mmHg)';
     UV_vgubn_seksin[0] := '��         ��';
     UV_vhyungwi[0]     := '����(�����ѷ�)';
     UV_vbookwi[0]      := '����(�㸮�ѷ�)';
     Inc(index);

     if qryKicho.RecordCount > 0 then
     begin
        while Not Eof do
        begin
           UP_VarMemSet(index, 'K');
           UV_vHul_h[index]       := FieldByName('Hul_h').AsString;
           UV_vHul_l[index]       := FieldByName('Hul_l').AsString;
           UV_vBiman[index]       := FormatFloat('##0.0', strtofloat(FieldByName('Biman').AsString) );
           UV_vSinjang[index]     := FieldByName('Sinjang').AsString;
           UV_vChejung[index]     := FieldByName('Chejung').AsString;
           UV_vear_left5[index]   := FieldByName('ear_left5').AsString;
           UV_vear_left10[index]  := FieldByName('ear_left10').AsString;
           UV_vear_lef20[index]   := FieldByName('ear_left20').AsString;
           UV_vear_left30[index]  := FieldByName('ear_left30').AsString;
           UV_vear_left40[index]  := FieldByName('ear_left40').AsString;
           UV_vear_left60[index]  := FieldByName('ear_left60').AsString;
           UV_vear_right5[index]  := FieldByName('ear_right5').AsString;
           UV_vear_right10[index] := FieldByName('ear_right10').AsString;
           UV_vear_right20[index] := FieldByName('ear_right20').AsString;
           UV_vear_right30[index] := FieldByName('ear_right30').AsString;
           UV_vear_right40[index] := FieldByName('ear_right40').AsString;
           UV_vear_right60[index] := FieldByName('ear_right60').AsString;
           UV_vdesc_ear[index]    := FieldByName('desc_ear').AsString;
           UV_veye_lman[index]    := FieldByName('eye_lman').AsString;
           UV_veye_rman[index]    := FieldByName('eye_rman').AsString;
           UV_veye_lkyo[index]    := FieldByName('eye_lkyo').AsString;
           UV_veye_rkyo[index]    := FieldByName('eye_rkyo').AsString;
           UV_vanap_left[index]   := FieldByName('anap_left').AsString;
           UV_vanap_right[index]  := FieldByName('anap_right').AsString;
           UV_vgubn_seksin[index] := FieldByName('gubn_seksin').AsString;
           UV_vhyungwi[index]     := FieldByName('hyungwi').AsString;
           UV_vbookwi[index]      := FieldByName('bookwi').AsString;
           UV_vDat_gmgn[index]    := Copy(FieldByName('Dat_gmgn').AsString,1,4) + '-' + Copy(FieldByName('Dat_gmgn').AsString,5,2) + '-' + Copy(FieldByName('Dat_gmgn').AsString,7,2);

           Next;
           Inc(index);
        end;
     end
     else
     begin
        // �ڷᰡ ������ ǥ��
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid�� �ڷḦ �Ҵ�
   if index > 0 then
   begin
      grd_Kicho.Rows := 27;
      grd_Kicho.Cols := index;
   end;
end;

procedure TfrmCK5I17.grd_KichoCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;

   // Grid�� ȯ�� ����
   with grd_Kicho do
   begin
      // Column Title
      Col[1].Heading := '�˻��׸�';
      Col[1].width   := 136;

      if DataCol > 1 then
      begin
         Col[DataCol].Heading := UV_vDat_gmgn[DataCol-1];
         Col[DataCol].width   := 80;
      end;
   end;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataRow of
       1 : Value := UV_vSinjang[DataCol-1];
       2 : Value := UV_vChejung[DataCol-1];
       3 : Value := UV_vBiman[DataCol-1];
       4 : Value := UV_vHul_h[DataCol-1];
       5 : Value := UV_vHul_l[DataCol-1];
       6 : Value := UV_veye_lman[DataCol-1];
       7 : Value := UV_veye_rman[DataCol-1];
       8 : Value := UV_veye_lkyo[DataCol-1];
       9 : Value := UV_veye_rkyo[DataCol-1];
      10 : Value := UV_vear_left5[DataCol-1];
      11 : Value := UV_vear_left10[DataCol-1];
      12 : Value := UV_vear_lef20[DataCol-1];
      13 : Value := UV_vear_left30[DataCol-1];
      14 : Value := UV_vear_left40[DataCol-1];
      15 : Value := UV_vear_left60[DataCol-1];
      16 : Value := UV_vear_right5[DataCol-1];
      17 : Value := UV_vear_right10[DataCol-1];
      18 : Value := UV_vear_right20[DataCol-1];
      19 : Value := UV_vear_right30[DataCol-1];
      20 : Value := UV_vear_right40[DataCol-1];
      21 : Value := UV_vear_right60[DataCol-1];
      22 : Value := UV_vdesc_ear[DataCol-1];
      23 : Value := UV_vanap_left[DataCol-1];
      24 : Value := UV_vanap_right[DataCol-1];
      25 : Value := UV_vgubn_seksin[DataCol-1];
      26 : Value := UV_vhyungwi[DataCol-1];
      27 : Value := UV_vbookwi[DataCol-1];
   end;
end;

procedure TfrmCK5I17.CreateTextArray_Hul;
var I, J, K, index, iRet, iAge, Temp : Integer;
    cod_name, sMan, sGubn, sResult,  sCod_hm : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;
    sSelect, sWhere, sOrder : String;
    // �ӻ�����ġ
    eLow, eHigh, eIms : Extended;
    sMunja : String;
    vCod_Tkgum : Variant;
begin
  inherited;

   // Grid �ʱ�ȭ
   grd_Hul.Rows := 0;
   UP_VarMemSet(0, 'HG');
   sSelect := '';
   sWhere  := '';
   sOrder  := '';

   with qryHul_gumgin do
   begin
      Active := False;
      cmb_gmgn.items.Clear;
      qryHul.Filter := '';
      ParamByName('num_jumin').AsString := edtJumin.Text;
      Active := True;

      //���װ˻��׸�����
      sCod_hm := '';
      Index   := 0;
      if qryHul_gumgin.RecordCount > 0 then
      begin
         while not qryHul_gumgin.Eof do
         begin
            UP_VarMemSet(index, 'H');

            UV_vHul_Num_id[index]    := FieldByName('num_id').AsString;
            UV_vHul_dat_gmgn[index]  := FieldByName('dat_gmgn').AsString;

            //�������� ǥ��..
            cmb_gmgn.Items.Add(Copy(FieldByName('Dat_gmgn').AsString,1,4) + '-' + Copy(FieldByName('Dat_gmgn').AsString,5,2) + '-' + Copy(FieldByName('Dat_gmgn').AsString,7,2) +
                               '(' + FieldByName('cod_jisa').AsString  + ')' +
                               '(' + FieldByName('dat_bunju').AsString + ')' +
                               '(' + FieldByName('num_bunju').AsString + ')' +
                               '(' + FieldByName('num_id').AsString    + ')' );
                               
            if qryHul_gumgin.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if (pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) = 0) and
                        (Copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if qryHul_gumgin.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if (pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) = 0) and
                        (Copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if qryHul_gumgin.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if (pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) = 0) and
                        (Copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;
            iRet := GF_MulToSingle(qryHul_gumgin.FieldByName('cod_chuga').AsString, 4, UV_vCod_chuga);
            for i := 0 to iRet-1 do
            begin
               if (pos(UV_vCod_chuga[i], sCod_hm) = 0) and
                  (Copy(UV_vCod_chuga[i],1,2) <> 'JJ') then
                  sCod_hm := sCod_hm + UV_vCod_chuga[i];
            end;

            // ���, ���κ�, ������ ���� �˻��׸� ����
            cod_name := '';
            // �������Display
            if FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            //���κ�����Display
            if FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            //��������Display
            if FieldByName('gubn_agab').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            //������ȯ������Display
            if FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            //Ư������Display
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString  := FieldByName('cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := FieldByName('num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString  := FieldByName('dat_gmgn').AsString;
               qryTkgum.Active := True;
               if qryTkgum.RecordCount > 0 then
               begin
                  iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_Tkgum);
                  for Temp := 0 to iRet - 1 do
                  begin
                    qryPf_hm.Active := False;
                    qryPf_hm.ParamByName('In_Cod_Pf').AsString := vCod_Tkgum[Temp];
                    qryPf_hm.Active := True;
                    if qryPf_hm.RecordCount > 0 then
                    begin
                       while not qryPf_hm.Eof do
                       begin
                          cod_name := cod_name + qryPf_hm.FieldByName('COD_HM').AsString;
                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
                  cod_name := cod_name + qryTkgum.FieldByName('Cod_chuga').AsString;
               end;           //if(qryTkgum) end;
            end;

            iRet := GF_MulToSingle(cod_name, 4, UV_vCod_totyh);
            for i := 0 to iRet-1 do
            begin
               if (pos(UV_vCod_totyh[i], sCod_hm) = 0) and
                  (Copy(UV_vCod_totyh[i],1,2) <> 'JJ') and
                  (Copy(UV_vCod_totyh[i],1,2) <> 'JK') then
                  sCod_hm := sCod_hm + UV_vCod_totyh[i];
            end;

            inc(index);
            Next;
         end;

         if sCod_hm <> '' then
         begin
            iRet := GF_MulToSingle(sCod_hm, 4, UV_vTot_cod_hm);
         end;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   UP_VarMemSet(iRet, 'HG');
   for I := 0 to index-1 do
   begin
      // �ֹι�ȣ�� �̿��Ͽ� ������ ���̸� ����
      GP_JuminToAge(edtJumin.Text,UV_vHul_dat_gmgn[I], sMan, iAge);
      
      // ������ ���̸� �������� Column���� ����
      sGubn := '';
      if iAge < 10 then sGubn := '1'
      else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
      else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
      else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
      else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
      else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
      else sGubn := '7';
      
      if sMan = 'F' then sGubn := sGubn + 'f';

      // ���ֹ�ȣ�� ���� ���װ���� ȹ��
      With qryGlkwa do
      begin
         Active := False;

         sSelect := ' SELECT gubn_part, desc_glkwa1,  desc_glkwa2, desc_glkwa3 FROM gulkwa ';
         sWhere  := ' where num_id = ''' + UV_vHul_Num_id[I] + ''' ';
         sWhere  := sWhere + ' and dat_gmgn = ''' + UV_vHul_dat_gmgn[I] + ''' ';
         sOrder  := ' order by dat_gmgn DESC ';

         SQL.Clear;
         SQL.Add(sSelect + sWhere + sOrder );
         Active := True;
      end;

//      with qryGlkwa do
//      begin
//         Active := False;
//         ParamByName('dat_gmgn').AsString  := UV_vHul_dat_gmgn[I];
//         ParamByName('NUM_ID').AsString    := UV_sNum_id;
//         Active := True;
//      end;

      for J := 0 to VarArrayHighBound(UV_vTot_cod_hm, 1)-1 do
      begin
         // �ʱⰪ ����
         eLow   := 0;
         eHigh  := 0;
         eIms   := 0;
         sMunja := '';

         // �˻��׸� ���� �ӻ�����ġ�� ȹ��
         qryHm.Filter := 'COD_HM = ''' + UV_vTot_cod_hm[J] + ''' AND ' +
                         'DAT_APPLY <= ''' + UV_vHul_dat_gmgn[I] + '''';

         //�׸��
         if trim(UV_vTot_cod_hm_name[0, J]) = '' then
         begin
            UV_vTot_cod_hm_name[0, J] := qryHm.FIeldByName('desc_kor').AsString;
            if trim(UV_vTot_cod_hm_name[0, J]) = '' then
            begin
               with qryHm_name do
               begin
                  Active := False;
                  ParamByName('cod_hm').AsString  := UV_vTot_cod_hm[J];
                  Active := True;
                  UV_vTot_cod_hm_name[0, J] := qryHm_name.FIeldByName('desc_kor').AsString;
               end;
            end;
         end;

         // ���װ˻����� Check(���PF, �߰��˻� ����)
         if qryHm.FieldByName('GUBN_PART').AsString > '10' then continue;

         eLow   := GF_StrToFloat(qryHm.FieldByName('IMS_LOW'   + sGubn).AsString);
         eHigh  := GF_StrToFloat(qryHm.FieldByName('IMS_HIGH'  + sGubn).AsString);
         sMunja := qryHm.FIeldByName('IMS_MUNJA').AsString;

         // �ش� �˻翡 ���� ����� ����(Start Pos, End Pos�� �̿�)
         qryGlkwa.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';
         UV_fGulkwa := '';
         UV_fGulkwa1 := qryGlkwa.FieldByName('DESC_GLKWA1').AsString;
         UV_fGulkwa2 := qryGlkwa.FieldByName('DESC_GLKWA2').AsString;
         UV_fGulkwa3 := qryGlkwa.FieldByName('DESC_GLKWA3').AsString;
         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
         if (qryHm.FieldByName('POS_START').AsInteger <> 0) and (qryHm.FieldByName('POS_END').AsInteger <> 0) then
         begin
            sResult := Trim(Copy(UV_fGulkwa,
                         qryHm.FieldByName('POS_START').AsInteger,6));
         end
         else sResult := '';

         UV_vHul_Gulkwa[I, J] := sResult;

         if UV_vHul_Low_High[0,J] = '' then
         begin
            if qryHm.FieldByName('GUBN_HM').AsString = '1' then
            begin
               UV_vHul_Low_High[0,J] := floatTostr(eLow) + ' ~ ' + floatTostr(eHigh);
            end
            else if qryHm.FieldByName('GUBN_HM').AsString = '2' then
               UV_vHul_Low_High[0,J] := sMunja;
         end;

         if (qryHm.FieldByName('GUBN_HM').AsString = '1') and (trim(sResult) <> '') and (trim(sResult) < '999999') then
         begin
          if (StrToFloat(sResult) < eLow) or (StrToFloat(sResult) > eHigh) then
             UV_vFont_color[I, J] := 'clRed'
          else UV_vFont_color[I, J] := 'clBlack';
         end
         else if (qryHm.FieldByName('GUBN_HM').AsString = '2') and (trim(sResult) <> '') then
         begin
            if trim(sResult) <> trim(sMunja) then
               UV_vFont_color[I, J] := 'clRed'
            else UV_vFont_color[I, J] := 'clBlack';
         end;
      end;
   end;

   // Grid�� �ڷḦ �Ҵ�
   if index > 0 then
   begin
      grd_Hul.Rows := J;
      grd_Hul.Cols := I + 3;
   end;
end;

procedure TfrmCK5I17.grd_HulCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
var I, J : Integer;
begin
  inherited;
   // Grid�� ȯ�� ����
   with grd_Hul do
   begin
      J := grd_Hul.Rows;
      I := grd_Hul.Cols;

      // Column Title
      Col[1].Heading := '�ڵ�';
      Col[2].Heading := '�׸��';
      Col[3].Heading := '�ӻ�����ġ';

      // Column Align
      Col[1].Alignment := taCenter;
      Col[3].Alignment := taCenter;

      if DataCol > 3 then
      begin
         Col[DataCol].Heading := Copy(UV_vHul_dat_gmgn[DataCol-4],1,4) + '-' + Copy(UV_vHul_dat_gmgn[DataCol-4],5,2) + '-' + Copy(UV_vHul_dat_gmgn[DataCol-4],7,2);
         Col[DataCol].width   := 80;
      end;
   end;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vTot_cod_hm[DataRow-1];
       2 : Value := UV_vTot_cod_hm_name[0, DataRow-1];
       3 : Value := UV_vHul_Low_High[0, DataRow-1];
       4..50 :  Value := UV_vHul_Gulkwa[DataCol-4, DataRow-1];
   end;

   grd_Hul .AssignCellFont(DataCol , DataRow);
   case DataCol of
     4..50 :
     begin
        if UV_vFont_color[DataCol-4, DataRow-1] = 'clRed' then
           grd_Hul.CellFont[DataCol, DataRow].Color := clRed
        else
           grd_Hul.CellFont[DataCol, DataRow].Color := clBlack;
     end;
   end;   
end;

procedure TfrmCK5I17.grd_HulRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var I, J : Integer;  
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grd_Hul.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grd_Hul.Rows Do
      begin
         for J := 1 to grd_Hul.Cols Do
         begin
            grd_Hul.CellColor[J, i] := clWindow;
         end;

         if (i mod 2 = 0) then
         begin
            for J := 1 to grd_Hul.Cols Do
            begin
               grd_Hul.CellColor[J, i] := $00FEEEC2;
            end;
         end
         else
         begin
            for J := 1 to grd_Hul.Cols Do
            begin
               grd_Hul.CellColor[J, i] := clWindow;
            end;
         end;
         Next;
      end;
   end;

   if (grd_Hul.Rows > 0) and (NewRow > 0) then
   begin
      for J := 1 to grd_Hul.Cols Do
      begin
         grd_Hul.CellColor[J, NewRow] := $001AD7C9;
         Next;
      end;
   end;
end;


procedure TfrmCK5I17.CreateTextArray_Jangbi;
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;

  sSelect := '';
  sWhere  := '';
  sOrder  := '';

  // Column Align
  grd_Jangbi.Col[1].Alignment := taCenter;
  grd_Jangbi.Col[2].Alignment := taCenter;
  grd_Jangbi.Col[3].Alignment := taCenter;
  grd_Jangbi.Col[4].Alignment := taCenter;

  // Readonly
  grd_Jangbi.Col[1].ReadOnly := True;
  grd_Jangbi.Col[2].ReadOnly := False;
  grd_Jangbi.Col[3].ReadOnly := True;
  grd_Jangbi.Col[4].ReadOnly := True;

  with qryJangbi do
  begin
     Active := False;

     sSelect := ' select J.dat_gmgn, H.desc_kor, J.cod_sokun, J.gubn_pan, J.desc_sokun, J.desc_sbsg, J.desc_sbsg1, J.desc_sbsg2, J.desc_sbsg3 '
              + ' from gumgin G with(nolock) '
              + ' inner join jangbi J with(nolock) '
              + ' on G.Cod_jisa = J.Cod_jisa '
              + ' and G.num_id = J.num_id '
              + ' and G.dat_gmgn = J.dat_gmgn '
              + ' inner join hangmok H  On J.cod_jangbi = H.cod_hm and H.dat_apply in (select MAX(dat_apply) from hangmok where cod_hm = H.cod_hm) ';

     sWhere  := ' where G.num_jumin = ''' + edtJumin.Text + ''' ';

     sOrder  := ' order by G.dat_gmgn DESC, J.cod_jangbi ';

     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrder );
     Active := True;

     index := 0;
     if qryJangbi.RecordCount > 0 then
     begin
        while Not Eof do
        begin
           UP_VarMemSet(index, 'J');

           UV_vDat_gmgn_Jangbi[index]   := Copy(FieldByName('Dat_gmgn').AsString,1,4) + '-' + Copy(FieldByName('Dat_gmgn').AsString,5,2) + '-' + Copy(FieldByName('Dat_gmgn').AsString,7,2);
           UV_vDesc_kor[index]          := FieldByName('Desc_kor').AsString;
           UV_vCod_sokun_Jangbi[index]  := FieldByName('Cod_sokun').AsString;
           UV_vGubn_pan[index]          := FieldByName('Gubn_pan').AsString;
           UV_vDesc_sokun_Jangbi[index] := FieldByName('Desc_sokun').AsString;
           UV_vDesc_sbsg_Jangbi[index]  := FieldByName('Desc_sbsg').AsString + FieldByName('Desc_sbsg1').AsString + FieldByName('Desc_sbsg2').AsString  + FieldByName('Desc_sbsg3').AsString;

           Next;
           Inc(index);
        end;
     end
     else
     begin
        // �ڷᰡ ������ ǥ��
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid�� �ڷḦ �Ҵ�
   if index > 0 then
      grd_Jangbi.Rows := index;
end;

procedure TfrmCK5I17.grd_JangbiCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vDat_gmgn_Jangbi[DataRow-1];
       2 : Value := UV_vDesc_kor[DataRow-1];
       3 : Value := UV_vCod_sokun_Jangbi[DataRow-1];
       4 : Value := UV_vGubn_pan[DataRow-1];
       5 : Value := UV_vDesc_sokun_Jangbi[DataRow-1];
       6 : Value := UV_vDesc_sbsg_Jangbi[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grd_JangbiRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i : integer;
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grd_Jangbi.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grd_Jangbi.Rows Do
      begin
         grd_Jangbi.CellColor[1, i]  := clWindow;
         grd_Jangbi.CellColor[2, i]  := clWindow;
         grd_Jangbi.CellColor[3, i]  := clWindow;
         grd_Jangbi.CellColor[4, i]  := clWindow;
         grd_Jangbi.CellColor[5, i]  := clWindow;
         grd_Jangbi.CellColor[6, i]  := clWindow;

         if (i mod 2 = 0) then
         begin
            grd_Jangbi.CellColor[1, i]   := $00FEEEC2;
            grd_Jangbi.CellColor[2, i]   := $00FEEEC2;
            grd_Jangbi.CellColor[3, i]   := $00FEEEC2;
            grd_Jangbi.CellColor[4, i]   := $00FEEEC2;
            grd_Jangbi.CellColor[5, i]   := $00FEEEC2;
            grd_Jangbi.CellColor[6, i]   := $00FEEEC2;
         end
         else
         begin
            grd_Jangbi.CellColor[1, i]  := clWindow;
            grd_Jangbi.CellColor[2, i]  := clWindow;
            grd_Jangbi.CellColor[3, i]  := clWindow;
            grd_Jangbi.CellColor[4, i]  := clWindow;
            grd_Jangbi.CellColor[5, i]  := clWindow;
            grd_Jangbi.CellColor[6, i]  := clWindow;
         end;
         Next;
      end;
   end;

   if (grd_Jangbi.Rows > 0) and (NewRow > 0) then
   begin
      grd_Jangbi.CellColor[1, NewRow]  := $001AD7C9;
      grd_Jangbi.CellColor[2, NewRow]  := $001AD7C9;
      grd_Jangbi.CellColor[3, NewRow]  := $001AD7C9;
      grd_Jangbi.CellColor[4, NewRow]  := $001AD7C9;
      grd_Jangbi.CellColor[5, NewRow]  := $001AD7C9;
      grd_Jangbi.CellColor[6, NewRow]  := $001AD7C9;
   end;
end;

procedure TfrmCK5I17.CreateTextArray_Sokun;
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
  sSelect := '';
  sWhere  := '';
  sOrder  := '';

  // Column Align
  grd_sokun.Col[1].Alignment := taCenter;
  grd_sokun.Col[2].Alignment := taCenter;
  grd_sokun.Col[3].Alignment := taCenter;

  // Readonly
  grd_sokun.Col[1].ReadOnly := True;
  grd_sokun.Col[2].ReadOnly := True;
  grd_sokun.Col[3].ReadOnly := True;

  with qryTot_sokun do
  begin
     Active := False;

     sSelect := ' select cod_jisa, num_id, dat_gmgn, cod_sokun, desc_sokun, gubn_hsok, num_seq ' +
                ' from ( ' +
                ' select H.cod_jisa, H.num_id, H.dat_gmgn, H.cod_sokun, H.desc_sokun, H.gubn_hsok, H.num_seq ' +
                ' from gumgin G with (nolock) inner join hul_sokun H with (nolock) ' +
                '      On G.cod_jisa = H.cod_jisa and G.num_id = H.num_id and G.dat_gmgn = H.dat_gmgn ' +
                ' where H.cod_bjjs in (''11'',''15'',''12'',''34'',''41'',''43'',''51'',''61'',''71'') '+
                '   and G.num_jumin = ''' + edtJumin.Text + ''' ' +
                ' union ' +
                ' select T.cod_jisa, T.num_id, T.dat_gmgn, T.cod_sokun, T.desc_sokun, ''tot_sokun'' as gubn_hsok, T.num_seq ' +
                ' from gumgin G with (nolock) inner join tot_sokun T with (nolock) ' +
                '      On G.cod_jisa = T.cod_jisa and G.num_id = T.num_id and G.dat_gmgn = T.dat_gmgn ' +
                ' where G.cod_jisa in (''11'',''15'',''12'',''34'',''41'',''43'',''51'',''61'',''71'') ' +
                ' and G.num_jumin = ''' + edtJumin.Text + ''' ) A ' +
                ' order by A.dat_gmgn DESC, A.gubn_hsok, A.num_seq ';

     SQL.Clear;
     SQL.Add(sSelect);
     Active := True;

     index := 0;
     if qryTot_sokun.RecordCount > 0 then
     begin
        while Not Eof do
        begin
           UP_VarMemSet(index, 'S');
           UV_vDat_gmgn_sokun[index]  := Copy(FieldByName('Dat_gmgn').AsString,1,4) + '-' + Copy(FieldByName('Dat_gmgn').AsString,5,2) + '-' + Copy(FieldByName('Dat_gmgn').AsString,7,2);

           if      FieldByName('gubn_hsok').AsString = '1'         then UV_vgubn_sokun[index] := '���׼Ұ�'
           else if FieldByName('gubn_hsok').AsString = '2'         then UV_vgubn_sokun[index] := '���̿��'
           else if FieldByName('gubn_hsok').AsString = '3'         then UV_vgubn_sokun[index] := '����'
           else if FieldByName('gubn_hsok').AsString = 'tot_sokun' then UV_vgubn_sokun[index] := '���ռҰ�';

           UV_vCod_sokun[index]  := FieldByName('cod_sokun').AsString;
           UV_vDesc_sokun[index] := FieldByName('desc_sokun').AsString;
           Next;
           Inc(index);
        end;
     end
     else
     begin
        // �ڷᰡ ������ ǥ��
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid�� �ڷḦ �Ҵ�
   if index > 0 then
      grd_sokun.Rows := index;
end;

procedure TfrmCK5I17.CreateTextArray_Cell;
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin

     retDesc_Sokun.Lines.Clear;
     retDesc_Sokun1.Lines.Clear;
     retDESC_KGBL.Lines.Clear;
     Panel_PAP.Visible := True;

     edtDESC_BUWI.Text  := ''; mskDAT_RESULT.Text := ''; edtGUBN_PAN.Text := '';
     edtDESC_VIRUS.Text := ''; edtGum_Remark.Text := ''; edtGum_Form3_Etc.Text := '';
     edtSun_Form4.Text  := '';

     cmbCell_Hangmok.ItemIndex := -1;
     cmbCOD_DOCT.ItemIndex := -1;
     cmbGUBN_CLASS.ItemIndex := -1;
     CmbSang_Form1_gubn.ItemIndex := -1;

     ckYSNO_VIRUS1.Checked := False; ckYSNO_VIRUS2.Checked := False; ckYSNO_VIRUS3.Checked := False;
     ckYSNO_VIRUS4.Checked := False; ckYSNO_VIRUS5.Checked := False; ckYSNO_VIRUS6.Checked := False;
     ckYSNO_VIRUS7.Checked := False; ckYSNO_VIRUS8.Checked := False;

     edtDesc_Gum1.Checked := False; edtDesc_Gum2.Checked := False; edtGum_Chk1.Checked := False;
     edtGum_Chk2.Checked  := False;

     edtGum_Form1.Checked := False; edtGum_Form2.Checked := False; edtGum_Form3.Checked := False;
     edtSun_Form1.Checked := False; edtSun_Form2.Checked := False; edtSun_Form3.Checked := False;

     edtSang_Form1.Checked := False; edtSang_Form2.Checked := False; edtSang_Form3.Checked := False;
     edtSang_Form4.Checked := False;

     qryCell.Active := False;
     qryCell.Filter := '';
     qryCell.ParamByName('num_jumin').AsString := edtJumin.Text;
     qryCell.ParamByName('Dat_gmgn').AsString  := copy(cmb_gmgn.text,1,4) + copy(cmb_gmgn.text,6,2) + copy(cmb_gmgn.text,9,2);
     qryCell.Active := True;

     if qryCell.IsEmpty = False then
     begin
        cmbCell_Hangmok.Items.Clear;
        while Not qryCell.EOF do
        begin
           cmbCell_Hangmok.Items.Add(qryCell.FieldByName('cod_hm').AsString);
           qryCell.Next;
        end;

        // PAP�ڷ� Query
        cmbCell_Hangmok.ItemIndex := 0;
        UP_PAPQuery(cmbCell_Hangmok.Text);
     end;
end;


procedure TfrmCK5I17.UP_PAPQuery(sCod_hm : String);
var i : Integer;
begin
   qryCell.Filter := ' cod_hm = ''' + sCod_hm + '''';

   if qryCell.IsEmpty = False then
   begin
      if length(qryCell.FieldByname('Gubn_class').AsString) > 1 then
         GP_ComboDisplay(cmbGUBN_CLASS, qryCell.FieldByname('Gubn_class').AsString, 2)
      else
         GP_ComboDisplay(cmbGUBN_CLASS, qryCell.FieldByname('Gubn_class').AsString, 1);

      if copy(qryCell.FieldByname('cod_hm').AsString,1,1) = 'P' then
      begin
         if qryCell.FieldByName('dat_gmgn').AsString < '20101012' then
         begin
            if   qryCell.FieldByName('cod_pan').AsString = '1' then
               JPan13.Caption := GF_PanComparison(JPan13.Caption, 'A')
            else if qryCell.FieldByName('cod_pan').AsString = '2' then
            begin
               if (JPan13.Caption = '' ) or (JPan13.Caption = 'A') or
                  (JPan13.Caption = 'B') or (copy(JPan13.Caption,1,1) = 'C') then JPan13.Caption := 'C';
            end
            else if (qryCell.FieldByName('cod_pan').AsString = '3') or (qryCell.FieldByName('cod_pan').AsString = '4') or
                    (qryCell.FieldByName('cod_pan').AsString = '5') then
               JPan13.Caption := GF_PanComparison(JPan13.Caption, 'D(0)');
         end
         else if qryCell.FieldByName('dat_gmgn').AsString >= '20101012' then
         begin
            JPan13.Caption := GF_PanComparison(JPan13.Caption, qryCell.FieldByName('gubn_pan').AsString);
         end;
      end
      else if qryCell.FieldByname('cod_hm').AsString = 'B009' then              //20150601
      begin
         JPan13.Caption := GF_PanComparison(JPan13.Caption, qryCell.FieldByName('gubn_pan').AsString);
      end;

      ckYSNO_VIRUS1.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus1').AsString);
      ckYSNO_VIRUS2.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus2').AsString);
      ckYSNO_VIRUS3.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus3').AsString);
      ckYSNO_VIRUS4.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus4').AsString);
      ckYSNO_VIRUS5.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus5').AsString);
      If Trim(qryCell.FieldByname('Ysno_virus7').AsString) = '' Then ckYSno_virus6.Checked := True
      Else                                                   ckYSNO_VIRUS6.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus6').AsString);
      ckYSNO_VIRUS7.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus7').AsString);
      ckYSNO_VIRUS8.Checked := GF_YN2B(qryCell.FieldByname('Ysno_virus8').AsString);
      edtDESC_VIRUS.Text    := qryCell.FieldByname('Desc_virus').AsString;

      edtDESC_BUWI.Text     := qryCell.FieldByname('Desc_buwi').AsString;

      GP_ComboDisplay(cmbCOD_Doct,  qryCell.FieldByname('Cod_Doct').AsString, 6);

      If qryCell.FieldByname('Desc_Gum').AsString = '1' Then
      Begin
         edtDesc_Gum1.Checked := true;
         edtDesc_Gum2.Checked := False;
      End
      Else If qryCell.FieldByname('Desc_Gum').AsString = '2' Then
      Begin
         edtDesc_Gum1.Checked := False;
         edtDesc_Gum2.Checked := True;
      End;

      If qryCell.FieldByname('Gum_Chk').AsString = '1' Then
      Begin
         edtGum_Chk1.Checked := true;
         edtGum_Chk2.Checked := False;
      End
      Else If qryCell.FieldByname('Gum_Chk').AsString = '2' Then
      Begin
         edtGum_Chk1.Checked := False;
         edtGum_Chk2.Checked := True;
      End;

      edtGum_Remark.Text    := qryCell.FieldByname('Gum_Remark').AsString;

      edtGum_Form1.Checked  := GF_YN2B(qryCell.FieldByname('Gum_Form1').AsString);
      edtGum_Form2.Checked  := GF_YN2B(qryCell.FieldByname('Gum_Form2').AsString);
      edtGum_Form3.Checked  := GF_YN2B(qryCell.FieldByname('Gum_Form3').AsString);
      edtGum_Form3_Etc.Text := qryCell.FieldByname('Gum_Form3_Etc').AsString;
      edtSang_Form1.Checked := GF_YN2B(qryCell.FieldByname('Sang_Form1').AsString);

      if qryCell.FieldByname('Sang_Form1_Gubn').AsString <> '' then
         CmbSang_Form1_gubn.ItemIndex := StrToInt(qryCell.FieldByname('Sang_Form1_Gubn').AsString)
      else
         CmbSang_Form1_gubn.ItemIndex := 0;
      edtSang_Form2.Checked := GF_YN2B(qryCell.FieldByname('Sang_Form2').AsString);
      edtSang_Form3.Checked := GF_YN2B(qryCell.FieldByname('Sang_Form3').AsString);
      edtSang_Form4.Checked := GF_YN2B(qryCell.FieldByname('Sang_Form4').AsString);

      edtSun_Form1.Checked  := GF_YN2B(qryCell.FieldByname('Sun_Form1').AsString);
      edtSun_Form2.Checked  := GF_YN2B(qryCell.FieldByname('Sun_Form2').AsString);
      edtSun_Form3.Checked  := GF_YN2B(qryCell.FieldByname('Sun_Form3').AsString);
      edtSun_Form4.Text     := qryCell.FieldByname('Sun_Form4').AsString;

      mskDAT_RESULT.Text    := qryCell.FieldByname('Dat_result').AsString;
      edtCod_Sokun.Text     := UpperCase(qryCell.FieldByname('Cod_sokun').AsString);

      retDESC_KGBL.Text     := qryCell.FieldByname('Desc_kgbl').AsString;
      retDESC_SOKUN.Text    := qryCell.FieldByname('Desc_sokun1').AsString + qryCell.FieldByname('Desc_sokun2').AsString +
                               qryCell.FieldByname('Desc_sokun3').AsString + qryCell.FieldByname('Desc_sokun4').AsString +
                               qryCell.FieldByname('Desc_sokun5').AsString;
      with qryCell_sokun do
      begin
          qryCell_sokun.Active := False;
          ParamByName('num_id').AsString   := qryCell.FieldByname('num_id').AsString;
          ParamByName('dat_gmgn').AsString := qryCell.FieldByname('dat_gmgn').AsString;
          ParamByName('cod_cell').AsString := qryCell.FieldByname('cod_cell').AsString;
          qryCell_sokun.Active := True;

          if RecordCount >0 then
             retDESC_SOKUN1.Text    :=  qryCell_sokun.FieldByname('Desc_sokun').AsString + qryCell_sokun.FieldByname('Desc_sokun1').AsString
                                        + qryCell_sokun.FieldByname('Desc_sokun2').AsString;

          qryCell_sokun.Active := False;
      end;

      edtGUBN_PAN.Text      := qryCell.FieldByname('Gubn_pan').AsString;

      if qryCell.FieldByname('dat_gmgn').AsString >= '20070701' then
      begin
         if      Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'C' then Panel_PAP.Visible := False //�ǵ��Ϸ�
         else if Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'Y' then Panel_PAP.Visible := True  //*�ǵ���*
         else Panel_PAP.Visible := True;                                                                //���ǵ�
      end
      else
      begin
         if (Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'Y') or
            (Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'C') then Panel_PAP.Visible := False     //�ǵ��Ϸ�
         else
         begin
            if qryCell.FieldByname('Dat_result').AsString <> '' then Panel_PAP.Visible := False          //�ǵ��Ϸ�
            else Panel_PAP.Visible := True;                                                              //���ǵ�
         end;
      end;
   end;
end;

procedure TfrmCK5I17.CreateTextArrayTot_Munjin;
var
   i : Integer;
   sMunjinTemp, sMunjinTemp2 : string;   
begin

   //[2019.05.29-������]2019�⵵ ���� ����(���κ�)  ..2019.06.12 �߰� (���ξ�����)
   //---------------------------------------------------------------------------
   Mun4_2019.itemindex     := 0;
   Mun4_1_2019.itemindex   := 0;
   Mun4_2_2019.itemindex   := 0;
   Mun4_2_1_2019.itemindex := 0;
   Mun5_2019.itemindex     := 0;
   Mun5_1_2019.itemindex   := 0;

   Mun4_1_Year_2019.Text  := '';
   Mun4_1_Day_2019.Text   := '';
   Mun4_1_PYear_2019.Text := '';
   Mun4_2_Year_2019.Text  := '';
   Mun4_2_Day_2019.Text   := '';
   Mun4_2_PYear_2019.Text := '';
   //===========================================================================
   // �ǰ����� ����
   Mun1_J_YN.itemindex := 0;
   Mun1_D_YN.itemindex := 0;
   Ck_Mun1_J1.Checked := false;   Ck_Mun1_J2.Checked := false;
   Ck_Mun1_J3.Checked := false;   Ck_Mun1_J4.Checked := false;
   Ck_Mun1_J5.Checked := false;   Ck_Mun1_J6.Checked := false;
   Ck_Mun1_J7.Checked := false;   Ck_Mun1_J8.Checked := false;

   Ck_Mun1_D1.Checked := false;   Ck_Mun1_D2.Checked := false;
   Ck_Mun1_D3.Checked := false;   Ck_Mun1_D4.Checked := false;
   Ck_Mun1_D5.Checked := false;   Ck_Mun1_D6.Checked := false;
   Ck_Mun1_D7.Checked := false;   Ck_Mun1_D8.Checked := false;

   Mun2_YN.itemindex  := 0;
   Ck_Mun2_1.Checked  := false;   Ck_Mun2_2.Checked  := false;
   Ck_Mun2_3.Checked  := false;   Ck_Mun2_4.Checked  := false;
   Ck_Mun2_5.Checked  := false;   Ck_Mun2_6.Checked  := false;

   Mun3.ItemIndex   := 0;
   Mun4_1.ItemIndex := 0;
   Mun4_2_Year.Text := '';
   Mun4_2_Day.Text  := '';
   Mun4_3_Year.Text := '';
   Mun4_3_Day.Text  := '';

   Mun5_1.ItemIndex := 0;
   Mun5_2.ItemIndex := 0;
   Mun5_3_1.ItemIndex := 0;
   Mun5_3_2.Text    := '';
   Mun5_4.Text      := '';

   Mun6_1.ItemIndex := 0;
   Mun6_2.ItemIndex := 0;
   Mun6_3.ItemIndex := 0;
   Mun6_4.ItemIndex := 0;

   Mun8_1.ItemIndex := 0;
   Mun8_2.ItemIndex := 0;
   Mun8_3.ItemIndex := 0;
   Mun8_4.ItemIndex := 0;

   //Ư���Ϲ���
   CMun1.ItemIndex := 0;
   CMun1_Bung.Text := '';
   CMun2.ItemIndex := 0;
   CMun2_kg.Text   := '';

   CMun3_Can1_Yn.ItemIndex := 0;
   CMun3_Can2_Yn.ItemIndex := 0;
   CMun3_Can3_Yn.ItemIndex := 0;
   CMun3_Can4_Yn.ItemIndex := 0;
   CMun3_Can5_Yn.ItemIndex := 0;
   CMun3_Can6_Yn.ItemIndex := 0;
   CMun3_Can6_Bung.Text := '';

   CMun3_Can1_F1.Checked := false;
   CMun3_Can1_F2.Checked := false;
   CMun3_Can1_F3.Checked := false;
   CMun3_Can1_F4.Checked := false;
   CMun3_Can1_F5.Checked := false;

   CMun3_Can2_F1.Checked := false;
   CMun3_Can2_F2.Checked := false;
   CMun3_Can2_F3.Checked := false;
   CMun3_Can2_F4.Checked := false;
   CMun3_Can2_F5.Checked := false;

   CMun3_Can3_F1.Checked := false;
   CMun3_Can3_F2.Checked := false;
   CMun3_Can3_F3.Checked := false;
   CMun3_Can3_F4.Checked := false;
   CMun3_Can3_F5.Checked := false;

   CMun3_Can4_F1.Checked := false;
   CMun3_Can4_F2.Checked := false;
   CMun3_Can4_F3.Checked := false;
   CMun3_Can4_F4.Checked := false;
   CMun3_Can4_F5.Checked := false;

   CMun3_Can5_F1.Checked := false;
   CMun3_Can5_F2.Checked := false;
   CMun3_Can5_F3.Checked := false;
   CMun3_Can5_F4.Checked := false;
   CMun3_Can5_F5.Checked := false;

   CMun3_Can6_F1.Checked := false;
   CMun3_Can6_F2.Checked := false;
   CMun3_Can6_F3.Checked := false;
   CMun3_Can6_F4.Checked := false;
   CMun3_Can6_F5.Checked := false;

   CMun4_Can1.ItemIndex := 0;
   CMun4_Can2.ItemIndex := 0;
   CMun4_Can3.ItemIndex := 0;
   CMun4_Can4.ItemIndex := 0;
   CMun4_Can5.ItemIndex := 0;
   CMun4_Can6.ItemIndex := 0;
   CMun4_Can7.ItemIndex := 0;
   CMun4_Can8.ItemIndex := 0;

   CMun5_YN.ItemIndex := 0;
   CMun5_1.Checked := false;
   CMun5_2.Checked := false;
   CMun5_3.Checked := false;
   CMun5_4.Checked := false;
   CMun5_5.Checked := false;

   CMun6_YN.ItemIndex := 0;
   CMun6_1.Checked := false;
   CMun6_2.Checked := false;
   CMun6_3.Checked := false;
   CMun6_4.Checked := false;
   CMun6_5.Checked := false;

   CMun7_YN.ItemIndex := 0;
   CMun7_1.Checked := false;
   CMun7_2.Checked := false;
   CMun7_3.Checked := false;
   CMun7_4.Checked := false;
   CMun7_5.Checked := false;

  // �������� Ư���Ϲ���...
   CMun8.ItemIndex  := 0;
   CMun8_Year.Text  := '';
   CMun9.ItemIndex  := 0;
   CMun9_Year.Text  := '';
   CMun10.ItemIndex := 0;
   CMun11.ItemIndex := 0;
   CMun12.ItemIndex := 0;
   CMun13.ItemIndex := 0;
   CMun14.ItemIndex := 0;

   cmbYsno_agree.ItemIndex := 0;


   with qryTot_Munjin2010 do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString  := copy(cmb_gmgn.text,12,2);
      ParamByName('num_jumin').AsString := edtJumin.Text;
      ParamByName('dat_gmgn').AsString  := copy(cmb_gmgn.text,1,4) + copy(cmb_gmgn.text,6,2) + copy(cmb_gmgn.text,9,2);
      Active := True;

     // if cmb_gmgn.text <= '20100101' then
     //    Panel41.Visible := True
     // else Panel41.Visible := False;

      if IsEmpty = False then
      begin
          //[2017.06.20-������]���հ���� ���ŷ� ǥ��....
          //--------------------------------------------------------------------
          Mm_munjin2.Lines.Clear;
          sMunjinTemp := '';
          if FieldByName('mun1_jindan1').AsString = '1' then sMunjinTemp := sMunjinTemp + '������(��ǳ),';
          if FieldByName('mun1_jindan2').AsString = '1' then sMunjinTemp := sMunjinTemp + '���庴(�ɱٰ��/������),';
          if FieldByName('mun1_jindan3').AsString = '1' then sMunjinTemp := sMunjinTemp + '������,';
          if FieldByName('mun1_jindan4').AsString = '1' then sMunjinTemp := sMunjinTemp + '�索��,';
          if FieldByName('mun1_jindan5').AsString = '1' then sMunjinTemp := sMunjinTemp + '�̻���������,';
          if FieldByName('mun1_jindan6').AsString = '1' then sMunjinTemp := sMunjinTemp + '�����,';
          if FieldByName('mun1_jindan7').AsString = '1' then
          begin
             sMunjinTemp2 := '';
             if copy(FieldByName('CMun3_Can1_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '����,';
             if copy(FieldByName('CMun3_Can2_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�����,';
             if copy(FieldByName('CMun3_Can3_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�����,';
             if copy(FieldByName('CMun3_Can4_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '����,';
             if copy(FieldByName('CMun3_Can5_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�ڱð�ξ�,';
             if copy(FieldByName('CMun3_Can6_Family').AsString,1,1) = '2' then
             begin
                if pos('��',FieldByName('CMun3_Can6_Bung').AsString) > 0 then sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + ','
                else                                                          sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + '��,';
             end;

             if sMunjinTemp2 = '' then sMunjinTemp := sMunjinTemp + '��,'
             else                      sMunjinTemp := sMunjinTemp + sMunjinTemp2;
          end;
          if FieldByName('mun1_jindan8').AsString = '1' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then  Mm_munjin2.Lines.Add('��������(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun5').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '���˾�,';
          if copy(FieldByName('CMun5').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '���༺����,';
          if copy(FieldByName('CMun5').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '�����ȭ��,';
          if copy(FieldByName('CMun5').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '������,';
          if copy(FieldByName('CMun5').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('��������ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun6').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '�������,';
          if copy(FieldByName('CMun6').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '�˾缺���忰,';
          if copy(FieldByName('CMun6').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + 'ũ�к�,';
          if copy(FieldByName('CMun6').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + 'ġ��(ġ��,ġ��),';
          if copy(FieldByName('CMun6').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('�������׹���ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun7').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + 'B������������,';
          if copy(FieldByName('CMun7').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '����B������,';
          if copy(FieldByName('CMun7').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '����C������,';
          if copy(FieldByName('CMun7').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '���溯,';
          if copy(FieldByName('CMun7').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('������ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');
          //====================================================================

          // �ǰ���������.
          //1������ ����
          if FieldByName('Mun1_J_YN').AsString <> ''  then
             Mun1_J_YN.itemindex  := FieldByName('Mun1_J_YN').AsInteger;
          if FieldByName('Mun1_Jindan1').AsString = '1' then
             Ck_Mun1_J1.checked := true else Ck_Mun1_J1.checked := false;
          if FieldByName('Mun1_Jindan2').AsString = '1' then
             Ck_Mun1_J2.checked := true else Ck_Mun1_J2.checked := false;
          if FieldByName('Mun1_Jindan3').AsString = '1' then
             Ck_Mun1_J3.checked := true else Ck_Mun1_J3.checked := false;
          if FieldByName('Mun1_Jindan4').AsString = '1' then
             Ck_Mun1_J4.checked := true else Ck_Mun1_J4.checked := false;
          if FieldByName('Mun1_Jindan5').AsString = '1' then
             Ck_Mun1_J5.checked := true else Ck_Mun1_J5.checked := false;
          if FieldByName('Mun1_Jindan6').AsString = '1' then
             Ck_Mun1_J6.checked := true else Ck_Mun1_J6.checked := false;
          if FieldByName('Mun1_Jindan7').AsString = '1' then
             Ck_Mun1_J7.checked := true else Ck_Mun1_J7.checked := false;
          if FieldByName('Mun1_Jindan8').AsString = '1' then
             Ck_Mun1_J8.checked := true else Ck_Mun1_J8.checked := false;

          //1������ �๰
          if FieldByName('Mun1_D_YN').AsString <> ''  then
             Mun1_D_YN.itemindex  := FieldByName('Mun1_D_YN').AsInteger;
          if FieldByName('Mun1_Drug1').AsString = '1' then
             Ck_Mun1_D1.checked := true else Ck_Mun1_D1.checked := false;
          if FieldByName('Mun1_Drug2').AsString = '1' then
             Ck_Mun1_D2.checked := true else Ck_Mun1_D2.checked := false;
          if FieldByName('Mun1_Drug3').AsString = '1' then
          begin
             Ck_Mun1_D3.checked := true;
             JPan6.Caption      := 'T'
          end
          else Ck_Mun1_D3.checked := false;
          if FieldByName('Mun1_Drug4').AsString = '1' then
             Ck_Mun1_D4.checked := true else Ck_Mun1_D4.checked := false;
          if FieldByName('Mun1_Drug5').AsString = '1' then
             Ck_Mun1_D5.checked := true else Ck_Mun1_D5.checked := false;
          if FieldByName('Mun1_Drug6').AsString = '1' then
             Ck_Mun1_D6.checked := true else Ck_Mun1_D6.checked := false;
          if FieldByName('Mun1_Drug7').AsString = '1' then
             Ck_Mun1_D7.checked := true else Ck_Mun1_D7.checked := false;
          if FieldByName('Mun1_Drug8').AsString = '1' then
             Ck_Mun1_D8.checked := true else Ck_Mun1_D8.checked := false;

          //2������
          if FieldByName('Mun2_YN').AsString <> ''  then
             Mun2_YN.itemindex  := FieldByName('Mun2_YN').AsInteger;
          if FieldByName('Mun2_1').AsString = '1' then
             Ck_Mun2_1.checked := true else Ck_Mun2_1.checked := false;
          if FieldByName('Mun2_2').AsString = '1' then
             Ck_Mun2_2.checked := true else Ck_Mun2_2.checked := false;
          if FieldByName('Mun2_3').AsString = '1' then
             Ck_Mun2_3.checked := true else Ck_Mun2_3.checked := false;
          if FieldByName('Mun2_4').AsString = '1' then
             Ck_Mun2_4.checked := true else Ck_Mun2_4.checked := false;
          if FieldByName('Mun2_5').AsString = '1' then
             Ck_Mun2_5.checked := true else Ck_Mun2_5.checked := false;
          if FieldByName('Mun2_6').AsString = '1' then
             Ck_Mun2_6.checked := true else Ck_Mun2_6.checked := false;

          //3������ B������
          if FieldByName('Mun3').AsString <> ''  then
             Mun3.itemindex  := FieldByName('Mun3').AsInteger;

          //4������
          if FieldByName('Mun4_1').AsString <> ''  then
             Mun4_1.itemindex  := FieldByName('Mun4_1').AsInteger;
          Mun4_2_Year.Text := FieldByName('Mun4_2_Year').AsString;
          Mun4_2_Day.Text  := FieldByName('Mun4_2_Day').AsString;
          Mun4_3_Year.Text := FieldByName('Mun4_3_Year').AsString;
          Mun4_3_Day.Text  := FieldByName('Mun4_3_Day').AsString;

          //5������
          if FieldByName('Mun5_1').AsString <> ''  then
             Mun5_1.itemindex  := FieldByName('Mun5_1').AsInteger;
          if FieldByName('Mun5_2').AsString <> ''  then
             Mun5_2.itemindex  := FieldByName('Mun5_2').AsInteger;
          if FieldByName('Mun5_3_1').AsString <> '' then
             Mun5_3_1.itemindex  := FieldByName('Mun5_3_1').AsInteger;
          Mun5_3_2.Text  := FieldByName('Mun5_3_2').AsString;
          Mun5_4.Text    := FieldByName('Mun5_4').AsString;

          //6������
          if FieldByName('Mun6_1').AsString <> ''  then
             Mun6_1.itemindex  := FieldByName('Mun6_1').AsInteger;
          if FieldByName('Mun6_2').AsString <> ''  then
             Mun6_2.itemindex  := FieldByName('Mun6_2').AsInteger;
          if FieldByName('Mun6_3').AsString <> ''  then
             Mun6_3.itemindex  := FieldByName('Mun6_3').AsInteger;
          if FieldByName('Mun6_4').AsString <> ''  then
             Mun6_4.itemindex  := FieldByName('Mun6_4').AsInteger;

          //8������
          if FieldByName('Mun8_1').AsString <> ''  then
             Mun8_1.itemindex  := FieldByName('Mun8_1').AsInteger;
          if FieldByName('Mun8_2').AsString <> ''  then
             Mun8_2.itemindex  := FieldByName('Mun8_2').AsInteger;
          if FieldByName('Mun8_3').AsString <> ''  then
             Mun8_3.itemindex  := FieldByName('Mun8_3').AsInteger;
          if FieldByName('Mun8_4').AsString <> ''  then
             Mun8_4.itemindex  := FieldByName('Mun8_4').AsInteger;

          // Ư���Ϲ���
          // 1������
          if FieldByName('CMun1').AsString <> ''  then
             CMun1.itemindex  := FieldByName('CMun1').AsInteger;
          CMun1_Bung.Text := FieldByName('CMun1_Bung').AsString;

          // 2������
          if FieldByName('CMun2').AsString <> ''  then
             CMun2.itemindex  := FieldByName('CMun2').AsInteger;
          CMun2_kg.Text := FieldByName('CMun2_kg').AsString;

          // 3������
          //����
          if FieldByName('CMun3_Can1_Yn').AsString <> ''  then
             CMun3_Can1_Yn.itemindex  := FieldByName('CMun3_Can1_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 1, 1) = '2' then
             CMun3_Can1_F1.checked := true else CMun3_Can1_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 2, 1) = '2' then
             CMun3_Can1_F2.checked := true else CMun3_Can1_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 3, 1) = '2' then
             CMun3_Can1_F3.checked := true else CMun3_Can1_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 4, 1) = '2' then
             CMun3_Can1_F4.checked := true else CMun3_Can1_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 5, 1) = '2' then
             CMun3_Can1_F5.checked := true else CMun3_Can1_F5.checked := false;

          //�����
          if FieldByName('CMun3_Can2_Yn').AsString <> ''  then
             CMun3_Can2_Yn.itemindex  := FieldByName('CMun3_Can2_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 1, 1) = '2' then
             CMun3_Can2_F1.checked := true else CMun3_Can2_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 2, 1) = '2' then
             CMun3_Can2_F2.checked := true else CMun3_Can2_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 3, 1) = '2' then
             CMun3_Can2_F3.checked := true else CMun3_Can2_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 4, 1) = '2' then
             CMun3_Can2_F4.checked := true else CMun3_Can2_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 5, 1) = '2' then
             CMun3_Can2_F5.checked := true else CMun3_Can2_F5.checked := false;

          //�����
          if FieldByName('CMun3_Can3_Yn').AsString <> ''  then
             CMun3_Can3_Yn.itemindex  := FieldByName('CMun3_Can3_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 1, 1) = '2' then
             CMun3_Can3_F1.checked := true else CMun3_Can3_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 2, 1) = '2' then
             CMun3_Can3_F2.checked := true else CMun3_Can3_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 3, 1) = '2' then
             CMun3_Can3_F3.checked := true else CMun3_Can3_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 4, 1) = '2' then
             CMun3_Can3_F4.checked := true else CMun3_Can3_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 5, 1) = '2' then
             CMun3_Can3_F5.checked := true else CMun3_Can3_F5.checked := false;

          //����
          if FieldByName('CMun3_Can4_Yn').AsString <> ''  then
             CMun3_Can4_Yn.itemindex  := FieldByName('CMun3_Can4_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 1, 1) = '2' then
             CMun3_Can4_F1.checked := true else CMun3_Can4_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 2, 1) = '2' then
             CMun3_Can4_F2.checked := true else CMun3_Can4_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 3, 1) = '2' then
             CMun3_Can4_F3.checked := true else CMun3_Can4_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 4, 1) = '2' then
             CMun3_Can4_F4.checked := true else CMun3_Can4_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 5, 1) = '2' then
             CMun3_Can4_F5.checked := true else CMun3_Can4_F5.checked := false;

          //�ڱð�ξ�
          if FieldByName('CMun3_Can5_Yn').AsString <> ''  then
             CMun3_Can5_Yn.itemindex  := FieldByName('CMun3_Can5_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 1, 1) = '2' then
             CMun3_Can5_F1.checked := true else CMun3_Can5_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 2, 1) = '2' then
             CMun3_Can5_F2.checked := true else CMun3_Can5_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 3, 1) = '2' then
             CMun3_Can5_F3.checked := true else CMun3_Can5_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 4, 1) = '2' then
             CMun3_Can5_F4.checked := true else CMun3_Can5_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 5, 1) = '2' then
             CMun3_Can5_F5.checked := true else CMun3_Can5_F5.checked := false;

          //��Ÿ��
          CMun3_Can6_Bung.Text := FieldByName('CMun3_Can6_Bung').AsString;
          if FieldByName('CMun3_Can6_Yn').AsString <> ''  then
             CMun3_Can6_Yn.itemindex  := FieldByName('CMun3_Can6_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 1, 1) = '2' then
             CMun3_Can6_F1.checked := true else CMun3_Can6_F1.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 2, 1) = '2' then
             CMun3_Can6_F2.checked := true else CMun3_Can6_F2.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 3, 1) = '2' then
             CMun3_Can6_F3.checked := true else CMun3_Can6_F3.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 4, 1) = '2' then
             CMun3_Can6_F4.checked := true else CMun3_Can6_F4.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 5, 1) = '2' then
             CMun3_Can6_F5.checked := true else CMun3_Can6_F5.checked := false;

          // 4������
          if FieldByName('CMun4_Can1').AsString <> ''  then
             CMun4_Can1.itemindex  := FieldByName('CMun4_Can1').AsInteger;
          if FieldByName('CMun4_Can2').AsString <> ''  then
             CMun4_Can2.itemindex  := FieldByName('CMun4_Can2').AsInteger;
          if FieldByName('CMun4_Can3').AsString <> ''  then
             CMun4_Can3.itemindex  := FieldByName('CMun4_Can3').AsInteger;
          if FieldByName('CMun4_Can4').AsString <> ''  then
             CMun4_Can4.itemindex  := FieldByName('CMun4_Can4').AsInteger;
          if FieldByName('CMun4_Can5').AsString <> ''  then
             CMun4_Can5.itemindex  := FieldByName('CMun4_Can5').AsInteger;
          if FieldByName('CMun4_Can6').AsString <> ''  then
             CMun4_Can6.itemindex  := FieldByName('CMun4_Can6').AsInteger;
          if FieldByName('CMun4_Can7').AsString <> ''  then
             CMun4_Can7.itemindex  := FieldByName('CMun4_Can7').AsInteger;
          if FieldByName('CMun4_Can8').AsString <> ''  then
             CMun4_Can8.itemindex  := FieldByName('CMun4_Can8').AsInteger;

          // 5������
          if FieldByName('CMun5_YN').AsString <> ''  then
             CMun5_YN.itemindex := FieldByName('CMun5_YN').AsInteger;
          if Copy(FieldByName('CMun5').AsString, 1, 1) = '2' then
             CMun5_1.checked := true else CMun5_1.checked := false;
          if Copy(FieldByName('CMun5').AsString, 2, 1) = '2' then
             CMun5_2.checked := true else CMun5_2.checked := false;
          if Copy(FieldByName('CMun5').AsString, 3, 1) = '2' then
             CMun5_3.checked := true else CMun5_3.checked := false;
          if Copy(FieldByName('CMun5').AsString, 4, 1) = '2' then
             CMun5_4.checked := true else CMun5_4.checked := false;
          if Copy(FieldByName('CMun5').AsString, 5, 1) = '2' then
             CMun5_5.checked := true else CMun5_5.checked := false;

          // 6������
          if FieldByName('CMun6_YN').AsString <> ''  then
             CMun6_YN.itemindex := FieldByName('CMun6_YN').AsInteger;
          if Copy(FieldByName('CMun6').AsString, 1, 1) = '2' then
             CMun6_1.checked := true else CMun6_1.checked := false;
          if Copy(FieldByName('CMun6').AsString, 2, 1) = '2' then
             CMun6_2.checked := true else CMun6_2.checked := false;
          if Copy(FieldByName('CMun6').AsString, 3, 1) = '2' then
             CMun6_3.checked := true else CMun6_3.checked := false;
          if Copy(FieldByName('CMun6').AsString, 4, 1) = '2' then
             CMun6_4.checked := true else CMun6_4.checked := false;
          if Copy(FieldByName('CMun6').AsString, 5, 1) = '2' then
             CMun6_5.checked := true else CMun6_5.checked := false;

          // 7������
          if FieldByName('CMun7_YN').AsString <> ''  then
             CMun7_YN.itemindex := FieldByName('CMun7_YN').AsInteger;
          if Copy(FieldByName('CMun7').AsString, 1, 1) = '2' then
             CMun7_1.checked := true else CMun7_1.checked := false;
          if Copy(FieldByName('CMun7').AsString, 2, 1) = '2' then
             CMun7_2.checked := true else CMun7_2.checked := false;
          if Copy(FieldByName('CMun7').AsString, 3, 1) = '2' then
             CMun7_3.checked := true else CMun7_3.checked := false;
          if Copy(FieldByName('CMun7').AsString, 4, 1) = '2' then
             CMun7_4.checked := true else CMun7_4.checked := false;
          if Copy(FieldByName('CMun7').AsString, 5, 1) = '2' then
             CMun7_5.checked := true else CMun7_5.checked := false;

          // ���������ڸ� �ش�

          // 8������
          if FieldByName('CMun8').AsString <> ''  then
             CMun8.itemindex  := FieldByName('CMun8').AsInteger;
          CMun8_Year.Text := FieldByName('CMun8_Year').AsString;

          // 9������
          if FieldByName('CMun9').AsString <> ''  then
             CMun9.itemindex  := FieldByName('CMun9').AsInteger;
          CMun9_Year.Text := FieldByName('CMun9_Year').AsString;

          // 10������
          if FieldByName('CMun10').AsString <> ''  then
             CMun10.itemindex  := FieldByName('CMun10').AsInteger;

          // 11������
          if FieldByName('CMun11').AsString <> ''  then
             CMun11.itemindex  := FieldByName('CMun11').AsInteger;

          // 12������
          if FieldByName('CMun12').AsString <> ''  then
             CMun12.itemindex  := FieldByName('CMun12').AsInteger;

          // 13������
          if FieldByName('CMun13').AsString <> ''  then
             CMun13.itemindex  := FieldByName('CMun13').AsInteger;

          // 14������
          if FieldByName('CMun14').AsString <> ''  then
             CMun14.itemindex  := FieldByName('CMun14').AsInteger;

          //����Ȱ�� ���ǿ���
          if FieldByName('Ysno_agree').AsString <> '' then
             cmbYsno_agree.itemindex := FieldByname('Ysno_agree').AsInteger;
      end;
      Active := False;
   end;
end;

procedure TfrmCK5I17.CreateTextArrayTot_Munjin2018;
var
   i : Integer;
   sMunjinTemp, sMunjinTemp2 : string;   
begin
   // �ǰ����� ����
   Mun1_J_YN_2018.itemindex := 0;
   Mun1_D_YN_2018.itemindex := 0;
   Ck_Mun1_J1_2018.Checked := false;   Ck_Mun1_J2_2018.Checked := false;
   Ck_Mun1_J3_2018.Checked := false;   Ck_Mun1_J4_2018.Checked := false;
   Ck_Mun1_J5_2018.Checked := false;   Ck_Mun1_J6_2018.Checked := false;
   Ck_Mun1_J7_2018.Checked := false;  

   Ck_Mun1_D1_2018.Checked := false;   Ck_Mun1_D2_2018.Checked := false;
   Ck_Mun1_D3_2018.Checked := false;   Ck_Mun1_D4_2018.Checked := false;
   Ck_Mun1_D5_2018.Checked := false;   Ck_Mun1_D6_2018.Checked := false;
   Ck_Mun1_D7_2018.Checked := false;   

   Mun2_YN_2018.itemindex  := 0;
   Ck_Mun2_1_2018.Checked  := false;   Ck_Mun2_2_2018.Checked  := false;
   Ck_Mun2_3_2018.Checked  := false;   Ck_Mun2_4_2018.Checked  := false;
   Ck_Mun2_5_2018.Checked  := false;  
   
   Mun3_2018.itemindex  := 0;
   Mun4_2018.itemindex  := 0;
   
   Mun4_1_Year_2018.Text  := '';
   Mun4_1_Day_2018.Text  := '';
   Mun4_2_Year_2018.Text  := '';
   Mun4_2_Day_2018.Text  := '';
   
   Mun5_2018.itemindex  := 0;
   Mun5_1_2018.itemindex  := 0;
   Mun6_num_2018.itemindex  := 0;
   Mun6_count_2018.Text  := '';
   Mun6_1_2018.Text  := '';
   Mun6_2_2018.Text  := '';
   
   Mun7_1_2018.itemindex  := 0;
   Mun7_2_H_2018.Text  := '';
   Mun7_2_M_2018.Text  := '';
   Mun8_1_2018.itemindex  := 0;
   Mun8_2_H_2018.Text  := '';
   Mun8_2_M_2018.Text  := '';
   Mun9_2018.itemindex  := 0;


   //Ư���Ϲ���
   CMun1_2018.ItemIndex := 0;
   CMun1_Bung_2018.Text := '';
   CMun2_2018.ItemIndex := 0;
   CMun2_kg_2018.Text   := '';

   CMun3_Can1_Yn_2018.ItemIndex := 0;
   CMun3_Can2_Yn_2018.ItemIndex := 0;
   CMun3_Can3_Yn_2018.ItemIndex := 0;
   CMun3_Can4_Yn_2018.ItemIndex := 0;
   CMun3_Can5_Yn_2018.ItemIndex := 0;
   CMun3_Can6_Yn_2018.ItemIndex := 0;
   CMun3_Can6_Bung_2018.Text := '';

   CMun3_Can1_F1_2018.Checked := false;
   CMun3_Can1_F2_2018.Checked := false;
   CMun3_Can1_F3_2018.Checked := false;
   CMun3_Can1_F4_2018.Checked := false;
   CMun3_Can1_F5_2018.Checked := false;

   CMun3_Can2_F1_2018.Checked := false;
   CMun3_Can2_F2_2018.Checked := false;
   CMun3_Can2_F3_2018.Checked := false;
   CMun3_Can2_F4_2018.Checked := false;
   CMun3_Can2_F5_2018.Checked := false;

   CMun3_Can3_F1_2018.Checked := false;
   CMun3_Can3_F2_2018.Checked := false;
   CMun3_Can3_F3_2018.Checked := false;
   CMun3_Can3_F4_2018.Checked := false;
   CMun3_Can3_F5_2018.Checked := false;

   CMun3_Can4_F1_2018.Checked := false;
   CMun3_Can4_F2_2018.Checked := false;
   CMun3_Can4_F3_2018.Checked := false;
   CMun3_Can4_F4_2018.Checked := false;
   CMun3_Can4_F5_2018.Checked := false;

   CMun3_Can5_F1_2018.Checked := false;
   CMun3_Can5_F2_2018.Checked := false;
   CMun3_Can5_F3_2018.Checked := false;
   CMun3_Can5_F4_2018.Checked := false;
   CMun3_Can5_F5_2018.Checked := false;

   CMun3_Can6_F1_2018.Checked := false;
   CMun3_Can6_F2_2018.Checked := false;
   CMun3_Can6_F3_2018.Checked := false;
   CMun3_Can6_F4_2018.Checked := false;
   CMun3_Can6_F5_2018.Checked := false;

   CMun4_Can1_2018.ItemIndex := 0;
   CMun4_Can2_2018.ItemIndex := 0;
   CMun4_Can3_2018.ItemIndex := 0;
   CMun4_Can4_2018.ItemIndex := 0;
   CMun4_Can5_2018.ItemIndex := 0;
   CMun4_Can6_2018.ItemIndex := 0;
   CMun4_Can7_2018.ItemIndex := 0;
   CMun4_Can8_2018.ItemIndex := 0;

   CMun5_YN_2018.ItemIndex := 0;
   CMun5_1_2018.Checked := false;
   CMun5_2_2018.Checked := false;
   CMun5_3_2018.Checked := false;
   CMun5_4_2018.Checked := false;
   CMun5_5_2018.Checked := false;

   CMun6_YN_2018.ItemIndex := 0;
   CMun6_1_2018.Checked := false;
   CMun6_2_2018.Checked := false;
   CMun6_3_2018.Checked := false;
   CMun6_4_2018.Checked := false;
   CMun6_5_2018.Checked := false;

   CMun7_YN_2018.ItemIndex := 0;
   CMun7_1_2018.Checked := false;
   CMun7_2_2018.Checked := false;
   CMun7_3_2018.Checked := false;
   CMun7_4_2018.Checked := false;
   CMun7_5_2018.Checked := false;

  // �������� Ư���Ϲ���...
   CMun8_2018.ItemIndex  := 0;
   CMun8_Year_2018.Text  := '';
   CMun9_2018.ItemIndex  := 0;
   CMun9_Year_2018.Text  := '';
   CMun10_2018.ItemIndex := 0;
   CMun11_2018.ItemIndex := 0;
   CMun12_2018.ItemIndex := 0;
   CMun13_2018.ItemIndex := 0;
   CMun14_2018.ItemIndex := 0;

   cmbYsno_agree_2018.ItemIndex := 0;

   // ������ġ�� ����
   i := grdMaster.CurrentDataRow - 1;

   //[2019.05.29-������]2019�⵵ ���� ����(���κ�)  ..19.06.12
   if copy(cmb_gmgn.text,1,4) = '2018' then
   begin
      Pnl_Mun_2019.Visible := False;

      Label192.Caption := '6. �����ô� Ƚ��';
      Label195.Caption := '6-1. �����ô� ��';
      Label177.Caption := '6-2. �ִ� ���־�';
      Label190.Caption := '7-1. ��� 1���ϰ�, ������ Ȱ�� �ϼ�?    �ִ�               ��';
      Label196.Caption := '7-2. ��� ������ Ȱ�� �ð�?   �Ϸ翡          �ð�          ��';
      Label198.Caption := '8-1. ��� 1���ϰ�, �߰��� Ȱ�� �ϼ�?    �ִ�               ��';
      Label197.Caption := '8-2. ��� �߰��� Ȱ�� �ð�?   �Ϸ翡          �ð�          ��';
      Label199.Caption := '9. �ֱ� �ٷ¿ �� ��?                     �ִ�               ��';
   end
   else if copy(cmb_gmgn.text,1,4) = '2019' then
   begin
      Pnl_Mun_2019.Visible := True;

      Label192.Caption := '7. �����ô� Ƚ��';
      Label195.Caption := '7-1. �����ô� ��';
      Label177.Caption := '7-2. �ִ� ���־�';
      Label190.Caption := '8-1. ��� 1���ϰ�, ������ Ȱ�� �ϼ�?    �ִ�               ��';
      Label196.Caption := '8-2. ��� ������ Ȱ�� �ð�?   �Ϸ翡          �ð�          ��';
      Label198.Caption := '9-1. ��� 1���ϰ�, �߰��� Ȱ�� �ϼ�?    �ִ�               ��';
      Label197.Caption := '9-2. ��� �߰��� Ȱ�� �ð�?   �Ϸ翡          �ð�          ��';
      Label199.Caption := '10. �ֱ� �ٷ¿ �� ��?                    �ִ�               ��';
   end;

   with qryTot_Munjin2018 do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString  := copy(cmb_gmgn.text,12,2);
      ParamByName('num_jumin').AsString := edtJumin.Text;
      ParamByName('dat_gmgn').AsString  := copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2);
      Active := True;

      if IsEmpty = False then
      begin
      
          //[2017.06.20-������]���հ���� ���ŷ� ǥ��....
          //--------------------------------------------------------------------
          Mm_munjin2_2018.Lines.Clear;
          sMunjinTemp := '';
          if FieldByName('mun1_jindan1').AsString = '1' then sMunjinTemp := sMunjinTemp + '������(��ǳ),';
          if FieldByName('mun1_jindan2').AsString = '1' then sMunjinTemp := sMunjinTemp + '���庴(�ɱٰ��/������),';
          if FieldByName('mun1_jindan3').AsString = '1' then sMunjinTemp := sMunjinTemp + '������,';
          if FieldByName('mun1_jindan4').AsString = '1' then sMunjinTemp := sMunjinTemp + '�索��,';
          if FieldByName('mun1_jindan5').AsString = '1' then sMunjinTemp := sMunjinTemp + '�̻���������,';
          if FieldByName('mun1_jindan6').AsString = '1' then sMunjinTemp := sMunjinTemp + '�����,';
          if FieldByName('mun1_jindan7').AsString = '1' then
          begin
             sMunjinTemp2 := '';
             if copy(FieldByName('CMun3_Can1_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '����,';
             if copy(FieldByName('CMun3_Can2_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�����,';
             if copy(FieldByName('CMun3_Can3_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�����,';
             if copy(FieldByName('CMun3_Can4_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '����,';
             if copy(FieldByName('CMun3_Can5_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '�ڱð�ξ�,';
             if copy(FieldByName('CMun3_Can6_Family').AsString,1,1) = '2' then
             begin
                if pos('��',FieldByName('CMun3_Can6_Bung').AsString) > 0 then sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + ','
                else                                                          sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + '��,';
             end;

             if sMunjinTemp2 = '' then sMunjinTemp := sMunjinTemp + '��Ÿ(������),'
             else                      sMunjinTemp := sMunjinTemp + sMunjinTemp2;
          end;
          if Trim(sMunjinTemp) <> '' then  Mm_munjin2_2018.Lines.Add('��������(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun5').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '���˾�,';
          if copy(FieldByName('CMun5').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '���༺����,';
          if copy(FieldByName('CMun5').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '�����ȭ��,';
          if copy(FieldByName('CMun5').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '������,';
          if copy(FieldByName('CMun5').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('��������ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun6').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '�������,';
          if copy(FieldByName('CMun6').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '�˾缺���忰,';
          if copy(FieldByName('CMun6').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + 'ũ�к�,';
          if copy(FieldByName('CMun6').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + 'ġ��(ġ��,ġ��),';
          if copy(FieldByName('CMun6').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('�������׹���ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun7').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + 'B������������,';
          if copy(FieldByName('CMun7').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '����B������,';
          if copy(FieldByName('CMun7').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '����C������,';
          if copy(FieldByName('CMun7').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '���溯,';
          if copy(FieldByName('CMun7').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '��Ÿ,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('������ȯ(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          //�Ϲݹ���
          //1������ ����
          if FieldByName('Mun1_J_YN').AsString <> '' then
             Mun1_J_YN_2018.itemindex := FieldByName('Mun1_J_YN').AsInteger
          else
             Mun1_J_YN_2018.itemindex := 0;

          if FieldByName('Mun1_Jindan1').AsString = '1' then
             Ck_Mun1_J1_2018.checked := true else Ck_Mun1_J1_2018.checked := false;
          if FieldByName('Mun1_Jindan2').AsString = '1' then
             Ck_Mun1_J2_2018.checked := true else Ck_Mun1_J2_2018.checked := false;
          if FieldByName('Mun1_Jindan3').AsString = '1' then
             Ck_Mun1_J3_2018.checked := true else Ck_Mun1_J3_2018.checked := false;
          if FieldByName('Mun1_Jindan4').AsString = '1' then
             Ck_Mun1_J4_2018.checked := true else Ck_Mun1_J4_2018.checked := false;
          if FieldByName('Mun1_Jindan5').AsString = '1' then
             Ck_Mun1_J5_2018.checked := true else Ck_Mun1_J5_2018.checked := false;
          if FieldByName('Mun1_Jindan6').AsString = '1' then
             Ck_Mun1_J6_2018.checked := true else Ck_Mun1_J6_2018.checked := false;
          if FieldByName('Mun1_Jindan7').AsString = '1' then
             Ck_Mun1_J7_2018.checked := true else Ck_Mun1_J7_2018.checked := false;

          if FieldByName('Mun1_D_YN').AsString <> '' then
             Mun1_D_YN_2018.itemindex := FieldByName('Mun1_D_YN').AsInteger
          else
             Mun1_D_YN_2018.itemindex := 0;

          //1������ �๰
          if FieldByName('Mun1_Drug1').AsString = '1' then
             Ck_Mun1_D1_2018.checked := true else Ck_Mun1_D1_2018.checked := false;
          if FieldByName('Mun1_Drug2').AsString = '1' then
             Ck_Mun1_D2_2018.checked := true else Ck_Mun1_D2_2018.checked := false;
          if FieldByName('Mun1_Drug3').AsString = '1' then
          begin
             Ck_Mun1_D3_2018.checked := true;
          end
          else
          begin
             Ck_Mun1_D3_2018.checked := false;
          end;
          if FieldByName('Mun1_Drug4').AsString = '1' then
          begin
             Ck_Mun1_D4_2018.checked := true;
          end
          else
          begin
             Ck_Mun1_D4_2018.checked := false;
          end;
          if FieldByName('Mun1_Drug5').AsString = '1' then
          begin
             Ck_Mun1_D5_2018.checked := true;
          end
          else
          begin
             Ck_Mun1_D5_2018.checked := false;
          end;
          if FieldByName('Mun1_Drug6').AsString = '1' then
          begin
             Ck_Mun1_D6_2018.checked := true;
          end
          else
          begin
             Ck_Mun1_D6_2018.checked := false;
          end;
          if FieldByName('Mun1_Drug7').AsString = '1' then
             Ck_Mun1_D7_2018.checked := true else Ck_Mun1_D7_2018.checked := false;

          //2������
          if FieldByName('Mun2_YN').AsString <> '' then
             Mun2_YN_2018.itemindex := FieldByName('Mun2_YN').AsInteger
          else
             Mun2_YN_2018.itemindex := 0;

          if FieldByName('Mun2_1').AsString = '1' then
             Ck_Mun2_1_2018.checked := true else Ck_Mun2_1_2018.checked := false;
          if FieldByName('Mun2_2').AsString = '1' then
             Ck_Mun2_2_2018.checked := true else Ck_Mun2_2_2018.checked := false;
          if FieldByName('Mun2_3').AsString = '1' then
             Ck_Mun2_3_2018.checked := true else Ck_Mun2_3_2018.checked := false;
          if FieldByName('Mun2_4').AsString = '1' then
             Ck_Mun2_4_2018.checked := true else Ck_Mun2_4_2018.checked := false;
          if FieldByName('Mun2_5').AsString = '1' then
             Ck_Mun2_5_2018.checked := true else Ck_Mun2_5_2018.checked := false;

          //3������ B������
          if FieldByName('Mun3').AsString <> ''  then
             Mun3_2018.itemindex  := FieldByName('Mun3').AsInteger;

          //[2019.05.29-������]2019�⵵ ���� ����(���κ�) ..19.06.12 �߰�
          if copy(cmb_gmgn.text,1,4) = '2018' then
          begin
             //4������
             if FieldByName('Mun4').AsString        <> '' then Mun4_2018.itemindex  := FieldByName('Mun4').AsInteger;
             if FieldByName('Mun4_1_year').AsString <> '' then Mun4_1_year_2018.Text  := FieldByName('Mun4_1_year').AsString;
             if FieldByName('Mun4_1_day').AsString  <> '' then Mun4_1_day_2018.Text  := FieldByName('Mun4_1_day').AsString;
             if FieldByName('Mun4_2_year').AsString <> '' then Mun4_2_year_2018.Text  := FieldByName('Mun4_2_year').AsString;
             if FieldByName('Mun4_2_day').AsString  <> '' then Mun4_2_day_2018.Text  := FieldByName('Mun4_2_day').AsString;

             //5������
             if FieldByName('Mun5').AsString   <> '' then Mun5_2018.itemindex   := FieldByName('Mun5').AsInteger;
             if FieldByName('Mun5_1').AsString <> '' then Mun5_1_2018.itemindex := FieldByName('Mun5_1').AsInteger;
          end
          else  if copy(cmb_gmgn.text,1,4) = '2019' then
          begin
             //4������
             if FieldByName('Mun4').AsString   <> ''  then Mun4_2019.itemindex   := FieldByName('Mun4').AsInteger;
             if FieldByName('Mun4_1').AsString <> ''  then Mun4_1_2019.itemindex := FieldByName('Mun4_1').AsInteger;

             if FieldByName('Mun4_1_year').AsString  <> '' then Mun4_1_year_2019.Text  := FieldByName('Mun4_1_year').AsString;
             if FieldByName('Mun4_1_day').AsString   <> '' then Mun4_1_day_2019.Text   := FieldByName('Mun4_1_day').AsString;
             if FieldByName('Mun4_1_PYear').AsString <> '' then Mun4_1_PYear_2019.Text := FieldByName('Mun4_1_PYear').AsString;

             if FieldByName('Mun4_2').AsString   <> '' then Mun4_2_2019.itemindex   := FieldByName('Mun4_2').AsInteger;
             if FieldByName('Mun4_2_1').AsString <> '' then Mun4_2_1_2019.itemindex := FieldByName('Mun4_2_1').AsInteger;

             if FieldByName('Mun4_2_year').AsString  <> '' then Mun4_2_year_2019.Text  := FieldByName('Mun4_2_year').AsString;
             if FieldByName('Mun4_2_day').AsString   <> '' then Mun4_2_day_2019.Text   := FieldByName('Mun4_2_day').AsString;
             if FieldByName('Mun4_2_PYear').AsString <> '' then Mun4_2_PYear_2019.Text := FieldByName('Mun4_2_PYear').AsString;

             //5������
             if FieldByName('Mun5').AsString   <> '' then Mun5_2019.itemindex   := FieldByName('Mun5').AsInteger;
             if FieldByName('Mun5_1').AsString <> '' then Mun5_1_2019.itemindex := FieldByName('Mun5_1').AsInteger;
          end;

          //6������
          if FieldByName('Mun6_num').AsString <> ''  then
             Mun6_num_2018.itemindex  := FieldByName('Mun6_num').AsInteger;
          if FieldByName('Mun6_count').AsString <> ''  then
             Mun6_count_2018.Text  := FieldByName('Mun6_count').AsString;


          if FieldByName('Mun6_1_soju_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_soju_shot').AsString + '��/';
          if FieldByName('Mun6_1_soju_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_soju_btl').AsString + '��/';
          if FieldByName('Mun6_1_soju_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_soju_can').AsString + 'ĵ/';
          if FieldByName('Mun6_1_soju_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_soju_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_beer_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_beer_shot').AsString + '��/';
          if FieldByName('Mun6_1_beer_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_beer_btl').AsString + '��/';
          if FieldByName('Mun6_1_beer_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_beer_can').AsString + 'ĵ/';
          if FieldByName('Mun6_1_beer_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_beer_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_liquor_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_liquor_shot').AsString + '��/';
          if FieldByName('Mun6_1_liquor_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_liquor_btl').AsString + '��/';
          if FieldByName('Mun6_1_liquor_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_liquor_can').AsString + 'ĵ/';
          if FieldByName('Mun6_1_liquor_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_liquor_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_makgeolli_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '���ɸ�' + FieldByName('Mun6_1_makgeolli_shot').AsString + '��/';
          if FieldByName('Mun6_1_makgeolli_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '���ɸ�' + FieldByName('Mun6_1_makgeolli_btl').AsString + '��/';
          if FieldByName('Mun6_1_makgeolli_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '���ɸ�' + FieldByName('Mun6_1_makgeolli_can').AsString + 'ĵ/';
          if FieldByName('Mun6_1_makgeolli_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '���ɸ�' + FieldByName('Mun6_1_makgeolli_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_wine_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_wine_shot').AsString + '��/';
          if FieldByName('Mun6_1_wine_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_wine_btl').AsString + '��/';
          if FieldByName('Mun6_1_wine_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_wine_can').AsString + 'ĵ/';
          if FieldByName('Mun6_1_wine_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '����' + FieldByName('Mun6_1_wine_cc').AsString + 'CC/';
             

          //////////////////////////////////////
          if FieldByName('Mun6_2_soju_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_soju_shot').AsString + '��/';
          if FieldByName('Mun6_2_soju_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_soju_btl').AsString + '��/';
          if FieldByName('Mun6_2_soju_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_soju_can').AsString + 'ĵ/';
          if FieldByName('Mun6_2_soju_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_soju_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_beer_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_beer_shot').AsString + '��/';
          if FieldByName('Mun6_2_beer_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_beer_btl').AsString + '��/';
          if FieldByName('Mun6_2_beer_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_beer_can').AsString + 'ĵ/';
          if FieldByName('Mun6_2_beer_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_beer_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_liquor_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_liquor_shot').AsString + '��/';
          if FieldByName('Mun6_2_liquor_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_liquor_btl').AsString + '��/';
          if FieldByName('Mun6_2_liquor_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_liquor_can').AsString + 'ĵ/';
          if FieldByName('Mun6_2_liquor_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_liquor_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_makgeolli_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '���ɸ�' + FieldByName('Mun6_2_makgeolli_shot').AsString + '��/';
          if FieldByName('Mun6_2_makgeolli_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '���ɸ�' + FieldByName('Mun6_2_makgeolli_btl').AsString + '��/';
          if FieldByName('Mun6_2_makgeolli_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '���ɸ�' + FieldByName('Mun6_2_makgeolli_can').AsString + 'ĵ/';
          if FieldByName('Mun6_2_makgeolli_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '���ɸ�' + FieldByName('Mun6_2_makgeolli_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_wine_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_wine_shot').AsString + '��/';
          if FieldByName('Mun6_2_wine_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_wine_btl').AsString + '��/';
          if FieldByName('Mun6_2_wine_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_wine_can').AsString + 'ĵ/';
          if FieldByName('Mun6_2_wine_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '����' + FieldByName('Mun6_2_wine_cc').AsString + 'CC/';

          //7������
          if FieldByName('Mun7_1').AsString <> ''  then
             Mun7_1_2018.itemindex  := StrToInt(FieldByName('Mun7_1').AsString);
          if FieldByName('Mun7_2').AsString <> ''  then
          begin
             Mun7_2_H_2018.Text  := copy(FieldByName('Mun7_2').AsString, 1, 2);
             Mun7_2_M_2018.Text  := copy(FieldByName('Mun7_2').AsString, 3, 2);
          end;


          //8������
          if FieldByName('Mun8_1').AsString <> ''  then
             Mun8_1_2018.itemindex  := StrToInt(FieldByName('Mun8_1').AsString);
          if FieldByName('Mun8_2').AsString <> ''  then
          begin
             Mun8_2_H_2018.Text  := copy(FieldByName('Mun8_2').AsString, 1, 2);
             Mun8_2_M_2018.Text  := copy(FieldByName('Mun8_2').AsString, 3, 2);
          end;

          //9������
          if FieldByName('Mun9').AsString <> ''  then
             Mun9_2018.itemindex  := StrToInt(FieldByName('Mun9').AsString);
             
          // Ư���Ϲ���
          // 1������
          if FieldByName('CMun1').AsString <> ''  then
             CMun1_2018.itemindex  := FieldByName('CMun1').AsInteger;
          CMun1_Bung_2018.Text := FieldByName('CMun1_Bung').AsString;

          // 2������
          if FieldByName('CMun2').AsString <> ''  then
             CMun2_2018.itemindex  := FieldByName('CMun2').AsInteger;
          CMun2_kg_2018.Text := FieldByName('CMun2_kg').AsString;

          // 3������
          //����
          if FieldByName('CMun3_Can1_Yn').AsString <> ''  then
             CMun3_Can1_Yn_2018.itemindex  := FieldByName('CMun3_Can1_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 1, 1) = '2' then
             CMun3_Can1_F1_2018.checked := true else CMun3_Can1_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 2, 1) = '2' then
             CMun3_Can1_F2_2018.checked := true else CMun3_Can1_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 3, 1) = '2' then
             CMun3_Can1_F3_2018.checked := true else CMun3_Can1_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 4, 1) = '2' then
             CMun3_Can1_F4_2018.checked := true else CMun3_Can1_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 5, 1) = '2' then
             CMun3_Can1_F5_2018.checked := true else CMun3_Can1_F5_2018.checked := false;

          //�����
          if FieldByName('CMun3_Can2_Yn').AsString <> ''  then
             CMun3_Can2_Yn_2018.itemindex  := FieldByName('CMun3_Can2_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 1, 1) = '2' then
             CMun3_Can2_F1_2018.checked := true else CMun3_Can2_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 2, 1) = '2' then
             CMun3_Can2_F2_2018.checked := true else CMun3_Can2_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 3, 1) = '2' then
             CMun3_Can2_F3_2018.checked := true else CMun3_Can2_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 4, 1) = '2' then
             CMun3_Can2_F4_2018.checked := true else CMun3_Can2_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 5, 1) = '2' then
             CMun3_Can2_F5_2018.checked := true else CMun3_Can2_F5_2018.checked := false;

          //�����
          if FieldByName('CMun3_Can3_Yn').AsString <> ''  then
             CMun3_Can3_Yn_2018.itemindex  := FieldByName('CMun3_Can3_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 1, 1) = '2' then
             CMun3_Can3_F1_2018.checked := true else CMun3_Can3_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 2, 1) = '2' then
             CMun3_Can3_F2_2018.checked := true else CMun3_Can3_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 3, 1) = '2' then
             CMun3_Can3_F3_2018.checked := true else CMun3_Can3_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 4, 1) = '2' then
             CMun3_Can3_F4_2018.checked := true else CMun3_Can3_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 5, 1) = '2' then
             CMun3_Can3_F5_2018.checked := true else CMun3_Can3_F5_2018.checked := false;

          //����
          if FieldByName('CMun3_Can4_Yn').AsString <> ''  then
             CMun3_Can4_Yn_2018.itemindex  := FieldByName('CMun3_Can4_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 1, 1) = '2' then
             CMun3_Can4_F1_2018.checked := true else CMun3_Can4_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 2, 1) = '2' then
             CMun3_Can4_F2_2018.checked := true else CMun3_Can4_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 3, 1) = '2' then
             CMun3_Can4_F3_2018.checked := true else CMun3_Can4_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 4, 1) = '2' then
             CMun3_Can4_F4_2018.checked := true else CMun3_Can4_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 5, 1) = '2' then
             CMun3_Can4_F5_2018.checked := true else CMun3_Can4_F5_2018.checked := false;

          //�ڱð�ξ�
          if FieldByName('CMun3_Can5_Yn').AsString <> ''  then
             CMun3_Can5_Yn_2018.itemindex  := FieldByName('CMun3_Can5_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 1, 1) = '2' then
             CMun3_Can5_F1_2018.checked := true else CMun3_Can5_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 2, 1) = '2' then
             CMun3_Can5_F2_2018.checked := true else CMun3_Can5_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 3, 1) = '2' then
             CMun3_Can5_F3_2018.checked := true else CMun3_Can5_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 4, 1) = '2' then
             CMun3_Can5_F4_2018.checked := true else CMun3_Can5_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 5, 1) = '2' then
             CMun3_Can5_F5_2018.checked := true else CMun3_Can5_F5_2018.checked := false;

          //��Ÿ��
          CMun3_Can6_Bung_2018.Text := FieldByName('CMun3_Can6_Bung').AsString;
          if FieldByName('CMun3_Can6_Yn').AsString <> ''  then
             CMun3_Can6_Yn_2018.itemindex  := FieldByName('CMun3_Can6_Yn').AsInteger;
          //����
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 1, 1) = '2' then
             CMun3_Can6_F1_2018.checked := true else CMun3_Can6_F1_2018.checked := false;
          //�θ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 2, 1) = '2' then
             CMun3_Can6_F2_2018.checked := true else CMun3_Can6_F2_2018.checked := false;
          //����
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 3, 1) = '2' then
             CMun3_Can6_F3_2018.checked := true else CMun3_Can6_F3_2018.checked := false;
          //�ڸ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 4, 1) = '2' then
             CMun3_Can6_F4_2018.checked := true else CMun3_Can6_F4_2018.checked := false;
          //�ڳ�
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 5, 1) = '2' then
             CMun3_Can6_F5_2018.checked := true else CMun3_Can6_F5_2018.checked := false;

          // 4������
          if FieldByName('CMun4_Can1').AsString <> ''  then
             CMun4_Can1_2018.itemindex  := FieldByName('CMun4_Can1').AsInteger;
          if FieldByName('CMun4_Can2').AsString <> ''  then
             CMun4_Can2_2018.itemindex  := FieldByName('CMun4_Can2').AsInteger;
          if FieldByName('CMun4_Can3').AsString <> ''  then
             CMun4_Can3_2018.itemindex  := FieldByName('CMun4_Can3').AsInteger;
          if FieldByName('CMun4_Can4').AsString <> ''  then
             CMun4_Can4_2018.itemindex  := FieldByName('CMun4_Can4').AsInteger;
          if FieldByName('CMun4_Can5').AsString <> ''  then
             CMun4_Can5_2018.itemindex  := FieldByName('CMun4_Can5').AsInteger;
          if FieldByName('CMun4_Can6').AsString <> ''  then
             CMun4_Can6_2018.itemindex  := FieldByName('CMun4_Can6').AsInteger;
          if FieldByName('CMun4_Can7').AsString <> ''  then
             CMun4_Can7_2018.itemindex  := FieldByName('CMun4_Can7').AsInteger;
          if FieldByName('CMun4_Can8').AsString <> ''  then
             CMun4_Can8_2018.itemindex  := FieldByName('CMun4_Can8').AsInteger;

          // 5������
          if FieldByName('CMun5_YN').AsString <> ''  then
             CMun5_YN_2018.itemindex := FieldByName('CMun5_YN').AsInteger;
          if Copy(FieldByName('CMun5').AsString, 1, 1) = '2' then
             CMun5_1_2018.checked := true else CMun5_1_2018.checked := false;
          if Copy(FieldByName('CMun5').AsString, 2, 1) = '2' then
             CMun5_2_2018.checked := true else CMun5_2_2018.checked := false;
          if Copy(FieldByName('CMun5').AsString, 3, 1) = '2' then
             CMun5_3_2018.checked := true else CMun5_3_2018.checked := false;
          if Copy(FieldByName('CMun5').AsString, 4, 1) = '2' then
             CMun5_4_2018.checked := true else CMun5_4_2018.checked := false;
          if Copy(FieldByName('CMun5').AsString, 5, 1) = '2' then
             CMun5_5_2018.checked := true else CMun5_5_2018.checked := false;

          // 6������
          if FieldByName('CMun6_YN').AsString <> ''  then
             CMun6_YN_2018.itemindex := FieldByName('CMun6_YN').AsInteger;
          if Copy(FieldByName('CMun6').AsString, 1, 1) = '2' then
             CMun6_1_2018.checked := true else CMun6_1_2018.checked := false;
          if Copy(FieldByName('CMun6').AsString, 2, 1) = '2' then
             CMun6_2_2018.checked := true else CMun6_2_2018.checked := false;
          if Copy(FieldByName('CMun6').AsString, 3, 1) = '2' then
             CMun6_3_2018.checked := true else CMun6_3_2018.checked := false;
          if Copy(FieldByName('CMun6').AsString, 4, 1) = '2' then
             CMun6_4_2018.checked := true else CMun6_4_2018.checked := false;
          if Copy(FieldByName('CMun6').AsString, 5, 1) = '2' then
             CMun6_5_2018.checked := true else CMun6_5_2018.checked := false;

          // 7������
          if FieldByName('CMun7_YN').AsString <> ''  then
             CMun7_YN_2018.itemindex := FieldByName('CMun7_YN').AsInteger;
          if Copy(FieldByName('CMun7').AsString, 1, 1) = '2' then
             CMun7_1_2018.checked := true else CMun7_1_2018.checked := false;
          if Copy(FieldByName('CMun7').AsString, 2, 1) = '2' then
             CMun7_2_2018.checked := true else CMun7_2_2018.checked := false;
          if Copy(FieldByName('CMun7').AsString, 3, 1) = '2' then
             CMun7_3_2018.checked := true else CMun7_3_2018.checked := false;
          if Copy(FieldByName('CMun7').AsString, 4, 1) = '2' then
             CMun7_4_2018.checked := true else CMun7_4_2018.checked := false;
          if Copy(FieldByName('CMun7').AsString, 5, 1) = '2' then
             CMun7_5_2018.checked := true else CMun7_5_2018.checked := false;

          // ���������ڸ� �ش�

          // 8������
          if FieldByName('CMun8').AsString <> ''  then
             CMun8_2018.itemindex  := FieldByName('CMun8').AsInteger;
          CMun8_Year_2018.Text := FieldByName('CMun8_Year').AsString;

          // 9������
          if FieldByName('CMun9').AsString <> ''  then
             CMun9_2018.itemindex  := FieldByName('CMun9').AsInteger;
          CMun9_Year_2018.Text := FieldByName('CMun9_Year').AsString;

          // 10������
          if FieldByName('CMun10').AsString <> ''  then
             CMun10_2018.itemindex  := FieldByName('CMun10').AsInteger;

          // 11������
          if FieldByName('CMun11').AsString <> ''  then
             CMun11_2018.itemindex  := FieldByName('CMun11').AsInteger;

          // 12������
          if FieldByName('CMun12').AsString <> ''  then
             CMun12_2018.itemindex  := FieldByName('CMun12').AsInteger;

          // 13������
          if FieldByName('CMun13').AsString <> ''  then
             CMun13_2018.itemindex  := FieldByName('CMun13').AsInteger;

          // 14������
          if FieldByName('CMun14').AsString <> ''  then
             CMun14_2018.itemindex  := FieldByName('CMun14').AsInteger;

          //����Ȱ�� ���ǿ���
          if FieldByName('Ysno_agree').AsString <> '' then
             cmbYsno_agree_2018.itemindex := FieldByname('Ysno_agree').AsInteger;
      end;
      Active := False;
   end;
end;

procedure TfrmCK5I17.CreateTextArray_Pan;
var i, index : Integer;
    VEdt_BMI : single;
begin

     UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

     UV_HPan1_check := False; UV_HPan2_check := False; UV_HPan3_check := False; UV_HPan4_check := False; UV_HPan5_check := False;
     UV_HPan6_check := False; UV_HPan7_check := False; UV_HPan8_check := False; UV_HPan9_check := False; UV_HPan10_check := False;
     UV_HPan11_check := False; UV_HPan12_check := False;


     HPan1.Caption := '';  HPan2.Caption := '';  HPan3.Caption := '';  HPan4.Caption := ''; HPan5.Caption := '';
     HPan6.Caption := '';  HPan7.Caption := '';  HPan8.Caption := '';  HPan9.Caption := ''; HPan10.Caption := '';
     HPan11.Caption := ''; HPan12.Caption := '';

     JPan1.Caption := '';  JPan2.Caption := '';  JPan3.Caption := '';  JPan4.Caption := '';  JPan5.Caption := '';
     JPan6.Caption := '';  JPan7.Caption := '';  JPan8.Caption := '';  JPan9.Caption := '';  JPan10.Caption := '';
     JPan11.Caption := ''; JPan12.Caption := ''; JPan13.Caption := ''; JPan14.Caption := ''; JPan15.Caption := '';
     JPan16.Caption := ''; HPan13.Caption := '';

     chk_gmgn := cmb_gmgn.Text;

     with qryHsok do
     begin
        Active := False;
        ParamByName('cod_jisa').AsString   := copy(cmb_gmgn.text,12,2);
        ParamByName('num_id').AsString     := copy(cmb_gmgn.text,32,10);
        ParamByName('dat_gmgn').AsString   := copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2);
        ParamByName('dat_bunju').AsString  := copy(cmb_gmgn.text,16,8);
        ParamByName('num_bunju').AsString  := copy(cmb_gmgn.text,26,4); //�ڴ�� - 2016.03.08 - 24 -> 26���� ����
        //ParamByName('num_bunju').AsString  := copy(cmb_gmgn.text,24,4);
        Active := True;

     if RecordCount > 0 then
     begin
        while not Eof do
        begin
           //�����׸� ����
           if trim(qryHsok.FieldByName('gubn_hul').AsString) = '' then
           begin
                if (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NJ')   or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NH')   or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NHF')  or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NHH')  or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NHJ')  or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NHM')  or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'GNHG') or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'BBBB') or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NNNN') or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NH1') or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NH3')  or (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NJB') or
                    (trim(qryHsok.FieldByName('cod_sokun').AsString) = 'NJB')  then
                 begin
                    HPan1.Caption := GF_PanComparison(HPan1.Caption, 'A');
                    UV_HPan1_check := True;
                 end
                 else
                 begin
                    HPan12.Caption := GF_PanComparison(HPan12.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                    UV_HPan12_check := True;
                 end;
           end
           else
           begin
                 case qryHsok.FieldByName('gubn_hul').AsInteger of
                    1,2 : begin
                             HPan1.Caption := GF_PanComparison(HPan1.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan1_check := True;
                          end;
                      3 : begin
                             HPan2.Caption := GF_PanComparison(HPan2.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan2_check := True;
                          end;
                      4 : begin
                             HPan3.Caption := GF_PanComparison(HPan3.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan3_check := True;
                          end;
                      5 : begin
                             HPan4.Caption := GF_PanComparison(HPan4.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan4_check := True;
                          end;
                      6 : begin
                             HPan5.Caption := GF_PanComparison(HPan5.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan5_check := True;
                          end;
                      7 : begin
                             HPan6.Caption := GF_PanComparison(HPan6.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan6_check := True;
                          end;
                      8 : begin
                             HPan7.Caption := GF_PanComparison(HPan7.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan7_check := True;
                          end;
                     9 : begin
                             HPan8.Caption := GF_PanComparison(HPan8.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan8_check := True;
                          end;
                  10,11 : begin
                             HPan9.Caption := GF_PanComparison(HPan9.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan9_check := True;
                          end;
                     12 : begin
                             HPan10.Caption := GF_PanComparison(HPan10.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan10_check := True;
                          end;
                     13 : begin
                             HPan11.Caption := GF_PanComparison(HPan11.Caption, qryHsok.FieldByName('gubn_pan').AsString);
                             UV_HPan11_check := True;
                          end;
                 end;
           end;
              //========================================================
           Next;
        end;
        end;
        Active := False;
     end;

     with qryKicho2 do
     begin
        Active := False;
        Sql.Clear;

        if copy(cmb_gmgn.text,1,4) < '2018' then
        begin
           Sql.Text := ' SELECT K.*, T.Mun1_Drug3 FROM kicho K left outer join tot_munjin2010 T ' +
                       '        On K.cod_jisa = T.cod_jisa and K.num_id = T.num_id and K.dat_gmgn = T.dat_gmgn ' +
                       ' WHERE  K.num_id   = ''' + copy(cmb_gmgn.text,32,10) + ''' ' +
                       '   AND  K.dat_gmgn = ''' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + ''' ' +
                       ' order by K.dat_gmgn DESC ';
        end
        else
        begin
           Sql.Text := ' SELECT K.*, T.Mun1_Drug3 FROM kicho K left outer join tot_munjin2018 T ' +
                       '        On K.cod_jisa = T.cod_jisa and K.num_id = T.num_id and K.dat_gmgn = T.dat_gmgn ' +
                       ' WHERE  K.num_id   = ''' + copy(cmb_gmgn.text,32,10) + ''' ' +
                       '   AND  K.dat_gmgn = ''' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + ''' ' +
                       ' order by K.dat_gmgn DESC ';
        end;
        Active := True;

            //BMI ���
        if FieldByname('sinjang').Asfloat <> 0.0 then
        begin
           VEdt_BMI := StrToFloat(copy(FloatTostr(FieldByName('Chejung').AsFloat / (FieldByName('Sinjang').AsFloat / 100) / (FieldByName('Sinjang').AsFloat / 100)),1,4));
           if VEdt_BMI < 18.5 then HPan13.Caption := 'C'
           else if (VEdt_BMI>= 18.5) and (VEdt_BMI<= 24.9) then HPan13.Caption := 'A'
           else if (VEdt_BMI>= 25)   and (VEdt_BMI< 30)    then HPan13.Caption := 'C'
           else if (VEdt_BMI>= 30) then HPan13.Caption := 'E';
        end;
       
        if (FieldByname('hul_h').AsInteger > 0) AND (FieldByname('hul_l').AsInteger > 0) then
        begin
           if (FieldByname('hul_h').AsInteger >= 140 ) OR (FieldByname('hul_l').AsInteger >= 90 ) OR (FieldByName('Mun1_Drug3').AsString = '1' ) then
              JPan6.Caption := 'T'
           else if (FieldByname('hul_h').AsInteger< 90 ) OR (FieldByname('hul_l').AsInteger < 50 ) then
              JPan6.Caption := 'B'
           else if ((FieldByname('hul_h').AsInteger >= 120 ) and (FieldByname('hul_h').AsInteger < 140 )) or
                   ((FieldByname('hul_l').AsInteger >= 80 ) and (FieldByname('hul_l').AsInteger < 90 )) then
              JPan6.Caption := 'C(3)'
           else if (FieldByname('hul_h').AsInteger < 120 ) and
              (FieldByname('hul_l').AsInteger < 80 ) then
              JPan6.Caption := 'A';
        end;
       
       if (pos(FieldByname('desc_ear').AsString, '����') > 0) or
           (pos(FieldByname('desc_ear').AsString, 'NORMAL') > 0) or
           (pos(FieldByname('desc_ear').AsString, 'Normal') > 0) then JPan3.Caption := 'A'
        else if FieldByname('desc_ear').AsString = ''        then JPan3.Caption := ''
        else if (FieldByName('ear_left5').AsInteger   >= 40) or
                (FieldByName('ear_right5').AsInteger  >= 40) or
                (FieldByName('ear_left10').AsInteger  >= 40) or
                (FieldByName('ear_right10').AsInteger  >= 40) or
                (FieldByName('ear_left20').AsInteger   >= 40) or
                (FieldByName('ear_right20').AsInteger  >= 40) or
                (FieldByName('ear_left40').AsInteger   >= 40) or
                (FieldByName('ear_right40').AsInteger  >= 40) or
                (FieldByName('ear_left60').AsInteger   >= 40) or
                (FieldByName('ear_right60').AsInteger  >= 40) then
                JPan3.Caption := 'E'
        else                                      JPan3.Caption := 'C(12)';
        if copy(FieldByname('desc_ear').AsString, 1, 1) = 'D' then
           JPan3.Caption := copy(FieldByname('desc_ear').AsString, 1, 4);
       

       if (FieldByName('anap_left').AsInteger  > 0) AND (FieldByName('anap_right').AsInteger > 0) then
        begin
           if (FieldByName('anap_left').Asfloat >= 21.0) OR (FieldByName('anap_right').Asfloat >= 21.0) OR
              (FieldByName('anap_left').Asfloat- FieldByName('anap_right').Asfloat >= 5.0) OR
              (FieldByName('anap_right').Asfloat - FieldByName('anap_left').Asfloat >= 5.0) then
              JPan2.Caption := GF_PanComparison(JPan2.Caption, 'E')
           else JPan2.Caption := GF_PanComparison(JPan2.Caption, 'A');
        end;
       
       if pos(FieldByName('DESC_PEKI').AsString, '����') > 0 then JPan8.Caption := 'A'
        else if FieldByName('DESC_PEKI').AsString = ''        then JPan8.Caption := ''
        else if ((FieldByName('RSLT_PEKI1').Asfloat > 0.0) and (FieldByName('RSLT_PEKI1').Asfloat <= 79.9)) or
                ((FieldByName('RSLT_PEKI2').Asfloat > 0.0) and (FieldByName('RSLT_PEKI2').Asfloat <= 79.9)) or
                ((FieldByName('RSLT_PEKI4').Asfloat > 0.0) and (FieldByName('RSLT_PEKI4').Asfloat < 70.0)) then
                JPan8.Caption := 'E'
        else                                       JPan8.Caption := 'C(12)';
        if copy(FieldByName('DESC_PEKI').AsString, 1, 1) = 'D' then
           JPan8.Caption := copy(FieldByName('DESC_PEKI').AsString, 1, 4);

        Active := False;
     end;

    UP_ResultQuery;
end;

procedure TfrmCK5I17.UP_ResultQuery;
var sGubn, Tot_Code : String;
    iAge, Temp, temp1 : Integer;
    sResult, sTemp, cod_name: String;
    eResult : Extended;
    i, iResult_cnt, iRet, iCur, iCut, index : Integer;
    bCheck, bCheck_best,ckeck1 : Boolean;
    // �ӻ�����ġ
    eMin3, eMin2, eMin1, eLow, eHigh, eMax1, eMax2, eMax3, eIms : Extended;
    sMunja,sMan : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;
    vCod_Tkgum : Variant;
begin

         ckeck1 := True;

         with qryHul_gumgin2 do
         begin
              Active := False;
              qryHul.Filter := '';
              ParamByName('num_jumin').AsString := edtJumin.Text;
              ParamByName('dat_gmgn').AsString := copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2);
              Active := True;

              if qryHul_gumgin2.RecordCount > 0 then
              begin
                   index := 0;
                   with qryPf_hm do
                   begin
                      // 1. ���׿� ���� �˻��׸� ����
                      if qryHul_gumgin2.FieldByName('cod_hul').AsString  <> '' then
                      begin
                         Active := False;
                         ParamByName('In_Cod_Pf').AsString := qryHul_gumgin2.FieldByName('cod_hul').AsString;
                         Active := True;

                       if RecordCount > 0 then
                         begin
                            VarArrayRedim(UV_vCod_hm, RecordCount-1);
                            while not Eof do
                            begin
                               UV_vCod_hm[index] := FieldByName('COD_HM').AsString;
                               Inc(index);
                               Next;
                            end;
                         end;
                      end;

                    // 2. ���翡 ���� �˻��׸� ����
                      if qryHul_gumgin2.FieldByName('cod_cancer').AsString <> '' then
                      begin
                         Active := False;
                         ParamByName('In_Cod_Pf').AsString := qryHul_gumgin2.FieldByName('cod_cancer').AsString;
                         Active := True;

                       if RecordCount > 0 then
                         begin
                            VarArrayRedim(UV_vCod_hm, index+RecordCount-1);
                            while not Eof do
                            begin
                               UV_vCod_hm[index] := FieldByName('COD_HM').AsString;
                               Inc(index);
                               Next;
                            end;
                         end;
                      end;

                    // 3. ��� ���� �˻��׸� ����
                      if qryHul_gumgin2.FieldByName('cod_jangbi').AsString <> '' then
                      begin
                         Active := False;
                         ParamByName('In_Cod_Pf').AsString := qryHul_gumgin2.FieldByName('cod_jangbi').AsString;
                         Active := True;
               
                       if RecordCount > 0 then
                         begin
                            VarArrayRedim(UV_vCod_hm, index+RecordCount-1);
                            while not Eof do
                            begin
                               UV_vCod_hm[index] := FieldByName('COD_HM').AsString;
                               Inc(index);
                               Next;
                            end;
                         end;
                      end;


                    // 3. �߰��ڵ忡 ���� �˻��׸� ����
                      iRet := GF_MulToSingle(qryHul_gumgin2.FieldByName('cod_chuga').AsString, 4, UV_vCod_chuga2); /////////////////////��������
                      VarArrayRedim(UV_vCod_hm, index+iRet);
                      for i := 0 to iRet-1 do
                      begin
                         UV_vCod_hm[index] := UV_vCod_chuga2[i];
                         Inc(index);
                      end;
               
                     // ���, ���κ�, ������ ���� �˻��׸� ����
                      cod_name := '';
                      // �������Display
                      if qryHul_gumgin2.FieldByName('gubn_nosin').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '1', qryHul_gumgin2.FieldByName('gubn_nsyh').AsInteger);
                      //���κ�����Display
                      if qryHul_gumgin2.FieldByName('gubn_adult').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '4', qryHul_gumgin2.FieldByName('gubn_adyh').AsInteger);
                      //��������Display
                      if qryHul_gumgin2.FieldByName('gubn_agab').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '5', qryHul_gumgin2.FieldByName('gubn_agyh').AsInteger);
                      //������ȯ������Display
                      if qryHul_gumgin2.FieldByName('gubn_life').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '7', qryHul_gumgin2.FieldByName('gubn_lfyh').AsInteger);

                    //Ư������Display
                      if (qryHul_gumgin2.FieldByName('gubn_tkgm').AsString = '1') or (qryHul_gumgin2.FieldByName('gubn_tkgm').AsString = '2') then
                      begin
                         qryTkgum.Active := False;
                         qryTkgum.ParamByName('cod_jisa').AsString  := qryHul_gumgin2.FieldByName('cod_jisa').AsString;
                         qryTkgum.ParamByName('num_jumin').AsString := qryHul_gumgin2.FieldByName('num_jumin').AsString;
                         qryTkgum.ParamByName('dat_gmgn').AsString  := qryHul_gumgin2.FieldByName('dat_gmgn').AsString;
                         qryTkgum.Active := True;
                         if qryTkgum.RecordCount > 0 then
                         begin
                            iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_Tkgum);
                            for Temp := 0 to iRet - 1 do
                            begin
                              qryPf_hm.Active := False;
                              qryPf_hm.ParamByName('In_Cod_Pf').AsString := vCod_Tkgum[Temp];
                              qryPf_hm.Active := True;
                              if qryPf_hm.RecordCount > 0 then
                              begin
                                 while not qryPf_hm.Eof do
                                 begin
                                    cod_name := cod_name + qryPf_hm.FieldByName('COD_HM').AsString;
                                    qryPf_hm.next;
                                 end;   // while(qryPf_hm) end;
                              end;      //if(qryPf_hm) end;
                            end;        //for(qryTkgum) end;
                            cod_name := cod_name + qryTkgum.FieldByName('Cod_chuga').AsString;
                         end;           //if(qryTkgum) end;
                      end;

                      iRet := GF_MulToSingle(cod_name, 4, UV_vCod_totyh2);
                      for i := 0 to iRet-1 do
                      begin
                        ckeck1 := True;
                        if copy(UV_vCod_totyh2[i],1,2) <> 'JJ' then
                        begin
                          for temp1 := 0 to index do
                          begin
                            if UV_vCod_hm[temp1] = UV_vCod_totyh2[i] then ckeck1 := False;
                          end;
                          if ckeck1 = True then
                          begin
                            VarArrayRedim(UV_vCod_hm, index + 1);
                            UV_vCod_hm[index] := UV_vCod_totyh2[i];
                            Inc(index);
                          end;
                        end;
                      end;
                   end;
              end;
         end;



         // �ֹι�ȣ�� �̿��Ͽ� ������ ���̸� ����
         GP_JuminToAge(edtJumin.text,copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2), sMan, iAge);

         // ������ ���̸� �������� Column���� ����
         if iAge < 10 then sGubn := '1'
         else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
         else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
         else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
         else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
         else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
         else sGubn := '7';

         if sMan = 'F' then sGubn := sGubn + 'f';  

         Tot_Code := 'C001C002C003C004C005C006C007C009C010C011C012C013C021S007S008S016C019C025C026C027C028C022C023C024' +
                     'C032C034C083C041C042C043C039C017S004S001TT01T002T011T007T006C044T037E040' +
                     'E001E002E003E041E015C045C048C047C046C049H001H002H003H004H005H006H007H008H009H010H011H012H013H014H015H016H017H018H019H020H021H022' +
                     'S034S020U001U002U010U003U004U005U006U007U008U009U011U012U013';

         // �ʱⰪ ����
         iResult_cnt := 0;
         iCut := 0;

         with qryGlkwa2 do
         begin
              Active := False;
              ParamByName('dat_gmgn').AsString  := copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2);
              ParamByName('NUM_ID').AsString    := copy(cmb_gmgn.text,32,10);
              Active := True;
         end;

         for i := 0 to VarArrayHighBound(UV_vCod_hm, 1)-1 do
         begin
            // �ʱⰪ ����
            eMin3  := 0;
            eMin2  := 0;
            eMin1  := 0;
            eLow   := 0;
            eHigh  := 0;
            eMax1  := 0;
            eMax2  := 0;
            eMax3  := 0;
            eIms   := 0;
            sMunja := '';

            bCheck := False;
            bCheck_best := False;

            // �˻��׸� ���� �ӻ�����ġ�� ȹ��
            qryHm.Filter := '';
            qryHm.Filter := 'COD_HM = ''' + UV_vCod_hm[i] + ''' AND ' +
                            'DAT_APPLY <= ''' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + '''';

            // ���װ˻����� Check(���PF, �߰��˻� ����)
            if qryHm.FieldByName('GUBN_PART').AsString > '10' then continue;

            eMin3  := GF_StrToFloat(qryHm.FieldByName('IMS_MIN3_' + sGubn).AsString);
            eMin2  := GF_StrToFloat(qryHm.FieldByName('IMS_MIN2_' + sGubn).AsString);
            eMin1  := GF_StrToFloat(qryHm.FieldByName('IMS_MIN1_' + sGubn).AsString);
            eLow   := GF_StrToFloat(qryHm.FieldByName('IMS_LOW'   + sGubn).AsString);
            eHigh  := GF_StrToFloat(qryHm.FieldByName('IMS_HIGH'  + sGubn).AsString);
            eMax1  := GF_StrToFloat(qryHm.FieldByName('IMS_MAX1_' + sGubn).AsString);
            eMax2  := GF_StrToFloat(qryHm.FieldByName('IMS_MAX2_' + sGubn).AsString);
            eMax3  := GF_StrToFloat(qryHm.FieldByName('IMS_MAX3_' + sGubn).AsString);
            sMunja := qryHm.FIeldByName('IMS_MUNJA').AsString;

            // �ش� �˻翡 ���� ����� ����(Start Pos, End Pos�� �̿�)
            qryGlkwa2.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';
            UV_fGulkwa := '';
            UV_fGulkwa1 := qryGlkwa2.FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := qryGlkwa2.FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := qryGlkwa2.FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
            if (qryHm.FieldByName('POS_START').AsInteger <> 0) and (qryHm.FieldByName('POS_END').AsInteger <> 0) then
            begin
               sResult := Copy(UV_fGulkwa,
                            qryHm.FieldByName('POS_START').AsInteger,
                            qryHm.FieldByName('POS_END').AsInteger -
                            qryHm.FieldByName('POS_START').AsInteger + 1);
            end
            else sResult := '';

            // �˻��׸��� ����+������ ��� �˻��׸�(Ư��)���� �ӻ�����ġ ȹ��
            if qryHm.FieldByName('GUBN_HM').AsString = '3' then
            begin
               qryHm_sp.Active := False;
               qryHm_sp.ParamByName('COD_HM').AsString    := UV_vCod_hm[i];
               // �ӻ�����ġ�� �����ϴ� ������ �˻�����̹Ƿ� ��������(GV_sToday)�� ����
               qryHm_sp.ParamByName('DAT_APPLY').AsString := GV_sToday;
               qryHm_sp.Active := True;

               if qryHm_sp.RecordCount > 0 then
                  eIms := qryHm_sp.FieldByName('IMS_JUNG').AsFloat;

               // disconnect
               qryHm_sp.Active := False;
            end;

            // �ش� �˻����� �ӻ�����ġ�� ������� Check
            if qryHm.FieldByName('GUBN_HM').AsString = '1' then
            begin
               if Trim(sResult) = '' then
                  bCheck := True;

               // ������ ��� Low�� High�� ��
               eResult := GF_StrToFloat(sResult);

               if (eResult < eLow) or (eResult > eHigh) then
                  bCheck := True;

               if Trim(sResult) = '' then
                  bCheck := True;
            end
            else
            begin
               if qryHm.FieldByName('GUBN_HM').AsString = '2' then
                  sTemp := sResult
               else if qryHm.FieldByName('GUBN_HM').AsString = '3' then
                  UP_Separate(sResult, sTemp);

               // ������ ��� �����ӻ�ġ�� ��
               if GF_SpaceDel(sTemp) <> GF_SpaceDel(sMunja) then
                  bCheck := True;

               if Trim(sTemp) = '' then 
                  bCheck := True;
            end;

            

           
            //�ӻ�����ġ�� ����鼭 �Ұ��� ���� ��  ����ó��..
            if (UV_HPan1_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'C001') OR (qryHm.FieldByName('cod_hm').AsString = 'C002') OR (qryHm.FieldByName('cod_hm').AsString = 'C003') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C004') OR (qryHm.FieldByName('cod_hm').AsString = 'C005') OR (qryHm.FieldByName('cod_hm').AsString = 'C006') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C007') OR (qryHm.FieldByName('cod_hm').AsString = 'C008') OR (qryHm.FieldByName('cod_hm').AsString = 'C009') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C010') OR (qryHm.FieldByName('cod_hm').AsString = 'C011') OR (qryHm.FieldByName('cod_hm').AsString = 'C012') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C013') OR (qryHm.FieldByName('cod_hm').AsString = 'C021') OR (qryHm.FieldByName('cod_hm').AsString = 'S007') OR
               (qryHm.FieldByName('cod_hm').AsString = 'S008') OR (qryHm.FieldByName('cod_hm').AsString = 'S016')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan1.Caption := 'B';
            end;

            if (UV_HPan2_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'C019') OR (qryHm.FieldByName('cod_hm').AsString = 'C025') OR (qryHm.FieldByName('cod_hm').AsString = 'C026') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C027') OR (qryHm.FieldByName('cod_hm').AsString = 'C028') OR (qryHm.FieldByName('cod_hm').AsString = 'C022') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C023') OR (qryHm.FieldByName('cod_hm').AsString = 'C024')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan2.Caption := 'B';
            end;

            if (UV_HPan3_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'C032') OR (qryHm.FieldByName('cod_hm').AsString = 'C034') OR (qryHm.FieldByName('cod_hm').AsString = 'C083')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan3.Caption := 'B';
            end;

            if (UV_HPan4_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'C041') OR (qryHm.FieldByName('cod_hm').AsString = 'C042') OR (qryHm.FieldByName('cod_hm').AsString = 'C043')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan4.Caption := 'B';
            end;

            if (UV_HPan5_check = False) and
               ((qryHm.FieldByName('cod_hm').AsString = 'C039') AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan5.Caption := 'B';
            end;

            if (UV_HPan6_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'S003') OR (qryHm.FieldByName('cod_hm').AsString = 'S004') OR (qryHm.FieldByName('cod_hm').AsString = 'C017') OR
               (qryHm.FieldByName('cod_hm').AsString = 'S071') OR (qryHm.FieldByName('cod_hm').AsString = 'S001')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan6.Caption := 'B';
            end;

            if (UV_HPan7_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'TT01') OR (qryHm.FieldByName('cod_hm').AsString = 'TT02') OR (qryHm.FieldByName('cod_hm').AsString = 'T002') OR
               (qryHm.FieldByName('cod_hm').AsString = 'T011') OR (qryHm.FieldByName('cod_hm').AsString = 'T007') OR (qryHm.FieldByName('cod_hm').AsString = 'T006') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C044') OR (qryHm.FieldByName('cod_hm').AsString = 'T037') OR (qryHm.FieldByName('cod_hm').AsString = 'E040')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan7.Caption := 'B';
            end;

            if (UV_HPan8_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'S034') OR (qryHm.FieldByName('cod_hm').AsString = 'S020')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan8.Caption := 'B';
            end;

            if (UV_HPan9_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'C045') OR (qryHm.FieldByName('cod_hm').AsString = 'C048') OR (qryHm.FieldByName('cod_hm').AsString = 'C047') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C046') OR (qryHm.FieldByName('cod_hm').AsString = 'C049') OR (qryHm.FieldByName('cod_hm').AsString = 'H001') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H002') OR (qryHm.FieldByName('cod_hm').AsString = 'H003') OR (qryHm.FieldByName('cod_hm').AsString = 'H004') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H005') OR (qryHm.FieldByName('cod_hm').AsString = 'H006') OR (qryHm.FieldByName('cod_hm').AsString = 'H007') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H008') OR (qryHm.FieldByName('cod_hm').AsString = 'H009') OR (qryHm.FieldByName('cod_hm').AsString = 'H010') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H011') OR (qryHm.FieldByName('cod_hm').AsString = 'H012') OR (qryHm.FieldByName('cod_hm').AsString = 'H013') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H014') OR (qryHm.FieldByName('cod_hm').AsString = 'H015') OR (qryHm.FieldByName('cod_hm').AsString = 'H016') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H017') OR (qryHm.FieldByName('cod_hm').AsString = 'H018') OR (qryHm.FieldByName('cod_hm').AsString = 'H019') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H020') OR (qryHm.FieldByName('cod_hm').AsString = 'H021') OR (qryHm.FieldByName('cod_hm').AsString = 'H022')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan9.Caption := 'B';
            end;

            if (UV_HPan10_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'E001') OR (qryHm.FieldByName('cod_hm').AsString = 'E002') OR (qryHm.FieldByName('cod_hm').AsString = 'E003') OR
               (qryHm.FieldByName('cod_hm').AsString = 'E041') OR (qryHm.FieldByName('cod_hm').AsString = 'E015')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan10.Caption := 'B';
            end;

            if (UV_HPan11_check = False) and
               (((qryHm.FieldByName('cod_hm').AsString = 'U001') OR (qryHm.FieldByName('cod_hm').AsString = 'U002') OR (qryHm.FieldByName('cod_hm').AsString = 'U010') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U003') OR (qryHm.FieldByName('cod_hm').AsString = 'U004') OR (qryHm.FieldByName('cod_hm').AsString = 'U005') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U006') OR (qryHm.FieldByName('cod_hm').AsString = 'U007') OR (qryHm.FieldByName('cod_hm').AsString = 'U008') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U009') OR (qryHm.FieldByName('cod_hm').AsString = 'U011') OR (qryHm.FieldByName('cod_hm').AsString = 'U012') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U013')) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan11.Caption := 'B';
            end;

            if (UV_HPan12_check = False) and
               ((pos(qryHm.FieldByName('COD_HM').AsString, Tot_Code) <= 0) AND
               (bCheck = True) AND (trim(sResult) <> '')) then
            begin
               HPan12.Caption := 'B';
            end;
            //==========================================================
            //��Ʈ�� �Ұ��� ���� ��  ��������ó��..---------------------------------------------
            if (UV_HPan1_check = False) and (HPan1.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'C001') OR (qryHm.FieldByName('cod_hm').AsString = 'C002') OR (qryHm.FieldByName('cod_hm').AsString = 'C003') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C004') OR (qryHm.FieldByName('cod_hm').AsString = 'C005') OR (qryHm.FieldByName('cod_hm').AsString = 'C006') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C007') OR (qryHm.FieldByName('cod_hm').AsString = 'C008') OR (qryHm.FieldByName('cod_hm').AsString = 'C009') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C010') OR (qryHm.FieldByName('cod_hm').AsString = 'C011') OR (qryHm.FieldByName('cod_hm').AsString = 'C012') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C013') OR (qryHm.FieldByName('cod_hm').AsString = 'C021') OR (qryHm.FieldByName('cod_hm').AsString = 'S007') OR
               (qryHm.FieldByName('cod_hm').AsString = 'S008') OR (qryHm.FieldByName('cod_hm').AsString = 'S016')) AND
               (trim(sResult) <> '')) then
            begin
               HPan1.Caption := 'A';
            end;

            if (UV_HPan2_check = False) and (HPan2.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'C019') OR (qryHm.FieldByName('cod_hm').AsString = 'C025') OR (qryHm.FieldByName('cod_hm').AsString = 'C026') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C027') OR (qryHm.FieldByName('cod_hm').AsString = 'C028') OR (qryHm.FieldByName('cod_hm').AsString = 'C022') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C023') OR (qryHm.FieldByName('cod_hm').AsString = 'C024')) AND
               (trim(sResult) <> '')) then
            begin
               HPan2.Caption := 'A';
            end;

            if (UV_HPan3_check = False) and (HPan3.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'C032') OR (qryHm.FieldByName('cod_hm').AsString = 'C034') OR (qryHm.FieldByName('cod_hm').AsString = 'C083')) AND
               (trim(sResult) <> '')) then
            begin
               HPan3.Caption := 'A';
            end;

            if (UV_HPan4_check = False) and (HPan4.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'C041') OR (qryHm.FieldByName('cod_hm').AsString = 'C042') OR (qryHm.FieldByName('cod_hm').AsString = 'C043')) AND
               (trim(sResult) <> '')) then
            begin
               HPan4.Caption := 'A';
            end;

            if (UV_HPan5_check = False) and (HPan5.Caption = '') and
            ((qryHm.FieldByName('cod_hm').AsString = 'C039') AND
            (trim(sResult) <> '')) then
            begin
               HPan5.Caption := 'A';
            end;

            if (UV_HPan6_check = False) and (HPan6.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'S003') OR (qryHm.FieldByName('cod_hm').AsString = 'S004') OR (qryHm.FieldByName('cod_hm').AsString = 'C017') OR
               (qryHm.FieldByName('cod_hm').AsString = 'S071') OR (qryHm.FieldByName('cod_hm').AsString = 'S001')) AND
               (trim(sResult) <> '')) then
            begin
               HPan6.Caption := 'A';
            end;

            if (UV_HPan7_check = False) and (HPan7.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'TT01') OR (qryHm.FieldByName('cod_hm').AsString = 'TT02') OR (qryHm.FieldByName('cod_hm').AsString = 'T002') OR
               (qryHm.FieldByName('cod_hm').AsString = 'T011') OR (qryHm.FieldByName('cod_hm').AsString = 'T007') OR (qryHm.FieldByName('cod_hm').AsString = 'T006') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C044') OR (qryHm.FieldByName('cod_hm').AsString = 'T037') OR (qryHm.FieldByName('cod_hm').AsString = 'E040')) AND
               (trim(sResult) <> '')) then
            begin
               HPan7.Caption := 'A';
            end;

            if (UV_HPan8_check = False) and (HPan8.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'S034') OR (qryHm.FieldByName('cod_hm').AsString = 'S020')) AND
               (trim(sResult) <> '')) then
            begin
               HPan8.Caption := 'A';
            end;

            if (UV_HPan9_check = False) and (HPan9.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'C045') OR (qryHm.FieldByName('cod_hm').AsString = 'C048') OR (qryHm.FieldByName('cod_hm').AsString = 'C047') OR
               (qryHm.FieldByName('cod_hm').AsString = 'C046') OR (qryHm.FieldByName('cod_hm').AsString = 'C049') OR (qryHm.FieldByName('cod_hm').AsString = 'H001') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H002') OR (qryHm.FieldByName('cod_hm').AsString = 'H003') OR (qryHm.FieldByName('cod_hm').AsString = 'H004') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H005') OR (qryHm.FieldByName('cod_hm').AsString = 'H006') OR (qryHm.FieldByName('cod_hm').AsString = 'H007') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H008') OR (qryHm.FieldByName('cod_hm').AsString = 'H009') OR (qryHm.FieldByName('cod_hm').AsString = 'H010') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H011') OR (qryHm.FieldByName('cod_hm').AsString = 'H012') OR (qryHm.FieldByName('cod_hm').AsString = 'H013') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H014') OR (qryHm.FieldByName('cod_hm').AsString = 'H015') OR (qryHm.FieldByName('cod_hm').AsString = 'H016') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H017') OR (qryHm.FieldByName('cod_hm').AsString = 'H018') OR (qryHm.FieldByName('cod_hm').AsString = 'H019') OR
               (qryHm.FieldByName('cod_hm').AsString = 'H020') OR (qryHm.FieldByName('cod_hm').AsString = 'H021') OR (qryHm.FieldByName('cod_hm').AsString = 'H022')) AND
               (trim(sResult) <> '')) then
            begin
               HPan9.Caption := 'A';
            end;

            if (UV_HPan10_check = False) and (HPan10.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'E001') OR (qryHm.FieldByName('cod_hm').AsString = 'E002') OR (qryHm.FieldByName('cod_hm').AsString = 'E003') OR
               (qryHm.FieldByName('cod_hm').AsString = 'E041') OR (qryHm.FieldByName('cod_hm').AsString = 'E015')) AND
               (trim(sResult) <> '')) then
            begin
               HPan10.Caption := 'A';
            end;

            if (UV_HPan11_check = False) and (HPan11.Caption = '') and
               (((qryHm.FieldByName('cod_hm').AsString = 'U001') OR (qryHm.FieldByName('cod_hm').AsString = 'U002') OR (qryHm.FieldByName('cod_hm').AsString = 'U010') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U003') OR (qryHm.FieldByName('cod_hm').AsString = 'U004') OR (qryHm.FieldByName('cod_hm').AsString = 'U005') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U006') OR (qryHm.FieldByName('cod_hm').AsString = 'U007') OR (qryHm.FieldByName('cod_hm').AsString = 'U008') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U009') OR (qryHm.FieldByName('cod_hm').AsString = 'U011') OR (qryHm.FieldByName('cod_hm').AsString = 'U012') OR
               (qryHm.FieldByName('cod_hm').AsString = 'U013')) AND
               (trim(sResult) <> '')) then
            begin
               HPan11.Caption := 'A';
            end;



            if (UV_HPan12_check = False) and (HPan12.Caption = '') and
               ((pos(qryHm.FieldByName('COD_HM').AsString, Tot_Code) <= 0) AND
               (trim(sResult) <> '')) then
            begin
               HPan12.Caption := 'A';
            end;
            //===================================================================================
         end;   // End of for

         qryGlkwa2.Active := False;

         with qryJangbi2 do
         begin
              Active := False;
              ParamByName('dat_gmgn').AsString  := copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2);
              ParamByName('NUM_ID').AsString    := copy(cmb_gmgn.text,32,10);
              Active := True;

              if RecordCount > 0 then
              begin

                   for i := 0 to RecordCount - 1 do
                   begin   
                        //========================= ����׸�����
                        //����
                        if FieldByName('cod_jangbi').AsString = 'JJ03' then
                           JPan2.Caption := GF_PanComparison(JPan2.Caption, FieldByName('gubn_pan').AsString)
                        //������
                        else if FieldByName('cod_jangbi').AsString = 'JJ09' then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //�����Կ�
                        else if FieldByName('cod_jangbi').AsString = 'JJ14' then
                           JPan12.Caption := GF_PanComparison(JPan12.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ20' then
                           JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ21' then
                           JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ22' then
                           JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ23' then
                           JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ24' then
                           JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ27' then
                           JPan14.Caption := GF_PanComparison(JPan14.Caption, FieldByName('gubn_pan').AsString)
                        else if FieldByName('cod_jangbi').AsString = 'JJ2U' then
                           JPan14.Caption := GF_PanComparison(JPan14.Caption, FieldByName('gubn_pan').AsString)
                        //����������
                        else if FieldByName('cod_jangbi').AsString = 'JJ38' then
                           JPan4.Caption := GF_PanComparison(JPan4.Caption, FieldByName('gubn_pan').AsString)
                        //����������
                        else if FieldByName('cod_jangbi').AsString = 'JJ39' then
                           JPan12.Caption := GF_PanComparison(JPan12.Caption, FieldByName('gubn_pan').AsString)
                        //�ڱ�������
                        else if FieldByName('cod_jangbi').AsString = 'JJ41' then
                           JPan13.Caption := GF_PanComparison(JPan13.Caption, FieldByName('gubn_pan').AsString)
                        //����������
                        else if FieldByName('cod_jangbi').AsString = 'JJ56' then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //����CT
                        else if (FieldByName('cod_jangbi').AsString = 'JJHE') OR (FieldByName('cod_jangbi').AsString = 'JJHF') then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //���ư�ȭ����
                        else if (FieldByName('cod_jangbi').AsString = 'JJE5') or
                                (FieldByName('cod_jangbi').AsString = 'JJDM') or
                                (FieldByName('cod_jangbi').AsString = 'JJBC') then
                           JPan5.Caption := GF_PanComparison(JPan5.Caption, FieldByName('gubn_pan').AsString)
                        //�浿��������
                        else if FieldByName('cod_jangbi').AsString = 'JJB8' then
                           JPan5.Caption := GF_PanComparison(JPan5.Caption, FieldByName('gubn_pan').AsString)
                        //X-ray
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ12') or (FieldByName('cod_jangbi').AsString = 'JJB7')) then
                           JPan8.Caption := GF_PanComparison(JPan8.Caption, FieldByName('gubn_pan').AsString)
                        //[����]���ð�&�����Կ�
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ32')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ83')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ86')) and
                                 (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan10.Caption := GF_PanComparison(JPan10.Caption, FieldByName('gubn_pan').AsString)
                        //[��]���ð�&�����Կ�
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ13')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ34')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJB9')) and
                                 (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan9.Caption := GF_PanComparison(JPan9.Caption, FieldByName('gubn_pan').AsString)
                        //[��]�����˻�

                        else if ((FieldByName('cod_jangbi').AsString = 'JJ35')  or
                        (FieldByName('cod_jangbi').AsString = 'JJ78')) and  //20150120 ������ - JJ78�ڵ�� ��-�����˻翡 ��µǵ��� (����-�ֹ̼������)
                        (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan9.Caption := GF_PanComparison(JPan9.Caption, FieldByName('gubn_pan').AsString)
                        //[����]�����˻�
                        else if (FieldByName('cod_jangbi').AsString = 'JJDA') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                           JPan10.Caption := GF_PanComparison(JPan10.Caption, FieldByName('gubn_pan').AsString)
                        //��е��˻�
                        else  if ((FieldByName('cod_jangbi').AsString = 'JJ16') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH0') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH1') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH2') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH3') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH4'))and
                                 (FieldByName('cod_sokun').AsString <> '****') then
                           JPan15.Caption := GF_PanComparison(JPan15.Caption, FieldByName('gubn_pan').AsString)
                        //CT�˻�
                        else if (FieldByName('pos_glkwa').AsString = 'CT') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           //BRAIN(��) CT
                           if FieldByName('cod_jangbi').AsString = 'JJ50' then
                           begin
                              JPan1.Caption := GF_PanComparison(JPan1.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //LUNG(��) CT
                           else if FieldByName('cod_jangbi').AsString = 'JJ54' then
                           begin
                              JPan8.Caption := GF_PanComparison(JPan8.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //���� CT
                           else if FieldByName('cod_jangbi').AsString = 'JJ90' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //���κ�CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //����CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //���κ�CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //��Ÿ CT
                           else
                           begin
                              JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end
                        //MRI�˻�
                        else if ((FieldByName('pos_glkwa').AsString = 'MRI') or
                                 (FieldByName('pos_glkwa').AsString = 'MRA')) and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           if (FieldByName('cod_jangbi').AsString = 'JJC8') or (FieldByName('cod_jangbi').AsString = 'JJCN') then
                           begin
                              JPan1.Caption := GF_PanComparison(JPan1.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //��Ÿ MRI
                           else
                           begin
                              JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end
                        //PET�˻�
                        else if (FieldByName('pos_glkwa').AsString = 'PET') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                        end
                        //��Ÿ���Ұ�
                        else
                        begin
                           if (FieldByName('cod_sokun').AsString <> '****') then
                           begin
                              //10.09.04 ��ö. JJD2�� ��Ÿ������ �ݿ��� �ȵǰԲ�.ü��������(BMI)������ �־ �ȳ����°� ������ ����(������ �����)
                              if (FieldByName('cod_jangbi').AsString <> 'JJD2') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJE0') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJE7') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJEE') then
                                 JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end;
                        //=========================  �������

                        Next;
                   end;

                   // Table�� Disconnect
                   Active := False;
              end;
         end;


end;





procedure TfrmCK5I17.grd_sokunCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vDat_gmgn_sokun[DataRow-1];
       2 : Value := UV_vgubn_sokun[DataRow-1];
       3 : Value := UV_vCod_sokun[DataRow-1];
       4 : Value := UV_vDesc_sokun[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grd_sokunRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i, index, iAge : Integer;
    sMan : String;

begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grd_sokun.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grd_sokun.Rows Do
      begin
         grd_sokun.CellColor[1, i]  := clWindow;
         grd_sokun.CellColor[2, i]  := clWindow;
         grd_sokun.CellColor[3, i]  := clWindow;
         grd_sokun.CellColor[4, i]  := clWindow;

         if (i mod 2 = 0) then
         begin
            grd_sokun.CellColor[1, i]   := $00FEEEC2;
            grd_sokun.CellColor[2, i]   := $00FEEEC2;
            grd_sokun.CellColor[3, i]   := $00FEEEC2;
            grd_sokun.CellColor[4, i]   := $00FEEEC2;
         end
         else
         begin
            grd_sokun.CellColor[1, i]  := clWindow;
            grd_sokun.CellColor[2, i]  := clWindow;
            grd_sokun.CellColor[3, i]  := clWindow;
            grd_sokun.CellColor[4, i]  := clWindow;
         end;
         Next;
      end;
   end;

   if (grd_sokun.Rows > 0) and (NewRow > 0) then
   begin
      grd_sokun.CellColor[1, NewRow]  := $001AD7C9;
      grd_sokun.CellColor[2, NewRow]  := $001AD7C9;
      grd_sokun.CellColor[3, NewRow]  := $001AD7C9;
      grd_sokun.CellColor[4, NewRow]  := $001AD7C9;
   end;
end;

procedure TfrmCK5I17.SBtn_CTI_LoginClick(Sender: TObject);
var
  ctiSvrIp : String;
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  //���� ���� ��
  if      copy(GV_sUserId,1,2) = '15' then ctiSvrIp  := '125.128.11.121'  //��ȭ��
  else if copy(GV_sUserId,1,2) = '11' then ctiSvrIp  := '211.192.180.181' //����
  else if copy(GV_sUserId,1,2) = '12' then ctiSvrIp  := '121.131.48.173'  //���ǵ�
  else if copy(GV_sUserId,1,2) = '61' then ctiSvrIp  := '119.198.186.80'  //�λ�
  else if copy(GV_sUserId,1,2) = '43' then ctiSvrIp  := '59.11.26.58'     //����
  else if copy(GV_sUserId,1,2) = '51' then ctiSvrIp  := '220.93.183.118'  //����
  else
  begin
     ShowMessage('CTI�� ����� �� ���� �����Դϴ�.');
     exit;
  end;

  if Pnl_CTISvr.Caption = '����' then
  begin
     if Panel_CTI.Caption = '�α׾ƿ�' then UP_CTI_Login('IN')   //�α���...
     else                                   UP_CTI_Login('OUT'); //�α׾ƿ�...
  end
  else
  begin
     //CTI ���� ����...
     CAModule.ConnectCTIServer(ctiSvrIp);
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_AnswerClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  //��ȭ�ޱ�
  if CAModule.GetLoginState then CAModule.Answer();
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_DisconnectClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  //��ȭ����
  if CAModule.GetLoginState then CAModule.Hangup();
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_BreakClick(Sender: TObject);
var
   sTemp : string;
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  Timer_hu.Enabled := False;

  if CAModule.GetLoginState then
  begin
     sTemp := CAModule.GetTellerStatus();
     if sTemp <> 'R001' then begin CAModule.ChangeTellerStatus('R001'); end
     else                    begin CAModule.ChangeTellerStatus('0300'); end;
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_HuClick(Sender: TObject);
var
   sTemp : string;
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then
  begin
    sTemp := CAModule.GetTellerStatus();
    huTime := 0;

    if (sTemp = 'L001') or (sTemp = 'R001') then
    begin
      ShowMessage('�޽����̳� �Ļ��߿��� ��ó���� �� �� �����ϴ�.');
    end
    else if sTemp <> 'W004' then begin CAModule.ChangeTellerStatus('W004'); Timer_hu.Enabled := True;  end
    else                         begin CAModule.ChangeTellerStatus('0300'); Timer_hu.Enabled := False; end;
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_EtcClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  Timer_hu.Enabled := False;
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then
  begin
    if CAModule.GetTellerStatus <> 'W005' then CAModule.ChangeTellerStatus('W005')
    else                                       CAModule.ChangeTellerStatus('0300');
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_SettingClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then
  begin
     //���� ����� ����...
     if CAModule.GetTellerMode = '0' then CAModule.ChangeTellerMode('1')
     else                                 CAModule.ChangeTellerMode('0');
  end;
  //============================================================================
end;

procedure TfrmCK5I17.cmbRelationChange(Sender: TObject);
var sFull_FileName : string;
    AppHandle: HWND;   
begin
  inherited;
  //1 - ����÷��
  if cmbRelation.ItemIndex = 1 then
  begin
     if edtJumin.Text <> '' then
     begin
        frmCK5I171  := TFrmCK5I171.Create(self);
        frmCK5I171.ShowModal;
     end
     else
        showmessage('��ȸ�� ���������� �����ϴ�.');
  end
  //2 - ���޺��� ���
  else if cmbRelation.ItemIndex = 2 then
  begin
     frmCK5I161  := TFrmCK5I161.Create(self);
     frmCK5I161.ShowModal;
  end
  //3 - ���޺��� �Ұ��ڵ� ���
  else if cmbRelation.ItemIndex = 3 then
  begin
     frmCK5I162  := TFrmCK5I162.Create(self);
     frmCK5I162.ShowModal;
  end
  //4 - SMS���� ���ڰ���
  else if cmbRelation.ItemIndex = 4 then
  begin
     frmCK5I17B := TFrmCK5I17B.Create(self);
     frmCK5I17B.Show;
  end
  //5 - ������ ������
  else if cmbRelation.ItemIndex = 5 then
  begin
     frmCK5I163  := TFrmCK5I163.Create(self);
     frmCK5I163.ShowModal;
  end
  //7 - ���Ż�㳻�� �߰�
  else if (cmbRelation.ItemIndex = 7) and ((GV_sUserId ='151026') or (GV_sUserId='710297') or (GV_sUserId ='710298') or (GV_sUserId ='710138'))then
  begin
     frmCK5I164  := TFrmCK5I164.Create(self);
     frmCK5I164.ShowModal;
  end
  else if cmbRelation.ItemIndex = 6 then
  begin
     qrySawon.Active := False;
     qrySawon.ParamByName('cod_sawon').AsString := GV_sUserId;
     qrySawon.Active := True;

     if (qrySawon.RecordCount > 0) and (length(edtJumin.Text) = 13) and (length(copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2)) = 8) then
     begin
        if qrySawon.FieldByName('desc_licen').AsString = 'Y' then
        begin
           sFull_FileName := '\\222.222.1.6\MDMS_Chart_image\' + copy(cmb_gmgn.text,12,2) + '\' + copy(cmb_gmgn.text,1,4) + '\' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + '\' + edtJumin.Text + '.pdf';

           if not FileExists(sFull_FileName) then
           begin
              MessageDlg('PDF������ �������� �ʽ��ϴ�.. '+ #10#13 + '('+sFull_FileName+')', mtError, [mbOK], 0);
              exit;
           end
           else
           begin
              AppHandle := ShellExecute(Handle,'open',PChar(sFull_FileName),'','',SW_SHOWNORMAL);
           end;
        end
        else
        begin
           GF_MsgBox('Information', '��ȸ������ �����ϴ�..');
        end;
     end
  end
end;

procedure TfrmCK5I17.btnQueryClick(Sender: TObject);
begin
  inherited;
   edtJuminChange(edtJumin);
end;

procedure TfrmCK5I17.grd_KichoRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var I, J : Integer;
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grd_Kicho.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grd_Kicho.Rows Do
      begin
         for J := 1 to grd_Kicho.Cols Do
         begin
            grd_Kicho.CellColor[J, i] := clWindow;
         end;

         if (i mod 2 = 0) then
         begin
            for J := 1 to grd_Kicho.Cols Do
            begin
               grd_Kicho.CellColor[J, i] := $00FEEEC2;
            end;
         end
         else
         begin
            for J := 1 to grd_Kicho.Cols Do
            begin
               grd_Kicho.CellColor[J, i] := clWindow;
            end;
         end;
         Next;
      end;
   end;

   if (grd_Kicho.Rows > 0) and (NewRow > 0) then
   begin
      for J := 1 to grd_Kicho.Cols Do
      begin
         if newrow = 0 then break;
         grd_Kicho.CellColor[J, NewRow] := $001AD7C9;
         Next;
      end;
   end;
end;

procedure TfrmCK5I17.grd_KichoCellChanged(Sender: TObject; OldCol, NewCol,
  OldRow, NewRow: Integer);
  var I, J : Integer;
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewCol = 0 then exit;

   if (grd_Kicho.Cols > 1) and (NewCol > 1) then
   begin
      for i := 1 to grd_Kicho.Rows Do
      begin
         for J := 1 to grd_Kicho.Cols Do
         begin
            grd_Kicho.CellColor[J, i] := clWindow;
         end;

         if (i mod 2 = 0) then
         begin
            for J := 1 to grd_Kicho.Cols Do
            begin
               grd_Kicho.CellColor[J, i] := $00FEEEC2;
            end;
         end
         else
         begin
            for J := 1 to grd_Kicho.Cols Do
            begin
               grd_Kicho.CellColor[J, i] := clWindow;
            end;
         end;
         Next;
      end;
   end;

   if (grd_Kicho.Cols > 1) and (NewCol > 1) then
   begin
      for J := 1 to grd_Kicho.Cols Do
      begin
         if newrow = 0 then break;
         grd_Kicho.CellColor[J, NewRow] := $001AD7C9;
         Next;
      end;
   end;
end;

FUNCTION TfrmCK5I17.UF_CONNEC(sTemp : String) : String;
begin
   Result := '';
   if sTemp = '1' then Result := '��������';
   if sTemp = '2' then Result := '�λ꼾��';
   if sTemp = '3' then Result := '��������';
   if sTemp = '4' then Result := '��������';
   if sTemp = '5' then Result := '�뱸����';
   if sTemp = '6' then Result := '���ּ���';
   if sTemp = '7' then Result := '���ǵ�����';
   if sTemp = '8' then Result := '1599-7070';

   if sTemp = 'a' then Result := '��������->��ġ�ȳ�';
   if sTemp = 'b' then Result := '��������->��������';
   if sTemp = 'c' then Result := '��������->������';
   if sTemp = 'd' then Result := '��������->�����μ�';
   if sTemp = 'f' then Result := '��������->��Ÿ����';

   if sTemp = 'g' then Result := '�λ꼾��->��ġ�ȳ�';
   if sTemp = 'h' then Result := '�λ꼾��->��������';
   if sTemp = 'i' then Result := '�λ꼾��->��������';
   if sTemp = 'j' then Result := '�λ꼾��->������';
   if sTemp = 'k' then Result := '�λ꼾��->Ư������';
   if sTemp = 'l' then Result := '�λ꼾��->��Ÿ����';

   if sTemp = 'm' then Result := '��������->��ġ�ȳ�';
   if sTemp = 'n' then Result := '��������->��������';
   if sTemp = 'o' then Result := '��������->������';
   if sTemp = 'p' then Result := '��������->������';
   if sTemp = 'q' then Result := '��������->�ܷ�����';
   if sTemp = 'r' then Result := '��������->��Ÿ����';

   if sTemp = 's' then Result := '��������->��ġ�ȳ�';
   if sTemp = 't' then Result := '��������->��������';
   if sTemp = 'u' then Result := '��������->������';
   if sTemp = 'v' then Result := '��������->��Ÿ���';

   if sTemp = 'w' then Result := '�뱸����->��ġ�ȳ�';
   if sTemp = 'x' then Result := '�뱸����->��������';

   if sTemp = 'y' then Result := '���ּ���->��ġ�ȳ�';
   if sTemp = 'z' then Result := '���ּ���->��������';
   if sTemp = '!' then Result := '���ּ���->���ܰ���';
   if sTemp = '@' then Result := '���ּ���->������';
   if sTemp = '#' then Result := '���ּ���->ȯ������';
   if sTemp = '$' then Result := '���ּ���->��Ÿ����';

   if sTemp = '%' then Result := '���ǵ�����->��ġ�ȳ�';
   if sTemp = '^' then Result := '���ǵ�����->��������';
   if sTemp = '&' then Result := '���ǵ�����->������';
   if sTemp = '*' then Result := '���ǵ�����->��Ÿ����';
end;

procedure TfrmCK5I17.cmbCOD_JISAChange(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   with qryJc_sms_message do
   begin
      Active := False;
      sSelect := ' select * from Jc_sms_message ';
      if Copy(cmbCOD_JISA.Text,1,2) <> '00' then
         sWhere  := ' where cod_jisa = ''' + Copy(cmbCOD_JISA.Text,1,2) + ''' ';
      sOrder  := ' order by jisa_serial ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      if Recordcount > 0 then
      begin
         // DataSet�� ���� Variant������ �̵�
         index := 0;
         for i := 0 to RecordCount - 1 do
         begin
            UP_VarMemSet(index, 'SMS');
            UV_vSms_content[index] := FieldByName('sms_content').AsString;
            Next;
            Inc(index);
         end;
      end
      else
      begin
         index := 0;
         grdMaster.Rows := 0;
      end;
   end;

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
end;

procedure TfrmCK5I17.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := intToStr(DataRow);
       2 : Value := UV_vSms_content[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grdMasterClick(Sender: TObject);
var iPos : Integer;
begin
  inherited;
   if (Sender = grdMaster) then
   begin
      iPos := grdMaster.CurrentDataRow-1;
      meMess.Text := UV_vSms_content[iPos];
   end;
end;

procedure TfrmCK5I17.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i, J : Integer;
begin
  inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   if (grdMaster.Rows > 0) and (NewRow > 0) then
   begin
      for i := 1 to grdMaster.Rows Do
      begin
         for J := 1 to grdMaster.Cols Do
         begin
            grdMaster.CellColor[J, i] := clWindow;
         end;

         if (i mod 2 = 0) then
         begin
            for J := 1 to grdMaster.Cols Do
            begin
               grdMaster.CellColor[J, i] := $00FEEEC2;
            end;
         end
         else
         begin
            for J := 1 to grdMaster.Cols Do
            begin
               grdMaster.CellColor[J, i] := clWindow;
            end;
         end;
         Next;
      end;
   end;

   if (grdMaster.Rows > 0) and (NewRow > 0) then
   begin
      grdMaster.CellColor[1, NewRow]  := $001AD7C9;
      grdMaster.CellColor[2, NewRow]  := $001AD7C9;
   end;
end;

procedure TfrmCK5I17.Btn_SMSClick(Sender: TObject);
begin
  inherited;

  if callback_tel.Text = '' then
  begin
      GF_MsgBox('Warning', 'ȸ�Ź�ȣ�� Ȯ�����ּ���!!');
      exit;
  end;

  if Num_tel.Text = '' then
  begin
      GF_MsgBox('Warning', '���Ź�ȣ�� Ȯ�����ּ���!!');
      exit;
  end;

  if meMess.Text  <> '' then
  begin
       if (Length(meMess.Text) <= 80) then
       begin
            with qryI_SC_TRAN do
            begin
                 if Chk_yeyak.checked then
                    ParamByName('TR_SENDDATE').Asdatetime  := Strtodatetime(copy(date_yeyak.Text,1,4) + '-' +
                                                                            copy(date_yeyak.Text,5,2) + '-' +
                                                                            copy(date_yeyak.Text,7,2) + '-' +
                                                                            copy(Cmb_HH.Text,1,2) + ':' +
                                                                            copy(Cmb_MM.Text,1,2) + ':' + '00')
                 else
                    ParamByName('TR_SENDDATE').Asdatetime  := now();
                 ParamByName('TR_ID').AsString        := GV_sUserId;
                 ParamByName('TR_PHONE').AsString     := Num_tel.Text;
                 ParamByName('TR_CALLBACK').AsString  := callback_tel.Text;
                 ParamByName('TR_MSG').AsString       := meMess.Text;
                 ParamByName('TR_ETC1').AsString      := 'CK5I17';
                 ParamByName('TR_ETC2').AsString      := '';
                 ExecSql;
            end;
       end
       else if Length(meMess.Text) > 80 then
       begin
            with qryI_MMS_MSG do
            begin
                 if Chk_yeyak.checked then
                    ParamByName('REQDATE').Asdatetime  := Strtodatetime(copy(date_yeyak.Text,1,4) + '-' +
                                                                            copy(date_yeyak.Text,5,2) + '-' +
                                                                            copy(date_yeyak.Text,7,2) + '-' +
                                                                            copy(Cmb_HH.Text,1,2) + ':' +
                                                                            copy(Cmb_MM.Text,1,2) + ':' + '00')
                 else
                    ParamByName('REQDATE').Asdatetime  := now();
                 ParamByName('SUBJECT').AsString   := '�ѱ����п�����';
                 ParamByName('PHONE').AsString     := Num_tel.Text;
                 ParamByName('CALLBACK').AsString  := callback_tel.Text;
                 ParamByName('MSG').AsMemo         := meMess.Text;
                 ParamByName('ID').AsString        := GV_sUserId;
                 ParamByName('ETC1').AsString      := 'CK5I17';
                 ParamByName('ETC2').AsString      := '';

                 ExecSql;
            end;
       end;
  end
  else if meMess.Text = '' then
  begin
      GF_MsgBox('Warning', '�޼����� �Է����ּ���!!');
      exit;
  end;;

  GF_MsgBox('Information', '���ۿϷ�!');
  meMess.Text := '';
  UV_CkByte  := False;
end;

procedure TfrmCK5I17.btnYeyakClick(Sender: TObject);
begin
  inherited;
  if Sender = btnyeyak then
      GF_CalendarClick(date_yeyak);
end;

procedure TfrmCK5I17.meMessChange(Sender: TObject);
begin
  inherited;
  Label2.caption := IntTOStr(Length(meMess.Text));

  if meMess.Text = '' then
     UV_CkByte := False;

  if (Length(meMess.Text) > 80) and (UV_CkByte = False) then
  begin
     GF_MsgBox('Warning', '��Ƽ�޼����� ���۵˴ϴ�. ���ڸ޼���:13�� ��Ƽ�޼���:50��!!');
     UV_CkByte := True;
  end;
end;

procedure TfrmCK5I17.FormShow(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   with qryJc_sms_message do
   begin
      Active := False;
      sSelect := ' select * from Jc_sms_message with (Nolock) ';
      if Copy(cmbCOD_JISA.Text,1,2) <> '00' then
         sWhere  := ' where cod_jisa = ''' + Copy(cmbCOD_JISA.Text,1,2) + ''' ';
      sOrder  := ' order by jisa_serial ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      if Recordcount > 0 then
      begin
         // DataSet�� ���� Variant������ �̵�
         index := 0;
         for i := 0 to RecordCount - 1 do
         begin
            UP_VarMemSet(index, 'SMS');
            UV_vSms_content[index] := FieldByName('sms_content').AsString;
            Next;
            Inc(index);
         end;
      end
      else
      begin
         index := 0;
         grdMaster.Rows := 0;
      end;
   end;

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

end;

procedure TfrmCK5I17.SBtn_CallClick(Sender: TObject);
var
  i : Integer;
  sCall, sLost, sTel : string;
begin
  inherited;
  //[2016.08.05-������]�ű� CTI �۾�...
  //---------------------------------------------------------------------------
  sCall := ''; sLost := '';
  sTel  := Edt_Tel.Text;

  if GF_IsNumber(Edt_Tel.Text) then CAModule.MakeCall('',sTel,'')
  else
  begin
     for i := 1 to Length(sTel) do
     begin
        if (sTel[i] < '0') or (sTel[i] > '9') then sLost := sLost + sTel[i]
        else                                       sCall := sCall + sTel[i];
     end;

     CAModule.MakeCall('',sCall,'');  //�߽��ڹ�ȣ/������ȭ��ȣ/����ڵ�����
  end;
  //===========================================================================
end;

procedure TfrmCK5I17.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.',
                             '����', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster_Detail.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster_Detail.Rows + 1, 0, grdMaster_Detail.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster_Detail.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster_Detail.Cols - 1 do
               ArrV3[Row, Col] := grdMaster_Detail.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster_Detail.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster_Detail.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster_Detail.Rows + 1, grdMaster_Detail.Cols]].Value := ArrV3;


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

procedure TfrmCK5I17.Timer_huTimer(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then   
  begin
    inc(huTime);
    Pnl_CTIUser.Caption := '�����ó��(' + IntToStr(huTime) + '��)';
    if huTime >= 30 then
    begin
      //���� ���� ����...
      CAModule.ChangeTellerStatus('0300');
      Timer_hu.Enabled := False;
    end
  end;
  //============================================================================
end;

procedure TfrmCK5I17.FormDestroy(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  //���� �α׾ƿ�
  if CAModule.GetLoginState then UP_CTI_Login('OUT');
  //============================================================================
end;

procedure TfrmCK5I17.Timer1Timer(Sender: TObject);
begin
  inherited;
  //[2016.07.07-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  //CTI Server���� �������
  if CAModule.GetConnectionState then Pnl_CTISvr.Caption := '����'
  else                                Pnl_CTISvr.Caption := '����';

  //���� �α��� ����
  if CAModule.GetLoginState then
  begin
     Panel_CTI.Caption := '�α���';
     SBtn_CTI_Answer.Enabled     := True;
     SBtn_CTI_Disconnect.Enabled := True;
     SBtn_CTI_Break.Enabled      := True;
     SBtn_CTI_Hu.Enabled         := True;
     SBtn_CTI_Etc.Enabled        := True;
     SBtn_CTI_Monitor.Enabled    := True;
     SBtn_CTI_Callback.Enabled   := True;
     SBtn_CTI_Setting.Enabled    := True;

     CAModule.GetTellerStatus();
     pnl_WaitCnt.Caption := CAModule.GetWaitingCnt + '��';
  end
  else
  begin
     Panel_CTI.Caption := '�α׾ƿ�';
     SBtn_CTI_Answer.Enabled     := False;
     SBtn_CTI_Disconnect.Enabled := False;
     SBtn_CTI_Break.Enabled      := False;
     SBtn_CTI_Hu.Enabled         := False;
     SBtn_CTI_Etc.Enabled        := False;
     SBtn_CTI_Monitor.Enabled    := False;
     SBtn_CTI_Callback.Enabled   := False;
     SBtn_CTI_Setting.Enabled    := False;
  end;
  //============================================================================
end;
                                   
procedure TfrmCK5I17.UP_Change(Sender: TObject);
var sMsg : string;
begin
   inherited;
   Panel41.SetFocus;
   if Sender = cmbCell_Hangmok then UP_PAPQuery(cmbCell_Hangmok.Text);
   if Sender = cmb_gmgn then
   begin
      with qryTemp do
      begin
         Active := False;
         sMsg := '';
         Sql.Clear;
         Sql.Text := ' SELECT Num_id, dat_gmgn, cod_jangbi, cod_hul, cod_chuga FROM GUMGIN with(nolock) ' + #13 +
                     ' WHERE cod_jisa  = ''' + copy(cmb_gmgn.text,12,2) + ''' ' + #13 +
                     '   and num_jumin = ''' + edtJumin.Text + ''' ' + #13 +
                     '   and dat_gmgn  = ''' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + ''' ';

         Active := True;

         if RecordCount > 0 then
         begin
            UV_sNum_id := FieldByName('Num_id').AsString;
            if (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'SS') Or
               (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'GS') Or
               (Copy(FieldByName('cod_hul').AsString,1,2) = 'SS') Or
               (Copy(FieldByName('cod_hul').AsString,1,2) = 'GS') Or
               (POS('JJ1I',FieldByName('cod_chuga').AsString) > 0) then
            begin
               if (POS('JJ1I',FieldByName('cod_chuga').AsString) > 0) then sMsg := '[���]';

               showmessage('�ŽŰ���(������ : ' + GF_DateFormat(FieldByName('dat_gmgn').AsString) + ')' + sMsg);
            end;
         end;
      end;

      if      PageControl1.ActivePage = TabSheet7  then CreateTextArray_Cell
      else if PageControl1.ActivePage = TabSheet8  then CreateTextArrayTot_Munjin
      else if PageControl1.ActivePage = TabSheet12 then CreateTextArrayTot_Munjin2018
      else if PageControl1.ActivePage = TabSheet9  then CreateTextArray_Pan
      else if PageControl1.ActivePage = TabSheet10 then CreateTextArrayTkgum;
   end;
end;

procedure TfrmCK5I17.UP_Separate(sData : String; var sValue : String);
var i : Integer;
    sTemp : String;
begin
   // ����+���� ������ ���ڿ����� ���ڸ� �и�

   // �ڷᰡ ���� ����� ó��
   if sData = '' then
   begin
      sValue := '';
      exit;
   end;

   // ���ڿ����� Space����
   sData := GF_SpaceDel(sData);
   sValue := '';

   for i := 1 to Length(sData) do
   begin
      // �ڷᰡ �����̸� �۾��Ϸ�(�����ڷ������ ���ڷ� �ν�)
      if GF_IsNumber(sData[i]) then break;

      sValue := sValue + sData[i];
   end;
end;

//2016.02.03 - �ڴ�� - ��������(CK3I01)�� ���÷��ι�ư�� ����
procedure TfrmCK5I17.Btn_ComplainClick(Sender: TObject);
begin
   inherited;
   frmCK3I015 := TfrmCK3I015.Create(self);
   frmCK3I015.MEdt_jumin.Text    := edtJumin.Text;
   frmCK3I015.Edt_desc_name.Text := edtName.Text;
   frmCK3I015.Date.Text          := GV_sToday;
   frmCK3I015.edtCOD_DC.Text     := UV_sCod_dc;
   frmCK3I015.edtDESC_DC.Text    := Edt_Company.Text;
   frmCK3I015.ShowModal;
end;

procedure TfrmCK5I17.CAModuleCTIConnect(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  UP_CTI_Login('IN');
  //============================================================================
end;

procedure TfrmCK5I17.CAModuleCTIConnectError(Sender: TObject);
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  ShowMessage('CTI Server ���ῡ �����Ͽ����ϴ�.');
  //============================================================================
end;

procedure TfrmCK5I17.CAModuleCTIEvent(Sender: TObject;
  const strMsg: WideString);
var
   sCMD, sMsg_Cut, sMsg_Tmp, sTemp : string;
   iCnt : integer;
   aMsgList : array [0..200] of string;
begin
  inherited;
  //[2016.08.05-������]�ű�CTI�۾�...
  //----------------------------------------------------------------------------
  sMsg_Tmp := strMsg;
  for iCnt := 0 to 200 do begin aMsgList[iCnt] := ''; end;

  //ShowMessage(strMsg);

  //CMD CODE(strMsg : ���ʵ帶�� ������ ��� "^" / �ð��� ���ýð� HH:MM:SS ����)
  sCMD := copy(sMsg_Tmp,1,2);

  //�� �ʵ忡 ���� �� ���...
  iCnt := 0;
  while pos('^',sMsg_Tmp) > 0 do
  begin
     sMsg_Cut := copy(sMsg_Tmp,1,pos('^',sMsg_Tmp)-1);
     sMsg_Tmp := copy(sMsg_Tmp,pos('^',sMsg_Tmp)+1,length(sMsg_Tmp));

     aMsgList[iCnt] := sMsg_Cut;

     Inc(iCnt);
  end;
  //������ ������ �� ���...
  if sMsg_Tmp <> '' then aMsgList[iCnt] := sMsg_Tmp;


  //�α��� ����(00)
  //aMsgList(0:CMD,1: ���̵�,2:������ȣ,3:�̸�,4:�����ڵ�,5:���ʻ���,6:��׷�,7:�߱׷�,8:�ұ׷�,9:����ڼ�,10:�ð�)
  if sCMD = '00' then //�α��� ����
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        if      aMsgList[4] = '2' then ShowMessage('���̵� �������� �ʽ��ϴ�.')
        else if aMsgList[4] = '3' then ShowMessage('��й�ȣ�� ��ġ���� �ʽ��ϴ�.')
        else if aMsgList[4] = '4' then ShowMessage('�α����� �ߺ��Ǿ����ϴ�.')
        else if aMsgList[4] = '5' then ShowMessage('������ȣ �ߺ��Ǿ����ϴ�.')
        else if aMsgList[4] = '6' then ShowMessage('�̵�� �����Դϴ�.')
        else if aMsgList[4] = '7' then ShowMessage('���� �⺻�׷� �����Դϴ�.');   
     end;
  end
  else if sCMD = '02' then //��й�ȣ���� ����
  begin
     if      aMsgList[4] = '1' then ShowMessage('����')
     else if aMsgList[4] = '2' then ShowMessage('�߸��� ���̵��Դϴ�.')
     else if aMsgList[4] = '3' then ShowMessage('�߸��� ��й�ȣ�Դϴ�.');
  end
  else if sCMD = '85' then //���� ��� ����
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        if      aMsgList[2] = '0' then begin Pnl_Mod.Color := $00FDD0B5; Pnl_Mod.Caption := '�ιٿ��';   end
        else if aMsgList[2] = '1' then begin Pnl_Mod.Color := $00D6E9F3; Pnl_Mod.Caption := '�ƿ��ٿ��'; end;
     end;
  end
  else if sCMD = '86' then //��ũ�� �˾�
  begin
     if      aMsgList[2] = '1' then sTemp := 'Ring'
     else if aMsgList[2] = '2' then sTemp := 'Answer';

     if      aMsgList[1] = '01' then Pnl_PopupType.Caption := sTemp + '[ȣ�й�]'
     else if aMsgList[1] = '02' then Pnl_PopupType.Caption := sTemp + '[�ιٿ��]'
     else if aMsgList[1] = '03' then Pnl_PopupType.Caption := sTemp + '[�ιٿ��(�����ֱ�)]'
     else if aMsgList[1] = '04' then Pnl_PopupType.Caption := sTemp + '[�ƿ��ٿ��]'
     else if aMsgList[1] = '05' then Pnl_PopupType.Caption := sTemp + '[������]'
     else if aMsgList[1] = '06' then Pnl_PopupType.Caption := sTemp + '[��ܹޱ�]'
     else if aMsgList[1] = '07' then Pnl_PopupType.Caption := sTemp + '[�����ޱ�]'
     else if aMsgList[1] = '08' then Pnl_PopupType.Caption := sTemp + '[3����ȭ]';

     if      aMsgList[7] = 'S511' then Panel_CTI_Connec.Caption := '���� �� ����(����)'
     else if aMsgList[7] = '2100' then Panel_CTI_Connec.Caption := '���� �� ����(����)'
     else if aMsgList[7] = '2200' then Panel_CTI_Connec.Caption := '���� �� ����(���ǵ�)'
     else if aMsgList[7] = 'S513' then Panel_CTI_Connec.Caption := '���� ���� ���'
     else if aMsgList[7] = 'S512' then Panel_CTI_Connec.Caption := '��� ���'
     else if aMsgList[7] = 'S514' then Panel_CTI_Connec.Caption := '��Ÿ ����';


     if (sTemp = 'Ring') and (Length(Trim(UV_vCallNo)) > 4) then Pnl_tel.Text := UV_vCallNo;

     UV_sCustNo := aMsgList[5];
     UV_vCallNo := aMsgList[4];

     Panel_CTI_Custom.Text := UV_vCallNo;
     Timer_hu.Enabled         := False;

     if (Trim(UV_vCallNo) <> '') and (sTemp = 'Ring') then
     begin
        if      aMsgList[1] = '04'           then Panel_CTI_Connec.Caption := '�߽���ȭ'
        else if Length(Trim(UV_vCallNo)) < 5 then Panel_CTI_Connec.Caption := '������ȭ';
     end;

{
       ShowMessage('[��ũ���˾�����]\n�˾�Ÿ��:'             + aMsgList[1]  + '\n�˾�����:'           + aMsgList[2]  + '\nA-1 ��ǥ��ȣ:'    + aMsgList[3]  +
            	                   '\nA-2 �߽��ڹ�ȣ:'       + aMsgList[4]  + '\nA-3 IVR ����������:' + aMsgList[5]  + '\nA-4 CALL_ID:'     + aMsgList[6]  +
    			     	   '\nA-5 IVR�޴���ȣ:'      + aMsgList[7]  + '\nA-6 ����ڵ�����:'   + aMsgList[8]  + '\nA-7 �������ϸ�:'  + aMsgList[9]  +
    				   '\nA-8 ��������(������):' + aMsgList[10] + '\nB-1 ��ǥ��ȣ:'       + aMsgList[11] + '\nB-2 �߽��ڹ�ȣ:'  + aMsgList[12] +
    				   '\nB-3 IVR ����������:'   + aMsgList[13] + '\nB-4 �ݾ��̵�:'       + aMsgList[14] + '\nB-5 IVR�޴���ȣ:' + aMsgList[15] +
    				   '\nB-6 ����ڵ�����:'     + aMsgList[16]);
}
  end
  else if sCMD = '89' then //������ ������ �� �ִ� �������� �䱸�� ���� ����
  begin
    iCnt := 1;
    Cmb_CTI_status.Items.Clear;
    while iCnt < 20 do
    begin
       if length(aMsgList[iCnt]) = 4 then Cmb_CTI_status.Items.Add(aMsgList[iCnt] + '-' + aMsgList[iCnt+1]);
       iCnt := iCnt + 2;
    end;
  end
  else if sCMD = '90' then //CTI �������� ����(������ ���� ǥ�⿡ ���Ǵ� ����)
  begin
    iCnt := 1;
    Cmb_CTI_status_all.Items.Clear;
    while iCnt < 200 do
    begin
       if length(aMsgList[iCnt]) = 4 then Cmb_CTI_status_all.Items.Add(aMsgList[iCnt] + '-' + aMsgList[iCnt+1]);
       iCnt := iCnt + 2;
    end;
  end
  else if sCMD = '93' then //��� ���� ���� �䱸�� ���� ����
  begin
     //��� ���� ���� �䱸�� ���� ����
  end
  else if sCMD = '94' then //���� ���� ���� �˸�
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        sTemp := Trim(aMsgList[5]);
        if sTemp <> '' then
        begin
           //�����ó��
           if sTemp = 'W004' then begin huTime := 0; Timer_hu.Enabled := True; end;
           // if sTemp = '0200' then sTemp := '3000'; //����ڵ� �϶� ��ȭ���� ����

           GP_ComboDisplay(Cmb_CTI_status_all, sTemp, 4);

           if sTemp <> '0300' then Pnl_CTIUser.Color := $00D6E9F3
           else                    Pnl_CTIUser.Color := $00FDD0B5;

           Pnl_CTIUser.Caption := copy(Cmb_CTI_status_all.Text,6,length(Cmb_CTI_status_all.Text));
        end;

        if      aMsgList[4] = '0' then begin Pnl_Mod.Color := $00FDD0B5; Pnl_Mod.Caption := '�ιٿ��';   end
        else if aMsgList[4] = '1' then begin Pnl_Mod.Color := $00D6E9F3; Pnl_Mod.Caption := '�ƿ��ٿ��'; end;
     end;
  end
  else if sCMD = '95' then //���� ����ڼ� �˸�
  begin
     pnl_WaitCnt.Caption := aMsgList[1] + '��';
  end;

  //============================================================================
end;

procedure TfrmCK5I17.SBtn_Call2Click(Sender: TObject);
var
   i : Integer;
   sCall, sLost, sTel : string;
begin
   inherited;
   //[2016.08.05-������]�ű� CTI �۾�...
   //---------------------------------------------------------------------------
   sCall := ''; sLost := '';
   sTel  := Panel_CTI_Custom.Text;

   if GF_IsNumber(Panel_CTI_Custom.Text) then CAModule.MakeCall('',sTel,'')
   else
   begin
      for i := 1 to Length(sTel) do
      begin
         if (sTel[i] < '0') or (sTel[i] > '9') then sLost := sLost + sTel[i]
         else                                       sCall := sCall + sTel[i];
      end;

      //[2016.07.07-������]�ű� CTI �۾�...
      CAModule.MakeCall('',sCall,'');  //�߽��ڹ�ȣ/������ȭ��ȣ/����ڵ�����
   end;
   //===========================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_CallbackClick(Sender: TObject);
var
   sTemp : string;
begin
   inherited;
   sTemp := 'http://125.128.11.121:8090/callBackUnConfirmListForm.do?id=' + Panel_CTINo.Caption;
   ShellExecute(0, 'open', PChar(sTemp), '', nil, SW_SHOW);
end;

procedure TfrmCK5I17.SBtn_CTI_MonitorClick(Sender: TObject);
var
   sTemp : string;
begin
   inherited;
   sTemp:= GV_sPath + '\OnWeb.exe';

   if FileExists(sTemp) then ShellExecute(Handle, 'open', PChar(sTemp), nil, nil, SW_SHOW)
   else                      showMessage('���������� �������� �ʽ��ϴ�.');
end;

procedure TfrmCK5I17.grdTkSokunCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vTCod_matter[DataRow-1];
      2 : Value := UV_vTDesc_matter[DataRow-1];
      3 : Value := UV_vTGubn_LR[DataRow-1];
      4 : Value := UV_vTNum_seq[DataRow-1];
      5 : Value := UV_vTCod_pan[DataRow-1];
      6 : Value := UV_vTcod_sokun[DataRow-1];
      7 : Value := UV_vTcod_jochi[DataRow-1];
      8 : Value := UV_vTcod_ujs[DataRow-1];
      9 : Value := UV_vTCod_sahu1[DataRow-1];
     10 : Value := UV_vTCod_sahu2[DataRow-1];
     11 : Value := UV_vTCod_jilhwan[DataRow-1];
     12 : Value := UV_vTCod_janggi[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grdTkSokunRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var
   sName : string;
begin
   inherited;
   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   //�Ұ߼���
   EdtTNum_seq.Text     := UV_vTNum_seq[NewRow-1];

   //���ع���
   EdtTCod_HMatter.Text := UV_vTCod_matter[NewRow-1];
   EdtDesc_HMatter.Text := UV_vTDesc_matter[NewRow-1];

   //��/��
   edtTGubn_LR.Text     := UV_vTGubn_LR[NewRow-1];

   //����
   EdtTCod_pan.Text     := UV_vTCod_pan[NewRow-1];

   //�Ұ�
   EdtTcod_sokun.Text  := UV_vTcod_sokun[NewRow-1];
   EdtTDesc_sokun.Text := UF_CodeEdit(1, EdtTcod_sokun.Text);

   //��ġ
   EdtTCod_jochi.Text  := UV_vTcod_jochi[NewRow-1];
   EdtTDesc_jochi.Text := UF_CodeEdit(2, EdtTCod_jochi.Text);

   //ǥ�����
   EdtTCod_janggi.Text  := UV_vTCod_janggi[NewRow-1];
   EdtTDesc_janggi.Text := UF_CodeEdit(3, EdtTCod_janggi.Text);

   //�����ڵ�
   EdtTTot_Jilhan.Text  := UV_vTTot_jilhan[NewRow-1];

   //��ȯ
   sName := '';
   EdtTCod_jilhan.Text  := UV_vTCod_jilhwan[NewRow-1];
   if GF_SickEdit(EdtTCod_jilhan, sName) then EdtTDesc_jilhan.Text := sName;

   //�������ռ� ��
   sName := '';
   EdtTCod_ujs.Text     := UV_vTcod_ujs[NewRow-1];
   if GF_ujsEdit(EdtTCod_ujs, sName) then EdtTDesc_ujs.Text  := sName;
   if EdtTDesc_ujs.Text = '004'      then EdtTDesc_ujs.Color := clRed
   else                                   EdtTDesc_ujs.Color := clSilver;

   //���İ���(1)
   sName := '';
   EdtTCod_sahu1.Text   := UV_vTCod_sahu1[NewRow-1];
   if GF_sahuEdit(EdtTCod_sahu1, sName) then EdtTDesc_sahu1.Text  := sName;
   if EdtTDesc_sahu1.Text = '06'        then EdtTDesc_sahu1.Color := clRed
   else                                      EdtTDesc_sahu1.Color := clSilver;

   //���İ���(2)
   sName := '';
   EdtTCod_sahu2.Text   := UV_vTCod_sahu2[NewRow-1];
   if GF_sahuEdit(EdtTCod_sahu2, sName) then EdtTDesc_sahu2.Text  := sName;
   if EdtTDesc_sahu2.Text = '06'        then EdtTDesc_sahu2.Color := clRed
   else                                      EdtTDesc_sahu2.Color := clSilver;

   //���μҰ�
   MM_Tetcsokun.Text := UV_vTDesc_etcSokun[NewRow-1];
   //�����(�ۺ�)
   MM_Tetcbigo.Text  := UV_vTDesc_bigo[NewRow-1];
   //������ġ
   MM_TetcJochi.Text := UV_vTDesc_etcJochi[NewRow-1];
   //�����(KMI)
   MM_TKMI_bigo.Text := UV_vTDesc_KMI_bigo[NewRow-1];

end;

procedure TfrmCK5I17.grdResultCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := '[' + UV_vCod_Mhm[DataRow-1] + ']' + UV_vCod_Hhm[DataRow-1];
      2 : Value := UV_vTDesc_kor[DataRow-1];
      3 : Value := UV_vIms_low[DataRow-1];
      4 : Value := UV_vIms_high[DataRow-1];
      5 : Value := UV_vDesc_unit[DataRow-1];
      6 : Value := UV_vGulkwa[DataRow-1];
   end;
   if UV_vCod_Mhm[DataRow-1] <> 'JJ07' then
   begin
      if GF_IsNumber(UV_vIms_low[DataRow-1]) and GF_IsNumber(UV_vIms_high[DataRow-1]) and
         GF_IsNumber(UV_vGulkwa[DataRow-1]) then
      begin
         if ((GF_StrToFloat(UV_vGulkwa[DataRow-1]) < GF_StrToFloat(UV_vIms_low[DataRow-1]))) or
            ((GF_StrToFloat(UV_vGulkwa[DataRow-1]) > GF_StrToFloat(UV_vIms_high[DataRow-1]))) then
         begin
            // ������ Column�� ǥ��
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
         //2014.10.16 ��â��
         if GF_StrToFloat(UV_vGulkwa[DataRow-1]) = 0 Then
         begin
            // ������ Column�� ǥ��
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end
      else
      begin
         // ������ Column�� ǥ��
         grdResult.AssignCellFont(DataCol, DataRow);
         grdResult.CellFont[DataCol, DataRow].Color := clBlue;
      end;
   end
   else if UV_vCod_Mhm[DataRow-1] = 'JJ07' then
   begin
      if (UV_vCod_Hhm[DataRow-1] = '0516') or (UV_vCod_Hhm[DataRow-1] = '0518') or (UV_vCod_Hhm[DataRow-1] = '0519') then
      begin
         if (GF_StrToFloat(UV_vGulkwa[DataRow-1]) <= 70) then
         begin
            // ������ Column�� ǥ��
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end
      else if (UV_vCod_Hhm[DataRow-1] = '0515') or (UV_vCod_Hhm[DataRow-1] = '0517') then
      begin
         if (GF_StrToFloat(UV_vGulkwa[DataRow-1]) <= 3) then
         begin
            // ������ Column�� ǥ��
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end;
   end;
   if UV_vCod_Mhm[DataRow-1] = 'C027' then
   begin
      // ������ Column�� ǥ��
      grdResult.AssignCellFont(DataCol, DataRow);
      grdResult.CellFont[DataCol, DataRow].Color := clBlue;
   end;
   if (UV_vCod_Mhm[DataRow-1] = 'C083') and (EdtGubn_Cha.Text = '1') then
   begin
      // ������ Column�� ǥ��
      grdResult.AssignCellFont(DataCol, DataRow);
      grdResult.CellFont[DataCol, DataRow].Color := clBlue;
   end;
   if UV_vCod_Mhm[DataRow-1] = '----' then
   begin
      // ������ Column�� ǥ��
      grdResult.AssignCellFont(DataCol, DataRow);
      grdResult.CellFont[DataCol, DataRow].Color := clBlack;
   end;
end;

procedure TfrmCK5I17.grdResult2CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := '[' + UV_vCod_Mhm2[DataRow-1] + ']' + UV_vCod_Hhm2[DataRow-1];
      2 : Value := UV_vDesc_kor2[DataRow-1];
      3 : Begin
             Value := UV_vCKMSGulkwa2[DataRow-1];

             //������� ��������
             If UV_vCod_Mhm2[DataRow-1] = 'JJUN' Then
             Begin
                If Pos('�����豺',Value) > 0 Then
                Begin
                   grdResult2.AssignCellFont(DataCol, DataRow);
                   grdResult2.CellFont[DataCol, DataRow].Color := clBlack;
                End
                Else
                Begin
                   grdResult2.AssignCellFont(DataCol, DataRow);
                   grdResult2.CellFont[DataCol, DataRow].Color := clRed;
                End;
             End;

             //[2016.12.26-������]���к�� ��������
             If UV_vCod_Mhm2[DataRow-1] = 'JJUL' Then
             Begin
                If Pos('������',Value) > 0 Then
                Begin
                   grdResult2.AssignCellFont(DataCol, DataRow);
                   grdResult2.CellFont[DataCol, DataRow].Color := clRed;
                End
                Else
                Begin
                   grdResult2.AssignCellFont(DataCol, DataRow);
                   grdResult2.CellFont[DataCol, DataRow].Color := clBlack;
                End;
             End;
          End;
      4 : Value := UV_vHEMSGulkwa2[DataRow-1];
   end;
end;

procedure TfrmCK5I17.grdGmgnCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
     1 : Value := UV_vNum_id_G[DataRow-1];
     2 : Value := GF_DateFormat(UV_vDat_gmgn_G[DataRow-1]);
     3 : Value := UV_vCod_janbi_G[DataRow-1];
     4 : Value := UV_vCod_hul_G[DataRow-1];
     5 : Value := UV_vCod_cancer_G[DataRow-1];
     6 : Value := UV_vCod_chuga_G[DataRow-1];
     7 : Value := UV_vDesc_saler_G[DataRow-1];
     8 : Value := UV_vGubn_jinch_G[DataRow-1];
     9 : Value := UV_vNum_tel_G[DataRow-1];
    10 : Value := UV_vGubn_nosin_G[DataRow-1];
    11 : Value := UV_vGubn_nscg_G[DataRow-1];
    12 : Value := UV_vGubn_adult_G[DataRow-1];
    13 : Value := UV_vGubn_adcg_G[DataRow-1];
    14 : Value := UV_vGubn_life_G[DataRow-1];
    15 : Value := UV_vGubn_lfcg_G[DataRow-1];
    16 : Value := UV_vDesc_yhname_G[DataRow-1];
   end;
end;

procedure TfrmCK5I17.btn_CancelClick(Sender: TObject);
begin
   inherited;
   // �۾������� Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   grdMaster_DetailRowChanged(grdMaster_Detail, grdMaster_Detail.CurrentDataRow, grdMaster_Detail.CurrentDataRow);
end;

procedure TfrmCK5I17.grdMaster_DetailClickCell(Sender: TObject;
  DataColDown, DataRowDown, DataColUp, DataRowUp: Integer; DownPos,
  UpPos: TtsClickPosition);
begin
  inherited;
  if  grdMaster_Detail.CurrentDataRow = 1 then
      grdMaster_DetailRowChanged(grdMaster_Detail,0, 1);
end;

procedure TfrmCK5I17.Button1Click(Sender: TObject);
begin
  inherited;
  meMess.Font.Size := meMess.Font.Size+1 ;
end;

procedure TfrmCK5I17.Button2Click(Sender: TObject);
begin
  inherited;
  meMess.Font.Size := meMess.Font.Size -1 ;
end;

procedure TfrmCK5I17.N1Click(Sender: TObject);
Var
   j : Integer;
   stage, D_File_name : string;
begin

   if DirectoryExists('\\222.222.1.6\hospital\') then
   begin
      if edtJumin.Text <> '' then
      begin
         for j := 0 to grd_file.Rows - 1 do
         begin
            if grd_file.RowSelected[j + 1] = true then
            begin
               D_File_name := trim(grd_file.cell[1, j + 1]);

               stage := '\\222.222.1.6\hospital\' + D_File_name;

               DeleteFile(PChar(stage));

               with qryD_File_directory do
               begin
                  ParamByName('num_jumin').AsString      := edtJumin.Text;
                  ParamByName('gubn_use').AsString       := 'Hospital';
                  ParamByName('file_directory').AsString := '\\222.222.1.6\hospital\';
                  ParamByName('file_name').AsString      := D_File_name;
                  Execsql;
               end;

               showmessage(D_File_name + ' ���� ������ �Ϸ�Ǿ����ϴ�.');
               grd_file.Refresh;
               Break;
            end;
         end;
         btnQuery.Click;
      end;
   end
   else ShowMessage('\\222.222.1.6\chart\ ��Ʈ��ũ ����̺� ������ �Ǿ����� �ʽ��ϴ�. ' + #13#10 +
                    '��Ʈ��ũ ����̺� ������ ���ε� �Ͻñ� �ٶ��ϴ�.');
end;

end.