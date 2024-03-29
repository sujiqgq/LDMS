//==============================================================================
// 프로그램 설명 : 개별상담일지
// 시스템        :
// 수정일자      : 2012.12.07
// 수정자        : 김승철
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 개별상담일지
// 시스템        :
// 수정일자      : 2016.02.03
// 수정자        : 박대용
// 수정내용      : 1. 개인정보 활용동의 최근정보 적용
//                 2. 상담종류 추가 - 5.안내문자발송(DM문자발송)
//                 3. 접수창의 컴플레인 버튼 해당 페이지에서도 접근가능
//                 4. VIP검진자 확인 가능 변경
// 참고사항      : 한의 여의CRM팀 1500087 - 기안자 : 유정주
//==============================================================================
// 수정일자      : 2016.08.06
// 수정자        : 유동구
// 수정내용      : CTI신규개발(멘트찾아서복사:신규CTI작업...)
//==============================================================================
// 수정일자      : 2017.05.01
// 수정자        : 유동구
// 수정내용      : [특수]검진자료 추가...
//==============================================================================
// 수정일자      : 2017.05.11
// 수정자        : 유동구
// 수정내용      : 검진일자 표기(혈액존재->검진진행)
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
    // Grid와 연관된 Variant 변수 선언
    UV_vcod_jisa, UV_vDesc_name, UV_vNum_jumin, UV_vNum_id, UV_vDesc_dc, UV_vDesc_Dept, UV_vConsult_date, UV_vConsult_time, UV_vFile_name,
    UV_vCod_sawon, UV_vSawon_name, UV_vConsult_note, UV_vNext_gn_date, UV_nH_name, UV_nsokun_name, UV_vIdx,
    UV_vDconsult_date, UV_vDconsult_time, UV_vDSawon_name, UV_vDConsult_note, UV_vDNext_gn_date, UV_vDH_Code, UV_vDH_name, UV_vDsokun_name,
    UV_vHul_h, UV_vHul_l, UV_vBiman, UV_vSinjang, UV_vChejung, UV_vear_left5, UV_vear_left10, UV_vear_lef20, UV_vear_left30, UV_vear_left40,
    UV_vear_left60, UV_vear_right5, UV_vear_right10, UV_vear_right20, UV_vear_right30, UV_vear_right40, UV_vear_right60, UV_vdesc_ear,
    UV_veye_lman, UV_veye_rman, UV_veye_lkyo, UV_veye_rkyo, UV_vanap_left, UV_vanap_right, UV_vgubn_seksin, UV_vhyungwi, UV_vbookwi, UV_vDat_gmgn,
    UV_vInjouk_num_id, UV_vCod_hm, UV_vDCon_type  : Variant;

    UV_vDat_gmgn_G, UV_vCod_janbi_G, UV_vCod_hul_G, UV_vCod_cancer_G, UV_vCod_chuga_G, UV_vNum_tel_G, UV_vGubn_nosin_G, UV_vNum_id_G,
    UV_vGubn_nscg_G, UV_vGubn_adult_G, UV_vGubn_adcg_G, UV_vGubn_life_G, UV_vGubn_lfcg_G, UV_vDesc_yhname_G, UV_vDesc_saler_G, UV_vGubn_jinch_G : Variant;

    //혈액관련
    UV_HPan1_check, UV_HPan2_check, UV_HPan3_check, UV_HPan4_check, UV_HPan5_check, UV_HPan6_check,
    UV_HPan7_check, UV_HPan8_check, UV_HPan9_check, UV_HPan10_check, UV_HPan11_check, UV_HPan12_check : Boolean;

    //혈액검사
    UV_vCod_chuga, UV_vCod_totyh, UV_vCod_chuga2, UV_vCod_totyh2,UV_vTot_cod_hm, UV_vTot_cod_hm_name, UV_vHul_Gulkwa, UV_vHul_dat_gmgn, UV_vHul_Low_High, UV_vFont_color, UV_vHul_num_id : Variant;
    //장비검사
    UV_vDat_gmgn_Jangbi, UV_vDesc_kor, UV_vCod_sokun_Jangbi, UV_vGubn_pan, UV_vDesc_sokun_Jangbi, UV_vDesc_sbsg_Jangbi : Variant;
    //개인소견
    UV_vDat_gmgn_sokun, UV_vgubn_sokun, UV_vCod_sokun, UV_vDesc_sokun : Variant;

    //특수검진자료(물질별 소견)
    UV_vTCod_matter, UV_vTDesc_matter, UV_vTGubn_LR, UV_vTNum_seq, UV_vTCod_pan, UV_vTcod_sokun, UV_vTcod_jochi, UV_vTcod_ujs, UV_vTCod_sahu1, UV_vTCod_sahu2, UV_vTCod_jilhwan, UV_vTCod_janggi,
    UV_vTDesc_etcSokun, UV_vTDesc_etcJochi, UV_vTDesc_bigo, UV_vTDesc_KMI_bigo, UV_vTTot_jilhan : Variant;

    //특수검진결과(수치형)
    UV_vCod_Mhm, UV_vCod_Hhm, UV_vTDesc_kor, UV_vIms_low, UV_vIms_high, UV_vDesc_unit, UV_vGulkwa : Variant;

    //특수검진결과(문자형)
    UV_vCod_Mhm2, UV_vCod_Hhm2, UV_vDesc_kor2, UV_vCKMSGulkwa2, UV_vHEMSGulkwa2, UV_Gubun2, UV_vCKMSCODE : Variant;

    //인적 여러개인경우 체크
    bCK_Num_id : Boolean;
    index3     : integer;
    //SMS전송
    UV_CkByte  : Boolean;
    UV_vSms_content : Variant;

    //CTI 변수..
    UV_sCallNo : Extended;               // CID 번호
    UV_sCustNo, UV_vCallNo : String;     // 고객정보 연동

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

    //CTI연동
    huTime : Integer; //후처리 시간

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
   //sMode : [1]소견 [2]조치 [3]표적장기
   Result := ''; sSqlText := '';

   case iMode of
      1 : sSqlText := ' SELECT 소견명  Desc_Code FROM HEMS..Comm20  WHERE 소견코드 = ''' + sCode + ''' ';
      2 : sSqlText := ' SELECT 조치명  Desc_Code FROM HEMS..Comm21  WHERE 조치코드 = ''' + sCode + ''' ';
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
         // Variant변수를 사용하기 위해서 생성
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
         // 이미 생성된 변수들의 크기를 조정
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
         // Variant변수를 사용하기 위해서 생성
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
         // 이미 생성된 변수들의 크기를 조정
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
         // Variant변수를 사용하기 위해서 생성
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
         // 이미 생성된 변수들의 크기를 조정
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
   //특수검진자료 초기화...
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
         //판독의사코드
         EdtCod_Doctor.Text  := FieldByName('cod_Doctor').AsString;
         //판독의사명
         EdtDesc_Doctor.Text := FieldByName('desc_Doctor').AsString;
         //특검차수
         EdtGubn_Cha.Text    := FieldByName('gubn_cha').AsString;

         //판정완료 체크...
         if FieldByName('Ysno_Complete').AsString = 'Y' then Pnl_Complete.Caption := '판정완료'
         else                                                Pnl_Complete.Caption := '판정중(미완료)';

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

         //소견 Display...
         UP_TkSokun_Display;
         //결과 Display...
         UP_TkGulkwa_Display(qryTkgumCheck.FieldByName('cod_prf').AsString,     //[특검]프로파일
                             qryTkgumCheck.FieldByName('cod_chuga').AsString,   //[특검]검사추가
                             FieldByName('dan_sokun').AsString,                 //치과소견
                             FieldByName('Desc_EK').AsString,                   //이경소견
                             FieldByName('Desc_UBio').AsString,                 //소변세포 소견
                             FieldByName('Desc_PBio').AsString,                 //객담세포 소견
                             iSex);                                             //성별(1:남, 2:여)
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

         UV_vTCod_matter[iCnt]    := qryTk_sokun2009.FieldByName('물질코드').AsString;
         UV_vTDesc_matter[iCnt]   := qryTk_sokun2009.FieldByName('유해물질명').AsString;
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

   //검사항목 계산...
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

   //중복검사항목 제외...
   iRet := GF_MulToSingle(sHmCode, 4, vCod_Hangmok);
   sHmCodeDump := '';
   for iCountX := 0 to iRet - 1 do
   begin
      if pos(vCod_Hangmok[iCountX],sHmCodeDump) <= 0 then
         sHmCodeDump := sHmCodeDump + vCod_Hangmok[iCountX]
   end;
   sHmCode := sHmCodeDump;

   //검사항목 전역변수..
   //UV_sTkHangmokList := sHmCode;
   //UV_sTkPrfList     := UV_vCod_prf[NewRow-1];

   //검사항목 Display
   iRet := GF_MulToSingle(sHmCode, 4, vCod_Hangmok);

   //기초 내역...
   //---------------------------------------------------------------------------
   with qryTKicho do
   begin
      if RecordCount > 0 then
      begin
         //검진일자
         sDat_gmgn := FieldByName('Dat_gmgn').AsString;
         //신장
         if FieldByName('Sinjang').AsString <> '' Then eSinjang := FieldByName('Sinjang').AsFloat;

         //체중
         if FieldByName('Chejung').AsString <> '' Then eChejung := FieldByName('Chejung').AsFloat;

         //시력
         if (FieldByName('Eye_lman').AsString > '0') or (FieldByName('Eye_lkyo').AsString > '0') then
         begin
            if (FieldByName('Eye_lman').AsString = '9.9')  Or
               (FieldByName('Eye_lkyo').AsString = '9.9')  Then sEye_Left := '(좌)실명'
            else if FieldByName('Eye_lman').AsString > '0' Then sEye_Left := '(좌)' + FieldByName('Eye_lman').AsString
            else if FieldByName('Eye_lkyo').AsString > '0' Then sEye_Left := '(좌)' + FieldByName('Eye_lkyo').AsString;
         End;

         if (FieldByName('Eye_rman').AsString > '0') or
            (FieldByName('Eye_rkyo').AsString > '0') then
         begin
            if (FieldByName('Eye_rman').AsString = '9.9')  or
               (FieldByName('Eye_rkyo').AsString = '9.9')  then sEye_Right := '(우)실명'
            else if FieldByName('Eye_lman').AsString > '0' then sEye_Right := '(우)' + FieldByName('Eye_rman').AsString
            else if FieldByName('Eye_lkyo').AsString > '0' then sEye_Right := '(우)' + FieldByName('Eye_rkyo').AsString;
         end;

         //복위
         if (pos('JJ05', sHmCode) > 0) then eBookwi   := FieldByName('Bookwi').AsFloat;

         //혈압
         if (pos('JJ08', sHmCode) > 0) then
         begin
            eHul_h   := FieldByName('hul_h').AsFloat;
            eHul_l   := FieldByName('hul_l').AsFloat;
         end;

         //청력(기본)
         if pos('JJ06', sHmCode) > 0 then
         begin
            eEAR_LEFT20   := FieldByName('EAR_LEFT20').AsFloat;
            eEAR_RIGHT20  := FieldByName('EAR_RIGHT20').AsFloat;
            eEAR_LEFT30   := FieldByName('EAR_LEFT30').AsFloat;
            eEAR_RIGHT30  := FieldByName('EAR_RIGHT30').AsFloat;
            eEAR_LEFT40   := FieldByName('EAR_LEFT40').AsFloat;
            eEAR_RIGHT40  := FieldByName('EAR_RIGHT40').AsFloat;
         end;

         //청력(정밀)
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

         //페기능
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

         //요검사(PH)
         if (pos('U003', sHmCode) > 0) then eU003 := FieldByName('gubn_u003').AsFloat;

         //요검사(단백)
         if (pos('U004', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u004').AsString = '1' then sU004 := '음성'
            else if  FieldByName('gubn_u004').AsString = '2' then sU004 := '±'
            else if  FieldByName('gubn_u004').AsString = '3' then sU004 := '+1'
            else if  FieldByName('gubn_u004').AsString = '4' then sU004 := '+2'
            else if  FieldByName('gubn_u004').AsString = '5' then sU004 := '+3'
            else if  FieldByName('gubn_u004').AsString = '6' then sU004 := '+4';
         end;

         //요검사(당)
         if (pos('U005', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u005').AsString = '1' then sU005 := '음성'
            else if  FieldByName('gubn_u005').AsString = '2' then sU005 := '±'
            else if  FieldByName('gubn_u005').AsString = '3' then sU005 := '+1'
            else if  FieldByName('gubn_u005').AsString = '4' then sU005 := '+2'
            else if  FieldByName('gubn_u005').AsString = '5' then sU005 := '+3'
            else if  FieldByName('gubn_u005').AsString = '6' then sU005 := '+4';
         end;

         //요검사(적혈구)
         if (pos('U009', sHmCode) > 0) then
         begin
            if       FieldByName('gubn_u009').AsString = '1' then sU009 := '음성'
            else if  FieldByName('gubn_u009').AsString = '2' then sU009 := '±'
            else if  FieldByName('gubn_u009').AsString = '3' then sU009 := '+1'
            else if  FieldByName('gubn_u009').AsString = '4' then sU009 := '+2'
            else if  FieldByName('gubn_u009').AsString = '5' then sU009 := '+3'
            else if  FieldByName('gubn_u009').AsString = '6' then sU009 := '+4';
         end;

      end;
      Active := False;
   end;
   //===========================================================================

   ihangmok := 0;      //검사 개수...
   iHamgmokCnt := 0;   //수치검사 Cnt...
   iHamgmokCnt2 := 0;  //문자검사 Cnt...

   //[수치]신장
   UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
   UP_Grid_Clear(iHamgmokCnt,'grdResult');

   UV_vCod_Mhm[iHamgmokCnt]  := '----';
   UV_vCod_Hhm[iHamgmokCnt]  := '****';
   UV_vTDesc_kor[iHamgmokCnt] := '신장';
   UV_vGulkwa[iHamgmokCnt]   := eSinjang;
   UV_vIms_low[iHamgmokCnt]  := '';
   UV_vIms_high[iHamgmokCnt] := '';
   UV_vDesc_unit[iHamgmokCnt] := 'Cm';
   Inc(iHamgmokCnt);

   //[수치]체중
   UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
   UP_Grid_Clear(iHamgmokCnt,'grdResult');

   UV_vCod_Mhm[iHamgmokCnt]  := '----';
   UV_vCod_Hhm[iHamgmokCnt]  := '****';
   UV_vTDesc_kor[iHamgmokCnt] := '체중';
   UV_vGulkwa[iHamgmokCnt]   := eChejung;
   UV_vIms_low[iHamgmokCnt]  := '';
   UV_vIms_high[iHamgmokCnt] := '';
   UV_vDesc_unit[iHamgmokCnt] := 'Kg';
   Inc(iHamgmokCnt);

   while ihangmok <= iRet - 1 do
   begin
      if vCod_Hangmok[ihangmok] = 'JJD7' then
      begin
         //[문자]흉부방사선(Lateral_Lt) 결과...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJD7';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0433';
         UV_vDesc_kor2[iHamgmokCnt2] := '흉부방사선(Lateral_Lt)';

         qryTk_Comm12.Filter := ' 검사항목코드 = ''0433''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sJJD7;

         qryTk_gulkwa.Filter := 'cod_Hhm = ''0433''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'SE81' then
      begin
         //[문자]구강검사 결과...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'SE81';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0373';
         UV_vDesc_kor2[iHamgmokCnt2] := '구강검사(치은염,치주염등)';

         //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' 검사항목코드 = ''0373''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDan_sokun;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0373''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJF4' then
      begin
         for i := 1 to 2 do
         begin
            case i of
               1 : begin
                      //[문자]이경검사 결과...
                      UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
                      UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

                      UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJF4';
                      UV_vCod_Hhm2[iHamgmokCnt2]  := '0507';
                      UV_vDesc_kor2[iHamgmokCnt2] := '정밀 진찰(이경)-좌';

                      //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
                      //---------------------------------------------------------------------
                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0507''';
                      if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
                      //=====================================================================

                      UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_EK;
                      qryTk_gulkwa.Filter := 'cod_Hhm = ''0507''';
                      if qryTk_gulkwa.RecordCount > 0 then
                      begin
                         UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
                         UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
                      end;
                   end;
               2 : begin
                      //[문자]이경검사 결과...
                      UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
                      UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

                      UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJF4';
                      UV_vCod_Hhm2[iHamgmokCnt2]  := '0532';
                      UV_vDesc_kor2[iHamgmokCnt2] := '정밀 진찰(이경)-우';

                      //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
                      //---------------------------------------------------------------------
                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0532''';
                      if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
                      //=====================================================================

                      UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_EK;
                      qryTk_gulkwa.Filter := 'cod_Hhm = ''0532''';
                      if qryTk_gulkwa.RecordCount > 0 then
                      begin
                         UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
                         UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
                      end;
                   end;
            end;
            Inc(iHamgmokCnt2);
         end;
      end
      else if vCod_Hangmok[ihangmok] = 'JJFG' then
      begin
         //[문자]소변세포병리검사 결과...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJFG';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0399';
         UV_vDesc_kor2[iHamgmokCnt2] := '파파니콜라우검사(요세포검사)';

         //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' 검사항목코드 = ''0399''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_UBio;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0399''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJFH' then
      begin
         //[문자]객담세포검사 결과...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJFH';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0270';
         UV_vDesc_kor2[iHamgmokCnt2] := '객담세포학적검사';

         //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' 검사항목코드 = ''0270''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sDesc_PBio;

         qryTk_gulkwa.Filter := 'cod_Hhm = ''0270''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
         end;

         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJ05' then
      begin
         //[수치]복부둘레...
         UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
         UP_Grid_Clear(iHamgmokCnt,'grdResult');

         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ05';
         UV_vCod_Hhm[iHamgmokCnt]  := '0522';
         UV_vTDesc_kor[iHamgmokCnt] := '복부둘레';
         UV_vGulkwa[iHamgmokCnt]   := eBookwi;

         qryTk_Comm12.Filter := ' 검사항목코드 = ''0522''';
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
         //혈압 결과 LSIT...
         for i := 1 to 2 do
         begin
            UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
            UP_Grid_Clear(iHamgmokCnt,'grdResult');

            case i of
               1 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ08';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0181';
                      UV_vTDesc_kor[iHamgmokCnt] := '혈압(최고)';
                      UV_vGulkwa[iHamgmokCnt]   := eHul_h;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0181''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '혈압(최저)';
                      UV_vGulkwa[iHamgmokCnt]   := eHul_l;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0182''';
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
         //폐기능 결과 LSIT...
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0515''';
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0516''';
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0517''';
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0518''';
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0519''';
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

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0521''';
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

         //폐기능 소견...
         UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
         UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

         UV_vCod_Mhm2[iHamgmokCnt2]  := 'JJ07';
         UV_vCod_Hhm2[iHamgmokCnt2]  := '0437';
         UV_vDesc_kor2[iHamgmokCnt2] := '폐활량검사';

         //[2017.02.06-유동구]결과값유형 항목에서 가져오기...
         //---------------------------------------------------------------------
         qryTk_Comm12.Filter := ' 검사항목코드 = ''0437''';
         if qryTk_Comm12.RecordCount > 0 then UV_Gubun2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;
         //UV_Gubun2[iHamgmokCnt2]     := '018';
         //=====================================================================

         UV_vCKMSGulkwa2[iHamgmokCnt2] := sPekiSokun;
         qryTk_gulkwa.Filter := 'cod_Hhm = ''0437''';
         if qryTk_gulkwa.RecordCount > 0 then
         begin
            UV_vCKMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('desc_etc').AsString;
            UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
         end;
         
         Inc(iHamgmokCnt2);
      end
      else if vCod_Hangmok[ihangmok] = 'JJ06' then
      begin
         if pos('JJ76',sHmCode) <= 0 then
         begin
            //1차 청력 항목 LSIT...
            for i := 1 to 6 do
            begin
               UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
               UP_Grid_Clear(iHamgmokCnt,'grdResult');

               case i of
                  1 : begin
                         UV_vCod_Mhm[iHamgmokCnt]  := 'JJ06';
                         UV_vCod_Hhm[iHamgmokCnt]  := '0160';
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 2000Hz(좌)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0160''';
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
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 2000Hz(우)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0162''';
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
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 3000Hz(좌)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0363''';
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
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 3000Hz(우)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0364''';
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
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 4000Hz(좌)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0164''';
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
                         UV_vTDesc_kor[iHamgmokCnt] := '기도청력 4000Hz(우)';
                         UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40;

                         qryTk_Comm12.Filter := ' 검사항목코드 = ''0166''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '악력검사(좌)';

                      //장비결과..
                      qryJangbi.Filter := 'cod_jangbi = ''' + vCod_Hangmok[ihangmok] + '''';
                      if qryJangbi.RecordCount > 0 then
                         UV_vGulkwa[iHamgmokCnt] := copy(Trim(qryJangbi.FieldByName('desc_sokun').AsString),
                                                    1,
                                                    pos('/',Trim(qryJangbi.FieldByName('desc_sokun').AsString))-1);

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0106''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '악력검사(우)';

                      //장비결과..
                      qryJangbi.Filter := 'cod_jangbi = ''' + vCod_Hangmok[ihangmok] + '''';
                      if qryJangbi.RecordCount > 0 then
                         UV_vGulkwa[iHamgmokCnt] := copy(Trim(qryJangbi.FieldByName('desc_sokun').AsString),
                                                    pos('/',Trim(qryJangbi.FieldByName('desc_sokun').AsString))+1,
                                                    length(Trim(qryJangbi.FieldByName('desc_sokun').AsString)));

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0107''';
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
         //2차 정밀청력 항목 LSIT...
         for i := 1 to 22 do
         begin
            UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
            UP_Grid_Clear(iHamgmokCnt,'grdResult');

            case i of
               1 : begin
                      UV_vCod_Mhm[iHamgmokCnt]  := 'JJ76';
                      UV_vCod_Hhm[iHamgmokCnt]  := '0152';
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 500Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT5;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0152''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 500Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT5;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0154''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 1000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT10;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0156''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 1000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT10;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0158''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 2000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0160''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 2000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0162''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 3000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0363''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 3000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0364''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 4000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0164''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 4000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0166''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 6000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT60;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0367''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '기도청력 6000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT60;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0368''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 500Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT5k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0153''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 500Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT5k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0155''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 1000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT10k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0157''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 1000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT10k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0159''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 2000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT20k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0161''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 2000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT20k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0163''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 3000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT30k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0365''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 3000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT30k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0366''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 4000Hz(좌)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_LEFT40k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0165''';
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
                      UV_vTDesc_kor[iHamgmokCnt] := '골전도청력 4000Hz(우)';
                      UV_vGulkwa[iHamgmokCnt]   := eEAR_RIGHT40k;

                      qryTk_Comm12.Filter := ' 검사항목코드 = ''0167''';
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
         //진찰코드로 화면에 Display 제외...
      end
      else
      begin
         if vCod_Hangmok[ihangmok] = 'JJB9' then vCod_Hangmok[ihangmok] := 'JJ34';

         qryTk_Comm12.Filter := ' MDMS항목코드 = ''' + vCod_Hangmok[ihangmok] + '''';
         if qryTk_Comm12.RecordCount > 0 then
         begin
            if qryTk_Comm12.FieldByName('HM_GUM_INFO').AsString = '1' then
            begin
               UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');
               UP_Grid_Clear(iHamgmokCnt,'grdResult');

               UV_vCod_Mhm[iHamgmokCnt]  := vCod_Hangmok[ihangmok];
               UV_vCod_Hhm[iHamgmokCnt]  := qryTk_Comm12.FieldByName('검사항목코드').AsString;
               UV_vTDesc_kor[iHamgmokCnt] := qryTk_Comm12.FieldByName('검사항목명_한').AsString;
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
               // 결과값 입력...
               dmComm.qryHangmok.Filter := 'COD_HM = ''' + vCod_Hangmok[ihangmok] + ''' AND ' + 'DAT_APPLY <= ''' + sDat_gmgn + '''';
               if dmComm.qryHangmok.RecordCount > 0 then
               begin
                  if dmComm.qryHangmok.FieldByName('gubn_part').AsString < '10' then
                  begin
                     //혈액결과..
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

                        //대사증후군 Check - C026,C028, C032 검사자료 추출;
                        If (Trim(UV_vGulkwa[iHamgmokCnt]) <> '') And
                           ((vCod_Hangmok[ihangmok] = 'C026') Or
                            (vCod_Hangmok[ihangmok] = 'C028') Or
                            (vCod_Hangmok[ihangmok] = 'C032')) Then
                        Begin
                           If vCod_Hangmok[ihangmok] = 'C026' Then eC026 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C028' Then eC028 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                           If vCod_Hangmok[ihangmok] = 'C032' Then eC032 := StrToFloat(UV_vGulkwa[iHamgmokCnt]);
                        End;

                        //C025,C026, C028 검사자료 추출;
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
                              //변수 할당
                              Inc(iHamgmokCnt);
                              UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');

                              UV_vCod_Mhm[iHamgmokCnt]   := 'C027';
                              UV_vCod_Hhm[iHamgmokCnt]   := '****';
                              UV_vTDesc_kor[iHamgmokCnt]  := '**LDL콜레스테롤';
                              UV_vIms_low[iHamgmokCnt]   := '50.0';
                              UV_vIms_high[iHamgmokCnt]  := '129.9';
                              UV_vDesc_unit[iHamgmokCnt] := 'mg/dL';
                              UV_vGulkwa[iHamgmokCnt]    := '---';
                              If (eC025 >0) And (eC026 > 0) And (eC028 > 0) Then
                              begin
                                  If eC028 >= 400 Then
                                  begin
                                     UV_vGulkwa[iHamgmokCnt]  := '???';

                                     //[2017.02.09-유동구]한의 재단특수행정파트1700412
                                     //TG(C028) 값이 400 이 넘으면 LDL(CO27)에 직접검사한 수치 넣음...
                                     //전국센터 통일되어 있음..
                                     if Trim(Copy(UV_fGulkwa, 133, 6)) <> '' then
                                     begin
                                        Inc(iHamgmokCnt);
                                        UP_VarMemSet_Tkgum(iHamgmokCnt,'grdResult');

                                        UV_vCod_Mhm[iHamgmokCnt]   := 'C027';
                                        UV_vCod_Hhm[iHamgmokCnt]   := '****';
                                        UV_vTDesc_kor[iHamgmokCnt]  := '**LDL-콜레스테롤(직접)';
                                        UV_vIms_low[iHamgmokCnt]   := '50.0';
                                        UV_vIms_high[iHamgmokCnt]  := '129.9';
                                        UV_vDesc_unit[iHamgmokCnt] := 'mg/dL';
                                        UV_vGulkwa[iHamgmokCnt]    := Trim(Copy(UV_fGulkwa, 133, 6));
                                     end;
                                  end
                                  Else UV_vGulkwa[iHamgmokCnt] := eC025 - eC026 - (eC028 /5);
                              end;

                              // Hb A1C 결과 확인 건 2015.04.23 한의 재단보건관리대행파트1500252
                              sC083 := Trim(Copy(UV_fGulkwa, 385, 6));
                              If Trim(sC083) <> '' Then
                              Begin
                                 //변수 할당
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
                     //장비결과..
                     qryJangbi.Filter := 'cod_jangbi = ''' + dmComm.qryHangmok.FieldByName('cod_hm').AsString + '''';
                     if qryJangbi.RecordCount > 0 then UV_vGulkwa[iHamgmokCnt] := Trim(qryJangbi.FieldByName('desc_sokun').AsString);
                  end;
               end;

               Inc(iHamgmokCnt);
            end
            else
            begin
               // 그외 검사....
               UP_VarMemSet_Tkgum(iHamgmokCnt2,'grdResult2');
               UP_Grid_Clear(iHamgmokCnt2,'grdResult2');

               UV_vCod_Mhm2[iHamgmokCnt2]  := vCod_Hangmok[ihangmok];
               UV_vCod_Hhm2[iHamgmokCnt2]  := qryTk_Comm12.FieldByName('검사항목코드').AsString;
               UV_vDesc_kor2[iHamgmokCnt2] := qryTk_Comm12.FieldByName('검사항목명_한').AsString;
               UV_Gubun2[iHamgmokCnt2]     := qryTk_Comm12.FieldByName('HM_RFM_CODE').AsString;

               //--------------------------------------------------------------------------
               // 결과값 입력...
               dmComm.qryHangmok.Filter := 'COD_HM = ''' + vCod_Hangmok[ihangmok] + ''' AND ' + 'DAT_APPLY <= ''' + sDat_gmgn + '''';
               if dmComm.qryHangmok.RecordCount > 0 then
               begin
                  if dmComm.qryHangmok.FieldByName('gubn_part').AsString < '10' then
                  begin
                     //혈액결과..
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
                     //시력결과..
                     if dmComm.qryHangmok.FieldByName('cod_hm').AsString = 'JJ02' then
                     begin
                        UV_vCKMSGulkwa2[iHamgmokCnt2] := sEye_Left +'/' + sEye_Right;
                     end
                     //장비결과..
                     Else if dmComm.qryHangmok.FieldByName('cod_hm').AsString = 'JJ34' then
                     begin
                        // 위 내시경이 아니라 수면 내시경일 경우
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
                              UV_vCKMSGulkwa2[iHamgmokCnt2] := '미촬영';   //2014.05.08 한의 본원의료지원본부장1400018
                        end;
                     end;
                  end;
               end;

               qryTk_gulkwa.Filter := 'cod_Hhm = ''' + qryTk_Comm12.FieldByName('검사항목코드').AsString + '''';
               if qryTk_gulkwa.RecordCount > 0 then
               begin
                  UV_vHEMSGulkwa2[iHamgmokCnt2] := qryTk_gulkwa.FieldByName('cod_gulkwa').AsString + '-' + qryTk_gulkwa.FieldByName('검사결과명').AsString;
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
  sStrid  := '';       //상담원 ID
  sStrPwd := '';       //상담원 비밀번호
  sStrExtension := ''; //상담원 내선번호

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
        showmessage('CTI 사용자가 아닙니다.' + #13#10 + '접속 ID를 확인 바랍니다.');
        exit;
     end;
  end
  else if sLogin = 'OUT' then
  begin
     CAModule.Logout();
     Panel_CTINo.Caption := '내선';
     Panel_CTI.Caption := '로그아웃';
     Pnl_CTIUser.Caption := '';
     Pnl_Mod.Caption := '';
     pnl_WaitCnt.Caption := '';
  end;
end;

procedure TfrmCK5I17.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster_Detail.Refresh; //수정
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

   // Excel 변환 Dialog Box를 Open한다.
   GP_Excel(Self);
end;

procedure TfrmCK5I17.btnOpenOfficeClick(Sender: TObject);
begin
   inherited;

   // Excel 변환 Dialog Box를 Open한다.
   GP_OpenOffice(Self);
end;

procedure TfrmCK5I17.FormKeyPress(Sender: TObject; var Key: Char);
begin
   inherited;

   // ESC Key를 누르면 취소버튼 Click
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
       // Variant변수를 사용하기 위해서 생성
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
       // 이미 생성된 변수들의 크기를 조정
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
         // Variant변수를 사용하기 위해서 생성
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
         // 이미 생성된 변수들의 크기를 조정
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
         // Variant변수를 사용하기 위해서 생성
         UV_vInjouk_num_id := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // 이미 생성된 변수들의 크기를 조정
         VarArrayRedim(UV_vInjouk_num_id,      iLength);
      end;
   end
   else if sGubun = 'F' then
   begin
      if iLength = 0 then
      begin
         // Variant변수를 사용하기 위해서 생성
         UV_vFile_name    := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // 이미 생성된 변수들의 크기를 조정
         VarArrayRedim(UV_vFile_name,      iLength);
      end;
   end
   else if sGubun = 'H' then
   begin
      if iLength = 0 then
      begin
         // Variant변수를 사용하기 위해서 생성
         UV_vHul_dat_gmgn    := VarArrayCreate([0, 0], varOleStr);
         UV_vHul_num_id      := VarArrayCreate([0, 0], varOleStr);
      end
      else
      begin
         // 이미 생성된 변수들의 크기를 조정
         VarArrayRedim(UV_vHul_dat_gmgn,      iLength);
         VarArrayRedim(UV_vHul_num_id,        iLength);
      end;
   end
   else if sGubun = 'HG' then
   begin
      if iLength = 0 then
      begin
         // Variant변수를 사용하기 위해서 생성
         UV_vHul_Gulkwa      := VarArrayCreate([0, 50, 0, 0], varOleStr);
         UV_vTot_cod_hm_name := VarArrayCreate([0, 0, 0, 0], varOleStr);
         UV_vHul_Low_High    := VarArrayCreate([0, 0, 0, 0], varOleStr);
         UV_vFont_color      := VarArrayCreate([0, 50, 0, 0], varOleStr);
      end
      else
      begin
         // 이미 생성된 변수들의 크기를 조정
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
         // 이미 생성된 변수들의 크기를 조정
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
         // 이미 생성된 변수들의 크기를 조정
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

   // Grid의 환경 설정
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

   // 항목자료 Query
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
   // Grid 초기화
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
            if      FieldByName('Sawon_cod_jisa').AsString = '11' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (강남센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '12' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (여의도센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '15' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (본원센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '43' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (수원센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '51' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (광주센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '61' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (부산센터)'
            else if FieldByName('Sawon_cod_jisa').AsString = '71' then UV_vDSawon_name[index] := UV_vDSawon_name[index] + ' (대구센터)';

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
      // Table과 Disconnect
      Active := False;
   end;

   // Grid에 자료를 할당
   if index > 0 then
      grdMaster_Detail.Rows := index;

end;

procedure TfrmCK5I17.grdMaster_DetailCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;

   // 자룔를 화면에 조회
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
   // 자료가 존재하는지 Check
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
            //소견
            Cmb_sokun_name.Text := qryJc_consults.FieldByName('sokun_name').AsString;

            //만족도
            if trim(FieldByName('con_type').AsString) = 'N' then
               con_type.ItemIndex := 0
            else if trim(FieldByName('con_type').AsString) = 'G' then
               con_type.ItemIndex := 1
            else if trim(FieldByName('con_type').AsString) = 'C' then
               con_type.ItemIndex := 2
            else if trim(FieldByName('con_type').AsString) = 'E' then
               con_type.ItemIndex := 3
            else con_type.ItemIndex := -1;

            //부재중
            if trim(FieldByName('Gubn_Recall').AsString) = 'Y' then
               chk_ReCall.Checked := True
            Else
               chk_ReCall.Checked := False;

            //상담종류
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
      // Table과 Disconnect
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
   if not GF_MsgBox('Confirmation', '저장 하시겠습니까.?') then exit;
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
   if not GF_MsgBox('Confirmation', '삭제 하시겠습니까.?') then exit;
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

      // Grid 초기화
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
            //query_log남기기...
            GP_query_log(qryGumgin,'CK5I17', 'Q', 'Y');

            //[2018.08.09-유동구]한의 부산CRM팀장1800029 의거]
            //타인에게 개인정보 노출 금지(개인정보노출 차단 요청시)
            //------------------------------------------------------------
            if FieldByName('Ysno_Privacy_Forbid').AsString = 'Y' then
            begin
               PnlYsno_Privacy_Forbid.Visible := True;
               GF_MsgBox('Warning', '[★개인정보노출 사고예방★]' + #13#10#13#10 +
                                    '타인에게 개인정보 노출 금지 요청된 고객입니다.' + #13#10 +
                                    '(본인 외 검진확인불가 및 모든 정보 발설(노출) 금지)');
            end
            else
            begin
               PnlYsno_Privacy_Forbid.Visible := False;
            end;
            //==================================================================

            // DataSet의 값을 Variant변수로 이동
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

               //인적정보 표기..
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
            // 자료가 없음을 표시
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
         // Grid에 자료를 할당
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
            if qryInjouk.RecordCount > 1 then bCK_Num_id := True    //ID 여러개인 경우 True...
            else                              bCK_Num_id := False;

            while Not Eof do
            begin
               UP_VarMemSet(index3, 'I');
               UV_vInjouk_num_id[index3] := FieldByName('num_id').AsString;


               with qryGmgn3 do //2016.02.03 - 박대용 - 인적테이블의 개인정보동의에서 검진테이블 개인정보동의로 변경
               begin
                  Active := False;
                  ParamByName('num_jumin').AsString  := edtJumin.Text;
                  Active := True;

                  if FieldByName('ysno_useinfo').AsString = 'Y' then pnlYsno_useinfo.Visible := False //개인정보 활용 동의
                  else                                               pnlYsno_useinfo.Visible := True;

                  with qryKicho3 do //2016.02.03 - 박대용 - 기초테이블의 가장 최근의 VIP기록 출력
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

      //Sheet 선택 시 재 계산.. 처음부터 표기 필요 없음...
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
      
      //[2017.08.17-유동구]한의 부산CRM팀장1700014의거 작업
      //마지막 상담내용이 아닌 신규창 요청...
      btn_NewClick(Self);
   end;
end;

procedure TfrmCK5I17.grd_fileCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // 자룔를 화면에 조회
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
   else ShowMessage('\\222.222.1.6\chart\ 네트워크 드라이브 연결이 되어있지 않습니다. ' + #13#10 +
                    '네트워크 드라이브 연결후 파일 확인이 가능합니다.');
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
   qryComm06.Filter := ' 물질코드 = ''' + sCod_Hmatter + '''';

   if qryComm06.RecordCount > 0 then Result := qryComm06.FieldByName('유해물질명').AsString;
end;

procedure TfrmCK5I17.UP_GumsaCheck(gubun, sID_1, sJisa_1, sNsGubn_1, sNsCg_1, sDate1, sDate2  : String);
var
   sSelect, sWhere, sOrderBy,
   sTemp, sTemp_NS, sTemp_Tk, sTk_YHGM, sTkCg : String;

begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sTemp := ''; sTemp_NS := ''; sTemp_Tk  := '';  sTk_YHGM:= '';{검진구분} sTkCg := '';{특검청구구분}

   if gubun = '1' then
   begin
      {노신}
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
      {특검}
      //[20160905-유동구]특검 검진구분 찾기...
      //------------------------------------------------------------------------
      //[20170419-유동구]1차 검진 청구 유형 가져오기...
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
         if      pos('TC77', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := '특검[TC77-배치전]'  //배치전
         else if pos('TC93', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := '특검[TC93-배치후]'  //배치후
         else if pos('TC79', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := '특검[TC79-임시]'    //임시
         else if pos('TC78', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := '특검[TC78-수시]'    //수시
         else if pos('TCA7', qryTemp.FieldByName('cod_prf').AsString) > 0 then sTk_YHGM := '특검[TCA7-추적]'    //추적
         else if Trim(sTk_YHGM) = ''                                      then sTk_YHGM := '특검';              //특검

         if      qryTemp.FieldByName('gubn_Tkcg').AsString = '1' then begin sTk_YHGM := sTk_YHGM + '(회청)'; sTkCg := '1'; end
         else if qryTemp.FieldByName('gubn_Tkcg').AsString = '2' then begin sTk_YHGM := sTk_YHGM + '(공청)'; sTkCg := '2'; end;
      end;
      //========================================================================

      //재검 항목 찾기..
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
      {일반}
      if (sTemp_NS <> '') and (sNsGubn_1 <> 'L') and (sNsCg_1 = '2') then
      begin
         if sNsGubn_1 = 'N' then sTemp := '일반검진';
         if sNsGubn_1 = 'A' then sTemp := '성인병검진';
//         if sNsGubn_1 = 'L' then sTemp := '생애전환기검진';
      
         {if      sNsCg_1 = '1' then sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(회사)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else if sNsCg_1 = '2' then sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(조합)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else                       sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(' + sNsCg_1 + ')]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')' ;
         }
         sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(조합)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')';
      end
      else if (sNsGubn_1 = 'L') and (sNsCg_1 = '2') then
      begin
         if sNsGubn_1 = 'L' then sTemp := '생애전환기검진';
         {if      sNsCg_1 = '1' then sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(회사)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else if sNsCg_1 = '2' then sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(조합)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')'
         else                       sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(' + sNsCg_1 + ')]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')' ;
         }
         sTemp := sTemp + '[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')|청구(조합)]' + '재검(' + copy(sTemp_NS,1,length(sTemp_NS)-1) + ')';

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
                  sTemp := sTemp + #10#13#13 + '생애재검 대상, 1차 소견정리 안됨';
               end;
            end
            else
            begin
                  sTemp := sTemp + #10#13#13 + '생애재검 대상, 1차 소견정리 안됨';
            end;
            Active := False;
         End;
      end;
   end;

   {if gubun = '2' then
   begin
      if sTemp_Tk <> '' then GP_ComboDisplay(cmbGUBN_TKCG, sTkCg, 1);
      //특검
      if (sTemp_Tk <> '') and (sTemp_NS <> '') then sTemp := sTemp + #10#13 + '특검[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')]' + '재검(' + copy(sTemp_Tk,1,length(sTemp_Tk)-1) + ')'
      else if sTemp_Tk <> '' then sTemp := '특검[센터(' + sJisa_1 + ')|검진일(' + sDate1 + ')]' + '재검(' + copy(sTemp_Tk,1,length(sTemp_Tk)-1) + ')';
   end;}

   if sTemp <> '' then ShowMessage(sTemp + #10#13#13 + ' ## ' + sTk_YHGM + ' 재검대상자 입니다.');
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
     // Table과 Disconnect
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
        // 자료가 없음을 표시
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
     UV_vSinjang[0]     := '신        장   (cm)';
     UV_vChejung[0]     := '체        중   (kg)';
     UV_vBiman[0]       := '비    만   도   (%)';
     UV_vHul_h[0]       := '최고 혈압    (mmHg)';
     UV_vHul_l[0]       := '최저 혈압    (mmHg)';
     UV_vear_left5[0]   := '청력(좌 500Hz dB)';
     UV_vear_left10[0]  := '청력(좌1000Hz dB)';
     UV_vear_lef20[0]   := '청력(좌2000Hz dB)';
     UV_vear_left30[0]  := '청력(좌3000Hz dB)';
     UV_vear_left40[0]  := '청력(좌4000Hz dB)';
     UV_vear_left60[0]  := '청력(좌6000Hz dB)';
     UV_vear_right5[0]  := '청력(우 500Hz dB)';
     UV_vear_right10[0] := '청력(우1000Hz dB)';
     UV_vear_right20[0] := '청력(우2000Hz dB)';
     UV_vear_right30[0] := '청력(우3000Hz dB)';
     UV_vear_right40[0] := '청력(우4000Hz dB)';
     UV_vear_right60[0] := '청력(우6000Hz dB)';
     UV_vdesc_ear[0]    := '청력소견';
     UV_veye_lman[0]    := '시         력(좌)';
     UV_veye_rman[0]    := '시         력(우)';
     UV_veye_lkyo[0]    := '교 정   시 력(좌)';
     UV_veye_rkyo[0]    := '교 정   시 력(우)';
     UV_vanap_left[0]   := '안   압(좌)  (mmHg)';
     UV_vanap_right[0]  := '안   압(우)  (mmHg)';
     UV_vgubn_seksin[0] := '색         신';
     UV_vhyungwi[0]     := '흉위(가슴둘레)';
     UV_vbookwi[0]      := '복위(허리둘레)';
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
        // 자료가 없음을 표시
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid에 자료를 할당
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

   // Grid의 환경 설정
   with grd_Kicho do
   begin
      // Column Title
      Col[1].Heading := '검사항목';
      Col[1].width   := 136;

      if DataCol > 1 then
      begin
         Col[DataCol].Heading := UV_vDat_gmgn[DataCol-1];
         Col[DataCol].width   := 80;
      end;
   end;

   // 자룔를 화면에 조회
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
    // 임상참고치
    eLow, eHigh, eIms : Extended;
    sMunja : String;
    vCod_Tkgum : Variant;
begin
  inherited;

   // Grid 초기화
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

      //혈액검사항목추출
      sCod_hm := '';
      Index   := 0;
      if qryHul_gumgin.RecordCount > 0 then
      begin
         while not qryHul_gumgin.Eof do
         begin
            UP_VarMemSet(index, 'H');

            UV_vHul_Num_id[index]    := FieldByName('num_id').AsString;
            UV_vHul_dat_gmgn[index]  := FieldByName('dat_gmgn').AsString;

            //검진일자 표기..
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

            // 노신, 성인병, 간염에 대한 검사항목 추출
            cod_name := '';
            // 노신유형Display
            if FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            //성인병유형Display
            if FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            //간염유형Display
            if FieldByName('gubn_agab').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            //생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(FieldByName('dat_gmgn').AsString,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            //특검유형Display
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
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   UP_VarMemSet(iRet, 'HG');
   for I := 0 to index-1 do
   begin
      // 주민번호를 이용하여 성별과 나이를 구함
      GP_JuminToAge(edtJumin.Text,UV_vHul_dat_gmgn[I], sMan, iAge);
      
      // 성별과 나이를 기준으로 Column명을 추출
      sGubn := '';
      if iAge < 10 then sGubn := '1'
      else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
      else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
      else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
      else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
      else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
      else sGubn := '7';
      
      if sMan = 'F' then sGubn := sGubn + 'f';

      // 분주번호에 대한 혈액결과를 획득
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
         // 초기값 설정
         eLow   := 0;
         eHigh  := 0;
         eIms   := 0;
         sMunja := '';

         // 검사항목에 대한 임상참고치를 획득
         qryHm.Filter := 'COD_HM = ''' + UV_vTot_cod_hm[J] + ''' AND ' +
                         'DAT_APPLY <= ''' + UV_vHul_dat_gmgn[I] + '''';

         //항목명
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

         // 혈액검사인지 Check(장비PF, 추가검사 때문)
         if qryHm.FieldByName('GUBN_PART').AsString > '10' then continue;

         eLow   := GF_StrToFloat(qryHm.FieldByName('IMS_LOW'   + sGubn).AsString);
         eHigh  := GF_StrToFloat(qryHm.FieldByName('IMS_HIGH'  + sGubn).AsString);
         sMunja := qryHm.FIeldByName('IMS_MUNJA').AsString;

         // 해당 검사에 대한 결과를 추출(Start Pos, End Pos를 이용)
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

   // Grid에 자료를 할당
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
   // Grid의 환경 설정
   with grd_Hul do
   begin
      J := grd_Hul.Rows;
      I := grd_Hul.Cols;

      // Column Title
      Col[1].Heading := '코드';
      Col[2].Heading := '항목명';
      Col[3].Heading := '임상참고치';

      // Column Align
      Col[1].Alignment := taCenter;
      Col[3].Alignment := taCenter;

      if DataCol > 3 then
      begin
         Col[DataCol].Heading := Copy(UV_vHul_dat_gmgn[DataCol-4],1,4) + '-' + Copy(UV_vHul_dat_gmgn[DataCol-4],5,2) + '-' + Copy(UV_vHul_dat_gmgn[DataCol-4],7,2);
         Col[DataCol].width   := 80;
      end;
   end;

   // 자룔를 화면에 조회
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
   // 자료가 존재하는지 Check
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
        // 자료가 없음을 표시
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid에 자료를 할당
   if index > 0 then
      grd_Jangbi.Rows := index;
end;

procedure TfrmCK5I17.grd_JangbiCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // 자룔를 화면에 조회
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
   // 자료가 존재하는지 Check
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

           if      FieldByName('gubn_hsok').AsString = '1'         then UV_vgubn_sokun[index] := '혈액소견'
           else if FieldByName('gubn_hsok').AsString = '2'         then UV_vgubn_sokun[index] := '식이요법'
           else if FieldByName('gubn_hsok').AsString = '3'         then UV_vgubn_sokun[index] := '운동요법'
           else if FieldByName('gubn_hsok').AsString = 'tot_sokun' then UV_vgubn_sokun[index] := '종합소견';

           UV_vCod_sokun[index]  := FieldByName('cod_sokun').AsString;
           UV_vDesc_sokun[index] := FieldByName('desc_sokun').AsString;
           Next;
           Inc(index);
        end;
     end
     else
     begin
        // 자료가 없음을 표시
        GF_MsgBox('Information', 'NODATA');
        exit;
     end;
  end;

   // Grid에 자료를 할당
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

        // PAP자료 Query
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
         if      Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'C' then Panel_PAP.Visible := False //판독완료
         else if Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'Y' then Panel_PAP.Visible := True  //*판독중*
         else Panel_PAP.Visible := True;                                                                //미판독
      end
      else
      begin
         if (Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'Y') or
            (Trim(qryCell.FieldByname('Ysno_sokun').AsString) = 'C') then Panel_PAP.Visible := False     //판독완료
         else
         begin
            if qryCell.FieldByname('Dat_result').AsString <> '' then Panel_PAP.Visible := False          //판독완료
            else Panel_PAP.Visible := True;                                                              //미판독
         end;
      end;
   end;
end;

procedure TfrmCK5I17.CreateTextArrayTot_Munjin;
var
   i : Integer;
   sMunjinTemp, sMunjinTemp2 : string;   
begin

   //[2019.05.29-유동구]2019년도 문진 적용(담배부분)  ..2019.06.12 추가 (정부업수석)
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
   // 건강검진 문진
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

   //특정암문진
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

  // 여성검진 특정암문진...
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
          //[2017.06.20-유동구]종합결과지 과거력 표기....
          //--------------------------------------------------------------------
          Mm_munjin2.Lines.Clear;
          sMunjinTemp := '';
          if FieldByName('mun1_jindan1').AsString = '1' then sMunjinTemp := sMunjinTemp + '뇌졸중(중풍),';
          if FieldByName('mun1_jindan2').AsString = '1' then sMunjinTemp := sMunjinTemp + '심장병(심근경색/협심증),';
          if FieldByName('mun1_jindan3').AsString = '1' then sMunjinTemp := sMunjinTemp + '고혈압,';
          if FieldByName('mun1_jindan4').AsString = '1' then sMunjinTemp := sMunjinTemp + '당뇨병,';
          if FieldByName('mun1_jindan5').AsString = '1' then sMunjinTemp := sMunjinTemp + '이상지질혈증,';
          if FieldByName('mun1_jindan6').AsString = '1' then sMunjinTemp := sMunjinTemp + '폐결핵,';
          if FieldByName('mun1_jindan7').AsString = '1' then
          begin
             sMunjinTemp2 := '';
             if copy(FieldByName('CMun3_Can1_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '위암,';
             if copy(FieldByName('CMun3_Can2_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '유방암,';
             if copy(FieldByName('CMun3_Can3_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '대장암,';
             if copy(FieldByName('CMun3_Can4_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '간암,';
             if copy(FieldByName('CMun3_Can5_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '자궁경부암,';
             if copy(FieldByName('CMun3_Can6_Family').AsString,1,1) = '2' then
             begin
                if pos('암',FieldByName('CMun3_Can6_Bung').AsString) > 0 then sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + ','
                else                                                          sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + '암,';
             end;

             if sMunjinTemp2 = '' then sMunjinTemp := sMunjinTemp + '암,'
             else                      sMunjinTemp := sMunjinTemp + sMunjinTemp2;
          end;
          if FieldByName('mun1_jindan8').AsString = '1' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then  Mm_munjin2.Lines.Add('⊙질병명(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun5').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '위궤양,';
          if copy(FieldByName('CMun5').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '위축성위염,';
          if copy(FieldByName('CMun5').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '장상피화생,';
          if copy(FieldByName('CMun5').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '위용종,';
          if copy(FieldByName('CMun5').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('⊙위장질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun6').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '대장용종,';
          if copy(FieldByName('CMun6').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '궤양성대장염,';
          if copy(FieldByName('CMun6').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '크론병,';
          if copy(FieldByName('CMun6').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '치질(치핵,치열),';
          if copy(FieldByName('CMun6').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('⊙대장항문질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun7').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + 'B형간염보균자,';
          if copy(FieldByName('CMun7').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '만성B형간염,';
          if copy(FieldByName('CMun7').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '만성C형간염,';
          if copy(FieldByName('CMun7').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '간경변,';
          if copy(FieldByName('CMun7').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2.Lines.Add('⊙간질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');
          //====================================================================

          // 건강검진문진.
          //1번문항 진단
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

          //1번문항 약물
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

          //2번문항
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

          //3번문항 B형간염
          if FieldByName('Mun3').AsString <> ''  then
             Mun3.itemindex  := FieldByName('Mun3').AsInteger;

          //4번문항
          if FieldByName('Mun4_1').AsString <> ''  then
             Mun4_1.itemindex  := FieldByName('Mun4_1').AsInteger;
          Mun4_2_Year.Text := FieldByName('Mun4_2_Year').AsString;
          Mun4_2_Day.Text  := FieldByName('Mun4_2_Day').AsString;
          Mun4_3_Year.Text := FieldByName('Mun4_3_Year').AsString;
          Mun4_3_Day.Text  := FieldByName('Mun4_3_Day').AsString;

          //5번문항
          if FieldByName('Mun5_1').AsString <> ''  then
             Mun5_1.itemindex  := FieldByName('Mun5_1').AsInteger;
          if FieldByName('Mun5_2').AsString <> ''  then
             Mun5_2.itemindex  := FieldByName('Mun5_2').AsInteger;
          if FieldByName('Mun5_3_1').AsString <> '' then
             Mun5_3_1.itemindex  := FieldByName('Mun5_3_1').AsInteger;
          Mun5_3_2.Text  := FieldByName('Mun5_3_2').AsString;
          Mun5_4.Text    := FieldByName('Mun5_4').AsString;

          //6번문항
          if FieldByName('Mun6_1').AsString <> ''  then
             Mun6_1.itemindex  := FieldByName('Mun6_1').AsInteger;
          if FieldByName('Mun6_2').AsString <> ''  then
             Mun6_2.itemindex  := FieldByName('Mun6_2').AsInteger;
          if FieldByName('Mun6_3').AsString <> ''  then
             Mun6_3.itemindex  := FieldByName('Mun6_3').AsInteger;
          if FieldByName('Mun6_4').AsString <> ''  then
             Mun6_4.itemindex  := FieldByName('Mun6_4').AsInteger;

          //8번문항
          if FieldByName('Mun8_1').AsString <> ''  then
             Mun8_1.itemindex  := FieldByName('Mun8_1').AsInteger;
          if FieldByName('Mun8_2').AsString <> ''  then
             Mun8_2.itemindex  := FieldByName('Mun8_2').AsInteger;
          if FieldByName('Mun8_3').AsString <> ''  then
             Mun8_3.itemindex  := FieldByName('Mun8_3').AsInteger;
          if FieldByName('Mun8_4').AsString <> ''  then
             Mun8_4.itemindex  := FieldByName('Mun8_4').AsInteger;

          // 특정암문진
          // 1번문항
          if FieldByName('CMun1').AsString <> ''  then
             CMun1.itemindex  := FieldByName('CMun1').AsInteger;
          CMun1_Bung.Text := FieldByName('CMun1_Bung').AsString;

          // 2번문항
          if FieldByName('CMun2').AsString <> ''  then
             CMun2.itemindex  := FieldByName('CMun2').AsInteger;
          CMun2_kg.Text := FieldByName('CMun2_kg').AsString;

          // 3번문항
          //위암
          if FieldByName('CMun3_Can1_Yn').AsString <> ''  then
             CMun3_Can1_Yn.itemindex  := FieldByName('CMun3_Can1_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 1, 1) = '2' then
             CMun3_Can1_F1.checked := true else CMun3_Can1_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 2, 1) = '2' then
             CMun3_Can1_F2.checked := true else CMun3_Can1_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 3, 1) = '2' then
             CMun3_Can1_F3.checked := true else CMun3_Can1_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 4, 1) = '2' then
             CMun3_Can1_F4.checked := true else CMun3_Can1_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 5, 1) = '2' then
             CMun3_Can1_F5.checked := true else CMun3_Can1_F5.checked := false;

          //유방암
          if FieldByName('CMun3_Can2_Yn').AsString <> ''  then
             CMun3_Can2_Yn.itemindex  := FieldByName('CMun3_Can2_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 1, 1) = '2' then
             CMun3_Can2_F1.checked := true else CMun3_Can2_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 2, 1) = '2' then
             CMun3_Can2_F2.checked := true else CMun3_Can2_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 3, 1) = '2' then
             CMun3_Can2_F3.checked := true else CMun3_Can2_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 4, 1) = '2' then
             CMun3_Can2_F4.checked := true else CMun3_Can2_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 5, 1) = '2' then
             CMun3_Can2_F5.checked := true else CMun3_Can2_F5.checked := false;

          //대장암
          if FieldByName('CMun3_Can3_Yn').AsString <> ''  then
             CMun3_Can3_Yn.itemindex  := FieldByName('CMun3_Can3_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 1, 1) = '2' then
             CMun3_Can3_F1.checked := true else CMun3_Can3_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 2, 1) = '2' then
             CMun3_Can3_F2.checked := true else CMun3_Can3_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 3, 1) = '2' then
             CMun3_Can3_F3.checked := true else CMun3_Can3_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 4, 1) = '2' then
             CMun3_Can3_F4.checked := true else CMun3_Can3_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 5, 1) = '2' then
             CMun3_Can3_F5.checked := true else CMun3_Can3_F5.checked := false;

          //간암
          if FieldByName('CMun3_Can4_Yn').AsString <> ''  then
             CMun3_Can4_Yn.itemindex  := FieldByName('CMun3_Can4_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 1, 1) = '2' then
             CMun3_Can4_F1.checked := true else CMun3_Can4_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 2, 1) = '2' then
             CMun3_Can4_F2.checked := true else CMun3_Can4_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 3, 1) = '2' then
             CMun3_Can4_F3.checked := true else CMun3_Can4_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 4, 1) = '2' then
             CMun3_Can4_F4.checked := true else CMun3_Can4_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 5, 1) = '2' then
             CMun3_Can4_F5.checked := true else CMun3_Can4_F5.checked := false;

          //자궁경부암
          if FieldByName('CMun3_Can5_Yn').AsString <> ''  then
             CMun3_Can5_Yn.itemindex  := FieldByName('CMun3_Can5_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 1, 1) = '2' then
             CMun3_Can5_F1.checked := true else CMun3_Can5_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 2, 1) = '2' then
             CMun3_Can5_F2.checked := true else CMun3_Can5_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 3, 1) = '2' then
             CMun3_Can5_F3.checked := true else CMun3_Can5_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 4, 1) = '2' then
             CMun3_Can5_F4.checked := true else CMun3_Can5_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 5, 1) = '2' then
             CMun3_Can5_F5.checked := true else CMun3_Can5_F5.checked := false;

          //기타암
          CMun3_Can6_Bung.Text := FieldByName('CMun3_Can6_Bung').AsString;
          if FieldByName('CMun3_Can6_Yn').AsString <> ''  then
             CMun3_Can6_Yn.itemindex  := FieldByName('CMun3_Can6_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 1, 1) = '2' then
             CMun3_Can6_F1.checked := true else CMun3_Can6_F1.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 2, 1) = '2' then
             CMun3_Can6_F2.checked := true else CMun3_Can6_F2.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 3, 1) = '2' then
             CMun3_Can6_F3.checked := true else CMun3_Can6_F3.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 4, 1) = '2' then
             CMun3_Can6_F4.checked := true else CMun3_Can6_F4.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 5, 1) = '2' then
             CMun3_Can6_F5.checked := true else CMun3_Can6_F5.checked := false;

          // 4번문항
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

          // 5번문항
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

          // 6번문항
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

          // 7번문항
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

          // 여성수검자만 해당

          // 8번문항
          if FieldByName('CMun8').AsString <> ''  then
             CMun8.itemindex  := FieldByName('CMun8').AsInteger;
          CMun8_Year.Text := FieldByName('CMun8_Year').AsString;

          // 9번문항
          if FieldByName('CMun9').AsString <> ''  then
             CMun9.itemindex  := FieldByName('CMun9').AsInteger;
          CMun9_Year.Text := FieldByName('CMun9_Year').AsString;

          // 10번문항
          if FieldByName('CMun10').AsString <> ''  then
             CMun10.itemindex  := FieldByName('CMun10').AsInteger;

          // 11번문항
          if FieldByName('CMun11').AsString <> ''  then
             CMun11.itemindex  := FieldByName('CMun11').AsInteger;

          // 12번문항
          if FieldByName('CMun12').AsString <> ''  then
             CMun12.itemindex  := FieldByName('CMun12').AsInteger;

          // 13번문항
          if FieldByName('CMun13').AsString <> ''  then
             CMun13.itemindex  := FieldByName('CMun13').AsInteger;

          // 14번문항
          if FieldByName('CMun14').AsString <> ''  then
             CMun14.itemindex  := FieldByName('CMun14').AsInteger;

          //정보활용 동의여부
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
   // 건강검진 문진
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


   //특정암문진
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

  // 여성검진 특정암문진...
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

   // 현재위치를 추출
   i := grdMaster.CurrentDataRow - 1;

   //[2019.05.29-유동구]2019년도 문진 적용(담배부분)  ..19.06.12
   if copy(cmb_gmgn.text,1,4) = '2018' then
   begin
      Pnl_Mun_2019.Visible := False;

      Label192.Caption := '6. 술마시는 횟수';
      Label195.Caption := '6-1. 술마시는 양';
      Label177.Caption := '6-2. 최대 음주양';
      Label190.Caption := '7-1. 평소 1주일간, 고강도 활동 일수?    주당               일';
      Label196.Caption := '7-2. 평소 고강도 활동 시간?   하루에          시간          분';
      Label198.Caption := '8-1. 평소 1주일간, 중강도 활동 일수?    주당               일';
      Label197.Caption := '8-2. 평소 중강도 활동 시간?   하루에          시간          분';
      Label199.Caption := '9. 최근 근력운동 날 수?                     주당               일';
   end
   else if copy(cmb_gmgn.text,1,4) = '2019' then
   begin
      Pnl_Mun_2019.Visible := True;

      Label192.Caption := '7. 술마시는 횟수';
      Label195.Caption := '7-1. 술마시는 양';
      Label177.Caption := '7-2. 최대 음주양';
      Label190.Caption := '8-1. 평소 1주일간, 고강도 활동 일수?    주당               일';
      Label196.Caption := '8-2. 평소 고강도 활동 시간?   하루에          시간          분';
      Label198.Caption := '9-1. 평소 1주일간, 중강도 활동 일수?    주당               일';
      Label197.Caption := '9-2. 평소 중강도 활동 시간?   하루에          시간          분';
      Label199.Caption := '10. 최근 근력운동 날 수?                    주당               일';
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
      
          //[2017.06.20-유동구]종합결과지 과거력 표기....
          //--------------------------------------------------------------------
          Mm_munjin2_2018.Lines.Clear;
          sMunjinTemp := '';
          if FieldByName('mun1_jindan1').AsString = '1' then sMunjinTemp := sMunjinTemp + '뇌졸중(중풍),';
          if FieldByName('mun1_jindan2').AsString = '1' then sMunjinTemp := sMunjinTemp + '심장병(심근경색/협심증),';
          if FieldByName('mun1_jindan3').AsString = '1' then sMunjinTemp := sMunjinTemp + '고혈압,';
          if FieldByName('mun1_jindan4').AsString = '1' then sMunjinTemp := sMunjinTemp + '당뇨병,';
          if FieldByName('mun1_jindan5').AsString = '1' then sMunjinTemp := sMunjinTemp + '이상지질혈증,';
          if FieldByName('mun1_jindan6').AsString = '1' then sMunjinTemp := sMunjinTemp + '폐결핵,';
          if FieldByName('mun1_jindan7').AsString = '1' then
          begin
             sMunjinTemp2 := '';
             if copy(FieldByName('CMun3_Can1_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '위암,';
             if copy(FieldByName('CMun3_Can2_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '유방암,';
             if copy(FieldByName('CMun3_Can3_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '대장암,';
             if copy(FieldByName('CMun3_Can4_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '간암,';
             if copy(FieldByName('CMun3_Can5_Family').AsString,1,1) = '2' then sMunjinTemp2 := sMunjinTemp2 + '자궁경부암,';
             if copy(FieldByName('CMun3_Can6_Family').AsString,1,1) = '2' then
             begin
                if pos('암',FieldByName('CMun3_Can6_Bung').AsString) > 0 then sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + ','
                else                                                          sMunjinTemp2 := sMunjinTemp2 + FieldByName('CMun3_Can6_Bung').AsString + '암,';
             end;

             if sMunjinTemp2 = '' then sMunjinTemp := sMunjinTemp + '기타(암포함),'
             else                      sMunjinTemp := sMunjinTemp + sMunjinTemp2;
          end;
          if Trim(sMunjinTemp) <> '' then  Mm_munjin2_2018.Lines.Add('⊙질병명(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun5').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '위궤양,';
          if copy(FieldByName('CMun5').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '위축성위염,';
          if copy(FieldByName('CMun5').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '장상피화생,';
          if copy(FieldByName('CMun5').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '위용종,';
          if copy(FieldByName('CMun5').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('⊙위장질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun6').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + '대장용종,';
          if copy(FieldByName('CMun6').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '궤양성대장염,';
          if copy(FieldByName('CMun6').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '크론병,';
          if copy(FieldByName('CMun6').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '치질(치핵,치열),';
          if copy(FieldByName('CMun6').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('⊙대장항문질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          sMunjinTemp := '';
          if copy(FieldByName('CMun7').AsString,1,1) = '2' then sMunjinTemp := sMunjinTemp + 'B형간염보균자,';
          if copy(FieldByName('CMun7').AsString,2,1) = '2' then sMunjinTemp := sMunjinTemp + '만성B형간염,';
          if copy(FieldByName('CMun7').AsString,3,1) = '2' then sMunjinTemp := sMunjinTemp + '만성C형간염,';
          if copy(FieldByName('CMun7').AsString,4,1) = '2' then sMunjinTemp := sMunjinTemp + '간경변,';
          if copy(FieldByName('CMun7').AsString,5,1) = '2' then sMunjinTemp := sMunjinTemp + '기타,';
          if Trim(sMunjinTemp) <> '' then Mm_munjin2_2018.Lines.Add('⊙간질환(' + copy(sMunjinTemp,1,length(sMunjinTemp)-1) + ')');

          //일반문진
          //1번문항 진단
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

          //1번문항 약물
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

          //2번문항
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

          //3번문항 B형간염
          if FieldByName('Mun3').AsString <> ''  then
             Mun3_2018.itemindex  := FieldByName('Mun3').AsInteger;

          //[2019.05.29-유동구]2019년도 문진 적용(담배부분) ..19.06.12 추가
          if copy(cmb_gmgn.text,1,4) = '2018' then
          begin
             //4번문항
             if FieldByName('Mun4').AsString        <> '' then Mun4_2018.itemindex  := FieldByName('Mun4').AsInteger;
             if FieldByName('Mun4_1_year').AsString <> '' then Mun4_1_year_2018.Text  := FieldByName('Mun4_1_year').AsString;
             if FieldByName('Mun4_1_day').AsString  <> '' then Mun4_1_day_2018.Text  := FieldByName('Mun4_1_day').AsString;
             if FieldByName('Mun4_2_year').AsString <> '' then Mun4_2_year_2018.Text  := FieldByName('Mun4_2_year').AsString;
             if FieldByName('Mun4_2_day').AsString  <> '' then Mun4_2_day_2018.Text  := FieldByName('Mun4_2_day').AsString;

             //5번문항
             if FieldByName('Mun5').AsString   <> '' then Mun5_2018.itemindex   := FieldByName('Mun5').AsInteger;
             if FieldByName('Mun5_1').AsString <> '' then Mun5_1_2018.itemindex := FieldByName('Mun5_1').AsInteger;
          end
          else  if copy(cmb_gmgn.text,1,4) = '2019' then
          begin
             //4번문항
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

             //5번문항
             if FieldByName('Mun5').AsString   <> '' then Mun5_2019.itemindex   := FieldByName('Mun5').AsInteger;
             if FieldByName('Mun5_1').AsString <> '' then Mun5_1_2019.itemindex := FieldByName('Mun5_1').AsInteger;
          end;

          //6번문항
          if FieldByName('Mun6_num').AsString <> ''  then
             Mun6_num_2018.itemindex  := FieldByName('Mun6_num').AsInteger;
          if FieldByName('Mun6_count').AsString <> ''  then
             Mun6_count_2018.Text  := FieldByName('Mun6_count').AsString;


          if FieldByName('Mun6_1_soju_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '소주' + FieldByName('Mun6_1_soju_shot').AsString + '잔/';
          if FieldByName('Mun6_1_soju_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '소주' + FieldByName('Mun6_1_soju_btl').AsString + '병/';
          if FieldByName('Mun6_1_soju_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '소주' + FieldByName('Mun6_1_soju_can').AsString + '캔/';
          if FieldByName('Mun6_1_soju_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '소주' + FieldByName('Mun6_1_soju_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_beer_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '맥주' + FieldByName('Mun6_1_beer_shot').AsString + '잔/';
          if FieldByName('Mun6_1_beer_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '맥주' + FieldByName('Mun6_1_beer_btl').AsString + '병/';
          if FieldByName('Mun6_1_beer_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '맥주' + FieldByName('Mun6_1_beer_can').AsString + '캔/';
          if FieldByName('Mun6_1_beer_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '맥주' + FieldByName('Mun6_1_beer_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_liquor_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '양주' + FieldByName('Mun6_1_liquor_shot').AsString + '잔/';
          if FieldByName('Mun6_1_liquor_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '양주' + FieldByName('Mun6_1_liquor_btl').AsString + '병/';
          if FieldByName('Mun6_1_liquor_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '양주' + FieldByName('Mun6_1_liquor_can').AsString + '캔/';
          if FieldByName('Mun6_1_liquor_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '양주' + FieldByName('Mun6_1_liquor_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_makgeolli_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '막걸리' + FieldByName('Mun6_1_makgeolli_shot').AsString + '잔/';
          if FieldByName('Mun6_1_makgeolli_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '막걸리' + FieldByName('Mun6_1_makgeolli_btl').AsString + '병/';
          if FieldByName('Mun6_1_makgeolli_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '막걸리' + FieldByName('Mun6_1_makgeolli_can').AsString + '캔/';
          if FieldByName('Mun6_1_makgeolli_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '막걸리' + FieldByName('Mun6_1_makgeolli_cc').AsString + 'CC/';

          if FieldByName('Mun6_1_wine_shot').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '와인' + FieldByName('Mun6_1_wine_shot').AsString + '잔/';
          if FieldByName('Mun6_1_wine_btl').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '와인' + FieldByName('Mun6_1_wine_btl').AsString + '병/';
          if FieldByName('Mun6_1_wine_can').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '와인' + FieldByName('Mun6_1_wine_can').AsString + '캔/';
          if FieldByName('Mun6_1_wine_cc').AsString <> ''  then
             Mun6_1_2018.Text := Mun6_1_2018.Text + '와인' + FieldByName('Mun6_1_wine_cc').AsString + 'CC/';
             

          //////////////////////////////////////
          if FieldByName('Mun6_2_soju_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '소주' + FieldByName('Mun6_2_soju_shot').AsString + '잔/';
          if FieldByName('Mun6_2_soju_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '소주' + FieldByName('Mun6_2_soju_btl').AsString + '병/';
          if FieldByName('Mun6_2_soju_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '소주' + FieldByName('Mun6_2_soju_can').AsString + '캔/';
          if FieldByName('Mun6_2_soju_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '소주' + FieldByName('Mun6_2_soju_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_beer_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '맥주' + FieldByName('Mun6_2_beer_shot').AsString + '잔/';
          if FieldByName('Mun6_2_beer_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '맥주' + FieldByName('Mun6_2_beer_btl').AsString + '병/';
          if FieldByName('Mun6_2_beer_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '맥주' + FieldByName('Mun6_2_beer_can').AsString + '캔/';
          if FieldByName('Mun6_2_beer_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '맥주' + FieldByName('Mun6_2_beer_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_liquor_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '양주' + FieldByName('Mun6_2_liquor_shot').AsString + '잔/';
          if FieldByName('Mun6_2_liquor_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '양주' + FieldByName('Mun6_2_liquor_btl').AsString + '병/';
          if FieldByName('Mun6_2_liquor_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '양주' + FieldByName('Mun6_2_liquor_can').AsString + '캔/';
          if FieldByName('Mun6_2_liquor_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '양주' + FieldByName('Mun6_2_liquor_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_makgeolli_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '막걸리' + FieldByName('Mun6_2_makgeolli_shot').AsString + '잔/';
          if FieldByName('Mun6_2_makgeolli_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '막걸리' + FieldByName('Mun6_2_makgeolli_btl').AsString + '병/';
          if FieldByName('Mun6_2_makgeolli_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '막걸리' + FieldByName('Mun6_2_makgeolli_can').AsString + '캔/';
          if FieldByName('Mun6_2_makgeolli_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '막걸리' + FieldByName('Mun6_2_makgeolli_cc').AsString + 'CC/';

          if FieldByName('Mun6_2_wine_shot').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '와인' + FieldByName('Mun6_2_wine_shot').AsString + '잔/';
          if FieldByName('Mun6_2_wine_btl').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '와인' + FieldByName('Mun6_2_wine_btl').AsString + '병/';
          if FieldByName('Mun6_2_wine_can').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '와인' + FieldByName('Mun6_2_wine_can').AsString + '캔/';
          if FieldByName('Mun6_2_wine_cc').AsString <> ''  then
             Mun6_2_2018.Text := Mun6_2_2018.Text + '와인' + FieldByName('Mun6_2_wine_cc').AsString + 'CC/';

          //7번문항
          if FieldByName('Mun7_1').AsString <> ''  then
             Mun7_1_2018.itemindex  := StrToInt(FieldByName('Mun7_1').AsString);
          if FieldByName('Mun7_2').AsString <> ''  then
          begin
             Mun7_2_H_2018.Text  := copy(FieldByName('Mun7_2').AsString, 1, 2);
             Mun7_2_M_2018.Text  := copy(FieldByName('Mun7_2').AsString, 3, 2);
          end;


          //8번문항
          if FieldByName('Mun8_1').AsString <> ''  then
             Mun8_1_2018.itemindex  := StrToInt(FieldByName('Mun8_1').AsString);
          if FieldByName('Mun8_2').AsString <> ''  then
          begin
             Mun8_2_H_2018.Text  := copy(FieldByName('Mun8_2').AsString, 1, 2);
             Mun8_2_M_2018.Text  := copy(FieldByName('Mun8_2').AsString, 3, 2);
          end;

          //9번문항
          if FieldByName('Mun9').AsString <> ''  then
             Mun9_2018.itemindex  := StrToInt(FieldByName('Mun9').AsString);
             
          // 특정암문진
          // 1번문항
          if FieldByName('CMun1').AsString <> ''  then
             CMun1_2018.itemindex  := FieldByName('CMun1').AsInteger;
          CMun1_Bung_2018.Text := FieldByName('CMun1_Bung').AsString;

          // 2번문항
          if FieldByName('CMun2').AsString <> ''  then
             CMun2_2018.itemindex  := FieldByName('CMun2').AsInteger;
          CMun2_kg_2018.Text := FieldByName('CMun2_kg').AsString;

          // 3번문항
          //위암
          if FieldByName('CMun3_Can1_Yn').AsString <> ''  then
             CMun3_Can1_Yn_2018.itemindex  := FieldByName('CMun3_Can1_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 1, 1) = '2' then
             CMun3_Can1_F1_2018.checked := true else CMun3_Can1_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 2, 1) = '2' then
             CMun3_Can1_F2_2018.checked := true else CMun3_Can1_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 3, 1) = '2' then
             CMun3_Can1_F3_2018.checked := true else CMun3_Can1_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 4, 1) = '2' then
             CMun3_Can1_F4_2018.checked := true else CMun3_Can1_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can1_Family').AsString, 5, 1) = '2' then
             CMun3_Can1_F5_2018.checked := true else CMun3_Can1_F5_2018.checked := false;

          //유방암
          if FieldByName('CMun3_Can2_Yn').AsString <> ''  then
             CMun3_Can2_Yn_2018.itemindex  := FieldByName('CMun3_Can2_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 1, 1) = '2' then
             CMun3_Can2_F1_2018.checked := true else CMun3_Can2_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 2, 1) = '2' then
             CMun3_Can2_F2_2018.checked := true else CMun3_Can2_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 3, 1) = '2' then
             CMun3_Can2_F3_2018.checked := true else CMun3_Can2_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 4, 1) = '2' then
             CMun3_Can2_F4_2018.checked := true else CMun3_Can2_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can2_Family').AsString, 5, 1) = '2' then
             CMun3_Can2_F5_2018.checked := true else CMun3_Can2_F5_2018.checked := false;

          //대장암
          if FieldByName('CMun3_Can3_Yn').AsString <> ''  then
             CMun3_Can3_Yn_2018.itemindex  := FieldByName('CMun3_Can3_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 1, 1) = '2' then
             CMun3_Can3_F1_2018.checked := true else CMun3_Can3_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 2, 1) = '2' then
             CMun3_Can3_F2_2018.checked := true else CMun3_Can3_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 3, 1) = '2' then
             CMun3_Can3_F3_2018.checked := true else CMun3_Can3_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 4, 1) = '2' then
             CMun3_Can3_F4_2018.checked := true else CMun3_Can3_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can3_Family').AsString, 5, 1) = '2' then
             CMun3_Can3_F5_2018.checked := true else CMun3_Can3_F5_2018.checked := false;

          //간암
          if FieldByName('CMun3_Can4_Yn').AsString <> ''  then
             CMun3_Can4_Yn_2018.itemindex  := FieldByName('CMun3_Can4_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 1, 1) = '2' then
             CMun3_Can4_F1_2018.checked := true else CMun3_Can4_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 2, 1) = '2' then
             CMun3_Can4_F2_2018.checked := true else CMun3_Can4_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 3, 1) = '2' then
             CMun3_Can4_F3_2018.checked := true else CMun3_Can4_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 4, 1) = '2' then
             CMun3_Can4_F4_2018.checked := true else CMun3_Can4_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can4_Family').AsString, 5, 1) = '2' then
             CMun3_Can4_F5_2018.checked := true else CMun3_Can4_F5_2018.checked := false;

          //자궁경부암
          if FieldByName('CMun3_Can5_Yn').AsString <> ''  then
             CMun3_Can5_Yn_2018.itemindex  := FieldByName('CMun3_Can5_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 1, 1) = '2' then
             CMun3_Can5_F1_2018.checked := true else CMun3_Can5_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 2, 1) = '2' then
             CMun3_Can5_F2_2018.checked := true else CMun3_Can5_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 3, 1) = '2' then
             CMun3_Can5_F3_2018.checked := true else CMun3_Can5_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 4, 1) = '2' then
             CMun3_Can5_F4_2018.checked := true else CMun3_Can5_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can5_Family').AsString, 5, 1) = '2' then
             CMun3_Can5_F5_2018.checked := true else CMun3_Can5_F5_2018.checked := false;

          //기타암
          CMun3_Can6_Bung_2018.Text := FieldByName('CMun3_Can6_Bung').AsString;
          if FieldByName('CMun3_Can6_Yn').AsString <> ''  then
             CMun3_Can6_Yn_2018.itemindex  := FieldByName('CMun3_Can6_Yn').AsInteger;
          //본인
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 1, 1) = '2' then
             CMun3_Can6_F1_2018.checked := true else CMun3_Can6_F1_2018.checked := false;
          //부모
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 2, 1) = '2' then
             CMun3_Can6_F2_2018.checked := true else CMun3_Can6_F2_2018.checked := false;
          //형제
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 3, 1) = '2' then
             CMun3_Can6_F3_2018.checked := true else CMun3_Can6_F3_2018.checked := false;
          //자매
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 4, 1) = '2' then
             CMun3_Can6_F4_2018.checked := true else CMun3_Can6_F4_2018.checked := false;
          //자녀
          if Copy(FieldByName('CMun3_Can6_Family').AsString, 5, 1) = '2' then
             CMun3_Can6_F5_2018.checked := true else CMun3_Can6_F5_2018.checked := false;

          // 4번문항
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

          // 5번문항
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

          // 6번문항
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

          // 7번문항
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

          // 여성수검자만 해당

          // 8번문항
          if FieldByName('CMun8').AsString <> ''  then
             CMun8_2018.itemindex  := FieldByName('CMun8').AsInteger;
          CMun8_Year_2018.Text := FieldByName('CMun8_Year').AsString;

          // 9번문항
          if FieldByName('CMun9').AsString <> ''  then
             CMun9_2018.itemindex  := FieldByName('CMun9').AsInteger;
          CMun9_Year_2018.Text := FieldByName('CMun9_Year').AsString;

          // 10번문항
          if FieldByName('CMun10').AsString <> ''  then
             CMun10_2018.itemindex  := FieldByName('CMun10').AsInteger;

          // 11번문항
          if FieldByName('CMun11').AsString <> ''  then
             CMun11_2018.itemindex  := FieldByName('CMun11').AsInteger;

          // 12번문항
          if FieldByName('CMun12').AsString <> ''  then
             CMun12_2018.itemindex  := FieldByName('CMun12').AsInteger;

          // 13번문항
          if FieldByName('CMun13').AsString <> ''  then
             CMun13_2018.itemindex  := FieldByName('CMun13').AsInteger;

          // 14번문항
          if FieldByName('CMun14').AsString <> ''  then
             CMun14_2018.itemindex  := FieldByName('CMun14').AsInteger;

          //정보활용 동의여부
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
        ParamByName('num_bunju').AsString  := copy(cmb_gmgn.text,26,4); //박대용 - 2016.03.08 - 24 -> 26으로 수정
        //ParamByName('num_bunju').AsString  := copy(cmb_gmgn.text,24,4);
        Active := True;

     if RecordCount > 0 then
     begin
        while not Eof do
        begin
           //혈액항목 판정
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

            //BMI 계산
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
       
       if (pos(FieldByname('desc_ear').AsString, '정상') > 0) or
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
       
       if pos(FieldByName('DESC_PEKI').AsString, '정상') > 0 then JPan8.Caption := 'A'
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
    // 임상참고치
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
                      // 1. 혈액에 대한 검사항목 추출
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

                    // 2. 종양에 대한 검사항목 추출
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

                    // 3. 장비에 대한 검사항목 추출
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


                    // 3. 추가코드에 대한 검사항목 추출
                      iRet := GF_MulToSingle(qryHul_gumgin2.FieldByName('cod_chuga').AsString, 4, UV_vCod_chuga2); /////////////////////수정시작
                      VarArrayRedim(UV_vCod_hm, index+iRet);
                      for i := 0 to iRet-1 do
                      begin
                         UV_vCod_hm[index] := UV_vCod_chuga2[i];
                         Inc(index);
                      end;
               
                     // 노신, 성인병, 간염에 대한 검사항목 추출
                      cod_name := '';
                      // 노신유형Display
                      if qryHul_gumgin2.FieldByName('gubn_nosin').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '1', qryHul_gumgin2.FieldByName('gubn_nsyh').AsInteger);
                      //성인병유형Display
                      if qryHul_gumgin2.FieldByName('gubn_adult').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '4', qryHul_gumgin2.FieldByName('gubn_adyh').AsInteger);
                      //간염유형Display
                      if qryHul_gumgin2.FieldByName('gubn_agab').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '5', qryHul_gumgin2.FieldByName('gubn_agyh').AsInteger);
                      //생애전환기유형Display
                      if qryHul_gumgin2.FieldByName('gubn_life').AsString = '1' then
                         cod_name := cod_name + UF_No_Hangmok(copy(qryHul_gumgin2.FieldByName('dat_gmgn').AsString,1,4), '7', qryHul_gumgin2.FieldByName('gubn_lfyh').AsInteger);

                    //특검유형Display
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



         // 주민번호를 이용하여 성별과 나이를 구함
         GP_JuminToAge(edtJumin.text,copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2), sMan, iAge);

         // 성별과 나이를 기준으로 Column명을 추출
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

         // 초기값 설정
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
            // 초기값 설정
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

            // 검사항목에 대한 임상참고치를 획득
            qryHm.Filter := '';
            qryHm.Filter := 'COD_HM = ''' + UV_vCod_hm[i] + ''' AND ' +
                            'DAT_APPLY <= ''' + copy(cmb_gmgn.text,1,4)+copy(cmb_gmgn.text,6,2)+copy(cmb_gmgn.text,9,2) + '''';

            // 혈액검사인지 Check(장비PF, 추가검사 때문)
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

            // 해당 검사에 대한 결과를 추출(Start Pos, End Pos를 이용)
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

            // 검사항목이 숫자+문자인 경우 검사항목(특수)에서 임상참고치 획득
            if qryHm.FieldByName('GUBN_HM').AsString = '3' then
            begin
               qryHm_sp.Active := False;
               qryHm_sp.ParamByName('COD_HM').AsString    := UV_vCod_hm[i];
               // 임상참고치를 변경하는 시점이 검사시점이므로 오늘일자(GV_sToday)를 전달
               qryHm_sp.ParamByName('DAT_APPLY').AsString := GV_sToday;
               qryHm_sp.Active := True;

               if qryHm_sp.RecordCount > 0 then
                  eIms := qryHm_sp.FieldByName('IMS_JUNG').AsFloat;

               // disconnect
               qryHm_sp.Active := False;
            end;

            // 해당 검사결과가 임상참고치를 벗어나는지 Check
            if qryHm.FieldByName('GUBN_HM').AsString = '1' then
            begin
               if Trim(sResult) = '' then
                  bCheck := True;

               // 숫자인 경우 Low와 High를 비교
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

               // 문자인 경우 문자임상치와 비교
               if GF_SpaceDel(sTemp) <> GF_SpaceDel(sMunja) then
                  bCheck := True;

               if Trim(sTemp) = '' then 
                  bCheck := True;
            end;

            

           
            //임상참고치가 벗어나면서 소견이 없을 때  판정처리..
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
            //파트별 소견이 없을 때  정상판정처리..---------------------------------------------
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
                        //========================= 장비항목판정
                        //안저
                        if FieldByName('cod_jangbi').AsString = 'JJ03' then
                           JPan2.Caption := GF_PanComparison(JPan2.Caption, FieldByName('gubn_pan').AsString)
                        //심전도
                        else if FieldByName('cod_jangbi').AsString = 'JJ09' then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //유방촬영
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
                        //갑상선초음파
                        else if FieldByName('cod_jangbi').AsString = 'JJ38' then
                           JPan4.Caption := GF_PanComparison(JPan4.Caption, FieldByName('gubn_pan').AsString)
                        //유방초음파
                        else if FieldByName('cod_jangbi').AsString = 'JJ39' then
                           JPan12.Caption := GF_PanComparison(JPan12.Caption, FieldByName('gubn_pan').AsString)
                        //자궁초음파
                        else if FieldByName('cod_jangbi').AsString = 'JJ41' then
                           JPan13.Caption := GF_PanComparison(JPan13.Caption, FieldByName('gubn_pan').AsString)
                        //심장초음파
                        else if FieldByName('cod_jangbi').AsString = 'JJ56' then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //심장CT
                        else if (FieldByName('cod_jangbi').AsString = 'JJHE') OR (FieldByName('cod_jangbi').AsString = 'JJHF') then
                           JPan7.Caption := GF_PanComparison(JPan7.Caption, FieldByName('gubn_pan').AsString)
                        //동맥경화지수
                        else if (FieldByName('cod_jangbi').AsString = 'JJE5') or
                                (FieldByName('cod_jangbi').AsString = 'JJDM') or
                                (FieldByName('cod_jangbi').AsString = 'JJBC') then
                           JPan5.Caption := GF_PanComparison(JPan5.Caption, FieldByName('gubn_pan').AsString)
                        //경동맥초음파
                        else if FieldByName('cod_jangbi').AsString = 'JJB8' then
                           JPan5.Caption := GF_PanComparison(JPan5.Caption, FieldByName('gubn_pan').AsString)
                        //X-ray
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ12') or (FieldByName('cod_jangbi').AsString = 'JJB7')) then
                           JPan8.Caption := GF_PanComparison(JPan8.Caption, FieldByName('gubn_pan').AsString)
                        //[대장]내시경&조영촬영
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ32')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ83')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ86')) and
                                 (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan10.Caption := GF_PanComparison(JPan10.Caption, FieldByName('gubn_pan').AsString)
                        //[위]내시경&조영촬영
                        else if ((FieldByName('cod_jangbi').AsString = 'JJ13')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJ34')  or
                                 (FieldByName('cod_jangbi').AsString = 'JJB9')) and
                                 (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan9.Caption := GF_PanComparison(JPan9.Caption, FieldByName('gubn_pan').AsString)
                        //[위]조직검사

                        else if ((FieldByName('cod_jangbi').AsString = 'JJ35')  or
                        (FieldByName('cod_jangbi').AsString = 'JJ78')) and  //20150120 곽윤설 - JJ78코드는 위-조직검사에 출력되도록 (본원-최미숙팀장님)
                        (FieldByName('cod_sokun').AsString <> '****')  then
                           JPan9.Caption := GF_PanComparison(JPan9.Caption, FieldByName('gubn_pan').AsString)
                        //[대장]조직검사
                        else if (FieldByName('cod_jangbi').AsString = 'JJDA') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                           JPan10.Caption := GF_PanComparison(JPan10.Caption, FieldByName('gubn_pan').AsString)
                        //골밀도검사
                        else  if ((FieldByName('cod_jangbi').AsString = 'JJ16') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH0') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH1') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH2') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH3') or
                                  (FieldByName('cod_jangbi').AsString = 'JJH4'))and
                                 (FieldByName('cod_sokun').AsString <> '****') then
                           JPan15.Caption := GF_PanComparison(JPan15.Caption, FieldByName('gubn_pan').AsString)
                        //CT검사
                        else if (FieldByName('pos_glkwa').AsString = 'CT') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           //BRAIN(뇌) CT
                           if FieldByName('cod_jangbi').AsString = 'JJ50' then
                           begin
                              JPan1.Caption := GF_PanComparison(JPan1.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //LUNG(폐) CT
                           else if FieldByName('cod_jangbi').AsString = 'JJ54' then
                           begin
                              JPan8.Caption := GF_PanComparison(JPan8.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //복부 CT
                           else if FieldByName('cod_jangbi').AsString = 'JJ90' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //복부비만CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //심장CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //복부비만CT
                           else if FieldByName('cod_jangbi').AsString = 'JJHB' then
                           begin
                              JPan11.Caption := GF_PanComparison(JPan11.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //기타 CT
                           else
                           begin
                              JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end
                        //MRI검사
                        else if ((FieldByName('pos_glkwa').AsString = 'MRI') or
                                 (FieldByName('pos_glkwa').AsString = 'MRA')) and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           if (FieldByName('cod_jangbi').AsString = 'JJC8') or (FieldByName('cod_jangbi').AsString = 'JJCN') then
                           begin
                              JPan1.Caption := GF_PanComparison(JPan1.Caption, FieldByName('gubn_pan').AsString);
                           end
                           //기타 MRI
                           else
                           begin
                              JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end
                        //PET검사
                        else if (FieldByName('pos_glkwa').AsString = 'PET') and
                                (FieldByName('cod_sokun').AsString <> '****') then
                        begin
                           JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                        end
                        //기타장비소견
                        else
                        begin
                           if (FieldByName('cod_sokun').AsString <> '****') then
                           begin
                              //10.09.04 승철. JJD2는 기타판정에 반영이 안되게끔.체질량지수(BMI)판정이 있어서 안나오는게 낳을꺼 같다(신현숙 부장님)
                              if (FieldByName('cod_jangbi').AsString <> 'JJD2') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJE0') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJE7') and
                                 (FieldByName('cod_jangbi').AsString <> 'JJEE') then
                                 JPan16.Caption := GF_PanComparison(JPan16.Caption, FieldByName('gubn_pan').AsString);
                           end;
                        end;
                        //=========================  장비판정

                        Next;
                   end;

                   // Table과 Disconnect
                   Active := False;
              end;
         end;


end;





procedure TfrmCK5I17.grd_sokunCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // 자룔를 화면에 조회
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
   // 자료가 존재하는지 Check
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  //서버 연결 중
  if      copy(GV_sUserId,1,2) = '15' then ctiSvrIp  := '125.128.11.121'  //광화문
  else if copy(GV_sUserId,1,2) = '11' then ctiSvrIp  := '211.192.180.181' //강남
  else if copy(GV_sUserId,1,2) = '12' then ctiSvrIp  := '121.131.48.173'  //여의도
  else if copy(GV_sUserId,1,2) = '61' then ctiSvrIp  := '119.198.186.80'  //부산
  else if copy(GV_sUserId,1,2) = '43' then ctiSvrIp  := '59.11.26.58'     //수원
  else if copy(GV_sUserId,1,2) = '51' then ctiSvrIp  := '220.93.183.118'  //광주
  else
  begin
     ShowMessage('CTI를 사용할 수 없는 센터입니다.');
     exit;
  end;

  if Pnl_CTISvr.Caption = '연결' then
  begin
     if Panel_CTI.Caption = '로그아웃' then UP_CTI_Login('IN')   //로그인...
     else                                   UP_CTI_Login('OUT'); //로그아웃...
  end
  else
  begin
     //CTI 서버 연결...
     CAModule.ConnectCTIServer(ctiSvrIp);
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_AnswerClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  //전화받기
  if CAModule.GetLoginState then CAModule.Answer();
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_DisconnectClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  //전화끊기
  if CAModule.GetLoginState then CAModule.Hangup();
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_BreakClick(Sender: TObject);
var
   sTemp : string;
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then
  begin
    sTemp := CAModule.GetTellerStatus();
    huTime := 0;

    if (sTemp = 'L001') or (sTemp = 'R001') then
    begin
      ShowMessage('휴식중이나 식사중에는 후처리를 할 수 없습니다.');
    end
    else if sTemp <> 'W004' then begin CAModule.ChangeTellerStatus('W004'); Timer_hu.Enabled := True;  end
    else                         begin CAModule.ChangeTellerStatus('0300'); Timer_hu.Enabled := False; end;
  end;
  //============================================================================
end;

procedure TfrmCK5I17.SBtn_CTI_EtcClick(Sender: TObject);
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then
  begin
     //상담원 상담모드 변경...
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
  //1 - 파일첨부
  if cmbRelation.ItemIndex = 1 then
  begin
     if edtJumin.Text <> '' then
     begin
        frmCK5I171  := TFrmCK5I171.Create(self);
        frmCK5I171.ShowModal;
     end
     else
        showmessage('조회된 고객정보가 없습니다.');
  end
  //2 - 제휴병원 목록
  else if cmbRelation.ItemIndex = 2 then
  begin
     frmCK5I161  := TFrmCK5I161.Create(self);
     frmCK5I161.ShowModal;
  end
  //3 - 제휴병원 소견코드 목록
  else if cmbRelation.ItemIndex = 3 then
  begin
     frmCK5I162  := TFrmCK5I162.Create(self);
     frmCK5I162.ShowModal;
  end
  //4 - SMS전송 문자관리
  else if cmbRelation.ItemIndex = 4 then
  begin
     frmCK5I17B := TFrmCK5I17B.Create(self);
     frmCK5I17B.Show;
  end
  //5 - 검진자 해피콜
  else if cmbRelation.ItemIndex = 5 then
  begin
     frmCK5I163  := TFrmCK5I163.Create(self);
     frmCK5I163.ShowModal;
  end
  //7 - 과거상담내용 추가
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
              MessageDlg('PDF파일이 존재하지 않습니다.. '+ #10#13 + '('+sFull_FileName+')', mtError, [mbOK], 0);
              exit;
           end
           else
           begin
              AppHandle := ShellExecute(Handle,'open',PChar(sFull_FileName),'','',SW_SHOWNORMAL);
           end;
        end
        else
        begin
           GF_MsgBox('Information', '조회권한이 없습니다..');
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
   // 자료가 존재하는지 Check
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
   // 자료가 존재하는지 Check
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
   if sTemp = '1' then Result := '본원센터';
   if sTemp = '2' then Result := '부산센터';
   if sTemp = '3' then Result := '강남센터';
   if sTemp = '4' then Result := '수원센터';
   if sTemp = '5' then Result := '대구센터';
   if sTemp = '6' then Result := '광주센터';
   if sTemp = '7' then Result := '여의도센터';
   if sTemp = '8' then Result := '1599-7070';

   if sTemp = 'a' then Result := '본원센터->위치안내';
   if sTemp = 'b' then Result := '본원센터->예약접수';
   if sTemp = 'c' then Result := '본원센터->결과상담';
   if sTemp = 'd' then Result := '본원센터->행정부서';
   if sTemp = 'f' then Result := '본원센터->기타연결';

   if sTemp = 'g' then Result := '부산센터->위치안내';
   if sTemp = 'h' then Result := '부산센터->예약접수';
   if sTemp = 'i' then Result := '부산센터->검진문의';
   if sTemp = 'j' then Result := '부산센터->결과상담';
   if sTemp = 'k' then Result := '부산센터->특수검진';
   if sTemp = 'l' then Result := '부산센터->기타연결';

   if sTemp = 'm' then Result := '강남센터->위치안내';
   if sTemp = 'n' then Result := '강남센터->예약접수';
   if sTemp = 'o' then Result := '강남센터->보험계약';
   if sTemp = 'p' then Result := '강남센터->결과상담';
   if sTemp = 'q' then Result := '강남센터->외래센터';
   if sTemp = 'r' then Result := '강남센터->기타연결';

   if sTemp = 's' then Result := '수원센터->위치안내';
   if sTemp = 't' then Result := '수원센터->예약접수';
   if sTemp = 'u' then Result := '수원센터->결과상담';
   if sTemp = 'v' then Result := '수원센터->기타상담';

   if sTemp = 'w' then Result := '대구센터->위치안내';
   if sTemp = 'x' then Result := '대구센터->예약접수';

   if sTemp = 'y' then Result := '광주센터->위치안내';
   if sTemp = 'z' then Result := '광주센터->예약접수';
   if sTemp = '!' then Result := '광주센터->공단검진';
   if sTemp = '@' then Result := '광주센터->결과상담';
   if sTemp = '#' then Result := '광주센터->환경측정';
   if sTemp = '$' then Result := '광주센터->기타문의';

   if sTemp = '%' then Result := '여의도센터->위치안내';
   if sTemp = '^' then Result := '여의도센터->예약접수';
   if sTemp = '&' then Result := '여의도센터->결과상담';
   if sTemp = '*' then Result := '여의도센터->기타업무';
end;

procedure TfrmCK5I17.cmbCOD_JISAChange(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
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
         // DataSet의 값을 Variant변수로 이동
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

   // Grid에 자료를 할당
   grdMaster.Rows := index;
end;

procedure TfrmCK5I17.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // 자룔를 화면에 조회
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
   // 자료가 존재하는지 Check
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
      GF_MsgBox('Warning', '회신번호를 확인해주세요!!');
      exit;
  end;

  if Num_tel.Text = '' then
  begin
      GF_MsgBox('Warning', '수신번호를 확인해주세요!!');
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
                 ParamByName('SUBJECT').AsString   := '한국의학연구소';
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
      GF_MsgBox('Warning', '메세지를 입력해주세요!!');
      exit;
  end;;

  GF_MsgBox('Information', '전송완료!');
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
     GF_MsgBox('Warning', '멜티메세지로 전송됩니다. 문자메세지:13원 멀티메세지:50원!!');
     UV_CkByte := True;
  end;
end;

procedure TfrmCK5I17.FormShow(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrder : String;
begin
  inherited;
   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
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
         // DataSet의 값을 Variant변수로 이동
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

   // Grid에 자료를 할당
   grdMaster.Rows := index;

end;

procedure TfrmCK5I17.SBtn_CallClick(Sender: TObject);
var
  i : Integer;
  sCall, sLost, sTel : string;
begin
  inherited;
  //[2016.08.05-유동구]신규 CTI 작업...
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

     CAModule.MakeCall('',sCall,'');  //발신자번호/수신전화번호/사용자데이터
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
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  if CAModule.GetLoginState then   
  begin
    inc(huTime);
    Pnl_CTIUser.Caption := '상담후처리(' + IntToStr(huTime) + '초)';
    if huTime >= 30 then
    begin
      //상담원 상태 변경...
      CAModule.ChangeTellerStatus('0300');
      Timer_hu.Enabled := False;
    end
  end;
  //============================================================================
end;

procedure TfrmCK5I17.FormDestroy(Sender: TObject);
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  //상담원 로그아웃
  if CAModule.GetLoginState then UP_CTI_Login('OUT');
  //============================================================================
end;

procedure TfrmCK5I17.Timer1Timer(Sender: TObject);
begin
  inherited;
  //[2016.07.07-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  //CTI Server와의 연결상태
  if CAModule.GetConnectionState then Pnl_CTISvr.Caption := '연결'
  else                                Pnl_CTISvr.Caption := '끊김';

  //상담원 로그인 상태
  if CAModule.GetLoginState then
  begin
     Panel_CTI.Caption := '로그인';
     SBtn_CTI_Answer.Enabled     := True;
     SBtn_CTI_Disconnect.Enabled := True;
     SBtn_CTI_Break.Enabled      := True;
     SBtn_CTI_Hu.Enabled         := True;
     SBtn_CTI_Etc.Enabled        := True;
     SBtn_CTI_Monitor.Enabled    := True;
     SBtn_CTI_Callback.Enabled   := True;
     SBtn_CTI_Setting.Enabled    := True;

     CAModule.GetTellerStatus();
     pnl_WaitCnt.Caption := CAModule.GetWaitingCnt + '명';
  end
  else
  begin
     Panel_CTI.Caption := '로그아웃';
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
               if (POS('JJ1I',FieldByName('cod_chuga').AsString) > 0) then sMsg := '[재검]';

               showmessage('신신검진(검진일 : ' + GF_DateFormat(FieldByName('dat_gmgn').AsString) + ')' + sMsg);
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
   // 문자+숫자 형태의 문자열에서 문자만 분리

   // 자료가 없을 경우의 처리
   if sData = '' then
   begin
      sValue := '';
      exit;
   end;

   // 문자열에서 Space제거
   sData := GF_SpaceDel(sData);
   sValue := '';

   for i := 1 to Length(sData) do
   begin
      // 자료가 숫자이면 작업완료(이전자료까지만 문자로 인식)
      if GF_IsNumber(sData[i]) then break;

      sValue := sValue + sData[i];
   end;
end;

//2016.02.03 - 박대용 - 접수관리(CK3I01)의 컴플레인버튼과 연결
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  UP_CTI_Login('IN');
  //============================================================================
end;

procedure TfrmCK5I17.CAModuleCTIConnectError(Sender: TObject);
begin
  inherited;
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  ShowMessage('CTI Server 연결에 실패하였습니다.');
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
  //[2016.08.05-유동구]신규CTI작업...
  //----------------------------------------------------------------------------
  sMsg_Tmp := strMsg;
  for iCnt := 0 to 200 do begin aMsgList[iCnt] := ''; end;

  //ShowMessage(strMsg);

  //CMD CODE(strMsg : 각필드마다 구분자 사용 "^" / 시간은 로컬시간 HH:MM:SS 형식)
  sCMD := copy(sMsg_Tmp,1,2);

  //각 필드에 대한 값 등록...
  iCnt := 0;
  while pos('^',sMsg_Tmp) > 0 do
  begin
     sMsg_Cut := copy(sMsg_Tmp,1,pos('^',sMsg_Tmp)-1);
     sMsg_Tmp := copy(sMsg_Tmp,pos('^',sMsg_Tmp)+1,length(sMsg_Tmp));

     aMsgList[iCnt] := sMsg_Cut;

     Inc(iCnt);
  end;
  //마지막 구분자 값 등록...
  if sMsg_Tmp <> '' then aMsgList[iCnt] := sMsg_Tmp;


  //로그인 응답(00)
  //aMsgList(0:CMD,1: 아이디,2:내선번호,3:이름,4:응답코드,5:최초상태,6:대그룹,7:중그룹,8:소그룹,9:대기자수,10:시간)
  if sCMD = '00' then //로그인 응답
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        if      aMsgList[4] = '2' then ShowMessage('아이디가 존재하지 않습니다.')
        else if aMsgList[4] = '3' then ShowMessage('비밀번호가 일치하지 않습니다.')
        else if aMsgList[4] = '4' then ShowMessage('로그인이 중복되었습니다.')
        else if aMsgList[4] = '5' then ShowMessage('내선번호 중복되었습니다.')
        else if aMsgList[4] = '6' then ShowMessage('미등록 내선입니다.')
        else if aMsgList[4] = '7' then ShowMessage('상담원 기본그룹 오류입니다.');   
     end;
  end
  else if sCMD = '02' then //비밀번호변경 응답
  begin
     if      aMsgList[4] = '1' then ShowMessage('성공')
     else if aMsgList[4] = '2' then ShowMessage('잘못된 아이디입니다.')
     else if aMsgList[4] = '3' then ShowMessage('잘못된 비밀번호입니다.');
  end
  else if sCMD = '85' then //상담원 모드 응답
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        if      aMsgList[2] = '0' then begin Pnl_Mod.Color := $00FDD0B5; Pnl_Mod.Caption := '인바운드';   end
        else if aMsgList[2] = '1' then begin Pnl_Mod.Color := $00D6E9F3; Pnl_Mod.Caption := '아웃바운드'; end;
     end;
  end
  else if sCMD = '86' then //스크린 팝업
  begin
     if      aMsgList[2] = '1' then sTemp := 'Ring'
     else if aMsgList[2] = '2' then sTemp := 'Answer';

     if      aMsgList[1] = '01' then Pnl_PopupType.Caption := sTemp + '[호분배]'
     else if aMsgList[1] = '02' then Pnl_PopupType.Caption := sTemp + '[인바운드]'
     else if aMsgList[1] = '03' then Pnl_PopupType.Caption := sTemp + '[인바운드(돌려주기)]'
     else if aMsgList[1] = '04' then Pnl_PopupType.Caption := sTemp + '[아웃바운드]'
     else if aMsgList[1] = '05' then Pnl_PopupType.Caption := sTemp + '[오토콜]'
     else if aMsgList[1] = '06' then Pnl_PopupType.Caption := sTemp + '[당겨받기]'
     else if aMsgList[1] = '07' then Pnl_PopupType.Caption := sTemp + '[돌려받기]'
     else if aMsgList[1] = '08' then Pnl_PopupType.Caption := sTemp + '[3자통화]';

     if      aMsgList[7] = 'S511' then Panel_CTI_Connec.Caption := '예약 및 변경(본원)'
     else if aMsgList[7] = '2100' then Panel_CTI_Connec.Caption := '예약 및 변경(강남)'
     else if aMsgList[7] = '2200' then Panel_CTI_Connec.Caption := '예약 및 변경(여의도)'
     else if aMsgList[7] = 'S513' then Panel_CTI_Connec.Caption := '보험 관련 상담'
     else if aMsgList[7] = 'S512' then Panel_CTI_Connec.Caption := '결과 상담'
     else if aMsgList[7] = 'S514' then Panel_CTI_Connec.Caption := '기타 문의';


     if (sTemp = 'Ring') and (Length(Trim(UV_vCallNo)) > 4) then Pnl_tel.Text := UV_vCallNo;

     UV_sCustNo := aMsgList[5];
     UV_vCallNo := aMsgList[4];

     Panel_CTI_Custom.Text := UV_vCallNo;
     Timer_hu.Enabled         := False;

     if (Trim(UV_vCallNo) <> '') and (sTemp = 'Ring') then
     begin
        if      aMsgList[1] = '04'           then Panel_CTI_Connec.Caption := '발신전화'
        else if Length(Trim(UV_vCallNo)) < 5 then Panel_CTI_Connec.Caption := '내부전화';
     end;

{
       ShowMessage('[스크린팝업정보]\n팝업타입:'             + aMsgList[1]  + '\n팝업시점:'           + aMsgList[2]  + '\nA-1 대표번호:'    + aMsgList[3]  +
            	                   '\nA-2 발신자번호:'       + aMsgList[4]  + '\nA-3 IVR 연동데이터:' + aMsgList[5]  + '\nA-4 CALL_ID:'     + aMsgList[6]  +
    			     	   '\nA-5 IVR메뉴번호:'      + aMsgList[7]  + '\nA-6 사용자데이터:'   + aMsgList[8]  + '\nA-7 녹취파일명:'  + aMsgList[9]  +
    				   '\nA-8 녹취폴더(녹취일):' + aMsgList[10] + '\nB-1 대표번호:'       + aMsgList[11] + '\nB-2 발신자번호:'  + aMsgList[12] +
    				   '\nB-3 IVR 연동데이터:'   + aMsgList[13] + '\nB-4 콜아이디:'       + aMsgList[14] + '\nB-5 IVR메뉴번호:' + aMsgList[15] +
    				   '\nB-6 사용자데이터:'     + aMsgList[16]);
}
  end
  else if sCMD = '89' then //상담원이 변경할 수 있는 상태정보 요구에 대한 응답
  begin
    iCnt := 1;
    Cmb_CTI_status.Items.Clear;
    while iCnt < 20 do
    begin
       if length(aMsgList[iCnt]) = 4 then Cmb_CTI_status.Items.Add(aMsgList[iCnt] + '-' + aMsgList[iCnt+1]);
       iCnt := iCnt + 2;
    end;
  end
  else if sCMD = '90' then //CTI 상태정보 응답(상담원의 상태 표출에 사용되는 정보)
  begin
    iCnt := 1;
    Cmb_CTI_status_all.Items.Clear;
    while iCnt < 200 do
    begin
       if length(aMsgList[iCnt]) = 4 then Cmb_CTI_status_all.Items.Add(aMsgList[iCnt] + '-' + aMsgList[iCnt+1]);
       iCnt := iCnt + 2;
    end;
  end
  else if sCMD = '93' then //모든 상담원 상태 요구에 대한 응답
  begin
     //모든 상담원 상태 요구에 대한 응답
  end
  else if sCMD = '94' then //상담원 상태 변경 알림
  begin
     if aMsgList[1] = Panel_CTINo.Caption then
     begin
        sTemp := Trim(aMsgList[5]);
        if sTemp <> '' then
        begin
           //상담후처리
           if sTemp = 'W004' then begin huTime := 0; Timer_hu.Enabled := True; end;
           // if sTemp = '0200' then sTemp := '3000'; //출근코드 일때 전화대기로 변경

           GP_ComboDisplay(Cmb_CTI_status_all, sTemp, 4);

           if sTemp <> '0300' then Pnl_CTIUser.Color := $00D6E9F3
           else                    Pnl_CTIUser.Color := $00FDD0B5;

           Pnl_CTIUser.Caption := copy(Cmb_CTI_status_all.Text,6,length(Cmb_CTI_status_all.Text));
        end;

        if      aMsgList[4] = '0' then begin Pnl_Mod.Color := $00FDD0B5; Pnl_Mod.Caption := '인바운드';   end
        else if aMsgList[4] = '1' then begin Pnl_Mod.Color := $00D6E9F3; Pnl_Mod.Caption := '아웃바운드'; end;
     end;
  end
  else if sCMD = '95' then //고객 대기자수 알림
  begin
     pnl_WaitCnt.Caption := aMsgList[1] + '명';
  end;

  //============================================================================
end;

procedure TfrmCK5I17.SBtn_Call2Click(Sender: TObject);
var
   i : Integer;
   sCall, sLost, sTel : string;
begin
   inherited;
   //[2016.08.05-유동구]신규 CTI 작업...
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

      //[2016.07.07-유동구]신규 CTI 작업...
      CAModule.MakeCall('',sCall,'');  //발신자번호/수신전화번호/사용자데이터
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
   else                      showMessage('실행파일이 존재하지 않습니다.');
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
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   //소견순서
   EdtTNum_seq.Text     := UV_vTNum_seq[NewRow-1];

   //유해물질
   EdtTCod_HMatter.Text := UV_vTCod_matter[NewRow-1];
   EdtDesc_HMatter.Text := UV_vTDesc_matter[NewRow-1];

   //좌/우
   edtTGubn_LR.Text     := UV_vTGubn_LR[NewRow-1];

   //판정
   EdtTCod_pan.Text     := UV_vTCod_pan[NewRow-1];

   //소견
   EdtTcod_sokun.Text  := UV_vTcod_sokun[NewRow-1];
   EdtTDesc_sokun.Text := UF_CodeEdit(1, EdtTcod_sokun.Text);

   //조치
   EdtTCod_jochi.Text  := UV_vTcod_jochi[NewRow-1];
   EdtTDesc_jochi.Text := UF_CodeEdit(2, EdtTCod_jochi.Text);

   //표적장기
   EdtTCod_janggi.Text  := UV_vTCod_janggi[NewRow-1];
   EdtTDesc_janggi.Text := UF_CodeEdit(3, EdtTCod_janggi.Text);

   //질병코드
   EdtTTot_Jilhan.Text  := UV_vTTot_jilhan[NewRow-1];

   //질환
   sName := '';
   EdtTCod_jilhan.Text  := UV_vTCod_jilhwan[NewRow-1];
   if GF_SickEdit(EdtTCod_jilhan, sName) then EdtTDesc_jilhan.Text := sName;

   //업무적합성 평가
   sName := '';
   EdtTCod_ujs.Text     := UV_vTcod_ujs[NewRow-1];
   if GF_ujsEdit(EdtTCod_ujs, sName) then EdtTDesc_ujs.Text  := sName;
   if EdtTDesc_ujs.Text = '004'      then EdtTDesc_ujs.Color := clRed
   else                                   EdtTDesc_ujs.Color := clSilver;

   //사후관리(1)
   sName := '';
   EdtTCod_sahu1.Text   := UV_vTCod_sahu1[NewRow-1];
   if GF_sahuEdit(EdtTCod_sahu1, sName) then EdtTDesc_sahu1.Text  := sName;
   if EdtTDesc_sahu1.Text = '06'        then EdtTDesc_sahu1.Color := clRed
   else                                      EdtTDesc_sahu1.Color := clSilver;

   //사후관리(2)
   sName := '';
   EdtTCod_sahu2.Text   := UV_vTCod_sahu2[NewRow-1];
   if GF_sahuEdit(EdtTCod_sahu2, sName) then EdtTDesc_sahu2.Text  := sName;
   if EdtTDesc_sahu2.Text = '06'        then EdtTDesc_sahu2.Color := clRed
   else                                      EdtTDesc_sahu2.Color := clSilver;

   //세부소견
   MM_Tetcsokun.Text := UV_vTDesc_etcSokun[NewRow-1];
   //비고란(송부)
   MM_Tetcbigo.Text  := UV_vTDesc_bigo[NewRow-1];
   //세부조치
   MM_TetcJochi.Text := UV_vTDesc_etcJochi[NewRow-1];
   //비고란(KMI)
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
            // 수정된 Column을 표시
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
         //2014.10.16 김창규
         if GF_StrToFloat(UV_vGulkwa[DataRow-1]) = 0 Then
         begin
            // 수정된 Column을 표시
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end
      else
      begin
         // 수정된 Column을 표시
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
            // 수정된 Column을 표시
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end
      else if (UV_vCod_Hhm[DataRow-1] = '0515') or (UV_vCod_Hhm[DataRow-1] = '0517') then
      begin
         if (GF_StrToFloat(UV_vGulkwa[DataRow-1]) <= 3) then
         begin
            // 수정된 Column을 표시
            grdResult.AssignCellFont(DataCol, DataRow);
            grdResult.CellFont[DataCol, DataRow].Color := clRed;
         end;
      end;
   end;
   if UV_vCod_Mhm[DataRow-1] = 'C027' then
   begin
      // 수정된 Column을 표시
      grdResult.AssignCellFont(DataCol, DataRow);
      grdResult.CellFont[DataCol, DataRow].Color := clBlue;
   end;
   if (UV_vCod_Mhm[DataRow-1] = 'C083') and (EdtGubn_Cha.Text = '1') then
   begin
      // 수정된 Column을 표시
      grdResult.AssignCellFont(DataCol, DataRow);
      grdResult.CellFont[DataCol, DataRow].Color := clBlue;
   end;
   if UV_vCod_Mhm[DataRow-1] = '----' then
   begin
      // 수정된 Column을 표시
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

             //위장관계 관련증상문
             If UV_vCod_Mhm2[DataRow-1] = 'JJUN' Then
             Begin
                If Pos('저위험군',Value) > 0 Then
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

             //[2016.12.26-유동구]내분비계 관련증상문
             If UV_vCod_Mhm2[DataRow-1] = 'JJUL' Then
             Begin
                If Pos('비정상',Value) > 0 Then
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
   // 작업중인지 Check & Cancel Confirm Message
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

               showmessage(D_File_name + ' 파일 삭제가 완료되었습니다.');
               grd_file.Refresh;
               Break;
            end;
         end;
         btnQuery.Click;
      end;
   end
   else ShowMessage('\\222.222.1.6\chart\ 네트워크 드라이브 연결이 되어있지 않습니다. ' + #13#10 +
                    '네트워크 드라이브 연결후 업로드 하시기 바랍니다.');
end;

end.
