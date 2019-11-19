//==============================================================================
// 프로그램 설명 : 혈액학결과등록
// 시스템        : 통합검진
// 수정일자      : 1999.10.13                                                                                                                       
// 수정자        : 김은석
// 수정내용      : 항목에 해당되는 코드 추가
// 참고사항      : 2006.09.08 [B100,A100,AB01,AB02 - 검사 중 몇가지 검사 체크]
//==============================================================================
// 수정일자      : 2006.09.08
// 수정자        : 유동구
// 수정내용      : 2006.09.08 [B100,A100,AB01,AB02 - 검사 중 몇가지 검사 체크]
// 참고사항      : A100 -- C014,C056, C057, C058,C008,S001.
//                 B100,AB01,AB02 -- C014,C056. 혈액형은 검사안하니 상관없어요.
//==============================================================================
// 수정일자      : 2006.11
// 수정자        : 유동구
// 수정내용      : 특수검진(혈액기타2) - 광주센터 관련하여 수정..
// 참고사항      : 기본적으로 광주센터에서 발생한 특검 항목은 광주분석부에서 입력함.
//                 본원 분주라도 분석부에서는 보이지 않고, 광주센터에서만 보이게 함.
//==============================================================================
// 수정일자      : 2009.04.15
// 수정자        : 유동구
// 수정내용      : 2009년도 노신항목 추가.. HDL,LDL,중성지방,크레아티닌
// 참고사항      : 4개항목 추가..
//==============================================================================
// 수정일자      : 2009.06.27
// 수정자        : 김승철
// 수정내용      : 09.06.27 철. 호모시스테임 검사있을시 표시
//==============================================================================
// 수정일자      : 2011.01.26
// 수정자        : 송재호
// 수정내용      : 검색구분에 혈액형검사자를 추가하여 혈액형검사자만 Display되도록 수정
//==============================================================================
// 수정일자      : 2011.04.22
// 수정자        : 송재호
// 수정내용      : 전결과 보이도록 수정(노신이상) 및 다음항목조회 가능하도록 수정 및 입력안된 text 생화학만 녹색으로 표시(한의 재단생화학팀00066)
//==============================================================================
// 수정일자      : 2011.05.31
// 수정자        : 송재호
// 수정내용      : 전결과 생화학결과등록에만 보이도록.. (속도저하로 인해)
//==============================================================================
// 수정일자      : 2011.06.13
// 수정자        : 김승철
// 수정내용      : 혈액형검사자인경우 혈액형으로 바로커서이동, 해당칸 대문자, 엔터엔터후엔 다음항목에서 저장으로 이동.
//==============================================================================
// 수정일자      : 2011.06.16
// 수정자        : 송재호
// 수정내용      : 전결과 혈액학결과등록에도 보이도록.. (주임 엄주호 요청)
//==============================================================================
// 수정일자      : 2011.12.28
// 수정자        : 유동구
// 수정내용      : 혈액학, 뇨화학 자체검진 관련 추가...
//==============================================================================
// 프로그램 설명 : 혈액형분주현황
// 시스템        : 통합검진
// 수정일자      : 2012.03.13
// 수정자        : 송재호
// 수정내용      : 혈액학결과등록시 본사에서 ESR분주현황 조회할땐 검진센터 선택 가능토록..
// 참고사항      : 한의 재단혈액학팀1200022
//==============================================================================
// 수정일자      : 2012.6.22
// 수정자        : 김희석
// 수정내용      : 보류 사유 추가
//=============================================================================
// 수정일자      : 2012.10.30
// 수정자        : 유동구
// 수정내용      : 대전협력병원 자체검진관련 추가(생화학 3일때 모두 검사할수있게...
//=============================================================================
// 수정일자      : 2012.11.14
// 수정자        : 김희석
// 수정내용      : 키, 몸무게,임신및 생리
//=============================================================================
// 수정일자      : 2013.06.24
// 수정자        : 김희석
// 수정내용      : 이전 결과 EIA,혈청학 추가
//=============================================================================
// 수정일자      : 2013.7.29
// 수정자        : 김희석
// 수정내용      :한의 재단진단면역학팀1300101 및 지사 25 종 추가
//=============================================================================
// 수정일자      : 2014.04.18
// 수정자        : 곽윤설
// 수정내용      : 혈액구분 3-'연구소+센터'인 C077항목 입력 비활성화
//=============================================================================
// 수정일자      : 2014.04.26
// 수정자        : 곽윤설
// 수정내용      : 1. Title명 -> STMS에 있는것과 동일하게 수정
//                 2. C083 -> 3.0미만 입력 못하도록
//=============================================================================
// 수정일자      : 2014.04.30
// 수정자        : 곽윤설
// 수정내용      : 생화학파트 일괄 자동계산  [본원-최은희 요청]
//참고사항       :
//=============================================================================
// 수정일자      : 2014.05.07
// 수정자        : 곽윤설
// 수정내용      : C083 -> 3.0미만은 '<3.0'으로 입력 가능
//참고사항       : [본원 - 최은희]
//=============================================================================
// 수정일자      : 2014.05.15
// 수정자        : 곽윤설
// 수정내용      : 값이 없을 경우 결과값 안찍히도록.. (0.0도 안되게..)
//참고사항       : [본원 - 최은희]
//=============================================================================
// 수정일자      : 2014.05.23
// 수정자        : 곽윤설
// 수정내용      : 'Cmb_gubn=혈액형검사지' 일때 혈액프로파일 모두 읽어도록..
//참고사항       : LD4Q02도 동일하게 수정
//=============================================================================
// 수정일자      : 2014.11.17
// 수정자        : 곽윤설
// 수정내용      : 채용검진 & 센터자체검진(노신9종)일 경우 - 연구소검사에서 센터 자체검사로
// 참고사항      :
//=============================================================================
// 수정일자      : 2015.04.08
// 수정자        : 곽윤설
// 수정내용      : 노신,성인병 재검 & C032(공복시혈당)은 센터 자체검사
// 참고사항      : 한의 수원진단검사의학팀1500248  // 전국센터 진단검사의학팀 요청
//=============================================================================
// 수정일자      : 2015.06.27
// 수정자        : 곽윤설
// 수정내용      : 특검 재검 & C032(공복시혈당)은 센터 자체검사
// 참고사항      : 부산 유희짐, 본원 박연숙 팀자 요청
//=============================================================================
// 수정일자      : 2015.08.19
// 수정자        : 곽윤설
// 수정내용      : 2015.06.27 요청 번복, 특검 재검 C032 센터자체검사 제외
// 참고사항      :
//=============================================================================
// 수정일자      : 2015.09.17
// 수정자        : 곽윤설
// 수정내용      : 생화학결과 출력 (LD1P011)
// 참고사항      : 한의 부산진단검사의학팀1500839
//=============================================================================
// 수정일자      : 2015.09.24
// 수정자        : 곽윤설
// 수정내용      : 저장 시 알림창 안뜨도록 수정
// 참고사항      : 본원 혈액학팀 요청
//=============================================================================
// 수정일자      : 2015.11.02
// 수정자        : 곽윤설
// 수정내용      : C011 r-GTP 에서 γ-GTP로 검사항목명 변경에 따른 프로그램 수정
// 참고사항      : 글로벌헬스케어팀장 이춘희 수석 요청
//=============================================================================
// 수정일자      : 2015.11.03
// 수정자        : 곽윤설
// 수정내용      : 노신재검&C032 검사 정리
//                 노신/성인병/생애 재검(2차)이면서 공복시 혈당(C032) 검사 있을 경우 C032는 센터자체검사
//                 노신재검과 특검재검이 같이 있을 경우 노신재검이 우선순위, C032 센터자체검사
//                 특검재검만 있을 경우에는 C032 연구소 검사로 진행함.
// 참고사항      : 문지혜 선임님 확인
//=============================================================================
// 수정일자      : 2015.11.09
// 수정자        : 곽윤설
// 수정내용      : 요화학 결과등록 　
//                 Y004검사자/니코틴검사자/요침사'검사자 추가 [본원만 적용]
//                 조회조건 중 혈액형검사자/ESR검사자 방향키로도 확인할수 있도록 적용
// 참고사항      : 한의 재단진단검사의학팀1500832
//=============================================================================
// 수정일자      : 2017.08.25
// 수정자        : 박수지
// 수정내용      :  c074 결과 있을 시에 c027에 동일결과 자동 저장
// 참고사항      :
//=============================================================================

unit LD1I01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I01 = class(TfrmSingle)
    qryGlkwa: TQuery;
    grdMaster: TtsGrid;
    qryU_Glkwa: TQuery;
    pnlDetail: TPanel;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlName1: TPanel;
    pnlCond: TPanel;
    rbDate: TRadioButton;
    rbJumin: TRadioButton;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    Panel47: TPanel;
    Panel17: TPanel;
    pnlNum1: TPanel;
    pnlNum2: TPanel;
    pnlName2: TPanel;
    edtValue2: TEdit;
    pnlNum3: TPanel;
    pnlName3: TPanel;
    edtValue3: TEdit;
    pnlNum4: TPanel;
    pnlName4: TPanel;
    edtValue4: TEdit;
    pnlNum5: TPanel;
    pnlName5: TPanel;
    edtValue5: TEdit;
    pnlNum6: TPanel;
    pnlName6: TPanel;
    edtValue6: TEdit;
    pnlNum7: TPanel;
    pnlName7: TPanel;
    edtValue7: TEdit;
    pnlNum8: TPanel;
    pnlName8: TPanel;
    edtValue8: TEdit;
    pnlNum9: TPanel;
    pnlName9: TPanel;
    edtValue9: TEdit;
    pnlNum10: TPanel;
    pnlName10: TPanel;
    edtValue10: TEdit;
    pnlNum11: TPanel;
    pnlName11: TPanel;
    edtValue11: TEdit;
    pnlNum12: TPanel;
    pnlName12: TPanel;
    edtValue12: TEdit;
    pnlNum13: TPanel;
    pnlName13: TPanel;
    edtValue13: TEdit;
    pnlNum14: TPanel;
    pnlName14: TPanel;
    edtValue14: TEdit;
    pnlNum15: TPanel;
    pnlName15: TPanel;
    edtValue15: TEdit;
    Panel33: TPanel;
    Panel34: TPanel;
    pnlName16: TPanel;
    pnlNum16: TPanel;
    pnlNum17: TPanel;
    pnlName17: TPanel;
    edtValue17: TEdit;
    pnlNum18: TPanel;
    pnlName18: TPanel;
    edtValue18: TEdit;
    pnlNum19: TPanel;
    pnlName19: TPanel;
    edtValue19: TEdit;
    pnlNum20: TPanel;
    pnlName20: TPanel;
    edtValue20: TEdit;
    pnlNum21: TPanel;
    pnlName21: TPanel;
    edtValue21: TEdit;
    pnlNum22: TPanel;
    pnlName22: TPanel;
    edtValue22: TEdit;
    pnlNum23: TPanel;
    pnlName23: TPanel;
    edtValue23: TEdit;
    pnlNum24: TPanel;
    pnlName24: TPanel;
    edtValue24: TEdit;
    pnlNum25: TPanel;
    pnlName25: TPanel;
    edtValue25: TEdit;
    pnlNum26: TPanel;
    pnlName26: TPanel;
    edtValue26: TEdit;
    pnlNum27: TPanel;
    pnlName27: TPanel;
    edtValue27: TEdit;
    pnlNum28: TPanel;
    pnlName28: TPanel;
    edtValue28: TEdit;
    pnlNum29: TPanel;
    pnlName29: TPanel;
    edtValue29: TEdit;
    pnlNum30: TPanel;
    pnlName30: TPanel;
    edtValue30: TEdit;
    pnlName31: TPanel;
    pnlNum31: TPanel;
    pnlNum32: TPanel;
    pnlName32: TPanel;
    edtValue32: TEdit;
    pnlNum33: TPanel;
    pnlName33: TPanel;
    edtValue33: TEdit;
    pnlNum34: TPanel;
    pnlName34: TPanel;
    edtValue34: TEdit;
    pnlNum35: TPanel;
    pnlName35: TPanel;
    edtValue35: TEdit;
    pnlNum36: TPanel;
    pnlName36: TPanel;
    edtValue36: TEdit;
    pnlNum37: TPanel;
    pnlName37: TPanel;
    edtValue37: TEdit;
    pnlNum38: TPanel;
    pnlName38: TPanel;
    edtValue38: TEdit;
    pnlNum39: TPanel;
    pnlName39: TPanel;
    edtValue39: TEdit;
    pnlNum40: TPanel;
    pnlName40: TPanel;
    edtValue40: TEdit;
    pnlNum41: TPanel;
    pnlName41: TPanel;
    edtValue41: TEdit;
    pnlNum42: TPanel;
    pnlName42: TPanel;
    edtValue42: TEdit;
    pnlNum43: TPanel;
    pnlName43: TPanel;
    edtValue43: TEdit;
    pnlNum44: TPanel;
    pnlName44: TPanel;
    edtValue44: TEdit;
    pnlNum45: TPanel;
    pnlName45: TPanel;
    edtValue45: TEdit;
    mskDAT_BUNJU: TDateEdit;
    edtNUM_BUNJU: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    mskNUM_JUMIN: TMaskEdit;
    Label6: TLabel;
    edtDESC_NAME: TEdit;
    qryHangmok: TQuery;
    btnNextHm: TBitBtn;
    btnPrevHm: TBitBtn;
    qryPf_hm: TQuery;
    qryGmgn: TQuery;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    edtFind: TEdit;
    Label2: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    btnDefault: TBitBtn;
    mskNum: TMaskEdit;
    Label11: TLabel;
    mskDAT_GMGN: TDateEdit;
    btnPrevNum: TBitBtn;
    btnNextNum: TBitBtn;
    qryU_Bunju: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    valChange: TValEdit;
    qryTkgum: TQuery;
    qryProfileG: TQuery;
    qryUser_priv: TQuery;
    qryGubn_hulgum: TQuery;
    Cmb_gubn: TComboBox;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
    Bevel2: TBevel;
    Label14: TLabel;
    EdtDesc_dc: TEdit;
    EdtNum_samp: TEdit;
    Label15: TLabel;
    Label16: TLabel;
    EdtDesc_dept: TEdit;
    Lbl_C078: TLabel;
    qryCheck: TQuery;
    qryProfile: TQuery;
    qryinjouk: TQuery;
    edtPValue1: TEdit;
    edtPValue2: TEdit;
    edtPValue3: TEdit;
    edtPValue4: TEdit;
    edtPValue5: TEdit;
    edtPValue6: TEdit;
    edtPValue7: TEdit;
    edtPValue8: TEdit;
    edtPValue9: TEdit;
    edtPValue10: TEdit;
    edtPValue11: TEdit;
    edtPValue12: TEdit;
    edtPValue13: TEdit;
    edtPValue14: TEdit;
    edtPValue15: TEdit;
    edtPValue16: TEdit;
    edtPValue17: TEdit;
    edtPValue18: TEdit;
    edtPValue19: TEdit;
    edtPValue20: TEdit;
    edtPValue21: TEdit;
    edtPValue22: TEdit;
    edtPValue23: TEdit;
    edtPValue24: TEdit;
    edtPValue25: TEdit;
    edtPValue26: TEdit;
    edtPValue27: TEdit;
    edtPValue28: TEdit;
    edtPValue29: TEdit;
    edtPValue30: TEdit;
    edtPValue31: TEdit;
    edtPValue32: TEdit;
    edtPValue33: TEdit;
    edtPValue34: TEdit;
    edtPValue35: TEdit;
    edtPValue36: TEdit;
    edtPValue37: TEdit;
    edtPValue38: TEdit;
    edtPValue39: TEdit;
    edtPValue40: TEdit;
    edtPValue41: TEdit;
    edtPValue42: TEdit;
    edtPValue43: TEdit;
    edtPValue44: TEdit;
    edtPValue45: TEdit;
    qryPrev: TQuery;
    qryProfile_H025: TQuery;
    qryDanche: TQuery;
    edtValue1: TEdit;
    edtValue16: TEdit;
    Edt_jisa: TEdit;
    Label17: TLabel;
    mskPDAT_BUNJU: TDateEdit;
    Label18: TLabel;
    edtPNUM_BUNJU: TEdit;
    Label19: TLabel;
    Panel2: TPanel;
    Chk_15: TCheckBox;
    Chk_12: TCheckBox;
    Chk_11: TCheckBox;
    Chk_71: TCheckBox;
    Chk_61: TCheckBox;
    Chk_51: TCheckBox;
    Chk_43: TCheckBox;
    Chk_45: TCheckBox;
    Chk_52: TCheckBox;
    Chk_34: TCheckBox;
    Chk_41: TCheckBox;
    Chk_Etc: TCheckBox;
    Label20: TLabel;
    CheckBox: TCheckBox;
    Panel3: TPanel;
    P_1: TPanel;
    ckGUBN_JUNGDO1: TCheckBox;
    mskDAT_SOLVE1: TDateEdit;
    mskDAT_SOLVE2: TDateEdit;
    ckGUBN_JUNGDO2: TCheckBox;
    P_2: TPanel;
    P_3: TPanel;
    ckGUBN_JUNGDO3: TCheckBox;
    mskDAT_SOLVE3: TDateEdit;
    mskDAT_SOLVE8: TDateEdit;
    ckGUBN_JUNGDO8: TCheckBox;
    P_8: TPanel;
    P_9: TPanel;
    ckGUBN_JUNGDO9: TCheckBox;
    mskDAT_SOLVE9: TDateEdit;
    P_6: TPanel;
    ckGUBN_JUNGDO6: TCheckBox;
    mskDAT_SOLVE6: TDateEdit;
    mskDAT_SOLVE5: TDateEdit;
    ckGUBN_JUNGDO5: TCheckBox;
    P_5: TPanel;
    P_4: TPanel;
    ckGUBN_JUNGDO4: TCheckBox;
    mskDAT_SOLVE4: TDateEdit;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel30: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    RGrp_preGulkwa: TRadioGroup;
    qryGum_bul: TQuery;
    qryinjouk1: TQuery;
    Panel4: TPanel;
    JD_sayu: TEdit;
    qrykicho: TQuery;
    edt_sinjang: TEdit;
    edt_chejung: TEdit;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    edt_gumgin_Check: TEdit;
    qryGumgin_Check: TQuery;
    edtValue31: TEdit;
    qry_munjin2010: TQuery;
    Label24: TLabel;
    Edt_munjin: TEdit;
    Panel8: TPanel;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    RB_bunju: TRadioButton;
    RB_samp: TRadioButton;
    chk_CRM: TCheckBox;
    qryU_check_CRM: TQuery;
    qryI_Check_CRM: TQuery;
    qry_Check_CRM: TQuery;
    Lbl_JJXE: TLabel;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnPInsertClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_MoveHm(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Exit(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure edtFindExit(Sender: TObject);
    procedure UP_KeyMove(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure mskNumExit(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure UP_MoveNum(Sender: TObject);
    procedure grdMasterKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnDefaultClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    procedure edtValue1Change(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure Cmb_gubnChange(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,    UV_vNum_id,     UV_vDat_bunju,  UV_vNum_bunju,  UV_vDesc_glkwa,
    UV_vDesc_Pglkwa, UV_vNum_jumin,  UV_vDesc_name,  UV_vCod_hul,    UV_vCod_cancer,
    UV_vCod_chuga,   UV_vCod_jangbi, UV_vDat_gmgn,   UV_vGubn_nosin, UV_vGubn_nsyh,
    UV_vGubn_bogun,  UV_vGubn_bgyh,  UV_vGubn_adult, UV_vGubn_adyh,  UV_vGubn_agab,
    UV_vGubn_agyh,   UV_vGubn_gong,  UV_vGubn_goyh,  UV_vCod_jisa,   UV_vGubn_tkgm,
    UV_vGubn_hulgum, UV_vNum_Sample, UV_vABO_chk,    UV_vGubn_life,  UV_vGubn_lfyh,
    UV_vCod_dc,      UV_vDesc_dept,  UV_vPDat_bunju, UV_vPNum_bunju, UV_vGumgin_check : Variant;

    // 검사항목코드
    UV_vCod_hm, UV_vTCod_hm : Variant;

    // 항목코드, 시작위치, 마지막위치, 값
    UV_sCode  : array[1..90] of String;
    UV_iStart : array[1..90] of Integer;
    UV_iEnd   : array[1..90] of Integer;
    UV_sValue : array[1..90] of String;
    UV_sPValue : array[1..90] of String;
    UV_sGubn  : array[1..90] of String;
    UV_sScale : array[1..90] of String;
    UV_fLow   : array[1..90] of Extended;
    UV_fHigh  : array[1..90] of Extended;
    UV_sUnit  : array[1..90] of String;

    // Plus되는값
    UV_iPlus : Integer;

    // part구분, 본사+지사 분주 표시
    UV_sCod_jangbi,UV_sCod_hul,UV_sGubn_part, UV_sHulgum, sHangmok : String;

    // 화면페이지
    UV_iPage : Integer;

    // 백분율Error 체크
    UV_bPerError : Boolean;
    sChk_Hm  : Boolean;
    sChk_Hm2 : Boolean;

    //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 체크...
    UV_bSE42_Check : Boolean;

    //나이
    iAge : Integer;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
    procedure UP_HangmokLocate(iStartPos : Integer);
    procedure UP_ArrayClear;
    function  UF_GetValue(sCode : String) : Extended;
    procedure UP_SetValue(sCode : String; eValue : Extended);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    procedure UP_EditDisplay(Sender : TObject);
    procedure UP_Default(Sender: TObject);
    Procedure Hangmok_Set;
    Procedure Hangmok_Set2;

  public
    { Public declarations }
    procedure UP_EnvSet(sGubn_part : String);
  end;

  const UC_First = 1;
        UC_Last  = 2;

var
  frmLD1I01: TfrmLD1I01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I011, LD1I012, LD1I013, LD1I014, math,
  LD1I015, LD1I018, LD1P011;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD1I01.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //추가
      end;
   end;
end;

procedure TfrmLD1I01.UP_SetValue(sCode : String; eValue : Extended);
var i, j, index : Integer;
    sValue : String;
begin
   sValue := FloatToStr(eValue);

   for i := 1 to 90 do
   begin
      if UV_sCode[i] = sCode then
      begin
         // 1. 정해진 자리수만큼 자름
         valChange.Scale := StrToInt(UV_sScale[i]);
         valChange.Value := eValue;
         sValue := valChange.Text;

         //sValue := Copy(sValue, 1, UV_iEnd[i] - UV_iStart[i] + 1);

         // 2. Memory에 값을 할당
         UV_sValue[i] := sValue;
         UV_sPValue[i] := sValue;

         // 3. 화면에 보이는 필드인지 Check
         if (UV_iPlus = 0)  and (i > 45) then exit;
         if (UV_iPlus = 45) and (i < 46) then exit;

         // 4. 해당 Field에 값을 보여줌
         for j := 0 to pnlDetail.ControlCount - 1 do
         begin
            if pnlDetail.Controls[j].ClassType = TEdit then
            begin
               if Pos('edtValue', TEdit(pnlDetail.Controls[j]).Name) > 0 then
               begin
                  index := StrToInt(Copy(TEdit(pnlDetail.Controls[j]).Name, 9, 2));
                  if index = i then
                     TEdit(pnlDetail.Controls[j]).Text := sValue;
               end;
            end;
         end;
      end;
   end;
end;

function  TfrmLD1I01.UF_GetValue(sCode : String) : Extended;
var i : Integer;
begin
   for i := 1 to 90 do
   begin
      if (UV_sCode[i] = sCode) and (copy(Trim(UV_sValue[65]),1,1) <> '<') then  //20140507 곽윤설
      begin
         Result := GF_StrToNum(UV_sValue[i]);
         exit;
      end;
   end;
   Result := 0;
end;

procedure TfrmLD1I01.UP_EnvSet(sGubn_part : String);
var sTitle : String;
    i : integer;
begin
   // Part구분에 따라서 환경을 설정한다.
  Label19.Visible := False;
  Panel2.Visible  := False;

   if      sGubn_part = '01' then sTitle := '생  화  학  결  과  등  록'
   else if sGubn_part = '02' then sTitle := '혈  액  학  결  과  등  록'
   else if sGubn_part = '03' then sTitle := 'URIN  결  과  등  록'
   else if sGubn_part = '04' then sTitle := 'RIA  결  과  등  록'
   else if sGubn_part = '05' then sTitle := 'EIA  결  과  등  록'
   else if sGubn_part = '07' then sTitle := '혈  청  학  결  과  등  록'
   else if sGubn_part = '08' then sTitle := '기  타  결  과  1  등  록'
   else if sGubn_part = '09' then sTitle := '기  타  결  과  2  등  록';

   // Global변수에 할당
   UV_sGubn_part := sGubn_part;

   if      sGubn_part = '01' then frmLD1I01.Caption := 'frmLD1I02_' + sTitle
   else if sGubn_part = '02' then frmLD1I01.Caption := 'frmLD1I01_' + sTitle
   else if sGubn_part = '03' then frmLD1I01.Caption := 'frmLD1I04_' + sTitle
   else if sGubn_part = '04' then frmLD1I01.Caption := 'frmLD1I05_' + sTitle
   else if sGubn_part = '05' then frmLD1I01.Caption := 'frmLD1I06_' + sTitle
   else if sGubn_part = '07' then frmLD1I01.Caption := 'frmLD1I03_' + sTitle
   else if sGubn_part = '08' then frmLD1I01.Caption := 'frmLD1I08_' + sTitle
   else if sGubn_part = '09' then frmLD1I01.Caption := 'frmLD1I09_' + sTitle;

   Panel1.Caption    := sTitle;


   if sGubn_part = '02' then   //20151109
   begin
      Cmb_gubn.Clear;
      Cmb_gubn.Items.Add('      ');
      Cmb_gubn.Items.Add('혈액형검사자');
      Cmb_gubn.Items.Add('ESR검사자');
   end
   else if (sGubn_part = '03') and (copy(GV_sUserId,1,2)='15') then
   begin
      Cmb_gubn.Clear;
      Cmb_gubn.Items.Add('      ');
      Cmb_gubn.Items.Add('Y004검사자');
      Cmb_gubn.Items.Add('니코틴검사자');
      Cmb_gubn.Items.Add('뇨침사검사자');
      Cmb_gubn.Items.Add('M2-PK검사자');
   end
   else
   begin
      Cmb_gubn.Clear;
      Cmb_gubn.Items.Add('      ');
   end;
   Cmb_gubn.ItemIndex := 0;

   if (sGubn_part = '01') or (sGubn_part = '02') then chk_CRM.Visible := True
   else chk_CRM.Visible := False;

   if (UV_sGubn_part <> '01') AND (UV_sGubn_part <> '02') and (UV_sGubn_part <> '05') and (UV_sGubn_part <> '07')   then
   begin
      Lbl_C078.Caption := '※종전결과는 생화학,혈액학,혈청,EIA 결과등록에만 표시됩니다.';
      Lbl_C078.Visible := True;
   end
   else
      Lbl_C078.Visible := False;

   // 항목Query
   with qryHangmok do
   begin
      Active := False;
      ParamByName('GUBN_PART').AsString := UV_sGubn_part;
      ParamByName('DAT_APPLY').AsString := mskDate.Text;
      Active := True;

      if RecordCount <= 0 then
      begin
         GF_MsgBox('Warning', '해당 Part에 대한 항목자료가 존재하지 않습니다.');
         exit;
      end;
   end;

   // 파트별 작업 가능케.. 04.10.01
   with qryUser_priv do
   begin
      close;
      ParamByName('cod_sawon').AsString := GV_sUserId;
      open;
      if recordcount > 0 then
      begin
         for i := 0 to RecordCount - 1 do
         begin
            if ((sGubn_part = '01') and
                ((FieldByName('cod_prog').AsString = 'LD1I02') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '02') and
                ((FieldByName('cod_prog').AsString = 'LD1I01') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '03') and
                ((FieldByName('cod_prog').AsString = 'LD1I04') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '04') and
                ((FieldByName('cod_prog').AsString = 'LD1I05') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '05') and
                ((FieldByName('cod_prog').AsString = 'LD1I06') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '07') and
                ((FieldByName('cod_prog').AsString = 'LD1I03') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '08') and
                ((FieldByName('cod_prog').AsString = 'LD1I08') and (FieldBYName('ysno_read').AsString = 'Y'))) or
               ((sGubn_part = '09') and
                ((FieldByName('cod_prog').AsString = 'LD1I09') and (FieldBYName('ysno_read').AsString = 'Y'))) then
            begin
               btnInsert.enabled := false;
               btnDelete.enabled := false;
               btnSave.enabled   := false;
               btnPSave.enabled  := false;
               btnPCancel.enabled := false;
               btnDefault.enabled := false;
               exit;
            end else
            begin
               btnInsert.enabled := true;
               btnDelete.enabled := true;
               btnSave.enabled   := true;
               btnPSave.enabled  := true;
               btnPCancel.enabled := true;
               btnDefault.enabled := true;
            end;
            next;
         end;
      end;
   end;
    qryHangmok.Active := false;
end;

procedure TfrmLD1I01.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading := '분주번호';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';
         {     20151117 주석처리
         case Cmb_gubn.ItemIndex of

            0 : begin //분주
                   Col[1].Color := $00D7FBDC;
                   Col[2].Color := clBtnFace;
                   Col[3].Color := clBtnFace;
                end;
            1 : begin //샘플
                   Col[1].Color := clBtnFace;
                   Col[2].Color := $00D7FBDC;
                   Col[3].Color := clBtnFace;
                end;
            2 : begin
                   Col[1].Color := $00D7FBDC;
                   Col[2].Color := clBtnFace;
                   Col[3].Color := clBtnFace;
                end;
         end;
         }
         if Trim(Cmb_gubn.Text) = '혈액형검사자' then
         begin
            Col[1].Color := $00D7FBDC;
            Col[2].Color := clBtnFace;
            Col[3].Color := clBtnFace;
         end
         else
         begin
            if RB_bunju.Checked then
            begin
               Col[1].Color := $00D7FBDC;
               Col[2].Color := clBtnFace;
               Col[3].Color := clBtnFace;
            end
            else if RB_samp.Checked then
            begin
               Col[1].Color := clBtnFace;
               Col[2].Color := $00D7FBDC;
               Col[3].Color := clBtnFace;
            end;
         end;
      end
      else if rbJumin.Checked then
      begin
         Col[1].Heading := '검진일자';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';

         Col[1].Color := clBtnFace;
         Col[2].Color := $00D7FBDC;
         Col[3].Color := clBtnFace;
      end
      else if rbBarcode.Checked then
      begin
         Col[1].Heading := '분주번호';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';

         Col[1].Color := clBtnFace;
         Col[2].Color := $00D7FBDC;
         Col[3].Color := clBtnFace;
      end;

      // Set Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
   end;
end;

procedure TfrmLD1I01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_glkwa := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pglkwa := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cancer := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_tkgm  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nosin := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nsyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bogun := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bgyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adult := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agab  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_gong  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_goyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_hulgum := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Sample := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_life  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_lfyh  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_dc     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dept  := VarArrayCreate([0, 0], varOleStr);
      UV_vPDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vPNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vGumgin_check := VarArrayCreate([0, 0], varOleStr);
      UV_vABO_chk    := VarArrayCreate([0, 0], varOleStr);
      UV_vTCod_HM    := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vDat_bunju,  iLength);
      VarArrayRedim(UV_vNum_bunju,  iLength);
      VarArrayRedim(UV_vDesc_glkwa, iLength);
      VarArrayRedim(UV_vDesc_Pglkwa, iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vCod_jangbi, iLength);
      VarArrayRedim(UV_vCod_hul,    iLength);
      VarArrayRedim(UV_vCod_cancer, iLength);
      VarArrayRedim(UV_vCod_chuga,  iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vGubn_tkgm,  iLength);
      VarArrayRedim(UV_vGubn_nosin, iLength);
      VarArrayRedim(UV_vGubn_nsyh,  iLength);
      VarArrayRedim(UV_vGubn_bogun, iLength);
      VarArrayRedim(UV_vGubn_bgyh,  iLength);
      VarArrayRedim(UV_vGubn_adult, iLength);
      VarArrayRedim(UV_vGubn_adyh,  iLength);
      VarArrayRedim(UV_vGubn_agab,  iLength);
      VarArrayRedim(UV_vGubn_agyh,  iLength);
      VarArrayRedim(UV_vGubn_gong,  iLength);
      VarArrayRedim(UV_vGubn_goyh,  iLength);
      VarArrayRedim(UV_vGubn_hulgum, iLength);
      VarArrayRedim(UV_vNum_Sample, iLength);
      VarArrayRedim(UV_vGubn_life,  iLength);
      VarArrayRedim(UV_vGubn_lfyh,  iLength);
      VarArrayRedim(UV_vCod_dc,     iLength);
      VarArrayRedim(UV_vDesc_dept,  iLength);
      VarArrayRedim(UV_vPDat_bunju, iLength);
      VarArrayRedim(UV_vPNum_bunju, iLength);
      VarArrayRedim(UV_vGumgin_check, iLength);
      VarArrayRedim(UV_vABO_chk,    iLength);
      VarArrayRedim(UV_vTCod_HM,    iLength);
   end;
end;

function TfrmLD1I01.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (pnlJisa.Visible and (cmbJisa.ItemIndex = -1)) or
      (rbDate.Checked and (mskDate.Text = '')) or
      (rbJumin.Checked and (mskJumin.Text = '')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

function TfrmLD1I01.UF_Save : Boolean;
var i : Integer;
    sGlkwa : String;
begin
   Result := False;
   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;
   //CRM 체크 버튼, 작업 여부와 상관없이 체크 되면 저장 되도록 요청
   with qry_Check_CRM do
   begin
      Close;
      qry_Check_CRM.ParamByName('COD_jisa').AsString := UV_vCod_jisa[i];
      qry_Check_CRM.ParamByName('DAT_gmgn').AsString := UV_vdat_gmgn[i];
      qry_Check_CRM.ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
      Open;
      if qry_Check_CRM.RecordCount > 0 then
      Begin
         if chk_CRM.Checked = false then
         begin
         with qryU_Check_CRM do
             begin
                ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
                ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
                ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
                ParamByName('YSNO_CRM').AsString    := 'N';
                ParamByName('DAT_update').AsString  := GV_sToday;
                ParamByName('COD_update').AsString  := GV_sUserId;
                Execsql;

                GP_query_log(qryU_Check_CRM, 'LD1I13', 'U', 'Y');
             end;
          end
          else
             with qryU_Check_CRM do
             begin
                ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
                ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
                ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
                ParamByName('YSNO_CRM').AsString    := 'Y';
                ParamByName('DAT_update').AsString  := GV_sToday;
                ParamByName('COD_update').AsString  := GV_sUserId;
                Execsql;

                GP_query_log(qryU_Check_CRM, 'LD1I13', 'U', 'Y');
             end;
      end
      else
      if chk_CRM.Checked = TRUE then
      begin
      with qryI_Check_CRM do
          begin
             ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
             ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
             ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
             ParamByName('YSNO_CRM').AsString    := 'Y';
             ParamByName('DAT_INSERT').AsString  := GV_sToday;
             ParamByName('COD_INSERT').AsString  := GV_sUserId;
             ParamByName('DAT_update').AsString  := '';
             ParamByName('COD_update').AsString  := '';
             Execsql;

             GP_query_log(qryI_Check_CRM, 'LD1I13', 'I', 'Y');
          end;
      end;
   end;

   // 작업중인지 Check
   if not UV_bEdit then exit;

   // Validation Check
   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   //if not GF_MsgBox('Confirmation', 'SAVE') then exit; //2015.09.24 본원 진단검사 혈액학팀 요청

   if (UV_sGubn_part = '02') and (UV_iPage = 1) then
   begin
      if  Trim(edtValue12.Text) <> '' then
      begin
         if GF_StrToNum(edtValue12.Text) + GF_StrToNum(edtValue13.Text) + GF_StrToNum(edtValue14.Text)
          + GF_StrToNum(edtValue15.Text) + GF_StrToNum(edtValue16.Text) + GF_StrToNum(edtValue17.Text)
          + GF_StrToNum(edtValue18.Text) + GF_StrToNum(edtValue19.Text) + GF_StrToNum(edtValue20.Text)
          + GF_StrToNum(edtValue21.Text) + GF_StrToNum(edtValue22.Text) <> 100 then
          begin
            GF_MsgBox('Warning', '백혈구 백분율 Error');
            exit;
          end;
      end;
   end
   else if (UV_sGubn_part = '02') and (UV_iPage = 2) then
   begin
      if UV_bPerError then
      begin
         UP_HangmokLocate(UC_First);
         GF_MsgBox('Warning', '백혈구 백분율 Error');
         exit;
      end;
   end;
   sGlkwa := '';
   for i := 1 to 15 do
   begin
      sGlkwa := sGlkwa + '                                                            ';
   end;

   // 결과자료를 생성
   for i := 1 to 90 do
   begin

      // 자료가 존재할경우
      if UV_sCode[i] <> '' then
      begin
         sGlkwa := Copy(sGlkwa, 1, UV_iStart[i]-1) + GF_RPad(UV_sValue[i], UV_iEnd[i] - UV_iStart[i] + 1, ' ')
                 + Copy(sGlkwa, UV_iEnd[i]+1, Length(sGlkwa));
      end;
   end;
   sGlkwa := 'A' + sGlkwa;
   sGlkwa := Trim(sGlkwa);
   sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   dmComm.database.StartTransaction;
   try
      with qryU_Glkwa do
      begin
         ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[i];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString  := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString  := edtNUM_BUNJU.Text;
         ParamByName('GUBN_PART').AsString  := UV_sGubn_part;
         UV_fGulkwa1 := '';
         UV_fGulkwa2 := '';
         UV_fGulkwa3 := '';
         GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);
         ParamByName('DESC_GLKWA1').AsString := UV_fGulkwa1;
         ParamByName('DESC_GLKWA2').AsString := UV_fGulkwa2;
         ParamByName('DESC_GLKWA3').AsString := UV_fGulkwa3;
         ParamByName('DAT_UPDATE').AsString := GV_sToday;
         ParamByName('COD_UPDATE').AsString := GV_sUserId;

         Execsql;
         GP_query_log(qryU_Glkwa, 'LD1I01', 'Q', 'Y');
      end;

      with qryU_Bunju do
      begin
         ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[i];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString  := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString  := edtNUM_BUNJU.Text;

         Execsql;
         GP_query_log(qryU_Bunju, 'LD1I01', 'Q', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;

   // database commit
   dmComm.database.Commit;

   // Grid에 자료를 반영
   with grdMaster do
   begin
      // 현재 Field의 자료를 Grid관련 Variant변수에 할당
      UV_vDesc_glkwa[i] := sGlkwa;
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD1I01.UP_HangmokLocate(iStartPos : Integer);
var sName, sTemp, sMan, sGubn : String;
    i, j, iPos, iTemp, iHulgum, iSpace : Integer;
    bFocus : boolean; // 두번째 커서 위치
begin
   if iStartPos = UC_First then
   begin
      UV_iPlus := 0;
      UV_iPage := 1;
   end
   else if iStartPos = UC_Last  then
   begin
      UV_iPlus := 45;
      UV_iPage := 2;
      bFocus := False;
   end;

   // 주민번호를 이용하여 성별과 나이를 구함
   GP_JuminToAge(mskNUM_JUMIN.Text,mskDAT_GMGN.Text, sMan, iAge);

   if iAge < 10 then sGubn := '1'
   else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
   else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
   else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
   else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
   else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
   else sGubn := '7';

   if sMan = 'F' then sGubn := sGubn + 'f';

   // 번호설정 & 검사항목을 Clear
   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if pnlDetail.Controls[i].ClassType = TPanel then
      begin
         sName := TPanel(pnlDetail.Controls[i]).Name;
         // 번호에 대한 작업
         if Pos('Num', sName) > 0 then
            TPanel(pnlDetail.Controls[i]).Caption := IntToStr(StrToInt(Copy(sName, 7, 2)) + UV_iPlus);

         // 항목이름에 대한 작업(Clear)
         if Pos('Name', sName) > 0 then
         begin
            TPanel(pnlDetail.Controls[i]).Caption := '';
            TPanel(pnlDetail.Controls[i]).Color   := clBtnFace;
         end;
      end;

      if pnlDetail.Controls[i].ClassType = TEdit then
      begin
         sName := TEdit(pnlDetail.Controls[i]).Name;
         // 결과값에 대한 작업
         if Pos('edtValue', sName) > 0 then
         begin
            TEdit(pnlDetail.Controls[i]).Enabled := False;
            TEdit(pnlDetail.Controls[i]).Color   := clBtnFace;
            TEdit(pnlDetail.Controls[i]).Text    := '';
         end
         else if Pos('edtPValue', sName) > 0 then
            TEdit(pnlDetail.Controls[i]).Text    := '';
      end;
   end;

   with qryHangmok do
   begin
      // 첫번째 자료
   //   First;
      Active := False;
      ParamByName('GUBN_PART').AsString := UV_sGubn_part;
      ParamByName('DAT_APPLY').AsString := mskDAT_GMGN.Text;
      Active := True;
      while not Eof do
      begin
         sName := FieldByName('DESC_KOR').AsString;
         iPos  := FieldByName('POS_GLKWA').AsInteger;
{
         // 처음검사일 경우 45이상이면 Break
         if iStartPos = UC_First then
            if iPos > 45 then break;
}
         // 마지막 자료일 경우 45이후에 작업수행
         if iStartPos = UC_Last then
            if iPos <= 45 then
            begin
               Next;
               continue;
            end;

         // 검사한 항목인지 Check
         for i := 0 to VarArrayHighBound(UV_vCod_hm, 1) - 1 do
         begin
            iHulgum := 0;

            if UV_vCod_hm[i] = FieldByName('COD_HM').AsString then
            begin
               if UV_sCode[iPos] = '' then
               begin
                  UV_sCode[iPos]  := qryHangmok.FieldByName('COD_HM').AsString;
                  UV_iStart[iPos] := qryHangmok.FieldByName('POS_START').AsInteger;
                  UV_iEnd[iPos]   := qryHangmok.FieldByName('POS_END').AsInteger;
                  UV_sGubn[iPos]  := qryHangmok.FieldByName('GUBN_HM').AsString;
                  UV_sScale[iPos] := qryHangmok.FieldByName('POS_POINT').AsString;
                  UV_fLow[iPos]   := qryHangmok.FieldByName('IMS_LOW'   + sGubn).AsFloat;
                  UV_fHigh[iPos]  := qryHangmok.FieldByName('IMS_HIGH'  + sGubn).AsFloat;
                  UV_sUnit[iPos]  := qryHangmok.FieldByName('DESC_UNIT').AsString;
                  if (UV_iStart[iPos] <> 0) and (UV_iEnd[iPos] <> 0) then
                  begin
                     sTemp := Copy(UV_vDesc_glkwa[grdMaster.CurrentDataRow-1],
                                   UV_iStart[iPos], UV_iEnd[iPos] - UV_iStart[iPos] + 1);
                     // 문자열에 존재하는 Space를 제거
                     UV_sValue[iPos] := GF_SpaceDel(sTemp);
                     sTemp := Copy(UV_vDesc_Pglkwa[grdMaster.CurrentDataRow-1],
                                   UV_iStart[iPos], UV_iEnd[iPos] - UV_iStart[iPos] + 1);
                     UV_sPValue[iPos] := GF_SpaceDel(sTemp);
                  end;
               end;

               // 처음검사일 경우 45이상이면 Break
               if iStartPos = UC_First then
                 if iPos > 45 then break;

               // 해당 위치필드에 대한 작업(항목명칭, 활성화조절)
               for j := 0 to pnlDetail.ControlCount - 1 do
               begin
                  if pnlDetail.Controls[j].ClassType = TPanel then
                  begin
                     // Name Panel에 대헤서 항목명을 할당
                     if Pos('Name', TPanel(pnlDetail.Controls[j]).Name) > 0 then
                     begin
                        iTemp := StrToInt(Copy(TPanel(pnlDetail.Controls[j]).Name, 8, 2));

                        if (iTemp + UV_iPlus) = iPos then
                           TPanel(pnlDetail.Controls[j]).Caption := '  ' + sName;

                        TPanel(pnlDetail.Controls[j]).Color := clBtnFace;

                        //[2018.07.19-유동구]일산화탄소 검사 중 호기중 검사 진행 시 카복시 미 진행으로 인해 색깔 표기..
                        if (UV_bSE42_Check) and
                           (trim(TPanel(pnlDetail.Controls[j]).caption) = '혈중카복시헤모글로빈') then
                        begin
                           TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                        end;

                        // 혈액학(2011220)
                        if (mskDate.Text >= '20120301') and (UV_sHulgum = '3') and
                           ((trim(TPanel(pnlDetail.Controls[j]).caption) = '적혈구수')    or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '혈색소')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Hematocrit')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'MCV')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'MCH')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'MCHC')        or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'RDW')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'MPV')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'PDW')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '혈소판수')    or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '백혈구수')    or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '분획호중구')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '봉상호중구')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '임파구')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '단핵구')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '호산구')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '호염기성구')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '후골수구')    or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '골수구')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '전골수세포')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '골수아세포')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '유핵적혈구')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ABO형')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Rh형')        or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ESR')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'PCT'))        then
                        begin
                           iHulgum := iTemp;
                           TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                        end;

                        // 생화학(20090531)
                       //2015.12.24 수원자체 검사 9종
                       //2016.08.20채용만 있으면 연구소검사, 그외 나머지 자체9종은 전부 수원지사 검사
                       //2017.10.13 수원지사 -> 전국센터로 생화학 자체 9종 확대 ..20171016일자 검진 일자 부터
                       //2018.02.20 채용상관없이 생화학 9종 자체검사
                       //19.07.02 한두례 책임 센터 수검자 식후혈당 검사 시 본원에서 결과창 막도록(전국센터 합의)
                       if (Edt_jisa.Text <> '15') and (mskDAT_GMGN.Text >= '20180302') and (copy(GV_sUserId,1,2)='15') then
                       begin
                           if (UV_sHulgum = '3') and
                           ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')             or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')             or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')                 or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')          or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')              or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')        or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')        or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')              or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤(효소)')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '식후혈당')              or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌'))           then
                           begin
                              iHulgum := iTemp;
                              TEdit(pnlDetail.Controls[j]).Enabled := FALSE;
                              TPanel(pnlDetail.Controls[j]).Color :=  $00F9D9F5;
                           end;
                       END
                       else if       (Edt_jisa.Text <> '15') and (mskDAT_GMGN.Text >= '20171016')
                                 and (copy(GV_sUserId,1,2)='15') and ((copy(UV_sCod_jangbi,1,2) <> 'SS') AND (copy(UV_sCod_jangbi,1,2) <> 'GS')) then
                       begin
                              if (UV_sHulgum = '3') and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')                 or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤(효소)')  or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식후혈당')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌'))           then
                              begin
                                 iHulgum := iTemp;
                                 TEdit(pnlDetail.Controls[j]).Enabled := FALSE;
                                 TPanel(pnlDetail.Controls[j]).Color :=  $00F9D9F5;
                              end;
                        END
                        else if (Edt_jisa.Text <> '15') and (mskDAT_GMGN.Text >= '20171016')
                            and (copy(GV_sUserId,1,2)='15')
                            and ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '1')  or (Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '2') or
                                 (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '1')  or (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '2') or
                                 (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '1')  or (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '2'))and
                               ((copy(UV_sCod_jangbi,1,2) = 'SS') or (copy(UV_sCod_jangbi,1,2) = 'GS')) then
                        begin
                              if (UV_sHulgum = '3') and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')                 or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤(효소)')  or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식후혈당')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌'))           then
                              begin
                                 iHulgum := iTemp;
                                 TEdit(pnlDetail.Controls[j]).Enabled := FALSE;
                                 TPanel(pnlDetail.Controls[j]).Color :=  $00F9D9F5;
                              end;
                        END
                        else if (mskDAT_GMGN.Text >= '20150728') then                //20150729 문지혜 요청(20150729분주일부터)
                        begin
                           if (UV_sHulgum = '3') and
                              ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '1')  or
                               (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '1')  or
                               (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '1')) and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')           or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')           or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')         ) then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                         {  if (UV_sHulgum = '3') and                   //20150408 곽윤설. 노신,성인병 재검일때 센터자체검사(C032)는 센터자체검사로..
                              (((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '2') or
                                (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '2') or
                                (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '2')) and
                               (pos('C032',UV_vCod_chuga[grdMaster.CurrentDataRow-1]) > 0)) and
                              (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당') then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end; }
                        end
                        // 생화학(20090531)
                        else if (mskDAT_GMGN.Text >= '20140623') and (mskDAT_GMGN.Text <= '20150727') then                //20140620 곽윤설
                        begin
                           if (UV_sHulgum = '3') and
                              ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '1')  or
                               (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '1')  or
                               (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '1')) and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')             or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')           or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식전자가혈당측정기') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤(효소)')) then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                           if (UV_sHulgum = '3') and
                              ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '2')  or
                               (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '2')  or
                               (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '2')) and
                              (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당') then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end
                        else if (mskDAT_GMGN.Text >= '20140308')  and (mskDAT_GMGN.Text < '20140623') and
                                (copy(UV_sCod_jangbi,1,2)<> 'SS') and (copy(UV_sCod_Hul,1,2)<> 'SS') and
                                (copy(UV_sCod_jangbi,1,2)<> 'GS') and (copy(UV_sCod_Hul,1,2)<> 'GS') then
                        begin
                           if (UV_sHulgum = '3') and
                              ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '1')  or
                               (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '1')  or
                               (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '1')) and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')       or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식전 자가혈당측정기')) then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end
                        else if (mskDAT_GMGN.Text >= '20130805')  and (mskDAT_GMGN.Text < '20140308') and
                                (Edt_jisa.Text <>'12') and (Edt_jisa.Text <>'15')  and
                                (copy(UV_sCod_jangbi,1,2)<> 'SS') and (copy(UV_sCod_Hul,1,2)<> 'SS') and
                                (copy(UV_sCod_jangbi,1,2)<> 'GS') and (copy(UV_sCod_Hul,1,2)<> 'GS')then
                           begin
                           if (UV_sHulgum = '3') and
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총단백')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '알부민')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '글로부린')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'A/G 비율')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총빌리루빈')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '직접빌리루빈')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '간접빌리루빈')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LAP')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALP')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '요산')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'CPK')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDH')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')    or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')    or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'cardiac risk index')or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식후혈당')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '요소질소')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'BUN/Cr 비율')       then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end
                        else if (mskDAT_GMGN.Text >= '20130729')  and (mskDAT_GMGN.Text < '20140308') and
                                (Edt_jisa.Text ='12') and
                                (copy(UV_sCod_jangbi,1,2)<> 'SS') and (copy(UV_sCod_Hul,1,2)<> 'SS') and
                                (copy(UV_sCod_jangbi,1,2)<> 'GS') and (copy(UV_sCod_Hul,1,2)<> 'GS')  then
                        begin
                           if (UV_sHulgum = '3') and
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총단백')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '알부민')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '글로부린')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'A/G 비율')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총빌리루빈')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '직접빌리루빈')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '간접빌리루빈')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')            or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LAP')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALP')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '요산')              or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'CPK')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDH')               or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤')    or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤')    or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'cardiac risk index')or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식후혈당')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '요소질소')          or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')        or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'BUN/Cr 비율')       then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end
                        else if (mskDAT_GMGN.Text > '20120301')  and  (mskDAT_GMGN.Text <= '20130729')then
                        begin
                           if (UV_sHulgum = '3') and
                              ((Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '1')  or
                               (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1])  = '1')  or
                               (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '1')) and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')      or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')         or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')       or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌')     or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = '식전 자가혈당측정기')) then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end
                        else if (mskDate.Text > '20090531') and (mskDAT_GMGN.Text < '20120301') then
                        begin
                           if (UV_sHulgum = '3')  and
                           ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'AST(SGOT)')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'ALT(SGPT)')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'r-GTP')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '총콜레스테롤')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '공복혈당')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'HDL-콜레스테롤') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'LDL-콜레스테롤') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '중성지방')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '크레아티닌'))    then
                           begin
                              iHulgum := iTemp;
                              TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                           end;
                        end;
                        // 뇨화학(20111220)
                        if (mskDate.Text >= '20120301') and (UV_sHulgum = '3') and
                           ((trim(TPanel(pnlDetail.Controls[j]).caption) = '백혈구(요검사)')   or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Nitrite(요검사)')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'PH (요검사)')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '단백(요검사)')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '당(요검사)')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '케톤(요검사)')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '유로빌리노겐(요)') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '빌리루빈(요검사)') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '적혈구(요검사)')   or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '요비중')           or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '뇨침사 WBC')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '뇨침사 RBC')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '상피세포')         or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '분변잠혈'))        then
                        begin
                           iHulgum := iTemp;
                           TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                        end
                        //수원센터 테스트 위함.
                        else if (mskDate.Text >= '20120131') and (Edt_jisa.Text = '43') and
                           ((trim(TPanel(pnlDetail.Controls[j]).caption) = '백혈구(요검사)')   or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Nitrite(요검사)')  or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = 'PH (요검사)')      or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '단백(요검사)')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '당(요검사)')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '케톤(요검사)')     or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '유로빌리노겐(요)') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '빌리루빈(요검사)') or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '적혈구(요검사)')   or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '요비중')           or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '뇨침사 WBC')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '뇨침사 RBC')       or
                            (trim(TPanel(pnlDetail.Controls[j]).caption) = '상피세포'))        then
                        begin
                           iHulgum := iTemp;
                           TPanel(pnlDetail.Controls[j]).Color := $00F9D9F5;
                        end;

                        if Trim(UV_vDat_bunju[grdMaster.CurrentDataRow-1]) >= '20060911' then
                        begin
                           // 2006.09.08 [B100,A100,AB01,AB02 - 검사 중 몇가지 검사 체크]
                           if ((Trim(UV_vCod_hul[grdMaster.CurrentDataRow-1]) = 'B100') or
                               (Trim(UV_vCod_hul[grdMaster.CurrentDataRow-1]) = 'AB01') or
                               (Trim(UV_vCod_hul[grdMaster.CurrentDataRow-1]) = 'AB02'))
                              and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'Cholinesterase') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Ca')) then
                           begin
                              TPanel(pnlDetail.Controls[j]).Color := $00C8FFFF;
                           end;

                           // 2006.09.08 [B100,A100,AB01,AB02 - 검사 중 몇가지 검사 체크]
                           if (UV_vCod_hul[grdMaster.CurrentDataRow-1] = 'A100') and
                              ((trim(TPanel(pnlDetail.Controls[j]).caption) = 'Cholinesterase') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Ca') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'I .Phosphorus') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'Mg') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'TTT') or
                               (trim(TPanel(pnlDetail.Controls[j]).caption) = 'CRP')) then
                           begin
                              TPanel(pnlDetail.Controls[j]).Color := $00C8FFFF;
                           end;
                        end;
                     end;
                  end;

                  if pnlDetail.Controls[j].ClassType = TEdit then
                  begin
                     // 값항목에 대해서 활성화 조절 및 값을 할당
                     //------------------------------------------------------
                     if Pos('edtValue', TEdit(pnlDetail.Controls[j]).Name) > 0 then
                     begin
                        iTemp := StrToInt(Copy(TEdit(pnlDetail.Controls[j]).Name, 9, 2));
                        if (iTemp + UV_iPlus) = iPos then
                        begin
                           TEdit(pnlDetail.Controls[j]).Enabled   := True;
                           TEdit(pnlDetail.Controls[j]).Text      := UV_sValue[iPos];
                           TEdit(pnlDetail.Controls[j]).MaxLength := UV_iEnd[iPos] - UV_iStart[iPos] + 1;

                           // '본사+지사'- '분주:연구소'일 경우 본사에서 입력못하도록 false
                           if (UV_sHulgum = '3') and (iHulgum = iTemp) then
                              TEdit(pnlDetail.Controls[j]).Enabled := False;

                           // 다음페이지 항목에 선택되도록..
                           if iStartPos = UC_Last then
                           begin                       //20140620 곽윤설
                              if (bFocus = False) and (TEdit(pnlDetail.Controls[j]).Enabled = True) then
                              begin
                                 TEdit(pnlDetail.Controls[j]).SetFocus;
                                 bFocus := True;
                              end;
                           end;

                           if (Trim(UV_sValue[iPos]) <> '')  then
                           begin
                              if UV_sGubn[iPos] = '1' then
                              begin
                                 if (copy(Trim(UV_sValue[65]),1,1) = '<') or //20140507 곽윤설
                                    (GF_StrToFloat(UV_sValue[iPos]) < GF_StrToFloat(FloatToStr(UV_fLow[iPos]))) or
                                    (GF_StrToFloat(UV_sValue[iPos]) > GF_StrToFloat(FloatToStr(UV_fHigh[iPos]))) then
                                     TEdit(pnlDetail.Controls[j]).Color     := $00FACDF4
                                 else
                                    TEdit(pnlDetail.Controls[j]).Color     := clWindow;
                              end
                              else if UV_sGubn[iPos] = '2' then
                              begin
                                 if UV_sValue[iPos] = UV_sUnit[iPos] then
                                    TEdit(pnlDetail.Controls[j]).Color     := clWindow
                                 else
                                    TEdit(pnlDetail.Controls[j]).Color     := $00FACDF4;
                              end
                              else if UV_sGubn[iPos] = '4' then
                              begin
                                 if UV_sValue[iPos] = UV_sUnit[iPos] then
                                    TEdit(pnlDetail.Controls[j]).Color     := clWindow
                                 else
                                    TEdit(pnlDetail.Controls[j]).Color     := $00FACDF4;
                              end;
                           end
                           else
                           begin
                              if (TEdit(pnlDetail.Controls[j]).Text = '') AND ((UV_sGubn_part = '01') or  (UV_sGubn_part = '02') OR (UV_sGubn_part = '04')) then
                                 TEdit(pnlDetail.Controls[j]).Color := $0088FF88
                              else
                                 TEdit(pnlDetail.Controls[j]).Color := clWindow;
                           end;

                           //[2018.08.13-유동구]일산화탄소 검사 중 호기중 검사 진행 시 카복시 미 진행으로 인해 색깔 표기..
                           //---------------------------------------------------
                           if UV_sCode[iPos] = 'SE42' then
                           begin
                              if UV_bSE42_Check then TEdit(pnlDetail.Controls[j]).Color := $00D5D5AA
                              else                   TEdit(pnlDetail.Controls[j]).Color := clWindow;
                           end;
                           //===================================================

                           //20140418 곽윤설 - 혈액구분이 센터, 연구소+센터인 C077항목은 입력 못하도록
                           if (UV_sHulgum = '3') and (Edt_jisa.Text <> '15') then
                           begin
                              if UV_sCode[iPos]='C077' then
                              begin
                                 TPanel(pnlDetail.Controls[j-1]).Color := $00F9D9F5;
                                 TEdit(pnlDetail.Controls[j]).Enabled := False;
                              end
                               //20151103
                              else if (UV_vGubn_tkgm[grdMaster.CurrentDataRow-1] = '2') and
                                      (Trim(UV_vGubn_nosin[grdMaster.CurrentDataRow-1]) = '') and
                                      (Trim(UV_vGubn_adult[grdMaster.CurrentDataRow-1]) = '') and
                                      (Trim(UV_vGubn_life[grdMaster.CurrentDataRow-1]) = '') and
                                      (pos('C032', UV_vTCod_HM[grdMaster.CurrentDataRow-1]) > 0) then
                              begin
                                 TEdit(pnlDetail.Controls[j]).Enabled := True;
                              end;
                           end;
                        end;
                     end
                     else if Pos('edtPValue', TEdit(pnlDetail.Controls[j]).Name) > 0 then
                     begin
                        iTemp := StrToInt(Copy(TEdit(pnlDetail.Controls[j]).Name, 10, 2));
                        if (iTemp + UV_iPlus) = iPos then
                        begin
                           TEdit(pnlDetail.Controls[j]).Text      := UV_sPValue[iPos];
                           TEdit(pnlDetail.Controls[j]).MaxLength := UV_iEnd[iPos] - UV_iStart[iPos] + 1;
                        end;
                     end;
                     //------------------------------------------------------
                  end;
               end;   // End of for(j)
            end;   // End of for(i)
         end;
      Next;
      end;   // End of while(unitl Eof)
   end;
  qryHangmok.Active := False;
end;

procedure TfrmLD1I01.UP_ArrayClear;
var i : Integer;
begin
   for i := 1 to 90 do
   begin
      UV_sCode[i]  := '';
      UV_iStart[i] := 0;
      UV_iEnd[i]   := 0;
      UV_sValue[i] := '';
      UV_sPValue[i] := '';
      UV_sGubn[i]  := '';
      UV_sScale[i] := '';
      UV_sUnit[i]  := '';
      UV_fLow[i]   := 0;
      UV_fHigh[i]  := 0;
   end;
end;

procedure TfrmLD1I01.btnSaveClick(Sender: TObject);
begin
   inherited;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I01.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);
   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   if      copy(GV_sUserId,1,2)='15' then Chk_15.Checked := True
   else if copy(GV_sUserId,1,2)='12' then Chk_12.Checked := True
   else if copy(GV_sUserId,1,2)='11' then Chk_11.Checked := True
   else if copy(GV_sUserId,1,2)='43' then Chk_43.Checked := True
   else if copy(GV_sUserId,1,2)='71' then Chk_71.Checked := True
   else if copy(GV_sUserId,1,2)='61' then Chk_61.Checked := True
   else if copy(GV_sUserId,1,2)='51' then Chk_51.Checked := True;

   Lbl_JJXE.Visible := False;   
end;

procedure TfrmLD1I01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
    // 자료를 화면에 조회
    case DataCol of
      1 : begin
             if rbDate.Checked then
                Value := UV_vNum_bunju[DataRow-1]
             else if rbJumin.Checked then
                Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1])
             else if rbBarcode.Checked then
                Value := UV_vNum_bunju[DataRow-1];

             // 본사+지사 분주 색깔표시
             if (UV_vGubn_hulgum[DataRow-1] = '3') and (UV_sGubn_part = '01') and
                ((UV_vGubn_nosin[DataRow-1] = '1') or (UV_vGubn_life[DataRow-1] = '1') or (UV_vGubn_adult[DataRow-1] = '1')) then
             begin
                grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
             end
             else if (UV_vGubn_hulgum[DataRow-1] = '3') and ((UV_sGubn_part = '02') or (UV_sGubn_part = '03')) then
             begin
                grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
             end;
             //20180220 채용 상관없이 센터 생화학 자체검사 시행
             If (UV_vGubn_hulgum[DataRow-1] = '3') and (UV_sGubn_part = '01') and (UV_vCod_jisa[DataRow-1] <> '15') then grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5

             //채용만 있으면 연구소검사, 그외 나머지 자체9종은 전부 수원지사 검사   -> 전체센터로 확대(171014일 검진일자 부터)
             {If   (UV_vGubn_hulgum[DataRow-1] = '3') and (UV_sGubn_part = '01') and (UV_vCod_jisa[DataRow-1] <> '15') and
                  ((Copy(UV_vCod_Jangbi[DataRow-1],1,2) <> 'SS') AND (Copy(UV_vCod_Jangbi[DataRow-1],1,2) <> 'GS')) then grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5
         else if  (UV_vGubn_hulgum[DataRow-1] = '3') and (UV_sGubn_part = '01') and (UV_vCod_jisa[DataRow-1] <> '15') and
                  (((UV_vGubn_nosin[DataRow-1] = '1') or (UV_vGubn_life[DataRow-1] = '1') or (UV_vGubn_adult[DataRow-1] = '1')) and
                  ((Copy(UV_vCod_Jangbi[DataRow-1],1,2) = 'SS') OR (Copy(UV_vCod_Jangbi[DataRow-1],1,2) = 'GS')))
                  then grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5 ;  }
         end;
      2 : begin
             Value := UV_vNum_Sample[DataRow-1];
          end;
      3 : begin
             Value := UV_vDesc_name[DataRow-1];
          end;
   end;
 end;
procedure TfrmLD1I01.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sJisa : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   UV_bPerError := False;

   UV_iPlus := 1;

   // Grid 환경설정
   UP_GridSet;

   // 분주번호에 0을 채워줌
   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   // Query 수행 & 배열에 저장
   index := 0;
   sJisa := '';
   with qryGlkwa do
   begin
      // 조건이 없으므로 바로 Open시킴
      Active := False;

      // SQL문을 생성
      if GV_sJisa = '00' then
      begin
         // '00'이면 지사선택한 자료만 조회
         if (Panel1.Caption = '기  타  결  과  2  등  록') and (Copy(cmbJisa.Text, 1, 2) = '15') then
         begin
            sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
            sWhere := sWhere + ' AND B.cod_jisa <> ''51''';
         end
         else if (Panel1.Caption = '혈  액  학  결  과  등  록') and (Cmb_gubn.Text = 'ESR검사자') and (copy(GV_sUserId,1,2)='15') then
         begin
            if Chk_15.Checked then sJisa := sJisa + '''15'',';
            if Chk_12.Checked then sJisa := sJisa + '''12'',';
            if Chk_11.Checked then sJisa := sJisa + '''11'',';
            if Chk_43.Checked then sJisa := sJisa + '''43'',';
            if Chk_71.Checked then sJisa := sJisa + '''71'',';
            if Chk_61.Checked then sJisa := sJisa + '''61'',';
            if Chk_51.Checked then sJisa := sJisa + '''51'',';
            if Chk_45.Checked then sJisa := sJisa + '''45'',';
            if Chk_52.Checked then sJisa := sJisa + '''52'',';
            if Chk_34.Checked then sJisa := sJisa + '''34'',';
            if Chk_41.Checked then sJisa := sJisa + '''41'',';
            if Chk_Etc.Checked then sJisa := sJisa + '''62'',''80'',''80'',';

            if sJisa <> '' then
            begin
               sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
               sWhere := sWhere + ' AND B.cod_jisa in (' + copy(sJisa,1, Length(sJisa)-1) + ')';
            end
            else
            begin
               showmessage('검진센터를 하나 이상 선택해주세요!');
               exit;
            end;
         end
         else
            sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
      end
      else
      begin
         // '00'이 아니면 해당지사에 대한 자료를 조회
         if (Panel1.Caption = '기  타  결  과  2  등  록') and (GV_sJisa = '51') then
         begin
            sWhere := 'WHERE (A.cod_bjjs = ''15'' or A.cod_bjjs = ''' + GV_sJisa + ''')';
            sWhere := sWhere + ' AND B.cod_jisa = ''' + GV_sJisa + '''';
         end
         else
         begin
            if (Panel1.Caption = '기  타  결  과  2  등  록') and (GV_sJisa = '15') then
            begin
               sWhere := 'WHERE A.cod_bjjs = ''' + GV_sJisa + '''';
               sWhere := sWhere + ' AND B.cod_jisa <> ''51''';
            end
            else if (Panel1.Caption = '혈  액  학  결  과  등  록') and (Cmb_gubn.Text = 'ESR검사자') and (copy(GV_sUserId,1,2)='15') then
            begin
               if Chk_15.Checked then sJisa := sJisa + '''15'',';
               if Chk_12.Checked then sJisa := sJisa + '''12'',';
               if Chk_11.Checked then sJisa := sJisa + '''11'',';
               if Chk_43.Checked then sJisa := sJisa + '''43'',';
               if Chk_71.Checked then sJisa := sJisa + '''71'',';
               if Chk_61.Checked then sJisa := sJisa + '''61'',';
               if Chk_51.Checked then sJisa := sJisa + '''51'',';
               if Chk_45.Checked then sJisa := sJisa + '''45'',';
               if Chk_52.Checked then sJisa := sJisa + '''52'',';
               if Chk_34.Checked then sJisa := sJisa + '''34'',';
               if Chk_41.Checked then sJisa := sJisa + '''41'',';
               if Chk_Etc.Checked then sJisa := sJisa + '''62'',''80'',''80'',';
               
               if sJisa <> '' then
               begin
                  sWhere := 'WHERE A.cod_bjjs = ''' + GV_sJisa + '''';
                  sWhere := sWhere + ' AND B.cod_jisa in (' + copy(sJisa,1, Length(sJisa)-1) + ')';
               end
               else
               begin
                  showmessage('검진센터를 하나 이상 선택해주세요!');
                  exit;
               end;
            end
            else
               sWhere := 'WHERE A.cod_bjjs = ''' + GV_sJisa + '''';
         end;
      end;

      if rbDate.Checked then
      begin
         sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';
         sWhere := sWhere + ' AND A.gubn_part = ''' + UV_sGubn_part + '''';

         //if Cmb_gubn.Text = '분주번호' then
         if RB_bunju.Checked then
         begin
//------------------------------------------------------------------------------
            if (mskDate.Text = '20081025') then
            begin
               if GV_sJisa = '00' then
               begin
                  if Copy(cmbJisa.Text, 1, 2) = '15' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''1924'''
                  else if Copy(cmbJisa.Text, 1, 2) = '71' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end
               else if GV_sJisa = '15' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''1924''';
               end
               else if GV_sJisa = '71' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end;
            end;
//------------------------------------------------------------------------------
            if Trim(mskFrom.Text) <> '' then
               sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
            if Trim(mskTo.Text) <> '' then
               sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';

            sOrderBy := ' ORDER BY A.num_bunju ';
         end
         //else if Cmb_gubn.Text = '샘플번호' then
         else if RB_samp.Checked then
         begin
//------------------------------------------------------------------------------
            if (mskDate.Text = '20081025') then
            begin
               if GV_sJisa = '00' then
               begin
                  if Copy(cmbJisa.Text, 1, 2) = '15' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''1924'''
                  else if Copy(cmbJisa.Text, 1, 2) = '71' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end
               else if GV_sJisa = '15' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''1924''';
               end
               else if GV_sJisa = '71' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end;
            end;
            if Trim(MEdt_SampS.Text) <> '' then
               sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
            if Trim(MEdt_SampE.Text) <> '' then
               sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
//---------------------------------------------------------------------------
            sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju '
         end
         else
         begin
            if (mskDate.Text = '20081025') then
            begin
               if GV_sJisa = '00' then
               begin
                  if Copy(cmbJisa.Text, 1, 2) = '15' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''1924'''
                  else if Copy(cmbJisa.Text, 1, 2) = '71' then
                     sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end
               else if GV_sJisa = '15' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''1924''';
               end
               else if GV_sJisa = '71' then
               begin
                  sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
               end;
               if Trim(MEdt_SampS.Text) <> '' then
                  sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
               if Trim(MEdt_SampE.Text) <> '' then
                  sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';

            end
            else //20151110
            begin
               if Trim(mskFrom.Text) <> '' then
                  sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
               if Trim(mskTo.Text) <> '' then
                  sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';

               if Trim(MEdt_SampS.Text) <> '' then
                  sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
               if Trim(MEdt_SampE.Text) <> '' then
                  sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
            end;
//------------------------------------------------------------------------------
            sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju '
         end;
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' AND B.num_jumin = ''' + mskJumin.Text + '''';
         sWhere := sWhere + ' AND A.gubn_part = ''' + UV_sGubn_part + '''';

         sOrderBy := ' ORDER BY B.dat_gmgn ';
      end
      else if rbBarcode.Checked then
      begin
         sWhere := sWhere + ' AND B.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
         sWhere := sWhere + ' AND B.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';
         sWhere := sWhere + ' AND A.gubn_part = ''' + UV_sGubn_part + '''';

         sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju ';
      end;

      ParamByName('sWhere').Asmemo     :=  sWhere;
      ParamByName('sOrderBy').AsString :=  sOrderBy;
      Active := True;

      if qryGlkwa.RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I01', 'Q', 'N');
         if Cmb_gubn.Text = '혈액형검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
//[2012.03.07]혈액형 실제 검사자만 일괄정리 가능하게...
//------------------------------------------------------------------------------
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
{  //20140523 곽윤설
               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then
}              //혈액프로파일 모두 읽어오도록
               If (FieldByname('Cod_hul').Asstring <> '') then sCheck := '99';

               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';
               If (sCheck = '99') Or (FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
//------------------------------------------------------------------------------
               If (Pos('H023',sHangmok) > 0) or (Pos('H024',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '01') OR (UV_sGubn_part = '02') OR (UV_sGubn_part = '05') OR (UV_sGubn_part = '07')then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  Inc(index);
               end;
               Next;
            end;
         end
         else if Cmb_gubn.Text = 'ESR검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
//[2012.03.07]혈액형 실제 검사자만 일괄정리 가능하게...
//------------------------------------------------------------------------------
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';

               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then
                  sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';
               If FieldByName('Num_Bunju').AsInteger > 4999 Then  sCheck := '99';
//------------------------------------------------------------------------------

               Hangmok_Set2;

               If (Pos('H025',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '01') OR (UV_sGubn_part = '02') OR (UV_sGubn_part = '04') OR (UV_sGubn_part = '05') OR (UV_sGubn_part = '07') then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    :='임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  if (sCheck = '99') then UV_vABO_chk[index] := 'Y'
                  else                    UV_vABO_chk[index] := 'N';

                  Inc(index);
               end;
               Next;
            end;
         end
         //---------------------------------------------------------------------- 20151109 곽윤설
         else if Cmb_gubn.Text = 'Y004검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('Y004',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
               //혈액프로파일 모두 읽어오도록
               If FieldByname('Cod_hul').Asstring <> '' then sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';

               If (sCheck = '99') then //Or (FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
               If Pos('Y004',sHangmok) > 0 Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '03') then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  Inc(index);
               end;
               Next;
            end;
         end
         else if Cmb_gubn.Text = '니코틴검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('U033',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
               //혈액프로파일 모두 읽어오도록
               If FieldByname('Cod_hul').Asstring <> '' then sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';

               If (sCheck = '99') then //Or (FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
               If Pos('U033',sHangmok) > 0 Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '03') then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  Inc(index);
               end;
               Next;
            end;
         end
         else if Cmb_gubn.Text = '뇨침사검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('U011',FieldByName('cod_chuga').AsString) > 0) or (pos('U012',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
               //혈액프로파일 모두 읽어오도록
               If FieldByname('Cod_hul').Asstring <> '' then sCheck := '99';
               If FieldByname('Cod_jangbi').Asstring <> '' then sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';

               If (sCheck = '99') then// Or (FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
               If (Pos('U011',sHangmok) > 0) or (Pos('U012',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '03')then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  Inc(index);
               end;
               Next;
            end;
         end
         //----------------------------------------------------------------------
         else if Cmb_gubn.Text = 'M2-PK검사자' then
         begin
            for i := 0 to qryGlkwa.RecordCount - 1 do
            Begin
               sHangmok := '';
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               if (pos('Y010',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
               //혈액프로파일 모두 읽어오도록
               If FieldByname('Cod_hul').Asstring <> '' then sCheck := '99';
               If FieldByname('Cod_jangbi').Asstring <> '' then sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';

               If (sCheck = '99') then// Or (FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
               If (Pos('Y010',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

                  UV_fGulkwa := '';
                  UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  UV_vDesc_glkwa[index] := UV_fGulkwa;

                  UV_vPDat_bunju[index]  := '';
                  UV_vPNum_bunju[index]  := '';
                  // 해당검진자에 대한 이전검사결과를 획득
                  if (UV_sGubn_part = '03')then
                  begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                     case RGrp_preGulkwa.ItemIndex of
                        0 : begin
                               qryinjouk.Active := False;
                               qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                               qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                               qryinjouk.Active := True;
                               if qryinjouk.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                        1 : begin
                               qryinjouk1.Active := False;

                               sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                          ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                               sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                               sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                               sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                               sOrderBy := 'order by G.dat_gmgn DESC';

                               qryinjouk1.SQL.Clear;
                               qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                               qryinjouk1.Active := True;
                               if qryinjouk1.RecordCount > 0 then
                               begin
                                  qryPrev.Active := False;
                                  qryPrev.Filter := '';
                                  qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                                  qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                                  qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                                  qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                                  qryPrev.Active := True;
                                  if qryPrev.RecordCount > 0 then
                                  begin
                                     UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                     UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                     UV_fGulkwa := '';
                                     UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                     UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                     UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                     UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                                  end;
                                  qryPrev.Active := False;
                               end
                               else
                                  UV_vDesc_Pglkwa[index] := '';
                               qryinjouk.Active := False;
                            end;
                     end;
//------------------------------------------------------------------------------
                  end;
                  UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
                  UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                  UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                  UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                  UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
                  UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
                  UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
                  UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
                  UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
                  UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
                  UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
                  UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';


                  qryGumgin_Check.Active := False;
                  qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                  qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryGumgin_Check.Active := True;
                  if qryGumgin_Check.RecordCount > 0 then
                  begin
                       if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                          UV_vGumgin_check[index]   :='생리중';

                  end;

                  Inc(index);
               end;
               Next;
            end;
         end
         //----------------------------------------------------------------------
         else
         begin
            UP_VarMemSet(RecordCount-1);
            for i := 0 to RecordCount - 1 do
            begin
//[2012.03.07]혈액형 실제 검사자만 일괄정리 가능하게...
//------------------------------------------------------------------------------
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';

               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then
                  sCheck := '99';

               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then sCheck := '99';

               If FieldByName('Num_Bunju').AsInteger > 4999 Then  sCheck := '99';
//------------------------------------------------------------------------------

               UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
               UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
               UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
               UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
               UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;

               UV_fGulkwa := '';
               UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

               UV_vDesc_glkwa[index] := UV_fGulkwa;

               UV_vPDat_bunju[index]  := '';
               UV_vPNum_bunju[index]  := '';
               // 해당검진자에 대한 이전검사결과를 획득
               if (UV_sGubn_part = '01') OR (UV_sGubn_part = '02')  OR (UV_sGubn_part = '04') OR (UV_sGubn_part = '05') OR (UV_sGubn_part = '07') then
               begin
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
                  case RGrp_preGulkwa.ItemIndex of
                     0 : begin
                            qryinjouk.Active := False;
                            qryinjouk.ParamByName('NUM_JUMIN').AsString := qryGlkwa.FieldByName('num_jumin').AsString;
                            qryinjouk.ParamByName('DAT_GMGN').AsString  := qryGlkwa.FieldByName('dat_gmgn').AsString;
                            qryinjouk.Active := True;
                            if qryinjouk.RecordCount > 0 then
                            begin
                               qryPrev.Active := False;
                               qryPrev.Filter := '';
                               qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                               qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                               qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                               qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                               qryPrev.Active := True;
                               if qryPrev.RecordCount > 0 then
                               begin
                                  UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                  UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                  UV_fGulkwa := '';
                                  UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                  UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                  UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                  UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                               end;
                               qryPrev.Active := False;
                            end
                            else
                               UV_vDesc_Pglkwa[index] := '';
                            qryinjouk.Active := False;
                         end;
                     1 : begin
                            qryinjouk1.Active := False;

                            sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                                       ' FROM gumgin G with(nolock) INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
                            sWhere  := ' WHERE G.num_jumin = ''' + qryGlkwa.FieldByName('num_jumin').AsString + '''';
                            sWhere  := sWhere + ' AND G.dat_gmgn < ''' + qryGlkwa.FieldByName('dat_gmgn').AsString + '''';
                            sWhere  := sWhere + ' AND B.gubn_part = ''' + UV_sGubn_part + '''';
                            sOrderBy := 'order by G.dat_gmgn DESC';

                            qryinjouk1.SQL.Clear;
                            qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
                            qryinjouk1.Active := True;
                            if qryinjouk1.RecordCount > 0 then
                            begin
                               qryPrev.Active := False;
                               qryPrev.Filter := '';
                               qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                               qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                               qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                               qryPrev.ParamByName('GUBN_PART').AsString := UV_sGubn_part;
                               qryPrev.Active := True;
                               if qryPrev.RecordCount > 0 then
                               begin
                                  UV_vPDat_bunju[index]  := qryPrev.FieldByName('dat_bunju').AsString;
                                  UV_vPNum_bunju[index]  := qryPrev.FieldByName('num_bunju').AsString;

                                  UV_fGulkwa := '';
                                  UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                                  UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                                  UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                                  UV_vDesc_Pglkwa[index] := UV_fGulkwa;
                               end;
                               qryPrev.Active := False;
                            end
                            else
                               UV_vDesc_Pglkwa[index] := '';
                            qryinjouk.Active := False;
                         end;
                  end;
//------------------------------------------------------------------------------
               end;

               UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
               UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
               UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
               UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
               UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
               UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
               UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
               UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
               UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vGubn_tkgm[index]  := FieldByName('GUBN_tkgm').AsString;
               UV_vGubn_nosin[index] := FieldByName('GUBN_NOSIN').AsString;
               UV_vGubn_nsyh[index]  := FieldByName('GUBN_NSYH').AsString;
               UV_vGubn_bogun[index] := FieldByName('GUBN_BOGUN').AsString;
               UV_vGubn_bgyh[index]  := FieldByName('GUBN_BGYH').AsString;
               UV_vGubn_adult[index] := FieldByName('GUBN_ADULT').AsString;
               UV_vGubn_adyh[index]  := FieldByName('GUBN_ADYH').AsString;
               UV_vGubn_agab[index]  := FieldByName('GUBN_AGAB').AsString;
               UV_vGubn_agyh[index]  := FieldByName('GUBN_AGYH').AsString;
               UV_vGubn_gong[index]  := FieldByName('GUBN_GONG').AsString;
               UV_vGubn_goyh[index]  := FieldByName('GUBN_GOYH').AsString;
               UV_vGubn_hulgum[index]:= FieldByName('Gubn_hulgum').AsString;
               UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
               UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;

               if FieldByName('gubn_preg').AsString ='Y' then
                  UV_vGumgin_check[index]    := '임신중'
               else if FieldByName('gubn_preg').AsString ='P' then
                  UV_vGumgin_check[index]   := '임신가능성';

               qryGumgin_Check.Active := False;
               qryGumgin_Check.ParamByName('cod_jisa').AsString  := FieldByName('cod_jisa').AsString;
               qryGumgin_Check.ParamByName('NUM_JUMIN').AsString := FieldByName('NUM_JUMIN').AsString;
               qryGumgin_Check.ParamByName('dat_gmgn').AsString  := FieldByName('dat_gmgn').AsString;
               qryGumgin_Check.Active := True;
               if qryGumgin_Check.RecordCount > 0 then
               begin
                    if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                       UV_vGumgin_check[index]   :='생리중';

               end;

               if (sCheck = '99') then UV_vABO_chk[index] := 'Y'
               else                    UV_vABO_chk[index] := 'N';

               Next;
               Inc(index);
            end;
         end;
         // Table과 Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   //여기

   // Query Mode & Msg

   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD1I01.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // 취소시
   // 1.입력이면 화면 Clear
   // 2.수정이면 Re-query
   if UV_iStatus = GC_Insert then
      GP_FieldClear(pnlDetail);

   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD1I01.UP_Change(Sender: TObject);
var index, iPos : Integer;
    sCode : String;
    e1, e2, e3, e5, e6, e22, e23, e25, e26, e28, e41, e42, e45, e48, e74,e83 : Extended;
    eRslt : Extended;
begin
   inherited;
   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   index  := StrToInt(Copy(TEdit(Sender).Name, 9, 2));

   // 값을 Memory에 할당
   UV_sValue[index + UV_iPlus] := TEdit(Sender).Text;

   iPos := index + UV_iPlus;

   if (Trim(UV_sValue[iPos]) <> '') and (copy(Trim(UV_sValue[65]),1,1) <> '<') then  // 20140507 곽윤설
   begin
      if UV_sGubn[iPos] = '1' then
      begin
         if (GF_StrToFloat(UV_sValue[iPos]) < GF_StrToFloat(FloatToStr(UV_fLow[iPos]))) or (GF_StrToFloat(UV_sValue[iPos]) > GF_StrToFloat(FloatToStr(UV_fHigh[iPos]))) then
            TEdit(Sender).Color     := $00FACDF4
         else
            TEdit(Sender).Color     := clWindow;
      end
      else if UV_sGubn[iPos] = '2' then
      begin
         if UV_sValue[iPos] = UV_sUnit[iPos] then
            TEdit(Sender).Color     := clWindow
         else
            TEdit(Sender).Color     := $00FACDF4;
      end;
   end;

   eRslt := 0;
   sCode := UV_sCode[index + UV_iPlus];
   // 계산공식에 의해서 결과값을 구한다.

   // 1. 생화학
   if (sCode = 'C001') or (sCode = 'C002') or
      (sCode = 'C005') or (sCode = 'C006') or
      (sCode = 'C022') or (sCode = 'C023') or
      (sCode = 'C025') or (sCode = 'C026') or
      (sCode = 'C028') or (sCode = 'C029') or
      (sCode = 'C041') or (sCode = 'C042') or
      (sCode = 'C045') or (sCode = 'C048') or
      (sCode = 'C074') or (sCode = 'C083') then
   begin
      e1  := UF_GetValue('C001');
      e2  := UF_GetValue('C002');
      e3  := 0;
      e5  := UF_GetValue('C005');
      e6  := UF_GetValue('C006');
      e22 := UF_GetValue('C022');
      e23 := UF_GetValue('C023');
      e25 := UF_GetValue('C025');
      e26 := UF_GetValue('C026');
      e28 := UF_GetValue('C028');
      e41 := UF_GetValue('C041');
      e42 := UF_GetValue('C042');
      e45 := UF_GetValue('C045');
      e48 := UF_GetValue('C048');
      e74 := UF_GetValue('C074');
      e83 := UF_GetValue('C083');

      if (sCode = 'C001') or (sCode = 'C002') then
      begin
         // C003 = C001 - C002
         if (e1 <> 0) and (e2 <> 0) then
         begin
            eRslt := 0;
            eRslt := e1 - e2;
            e3 := eRslt ;
            UP_SetValue('C003', eRslt);
         end;
         // C004 = C002 / (C001 - C002)
         if (e3 <> 0) and (e2 <> 0) then
         begin
            eRslt := 0;
            eRslt := (e2/e3) * 10;
            eRslt := trunc(eRslt + 0.5);
            eRslt := eRslt / 10;
            UP_SetValue('C004', eRslt);
         end;
      end;
      if (sCode = 'C005') or (sCode = 'C006') then
      begin
         // C007 = C005 - C006
         if (e5 <> 0) and (e6 <> 0) then
         begin
            eRslt := 0;
            eRslt := e5 - e6;
            UP_SetValue('C007', eRslt);
         end;
      end;
      if (sCode = 'C022') or (sCode = 'C023') then
      begin
         // C024 = C023 / C022
         if (e22 <> 0) and (e23 <> 0) then
         begin
            eRslt := 0;
            eRslt := e23 / e22;
            UP_SetValue('C024', eRslt);
         end;
      end;

      if (sCode = 'C029') then
      begin
         // C029 = (C026 / C025) * 100
         if (e25 <> 0) and (e26 <> 0) then
         begin
            eRslt := 0;
            eRslt := (e26 / e25) * 100;
            UP_SetValue('C029', eRslt);
         end;
         if (eRslt < GF_StrToFloat(FloatToStr(UV_fLow[iPos]))) or (eRslt > GF_StrToFloat(FloatToStr(UV_fHigh[iPos]))) then
            TEdit(Sender).Color := $00FACDF4
         else
            TEdit(Sender).Color := clWindow;
      end;

      if sCode = 'C028' then
      begin
         // C027 = C025 / ((C028/5) + C026)
         //C025(총콜레스테롤) - C026 (HDL콜레스테롤) - [C028 (중성지방)/5]
         if (e25 <> 0) and (e26 <> 0) and (e25 <> 0) then
         begin
            eRslt := 0;
            eRslt := e25 - e26 - (e28/5);
            eRslt := eRslt + 0.5;
            eRslt := trunc(eRslt);

            UP_SetValue('C027', eRslt);
         end;
      end;

      if (sCode = 'C041') or (sCode = 'C042') then
      begin
         // C043 = C041 / C042
         if (e41 <> 0) and (e42 <> 0) then
         begin
            eRslt := 0;
            eRslt := e41 / e42;
            UP_SetValue('C043', eRslt);
         end;
      end;
      if (sCode = 'C045') or (sCode = 'C048') then
      begin
         if (e45 <> 0) and (e48 <> 0) then
         begin
            // C046 = C045 + C048
            eRslt := 0;
            eRslt := e45 + e48;
            UP_SetValue('C046', eRslt);

            eRslt := (e45 / (e45+e48)) * 100;
            UP_SetValue('C047', eRslt);
         end;
      end;
      //2017.08.25 수지
      if (sCode = 'C074') then
      begin
        eRslt := e74;
        UP_SetValue('C027', eRslt);
      end;

      if (e42 <> 0) then
      begin
         if (Copy(mskNUM_JUMIN.Text,7,1) = '1') or
            (Copy(mskNUM_JUMIN.Text,7,1) = '3') or
            (Copy(mskNUM_JUMIN.Text,7,1) = '5') or
            (Copy(mskNUM_JUMIN.Text,7,1) = '7') then
         begin
            if (GV_sToday >= '20180101') then eRslt := 175 * Power(e42,-1.154) * Power(iAge, -0.203)
            else eRslt := 186 * Power(e42,-1.154) * Power(iAge, -0.203);
         end
         else if (Copy(mskNUM_JUMIN.Text,7,1) = '2') or
                 (Copy(mskNUM_JUMIN.Text,7,1) = '4') or
                 (Copy(mskNUM_JUMIN.Text,7,1) = '6') or
                 (Copy(mskNUM_JUMIN.Text,7,1) = '8') then
         begin
            eRslt := 0;
            eRslt := 175 * Power(e42,-1.154) * Power(iAge, -0.203) * 0.742 ;
         end;
         UP_SetValue('C087', eRslt);
      end;

      if sCode = 'C083' then   //곽윤설 20140425
      begin
         if (e83 > 0) and (e83 < 3.0) then
         begin
           showmessage('HbA1C : 3.0이상 입력하세요.');
           grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);
           UP_MoveHm(btnNextHm);
         end;
      end;

   end;

   // 2. 혈액학
   if (sCode = 'H001') or(sCode = 'H002') or (sCode = 'H003') then
   begin
      e1 := UF_GetValue('H001');
      e2 := UF_GetValue('H002');
      e3 := UF_GetValue('H003');

      if sCode = 'H001' then
      begin
         // H004 = (H003 * 10) / H001
         if e1 <> 0 then eRslt := (e3 * 10) / e1;
         UP_SetValue('H004', eRslt);

         // H005 = (H002 * 10) / H001
         if e1 <> 0 then eRslt := (e2 * 10) / e1;
         UP_SetValue('H005', eRslt);
      end
      else if sCode = 'H002' then
      begin
         // H005 = (H002 * 10) / H001
         if e1 <> 0 then eRslt := (e2 * 10) / e1;
         UP_SetValue('H005', eRslt);

         // H006 = (H002 * 10) / H003
         if e3 <> 0 then eRslt := (e2 * 10) / e3;
         UP_SetValue('H006', eRslt);
      end
      else if sCode = 'H003' then
      begin
         // H004 = (H003 * 10) / H001
         if e1 <> 0 then eRslt := (e3 * 10) / e1;
         UP_SetValue('H004', eRslt);

         // H006 = (H002 * 100) / H003
         if e3 <> 0 then eRslt := (e2 * 100) / e3;
         UP_SetValue('H006', eRslt);
      end;
   END;
   // 4. RIA
   if (sCode = 'T023') OR (sCode = 'T011') then
   begin
      e1 := UF_GetValue('T023');
      e2 := UF_GetValue('T011');
      e3  := 0;

      if (sCode = 'T023') OR (sCode = 'T011') then
      begin
         // T058 = T023 / T011
         if (e1 <> 0) and (e2 <> 0) then
         begin
            eRslt := 0;
            eRslt := e1 / e2;
            e3 := eRslt ;
            UP_SetValue('T058', eRslt);
         end;
      end;
   end;
   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I01.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, iRet, i, iTemp : Integer;
    vCod_chuga : Variant;
    sHmCode, sCode : String;
    sSelect, sWhere, sOrderby : string;

begin
   inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 체크...
   UV_bSE42_Check := False;

   // Grid의 Row가 변경되면 Detail에 자료를 할당
   mskNUM_JUMIN.Text := UV_vNum_jumin[NewRow-1];
   edtDESC_NAME.Text := UV_vDesc_name[NewRow-1];
   mskDAT_GMGN.Text  := UV_vDat_gmgn[NewRow-1];
   mskDAT_BUNJU.Text := UV_vDat_bunju[NewRow-1];
   edtNUM_BUNJU.Text := UV_vNum_bunju[NewRow-1];

   mskPDAT_BUNJU.Text := UV_vPDat_bunju[NewRow-1];
   edtPNUM_BUNJU.Text := UV_vPNum_bunju[NewRow-1];

   Edt_jisa.Text     := UV_vCod_jisa[NewRow-1];

   UV_sHulgum        := UV_vGubn_hulgum[NewRow-1];
   UV_sCod_jangbi    := UV_vCod_jangbi[NewRow-1];
   UV_sCod_hul       := UV_vCod_hul[NewRow-1];

   EdtDesc_dept.Text := UV_vDesc_dept[NewRow-1];
   EdtNum_samp.Text  := UV_vNum_Sample[NewRow-1];
   edt_gumgin_Check.Text   := UV_vGumgin_check[NewRow-1];
   qry_Check_CRM.Active := False;
   qry_Check_CRM.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qry_Check_CRM.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qry_Check_CRM.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qry_Check_CRM.Active := True;
   if qry_Check_CRM.FieldByName('ysno_crm').AsString = 'Y' then chk_CRM.checked := true
   else chk_CRM.checked := false;

//[2012.003.15]검체불량관리 추가..
//------------------------------------------------------------------------------
   ckGUBN_JUNGDO1.Checked := False;
   ckGUBN_JUNGDO2.Checked := False;
   ckGUBN_JUNGDO3.Checked := False;
   ckGUBN_JUNGDO4.Checked := False;
   ckGUBN_JUNGDO5.Checked := False;
   ckGUBN_JUNGDO6.Checked := False;
   ckGUBN_JUNGDO8.Checked := False;
   ckGUBN_JUNGDO9.Checked := False;
   mskDAT_SOLVE1.Text := '';
   mskDAT_SOLVE2.Text := '';
   mskDAT_SOLVE3.Text := '';
   mskDAT_SOLVE4.Text := '';
   mskDAT_SOLVE5.Text := '';
   mskDAT_SOLVE6.Text := '';
   mskDAT_SOLVE8.Text := '';
   mskDAT_SOLVE9.Text := '';
   JD_sayu.Text := '' ;

   qrykicho.Active := False;
   qrykicho.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qrykicho.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qrykicho.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qrykicho.Active := True;
   if qrykicho.RecordCount > 0 then
   begin
        Edt_Sinjang.text:= qrykicho.FieldByname('Sinjang').Asstring;
        Edt_chejung.Text:= qrykicho.FieldByname('chejung').Asstring;
   end;
   Edt_munjin.Text:='';
   qry_munjin2010.Active := False;
   qry_munjin2010.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qry_munjin2010.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qry_munjin2010.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qry_munjin2010.Active := True;

   if qry_munjin2010.RecordCount > 0 then
   begin
        if qry_munjin2010.FieldByname('mun3').Asstring='1' then
           Edt_munjin.Text := 'B형간염항원 보유자';
        if qry_munjin2010.FieldByname('CMun7_Yn').Asstring='1' then
           begin
           if qry_munjin2010.FieldByname('CMun7').Asstring='21111' Then
              Edt_munjin.Text := Edt_munjin.Text+' ,B형 간염보유자 '
           else if qry_munjin2010.FieldByname('CMun7').Asstring='12111' Then
              Edt_munjin.Text := Edt_munjin.Text+' ,만성 B형 간염'
           else if qry_munjin2010.FieldByname('CMun7').Asstring='11211' Then 
              Edt_munjin.Text := Edt_munjin.Text+' ,만성 C형 간염';
           end;
   end;

   qryGum_bul.Active := False;
   qryGum_bul.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qryGum_bul.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qryGum_bul.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qryGum_bul.Active := True;
   if qryGum_bul.RecordCount > 0 then
   begin
      while not qryGum_bul.Eof  do
      begin
         if qryGum_bul.FieldByName('gubn_bul').AsString = '01' then
         begin
            ckGUBN_JUNGDO1.Checked := True;
            mskDAT_SOLVE1.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '02' then
         begin
            ckGUBN_JUNGDO2.Checked := True;
            mskDAT_SOLVE2.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '03' then
         begin
            ckGUBN_JUNGDO3.Checked := True;
            mskDAT_SOLVE3.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '04' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '05' then
         begin
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '06' then
         begin
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '07' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qryGum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qryGum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '08' then
         begin
            ckGUBN_JUNGDO8.Checked := True;
            mskDAT_SOLVE8.Text := qryGum_bul.FieldByName('dat_solve').AsString;
            JD_sayu.Text :=qryGum_bul.FieldByName('JD_sayu').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '09' then
         begin
            ckGUBN_JUNGDO9.Checked := True;
            mskDAT_SOLVE9.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;

         qryGum_bul.Next;
      end;
   end;
//------------------------------------------------------------------------------

   with qryDanche do
   begin
      Active := False;
      ParamByName('COD_dc').AsString := UV_vCod_dc[NewRow-1];
      Active := True;

      if RecordCount > 0 then
         EdtDesc_dc.Text := FieldByName('desc_dc').AsString;

      Active := False;
   end;

   // 검사코드를 추출
   index := 0;
   UV_bPerError := False;
   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if UV_vCod_hul[NewRow-1] <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := UV_vCod_hul[NewRow-1];
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
      if UV_vCod_cancer[NewRow-1] <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := UV_vCod_cancer[NewRow-1];
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
      if UV_vCod_jangbi[NewRow-1] <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := UV_vCod_jangbi[NewRow-1];
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
   end;

   // 3. 추가코드에 대한 검사항목 추출
   iRet := GF_MulToSingle(UV_vCod_chuga[NewRow-1], 4, vCod_chuga);
   VarArrayRedim(UV_vCod_hm, index+iRet);
   for i := 0 to iRet-1 do
   begin
      UV_vCod_hm[index] := vCod_chuga[i];
      Inc(index);
   end;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if UV_vGubn_nosin[NewRow-1] = '1' then
      sHmCode := UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '1', UV_vGubn_nsyh[NewRow-1])
   else if UV_vGubn_adult[NewRow-1] = '1' then
      sHmCode := UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '4', UV_vGubn_adyh[NewRow-1]);

   if UV_vGubn_life[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '7', UV_vGubn_lfyh[NewRow-1]);

   if UV_vGubn_bogun[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '3',UV_vGubn_bgyh[NewRow-1])
   else if UV_vGubn_bogun[NewRow-1] = '2' then
      sHmCode := sHmCode + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '3', UV_vGubn_bgyh[NewRow-1]);

   if UV_vGubn_agab[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '5',UV_vGubn_agyh[NewRow-1])
   else if UV_vGubn_agab[NewRow-1] = '2' then
      sHmCode := sHmCode + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '5', UV_vGubn_agyh[NewRow-1]);

   if (UV_vGubn_tkgm[NewRow-1] = '1') or (UV_vGubn_tkgm[NewRow-1] = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
      qryTkgum.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryTkgum.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
//[2012.03.15]특검항목 자유롭게..
//------------------------------------------------------------------------------
         if Trim(qryTkgum.FieldByName('cod_prf').AsString) <> '' then
         begin
            qryProfileG.Active := False;

            sSelect := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';
            for iTemp := 1 to Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 do
            begin
               if iTemp = Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 then
                  sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''''
               else
                  sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''' or ';
            end;
            sOrderby := ' ) and B.cod_pf <> ''SE81'' GROUP BY B.cod_hm ';

            qryProfileG.SQL.Clear;
            qryProfileG.SQL.Add(sSelect + sWhere + sOrderby);
            qryProfileG.Active := True;

            if qryProfileG.RecordCount > 0 then
            begin
               while not qryProfileG.Eof do
               begin
                  //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                  //----------------------------------------------------------
                  if qryProfileG.FieldByName('cod_hm').AsString = 'SE42' then
                  begin
                     if (pos('TC41', qryTkgum.FieldByName('cod_prf').AsString)   > 0) and
                        (pos('JJXE', qryTkgum.FieldByName('cod_chuga').AsString) > 0) then
                     begin
                        UV_bSE42_Check   := True;
                        Lbl_JJXE.Visible := True;
                     end
                     else
                     begin
                        UV_bSE42_Check   := False;
                        Lbl_JJXE.Visible := False;
                     end;
                  end;

                  sHmCode := sHmCode + qryProfileG.FieldByName('cod_hm').AsString;
                  UV_vTCod_hm[NewRow-1] := UV_vTCod_hm[NewRow-1] + qryProfileG.FieldByName('cod_hm').AsString; //20151126

                  qryProfileG.Next;
               end;
            end;
            qryProfileG.Active := False;
         end;
//------------------------------------------------------------------------------
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
         UV_vTCod_hm[NewRow-1] := UV_vTCod_hm[NewRow-1] + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      VarArrayRedim(UV_vCod_hm, index+iRet);
      for i := 0 to iRet-1 do
      begin
         UV_vCod_hm[index] := vCod_chuga[i];
         Inc(index);
      end;
   end;

   //09.06.27 철. 호모시스테임 검사있을시 표시
   Lbl_C078.Visible := False;
   for i := 0 to index-1 do
   begin
      if UV_vCod_hm[i] = 'C078' then
      begin
         Lbl_C078.Caption := '호모시스테인 검사있음';
         Lbl_C078.Visible := True;
      end;
   end;

   // Array Clear
   UP_ArrayClear;

   // 검사항목을 화면에 배치
   UP_HangmokLocate(UC_First);

   // Grid Focus
   grdMaster.SetFocus;

   if (Cmb_gubn.Text = '혈액형검사자') and (edtValue23.Enabled = True) then edtValue23.SetFocus;
   if (Cmb_gubn.Text = 'ESR검사자')    and (edtValue25.Enabled = True) then edtValue25.SetFocus;

   if (UV_vABO_chk[NewRow-1] = 'Y') and (edtValue23.Color <> clBtnFace) then pnlNum23.Color := clWhite
   else                                                                      pnlNum23.Color := $00FFFF80;

   if (UV_vABO_chk[NewRow-1] = 'Y') and (edtValue24.Color <> clBtnFace) then pnlNum24.Color := clWhite
   else                                                                      pnlNum24.Color := $00FFFF80;


   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;

procedure TfrmLD1I01.btnPInsertClick(Sender: TObject);
begin
   inherited;

   // 작업절차
   // 1. 사용자가 작업중이라면 저장이 입력상태
   // 2. 사용자가 작업중이아니라면 바로 입력상태
   if UV_bEdit then
   begin
      if UF_Save then
         btnInsert.Click;
   end
   else
      btnInsert.Click;
end;

procedure TfrmLD1I01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   // Button Click시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = btnJumin then
   begin
      if GF_InjoukClick(Self, sCode, sName) then
      begin
         mskJumin.Text := sCode;
         edtName.Text  := sName;
      end;   
   end
   else if Sender = btnDate then GF_CalendarClick(mskDate)
   else if Sender = CheckBox then
   begin
      if CheckBox.Checked then
      begin
         CheckBox.Caption := 'UnCheck All';
         Chk_15.Checked := True; Chk_12.Checked := True;
         Chk_11.Checked := True; Chk_43.Checked := True;
         Chk_71.Checked := True; Chk_61.Checked := True;
         Chk_51.Checked := True; Chk_45.Checked := True;
         Chk_52.Checked := True; Chk_34.Checked := True;
         Chk_41.Checked := True; Chk_Etc.Checked := True;
      end
      else
      begin
         CheckBox.Caption := 'Check All';
         Chk_15.Checked := False; Chk_12.Checked  := False;
         Chk_11.Checked := False; Chk_43.Checked  := False;
         Chk_71.Checked := False; Chk_61.Checked  := False;
         Chk_51.Checked := False; Chk_45.Checked  := False;
         Chk_52.Checked := False; Chk_34.Checked  := False;
         Chk_41.Checked := False; Chk_Etc.Checked := False;
      end;
   end;
end;

procedure TfrmLD1I01.UP_MoveHm(Sender: TObject);
begin
   inherited;

   if      Sender = btnPrevHm then UP_HangmokLocate(UC_First)
   else if Sender = btnNextHm then
   begin
      //첫페이지에는 다음페이지 버튼으로.. 다음페이지에선 저장버튼으로 가도록 하기위해.. (by.SSong)
      if (UV_iPage = 1) and (Cmb_gubn.Text <> '혈액형검사자') and (Cmb_gubn.Text <> 'ESR검사자') then
      begin
         // 다음페이지로 넘어갔을 경우 첫번째페이지 백분율 계산 체크를 위해..
         if (edtValue12.Text <> '') AND (UV_sGubn_part = '02') then
         begin
            if GF_StrToFloat(edtValue12.Text) + GF_StrToFloat(edtValue13.Text) + GF_StrToFloat(edtValue14.Text) + GF_StrToFloat(edtValue15.Text) +
               GF_StrToFloat(edtValue16.Text) + GF_StrToFloat(edtValue17.Text) + GF_StrToFloat(edtValue18.Text) + GF_StrToFloat(edtValue19.Text) +
               GF_StrToFloat(edtValue20.Text) + GF_StrToFloat(edtValue21.Text) + GF_StrToFloat(edtValue22.Text) <> 100 then
               UV_bPerError := True;
         end;

         UP_HangmokLocate(UC_Last);
      end
      else
         btnPSave.SetFocus;
   end;
end;
procedure TfrmLD1I01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD1I01.UP_Exit(Sender: TObject);
begin
   inherited;

   if Sender = mskJumin then
   begin
      // 주민번호의 입력오류일 경우 처리
      if Length(mskJumin.Text) < 13 then exit;

      with qryGmgn do
      begin
         Active := False;
         ParamByName('NUM_JUMIN').AsString := mskJumin.Text;
         Active := True;

         if RecordCount > 0 then
            edtName.Text := FieldByName('DESC_NAME').AsString
         else
         begin
            GF_MsgBox('Warning', '해당 주민번호는 검진자가 아닙니다.');
            edtName.Text := '';
            mskJumin.SetFocus;
         end;
      end;
   end
   else if Sender = MEdt_Barcode then
   begin
      // 주민번호의 입력오류일 경우 처리
      if Length(MEdt_Barcode.Text) < 12 then exit;

      btnQueryClick(self);
   end;
end;

procedure TfrmLD1I01.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      mskDate.Color      := $00E6FFFE;
      RB_bunju.Enabled   := True; RB_samp.Enabled    := True;
      mskDate.Enabled    := True; btnDate.Enabled    := True;
      mskFrom.Enabled    := True; mskTo.Enabled      := True;
      MEdt_SampS.Enabled := True; MEdt_SampE.Enabled := True;
      Cmb_gubn.Enabled   := True;

      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;  btnJumin.Enabled := False;

      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled := False;

      mskDate.SetFocus;

      // 찾기 Enable
      edtFind.Enabled  := True;
   end
   else if Sender = rbJumin then
   begin
      mskDate.Color      := clWindow;
      RB_bunju.Enabled   := False;  RB_samp.Enabled    := False;
      mskDate.Enabled    := False;  btnDate.Enabled    := False;
      mskFrom.Enabled    := False;  mskTo.Enabled      := False;
      MEdt_SampS.Enabled := False;  MEdt_SampE.Enabled := False;
      Cmb_gubn.Enabled   := False;

      mskJumin.Color   := $00E6FFFE;
      mskJumin.Enabled := True;    btnJumin.Enabled := True;

      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled := False;

      mskJumin.SetFocus;

      // 찾기 Disable
      edtFind.Enabled  := False;
   end
   else if Sender = rbBarcode then
   begin
      mskDate.Color      := clWindow;
      RB_bunju.Enabled   := False;  RB_samp.Enabled    := False;
      mskDate.Enabled    := False;  btnDate.Enabled  := False;
      mskFrom.Enabled    := False;  mskTo.Enabled    := False;
      MEdt_SampS.Enabled := False;  MEdt_SampE.Enabled := False;
      Cmb_gubn.Enabled   := False;

      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;    btnJumin.Enabled := False;

      MEdt_Barcode.Color   := $00E6FFFE;
      MEdt_Barcode.Enabled := True;

      MEdt_Barcode.SetFocus;

      // 찾기 Disable
      edtFind.Enabled  := False;
   end;
end;

procedure TfrmLD1I01.edtFindExit(Sender: TObject);
begin
   inherited;

   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
//   if Cmb_gubn.Text = '분주번호' then GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0)
   if RB_bunju.Checked  then GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0)
   else                      GF_Find(grdMaster, edtFind.Text, 2, 1, 1, 0);
end;

procedure TfrmLD1I01.UP_KeyMove(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var i, iTemp : Integer;
    sName : String;
begin
   inherited;

   // 정해진 Key가 입력되었는지 Check
   if Shift <> [ssAlt] then exit;
   if (Key <> VK_UP) and (Key <> VK_Down) and
      (Key <> VK_Left) and (Key <> VK_Right) then exit;

   // 이동할 위치를 계산
   iTemp := StrToInt(Copy((Sender as TEdit).Name, 9, 2));

   if Key = VK_Up then
   begin
      if      iTemp = 1  then iTemp := 15
      else if iTemp = 16 then iTemp := 30
      else if iTemp = 31 then iTemp := 45
      else                    iTemp := iTemp - 1;
   end
   else if Key = VK_Down then
   begin
      if      iTemp = 15 then iTemp := 1
      else if iTemp = 30 then iTemp := 16
      else if iTemp = 45 then iTemp := 31
      else                    iTemp := iTemp + 1;
   end
   else if Key = VK_Left then
   begin
      if iTemp <= 15 then iTemp := iTemp + 30
      else                iTemp := iTemp - 15;
   end
   else if Key = VK_Right then
   begin
      if iTemp >= 31 then iTemp := iTemp - 30
      else                iTemp := iTemp + 15;
   end;

   // 해당위치로 Focus이동
   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if pnlDetail.Controls[i].ClassType = TEdit then
      begin
         sName := TEdit(pnlDetail.Controls[i]).Name;
         if Pos('edtValue' + IntToStr(iTemp), sName) > 0 then
         begin
            if TEdit(pnlDetail.Controls[i]).Enabled then
               TEdit(pnlDetail.Controls[i]).SetFocus;

            break;
         end;

      end;
   end;
end;

procedure TfrmLD1I01.mskNumExit(Sender: TObject);
var i, iTemp : Integer;
    sName : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if mskNum.Text = '' then exit;

   // 현재항목에 해당값이 존재하는지 Check
   if (pnlNum1.Caption = '1') and (mskNum.Text > '45') then
      UP_HangmokLocate(UC_Last)
   else if (pnlNum1.Caption = '46') and (mskNum.Text < '46') then
      UP_HangmokLocate(UC_First);

   // 위치를 추출
   iTemp := StrToInt(mskNum.Text);
   if iTemp > 45 then iTemp := iTemp - 45;

   // 해당위치로 Focus이동
   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if pnlDetail.Controls[i].ClassType = TEdit then
      begin
         sName := TEdit(pnlDetail.Controls[i]).Name;
         if ('edtValue' + IntToStr(iTemp)) = sName then
         begin
            if TEdit(pnlDetail.Controls[i]).Enabled then
               TEdit(pnlDetail.Controls[i]).SetFocus
            else
               mskNum.SetFocus;

            break;
         end;
      end;
   end;
end;
procedure TfrmLD1I01.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F3 then
   begin
      if edtFind.Enabled then edtFind.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum)
   else if Key = VK_F7 then UP_MoveHm(btnPrevHm)
   else if Key = VK_F8 then UP_MoveHm(btnNextHm)
   else if Key = VK_RCONTROL then UP_MoveHm(btnNextHm)
   else if Key = VK_F9 then UP_Default(btnDefault)
   else if Key = VK_F10 then edtValue5.setFocus
   else if Key = VK_F11 then edtValue38.setFocus;   
end;

procedure TfrmLD1I01.UP_MoveNum(Sender: TObject);
begin
   inherited;

   if Sender = btnPrevNum then
   begin
      if grdMaster.CurrentDataRow > 1 then
         grdMaster.CurrentDataRow := grdMaster.CurrentDataRow - 1
      else
         exit;
   end
   else if Sender = btnNextNum then
   begin
      if grdMaster.CurrentDataRow < grdMaster.Rows then
         grdMaster.CurrentDataRow := grdMaster.CurrentDataRow + 1
      else
         exit;
   end;

   grdMaster.TopRow := grdMaster.CurrentDataRow;
end;

procedure TfrmLD1I01.grdMasterKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if      Key = VK_LEFT  then UP_HangmokLocate(UC_First)
   else if Key = VK_RIGHT then UP_HangmokLocate(UC_Last);
end;

function  TfrmLD1I01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

function  TfrmLD1I01.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_id').AsString      := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD1I01.UP_EditDisplay(Sender : TObject);
begin
   if Sender.ClassType <> TEdit then exit;

   // 자료가 선택되었는지 Check
   if TEdit(Sender).Text = '~' then TEdit(Sender).Text := '음성'
   else if TEdit(Sender).Text = '!' then TEdit(Sender).Text := '양성'
   else if TEdit(Sender).Text = '@' then TEdit(Sender).Text := '약양성'
   else if TEdit(Sender).Text = '#' then TEdit(Sender).Text := '경계역'
   else if TEdit(Sender).Text = '$' then TEdit(Sender).Text := '보류'
   else if TEdit(Sender).Text = '%' then TEdit(Sender).Text := 'Rh+'
   else if TEdit(Sender).Text = '^' then TEdit(Sender).Text := 'Rh-'
   else if TEdit(Sender).Text = '&' then TEdit(Sender).Text := 'A'
   else if TEdit(Sender).Text = '*' then TEdit(Sender).Text := 'B'
   else if TEdit(Sender).Text = '(' then TEdit(Sender).Text := 'O'
   else if TEdit(Sender).Text = ')' then TEdit(Sender).Text := 'AB';

end;

procedure TfrmLD1I01.btnDefaultClick(Sender: TObject);
begin
   inherited;
   UP_Default(btnDefault);
end;

procedure TfrmLD1I01.UP_Default(Sender: TObject);
begin
   inherited;

end;

procedure TfrmLD1I01.cmbRelationChange(Sender: TObject);
begin
  inherited;

   case cmbRelation.ItemIndex of
      0 : begin
             // 생화학
             frmLD1I018 := TFrmLD1I018.Create(self);
             frmLD1I018.ShowModal;
          end;
      1 : begin
             // 혈액학
             frmLD1I011 := TFrmLD1I011.Create(self);
             frmLD1I011.ShowModal;
          end;
      2 : begin
             // URIN
             frmLD1I012 := TFrmLD1I012.Create(self);
             frmLD1I012.ShowModal;
          end;
      3 : begin
             // EIA
             frmLD1I013 := TFrmLD1I013.Create(self);
             frmLD1I013.ShowModal;
          end;
      4 : begin
             // 혈청학
             frmLD1I014 := TFrmLD1I014.Create(self);
             frmLD1I014.ShowModal;
          end;
      5 : begin
             // 기타혈액1
             frmLD1I015 := TFrmLD1I015.Create(self);
             frmLD1I015.ShowModal;
          end;
   end;

   cmbRelation.ItemIndex := -1;
end;

procedure TfrmLD1I01.edtValue1Change(Sender: TObject);
begin
  inherited;

   //해당하는 값을 넣어줌.
   if Copy(TEdit(Sender).Name, 1, 8)  = 'edtValue' then
      UP_EditDisplay(Sender)
   else if Copy(TEdit(Sender).Name, 1, 8)  = 'edtPValue' then
      UP_EditDisplay(Sender)
end;

procedure TfrmLD1I01.Cmb_gubnChange(Sender: TObject);
begin
  inherited;

  //if Cmb_gubn.ItemIndex = 0 then
  if RB_bunju.Checked then
  begin
     Label2.Caption := '분주번호(F3)';
     edtFind.Text := '';
     edtFind.MaxLength := 4;
     mskFrom.Enabled := True;
     mskTo.Enabled := True;
     MEdt_SampS.Enabled := False;
     MEdt_SampE.Enabled := False;

     Label20.Visible := False;
     Chk_15.Visible  := False;
     Chk_12.Visible  := False;
     Chk_11.Visible  := False;
     Chk_43.Visible  := False;
     Chk_71.Visible  := False;
     Chk_61.Visible  := False;
     Chk_51.Visible  := False;
     Chk_45.Visible  := False;
     Chk_52.Visible  := False;
     Chk_34.Visible  := False;
     Chk_41.Visible  := False;
     Chk_Etc.Visible := False;
     CheckBox.Visible := False;
  end
  //else if Cmb_gubn.ItemIndex = 1 then
  else if RB_samp.Checked then
  begin
     Label2.Caption := '샘플번호(F3)';
     edtFind.Text := '';
     edtFind.MaxLength := 6;
     mskFrom.Enabled := False;
     mskTo.Enabled := False;
     MEdt_SampS.Enabled := True;
     MEdt_SampE.Enabled := True;

     Label20.Visible := False;
     Chk_15.Visible  := False;
     Chk_12.Visible  := False;
     Chk_11.Visible  := False;
     Chk_43.Visible  := False;
     Chk_71.Visible  := False;
     Chk_61.Visible  := False;
     Chk_51.Visible  := False;
     Chk_45.Visible  := False;
     Chk_52.Visible  := False;
     Chk_34.Visible  := False;
     Chk_41.Visible  := False;
     Chk_Etc.Visible := False;
     CheckBox.Visible := False;
  end;

  if Cmb_gubn.Text = 'ESR검사자' then
  begin
     if (Panel1.Caption = '혈  액  학  결  과  등  록') AND (copy(GV_sUserId,1,2)='15') then
     begin
        Label20.Visible := True;
        Chk_15.Visible := True;
        Chk_12.Visible := True;
        Chk_11.Visible := True;
        Chk_43.Visible := True;
        Chk_71.Visible := True;
        Chk_61.Visible := True;
        Chk_51.Visible := True;
        Chk_45.Visible := True;
        Chk_52.Visible := True;
        Chk_34.Visible := True;
        Chk_41.Visible := True;
        Chk_Etc.Visible := True;
        CheckBox.Visible := True;
     end;
  end;
end;

Procedure TfrmLD1I01.Hangmok_Set;
Var I, J, K : Integer;
Begin
      Inherited;
//PRofile Reading
      With qryProfile do
         Begin
            Close;
//장비
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_jangbi').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//혈액
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Hul').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//Cancer
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Cancer').AsString;
            Open;
            For I := 1 To Recordcount Do
                Begin
                   sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                   Next;
                End;
             Close;
         End;
//추가코드
         sHangmok := sHangmok + qryGlkwa.FieldByName('Cod_chuga').AsString;
//특검
         If qryGlkwa.FieldByName('Gubn_tkgm').AsString <> '' Then
            With qryTkgum Do
            Begin
               Close;
               ParamByName('Cod_jisa').AsString  := qryGlkwa.FieldByname('Cod_jisa').AsString;
               ParamByName('Dat_gmgn').AsString  := qryGlkwa.FieldByname('DaT_gmgn').AsString;
               ParamByName('NuM_Jumin').AsString := qryGlkwa.FieldByname('NuM_Jumin').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  I := Length(FieldByName('Cod_prf').AsString);
                  J := Round(I / 4);
                  For K := 0 To J - 1 Do
                  Begin
                     With qryProfile do
//특검 Profile
                     Begin
                        Close;
                        ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4);
                        Open;
                        For I := 1 To Recordcount Do
                        Begin
                           //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                           //----------------------------------------------------------
                           if (Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4) = 'TC41') and
                              (pos('JJXE', qryTkgum.FieldByName('Cod_chuga').AsString) > 0) and
                              (FieldByName('Cod_hm').AsString = 'SE42') then
                           begin
                              // 혈중카복시(SE42) Skip...
                           end
                           else
                           begin
                              sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                           end;

                           Next;
                        End;
                        Close;
                     End;
                  End;
               End;
               sHangmok := sHangmok + qryTkgum.FieldByName('Cod_chuga').AsString;
               Close;
            End;
End;

Procedure TfrmLD1I01.Hangmok_Set2;
Var I, J, K : Integer;
Begin
      Inherited;
//PRofile Reading
      With qryProfile_H025 do
         Begin
            Close;
//장비
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_jangbi').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//혈액
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Hul').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//Cancer
            ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Cancer').AsString;
            Open;
            For I := 1 To Recordcount Do
                Begin
                   sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                   Next;
                End;
             Close;
         End;
//추가코드
         sHangmok := sHangmok + qryGlkwa.FieldByName('Cod_chuga').AsString;
//특검
         If qryGlkwa.FieldByName('Gubn_tkgm').AsString <> '' Then
            With qryTkgum Do
               Begin
                  Close;
                  ParamByName('Cod_jisa').AsString  := qryGlkwa.FieldByname('Cod_jisa').AsString;
                  ParamByName('Dat_gmgn').AsString  := qryGlkwa.FieldByname('DaT_gmgn').AsString;
                  ParamByName('NuM_Jumin').AsString := qryGlkwa.FieldByname('NuM_Jumin').AsString;
                  Open;
                  If RecordCount > 0 Then
                     Begin
                        I := Length(FieldByName('Cod_prf').AsString);
                        J := Round(I / 4);
                        For K := 0 To J - 1 Do
                           Begin
//특검 Profile
                              With qryProfile_H025 do
                                 Begin
                                    Close;
                                    ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4);
                                    Open;
                                    For I := 1 To Recordcount Do
                                    Begin
                                       //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                                       //---------------------------------------
                                       if (Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4) = 'TC41') and
                                          (pos('JJXE', qryTkgum.FieldByName('Cod_chuga').AsString) > 0) and
                                          (FieldByName('Cod_hm').AsString = 'SE42') then
                                       begin
                                          // 혈중카복시(SE42) Skip...
                                       end
                                       else
                                       begin
                                          sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                                       end;
                                       //=======================================
                                       Next;
                                    End;

                                    Close;
                                 End;
                           End;
                      End;
                   sHangmok := sHangmok + qryTkgum.FieldByName('Cod_chuga').AsString;
                   Close;
                End;
End;

procedure TfrmLD1I01.btnPrintClick(Sender: TObject);
begin
  inherited;
  if pos('LD1I02',frmLD1I01.Caption) > 0 then
  begin
     frmLD1P011 := TfrmLD1P011.Create(Self);
     frmLD1P011.QuickRep.Preview;
  end;
end;

end.
 