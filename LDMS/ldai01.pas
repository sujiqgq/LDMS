//==============================================================================
// 프로그램 설명 : 2007년도 조직실 결과등록(총센터)
// 시스템        : MDMS
// 수정일자      : 2007.09.06
// 수정자        : 유동구
// 수정내용      : 추가...
//==============================================================================
// 수정일자      : 2008.11.10
// 수정자        : 김승철
// 수정내용      : Class항목 추가관련 수정.
//==============================================================================
// 수정일자      : 2010.06.28
// 수정자        : 유동구
// 수정내용      : 객담결과 입력추가..
//==============================================================================
// 수정일자      : 2010.10.04
// 수정자        : 구수정
// 수정내용      : 소견코드입력시 자동판정 되도록..
//==============================================================================
// 수정일자      : 2011.05.03
// 수정자        : 김승철
// 수정내용      : B005추가
//==============================================================================
// 수정일자      : 2011.11.17
// 수정자        : 유동구
// 수정내용      : 소변세포병리검사 추가(P013,JJFG)
//==============================================================================
// 수정일자      : 2013.07.25
// 수정자        : 박신정
// 수정내용      : 판정중일 경우 결과 보이지 않기
//==============================================================================
// 수정일자      : 2014.04.11
// 수정자        : 곽윤설
// 수정내용      :  1) 검체종류 (B001, B005, B007, P010) 는 4. Other 로 기본설정
//                  2) 조직진단, 조직검사갯수 기본설정
//                     (B001 = 2. 위염, 1~3개 / B007 = 3. 저도선종 혹은 이형성, 1~3개)
//                    => 권한설정 ID (150007, 150681) 기본설정
//                  4) 소견코드 입력&저장 후 소견코드가 사라지지 않게 수정 -> 2014.05.10 백승현 추가 수정
//==============================================================================
// 수정일자      : 2014.05.09
// 수정자        : 곽윤설
// 수정내용      : P003, P010, P011, P012, P013 - 기본 검사자 '150674-감원연'으로 변경
//                 B   - 기본 검사자 '150729-정종우'로 변경(기존 '150582-김병권')
//==============================================================================
// 수정일자      : 2014.05.14
// 수정자        : 곽윤설
// 수정내용      : 비정형편평상피세포 콤보박스 (CmbSang_Form1_gubn) -> 1,2 이외 값 입력 안되도록
//==============================================================================
// 수정일자      : 2014.06.19
// 수정자        : 곽윤설
// 수정내용      : B009 추가 [병리-김원연]
//==============================================================================
// 수정일자      : 2014.07.15
// 수정자        : 곽윤설
// 수정내용      : P001,3,6,9,10,11,12,13  - 기본 검사자 '150674-감원연' ->> '150943-한영애'로 변경
//==============================================================================
// 수정일자      : 2014.09.30
// 수정자        : 곽윤설
// 수정내용      : B   - 기본 검사자 '150574-윤형빈'으로 변경(기존 '150729-정종우')
//==============================================================================
// 수정일자      : 2014.12.09
// 수정자        : 곽윤설
// 수정내용      : B012추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.01.21
// 수정자        : 곽윤설
// 수정내용      : 검사자 변경
// 참고사항      : 한의 재단병리팀1500011[병리팀-이영신]
//==============================================================================
// 수정일자      : 2015.03.04
// 수정자        : 곽윤설
// 수정내용      : 소견판정 저장값 끌어오기
// 참고사항      : [병리팀-윤형빈]
//==============================================================================
// 수정일자      : 2015.03.27
// 수정자        : 곽윤설
// 수정내용      : P012[판정관리] 아이템 변경 (종검,준종검 결과지)
// 참고사항      : [병리팀-윤형빈]
//==============================================================================
// 수정일자      : 2015.05.19
// 수정자        : 곽윤설
// 수정내용      : B001,B005,B007 이영신책임 검사 → 윤형빈선임 검사
// 참고사항      : [병리팀-박예진]
//==============================================================================
// 수정일자      : 2015.07.07
// 수정자        : 곽윤설
// 수정내용      : 검사자 변경 B001 / B005 / B007 / P009 :  151021 - 최재현
// 참고사항      : [병리팀-박예진]
//==============================================================================
// 수정일자      : 2015.08.01
// 수정자        : 곽윤설
// 수정내용      : UV_vCod_bjjs -> invisible, UV_vDoctor -> 신규추가
// 참고사항      : 한의 재단병리팀1500161 [병리팀 - 박예진 요청]
//==============================================================================
// 수정일자      : 2015.08.04
// 수정자        : 곽윤설
// 수정내용      : P003검사 관련 소견코드 수정 시 체크된 항목 모두 해제되도록
// 참고사항      : 한의 재단병리팀장1500028 [병리팀 - 이유경 요청]
//==============================================================================
// 수정일자      : 2015.09.18
// 수정자        : 곽윤설
// 수정내용      : 윤형빈 선임(150574) 퇴사로 인해 담당업무 최윤선 선임(151030)으로 변경
// 참고사항      : 박예진선임 요청
//==============================================================================
// 수정일자      : 2015.11.09
// 수정자        : 곽윤설
// 수정내용      : PO46 신규 소견코드 생성에 따른 프로그램 수정요청.
// 참고사항      : 한의 재단병리팀장1500048
//==============================================================================
// 수정일자      : 2016.02.29
// 수정자        : 박대용
// 수정내용      : 2016년 청구파일 양식 변경에 따라 위, 대장관련 변경
// 참고사항      :
//==============================================================================
// 수정일자      : 2016.03.17
// 수정자        : 박대용
// 수정내용      : ICMS상에(IC2P201) 자궁암 관련해서 자궁암 의사 싸인 나오지 않던분제 수정
//               : 서명저장시 ca_cancer2009 DB에 can5_doctor에 의사 코드 저장
// 참고사항      : 한의 본원의료정보팀1600133
//============================================================================== 
// 수정일자      : 2016.03.23
// 수정자        : 박대용
// 수정내용      : 위 수정사항에 대한 에러 수정 - 타 지사 의사서명에도 입력되는 것을 수정
//               : qryU_ca_cancer2009_Can5_15 - 15지사용 개별 쿼리 추가 생성
// 참고사항      : 한의 본원의료정보팀1600133
//==============================================================================
// 수정일자      : 2016.09.26
// 수정자        : 유동구
// 수정내용      : P012 - 판정값 변경(코드는 기존과 동일 멘트 변경)
// 참고사항      : 한의 재단특검행정파트1601278
//==============================================================================
// 수정일자      : 2017.4.24
// 수정자        : 박수지
// 수정내용      : 병리팀을 제외하고 조회는 판독 완료 한 것들만 보여지게 적용
// 참고사항      : 한의 재단조직병리파트1700079
//==============================================================================
// 수정일자      : 2017.07.20
// 수정자        : 박수지
// 수정내용      : 기존 병리 진단 등록 -> 조직병리, 세포병리로 분리하여 신규프로그램 적용
// 참고사항      : 한의 재단조직병리파트1700126
//==============================================================================
// 수정일자      : 2017.11.24
// 수정자        : 박수지
// 수정내용      : PO43소견 클래스, 노신판정 부분 새로 입력
// 참고사항      : 한의 재단조직병리파트1700056
//==============================================================================
unit LDAI01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, jpeg,
  QuickRpt, Qrctrls;

type
  TfrmLDAI01 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    pnlDetail: TPanel;
    rbDate: TRadioButton;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    Panel2: TPanel;
    pnlJisa: TPanel;
    Label2: TLabel;
    cmbJisa: TComboBox;
    cmbGUBN_JONGRU: TComboBox;
    Panel7: TPanel;
    mskDAT_LAST: TDateEdit;
    btnDAT_LAST: TSpeedButton;
    rbJumin: TRadioButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    btnDate: TSpeedButton;
    mskDate: TDateEdit;
    cmbGubn: TComboBox;
    Panel3: TPanel;
    mskDAT_RESULT: TDateEdit;
    btnDAT_RESULT: TSpeedButton;
    Panel4: TPanel;
    cmbCOD_SAWON: TComboBox;
    Panel5: TPanel;
    cmbCOD_DOCT: TComboBox;
    retDESC_KGBL: TRichEdit;
    cmbGUBN_CLASS: TComboBox;
    Label1: TLabel;
    ckMiSokun: TCheckBox;
    Panel6: TPanel;
    GroupBox1: TGroupBox;
    ckYSNO_VIRUS1: TCheckBox;
    ckYSNO_VIRUS2: TCheckBox;
    ckYSNO_VIRUS3: TCheckBox;
    ckYSNO_VIRUS4: TCheckBox;
    ckYSNO_VIRUS5: TCheckBox;
    Label3: TLabel;
    edtDESC_VIRUS: TEdit;
    GroupBox2: TGroupBox;
    ckYSNO_CLIN1: TCheckBox;
    ckYSNO_CLIN3: TCheckBox;
    ckYSNO_CLIN5: TCheckBox;
    ckYSNO_CLIN2: TCheckBox;
    ckYSNO_CLIN4: TCheckBox;
    ckYSNO_CLIN6: TCheckBox;
    ckYSNO_CLIN7: TCheckBox;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel11: TPanel;
    Label4: TLabel;
    edtDESC_JONGRU: TEdit;
    edtDESC_BANG: TEdit;
    edtDESC_BUWI: TEdit;
    edtCod_sokun: TEdit;
    retDesc_Sokun: TMemo;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    ckYsno_Virus7: TCheckBox;
    ckYsno_Virus6: TCheckBox;
    edtDesc_Gum1: TCheckBox;
    edtDesc_Gum2: TCheckBox;
    Label5: TLabel;
    edtGum_Remark: TEdit;
    Label6: TLabel;
    edtGum_Chk1: TCheckBox;
    edtGum_Chk2: TCheckBox;
    edtGum_Form1: TCheckBox;
    edtGum_Form2: TCheckBox;
    edtGum_Form3: TCheckBox;
    edtSang_Form1: TCheckBox;
    edtSang_Form2: TCheckBox;
    edtSang_Form4: TCheckBox;
    edtSang_Form3: TCheckBox;
    edtSun_Form1: TCheckBox;
    edtSun_Form2: TCheckBox;
    Label7: TLabel;
    edtSun_Form4: TEdit;
    edtSun_Form3: TCheckBox;
    qrySokun: TQuery;
    qryCell: TQuery;
    qryGmgn: TQuery;
    qryU_Cell: TQuery;
    qryD_Cell: TQuery;
    GroupBox7: TGroupBox;
    edtCod_Pan1: TRadioButton;
    edtCod_Pan3: TRadioButton;
    edtCod_Pan2: TRadioButton;
    edtCod_Pan5: TRadioButton;
    edtCod_Pan6: TRadioButton;
    edtDesc_Pan: TEdit;
    qryUser_priv: TQuery;
    qryCommon: TQuery;
    CmbSang_Form1_gubn: TComboBox;
    edtGum_Form3_Etc: TEdit;
    Chk_order: TCheckBox;
    btn_sign: TBitBtn;
    Qry_CK_Gulkwa: TQuery;
    Qry_Sign: TQuery;
    QryI_Sign: TQuery;
    Can_Advice: TMemo;
    Panel12: TPanel;
    qryCa_cancer2009: TQuery;
    qryI_Ca_Cancer2009_Can5_15: TQuery;
    Label8: TLabel;
    Label9: TLabel;
    qryI_Ca_Cancer2009_Can1: TQuery;
    qryU_ca_cancer2009_Can1: TQuery;
    qryI_Ca_Cancer2009_Can2: TQuery;
    qryU_ca_cancer2009_Can2: TQuery;
    ckYsno_Virus8: TCheckBox;
    cmbOrder: TComboBox;
    Label10: TLabel;
    Panel14: TPanel;
    edtGUBN_PAN: TEdit;
    qryCa_Request: TQuery;
    qryJangbi: TQuery;
    retDesc_Sokun1: TMemo;
    qryI_cellsokun: TQuery;
    qryU_cellsokun: TQuery;
    qryCell_sokun: TQuery;
    pnlHide: TPanel;
    Panel16: TPanel;
    Label12: TLabel;
    qryU_ca_cancer2009_Can5_15: TQuery;
    qryU_ca_cancer2009_Can5: TQuery;
    qryI_Ca_Cancer2009_Can5: TQuery;
    PnlHide4: TPanel;
    Label15: TLabel;
    Panel13: TPanel;
    qryFloor: TQuery;
    Label14: TLabel;
    qryHangmok: TQuery;
    Edt_cod_chuga: TMemo;
    Label16: TLabel;
    Panel15: TPanel;
    Label13: TLabel;
    Edt_floor: TEdit;
    Pnl_P012: TPanel;
    Label11: TLabel;
    Cmb_P012: TComboBox;
    Label22: TLabel;
    Label23: TLabel;
    GroupBox8: TGroupBox;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Lab_slide: TLabel;
    qry_slide: TQuery;
    qryU_slide: TQuery;
    qryI_slide: TQuery;
    msk_dat_rent: TMaskEdit;
    btn_rent: TRadioButton;
    btn_return: TRadioButton;
    msk_dat_return: TMaskEdit;
    Button1: TButton;
    Label20: TLabel;
    Edt_cmt: TMaskEdit;
    qry_cancel: TQuery;
    asdf: TLabel;
    memo1: TMemo;
    qryComplain: TQuery;
    Pnl_P013: TPanel;
    Label21: TLabel;
    Cmb_P013: TComboBox;
    qryCa_cancer2018: TQuery;
    qryU_ca_cancer2018_Can1: TQuery;
    qryU_ca_cancer2018_Can2: TQuery;
    qryU_ca_cancer2018_Can5: TQuery;
    qryU_ca_cancer2018_Can5_15: TQuery;
    canSun_form5: TComboBox;
    qryI_Ca_Cancer2018_Can1: TQuery;
    qryI_Ca_Cancer2018_Can2: TQuery;
    qryI_Ca_Cancer2018_Can5: TQuery;
    qryI_Ca_Cancer2018_Can5_15: TQuery;
    ckYsno_Virus10: TCheckBox;
    rbCell: TRadioButton;
    mskcell: TMaskEdit;
    edt_time_S: TEdit;
    edt_time_R: TEdit;
    Panel10: TPanel;
    cmbCOD_JINDAN_DOCT: TComboBox;
    chk_jindan: TCheckBox;
    rbRDate: TRadioButton;
    mskRDate: TDateEdit;
    btnRDate: TSpeedButton;
    Panel18: TPanel;
    Panel17: TPanel;
    edt_num1: TRadioButton;
    edt_num2: TRadioButton;
    edt_num3: TRadioButton;
    edt_kita: TEdit;
    btn_del_num: TBitBtn;
    edt_del_num: TEdit;
    qryU_DelNumber: TQuery;
    Label24: TLabel;
    qryI_DelNumber: TQuery;
    qry_DelNumber: TQuery;
    edtCod_Pan4: TRadioButton;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Exit(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure edtCod_sokunExit(Sender: TObject);
    procedure edtCod_Pan6Click(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure cmbRelationChange(Sender: TObject);
    procedure btn_signClick(Sender: TObject);
    procedure cmbGUBN_CLASSExit(Sender: TObject);
    procedure cmbGubnChange(Sender: TObject);
    procedure CmbSang_Form1_gubnKeyPress(Sender: TObject; var Key: Char);
    procedure btnPrintClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure rbRDateClick(Sender: TObject);
    procedure btn_del_numClick(Sender: TObject);
    procedure edt_num3Click(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,  UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vCod_cell, UV_vFloor, UV_vChuga_HM, UV_vRent, UV_vCmt, UV_vDat_Rent, UV_vDAT_Return,
    UV_vCod_hm, UV_vGubn_jongru, UV_vDesc_jongru, UV_vGubn_class, UV_vYsno_virus1,
    UV_vYsno_virus2, UV_vYsno_virus3, UV_vYsno_virus4, UV_vYsno_virus5, UV_vYsno_Virus6, UV_vYsno_Virus7, UV_vYsno_Virus8, UV_vYsno_Virus10,
    UV_vDesc_virus, UV_vYsno_clin1, UV_vYsno_clin2, UV_vYsno_clin3, UV_vYsno_clin4,
    UV_vYsno_clin5, UV_vYsno_clin6, UV_vYsno_clin7, UV_vDesc_bang, UV_vDesc_buwi,
    UV_vCod_sawon, UV_vCod_doct, UV_vCod_Jindan_doct, UV_vCheck_jindan, UV_vDat_last, UV_vDAT_TIME, UV_vDAT_TIME_R, UV_vDat_result, UV_vDesc_kgbl,
    UV_vDesc_sokun, UV_vDesc_sokun1, UV_vCod_jisa, UV_vDesc_Gum, UV_vGum_Chk, UV_vGum_Remark,
    UV_vGum_Form1, UV_vGum_Form2, UV_vGum_Form3, UV_vGum_Form3_Etc,
    UV_vSang_Form1, UV_vSang_Form1_Gubn, UV_vSang_Form2, UV_vSang_Form3, UV_vSang_Form4,
    UV_vSun_Form1, UV_vSun_Form2, UV_vSun_Form3, UV_vSun_Form4, UV_vSun_Form5, UV_vCod_Pan, UV_vDesc_Pan, UV_vYsno_sokun, UV_vCheck_Year,
    UV_vGubn_P012, UV_vGubn_pan, UV_vCod_sokun, UV_vDocto, UV_vEdt_delnum, UV_vdesc_kita, UV_vDoctor   : Variant;

    // Select문
    UV_sBasicSQL, UV_sBasicSQL_cs : String;

    //검진일
    UV_bGmgn :Boolean;

    // 작업 지사코드
    UV_sCod_jisa : String;
    UV_iCod_doct : Integer;
    SUBSOKUN, UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4 : String;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;

    function UF_B201(bData : Boolean) : String;
    function UF_B201_Jong1(bData, bData_H, bData_M, bData_L : Boolean) : String;
    function UF_B201_Jong4(bData, bData_H, bData_L : Boolean) : String;
    Function UF_Check(iTemp : Integer) : Integer;
    Function UF_Check01(iTemp : Integer) : String;
  public
    { Public declarations }
  end;

var
  frmLDAI01: TfrmLDAI01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     LD1I07F, LD1I071, LD1P13, LD1P132,  LDAI012;//, LD1P121;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLDAI01.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLDAI01.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '주민번호';
         Col[3].Heading   := '이  름';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taLeftJustify;
         Col[3].Color     := clWindow;
         Col[3].Alignment := taLeftJustify;
      end
      else if rbRDate.Checked then
      begin
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '주민번호';
         Col[3].Heading   := '이  름';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taLeftJustify;
         Col[3].Color     := clWindow;
         Col[3].Alignment := taLeftJustify;
      end
      else if  rbJumin.Checked then
      begin
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '검진일자';
         Col[3].Heading   := 'None';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taCenter;
         Col[3].Color     := clBtnFace;
      end
      else if rbCell.Checked then
      begin
         Col[1].Heading   := '지사';
         Col[2].Heading   := '검진일자';
         Col[3].Heading   := '성 명';
         Col[4].Heading   := '주민번호';
         Col[5].Heading   := '접수 No';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taCenter;
         Col[3].Color     := clBtnFace;
      end;

      if rbCell.Checked = false then Col[4].Heading := '접수 No';
      if rbCell.Checked = false then Col[5].Heading := '세포코드';
      Col[6].Heading := '항목코드';
      Col[7].Heading := '분주지사';
      Col[8].Heading := '임상의사'; //20150803 곽윤설
      Col[9].Heading := '판독유무';

      // Set Alignment
      Col[4].Alignment := taCenter;
      Col[5].Alignment := taCenter;
      Col[6].Alignment := taCenter;
      Col[7].Alignment := taCenter;
      Col[8].Alignment := taCenter;
      Col[9].Alignment := taCenter;
   end;
end;

procedure TfrmLDAI01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_jongru := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_jongru := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_class  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus1 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus2 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus3 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus4 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus5 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus6 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus7 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus8 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus10 := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_virus  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin1  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin2  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin3  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin4  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin5  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin6  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin7  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Gum    := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Chk     := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Remark  := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form1   := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form2   := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form3   := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form3_Etc   := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form1  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form1_Gubn  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form2  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form3  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form4  := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form1   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form2   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form3   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form4   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form5   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_bang   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sawon   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jindan_doct := VarArrayCreate([0, 0], varOleStr);
      UV_vCheck_jindan:= VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_last    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_time    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_time_R  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_result  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_kgbl   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun1 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Pan     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pan    := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_sokun  := VarArrayCreate([0, 0], varOleStr);
      UV_vCheck_Year  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_P012   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_Pan    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sokun   := VarArrayCreate([0, 0], varOleStr);
      UV_vDoctor      := VarArrayCreate([0, 0], varOleStr);
      UV_vFloor       := VarArrayCreate([0, 0], varOleStr);
      UV_vChuga_HM    := VarArrayCreate([0, 0], varOleStr);
      UV_vRent        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Rent    := VarArrayCreate([0, 0], varOleStr);
      UV_vDAT_Return  := VarArrayCreate([0, 0], varOleStr);
      UV_vCmt         := VarArrayCreate([0, 0], varOleStr);
      UV_vEdt_delnum  := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_kita   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vCod_cell,    iLength);
      VarArrayRedim(UV_vCod_hm,      iLength);
      VarArrayRedim(UV_vGubn_jongru, iLength);
      VarArrayRedim(UV_vDesc_jongru, iLength);
      VarArrayRedim(UV_vGubn_class,  iLength);
      VarArrayRedim(UV_vYsno_virus1, iLength);
      VarArrayRedim(UV_vYsno_virus2, iLength);
      VarArrayRedim(UV_vYsno_virus3, iLength);
      VarArrayRedim(UV_vYsno_virus4, iLength);
      VarArrayRedim(UV_vYsno_virus5, iLength);
      VarArrayRedim(UV_vYsno_virus6, iLength);
      VarArrayRedim(UV_vYsno_virus7, iLength);
      VarArrayRedim(UV_vYsno_virus8, iLength);
      VarArrayRedim(UV_vYsno_virus10, iLength);
      VarArrayRedim(UV_vDesc_virus,  iLength);
      VarArrayRedim(UV_vYsno_clin1,  iLength);
      VarArrayRedim(UV_vYsno_clin2,  iLength);
      VarArrayRedim(UV_vYsno_clin3,  iLength);
      VarArrayRedim(UV_vYsno_clin4,  iLength);
      VarArrayRedim(UV_vYsno_clin5,  iLength);
      VarArrayRedim(UV_vYsno_clin6,  iLength);
      VarArrayRedim(UV_vYsno_clin7,  iLength);
      VarArrayRedim(UV_vDesc_Gum,    iLength);
      VarArrayRedim(UV_vGum_Chk,     iLength);
      VarArrayRedim(UV_vGum_Remark,  iLength);
      VarArrayRedim(UV_vGum_Form1,   iLength);
      VarArrayRedim(UV_vGum_Form2,   iLength);
      VarArrayRedim(UV_vGum_Form3,   iLength);
      VarArrayRedim(UV_vGum_Form3_Etc,   iLength);
      VarArrayRedim(UV_vSang_Form1,  iLength);
      VarArrayRedim(UV_vSang_Form1_Gubn,  iLength);
      VarArrayRedim(UV_vSang_Form2,  iLength);
      VarArrayRedim(UV_vSang_Form3,  iLength);
      VarArrayRedim(UV_vSang_Form4,  iLength);
      VarArrayRedim(UV_vSun_Form1,   iLength);
      VarArrayRedim(UV_vSun_Form2,   iLength);
      VarArrayRedim(UV_vSun_Form3,   iLength);
      VarArrayRedim(UV_vSun_Form4,   iLength);
      VarArrayRedim(UV_vSun_Form5,   iLength);
      VarArrayRedim(UV_vDesc_bang,   iLength);
      VarArrayRedim(UV_vDesc_buwi,   iLength);
      VarArrayRedim(UV_vCod_sawon,   iLength);
      VarArrayRedim(UV_vCod_doct,    iLength);
      VarArrayRedim(UV_vCod_jindan_doct,    iLength);
      VarArrayRedim(UV_vcheck_jindan,    iLength);
      VarArrayRedim(UV_vDat_last,    iLength);
      VarArrayRedim(UV_vDat_time,    iLength);
      VarArrayRedim(UV_vDat_time_R,  iLength);
      VarArrayRedim(UV_vDat_result,  iLength);
      VarArrayRedim(UV_vDesc_kgbl,   iLength);
      VarArrayRedim(UV_vDesc_sokun,  iLength);
      VarArrayRedim(UV_vDesc_sokun1,  iLength);
      VarArrayRedim(UV_vDesc_Pan,    iLength);
      VarArrayRedim(UV_vCod_Pan,     iLength);
      VarArrayRedim(UV_vYsno_sokun,  iLength);
      VarArrayRedim(UV_vCheck_Year,  iLength);
      VarArrayRedim(UV_vGubn_P012,   iLength);
      VarArrayRedim(UV_vGubn_Pan,    iLength);
      VarArrayRedim(UV_vCod_sokun,   iLength);
      VarArrayRedim(UV_vDoctor,      iLength);
      VarArrayRedim(UV_vFloor,       iLength);
      VarArrayRedim(UV_vChuga_HM,    iLength);
      VarArrayRedim(UV_vRent,        iLength);
      VarArrayRedim(UV_vCmt,         iLength);
      VarArrayRedim(UV_vDat_Rent,    iLength);
      VarArrayRedim(UV_vDAT_Return,  iLength);
      VarArrayRedim(UV_vEdt_delnum,  iLength);
      VarArrayRedim(UV_vdesc_kita,   iLength);

   end;
end;

function TfrmLDAI01.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (pnlJisa.Visible and (cmbJisa.ItemIndex = -1)) or
      (rbDate.Checked and (mskDate.Text = '')) or
      (rbRDate.Checked and (mskRDate.Text = '')) or
      (rbJumin.Checked and (mskJumin.Text = '')) or
      (rbCell.Checked and (mskcell.Text ='')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

function TfrmLDAI01.UF_Save : Boolean;
var i : Integer;
sWhere,sOrderBy,a,b,c : String;
begin
   Result := False;

   // 작업중인지 Check
   //if not UV_bEdit then exit;

   // Validation Check
   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   if (copy(retDESC_SOKUN.Text,250,1) = ' ') and (copy(retDESC_SOKUN.Text,251,1) <> ' ') then
      retDESC_SOKUN.Text := copy(retDESC_SOKUN.Text,1,250) + ' ' + copy(retDESC_SOKUN.Text,251,length(retDESC_SOKUN.Text) - 250);
   if (copy(retDESC_SOKUN.Text,500,1) = ' ') and (copy(retDESC_SOKUN.Text,501,1) <> ' ') then
      retDESC_SOKUN.Text := copy(retDESC_SOKUN.Text,1,500) + ' ' + copy(retDESC_SOKUN.Text,501,length(retDESC_SOKUN.Text) - 500);
   if (copy(retDESC_SOKUN.Text,750,1) = ' ') and (copy(retDESC_SOKUN.Text,751,1) <> ' ') then
      retDESC_SOKUN.Text := copy(retDESC_SOKUN.Text,1,750) + ' ' + copy(retDESC_SOKUN.Text,751,length(retDESC_SOKUN.Text) - 750);

   if (copy(retDESC_SOKUN1.Text,250,1) = ' ') and (copy(retDESC_SOKUN1.Text,251,1) <> ' ') then
      retDESC_SOKUN1.Text := copy(retDESC_SOKUN1.Text,1,250) + ' ' + copy(retDESC_SOKUN1.Text,251,length(retDESC_SOKUN1.Text) - 250);
   if (copy(retDESC_SOKUN1.Text,500,1) = ' ') and (copy(retDESC_SOKUN1.Text,501,1) <> ' ') then
      retDESC_SOKUN1.Text := copy(retDESC_SOKUN1.Text,1,500) + ' ' + copy(retDESC_SOKUN1.Text,501,length(retDESC_SOKUN1.Text) - 500);
   if (copy(retDESC_SOKUN1.Text,750,1) = ' ') and (copy(retDESC_SOKUN1.Text,751,1) <> ' ') then
      retDESC_SOKUN1.Text := copy(retDESC_SOKUN1.Text,1,750) + ' ' + copy(retDESC_SOKUN1.Text,751,length(retDESC_SOKUN1.Text) - 750);

     a :=  copy(retDESC_SOKUN1.Text,1,250);
     b :=  copy(retDESC_SOKUN1.Text,251,250);
     c := copy(retDESC_SOKUN1.Text,501,250);




   // DB 작업
   try
      if (UV_vCod_hm[i] = 'P003') then
      begin
    
          with qryCell_sokun do
          begin
           Active := False;

           sWhere := 'WHERE ';

            if Copy(cmbJisa.Text, 1, 2) <> '00' then
              sWhere := sWhere + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';
    
           sWhere := sWhere + ' A.cod_cell = ''' + UV_vCod_cell[i] + ''' AND ';

            if rbDate.Checked then
            begin
              sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''';
            end
            else if rbRDate.Checked then
            begin
              sWhere := sWhere + ' C.dat_last = ''' + mskRDate.Text + '''';
            end
            else if rbJumin.Checked then
            begin
                sWhere := sWhere + ' B.num_jumin = ''' + mskJumin.Text + '''';
            end
            else if rbCell.Checked then
            begin
                sWhere := sWhere + ' A.cod_cell = ''' + UV_vCod_cell[i] + '''';
            end;

            SQL.Clear;
            SQL.Add(UV_sBasicSQL_cs + sWhere + sOrderBy);
           Active := True;
    
            if RecordCount >0 then
            begin
              with qryU_cellsokun do
              begin
                ParamByName('cod_bjjs').AsString   := UV_vCod_bjjs[i];
                ParamByName('num_id').AsString     := UV_vNum_id[i];
                ParamByName('cod_cell').AsString   := UV_vCod_cell[i];
                ParamByName('desc_sokun').AsString := copy(retDESC_SOKUN1.Text,1,250);
                ParamByName('desc_sokun1').AsString:= copy(retDESC_SOKUN1.Text,251,250);
                ParamByName('desc_sokun2').AsString:= copy(retDESC_SOKUN1.Text,501,250);
                ParamByName('cod_sokun').AsString  := Uppercase(edtCod_sokun.Text);
                ParamByName('cod_update').AsString := GV_sUserId;
               ParamByName('dat_update').AsString := GV_sToday;

                Execsql;
              end;
            end
            else
            begin
              with qryI_cellsokun do
              begin
                ParamByName('cod_bjjs').AsString   := UV_vCod_bjjs[i];
                ParamByName('num_id').AsString     := UV_vNum_id[i];
                ParamByName('cod_cell').AsString   := UV_vCod_cell[i];
                ParamByName('dat_gmgn').AsString   := UV_vDat_gmgn[i];
                ParamByName('cod_jisa').AsString   := UV_vCod_jisa[i];
                ParamByName('cod_hm').AsString     := UV_vCod_hm[i];
                ParamByName('desc_sokun').AsString := copy(retDESC_SOKUN1.Text,1,250);
                ParamByName('desc_sokun1').AsString := copy(retDESC_SOKUN1.Text,251,250);
                ParamByName('desc_sokun2').AsString := copy(retDESC_SOKUN1.Text,501,250);
                ParamByName('cod_sokun').AsString  := Uppercase(edtCod_sokun.Text);
                ParamByName('cod_insert').AsString := GV_sUserId;
                ParamByName('dat_insert').AsString := GV_sToday;
    
                Execsql;
              end;
           end;
    
         end;

      end;




      with qryU_Cell do
      begin
         //P010일 경우 검체접수일 최초저장시 당일날짜 저장
         if (UV_vCod_hm[i] = 'P010') then
         begin
           if mskDAT_LAST.Text = '' then ParamByName('DAT_LAST').AsString  := GV_sToday
           else                          ParamByName('DAT_LAST').AsString  := mskDAT_LAST.Text;
         end
         else
         ParamByName('DAT_LAST').AsString         := mskDAT_LAST.Text;

         ParamByName('GUBN_JONGRU').AsString := Copy(cmbGUBN_JONGRU.Text, 1, 1);
         ParamByName('DESC_JONGRU').AsString := edtDESC_JONGRU.Text;
         if Pos('-', Copy(cmbGUBN_CLASS.Text, 1, 2)) > 0 then
            ParamByName('GUBN_CLASS').AsString  := Copy(cmbGUBN_CLASS.Text, 1, 1)
         else
            ParamByName('GUBN_CLASS').AsString  := Copy(cmbGUBN_CLASS.Text, 1, 2);
         ParamByName('YSNO_VIRUS1').AsString := GF_B2YN(ckYSNO_VIRUS1.Checked);
         ParamByName('YSNO_VIRUS2').AsString := GF_B2YN(ckYSNO_VIRUS2.Checked);
         ParamByName('YSNO_VIRUS3').AsString := GF_B2YN(ckYSNO_VIRUS3.Checked);
         ParamByName('YSNO_VIRUS4').AsString := GF_B2YN(ckYSNO_VIRUS4.Checked);
         ParamByName('YSNO_VIRUS5').AsString := GF_B2YN(ckYSNO_VIRUS5.Checked);
         ParamByName('YSNO_VIRUS6').AsString := GF_B2YN(ckYSNO_VIRUS6.Checked);
         ParamByName('YSNO_VIRUS7').AsString := GF_B2YN(ckYSNO_VIRUS7.Checked);
         ParamByName('YSNO_VIRUS8').AsString := GF_B2YN(ckYSNO_VIRUS8.Checked);
         ParamByName('YSNO_VIRUS10').AsString := GF_B2YN(ckYSNO_VIRUS10.Checked);
         ParamByName('DESC_VIRUS').AsString  := edtDESC_VIRUS.Text;
         ParamByName('YSNO_CLIN1').AsString  := GF_B2YN(ckYSNO_CLIN1.Checked);
         ParamByName('YSNO_CLIN2').AsString  := GF_B2YN(ckYSNO_CLIN2.Checked);
         ParamByName('YSNO_CLIN3').AsString  := GF_B2YN(ckYSNO_CLIN3.Checked);
         ParamByName('YSNO_CLIN4').AsString  := GF_B2YN(ckYSNO_CLIN4.Checked);
         ParamByName('YSNO_CLIN5').AsString  := GF_B2YN(ckYSNO_CLIN5.Checked);
         ParamByName('YSNO_CLIN6').AsString  := GF_B2YN(ckYSNO_CLIN6.Checked);
         ParamByName('YSNO_CLIN7').AsString  := GF_B2YN(ckYSNO_CLIN7.Checked);
         If edtDesc_Gum1.Checked Then
            ParamByName('Desc_Gum').AsString := '1'
         Else If edtDEsc_Gum2.Checked Then
            ParamByName('Desc_Gum').AsString := '2'
         Else
            ParamByName('Desc_Gum').AsString := '';
         If edtGum_Chk1.Checked Then
            ParamByName('Gum_Chk').AsString  := '1'
         Else If edtGum_Chk2.Checked Then
            ParamByName('Gum_Chk').AsString  := '2'
         Else
            ParamByName('Gum_Chk').AsString  := '';
         ParamByName('Gum_Remark').AsString       := edtGum_Remark.Text;
         ParamByName('Gum_Form1').AsString        := GF_B2YN(edtGum_Form1.Checked);
         ParamByName('Gum_Form2').AsString        := GF_B2YN(edtGum_Form2.Checked);
         ParamByName('Gum_Form3').AsString        := GF_B2YN(edtGum_Form3.Checked);
         ParamByName('Gum_Form3_Etc').AsString    := edtGum_Form3_Etc.Text;
         ParamByName('Sang_Form1').AsString       := GF_B2YN(edtSang_Form1.Checked);
         ParamByName('Sang_Form1_Gubn').AsInteger := CmbSang_Form1_gubn.Itemindex;
         ParamByName('Sang_Form2').AsString       := GF_B2YN(edtSang_Form2.Checked);
         ParamByName('Sang_Form3').AsString       := GF_B2YN(edtSang_Form3.Checked);
         ParamByName('Sang_Form4').AsString       := GF_B2YN(edtSang_Form4.Checked);
         ParamByName('Sun_Form1').AsString        := GF_B2YN(edtSun_Form1.Checked);
         ParamByName('Sun_Form2').AsString        := GF_B2YN(edtSun_Form2.Checked);
         ParamByName('Sun_Form3').AsString        := GF_B2YN(edtSun_Form3.Checked);
         ParamByName('Sun_Form5').AsString        := Copy(canSun_Form5.text,1,1);
         ParamByName('Sun_Form4').AsString        := edtSun_Form4.Text;
         ParamByName('DESC_BANG').AsString        := edtDESC_BANG.Text;
         ParamByName('DESC_BUWI').AsString        := edtDESC_BUWI.Text;
         ParamByName('COD_SAWON').AsString        := Copy(cmbCOD_SAWON.Text, 1, 6);
         //외부 입력 의사 추가
         If chk_jindan.Checked Then
         ParamByName('check_jindan').AsString      := 'Y'
         else ParamByName('check_jindan').AsString := 'N';
         ParamByName('COD_jindan_doctor').AsString  := Copy(cmbCOD_Jindan_DOCT.Text, 1, 6);

         ParamByName('COD_DOCT').AsString         := Copy(cmbCOD_DOCT.Text, 1, 6);
         ParamByName('DAT_RESULT').AsString       := mskDAT_RESULT.Text;
         ParamByName('DESC_KGBL').AsString        := retDESC_KGBL.Text;
         ParamByName('DESC_SOKUN1').AsString      := copy(retDESC_SOKUN.Text,1,250);//copy((StringReplace(retDESC_SOKUN.Text,''#$D#$A'','',[rfReplaceAll])),1,250);   //20141119 곽윤설
         ParamByName('DESC_SOKUN2').AsString      := copy(retDESC_SOKUN.Text,251,250);
         ParamByName('DESC_SOKUN3').AsString      := copy(retDESC_SOKUN.Text,501,250);
         ParamByName('DESC_SOKUN4').AsString      := copy(retDESC_SOKUN.Text,751,250);
         ParamByName('DESC_SOKUN5').AsMemo        := copy(retDESC_SOKUN.Text,1001,length(retDESC_SOKUN.Text) - 1000);

         //*19.03.20 일자 부터 공단 추가된 내용 반영(자궁경부암 전구단계 의심 4번에 추가), 한칸씩 밀림. 전 결과 문제시 내용 이해 바람. 이유경 팀장//
         If edtCod_Pan1.Checked Then
            ParamByName('Cod_Pan').AsString := '1'
         Else If edtCod_Pan2.Checked Then
            ParamByName('Cod_Pan').AsString := '2'
         Else If edtCod_Pan3.Checked Then
            ParamByName('Cod_Pan').AsString := '3'
         Else If edtCod_Pan4.Checked Then
            ParamByName('Cod_Pan').AsString := '4'
         Else If edtCod_Pan5.Checked Then
            ParamByName('Cod_Pan').AsString := '5'
         Else If edtCod_Pan6.Checked Then
            ParamByName('Cod_Pan').AsString := '6'
         Else
            ParamByName('Cod_Pan').AsString := '';
         If (edtCod_Pan6.Checked)Then
            ParamByName('Desc_Pan').AsString := Trim(edtDesc_Pan.Text)
         Else
            ParamByName('Desc_Pan').AsString := '';

         // 결과일자에 따라서 소견유무 표시 (Y:판독중, C:판독완료, N: 미판독)
         if Trim(mskDAT_RESULT.Text) <> '' then
            ParamByName('YSNO_SOKUN').AsString  := 'Y'
         else
            ParamByName('YSNO_SOKUN').AsString  := 'N';
         if UV_vCod_hm[i] = 'P012' then ParamByName('GUBN_P012').AsString   := Copy(Cmb_P012.Text, 1, 2)
         else if UV_vCod_hm[i] = 'P013' then ParamByName('GUBN_P012').AsString   := Copy(Cmb_P013.Text, 1, 2)
         else ParamByName('GUBN_P012').AsString   := Copy(Cmb_P013.Text, 1, 2);

         ParamByName('GUBN_PAN').AsString    := edtGUBN_PAN.Text;

         ParamByName('Check_Year').AsString  := Copy(GV_sToday, 3, 2);
         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         ParamByName('COD_BJJS').AsString    := UV_vCod_bjjs[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('COD_CELL').AsString    := UV_vCod_cell[i];
         ParamByName('COD_SOKUN').AsString   := edtCod_sokun.Text;

         Execsql;
         GP_query_log(qryU_Cell, 'LD1I12', 'U', 'Y');
      end;

      if UV_vDat_gmgn[i] < '20180101' then
      begin
      qryCa_cancer2009.Active := false;
      qryCa_cancer2009.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[i];
      qryCa_cancer2009.ParamByName('Num_id').AsString   := UV_vNum_id[i];
      qryCa_cancer2009.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[i];
      qryCa_cancer2009.Active := True;
      if qryCa_cancer2009.IsEmpty = False then
      begin
         //자궁경부
         if UV_vCod_hm[i] = 'P003' then
         begin
            if UV_vCod_Jisa[i] ='15' then
            begin
                 with qryU_ca_cancer2009_Can5_15 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';

                    if (b_sgCAppSet = True) and (UV_vCod_Jisa[i] ='15') then
                       ParamByName('can5_doctor').AsString := GV_sUserId;
                    ParamByName('DAT_UPDATE').AsString  := GV_sToday;
                    ParamByName('COD_UPDATE').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryU_ca_cancer2009_Can5, 'LD1I12', 'U', 'Y');
                 end;
            end
            else
            begin
                 with qryU_ca_cancer2009_Can5 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';
                    ParamByName('DAT_UPDATE').AsString  := GV_sToday;
                    ParamByName('COD_UPDATE').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryU_ca_cancer2009_Can5, 'LD1I12', 'U', 'Y');
                 end;
            end;
         End;
      end
      else
      begin
         //자궁경부
         if UV_vCod_hm[i] = 'P003' then
         begin
            if UV_vCod_Jisa[i] ='15' then
            begin
                 with qryI_Ca_Cancer2009_Can5_15 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';

                    //if (b_sgCAppSet = True) and (UV_vCod_Jisa[i]='15') then
                       ParamByName('can5_doctor').AsString := GV_sUserId;

                    ParamByName('DAT_Insert').AsString  := GV_sToday;
                    ParamByName('COD_Insert').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryI_Ca_Cancer2009_Can5, 'LD1I12', 'I', 'Y');
                 end;
            end
            else
            begin
                 with qryI_Ca_Cancer2009_Can5 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';
                    ParamByName('DAT_Insert').AsString  := GV_sToday;
                    ParamByName('COD_Insert').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryI_Ca_Cancer2009_Can5, 'LD1I12', 'I', 'Y');
                 end;
            end;
         End;
      end;
   end
   else
   begin
      qryCa_cancer2018.Active := false;
      qryCa_cancer2018.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[i];
      qryCa_cancer2018.ParamByName('Num_id').AsString   := UV_vNum_id[i];
      qryCa_cancer2018.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[i];
      qryCa_cancer2018.Active := True;
      if qryCa_cancer2018.IsEmpty = False then
      begin
         //자궁경부
         if UV_vCod_hm[i] = 'P003' then
         begin
            if UV_vCod_Jisa[i] ='15' then
            begin
                 with qryU_ca_cancer2018_Can5_15 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';

                    //if (b_sgCAppSet = True) and (UV_vCod_Jisa[i] ='15') then
                       ParamByName('can5_doctor').AsString := GV_sUserId;
                    ParamByName('DAT_UPDATE').AsString  := GV_sToday;
                    ParamByName('COD_UPDATE').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryU_ca_cancer2018_Can5_15, 'LD1I12', 'U', 'Y');
                 end;
            end
            else
            begin
                 with qryU_ca_cancer2018_Can5 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';
                    ParamByName('DAT_UPDATE').AsString  := GV_sToday;
                    ParamByName('COD_UPDATE').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryU_ca_cancer2018_Can5, 'LD1I12', 'U', 'Y');
                 end;
            end;
         End;
      end
      else
      begin
         //자궁경부
         if UV_vCod_hm[i] = 'P003' then
         begin
            if UV_vCod_Jisa[i] ='15' then
            begin
                 with qryI_Ca_Cancer2018_Can5_15 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';

                    //if (b_sgCAppSet = True) and (UV_vCod_Jisa[i]='15') then
                       ParamByName('can5_doctor').AsString := GV_sUserId;

                    ParamByName('DAT_Insert').AsString  := GV_sToday;
                    ParamByName('COD_Insert').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryI_Ca_Cancer2018_Can5_15, 'LD1I12', 'I', 'Y');
                 end;
            end
            else
            begin
                 with qryI_Ca_Cancer2018_Can5 do
                 begin
                    ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
                    ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
                    ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
                    SUBSOKUN    := '';
                    UV_SBSOKUN1 := '';
                    UV_SBSOKUN2 := '';
                    UV_SBSOKUN3 := '';
                    UV_SBSOKUN4 := '';
                    SUBSOKUN    := Can_Advice.Text;
                    GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                    ParamByName('can5_advice').AsString  := UV_SBSOKUN1;
                    ParamByName('can5_advice1').AsString := UV_SBSOKUN2;
                    If edtCod_Pan1.Checked Then
                       ParamByName('can5_panjung').AsString := '1'
                    Else If edtCod_Pan2.Checked Then
                       ParamByName('can5_panjung').AsString := '2'
                    Else If edtCod_Pan3.Checked Then
                       ParamByName('can5_panjung').AsString := '3'
                    Else If edtCod_Pan4.Checked Then
                       ParamByName('can5_panjung').AsString := '4'
                    Else If edtCod_Pan5.Checked Then
                       ParamByName('can5_panjung').AsString := '5'
                    Else If edtCod_Pan6.Checked Then
                       ParamByName('can5_panjung').AsString := '6'
                    Else
                       ParamByName('can5_panjung').AsString := '';
                    If (edtCod_Pan6.Checked)Then
                       ParamByName('can5_panjung_etc').AsString := Trim(edtDesc_Pan.Text)
                    Else
                       ParamByName('can5_panjung_etc').AsString := '';
                    ParamByName('DAT_Insert').AsString  := GV_sToday;
                    ParamByName('COD_Insert').AsString  := GV_sUserId;
                    Execsql;
                    GP_query_log(qryI_Ca_Cancer2018_Can5, 'LD1I12', 'I', 'Y');
                 end;
            end;
         End;
      end;
   end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end;




   // Grid에 자료를 반영
   with grdMaster do
   begin
      // 현재 Field의 자료를 Grid관련 Variant변수에 할당
      UV_vGubn_jongru[i] := Copy(cmbGUBN_JONGRU.Text, 1, 1);
      UV_vDesc_jongru[i] := edtDESC_JONGRU.Text;

      if Pos('-', Copy(cmbGUBN_CLASS.Text, 1, 2)) > 0 then
         UV_vGubn_class[i]  := Copy(cmbGUBN_CLASS.Text, 1, 1)
      else
         UV_vGubn_class[i]  := Copy(cmbGUBN_CLASS.Text, 1, 2);
      UV_vYsno_virus1[i] := GF_B2YN(ckYSNO_VIRUS1.Checked);
      UV_vYsno_virus2[i] := GF_B2YN(ckYSNO_VIRUS2.Checked);
      UV_vYsno_virus3[i] := GF_B2YN(ckYSNO_VIRUS3.Checked);
      UV_vYsno_virus4[i] := GF_B2YN(ckYSNO_VIRUS4.Checked);
      UV_vYsno_virus5[i] := GF_B2YN(ckYSNO_VIRUS5.Checked);
      UV_vYsno_virus6[i] := GF_B2YN(ckYSNO_VIRUS6.Checked);
      UV_vYsno_virus7[i] := GF_B2YN(ckYSNO_VIRUS7.Checked);
      UV_vYsno_virus8[i] := GF_B2YN(ckYSNO_VIRUS8.Checked);
      UV_vYsno_virus10[i] := GF_B2YN(ckYSNO_VIRUS10.Checked);
      UV_vDesc_virus[i]  := edtDESC_VIRUS.Text;
      UV_vYsno_clin1[i]  := GF_B2YN(ckYSNO_CLIN1.Checked);
      UV_vYsno_clin2[i]  := GF_B2YN(ckYSNO_CLIN2.Checked);
      UV_vYsno_clin3[i]  := GF_B2YN(ckYSNO_CLIN3.Checked);
      UV_vYsno_clin4[i]  := GF_B2YN(ckYSNO_CLIN4.Checked);
      UV_vYsno_clin5[i]  := GF_B2YN(ckYSNO_CLIN5.Checked);
      UV_vYsno_clin6[i]  := GF_B2YN(ckYSNO_CLIN6.Checked);
      UV_vYsno_clin7[i]  := GF_B2YN(ckYSNO_CLIN7.Checked);
      UV_vGum_Form1[i]   := GF_B2YN(edtGum_Form1.Checked);
      UV_vGum_Form2[i]   := GF_B2YN(edtGum_Form2.Checked);
      UV_vGum_Form3[i]   := GF_B2YN(edtGum_Form3.Checked);
      UV_vGum_Form3_Etc[i]   := edtGum_Form3_Etc.Text;
      UV_vSang_Form1[i]  := GF_B2YN(edtSang_Form1.Checked);
      UV_vSang_Form1_Gubn[i]  := CmbSang_Form1_gubn.ItemIndex;
      UV_vSang_Form2[i]  := GF_B2YN(edtSang_Form2.Checked);
      UV_vSang_Form3[i]  := GF_B2YN(edtSang_Form3.Checked);
      UV_vSang_Form4[i]  := GF_B2YN(edtSang_Form4.Checked);
      UV_vSun_Form1[i]   := GF_B2YN(edtSun_Form1.Checked);
      UV_vSun_Form2[i]   := GF_B2YN(edtSun_Form2.Checked);
      UV_vSun_Form3[i]   := GF_B2YN(edtSun_Form3.Checked);
      UV_vSun_Form4[i]   := edtSun_Form4.Text;
      UV_vSun_Form5[i]   := canSun_form5.ItemIndex;
      If edtDesc_Gum1.Checked Then
         UV_vDesc_Gum[i] := '1'
      Else If edtDesc_Gum2.Checked Then
         UV_vDesc_Gum[i] := '2';
      If edtGum_Chk1.Checked Then
         UV_vGum_Chk[i] := '1'
      Else If edtGum_Chk2.Checked Then
         UV_vGum_Chk[i] := '2';
      UV_vGum_Remark[i] := edtGum_Remark.Text;
      UV_vDesc_bang[i]       := edtDESC_BANG.Text;
      UV_vDesc_buwi[i]       := edtDESC_BUWI.Text;
      UV_vCod_sawon[i]       := Copy(cmbCOD_SAWON.Text, 1, 6);
      UV_vCod_doct[i]        := Copy(cmbCOD_DOCT.Text, 1, 6);
      //외부 입력 의사 추가
      If chk_jindan.Checked Then
         UV_vCheck_jindan[i]   := 'Y'
      else UV_vCheck_jindan[i] := 'N';
      UV_vCod_jindan_doct[i] := Copy(cmbCOD_jindan_DOCT.Text, 1, 6);

      UV_vDat_last[i]    := mskDAT_LAST.Text;
      UV_vDat_time[i]    := edt_time_S.Text;
      UV_vDat_time_R[i]  := edt_time_R.Text;
      UV_vDat_result[i]  := mskDAT_RESULT.Text;
      UV_vDesc_kgbl[i]   := retDESC_KGBL.Text;
      UV_vDesc_sokun[i]  := retDESC_SOKUN.Text;
      UV_vDesc_sokun1[i] := retDESC_SOKUN1.Text;
      If edtCod_Pan1.Checked Then
         UV_vCod_Pan[i] := '1'
      Else If edtCod_Pan2.Checked Then
         UV_VCod_Pan[i] := '2'
      Else If edtCod_Pan3.Checked Then
         UV_VCod_Pan[i] := '3'
      Else If edtCod_Pan4.Checked Then
         UV_VCod_Pan[i] := '4'
      Else If edtCod_Pan5.Checked Then
         UV_VCod_Pan[i] := '5'
      Else If edtCod_Pan6.Checked Then
         UV_VCod_Pan[i] := '6';
      UV_vDesc_Pan[i]   := Trim(edtDesc_Pan.Text);
      UV_vdesc_kita[i]  := Trim(edt_kita.Text);

      if Trim(mskDAT_RESULT.Text) <> '' then
         UV_vYsno_sokun[i] := 'Y'
      else
         UV_vYsno_sokun[i] := 'N';

      UV_vGubn_pan[i]     := edtGUBN_PAN.Text;
      UV_vCod_sokun[i]    := edtCod_sokun.Text;
      UV_vFloor[i]        := Edt_floor.Text;
      UV_vChuga_HM[i]     := Edt_cod_chuga.Text;
      UV_vRent[i]         := Lab_slide.Caption;
      UV_vCmt[i]          := Edt_cmt.Text;
      UV_vDat_Rent[i]     := msk_dat_rent.Text;
      UV_vDAT_Return[i]   := msk_dat_Return.Text;
      if (UV_vCod_hm[i] = 'P012') then UV_vGubn_P012[i]    := Copy(Cmb_P012.Text, 1, 2);

      // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;
   end;

   UV_iCod_doct := cmbCOD_DOCT.ItemIndex;
   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLDAI01.btnSaveClick(Sender: TObject);
begin
   inherited;
   UF_Save;
end;

procedure TfrmLDAI01.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);

   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   pnlJisa.Visible := True;                                                     //06.11.24 철. 지사선택 보이게끔.

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   // Combo Populate
   GP_ComboFill(cmbGUBN_JONGRU, 'GCJR');
//   GP_ComboFill(cmbGUBN_CLASS,  'CLAS');
   with qryCommon do
   begin
//      Close;
      Open;
      while not Eof do
      begin
         cmbGUBN_CLASS.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                             FieldByName('DESC_CODE').AsString);
         Next;
      end;
   end;
   GP_ComboSawonAll(cmbCOD_SAWON, 0);
   GP_ComboSawonAll(cmbCOD_DOCT, 1);
   GP_ComboSawonAll(cmbCOD_jindan_DOCT, 1);

   // 분석부 조직실 전체...
   if (GV_sUserID = '150007') or (GV_sUserID = '150365') or (GV_sUserID = '410015') or
      (GV_sUserID = '150281') or (GV_sUserID = '150627') or (GV_sUserID = '150332') or
      (GV_sUserID = '150644') or (GV_sUserID = '150486') or (GV_sUserID = '150681') or
      (GV_sUserID = '150875') or (GV_sUserID = '150344') or (GV_sUserID = '151026') or
      (GV_sUserID = '150879') then cmbRelation.Enabled := True
   else
   cmbRelation.Enabled := False;

   if (GV_sUserID = '150007') or (GV_sUserID = '150365') or (GV_sUserID = '410015') or
      (GV_sUserID = '150281') or (GV_sUserID = '150627') or (GV_sUserID = '150332') or
      (GV_sUserID = '150644') or (GV_sUserID = '150486') or (GV_sUserID = '150681') or
      (GV_sUserID = '150875') or (GV_sUserID = '150344') or (GV_sUserID = '151026') or
      (GV_sUserID = '150879') then GroupBox8.Visible := True
   else
   GroupBox8.Visible := False;
   // SQL문 저장
   UV_sBasicSQL := qryCell.SQL.Text;
   UV_sBasicSQL_cs := qryCell_sokun.SQL.Text;
   pnlHide.Visible := False;
   pnlHide4.Visible := False;
   pnlHide.Left := 80;
   pnlHide.Top := 225;

   UV_iCod_doct := -1;
   cmbGubn.ItemIndex := 0;
   mskDate.Text := GV_sToday;
end;

procedure TfrmLDAI01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자료를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_Jisa[DataRow-1];
      2 : begin
             if rbDate.Checked  then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
             if rbRDate.Checked  then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
             if rbJumin.Checked then Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
             if rbCell.Checked then Value  := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
          end;
      3 : begin
             if rbDate.Checked  then Value := UV_vDesc_name[DataRow-1];
             if rbRDate.Checked  then Value := UV_vDesc_name[DataRow-1];
             if rbJumin.Checked then Value := '';
             if rbCell.Checked  then Value := UV_vDesc_name[DataRow-1];
          end;
      4 : begin
             if rbCell.Checked then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1])
             else                   Value := UV_vDesc_buwi[DataRow-1];
          end;
      5 : begin
             if rbCell.Checked then Value := UV_vDesc_buwi[DataRow-1]
             else                   Value := UV_vCod_cell[DataRow-1];
          end;
      6 : Value := UV_vCod_hm[DataRow-1];
      7 : Value := UV_vCod_bjjs[DataRow-1];  //Visible := False;
      8 : Value := UV_vDoctor[DataRow-1];
      9 : begin
             if UV_vDat_gmgn[DataRow-1] >= '20070701' then
             begin
                if      Trim(UV_vYsno_sokun[DataRow-1]) = 'C' then Value := '판독완료'
                else if Trim(UV_vYsno_sokun[DataRow-1]) = 'Y' then Value := '*판독중*'
                else Value := '미판독';
             end
             else
             begin
                if (Trim(UV_vYsno_sokun[DataRow-1]) = 'Y') or
                   (Trim(UV_vYsno_sokun[DataRow-1]) = 'C') then Value := '판독완료'
                else
                begin
                   if UV_vDat_result[DataRow-1] <> '' then Value := '판독완료'
                   else                                    Value := '미판독';
                end;
             end;
          end;
   end;

   // 조직검사일 경우 색깔 변경...
   if (Copy(UV_vCod_cell[DataRow-1],1,1) = 'J') then
      grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
end;

procedure TfrmLDAI01.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderby, sWhere_cs: String;
begin
   inherited;
   UV_bEdit := True;
   btnPsave.Enabled   := True;
   btnSave.Enabled    := True;
   btnPcancel.Enabled := True;

   if mskDate.Text >= '20160107' then
        UV_bGmgn := True
   else
        UV_bGmgn := False;

   with qryUser_priv do                              //06.11.24 철. 조회지사가 로그인한 아이디의 지사와 일치할 경우 저장할 수 있음.
   begin
      close;
      ParamByName('cod_sawon').AsString := GV_sUserId;
      open;
      if recordcount > 0 then
      begin
         if (copy(cmbJisa.Text,1,2) = copy(GV_sUserID,1,2)) and (FieldBYName('ysno_read').AsString = 'N') then
         begin
            btnPSave.Enabled := True;
            btnSave.Enabled  := True;
         end
         else
         begin
            btnPSave.Enabled := False;
            btnSave.Enabled  := False;
         end;
      end;
   end;

   // 분석부 조직실 전체...
   if  (GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026') then       //이유경 이시내 전산
   begin
      btnPSave.Enabled := True;
      btnSave.Enabled  := True;
   end
   else
   begin
      btnPSave.Enabled := False;
      btnSave.Enabled  := False;
   end;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   retDesc_sokun.text :='';
   retDesc_sokun1.text:='';
   // Grid Setting
   UP_GridSet;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryCell do
   begin
      Active := False;

      sWhere := 'WHERE ';

      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''';
         //201808020 최윤선 선임 요청으로 변경
         if Chk_order.Checked then sWhere := sWhere + 'and (A.gubn_order = ''Y'' or A.check_jindan = ''Y'' ) ';
      end
      else if rbRDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''';
         //201808020 최윤선 선임 요청으로 변경
         if Chk_order.Checked then sWhere := sWhere + 'and (A.gubn_order = ''Y'' or A.check_jindan = ''Y'' ) ';
      end
      else if rbRDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''';
         //201808020 최윤선 선임 요청으로 변경
         if Chk_order.Checked then sWhere := sWhere + 'and (A.gubn_order = ''Y'' or A.check_jindan = ''Y'' ) ';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' B.num_jumin = ''' + mskJumin.Text + '''';
      end
      else if rbCell.Checked then
      begin
         sWhere := sWhere + ' A.desc_buwi = ''' + mskCell.Text + '''';
      end;

      // 검사구분에 따라서 Where 조건추가
      if      cmbGubn.ItemIndex = 0  then sWhere := sWhere + ' AND ((SUBSTRING(cod_cell,1,1) = ''P'') or (A.cod_hm = ''B008'') or (A.cod_hm = ''B009''))'
      else if cmbGubn.ItemIndex = 1  then sWhere := sWhere + ' AND A.cod_hm = ''B008''  '
      else if cmbGubn.ItemIndex = 2  then sWhere := sWhere + ' AND A.cod_hm = ''B009''  '
      else if cmbGubn.ItemIndex = 3 then sWhere := sWhere + ' AND A.cod_hm = ''P003''   '
      else if cmbGubn.ItemIndex = 4 then sWhere := sWhere + ' AND A.cod_hm = ''P006''   '
      else if cmbGubn.ItemIndex = 5 then sWhere := sWhere + ' AND A.cod_hm = ''P010''   '
      else if cmbGubn.ItemIndex = 6 then sWhere := sWhere + ' AND A.cod_hm = ''P011''   '
      else if cmbGubn.ItemIndex = 7 then sWhere := sWhere + ' AND A.cod_hm = ''P012''   '
      else if cmbGubn.ItemIndex = 8 then sWhere := sWhere + ' AND A.cod_hm = ''P013''   ';

      if cmbOrder.ItemIndex = 1 then
         sOrderBy := ' Order by A.desc_buwi '
      else
         sOrderBy := ' ORDER BY A.dat_gmgn, A.cod_jisa, B.desc_name ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
      Active := True;

      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD1I12', 'Q', 'N');
         while Not Eof do
         begin
            UP_VarMemSet(index);
            UV_vCod_bjjs[index]    := FieldByName('COD_BJJS').AsString;
            UV_vNum_id[index]      := FieldByName('NUM_ID').AsString;
            UV_vNum_jumin[index]   := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]   := FieldByName('DESC_NAME').AsString;
            UV_vCod_jisa[index]    := FieldByName('COD_JISA').AsString;
            UV_vDat_gmgn[index]    := FieldByName('DAT_GMGN').AsString;
            UV_vCod_cell[index]    := FieldByName('COD_CELL').AsString;
            UV_vCod_hm[index]      := FieldByName('COD_HM').AsString;
            UV_vGubn_jongru[index] := FieldByName('GUBN_JONGRU').AsString;
            UV_vDesc_jongru[index] := FieldByName('DESC_JONGRU').AsString;
            UV_vGubn_class[index]  := FieldByName('GUBN_CLASS').AsString;
            UV_vYsno_virus1[index] := FieldByName('YSNO_VIRUS1').AsString;
            UV_vYsno_virus2[index] := FieldByName('YSNO_VIRUS2').AsString;
            UV_vYsno_virus3[index] := FieldByName('YSNO_VIRUS3').AsString;
            UV_vYsno_virus4[index] := FieldByName('YSNO_VIRUS4').AsString;
            UV_vYsno_virus5[index] := FieldByName('YSNO_VIRUS5').AsString;
            UV_vYsno_virus6[index] := FieldByName('YSNO_VIRUS6').AsString;
            UV_vYsno_virus7[index] := FieldByName('YSNO_VIRUS7').AsString;
            UV_vYsno_virus8[index] := FieldByName('YSNO_VIRUS8').AsString;
            UV_vYsno_virus10[index] := FieldByName('YSNO_VIRUS10').AsString;
            UV_vDesc_virus[index]  := FieldByName('DESC_VIRUS').AsString;
            UV_vYsno_clin1[index]  := FieldByName('YSNO_CLIN1').AsString;
            UV_vYsno_clin2[index]  := FieldByName('YSNO_CLIN2').AsString;
            UV_vYsno_clin3[index]  := FieldByName('YSNO_CLIN3').AsString;
            UV_vYsno_clin4[index]  := FieldByName('YSNO_CLIN4').AsString;
            UV_vYsno_clin5[index]  := FieldByName('YSNO_CLIN5').AsString;
            UV_vYsno_clin6[index]  := FieldByName('YSNO_CLIN6').AsString;
            UV_vYsno_clin7[index]  := FieldByName('YSNO_CLIN7').AsString;
            UV_vDesc_Gum[Index]    := FieldByName('Desc_Gum').AsString;
            UV_vGum_Chk[Index]     := FieldByName('Gum_Chk').AsString;
            UV_vGum_Remark[Index]  := FieldByName('Gum_Remark').AsString;
            UV_vGum_Form1[Index]   := FieldByName('Gum_Form1').AsString;
            UV_vGum_Form2[Index]   := FieldByName('Gum_Form2').AsString;
            UV_vGum_Form3[Index]   := FieldByName('Gum_Form3').AsString;
            UV_vGum_Form3_Etc[Index]   := FieldByName('Gum_Form3_Etc').AsString;
            UV_vSang_Form1[Index]  := FieldByName('Sang_Form1').AsString;
            UV_vSang_Form1_Gubn[Index]  := FieldByName('Sang_Form1_Gubn').AsString;
            UV_vSang_Form2[Index]  := FieldByName('Sang_Form2').AsString;
            UV_vSang_Form3[Index]  := FieldByName('Sang_Form3').AsString;
            UV_vSang_Form4[Index]  := FieldByName('Sang_Form4').AsString;
            UV_vSun_Form1[Index]   := FieldByName('Sun_Form1').AsString;
            UV_vSun_Form2[Index]   := FieldByName('Sun_Form2').AsString;
            UV_vSun_Form3[Index]   := FieldByName('Sun_Form3').AsString;
            UV_vSun_Form4[Index]   := FieldByName('Sun_Form4').AsString;
            UV_vSun_Form5[Index]   := FieldByName('Sun_Form5').AsString;
            UV_vDesc_bang[index]   := FieldByName('DESC_BANG').AsString;
            UV_vDesc_buwi[index]   := FieldByName('DESC_BUWI').AsString;
            UV_vCod_sawon[index]   := FieldByName('COD_SAWON').AsString;
            UV_vCod_doct[index]    := FieldByName('COD_DOCT').AsString;
            UV_vCod_jindan_doct[index] := FieldByName('COD_jindan_doctor').AsString;
            UV_vCheck_jindan[index]    := FieldByName('check_jindan').AsString;
            UV_vCod_Pan[index]     := FieldByName('COD_PAN').AsString;
            UV_vDESC_Pan[index]    := FieldByName('DESC_PAN').AsString;
            UV_vDat_last[index]    := FieldByName('DAT_LAST').AsString;
            UV_vDat_time[index]    := FieldByName('dat_time').AsString;
            UV_vDat_time_R[index]  := FieldByName('dat_time_R').AsString;
            UV_vDat_result[index]  := FieldByName('DAT_RESULT').AsString;
            UV_vDesc_kgbl[index]   := FieldByName('DESC_KGBL').AsString;
            UV_vDesc_sokun[index]  := FieldByName('DESC_SOKUN1').AsString +
                                      FieldByName('DESC_SOKUN2').AsString +
                                      FieldByName('DESC_SOKUN3').AsString +
                                      FieldByName('DESC_SOKUN4').AsString +
                                      FieldByName('DESC_SOKUN5').AsString;

            UV_vYsno_sokun[index]  := FieldByName('Ysno_sokun').AsString;
            UV_vCheck_Year[index]  := FieldByName('Check_Year').AsString;
            UV_vGubn_P012 [index]  := FieldByName('Gubn_P012').AsString;
            UV_vGubn_Pan [index]   := FieldByName('Gubn_Pan').AsString;
            UV_vCod_sokun[index]   := FieldByName('cod_sokun').AsString;
            UV_vChuga_HM[index]    := FieldByName('cod_Chuga').AsString;
            UV_vEdt_delnum[index]  := FieldByName('num_delete').AsString;
            if FieldByName('desc_delete').AsString <> '' then
            begin
               UV_vdesc_kita[index]   := FieldByName('desc_delete').AsString;
               edt_kita.color := $007171FF;
            end
            else
            begin
                 UV_vdesc_kita[index]   := '';
                 edt_kita.color := clWindow;
            end;
            
            if qryCell.FieldByName('gubn_jinch').AsString <> '' then
              begin
              qryFloor.Active := False;
              qryFloor.ParamByName('cod_group').AsString  := 'FLOR';
              qryFloor.ParamByName('cod_detail').AsString := qryCell.FieldByName('gubn_jinch').AsString;
              qryFloor.Active := True;
              // 층구분 표시 ..20170720
              UV_vFloor[index] := qryFloor.fieldByName('desc_code').AsString;
              end;
              UV_vChuga_HM[index]    := FieldByName('cod_Chuga').AsString;

            //슬라이드 대여 표시
            with qry_slide do
            begin
            qry_slide.Active := False;
            qry_slide.ParamByName('DAT_GMGN').AsString  := qryCell.FieldByName('dat_gmgn').AsString;
            qry_slide.ParamByName('num_id').AsString    := qryCell.FieldByName('num_id').AsString;
            qry_slide.ParamByName('desc_buwi').AsString := qryCell.FieldByName('desc_buwi').AsString;
            qry_slide.Active := True;
            if RecordCount > 0 then
              begin
                    UV_vRent[index]          := qry_slide.FieldByName('rent_yn').AsString;
                    UV_vDat_Rent[index]      := qry_slide.FieldByName('dat_rent').AsString;
                    UV_vDAT_Return[index]    := qry_slide.FieldByName('dat_return').AsString;
                    UV_vCmt[Index]           := qry_slide.FieldByName('cell_cmt').AsString;
              end
              else  UV_vRent[index]          := '';
            end;

            if (UV_vCod_hm[index] = 'P003') then
            begin
                with qryCell_sokun do
                begin
                  Active := False;

                  sWhere_cs := 'WHERE ';

                  if Copy(cmbJisa.Text, 1, 2) <> '00' then
                     sWhere_cs := sWhere_cs + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

                  if rbDate.Checked then
                  begin
                  sWhere_cs := sWhere_cs + ' A.dat_gmgn = ''' + mskDate.Text + '''';
                  end
                  else if rbRDate.Checked then
                  begin
                  sWhere_cs := sWhere_cs + ' C.dat_last = ''' + mskRDate.Text + '''';
                  end
                  else if rbJumin.Checked then
                  begin
                  sWhere_cs := sWhere_cs + ' B.num_jumin = ''' + mskJumin.Text + '''';
                  end
                  else if rbCell.Checked then
                  begin
                  sWhere_cs := sWhere_cs + ' A.cod_cell = ''' + qryCell.FieldByName('COD_CELL').AsString + '''';
                  end;

                  sWhere_cs := sWhere_cs + 'And A.cod_cell = ''' + qryCell.FieldByName('COD_CELL').AsString + '''';

                  SQL.Clear;
                  SQL.Add(UV_sBasicSQL_cs + sWhere_cs);
                  Active := True;

                  if RecordCount > 0 then
                  begin
                    UV_vDesc_sokun1[index] := qryCell_sokun.FieldByName('desc_sokun').AsString
                                              + qryCell_sokun.FieldByName('desc_sokun1').AsString
                                              + qryCell_sokun.FieldByName('desc_sokun2').AsString;
                  end
                  else
                    UV_vDesc_sokun1[index] :='';

                  Active := False;
                end;
            end;

            if (UV_vCod_hm[index] = 'P001') or
               (UV_vCod_hm[index] = 'P003') or
               (UV_vCod_hm[index] = 'B009') or
               (UV_vCod_hm[index] = 'P010') or
               (UV_vCod_hm[index] = 'P012') or
               (UV_vCod_hm[index] = 'P013')then
            begin
               with qryCa_Request do
               begin
                  qryCa_Request.Active := False;
                  qryCa_Request.ParamByName('cod_jisa').AsString := UV_vCod_jisa[index];
                  qryCa_Request.ParamByName('num_id').AsString   := UV_vNum_id[index];
                  qryCa_Request.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[index];
                  qryCa_Request.Active := True;
                  if qryCa_Request.RecordCount > 0 then
                  begin
                     UV_vDoctor[index] := qryCa_Request.FieldByName('DOCTOR').AsString;
                  end;
               end;
            end;


            Next;
            Inc(index);
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


   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLDAI01.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // 취소시
   // 1.입력이면 화면 Clear
   // 2.수정이면 Re-query
   if UV_iStatus = GC_Insert then
   begin
      GP_FieldClear(pnlDetail);
   end;

   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLDAI01.UP_Change(Sender: TObject);
var sPan1, sSokun1 : String;
begin
   inherited;
   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = edtCOD_SOKUN then
   begin
      //IF edtCOD_SOKUN.Text <> 'po03' then  edtGUBN_PAN.Text := '';   // 2018.04.08 최윤선 선임 요청

      // '****'이면 미검이므로 자료를 Clear
      if Trim(edtGUBN_PAN.Text) = '' then  //2015.03.04 곽윤설 .. 저장값 우선 끌어오기
      begin
         if GF_CellSokun_P(edtCod_sokun.Text, sSokun1, sPan1) then
         begin
            edtGUBN_PAN.Text   := sPan1;
         end;
      end
      else
      begin
         //2018.08.08 최윤선 소견코드 지웠을때 소견 판정도 지워지게 해달라
         if GF_CellSokun_P(edtCod_sokun.Text, sSokun1, sPan1) then
         begin
         if edtCOD_SOKUN.Text <> 'po03' then
         begin
            if sPan1 <> edtGUBN_PAN.Text then
            begin
               edtGUBN_PAN.Text   := sPan1;
            end;
         end;
         end;
      end;

      if grdMaster.CurrentDataRow > 0 then
      begin
         if (UV_vCod_hm[grdMaster.CurrentDataRow-1] = 'P003') and (Length(Trim(edtCod_sokun.Text)) <> 4) then  //20150804
         begin
            cmbGUBN_CLASS.Text      := '';
            ckYSNO_VIRUS1.Checked   := False;
            ckYSNO_VIRUS2.Checked   := False;
            ckYSNO_VIRUS3.Checked   := False;
            ckYSNO_VIRUS4.Checked   := False;
            ckYSNO_VIRUS5.Checked   := False;
            ckYSNO_VIRUS6.Checked   := False;
            ckYSNO_VIRUS7.Checked   := False;
            ckYSNO_VIRUS8.Checked   := False;
            ckYSNO_VIRUS10.Checked  := False;
            edtDESC_VIRUS.Text      := '';
            edtGum_Form1.Checked    := False;
            edtGum_Form2.Checked    := False;
            edtGum_Form3.Checked    := False;
            edtGum_Form3_Etc.Text   := '';
            edtDesc_Gum1.Checked    := False;
            edtDesc_Gum2.Checked    := False;
            edtGum_Chk1.Checked     := False;
            edtGum_Chk2.Checked     := False;
            edtSun_Form1.Checked    := False;
            edtSun_Form2.Checked    := False;
            edtSun_Form3.Checked    := False;
            canSun_form5.ItemIndex  := 0;
            edtSun_Form4.Text       := '';
            edtSang_Form1.Checked   := False;
            edtSang_Form2.Checked   := False;
            edtSang_Form3.Checked   := False;
            edtSang_Form4.Checked   := False;
            CmbSang_Form1_gubn.Text := '';
            edtCod_Pan1.checked     := false;
            edtCod_Pan2.checked     := false;
            edtCod_Pan3.checked     := false;
            edtCod_Pan4.checked     := false;
            edtCod_Pan5.checked     := false;
            edtCod_Pan6.checked     := false;
            ckYSNO_CLIN1.Checked    := False;
            ckYSNO_CLIN2.Checked    := False;
            ckYSNO_CLIN3.Checked    := False;
            ckYSNO_CLIN4.Checked    := False;
            ckYSNO_CLIN5.Checked    := False;
            ckYSNO_CLIN6.Checked    := False;
            ckYSNO_CLIN7.Checked    := False;
         end;
      end;

      if      Uppercase(edtCod_sokun.Text) = 'QW01' then Cmb_P012.ItemIndex := 1
      else if Uppercase(edtCod_sokun.Text) = 'QW02' then Cmb_P012.ItemIndex := 2
      else if Uppercase(edtCod_sokun.Text) = 'QW03' then Cmb_P012.ItemIndex := 3
      else if Uppercase(edtCod_sokun.Text) = 'QW04' then Cmb_P012.ItemIndex := 4
      else if Uppercase(edtCod_sokun.Text) = 'QW05' then Cmb_P012.ItemIndex := 5
      else if Uppercase(edtCod_sokun.Text) = 'QW06' then Cmb_P012.ItemIndex := 6
      else if Uppercase(edtCod_sokun.Text) = 'QW07' then Cmb_P012.ItemIndex := 2
      else if Uppercase(edtCod_sokun.Text) = 'QW08' then Cmb_P012.ItemIndex := 2;

      if      Uppercase(edtCod_sokun.Text) = 'UI01' then Cmb_P013.ItemIndex := 1
      else if Uppercase(edtCod_sokun.Text) = 'UI02' then Cmb_P013.ItemIndex := 2
      else if Uppercase(edtCod_sokun.Text) = 'UI03' then Cmb_P013.ItemIndex := 3
      else if Uppercase(edtCod_sokun.Text) = 'UI04' then Cmb_P013.ItemIndex := 4
      else if Uppercase(edtCod_sokun.Text) = 'UI05' then Cmb_P013.ItemIndex := 5
      else if Uppercase(edtCod_sokun.Text) = 'UI06' then Cmb_P013.ItemIndex := 6;

      //[2017.06.23-유동구]PO37,PO38코드는 기존에 사용하던 코드로 변경으로 문제 발생(화면표기 문제)
      //신규코드(PO50,PO51)로 재 추진...
      if (Uppercase(edtCod_sokun.Text) = 'PO01') or
         (Uppercase(edtCod_sokun.Text) = 'PO50') or
         (Uppercase(edtCod_sokun.Text) = 'PO51') or
         (Uppercase(edtCod_sokun.Text) = 'PO36') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 6;        //Class       7-정상
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan1.Checked      := True;     //노신판정    정상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO02') or
              (Uppercase(edtCod_sokun.Text) = 'PO05') or
              (Uppercase(edtCod_sokun.Text) = 'PO37') or
              (Uppercase(edtCod_sokun.Text) = 'PO38') or
              (Uppercase(edtCod_sokun.Text) = 'PO45') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO03') or
              (Uppercase(edtCod_sokun.Text) = 'PO04') or
              (Uppercase(edtCod_sokun.Text) = 'PO29') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
         ckYsno_Virus8.Checked    := True;     //Virus 구분  염증
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if Uppercase(edtCod_sokun.Text) = 'PO06' then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYSNO_VIRUS3.Checked    := True;     //Virus 구분  트리코모나스
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if Uppercase(edtCod_sokun.Text) = 'PO07' then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         //edtDESC_VIRUS.Text := '질세균분포변화';   //Virus 구분  기타
         ckYsno_Virus10.Checked := True;
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if Uppercase(edtCod_sokun.Text) = 'PO08' then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYSNO_VIRUS4.Checked    := True;     //Virus 구분  캔디다
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if Uppercase(edtCod_sokun.Text) = 'PO09' then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYsno_Virus7.Checked    := True;     //Virus 구분  방선균
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if Uppercase(edtCod_sokun.Text) = 'PO10' then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화
         ckYSNO_VIRUS2.Checked    := True;     //Virus 구분  헤르페스
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO11') or
              (Uppercase(edtCod_sokun.Text) = 'PO12') or
              (Uppercase(edtCod_sokun.Text) = 'PO13') or
              (Uppercase(edtCod_sokun.Text) = 'PO14') or
              (Uppercase(edtCod_sokun.Text) = 'PO39') or
              (Uppercase(edtCod_sokun.Text) = 'PO40') or
              (Uppercase(edtCod_sokun.Text) = 'PO41') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 8;        //Class       9-비정형 편평세포
         if (Uppercase(edtCod_sokun.Text) = 'PO40') then    //Virus 구분  염증
            ckYsno_Virus8.Checked := True;
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSang_Form1.Checked    := True;     //비정형편평상피세포
         if (Uppercase(edtCod_sokun.Text) = 'PO41') then
            CmbSang_Form1_gubn.ItemIndex := 2    //(고위험)
         else
            CmbSang_Form1_gubn.ItemIndex := 1;   //(일반)
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO15') or
              (Uppercase(edtCod_sokun.Text) = 'PO16') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 10;       //Class       11-저등급 이형성
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSang_Form2.Checked    := True;     //저등급편평상피세포
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO17') or
              (Uppercase(edtCod_sokun.Text) = 'PO18') or
              (Uppercase(edtCod_sokun.Text) = 'PO19') or
              (Uppercase(edtCod_sokun.Text) = 'PO35') or
              (Uppercase(edtCod_sokun.Text) = 'PO42') or
              (Uppercase(edtCod_sokun.Text) = 'PO46') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 11;       //Class       12-고등급 이형성
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSang_Form3.Checked    := True;     //고등급편평상피세포
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO43') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 12;       //Class       13-상피내암좀
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSun_Form2.Checked    := True;     //상피내선암종
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO20') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 13;       //Class       14-침윤성암
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSang_Form4.Checked    := True;     //침윤성편평세포암종
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan5.Checked      := True;     //노신판정    자궁경부암 의심
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO21') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 9;        //Class       10-비정형 선세포
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSun_Form1.Checked     := True;     //비정형선상피세포
         canSun_form5.ItemIndex   := 1;
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO22') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 9;        //Class       10-비정형 선세포
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSun_Form1.Checked     := True;     //비정형선상피세포
         canSun_form5.ItemIndex   := 2;
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan3.Checked      := True;     //노신판정    상피세포 이상
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO23') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 13;       //Class       14-침윤성암
         edtGum_Form2.Checked     := True;     //유형별진단  상피세포이상
         edtSun_Form3.Checked     := True;     //침윤성선암종
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan5.Checked      := True;     //노신판정    자궁경부암 의심
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO24') or
              (Uppercase(edtCod_sokun.Text) = 'PO25') or
              (Uppercase(edtCod_sokun.Text) = 'PO26') or
              (Uppercase(edtCod_sokun.Text) = 'PO27') or
              (Uppercase(edtCod_sokun.Text) = 'PO28') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 14;       //Class       15-부적절검체
         edtDesc_Gum2.Checked     := True;     //검체유형    부적절
         edtCod_Pan6.Checked      := True;     //노신판정    기타
         if (Uppercase(edtCod_sokun.Text) = 'PO24') then
            edtGum_Remark.Text := '세포부족'
         else if (Uppercase(edtCod_sokun.Text) = 'PO25') then
            edtGum_Remark.Text := '혈액도말'
         else if (Uppercase(edtCod_sokun.Text) = 'PO26') then
            edtGum_Remark.Text := '염증세포도말'
         else if (Uppercase(edtCod_sokun.Text) = 'PO27') then
            edtGum_Remark.Text := '건조도말'
         else if (Uppercase(edtCod_sokun.Text) = 'PO28') then
            edtGum_Remark.Text := '두꺼운도말';
      end
      else if (Uppercase(edtCod_sokun.Text) = 'PO30') or
              (Uppercase(edtCod_sokun.Text) = 'PO31') or
              (Uppercase(edtCod_sokun.Text) = 'PO32') or
              (Uppercase(edtCod_sokun.Text) = 'PO33') or
              (Uppercase(edtCod_sokun.Text) = 'PO34') then
      begin
         cmbGUBN_JONGRU.ItemIndex := 0;        //검체종류    1-Cervical
         cmbGUBN_CLASS.ItemIndex  := 7;        //Class       8-반응성 세포변화

         if (Uppercase(edtCod_sokun.Text) = 'PO30') then
            edtDESC_VIRUS.Text := '복합적세균' //Virus 구분  기타
         else if (Uppercase(edtCod_sokun.Text) = 'PO31') then
         begin
            ckYSNO_VIRUS3.Checked := True;     //Virus 구분  트리코모나스
            ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
            ckYsno_Virus8.Checked := True;     //Virus 구분  염증
         end
         else if (Uppercase(edtCod_sokun.Text) = 'PO32') then
         begin
            ckYSNO_VIRUS10.Checked := True; //Virus 구분  기타
            ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
            ckYsno_Virus8.Checked := True;     //Virus 구분  염증
         end
         else if (Uppercase(edtCod_sokun.Text) = 'PO33') then
         begin
            ckYSNO_VIRUS4.Checked := True;     //Virus 구분  캔디다
            ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
            ckYsno_Virus8.Checked := True;     //Virus 구분  염증
         end
         else if (Uppercase(edtCod_sokun.Text) = 'PO34') then
         begin
            ckYsno_Virus6.Checked    := True;     //Virus 구분  반응성세포변화
            ckYsno_Virus7.Checked := True;     //Virus 구분  방선균
            ckYsno_Virus8.Checked := True;     //Virus 구분  염증
         end;
         edtGum_Form1.Checked     := True;     //유형별진단  음성
         edtDesc_Gum1.Checked     := True;     //검체유형    적절
         edtGum_Chk1.Checked      := True;     //선상피세포  유
         edtCod_Pan2.Checked      := True;     //노신판정    염증성 또는 감염성 질환
      end; ////

   end
   else if Sender = cmbGUBN_CLASS then
   begin
      if      copy(cmbGUBN_CLASS.Text,1,1) = '7'  then edtCod_Pan1.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,1) = '8'  then edtCod_Pan2.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,1) = '9'  then edtCod_Pan3.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '10' then edtCod_Pan3.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '11' then edtCod_Pan3.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '12' then edtCod_Pan3.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '13' then edtCod_Pan3.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '14' then edtCod_Pan5.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '15' then edtCod_Pan6.checked :=  true
      else if copy(cmbGUBN_CLASS.Text,1,2) = '16' then edtCod_Pan1.checked :=  true
      else
      begin
         edtCod_Pan1.checked :=  false;
         edtCod_Pan2.checked :=  false;
         edtCod_Pan3.checked :=  false;
         edtCod_Pan4.checked :=  false;
         edtCod_Pan5.checked :=  false;
         edtCod_Pan6.checked :=  false;
      end;

      if cmbGUBN_CLASS.Text = '' then
      begin
         edtCod_Pan1.checked :=  false;
         edtCod_Pan2.checked :=  false;
         edtCod_Pan3.checked :=  false;
         edtCod_Pan4.checked :=  false;
         edtCod_Pan5.checked :=  false;
         edtCod_Pan6.checked :=  false;
      end;


   end;
  if grdMaster.CurrentDataRow > 0 then
  begin
    if Sender = chk_jindan then
    begin
       if chk_jindan.Checked then
       begin
        IF (UV_vCod_Jindan_doct[grdMaster.CurrentDataRow-1] = '' ) then GP_ComboDisplay(cmbCOD_jindan_Doct, '151227', 6);
        IF (UV_vCod_sawon[grdMaster.CurrentDataRow-1] = '' ) then GP_ComboDisplay(cmbCOD_SAWON, '800016', 6);
       end
       else if chk_jindan.Checked = false then
       begin
         cmbCOD_JINDAN_DOCT.ItemIndex := -1;
         if (copy(GV_sUserId,1,2) = '15') then

         if Trim(UV_vCod_sawon[grdMaster.CurrentDataRow-1]) = '' then
         begin
            if (copy(GV_sUserId,1,2) = '15') then
            begin
              cmbCOD_SAWON.Font.Color := clRed;
                   if (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'P001') or
                      (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'P006') or
                      (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'P011')then   GP_ComboDisplay(cmbCOD_SAWON, '150943', 6) //한영애
              else if (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,1) = 'P')   then   GP_ComboDisplay(cmbCOD_SAWON, '151030', 6) //최윤선
              else if (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,1) = 'B')   then   GP_ComboDisplay(cmbCOD_SAWON, '151030', 6) //최윤선
                 else                                                                   GP_ComboDisplay(cmbCOD_SAWON, GV_sUserId, 6);
            end
            else cmbCOD_SAWON.ItemIndex := -1;
         end
         else
         begin
           cmbCOD_SAWON.Font.Color := clBlack;
           GP_ComboDisplay(cmbCOD_SAWON, UV_vCod_sawon[grdMaster.CurrentDataRow-1], 6);
         end;
        end;
        {else cmbCOD_SAWON.ItemIndex := -1;}

    end;
  end;

   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLDAI01.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
Var
   Max_Num : integer;
   sWhere, sOrderby, sWhere_cs: String;
   i, index, itemp : Integer;
   sList, sHangmok, sName : String;
   eSuga : Extended;
   sCancel: String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;
   canSun_form5.visible := False;
   edtDesc_Gum1.checked := false;
   edtDesc_Gum2.checked := false;
   edtGum_Remark.Text   := '';
   edtGum_Chk1.checked  := false;
   edtGum_Chk2.checked  := false;
   Can_Advice.Text := '';
   edt_kita.Text := '';
   edt_del_num.Text := '';
   edt_num1.Checked := false;
   edt_num2.Checked := false;
   edt_num3.Checked := false;

   edtGUBN_PAN.Text      := UV_vGubn_pan[NewRow-1];
   edtCod_sokun.Text     := UV_vCod_sokun[NewRow-1];
//   edtCod_sokun.Text := '';   //2014.04.10 곽윤설 주석처리
//   edtGUBN_PAN.Text := '';    //2015.03.04 곽윤설 주석처리

   edtGUBN_PAN.Enabled   := True;
   //20150327 곽윤설
   if UV_vCod_hm[Newrow - 1] = 'P012' then
   begin
       if ((rbDate.Checked) and (mskDate.Text >= '20150401')) or
          ((rbJumin.Checked) and (UV_vDat_gmgn[Newrow - 1] >= '20150401')) then
       begin
          Cmb_p012.Items.Clear;
          Cmb_P012.Items.Add('');
          Cmb_P012.Items.Add('01-검체불량(부적합)');
          Cmb_P012.Items.Add('02-Class I (정상)');
          Cmb_P012.Items.Add('03-Class II(반응성)');
          Cmb_P012.Items.Add('04-Class III(비정형세포)');
          Cmb_P012.Items.Add('05-Class IV(암의심)');
          Cmb_P012.Items.Add('06-Class V(암)');
       end
       else
       begin
          Cmb_p012.Items.Clear;
          Cmb_P012.Items.Add('');
          Cmb_P012.Items.Add('01-검체불량');
          Cmb_P012.Items.Add('02-음성ⅰ');
          Cmb_P012.Items.Add('03-음성ⅱ');
          Cmb_P012.Items.Add('04-의양성ⅲ');
          Cmb_P012.Items.Add('05-양성ⅳ');
          Cmb_P012.Items.Add('06-양성ⅴ');
       end;
   end;
   Cmb_P012.ItemIndex := 0;

   pnlHide.Visible := False;
   pnlHide4.Visible := False;
   //병리 제외한 나머지는 판독중일때 결과 조회 불가능 하도록 적용 본원이유굥  ..20170424 수지
   if ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '150644') or (GV_sUserID = '151021') or
       (GV_sUserID = '151022') or (GV_sUserID = '150879') or (GV_sUserID = '101030') or (GV_sUserID = '151026')) then
   begin
       pnlHide.Visible := False;
       pnlHide4.Visible := False;
   end
   ELSE  if (UV_vYsno_sokun[Newrow - 1] = 'Y') or (UV_vYsno_sokun[Newrow - 1] = 'N') then
   begin
     if UV_vCod_hm[Newrow - 1] = 'P003' then pnlHide4.Visible := True;
     pnlHide.Visible := true;
   end;
   if (UV_vCod_hm[Newrow - 1] = 'P012')  then
   begin
      Pnl_P012.Visible := True;
   end
   else
   begin
      Pnl_P012.Visible := False;
   end;
   if  (UV_vCod_hm[Newrow - 1] = 'P013') then
   begin
      Pnl_P013.Visible := True;
   end
   else
   begin
      Pnl_P013.Visible := False;
   end;

   if (UV_vCod_hm[Newrow - 1] = 'P003') then
   begin
      //공단 자궁암관련 항목 변경 .. 20150102  수지
      if ((rbDate.Checked) and (mskDate.Text >= '20180101')) or
         ((rbJumin.Checked) and (UV_vDat_gmgn[Newrow - 1] >= '20180101')) or
         ((rbRDate.Checked) and (mskRDate.Text >= '20180101')) or
         ((rbCell.Checked) and (UV_vDat_gmgn[Newrow - 1] >= '20180101'))  then canSun_form5.visible := True
      else  canSun_form5.visible := False;

      with Qry_Sign do
      begin
         Close;
         ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Newrow - 1];
         ParamByName('Num_id').AsString   := UV_vNum_id[Newrow - 1];
         ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[Newrow - 1];
         if UV_vCod_hm[Newrow - 1] = 'P003' then
            ParamByName('Program_id').AsString := 'LD1I123';
         open;
         Max_Num := Qry_Sign.FieldByName('seq_no').AsInteger;
         if Max_Num <> 0 then                                                                                              
         begin
            btnPSave.Enabled := False;
            btnPSave.Caption := '서명저장됨';
         end
         else
         begin
            btnPSave.Enabled := True;
            btnPSave.Caption := '저장(&S)';
         end;
      end;
   end;
   //2018.04.27 ubnormal 소견 코드시 한글 소견, 영문소견 글자 빨간색으로 표시 요청 ..본원 최윤선 선임
   if   (edtCod_sokun.text = 'po11') or (edtCod_sokun.text = 'po12')  or (edtCod_sokun.text = 'po13') or (edtCod_sokun.text = 'po14') or (edtCod_sokun.text = 'po39')
   or   (edtCod_sokun.text = 'po40') or (edtCod_sokun.text = 'po41')  or (edtCod_sokun.text = 'po21') or (edtCod_sokun.text = 'po22') or (edtCod_sokun.text = 'po15')
   or   (edtCod_sokun.text = 'po16') or (edtCod_sokun.text = 'po17')  or (edtCod_sokun.text = 'po18') or (edtCod_sokun.text = 'po19') or (edtCod_sokun.text = 'po35')
   or   (edtCod_sokun.text = 'po42') or (edtCod_sokun.text = 'po46')  or (edtCod_sokun.text = 'po43') or (edtCod_sokun.text = 'po20') or (edtCod_sokun.text = 'po23') THEN
   begin
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
   end
   else
   begin
       retDesc_Sokun.Font.Color := clBlack;
       retDesc_Sokun1.Font.Color := clBlack;
   end;

   with Qry_CK_Gulkwa do
   begin
      Close;
      Qry_CK_Gulkwa.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Newrow - 1];
      Qry_CK_Gulkwa.ParamByName('Num_id').AsString   := UV_vNum_id[Newrow - 1];
      Qry_CK_Gulkwa.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[Newrow - 1];
      Qry_CK_Gulkwa.ParamByName('COD_CELL').AsString := UV_vCod_cell[Newrow - 1];
      Qry_CK_Gulkwa.ParamByName('COD_HM').AsString   := UV_vCod_hm[Newrow - 1];

      Open;
      if Qry_CK_Gulkwa.IsEmpty = False then
      begin
         if UV_vCod_hm[Newrow - 1] = 'P003' then
         begin
            if UV_vDat_gmgn[Newrow - 1] < '20180101' then
            begin
               Can_Advice.Text := Trim(FieldByName('Can5_Advice').AsString);
               Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can5_Advice1').AsString);
            end
            else
            begin    
               Can_Advice.Text := Trim(FieldByName('Can5_Advice2018').AsString);
               Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can5_Advice12018').AsString);     
            end;
         end;
      end
      else
         Can_Advice.Text := '';
   end;

   //2014.04.10 곽윤설~
   if (UV_vCod_hm[Newrow - 1] = 'P010') and (UV_vGubn_jongru[NewRow-1] = '') and
      ((GV_sUserID = '150007') or (GV_sUserID = '150681')or (GV_sUserID = '151026') )then
       cmbGUBN_JONGRU.ItemIndex := 3  //~2014.04.10 곽윤설
   // Grid의 Row가 변경되면 Detail에 자료를 할당
   else GP_ComboDisplay(cmbGUBN_JONGRU, UV_vGubn_jongru[NewRow-1], 1);

   edtDESC_JONGRU.Text   := UV_vDesc_jongru[NewRow-1];

   if length(UV_vGubn_class[NewRow-1]) > 1 then
      GP_ComboDisplay(cmbGUBN_CLASS, UV_vGubn_class[NewRow-1], 2)
   else
      GP_ComboDisplay(cmbGUBN_CLASS, UV_vGubn_class[NewRow-1], 1);
   ckYSNO_VIRUS1.Checked := GF_YN2B(UV_vYsno_virus1[NewRow-1]);
   ckYSNO_VIRUS2.Checked := GF_YN2B(UV_vYsno_virus2[NewRow-1]);
   ckYSNO_VIRUS3.Checked := GF_YN2B(UV_vYsno_virus3[NewRow-1]);
   ckYSNO_VIRUS4.Checked := GF_YN2B(UV_vYsno_virus4[NewRow-1]);
   ckYSNO_VIRUS5.Checked := GF_YN2B(UV_vYsno_virus5[NewRow-1]);
   ckYSNO_VIRUS6.Checked := GF_YN2B(UV_vYsno_virus6[NewRow-1]);
   ckYSNO_VIRUS7.Checked := GF_YN2B(UV_vYsno_virus7[NewRow-1]);
   ckYSNO_VIRUS8.Checked := GF_YN2B(UV_vYsno_virus8[NewRow-1]);
   ckYSNO_VIRUS10.Checked := GF_YN2B(UV_vYsno_virus10[NewRow-1]);
   edtDESC_VIRUS.Text    := UV_vDesc_virus[NewRow-1];
   ckYSNO_CLIN1.Checked  := GF_YN2B(UV_vYsno_clin1[NewRow-1]);
   ckYSNO_CLIN2.Checked  := GF_YN2B(UV_vYsno_clin2[NewRow-1]);
   ckYSNO_CLIN3.Checked  := GF_YN2B(UV_vYsno_clin3[NewRow-1]);
   ckYSNO_CLIN4.Checked  := GF_YN2B(UV_vYsno_clin4[NewRow-1]);
   ckYSNO_CLIN5.Checked  := GF_YN2B(UV_vYsno_clin5[NewRow-1]);
   ckYSNO_CLIN6.Checked  := GF_YN2B(UV_vYsno_clin6[NewRow-1]);
   ckYSNO_CLIN7.Checked  := GF_YN2B(UV_vYsno_clin7[NewRow-1]);
   edtDESC_BANG.Text     := UV_vDesc_bang[NewRow-1];
   edtDESC_BUWI.Text     := UV_vDesc_buwi[NewRow-1];

   if (UV_vCod_hm[Newrow - 1] = 'P012') then GP_ComboDisplay(Cmb_P012, UV_vGubn_P012[NewRow-1], 2)
   else if  (UV_vCod_hm[Newrow - 1] = 'P013') then GP_ComboDisplay(Cmb_P013, UV_vGubn_P012[NewRow-1], 2);


   if Trim(UV_vCod_sawon[NewRow-1]) = '' then
   begin
      if (copy(GV_sUserId,1,2) = '15') then
      begin
         cmbCOD_SAWON.Font.Color := clRed;
         if (Copy(UV_vCod_hm[Newrow-1],1,4) = 'P001') or
            (Copy(UV_vCod_hm[Newrow-1],1,4) = 'P006') or
            (Copy(UV_vCod_hm[Newrow-1],1,4) = 'P011')then
            GP_ComboDisplay(cmbCOD_SAWON, '150943', 6) //한영애
         else if (Copy(UV_vCod_hm[Newrow-1],1,1) = 'P') then
            GP_ComboDisplay(cmbCOD_SAWON, '151030', 6) //최윤선
         else if (Copy(UV_vCod_hm[Newrow-1],1,1) = 'B') then
            GP_ComboDisplay(cmbCOD_SAWON, '151030', 6) //최윤선
         else
            GP_ComboDisplay(cmbCOD_SAWON, GV_sUserId, 6);
      end
      else cmbCOD_SAWON.ItemIndex := -1;
   end
   else
   begin
      cmbCOD_SAWON.Font.Color := clBlack;
      GP_ComboDisplay(cmbCOD_SAWON, UV_vCod_sawon[NewRow-1], 6);
   end;
//kmi 2003-08-09 의사선생 Fixed
   if Trim(UV_vCod_Doct[NewRow-1]) = '' then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         cmbCOD_Doct.Font.Color := clRed;

         if GV_sUserId = '150007' then
            GP_ComboDisplay(cmbCOD_Doct, '150007', 6)
         else if GV_sUserId = '150681' then
            GP_ComboDisplay(cmbCOD_Doct, '150681', 6)
         else
            GP_ComboDisplay(cmbCOD_Doct, '150007', 6);
      end
      else cmbCOD_Doct.ItemIndex := -1;
   end
   else
   begin
      cmbCOD_Doct.Font.Color := clBlack;
      GP_ComboDisplay(cmbCOD_Doct, UV_vCod_Doct[NewRow-1], 6);
   end;


   if Trim(UV_vCod_jindan_Doct[NewRow-1]) = '' then
   begin
      cmbCOD_jindan_Doct.ItemIndex := -1;
   end
   else
   begin
      cmbCOD_jindan_Doct.Font.Color := clBlack;
      GP_ComboDisplay(cmbCOD_jindan_Doct, UV_vCod_jindan_Doct[NewRow-1], 6);
   end;

   if (UV_vCheck_jindan[NewRow-1]) = 'Y' then chk_jindan.checked := True
   else if (UV_vCheck_jindan[NewRow-1]) = 'N' then chk_jindan.checked := False
   else chk_jindan.checked := False;

   If UV_vDesc_Gum[NewRow-1] = '1' Then
      Begin
         edtDesc_Gum1.Checked := true;
         edtDesc_Gum2.Checked := False;
      End
   Else If UV_vDesc_Gum[NewRow-1] = '2' Then
      Begin
         edtDesc_Gum1.Checked := False;
         edtDesc_Gum2.Checked := True;
      End;
{   Else
      Begin
         edtDesc_Gum1.Checked := True;
         edtDesc_Gum2.Checked := False;
      End;  }
   If UV_vGum_Chk[NewRow-1] = '1' Then
      Begin
         edtGum_Chk1.Checked := true;
         edtGum_Chk2.Checked := False;
      End
   Else If UV_vGum_Chk[NewRow-1] = '2' Then
      Begin
         edtGum_Chk1.Checked := False;
         edtGum_Chk2.Checked := True;
      End;
{   Else
      Begin
         edtGum_Chk1.Checked := True;
         edtGum_Chk2.Checked := False;
      End;     }

   edtGum_Remark.Text    := UV_vGum_Remark[NewRow-1];
   edtGum_Form1.Checked  := GF_YN2B(UV_vGum_Form1[NewRow-1]);
   edtGum_Form2.Checked  := GF_YN2B(UV_vGum_Form2[NewRow-1]);
   edtGum_Form3.Checked  := GF_YN2B(UV_vGum_Form3[NewRow-1]);
   edtGum_Form3_Etc.Text := UV_vGum_Form3_Etc[NewRow-1];
   edtSang_Form1.Checked := GF_YN2B(UV_vSang_Form1[NewRow-1]);
   if UV_vSang_Form1_Gubn[NewRow-1] <> '' then
      CmbSang_Form1_gubn.ItemIndex := StrToInt(UV_vSang_Form1_Gubn[NewRow-1])
   else
      CmbSang_Form1_gubn.ItemIndex := 0;
   edtSang_Form2.Checked := GF_YN2B(UV_vSang_Form2[NewRow-1]);
   edtSang_Form3.Checked := GF_YN2B(UV_vSang_Form3[NewRow-1]);
   edtSang_Form4.Checked := GF_YN2B(UV_vSang_Form4[NewRow-1]);
   edtSun_Form1.Checked  := GF_YN2B(UV_vSun_Form1[NewRow-1]);
   edtSun_Form2.Checked  := GF_YN2B(UV_vSun_Form2[NewRow-1]);
   edtSun_Form3.Checked  := GF_YN2B(UV_vSun_Form3[NewRow-1]);
   edtSun_Form4.Text     := UV_vSun_Form4[NewRow-1];
   if UV_vSun_Form5[NewRow-1] <> '' then
      canSun_Form5.ItemIndex := StrToInt(UV_vSun_Form5[NewRow-1])
   else
      canSun_Form5.ItemIndex := 0;

   mskDAT_LAST.Text      := UV_vDat_last[NewRow-1];
   edt_time_S.Text       := UV_vDat_time[NewRow-1];
   edt_time_R.Text       := UV_vDat_time_R[NewRow-1];
   if Trim(UV_vDat_result[NewRow-1]) = '' Then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         mskDat_Result.Font.Color := clRed;
         mskDat_Result.Text       := GV_sToday;
      end
      else mskDat_Result.Text := '';
   end
   else
   begin
      mskDat_Result.Font.Color := clBlack;
      mskDAT_RESULT.Text       := UV_vDat_result[NewRow-1];
   end;
   retDESC_KGBL.Text     := UV_vDesc_kgbl[NewRow-1];
   retDESC_SOKUN.Text    := '';
   retDESC_SOKUN1.Text   := '';
   retDESC_SOKUN.Text    := UV_vDesc_sokun[NewRow-1];
   retDESC_SOKUN1.Text   := UV_vDesc_sokun1[NewRow-1];


   Edt_floor.Text        := UV_vFloor[NewRow-1];
   if UV_vRent[NewRow-1] = 'Y' THEN
   begin
   Lab_slide.Caption     := '대여중';
   Lab_slide.Font.color  :=  clred;
   end
   else IF UV_vRent[NewRow-1] = 'N' THEN
   begin
   Lab_slide.Caption     := '반  납';
   Lab_slide.Font.color  :=  clBlack;
   end
   Else Lab_slide.Caption     := '';
   msk_dat_rent.text     := UV_vDat_Rent[NewRow-1];
   msk_dat_return.text   := UV_vDAT_Return[NewRow-1];
   Edt_cmt.Text          := UV_vCmt[NewRow-1];

   Edt_cod_chuga.Text    := '';
   sCancel :='';
   if UV_vChuga_HM[NewRow-1] <> '' then
      begin
         sHangmok := UV_vChuga_HM[NewRow-1];
         for itemp := 0 to (Length(sHangmok) div 4) - 1 do
         begin
            if (copy(sHangmok, (itemp*4)+1, 4) ='P003') or (copy(sHangmok, (itemp*4)+1, 4) ='P006') or (copy(sHangmok, (itemp*4)+1, 4) ='P010') or (copy(sHangmok, (itemp*4)+1, 4) ='P011') or (sHangmok ='P012') or
                (copy(sHangmok, (itemp*4)+1, 4) ='P013') or (copy(sHangmok, (itemp*4)+1, 4) ='B008') OR (copy(sHangmok, (itemp*4)+1, 4) ='B009') Then
                begin
                GF_HangmokEdit(copy(sHangmok, (itemp*4)+1, 4), GV_sToday, sName, eSuga);
                qry_cancel.Active := False;
                qry_cancel.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
                qry_cancel.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
                qry_cancel.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
                qry_cancel.ParamByName('cod_hm').AsString    := copy(sHangmok, (itemp*4)+1, 4);
                qry_cancel.Active := True;
                     if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '01' then   sCancel := '-> 취소'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '02' then   sCancel := '-> 거부'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '03' then   sCancel := '-> 연기'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '04' then   sCancel := '-> 임신'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '05' then   sCancel := '-> 기타'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '06' then   sCancel := '-> 임신가능'
                else if copy(qry_cancel.fieldbyName('gubn_cancel').AsString,1,2) = '07' then   sCancel := '-> 수유중'
                else   sCancel := '';

                Edt_cod_chuga.lines.add('[' + copy(sHangmok, (itemp*4)+1, 4) + ']' + sName + sCancel);
                end;

                memo1.text := '';
                qryComplain.Active := False;
                qryComplain.ParamByName('num_jumin').AsString  := UV_vNum_jumin[NewRow-1];
                qryComplain.Active := True;
                if qryComplain.RecordCount > 0 then
                begin
                  memo1.text := qryComplain.fieldByname('dat_sangdam').AsString +' : '+ qryComplain.fieldbyName('desc_sangdam').AsString;
                end;


         end;
      end;
   Edt_cod_chuga.Perform(WM_VSCROLL, SB_top, 0);

   edtCod_Pan1.Checked := False;
   edtCod_Pan2.Checked := False;
   edtCod_Pan3.Checked := False;
   edtCod_Pan4.Checked := False;
   edtCod_Pan5.Checked := False;
   edtCod_Pan6.checked := False;
   If UV_vCod_Pan[NewRow-1] = '1' Then
      edtCod_Pan1.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '2' Then
      edtCod_Pan2.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '3' Then
      edtCod_Pan3.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '4' Then
      edtCod_Pan4.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '5' Then
      edtCod_Pan5.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '6' Then
      edtCod_Pan6.Checked := True;
   If  (UV_vCod_Pan[NewRow-1] = '2') Or (UV_vCod_Pan[NewRow-1] = '6') Then
      edtDesc_Pan.Text := UV_vDesc_Pan[NewRow-1]
   Else
      edtDesc_Pan.Text := '';

   if cmbCod_doct.ItemIndex = -1 then
      cmbCod_doct.ItemIndex := UV_iCod_doct;

   if (UV_vCheck_Year[NewRow-1] = '09') then
   begin
      GroupBox7.Caption := '노신 판정(2009년)';
      GroupBox7.Font.Color := clBlue;
      GroupBox7.Font.Style := [fsBold];
      edtDesc_Pan.Font.Style := [];
      edtDesc_Pan.Font.Color := clNone;
   end
   else if (UV_vCheck_Year[NewRow-1] = '') and (trim(UV_vCod_Pan[NewRow-1]) <> '') then
   begin
      GroupBox7.Caption := '2009년도 노신기준이 아님';
      GroupBox7.Font.Color := clBlue;
      GroupBox7.Font.Style := [fsBold];
      edtDesc_Pan.Font.Style := [];
      edtDesc_Pan.Font.Color := clNone;
   end
   else
   begin
      GroupBox7.Caption := '노신 판정';
      GroupBox7.Font.Color := clNone;
      GroupBox7.Font.Style := [];
      edtDesc_Pan.Font.Style := [];
      edtDesc_Pan.Font.Color := clNone;
   end;

   // 2014.05.12 백승현
   // 검사코드 P003 소견코드 입력 & 저장 후 소견코드가 사라지지 않고 보이도록 수정
   if (edtCod_sokun.Text <> '') and ((UV_vCod_hm[NewRow-1] = 'P003') or (UV_vCod_hm[NewRow-1] = 'P012')) then
   begin
     if (GV_sDept <> '04') and (GV_sDept <> '12') and (GV_sDept <> '23') then
       edtCod_sokun.Text := '';
   end
   else
     edtCod_sokun.Text := UV_vCod_sokun[NewRow-1];

   //2019.01.11 수지 접수번호 삭제 이력 기능 추가
   if UV_vEdt_delnum[NewRow-1] <> '' then edt_del_num.text := UV_vEdt_delnum[NewRow-1]
   else edt_del_num.text := '';

   if UV_vdesc_kita[NewRow-1] <> '' then
   begin
      if UV_vdesc_kita[NewRow-1] = '이중접수' then
      begin
         edt_num1.checked;
         edt_kita.Text := UV_vdesc_kita[NewRow-1];
         edt_kita.Color :=  $007171FF;
      end
      else if UV_vdesc_kita[NewRow-1] = '수정건 삭제' then
      begin
         edt_num2.checked;
         edt_kita.Text := UV_vdesc_kita[NewRow-1];
         edt_kita.Color :=  $007171FF;
      end
      else
      begin
        edt_num3.checked;
        edt_kita.Text := UV_vdesc_kita[NewRow-1] ;
        edt_kita.Color :=  $007171FF;
      end;
   end;

   if UV_vdesc_kita[NewRow-1] <> '' then edt_kita.text := UV_vdesc_kita[NewRow-1]
   else edt_kita.text := '';

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLDAI01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   if Sender = btnJumin then
   begin
      if GF_InjoukClick(Self, sCode, sName) then
      begin
         mskJumin.Text := sCode;
         edtName.Text  := sName;
      end;
   end
   else if Sender = btnDate       then GF_CalendarClick(mskDate)
   else if Sender = btnRDate       then GF_CalendarClick(mskRDate)
   else if Sender = btnDAT_LAST   then GF_CalendarClick(mskDAT_LAST)
   else if Sender = btnDAT_RESULT then GF_CalendarClick(mskDAT_RESULT);

end;

procedure TfrmLDAI01.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      // 찾기기능 활성화
      btnFind.Enabled   := True;

      mskDate.Color     := $00E6FFFE;
      mskDate.Enabled   := True;
      btnDate.Enabled   := True;
      ckMiSokun.Enabled := True;
      mskJumin.Color    := clWindow;
      mskJumin.Enabled  := False;
      btnJumin.Enabled  := False;
      mskDate.SetFocus;
   end
   else if Sender = rbJumin then
   begin
      // 찾기기능 비활성화
      btnFind.Enabled   := False;

      mskDate.Color     := clWindow;
      mskDate.Enabled   := False;
      btnDate.Enabled   := False;
      ckMiSokun.Enabled := False;
      mskJumin.Color    := $00E6FFFE;
      mskJumin.Enabled  := True;
      btnJumin.Enabled  := True;
      mskJumin.SetFocus;
   end;
end;

procedure TfrmLDAI01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // Reference Key가 아니면 종료
   if Key <> GC_Refer then exit;

   if      Sender = mskDate       then UP_Click(btnDate)
   else if Sender = mskJumin      then UP_Click(btnJumin)
   else if Sender = mskDAT_LAST   then UP_Click(btnDAT_LAST)
   else if Sender = mskDAT_RESULT then UP_Click(btnDAT_RESULT);
end;

procedure TfrmLDAI01.UP_Exit(Sender: TObject);
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
   end;
end;


procedure TfrmLDAI01.btnFindClick(Sender: TObject);
begin
   inherited;

   frmLD1I07F := TfrmLD1I07F.Create(Self);
   frmLD1I07F.ShowModal;
end;

procedure TfrmLDAI01.btnDeleteClick(Sender: TObject);
begin
   inherited;

   // Delete Confirm Message
   if (grdMaster.Rows = 0) or (grdMaster.VisibleRows < 1) then exit;
   if not GF_MsgBox('Confirmation', '해당 조직검사 자료를 삭제하면 복구가 불가능합니다.' + Chr(13) +
                                    '정말로 삭제하시겠습니까 ?') then exit;

   // Delete 작업수행
   try
      with qryD_Cell do
      begin
         ParamByName('COD_BJJS').AsString := UV_vCod_bjjs[grdMaster.CurrentDataRow-1];
         ParamByName('NUM_ID').AsString   := UV_vNum_id[grdMaster.CurrentDataRow-1];
         ParamByName('COD_CELL').AsString := UV_vCod_cell[grdMaster.CurrentDataRow-1];

         ExecSql;
         GP_query_log(qryD_Cell, 'LD1I12', 'D', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end; 

   // 화면에서 안보이게함
   grdMaster.RowVisible[grdMaster.CurrentDataRow] := False;
   if grdMaster.VisibleRows < 1 then GP_FieldClear(pnlDetail);

   // Delete Mode & Msg
   UP_SetMode('DELETE');
end;

procedure TfrmLDAI01.edtCod_sokunExit(Sender: TObject);
begin
  inherited;
   qrySokun.Active := false;
   qrySokun.ParamByname('Cod_sokun').AsString := edtCod_sokun.text;
   qrySokun.Active := true;

    if  qrySokun.RecordCount > 0 then
        retDesc_sokun.lines.add(qrySokun.FieldByName('desc_sbsg').AsString +
                                qrySokun.FieldByName('desc_sbsg1').AsString +                                qrySokun.FieldByName('desc_sbsg2').AsString)
    else
         GF_MsgBox('Information', ' 없는 소견 코드입니다. ' + Chr(13) +
                                  ' 다시한번 확인하십시오');

    //edtCod_sokun.Text := '';
    retDesc_Sokun.Font.Color  := clBlack;
    retDesc_Sokun1.Font.Color := clBlack;

     if Uppercase(edtCod_sokun.Text) = 'PO01' then
       retDesc_sokun1.text :='자궁세포진 검사상 정상소견입니다. 자궁 경부암의 조기 발견을 위해 1년 간격의 정기적인 검사를 권유합니다. 이 결과는 "Screening Test" 소견 이므로 임상적으로 의심이 되는 병변이나 이상 증상이 관찰되시면 반드시 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO43' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 선상피내암종 소견이 관찰됩니다. 산부인과 진료를 통한 조직 검사와 적절한 치료가 반드시 필요합니다. 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO02' then
       retDesc_sokun1.text :='자궁세포진 검사상 양성 반응성세포와 함께 노인성 위축성세포가 관찰됩니다. 6개월 간격의 정기검진을 요합니다. '

     else if Uppercase(edtCod_sokun.Text) = 'PO03' then
       retDesc_sokun1.text :='자궁세포진 검사상 노인성 위축성 세포 및 염증세포가 관찰됩니다. 냉이 많거나 가려움증 등의 부인과적 증상이 있으면 산부인과 진찰을 요하며, 6개월이내 검사 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO04' then
       retDesc_sokun1.text :='자궁세포진 검사상 양성의 반응성 세포와 함께 다소 염증세포가 관찰됩니다. 염증 또는 외부자극에 의해 관찰될 수 있으니 6개월 간격의 정기적인 검사를 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO05' then
       retDesc_sokun1.text :='자궁세포진 검사상 반응성 세포 변화가 관찰됩니다. 염증 또는 외부자극에 의해 관찰될 수 있으니 6개월 간격으로 정기적인 검사를 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO06' then
       retDesc_sokun1.text :='자궁세포진 검사상 트리코모나스균 감염소견을 나타냅니다. 산부인과 진료를 통해 적절한 치료를 받으시고, 치료후 정기검진 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO07' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증세포와 함께 구균간선균이 관찰됩니다. 산부인과 진료를 통해 적절한 치료를 받으시고, 치료후 정기검진 받으시기 바랍니다.'
     
     else if Uppercase(edtCod_sokun.Text) = 'PO08' then
       retDesc_sokun1.text :='자궁세포진 검사상 캔디다류에 해당되는 진균 감염소견을 나타냅니다. 산부인과 진료를 통해 적절한 치료를 받으시고, 치료후 정기검진 받으시기 바랍니다.'
     
     else if Uppercase(edtCod_sokun.Text) = 'PO09' then
       retDesc_sokun1.text :='자궁세포진 검사상 방선균(Actinomyces) 종류로 보이는 세균 감염 소견입니다. 산부인과 진료를 통해 적절한 치료를 받으시고, 치료후 정기검진 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO10' then
       retDesc_sokun1.text :='자궁세포진 검사상 단순포진 바이러스 감염이 의심되는 세포변화가 관찰됩니다. 산부인과 전문의의 진찰 및 치료후 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO11' then
     Begin
       retDesc_sokun1.text :='자궁세포진 검사상 비정형편평상피세포(ASC-US)가 관찰됩니다. 산부인과 전문의의 진찰 및 3개월 내 재검하여 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO12' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 반응성으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 3-6개월 내 재검하여 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO13' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL) 으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO14' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 원반세포변화(koilocytotic change)를 동반한 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 인유두종바이러스 검사 (PCR for HPV)를 포함한 추적검사로 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO15' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL, 경도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료 의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO16' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL, HPV/경도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료 의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO17' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL/중등도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO18' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL/ 침윤이 의심되는 고도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO19' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급평평상피내병변 (HSIL/침윤이 의심되는 상피내암을 포함한 고도의 이형성세포)이 관찰됩니다. 산부인과 진료를 통한 조직검사와 적절한 치료가 필요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO20' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 편평상피세포암이 관찰됩니다. 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO21' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 비정형 선세포가 관찰됩니다. 산부인과 진료를 통한 재검사 또는 추적관찰이 필요합니다. 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO22' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 암이 의심되는 비정형 선세포가 관찰됩니다. 산부인과 진료를 통한 재검사가 필요합니다. 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO23' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 선암종 소견 관찰됩니다. 암 확진을 위해 산부인과 진료를 통한 조직 검사와 적절한 치료가 반드시 필요합니다. 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO24' then
       retDesc_sokun1.text :='자궁세포진 검사상 진단에 필요한 세포가 부족합니다. 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO25' then
       retDesc_sokun1.text :='자궁세포진 검사상 과다 혈액소견으로 진단에 부적절합니다. 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO26' then
       retDesc_sokun1.text :='자궁세포진 검사상 과다 염증세포소견으로 진단에 부적절합니다. 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO27' then
       retDesc_sokun1.text :='자궁세포진 검사상 도말된 세포가 건조되어 진단에 부적절합니다. 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO28' then
       retDesc_sokun1.text :='자궁세포진 검사상 세포가 두껍게 도말되어 진단에 부적절합니다. 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO29' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증소견과 함께 반응성 세포변화가 관찰됩니다. 자궁경부염 또는 질염의 가능성이 있으므로 냉이 많거나 가려움등의 증상이 있을 경우 염증치료 후 6개월 내 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO30' then
       retDesc_sokun1.text :='과각화증 세균성 질염이 의심됩니다.치료하시고 6개월후 재검사 하시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO31' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증을 동반한 반응성 세포변화와 함께 트리코모나스가 동반된 소견입니다. 산부인과의 진찰 및 적절한 치료를 요하며 6개월 내 추적검사로 경과 관찰 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO32' then
       retDesc_sokun1.text :='자궁세포진 검사상 양성 반응성 세포변화, 염증세포와 함께 구균간선균이 관찰됩니다. 산부인과의 진찰 및 적절한 치료를 하시고 치료 후 6개월 내 검사 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO33' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증을 동반한 반응성 세포변화와 함께 칸디다균으로 보이는 진균 감염이 관찰됩니다. 산부인과 진찰 및 적절한 치료를 받으시고, 치료 후 6개월 내 재검 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO34' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증을 동반한 반응성 세포변화와 함께 방선균으로 보이는 세균감염이 동반된 소견입니다. 산부인과의 진찰 및 적절한 치료를 요하며 6개월 내 추적검사로 경과 관찰 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO35' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HIS/ 중등도 또는 고도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여진료의뢰 받으시기 바랍니다.' ;
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO36' then
       retDesc_sokun1.text :='자궁세포진 검사상 정상소견입니다. 자궁 경부암의 조기 발견을 위해 1년 간격의 정기적인 검사를 권유합니다. 이 결과는 "Screening Test" 소견 이므로 임상적으로 의심이 되는 병변이나 이상 증상이 관찰되시면 반드시 재검사 받으시기 바랍니다.'

     else if (Uppercase(edtCod_sokun.Text) = 'PO37') or (Uppercase(edtCod_sokun.Text) = 'PO50') then
       retDesc_sokun1.text :='자궁세포진 검사상 약간의 위축성세포가 동반된 정상소견을 나타냅니다. 1년 간격으로 정기검진 받으시기 바랍니다.'

     else if (Uppercase(edtCod_sokun.Text) = 'PO38') or (Uppercase(edtCod_sokun.Text) = 'PO51') then
       retDesc_sokun1.text :='자궁세포진 검사상 약간의 위축성세포가 동반된 정상소견을 나타냅니다. 1년 간격으로 정기검진 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO39' then
     Begin
       retDesc_sokun1.text :='자궁세포진 검사상 비정형편평상피세포(ASC-US)가 관찰됩니다.산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO40' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 염증을 동반한 반응성(양성)으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 진료와 함께 3~6개월 내 재검사 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO41' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL)을 배제하기어려운 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO42' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내 병변(HSIL)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end
     else if Uppercase(edtCod_sokun.Text) = 'PO45' then
       retDesc_sokun1.text :='자궁세포진 검사상 수복세포를 포함한 반응성 세포변화가 관찰됩니다.염증 또는 외부자극에 의할 수 있으나 6개월 내 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO46' then
     begin
       retDesc_sokun1.text :='자궁세포진 검사상 경도 또는 중등도 이형성 세포(CIN 1-2)가 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.';
       retDesc_Sokun.Font.Color := clRed;
       retDesc_Sokun1.Font.Color := clRed;
     end;

end;

procedure TfrmLDAI01.edtCod_Pan6Click(Sender: TObject);
begin
     inherited;
     If edtCod_Pan6.Checked Then
        Begin
           edtDesc_Pan.Visible := True;
        End
     Else edtDesc_Pan.Visible := False;
     UV_bEdit := True;
end;

procedure TfrmLDAI01.cmbRelationChange(Sender: TObject);
begin
   inherited;
   if cmbRelation.ItemIndex = 0 then
   begin
      Application.CreateForm(TfrmLD1I071, frmLD1I071);
      frmLD1I071.Show;
   end;
end;

procedure TfrmLDAI01.btn_signClick(Sender: TObject);
Var
   index, Max_Num : integer;
begin
  inherited;
  index := grdMaster.CurrentDataRow - 1;
    { If b_sgCAppSet = False Then              //인증로그인이 안되어 있으면
      begin
         KMIConnect();
         GF_PubCertifyLoad();      //ysys
      end; 

      if not UV_bEdit then
      begin
         showmessage('변경된 내용이 없습니다.');
         exit;
      end; }

      If (b_sgCAppSet = True) and (Index >= 0) Then
      Begin
         //저장  프로시져 호출
         UF_Save;
         with Qry_CK_Gulkwa do
         begin
            Close;
            Qry_CK_Gulkwa.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Index];
            Qry_CK_Gulkwa.ParamByName('Num_id').AsString   := UV_vNum_id[index];
            Qry_CK_Gulkwa.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[index];
            Qry_CK_Gulkwa.ParamByName('COD_CELL').AsString := UV_vCod_cell[index];
            Qry_CK_Gulkwa.ParamByName('COD_HM').AsString   := UV_vCod_hm[index];

            Open;
            if (Qry_CK_Gulkwa.IsEmpty = False) and (UV_vCod_hm[index] = 'P003') then
            begin
               with Qry_Sign do
               begin
                  Close;
                  ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString   := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'P003' then
                     ParamByName('Program_id').AsString := 'LD1I123';
                  open;
                  Max_Num := Qry_Sign.FieldByName('seq_no').AsInteger;
               end;


               GV_sGulkwa := 'cod_bjjs[' + FieldByName('cod_bjjs').AsString + ']' +
                             'num_id['   + FieldByName('num_id').AsString + ']' +
                             'cod_cell['   + FieldByName('cod_cell').AsString + ']' +
                             'cod_jisa[' + FieldByName('cod_jisa').AsString + ']' +
                             'dat_gmgn[' + FieldByName('dat_gmgn').AsString + ']' +
                             'cod_hm[' + FieldByName('cod_hm').AsString + ']' +
                             'gubn_jongru[' + FieldByName('gubn_jongru').AsString + ']' +
                             'desc_jongru[' + FieldByName('desc_jongru').AsString + ']' +
                             'desc_kgbl[' + FieldByName('desc_kgbl').AsString + ']' +
                             'gubn_class[' + FieldByName('gubn_class').AsString + ']' +
                             'ysno_virus1[' + FieldByName('ysno_virus1').AsString + ']' +
                             'ysno_virus2[' + FieldByName('ysno_virus2').AsString + ']' +
                             'ysno_virus3[' + FieldByName('ysno_virus3').AsString + ']' +
                             'ysno_virus4[' + FieldByName('ysno_virus4').AsString + ']' +
                             'ysno_virus5[' + FieldByName('ysno_virus5').AsString + ']' +
                             'ysno_virus6[' + FieldByName('ysno_virus6').AsString + ']' +
                             'ysno_virus7[' + FieldByName('ysno_virus7').AsString + ']' +
                             'ysno_virus8[' + FieldByName('ysno_virus8').AsString + ']' +
                             'ysno_virus9[' + FieldByName('ysno_virus9').AsString + ']' +
                             'ysno_virus10[' + FieldByName('ysno_virus10').AsString + ']' +
                             'desc_virus[' + FieldByName('desc_virus').AsString + ']' +
                             'ysno_clin1[' + FieldByName('ysno_clin1').AsString + ']' +
                             'ysno_clin2[' + FieldByName('ysno_clin2').AsString + ']' +
                             'ysno_clin3[' + FieldByName('ysno_clin3').AsString + ']' +
                             'ysno_clin4[' + FieldByName('ysno_clin4').AsString + ']' +
                             'ysno_clin5[' + FieldByName('ysno_clin5').AsString + ']' +
                             'ysno_clin6[' + FieldByName('ysno_clin6').AsString + ']' +
                             'ysno_clin7[' + FieldByName('ysno_clin7').AsString + ']' +
                             'Desc_Gum[' + FieldByName('Desc_Gum').AsString + ']' +
                             'Gum_Chk[' + FieldByName('Gum_Chk').AsString + ']' +
                             'Gum_Remark[' + FieldByName('Gum_Remark').AsString + ']' +
                             'Gum_Form1[' + FieldByName('Gum_Form1').AsString + ']' +
                             'Gum_Form2[' + FieldByName('Gum_Form2').AsString + ']' +
                             'Gum_Form3[' + FieldByName('Gum_Form3').AsString + ']' +
                             'Gum_Form3_Etc[' + FieldByName('Gum_Form3_Etc').AsString + ']' +
                             'Sang_Form1[' + FieldByName('Sang_Form1').AsString + ']' +
                             'Sang_Form1_Gubn[' + FieldByName('Sang_Form1_Gubn').AsString + ']' +
                             'Sang_Form2[' + FieldByName('Sang_Form2').AsString + ']' +
                             'Sang_Form3[' + FieldByName('Sang_Form3').AsString + ']' +
                             'Sang_Form4[' + FieldByName('Sang_Form4').AsString + ']' +
                             'Sun_Form1[' + FieldByName('Sun_Form1').AsString + ']' +
                             'Sun_Form2[' + FieldByName('Sun_Form2').AsString + ']' +
                             'Sun_Form3[' + FieldByName('Sun_Form3').AsString + ']' +
                             'Sun_Form4[' + FieldByName('Sun_Form4').AsString + ']' +
                             'desc_bang[' + FieldByName('desc_bang').AsString + ']' +
                             'desc_buwi[' + FieldByName('desc_buwi').AsString + ']' +
                             'dat_last[' + FieldByName('dat_last').AsString + ']' +
                             'dat_result[' + FieldByName('dat_result').AsString + ']' +
                             'cod_sawon[' + FieldByName('cod_sawon').AsString + ']' +
                             'cod_doct[' + FieldByName('cod_doct').AsString + ']' +
                             'desc_sokun1[' + FieldByName('desc_sokun1').AsString + ']' +
                             'desc_sokun2[' + FieldByName('desc_sokun2').AsString + ']' +
                             'desc_sokun3[' + FieldByName('desc_sokun3').AsString + ']' +
                             'desc_sokun4[' + FieldByName('desc_sokun4').AsString + ']' +
                             'desc_sokun5[' + FieldByName('desc_sokun5').AsString + ']' +
                             'Cod_Pan[' + FieldByName('Cod_Pan').AsString + ']' +
                             'Desc_Pan[' + FieldByName('Desc_Pan').AsString + ']' +
                             'ysno_sokun[' + FieldByName('ysno_sokun').AsString + ']' +
                             'Check_Year[' + FieldByName('Check_Year').AsString + ']' +
                             'gubn_order[' + FieldByName('gubn_order').AsString + ']' +
                             'dat_insert[' + FieldByName('dat_insert').AsString + ']' +
                             'cod_insert[' + FieldByName('cod_insert').AsString + ']' +
                             'dat_update[' + FieldByName('dat_update').AsString + ']' +
                             'cod_update[' + FieldByName('cod_update').AsString + ']';


                             if UV_vDat_gmgn[Index] < '20180101' then
                             begin
                                GV_sGulkwa := GV_sGulkwa +
                                'cod_jisa_R[' + FieldByName('cod_jisa_R').AsString + ']' +
                                'num_id_R[' + FieldByName('num_id_R').AsString + ']' +
                                'dat_gmgn_R[' + FieldByName('dat_gmgn_R').AsString + ']';
                             end
                             else
                             begin
                                GV_sGulkwa := GV_sGulkwa +
                                'cod_jisa_R2[' + FieldByName('cod_jisa_R2018').AsString + ']' +
                                'num_id_R2[' + FieldByName('num_id_R2018').AsString + ']' +
                                'dat_gmgn_R2[' + FieldByName('dat_gmgn_R2018').AsString + ']';
                             end;



               if (UV_vCod_hm[index] = 'P003') and (UV_vDat_gmgn[Index] < '20180101') then
               begin
               GV_sGulkwa := GV_sGulkwa +
                             'can5_panjung[' + FieldByName('can5_panjung').AsString + ']' +
                             'can5_panjung_etc[' + FieldByName('can5_panjung_etc').AsString + ']' +
                             'can5_advice[' + FieldByName('can5_advice').AsString + ']' +
                             'can5_advice1[' + FieldByName('can5_advice1').AsString + ']';

               if UV_vCod_Jisa[Index] ='15' then
                  GV_sGulkwa := GV_sGulkwa + 'can5_doctor[' + FieldByName('can5_doctor').AsString + ']' ;
               GV_sGulkwa := GV_sGulkwa + 'can5_pan_date[' + FieldByName('can5_pan_date').AsString + ']';
               end
               else if (UV_vCod_hm[index] = 'P003') and (UV_vDat_gmgn[Index] >= '20180101') then
               begin  
               GV_sGulkwa := GV_sGulkwa +
                             'can5_panjung[' + FieldByName('can5_panjung2018').AsString + ']' +
                             'can5_panjung_etc[' + FieldByName('can5_panjung_etc2018').AsString + ']' +
                             'can5_advice[' + FieldByName('can5_advice2018').AsString + ']' +
                             'can5_advice1[' + FieldByName('can5_advice12018').AsString + ']';

               if UV_vCod_Jisa[Index] ='15' then
                  GV_sGulkwa := GV_sGulkwa + 'can5_doctor[' + FieldByName('can5_doctor2018').AsString + ']' ;
               GV_sGulkwa := GV_sGulkwa + 'can5_pan_date[' + FieldByName('can5_pan_date2018').AsString + ']';
               end;



               if UV_vDat_gmgn[Index] < '20180101' then
               begin
                  GV_sGulkwa := GV_sGulkwa +
                             'dat_insert_R[' + FieldByName('dat_insert_R').AsString + ']' +
                             'cod_insert_R[' + FieldByName('cod_insert_R').AsString + ']' +
                             'dat_update_R[' + FieldByName('dat_update_R').AsString + ']' +
                             'cod_update_R[' + FieldByName('cod_update_R').AsString + ']';
               end
               else
               begin
                  GV_sGulkwa := GV_sGulkwa +
                             'dat_insert_R2[' + FieldByName('dat_insert_R2018').AsString + ']' +
                             'cod_insert_R2[' + FieldByName('cod_insert_R2018').AsString + ']' +
                             'dat_update_R2[' + FieldByName('dat_update_R2018').AsString + ']' +
                             'cod_update_R2[' + FieldByName('cod_update_R2018').AsString + ']';
               end;
               GF_PubSignCertify;

               with QryI_Sign do
               begin
                  ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString     := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'P003' then
                     ParamByName('Program_id').AsString := 'LD1I123';
                  ParamByName('Seq_no').Asinteger    := Max_Num + 1;
                  ParamByName('Result').AsMemo       := GV_sGulkwa;
                  ParamByName('Sign_Key').AsMemo   := gv_strSignedMsg;
                  ParamByName('cod_sawon').AsString  := GV_sUserID;
                  Execsql;
                  GP_query_log(QryI_Sign, 'LD1I12', 'I', 'Y');
               end;
            end
            else
            begin
               ShowMessage('서명저장할 내용이 없습니다.');
               Exit;
            end;
         end;
      end
      Else If (b_sgCAppSet = False) then //처음 로그인 할때 인증서 로그인을 안했을경우
      Begin
           ShowMessage('서명 저장을 다시 하시기 바랍니다.');
           Exit;
      end
      Else If Index < 0 then
      Begin
           ShowMessage('조회된 내용이 없습니다. 조회후 서명저장하시기 바랍니다.');
           Exit;
      end;

      //저장이 완료된후에는 글로벌 변수 초기화
      gv_strSignDate  := '';
      gv_strSignedMsg := '';
      GV_sGulkwa      := '';
end;

function TfrmLDAI01.UF_B201_Jong1(bData, bData_H, bData_M, bData_L : Boolean) : String;
begin
   if bData then
   begin
      if bData_H then Result := '1'
      else if bData_M then Result := '2'
      else if bData_L then Result := '3'
      else Result := '1';
   end
   else          Result := '0';
end;
function TfrmLDAI01.UF_B201_Jong4(bData, bData_H, bData_L : Boolean) : String;
begin
   if bData then
   begin
      if bData_H then Result := '1'
      else if bData_L then Result := '2'
      else  Result := '1';
   end
   else          Result := '0';
end;
function TfrmLDAI01.UF_B201(bData : Boolean) : String;
begin
   if bData then Result := '1'
   else          Result := '0';
end;
Function TfrmLDAI01.UF_Check(iTemp : Integer) : Integer;
begin
   if iTemp >= 0 then Result := iTemp
   else               Result := 0;
end;
Function TfrmLDAI01.UF_Check01(iTemp : Integer) : String;
begin
   if (iTemp >= 0) and
      (iTemp < 10) then      Result := '0' + intToStr(iTemp)
   else if iTemp >= 10 then  Result := intToStr(iTemp)
   else                      Result := '0';
end;

procedure TfrmLDAI01.cmbGUBN_CLASSExit(Sender: TObject);
begin
  inherited;
  UP_Change(cmbGUBN_CLASS);
end;

procedure TfrmLDAI01.cmbGubnChange(Sender: TObject);
begin
  inherited;
   if cmbGubn.ItemIndex = 16 then
      Label11.Caption := 'P013[판정관리]'
   else
      Label11.Caption := 'P012[판정관리]';
end;

procedure TfrmLDAI01.CmbSang_Form1_gubnKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['1','2',#8, #13]) then
  begin
     showmessage('입력 값을 다시 확인해주십시오.');
     key := #0;
  end;
end;

procedure TfrmLDAI01.btnPrintClick(Sender: TObject);
begin
  inherited;
    Application.CreateForm(TfrmLD1P13, frmLD1P13);
    frmLD1P13.Show;
end;

procedure TfrmLDAI01.Button1Click(Sender: TObject);
var i : Integer;
sWhere,sOrderBy,a,b,c : String;
begin
  inherited;
   with qryCell do
   begin
      Active := False;

      sWhere := 'WHERE ';

      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''';

         if Chk_order.Checked then sWhere := sWhere + 'and (A.gubn_order = ''Y'' or A.check_jindan = ''Y'')';
      end
      else if rbRDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''';

         if Chk_order.Checked then sWhere := sWhere + 'and (A.gubn_order = ''Y'' or A.check_jindan = ''Y'')';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' B.num_jumin = ''' + mskJumin.Text + '''';
      end;

      // 검사구분에 따라서 Where 조건추가
      if      cmbGubn.ItemIndex = 0  then sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''P'''
      else if cmbGubn.ItemIndex = 1  then sWhere := sWhere + ' AND cod_hm = ''B008'''
      else if cmbGubn.ItemIndex = 2  then sWhere := sWhere + ' AND cod_hm = ''B009'''
      else if cmbGubn.ItemIndex = 3 then sWhere := sWhere + ' AND cod_hm = ''P003'''
      else if cmbGubn.ItemIndex = 4 then sWhere := sWhere + ' AND cod_hm = ''P006'''
      else if cmbGubn.ItemIndex = 5 then sWhere := sWhere + ' AND cod_hm = ''P010'''
      else if cmbGubn.ItemIndex = 6 then sWhere := sWhere + ' AND cod_hm = ''P011'''
      else if cmbGubn.ItemIndex = 7 then sWhere := sWhere + ' AND cod_hm = ''P012'''
      else if cmbGubn.ItemIndex = 8 then sWhere := sWhere + ' AND cod_hm = ''P013''';

      if cmbOrder.ItemIndex = 1 then
         sOrderBy := ' Order by A.desc_buwi '
      else
         sOrderBy := ' ORDER BY A.dat_gmgn, A.cod_jisa, B.desc_name ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
      Active := True;

      if qryCell.IsEmpty = False then
      begin
      with qry_slide do
      begin
      qry_slide.Active := False;
      qry_slide.ParamByName('DAT_GMGN').AsString  := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];
      qry_slide.ParamByName('num_id').AsString    := UV_vNum_id[grdMaster.CurrentDataRow-1];
      qry_slide.ParamByName('desc_buwi').AsString := UV_vDesc_buwi[grdMaster.CurrentDataRow-1];
      qry_slide.Active := True;
      if qry_slide.RecordCount > 0 then
      begin
          with qryU_slide do
          begin
             ParamByName('desc_buwi').AsString := edtDESC_BUWI.Text;
             ParamByName('num_id').AsString    := UV_vNum_id[grdMaster.CurrentDataRow-1];
             ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];

             If btn_rent.Checked Then
             begin
                ParamByName('rent_yn').AsString := 'Y';
                ParamByName('DAT_RENT').AsString    := msk_dat_rent.Text;
                ParamByName('DAT_RETURN').AsString  := qry_slide.fieldbyname('dat_return').AsString;
                ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
             end
             Else If btn_return.Checked Then
             begin
                ParamByName('rent_yn').AsString := 'N';
                ParamByName('DAT_RENT').AsString    := qry_slide.fieldbyname('dat_rent').AsString;
                ParamByName('DAT_RETURN').AsString  := msk_dat_return.text;
                ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
             end
             else
             begin
               GF_MsgBox('Warning', '대여상태를 입력해주세요');
               exit;
             end;
             Execsql;
             GP_query_log(qryU_slide, 'LDAI01', 'U', 'Y');
          end;
      end
      else
      begin
            with qryI_slide do
            begin
               ParamByName('desc_buwi').AsString := edtDESC_BUWI.Text;
               ParamByName('num_id').AsString    := UV_vNum_id[grdMaster.CurrentDataRow-1];
               ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[grdMaster.CurrentDataRow-1]; 
               If btn_rent.Checked Then
               begin
                  ParamByName('rent_yn').AsString := 'Y';
                  ParamByName('DAT_RENT').AsString    := msk_dat_rent.Text;
                  ParamByName('DAT_RETURN').AsString  := qry_slide.fieldbyname('dat_return').AsString;
                  ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
                  ParamByName('cod_jisa').AsString    := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
               end
               Else If btn_return.Checked Then
               begin
                  ParamByName('rent_yn').AsString := 'N';
                  ParamByName('DAT_RENT').AsString    := qry_slide.fieldbyname('dat_rent').AsString;
                  ParamByName('DAT_RETURN').AsString  := msk_dat_return.text;
                  ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
                  ParamByName('cod_jisa').AsString    := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
               end
               else
               begin
                GF_MsgBox('Warning', '대여상태를 입력해주세요');
               end;
               Execsql;
               GP_query_log(qryI_slide, 'LDAI01', 'I', 'Y');
            end;
      end;
      END;
      {except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end; }
   end;
end;
end;

procedure TfrmLDAI01.rbRDateClick(Sender: TObject);
begin
  inherited;
  inherited;
   if Sender = rbRDate then
   begin
      // 찾기기능 활성화
      btnFind.Enabled   := True;

      mskRDate.Color     := $00E6FFFE;
      mskRDate.Enabled   := True;
      btnRDate.Enabled   := True;
      ckMiSokun.Enabled := True;
      mskJumin.Color    := clWindow;
      mskJumin.Enabled  := False;
      btnJumin.Enabled  := False;
      mskRDate.SetFocus;
   end
   else if Sender = rbJumin then
   begin
      // 찾기기능 비활성화
      btnFind.Enabled   := False;

      mskRDate.Color     := clWindow;
      mskRDate.Enabled   := False;
      btnRDate.Enabled   := False;
      ckMiSokun.Enabled := False;
      mskJumin.Color    := $00E6FFFE;
      mskJumin.Enabled  := True;
      btnJumin.Enabled  := True;
      mskJumin.SetFocus;
   end;
end;

procedure TfrmLDAI01.btn_del_numClick(Sender: TObject);
var i : Integer;
begin
  inherited;

   with qry_DelNumber do
   begin
      close;
       ParamByName('cod_jisa').AsString := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
       ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];
       ParamByName('num_id').AsString := UV_vNum_id[grdMaster.CurrentDataRow-1];
       ParamByName('cod_hm').AsString := UV_vCod_hm[grdMaster.CurrentDataRow-1];
      open;
      if qry_DelNumber.recordcount = 0 then
      begin
            with qryI_DelNumber do
            begin

                  ParamByName('cod_jisa').AsString := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
                  ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];
                  ParamByName('num_id').AsString := UV_vNum_id[grdMaster.CurrentDataRow-1];
                  ParamByName('cod_hm').AsString := UV_vCod_hm[grdMaster.CurrentDataRow-1];
                  ParamByName('cod_insert').AsString := GV_sUserId;
                  ParamByName('dat_insert').AsString := GV_sToday;

                  if (edt_del_num.Text <> '') and (copy(edt_del_num.Text,1,1) <> 'D') then ParamByName('num_delete').AsString := 'D('+edt_del_num.Text+')'
                  else if edt_del_num.Text = '' then
                  begin
                    if not GF_MsgBox('Confirmation','접수번호를 초기화 하시겠습니까?') then exit;
                    ParamByName('num_delete').AsString := '';
                  end
                  else if  edt_del_num.Text <> '' then
                  begin
                      if not GF_MsgBox('Confirmation', edt_del_num.Text + '로 저장하시겠습니까?') then exit;
                       ParamByName('num_delete').AsString := edt_del_num.Text
                  end;

                  if (edt_num1.Checked) or (edt_num2.Checked) or (edt_num3.Checked) then
                  begin
                       if edt_num1.Checked then ParamByName('desc_delete').AsString := '이중접수'
                       else if edt_num2.Checked then ParamByName('desc_delete').AsString := '수정건 삭제'
                       else if edt_num3.Checked then ParamByName('desc_delete').AsString := '기타:' + edt_kita.Text;
                  end
                  else
                  begin
                       if not GF_MsgBox ('Warning', '사유를 초기화 하시겠습니까?') then exit;
                       ParamByName('desc_delete').AsString := '';
                  end;

                  Execsql;
                  GP_query_log(qryU_DelNumber, 'LDAI01', 'Q', 'Y');
            end;
      end
      else
      begin
          with qryU_DelNumber do
          begin
                ParamByName('cod_jisa').AsString := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
                ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];
                ParamByName('num_id').AsString := UV_vNum_id[grdMaster.CurrentDataRow-1];
                ParamByName('cod_hm').AsString := UV_vCod_hm[grdMaster.CurrentDataRow-1];
                ParamByName('cod_update').AsString := GV_sUserId;
                ParamByName('dat_update').AsString := GV_sToday;

                if (edt_del_num.Text <> '') and (copy(edt_del_num.Text,1,1) <> 'D') then ParamByName('num_delete').AsString := 'D('+edt_del_num.Text+')'
                else if edt_del_num.Text = '' then
                begin
                  if not GF_MsgBox('Confirmation','접수번호를 초기화 하시겠습니까?') then exit;
                  ParamByName('num_delete').AsString := '';
                end
                else if  edt_del_num.Text <> '' then
                begin
                    if not GF_MsgBox('Confirmation', edt_del_num.Text + '로 저장하시겠습니까?') then exit;
                     ParamByName('num_delete').AsString := edt_del_num.Text
                end;

                if (edt_num1.Checked) or (edt_num2.Checked) or (edt_num3.Checked) then
                begin
                     if edt_num1.Checked then ParamByName('desc_delete').AsString := '이중접수'
                     else if edt_num2.Checked then ParamByName('desc_delete').AsString := '수정건 삭제'
                     else if edt_num3.Checked then ParamByName('desc_delete').AsString := '기타:' + edt_kita.Text;
                end
                else
                begin
                     if not GF_MsgBox ('Warning', '사유를 초기화 하시겠습니까?') then exit;
                     ParamByName('desc_delete').AsString := '';
                end;

                Execsql;
                GP_query_log(qryU_DelNumber, 'LDAI01', 'Q', 'Y');
          end;
      end;
      end;
end;
procedure TfrmLDAI01.edt_num3Click(Sender: TObject);
begin
  inherited;
     If edt_num3.Checked Then
        Begin
           edt_kita.Visible := True;
        End
     Else edt_kita.Visible := False;
     UV_bEdit := True;
end;

end.
