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

unit LDAI02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLDAI02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    pnlDetail: TPanel;
    rbDate: TRadioButton;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlJisa: TPanel;
    Label2: TLabel;
    cmbJisa: TComboBox;
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
    Label1: TLabel;
    ckMiSokun: TCheckBox;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    edtDESC_BANG: TEdit;
    edtDESC_BUWI: TEdit;
    edtCod_sokun: TEdit;
    retDesc_Sokun: TMemo;
    qrySokun: TQuery;
    qryCell: TQuery;
    qryGmgn: TQuery;
    qryU_Cell: TQuery;
    qryD_Cell: TQuery;
    qryUser_priv: TQuery;
    qryCommon: TQuery;
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
    Pnl_Can1: TPanel;
    Label72: TLabel;
    Label73: TLabel;
    Panel13: TPanel;
    GroupBox8: TGroupBox;
    Can1_Bung3_Jong1: TCheckBox;
    Can1_Bung3_Jong4: TCheckBox;
    Can1_Bung3_Jong7: TCheckBox;
    Can1_Bung3_Jong2: TCheckBox;
    Can1_Bung3_Jong5: TCheckBox;
    Can1_Bung3_Jong8: TCheckBox;
    Can1_Bung3_Jong3: TCheckBox;
    Can1_Bung3_Jong6: TCheckBox;
    Can1_Bung3_Jong9: TCheckBox;
    Can1_Bung3_Jong10: TCheckBox;
    Can1_Bung3_Jong11: TCheckBox;
    Can1_Bung3_Jong_Direct: TEdit;
    GroupBox9: TGroupBox;
    Can1_Bung3_Etc1: TCheckBox;
    Can1_Bung3_Etc6: TCheckBox;
    Can1_Bung3_Etc4: TCheckBox;
    Can1_Bung3_Etc3: TCheckBox;
    Can1_Bung3_Etc8: TCheckBox;
    Can1_Bung3_Etc2: TCheckBox;
    Can1_Bung3_Etc7: TCheckBox;
    Can1_Bung3_Etc_Direct: TEdit;
    Can1_Bung3_Etc5: TCheckBox;
    Can1_Bung3_Jindan: TComboBox;
    Can1_Bung3_Count: TComboBox;
    qryI_Ca_Cancer2009_Can1: TQuery;
    qryU_ca_cancer2009_Can1: TQuery;
    Pnl_Can2: TPanel;
    Label74: TLabel;
    Label75: TLabel;
    Panel48: TPanel;
    GroupBox10: TGroupBox;
    Can2_Bung4_Jong1: TCheckBox;
    Can2_Bung4_Jong4: TCheckBox;
    Can2_Bung4_Jong2: TCheckBox;
    Can2_Bung4_Jong5: TCheckBox;
    Can2_Bung4_Jong8: TCheckBox;
    Can2_Bung4_Jong3: TCheckBox;
    Can2_Bung4_Jong9: TCheckBox;
    Can2_Bung4_Jong10: TCheckBox;
    Can2_Bung4_Jong11: TCheckBox;
    Can2_Bung4_Jong_Direct: TEdit;
    Can2_Bung4_Jong7: TCheckBox;
    Can2_Bung4_Jong6: TCheckBox;
    GroupBox11: TGroupBox;
    Can2_Bung4_Etc1: TCheckBox;
    Can2_Bung4_Etc5: TCheckBox;
    Can2_Bung4_Etc4: TCheckBox;
    Can2_Bung4_Etc2: TCheckBox;
    Can2_Bung4_Etc_Direct: TEdit;
    Can2_Bung4_Etc3: TCheckBox;
    Can2_Bung4_Jindan: TComboBox;
    Can2_Bung4_Count: TComboBox;
    qryI_Ca_Cancer2009_Can2: TQuery;
    qryU_ca_cancer2009_Can2: TQuery;
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
    GroupBox12: TGroupBox;
    GroupBox13: TGroupBox;
    GroupBox14: TGroupBox;
    Can1_Bung3_Jong1_H: TRadioButton;
    Can1_Bung3_Jong1_M: TRadioButton;
    Can1_Bung3_Jong1_L: TRadioButton;
    Can1_Bung3_Jong4_H: TRadioButton;
    Can1_Bung3_Jong4_l: TRadioButton;
    GroupBox15: TGroupBox;
    Can2_Bung4_Jong1_H: TRadioButton;
    Can2_Bung4_Jong1_M: TRadioButton;
    Can2_Bung4_Jong1_L: TRadioButton;
    qryU_ca_cancer2009_Can5_15: TQuery;
    qryU_ca_cancer2009_Can5: TQuery;
    qryI_Ca_Cancer2009_Can5: TQuery;
    PnlHide2: TPanel;
    Panel17: TPanel;
    Label13: TLabel;
    PnlHide3: TPanel;
    Panel19: TPanel;
    Label14: TLabel;
    GroupBox6: TGroupBox;
    Memo_Psokun: TMemo;
    qryPsokun: TQuery;
    qryPCa_cancer: TQuery;
    edt_pdat_gmgn: TDateEdit;
    Panel2: TPanel;
    Panel6: TPanel;
    edt_jindan: TEdit;
    edt_jindan_count: TEdit;
    GroupBox1: TGroupBox;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Lab_slide: TLabel;
    Label20: TLabel;
    msk_dat_rent: TMaskEdit;
    btn_rent: TRadioButton;
    btn_return: TRadioButton;
    msk_dat_return: TMaskEdit;
    Button1: TButton;
    Edt_cmt: TMaskEdit;
    Panel15: TPanel;
    Label3: TLabel;
    Edt_floor: TEdit;
    qryI_slide: TQuery;
    qry_slide: TQuery;
    qryU_slide: TQuery;
    qryFloor: TQuery;
    pnlHide4: TPanel;
    Panel20: TPanel;
    Label4: TLabel;
    Can1_Bung3_Jindan_2018: TComboBox;
    qryCa_cancer2018: TQuery;
    qryU_ca_cancer2018_Can1: TQuery;
    qryU_ca_cancer2018_Can2: TQuery;
    qryU_ca_cancer2018_Can5: TQuery;
    qryU_ca_cancer2018_Can5_15: TQuery;
    qryI_Ca_Cancer2018_Can1: TQuery;
    qryI_Ca_Cancer2018_Can2: TQuery;
    qryI_Ca_Cancer2018_Can5: TQuery;
    qryI_Ca_Cancer2018_Can5_15: TQuery;
    Can2_Bung4_Jindan_2018: TComboBox;
    Qry_CK_Gulkwa_2018: TQuery;
    qryPCa_cancer2018: TQuery;
    edt_buwi: TEdit;
    qrySawon: TQuery;
    edt_pDoctor: TEdit;
    rbCell: TRadioButton;
    mskcell: TMaskEdit;
    edt_data: TEdit;
    Label5: TLabel;
    cmb_level: TComboBox;
    mskdat_time: TDateEdit;
    edt_time_S: TEdit;
    edt_time_R: TEdit;
    mskdat_time_r: TDateEdit;
    Panel18: TPanel;
    cmbCOD_jindan_DOCT: TComboBox;
    chk_jindan: TCheckBox;
    rbRDate: TRadioButton;
    mskRDate: TDateEdit;
    btnRDate: TSpeedButton;
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
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure cmbRelationChange(Sender: TObject);
    procedure btn_signClick(Sender: TObject);
    procedure CmbSang_Form1_gubnKeyPress(Sender: TObject; var Key: Char);
    procedure btnPrintClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure retDesc_SokunKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure rbRDateClick(Sender: TObject);


  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,  UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vCod_cell, UV_vPsokun, UV_vFloor, UV_vRent, UV_vCmt, UV_vDat_Rent, UV_vDAT_Return,
    UV_vCod_hm, UV_vGubn_jongru, UV_vDesc_jongru, UV_vGubn_class, UV_vYsno_virus1,
    UV_vYsno_virus2, UV_vYsno_virus3, UV_vYsno_virus4, UV_vYsno_virus5, UV_vYsno_Virus6, UV_vYsno_Virus7, UV_vYsno_Virus8,
    UV_vDesc_virus, UV_vYsno_clin1, UV_vYsno_clin2, UV_vYsno_clin3, UV_vYsno_clin4,
    UV_vYsno_clin5, UV_vYsno_clin6, UV_vYsno_clin7, UV_vDesc_bang, UV_vDesc_buwi,
    UV_vCod_sawon, UV_vCod_doct, UV_vCod_Jindan_doct, UV_vCheck_jindan, UV_vDat_last, UV_vDat_Time, UV_vDat_result, UV_vDat_Time_R, UV_vDesc_kgbl,
    UV_vDesc_sokun, UV_vDesc_sokun1, UV_vCod_jisa, UV_vDesc_Gum, UV_vGum_Chk, UV_vGum_Remark,
    UV_vGum_Form1, UV_vGum_Form2, UV_vGum_Form3, UV_vGum_Form3_Etc,
    UV_vSang_Form1, UV_vSang_Form1_Gubn, UV_vSang_Form2, UV_vSang_Form3, UV_vSang_Form4,
    UV_vSun_Form1, UV_vSun_Form2, UV_vSun_Form3, UV_vSun_Form4, UV_vCod_Pan, UV_vDesc_Pan, UV_vYsno_sokun, UV_vCheck_Year, UV_vCell_level,
    UV_vGubn_P012, UV_vGubn_pan, UV_vCod_sokun, UV_vDoctor : Variant;

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
    function UF_insertSokun(sokun : string) : String ;
    function UF_insertSokun_drag(data: string) : String ;
    function GetCurrLine(Memo : TMemo) : integer;
  public
    { Public declarations }
  end;

var
  frmLDAI02: TfrmLDAI02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     LD1I07F, LD1I071, LD1P12, LDAI012, LDAI021;//, LD1P121;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLDAI02.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLDAI02.UP_GridSet;
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
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '주민번호';
         Col[3].Heading   := '검진일자';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taCenter;
         Col[3].Color     := clBtnFace;
      end;

      Col[4].Heading := '접수 No';
      Col[5].Heading := '조직코드';
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

procedure TfrmLDAI02.UP_VarMemSet(iLength : Integer);
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
      UV_vDesc_bang   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sawon   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jindan_doct := VarArrayCreate([0, 0], varOleStr);
      UV_vCheck_jindan:= VarArrayCreate([0, 0], varOleStr);
      UV_vDat_last    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Time    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_result  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Time_R  := VarArrayCreate([0, 0], varOleStr);
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
      UV_vCell_level  := VarArrayCreate([0, 0], varOleStr);
      UV_vDoctor      := VarArrayCreate([0, 0], varOleStr);
      UV_vPsokun      := VarArrayCreate([0, 0], varOleStr);
      UV_vFloor       := VarArrayCreate([0, 0], varOleStr);
      UV_vRent        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Rent    := VarArrayCreate([0, 0], varOleStr);
      UV_vDAT_Return  := VarArrayCreate([0, 0], varOleStr);
      UV_vCmt         := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_vDesc_bang,   iLength);
      VarArrayRedim(UV_vDesc_buwi,   iLength);
      VarArrayRedim(UV_vCod_sawon,   iLength);
      VarArrayRedim(UV_vCod_doct,    iLength);
      VarArrayRedim(UV_vCod_jindan_doct,    iLength);
      VarArrayRedim(UV_vcheck_jindan,    iLength);
      VarArrayRedim(UV_vDat_last,    iLength);
      VarArrayRedim(UV_vDat_time,    iLength);
      VarArrayRedim(UV_vDat_result,  iLength);
      VarArrayRedim(UV_vDat_time_R,    iLength);
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
      VarArrayRedim(UV_vCell_level,  iLength);
      VarArrayRedim(UV_vDoctor,      iLength);
      VarArrayRedim(UV_vPsokun,      iLength);
      VarArrayRedim(UV_vFloor,       iLength);
      VarArrayRedim(UV_vRent,        iLength);
      VarArrayRedim(UV_vCmt,         iLength);
      VarArrayRedim(UV_vDat_Rent,    iLength);
      VarArrayRedim(UV_vDAT_Return,  iLength);
   end;
end;

function TfrmLDAI02.UF_Condition : Boolean;
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

function TfrmLDAI02.UF_Save : Boolean;
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
      with qryU_Cell do
      begin
         ParamByName('DESC_BANG').AsString        := edtDESC_BANG.Text;
         ParamByName('DESC_BUWI').AsString        := edtDESC_BUWI.Text;
         ParamByName('COD_SAWON').AsString        := Copy(cmbCOD_SAWON.Text, 1, 6);
         ParamByName('COD_DOCT').AsString         := Copy(cmbCOD_DOCT.Text, 1, 6);
         //외부 입력 의사 추가
         If chk_jindan.Checked Then
         ParamByName('check_jindan').AsString      := 'Y'
         else ParamByName('check_jindan').AsString := 'N';
         ParamByName('COD_jindan_doctor').AsString  := Copy(cmbCOD_Jindan_DOCT.Text, 1, 6);

         //검체접수일 - 처음 저장 때, 최초 입력 값
         If mskDAT_LAST.Text = '' then ParamByName('DAT_LAST').AsString  := GV_sToday
         else ParamByName('DAT_LAST').AsString    := mskDAT_LAST.Text;
         If edt_time_S.Text = '' then ParamByName('dat_time').AsString := formatdatetime('HH:NN:SS', now)
         else  ParamByName('dat_time').AsString := edt_time_S.Text;
         //결과입력일 - 서명 저장 시에 저장됨
         ParamByName('DAT_RESULT').AsString       := mskDAT_RESULT.Text;
         //ParamByName('dat_time_R').AsString := formatdatetime('HH:NN:SS', now);

         ParamByName('cell_level').AsString       := copy(cmb_level.Text, 1, 1) ;

         ParamByName('DESC_KGBL').AsString        := retDESC_KGBL.Text;
         ParamByName('DESC_SOKUN1').AsString      := copy(retDESC_SOKUN.Text,1,250);//copy((StringReplace(retDESC_SOKUN.Text,''#$D#$A'','',[rfReplaceAll])),1,250);   //20141119 곽윤설
         ParamByName('DESC_SOKUN2').AsString      := copy(retDESC_SOKUN.Text,251,250);
         ParamByName('DESC_SOKUN3').AsString      := copy(retDESC_SOKUN.Text,501,250);
         ParamByName('DESC_SOKUN4').AsString      := copy(retDESC_SOKUN.Text,751,250);
         ParamByName('DESC_SOKUN5').AsMemo        := copy(retDESC_SOKUN.Text,1001,length(retDESC_SOKUN.Text) - 1000);
         // 결과일자에 따라서 소견유무 표시 (Y:판독중, C:판독완료, N: 미판독)
         if Trim(mskDAT_RESULT.Text) <> '' then
         begin
            ParamByName('YSNO_SOKUN').AsString  := 'Y';
            ParamByName('dat_time_R').AsString  := UV_vDat_time_R[i];
         end
         else
         begin
            ParamByName('YSNO_SOKUN').AsString  := 'N';
            ParamByName('dat_time_R').AsString  := '';
         end;
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

      //2018 공단 위장, 대장 내시경 관련 변경 ..20180103 수지
      If UV_vDat_gmgn[i] < '20180101' then
      begin
      qryCa_cancer2009.Active := false;
      qryCa_cancer2009.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[i];
      qryCa_cancer2009.ParamByName('Num_id').AsString   := UV_vNum_id[i];
      qryCa_cancer2009.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[i];
      qryCa_cancer2009.Active := True;
      if qryCa_cancer2009.IsEmpty = False then
      begin
         //위조직
         if UV_vCod_hm[i] = 'B001' then
         begin
            with qryU_ca_cancer2009_Can1 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               ParamByName('can1_bung3_jindan').AsInteger := UF_Check(Can1_Bung3_Jindan.ItemIndex);
               ParamByName('can1_bung3_count').AsInteger  := UF_Check(Can1_Bung3_Count.ItemIndex);

               ParamByName('can1_bung3_jong').AsString := UF_B201_Jong1(can1_bung3_jong1.Checked, can1_bung3_jong1_H.Checked, can1_bung3_jong1_M.Checked, can1_bung3_jong1_L.Checked) +
                                                          UF_B201(can1_bung3_jong2.Checked) +
                                                          UF_B201(can1_bung3_jong3.Checked) +
                                                          UF_B201_Jong4(can1_bung3_jong4.Checked, can1_bung3_jong4_H.Checked, can1_bung3_jong4_L.Checked) +
                                                          UF_B201(can1_bung3_jong5.Checked) +
                                                          UF_B201(can1_bung3_jong6.Checked) +
                                                          UF_B201(can1_bung3_jong7.Checked) +
                                                          UF_B201(can1_bung3_jong8.Checked) +
                                                          UF_B201(can1_bung3_jong9.Checked) +
                                                          UF_B201(can1_bung3_jong10.Checked) +
                                                          UF_B201(can1_bung3_jong11.Checked);
               if can1_bung3_jong11.Checked then ParamByName('can1_bung3_jong_direct').AsString := GF_CutString(can1_bung3_jong_direct.Text, 20)
               else                              ParamByName('can1_bung3_jong_direct').AsString := '';

               ParamByName('can1_bung3_etc').AsString := UF_B201(can1_bung3_etc1.Checked) +
                                                         UF_B201(can1_bung3_etc2.Checked) +
                                                         UF_B201(can1_bung3_etc3.Checked) +
                                                         UF_B201(can1_bung3_etc4.Checked) +
                                                         UF_B201(Can1_Bung3_Etc5.Checked) +
                                                         UF_B201(Can1_Bung3_Etc6.Checked) +
                                                         UF_B201(Can1_Bung3_Etc7.Checked) +
                                                         UF_B201(Can1_Bung3_Etc8.Checked);
               if Can1_Bung3_Etc8.Checked then ParamByName('can1_bung3_etc_direct').AsString := GF_CutString(can1_bung3_etc_direct.Text, 20)
               else                            ParamByName('can1_bung3_etc_direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can1_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can1_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_UPDATE').AsString  := GV_sToday;
               ParamByName('COD_UPDATE').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryU_ca_cancer2009_Can1, 'LD1I12', 'U', 'Y');
            end;
         end
         //대장조직
         else if UV_vCod_hm[i] = 'B007' then
         begin
            with qryU_ca_cancer2009_Can2 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               ParamByName('can2_bung4_jindan').AsInteger := UF_Check(can2_bung4_jindan.ItemIndex);
               ParamByName('can2_bung4_count').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);
               ParamByName('can2_bung4_count_Gong').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);

               ParamByName('can2_bung4_jong').AsString := UF_B201_Jong1(can2_bung4_jong1.Checked, can2_bung4_jong1_H.Checked, can2_bung4_jong1_M.Checked, can2_bung4_jong1_L.Checked) +
                                                          UF_B201(can2_bung4_jong2.Checked) +
                                                          UF_B201(can2_bung4_jong3.Checked) +
                                                          UF_B201(can2_bung4_jong4.Checked) +
                                                          UF_B201(can2_bung4_jong5.Checked) +
                                                          UF_B201(can2_bung4_jong6.Checked) +
                                                          UF_B201(can2_bung4_jong7.Checked) +
                                                          UF_B201(can2_bung4_jong8.Checked) +
                                                          UF_B201(can2_bung4_jong9.Checked) +
                                                          UF_B201(can2_bung4_jong10.Checked) +
                                                          UF_B201(can2_bung4_jong11.Checked);
               if can2_bung4_jong11.Checked then ParamByName('Can2_Bung4_Jong_Direct').AsString := GF_CutString(Can2_Bung4_Jong_Direct.Text, 20)
               else                              ParamByName('Can2_Bung4_Jong_Direct').AsString := '';


               ParamByName('can2_bung4_etc').AsString := UF_B201(Can2_Bung4_Etc1.Checked) +
                                                         UF_B201(Can2_Bung4_Etc2.Checked) +
                                                         UF_B201(Can2_Bung4_Etc3.Checked) +
                                                         UF_B201(Can2_Bung4_Etc4.Checked) +
                                                         UF_B201(Can2_Bung4_Etc5.Checked);
               if Can2_Bung4_Etc5.Checked then ParamByName('Can2_Bung4_Etc_Direct').AsString := GF_CutString(Can2_Bung4_Etc_Direct.Text, 20)
               else                            ParamByName('Can2_Bung4_Etc_Direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can2_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can2_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_UPDATE').AsString  := GV_sToday;
               ParamByName('COD_UPDATE').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryU_ca_cancer2009_Can2, 'LD1I12', 'U', 'Y');
            End;
         end;
      end
      else
      begin
         //위조직
         if UV_vCod_hm[i] = 'B001' then
         begin
            with qryI_ca_cancer2009_Can1 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
               ParamByName('can1_bung3_jindan').AsInteger := UF_Check(Can1_Bung3_Jindan.ItemIndex);
               ParamByName('can1_bung3_count').AsInteger  := UF_Check(Can1_Bung3_Count.ItemIndex);

               ParamByName('can1_bung3_jong').AsString := UF_B201_Jong1(can1_bung3_jong1.Checked, can1_bung3_jong1_H.Checked, can1_bung3_jong1_M.Checked, can1_bung3_jong1_L.Checked) +
                                                          UF_B201(can1_bung3_jong2.Checked) +
                                                          UF_B201(can1_bung3_jong3.Checked) +
                                                          UF_B201_Jong4(can1_bung3_jong4.Checked, can1_bung3_jong4_H.Checked, can1_bung3_jong4_L.Checked) +
                                                          UF_B201(can1_bung3_jong5.Checked) +
                                                          UF_B201(can1_bung3_jong6.Checked) +
                                                          UF_B201(can1_bung3_jong7.Checked) +
                                                          UF_B201(can1_bung3_jong8.Checked) +
                                                          UF_B201(can1_bung3_jong9.Checked) +
                                                          UF_B201(can1_bung3_jong10.Checked) +
                                                          UF_B201(can1_bung3_jong11.Checked);
               if can1_bung3_jong11.Checked then ParamByName('can1_bung3_jong_direct').AsString := GF_CutString(can1_bung3_jong_direct.Text, 20)
               else                              ParamByName('can1_bung3_jong_direct').AsString := '';

               ParamByName('can1_bung3_etc').AsString := UF_B201(can1_bung3_etc1.Checked) +
                                                         UF_B201(can1_bung3_etc2.Checked) +
                                                         UF_B201(can1_bung3_etc3.Checked) +
                                                         UF_B201(can1_bung3_etc4.Checked) +
                                                         UF_B201(Can1_Bung3_Etc5.Checked) +
                                                         UF_B201(Can1_Bung3_Etc6.Checked) +
                                                         UF_B201(Can1_Bung3_Etc7.Checked) +
                                                         UF_B201(Can1_Bung3_Etc8.Checked);
               if Can1_Bung3_Etc8.Checked then ParamByName('can1_bung3_etc_direct').AsString := GF_CutString(can1_bung3_etc_direct.Text, 20)
               else                            ParamByName('can1_bung3_etc_direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can1_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can1_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_Insert').AsString  := GV_sToday;
               ParamByName('COD_Insert').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryI_ca_cancer2009_Can1, 'LD1I12', 'I', 'Y');
            end;
         end
         //대장조직
         else if UV_vCod_hm[i] = 'B007' then
         begin
            with qryI_ca_cancer2009_Can2 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               ParamByName('can2_bung4_jindan').AsInteger := UF_Check(can2_bung4_jindan.ItemIndex);
               ParamByName('can2_bung4_count').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);
               ParamByName('can2_bung4_count_Gong').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);

               ParamByName('can2_bung4_jong').AsString := UF_B201_Jong1(can2_bung4_jong1.Checked, can2_bung4_jong1_H.Checked, can2_bung4_jong1_M.Checked, can2_bung4_jong1_L.Checked) +
                                                          UF_B201(can2_bung4_jong2.Checked) +
                                                          UF_B201(can2_bung4_jong3.Checked) +
                                                          UF_B201(can2_bung4_jong4.Checked) +
                                                          UF_B201(can2_bung4_jong5.Checked) +
                                                          UF_B201(can2_bung4_jong6.Checked) +
                                                          UF_B201(can2_bung4_jong7.Checked) +
                                                          UF_B201(can2_bung4_jong8.Checked) +
                                                          UF_B201(can2_bung4_jong9.Checked) +
                                                          UF_B201(can2_bung4_jong10.Checked) +
                                                          UF_B201(can2_bung4_jong11.Checked);
               if can2_bung4_jong11.Checked then ParamByName('Can2_Bung4_Jong_Direct').AsString := GF_CutString(Can2_Bung4_Jong_Direct.Text, 20)
               else                              ParamByName('Can2_Bung4_Jong_Direct').AsString := '';


               ParamByName('can2_bung4_etc').AsString := UF_B201(Can2_Bung4_Etc1.Checked) +
                                                         UF_B201(Can2_Bung4_Etc2.Checked) +
                                                         UF_B201(Can2_Bung4_Etc3.Checked) +
                                                         UF_B201(Can2_Bung4_Etc4.Checked) +
                                                         UF_B201(Can2_Bung4_Etc5.Checked);
               if Can2_Bung4_Etc5.Checked then ParamByName('Can2_Bung4_Etc_Direct').AsString := GF_CutString(Can2_Bung4_Etc_Direct.Text, 20)
               else                            ParamByName('Can2_Bung4_Etc_Direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can2_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can2_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_Insert').AsString  := GV_sToday;
               ParamByName('COD_Insert').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryI_ca_cancer2009_Can2, 'LD1I12', 'I', 'Y');
            End;
         end;
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
         //위조직
         if UV_vCod_hm[i] = 'B001' then
         begin
            with qryU_ca_cancer2018_Can1 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               if copy(Can1_Bung3_Jindan_2018.text, 1, 1) <> '2' then
                  ParamByName('can1_bung3_jindan').AsString := copy(Can1_Bung3_Jindan_2018.text, 1,1)
               else
                  ParamByName('can1_bung3_jindan').AsString := copy(Can1_Bung3_Jindan_2018.text, 1,2);

               ParamByName('can1_bung3_count').AsInteger  := UF_Check(Can1_Bung3_Count.ItemIndex);

               ParamByName('can1_bung3_jong').AsString := UF_B201_Jong1(can1_bung3_jong1.Checked, can1_bung3_jong1_H.Checked, can1_bung3_jong1_M.Checked, can1_bung3_jong1_L.Checked) +
                                                          UF_B201(can1_bung3_jong2.Checked) +
                                                          UF_B201(can1_bung3_jong3.Checked) +
                                                          UF_B201_Jong4(can1_bung3_jong4.Checked, can1_bung3_jong4_H.Checked, can1_bung3_jong4_L.Checked) +
                                                          UF_B201(can1_bung3_jong5.Checked) +
                                                          UF_B201(can1_bung3_jong6.Checked) +
                                                          UF_B201(can1_bung3_jong7.Checked) +
                                                          UF_B201(can1_bung3_jong8.Checked) +
                                                          UF_B201(can1_bung3_jong9.Checked) +
                                                          UF_B201(can1_bung3_jong10.Checked) +
                                                          UF_B201(can1_bung3_jong11.Checked);
               if can1_bung3_jong11.Checked then ParamByName('can1_bung3_jong_direct').AsString := GF_CutString(can1_bung3_jong_direct.Text, 20)
               else                              ParamByName('can1_bung3_jong_direct').AsString := '';

               ParamByName('can1_bung3_etc').AsString := UF_B201(can1_bung3_etc1.Checked) +
                                                         UF_B201(can1_bung3_etc2.Checked) +
                                                         UF_B201(can1_bung3_etc3.Checked) +
                                                         UF_B201(can1_bung3_etc4.Checked) +
                                                         UF_B201(Can1_Bung3_Etc5.Checked) +
                                                         UF_B201(Can1_Bung3_Etc6.Checked) +
                                                         UF_B201(Can1_Bung3_Etc7.Checked) +
                                                         UF_B201(Can1_Bung3_Etc8.Checked);
               if Can1_Bung3_Etc8.Checked then ParamByName('can1_bung3_etc_direct').AsString := GF_CutString(can1_bung3_etc_direct.Text, 20)
               else                            ParamByName('can1_bung3_etc_direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can1_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can1_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_UPDATE').AsString  := GV_sToday;
               ParamByName('COD_UPDATE').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryU_ca_cancer2018_Can1, 'LD1I12', 'U', 'Y');
            end;
         end
         //대장조직
         else if UV_vCod_hm[i] = 'B007' then
         begin
            with qryU_ca_cancer2018_Can2 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               ParamByName('can2_bung4_jindan').AsInteger := UF_Check(can2_bung4_jindan_2018.ItemIndex);
               ParamByName('can2_bung4_count').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);
               ParamByName('can2_bung4_count_Gong').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);

               ParamByName('can2_bung4_jong').AsString := UF_B201_Jong1(can2_bung4_jong1.Checked, can2_bung4_jong1_H.Checked, can2_bung4_jong1_M.Checked, can2_bung4_jong1_L.Checked) +
                                                          UF_B201(can2_bung4_jong2.Checked) +
                                                          UF_B201(can2_bung4_jong3.Checked) +
                                                          UF_B201(can2_bung4_jong4.Checked) +
                                                          UF_B201(can2_bung4_jong5.Checked) +
                                                          UF_B201(can2_bung4_jong6.Checked) +
                                                          UF_B201(can2_bung4_jong7.Checked) +
                                                          UF_B201(can2_bung4_jong8.Checked) +
                                                          UF_B201(can2_bung4_jong9.Checked) +
                                                          UF_B201(can2_bung4_jong10.Checked) +
                                                          UF_B201(can2_bung4_jong11.Checked);
               if can2_bung4_jong11.Checked then ParamByName('Can2_Bung4_Jong_Direct').AsString := GF_CutString(Can2_Bung4_Jong_Direct.Text, 20)
               else                              ParamByName('Can2_Bung4_Jong_Direct').AsString := '';


               ParamByName('can2_bung4_etc').AsString := UF_B201(Can2_Bung4_Etc1.Checked) +
                                                         UF_B201(Can2_Bung4_Etc2.Checked) +
                                                         UF_B201(Can2_Bung4_Etc3.Checked) +
                                                         UF_B201(Can2_Bung4_Etc4.Checked) +
                                                         UF_B201(Can2_Bung4_Etc5.Checked);
               if Can2_Bung4_Etc5.Checked then ParamByName('Can2_Bung4_Etc_Direct').AsString := GF_CutString(Can2_Bung4_Etc_Direct.Text, 20)
               else                            ParamByName('Can2_Bung4_Etc_Direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can2_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can2_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_UPDATE').AsString  := GV_sToday;
               ParamByName('COD_UPDATE').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryU_ca_cancer2018_Can2, 'LD1I12', 'U', 'Y');
            End;
         end;
      end
      else
      begin
         //위조직
         if UV_vCod_hm[i] = 'B001' then
         begin
            with qryI_ca_cancer2018_Can1 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];
               
               if copy(Can1_Bung3_Jindan_2018.text, 1, 1) <> '2' then
                  ParamByName('can1_bung3_jindan').AsString := copy(Can1_Bung3_Jindan_2018.text, 1,1)
               else
                  ParamByName('can1_bung3_jindan').AsString := copy(Can1_Bung3_Jindan_2018.text, 1,2);

               ParamByName('can1_bung3_count').AsInteger  := UF_Check(Can1_Bung3_Count.ItemIndex);

               ParamByName('can1_bung3_jong').AsString := UF_B201_Jong1(can1_bung3_jong1.Checked, can1_bung3_jong1_H.Checked, can1_bung3_jong1_M.Checked, can1_bung3_jong1_L.Checked) +
                                                          UF_B201(can1_bung3_jong2.Checked) +
                                                          UF_B201(can1_bung3_jong3.Checked) +
                                                          UF_B201_Jong4(can1_bung3_jong4.Checked, can1_bung3_jong4_H.Checked, can1_bung3_jong4_L.Checked) +
                                                          UF_B201(can1_bung3_jong5.Checked) +
                                                          UF_B201(can1_bung3_jong6.Checked) +
                                                          UF_B201(can1_bung3_jong7.Checked) +
                                                          UF_B201(can1_bung3_jong8.Checked) +
                                                          UF_B201(can1_bung3_jong9.Checked) +
                                                          UF_B201(can1_bung3_jong10.Checked) +
                                                          UF_B201(can1_bung3_jong11.Checked);
               if can1_bung3_jong11.Checked then ParamByName('can1_bung3_jong_direct').AsString := GF_CutString(can1_bung3_jong_direct.Text, 20)
               else                              ParamByName('can1_bung3_jong_direct').AsString := '';

               ParamByName('can1_bung3_etc').AsString := UF_B201(can1_bung3_etc1.Checked) +
                                                         UF_B201(can1_bung3_etc2.Checked) +
                                                         UF_B201(can1_bung3_etc3.Checked) +
                                                         UF_B201(can1_bung3_etc4.Checked) +
                                                         UF_B201(Can1_Bung3_Etc5.Checked) +
                                                         UF_B201(Can1_Bung3_Etc6.Checked) +
                                                         UF_B201(Can1_Bung3_Etc7.Checked) +
                                                         UF_B201(Can1_Bung3_Etc8.Checked);
               if Can1_Bung3_Etc8.Checked then ParamByName('can1_bung3_etc_direct').AsString := GF_CutString(can1_bung3_etc_direct.Text, 20)
               else                            ParamByName('can1_bung3_etc_direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can1_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can1_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_Insert').AsString  := GV_sToday;
               ParamByName('COD_Insert').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryI_ca_cancer2018_Can1, 'LD1I12', 'I', 'Y');
            end;
         end
         //대장조직
         else if UV_vCod_hm[i] = 'B007' then
         begin
            with qryI_ca_cancer2018_Can2 do
            begin
               ParamByName('COD_JISA').AsString := UV_vCod_Jisa[i];
               ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
               ParamByName('DAT_GMGN').AsString := UV_vDat_gmgn[i];

               ParamByName('can2_bung4_jindan').AsInteger := UF_Check(can2_bung4_jindan_2018.ItemIndex);
               ParamByName('can2_bung4_count').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);
               ParamByName('can2_bung4_count_Gong').AsInteger  := UF_Check(can2_bung4_count.ItemIndex);

               ParamByName('can2_bung4_jong').AsString := UF_B201_Jong1(can2_bung4_jong1.Checked, can2_bung4_jong1_H.Checked, can2_bung4_jong1_M.Checked, can2_bung4_jong1_L.Checked) +
                                                          UF_B201(can2_bung4_jong2.Checked) +
                                                          UF_B201(can2_bung4_jong3.Checked) +
                                                          UF_B201(can2_bung4_jong4.Checked) +
                                                          UF_B201(can2_bung4_jong5.Checked) +
                                                          UF_B201(can2_bung4_jong6.Checked) +
                                                          UF_B201(can2_bung4_jong7.Checked) +
                                                          UF_B201(can2_bung4_jong8.Checked) +
                                                          UF_B201(can2_bung4_jong9.Checked) +
                                                          UF_B201(can2_bung4_jong10.Checked) +
                                                          UF_B201(can2_bung4_jong11.Checked);
               if can2_bung4_jong11.Checked then ParamByName('Can2_Bung4_Jong_Direct').AsString := GF_CutString(Can2_Bung4_Jong_Direct.Text, 20)
               else                              ParamByName('Can2_Bung4_Jong_Direct').AsString := '';


               ParamByName('can2_bung4_etc').AsString := UF_B201(Can2_Bung4_Etc1.Checked) +
                                                         UF_B201(Can2_Bung4_Etc2.Checked) +
                                                         UF_B201(Can2_Bung4_Etc3.Checked) +
                                                         UF_B201(Can2_Bung4_Etc4.Checked) +
                                                         UF_B201(Can2_Bung4_Etc5.Checked);
               if Can2_Bung4_Etc5.Checked then ParamByName('Can2_Bung4_Etc_Direct').AsString := GF_CutString(Can2_Bung4_Etc_Direct.Text, 20)
               else                            ParamByName('Can2_Bung4_Etc_Direct').AsString := '';

               SUBSOKUN    := '';
               UV_SBSOKUN1 := '';
               UV_SBSOKUN2 := '';
               UV_SBSOKUN3 := '';
               UV_SBSOKUN4 := '';
               SUBSOKUN    := Can_Advice.Text;
               GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
               ParamByName('can2_advice').AsString  := UV_SBSOKUN1;
               ParamByName('can2_advice1').AsString := UV_SBSOKUN2;

               ParamByName('DAT_Insert').AsString  := GV_sToday;
               ParamByName('COD_Insert').AsString  := GV_sUserId;
               Execsql;
               GP_query_log(qryI_ca_cancer2018_Can2, 'LD1I12', 'I', 'Y');
            End;
         end;
      end;
      end
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
      UV_vDesc_bang[i]   := edtDESC_BANG.Text;
      UV_vDesc_buwi[i]   := edtDESC_BUWI.Text;
      UV_vCod_sawon[i]   := Copy(cmbCOD_SAWON.Text, 1, 6);
      UV_vCod_doct[i]    := Copy(cmbCOD_DOCT.Text, 1, 6);
      UV_vDat_last[i]    := mskDAT_LAST.Text;
      UV_vDat_result[i]  := mskDAT_RESULT.Text;
      UV_vDesc_kgbl[i]   := retDESC_KGBL.Text;
      UV_vDesc_sokun[i]  := retDESC_SOKUN.Text;
      UV_vDesc_sokun1[i] := retDESC_SOKUN1.Text;

      if Trim(mskDAT_RESULT.Text) <> '' then
         UV_vYsno_sokun[i] := 'Y'
      else
         UV_vYsno_sokun[i] := 'N';

      UV_vGubn_pan[i]     := edtGUBN_PAN.Text;
      UV_vCod_sokun[i]    := edtCod_sokun.Text;
      UV_vCell_level[i]   := copy(cmb_level.Text,1,1);
      UV_vPsokun[i]       := Memo_Psokun.Text;
      UV_vFloor[i]        := Edt_floor.Text;
      UV_vRent[i]         := Lab_slide.Caption;
      UV_vCmt[i]          := Edt_cmt.Text;
      UV_vDat_Rent[i]     := msk_dat_rent.Text;
      UV_vDAT_Return[i]   := msk_dat_Return.Text;

      //외부 입력 의사 추가
      If chk_jindan.Checked Then
         UV_vCheck_jindan[i]   := 'Y'
      else UV_vCheck_jindan[i] := 'N';
      UV_vCod_jindan_doct[i] := Copy(cmbCOD_jindan_DOCT.Text, 1, 6);
      // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;
   end;

   UV_iCod_doct := cmbCOD_DOCT.ItemIndex;
   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLDAI02.btnSaveClick(Sender: TObject);
begin
   inherited;
   UF_Save;
end;

procedure TfrmLDAI02.FormCreate(Sender: TObject);
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

//   GP_ComboFill(cmbGUBN_CLASS,  'CLAS');
   with qryCommon do
   begin
//      Close;
      Open;
   end;
   GP_ComboSawonAll(cmbCOD_SAWON, 0);
   GP_ComboSawonAll(cmbCOD_DOCT, 1);
   GP_ComboSawonAll(cmbCOD_jindan_DOCT, 1);

   // 분석부 조직실 전체...
   if (GV_sUserID = '150007') or (GV_sUserID = '150365') or (GV_sUserID = '410015') or
      (GV_sUserID = '150281') or (GV_sUserID = '150627') or (GV_sUserID = '150332') or
      (GV_sUserID = '150644') or (GV_sUserID = '150486') or (GV_sUserID = '150681') or
      (GV_sUserID = '150875') or (GV_sUserID = '150344') or (GV_sUserID = '151026') then cmbRelation.Enabled := True
   else                                                                                  cmbRelation.Enabled := False;


   // SQL문 저장
   UV_sBasicSQL := qryCell.SQL.Text;
   UV_sBasicSQL_cs := qryCell_sokun.SQL.Text;
   pnlHide.Visible  := False;
   pnlHide2.Visible := False;
   pnlHide3.Visible := False;
   if (GV_sUserID = '150007') or (GV_sUserID = '150365') or (GV_sUserID = '410015') or
      (GV_sUserID = '150281') or (GV_sUserID = '150627') or (GV_sUserID = '150332') or
      (GV_sUserID = '150644') or (GV_sUserID = '150486') or (GV_sUserID = '150681') or
      (GV_sUserID = '150875') or (GV_sUserID = '150344') or (GV_sUserID = '151026') then pnlHide4.Visible := False
   else                                                                                  pnlHide4.Visible := True;  //병리팀만 조회 가능하도록
   pnlHide.Left := 80;
   pnlHide.Top := 225;

   UV_iCod_doct := -1;
   cmbGubn.ItemIndex := 0;
   cmb_Level.ItemIndex := 1;
   mskDate.Text := GV_sToday;
end;

procedure TfrmLDAI02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자료를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_Jisa[DataRow-1];
      2 : begin
             if rbDate.Checked  then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
             if rbRDate.Checked then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
             if rbJumin.Checked then Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
             if rbCell.Checked  then Value  := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
          end;
      3 : begin
             if rbDate.Checked  then Value := UV_vDesc_name[DataRow-1];
             if rbRDate.Checked then Value := UV_vDesc_name[DataRow-1];
             if rbJumin.Checked then Value := '';
             if rbCell.Checked  then Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
          end;
      4 : Value := UV_vDesc_buwi[DataRow-1];
      5 : Value := UV_vCod_cell[DataRow-1];
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

procedure TfrmLDAI02.btnQueryClick(Sender: TObject);
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

         if Chk_order.Checked then sWhere := sWhere + 'and A.gubn_order = ''Y''';
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

         if Chk_order.Checked then sWhere := sWhere + 'and A.gubn_order = ''Y''';
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
           if cmbGubn.ItemIndex = 0  then sWhere := sWhere + ' AND ((SUBSTRING(cod_cell,1,1) = ''J'') or (cod_hm = ''P009'')) '
      else if cmbGubn.ItemIndex = 1  then sWhere := sWhere + ' AND cod_hm = ''B001'''
      else if cmbGubn.ItemIndex = 2  then sWhere := sWhere + ' AND cod_hm = ''B005'''
      else if cmbGubn.ItemIndex = 3  then sWhere := sWhere + ' AND cod_hm = ''B007'''
      else if cmbGubn.ItemIndex = 4  then sWhere := sWhere + ' AND cod_hm = ''B012'''
      else if cmbGubn.ItemIndex = 5  then sWhere := sWhere + ' AND cod_hm = ''P009''';

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
            UV_vDesc_bang[index]   := FieldByName('DESC_BANG').AsString;
            UV_vDesc_buwi[index]   := FieldByName('DESC_BUWI').AsString;
            UV_vCod_sawon[index]   := FieldByName('COD_SAWON').AsString;
            UV_vCod_doct[index]    := FieldByName('COD_DOCT').AsString;
            UV_vCod_jindan_doct[index] := FieldByName('COD_jindan_doctor').AsString;
            UV_vCheck_jindan[index]    := FieldByName('check_jindan').AsString;
            UV_vCod_Pan[index]     := FieldByName('COD_PAN').AsString;
            UV_vDESC_Pan[index]    := FieldByName('DESC_PAN').AsString;
            UV_vDat_last[index]    := FieldByName('DAT_LAST').AsString;
            UV_vDat_Time[index]    := FieldByName('DAT_Time').AsString;
            UV_vDat_result[index]  := FieldByName('DAT_RESULT').AsString;
            UV_vDat_Time_R[index]  := FieldByName('DAT_Time_R').AsString;
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
            UV_vCell_level[index]  := FieldByName('cell_level').AsString;
            
            // 층구분 표시 ..20170720
            if qryCell.FieldByName('gubn_jinch').AsString <> '' then
              begin
              qryFloor.Active := False;
              qryFloor.ParamByName('cod_group').AsString  := 'FLOR';
              qryFloor.ParamByName('cod_detail').AsString := qryCell.FieldByName('gubn_jinch').AsString;
              qryFloor.Active := True;
              UV_vFloor[index] := qryFloor.fieldByName('desc_code').AsString;
              end;

            //슬라이드 대여 표시 ..20170720
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

            if (UV_vCod_hm[index] = 'B001') or (UV_vCod_hm[index] = 'B007') then
            begin
               with qryJangbi do
               begin
                  // Filter 제거
                  Filtered := False;
                  qryJangbi.Active := False;
                  qryJangbi.ParamByName('cod_jisa').AsString := UV_vCod_jisa[index];
                  qryJangbi.ParamByName('num_id').AsString   := UV_vNum_id[index];
                  qryJangbi.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[index];
                  qryJangbi.Active := True;
                  Filtered := True;

                  if UV_vCod_hm[index] = 'B001' then        //위내시경
                  begin
                     qryJangbi.Filter := ' cod_jangbi = ''JJ34'' ';
                     if qryJangbi.RecordCount > 0 then
                     begin
                        UV_vDoctor[index] := qryJangbi.FieldByName('DOCTOR').AsString;
                     end
                     else
                     begin
                        qryJangbi.Filter := ' cod_jangbi = ''JJB9'' ';
                        if qryJangbi.RecordCount > 0 then
                        begin
                           UV_vDoctor[index] := qryJangbi.FieldByName('DOCTOR').AsString;
                        end;
                     end;
                  end
                  else if UV_vCod_hm[index] = 'B007' then   //대장내시경
                  begin
                     qryJangbi.Filter := ' cod_jangbi = ''JJ83'' ';
                     if qryJangbi.RecordCount > 0 then
                     begin
                        UV_vDoctor[index] := qryJangbi.FieldByName('DOCTOR').AsString;
                     end
                     else
                     begin
                        qryJangbi.Filter := ' cod_jangbi = ''JJ86'' ';
                        if qryJangbi.RecordCount > 0 then
                        begin
                           UV_vDoctor[index] := qryJangbi.FieldByName('DOCTOR').AsString;
                        end;
                     end;
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

procedure TfrmLDAI02.btnCancelClick(Sender: TObject);
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

procedure TfrmLDAI02.UP_Change(Sender: TObject);
var sPan1, sSokun1, sPsokun : String;
begin
   inherited;
   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = edtCOD_SOKUN then
   begin
      IF edtCOD_SOKUN.Text <> 'po03' then  edtGUBN_PAN.Text := '';

      // '****'이면 미검이므로 자료를 Clear
      if Trim(edtGUBN_PAN.Text) = '' then  //2015.03.04 곽윤설 .. 저장값 우선 끌어오기
      begin
         if GF_CellSokun_P(edtCod_sokun.Text, sSokun1, sPan1) then
         begin
            edtGUBN_PAN.Text   := sPan1;
         end;
      end;
   end
   else if Sender = Can1_Bung3_Jong1 then //위 관상샘암종
   begin
        if ((Can1_Bung3_Jong1.checked) or (Can1_Bung3_Jong4.checked)) and (UV_bGmgn) then  //샘암종과 위름프종중 체크시 세분화 그룹 생성
            GroupBox12.Visible := True
        else
            GroupBox12.Visible := False;

        if (Can1_Bung3_Jong1.Checked) and (UV_bGmgn) then
            GroupBox13.Visible := True
        else
            GroupBox13.Visible := False;
   end
   else if Sender = Can1_Bung3_Jong4 then
   begin
        if ((Can1_Bung3_Jong1.checked) or (Can1_Bung3_Jong4.checked)) and (UV_bGmgn) then
            GroupBox12.Visible := True
        else
            GroupBox12.Visible := False;

        if (Can1_Bung3_Jong4.Checked) and (UV_bGmgn) then
            GroupBox14.Visible := True
        else
            GroupBox14.Visible := False;
   end
   else if Sender = Can2_Bung4_Jong1 then
   begin

        if (Can2_Bung4_Jong1.Checked) and (UV_bGmgn) then
            GroupBox15.Visible := True
        else
            GroupBox15.Visible := False;
   end;

  if grdMaster.CurrentDataRow > 0 then
  begin
      if Sender = chk_jindan then
      begin
         if chk_jindan.Checked then
         begin
            IF (UV_vCod_Jindan_doct[grdMaster.CurrentDataRow-1] = '' ) then GP_ComboDisplay(cmbCOD_jindan_Doct, '151238', 6);
            GP_ComboDisplay(cmbCOD_SAWON, '800016', 6);
         end
         else
         begin
          cmbCOD_JINDAN_DOCT.ItemIndex := -1;
          //검사자
          if Trim(UV_vCod_sawon[grdMaster.CurrentDataRow-1]) = '' then
           begin
              if (copy(GV_sUserId,1,2) = '15') then
              begin
                 cmbCOD_SAWON.Font.Color := clRed;
                 if      (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'B001') or
                         (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'B005') or
                         (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'B007') or
                         (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,4) = 'P009') then
                    GP_ComboDisplay(cmbCOD_SAWON, '151021', 6) //최재현
                 else if (Copy(UV_vCod_hm[grdMaster.CurrentDataRow-1],1,1) = 'B') then
                    GP_ComboDisplay(cmbCOD_SAWON, '151030', 6) //최윤선
                 else
                    GP_ComboDisplay(cmbCOD_SAWON, GV_sUserId, 6);
              end
              else cmbCOD_SAWON.ItemIndex := -1;
           end
           else
           begin
              cmbCOD_SAWON.Font.Color := clBlack;
              GP_ComboDisplay(cmbCOD_SAWON, UV_vCod_sawon[grdMaster.CurrentDataRow-1], 6);
           end;
         end;
        //전문의
         if Trim(UV_vCod_Doct[grdMaster.CurrentDataRow-1]) = '' then
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
            GP_ComboDisplay(cmbCOD_Doct, UV_vCod_Doct[grdMaster.CurrentDataRow-1], 6);
         end;
      end;
  end;
   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLDAI02.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
Var
   Max_Num : integer;
   sWhere, sOrderby, sWhere_cs, sPsokun: String;
   a, b : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   Can2_Bung4_Jindan.ItemIndex := 0;
   Can2_Bung4_Jindan_2018.ItemIndex := 0;
   Can2_Bung4_Count.ItemIndex  := 0;
   cmb_level.ItemIndex   := 1;
   Can2_Bung4_Jong1.Checked := False;
   Can2_Bung4_Jong2.Checked := False;
   Can2_Bung4_Jong3.Checked := False;
   Can2_Bung4_Jong4.Checked := False;
   Can2_Bung4_Jong5.Checked := False;
   Can2_Bung4_Jong6.Checked := False;
   Can2_Bung4_Jong7.Checked := False;
   Can2_Bung4_Jong8.Checked := False;
   Can2_Bung4_Jong9.Checked := False;
   Can2_Bung4_Jong10.Checked := False;
   Can2_Bung4_Jong11.Checked := False;
   Can2_Bung4_Jong1_H.Checked := False;
   Can2_Bung4_Jong1_M.Checked := False;
   Can2_Bung4_Jong1_L.Checked := False;
   Can2_Bung4_Jong_Direct.Text := '';
   Can2_Bung4_Etc1.Checked := False;
   Can2_Bung4_Etc2.Checked := False;
   Can2_Bung4_Etc3.Checked := False;
   Can2_Bung4_Etc4.Checked := False;
   Can2_Bung4_Etc5.Checked := False;
   Can2_Bung4_Etc_Direct.Text := '';

   Can1_Bung3_Jindan.ItemIndex := 0;
   Can1_Bung3_Jindan_2018.ItemIndex := 0;
   Can1_Bung3_Count.ItemIndex  := 0;
   Can1_Bung3_Jong1.Checked := False;
   Can1_Bung3_Jong2.Checked := False;
   Can1_Bung3_Jong3.Checked := False;
   Can1_Bung3_Jong4.Checked := False;
   Can1_Bung3_Jong5.Checked := False;
   Can1_Bung3_Jong6.Checked := False;
   Can1_Bung3_Jong7.Checked := False;
   Can1_Bung3_Jong8.Checked := False;
   Can1_Bung3_Jong9.Checked := False;
   Can1_Bung3_Jong10.Checked := False;
   Can1_Bung3_Jong11.Checked := False;
   Can1_Bung3_Jong_Direct.Text := '';
   Can1_Bung3_Etc1.Checked := False;
   Can1_Bung3_Etc2.Checked := False;
   Can1_Bung3_Etc3.Checked := False;
   Can1_Bung3_Etc4.Checked := False;
   Can1_Bung3_Etc5.Checked := False;
   Can1_Bung3_Etc6.Checked := False;
   Can1_Bung3_Etc7.Checked := False;
   Can1_Bung3_Etc8.Checked := False;
   Can1_Bung3_Jong1_H.Checked := False;
   Can1_Bung3_Jong1_M.Checked := False;
   Can1_Bung3_Jong1_L.Checked := False;
   Can1_Bung3_Jong4_H.Checked := False;
   Can1_Bung3_Jong4_L.Checked := False;
   Can1_Bung3_Etc_Direct.Text := '';
   Can_Advice.Text := '';
   //edtCod_sokun.Text := '';   //2014.04.10 곽윤설 주석처리
   //edtGUBN_PAN.Text := '';    //2015.03.04 곽윤설 주석처리
   edtGUBN_PAN.Enabled   := True;
   //20150327 곽윤설

   pnlHide.Visible := False;
   pnlHide2.Visible := False;
   pnlHide3.Visible := False;
   pnlHide4.Visible := True;

   //공단 자궁암관련 항목 변경 .. 20150102  수지
      if (UV_vCod_hm[Newrow - 1] = 'B001') and
      (((rbDate.Checked) and (mskDate.Text >= '20180101')) or ((rbJumin.Checked) and (UV_vDat_gmgn[Newrow - 1] >= '20180101')) or
       ((rbRDate.Checked) and (mskRDate.Text >= '20180101')) or ((rbCell.Checked) and (mskcell.text <> '') and (UV_vDat_gmgn[Newrow - 1] >= '20180101'))) then
         Can1_Bung3_Jindan_2018.visible := True
      else  Can1_Bung3_Jindan_2018.visible := False;
      if (UV_vCod_hm[Newrow - 1] = 'B007') and
      (((rbDate.Checked) and (mskDate.Text >= '20180101')) or ((rbJumin.Checked) and (UV_vDat_gmgn[Newrow - 1] >= '20180101')) or
       ((rbRDate.Checked) and (mskRDate.Text >= '20180101')) or ((rbCell.Checked) and (mskcell.text <> '') and (UV_vDat_gmgn[Newrow - 1] >= '20180101'))) then
         Can2_Bung4_Jindan_2018.visible := True
      else  Can2_Bung4_Jindan_2018.visible := False;

   //병리 제외한 나머지는 판독중일때 결과 조회 불가능 하도록 적용 본원이유굥  ..20170424 수지
   if ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '150644') or (GV_sUserID = '151021') or
       (GV_sUserID = '151022') or (GV_sUserID = '150879') or (GV_sUserID = '101030') or (GV_sUserID = '151026')) then
   begin
       pnlHide.Visible := False;
       pnlHide2.Visible := False;
       pnlHide3.Visible := False;
       pnlHide4.Visible := False;
   end
   ELSE  if (UV_vYsno_sokun[Newrow - 1] = 'Y') or (UV_vYsno_sokun[Newrow - 1] = 'N') then
   begin
     if UV_vCod_hm[Newrow - 1] = 'B001' then pnlHide3.Visible := True
     else if UV_vCod_hm[Newrow - 1] = 'B007' then pnlHide2.Visible := True;
     pnlHide.Visible := true;
   end;
      //위내시경
      if UV_vCod_hm[Newrow - 1] = 'B001' then
      begin
         Pnl_Can1.Visible := True;
         Pnl_Can2.Visible := False;
      end
      else if UV_vCod_hm[Newrow - 1] = 'B007' then
      begin

         Pnl_Can1.Visible := False;
         Pnl_Can2.Visible := True;
      end
      else begin
         Pnl_Can1.Visible := False;
         Pnl_Can2.Visible := False;
      end;

   if (UV_vCod_hm[Newrow - 1] = 'B001') or
      (UV_vCod_hm[Newrow - 1] = 'B007') then
   begin
      with Qry_Sign do
      begin
         Close;
         ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Newrow - 1];
         ParamByName('Num_id').AsString   := UV_vNum_id[Newrow - 1];
         ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[Newrow - 1];
         if UV_vCod_hm[Newrow - 1] = 'B001' then
            ParamByName('Program_id').AsString := 'LD1I121'
         else if UV_vCod_hm[Newrow - 1] = 'B007' then
            ParamByName('Program_id').AsString := 'LD1I127';
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

   if (UV_vDat_gmgn[Newrow - 1] < '20180101') then
   begin
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
         if UV_vCod_hm[Newrow - 1] = 'B001' then
         begin
            if FieldByName('Can1_Bung3_Jindan').AsString <> '' then
               Can1_Bung3_Jindan.Itemindex := FieldByName('Can1_Bung3_Jindan').AsInteger  ;
            if (FieldByName('Can1_Bung3_Jindan').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026'))then
               Can1_Bung3_Jindan.Itemindex := 2;

            if FieldByName('Can1_Bung3_Count').AsString  <> '' then
               Can1_Bung3_Count.Itemindex := FieldByName('Can1_Bung3_Count').AsInteger;
            if (FieldByName('Can1_Bung3_Count').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026')) then    //2014.04.11 곽윤설
               Can1_Bung3_Count.Itemindex  := 1;

            if FieldByName('Can1_Bung3_Jong').AsString  <> '' then
            begin
               if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  <> '0' then
               begin
                    Can1_Bung3_Jong1.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '1' then  Can1_Bung3_Jong1_H.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '2' then  Can1_Bung3_Jong1_M.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '3' then  Can1_Bung3_Jong1_L.checked  := True;
               end;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,2,1)  = '1' then Can1_Bung3_Jong2.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,3,1)  = '1' then Can1_Bung3_Jong3.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  <> '0' then
               begin
                    Can1_Bung3_Jong4.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  = '1' then  Can1_Bung3_Jong4_H.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  = '2' then  Can1_Bung3_Jong4_L.checked  := True;
               end;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,5,1)  = '1' then Can1_Bung3_Jong5.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,6,1)  = '1' then Can1_Bung3_Jong6.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,7,1)  = '1' then Can1_Bung3_Jong7.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,8,1)  = '1' then Can1_Bung3_Jong8.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,9,1)  = '1' then Can1_Bung3_Jong9.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,10,1) = '1' then Can1_Bung3_Jong10.checked := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,11,1) = '1' then
               begin
                  Can1_Bung3_Jong11.checked    := True;
                  Can1_Bung3_Jong_Direct.Text := FieldByName('Can1_Bung3_Jong_Direct').AsString;
               end;
            end;

            if FieldByName('Can1_Bung3_Etc').AsString  <> '' then
            begin
               if copy(FieldByName('Can1_Bung3_Etc').AsString,1,1) = '1' then Can1_Bung3_Etc1.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,2,1) = '1' then Can1_Bung3_Etc2.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,3,1) = '1' then Can1_Bung3_Etc3.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,4,1) = '1' then Can1_Bung3_Etc4.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,5,1) = '1' then Can1_Bung3_Etc5.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,6,1) = '1' then Can1_Bung3_Etc6.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,7,1) = '1' then Can1_Bung3_Etc7.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,8,1) = '1' then
               begin
                  Can1_Bung3_Etc8.checked    := True;
                  Can1_Bung3_Etc_Direct.Text := FieldByName('Can1_Bung3_Etc_Direct').AsString;
               end;
            end;

            Can_Advice.Text := Trim(FieldByName('Can1_Advice').AsString);
            Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can1_Advice1').AsString);
         End
         else if UV_vCod_hm[Newrow - 1] = 'B007' then
         begin
            if FieldByName('Can2_Bung4_Jindan').AsString <> '' then
               Can2_Bung4_Jindan.Itemindex := FieldByName('Can2_Bung4_Jindan').AsInteger;
            if (FieldByName('Can2_Bung4_Jindan').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026'))then      //2014.04.11 곽윤설
               Can2_Bung4_Jindan.Itemindex := 3 ;

            if FieldByName('Can2_Bung4_Count').AsString  <> '' then
               Can2_Bung4_Count.Itemindex := FieldByName('Can2_Bung4_Count').AsInteger;
            if (FieldByName('Can2_Bung4_Count').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681')  or (GV_sUserID = '151026')) then   //2014.04.11 곽윤설
                   Can2_Bung4_Count.Itemindex  := 1;

            if FieldByName('Can2_Bung4_Jong').AsString  <> '' then
            begin
               if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)<> '0' then
               begin
                    Can2_Bung4_Jong1.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '1' then  Can2_Bung4_Jong1_H.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '2' then  Can2_Bung4_Jong1_M.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '3' then  Can2_Bung4_Jong1_L.checked  := True;
               end;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,2,1)  = '1' then Can2_Bung4_Jong2.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,3,1)  = '1' then Can2_Bung4_Jong3.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,4,1)  = '1' then Can2_Bung4_Jong4.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,5,1)  = '1' then Can2_Bung4_Jong5.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,6,1)  = '1' then Can2_Bung4_Jong6.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,7,1)  = '1' then Can2_Bung4_Jong7.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,8,1)  = '1' then Can2_Bung4_Jong8.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,9,1)  = '1' then Can2_Bung4_Jong9.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,10,1) = '1' then Can2_Bung4_Jong10.checked := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,11,1) = '1' then
               begin
                  Can2_Bung4_Jong11.checked := True;
                  Can2_Bung4_Jong_Direct.Text := FieldByName('Can2_Bung4_Jong_Direct').AsString;
               end;
            end;

            if FieldByName('Can2_Bung4_Etc').AsString  <> '' then
            begin
               if copy(FieldByName('Can2_Bung4_Etc').AsString,1,1)  = '1' then Can2_Bung4_Etc1.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,2,1)  = '1' then Can2_Bung4_Etc2.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,3,1)  = '1' then Can2_Bung4_Etc3.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,4,1)  = '1' then Can2_Bung4_Etc4.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,5,1)  = '1' then
               begin
                  Can2_Bung4_Etc5.checked := True;
                  Can2_Bung4_Etc_Direct.Text := FieldByName('Can2_Bung4_Etc_Direct').AsString;
               end;
            end;

            Can_Advice.Text := Trim(FieldByName('Can2_Advice').AsString);
            Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can2_Advice1').AsString);
         end;
      end
      else
         Can_Advice.Text := '';
   end;
   Qry_CK_Gulkwa.close;

   end
   else
   begin
   Qry_CK_Gulkwa.close;
   with Qry_CK_Gulkwa_2018 do
   begin
      Close;
      Qry_CK_Gulkwa_2018.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Newrow - 1];
      Qry_CK_Gulkwa_2018.ParamByName('Num_id').AsString   := UV_vNum_id[Newrow - 1];
      Qry_CK_Gulkwa_2018.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[Newrow - 1];
      Qry_CK_Gulkwa_2018.ParamByName('COD_CELL').AsString := UV_vCod_cell[Newrow - 1];
      Qry_CK_Gulkwa_2018.ParamByName('COD_HM').AsString   := UV_vCod_hm[Newrow - 1];

      Open;
      if Qry_CK_Gulkwa_2018.IsEmpty = False then
      begin
         if UV_vCod_hm[Newrow - 1] = 'B001' then
         begin
            if FieldByName('Can1_Bung3_Jindan').AsString <> '' then
            begin
               if fieldbyname('can1_bung3_jindan').asString = '1' then Can1_Bung3_Jindan_2018.Itemindex := 1
               else if fieldbyname('can1_bung3_jindan').asString = '21' then Can1_Bung3_Jindan_2018.Itemindex := 2
               else if fieldbyname('can1_bung3_jindan').asString = '22' then Can1_Bung3_Jindan_2018.Itemindex := 3
               else if fieldbyname('can1_bung3_jindan').asString = '3' then Can1_Bung3_Jindan_2018.Itemindex := 4
               else if fieldbyname('can1_bung3_jindan').asString = '4' then Can1_Bung3_Jindan_2018.Itemindex := 5
               else if fieldbyname('can1_bung3_jindan').asString = '5' then Can1_Bung3_Jindan_2018.Itemindex := 6
               else if fieldbyname('can1_bung3_jindan').asString = '6' then Can1_Bung3_Jindan_2018.Itemindex := 7
               else if fieldbyname('can1_bung3_jindan').asString = '7' then Can1_Bung3_Jindan_2018.Itemindex := 8
               else if fieldbyname('can1_bung3_jindan').asString = '8' then Can1_Bung3_Jindan_2018.Itemindex := 9;
            end;
               //Can1_Bung3_Jindan_2018.Itemindex := FieldByName('Can1_Bung3_Jindan').AsInteger  ;
            if (FieldByName('Can1_Bung3_Jindan').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026'))then
               Can1_Bung3_Jindan_2018.Itemindex := 2;

            if FieldByName('Can1_Bung3_Count').AsString  <> '' then
               Can1_Bung3_Count.Itemindex := FieldByName('Can1_Bung3_Count').AsInteger;
            if (FieldByName('Can1_Bung3_Count').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026')) then    //2014.04.11 곽윤설
               Can1_Bung3_Count.Itemindex  := 1;

            if FieldByName('Can1_Bung3_Jong').AsString  <> '' then
            begin
               if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  <> '0' then
               begin
                    Can1_Bung3_Jong1.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '1' then  Can1_Bung3_Jong1_H.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '2' then  Can1_Bung3_Jong1_M.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,1,1)  = '3' then  Can1_Bung3_Jong1_L.checked  := True;
               end;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,2,1)  = '1' then Can1_Bung3_Jong2.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,3,1)  = '1' then Can1_Bung3_Jong3.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  <> '0' then
               begin
                    Can1_Bung3_Jong4.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  = '1' then  Can1_Bung3_Jong4_H.checked  := True;
                    if copy(FieldByName('Can1_Bung3_Jong').AsString,4,1)  = '2' then  Can1_Bung3_Jong4_L.checked  := True;
               end;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,5,1)  = '1' then Can1_Bung3_Jong5.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,6,1)  = '1' then Can1_Bung3_Jong6.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,7,1)  = '1' then Can1_Bung3_Jong7.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,8,1)  = '1' then Can1_Bung3_Jong8.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,9,1)  = '1' then Can1_Bung3_Jong9.checked  := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,10,1) = '1' then Can1_Bung3_Jong10.checked := True;
               if copy(FieldByName('Can1_Bung3_Jong').AsString,11,1) = '1' then
               begin
                  Can1_Bung3_Jong11.checked    := True;
                  Can1_Bung3_Jong_Direct.Text := FieldByName('Can1_Bung3_Jong_Direct').AsString;
               end;
            end;

            if FieldByName('Can1_Bung3_Etc').AsString  <> '' then
            begin
               if copy(FieldByName('Can1_Bung3_Etc').AsString,1,1) = '1' then Can1_Bung3_Etc1.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,2,1) = '1' then Can1_Bung3_Etc2.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,3,1) = '1' then Can1_Bung3_Etc3.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,4,1) = '1' then Can1_Bung3_Etc4.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,5,1) = '1' then Can1_Bung3_Etc5.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,6,1) = '1' then Can1_Bung3_Etc6.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,7,1) = '1' then Can1_Bung3_Etc7.checked := True;
               if copy(FieldByName('Can1_Bung3_Etc').AsString,8,1) = '1' then
               begin
                  Can1_Bung3_Etc8.checked    := True;
                  Can1_Bung3_Etc_Direct.Text := FieldByName('Can1_Bung3_Etc_Direct').AsString;
               end;
            end;

            Can_Advice.Text := Trim(FieldByName('Can1_Advice').AsString);
            Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can1_Advice1').AsString);
         End
         else if UV_vCod_hm[Newrow - 1] = 'B007' then
         begin
            if FieldByName('Can2_Bung4_Jindan').AsString <> '' then
               Can2_Bung4_Jindan_2018.Itemindex := FieldByName('Can2_Bung4_Jindan').AsInteger;
            if (FieldByName('Can2_Bung4_Jindan').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '151026'))then      //2014.04.11 곽윤설
               Can2_Bung4_Jindan_2018.Itemindex := 3 ;

            if FieldByName('Can2_Bung4_Count').AsString  <> '' then
               Can2_Bung4_Count.Itemindex := FieldByName('Can2_Bung4_Count').AsInteger;
            if (FieldByName('Can2_Bung4_Count').AsString = '') or (FieldByName('ysno_sokun').AsString = 'N') and
               ((GV_sUserID = '150007') or (GV_sUserID = '150681')  or (GV_sUserID = '151026')) then   //2014.04.11 곽윤설
                   Can2_Bung4_Count.Itemindex  := 1;

            if FieldByName('Can2_Bung4_Jong').AsString  <> '' then
            begin
               if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)<> '0' then
               begin
                    Can2_Bung4_Jong1.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '1' then  Can2_Bung4_Jong1_H.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '2' then  Can2_Bung4_Jong1_M.checked  := True;
                    if copy(FieldByName('Can2_Bung4_Jong').AsString,1,1)  = '3' then  Can2_Bung4_Jong1_L.checked  := True;
               end;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,2,1)  = '1' then Can2_Bung4_Jong2.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,3,1)  = '1' then Can2_Bung4_Jong3.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,4,1)  = '1' then Can2_Bung4_Jong4.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,5,1)  = '1' then Can2_Bung4_Jong5.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,6,1)  = '1' then Can2_Bung4_Jong6.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,7,1)  = '1' then Can2_Bung4_Jong7.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,8,1)  = '1' then Can2_Bung4_Jong8.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,9,1)  = '1' then Can2_Bung4_Jong9.checked  := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,10,1) = '1' then Can2_Bung4_Jong10.checked := True;
               if copy(FieldByName('Can2_Bung4_Jong').AsString,11,1) = '1' then
               begin
                  Can2_Bung4_Jong11.checked := True;
                  Can2_Bung4_Jong_Direct.Text := FieldByName('Can2_Bung4_Jong_Direct').AsString;
               end;
            end;

            if FieldByName('Can2_Bung4_Etc').AsString  <> '' then
            begin
               if copy(FieldByName('Can2_Bung4_Etc').AsString,1,1)  = '1' then Can2_Bung4_Etc1.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,2,1)  = '1' then Can2_Bung4_Etc2.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,3,1)  = '1' then Can2_Bung4_Etc3.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,4,1)  = '1' then Can2_Bung4_Etc4.checked  := True;
               if copy(FieldByName('Can2_Bung4_Etc').AsString,5,1)  = '1' then
               begin
                  Can2_Bung4_Etc5.checked := True;
                  Can2_Bung4_Etc_Direct.Text := FieldByName('Can2_Bung4_Etc_Direct').AsString;
               end;
            end;

            Can_Advice.Text := Trim(FieldByName('Can2_Advice').AsString);
            Can_Advice.Text := Can_Advice.Text + Trim(FieldByName('Can2_Advice1').AsString);
         end;
      end
      else
         Can_Advice.Text := '';
   end;
   end;

   if Trim(UV_vCod_sawon[NewRow-1]) = '' then
   begin
      if (copy(GV_sUserId,1,2) = '15') then
      begin
         cmbCOD_SAWON.Font.Color := clRed;
         if      (Copy(UV_vCod_hm[Newrow-1],1,4) = 'B001') or
                 (Copy(UV_vCod_hm[Newrow-1],1,4) = 'B005') or
                 (Copy(UV_vCod_hm[Newrow-1],1,4) = 'B007') or
                 (Copy(UV_vCod_hm[Newrow-1],1,4) = 'P009') then
            GP_ComboDisplay(cmbCOD_SAWON, '151021', 6) //최재현
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
//
//   GP_ComboDisplay(cmbCOD_DOCT,  UV_vCod_doct[NewRow-1], 6);

//외부의사 추가 부분 
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


   mskDAT_LAST.Text      := UV_vDat_last[NewRow-1];
   EDT_TIME_S.Text       := UV_vDat_TIME[NewRow-1];
   EDT_TIME_R.Text       := UV_vDat_TIME_R[NewRow-1];

   if Trim(UV_vDat_result[NewRow-1]) = '' Then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         mskDat_Result.Font.Color := clRed;
         mskDat_Result.Text       := GV_sToday;
      end
      else
      begin
       mskDat_Result.Text := '';
      end;
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
   retDesc_sokun.SelStart := length(retDesc_sokun.Lines.Text);

   edtDESC_BUWI.Text     := UV_vDesc_buwi[NewRow-1];
   edtDESC_BANG.Text     := UV_vDesc_bang[NewRow-1];
   edtGUBN_PAN.Text      := UV_vGubn_pan[NewRow-1];
   edtCod_sokun.Text     := UV_vCod_sokun[NewRow-1];
   if UV_vCell_level[NewRow-1] <> '' then cmb_level.ItemIndex   := strtoInt(UV_vCell_level[NewRow-1]) - 1 
   else cmb_level.ItemIndex   := 1;

   Edt_floor.Text        := UV_vFloor[NewRow-1];
   //슬라이드 렌트
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

   //if  UV_vDat_Rent[NewRow-1] <> '' then msk_dat_rent.text   := copy(UV_vDat_Rent[NewRow-1],1,4)+ copy(UV_vDat_Rent[NewRow-1],6,2) + copy(UV_vDat_Rent[NewRow-1],9,2)
   if  UV_vDat_Rent[NewRow-1] <> '' then msk_dat_rent.text   := UV_vDat_Rent[NewRow-1]
   else  msk_dat_rent.text     := copy(GV_sToday, 1,4) + '-' + copy(GV_sToday, 5,2) + '-' +copy(GV_sToday, 7,2) ;

   msk_dat_return.text   := UV_vDAT_Return[NewRow-1];
   Edt_cmt.Text          := UV_vCmt[NewRow-1];
   //20170724 박예진 선임 이전결과 조회는 B001,B007, P009만 P009는 조직진단이랑 병리 조직검사개수가 나오지 않음.
   sPsokun := '';
   Memo_Psokun.Text := '';
   edt_jindan.Text := '';
   edt_jindan_count.Text := '';
   edt_buwi.text := '';
   edt_pDoctor.text := '';
   edt_pdat_gmgn.text := '';

   if (UV_vCod_hm[NewRow-1] = 'B001') or (UV_vCod_hm[NewRow-1] = 'B007') or (UV_vCod_hm[NewRow-1] = 'P009') then
   with qryPsokun do
   begin
          qryPsokun.Active := False;
          qryPsokun.ParamByName('cod_jisa').AsString    := UV_vCod_jisa[NewRow-1];
          qryPsokun.ParamByName('num_id').AsString      := UV_vNum_id[NewRow-1];
          qryPsokun.ParamByName('cod_hm').AsString      := UV_vCod_hm[NewRow-1];
          qryPsokun.ParamByName('dat_gmgn').AsString    := UV_vDat_gmgn[NewRow-1];
          qryPsokun.Active := True;

          if qryPsokun.IsEmpty = False then
          begin
                edt_pdat_gmgn.Text := fieldByName('dat_gmgn').AsString;
                sPsokun := (fieldByName('desc_sokun1').AsString) + (fieldByName('desc_sokun2').AsString) + (fieldByName('desc_sokun3').AsString) + #13#10 + #13#10 ;
                edt_buwi.Text := fieldByName('desc_buwi').AsString;

                 with qrySawon do
                   qrySawon.Active := False;
                   qrySawon.ParamByName('cod_sawon').AsString    := qryPsokun.fieldbyName('cod_doct').AsString;
                   qrySawon.Active := True;

                edt_pDoctor.text := qrySawon.fieldbyName('desc_name').AsString;

                if qryPsokun.fieldbyName('dat_gmgn').AsString < '20180101' then
                begin
                with qryPCa_cancer do
                   qryPCa_cancer.Active := False;
                   qryPCa_cancer.ParamByName('cod_jisa').AsString    := UV_vCod_jisa[NewRow-1];
                   qryPCa_cancer.ParamByName('num_id').AsString      := UV_vNum_id[NewRow-1];
                   qryPCa_cancer.ParamByName('dat_gmgn').AsString   := qryPsokun.fieldbyName('dat_gmgn').AsString;
                   qryPCa_cancer.Active := True;
                end
                else
                begin
                with qryPCa_cancer2018 do
                     qryPCa_cancer2018.Active := False;
                     qryPCa_cancer2018.ParamByName('cod_jisa').AsString    := UV_vCod_jisa[NewRow-1];
                     qryPCa_cancer2018.ParamByName('num_id').AsString      := UV_vNum_id[NewRow-1];
                     qryPCa_cancer2018.ParamByName('dat_gmgn').AsString   := qryPsokun.fieldbyName('dat_gmgn').AsString;
                     qryPCa_cancer2018.Active := True;
                end;

                   IF UV_vCod_hm[NewRow-1] = 'B001' then
                   begin
                   if qryPsokun.fieldbyName('dat_gmgn').AsString < '20180101' then
                   begin
                     qryPsokun.Active := False;
                     qryPsokun.ParamByName('cod_jisa').AsString    := UV_vCod_jisa[NewRow-1];
                     qryPsokun.ParamByName('num_id').AsString      := UV_vNum_id[NewRow-1];
                     qryPsokun.ParamByName('cod_hm').AsString      := UV_vCod_hm[NewRow-1];
                     qryPsokun.ParamByName('dat_gmgn').AsString    := UV_vDat_gmgn[NewRow-1];
                     qryPsokun.Active := True;

                          if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '1'  then a := '정상'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '2'  then a := '위염'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '3'  then a := '염증성 또는 증식성 병변'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '4'  then a := '저도선종 혹은 이형성'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '5'  then a := '고도선종 혹은 이형성'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '6'  then a := '암의심'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '7'  then a := '암'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_jindan').AsString = '8'  then a := '기타'
                     else a := '';
                          if qryPCa_cancer.fieldByName('Can1_bung3_count').AsString = '1'  then b := '１~３개 '
                     else if qryPCa_cancer.fieldByName('Can1_bung3_count').AsString = '2'  then b := '４~６개'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_count').AsString = '3'  then b := '７~９개 '
                     else if qryPCa_cancer.fieldByName('Can1_bung3_count').AsString = '4'  then b := '10 ~ 12개'
                     else if qryPCa_cancer.fieldByName('Can1_bung3_count').AsString = '5'  then b := '13개 이상'
                     else  b := '' ;
                     edt_jindan.Text := a;
                     edt_jindan_count.Text := b;
                   end
                   else
                   begin
                       qryPCa_cancer.Active := False;

                       qryPsokun.Active := False;
                       qryPsokun.ParamByName('cod_jisa').AsString    := UV_vCod_jisa[NewRow-1];
                       qryPsokun.ParamByName('num_id').AsString      := UV_vNum_id[NewRow-1];
                       qryPsokun.ParamByName('cod_hm').AsString      := UV_vCod_hm[NewRow-1];
                       qryPsokun.ParamByName('dat_gmgn').AsString    := UV_vDat_gmgn[NewRow-1];
                       qryPsokun.Active := True;

                          if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '1'  then a := '이상소견없음'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '21'  then a := '위염'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '22'  then a := '위축성위염/장상피화생'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '3'  then a := '염증성 또는 증식성 병변'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '4'  then a := '저도샘종 혹은 이형성'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '5'  then a := '고도샘종 혹은 이형성'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '6'  then a := '암의심'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '7'  then a := '암'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_jindan').AsString = '8'  then a := '기타'
                     else a := '';
                          if qryPCa_cancer2018.fieldByName('Can1_bung3_count').AsString = '1'  then b := '１~３개 '
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_count').AsString = '2'  then b := '４~６개'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_count').AsString = '3'  then b := '７~９개 '
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_count').AsString = '4'  then b := '10 ~ 12개'
                     else if qryPCa_cancer2018.fieldByName('Can1_bung3_count').AsString = '5'  then b := '13개 이상'
                     else  b := '' ;
                     edt_jindan.Text := a;
                     edt_jindan_count.Text := b;
                   end;
                   end
                   Else if UV_vCod_hm[NewRow-1] = 'B007' then
                   begin
                     if qryPsokun.fieldbyName('dat_gmgn').AsString < '20180101' then
                     begin
                            if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '1'  then a := '정상'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '2'  then a := '염증성 또는 증식성 병변'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '3'  then a := '저도선종 혹은 이형성'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '4'  then a := '고도선종 혹은 이형성'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '5'  then a := '암의심'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '6'  then a := '암'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_jindan').AsString = '7'  then a := '기타'
                       else a := '';

                            if qryPCa_cancer.fieldByName('Can2_bung4_count').AsString = '1'  then b := '１~３개 '
                       else if qryPCa_cancer.fieldByName('Can2_bung4_count').AsString = '2'  then b := '４~６개'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_count').AsString = '3'  then b := '７~９개 '
                       else if qryPCa_cancer.fieldByName('Can2_bung4_count').AsString = '4'  then b := '10 ~ 12개'
                       else if qryPCa_cancer.fieldByName('Can2_bung4_count').AsString = '5'  then b := '13개 이상'
                       else  b := '' ;
                       edt_jindan.Text := a;
                       edt_jindan_count.Text := b;
                     end
                     else
                     begin
                            if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '1'  then a := '정상'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '2'  then a := '염증성 또는 증식성 병변'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '3'  then a := '저도선종 혹은 이형성'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '4'  then a := '고도선종 혹은 이형성'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '5'  then a := '암의심'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '6'  then a := '암'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_jindan').AsString = '7'  then a := '기타'
                       else a := '';

                            if qryPCa_cancer2018.fieldByName('Can2_bung4_count').AsString = '1'  then b := '１~３개 '
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_count').AsString = '2'  then b := '４~６개'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_count').AsString = '3'  then b := '７~９개 '
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_count').AsString = '4'  then b := '10 ~ 12개'
                       else if qryPCa_cancer2018.fieldByName('Can2_bung4_count').AsString = '5'  then b := '13개 이상'
                       else  b := '' ;
                       edt_jindan.Text := a;
                       edt_jindan_count.Text := b;
                     end;
                     end
                   else  sPsokun := sPsokun + '';
                end;

             Memo_Psokun.lines.add(sPsokun);
             Memo_Psokun.Perform(WM_VSCROLL, SB_top, 0);
   end;

   if cmbCod_doct.ItemIndex = -1 then
      cmbCod_doct.ItemIndex := UV_iCod_doct;

   // 2014.05.12 백승현
   // 검사코드 P003 소견코드 입력 & 저장 후 소견코드가 사라지지 않고 보이도록 수정
   if (edtCod_sokun.Text <> '') and ((UV_vCod_hm[NewRow-1] = 'P003') or (UV_vCod_hm[NewRow-1] = 'P012')) then
   begin
     if (GV_sDept <> '04') and (GV_sDept <> '12') and (GV_sDept <> '23') then
       edtCod_sokun.Text := '';
   end
   else
     edtCod_sokun.Text := '';


   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLDAI02.UP_Click(Sender: TObject);
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

   //대장 조직진단
   if (Can2_Bung4_Jindan.ItemIndex = 6) or (Can2_Bung4_Jindan_2018.ItemIndex = 6)  then
   begin
        GroupBox10.Visible := True;
        if Can2_Bung4_Jong1.Checked then
           GroupBox15.Visible := True;
   end
   else
   begin
        GroupBox10.Visible := False;
        GroupBox15.Visible := False;
   end;

   if (Can2_Bung4_Jindan.ItemIndex = 7) or (Can2_Bung4_Jindan_2018.ItemIndex = 7)  then
        GroupBox11.Visible := True
   else GroupBox11.Visible := False;

   if Can2_Bung4_Jong11.Checked then
      Can2_Bung4_Jong_Direct.Visible := True
   else if Can2_Bung4_Jong11.Checked then
      Can2_Bung4_Jong_Direct.Visible := False;

   if Can2_Bung4_Etc5.Checked then
      Can2_Bung4_Etc_Direct.Visible := True
   else if Can2_Bung4_Etc5.Checked then
      Can2_Bung4_Etc_Direct.Visible := False;

   //위 조직진단
   
   if (Can1_Bung3_Jindan.ItemIndex = 7) or (Can1_Bung3_Jindan_2018.ItemIndex = 8) then
   begin
      GroupBox8.Visible := True;
      GroupBox12.Visible := True;
   end
   else
   begin
      GroupBox8.Visible := False;
      GroupBox12.Visible := False;
   end;

   if (Can1_Bung3_Jindan.ItemIndex = 8) or (Can1_Bung3_Jindan_2018.ItemIndex = 9) then
      GroupBox9.Visible := True
   else GroupBox9.Visible := False;

   if Can1_Bung3_Jong11.Checked then
      Can1_Bung3_Jong_Direct.Visible := True
   else if Can1_Bung3_Jong11.Checked then
      Can1_Bung3_Jong_Direct.Visible := False;

   if Can1_Bung3_Etc8.Checked then
      Can1_Bung3_Etc_Direct.Visible := True
   else if Can1_Bung3_Etc8.Checked then
      Can1_Bung3_Etc_Direct.Visible := False;

   UV_bEdit := True;
end;

procedure TfrmLDAI02.rbDateClick(Sender: TObject);
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
      mskCell.Enabled   := False;
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
      mskCell.Enabled   := False;
      mskJumin.SetFocus;
   end
   else if Sender = rbCell then
   begin
      // 찾기기능 비활성화
      btnFind.Enabled   := False;

      mskDate.Color     := clWindow;
      mskDate.Enabled   := False;
      btnDate.Enabled   := False;
      ckMiSokun.Enabled := False;
      mskJumin.Color    := $00E6FFFE;
      mskJumin.Enabled  := False;
      btnJumin.Enabled  := False;
      mskCell.Enabled  := True;
      mskCell.SetFocus;
   end;
end;

procedure TfrmLDAI02.UP_KeyUp(Sender: TObject; var Key: Word;
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

procedure TfrmLDAI02.UP_Exit(Sender: TObject);
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

procedure TfrmLDAI02.btnFindClick(Sender: TObject);
begin
   inherited;

   frmLD1I07F := TfrmLD1I07F.Create(Self);
   frmLD1I07F.ShowModal;
end;

procedure TfrmLDAI02.btnDeleteClick(Sender: TObject);
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

function  TfrmLDAI02.GetCurrLine(Memo : TMemo) : integer;
begin 
 Result := retDesc_sokun.Perform(EM_LINEFROMCHAR, retDesc_sokun.SelStart, 0);
end;

function TfrmLDAI02.UF_insertSokun(sokun : string) : String ;
var
sTemp, sTemp1, sTemp2: String;
iposition, a  : integer;

begin
   sTemp :=  retDesc_sokun.text;
   iposition := retDesc_sokun.SelStart;

   sTemp1 := copy(retDesc_sokun.Lines.Text,1,iposition);
   sTemp2 := copy(retDesc_sokun.Lines.Text,iposition+1,length(retDesc_sokun.Lines.Text));

   retDesc_sokun.Lines.Clear;
   retDesc_sokun.Lines.Text := sTemp1 + sokun +  sTemp2 ;
   //곧바로 컨트롤 기능 쓸경우에 입력 한 부분 뒤에 바로 셀 스타트 자리 잡도록 수정 ..박예진
   //스크롤도 커서 있는 부분으로 박에진....
   retDesc_sokun.SelStart := length( sTemp1 + sokun);
   a := GetCurrLine(retDesc_sokun);
   retDesc_sokun.Perform(EM_LINESCROLL, 0, a);


end;

function TfrmLDAI02.UF_insertSokun_drag(data: string) : String ;
var
sTemp, sTemp1, sTemp2: String;
iposition,a : integer;
begin
   sTemp :=  retDesc_sokun.text;
   iposition := retDesc_sokun.SelStart;

   sTemp1 := copy(retDesc_sokun.Lines.Text,1,iposition);
   sTemp2 := copy(retDesc_sokun.Lines.Text,iposition+1,length(retDesc_sokun.Lines.Text));

   retDesc_sokun.Lines.Clear;
   retDesc_sokun.Lines.Text := sTemp1 + data + ' ' + sTemp2;
   //곧바로 컨트롤 기능 쓸경우에 입력 한 부분 뒤에 바로 셀 스타트 자리 잡도록 수정 ..박예진
   //스크롤도 커서 있는 부분으로 박에진....
   retDesc_sokun.SelStart := length( sTemp1 + data + ' ');
   a := GetCurrLine(retDesc_sokun);
   retDesc_sokun.Perform(EM_LINESCROLL, 0, a);
end;


procedure TfrmLDAI02.edtCod_sokunExit(Sender: TObject);
begin
  inherited;
   qrySokun.Active := false;
   qrySokun.ParamByname('Cod_sokun').AsString := edtCod_sokun.text;
   qrySokun.Active := true;


   if  qrySokun.RecordCount > 0 then
    UF_insertSokun(qrySokun.FieldByName('desc_sbsg').AsString + qrySokun.FieldByName('desc_sbsg1').AsString + qrySokun.FieldByName('desc_sbsg2').AsString)
    else
         GF_MsgBox('Information', ' 없는 소견 코드입니다. ' + Chr(13) +
                                  ' 다시한번 확인하십시오');

    //edtCod_sokun.Text := '';
     if Uppercase(edtCod_sokun.Text) = 'PO01' then
       retDesc_sokun1.text :='자궁세포진 검사상 정상소견입니다. 자궁 경부암의 조기 발견을 위해 1년 간격의 정기적인 검사를 권유합니다. 이 결과는 "Screening Test" 소견 이므로 임상적으로 의심이 되는 병변이나 이상 증상이 관찰되시면 반드시 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO43' then
       retDesc_sokun1.text :='자궁세포진 검사상 선상피내암종 소견이 관찰됩니다. 산부인과 진료를 통한 조직 검사와 적절한 치료가 반드시 필요합니다. 진료의뢰 받으시기 바랍니다.'

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
       retDesc_sokun1.text :='자궁세포진 검사상 비정형편평상피세포(ASC-US)가 관찰됩니다. 산부인과 전문의의 진찰 및 3개월 내 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO12' then
       retDesc_sokun1.text :='자궁세포진 검사상 반응성으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 3-6개월 내 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO13' then
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL) 으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO14' then
       retDesc_sokun1.text :='자궁세포진 검사상 원반세포변화(koilocytotic change)를 동반한 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 인유두종바이러스 검사 (PCR for HPV)를 포함한 추적검사로 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO15' then
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL, 경도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료 의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO16' then
       retDesc_sokun1.text :='자궁세포진 검사상 저등급편평상피내병변 (LSIL, HPV/경도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료 의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO17' then
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL/중등도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO18' then
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL/ 침윤이 의심되는 고도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.'

      else if Uppercase(edtCod_sokun.Text) = 'PO19' then
       retDesc_sokun1.text :='자궁세포진 검사상 고등급평평상피내병변 (HSIL/침윤이 의심되는 상피내암을 포함한 고도의 이형성세포)이 관찰됩니다. 산부인과 진료를 통한 조직검사와 적절한 치료가 필요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO20' then
       retDesc_sokun1.text :='자궁세포진 검사상 편평상피세포암이 관찰됩니다. 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO21' then
       retDesc_sokun1.text :='자궁세포진 검사상 비정형 선세포가 관찰됩니다. 산부인과 진료를 통한 재검사 또는 추적관찰이 필요합니다. 진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO22' then
       retDesc_sokun1.text :='자궁세포진 검사상 암이 의심되는 비정형 선세포가 관찰됩니다. 산부인과 진료를 통한 재검사가 필요합니다. 진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO23' then
       retDesc_sokun1.text :='자궁세포진 검사상 선암종 소견 관찰됩니다. 암 확진을 위해 산부인과 진료를 통한 조직 검사와 적절한 치료가 반드시 필요합니다. 진료의뢰 받으시기 바랍니다.'

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
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HIS/ 중등도 또는 고도의 이형성 세포)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO36' then
       retDesc_sokun1.text :='자궁세포진 검사상 정상소견입니다. 자궁 경부암의 조기 발견을 위해 1년 간격의 정기적인 검사를 권유합니다. 이 결과는 "Screening Test" 소견 이므로 임상적으로 의심이 되는 병변이나 이상 증상이 관찰되시면 반드시 재검사 받으시기 바랍니다.'

     else if (Uppercase(edtCod_sokun.Text) = 'PO37') or (Uppercase(edtCod_sokun.Text) = 'PO50') then
       retDesc_sokun1.text :='자궁세포진 검사상 약간의 위축성세포가 동반된 정상소견을 나타냅니다. 1년 간격으로 정기검진 받으시기 바랍니다.'

     else if (Uppercase(edtCod_sokun.Text) = 'PO38') or (Uppercase(edtCod_sokun.Text) = 'PO51') then
       retDesc_sokun1.text :='자궁세포진 검사상 약간의 위축성세포가 동반된 정상소견을 나타냅니다. 1년 간격으로 정기검진 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO39' then
       retDesc_sokun1.text :='자궁세포진 검사상 비정형편평상피세포(ASC-US)가 관찰됩니다.산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO40' then
       retDesc_sokun1.text :='자궁세포진 검사상 염증을 동반한 반응성(양성)으로 보이는 비정형 편평상피세포가 관찰됩니다. 산부인과 진료와 함께 3~6개월 내 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO41' then
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내병변(HSIL)을 배제하기어려운 비정형 편평상피세포가 관찰됩니다. 산부인과 전문의의 진찰 및 재검하여 결과 확인 요합니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO42' then
       retDesc_sokun1.text :='자궁세포진 검사상 고등급편평상피내 병변(HSIL)이 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO45' then
       retDesc_sokun1.text :='자궁세포진 검사상 수복세포를 포함한 반응성 세포변화가 관찰됩니다.염증 또는 외부자극에 의할 수 있으나 6개월 내 재검사 받으시기 바랍니다.'

     else if Uppercase(edtCod_sokun.Text) = 'PO46' then
       retDesc_sokun1.text :='자궁세포진 검사상 경도 또는 중등도 이형성 세포(CIN 1-2)가 관찰됩니다. 자궁경부암으로 이행할 가능성이 높으므로, 산부인과 진료를 통한 조직생검 및 치료가 반드시 필요합니다. 외래로 방문하여 진료의뢰 받으시기 바랍니다.';


end;

procedure TfrmLDAI02.cmbRelationChange(Sender: TObject);
begin
   inherited;
   if cmbRelation.ItemIndex = 0 then
   begin
      Application.CreateForm(TfrmLD1I071, frmLD1I071);
      frmLD1I071.Show;
   end
   else if cmbRelation.ItemIndex = 1 then
   begin
      Application.CreateForm(TfrmLDAI012, frmLDAI012);
      frmLDAI012.Show;
   end;
end;

procedure TfrmLDAI02.btn_signClick(Sender: TObject);
Var
   index, Max_Num : integer;
begin
  inherited;
  index := grdMaster.CurrentDataRow - 1;
    {  If b_sgCAppSet = False Then              //인증로그인이 안되어 있으면
      begin
         KMIConnect();
         GF_PubCertifyLoad();      //ysys
      end;

      if not UV_bEdit then
      begin
         showmessage('변경된 내용이 없습니다.');
         exit;
      end;  }

      if (UV_vDat_gmgn[index] < '20180101') then
      begin
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
            if (Qry_CK_Gulkwa.IsEmpty = False) and
               ((UV_vCod_hm[index] = 'B001') or
                (UV_vCod_hm[index] = 'B007')) then
            begin
               with Qry_Sign do
               begin
                  Close;
                  ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString   := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'B001' then
                     ParamByName('Program_id').AsString := 'LD1I121'
                  else if UV_vCod_hm[index] = 'B007' then
                     ParamByName('Program_id').AsString := 'LD1I127';
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
                             'dat_time[' + FieldByName('dat_time').AsString + ']' +
                             'dat_result[' + FieldByName('dat_result').AsString + ']' +
                             'dat_time_R[' + FieldByName('dat_time_R').AsString + ']' +
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
                             'cod_update[' + FieldByName('cod_update').AsString + ']' +
                             'cod_jisa_R[' + FieldByName('cod_jisa_R').AsString + ']' +
                             'num_id_R[' + FieldByName('num_id_R').AsString + ']' +
                             'dat_gmgn_R[' + FieldByName('dat_gmgn_R').AsString + ']';

               if UV_vCod_hm[index] = 'B001' then
               GV_sGulkwa := GV_sGulkwa +
                             'can1_bung3_jindan['       + FieldByName('can1_bung3_jindan').AsString + ']' +
                             'can1_bung3_count['        + FieldByName('can1_bung3_count').AsString + ']' +
                             'can1_bung3_jong['         + FieldByName('can1_bung3_jong').AsString + ']' +
                             'can1_bung3_jong_direct['  + FieldByName('can1_bung3_jong_direct').AsString + ']' +
                             'can1_bung3_etc['          + FieldByName('can1_bung3_etc').AsString + ']' +
                             'can1_bung3_etc_direct['   + FieldByName('can1_bung3_etc_direct').AsString + ']' +
                             'can1_advice['             + FieldByName('can1_advice').AsString + ']' +
                             'can1_advice1['            + FieldByName('can1_advice1').AsString + ']';

               if UV_vCod_hm[index] = 'B007' then
               GV_sGulkwa := GV_sGulkwa +
                             'can2_bung4_jindan['       + FieldByName('can2_bung4_jindan').AsString + ']' +
                             'can2_bung4_count['        + FieldByName('can2_bung4_count').AsString + ']' +
                             'can2_bung4_count_Gong['   + FieldByName('can2_bung4_count_Gong').AsString + ']' +
                             'can2_bung4_jong['         + FieldByName('can2_bung4_jong').AsString + ']' +
                             'can2_bung4_jong_direct['  + FieldByName('can2_bung4_jong_direct').AsString + ']' +
                             'can2_bung4_etc['          + FieldByName('can2_bung4_etc').AsString + ']' +
                             'can2_bung4_etc_direct['   + FieldByName('can2_bung4_etc_direct').AsString + ']' +
                             'can2_advice['             + FieldByName('can2_advice').AsString + ']' +
                             'can2_advice1['            + FieldByName('can2_advice1').AsString + ']';


               GV_sGulkwa := GV_sGulkwa +
                             'dat_insert_R[' + FieldByName('dat_insert_R').AsString + ']' +
                             'cod_insert_R[' + FieldByName('cod_insert_R').AsString + ']' +
                             'dat_update_R[' + FieldByName('dat_update_R').AsString + ']' +
                             'cod_update_R[' + FieldByName('cod_update_R').AsString + ']';
               GF_PubSignCertify;

               with QryI_Sign do
               begin
                  ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString     := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'B001' then
                     ParamByName('Program_id').AsString := 'LD1I121'
                  else if UV_vCod_hm[index] = 'B007' then
                     ParamByName('Program_id').AsString := 'LD1I127';
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
      end;
      end
      else if (UV_vDat_gmgn[index] >= '20180101') then
      begin
      if (b_sgCAppSet = True) and (Index >= 0) Then
      Begin
         //저장  프로시져 호출
         UF_Save;
         with Qry_CK_Gulkwa_2018 do
         begin
            Close;
            Qry_CK_Gulkwa_2018.ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Index];
            Qry_CK_Gulkwa_2018.ParamByName('Num_id').AsString   := UV_vNum_id[index];
            Qry_CK_Gulkwa_2018.ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[index];
            Qry_CK_Gulkwa_2018.ParamByName('COD_CELL').AsString := UV_vCod_cell[index];
            Qry_CK_Gulkwa_2018.ParamByName('COD_HM').AsString   := UV_vCod_hm[index];

            Open;
            if (Qry_CK_Gulkwa_2018.IsEmpty = False) and
               ((UV_vCod_hm[index] = 'B001') or
                (UV_vCod_hm[index] = 'B007')) then
            begin
               with Qry_Sign do
               begin
                  Close;
                  ParamByName('Cod_jisa').AsString := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString   := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'B001' then
                     ParamByName('Program_id').AsString := 'LD1I121'
                  else if UV_vCod_hm[index] = 'B007' then
                     ParamByName('Program_id').AsString := 'LD1I127';
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
                             'dat_time[' + FieldByName('dat_time').AsString + ']' +
                             'dat_time_R[' + FieldByName('dat_time_R').AsString + ']' +
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
                             'cod_update[' + FieldByName('cod_update').AsString + ']' +
                             'cod_jisa_R[' + FieldByName('cod_jisa_R').AsString + ']' +
                             'num_id_R[' + FieldByName('num_id_R').AsString + ']' +
                             'dat_gmgn_R[' + FieldByName('dat_gmgn_R').AsString + ']';

               if UV_vCod_hm[index] = 'B001' then
               GV_sGulkwa := GV_sGulkwa +
                             'can1_bung3_jindan['       + FieldByName('can1_bung3_jindan').AsString + ']' +
                             'can1_bung3_count['        + FieldByName('can1_bung3_count').AsString + ']' +
                             'can1_bung3_jong['         + FieldByName('can1_bung3_jong').AsString + ']' +
                             'can1_bung3_jong_direct['  + FieldByName('can1_bung3_jong_direct').AsString + ']' +
                             'can1_bung3_etc['          + FieldByName('can1_bung3_etc').AsString + ']' +
                             'can1_bung3_etc_direct['   + FieldByName('can1_bung3_etc_direct').AsString + ']' +
                             'can1_advice['             + FieldByName('can1_advice').AsString + ']' +
                             'can1_advice1['            + FieldByName('can1_advice1').AsString + ']';

               if UV_vCod_hm[index] = 'B007' then
               GV_sGulkwa := GV_sGulkwa +
                             'can2_bung4_jindan['       + FieldByName('can2_bung4_jindan').AsString + ']' +
                             'can2_bung4_count['        + FieldByName('can2_bung4_count').AsString + ']' +
                             'can2_bung4_count_Gong['   + FieldByName('can2_bung4_count_Gong').AsString + ']' +
                             'can2_bung4_jong['         + FieldByName('can2_bung4_jong').AsString + ']' +
                             'can2_bung4_jong_direct['  + FieldByName('can2_bung4_jong_direct').AsString + ']' +
                             'can2_bung4_etc['          + FieldByName('can2_bung4_etc').AsString + ']' +
                             'can2_bung4_etc_direct['   + FieldByName('can2_bung4_etc_direct').AsString + ']' +
                             'can2_advice['             + FieldByName('can2_advice').AsString + ']' +
                             'can2_advice1['            + FieldByName('can2_advice1').AsString + ']';


               GV_sGulkwa := GV_sGulkwa +
                             'dat_insert_R[' + FieldByName('dat_insert_R').AsString + ']' +
                             'cod_insert_R[' + FieldByName('cod_insert_R').AsString + ']' +
                             'dat_update_R[' + FieldByName('dat_update_R').AsString + ']' +
                             'cod_update_R[' + FieldByName('cod_update_R').AsString + ']';
               GF_PubSignCertify;

               with QryI_Sign do
               begin
                  ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString     := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[index];
                  if UV_vCod_hm[index] = 'B001' then
                     ParamByName('Program_id').AsString := 'LD1I121'
                  else if UV_vCod_hm[index] = 'B007' then
                     ParamByName('Program_id').AsString := 'LD1I127';
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

function TfrmLDAI02.UF_B201_Jong1(bData, bData_H, bData_M, bData_L : Boolean) : String;
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
function TfrmLDAI02.UF_B201_Jong4(bData, bData_H, bData_L : Boolean) : String;
begin
   if bData then
   begin
      if bData_H then Result := '1'
      else if bData_L then Result := '2'
      else  Result := '1';
   end
   else          Result := '0';
end;
function TfrmLDAI02.UF_B201(bData : Boolean) : String;
begin
   if bData then Result := '1'
   else          Result := '0';
end;
Function TfrmLDAI02.UF_Check(iTemp : Integer) : Integer;
begin
   if iTemp >= 0 then Result := iTemp
   else               Result := 0;
end;
Function TfrmLDAI02.UF_Check01(iTemp : Integer) : String;
begin
   if (iTemp >= 0) and
      (iTemp < 10) then      Result := '0' + intToStr(iTemp)
   else if iTemp >= 10 then  Result := intToStr(iTemp)
   else                      Result := '0';
end;

procedure TfrmLDAI02.CmbSang_Form1_gubnKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['1','2',#8, #13]) then
  begin
     showmessage('입력 값을 다시 확인해주십시오.');
     key := #0;
  end;
end;

procedure TfrmLDAI02.btnPrintClick(Sender: TObject);
begin
  inherited;
    Application.CreateForm(TfrmLD1P12, frmLD1P12);
    frmLD1P12.Show;
end;

procedure TfrmLDAI02.Button1Click(Sender: TObject);
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

         if Chk_order.Checked then sWhere := sWhere + 'and A.gubn_order = ''Y''';
      end
      else if rbRDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_last= ''' + mskRDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_last = ''' + mskRDate.Text + '''';

         if Chk_order.Checked then sWhere := sWhere + 'and A.gubn_order = ''Y''';
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
           if cmbGubn.ItemIndex = 0  then sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''J'''
      else if cmbGubn.ItemIndex = 1  then sWhere := sWhere + ' AND cod_hm = ''B001'''
      else if cmbGubn.ItemIndex = 2  then sWhere := sWhere + ' AND cod_hm = ''B005'''
      else if cmbGubn.ItemIndex = 3  then sWhere := sWhere + ' AND cod_hm = ''B007'''
      else if cmbGubn.ItemIndex = 4  then sWhere := sWhere + ' AND cod_hm = ''B012'''
      else if cmbGubn.ItemIndex = 5  then sWhere := sWhere + ' AND cod_hm = ''P009''';

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
             ParamByName('desc_buwi').AsString := UV_vDesc_buwi[grdMaster.CurrentDataRow-1];
             ParamByName('num_id').AsString    := UV_vNum_id[grdMaster.CurrentDataRow-1];
             ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[grdMaster.CurrentDataRow-1];

             If btn_rent.Checked Then
             begin
                ParamByName('rent_yn').AsString := 'Y';
                ParamByName('DAT_RENT').AsString    := msk_dat_rent.Text;
                ParamByName('DAT_RETURN').AsString  := msk_dat_return.text;
                ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
             end
             Else If btn_return.Checked Then
             begin
                ParamByName('rent_yn').AsString := 'N';
                ParamByName('DAT_RENT').AsString    := msk_dat_rent.Text;
                ParamByName('DAT_RETURN').AsString  := msk_dat_return.text;
                ParamByName('cell_cmt').AsString    := Edt_cmt.Text;
             end
             else
             begin
               GF_MsgBox('Warning', '대여상태를 입력해주세요');
               exit;
             end;
             Execsql;
             GP_query_log(qryU_ca_cancer2009_Can5, 'LDAI01', 'U', 'Y');
          end;
      end
      else
      begin
            with qryI_slide do
            begin
               ParamByName('desc_buwi').AsString := UV_vDesc_buwi[grdMaster.CurrentDataRow-1];
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
procedure TfrmLDAI02.retDesc_SokunKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var i, iposition : Integer;

begin
  // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

  inherited;
  if (ssCtrl in Shift) And (Char(key) in ['A','a']) then
  begin
      if UV_vCod_hm[i] = 'B001' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL1';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end
      else if UV_vCod_hm[i] = 'B007' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL5';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end;
  end;

  if (ssCtrl in Shift) And (Char(key) in ['S','s']) then
  begin
      if UV_vCod_hm[i] = 'B001' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL2';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end
      else if UV_vCod_hm[i] = 'B007' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL6';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end;
  end;

  if (ssCtrl in Shift) And (Char(key) in ['D','d']) then
  begin
      if UV_vCod_hm[i] = 'B001' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL3';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end
      else if UV_vCod_hm[i] = 'B007' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL7';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end;
  end;

  if (ssCtrl in Shift) And (Char(key) in ['Q','q']) then
  begin
      if UV_vCod_hm[i] = 'B007' then
      begin
         frmLDAI021 := TfrmLDAI021.Create(self);
         frmLDAI021.edt_cod_group.text  := 'CEL4';
         frmLDAI021.ShowModal;

         UF_insertSokun_drag(edt_data.text);
         edt_data.text := '';
         exit;
      end;
  end;

end;

procedure TfrmLDAI02.rbRDateClick(Sender: TObject);
begin
  inherited;
  if Sender = rbRDate then
   begin
      // 찾기기능 활성화
      btnFind.Enabled   := True;

      mskRDate.Color     := $00E6FFFE;
      mskRDate.Enabled   := True;
      btnRDate.Enabled   := True;
      mskDate.Color     := clWindow;
      mskDate.Enabled   := False;
      btnDate.Enabled   := False;
      ckMiSokun.Enabled := True;
      mskJumin.Color    := clWindow;
      mskJumin.Enabled  := False;
      btnJumin.Enabled  := False;
      mskCell.Enabled   := False;
      mskRDate.SetFocus;
   end ;

end;

end.
