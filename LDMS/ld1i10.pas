//==============================================================================
// 프로그램 설명 : [센터]자체검진 생화학결과등록
// 시스템        : LDMS
// 수정일자      : 2011.12.14
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2011.12.14
// 수정자        : 유동구
// 수정내용      : [20120516-유동구]특수검진은 제외.(연구소+센터인 경우)
//==============================================================================
// 수정일자      : 2013.5.9
// 수정자        : 김희석
// 수정내용      : 키, 몸무게,임신및 생리
//=============================================================================
// 수정일자      : 2013.6.15
// 수정자        : 김승철
// 수정내용      : 임신중 안나오는 부분 수정
//=============================================================================
// 수정일자      : 2014.04.30
// 수정자        : 곽윤설
// 수정내용      : 식전자가혈당측정기(C077) Pos_End값 변경 (356->355)
//=============================================================================
// 수정일자      : 2014.06.20
// 수정자        : 곽윤설
// 수정내용      : 생화학 C074 센터 자체검사
// 참고사항      : [본원-최은희]
//=============================================================================
// 수정일자      : 2014.07.02
// 수정자        : 곽윤설
// 수정내용      : gulkwa1, gulkwa2 중복오류 수정
// 참고사항      :
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
// 수정일자      : 2015.04.23
// 수정자        : 곽윤설
// 수정내용      : 특검 재검일때 LDL-콜레스테롤(효소)(C074) 센터 자체검사
// 참고사항      :
//=============================================================================
// 수정일자      : 2015.06.15
// 수정자        : 곽윤설
// 수정내용      : 노신,성인병 재검 -> 혈액 프로파일도 조회 가능 (1년에 1~2건 정도 있음..)
// 참고사항      :
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
// 수정내용      : 생화학결과 출력 (LD1P101)
// 참고사항      : 한의 부산진단검사의학팀1500839
//=============================================================================
// 수정일자      : 2015.11.03
// 수정자        : 곽윤설
// 수정내용      : 노신재검&C032 검사 정리
//                 노신/성인병/생애 재검(2차)이면서 공복시 혈당(C032) 검사 있을 경우 C032는 센터자체검사
//                 노신재검과 특검재검이 같이 있을 경우 노신재검이 우선순위, C032 센터자체검사
//                 특검재검만 있을 경우에는 C032 연구소 검사로 진행함.
// 참고사항      : 문지혜 선임님 확인
//=============================================================================
// 수정일자      : 2017.08.25
// 수정자        : 박수지
// 수정내용      :  c074 결과 있을 시에 c027에 동일결과 자동 저장
// 참고사항      :
//=============================================================================
// 수정일자      : 2019.04.08
// 수정자        : 박수지
// 수정내용      : c083결과창 생성
// 참고사항      :
//=============================================================================
unit LD1I10;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I10 = class(TfrmSingle)
    qryGlkwa: TQuery;
    grdMaster: TtsGrid;
    qryU_Glkwa: TQuery;
    pnlDetail: TPanel;
    edtValue1: TEdit;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlName1: TPanel;
    pnlCond: TPanel;
    btnDate: TSpeedButton;
    btnJumin: TSpeedButton;
    edtName: TEdit;
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
    pnlNum9: TPanel;
    pnlName9: TPanel;
    edtValue9: TEdit;
    Panel33: TPanel;
    Panel34: TPanel;
    mskDAT_BUNJU: TDateEdit;
    edtNUM_BUNJU: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    mskNUM_JUMIN: TMaskEdit;
    Label6: TLabel;
    edtDESC_NAME: TEdit;
    qryHangmok: TQuery;
    qryGmgn: TQuery;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    edtFind: TEdit;
    Label2: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    Label11: TLabel;
    mskDAT_GMGN: TDateEdit;
    btnNextNum: TBitBtn;
    qryU_Bunju: TQuery;
    valChange: TValEdit;
    btnPrevNum: TBitBtn;
    qryPart01: TQuery;
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
    edtPValue1: TEdit;
    edtPValue2: TEdit;
    edtPValue3: TEdit;
    edtPValue4: TEdit;
    edtPValue5: TEdit;
    edtPValue6: TEdit;
    edtPValue7: TEdit;
    edtPValue8: TEdit;
    edtPValue9: TEdit;
    qryPrev: TQuery;
    qryinjouk: TQuery;
    Label7: TLabel;
    Edt_smp: TEdit;
    Label12: TLabel;
    Edt_ID: TEdit;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    qryJaegumhm: TQuery;
    qryNo_hangmok: TQuery;
    Label13: TLabel;
    Edt_bjjs: TEdit;
    Cmb_gubn: TComboBox;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    mskPDAT_BUNJU: TDateEdit;
    Label17: TLabel;
    edtPNUM_BUNJU: TEdit;
    pnlNum10: TPanel;
    pnlName10: TPanel;
    edtPValue10: TEdit;
    edtValue10: TEdit;
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
    qryinjouk1: TQuery;
    qryGum_bul: TQuery;
    Label21: TLabel;
    edt_sinjang: TEdit;
    Label22: TLabel;
    edt_chejung: TEdit;
    Label23: TLabel;
    edt_gumgin_Check: TEdit;
    qryGumgin_Check: TQuery;
    qrykicho: TQuery;
    pnlNum11: TPanel;
    Panel4: TPanel;
    edtPValue11: TEdit;
    edtValue11: TEdit;
    Label18: TLabel;
    rbJumin: TRadioButton;
    mskJumin: TMaskEdit;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
    RG_sort: TRadioGroup;
    rbDate: TRadioButton;
    mskDate: TDateEdit;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    EdtFind2: TEdit;
    mskNum: TMaskEdit;
    qryHmCheck: TQuery;
    chk_CRM: TCheckBox;
    qryI_Check_CRM: TQuery;
    qry_Check_CRM: TQuery;
    qryU_check_CRM: TQuery;
    Panel2: TPanel;
    Panel8: TPanel;
    edtPValue12: TEdit;
    edtValue12: TEdit;
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
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Exit(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure edtFindExit(Sender: TObject);
    procedure mskNumExit(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure UP_MoveNum(Sender: TObject);
    procedure edtValue1Change(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure EdtFind2Exit(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_jisa,   UV_vNum_id,     UV_vNum_jumin,  UV_vDesc_name,  UV_vDat_gmgn,    UV_vNum_samp,
    UV_vCod_hul,    UV_vCod_cancer, UV_vCod_jangbi, UV_vCod_chuga,
    UV_vGubn_nosin, UV_vGubn_nsyh,  UV_vGubn_adult, UV_vGubn_adyh,
    UV_vGubn_life,  UV_vGubn_lfyh,  UV_vGubn_bogun, UV_vGubn_bgyh,
    UV_vGubn_agab,  UV_vGubn_agyh,  UV_vGubn_tkgm,
    UV_vCod_bjjs,   UV_vDat_bunju,  UV_vNum_bunju,  UV_vDesc_glkwa, UV_vDesc_Pglkwa, UV_vPart01, 
    UV_vC009, UV_vC010, UV_vC011, UV_vC032, UV_vC033, UV_vC025, UV_vC026, UV_vC027, UV_vC028, UV_vC042, UV_vC074, UV_vC077, UV_vGumgin_check : Variant;

    iCntCheck : Integer;

    // 항목코드, 시작위치, 마지막위치, 값
    UV_sValue : array[1..6] of String;
    UV_sGubn  : array[1..6] of String;
    UV_sUnit  : array[1..6] of String;

    // part구분
    UV_sGubn_part : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_EditDisplay(Sender : TObject);
    
    function  UF_Save : Boolean;
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;

    procedure UP_DChange(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmLD1I10: TfrmLD1I10;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1P101;

{$R *.DFM}

function  TfrmLD1I10.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I10.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD1I10.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading := '분주번호';
         Col[2].Heading := '혈액샘플No';
         Col[3].Heading := '검진자명';
      end
      else if rbJumin.Checked then
      begin
         Col[1].Heading := '검진일자';
         Col[2].Heading := '혈액샘플No';
         Col[3].Heading := '검진자명';
      end;
      // Set Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
   end;
end;

procedure TfrmLD1I10.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_glkwa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pglkwa := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vPart01      := VarArrayCreate([0, 0], varOleStr);
      
      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nosin  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nsyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adult  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_life   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_lfyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bogun  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bgyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agab   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_tkgm   := VarArrayCreate([0, 0], varOleStr);

      UV_vC009        := VarArrayCreate([0, 0], varOleStr);
      UV_vC010        := VarArrayCreate([0, 0], varOleStr);
      UV_vC011        := VarArrayCreate([0, 0], varOleStr);
      UV_vC032        := VarArrayCreate([0, 0], varOleStr);
      UV_vC033        := VarArrayCreate([0, 0], varOleStr);
      UV_vC025        := VarArrayCreate([0, 0], varOleStr);
      UV_vC026        := VarArrayCreate([0, 0], varOleStr);
      UV_vC027        := VarArrayCreate([0, 0], varOleStr);
      UV_vC028        := VarArrayCreate([0, 0], varOleStr);
      UV_vC042        := VarArrayCreate([0, 0], varOleStr);
      UV_vC074        := VarArrayCreate([0, 0], varOleStr);
      UV_vC077        := VarArrayCreate([0, 0], varOleStr);
      UV_vGumgin_check := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vDat_bunju,   iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_glkwa,  iLength);
      VarArrayRedim(UV_vDesc_Pglkwa, iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);
      VarArrayRedim(UV_vPart01,      iLength);

      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_cancer,  iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vGubn_nosin,  iLength);
      VarArrayRedim(UV_vGubn_nsyh,   iLength);
      VarArrayRedim(UV_vGubn_adult,  iLength);
      VarArrayRedim(UV_vGubn_adyh,   iLength);
      VarArrayRedim(UV_vGubn_life,   iLength);
      VarArrayRedim(UV_vGubn_lfyh,   iLength);
      VarArrayRedim(UV_vGubn_bogun,  iLength);
      VarArrayRedim(UV_vGubn_bgyh,   iLength);
      VarArrayRedim(UV_vGubn_agab,   iLength);
      VarArrayRedim(UV_vGubn_agyh,   iLength);
      VarArrayRedim(UV_vGubn_tkgm,   iLength);

      VarArrayRedim(UV_vC009,        iLength);
      VarArrayRedim(UV_vC010,        iLength);
      VarArrayRedim(UV_vC011,        iLength);
      VarArrayRedim(UV_vC032,        iLength);
      VarArrayRedim(UV_vC033,        iLength);
      VarArrayRedim(UV_vC025,        iLength);
      VarArrayRedim(UV_vC026,        iLength);
      VarArrayRedim(UV_vC027,        iLength);
      VarArrayRedim(UV_vC028,        iLength);
      VarArrayRedim(UV_vC042,        iLength);
      VarArrayRedim(UV_vC074,        iLength);
      VarArrayRedim(UV_vC077,        iLength);
      VarArrayRedim(UV_vGumgin_check, iLength);
   end;
end;

function TfrmLD1I10.UF_Condition : Boolean;
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

function TfrmLD1I10.UF_Save : Boolean;
var i, iTemp : Integer;
    Part01 , Part02 : String;
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

                GP_query_log(qryU_Check_CRM, 'LD1I10', 'U', 'Y');
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

                GP_query_log(qryU_Check_CRM, 'LD1I10', 'U', 'Y');
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

             GP_query_log(qryI_Check_CRM, 'LD1I10', 'I', 'Y');
          end;
      end;
   end;
   // 작업중인지 Check
   if not UV_bEdit  then exit;

   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;



   with qryPart01 do
   begin
      Close;
      qryPart01.ParamByName('dat_bunju').AsString := mskDAT_BUNJU.Text;
      qryPart01.ParamByName('num_bunju').AsString := edtNUM_BUNJU.Text;
      Open;

      if qryPart01.RecordCount > 0 then
      begin
         Part01 := '';
         if FieldByName('desc_glkwa1').AsString <> '' then
         begin
            for iTemp := 1 to 200 - Length(FieldByName('desc_glkwa1').AsString) do
                Part01 := Part01 + ' ';
            Part01 := FieldByName('desc_glkwa1').AsString + Part01;
         end
         else
         begin
            for iTemp := 1 to 2 do
                Part01 := Part01 + '                                                                                                    ';
         end;
         Part02 := '';
         if FieldByName('desc_glkwa2').AsString <> '' then
         begin
            for iTemp := 1 to 200 - Length(FieldByName('desc_glkwa2').AsString) do
                Part02 := Part02 + ' ';
            Part02 := FieldByName('desc_glkwa2').AsString + Part02;
         end
         else
         begin
            for iTemp := 1 to 2 do
                Part02 := Part02 + '                                                                                                    ';
         end;
      end;
   end;

   // SGOT(AST)	          C009	        01	09 	49	54
   Part01 := Copy(Part01, 1, 48)  + GF_RPad(edtValue1.text, 6, ' ') + Copy(Part01, 55, Length(Part01));
   //  SGPT(ALT)	  C010  	01	10 	55	60
   Part01 := Copy(Part01, 1, 54)  + GF_RPad(edtValue2.text, 6, ' ') + Copy(Part01, 61, Length(Part01));
   //  r-GTP     	  C011	        01	11 	61	66
   Part01 := Copy(Part01, 1, 60)  + GF_RPad(edtValue3.text, 6, ' ') + Copy(Part01, 67, Length(Part01));
   //  콜레스테롤	  C025	        01	21 	121	126
   Part01 := Copy(Part01, 1, 120) + GF_RPad(edtValue4.text, 6, ' ') + Copy(Part01, 127, Length(Part01));

   //  HDL-콜레스테롤	  C026	        01	22 	127	132
   Part01 := Copy(Part01, 1, 126) + GF_RPad(edtValue5.text, 6, ' ') + Copy(Part01, 133, Length(Part01));
   //  LDL-콜레스테롤	  C027	        01	23 	133	138
   Part01 := Copy(Part01, 1, 132) + GF_RPad(edtValue6.text, 6, ' ') + Copy(Part01, 139, Length(Part01));
   //  중성지방  	  C028	        01	24 	139	144
   Part01 := Copy(Part01, 1, 138) + GF_RPad(edtValue7.text, 6, ' ') + Copy(Part01, 145, Length(Part01));
   //  공복시혈당	  C032	        01	27 	157	162
   Part01 := Copy(Part01, 1, 156) + GF_RPad(edtValue8.text, 6, ' ') + Copy(Part01, 163, Length(Part01));
   //  크레아티닌	  C042	        01	32 	187	192
   Part01 := Copy(Part01, 1, 186) + GF_RPad(edtValue9.text, 6, ' ') + Copy(Part01, 193, Length(Part01));

   //LDL-콜레스테롤(효소) C074          01      56      331     336
   Part02 := Copy(Part01 + Part02, 1, 330) + GF_RPad(edtValue10.text, 6, ' ') + Copy(Part01 + Part02, 337, Length(Part01 + Part02));  //20140620 곽윤설
   //식전 자가혈당측정기  C077	        01	59 	349	354
   Part02 := Copy(Part02, 1, 348) + GF_RPad(edtValue11.text, 6, ' ') + Copy(Part02, 355, Length(Part02));   //20140430 356->355  //20140702  Part01 + Part02 -> Part2
   //  식후두시간혈당	  C033	        01	27 	163	168
   Part01 := Copy(Part01, 1, 162) + GF_RPad(edtValue12.text, 6, ' ') + Copy(Part01, 169, Length(Part01));

   Part01 := 'A' + Part01;
   Part01 := Trim(Part01);
   Part01 := Copy(Part01, 2, Length(Part01)-1);

   Part02 := 'A' + Part02;
   Part02 := Trim(Part02);
   Part02 := Copy(Part02, 202, Length(Part02)-1);

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   dmComm.database.StartTransaction;
   try
      with qryU_Glkwa do
      begin
         ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString   := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString   := edtNUM_BUNJU.Text;
         ParamByName('GUBN_PART').AsString   := '01';
         ParamByName('desc_glkwa1').AsString := Part01;
         ParamByName('desc_glkwa2').AsString := Part02;
         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         Execsql;

         GP_query_log(qryU_Glkwa, 'LD1I10', 'U', 'Y');
      end;

      with qryU_Bunju do
      begin
         ParamByName('COD_jisa').AsString   := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString  := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString  := edtNUM_BUNJU.Text;
         Execsql;

         GP_query_log(qryU_Bunju, 'LD1I10', 'U', 'Y');
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
      UV_vPart01[i] := Part01;
      UV_vC009[i]   := Trim(copy(Part01,  49, 6));
      UV_vC010[i]   := Trim(copy(Part01,  55, 6));
      UV_vC011[i]   := Trim(copy(Part01,  61, 6));
      UV_vC025[i]   := Trim(copy(Part01, 121, 6));
      UV_vC032[i]   := Trim(copy(Part01, 157, 6));
      UV_vC033[i]   := Trim(copy(Part01, 163, 6));

      UV_vC026[i]   := Trim(copy(Part01, 127, 6));
      UV_vC027[i]   := Trim(copy(Part01, 133, 6));
      UV_vC028[i]   := Trim(copy(Part01, 139, 6));
      UV_vC042[i]   := Trim(copy(Part01, 187, 6));

      UV_vC074[i]   := Trim(copy(Part02, 131, 6));    //20140620 곽윤설
      UV_vC077[i]   := Trim(copy(Part02, 149, 6));    //200뺀값으로 위치값 넣음.
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD1I10.btnSaveClick(Sender: TObject);
begin
   inherited;

   if (edtValue10.text <> '') and (edtValue6.Enabled = true) then
   begin
         edtValue6.Text := edtValue10.text;
         UV_bEdit := True;
   end;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I10.FormCreate(Sender: TObject);
var i : integer;
begin
   inherited;

   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   UV_sBasicSQL := qryGlkwa.SQL.Text;
   Cmb_gubn.ItemIndex := 0
end;

procedure TfrmLD1I10.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : begin
             if rbDate.Checked then
                Value := UV_vNum_bunju[DataRow-1]
             else if rbJumin.Checked then
                Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
          end;
      2 : Value := UV_vNum_Samp[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
   end;
end;

procedure TfrmLD1I10.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrderBy : String;
    chk_jisagum,chk_jisagum2  : boolean;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;
   // Grid 초기화
   grdMaster.Rows := 0;
   // Grid 환경설정
   UP_GridSet;

   // 분주번호에 0을 채워줌
   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> ''   then mskTo.Text   := GF_LPad(Trim(mskTo.Text), 4, '0');

   // Query 수행 & 배열에 저장
   index := 0;
   with qryGlkwa do
   begin
      Active := False;

      sSelect := ' SELECT A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.gubn_part,            ' +
                 '        A.desc_glkwa1, A.desc_glkwa2, A.desc_glkwa3, B.desc_dept, B.num_samp,   ' +
                 '        B.num_jumin, B.desc_name, B.cod_dc, B.dat_gmgn, B.gubn_hulgum,          ' +
                 '        B.cod_hul, B.cod_cancer, B.cod_chuga, B.cod_jangbi, B.cod_jisa,         ' +
                 '        B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh, B.Num_samp,       ' +
                 '        B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, B.gubn_tkgm,       ' +
                 '        B.gubn_gong, B.gubn_goyh, B.Gubn_life, B.Gubn_lfyh, B.gubn_preg,        ' +
                 '        B.cod_jangbi, B.cod_hul, C.cod_chuga as TK_chuga                        ' +
                 ' FROM   dbo.gulkwa A WITH(NOLOCK) INNER JOIN dbo.gumgin B WITH(NOLOCK)          ' +
                 ' ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ' +
                 ' LEFT OUTER JOIN dbo.tkgum C WITH(NOLOCK) ON A.cod_jisa = C.cod_jisa and        ' +
                 ' A.dat_gmgn = C.dat_gmgn and B.num_jumin = C.num_jumin                          ' ;

      if GV_sJisa = '00' then
      begin
         sWhere := ' B.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' ';
      end
      else
      begin
         sWhere := ' B.cod_jisa = ''' + Copy(GV_sUserId, 1, 2) + ''' ';
      end;

      //연구소+센터만 조회
      sWhere := sWhere + ' and B.gubn_hulgum = ''3'' ';


      IF  (mskDate.Text < '20171011') and (Copy(cmbJisa.Text, 1, 2) <> '43')  THEN
      BEGIN
         {  //  곽윤설
         sWher e := sWhere + ' and substring(B.cod_jangbi,1,2) <> ''SS'' ';
         sWher e := sWhere +  ';
         sWher e := sWhere + ' and substring(B.cod_hul,1,2) <> ''SS''';
         sWher e := sWhere + ' and substring(B.cod_hul,1,2) <> ''GS''';
         }
         sWhere := sWhere +  '  and ((B.gubn_nosin = ''1'' or B.gubn_adult = ''1'' or B.gubn_life = ''1'') or (B.cod_chuga like ''%C077%'') or       ' +
                             '     ((B.gubn_nosin = ''2'' or B.gubn_adult = ''2'' or B.gubn_life = ''2'') and                                        ' +
                             '     ((B.cod_chuga like ''%C032%'') or (C.cod_chuga like ''%C032%'') or                                                ' +
                             '      (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C032'')) or                          ' +
                             '      (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C032'')))))                             ' ;
      END
      ELSE
      BEGIN
        sWhere := sWhere + '  and     ((B.gubn_nosin = ''1'' or B.gubn_adult = ''1'' or B.gubn_life = ''1'') or (B.cod_chuga like ''%C077%'') or                         ' +
                           '           (B.gubn_nosin = ''2'' or B.gubn_adult = ''2'' or B.gubn_life = ''2'') or (B.gubn_hulgum =''3'')                                   ' +
                           '  or       (B.gubn_tkgm  = ''1'' or B.gubn_tkgm = ''2'')                                                                                     ' +
                           '  or       (B.cod_chuga like ''%C009%'') OR  (B.cod_chuga like ''%C010%'') or (C.cod_chuga like ''%C011%'') or (C.cod_chuga like ''%C028%'') ' +
                           '  or       (B.cod_chuga like ''%C025%'')  OR  (B.cod_chuga like ''%C026%'') or (C.cod_chuga like ''%C027%'') or (C.cod_chuga like ''%C032%'')' +
                           '  or       (B.cod_chuga like ''%C042%'')  or  (B.cod_chuga like ''%C074%'') or (B.cod_chuga like ''%C033%'') or                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C009'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C009'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C010'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C010'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C011'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C011'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C025'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C025'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C026'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C026'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C027'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C027'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C028'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C028'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C074'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C074'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C032'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C032'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C033'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C033'')) or                                              ' +
                           '           (B.cod_jangbi in (select cod_pf from profile_hm with(nolock) where cod_hm=''C042'')) or                                           ' +
                           '           (B.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C042'')))                                                ' ;


      END;

      if rbDate.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND A.dat_bunju = ''' + mskDate.Text + ''' ';

         if Cmb_gubn.ItemIndex = 0 then
         begin
            if Trim(mskFrom.Text) <> '' then
            begin
                sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + ''' ';
            end;
            if Trim(mskTo.Text) <> '' then
            begin
                sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + ''' ';
            end;
         end
         else if Cmb_gubn.ItemIndex = 1 then
         begin
            if Trim(MEdt_SampS.Text) <> '' then
            begin
                sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + ''' ';
            end;
            if Trim(MEdt_SampE.Text) <> '' then
            begin
                sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + ''' ';
            end;
         end
         else
         begin
            if Trim(MEdt_SampS.Text) <> '' then
            begin
                sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + ''' ';
            end;
            if Trim(MEdt_SampE.Text) <> '' then
            begin
                sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + ''' ';
            end;
         end;

//------------------------------------------------------------------------------
            if (mskDate.Text = '20081025') then
            begin
               sWhere := sWhere + ' AND A.num_bunju <> ''1924''';
               if Copy(GV_sUserId, 1, 2) = '71' then
                  sWhere := sWhere + ' AND A.num_bunju <> ''7243''';
            end;
//------------------------------------------------------------------------------

         case RG_sort.ItemIndex of
            0 : sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
            1 : sOrderBy := ' ORDER BY A.num_bunju, A.gubn_part ';
            else
               sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
         end
      end
      else if rbJumin.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND B.num_jumin = ''' + mskJumin.Text + '''';
         sOrderBy := ' ORDER BY B.dat_gmgn, A.gubn_part ';
      end
      else if rbBarcode.Checked then
      begin
         sWhere := sWhere + ' AND B.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
         sWhere := sWhere + ' AND B.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';
         sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
      end;

      sWhere := sWhere + ' and A.gubn_part = ''01'' ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I10', 'Q', 'Y');
         UP_VarMemSet(RecordCount - 1);

         while not qryGlkwa.Eof do
         begin

         //2018.03.02일자 부터 순수채용인 생화학 자체 검사 -> 센터
         //순수채용 거르기
         chk_jisagum := false;
         if (mskDate.Text <= '20180301') and
            ((qryGlkwa.FieldByName('gubn_nosin').AsString <> '1') and (qryGlkwa.FieldByName('gubn_adult').AsString <> '1') and (qryGlkwa.FieldByName('gubn_life').AsString <> '1') and
             (qryGlkwa.FieldByName('gubn_nosin').AsString <> '2') and (qryGlkwa.FieldByName('gubn_adult').AsString <> '2') and (qryGlkwa.FieldByName('gubn_life').AsString <> '2')) then
         begin
            if (copy(qryGlkwa.FieldByName('cod_jangbi').AsString, 1, 2) = 'SS') or  (copy(qryGlkwa.FieldByName('cod_jangbi').AsString, 1, 2) = 'GS') or
               (copy(qryGlkwa.FieldByName('cod_hul').AsString, 1, 2) = 'SS')    or  (copy(qryGlkwa.FieldByName('cod_hul').AsString, 1, 2) = 'GS') then chk_jisagum := true;
         end;
         //자체검사 외 제외하기
         with qryHmCheck do
         begin
           qryHmCheck.Active := False;
           chk_jisagum2 := false;
           qryHmCheck.Sql.Clear;
           qryHmCheck.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryGlkwa.FieldByName('cod_jisa').AsString + ''','''
                                                                                  + qryGlkwa.FieldByName('num_id').AsString   + ''','''
                                                                                  + qryGlkwa.FieldByName('dat_gmgn').AsString + ''',''1'') ';
           qryHmCheck.Sql.Text := Sql.Text + ' WHERE Cod_hm IN (''C009'', ''C010'', ''C011'', ''C025'', ''C026'', ''C027'', ''C028'', ''C032'', ''C033'', ''C042'', ''C074'', ''C033'')';
           qryHmCheck.Active := True;
           if (qryHmCheck.RecordCount > 0) then chk_jisagum2 := true;
         end;

            if (chk_jisagum = false) and (chk_jisagum2 = true) then
            begin
            UV_vCod_jisa[index]   := FieldByName('COD_jisa').AsString;
            UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
            UV_vCod_bjjs[index]   := FieldByName('cod_bjjs').AsString;
            UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
            UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
            UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
            UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
            UV_vNum_samp[index]   := FieldByName('NUM_SAMP').AsString;

            UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
            UV_vCod_cancer[index] := FieldByName('Cod_cancer').AsString;
            UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
            UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;
            UV_vGubn_nosin[index] := FieldByName('Gubn_nosin').AsString;
            UV_vGubn_nsyh[index]  := FieldByName('Gubn_nsyh').AsString;
            UV_vGubn_adult[index] := FieldByName('Gubn_adult').AsString;
            UV_vGubn_adyh[index]  := FieldByName('Gubn_adyh').AsString;
            UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
            UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
            UV_vGubn_bogun[index] := FieldByName('Gubn_bogun').AsString;
            UV_vGubn_bgyh[index]  := FieldByName('Gubn_bgyh').AsString;
            UV_vGubn_agab[index]  := FieldByName('Gubn_agab').AsString;
            UV_vGubn_agyh[index]  := FieldByName('Gubn_agyh').AsString;
            UV_vGubn_tkgm[index]  := FieldByName('Gubn_tkgm').AsString;

            UV_vC009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  49,6));
            UV_vC010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  55,6));
            UV_vC011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  61,6));
            UV_vC025[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 121,6));

            UV_vC026[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 127,6));
            UV_vC027[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 133,6));
            UV_vC028[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 139,6));
            UV_vC042[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 187,6));

            UV_vC032[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 157,6));
            UV_vC033[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 163,6));
            UV_vC074[index]       := Trim(copy(FieldByName('desc_glkwa2').AsString, 131,6));
            UV_vC077[index]       := Trim(copy(FieldByName('desc_glkwa2').AsString, 149,6));  //200뺀값으로 위치값 넣음.

            UV_vPart01[index]     := FieldByName('desc_glkwa1').AsString;
            if FieldByName('gubn_preg').AsString ='Y' then
               UV_vGumgin_check[index]    := '임신중'
            else if FieldByName('gubn_preg').AsString ='P' then
                 UV_vGumgin_check[index]   := '임신가능성'
            else UV_vGumgin_check[index]   :=' ';
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

procedure TfrmLD1I10.btnCancelClick(Sender: TObject);
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

procedure TfrmLD1I10.UP_Change(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    eRslt, e25, e26, e28 , e74: Extended;
    cod_HM  : array[1..12] of String;
begin
   inherited;

   index := grdMaster.CurrentDataRow-1;

   cod_HM[1] := 'C009'; cod_HM[2] := 'C010'; cod_HM[3] := 'C011';
   cod_HM[4] := 'C025'; cod_HM[5] := 'C026'; cod_HM[6] := 'C027';
   cod_HM[7] := 'C028'; cod_HM[8] := 'C032'; cod_HM[9] := 'C042';
   cod_HM[10] := 'C074';  cod_HM[11] := 'C077';  cod_HM[12] := 'C033';

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

   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   with qryHangmok do
   begin
      close;
      ParamByName('dat_apply').AsString := mskDAT_GMGN.Text;
      ParamByName('cod_hm').AsString    := cod_HM[StrToInt(Copy(TEdit(Sender).Name, 9, 2))];
      open;
      ims_Low   := FieldByName('IMS_LOW'   + sGubn).AsFloat;
      ims_High  := FieldByName('IMS_HIGH'  + sGubn).AsFloat;
   end;

   if (Trim(TEdit(Sender).Text) <> '') and (TEdit(Sender).Color <> clBtnFace)then
   begin
      if qryHangmok.FieldByName('gubn_hm').AsString = '1' then
      begin
         if (GF_StrToFloat(Trim(TEdit(Sender).Text)) < ims_Low) or (GF_StrToFloat(Trim(TEdit(Sender).Text)) > ims_High) then
            TEdit(Sender).Color     := $00FACDF4
         else
            TEdit(Sender).Color     := clWindow;
      end
      else if qryHangmok.FieldByName('gubn_hm').AsString = '2' then
      begin
         if TEdit(Sender).text = qryHangmok.FieldByName('DESC_UNIT').AsString then
            TEdit(Sender).Color     := clWindow
         else
            TEdit(Sender).Color     := $00FACDF4;
      end;
   end;

   eRslt := 0;

   // 1. 생화학
   if (Sender = edtValue4) or (Sender = edtValue5) or (Sender = edtValue7) then
   begin
      Inc(iCntCheck);

      if (UV_vC074[index] = '') and (edtValue6.Enabled = true) then  //2017.11.13  항목코드 없으면 계산 안함.
      begin
        if (edtValue4.Text <> '') and (edtValue5.Text <> '') and (edtValue7.Text <> '') and (iCntCheck > 3) then
        begin
           e25 := GF_StrToFloat(edtValue4.Text);
           e26 := GF_StrToFloat(edtValue5.Text);
           e28 := GF_StrToFloat(edtValue7.Text);

           eRslt := e25 - e26 - (e28/5);
           eRslt := eRslt + 0.5;
           eRslt := trunc(eRslt);

           valChange.Scale := 1;
           valChange.Value := eRslt;
           sValue := valChange.Text;

           edtValue6.Text := sValue;
        end;
      end;
   end;

   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I10.UP_DChange(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    cod_HM  : array[1..12] of String;
begin
   inherited;

   cod_HM[1] := 'C009';  cod_HM[2] := 'C010';  cod_HM[3] := 'C011';
   cod_HM[4] := 'C025';  cod_HM[5] := 'C026';  cod_HM[6] := 'C027';
   cod_HM[7] := 'C028';  cod_HM[8] := 'C032';  cod_HM[9] := 'C042';
   cod_HM[10] := 'C074'; cod_HM[11] := 'C077'; cod_HM[12] := 'C033';

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

   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   with qryHangmok do
   begin
      close;
      ParamByName('dat_apply').AsString := mskDAT_GMGN.Text;
      ParamByName('cod_hm').AsString    := cod_HM[StrToInt(Copy(TEdit(Sender).Name, 9, 2))];
      open;
      ims_Low   := FieldByName('IMS_LOW'   + sGubn).AsFloat;
      ims_High  := FieldByName('IMS_HIGH'  + sGubn).AsFloat;
   end;

   if (Trim(TEdit(Sender).Text) <> '') and (TEdit(Sender).Color <> clBtnFace)then
   begin
      if qryHangmok.FieldByName('gubn_hm').AsString = '1' then
      begin
         if (GF_StrToFloat(Trim(TEdit(Sender).Text)) < ims_Low) or (GF_StrToFloat(Trim(TEdit(Sender).Text)) > ims_High) then
            TEdit(Sender).Color     := $00FACDF4
         else
            TEdit(Sender).Color     := clWindow;
      end
      else if qryHangmok.FieldByName('gubn_hm').AsString = '2' then
      begin
         if TEdit(Sender).text = qryHangmok.FieldByName('DESC_UNIT').AsString then
            TEdit(Sender).Color     := clWindow
         else
            TEdit(Sender).Color     := $00FACDF4;
      end;
   end;
end;

procedure TfrmLD1I10.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, iRet, i, iTemp : Integer;
    vCod_chuga : Variant;
    sHmCode, sHmList_01, sSelect, sWhere, sOrderby  : String;
begin
   inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   iCntCheck := 0;
   sHmList_01 := '';

   // Grid의 Row가 변경되면 Detail에 자료를 할당
   mskNUM_JUMIN.Text := UV_vNum_jumin[NewRow-1];
   edtDESC_NAME.Text := UV_vDesc_name[NewRow-1];
   mskDAT_GMGN.Text  := UV_vDat_gmgn[NewRow-1];

   Edt_bjjs.Text     := UV_vcod_bjjs[NewRow-1];
   mskDAT_BUNJU.Text := UV_vDat_bunju[NewRow-1];
   edtNUM_BUNJU.Text := UV_vNum_bunju[NewRow-1];
   Edt_smp.Text      := UV_vNum_samp[NewRow-1];
   Edt_ID.Text       := UV_vNum_id[NewRow-1];
   edt_gumgin_Check.Text   := UV_vGumgin_check[NewRow-1];
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
   qry_Check_CRM.Active := False;
   qry_Check_CRM.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qry_Check_CRM.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qry_Check_CRM.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qry_Check_CRM.Active := True;
   if qry_Check_CRM.FieldByName('ysno_crm').AsString = 'Y' then chk_CRM.checked := true
   else chk_CRM.checked := false;

   //항목찾기
   //---------------------------------------------------------------------------
   sSelect := '';
   with qryPf_hm do
   begin
      Active := False;

      sSelect := ' SELECT P.cod_hm FROM profile_hm P INNER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.gubn_part = ''01'' and H.dat_apply <= ''' + mskDAT_GMGN.Text + '''';
      sSelect := sSelect + ' WHERE P.cod_pf IN (''' + UV_vCod_hul[NewRow-1] + ''',''' + UV_vCod_cancer[NewRow-1] + ''',''' + UV_vCod_jangbi[NewRow-1] + ''')';
      sSelect := sSelect + ' GROUP BY P.cod_hm ';

      SQL.Clear;
      SQL.Add(sSelect);
      Active := True;

      if RecordCount > 0 then
      begin
         while not Eof do
         begin
            sHmList_01 := sHmList_01 + FieldByName('COD_HM').AsString;
            Next;
         end;
      end;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   iRet := GF_MulToSingle(UV_vCod_chuga[NewRow-1], 4, vCod_chuga);
   for i := 0 to iRet-1 do
   begin
      sHmList_01 := sHmList_01 + vCod_chuga[i];
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
                  sHmCode := sHmCode + qryProfileG.FieldByName('cod_hm').AsString;
                  qryProfileG.Next;
               end;
            end;
            qryProfileG.Active := False;
         end;
//------------------------------------------------------------------------------
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;

      end;
      qryTkgum.Active := False;
   end;

//특검항목..//20150423
//------------------------------------------------------------------------------
   if (UV_vGubn_tkgm[NewRow-1] = '2')  then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
      qryTkgum.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryTkgum.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         if pos('c009', qryTkgum.FieldByName('cod_chuga').AsString) > 0 then
         begin
            sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
         end;
         if pos('C074', qryTkgum.FieldByName('cod_chuga').AsString) > 0 then
         begin
            sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
         end;
      end;
      qryTkgum.Active := False;
   end;


   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         sHmList_01 := sHmList_01 + vCod_chuga[i];
      end;
   end;

   //--------------------------------------------------------------------------
  {if ((UV_vGubn_nosin[NewRow-1] = '2') or (UV_vGubn_adult[NewRow-1] = '2') or (UV_vGubn_life[NewRow-1] = '2'))
       and (pos('C032', sHmList_01) > 0) and (Copy(GV_sUserId, 1, 2) <> '43') then
   begin
      edtValue1.Color := clBtnFace;
      edtValue2.Color := clBtnFace;
      edtValue3.Color := clBtnFace;
      edtValue4.Color := clBtnFace;
      edtValue5.Color := clBtnFace;
      edtValue6.Color := clBtnFace;
      edtValue7.Color := clBtnFace;
      edtValue9.Color := clBtnFace;
      edtValue10.Color := clBtnFace;
      edtValue11.Color := clBtnFace;

      edtValue1.Enabled := False;
      edtValue2.Enabled := False;
      edtValue3.Enabled := False;
      edtValue4.Enabled := False;
      edtValue5.Enabled := False;
      edtValue6.Enabled := False;
      edtValue7.Enabled := False;
      edtValue9.Enabled := False;
      edtValue10.Enabled := False;
      edtValue11.Enabled := False;
      if pos('C032',sHmList_01) > 0 then
      begin
         edtValue8.Color := clWindow;
         edtValue8.Enabled := True;
      end
      else
      begin
         edtValue8.Color := clBtnFace;
         edtValue8.Enabled := False;
      end;
   end}

    //2018.03.02 순수채용 생화학 자체검사 -> 센터 -------------------------------------------------------------------
    if  (mskDate.text <= '20180301') then
    begin
       if (((Copy(UV_vCod_jangbi[NewRow-1],1,2) = 'SS') or (Copy(UV_vCod_jangbi[NewRow-1],1,2) = 'GS')) or
           ((Copy(UV_vCod_hul[NewRow-1],1,2) = 'SS') or (Copy(UV_vCod_hul[NewRow-1],1,2) = 'GS'))) Then
        begin
        if (UV_vGubn_nosin[NewRow-1] = '1') or (UV_vGubn_life[NewRow-1]  =  '1') or (UV_vGubn_adult[NewRow-1] = '1') or
           (UV_vGubn_nosin[NewRow-1] = '2') or (UV_vGubn_life[NewRow-1]  =  '2') or (UV_vGubn_adult[NewRow-1] = '2') then
          begin
            if pos('C009',sHmList_01) > 0 then edtValue1.Color := clWindow
            else                               edtValue1.Color := clBtnFace;
            if pos('C010',sHmList_01) > 0 then edtValue2.Color := clWindow
            else                               edtValue2.Color := clBtnFace;
            if pos('C011',sHmList_01) > 0 then edtValue3.Color := clWindow
            else                               edtValue3.Color := clBtnFace;
            if pos('C025',sHmList_01) > 0 then edtValue4.Color := clWindow
            else                               edtValue4.Color := clBtnFace;
            if pos('C026',sHmList_01) > 0 then edtValue5.Color := clWindow
            else                               edtValue5.Color := clBtnFace;
            if pos('C027',sHmList_01) > 0 then edtValue6.Color := clWindow
            else                               edtValue6.Color := clBtnFace;
            if pos('C028',sHmList_01) > 0 then edtValue7.Color := clWindow
            else                               edtValue7.Color := clBtnFace;
            if pos('C032',sHmList_01) > 0 then edtValue8.Color := clWindow
            else                               edtValue8.Color := clBtnFace;
            if pos('C042',sHmList_01) > 0 then edtValue9.Color := clWindow
            else                               edtValue9.Color := clBtnFace;
            if pos('C074',sHmList_01) > 0 then edtValue10.Color:= clWindow
            else                               edtValue10.Color:= clBtnFace;
            if pos('C077',sHmList_01) > 0 then edtValue11.Color:= clWindow
            else                               edtValue11.Color:= clBtnFace;
            if pos('C033',sHmList_01) > 0 then edtValue12.Color := clWindow
            else                               edtValue12.Color := clBtnFace;
            if pos('C009',sHmList_01) > 0 then edtValue1.Enabled := True
            else                               edtValue1.Enabled := False;
            if pos('C010',sHmList_01) > 0 then edtValue2.Enabled := True
            else                               edtValue2.Enabled := False;
            if pos('C011',sHmList_01) > 0 then edtValue3.Enabled := True
            else                               edtValue3.Enabled := False;
            if pos('C025',sHmList_01) > 0 then edtValue4.Enabled := True
            else                               edtValue4.Enabled := False;
            if pos('C026',sHmList_01) > 0 then edtValue5.Enabled := True
            else                               edtValue5.Enabled := False;
            if pos('C027',sHmList_01) > 0 then edtValue6.Enabled := True
            else                               edtValue6.Enabled := False;
            if pos('C028',sHmList_01) > 0 then edtValue7.Enabled := True
            else                               edtValue7.Enabled := False;
            if pos('C032',sHmList_01) > 0 then edtValue8.Enabled := True
            else                               edtValue8.Enabled := False;
            if pos('C042',sHmList_01) > 0 then edtValue9.Enabled := True
            else                               edtValue9.Enabled := False;
            if pos('C074',sHmList_01)> 0 then  edtValue10.Enabled := True
            else                               edtValue10.Enabled := False;
            if pos('C077',sHmList_01) > 0 then edtValue11.Enabled := True
            else                               edtValue11.Enabled := False;
            if pos('C033',sHmList_01) > 0 then edtValue12.Enabled := True
            else                               edtValue12.Enabled := False;
         end
         else
         begin
            edtValue1.Color := clBtnFace;
            edtValue2.Color := clBtnFace;
            edtValue3.Color := clBtnFace;
            edtValue4.Color := clBtnFace;
            edtValue5.Color := clBtnFace;
            edtValue6.Color := clBtnFace;
            edtValue7.Color := clBtnFace;
            edtValue8.Color := clBtnFace;
            edtValue9.Color := clBtnFace;
            edtValue10.Color := clBtnFace;
            edtValue11.Color := clBtnFace;
            edtValue12.Color := clBtnFace;

            edtValue1.Enabled := False;
            edtValue2.Enabled := False;
            edtValue3.Enabled := False;
            edtValue4.Enabled := False;
            edtValue5.Enabled := False;
            edtValue6.Enabled := False;
            edtValue7.Enabled := False;
            edtValue8.Enabled := False;
            edtValue9.Enabled := False;
            edtValue10.Enabled := False;
            edtValue11.Enabled := False;
            edtValue12.Enabled := False;
         end;
       end
    end
    else
    begin
        if pos('C009',sHmList_01) > 0 then edtValue1.Color := clWindow
        else                               edtValue1.Color := clBtnFace;
        if pos('C010',sHmList_01) > 0 then edtValue2.Color := clWindow
        else                               edtValue2.Color := clBtnFace;
        if pos('C011',sHmList_01) > 0 then edtValue3.Color := clWindow
        else                               edtValue3.Color := clBtnFace;
        if pos('C025',sHmList_01) > 0 then edtValue4.Color := clWindow
        else                               edtValue4.Color := clBtnFace;
        if pos('C026',sHmList_01) > 0 then edtValue5.Color := clWindow
        else                               edtValue5.Color := clBtnFace;
        if pos('C027',sHmList_01) > 0 then edtValue6.Color := clWindow
        else                               edtValue6.Color := clBtnFace;
        if pos('C028',sHmList_01) > 0 then edtValue7.Color := clWindow
        else                               edtValue7.Color := clBtnFace;
        if pos('C032',sHmList_01) > 0 then edtValue8.Color := clWindow
        else                               edtValue8.Color := clBtnFace;
        if pos('C042',sHmList_01) > 0 then edtValue9.Color := clWindow
        else                               edtValue9.Color := clBtnFace;
        if pos('C074',sHmList_01) > 0 then edtValue10.Color := clWindow
        else                               edtValue10.Color := clBtnFace;
        if pos('C077',sHmList_01) > 0 then edtValue11.Color := clWindow
        else                               edtValue11.Color := clBtnFace;
        if pos('C033',sHmList_01) > 0 then edtValue12.Color := clWindow
        else                               edtValue12.Color := clBtnFace;

        if pos('C009',sHmList_01) > 0 then edtValue1.Enabled := True
        else                               edtValue1.Enabled := False;
        if pos('C010',sHmList_01) > 0 then edtValue2.Enabled := True
        else                               edtValue2.Enabled := False;
        if pos('C011',sHmList_01) > 0 then edtValue3.Enabled := True
        else                               edtValue3.Enabled := False;
        if pos('C025',sHmList_01) > 0 then edtValue4.Enabled := True
        else                               edtValue4.Enabled := False;
        if pos('C026',sHmList_01) > 0 then edtValue5.Enabled := True
        else                               edtValue5.Enabled := False;
        if pos('C027',sHmList_01) > 0 then edtValue6.Enabled := True
        else                               edtValue6.Enabled := False;
        if pos('C028',sHmList_01) > 0 then edtValue7.Enabled := True
        else                               edtValue7.Enabled := False;
        if pos('C032',sHmList_01) > 0 then edtValue8.Enabled := True
        else                               edtValue8.Enabled := False;
        if pos('C042',sHmList_01) > 0 then edtValue9.Enabled := True
        else                               edtValue9.Enabled := False;
        if pos('C074',sHmList_01)> 0  then edtValue10.Enabled := True
        else                               edtValue10.Enabled := False;
        if pos('C077',sHmList_01) > 0 then edtValue11.Enabled := True
        else                               edtValue11.Enabled := False;
        if pos('C033',sHmList_01) > 0 then edtValue12.Enabled := True
        else                               edtValue12.Enabled := False;
   end;

   edtValue1.Text    := UV_vC009[NewRow-1];
   edtValue2.Text    := UV_vC010[NewRow-1];
   edtValue3.Text    := UV_vC011[NewRow-1];
   edtValue4.Text    := UV_vC025[NewRow-1];
   edtValue5.Text    := UV_vC026[NewRow-1];
   edtValue6.Text    := UV_vC027[NewRow-1];
   edtValue7.Text    := UV_vC028[NewRow-1];
   edtValue8.Text    := UV_vC032[NewRow-1];
   edtValue9.Text    := UV_vC042[NewRow-1];
   edtValue10.Text   := UV_vC074[NewRow-1];
   edtValue11.Text   := UV_vC077[NewRow-1];
   edtValue12.Text   := UV_vC033[NewRow-1];

   edtPValue1.Text  := '';
   edtPValue2.Text  := '';
   edtPValue3.Text  := '';
   edtPValue4.Text  := '';
   edtPValue5.Text  := '';
   edtPValue6.Text  := '';
   edtPValue7.Text  := '';
   edtPValue8.Text  := '';
   edtPValue9.Text  := '';
   edtPValue10.Text := '';
   edtPValue11.Text := '';
   edtPValue12.Text := '';


   mskPDAT_BUNJU.Text := '';
   edtPNUM_BUNJU.Text := '';
//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
   case RGrp_preGulkwa.ItemIndex of
      0 : begin
             qryinjouk.Active := False;
             qryinjouk.ParamByName('NUM_JUMIN').AsString := UV_vNum_jumin[NewRow-1];
             qryinjouk.ParamByName('DAT_GMGN').AsString  := UV_vDat_gmgn[NewRow-1];
             qryinjouk.Active := True;
             if qryinjouk.RecordCount > 0 then
             begin
                qryPrev.Active := False;
                qryPrev.Filter := '';
                qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                qryPrev.ParamByName('GUBN_PART').AsString := '01';
                qryPrev.Active := True;
                if qryPrev.RecordCount > 0 then
                begin
                   mskPDAT_BUNJU.Text := qryPrev.FieldByName('dat_bunju').AsString;
                   edtPNUM_BUNJU.Text := qryPrev.FieldByName('num_bunju').AsString;

                   UV_fGulkwa := '';
                   UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                   UV_vDesc_Pglkwa[index] := UV_fGulkwa;

                   //SGOT(AST)(C009)
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
                      edtPValue1.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else
                      edtPValue1.Text := '';
                   //SGPT(ALT)(C010)
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
                      edtPValue2.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else
                      edtPValue2.Text := '';
                   //r-GPT(C011)
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
                      edtPValue3.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else
                      edtPValue3.Text := '';
                   //콜레스테롤(C025)
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
                      edtPValue4.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else
                      edtPValue4.Text := '';
                   //HDL-콜레스테롤(C026)
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
                      edtPValue5.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else
                      edtPValue5.Text := '';
                   //LDL-콜레스테롤(C027)
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
                      edtPValue6.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else
                      edtPValue6.Text := '';
                   //중성지방(C028)
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
                      edtPValue7.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else
                      edtPValue7.Text := '';
                   //공복시혈당(C032)
                   if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
                      edtPValue8.Text := Trim(Copy(UV_fGulkwa, 157,  6))
                   else
                      edtPValue8.Text := '';
                   //크레아티닌(C042)
                   if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
                      edtPValue9.Text := Trim(Copy(UV_fGulkwa, 187,  6))
                   else
                      edtPValue9.Text := '';
                   //LDL-콜레스테롤(효소)(C074)
                   if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 331,  6))
                   else                                          edtPValue10.Text := '';
                   //식전 자가혈당측정기(C077)
                   if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 349,  6))
                   else                                          edtPValue11.Text := '';
                   //식후 2시간 후 혈당(C033)
                   if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then edtPValue12.Text := Trim(Copy(UV_fGulkwa, 163,  6))
                   else                                          edtPValue12.Text := '';
                end; // end of if[qryPrev];
                qryPrev.Active := False;
             end; // end of if[qryinjouk]
             qryinjouk.Active := False;
          end; // end of[0]
      1 : begin
             qryinjouk1.Active := False;

             sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                        ' FROM gumgin G INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
             sWhere  := ' WHERE G.num_jumin = ''' + UV_vNum_jumin[NewRow-1] + '''';
             sWhere  := sWhere + ' AND G.dat_gmgn < ''' + UV_vDat_gmgn[NewRow-1] + '''';
             sWhere  := sWhere + ' AND B.gubn_part = ''02''';
             sOrderBy := ' order by G.dat_gmgn DESC ';

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
                qryPrev.ParamByName('GUBN_PART').AsString := '02';
                qryPrev.Active := True;
                if qryPrev.RecordCount > 0 then
                begin
                   mskPDAT_BUNJU.Text := qryPrev.FieldByName('dat_bunju').AsString;
                   edtPNUM_BUNJU.Text := qryPrev.FieldByName('num_bunju').AsString;

                   UV_fGulkwa := '';
                   UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                   UV_vDesc_Pglkwa[index] := UV_fGulkwa;

                   //SGOT(AST)(C009)
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
                      edtPValue1.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else
                      edtPValue1.Text := '';
                   //SGPT(ALT)(C010)
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
                      edtPValue2.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else
                      edtPValue2.Text := '';
                   //r-GPT(C011)
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
                      edtPValue3.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else
                      edtPValue3.Text := '';
                   //콜레스테롤(C025)
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
                      edtPValue4.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else
                      edtPValue4.Text := '';
                   //HDL-콜레스테롤(C026)
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
                      edtPValue5.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else
                      edtPValue5.Text := '';
                   //LDL-콜레스테롤(C027)
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
                      edtPValue6.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else
                      edtPValue6.Text := '';
                   //중성지방(C028)
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
                      edtPValue7.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else
                      edtPValue7.Text := '';
                   //공복시혈당(C032)
                   if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
                      edtPValue8.Text := Trim(Copy(UV_fGulkwa, 157,  6))
                   else
                      edtPValue8.Text := '';
                   //크레아티닌(C042)
                   if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
                      edtPValue9.Text := Trim(Copy(UV_fGulkwa, 187,  6))
                   else
                      edtPValue9.Text := '';
                   //LDL-콜레스테롤(효소)(C074)
                   if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 331,  6))
                   else                                          edtPValue10.Text := '';
                   //식전 자가혈당측정기(C077)
                   if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 349,  6))
                   else                                          edtPValue11.Text := '';
                   //공복시혈당(C033)
                   if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then
                      edtPValue12.Text := Trim(Copy(UV_fGulkwa, 163,  6))
                   else
                      edtPValue12.Text := '';
                end; // end of if[qryPrev];
                qryPrev.Active := False;
             end; // end of if[qryinjouk]
             qryinjouk.Active := False;
          end; // end of[1];
   end; // end of case
   with qryinjouk do
   begin

      qryinjouk.Active := False;
      qryinjouk.ParamByName('num_Jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryinjouk.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryinjouk.Active := True;
      if RecordCount > 0 then
      begin
         qryPrev.Active := False;
         qryPrev.Filter := '';
         qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
         qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
         qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
         qryPrev.ParamByName('GUBN_PART').AsString := '01';
         qryPrev.Active := True;
         if qryPrev.RecordCount > 0 then
         begin
            mskPDAT_BUNJU.Text := qryPrev.FieldByName('dat_bunju').AsString;
            edtPNUM_BUNJU.Text := qryPrev.FieldByName('num_bunju').AsString;

            UV_fGulkwa := '';
            UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

            UV_vDesc_Pglkwa[index] := UV_fGulkwa;

            //SGOT(AST)(C009)
            if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
               edtPValue1.Text := Trim(Copy(UV_fGulkwa, 49,  6))
            else
               edtPValue1.Text := '';
            //SGPT(ALT)(C010)
            if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
               edtPValue2.Text := Trim(Copy(UV_fGulkwa, 55,  6))
            else
               edtPValue2.Text := '';
            //r-GPT(C011)
            if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
               edtPValue3.Text := Trim(Copy(UV_fGulkwa, 61,  6))
            else
               edtPValue3.Text := '';
            //콜레스테롤(C025)
            if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
               edtPValue4.Text := Trim(Copy(UV_fGulkwa, 121,  6))
            else
               edtPValue4.Text := '';
            //HDL-콜레스테롤(C026)
            if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
               edtPValue5.Text := Trim(Copy(UV_fGulkwa, 127,  6))
            else
               edtPValue5.Text := '';
            //LDL-콜레스테롤(C027)
            if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
               edtPValue6.Text := Trim(Copy(UV_fGulkwa, 133,  6))
            else
               edtPValue6.Text := '';
            //중성지방(C028)
            if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
               edtPValue7.Text := Trim(Copy(UV_fGulkwa, 139,  6))
            else
               edtPValue7.Text := '';
            //공복시혈당(C032)
            if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
               edtPValue8.Text := Trim(Copy(UV_fGulkwa, 157,  6))
            else
               edtPValue8.Text := '';
            //크레아티닌(C042)
            if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
               edtPValue9.Text := Trim(Copy(UV_fGulkwa, 187,  6))
            else
               edtPValue9.Text := '';
            //LDL-콜레스테롤(효소)(C074)
            if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 331,  6))
            else                                          edtPValue10.Text := '';
            //식전 자가혈당측정기(C077)
            if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 349,  6))
            else                                          edtPValue11.Text := '';
            //식후두시간후혈당(C033)
            if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then
               edtPValue12.Text := Trim(Copy(UV_fGulkwa, 163,  6))
            else
               edtPValue12.Text := '';
         end;
         qryPrev.Active := False;
      end
      else
      begin

      end;
      qryinjouk.Active := False;
   end;

//[2012.003.16]검체불량관리 추가..
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

   // 값변형 체크
   UP_DChange(edtValue1);
   UP_DChange(edtValue2);
   UP_DChange(edtValue3);
   UP_DChange(edtValue4);
   UP_DChange(edtValue5);
   UP_DChange(edtValue6);
   UP_DChange(edtValue7);
   UP_DChange(edtValue8);
   UP_DChange(edtValue9);
   UP_DChange(edtValue10);
   UP_DChange(edtValue11);
   UP_DChange(edtValue12);

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD1I10.btnPInsertClick(Sender: TObject);
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

procedure TfrmLD1I10.UP_Click(Sender: TObject);
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
   else if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD1I10.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I10.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD1I10.UP_Exit(Sender: TObject);
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

procedure TfrmLD1I10.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      mskDate.Color    := $00E6FFFE;
      mskDate.Enabled  := True;
      btnDate.Enabled  := True;
      mskFrom.Enabled  := True;
      mskTo.Enabled    := True;
      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;
      btnJumin.Enabled := False;
      mskDate.SetFocus;

      // 찾기 Enable
      //edtFind.Enabled  := True;
      edtFind2.Enabled  := True;
   end
   else if Sender = rbJumin then
   begin
      mskDate.Color    := clWindow;
      mskDate.Enabled  := False;
      btnDate.Enabled  := False;
      mskFrom.Enabled  := False;
      mskTo.Enabled    := False;
      mskJumin.Color   := $00E6FFFE;
      mskJumin.Enabled := True;
      btnJumin.Enabled := True;
      mskJumin.SetFocus;

      // 찾기 Disable
      //edtFind.Enabled  := False;
      edtFind2.Enabled := False;
   end;
end;

procedure TfrmLD1I10.edtFindExit(Sender: TObject);
begin
   inherited;

   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0);
end;

procedure TfrmLD1I10.EdtFind2Exit(Sender: TObject);
begin
  inherited;
    // 자료가 존재하는지 Check
   if Length(edtFind2.Text) < edtFind2.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind2.Text, 2, 1, 1, 0);
end;

procedure TfrmLD1I10.mskNumExit(Sender: TObject);
var i, iTemp : Integer;
    sName : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if mskNum.Text = '' then exit;

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

procedure TfrmLD1I10.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F7 then
   begin
      if edtFind2.Enabled then edtFind2.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);

   if Key = VK_F3 then
   begin
      if edtFind.Enabled then edtFind.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I10.UP_EditDisplay(Sender : TObject);
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

   UP_Change(Sender);
end;

procedure TfrmLD1I10.edtValue1Change(Sender: TObject);
begin
  inherited;
   //해당하는 값을 넣어줌.
   if Copy(TEdit(Sender).Name, 1, 8)  = 'edtValue' then
      UP_EditDisplay(Sender);
end;

procedure TfrmLD1I10.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD1P101 := TfrmLD1P101.Create(Self);
  frmLD1P101.QuickRep.Preview;

  //   frmLD1I101.QuickRep.Print;
end;


end.
