//==============================================================================
// 프로그램 설명 : 혈액소견관리
// 시스템        : 통합검진
// 수정일자      : 2005.01.26
// 수정자        : 유동구
// 수정내용      : 분주실 인증 수정..(조회부분)..
// 참고사항      :
//==============================================================================
// 수정일자      : 2006.04.24
// 수정자        : 유동구
// 수정내용      : 종전결과 비교일자를 혈액종검일때로 수정..
//==============================================================================
// 수정일자      : 2006.06.01
// 수정자        : 유동구
// 수정내용      : 혈액소견 10개로 늘림..
//==============================================================================
// 수정일자      : 2006.09.08
// 수정자        : 유동구
// 수정내용      : 화면구성 변경..(혈액코드,종양코드 추가)
//==============================================================================
// 수정일자      : 2006.05.08
// 수정자        : 김승철
// 수정내용      : 생애전환기 추가
//==============================================================================
// 수정일자      : 2008.03.28
// 수정자        : 김승철
// 수정내용      : 김소영선생님 소견저장시 '저장하시겠습니까' 멘트 안나오도록
//==============================================================================
// 수정일자      : 2009.06.15
// 수정자        : 송재호
// 수정내용      : 중요항목(C026,C027,S014,S015,S018,S019)포함 표시
//==============================================================================
// 수정일자      : 2009.06.18
// 수정자        : 김승철
// 수정내용      : 소견12개로 늘림
//==============================================================================
// 수정일자      : 2009.08.14
// 수정자        : 송재호
// 수정내용      : 중요항목에 SE01, SE02 추가
//==============================================================================
// 수정일자      : 2011.07.13
// 수정자        : 김승철
// 수정내용      : 항원, 항체에 S007,S008만 되는거 SE07, SE08 도 되도록.(sTot_hangmok:전체항목코드)추가.
//==============================================================================
// 수정일자      : 2012.03.22
// 수정자        : 송재호
// 수정내용      : H018~H022는 폐기항목이므로 이상자료 표시 안되도록 수정(김소영 선생님&김미래 차장님 요청)
//==============================================================================
// 수정일자      : 2012.6.22
// 수정자        : 김희석
// 수정내용      : 불량사유 
//==============================================================================
// 수정일자      : 2012.9.17
// 수정자        : 김희석
// 수정내용      : 중요항목 E003 추가 
//==============================================================================
// 수정일자      : 2012.11.14
// 수정자        : 김희석
// 수정내용      : 출력및 엑셀변환 추가 
//==============================================================================
// 수정일자      : 2012.12.11
// 수정자        : 김희석
// 수정내용      : 임신,센터,추가 .   
//==============================================================================
// 수정일자      : 2013.4.10
// 수정자        : 김희석
// 수정내용      : 18세이하 조회 조건 추가
//==============================================================================
// 수정일자      : 2013.6.28
// 수정자        : 김희석
// 수정내용      : 바코드 조회 추가, 샘플번호추가
//==============================================================================
// 수정일자      : 2014.10.20
// 수정자        : 곽윤설
// 수정내용      : 중요항목 포함 표시 : SE31, T043, C083, Y004, U017 추가
// 요청자        : 본원 - 김소영 전문의
//==============================================================================
// 수정일자      : 2015.05.07
// 수정자        : 곽윤설
// 수정내용      : 뇨침사(U011, U012, U013) 임상참고치&판정 안나오는 부분 수정
// 요청자        : 한의 재단윤리경영팀장1500026 [본원 - 이춘희 수석, 박연숙 수석 요청]
//==============================================================================
// 수정일자      : 2015.10.23
// 수정자        : 곽윤설
// 수정내용      : UV_vCod_hm 초기화 미적용으로 인한 에러 정정
// 요청자        :
//==============================================================================
// 수정일자      : 2016.11.09
// 수정자        : 박수지
// 수정내용      : 대전지사 세부소견 입력시 저장이 가능하도록 수정
// 요청자        :
//==============================================================================

unit LD6I02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, Gauges, ComObj;

type
  TfrmLD6I02 = class(TfrmSingle)
    qryBunju: TQuery;
    grdMaster: TtsGrid;
    qryGlkwa: TQuery;
    pnlDetail: TPanel;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    mskDAT_BUNJU: TDateEdit;
    edtNUM_BUNJU: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    mskNUM_JUMIN: TMaskEdit;
    Label6: TLabel;
    edtDESC_NAME: TEdit;
    qryHm: TQuery;
    qryPf_hm: TQuery;
    qryGmgn: TQuery;
    edtFind: TEdit;
    Label2: TLabel;
    pnlCond: TPanel;
    btnDate: TSpeedButton;
    btnJumin: TSpeedButton;
    Label9: TLabel;
    Label10: TLabel;
    rbDate: TRadioButton;
    rbJumin: TRadioButton;
    mskDate: TDateEdit;
    mskJumin: TMaskEdit;
    edtName: TEdit;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    grdResult: TtsGrid;
    pgcSokun: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    grdSokun1: TtsGrid;
    edtMan: TEdit;
    edtAge: TEdit;
    qryHsok: TQuery;
    qryI_Hsok: TQuery;
    qryU_Hsok: TQuery;
    qryD_Hsok: TQuery;
    qryU_Bnju: TQuery;
    btnPrevNum: TBitBtn;
    btnNextNum: TBitBtn;
    Label7: TLabel;
    edtBIMAN: TEdit;
    Label11: TLabel;
    edtH023: TEdit;
    Label12: TLabel;
    edtCOD_PF: TEdit;
    Label13: TLabel;
    edtHUL: TEdit;
    Label16: TLabel;
    edtS007: TEdit;
    Label17: TLabel;
    edtS008: TEdit;
    Label18: TLabel;
    edtS034: TEdit;
    Label19: TLabel;
    edtS020: TEdit;
    Label20: TLabel;
    edtS003: TEdit;
    Label21: TLabel;
    edtTT01: TEdit;
    Label22: TLabel;
    edtS001: TEdit;
    ckCheck: TCheckBox;
    qryPrev2006: TQuery;
    Label8: TLabel;
    mskDAT_GMGN: TDateEdit;
    edtH024: TEdit;
    qryGum_bul: TQuery;
    Label23: TLabel;
    grdSokun2: TtsGrid;
    grdSokun3: TtsGrid;
    qrySk_h: TQuery;
    Label59: TLabel;
    cmbDoct: TComboBox;
    qryU_gumgin: TQuery;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    qryInjouk2006: TQuery;
    Label24: TLabel;
    mskPDAT_GMGN: TDateEdit;
    qryPrev: TQuery;
    qryinjouk: TQuery;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    EdtCOD_HUL: TEdit;
    EdtCOD_CANCER: TEdit;
    Bevel2: TBevel;
    Bevel3: TBevel;
    ck6Check: TCheckBox;
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
    Panel4: TPanel;
    JD_sayu: TEdit;
    qry_Gum_bul: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Label14: TLabel;
    edt_preg: TEdit;
    Bevel4: TBevel;
    Label15: TLabel;
    edt_jisa: TEdit;
    Chk_Age: TCheckBox;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
    Label29: TLabel;
    edtNum_Samp: TEdit;
    Cmb_gubn: TComboBox;
    MEdt_SampE: TMaskEdit;
    Label30: TLabel;
    MEdt_SampS: TMaskEdit;
    Label31: TLabel;
    qryHM_sp: TQuery;
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
    procedure grdResultCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdSokun1CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdSokun1CellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure grdSokun1KeyPress(Sender: TObject; var Key: Char);
    procedure ckCheckClick(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure UP_MoveNum(Sender: TObject);
    procedure grdSokun2CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdSokun3CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdSokun2CellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure grdSokun3CellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs, UV_vNum_id, UV_vCod_jisa, UV_vDat_gmgn, UV_vNum_jumin, UV_vDesc_name,
    UV_vNum_bunju, UV_vDat_bunju, UV_vCod_hul, UV_vCod_cancer, UV_vCod_jangbi, UV_vCod_chuga,
    UV_vBiman, UV_vHul_h, UV_vHul_l, UV_vCod_nosin, UV_vCod_adult, UV_vCod_agab, UV_vCod_nsyh,
    UV_vCod_adyh, UV_vCod_agyh, UV_vCod_tkgm, UV_vCod_life, UV_vCod_lfyh,UV_vPreg,UV_vNum_Samp : Variant;

    // 검사결과 Grid
    UV_vSeq, UV_vDesc_hm, UV_vIms_low, UV_vIms_high, UV_vPrev_rslt, UV_vCur_rslt,
    UV_vMin3, UV_vMin2, UV_vMin1, UV_vJung, UV_vMax1, UV_vMax2, UV_vMax3, UV_vDesc_unit : Variant;

    // 혈액소견, 식이요법, 운동요법 Grid
    UV_vCod_sokun1, UV_vDesc_sokun1, UV_vPrev_cod1, UV_vInfo_numb1, UV_vCod_Doct, UV_vGubn_pan : Variant;
    UV_vCod_sokun2, UV_vDesc_sokun2, UV_vPrev_cod2, UV_vInfo_numb2 : Variant;
    UV_vCod_sokun3, UV_vDesc_sokun3, UV_vPrev_cod3, UV_vInfo_numb3 : Variant;

    // 검진자가 검사한 항목코드
    UV_vCod_hm : Variant;

    iCmb_Index : Integer;

    //종전결과가 종검일때 'Y' 체크...
    UV_sBunjuchk : String;
    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_ResultVarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
    procedure UP_Separate(sData : String; var sValue : String);
    procedure UP_SokunQuery;
    procedure UP_ResultQuery;
    function  UF_HangmokResult(qryResult : TQuery; sCod_hm : String) : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

    procedure MouseWheelHandler(var Message: TMessage); override;    
  public
    { Public declarations }
  end;

var
  frmLD6I02: TfrmLD6I02;
  UV_sHul_Part : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD6I021;

{$R *.DFM}

procedure TfrmLD6I02.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
//         grdResult.Refresh; //추가
      end;
   end;
end;

function  TfrmLD6I02.UF_HangmokResult(qryResult : TQuery; sCod_hm : String) : String;
var iCur : Integer;
    sResult : String;
begin
   iCur := grdMaster.CurrentDataRow - 1;

   // 항목코드에 대한 결과위치를 추출
   qryHm.Filter := 'COD_HM = ''' + sCod_hm + ''' AND ' +
                   'DAT_APPLY <= ''' + UV_vDat_gmgn[iCur] + '''';

   // 해당 검사에 대한 결과를 추출(Start Pos, End Pos를 이용)
   qryResult.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';

   if qryResult.RecordCount > 0 then
   begin
      UV_fGulkwa := '';
      UV_fGulkwa1 := qryResult.FieldByName('DESC_GLKWA1').AsString;
      UV_fGulkwa2 := qryResult.FieldByName('DESC_GLKWA2').AsString;
      UV_fGulkwa3 := qryResult.FieldByName('DESC_GLKWA3').AsString;
      GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
      if (qryHm.FieldByName('POS_START').AsInteger <> 0) and (qryHm.FieldByName('POS_END').AsInteger <> 0) then
      begin
           sResult := Copy(UV_fGulkwa,
                      qryHm.FieldByName('POS_START').AsInteger,
                      qryHm.FieldByName('POS_END').AsInteger -
                      qryHm.FieldByName('POS_START').AsInteger + 1);
      end
      else sResult := '';
   end;
   Result := sResult;
end;

procedure TfrmLD6I02.UP_ResultQuery;
var sMan, sGubn : String;
    iAge : Integer;
    sResult, sTemp : String;
    eResult : Extended;
    i, iResult_cnt, iCur, iCut : Integer;
    bCheck : Boolean;

    // 임상참고치
    eMin3, eMin2, eMin1, eLow, eHigh, eMax1, eMax2, eMax3, eIms : Extended;
    sMunja : String;
begin
   // 현재위치 추출
   iCur := grdMaster.CurrentDataRow - 1;

   // 주민번호를 이용하여 성별과 나이를 구함
   GP_JuminToAge(mskNUM_JUMIN.Text,mskDAT_GMGN.Text, sMan, iAge);
   edtMan.Text := sMan;
   edtAge.Text := IntToStr(iAge);

   // 성별과 나이를 기준으로 Column명을 추출
   sGubn := '';
   if iAge < 10 then sGubn := '1'
   else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
   else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
   else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
   else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
   else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
   else sGubn := '7';

   if edtMan.Text = 'F' then sGubn := sGubn + 'f';

   // 초기값 설정
   iResult_cnt := 0;
   grdResult.Rows := 0;
   iCut := 0;

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

      // 검사항목에 대한 임상참고치를 획득
      qryHm.Filter := 'COD_HM = ''' + UV_vCod_hm[i] + ''' AND ' +
                      'DAT_APPLY <= ''' + UV_vDat_gmgn[iCur] + '''';

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
      qryGlkwa.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';
      UV_fGulkwa := '';
      UV_fGulkwa1 := qryGlkwa.FieldByName('DESC_GLKWA1').AsString;
      UV_fGulkwa2 := qryGlkwa.FieldByName('DESC_GLKWA2').AsString;
      UV_fGulkwa3 := qryGlkwa.FieldByName('DESC_GLKWA3').AsString;
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
         if (Trim(sResult) = '') AND
            (qryHm.FieldByName('cod_hm').AsString <> 'H018') AND (qryHm.FieldByName('cod_hm').AsString <> 'H019') AND
            (qryHm.FieldByName('cod_hm').AsString <> 'H020') AND (qryHm.FieldByName('cod_hm').AsString <> 'H021') AND
            (qryHm.FieldByName('cod_hm').AsString <> 'H022') then
            bCheck := True;

         // 숫자인 경우 Low와 High를 비교
         eResult := GF_StrToFloat(sResult);

         if (eResult < eLow) or (eResult > eHigh) then
            bCheck := True;
      end
      else
      begin
         if qryHm.FieldByName('GUBN_HM').AsString = '2' then
            sTemp := sResult
         else if qryHm.FieldByName('GUBN_HM').AsString = '3' then
         begin
            UP_Separate(sResult, sTemp);
         end;
         // 문자인 경우 문자임상치와 비교
         if GF_SpaceDel(sTemp) <> GF_SpaceDel(sMunja) then
            bCheck := True;

         if Trim(sTemp) = '' then
            bCheck := True;
      end;

      // 중요6개 항목 체크..
      if ck6Check.Checked then
      begin
         if (qryHm.FieldByName('cod_hm').AsString = 'C026') OR (qryHm.FieldByName('cod_hm').AsString = 'C027') OR
            (qryHm.FieldByName('cod_hm').AsString = 'S014') OR (qryHm.FieldByName('cod_hm').AsString = 'S015') OR
            (qryHm.FieldByName('cod_hm').AsString = 'S018') OR (qryHm.FieldByName('cod_hm').AsString = 'S019') OR
            (qryHm.FieldByName('cod_hm').AsString = 'SE01') OR (qryHm.FieldByName('cod_hm').AsString = 'SE02') OR
            (qryHm.FieldByName('cod_hm').AsString = 'E003') OR
            (qryHm.FieldByName('cod_hm').AsString = 'SE31') OR (qryHm.FieldByName('cod_hm').AsString = 'T043') OR
            (qryHm.FieldByName('cod_hm').AsString = 'C083') OR (qryHm.FieldByName('cod_hm').AsString = 'Y004') OR
            (qryHm.FieldByName('cod_hm').AsString = 'U017') then
            bCheck := True;
      end;

      // 임상참고치를 벗어날경우
      if bCheck or not ckCheck.Checked then
      begin
         UP_ResultVarMemSet(iResult_cnt);

         UV_vSeq[iResult_cnt]      := (iResult_cnt+1);
         UV_vDesc_hm[iResult_cnt]  := qryHm.FieldByName('DESC_KOR').AsString;

         // 항목에 대한 이전결과자료 표시
         if UV_sBunjuchk = 'Y' then
         begin
            // 해당 Part검사에 대한 결과를 추출(Start Pos, End Pos를 이용)
            qryPrev.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';

            if qryPrev.RecordCount > 0 then
            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
               if (qryHm.FieldByName('POS_START').AsInteger <> 0) and (qryHm.FieldByName('POS_END').AsInteger <> 0) then
               begin
                    UV_vPrev_rslt[iResult_cnt] :=
                    Copy(UV_fGulkwa,
                       qryHm.FieldByName('POS_START').AsInteger,
                       qryHm.FieldByName('POS_END').AsInteger -
                       qryHm.FieldByName('POS_START').AsInteger + 1);
               end;        
            end;
         end
         else if UV_sBunjuchk = 'N' then
         begin
            // 해당 Part검사에 대한 결과를 추출(Start Pos, End Pos를 이용)
            qryPrev2006.Filter := 'GUBN_PART = ''' + qryHm.FieldByName('GUBN_PART').AsString + '''';

            if qryPrev2006.RecordCount > 0 then
            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryPrev2006.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryPrev2006.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryPrev2006.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
               if (qryHm.FieldByName('POS_START').AsInteger <> 0) and (qryHm.FieldByName('POS_END').AsInteger <> 0) then
               begin
                    UV_vPrev_rslt[iResult_cnt] :=
                    Copy(UV_fGulkwa,
                       qryHm.FieldByName('POS_START').AsInteger,
                       qryHm.FieldByName('POS_END').AsInteger -
                       qryHm.FieldByName('POS_START').AsInteger + 1);
               end;
            end;
         end;

         // 현재결과를 표시
         UV_vCur_rslt[iResult_cnt] := sResult;

         if qryHm.FieldByName('GUBN_HM').AsString = '1' then
         begin
            UV_vIms_low[iResult_cnt]   := eLow;
            UV_vIms_high[iResult_cnt]  := eHigh;
            UV_vDesc_unit[iResult_cnt] := qryHm.FieldByName('desc_unit').AsString;

            // 판정구분 판별
            if eResult < eMin2 then
               UV_vMin3[iResult_cnt] := '*'
            else if (eResult >= eMin2) and (eResult < eMin1) then
               UV_vMin2[iResult_cnt] := '*'
            else if (eResult >= eMin1) and (eResult < eLow) then
               UV_vMin1[iResult_cnt] := '*'
            else if (eResult >= eLow) and (eResult <= eHigh) then
               UV_vJung[iResult_cnt] := '*'
            else if (eResult > eHigh) and (eResult <= eMax1) then
               UV_vMax1[iResult_cnt] := '*'
            else if (eResult > eMax1) and (eResult <= eMax2) then
               UV_vMax2[iResult_cnt] := '*'
            else if eResult > eMax2 then
               UV_vMax3[iResult_cnt] := '*';
         end
         else if qryHm.FieldByName('GUBN_HM').AsString = '3' then
            begin
               UV_vIms_low[iResult_cnt] := eIms;
               UV_vDesc_unit[iResult_cnt] := qryHm.FieldByName('desc_unit').AsString;
               if  ((qryHm.FieldByName('cod_hm').AsString = 'U011') or           //20150507 뇨침사 검사는 따로 처리.
                    (qryHm.FieldByName('cod_hm').AsString = 'U012') or
                    (qryHm.FieldByName('cod_hm').AsString = 'U013')) and (Trim(sResult) <> '') then
               begin
                  UV_vIms_high[iResult_cnt]  := eHigh;               
                  if Trim(sResult) = 'many' then eResult := eMax2+1
                  else
                  begin
                    iCut := GF_CountInt(Trim(sResult));
                    eResult := StrToInt(Copy(sResult, 1, iCut));
                  end;
                  if (eResult >= eLow) and (eResult <= eHigh) then
                     UV_vJung[iResult_cnt] := '*'
                  else if (eResult > eHigh) and (eResult <= eMax1) then
                     UV_vMax1[iResult_cnt] := '*'
                  else if (eResult > eMax1) and (eResult <= eMax2) then
                     UV_vMax2[iResult_cnt] := '*'
                  else if eResult > eMax2 then
                     UV_vMax3[iResult_cnt] := '*';
               end;
            end;
         Inc(iResult_cnt);
      end;
   end;   // End of for

   grdResult.Rows := iResult_cnt;
end;

procedure TfrmLD6I02.UP_SokunQuery;
var i, index1, index2, index3 : Integer;
begin
   grdSokun1.Rows := 0;
   grdSokun2.Rows := 0;
   grdSokun3.Rows := 0;

   // 소견 Clear
   //2009.06.18 철. 소견15개로 늘림
   for i := 0 to 29 do
   begin
      UV_vPrev_cod1[i]   := '';
      UV_vCod_sokun1[i]  := '';
      UV_vDesc_sokun1[i] := '';
      UV_vInfo_numb1[i]  := '';
      UV_vCod_Doct[i]    := '';
      UV_vGubn_pan[i]    := '';

      UV_vPrev_cod2[i]   := '';
      UV_vCod_sokun2[i]  := '';
      UV_vDesc_sokun2[i] := '';
      UV_vInfo_numb2[i]  := '';

      UV_vPrev_cod3[i]   := '';
      UV_vCod_sokun3[i]  := '';
      UV_vDesc_sokun3[i] := '';
      UV_vInfo_numb3[i]  := '';
   end;

   // 현재위치를 추출
   i := grdMaster.CurrentDataRow - 1;

   index1 := 0;
   index2 := 0;
   index3 := 0;
   with qryHsok do
   begin
      Active := False;
      ParamByName('COD_BJJS').AsString  := UV_vCod_bjjs[i];
      ParamByName('NUM_ID').AsString    := UV_vNum_id[i];
      ParamByName('DAT_BUNJU').AsString := UV_vDat_bunju[i];
      ParamByName('NUM_BUNJU').AsString := UV_vNum_bunju[i];
      Active := True;

      if RecordCount > 0 then
      begin
         while not Eof do
         begin
            if FieldByName('GUBN_HSOK').AsString = '1' then
            begin
               // 혈액소견에 대한 자료
               UV_vPrev_cod1[index1]   := FieldByName('COD_SOKUN').AsString;
               UV_vCod_sokun1[index1]  := FieldByName('COD_SOKUN').AsString;
               UV_vDesc_sokun1[index1] := FieldByName('DESC_SOKUN').AsString;
               UV_vCod_doct[index1]    := FieldByname('COD_DOCT').AsString;
               UV_vGubn_pan[index1]    := FieldByname('GUBN_PAN').AsString;
               Inc(index1);
            end
            else if FieldByName('GUBN_HSOK').AsString = '2' then
            begin
               // 식이요법에 대한 자료
               UV_vPrev_cod2[index2]   := FieldByName('COD_SOKUN').AsString;
               UV_vCod_sokun2[index2]  := FieldByName('COD_SOKUN').AsString;
               UV_vDesc_sokun2[index2] := FieldByName('DESC_SOKUN').AsString;

               Inc(index2);
            end
            else if FieldByName('GUBN_HSOK').AsString = '3' then
            begin
               // 운동요법에 대한 자료
               UV_vPrev_cod3[index3]   := FieldByName('COD_SOKUN').AsString;
               UV_vCod_sokun3[index3]  := FieldByName('COD_SOKUN').AsString;
               UV_vDesc_sokun3[index3] := FieldByName('DESC_SOKUN').AsString;

               Inc(index3);
            end;

            Next;
         end;   // End of while
      end;

      Active := False;
   end;
   //[2016.03.19-소견30개로 늘림]
   grdSokun1.Rows := 30;
   grdSokun2.Rows := 5;
   grdSokun3.Rows := 5;
   if Trim(UV_vCod_doct[0]) <> '' then
   GP_ComboDisplay(cmbDoct, UV_vCod_doct[0],  6);

end;

procedure TfrmLD6I02.UP_Separate(sData : String; var sValue : String);
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

procedure TfrmLD6I02.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading := '분주번호';
         Col[1].Width := 67;
      end
      else if rbJumin.Checked then
      begin
         Col[1].Heading := '검진일자';
         Col[1].Width := 87;
      end;

      // Set Alignment
      Col[1].Alignment := taCenter;
   end;

   with grdResult do
   begin
      Rows := 0;

      // Column Title
      Col[1].Heading  := 'NO';
      Col[2].Heading  := '검사항목명';
      Col[3].Heading  := '임상';
      Col[4].Heading  := '참고치';
      Col[5].Heading  := '단위';

      Col[6].Heading  := '종전결과';
      Col[7].Heading  := '결과';
      Col[8].Heading  := '';
      Col[9].Heading  := '';
      Col[10].Heading  := '판';
      Col[11].Heading := '';
      Col[12].Heading := '정';
      Col[13].Heading := '';
      Col[14].Heading := '';

      // Column Color 설정
      Col[8].Color  := $00F9B0EE;
      Col[9].Color  := $00FBCCF3;
      Col[10].Color  := $00FDDFFA;
      Col[11].Color := $00FDF7AC;
      Col[12].Color := $00FDDFFA;
      Col[13].Color := $00FBCCF3;
      Col[14].Color := $00F9B0EE;

      // Column Title Alignment
      Col[3].HeadingAlign := True;
      Col[3].HeadingAlignment := taRightJustify;
      Col[4].HeadingAlign := True;
      Col[4].HeadingAlignment := taLeftJustify;

      // Column Align
      Col[1].Alignment := taCenter;
      Col[3].Alignment := taRightJustify;
      Col[4].Alignment := taRightJustify;
      Col[5].Alignment := taRightJustify;
      Col[6].Alignment := taRightJustify;
      Col[7].Alignment := taRightJustify;

      for i := 1 to Cols do
      begin
         Col[i].ReadOnly := True;

         if i < 7 then continue;
         Col[i].Alignment := taCenter;
      end;
   end;

   with grdSokun1 do
   begin
      //[2016.03.19-소견30개로 늘림]   
      Rows := 30;

      // Column Title
      Col[1].Heading  := '판정코드';
      Col[2].Heading  := '소견코드';
      Col[3].Heading  := '세부소견';

      // Column Align
      Col[1].Alignment := taCenter;

      // Max Length
      Col[1].MaxLength := 5;
      Col[2].MaxLength := 4;
      Col[3].MaxLength := 600;
   end;

   with grdSokun2 do
   begin
      Rows := 5;

      // Column Title
      Col[1].Heading  := '소견코드';
      Col[2].Heading  := '세부소견';

      // Column Align
      Col[1].Alignment := taCenter;

      // Max Length
      Col[1].MaxLength := 4;
      Col[2].MaxLength := 600;
   end;

   with grdSokun3 do
   begin
      Rows := 5;

      // Column Title
      Col[1].Heading  := '소견코드';
      Col[2].Heading  := '세부소견';

      // Column Align
      Col[1].Alignment := taCenter;

      // Max Length
      Col[1].MaxLength := 4;
      Col[2].MaxLength := 600;
   end;
end;

procedure TfrmLD6I02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Samp   := VarArrayCreate([0, 0], varOleStr);      
      UV_vDat_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cancer := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga  := VarArrayCreate([0, 0], varOleStr);
      UV_vBiman      := VarArrayCreate([0, 0], varOleStr);
      UV_vHul_h      := VarArrayCreate([0, 0], varOleStr);
      UV_vHul_l      := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_nosin  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_adult  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_agab   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_nsyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_adyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_agyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_tkgm   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_life   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_lfyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vPreg       := VarArrayCreate([0, 0], varOleStr);      
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_bunju,  iLength);
      VarArrayRedim(UV_vNum_Samp,   iLength);      
      VarArrayRedim(UV_vDat_bunju,  iLength);
      VarArrayRedim(UV_vCod_hul,    iLength);
      VarArrayRedim(UV_vCod_cancer, iLength);
      VarArrayRedim(UV_vCod_jangbi, iLength);
      VarArrayRedim(UV_vCod_chuga,  iLength);
      VarArrayRedim(UV_vBiman,      iLength);
      VarArrayRedim(UV_vHul_h,      iLength);
      VarArrayRedim(UV_vHul_l,      iLength);
      VarArrayRedim(UV_vCod_nosin,  iLength);
      VarArrayRedim(UV_vCod_adult,  iLength);
      VarArrayRedim(UV_vCod_agab,   iLength);
      VarArrayRedim(UV_vCod_nsyh,   iLength);
      VarArrayRedim(UV_vCod_adyh,   iLength);
      VarArrayRedim(UV_vCod_agyh,   iLength);
      VarArrayRedim(UV_vCod_tkgm,   iLength);
      VarArrayRedim(UV_vCod_life,   iLength);
      VarArrayRedim(UV_vCod_lfyh,   iLength);
      VarArrayRedim(UV_vPreg,       iLength);      

   end;
end;

procedure TfrmLD6I02.UP_ResultVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vSeq       := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hm   := VarArrayCreate([0, 0], varOleStr);
      UV_vIms_low   := VarArrayCreate([0, 0], varOleStr);
      UV_vIms_high  := VarArrayCreate([0, 0], varOleStr);
      UV_vPrev_rslt := VarArrayCreate([0, 0], varOleStr);
      UV_vCur_rslt  := VarArrayCreate([0, 0], varOleStr);
      UV_vMin3      := VarArrayCreate([0, 0], varOleStr);
      UV_vMin2      := VarArrayCreate([0, 0], varOleStr);
      UV_vMin1      := VarArrayCreate([0, 0], varOleStr);
      UV_vJung      := VarArrayCreate([0, 0], varOleStr);
      UV_vMax1      := VarArrayCreate([0, 0], varOleStr);
      UV_vMax2      := VarArrayCreate([0, 0], varOleStr);
      UV_vMax3      := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_unit := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vSeq,       iLength);
      VarArrayRedim(UV_vDesc_hm,   iLength);
      VarArrayRedim(UV_vIms_low,   iLength);
      VarArrayRedim(UV_vIms_high,  iLength);
      VarArrayRedim(UV_vPrev_rslt, iLength);
      VarArrayRedim(UV_vCur_rslt,  iLength);
      VarArrayRedim(UV_vMin3,      iLength);
      VarArrayRedim(UV_vMin2,      iLength);
      VarArrayRedim(UV_vMin1,      iLength);
      VarArrayRedim(UV_vJung,      iLength);
      VarArrayRedim(UV_vMax1,      iLength);
      VarArrayRedim(UV_vMax2,      iLength);
      VarArrayRedim(UV_vMax3,      iLength);
      VarArrayRedim(UV_vDesc_unit, iLength);
   end;
end;

function TfrmLD6I02.UF_Condition : Boolean;
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

function TfrmLD6I02.UF_Save : Boolean;
var i, iCur : Integer;
begin
   Result := False;

   // 현재위치 추출
   iCur := grdMaster.CurrentDataRow - 1;

   // 작업중인지 Check
   if not UV_bEdit then
   begin
      // 작업중이 아니더라도 분주번호에 소견확인 Update
      // 이유 : 자동소견에 의하여 만들어진 자료를 그대로 사용할 수 있기때문
      with qryU_Bnju do
      begin
         ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
         ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
         ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];

         Execsql;
         GP_query_log(qryU_Bnju, 'LD6I02', 'Q', 'Y');
      end;

      with qryU_Gumgin do
      begin
         ParamByName('COD_jisa').AsString   := UV_vCod_jisa[iCur];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
         ParamByName('DAT_GMGN').AsString   := UV_vDat_gmgn[iCur];
         ParamByName('NUM_Jumin').AsString  := UV_vNum_jumin[iCur];

         Execsql;
         GP_query_log(qryU_Gumgin, 'LD6I02', 'Q', 'Y');
      end;

      Result := True;
      exit;
   end;

   // Validation Check
   // 1. Not Null Check
   if cmbDoct.ItemIndex = -1 then
   begin
      GF_MsgBox('Warning', '담당의사를 먼저 입력하십시요.');
      cmbDoct.SetFocus;
      exit;
   end;

   if not GF_NotNullCheck(pnlDetail) then exit;

   //090328 철. 김소영선생님 소견저장시 '저장하시겠습니까' 멘트 안나오도록
   // Save Confirm message
   if (GV_sUserId <> '150005') and (GV_sUserId <> '150783') then
   begin
      if not GF_MsgBox('Confirmation', 'SAVE') then exit;
   end;

   // DB 작업
   dmComm.database.StartTransaction;
   try
      // 혈액소견
      with qryD_Hsok do
      begin
           ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
           ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
           ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
           ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];
           Execsql;
           GP_query_log(qryD_Hsok, 'LD6I02', 'Q', 'Y');
      end;

      for i := 0 to grdSokun1.Rows - 1 do
      begin
         // 새로운 소견이 입력될 경우
         If (GV_sJisa = '34') then
         begin
            with qryI_Hsok do
            begin
                 ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
                 ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
                 ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
                 ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];
                 ParamByName('COD_JISA').AsString   := UV_vCod_jisa[iCur];
                 ParamByName('DAT_GMGN').AsString   := UV_vDat_gmgn[iCur];
                 ParamByName('COD_SOKUN').AsString  := UV_vCod_sokun1[i];
                 ParamByName('DESC_SOKUN').AsString := UV_vDesc_sokun1[i];
                 ParamByName('COD_DOCT').AsString   := copy(cmbDoct.Text,1,6);
                 ParamByName('GUBN_PAN').AsString   := UV_vGubn_pan[i];
                 ParamByName('GUBN_HSOK').AsString  := '1';
                 ParamByName('NUM_SEQ').AsInteger   := i + 1;
                 ParamByName('DAT_INSERT').AsString := GV_sToday;
                 ParamByName('COD_INSERT').AsString := GV_sUserId;

                 Execsql;
                 GP_query_log(qryI_Hsok, 'LD6I02', 'Q', 'Y');
            end;
         end
         else if (trim(UV_vDesc_sokun1[i]) <> '') and (trim(UV_vCod_sokun1[i]) <> '') then
         begin
            with qryI_Hsok do
            begin
                 ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
                 ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
                 ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
                 ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];
                 ParamByName('COD_JISA').AsString   := UV_vCod_jisa[iCur];
                 ParamByName('DAT_GMGN').AsString   := UV_vDat_gmgn[iCur];
                 ParamByName('COD_SOKUN').AsString  := UV_vCod_sokun1[i];
                 ParamByName('DESC_SOKUN').AsString := UV_vDesc_sokun1[i];
                 ParamByName('COD_DOCT').AsString   := copy(cmbDoct.Text,1,6);
                 ParamByName('GUBN_PAN').AsString   := UV_vGubn_pan[i];
                 ParamByName('GUBN_HSOK').AsString  := '1';
                 ParamByName('NUM_SEQ').AsInteger   := i + 1;
                 ParamByName('DAT_INSERT').AsString := GV_sToday;
                 ParamByName('COD_INSERT').AsString := GV_sUserId;

                 Execsql;
                 GP_query_log(qryI_Hsok, 'LD6I02', 'Q', 'Y');
            end;
         end;

         // 작업Flag Clear
         UV_vInfo_numb2[i] := '';
         UV_vPrev_cod2[i]  := UV_vCod_sokun2[i];
      end;

      // 식이요법
      for i := 0 to grdSokun2.Rows - 1 do
      begin
            if UV_vDesc_sokun2[i] <> '' then
            begin
               // 새로운 소견이 입력될 경우
               with qryI_Hsok do
               begin
                  ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
                  ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
                  ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
                  ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];
                  ParamByName('COD_JISA').AsString   := UV_vCod_jisa[iCur];
                  ParamByName('DAT_GMGN').AsString   := UV_vDat_gmgn[iCur];
                  ParamByName('COD_SOKUN').AsString  := UV_vCod_sokun2[i];
                  ParamByName('DESC_SOKUN').AsString := UV_vDesc_sokun2[i];
                  ParamByName('COD_DOCT').AsString   := copy(cmbDoct.Text,1,6);
                  ParamByName('GUBN_PAN').AsString   := UV_vGubn_pan[i];
                  ParamByName('GUBN_HSOK').AsString  := '2';
                  ParamByName('NUM_SEQ').AsInteger   := i + 1;
                  ParamByName('DAT_INSERT').AsString := GV_sToday;
                  ParamByName('COD_INSERT').AsString := GV_sUserId;

                  Execsql;
                  GP_query_log(qryI_Hsok, 'LD6I02', 'Q', 'Y');
               end;
            end;

         // 작업Flag Clear
         UV_vInfo_numb2[i] := '';
         UV_vPrev_cod2[i]  := UV_vCod_sokun2[i];
      end;

      // 운동요법
      for i := 0 to grdSokun3.Rows - 1 do
      begin
         if UV_vDesc_sokun3[i] <> '' then
         begin
            // 새로운 소견이 입력될 경우
            with qryI_Hsok do
            begin
               ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
               ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
               ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
               ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];
               ParamByName('COD_JISA').AsString   := UV_vCod_jisa[iCur];
               ParamByName('DAT_GMGN').AsString   := UV_vDat_gmgn[iCur];
               ParamByName('COD_SOKUN').AsString  := UV_vCod_sokun3[i];
               ParamByName('DESC_SOKUN').AsString := UV_vDesc_sokun3[i];
               ParamByName('GUBN_HSOK').AsString  := '3';
               ParamByName('COD_DOCT').AsString   := copy(cmbDoct.Text,1,6);
               ParamByName('GUBN_PAN').AsString   := UV_vGubn_pan[i];
               ParamByName('NUM_SEQ').AsInteger   := i + 1;
               ParamByName('DAT_INSERT').AsString := GV_sToday;
               ParamByName('COD_INSERT').AsString := GV_sUserId;

               Execsql;
               GP_query_log(qryI_Hsok, 'LD6I02', 'Q', 'Y');
            end;
         end;

         // 작업Flag Clear
         UV_vInfo_numb3[i] := '';
         UV_vPrev_cod3[i]  := UV_vCod_sokun3[i];
      end;

      // 소견확인 Update
      with qryU_Bnju do
      begin
         ParamByName('COD_BJJS').AsString   := UV_vCod_bjjs[iCur];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[iCur];
         ParamByName('DAT_BUNJU').AsString  := UV_vDat_bunju[iCur];
         ParamByName('NUM_BUNJU').AsString  := UV_vNum_bunju[iCur];

         Execsql;
         GP_query_log(qryU_Bnju, 'LD6I02', 'Q', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;

   // Database Commit
   dmComm.database.Commit;

   // Cell Color 원래대로 표시
   grdSokun1.ResetCellProperties([prFont]);
   grdSokun2.ResetCellProperties([prFont]);
   grdSokun3.ResetCellProperties([prFont]);

   iCmb_Index := cmbDoct.ItemIndex;
   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD6I02.btnSaveClick(Sender: TObject);
begin
   inherited;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD6I02.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboSawon(cmbDoct, 1);
   Cmb_gubn.ItemIndex:=1;

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);
   UP_ResultVarMemSet(0);
   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

   UV_vCod_sokun1  := VarArrayCreate([0, 29], varOleStr);
   UV_vDesc_sokun1 := VarArrayCreate([0, 29], varOleStr);
   UV_vPrev_cod1   := VarArrayCreate([0, 29], varOleStr);
   UV_vInfo_numb1  := VarArrayCreate([0, 29], varOleStr);
   UV_vCod_doct    := VarArrayCreate([0, 29], varOleStr);
   UV_vGubn_pan    := VarArrayCreate([0, 29], varOleStr);

   UV_vCod_sokun2  := VarArrayCreate([0, 29], varOleStr);
   UV_vDesc_sokun2 := VarArrayCreate([0, 29], varOleStr);
   UV_vPrev_cod2   := VarArrayCreate([0, 29], varOleStr);
   UV_vInfo_numb2  := VarArrayCreate([0, 29], varOleStr);

   UV_vCod_sokun3  := VarArrayCreate([0, 29], varOleStr);
   UV_vDesc_sokun3 := VarArrayCreate([0, 29], varOleStr);
   UV_vPrev_cod3   := VarArrayCreate([0, 29], varOleStr);
   UV_vInfo_numb3  := VarArrayCreate([0, 29], varOleStr);

   // 항목자료 Query
   if not qryHm.Active then qryHm.Active := True;

   UV_sBasicSQL := qryBunju.SQL.Text;
   iCmb_Index := -1;
   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;
end;

procedure TfrmLD6I02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : begin
             if (rbDate.Checked) or (rbBarcode.Checked) then
                Value := UV_vNum_bunju[DataRow-1]
             else if rbJumin.Checked then
                Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);

          end;
   end;
end;

procedure TfrmLD6I02.btnQueryClick(Sender: TObject);
var i, index ,iage: Integer;
    sWhere, sOrderBy,Sman : String;
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
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   // 각 혈액파트 완료체크[분주번호]...
   if (rbDate.Checked) and ((GV_sJisa = '00') or (GV_sJisa = '15')) then
   begin
      btnSave.Enabled    := True;
      btnCancel.Enabled  := True;
      btnPSave.Enabled   := True;
      btnPCancel.Enabled := True;

      if GF_HulGulkwaCheck(1, mskDate.Text, copy(GV_sUserId,1,2), UV_sHul_Part) then
      begin
         if not UV_bRead then
         begin
            if not GF_MsgBox('Confirmation', '다음 혈액파트[' + UV_sHul_Part + ']가 검사중입니다.' + #13#13 +
                                             '그래도 작업하시겠습니까?[Yes - 소견작업, No - 조회]') then
            begin
               btnSave.Enabled    := False;
               btnCancel.Enabled  := False;
               btnPSave.Enabled   := False;
               btnPCancel.Enabled := False;
            end
            else
            begin
               btnSave.Enabled    := True;
               btnCancel.Enabled  := True;
               btnPSave.Enabled   := True;
               btnPCancel.Enabled := True;
            end; // end of if[GF_MsgBox];
         end
         else if UV_bRead then
         begin
            btnSave.Enabled    := False;
            btnCancel.Enabled  := False;
            btnPSave.Enabled   := False;
            btnPCancel.Enabled := False;
            if not GF_MsgBox('Confirmation', '다음 혈액파트[' + UV_sHul_Part + ']가 검사중입니다.' + #13#13 +
                                             '그래도 조회하시겠습니까?[Yes - 조회, No - 미조회]') then
            begin
               // Field Clear & 첫번째 Field Focus
               GP_FieldClear(pnlDetail);
               grdSokun1.Rows := 0;
               grdSokun2.Rows := 0;
               grdSokun3.Rows := 0;
               exit;
            end;
         end;
      end; // end of if[GF_HulGulkwaCheck];
   end; // end of if[rbDate];

   // Query 수행 & 배열에 저장
   index := 0;
   with qryBunju do
   begin
      // SQL문 생성
      Active := False;

      if GV_sJisa = '00' then
      begin
         // '00'이면 지사선택한 자료만 조회
         sWhere := ' A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
      end
      else
      begin
         // '00'이 아니면 해당지사에 대한 자료를 조회
         sWhere := ' A.cod_bjjs = ''' + GV_sJisa + '''';
      end;

      if rbDate.Checked then
      begin
         if Cmb_gubn.Text = '분주번호' then
         begin
              sWhere := ' WHERE ' + sWhere + 'AND A.dat_bunju = ''' + mskDate.Text + '''';
              if Trim(mskFrom.Text) <> '' then
                 sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
              if Trim(mskTo.Text) <> '' then
                 sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
              sOrderBy := ' ORDER BY A.dat_bunju, A.num_bunju ';
         end
         else
         begin
             sWhere := ' WHERE ' + sWhere + 'AND A.dat_bunju = ''' + mskDate.Text + '''';
             if Trim(MEdt_SampS.Text) <> '' then
                sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
             if Trim(MEdt_SampE.Text) <> '' then
                sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
             sOrderBy := ' ORDER BY A.dat_bunju, B.num_samp ';

         end;
      end
      else if rbJumin.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND B.num_jumin = ''' + mskJumin.Text + '''';

         sOrderBy := ' ORDER BY B.dat_gmgn ';
      end
      else if rbBarcode.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + ' AND B.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
         sWhere :=    sWhere + ' AND B.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';
         sOrderBy := ' ORDER BY B.dat_gmgn ';
      end;


      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD6I02', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);
         
         for i := 0 to RecordCount - 1 do
         begin
            GP_JuminToAge(FieldByName('NUM_JUMIN').AsString,FieldByName('DAt_gmgn').AsString, sMan, iAge);
            if (iAge<=18) and (Chk_Age.checked=true)  then
            begin
                 UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                 UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                 UV_vCod_jisa[index]   := FieldByName('COD_JISA').AsString;
                 UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                 UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                 UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                 UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                 UV_vNum_Samp[index]  := FieldByName('NUM_Samp').AsString;                 
                 UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                 UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                 UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                 UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                 UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                 UV_vBiman[index]      := FieldByName('BIMAN').AsString;
                 UV_vHul_h[index]      := FieldByName('HUL_H').AsString;
                 UV_vHul_l[index]      := FieldByName('HUL_L').AsString;
                 UV_vCod_nosin[index]  := FieldByName('GUBN_NOSIN').AsString;
                 UV_vCod_adult[index]  := FieldByName('GUBN_ADULT').AsString;
                 UV_vCod_agab[index]   := FieldByName('GUBN_AGAB').AsString;
                 UV_vCod_nsyh[index]   := FieldByName('GUBN_NSYH').AsString;
                 UV_vCod_adyh[index]   := FieldByName('GUBN_ADYH').AsString;
                 UV_vCod_agyh[index]   := FieldByName('GUBN_AGYH').AsString;
                 UV_vCod_tkgm[index]   := FieldByName('GUBN_TKGM').AsString;
                 UV_vCod_life[index]   := FieldByName('GUBN_life').AsString;
                 UV_vCod_lfyh[index]   := FieldByName('GUBN_lfyh').AsString;
                 UV_vPreg[index]       := FieldByName('GUBN_preg').AsString;
                 Inc(index);
            end
            else if  (Chk_Age.checked=false)  then
            begin
                 UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
                 UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                 UV_vCod_jisa[index]   := FieldByName('COD_JISA').AsString;
                 UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                 UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                 UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                 UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                 UV_vNum_Samp[index]   := FieldByName('NUM_Samp').AsString;                 
                 UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                 UV_vCod_hul[index]    := FieldByName('COD_HUL').AsString;
                 UV_vCod_cancer[index] := FieldByName('COD_CANCER').AsString;
                 UV_vCod_jangbi[index] := FieldByName('COD_JANGBI').AsString;
                 UV_vCod_chuga[index]  := FieldByName('COD_CHUGA').AsString;
                 UV_vBiman[index]      := FieldByName('BIMAN').AsString;
                 UV_vHul_h[index]      := FieldByName('HUL_H').AsString;
                 UV_vHul_l[index]      := FieldByName('HUL_L').AsString;
                 UV_vCod_nosin[index]  := FieldByName('GUBN_NOSIN').AsString;
                 UV_vCod_adult[index]  := FieldByName('GUBN_ADULT').AsString;
                 UV_vCod_agab[index]   := FieldByName('GUBN_AGAB').AsString;
                 UV_vCod_nsyh[index]   := FieldByName('GUBN_NSYH').AsString;
                 UV_vCod_adyh[index]   := FieldByName('GUBN_ADYH').AsString;
                 UV_vCod_agyh[index]   := FieldByName('GUBN_AGYH').AsString;
                 UV_vCod_tkgm[index]   := FieldByName('GUBN_TKGM').AsString;
                 UV_vCod_life[index]   := FieldByName('GUBN_life').AsString;
                 UV_vCod_lfyh[index]   := FieldByName('GUBN_lfyh').AsString;
                 UV_vPreg[index]       := FieldByName('GUBN_preg').AsString;
                 Inc(index);
            end;
            next;
         end;
        
         // Table과 Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         grdSokun1.Rows := 0;
         grdSokun2.Rows := 0;
         grdSokun3.Rows := 0;
         
         GF_MsgBox('Information', 'NODATA');
      end;
   end;


   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD6I02.btnCancelClick(Sender: TObject);
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
      cmbDoct.ItemIndex := iCmb_Index;
   end;
   // 자료 Display
   UP_SokunQuery;

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD6I02.UP_Change(Sender: TObject);
begin
   inherited;

   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행


   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD6I02.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, iRet, i, temp1, Temp : Integer;
    vCod_chuga, vCod_totyh, vCod_Tkgum : Variant;
    sName, cod_name, sTot_hangmok,sJisaName : String;
    eSuga : Extended;
    ckeck1 : Boolean;
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // 각 혈액파트 완료체크[주민번호]...
   if (rbJumin.Checked) and ((GV_sJisa = '00') or (GV_sJisa = '15')) then
   begin
      btnSave.Enabled    := True;
      btnCancel.Enabled  := True;
      btnPSave.Enabled   := True;
      btnPCancel.Enabled := True;

      UV_sHul_Part := UV_vDat_gmgn[NewRow-1];
      if GF_HulGulkwaCheck(2, mskJumin.Text, UV_vCod_bjjs[NewRow-1], UV_sHul_Part) then
      begin
         if not UV_bRead then
         begin
            if not GF_MsgBox('Confirmation', '다음 혈액파트[' + UV_sHul_Part + ']가 검사중입니다.' + #13#13 +
                                             '그래도 작업하시겠습니까?[Yes - 소견작업, No - 조회]') then
            begin
               btnSave.Enabled    := False;
               btnCancel.Enabled  := False;
               btnPSave.Enabled   := False;
               btnPCancel.Enabled := False;
            end
            else
            begin
               btnSave.Enabled    := True;
               btnCancel.Enabled  := True;
               btnPSave.Enabled   := True;
               btnPCancel.Enabled := True;
            end; // end of if[GF_MsgBox];
         end
         else if UV_bRead then
         begin
            btnSave.Enabled    := False;
            btnCancel.Enabled  := False;
            btnPSave.Enabled   := False;
            btnPCancel.Enabled := False;
            if not GF_MsgBox('Confirmation', '다음 혈액파트[' + UV_sHul_Part + ']가 검사중입니다.' + #13#13 +
                                             '그래도 조회하시겠습니까?[Yes - 조회, No - 미조회]') then
            begin
               // Field Clear & 첫번째 Field Focus
               GP_FieldClear(pnlDetail);
               grdSokun1.Rows := 0;
               grdSokun2.Rows := 0;
               grdSokun3.Rows := 0;
               exit;
            end;
         end;
      end; // end of if[GF_HulGulkwaCheck];
   end; // end of if[rbJumin];

   ckeck1 := True;
   sJisaName:='';
   // Grid의 Row가 변경되면 Detail에 자료를 할당
   mskNUM_JUMIN.Text := UV_vNum_jumin[NewRow-1];
   edtDESC_NAME.Text := UV_vDesc_name[NewRow-1];
   mskDAT_GMGN.Text  := UV_vDat_gmgn[NewRow-1];
   mskDAT_BUNJU.Text := UV_vDat_bunju[NewRow-1];
   edtNUM_BUNJU.Text := UV_vNum_bunju[NewRow-1];
   edtNUM_Samp.Text  := UV_vNum_Samp[NewRow-1];
   edt_Jisa.Text := UV_vCod_jisa[NewRow-1];
   GF_JisaEdit(edt_Jisa, sJisaName);
   edt_Jisa.Text     := sJisaName;



   if  UV_vPreg[NewRow-1] ='Y' then
       edt_Preg.Text:='임신'
   else if  UV_vPreg[NewRow-1] ='P' then
       edt_Preg.Text:='임신가능성(유)'
   else edt_Preg.Text:='';

   if (UV_vCod_hul[NewRow-1] = 'B100') or (UV_vCod_hul[NewRow-1] = 'A100') or
      (UV_vCod_hul[NewRow-1] = 'AB01') or (UV_vCod_hul[NewRow-1] = 'AB02') then
   begin
       EdtCOD_HUL.Color := $00C8FFFF;
       EdtCOD_HUL.Text  := UV_vCod_hul[NewRow-1];
   end
   else
   begin
       EdtCOD_HUL.Color := clBtnFace;
       EdtCOD_HUL.Text  := UV_vCod_hul[NewRow-1];
   end;
   EdtCOD_CANCER.Text := UV_vCod_cancer[NewRow-1];

   if (UV_vHul_h[NewRow-1] >= 140) or (UV_vHul_l[NewRow-1] >= 100) then
   begin
       edtHUL.Color :=  clAqua;
       edtHUL.Text   := UV_vHul_h[NewRow-1] + '/' + UV_vHul_l[NewRow-1];
   end
   else
   begin
       edtHUL.Color :=  clBtnFace;
       edtHUL.Text   := UV_vHul_h[NewRow-1] + '/' + UV_vHul_l[NewRow-1];
   end;
   edtBIMAN.Text     := UV_vBiman[NewRow-1];

   if (GF_StrToNum(edtBiman.Text) >= 110) and (GF_StrToNum(edtBiman.Text) < 120) then
      edtBiman.Color := clAqua
   else
   if GF_StrToNum(edtBiman.Text) >= 120 then
      edtBiman.Color := $00FACFFA
   else
      edtBiman.Color := clBtnFace;


   // 혈액검사 Profile명 조회
   edtCOD_PF.Text    := UV_vCod_hul[NewRow-1];
   GF_ProfileEdit(edtCOD_PF, '2', sName, eSuga);
   edtCOD_PF.Text    := sName;

   // 검사코드를 추출
   index := 0;
   sTot_hangmok := '';

   FillChar(UV_vCod_hm, SizeOf(UV_vCod_hm), 0);     //20151023 곽윤설, UV_vCod_hm 초기화
   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

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
               sTot_hangmok      := sTot_hangmok + FieldByName('COD_HM').AsString;
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
               sTot_hangmok      := sTot_hangmok + FieldByName('COD_HM').AsString;
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
               sTot_hangmok      := sTot_hangmok + FieldByName('COD_HM').AsString;
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
      sTot_hangmok      := sTot_hangmok + vCod_chuga[i];
      Inc(index);
   end;

   // 4. 노신, 성인병, 간염에 대한 검사항목 추출
   cod_name := '';
   // 노신유형Display
   if UV_vCod_nosin[NewRow-1] = '1' then
      cod_name := cod_name + UF_No_Hangmok(copy(UV_vDat_gmgn[NewRow-1],1,4), '1', UV_vCod_nsyh[NewRow-1] );
   // 재검테이블을 사용안함.
{ else if FieldByName('gubn_nosin').AsString = '2' then
   cod_name := cod_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) }

   //성인병유형Display
   if UV_vCod_adult[NewRow-1] = '1' then
      cod_name := cod_name + UF_No_Hangmok(copy(UV_vDat_gmgn[NewRow-1],1,4), '4', UV_vCod_adyh[NewRow-1]);
   // 재검테이블을 사용안함.
{  else if UV_vCod_adult[NewRow-1] = '2' then
      cod_name := cod_name + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '1', UV_vCod_adyh[NewRow-1]);}

   //간염유형Display
   if UV_vCod_agab[NewRow-1] = '1' then
      cod_name := cod_name + UF_No_Hangmok(copy(UV_vDat_gmgn[NewRow-1],1,4), '5', UV_vCod_agyh[NewRow-1]);
{   else if UV_vCod_agab[NewRow-1] = '2' then
      cod_name := cod_name + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '1', UV_vCod_agyh[NewRow-1]) ;}

   //07.05.08 철. 생애전환기검사유형 Display
   if UV_vCod_life[NewRow-1] = '1' then
      cod_name := cod_name + UF_No_Hangmok(copy(UV_vDat_gmgn[NewRow-1],1,4), '7', UV_vCod_lfyh[NewRow-1]);

   //특검유형Display
   if (UV_vCod_tkgm[NewRow-1] = '1') or (UV_vCod_tkgm[NewRow-1] = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
      qryTkgum.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryTkgum.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_Tkgum);
         for Temp := 0 to iRet - 1 do
         begin
           qryPf_hm.Active := False;
           qryPf_hm.ParamByName('cod_pf').AsString := vCod_Tkgum[Temp];
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

   iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
   for i := 0 to iRet-1 do
   begin
     if copy(vCod_totyh[i],1,2) <> 'JJ' then
     begin
       ckeck1 := True;
       if index <> 0 then
       begin
         for temp1 := 0 to index do
         begin
           if UV_vCod_hm[temp1] = vCod_totyh[i] then ckeck1 := False;
         end;
       end;

       if ckeck1 = True then
       begin
         VarArrayRedim(UV_vCod_hm, index + 1);
         UV_vCod_hm[index] := vCod_totyh[i];
         Inc(index);
       end;
     end;
   end;

   // 분주번호에 대한 혈액결과를 획득
   with qryGlkwa do
   begin
      Active := False;
      ParamByName('COD_BJJS').AsString  := UV_vCod_bjjs[NewRow-1];
      ParamByName('NUM_ID').AsString    := UV_vNum_id[NewRow-1];
      ParamByName('DAT_BUNJU').AsString := UV_vDat_bunju[NewRow-1];
      ParamByName('NUM_BUNJU').AsString := UV_vNum_bunju[NewRow-1];
      Active := True;
   end;

   UV_sBunjuchk := 'Y'; mskPDAT_GMGN.Text := '';
   // 해당검진자에 대한 이전검사결과를 획득
   qryinjouk.Active := False;
   qryinjouk.ParamByName('NUM_JUMIN').AsString := UV_vNum_jumin[NewRow-1];
   qryinjouk.ParamByName('DAT_GMGN').AsString  := UV_vDat_gmgn[NewRow-1];
   qryinjouk.Active := True;


   qryPrev.Active := False;
   qryPrev.Filter := '';
   qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
   qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
   qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
   qryPrev.Active := True;

   mskPDAT_GMGN.Text := qryinjouk.FieldByName('dat_gmgn').AsString;
   if qryPrev.RecordCount = 0 then
   begin
      qryPrev.Active := False;
      UV_sBunjuchk := 'N';
      qryInjouk2006.Active := False;
      qryInjouk2006.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryInjouk2006.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryInjouk2006.Active := True;

      qryPrev2006.Active := False;
      qryPrev2006.Filter := '';
      qryPrev2006.ParamByName('COD_JISA').AsString := qryInjouk2006.FieldByName('cod_jisa').AsString;;
      qryPrev2006.ParamByName('NUM_ID').AsString   := qryInjouk2006.FieldByName('num_id').AsString;
      qryPrev2006.ParamByName('DAT_GMGN').AsString := qryInjouk2006.FieldByName('dat_gmgn').AsString;
      qryPrev2006.Active := True;
      mskPDAT_GMGN.Text := qryPrev2006.FieldByName('dat_gmgn').AsString;
   end;

  { // 검체불량자에 대한 자료 획득
   with qryGum_bul do
   begin
      Active := False;
      ParamByName('COD_BJJS').AsString  := UV_vCod_bjjs[NewRow-1];
      ParamByName('NUM_ID').AsString    := UV_vNum_id[NewRow-1];
      ParamByName('DAT_BUNJU').AsString := UV_vDat_bunju[NewRow-1];
      ParamByName('NUM_BUNJU').AsString := UV_vNum_bunju[NewRow-1];
      Active := True;

  //    edtHemo.Text := '';
  //    edtLipe.Text := '';

      if RecordCount > 0 then
      begin
         // 검체용열 자료가 존재하는지 Check
         Filter := ' GUBN_BUL = ''02''';
         if (RecordCount > 0) and (FieldByName('DAT_SOLVE').AsString = '') then edtHemo.Text := 'Y';

         // 검체혼탁 자료가 존재하는지 Check
         Filter := ' GUBN_BUL = ''01''';
         if (RecordCount > 0) and (FieldByName('DAT_SOLVE').AsString = '') then edtLipe.Text := 'Y';
      end;
   end;
        }
   sTot_hangmok := sTot_hangmok + cod_name;



   
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


   qry_Gum_bul.Active := False;
   qry_Gum_bul.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qry_Gum_bul.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qry_Gum_bul.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qry_Gum_bul.Active := True;
   if qry_Gum_bul.RecordCount > 0 then
   begin
      while not qry_Gum_bul.Eof  do
      begin
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '01' then
         begin
            ckGUBN_JUNGDO1.Checked := True;
            mskDAT_SOLVE1.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '02' then
         begin
            ckGUBN_JUNGDO2.Checked := True;
            mskDAT_SOLVE2.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '03' then
         begin
            ckGUBN_JUNGDO3.Checked := True;
            mskDAT_SOLVE3.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '04' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '05' then
         begin
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '06' then
         begin
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '07' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '08' then
         begin
            ckGUBN_JUNGDO8.Checked := True;
            mskDAT_SOLVE8.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
            JD_sayu.Text :=qry_Gum_bul.FieldByName('JD_sayu').AsString;
         end;
         if qry_Gum_bul.FieldByName('gubn_bul').AsString = '09' then
         begin
            ckGUBN_JUNGDO9.Checked := True;
            mskDAT_SOLVE9.Text := qry_Gum_bul.FieldByName('dat_solve').AsString;
         end;

         qry_Gum_bul.Next;
      end;
   end;
//------------------------------------------------------------------------------


   // Header에 존재하는 검사항목에 대한 자료 획득
   edtH023.Text := UF_HangmokResult(qryGlkwa, 'H023');
   edtH024.Text := UF_HangmokResult(qryGlkwa, 'H024');
   if UV_sBunjuchk = 'Y' then
   begin
      if pos('SE07', sTot_hangmok) > 0 then
      begin
         if (Trim(UF_HangmokResult(qryPrev, 'SE07')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE07')) = '양성')
             or (Trim(UF_HangmokResult(qryPrev, 'SE07')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE07')) = '약양성') then
         begin
            edtS007.Color := clAqua;
            edtS007.Text  := Trim(UF_HangmokResult(qryPrev,  'SE07')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE07'));
         end
         else
         begin
            edtS007.Color := clBtnFace;
            edtS007.Text  := Trim(UF_HangmokResult(qryPrev,  'SE07')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE07'));
         end;
      end;
      if pos('SE08', sTot_hangmok) > 0 then
      begin
         if (Trim(UF_HangmokResult(qryPrev, 'SE08')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE08')) = '양성')
             or (Trim(UF_HangmokResult(qryPrev, 'SE08')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE08')) = '약양성') then
         begin
            edtS008.Color := clAqua;
            edtS008.Text  := Trim(UF_HangmokResult(qryPrev,  'SE08')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE08'));
         end
         else
         begin
            edtS008.Color := clBtnFace;
            edtS008.Text  := Trim(UF_HangmokResult(qryPrev,  'SE08')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE08'));
         end;
      end;

      if (Trim(UF_HangmokResult(qryPrev, 'S007')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S007')) = '양성')
          or (Trim(UF_HangmokResult(qryPrev, 'S007')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S007')) = '약양성') then
      begin
         edtS007.Color := clAqua;
         edtS007.Text  := Trim(UF_HangmokResult(qryPrev,  'S007')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S007'));
      end
      else
      begin
         edtS007.Color := clBtnFace;
         edtS007.Text  := Trim(UF_HangmokResult(qryPrev,  'S007')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S007'));
      end;
      if (Trim(UF_HangmokResult(qryPrev, 'S008')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S008')) = '양성')
          or (Trim(UF_HangmokResult(qryPrev, 'S008')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S008')) = '약양성') then
      begin
         edtS008.Color := clAqua;
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev,  'S008')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S008'));
      end
      else
      begin
         edtS008.Color := clBtnFace;
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev,  'S008')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S008'));
      end;
               
      //2012.04.19 S091 추가
      if pos('S091', sTot_hangmok) > 0 then
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev,  'S091')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S091'));


      if (Trim(UF_HangmokResult(qryPrev,  'S034')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S034')) = '양성') then
      begin
         edtS034.Color := clAqua;
         edtS034.Text := Trim(UF_HangmokResult(qryPrev,  'S034')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S034'));
      end
      else
      begin
         edtS034.Color := clBtnFace;
         edtS034.Text := Trim(UF_HangmokResult(qryPrev,  'S034')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S034'));
      end;

      if (Trim(UF_HangmokResult(qryPrev,  'S020')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S020')) = '양성') then
      begin
         edtS020.Color := clAqua;
         edtS020.Text  := Trim(UF_HangmokResult(qryPrev,  'S020')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S020'));
      end
      else
      begin
         edtS020.Color := clBtnFace;
         edtS020.Text  := Trim(UF_HangmokResult(qryPrev,  'S020')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S020'));
      end;

      if (Trim(UF_HangmokResult(qryPrev,  'S003')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S003')) = '양성') then
      begin
         edtS003.Color := clAqua;
         edtS003.Text  := Trim(UF_HangmokResult(qryPrev,  'S003')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S003'));
      end
      else
      begin
         edtS003.Color := clBtnFace;
         edtS003.Text  := Trim(UF_HangmokResult(qryPrev,  'S003')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S003'));
      end;

      if (Trim(UF_HangmokResult(qryPrev,  'TT01')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'TT01')) = '양성') then
      begin
         edtTT01.Color := clAqua;
         edtTT01.Text := Trim(UF_HangmokResult(qryPrev,  'TT01')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'TT01'));
      end
      else
      begin
         edtTT01.Color := clBtnFace;
         edtTT01.Text := Trim(UF_HangmokResult(qryPrev,  'TT01')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'TT01'));
      end;

      if (Trim(UF_HangmokResult(qryPrev,  'S001')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S001')) = '양성') then
      begin
         edtS001.Color := clAqua;
         edtS001.Text := Trim(UF_HangmokResult(qryPrev,  'S001')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S001'));
      end
      else
      begin
         edtS001.Color := clBtnFace;
         edtS001.Text := Trim(UF_HangmokResult(qryPrev,  'S001')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S001'));
      end;
   end
   else
   begin
      if pos('SE07', sTot_hangmok) > 0 then
      begin
         if (Trim(UF_HangmokResult(qryPrev2006, 'SE07')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE07')) = '양성')
             or (Trim(UF_HangmokResult(qryPrev2006, 'SE07')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE07')) = '약양성') then
         begin
            edtS007.Color := clAqua;
            edtS007.Text  := Trim(UF_HangmokResult(qryPrev2006,  'SE07')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE07'));
         end
         else
         begin
            edtS007.Color := clBtnFace;
            edtS007.Text  := Trim(UF_HangmokResult(qryPrev2006,  'SE07')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE07'));
         end;
      end;
      if pos('SE08', sTot_hangmok) > 0 then
      begin
         if (Trim(UF_HangmokResult(qryPrev2006, 'SE08')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE08')) = '양성')
             or (Trim(UF_HangmokResult(qryPrev2006, 'SE08')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'SE08')) = '약양성') then
         begin
            edtS008.Color := clAqua;
            edtS008.Text  := Trim(UF_HangmokResult(qryPrev2006,  'SE08')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE08'));
         end
         else
         begin
            edtS008.Color := clBtnFace;
            edtS008.Text  := Trim(UF_HangmokResult(qryPrev2006,  'SE08')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'SE08'));
         end;
      end;

      if (Trim(UF_HangmokResult(qryPrev2006, 'S007')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S007')) = '양성')
          or (Trim(UF_HangmokResult(qryPrev2006, 'S007')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S007')) = '약양성')then
      begin
         edtS007.Color := clAqua;
         edtS007.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S007')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S007'));
      end
      else
      begin
         edtS007.Color := clBtnFace;
         edtS007.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S007')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S007'));
      end;
      if (Trim(UF_HangmokResult(qryPrev2006, 'S008')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S008')) = '양성')
          or (Trim(UF_HangmokResult(qryPrev2006, 'S008')) = '약양성') or (Trim(UF_HangmokResult(qryGlkwa, 'S008')) = '약양성')then
      begin
         edtS008.Color := clAqua;
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S008')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S008'));
      end
      else
      begin
         edtS008.Color := clBtnFace;
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S008')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S008'));
      end;
               
      //2012.04.19 S091 추가
      if pos('S091', sTot_hangmok) > 0 then
         edtS008.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S091')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S091'));

      if (Trim(UF_HangmokResult(qryPrev2006,  'S034')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S034')) = '양성') then
      begin
         edtS034.Color := clAqua;
         edtS034.Text := Trim(UF_HangmokResult(qryPrev2006,  'S034')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S034'));
      end
      else
      begin
         edtS034.Color := clBtnFace;
         edtS034.Text := Trim(UF_HangmokResult(qryPrev2006,  'S034')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S034'));
      end;

      if (Trim(UF_HangmokResult(qryPrev2006,  'S020')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S020')) = '양성') then
      begin
         edtS020.Color := clAqua;
         edtS020.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S020')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S020'));
      end
      else
      begin
         edtS020.Color := clBtnFace;
         edtS020.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S020')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S020'));
      end;

      if (Trim(UF_HangmokResult(qryPrev2006,  'S003')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S003')) = '양성') then
      begin
         edtS003.Color := clAqua;
         edtS003.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S003')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S003'));
      end
      else
      begin
         edtS003.Color := clBtnFace;
         edtS003.Text  := Trim(UF_HangmokResult(qryPrev2006,  'S003')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S003'));
      end;

      if (Trim(UF_HangmokResult(qryPrev2006,  'TT01')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa, 'TT01')) = '양성') then
      begin
         edtTT01.Color := clAqua;
         edtTT01.Text := Trim(UF_HangmokResult(qryPrev2006,  'TT01')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'TT01'));
      end
      else
      begin
         edtTT01.Color := clBtnFace;
         edtTT01.Text := Trim(UF_HangmokResult(qryPrev2006,  'TT01')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'TT01'));
      end;

      if (Trim(UF_HangmokResult(qryPrev2006,  'S001')) = '양성') or (Trim(UF_HangmokResult(qryGlkwa,  'S001')) = '양성') then
      begin
         edtS001.Color := clAqua;
         edtS001.Text := Trim(UF_HangmokResult(qryPrev2006,  'S001')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S001'));
      end
      else
      begin
         edtS001.Color := clBtnFace;
         edtS001.Text := Trim(UF_HangmokResult(qryPrev2006,  'S001')) + '/' + Trim(UF_HangmokResult(qryGlkwa, 'S001'));
      end;
   end;

   if Trim(edtH024.Text) = 'Rh-' then
      edtH024.Color :=  $00FACFFA
   else edtH024.Color := clBtnFace;

   // 결과자료 Query
   UP_ResultQuery;
   // 의사Setting
   if Trim(cmbDoct.Text) = '' then
      cmbDoct.ItemIndex := iCmb_Index;
   // 소견자료 Query
   UP_SokunQuery;

   // 첫번째 소견자료를 Default로 표시
   pgcSokun.ActivePage := TabSheet1;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY') 
end;

procedure TfrmLD6I02.btnPInsertClick(Sender: TObject);
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

procedure TfrmLD6I02.UP_Click(Sender: TObject);
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

procedure TfrmLD6I02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD6I02.UP_Exit(Sender: TObject);
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

procedure TfrmLD6I02.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      mskDate.Color        := $00E6FFFE;
      mskDate.Enabled      := True;
      btnDate.Enabled      := True;
      mskFrom.Enabled      := True;
      mskTo.Enabled        := True;
      mskJumin.Color       := clWindow;
      mskJumin.Enabled     := False;
      btnJumin.Enabled     := False;
      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled :=False;

      mskDate.SetFocus;
      // 찾기 Enable
      edtFind.Enabled  := True;
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
      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled         :=False;


      mskJumin.SetFocus;

      // 찾기 Disable
      edtFind.Enabled  := False;
   end
   else if Sender = rbBarcode then
   begin
      mskDate.Color        := clWindow;
      mskDate.Enabled      := False;
      btnDate.Enabled      := False;
      mskFrom.Enabled      := False;
      mskTo.Enabled        := False;
      mskJumin.Color       := clWindow;
      mskJumin.Enabled     := False;
      btnJumin.Enabled     := False;
      MEdt_Barcode.Color   := $00E6FFFE;
      MEdt_Barcode.Enabled         := True;
     

      MEdt_Barcode.SetFocus;

      // 찾기 Disable
      edtFind.Enabled  := False;
   end;;
end;

procedure TfrmLD6I02.edtFindExit(Sender: TObject);
begin
   inherited;

   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0);
end;

procedure TfrmLD6I02.grdResultCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := UV_vSeq[DataRow-1];         // No
      2 : Value := UV_vDesc_hm[DataRow-1];     // 항목명
      3 : Value := UV_vIms_low[DataRow-1];     // 임상참고치(low)
      4 : Value := UV_vIms_high[DataRow-1];    // 임상참고치(high)
      5 : Value := UV_vDesc_unit[DataRow-1];   // 단위
      6 : Value := UV_vPrev_rslt[DataRow-1];   // 종전결과
      7 : Value := UV_vCur_rslt[DataRow-1];    // 결과
      8 : Value := UV_vMin3[DataRow-1];
      9 : Value := UV_vMin2[DataRow-1];
     10 : Value := UV_vMin1[DataRow-1];
     11 : Value := UV_vJung[DataRow-1];
     12 : Value := UV_vMax1[DataRow-1];
     13 : Value := UV_vMax2[DataRow-1];
     14 : Value := UV_vMax3[DataRow-1];
   end;
end;

procedure TfrmLD6I02.grdSokun1CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := UV_vGubn_pan[DataRow-1];
      2 : Value := UV_vCod_sokun1[DataRow-1];
      3 : Value := UV_vDesc_sokun1[DataRow-1];
   end;
end;

procedure TfrmLD6I02.grdSokun1CellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
var sSokun, sPan : String;
begin
   inherited;

   // Edit한 현재의 자료를 Variant 변수에 할당
   case DataCol of
      1 : UV_vGubn_pan[DataRow-1] := grdSokun1.CurrentCell.Value;

      2 : begin
             UV_vCod_sokun1[DataRow-1] := grdSokun1.CurrentCell.Value;

             if (Length(UV_vCod_sokun1[DataRow-1]) = 4) and
                (UV_vCod_sokun1[DataRow-1] <> '9999') then
             begin
                if GF_HulSokun_P(UV_vCod_sokun1[DataRow-1], '2', sSokun, sPan) then
                begin
                   UV_vGubn_pan[DataRow-1] := sPan;
                   grdSokun1.Cell[1, DataRow] := UV_vGubn_pan[DataRow-1];

                   UV_vDesc_sokun1[DataRow-1] := sSokun;
                   grdSokun1.Cell[3, DataRow] := UV_vDesc_sokun1[DataRow-1];
                end
                else
                begin
                   UV_vGubn_pan[DataRow-1]    := '';
                   UV_vDesc_sokun1[DataRow-1] := '';
                   grdSokun1.Cell[1, DataRow] := '';
                   grdSokun1.Cell[2, DataRow] := '';
                   grdSokun1.Cell[3, DataRow] := '';
                end;
             end
             else
             begin
                UV_vGubn_pan[DataRow-1]    := '';
                UV_vDesc_sokun1[DataRow-1] := '';
                grdSokun1.Cell[1, DataRow] := '';
                grdSokun1.Cell[3, DataRow] := '';
             end;
          end;
      3 : UV_vDesc_sokun1[DataRow-1] := grdSokun1.CurrentCell.Value;
   end;

   // 수정된 Column을 표시
   grdSokun1.AssignCellFont(DataCol, DataRow);
   grdSokun1.CellFont[DataCol, DataRow].Color := clBlue;

   // 작업Flag Check
   if UV_vInfo_numb1[DataRow-1] = '' then
      UV_vInfo_numb1[DataRow-1] := '2';

   // 수정되었음을 표시 -> 내부적으로 사용하는 Edit Flag
   UV_bEdit := True;
end;

procedure TfrmLD6I02.grdSokun1KeyPress(Sender: TObject; var Key: Char);
begin
   inherited;

   if (Sender as TtsGrid).CurrentCell.DataCol = 1 then
      Key := GF_Upper(Key);
   if (Sender as TtsGrid).CurrentCell.DataCol = 2 then
      Key := GF_Upper(Key);
end;

procedure TfrmLD6I02.ckCheckClick(Sender: TObject);
begin
   inherited;

   if Sender = ckCheck then
   begin
      if ckCheck.Checked = False then
         ck6Check.Checked := False;
   end;
   if Sender = ck6Check then
   begin
      if ck6Check.Checked then
         ckCheck.Checked := True;
   end;

   // 결과자료를 Query
   UP_ResultQuery;
end;

procedure TfrmLD6I02.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F3 then
   begin
      if edtFind.Enabled then edtFind.SetFocus;
   end
   else if Key = VK_F5  then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6  then UP_MoveNum(btnNextNum)
   else if Key = VK_F9  then
   begin
      pgcSokun.ActivePage := TabSheet1;
      grdSokun1.SetFocus;
   end
   else if Key = VK_F11 then
   begin
      pgcSokun.ActivePage := TabSheet2;
      grdSokun2.SetFocus;
   end
   else if Key = VK_F12 then
   begin
      pgcSokun.ActivePage := TabSheet3;
      grdSokun3.SetFocus;
   end
   else if Key = VK_F8 then grdResult.SetFocus;
end;

procedure TfrmLD6I02.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD6I02.grdSokun2CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := UV_vCod_sokun2[DataRow-1];
      2 : Value := UV_vDesc_sokun2[DataRow-1];
   end;
end;

procedure TfrmLD6I02.grdSokun3CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := UV_vCod_sokun3[DataRow-1];
      2 : Value := UV_vDesc_sokun3[DataRow-1];
   end;
end;

procedure TfrmLD6I02.grdSokun2CellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
var sSokun : String;
begin
   inherited;

   // Edit한 현재의 자료를 Variant 변수에 할당
   case DataCol of
      1 : begin
             UV_vCod_sokun2[DataRow-1] := grdSokun2.CurrentCell.Value;

             if (Length(UV_vCod_sokun2[DataRow-1]) = 4) and
                (UV_vCod_sokun2[DataRow-1] <> '9999') then
             begin
                if GF_HulSokun(UV_vCod_sokun2[DataRow-1], '5', sSokun) then
                begin
                   UV_vDesc_sokun2[DataRow-1] := sSokun;
                   grdSokun2.Cell[2, DataRow] := UV_vDesc_sokun2[DataRow-1];
                end
                else
                begin
                   UV_vCod_sokun2[DataRow-1]  := '';
                   UV_vDesc_sokun2[DataRow-1] := '';
                   grdSokun2.Cell[1, DataRow] := '';
                   grdSokun2.Cell[2, DataRow] := '';
                end;
             end
             else
             begin
                UV_vDesc_sokun2[DataRow-1] := '';
                grdSokun2.Cell[2, DataRow] := '';
             end;
          end;
      2 : UV_vDesc_sokun2[DataRow-1] := grdSokun2.CurrentCell.Value;
   end;

   // 수정된 Column을 표시
   grdSokun2.AssignCellFont(DataCol, DataRow);
   grdSokun2.CellFont[DataCol, DataRow].Color := clBlue;

   // 작업Flag Check
   if UV_vInfo_numb2[DataRow-1] = '' then
      UV_vInfo_numb2[DataRow-1] := '2';

   // 수정되었음을 표시 -> 내부적으로 사용하는 Edit Flag
   UV_bEdit := True;
end;

procedure TfrmLD6I02.grdSokun3CellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
var sSokun : String;
begin
   inherited;


   // Edit한 현재의 자료를 Variant 변수에 할당
   case DataCol of
      1 : begin
             UV_vCod_sokun3[DataRow-1] := grdSokun3.CurrentCell.Value;

             if (Length(UV_vCod_sokun3[DataRow-1]) = 4) and
                (UV_vCod_sokun3[DataRow-1] <> '9999') then
             begin
                if GF_HulSokun(UV_vCod_sokun3[DataRow-1], '6', sSokun) then
                begin
                   UV_vDesc_sokun3[DataRow-1] := sSokun;
                   grdSokun3.Cell[2, DataRow] := UV_vDesc_sokun3[DataRow-1];
                end
                else
                begin
                   UV_vCod_sokun3[DataRow-1]  := '';
                   UV_vDesc_sokun3[DataRow-1] := '';
                   grdSokun3.Cell[1, DataRow] := '';
                   grdSokun3.Cell[2, DataRow] := '';
                end;
             end
             else
             begin
                UV_vDesc_sokun3[DataRow-1] := '';
                grdSokun3.Cell[2, DataRow] := '';
             end;
          end;
      2 : UV_vDesc_sokun3[DataRow-1] := grdSokun3.CurrentCell.Value;
   end;

   // 수정된 Column을 표시
   grdSokun3.AssignCellFont(DataCol, DataRow);
   grdSokun3.CellFont[DataCol, DataRow].Color := clBlue;

   // 작업Flag Check
   if UV_vInfo_numb3[DataRow-1] = '' then
      UV_vInfo_numb3[DataRow-1] := '2';

   // 수정되었음을 표시 -> 내부적으로 사용하는 Edit Flag
   UV_bEdit := True;
end;

function  TfrmLD6I02.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD6I02.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD6I021 := TfrmLD6I021.Create(Self);
     frmLD6I021.QuickRep.Preview;
  end
  else
  begin
     frmLD6I021 := TfrmLD6I021.Create(Self);
     frmLD6I021.QuickRep.Print;
  end;

end;


procedure TfrmLD6I02.SBut_ExcelClick(Sender: TObject);
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

   Gauge2.MaxValue := grdResult.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

     //Data import
      ArrV3 := VarArrayCreate([0, grdResult.Rows + 2, 0, grdResult.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdResult.Rows + 2 do
      begin
         if Row = 0 then
         begin
               ArrV3[Row, 0] := '분주번호:'+ edtNUM_BUNJU.text;
               ArrV3[Row, 1] := '분주일:'+ mskDAT_BUNJU.text;
               ArrV3[Row, 2] := '성 명 :'+ edtDESC_NAME.text;
         end
         else if Row = 1 then
         begin
            for Col := 0 to 5 do
               ArrV3[Row, Col] := grdResult.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to 5 do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdResult.cell[Col + 1, Row-1];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdResult.Rows+2, grdResult.Cols]].Value := ArrV3;

      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.


