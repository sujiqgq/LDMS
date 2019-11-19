//==============================================================================
// 수정일자      : 2013.5.9
// 수정자        : 김희석
// 수정내용      : 키, 몸무게,임신및 생리
//=============================================================================
// 수정일자      : 2013.6.15
// 수정자        : 김승철
// 수정내용      : 임신중 안나오는 부분 수정
//=============================================================================
// 수정일자      : 2014.3.7
// 수정자        : 김희석
// 수정내용      : 검진일 3월8일부터 막음 
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

unit LD1I17;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I17 = class(TfrmSingle)
    qryGlkwa: TQuery;
    grdMaster: TtsGrid;
    qryU_Glkwa: TQuery;
    pnlDetail: TPanel;
    edtValue1: TEdit;
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
    mskNum: TMaskEdit;
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
    RG_sort: TRadioGroup;
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
    MEdt_SampS: TMaskEdit;
    Label15: TLabel;
    MEdt_SampE: TMaskEdit;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
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
    pnlName14: TPanel;
    Panel4: TPanel;
    edtPValue14: TEdit;
    edtValue14: TEdit;
    Panel8: TPanel;
    pnlName11: TPanel;
    edtPValue11: TEdit;
    edtValue11: TEdit;
    Panel10: TPanel;
    pnlName15: TPanel;
    edtPValue15: TEdit;
    edtValue15: TEdit;
    Panel12: TPanel;
    pnlName12: TPanel;
    edtPValue12: TEdit;
    edtValue12: TEdit;
    Panel14: TPanel;
    pnlName13: TPanel;
    edtPValue13: TEdit;
    edtValue13: TEdit;
    Panel16: TPanel;
    pnlName16: TPanel;
    edtPValue16: TEdit;
    edtValue16: TEdit;
    Panel18: TPanel;
    pnlName17: TPanel;
    edtPValue17: TEdit;
    edtValue17: TEdit;
    edtPValue18: TEdit;
    pnlName18: TPanel;
    Panel21: TPanel;
    pnlName19: TPanel;
    Panel23: TPanel;
    edtPValue19: TEdit;
    edtValue19: TEdit;
    Panel24: TPanel;
    edtPValue20: TEdit;
    edtValue20: TEdit;
    pnlName20: TPanel;
    Panel26: TPanel;
    edtPValue21: TEdit;
    edtValue21: TEdit;
    pnlName21: TPanel;
    pnlName22: TPanel;
    Panel32: TPanel;
    edtPValue22: TEdit;
    edtValue22: TEdit;
    Panel35: TPanel;
    pnlName23: TPanel;
    edtPValue23: TEdit;
    edtValue23: TEdit;
    Panel2: TPanel;
    pnlName25: TPanel;
    edtPValue25: TEdit;
    edtValue25: TEdit;
    Panel11: TPanel;
    pnlName24: TPanel;
    edtPValue24: TEdit;
    edtValue24: TEdit;
    Panel15: TPanel;
    pnlName26: TPanel;
    edtPValue26: TEdit;
    edtValue26: TEdit;
    edtValue18: TEdit;
    edtValue27: TEdit;
    edtPValue27: TEdit;
    Panel9: TPanel;
    Panel13: TPanel;
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
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_jisa, UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vNum_samp,
    UV_vCod_hul,    UV_vCod_cancer, UV_vCod_jangbi, UV_vCod_chuga,
    UV_vGubn_nosin, UV_vGubn_nsyh, UV_vGubn_adult, UV_vGubn_adyh,
    UV_vGubn_life,  UV_vGubn_lfyh, UV_vGubn_bogun, UV_vGubn_bgyh,
    UV_vGubn_agab,  UV_vGubn_agyh, UV_vGubn_tkgm,
    UV_vCod_bjjs, UV_vDat_bunju, UV_vNum_bunju, UV_vDesc_glkwa, UV_vDesc_Pglkwa, UV_vPart01,

    UV_vC001,UV_vC002,UV_vC003,UV_vC004,UV_vC005,UV_vC006,UV_vC007,UV_vC009,UV_vC010,UV_vC011,UV_vC012,UV_vC013,UV_vC017,UV_vC019,
    UV_vC021,UV_vC025,UV_vC026,UV_vC027,UV_vC028,UV_vC029,UV_vC032,UV_vC033,UV_vC041,UV_vC042,UV_vC043,UV_vC074,UV_vC077 ,UV_vGumgin_check : Variant;

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
  frmLD1I17: TfrmLD1I17;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

function  TfrmLD1I17.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I17.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD1I17.UP_GridSet;
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

procedure TfrmLD1I17.UP_VarMemSet(iLength : Integer);
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

      UV_vC001         := VarArrayCreate([0, 0], varOleStr);
      UV_vC002         := VarArrayCreate([0, 0], varOleStr);
      UV_vC003         := VarArrayCreate([0, 0], varOleStr);
      UV_vC004         := VarArrayCreate([0, 0], varOleStr);
      UV_vC005         := VarArrayCreate([0, 0], varOleStr);
      UV_vC006         := VarArrayCreate([0, 0], varOleStr);
      UV_vC007         := VarArrayCreate([0, 0], varOleStr);
      UV_vC009         := VarArrayCreate([0, 0], varOleStr);
      UV_vC010         := VarArrayCreate([0, 0], varOleStr);
      UV_vC011         := VarArrayCreate([0, 0], varOleStr);
      UV_vC012         := VarArrayCreate([0, 0], varOleStr);
      UV_vC013         := VarArrayCreate([0, 0], varOleStr);
      UV_vC017         := VarArrayCreate([0, 0], varOleStr);
      UV_vC019         := VarArrayCreate([0, 0], varOleStr);
      UV_vC021         := VarArrayCreate([0, 0], varOleStr);
      UV_vC025         := VarArrayCreate([0, 0], varOleStr);
      UV_vC026         := VarArrayCreate([0, 0], varOleStr);
      UV_vC027         := VarArrayCreate([0, 0], varOleStr);
      UV_vC028         := VarArrayCreate([0, 0], varOleStr);
      UV_vC029         := VarArrayCreate([0, 0], varOleStr);
      UV_vC032         := VarArrayCreate([0, 0], varOleStr);
      UV_vC033         := VarArrayCreate([0, 0], varOleStr);
      UV_vC041         := VarArrayCreate([0, 0], varOleStr);
      UV_vC042         := VarArrayCreate([0, 0], varOleStr);
      UV_vC043         := VarArrayCreate([0, 0], varOleStr);
      UV_vC074         := VarArrayCreate([0, 0], varOleStr);
      UV_vC077         := VarArrayCreate([0, 0], varOleStr);
      UV_vPart01       := VarArrayCreate([0, 0], varOleStr);
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

      VarArrayRedim(UV_vC001,        iLength);
      VarArrayRedim(UV_vC002,        iLength);
      VarArrayRedim(UV_vC003,        iLength);
      VarArrayRedim(UV_vC004,        iLength);
      VarArrayRedim(UV_vC005,        iLength);
      VarArrayRedim(UV_vC006,        iLength);
      VarArrayRedim(UV_vC007,        iLength);
      VarArrayRedim(UV_vC009,        iLength);
      VarArrayRedim(UV_vC010,        iLength);
      VarArrayRedim(UV_vC011,        iLength);
      VarArrayRedim(UV_vC012,        iLength);
      VarArrayRedim(UV_vC013,        iLength);
      VarArrayRedim(UV_vC017,        iLength);
      VarArrayRedim(UV_vC019,        iLength);
      VarArrayRedim(UV_vC021,        iLength);
      VarArrayRedim(UV_vC025,        iLength);
      VarArrayRedim(UV_vC026,        iLength);
      VarArrayRedim(UV_vC027,        iLength);
      VarArrayRedim(UV_vC028,        iLength);
      VarArrayRedim(UV_vC029,        iLength);
      VarArrayRedim(UV_vC032,        iLength);
      VarArrayRedim(UV_vC033,        iLength);
      VarArrayRedim(UV_vC041,        iLength);
      VarArrayRedim(UV_vC042,        iLength);
      VarArrayRedim(UV_vC043,        iLength);
      VarArrayRedim(UV_vC074,        iLength);
      VarArrayRedim(UV_vC077,        iLength);
      VarArrayRedim(UV_vPart01,      iLength);
      VarArrayRedim(UV_vGumgin_check, iLength);
   end;
end;

function TfrmLD1I17.UF_Condition : Boolean;
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

function TfrmLD1I17.UF_Save : Boolean;
var i, iTemp : Integer;
    Part01 , Part02 : String;
begin
   Result := False;

   // 작업중인지 Check
   if not UV_bEdit then exit;

   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   with qryPart01 do
   begin
      Close;
      qryPart01.ParamByName('dat_bunju').AsString := mskDAT_BUNJU.Text;
      qryPart01.ParamByName('num_bunju').AsString := edtNUM_BUNJU.Text;
      qryPart01.ParamByName('cod_bjjs').AsString  := UV_vCod_bjjs[i];

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
   // C001	총단백	1	6
   Part01 := GF_RPad(edtValue1.text, 6, ' ')+ Copy(Part01, 7, Length(Part01));
   //C002	알부민	7	12
   Part01 := Copy(Part01, 1,  6)  + GF_RPad(edtValue2.text, 6, ' ') + Copy(Part01, 13, Length(Part01));
   //C003(계산)	글로부린	13	18
   Part01 := Copy(Part01, 1, 12)  + GF_RPad(edtValue3.text, 6, ' ') + Copy(Part01, 19, Length(Part01));
   //C004(계산)	A/G 비율	19	24
   Part01 := Copy(Part01, 1, 18)  + GF_RPad(edtValue4.text, 6, ' ') + Copy(Part01, 25, Length(Part01));
   //C005	총빌리루빈	25	30
   Part01 := Copy(Part01, 1, 24)  + GF_RPad(edtValue5.text, 6, ' ') + Copy(Part01, 31, Length(Part01));
   //C006	직접빌리루빈	31	36
   Part01 := Copy(Part01, 1, 30)  + GF_RPad(edtValue6.text, 6, ' ') + Copy(Part01, 37, Length(Part01));
   //C007(계산)	간접빌리루빈	37	42
   Part01 := Copy(Part01, 1, 36)  + GF_RPad(edtValue7.text, 6, ' ') + Copy(Part01, 43, Length(Part01));


   // C009	AST(SGOT)	49	54
   Part01 := Copy(Part01, 1, 48)  + GF_RPad(edtValue8.text, 6, ' ') + Copy(Part01, 55, Length(Part01));
   //C010	ALT(SGPT)	55	60
   Part01 := Copy(Part01, 1, 54)  + GF_RPad(edtValue9.text, 6, ' ') + Copy(Part01, 61, Length(Part01));
   //C011	r-GTP   	61	66
   Part01 := Copy(Part01, 1, 60)  + GF_RPad(edtValue10.text, 6, ' ') + Copy(Part01, 67, Length(Part01));

   //C012	LAP   	        67	72
   Part01 := Copy(Part01, 1, 66)  + GF_RPad(edtValue11.text, 6, ' ') + Copy(Part01, 73, Length(Part01));
   //C013	ALP	        73	78
   Part01 := Copy(Part01, 1, 72)  + GF_RPad(edtValue12.text, 6, ' ') + Copy(Part01, 79, Length(Part01));
   //C017	뇨산	        85	90
   Part01 := Copy(Part01, 1, 84)  + GF_RPad(edtValue13.text, 6, ' ') + Copy(Part01, 91, Length(Part01));
   //C019	CPK	91	96
   Part01 := Copy(Part01, 1, 90)  + GF_RPad(edtValue14.text, 6, ' ') + Copy(Part01, 97, Length(Part01));
   //C021	LDH	97	102
   Part01 := Copy(Part01, 1, 96)  + GF_RPad(edtValue15.text, 6, ' ') + Copy(Part01, 103, Length(Part01));




   //C025	콜레스테롤	121	126
   Part01 := Copy(Part01, 1, 120)  + GF_RPad(edtValue16.text, 6, ' ') + Copy(Part01, 127, Length(Part01));
   //C026	HDL-콜레스테롤	127	132
   Part01 := Copy(Part01, 1, 126)  + GF_RPad(edtValue17.text, 6, ' ') + Copy(Part01, 133, Length(Part01));
   //C027(계산)	LDL-콜레스테롤	133	138
   Part01 := Copy(Part01, 1, 132)  + GF_RPad(edtValue18.text, 6, ' ') + Copy(Part01, 139, Length(Part01));
   //C028	중성지방	139	144
   Part01 := Copy(Part01, 1, 138) + GF_RPad(edtValue19.text, 6, ' ') + Copy(Part01, 145, Length(Part01));
   //C029(계산)	cardiac risk index	151	156
   Part01 := Copy(Part01, 1, 150) + GF_RPad(edtValue20.text, 6, ' ') + Copy(Part01, 157, Length(Part01));
   //C032  공복시혈당	 	157	162
   Part01 := Copy(Part01, 1, 156) + GF_RPad(edtValue21.text, 6, ' ') + Copy(Part01, 163, Length(Part01));
   //C033  식후혈당	 	163	168
   Part01 := Copy(Part01, 1, 162) + GF_RPad(edtValue22.text, 6, ' ') + Copy(Part01, 169, Length(Part01));

   //C041	뇨소질소	181	186
   Part01 := Copy(Part01, 1, 180) + GF_RPad(edtValue23.text, 6, ' ') + Copy(Part01, 187, Length(Part01));
   //C042	  크레아티닌 	187	192
   Part01 := Copy(Part01, 1, 186) + GF_RPad(edtValue24.text, 6, ' ') + Copy(Part01, 193, Length(Part01));
   //C043(계산)	BUN/Cr 비율	193	198
   Part01 := Copy(Part01, 1, 192) + GF_RPad(edtValue25.text, 6, ' ') + Copy(Part01, 199, Length(Part01));

   //LDL-콜레스테롤(효소) C074          01      56      331     336
   Part02 := Copy(Part01 + Part02, 1, 330) + GF_RPad(edtValue26.text, 6, ' ') + Copy(Part01 + Part02, 337, Length(Part01 + Part02));  //20140620 곽윤설
   //식전 자가혈당측정기C077	01	59 	349	354
   Part02 := Copy(Part02, 1, 348) + GF_RPad(edtValue27.text, 6, ' ') + Copy(Part02, 355, Length(Part02));

   Part01 := 'A' + Part01;
   Part01 := Trim(Part01);
   Part01 := Copy(Part01, 2, Length(Part01)-1);

   Part02 := 'A' + Part02;
   Part02 := Trim(Part02);
   Part02 := Copy(Part02, 202, Length(Part02)-1);

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
      UV_vC001[i]   := Trim(copy(Part01,    1, 6));
      UV_vC002[i]   := Trim(copy(Part01,    7, 6));
      UV_vC003[i]   := Trim(copy(Part01,   13, 6));
      UV_vC004[i]   := Trim(copy(Part01,   19, 6));
      UV_vC005[i]   := Trim(copy(Part01,   25, 6));
      UV_vC006[i]   := Trim(copy(Part01,   31, 6));
      UV_vC007[i]   := Trim(copy(Part01,   37, 6));
      UV_vC009[i]   := Trim(copy(Part01,   49, 6));
      UV_vC010[i]   := Trim(copy(Part01,   55, 6));
      UV_vC011[i]   := Trim(copy(Part01,   61, 6));
      UV_vC012[i]   := Trim(copy(Part01,   67, 6));
      UV_vC013[i]   := Trim(copy(Part01,   73, 6));
      UV_vC017[i]   := Trim(copy(Part01,   85, 6));
      UV_vC019[i]   := Trim(copy(Part01,   91, 6));
      UV_vC021[i]   := Trim(copy(Part01,   97, 6));
      UV_vC025[i]   := Trim(copy(Part01,  121, 6));
      UV_vC026[i]   := Trim(copy(Part01,  127, 6));
      UV_vC027[i]   := Trim(copy(Part01,  133, 6));
      UV_vC028[i]   := Trim(copy(Part01,  139, 6));
      UV_vC029[i]   := Trim(copy(Part01,  151, 6));
      UV_vC032[i]   := Trim(copy(Part01,  157, 6));
      UV_vC033[i]   := Trim(copy(Part01,  163, 6));
      UV_vC041[i]   := Trim(copy(Part01,  181, 6));
      UV_vC042[i]   := Trim(copy(Part01,  187, 6));
      UV_vC043[i]   := Trim(copy(Part01,  193, 6));
      UV_vC074[i]   := Trim(copy(Part01,  193, 6));
      UV_vC077[i]   := Trim(copy(Part02,  149, 6));    //200뺀값으로 위치값 넣음.
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD1I17.btnSaveClick(Sender: TObject);
begin
   inherited;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I17.FormCreate(Sender: TObject);
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
  
end;

procedure TfrmLD1I17.grdMasterCellLoaded(Sender: TObject; DataCol,
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
   if  mskDAT_GMGN.text<='20140307' then    btnPSave.Enabled:=true
   else btnPSave.Enabled:=false;
         //3월8일접수분 부터 막음
end;

procedure TfrmLD1I17.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy : String;
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
      if GV_sJisa = '00' then
         sWhere := ' B.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''''
      else
         sWhere := ' B.cod_jisa = ''' + Copy(GV_sUserId, 1, 2) + '''';

      //연구소+센터만 조회
      sWhere := sWhere + ' and B.gubn_hulgum = ''3'' ';
      {  //20141117 곽윤설
      sWhere := sWhere + ' and substring(B.cod_jangbi,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(B.cod_jangbi,1,2) <> ''GS''';
      sWhere := sWhere + ' and substring(B.cod_hul,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(B.cod_hul,1,2) <> ''GS''';
      }
      sWhere := sWhere + '  and ((B.gubn_nosin = ''1'' or B.gubn_adult = ''1'' or B.gubn_life = ''1'' or B.cod_chuga like ''%C077%'') or  ' +
                         '       ((B.gubn_nosin = ''2'' or B.gubn_adult = ''2'') and B.cod_chuga like ''%C032%'') or ' +
                         '       (B.gubn_tkgm = ''2'')) ';

      if rbDate.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND A.dat_bunju = ''' + mskDate.Text + '''';

         if Cmb_gubn.Text = '분주번호' then
         begin
            if Trim(mskFrom.Text) <> '' then  sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
            if Trim(mskTo.Text) <> '' then    sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
         end
         else if Cmb_gubn.Text = '샘플번호' then
         begin
            if Trim(MEdt_SampS.Text) <> '' then sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
            if Trim(MEdt_SampE.Text) <> '' then sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
         end
         else
         begin
            if Trim(MEdt_SampS.Text) <> '' then sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
            if Trim(MEdt_SampE.Text) <> '' then sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
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

      ParamByName('sWhere').Asmemo     :=  sWhere;
      ParamByName('sOrderBy').AsString :=  sOrderBy;
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I10', 'Q', 'Y');
         UP_VarMemSet(RecordCount - 1);

         while not qryGlkwa.Eof do
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


            UV_vC001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 1,6));
            UV_vC002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 7,6));
            UV_vC003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 13,6));
            UV_vC004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 19,6));
            UV_vC005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 25,6));
            UV_vC006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 31,6));
            UV_vC007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 37,6));
            UV_vC009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 49,6));
            UV_vC010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 55,6));
            UV_vC011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 61,6));
            UV_vC012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 67,6));
            UV_vC013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 73,6));
            UV_vC017[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 85,6));
            UV_vC019[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 91,6));
            UV_vC021[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 97,6));
            UV_vC025[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 121,6));
            UV_vC026[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 127,6));
            UV_vC027[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 133,6));
            UV_vC028[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 139,6));
            UV_vC029[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 151,6));
            UV_vC032[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 157,6));
            UV_vC033[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 163,6));
            UV_vC041[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 181,6));
            UV_vC042[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 187,6));
            UV_vC043[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 193,6));
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
   if (Trim(TEdit(Sender).Text) = '')  and (TEdit(Sender).Color <> clBtnFace)then
   begin
      TEdit(Sender).Color     := $0088FF88;
   end;



   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD1I17.btnCancelClick(Sender: TObject);
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

procedure TfrmLD1I17.UP_Change(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    //eRslt, e25, e26, e28 : Extended;

    e1, e2, e3, e5, e6, e22, e23, e25, e26, e28, e41, e42, e45, e48 ,eRslt: Extended;
    cod_HM  : array[1..27] of String;
begin
   inherited;
   e1 :=0;
   e2 :=0;
   e3 :=0;
   e5 :=0;
   e6 :=0;
   e22:=0;
   e23:=0;
   e25:=0;
   e26:=0;
   e28:=0;
   e41:=0;
   e42:=0;
   e45:=0;
   e48:=0;
   eRslt:=0;


   cod_HM[1] := 'C001';
   cod_HM[2] := 'C002';
   cod_HM[3] := 'C003';
   cod_HM[4] := 'C004';
   cod_HM[5] := 'C005';
   cod_HM[6] := 'C006';
   cod_HM[7] := 'C007';
   cod_HM[8] := 'C009';
   cod_HM[9] := 'C010';
   cod_HM[10] := 'C011';
   cod_HM[11] := 'C012';
   cod_HM[12] := 'C013';
   cod_HM[13] := 'C017';
   cod_HM[14] := 'C019';
   cod_HM[15] := 'C021';
   cod_HM[16] := 'C025';
   cod_HM[17] := 'C026';
   cod_HM[18] := 'C027';
   cod_HM[19] := 'C028';
   cod_HM[20] := 'C029';
   cod_HM[21] := 'C032';
   cod_HM[22] := 'C033';
   cod_HM[23] := 'C041';
   cod_HM[24] := 'C042';
   cod_HM[25] := 'C043';
   cod_HM[26] := 'C074';
   cod_HM[27] := 'C077';
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
   if (Trim(TEdit(Sender).Text) = '')  and (TEdit(Sender).Color <> clBtnFace)then
   begin
      TEdit(Sender).Color     := $0088FF88;
   end;


    eRslt := 0;
    if ((Sender = edtValue1) or (Sender = edtValue2)) and  (edtValue3.Enabled=True)  then
    begin
   //     Inc(iCntCheck);
         if  (edtValue1.Text <> '') and (edtValue2.Text <> '')  then
         begin
             e1 := GF_StrToFloat(edtValue1.Text);
             e2 := GF_StrToFloat(edtValue2.Text);
             if e1>e2 then  eRslt := e1-e2;
             if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue3.Text := sValue;
         end;
    end;
       eRslt := 0;

   if ((Sender = edtValue1) or (Sender = edtValue2)) and   (edtValue4.Enabled=True)  then
    begin
      //  Inc(iCntCheck);
         if  (edtValue1.Text <> '') and (edtValue2.Text <> '') then
         begin
             e1 := GF_StrToFloat(edtValue1.Text);
             e2 := GF_StrToFloat(edtValue2.Text);
             if e1>e2 then   eRslt :=e2/(e1-e2);
             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue4.Text := sValue;
         end;
    end;
    eRslt := 0;

    if ((Sender = edtValue5) or (Sender = edtValue6)) and   (edtValue7.Enabled=True) then
    begin
      //  Inc(iCntCheck);
         if  (edtValue5.Text <> '') and (edtValue6.Text <> '')  then
         begin
             e5 := GF_StrToFloat(edtValue5.Text);
             e6 := GF_StrToFloat(edtValue6.Text);
             if e5>e6 then  eRslt := e5-e6;
             if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue7.Text := sValue;
         end;
    end;
//
    eRslt := 0;

    if ((Sender = edtValue16) or (Sender = edtValue17) or (Sender = edtValue19)) and  (edtValue18.Enabled=True) then
    begin
        Inc(iCntCheck);
         if  (edtValue16.Text <> '') and (edtValue17.Text <> '') and  (edtValue19.Text <> '') and  (iCntCheck > 3)   then
         begin
             e25 := GF_StrToFloat(edtValue16.Text); //   c025
             e26 := GF_StrToFloat(edtValue17.Text); //   c026
             e28 := GF_StrToFloat(edtValue19.Text); //   C028

             if (e25>0) and (e26>0) and (e28>0) then  eRslt := e25 - ((e28/5) + e26);
             if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue18.Text := sValue;
         end;
    end;
    eRslt := 0;

    if ((Sender = edtValue16) or (Sender = edtValue17)) and (edtValue20.Enabled=True)  then
    begin
      //   Inc(iCntCheck);
         if  (edtValue16.Text <> '') and (edtValue17.Text <> '')  then
         begin
             e25 := GF_StrToFloat(edtValue16.Text);
             e26 := GF_StrToFloat(edtValue17.Text);
             if (e25>0) and (e26>0) then   eRslt := (e26 / e25) * 100;
             if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue20.Text := sValue;
         end;
    end;
    eRslt := 0;

    if ((Sender = edtValue23) or (Sender = edtValue24)) and (edtValue25.Enabled=True) then
    begin
        // Inc(iCntCheck);
         if  (edtValue23.Text <> '') and (edtValue24.Text <> '')  then
         begin
             e41 := GF_StrToFloat(edtValue23.Text);
             e42 := GF_StrToFloat(edtValue24.Text);
             if (e41>0) and (e42>0) then  eRslt := e41 / e42;
             if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

             valChange.Scale := 1;
             valChange.Value := eRslt;
             sValue := valChange.Text;

             edtValue25.Text := sValue;
         end;
    end; 


   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I17.UP_DChange(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    cod_HM  : array[1..27] of String;
begin
   inherited;

   cod_HM[1] := 'C001';
   cod_HM[2] := 'C002';
   cod_HM[3] := 'C003';
   cod_HM[4] := 'C004';
   cod_HM[5] := 'C005';
   cod_HM[6] := 'C006';
   cod_HM[7] := 'C007';
   cod_HM[8] := 'C009';
   cod_HM[9] := 'C010';
   cod_HM[10] := 'C011';
   cod_HM[11] := 'C012';
   cod_HM[12] := 'C013';
   cod_HM[13] := 'C017';
   cod_HM[14] := 'C019';
   cod_HM[15] := 'C021';
   cod_HM[16] := 'C025';
   cod_HM[17] := 'C026';
   cod_HM[18] := 'C027';
   cod_HM[19] := 'C028';
   cod_HM[20] := 'C029';
   cod_HM[21] := 'C032';
   cod_HM[22] := 'C033';
   cod_HM[23] := 'C041';
   cod_HM[24] := 'C042';
   cod_HM[25] := 'C043';
   cod_HM[26] := 'C074';
   cod_HM[27] := 'C077';

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

procedure TfrmLD1I17.grdMasterRowChanged(Sender: TObject; OldRow,
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

//[20120516-유동구]특수검진은 제외.(연구소+센터인 경우)

   if (UV_vGubn_tkgm[NewRow-1] = '1') or (UV_vGubn_tkgm[NewRow-1] = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
      qryTkgum.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryTkgum.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
//[2012.03.16]특검항목 자유롭게..
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

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         sHmList_01 := sHmList_01 + vCod_chuga[i];
      end;
   end;

   //--------------------------------------------------------------------------
   if (UV_vGubn_nosin[NewRow-1] = '2') or (UV_vGubn_adult[NewRow-1] = '2') then
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
      edtValue13.Color := clBtnFace;
      edtValue14.Color := clBtnFace;
      edtValue15.Color := clBtnFace;
      edtValue16.Color := clBtnFace;
      edtValue17.Color := clBtnFace;
      edtValue18.Color := clBtnFace;
      edtValue19.Color := clBtnFace;
      edtValue20.Color := clBtnFace;
      if pos('C032',sHmList_01) > 0 then edtValue21.Color := clWindow
      else                               edtValue21.Color := clBtnFace;
      edtValue22.Color := clBtnFace;
      edtValue23.Color := clBtnFace;
      edtValue24.Color := clBtnFace;
      edtValue25.Color := clBtnFace;
      edtValue26.Color := clBtnFace;
      edtValue27.Color := clBtnFace;

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
      edtValue13.Enabled := False;
      edtValue14.Enabled := False;
      edtValue15.Enabled := False;
      edtValue16.Enabled := False;
      edtValue17.Enabled := False;
      edtValue18.Enabled := False;
      edtValue19.Enabled := False;
      edtValue20.Enabled := False;
      if pos('C032',sHmList_01) > 0 then edtValue21.Enabled := True
      else                               edtValue21.Enabled := False;
      edtValue22.Enabled := False;
      edtValue23.Enabled := False;
      edtValue24.Enabled := False;
      edtValue25.Enabled := False;
      edtValue26.Enabled := False;
      edtValue27.Enabled := False;
   end
   else
   begin
      if pos('C001',sHmList_01) > 0 then edtValue1.Color := clWindow
      else                               edtValue1.Color := clBtnFace;
      if pos('C002',sHmList_01) > 0 then edtValue2.Color := clWindow
      else                               edtValue2.Color := clBtnFace;
      if pos('C003',sHmList_01) > 0 then edtValue3.Color := clWindow
      else                               edtValue3.Color := clBtnFace;
      if pos('C004',sHmList_01) > 0 then edtValue4.Color := clWindow
      else                               edtValue4.Color := clBtnFace;
      if pos('C005',sHmList_01) > 0 then edtValue5.Color := clWindow
      else                               edtValue5.Color := clBtnFace;
      if pos('C006',sHmList_01) > 0 then edtValue6.Color := clWindow
      else                               edtValue6.Color := clBtnFace;
      if pos('C007',sHmList_01) > 0 then edtValue7.Color := clWindow
      else                               edtValue7.Color := clBtnFace;
      if pos('C009',sHmList_01) > 0 then edtValue8.Color := clWindow
      else                               edtValue8.Color := clBtnFace;
      if pos('C010',sHmList_01) > 0 then edtValue9.Color := clWindow
      else                               edtValue9.Color := clBtnFace;
      if pos('C011',sHmList_01) > 0 then edtValue10.Color := clWindow
      else                               edtValue10.Color := clBtnFace;

      if pos('C012',sHmList_01) > 0 then edtValue11.Color := clWindow
      else                               edtValue11.Color := clBtnFace;
      if pos('C013',sHmList_01) > 0 then edtValue12.Color := clWindow
      else                               edtValue12.Color := clBtnFace;
      if pos('C017',sHmList_01) > 0 then edtValue13.Color := clWindow
      else                               edtValue13.Color := clBtnFace;
      if pos('C019',sHmList_01) > 0 then edtValue14.Color := clWindow
      else                               edtValue14.Color := clBtnFace;
      if pos('C021',sHmList_01) > 0 then edtValue15.Color := clWindow
      else                               edtValue15.Color := clBtnFace;
      if pos('C025',sHmList_01) > 0 then edtValue16.Color := clWindow
      else                               edtValue16.Color := clBtnFace;
      if pos('C026',sHmList_01) > 0 then edtValue17.Color := clWindow
      else                               edtValue17.Color := clBtnFace;
      if pos('C027',sHmList_01) > 0 then edtValue18.Color := clWindow
      else                               edtValue18.Color := clBtnFace;
      if pos('C028',sHmList_01) > 0 then edtValue19.Color := clWindow
      else                               edtValue19.Color := clBtnFace;
      if pos('C029',sHmList_01) > 0 then edtValue20.Color := clWindow
      else                               edtValue20.Color := clBtnFace;

      if pos('C032',sHmList_01) > 0 then edtValue21.Color := clWindow
      else                               edtValue21.Color := clBtnFace;
      if pos('C033',sHmList_01) > 0 then edtValue22.Color := clWindow
      else                               edtValue22.Color := clBtnFace;
      if pos('C041',sHmList_01) > 0 then edtValue23.Color := clWindow
      else                               edtValue23.Color := clBtnFace;
      if pos('C042',sHmList_01) > 0 then edtValue24.Color := clWindow
      else                               edtValue24.Color := clBtnFace;
      if pos('C043',sHmList_01) > 0 then edtValue25.Color := clWindow
      else                               edtValue25.Color := clBtnFace;
      if pos('C074',sHmList_01) > 0 then edtValue26.Color := clWindow
      else                               edtValue26.Color := clBtnFace;
      if pos('C077',sHmList_01) > 0 then edtValue27.Color := clWindow
      else                               edtValue27.Color := clBtnFace;

      if pos('C001',sHmList_01) > 0 then edtValue1.Enabled := True
      else                               edtValue1.Enabled := False;
      if pos('C002',sHmList_01) > 0 then edtValue2.Enabled := True
      else                               edtValue2.Enabled := False;
      if pos('C003',sHmList_01) > 0 then edtValue3.Enabled := True
      else                               edtValue3.Enabled := False;
      if pos('C004',sHmList_01) > 0 then edtValue4.Enabled := True
      else                               edtValue4.Enabled := False;
      if pos('C005',sHmList_01) > 0 then edtValue5.Enabled := True
      else                               edtValue5.Enabled := False;
      if pos('C006',sHmList_01) > 0 then edtValue6.Enabled := True
      else                               edtValue6.Enabled := False;
      if pos('C007',sHmList_01) > 0 then edtValue7.Enabled := True
      else                               edtValue7.Enabled := False;
      if pos('C009',sHmList_01) > 0 then edtValue8.Enabled := True
      else                               edtValue8.Enabled := False;
      if pos('C010',sHmList_01) > 0 then edtValue9.Enabled := True
      else                               edtValue9.Enabled := False;
      if pos('C011',sHmList_01) > 0 then edtValue10.Enabled := True
      else                               edtValue10.Enabled := False;

      if pos('C012',sHmList_01) > 0 then edtValue11.Enabled := True
      else                               edtValue11.Enabled := False;
      if pos('C013',sHmList_01) > 0 then edtValue12.Enabled := True
      else                               edtValue12.Enabled := False;
      if pos('C017',sHmList_01) > 0 then edtValue13.Enabled := True
      else                               edtValue13.Enabled := False;
      if pos('C019',sHmList_01) > 0 then edtValue14.Enabled := True
      else                               edtValue14.Enabled := False;
      if pos('C021',sHmList_01) > 0 then edtValue15.Enabled := True
      else                               edtValue15.Enabled := False;
      if pos('C025',sHmList_01) > 0 then edtValue16.Enabled := True
      else                               edtValue16.Enabled := False;
      if pos('C026',sHmList_01) > 0 then edtValue17.Enabled := True
      else                               edtValue17.Enabled := False;
      if pos('C027',sHmList_01) > 0 then edtValue18.Enabled := True
      else                               edtValue18.Enabled := False;
      if pos('C028',sHmList_01) > 0 then edtValue19.Enabled := True
      else                               edtValue19.Enabled := False;
      if pos('C029',sHmList_01) > 0 then edtValue20.Enabled := True
      else                               edtValue20.Enabled := False;

      if pos('C032',sHmList_01) > 0 then edtValue21.Enabled := True
      else                               edtValue21.Enabled := False;
      if pos('C033',sHmList_01) > 0 then edtValue22.Enabled := True
      else                               edtValue22.Enabled := False;
      if pos('C041',sHmList_01) > 0 then edtValue23.Enabled := True
      else                               edtValue23.Enabled := False;
      if pos('C042',sHmList_01) > 0 then edtValue24.Enabled := True
      else                               edtValue24.Enabled := False;
      if pos('C043',sHmList_01) > 0 then edtValue25.Enabled := True
      else                               edtValue25.Enabled := False;
      if pos('C074',sHmList_01) > 0 then edtValue26.Enabled := True
      else                               edtValue26.Enabled := False;
      if pos('C077',sHmList_01) > 0 then edtValue27.Enabled := True
      else                               edtValue27.Enabled := False;
   end;

   edtValue1.Text     := UV_vC001[NewRow-1];
   edtValue2.Text     := UV_vC002[NewRow-1];
   edtValue3.Text     := UV_vC003[NewRow-1];
   edtValue4.Text     := UV_vC004[NewRow-1];
   edtValue5.Text     := UV_vC005[NewRow-1];
   edtValue6.Text     := UV_vC006[NewRow-1];
   edtValue7.Text     := UV_vC007[NewRow-1];
   edtValue8.Text     := UV_vC009[NewRow-1];
   edtValue9.Text     := UV_vC010[NewRow-1];
   edtValue10.Text    := UV_vC011[NewRow-1];
   edtValue11.Text    := UV_vC012[NewRow-1];
   edtValue12.Text    := UV_vC013[NewRow-1];
   edtValue13.Text    := UV_vC017[NewRow-1];
   edtValue14.Text    := UV_vC019[NewRow-1];
   edtValue15.Text    := UV_vC021[NewRow-1];
   edtValue16.Text    := UV_vC025[NewRow-1];
   edtValue17.Text    := UV_vC026[NewRow-1];
   edtValue18.Text    := UV_vC027[NewRow-1];
   edtValue19.Text    := UV_vC028[NewRow-1];
   edtValue20.Text    := UV_vC029[NewRow-1];
   edtValue21.Text    := UV_vC032[NewRow-1];
   edtValue22.Text    := UV_vC033[NewRow-1];
   edtValue23.Text    := UV_vC041[NewRow-1];
   edtValue24.Text    := UV_vC042[NewRow-1];
   edtValue25.Text    := UV_vC043[NewRow-1];
   edtValue26.Text    := UV_vC074[NewRow-1];
   edtValue27.Text    := UV_vC077[NewRow-1];


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
   edtPValue11.Text  := '';
   edtPValue12.Text  := '';
   edtPValue13.Text  := '';
   edtPValue14.Text  := '';
   edtPValue15.Text  := '';
   edtPValue16.Text  := '';
   edtPValue17.Text  := '';
   edtPValue18.Text  := '';
   edtPValue19.Text  := '';
   edtPValue20.Text := '';
   edtPValue21.Text  := '';
   edtPValue22.Text  := '';
   edtPValue23.Text  := '';
   edtPValue24.Text  := '';
   edtPValue25.Text  := '';
   edtPValue26.Text  := '';
   edtPValue27.Text  := '';


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
                  //총단백(C001)
                   if Trim(Copy(UV_fGulkwa, 1,  6)) <> '' then
                      edtPValue1.Text := Trim(Copy(UV_fGulkwa, 1,  6))
                   else
                      edtPValue1.Text := '';
                   //알부민(C002)
                   if Trim(Copy(UV_fGulkwa, 7,  6)) <> '' then
                      edtPValue2.Text := Trim(Copy(UV_fGulkwa, 7,  6))
                   else
                      edtPValue2.Text := '';
                   //C003(계산)	글로부린
                   if Trim(Copy(UV_fGulkwa, 13,  6)) <> '' then
                      edtPValue3.Text := Trim(Copy(UV_fGulkwa, 13,  6))
                   else
                      edtPValue3.Text := '';
                   //C004(계산)	A/G 비율
                   if Trim(Copy(UV_fGulkwa, 19,  6)) <> '' then
                      edtPValue4.Text := Trim(Copy(UV_fGulkwa, 19,  6))
                   else
                      edtPValue4.Text := '';
                   //C005	총빌리루빈
                   if Trim(Copy(UV_fGulkwa, 25,  6)) <> '' then
                      edtPValue5.Text := Trim(Copy(UV_fGulkwa, 25,  6))
                   else
                      edtPValue5.Text := '';
                   //C006	직접빌리루빈
                   if Trim(Copy(UV_fGulkwa, 31,  6)) <> '' then
                      edtPValue6.Text := Trim(Copy(UV_fGulkwa, 31,  6))
                   else
                      edtPValue6.Text := '';
                   //C007             간접빌리루빈
                   if Trim(Copy(UV_fGulkwa, 37,  6)) <> '' then
                      edtPValue7.Text := Trim(Copy(UV_fGulkwa, 37,  6))
                   else
                      edtPValue7.Text := '';
                   //C009	AST(SGOT)
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
                      edtPValue8.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else
                      edtPValue8.Text := '';
                   //C010	ALT(SGPT)
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
                      edtPValue9.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else
                      edtPValue9.Text := '';
                   //C011	r-GTP
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
                      edtPValue10.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else
                      edtPValue10.Text := '';
                   //C012	LAP
                   if Trim(Copy(UV_fGulkwa, 67,  6)) <> '' then
                      edtPValue11.Text := Trim(Copy(UV_fGulkwa, 67,  6))
                   else
                      edtPValue11.Text := '';
                   //C013	ALP
                   if Trim(Copy(UV_fGulkwa, 73,  6)) <> '' then
                      edtPValue12.Text := Trim(Copy(UV_fGulkwa, 73,  6))
                   else
                      edtPValue12.Text := '';
                  //C017	뇨산
                   if Trim(Copy(UV_fGulkwa, 85,  6)) <> '' then
                      edtPValue13.Text := Trim(Copy(UV_fGulkwa, 85,  6))
                   else
                      edtPValue13.Text := '';
                  //C019	CPK
                   if Trim(Copy(UV_fGulkwa, 91,  6)) <> '' then
                      edtPValue14.Text := Trim(Copy(UV_fGulkwa, 91,  6))
                   else
                      edtPValue14.Text := '';
                  //C021	LDH
                   if Trim(Copy(UV_fGulkwa, 97,  6)) <> '' then
                      edtPValue15.Text := Trim(Copy(UV_fGulkwa, 97,  6))
                   else
                      edtPValue15.Text := '';
                   //C025	콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
                      edtPValue16.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else
                      edtPValue16.Text := '';
                  //C026	HDL-콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
                      edtPValue17.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else
                      edtPValue17.Text := '';
                  //C027(계산)	LDL-콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
                      edtPValue18.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else
                      edtPValue18.Text := '';
                   //C028	중성지방
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
                      edtPValue19.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else
                      edtPValue19.Text := '';
                   //C029(계산)	cardiac risk index
                   if Trim(Copy(UV_fGulkwa, 151,  6)) <> '' then
                      edtPValue20.Text := Trim(Copy(UV_fGulkwa, 151,  6))
                   else
                      edtPValue20.Text := '';
                   //C032	공복시혈당
                   if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
                      edtPValue21.Text := Trim(Copy(UV_fGulkwa, 157,  6))
                   else
                      edtPValue21.Text := '';
                   //C033	식후혈당
                   if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then
                      edtPValue22.Text := Trim(Copy(UV_fGulkwa, 163,  6))
                   else
                      edtPValue22.Text := '';
                   //C041	뇨소질소
                   if Trim(Copy(UV_fGulkwa, 181,  6)) <> '' then
                      edtPValue23.Text := Trim(Copy(UV_fGulkwa, 181,  6))
                   else
                      edtPValue23.Text := '';
                   //C042	크레아티닌
                   if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
                      edtPValue24.Text := Trim(Copy(UV_fGulkwa, 187,  6))
                   else
                      edtPValue24.Text := '';
                   //C043(계산)	BUN/Cr 비율
                   if Trim(Copy(UV_fGulkwa, 193,  6)) <> '' then
                      edtPValue25.Text := Trim(Copy(UV_fGulkwa, 193,  6))
                   else
                      edtPValue25.Text := '';
                   //LDL-콜레스테롤(효소)(C074)
                   if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue26.Text := Trim(Copy(UV_fGulkwa, 331,  6))
                   else                                          edtPValue26.Text := '';
                   //식전 자가혈당측정기(C077)
                   if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue27.Text := Trim(Copy(UV_fGulkwa, 349,  6))
                   else                                          edtPValue27.Text := '';
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
                     //총단백(C001)
                   if Trim(Copy(UV_fGulkwa, 1,  6)) <> '' then
                      edtPValue1.Text := Trim(Copy(UV_fGulkwa, 1,  6))
                   else
                      edtPValue1.Text := '';
                   //알부민(C002)
                   if Trim(Copy(UV_fGulkwa, 7,  6)) <> '' then
                      edtPValue2.Text := Trim(Copy(UV_fGulkwa, 7,  6))
                   else
                      edtPValue2.Text := '';
                   //C003(계산)	글로부린
                   if Trim(Copy(UV_fGulkwa, 13,  6)) <> '' then
                      edtPValue3.Text := Trim(Copy(UV_fGulkwa, 13,  6))
                   else
                      edtPValue3.Text := '';
                   //C004(계산)	A/G 비율
                   if Trim(Copy(UV_fGulkwa, 19,  6)) <> '' then
                      edtPValue4.Text := Trim(Copy(UV_fGulkwa, 19,  6))
                   else
                      edtPValue4.Text := '';
                   //C005	총빌리루빈
                   if Trim(Copy(UV_fGulkwa, 25,  6)) <> '' then
                      edtPValue5.Text := Trim(Copy(UV_fGulkwa, 25,  6))
                   else
                      edtPValue5.Text := '';
                   //C006	직접빌리루빈
                   if Trim(Copy(UV_fGulkwa, 31,  6)) <> '' then
                      edtPValue6.Text := Trim(Copy(UV_fGulkwa, 31,  6))
                   else
                      edtPValue6.Text := '';
                   //C007             간접빌리루빈
                   if Trim(Copy(UV_fGulkwa, 37,  6)) <> '' then
                      edtPValue7.Text := Trim(Copy(UV_fGulkwa, 37,  6))
                   else
                      edtPValue7.Text := '';
                   //C009	AST(SGOT)
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
                      edtPValue8.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else
                      edtPValue8.Text := '';
                   //C010	ALT(SGPT)
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
                      edtPValue9.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else
                      edtPValue9.Text := '';
                   //C011	r-GTP
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
                      edtPValue10.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else
                      edtPValue10.Text := '';
                   //C012	LAP
                   if Trim(Copy(UV_fGulkwa, 67,  6)) <> '' then
                      edtPValue11.Text := Trim(Copy(UV_fGulkwa, 67,  6))
                   else
                      edtPValue11.Text := '';
                   //C013	ALP
                   if Trim(Copy(UV_fGulkwa, 73,  6)) <> '' then
                      edtPValue12.Text := Trim(Copy(UV_fGulkwa, 73,  6))
                   else
                      edtPValue12.Text := '';
                  //C017	뇨산
                   if Trim(Copy(UV_fGulkwa, 85,  6)) <> '' then
                      edtPValue13.Text := Trim(Copy(UV_fGulkwa, 85,  6))
                   else
                      edtPValue13.Text := '';
                  //C019	CPK
                   if Trim(Copy(UV_fGulkwa, 91,  6)) <> '' then
                      edtPValue14.Text := Trim(Copy(UV_fGulkwa, 91,  6))
                   else
                      edtPValue14.Text := '';
                  //C021	LDH
                   if Trim(Copy(UV_fGulkwa, 97,  6)) <> '' then
                      edtPValue15.Text := Trim(Copy(UV_fGulkwa, 97,  6))
                   else
                      edtPValue15.Text := '';
                   //C025	콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
                      edtPValue16.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else
                      edtPValue16.Text := '';
                  //C026	HDL-콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
                      edtPValue17.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else
                      edtPValue17.Text := '';
                  //C027(계산)	LDL-콜레스테롤
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
                      edtPValue18.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else
                      edtPValue18.Text := '';
                   //C028	중성지방
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
                      edtPValue19.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else
                      edtPValue19.Text := '';
                   //C029(계산)	cardiac risk index
                   if Trim(Copy(UV_fGulkwa, 151,  6)) <> '' then
                      edtPValue20.Text := Trim(Copy(UV_fGulkwa, 151,  6))
                   else
                      edtPValue20.Text := '';
                   //C032	공복시혈당
                   if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
                      edtPValue21.Text := Trim(Copy(UV_fGulkwa, 157,  6))
                   else
                      edtPValue21.Text := '';
                   //C033	식후혈당
                   if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then
                      edtPValue22.Text := Trim(Copy(UV_fGulkwa, 163,  6))
                   else
                      edtPValue22.Text := '';
                   //C041	뇨소질소
                   if Trim(Copy(UV_fGulkwa, 181,  6)) <> '' then
                      edtPValue23.Text := Trim(Copy(UV_fGulkwa, 181,  6))
                   else
                      edtPValue23.Text := '';
                   //C042	크레아티닌
                   if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
                      edtPValue24.Text := Trim(Copy(UV_fGulkwa, 187,  6))
                   else
                      edtPValue24.Text := '';
                   //C043(계산)	BUN/Cr 비율
                   if Trim(Copy(UV_fGulkwa, 193,  6)) <> '' then
                      edtPValue25.Text := Trim(Copy(UV_fGulkwa, 193,  6))
                   else
                      edtPValue25.Text := '';
                   //LDL-콜레스테롤(효소)(C074)
                   if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue26.Text := Trim(Copy(UV_fGulkwa, 331,  6))
                   else                                          edtPValue26.Text := '';
                   //식전 자가혈당측정기(C077)
                   if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue27.Text := Trim(Copy(UV_fGulkwa, 349,  6))
                   else                                          edtPValue27.Text := '';

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

             //총단백(C001)
             if Trim(Copy(UV_fGulkwa, 1,  6)) <> '' then
                edtPValue1.Text := Trim(Copy(UV_fGulkwa, 1,  6))
             else
                edtPValue1.Text := '';
             //알부민(C002)
             if Trim(Copy(UV_fGulkwa, 7,  6)) <> '' then
                edtPValue2.Text := Trim(Copy(UV_fGulkwa, 7,  6))
             else
                edtPValue2.Text := '';
             //C003(계산)	글로부린
             if Trim(Copy(UV_fGulkwa, 13,  6)) <> '' then
                edtPValue3.Text := Trim(Copy(UV_fGulkwa, 13,  6))
             else
                edtPValue3.Text := '';
             //C004(계산)	A/G 비율
             if Trim(Copy(UV_fGulkwa, 19,  6)) <> '' then
                edtPValue4.Text := Trim(Copy(UV_fGulkwa, 19,  6))
             else
                edtPValue4.Text := '';
             //C005	총빌리루빈
             if Trim(Copy(UV_fGulkwa, 25,  6)) <> '' then
                edtPValue5.Text := Trim(Copy(UV_fGulkwa, 25,  6))
             else
                edtPValue5.Text := '';
             //C006	직접빌리루빈
             if Trim(Copy(UV_fGulkwa, 31,  6)) <> '' then
                edtPValue6.Text := Trim(Copy(UV_fGulkwa, 31,  6))
             else
                edtPValue6.Text := '';
             //C007             간접빌리루빈
             if Trim(Copy(UV_fGulkwa, 37,  6)) <> '' then
                edtPValue7.Text := Trim(Copy(UV_fGulkwa, 37,  6))
             else
                edtPValue7.Text := '';
             //C009	AST(SGOT)
             if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then
                edtPValue8.Text := Trim(Copy(UV_fGulkwa, 49,  6))
             else
                edtPValue8.Text := '';
             //C010	ALT(SGPT)
             if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then
                edtPValue9.Text := Trim(Copy(UV_fGulkwa, 55,  6))
             else
                edtPValue9.Text := '';
             //C011	r-GTP
             if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then
                edtPValue10.Text := Trim(Copy(UV_fGulkwa, 61,  6))
             else
                edtPValue10.Text := '';
             //C012	LAP
             if Trim(Copy(UV_fGulkwa, 67,  6)) <> '' then
                edtPValue11.Text := Trim(Copy(UV_fGulkwa, 67,  6))
             else
                edtPValue11.Text := '';
             //C013	ALP
             if Trim(Copy(UV_fGulkwa, 73,  6)) <> '' then
                edtPValue12.Text := Trim(Copy(UV_fGulkwa, 73,  6))
             else
                edtPValue12.Text := '';
            //C017	뇨산
             if Trim(Copy(UV_fGulkwa, 85,  6)) <> '' then
                edtPValue13.Text := Trim(Copy(UV_fGulkwa, 85,  6))
             else
                edtPValue13.Text := '';
            //C019	CPK
             if Trim(Copy(UV_fGulkwa, 91,  6)) <> '' then
                edtPValue14.Text := Trim(Copy(UV_fGulkwa, 91,  6))
             else
                edtPValue14.Text := '';
            //C021	LDH
             if Trim(Copy(UV_fGulkwa, 97,  6)) <> '' then
                edtPValue15.Text := Trim(Copy(UV_fGulkwa, 97,  6))
             else
                edtPValue15.Text := '';
             //C025	콜레스테롤
             if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then
                edtPValue16.Text := Trim(Copy(UV_fGulkwa, 121,  6))
             else
                edtPValue16.Text := '';
            //C026	HDL-콜레스테롤
             if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then
                edtPValue17.Text := Trim(Copy(UV_fGulkwa, 127,  6))
             else
                edtPValue17.Text := '';
            //C027(계산)	LDL-콜레스테롤
             if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then
                edtPValue18.Text := Trim(Copy(UV_fGulkwa, 133,  6))
             else
                edtPValue18.Text := '';
             //C028	중성지방
             if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then
                edtPValue19.Text := Trim(Copy(UV_fGulkwa, 139,  6))
             else
                edtPValue19.Text := '';
             //C029(계산)	cardiac risk index
             if Trim(Copy(UV_fGulkwa, 151,  6)) <> '' then
                edtPValue20.Text := Trim(Copy(UV_fGulkwa, 151,  6))
             else
                edtPValue20.Text := '';
             //C032	공복시혈당
             if Trim(Copy(UV_fGulkwa, 157,  6)) <> '' then
                edtPValue21.Text := Trim(Copy(UV_fGulkwa, 157,  6))
             else
                edtPValue21.Text := '';
             //C033	식후혈당
             if Trim(Copy(UV_fGulkwa, 163,  6)) <> '' then
                edtPValue22.Text := Trim(Copy(UV_fGulkwa, 163,  6))
             else
                edtPValue22.Text := '';
             //C041	뇨소질소
             if Trim(Copy(UV_fGulkwa, 181,  6)) <> '' then
                edtPValue23.Text := Trim(Copy(UV_fGulkwa, 181,  6))
             else
                edtPValue23.Text := '';
             //C042	크레아티닌
             if Trim(Copy(UV_fGulkwa, 187,  6)) <> '' then
                edtPValue24.Text := Trim(Copy(UV_fGulkwa, 187,  6))
             else
                edtPValue24.Text := '';
             //C043(계산)	BUN/Cr 비율
             if Trim(Copy(UV_fGulkwa, 193,  6)) <> '' then
                edtPValue25.Text := Trim(Copy(UV_fGulkwa, 193,  6))
             else
                edtPValue25.Text := '';
             //LDL-콜레스테롤(효소)(C074)
             if Trim(Copy(UV_fGulkwa, 331,  6)) <> '' then edtPValue26.Text := Trim(Copy(UV_fGulkwa, 331,  6))
             else                                          edtPValue26.Text := '';
             //식전 자가혈당측정기(C077)
             if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then edtPValue27.Text := Trim(Copy(UV_fGulkwa, 349,  6))
             else                                          edtPValue27.Text := '';
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
   UP_DChange(edtValue13);
   UP_DChange(edtValue14);
   UP_DChange(edtValue15);
   UP_DChange(edtValue16);
   UP_DChange(edtValue17);
   UP_DChange(edtValue18);
   UP_DChange(edtValue19);
   UP_DChange(edtValue20);
   UP_DChange(edtValue21);
   UP_DChange(edtValue22);
   UP_DChange(edtValue23);
   UP_DChange(edtValue24);
   UP_DChange(edtValue25);
   UP_DChange(edtValue26);      



   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD1I17.btnPInsertClick(Sender: TObject);
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

procedure TfrmLD1I17.UP_Click(Sender: TObject);
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

procedure TfrmLD1I17.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I17.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD1I17.UP_Exit(Sender: TObject);
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

procedure TfrmLD1I17.rbDateClick(Sender: TObject);
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
      mskJumin.SetFocus;

      // 찾기 Disable
      edtFind.Enabled  := False;
   end;
end;

procedure TfrmLD1I17.edtFindExit(Sender: TObject);
begin
   inherited;

   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0);
end;

procedure TfrmLD1I17.mskNumExit(Sender: TObject);
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

procedure TfrmLD1I17.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F3 then
   begin
      if edtFind.Enabled then edtFind.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I17.UP_EditDisplay(Sender : TObject);
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

procedure TfrmLD1I17.edtValue1Change(Sender: TObject);
begin
  inherited;
   //해당하는 값을 넣어줌.
   if Copy(TEdit(Sender).Name, 1, 8)  = 'edtValue' then
      UP_EditDisplay(Sender);
end;

end.
