//==============================================================================
// 프로그램 설명 : 생화확 검사항목 결과오류
// 시스템        : 통합검진
// 수정일자      : 20130614
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.09.03
// 수정자        : 곽윤설
// 수정내용      : Part_02 - C074추가
// 참고사항      : [본원-최은희]
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================

unit LD4Q47;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q47 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label4: TLabel;
    Cmb_gubn: TComboBox;
    Chk_special: TRadioButton;
    Chk_part_01: TRadioButton;
    Chk_part_02: TRadioButton;
    qryJHangmokList: TQuery;
    group_radio: TGroupBox;
    Query1: TQuery;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN,
    Uv_vC022,Uv_vC023,Uv_vC024,Uv_vC025,Uv_vC026,Uv_vC027,Uv_vC028,Uv_vC029,Uv_vC041,Uv_vC042,Uv_vC043,Uv_vC045,Uv_vC046,Uv_vC047,Uv_vC048,Uv_vC056,
    Uv_vC001,Uv_vC002,Uv_vC003,Uv_vC004,Uv_vC005,Uv_vC006,Uv_vC007,Uv_vC009,Uv_vC010,Uv_vC011,Uv_vC012,Uv_vC013,Uv_vC017,Uv_vC019,Uv_vC021,Uv_vC030,
    Uv_vC032,Uv_vC033,Uv_vC034,Uv_vC039,Uv_vC057,Uv_vC058,Uv_vC078,Uv_vC080,Uv_vC082,Uv_vC083,Uv_vC050,Uv_vC051,Uv_vC052,Uv_vC074,
    UV_vSampNo, UV_vDesc_dc : Variant;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q47: TfrmLD4Q47;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q471;

{$R *.DFM}

procedure TfrmLD4Q47.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q47.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //수정
      end;
   end;
end;

procedure TfrmLD4Q47.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      Uv_vC022:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC023:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC024:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC025:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC026:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC027:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC028:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC029:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC041:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC042:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC043:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC045:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC046:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC047:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC048:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC056:= VarArrayCreate([0, 0], varOleStr);

      Uv_vC001:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC002:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC003:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC004:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC005:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC006:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC007:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC009:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC010:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC011:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC012:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC013:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC017:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC019:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC021:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC030:= VarArrayCreate([0, 0], varOleStr);

      Uv_vC032:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC033:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC034:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC039:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC057:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC058:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC078:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC080:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC082:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC083:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC050:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC051:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC052:= VarArrayCreate([0, 0], varOleStr);
      Uv_vC074:= VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);

      VarArrayRedim(Uv_vC022,  iLength);
      VarArrayRedim(Uv_vC023,  iLength);
      VarArrayRedim(Uv_vC024,  iLength);
      VarArrayRedim(Uv_vC025,  iLength);
      VarArrayRedim(Uv_vC026,  iLength);
      VarArrayRedim(Uv_vC027,  iLength);
      VarArrayRedim(Uv_vC028,  iLength);
      VarArrayRedim(Uv_vC029,  iLength);
      VarArrayRedim(Uv_vC041,  iLength);
      VarArrayRedim(Uv_vC042,  iLength);
      VarArrayRedim(Uv_vC043,  iLength);
      VarArrayRedim(Uv_vC045,  iLength);
      VarArrayRedim(Uv_vC046,  iLength);
      VarArrayRedim(Uv_vC047,  iLength);
      VarArrayRedim(Uv_vC048,  iLength);
      VarArrayRedim(Uv_vC056,  iLength);
      VarArrayRedim(Uv_vC001,  iLength);
      VarArrayRedim(Uv_vC002,  iLength);
      VarArrayRedim(Uv_vC003,  iLength);
      VarArrayRedim(Uv_vC004,  iLength);
      VarArrayRedim(Uv_vC005,  iLength);
      VarArrayRedim(Uv_vC006,  iLength);
      VarArrayRedim(Uv_vC007,  iLength);
      VarArrayRedim(Uv_vC009,  iLength);
      VarArrayRedim(Uv_vC010,  iLength);
      VarArrayRedim(Uv_vC011,  iLength);
      VarArrayRedim(Uv_vC012,  iLength);
      VarArrayRedim(Uv_vC013,  iLength);
      VarArrayRedim(Uv_vC017,  iLength);
      VarArrayRedim(Uv_vC019,  iLength);
      VarArrayRedim(Uv_vC021,  iLength);
      VarArrayRedim(Uv_vC030,  iLength);
      VarArrayRedim(Uv_vC032,  iLength);
      VarArrayRedim(Uv_vC033,  iLength);
      VarArrayRedim(Uv_vC034,  iLength);
      VarArrayRedim(Uv_vC039,  iLength);
      VarArrayRedim(Uv_vC057,  iLength);
      VarArrayRedim(Uv_vC058,  iLength);
      VarArrayRedim(Uv_vC078,  iLength);
      VarArrayRedim(Uv_vC080,  iLength);
      VarArrayRedim(Uv_vC082,  iLength);
      VarArrayRedim(Uv_vC083,  iLength);
      VarArrayRedim(Uv_vC050,  iLength);
      VarArrayRedim(Uv_vC051,  iLength);
      VarArrayRedim(Uv_vC052,  iLength);
      VarArrayRedim(Uv_vC074,  iLength);

      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD4Q47.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q47.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q47.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;

   // 자룔를 화면에 조회

   if Chk_special.Checked =true  then
   begin
        case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := Uv_vC022[DataRow-1];
           6 : Value := Uv_vC023[DataRow-1];
           7 : Value := Uv_vC024[DataRow-1];
           8 : Value := Uv_vC025[DataRow-1];
           9 : Value := Uv_vC026[DataRow-1];
          10 : Value := Uv_vC027[DataRow-1];
          11 : Value := Uv_vC028[DataRow-1];
          12 : Value := Uv_vC029[DataRow-1];
          13 : Value := Uv_vC041[DataRow-1];
          14 : Value := Uv_vC042[DataRow-1];
          15 : Value := Uv_vC043[DataRow-1];
          16 : Value := Uv_vC045[DataRow-1];
          17 : Value := Uv_vC046[DataRow-1];
          18 : Value := Uv_vC047[DataRow-1];
          19 : Value := Uv_vC048[DataRow-1];
          20 : Value := Uv_vC056[DataRow-1];

          end;
          for i:=5 to 20 do
          begin
               grdMaster.Col[i].heading :=' ';
          end;
          grdMaster.Col[5].heading  :='APOA C022';
          grdMaster.Col[6].heading  :='APOB C023';
          grdMaster.Col[7].heading  :='B/A C024';
          grdMaster.Col[8].heading  :='Tich C025';
          grdMaster.Col[9].heading  :='HDL C026';
          grdMaster.Col[10].heading :='LDL C027';
          grdMaster.Col[11].heading :='TG C028';
          grdMaster.Col[12].heading :='CRI C029';
          grdMaster.Col[13].heading :='Bun C041';
          grdMaster.Col[14].heading :='Crea C042';
          grdMaster.Col[15].heading :='B/C C043';
          grdMaster.Col[16].heading :='FE C045';
          grdMaster.Col[17].heading :='TIBC C046';
          grdMaster.Col[18].heading :='TIBCS C047';
          grdMaster.Col[19].heading :='UIBC C048';
          grdMaster.Col[20].heading :='Ca C056';

          grdMaster.Col[19].Width := 45;
          grdMaster.Col[20].Width := 48;

    end
    else if Chk_part_01.Checked =true  then
    begin
        case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := Uv_vC001[DataRow-1];
           6 : Value := Uv_vC002[DataRow-1];
           7 : Value := Uv_vC003[DataRow-1];
           8 : Value := Uv_vC004[DataRow-1];
           9 : Value := Uv_vC005[DataRow-1];
          10 : Value := Uv_vC006[DataRow-1];
          11 : Value := Uv_vC007[DataRow-1];
          12 : Value := Uv_vC009[DataRow-1];
          13 : Value := Uv_vC010[DataRow-1];
          14 : Value := Uv_vC011[DataRow-1];
          15 : Value := Uv_vC012[DataRow-1];
          16 : Value := Uv_vC013[DataRow-1];
          17 : Value := Uv_vC017[DataRow-1];
          18 : Value := Uv_vC019[DataRow-1];
          19 : Value := Uv_vC021[DataRow-1];
          20 : Value := Uv_vC030[DataRow-1];

          end;
          for i:=5 to 20 do
          begin
               grdMaster.Col[i].heading :=' ';
          end;
          grdMaster.Col[5].heading  :='TP C001';
          grdMaster.Col[6].heading  :='Alb C002';
          grdMaster.Col[7].heading  :='Glo C003';
          grdMaster.Col[8].heading  :='A/G C004';
          grdMaster.Col[9].heading  :='TB C005';
          grdMaster.Col[10].heading :='DB C006';
          grdMaster.Col[11].heading :='IB C007';
          grdMaster.Col[12].heading :='AST C009';
          grdMaster.Col[13].heading :='ALT C010';
          grdMaster.Col[14].heading :='GTP C011';
          grdMaster.Col[15].heading :='LAP C012';
          grdMaster.Col[16].heading :='ALP C013';
          grdMaster.Col[17].heading :='UA C017';
          grdMaster.Col[18].heading :='CPK C019';
          grdMaster.Col[19].heading :='LDH C021';
          grdMaster.Col[20].heading :='B-Lipo C030';

         grdMaster.Col[19].Width := 45;
          grdMaster.Col[20].Width := 48;
    end
    else if Chk_part_02.Checked =true  then
    begin
       case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := Uv_vC032[DataRow-1];
           6 : Value := Uv_vC033[DataRow-1];
           7 : Value := Uv_vC034[DataRow-1];
           8 : Value := Uv_vC039[DataRow-1];
           9 : Value := Uv_vC057[DataRow-1];
          10 : Value := Uv_vC058[DataRow-1];
          11 : Value := Uv_vC078[DataRow-1];
          12 : Value := Uv_vC080[DataRow-1];
          13 : Value := Uv_vC082[DataRow-1];
          14 : Value := Uv_vC083[DataRow-1];
          15 : Value := Uv_vC074[DataRow-1];
          16 : Value := Uv_vC050[DataRow-1];
          17 : Value := Uv_vC051[DataRow-1];
          18 : Value := Uv_vC052[DataRow-1];

          end;
          for i:=5 to 20 do
          begin
               grdMaster.Col[i].heading :=' ';
          end;
          grdMaster.Col[5].heading  :='Glu C032';
          grdMaster.Col[6].heading  :='PP2 C033';
          grdMaster.Col[7].heading  :='FA C034';
          grdMaster.Col[8].heading  :='AMY C039';
          grdMaster.Col[9].heading  :='I.P C057';
          grdMaster.Col[10].heading :='Mg C058';
          grdMaster.Col[11].heading :='Homo C078';
          grdMaster.Col[12].heading :='Lp(a) C080';
          grdMaster.Col[13].heading :='Lip C082';
          grdMaster.Col[14].heading :='A1C C083';
          grdMaster.Col[15].heading :='LDL-C C074';
          grdMaster.Col[16].heading :='Na C050';
          grdMaster.Col[17].heading :='K C051';
          grdMaster.Col[18].heading :='CI C052';
          grdMaster.Col[19].heading :='';
          grdMaster.Col[20].heading :='';
          grdMaster.Col[19].Width := 0;
          grdMaster.Col[20].Width := 0;
    end;                               

end;

procedure TfrmLD4Q47.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
   Sum_vC022,Sum_vC023,Sum_vC024,Sum_vC025,Sum_vC026,Sum_vC027,Sum_vC028,Sum_vC029,Sum_vC041,Sum_vC042,Sum_vC043,Sum_vC045,Sum_vC046,Sum_vC047,Sum_vC048,Sum_vC056,
   Sum_vC001,Sum_vC002,Sum_vC003,Sum_vC004,Sum_vC005,Sum_vC006,Sum_vC007,Sum_vC009,Sum_vC010,Sum_vC011,Sum_vC012,Sum_vC013,Sum_vC017,Sum_vC019,Sum_vC021,Sum_vC030,
   Sum_vC032,Sum_vC033,Sum_vC034,Sum_vC039,Sum_vC057,Sum_vC058,Sum_vC078,Sum_vC080,Sum_vC082,Sum_vC083,Sum_vC050,Sum_vC051,Sum_vC052,Sum_vC074 :  integer;


    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    UV_fGulkwa_07, UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07  : String;
    vCod_chuga : Variant;
    Check :Boolean;

begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   Sum_vC022 :=0; Sum_vC023 :=0; Sum_vC024 :=0; Sum_vC025 :=0; Sum_vC026 :=0; Sum_vC027 :=0; Sum_vC028 :=0; Sum_vC029 :=0;
   Sum_vC041 :=0; Sum_vC042 :=0; Sum_vC043 :=0; Sum_vC045 :=0; Sum_vC046 :=0; Sum_vC047 :=0; Sum_vC048 :=0; Sum_vC056 :=0;

   Sum_vC001 :=0; Sum_vC002 :=0; Sum_vC003 :=0; Sum_vC004 :=0; Sum_vC005 :=0; Sum_vC006 :=0; Sum_vC007 :=0; Sum_vC009 :=0;
   Sum_vC010 :=0; Sum_vC011 :=0; Sum_vC012 :=0; Sum_vC013 :=0; Sum_vC017 :=0; Sum_vC019 :=0; Sum_vC021 :=0; Sum_vC030 :=0;

   Sum_vC032 :=0; Sum_vC033 :=0; Sum_vC034 :=0; Sum_vC039 :=0; Sum_vC057 :=0; Sum_vC058 :=0; Sum_vC078 :=0; Sum_vC080 :=0;
   Sum_vC082 :=0; Sum_vC083 :=0; Sum_vC050 :=0; Sum_vC051 :=0; Sum_vC052 :=0;





   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B inner join gumgin G on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha'+
                 '         left outer join gulkwa K  ON G.cod_jisa = k.cod_jisa and G.num_id = K.num_id and G.dat_gmgn = K.dat_gmgn ';
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + '''  and gubn_part=''01''';



      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';   
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno ';



      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q36', 'Q', 'N');
         while not qryBunju.Eof do
         Begin

            sHul_List := '';

            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;
            
            sHul_List := UF_hangmokList;

            if ((Pos('C022' ,sHul_List)> 0) or (Pos('C026' ,sHul_List)> 0) or (Pos('C041' ,sHul_List)> 0) or (Pos('C046' ,sHul_List)> 0)or
               (Pos('C023' ,sHul_List)> 0) or (Pos('C027' ,sHul_List)> 0) or (Pos('C042' ,sHul_List)> 0) or (Pos('C047' ,sHul_List)> 0)or
               (Pos('C024' ,sHul_List)> 0) or (Pos('C028' ,sHul_List)> 0) or (Pos('C043' ,sHul_List)> 0) or (Pos('C048' ,sHul_List)> 0)or
               (Pos('C025' ,sHul_List)> 0) or (Pos('C029' ,sHul_List)> 0) or (Pos('C045' ,sHul_List)> 0) or (Pos('C056' ,sHul_List)> 0))and (Chk_special.Checked =True )Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         Check := False;
                         UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('C022', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 103,  6)) = '') then
                                  begin
                                  UV_vC022[Index] := '결과無';
                                  Sum_vC022 := Sum_vC022 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C023', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 109,  6)) = '') then
                                  begin
                                  UV_vC023[Index] := '결과無';
                                  Sum_vC023 := Sum_vC023 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('C024', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 115,  6)) = '') then
                                  begin
                                  UV_vC024[Index] := '결과無';
                                  Sum_vC024 := Sum_vC024 + 1;
                                  Check := True;
                               end;
                          end;
                             ///////////////////////////////////////////////////////////
                          if (pos('C025', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 121,  6)) = '') then
                                  begin
                                  UV_vC025[Index] := '결과無';
                                  Sum_vC025 := Sum_vC025 + 1;
                                  Check := True;
                               end;
                          end;
                              ///////////////////////////////////////////////////////////
                          if (pos('C026', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 127,  6)) = '') then
                                  begin
                                  UV_vC026[Index] := '결과無';
                                  Sum_vC026 := Sum_vC026 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('C027', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 133,  6)) = '') then
                                  begin
                                  UV_vC027[Index] := '결과無';
                                  Sum_vC027 := Sum_vC027 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('C028', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 139,  6)) = '') then
                                  begin
                                  UV_vC028[Index] := '결과無';
                                  Sum_vC028 := Sum_vC028 + 1;
                                  Check := True;
                               end;
                          end;
                                 ///////////////////////////////////////////////////////////
                          if (pos('C029', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 151,  6)) = '') then
                                  begin
                                  UV_vC029[Index] := '결과無';
                                  Sum_vC029 := Sum_vC029 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('C041', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 181,  6)) = '') then
                                  begin
                                  UV_vC041[Index] := '결과無';
                                  Sum_vC041 := Sum_vC041 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('C042', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 187,  6)) = '') then
                                  begin
                                  UV_vC042[Index] := '결과無';
                                  Sum_vC042 := Sum_vC042 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('C043', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 193,  6)) = '') then
                                  begin
                                  UV_vC043[Index] := '결과無';
                                  Sum_vC043:= Sum_vC043 + 1;
                                  Check := True;
                               end;
                          end;
                                    ///////////////////////////////////////////////////////////
                          if (pos('C045', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 199,  6)) = '') then
                                  begin
                                  UV_vC045[Index] := '결과無';
                                  Sum_vC045:= Sum_vC045 + 1;
                                  Check := True;
                               end;
                          end;
                                     ///////////////////////////////////////////////////////////
                          if (pos('C046', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 205,  6)) = '') then
                                  begin
                                  UV_vC046[Index] := '결과無';
                                  Sum_vC046:= Sum_vC046 + 1;
                                  Check := True;
                               end;
                          end;
                                      ///////////////////////////////////////////////////////////
                          if (pos('C047', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 211,  6)) = '') then
                                  begin
                                  UV_vC047[Index] := '결과無';
                                  Sum_vC047:= Sum_vC047 + 1;
                                  Check := True;
                               end;
                          end;
                                      ///////////////////////////////////////////////////////////
                          if (pos('C048', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 217,  6)) = '') then
                                  begin
                                  UV_vC048[Index] := '결과無';
                                  Sum_vC048:= Sum_vC048 + 1;
                                  Check := True;
                               end;
                          end;
                                       ///////////////////////////////////////////////////////////
                          if (pos('C056', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 247,  6)) = '') then
                                  begin
                                  UV_vC056[Index] := '결과無';
                                  Sum_vC056:= Sum_vC056 + 1;
                                  Check := True;
                               end;
                          end;

                          if Check = True then
                          begin
                               UP_VarMemSet(Index+1);
                               inc(index);
                          end;
                   qryGulkwa.Next;
                   end;
               End;
               qryGulkwa.Close;

              end
              else if ((Pos('C001' ,sHul_List)> 0) or (Pos('C005' ,sHul_List)> 0) or (Pos('C010' ,sHul_List)> 0) or (Pos('C017' ,sHul_List)> 0)or
                       (Pos('C002' ,sHul_List)> 0) or (Pos('C006' ,sHul_List)> 0) or (Pos('C011' ,sHul_List)> 0) or (Pos('C019' ,sHul_List)> 0)or
                       (Pos('C003' ,sHul_List)> 0) or (Pos('C007' ,sHul_List)> 0) or (Pos('C012' ,sHul_List)> 0) or (Pos('C021' ,sHul_List)> 0)or
                       (Pos('C004' ,sHul_List)> 0) or (Pos('C009' ,sHul_List)> 0) or (Pos('C013' ,sHul_List)> 0) or (Pos('C030' ,sHul_List)> 0))and (Chk_part_01.Checked =True )Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         Check := False;
                         UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('C001', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 1,  6)) = '') then
                                  begin
                                  UV_vC001[Index] := '결과無';
                                  Sum_vC001 := Sum_vC001 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C002', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                                  begin
                                  UV_vC002[Index] := '결과無';
                                  Sum_vC002 := Sum_vC002 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('C003', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 13,  6)) = '') then
                                  begin
                                  UV_vC003[Index] := '결과無';
                                  Sum_vC003 := Sum_vC003 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('C004', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                                  begin
                                  UV_vC004[Index] := '결과無';
                                  Sum_vC004 := Sum_vC004 + 1;
                                  Check := True;
                               end;
                          end;
                          ////////////////
                          if (pos('C005', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                                  begin
                                  UV_vC005[Index] := '결과無';
                                  Sum_vC005 := Sum_vC005 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////
                          if (pos('C006', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 31,  6)) = '') then
                                  begin
                                  UV_vC006[Index] := '결과無';
                                  Sum_vC006 := Sum_vC006 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('C007', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                                  begin
                                  UV_vC007[Index] := '결과無';
                                  Sum_vC007 := Sum_vC007 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('C009', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 49,  6)) = '') then
                                  begin
                                  UV_vC009[Index] := '결과無';
                                  Sum_vC009 := Sum_vC009 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('C010', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 55,  6)) = '') then
                                  begin
                                  UV_vC010[Index] := '결과無';
                                  Sum_vC010 := Sum_vC010 + 1;
                                  Check := True;
                               end;
                          end;
                             ///////////////////////////////////////////////////////////
                          if (pos('C011', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 61,  6)) = '') then
                                  begin
                                  UV_vC011[Index] := '결과無';
                                  Sum_vC011 := Sum_vC011 + 1;
                                  Check := True;
                               end;
                          end;
                              ///////////////////////////////////////////////////////////
                          if (pos('C012', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                                  begin
                                  UV_vC012[Index] := '결과無';
                                  Sum_vC012 := Sum_vC012 + 1;
                                  Check := True;
                               end;
                          end;
                               ///////////////////////////////////////////////////////////
                          if (pos('C013', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 73,  6)) = '') then
                                  begin
                                  UV_vC013[Index] := '결과無';
                                  Sum_vC013 := Sum_vC013 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('C017', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 85,  6)) = '') then
                                  begin
                                  UV_vC017[Index] := '결과無';
                                  Sum_vC017 := Sum_vC017 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('C019', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 91,  6)) = '') then
                                  begin
                                  UV_vC019[Index] := '결과無';
                                  Sum_vC019 := Sum_vC019 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('C021', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 97,  6)) = '') then
                                  begin
                                  UV_vC021[Index] := '결과無';
                                  Sum_vC021 := Sum_vC021 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('C030', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 145,  6)) = '') then
                                  begin
                                  UV_vC030[Index] := '결과無';
                                  Sum_vC030 := Sum_vC030 + 1;
                                  Check := True;
                               end;
                          end;
                          if Check = True then
                          begin
                               UP_VarMemSet(Index+1);
                               inc(index);
                          end;
                   qryGulkwa.Next;
                   end;
               End;
               qryGulkwa.Close;

              end
              else if ((Pos('C032' ,sHul_List)> 0) or (Pos('C057' ,sHul_List)> 0) or (Pos('C082' ,sHul_List)> 0) or (Pos('C050' ,sHul_List)> 0)or
                       (Pos('C033' ,sHul_List)> 0) or (Pos('C058' ,sHul_List)> 0) or (Pos('C083' ,sHul_List)> 0) or (Pos('C051' ,sHul_List)> 0)or
                       (Pos('C034' ,sHul_List)> 0) or (Pos('C078' ,sHul_List)> 0) or (Pos('C052' ,sHul_List)> 0) or (Pos('C039' ,sHul_List)> 0)or
                       (Pos('C080' ,sHul_List)> 0) or (Pos('C074' ,sHul_List)> 0))and (Chk_part_02.Checked =True )Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         Check := False;
                         UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('C032', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 157,  6)) = '') then
                                  begin
                                  UV_vC032[Index] := '결과無';
                                  Sum_vC032 := Sum_vC032 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C033', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 163,  6)) = '') then
                                  begin
                                  UV_vC033[Index] := '결과無';
                                  Sum_vC033 := Sum_vC033 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C034', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 169,  6)) = '') then
                                  begin
                                  UV_vC034[Index] := '결과無';
                                  Sum_vC034 := Sum_vC034 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C039', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 175,  6)) = '') then
                                  begin
                                  UV_vC039[Index] := '결과無';
                                  Sum_vC039 := Sum_vC039 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C050', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 223,  6)) = '') then
                                  begin
                                  UV_vC050[Index] := '결과無';
                                  Sum_vC050 := Sum_vC050 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C051', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 229,  6)) = '') then
                                  begin
                                  UV_vC051[Index] := '결과無';
                                  Sum_vC051 := Sum_vC051 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C052', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 235,  6)) = '') then
                                  begin
                                  UV_vC052[Index] := '결과無';
                                  Sum_vC052 := Sum_vC052 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C057', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 253,  6)) = '') then
                                  begin
                                  UV_vC057[Index] := '결과無';
                                  Sum_vC057 := Sum_vC057 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C058', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 259,  6)) = '') then
                                  begin
                                  UV_vC058[Index] := '결과無';
                                  Sum_vC058 := Sum_vC058 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('C074', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 331,  6)) = '') then
                                  begin
                                  UV_vC074[Index] := '결과無';
                                  Sum_vC074 := Sum_vC074 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C078', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 355,  6)) = '') then
                                  begin
                                  UV_vC078[Index] := '결과無';
                                  Sum_vC078 := Sum_vC078 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C080', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 367,  6)) = '') then
                                  begin
                                  UV_vC080[Index] := '결과無';
                                  Sum_vC080 := Sum_vC080 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C082', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 379,  6)) = '') then
                                  begin
                                  UV_vC082[Index] := '결과無';
                                  Sum_vC082 := Sum_vC082 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('C083', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 385,  6)) = '') then
                                  begin
                                  UV_vC083[Index] := '결과無';
                                  Sum_vC083 := Sum_vC083 + 1;
                                  Check := True;
                               end;
                          end;



                          if Check = True then
                          begin
                               UP_VarMemSet(Index+1);
                               inc(index);
                          end;
                   qryGulkwa.Next;
                   end;
               End;
               qryGulkwa.Close;

              end;



           qryBunju.Next;
         End;


         // Table과 Disconnect
         qryBunju.Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '총  합  계';
         if (Chk_part_01.Checked =True ) then
         begin
              Uv_vC001[Index]:= Sum_vC001;
              Uv_vC002[Index]:= Sum_vC002;
              Uv_vC003[Index]:= Sum_vC003;
              Uv_vC004[Index]:= Sum_vC004;
              Uv_vC005[Index]:= Sum_vC005;
              Uv_vC006[Index]:= Sum_vC006;
              Uv_vC007[Index]:= Sum_vC007;
              Uv_vC009[Index]:= Sum_vC009;
              Uv_vC010[Index]:= Sum_vC010;
              Uv_vC011[Index]:= Sum_vC011;
              Uv_vC012[Index]:= Sum_vC012;
              Uv_vC013[Index]:= Sum_vC013;
              Uv_vC017[Index]:= Sum_vC017;
              Uv_vC019[Index]:= Sum_vC019;
              Uv_vC021[Index]:= Sum_vC021;
              Uv_vC030[Index]:= Sum_vC030;
         end
         else if (Chk_part_02.Checked =True ) then
         begin
              Uv_vC032[Index]:= Sum_vC032;
              Uv_vC033[Index]:= Sum_vC033;
              Uv_vC034[Index]:= Sum_vC034;
              Uv_vC039[Index]:= Sum_vC039;
              Uv_vC057[Index]:= Sum_vC057;
              Uv_vC058[Index]:= Sum_vC058;
              Uv_vC078[Index]:= Sum_vC078;
              Uv_vC080[Index]:= Sum_vC080;
              Uv_vC082[Index]:= Sum_vC082;
              Uv_vC083[Index]:= Sum_vC083;
              Uv_vC050[Index]:= Sum_vC050;
              Uv_vC051[Index]:= Sum_vC051;
              Uv_vC052[Index]:= Sum_vC052;
              Uv_vC074[Index]:= Sum_vC074;
         end
         else  if (Chk_special.Checked =True ) then
         begin
              Uv_vC022[Index]:= Sum_vC022;
              Uv_vC023[Index]:= Sum_vC023;
              Uv_vC024[Index]:= Sum_vC024;
              Uv_vC025[Index]:= Sum_vC025;
              Uv_vC026[Index]:= Sum_vC026;
              Uv_vC027[Index]:= Sum_vC027;
              Uv_vC028[Index]:= Sum_vC028;
              Uv_vC029[Index]:= Sum_vC029;
              Uv_vC041[Index]:= Sum_vC041;
              Uv_vC042[Index]:= Sum_vC042;
              Uv_vC043[Index]:= Sum_vC043;
              Uv_vC045[Index]:= Sum_vC045;
              Uv_vC046[Index]:= Sum_vC046;
              Uv_vC047[Index]:= Sum_vC047;
              Uv_vC048[Index]:= Sum_vC048;
              Uv_vC056[Index]:= Sum_vC056;
         end;


         UV_vSampNo[Index]  := '';
         Inc(index);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q47.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q47.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q471 := TfrmLD4Q471.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD4Q471.QuickRep.Preview
  else                                frmLD4Q471.QuickRep.Print;

end;

procedure TfrmLD4Q47.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;


end;

function TfrmLD4Q47.UF_hangmokList : String;
var
   i, j, k : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
       sHmCode :string;
   vCod_chuga : Variant;
   iRet : integer;
begin
   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;

   result := sTemp;
end;

function TfrmLD4Q47.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.
