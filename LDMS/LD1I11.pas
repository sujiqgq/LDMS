//==============================================================================
// 프로그램 설명 : 혈액결과,소견체크
// 시스템        : 공통(등록 & 조회)
// 수정일자      : 2004.10.
// 수정자        : 최지혜
// 수정내용      :                                                                   
// 참고사항      :
//==============================================================================
// 수정일자      : 2007.04.09
// 수정자        : 유동구
// 수정내용      : 조직학 추가..[단, 분주일자가 검진일자로 사용]
//==============================================================================
// 수정일자      : 2008.10.28
// 수정자        : 김승철
// 수정내용      : 김소영 선생님 권한추가.
//==============================================================================
// 수정일자      : 2011.12.21
// 수정자        : 유동구
// 수정내용      : [센터]자체검사 실시관련 추가...(qryHgulkwa_EtcChk)
//==============================================================================
unit LD1I11;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, StdCtrls, ExtCtrls, Buttons, ComCtrls, ToolWin, Mask, DateEdit,
  Grids_ts, TSGrid, Db, DBTables;

type
  TfrmLD1I11 = class(TfrmSingle)
    pnlCond: TPanel;
    btnDate: TSpeedButton;
    Label2: TLabel;
    btnDATE_1: TSpeedButton;
    Label7: TLabel;
    mskST_Date: TDateEdit;
    mskED_Date: TDateEdit;
    pnlJisa: TPanel;
    Label8: TLabel;
    cmbJisa: TComboBox;
    grdMaster: TtsGrid;
    qryHgulkwa_chk: TQuery;
    qryUser_priv: TQuery;
    qrySawon: TQuery;
    pnlDetail: TPanel;
    grp01: TGroupBox;
    Label1: TLabel;
    Label3: TLabel;
    dat01: TLabel;
    cod01: TLabel;
    Label40: TLabel;
    cmb01: TComboBox;
    grp02: TGroupBox;
    Label6: TLabel;
    dat02: TLabel;
    Label10: TLabel;
    cod02: TLabel;
    Label41: TLabel;
    cmb02: TComboBox;
    grp03: TGroupBox;
    Label12: TLabel;
    dat03: TLabel;
    Label14: TLabel;
    cod03: TLabel;
    Label42: TLabel;
    cmb03: TComboBox;
    grp04: TGroupBox;
    Label16: TLabel;
    Label18: TLabel;
    dat04: TLabel;
    cod04: TLabel;
    Label43: TLabel;
    cmb04: TComboBox;
    grp07: TGroupBox;
    Label36: TLabel;
    dat07: TLabel;
    Label38: TLabel;
    cod07: TLabel;
    Label46: TLabel;
    cmb07: TComboBox;
    grpHS: TGroupBox;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    datHS: TLabel;
    codHS: TLabel;
    cmbHS: TComboBox;
    grp08: TGroupBox;
    Label32: TLabel;
    dat08: TLabel;
    Label34: TLabel;
    cod08: TLabel;
    Label47: TLabel;
    cmb08: TComboBox;
    grp05: TGroupBox;
    Label20: TLabel;
    dat05: TLabel;
    Label22: TLabel;
    cod05: TLabel;
    Label44: TLabel;
    cmb05: TComboBox;
    qryU_Hgulkwa_chk: TQuery;
    qryNameToCode: TQuery;
    grp06: TGroupBox;
    Label4: TLabel;
    dat06: TLabel;
    Label9: TLabel;
    cod06: TLabel;
    Label13: TLabel;
    cmb06: TComboBox;
    Label15: TLabel;
    Label17: TLabel;
    DEdt_gumgin: TDateEdit;
    Label19: TLabel;
    CombDate: TComboBox;
    qryCell: TQuery;
    Label5: TLabel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Label11: TLabel;
    Label21: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    ck01_12: TLabel;
    ck01_11: TLabel;
    ck01_43: TLabel;
    ck01_71: TLabel;
    ck01_61: TLabel;
    ck01_51: TLabel;
    cod01_12: TLabel;
    cod01_11: TLabel;
    Label39: TLabel;
    Label45: TLabel;
    cod01_43: TLabel;
    cod01_71: TLabel;
    cod01_61: TLabel;
    cod01_51: TLabel;
    Label55: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    Label75: TLabel;
    Label76: TLabel;
    Label77: TLabel;
    Label78: TLabel;
    dat01_12: TLabel;
    dat01_11: TLabel;
    dat01_43: TLabel;
    dat01_71: TLabel;
    dat01_61: TLabel;
    dat01_51: TLabel;
    ck02_12: TLabel;
    ck02_11: TLabel;
    ck02_43: TLabel;
    ck02_71: TLabel;
    ck02_61: TLabel;
    ck02_51: TLabel;
    cod02_12: TLabel;
    cod02_11: TLabel;
    cod02_43: TLabel;
    cod02_71: TLabel;
    cod02_61: TLabel;
    cod02_51: TLabel;
    dat02_12: TLabel;
    dat02_11: TLabel;
    dat02_43: TLabel;
    dat02_71: TLabel;
    dat02_61: TLabel;
    dat02_51: TLabel;
    ck03_12: TLabel;
    ck03_11: TLabel;
    ck03_43: TLabel;
    ck03_71: TLabel;
    ck03_61: TLabel;
    ck03_51: TLabel;
    cod03_12: TLabel;
    cod03_11: TLabel;
    cod03_43: TLabel;
    cod03_71: TLabel;
    cod03_61: TLabel;
    cod03_51: TLabel;
    dat03_12: TLabel;
    dat03_11: TLabel;
    dat03_43: TLabel;
    dat03_71: TLabel;
    dat03_61: TLabel;
    dat03_51: TLabel;
    Bevel5: TBevel;
    qryHgulkwa_EtcChk: TQuery;
    Part_01: TLabel;
    Part_02: TLabel;
    Part_03: TLabel;
    procedure UP_Click(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnSaveClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure cmbJisaChange(Sender: TObject);
  private
    { Private declarations }
    UV_vCod_bjjs, UV_vDat_bunju,
    UV_vGubn_01, UV_vDat_update01, UV_vCod_update01,
    UV_vGubn_02, UV_vDat_update02, UV_vCod_update02,
    UV_vGubn_03, UV_vDat_update03, UV_vCod_update03,
    UV_vGubn_04, UV_vDat_update04, UV_vCod_update04,
    UV_vGubn_05, UV_vDat_update05, UV_vCod_update05,
    UV_vGubn_06, UV_vDat_update06, UV_vCod_update06,
    UV_vGubn_07, UV_vDat_update07, UV_vCod_update07,
    UV_vGubn_08, UV_vDat_update08, UV_vCod_update08,
    UV_vGubn_09, UV_vDat_update09, UV_vCod_update09,
    UV_vGubn_Hsokun, UV_vDat_updateHS, UV_vCod_updateHS  : Variant;

    UV_sJisa : String;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_GridSet;
    procedure UP_Save(part : string; itemindex : integer; iRow : integer);
    procedure UP_Hgulkwa_JisaChk(sCod_Bjjs, sDat_bunju : string);

    function  UF_PartGubn : Boolean;
    function  UF_Condition : Boolean;
    function  UF_Check(iTemp : Integer) : String;
    function  UF_ToItemindex(sTemp : String) : integer;
    function  UF_SawonChk(sTemp : String) : String;
    function  UF_DateChk(sTemp : String) : String;
    function  UF_NameToCode(sTemp : String) : String;

  public
    UV_Edit1, UV_Edit2, UV_Edit3, UV_Edit4, UV_Edit5,
    UV_Edit6, UV_Edit7, UV_Edit8, UV_Edit9, UV_EditHS : boolean;
    { Public declarations }
  end;

var
  frmLD1I11: TfrmLD1I11;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD1I11.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   // Button Click시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = btnDate then GF_CalendarClick(mskST_Date)
   else if Sender = btnDATE_1 then GF_CalendarClick(mskED_Date);
end;

function TfrmLD1I11.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (mskSt_Date.Text = '') or (mskEd_Date.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD1I11.btnQueryClick(Sender: TObject);
var i, index : integer;
begin
  inherited;
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   UV_Edit1 := False; UV_Edit2 := False; UV_Edit3 := False; UV_Edit4 := False; UV_Edit5  := False;
   UV_Edit6 := False; UV_Edit7 := False; UV_Edit8 := False; UV_Edit9 := False; UV_EditHS := False;

   with qryHgulkwa_chk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := copy(cmbJisa.text,1,2);
      ParamByName('sdat_bunju').AsString := mskST_date.text;
      ParamByName('edat_bunju').AsString := mskED_date.text;
      open;
      if recordcount > 0 then
      begin
         UP_VarMemSet(RecordCount-1);
         GP_query_log(qryHgulkwa_chk, 'LD1I11', 'Q', 'Y');

         for i := 0 to RecordCount - 1 do
         begin
            UV_vCod_bjjs[index]     := FieldByName('cod_bjjs').AsString;
            UV_vDat_bunju[index]    := FieldBYName('dat_bunju').AsString;
            UV_vGubn_01[index]      := FieldByName('Gubn_01').AsString;
            UV_vDat_update01[index] := FieldByName('Dat_update01').AsString;
            UV_vCod_update01[index] := UF_SawonChk(FieldByName('cod_update01').AsString);
            UV_vGubn_02[index]      := FieldByName('Gubn_02').AsString;
            UV_vDat_update02[index] := FieldByName('Dat_update02').AsString;
            UV_vCod_update02[index] := UF_SawonChk(FieldByName('cod_update02').AsString);
            UV_vGubn_03[index]      := FieldByName('Gubn_03').AsString;
            UV_vDat_update03[index] := FieldByName('Dat_update03').AsString;
            UV_vCod_update03[index] := UF_SawonChk(FieldByName('cod_update03').AsString);
            UV_vGubn_04[index]      := FieldByName('Gubn_04').AsString;
            UV_vDat_update04[index] := FieldByName('Dat_update04').AsString;
            UV_vCod_update04[index] := UF_SawonChk(FieldByName('cod_update04').AsString);
            UV_vGubn_05[index]      := FieldByName('Gubn_05').AsString;
            UV_vDat_update05[index] := FieldByName('Dat_update05').AsString;
            UV_vCod_update05[index] := UF_SawonChk(FieldByName('cod_update05').AsString);
            UV_vGubn_06[index]      := FieldByName('Gubn_06').AsString;
            UV_vDat_update06[index] := FieldByName('Dat_update06').AsString;
            UV_vCod_update06[index] := UF_SawonChk(FieldByName('cod_update06').AsString);
            UV_vGubn_07[index]      := FieldByName('Gubn_07').AsString;
            UV_vDat_update07[index] := FieldByName('Dat_update07').AsString;
            UV_vCod_update07[index] := UF_SawonChk(FieldByName('cod_update07').AsString);
            UV_vGubn_08[index]      := FieldByName('Gubn_08').AsString;
            UV_vDat_update08[index] := FieldByName('Dat_update08').AsString;
            UV_vCod_update08[index] := UF_SawonChk(FieldByName('cod_update08').AsString);
            UV_vGubn_09[index]      := FieldByName('Gubn_09').AsString;
            UV_vDat_update09[index] := FieldByName('Dat_update09').AsString;
            UV_vCod_update09[index] := UF_SawonChk(FieldByName('cod_update09').AsString);
            UV_vGubn_Hsokun[index]  := FieldByName('Gubn_Hsokun').AsString;
            UV_vDat_updateHS[index] := FieldByName('Dat_updateHS').AsString;
            UV_vCod_updateHS[index] := UF_SawonChk(FieldByName('cod_updateHS').AsString);

            Next;
            Inc(index);
         end;
      end
      else
         GF_MsgBox('Information', 'NODATA');
   end;
   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

   grdmaster.setfocus;

   UF_SawonChk(GV_sUserId);
end;

procedure TfrmLD1I11.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      Col[1].Alignment := taCenter;
   end;
end;

procedure TfrmLD1I11.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;
   // 지사권한관리
   GP_ComboJisa(cmbJisa);

   UF_SawonChk(GV_sUserId);
      
   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then  sJisa := Copy(GV_sUserId,1,2)
   else                     sJisa := GV_sJisa;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

   // 환경설정
   UP_VarMemSet(0);
   UF_PartGubn;
end;

procedure TfrmLD1I11.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs     := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_01      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update01 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update01 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_02      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update02 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update02 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_03      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update03 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update03 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_04      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update04 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update04 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_05      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update05 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update05 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_06      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update06 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update06 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_07      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update07 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update07 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_08      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update08 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update08 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_09      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_update09 := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_update09 := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_Hsokun  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_updateHS := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_updateHS := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,       iLength);
      VarArrayRedim(UV_vDat_bunju,      iLength);
      VarArrayRedim(UV_vGubn_01,        iLength);
      VarArrayRedim(UV_vDat_Update01,   iLength);
      VarArrayRedim(UV_vCod_Update01,   iLength);
      VarArrayRedim(UV_vGubn_02,        iLength);
      VarArrayRedim(UV_vDat_Update02,   iLength);
      VarArrayRedim(UV_vCod_Update02,   iLength);
      VarArrayRedim(UV_vGubn_03,        iLength);
      VarArrayRedim(UV_vDat_Update03,   iLength);
      VarArrayRedim(UV_vCod_Update03,   iLength);
      VarArrayRedim(UV_vGubn_04,        iLength);
      VarArrayRedim(UV_vDat_Update04,   iLength);
      VarArrayRedim(UV_vCod_Update04,   iLength);
      VarArrayRedim(UV_vGubn_05,        iLength);
      VarArrayRedim(UV_vDat_Update05,   iLength);
      VarArrayRedim(UV_vCod_Update05,   iLength);
      VarArrayRedim(UV_vGubn_06,        iLength);
      VarArrayRedim(UV_vDat_Update06,   iLength);
      VarArrayRedim(UV_vCod_Update06,   iLength);
      VarArrayRedim(UV_vGubn_07,        iLength);
      VarArrayRedim(UV_vDat_Update07,   iLength);
      VarArrayRedim(UV_vCod_Update07,   iLength);
      VarArrayRedim(UV_vGubn_08,        iLength);
      VarArrayRedim(UV_vDat_Update08,   iLength);
      VarArrayRedim(UV_vCod_Update08,   iLength);
      VarArrayRedim(UV_vGubn_09,        iLength);
      VarArrayRedim(UV_vDat_Update09,   iLength);
      VarArrayRedim(UV_vCod_Update09,   iLength);
      VarArrayRedim(UV_vGubn_Hsokun,    iLength);
      VarArrayRedim(UV_vDat_UpdateHS,   iLength);
      VarArrayRedim(UV_vCod_UpdateHS,   iLength);
   end;
end;

procedure TfrmLD1I11.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자료를 화면에 조회
   case DataCol of
       1 :  Value := GF_DateFormat(UV_vDat_bunju[DataRow-1]);
   end;

   UF_SawonChk(GV_sUserId);
end;

procedure TfrmLD1I11.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, i : Integer;
begin
   inherited;
   UV_sJisa := '';
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   UV_sJisa := UV_vCod_bjjs[NewRow-1];
   // Grid의 Row가 변경되면 Detail에 자료를 할당
   cmb01.itemindex := UF_ToItemindex(UV_vGubn_01[NewRow-1]);
   cmb02.itemindex := UF_ToItemindex(UV_vGubn_02[NewRow-1]);
   cmb03.itemindex := UF_ToItemindex(UV_vGubn_03[NewRow-1]);
   cmb04.itemindex := UF_ToItemindex(UV_vGubn_04[NewRow-1]);
   cmb05.itemindex := UF_ToItemindex(UV_vGubn_05[NewRow-1]);
   cmb06.itemindex := UF_ToItemindex(UV_vGubn_06[NewRow-1]);
   cmb07.itemindex := UF_ToItemindex(UV_vGubn_07[NewRow-1]);
   cmb08.itemindex := UF_ToItemindex(UV_vGubn_08[NewRow-1]);
   cmbHS.itemindex := UF_ToItemindex(UV_vGubn_HSokun[NewRow-1]);
   dat01.caption   := UV_vDat_update01[NewRow-1];
   dat02.caption   := UV_vDat_update02[NewRow-1];
   dat03.caption   := UV_vDat_update03[NewRow-1];
   dat04.caption   := UV_vDat_update04[NewRow-1];
   dat05.caption   := UV_vDat_update05[NewRow-1];
   dat06.caption   := UV_vDat_update06[NewRow-1];
   dat07.caption   := UV_vDat_update07[NewRow-1];
   dat08.caption   := UV_vDat_update08[NewRow-1];
   datHS.caption   := UV_vDat_updateHS[NewRow-1];
   cod01.caption   := UV_vCod_update01[NewRow-1];
   cod02.caption   := UV_vCod_update02[NewRow-1];
   cod03.caption   := UV_vCod_update03[NewRow-1];
   cod04.caption   := UV_vCod_update04[NewRow-1];
   cod05.caption   := UV_vCod_update05[NewRow-1];
   cod06.caption   := UV_vCod_update06[NewRow-1];
   cod07.caption   := UV_vCod_update07[NewRow-1];
   cod08.caption   := UV_vCod_update08[NewRow-1];
   codHS.caption   := UV_vCod_updateHS[NewRow-1];
   if cmb01.ItemIndex = 0 then cmb01.Color := clGray else cmb01.color := clWindow;
   if cmb02.ItemIndex = 0 then cmb02.Color := clGray else cmb02.color := clWindow;
   if cmb03.ItemIndex = 0 then cmb03.Color := clGray else cmb03.color := clWindow;
   if cmb04.ItemIndex = 0 then cmb04.Color := clGray else cmb04.color := clWindow;
   if cmb05.ItemIndex = 0 then cmb05.Color := clGray else cmb05.color := clWindow;
   if cmb06.ItemIndex = 0 then cmb06.Color := clGray else cmb06.color := clWindow;
   if cmb07.ItemIndex = 0 then cmb07.Color := clGray else cmb07.color := clWindow;
   if cmb08.ItemIndex = 0 then cmb08.Color := clGray else cmb08.color := clWindow;
   if cmbHS.ItemIndex = 0 then cmbHS.Color := clGray else cmbHS.color := clWindow;

   // 조직학 Display 관련..-----------------------------------------------------
   DEdt_gumgin.Text := UV_vDat_bunju[NewRow-1];
   CombDate.Clear;
   CombDate.Color := clWindow;

   qryCell.Active := False;
   qryCell.ParamByName('cod_bjjs').AsString := UV_vCod_bjjs[NewRow-1];
   qryCell.ParamByName('dat_gmgn').AsString := UV_vDat_bunju[NewRow-1];
   qryCell.Active := True;
   if qryCell.RecordCount > 0 then
   begin
      while not qryCell.Eof do
      begin
         if Trim(qryCell.FieldByName('dat_result').AsString) <> '' then
            CombDate.Items.Add(copy(qryCell.FieldByName('dat_result').AsString,1,4) + '-' +
                               copy(qryCell.FieldByName('dat_result').AsString,5,2) + '-' +
                               copy(qryCell.FieldByName('dat_result').AsString,7,2))
         else
            CombDate.Color := $00CECEFF;
         qryCell.Next;
      end;
      CombDate.ItemIndex := 0;
   end;
   //---------------------------------------------------------------------------
   //---------------------------------------------------------------------------   
   //part 1,2,3 [결과체크]
   UP_Hgulkwa_JisaChk(UV_vCod_bjjs[NewRow-1], UV_vDat_bunju[NewRow-1]);

   if UV_vDat_bunju[NewRow-1] >= '20120303' then
   begin
      if (ck01_12.Caption = 'Y') and (ck01_11.Caption = 'Y') and
         (ck01_43.Caption = 'Y') and (ck01_51.Caption = 'Y') and
         (ck01_61.Caption = 'Y') and (ck01_71.Caption = 'Y') then
      begin
         cmb01.Enabled := True; Part_01.Visible := False;
      end
      else
      begin
         cmb01.Enabled := False; Part_01.Visible := True;
      end;

      if (ck02_12.Caption = 'Y') and (ck02_11.Caption = 'Y') and
         (ck02_43.Caption = 'Y') and (ck02_51.Caption = 'Y') and
         (ck02_61.Caption = 'Y') and (ck02_71.Caption = 'Y') then
      begin
         cmb02.Enabled := True; Part_02.Visible := False;
      end
      else
      begin
         cmb02.Enabled := False; Part_02.Visible := True;
      end;

      if (ck03_12.Caption = 'Y') and (ck03_11.Caption = 'Y') and
         (ck03_43.Caption = 'Y') and (ck03_51.Caption = 'Y') and
         (ck03_61.Caption = 'Y') and (ck03_71.Caption = 'Y') then
      begin
         cmb03.Enabled := True; Part_03.Visible := False;
      end
      else
      begin
         cmb03.Enabled := False; Part_03.Visible := True;
      end;
   end;
   //---------------------------------------------------------------------------

   cmbJisaChange(self);

   UF_SawonChk(GV_sUserId);
end;

function TfrmLD1I11.UF_PartGubn;
var i : integer;
begin
   grp01.enabled := false;  grp02.enabled := false;  grp03.enabled := false;
   grp04.enabled := false;  grp05.enabled := false;  grp06.enabled := false;
   grp07.enabled := false;  grp08.enabled := false;
   grpHS.enabled := false;  btnSave.enabled := false;
   // 파트별 작업 가능케..
   with qryUser_priv do
   begin
      close;
      ParamByName('cod_jisa').AsString  := copy(cmbJisa.text,1,2);
      ParamByName('cod_sawon').AsString := GV_sUserId;
      open;
      if recordcount > 0 then
      begin
         for i := 0 to RecordCount - 1 do
         begin
            // 01
            if (FieldByName('cod_prog').AsString = 'LD1I02') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp01.enabled   := true;
               btnSave.enabled := true;
            end;
            // 02
            if (FieldByName('cod_prog').AsString = 'LD1I01') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp02.enabled   := true;
               btnSave.enabled := true;
            end;
            // 03
            if (FieldByName('cod_prog').AsString = 'LD1I04') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp03.enabled   := true;
               btnSave.enabled := true;
            end;
            // 04
            if (FieldByName('cod_prog').AsString = 'LD1I05') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp04.enabled   := true;
               btnSave.enabled := true;
            end;
            // 05
            if (FieldByName('cod_prog').AsString = 'LD1I06') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp05.enabled   := true;
               btnSave.enabled := true;
            end;
            // 06
            if (FieldByName('cod_prog').AsString = 'LD1I07') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp06.enabled   := true;
               btnSave.enabled := true;
            end;
            // 07
            if (FieldByName('cod_prog').AsString = 'LD1I03') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp07.enabled   := true;
               btnSave.enabled := true;
            end;
            // 08
            if (FieldByName('cod_prog').AsString = 'LD1I08') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grp08.enabled   := true;
               btnSave.enabled := true;
            end;
            // 혈액소견
            if (FieldByName('cod_prog').AsString = 'LD6I02') and (FieldBYName('ysno_read').AsString = 'N') then
            begin
               grpHS.enabled   := true;
               btnSave.enabled := true;
            end;
            next;
         end;
      end;
   end;
end;

procedure TfrmLD1I11.btnSaveClick(Sender: TObject);
var i : integer;
begin
   inherited;

   // 현재위치 추출
   // i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   dmComm.database.StartTransaction;
   try
      for i := 0 to grdMaster.Rows - 1 do
      begin
         if grp01.enabled and UV_Edit1  then  UP_Save('01',UF_ToItemindex(UV_vGubn_01[i]),i);
         if grp02.enabled and UV_Edit2  then  UP_Save('02',UF_ToItemindex(UV_vGubn_02[i]),i);
         if grp03.enabled and UV_Edit3  then  UP_Save('03',UF_ToItemindex(UV_vGubn_03[i]),i);
         if grp04.enabled and UV_Edit4  then  UP_Save('04',UF_ToItemindex(UV_vGubn_04[i]),i);
         if grp05.enabled and UV_Edit5  then  UP_Save('05',UF_ToItemindex(UV_vGubn_05[i]),i);
         if grp06.enabled and UV_Edit6  then  UP_Save('06',UF_ToItemindex(UV_vGubn_06[i]),i);
         if grp07.enabled and UV_Edit7  then  UP_Save('07',UF_ToItemindex(UV_vGubn_07[i]),i);
         if grp08.enabled and UV_Edit8  then  UP_Save('08',UF_ToItemindex(UV_vGubn_08[i]),i);
         if grpHS.enabled and UV_EditHS then  UP_Save('HS',UF_ToItemindex(UV_vGubn_HSokun[i]),i);
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

   // Save Mode & Msg
   UP_SetMode('SAVE');
   btnQuery.click;
end;

procedure TfrmLD1I11.UP_Save(part : string; itemindex : integer; iRow : integer);
var sUpdate, sSet, sWhere : String;
begin
   inherited;
   sUpdate := ''; sSet := ''; sWhere := '';
   with qryU_Hgulkwa_chk do
   begin
      Active  := False;
      sUpdate := ' UPDATE Hgulkwa_chk ';
      sWhere  := ' Where cod_bjjs = ''' + UV_vCod_bjjs[iRow] + '''';
      sWhere  := sWhere + ' and dat_bunju = ''' + UV_vDat_bunju[iRow] + '''';

      if part = 'HS' then
         sSet    := ' set gubn_Hsokun = ''' + UF_Check(itemindex) + ''','
      else
         sSet    := ' set gubn_'+ part +' = ''' + UF_Check(itemindex) + ''',';

      if part = '01' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update01[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '02' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update02[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '03' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update03[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '04' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update04[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '05' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update05[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '06' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update06[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '07' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update07[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '08' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update08[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = '09' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_update09[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;
      if part = 'HS' then
      begin
         sSet := sSet + ' dat_update'+ part +' = ''' + UV_vDat_updateHS[iRow] + ''',';
         sSet := sSet + ' cod_update'+ part +' = ''' + GV_sUserId + '''';
      end;

      sql.clear;
      sql.add(sUpdate + sSet + sWhere);
      Execsql;
      GP_query_log(qryU_Hgulkwa_chk, 'LD1I11', 'Q', 'N');

      if part = '01' then
      begin
         UV_vGubn_01[iRow]      := UF_Check(cmb01.itemindex);
         UV_vDat_update01[iRow] := GV_sToday;
         UV_vCod_update01[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '02' then
      begin
         UV_vGubn_02[iRow]      := UF_Check(cmb02.itemindex);
         UV_vDat_update02[iRow] := GV_sToday;
         UV_vCod_update02[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '03' then
      begin
         UV_vGubn_03[iRow]      := UF_Check(cmb03.itemindex);
         UV_vDat_update03[iRow] := GV_sToday;
         UV_vCod_update03[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '04' then
      begin
         UV_vGubn_04[iRow]      := UF_Check(cmb04.itemindex);
         UV_vDat_update04[iRow] := GV_sToday;
         UV_vCod_update04[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '05' then
      begin
         UV_vGubn_05[iRow]      := UF_Check(cmb05.itemindex);
         UV_vDat_update05[iRow] := GV_sToday;
         UV_vCod_update05[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '06' then
      begin
         UV_vGubn_06[iRow]      := UF_Check(cmb06.itemindex);
         UV_vDat_update06[iRow] := GV_sToday;
         UV_vCod_update06[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '07' then
      begin
         UV_vGubn_07[iRow]      := UF_Check(cmb07.itemindex);
         UV_vDat_update07[iRow] := GV_sToday;
         UV_vCod_update07[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '08' then
      begin
         UV_vGubn_08[iRow]      := UF_Check(cmb08.itemindex);
         UV_vDat_update08[iRow] := GV_sToday;
         UV_vCod_update08[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = 'HS' then
      begin
         UV_vGubn_HSokun[iRow]  := UF_Check(cmbHS.itemindex);
         UV_vDat_updateHS[iRow] := GV_sToday;
         UV_vCod_updateHS[iRow] := UF_SawonChk(GV_sUserId);
      end;
   end;
end;

function TfrmLD1I11.UF_SawonChk(sTemp : String) : String;
begin
   with qrySawon do
   begin
      close;
      ParamByName('cod_sawon').AsString := sTemp;
      open;

      if recordcount > 0 then
      begin
         Result := FieldBYName('desc_name').AsString;
         if (copy(FieldByName('cod_sawon').AsString,1,2) = '12') or
            (copy(FieldByName('cod_sawon').AsString,1,2) = '15') then
         begin
            if Trim(FieldByName('gubn_dept').AsString) = '' then
               btnSave.Enabled  := False
            else
               case FieldByName('gubn_dept').AsInteger of
                  10 : btnSave.Enabled := True;                           //여의도 분석부
                  12 : btnSave.Enabled := True;                           //본사   분석부
                  19 : btnSave.Enabled := True;                           //본사   19-진단검사의학과 (김소영 선생님)
                  else
                     btnSave.Enabled := False;
               end;
         end;
      end
      else Result := '';
   end;
end;

function TfrmLD1I11.UF_DateChk(sTemp : String) : String;
begin
   if sTemp = '19000101' then Result := ''
   else                       Result := sTemp;
end;


function TfrmLD1I11.UF_ToItemindex(sTemp : String) : integer;
begin
   if sTemp = 'Y' then Result := 1
   else                Result := 0;
end;

function TfrmLD1I11.UF_Check(iTemp : Integer) : String;
begin
   Result := 'N';
   if iTemp = 1 then Result := 'Y';
end;

function TfrmLD1I11.UF_NameToCode(sTemp : String) : String;
begin
   with qryNameToCode do
   begin
      close;
      ParamByName('desc_name').AsString := sTemp;
      open;
      Result := FieldByName('cod_sawon').AsString;
   end;
end;

procedure TfrmLD1I11.UP_Change(Sender: TObject);
var i : integer;
begin
  inherited;
  //현재위치 추출
  i := grdMaster.CurrentDataRow - 1;
  if sender = cmb01 then
  begin
     UV_Edit1 := True;
     UV_vGubn_01[i]      := UF_Check(cmb01.itemindex);
     UV_vDat_update01[i] := GV_sToday;
     UV_vCod_update01[i] := UF_SawonChk(GV_sUserId);
     dat01.caption := UV_vDat_update01[i];
     cod01.caption := UV_vCod_update01[i];
  end else
  if sender = cmb02 then
  begin
     UV_Edit2 := True;
     UV_vGubn_02[i]      := UF_Check(cmb02.itemindex);
     UV_vDat_update02[i] := GV_sToday;
     UV_vCod_update02[i] := UF_SawonChk(GV_sUserId);
     dat02.caption := UV_vDat_update02[i];
     cod02.caption := UV_vCod_update02[i];
  end else
  if sender = cmb03 then
  begin
     UV_Edit3 := True;
     UV_vGubn_03[i]      := UF_Check(cmb03.itemindex);
     UV_vDat_update03[i] := GV_sToday;
     UV_vCod_update03[i] := UF_SawonChk(GV_sUserId);
     dat03.caption := UV_vDat_update03[i];
     cod03.caption := UV_vCod_update03[i];
  end else
  if sender = cmb04 then
  begin
     UV_Edit4 := True;
     UV_vGubn_04[i]      := UF_Check(cmb04.itemindex);
     UV_vDat_update04[i] := GV_sToday;
     UV_vCod_update04[i] := UF_SawonChk(GV_sUserId);
     dat04.caption := UV_vDat_update04[i];
     cod04.caption := UV_vCod_update04[i];
  end else
  if sender = cmb05 then
  begin
     UV_Edit5 := True;
     UV_vGubn_05[i]      := UF_Check(cmb05.itemindex);
     UV_vDat_update05[i] := GV_sToday;
     UV_vCod_update05[i] := UF_SawonChk(GV_sUserId);
     dat05.caption := UV_vDat_update05[i];
     cod05.caption := UV_vCod_update05[i];
  end else
  if sender = cmb06 then
  begin
     UV_Edit6 := True;
     UV_vGubn_06[i]      := UF_Check(cmb06.itemindex);
     UV_vDat_update06[i] := GV_sToday;
     UV_vCod_update06[i] := UF_SawonChk(GV_sUserId);
     dat06.caption := UV_vDat_update06[i];
     cod06.caption := UV_vCod_update06[i];
  end else
  if sender = cmb07 then
  begin
     UV_Edit7 := True;
     UV_vGubn_07[i]      := UF_Check(cmb07.itemindex);
     UV_vDat_update07[i] := GV_sToday;
     UV_vCod_update07[i] := UF_SawonChk(GV_sUserId);
     dat07.caption := UV_vDat_update07[i];
     cod07.caption := UV_vCod_update07[i];
  end else
  if sender = cmb08 then
  begin
     UV_Edit8 := True;
     UV_vGubn_08[i]      := UF_Check(cmb08.itemindex);
     UV_vDat_update08[i] := GV_sToday;
     UV_vCod_update08[i] := UF_SawonChk(GV_sUserId);
     dat08.caption := UV_vDat_update08[i];
     cod08.caption := UV_vCod_update08[i];
  end else
  if sender = cmbHS then
  begin
     if (cmb01.itemindex = 0) or (cmb02.itemindex = 0) or (cmb03.itemindex = 0) or
        (cmb04.itemindex = 0) or (cmb05.itemindex = 0) or
        (cmb07.itemindex = 0) or (cmb08.itemindex = 0) then
     begin
        showmessage('혈액결과처리가 완료되지 않았습니다.');
        cmbHS.itemindex := 0;
     end else
     begin
        UV_EditHS := True;
        UV_vGubn_HSokun[i]    := UF_Check(cmbHS.itemindex);
        UV_vDat_updateHS[i]   := GV_sToday;
        UV_vCod_updateHS[i]   := UF_SawonChk(GV_sUserId);
        datHS.caption := UV_vDat_updateHS[i];
        codHS.caption := UV_vCod_updateHS[i];
     end;
  end;
end;

procedure TfrmLD1I11.cmbJisaChange(Sender: TObject);
begin
   inherited;
   if (copy(cmbJisa.Text,1,2) = copy(GV_sUserId,1,2)) and (UV_sJisa = copy(GV_sUserId,1,2)) then
   begin
      pnlDetail.Enabled := True;
      btnSave.Enabled   := True;
      UF_PartGubn;
   end
   else
   begin
      pnlDetail.Enabled := False;
      btnSave.Enabled   := False;
   end;
end;

procedure TfrmLD1I11.UP_Hgulkwa_JisaChk(sCod_Bjjs, sDat_bunju : string);
begin
   ck01_11.Caption := 'N'; dat01_11.Caption := ''; cod01_11.Caption := '';
   ck01_12.Caption := 'N'; dat01_12.Caption := ''; cod01_12.Caption := '';
   ck01_43.Caption := 'N'; dat01_43.Caption := ''; cod01_43.Caption := '';
   ck01_51.Caption := 'N'; dat01_51.Caption := ''; cod01_51.Caption := '';
   ck01_61.Caption := 'N'; dat01_61.Caption := ''; cod01_61.Caption := '';
   ck01_71.Caption := 'N'; dat01_71.Caption := ''; cod01_71.Caption := '';

   ck02_11.Caption := 'N'; dat02_11.Caption := ''; cod02_11.Caption := '';
   ck02_12.Caption := 'N'; dat02_12.Caption := ''; cod02_12.Caption := '';
   ck02_43.Caption := 'N'; dat02_43.Caption := ''; cod02_43.Caption := '';
   ck02_51.Caption := 'N'; dat02_51.Caption := ''; cod02_51.Caption := '';
   ck02_61.Caption := 'N'; dat02_61.Caption := ''; cod02_61.Caption := '';
   ck02_71.Caption := 'N'; dat02_71.Caption := ''; cod02_71.Caption := '';

   ck03_11.Caption := 'N'; dat03_11.Caption := ''; cod03_11.Caption := '';
   ck03_12.Caption := 'N'; dat03_12.Caption := ''; cod03_12.Caption := '';
   ck03_43.Caption := 'N'; dat03_43.Caption := ''; cod03_43.Caption := '';
   ck03_51.Caption := 'N'; dat03_51.Caption := ''; cod03_51.Caption := '';
   ck03_61.Caption := 'N'; dat03_61.Caption := ''; cod03_61.Caption := '';
   ck03_71.Caption := 'N'; dat03_71.Caption := ''; cod03_71.Caption := '';

   qryHgulkwa_EtcChk.Active := False;
   qryHgulkwa_EtcChk.ParamByName('cod_bjjs').AsString  := sCod_Bjjs;
   qryHgulkwa_EtcChk.ParamByName('dat_bunju').AsString := sDat_bunju;
   qryHgulkwa_EtcChk.Active := True;

   if qryHgulkwa_EtcChk.RecordCount > 0 then
   begin
      //------------------------------------------------------------------------
      ck01_11.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_11').AsString;
      ck01_12.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_12').AsString;
      ck01_43.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_43').AsString;
      ck01_51.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_51').AsString;
      ck01_61.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_61').AsString;
      ck01_71.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_01_71').AsString;

      ck02_11.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_11').AsString;
      ck02_12.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_12').AsString;
      ck02_43.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_43').AsString;
      ck02_51.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_51').AsString;
      ck02_61.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_61').AsString;
      ck02_71.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_02_71').AsString;

      ck03_11.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_11').AsString;
      ck03_12.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_12').AsString;
      ck03_43.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_43').AsString;
      ck03_51.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_51').AsString;
      ck03_61.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_61').AsString;
      ck03_71.Caption := qryHgulkwa_EtcChk.FieldByName('Chk_03_71').AsString;
      //------------------------------------------------------------------------
      dat01_11.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_11').AsString,1,8));
      dat01_12.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_12').AsString,1,8));
      dat01_43.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_43').AsString,1,8));
      dat01_51.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_51').AsString,1,8));
      dat01_61.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_61').AsString,1,8));
      dat01_71.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_71').AsString,1,8));

      dat02_11.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_11').AsString,1,8));
      dat02_12.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_12').AsString,1,8));
      dat02_43.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_43').AsString,1,8));
      dat02_51.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_51').AsString,1,8));
      dat02_61.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_61').AsString,1,8));
      dat02_71.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_71').AsString,1,8));

      dat03_11.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_11').AsString,1,8));
      dat03_12.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_12').AsString,1,8));
      dat03_43.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_43').AsString,1,8));
      dat03_51.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_51').AsString,1,8));
      dat03_61.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_61').AsString,1,8));
      dat03_71.Caption := UF_DateChk(copy(qryHgulkwa_EtcChk.FieldByName('Dat_01_71').AsString,1,8));
      //------------------------------------------------------------------------
      cod01_11.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_11').AsString);
      cod01_12.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_12').AsString);
      cod01_43.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_43').AsString);
      cod01_51.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_51').AsString);
      cod01_61.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_61').AsString);
      cod01_71.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_01_71').AsString);

      cod02_11.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_11').AsString);
      cod02_12.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_12').AsString);
      cod02_43.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_43').AsString);
      cod02_51.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_51').AsString);
      cod02_61.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_61').AsString);
      cod02_71.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_02_71').AsString);

      cod03_11.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_11').AsString);
      cod03_12.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_12').AsString);
      cod03_43.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_43').AsString);
      cod03_51.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_51').AsString);
      cod03_61.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_61').AsString);
      cod03_71.Caption := UF_SawonChk(qryHgulkwa_EtcChk.FieldByName('Cod_03_71').AsString);
      //------------------------------------------------------------------------
   end;
end;


end.
