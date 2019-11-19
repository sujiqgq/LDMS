//==============================================================================
// 프로그램 설명 : [센터]혈액결과,소견체크
// 시스템        : LDMS
// 수정일자      : 2012.12.21
// 수정자        : 유동구
// 참고사항      : [센터]자체검사 실시관련 추가...(qryHgulkwa_EtcChk)
//==============================================================================
unit LD1I15;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, StdCtrls, ExtCtrls, Buttons, ComCtrls, ToolWin, Mask, DateEdit,
  Grids_ts, TSGrid, Db, DBTables;

type
  TfrmLD1I15 = class(TfrmSingle)
    pnlCond: TPanel;
    btnDate: TSpeedButton;
    Label2: TLabel;
    btnDATE_1: TSpeedButton;
    Label7: TLabel;
    mskST_Date: TDateEdit;
    mskED_Date: TDateEdit;
    grdMaster: TtsGrid;
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
    qryU_Hgulkwa_EtcChk: TQuery;
    qryNameToCode: TQuery;
    Label39: TLabel;
    Label45: TLabel;
    qryHgulkwa_EtcChk: TQuery;
    procedure UP_Click(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnSaveClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
  private
    { Private declarations }
    UV_vCod_bjjs, UV_vDat_bunju,
    UV_vChk_01, UV_vDat_01, UV_vCod_01,
    UV_vChk_02, UV_vDat_02, UV_vCod_02,
    UV_vChk_03, UV_vDat_03, UV_vCod_03 : Variant;

    UV_sJisa : String;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_GridSet;
    procedure UP_Save(part : string; itemindex : integer; iRow : integer);

    function  UF_PartGubn : Boolean;
    function  UF_Condition : Boolean;
    function  UF_Check(iTemp : Integer) : String;
    function  UF_ToItemindex(sTemp : String) : integer;
    function  UF_SawonChk(sTemp : String) : String;
    function  UF_NameToCode(sTemp : String) : String;

  public
    UV_Edit1, UV_Edit2, UV_Edit3 : boolean;
    { Public declarations }
  end;

var
  frmLD1I15: TfrmLD1I15;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD1I15.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   // Button Click시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = btnDate then GF_CalendarClick(mskST_Date)
   else if Sender = btnDATE_1 then GF_CalendarClick(mskED_Date);
end;

function TfrmLD1I15.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (mskSt_Date.Text = '') or (mskEd_Date.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD1I15.btnQueryClick(Sender: TObject);
var i, index : integer;
begin
  inherited;
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   UV_Edit1 := False; UV_Edit2 := False; UV_Edit3 := False;

   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('sdat_bunju').AsString := mskST_date.text;
      ParamByName('edat_bunju').AsString := mskED_date.text;
      open;
      if recordcount > 0 then
      begin
         UP_VarMemSet(RecordCount-1);
         GP_query_log(qryHgulkwa_EtcChk, 'LD1I15', 'Q', 'Y');

         for i := 0 to RecordCount - 1 do
         begin
            UV_vCod_bjjs[index]  := FieldByName('cod_bjjs').AsString;
            UV_vDat_bunju[index] := FieldBYName('dat_bunju').AsString;
            UV_vChk_01[index]    := FieldByName('Chk_01_' + Copy(GV_sUserId,1,2)).AsString;
            UV_vChk_02[index]    := FieldByName('Chk_02_' + Copy(GV_sUserId,1,2)).AsString;
            UV_vChk_03[index]    := FieldByName('Chk_03_' + Copy(GV_sUserId,1,2)).AsString;

            UV_vDat_01[index]    := FieldByName('Dat_01_' + Copy(GV_sUserId,1,2)).AsString;
            UV_vDat_02[index]    := FieldByName('Dat_02_' + Copy(GV_sUserId,1,2)).AsString;
            UV_vDat_03[index]    := FieldByName('Dat_03_' + Copy(GV_sUserId,1,2)).AsString;

            UV_vCod_01[index]    := UF_SawonChk(FieldByName('cod_01_' + Copy(GV_sUserId,1,2)).AsString);
            UV_vCod_02[index]    := UF_SawonChk(FieldByName('cod_02_' + Copy(GV_sUserId,1,2)).AsString);
            UV_vCod_03[index]    := UF_SawonChk(FieldByName('cod_03_' + Copy(GV_sUserId,1,2)).AsString);

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
end;

procedure TfrmLD1I15.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      Col[1].Alignment  := taCenter;
      Col[2].Alignment  := taCenter;
      Col[4].Alignment  := taCenter;
      Col[5].Alignment  := taCenter;
      Col[7].Alignment  := taCenter;
      Col[8].Alignment  := taCenter;
      Col[10].Alignment := taCenter;
   end;
end;

procedure TfrmLD1I15.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;
   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then  sJisa := Copy(GV_sUserId,1,2)
   else                     sJisa := GV_sJisa;

   // 환경설정
   UP_VarMemSet(0);

   UF_PartGubn;
end;

procedure TfrmLD1I15.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vChk_01    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_01    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_01    := VarArrayCreate([0, 0], varOleStr);
      UV_vChk_02    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_02    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_02    := VarArrayCreate([0, 0], varOleStr);
      UV_vChk_03    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_03    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_03    := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,  iLength);
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vChk_01,    iLength);
      VarArrayRedim(UV_vDat_01,    iLength);
      VarArrayRedim(UV_vCod_01,    iLength);
      VarArrayRedim(UV_vChk_02,    iLength);
      VarArrayRedim(UV_vDat_02,    iLength);
      VarArrayRedim(UV_vCod_02,    iLength);
      VarArrayRedim(UV_vChk_03,    iLength);
      VarArrayRedim(UV_vDat_03,    iLength);
      VarArrayRedim(UV_vCod_03,    iLength);
   end;
end;

procedure TfrmLD1I15.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자료를 화면에 조회
   case DataCol of
       1 :  Value := GF_DateFormat(UV_vDat_bunju[DataRow-1]);
       2 :  Value := UV_vChk_01[DataRow-1];
       3 :  Value := UV_vDat_01[DataRow-1];
       4 :  Value := UV_vCod_01[DataRow-1];
       5 :  Value := UV_vChk_02[DataRow-1];
       6 :  Value := UV_vDat_02[DataRow-1];
       7 :  Value := UV_vCod_02[DataRow-1];
       8 :  Value := UV_vChk_03[DataRow-1];
       9 :  Value := UV_vDat_03[DataRow-1];
      10 :  Value := UV_vCod_03[DataRow-1];
   end;

   UF_SawonChk(GV_sUserId);   
end;

procedure TfrmLD1I15.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, i : Integer;
begin
   inherited;
   UV_sJisa := '';
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   UV_sJisa := UV_vCod_bjjs[NewRow-1];
   // Grid의 Row가 변경되면 Detail에 자료를 할당
   cmb01.itemindex := UF_ToItemindex(UV_vChk_01[NewRow-1]);
   cmb02.itemindex := UF_ToItemindex(UV_vChk_02[NewRow-1]);
   cmb03.itemindex := UF_ToItemindex(UV_vChk_03[NewRow-1]);

   dat01.caption   := UV_vDat_01[NewRow-1];
   dat02.caption   := UV_vDat_02[NewRow-1];
   dat03.caption   := UV_vDat_03[NewRow-1];

   cod01.caption   := UV_vCod_01[NewRow-1];
   cod02.caption   := UV_vCod_02[NewRow-1];
   cod03.caption   := UV_vCod_03[NewRow-1];

   if cmb01.ItemIndex = 0 then cmb01.Color := clGray else cmb01.color := clWindow;
   if cmb02.ItemIndex = 0 then cmb02.Color := clGray else cmb02.color := clWindow;
   if cmb03.ItemIndex = 0 then cmb03.Color := clGray else cmb03.color := clWindow;
end;

function TfrmLD1I15.UF_PartGubn;
var i : integer;
begin
   grp01.enabled := false;  grp02.enabled := false;  grp03.enabled := false;
   btnSave.enabled := false;
   // 파트별 작업 가능케..
   with qryUser_priv do
   begin
      close;
      ParamByName('cod_jisa').AsString  := copy(GV_sUserId,1,2);
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

            next;
         end;
      end;
   end;
end;

procedure TfrmLD1I15.btnSaveClick(Sender: TObject);
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
         if grp01.enabled and UV_Edit1  then  UP_Save('01',UF_ToItemindex(UV_vChk_01[i]),i);
         if grp02.enabled and UV_Edit2  then  UP_Save('02',UF_ToItemindex(UV_vChk_02[i]),i);
         if grp03.enabled and UV_Edit3  then  UP_Save('03',UF_ToItemindex(UV_vChk_03[i]),i);
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

procedure TfrmLD1I15.UP_Save(part : string; itemindex : integer; iRow : integer);
var sUpdate, sSet, sWhere : String;
begin
   inherited;
   sUpdate := ''; sSet := 'SET '; sWhere := '';
   with qryU_Hgulkwa_EtcChk do
   begin
      Active  := False;
      sUpdate := ' UPDATE Hgulkwa_EtcChk ';
      sWhere  := ' Where cod_bjjs = ''' + UV_vCod_bjjs[iRow] + '''';
      sWhere  := sWhere + ' and dat_bunju = ''' + UV_vDat_bunju[iRow] + '''';

      if part = '01' then
      begin
         sSet := sSet + ' Chk_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + UV_vChk_01[iRow] + ''',';
         sSet := sSet + ' dat_'+ part + '_' + copy(GV_sUserId,1,2) +' = getdate(),';
         sSet := sSet + ' cod_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + GV_sUserId + '''';
      end;

      if part = '02' then
      begin
         sSet := sSet + ' Chk_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + UV_vChk_02[iRow] + ''',';
         sSet := sSet + ' dat_'+ part + '_' + copy(GV_sUserId,1,2) +' = getdate(),';
         sSet := sSet + ' cod_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + GV_sUserId + '''';
      end;

      if part = '03' then
      begin
         sSet := sSet + ' Chk_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + UV_vChk_03[iRow] + ''',';
         sSet := sSet + ' dat_'+ part + '_' + copy(GV_sUserId,1,2) +' = getdate(),';
         sSet := sSet + ' cod_'+ part + '_' + copy(GV_sUserId,1,2) +' = ''' + GV_sUserId + '''';
      end;

      sql.clear;
      sql.add(sUpdate + sSet + sWhere);
      Execsql;

      GP_query_log(qryU_Hgulkwa_EtcChk, 'LD1I15', 'Q', 'N');

      if part = '01' then
      begin
         UV_vChk_01[iRow] := UF_Check(cmb01.itemindex);
         UV_vDat_01[iRow] := GV_sToday;
         UV_vCod_01[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '02' then
      begin
         UV_vChk_02[iRow] := UF_Check(cmb02.itemindex);
         UV_vDat_02[iRow] := GV_sToday;
         UV_vCod_02[iRow] := UF_SawonChk(GV_sUserId);
      end else
      if part = '03' then
      begin
         UV_vChk_03[iRow] := UF_Check(cmb03.itemindex);
         UV_vDat_03[iRow] := GV_sToday;
         UV_vCod_03[iRow] := UF_SawonChk(GV_sUserId);
      end;
   end;
end;

function TfrmLD1I15.UF_SawonChk(sTemp : String) : String;
begin
   Result := '';

   with qrySawon do
   begin
      close;
      ParamByName('cod_sawon').AsString := sTemp;
      open;

      if recordcount > 0 then Result := FieldBYName('desc_name').AsString;
   end;
end;

function TfrmLD1I15.UF_ToItemindex(sTemp : String) : integer;
begin
   if sTemp = 'Y' then Result := 1
   else                Result := 0;
end;

function TfrmLD1I15.UF_Check(iTemp : Integer) : String;
begin
   Result := 'N';
   if iTemp = 1 then Result := 'Y';
end;

function TfrmLD1I15.UF_NameToCode(sTemp : String) : String;
begin
   with qryNameToCode do
   begin
      close;
      ParamByName('desc_name').AsString := sTemp;
      open;
      Result := FieldByName('cod_sawon').AsString;
   end;
end;

procedure TfrmLD1I15.UP_Change(Sender: TObject);
var i : integer;
begin
  inherited;
  //현재위치 추출
  i := grdMaster.CurrentDataRow - 1;
  if sender = cmb01 then
  begin
     UV_Edit1 := True;
     UV_vChk_01[i] := UF_Check(cmb01.itemindex);
     UV_vDat_01[i] := GV_sToday;
     UV_vCod_01[i] := UF_SawonChk(GV_sUserId);
     dat01.caption := UV_vDat_01[i];
     cod01.caption := UV_vCod_01[i];
  end else
  if sender = cmb02 then
  begin
     UV_Edit2 := True;
     UV_vChk_02[i] := UF_Check(cmb02.itemindex);
     UV_vDat_02[i] := GV_sToday;
     UV_vCod_02[i] := UF_SawonChk(GV_sUserId);
     dat02.caption := UV_vDat_02[i];
     cod02.caption := UV_vCod_02[i];
  end else
  if sender = cmb03 then
  begin
     UV_Edit3 := True;
     UV_vChk_03[i] := UF_Check(cmb03.itemindex);
     UV_vDat_03[i] := GV_sToday;
     UV_vCod_03[i] := UF_SawonChk(GV_sUserId);
     dat03.caption := UV_vDat_03[i];
     cod03.caption := UV_vCod_03[i];
  end;
end;

end.
