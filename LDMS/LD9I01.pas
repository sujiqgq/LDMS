//==============================================================================
// 프로그램 설명 : 조회코드 추가/삭제
// 시스템        : LDMS
// 수정일자      : 2015.02.28
// 수정자        : 곽윤설
//==============================================================================

unit LD9I01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, Gauges;
type
  TfrmLD9I01 = class(TfrmSingle)
    pnlCond: TPanel;
    grdMaster: TtsGrid;
    Panel2: TPanel;
    Panel3: TPanel;
    qryHangmok: TQuery;
    qryHM_List: TQuery;
    btn_chuga: TButton;
    btn_delete: TButton;
    CB_Program: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    CB_Group: TComboBox;
    grdProgram: TtsGrid;
    qryGroup: TQuery;
    qrySawon: TQuery;
    btnPSave: TBitBtn;
    qryIHM_List: TQuery;
    qryUHM_List: TQuery;
    qryU_PosHM: TQuery;
    Edt_Create: TEdit;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure btn_chugaClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdProgramRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btn_deleteClick(Sender: TObject);
    procedure grdProgramCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnPSaveClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_Part, UV_Cod_HM, UV_Desc_HM, UV_Dat_Apply : Variant;
    UV_vPart, UV_vCod_HM, UV_vDesc_HM, UV_vDat_Insert, UV_vCod_Insert : Variant;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    function  UP_ProgList(vGubn : Variant) : Variant;
    function  UP_SawonName(sCod_sawon : String) : String;
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_ProgVarMemSet(iLength : Integer);
    public
    { Public declarations }
  end;

var
  frmLD9I01: TfrmLD9I01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD9I01.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      // Column Title
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
      Col[4].Alignment := taCenter;
   end;

   with grdProgram do
   begin
      Rows := 0;
      // Column Title
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
      Col[4].Alignment := taCenter;
      Col[5].Alignment := taCenter;
   end;
end;

procedure TfrmLD9I01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_Part      := VarArrayCreate([0, 0], varOleStr);
      UV_Cod_HM    := VarArrayCreate([0, 0], varOleStr);
      UV_Desc_HM   := VarArrayCreate([0, 0], varOleStr);
      UV_Dat_Apply := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_Part     , iLength);
      VarArrayRedim(UV_Cod_HM   , iLength);
      VarArrayRedim(UV_Desc_HM  , iLength);
      VarArrayRedim(UV_Dat_Apply, iLength);
   end;
end;

procedure TfrmLD9I01.UP_ProgVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vPart       := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_HM     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_HM    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Insert := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Insert := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vPart      , iLength);
      VarArrayRedim(UV_vCod_HM    , iLength);
      VarArrayRedim(UV_vDesc_HM   , iLength);
      VarArrayRedim(UV_vDat_Insert, iLength);
      VarArrayRedim(UV_vCod_Insert, iLength);
   end;
end;

procedure TfrmLD9I01.FormCreate(Sender: TObject);
   var Index, i : Integer;
begin
   inherited;
   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);
   UP_ProgVarMemSet(0);
   with qryGroup do
   begin
      Active := False;
      ParamByName('PROGRAM_ID').AsString := 'LD9Q01';
      Active := True;
      if RecordCount > 0 then
      begin
         while not Eof do
         begin
             CB_Group.Items.Add(FieldByName('GROUP_HM').AsString);
             Next;
         end; //End of while
      end; //End of RecordCount
      Active := False;
   end; //End of with

   // Grid 초기화
   grdMaster.Rows  := 0;
   // Query 수행 & 배열에 저장
   Index := 0;
   with qryHangmok do
   begin
      Active := False;
      Active := True;

      if RecordCount > 0 then
      begin
         UP_VarMemSet(RecordCount-1);

         for i := 0 to RecordCount - 1 do
         begin
            UV_Part[i]      := FieldByName('gubn_part').AsString;
            UV_Cod_HM[i]    := FieldByName('cod_hm').AsString;
            UV_Desc_HM[i]   := FieldByName('desc_kor').AsString;
            UV_Dat_Apply[i] := FieldByName('dat_apply').AsString;
            Next;
            Inc(Index);
         end;
         // Table과 Disconnect
         Active := False;
      end;
   end;
   // Grid에 자료를 할당
   grdMaster.Rows := Index;
   CB_Program.ItemIndex := 0;
   CB_Group.ItemIndex := 0;
end;

procedure TfrmLD9I01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_Part[DataRow-1];
      2 : Value := UV_Cod_HM[DataRow-1];
      3 : Value := UV_Desc_HM[DataRow-1];
      4 : Value := UV_Dat_Apply[DataRow-1];
   end;
end;

procedure TfrmLD9I01.grdProgramCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vPart[DataRow-1];
      2 : Value := UV_vCod_HM[DataRow-1];
      3 : Value := UV_vDesc_HM[DataRow-1];
      4 : Value := UV_vDat_Insert[DataRow-1];
      5 : Value := UV_vCod_Insert[DataRow-1];
   end;
end;

procedure TfrmLD9I01.btnQueryClick(Sender: TObject);
var i, vIndex : Integer;
begin
   inherited;
   // Grid 초기화
   grdProgram.Rows := 0;

   // Query 수행 & 배열에 저장
   vIndex := 0;
   with qryHM_List do
   begin
      Active := False;
      ParamByName('PROGRAM_ID').AsString := UP_ProgList(CB_Program.ItemIndex + 1);
      ParamByName('GROUP_HM').AsString := Trim(CB_Group.Text);
      Active := True;
      if RecordCount > 0 then
      begin
         UP_ProgVarMemSet(RecordCount-1);
         for i := 0 to RecordCount - 1 do
         begin
            UV_vPart[i]       := FieldByName('gubn_part').AsString;
            UV_vCod_HM[i]     := FieldByName('cod_hm').AsString;
            UV_vDesc_HM[i]    := FieldByName('desc_kor').AsString;
            UV_vDat_Insert[i] := FieldByName('dat_insert').AsString;
            UV_vCod_Insert[i] := UP_SawonName(FieldByName('cod_insert').AsString);
            Next;
            Inc(vIndex);
         end;
         // Table과 Disconnect
         Active := False;
      end;
   end;
   // Grid에 자료를 할당
   grdProgram.Rows := vIndex;
end;

procedure TfrmLD9I01.UP_Change(Sender: TObject);
begin
   inherited;
   if Sender = CB_Program then
   begin
      CB_Group.Items.Clear;
      with qryGroup do
      begin
         Active := False;
         ParamByName('PROGRAM_ID').AsString := UP_ProgList(CB_Program.ItemIndex + 1);
         Active := True;
         if RecordCount > 0 then
         begin
            while not Eof do
            begin
                CB_Group.Items.Add(FieldByName('GROUP_HM').AsString);
                Next;
            end; //End of while
         end; //End of RecordCount
         Active := False;
      end; //End of with
      CB_Group.ItemIndex := 0;
   end;
end;

procedure TfrmLD9I01.btn_chugaClick(Sender: TObject);
var i,j,Cnt:integer;
    Change : Boolean;
begin
   inherited;
   for i := 1 to grdMaster.Rows do
   begin
      Change := False;
      if grdMaster.RowSelected[i] = True then Change := True;
      for j := 1 to grdProgram.Rows do
      begin
         if grdMaster.Cell[1, i] = UV_vCod_HM[j-1] then
         begin
            Change := False;
            break;
         end;
      end;
      if Change then
      begin
           UP_ProgVarMemSet(grdProgram.Rows);
           UV_vPART[grdProgram.Rows]       := grdMaster.Cell[1, i];
           UV_vCOD_HM[grdProgram.Rows]     := grdMaster.Cell[2, i];
           UV_vDESC_HM[grdProgram.Rows]    := grdMaster.Cell[3, i];
           UV_vDAT_INSERT[grdProgram.Rows] := GV_sToday;
           UV_vCOD_INSERT[grdProgram.Rows] := UP_SawonName(GV_sUserId);
           grdProgram.Rows := grdProgram.Rows + 1;
           grdProgram.CurrentDataRow := grdProgram.Rows;

           with qryIHM_List do
           begin
              Active := False;
              ParamByName('PROGRAM_ID').AsString := UP_ProgList(CB_Program.ItemIndex + 1);
              ParamByName('COD_HM').AsString     := grdMaster.Cell[2, i];
              ParamByName('GUBN_PART').AsString  := grdMaster.Cell[1, i];
              if (edt_create.enabled) and (Trim(edt_create.Text) <> '') then
                   ParamByName('GROUP_HM').AsString := Trim(edt_create.Text)
              else ParamByName('GROUP_HM').AsString := Trim(CB_Group.Text);
              ParamByName('DAT_INSERT').AsString := GV_sToday;
              ParamByName('COD_INSERT').AsString := GV_sUserId;
              ExecSQL;
              Active := False;
           end;
      end;
      if grdProgram.Rows > 0 then grdProgram.Enabled := true
      else grdProgram.Enabled := false;
   end;
end;

procedure TfrmLD9I01.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
  inherited;
  if NewRow = 0 then exit;
  GP_SetDataCnt(grdMaster);
end;

procedure TfrmLD9I01.grdProgramRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
  inherited;
  if NewRow = 0 then exit;
  GP_SetDataCnt(grdProgram);
end;

procedure TfrmLD9I01.btn_deleteClick(Sender: TObject);
var i, j, k : Integer;
begin
  inherited;
  UP_ProgVarMemSet(grdProgram.Rows);
  for  j := 1 to grdProgram.Rows do
     if grdProgram.RowSelected[j] = True then
     begin
         i := grdProgram.CurrentDataRow-1;
         with qryUHM_List do
         begin
            Active := False;
            ParamByName('DAT_UPDATE').AsString := GV_sToday;
            ParamByName('COD_UPDATE').AsString := GV_sUserId;
            ParamByName('PROGRAM_ID').AsString := UP_ProgList(CB_Program.ItemIndex + 1);
            if (edt_create.enabled) and (Trim(edt_create.Text) <> '') then
                 ParamByName('GROUP_HM').AsString := Trim(edt_create.Text)
            else ParamByName('GROUP_HM').AsString := Trim(CB_Group.Text);
            ParamByName('COD_HM').AsString     := UV_vCOD_HM[i];
            ExecSQL;
            Active := False;
         end;
         UV_vPart[i]       := '';
         UV_vCod_HM[i]     := '';
         UV_vDesc_HM[i]    := '';
         UV_vDat_Insert[i] := '';
         UV_vCod_Insert[i] := '';
         grdProgram.RowVisible[i+1] := False;
     end;
end;

function  TfrmLD9I01.UP_ProgList(vGubn : Variant) : Variant;
begin
   Result := '';
   case vGubn of
      1 : Result := 'LD9Q01';
      2 : Result := 'LD9Q02';
      3 : Result := 'LD9Q03';
      4 : Result := 'LD9Q04';
   end;
end;

function  TfrmLD9I01.UP_SawonName(sCod_sawon : String) : String;
begin
   Result := '';
   with qrySawon do
   begin
      Active := False;
      ParamByName('cod_sawon').AsString := sCod_sawon;
      Active := True;
      if RecordCount > 0 then
      begin
         while not Eof do
         begin
            Result := FieldByName('desc_name').AsString;
            next;
         end;
      end;
      Active := False;
   end;
end;

procedure TfrmLD9I01.btnPSaveClick(Sender: TObject);
var i, j : Integer;
begin
  inherited;
  UP_ProgVarMemSet(grdProgram.Rows);
  j := 1;
  for i := 1 to grdProgram.Rows do
  begin
     if UV_vCOD_HM[i] <> '' then
     begin
        with qryU_posHM do
        begin
           Active := False;
           ParamByName('POS_HM').AsString     := IntToStr(j);
           ParamByName('PROGRAM_ID').AsString := UP_ProgList(CB_Program.ItemIndex + 1);
           if (Edt_Create.Enabled) and (Trim(Edt_Create.Text) <> '') then
                ParamByName('GROUP_HM').AsString := Trim(Edt_Create.Text)
           else ParamByName('GROUP_HM').AsString := Trim(CB_Group.Text);
           ParamByName('COD_HM').AsString     := UV_vCOD_HM[i];
           ExecSQL;
           Active := False;
        end;
        inc(j);
     end;
  end;
  UP_Change(CB_Program);

  Edt_Create.Enabled := False;
  Edt_Create.Text    := '';
  Edt_Create.Color   := clBtnFace;
  CB_Group.Enabled   := True;

  Showmessage('저장완료');
end;

procedure TfrmLD9I01.btnInsertClick(Sender: TObject);
begin
  inherited;
  Edt_Create.Enabled := True;
  Edt_Create.Color   := clWindow;
  CB_Group.Enabled   := False;
  grdProgram.Rows    := 0;
end;

procedure TfrmLD9I01.FormActivate(Sender: TObject);
begin
  inherited;
  if btnPSave.Enabled then
  begin
     btn_Chuga.Enabled  := True;
     btn_Delete.Enabled := True;
  end
  else
  begin
     btn_Chuga.Enabled  := False;
     btn_Delete.Enabled := False;
  end;
end;

procedure TfrmLD9I01.btnCancelClick(Sender: TObject);
begin
  inherited;
  Edt_Create.Enabled := False;
  Edt_Create.Color   := clBtnFace;
  CB_Group.Enabled   := True;
  grdProgram.Rows    := 0;
end;

end.
