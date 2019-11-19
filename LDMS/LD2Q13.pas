//==============================================================================
// 프로그램 설명 : 분주번호 조회
// 시스템        : 통합검진
// 수정일자      : 2015.11.02
// 수정자        : 박수지 
// 수정내용      : 신규개발
// 참고사항      :
//==============================================================================
unit LD2Q13;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;

type
  TfrmLD2Q13 = class(TfrmSingle)
    pnlCond: TPanel;
    msk_date: TDateEdit;
    asdf: TLabel;
    cmbJisa: TComboBox;
    Label12: TLabel;
    qryHul: TQuery;
    qryGumgin: TQuery;
    Gauge1: TGauge;
    grdMaster: TtsGrid;
    Label2: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryJangbi: TQuery;
    qryPf_hm: TQuery;
    btnDate: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
              DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
   // procedure UP_KeyUp(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure btnQuitClick(Sender: TObject);
//    procedure UP_Change(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
 private
    { Private declarations }
    UV_vDat_bunju, UV_vNum_bunju, UV_vCod_jisa, UV_vNum_samp, UV_vName,
    UV_vC078, UV_vE069, UV_vC022, UV_vC023, UV_vC080, UV_vSex : Variant;
    sSex : String;

    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Clear(Temp : Integer);

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;


var
  frmLD2Q13 : TfrmLD2Q13;
  UV_vJJ12 : String;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}

procedure TfrmLD2Q13.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vDat_bunju      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju      := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa       := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp       := VarArrayCreate([0, 0], varOleStr);
      UV_vName           := VarArrayCreate([0, 0], varOleStr);
      UV_vSex            := VarArrayCreate([0, 0], varOleStr);
      UV_vC078           := VarArrayCreate([0, 0], varOleStr);
      UV_vE069           := VarArrayCreate([0, 0], varOleStr);
      UV_vC022           := VarArrayCreate([0, 0], varOleStr);
      UV_vC023           := VarArrayCreate([0, 0], varOleStr);
      UV_vC080           := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vDat_bunju          , iLength);
      VarArrayRedim(UV_vNum_bunju          , iLength);
      VarArrayRedim(UV_vCod_jisa           , iLength);
      VarArrayRedim(UV_vNum_samp           , iLength);
      VarArrayRedim(UV_vName               , iLength);
      VarArrayRedim(UV_vSex                , iLength);
      VarArrayRedim(UV_vC078               , iLength);
      VarArrayRedim(UV_vE069               , iLength);
      VarArrayRedim(UV_vC022               , iLength);
      VarArrayRedim(UV_vC023               , iLength);
      VarArrayRedim(UV_vC080               , iLength);
   end;
end;

procedure TfrmLD2Q13.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vDat_bunju[DataRow - 1];
      2 : Value := UV_vNum_bunju[DataRow - 1];
      3 : Value := UV_vCod_jisa[DataRow - 1];
      4 : Value := UV_vNum_samp[DataRow - 1];
      5 : Value := UV_vSex[DataRow - 1];
      6 : Value := UV_vC078[DataRow - 1];
      7 : Value := UV_vE069[DataRow - 1];
      8 : Value := UV_vC022[DataRow - 1];
      9 : Value := UV_vC023[DataRow - 1];
      10: Value := UV_vC080[DataRow - 1];
   end;
end;

procedure TfrmLD2Q13.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;

   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      sJisa := '15';
   end
   else
   begin
      sJisa := GV_sJisa;
   end;
end;

procedure TfrmLD2Q13.btnQueryClick(Sender: TObject);
var
  Index : Integer;
  sSELECT, sWHERE, sORDER : String;

begin
   inherited;
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid 환경설정
   Index := 0; sWHERE := '';  sWHERE := ''; sORDER := '';

   with qryGumgin do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sSELECT := 'Select G.*, B.num_bunju, B.dat_bunju                                                                '+
                 'FROM Gumgin G WITH(NOLOCK)                                                                          '+
                 'INNER Join Bunju B with(nolock) On g.dat_gmgn = b.dat_gmgn and g.num_id = B.num_id                  ';
      sWHERE :=  'Where B.dat_bunju = ''' + msk_date.Text + '''                                                       ';
      if copy(cmbJisa.Text,1,2) <> '00' then
      sWHERE := sWHERE + ' AND G.cod_jisa =  '''+ copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And (G.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
      sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''C078'' or cod_hm = ''E069'' or cod_hm = ''C022'' or cod_hm = ''C023'' or cod_hm = ''C080'' ))) ';
      sWhere := sWhere + ' or G.cod_hul in ( select cod_pf from profile where cod_pf in ';
      sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''C078'' or cod_hm = ''E069'' or cod_hm = ''C022'' or cod_hm = ''C023'' or cod_hm = ''C080'' ))) ';
      sWhere := sWhere + ' or G.cod_cancer in (select cod_pf from profile where cod_pf in ';
      sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''C078'' or cod_hm = ''E069'' or cod_hm = ''C022'' or cod_hm = ''C023'' or cod_hm = ''C080'' ))) ';
      sWhere := sWhere + ' or ( cod_chuga like ''%C078%'' or cod_chuga like ''%E069%'' or  cod_chuga like ''%C022%'' or cod_chuga like ''%C023%'' or cod_chuga like ''%C080%'')) ';

      SQL.Clear;
      SQL.Add (sSELECT + sWHERE + sORDER);
      Active := True;

      Gauge1.Progress := 0;
      if qryGumgin.RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'ED7Q46', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         UP_VarMemSet(RecordCount - 1);
         // DataSet의 값을 Variant변수로 이동
         while not qryGumgin.Eof do
         begin
            UP_Clear(Index);

            Gauge1.Progress := Gauge1.Progress + 1;
            Label2.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;
            case StrToInt(Copy(FieldByName('Num_jumin').AsString, 7, 1)) of
               1,3,5,7,9 : sSex := '남';
               2,4,6,8   : sSex := '여';
            end;
               //분주일자
               UV_vDat_bunju[index] := FieldByName('dat_bunju').AsString;
               //분주번호
               UV_vNum_bunju[index] := FieldByName('num_bunju').AsString;
               //센터
               UV_vCod_jisa[index]  := FieldByName('cod_jisa').AsString;
               //샘플번호
               UV_vNum_samp[Index] := FieldByName('Num_samp').AsString;
               //성별
               UV_vSex[Index] := sSex;
               // 혈액결과...
               qryHul.Active := False;
               qryHul.ParamByName('num_id').AsString   := qryGumgin.FieldByName('Num_id').AsString;
               qryHul.ParamByName('Dat_gmgn').AsString := qryGumgin.FieldByName('Dat_gmgn').AsString;

               qryHul.Active := True;
               if qryHul.RecordCount > 0 then
               begin
                  while not qryHul.Eof do
                  begin
                     UV_fGulkwa := '';
                     UV_fGulkwa1 := qryHul.FieldByName('DESC_GLKWA1').AsString;
                     UV_fGulkwa2 := qryHul.FieldByName('DESC_GLKWA2').AsString;
                     UV_fGulkwa3 := qryHul.FieldByName('DESC_GLKWA3').AsString;
                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                     Case qryHul.FieldByName('gubn_part').AsInteger of

                        1 : begin
                               if (Trim(Copy(UV_fGulkwa, 355, 6)) <> '') then
                                  UV_vC078[Index] := Trim(Copy(UV_fGulkwa, 355, 6)); //c078
                               if (Trim(Copy(UV_fGulkwa, 103,6)) <> '') then
                                  UV_vC022[Index] := Trim(Copy(UV_fGulkwa, 103, 6)); //c022
                               if (Trim(Copy(UV_fGulkwa, 109,6)) <> '') then
                                  UV_vC023[Index] := Trim(Copy(UV_fGulkwa, 109, 6)); //c023
                               if (Trim(Copy(UV_fGulkwa, 367,6)) <> '') then
                                  UV_vC080[Index] := Trim(Copy(UV_fGulkwa, 367, 6)); //c080
                            end;
                        5 : begin
                               if (Trim(Copy(UV_fGulkwa, 433,6)) <> '') then
                                  UV_vE069[Index] := Trim(Copy(UV_fGulkwa, 433, 6)); //e069
                            end;
                     end; // end of Case[Gubn_Part];
                     qryHul.Next;
                  end; // end of while[qryHul];
               end; // end of if[qryHul];
               qryHul.Active := False;



            Inc(Index);
            Next;
         end;
      end
      else
      begin
         ShowMessage('조건에 맞는 자료가 존재하지 않습니다.');
      end;
      // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := Index;
   // Grid Focus
   grdMaster.SetFocus;
   // Data건수 표시
   GP_SetDataCnt(grdMaster);
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD2Q13.UP_Clear(Temp : Integer);
begin
    UV_vNum_bunju[Temp]:= '';   UV_vDat_Bunju[Temp] := '';  UV_vCod_jisa[Temp] := '';
    UV_vNum_samp[Temp] := '';   UV_vC078[Temp] := '';       UV_vE069[Temp] := '';
    UV_vC022[Temp] := '';       UV_vC023[Temp] := '';       UV_vC080[Temp] := '';
end;

function TfrmLD2Q13.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (cmbJisa.ItemIndex = -1) or
      (msk_date.Text = '')   then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;


procedure TfrmLD2Q13.btnQuitClick(Sender: TObject);
begin
  inherited;
  Close;
end;
procedure TfrmLD2Q13.SBut_ExcelClick(Sender: TObject);
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

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

procedure TfrmLD2Q13.UP_Click(Sender: TObject);
begin
  inherited;
  if Sender = btnDate       then GF_CalendarClick(msk_date);          
end;



end.




