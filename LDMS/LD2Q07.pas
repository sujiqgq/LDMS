//==============================================================================
// 프로그램 설명 : 분주실 비번 삭제
// 시스템        : LDMS
// 수정일자      : 2010.10.29
// 수정자        : 김희석
//==============================================================================
// 수정일자      : 2014.08.30
// 수 정 자      : 곽윤설
// 수정내용      : 검진일자, 샘플번호 조회 추가
//==============================================================================

unit LD2Q07;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj,  Gauges;
type
  TfrmLD2Q07 = class(TfrmSingle)
    pnlCond: TPanel;
    grdMibun: TtsGrid;
    btnSLeft: TSpeedButton;
    mskDate_bunju: TDateEdit;
    btnDate_Bunju: TSpeedButton;
    grdMaster: TtsGrid;
    Chk_apply: TCheckBox;
    Dat_gmgn: TDateEdit;
    Label1: TLabel;
    Button1: TButton;
    mskFrom_B: TMaskEdit;
    mskTo_B: TMaskEdit;
    Label10: TLabel;
    Gauge2: TGauge;
    Panel2: TPanel;
    Panel3: TPanel;
    btnMibun: TBitBtn;
    qrybunju: TQuery;
    btn_Excel: TSpeedButton;
    qryD_bunju: TQuery;
    qryD_gulkwa: TQuery;
    qryU_Gumgin: TQuery;
    RB_Dat_B: TRadioButton;
    RB_Dat_G: TRadioButton;
    mskDate_gumgin: TDateEdit;
    btnDate_Gumgin: TSpeedButton;
    mskFrom_G: TEdit;
    mskTo_G: TEdit;
    Label2: TLabel;
    CB_01: TComboBox;
    qryD_Allergy: TQuery;
    SBtn_check: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMibunCellLoaded(Sender: TObject; DataCol,
    DataRow: Integer; var Value: Variant);
    procedure UP_Click(Sender: TObject);
    procedure btnSLeftClick(Sender: TObject);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer; ByUser: Boolean);
    procedure Button1Click(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdProgramRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnDeleteClick(Sender: TObject);
    procedure btn_ExcelClick(Sender: TObject); 
    procedure btnMibunClick(Sender: TObject);
    procedure SBtn_checkClick(Sender: TObject);


  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_Check,UV_Num_bunju,UV_Desc_name,UV_Num_jumin,UV_Dat_gmgn,UV_Num_samp,UV_Cod_jisa,UV_gubn_hulgum,UV_num_id: Variant;

    UV_vNum_bunju, UV_vNum_samp,UV_vdesc_name,UV_vnum_jumin,UV_vNum_id,UV_vCod_jisa,UV_vDat_gmgn : Variant;

     UV_BasicSQL,sWhere ,sOrder: String;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_ProgVarMemSet(iLength : Integer);
    public
    { Public declarations }
  end;

var
  frmLD2Q07: TfrmLD2Q07;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD2Q07.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      // Column Title
      Col[1].Heading := '';
      Col[2].Heading := '분주';
      Col[3].Heading := '이름';
      Col[4].Heading := '주민번호';
      Col[5].Heading := '검진일';
      Col[6].Heading := '핼액샘플';
      Col[7].Heading := '지사코드';
      Col[8].Heading := '아이디';



      // Set Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
      Col[4].Alignment := taCenter;
      Col[5].Alignment := taCenter;
      Col[6].Alignment := taCenter;
      Col[1].ControlType := ctCheck;
   end;

   with grdMibun do
   begin
      Rows := 0;

      // Column Title
      Col[1].Heading := '성명';
      Col[2].Heading := '분주번호';
      Col[3].Heading := '샘플번호 ';
      Col[4].Heading := '주민번호';
      Col[5].Heading := '지사코드';
      Col[6].Heading := '검진일';
      Col[7].Heading := '아이디';
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
      Col[4].Alignment := taCenter;


   end;
end;

procedure TfrmLD2Q07.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_Check       := VarArrayCreate([0, 0], varOleStr);
      UV_num_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_Desc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_Num_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_Dat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_Num_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_Gubn_hulgum := VarArrayCreate([0, 0], varOleStr);
      UV_Cod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_Num_id      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_Check,  iLength);
      VarArrayRedim(UV_num_bunju, iLength);
      VarArrayRedim(UV_Desc_name, iLength);
      VarArrayRedim(UV_Num_jumin, iLength);
      VarArrayRedim(UV_Dat_gmgn, iLength);
      VarArrayRedim(UV_Num_samp, iLength);
      VarArrayRedim(UV_Gubn_hulgum, iLength);
      VarArrayRedim(UV_Cod_jisa, iLength);
      VarArrayRedim(UV_Num_id, iLength);
   end;
end;

procedure TfrmLD2Q07.UP_ProgVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_vnum_jumin := VarArrayCreate([0, 0], varOleStr);
      UV_vnum_id := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      VarArrayRedim(UV_vNum_bunju,  iLength);
      VarArrayRedim(UV_vNum_samp, iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vnum_jumin, iLength);
      VarArrayRedim(UV_vNum_id, iLength);
      VarArrayRedim(UV_vCod_jisa, iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
   end;
end;


procedure TfrmLD2Q07.FormCreate(Sender: TObject);
begin
   inherited;

   CB_01.ItemIndex := 0;

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);
   UP_ProgVarMemSet(0);
   UV_BasicSQL := qryBunju.SQL.Text;
   btnMibun.Enabled := False;

end;

procedure TfrmLD2Q07.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회

     case DataCol of
      1 :if UV_Check[DataRow-1] = 'Y' then
             Value := 1
          else
             Value := 0;
      2 : Value := UV_Num_bunju[DataRow-1];
      3 : Value := UV_Desc_name[DataRow-1];
      4 : Value := UV_Num_jumin[DataRow-1];
      5 : Value := UV_Dat_gmgn[DataRow-1];
      6 : Value := UV_Num_samp[DataRow-1];
      7 : Value := UV_Cod_jisa[DataRow-1];
      8 : Value := UV_num_id[DataRow-1];
    end;



end;

procedure TfrmLD2Q07.btnQueryClick(Sender: TObject);
var i, index : Integer;
begin
   inherited;

   // 조회조건 Check
 //  if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows  := 0;
   grdMibun.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   with qrybunju do
   begin
      Active := False;
        sWhere := ' Where B.cod_bjjs=''15'' ' ;
      //분주일자 or 검진일자
      if RB_Dat_B.Checked then
         sWhere := sWhere + ' And B.dat_bunju = ''' + mskDate_bunju.Text + ''' '
      else if RB_Dat_G.Checked then
         sWhere := sWhere + ' And A.dat_gmgn = ''' + mskDate_gumgin.Text + ''' ';

      //분주번호 or 샘플번호
      if Trim(CB_01.Text) = '분주번호' then
      begin
         sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom_B.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo_B.Text), 4, '9') + '''';
         sOrder := ' order by A.cod_jisa,A.Dat_gmgn,B.num_bunju';
      end
      else if Trim(CB_01.Text) = '샘플번호' then
      begin
         sWhere := sWhere + ' AND A.Num_samp >= ''' + GF_LPad(Trim(mskFrom_G.Text), 6, '0') + '''';
         sWhere := sWhere + ' AND A.Num_samp <= ''' + GF_LPad(Trim(mskTo_G.Text), 6, '9') + '''';
         sOrder := ' order by A.cod_jisa,A.Dat_gmgn,A.num_samp';
      end;


      SQL.Clear;
      SQL.Add(UV_BasicSQL + sWhere + sOrder );
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qrybunju, 'LD2Q07', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);

         for i := 0 to RecordCount - 1 do
         begin
            Uv_check[i]     := 'N';
            UV_Num_bunju[i]  := FieldByName('Num_bunju').AsString;
            UV_Desc_name[i]  := FieldByName('desc_name').AsString;
            UV_Num_jumin[i]  := FieldByName('num_jumin').AsString;
            UV_Dat_gmgn[i]   := FieldByName('dat_gmgn').AsString;
            UV_Num_samp[i]   := FieldByName('Num_samp').AsString;
            UV_gubn_hulgum[i]:= FieldByName('gubn_hulgum').AsString;
            UV_Cod_jisa[i]   := FieldByName('cod_jisa').AsString;
            UV_Num_id[i]     := FieldByName('num_id').AsString;

            Next;
            Inc(index);
         end;

         // Table과 Disconnect
         Active := False;
      end;
   end;
   // Grid에 자료를 할당
 grdMaster.Rows := index;
  if index<1 then
  begin
  Showmessage('자료가 없습니다.');
  end;

end;

procedure TfrmLD2Q07.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // 자료를 Re-query
   if UV_bQuery then
      btnQueryClick(Sender)
   else
   begin
      grdMaster.Rows  := 0;
      grdmibun.Rows := 0;
   end;

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD2Q07.UP_Change(Sender: TObject);
begin
   inherited;
     //분주일자 or 검진일자 활성화
     if RB_Dat_B.Checked then
     begin
        mskDate_bunju.Enabled := True;
        btnDate_Bunju.Enabled := True;
        mskDate_gumgin.Enabled := False;
        btnDate_gumgin.Enabled := False;
     end
     else if RB_Dat_G.Checked then
     begin
        mskDate_bunju.Enabled := False;
        btnDate_Bunju.Enabled := False;
        mskDate_gumgin.Enabled := True;
        btnDate_gumgin.Enabled := True;
     end;

     if Trim(CB_01.Text) = '분주번호' then
     begin
        mskFrom_B.Enabled := True;
        mskTo_B.Enabled := True;
        mskFrom_G.Enabled := False;
        mskTo_G.Enabled := False;
     end
     else
     begin
        mskFrom_B.Enabled := False;
        mskTo_B.Enabled := False;
        mskFrom_G.Enabled := True;
        mskTo_G.Enabled := True;
     end;

   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행

   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD2Q07.grdMibunCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
 
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vdesc_name[DataRow-1];
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vNum_samp [DataRow-1];
      4 : Value := UV_vNum_jumin[DataRow-1];
      5 : Value := UV_vCod_jisa[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vNum_id[DataRow-1];
    end;
end;

procedure TfrmLD2Q07.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate_Bunju then GF_CalendarClick(mskDate_bunju)
   else if Sender = btnDate_Gumgin then GF_CalendarClick(mskDate_gumgin);
end;
procedure TfrmLD2Q07.btnSLeftClick(Sender: TObject);
var i,j,Cnt:integer;
     Change : Boolean;
begin
   inherited;

     if grdMibun.Rows >0 then
         Cnt := grdMibun.Rows
      else
         Cnt :=1;

      for i := 1 to grdMaster.Rows do
       begin
       Change := False;
        if grdMaster.Cell[1, i]=1 then
         begin
           for j:=1 to grdMibun.Rows do
               if  grdMaster.Cell[2, i] = UV_vNum_bunju[j-1] then
                     begin
                     Change := True;
                     break;
               end;
           if not Change then
              begin
                UP_ProgVarMemSet(cnt);
                UV_vNum_bunju[grdMibun.Rows]  := grdMaster.Cell[2, i];
                UV_vDesc_name[grdMibun.Rows]  := grdMaster.Cell[3, i];
                UV_vNum_jumin[grdMibun.Rows]  := grdMaster.Cell[4, i];
                UV_vNum_samp[grdMibun.Rows]   := grdMaster.Cell[6, i];
                UV_vCod_jisa[grdMibun.Rows]   := grdMaster.Cell[7, i];
                UV_vDat_gmgn[grdMibun.Rows]   := grdMaster.Cell[5, i];
                UV_vNum_id[grdMibun.Rows]     := grdMaster.Cell[8, i];
                grdMibun.Rows := grdMibun.Rows + 1;
                cnt:=cnt+1;
                grdMibun.CurrentDataRow := grdMibun.Rows;
                grdMibun.TopRow := grdMibun.Rows;

          end;
      end;
      if grdMibun.Rows >0 then
         btnMibun.Enabled := true
      else
         btnMibun.Enabled := false;
   end;
end;
procedure TfrmLD2Q07.Button1Click(Sender: TObject);
var i,j:integer;
begin
 inherited;
    if Button1.caption='적용' then
      begin
      if  Chk_apply.checked =true then
          begin
          for i:=0 to grdMaster.rows - 1 do
            begin
               if ((UV_Cod_jisa[i] <> '15') and (UV_dat_gmgn[i]=dat_gmgn.text))  then
                   UV_check[i]:='Y';
                   grdMaster.repaint;
                   Button1.caption:='해제';
                   //((UV_gubn_hulgum[i]= '3') or (UV_gubn_hulgum[i]= '1') ) and
          end;
        end
      end
    else if ((Button1.caption='해제') and (Chk_apply.checked =true)) then
       for j:=0 to grdMaster.rows - 1 do
           begin
           UV_check[j]:='N';
           Button1.caption:='적용';
           grdMaster.repaint;
    end;
end;




procedure TfrmLD2Q07.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   // Edit한 현재의 자료를 Variant 변수에 할당
    case DataCol of
    1 : if grdMaster.CurrentCell.Value = 1   then
             UV_check[DataRow-1] := 'Y'
          else
             UV_Check[DataRow-1] := 'N';

    end;
end;


procedure TfrmLD2Q07.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
  inherited;

   if NewRow = 0 then exit;

   GP_SetDataCnt(grdMaster);

end;

procedure TfrmLD2Q07.grdProgramRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
  inherited;
    if NewRow = 0 then exit;

   GP_SetDataCnt(grdMibun);

end;


procedure TfrmLD2Q07.btnDeleteClick(Sender: TObject);
var i,j,k : integer;
begin
   inherited;
       UP_ProgVarMemSet(grdMibun.rows);
          for  j:=1 to grdMibun.rows do
            if grdMibun.RowSelected[j] = true then
            begin
              i:=grdMibun.CurrentDataRow-1;
              UV_vNum_bunju[i]  := '';
              UV_vDesc_name[i]  := '';
              UV_vNum_jumin[i]  := '';
              UV_vNum_samp[i]   := '';
              UV_vCod_jisa[i]   := '';
              UV_vDat_gmgn[i]   := '';
              UV_vNum_id[i]     := '';
              grdMibun.RowVisible[i+1] := False;

            end;
            
  end;

procedure TfrmLD2Q07.btn_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col,T_row : Integer;
  sString :String;


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

   Gauge2.MaxValue := grdMibun.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMibun.Rows + 1, 0, grdMibun.Cols], varOleStr);

      I := 0;
      T_row:=0;
      for Row := 0 to grdMibun.Rows  do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMibun.Cols - 1 do
               ArrV3[T_row, Col] := grdMibun.Col[Col + 1].Heading;
            Inc(T_row);
         end
         else
         begin
            if grdMibun.cell[1, Row] <> '' then
            begin
               for Col := 0 to grdMibun.Cols do
               begin
                   Application.ProcessMessages;
                   ArrV3[T_row, Col] := grdMibun.cell[Col + 1, Row];
               end;
               Inc(T_row);
            end;
          end;

         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMibun.Rows + 1, grdMibun.Cols]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;



procedure TfrmLD2Q07.btnMibunClick(Sender: TObject);
Var I,J,qcnt : Integer;
   Ex :Boolean;

begin
     inherited;

  try
  dmcomm.database.starttransaction;
     Ex := false;
     for qcnt:=1 to  grdMibun.Rows do
         if  grdMibun.Cell[1, qcnt]<>'' then
              begin
              Ex := true;
              break
              end
          else     Ex := false;


     if  Ex then
     For I := 1 To grdMibun.Rows do
        Begin
           With qryD_Gulkwa Do
              Begin
                 Close;
                 ParamByName('Dat_gmgn').AsString   := grdMibun.Cell[6, i];
                 ParamByName('Num_Bunju').AsString  := grdMibun.Cell[2, i];
                 ParamByName('cod_jisa').AsString   := grdMibun.Cell[5, i];
                 ParamByName('Num_id').AsString     := grdMibun.Cell[7, i];
                 ExecSql;
                 GP_query_log(qryD_Gulkwa, 'LD2Q07', 'Q', 'Y');
              End;
          With qryU_Gumgin Do
              Begin
                 Close;
                 ParamByName('Ysno_Bunju').AsString  := 'N';
                 ParamByName('Cod_Jisa').AsString    := grdMibun.Cell[5, i];
                 ParamByName('Num_id').AsString      := grdMibun.Cell[7, i];
                 ParamByName('Dat_Gmgn').AsString    := grdMibun.Cell[6, i];
                 ParamByName('Cod_Update').AsString  := GV_sUserId;
                 ParamByName('Dat_Update').AsString  := GV_sToDay;
                 ExecSql;
                 GP_query_log(qryU_Gumgin, 'LD2Q07', 'Q', 'Y');
              End;
           With qryD_Bunju Do
              Begin
                 Close;
                 ParamByName('Cod_Jisa').AsString  := grdMibun.Cell[5, i];
                 ParamByName('Num_id').AsString    := grdMibun.Cell[7, i];
                 ParamByName('Dat_gmgn').AsString   := grdMibun.Cell[6, i];
                 ParamByName('Num_Bunju').AsString := grdMibun.Cell[2, i];
                 ExecSql;
                 GP_query_log(qryD_Bunju, 'LD2Q07', 'Q', 'Y');
              End;
           With qryD_Allergy Do
              Begin
                 Close;
                 ParamByName('Cod_Jisa').AsString  := grdMibun.Cell[5, i];
                 ParamByName('Num_id').AsString    := grdMibun.Cell[7, i];
                 ParamByName('Dat_gmgn').AsString   := grdMibun.Cell[6, i];
                 ParamByName('Num_Bunju').AsString := grdMibun.Cell[2, i];
                 ExecSql;
                 GP_query_log(qryD_Allergy, 'LD2Q07', 'Q', 'Y');
              End;
            next;
        End;
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

    // qry_Bunju.Active := False;
     Showmessage(' ☞ 미번작업이 완료되었습니다.');
     btnMibun.Enabled := False;
     UP_ProgVarMemSet(grdMibun.rows);
       i:=0;
      for  j:=1 to grdMibun.rows do
         begin

         UV_vNum_bunju[i]  := '';
         UV_vDesc_name[i]  := '';
         UV_vNum_jumin[i]  := '';
         UV_vNum_samp[i]   := '';
         UV_vCod_jisa[i]   := '';
         UV_vDat_gmgn[i]   := '';
         UV_vNum_id[i]     := '';
         grdMibun.RowVisible[i+1] := False;
         i:=i+1;
         end;
         if UV_vNum_bunju[1]  = '' then
              btnMibun.Enabled := false


end;

procedure TfrmLD2Q07.SBtn_checkClick(Sender: TObject);
var  I, iTemp : Integer;
begin
  inherited;
  iTemp := grdMaster.Rows;
   if SBtn_check.Caption = '[전체선택]' then
   begin
     SBtn_check.Caption  := '[제외]';
     for I := 0 to grdMaster.Rows - 1 do
        UV_Check[I] := 'Y';
     grdMaster.RowS := 0;
     grdMaster.RowS := iTemp;
   end // end of if[btnall(전체선택)];
   else if SBtn_check.Caption = '[제외]' then
   begin
     SBtn_check.Caption  := '[전체선택]';
     for I := 0 to grdMaster.Rows - 1 do
        UV_Check[I] := '0';
     grdMaster.RowS := 0;
     grdMaster.RowS := iTemp;// end of if[btnall(해제)];

     end;
end;

end.
