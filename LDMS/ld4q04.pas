//==============================================================================
// 프로그램 설명 : RIA분주현황
// 시스템        : 통합검진
// 수정일자      : 2005.11.14
// 수정자        : 유동구
// 수정내용      : 새롭게 개발..
// 참고사항      :
//==============================================================================
unit LD4Q04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q04 = class(TfrmSingle)
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
    qryPf_hm: TQuery;
    qryHul: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vRIA, UV_vC044, UV_vC049, UV_vE001, UV_vE002, UV_vE003, UV_vE015,  UV_vT007,UV_vT006,
    UV_vE040, UV_vT011, UV_vT002, UV_vTT01, UV_vTT02 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q04: TfrmLD4Q04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q041;

{$R *.DFM}

procedure TfrmLD4Q04.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo   := VarArrayCreate([0, 0], varOleStr);
      UV_vRIA  := VarArrayCreate([0, 0], varOleStr);
      UV_vC044 := VarArrayCreate([0, 0], varOleStr);
      UV_vC049 := VarArrayCreate([0, 0], varOleStr);
      UV_vE001 := VarArrayCreate([0, 0], varOleStr);
      UV_vE002 := VarArrayCreate([0, 0], varOleStr);
      UV_vE003 := VarArrayCreate([0, 0], varOleStr);

      UV_vE040 := VarArrayCreate([0, 0], varOleStr);
      UV_vT011 := VarArrayCreate([0, 0], varOleStr);
      UV_vT002 := VarArrayCreate([0, 0], varOleStr);
      UV_vE015 := VarArrayCreate([0, 0], varOleStr);
      UV_vT007 := VarArrayCreate([0, 0], varOleStr);
      UV_vT006 := VarArrayCreate([0, 0], varOleStr);
      UV_vTT01 := VarArrayCreate([0, 0], varOleStr);
      UV_vTT02 := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,   iLength);
      VarArrayRedim(UV_vRIA,  iLength);
      VarArrayRedim(UV_vC044, iLength);
      VarArrayRedim(UV_vC049, iLength);
      VarArrayRedim(UV_vE001, iLength);
      VarArrayRedim(UV_vE002, iLength);
      VarArrayRedim(UV_vE003, iLength);

      VarArrayRedim(UV_vE040, iLength);
      VarArrayRedim(UV_vT011, iLength);
      VarArrayRedim(UV_vT002, iLength);
      VarArrayRedim(UV_vE015, iLength);
      VarArrayRedim(UV_vT007, iLength);
      VarArrayRedim(UV_vT006, iLength);
      VarArrayRedim(UV_vTT01, iLength);
      VarArrayRedim(UV_vTT02, iLength);
   end;
end;

function TfrmLD4Q04.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q04.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRIA[DataRow-1];
      3 : Value := UV_vC044[DataRow-1];
      4 : Value := UV_vC049[DataRow-1];
      5 : Value := UV_vE001[DataRow-1];
      6 : Value := UV_vE002[DataRow-1];
      7 : Value := UV_vE003[DataRow-1];
      8 : Value := UV_vE040[DataRow-1];
      9 : Value := UV_vT011[DataRow-1];
     10 : Value := UV_vT002[DataRow-1];
     11 : Value := UV_vE015[DataRow-1];
     12 : Value := UV_vT007[DataRow-1];
     13 : Value := UV_vT006[DataRow-1];
     14 : Value := UV_vTT01[DataRow-1];
     15 : Value := UV_vTT02[DataRow-1];

   end;
end;

procedure TfrmLD4Q04.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT K.num_Bunju, G.num_id, G.dat_gmgn, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga ' +
                 ' FROM Gulkwa K left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';
      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' +
                          ' AND (K.Gubn_Part  = ''04'' OR K.Gubn_Part  = ''07'' ) ';
      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And K.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      sOrderBy := ' ORDER BY K.Num_Bunju';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q04', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            UP_VarMemSet(Index);
            UV_vNo[Index]  := IntToStr(Index+1);
            UV_VRIA[Index] := qryBunju.FieldByName('Num_Bunju').AsString;

            // 항목(장비프로파일)추출..
            sJangbi_List := ''; sHul_List := '';
            if qryBunju.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            // 항목(혈액프로파일)추출..
            if qryBunju.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;
         
            // 항목(종양프로파일)추출..
            if qryBunju.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //추가검사항목..
            iRet := GF_MulToSingle(qryBunju.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            //(특검)추가검사항목..
            iRet := GF_MulToSingle(qryBunju.FieldByName('Tcod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            if pos('C044', sHul_List) > 0 then UV_vC044[Index] := ''
            else                               UV_vC044[Index] := '*';
            if pos('C049', sHul_List) > 0 then UV_vC049[Index] := ''
            else                               UV_vC049[Index] := '*';
            if pos('E001', sHul_List) > 0 then UV_vE001[Index] := ''
            else                               UV_vE001[Index] := '*';
            if pos('E002', sHul_List) > 0 then UV_vE002[Index] := ''
            else                               UV_vE002[Index] := '*';
            if pos('E003', sHul_List) > 0 then UV_vE003[Index] := ''
            else                               UV_vE003[Index] := '*';

            if pos('E040', sHul_List) > 0 then UV_vE040[Index] := ''
            else                               UV_vE040[Index] := '*';
            if pos('T011', sHul_List) > 0 then UV_vT011[Index] := ''
            else                               UV_vT011[Index] := '*';
            if pos('T002', sHul_List) > 0 then UV_vT002[Index] := ''
            else                               UV_vT002[Index] := '*';
            if pos('E015', sHul_List) > 0 then UV_vE015[Index] := ''
            else                               UV_vE015[Index] := '*';
            if pos('T007', sHul_List) > 0 then UV_vT007[Index] := ''
            else                               UV_vT007[Index] := '*';
            if pos('T006', sHul_List) > 0 then UV_vT006[Index] := ''
            else                               UV_vT006[Index] := '*';
            if pos('TT01', sHul_List) > 0 then UV_vTT01[Index] := ''
            else                               UV_vTT01[Index] := '*';
            if pos('TT02', sHul_List) > 0 then UV_vTT02[Index] := ''
            else                               UV_vTT02[Index] := '*';

            // 혈액결과...
            qryHul.Active := False;
            qryHul.ParamByName('num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
            qryHul.ParamByName('Dat_gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
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
                     4 : begin
                            //B2 -MG
                            if pos('C044', sHul_List) > 0 then
                               UV_vC044[Index] := Trim(Copy(UV_fGulkwa,   1, 6));
                            //Ferritin
                            if pos('C049', sHul_List) > 0 then
                               UV_vC049[Index] := Trim(Copy(UV_fGulkwa,   7, 6));
                            //T3
                            if pos('E001', sHul_List) > 0 then
                               UV_vE001[Index] := Trim(Copy(UV_fGulkwa,  13, 6));
                            //T4
                            if pos('E002', sHul_List) > 0 then
                               UV_vE002[Index] := Trim(Copy(UV_fGulkwa,  19, 6));
                            //TSH
                            if pos('E003', sHul_List) > 0 then
                               UV_vE003[Index] := Trim(Copy(UV_fGulkwa,  25, 6));

                            //CYFRA
                            if pos('E040', sHul_List) > 0 then
                               UV_vE040[Index] := Trim(Copy(UV_fGulkwa,  61, 6));
                            //PSA ( 남자 )
                            if pos('T011', sHul_List) > 0 then
                               UV_vT011[Index] := Trim(Copy(UV_fGulkwa, 157, 6));
                            //CEA
                            if pos('T002', sHul_List) > 0 then
                               UV_vT002[Index] := Trim(Copy(UV_fGulkwa, 145, 6));
                         end;
                     5 : begin
                            //CA 125II ( 여자 )
                            if pos('T007', sHul_List) > 0 then
                               UV_vT007[Index] := Trim(Copy(UV_fGulkwa, 115, 6));
                            if pos('T006', sHul_List) > 0 then
                               UV_vT006[Index] := Trim(Copy(UV_fGulkwa, 127, 6));

                         end;
                     7 : begin
                            //AFP
                            if pos('TT01', sHul_List) > 0 then
                               UV_vTT01[Index] := Trim(Copy(UV_fGulkwa, 85, 6));
                            if pos('TT02', sHul_List) > 0 then
                               UV_vTT02[Index] := Trim(Copy(UV_fGulkwa, 169, 6));

                         end;
                      8 : begin
                            //Free T4
                            if pos('E015', sHul_List) > 0 then
                               UV_vE015[Index] := Trim(Copy(UV_fGulkwa, 151, 6));

                         end
                  end; // end of Case[Gubn_Part];
                  qryHul.Next;
               end; // end of while[qryHul];
            end; // end of if[qryHul];
            qryHul.Active := False;


            Inc(Index);
            Next;
         End;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := Index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q04.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q04.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q041 := TfrmLD4Q041.Create(Self);
     frmLD4Q041.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q041 := TfrmLD4Q041.Create(Self);
     frmLD4Q041.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q04.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;


procedure TfrmLD4Q04.SBut_ExcelClick(Sender: TObject);
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


end.
