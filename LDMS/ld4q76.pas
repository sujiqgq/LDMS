//==============================================================================
// 프로그램 설명 : URIN 검사결과 김소영 선생 요청
// 시스템        : LDMS
// 수정일자      : 2016-07-13
// 수정자        : 박수지
// 수정내용      : 신규...
// 참고사항      : 한의 본원진료팀 1600410
//==============================================================================

unit LD4Q76;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;

type
  TfrmLD4Q76 = class(TfrmSingle)
    pnlCond: TPanel;
    mskST_date: TDateEdit;
    btnGmgnF: TSpeedButton;
    mskEd_date: TDateEdit;
    btnGmgnT: TSpeedButton;
    Label10: TLabel;
    Label11: TLabel;
    cmbJisa: TComboBox;
    Label12: TLabel;
    qryHul: TQuery;
    qryJangbi: TQuery;
    qryGumgin: TQuery;
    Gauge1: TGauge;
    grdMaster: TtsGrid;
    Label2: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryNsSokun: TQuery;
    qrySokun: TQuery;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryNs_Sokun: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
              DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnQuitClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure UP_SetInitGrid;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

 private
    { Private declarations }

    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010,
    UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016 : Variant;

    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Clear(Temp : Integer);

    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;


var
  frmLD4Q76 : TfrmLD4Q76;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_sUrine_Char, UV_sDat_gmgn : String;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}

procedure TfrmLD4Q76.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
      UV_v013 := VarArrayCreate([0, 0], varOleStr);
      UV_v014 := VarArrayCreate([0, 0], varOleStr);
      UV_v015 := VarArrayCreate([0, 0], varOleStr);
      UV_v016 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_v001, iLength);
      VarArrayRedim(UV_v002, iLength);
      VarArrayRedim(UV_v003, iLength);
      VarArrayRedim(UV_v004, iLength);
      VarArrayRedim(UV_v005, iLength);
      VarArrayRedim(UV_v006, iLength);
      VarArrayRedim(UV_v007, iLength);
      VarArrayRedim(UV_v008, iLength);
      VarArrayRedim(UV_v009, iLength);
      VarArrayRedim(UV_v010, iLength);
      VarArrayRedim(UV_v011, iLength);
      VarArrayRedim(UV_v012, iLength);
      VarArrayRedim(UV_v013, iLength);
      VarArrayRedim(UV_v014, iLength);
      VarArrayRedim(UV_v015, iLength);
      VarArrayRedim(UV_v016, iLength);
   end;
end;

procedure TfrmLD4Q76.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_v001[DataRow - 1];
      2 : Value := UV_v016[DataRow - 1];
      3 : Value := UV_v002[DataRow - 1];
      4 : Value := UV_v003[DataRow - 1];
      5 : Value := UV_v004[DataRow - 1];
      6 : Value := UV_v005[DataRow - 1];
      7 : Value := UV_v006[DataRow - 1];
      8 : Value := UV_v007[DataRow - 1];
      9 : Value := UV_v008[DataRow - 1];
     10 : Value := UV_v009[DataRow - 1];
     11 : Value := UV_v010[DataRow - 1];
     12 : Value := UV_v011[DataRow - 1];
     13 : Value := UV_v012[DataRow - 1];
     14 : Value := UV_v013[DataRow - 1];
     15 : Value := UV_v014[DataRow - 1];
     16 : Value := UV_v015[DataRow - 1];
   end;
end;

function  TfrmLD4Q76.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q76.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;

   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
      sJisa := '15'
   else
      sJisa := GV_sJisa;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

   UP_SetInitGrid;
end;

procedure TfrmLD4Q76.btnQueryClick(Sender: TObject);
var
   sMan, sSelect, sWhere, sOrder,
   sJJ12, sCod_hm, cod_name, nSokun, Temp_gulkwa, sPJindan, sDrug,
   sFamily_history, sPan : String;
   Index, iTemp, iRet, i, sokun_count, iAge : Integer;
   vCod_chuga, vCod_totyh : Variant;
begin
   inherited;
   sSelect := '';  sWhere := '';  sOrder := '';
   sMan := '';  iAge := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid 환경설정
   Index := 0;

   with qryGumgin do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sSelect := ' SELECT G.cod_jisa, G.num_id, G.num_sabun, G.dat_gmgn, G.desc_name, G.num_jumin, G.desc_dept, G.cod_hul, G.cod_jangbi, G.cod_cancer,             ' +
                 '        G.gubn_nosin, G.gubn_adult, G.gubn_agab, G.gubn_life, G.cod_chuga, G.gubn_nsyh, G.gubn_adyh, G.gubn_agyh, G.gubn_lfyh, B.* ,A.gubn_part  ' +
                 ' FROM Gumgin G with(nolock) left outer join Bunju B with(nolock) ON G.cod_jisa = B.cod_jisa and G.num_id = B.Num_id AND G.dat_gmgn = B.dat_gmgn  ' +
                 ' left outer join Gulkwa A with(nolock) ON G.cod_jisa = A.cod_jisa and G.num_id = A.Num_id AND G.dat_gmgn = A.dat_gmgn                            ' ;

      if Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) <> '00' then
      begin
         sWhere := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) + '''';
         sWhere := sWhere + '   AND B.dat_bunju >= ''' + mskST_date.Text + '''' ;
         sWhere := sWhere + '   AND B.dat_bunju <= ''' + mskEd_date.Text + '''';
         sWhere := sWhere + 'AND  A.Gubn_Part = ''03'' ' ;
         sWhere := sWhere + 'AND  G.ysno_bunju = ''Y'' ' ;
         sOrder := ' Order BY  B.num_bunju, G.Dat_gmgn' ;
      end
      else
      begin
         sWhere := sWhere + ' WHERE B.dat_bunju >= ''' + mskST_date.Text + '''';
         sWhere := sWhere + '   AND B.dat_bunju <= ''' + mskEd_date.Text + '''';
         sWhere := sWhere + 'AND  A.Gubn_Part = ''03'' ';
         sWhere := sWhere + 'AND  G.ysno_bunju = ''Y'' ';
         sOrder := ' Order BY G.cod_jisa, G.num_jumin, G.Dat_gmgn';
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);

      Active := True;

      Gauge1.Progress := 0;
      if qryGumgin.RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'ED7Q35', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         UP_VarMemSet(qryGumgin.RecordCount);

         // DataSet의 값을 Variant변수로 이동
         while not qryGumgin.Eof do
         begin
            UP_Clear(Index);
            Gauge1.Progress := Gauge1.Progress + 1;
            Label2.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;

            sPJindan := ''; sDrug := ''; sFamily_history := '';

            //검사항목축출
            sCod_hm := '';
            if FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;
            iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_chuga[i];

           // 노신, 성인병, 간염에 대한 검사항목 추출
            cod_name := '';
            // 노신유형Display
            if FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            //성인병유형Display
            if FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            //간염유형Display
            if FieldByName('gubn_agab').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            //생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_totyh[i];

            GP_JuminToAge(FieldByname('num_jumin').AsString,FieldByname('Dat_gmgn').AsString, sMan, iAge);

            UV_v001[Index] := index+1;
            UV_v002[Index] := FieldByName('dat_bunju').AsString;
            UV_v003[Index] := FieldByName('num_bunju').AsString;
            UV_v004[Index] := sMan;
            UV_v016[Index] := FieldByName('cod_jisa').AsString;
            // 혈액결과...
            qryHul.Active := False;
            qryHul.ParamByName('num_id').AsString   := qryGumgin.FieldByName('Num_id').AsString;
            qryHul.ParamByName('Dat_gmgn').AsString := qryGumgin.FieldByName('Dat_gmgn').AsString;
            qryHul.Active := True;
            if qryHul.RecordCount > 0 then
            begin
               while not qryHul.Eof do
               begin
                  UV_fGulkwa := ''; UV_sDat_gmgn := ''; UV_sUrine_Char := '';
                  UV_sDat_gmgn := qryGumgin.FieldByName('Dat_gmgn').AsString;
                  UV_fGulkwa1 := qryHul.FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := qryHul.FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := qryHul.FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                  Case qryHul.FieldByName('gubn_part').AsInteger of
                     1 : begin
                            if pos('C042', sCod_hm) > 0 then UV_v015[Index] := Trim(Copy(UV_fGulkwa, 187, 6));
                         end;
                     3 : begin
                            if pos('U001', sCod_hm) > 0 then UV_v005[Index] := Trim(Copy(UV_fGulkwa,   7, 6));
                            if pos('U002', sCod_hm) > 0 then UV_v006[Index] := Trim(Copy(UV_fGulkwa,  13, 6));
                            if pos('U003', sCod_hm) > 0 then UV_v007[Index] := Trim(Copy(UV_fGulkwa,  19, 6));
                            if pos('U004', sCod_hm) > 0 then UV_v008[Index] := Trim(Copy(UV_fGulkwa,  25, 6));
                            if pos('U005', sCod_hm) > 0 then UV_v009[Index] := Trim(Copy(UV_fGulkwa,  31, 6));
                            if pos('U006', sCod_hm) > 0 then UV_v010[Index] := Trim(Copy(UV_fGulkwa,  37, 6));
                            if pos('U007', sCod_hm) > 0 then UV_v011[Index] := Trim(Copy(UV_fGulkwa,  43, 6));
                            if pos('U008', sCod_hm) > 0 then UV_v012[Index] := Trim(Copy(UV_fGulkwa,  49, 6));
                            if pos('U009', sCod_hm) > 0 then UV_v013[Index] := Trim(Copy(UV_fGulkwa,  55, 6));
                            if pos('U010', sCod_hm) > 0 then UV_v014[Index] := Trim(Copy(UV_fGulkwa,  61, 6));
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


procedure TfrmLD4Q76.UP_Clear(Temp : Integer);
begin
    UV_v001[Temp] := ''; UV_v002[Temp] := ''; UV_v003[Temp] := ''; UV_v004[Temp] := ''; UV_v005[Temp] := '';
    UV_v006[Temp] := ''; UV_v007[Temp] := ''; UV_v008[Temp] := ''; UV_v009[Temp] := ''; UV_v010[Temp] := '';
    UV_v011[Temp] := ''; UV_v012[Temp] := ''; UV_v013[Temp] := ''; UV_v014[Temp] := ''; UV_v015[Temp] := '';
end;

procedure TfrmLD4Q76.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnGmgnF then GF_CalendarClick(mskST_date)
   else if Sender = btnGmgnT then GF_CalendarClick(mskEd_date);
end;

function TfrmLD4Q76.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (cmbJisa.ItemIndex = -1) or
      ((mskST_date.Text = '') and (mskEd_date.Text = '' )) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;


procedure TfrmLD4Q76.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;
   if Sender = mskST_date then UP_Click(btnGmgnF)
   else if Sender = mskEd_date then UP_Click(btnGmgnT);

end;

procedure TfrmLD4Q76.btnQuitClick(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmLD4Q76.SBut_ExcelClick(Sender: TObject);
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

procedure TfrmLD4Q76.UP_SetInitGrid;
var
   i : Integer;
begin
  with grdMaster do
  begin
    RowS := 0;
    cols := 16;
    Col[1].Heading   := '순번';
    Col[2].Heading   := '지사';
    Col[3].Heading   := '분주일자';
    Col[4].Heading   := '분주번호';
    Col[5].Heading   := '성별';
    Col[6].Heading   := 'U001';
    Col[7].Heading   := 'U002';
    Col[8].Heading   := 'U003';
    Col[9].Heading   := 'U004';
    Col[10].Heading   := 'U005';
    Col[11].Heading  := 'U006';
    Col[12].Heading  := 'U007';
    Col[13].Heading  := 'U008';
    Col[14].Heading  := 'U009';
    Col[15].Heading  := 'U010';
    Col[16].Heading  := 'C042';
  end;
end;

end.
