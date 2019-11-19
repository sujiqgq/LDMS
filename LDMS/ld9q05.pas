unit LD9Q05;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges, Math;

type
  TfrmLD9Q05 = class(TfrmSingle)
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
    qryGumgin: TQuery;
    Gauge1: TGauge;
    grdMaster: TtsGrid;
    Label2: TLabel;
    qryCell: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm2: TQuery;
    CheckBox1: TCheckBox;
    chk_Gong: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
              DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnQuitClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

 private
    { Private declarations }

    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010,
    UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016, UV_v017, UV_v018, UV_v019, UV_v020 : Variant;


    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Clear(Temp : Integer);

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;


var
  frmLD9Q05 : TfrmLD9Q05;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_sUrine_Char,UV_sDat_gmgn : String;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}

procedure TfrmLD9Q05.UP_VarMemSet(iLength : Integer);
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
      UV_v017 := VarArrayCreate([0, 0], varOleStr);
      UV_v018 := VarArrayCreate([0, 0], varOleStr);
      UV_v019 := VarArrayCreate([0, 0], varOleStr);
      UV_v020 := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_v017, iLength);
      VarArrayRedim(UV_v018, iLength);
      VarArrayRedim(UV_v019, iLength);
      VarArrayRedim(UV_v020, iLength);
   end;
end;

procedure TfrmLD9Q05.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
       1: Value := UV_v001[DataRow - 1];
       2: Value := UV_v002[DataRow - 1];
       3: Value := UV_v003[DataRow - 1];
       4: Value := UV_v004[DataRow - 1];
       5: Value := UV_v005[DataRow - 1];
       6: Value := UV_v006[DataRow - 1];
       7: Value := UV_v007[DataRow - 1];
       8: Value := UV_v008[DataRow - 1];
       9: Value := UV_v009[DataRow - 1];
      10: Value := UV_v010[DataRow - 1];
      11: Value := UV_v011[DataRow - 1];
      12: Value := UV_v012[DataRow - 1];
      13: Value := UV_v013[DataRow - 1];
      14: Value := UV_v014[DataRow - 1];
      15: Value := UV_v015[DataRow - 1];
      16: Value := UV_v016[DataRow - 1];
      17: Value := UV_v017[DataRow - 1];
      18: Value := UV_v018[DataRow - 1];
      19: Value := UV_v019[DataRow - 1];
      20: Value := UV_v020[DataRow - 1];

   end;
end;

function  TfrmLD9Q05.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD9Q05.FormCreate(Sender: TObject);
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

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

end;

procedure TfrmLD9Q05.btnQueryClick(Sender: TObject);
var
   sSex, sSelect, sWhere, sOrder,
   sJJ12, sCod_hm, cod_name, nSokun, tSokun, YsNoFirst, a, b,sMan: String;
   Index, iAge, iTemp, iRet, i, num : Integer;
   vCod_chuga, vCod_totyh, vCod_profile : Variant;

   check_tk : boolean;
begin
   inherited;
   sSelect := '';  sWhere := '';  sOrder := '';
   sSex := '';
   iAge := 0;

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

      sSelect := 'SELECT Distinct G.*, T.cod_prf, T.cod_chuga as TkGum_chuga  FROM Gumgin G WITH(nolock)      ' +
                 'left outer join Tkgum  T  ON G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn ' ;
      IF chk_Gong.Checked then
      begin
      sSelect := sSelect + 'Join Send_gongdan S on S.cod_jisa = G. Cod_Jisa and S.dat_gmgn = G.dat_gmgn and S.num_id = G.num_id';
      end;

      if Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) <> '00' then
      begin
         sWhere := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) + '''';
      end;

      if CheckBox1.Checked then sWhere := sWhere + '   AND G.num_samp <> '''' ';

      if chk_Gong.Checked  then sWhere := sWhere + '   AND (S.dat_chgu <> '''') AND (S.gubn_gongdan = ''C'' OR S.gubn_gongdan = ''G'')';

      sWhere := sWhere + '   AND G.dat_gmgn >= ''' + mskST_date.Text + '''';
      sWhere := sWhere + '   AND G.dat_gmgn <= ''' + mskEd_date.Text + '''';

      sWhere := sWhere + '   AND (((G.gubn_nosin =''1'' or G.gubn_nosin =''2'') and  G.gubn_nscg = ''2'') or ' +
                         '        ((G.gubn_life =''1''  or G.gubn_life =''2'' ) and  G.gubn_lfcg = ''2'') or ' +
                         '        ((G.gubn_adult =''1'' or G.gubn_adult =''2'') and  G.gubn_adcg = ''2'') or ' +
                         '         (gubn_gong =''1'' and G.gubn_gocg =''2''))                                ' ;

      sOrder := ' Order BY G.desc_name, G.num_jumin ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);

      Active := True;

      Gauge1.Progress := 0;
      if qryGumgin.RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD9Q05', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         UP_VarMemSet(qryGumgin.RecordCount-1);

         UP_Clear(0);

         // DataSet의 값을 Variant변수로 이동
         while not qryGumgin.Eof do
         begin
            UP_Clear(Index);
            Gauge1.Progress := Gauge1.Progress + 1;
            Label2.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;

            
            UV_v001[Index]  := Index + 1;                                         //1
            UV_v002[Index]  := FieldByName('Dat_gmgn').AsString;                  //2
            UV_v003[Index]  := FieldByName('Num_samp').AsString;                  //3
            UV_v004[Index]  := copy(FieldByName('num_jumin').AsString,1,6);       //4
            UV_v005[Index]  := copy(FieldByName('desc_name').AsString,1,4) + '*'; //5

            //검사항목추출
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

           //-------------------------------------------------------------------

            //---- 특검항목 추출...
            iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_profile);
            for i := 0 to iRet-1 do
            begin
               if Trim(vCod_profile[i]) <> '' then
               begin
                  qryPf_hm2.Active := False;
                  qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[i];
                  qryPf_hm2.Active := True;
                  if qryPf_hm2.RecordCount > 0 then
                  begin
                     while not qryPf_hm2.Eof do
                     begin
                        check_tk := True;
                        for num := 1 to round(Length(sCod_hm)/4) do
                        begin
                           if copy(sCod_hm, (num * 4) - 3,4) =
                              qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                        end;
                        if check_tk = True then
                           sCod_hm := sCod_hm + qryPf_hm2.FieldByName('COD_HM').AsString;
                        qryPf_hm2.Next;
                     end; // end of while(항목 처리)
                  end; // end of if
               end; //end of if
            end; //end of for
            sCod_hm := sCod_hm + FieldByName('TkGum_chuga').AsString;

//------------------------------------------------------------------------------

           // 노신, 성인병, 생애에 대한 검사항목 추출
            cod_name := '';
            // 노신유형Display
            if FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            //성인병유형Display
            if FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            //생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_totyh[i];

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
                            if pos('C032', sCod_hm) > 0 then                     //2.식전혈당(공복시혈당)
                               UV_v007[Index] := Trim(Copy(UV_fGulkwa, 157, 6));
                            if pos('C025', sCod_hm) > 0 then                     //3.총콜레스테롤
                               UV_v008[Index] := Trim(Copy(UV_fGulkwa, 121, 6));
                            if pos('C026', sCod_hm) > 0 then                     //4.HDL-콜레스테롤
                               UV_v009[Index] := Trim(Copy(UV_fGulkwa, 127, 6));
                            if pos('C027', sCod_hm) > 0 then                     //5.LDL-콜레스테롤
                               UV_v010[Index] := Trim(Copy(UV_fGulkwa, 133, 6));
                            if pos('C028', sCod_hm) > 0 then                     //6.중성지방
                               UV_v011[Index] := Trim(Copy(UV_fGulkwa, 139, 6));
                            if pos('C042', sCod_hm) > 0 then                     //7.크레아티닌
                               UV_v012[Index] := Trim(Copy(UV_fGulkwa, 187, 6));
                            if pos('C009', sCod_hm) > 0 then                     //8.혈청지오티(C009)
                               UV_v013[Index] := Trim(Copy(UV_fGulkwa, 49,  6));
                            if pos('C010', sCod_hm) > 0 then                     //9.혈청지피티(C010)
                               UV_v014[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                            if pos('C011', sCod_hm) > 0 then                     //10.감마지티피(C011)
                               UV_v015[Index] := Trim(Copy(UV_fGulkwa, 61,  6));
                         end;
                     2 : begin
                             if pos('H002', sCod_hm) > 0 then                    //1.혈색소(H002)
                               UV_v006[Index] := Trim(Copy(UV_fGulkwa,  7,  6));
                         end;
                     3 : begin
                            if pos('U004', sCod_hm) > 0 then                     //11.요단백(U004)
                            UV_v016[Index] := Trim(Copy(UV_fGulkwa,  25, 6));

                            if pos('Y004', sCod_hm) > 0 then                     //13.분별잠혈
                               UV_v018[Index] := Trim(Copy(UV_fGulkwa, 85, 6));
                         end;
                     5 : begin
                            if pos('TT03', sCod_hm) > 0 then                    //12.AFP(E)-수치
                               UV_v017[Index] := Trim(Copy(UV_fGulkwa, 451, 6));
                         end;
                     7 : begin
                            if pos('S099', sCod_hm) > 0 then                    //14.HBs Ag
                               UV_v019[Index] := Trim(Copy(UV_fGulkwa, 229, 6));
                            if pos('S091', sCod_hm) > 0 then                    //15.HBs Ab
                               UV_v020[Index] := Trim(Copy(UV_fGulkwa, 186, 6));
                         end;
                  end; // end of Case[Gubn_Part];
                  qryHul.Next;
               end; // end of while[qryHul];
            end; // end of if[qryHul];
            qryHul.Active := False;

            if (UV_v006[index] = '') and (UV_v007[index] = '') and (UV_v008[index] = '') and (UV_v009[index] = '') and (UV_v010[index] = '') and (UV_v011[index] = '') and
               (UV_v012[index] = '') and (UV_v013[index] = '') and (UV_v014[index] = '') and (UV_v015[index] = '') and (UV_v016[index] = '') and (UV_v017[index] = '') and
               (UV_v018[index] = '') and (UV_v019[index] = '') and (UV_v020[index] = '') then Next
            else
            begin
                 Inc(Index);
                 Next;
            end;
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
   grdMaster.Rows := Index ;
   // Grid Focus
   grdMaster.SetFocus;
   // Data건수 표시
   GP_SetDataCnt(grdMaster);
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;


procedure TfrmLD9Q05.UP_Clear(Temp : Integer);
begin
    UV_v001[Temp] := ''; UV_v002[Temp] := ''; UV_v003[Temp] := ''; UV_v004[Temp] := ''; UV_v005[Temp] := '';
    UV_v006[Temp] := ''; UV_v007[Temp] := ''; UV_v008[Temp] := ''; UV_v009[Temp] := ''; UV_v010[Temp] := '';
    UV_v011[Temp] := ''; UV_v012[Temp] := ''; UV_v013[Temp] := ''; UV_v014[Temp] := ''; UV_v015[Temp] := '';
    UV_v016[Temp] := ''; UV_v017[Temp] := ''; UV_v018[Temp] := ''; UV_v019[Temp] := ''; UV_v020[Temp] := '';
end;

procedure TfrmLD9Q05.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnGmgnF then GF_CalendarClick(mskST_date)
   else if Sender = btnGmgnT then GF_CalendarClick(mskEd_date);
end;

function TfrmLD9Q05.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (cmbJisa.ItemIndex = -1) or
      ((mskST_date.Text = '') and (mskEd_date.Text = '' ))  then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;


procedure TfrmLD9Q05.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_date then UP_Click(btnGmgnF)
   else if Sender = mskEd_date then UP_Click(btnGmgnT);

end;

procedure TfrmLD9Q05.btnQuitClick(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmLD9Q05.SBut_ExcelClick(Sender: TObject);
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




