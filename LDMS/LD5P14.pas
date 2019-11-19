//==============================================================================
// 프로그램 설명 : 건강검진 검체검사 의뢰서 & 결과지
// 시스템        : LDMS
// 수정일자      : 2016.4.15
// 수정자        : 박대용
//==============================================================================
unit LD5P14;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables, Grids, DBGrids, Grids_ts, TSGrid, Gauges,ComObj;

type
  TfrmLD5P14 = class(TfrmPrint)
    cmbJisa: TComboBox;
    mskSt_date: TDateEdit;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    edtST_Name: TEdit;
    qrySaler: TQuery;
    Query1: TQuery;
    edtED_name: TEdit;
    Label7: TLabel;
    mskED_Date: TDateEdit;
    btnDATE1: TSpeedButton;
    Label9: TLabel;
    edtCod_dc: TEdit;
    btnCod_dc: TSpeedButton;
    Label10: TLabel;
    Label8: TLabel;
    cmbnaewon: TComboBox;
    edtDESC_DC: TEdit;
    Label14: TLabel;
    sorting: TComboBox;
    qryDan: TQuery;
    qryTemp: TQuery;
    tsGrid1: TtsGrid;
    Gauge2: TGauge;
    SBut_Excel: TBitBtn;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    btnPreview2: TBitBtn;
    btnPrint2: TBitBtn;
    btnPrintSetup2: TBitBtn;
    btnClose2: TBitBtn;
    Gauge1: TGauge;
    edtED_Hul: TEdit;
    rDHul: TRadioButton;
    edtED_Jumin: TMaskEdit;
    rDJumin: TRadioButton;
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure UP_ReportClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure tsGrid1CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
    var Value: Variant);    
    procedure UP_VarMemSet(iLength : Integer); 
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }  
    UV_iCount    : integer;
    UV_vNum_samp ,UV_vNum_jumin, UV_vDesc_name, UV_vS099, UV_vS091, UV_vS016, UV_vTT03, UV_vY004  : Variant;
  public

    // 지사코드
    UV_sCod_jisa : String;
    // SQL문 저장
    UV_sBasicSQL : String;

    procedure UP_Select;
    procedure UP_Select2;
    { Public declarations }
  end;

var
  frmLD5P14: TfrmLD5P14;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc, LD5P141, LD5P142, ZZExcel;

procedure TfrmLD5P14.UP_Click(Sender: TObject);
var sCode, sName : string;
begin
  inherited;
  if Sender = btnCOD_DC then
  begin
     if GF_DancheClick(Self, sCode, sName) then
     begin
        edtCOD_DC.Text  := sCode;
        edtDesc_dc.Text := sName;
     end;
  end
  else
  if Sender = btnDate then GF_CalendarClick(mskST_Date)
  else
  if Sender = btnDate1 then GF_CalendarClick(mskED_Date);;

end;

procedure TfrmLD5P14.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate)
   else
   if Sender = mskED_Date then UP_Click(btnDate1)
   else if Sender = edtCOD_DC then Up_Click(btnCOD_DC);
end;

procedure TfrmLD5P14.FormCreate(Sender: TObject);
var sJisa: String;
begin
   inherited;

   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   SBut_Excel.Enabled := True;
{
   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbJisa.Enabled := True;
      sJisa := copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbJisa.Enabled := False;
      sJisa := GV_sJisa;
   end;
}
   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

   // 오늘일자 설정
   mskST_Date.Text := GV_sToday;
   mskED_Date.Text := GV_sToday;
   cmbNaewon.itemindex := 0;
   sorting.itemindex := 6;

end;

procedure TfrmLD5P14.UP_ReportClick(Sender: TObject);
begin
   inherited;

   // 조회조건 Check
   if not GF_NotNullCheck(pnlPrint) then exit;
   
   Gauge1.Progress := 0;
   if (Sender = btnPreview) or (Sender = btnPrint) then
   begin
      // Report Form Create
      frmLD5P141 := TfrmLD5P141.Create(Self);

      // SQL문을 저장
      UV_sBasicSQL := frmLD5P141.qryGumgin.SQL.Text;

      with frmLD5P141.QuickRep do
      begin
        // Query 실행
         Up_Select;

        // 인쇄범위 설정
         if rbAll.Checked then
         begin
            PrinterSettings.FirstPage := 1;
            PrinterSettings.LastPage := 1000;
         end
         else
         begin
            PrinterSettings.FirstPage := Trunc(valFrom.Value);
            PrinterSettings.LastPage := Trunc(valTo.Value);
         end;
   
         // 인쇄매수 설정
         PrinterSettings.Copies := spnCopy.Value;

         // Preview or Print
         if Sender = btnPreview then Preview
         else if Sender = btnPrint then Print;
      end;
   end
   else if (Sender = btnPreview2) or (Sender = btnPrint2) then
   begin
      // Report Form Create
      frmLD5P142 := TfrmLD5P142.Create(Self);

      // SQL문을 저장
      UV_sBasicSQL := frmLD5P142.qryGumgin.SQL.Text;

      with frmLD5P142.QuickRep do
      begin
        // Query 실행
         Up_Select2;

        // 인쇄범위 설정
         if rbAll.Checked then
         begin
            PrinterSettings.FirstPage := 1;
            PrinterSettings.LastPage := 1000;
         end
         else
         begin
            PrinterSettings.FirstPage := Trunc(valFrom.Value);
            PrinterSettings.LastPage := Trunc(valTo.Value);
         end;
   
         // 인쇄매수 설정
         PrinterSettings.Copies := spnCopy.Value;

         // Preview or Print
         if Sender = btnPreview2 then Preview
         else if Sender = btnPrint2 then Print;
      end;
   end;
end;

procedure TfrmLD5P14.UP_Select;
var    sWhere, sHangmok_Iist : String;
       i, z, index : Integer;
begin
  UP_VarMemSet(0);
  tsGrid1.Rows := 0;
   with frmLD5P141.qryGumgin do
   begin
      Active := False;
      sWhere := ' WHERE G.Cod_jisa  = ''' + copy(cmbJisa.items[cmbjisa.itemindex],1,2) + ''''
              + ' AND G.Dat_gmgn >= ''' + mskSt_Date.Text + ''''
              + ' AND G.Dat_gmgn <= ''' + mskEd_Date.Text + '''';

      sWhere := sWhere + '  AND (G.gubn_nosin <> '''' or G.gubn_adult <> '''' or G.gubn_life <> '''' or G.gubn_gong <> '''')  ';

      if Trim(edtCod_Dc.Text) <> '' then
         sWhere := sWhere + '  AND G.Cod_dc = ''' + edtCod_Dc.Text + '''';

      if Trim(edtST_name.Text) <> '' then
      begin
         sWhere := sWhere + '  AND G.Desc_name >= ''' + edtST_name.Text + ''''
                          + '  AND G.Desc_name <= ''' + edtED_name.Text + '''';
      end;
      if  cmbnaewon.itemindex = 1 then
          sWhere := sWhere + ' AND G.Gubn_neawon <> ''' + inttostr(2) + '''' //내원
      else if  cmbnaewon.itemindex = 2 then
          sWhere := sWhere + ' AND G.Gubn_neawon = ''' + inttostr(2) + ''''; //출장
      if (rDHul.Checked) and (edtED_Hul.Text <> '') then
         sWhere := sWhere +  ' AND G.num_samp  = ''' + edtED_Hul.Text + '''' ; 
      if (rDJumin.Checked) and (edtED_Jumin.Text <> '') then
         sWhere := sWhere +  ' AND G.num_jumin  = ''' + edtED_Jumin.Text + '''' ;

      case sorting.itemindex of
        0 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_name, G.Num_Jumin';
            end;
        1 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_dept, G.desc_name, G.Num_Jumin';
            end;
        2 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_dept, G.desc_name, G.Num_Jumin';
            end;
        3 : begin
              sWhere := sWhere + '  ORDER BY G.dat_gmgn, G.desc_name, G.Num_Jumin';
            end;
        4 : begin
              sWhere := sWhere + '  ORDER BY G.desc_name, G.Num_Jumin';
            end;
        5 : begin
              sWhere := sWhere + '  ORDER BY G.num_sabun';
            end;
        6 : begin
              sWhere := sWhere + '  ORDER BY G.DAT_GMGN, G.NUM_SAMP';
            end;  
      end;

      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);
      Active := True;
      Gauge1.MinValue := 0;
      Gauge1.MaxValue := frmLD5P141.qryGumgin.RecordCount;

      UP_VarMemSet(frmLD5P141.qryGumgin.RecordCount );
      z := 0;
      index := 0;
      if RecordCount > 0 then
      begin
           with tsGrid1 do
           begin
              if (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) > '2012') then
              begin
                   Col[5].Heading := 'HBs Ag(S099)';
                   Col[6].Heading := 'HBs Ab(S091)';
                   if (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) > '2015') then
                      Col[8].Heading := 'AFP(TT03)'
                   else
                      Col[8].Heading := 'AFP(TT02)';

              end
              else
              begin
                   Col[5].Heading := 'HBs Ag(S007)';
                   Col[6].Heading := 'HBs Ab(S008)';
              end;
           end;

           for i := 0 to frmLD5P141.qryGumgin.RecordCount - 1 do
           begin
                sHangmok_Iist             := frmLD5P141.qryGumgin.FieldByName('cod_chuga').AsString;

                if  (frmLD5P141.qryGumgin.FieldByName('gubn_life').AsString = '1') and
                    (frmLD5P141.qryGumgin.FieldByName('gubn_lfyh').AsString = '1') then
                begin
                    if (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
                       (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') then
                        sHangmok_Iist := sHangmok_Iist + 'S099S091'
                    else
                        sHangmok_Iist := sHangmok_Iist + 'S007S008'
                end;

                with frmLD5P141.qryProfile do
                begin
                   Active := False;
                   ParamByName('cod_pf').AsString    := frmLD5P141.qryGumgin.FieldByName('cod_hul').AsString;
                   Active := True;
                   if  0 < frmLD5P141.qryProfile.RecordCount then
                   begin
                       while not frmLD5P141.qryprofile.eof do
                       begin
                          sHangmok_Iist := sHangmok_Iist +  FieldByName('cod_hm').AsString;
                          next;
                       end;
                   end;
                   Active := False;
                end;

                UV_vNum_samp[index]    :=  frmLD5P141.qryGumgin.FieldByName('num_samp').AsString;
                UV_vNum_jumin[index]   :=  frmLD5P141.qryGumgin.FieldByName('Num_jumin').AsString;
                UV_vDesc_name[index]   :=  frmLD5P141.qryGumgin.FieldByName('desc_name').AsString;

                if  (frmLD5P141.qryGumgin.FieldByName('gubn_life').AsString = '1') and
                    (frmLD5P141.qryGumgin.FieldByName('gubn_lfyh').AsString = '1') then
                begin
                     if (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
                        (copy(frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') then
                     begin
                          if (pos('S099',sHangmok_Iist) > 0) then  //B형간염표면항원
                          begin
                               UV_vS099[index]  := 'V';
                          end
                          else
                          begin
                               UV_vS099[index]  := '';
                          end;

                          if (pos('S091',sHangmok_Iist) > 0) then  //B형간염표면항체
                          begin
                               UV_vS091[index]  := 'V';
                          end
                          else
                          begin
                               UV_vS091[index]  := '';
                          end;
                     end
                     else
                     begin
                          if (pos('S007',sHangmok_Iist) > 0) then  //B형간염표면항원
                          begin
                               UV_vS099[index]  := 'V';
                          end
                          else
                          begin
                               UV_vS099[index]  := '';
                          end;

                          if (pos('S008',sHangmok_Iist) > 0) then  //B형간염표면항체
                          begin
                               UV_vS091[index]  := 'V';
                          end
                          else
                          begin
                               UV_vS091[index]  := '';
                          end;
                     end;
                 end;

                if (pos('S016',sHangmok_Iist) > 0) and
                   (pos('간',frmLD5P141.qryGumgin.FieldByName('desc_yhname').AsString) > 0) then  //C형간염표면항체
                begin
                     UV_vS016[index]  := 'V';
                end
                else
                begin
                     UV_vS016[index]  := '';
                end;

                if (pos('간',frmLD5P141.qryGumgin.FieldByName('desc_yhname').AsString) > 0) then  //혈청알파태아단백검사
                begin
                     if ((pos('TT03',sHangmok_Iist) > 0) and (frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString > '20160101')) or
                        ((pos('TT02',sHangmok_Iist) > 0) and (frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString < '20160101')) then
                             UV_vTT03[index]  := 'V'
                     else
                             UV_vTT03[index]  := '';
                end
                else
                begin
                     UV_vTT03[index]  := '';
                end;

                if (pos('Y004',sHangmok_Iist) > 0) and
                   (pos('대장',frmLD5P141.qryGumgin.FieldByName('desc_yhname').AsString) > 0) and
                   (frmLD5P141.qryGumgin.FieldByName('dat_gmgn').AsString < '20160101')   then  //분변잠혈검사
                begin
                     UV_vY004[index]  := 'V';
                end
                else
                begin
                     UV_vY004[index]  := '';
                end;
                Gauge1.Progress := Gauge1.Progress + 1;

                frmLD5P141.qryGumgin.Next;
                if (UV_vS099[index] = '') and (UV_vS091[index] = '') and (UV_vS016[index] = '') and (UV_vTT03[index] = '') and (UV_vY004[index] = '') then
                begin
                    inc(z);
                    continue;
                end;  
                Inc(index);
           end; //end for
           frmLD5P141.qryGumgin.First;
           tsGrid1.Rows := frmLD5P141.qryGumgin.RecordCount - z;
      end; // end gumgin recordcount

      GP_query_log(frmLD5P141.qryGumgin, 'LD5P14', 'Q', 'N');
   end;

end;

procedure TfrmLD5P14.UP_Select2;
var    sWhere, sHangmok_Iist : String;
       i, z, index : Integer;
begin
  UP_VarMemSet(0);
  tsGrid1.Rows := 0;
   with frmLD5P142.qryGumgin do
   begin
      Active := False;
      sWhere := ' WHERE G.Cod_jisa  = ''' + copy(cmbJisa.items[cmbjisa.itemindex],1,2) + ''''
              + ' AND G.Dat_gmgn >= ''' + mskSt_Date.Text + ''''
              + ' AND G.Dat_gmgn <= ''' + mskEd_Date.Text + '''';

      if Trim(edtCod_Dc.Text) <> '' then
         sWhere := sWhere + '  AND G.Cod_dc = ''' + edtCod_Dc.Text + '''';

      if Trim(edtST_name.Text) <> '' then
      begin
         sWhere := sWhere + '  AND G.Desc_name >= ''' + edtST_name.Text + ''''
                          + '  AND G.Desc_name <= ''' + edtED_name.Text + '''';
      end;
      if  cmbnaewon.itemindex = 1 then
          sWhere := sWhere + ' AND G.Gubn_neawon <> ''' + inttostr(2) + '''' //내원
      else if  cmbnaewon.itemindex = 2 then
          sWhere := sWhere + ' AND G.Gubn_neawon = ''' + inttostr(2) + ''''; //출장
      if (rDHul.Checked) and (edtED_Hul.Text <> '') then
         sWhere := sWhere +  ' AND G.num_samp  = ''' + edtED_Hul.Text + '''' ; 
      if (rDJumin.Checked) and (edtED_Jumin.Text <> '') then
         sWhere := sWhere +  ' AND G.num_jumin  = ''' + edtED_Jumin.Text + '''' ;

      case sorting.itemindex of
        0 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_name, G.Num_Jumin';
            end;
        1 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_dept, G.desc_name, G.Num_Jumin';
            end;
        2 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_dept, G.desc_name, G.Num_Jumin';
            end;
        3 : begin
              sWhere := sWhere + '  ORDER BY G.dat_gmgn, G.desc_name, G.Num_Jumin';
            end;
        4 : begin
              sWhere := sWhere + '  ORDER BY G.desc_name, G.Num_Jumin';
            end;
        5 : begin
              sWhere := sWhere + '  ORDER BY G.num_sabun';
            end;
        6 : begin
              sWhere := sWhere + '  ORDER BY G.DAT_GMGN, G.NUM_SAMP';
            end;
      end;

      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);
      Active := True;

      GP_query_log(frmLD5P142.qryGumgin, 'LD5P14', 'Q', 'N');
   end;

end;

procedure TfrmLD5P14.UP_Change(Sender: TObject);
var sName, sSelect, sWhere, sOrderby : String;
begin
  inherited;
   if Sender = edtCod_dc then
   begin
      qryTemp.Active := False;
      qryTemp.ParamByName('cod_dc').AsString := edtCod_Dc.Text;
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then edtDesc_dc.Text := qryTemp.FieldByName('desc_Dc').AsString
      else                            edtDesc_dc.Text := '';
   end;
end;


procedure TfrmLD5P14.tsGrid1CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
  var Value: Variant);
begin
     inherited;
   // 자룔를 화면에 조회    //대용
      case DataCol of
           1 : Value :=  inttostr(DataRow);
           2 : value :=  UV_vNum_samp[DataRow - 1];
           3 : Value :=  UV_vDesc_name[DataRow - 1];
           4 : Value :=  UV_vNum_jumin[DataRow - 1];
           5 : Value :=  UV_vS099[DataRow - 1];
           6 : Value :=  UV_vS091[DataRow - 1];
           7 : Value :=  UV_vS016[DataRow - 1];
           8 : Value :=  UV_vTT03[DataRow - 1];
           9 : Value :=  UV_vY004[DataRow - 1];
      end;
end;


procedure TfrmLD5P14.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_vS099      := VarArrayCreate([0, 0], varOleStr);
      UV_vS091      := VarArrayCreate([0, 0], varOleStr);
      UV_vS016      := VarArrayCreate([0, 0], varOleStr);
      UV_vTT03      := VarArrayCreate([0, 0], varOleStr);
      UV_vY004      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vNum_samp , iLength);
      VarArrayRedim(UV_vNum_jumin, iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vS099     , iLength);
      VarArrayRedim(UV_vS091     , iLength);
      VarArrayRedim(UV_vS016     , iLength);
      VarArrayRedim(UV_vTT03     , iLength);
      VarArrayRedim(UV_vY004     , iLength);
   end;
end;

procedure TfrmLD5P14.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook, Worksheet: Variant;
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

   Gauge2.MaxValue := tsGrid1.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;

   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, tsGrid1.Rows + 1, 0, tsGrid1.Cols], varOleStr);

      I := 0;
      for Row := 0 to tsGrid1.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to tsGrid1.Cols - 1 do
               ArrV3[Row, Col] := tsGrid1.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to tsGrid1.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := tsGrid1.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[tsGrid1.Rows + 1, tsGrid1.Cols]].Value := ArrV3;

      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;

end;



end.
