//==============================================================================
// 프로그램 설명 : 특검 결과 분석값 입력
// 시스템        : 통합검진
// 수정일자      : 2016.07.26
// 수정자        : 박수지
// 수정내용      :
// 참고사항      : 한의 재단특수분석파트 1600138 (신진주 책임)
//==============================================================================
unit LD3I12;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD3I12 = class(TfrmSingle)
    OpenDialog: TOpenDialog;
    GroupBox1: TGroupBox;
    gkPie: TGauge;
    pnlJisa: TPanel;
    Btn_Insert: TBitBtn;
    pnlMaster: TPanel;
    btnDESC_DEPT: TSpeedButton;
    btnStart: TBitBtn;
    Panel5: TPanel;
    edtFile: TEdit;
    Panel3: TPanel;
    Panel6: TPanel;
    qrdSPList: TtsGrid;
    qryGumgin: TQuery;
    qryGulkwa: TQuery;
    qryU_Gulkwa: TQuery;
    TG_gumgin: TStringGrid;
    qrdSList: TtsGrid;
    Panel2: TPanel;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    qryProfile: TQuery;
    qryNo_hangmok: TQuery;
    ComboBox1: TComboBox;
  procedure FormCreate(Sender: TObject);
    procedure btnDESC_DEPTClick(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure Btn_InsertClick(Sender: TObject);
    procedure qrdSListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure qrdSPListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnStartClick(Sender: TObject);

 private
    { Private declarations }
    UV_sJisa : String;
    Excel, Worksheet  : Variant;

    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005, UV_vS006, UV_vS007 : Variant;
    UV_vSP001, UV_vSP002, UV_vSP003, UV_vSP004, UV_vSP005 : Variant;

    procedure UP_SVarMemSet(iLength : Integer);
    procedure UP_SPVarMemSet(iLength : Integer);


  public
    { Public declarations }
    UV_sValue  : array[1..8] of String;

  end;


var
  frmLD3I12 : TfrmLD3I12;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I12.FormCreate(Sender: TObject);
begin
   inherited;
   if  copy(GV_sUserId,1,2) = '51' then ComboBox1.Visible := true
   else ComboBox1.Visible := False;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I12.UP_SVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vS001 := VarArrayCreate([0, 0], varOleStr);
      UV_vS002 := VarArrayCreate([0, 0], varOleStr);
      UV_vS003 := VarArrayCreate([0, 0], varOleStr);
      UV_vS004 := VarArrayCreate([0, 0], varOleStr);
      UV_vS005 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vS001, iLength);
      VarArrayRedim(UV_vS002, iLength);
      VarArrayRedim(UV_vS003, iLength);
      VarArrayRedim(UV_vS004, iLength);
      VarArrayRedim(UV_vS005, iLength);
   end;
end;
procedure TfrmLD3I12.qrdSListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vS001[DataRow - 1];
      2 : Value := UV_vS002[DataRow - 1];
      3 : Value := UV_vS003[DataRow - 1];
      4 : Value := UV_vS004[DataRow - 1];
      5 : Value := UV_vS005[DataRow - 1];
   end;
end;
procedure TfrmLD3I12.UP_SPVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vSP001 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP002 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP003 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP004 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP005 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vSP001, iLength);
      VarArrayRedim(UV_vSP002, iLength);
      VarArrayRedim(UV_vSP003, iLength);
      VarArrayRedim(UV_vSP004, iLength);
      VarArrayRedim(UV_vSP005, iLength);
   end;
end;
procedure TfrmLD3I12.qrdSPListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vSP001[DataRow - 1];
      2 : Value := UV_vSP002[DataRow - 1];
      3 : Value := UV_vSP003[DataRow - 1];
      4 : Value := UV_vSP004[DataRow - 1];
      5 : Value := UV_vSP005[DataRow - 1];
   end;
end;

procedure TfrmLD3I12.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I12.btnStartClick(Sender: TObject);
var
  icol, sCol, iRow, sRow : Integer;
  sLine, FileName, sSheetName : String;
begin
   Screen.Cursor := crHourGlass;
   sCol := 0;  sRow := 0;
   try
      Excel:= CreateOleObject('Excel.Application');
//      Excel.Visible := True;
      Excel.Workbooks.Open(OpenDialog.FileName);
      if Trim(sSheetName) = '' Then
        Worksheet := Excel.Workbooks[1].WorkSheetS[Excel.WorkBooks[1].Worksheets[1].Name]
      else
        Worksheet := Excel.Workbooks[1].WorkSheetS[SSheetName];
   except
      MessageDlg('Excel이 정상적으로 수행되지 않았습니다.' + Chr(13) +
                         '환경설정을 다시 하시기 바랍니다.',mtError,[mbOk], 0);
   end;
   For iRow := 1 to 65000 do
   begin
      if WorkSheet.Cells[iRow,1].Text <> '' then
         Inc(sRow)
      else
         Break;
   end;
   For iCol := 1 to 20 do
   begin
      if WorkSheet.Cells[1,iCol].Text <> '' then
         Inc(sCol)
      else
         Break;
   end;

   TG_gumgin.RowCount := sRow;
   TG_gumgin.ColCount := sCol;
   
   For iRow := 0 to sRow do
   begin
      For iCol := 0 to sCol do
         TG_gumgin.Cells[iCol,iRow] := Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text);
   end;

   if not VarIsEmpty(Excel) then
     Excel.Quit;
   Screen.Cursor := crDefault;
end;


procedure TfrmLD3I12.Btn_InsertClick(Sender: TObject);
var
    sGlkwa,S, sHul_List : String;
    I, iTemp, sIndex, spIndex: Integer;
begin
   inherited;
   sIndex := 0;
   spIndex := 0;
   if TG_gumgin.RowCount = 0 then
   begin
      showmessage('데이터가 존재하지 않습니다.');
      exit;
   end;
   if not GF_MsgBox('Confirmation', 'Transfer를 시작 하시겠습니까.?') then exit;
   gkPie.MaxValue := TG_gumgin.RowCount -1;
   gkPie.Progress := 0;

   for I := 1 to TG_gumgin.RowCount -1 do
   begin
      gkPie.Progress := gkPie.Progress + 1;
      // DB 작업
      try

         // 초기결과값.
         with qryGulkwa do
         begin
              Active := False;
              if copy(GV_sUserId,1,2) <> '51' then ParamByName('cod_bjjs').AsString   := '15'
              else ParamByName('cod_bjjs').AsString   := copy(ComboBox1.Text,1,2);

              qryGulkwa.ParamByName('num_bunju').AsString  := GF_LPad(TG_gumgin.Cells[0,I], 4, '0');
              qryGulkwa.ParamByName('dat_bunju').AsString  := TG_gumgin.Cells[1,I];
              //part구분
              if      (TG_gumgin.Cells[3,I] = 'SE42') or (TG_gumgin.Cells[3,I] = 'SE85')
                   or (TG_gumgin.Cells[3,I] = 'SE87') or (TG_gumgin.Cells[3,I] = 'SE89')
                   or (TG_gumgin.Cells[3,I] = 'SE90') or (TG_gumgin.Cells[3,I] = 'SE92')
                   or (TG_gumgin.Cells[3,I] = 'SE93') or (TG_gumgin.Cells[3,I] = 'SE96')
                   or (TG_gumgin.Cells[3,I] = 'SE97') or (TG_gumgin.Cells[3,I] = 'SE98')
                   or (TG_gumgin.Cells[3,I] = 'SEA1') or (TG_gumgin.Cells[3,I] = 'SEA2')
                   or (TG_gumgin.Cells[3,I] = 'SEA6') then
                    ParamByName('gubn_part').AsString  := '09'
              Else
                    ParamByName('gubn_part').AsString  := '';

              Active := True;
              if RecordCount > 0 then
              begin
                sHul_List := UF_hangmokList;
                if (pos('SE42',sHul_List) > 0) or (pos('SE85',sHul_List) > 0) or (pos('SE87',sHul_List) > 0) or (pos('SE89',sHul_List) > 0) or
                   (pos('SE90',sHul_List) > 0) or (pos('SE92',sHul_List) > 0) or (pos('SE93',sHul_List) > 0) or (pos('SE96',sHul_List) > 0) or
                   (pos('SE97',sHul_List) > 0) or (pos('SE98',sHul_List) > 0) or (pos('SEA1',sHul_List) > 0) or (pos('SEA2',sHul_List) > 0) or (pos('SEA6',sHul_List) > 0) then
                BEGIN
                   UP_SVarMemSet(sIndex);

                   UV_vS001[sIndex] := TG_gumgin.Cells[0,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[1,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[2,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[3,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[4,I];

                   UV_fGulkwa := '';
                   UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                   sGlkwa := UV_fGulkwa;
                   //Active := False;
                   if sGlkwa = '' then
                   begin
                        for iTemp := 1 to 10 do
                            begin
                                 sGlkwa := sGlkwa + '                                                            ';
                            end; // end of for;
                   end; // end of if

                   if copy(GV_sUserId,1,2) <> '51' then
                   begin
                      if FieldByName('Cod_jisa').AsString = '51' then
                      begin
                           showmessage(qryGulkwa.ParamByName('num_bunju').AsString +'/'+ qryGulkwa.ParamByName('dat_bunju').AsString+ ' '+
                           TG_gumgin.Cells[2,I] +' => 광주지사(51) 수검자가 포함되어 있습니다.' );
                           exit;
                      end;
                   end;

                   // 1.SEA6
                   if (pos('SEA6',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SEA6') then
                        sGlkwa := Copy(sGlkwa, 1, 414) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 421, Length(sGlkwa))
                   // 2.SE87
                   Else if (pos('SE87',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE87') then
                        sGlkwa := Copy(sGlkwa, 1, 330) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 337, Length(sGlkwa))
                   // 3.SE90
                   Else if (pos('SE90',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE90') then
                        sGlkwa := Copy(sGlkwa, 1, 348) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 355, Length(sGlkwa))
                   // 4.SE93
                   Else if (pos('SE93',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE93') then
                        sGlkwa := Copy(sGlkwa, 1, 366) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 373, Length(sGlkwa))
                   // 5.SE97
                   Else if (pos('SE97',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE97') then
                        sGlkwa := Copy(sGlkwa, 1, 228) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 235, Length(sGlkwa))
                   // 6.SEA1
                   Else if (pos('SEA1',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SEA1') then
                        sGlkwa := Copy(sGlkwa, 1, 396) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 403, Length(sGlkwa))
                   // 7.SE42
                   Else if (pos('SE42',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE42') then
                        sGlkwa := Copy(sGlkwa, 1, 6) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 13, Length(sGlkwa))
                   // 8.SE85
                   Else if (pos('SE85',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE85') then
                        sGlkwa := Copy(sGlkwa, 1, 318) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 325, Length(sGlkwa))
                   // 9.SE89
                   Else if (pos('SE89',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE89') then
                        sGlkwa := Copy(sGlkwa, 1, 342) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 349, Length(sGlkwa))
                   // 10.SE92
                   Else if (pos('SE92',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE92') then
                        sGlkwa := Copy(sGlkwa, 1, 360) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 367, Length(sGlkwa))
                   // 11.SE96
                   Else if (pos('SE96',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE96') then
                        sGlkwa := Copy(sGlkwa, 1, 222) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 229, Length(sGlkwa))
                   // 12.SE98
                   Else if (pos('SE98',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SE98') then
                        sGlkwa := Copy(sGlkwa, 1, 384) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 391, Length(sGlkwa))
                   // 13.SEA2
                   Else if (pos('SEA2',sHul_List) > 0) and (TG_gumgin.Cells[3,I] = 'SEA2') then
                        sGlkwa := Copy(sGlkwa, 1, 234) + GF_RPad(TG_gumgin.Cells[4,I], 6, ' ')
                             + Copy(sGlkwa, 241, Length(sGlkwa))
                   ELSE
                   begin
                     showmessage(qryGulkwa.ParamByName('num_bunju').AsString +'/'+ qryGulkwa.ParamByName('dat_bunju').AsString+ ' '+
                     TG_gumgin.Cells[2,I] +' => 09파트 항목을 확인해주세요' );
                     exit;
                   end;


                   sGlkwa := 'A' + sGlkwa;
                   sGlkwa := Trim(sGlkwa);
                   sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                   UV_fGulkwa1 := '';
                   UV_fGulkwa2 := '';
                   UV_fGulkwa3 := '';
                   GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                   if copy(GV_sUserId,1,2) <> '51' then qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15'
                   else qryU_Gulkwa.ParamByName('cod_bjjs').AsString   := copy(ComboBox1.Text,1,2);
                   qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := GF_LPad(TG_gumgin.Cells[0,I], 4, '0');
                   qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                   qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                   qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                   qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                   //part구분
                   if      (TG_gumgin.Cells[3,I] = 'SE42') or (TG_gumgin.Cells[3,I] = 'SE85')
                        or (TG_gumgin.Cells[3,I] = 'SE87') or (TG_gumgin.Cells[3,I] = 'SE89')
                        or (TG_gumgin.Cells[3,I] = 'SE90') or (TG_gumgin.Cells[3,I] = 'SE92')
                        or (TG_gumgin.Cells[3,I] = 'SE93') or (TG_gumgin.Cells[3,I] = 'SE96')
                        or (TG_gumgin.Cells[3,I] = 'SE97') or (TG_gumgin.Cells[3,I] = 'SE98')
                        or (TG_gumgin.Cells[3,I] = 'SEA1') or (TG_gumgin.Cells[3,I] = 'SEA2')
                        or (TG_gumgin.Cells[3,I] = 'SEA6') then
                           qryU_Gulkwa.ParamByName('gubn_part').AsString  := '09';


                   qryU_Gulkwa.Execsql;
                   GP_query_log(qryU_Gulkwa, 'LD3I12', 'Q', 'Y');
                   Inc(sIndex);
                END
                ELSE
                begin
                   showmessage('접수되지 않은 항목입니다.');
                   exit;
                end;
              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[0,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[1,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[2,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[3,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[4,I];

                   Inc(spIndex);
              end; // end of else;
         end; // end of with;
         except
               on E : EDBEngineError do
                  begin
                       GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
                       exit;
                  end; // end of except;
         end; // end of try;

   end; // end of for;

   qrdSList.Rows  := sIndex;
   qrdSPList.Rows  := spIndex;

end;

function TfrmLD3I12.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
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
      if qryGulkwa.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_cancer').AsString;
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
      if qryGulkwa.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryGulkwa.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryGulkwa.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '1', qryGulkwa.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGulkwa.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '4', qryGulkwa.FieldByName('Gubn_adyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '7', qryGulkwa.FieldByName('Gubn_lfyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '3',qryGulkwa.FieldByName('Gubn_bgyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '5',qryGulkwa.FieldByName('Gubn_agyh').AsInteger);

   if (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '1') or (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGulkwa.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGulkwa.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGulkwa.FieldByName('Dat_gmgn').AsString;
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
function TfrmLD3I12.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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






