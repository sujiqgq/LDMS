//==============================================================================
// 프로그램 설명 : 외주항목 혈액결과 입력(씨젠)
// 시스템        : 통합검진
// 수정일자      : 2016.03.29
// 수정자        : 박수지
// 수정내용      :
// 참고사항      : 한의 재단진단검사의학팀1600362
//==============================================================================
unit LD3I11;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD3I11 = class(TfrmSingle)
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
  procedure FormCreate(Sender: TObject);
    procedure btnDESC_DEPTClick(Sender: TObject);
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

    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005, UV_vS006 : Variant;
    UV_vSP001, UV_vSP002, UV_vSP003, UV_vSP004, UV_vSP005, UV_vSP006 : Variant;
    
    procedure UP_SVarMemSet(iLength : Integer);
    procedure UP_SPVarMemSet(iLength : Integer);


  public
    { Public declarations }
    UV_sValue  : array[1..8] of String;

  end;


var
  frmLD3I11 : TfrmLD3I11;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I11.FormCreate(Sender: TObject);
begin
   inherited;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I11.UP_SVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vS001 := VarArrayCreate([0, 0], varOleStr);
      UV_vS002 := VarArrayCreate([0, 0], varOleStr);
      UV_vS003 := VarArrayCreate([0, 0], varOleStr);
      UV_vS004 := VarArrayCreate([0, 0], varOleStr);
      UV_vS005 := VarArrayCreate([0, 0], varOleStr);
      UV_vS006 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vS001, iLength);
      VarArrayRedim(UV_vS002, iLength);
      VarArrayRedim(UV_vS003, iLength);
      VarArrayRedim(UV_vS004, iLength);
      VarArrayRedim(UV_vS005, iLength);
      VarArrayRedim(UV_vS006, iLength);
   end;
end;
procedure TfrmLD3I11.qrdSListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vS001[DataRow - 1];
      2 : Value := UV_vS002[DataRow - 1];
      3 : Value := UV_vS003[DataRow - 1];
      4 : Value := UV_vS004[DataRow - 1];
      5 : Value := UV_vS005[DataRow - 1];
      6 : Value := UV_vS006[DataRow - 1];
   end;
end;
procedure TfrmLD3I11.UP_SPVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vSP001 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP002 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP003 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP004 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP005 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP006 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vSP001, iLength);
      VarArrayRedim(UV_vSP002, iLength);
      VarArrayRedim(UV_vSP003, iLength);
      VarArrayRedim(UV_vSP004, iLength);
      VarArrayRedim(UV_vSP005, iLength);
      VarArrayRedim(UV_vSP006, iLength);
   end;
end;
procedure TfrmLD3I11.qrdSPListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vSP001[DataRow - 1];
      2 : Value := UV_vSP002[DataRow - 1];
      3 : Value := UV_vSP003[DataRow - 1];
      4 : Value := UV_vSP004[DataRow - 1];
      5 : Value := UV_vSP005[DataRow - 1];
      6 : Value := UV_vSP006[DataRow - 1];
   end;
end;

procedure TfrmLD3I11.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I11.btnStartClick(Sender: TObject);
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


procedure TfrmLD3I11.Btn_InsertClick(Sender: TObject);
var
    sGlkwa,S : String;
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
              ParamByName('cod_bjjs').AsString   := '15';
              qryGulkwa.ParamByName('dat_bunju').AsString  := TG_gumgin.Cells[0,I];
              qryGulkwa.ParamByName('num_bunju').AsString  := TG_gumgin.Cells[1,I];

              //part구분
              if (TG_gumgin.Cells[4,I] = 'U059') or (TG_gumgin.Cells[4,I] = 'U060')
                   or (TG_gumgin.Cells[4,I] = 'U061') or (TG_gumgin.Cells[4,I] = 'U063')
                   or (TG_gumgin.Cells[4,I] = 'U064') or (TG_gumgin.Cells[4,I] = 'U065') then
                      ParamByName('gubn_part').AsString  := '03'
              Else if (TG_gumgin.Cells[4,I] = 'P007') then
                      ParamByName('gubn_part').AsString  := '07'
              Else
                      ParamByName('gubn_part').AsString  := '';

              Active := True;
              if RecordCount > 0 then
              begin

                   if (copy(TG_gumgin.Cells[5,I], 1,8) = 'Negative') or (copy(TG_gumgin.Cells[5,I], 1,8) = 'negative') then
                      TG_gumgin.Cells[5,I] := '음성'
                   else if (copy(TG_gumgin.Cells[5,I], 1,8) = 'Positive') or (copy(TG_gumgin.Cells[5,I], 1,8) = 'positive') then
                      TG_gumgin.Cells[5,I] := '양성';


                   UP_SVarMemSet(sIndex);

                   UV_vS001[sIndex] := TG_gumgin.Cells[0,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[1,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[2,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[3,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[4,I];
                   UV_vS006[sIndex] := TG_gumgin.Cells[5,I];

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

                   // 1.P007  ----------------------------------------------------------------------------------------- 07
                   if TG_gumgin.Cells[4,I] = 'P007' then
                        sGlkwa := Copy(sGlkwa, 1, 96) + GF_RPad(TG_gumgin.Cells[5,I], 6, ' ')
                             + Copy(sGlkwa, 103, Length(sGlkwa))
                   // 2.U059 ------------------------------------------------------------------------------------------ 03
                   Else if TG_gumgin.Cells[4,I] = 'U059' then
                        sGlkwa := Copy(sGlkwa, 1, 372) + GF_RPad(TG_gumgin.Cells[5,I], 6, ' ')
                             + Copy(sGlkwa, 379, Length(sGlkwa))
                   // 3.U063
                   Else if TG_gumgin.Cells[4,I] = 'U063' then
                        sGlkwa := Copy(sGlkwa, 1, 397) + GF_RPad(TG_gumgin.Cells[5,I], 6, ' ')
                             + Copy(sGlkwa, 404, Length(sGlkwa))
                   // 4.U064
                   Else if TG_gumgin.Cells[4,I] = 'U064' then
                        sGlkwa := Copy(sGlkwa, 1, 403) + GF_RPad(TG_gumgin.Cells[5,I], 6, ' ')
                             + Copy(sGlkwa, 410, Length(sGlkwa))
                   // 5.U065
                   Else if TG_gumgin.Cells[4,I] = 'U065' then
                        sGlkwa := Copy(sGlkwa, 1, 409) + GF_RPad(TG_gumgin.Cells[5,I], 6, ' ')
                             + Copy(sGlkwa, 416, Length(sGlkwa)) ;

                   sGlkwa := 'A' + sGlkwa;
                   sGlkwa := Trim(sGlkwa);
                   sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                   UV_fGulkwa1 := '';
                   UV_fGulkwa2 := '';
                   UV_fGulkwa3 := '';
                   GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                   qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                   qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[0,I];
                   qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                   qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                   qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                   qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                   //part구분
                   if    (TG_gumgin.Cells[4,I] = 'U059') or (TG_gumgin.Cells[4,I] = 'U063')
                      or (TG_gumgin.Cells[4,I] = 'U064') or (TG_gumgin.Cells[4,I] = 'U065') then
                           qryU_Gulkwa.ParamByName('gubn_part').AsString  := '03'
                   Else if (TG_gumgin.Cells[4,I] = 'P007') then
                           qryU_Gulkwa.ParamByName('gubn_part').AsString  := '07';


                   qryU_Gulkwa.Execsql;
                   GP_query_log(qryU_Gulkwa, 'LD3I11', 'Q', 'Y');
                   Inc(sIndex);

              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[0,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[1,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[2,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[3,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[4,I];
                   UV_vSP006[spIndex] := TG_gumgin.Cells[5,I];

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

end.






