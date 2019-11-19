//==============================================================================
// 프로그램 설명 : 외주항목 혈액결과 입력
// 시스템        : 통합검진
// 수정일자      : 2011.06.10
// 수정자        : 송재호
// 수정내용      : 혈액결과 업로드...
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.16
// 수정자        : 곽윤설
// 수정내용      : 기존코드 모두 제거, U057추가
// 참고사항      :
//==============================================================================
unit LD3I08;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD3I08 = class(TfrmSingle)
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

    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005 : Variant;
    UV_vSP001, UV_vSP002, UV_vSP003, UV_vSP004, UV_vSP005 : Variant;
    
    procedure UP_SVarMemSet(iLength : Integer);
    procedure UP_SPVarMemSet(iLength : Integer);


  public
    { Public declarations }
    UV_sValue  : array[1..8] of String;

  end;


var
  frmLD3I08 : TfrmLD3I08;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I08.FormCreate(Sender: TObject);
begin
   inherited;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I08.UP_SVarMemSet(iLength : Integer);
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
procedure TfrmLD3I08.qrdSListCellLoaded(Sender: TObject; DataCol,
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
procedure TfrmLD3I08.UP_SPVarMemSet(iLength : Integer);
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
procedure TfrmLD3I08.qrdSPListCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD3I08.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I08.btnStartClick(Sender: TObject);
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
   For iCol := 1 to 10 do
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


procedure TfrmLD3I08.Btn_InsertClick(Sender: TObject);
var
    sGlkwa : String;
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
              if (TG_gumgin.Cells[4,I] = 'U057') then
                 ParamByName('gubn_part').AsString  := '03'
              Else
                 ParamByName('gubn_part').AsString  := '99'; //해당항목이 아닐경우 조회 안 되게 하기 위해서..

              Active := True;
              if RecordCount > 0 then
              begin
                   GP_query_log(qryGulkwa, 'LD3I08', 'Q', 'Y');
                   if (copy(TG_gumgin.Cells[6,I], 1,8) = 'Negative') or (copy(TG_gumgin.Cells[6,I], 1,8) = 'negative') then
                      TG_gumgin.Cells[6,I] := '음성'
                   Else if (copy(TG_gumgin.Cells[6,I], 1,8) = 'Positive') or (copy(TG_gumgin.Cells[6,I], 1,8) = 'positive') then
                      TG_gumgin.Cells[6,I] := '양성'
                   Else if (copy(TG_gumgin.Cells[6,I], 1,15) = 'Weakly Positive') or (copy(TG_gumgin.Cells[6,I], 1,15) = 'weakly Positive') or
                           (copy(TG_gumgin.Cells[6,I], 1,9)  = 'Equivocal') then
                      TG_gumgin.Cells[6,I] := '약양성'
                   Else if copy(TG_gumgin.Cells[6,I], 1,2) = '< ' then   // 부등호가 들어간 10.0 미만 결과 처리
                      TG_gumgin.Cells[6,I] := copy(TG_gumgin.Cells[6,I], 4, length(TG_gumgin.Cells[6,I]) - 3);

                   UP_SVarMemSet(sIndex);

                   UV_vS001[sIndex] := TG_gumgin.Cells[0,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[1,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[3,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[4,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[6,I];

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

                   // T039
                   if TG_gumgin.Cells[4,I] = 'T039' then
                        sGlkwa := Copy(sGlkwa, 1, 360) + GF_RPad(TG_gumgin.Cells[6,I], 6, ' ')
                               + Copy(sGlkwa, 367, Length(sGlkwa));

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
                   if (TG_gumgin.Cells[4,I] = 'U057') then
                      qryU_Gulkwa.ParamByName('gubn_part').AsString  := '03'
                   else
                      qryU_Gulkwa.ParamByName('gubn_part').AsString  := '99';   // 해당항목이 아닐경우 update 안되게 하기 위해서..

                   qryU_Gulkwa.Execsql;
                   GP_query_log(qryU_Gulkwa, 'LD3I08', 'Q', 'Y');
                   Inc(sIndex);

              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[0,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[1,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[3,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[4,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[6,I];

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






