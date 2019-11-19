//==============================================================================
// 프로그램 설명 : [조직]외주항목결과업로드
// 시스템        : 검진관리
// 수정일자      : 2008.06.16
// 수정자        : 유동구
// 수정내용      : 프로그램 추가..
// 참고사항      :
//==============================================================================
unit LD4I01;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD4I01 = class(TfrmSingle)
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
    qryCell: TQuery;
    qryU_Cell: TQuery;
    TG_gumgin: TStringGrid;
    qrdSList: TtsGrid;
    Panel2: TPanel;
    retDesc_Sokun: TMemo;
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
    UV_sJisa, UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4 : String;
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
  frmLD4I01 : TfrmLD4I01;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD4I01.FormCreate(Sender: TObject);
begin
   inherited;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD4I01.UP_SVarMemSet(iLength : Integer);
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
procedure TfrmLD4I01.qrdSListCellLoaded(Sender: TObject; DataCol,
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
procedure TfrmLD4I01.UP_SPVarMemSet(iLength : Integer);
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
procedure TfrmLD4I01.qrdSPListCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD4I01.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD4I01.btnStartClick(Sender: TObject);
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
   For iCol := 1 to 12 do
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
      begin
         if (iCol = 3) and (iRow <> 0) then
         begin
            if pos('-',Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text)) > 0 then
               TG_gumgin.Cells[iCol,iRow] := copy(Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text),1,6) +
                                             copy(Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text),8,7)
            else
               TG_gumgin.Cells[iCol,iRow] := Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text);
         end
         else if (iCol = 6) and (iRow <> 0) then
         begin
            TG_gumgin.Cells[iCol,iRow] := 'T' +
                                          copy(Trim(WorkSheet.Cells[iRow + 1,iCol + 4].Text),3,2) + '-' +
                                          copy(Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text),2,5);
         end
         else
            TG_gumgin.Cells[iCol,iRow] := Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text);
      end;
   end;

   if not VarIsEmpty(Excel) then
     Excel.Quit;
   Screen.Cursor := crDefault;
end;


procedure TfrmLD4I01.Btn_InsertClick(Sender: TObject);
var
    sGlkwa, stemp, sValue : String;
    SUBSOKUN : String;
    I, iTemp, sIndex, spIndex, iValue, sokun_length : Integer;
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
         with qryCell do
         begin
              Active := False;
              qryCell.ParamByName('num_jumin').AsString := TG_gumgin.Cells[3,I];
              qryCell.ParamByName('dat_gmgn').AsString  := TG_gumgin.Cells[9,I];
              qryCell.ParamByName('cod_hm').AsString    := TG_gumgin.Cells[0,I];
              Active := True;

              if qryCell.IsEmpty = False then
              begin
                   GP_query_log(qryCell, 'LD4I01', 'Q', 'Y');
                   UP_SVarMemSet(sIndex);
                   stemp := ''; iValue := 0; sValue := '';
                   
                   UV_vS001[sIndex] := TG_gumgin.Cells[6,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[9,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[3,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[2,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[11,I];
                   UV_vS006[sIndex] := TG_gumgin.Cells[10,I];

                   retDesc_Sokun.Lines.Clear;
                   sValue := TG_gumgin.Cells[10,I];
                   while pos(#10,sValue) > 0 do
                   begin
                      iValue := pos(#10,sValue);
                      stemp  := copy(sValue,1, iValue - 1);
                      retDesc_Sokun.Lines.Add(stemp);
                      sValue := copy(sValue,iValue + 1, length(sValue));
                   end;

                   retDesc_Sokun.Lines.Add(sValue);

                   qryU_Cell.ParamByName('COD_BJJS').AsString    := qryCell.FieldByName('COD_BJJS').AsString;
                   qryU_Cell.ParamByName('NUM_ID').AsString      := qryCell.FieldByName('NUM_ID').AsString;
                   qryU_Cell.ParamByName('COD_CELL').AsString    := qryCell.FieldByName('COD_CELL').AsString;
                   qryU_Cell.ParamByName('DAT_GMGN').AsString    := qryCell.FieldByName('DAT_GMGN').AsString;
                   qryU_Cell.ParamByName('gubn_order').AsString  := qryCell.FieldByName('gubn_order').AsString;


                   qryU_Cell.ParamByName('gubn_jongru').AsString := '4';
                   qryU_Cell.ParamByName('dat_result').AsString  := GV_sToday;
                   qryU_Cell.ParamByName('cod_sawon').AsString   := GV_sUserId;
                   qryU_Cell.ParamByName('cod_doct').AsString    := '158877';

                   SUBSOKUN    := retDesc_Sokun.Text;
                   GF_AllSokun('1', UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4, SUBSOKUN);
                   qryU_Cell.ParamByName('desc_sokun1').AsString := UV_SBSOKUN1;
                   qryU_Cell.ParamByName('desc_sokun2').AsString := UV_SBSOKUN2;
                   qryU_Cell.ParamByName('desc_sokun3').AsString := UV_SBSOKUN3;
                   qryU_Cell.ParamByName('desc_sokun4').AsString := UV_SBSOKUN4;
//                   qryU_Cell.ParamByName('desc_sokun1').AsString := copy(retDesc_Sokun.Text,1,250);
//                   qryU_Cell.ParamByName('desc_sokun2').AsString := copy(retDesc_Sokun.Text,251,250);
//                   qryU_Cell.ParamByName('desc_sokun3').AsString := copy(retDesc_Sokun.Text,501,250);
//                   qryU_Cell.ParamByName('desc_sokun4').AsString := copy(retDesc_Sokun.Text,751,250);
                   sokun_length := length(UV_SBSOKUN1) + length(UV_SBSOKUN2) + length(UV_SBSOKUN3) + length(UV_SBSOKUN4);
                   qryU_Cell.ParamByName('desc_sokun5').AsString := copy(retDesc_Sokun.Text, sokun_length + 1, length(retDesc_Sokun.Text) - sokun_length);
                   qryU_Cell.ParamByName('ysno_sokun').AsString  := 'Y';
                   qryU_Cell.ParamByName('desc_buwi').AsString   := TG_gumgin.Cells[6,I];
                   qryU_Cell.ParamByName('Check_Year').AsString  := copy(TG_gumgin.Cells[9,I],3,2);
                   qryU_Cell.ParamByName('cod_hm').AsString      := TG_gumgin.Cells[0,I];
                   qryU_Cell.ParamByName('DAT_UPDATE').AsString  := GV_sToday;
                   qryU_Cell.ParamByName('COD_UPDATE').AsString  := GV_sUserId;
                   qryU_Cell.Execsql;
                   GP_query_log(qryU_Cell, 'LD4I01', 'Q', 'Y');
                   Inc(sIndex);
              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[6,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[9,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[3,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[2,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[11,I];
                   UV_vSP006[spIndex] := TG_gumgin.Cells[10,I];

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






