//==============================================================================
// 프로그램 설명 : 외주항목 혈액결과 입력 (랩커어)
// 시스템        : 통합검진
// 수정일자      : 2013.12.7
// 수정자        : 김희석
// 수정내용      : 혈액결과 업로드...
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.03.21
// 수정자        : 곽윤설
// 수정내용      : 랩케어2014 항목 추가 (T012, T009, S019, S018, C070, C088, C089)
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.26
// 수정자        : 곽윤설
// 수정내용      : T040항목 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.11.07
// 수정자        : 곽윤설
// 수정내용      : S021항목 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.30
// 수정자        : 곽윤설
// 수정내용      : T009, T012 삭제
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.01.22
// 수정자        : 곽윤설
// 수정내용      : T009 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.04.02
// 수정자        : 곽윤설
// 수정내용      : T012 추가
// 참고사항      :
//==============================================================================

unit LD3I09;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD3I09 = class(TfrmSingle)
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
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    qryJHangmokList: TQuery;
    qryNo_hangmok: TQuery;
    qryProfileG: TQuery;
    qryPf_hm2: TQuery;
  procedure FormCreate(Sender: TObject);
    procedure btnDESC_DEPTClick(Sender: TObject);
    procedure Btn_InsertClick(Sender: TObject);
    procedure qrdSListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure qrdSPListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnStartClick(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

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
  frmLD3I09 : TfrmLD3I09;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I09.FormCreate(Sender: TObject);
begin
   inherited;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I09.UP_SVarMemSet(iLength : Integer);
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
procedure TfrmLD3I09.qrdSListCellLoaded(Sender: TObject; DataCol,
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
procedure TfrmLD3I09.UP_SPVarMemSet(iLength : Integer);
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
procedure TfrmLD3I09.qrdSPListCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD3I09.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I09.btnStartClick(Sender: TObject);
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


procedure TfrmLD3I09.Btn_InsertClick(Sender: TObject);
var
    sGlkwa,sHul_List,temp2,Cell_temp : String;
    I, iTemp, sIndex,temp1, spIndex: Integer;
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
              qryGulkwa.ParamByName('dat_bunju').AsString  := TG_gumgin.Cells[1,I];
              qryGulkwa.ParamByName('num_bunju').AsString  := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
              //part구분
              if ((TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or
                  (TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099') or
                  (TG_gumgin.Cells[6,I] = 'TT01') or (TG_gumgin.Cells[6,I] = 'TT02')) and (copy(TG_gumgin.Cells[8,I], 1,8) <> 'Grayzone') then
                  ParamByName('gubn_part').AsString  := '07'
              else if ((TG_gumgin.Cells[6,I] = 'T043') or (TG_gumgin.Cells[6,I] = 'T041') or
                       (TG_gumgin.Cells[6,I] = 'SE01') or (TG_gumgin.Cells[6,I] = 'SE02') or
                       (TG_gumgin.Cells[6,I] = 'S019') or
                       (TG_gumgin.Cells[6,I] = 'S018') or                                     
                       (TG_gumgin.Cells[6,I] = 'S010') or (TG_gumgin.Cells[6,I] = 'E069') or
                       (TG_gumgin.Cells[6,I] = 'S020') or (TG_gumgin.Cells[6,I] = 'S021') or
                       (TG_gumgin.Cells[6,I] = 'T040') or (TG_gumgin.Cells[6,I] = 'T009') or
                       (TG_gumgin.Cells[6,I] = 'E068')) then
                       ParamByName('gubn_part').AsString  := '05'
              else if ((TG_gumgin.Cells[6,I] = 'C070') or (TG_gumgin.Cells[6,I] = 'C088') or
                       (TG_gumgin.Cells[6,I] = 'C089')) then
                       ParamByName('gubn_part').AsString  := '01'
              else if ((TG_gumgin.Cells[6,I] = 'T012')) then
                       ParamByName('gubn_part').AsString  := '08'
              else if ((TG_gumgin.Cells[6,I] = 'E040')) then
                       ParamByName('gubn_part').AsString  := '04'
              Else
                  ParamByName('gubn_part').AsString  := '99'; //해당항목이 아닐경우 조회 안 되게 하기 위해서..

              Active := True;
              if RecordCount > 0 then
              begin
                   Cell_temp:='';
                   GP_query_log(qryGulkwa, 'LD3I09', 'Q', 'Y');
                   sHul_List :='';
                   sHul_List := UF_hangmokList;
                   Cell_temp:= TG_gumgin.Cells[8,I]; // 결과
                   //랩케어 액양성 결과 없음으로 인해 약양성 기준치 추가
                   if ((TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or (TG_gumgin.Cells[6,I] = 'SE01') or
                       (TG_gumgin.Cells[6,I] = 'SE02') or (TG_gumgin.Cells[6,I] = 'S020') or (TG_gumgin.Cells[6,I] = 'S010') or
                       (TG_gumgin.Cells[6,I] = 'S018') or (TG_gumgin.Cells[6,I] = 'S019') or (TG_gumgin.Cells[6,I] = 'T040')) then
                   begin
                        if ((Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) >= 1) and
                            (Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) < 10)) and
                            (TG_gumgin.Cells[6,I] = 'S007') then
                             TG_gumgin.Cells[8,I] := '약양성'
                        else  if ((Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) >= 10) and
                            (Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) < 30)) and
                            (TG_gumgin.Cells[6,I] = 'S008') then
                             TG_gumgin.Cells[8,I] := '약양성'
                        else  if ((Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) >= 0.80) and
                            (Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) < 1.20)) and
                            (TG_gumgin.Cells[6,I] = 'SE01') then
                             TG_gumgin.Cells[8,I] := '약양성'
                        else  if ((Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) >= 1) and
                            (Strtofloat((copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1))) < 1.19)) and
                            ((TG_gumgin.Cells[6,I] = 'SE02') or (TG_gumgin.Cells[6,I] = 'S010')) then
                             TG_gumgin.Cells[8,I] := '약양성'
                        else if (copy(TG_gumgin.Cells[8,I], 1,8) = 'Negative') or (copy(TG_gumgin.Cells[8,I], 1,8) = 'negative') then
                             TG_gumgin.Cells[8,I] := '음성'
                        else if (copy(TG_gumgin.Cells[8,I], 1,8) = 'Positive') or (copy(TG_gumgin.Cells[8,I], 1,8) = 'positive') then
                             TG_gumgin.Cells[8,I] := '양성';
                   end
                   else if ((TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099')) then
                   begin
                        TG_gumgin.Cells[8,I] :=(copy(TG_gumgin.Cells[8,I],pos('(',TG_gumgin.Cells[8,I])+1,length(TG_gumgin.Cells[8,I])-pos('(',TG_gumgin.Cells[8,I])-1));
                   end
                   else if ((TG_gumgin.Cells[6,I] = 'TT01')) then
                   begin
                        if StrTofloat(TG_gumgin.Cells[8,I])<10 then
                           TG_gumgin.Cells[8,I]:='음성'
                        else if StrTofloat(TG_gumgin.Cells[8,I])>=10 then
                           TG_gumgin.Cells[8,I]:='양성';
                   end
                   else if ((TG_gumgin.Cells[6,I] = 'TT02') or (TG_gumgin.Cells[6,I] = 'T043') or (TG_gumgin.Cells[6,I] = 'T041') or (TG_gumgin.Cells[6,I] = 'E069') or
                            (TG_gumgin.Cells[6,I] = 'T012') or (TG_gumgin.Cells[6,I] = 'C070') or (TG_gumgin.Cells[6,I] = 'C088') or (TG_gumgin.Cells[6,I] = 'C089') or
                            (TG_gumgin.Cells[6,I] = 'T009') or (TG_gumgin.Cells[6,I] = 'E040') or (TG_gumgin.Cells[6,I] = 'E068')) then
                   begin
                        TG_gumgin.Cells[8,I]:=TG_gumgin.Cells[8,I];
                   end;

                   UP_SVarMemSet(sIndex);
                   UV_vS001[sIndex] := TG_gumgin.Cells[1,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[2,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[4,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[6,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[8,I];
                   Inc(sIndex);

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


                   //S099
                   if TG_gumgin.Cells[6,I] = 'S099' then
                        sGlkwa := Copy(sGlkwa, 1, 228) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 235, Length(sGlkwa))
                   //S091
                   else if TG_gumgin.Cells[6,I] = 'S091' then
                        sGlkwa := Copy(sGlkwa, 1, 180) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 187, Length(sGlkwa))
                   //TT02
                   else if TG_gumgin.Cells[6,I] = 'TT02' then
                        sGlkwa := Copy(sGlkwa, 1, 168) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 175, Length(sGlkwa))
                   //TT01
                   else if TG_gumgin.Cells[6,I] = 'TT01' then
                        sGlkwa := Copy(sGlkwa, 1, 84) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 91, Length(sGlkwa))
                   //S008
                   else if TG_gumgin.Cells[6,I] = 'S008' then
                        sGlkwa := Copy(sGlkwa, 1, 36) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 43, Length(sGlkwa))
                   //S007
                   else if TG_gumgin.Cells[6,I] = 'S007' then
                        sGlkwa := Copy(sGlkwa, 1, 30) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 37, Length(sGlkwa))
                   //T043
                   else if TG_gumgin.Cells[6,I] = 'T043' then
                        sGlkwa := Copy(sGlkwa, 1, 360) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 367, Length(sGlkwa))
                   //T041
                   else if TG_gumgin.Cells[6,I] = 'T041' then
                        sGlkwa := Copy(sGlkwa, 1, 348) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 355, Length(sGlkwa))
                   //SE01
                   else if TG_gumgin.Cells[6,I] = 'SE01' then
                        sGlkwa := Copy(sGlkwa, 1, 246) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 253, Length(sGlkwa))
                   //SE02
                   else if TG_gumgin.Cells[6,I] = 'SE02' then
                        sGlkwa := Copy(sGlkwa, 1, 252) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 259, Length(sGlkwa))
                   //S020
                   else if TG_gumgin.Cells[6,I] = 'S020' then
                        sGlkwa := Copy(sGlkwa, 1, 0) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 7, Length(sGlkwa))
                   //S021
                   else if TG_gumgin.Cells[6,I] = 'S021' then
                        sGlkwa := Copy(sGlkwa, 43, 0) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 49, Length(sGlkwa))
                   //E069
                   else if TG_gumgin.Cells[6,I] = 'E069' then
                        sGlkwa := Copy(sGlkwa, 1, 432) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 439, Length(sGlkwa))
                   //S010
                   else if TG_gumgin.Cells[6,I] = 'S010' then
                        sGlkwa := Copy(sGlkwa, 1, 36) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 43, Length(sGlkwa))
                   //S018
                   else if TG_gumgin.Cells[6,I] = 'S018' then
                        sGlkwa := Copy(sGlkwa, 1, 72) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 79, Length(sGlkwa))
                   //S019
                   else if TG_gumgin.Cells[6,I] = 'S019' then
                        sGlkwa := Copy(sGlkwa, 1, 78) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 85, Length(sGlkwa))
                 
                   //C070                                     //T012
                   else if TG_gumgin.Cells[6,I] = 'T012' then
                        sGlkwa := Copy(sGlkwa, 1, 504) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 511, Length(sGlkwa))
                   else if TG_gumgin.Cells[6,I] = 'C070' then
                        sGlkwa := Copy(sGlkwa, 1, 312) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 319, Length(sGlkwa))
                   //C088
                   else if TG_gumgin.Cells[6,I] = 'C088' then
                        sGlkwa := Copy(sGlkwa, 1, 420) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 427, Length(sGlkwa))
                   //C089
                   else if TG_gumgin.Cells[6,I] = 'C089' then
                        sGlkwa := Copy(sGlkwa, 1, 426) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 433, Length(sGlkwa))
                   //T040
                   else if TG_gumgin.Cells[6,I] = 'T040' then
                        sGlkwa := Copy(sGlkwa, 1, 342) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 349, Length(sGlkwa))
                   //T009
                   else if TG_gumgin.Cells[6,I] = 'T009' then
                        sGlkwa := Copy(sGlkwa, 1, 48) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 55, Length(sGlkwa))

                   //E040
                   else if TG_gumgin.Cells[6,I] = 'E040' then
                        sGlkwa := Copy(sGlkwa, 1, 60) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 67, Length(sGlkwa))

                   //E068
                   else if TG_gumgin.Cells[6,I] = 'E068' then
                        sGlkwa := Copy(sGlkwa, 1, 426) + GF_RPad(TG_gumgin.Cells[8,I], 6, ' ')
                               + Copy(sGlkwa, 433, Length(sGlkwa));


                   sGlkwa := 'A' + sGlkwa;
                   sGlkwa := Trim(sGlkwa);
                   sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                   UV_fGulkwa1 := '';
                   UV_fGulkwa2 := '';
                   UV_fGulkwa3 := '';
                   GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                   qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                   qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                   qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
                   qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                   qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                   qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                   //part구분
                   if (TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or (TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099') or
                      (TG_gumgin.Cells[6,I] = 'TT01') or (TG_gumgin.Cells[6,I] = 'TT02')  then
                      qryU_Gulkwa.ParamByName('gubn_part').AsString  := '07'
                   else if ((TG_gumgin.Cells[6,I] = 'T043') or (TG_gumgin.Cells[6,I] = 'T041') or
                            (TG_gumgin.Cells[6,I] = 'SE01') or (TG_gumgin.Cells[6,I] = 'SE02') or
                            (TG_gumgin.Cells[6,I] = 'S020') or (TG_gumgin.Cells[6,I] = 'E069') or
                            (TG_gumgin.Cells[6,I] = 'S010') or (TG_gumgin.Cells[6,I] = 'S018') or
                            (TG_gumgin.Cells[6,I] = 'S019') or (TG_gumgin.Cells[6,I] = 'T040') or
                            (TG_gumgin.Cells[6,I] = 'T009') or (TG_gumgin.Cells[6,I] = 'E068')) then
                        qryU_Gulkwa.ParamByName('gubn_part').AsString  := '05'
                   else if ((TG_gumgin.Cells[6,I] = 'C070') or (TG_gumgin.Cells[6,I] = 'C088') or
                            (TG_gumgin.Cells[6,I] = 'C089')) then
                        qryU_Gulkwa.ParamByName('gubn_part').AsString  := '01'
                   else if ((TG_gumgin.Cells[6,I] = 'T012')) then
                        qryU_Gulkwa.ParamByName('gubn_part').AsString  := '08'
                   else if ((TG_gumgin.Cells[6,I] = 'E040')) then
                        qryU_Gulkwa.ParamByName('gubn_part').AsString  := '04'
                   else
                      qryU_Gulkwa.ParamByName('gubn_part').AsString  := '99';   // 해당항목이 아닐경우 update 안되게 하기 위해서..

                   qryU_Gulkwa.Execsql;

                   if ((TG_gumgin.Cells[6,I] = 'S007') and (pos('S099',sHul_List)>0)) or
                      ((TG_gumgin.Cells[6,I] = 'S008') and (pos('S091',sHul_List)>0)) or
                      ((TG_gumgin.Cells[6,I] = 'TT01') and (pos('TT02',sHul_List)>0)) then
                   begin
                      Active := False;
                      ParamByName('cod_bjjs').AsString   := '15';
                      qryGulkwa.ParamByName('dat_bunju').AsString  := TG_gumgin.Cells[1,I];
                      qryGulkwa.ParamByName('num_bunju').AsString  := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
                      qryGulkwa.ParamByName('gubn_part').AsString  := '07';
                      Active := True;
                      if ((TG_gumgin.Cells[6,I] = 'S007') and (pos('S099',sHul_List)>0)) then
                      begin
                         UP_SVarMemSet(sIndex);
                         Cell_temp:=(copy(Cell_temp,pos('(',Cell_temp)+1,length(Cell_temp)-pos('(',Cell_temp)-1));

                         UV_vS001[sIndex] := TG_gumgin.Cells[1,I];
                         UV_vS002[sIndex] := TG_gumgin.Cells[2,I];
                         UV_vS003[sIndex] := TG_gumgin.Cells[4,I];
                         UV_vS004[sIndex] := 'S099';
                         UV_vS005[sIndex] := Cell_temp;

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

                          sGlkwa := Copy(sGlkwa, 1, 228) + GF_RPad(Cell_temp, 6, ' ')
                                 + Copy(sGlkwa, 235, Length(sGlkwa));

                         sGlkwa := 'A' + sGlkwa;
                         sGlkwa := Trim(sGlkwa);
                         sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                         UV_fGulkwa1 := '';
                         UV_fGulkwa2 := '';
                         UV_fGulkwa3 := '';
                         GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                         qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                         qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                         qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
                         qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                         qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                         qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                         qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                         qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                         //part구분
                         if (TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or (TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099') or
                            (TG_gumgin.Cells[6,I] = 'TT01') or (TG_gumgin.Cells[6,I] = 'TT02')then
                            qryU_Gulkwa.ParamByName('gubn_part').AsString  := '07'
                         else
                            qryU_Gulkwa.ParamByName('gubn_part').AsString  := '99';   // 해당항목이 아닐경우 update 안되게 하기 위해서..

                         qryU_Gulkwa.Execsql;
                         Inc(sIndex);
                      end;
                      if ((TG_gumgin.Cells[6,I] = 'S008') and (pos('S091',sHul_List)>0)) then
                      begin
                         UP_SVarMemSet(sIndex);
                         Cell_temp :=(copy(Cell_temp,pos('(',Cell_temp)+1,length(Cell_temp)-pos('(',Cell_temp)-1));

                         UV_vS001[sIndex] := TG_gumgin.Cells[1,I];
                         UV_vS002[sIndex] := TG_gumgin.Cells[2,I];
                         UV_vS003[sIndex] := TG_gumgin.Cells[4,I];
                         UV_vS004[sIndex] := 'S091';
                         UV_vS005[sIndex] := Cell_temp;

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
                         sGlkwa := Copy(sGlkwa, 1, 180) + GF_RPad(Cell_temp, 6, ' ')
                                 + Copy(sGlkwa, 187, Length(sGlkwa));

                         sGlkwa := 'A' + sGlkwa;
                         sGlkwa := Trim(sGlkwa);
                         sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                         UV_fGulkwa1 := '';
                         UV_fGulkwa2 := '';
                         UV_fGulkwa3 := '';
                         GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                         qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                         qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                         qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
                         qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                         qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                         qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                         qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                         qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                         //part구분
                         if (TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or (TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099') or
                            (TG_gumgin.Cells[6,I] = 'TT01') or (TG_gumgin.Cells[6,I] = 'TT02')then
                            qryU_Gulkwa.ParamByName('gubn_part').AsString  := '07'
                         else
                            qryU_Gulkwa.ParamByName('gubn_part').AsString  := '99';   // 해당항목이 아닐경우 update 안되게 하기 위해서..

                         qryU_Gulkwa.Execsql;
                         Inc(sIndex);
                      end;
                      if ((TG_gumgin.Cells[6,I] = 'TT01') and (pos('TT02',sHul_List)>0)) then
                      begin
                      UP_SVarMemSet(sIndex);

                      UV_vS001[sIndex] := TG_gumgin.Cells[1,I];
                      UV_vS002[sIndex] := TG_gumgin.Cells[2,I];
                      UV_vS003[sIndex] := TG_gumgin.Cells[4,I];
                      UV_vS004[sIndex] := 'TT02';
                      UV_vS005[sIndex] := Cell_temp;

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

                       sGlkwa := Copy(sGlkwa, 1, 168) + GF_RPad(Cell_temp, 6, ' ')
                               + Copy(sGlkwa, 175, Length(sGlkwa));

                       sGlkwa := 'A' + sGlkwa;
                       sGlkwa := Trim(sGlkwa);
                       sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                       UV_fGulkwa1 := '';
                       UV_fGulkwa2 := '';
                       UV_fGulkwa3 := '';
                       GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                       qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                       qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[1,I];
                       qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := GF_LPad(TG_gumgin.Cells[2,I], 4, '0');
                       qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                       qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                       qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                       qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                       qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;

                       //part구분
                       if (TG_gumgin.Cells[6,I] = 'S007') or (TG_gumgin.Cells[6,I] = 'S008') or (TG_gumgin.Cells[6,I] = 'S091') or (TG_gumgin.Cells[6,I] = 'S099') or
                          (TG_gumgin.Cells[6,I] = 'TT01') or (TG_gumgin.Cells[6,I] = 'TT02') then
                          qryU_Gulkwa.ParamByName('gubn_part').AsString  := '07'
                       else
                          qryU_Gulkwa.ParamByName('gubn_part').AsString  := '99';   // 해당항목이 아닐경우 update 안되게 하기 위해서..

                       qryU_Gulkwa.Execsql;
                       Inc(sIndex);
                   end;
                   Active := False;
                   end;

              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[1,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[2,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[4,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[7,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[8,I];

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
function TfrmLD3I09.UF_hangmokList : String;
var
   i ,k : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList,sCod_hm,cod_name : string;
       sHmCode :string;
   vCod_chuga, vCod_totyh, vCod_profile : Variant;
   iRet,num : integer;
      check_tk : boolean;
begin
   //검사항목축출
            sCod_hm := '';
            if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
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

            if qryGulkwa.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := qryGulkwa.FieldByName('cod_cancer').AsString;
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

            if qryGulkwa.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := qryGulkwa.FieldByName('cod_jangbi').AsString;
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
            iRet := GF_MulToSingle(qryGulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_chuga[i];

           //-------------------------------------------------------------------

            //---- 특검항목 추출...
            iRet := GF_MulToSingle(qryGulkwa.FieldByName('cod_prf').AsString, 4, vCod_profile);
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
            sCod_hm := sCod_hm + qryGulkwa.FieldByName('TkGum_chuga').AsString;

//------------------------------------------------------------------------------

           // 노신, 성인병, 간염에 대한 검사항목 추출
            cod_name := '';
            // 노신유형Display
            if qryGulkwa.FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qryGulkwa.FieldByName('gubn_nsyh').AsInteger);
            //성인병유형Display
            if qryGulkwa.FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qryGulkwa.FieldByName('gubn_adyh').AsInteger);
            //간염유형Display
            if qryGulkwa.FieldByName('gubn_agab').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qryGulkwa.FieldByName('gubn_agyh').AsInteger);
            //생애전환기유형Display
            if qryGulkwa.FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qryGulkwa.FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_totyh[i];

   for k:=0 to  Round(Length(sCod_hm)/4)-1 do
   begin
   result:=result+Copy(sCod_hm, (k+1)*4-3, 4)+'.';
   end;

end;

function TfrmLD3I09.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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






