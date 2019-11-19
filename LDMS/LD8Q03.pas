//==============================================================================
// 프로그램 설명 : 항목별 지사별 검진인원 현황
// 시스템        : 통합검진
// 수정일자      : 1999.11.3
// 수정자        : 허정남
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8Q03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD8Q03 = class(TfrmSingle)
    qrygulkwa: TQuery;
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    mskST_date: TDateEdit;
    btnDate: TSpeedButton;
    mskED_date: TDateEdit;
    btnDATE1: TSpeedButton;
    Label1: TLabel;
    pnlJisa: TPanel;
    Label3: TLabel;
    cmbJisa: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_vCod_hm, UV_vDesc_kor,  UV_vJisa_15, UV_vJisa_11, UV_vJisa_33,  UV_vJisa_34,
    UV_vJisa_35, UV_vJisa_41, UV_vJisa_43,  UV_vJisa_45, UV_vJisa_46,  UV_vJisa_47,
    UV_vJisa_51, UV_vJisa_52, UV_vJisa_61,  UV_vJisa_63, UV_vJisa_71,  UV_vJisa_72,
    UV_vJisa_kita, UV_vTotal : Variant;
  end;

var
  frmLD8Q03: TfrmLD8Q03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q031;

{$R *.DFM}
procedure TfrmLD8Q03.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading := '검사코드';
      Col[2].Heading := '검 사 명';
      Col[3].Heading := '연구소';
      Col[4].Heading := '강남';
      Col[5].Heading := '청주';
      Col[6].Heading := '대전';
      Col[7].Heading := '천안';
      Col[8].Heading := '인천';
      Col[9].Heading := '수원';
      Col[10].Heading := '안양';
      Col[11].Heading := '의정부';
      Col[12].Heading := '시흥';
      Col[13].Heading := '광주';
      Col[14].Heading := '전주';
      Col[15].Heading := '부산';
      Col[16].Heading := '울산';
      Col[17].Heading := '대구';
      Col[18].Heading := '포항';
      Col[19].Heading := '기타';
      Col[20].Heading := '합계';
      // Column Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
   end;
end;

procedure TfrmLD8Q03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vCod_hm    := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_kor  := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_15   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_11   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_33   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_34   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_35   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_41   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_43   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_45   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_46   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_47   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_51   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_52   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_61   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_63   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_71   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_72   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_kita := VarArrayCreate([0, 0], varOleStr);
      UV_vTotal     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vCod_hm,    iLength);
      VarArrayRedim(UV_vdesc_kor,  iLength);
      VarArrayRedim(UV_vJisa_15,   iLength);
      VarArrayRedim(UV_vJisa_11,   iLength);
      VarArrayRedim(UV_vJisa_33,   iLength);
      VarArrayRedim(UV_vJisa_34,   iLength);
      VarArrayRedim(UV_vJisa_35,   iLength);
      VarArrayRedim(UV_vJisa_41,   iLength);
      VarArrayRedim(UV_vJisa_43,   iLength);
      VarArrayRedim(UV_vJisa_45,   iLength);
      VarArrayRedim(UV_vJisa_46,   iLength);
      VarArrayRedim(UV_vJisa_47,   iLength);
      VarArrayRedim(UV_vJisa_51,   iLength);
      VarArrayRedim(UV_vJisa_52,   iLength);
      VarArrayRedim(UV_vJisa_61,   iLength);
      VarArrayRedim(UV_vJisa_63,   iLength);
      VarArrayRedim(UV_vJisa_71,   iLength);
      VarArrayRedim(UV_vJisa_72,   iLength);
      VarArrayRedim(UV_vJisa_kita, iLength);
      VarArrayRedim(UV_vTotal,     iLength);
   end;
end;

function TfrmLD8Q03.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskST_Date.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD8Q03.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;
   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);
   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

end;

procedure TfrmLD8Q03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_hm[DataRow-1];
      2 : Value := UV_vDesc_kor[DataRow-1];
      3 : Value := UV_vJisa_15[DataRow-1];
      4 : Value := UV_vJisa_11[DataRow-1];
      5 : Value := UV_vJisa_33[DataRow-1];
      6 : Value := UV_vJisa_34[DataRow-1];
      7 : Value := UV_vJisa_35[DataRow-1];
      8 : Value := UV_vJisa_41[DataRow-1];
      9 : Value := UV_vJisa_43[DataRow-1];
     10 : Value := UV_vJisa_45[DataRow-1];
     11 : Value := UV_vJisa_46[DataRow-1];
     12 : Value := UV_vJisa_47[DataRow-1];
     13 : Value := UV_vJisa_51[DataRow-1];
     14 : Value := UV_vJisa_52[DataRow-1];
     15 : Value := UV_vJisa_61[DataRow-1];
     16 : Value := UV_vJisa_63[DataRow-1];
     17 : Value := UV_vJisa_71[DataRow-1];
     18 : Value := UV_vJisa_72[DataRow-1];
     19 : Value := UV_vJisa_kita[DataRow-1];
     20 : Value := UV_vTotal[DataRow-1];
   end;
end;

procedure TfrmLD8Q03.btnQueryClick(Sender: TObject);
var i, index, iJisa_15, iJisa_11, iJisa_33, iJisa_34, iJisa_35, iJisa_41 : Integer;
    iJisa_43, iJisa_45, iJisa_46, iJisa_47, iJisa_51, iJisa_52, iJisa_61 : Integer;
    iJisa_63, iJisa_71, iJisa_72, iJisa_kita, iTotal, iStart, iEnd : Integer;
    sCod_hm, sDesc_kor, sGulkwa, gul  : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid  초기화
   grdMaster.Rows := 0;
   iJisa_15 := 0;    iJisa_11 := 0;
   iJisa_33 := 0;    iJisa_34 := 0;
   iJisa_35 := 0;    iJisa_41 := 0;
   iJisa_43 := 0;    iJisa_45 := 0;
   iJisa_46 := 0;    iJisa_47 := 0;
   iJisa_51 := 0;    iJisa_52 := 0;
   iJisa_61 := 0;    iJisa_63 := 0;
   iJisa_71 := 0;    iJisa_72 := 0;
   iJisa_kita := 0;  iTotal   := 0;
   index := 0;

   with  qryGulkwa do
   begin
      Active := False;
      ParamByName('St_Date').AsString := mskSt_Date.text;
      ParamByName('ED_Date').AsString := mskED_Date.text;
      ParamByName('Cod_bjjs').AsString := Copy(cmbJisa.Text, 1, 2);
      Active := true;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD8Q03', 'Q', 'Y');
         UP_VarMemSet(RecordCount-1);
         for i := 0 to RecordCount - 1 do
         begin
           iStart  := FieldByname('Pos_start').Asinteger;
           iEnd    := FieldByname('Pos_End').Asinteger;
           if  (sCod_hm <> FieldByName('Cod_hm').AsString) and (i <> 0) then
           begin
              UV_vCod_hm[index]     := sCod_hm;
              UV_vdesc_kor[index]   := sDesc_kor;
              UV_vJisa_15[index]    := IntToStr(iJisa_15);
              UV_vJisa_11[index]    := IntToStr(iJisa_11);
              UV_vJisa_33[index]    := IntToStr(iJisa_33);
              UV_vJisa_34[index]    := IntToStr(iJisa_34);
              UV_vJisa_35[index]    := IntToStr(iJisa_35);
              UV_vJisa_41[index]    := IntToStr(iJisa_45);
              UV_vJisa_43[index]    := IntToStr(iJisa_43);
              UV_vJisa_45[index]    := IntToStr(iJisa_45);
              UV_vJisa_46[index]    := IntToStr(iJisa_46);
              UV_vJisa_47[index]    := IntToStr(iJisa_47);
              UV_vJisa_51[index]    := IntToStr(iJisa_51);
              UV_vJisa_52[index]    := IntToStr(iJisa_52);
              UV_vJisa_61[index]    := IntToStr(iJisa_61);
              UV_vJisa_63[index]    := IntToStr(iJisa_63);
              UV_vJisa_71[index]    := IntToStr(iJisa_71);
              UV_vJisa_72[index]    := IntToStr(iJisa_72);
              UV_vJisa_kita[index]  := IntToStr(iJisa_kita);
              UV_vTotal[index]      := IntToStr(iTotal);
              iJisa_15 := 0;    iJisa_11 := 0;
              iJisa_33 := 0;    iJisa_34 := 0;
              iJisa_35 := 0;    iJisa_41 := 0;
              iJisa_43 := 0;    iJisa_45 := 0;
              iJisa_46 := 0;    iJisa_47 := 0;
              iJisa_51 := 0;    iJisa_52 := 0;
              iJisa_61 := 0;    iJisa_63 := 0;
              iJisa_71 := 0;    iJisa_72 := 0;
              iJisa_kita := 0;  iTotal   := 0;
              Inc(index);
           end;
           gul :=  FieldByName('Desc_glkwa1').AsString +
                   FieldByName('Desc_glkwa2').AsString +
                   FieldByName('Desc_glkwa3').AsString;
           sGulkwa := copy(gul, iStart,  iEnd - iStart + 1);
           if  copy(trim(sGulkwa),1,1) <> '' then
           begin
               if  FieldByName('Cod_jisa').AsString = '15' then
                   iJisa_15 := iJisa_15 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '11' then
                   iJisa_11 := iJisa_11 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '33' then
                   iJisa_33 := iJisa_33 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '34' then
                   iJisa_34 := iJisa_34 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '35' then
                   iJisa_35 := iJisa_35 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '41' then
                   iJisa_41 := iJisa_41 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '43' then
                   iJisa_43 := iJisa_43 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '45' then
                   iJisa_45 := iJisa_45 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '46' then
                   iJisa_46 := iJisa_46 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '47' then
                   iJisa_47 := iJisa_47 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '51' then
                   iJisa_51 := iJisa_51 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '52' then
                   iJisa_52 := iJisa_52 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '61' then
                   iJisa_61 := iJisa_61 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '63' then
                   iJisa_63 := iJisa_63 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '71' then
                   iJisa_71 := iJisa_71 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '72' then
                   iJisa_72 := iJisa_72 + 1
               else
                   iJisa_Kita := iJisa_Kita + 1 ;
               iTotal := iTotal + 1;
            end;
           if  (i = RecordCount - 1 )  then
           begin
              UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
              UV_vdesc_kor[index]   := FieldByName('desc_kor').AsString;
              UV_vJisa_15[index]    := IntToStr(iJisa_15);
              UV_vJisa_11[index]    := IntToStr(iJisa_11);
              UV_vJisa_33[index]    := IntToStr(iJisa_33);
              UV_vJisa_34[index]    := IntToStr(iJisa_34);
              UV_vJisa_35[index]    := IntToStr(iJisa_35);
              UV_vJisa_41[index]    := IntToStr(iJisa_45);
              UV_vJisa_43[index]    := IntToStr(iJisa_43);
              UV_vJisa_45[index]    := IntToStr(iJisa_45);
              UV_vJisa_46[index]    := IntToStr(iJisa_46);
              UV_vJisa_47[index]    := IntToStr(iJisa_47);
              UV_vJisa_51[index]    := IntToStr(iJisa_51);
              UV_vJisa_52[index]    := IntToStr(iJisa_52);
              UV_vJisa_61[index]    := IntToStr(iJisa_61);
              UV_vJisa_63[index]    := IntToStr(iJisa_63);
              UV_vJisa_71[index]    := IntToStr(iJisa_71);
              UV_vJisa_72[index]    := IntToStr(iJisa_72);
              UV_vJisa_kita[index]  := IntToStr(iJisa_kita);
              UV_vTotal[index]      := IntToStr(iTotal);
              Inc(index);
           end;

           sCod_hm :=  FieldByName('Cod_hm').AsString;
           sDesc_kor :=  FieldByName('Desc_kor').AsString;
           Next;
        end;
         // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
       Active := False;
    end
    else
    begin
      GF_MsgBox('Information', 'NODATA');
      exit;
    end;
 end;

 // Grid에 자료를 할당
 grdMaster.Rows := index;

 // Query Mode & Msg
 if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD8Q03.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 없을경우의 처리
   if NewRow = 0 then exit;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // Grid Focus
   grdMaster.SetFocus;
end;

procedure TfrmLD8Q03.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate  then GF_CalendarClick(mskST_Date)
   else
   if Sender = btnDate1 then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD8Q03.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskST_Date then UP_Click(btnDate)
   else if Sender = mskED_Date then UP_Click(btnDate1)
end;

procedure TfrmLD8Q03.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD8Q031 := TfrmLD8Q031.Create(Self);
   frmLD8Q031.QuickRep.Preview;
//   frmMD3Q311.QuickRep.Print;
end;

procedure TfrmLD8Q03.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

end.
