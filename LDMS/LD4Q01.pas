unit LD4Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, Gauges;

type
  TfrmLD4Q01 = class(TfrmSingle)
    pnlCond: TPanel;
    mskST_date: TDateEdit;
    btnSt_date: TSpeedButton;
    mskEd_date: TDateEdit;
    btnEd_date: TSpeedButton;
    Label10: TLabel;
    Label11: TLabel;
    grdMaster: TtsGrid;
    qryBunju: TQuery;
    qryGumgin: TQuery;
    qryGulkwa: TQuery;
    SPrct: TGauge;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
 private
     { Private declarations }
    UV_vDat_gmgn, UV_vDat_Bunju, UV_vNum_bunju, UV_vDesc_Name, UV_vNum_Jumin, UV_vCod_jisa,
    UV_vH023, UV_vH024, UV_vS007, UV_vS008, UV_vS034, UV_vTT01, UV_vS016 : Variant;

    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
  public
    { Public declarations }
  end;


var
  frmLD4Q01 : TfrmLD4Q01;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}

procedure TfrmLD4Q01.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      Col[1].Heading := '분주일자';
      Col[2].Heading := '분주번호';
      Col[3].Heading := '지사';
      Col[4].Heading := '성명 ';
      Col[5].Heading := '검진일자';
      Col[6].Heading := '주민번호';
      Col[7].Heading := 'ABO형';
      Col[8].Heading := 'Rh형';
      Col[9].Heading := 'HBs Ag(RPHA)';
      Col[10].Heading := 'HBs Ab(PHA)';
      Col[11].Heading := 'VDRL(정성)';
      Col[12].Heading := 'AFP(RPHA)';
      Col[13].Heading := 'HCV Ab';

      // Set Alignment
      Col[1].Alignment := taLeftJustify;
      Col[2].Alignment := taLeftJustify;
      Col[3].Alignment := taLeftJustify;
      Col[4].Alignment := taLeftJustify;
      Col[5].Alignment := taLEftJustify;
      Col[6].Alignment := taLEftJustify;
      Col[7].Alignment := taLeftJustify;
      Col[8].Alignment := taLeftJustify;
      Col[9].Alignment := taLeftJustify;
      Col[10].Alignment := taLeftJustify;
      Col[11].Alignment := taLeftJustify;
      Col[12].Alignment := taLeftJustify;
      Col[13].Alignment := taLeftJustify;

//      Col[1].SortPicture := spDown;
//      Col[2].SortPicture := spDown;
//      Col[3].SortPicture := spDown;


   end;
end;

procedure TfrmLD4Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_jisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Name := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Jumin := VarArrayCreate([0, 0], varOleStr);
      UV_vH023      := VarArrayCreate([0, 0], varOleStr);
      UV_vH024      := VarArrayCreate([0, 0], varOleStr);
      UV_vS007      := VarArrayCreate([0, 0], varOleStr);
      UV_vS008      := VarArrayCreate([0, 0], varOleStr);
      UV_vS034      := VarArrayCreate([0, 0], varOleStr);
      UV_vTT01      := VarArrayCreate([0, 0], varOleStr);
      UV_vS016      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vDat_Bunju,  iLength);
      VarArrayRedim(UV_vNum_Bunju,  iLength);
      VarArrayRedim(UV_vDesc_Name,  iLength);
      VarArrayRedim(UV_vNum_Jumin,  iLength);
      VarArrayRedim(UV_vH023,       iLength);
      VarArrayRedim(UV_vH024,       iLength);
      VarArrayRedim(UV_vS007,       iLength);
      VarArrayRedim(UV_vS008,       iLength);
      VarArrayRedim(UV_vS034,       iLength);
      VarArrayRedim(UV_vTT01,       iLength);
      VarArrayRedim(UV_vS016,       iLength);
   end;
end;
procedure TfrmLD4Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := copy(UV_vDat_Bunju[DataRow-1],1,4) + '/' + copy(UV_vDat_Bunju[DataRow-1],5,2) + '/' + copy(UV_vDat_Bunju[DataRow-1],7,2);
      2 : Value := UV_vNum_Bunju[DataRow-1];
      3 : Value := UV_vCod_jisa[DataRow-1];
      4 : Value := UV_vDesc_Name[DataRow-1];
      5 : Value := copy(UV_vDat_Gmgn[DataRow-1],1,4) + '/' + copy(UV_vDat_Gmgn[DataRow-1],5,2) + '/' + copy(UV_vDat_Gmgn[DataRow-1],7,2);
      6 : Value := copy(UV_vNum_Jumin[DataRow-1],1,6) + '-' + copy(UV_vNum_Jumin[DataRow-1],7,7);
      7 : Value := UV_vH023[DataRow-1];
      8 : Value := UV_vH024[DataRow-1];
      9 : Value := UV_vS007[DataRow-1];
      10 : Value := UV_vS008[DataRow-1];
      11 : Value := UV_vS034[DataRow-1];
      12 : Value := UV_vTT01[DataRow-1];
      13 : Value := UV_vS016[DataRow-1];
   end;
end;

procedure TfrmLD4Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid 설정 & Memory Allocation
   UP_GridSet;
   UP_VarMemSet(0);
   // Grid 초기화
   grdMaster.Rows := 0;

end;

procedure TfrmLD4Q01.btnQueryClick(Sender: TObject);
var
    I, J, index, Sw : Integer;
    H023, H024, S007, S008, S034, TT01, S016, sHangmok : String;
begin
   inherited;
   index := 0;
   sHangmok := 'H023H024S007S008S034TT01S016';
   with qryBunju do
   begin
      Close;
      ParamByname('In_St_Date').AsString := mskSt_Date.Text;
      ParamByname('In_Ed_Date').AsString := mskEd_Date.Text;
      Open;
      sPrct.MaxValue := RecordCount;
      UP_VarMemSet(0);
      For I := 1 To RecordCount Do
         Begin
            GP_query_log(qryBunju, 'LD4Q01', 'Q', 'Y');
            H023 := '';
            H024 := '';
            S007 := '';
            S008 := '';
            S034 := '';
            TT01 := '';
            S016 := '';
            Sw := 0;
            With QryGulkwa Do
               Begin
//혈액학 결과
                  Close;
                  ParamByName('In_Num_id').AsString    := qryBunju.FieldByName('Num_id').ASString;
                  ParamByName('In_Gubn_Part').AsString := '02';
                  Open;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If H023 = '' Then
                              H023 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,133,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,133,6)) <> '' Then
                                    If H023 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,133,6)) Then
                                       Begin
                                          H023 := H023 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,133,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',H023) = 0 Then
                     H023 := '';
                  First;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If H024 = '' Then
                              H024 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,139,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,139,6)) <> '' Then
                                    If H024 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,139,6)) Then
                                       Begin
                                          H024 := H024 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,139,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',H024) = 0 Then
                     H024 := '';
//혈청학 결과
                  Close;
                  ParamByName('In_Num_id').AsString    := qryBunju.FieldByName('Num_id').ASString;
                  ParamByName('In_Gubn_Part').AsString := '07';
                  Open;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If S007 = '' Then
                              S007 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,31,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,31,6)) <> '' Then
                                    If S007 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,31,6)) Then
                                       Begin
                                          S007 := S007 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,31,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',S007) = 0 Then
                     S007 := '';
                  First;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If S008 = '' Then
                              S008 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,37,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,37,6)) <> '' Then
                                    If S008 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,37,6)) Then
                                       Begin
                                          S008 := S008 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,37,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',S008) = 0 Then
                     S008 := '';
                  First;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If S034 = '' Then
                              S034 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,43,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,43,6)) <> '' Then
                                    If S034 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,43,6)) Then
                                       Begin
                                          S034 := S034 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,43,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',S034) = 0 Then
                     S034 := '';
                  First;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If TT01 = '' Then
                              TT01 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,85,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,85,6)) <> '' Then
                                    If TT01 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,85,6)) Then
                                       Begin
                                          TT01 := TT01 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,85,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',TT01) = 0 Then
                     TT01 := '';
//EIA 결과
                  Close;
                  ParamByName('In_Num_id').AsString    := qryBunju.FieldByName('Num_id').ASString;
                  ParamByName('In_Gubn_Part').AsString := '05';
                  Open;
                  If RecordCount > 1 Then
                     For J := 1 to RecordCount Do
                        Begin
                           If S016 = '' Then
                              S016 := Trim(Copy(FieldByname('Desc_Glkwa1').AsString,7,6))
                           Else If Trim(Copy(FieldByname('Desc_Glkwa1').AsString,7,6)) <> '' Then
                                    If S016 <> Trim(Copy(FieldByname('Desc_Glkwa1').AsString,7,6)) Then
                                       Begin
                                          S016 := S016 + '(' + Trim(Copy(FieldByname('Desc_Glkwa1').AsString,7,6)) + ')';
                                          Sw := 99;
                                          Break;
                                       End;
                           Next;
                        End;
                  If Pos('(',S016) = 0 Then
                     S016 := '';
               End;
            If Sw = 99 Then
               Begin
                  With qryGumgin Do
                     Begin
                        Close;
                        ParamByName('In_Cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
                        ParamByName('In_Num_id').AsString    := qryBunju.FieldByName('Num_id').AsString;
                        Open;
                        UV_vDesc_Name[index]   := FieldByname('Desc_Name').AsString;
                        UV_vDat_Gmgn[index]    := FieldByname('Dat_Gmgn').AsString;
                        UV_vNum_Jumin[index]   := FieldByname('Num_Jumin').AsString;
                        Close;
                     End;
                  UV_vDat_Bunju[index]   := FieldByname('Dat_Bunju').AsString;
                  UV_vNum_Bunju[index]   := FieldByname('Num_Bunju').AsString;
                  UV_vH023[Index] := H023;
                  UV_vH024[Index] := H024;
                  UV_vS007[Index] := S007;
                  UV_vS008[Index] := S008;
                  UV_vS034[Index] := S034;
                  UV_vTT01[Index] := TT01;
                  UV_vS016[Index] := S016;
                  Index := Index + 1;
                  UP_VarMemSet(Index);
               End;
            Sprct.Progress := I;
            Next;
         End;
      Close;
   End;
   grdMaster.Rows := index;
   If Index = 0 Then
      ShowMessage('☞ 이상 자료가 없읍니다......');
   grdMaster.SetFocus;
end;

procedure TfrmLD4Q01.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnST_Date then GF_CalendarClick(mskST_Date);
   if Sender = btnED_Date then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD4Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if      Sender = mskST_Date       then UP_Click(mskST_Date)
   else if  Sender = mskED_Date       then UP_Click(mskED_Date);

end;

end.
