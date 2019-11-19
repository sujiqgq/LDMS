//==============================================================================
// 프로그램 설명 : 항목별 지사별 검진인원 현황
// 시스템        : 통합검진
// 수정일자      : 1999.11.3
// 수정자        : 허정남
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD6Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD6Q01 = class(TfrmSingle)
    qrygumgin: TQuery;
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
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);    
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_vDat_bunju, UV_vNum_bunju, UV_vDesc_name, UV_vNum_jumin, UV_vCod_jisa, UV_vDat_gmgn,
    UV_vnum_id : Variant;
  end;

var
  frmLD6Q01: TfrmLD6Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q031;

{$R *.DFM}

procedure TfrmLD6Q01.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];

   //--------------------------------------------------------------------------
   // 내림차순
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vdat_bunju[Lo];  UV_vDat_bunju[Lo] := UV_vDat_bunju[Hi]; UV_vDat_bunju[Hi] := T;
            T := UV_vNum_bunju[Lo];  UV_vNum_bunju[Lo] := UV_vNum_bunju[Hi]; UV_vNum_bunju[Hi] := T;
            T := UV_vDesc_name[Lo];  UV_vDesc_name[Lo] := UV_vDesc_name[Hi]; UV_vDesc_name[Hi] := T;
            T := UV_vNum_Jumin[Lo];  UV_vNum_jumin[Lo] := UV_vNum_jumin[Hi]; UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];   UV_vDat_gmgn[Lo]  := UV_vDat_gmgn[Hi];  UV_vDat_gmgn[Hi]  := T;
            T := UV_vcod_jisa[Lo];   UV_vCod_jisa[Lo]  := UV_vCod_jisa[Hi];  UV_vCod_jisa[Hi]  := T;
            T := UV_vNum_id[Lo];     UV_vNum_id[Lo]    := UV_vNum_id[Hi];    UV_vNum_id[Hi]    := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // 오름차순
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vdat_bunju[Lo];  UV_vDat_bunju[Lo] := UV_vDat_bunju[Hi]; UV_vDat_bunju[Hi] := T;
            T := UV_vNum_bunju[Lo];  UV_vNum_bunju[Lo] := UV_vNum_bunju[Hi]; UV_vNum_bunju[Hi] := T;
            T := UV_vDesc_name[Lo];  UV_vDesc_name[Lo] := UV_vDesc_name[Hi]; UV_vDesc_name[Hi] := T;
            T := UV_vNum_Jumin[Lo];  UV_vNum_jumin[Lo] := UV_vNum_jumin[Hi]; UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];   UV_vDat_gmgn[Lo]  := UV_vDat_gmgn[Hi];  UV_vDat_gmgn[Hi]  := T;
            T := UV_vcod_jisa[Lo];   UV_vCod_jisa[Lo]  := UV_vCod_jisa[Hi];  UV_vCod_jisa[Hi]  := T;
            T := UV_vNum_id[Lo];     UV_vNum_id[Lo]    := UV_vNum_id[Hi];    UV_vNum_id[Hi]    := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD6Q01.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading := '분주일';
      Col[2].Heading := '분주번호';
      Col[3].Heading := '성  명';
      Col[4].Heading := '주민번호';
      Col[5].Heading := '검진일';
      Col[6].Heading := '지사';
      col[7].Heading := '회원번호' ;
      // Column Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
       for i := 1 to Cols do
       Col[i].SortPicture := spDown;

   end;
end;

procedure TfrmLD6Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성

      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vcod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vNum_jumin, iLength);
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vcod_jisa,  iLength);
      VarArrayRedim(UV_vNum_id,    iLength);
   end;
end;

function TfrmLD6Q01.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskST_Date.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD6Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;
   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);
   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := '15'
   else                    UV_sCod_jisa := GV_sJisa;

end;

procedure TfrmLD6Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vdat_bunju[DataRow-1];
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := UV_vNum_jumin[DataRow-1];
      5 : Value := UV_vDat_gmgn[DataRow-1];
      6 : Value := UV_vCod_jisa[DataRow-1];
      7 : Value := UV_vNum_id[DataRow-1];

   end;
end;

procedure TfrmLD6Q01.btnQueryClick(Sender: TObject);
var i, index : integer;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid  초기화
   grdMaster.Rows := 0;
   index := 0;

   with  qryGumgin do
   begin
      Active := False;
      ParamByName('St_Date').AsString := mskSt_Date.text;
      ParamByName('ED_Date').AsString := mskED_Date.text;
      ParamByName('Cod_jisa').AsString := Copy(cmbJisa.Text, 1, 2);
      Active := true;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD6Q01', 'Q', 'Y');
         for i := 0 to RecordCount - 1 do
         begin
            UP_VarMemSet(RecordCount-1);
            UV_vDat_bunju[index]   := FieldByName('Dat_bunju').AsString;
            UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
            UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
            UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
            UV_vDat_gmgn[index]    := FieldByName('dat_gmgn').AsString;
            UV_vCod_jisa[index]    := FieldByName('Cod_jisa').AsString;
            UV_vNum_id[index]      := FieldByName('Num_id').AsString;
            Inc(index);
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

procedure TfrmLD6Q01.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD6Q01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate  then GF_CalendarClick(mskST_Date)
   else
   if Sender = btnDate1 then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD6Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskST_Date then UP_Click(btnDate)
   else if Sender = mskED_Date then UP_Click(btnDate1)
end;

procedure TfrmLD6Q01.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

//   frmMD6Q311 := TfrmMD6Q311.Create(Self);
//   frmMD6Q311.QuickRep.Preview;
//   frmMD3Q311.QuickRep.Print;
end;

procedure TfrmLD6Q01.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

procedure TfrmLD6Q01.grdMasterHeadingClick(Sender: TObject;
  DataCol: Integer);
  var  iCnt : integer;
  sOrder : String;
begin
  inherited;
  // 자료가 존재하는지 Check
   if grdMaster.Rows = 0 then exit;

   iCnt := grdMaster.Rows;

   if grdMaster.Col[DataCol].SortPicture = spUp then
   begin
      grdMaster.Col[DataCol].SortPicture := spDown;
      sOrder := 'A';
   end
   else
   begin
      grdMaster.Col[DataCol].SortPicture := spUp;
      sOrder := 'D';
   end;

   case DataCol of
      1 : UP_QuickSort(sOrder, UV_vDat_Bunju,    0, iCnt-1);
      2 : UP_QuickSort(sOrder, UV_vNum_bunju,    0, iCnt-1);
      3 : UP_QuickSort(sOrder, UV_vDesc_name,    0, iCnt-1);
      4 : UP_QuickSort(sOrder, UV_vNum_jumin,      0, iCnt-1);
      5 : UP_QuickSort(sOrder, UV_vDat_gmgn,   0, iCnt-1);
      6 : UP_QuickSort(sOrder, UV_vCod_jisa,    0, iCnt-1);
      7 : UP_QuickSort(sOrder, UV_vNum_id,   0, iCnt-1);
      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;

end;

end.
