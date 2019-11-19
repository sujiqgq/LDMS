//==============================================================================
// 프로그램 설명 : [지사]노신분주현황
// 시스템        : 통합검진
// 수정일자      : 2005.06.09
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD8Q04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD8Q04 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vRIA : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD8Q04: TfrmLD8Q04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q041;

{$R *.DFM}

procedure TfrmLD8Q04.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vRIA   := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vRIA,   iLength);
   end;
end;

function TfrmLD8Q04.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD8Q04.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD8Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRIA[DataRow-1];
   end;
end;

procedure TfrmLD8Q04.btnQueryClick(Sender: TObject);
var index, I : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := 'select b.Num_Bunju From gumgin a left outer join bunju b ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sWhere  := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';
      sWhere  := sWhere + ' And  a.Gubn_hulgum = ''3'' ';
      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And b.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And b.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      sWhere := sWhere + ' And a.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sOrderBy := ' ORDER BY Num_Bunju';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD8Q04', 'Q', 'N');
         For I := 1 to RecordCount do
            Begin
               UP_VarMemSet(Index);
               UV_VRIA[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               UV_vNo[Index]  := IntToStr(Index+1);
               Inc(Index);
               Gride.Progress := I;
               Next;
            End;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD8Q04.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD8Q04.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD8Q041 := TfrmLD8Q041.Create(Self);
     frmLD8Q041.QuickRep.Preview;
  end
  else
  begin
     frmLD8Q041 := TfrmLD8Q041.Create(Self);
     frmLD8Q041.QuickRep.Print;
  end;
end;

procedure TfrmLD8Q04.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

end.
