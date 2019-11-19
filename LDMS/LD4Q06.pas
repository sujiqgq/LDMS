//==============================================================================
// 프로그램 설명 : URIN분주현황
// 시스템        : 통합검진
// 수정일자      : 2003.10.02
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.03.13
// 수정자        : 곽윤설
// 수정내용      : Y005검사 안함 -> U017, U046 검사로 변경
// 참고사항      : 본원-한두례
//==============================================================================
unit LD4Q06;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q06 = class(TfrmSingle)
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
    Gride: TGauge;
    RB_Y004: TRadioButton;
    RB_U017U046: TRadioButton;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
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
    UV_vNo, UV_vURIN, UV_vNAME : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q06: TfrmLD4Q06;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q061;

{$R *.DFM}

procedure TfrmLD4Q06.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vURIN  := VarArrayCreate([0, 0], varOleStr);
      UV_vNAME  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vURIN,  iLength);
      VarArrayRedim(UV_vNAME,  iLength);

   end;
end;

function TfrmLD4Q06.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q06.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa,  copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q06.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vURIN[DataRow-1];
      3 : Value := UV_vNAME[DataRow-1];
   end;
end;

procedure TfrmLD4Q06.btnQueryClick(Sender: TObject);
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

      sSelect := 'SELECT A.Num_Bunju, B.desc_name FROM Gulkwa A WITH(NOLOCK) left outer join gumgin B WITH(NOLOCK) ' +
                 ' on A.cod_jisa = B.cod_jisa and A.num_id = B.num_id and A.dat_gmgn = B.dat_gmgn                  ';

      sWhere := 'WHERE A.Dat_Bunju = ''' + edtBDate.Text + ''' ' +
                ' And  A.Gubn_Part = ''' + '03' + ''' ';
      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And a.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And a.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      sWhere := sWhere + ' And a.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      if RB_Y004.Checked then
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'Y004' + '%'+''') '
      else sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'U017' + '%' +''' or ' +
                              '      b.cod_chuga like ''' +'%' + 'U046' + '%' +''' ) ';
      sOrderBy := ' ORDER BY a.Num_Bunju';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q06', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            UP_VarMemSet(Index);
            UV_vURIN[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
            UV_vName[Index] := qryBunju.FieldByName('Desc_Name').AsString;
            UV_vNo[Index]  := IntToStr(Index+1);
            Inc(Index);
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

procedure TfrmLD4Q06.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q06.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q061 := TfrmLD4Q061.Create(Self);
     frmLD4Q061.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q061 := TfrmLD4Q061.Create(Self);
     frmLD4Q061.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q06.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

end.
