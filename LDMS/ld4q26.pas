//==============================================================================
// 프로그램 설명 : 채용 재검 명단 조회
// 시스템        : 통합검진
// 수정일자      : 2015.11.15
// 수정자        : 곽윤설
// 수정내용      :
// 참고사항      :
//==============================================================================
unit ld4q26;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan;

type
  TfrmLD4Q26 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    pnlgumgin: TPanel;
    Label1: TLabel;
    msk_SGmgn: TDateEdit;
    btnsDate: TSpeedButton;
    edtDESC_DC: TEdit;
    btnCOD_DC: TSpeedButton;
    edtCOD_DC: TEdit;
    Label4: TLabel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryGumgin: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure edtCOD_DCChange(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vdat_gmgn, UV_vNum_samp, UV_vDesc_name, UV_vNum_jumin : Variant;

    // SQL 임시변수
    UV_sSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q26: TfrmLD4Q26;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q261;

{$R *.DFM}

procedure TfrmLD4Q26.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading  := '검진일';
      Col[2].Heading  := '샘플번호';
      Col[3].Heading  := '이 름';
      Col[4].Heading  := '주민번호';
      // Column Alignment
      for i := 1 to 4 do
      begin
         if i = 4 then Col[i].Alignment := taLeftJustify;
         Col[i].Alignment := taCenter;
      end;
   end;
end;

procedure TfrmLD4Q26.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vDat_gmgn     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin    := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vNum_jumin, iLength);
   end;
end;

function TfrmLD4Q26.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (cmbJisa.ItemIndex = -1) or (msk_SGmgn.Text = '') or (msk_SGmgn.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q26.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;

   // SQL문 저장
   UV_sSQL := qryGumgin.SQL.Text;
end;

procedure TfrmLD4Q26.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := copy(UV_vDat_gmgn[DataRow-1],1,4) + '-' +
                   copy(UV_vDat_gmgn[DataRow-1],5,2) + '-' +
                   copy(UV_vDat_gmgn[DataRow-1],7,2);
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := UV_vNum_jumin[DataRow-1];
   end;
end;

procedure TfrmLD4Q26.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   sWhere := '';
   sOrderBy := '';

   with qryGumgin do
   begin
      // SQL문 생성
      Active := False;

      sWhere := sWhere + ' and cod_jisa = ''' + copy(cmbJisa.text,1,2) + ''' ';
      sWhere := sWhere + ' and dat_gmgn = ''' + msk_Sgmgn.text + ''' ';

      if edtCod_dc.text <> '' then
         sWhere := sWhere + ' and cod_dc = ''' + edtCod_dc.text + ''' ';

      sOrderBy := ' ORDER BY num_samp';
      SQL.Clear;
      SQL.Add(UV_sSQL + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD4Q26', 'Q', 'N');
         // 자료의 갯수만큼 Variant 변수의 Memory 재할당
         UP_VarMemSet(RecordCount-1);

         // DataSet의 값을 Variant변수로 이동
         index := -1;
         for i := 0 to RecordCount - 1 do
         begin
            Inc(index);
            UV_vDat_gmgn[index]  := FieldByName('dat_gmgn').AsString;
            UV_vNum_samp[index]  := FieldBYName('num_samp').AsString;
            UV_vDesc_name[index] := FieldByName('desc_name').AsString;
            UV_vNum_jumin[index] := copy(FieldByName('num_jumin').AsString,1,6);
            Next;
         end;  // for i := 0 to RecordCount - 1 do
         // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
         Active := False;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qryGumgin do

   // Grid에 자료를 할당
   grdMaster.Rows := index+1;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q26.edtCOD_DCChange(Sender: TObject);
var sName : String;
begin
   inherited;

   if Sender = edtCOD_DC then
   begin
      if GF_DancheEdit(edtCOD_DC, sName) then
         edtDESC_DC.Text := sName;
   end;
end;

procedure TfrmLD4Q26.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   if Sender = btnCOD_DC then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtCOD_DC.Text  := sCode;
         edtDESC_DC.Text := sName;
      end;
   end
   else
   if Sender = btnsDate then GF_CalendarClick(msk_SGmgn);
end;
procedure TfrmLD4Q26.btnPrintClick(Sender: TObject);
begin
  inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q261 := TfrmLD4Q261.Create(Self);
   frmLD4Q261.QuickRep.Preview;

end;

end.
