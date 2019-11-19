//==============================================================================
// 프로그램 설명 : 혈액형보기
// 시스템        : 통합검진
// 수정일자      : 2002.03.19
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.25
// 수정자        : 곽윤설
// 수정내용      : 샘플번호 추가 및 샘플번호 순으로 조회
// 참고사항      :
//==============================================================================
unit LD4Q03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan;

type
  TfrmLD4Q03 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Panel2: TPanel;
    rdBunju: TRadioButton;
    rdGumgin: TRadioButton;
    pnlbunju: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    msk_Bgmgn: TDateEdit;
    btnBdate: TSpeedButton;
    bunjuno1: TEdit;
    pnlgumgin: TPanel;
    Label1: TLabel;
    msk_SGmgn: TDateEdit;
    btnsDate: TSpeedButton;
    Label8: TLabel;
    msk_EGmgn: TDateEdit;
    btneDate: TSpeedButton;
    edtDESC_DC: TEdit;
    btnCOD_DC: TSpeedButton;
    edtCOD_DC: TEdit;
    Label4: TLabel;
    bunjuno2: TEdit;
    Label5: TLabel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qry_gulkwa: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure edtCOD_DCChange(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure rdBunjuClick(Sender: TObject);
    procedure rdGumginClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vDat_bunju, UV_vNum_bunju,  UV_vNum_samp, UV_vDesc_name, UV_vNum_jumin, UV_vABO,
    UV_vdat_gmgn,  UV_vCod_dc,    UV_vCod_jisa,  UV_vNum_id : Variant;

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
  frmLD4Q03: TfrmLD4Q03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD4Q03.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading  := '분주일';
      Col[2].Heading  := '분주번호';
      Col[3].Heading  := '샘플번호';
      Col[4].Heading  := '이 름';
      Col[5].Heading  := '주민번호';
      Col[6].Heading  := '혈액형';
      Col[7].Heading  := '검진일';
      Col[8].Heading  := '단체코드';
      Col[9].Heading  := '지 사';

      // Column Alignment
      for i := 1 to 9 do
      begin
         if i = 9 then Col[i].Alignment := taLeftJustify;
         Col[i].Alignment := taCenter;
      end;
   end;
end;

procedure TfrmLD4Q03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vdat_bunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin    := VarArrayCreate([0, 0], varOleStr);
      UV_vABO          := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_dc       := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id       := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vdat_bunju,    iLength);
      VarArrayRedim(UV_vNum_bunju,    iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);
      VarArrayRedim(UV_vDesc_name,    iLength);
      VarArrayRedim(UV_vNum_jumin,    iLength);
      VarArrayRedim(UV_vABO,          iLength);
      VarArrayRedim(UV_vDat_gmgn,     iLength);
      VarArrayRedim(UV_vCod_dc,       iLength);
      VarArrayRedim(UV_vCod_jisa,     iLength);
      VarArrayRedim(UV_vNum_id,       iLength);
   end;
end;

function TfrmLD4Q03.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (rdbunju.checked = true) and (msk_Bgmgn.text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
   if (rdgumgin.checked = true) and ((cmbJisa.ItemIndex = -1) or
      (msk_SGmgn.Text = '') or (msk_SGmgn.Text = '')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q03.FormCreate(Sender: TObject);
begin
   inherited;

   rdbunju.checked := True;
   pnlgumgin.enabled := False;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   cmbjisa.Items.Add('99-총지사');

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;

   // SQL문 저장
   UV_sSQL := qry_gulkwa.SQL.Text;
end;

procedure TfrmLD4Q03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := copy(UV_vdat_bunju[DataRow-1],1,4) + '-' +
                   copy(UV_vdat_bunju[DataRow-1],5,2) + '-' +
                   copy(UV_vdat_bunju[DataRow-1],7,2);
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vNum_samp[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      6 : Value := UV_vABO[DataRow-1];
      7 : Value := copy(UV_vDat_gmgn[DataRow-1],1,4) + '-' +
                   copy(UV_vDat_gmgn[DataRow-1],5,2) + '-' +
                   copy(UV_vDat_gmgn[DataRow-1],7,2);
      8 : Value := UV_vCod_dc[DataRow-1];
      9 : Value := UV_vCod_jisa[DataRow-1];
   end;
end;

procedure TfrmLD4Q03.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qry_gulkwa do
   begin
      if rdbunju.checked = true then
      begin
         // SQL문 생성
         Active := False;
         sWhere := ' and G.dat_bunju = ''' + msk_Bgmgn.text + '''';
         if (bunjuno1.text <> '') and (bunjuno2.text <> '') then
         begin
            if (length(bunjuno1.text) = 4) and (length(bunjuno2.text) = 4) then
            begin
               sWhere := sWhere + ' and G.num_bunju >= ''' + bunjuno1.text + '''';
               sWhere := sWhere + ' and G.num_bunju <= ''' + bunjuno2.text + '''';
            end;
         end;
         sOrderBy := ' ORDER BY num_samp';
         SQL.Clear;
         SQL.Add(UV_sSQL + sWhere + sOrderBy);
         Active := True
      end
      else
      if rdgumgin.checked = true then
      begin
         // SQL문 생성
         Active := False;
         if (GV_sJisa = '00') or (GV_sJisa = '15') or (GV_sJisa = '12') then
            if copy(cmbJisa.Text,1,2) <> '99' then
               sWhere := ' and M.cod_jisa = ''' + copy(cmbJisa.text,1,2) + ''''
            else
         else
            sWhere := ' and M.cod_jisa = ''' + GV_sJisa + '''';
         sWhere := sWhere + ' and M.dat_gmgn >= ''' + msk_Sgmgn.text + '''';
         sWhere := sWhere + ' and M.dat_gmgn <= ''' + msk_Egmgn.text + '''';

         if edtCod_dc.text <> '' then
            sWhere := sWhere + ' and M.cod_dc = ''' + edtCod_dc.text + '''';

         sOrderBy := ' ORDER BY num_samp';
         SQL.Clear;
         SQL.Add(UV_sSQL + sWhere + sOrderBy);
         Active := True;
      end;

      if RecordCount > 0 then
      begin
         GP_query_log(qry_gulkwa, 'LD4Q03', 'Q', 'N');
         // 자료의 갯수만큼 Variant 변수의 Memory 재할당
         UP_VarMemSet(RecordCount-1);

         // DataSet의 값을 Variant변수로 이동
         index := -1;
         for i := 0 to RecordCount - 1 do
         begin
            Inc(index);
            UV_vDat_bunju[index]  := FieldByName('dat_bunju').AsString;
            UV_vNum_bunju[index]  := FieldBYName('num_bunju').AsString;
            UV_vNum_samp[index]   := FieldBYName('num_samp').AsString;
            UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
            UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
            UV_vABO[index]        := FieldByName('ABO').AsString;
            UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsString;
            UV_vCod_dc[index]     := FieldByName('cod_dc').AsString;
            UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
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
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index+1;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q03.edtCOD_DCChange(Sender: TObject);
var sName : String;
begin
   inherited;

   if Sender = edtCOD_DC then
   begin
      if GF_DancheEdit(edtCOD_DC, sName) then
         edtDESC_DC.Text := sName;
   end;
end;

procedure TfrmLD4Q03.UP_Click(Sender: TObject);
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
   if Sender = btnsDate then GF_CalendarClick(msk_SGmgn)
   else
   if Sender = btneDate then GF_CalendarClick(msk_EGmgn);
end;
procedure TfrmLD4Q03.rdBunjuClick(Sender: TObject);
begin
  inherited;
   pnlgumgin.enabled := False;
   pnlbunju.enabled := True;
end;

procedure TfrmLD4Q03.rdGumginClick(Sender: TObject);
begin
  inherited;
   pnlbunju.enabled := False;
   pnlgumgin.Enabled := True;
end;

end.
