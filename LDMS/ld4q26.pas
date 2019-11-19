//==============================================================================
// ���α׷� ���� : ä�� ��� ��� ��ȸ
// �ý���        : ���հ���
// ��������      : 2015.11.15
// ������        : ������
// ��������      :
// �������      :
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
    // Grid�� ������ Variant ���� ����
    UV_vdat_gmgn, UV_vNum_samp, UV_vDesc_name, UV_vNum_jumin : Variant;

    // SQL �ӽú���
    UV_sSQL : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
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
   // Grid�� ȯ�� ����
   with grdMaster do
   begin
      // Row���� �ʱ�ȭ
      Rows := 0;

      // Column Title
      Col[1].Heading  := '������';
      Col[2].Heading  := '���ù�ȣ';
      Col[3].Heading  := '�� ��';
      Col[4].Heading  := '�ֹι�ȣ';
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
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vDat_gmgn     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin    := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vNum_jumin, iLength);
   end;
end;

function TfrmLD4Q26.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if (cmbJisa.ItemIndex = -1) or (msk_SGmgn.Text = '') or (msk_SGmgn.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q26.FormCreate(Sender: TObject);
begin
   inherited;
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_GridSet;

   // SQL�� ����
   UV_sSQL := qryGumgin.SQL.Text;
end;

procedure TfrmLD4Q26.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
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

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   sWhere := '';
   sOrderBy := '';

   with qryGumgin do
   begin
      // SQL�� ����
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
         // �ڷ��� ������ŭ Variant ������ Memory ���Ҵ�
         UP_VarMemSet(RecordCount-1);

         // DataSet�� ���� Variant������ �̵�
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
         // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
         Active := False;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qryGumgin do

   // Grid�� �ڷḦ �Ҵ�
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
