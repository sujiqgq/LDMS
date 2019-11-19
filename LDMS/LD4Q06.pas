//==============================================================================
// ���α׷� ���� : URIN������Ȳ
// �ý���        : ���հ���
// ��������      : 2003.10.02
// ������        : ������
// ��������      : 
// �������      :
//==============================================================================
// ��������      : 2015.03.13
// ������        : ������
// ��������      : Y005�˻� ���� -> U017, U046 �˻�� ����
// �������      : ����-�ѵη�
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
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vURIN, UV_vNAME : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
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
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vURIN  := VarArrayCreate([0, 0], varOleStr);
      UV_vNAME  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vURIN,  iLength);
      VarArrayRedim(UV_vNAME,  iLength);

   end;
end;

function TfrmLD4Q06.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
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
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa,  copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q06.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
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
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
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
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
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
