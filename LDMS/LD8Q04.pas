//==============================================================================
// ���α׷� ���� : [����]��ź�����Ȳ
// �ý���        : ���հ���
// ��������      : 2005.06.09
// ������        : ������
// ��������      : 
// �������      :
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
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vRIA : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
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
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vRIA   := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vRIA,   iLength);
   end;
end;

function TfrmLD8Q04.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
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
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD8Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
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
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
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
