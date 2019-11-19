//==============================================================================
// ���α׷� ���� : H.pylori������Ȳ
// �ý���        : ���հ���
// ��������      : 2005.12.24
// ������        : ������
// ��������      : Ư���׸� üũ...
// �������      :
//==============================================================================

unit LD8Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD8Q01 = class(TfrmSingle)
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
    cmbGubn: TComboBox;
    Label1: TLabel;
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
  frmLD8Q01: TfrmLD8Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q011;

{$R *.DFM}

procedure TfrmLD8Q01.UP_VarMemSet(iLength : Integer);
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

function TfrmLD8Q01.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD8Q01.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   cmbGubn.itemindex := 0;
end;

procedure TfrmLD8Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRIA[DataRow-1];
   end;
end;

procedure TfrmLD8Q01.btnQueryClick(Sender: TObject);
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
      if (cmbGubn.itemindex = 0) or (cmbGubn.itemindex = 1) or (cmbGubn.itemindex = 2) or
         (cmbGubn.itemindex = 7) or (cmbGubn.itemindex = 8) or (cmbGubn.itemindex = 9) or
         (cmbGubn.itemindex = 10)then
      begin
         sSelect := 'select distinct Num_Bunju From Gulkwa a ';
         sSelect := sSelect + ' left outer join gumgin b ';
         sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
         sSelect := sSelect + ' and a.num_id = b.num_id ';
         sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
         sSelect := sSelect + ' left outer join profile_hm c ';
         sSelect := sSelect + ' on (b.cod_hul = c.cod_pf ';
         sSelect := sSelect + ' or b.cod_jangbi = c.cod_pf ';
         sSelect := sSelect + ' or b.cod_cancer = c.cod_pf) ';
         sSelect := sSelect + ' left outer join Tkgum T ';
         sSelect := sSelect + ' on a.cod_jisa = T.cod_jisa and b.num_jumin = T.num_jumin and a.dat_gmgn = T.dat_gmgn ';

         sWhere := ' where a.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
         sWhere := sWhere + ' And  a.Dat_Bunju = ''' + edtBDate.Text + '''';
         If BunjuNo1.Text <> '' Then
            sWhere := sWhere + ' And a.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                               ' And a.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      end else
      if (cmbGubn.itemindex = 3) or (cmbGubn.itemindex = 4) or
         (cmbGubn.itemindex = 5) or (cmbGubn.itemindex = 6) then
      begin
         sSelect := 'select distinct c.num_bunju from gumgin a ';
         sSelect := sSelect + ' left outer join tkgum b ';
         sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
         sSelect := sSelect + ' and a.num_jumin = b.num_jumin ';
         sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
         sSelect := sSelect + ' left outer join gulkwa c ';
         sSelect := sSelect + ' on a.cod_jisa = c.cod_jisa ';
         sSelect := sSelect + ' and a.num_id = c.num_id ';
         sSelect := sSelect + ' and a.dat_gmgn = c.dat_gmgn ';
         sSelect := sSelect + ' left outer join profile_hm d ';
         sSelect := sSelect + ' on  (substring(b.cod_prf,1,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,5,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,9,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,13,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,17,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,21,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,25,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,29,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,33,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,37,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,41,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,45,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,49,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,53,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,57,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,61,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,65,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,69,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,73,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,77,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,81,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,85,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,89,4) = d.cod_pf ';
         sSelect := sSelect + ' or substring(b.cod_prf,93,4) = d.cod_pf) ';
         sWhere :=  ' where c.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
         sWhere  := sWhere + ' And c.dat_bunju = ''' + edtBDate.Text + '''';
         If BunjuNo1.Text <> '' Then
            sWhere := sWhere + ' And c.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                               ' And c.Num_Bunju <= ''' + BunjuNo2.Text + '''';

      end;

      if cmbGubn.itemindex = 0 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '05' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'S021' + '%'+'''';
         sWhere := sWhere + ' or T.cod_chuga like ''' +'%' + 'S021' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'S021' + ''' )';
      end else
      if cmbGubn.itemindex = 1 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '08' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'E015' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'E015' + ''' )';
      end else
      if cmbGubn.itemindex = 2 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '03' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'H025' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'H025' + ''' )';
      end else
      if cmbGubn.itemindex = 7 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '05' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'T009' + '%'+'''';
         sWhere := sWhere + ' or T.cod_chuga like ''' +'%' + 'T009' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'T009' + ''' )';
      end else
      if cmbGubn.itemindex = 8 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '05' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'S019' + '%'+'''';
         sWhere := sWhere + ' or T.cod_chuga like ''' +'%' + 'S019' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'S019' + ''' )';
      end else
      if cmbGubn.itemindex = 9 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '04' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'C044' + '%'+'''';
         sWhere := sWhere + ' or T.cod_chuga like ''' +'%' + 'C044' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'C044' + ''' )';
      end else
      if cmbGubn.itemindex = 10 then
      begin
         sWhere := sWhere + ' And  a.Gubn_Part = ''' + '08' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'E041' + '%'+'''';
         sWhere := sWhere + ' or T.cod_chuga like ''' +'%' + 'E041' + '%'+'''';
         sWhere := sWhere + ' or c.cod_hm = ''' + 'E041' + ''' )';
      end;

      if cmbGubn.itemindex = 3 then
      begin
         sWhere := sWhere + ' And  c.Gubn_Part = ''' + '09' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'SE24' + '%'+'''';
         sWhere := sWhere + ' or d.cod_hm = ''' + 'SE24' + '''';
         sWhere := sWhere + ' or a.cod_chuga like ''' + '%' + 'SE24' + '%' + ''')';
      end;
      if cmbGubn.itemindex = 4 then
      begin
         sWhere := sWhere + ' And  c.Gubn_Part = ''' + '09' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'SE50' + '%'+'''';
         sWhere := sWhere + ' or d.cod_hm = ''' + 'SE50' + '''';
         sWhere := sWhere + ' or a.cod_chuga like ''' + '%' + 'SE50' + '%' + ''')';
      end;
      if cmbGubn.itemindex = 5 then
      begin
         sWhere := sWhere + ' And  c.Gubn_Part = ''' + '09' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'SE53' + '%'+'''';
         sWhere := sWhere + ' or d.cod_hm = ''' + 'SE53' + '''';
         sWhere := sWhere + ' or a.cod_chuga like ''' + '%' + 'SE53' + '%' + ''')';
      end;
      if cmbGubn.itemindex = 6 then
      begin
         sWhere := sWhere + ' And  c.Gubn_Part = ''' + '09' + '''';
         sWhere := sWhere + ' and (b.cod_chuga like ''' +'%' + 'SE57' + '%'+'''';
         sWhere := sWhere + ' or d.cod_hm = ''' + 'SE57' + '''';
         sWhere := sWhere + ' or a.cod_chuga like ''' + '%' + 'SE57' + '%' + ''')';
      end;

      if (cmbGubn.itemindex = 0) or (cmbGubn.itemindex = 1) or (cmbGubn.itemindex = 2) or
         (cmbGubn.itemindex = 7) or (cmbGubn.itemindex = 8) or (cmbGubn.itemindex = 9) or
         (cmbGubn.itemindex = 10) then
         sOrderBy := ' ORDER BY a.Num_Bunju'
      else if (cmbGubn.itemindex = 3) or (cmbGubn.itemindex = 4) or
              (cmbGubn.itemindex = 5) or (cmbGubn.itemindex = 6) then
         sOrderBy := ' ORDER BY c.Num_Bunju';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD8Q01', 'Q', 'N');
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

procedure TfrmLD8Q01.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD8Q01.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD8Q011 := TfrmLD8Q011.Create(Self);
     frmLD8Q011.QuickRep.Preview;
  end
  else
  begin
     frmLD8Q011 := TfrmLD8Q011.Create(Self);
     frmLD8Q011.QuickRep.Print;
  end;
end;

procedure TfrmLD8Q01.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

end.
