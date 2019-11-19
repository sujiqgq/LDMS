//==============================================================================
// ���α׷� ���� : �˻��׸������Ȳ(��ȭ������)
// �ý���        : ���հ���
// ��������      : 2010.04.13
// ������        : ����ȣ
// ��������      :
// �������      :
//==============================================================================
// ��������      : 2014.11.03
// ������        : ������
// ��������      : ��� ��� ����
// �������      : [����-������]
//==============================================================================
// ��������      : 2014.12.11
// ������        : ������
// ��������      : Ư�� ���� �������� ���� ��������� ����
// �������      :
//==============================================================================
// ��������      : 2015.05.29
// ������        : ������
// ��������      : TPLA(S036) �׸��߰�.
// �������      : ���� ������ܰ˻�������1500435
//==============================================================================
// ��������      : 2015.06.17
// ������        : ������
// ��������      : C022 �������ܹ�A, C023 �������ܹ� B, C080 LP(a) ���� ��û
// �������      : ���� ������ܰ˻�������1500238
//==============================================================================
// ��������      : 2015.07.10
// ������        : ������
// ��������      : C022 �������ܹ�A, C023 �������ܹ� B, C080 LP(a) ��ȸ �����ϵ��� ����
// �������      :
//==============================================================================
// ��������      : 2015.07.15
// ������        : ������
// ��������      : C074 ���� �˻縸 ��ȸ�ǵ��� (cod_jisa='15' or gubn_hulgum='3')
// �������      : ����-���ϴ�
//==============================================================================
// ��������      : 2016.03.29
// ������        : �ڼ���
// ��������      : CO78�׸��ڵ� �߰�
// �������      : ���� ������ܰ˻������� 1600348
//==============================================================================
// ��������      : 2018.04.25
// ������        : �ڼ���
// ��������      : CO78�׸� ���� �����ڸ� (cod_jisa='15')
// �������      : ���� ����Ϲ���Ʈ 1800324
//==============================================================================
unit LD4Q28;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q28 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    RB_02: TRadioButton;
    RB_01: TRadioButton;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Label4: TLabel;
    RadioGroup1: TRadioGroup;
    qryProfile: TQuery;
    chk_noHomo: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vBUN, UV_vName, UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006,
    UV_v007, UV_v008, UV_v009, UV_v010, UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016, UV_v017, UV_v018    : Variant;
    //20150617 {UV_v001(C022), UV_v002(C023), UV_v005(C080)}

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q28: TfrmLD4Q28;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q281;

{$R *.DFM}

procedure TfrmLD4Q28.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q28.MouseWheelHandler(var Message: TMessage);  //����
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //����
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster.Refresh; //����
      end;
   end;
end;

procedure TfrmLD4Q28.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_v001   := VarArrayCreate([0, 0], varOleStr);
      UV_v002   := VarArrayCreate([0, 0], varOleStr);
      UV_v003   := VarArrayCreate([0, 0], varOleStr);
      UV_v004   := VarArrayCreate([0, 0], varOleStr);
      UV_v005   := VarArrayCreate([0, 0], varOleStr);
      UV_v006   := VarArrayCreate([0, 0], varOleStr);
      UV_v007   := VarArrayCreate([0, 0], varOleStr);
      UV_v008   := VarArrayCreate([0, 0], varOleStr);
      UV_v009   := VarArrayCreate([0, 0], varOleStr);
      UV_v010   := VarArrayCreate([0, 0], varOleStr);
      UV_v011   := VarArrayCreate([0, 0], varOleStr);
      UV_v012   := VarArrayCreate([0, 0], varOleStr);
      UV_v013   := VarArrayCreate([0, 0], varOleStr);
      UV_v014   := VarArrayCreate([0, 0], varOleStr);
      UV_v015   := VarArrayCreate([0, 0], varOleStr);
      UV_v016   := VarArrayCreate([0, 0], varOleStr);
      UV_v017   := VarArrayCreate([0, 0], varOleStr);
      UV_v018   := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,   iLength);
      VarArrayRedim(UV_vName, iLength);
      VarArrayRedim(UV_vBUN,  iLength);
      VarArrayRedim(UV_v001,  iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
      VarArrayRedim(UV_v008,  iLength);
      VarArrayRedim(UV_v009,  iLength);
      VarArrayRedim(UV_v010,  iLength);
      VarArrayRedim(UV_v011,  iLength);
      VarArrayRedim(UV_v012,  iLength);
      VarArrayRedim(UV_v013,  iLength);
      VarArrayRedim(UV_v014,  iLength);
      VarArrayRedim(UV_v015,  iLength);
      VarArrayRedim(UV_v016,  iLength);
      VarArrayRedim(UV_v017,  iLength);
      VarArrayRedim(UV_v018,  iLength);
   end;
end;

function TfrmLD4Q28.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q28.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // �������ڷ� �⺻����
   mskDate.Text := GV_sToday;
end;

procedure TfrmLD4Q28.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;


   if chk_noHomo.checked = false then
   begin
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1  : Value := UV_vNo[DataRow-1];   //Rack No
      2  : Value := UV_vBUN[DataRow-1];  //���ֹ�ȣ or ���ù�ȣ
      3  : Value := UV_vName[DataRow-1]; //����
      4  : Value := UV_v001[DataRow-1];  //C022
      5  : Value := UV_v002[DataRow-1];  //C023
      6  : Value := UV_v005[DataRow-1];  //C080
      7  : Value := UV_v003[DataRow-1];  //C033
      8  : Value := UV_v004[DataRow-1];  //C034
      9  : Value := UV_v006[DataRow-1];  //C074
      10 : Value := UV_v007[DataRow-1];  //C078
      11 : Value := UV_v015[DataRow-1];  //S036
      12 : Value := UV_v016[DataRow-1];  //C084
      13 : Value := UV_v017[DataRow-1];  //C085 ..C084,C085 �߰� 20160708  ����
      14 : Value := UV_v008[DataRow-1];  //C082
      15 : Value := UV_v018[DataRow-1];  //C090 ..20180425
      16 : Value := UV_v010[DataRow-1];  //GL02
      17 : Value := UV_v011[DataRow-1];  //GL03
      18 : Value := UV_v012[DataRow-1];  //GL04
      19 : Value := UV_v013[DataRow-1];  //GL05
      20 : Value := UV_v014[DataRow-1]; //GL06
   end;
    grdMaster.Col[1].Heading   := 'Rack No.';
    grdMaster.Col[2].Heading   := '���ù�ȣ';
    grdMaster.Col[3].Heading   := '����';
    grdMaster.Col[4].Heading   := '�������ܹ�A';
    grdMaster.Col[5].Heading   := '�������ܹ�B';
    grdMaster.Col[6].Heading   := 'LP(a)';
    grdMaster.Col[7].Heading   := '���� ����';
    grdMaster.Col[8].Heading   := '��ȭ�ܹ�';
    grdMaster.Col[9].Heading   := 'LDL-�ݷ����׷�';
    grdMaster.Col[10].Heading  := 'Homocysteine';
    grdMaster.Col[11].Heading  := 'TPLA';
    grdMaster.Col[12].Heading  := 'Ȱ����ҷ�(d-ROMs)';
    grdMaster.Col[13].Heading  := '���׻�ȭ�°˻�(BAP)';
    grdMaster.Col[14].Heading  := 'Lipase';
    grdMaster.Col[15].Heading  := 'H.pylori';
    grdMaster.Col[16].Heading  := '���� 30�� Glucose';
    grdMaster.Col[17].Heading  := '���� 60�� Glucose';
    grdMaster.Col[18].Heading  := '���� 90�� Glucose';
    grdMaster.Col[19].Heading  := '���� 120�� Glucose';
    grdMaster.Col[20].Heading  := '���� 180�� Glucose';
   end
   else
   begin
      case DataCol of
      1  : Value := UV_vNo[DataRow-1];   //Rack No
      2  : Value := UV_vBUN[DataRow-1];  //���ֹ�ȣ or ���ù�ȣ
      3  : Value := UV_vName[DataRow-1]; //����
      4  : Value := UV_v001[DataRow-1];  //C022
      5  : Value := UV_v002[DataRow-1];  //C023
      6  : Value := UV_v005[DataRow-1];  //C080
      7  : Value := UV_v003[DataRow-1];  //C033
      8  : Value := UV_v004[DataRow-1];  //C034
      9  : Value := UV_v006[DataRow-1];  //C074
      10 : Value := UV_v015[DataRow-1];  //S036
      11 : Value := UV_v016[DataRow-1];  //C084
      12 : Value := UV_v017[DataRow-1];  //C085 ..C084,C085 �߰� 20160708  ����
      13 : Value := UV_v008[DataRow-1];  //C082
      14 : Value := UV_v018[DataRow-1];  //c090 ..20180425 ����
      15 : Value := UV_v010[DataRow-1];  //GL02
      16 : Value := UV_v011[DataRow-1];  //GL03
      17 : Value := UV_v012[DataRow-1];  //GL04
      18 : Value := UV_v013[DataRow-1];  //GL05
      19 : Value := UV_v014[DataRow-1];  //GL06
  end;
    grdMaster.Col[1].Heading   := 'Rack No.';
    grdMaster.Col[2].Heading   := '���ù�ȣ';
    grdMaster.Col[3].Heading   := '����';
    grdMaster.Col[4].Heading   := '�������ܹ�A';
    grdMaster.Col[5].Heading   := '�������ܹ�B';
    grdMaster.Col[6].Heading   := 'LP(a)';
    grdMaster.Col[7].Heading   := '���� ����';
    grdMaster.Col[8].Heading   := '��ȭ�ܹ�';
    grdMaster.Col[9].Heading   := 'LDL-�ݷ����׷�';
    grdMaster.Col[10].Heading  := 'TPLA';
    grdMaster.Col[11].Heading  := 'Ȱ����ҷ�(d-ROMs)';
    grdMaster.Col[12].Heading  := '���׻�ȭ�°˻�(BAP)';
    grdMaster.Col[13].Heading  := 'Lipase';
    grdMaster.Col[14].Heading  := 'H.pylori';
    grdMaster.Col[15].Heading  := '���� 30�� Glucose';
    grdMaster.Col[16].Heading  := '���� 60�� Glucose';
    grdMaster.Col[17].Heading  := '���� 90�� Glucose';
    grdMaster.Col[18].Heading  := '���� 120�� Glucose';
    grdMaster.Col[19].Heading  := '���� 180�� Glucose';
   end;
end;

procedure TfrmLD4Q28.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    Sum_v001, Sum_v002, Sum_v003, Sum_v004, Sum_v005, Sum_v006, Sum_v007,
    Sum_v008, Sum_v009, Sum_v010, Sum_v011, Sum_v012, Sum_v013, Sum_v014, Sum_v015, Sum_v016, Sum_v017, Sum_v018 : Integer;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Sum_v001 := 0; Sum_v002 := 0; Sum_v003 := 0; Sum_v004 := 0; Sum_v005 := 0;
   Sum_v006 := 0; Sum_v007 := 0; Sum_v008 := 0; Sum_v009 := 0; Sum_v010 := 0;
   Sum_v011 := 0; Sum_v012 := 0; Sum_v013 := 0; Sum_v014 := 0; Sum_v015 := 0;
   Sum_v016 := 0; Sum_v017 := 0; Sum_v018 := 0; Index := 0;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
      // SQL�� ����
      Close;
      sSelect := ' SELECT  D.Desc_dc, B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp, G.desc_name, G.num_id, G.dat_gmgn, B.dat_bunju, G.cod_jisa, G.gubn_nosin, G.gubn_adult,  ' +
                 '         G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul,          ' +
                 '         G.cod_cancer, G.cod_chuga, G.gubn_hulgum, T.cod_prf, T.cod_chuga As Tcod_chuga FROM bunju B with(nolock)                                                ' +
                 '         Inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn                                         ' +
                 '         Left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ' +
                 '         Left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc                                                                                            ' ;

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ';
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + mskDate.Text + ''' and (Gubn_Part  = ''01''))';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if RB_01.Checked then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY B.Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q28', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

         if (chk_noHomo.Checked = false) then
         begin
           if (pos('C022',sHul_List) > 0) or (pos('C023',sHul_List) > 0) or
              (pos('C033',sHul_List) > 0) or (pos('C034',sHul_List) > 0) or
              (pos('C080',sHul_List) > 0) or ((pos('C074',sHul_List) > 0) and (qryBunju.FieldByName('cod_jisa').AsString = '15'))  or
              (pos('C078',sHul_List) > 0) or (pos('C082',sHul_List) > 0) or
              (pos('GL01',sHul_List) > 0) or (pos('GL02',sHul_List) > 0) or
              (pos('GL03',sHul_List) > 0) or (pos('GL04',sHul_List) > 0) or
              (pos('GL05',sHul_List) > 0) or (pos('GL06',sHul_List) > 0) or
              (pos('S036',sHul_List) > 0) or (pos('C084',sHul_List) > 0) or
              (pos('C085',sHul_List) > 0) or (pos('C090',sHul_List) > 0) then
           begin
              UP_VarMemSet(Index+1);
              UV_vNo[Index]  := qryBunju.FieldByName('desc_rackno').AsString;
              if RB_01.Checked then
                 UV_vBUN[Index] := qryBunju.FieldByName('num_samp').AsString
              else
                 UV_vBUN[Index] := qryBunju.FieldByName('num_bunju').AsString;
              UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;


              if pos('C022', sHul_List) > 0 then
              begin
                 UV_v001[Index] := '*';
                 Sum_v001 := Sum_v001 + 1;
              end
              else                               UV_v001[Index] := '';          //Sum_v001
              if pos('C023', sHul_List) > 0 then
              begin
                 UV_v002[Index] := '*';
                 Sum_v002 := Sum_v002 + 1;
              end
              else                               UV_v002[Index] := '';          //Sum_v002
              if pos('C033', sHul_List) > 0 then
              begin
                 UV_v003[Index] := '*';
                 Sum_v003 := Sum_v003 + 1;
              end
              else                               UV_v003[Index] := '';          //Sum_v003
              if pos('C034', sHul_List) > 0 then
              begin
                 UV_v004[Index] := '*';
                 Sum_v004 := Sum_v004 + 1;
              end
              else                               UV_v004[Index] := '';          //Sum_v004
              if pos('C080', sHul_List) > 0 then
              begin
                 UV_v005[Index] := '*';
                 Sum_v005 := Sum_v005 + 1;
              end
              else                               UV_v005[Index] := '';          //Sum_v005
              if (pos('C074', sHul_List) > 0) and
                 (qryBunju.FieldByName('cod_jisa').AsString = '15') then   //20150715 ������
              begin
                 UV_v006[Index] := '*';
                 Sum_v006 := Sum_v006 + 1;
              end
              else                               UV_v006[Index] := '';          //Sum_v006
              if (pos('C078', sHul_List) > 0) and (chk_noHomo.checked = false) then
              begin
                 UV_v007[Index] := '*';
                 Sum_v007 := Sum_v007 + 1;
              end
              else                               UV_v007[Index] := '';          //Sum_v007
              if pos('C082', sHul_List) > 0 then
              begin
                 UV_v008[Index] := '*';
                 Sum_v008 := Sum_v008 + 1;
              end
              else                               UV_v008[Index] := '';          //Sum_v008
              if pos('C084', sHul_List) > 0 then
              begin
                 UV_v016[Index] := '*';
                 Sum_v016 := Sum_v016 + 1;
              end
              else                               UV_v016[Index] := '';          //Sum_v016
              if pos('C085', sHul_List) > 0 then
              begin
                 UV_v017[Index] := '*';
                 Sum_v017 := Sum_v017 + 1;
              end
              else                               UV_v017[Index] := '';          //Sum_v017
              if pos('C090', sHul_List) > 0 then
              begin
                 UV_v018[Index] := '*';
                 Sum_v018 := Sum_v018 + 1;
              end
              else                               UV_v018[Index] := '';          //Sum_v018
              if pos('GL01', sHul_List) > 0 then
              begin
                 UV_v009[Index] := '*';
                 Sum_v009 := Sum_v009 + 1;
              end
              else                               UV_v009[Index] := '';          //Sum_v009
              if pos('GL02', sHul_List) > 0 then
              begin
                 UV_v010[Index] := '*';
                 Sum_v010 := Sum_v010 + 1;
              end
              else                               UV_v010[Index] := '';          //Sum_v010
              if pos('GL03', sHul_List) > 0 then
              begin
                 UV_v011[Index] := '*';
                 Sum_v011 := Sum_v011 + 1;
              end
              else                               UV_v011[Index] := '';          //Sum_v011
              if pos('GL04', sHul_List) > 0 then
              begin
                 UV_v012[Index] := '*';
                 Sum_v012 := Sum_v012 + 1;
              end
              else                               UV_v012[Index] := '';          //Sum_v012
              if pos('GL05', sHul_List) > 0 then
              begin
                 UV_v013[Index] := '*';
                 Sum_v013 := Sum_v013 + 1;
              end
              else                               UV_v013[Index] := '';          //Sum_v013
              if pos('GL06', sHul_List) > 0 then
              begin
                 UV_v014[Index] := '*';
                 Sum_v014 := Sum_v014 + 1;
              end
              else                               UV_v014[Index] := '';          //Sum_v014
              if pos('S036', sHul_List) > 0 then
              begin
                 UV_v015[Index] := '*';
                 Sum_v015 := Sum_v015 + 1;
              end
              else                               UV_v015[Index] := '';          //Sum_v015

              Inc(Index);
           end;
      end
      else if  chk_noHomo.checked = true then
      begin
           if (pos('C022',sHul_List) > 0) or (pos('C023',sHul_List) > 0) or
              (pos('C033',sHul_List) > 0) or (pos('C034',sHul_List) > 0) or
              (pos('C080',sHul_List) > 0) or ((pos('C074',sHul_List) > 0) and (qryBunju.FieldByName('cod_jisa').AsString = '15')) or
              (pos('C082',sHul_List) > 0) or (pos('C085',sHul_List) > 0) or
              (pos('GL01',sHul_List) > 0) or (pos('GL02',sHul_List) > 0) or
              (pos('GL03',sHul_List) > 0) or (pos('GL04',sHul_List) > 0) or
              (pos('GL05',sHul_List) > 0) or (pos('GL06',sHul_List) > 0) or
              (pos('S036',sHul_List) > 0) or (pos('C084',sHul_List) > 0) or
              (pos('C090',sHul_List) > 0) then
           begin
              UP_VarMemSet(Index+1);
              UV_vNo[Index]  := qryBunju.FieldByName('desc_rackno').AsString;
              if RB_01.Checked then
                 UV_vBUN[Index] := qryBunju.FieldByName('num_samp').AsString
              else
                 UV_vBUN[Index] := qryBunju.FieldByName('num_bunju').AsString;
              UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;


              if pos('C022', sHul_List) > 0 then
              begin
                 UV_v001[Index] := '*';
                 Sum_v001 := Sum_v001 + 1;
              end
              else                               UV_v001[Index] := '';          //Sum_v001
              if pos('C023', sHul_List) > 0 then
              begin
                 UV_v002[Index] := '*';
                 Sum_v002 := Sum_v002 + 1;
              end
              else                               UV_v002[Index] := '';          //Sum_v002
              if pos('C033', sHul_List) > 0 then
              begin
                 UV_v003[Index] := '*';
                 Sum_v003 := Sum_v003 + 1;
              end
              else                               UV_v003[Index] := '';          //Sum_v003
              if pos('C034', sHul_List) > 0 then
              begin
                 UV_v004[Index] := '*';
                 Sum_v004 := Sum_v004 + 1;
              end
              else                               UV_v004[Index] := '';          //Sum_v004
              if pos('C080', sHul_List) > 0 then
              begin
                 UV_v005[Index] := '*';
                 Sum_v005 := Sum_v005 + 1;
              end
              else                               UV_v005[Index] := '';          //Sum_v005
              if (pos('C074', sHul_List) > 0) and
                 ((qryBunju.FieldByName('cod_jisa').AsString = '15')   or
                  (qryBunju.FieldByName('gubn_hulgum').AsString = '1') or
                  (qryBunju.FieldByName('gubn_hulgum').AsString = '3')) then   //20150715 ������
              begin
                 UV_v006[Index] := '*';
                 Sum_v006 := Sum_v006 + 1;
              end
              else                               UV_v006[Index] := '';          //Sum_v006
              if pos('C082', sHul_List) > 0 then
              begin
                 UV_v008[Index] := '*';
                 Sum_v008 := Sum_v008 + 1;
              end
              else                               UV_v008[Index] := '';          //Sum_v008
              if pos('C084', sHul_List) > 0 then
              begin
                 UV_v016[Index] := '*';
                 Sum_v016 := Sum_v016 + 1;
              end
              else                               UV_v016[Index] := '';          //Sum_v016
              if pos('C085', sHul_List) > 0 then
              begin
                 UV_v017[Index] := '*';
                 Sum_v017 := Sum_v017 + 1;
              end
              else                               UV_v017[Index] := '';          //Sum_v017
              if pos('C090', sHul_List) > 0 then
              begin
                 UV_v018[Index] := '*';
                 Sum_v018 := Sum_v018 + 1;
              end
              else                               UV_v018[Index] := '';          //Sum_v018
              if pos('GL01', sHul_List) > 0 then
              begin
                 UV_v009[Index] := '*';
                 Sum_v009 := Sum_v009 + 1;
              end
              else                               UV_v009[Index] := '';          //Sum_v009
              if pos('GL02', sHul_List) > 0 then
              begin
                 UV_v010[Index] := '*';
                 Sum_v010 := Sum_v010 + 1;
              end
              else                               UV_v010[Index] := '';          //Sum_v010
              if pos('GL03', sHul_List) > 0 then
              begin
                 UV_v011[Index] := '*';
                 Sum_v011 := Sum_v011 + 1;
              end
              else                               UV_v011[Index] := '';          //Sum_v011
              if pos('GL04', sHul_List) > 0 then
              begin
                 UV_v012[Index] := '*';
                 Sum_v012 := Sum_v012 + 1;
              end
              else                               UV_v012[Index] := '';          //Sum_v012
              if pos('GL05', sHul_List) > 0 then
              begin
                 UV_v013[Index] := '*';
                 Sum_v013 := Sum_v013 + 1;
              end
              else                               UV_v013[Index] := '';          //Sum_v013
              if pos('GL06', sHul_List) > 0 then
              begin
                 UV_v014[Index] := '*';
                 Sum_v014 := Sum_v014 + 1;
              end
              else                               UV_v014[Index] := '';          //Sum_v014
              if pos('S036', sHul_List) > 0 then
              begin
                 UV_v015[Index] := '*';
                 Sum_v015 := Sum_v015 + 1;
              end
              else                               UV_v015[Index] := '';          //Sum_v015

              Inc(Index);
           end;
         end;
            Next;
         End;
         // Table�� Disconnect
         Close;

         UV_vNo[Index]   := '�� �� ��';
         UV_vBUN[Index]  := '';
         UV_vName[Index] := '';
         UV_v001[Index]  := Sum_v001;
         UV_v002[Index]  := Sum_v002;
         UV_v003[Index]  := Sum_v003;
         UV_v004[Index]  := Sum_v004;
         UV_v005[Index]  := Sum_v005;
         UV_v006[Index]  := Sum_v006;
         UV_v007[Index]  := Sum_v007;
         UV_v008[Index]  := Sum_v008;
         UV_v009[Index]  := Sum_v009;
         UV_v010[Index]  := Sum_v010;
         UV_v011[Index]  := Sum_v011;
         UV_v012[Index]  := Sum_v012;
         UV_v013[Index]  := Sum_v013;
         UV_v014[Index]  := Sum_v014;
         UV_v015[Index]  := Sum_v015;
         UV_v016[Index]  := Sum_v016;
         UV_v017[Index]  := Sum_v017;
         UV_v018[Index]  := Sum_v018;
         Inc(Index);
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q28.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

   if RB_01.Checked then
   begin
       grdMaster.Col[2].Heading   := '���ù�ȣ';
       MEdt_SampS.Enabled := True;
       MEdt_SampE.Enabled := True;
       mskFrom.Enabled := False;
       mskTo.Enabled := False;
   end
   else
   begin
       grdMaster.Col[2].Heading   := '���ֹ�ȣ';
       MEdt_SampS.Enabled := False;
       MEdt_SampE.Enabled := False;
       mskFrom.Enabled := True;
       mskTo.Enabled := True;
   end;

end;

procedure TfrmLD4Q28.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end
   else
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD4Q28.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True   then QRCompositeReport.Preview
  else                                  QRCompositeReport.Print;


end;

procedure TfrmLD4Q28.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   mskDate.SetFocus;
end;

procedure TfrmLD4Q28.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q281 := TfrmLD4Q281.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q281.QuickRep);
  end;
end;

function TfrmLD4Q28.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';
   with qryPf_hm do
   begin
      // 1. ���׿� ���� �˻��׸� ����
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. ���翡 ���� �˻��׸� ����
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. ��� ���� �˻��׸� ����
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. �߰��ڵ忡 ���� �˻��׸� ����
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. �׸��ڵ忡 ���� �˻��׸� ����
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
            With qryProfile do
            Begin
               Close;
               ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
               Open;
               For I := 1 To Recordcount Do
               Begin
                  if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                     sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                  Next;
               End;
               Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;

   result := sTemp;
end;

function TfrmLD4Q28.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryNo_hangmok do
   begin
      Active := False;
      ParamByName('dat_apply').AsString   := sDate;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('gubn_yh').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
      end;
      Active := False;
   end;
end;

end.

