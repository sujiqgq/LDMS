//==============================================================================
// ���α׷� ���� : EIA������Ȳ
// �ý���        : ���հ���
// ��������      : 2014.09.03
// ������        : ������
// ��������      : T039 ��ȸ
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.12.11
// ������        : ������
// ��������      : Ư�� ���� �������� ���� ��������� ����
// �������      :
//==============================================================================

unit LD4Q61;

interface
                                                
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q61 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    qryProfile: TQuery;
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
    UV_vNo, UV_vName, UV_vT039, UV_vBUN, UV_vDesc_dc : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;    
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q61: TfrmLD4Q61;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q611;

{$R *.DFM}

procedure TfrmLD4Q61.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q61.MouseWheelHandler(var Message: TMessage);  //����
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

procedure TfrmLD4Q61.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vT039 := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vBUN,  iLength);
      VarArrayRedim(UV_vT039, iLength);
      VarArrayRedim(UV_vDesc_dc, iLength);
   end;
end;

function TfrmLD4Q61.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q61.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);
      Cmb_gubn.ItemIndex := 1;

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q61.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vBUN[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vT039[DataRow-1];
   end;

end;

procedure TfrmLD4Q61.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vT039 : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;
   Sum_vT039:=0;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
      Close;

      sSelect := ' SELECT  D.Desc_dc, B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, B.dat_bunju, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B with(nolock) inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha'+
                 '          left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc';
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ';
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + mskDate.Text + ''' and (Gubn_Part  = ''05'' OR Gubn_part = ''07''))';



      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if Cmb_gubn.Text = '���ֹ�ȣ' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';   
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q61', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

            if (pos('T039',sHul_List) > 0) then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := qryBunju.FieldByName('desc_rackno').AsString;
//               UV_vBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               if Cmb_gubn.ItemIndex = 0 then
                  UV_VBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString
               else if  Cmb_gubn.ItemIndex = 1 then
               begin
                    UV_VBUN[Index] := qryBunju.FieldByName('num_samp').AsString;
                    grdMaster.Col[2].Heading :='���ù�ȣ';
               end;
               UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;

               UV_vDesc_dc[Index] := qryBunju.FieldByName('desc_dc').AsString;
               if pos('T039', sHul_List) > 0 then
               begin
                  UV_vT039[Index] := '';
                  Sum_vT039 := Sum_vT039 + 1;
               end
               else                               UV_vT039[Index] := '*';          //Sum_vT039

               Inc(Index);
            end;

            Next;
         End;

         // Table�� Disconnect
         Close;

         UV_vNo[Index]     := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '�� �� ��';
         UV_vT039[Index]   := Sum_vT039;

         UV_vDesc_dc[Index] := '';
         Inc(index);
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

procedure TfrmLD4Q61.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q61.UP_Click(Sender: TObject);
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

procedure TfrmLD4Q61.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then QRCompositeReport.Preview
  else                                QRCompositeReport.Print;

end;

procedure TfrmLD4Q61.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   Cmb_gubn.ItemIndex := 1;
   mskDate.SetFocus;
end;

procedure TfrmLD4Q61.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q611 := TfrmLD4Q611.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q611.QuickRep);
  end;
end;

function TfrmLD4Q61.UF_hangmokList : String;
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

function TfrmLD4Q61.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
