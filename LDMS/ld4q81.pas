//==============================================================================
// ���α׷� ���� : RIA������Ȳ
// �ý���        : ���հ���
// ��������      : 2005.11.14
// ������        : ������
// ��������      : ���Ӱ� ����..
// �������      :
//==============================================================================
// ���α׷� ���� : �����׸� List (���ֽ� ����)
// �ý���        : ���հ���
// ��������      : 2006.04.11
// ������        : ���ö
//==============================================================================
// ��������      : 2006.06.26
// ������        : ���ö
// ��������      : E008, E022, E027 �߰�
// �������      :
//==============================================================================
// ��������      : 2006.09.11
// ������        : ���ö
// ��������      : E045 �߰�
// �������      :
//==============================================================================
// ��������      : 2007.01.24
// ������        : ������
// ��������      : C070,C065,E019,SE11,SE12,T005,S080,C020 �߰�
// �������      : ����� ���� �����ؼ� �� 2�������� ���..
//==============================================================================
// ��������      : 2007.06.21
// ������        : ������
// ��������      : 007 �߰�
// �������      :
//==============================================================================
// ��������      : 2008.09.11
// ������        : ���ö
// ��������      : S011, S012 �׸� ������Ȳ���� ����.
// �������      :
//==============================================================================
// ��������      : 2013.05.22
// ������        : ����
// ��������      : SE01,E069,S079,S090 �߰�
// �������      :
//==============================================================================
// ��������      : 2014.05.03
// ������        : ������
// ��������      : ���ֱⰣ ��ȸ ����, ��ȸ����-������ �߰�
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.12.11
// ������        : ������
// ��������      : Ư�� ���� �������� ���� ��������� ����
// �������      :
//==============================================================================
// ��������      : 2015.06.17
// ������        : ������
// ��������      : ���� �߰�
// �������      : ���� ������ܰ˻�������1500460
//==============================================================================
unit LD4Q81;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q81 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    Label12: TLabel;
    Label13: TLabel;
    Label10: TLabel;
    Label9: TLabel;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    qryProfile: TQuery;
    Btn_NamePrint: TBitBtn;
    RBtn_2: TRadioButton;
    RBtn_1: TRadioButton;
    btnEdate: TSpeedButton;
    Label3: TLabel;
    mskDate_start: TDateEdit;
    mskDate_end: TDateEdit;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Cmb_gubn: TComboBox;
    edtDc: TEdit;
    cmbJisa: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Btn_NamePrintClick(Sender: TObject);


  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vName, UV_vJumin, UV_vNum_samp, UV_vBUN, UV_vHM, UV_vHM2, UV_vdat_bunju, UV_vdat_gmgn, UV_vnum_jumin, UV_vcod_jisa,
    UV_vNo : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;    
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q81: TfrmLD4Q81;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q811, LD4Q812, LD4Q813;

{$R *.DFM}

procedure TfrmLD4Q81.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q81.MouseWheelHandler(var Message: TMessage);  //����
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

procedure TfrmLD4Q81.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN       := VarArrayCreate([0, 0], varOleStr);
      UV_vHm        := VarArrayCreate([0, 0], varOleStr);
      UV_vHm2       := VarArrayCreate([0, 0], varOleStr);
      UV_vdat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vdat_gmgn  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vHM,  iLength);
      VarArrayRedim(UV_vHM2,  iLength);
      VarArrayRedim(UV_vdat_bunju,  iLength);
      VarArrayRedim(UV_vNo,  iLength);
      VarArrayRedim(UV_vdat_gmgn,  iLength);
   end;
end;

function TfrmLD4Q81.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if (mskDate_start.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end
   else if (mskDate_end.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q81.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Cmb_gubn.ItemIndex := 1;


end;

procedure TfrmLD4Q81.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vdat_bunju[DataRow-1];
      4 : Value := UV_vNum_samp[DataRow-1];
      5 : Value := UV_vBUN[DataRow-1];
      6 : Value := UV_vName[DataRow-1];
      7 : Value := UV_vHM[DataRow-1];
      8 : Value := UV_vHM2[DataRow-1];

   end;

end;

procedure TfrmLD4Q81.btnQueryClick(Sender: TObject);
var index, i, iRet, iTemp, num,
    Sum_vHM,Sum_vHM1 : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
    sjisa : string;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;


    Sum_vHM:=0;
    Sum_vHM1:=0;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
      Close;

      sSelect := ' select B.desc_rackno, B.dat_bunju, G.cod_jisa, G.num_id, G.dat_gmgn, A.num_samp, G.desc_glkwa1, G.num_bunju, G.gubn_part, A.num_Jumin, G.Dat_gmgn, D.Desc_dc, A.desc_name, ' +
                 ' A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.num_jumin, ' +
                 ' A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga ';

      sSelect := sSelect + ' From gulkwa G   left outer join Gumgin A ';
      sSelect := sSelect + ' On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';
      sSelect := sSelect + ' left outer join bunju B On G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and G.dat_gmgn = B.dat_gmgn ';
      sSelect := sSelect + ' left outer join Danche D ON A.cod_dc = D.cod_dc ';

      sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere := sWhere + ' AND B.Dat_Bunju  >= ''' + mskDate_start.Text + ''' ' +
                         ' AND B.Dat_Bunju  <= ''' + mskDate_end.Text + ''' ' +
                         ' AND  G.Gubn_Part = ''09''';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if Cmb_gubn.Text = '���ֹ�ȣ' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY B.dat_bunju, A.num_samp ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY B.dat_bunju, A.num_samp '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q15', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;



            if (pos('SE89',sHul_List) > 0) Then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index] := Index+1;
               UV_vcod_jisa[Index]   := qryBunju.FieldByName('cod_Jisa').AsString;
               UV_vdat_bunju[Index]  := qryBunju.FieldByName('dat_bunju').AsString;
               UV_VBUN[Index]        := qryBunju.FieldByName('Num_Bunju').AsString;
               UV_vNum_samp[Index]   := qryBunju.FieldByName('Num_samp').AsString;
               UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;

               if pos('E068', sHul_List) > 0 then
               begin
                    UV_vHM[Index] := '';
                    Sum_vHM := Sum_vHM + 1;
               end
               else                               UV_vHM[Index] := '*';          //Sum_vHM



               Inc(Index);
            end;
            Next;
         End;

         // Table�� Disconnect
         Close;
         UV_vNo[Index] := '';
         UV_vdat_bunju[Index]  := '';
         UV_vnum_samp[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vDat_gmgn[Index]   := '';
         UV_vName[Index]   := '��  ��  ��';
         UV_vHM[Index]   := Sum_vHM;
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

procedure TfrmLD4Q81.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q81.UP_Click(Sender: TObject);
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
   else if Sender = btnBdate then GF_CalendarClick(mskDate_start)
   else if Sender = btnEdate then GF_CalendarClick(mskDate_end);
end;

procedure TfrmLD4Q81.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then QRCompositeReport.Preview
  else                                QRCompositeReport.Print;

end;

procedure TfrmLD4Q81.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q811 := TfrmLD4Q811.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q811.QuickRep);
  end;
end;

function TfrmLD4Q81.UF_hangmokList : String;
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

function TfrmLD4Q81.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q81.FormActivate(Sender: TObject);
begin
  inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   Cmb_gubn.ItemIndex := 1;
   mskDate_start.SetFocus;
end;

procedure TfrmLD4Q81.Btn_NamePrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_1.Checked then
  begin
   frmLD4Q812 := TfrmLD4Q812.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q812.QuickRep.Preview
   else                         frmLD4Q812.QuickRep.Print;
  end
  else
   begin
   frmLD4Q813 := TfrmLD4Q813.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q813.QuickRep.Preview
   else                         frmLD4Q813.QuickRep.Print;
  end;
END;  

end.
