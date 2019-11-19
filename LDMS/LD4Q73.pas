//==============================================================================
// ���α׷� ���� : �����׸�_���� ������Ȳ(�ﱤ)
// �ý���        : ���հ���
// ��������      : 2015.11.25
// ������        : ������
// ��������      : �ű԰���..
// �������      : ���� ������ܰ˻�������1500972(�ѹ̿� å��)
//                 **2015.12 �Ѵ޸� ��� ����
//==============================================================================
unit LD4Q73;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD4Q73 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    RadioGroup1: TRadioGroup;
    qryProfile: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    pnlCond: TPanel;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    cmbJisa: TComboBox;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Btn_NamePrint: TBitBtn;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure Btn_NamePrintClick(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure SBut_ExcelClick(Sender: TObject);
    private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vName, UV_vJumin, UV_vSampNo, UV_vRack, UV_vS008, UV_vS091 : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
     UV_vJisa : Variant;
  end;

var
  frmLD4Q73: TfrmLD4Q73;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q730, LD4Q731;

{$R *.DFM}

procedure TfrmLD4Q73.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q73.MouseWheelHandler(var Message: TMessage);  //����
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

procedure TfrmLD4Q73.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo     := VarArrayCreate([0, 0], varOleStr);
      UV_vName   := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vRack   := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo := VarArrayCreate([0, 0], varOleStr);
      UV_vS008   := VarArrayCreate([0, 0], varOleStr);
      UV_vS091   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa   := VarArrayCreate([0, 0], varOleStr);

  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo, iLength);
      VarArrayRedim(UV_vName, iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vRack, iLength);
      VarArrayRedim(UV_vSampNo, iLength);
      VarArrayRedim(UV_vS008, iLength);
      VarArrayRedim(UV_vS091, iLength);
      VarArrayRedim(UV_vJisa, iLength);
   end;
end;

function TfrmLD4Q73.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q73.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   RadioGroup1.ItemIndex:=2;
   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   edtBDate.Text := GV_sToday;
end;

procedure TfrmLD4Q73.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRack[DataRow-1];
      3 : Value := UV_vSampNo[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vS008[DataRow-1];
      7 : Value := UV_vS091[DataRow-1];
   end;
end;

procedure TfrmLD4Q73.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_nTCnt, Sum_vS008, Sum_vS091 : Integer;

    sSelect, sWhere, sGroupBy, sOrderBy, sdate : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;
   Sum_nTCnt := 0; Sum_vS008 := 0; Sum_vS091 := 0;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
      Close;
      sSelect := ' SELECT G.Cod_jisa, G.desc_name, G.num_jumin, G.dat_gmgn, G.num_samp, B.desc_RackNo, G.gubn_nosin, G.gubn_adult, G.gubn_life,  ' +
                 '        G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh,                         ' +
                 '        G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga FROM bunju B with(nolock)     ' +
                 ' join Gumgin G with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn                     ' +
                 ' left outer join Tkgum  T with(nolock) ON B.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and B.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';

      sWhere := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
      sWhere := sWhere + ' AND B.desc_RackNo like ''A%'' ';

      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And B.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And B.Num_Bunju <= ''' + BunjuNo2.Text + '''';

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(num_samp AS INT) ';
         2 : sOrderBy := ' ORDER BY CAST(num_samp AS INT), G.cod_jisa, num_Bunju, desc_name  ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if qryBunju.RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q73', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;
            if (pos('S008',sHul_List) > 0) or (pos('S091',sHul_List) > 0) Then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index] := IntToStr(82000 + Index + 1);
               Sum_nTCnt := Sum_nTCnt + 1;
               UV_vRack[Index] := qryBunju.FieldByName('desc_RackNo').AsString;
               UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[Index] := copy(qryBunju.FieldByName('num_jumin').AsString,1,7);
               UV_vJisa[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               if pos('S008', sHul_List) > 0 then
               begin
                  UV_vS008[Index] := '';
                  Sum_vS008 := Sum_vS008 + 1;
               end
               else                               UV_vS008[Index] := '*';
               if pos('S091', sHul_List) > 0 then
               begin
                  UV_vS091[Index] := '';
                  Sum_vS091 := Sum_vS091 + 1;
               end
               else                               UV_vS091[Index] := '*';
               Inc(Index);
            End;
            Next;
         End;
         // Table�� Disconnect
         Close;

         UV_vNo[Index]     := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '��  ��  ��';
         UV_vS008[Index] := Sum_vS008;
         UV_vS091[Index] := Sum_vS091;
         UV_vSampNo[Index] := IntToStr(Sum_nTCnt) + '��';
         Inc(index);
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q73.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q73.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q730 := TfrmLD4Q730.Create(Self);
  if RBtn_preveiw.Checked then frmLD4Q730.QuickRep.Preview
  else                         frmLD4Q730.QuickRep.Print;
end;

procedure TfrmLD4Q73.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

procedure TfrmLD4Q73.Btn_NamePrintClick(Sender: TObject);
begin
   inherited;
   frmLD4Q731 := TfrmLD4Q731.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q731.QuickRep.Preview
   else                         frmLD4Q731.QuickRep.Print;
end;

function TfrmLD4Q73.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, J, K : integer;
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

function TfrmLD4Q73.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q73.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.',
                             '����', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
