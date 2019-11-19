//==============================================================================
// ���α׷� ���� : ���� ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2004.05.13
// ������        : ������
// ��������      : 
// �������      :
//==============================================================================
// ���α׷� ���� : C�� ���� ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2006.04.20
// ������        : ���ö
//==============================================================================
// ���α׷� ���� : ����(A��,C��)����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2010.11.03
// ��������      : A�� ���� ����� �� ������Ȳ�߰�
// ������        : �����
//==============================================================================
// ���α׷� ���� : ����(A��,C��), S019 ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2010.12.03
// ��������      : S019 ����� �� ������Ȳ�߰�
// ������        : �����
//==============================================================================
// ���α׷� ���� : ����(A��,C��), S019 ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2011.12.28
// ��������      : �̸��߰�
// ������        : ����ȣ
//==============================================================================
// ���α׷� ���� : ����(A��,C��), S019 ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2012.06.07
// ��������      : S010 �߰�
// ������        : ����
//==============================================================================


unit LD4Q54;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q54 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    pnlCond: TPanel;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    cmbJisa: TComboBox;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    GroupBox1: TGroupBox;
    RBtn_H002: TRadioButton;
    RBtn_C027: TRadioButton;
    RBtn_C028: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vSBun, UV_vName, UV_vhm, UV_vPhm,UV_vPDat_bunju,UV_vMan : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q54: TfrmLD4Q54;


implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q541;

{$R *.DFM}

procedure TfrmLD4Q54.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo      := VarArrayCreate([0, 0], varOleStr);
      UV_vSBun    := VarArrayCreate([0, 0], varOleStr);
      UV_vName    := VarArrayCreate([0, 0], varOleStr);
      UV_vhm      := VarArrayCreate([0, 0], varOleStr);
      UV_vPhm     := VarArrayCreate([0, 0], varOleStr);
      UV_vPDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vMan:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vSBun,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vhm,   iLength);
      VarArrayRedim(UV_vPhm,  iLength);
      VarArrayRedim(UV_vMan,  iLength);
      VarArrayRedim(UV_vPDat_bunju,  iLength);
   end;
end;

function TfrmLD4Q54.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q54.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD4Q54.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSBun[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vhm[DataRow-1];
      5 : Value := UV_vPhm[DataRow-1];
      6 : Value := UV_vPDat_bunju[DataRow-1]; 
      7 : Value := UV_vMan[DataRow-1];
   end;
end;
procedure TfrmLD4Q54.btnQueryClick(Sender: TObject);
var index, I, rowCount, iRet, iTemp,iAge : Integer;
    sSelect, sWhere, sOrderBy,sMan : String;
    sJangbi_List, sHul_List,hangmok_temp : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGulkwa do
   begin
      // SQL�� ����
      Close;
      sSelect := 'select K.desc_rackno,a.num_jumin, a.cod_jisa, a.num_id, a.desc_name, a.cod_jangbi, a.dat_gmgn, a.cod_hul, a.cod_chuga, b.desc_glkwa1, b.desc_glkwa2, b.desc_glkwa3, b.num_bunju From Gumgin a ';
      sSelect := sSelect + ' left outer join gulkwa b ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sSelect := sSelect + ' left  join bunju K ';
      sSelect := sSelect + ' on a.cod_jisa = K.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = K.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = K.dat_gmgn ';

      sWhere := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';
      if  (RBtn_H002.Checked =True)  then     sWhere := sWhere + ' And  b.Gubn_Part = ''' + '02' + ''' '
      else sWhere := sWhere + ' And  b.Gubn_Part = ''' + '01' + ''' ';

      //20170816 ����  //���͸��� �����Ͽ� ���� .. ��ҿ�
      {sWhere := sWhere + ' and substring(a.cod_jangbi,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(a.cod_jangbi,1,2) <> ''GS''';
      sWhere := sWhere + ' and substring(a.cod_hul,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(a.cod_hul,1,2) <> ''GS''';   }

      sWhere := sWhere + ' And b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';

       if Trim(bunjuno1.Text) <> '' then
          sWhere := sWhere + ' AND b.num_bunju >= ''' + bunjuno1.Text + '''';
       if Trim(bunjuno2.Text) <> '' then
          sWhere := sWhere + ' AND b.num_bunju <= ''' + bunjuno2.Text + '''';
       sOrderBy := ' ORDER BY G.num_bunju ';




      sOrderBy := ' ORDER BY b.Num_Bunju ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q54', 'Q', 'N');
         For I := 1 to RecordCount do
         begin
            UP_VarMemSet(index);

            // �׸�(������������)����..
            sJangbi_List := ''; sHul_List := '';
            if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //�߰��˻��׸�..
            iRet := GF_MulToSingle(qryGulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            GP_JuminToAge(qryGulkwa.FieldByname('num_jumin').AsString,qryGulkwa.FieldByname('dat_gmgn').AsString, sMan, iAge);

            with qryPGulkwa do
            begin
               close;
               ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
               ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
               ParamByName('dat_bunju').AsString := edtBDate.text;
               if (RBtn_H002.Checked = True) Then  ParamByName('Gubn_part').AsString := '02'
               else ParamByName('Gubn_part').AsString := '01';
               open;
               
               if RecordCount >= 0 then
               /////////////////////////////////////////////////////////////////
               begin
                   ////////////////////////////////
                  if (RBtn_H002.Checked) and (pos('H002', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           UV_vPhm[index] :=trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='��'
                           else if sMan='F' then UV_vMan[index] :='��';
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_C027.Checked) and (pos('C027', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,133,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,133,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='��'
                           else if sMan='F' then UV_vMan[index] :='��';
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_C028.Checked) and (pos('C028', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,139,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,139,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='��'
                           else if sMan='F' then UV_vMan[index] :='��';
                           Inc(index);
                           Gride.Progress := I;
                  end;

               end;
               ///////////////////////////////////////////////////////////////
            end;
            next;
         end;
         Gride.Progress := RecordCount;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGulkwa

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q54.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);

   if Sender = RBtn_H002 then
   Begin
         grdMaster.Col[4].Heading := '������(��)';
         grdMaster.Col[5].Heading := '������(��)';

   end
   else if Sender = RBtn_C027 then
   Begin
         grdMaster.Col[4].Heading := 'LDL(��)';
         grdMaster.Col[5].Heading := 'LDL(��)';

   end
   else if Sender = RBtn_C028 then
   Begin
         grdMaster.Col[4].Heading := '�߼�����(��)';
         grdMaster.Col[5].Heading := '�߼�����(��)';

   end;



end;



procedure TfrmLD4Q54.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q541 := TfrmLD4Q541.Create(Self);
     frmLD4Q541.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q541 := TfrmLD4Q541.Create(Self);
     frmLD4Q541.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q54.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
procedure TfrmLD4Q54.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;

end;

end.

