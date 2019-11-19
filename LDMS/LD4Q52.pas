//==============================================================================
// ���α׷� ���� : ������ ���� ����� ����Ʈ
// �ý���        : ���հ���
// ��������      : 2013.10.30
// ������        : ����
// ��������      : 
// �������      :
//==============================================================================


unit LD4Q52;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q52 = class(TfrmSingle)
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
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    chk_300: TCheckBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);

    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vSBun, UV_vName, UV_vhm, UV_vPhm,UV_vhm_1,UV_vPhm_1,UV_vNum_samp : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q52: TfrmLD4Q52;


implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q091;

{$R *.DFM}

procedure TfrmLD4Q52.UP_VarMemSet(iLength : Integer);
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
      UV_vNum_samp:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vSBun,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vhm,   iLength);
      VarArrayRedim(UV_vPhm,  iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
   end;
end;

function TfrmLD4Q52.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q52.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
      Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD4Q52.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vSBun[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vhm[DataRow-1];
      6 : Value := UV_vPhm[DataRow-1];

   end;

end;

procedure TfrmLD4Q52.btnQueryClick(Sender: TObject);
var index, I, rowCount, iRet, iTemp : Integer;
    sSelect, sWhere, sOrderBy : String;
    sJangbi_List, sHul_List : String;
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
      sSelect := 'select  a.num_samp,a.cod_jisa, a.num_id, a.desc_name, a.dat_gmgn, a.cod_hul, a.cod_chuga, b.desc_glkwa1, b.desc_glkwa2, b.desc_glkwa3, b.num_bunju From Gumgin a ';
      sSelect := sSelect + ' left outer join gulkwa b ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sWhere := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';

       sWhere := sWhere + ' And  b.Gubn_Part = ''' + '01' + ''' ';

       sWhere := sWhere + ' And b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      if Cmb_gubn.Text = '���ֹ�ȣ' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju <= ''' + bunjuno2.Text + '''';  
         sOrderBy := ' ORDER BY G.num_bunju ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), G.num_bunju '
      end;

      sOrderBy := ' ORDER BY b.Num_Bunju ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q52', 'Q', 'N');
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

            with qryPGulkwa do
            begin
               close;
               ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
               ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
               ParamByName('dat_bunju').AsString := edtBDate.text;
               ParamByName('Gubn_part').AsString := '01';


               open;

               if RecordCount >= 0 then
               /////////////////////////////////////////////////////////////////
               begin

                  if (pos('C032', sHul_List) > 0) and (chk_300.Checked=false)then
                      begin
                      UV_vNo[index]        := intTostr(index+1);
                      UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                      UV_vSBun[index]      := qryGulkwa.FieldByName('num_bunju').AsString;
                      UV_vName[index]      := qryGulkwa.FieldByName('desc_name').AsString;
                      UV_vhm[index]        := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,157,6));
                      UV_vPhm[index]       := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,157,6));
                      Inc(index);
                      Gride.Progress := I;
                  end
                  else if (pos('C032', sHul_List) > 0) and (chk_300.Checked=true) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,157,6)) <> '')  then

                      begin
                      if   (StrToInt(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,157,6)))>=300)  then
                      begin
                      UV_vNo[index]    := intTostr(index+1);
                      UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                      UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                      UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                      UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,157,6));
                      Inc(index);
                      Gride.Progress := I;
                      end;
                  end;




                  /////////////////////////////////////

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

procedure TfrmLD4Q52.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);




end;





procedure TfrmLD4Q52.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
    procedure TfrmLD4Q52.SBut_ExcelClick(Sender: TObject);
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

