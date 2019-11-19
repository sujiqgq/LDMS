//==============================================================================
// ���α׷� ���� : ���� ��� ���� ����Ʈ
// �ý���        : ���հ���
// ��������      : 2019.01.14
// ������        : �ڼ���
// ��������      :
// ��������      :
//==============================================================================

unit LD4Q35;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;
type
  TfrmLD4Q35 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    qryCell: TQuery;
    Gride: TGauge;
    QRCompositeReport: TQRCompositeReport;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    mskST_date: TDateEdit;
    btnSt_date: TSpeedButton;
    Label21: TLabel;
    mskEd_date: TDateEdit;
    btnEd_date: TSpeedButton;
    btn_gmgn: TRadioButton;
    btn_regi: TRadioButton;
    btn_regiNo: TRadioButton;
    Edt_sbuwi: TEdit;
    Edt_Ebuwi: TEdit;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;


    procedure SBut_ExcelClick(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vJisa, UV_vName, UV_vJumin, UV_vNO_GUM, UV_vDat_gmgn, UV_vdel_number, UV_vdesc_sayu : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q35: TfrmLD4Q35;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LDAI01;

{$R *.DFM}

procedure TfrmLD4Q35.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q35.MouseWheelHandler(var Message: TMessage);  //����
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

procedure TfrmLD4Q35.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo          := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa        := VarArrayCreate([0, 0], varOleStr);
      UV_vName        := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin       := VarArrayCreate([0, 0], varOleStr);
      UV_vNO_GUM      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vdel_number  := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_sayu   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo         ,  iLength);
      VarArrayRedim(UV_vJisa       ,  iLength);
      VarArrayRedim(UV_vName       ,  iLength);
      VarArrayRedim(UV_vJumin      ,  iLength);
      VarArrayRedim(UV_vNO_GUM     ,  iLength);
      VarArrayRedim(UV_vDat_gmgn   ,  iLength);
      VarArrayRedim(UV_vdel_number ,  iLength);
      VarArrayRedim(UV_vdesc_sayu  ,  iLength);
   end;
end;

function TfrmLD4Q35.UF_Condition : Boolean;
begin
   Result := True;
end;

procedure TfrmLD4Q35.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
end;

procedure TfrmLD4Q35.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vJumin[DataRow-1];
      5 : Value := UV_vNO_GUM[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vDel_number[DataRow-1];
      8 : Value := UV_vdesc_sayu[DataRow-1];
      9 : Value := '';
   end;
end;

procedure TfrmLD4Q35.btnQueryClick(Sender: TObject);
var index, i, iRet, iTemp, num : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;


   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryCell do
   begin
      // SQL�� ����
      Close;
      sSelect := ' select C.cod_jisa, A.desc_name, C.desc_buwi, C.dat_gmgn, R.num_delete, R.desc_delete, A.num_jumin From cell C';
      sSelect := sSelect + ' left outer join Gumgin A ';
      sSelect := sSelect + ' On C.cod_jisa = A.cod_jisa and C.num_id = A.num_id and C.dat_gmgn = A.dat_gmgn';
      sSelect := sSelect + ' left outer join record_cell R ';
      sSelect := sSelect + ' On C.cod_jisa = R.cod_jisa and C.num_id = R.num_id and C.dat_gmgn = R.dat_gmgn and C.cod_hm  = R.cod_hm';
      if btn_gmgn.checked then
           sWhere := ' WHERE C.dat_gmgn >= ''' + mskST_date.Text + ''' AND C.dat_gmgn <= ''' +  mskEd_date.Text + ''' '
      else if btn_regi.checked then sWhere := ' WHERE C.dat_last >= ''' + mskST_date.Text + ''' AND C.dat_last <= ''' +  mskEd_date.Text + ''' '
      else if btn_regiNo.checked then sWhere := ' WHERE C.desc_buwi >= ''' + Edt_sbuwi.Text + ''' AND C.desc_buwi <= ''' +  Edt_Ebuwi.Text + ''' ';


      sWhere := sWhere + ' AND ((R.desc_delete <> '''') or (R.num_delete <> ''''))';
      sOrderBy := ' ORDER BY CAST(C.cod_jisa AS INT), desc_buwi';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryCell, 'LDAP01', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
               Gride.Progress := I;
               UP_VarMemSet(Index+1);
               UV_vNo[Index]        := IntToStr(Index+1);
               UV_vJisa[Index]      := qryCell.FieldByName('Cod_jisa').AsString;
               UV_vName[Index]      := qryCell.FieldByName('desc_name').AsString;
               UV_vJumin[Index]     := copy(qryCell.FieldByName('num_jumin').AsString,1,6) + '-' + copy(qryCell.FieldByName('num_jumin').AsString,7,1);
               UV_vNo_Gum[Index]    := qryCell.FieldByName('desc_buwi').AsString;
               UV_vDat_gmgn[Index]  := qryCell.FieldByName('dat_gmgn').AsString;
               UV_vDel_number[Index]:= qryCell.FieldByName('num_delete').AsString;
               UV_vDesc_sayu[Index] := qryCell.FieldByName('desc_delete').AsString;

               Inc(Index);
            Next;
         End;

         // Table�� Disconnect
         Close;
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


procedure TfrmLD4Q35.SBut_ExcelClick(Sender: TObject);
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



procedure TfrmLD4Q35.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If Key = VK_F12 Then SBut_ExcelClick(Self);
end;

procedure TfrmLD4Q35.UP_Click(Sender: TObject);
begin
   inherited;

   if      Sender = btnSt_date       then GF_CalendarClick(mskST_date)
   else if Sender = btnEd_date       then GF_CalendarClick(mskEd_date);
end;

end.