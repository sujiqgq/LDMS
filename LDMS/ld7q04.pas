//==============================================================================
// ���α׷� ���� : 2015�⿬�������״���(�ݿ�,ȣ����)
// �ý���        : ���հ���
// ��������      : 2015.04.28
// ������        : ��â��
// ��������      :
// �������      :
//==============================================================================
unit LD7Q04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, ComObj;

type
  TfrmLD7Q04 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    msksDate: TDateEdit;
    btnDate: TSpeedButton;
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryGulkwa: TQuery;
    CBox_gubn: TComboBox;
    Label3: TLabel;
    qryGumgin: TQuery;
    mskeDate: TDateEdit;
    btnDate1: TSpeedButton;
    Label4: TLabel;
    SBut_Excel: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����(Report���� �����ϹǷ� Public�� ����)
    UV_vV001, UV_vV002, UV_vV003, UV_vV004, UV_vV005,
    UV_vV006, UV_vV007, UV_vV008, UV_vV009, UV_vV010,
    UV_vV011, UV_vV012 : Variant;

    UV_sCod_jisa : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }

  end;

var
  frmLD7Q04: TfrmLD7Q04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD7Q04.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vV001 := VarArrayCreate([0, 0], varOleStr);
      UV_vV002 := VarArrayCreate([0, 0], varOleStr);
      UV_vV003 := VarArrayCreate([0, 0], varOleStr);
      UV_vV004 := VarArrayCreate([0, 0], varOleStr);
      UV_vV005 := VarArrayCreate([0, 0], varOleStr);
      UV_vV006 := VarArrayCreate([0, 0], varOleStr);
      UV_vV007 := VarArrayCreate([0, 0], varOleStr);
      UV_vV008 := VarArrayCreate([0, 0], varOleStr);
      UV_vV009 := VarArrayCreate([0, 0], varOleStr);
      UV_vV010 := VarArrayCreate([0, 0], varOleStr);
      UV_vV011 := VarArrayCreate([0, 0], varOleStr);
      UV_vV012 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vV001, iLength);
      VarArrayRedim(UV_vV002, iLength);
      VarArrayRedim(UV_vV003, iLength);
      VarArrayRedim(UV_vV004, iLength);
      VarArrayRedim(UV_vV005, iLength);
      VarArrayRedim(UV_vV006, iLength);
      VarArrayRedim(UV_vV007, iLength);
      VarArrayRedim(UV_vV008, iLength);
      VarArrayRedim(UV_vV009, iLength);
      VarArrayRedim(UV_vV010, iLength);
      VarArrayRedim(UV_vV011, iLength);
      VarArrayRedim(UV_vV012, iLength);
   end;
end;

function TfrmLD7Q04.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if (msksDate.Text = '') Or (mskeDate.Text = '') then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD7Q04.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_VarMemSet(0);

   // Login ���簡 '00'�̸� '15'�� ��ü
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   Label1.Caption := '�� ��:';
   GP_ComboJisa(cmbCOD_JISA);
   GP_ComboDisplay(cmbCOD_JISA, copy(GV_sUserId,1,2), 2);
   CBox_gubn.ItemIndex := 0;
end;

procedure TfrmLD7Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := GF_DateFormat(UV_vV001[DataRow-1]);
      2 : Value := UV_vV002[DataRow-1];
      3 : Value := UV_vV003[DataRow-1];
      4 : Value := UV_vV004[DataRow-1];
      5 : Value := UV_vV005[DataRow-1];
      6 : Value := UV_vV006[DataRow-1];
      7 : Value := UV_vV007[DataRow-1];
      8 : Value := UV_vV008[DataRow-1];
   end;
end;

procedure TfrmLD7Q04.btnQueryClick(Sender: TObject);
var
   sSelect, sWhere, sOrderBy : String;
   index : Integer;
   GfGulkwa1, GfGulkwa2, GfGulkwa3, GfGulkwa, sTemp : String;
begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid & Chart �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   index := 0;
   sSelect := ''; sWhere := ''; sOrderBy := '';

   with qryGumgin do
   begin
      // SQL���� ������ �ڷḦ Query
      Active := False;
      sSelect := ' SELECT G.num_id, G.dat_gmgn, G.desc_name, G.Cod_jisa,  G.num_samp, G.num_tel, G.desc_addr, G.Cod_Chuga, ' +
                 '        SUBSTRING(G.num_jumin,1,2) + ''.'' + SUBSTRING(G.num_jumin,3,2) + ''.'' + SUBSTRING(G.num_jumin,5,2) AS BIRTH ' +
                 ' FROM gumgin G ';

      sWhere := sWhere + ' Where G.dat_gmgn >= ''' + msksDate.Text + '''';
      sWhere := sWhere + ' And   G.dat_gmgn <= ''' + mskeDate.Text + '''';
      //���� ���� ���� ��ü �۾�
//      sWhere := sWhere + ' And G.Cod_Jisa = ''' + copy(cmbCOD_JISA.Text,1,2) + '''';

      case CBox_gubn.ItemIndex of
         0 : sWhere := sWhere + ' AND (G.cod_chuga like ''%DI04%'' or G.cod_chuga like ''%DI05%'')';
         1 : sWhere := sWhere + ' AND G.cod_chuga like ''%DI05%''';
         2 : sWhere := sWhere + ' AND G.cod_chuga like ''%DI04%''';
      end;

      sOrderBy := ' ORDER BY G.dat_Gmgn, G.Desc_Name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD7Q04', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);
         // DataSet�� ���� Variant������ �̵�
         while Not qryGumgin.EOF do
         begin
            UV_vV001[index] := FieldByName('dat_gmgn').AsString;
            UV_vV002[index] := FieldByName('desc_name').AsString;
            UV_vV003[index] := FieldByName('BIRTH').AsString;
            UV_vV004[index] := FieldByName('num_id').AsString;
            With qryGulkwa Do
               Begin
                   Active := False;
                   ParamByName('Dat_Gmgn').AsSTring := qryGumgin.FieldByName('Dat_Gmgn').AsString;
                   ParamByName('Num_Id').AsSTring   := qryGumgin.FieldByName('Num_Id').AsString;
                   Active := True;
                   While Not Eof Do
                      Begin
                         GfGulkwa1 := Copy(FieldByName('DESC_GLKWA1').AsString,1,200);
                         GfGulkwa2 := Copy(FieldByName('DESC_GLKWA2').AsString,1,200);
                         GfGulkwa3 := Copy(FieldByName('DESC_GLKWA3').AsString,1,200);
                         GF_HulGulkwa('2', GfGulkwa1, GfGulkwa2, GfGulkwa3, GfGulkwa);
                         If FieldByName('Gubn_part').AsString = '01' Then
                            Begin
                                UV_vV006[index] := Trim(Copy(GfGulkwa,121,6));
                                UV_vV007[index] := Trim(Copy(GfGulkwa,139,6));
                            End;
                         If FieldByName('Gubn_part').AsString = '03' Then
                            Begin
                                UV_vV005[index] := Trim(Copy(GfGulkwa,391,6));
                            End;
                         If FieldByName('Gubn_part').AsString = '05' Then
                            Begin
                                UV_vV008[index] := Trim(Copy(GfGulkwa,313,6));
                            End;
                         Next;
                      End;
                   Active := False;
               End;
            Next;
            Inc(index);
         end;

         // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
         Active := False;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD7Q04.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // �ڷᰡ ��������� ó��
   if NewRow = 0 then exit;

   // Data�Ǽ� ǥ��
   GP_SetDataCnt(grdMaster);

   // Grid Focus
   grdMaster.SetFocus;
end;

procedure TfrmLD7Q04.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate  then GF_CalendarClick(msksDate);
   if Sender = btnDate1 then GF_CalendarClick(mskeDate);
end;

procedure TfrmLD7Q04.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = msksDate then UP_Click(btnDate);
   if Sender = mskeDate then UP_Click(btnDate1);
end;

procedure TfrmLD7Q04.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
  sTemp, sTemp2 : String;
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
//               Application.ProcessMessages;
               sTemp := grdMaster.cell[Col + 1, Row];
               while pos(#13,sTemp) > 0 do
               begin
                  sTemp2 := copy(sTemp, pos(#13,sTemp) + 1, Length(sTemp));
                  sTemp  := copy(sTemp,1,pos(#13,sTemp) -1);
                  sTemp := sTemp + sTemp2;
               end;

               ArrV3[Row, Col] := sTemp;
            end;
         end;
//         Gauge2.Progress:= i;
//         Application.ProcessMessages;
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
