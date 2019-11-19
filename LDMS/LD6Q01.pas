//==============================================================================
// ���α׷� ���� : �׸� ���纰 �����ο� ��Ȳ
// �ý���        : ���հ���
// ��������      : 1999.11.3
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
unit LD6Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD6Q01 = class(TfrmSingle)
    qrygumgin: TQuery;
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    mskST_date: TDateEdit;
    btnDate: TSpeedButton;
    mskED_date: TDateEdit;
    btnDATE1: TSpeedButton;
    Label1: TLabel;
    pnlJisa: TPanel;
    Label3: TLabel;
    cmbJisa: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL��
    UV_sBasicSQL : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);    
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_vDat_bunju, UV_vNum_bunju, UV_vDesc_name, UV_vNum_jumin, UV_vCod_jisa, UV_vDat_gmgn,
    UV_vnum_id : Variant;
  end;

var
  frmLD6Q01: TfrmLD6Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q031;

{$R *.DFM}

procedure TfrmLD6Q01.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];

   //--------------------------------------------------------------------------
   // ��������
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vdat_bunju[Lo];  UV_vDat_bunju[Lo] := UV_vDat_bunju[Hi]; UV_vDat_bunju[Hi] := T;
            T := UV_vNum_bunju[Lo];  UV_vNum_bunju[Lo] := UV_vNum_bunju[Hi]; UV_vNum_bunju[Hi] := T;
            T := UV_vDesc_name[Lo];  UV_vDesc_name[Lo] := UV_vDesc_name[Hi]; UV_vDesc_name[Hi] := T;
            T := UV_vNum_Jumin[Lo];  UV_vNum_jumin[Lo] := UV_vNum_jumin[Hi]; UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];   UV_vDat_gmgn[Lo]  := UV_vDat_gmgn[Hi];  UV_vDat_gmgn[Hi]  := T;
            T := UV_vcod_jisa[Lo];   UV_vCod_jisa[Lo]  := UV_vCod_jisa[Hi];  UV_vCod_jisa[Hi]  := T;
            T := UV_vNum_id[Lo];     UV_vNum_id[Lo]    := UV_vNum_id[Hi];    UV_vNum_id[Hi]    := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // ��������
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vdat_bunju[Lo];  UV_vDat_bunju[Lo] := UV_vDat_bunju[Hi]; UV_vDat_bunju[Hi] := T;
            T := UV_vNum_bunju[Lo];  UV_vNum_bunju[Lo] := UV_vNum_bunju[Hi]; UV_vNum_bunju[Hi] := T;
            T := UV_vDesc_name[Lo];  UV_vDesc_name[Lo] := UV_vDesc_name[Hi]; UV_vDesc_name[Hi] := T;
            T := UV_vNum_Jumin[Lo];  UV_vNum_jumin[Lo] := UV_vNum_jumin[Hi]; UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];   UV_vDat_gmgn[Lo]  := UV_vDat_gmgn[Hi];  UV_vDat_gmgn[Hi]  := T;
            T := UV_vcod_jisa[Lo];   UV_vCod_jisa[Lo]  := UV_vCod_jisa[Hi];  UV_vCod_jisa[Hi]  := T;
            T := UV_vNum_id[Lo];     UV_vNum_id[Lo]    := UV_vNum_id[Hi];    UV_vNum_id[Hi]    := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD6Q01.UP_GridSet;
var i : Integer;
begin
   // Grid�� ȯ�� ����
   with grdMaster do
   begin
      // Row���� �ʱ�ȭ
      Rows := 0;

      // Column Title
      Col[1].Heading := '������';
      Col[2].Heading := '���ֹ�ȣ';
      Col[3].Heading := '��  ��';
      Col[4].Heading := '�ֹι�ȣ';
      Col[5].Heading := '������';
      Col[6].Heading := '����';
      col[7].Heading := 'ȸ����ȣ' ;
      // Column Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
       for i := 1 to Cols do
       Col[i].SortPicture := spDown;

   end;
end;

procedure TfrmLD6Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����

      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vcod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vNum_jumin, iLength);
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vcod_jisa,  iLength);
      VarArrayRedim(UV_vNum_id,    iLength);
   end;
end;

function TfrmLD6Q01.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskST_Date.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD6Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_GridSet;
   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);
   // Login ���簡 '00'�̸� '15'�� ��ü
   if GV_sJisa = '00' then UV_sCod_jisa := '15'
   else                    UV_sCod_jisa := GV_sJisa;

end;

procedure TfrmLD6Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vdat_bunju[DataRow-1];
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := UV_vNum_jumin[DataRow-1];
      5 : Value := UV_vDat_gmgn[DataRow-1];
      6 : Value := UV_vCod_jisa[DataRow-1];
      7 : Value := UV_vNum_id[DataRow-1];

   end;
end;

procedure TfrmLD6Q01.btnQueryClick(Sender: TObject);
var i, index : integer;
begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid  �ʱ�ȭ
   grdMaster.Rows := 0;
   index := 0;

   with  qryGumgin do
   begin
      Active := False;
      ParamByName('St_Date').AsString := mskSt_Date.text;
      ParamByName('ED_Date').AsString := mskED_Date.text;
      ParamByName('Cod_jisa').AsString := Copy(cmbJisa.Text, 1, 2);
      Active := true;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD6Q01', 'Q', 'Y');
         for i := 0 to RecordCount - 1 do
         begin
            UP_VarMemSet(RecordCount-1);
            UV_vDat_bunju[index]   := FieldByName('Dat_bunju').AsString;
            UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
            UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
            UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
            UV_vDat_gmgn[index]    := FieldByName('dat_gmgn').AsString;
            UV_vCod_jisa[index]    := FieldByName('Cod_jisa').AsString;
            UV_vNum_id[index]      := FieldByName('Num_id').AsString;
            Inc(index);
            Next;
          end;
         // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
       Active := False;
    end
    else
    begin
      GF_MsgBox('Information', 'NODATA');
      exit;
    end;
 end;

 // Grid�� �ڷḦ �Ҵ�
 grdMaster.Rows := index;

 // Query Mode & Msg
 if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD6Q01.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD6Q01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate  then GF_CalendarClick(mskST_Date)
   else
   if Sender = btnDate1 then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD6Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskST_Date then UP_Click(btnDate)
   else if Sender = mskED_Date then UP_Click(btnDate1)
end;

procedure TfrmLD6Q01.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

//   frmMD6Q311 := TfrmMD6Q311.Create(Self);
//   frmMD6Q311.QuickRep.Preview;
//   frmMD3Q311.QuickRep.Print;
end;

procedure TfrmLD6Q01.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

procedure TfrmLD6Q01.grdMasterHeadingClick(Sender: TObject;
  DataCol: Integer);
  var  iCnt : integer;
  sOrder : String;
begin
  inherited;
  // �ڷᰡ �����ϴ��� Check
   if grdMaster.Rows = 0 then exit;

   iCnt := grdMaster.Rows;

   if grdMaster.Col[DataCol].SortPicture = spUp then
   begin
      grdMaster.Col[DataCol].SortPicture := spDown;
      sOrder := 'A';
   end
   else
   begin
      grdMaster.Col[DataCol].SortPicture := spUp;
      sOrder := 'D';
   end;

   case DataCol of
      1 : UP_QuickSort(sOrder, UV_vDat_Bunju,    0, iCnt-1);
      2 : UP_QuickSort(sOrder, UV_vNum_bunju,    0, iCnt-1);
      3 : UP_QuickSort(sOrder, UV_vDesc_name,    0, iCnt-1);
      4 : UP_QuickSort(sOrder, UV_vNum_jumin,      0, iCnt-1);
      5 : UP_QuickSort(sOrder, UV_vDat_gmgn,   0, iCnt-1);
      6 : UP_QuickSort(sOrder, UV_vCod_jisa,    0, iCnt-1);
      7 : UP_QuickSort(sOrder, UV_vNum_id,   0, iCnt-1);
      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;

end;

end.
