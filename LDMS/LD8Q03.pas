//==============================================================================
// ���α׷� ���� : �׸� ���纰 �����ο� ��Ȳ
// �ý���        : ���հ���
// ��������      : 1999.11.3
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
unit LD8Q03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD8Q03 = class(TfrmSingle)
    qrygulkwa: TQuery;
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
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL��
    UV_sBasicSQL : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_vCod_hm, UV_vDesc_kor,  UV_vJisa_15, UV_vJisa_11, UV_vJisa_33,  UV_vJisa_34,
    UV_vJisa_35, UV_vJisa_41, UV_vJisa_43,  UV_vJisa_45, UV_vJisa_46,  UV_vJisa_47,
    UV_vJisa_51, UV_vJisa_52, UV_vJisa_61,  UV_vJisa_63, UV_vJisa_71,  UV_vJisa_72,
    UV_vJisa_kita, UV_vTotal : Variant;
  end;

var
  frmLD8Q03: TfrmLD8Q03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q031;

{$R *.DFM}
procedure TfrmLD8Q03.UP_GridSet;
var i : Integer;
begin
   // Grid�� ȯ�� ����
   with grdMaster do
   begin
      // Row���� �ʱ�ȭ
      Rows := 0;

      // Column Title
      Col[1].Heading := '�˻��ڵ�';
      Col[2].Heading := '�� �� ��';
      Col[3].Heading := '������';
      Col[4].Heading := '����';
      Col[5].Heading := 'û��';
      Col[6].Heading := '����';
      Col[7].Heading := 'õ��';
      Col[8].Heading := '��õ';
      Col[9].Heading := '����';
      Col[10].Heading := '�Ⱦ�';
      Col[11].Heading := '������';
      Col[12].Heading := '����';
      Col[13].Heading := '����';
      Col[14].Heading := '����';
      Col[15].Heading := '�λ�';
      Col[16].Heading := '���';
      Col[17].Heading := '�뱸';
      Col[18].Heading := '����';
      Col[19].Heading := '��Ÿ';
      Col[20].Heading := '�հ�';
      // Column Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
   end;
end;

procedure TfrmLD8Q03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vCod_hm    := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_kor  := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_15   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_11   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_33   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_34   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_35   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_41   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_43   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_45   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_46   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_47   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_51   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_52   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_61   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_63   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_71   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_72   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa_kita := VarArrayCreate([0, 0], varOleStr);
      UV_vTotal     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vCod_hm,    iLength);
      VarArrayRedim(UV_vdesc_kor,  iLength);
      VarArrayRedim(UV_vJisa_15,   iLength);
      VarArrayRedim(UV_vJisa_11,   iLength);
      VarArrayRedim(UV_vJisa_33,   iLength);
      VarArrayRedim(UV_vJisa_34,   iLength);
      VarArrayRedim(UV_vJisa_35,   iLength);
      VarArrayRedim(UV_vJisa_41,   iLength);
      VarArrayRedim(UV_vJisa_43,   iLength);
      VarArrayRedim(UV_vJisa_45,   iLength);
      VarArrayRedim(UV_vJisa_46,   iLength);
      VarArrayRedim(UV_vJisa_47,   iLength);
      VarArrayRedim(UV_vJisa_51,   iLength);
      VarArrayRedim(UV_vJisa_52,   iLength);
      VarArrayRedim(UV_vJisa_61,   iLength);
      VarArrayRedim(UV_vJisa_63,   iLength);
      VarArrayRedim(UV_vJisa_71,   iLength);
      VarArrayRedim(UV_vJisa_72,   iLength);
      VarArrayRedim(UV_vJisa_kita, iLength);
      VarArrayRedim(UV_vTotal,     iLength);
   end;
end;

function TfrmLD8Q03.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskST_Date.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD8Q03.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_GridSet;
   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);
   // Login ���簡 '00'�̸� '15'�� ��ü
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

end;

procedure TfrmLD8Q03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vCod_hm[DataRow-1];
      2 : Value := UV_vDesc_kor[DataRow-1];
      3 : Value := UV_vJisa_15[DataRow-1];
      4 : Value := UV_vJisa_11[DataRow-1];
      5 : Value := UV_vJisa_33[DataRow-1];
      6 : Value := UV_vJisa_34[DataRow-1];
      7 : Value := UV_vJisa_35[DataRow-1];
      8 : Value := UV_vJisa_41[DataRow-1];
      9 : Value := UV_vJisa_43[DataRow-1];
     10 : Value := UV_vJisa_45[DataRow-1];
     11 : Value := UV_vJisa_46[DataRow-1];
     12 : Value := UV_vJisa_47[DataRow-1];
     13 : Value := UV_vJisa_51[DataRow-1];
     14 : Value := UV_vJisa_52[DataRow-1];
     15 : Value := UV_vJisa_61[DataRow-1];
     16 : Value := UV_vJisa_63[DataRow-1];
     17 : Value := UV_vJisa_71[DataRow-1];
     18 : Value := UV_vJisa_72[DataRow-1];
     19 : Value := UV_vJisa_kita[DataRow-1];
     20 : Value := UV_vTotal[DataRow-1];
   end;
end;

procedure TfrmLD8Q03.btnQueryClick(Sender: TObject);
var i, index, iJisa_15, iJisa_11, iJisa_33, iJisa_34, iJisa_35, iJisa_41 : Integer;
    iJisa_43, iJisa_45, iJisa_46, iJisa_47, iJisa_51, iJisa_52, iJisa_61 : Integer;
    iJisa_63, iJisa_71, iJisa_72, iJisa_kita, iTotal, iStart, iEnd : Integer;
    sCod_hm, sDesc_kor, sGulkwa, gul  : String;
begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid  �ʱ�ȭ
   grdMaster.Rows := 0;
   iJisa_15 := 0;    iJisa_11 := 0;
   iJisa_33 := 0;    iJisa_34 := 0;
   iJisa_35 := 0;    iJisa_41 := 0;
   iJisa_43 := 0;    iJisa_45 := 0;
   iJisa_46 := 0;    iJisa_47 := 0;
   iJisa_51 := 0;    iJisa_52 := 0;
   iJisa_61 := 0;    iJisa_63 := 0;
   iJisa_71 := 0;    iJisa_72 := 0;
   iJisa_kita := 0;  iTotal   := 0;
   index := 0;

   with  qryGulkwa do
   begin
      Active := False;
      ParamByName('St_Date').AsString := mskSt_Date.text;
      ParamByName('ED_Date').AsString := mskED_Date.text;
      ParamByName('Cod_bjjs').AsString := Copy(cmbJisa.Text, 1, 2);
      Active := true;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD8Q03', 'Q', 'Y');
         UP_VarMemSet(RecordCount-1);
         for i := 0 to RecordCount - 1 do
         begin
           iStart  := FieldByname('Pos_start').Asinteger;
           iEnd    := FieldByname('Pos_End').Asinteger;
           if  (sCod_hm <> FieldByName('Cod_hm').AsString) and (i <> 0) then
           begin
              UV_vCod_hm[index]     := sCod_hm;
              UV_vdesc_kor[index]   := sDesc_kor;
              UV_vJisa_15[index]    := IntToStr(iJisa_15);
              UV_vJisa_11[index]    := IntToStr(iJisa_11);
              UV_vJisa_33[index]    := IntToStr(iJisa_33);
              UV_vJisa_34[index]    := IntToStr(iJisa_34);
              UV_vJisa_35[index]    := IntToStr(iJisa_35);
              UV_vJisa_41[index]    := IntToStr(iJisa_45);
              UV_vJisa_43[index]    := IntToStr(iJisa_43);
              UV_vJisa_45[index]    := IntToStr(iJisa_45);
              UV_vJisa_46[index]    := IntToStr(iJisa_46);
              UV_vJisa_47[index]    := IntToStr(iJisa_47);
              UV_vJisa_51[index]    := IntToStr(iJisa_51);
              UV_vJisa_52[index]    := IntToStr(iJisa_52);
              UV_vJisa_61[index]    := IntToStr(iJisa_61);
              UV_vJisa_63[index]    := IntToStr(iJisa_63);
              UV_vJisa_71[index]    := IntToStr(iJisa_71);
              UV_vJisa_72[index]    := IntToStr(iJisa_72);
              UV_vJisa_kita[index]  := IntToStr(iJisa_kita);
              UV_vTotal[index]      := IntToStr(iTotal);
              iJisa_15 := 0;    iJisa_11 := 0;
              iJisa_33 := 0;    iJisa_34 := 0;
              iJisa_35 := 0;    iJisa_41 := 0;
              iJisa_43 := 0;    iJisa_45 := 0;
              iJisa_46 := 0;    iJisa_47 := 0;
              iJisa_51 := 0;    iJisa_52 := 0;
              iJisa_61 := 0;    iJisa_63 := 0;
              iJisa_71 := 0;    iJisa_72 := 0;
              iJisa_kita := 0;  iTotal   := 0;
              Inc(index);
           end;
           gul :=  FieldByName('Desc_glkwa1').AsString +
                   FieldByName('Desc_glkwa2').AsString +
                   FieldByName('Desc_glkwa3').AsString;
           sGulkwa := copy(gul, iStart,  iEnd - iStart + 1);
           if  copy(trim(sGulkwa),1,1) <> '' then
           begin
               if  FieldByName('Cod_jisa').AsString = '15' then
                   iJisa_15 := iJisa_15 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '11' then
                   iJisa_11 := iJisa_11 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '33' then
                   iJisa_33 := iJisa_33 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '34' then
                   iJisa_34 := iJisa_34 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '35' then
                   iJisa_35 := iJisa_35 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '41' then
                   iJisa_41 := iJisa_41 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '43' then
                   iJisa_43 := iJisa_43 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '45' then
                   iJisa_45 := iJisa_45 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '46' then
                   iJisa_46 := iJisa_46 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '47' then
                   iJisa_47 := iJisa_47 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '51' then
                   iJisa_51 := iJisa_51 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '52' then
                   iJisa_52 := iJisa_52 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '61' then
                   iJisa_61 := iJisa_61 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '63' then
                   iJisa_63 := iJisa_63 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '71' then
                   iJisa_71 := iJisa_71 + 1
               else
               if  FieldByName('Cod_jisa').AsString = '72' then
                   iJisa_72 := iJisa_72 + 1
               else
                   iJisa_Kita := iJisa_Kita + 1 ;
               iTotal := iTotal + 1;
            end;
           if  (i = RecordCount - 1 )  then
           begin
              UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
              UV_vdesc_kor[index]   := FieldByName('desc_kor').AsString;
              UV_vJisa_15[index]    := IntToStr(iJisa_15);
              UV_vJisa_11[index]    := IntToStr(iJisa_11);
              UV_vJisa_33[index]    := IntToStr(iJisa_33);
              UV_vJisa_34[index]    := IntToStr(iJisa_34);
              UV_vJisa_35[index]    := IntToStr(iJisa_35);
              UV_vJisa_41[index]    := IntToStr(iJisa_45);
              UV_vJisa_43[index]    := IntToStr(iJisa_43);
              UV_vJisa_45[index]    := IntToStr(iJisa_45);
              UV_vJisa_46[index]    := IntToStr(iJisa_46);
              UV_vJisa_47[index]    := IntToStr(iJisa_47);
              UV_vJisa_51[index]    := IntToStr(iJisa_51);
              UV_vJisa_52[index]    := IntToStr(iJisa_52);
              UV_vJisa_61[index]    := IntToStr(iJisa_61);
              UV_vJisa_63[index]    := IntToStr(iJisa_63);
              UV_vJisa_71[index]    := IntToStr(iJisa_71);
              UV_vJisa_72[index]    := IntToStr(iJisa_72);
              UV_vJisa_kita[index]  := IntToStr(iJisa_kita);
              UV_vTotal[index]      := IntToStr(iTotal);
              Inc(index);
           end;

           sCod_hm :=  FieldByName('Cod_hm').AsString;
           sDesc_kor :=  FieldByName('Desc_kor').AsString;
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

procedure TfrmLD8Q03.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD8Q03.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate  then GF_CalendarClick(mskST_Date)
   else
   if Sender = btnDate1 then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD8Q03.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskST_Date then UP_Click(btnDate)
   else if Sender = mskED_Date then UP_Click(btnDate1)
end;

procedure TfrmLD8Q03.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD8Q031 := TfrmLD8Q031.Create(Self);
   frmLD8Q031.QuickRep.Preview;
//   frmMD3Q311.QuickRep.Print;
end;

procedure TfrmLD8Q03.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

end.
