//==============================================================================
// ���α׷� ���� : ��ü�ҷ�����
// �ý���        : ���հ���
// ��������      : 2012.03.02
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
// ��������      : 2012.03.05
// ������        : ������
//[2012.03.15] ��ü�̳��� ������ ó���ϰ� �ٲ�..
//==============================================================================
// ��������      : 2012.06.21
// ������        : ����
//��������       : ���� ���� �߰�
//==============================================================================
// ��������      : 2013.12.19
// ������        : ����
//��������       : ���� �̰� üũ
//==============================================================================

unit LD2I03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit;

type
  TfrmLD2I03 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryI_Gum_bul: TQuery;
    qryD_Gum_bul: TQuery;
    pnlDetail: TPanel;
    edtNum_bunju: TEdit;
    edtNUM_TEL: TEdit;
    edtCod_jisa: TEdit;
    edtDesc_name: TEdit;
    edtCOD_JANGBI: TEdit;
    edtCod_dc: TEdit;
    edtDESC_DC: TEdit;
    mskNum_jumin: TMaskEdit;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    edtDesc_dept: TEdit;
    pnlCond: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel10: TPanel;
    Panel16: TPanel;
    Label1: TLabel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    Label2: TLabel;
    rbBunju: TRadioButton;
    rbGum_bul: TRadioButton;
    qryBunju: TQuery;
    qrySubGum_bul: TQuery;
    mskDat_bunju: TDateEdit;
    Panel29: TPanel;
    Panel18: TPanel;
    Panel9: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel17: TPanel;
    ckGUBN_JUNGDO1: TCheckBox;
    ckGUBN_JUNGDO2: TCheckBox;
    ckGUBN_JUNGDO3: TCheckBox;
    ckGUBN_JUNGDO4: TCheckBox;
    ckGUBN_JUNGDO5: TCheckBox;
    mskDAT_SOLVE1: TDateEdit;
    mskDAT_SOLVE2: TDateEdit;
    mskDAT_SOLVE3: TDateEdit;
    mskDAT_SOLVE4: TDateEdit;
    mskDAT_SOLVE5: TDateEdit;
    Panel19: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    ckGUBN_JUNGDO6: TCheckBox;
    ckGUBN_JUNGDO7: TCheckBox;
    ckGUBN_JUNGDO8: TCheckBox;
    ckGUBN_JUNGDO9: TCheckBox;
    mskDAT_SOLVE6: TDateEdit;
    mskDAT_SOLVE7: TDateEdit;
    mskDAT_SOLVE8: TDateEdit;
    mskDAT_SOLVE9: TDateEdit;
    edtDesc_jisa: TEdit;
    Bevel2: TBevel;
    Panel23: TPanel;
    edtCod_bjjs: TEdit;
    Panel24: TPanel;
    edtGubn_hulgum: TEdit;
    Edit3: TEdit;
    Bevel3: TBevel;
    edtCOD_HUL: TEdit;
    Edit5: TEdit;
    edtCOD_CANCER: TEdit;
    Edit7: TEdit;
    edtCod_chuga: TEdit;
    Panel15: TPanel;
    Panel25: TPanel;
    Panel26: TPanel;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel30: TPanel;
    Panel31: TPanel;
    Panel32: TPanel;
    Panel33: TPanel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    Bevel8: TBevel;
    Bevel9: TBevel;
    Label3: TLabel;
    cmbCOD_JISA: TComboBox;
    Cmb_gubn: TComboBox;
    Label14: TLabel;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    Label10: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label15: TLabel;
    MEdt_SampS: TMaskEdit;
    CbB_GUBN: TComboBox;
    Panel34: TPanel;
    Label5: TLabel;
    RD_print: TRadioButton;
    RD_preview: TRadioButton;
    JD_sayu: TEdit;
    Panel35: TPanel;
    Panel36: TPanel;
    edtGum_Check: TEdit;
    edt_temp: TEdit;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Exit(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ ������
    UV_vCod_bjjs, UV_vNum_id, UV_vDat_bunju, UV_vNum_bunju, UV_vCod_jisa, UV_vDat_gmgn,
    UV_vGubn_hulgum, UV_vDesc_dept, UV_vNum_tel, UV_vDesc_name, UV_vNum_jumin, UV_vCod_dc,
    UV_vCod_hul, UV_vCod_jangbi, UV_vCod_cancer, UV_vCod_chuga, UV_vDesc_dc,UV_vGUBN_JUNGDO,UV_vGum_Urin,UV_vGum_Blood: Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
  public
    { Public declarations }

    // SQL �ӽú���
    UV_sBasicSQL : String;

    // �����ڵ�
    UV_sJisa : String;

  end;

var
  frmLD2I03: TfrmLD2I03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2I031;

{$R *.DFM}

procedure TfrmLD2I03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_hulgum := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dept   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_tel     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_dc      := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc     := VarArrayCreate([0, 0], varOleStr);
      UV_vGUBN_JUNGDO := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Urin        := VarArrayCreate([0, 0], varOleStr);//���� �Һ��̰�
      UV_vGum_Blood       := VarArrayCreate([0, 0], varOleStr);//���� ���׹̰�
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vDat_bunju,   iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vGubn_hulgum, iLength);
      VarArrayRedim(UV_vDesc_dept,   iLength);
      VarArrayRedim(UV_vNum_tel,     iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vCod_dc,      iLength);
      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_cancer,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vDesc_dc,     iLength);
      VarArrayRedim(UV_vGUBN_JUNGDO, iLength);
      VarArrayRedim(UV_vGum_Urin,    iLength);
      VarArrayRedim(UV_vGum_Blood,   iLength);
   end;
end;

function TfrmLD2I03.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

end;

function TfrmLD2I03.UF_Save : Boolean;
var i, index : Integer;
begin
   Result := False;

   // �۾������� Check
   if not UV_bEdit then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   index := grdMaster.CurrentDataRow - 1;
   // DB �۾�
   dmComm.database.StartTransaction;
   try
      if qrySubGum_bul.RecordCount > 0 then
      begin
         with qryD_Gum_bul do
         begin
            ParamByName('COD_BJJS').AsString    := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString      := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString   := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString   := UV_vNum_bunju[index];

            Execsql;
            GP_query_log(qryD_Gum_bul, 'LD2I03', 'Q', 'Y');
         end;
      end;
      with qryI_Gum_bul do
      begin
         if ckGUBN_JUNGDO1.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '01';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE1.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE1.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
            GP_query_log(qryI_Gum_bul, 'LD2I03', 'Q', 'Y');
         end;
         if ckGUBN_JUNGDO2.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '02';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE2.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE2.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
         if ckGUBN_JUNGDO3.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '03';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE3.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE3.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
         if ckGUBN_JUNGDO4.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '04';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE4.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE4.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
         if ckGUBN_JUNGDO5.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '05';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE5.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE5.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
         if ckGUBN_JUNGDO6.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '06';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE6.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE6.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
{
//[2012.03.15] ��ü�̳��� ������ ó���ϰ� �ٲ�..
         if ckGUBN_JUNGDO7.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '07';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE7.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE7.Text;
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
}
         if ckGUBN_JUNGDO8.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '08';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE8.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE8.Text;

            
            ParamByName('JD_sayu').AsString     := JD_sayu.text;
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
         if ckGUBN_JUNGDO9.Checked = True then
         begin
            ParamByName('COD_BJJS').AsString     := UV_vCod_bjjs[index];
            ParamByName('NUM_ID').AsString       := UV_vNum_id[index];
            ParamByName('DAT_BUNJU').AsString    := UV_vDat_bunju[index];
            ParamByName('NUM_BUNJU').AsString    := UV_vNum_bunju[index];
            ParamByName('GUBN_BUL').AsString     := '09';
            ParamByName('GUBN_JUNGDO').AsString  := 'Y';
            ParamByName('DAT_GMGN').AsString     := UV_vDat_gmgn[index];
            if mskDAT_SOLVE9.Text = '' then ParamByName('DAT_SOLVE').AsString := GV_sToday
            else                            ParamByName('DAT_SOLVE').AsString := mskDAT_SOLVE9.Text;
            ParamByName('JD_sayu').AsString     := '';
            ParamByName('COD_JISA').AsString     := UV_vCod_jisa[index];
            ParamByName('DAT_INSERT').AsString   := GV_sToday;
            ParamByName('COD_INSERT').AsString   := GV_sUserId;

            Execsql;
         end;
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;

   dmComm.database.Commit;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD2I03.btnSaveClick(Sender: TObject);
begin
   inherited;

   UF_Save;
end;

procedure TfrmLD2I03.FormCreate(Sender: TObject);
begin
   inherited;
   mskDate.Text := dateTostr(date);

   // ���� Setting
   if GV_sJisa = '00' then UV_sJisa := Copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   if UV_sJisa = '15' then
   begin
      Label3.Caption := '��    ��:';
      GP_ComboJisa(cmbCOD_JISA);
      cmbCod_jisa.ItemIndex := 0;
   end
   else
   begin
      Label3.Caption := '���ּ���:';
      cmbCod_jisa.Items.Add('�� �� ��');
      cmbCod_jisa.Items.Add('��    ��');
      cmbCod_jisa.ItemIndex := 0;
   end;

   // ȯ�漳��
   UP_VarMemSet(0);

   CbB_GUBN.ItemIndex := 0;
end;


procedure TfrmLD2I03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
var sName : String;
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vCod_bjjs[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
      4 : Value := UV_vDesc_dept[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : Value := Copy(UV_vNum_jumin[DataRow-1], 1, 6) + '-*******';
      7 : Value := '[' +  UV_vCod_dc[DataRow-1] + ']' + UV_vDesc_dc[DataRow-1];
      8 : begin
             case UV_vGubn_hulgum[DataRow-1] of
                1 : Value := '[1]������';
                2 : Value := '[2]����';
                3 : Value := '[3]������ + ����';
             end;
          end;
      9 : begin
             case StrToInt(UV_vGUBN_JUNGDO[DataRow-1]) of
                1 : Value := '��ü ȥŹ';
                2 : Value := '��ü ����';
                3 : Value := 'CBC CLOT';
                4 : Value := '��û �̳�';
                5 : Value := '���� �̳�';
                6 : Value := '�Һ� �̳�';
                7 : Value := '��ü �̳�';
                8 : Value := '��       ��';
                9 : Value := '��  ä  ��';
             else
                Value := '';
             end;
         end;
   end;
end;

procedure TfrmLD2I03.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy, sTemp, sSelect, sName : String;
begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   index := 0;

   // ���ֹ�ȣ�� 0�� ä����
   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   // Query ���� & �迭�� ����
   with qryBunju do
   begin
      // ������ �����Ƿ� �ٷ� Open��Ŵ
      Active := False;

      if rbBunju.Checked = True then
      begin
         // SQL���� ����
         sSelect := 'SELECT A.cod_bjjs,  A.num_id,     A.dat_bunju,  A.num_bunju, A.dat_gmgn,    A.cod_jisa, ' +
                    '       B.desc_name, B.num_jumin,  B.num_tel,    B.cod_dc,    B.Gubn_hulgum, B.num_samp, ' +
                    '       B.Cod_hul,   B.Cod_jangbi, B.Cod_cancer, B.Cod_chuga, B.Desc_dept, C.GUBN_BUL, C.GUBN_JUNGDO,K.CK_Urin,CK_blood ' +
                    'FROM   bunju A INNER JOIN gumgin  B ON  A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ' +
                    '          LEFT OUTER JOIN gum_bul C ON  A.cod_jisa = C.cod_jisa AND A.num_id = C.num_id AND A.dat_gmgn = C.dat_gmgn '+
                    '          LEFT OUTER JOIN gumgin_Check K ON  A.cod_jisa = K.cod_jisa AND A.num_id = K.num_id AND A.dat_gmgn = K.dat_gmgn ';

        sOrderBy := ' GROUP BY A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.dat_gmgn, A.cod_jisa, ' +
                    ' B.desc_name, B.num_jumin, B.num_tel, B.cod_dc, B.Gubn_hulgum, B.num_samp, B.Cod_hul, ' +
                    ' B.Cod_jangbi, B.Cod_cancer, B.Cod_chuga, B.Desc_dept, C.GUBN_BUL, C.GUBN_JUNGDO,K.CK_Urin,CK_blood ';
      end
      else
      begin
         // SQL���� ����
         sSelect := 'SELECT A.cod_bjjs,  A.num_id,     A.dat_bunju,  A.num_bunju, A.dat_gmgn,    A.cod_jisa, A.GUBN_BUL, A.GUBN_JUNGDO, ' +
                    '       B.desc_name, B.num_jumin,  B.num_tel,    B.cod_dc,    B.Gubn_hulgum, B.num_samp, ' +
                    '       B.Cod_hul,   B.Cod_jangbi, B.Cod_cancer, B.Cod_chuga, B.Desc_dept, K.CK_Urin,CK_blood  ' +
                    'FROM   gum_bul A INNER JOIN gumgin B  ON  A.cod_jisa = B.cod_jisa  AND A.num_id   = B.num_id  AND A.dat_gmgn = B.dat_gmgn '+
                    '       LEFT OUTER JOIN gumgin_Check K ON  A.cod_jisa = K.cod_jisa AND A.num_id = K.num_id AND A.dat_gmgn = K.dat_gmgn ';
                    
        sOrderBy := ' GROUP BY A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.dat_gmgn, A.cod_jisa, A.GUBN_BUL, A.GUBN_JUNGDO, ' +
                    ' B.desc_name, B.num_jumin, B.num_tel, B.cod_dc, B.Gubn_hulgum, B.num_samp, B.Cod_hul,K.CK_Urin,CK_blood, ' +
                    ' B.Cod_jangbi, B.Cod_cancer, B.Cod_chuga, B.Desc_dept ';
      end;

      if Label3.Caption = '��    ��:' then sWhere := ' WHERE  A.cod_bjjs = ''' + Trim(copy(cmbCOD_JISA.Text,1,2)) + ''''
      else
      begin
         sWhere := ' WHERE  A.cod_bjjs IN (''15'',''' + UV_sJisa + ''')';
         sWhere := sWhere + ' AND  A.cod_jisa = ''' + UV_sJisa + '''';
      end;

      sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';

      if Cmb_gubn.Text = '���ֹ�ȣ' then
      begin
         if Trim(mskFrom.Text) <> '' then  sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then    sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
      end
      else if Cmb_gubn.Text = '���ù�ȣ' then
      begin
         if Trim(MEdt_SampS.Text) <> '' then sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then  sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then    sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
      end;

      if rbGum_bul.Checked then
      begin
         case CbB_GUBN.ItemIndex of
            1 : sWhere := sWhere + ' AND A.gubn_bul in (''01'')';
            2 : sWhere := sWhere + ' AND A.gubn_bul in (''02'')';
            3 : sWhere := sWhere + ' AND A.gubn_bul in (''03'')';
            4 : sWhere := sWhere + ' AND A.gubn_bul in (''04'',''07'')';
            5 : sWhere := sWhere + ' AND A.gubn_bul in (''05'',''07'')';
            6 : sWhere := sWhere + ' AND A.gubn_bul in (''06'',''07'')';
            7 : begin
                   sWhere := sWhere + ' AND A.gubn_bul = ''04''';
                   sWhere := sWhere + ' AND A.gubn_bul = ''05''';
                   sWhere := sWhere + ' AND A.gubn_bul = ''06''';
                end;
            8 : sWhere := sWhere + ' AND A.gubn_bul in (''08'')';
            9 : sWhere := sWhere + ' AND A.gubn_bul in (''09'')';
         end;
      end;
      
      if Cmb_gubn.Text = '���ֹ�ȣ' then
         sOrderBy := sOrderBy + ' ORDER BY A.cod_bjjs, A.cod_jisa, A.dat_bunju, A.num_bunju '
      else if Cmb_gubn.Text = '���ù�ȣ' then
         sOrderBy := sOrderBy + ' ORDER BY A.cod_bjjs, A.cod_jisa, A.dat_bunju, CAST(B.num_samp AS INT), A.num_bunju '
      else
         sOrderBy := sOrderBy + ' ORDER BY A.cod_bjjs, A.cod_jisa, A.dat_bunju, A.num_bunju ';


      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2I03', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);

         for i := 0 to RecordCount - 1 do
         begin
            UV_vCod_bjjs[index]    := FieldByName('COD_BJJS').AsString;
            UV_vNum_id[index]      := FieldByName('NUM_ID').AsString;
            UV_vDat_bunju[index]   := FieldByName('DAT_BUNJU').AsString;
            UV_vNum_bunju[index]   := FieldByName('NUM_BUNJU').AsString;
            UV_vCod_jisa[index]    := FieldByName('COD_JISA').AsString;
            UV_vDat_gmgn[index]    := FieldByName('DAT_GMGN').AsString;
            UV_vGubn_hulgum[index] := FieldByName('Gubn_hulgum').AsString;
            UV_vDesc_dept[index]   := FieldByName('num_samp').AsString;
            UV_vNum_tel[index]     := FieldByName('NUM_TEL').AsString;
            UV_vDesc_name[index]   := FieldByName('DESC_NAME').AsString;
            UV_vNum_jumin[index]   := FieldByName('NUM_JUMIN').AsString;
            UV_vCod_dc[index]      := FieldByName('COD_DC').AsString;
            UV_vCod_hul[index]     := FieldByName('COD_HUL').AsString;
            UV_vCod_jangbi[index]  := FieldByName('COD_JANGBI').AsString;
            UV_vCod_cancer[index]  := FieldByName('COD_CANCER').AsString;
            UV_vCod_chuga[index]   := FieldByName('Cod_chuga').AsString;
            UV_vGum_Urin[index]        := FieldByName('CK_urin').AsString;
            UV_vGum_Blood[index]       := FieldByName('CK_blood').AsString;


            if FieldByName('GUBN_JUNGDO').AsString = 'Y' then UV_vGUBN_JUNGDO[index] := FieldByName('GUBN_BUL').AsString
            else                                              UV_vGUBN_JUNGDO[index] := '0';

            edt_temp.Text := FieldByName('COD_DC').AsString;
            if GF_DancheEdit(edt_temp, sName) then UV_vDesc_dc[index] := sName;

            Next;
            Inc(index);
         end;

         // Table�� Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;
   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD2I03.btnCancelClick(Sender: TObject);
begin
   inherited;

   // �۾������� Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // ��ҽ�
   // 1.�Է��̸� ȭ�� Clear
   // 2.�����̸� Re-query
   if UV_iStatus = GC_Insert then
      GP_FieldClear(pnlDetail);

   // �ڷ� Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD2I03.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;

   // Edit Change�� Event Procedure�� ������ �� ������ Sender�� ����
   if (Sender = ckGUBN_JUNGDO7)then
   begin
      if ckGUBN_JUNGDO7.Checked then
      begin
         ckGUBN_JUNGDO4.Checked := True;
         ckGUBN_JUNGDO5.Checked := True;
         ckGUBN_JUNGDO6.Checked := True;
     end
     else
      begin
         ckGUBN_JUNGDO4.Checked := False;
         ckGUBN_JUNGDO5.Checked := False;
         ckGUBN_JUNGDO6.Checked := False;
     end;
   end;

   if  ckGUBN_JUNGDO8.Checked  then
       JD_sayu.enabled:= true
   else JD_sayu.enabled:= false;


   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD2I03.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var sName : String;
begin
   inherited;

   GP_FieldClear(pnlDetail);

   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   // Grid�� Row�� ����Ǹ� Detail�� �ڷḦ �Ҵ�
   edtCod_bjjs.Text  := UV_vCod_bjjs[NewRow-1];
   mskDat_bunju.Text := UV_vDat_bunju[NewRow-1];
   edtNum_bunju.Text := UV_vNum_bunju[NewRow-1];

   edtCod_jisa.Text  := UV_vCod_jisa[NewRow-1];
   edtDesc_name.Text := UV_vDesc_name[NewRow-1];
   mskNum_jumin.Text := UV_vNum_jumin[NewRow-1];
   edtNUM_TEL.Text   := UV_vNum_tel[NewRow-1];
   edtDesc_dept.Text := UV_vDesc_dept[NewRow-1];
   
   case UV_vGubn_hulgum[NewRow-1] of
      1 : edtGubn_hulgum.Text := '[1]������';
      2 : edtGubn_hulgum.Text := '[2]����';
      3 : edtGubn_hulgum.Text := '[3]������ + ����';
   end;

   edtCod_dc.Text      := UV_vCod_dc[NewRow-1];
   edtCOD_HUL.Text     := UV_vCod_hul[NewRow-1];
   edtCOD_JANGBI.Text  := UV_vCod_jangbi[NewRow-1];
   edtCOD_CANCER.Text  := UV_vCod_cancer[NewRow-1];
   edtCod_chuga.Text   := UV_vCod_chuga[NewRow-1];

    //�����̰� üũ
   if UV_vGum_Urin[NewRow-1] ='Y'  then
      edtGum_Check.Text:=' �Һ��̰�';
   if UV_vGum_Blood[NewRow-1] ='Y' then
      edtGum_Check.Text:=edtGum_Check.Text+' ���׹̰�';

   if GF_JisaEdit(edtCOD_JISA, sName) then edtDESC_JISA.Text := sName;
   if GF_DancheEdit(edtCOD_DC, sName) then edtDESC_DC.Text   := sName;

   with qrySubGum_bul do
   begin
      Active := False;
      ParamByName('cod_bjjs').AsString := UV_vCod_bjjs[NewRow-1];
      ParamByName('dat_bunju').AsString := UV_vDat_bunju[NewRow-1];
      ParamByName('num_bunju').AsString := UV_vnum_bunju[NewRow-1];
      Active := True;
      if Recordcount > 0 then
      begin
         while not EOF do
         begin
            if FieldByName('gubn_bul').AsString = '01' then
            begin
               ckGUBN_JUNGDO1.Checked := True;
               mskDAT_SOLVE1.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '02' then
            begin
               ckGUBN_JUNGDO2.Checked := True;
               mskDAT_SOLVE2.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '03' then
            begin
               ckGUBN_JUNGDO3.Checked := True;
               mskDAT_SOLVE3.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '04' then
            begin
               ckGUBN_JUNGDO4.Checked := True;
               mskDAT_SOLVE4.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '05' then
            begin
               ckGUBN_JUNGDO5.Checked := True;
               mskDAT_SOLVE5.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '06' then
            begin
               ckGUBN_JUNGDO6.Checked := True;
               mskDAT_SOLVE6.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '07' then
            begin
               ckGUBN_JUNGDO7.Checked := True;
               mskDAT_SOLVE7.Text := FieldByName('dat_solve').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '08' then
            begin
               ckGUBN_JUNGDO8.Checked := True;
               mskDAT_SOLVE8.Text := FieldByName('dat_solve').AsString;
               JD_sayu.Text := FieldByName('JD_sayu').AsString;
            end;
            if FieldByName('gubn_bul').AsString = '09' then
            begin
               ckGUBN_JUNGDO9.Checked := True;
               mskDAT_SOLVE9.Text := FieldByName('dat_solve').AsString;
            end;

            Next;
         end;

      end;
   end;

   // Grid Focus
   grdMaster.SetFocus;

   // Data�Ǽ� ǥ��
   GP_SetDataCnt(grdMaster);

   // ��ȸ Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD2I03.UP_Exit(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

end;

procedure TfrmLD2I03.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;
   if      Sender = mskDate then UP_Click(mskDate);

end;

procedure TfrmLD2I03.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate)
   else if Sender = rbGum_bul then CbB_GUBN.Enabled := True
   else if Sender = rbBunju   then CbB_GUBN.Enabled := False;
end;

procedure TfrmLD2I03.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD2I031 := TfrmLD2I031.Create(Self);
   if      RD_print.Checked   then frmLD2I031.QuickRep.Print
   else if RD_preview.Checked then frmLD2I031.QuickRep.Preview;
end;

end.
