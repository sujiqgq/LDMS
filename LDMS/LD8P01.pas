//==============================================================================
// ���α׷� ���� : ���װ�����
// �ý���        : ���հ���
// ��������      : 1999.10.21
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
unit LD8P01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD8P01 = class(TfrmPrint)
    cmbJisa: TComboBox;
    mskSt_date: TDateEdit;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    rbGubn_part: TRadioGroup;
    CB_Hulgum: TCheckBox;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure UP_ReportClick(Sender: TObject);

  private
    { Private declarations }
     public


    // �����ڵ�
    UV_sCod_jisa : String;
    // SQL�� ����
   UV_sBasicSQL : String;

// Procedure UP_gulkwa;
 //procedure UP_Select;
    { Public declarations }
  end;

var
  frmLD8P01: TfrmLD8P01;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,
     LD8P011;

procedure TfrmLD8P01.UP_Click(Sender: TObject);
var  sCode, sName : string;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date);

end;

procedure TfrmLD8P01.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate)
end;

procedure TfrmLD8P01.FormCreate(Sender: TObject);
var sJisa, sName : String;
begin
   inherited;

   GP_ComboJisa(cmbJisa);

   // ����ڱ��ѿ� ���� ����Combo Ȱ��ȭ ����
   if GV_sJisa = '00' then
   begin
      cmbJisa.Enabled := True;
      sJisa := copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbJisa.Enabled := False;
      sJisa := GV_sJisa;
   end;

   // Default�� �������縦 ����
   GP_ComboDisplay(cmbJisa, sJisa, 2);


   // �������� ����
   mskST_Date.Text := GV_sToday;
   Cmb_gubn.ItemIndex := 0;
end;

procedure TfrmLD8P01.UP_ReportClick(Sender: TObject);
var
   sWhere : string;
   index, i  : integer;
begin
   inherited;

   // ��ȸ���� Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD8P011 := TfrmLD8P011.Create(Self);

  with frmLD8P011.QuickRep do
   begin
      // ��ȸ���� & Date �Ҵ�
      frmLD8P011.QRL_Date.Caption   := GV_sPrintToday;
      frmLD8P011.QRLfrom.Caption    := mskST_Date.Text;
       if  rbGubn_part.itemindex = 0 then
           frmLD8P011.QRL_Title.Caption := '�� ȭ ��'
       else if  rbGubn_part.itemindex = 1 then
           frmLD8P011.QRL_Title.Caption := '�� �� ��'
       else if  rbGubn_part.itemindex = 2 then
           frmLD8P011.QRL_Title.Caption := 'U R I N '
       else if  rbGubn_part.itemindex = 3 then
           frmLD8P011.QRL_Title.Caption := 'R  I  A '
       else if  rbGubn_part.itemindex = 4 then
           frmLD8P011.QRL_Title.Caption := 'E  I  A '
       else if  rbGubn_part.itemindex = 5 then
           frmLD8P011.QRL_Title.Caption := '�� û ��';

  // Query ����
//    Up_Select;

  // �μ���� ����

      if rbAll.Checked then
      begin
         PrinterSettings.FirstPage := 1;
         PrinterSettings.LastPage := 1000;
      end
      else
      begin
         PrinterSettings.FirstPage := Trunc(valFrom.Value);
         PrinterSettings.LastPage := Trunc(valTo.Value);
      end;

      // �μ�ż� ����
      PrinterSettings.Copies := spnCopy.Value;

      // Preview or Print
      if Sender = btnPreview then Preview
      else if Sender = btnPrint then Print;
   end;
end;


end.
