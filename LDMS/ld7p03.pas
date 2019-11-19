//==============================================================================
// ���α׷� ���� : 2015�⿬�������״���(�ݿ�,ȣ����)
// �ý���        : ���հ���
// ��������      : 2015.04.06
// ������        : ��â��
// ��������      : �߰�..
// �������      :
//==============================================================================
unit LD7P03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD7P03 = class(TfrmPrint)
    cmbJisa: TComboBox;
    mskSt_date: TDateEdit;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    qrySaler: TQuery;
    Query1: TQuery;
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
  frmLD7P03: TfrmLD7P03;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,
     LD7P031;

procedure TfrmLD7P03.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date);

end;

procedure TfrmLD7P03.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate)
end;

procedure TfrmLD7P03.FormCreate(Sender: TObject);
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
end;

procedure TfrmLD7P03.UP_ReportClick(Sender: TObject);
var
   sWhere : string;
   index, i  : integer;
begin
   inherited;

   // ��ȸ���� Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD7P031 := TfrmLD7P031.Create(Self);

   with frmLD7P031.qryGulkwa do
   begin
      Active := False;
      ParambyName('dat_gmgn').AsString := mskSt_date.Text;
      Active := True;
      if RecordCount > 0 then
      begin
         GP_query_log(frmLD7P031.qryGulkwa, 'LD7P03', 'Q', 'Y');
         if rbAll.Checked then
         begin
            frmLD7P031.QuickRep.PrinterSettings.FirstPage := 0;
            frmLD7P031.QuickRep.PrinterSettings.LastPage  := 0;
         end
         else if rbPart.Checked then
         begin
            frmLD7P031.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
            frmLD7P031.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
         end;
         // �μ�ż� ����
         frmLD7P031.QuickRep.PrinterSettings.Copies := spnCopy.Value;

      // Preview or Print
      if Sender = btnPreview then frmLD7P031.QuickRep.Preview
      else if Sender = btnPrint then frmLD7P031.QuickRep.Print;

      end;
  end; // end of with;
  
end;


end.
