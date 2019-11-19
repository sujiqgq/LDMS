//==============================================================================
// ���α׷� ���� : ����List ���
// �ý���        : ���հ���
// ��������      : 1999.08.09
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
unit LD2I031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD2I031 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_Date: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRS_Page: TQRSysData;
    QRLabel12: TQRLabel;
    qrlFrom: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel5: TQRLabel;
    QRBand2: TQRBand;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRL5: TQRLabel;
    QRL6: TQRLabel;
    QRL8: TQRLabel;
    QRL9: TQRLabel;
    etc: TQRLabel;
    QRLabel17: TQRLabel;
    QRL7: TQRLabel;
    QRShape1: TQRShape;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD2I031: TfrmLD2I031;

implementation

{$R *.DFM}

uses ZZfunc, LD2I03;

procedure TfrmLD2I031.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2I03.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // ����ڷ� ����
      QRL1.Caption := Cell[1, UV_iCount];
      QRL2.Caption := Cell[2, UV_iCount];
      QRL3.Caption := Cell[3, UV_iCount];
      QRL4.Caption := Cell[4, UV_iCount];
      QRL5.Caption := Cell[5, UV_iCount];
      QRL6.Caption := Cell[6, UV_iCount];
      QRL7.Caption := Cell[7, UV_iCount];
      QRL8.Caption := Cell[8, UV_iCount];
      QRL9.Caption := Cell[9, UV_iCount];
      etc.Caption := '';
      
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD2I031.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // ��ȸ������ ����
   with frmLD2I03 do
   begin
      QRLabel1.Caption := '��ü�ҷ����� ����Ʈ[' + CbB_GUBN.Text + ']';
      qrlFrom.Caption  := GF_DateFormat(mskDate.Text);
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

end.
