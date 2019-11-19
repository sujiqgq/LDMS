unit LD4Q681;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q681 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRLabel37: TQRLabel;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRL_S007: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Sample: TQRLabel;
    QRLabel4: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q681: TfrmLD4Q681;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q68;

{$R *.DFM}

procedure TfrmLD4Q681.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   if frmLD4Q68.CB01.Text <> '����' then
      QRLabel37.caption := frmLD4Q68.CB01.Text
   else
      QRLabel37.caption := frmLD4Q68.CB01.Text + '-' +
                           copy(frmLD4Q68.CB_Sub.Text, 3, length(frmLD4Q68.CB_Sub.Text));

   with frmLD4Q68.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // ����ڷ� ����
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption      := Cell[ 1,UV_iCount];
      QRL_Sample.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption   := Cell[ 3,UV_iCount];
      QRL_Name.Caption    := Cell[ 4,UV_iCount];
      QRL_S007.Caption    := Cell[ 5,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q681.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q68.EdtBDate.Text,1,4) + '��'  +
                       copy(frmLD4Q68.EdtBDate.Text,5,2) + '��'  +
                       copy(frmLD4Q68.EdtBDate.Text,7,2) + '��';
   UV_iCount := 0;
end;

end.
