unit LD4Q171;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q171 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape32: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel4: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_S016: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape14: TQRShape;
    QRShape3: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRLabel2: TQRLabel;
    QRL_S033: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRL_Samp: TQRLabel;
    QRL_Index: TQRLabel;
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
  frmLD4Q171: TfrmLD4Q171;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q17;

{$R *.DFM}

procedure TfrmLD4Q171.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD4Q17.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // ����ڷ� ����
      UV_iCount := UV_iCount + 1;
      QRL_Index.Caption := cell[ 1,UV_iCount];
      QRL_samp.Caption  := cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      QRL_No.Caption    := Cell[ 4,UV_iCount];
      QRL_Name.Caption  := Cell[ 5,UV_iCount];
      QRL_S016.Caption  := Cell[ 6,UV_iCount];
      QRL_S033.Caption  := Cell[ 7,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q171.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q17.edtBDate.Text,1,4) + '��'  +
                       copy(frmLD4Q17.edtBDate.Text,5,2) + '��'  +
                       copy(frmLD4Q17.edtBDate.Text,7,2) + '��';
   UV_iCount := 0;
end;

end.
