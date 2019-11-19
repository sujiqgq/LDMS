unit LD4Q691;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q691 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_hm: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRLabel24: TQRLabel;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape18: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_rack: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape36: TQRShape;
    QRLabel37: TQRLabel;
    QRShape40: TQRShape;
    QRL_gulkwa: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Sample: TQRLabel;
    QRShape19: TQRShape;
    QRShape2: TQRShape;
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
  frmLD4Q691: TfrmLD4Q691;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q69;

{$R *.DFM}

procedure TfrmLD4Q691.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q69.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      qrl_Sample.Caption := Cell[ 1,UV_iCount];
      qrl_Name.Caption   := Cell[ 2,UV_iCount];
      qrl_Rack.Caption   := Cell[ 3,UV_iCount];
      qrl_HM.Caption     := Cell[ 4,UV_iCount];
      qrl_Gulkwa.Caption := Cell[ 5,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q691.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q69.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q69.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q69.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
