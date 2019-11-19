unit LD4Q811;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, StdCtrls;

type
  TfrmLD4Q811 = class(TForm)
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
    QRL_Rack: TQRLabel;
    QRLabel21: TQRLabel;
    QRL_Numsamp: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date_start: TQRLabel;
    QRShape36: TQRShape;
    QRShape40: TQRShape;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_hm: TQRLabel;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRLabel4: TQRLabel;
    QRL_dat_bunju: TQRLabel;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRLabel6: TQRLabel;
    QRL_No: TQRLabel;
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
  frmLD4Q811: TfrmLD4Q811;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q81;

{$R *.DFM}

procedure TfrmLD4Q811.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q81.grdMaster do
   begin
      QRLabel2.Caption  := Col[7].Heading;
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;
      QRL_No.Caption        := Cell[ 1,UV_iCount];
      QRL_dat_bunju.Caption := Cell[ 2,UV_iCount];
      QRL_Rack.Caption      := Cell[ 3,UV_iCount];
      QRL_Numsamp.Caption   := Cell[ 4,UV_iCount];
      QRL_Bunju.Caption     := Cell[ 5,UV_iCount];
      QRL_Name.Caption      := Cell[ 6,UV_iCount];
      QRL_hm.Caption        := Cell[ 7,UV_iCount];

      if (UV_iCount mod 5) = 0 then
         QRShape13.Pen.Width := 8
      else QRShape13.Pen.Width := 2;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q811.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date_start.Caption := copy(frmLD4Q81.mskDate_start.Text,1,4) + '년'  +
                             copy(frmLD4Q81.mskDate_Start.Text,5,2) + '월'  +
                             copy(frmLD4Q81.mskDate_Start.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
