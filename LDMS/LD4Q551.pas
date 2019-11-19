unit LD4Q551;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q551 = class(TForm)
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
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRLabel3: TQRLabel;
    Qrl_rack: TQRLabel;
    QRLabel4: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q551: TfrmLD4Q551;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q55;

{$R *.DFM}

procedure TfrmLD4Q551.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   if frmLD4Q55.cmbHM.Text = 'S007S099' then
       QRLabel37.caption :='HBsAg'
   else if frmLD4Q55.cmbHM.Text = 'S016' then
       QRLabel37.caption :='HCV Ab'
   else if frmLD4Q55.cmbHM.Text = 'S021' then
       QRLabel37.caption :='H.pylori'
   else if frmLD4Q55.cmbHM.Text = 'S012' then
       QRLabel37.caption :='HBe Ag'
   else if frmLD4Q55.cmbHM.Text = 'S011' then
       QRLabel37.caption :='HBe Ab'
   else if frmLD4Q55.cmbHM.Text = 'S010' then
       QRLabel37.caption :='HBc Ab';

   with frmLD4Q55.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption      := Cell[ 1,UV_iCount];
      Qrl_rack.Caption    := Cell[ 2,UV_iCount];
      QRL_Sample.Caption  := Cell[ 3,UV_iCount];
      QRL_Bunju.Caption   := Cell[ 4,UV_iCount];
      QRL_Name.Caption    := Cell[ 5,UV_iCount];
      QRL_S007.Caption    := Cell[ 6,UV_iCount];

//      if (UV_iCount mod 5) = 0 then QRShape13.Pen.Width := 5
//      else QRShape13.Pen.Width := 2;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q551.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q55.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q55.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q55.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
