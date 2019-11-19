unit LD4Q421;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q421 = class(TForm)
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
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_SEX: TQRLabel;
    QRShape18: TQRShape;
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
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    lavel234: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_NAME: TQRLabel;
    QRShape8: TQRShape;
    QRShape3: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Samp: TQRLabel;
    QRLabel3: TQRLabel;
    QRL_PSA: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_CA125: TQRLabel;
    QRLabel7: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow : Integer;
    sum_PSA, sum_CA125 : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q421: TfrmLD4Q421;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q42;

{$R *.DFM}

procedure TfrmLD4Q421.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD4Q42.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption    := Cell[ 1,UV_iCount];
      QRL_Samp.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      QRL_SEX.Caption   := Cell[ 4,UV_iCount];
      QRL_NAME.Caption  := Cell[ 5,UV_iCount];

      QRL_PSA.Caption   := Cell[ 6,UV_iCount];
      QRL_CA125.Caption := Cell[ 7,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q421.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q42.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q42.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q42.mskDate.Text,7,2) + '일';
   UV_iCount := 0;
   Sum_PSA := 0; Sum_CA125 := 0;
end;

end.
