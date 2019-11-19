unit LD4Q831;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q831 = class(TForm)
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
    QRShape1: TQRShape;
    lavel234: TQRLabel;
    QRShape7: TQRShape;
    QRL_NAME: TQRLabel;
    QRShape8: TQRShape;
    QRShape3: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Samp: TQRLabel;
    QRLabel7: TQRLabel;
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
  frmLD4Q831: TfrmLD4Q831;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q83;

{$R *.DFM}

procedure TfrmLD4Q831.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD4Q83.grdMaster do
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
      QRL_NAME.Caption  := Cell[ 4,UV_iCount];
      QRL_SEX.Caption   := Cell[ 5,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q831.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q83.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q83.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q83.mskDate.Text,7,2) + '일';

end;

end.
