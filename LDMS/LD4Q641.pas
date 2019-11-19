unit LD4Q641;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q641 = class(TForm)
    QuickRep: TQuickRep;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape47: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape18: TQRShape;
    QRLabel19: TQRLabel;
    QRBand4: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;

    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRLabel14: TQRLabel;
    QRBand2: TQRBand;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape23: TQRShape;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRShape21: TQRShape;
    QRL_No: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_Bunju: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape13: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRL_Cod_hm: TQRLabel;
    QRL_Desc_hm: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel4: TQRLabel;
    QRL_dat_Bunju: TQRLabel;

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
  frmLD4Q641: TfrmLD4Q641;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q64;

{$R *.DFM}

procedure TfrmLD4Q641.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q64.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption        := Cell[ 1,UV_iCount];
      QRL_dat_Bunju.Caption := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption     := Cell[ 3,UV_iCount];
      QRL_samp.Caption      := Cell[ 4,UV_iCount];
      QRL_Name.Caption      := Cell[ 5,UV_iCount];
      QRL_Jumin.Caption     := Cell[ 6,UV_iCount];
      QRL_Cod_hm.Caption    := Cell[ 7,UV_iCount];
      QRL_Desc_hm.Caption   := Cell[ 8,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q641.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q64.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q64.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q64.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
