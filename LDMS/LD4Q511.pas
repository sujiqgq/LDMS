unit LD4Q511;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q511 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Num_samp: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape35: TQRShape;
    QRLabel38: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_S018: TQRLabel;
    QRShape5: TQRShape;
    QRL_S021: TQRLabel;
    QRShape6: TQRShape;
    QRL_Bunju: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape8: TQRShape;
    QRLabel6: TQRLabel;
    QRL_S019: TQRLabel;
    QRShape14: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape2: TQRShape;
    QRLabel3: TQRLabel;
    QRL_C044: TQRLabel;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
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
  frmLD4Q511: TfrmLD4Q511;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q51;

{$R *.DFM}

procedure TfrmLD4Q511.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q51.grdMaster do
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
      QRL_Num_samp.Caption    := Cell[ 2,UV_iCount];
      QRL_Name.Caption  := Cell[ 3,UV_iCount];
      QRL_Bunju.Caption := Cell[ 4,UV_iCount];
      QRL_S021.Caption  := Cell[ 5,UV_iCount];
      QRL_S018.Caption  := Cell[ 6,UV_iCount];
      QRL_S019.Caption  := Cell[ 7,UV_iCount];
      QRL_C044.Caption  := Cell[ 8,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q511.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q51.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q51.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q51.mskDate.Text,7,2) + '일';
                       
   UV_iCount := 0;
end;

end.
