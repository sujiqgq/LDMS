unit LD7Q071;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD7Q071 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRShape10: TQRShape;
    QRL_Name: TQRLabel;
    QRShape24: TQRShape;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRShape5: TQRShape;
    QRShape7: TQRShape;
    QRL_SE89: TQRLabel;
    QRShape8: TQRShape;
    QRShape17: TQRShape;
    QRShape21: TQRShape;
    QRShape23: TQRShape;
    QRL_SE90: TQRLabel;
    QRL_Samp: TQRLabel;
    QRL_Index: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape11: TQRShape;
    QRLabel8: TQRLabel;
    QRShape19: TQRShape;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape32: TQRShape;
    QRShape36: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape3: TQRShape;
    QRShape16: TQRShape;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRL_date_2: TQRLabel;
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
  frmLD7Q071: TfrmLD7Q071;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD7Q07 ;

{$R *.DFM}

procedure TfrmLD7Q071.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD7Q07.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 출력자료 전달
      UV_iCount := UV_iCount + 1;
      QRL_Index.Caption := cell[ 1,UV_iCount];
      QRL_samp.Caption  := cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_SE89.Caption  := Cell[ 5,UV_iCount];
      QRL_SE90.Caption  := Cell[ 6,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD7Q071.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD7Q07.mskDate_start.Text,1,4) + '년'  +
                       copy(frmLD7Q07.mskDate_start.Text,5,2) + '월'  +
                       copy(frmLD7Q07.mskDate_start.Text,7,2) + '일';
   QRL_date_2.Caption := '~ ' +
                       copy(frmLD7Q07.mskDate_end.Text,1,4) + '년'  +
                       copy(frmLD7Q07.mskDate_end.Text,5,2) + '월'  +
                       copy(frmLD7Q07.mskDate_end.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
