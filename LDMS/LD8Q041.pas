unit LD8Q041;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD8Q041 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_Bunju1: TQRLabel;
    Qrl_Bunju2: TQRLabel;
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
  frmLD8Q041: TfrmLD8Q041;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q04;

{$R *.DFM}

procedure TfrmLD8Q041.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD8Q04.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 30 Then
         Begin
            iRow := iRow + 2;
            cRow := 1;
            UV_iCount := UV_iCount + 1;
         End;
//      Qrl_No1.Caption := FormatFloat('#00',cRow);
      Qrl_Bunju1.Caption := Cell[2,(iRow * 30) + cRow];
//      Qrl_No2.Caption := FormatFloat('#00',cRow);
      Qrl_Bunju2.Caption := Cell[2,(iRow * 30) + cRow + 30];
      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD8Q041.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD8Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD8Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD8Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
   sRec := Round(frmLD8Q04.grdMaster.Rows / 60);
   cRow := 0; iRow := 0;
end;

end.
