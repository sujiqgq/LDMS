unit LD5Q121;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD5Q121 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape6: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape26: TQRShape;
    QRShape31: TQRShape;
    QRLabel15: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape21: TQRShape;
    QRL1: TQRLabel;
    QRL5: TQRLabel;
    QRShape24: TQRShape;
    QRL6: TQRLabel;
    QRShape28: TQRShape;
    QRL2: TQRLabel;
    QRShape32: TQRShape;
    QRL7: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape5: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRShape29: TQRShape;
    QRShape15: TQRShape;
    QRShape19: TQRShape;
    QRShape23: TQRShape;
    QRShape30: TQRShape;
    QRLabel9: TQRLabel;
    QRL8: TQRLabel;
    QRLabel11: TQRLabel;
    QRL9: TQRLabel;
    QRShape7: TQRShape;
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
  frmLD5Q121: TfrmLD5Q121;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q12;

{$R *.DFM}

procedure TfrmLD5Q121.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q12.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      QRL1.Caption  := Cell[  1,UV_iCount];
      QRL2.Caption  := Cell[  3,UV_iCount];
      QRL3.Caption  := Cell[  4,UV_iCount];
      QRL4.Caption  := Cell[  2,UV_iCount];

      QRL5.Caption  := copy(Cell[  5,UV_iCount],1,6);
      QRL6.Caption  := Cell[  6,UV_iCount];
      QRL7.Caption  := Cell[  7,UV_iCount];
      QRL8.Caption  := Copy(Cell[  9,UV_iCount], 1,18) + #10#13 +
                       Copy(Cell[  9,UV_iCount],19,36) + #10#13 +
                       Copy(Cell[  9,UV_iCount],37,54) + #10#13 +
                       Copy(Cell[  9,UV_iCount],55,length(Cell[  9,UV_iCount])) + #10#13;
      QRL9.Caption  := Cell[ 10,UV_iCount];


      if UV_iCount <= Rows then
      begin
         MoreData := True;
         // 출력자료 전달
         UV_iCount := UV_iCount + 1;
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD5Q121.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   QRL_date.Caption := copy(frmLD5Q12.EdtBDate.Text,1,4) + '년' +
                       copy(frmLD5Q12.EdtBDate.Text,5,2) + '월' +
                       copy(frmLD5Q12.EdtBDate.Text,7,2) + '일';

   UV_iCount := 1;
   cRow := 0; iRow := 0;
end;

end.
