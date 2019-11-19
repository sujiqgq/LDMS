unit LD4Q291;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q291 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Samp: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
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
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRShape2: TQRShape;
    QRL_U004: TQRLabel;
    QRShape5: TQRShape;
    QRL_U003: TQRLabel;
    QRShape6: TQRShape;
    QRLabel11: TQRLabel;
    QRShape8: TQRShape;
    QRL_Jumin: TQRLabel;
    QRLabel3: TQRLabel;
    QRL_U005: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_U009: TQRLabel;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape9: TQRShape;
    QRShape14: TQRShape;
    QRShape16: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
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
  frmLD4Q291: TfrmLD4Q291;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q29;

{$R *.DFM}

procedure TfrmLD4Q291.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;


     // QRLabel10.caption:=frmLD4Q29.hm_temp;

   with frmLD4Q29.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption    := Cell[  1,UV_iCount];
      QRL_Samp.Caption  := Cell[  3,UV_iCount];
      QRL_Name.Caption  := Cell[  4,UV_iCount];
      QRL_Jumin.Caption := Cell[  5,UV_iCount];
      QRL_U003.Caption  := Cell[  6,UV_iCount];
      QRL_U004.Caption  := Cell[  7,UV_iCount];
      QRL_U005.Caption  := Cell[  8,UV_iCount];
      QRL_U009.Caption  := Cell[  9,UV_iCount];
      QRLabel10.Caption := Cell[ 10,UV_iCount];
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q291.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q29.edtDate.Text,1,4) + '년'  +
                       copy(frmLD4Q29.edtDate.Text,5,2) + '월'  +
                       copy(frmLD4Q29.edtDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
