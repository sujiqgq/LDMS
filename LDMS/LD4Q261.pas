unit LD4Q261;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q261 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Sample: TQRLabel;
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
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape36: TQRShape;
    QRLabel37: TQRLabel;
    QRShape40: TQRShape;
    QRL_Jumin: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_DatGmgn: TQRLabel;
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
  frmLD4Q261: TfrmLD4Q261;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, ld4q26;

{$R *.DFM}

procedure TfrmLD4Q261.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmld4q26.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption      := IntToStr(UV_iCount);
      QRL_DatGmgn.Caption := Cell[ 1,UV_iCount];
      QRL_Sample.Caption  := Cell[ 2,UV_iCount];
      QRL_Name.Caption    := Cell[ 3,UV_iCount];
      QRL_Jumin.Caption   := Cell[ 4,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q261.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmld4q26.msk_SGmgn.Text,1,4) + '년'  +
                       copy(frmld4q26.msk_SGmgn.Text,5,2) + '월'  +
                       copy(frmld4q26.msk_SGmgn.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
