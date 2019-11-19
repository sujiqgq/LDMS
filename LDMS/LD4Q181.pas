unit ld4q181;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q181 = class(TForm)
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
    QRShape36: TQRShape;
    QRShape40: TQRShape;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Sample: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRLabel3: TQRLabel;
    Qrl_rack: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_Munjin: TQRLabel;
    QRLabel6: TQRLabel;
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
  frmLD4Q181: TfrmLD4Q181;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q18;

{$R *.DFM}

procedure TfrmLD4Q181.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q18.grdMaster do
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
      if frmLD4Q18.cmbHM.Text = 'S016' then
      begin
      QRL_Munjin.Caption := '문진[만성C형간염체크] :' + copy(Cell[ 6,UV_iCount],1,1);
      end
      else QRL_Munjin.Caption := '';

//      if (UV_iCount mod 5) = 0 then QRShape13.Pen.Width := 5
//      else QRShape13.Pen.Width := 2;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q181.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q18.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q18.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q18.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
