unit LD2Q120;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, Math;

type
  TfrmLD2Q120 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape23: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_S008: TQRLabel;
    QRL_S091: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRShape42: TQRShape;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape18: TQRShape;
    QRShape21: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel10: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRShape51: TQRShape;
    QRLabel14: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRShape2: TQRShape;
    QRL_Rack: TQRLabel;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape9: TQRShape;
    QRShape16: TQRShape;
    QRShape19: TQRShape;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape20: TQRShape;
    QRLabel12: TQRLabel;
    QRShape22: TQRShape;
    QRLabel13: TQRLabel;
    QRShape24: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape25: TQRShape;
    QRL_Bunju2: TQRLabel;
    QRL_S0082: TQRLabel;
    QRL_S0912: TQRLabel;
    QRShape26: TQRShape;
    QRShape28: TQRShape;
    QRShape31: TQRShape;
    QRL_Jumin2: TQRLabel;
    QRShape32: TQRShape;
    QRL_No2: TQRLabel;
    QRShape33: TQRShape;
    QRL_samp2: TQRLabel;
    QRL_Name2: TQRLabel;
    QRShape34: TQRShape;
    QRL_Rack2: TQRLabel;
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
  frmLD2Q120: TfrmLD2Q120;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q12;

{$R *.DFM}

procedure TfrmLD2Q120.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q12.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 50 Then
      Begin
         iRow := iRow + 2;
         cRow := 1;
         UV_iCount := UV_iCount + 1;
      End;

      QRL_No.Caption    := Cell[ 1,(iRow * 50) + cRow];
      QRL_Rack.Caption  := Cell[ 2,(iRow * 50) + cRow];
      QRL_samp.Caption  := Cell[10,(iRow * 50) + cRow];
      QRL_Bunju.Caption := Cell[ 4,(iRow * 50) + cRow];
      if Length(Cell[ 5,(iRow * 50) + cRow]) >= 10 then QRL_Name.Caption  :=  copy(Cell[ 5,(iRow * 50) + cRow],1,10)
      else QRL_Name.Caption  := Cell[ 5,(iRow * 50) + cRow];
      QRL_Jumin.Caption := Cell[ 6,(iRow * 50) + cRow];
      QRL_S008.Caption  := Cell[ 7,(iRow * 50) + cRow];
      QRL_S091.Caption  := Cell[ 8,(iRow * 50) + cRow];

      QRL_No2.Caption    := Cell[ 1,(iRow * 50) + cRow + 50];
      QRL_Rack2.Caption  := Cell[ 2,(iRow * 50) + cRow + 50];
      QRL_samp2.Caption  := Cell[10,(iRow * 50) + cRow + 50];
      QRL_Bunju2.Caption := Cell[ 4,(iRow * 50) + cRow + 50];
      if Length(Cell[ 5,(iRow * 50) + cRow + 50]) >= 8 then QRL_Name2.Caption  :=  copy(Cell[ 5,(iRow * 50) + cRow + 50],1,8)
      else QRL_Name2.Caption  := Cell[ 5,(iRow * 50) + cRow + 50];
      QRL_Jumin2.Caption := Cell[ 6,(iRow * 50) + cRow + 50];
      QRL_S0082.Caption  := Cell[ 7,(iRow * 50) + cRow + 50];
      QRL_S0912.Caption  := Cell[ 8,(iRow * 50) + cRow + 50];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q120.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q12.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q12.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q12.EdtBDate.Text,7,2) + '일';
   sRec := Ceil(frmLD2Q12.grdMaster.Rows / 99);
   UV_iCount := 1; cRow := 0; iRow := 0;
end;

end.
