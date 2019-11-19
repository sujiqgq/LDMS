unit LD4Q061;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, Math;

type
  TfrmLD4Q061 = class(TForm)
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
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_Bunju1: TQRLabel;
    Qrl_Bunju2: TQRLabel;
    QRLabel3: TQRLabel;
    qrl_bunju3: TQRLabel;
    QRShape10: TQRShape;
    QRShape14: TQRShape;
    QRShape18: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRL_Name1: TQRLabel;
    QRL_Name2: TQRLabel;
    QRL_Name3: TQRLabel;
    QRShape13: TQRShape;
    QRShape4: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRL_No1: TQRLabel;
    QRL_No2: TQRLabel;
    QRL_No3: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
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
  frmLD4Q061: TfrmLD4Q061;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q06;

{$R *.DFM}

procedure TfrmLD4Q061.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q06.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 33 Then
      Begin
         iRow := iRow + 3;
         cRow := 1;
         UV_iCount := UV_iCount + 1;
      End;
      QRL_No1.Caption    := Cell[1,(iRow * 33) + cRow];
      QRL_No2.Caption    := Cell[1,(iRow * 33) + cRow + 33];
      QRL_No3.Caption    := Cell[1,(iRow * 33) + cRow + 66];
      Qrl_Bunju1.Caption := Cell[2,(iRow * 33) + cRow];
      Qrl_Bunju2.Caption := Cell[2,(iRow * 33) + cRow + 33];
      Qrl_Bunju3.Caption := Cell[2,(iRow * 33) + cRow + 66];
      QRL_Name1.Caption  := copy(Cell[3,(iRow * 33) + cRow],1,6);
      QRL_Name2.Caption  := copy(Cell[3,(iRow * 33) + cRow + 33],1,6);
      QRL_Name3.Caption  := copy(Cell[3,(iRow * 33) + cRow + 66],1,6);
   end;
   if UV_iCount <= sRec then
      MoreData := True
   else
      MoreData := False;
end;

procedure TfrmLD4Q061.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   if frmLD4Q06.RB_Y004.Checked then QRLabel1.Caption := 'Y004 분 주 표'
   else QRLabel1.Caption := 'U017,U046 분 주 표';
   QRL_date.Caption  := copy(frmLD4Q06.EdtBDate.Text,1,4) + '년'  +
                        copy(frmLD4Q06.EdtBDate.Text,5,2) + '월'  +
                        copy(frmLD4Q06.EdtBDate.Text,7,2) + '일';
   sRec := Ceil(frmLD4Q06.grdMaster.Rows / 99);
   UV_iCount := 1; cRow := 0; iRow := 0;
end;

end.
 