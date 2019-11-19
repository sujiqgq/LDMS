unit LD4Q022;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q022 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape8: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    Qrl_No: TQRLabel;
    Qrl_Num_samp: TQRLabel;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRLabel6: TQRLabel;
    Qrl_Num_Bunju: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    qrl_Name: TQRLabel;
    QRBand3: TQRBand;
    QRShape7: TQRShape;
    QRLabel10: TQRLabel;
    QRShape9: TQRShape;
    QRShape17: TQRShape;
    QRShape20: TQRShape;
    qrl_gulkwa: TQRLabel;
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
  frmLD4Q022: TfrmLD4Q022;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q02;

{$R *.DFM}

procedure TfrmLD4Q022.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q02.grd_New do
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
            iRow := iRow + 1;
            cRow := 1;
            UV_iCount := UV_iCount + 1;
         End;
     if frmLD4Q02.Chk_bunju.Checked then
     begin
       QRL_date.Caption := copy(frmLD4Q02.EdtBDate.Text,1,4) + '년'  +
                           copy(frmLD4Q02.EdtBDate.Text,5,2) + '월'  +
                           copy(frmLD4Q02.EdtBDate.Text,7,2) + '일';
     end
     else
     begin
       QRLabel5.Caption := '검진일자 :';
       QRL_date.Caption := copy(frmLD4Q02.mskSt_date.Text,1,4) + '년'  +
                           copy(frmLD4Q02.mskSt_date.Text,5,2) + '월'  +
                           copy(frmLD4Q02.mskSt_date.Text,7,2) + '일';
     end;

      Qrl_No.Caption        := FormatFloat('#00', (iRow * 30) + cRow);
      Qrl_Num_samp.Caption  := Cell[2,(iRow * 30) + cRow];
      Qrl_Name.Caption      := Cell[3,(iRow * 30) + cRow];
      Qrl_Num_Bunju.Caption := Cell[4,(iRow * 30) + cRow];
      Qrl_gulkwa.Caption    := Cell[5,(iRow * 30) + cRow];
//      Qrl_Rack.Caption      := Cell[5,(iRow * 30) + cRow];

//      Qrl_No2.Caption     := FormatFloat('#00', (iRow * 30) + cRow + 30);
//      Qrl_Name2.Caption   := Cell[2,(iRow * 30) + cRow + 30];
//      Qrl_Bunju2.Caption  := Cell[3,(iRow * 30) + cRow + 30];
      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q022.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   if ((frmLD4Q02.grdMaster.Rows mod 30) = 0) then
      sRec := Round(frmLD4Q02.grdMaster.Rows div 30)
   else
   begin
      sRec := Round(frmLD4Q02.grdMaster.Rows div 30) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
