unit LD4Q531;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q531 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape22: TQRShape;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRLabel15: TQRLabel;
    QRShape4: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    Qrl_004: TQRLabel;
    Qrl_002: TQRLabel;
    QRShape32: TQRShape;
    Qrl_003: TQRLabel;
    QRShape5: TQRShape;
    QRShape15: TQRShape;
    QRLabel3: TQRLabel;
    Qrl_007: TQRLabel;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel6: TQRLabel;
    QRLabel8: TQRLabel;
    Qrl_005: TQRLabel;
    Qrl_006: TQRLabel;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRLabel9: TQRLabel;
    QRL_008: TQRLabel;
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
  frmLD4Q531: TfrmLD4Q531;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q53;

{$R *.DFM}

procedure TfrmLD4Q531.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var     temp,tempa :String;
  begin

   inherited;

   with frmLD4Q53.grdMaster do
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

      if (cRow mod 10) = 0 then
      begin
         QRShape13.Pen.Style := psSolid;
         QRShape13.Pen.Width := 3;
      end
      else
      begin
         QRShape13.Pen.Style := psDot;
         QRShape13.Pen.width := 1;
      end;

      Qrl_001.Caption   := Cell[1,(iRow * 50) + cRow];
      Qrl_002.Caption   := Cell[2,(iRow * 50) + cRow];
      Qrl_003.Caption   := Cell[3,(iRow * 50) + cRow];
      Qrl_008.Caption   := Cell[4,(iRow * 50) + cRow];
      Qrl_004.Caption   := Cell[5,(iRow * 50) + cRow];
      Qrl_005.Caption   := Cell[6,(iRow * 50) + cRow];
      Qrl_006.Caption   := Cell[7,(iRow * 50) + cRow];
      Qrl_007.Caption   := Cell[8,(iRow * 50) + cRow];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q531.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q53.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q53.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q53.EdtBDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q53.grdMaster.Rows mod 50) = 0) then
      sRec := Round(frmLD4Q53.grdMaster.Rows div 50)
   else
   begin
      sRec := Round(frmLD4Q53.grdMaster.Rows div 50) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
