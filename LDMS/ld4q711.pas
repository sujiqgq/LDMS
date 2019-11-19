unit ld4q711;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q711 = class(TForm)
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
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel6: TQRLabel;
    Qrl_005: TQRLabel;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRShape23: TQRShape;
    QRLabel14: TQRLabel;
    QRShape28: TQRShape;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    Qrl_011: TQRLabel;
    Qrl_010: TQRLabel;
    Qrl_009: TQRLabel;
    Qrl_007: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel8: TQRLabel;
    Qrl_006: TQRLabel;
    QRL_012: TQRLabel;
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
  frmLD4Q711: TfrmLD4Q711;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q71;

{$R *.DFM}

procedure TfrmLD4Q711.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q71.grdMaster do
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

      Qrl_001.Caption   := FormatFloat('#00', (iRow * 50) + cRow);
      Qrl_003.Caption   := Cell[2,(iRow * 50) + cRow];
      Qrl_004.Caption   := Cell[3,(iRow * 50) + cRow];
      Qrl_005.Caption   := Cell[4,(iRow * 50) + cRow];
      Qrl_006.Caption   := Cell[5,(iRow * 50) + cRow];

      Qrl_007.Caption   := FormatFloat('#00', (iRow * 50) + cRow + 50);
      Qrl_009.Caption   := Cell[2,(iRow * 50) + cRow + 50];
      Qrl_010.Caption   := Cell[3,(iRow * 50) + cRow + 50];
      Qrl_011.Caption   := Cell[4,(iRow * 50) + cRow + 50];
      Qrl_012.Caption   := Cell[5,(iRow * 50) + cRow + 50];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q711.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q71.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q71.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q71.EdtBDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q71.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD4Q71.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD4Q71.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
