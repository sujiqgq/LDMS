unit LD4Q631;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q631 = class(TForm)
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
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape26: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRShape27: TQRShape;
    QRShape29: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_002: TQRLabel;
    QRShape24: TQRShape;
    Qrl_003: TQRLabel;
    QRShape28: TQRShape;
    Qrl_001: TQRLabel;
    QRShape30: TQRShape;
    Qrl_008: TQRLabel;
    Qrl_006: TQRLabel;
    Qrl_007: TQRLabel;
    Qrl_004: TQRLabel;
    Qrl_009: TQRLabel;
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
  frmLD4Q631: TfrmLD4Q631;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q63;

{$R *.DFM}

procedure TfrmLD4Q631.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q63.grdMaster do
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

      if copy(Cell[4,(iRow * 50) + cRow],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[4,(iRow * 50) + cRow],8,1)) of
            1,3,5,7,9 : Qrl_004.Caption := 'M / ' + copy(Cell[4,(iRow * 50) + cRow],1,6) + '-' + copy(Cell[4,(iRow * 50) + cRow],8,14);
            2,4,6,8   : Qrl_004.Caption := 'F / ' + copy(Cell[4,(iRow * 50) + cRow],1,6) + '-' + copy(Cell[4,(iRow * 50) + cRow],8,14);
            else  Qrl_004.Caption   := '';
         end;
      end
      else
         Qrl_004.Caption   := '';


      Qrl_006.Caption   := Cell[1,(iRow * 50) + cRow + 50];
      Qrl_007.Caption   := Cell[2,(iRow * 50) + cRow + 50];
      Qrl_008.Caption   := Cell[3,(iRow * 50) + cRow + 50];

      if copy(Cell[4,(iRow * 50) + cRow + 50],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[4,(iRow * 50) + cRow + 50],8,1)) of
            1,3,5,7,9 : Qrl_009.Caption := 'M / ' + copy(Cell[4,(iRow * 50) + cRow + 50],1,6) + '-' + copy(Cell[4,(iRow * 50) + cRow + 50],8,14);
            2,4,6,8   : Qrl_009.Caption := 'F / ' + copy(Cell[4,(iRow * 50) + cRow + 50],1,6) + '-' + copy(Cell[4,(iRow * 50) + cRow + 50],8,14);
            else  Qrl_009.Caption   := '';
         end;
      end
      else
         Qrl_009.Caption   := '';


      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q631.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q63.EdtDate.Text,1,4) + '년'  +
                       copy(frmLD4Q63.EdtDate.Text,5,2) + '월'  +
                       copy(frmLD4Q63.EdtDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q63.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD4Q63.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD4Q63.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
