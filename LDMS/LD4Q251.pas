unit LD4Q251;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q251 = class(TForm)
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
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape22: TQRShape;
    QRLabel6: TQRLabel;
    QRShape26: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape23: TQRShape;
    QRLabel14: TQRLabel;
    QRShape27: TQRShape;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape33: TQRShape;
    QRShape35: TQRShape;
    QRLabel16: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape37: TQRShape;
    QRBand2: TQRBand;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    Qrl_007: TQRLabel;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    Qrl_006: TQRLabel;
    QRShape28: TQRShape;
    Qrl_002: TQRLabel;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    QRShape30: TQRShape;
    Qrl_015: TQRLabel;
    Qrl_014: TQRLabel;
    Qrl_013: TQRLabel;
    Qrl_011: TQRLabel;
    Qrl_010: TQRLabel;
    Qrl_009: TQRLabel;
    QRShape32: TQRShape;
    Qrl_008: TQRLabel;
    QRShape34: TQRShape;
    Qrl_016: TQRLabel;
    Qrl_005: TQRLabel;
    QRShape36: TQRShape;
    Qrl_012: TQRLabel;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
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
  frmLD4Q251: TfrmLD4Q251;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q25;

{$R *.DFM}

procedure TfrmLD4Q251.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var sSelect, sWhere, sOrderBy, sDate : String;
begin
   inherited;

   with frmLD4Q25.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 25 Then
         Begin
            iRow := iRow + 2;
            cRow := 1;
            UV_iCount := UV_iCount + 1;
         End;

   {   if (cRow mod 10) = 0 then
      begin
         QRShape13.Pen.Style := psSolid;
         QRShape13.Pen.Width := 5;
      end
      else
      begin
         QRShape13.Pen.Style := psSolid;
         QRShape13.Pen.width := 1;
      end; }

      Qrl_001.Caption   := FormatFloat('#00', (iRow * 25)+ cRow);
      Qrl_002.Caption   := Cell[3,(iRow * 25)+ cRow];
      Qrl_003.Caption   := Cell[4,(iRow * 25)+ cRow];
      Qrl_004.Caption   := Cell[5,(iRow * 25)+ cRow];
      Qrl_005.Caption   := Cell[2,(iRow * 25)+ cRow];
      if length(Cell[6,(iRow * 25) + cRow]) > 8 then
         Qrl_006.Caption   := copy(Cell[6,(iRow * 25) + cRow], 1, 8)
      else
         Qrl_006.Caption   := Cell[6,(iRow * 25) + cRow];
      if copy(Cell[7,(iRow * 25) + cRow],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[7,(iRow * 25) + cRow],8,1)) of
            1,3,5,7,9 : Qrl_007.Caption := 'M/' + copy(Cell[7,(iRow * 25) + cRow],1,6);
            2,4,6,8   : Qrl_007.Caption := 'F/' + copy(Cell[7,(iRow * 25) + cRow],1,6);
            else  Qrl_007.Caption   := '';
         end;
      end
      else
         Qrl_007.Caption   := '';
      Qrl_008.Caption := Cell[8,(iRow * 25) + cRow];
      if Qrl_008.Caption <> '' then
         sDate := copy(Qrl_008.Caption,1,4) + copy(Qrl_008.Caption,6,2) + copy(Qrl_008.Caption,9,2)
      else
         sDate := '';



      Qrl_009.Caption   := FormatFloat('#00', (iRow * 25) + cRow + 25);
      Qrl_010.Caption   := Cell[3,(iRow * 25) + cRow + 25];
      Qrl_011.Caption   := Cell[4,(iRow * 25) + cRow + 25];
      Qrl_012.Caption   := Cell[5,(iRow * 25) + cRow + 25];
      Qrl_013.Caption   := Cell[2,(iRow * 25) + cRow + 25];
      if length(Cell[6,(iRow * 25) + cRow]) > 8 then
         Qrl_014.Caption   := copy(Cell[6,(iRow * 25) + cRow + 25], 1, 8)
      else
         Qrl_014.Caption   := Cell[6,(iRow * 25) + cRow + 25];
      if copy(Cell[7,(iRow * 25) + cRow + 25],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[7,(iRow * 25) + cRow + 25],8,1)) of
            1,3,5,7,9 : Qrl_015.Caption := 'M/' + copy(Cell[7,(iRow * 25) + cRow + 25],1,6);
            2,4,6,8   : Qrl_015.Caption := 'F/' + copy(Cell[7,(iRow * 25) + cRow + 25],1,6);
            else  Qrl_015.Caption   := '';
         end;
      end
      else
         Qrl_015.Caption   := '';
      Qrl_016.Caption := Cell[8,(iRow * 25) + cRow + 25];
      if Qrl_016.Caption <> '' then
         sDate := copy(Qrl_016.Caption,1,4) + copy(Qrl_016.Caption,6,2) + copy(Qrl_016.Caption,9,2)
      else
         sDate := '';



      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q251.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q25.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q25.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q25.EdtBDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q25.grdMaster.Rows mod 50) = 0) then
      sRec := Round(frmLD4Q25.grdMaster.Rows div 50)
   else
   begin
      sRec := Round(frmLD4Q25.grdMaster.Rows div 50) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
