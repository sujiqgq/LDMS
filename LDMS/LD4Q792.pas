unit LD4Q792;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q792 = class(TForm)
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
    QRShape22: TQRShape;
    QRLabel6: TQRLabel;
    QRShape9: TQRShape;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRShape23: TQRShape;
    QRLabel14: TQRLabel;
    QRShape29: TQRShape;
    QRShape37: TQRShape;
    QRBand2: TQRBand;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    Qrl_006: TQRLabel;
    QRShape28: TQRShape;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    Qrl_014: TQRLabel;
    Qrl_011: TQRLabel;
    Qrl_009: TQRLabel;
    Qrl_012: TQRLabel;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape5: TQRShape;
    QRLabel3: TQRLabel;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel10: TQRLabel;
    QRShape13: TQRShape;
    Qrl_002: TQRLabel;
    qrl_010: TQRLabel;
    QRShape15: TQRShape;
    QRShape19: TQRShape;
    QRLabel16: TQRLabel;
    QRShape20: TQRShape;
    QRShape35: TQRShape;
    QRLabel8: TQRLabel;
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
  frmLD4Q792: TfrmLD4Q792;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q79;

{$R *.DFM}

procedure TfrmLD4Q792.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var sSelect, sWhere, sOrderBy, sDate : String;
begin
   inherited;

   with frmLD4Q79.grdMaster do
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

      Qrl_001.Caption   := FormatFloat('#00', (iRow * 25)+ cRow);
      Qrl_003.Caption   := Cell[2,(iRow * 25)+ cRow];
      Qrl_004.Caption   := Cell[3,(iRow * 25)+ cRow];
      Qrl_006.Caption   := Cell[4,(iRow * 25) + cRow];
      Qrl_002.Caption   := Cell[5,(iRow * 25) + cRow];
//------------------------------------------------------------------------------
      Qrl_009.Caption   := FormatFloat('#00', (iRow * 25) + cRow + 25);
      Qrl_011.Caption   := Cell[2,(iRow * 25) + cRow + 25];
      Qrl_012.Caption   := Cell[3,(iRow * 25) + cRow + 25];
      Qrl_014.Caption   := Cell[4,(iRow * 25) + cRow + 25];
      Qrl_010.Caption   := Cell[5,(iRow * 25) + cRow + 25];
//------------------------------------------------------------------------------
      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q792.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q79.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q79.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q79.EdtBDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q79.grdMaster.Rows mod 50) = 0) then
      sRec := Round(frmLD4Q79.grdMaster.Rows div 50)
   else
   begin
      sRec := Round(frmLD4Q79.grdMaster.Rows div 50) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
