unit LD4Q132;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q132 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    Qrl_006: TQRLabel;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRLabel6: TQRLabel;
    Qrl_005: TQRLabel;
    QRShape26: TQRShape;
    QRShape28: TQRShape;
    Qrl_002: TQRLabel;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRLabel15: TQRLabel;
    Qrl_006_1: TQRLabel;
    QRShape7: TQRShape;
    QRShape17: TQRShape;
    QRShape2: TQRShape;
    QRLabel9: TQRLabel;
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
  frmLD4Q132: TfrmLD4Q132;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q13;

{$R *.DFM}

procedure TfrmLD4Q132.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
Var i : integer;
begin
   inherited;

   with frmLD4Q13.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      if (pos('※',Cell[5, UV_iCount]) > 0) then
      begin
         Qrl_001.Caption   := Cell[1, UV_iCount];
         Qrl_002.Caption   := Cell[2, UV_iCount];
         Qrl_003.Caption   := Cell[3, UV_iCount];
         Qrl_004.Caption   := Cell[4, UV_iCount];
         Qrl_005.Caption   := Cell[5, UV_iCount];
         if copy(Cell[6,UV_iCount],8,1) <> '' then
         begin
            case StrToInt(copy(Cell[6,UV_iCount],8,1)) of
               1,3,5,7,9 : Qrl_006.Caption := 'M/' + copy(Cell[6, UV_iCount],1,6);
               2,4,6,8   : Qrl_006.Caption := 'F/' + copy(Cell[6, UV_iCount],1,6);
               else  Qrl_006.Caption   := '';
            end;
         end
         else
            Qrl_006.Caption   := '';
         Qrl_006_1.Caption := Cell[7, UV_iCount];
      end
      else
      begin
         for i := UV_iCount to Rows do
         begin
            if (pos('※',Cell[5, UV_iCount]) > 0) then
            begin
               Qrl_001.Caption   := Cell[1, UV_iCount];
               Qrl_002.Caption   := Cell[2, UV_iCount];
               Qrl_003.Caption   := Cell[3, UV_iCount];
               Qrl_004.Caption   := Cell[4, UV_iCount];
               Qrl_005.Caption   := Cell[5, UV_iCount];
               if copy(Cell[6,UV_iCount],8,1) <> '' then
               begin
                  case StrToInt(copy(Cell[6,UV_iCount],8,1)) of
                     1,3,5,7,9 : Qrl_006.Caption := 'M/' + copy(Cell[6, UV_iCount],1,6);
                     2,4,6,8   : Qrl_006.Caption := 'F/' + copy(Cell[6, UV_iCount],1,6);
                     else  Qrl_006.Caption   := '';
                  end;
               end
               else
                  Qrl_006.Caption   := '';
               Qrl_006_1.Caption := Cell[7, UV_iCount];
               continue;
            end;
            Inc(UV_iCount);
         end;
      end;

      if (UV_iCount <= Rows) then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else if UV_iCount > Rows then
         MoreData := False;

   end;

end;

procedure TfrmLD4Q132.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q13.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q13.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q13.EdtBDate.Text,7,2) + '일';
   UV_iCount := 1;
   if ((frmLD4Q13.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD4Q13.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD4Q13.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
   UV_iCount := 1;
end;

end.
