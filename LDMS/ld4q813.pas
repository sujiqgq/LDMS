unit LD4Q813;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QBarcode, QuickRpt, Qrctrls, ExtCtrls;

type
  TfrmLD4Q813 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    BarCode: TBarCode;
    QRLabel9: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel8: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, UV_iRow  : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var  frmLD4Q813: TfrmLD4Q813;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q81;

{$R *.DFM}
procedure TfrmLD4Q813.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';
   MoreData := False;

   with frmLD4Q81.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if (Rows = 0) or (UV_iCount > Rows)  then
      begin
         MoreData := False;
         exit;
      end;

      repeat
         if UV_iCount < Rows then
         begin
           if UV_iCount > 1 then
           begin
              // 출력자료 전달
               if   ((Cell[ 4,UV_iCount]) = (Cell[ 4,UV_iCount-1])) and
                    ((Cell[ 6,UV_iCount]) = (Cell[ 6,UV_iCount-1])) then
                    sCheck := 'N'
               else sCheck := 'Y';
           end //end of 'if (UV_iCount > 1) - begin'

           else if UV_iCount = 1 then
                   sCheck := 'Y';
         UV_iCount := UV_iCount + 1;
         end //end of 'if (UV_iCount <= Rows) - begin'

         else if (UV_iCount = Rows) then
         begin
         MoreData := False;
         exit;
         end; //end of 'if(UV_iCount > Rows) - begin'

         until (sCheck = 'Y');

         MoreData := True;

         UV_iCount := UV_iCount - 1;

         BarCode.Text    := copy(Cell[ 8,UV_iCount],3,6)  + Cell[ 4,UV_iCount];
         QRLabel9.Caption  := Cell[ 4,UV_iCount];
         QRLabel1.Caption := Cell[ 3,UV_iCount];
         QRLabel7.Caption := copy(Cell[ 8,UV_iCount],1,4) + '-' +
                             copy(Cell[ 8,UV_iCount],5,2) + '-' +
                             copy(Cell[ 8,UV_iCount],7,2);
         QRLabel4.Caption := Cell[ 6,UV_iCount];
         QRLabel5.Caption := Cell[11,UV_iCount];
         QRLabel8.Caption := Cell[10,UV_iCount];
         UV_iRow := UV_iRow +1;


         UV_iCount := UV_iCount + 1;

   end;  //end of with
end;

procedure TfrmLD4Q813.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
   UV_iRow := 0;
end;

end.
