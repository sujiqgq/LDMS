//==============================================================================
// 프로그램 설명 : 폼텍 출력
// 시스템        : LDMS
// 수정일자      : 2017.03.23
// 수정자        : 박수지
// 수정내용      :
// 참고사항      :
//==============================================================================


unit LD4Q382;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables, QBarcode;

type
  TfrmLD4Q382 = class(TForm)
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

var
  frmLD4Q382: TfrmLD4Q382;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q38 ;
{$R *.DFM}

procedure TfrmLD4Q382.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';
   MoreData := False;

   with frmLD4Q38.grdMaster do
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
              sCheck := 'Y';
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

         UV_iCount := UV_iCount -1 ;
         if frmLD4Q38.Chk_ALL.Checked then
         begin
         BarCode.Text         := copy(Cell[ 11, UV_iCount],3,6) + Cell[ 12,UV_iCount];
         QRLabel9.Caption     := Cell[ 12,UV_iCount];
         QRLabel1.Caption     := Cell[ 1,UV_iCount];
         QRLabel7.Caption     := copy(Cell[ 11,UV_iCount],1,4) + '-' +
                                 copy(Cell[ 11,UV_iCount],5,2) + '-' +
                                 copy(Cell[ 11,UV_iCount],7,2) ;
         QRLabel4.Caption     := Cell[ 3,UV_iCount];
         QRLabel5.Caption     := Cell[13,UV_iCount];
         QRLabel8.Caption     := Cell[14,UV_iCount];
         end
         else
         begin
         BarCode.Text         := copy(Cell[ 5, UV_iCount],3,6) + Cell[ 6,UV_iCount];
         QRLabel9.Caption     := Cell[ 6,UV_iCount];
         QRLabel1.Caption     := Cell[ 1,UV_iCount];
         QRLabel7.Caption     := copy(Cell[ 5,UV_iCount],1,4) + '-' +
                                 copy(Cell[ 5,UV_iCount],5,2) + '-' +
                                 copy(Cell[ 5,UV_iCount],7,2) ;
         QRLabel4.Caption     := Cell[ 3,UV_iCount];
         QRLabel5.Caption     := Cell[ 7,UV_iCount];
         QRLabel8.Caption     := Cell[ 8,UV_iCount];
         end;

         UV_iCount := UV_iCount + 1;

   end;  //end of with
end;

procedure TfrmLD4Q382.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
   UV_iRow := 0;
end;

end.
