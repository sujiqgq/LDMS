//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : LDMS
// 수정일자      : 2012.01.11
// 수정자        : 송재호
// 수정내용      : S019 폼텍 양식
// 참고사항      :
//==============================================================================
unit LD4Q233;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q233 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_Bunju: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_date: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q233: TfrmLD4Q233;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q23;
{$R *.DFM}

procedure TfrmLD4Q233.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';

   with frmLD4Q23.grdS021 do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      repeat
         if UV_iCount <= Rows then
         begin
            // 출력자료 전달
            if UV_iCount > 1 then
            begin
               if Cell[ 4,UV_iCount] <> Cell[ 4,UV_iCount - 1] then
                  sCheck := 'Y';
            end
            else if UV_iCount = 1 then
               sCheck := 'Y';
         end
         else sCheck := 'Y';

         UV_iCount := UV_iCount + 1;
      until (sCheck = 'Y');

      UV_iCount := UV_iCount - 1;

      QRL_Name.Caption    := Cell[ 5,UV_iCount];
      QRL_Bunju.Caption   := Cell[ 4,UV_iCount];
      QRL_date.Caption    := Cell[ 3,UV_iCount];

      UV_iCount := UV_iCount + 1;
      if UV_iCount <= Rows + 1 then MoreData := True
      else MoreData := False;
   end;
end;

procedure TfrmLD4Q233.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
