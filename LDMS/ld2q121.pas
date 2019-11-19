//==============================================================================
// 프로그램 설명 : 폼텍 출력
// 시스템        : LDMS
// 수정일자      : 2015.11.25
// 수정자        : 곽윤설
// 수정내용      : 새로생성
// 참고사항      :
//==============================================================================


unit ld2q121;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD2Q121 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_Bunju: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_Jisa: TQRLabel;
    qryBunju: TQuery;
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
  frmLD2Q121: TfrmLD2Q121;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q12;
{$R *.DFM}

procedure TfrmLD2Q121.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';
   MoreData := False;

   with frmLD2Q12.grdMaster do
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
               if   ((Cell[ 5,UV_iCount]) = (Cell[ 5,UV_iCount-1])) and
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

         qryBunju.Active := False;
         QRL_Bunju.Caption := Cell[4,UV_iCount];
         QRL_Name.Caption  := Cell[5,UV_iCount];
         QRL_samp.Caption  := '(' + Cell[10,UV_iCount] + ')';
         QRL_Jisa.Caption  := '(' + Cell[3,UV_iCount] + ')';

         UV_iRow := UV_iRow +1;

         if UV_iRow mod 21 = 1 then
         begin
           QRL_Jisa.top := 9;
           QRL_Bunju.top:= 9;
           QRL_Name.top := 27;
           QRL_samp.top := 27;
         end;

         if UV_iRow mod 4 > 1 then
         begin
           QRL_Jisa.top := QRL_Jisa.Top +1;
           QRL_Bunju.top:= QRL_Bunju.Top +1;
           QRL_Name.top := QRL_Name.Top +1;
           QRL_samp.top := QRL_samp.Top +1;
         end;

         UV_iCount := UV_iCount + 1;

   end;  //end of with
end;

procedure TfrmLD2Q121.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
   UV_iRow := 0;
end;

end.
