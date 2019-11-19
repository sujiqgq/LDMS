//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : LDMS
// 수정일자      : 2011.06.01
// 수정자        : 송재호
// 수정내용      : 녹십자 폼텍 양식
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.17
// 수정자        : 곽윤설
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD2Q082;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD2Q082 = class(TForm)
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
    UV_iCount, UV_iRow : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD2Q082: TfrmLD2Q082;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q08;
{$R *.DFM}

procedure TfrmLD2Q082.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';

   with frmLD2Q08.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if (Rows = 0) OR (Rows = UV_iCount) then
      begin
         MoreData := False;
         exit;
      end;

      UV_iCount := UV_iCount + 1;
      repeat
         if UV_iCount <= Rows then
         begin
            // 출력자료 전달
            if UV_iCount > 1 then
            begin
               if (Cell[ 4,UV_iCount] <> Cell[ 4,UV_iCount - 1]) then
                  sCheck := 'Y';
            end
            else if UV_iCount = 1 then
               sCheck := 'Y';
         end
         else sCheck := 'Y';

         UV_iCount := UV_iCount + 1;
      until (sCheck = 'Y');

      UV_iCount := UV_iCount - 1;

//      if Cell[ 1,UV_iCount] <> '총합계' then
//      begin
         qryBunju.Active := False;
         qryBunju.ParamByName('cod_bjjs').AsString  := frmLD2Q08.UV_Bjjs;
         qryBunju.ParamByName('num_bunju').AsString := copy(Cell[ 1,UV_iCount],1,4);
         qryBunju.ParamByName('dat_bunju').AsString := Cell[ 12,UV_iCount];
         qryBunju.Active := True;
         if qryBunju.RecordCount > 0 then
         begin
            GP_query_log(qryBunju, 'LD2Q082', 'Q', 'Y');
            QRL_Jisa.Caption  := qryBunju.FieldByName('cod_jisa').AsString;
            QRL_Bunju.Caption := qryBunju.FieldByName('num_bunju').AsString;
         end;
         qryBunju.Active := False;
//      end;
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

      QRL_Name.Caption  := Cell[ 7,UV_iCount];
      QRL_samp.Caption  := '(' + Cell[ 4,UV_iCount] + ')';


      if UV_iCount <= Rows then MoreData := True
      else MoreData := False;
   end;
end;

procedure TfrmLD2Q082.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
   UV_iRow := 0;
end;

end.
