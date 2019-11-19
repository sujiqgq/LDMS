//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : LDMS
// 수정일자      : 2011.06.01
// 수정자        : 송재호
// 수정내용      : 녹십자 폼텍 양식
// 참고사항      :
//==============================================================================
unit LD4Q252;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q252 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_DAT_Bunju: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_Bunju: TQRLabel;
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
  frmLD4Q252: TfrmLD4Q252;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q25;
{$R *.DFM}

procedure TfrmLD4Q252.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   sCheck : String;
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
         if UV_iCount <= Rows then
         begin
         UV_iCount := UV_iCount + 1;
         QRL_DAT_Bunju.Caption  := '분주일:'+frmLD4Q25.edtBDate.Text;
         QRL_Bunju.Caption  :=Cell[ 5,UV_iCount];
         QRL_Name.Caption  := Cell[ 6,UV_iCount];
         end;
      if UV_iCount <= Rows then MoreData := True
      else MoreData := False;
     end;
end;

procedure TfrmLD4Q252.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
