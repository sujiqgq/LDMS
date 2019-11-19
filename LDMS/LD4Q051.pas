//==============================================================================
// 프로그램 설명 : EIA분주현황출력폼
// 시스템        : 통합검진
// 수정일자      : 2005.12.23
// 수정자        : 유동구
// 수정내용      : 새롭게 개발..
// 참고사항      :
//==============================================================================
unit LD4Q051;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q051 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRL_No: TQRLabel;
    QRL_bunju: TQRLabel;
    QRL_T006: TQRLabel;
    QRL_T007: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape47: TQRShape;
    QRShape48: TQRShape;
    QRLabel26: TQRLabel;
    QRLabel2: TQRLabel;
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
  frmLD4Q051: TfrmLD4Q051;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q05;

{$R *.DFM}

procedure TfrmLD4Q051.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q05.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_no.Caption    := Cell[ 1,UV_iCount];
      QRL_Bunju.Caption := Cell[ 2,UV_iCount];
      QRL_T006.Caption  := Cell[ 3,UV_iCount];
      QRL_T007.Caption  := Cell[ 4,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q051.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q05.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q05.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q05.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
