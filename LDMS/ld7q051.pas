//==============================================================================
// 프로그램 설명 : 혈액샘플대장(연세대연구목적) Form
// 시스템        : 통합검진
// 수정일자      : 2007.02.22
// 수정자        : 유동구
// 수정내용      : 추가..
// 참고사항      :
//==============================================================================
unit LD7Q051;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD7Q051 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRBand2: TQRBand;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRLabel10: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRL1: TQRLabel;
    QRShape29: TQRShape;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRShape18: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel20: TQRLabel;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape37: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRShape43: TQRShape;
    QRShape44: TQRShape;
    Qrl_Jisa: TQRLabel;
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
  frmLD7Q051: TfrmLD7Q051;

implementation

uses ZZFunc, MdmsFunc, LD7Q05;
{$R *.DFM}

procedure TfrmLD7Q051.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   with frmLD7Q05.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL1.Caption  := Trim(Cell[1, UV_iCount]);
      QRL2.Caption  := Trim(Cell[2, UV_iCount]);
      QRL3.Caption  := Trim(Cell[3, UV_iCount]);

      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD7Q051.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   Qrl_Jisa.Caption := '(센터 : ' + Copy(frmLD7Q05.cmbCOD_JISA.Text,4,50) + ')';
   UV_iCount := 1;
end;

end.
