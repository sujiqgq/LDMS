//==============================================================================
// 프로그램 설명 : 혈액샘플대장(연세대연구목적) Form
// 시스템        : 통합검진
// 수정일자      : 2007.02.22
// 수정자        : 유동구
// 수정내용      : 추가..
// 참고사항      :
//==============================================================================
unit LD7Q031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD7Q031 = class(TForm)
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
    QRL3: TQRLabel;
    QRL1: TQRLabel;
    QRShape29: TQRShape;
    QRL2: TQRLabel;
    QRL5: TQRLabel;
    QRL6: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    QRL4: TQRLabel;
    QRShape18: TQRShape;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel20: TQRLabel;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRLabel21: TQRLabel;
    QRShape37: TQRShape;
    QRLabel22: TQRLabel;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRShape41: TQRShape;
    QRLabel25: TQRLabel;
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
  frmLD7Q031: TfrmLD7Q031;

implementation

uses ZZFunc, MdmsFunc, LD7Q03;
{$R *.DFM}

procedure TfrmLD7Q031.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   with frmLD7Q03.grdMaster do
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
      QRL4.Caption  := Trim(Cell[4, UV_iCount]);
      QRL5.Caption  := Trim(Cell[5, UV_iCount]);
      QRL6.Caption  := Trim(Cell[6, UV_iCount]);
      QRL7.Caption  := Trim(Cell[7, UV_iCount]);
      QRL8.Caption  := Trim(Cell[8, UV_iCount]);
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD7Q031.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   Qrl_Jisa.Caption := '(센터 : ' + Copy(frmLD7Q03.cmbCOD_JISA.Text,4,50) + ')';
   UV_iCount := 1;
end;

end.
