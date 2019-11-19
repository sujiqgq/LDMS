//==============================================================================
// 프로그램 설명 : 항목별 지사별 검진인원 출력
// 시스템        : 통합검진
// 수정일자      : 1999.11.3
// 수정자        : 허정남
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8Q031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD8Q031 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    DetailBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_Date: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRS_Page: TQRSysData;
    QRBand1: TQRBand;
    QRLabel12: TQRLabel;
    qrlFrom: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRL5: TQRLabel;
    QRL6: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    QRLabel3: TQRLabel;
    QRL9: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    qrlTo: TQRLabel;
    QRL10: TQRLabel;
    QRL11: TQRLabel;
    QRL12: TQRLabel;
    QRL13: TQRLabel;
    QRL14: TQRLabel;
    QRL15: TQRLabel;
    QRL16: TQRLabel;
    QRL17: TQRLabel;
    QRL18: TQRLabel;
    QRL19: TQRLabel;
    QRL20: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD8Q031: TfrmLD8Q031;

implementation

{$R *.DFM}

uses LD8Q03;

procedure TfrmLD8Q031.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD8Q03.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL1.Caption := Cell[1, UV_iCount];
      QRL2.Caption := Cell[2, UV_iCount];
      QRL3.Caption := Cell[3, UV_iCount];
      QRL4.Caption := Cell[4, UV_iCount];
      QRL5.Caption := Cell[5, UV_iCount];
      QRL6.Caption := Cell[6, UV_iCount];
      QRL7.Caption := Cell[7, UV_iCount];
      QRL8.Caption := Cell[8, UV_iCount];
      QRL9.Caption := Cell[9, UV_iCount];
      QRL10.Caption := Cell[10, UV_iCount];
      QRL11.Caption := Cell[11, UV_iCount];
      QRL12.Caption := Cell[12, UV_iCount];
      QRL13.Caption := Cell[13, UV_iCount];
      QRL14.Caption := Cell[14, UV_iCount];
      QRL15.Caption := Cell[15, UV_iCount];
      QRL16.Caption := Cell[16, UV_iCount];
      QRL17.Caption := Cell[17, UV_iCount];
      QRL18.Caption := Cell[18, UV_iCount];
      QRL19.Caption := Cell[19, UV_iCount];
      QRL20.Caption := Cell[20, UV_iCount];
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD8Q031.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD8Q03 do
   begin
      qrlFrom.Caption  := mskST_Date.Text;
      qrlTo.Caption    := mskED_Date.Text;
   end;
end;

end.
