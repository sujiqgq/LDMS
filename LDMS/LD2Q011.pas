//==============================================================================
// 프로그램 설명 : 분주List 출력
// 시스템        : 통합검진
// 수정일자      : 1999.08.09
// 수정자        : 김은석
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD2Q011;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD2Q011 = class(TfrmPrintForm)
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
    etc: TQRLabel;
    QRLabel10: TQRLabel;
    QRL10: TQRLabel;
    QRL11: TQRLabel;
    QRLabel17: TQRLabel;
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
  frmLD2Q011: TfrmLD2Q011;

implementation

{$R *.DFM}

uses LD2Q01, ZZfunc;

procedure TfrmLD2Q011.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q01.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL1.Caption  := Cell[1,  UV_iCount];
      QRL11.Caption := Cell[2,  UV_iCount];
      QRL2.Caption  := Cell[3,  UV_iCount];
      QRL3.Caption  := Cell[4,  UV_iCount];
      QRL10.Caption := Cell[5,  UV_iCount];
      QRL4.Caption  := Cell[6,  UV_iCount];
      QRL5.Caption  := Cell[7,  UV_iCount];
      QRL6.Caption  := Cell[8,  UV_iCount];
      QRL7.Caption  := Cell[9,  UV_iCount];
      QRL8.Caption  := Cell[10, UV_iCount];

      if length(Cell[11, UV_iCount]) >= 60 then
         QRL9.Caption := copy(Cell[11, UV_iCount],1,60) + ' ' + copy(Cell[11, UV_iCount],61, length(Cell[11, UV_iCount]))
      else
         QRL9.Caption := Cell[11, UV_iCount];

      etc.Caption  := Cell[12, UV_iCount];
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD2Q011.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD2Q01 do
   begin
      qrlFrom.Caption  := GF_DateFormat(mskDate.Text);
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

end.
