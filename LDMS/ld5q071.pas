//==============================================================================
// 프로그램 설명    : 분주 list 현황 출력 폼
// 시스템           : 분석관리프로그램
// 수정일자         : 2006.05.17
// [수정자]수정내용 : 추가..
// 참고사항         :
//==============================================================================
unit LD5Q071;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls, Db, DBTables, ORPrintForm;

type
  TfrmLD5Q071 = class(TfrmPrintForm)
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
    QRL9: TQRLabel;
    QRL10: TQRLabel;
    QRLabel3: TQRLabel;
    QRL11: TQRLabel;
    QRLabel5: TQRLabel;
    QRL13: TQRLabel;
    QRLabel10: TQRLabel;
    QRL12: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel17: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    //procedure QRLabel9Print(sender: TObject; var Value: String);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD5Q071: TfrmLD5Q071;

implementation

{$R *.DFM}

uses LD5Q07, ZZFunc, ZZMenu, ZZComm, MdmsFunc;

procedure TfrmLD5Q071.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q07.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL1.Caption  := Cell[1,  UV_iCount];
      QRL2.Caption  := Cell[2,  UV_iCount];
      QRL3.Caption  := Cell[3,  UV_iCount];
      QRL4.Caption  := Cell[4,  UV_iCount];
      QRL5.Caption  := Cell[5,  UV_iCount];
      QRL6.Caption  := Cell[6,  UV_iCount];
      QRL7.Caption  := Cell[7,  UV_iCount];
      QRL8.Caption  := Cell[8,  UV_iCount];
      QRL9.Caption  := Cell[9,  UV_iCount];
      QRL10.Caption := Cell[10, UV_iCount];
      QRL11.Caption := Cell[11, UV_iCount];
      QRL12.Caption := Cell[12, UV_iCount];
      QRL13.Caption := Cell[13, UV_iCount];
      
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD5Q071.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD5Q07 do
   begin
      qrlFrom.Caption  := GF_DateFormat(mskDate.Text);
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;


end.
