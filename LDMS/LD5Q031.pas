//==============================================================================
// 프로그램 설명 : [WorkList]파트별 검사항목 대장출력
// 시스템        : 통합검진
// 수정일자      : 2007.03.21
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : QRL5 사이즈 늘림(외국인의경우 이름때문에 다음장으로 넘어감)
// 수정일자      : 2011.05.18
// 수정자        : 김승철
//==============================================================================

unit LD5Q031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD5Q031 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
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
    QRLabel5: TQRLabel;
    QRBand2: TQRBand;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRL5: TQRLabel;
    QRL6: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
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
  frmLD5Q031: TfrmLD5Q031;

implementation

{$R *.DFM}

uses ZZfunc, LD5Q03;

procedure TfrmLD5Q031.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q03.grdMaster do
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

      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD5Q031.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD5Q03 do
   begin
      qrlFrom.Caption  := GF_DateFormat(mskDate.Text);
      QRL_Date.Caption := GV_sPrintToday;
      case Com_Part.ItemIndex of
          0 : QRLabel1.Caption := '파트별 검사항목 대장 - 생화학';
          1 : QRLabel1.Caption := '파트별 검사항목 대장 - 혈액학';
          2 : QRLabel1.Caption := '파트별 검사항목 대장 - Urin';
          3 : QRLabel1.Caption := '파트별 검사항목 대장 - RIA';
          4 : QRLabel1.Caption := '파트별 검사항목 대장 - EIA';
          5 : QRLabel1.Caption := '파트별 검사항목 대장 - 조직학';
          6 : QRLabel1.Caption := '파트별 검사항목 대장 - 혈청학';
          7 : QRLabel1.Caption := '파트별 검사항목 대장 - 혈액기타(외주)';
          8 : QRLabel1.Caption := '파트별 검사항목 대장 - 혈액기타(특검)';
      end;
   end;
end;

end.
