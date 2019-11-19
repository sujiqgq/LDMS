//==============================================================================
// 프로그램 설명 : [WorkList]특수검진 검사항목 대장출력
// 시스템        : 통합검진
// 수정일자      : 2008.01.09
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q121;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD4Q121 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRS_Page: TQRSysData;
    QRLabel12: TQRLabel;
    qrlFrom: TQRLabel;
    QRBand2: TQRBand;
    Qrl_3: TQRLabel;
    Qrl_4: TQRLabel;
    Qrl_7: TQRLabel;
    Qrl_2: TQRLabel;
    Qrl_1: TQRLabel;
    QRBand1: TQRBand;
    QRLabel5: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRLabel3: TQRLabel;
    Qrs_1: TQRShape;
    Qrs_2: TQRShape;
    Qrs_3: TQRShape;
    Qrs_4: TQRShape;
    Qrs_5: TQRShape;
    Qrs_6: TQRShape;
    Qrs_7: TQRShape;
    Qrs_8: TQRShape;
    Qrr_5: TQRRichText;
    Qrr_6: TQRRichText;
    Qrr_8: TQRRichText;
    Qrr_9: TQRRichText;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD4Q121: TfrmLD4Q121;

implementation

{$R *.DFM}

uses ZZfunc, LD4Q12;

procedure TfrmLD4Q121.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q12.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      Qrr_5.Lines.Clear;
      Qrr_6.Lines.Clear;
      Qrr_8.Lines.Clear;
      Qrr_9.Lines.Clear;

      // 출력자료 전달
      Qrl_1.Caption    := Cell[1, UV_iCount];
      Qrl_2.Caption    := Cell[2, UV_iCount];
      Qrl_3.Caption    := Cell[3, UV_iCount];
      Qrl_4.Caption    := Cell[4, UV_iCount];
      Qrr_5.Lines.Text := Cell[5, UV_iCount];
      Qrr_6.Lines.Text := Cell[6, UV_iCount];
      Qrl_7.Caption    := copy(Cell[7, UV_iCount],1,8);
      Qrr_8.Lines.Text := Cell[8, UV_iCount];
      Qrr_9.Lines.Text := Cell[9, UV_iCount];

      if UV_iCount <= Rows then begin MoreData := True; Inc(UV_iCount); end
      else                            MoreData := False;
   end;
end;

procedure TfrmLD4Q121.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q12 do
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

procedure TfrmLD4Q121.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
   iCnt : Integer;
begin
   inherited;

   //가장 라인수 많은 걸 찾기...
   iCnt := Qrr_5.Lines.Count;
   if iCnt < Qrr_6.Lines.Count then iCnt := Qrr_6.Lines.Count;
   if iCnt < Qrr_8.Lines.Count then iCnt := Qrr_8.Lines.Count;
   if iCnt < Qrr_9.Lines.Count then iCnt := Qrr_9.Lines.Count;

   if iCnt = 1 then
   begin
      QRBand2.Height := 25;

      Qrr_5.Height := 25;
      Qrr_6.Height := 25;
      Qrr_8.Height := 25;
      Qrr_9.Height := 25;

      Qrs_1.Height := 25;
      Qrs_2.Height := 25;
      Qrs_3.Height := 25;
      Qrs_4.Height := 25;
      Qrs_5.Height := 25;
      Qrs_6.Height := 25;
      Qrs_7.Height := 25;
      Qrs_8.Height := 25;
   end
   else
   begin
      QRBand2.Height := (iCnt * 18) + 3;

      Qrr_5.Height := (iCnt * 18) + 3;
      Qrr_6.Height := (iCnt * 18) + 3;
      Qrr_8.Height := (iCnt * 18) + 3;
      Qrr_9.Height := (iCnt * 18) + 3;

      Qrs_1.Height := (iCnt * 18) + 3;
      Qrs_2.Height := (iCnt * 18) + 3;
      Qrs_3.Height := (iCnt * 18) + 3;
      Qrs_4.Height := (iCnt * 18) + 3;
      Qrs_5.Height := (iCnt * 18) + 3;
      Qrs_6.Height := (iCnt * 18) + 3;
      Qrs_7.Height := (iCnt * 18) + 3;
      Qrs_8.Height := (iCnt * 18) + 3;
   end;
end;

end.
