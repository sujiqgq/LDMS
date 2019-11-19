//==============================================================================
// 프로그램 설명 : 특수건강진단 분석 비용 내역[명단]
// 시스템        : 통합검진
// 수정일자      : 2016.08.12
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q201;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD4Q201 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel12: TQRLabel;
    qrlFrom: TQRLabel;
    QRLabel11: TQRLabel;
    QRBand2: TQRBand;
    QRLabel10: TQRLabel;
    QRLabel14: TQRLabel;
    qrlTo: TQRLabel;
    QRLabel5: TQRLabel;
    QRBand1: TQRBand;
    QRS_Page: TQRSysData;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape9: TQRShape;
    QRLabel15: TQRLabel;
    QRShape10: TQRShape;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape11: TQRShape;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRL3: TQRLabel;
    QRL5: TQRLabel;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRShape_1: TQRShape;
    QRL7: TQRLabel;
    QRL11: TQRLabel;
    QRL12: TQRLabel;
    QRShape_2: TQRShape;
    QRShape_3: TQRShape;
    QRShape_4: TQRShape;
    QRShape_5: TQRShape;
    QRShape_6: TQRShape;
    QRShape_7: TQRShape;
    QRShape_8: TQRShape;
    QRShape_10: TQRShape;
    QRShape_9: TQRShape;
    QRShape_11: TQRShape;
    QRBand3: TQRBand;
    QRRT4: TQRRichText;
    QRRT6: TQRRichText;
    QRRT8: TQRRichText;
    QRRT9: TQRRichText;
    QRRT10: TQRRichText;
    qryTemp: TQuery;
    QRShape22: TQRShape;
    QRLSum: TQRLabel;
    QRShape33: TQRShape;
    QRLabel37: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
    UV_sSum : string;

    Procedure UP_Clear;

  public
    { Public declarations }
  end;

var
  frmLD4Q201: TfrmLD4Q201;

implementation

{$R *.DFM}

uses ZZfunc, LD4Q20;

Procedure TfrmLD4Q201.UP_Clear;
begin
   QRL1.Caption  := '';
   QRL2.Caption  := '';
   QRL3.Caption  := '';
   QRL5.Caption  := '';
   QRL7.Caption  := '';
   QRL11.Caption := '';
   QRL12.Caption := '';

   QRRT4.Lines.Clear;
   QRRT6.Lines.Clear;
   QRRT8.Lines.Clear;
   QRRT9.Lines.Clear;
   QRRT10.Lines.Clear;
end;

procedure TfrmLD4Q201.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q20.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      UP_Clear;

      if UV_iCount < Rows then
      begin
         // 출력자료 전달
         QRL1.Caption := Cell[1,   UV_iCount];
         QRL2.Caption := Cell[2,   UV_iCount];
         QRL3.Caption := Cell[3,   UV_iCount];
         QRRT4.Lines.Add(Cell[4,   UV_iCount]);
         QRL5.Caption := Cell[5,   UV_iCount];
         QRRT6.Lines.Add(Cell[6,   UV_iCount]);
         QRL7.Caption := Cell[7,   UV_iCount];
         QRRT8.Lines.Add(Cell[8,   UV_iCount]);
         QRRT9.Lines.Add(Cell[9,   UV_iCount]);
         QRRT10.Lines.Add(Cell[10, UV_iCount]);
         QRL11.Caption := Cell[11, UV_iCount];
         QRL12.Caption := Cell[12, UV_iCount];
      end
      else if UV_iCount = Rows then
      begin
         UV_sSum := Cell[12, UV_iCount];
      end;


      if UV_iCount < Rows  then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q201.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   UV_sSum := '';
   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q20 do
   begin
      qrlFrom.Caption  := GF_DateFormat(msksDate.Text);
      qrlTo.Caption    := GF_DateFormat(mskeDate.Text);
      QRL_Date.Caption := GV_sPrintToday;

      //수탁의뢰기관
      qryTemp.Active := False;
      qryTemp.Sql.Text := ' SELECT * FROM JISA WHERE COD_JISA = ''' + copy(cmbCOD_JISA.Text,1,2) + ''' ';
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then
      begin
         if pos('(재)한국의학연구소',qryTemp.FieldByName('desc_hlbw').AsString) > 0 then QRLabel17.Caption := qryTemp.FieldByName('desc_hlbw').AsString
         else                                                                            QRLabel17.Caption := '(재)한국의학연구소 ' + qryTemp.FieldByName('desc_hlbw').AsString;
      end;
   end;
end;

procedure TfrmLD4Q201.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
   inherited;
   QRLSum.Caption := UV_sSum;
end;

procedure TfrmLD4Q201.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
   iCnt : Integer;
begin
   inherited;
   //가장 라인수 많은 걸 찾기...
   iCnt := QRRT4.Lines.Count;
   if iCnt < QRRT6.Lines.Count  then iCnt := QRRT6.Lines.Count;
   if iCnt < QRRT8.Lines.Count  then iCnt := QRRT8.Lines.Count;
   if iCnt < QRRT9.Lines.Count  then iCnt := QRRT9.Lines.Count;
   if iCnt < QRRT10.Lines.Count then iCnt := QRRT10.Lines.Count;

   if iCnt = 1 then
   begin
      QRBand2.Height := 25;

      QRRT4.Height  := 25;
      QRRT6.Height  := 25;
      QRRT8.Height  := 25;
      QRRT9.Height  := 25;
      QRRT10.Height := 25;

      QRShape_1.Height  := 25;
      QRShape_2.Height  := 25;
      QRShape_3.Height  := 25;
      QRShape_4.Height  := 25;
      QRShape_5.Height  := 25;
      QRShape_6.Height  := 25;
      QRShape_7.Height  := 25;
      QRShape_8.Height  := 25;
      QRShape_9.Height  := 25;
      QRShape_10.Height := 25;
      QRShape_11.Height := 25;
   end
   else
   begin
      QRBand2.Height := (iCnt * 13) + 10;

      QRRT4.Height  := (iCnt * 13) + 10;
      QRRT6.Height  := (iCnt * 13) + 10;
      QRRT8.Height  := (iCnt * 13) + 10;
      QRRT9.Height  := (iCnt * 13) + 10;
      QRRT10.Height := (iCnt * 13) + 10;

      QRShape_1.Height  := (iCnt * 13) + 10;
      QRShape_2.Height  := (iCnt * 13) + 10;
      QRShape_3.Height  := (iCnt * 13) + 10;
      QRShape_4.Height  := (iCnt * 13) + 10;
      QRShape_5.Height  := (iCnt * 13) + 10;
      QRShape_6.Height  := (iCnt * 13) + 10;
      QRShape_7.Height  := (iCnt * 13) + 10;
      QRShape_8.Height  := (iCnt * 13) + 10;
      QRShape_9.Height  := (iCnt * 13) + 10;
      QRShape_10.Height := (iCnt * 13) + 10;
      QRShape_11.Height := (iCnt * 13) + 10;
   end;
end;

end.
