//==============================================================================
// 프로그램 설명 : 특수건강진단 분석 시료 외부 의뢰서 출력 폼
// 시스템        : 통합검진
// 수정일자      : 2016.09.01
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q271;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD4Q271 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel11: TQRLabel;
    QRBand2: TQRBand;
    QRLabel10: TQRLabel;
    QRLabel14: TQRLabel;
    QRBand1: TQRBand;
    QRS_Page: TQRSysData;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel4: TQRLabel;
    QRL_1: TQRLabel;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRShape10: TQRShape;
    QRLabel16: TQRLabel;
    QRL_3: TQRLabel;
    QRShape11: TQRShape;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel25: TQRLabel;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRShape_1: TQRShape;
    QRL3: TQRLabel;
    QRL10: TQRLabel;
    QRL11: TQRLabel;
    QRShape_2: TQRShape;
    QRShape_7: TQRShape;
    QRShape_8: TQRShape;
    QRShape_9: TQRShape;
    QRShape_11: TQRShape;
    qryTemp: TQuery;
    QRShape22: TQRShape;
    QRLabel3: TQRLabel;
    QRL_2: TQRLabel;
    QRL_5: TQRLabel;
    QRL_6: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel35: TQRLabel;
    QRL8: TQRLabel;
    QRL9: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_8: TQRLabel;
    QRL_7: TQRLabel;
    QRShape17: TQRShape;
    QRShape3: TQRShape;
    QRShape9: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRLabel13: TQRLabel;
    QRL_4: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape25: TQRShape;
    QRLabel9: TQRLabel;
    QRShape26: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape27: TQRShape;
    QRL4: TQRLabel;
    QRShape28: TQRShape;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel24: TQRLabel;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRL5: TQRLabel;
    QRL6: TQRLabel;
    QRL7: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
    UV_sSumCnt, UV_sSumSuga : string;

    Procedure UP_Clear;

  public
    { Public declarations }
  end;

var
  frmLD4Q271: TfrmLD4Q271;

implementation

{$R *.DFM}

uses ZZfunc, LD4Q27;

Procedure TfrmLD4Q271.UP_Clear;
begin
   QRL1.Caption  := '';
   QRL2.Caption  := '';
   QRL3.Caption  := '';
   QRL4.Caption  := '';
   QRL5.Caption  := '';
   QRL6.Caption  := '';
   QRL7.Caption  := '';
   QRL9.Caption  := '';
   QRL10.Caption := '';
   QRL11.Caption := '';
end;

procedure TfrmLD4Q271.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q27.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      UP_Clear;

      if Copy(frmLD4Q27.Comb_Hm.Text,1,4) = 'SE61' then
      begin
         QRLabel22.Caption := '결과(mg/g crea.)'+ #13 +'(≤1.000)';
         QRLabel22.Font.Size := 7;
      end
      else if Copy(frmLD4Q27.Comb_Hm.Text,1,4) = 'SE85' then
      begin
         QRLabel22.Caption := '결과(㎍/L)'+ #13 +'(≤200.0㎍/L)';
         QRLabel22.Font.Size := 7;
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


      if UV_iCount <= Rows  then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q271.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q27 do
   begin
      //위탁기관
      QRL_1.Caption := Edt_Ekikwan.Text;
      //수탁기관
      QRL_2.Caption := Edt_Skikwan.Text;
      //수탁담당자
      if Trim(Edt_Sdangdam.Text) <> '' then QRL_3.Caption := Edt_Sdangdam.Text + '   (인)'
      else                                  QRL_3.Caption := '             (인)';
      //수탁연락처
      QRL_4.Caption := Edt_STel.Text;
      //분석책임자
      QRL_5.Caption := Edt_EManager.Text + '   (인)';
      //분석확인자
      QRL_6.Caption := Edt_Edangdam.Text + '   (인)';
      //의뢰일자
      QRL_7.Caption := GF_DateFormat(DEdt_EDate.Text);
      //의뢰연락처
      QRL_8.Caption := Edt_ETel.Text;

      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

end.
