unit LD8Q101;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD8Q101 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRS_Page: TQRSysData;
    QRBand1: TQRBand;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRL_STDATE: TQRLabel;
    QRL_EDDATE: TQRLabel;
    QRL_DASH: TQRLabel;
    QRLabel17: TQRLabel;
    DetailBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel3: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel16: TQRLabel;
    QRL_part: TQRLabel;
    QRL_code: TQRLabel;
    QRL_hm: TQRLabel;
    QRL_15: TQRLabel;
    QRL_12: TQRLabel;
    QRL_11: TQRLabel;
    QRL_43: TQRLabel;
    QRL_71: TQRLabel;
    QRL_61: TQRLabel;
    QRL_51: TQRLabel;
    QRL_sum: TQRLabel;
    QRLabel49: TQRLabel;
    QRL_S15: TQRLabel;
    QRL_S12: TQRLabel;
    QRL_S11: TQRLabel;
    QRL_S43: TQRLabel;
    QRL_S71: TQRLabel;
    QRL_S61: TQRLabel;
    QRL_S51: TQRLabel;
    QRL_SSum: TQRLabel;
    QRLabel18: TQRLabel;

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
  frmLD8Q101: TfrmLD8Q101;

implementation

uses ZZFunc, ZZComm, MdmsFunc, LD8Q10;

{$R *.DFM}


procedure TfrmLD8Q101.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD8Q10.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL_part.Caption := Cell[1, UV_iCount];
      QRL_code.Caption := Cell[2, UV_iCount];
      QRL_hm.Caption   := Cell[3, UV_iCount];
      QRL_15.Caption   := Cell[4, UV_iCount];
      QRL_12.Caption   := Cell[5, UV_iCount];
      QRL_11.Caption   := Cell[6, UV_iCount];
      QRL_43.Caption   := Cell[7, UV_iCount];
      QRL_71.Caption   := Cell[8, UV_iCount];
      QRL_61.Caption   := Cell[9, UV_iCount];
      QRL_51.Caption   := Cell[10, UV_iCount];
      QRL_sum.Caption  := Cell[11, UV_iCount];

      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;

end;

procedure TfrmLD8Q101.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
Var
   tCnt15, tCnt12, tCnt11, tCnt43, tCnt71, tCnt61, tCnt51, tCnt34, tCntkita, tCntTot : Integer;
   aCnt15, aCnt12, aCnt11, aCnt43, aCnt71, aCnt61, aCnt51, aCnt34, aCntkita, aCntTot : Integer;
   UV_i : Integer;
begin
   inherited;
   tCnt15 := 0; tCnt12 := 0; tCnt11 := 0; tCnt43 := 0; tCnt71 := 0; tCnt61 := 0; tCnt51 := 0; tCnt34 := 0; tCntkita := 0; tCntTot := 0;
   aCnt15 := 0; aCnt12 := 0; aCnt11 := 0; aCnt43 := 0; aCnt71 := 0; aCnt61 := 0; aCnt51 := 0; aCnt34 := 0; aCntkita := 0; aCntTot := 0;

   with frmLD8Q10.grdMaster do
   begin
      For UV_i := 0 To Rows Do
      begin
         tCnt15   := Cell[4,  UV_i];
         tCnt12   := Cell[5,  UV_i];
         tCnt11   := Cell[6,  UV_i];
         tCnt43   := Cell[7,  UV_i];
         tCnt71   := Cell[8,  UV_i];
         tCnt61   := Cell[9,  UV_i];
         tCnt51   := Cell[10, UV_i];
         tCntTot  := Cell[11, UV_i];

         aCnt15   := aCnt15   + tCnt15;
         aCnt12   := aCnt12   + tCnt12;
         aCnt11   := aCnt11   + tCnt11;
         aCnt43   := aCnt43   + tCnt43;
         aCnt71   := aCnt71   + tCnt71;
         aCnt61   := aCnt61   + tCnt61;
         aCnt51   := aCnt51   + tCnt51;
         aCntTot  := aCntTot  + tCntTot;
      end;

      QRL_S15.Caption   := FormatFloat('#,##0',aCnt15);
      QRL_S12.Caption   := FormatFloat('#,##0',aCnt12);
      QRL_S11.Caption   := FormatFloat('#,##0',aCnt11);
      QRL_S43.Caption   := FormatFloat('#,##0',aCnt43);
      QRL_S71.Caption   := FormatFloat('#,##0',aCnt71);
      QRL_S61.Caption   := FormatFloat('#,##0',aCnt61);
      QRL_S51.Caption   := FormatFloat('#,##0',aCnt51);
      QRL_SSum.Caption  := FormatFloat('#,##0',aCntTot);
   end;

   UV_iCount := 1;

   // 조회조건을 전달
   QRL_STDATE.Caption   := Copy(frmLD8Q10.mskST_date.Text,1,4) + '/' +
                           Copy(frmLD8Q10.mskST_date.Text,5,2) + '/' +
                           Copy(frmLD8Q10.mskST_date.Text,7,2);
   QRL_EDDATE.Caption   := Copy(frmLD8Q10.mskED_date.Text,1,4) + '/' +
                           Copy(frmLD8Q10.mskED_date.Text,5,2) + '/' +
                           Copy(frmLD8Q10.mskED_date.Text,7,2);
   QRLabel17.Caption    := FormatDateTime('yyyy/mm/dd', Date);

end;
end.

