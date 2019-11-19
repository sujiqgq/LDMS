unit LD8Q011;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD8Q011 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRBand2: TQRBand;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    Qrl_Bunju1: TQRLabel;
    Qrl_Bunju2: TQRLabel;
    QRLabel3: TQRLabel;
    qrl_bunju3: TQRLabel;
    QRShape10: TQRShape;
    QRLabel4: TQRLabel;
    QRShape18: TQRShape;
    Qrl_bunju4: TQRLabel;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape9: TQRShape;
    QRShape14: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel10: TQRLabel;
    QRL_No1: TQRLabel;
    QRL_No2: TQRLabel;
    QRL_No3: TQRLabel;
    QRL_No4: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow, No, count : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD8Q011: TfrmLD8Q011;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q01;

{$R *.DFM}

procedure TfrmLD8Q011.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD8Q01.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      if frmLD8Q01.cmbGubn.itemindex = 0 then qrlabel1.caption := 'H.pylori 분 주 표'
      else if frmLD8Q01.cmbGubn.itemindex = 1 then qrlabel1.caption := 'Free T4 분 주 표'
      else if frmLD8Q01.cmbGubn.itemindex = 2 then qrlabel1.caption := 'ESR 분 주 표'
      else if frmLD8Q01.cmbGubn.itemindex = 3 then qrlabel1.caption := '요중마뇨산(SE24) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 4 then qrlabel1.caption := '요중만델릭산(SE50) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 5 then qrlabel1.caption := '요중메틸마뇨산(SE53) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 6 then qrlabel1.caption := '요중NMF(SE57) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 7 then qrlabel1.caption := 'SCC(T009) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 8 then qrlabel1.caption := 'Rubella IgG (여자)(S019) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 9 then qrlabel1.caption := 'B2-MG(C044) 분주표'
      else if frmLD8Q01.cmbGubn.itemindex = 10 then qrlabel1.caption := 'Free T3(E041) 분주표';

      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 30 Then
      Begin
         iRow := iRow + 4;
         cRow := 1;
         UV_iCount := UV_iCount + 1;
         count := 0;
      End;
      Qrl_Bunju1.Caption := Cell[2,(iRow * 30) + cRow];
      Qrl_Bunju2.Caption := Cell[2,(iRow * 30) + cRow + 30];
      Qrl_Bunju3.Caption := Cell[2,(iRow * 30) + cRow + 60];
      Qrl_Bunju4.Caption := Cell[2,(iRow * 30) + cRow + 90];

      count := count + 1;
      Qrl_No1.caption := intTostr(count);
      Qrl_No2.caption := intTostr(count + 30);
      Qrl_No3.caption := intTostr(count + 60);
      Qrl_No4.caption := intTostr(count + 90);

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD8Q011.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD8Q01.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD8Q01.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD8Q01.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;

  { sRec := Round(frmMD6Q85.grdMaster.Rows / 120);
   if (((frmMD6Q85.grdMaster.Rows / 120) - Round(frmMD6Q85.grdMaster.Rows / 120)) <= 0.5) and
      (((frmMD6Q85.grdMaster.Rows / 120) - Round(frmMD6Q85.grdMaster.Rows / 120)) > 0)    and
      (frmMD6Q85.grdMaster.Rows > 120) then
       sRec := sRec + 1; }
   if frmLD8Q01.grdMaster.Rows < 120 then sRec := 0
   else if frmLD8Q01.grdMaster.Rows < 240 then sRec := 1
   else if frmLD8Q01.grdMaster.Rows < 360 then sRec := 2
   else if frmLD8Q01.grdMaster.Rows < 480 then sRec := 3
   else if frmLD8Q01.grdMaster.Rows < 600 then sRec := 4
   else if frmLD8Q01.grdMaster.Rows < 720 then sRec := 5
   else if frmLD8Q01.grdMaster.Rows < 840 then sRec := 6
   else if frmLD8Q01.grdMaster.Rows < 960 then sRec := 7
   else if frmLD8Q01.grdMaster.Rows < 1080 then sRec := 8
   else if frmLD8Q01.grdMaster.Rows < 1200 then sRec := 9
   else if frmLD8Q01.grdMaster.Rows < 1320 then sRec := 10  
   else if frmLD8Q01.grdMaster.Rows < 1440 then sRec := 11
   else if frmLD8Q01.grdMaster.Rows < 1560 then sRec := 12
   else if frmLD8Q01.grdMaster.Rows < 1680 then sRec := 13
   else if frmLD8Q01.grdMaster.Rows < 1800 then sRec := 14
   else if frmLD8Q01.grdMaster.Rows < 1920 then sRec := 15
   else if frmLD8Q01.grdMaster.Rows < 2040 then sRec := 16
   else if frmLD8Q01.grdMaster.Rows < 2160 then sRec := 17
   else if frmLD8Q01.grdMaster.Rows < 2280 then sRec := 18
   else if frmLD8Q01.grdMaster.Rows < 2400 then sRec := 19
   else if frmLD8Q01.grdMaster.Rows < 2520 then sRec := 20;

   cRow := 0; iRow := 0; count := 0;
end;

end.
