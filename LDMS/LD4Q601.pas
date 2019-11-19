unit LD4Q601;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q601 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRLabel37: TQRLabel;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRL_S007: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Sample: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_PS007: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q601: TfrmLD4Q601;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q60;

{$R *.DFM}

procedure TfrmLD4Q601.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   //검사항목명
   if frmLD4Q60.cmbHM.Text = 'C044' then
       QRLabel37.caption :='B2-MG'
   else if frmLD4Q60.cmbHM.Text = 'C049' then
       QRLabel37.caption :='Ferritin'
   else if frmLD4Q60.cmbHM.Text = 'E001' then
       QRLabel37.caption :='T3'
   else if frmLD4Q60.cmbHM.Text = 'E002' then
       QRLabel37.caption :='T4'
   else if frmLD4Q60.cmbHM.Text = 'E003' then
       QRLabel37.caption :='TSH'
   else if frmLD4Q60.cmbHM.Text = 'E015' then
       QRLabel37.caption :='Free T4'
   else if frmLD4Q60.cmbHM.Text = 'E040' then
       QRLabel37.caption :='Cyfra'
   else if frmLD4Q60.cmbHM.Text = 'T002' then
       QRLabel37.caption :='CEA'
   else if frmLD4Q60.cmbHM.Text = 'T009' then
       QRLabel37.caption :='CA 19-9'
   else if frmLD4Q60.cmbHM.Text = 'T007' then
       QRLabel37.caption :='CA 125'
   else if frmLD4Q60.cmbHM.Text = 'T011' then
       QRLabel37.caption :='PSA'
   else if frmLD4Q60.cmbHM.Text = 'T037' then
       QRLabel37.caption :='CA 15-3'
   else if frmLD4Q60.cmbHM.Text = 'TT01' then
       QRLabel37.caption :='AFP'
   else if frmLD4Q60.cmbHM.Text = 'TT02' then
       QRLabel37.caption :='AFP수치';

   //종전양성 체크시
   if frmLD4Q60.Ck_PPlus.Checked = True then
   begin
      QRLabel4.Caption :='종전양성';
      QRL_PS007.Visible := True;
   end
   else
   begin
      QRLabel4.Caption :='비고';
      QRL_PS007.Visible := False;
   end;

   with frmLD4Q60.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption      := Cell[ 1,UV_iCount];
      QRL_Sample.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption   := Cell[ 3,UV_iCount];
      QRL_Name.Caption    := Cell[ 4,UV_iCount];
      QRL_S007.Caption    := Cell[ 5,UV_iCount];
      QRL_PS007.Caption   := Cell[ 6,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q601.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q60.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q60.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q60.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
