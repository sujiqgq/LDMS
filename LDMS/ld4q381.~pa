unit LD4Q381;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q381 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape15: TQRShape;
    QRLabel8: TQRLabel;
    QRL_Name: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape32: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRLabel38: TQRLabel;
    QRShape40: TQRShape;
    QRShape44: TQRShape;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRShape2: TQRShape;
    QRLabel4: TQRLabel;
    QRL_T042: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_S010: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRL_S012: TQRLabel;
    QRShape14: TQRShape;
    QRShape3: TQRShape;
    QRLabel7: TQRLabel;
    QRL_T041: TQRLabel;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRLabel10: TQRLabel;
    QRL_S011: TQRLabel;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRLabel3: TQRLabel;
    QRL_TT03: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_S033: TQRLabel;
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
  frmLD4Q381: TfrmLD4Q381;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q38;

{$R *.DFM}

procedure TfrmLD4Q381.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   if  frmLD4Q38.Cmb_gubn.ItemIndex = 1 then
       QRLabel11.caption:='샘플번호'
   else QRLabel11.caption:='분주번호';
   with frmLD4Q38.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;


      QRL_No.Caption    := Cell[ 1,UV_iCount];
      QRL_Bunju.Caption := Cell[ 2,UV_iCount];
      QRL_Name.Caption  := Cell[ 3,UV_iCount];
      QRL_S010.Caption  := Cell[ 4,UV_iCount];
      QRL_T041.Caption  := Cell[ 5,UV_iCount];
      QRL_S011.Caption  := Cell[ 6,UV_iCount];
      QRL_S012.Caption  := Cell[ 7,UV_iCount];
      QRL_TT03.Caption  := Cell[ 8,UV_iCount];
      QRL_T042.Caption  := Cell[ 9,UV_iCount];
      QRL_S033.Caption  := Cell[10,UV_iCount];


      if frmLD4Q38.mskDate.Text >= '20140901' then
      begin
         QRLabel2.Caption := '';
         QRL_T042.Caption := '';
      end;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q381.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q38.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q38.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q38.mskDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.

