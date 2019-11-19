unit LD4Q561;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q561 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_NumBunju: TQRLabel;
    QRShape4: TQRShape;//
    QRLabel4: TQRLabel;
    qrl_no: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape10: TQRShape;
    QRLabel8: TQRLabel;
    Qrl_C083: TQRLabel;
    Qrl_C045: TQRLabel;
    QRShape14: TQRShape;
    QRLabel14: TQRLabel;
    QRShape18: TQRShape;//
    QRShape22: TQRShape;//
    QRShape23: TQRShape;//
    QRShape24: TQRShape;
    QRL_name: TQRLabel;
    QRLabel3: TQRLabel;
    Qrl_Datbunju: TQRLabel;
    QRLabel9: TQRLabel;
    Qrl_sample: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q561: TfrmLD4Q561;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q56;

{$R *.DFM}

procedure TfrmLD4Q561.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
  inherited;
  with frmLD4Q56.grdMaster do
   begin
      if (frmLD4Q56.gubn_Comb.ItemIndex = 0) then
      begin
         QRLabel1.Caption := 'Fe & Hemoglobin 결과비교';
         QRLabel8.Caption := 'Fe';
         QRLabel7.Caption := 'Hemoglobin';
      end
      else if (frmLD4Q56.gubn_Comb.ItemIndex = 1) then
      begin
         QRLabel1.Caption := 'HbA1C & Glucose 결과비교';
         QRLabel8.Caption := 'HbA1C';
         QRLabel7.Caption := 'Glucose';
      end;

      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      qrl_no.Caption       := Cell[ 1,UV_iCount];
      Qrl_DatBunju.Caption := Cell[ 2,UV_iCount];
      Qrl_NumBunju.Caption := Cell[ 3,UV_iCount];
      QRL_sample.Caption   := Cell[ 4,UV_iCount];
      QRL_name.Caption     := Cell[ 5,UV_iCount];
      Qrl_C045.Caption     := Cell[ 6,UV_iCount];
      Qrl_C083.Caption     := Cell[ 7,UV_iCount];
      if frmLD4Q56.grdMaster.CellColor[6,UV_iCount] = clYellow then
          Qrl_C045.Color := clYellow
      else Qrl_C045.Color := clWhite;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q561.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q56.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q56.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q56.mskDate.Text,7,2) + '일';
   UV_iCount := 0;

end;

end.
