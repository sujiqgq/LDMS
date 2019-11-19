unit LD4Q461;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q461 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRSysData2: TQRSysData;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape22: TQRShape;
    QRLabel7: TQRLabel;
    QRShape23: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel1: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    Qrl_no: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    qrl_jisa: TQRLabel;
    qrl_samp: TQRLabel;
    QRShape21: TQRShape;
    QRShape4: TQRShape;
    QRShape9: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRShape10: TQRShape;
    QRLabel6: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel10: TQRLabel;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    qrl_name: TQRLabel;
    qrl_jumin: TQRLabel;
    qrl_dat_gmgn: TQRLabel;
    qrl_blood: TQRLabel;
    qrl_urin: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel5: TQRLabel;
    QRL_date1: TQRLabel;
    QRLabel12: TQRLabel;
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
  frmLD4Q461: TfrmLD4Q461;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q46;

{$R *.DFM}

procedure TfrmLD4Q461.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD4Q46.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;
      Qrl_no.Caption      := Cell[1,UV_iCount];
      Qrl_jisa.Caption    := Cell[2,UV_iCount];
      Qrl_samp.Caption    := Cell[3,UV_iCount];
      Qrl_name.Caption    := Cell[4,UV_iCount];
      Qrl_jumin.Caption   := Cell[5,UV_iCount];
      Qrl_dat_gmgn.Caption:= Cell[6,UV_iCount];
      Qrl_blood.Caption   := Cell[7,UV_iCount];
      Qrl_urin.Caption    := Cell[8,UV_iCount];
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q461.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   QRL_date.Caption := copy(frmLD4Q46.mskST_date.Text,1,4) + '년'  +
                       copy(frmLD4Q46.mskST_date.Text,5,2) + '월'  +
                       copy(frmLD4Q46.mskST_date.Text,7,2) + '일';

   QRL_date1.Caption := copy(frmLD4Q46.mskEd_date.Text,1,4) + '년'  +
                       copy(frmLD4Q46.mskEd_date.Text,5,2) + '월'  +
                       copy(frmLD4Q46.mskEd_date.Text,7,2) + '일';

   UV_iCount := 0;


end;

end.
