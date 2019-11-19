unit LD6I021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD6I021 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape21: TQRShape;
    QRShape5: TQRShape;
    QRLabel10: TQRLabel;
    QRShape10: TQRShape;
    QRLabel11: TQRLabel;
    QRShape4: TQRShape;
    QRShape11: TQRShape;
    qrl_No: TQRLabel;
    qrl_Ims_low: TQRLabel;
    qrl_Ims_high: TQRLabel;
    QRShape17: TQRShape;
    QRShape20: TQRShape;
    qrl_Prev_rslt: TQRLabel;
    qrl_Cur_rslt: TQRLabel;
    QRShape9: TQRShape;
    QRLabel6: TQRLabel;
    QRShape16: TQRShape;
    qrl_desc_kor: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    qrl_desc_name: TQRLabel;
    Qrl_Bunju: TQRLabel;
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
  frmLD6I021: TfrmLD6I021;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD6I02 ;

{$R *.DFM}

procedure TfrmLD6I021.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD6I02.grdResult do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      Inc(UV_iCount);

      Qrl_No.caption         := Cell[1,UV_iCount];

      qrl_desc_kor.Caption   := Cell[2,UV_iCount];
      qrl_Ims_low.Caption    := Cell[3,UV_iCount];
      qrl_Ims_high.Caption   := Cell[4,UV_iCount];
      qrl_Prev_rslt.Caption  := Cell[5,UV_iCount];
      qrl_Cur_rslt.Caption   := Cell[6,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD6I021.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD6I02.mskDAT_BUNJU.Text,1,4) + '년'  +
                       copy(frmLD6I02.mskDAT_BUNJU.Text,5,2) + '월'  +
                       copy(frmLD6I02.mskDAT_BUNJU.Text,7,2) + '일';

   qrl_desc_name.Caption  := frmLD6I02.edtDESC_NAME.text;
   Qrl_Bunju.Caption      := frmLD6I02.edtNUM_BUNJU.text;
   UV_iCount := 0;
end;

end.
