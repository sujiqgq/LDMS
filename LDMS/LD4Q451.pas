unit LD4Q451;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q451 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand2: TQRBand;
    QRBand3: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape1: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    qrlDesc_RackNo: TQRLabel;
    qrlNum_samp: TQRLabel;
    qrlDesc_Name: TQRLabel;
    qrlNum_bunju: TQRLabel;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRLabel8: TQRLabel;
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
  frmLD4Q451: TfrmLD4Q451;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q45;

{$R *.DFM}

procedure TfrmLD4Q451.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q45.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;


      UV_iCount := UV_iCount + 1;

      qrlDesc_RackNo.Caption := Cell[1,UV_iCount];
      qrlNum_samp.Caption    := Cell[2,UV_iCount];
      qrlDesc_Name.Caption   := Cell[3,UV_iCount];
      qrlNum_bunju.Caption   := Cell[4,UV_iCount];
      QRLabel8.Caption   :=  IntToStr(UV_iCount);

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q451.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q45.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q45.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q45.mskDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
