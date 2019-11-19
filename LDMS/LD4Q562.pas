unit LD4Q562;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q562 = class(TForm)
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
    QRLabel4: TQRLabel;
    qrl_no: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape10: TQRShape;
    QRLabel8: TQRLabel;
    Qrl_002: TQRLabel;
    Qrl_001: TQRLabel;
    QRLabel14: TQRLabel;
    QRShape24: TQRShape;
    QRL_name: TQRLabel;
    QRLabel9: TQRLabel;
    Qrl_sample: TQRLabel;
    QRShape25: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    Qrl_005: TQRLabel;
    Qrl_006: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
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
  frmLD4Q562: TfrmLD4Q562;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q56;

{$R *.DFM}

procedure TfrmLD4Q562.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
  inherited;
  with frmLD4Q56.grdMaster do
   begin

      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      qrl_no.Caption       := Cell[ 1,UV_iCount];
      Qrl_NumBunju.Caption := Cell[ 3,UV_iCount];
      QRL_sample.Caption   := Cell[ 4,UV_iCount];
      QRL_name.Caption     := Cell[ 5,UV_iCount];
      Qrl_001.Caption     := Cell[ 6,UV_iCount];
      Qrl_002.Caption     := Cell[ 7,UV_iCount];
      Qrl_003.Caption     := Cell[ 8,UV_iCount];
      Qrl_004.Caption     := Cell[ 10,UV_iCount];
      Qrl_005.Caption     := Cell[ 11,UV_iCount];
      Qrl_006.Caption     := Cell[ 12,UV_iCount];

      if frmLD4Q56.grdMaster.CellColor[6,UV_iCount] = clYellow then
          Qrl_001.Color := clYellow
      else Qrl_001.Color := clWhite;
      if frmLD4Q56.grdMaster.CellColor[7,UV_iCount] = clYellow then
          Qrl_002.Color := clYellow
      else Qrl_002.Color := clWhite;
      if frmLD4Q56.grdMaster.CellColor[8,UV_iCount] = clYellow then
          Qrl_003.Color := clYellow
      else Qrl_003.Color := clWhite;
      if frmLD4Q56.grdMaster.CellColor[10,UV_iCount] = clYellow then
          Qrl_004.Color := clYellow
      else Qrl_004.Color := clWhite;
      if frmLD4Q56.grdMaster.CellColor[11,UV_iCount] = clYellow then
          Qrl_005.Color := clYellow
      else Qrl_005.Color := clWhite;
      if frmLD4Q56.grdMaster.CellColor[12,UV_iCount] = clYellow then
          Qrl_006.Color := clYellow
      else Qrl_006.Color := clWhite;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q562.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q56.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q56.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q56.mskDate.Text,7,2) + '일';
   UV_iCount := 0;

end;

end.
