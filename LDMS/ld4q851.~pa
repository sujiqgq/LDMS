unit ld4q851;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q851 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRL_Name: TQRLabel;
    QRShape5: TQRShape;
    QRL_001: TQRLabel;
    QRShape21: TQRShape;
    QRShape14: TQRShape;
    QRShape22: TQRShape;
    QRShape38: TQRShape;
    QRL_002: TQRLabel;
    QRL_003: TQRLabel;
    QRL_004: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape11: TQRShape;
    QRLabel8: TQRLabel;
    QRShape18: TQRShape;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape2: TQRShape;
    QRLabel11: TQRLabel;
    QRShape9: TQRShape;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape16: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel4: TQRLabel;
    QRShape30: TQRShape;
    QRLabel2: TQRLabel;
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
  frmLD4Q851: TfrmLD4Q851;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q85;

{$R *.DFM}

procedure TfrmLD4Q851.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   if  frmLD4Q85.RB_01.Checked then
       QRLabel11.caption:='샘플번호'
   else QRLabel11.caption:='분주번호';

   with frmLD4Q85.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      begin
      QRL_Bunju.Caption := Cell[  1, UV_iCount];
      QRL_Name.Caption  := Cell[  2, UV_iCount];
      QRL_001.Caption   := Cell[  3, UV_iCount];
      QRL_002.Caption   := Cell[  4, UV_iCount];
      QRL_003.Caption   := Cell[  5, UV_iCount];
      QRL_004.Caption   := Cell[  6, UV_iCount];
      end;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q851.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q85.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q85.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q85.mskDate.Text,7,2) + '일';
   UV_iCount := 0;
end;
end.
