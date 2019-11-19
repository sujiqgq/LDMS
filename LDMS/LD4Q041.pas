unit LD4Q041;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q041 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_C044: TQRLabel;
    QRL_C049: TQRLabel;
    QRL_E001: TQRLabel;
    QRL_E002: TQRLabel;
    QRL_E040: TQRLabel;
    QRL_T002: TQRLabel;
    QRL_E003: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape44: TQRShape;
    QRShape45: TQRShape;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRShape48: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel2: TQRLabel;
    QRL_T011: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_E015: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape4: TQRShape;
    QRShape3: TQRShape;
    QRL_T007: TQRLabel;
    QRL_T006: TQRLabel;
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
  frmLD4Q041: TfrmLD4Q041;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q04;

{$R *.DFM}

procedure TfrmLD4Q041.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q04.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_Bunju.Caption := Cell[ 2,UV_iCount];
      QRL_C044.Caption  := Cell[ 3,UV_iCount];
      QRL_C049.Caption  := Cell[ 4,UV_iCount];
      QRL_E001.Caption  := Cell[ 5,UV_iCount];
      QRL_E002.Caption  := Cell[ 6,UV_iCount];
      QRL_E003.Caption  := Cell[ 7,UV_iCount];
    
      QRL_E040.Caption  := Cell[8,UV_iCount];
      QRL_T011.Caption  := Cell[9,UV_iCount];
      QRL_T002.Caption  := Cell[10,UV_iCount];
      QRL_E015.Caption  := Cell[11,UV_iCount];
      QRL_T007.Caption  := Cell[12,UV_iCount];
      QRL_T006.Caption  := Cell[13,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q041.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
