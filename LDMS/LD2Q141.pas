unit LD2Q141;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q141 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_P007: TQRLabel;
    QRL_P008: TQRLabel;
    QRL_U059: TQRLabel;
    QRL_U063: TQRLabel;
    QRL_U064: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel32: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape45: TQRShape;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRLabel6: TQRLabel;
    QRShape4: TQRShape;
    QRShape6: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape18: TQRShape;
    QRShape21: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRBand4: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRShape51: TQRShape;
    QRLabel14: TQRLabel;
    QRL_samp: TQRLabel;
    QRShape5: TQRShape;
    QRShape8: TQRShape;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRShape28: TQRShape;
    v: TQRShape;
    QRShape31: TQRShape;
    QRShape36: TQRShape;
    QRL_Name: TQRLabel;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRL_U065: TQRLabel;
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
  frmLD2Q141: TfrmLD2Q141;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q14;

{$R *.DFM}

procedure TfrmLD2Q141.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q14.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption    := copy(Cell[  1,UV_iCount],1, pos('/',Cell[  1,UV_iCount])-1);
//      QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      if Length(Cell[ 4,UV_iCount]) >= 8 then
         QRL_Name.Caption  :=  copy(Cell[ 4,UV_iCount],1,8)
      else
         QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_Jumin.Caption := Cell[ 5,UV_iCount];

      QRL_P007.Caption  := Cell[ 6,UV_iCount];
      QRL_P008.Caption  := Cell[ 7,UV_iCount];
      QRL_U059.Caption  := Cell[ 8,UV_iCount];
      QRL_U063.Caption  := Cell[ 9,UV_iCount];
      QRL_U064.Caption  := Cell[10,UV_iCount];
      QRL_U065.Caption  := Cell[11,UV_iCount];


      QRL_samp.Caption       := Cell[13,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q141.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q14.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q14.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q14.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
