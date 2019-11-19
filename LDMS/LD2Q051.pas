unit LD2Q051;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q051 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_E001: TQRLabel;
    QRL_E002: TQRLabel;
    QRL_E003: TQRLabel;
    QRL_E040: TQRLabel;
    QRL_C044: TQRLabel;
    QRL_C049: TQRLabel;
    QRL_T001: TQRLabel;
    QRL_E015: TQRLabel;
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
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape45: TQRShape;
    QRLabel33: TQRLabel;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRShape48: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel2: TQRLabel;
    QRL_T011: TQRLabel;
    QRShape4: TQRShape;
    QRLabel7: TQRLabel;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_T002: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel21: TQRLabel;
    QRL_Jisa: TQRLabel;
    QRLabel20: TQRLabel;
    QRL_Desc_name: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRShape44: TQRShape;
    QRShape49: TQRShape;
    QRShape50: TQRShape;
    QRBand4: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel17: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel41: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_TT01: TQRLabel;
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
  frmLD2Q051: TfrmLD2Q051;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q05;

{$R *.DFM}

procedure TfrmLD2Q051.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q05.grdMaster do
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
      QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_Jumin.Caption := Cell[ 5,UV_iCount];
      QRL_E001.Caption  := Cell[ 6,UV_iCount];
      QRL_E002.Caption  := Cell[ 7,UV_iCount];
      QRL_E003.Caption  := Cell[ 8,UV_iCount];
      QRL_E015.Caption  := Cell[ 9,UV_iCount];
      QRL_E040.Caption  := Cell[10,UV_iCount];
      QRL_C044.Caption  := Cell[11,UV_iCount];
      QRL_C049.Caption  := Cell[12,UV_iCount];
      QRL_T001.Caption  := Cell[13,UV_iCount];
      QRL_T002.Caption  := Cell[14,UV_iCount];
      QRL_T011.Caption  := Cell[15,UV_iCount];
      QRL_TT01.Caption  := Cell[16,UV_iCount];
      QRL_Desc_name.Caption  := Cell[17,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q051.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q05.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q05.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q05.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
