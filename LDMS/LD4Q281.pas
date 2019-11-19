unit LD4Q281;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q281 = class(TForm)
    QuickRep: TQuickRep;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape11: TQRShape;
    QRLabel8: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRLabel19: TQRLabel;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape32: TQRShape;
    QRShape35: TQRShape;
    QRShape2: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape9: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape20: TQRShape;
    QRShape16: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRL_Bunju: TQRLabel;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRL_Name: TQRLabel;
    QRShape24: TQRShape;
    QRL_No: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_C022: TQRLabel;
    QRShape21: TQRShape;
    QRShape8: TQRShape;
    QRShape14: TQRShape;
    QRShape22: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRL_C023: TQRLabel;
    QRL_C033: TQRLabel;
    QRL_C034: TQRLabel;
    QRL_C080: TQRLabel;
    QRL_C074: TQRLabel;
    QRL_C078: TQRLabel;
    QRL_C082: TQRLabel;
    QRL_C090: TQRLabel;
    QRL_GL02: TQRLabel;
    QRL_GL03: TQRLabel;
    QRL_GL04: TQRLabel;
    QRL_GL05: TQRLabel;
    QRL_GL06: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRShape23: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape1: TQRShape;
    QRShape17: TQRShape;
    QRShape31: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape36: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRL_S036: TQRLabel;
    QRLabel31: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape44: TQRShape;
    QRShape45: TQRShape;
    QRLabel30: TQRLabel;
    QRLabel32: TQRLabel;
    QRL_C084: TQRLabel;
    QRL_C085: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
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
  frmLD4Q281: TfrmLD4Q281;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q28;

{$R *.DFM}

procedure TfrmLD4Q281.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   if  frmLD4Q28.RB_01.Checked then
       QRLabel11.caption:='샘플번호'
   else QRLabel11.caption:='분주번호';

   with frmLD4Q28.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      if frmLD4Q28.Chk_noHomo.checked = true then
      begin
      QRL_No.Caption    := Cell[  1, UV_iCount];
      QRL_Bunju.Caption := Cell[  2, UV_iCount];
      QRL_Name.Caption  := Cell[  3, UV_iCount];
      QRL_C022.Caption  := Cell[  4, UV_iCount];
      QRL_C023.Caption  := Cell[  5, UV_iCount];
      QRL_C080.Caption  := Cell[  6, UV_iCount];
      QRL_C033.Caption  := Cell[  7, UV_iCount];
      QRL_C034.Caption  := Cell[  8, UV_iCount];
      QRL_C074.Caption  := Cell[  9, UV_iCount];


      QRL_S036.Caption  := '';
      QRLabel10.Caption := '';
      QRLabel21.Caption := '';

      QRL_C082.Caption  := Cell[ 13, UV_iCount];
      QRL_C084.Caption  := Cell[ 11, UV_iCount];
      QRL_C085.Caption  := Cell[ 12, UV_iCount];
      QRL_C090.Caption  := Cell[ 14, UV_iCount];
      QRL_GL02.Caption  := Cell[ 15, UV_iCount];
      QRL_GL03.Caption  := Cell[ 16, UV_iCount];
      QRL_GL04.Caption  := Cell[ 17, UV_iCount];
      QRL_GL05.Caption  := Cell[ 18, UV_iCount];
      QRL_GL06.Caption  := Cell[ 19, UV_iCount];
      QRL_C078.Caption  := Cell[ 10, UV_iCount];
      end
      else
      begin
      QRL_No.Caption    := Cell[  1, UV_iCount];
      QRL_Bunju.Caption := Cell[  2, UV_iCount];
      QRL_Name.Caption  := Cell[  3, UV_iCount];
      QRL_C022.Caption  := Cell[  4, UV_iCount];
      QRL_C023.Caption  := Cell[  5, UV_iCount];
      QRL_C080.Caption  := Cell[  6, UV_iCount];
      QRL_C033.Caption  := Cell[  7, UV_iCount];
      QRL_C034.Caption  := Cell[  8, UV_iCount];
      QRL_C074.Caption  := Cell[  9, UV_iCount];
      QRL_S036.Caption  := Cell[ 10, UV_iCount];
      QRL_C082.Caption  := Cell[ 14, UV_iCount];
      QRL_C084.Caption  := Cell[ 12, UV_iCount];
      QRL_C085.Caption  := Cell[ 13, UV_iCount];
      QRL_C090.Caption  := Cell[ 15, UV_iCount];
      QRL_GL02.Caption  := Cell[ 16, UV_iCount];
      QRL_GL03.Caption  := Cell[ 17, UV_iCount];
      QRL_GL04.Caption  := Cell[ 18, UV_iCount];
      QRL_GL05.Caption  := Cell[ 19, UV_iCount];
      QRL_GL06.Caption  := Cell[ 20, UV_iCount];
      QRL_C078.Caption  := Cell[ 11, UV_iCount];
      end;
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q281.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q28.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q28.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q28.mskDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
