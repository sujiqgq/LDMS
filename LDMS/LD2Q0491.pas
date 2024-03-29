unit LD2Q0491;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q0491 = class(TForm)
    QuickRep: TQuickRep;
    QRBand4: TQRBand;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel27: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape45: TQRShape;
    QRShape47: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape18: TQRShape;
    QRLabel19: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel9: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRLabel14: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_U019: TQRLabel;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape21: TQRShape;
    QRL_No: TQRLabel;
    QRShape51: TQRShape;
    QRL_samp: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape4: TQRShape;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRLabel15: TQRLabel;
    QRShape8: TQRShape;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRL_C018: TQRLabel;
    QRL_E010: TQRLabel;
    QRL_S023: TQRLabel;
    QRL_S038: TQRLabel;
    QRL_S040: TQRLabel;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape36: TQRShape;
    QRL_S053: TQRLabel;
    QRL_S054: TQRLabel;
    QRL_SE38: TQRLabel;
    QRL_T015: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRShape26: TQRShape;
    QRLabel23: TQRLabel;
    QRShape40: TQRShape;
    QRL_S029: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel28: TQRLabel;
    QRL_Y005: TQRLabel;
    QRL_Y008: TQRLabel;
    QRShape9: TQRShape;
    QRShape25: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRShape41: TQRShape;
    QRShape44: TQRShape;
    QRShape46: TQRShape;
    QRShape48: TQRShape;
    QRLabel34: TQRLabel;
    QRL_T051: TQRLabel;
    QRL_T052: TQRLabel;
    QRL_T053: TQRLabel;
    QRL_T054: TQRLabel;
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
  frmLD2Q0491: TfrmLD2Q0491;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q0491.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q04.grdMaster do
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
   //   QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      if Length(Cell[ 4,UV_iCount]) >= 8 then
         QRL_Name.Caption  :=  copy(Cell[ 4,UV_iCount],1,8)
      else
         QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_Jumin.Caption := Cell[ 5,UV_iCount];

      QRL_U019.Caption  := Cell[109,UV_iCount];
      QRL_C018.caption  := Cell[110,UV_iCount];
      QRL_E010.caption  := Cell[111,UV_iCount];
      QRL_S023.caption  := Cell[112,UV_iCount];
      QRL_S029.caption  := Cell[113,UV_iCount];
      QRL_S038.caption  := Cell[114,UV_iCount];
      QRL_S040.caption  := Cell[115,UV_iCount];
      QRL_S053.caption  := Cell[116,UV_iCount];
      QRL_S054.caption  := Cell[117,UV_iCount];
      QRL_SE38.caption  := Cell[118,UV_iCount];
      QRL_T015.caption  := Cell[119,UV_iCount];
      QRL_Y005.caption  := Cell[120,UV_iCount];
      QRL_Y008.caption  := Cell[121,UV_iCount];

      QRL_T051.caption  := Cell[126,UV_iCount];
      QRL_T052.caption  := Cell[127,UV_iCount];
      QRL_T053.caption  := Cell[128,UV_iCount];
      QRL_T054.caption  := Cell[129,UV_iCount];

    //  QRL_Desc_name.Caption  := Cell[115,UV_iCount];
      QRL_samp.Caption       := Cell[147,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q0491.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;


end.
