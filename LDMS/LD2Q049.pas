unit LD2Q049;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q049 = class(TForm)
    QuickRep: TQuickRep;
    QRBand4: TQRBand;
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
    QRL_E063: TQRLabel;
    QRL_E064: TQRLabel;
    QRL_E065: TQRLabel;
    QRL_T045: TQRLabel;
    QRL_T046: TQRLabel;
    QRL_T044: TQRLabel;
    QRShape2: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_T048: TQRLabel;
    QRL_T047: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape21: TQRShape;
    QRL_No: TQRLabel;
    QRShape29: TQRShape;
    QRL_E066: TQRLabel;
    QRL_E067: TQRLabel;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRL_U048: TQRLabel;
    QRShape51: TQRShape;
    QRL_samp: TQRLabel;
    QRL_S092: TQRLabel;
    QRL_T050: TQRLabel;
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
    QRLabel32: TQRLabel;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape45: TQRShape;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRShape48: TQRShape;
    QRShape1: TQRShape;
    QRLabel6: TQRLabel;
    QRShape4: TQRShape;
    QRLabel7: TQRLabel;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape18: TQRShape;
    QRLabel19: TQRLabel;
    QRShape28: TQRShape;
    QRLabel22: TQRLabel;
    QRLabel28: TQRLabel;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape44: TQRShape;
    QRLabel39: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel40: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape5: TQRShape;
    QRLabel20: TQRLabel;
    QRShape8: TQRShape;
    QRL_C053: TQRLabel;
    QRShape24: TQRShape;
    QRL_C055: TQRLabel;
    QRShape30: TQRShape;
    QRL_T049: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel21: TQRLabel;
    QRShape31: TQRShape;
    QRShape19: TQRShape;
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
  frmLD2Q049: TfrmLD2Q049;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q049.QuickRepNeedData(Sender: TObject;
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
     // QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      if Length(Cell[ 4,UV_iCount]) >= 8 then
         QRL_Name.Caption  :=  copy(Cell[ 4,UV_iCount],1,8)
      else
         QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_Jumin.Caption := Cell[ 5,UV_iCount];

      QRL_E063.Caption  := Cell[92,UV_iCount];
      QRL_E064.Caption  := Cell[93,UV_iCount];
      QRL_E065.Caption  := Cell[94,UV_iCount];
      QRL_T044.Caption  := Cell[95,UV_iCount];
      QRL_T045.Caption  := Cell[96,UV_iCount];
      QRL_T046.Caption  := Cell[97,UV_iCount];
      QRL_T047.Caption  := Cell[98,UV_iCount];
      QRL_T048.Caption  := Cell[99,UV_iCount];
      QRL_E066.Caption  := Cell[100,UV_iCount];
      QRL_E067.Caption  := Cell[101,UV_iCount];
      //RL_T049.Caption  := Cell[102,UV_iCount];
      QRL_U048.Caption  := Cell[103,UV_iCount];
      QRL_S092.Caption  := Cell[104,UV_iCount];
      QRL_T050.Caption  := Cell[105,UV_iCount];
      QRL_C053.Caption  := Cell[106,UV_iCount];
      QRL_C055.Caption  := Cell[107,UV_iCount];
      QRL_T049.Caption  := Cell[108,UV_iCount];

//      QRL_Desc_name.Caption  := Cell[115,UV_iCount];
      QRL_samp.Caption       := Cell[147,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q049.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;


end.
