unit LD2Q048;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q048 = class(TForm)
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
    QRL_T026: TQRLabel;
    QRL_T027: TQRLabel;
    QRL_T028: TQRLabel;
    QRL_T033: TQRLabel;
    QRL_T035: TQRLabel;
    QRL_T032: TQRLabel;
    QRShape2: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_H027: TQRLabel;
    QRL_T036: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape21: TQRShape;
    QRL_No: TQRLabel;
    QRShape29: TQRShape;
    QRL_H028: TQRLabel;
    QRL_E014: TQRLabel;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRL_E059: TQRLabel;
    QRL_E060: TQRLabel;
    QRShape51: TQRShape;
    QRL_samp: TQRLabel;
    QRL_E061: TQRLabel;
    QRL_E062: TQRLabel;
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
    QRLabel37: TQRLabel;
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
    QRLabel38: TQRLabel;
    QRLabel40: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRShape5: TQRShape;
    QRShape8: TQRShape;
    QRLabel13: TQRLabel;
    QRLabel20: TQRLabel;
    QRL_H046: TQRLabel;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRLabel21: TQRLabel;
    QRLabel31: TQRLabel;
    QRL_S103: TQRLabel;
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
  frmLD2Q048: TfrmLD2Q048;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q048.QuickRepNeedData(Sender: TObject;
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

      QRL_T026.Caption  := Cell[78,UV_iCount];
      QRL_T027.Caption  := Cell[79,UV_iCount];
      QRL_T028.Caption  := Cell[80,UV_iCount];
      QRL_T032.Caption  := Cell[81,UV_iCount];
      QRL_T033.Caption  := Cell[82,UV_iCount];
      QRL_T035.Caption  := Cell[83,UV_iCount];
      QRL_T036.Caption  := Cell[84,UV_iCount];
      QRL_H027.Caption  := Cell[85,UV_iCount];
      QRL_H028.Caption  := Cell[86,UV_iCount];
      QRL_E014.Caption  := Cell[87,UV_iCount];
      QRL_E059.Caption  := Cell[88,UV_iCount];
      QRL_E060.Caption  := Cell[89,UV_iCount];
      QRL_E061.Caption  := Cell[90,UV_iCount];
      QRL_E062.Caption  := Cell[91,UV_iCount];
      QRL_H046.Caption  := Cell[135,UV_iCount];
      QRL_S103.Caption  := Cell[137,UV_iCount];


//      QRL_Desc_name.Caption  := Cell[115,UV_iCount];
      QRL_samp.Caption       := Cell[147,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q048.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;


end.
