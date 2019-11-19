unit LD2Q042;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q042 = class(TForm)
    QuickRep: TQuickRep;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRLabel24: TQRLabel;
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
    QRLabel3: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel15: TQRLabel;
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
    QRL_H038: TQRLabel;
    QRL_H039: TQRLabel;
    QRShape2: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRL_C075: TQRLabel;
    QRL_U049: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape21: TQRShape;
    QRL_No: TQRLabel;
    QRShape29: TQRShape;
    QRL_H031: TQRLabel;
    QRL_C076: TQRLabel;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRL_E050: TQRLabel;
    QRL_U039: TQRLabel;
    QRShape51: TQRShape;
    QRL_samp: TQRLabel;
    QRL_U040: TQRLabel;
    QRL_U041: TQRLabel;
    QRBand4: TQRBand;
    QRLabel13: TQRLabel;
    QRShape5: TQRShape;
    QRL_E051: TQRLabel;
    QRShape8: TQRShape;
    QRLabel20: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape19: TQRShape;
    QRShape24: TQRShape;
    QRLabel21: TQRLabel;
    QRLabel31: TQRLabel;
    QRL_T055: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel41: TQRLabel;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRL_U016: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRL_S100: TQRLabel;
    QRLabel11: TQRLabel;
    QRL_T030: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRL_C016: TQRLabel;
    QRL_S039: TQRLabel;
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
  frmLD2Q042: TfrmLD2Q042;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q042.QuickRepNeedData(Sender: TObject;
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
      //QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
      if Length(Cell[ 4,UV_iCount]) >= 8 then
         QRL_Name.Caption  :=  copy(Cell[ 4,UV_iCount],1,8)
      else
         QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_Jumin.Caption := Cell[ 5,UV_iCount];

      QRL_S100.Caption  := Cell[18,UV_iCount];
      QRL_S039.Caption  := Cell[19,UV_iCount];
      QRL_H038.Caption  := Cell[20,UV_iCount];
      QRL_H039.Caption  := Cell[21,UV_iCount];
      QRL_U049.Caption  := Cell[22,UV_iCount];
      QRL_C075.Caption  := Cell[23,UV_iCount];
      QRL_H031.Caption  := Cell[24,UV_iCount];
      QRL_C076.Caption  := Cell[25,UV_iCount];
      QRL_E050.Caption  := Cell[26,UV_iCount];
      QRL_U039.Caption  := Cell[27,UV_iCount];
      QRL_U040.Caption  := Cell[28,UV_iCount];
      QRL_U041.Caption  := Cell[29,UV_iCount];
      QRL_E051.Caption  := Cell[30,UV_iCount];

      QRL_T055.Caption  := Cell[130,UV_iCount];

      QRL_U016.Caption  := Cell[132,UV_iCount];
      QRL_T030.Caption  := Cell[136,UV_iCount];
      QRL_C016.Caption  := Cell[138,UV_iCount];
     // QRL_Desc_name.Caption  := Cell[115,UV_iCount];
      QRL_samp.Caption       := Cell[147,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q042.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.

