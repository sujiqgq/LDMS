unit LD2Q046;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q046 = class(TForm)
    QuickRep: TQuickRep;
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
    QRLabel13: TQRLabel;
    QRShape5: TQRShape;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRShape19: TQRShape;
    QRLabel33: TQRLabel;
    QRBand4: TQRBand;
    QRBand1: TQRBand;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRShape49: TQRShape;
    QRShape50: TQRShape;
    QRShape52: TQRShape;
    QRShape53: TQRShape;
    QRShape54: TQRShape;
    QRShape55: TQRShape;
    QRL_Bunju: TQRLabel;
    QRL_C065: TQRLabel;
    QRL_C072: TQRLabel;
    QRL_C079: TQRLabel;
    QRL_S060: TQRLabel;
    QRL_S061: TQRLabel;
    QRL_S059: TQRLabel;
    QRShape56: TQRShape;
    QRShape57: TQRShape;
    QRShape58: TQRShape;
    QRL_S063: TQRLabel;
    QRL_S062: TQRLabel;
    QRShape59: TQRShape;
    QRShape60: TQRShape;
    QRShape61: TQRShape;
    QRShape62: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape63: TQRShape;
    QRL_No: TQRLabel;
    QRShape64: TQRShape;
    QRL_S070: TQRLabel;
    QRL_S081: TQRLabel;
    QRShape65: TQRShape;
    QRShape66: TQRShape;
    QRShape67: TQRShape;
    QRL_S082: TQRLabel;
    QRL_S030: TQRLabel;
    QRShape68: TQRShape;
    QRL_samp: TQRLabel;
    QRL_S031: TQRLabel;
    QRL_S032: TQRLabel;
    QRL_T018: TQRLabel;
    QRShape69: TQRShape;
    QRL_E013: TQRLabel;
    QRShape70: TQRShape;
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
  frmLD2Q046: TfrmLD2Q046;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q046.QuickRepNeedData(Sender: TObject;
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

      QRL_C065.Caption  := Cell[46,UV_iCount];
      QRL_C072.Caption  := Cell[47,UV_iCount];
      QRL_C079.Caption  := Cell[48,UV_iCount];
      QRL_S059.Caption  := Cell[49,UV_iCount];
      QRL_S060.Caption  := Cell[50,UV_iCount];
      QRL_S061.Caption  := Cell[51,UV_iCount];
      QRL_S062.Caption  := Cell[52,UV_iCount];
      QRL_S063.Caption  := Cell[53,UV_iCount];
      QRL_S070.Caption  := Cell[54,UV_iCount];
      QRL_S081.Caption  := Cell[55,UV_iCount];
      QRL_S082.Caption  := Cell[56,UV_iCount];
      QRL_S030.Caption  := Cell[57,UV_iCount];
      QRL_S031.Caption  := Cell[58,UV_iCount];
      QRL_S032.Caption  := Cell[59,UV_iCount];
      QRL_T018.Caption  := Cell[60,UV_iCount];
      QRL_E013.Caption  := Cell[61,UV_iCount];

     // QRL_Desc_name.Caption  := Cell[115,UV_iCount];
      QRL_samp.Caption  := Cell[147,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q046.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q04.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q04.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q04.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;


end.
