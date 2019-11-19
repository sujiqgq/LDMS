unit LD4Q581;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q581 = class(TForm)
    QuickRep: TQuickRep;
    Qrl_999: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape4: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape21: TQRShape;
    Qrl_001: TQRLabel;
    Qrl_005: TQRLabel;
    Qrl_003: TQRLabel;
    QRShape32: TQRShape;
    Qrl_004: TQRLabel;
    QRShape5: TQRShape;
    QRShape15: TQRShape;
    QRLabel3: TQRLabel;
    Qrl_006: TQRLabel;
    QRLabel6: TQRLabel;
    Qrl_002: TQRLabel;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape16: TQRShape;
    QRLabel23: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel53: TQRLabel;
    QRLabel63: TQRLabel;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRLabel73: TQRLabel;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    Qrl_009: TQRLabel;
    Qrl_007: TQRLabel;
    Qrl_008: TQRLabel;
    Qrl_010: TQRLabel;
    Qrl_012: TQRLabel;
    Qrl_011: TQRLabel;
    QRLabel15: TQRLabel;
    QRShape28: TQRShape;
    QRLabel8: TQRLabel;
    QRShape30: TQRShape;
    QRL_013: TQRLabel;
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
  frmLD4Q581: TfrmLD4Q581;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q58;

{$R *.DFM}

procedure TfrmLD4Q581.QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
  var temp,tempa :String;
  begin

   inherited;

   with frmLD4Q58.grdMaster do
   begin

      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      UV_iCount := UV_iCount + 1;

      Qrl_001.Caption   := Cell[1, UV_iCount];
      Qrl_002.Caption   := Cell[2, UV_iCount];
      Qrl_003.Caption   := Cell[3, UV_iCount];
      Qrl_004.Caption   := Cell[4, UV_iCount];
      Qrl_005.Caption   := Cell[5, UV_iCount];
      Qrl_006.Caption   := Cell[6, UV_iCount];
      Qrl_007.Caption   := Cell[7, UV_iCount];
      Qrl_008.Caption   := Cell[8, UV_iCount];
      Qrl_009.Caption   := Cell[9, UV_iCount];
      Qrl_010.Caption   := Cell[10, UV_iCount];
      Qrl_011.Caption   := Cell[11, UV_iCount];
      Qrl_012.Caption   := Cell[12, UV_iCount];
      Qrl_013.Caption   := Cell[13, UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q581.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   if copy(frmLD4Q58.cmbjisa.Text, 1, 2) = '12' then QRLabel5.caption := '검진일자 : '
   else QRLabel5.caption := '분주일자 : ';
   QRL_date.Caption := copy(frmLD4Q58.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q58.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q58.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
 