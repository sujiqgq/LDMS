unit LD5Q161;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD5Q161 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand3: TQRBand;
    QRLabel9: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel_101: TQRLabel;
    QRLabel_102: TQRLabel;
    QRLabel_103: TQRLabel;
    QRLabel_104: TQRLabel;
    QRLabel_105: TQRLabel;
    QRLabel_106: TQRLabel;
    QRLabel_107: TQRLabel;
    QRLabel_108: TQRLabel;
    QRLabel_109: TQRLabel;
    QRLabel_201: TQRLabel;
    QRLabel_202: TQRLabel;
    QRLabel_203: TQRLabel;
    QRLabel_204: TQRLabel;
    QRLabel_205: TQRLabel;
    QRLabel_206: TQRLabel;
    QRLabel_207: TQRLabel;
    QRLabel_208: TQRLabel;
    QRLabel_209: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD5Q161: TfrmLD5Q161;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q16;

{$R *.DFM}

procedure TfrmLD5Q161.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i: integer;
begin
   inherited;

   with frmLD5Q16.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      Inc(UV_iCount);

      QRL_NO.caption     := Cell[1,UV_iCount];
      QRL_SAMPLE.Caption := Cell[2,UV_iCount];
      QRL_NAME.Caption   := Cell[4,UV_iCount];
      QRL_BUNJU.Caption  := Cell[3,UV_iCount];

      QRLabel_201.Caption   := Cell[5,UV_iCount];
      QRLabel_202.Caption   := Cell[6,UV_iCount];
      QRLabel_203.Caption   := Cell[7,UV_iCount];
      QRLabel_204.Caption   := Cell[8,UV_iCount];
      QRLabel_205.Caption   := Cell[9,UV_iCount];
      QRLabel_206.Caption   := Cell[10,UV_iCount];
      QRLabel_207.Caption   := Cell[11,UV_iCount];
      QRLabel_208.Caption   := Cell[12,UV_iCount];
      QRLabel_209.Caption   := Cell[13,UV_iCount];



      QRLabel_101.Caption   := frmLD5Q16.grdMaster.Col[5].heading;
      QRLabel_102.Caption   := frmLD5Q16.grdMaster.Col[6].heading;
      QRLabel_103.Caption   := frmLD5Q16.grdMaster.Col[7].heading;
      QRLabel_104.Caption   := frmLD5Q16.grdMaster.Col[8].heading;
      QRLabel_105.Caption   := frmLD5Q16.grdMaster.Col[9].heading;
      QRLabel_106.Caption   := frmLD5Q16.grdMaster.Col[10].heading;
      QRLabel_107.Caption   := frmLD5Q16.grdMaster.Col[11].heading;
      QRLabel_108.Caption   := frmLD5Q16.grdMaster.Col[12].heading;
      QRLabel_109.Caption   := frmLD5Q16.grdMaster.Col[13].heading;


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD5Q161.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD5Q16.edtBDate.Text,1,4) + '년'  +
                       copy(frmLD5Q16.edtBDate.Text,5,2) + '월'  +
                       copy(frmLD5Q16.edtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
 