unit LD4Q341;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls;

type
  TfrmLD4Q341 = class(TForm)
    QuickRep1: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRBand2: TQRBand;
    QRLabel6: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRBand3: TQRBand;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL4: TQRLabel;
    QRL5: TQRLabel;
    QRL11: TQRLabel;
    QRL6: TQRLabel;
    QRL10: TQRLabel;
    QRShape1: TQRShape;
    QRSysData2: TQRSysData;
    QRLabel11: TQRLabel;
    QRLabel8: TQRLabel;
    QRL3: TQRLabel;
    QRLabel12: TQRLabel;
    QRL12: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel13: TQRLabel;
    QRL_samp: TQRLabel;
    QRL14: TQRLabel;
    procedure QuickRep1BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QuickRep1NeedData(Sender: TObject; var MoreData: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD4Q341: TfrmLD4Q341;

implementation

uses ZZfunc, LD4Q34;

{$R *.DFM}

procedure TfrmLD4Q341.QuickRep1BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q34 do
   begin
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

procedure TfrmLD4Q341.QuickRep1NeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q34.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL1.Caption  := Cell[1,  UV_iCount];
      QRL2.Caption  := Cell[2,  UV_iCount];
      QRL3.Caption  := Cell[3,  UV_iCount];
      QRL4.Caption  := Cell[4,  UV_iCount];
      QRL5.Caption  := Cell[5,  UV_iCount];
      QRL6.Caption  := Copy(Cell[6,  UV_iCount],1,6) + '-' + Copy(Cell[6,  UV_iCount],7,1) + '******';

      if  pos('/',Cell[  10,UV_iCount]) > 0 then
      begin
      QRL10.Caption := Copy(Cell[  10,UV_iCount],1, pos('/',Cell[  10,UV_iCount])-1);
      QRL14.Caption := Copy(Cell[  10,UV_iCount],11, 11);
      end
      else
      begin
      QRL10.Caption := Cell[10,  UV_iCount];
      QRL14.Caption := '';
      end;

      QRL11.Caption := copy(Cell[11, UV_iCount],1,22);
      QRL12.Caption := Cell[12, UV_iCount];
      QRL_samp.Caption := Cell[9, UV_iCount];
      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

end.
