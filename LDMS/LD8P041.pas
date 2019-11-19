unit LD8P041;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls;

type
  TfrmLD8P041 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
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
  frmLD8P041: TfrmLD8P041;

implementation

uses ZZFunc, MdmsFunc, LD8P04;
{$R *.DFM}

procedure TfrmLD8P041.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   with frmLD8P04.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      if UV_iCount <= Rows then
      begin
         QRLabel1.Caption  := Cell[4, UV_iCount];

        if Cell[3, UV_iCount] <> '' then QRLabel2.Caption  := '(' + Cell[ 3, UV_iCount] + ')'
         else                             QRLabel2.Caption  := '';
      end;
      
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD8P041.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
