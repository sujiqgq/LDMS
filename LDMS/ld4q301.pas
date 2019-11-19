unit LD4Q301;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls;

type
  TfrmLD4Q301 = class(TForm)
    QuickRep1: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRBand2: TQRBand;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel16: TQRLabel;
    QRBand3: TQRBand;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRL6: TQRLabel;
    QRL5: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    QRL9: TQRMemo;
    QRShape1: TQRShape;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel17: TQRLabel;
    QRL10: TQRLabel;
    QRL11: TQRLabel;
    QRL12: TQRLabel;
    QRShape2: TQRShape;
    procedure QuickRep1BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QuickRep1NeedData(Sender: TObject; var MoreData: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD4Q301: TfrmLD4Q301;

implementation

uses ZZfunc, LD4Q30;

{$R *.DFM}

procedure TfrmLD4Q301.QuickRep1BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q30 do
   begin
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

procedure TfrmLD4Q301.QuickRep1NeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q30.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      if frmLD4Q30.Ck_print.Checked = True then
      begin
         // 출력자료 전달
         QRL1.Caption := Cell[1, UV_iCount*10];
         QRL2.Caption := Cell[2, UV_iCount*10];
         QRL3.Caption := Cell[3, UV_iCount*10];
         QRL4.Caption := Cell[4, UV_iCount*10];
         QRL5.Caption := Copy(Cell[6, UV_iCount*10],1,6) + '-' + Copy(Cell[6, UV_iCount*10],7,1) + '******';
         QRL6.Caption := Cell[7, UV_iCount*10];
         QRL7.Caption := Cell[8, UV_iCount*10];
         QRL8.Caption := Cell[10, UV_iCount*10];
         QRL9.Lines.Text := trim(Cell[13, UV_iCount*10]);
         QRL10.Caption := Cell[14, UV_iCount*10];
         QRL11.Caption := Cell[15, UV_iCount*10];
         QRL12.Caption := Cell[16, UV_iCount*10];

         if UV_iCount*10 <= Rows then
         begin
            MoreData := True;
            Inc(UV_iCount);
         end
         else
            MoreData := False;
      end
      else
      begin
         // 출력자료 전달
         QRL1.Caption := Cell[1, UV_iCount];
         QRL2.Caption := Cell[2, UV_iCount];
         QRL3.Caption := Cell[3, UV_iCount];
         QRL4.Caption := Cell[4, UV_iCount];
         QRL5.Caption := Copy(Cell[6, UV_iCount],1,6) + '-' + Copy(Cell[6, UV_iCount],7,1) + '******';
         QRL6.Caption := Cell[7, UV_iCount];
         QRL7.Caption := Cell[8, UV_iCount];
         QRL8.Caption := Cell[10, UV_iCount];
         QRL9.Lines.Text := trim(Cell[13, UV_iCount]);
         QRL10.Caption := Cell[14, UV_iCount];
         QRL11.Caption := Cell[15, UV_iCount];
         QRL12.Caption := Cell[16, UV_iCount];

         if UV_iCount <= Rows then
         begin
            MoreData := True;
            Inc(UV_iCount);
         end
         else
            MoreData := False;
      end;
   end;
end;

procedure TfrmLD4Q301.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   free;
end;

end.
