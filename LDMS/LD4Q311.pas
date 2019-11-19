unit LD4Q311;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls;

type
  TfrmLD4Q311 = class(TForm)
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
    QRLabel10: TQRLabel;
    QRBand3: TQRBand;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRL6: TQRLabel;
    QRL5: TQRLabel;
    QRL7: TQRLabel;
    QRL8: TQRLabel;
    QRL10: TQRMemo;
    QRShape1: TQRShape;
    QRLabel11: TQRLabel;
    QRL9: TQRLabel;
    QRLabel12: TQRLabel;
    QRL12: TQRLabel;
    QRLabel13: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel14: TQRLabel;
    QRL13: TQRLabel;
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
  frmLD4Q311: TfrmLD4Q311;

implementation

uses ZZfunc, LD4Q31;

{$R *.DFM}

procedure TfrmLD4Q311.QuickRep1BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q31 do
   begin
      QRL_Date.Caption := GV_sPrintToday;
   end;
end;

procedure TfrmLD4Q311.QuickRep1NeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q31.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL_No.Caption   := trim(Cell[1,  UV_iCount]);
      QRL1.Caption     := trim(Cell[2,  UV_iCount]);
      QRL2.Caption     := trim(Cell[3,  UV_iCount]);
      QRL3.Caption     := trim(Cell[4,  UV_iCount]);
      QRL4.Caption     := trim(Cell[5,  UV_iCount]);
      QRL5.Caption     := Copy(trim(Cell[6,  UV_iCount]),1,6) + '-' + Copy(trim(Cell[6,  UV_iCount]),7,1) + '******';
      QRL6.Caption     := trim(Cell[7,  UV_iCount]);
      QRL7.Caption     := trim(Cell[8,  UV_iCount]);
      QRL8.Caption     := trim(Cell[9,  UV_iCount]);
      QRL9.Caption     := trim(Cell[10,  UV_iCount]);
      QRL10.Lines.Text := '';
      if trim(Cell[11, UV_iCount]) <> '' then
      begin
         QRL10.Lines.Text := '* 임상소견';
         QRL10.Lines.Add(trim(Cell[11, UV_iCount]));
      end;
      if trim(Cell[12, UV_iCount]) <> '' then
      begin
         QRL10.Lines.Add('* 환자의 임상정보');
         QRL10.Lines.Add(trim(Cell[12, UV_iCount]));
      end;
      QRL12.Caption    := trim(Cell[13, UV_iCount]);
      QRL13.Caption    := trim(Cell[14, UV_iCount]);
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
