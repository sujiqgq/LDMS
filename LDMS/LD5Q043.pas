unit LD5Q043;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD5Q043 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_No1: TQRLabel;
    Qrl_Bunju1: TQRLabel;
    Qrl_No2: TQRLabel;
    Qrl_Bunju2: TQRLabel;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel6: TQRLabel;
    QRLabel12: TQRLabel;
    qrl_Name1: TQRLabel;
    qrl_name2: TQRLabel;
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
  frmLD5Q043: TfrmLD5Q043;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q04;

{$R *.DFM}

procedure TfrmLD5Q043.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q04.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // ����ڷ� ����
      cRow := cRow + 1;
      If cRow > 30 Then
      Begin
         iRow := iRow + 2;
         cRow := 1;
         UV_iCount := UV_iCount + 1;
      End;
      
      Qrl_No1.Caption     := FormatFloat('#00', (iRow * 30) + cRow);
      Qrl_Name1.Caption   := Cell[6,(iRow * 30) + cRow];

      Qrl_No2.Caption     := FormatFloat('#00', (iRow * 30) + cRow + 30);
      Qrl_Name2.Caption   := Cell[6,(iRow * 30) + cRow + 30];

      if frmLD5Q04.RBtn_Bunju.Checked then
      begin
         Qrl_Bunju1.Caption  := Cell[3,(iRow * 30) + cRow];
         Qrl_Bunju2.Caption  := Cell[3,(iRow * 30) + cRow + 30];
      end
      else
      begin
         Qrl_Bunju1.Caption  := Cell[4,(iRow * 30) + cRow];
         Qrl_Bunju2.Caption  := Cell[4,(iRow * 30) + cRow + 30];
      end;

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD5Q043.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   if frmLD5Q04.CB_H025.Checked then QRLabel1.Caption := 'ESR �������';
   if frmLD5Q04.CB_H023.Checked then QRLabel1.Caption := '������ �������';

   if frmLD5Q04.RBtn_Bunju.Checked then
   begin
      QRLabel5.Caption := '�������� : ';

      QRLabel3.Caption := '���� ��ȣ';
      QRLabel9.Caption := '���� ��ȣ';
   end
   else
   begin
      QRLabel5.Caption := '�������� : ';

      QRLabel3.Caption := '���� ��ȣ';
      QRLabel9.Caption := '���� ��ȣ';
   end;


   QRL_date.Caption := copy(frmLD5Q04.EdtBDate.Text,1,4) + '��'  +
                       copy(frmLD5Q04.EdtBDate.Text,5,2) + '��'  +
                       copy(frmLD5Q04.EdtBDate.Text,7,2) + '��';
   UV_iCount := 1;
   if ((frmLD5Q04.grdMaster.Rows mod 60) = 0) then
      sRec := Round(frmLD5Q04.grdMaster.Rows div 60)
   else
   begin
      sRec := Round(frmLD5Q04.grdMaster.Rows div 60) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
