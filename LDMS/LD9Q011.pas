unit LD9Q011;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD9Q011 = class(TForm)
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
    QRShape6: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape26: TQRShape;
    QRShape31: TQRShape;
    QRLabel15: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape21: TQRShape;
    QRL1: TQRLabel;
    QRL5: TQRLabel;
    QRShape24: TQRShape;
    QRL6: TQRLabel;
    QRShape28: TQRShape;
    QRL2: TQRLabel;
    QRShape32: TQRShape;
    QRL7: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape5: TQRShape;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRL3: TQRLabel;
    QRL4: TQRLabel;
    QRShape29: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRLabel9: TQRLabel;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape25: TQRShape;
    QRShape27: TQRShape;
    QRShape30: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel16: TQRLabel;
    QRL14: TQRLabel;
    QRL13: TQRLabel;
    QRL12: TQRLabel;
    QRL11: TQRLabel;
    QRL10: TQRLabel;
    QRL9: TQRLabel;
    QRL8: TQRLabel;
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
  frmLD9Q011: TfrmLD9Q011;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD9Q01;

{$R *.DFM}

procedure TfrmLD9Q011.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD9Q01.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if (Rows = 0) or (UV_iCount = Rows) then
      begin
         MoreData := False;
         exit;
      end;

      // ����ڷ� ����
      cRow := cRow + 1;
      If cRow > 50 Then
         Begin
            iRow := iRow + 2;
            cRow := 1;
            UV_iCount := UV_iCount + 1;
         End;

      if (cRow mod 10) = 0 then
      begin
         QRShape13.Pen.Style := psSolid;
         QRShape13.Pen.Width := 3;
      end
      else
      begin
         QRShape13.Pen.Style := psDot;
         QRShape13.Pen.width := 1;
      end;

      QRL1.Caption  := ''; QRL2.Caption  := ''; QRL3.Caption  := ''; QRL4.Caption  := '';
      QRL5.Caption  := ''; QRL6.Caption  := ''; QRL7.Caption  := '';
      QRL8.Caption  := ''; QRL9.Caption  := ''; QRL10.Caption := ''; QRL11.Caption := '';
      QRL12.Caption := ''; QRL13.Caption := ''; QRL14.Caption := '';

      QRL1.Caption   := Cell[1,(iRow * 50) + cRow]; //no
      QRL4.Caption   := Cell[2,(iRow * 50) + cRow]; //����
      QRL2.Caption   := Cell[3,(iRow * 50) + cRow]; //���ֹ�ȣ
      QRL3.Caption   := Cell[4,(iRow * 50) + cRow]; //���ù�ȣ
      QRL5.Caption   := Cell[5,(iRow * 50) + cRow]; //����
      QRL6.Caption   := Cell[6,(iRow * 50) + cRow]; //�̸�
      QRL7.Caption   := Cell[8,(iRow * 50) + cRow]; //��������

      QRL8.Caption    := Cell[1,(iRow * 50) + cRow + 50];
      QRL9.Caption    := Cell[2,(iRow * 50) + cRow + 50];
      QRL10.Caption   := Cell[3,(iRow * 50) + cRow + 50];
      QRL11.Caption   := Cell[4,(iRow * 50) + cRow + 50];
      QRL12.Caption   := Cell[5,(iRow * 50) + cRow + 50];
      QRL13.Caption   := Cell[6,(iRow * 50) + cRow + 50];
      QRL14.Caption   := Cell[8,(iRow * 50) + cRow + 50];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD9Q011.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   if frmLD9Q01.RBtn_Bunju.Checked then QRLabel5.Caption := '�������� : '
   else                                 QRLabel5.Caption := '�������� : ';

   QRL_date.Caption := copy(frmLD9Q01.EdtBDate.Text,1,4) + '��' +
                       copy(frmLD9Q01.EdtBDate.Text,5,2) + '��' +
                       copy(frmLD9Q01.EdtBDate.Text,7,2) + '��';

   QRLabel1.Caption := '�˻���� ����Ʈ(' + frmLD9Q01.Cmb_Hm.Text + ')';

   UV_iCount := 1;
   if ((frmLD9Q01.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD9Q01.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD9Q01.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
