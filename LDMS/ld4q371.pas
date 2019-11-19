unit LD4Q371;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q371 = class(TForm)
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
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
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
    Qrl_001: TQRLabel;
    Qrl_006: TQRLabel;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel6: TQRLabel;
    Qrl_005: TQRLabel;
    QRShape26: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape23: TQRShape;
    QRLabel14: TQRLabel;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    Qrl_002: TQRLabel;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    QRShape30: TQRShape;
    Qrl_012: TQRLabel;
    Qrl_011: TQRLabel;
    Qrl_010: TQRLabel;
    Qrl_009: TQRLabel;
    Qrl_008: TQRLabel;
    Qrl_007: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRLabel15: TQRLabel;
    Qrl_006_1: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    Qrl_012_1: TQRLabel;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRLabel16: TQRLabel;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel18: TQRLabel;
    qrl_013: TQRLabel;
    qrl_013_1: TQRLabel;
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
  frmLD4Q371: TfrmLD4Q371;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q37, LD4Q13;

{$R *.DFM}

procedure TfrmLD4Q371.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q37.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if Rows = 0 then
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
      if frmLD4Q37.btn_gmgn.checked then QRLabel5.Caption := '�������� :'
      else                               QRLabel5.Caption := '�������� :';
      Qrl_001.Caption   := FormatFloat('#00', (iRow * 50) + cRow);
      Qrl_002.Caption   := Cell[2,(iRow * 50) + cRow];
      Qrl_003.Caption   := Cell[3,(iRow * 50) + cRow];
      Qrl_004.Caption   := Cell[4,(iRow * 50) + cRow];
      Qrl_005.Caption   := Cell[5,(iRow * 50) + cRow];
      if copy(Cell[6,(iRow * 50) + cRow],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[6,(iRow * 50) + cRow],8,1)) of
            1,3,5,7,9 : Qrl_006.Caption := 'M/' + copy(Cell[6,(iRow * 50) + cRow],1,6);
            2,4,6,8   : Qrl_006.Caption := 'F/' + copy(Cell[6,(iRow * 50) + cRow],1,6);
            else  Qrl_006.Caption   := '';
         end;
      end
      else
         Qrl_006.Caption   := '';
      Qrl_006_1.Caption := Cell[7,(iRow * 50) + cRow];
      Qrl_013.Caption   := Cell[8,(iRow * 50) + cRow];

      Qrl_007.Caption   := FormatFloat('#00', (iRow * 50) + cRow + 50);
      Qrl_008.Caption   := Cell[2,(iRow * 50) + cRow + 50];
      Qrl_009.Caption   := Cell[3,(iRow * 50) + cRow + 50];
      Qrl_010.Caption   := Cell[4,(iRow * 50) + cRow + 50];
      Qrl_011.Caption   := Cell[5,(iRow * 50) + cRow + 50];
      if copy(Cell[6,(iRow * 50) + cRow + 50],8,1) <> '' then
      begin
         case StrToInt(copy(Cell[6,(iRow * 50) + cRow + 50],8,1)) of
            1,3,5,7,9 : Qrl_012.Caption := 'M/' + copy(Cell[6,(iRow * 50) + cRow + 50],1,6);
            2,4,6,8   : Qrl_012.Caption := 'F/' + copy(Cell[6,(iRow * 50) + cRow + 50],1,6);
            else  Qrl_012.Caption   := '';
         end;
      end
      else
         Qrl_012.Caption   := '';
      Qrl_012_1.Caption := Cell[7,(iRow * 50) + cRow + 50];
      Qrl_013_1.Caption := Cell[8,(iRow * 50) + cRow + 50];
      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q371.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q37.mskDate.Text,1,4) + '��'  +
                       copy(frmLD4Q37.mskDate.Text,5,2) + '��'  +
                       copy(frmLD4Q37.mskDate.Text,7,2) + '��';
   UV_iCount := 1;
   if ((frmLD4Q37.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD4Q37.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD4Q37.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
