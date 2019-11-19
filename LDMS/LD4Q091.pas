unit LD4Q091;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q091 = class(TForm)
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
    QRShape7: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    Qrl_001: TQRLabel;
    Qrl_006: TQRLabel;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRLabel6: TQRLabel;
    Qrl_005: TQRLabel;
    QRShape26: TQRShape;
    QRShape10: TQRShape;
    QRShape28: TQRShape;
    Qrl_003: TQRLabel;
    Qrl_004: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRLabel15: TQRLabel;
    Qrl_006_1: TQRLabel;
    QRLabel17: TQRLabel;
    Qrl_006_2: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape5: TQRShape;
    QRShape8: TQRShape;
    Qrl_006_3: TQRLabel;
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
  frmLD4Q091: TfrmLD4Q091;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q09;

{$R *.DFM}

procedure TfrmLD4Q091.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var     temp,tempa :String;
  begin

   inherited;

   with frmLD4Q09.grdMaster do
   begin
      if frmLD4Q09.RBtn_S016.Checked then
      Begin
            QRLabel8.caption := 'HCV Ab(��)';
            QRLabel15.caption := 'HCV Ab(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S021.Checked then
      Begin
            QRLabel8.caption := 'H.pylori(��)';
            QRLabel15.caption := 'H.pylori(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S019.Checked then
      Begin
            QRLabel8.caption := 'Rubella IgG (��)';
            QRLabel15.caption := 'Rubella IgG (��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';

      end
      else if frmLD4Q09.RBtn_S018.Checked then
      Begin
            QRLabel8.caption := 'Rubella IgM(��)';
            QRLabel15.caption := 'Rubella IgM(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';

      end
      else if frmLD4Q09.RBtn_SE02.Checked then
      Begin
            QRLabel8.caption := 'HAV IgG(��)';
            QRLabel15.caption := 'HAV IgG(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_SE01.Checked then
      Begin
            QRLabel8.caption := 'HAV IgM(��)';
            QRLabel15.caption := 'HAV IgM(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S010.Checked then
      Begin
            QRLabel8.caption := 'HBc Ab(��)';
            QRLabel15.caption := 'HBc Ab(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S011.Checked then
      Begin
            QRLabel8.caption := 'HBe Ag(��)';
            QRLabel15.caption := 'HBe Ag(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S012.Checked then
      Begin
            QRLabel8.caption := 'HBe Ab(��)';
            QRLabel15.caption := 'HBe Ab(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S034.Checked then
      Begin
            QRLabel8.caption := 'RPR ����(��)';
            QRLabel15.caption := 'RPR ����(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';
      end
      else if frmLD4Q09.RBtn_S036.Checked then
      Begin
            QRLabel8.caption := 'TPLA ����(��)';
            QRLabel15.caption := 'TPLA ����(��)';
            QRLabel17.caption := '';
            QRLabel3.caption := '';

      end
      else if frmLD4Q09.RBtn_S033.Checked    then
      Begin
         QRLabel8.caption  := 'anti-HCV(��)';
         QRLabel15.caption := 'anti-HCV(��)';
         QRLabel17.caption := '';
         QRLabel3.caption  := '';

      end
      else if frmLD4Q09.RBtn_S018S019.Checked then
      Begin
            QRLabel8.caption := 'Rubella IgM (��)';
            QRLabel15.caption := 'Rubella igG (��)';
            QRLabel17.caption := 'Rubella IgM  (��)';
            QRLabel3.caption := 'Rubella igG (��)';
      end
      else if frmLD4Q09.RBtn_SE01SE02.Checked then
      Begin
            QRLabel8.caption := 'HAV IgM(��)';
            QRLabel15.caption := 'HAV IgG(��)';
            QRLabel17.caption := 'HAV IgM(��)';
            QRLabel3.caption := 'HAV IgG(��)';
      end
      else if frmLD4Q09.Rbtn_S011S012.Checked then
      Begin
            QRLabel8.caption := 'HBe Ag(��)';
            QRLabel15.caption := 'HBe Ab(��)';
            QRLabel17.caption := 'HBe Ag(��)';
            QRLabel3.caption := 'HBe Ab(��)';
      end
      else if frmLD4Q09.Rbtn_S034S036.Checked then
      Begin
            QRLabel8.caption :=  'RPR ����(��)';
            QRLabel15.caption := 'TPLA ����(��)';
            QRLabel17.caption := 'RPR ����(��)';
            QRLabel3.caption := 'TPLA ����(��)';

      end
      else if frmLD4Q09.Rbtn_S007S008.Checked then
      Begin
            QRLabel8.caption := 'HBS Ag(��)';
            QRLabel15.caption := 'HBS Ab(��)';
            QRLabel17.caption := 'HBS Ag(��)';
            QRLabel3.caption := 'HBS Ab(��)';

      end
      else if frmLD4Q09.Rbtn_SE31SE02.Checked then
      Begin
            QRLabel8.caption := 'HAVAB,TOTAL (��)';
            QRLabel15.caption := 'HAVAB,TOTAL (��)';
            QRLabel17.caption := 'HAV IgG(��)';
            QRLabel3.caption := '';

      end
      else if frmLD4Q09.Rbtn_SE02SE31.Checked then
      Begin
            QRLabel8.caption := 'HAVAB,TOTAL (��)';
            QRLabel15.caption := 'HAVAB,TOTAL (��)';
            QRLabel17.caption := 'HAV IgG(��)';
            QRLabel3.caption := '';

      end;

      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // ����ڷ� ����
      cRow := cRow + 1;
      {If cRow > 50 Then
         Begin
            iRow := iRow + 2;
            cRow := 1;
            UV_iCount := UV_iCount + 1;
         End;}

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

      Qrl_001.Caption   := Cell[2,cRow];
      Qrl_003.Caption   := Cell[3,cRow];
      Qrl_004.Caption   := Cell[4,cRow];
      Qrl_005.Caption   := Cell[5,cRow];
      Qrl_006.Caption   := Cell[6,cRow];
      Qrl_006_1.Caption := Cell[7,cRow];
      Qrl_006_2.Caption := Cell[8,cRow];
      Qrl_006_3.Caption := Cell[9,cRow];

      if UV_iCount <= Rows then
      begin
         MoreData := True;
         UV_iCount := UV_iCount + 1;
      end
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q091.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q09.EdtBDate.Text,1,4) + '��'  +
                       copy(frmLD4Q09.EdtBDate.Text,5,2) + '��'  +
                       copy(frmLD4Q09.EdtBDate.Text,7,2) + '��';
   UV_iCount := 1;
   if ((frmLD4Q09.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD4Q09.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD4Q09.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0;
end;

end.
