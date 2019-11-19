unit LD5Q051;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD5Q051 = class(TForm)
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
  frmLD5Q051: TfrmLD5Q051;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q05;

{$R *.DFM}

procedure TfrmLD5Q051.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q05.grdMaster do
   begin
      // 자료가 없을경우의 처리
     

      // 출력자료 전달
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

      //UV_iCount := UV_iCount + 1;

      QRL1.Caption  := ''; QRL2.Caption  := ''; QRL3.Caption  := ''; QRL4.Caption  := '';
      QRL5.Caption  := ''; QRL6.Caption  := ''; QRL7.Caption  := '';
      QRL8.Caption  := ''; QRL9.Caption  := ''; QRL10.Caption := ''; QRL11.Caption := '';
      QRL12.Caption := ''; QRL13.Caption := ''; QRL14.Caption := '';

      QRL1.Caption   := FormatFloat('#00', (iRow * 50) + cRow);                  //No.
      QRL2.Caption   := Cell[3,(iRow * 50) + cRow];                              //분주번호
      QRL3.Caption   := Cell[4,(iRow * 50) + cRow];                              //샘플번호
      QRL4.Caption   := Cell[2,(iRow * 50) + cRow];                              //지사

      if copy(Cell[5,(iRow * 50) + cRow],7,1) <> '' then                         //생년월일
      begin
         case StrToInt(copy(Cell[5,(iRow * 50) + cRow],7,1)) of
            1,3,5,7,9 : QRL5.Caption := 'M/' + copy(Cell[5,(iRow * 50) + cRow],1,6);
            2,4,6,8   : QRL5.Caption := 'F/' + copy(Cell[5,(iRow * 50) + cRow],1,6);
            else  QRL5.Caption   := '';
         end;
      end
      else
         QRL5.Caption   := '';
      QRL6.Caption   := Cell[6,(iRow * 50) + cRow];                              //이름
      QRL7.Caption   := Cell[7,(iRow * 50) + cRow];                              //단체명

      QRL8.Caption    := FormatFloat('#00', (iRow * 50) + cRow + 50);
      QRL9.Caption    := Cell[3,(iRow * 50) + cRow + 50];
      QRL10.Caption   := Cell[4,(iRow * 50) + cRow + 50];
      QRL11.Caption   := Cell[2,(iRow * 50) + cRow + 50];


      if copy(Cell[5,(iRow * 50) + cRow + 50],7,1) <> '' then
      begin
         case StrToInt(copy(Cell[5,(iRow * 50) + cRow + 50],7,1)) of
            1,3,5,7,9 : QRL12.Caption := 'M/' + copy(Cell[5,(iRow * 50) + cRow + 50],1,6);
            2,4,6,8   : QRL12.Caption := 'F/' + copy(Cell[5,(iRow * 50) + cRow + 50],1,6);
            else  QRL12.Caption   := '';
         end;
      end
      else
         QRL12.Caption   := '';

      QRL13.Caption   := Cell[6,(iRow * 50) + cRow + 50];
      QRL14.Caption   := Cell[7,(iRow * 50) + cRow + 50];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD5Q051.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   if frmLD5Q05.RBtn_Bunju.Checked then QRLabel5.Caption := '분주일자 : '
   else                                 QRLabel5.Caption := '검진일자 : ';

   QRL_date.Caption := copy(frmLD5Q05.EdtBDate.Text,1,4) + '년' +
                       copy(frmLD5Q05.EdtBDate.Text,5,2) + '월' +
                       copy(frmLD5Q05.EdtBDate.Text,7,2) + '일';

   UV_iCount := 1;
   if ((frmLD5Q05.grdMaster.Rows mod 100) = 0) then
      sRec := Round(frmLD5Q05.grdMaster.Rows div 100)
   else
   begin
      sRec := Round(frmLD5Q05.grdMaster.Rows div 100) + 1;
   end;
   cRow := 0; iRow := 0; 
end;

end.
