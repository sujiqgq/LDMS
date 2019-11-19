unit LD4Q232;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q232 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape10: TQRShape;
    QRShape14: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape22: TQRShape;
    QRLabel7: TQRLabel;
    QRShape23: TQRShape;
    QRLabel8: TQRLabel;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel11: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape5: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    Qrl_date1: TQRLabel;
    Qrl_date2: TQRLabel;
    qrl_no1: TQRLabel;
    qrl_no2: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    qrl_num1: TQRLabel;
    qrl_S0191: TQRLabel;
    qrl_name1: TQRLabel;
    qrl_num2: TQRLabel;
    qrl_S0192: TQRLabel;
    qrl_name2: TQRLabel;
    QRShape18: TQRShape;
    QRShape21: TQRShape;
    QRShape4: TQRShape;
    QRLabel10: TQRLabel;
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
  frmLD4Q232: TfrmLD4Q232;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q23;

{$R *.DFM}

procedure TfrmLD4Q232.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
  if  frmLD4Q23.Cmb_gubn.ItemIndex = 0 then
  begin
       QRLabel7.caption :='샘플번호';
       QRLabel11.caption :='샘플번호';
  end
  else
  begin
       QRLabel7.caption :='분주번호';
       QRLabel11.caption :='분주번호';
  end;
   with frmLD4Q23.grdS021 do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 출력자료 전달
      cRow := cRow + 1;
      If cRow > 30 Then
      Begin
         iRow := iRow + 2;
         cRow := 1;
         UV_iCount := UV_iCount + 1;
      End;
      Qrl_No1.caption     := Cell[1,(iRow * 30) + cRow];
      Qrl_No2.caption     := Cell[1,(iRow * 30) + cRow + 30];
      Qrl_date1.caption   := Cell[3,(iRow * 30) + cRow];
      Qrl_date2.caption   := Cell[3,(iRow * 30) + cRow + 30];
      Qrl_num1.caption    := Cell[4,(iRow * 30) + cRow];
      Qrl_num2.caption    := Cell[4,(iRow * 30) + cRow + 30];
      Qrl_name1.caption   := Cell[5,(iRow * 30) + cRow];
      Qrl_name2.caption   := Cell[5,(iRow * 30) + cRow + 30];
      Qrl_S0191.Caption   := Cell[6,(iRow * 30) + cRow];
      Qrl_S0192.Caption   := Cell[6,(iRow * 30) + cRow + 30];

      if UV_iCount <= sRec then
         MoreData := True
      else
         MoreData := False;
   end;



end;

procedure TfrmLD4Q232.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   Qrl_Date.Caption := copy(frmLD4Q23.edtSDate.Text,1,4) + '년'  +
                       copy(frmLD4Q23.edtSDate.Text,5,2) + '월'  +
                       copy(frmLD4Q23.edtSDate.Text,7,2) + '일';
   if frmLD4Q23.edtSDate.Text <> frmLD4Q23.edtEDate.Text then
      Qrl_Date.Caption := Qrl_Date.Caption + ' ~ ' +
                          copy(frmLD4Q23.edtEDate.Text,1,4) + '년'  +
                          copy(frmLD4Q23.edtEDate.Text,5,2) + '월'  +
                          copy(frmLD4Q23.edtEDate.Text,7,2) + '일';

   UV_iCount := 0;

   if frmLD4Q23.grdS021.Rows < 60 then sRec := 0
   else if frmLD4Q23.grdS021.Rows < 120 then sRec := 1
   else if frmLD4Q23.grdS021.Rows < 180 then sRec := 2
   else if frmLD4Q23.grdS021.Rows < 240 then sRec := 3;

   cRow := 0; iRow := 0;
end;

end.
