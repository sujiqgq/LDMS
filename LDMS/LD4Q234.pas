unit LD4Q234;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q234 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape22: TQRShape;
    QRLabel7: TQRLabel;
    QRShape23: TQRShape;
    QRShape25: TQRShape;
    QRShape5: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel4: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    Qrl_Bunju_date: TQRLabel;
    qrl_no: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    qrl_num: TQRLabel;
    qrl_E068: TQRLabel;
    qrl_name: TQRLabel;
    QRShape21: TQRShape;
    QRShape4: TQRShape;
    QRLabel3: TQRLabel;
    QRShape6: TQRShape;
    QRLabel6: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    Qrl_num_Bunju: TQRLabel;
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
  frmLD4Q234: TfrmLD4Q234;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q23;

{$R *.DFM}

procedure TfrmLD4Q234.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;  
   with frmLD4Q23.grdS021 do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;
      // 출력자료 전달

      Qrl_No.caption           := Cell[1, UV_iCount];
      Qrl_bunju_date.caption   := Cell[2, UV_iCount];
      Qrl_num.caption          := Cell[3, UV_iCount];
      Qrl_name.caption        := Cell[4, UV_iCount];
      Qrl_num_Bunju.Caption   := Cell[5, UV_iCount];
      qrl_E068.Caption   := Cell[6, UV_iCount];

      
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q234.QuickRepBeforePrint(Sender: TCustomQuickRep;
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
