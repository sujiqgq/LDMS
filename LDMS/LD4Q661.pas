unit LD4Q661;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q661 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
    QRShape2: TQRShape;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand3: TQRBand;
    QRLabel9: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel24: TQRLabel;
    QRShape1: TQRShape;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);

  private
    { Private declarations }
    UV_iCount, i, imax : Integer;
    qlArray : array[1..32] of TQRLabel;
    qlArray_HM : array[1..32] of TQRLabel;
  public
    { Public declarations }

  end;

var
  frmLD4Q661: TfrmLD4Q661;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q66;

{$R *.DFM}

procedure TfrmLD4Q661.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i : Integer;
begin
   inherited;
   with frmLD4Q66.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      Inc(UV_iCount);

      QRL_NO.caption     := Cell[1, UV_iCount]; //No.
      QRL_SAMPLE.Caption := Cell[2, UV_iCount]; //샘플번호
      QRL_BUNJU.Caption  := Cell[3, UV_iCount]; //분주번호
      QRL_NAME.Caption   := Cell[4, UV_iCount]; //성명
      for i := 0 to frmLD4Q66.iArr-1 do
      begin
          qlArray[i].Caption    := Cell[i+5, UV_iCount];
      end;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q661.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
  var i : Integer;
begin
   inherited;

   QRL_date.Caption := copy(frmLD4Q66.edtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q66.edtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q66.edtBDate.Text,7,2) + '일';
   for i := 0 to frmLD4Q66.iArr-1 do
   begin
      qlArray_HM[i] := TQRLabel.Create(Self);
      qlArray_HM[i].Parent     := QRBand3;
      qlArray_HM[i].Caption    := frmLD4Q66.vArr_HulHM[i][0];
      qlArray_HM[i].Alignment  := taCenter;
      qlArray_HM[i].Height     := 13;
      qlArray_HM[i].Font.color := clBlack;
      qlArray_HM[i].Font.Size  := 9;

      qlArray[i] := TQRLabel.Create(Self);
      qlArray[i].Parent     := QRBand2;
      qlArray[i].Alignment  := taCenter;
      qlArray[i].Height     := 13;
      qlArray[i].Width      := 35;
      qlArray[i].Font.color := clBlack;
      qlArray[i].Font.Size  := 9;

      if i <= 15 then
      begin
         qlArray_HM[i].Left := 236 + ((i+1)*50);
         qlArray_HM[i].Top  := 3;
         qlArray[i].Left    := 250 + ((i+1)*50);
         qlArray[i].Top     := 2;
      end
      else
      begin
         qlArray_HM[i].Left := 236 + (((i+1)-16)*50);
         qlArray_HM[i].Top  := 19;
         qlArray[i].Left    := 250 + (((i+1)-16)*50);
         qlArray[i].Top     := 18;
      end;
   end;

   UV_iCount := 0;
end;

end.
