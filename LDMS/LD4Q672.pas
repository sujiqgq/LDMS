unit LD4Q672;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q672 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
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
    QRLabel4: TQRLabel;
    QRL_BIRTH: TQRLabel;
    QRLabel7: TQRLabel;
    QRL_GMGN: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure Print;

  private
    { Private declarations }
    UV_iCount, i, iNo : Integer;
    qlArray : array of TQRLabel;
    qlArray_HM : array of TQRLabel;
    Check_Last : Boolean;
  public
    { Public declarations }

  end;

var
  frmLD4Q672: TfrmLD4Q672;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q67;

{$R *.DFM}

procedure TfrmLD4Q672.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i, j : Integer;
begin
   inherited;
   with frmLD4Q67.grdMaster do
   begin
      j := 0;
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      for i := 0 to StrToInt(frmLD4Q67.Label_iArr.Caption)-1 do
         qlArray[i].Caption := '';

      inc(UV_iCount);

      for j:=UV_iCount to Rows do
      begin
         if (Check_Last) or (J = Rows) then
         begin
            Print;
            Check_Last := False;
            break;
         end
         else if (Cell[3, UV_iCount] = Cell[3, UV_iCount+1]) and
                 (Cell[4, UV_iCount] = Cell[4, UV_iCount+1]) then
         begin
            Print;
            inc(UV_iCount);
         end
         else
         begin
            Check_Last := True;
         end;
      end;

      inc(iNo);
      QRL_NO.caption := IntToStr(iNo); //No.

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q672.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
  var i : Integer;
begin
   inherited;
   Check_Last := False;

   QRL_date.Caption := copy(frmLD4Q67.edtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q67.edtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q67.edtBDate.Text,7,2) + '일';

   SetLength(qlArray, StrToInt(frmLD4Q67.Label_iArr.Caption));
   SetLength(qlArray_HM, StrToInt(frmLD4Q67.Label_iArr.Caption));

   for i := 0 to StrToInt(frmLD4Q67.Label_iArr.Caption)-1 do
   begin
      qlArray_HM[i] := TQRLabel.Create(Self);
      qlArray_HM[i].Parent     := QRBand3;
      qlArray_HM[i].Caption    := frmLD4Q67.vArrHM[i];
      qlArray_HM[i].Alignment  := taCenter;
      qlArray_HM[i].Height     := 13;
      qlArray_HM[i].Font.color := clBlack;
      qlArray_HM[i].Font.Size  := 9;

      qlArray[i] := TQRLabel.Create(Self);
      qlArray[i].Parent     := QRBand2;
      qlArray[i].Alignment  := taCenter;
      qlArray[i].Height     := 13;
      qlArray[i].Width      := 27;
      qlArray[i].Font.color := clBlack;
      qlArray[i].Font.Size  := 9;

      qlArray_HM[i].Left := 371 + (i*36);
      qlArray_HM[i].Top  := 3;
      qlArray[i].Left    := 385 + (i*36);
      qlArray[i].Top     := 2;
   end;

   UV_iCount := 0;
   iNo := 0;
end;

procedure TfrmLD4Q672.Print;
var i : integer;
begin
   with frmLD4Q67.grdMaster do
   begin
      QRL_SAMPLE.Caption := Cell[2, UV_iCount]; //샘플번호
      QRL_BUNJU.Caption  := Cell[3, UV_iCount]; //분주번호
      QRL_NAME.Caption   := Cell[4, UV_iCount]; //성명
      if copy(Cell[5,UV_iCount],8,1) <> '' then //생년월일
      begin
        case StrToInt(copy(Cell[5,UV_iCount],8,1)) of
          1,3,5,7,9 : QRL_BIRTH.Caption := 'M/' + copy(Cell[5,UV_iCount],1,6);
          2,4,6,8   : QRL_BIRTH.Caption := 'F/' + copy(Cell[5,UV_iCount],1,6);
        end;
      end
      else QRL_BIRTH.Caption := '';
      QRL_GMGN.Caption   := Cell[6, UV_iCount]; //검진일자
      for i := 0 to StrToInt(frmLD4Q67.Label_iArr.Caption)-1 do
      begin
         if qlArray_HM[i].Caption = Cell[7, UV_iCount] then
         begin
            qlArray[i].Caption := '*';
            break;
         end;
      end;
   end;
end;


end.
