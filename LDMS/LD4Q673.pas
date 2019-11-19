unit LD4Q673;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q673 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape21: TQRShape;
    Qrl_No: TQRLabel;
    Qrl_birth: TQRLabel;
    Qrl_bunju: TQRLabel;
    QRShape32: TQRShape;
    Qrl_name: TQRLabel;
    QRShape5: TQRShape;
    QRShape15: TQRShape;
    Qrl_U024: TQRLabel;
    Qrl_sample: TQRLabel;
    QRShape10: TQRShape;
    QRShape16: TQRShape;
    QRShape18: TQRShape;
    QRShape23: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    Qrl_U029: TQRLabel;
    Qrl_U030: TQRLabel;
    Qrl_U032: TQRLabel;
    Qrl_U051: TQRLabel;
    Qrl_U044: TQRLabel;
    Qrl_999: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape29: TQRShape;
    QRShape31: TQRShape;
    QRShape4: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel23: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel53: TQRLabel;
    QRLabel63: TQRLabel;
    QRShape17: TQRShape;
    QRShape22: TQRShape;
    QRLabel73: TQRLabel;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel16: TQRLabel;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRLabel11: TQRLabel;
    QRL_gmgn: TQRLabel;
    QRShape2: TQRShape;
    QRShape28: TQRShape;
    QRLabel12: TQRLabel;
    QRLabel17: TQRLabel;
    QRL_U031: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure Print;
  private
    UV_iCount, sRec, cRow, iRow, iNo,
    sum_U024, sum_U029, sum_U030, sum_U031, sum_U032, sum_U044, sum_U051: Integer;
    qlArray : array of TQRLabel;
    Check_Last : Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q673: TfrmLD4Q673;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q67;

{$R *.DFM}

procedure TfrmLD4Q673.QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
  var temp,tempa :String;
      j : integer;
  begin
   inherited;

   with frmLD4Q67.grdMaster do
   begin

      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      inc(UV_iCount);

      qrl_U024.Caption := '';
      qrl_U029.Caption := '';
      qrl_U030.Caption := '';
      qrl_U031.Caption := '';
      qrl_U032.Caption := '';
      qrl_U044.Caption := '';
      qrl_U051.Caption := '';

      for j:=UV_iCount to Rows+1 do
      begin
         if j = Rows+1 then
         begin
            QRL_NO.caption     := '';
            QRL_SAMPLE.Caption := '';
            QRL_BUNJU.Caption  := '';
            QRL_NAME.Caption   := '총합계';
            QRL_BIRTH.Caption  := '';
            QRL_GMGN.Caption   := '';
            qrl_U024.Caption   := IntToStr(sum_U024);
            qrl_U029.Caption   := IntToStr(sum_U029);
            qrl_U030.Caption   := IntToStr(sum_U030);
            qrl_U031.Caption   := IntToStr(sum_U031);
            qrl_U032.Caption   := IntToStr(sum_U032);
            qrl_U044.Caption   := IntToStr(sum_U044);
            qrl_U051.Caption   := IntToStr(sum_U051);
         end
         else if (Check_Last) or (J = Rows) then
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

      if UV_iCount <= Rows+1 then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q673.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q67.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q67.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q67.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
   iNo := 0;
   sum_U024 := 0;   sum_U029 := 0;   sum_U030 := 0;   sum_U031 := 0;
   sum_U032 := 0;   sum_U044 := 0;   sum_U051 := 0;
end;


procedure TfrmLD4Q673.Print;
begin
   with frmLD4Q67.grdMaster do
   begin
      QRL_SAMPLE.Caption := Cell[2, UV_iCount]; //샘플번호
      QRL_BUNJU.Caption  := Cell[3, UV_iCount]; //분주번호
      QRL_NAME.Caption   := copy(Cell[4, UV_iCount],1,14); //성명
      if copy(Cell[5,UV_iCount],8,1) <> '' then //생년월일
      begin
        case StrToInt(copy(Cell[5,UV_iCount],8,1)) of
          1,3,5,7,9 : QRL_BIRTH.Caption := 'M/' + copy(Cell[5,UV_iCount],1,6);
          2,4,6,8   : QRL_BIRTH.Caption := 'F/' + copy(Cell[5,UV_iCount],1,6);
        end;
      end
      else QRL_BIRTH.Caption := '';
      QRL_GMGN.Caption   := Cell[6, UV_iCount]; //검진일자
      //검사항목
      if      Cell[7,UV_iCount] = 'U024' then
      begin
         qrl_U024.Caption := '*';
         inc(sum_U024);
      end
      else if Cell[7,UV_iCount] = 'U029' then
      begin
         qrl_U029.Caption := '*';
         inc(sum_U029);
      end
      else if Cell[7,UV_iCount] = 'U030' then
      begin
         qrl_U030.Caption := '*';
         inc(sum_U030);
      end
      else if Cell[7,UV_iCount] = 'U031' then
      begin
         qrl_U031.Caption := '*';
         inc(sum_U031);
      end
      else if Cell[7,UV_iCount] = 'U032' then
      begin
         qrl_U032.Caption := '*';
         inc(sum_U032);
      end
      else if Cell[7,UV_iCount] = 'U044' then
      begin
         qrl_U044.Caption := '*';
         inc(sum_U044);
      end
      else if Cell[7,UV_iCount] = 'U051' then
      begin
         qrl_U051.Caption := '*';
         inc(sum_U051);
      end;
   end;
end;

end.
