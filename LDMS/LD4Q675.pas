unit LD4Q675;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q675 = class(TForm)
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
    Qrl_sample: TQRLabel;
    QRShape10: TQRShape;
    QRShape16: TQRShape;
    QRShape18: TQRShape;
    QRShape23: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    Qrl_U023: TQRLabel;
    Qrl_U026: TQRLabel;
    Qrl_U037: TQRLabel;
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
    QRLabel6: TQRLabel;
    QRShape7: TQRShape;
    QRShape9: TQRShape;
    QRShape17: TQRShape;
    QRShape22: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel16: TQRLabel;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRLabel11: TQRLabel;
    QRL_gmgn: TQRLabel;
    QRShape28: TQRShape;
    QRShape30: TQRShape;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape34: TQRShape;
    QRShape36: TQRShape;
    QRL_U045: TQRLabel;
    QRL_U050: TQRLabel;
    QRL_U047: TQRLabel;
    QRShape2: TQRShape;
    QRLabel3: TQRLabel;
    QRL_U062: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure Print;
  private
    UV_iCount, sRec, cRow, iRow, iNo,
    sum_U020, sum_U023, sum_U026, sum_U028, sum_U037,
    sum_U043, sum_U045, sum_U047, sum_U050, sum_U062 : Integer;

    qlArray : array of TQRLabel;
    Check_Last : Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q675: TfrmLD4Q675;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q67;

{$R *.DFM}

procedure TfrmLD4Q675.QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
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

      qrl_U023.Caption := '';
      qrl_U026.Caption := '';
      qrl_U037.Caption := '';
      qrl_U045.Caption := '';
      qrl_U047.Caption := '';
      qrl_U050.Caption := '';
      qrl_U062.Caption := '';

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
            qrl_U023.Caption   := IntToStr(sum_U023);
            qrl_U026.Caption   := IntToStr(sum_U026);
            qrl_U037.Caption   := IntToStr(sum_U037);
            qrl_U045.Caption   := IntToStr(sum_U045);
            qrl_U047.Caption   := IntToStr(sum_U047);
            qrl_U050.Caption   := IntToStr(sum_U050);
            qrl_U062.Caption   := IntToStr(sum_U062);
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

procedure TfrmLD4Q675.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q67.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q67.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q67.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
   iNo := 0;
   sum_U020 := 0; sum_U023 := 0; sum_U026 := 0; sum_U028 := 0; sum_U037 := 0;
   sum_U043 := 0; sum_U045 := 0; sum_U047 := 0; sum_U050 := 0; sum_U062 := 0;
end;


procedure TfrmLD4Q675.Print;
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
      if  Cell[7,UV_iCount] = 'U023' then
      begin
          qrl_U023.Caption := '*';
          inc(sum_U023);
      end
      else if Cell[7,UV_iCount] = 'U026' then
      begin
          qrl_U026.Caption := '*';
          inc(sum_U026);
      end
      else if Cell[7,UV_iCount] = 'U037' then
      begin
          qrl_U037.Caption := '*';
          inc(sum_U037);
      end
      else if Cell[7,UV_iCount] = 'U045' then
      begin
          qrl_U045.Caption := '*';
          inc(sum_U045);
      end
      else if Cell[7,UV_iCount] = 'U047' then
      begin
          qrl_U047.Caption := '*';
          inc(sum_U047);
      end
      else if Cell[7,UV_iCount] = 'U050' then
      begin
         qrl_U050.Caption := '*';
         inc(sum_U050);
      end
      else if Cell[7,UV_iCount] = 'U062' then
      begin
         qrl_U062.Caption := '*';
         inc(sum_U062);
      end;
   end;
end;



end.
