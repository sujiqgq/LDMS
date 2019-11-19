unit LD4Q674;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q674 = class(TForm)
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
    Qrl_U060: TQRLabel;
    Qrl_sample: TQRLabel;
    QRShape10: TQRShape;
    QRShape16: TQRShape;
    Qrl_U061: TQRLabel;
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
    QRLabel15: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel16: TQRLabel;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRLabel11: TQRLabel;
    QRL_gmgn: TQRLabel;
    QRShape2: TQRShape;
    QRLabel8: TQRLabel;
    QRL_U059: TQRLabel;
    QRShape17: TQRShape;
    QRLabel9: TQRLabel;
    QRShape18: TQRShape;
    QRLabel10: TQRLabel;
    QRShape22: TQRShape;
    QRL_U063: TQRLabel;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRShape25: TQRShape;
    QRLabel20: TQRLabel;
    QRL_U064: TQRLabel;
    QRL_U065: TQRLabel;
    QRShape26: TQRShape;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure Print;
  private
    UV_iCount, sRec, cRow, iRow, iNo, sum_U060, sum_U061, sum_U059,
    sum_U063, sum_U064, sum_U065 : Integer;
    qlArray : array of TQRLabel;
    Check_Last : Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q674: TfrmLD4Q674;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q67;

{$R *.DFM}

procedure TfrmLD4Q674.QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
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
      qrl_U059.Caption := '';
      qrl_U060.Caption := '';
      qrl_U061.Caption := '';
      qrl_U063.Caption := '';
      qrl_U064.Caption := '';
      qrl_U065.Caption := '';

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
            qrl_U059.Caption   := IntToStr(sum_U059);
            qrl_U060.Caption   := IntToStr(sum_U060);
            qrl_U061.Caption   := IntToStr(sum_U061);
            qrl_U063.Caption   := IntToStr(sum_U063);
            qrl_U064.Caption   := IntToStr(sum_U064);
            qrl_U065.Caption   := IntToStr(sum_U065);
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

procedure TfrmLD4Q674.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q67.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q67.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q67.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
   iNo := 0;
   sum_U060:= 0; sum_U061:= 0; sum_U059:= 0;
   sum_U063:= 0; sum_U064:= 0; sum_U065:= 0;
end;


procedure TfrmLD4Q674.Print;
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
      if      Cell[7,UV_iCount] = 'U059' then
      begin
         qrl_U059.Caption := '*';
         inc(sum_U059);
      end
      else if      Cell[7,UV_iCount] = 'U060' then
      begin
         qrl_U060.Caption := '*';
         inc(sum_U060);
      end
      else if Cell[7,UV_iCount] = 'U061' then
      begin
         qrl_U061.Caption := '*';
         inc(sum_U061);
      end
      else if Cell[7,UV_iCount] = 'U063' then
      begin
         qrl_U063.Caption := '*';
         inc(sum_U063);
      end
      else if Cell[7,UV_iCount] = 'U064' then
      begin
         qrl_U064.Caption := '*';
         inc(sum_U064);
      end
      else if Cell[7,UV_iCount] = 'U065' then
      begin
         qrl_U065.Caption := '*';
         inc(sum_U065);
      end;
   end;
end;

end.
