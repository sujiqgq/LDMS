unit LD2Q081;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, Db, DBTables;

type
  TfrmLD2Q081 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape23: TQRShape;
    QRL_Bunju: TQRLabel;
    QRBand3: TQRBand;
    QRShape27: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRLabel24: TQRLabel;
    QRShape42: TQRShape;
    QRShape47: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel18: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel19: TQRLabel;
    QRL_No: TQRLabel;
    QRLabel21: TQRLabel;
    QRL_Jisa: TQRLabel;
    QRBand4: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape3: TQRShape;
    QRShape51: TQRShape;
    QRLabel14: TQRLabel;
    QRL_samp: TQRLabel;
    qryBunju: TQuery;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRL_U057: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow, UV_iSeq, Sum_iS018, Sum_iT039,Sum_iU019 ,Sum_iU057,Sum_iS064: Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD2Q081: TfrmLD2Q081;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q08;

{$R *.DFM}

procedure TfrmLD2Q081.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD2Q08.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if (Rows = 0) OR (Rows = UV_iCount) then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      QRL_No.Caption    := IntToStr(UV_iSeq);
      QRL_Name.Caption  := Cell[ 7,UV_iCount];
      QRL_Jumin.Caption := Cell[ 8,UV_iCount];
      QRL_samp.Caption  := Cell[ 4,UV_iCount];

      if Cell[ 2,UV_iCount] = 'U057' then
      begin
         QRL_U057.Caption  := '';
 //        Sum_iU057 := Sum_iU057 + 1;
      end
      else
         QRL_U057.Caption  := '*';
{
      if UV_iCount + 1 <= Rows then
      begin
         if Cell[ 6,UV_iCount] = Cell[ 6,UV_iCount + 1] then   //다음행에 항목이 있으면 표시해주고 다음 사람으로 넘어가기 위해..(같은사람 중복 표시 안하도록)
         begin
            if (Cell[ 2,UV_iCount + 1] = 'U057') AND (QRL_U057.Caption = '*') then
            begin
               QRL_U057.Caption  := '';
//               Sum_iU057 := Sum_iU057 + 1;
            end;
            UV_iCount := UV_iCount + 1;
         end;
      end;
}
//      if Cell[ 1,UV_iCount] <> '총합계' then
//      begin
         qryBunju.Active := False;
         qryBunju.ParamByName('cod_bjjs').AsString  := frmLD2Q08.UV_Bjjs;
         qryBunju.ParamByName('num_bunju').AsString := copy(Cell[ 1,UV_iCount],1,4);
         qryBunju.ParamByName('dat_bunju').AsString := Cell[ 12,UV_iCount];
         qryBunju.Active := True;
         if qryBunju.RecordCount > 0 then
         begin
            GP_query_log(qryBunju, 'LD2Q081', 'Q', 'Y');
            QRL_Jisa.Caption  := qryBunju.FieldByName('cod_jisa').AsString;
            QRL_Bunju.Caption := qryBunju.FieldByName('num_bunju').AsString;
         end;
         qryBunju.Active := False;
{
      end
      else
      begin
         QRL_Jisa.Caption  := '';
         QRL_Bunju.Caption := '';
      end;

      if UV_iCount = Rows then
      begin
         QRL_No.Caption    := '';
         QRL_Jisa.Caption  := '';
         QRL_samp.Caption  := '';
         QRL_Bunju.Caption := '';
         QRL_Name.Caption  := '';
         QRL_Jumin.Caption := '총합계';
         QRL_S018.Caption := IntToStr(Sum_iS018);
         QRL_T039.Caption := IntToStr(Sum_iT039);
         QRL_U019.Caption := IntToStr(Sum_iU019);
         QRL_U057.Caption := IntToStr(Sum_iU057);
         QRL_S064.Caption := IntToStr(Sum_iS064);
      end;
}
      inc(UV_iSeq);

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD2Q081.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q08.EdtBDate.Text,1,4) + '년'  +
                       copy(frmLD2Q08.EdtBDate.Text,5,2) + '월'  +
                       copy(frmLD2Q08.EdtBDate.Text,7,2) + '일';
   UV_iCount := 0;
   UV_iSeq := 1;
//   Sum_iS018 := 0;
//   Sum_iT039 := 0;
//   Sum_iU019 := 0;
//   Sum_iS064 := 0;         
end;

end.
