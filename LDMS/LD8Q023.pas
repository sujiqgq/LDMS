unit LD8Q023;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD8Q023 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRL_2: TQRLabel;
    QRL_9: TQRLabel;
    QRL_16: TQRLabel;
    QRL_23: TQRLabel;
    QRL_3: TQRLabel;
    QRL_4: TQRLabel;
    QRL_5: TQRLabel;
    QRL_6: TQRLabel;
    QRL_7: TQRLabel;
    QRL_10: TQRLabel;
    QRL_11: TQRLabel;
    QRL_12: TQRLabel;
    QRL_13: TQRLabel;
    QRL_14: TQRLabel;
    QRL_17: TQRLabel;
    QRL_18: TQRLabel;
    QRL_19: TQRLabel;
    QRL_20: TQRLabel;
    QRL_21: TQRLabel;
    QRL_24: TQRLabel;
    QRL_25: TQRLabel;
    QRL_26: TQRLabel;
    QRL_27: TQRLabel;
    QRL_28: TQRLabel;
    QRL_Hangmok: TQRLabel;
    QRL_1: TQRLabel;
    QRL_15: TQRLabel;
    QRL_22: TQRLabel;
    QRShape13: TQRShape;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRL_date: TQRLabel;
    QRL_8: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, UV_iPage, UV_iSpace, UV_iTemp : Integer;
    UV_sHangmok : String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD8Q023: TfrmLD8Q023;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q02, LD8Q022, LD8Q024;

{$R *.DFM}

procedure TfrmLD8Q023.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   i, j, iPage, iSpace : integer;
begin
   inherited;

   with frmLD8Q02.grdMaster do
   begin
      // 자료가 없을경우의 처리 1
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      // 자료가 없을경우의 처리 2
      if trim(Cell[3,UV_iSpace]) = '' then
      begin
         MoreData := False;
         exit;
      end;

      if UV_iTemp < UV_iPage then
         MoreData := True
      else
         MoreData := False;

      if UV_iTemp = 0 then
      begin
         QRL_Hangmok.Caption := 'RA정량[S004]';
         QRL_date.Caption := '분주일:' +
                       copy(frmLD8Q02.msk_Bgmgn.Text,1,4) + '-'  +
                       copy(frmLD8Q02.msk_Bgmgn.Text,5,2) + '-'  +
                       copy(frmLD8Q02.msk_Bgmgn.Text,7,2);
      end
      else
      begin
         QRL_Hangmok.Caption := '';
         QRL_date.Caption := '';
      end;

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 1 then
      begin
         QRL_1.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 1, 4);
         
         if (UV_iTemp <> 0) AND (copy(QRL_28.Caption,1,1) <> '/') then
         begin
            if (StrToInt(QRL_1.Caption) <> StrToInt(QRL_28.Caption) + 1) then
               QRL_1.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 1, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else if (UV_iTemp <> 0) AND (copy(QRL_28.Caption,1,1) = '/') then
         begin
            if (StrToInt(QRL_1.Caption) <> StrToInt(copy(QRL_28.Caption,2,4)) + 1) then
               QRL_1.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 1, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_1.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 6 then
      begin
         QRL_2.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 6, 4);

         if copy(QRL_1.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_2.Caption) <> StrToInt(copy(QRL_1.Caption,2,4)) + 1) then
               QRL_2.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 6, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_2.Caption) <> StrToInt(QRL_1.Caption) + 1) then
               QRL_2.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 6, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_2.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 11 then
      begin
         QRL_3.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 11, 4);

         if copy(QRL_2.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_3.Caption) <> StrToInt(copy(QRL_2.Caption,2,4)) + 1) then
               QRL_3.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 11, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_3.Caption) <> StrToInt(QRL_2.Caption) + 1) then
               QRL_3.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 11, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_3.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 16 then
      begin
         QRL_4.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 16, 4);

         if copy(QRL_3.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_4.Caption) <> StrToInt(copy(QRL_3.Caption,2,4)) + 1) then
               QRL_4.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 16, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_4.Caption) <> StrToInt(QRL_3.Caption) + 1) then
               QRL_4.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 16, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_4.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 21 then
      begin
         QRL_5.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 21, 4);

         if copy(QRL_4.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_5.Caption) <> StrToInt(copy(QRL_4.Caption,2,4)) + 1) then
               QRL_5.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 21, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_5.Caption) <> StrToInt(QRL_4.Caption) + 1) then
               QRL_5.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 21, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_5.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 26 then
      begin
         QRL_6.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 26, 4);

         if copy(QRL_5.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_6.Caption) <> StrToInt(copy(QRL_5.Caption,2,4)) + 1) then
               QRL_6.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 26, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_6.Caption) <> StrToInt(QRL_5.Caption) + 1) then
               QRL_6.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 26, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_6.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 31 then
      begin
         QRL_7.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 31, 4);

         if copy(QRL_6.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_7.Caption) <> StrToInt(copy(QRL_6.Caption,2,4)) + 1) then
               QRL_7.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 31, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_7.Caption) <> StrToInt(QRL_6.Caption) + 1) then
               QRL_7.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 31, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_7.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 36 then
      begin
         QRL_8.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 36, 4);

         if copy(QRL_7.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_8.Caption) <> StrToInt(copy(QRL_7.Caption,2,4)) + 1) then
               QRL_8.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 36, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_8.Caption) <> StrToInt(QRL_7.Caption) + 1) then
               QRL_8.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 36, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_8.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 41 then
      begin
         QRL_9.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 41, 4);

         if copy(QRL_8.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_9.Caption) <> StrToInt(copy(QRL_8.Caption,2,4)) + 1) then
               QRL_9.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 41, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_9.Caption) <> StrToInt(QRL_8.Caption) + 1) then
               QRL_9.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 41, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_9.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 46 then
      begin
         QRL_10.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 46, 4);

         if copy(QRL_9.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_10.Caption) <> StrToInt(copy(QRL_9.Caption,2,4)) + 1) then
               QRL_10.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 46, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_10.Caption) <> StrToInt(QRL_9.Caption) + 1) then
               QRL_10.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 46, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_10.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 51 then
      begin
         QRL_11.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 51, 4);

         if copy(QRL_10.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_11.Caption) <> StrToInt(copy(QRL_10.Caption,2,4)) + 1) then
               QRL_11.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 51, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_11.Caption) <> StrToInt(QRL_10.Caption) + 1) then
               QRL_11.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 51, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_11.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 56 then
      begin
         QRL_12.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 56, 4);

         if copy(QRL_11.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_12.Caption) <> StrToInt(copy(QRL_11.Caption,2,4)) + 1) then
               QRL_12.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 56, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_12.Caption) <> StrToInt(QRL_11.Caption) + 1) then
               QRL_12.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 56, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_12.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 61 then
      begin
         QRL_13.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 61, 4);

         if copy(QRL_12.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_13.Caption) <> StrToInt(copy(QRL_12.Caption,2,4)) + 1) then
               QRL_13.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 61, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_13.Caption) <> StrToInt(QRL_12.Caption) + 1) then
               QRL_13.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 61, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_13.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 66 then
      begin
         QRL_14.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 66, 4);

         if copy(QRL_13.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_14.Caption) <> StrToInt(copy(QRL_13.Caption,2,4)) + 1) then
               QRL_14.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 66, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_14.Caption) <> StrToInt(QRL_13.Caption) + 1) then
               QRL_14.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 66, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_14.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 71 then
      begin
         QRL_15.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 71, 4);

         if copy(QRL_14.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_15.Caption) <> StrToInt(copy(QRL_14.Caption,2,4)) + 1) then
               QRL_15.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 71, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_15.Caption) <> StrToInt(QRL_14.Caption) + 1) then
               QRL_15.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 71, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_15.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 76 then
      begin
         QRL_16.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 76, 4);

         if copy(QRL_15.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_16.Caption) <> StrToInt(copy(QRL_15.Caption,2,4)) + 1) then
               QRL_16.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 76, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_16.Caption) <> StrToInt(QRL_15.Caption) + 1) then
               QRL_16.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 76, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_16.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 81 then
      begin
         QRL_17.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 81, 4);

         if copy(QRL_16.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_17.Caption) <> StrToInt(copy(QRL_16.Caption,2,4)) + 1) then
               QRL_17.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 81, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_17.Caption) <> StrToInt(QRL_16.Caption) + 1) then
               QRL_17.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 81, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_17.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 86 then
      begin
         QRL_18.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 86, 4);

         if copy(QRL_17.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_18.Caption) <> StrToInt(copy(QRL_17.Caption,2,4)) + 1) then
               QRL_18.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 86, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_18.Caption) <> StrToInt(QRL_17.Caption) + 1) then
               QRL_18.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 86, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_18.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 91 then
      begin
         QRL_19.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 91, 4);

         if copy(QRL_18.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_19.Caption) <> StrToInt(copy(QRL_18.Caption,2,4)) + 1) then
               QRL_19.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 91, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_19.Caption) <> StrToInt(QRL_18.Caption) + 1) then
               QRL_19.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 91, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_19.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 96 then
      begin
         QRL_20.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 96, 4);

         if copy(QRL_19.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_20.Caption) <> StrToInt(copy(QRL_19.Caption,2,4)) + 1) then
               QRL_20.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 96, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_20.Caption) <> StrToInt(QRL_19.Caption) + 1) then
               QRL_20.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 96, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_20.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 101 then
      begin
         QRL_21.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 101, 4);

         if copy(QRL_20.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_21.Caption) <> StrToInt(copy(QRL_20.Caption,2,4)) + 1) then
               QRL_21.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 101, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_21.Caption) <> StrToInt(QRL_20.Caption) + 1) then
               QRL_21.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 101, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_21.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 106 then
      begin
         QRL_22.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 106, 4);

         if copy(QRL_21.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_22.Caption) <> StrToInt(copy(QRL_21.Caption,2,4)) + 1) then
               QRL_22.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 106, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_22.Caption) <> StrToInt(QRL_21.Caption) + 1) then
               QRL_22.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 106, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_22.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 111 then
      begin
         QRL_23.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 111, 4);

         if copy(QRL_22.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_23.Caption) <> StrToInt(copy(QRL_22.Caption,2,4)) + 1) then
               QRL_23.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 111, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_23.Caption) <> StrToInt(QRL_22.Caption) + 1) then
               QRL_23.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 111, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_23.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 116 then
      begin
         QRL_24.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 116, 4);

         if copy(QRL_23.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_24.Caption) <> StrToInt(copy(QRL_23.Caption,2,4)) + 1) then
               QRL_24.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 116, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_24.Caption) <> StrToInt(QRL_23.Caption) + 1) then
               QRL_24.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 116, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_24.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 121 then
      begin
         QRL_25.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 121, 4);

         if copy(QRL_24.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_25.Caption) <> StrToInt(copy(QRL_24.Caption,2,4)) + 1) then
               QRL_25.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 121, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_25.Caption) <> StrToInt(QRL_24.Caption) + 1) then
               QRL_25.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 121, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_25.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 126 then
      begin
         QRL_26.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 126, 4);

         if copy(QRL_25.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_26.Caption) <> StrToInt(copy(QRL_25.Caption,2,4)) + 1) then
               QRL_26.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 126, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_26.Caption) <> StrToInt(QRL_25.Caption) + 1) then
               QRL_26.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 126, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_26.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 131 then
      begin
         QRL_27.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 131, 4);

         if copy(QRL_26.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_27.Caption) <> StrToInt(copy(QRL_26.Caption,2,4)) + 1) then
               QRL_27.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 131, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_27.Caption) <> StrToInt(QRL_26.Caption) + 1) then
               QRL_27.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 131, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_27.Caption  := '';

      if length(UV_sHangmok) > (UV_iTemp) * 5 * 28 + 136 then
      begin
         QRL_28.Caption  := copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 136, 4);

         if copy(QRL_27.Caption,1,1) = '/' then
         begin
            if (StrToInt(QRL_28.Caption) <> StrToInt(copy(QRL_27.Caption,2,4)) + 1) then
               QRL_28.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 136, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end
         else
         begin
            if (StrToInt(QRL_28.Caption) <> StrToInt(QRL_27.Caption) + 1) then
               QRL_28.Caption  := '/' + copy(UV_sHangmok, (UV_iTemp) * 5 * 28 + 136, 4);  // 연속되는 분주번호가 아닐경우 '/'표시
         end;
      end
      else
         QRL_28.Caption  := '';

      UV_iTemp := UV_iTemp + 1;
   end;
end;

procedure TfrmLD8Q023.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var
   i : integer;
begin
   inherited;

   UV_iSpace := 0;
   for i := 0 to frmLD8Q02.grdMaster.Rows - 1 do
   begin
      if (trim(frmLD8Q02.grdMaster.Cell[1,i]) = 'S004') then
      begin
         UV_iSpace := i;
         if trim(frmLD8Q02.grdMaster.Cell[3,i]) <> '' then
         begin
            UV_sHangmok := trim(frmLD8Q02.grdMaster.Cell[3,i]);
         end;
      end;
   end;

   UV_iTemp := 0;
   UV_iCount := 0;   //항목갯수
   UV_iCount := (length(UV_sHangmok) + 1) div 5;

   UV_iPage := 0;   //페이지수
   UV_iPage := UV_iCount div 28;
   if UV_iCount mod 28 > 0 then
      UV_iPage := UV_iPage + 1;
end;

end.
