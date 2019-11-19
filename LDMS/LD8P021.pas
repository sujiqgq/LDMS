//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : MDMS
// 수정일자      : 2007.10.30
// 수정자        : 유동구
// 수정내용      : 퀘리에서 한줄 삭제(7000~8000 출력가능하도록)
// 참고사항      :   and substring(num_bunju,1,1) not in ('7','8')
//==============================================================================
// 수정일자      : 2015.03.06
// 수정자        : 곽윤설
// 수정내용      : 렉번호 출력 (QRLabel1)
// 참고사항      : 본원 - 엄하니
//==============================================================================

unit LD8P021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables, QBarcode;

type
  TfrmLD8P021 = class(TForm)
    qryGulkwa: TQuery;
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    BarCode: TBarCode;
    QRLabel1: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Extended;
    { Private declarations }
  public
    UV_istart, UV_iend : Extended;
    UV_Check : boolean;
    UV_Date : String;
    { Public declarations }
  end;

var
  frmLD8P021: TfrmLD8P021;

implementation

uses ZZFunc, MdmsFunc, LD8P02;
{$R *.DFM}

procedure TfrmLD8P021.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  VAR sjisa : string;
begin
   if UV_Check then
   begin
      if UV_iCount > UV_iend Then
      begin
         MoreData := False;
         Exit;
      end;
      BarCode.Text := '';
//      QRLabel3.Caption := '';
      QRLabel4.Caption := ''; QRLabel5.Caption := ''; QRLabel6.Caption := '';
      QRLabel7.Caption := ''; QRLabel8.Caption := ''; QRLabel1.Caption := '';

      BarCode.Text    := copy(qryGulkwa.FieldByName('dat_gmgn').AsString,3,6) + qryGulkwa.FieldByName('num_samp').AsString;
//      QRLabel3.Caption := '*' + copy(qryGulkwa.FieldByName('dat_gmgn').AsString,3,6) + qryGulkwa.FieldByName('num_samp').AsString + '*';
      QRLabel4.Caption := qryGulkwa.FieldByName('desc_name').AsString;
      QRLabel1.Caption := qryGulkwa.FieldByName('desc_RackNo').AsString;

      // 주민번호(생년월일) -> MM-DD-성별-센터-YY로 나열 ..20170922..수지
           if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '15' then sjisa := '0'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '12' then sjisa := '2'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '11' then sjisa := '1'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '43' then sjisa := '4'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '71' then sjisa := '7'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '61' then sjisa := '6'
      else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '51' then sjisa := '5';
      case StrToInt(copy(qryGulkwa.FieldByName('num_jumin').AsString,7,1)) of
         1,3,5,7 : QRLabel5.caption := Copy(qryGulkwa.FieldByName('num_jumin').AsString, 3, 4) + 'M' + sjisa + Copy(qryGulkwa.FieldByName('num_jumin').AsString, 1, 2);
         2,4,6,8 : QRLabel5.caption := Copy(qryGulkwa.FieldByName('num_jumin').AsString, 3, 4) + 'F' + sjisa + Copy(qryGulkwa.FieldByName('num_jumin').AsString, 1, 2);
      end;

      if (pos('DI01', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) or
         (pos('DI02', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) then
      begin
         QRLabel6.Caption := '대사동의';
      end
      else if (pos('DI03', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) then
      begin
         QRLabel6.Caption := '대사재검';
      end;

      QRLabel7.Caption := copy(qryGulkwa.FieldByName('dat_gmgn').AsString,1,4) + '-' +
                          copy(qryGulkwa.FieldByName('dat_gmgn').AsString,5,2) + '-' +
                          copy(qryGulkwa.FieldByName('dat_gmgn').AsString,7,2) + '[' +
                          qryGulkwa.FieldByName('num_samp').AsString + ']';

      QRLabel8.Caption  := qryGulkwa.FieldByName('desc_dc').AsString;
      QRLabel9.Caption  := qryGulkwa.FieldByName('num_bunju').AsString;
      QRLabel1.Caption  := qryGulkwa.FieldByName('desc_rackNo').AsString;

      if UV_iCount <= UV_iend Then
      begin
         MoreData := True;
         UV_iCount := UV_iCount + 1;
      end
      else Moredata := False;
   end
   else
   begin
      with qryGulkwa Do
      begin
         if Eof Then
         begin
            MoreData := False;
            Exit;
         end;

         BarCode.Text := '';
//         QRLabel3.Caption := '';
         QRLabel4.Caption := ''; QRLabel5.Caption := ''; QRLabel6.Caption := '';
         QRLabel7.Caption := ''; QRLabel8.Caption := ''; QRLabel1.Caption := '';

         BarCode.Text    := copy(qryGulkwa.FieldByName('dat_gmgn').AsString,3,6) + qryGulkwa.FieldByName('num_samp').AsString;
//         QRLabel3.Caption := '*' + copy(qryGulkwa.FieldByName('dat_gmgn').AsString,3,6) + qryGulkwa.FieldByName('num_samp').AsString + '*';
         QRLabel4.Caption := qryGulkwa.FieldByName('desc_name').AsString;
         // 주민번호(생년월일) -> MM-DD-성별-센터-YY로 나열 ..20170922..수지
             if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '15' then sjisa := '0'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '12' then sjisa := '2'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '11' then sjisa := '1'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '43' then sjisa := '4'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '71' then sjisa := '7'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '61' then sjisa := '6'
        else if Copy(frmLD8P02.cmbbjjs.Text, 1, 2) = '51' then sjisa := '5';
        case StrToInt(copy(qryGulkwa.FieldByName('num_jumin').AsString,7,1)) of
           1,3,5,7 : QRLabel5.caption := Copy(qryGulkwa.FieldByName('num_jumin').AsString, 3, 4) + 'M' + sjisa + Copy(qryGulkwa.FieldByName('num_jumin').AsString, 1, 2);
           2,4,6,8 : QRLabel5.caption := Copy(qryGulkwa.FieldByName('num_jumin').AsString, 3, 4) + 'F' + sjisa + Copy(qryGulkwa.FieldByName('num_jumin').AsString, 1, 2);
        end;

         if (pos('DI01', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) or
            (pos('DI02', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) then
         begin
            QRLabel6.Caption := '대사동의';
         end
         else if (pos('DI03', qryGulkwa.FieldByName('cod_chuga').AsString) > 0) then
         begin
            QRLabel6.Caption := '대사재검';
         end;

         QRLabel7.Caption := copy(qryGulkwa.FieldByName('dat_gmgn').AsString,1,4) + '-' +
                             copy(qryGulkwa.FieldByName('dat_gmgn').AsString,5,2) + '-' +
                             copy(qryGulkwa.FieldByName('dat_gmgn').AsString,7,2);
                             { + '[' + qryGulkwa.FieldByName('num_samp').AsString + ']';}

         QRLabel8.Caption  := qryGulkwa.FieldByName('desc_dc').AsString;
         QRLabel9.Caption  := qryGulkwa.FieldByName('num_samp').AsString;
         QRLabel1.Caption  := qryGulkwa.FieldByName('desc_rackNo').AsString;

         Next;

         if UV_iCount <= RecordCount Then
         begin
            MoreData := True;
            UV_iCount := UV_iCount + 1;
         end
         else
            Moredata := False;
      end; // end of with;
   end;
end;

procedure TfrmLD8P021.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := UV_istart;
end;

end.
