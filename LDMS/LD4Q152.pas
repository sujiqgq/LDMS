//==============================================================================
// ���α׷� ���� : ���� ���
// �ý���        : LDMS
// ��������      : 2015.06.17
// ������        : ������
// ��������      : ���λ���
// �������      : ���� ������ܰ˻�������1500460
//==============================================================================


unit LD4Q152;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q152 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_Bunju: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_Jisa: TQRLabel;
    qryBunju: TQuery;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, UV_iRow  : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q152: TfrmLD4Q152;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q15;
{$R *.DFM}

procedure TfrmLD4Q152.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';
   MoreData := False;

   with frmLD4Q15.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if (Rows = 0) or (UV_iCount > Rows)  then
      begin
         MoreData := False;
         exit;
      end;

      repeat
         if UV_iCount < Rows then
         begin
           if UV_iCount > 1 then
           begin
              // ����ڷ� ����
               if   ((Cell[ 4,UV_iCount]) = (Cell[ 4,UV_iCount-1])) and
                    ((Cell[ 6,UV_iCount]) = (Cell[ 6,UV_iCount-1])) then
                    sCheck := 'N'
               else sCheck := 'Y';
           end //end of 'if (UV_iCount > 1) - begin'

           else if UV_iCount = 1 then
                   sCheck := 'Y';
         UV_iCount := UV_iCount + 1;
         end //end of 'if (UV_iCount <= Rows) - begin'

         else if (UV_iCount = Rows) then
         begin
         MoreData := False;
         exit;
         end; //end of 'if(UV_iCount > Rows) - begin'

         until (sCheck = 'Y');

         MoreData := True;

         UV_iCount := UV_iCount - 1;

         qryBunju.Active := False;
         QRL_Bunju.Caption := Cell[ 5,UV_iCount];
         QRL_Name.Caption  := Cell[ 6,UV_iCount];
         QRL_samp.Caption  := '(' + Cell[4,UV_iCount] + ')';
         qryBunju.ParamByName('cod_bjjs').AsString   := copy(frmLD4Q15.cmbJisa.Text,1,2);;
         qryBunju.ParamByName('num_bunju').AsString  := QRL_Bunju.Caption;
         qryBunju.ParamByName('dat_bunju').AsString  := copy(Cell[ 2,UV_iCount],1,8);

         qryBunju.Active := True;
         if qryBunju.RecordCount > 0 then
         begin
            GP_query_log(qryBunju, 'LD4Q152', 'Q', 'Y');
            QRL_Jisa.Caption  := qryBunju.FieldByName('cod_jisa').AsString;
         end;
         qryBunju.Active := False;

         UV_iRow := UV_iRow +1;

         if UV_iRow mod 21 = 1 then
         begin
           QRL_Jisa.top := 9;
           QRL_Bunju.top:= 9;
           QRL_Name.top := 27;
           QRL_samp.top := 27;
         end;

         if UV_iRow mod 4 > 1 then
         begin
           QRL_Jisa.top := QRL_Jisa.Top +1;
           QRL_Bunju.top:= QRL_Bunju.Top +1;
           QRL_Name.top := QRL_Name.Top +1;
           QRL_samp.top := QRL_samp.Top +1;
         end;

         UV_iCount := UV_iCount + 1;

   end;  //end of with
end;

procedure TfrmLD4Q152.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
   UV_iRow := 0;
end;

end.
