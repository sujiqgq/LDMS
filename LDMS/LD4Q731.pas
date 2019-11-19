//==============================================================================
// ���α׷� ���� : ���� ���
// �ý���        : LDMS
// ��������      : 2015.11.25
// ������        : ������
// ��������      : ���λ���
// �������      :
//==============================================================================


unit LD4Q731;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD4Q731 = class(TForm)
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
  frmLD4Q731: TfrmLD4Q731;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q73;
{$R *.DFM}

procedure TfrmLD4Q731.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';
   MoreData := False;

   with frmLD4Q73.grdMaster do
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
                    ((Cell[ 5,UV_iCount]) = (Cell[ 5,UV_iCount-1])) then
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
         QRL_Bunju.Caption := Cell[5,UV_iCount];
         QRL_Name.Caption  := Cell[6,UV_iCount];
         QRL_samp.Caption  := '(' + Cell[3,UV_iCount] + ')';
         QRL_Jisa.Caption  := '(' + frmLD4Q73.UV_vJisa[UV_iCount-1] + ')';

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

procedure TfrmLD4Q731.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
   UV_iRow := 0;
end;

end.
