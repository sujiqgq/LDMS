unit LD5Q101;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls, StdCtrls;

type
  TfrmLD5Q101 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRShape1: TQRShape;
    QRL_date: TQRLabel;
    QRBand3: TQRBand;
    QRShape3: TQRShape;
    QRL_Hangmok: TQRLabel;
    QRL_Bunju: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    UV_iCount, UV_Cnt, UV_Count : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD5Q101: TfrmLD5Q101;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q10;

{$R *.DFM}

procedure TfrmLD5Q101.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   vHangmok : String;
begin
   inherited;

   with frmLD5Q10.grdMaster do
   begin
      // �ڷᰡ ��������� ó��
      if (Rows = 0) or (UV_iCount > Rows) then
      begin
         MoreData := False;
         exit;
      end;

      // ������ ��...
      if UV_Cnt = 0 then
      begin
         UV_Count := length(Cell[5,  UV_iCount]) div 3990;
         if UV_Count > 0 then
         begin
            if (length(Cell[5,  UV_iCount]) mod 3990) >= 1 then
               UV_Count := UV_Count + 1;
            UV_Cnt := 1;
         end;
      end;

      if UV_Count = 0 then
      begin
         // ����ڷ� ����
         QRL_Bunju.Caption   := Cell[1,  UV_iCount] +' [' + Cell[2,  UV_iCount] + '] (' + Cell[3,  UV_iCount] + ') ' + Cell[4,  UV_iCount];
         QRL_Hangmok.Caption := Cell[5,  UV_iCount];
      end
      else
      begin
         vHangmok := copy(Cell[5,  UV_iCount], (UV_Cnt * 3990) - 3989 ,3990);
         // ����ڷ� ����
         QRL_Bunju.Caption   := Cell[1,  UV_iCount] +' [' + Cell[2,  UV_iCount] + '] (' + Cell[3,  UV_iCount] + ') ' + Cell[4,  UV_iCount];
         QRL_Hangmok.Caption := vHangmok;
      end;

      if UV_iCount <= Rows then
      begin
         if UV_Cnt = UV_Count then
         begin
            MoreData := True;
            Inc(UV_iCount);
            UV_Cnt   := 0;
         end
         else
         begin
            MoreData := True;
            Inc(UV_Cnt);
         end;
      end
      else
         MoreData := False;
   end;

end;

procedure TfrmLD5Q101.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := '�� �������� : ' +
                       copy(frmLD5Q10.msk_Bgmgn.Text,1,4) + '-'  +
                       copy(frmLD5Q10.msk_Bgmgn.Text,5,2) + '-'  +
                       copy(frmLD5Q10.msk_Bgmgn.Text,7,2);;
   UV_iCount  := 1;
   UV_Cnt     := 0;
end;

procedure TfrmLD5Q101.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
    vCount :Integer;
begin
   vCount := 0;
   vCount := length(QRL_Hangmok.Caption) div 70;
   if (length(QRL_Hangmok.Caption) mod 70) >= 1 then;
      vCount := vCount + 1;

   if vCount = 0 then
   begin
      QRBand3.Height     := 25;
      QRL_BUNJU.Height   := 25;
      QRL_Hangmok.Height := 25;
      QRShape3.Height    := 25;
   end
   else if vCount <= 57 then
   begin
      QRBand3.Height     := (vCount * 17) + 10;
      QRL_BUNJU.Height   := (vCount * 17) +  5;
      QRL_Hangmok.Height := (vCount * 17) + 10;
      QRShape3.Height    := (vCount * 17) + 10;
   end
   else if vCount > 57 then
   begin
      QRBand3.Height     := (57 * 17);
      QRL_BUNJU.Height   := (57 * 17);
      QRL_Hangmok.Height := (57 * 17);
      QRShape3.Height    := (57 * 17);
   end;
end;

end.
