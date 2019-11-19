unit LD8Q021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD8Q021 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRBand3: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRL_BUNJU: TQRLabel;
    QRL_CODE: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_date: TQRLabel;
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
  frmLD8Q021: TfrmLD8Q021;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q02;

{$R *.DFM}

procedure TfrmLD8Q021.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   vHangmok : String;
begin
   inherited;

   with frmLD8Q02.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 페이지 수...
      if UV_Cnt = 0 then
      begin
         UV_Count := length(Cell[3,  UV_iCount]) div 3990;
         if UV_Count > 0 then
         begin
            if (length(Cell[3,  UV_iCount]) mod 3990) >= 1 then
               UV_Count := UV_Count + 1;
            UV_Cnt := 1;
         end;
      end;

      if UV_Count = 0 then
      begin
         // 출력자료 전달
         QRL_CODE.Caption  := Cell[1,  UV_iCount];
         QRL_NAME.Caption  := Cell[2,  UV_iCount];
         QRL_BUNJU.Caption := Cell[3,  UV_iCount];
      end
      else
      begin
         vHangmok := copy(Cell[3,  UV_iCount], (UV_Cnt * 3990) - 3989 ,3990);
         // 출력자료 전달
         QRL_CODE.Caption  := Cell[1,  UV_iCount];
         QRL_NAME.Caption  := Cell[2,  UV_iCount];
         QRL_BUNJU.Caption := vHangmok;
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

procedure TfrmLD8Q021.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := '※ 분주일자 : ' +
                       copy(frmLD8Q02.msk_Bgmgn.Text,1,4) + '-'  +
                       copy(frmLD8Q02.msk_Bgmgn.Text,5,2) + '-'  +
                       copy(frmLD8Q02.msk_Bgmgn.Text,7,2) + ' [' +
                       frmLD8Q02.Com_Part.Items[frmLD8Q02.Com_Part.ItemIndex] + ' ]';
   UV_iCount  := 1;
   UV_Cnt     := 0;
end;

procedure TfrmLD8Q021.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
    vCount :Integer;
begin
   vCount := 0;
   vCount := length(QRL_BUNJU.Caption) div 70;
   if (length(QRL_BUNJU.Caption) mod 70) >= 1 then;
      vCount := vCount + 1;

   if vCount = 0 then
   begin
      QRBand3.Height   := 25;
      QRL_BUNJU.Height := 17;
      QRShape3.Height  := 25;
      QRShape4.Height  := 25;
   end
   else if vCount <= 57 then
   begin
      QRBand3.Height   := (vCount * 17) + 20;
      QRL_BUNJU.Height := (vCount * 17) + 20;
      QRShape3.Height  := (vCount * 17) + 15;
      QRShape4.Height  := (vCount * 17) + 20;
   end


   else if vCount <= 57 then
   begin

   end
   else if vCount > 57 then
   begin
      QRBand3.Height   := (57 * 17);
      QRL_BUNJU.Height := (57 * 17);
      QRShape3.Height  := (57 * 17);
      QRShape4.Height  := (57 * 17);
   end;
end;

end.
