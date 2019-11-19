unit LD4Q431;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q431 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    QRLabel2: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    Qrl_Bunju: TQRLabel;
    QRShape4: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    qrl_no: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape10: TQRShape;
    QRLabel8: TQRLabel;
    Qrl_S008: TQRLabel;
    Qrl_pS008: TQRLabel;
    QRLabel11: TQRLabel;
    Qrl_pBunju: TQRLabel;
    Qrl_pDatBunju: TQRLabel;
    QRShape14: TQRShape;
    QRLabel14: TQRLabel;
    QRShape18: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRL_name: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, sRec, cRow, iRow : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q431: TfrmLD4Q431;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q43;

{$R *.DFM}

procedure TfrmLD4Q431.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
  if  frmLD4Q43.gubn_Combo.ItemIndex =3 then
      begin
      QRLabel8.caption:='S091(전)';
      QRLabel7.caption:='S091(현)';
      QRLabel1.caption:='S091 전결과 비교LIST';
      end
  else if  frmLD4Q43.gubn_Combo.ItemIndex =1 then
      begin
      QRLabel8.caption:='S008(전)';
      QRLabel7.caption:='S008(현)';
      QRLabel1.caption:='S008 전결과 비교LIST';
      end
  else if  frmLD4Q43.gubn_Combo.ItemIndex =2 then
      begin
      QRLabel8.caption:='S007(전)';
      QRLabel7.caption:='S007(현)';
      QRLabel1.caption:='S007 전결과 비교LIST';
      end
   else if  frmLD4Q43.gubn_Combo.ItemIndex =4 then
      begin
      QRLabel8.caption:='SE02(전)';
      QRLabel7.caption:='SE02(현)';
      QRLabel1.caption:='SE02 전결과 비교LIST';
      end
  else if  frmLD4Q43.gubn_Combo.ItemIndex =5 then
      begin
      QRLabel8.caption:='임신';
      QRLabel7.caption:='';
      QRLabel1.caption:='임신 및 임신 가능성 LIST ';
      end
  else if  frmLD4Q43.gubn_Combo.ItemIndex =6 then
      begin
      QRLabel8.caption:='C047';
      QRLabel7.caption:='';
      QRLabel1.caption:='C047 이상 분주 리스트 ';
      end
  else if  frmLD4Q43.gubn_Combo.ItemIndex =7 then
      begin
      QRLabel8.caption:='H011';
      QRLabel7.caption:='';
      QRLabel1.caption:='H011 이상 분주 리스트 ';
      end
   else if  frmLD4Q43.gubn_Combo.ItemIndex =8 then
      begin
      QRLabel8.caption:='H001';
      QRLabel7.caption:='';
      QRLabel1.caption:='H001 이상 분주 리스트 ';
      end
    else if  frmLD4Q43.gubn_Combo.ItemIndex =9 then
      begin
      QRLabel8.caption:='C083';
      QRLabel7.caption:='';
      QRLabel1.caption:='C083 이상 분주 리스트 ';
      end;


  with frmLD4Q43.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      qrl_no.Caption        := Cell[ 1,UV_iCount];
      Qrl_Bunju.Caption     := Cell[ 3,UV_iCount];
      QRL_name.Caption      := Cell[ 4,UV_iCount];
      Qrl_pDatBunju.Caption := Cell[ 5,UV_iCount];
      Qrl_pBunju.Caption    := Cell[ 6,UV_iCount];
      Qrl_pS008.Caption     := Cell[ 7,UV_iCount];
      Qrl_S008.Caption      := Cell[ 8,UV_iCount];


      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q431.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q43.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q43.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q43.mskDate.Text,7,2) + '일';
   UV_iCount := 0;

end;

end.
