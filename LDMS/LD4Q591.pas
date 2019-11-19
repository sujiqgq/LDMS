unit LD4Q591;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q591 = class(TForm)
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
    QRShape21: TQRShape;
    Qrl_NumBunju: TQRLabel;
    QRShape4: TQRShape;//
    QRLabel4: TQRLabel;
    qrl_no: TQRLabel;
    QRLabel_hm02: TQRLabel;
    QRShape10: TQRShape;
    QRLabel_hm01: TQRLabel;
    Qrl_02: TQRLabel;
    Qrl_01: TQRLabel;
    QRShape14: TQRShape;
    QRLabel14: TQRLabel;
    QRShape18: TQRShape;//
    QRShape22: TQRShape;//
    QRShape23: TQRShape;//
    QRShape24: TQRShape;
    QRL_name: TQRLabel;
    QRLabel3: TQRLabel;
    Qrl_Datbunju: TQRLabel;
    QRLabel9: TQRLabel;
    Qrl_sample: TQRLabel;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape20: TQRShape;
    QRShape28: TQRShape;
    Qrl_03: TQRLabel;
    Qrl_04: TQRLabel;
    QRLabel_hm03: TQRLabel;
    QRLabel_hm04: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q591: TfrmLD4Q591;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q59;

{$R *.DFM}

procedure TfrmLD4Q591.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
  inherited;
  with frmLD4Q59.grdMaster do
   begin
      if (frmLD4Q59.gubn_Comb.ItemIndex = 0) then
      begin
         QRLabel_hm01.Caption := 'APO A-I';
         QRLabel_hm02.Caption := 'APO- B';
         QRLabel_hm03.Caption := 'B/A-I ����';
         QRLabel_hm04.Caption := '���';
      end
      else if (frmLD4Q59.gubn_Comb.ItemIndex = 2) then
      begin
         QRLabel_hm01.Caption := 'Fe';
         QRLabel_hm02.Caption := 'TIBC';
         QRLabel_hm03.Caption := 'ö��ȭ��';
         QRLabel_hm04.Caption := 'UIBC';
      end
      else if (frmLD4Q59.gubn_Comb.ItemIndex = 3) then
      begin
         QRLabel_hm01.Caption := '�Ѵܹ�';
         QRLabel_hm02.Caption := '�˺ι�';
         QRLabel_hm03.Caption := '�۷κθ�';
         QRLabel_hm04.Caption := 'A/G ����';
      end
      else if (frmLD4Q59.gubn_Comb.ItemIndex = 4) then
      begin
         QRLabel_hm01.Caption := '��������';
         QRLabel_hm02.Caption := 'ũ����Ƽ��';
         QRLabel_hm03.Caption := 'BUN/Cr ����';
         QRLabel_hm04.Caption := 'GFR';
      end;

      // �ڷᰡ ��������� ó��
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // ����ڷ� ����
      UV_iCount := UV_iCount + 1;

      qrl_no.Caption       := Cell[ 1,UV_iCount];
      Qrl_DatBunju.Caption := Cell[ 2,UV_iCount];
      Qrl_NumBunju.Caption := Cell[ 3,UV_iCount];
      QRL_sample.Caption   := Cell[ 4,UV_iCount];
      QRL_name.Caption     := Cell[ 5,UV_iCount];
      Qrl_01.Caption       := Cell[ 6,UV_iCount];
      Qrl_02.Caption       := Cell[ 7,UV_iCount];
      Qrl_03.Caption       := Cell[ 8,UV_iCount];
      Qrl_04.Caption       := Cell[ 9,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q591.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q59.mskDate.Text,1,4) + '��'  +
                       copy(frmLD4Q59.mskDate.Text,5,2) + '��'  +
                       copy(frmLD4Q59.mskDate.Text,7,2) + '��';
   UV_iCount := 0;

end;

end.
