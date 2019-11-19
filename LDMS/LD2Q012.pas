unit LD2Q012;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD2Q012 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    Qrl_Bunju1: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    qrl_samp: TQRLabel;
    qrl_jumin: TQRLabel;
    qrl_name: TQRLabel;
    Qrl_chuga: TQRLabel;
    QRShape21: TQRShape;
    QRShape4: TQRShape;
    Qrl_Dat_Gumgin: TQRLabel;
    QRShape9: TQRShape;
    QRShape11: TQRShape;
    Qrl_jisa: TQRLabel;
    QRBand1: TQRBand;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape22: TQRShape;
    QRLabel7: TQRLabel;
    QRShape23: TQRShape;
    QRLabel8: TQRLabel;
    QRLchuga: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape5: TQRShape;
    QRShape10: TQRShape;
    QRLabel4: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount: Integer;   //, sRec, cRow, iRow
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD2Q012: TfrmLD2Q012;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q01;

{$R *.DFM}

procedure TfrmLD2Q012.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD2Q01.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      if (frmLD2Q01.RG_Query.ItemIndex= 1) then
      BEGIN
          QRLabel1.Caption := 'Y005,Y008,U017,U019,U046,U058,U035,U015,U016 분주List조회';
          QRLabel1.Font.Size :=  14;
      end
      else if (frmLD2Q01.RG_Query.ItemIndex = 2) then QRLabel1.Caption := 'U060,U061분주List조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 3) then QRLabel1.Caption := '메디젠 유전자 분주Lis조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 4) then QRLabel1.Caption := 'SK 유전자 분주List조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 5) then QRLabel1.Caption := 'NK 세포활성[E068] 분주List조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 6) then QRLabel1.Caption := '케어링크 유전자 검사List조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 7) then QRLabel1.Caption := '중금속4종검사 List조회'
      else if (frmLD2Q01.RG_Query.ItemIndex = 8) then QRLabel1.Caption := 'QuantiFERON-TB(잠복결핵) 분주List조회';


      if (frmLD2Q01.RG_Query.ItemIndex = 7) or (frmLD2Q01.RG_Query.ItemIndex = 8) or (frmLD2Q01.RG_Query.ItemIndex = 11)  then
      begin
         QRLabel2.Caption := 'No.'; //20140415 곽윤설
         // 출력자료 전달
         UV_iCount := UV_iCount + 1;
         Qrl_Bunju1.Caption       := Cell[1,UV_iCount];
         Qrl_jisa.Caption         := Cell[7,UV_iCount];
         Qrl_samp.Caption         := Cell[2,UV_iCount];
         Qrl_name.Caption         := copy(Cell[3,UV_iCount],1,10);
         Qrl_jumin.Caption        := Cell[4,UV_iCount];
         Qrl_Dat_Gumgin.Caption   := Cell[5,UV_iCount];
         Qrl_chuga.Caption        := Cell[11,UV_iCount];
      end
      else
      begin
         // 출력자료 전달
         UV_iCount := UV_iCount + 1;
         Qrl_Bunju1.Caption       := Cell[1,UV_iCount];
         Qrl_jisa.Caption         := Cell[7,UV_iCount];
         Qrl_samp.Caption         := Cell[2,UV_iCount];
         Qrl_name.Caption         := copy(Cell[3,UV_iCount],1,10);
         Qrl_jumin.Caption        := Cell[4,UV_iCount];
         Qrl_Dat_Gumgin.Caption   := Cell[5,UV_iCount];
         Qrl_chuga.Caption        := Cell[11,UV_iCount];
      end;
      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD2Q012.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD2Q01.mskDate.Text,1,4) + '년'  +
                       copy(frmLD2Q01.mskDate.Text,5,2) + '월'  +
                       copy(frmLD2Q01.mskDate.Text,7,2) + '일';
   UV_iCount := 0;

  { sRec := Round(frmLD4Q08.grdMaster.Rows / 60);
   if (((frmLD4Q08.grdMaster.Rows / 60) - Round(frmLD4Q08.grdMaster.Rows / 60)) <= 0.5) and
      (((frmLD4Q08.grdMaster.Rows / 60) - Round(frmLD4Q08.grdMaster.Rows / 60)) > 0)    then
       sRec := sRec + 1; }
 { if frmLD2Q01.grdMaster.Rows < 60 then sRec := 0
   else if frmLD2Q01.grdMaster.Rows < 120 then sRec := 1
   else if frmLD2Q01.grdMaster.Rows < 180 then sRec := 2
   else if frmLD2Q01.grdMaster.Rows < 240 then sRec := 3
   else if frmLD2Q01.grdMaster.Rows < 300 then sRec := 4
   else if frmLD2Q01.grdMaster.Rows < 360 then sRec := 5
   else if frmLD2Q01.grdMaster.Rows < 420 then sRec := 6;
           }
 
end;

end.
