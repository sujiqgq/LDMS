unit LD4Q081;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q081 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
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
    QRLabel4: TQRLabel;
    QRShape22: TQRShape;
    QRLabel7: TQRLabel;
    QRShape23: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel10: TQRLabel;
    QRShape5: TQRShape;
    QRLabel15: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape15: TQRShape;
    QRShape32: TQRShape;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    Qrl_Bunju1: TQRLabel;
    qrl_no1: TQRLabel;
    QRShape28: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    qrl_S0071: TQRLabel;
    qrl_S0081: TQRLabel;
    qrl_pS0071: TQRLabel;
    qrl_pS0081: TQRLabel;
    QRShape21: TQRShape;
    QRL_Samp1: TQRLabel;
    QRL_Name1: TQRLabel;
    QRShape27: TQRShape;
    QRShape36: TQRShape;
    QRShape29: TQRShape;
    QRShape4: TQRShape;
    QRShape40: TQRShape;
    lae: TQRLabel;
    QRShape41: TQRShape;
    qrl_rack1: TQRLabel;

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
  frmLD4Q081: TfrmLD4Q081;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q08;

{$R *.DFM}

procedure TfrmLD4Q081.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q08.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
            // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      Qrl_No1.caption := Cell[ 1,UV_iCount];
      Qrl_Rack1.caption  := Cell[ 2,UV_iCount];
      Qrl_samp1.caption   := Cell[ 3,UV_iCount];
      Qrl_Bunju1.Caption  := Cell[ 4,UV_iCount];
      Qrl_name1.Caption := Cell[ 5,UV_iCount];
      if Cellcolor[6,(iRow * 30) + UV_iCount] = $00F9D9F5 then Qrl_pS0071.font.Color := clRed
      else Qrl_pS0071.font.Color := clWindowText;
      if Cellcolor[9,(iRow * 30) + UV_iCount] = $00F9D9F5 then Qrl_S0081.font.Color := clRed
      else Qrl_S0081.font.Color := clWindowText;
      Qrl_S0071.Caption   := Cell[8, UV_iCount];
      Qrl_S0081.Caption   := Cell[9, UV_iCount];
      Qrl_pS0071.Caption  := Cell[6, UV_iCount];
      Qrl_pS0081.Caption  := Cell[7, UV_iCount];



      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;


end;

procedure TfrmLD4Q081.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q08.mskDate.Text,1,4) + '년'  +
                       copy(frmLD4Q08.mskDate.Text,5,2) + '월'  +
                       copy(frmLD4Q08.mskDate.Text,7,2) + '일';

   if frmLD4Q08.CB_HM.ItemIndex = 0 then
   begin
      QRLabel7.Caption := 'S007/현';
      QRLabel9.Caption := 'S007/전';
      QRLabel8.Caption := 'S008/현';
      QRLabel10.Caption := 'S008/전';
  //    QRLabel11.Caption := 'S007/현';
  //    QRLabel13.Caption := 'S007/전';
  //    QRLabel12.Caption := 'S008/현';
  //    QRLabel14.Caption := 'S008/전';
   end
   else if frmLD4Q08.CB_HM.ItemIndex = 1 then
   begin
      QRLabel7.Caption := 'S011/현';
      QRLabel9.Caption := 'S011/전';
      QRLabel8.Caption := 'S012/현';
      QRLabel10.Caption := 'S012/전';
   //   QRLabel11.Caption := 'S011/현';
   //   QRLabel13.Caption := 'S011/전';
    //  QRLabel12.Caption := 'S012/현';
    //  QRLabel14.Caption := 'S012/전';
   end;

   UV_iCount := 0;

   {sRec := Round(frmLD4Q08.grdMaster.Rows / 60);
   if (((frmLD4Q08.grdMaster.Rows / 60) - Round(frmLD4Q08.grdMaster.Rows / 60)) <= 0.5) and
      (((frmLD4Q08.grdMaster.Rows / 60) - Round(frmLD4Q08.grdMaster.Rows / 60)) > 0)    then
       sRec := sRec + 1; }
{   if frmLD4Q08.grdMaster.Rows < 60 then sRec := 0
   else if frmLD4Q08.grdMaster.Rows < 120 then sRec := 1
   else if frmLD4Q08.grdMaster.Rows < 180 then sRec := 2
   else if frmLD4Q08.grdMaster.Rows < 240 then sRec := 3
   else if frmLD4Q08.grdMaster.Rows < 300 then sRec := 4;  }
    {
   if ((frmLD4Q08.grdMaster.Rows mod 60) = 0) then
      sRec := Round(frmLD4Q08.grdMaster.Rows div 60)
   else
   begin
      sRec := Round(frmLD4Q08.grdMaster.Rows div 60) + 1;
   end;

   cRow := 0; iRow := 0;  }
end;


end.
