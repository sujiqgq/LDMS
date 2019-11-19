//==============================================================================
// 프로그램 설명 : 생물학적 노출지표검사 대상자조회
// 시스템        : 통합검진
// 개발일자      : 2017.08.08
// 개발자        : 유동구
// 참고사항      : 신규개발(한의 재단특검행정파트1701202)
//==============================================================================
unit LD5Q151;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD5Q151 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRShape2: TQRShape;
    QRBand2: TQRBand;
    QRBand3: TQRBand;
    QRBand4: TQRBand;
    QRShape16: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape17: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape1: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRLabel3: TQRLabel;
    QrlNo: TQRLabel;
    QrrName: TQRRichText;
    QrrBirth: TQRRichText;
    QrrDanche: TQRRichText;
    QrrMatter: TQRRichText;
    QrrHangmok: TQRRichText;
    QrlSample_No: TQRLabel;
    QrlDat_gmgn: TQRLabel;
    QrlTemp: TQRLabel;
    QrlTime: TQRLabel;
    QRLabel4: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD5Q151: TfrmLD5Q151;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q15;

{$R *.DFM}

procedure TfrmLD5Q151.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q15.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      QrrName.Lines.Clear;
      QrrBirth.Lines.Clear;
      QrrDanche.Lines.Clear;
      QrrMatter.Lines.Clear;
      QrrHangmok.Lines.Clear;

      QrlNo.Caption         := Cell[ 1,UV_iCount];
      QrrName.Lines.Text    := Cell[ 2,UV_iCount];
      QrrBirth.Lines.Text   := copy(Cell[ 3,UV_iCount],1,8);
      QrlSample_No.Caption  := Cell[ 4,UV_iCount];
      QrrDanche.Lines.Text  := Cell[ 5,UV_iCount];
      QrlDat_gmgn.Caption   := Cell[ 6,UV_iCount];
      QrrMatter.Lines.Text  := Cell[ 7,UV_iCount];
      QrrHangmok.Lines.Text := Cell[ 8,UV_iCount];
      QrlTemp.Caption       := Cell[ 9,UV_iCount];
      QrlTime.Caption       := Cell[10,UV_iCount];

      if UV_iCount <= Rows then begin MoreData := True; Inc(UV_iCount); end
      else                            MoreData := False;

   end;
end;

procedure TfrmLD5Q151.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   UV_iCount := 1;
end;

procedure TfrmLD5Q151.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
   iCnt : Integer;
begin
   inherited;

   //가장 라인수 많은 걸 찾기...
   iCnt := QrrName.Lines.Count;
   if iCnt < QrrBirth.Lines.Count   then iCnt := QrrBirth.Lines.Count;
   if iCnt < QrrDanche.Lines.Count  then iCnt := QrrDanche.Lines.Count;
   if iCnt < QrrMatter.Lines.Count  then iCnt := QrrMatter.Lines.Count;
   if iCnt < QrrHangmok.Lines.Count then iCnt := QrrHangmok.Lines.Count;

   if iCnt = 1 then
   begin
      QRBand2.Height    := 25;

      QrrName.Height    := 25;
      QrrBirth.Height   := 25;
      QrrDanche.Height  := 25;
      QrrMatter.Height  := 25;
      QrrHangmok.Height := 25;

      QRShape18.Height := 25;
      QRShape19.Height := 25;
      QRShape20.Height := 25;
      QRShape21.Height := 25;
      QRShape22.Height := 25;
      QRShape23.Height := 25;
      QRShape24.Height := 25;
      QRShape25.Height := 25;
      QRShape26.Height := 25;
   end
   else
   begin
      QRBand2.Height    := (iCnt * 18) + 5;

      QrrName.Height    := (iCnt * 18) + 5;
      QrrBirth.Height   := (iCnt * 18) + 5;
      QrrDanche.Height  := (iCnt * 18) + 5;
      QrrMatter.Height  := (iCnt * 18) + 5;
      QrrHangmok.Height := (iCnt * 18) + 5;

      QRShape18.Height := (iCnt * 18) + 5;
      QRShape19.Height := (iCnt * 18) + 5;
      QRShape20.Height := (iCnt * 18) + 5;
      QRShape21.Height := (iCnt * 18) + 5;
      QRShape22.Height := (iCnt * 18) + 5;
      QRShape23.Height := (iCnt * 18) + 5;
      QRShape24.Height := (iCnt * 18) + 5;
      QRShape25.Height := (iCnt * 18) + 5;
      QRShape26.Height := (iCnt * 18) + 5;
   end;
end;

end.
