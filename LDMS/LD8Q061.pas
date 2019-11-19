//==============================================================================
// 프로그램 설명 : [WorkList]파트별 검사항목 대장출력
// 시스템        : 통합검진
// 수정일자      : 2007.03.21
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8Q061;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD8Q061 = class(TfrmPrintForm)
    Band1: TQRBand;
    QRLabel1: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel49: TQRLabel;
    QRLabel50: TQRLabel;
    QRLabel51: TQRLabel;
    QRLabel52: TQRLabel;
    QRLabel53: TQRLabel;
    QRLabel54: TQRLabel;
    QRLabel55: TQRLabel;
    QRLabel56: TQRLabel;
    QRLabel57: TQRLabel;
    QRLabel58: TQRLabel;
    QRLabel59: TQRLabel;
    QRLabel60: TQRLabel;
    QRLabel61: TQRLabel;
    QRLabel62: TQRLabel;
    QRLabel63: TQRLabel;
    QRLabel64: TQRLabel;
    QRLabel65: TQRLabel;
    QRLabel66: TQRLabel;
    QRLabel67: TQRLabel;
    QRLabel68: TQRLabel;
    QRLabel69: TQRLabel;
    QRLabel70: TQRLabel;
    QRLabel71: TQRLabel;
    QRLabel72: TQRLabel;
    QRLabel73: TQRLabel;
    QRLabel74: TQRLabel;
    QRLabel75: TQRLabel;
    QRLabel76: TQRLabel;
    QRLabel77: TQRLabel;
    QRLabel78: TQRLabel;
    QRLabel79: TQRLabel;
    QRLabel80: TQRLabel;
    QRLabel81: TQRLabel;
    QRLabel82: TQRLabel;
    QRLabel83: TQRLabel;
    QRLabel84: TQRLabel;
    QRLabel85: TQRLabel;
    QRLabel86: TQRLabel;
    QRLabel87: TQRLabel;
    QRLabel88: TQRLabel;
    QRLabel89: TQRLabel;
    QRLabel90: TQRLabel;
    QRLabel91: TQRLabel;
    QRLabel92: TQRLabel;
    QRLabel93: TQRLabel;
    QRLabel94: TQRLabel;
    QRLabel95: TQRLabel;
    QRLabel96: TQRLabel;
    QRLabel97: TQRLabel;
    QRLabel98: TQRLabel;
    QRLabel99: TQRLabel;
    QRLabel100: TQRLabel;
    QRLabel101: TQRLabel;
    QRLabel102: TQRLabel;
    QRLabel103: TQRLabel;
    QRLabel104: TQRLabel;
    QRLabel105: TQRLabel;
    QRLabel106: TQRLabel;
    QRLabel107: TQRLabel;
    QRLabel108: TQRLabel;
    QRLabel109: TQRLabel;
    QRLabel110: TQRLabel;
    QRLabel111: TQRLabel;
    QRLabel112: TQRLabel;
    QRLabel113: TQRLabel;
    QRLabel114: TQRLabel;
    QRLabel115: TQRLabel;
    QRLabel116: TQRLabel;
    QRLabel117: TQRLabel;
    QRLabel118: TQRLabel;
    QRLabel119: TQRLabel;
    QRLabel120: TQRLabel;
    QRLabel121: TQRLabel;
    QRLabel122: TQRLabel;
    QRLabel123: TQRLabel;
    QRLabel124: TQRLabel;
    QRLabel125: TQRLabel;
    QRLabel126: TQRLabel;
    QRLabel127: TQRLabel;
    QRLabel128: TQRLabel;
    QRLabel129: TQRLabel;
    QRLabel130: TQRLabel;
    QRLabel131: TQRLabel;
    QRLabel132: TQRLabel;
    QRLabel133: TQRLabel;
    QRLabel134: TQRLabel;
    QRLabel135: TQRLabel;
    QRLabel136: TQRLabel;
    QRLabel137: TQRLabel;
    QRLabel138: TQRLabel;
    QRLabel139: TQRLabel;
    QRLabel140: TQRLabel;
    QRLabel141: TQRLabel;
    QRLabel142: TQRLabel;
    QRLabel143: TQRLabel;
    QRLabel144: TQRLabel;
    QRLabel145: TQRLabel;
    QRLabel146: TQRLabel;
    QRLabel147: TQRLabel;
    QRLabel148: TQRLabel;
    QRLabel149: TQRLabel;
    QRLabel150: TQRLabel;
    QRLabel151: TQRLabel;
    QRLabel152: TQRLabel;
    QRLabel153: TQRLabel;
    QRLabel154: TQRLabel;
    QRLabel155: TQRLabel;
    QRLabel156: TQRLabel;
    QRLabel157: TQRLabel;
    QRLabel158: TQRLabel;
    QRLabel159: TQRLabel;
    QRLabel160: TQRLabel;
    QRLabel161: TQRLabel;
    QRLabel162: TQRLabel;
    QRLabel163: TQRLabel;
    QRLabel164: TQRLabel;
    QRLabel165: TQRLabel;
    QRLabel166: TQRLabel;
    QRLabel167: TQRLabel;
    QRLabel168: TQRLabel;
    QRLabel169: TQRLabel;
    QRLabel170: TQRLabel;
    QRLabel171: TQRLabel;
    QRLabel172: TQRLabel;
    QRLabel173: TQRLabel;
    QRLabel174: TQRLabel;
    QRLabel175: TQRLabel;
    QRLabel176: TQRLabel;
    QRLabel177: TQRLabel;
    QRLabel178: TQRLabel;
    QRLabel193: TQRLabel;
    QRLabel194: TQRLabel;
    QRLabel196: TQRLabel;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape25: TQRShape;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD8Q061: TfrmLD8Q061;

implementation

{$R *.DFM}

uses ZZfunc, LD8Q06;

procedure TfrmLD8Q061.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var i : Integer;
begin
   inherited;

   with frmLD8Q06.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      for i := 58 to 67 do
      begin

           if TQRLabel(Band1.Controls[i]).Name = 'QRLabel' + FloatToStr(i-19) then
            TQRLabel(Band1.Controls[i]).Caption := Cell[1,  UV_iCount];
           if TQRLabel(Band1.Controls[i+10]).Name = 'QRLabel' + FloatToStr(i-9) then
            TQRLabel(Band1.Controls[i+10]).Caption := Cell[2,  UV_iCount];
           if TQRLabel(Band1.Controls[i+20]).Name = 'QRLabel' + FloatToStr(i+1) then
            TQRLabel(Band1.Controls[i+20]).Caption := Cell[3,  UV_iCount];
           if TQRLabel(Band1.Controls[i+30]).Name = 'QRLabel' + FloatToStr(i+11) then
            TQRLabel(Band1.Controls[i+30]).Caption := Cell[4,  UV_iCount];
           if TQRLabel(Band1.Controls[i+40]).Name = 'QRLabel' + FloatToStr(i+21) then
            TQRLabel(Band1.Controls[i+40]).Caption := Cell[5,  UV_iCount];
           if TQRLabel(Band1.Controls[i+50]).Name = 'QRLabel' + FloatToStr(i+31) then
            TQRLabel(Band1.Controls[i+50]).Caption := Cell[6,  UV_iCount];
           if TQRLabel(Band1.Controls[i+60]).Name = 'QRLabel' + FloatToStr(i+41) then
            TQRLabel(Band1.Controls[i+60]).Caption := Cell[7,  UV_iCount];
           if TQRLabel(Band1.Controls[i+70]).Name = 'QRLabel' + FloatToStr(i+51) then
            TQRLabel(Band1.Controls[i+70]).Caption := Cell[8,  UV_iCount];
           if TQRLabel(Band1.Controls[i+80]).Name = 'QRLabel' + FloatToStr(i+61) then
            TQRLabel(Band1.Controls[i+80]).Caption := Cell[9,  UV_iCount];
           if TQRLabel(Band1.Controls[i+90]).Name = 'QRLabel' + FloatToStr(i+71) then
            TQRLabel(Band1.Controls[i+90]).Caption := Cell[10,  UV_iCount];
           if TQRLabel(Band1.Controls[i+100]).Name = 'QRLabel' + FloatToStr(i+81) then
            TQRLabel(Band1.Controls[i+100]).Caption := Cell[11,  UV_iCount];
           if TQRLabel(Band1.Controls[i+110]).Name = 'QRLabel' + FloatToStr(i+91) then
            TQRLabel(Band1.Controls[i+110]).Caption := Cell[12,  UV_iCount];
           if TQRLabel(Band1.Controls[i+120]).Name = 'QRLabel' + FloatToStr(i+101) then
            TQRLabel(Band1.Controls[i+120]).Caption := Cell[13,  UV_iCount];
           if TQRLabel(Band1.Controls[i+130]).Name = 'QRLabel' + FloatToStr(i+111) then
            TQRLabel(Band1.Controls[i+130]).Caption := Cell[14,  UV_iCount];

        Inc(UV_iCount);
      end;
   end;
end;

procedure TfrmLD8Q061.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD8Q06 do
   begin
      QRLabel3.Caption  := GF_DateFormat(mskTo.Text);
      QRLabel5.Caption  := GF_DateFormat(mskFrom.Text);
      QRLabel7.Caption := copy(cmbJisa.Text,4,100);
   end;
end;

end.
