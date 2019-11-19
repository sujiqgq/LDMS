//==============================================================================
// 프로그램 설명 : 2015년연세대혈액대장(금연,호르몬)
// 시스템        : 통합검진
// 수정일자      : 2015.04.06
// 수정자        : 김창규
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD7P031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables;

type
  TfrmLD7P031 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    qryGulkwa: TQuery;
    QRBand2: TQRBand;
    QRBand3: TQRBand;
    QRShape1: TQRShape;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRLabel9: TQRLabel;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRLabel8: TQRLabel;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText8: TQRDBText;
    QRLabel10: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape21: TQRShape;
    QRShape24: TQRShape;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRL_Gubn: TQRLabel;
    QRSysData3: TQRSysData;
    QRSysData4: TQRSysData;
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    UV_iCount, UV_iTemp : integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD7P031: TfrmLD7P031;

implementation

uses ZZFunc, MdmsFunc;
{$R *.DFM}

procedure TfrmLD7P031.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
   with qryGulkwa do
   begin
      if pos('DI05',FieldByName('cod_chuga').AsString) > 0 then
         QRL_Gubn.Caption := '금연'
      else if pos('DI04',FieldByName('cod_chuga').AsString) > 0 then
         QRL_Gubn.Caption := '환경'
      else QRL_Gubn.Caption := '';
   end;
end;

end.
