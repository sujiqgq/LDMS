//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : MDMS
// 개발일자      : 2010.05.10
// 개발자        : 유동구
//==============================================================================
unit LD5Q042;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables, QBarcode;

type
  TfrmLD5Q042 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel1: TQRLabel;
    BarCode: TBarCode;
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
  frmLD5Q042: TfrmLD5Q042;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q04;
{$R *.DFM}

procedure TfrmLD5Q042.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD5Q04.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if (Rows = 0) or (UV_iCount = Rows) then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      BarCode.Text := '';
      QRLabel1.Caption := ''; QRLabel2.Caption := ''; QRLabel3.Caption := '';
      QRLabel4.Caption := ''; QRLabel5.Caption := ''; QRLabel6.Caption := '';

      QRLabel1.Caption := Cell[4,UV_iCount];
      QRLabel2.Caption := Cell[6,UV_iCount];
      QRLabel3.Caption := Cell[7,UV_iCount];
      QRLabel4.Caption := Cell[8,UV_iCount];

      BarCode.Text := copy(Cell[8,UV_iCount],3,6) + Cell[4,UV_iCount];

      case StrToInt(copy(Cell[5,UV_iCount],7,1)) of
         1,3,5,7 : QRLabel5.Caption := 'M/' + copy(Cell[5,UV_iCount],1,6);
         2,4,6,8 : QRLabel5.Caption := 'F/' + copy(Cell[5,UV_iCount],1,6);
      end;

      QRLabel6.Caption := '자체검진';

      if UV_iCount <= Rows then
      begin
         MoreData := True;
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD5Q042.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
