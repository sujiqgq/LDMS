unit LD9Q012;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables, QBarcode;

type
  TfrmLD9Q012 = class(TForm)
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
    UV_HM : String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD9Q012: TfrmLD9Q012;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD9Q01;
{$R *.DFM}

procedure TfrmLD9Q012.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD9Q01.grdMaster do
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

      QRLabel1.Caption := Cell[4,UV_iCount]; //샘플번호
      QRLabel2.Caption := Cell[6,UV_iCount]; //성명
      QRLabel3.Caption := Cell[7,UV_iCount]; //단체명
      QRLabel4.Caption := Cell[8,UV_iCount]; //검진일자
      QRLabel5.Caption := Cell[5,UV_iCount]; //성별/생년월일
      QRLabel6.Caption := UV_HM;

      BarCode.Text := copy(Cell[8,UV_iCount],3,6) + Cell[4,UV_iCount];


      if UV_iCount <= Rows then
      begin
         MoreData := True;
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD9Q012.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   with frmLD9Q01 do
   begin
      //검사코드 추출
      qry_LdmsHm.Active := False;
      qry_LdmsHm.ParamByName('PROGRAM_ID').AsString := 'LD9Q01';
      qry_LdmsHm.ParamByName('GROUP_HM').AsString   := Cmb_Hm.Text;
      qry_LdmsHm.Active := True;
      if qry_LdmsHm.RecordCount > 0 then
      begin
         while not qry_LdmsHm.Eof do
         begin
             UV_HM := qry_LdmsHm.FieldByName('desc_kor').AsString;
             qry_LdmsHm.Next;
         end;
      end;
   end;
   UV_iCount := 0;

end;

end.
