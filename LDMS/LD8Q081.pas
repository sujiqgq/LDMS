//==============================================================================
// 프로그램 설명 : 바코드 출력
// 시스템        : LDMS
// 수정일자      : 2012.11.08
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8Q081;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, Db, DBTables, QBarcode;

type
  TfrmLD8Q081 = class(TForm)
    qryGulkwa: TQuery;
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_name: TQRLabel;
    QRL_birth: TQRLabel;
    QRL_uniq: TQRLabel;
    QRL_sample: TQRLabel;
    QRL_bunju: TQRLabel;
    BarCode: TBarCode;
    QRL_01: TQRLabel;
    QRL_05: TQRLabel;
    QRL_07: TQRLabel;
    qryHangmokList: TQuery;
    qryNo_hangmok: TQuery;
    QRL_03: TQRLabel;
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
  frmLD8Q081: TfrmLD8Q081;

implementation

uses ZZFunc, MdmsFunc, LD8Q08;
{$R *.DFM}

procedure TfrmLD8Q081.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
{ 1-분주일자
  2-분주번호
  3-센터코드
  4-성명
  5-검진일자
  6-샘플번호
  7-성별
  8-생년월일
  9-생화학
  10-EIA
  11-혈청학
  12-URINE
}

   with frmLD8Q08.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if (Rows = 0) or (UV_iCount = Rows)  then
      begin
         MoreData := False;
         exit;
      end;

      UV_iCount := UV_iCount + 1;

      BarCode.Text := '';
      QRL_01.Caption  := '';  QRL_05.Caption     := ''; QRL_07.Caption    := ''; QRL_03.Caption    := '';
      QRL_uniq.Caption := ''; QRL_bunju.Caption  := '';
      QRL_name.Caption := ''; QRL_sample.Caption := ''; QRL_birth.Caption := '';

      //이름
      QRL_name.Caption := Cell[4,UV_iCount];
      //생년월일
      QRL_birth.Caption := Cell[7,UV_iCount] + '/' + Cell[8,UV_iCount];

      //바코드구분
      case frmLD8Q08.RGrp_barcode.ItemIndex of
         0 : begin
                BarCode.Text := Cell[5,UV_iCount] + Cell[6,UV_iCount];
                QRL_uniq.Caption := copy(Cell[5,UV_iCount],3,2) + '-' +
                                    copy(Cell[5,UV_iCount],5,2) + '-' +
                                    copy(Cell[5,UV_iCount],7,2) + '/' +
                                    Cell[6,UV_iCount];
                QRL_bunju.Caption  := '[' + Cell[3,UV_iCount] + ']' + Cell[2,UV_iCount];
                QRL_sample.Caption := Cell[1,UV_iCount] + '/' + Cell[2,UV_iCount];
             end;
         1 : begin
                BarCode.Text := Cell[1,UV_iCount] + Cell[2,UV_iCount];
                QRL_uniq.Caption := copy(Cell[1,UV_iCount],1,4) + '-' +
                                    copy(Cell[1,UV_iCount],5,2) + '-' +
                                    copy(Cell[1,UV_iCount],7,2) + '/' +
                                    Cell[2,UV_iCount];
                QRL_bunju.Caption  := '[' + Cell[3,UV_iCount] + ']' + Cell[2,UV_iCount];
                QRL_sample.Caption := Cell[5,UV_iCount] + '/' + Cell[6,UV_iCount];
             end;
      end;

      if Cell[ 9,UV_iCount] = 'O' then QRL_01.Caption := '1';
      if Cell[10,UV_iCount] = 'O' then QRL_05.Caption := '5';
      if Cell[11,UV_iCount] = 'O' then QRL_07.Caption := '7';
      if Cell[12,UV_iCount] = 'O' then QRL_03.Caption := '3';
      
      if UV_iCount <= Rows then MoreData := True
      else                      MoreData := False;
   end;
end;

procedure TfrmLD8Q081.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
