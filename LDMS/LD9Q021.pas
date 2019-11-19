unit LD9Q021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD9Q021 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
    QRShape2: TQRShape;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand3: TQRBand;
    QRLabel9: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel24: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRL_RACKNO: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_BIRTH: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);

  private
    { Private declarations }
    UV_iCount, i, imax : Integer;
    qlArray : array[1..32] of TQRLabel;
    qlArray_HM : array[1..32] of TQRLabel;
  public
    { Public declarations }

  end;

var
  frmLD9Q021: TfrmLD9Q021;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD9Q02;

{$R *.DFM}

procedure TfrmLD9Q021.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i : Integer;
begin
   inherited;
   with frmLD9Q02.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      Inc(UV_iCount);
      
      QRL_NO.caption     := Cell[1, UV_iCount]; //No.
      QRL_RACKNO.Caption := Cell[2, UV_iCount]; //Rack No.
      QRL_SAMPLE.Caption := Cell[3, UV_iCount]; //샘플번호
      QRL_BUNJU.Caption  := Cell[4, UV_iCount]; //분주번호
      QRL_BIRTH.Caption  := Cell[5, UV_iCount]; //생년월일/성별
      QRL_NAME.Caption   := Cell[6, UV_iCount]; //성명

      for i := 1 to imax do
      begin
          qlArray[i].Caption := Cell[i+6, UV_iCount];
      end;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD9Q021.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;

   QRL_date.Caption := copy(frmLD9Q02.mskDate.Text,1,4) + '년'  +
                       copy(frmLD9Q02.mskDate.Text,5,2) + '월'  +
                       copy(frmLD9Q02.mskDate.Text,7,2) + '일';
   QRLabel1.Caption := QRLabel1.Caption + '(' + frmLD9Q02.Cmb_Hm.Text + ')';

   i := 1;
   frmLD9Q02.qry_LdmsHm.Active := False;
   frmLD9Q02.qry_LdmsHm.ParamByName('PROGRAM_ID').AsString := 'LD9Q02';
   frmLD9Q02.qry_LdmsHm.ParamByName('GROUP_HM').AsString   := frmLD9Q02.Cmb_Hm.Text;
   frmLD9Q02.qry_LdmsHm.Active := True;
   if frmLD9Q02.qry_LdmsHm.RecordCount > 0 then
   begin
        imax := frmLD9Q02.qry_LdmsHm.RecordCount;
        while not frmLD9Q02.qry_LdmsHm.Eof do
        begin
            qlArray_HM[i] := TQRLabel.Create(Self);
            qlArray_HM[i].Parent    := QRBand3;
            qlArray_HM[i].Caption   := frmLD9Q02.qry_LdmsHm.FieldByName('cod_hm').AsString;
            qlArray_HM[i].Alignment := taCenter;
            qlArray_HM[i].Left      := 230 + (i*50);
            qlArray_HM[i].Top       := 3;
            qlArray_HM[i].Height    := 13;
            qlArray_HM[i].Font.color:= clBlack;
            qlArray_HM[i].Font.Size := 9;

            qlArray[i] := TQRLabel.Create(Self);
            qlArray[i].Parent     := QRBand2;
            qlArray[i].Alignment  := taCenter;
            qlArray[i].Left       := 227 + (i*50);
            qlArray[i].Top        := 2;
            qlArray[i].Height     := 13;
            qlArray[i].Width      := 35;
            qlArray[i].Font.color := clBlack;
            qlArray[i].Font.Size  := 9;

            inc(i);
            frmLD9Q02.qry_LdmsHm.Next;
        end;
   end;
   frmLD9Q02.qry_LdmsHm.Active := False;
   UV_iCount := 0;
end;

end.
