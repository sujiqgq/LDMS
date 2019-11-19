//==============================================================================
// 프로그램 설명 : 센터 생화학 결과 출력
// 수정일자      : 2015.09.17
// 수정자        : 곽윤설
// 수정내용      : 신규 개발
// 참고사항      : 한의 부산진단검사의학팀1500839
//==============================================================================
unit LD1P101;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls, Db, DBTables, jpeg, StdCtrls;

type
  TfrmLD1P101 = class(TForm)
    QuickRep: TQuickRep;
    qryJisa: TQuery;
    QRBand1: TQRBand;
    qrl_date: TQRLabel;
    qrl_name: TQRLabel;
    qrl_birth: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel3: TQRLabel;
    QRL_jungdo1: TQRLabel;
    QRLabel9: TQRLabel;
    QRL_jungdo2: TQRLabel;
    QRLabel11: TQRLabel;
    Shape14: TShape;
    Shape15: TShape;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel34: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_jungdo8: TQRLabel;
    QRLabel24: TQRLabel;
    QRL_C009: TQRLabel;
    QRL_C010: TQRLabel;
    QRL_C026: TQRLabel;
    QRL_C011: TQRLabel;
    QRL_C025: TQRLabel;
    QRL_C032: TQRLabel;
    QRL_C027: TQRLabel;
    QRL_C028: TQRLabel;
    QRL_C042: TQRLabel;
    QRL_jisa: TQRLabel;
    QRShape1: TQRShape;
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
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRLabel2: TQRLabel;
    QRL_gmsa: TQRLabel;
    QRLabel10: TQRLabel;
    QRShape27: TQRShape;
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }
    procedure UP_FieldClear(frm : TWinControl);

  public
    { Public declarations }
  end;

var
  frmLD1P101: TfrmLD1P101;

implementation

uses ZZComm, ZZfunc, MdmsFunc, LD1I10;

{$R *.DFM}

procedure TfrmLD1P101.UP_FieldClear(frm : TWinControl);
var i : Integer;
begin
    QRL_date.Caption    := '';
    QRL_name.Caption    := '';
    QRL_birth.Caption   := '';
    QRL_samp.Caption    := '';
    QRL_jungdo1.Caption := '';
    QRL_jungdo2.Caption := '';
    QRL_jungdo8.Caption := '';
    QRL_C009.Caption    := '';
    QRL_C010.Caption    := '';
    QRL_C011.Caption    := '';
    QRL_C025.Caption    := '';
    QRL_C026.Caption    := '';
    QRL_C027.Caption    := '';
    QRL_C028.Caption    := '';
    QRL_C032.Caption    := '';
    QRL_C042.Caption    := '';
    qrl_jisa.Caption    := '';
end;

procedure TfrmLD1P101.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
   UP_FieldClear(QRBand1);

   QRL_date.Caption := Copy(frmLD1I10.mskDAT_GMGN.Text,1,4) + '-' + Copy(frmLD1I10.mskDAT_GMGN.Text,5,2) + '-' + Copy(frmLD1I10.mskDAT_GMGN.Text,7,2);
   QRL_gmsa.Caption := QRL_date.Caption;
   QRL_name.Caption := frmLD1I10.edtDESC_NAME.Text;
   QRL_birth.Caption := Copy(frmLD1I10.mskNUM_JUMIN.Text,1,6) + '-' + Copy(frmLD1I10.mskNUM_JUMIN.Text,7,1) + '******';
   QRL_samp.Caption := frmLD1I10.edt_smp.Text;
   if frmLD1I10.ckGUBN_JUNGDO1.Checked then
   begin
        QRL_jungdo1.Caption := 'Y';
        QRL_jungdo1.Caption := QRL_jungdo1.Caption + ' / ' + Copy(frmLD1I10.mskDAT_SOLVE1.Text,1,4) + '-' +
                               Copy(frmLD1I10.mskDAT_SOLVE1.Text,5,2) + '-' + Copy(frmLD1I10.mskDAT_SOLVE1.Text,7,2);
   end
   else QRL_jungdo1.Caption := 'N';
   if frmLD1I10.ckGUBN_JUNGDO2.Checked then
   begin
        QRL_jungdo2.Caption := 'Y';
        QRL_jungdo2.Caption := QRL_jungdo2.Caption + ' / ' + Copy(frmLD1I10.mskDAT_SOLVE2.Text,1,4) + '-' +
                               Copy(frmLD1I10.mskDAT_SOLVE2.Text,5,2) + '-' + Copy(frmLD1I10.mskDAT_SOLVE2.Text,7,2);
   end
   else QRL_jungdo2.Caption := 'N';
   if frmLD1I10.ckGUBN_JUNGDO8.Checked then
   begin
        QRL_jungdo8.Caption := 'Y';
        QRL_jungdo8.Caption := QRL_jungdo8.Caption + ' / ' + Copy(frmLD1I10.mskDAT_SOLVE8.Text,1,4) + '-' +
                               Copy(frmLD1I10.mskDAT_SOLVE8.Text,5,2) + '-' + Copy(frmLD1I10.mskDAT_SOLVE8.Text,7,2);
   end
   else QRL_jungdo8.Caption := 'N';
   if Trim(frmLD1I10.edtValue1.Text) <> '' then
     QRL_C009.Caption := frmLD1I10.edtValue1.Text + ' IU/L';
   if Trim(frmLD1I10.edtValue2.Text) <> '' then
     QRL_C010.Caption := frmLD1I10.edtValue2.Text + ' IU/L';
   if Trim(frmLD1I10.edtValue3.Text) <> '' then
     QRL_C011.Caption := frmLD1I10.edtValue3.Text + ' IU/L';
   if Trim(frmLD1I10.edtValue4.Text) <> '' then
     QRL_C025.Caption := frmLD1I10.edtValue4.Text + ' mg/dL';
   if Trim(frmLD1I10.edtValue5.Text) <> '' then
     QRL_C026.Caption := frmLD1I10.edtValue5.Text + ' mg/dL';
   if Trim(frmLD1I10.edtValue6.Text) <> '' then
     QRL_C027.Caption := frmLD1I10.edtValue6.Text + ' mg/dL';
   if Trim(frmLD1I10.edtValue7.Text) <> '' then
     QRL_C028.Caption := frmLD1I10.edtValue7.Text + ' mg/dL';
   if Trim(frmLD1I10.edtValue8.Text) <> '' then
     QRL_C032.Caption := frmLD1I10.edtValue8.Text + ' mg/dL';
   if Trim(frmLD1I10.edtValue9.Text) <> '' then
     QRL_C042.Caption := frmLD1I10.edtValue9.Text + ' mg/dL';

   with qryJisa do
   begin
      Close;
      ParamByName('cod_jisa').AsString := copy(GV_sUserId,1,2);
      FetchAll;
      Open;
      if RecordCount > 0 then
      begin
         if copy(frmLD1I10.cmbJisa.Text,1,2) = '15' then qrl_jisa.Left := 381
         else qrl_jisa.Left := 507;
         qrl_jisa.Caption := '검진기관 : ' + FieldByname('desc_hlbw').AsString;
      end;
      Close;
   end;
end;

end.

