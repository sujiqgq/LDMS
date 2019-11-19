//==============================================================================
// ���α׷� ���� : ��ȭ��, ������, Urine ����� ���  - ��ȭ��
// �ý���        : ���հ���
// ��������      : 2015.11.18
// ������        : ������
// ��������      : �ű԰���...
// �������      : ���� �λ����ܰ˻�������1500975 / �λ� ������ å��
//==============================================================================
unit ld4p723;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;
type
  TfrmLD4P723 = class(TForm)
    QuickRep: TQuickRep;
    qryJisa: TQuery;
    QRBand1: TQRBand;
    qrl_gmsa: TQRLabel;
    qrl_name: TQRLabel;
    qrl_birth: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
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
    QRL_HM1: TQRLabel;
    QRL_HM2: TQRLabel;
    QRL_HM3: TQRLabel;
    QRL_HM4: TQRLabel;
    QRL_HM8: TQRLabel;
    QRL_HM7: TQRLabel;
    QRL_HM6: TQRLabel;
    QRL_HM5: TQRLabel;
    QRL_HM9: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_jungdo8: TQRLabel;
    QRLabel24: TQRLabel;
    QRL_U001: TQRLabel;
    QRL_U002: TQRLabel;
    QRL_U005: TQRLabel;
    QRL_U003: TQRLabel;
    QRL_U004: TQRLabel;
    QRL_U008: TQRLabel;
    QRL_U006: TQRLabel;
    QRL_U007: TQRLabel;
    QRL_U009: TQRLabel;
    QRL_jisa: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
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
    QRShape27: TQRShape;
    QRLabel8: TQRLabel;
    QRL_date: TQRLabel;
    qry_GumBul: TQuery;
    QRShape3: TQRShape;
    QRShape6: TQRShape;
    QRL_U012: TQRLabel;
    QRL_U010: TQRLabel;
    QRL_U011: TQRLabel;
    QRL_U013: TQRLabel;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRL_HM13: TQRLabel;
    QRL_HM12: TQRLabel;
    QRL_HM11: TQRLabel;
    QRL_HM10: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure UP_FieldClear;
  private
    { Private declarations }
    UV_iCount, iNo : Integer;
  public
    { Public declarations }
  end;

var
  frmLD4P723: TfrmLD4P723;

implementation

uses ZZComm, ZZfunc, MdmsFunc, LD4Q72;

{$R *.DFM}

procedure TfrmLD4P723.UP_FieldClear;
var i : Integer;
begin
    QRL_date.Caption    := '';
    QRL_gmsa.Caption    := '';
    QRL_name.Caption    := '';
    QRL_birth.Caption   := '';
    QRL_samp.Caption    := '';
    QRL_jungdo8.Caption := 'N';
    QRL_U001.Caption    := '';
    QRL_U002.Caption    := '';
    QRL_U003.Caption    := '';
    QRL_U004.Caption    := '';
    QRL_U005.Caption    := '';
    QRL_U006.Caption    := '';
    QRL_U007.Caption    := '';
    QRL_U008.Caption    := '';
    QRL_U009.Caption    := '';
    QRL_U010.Caption    := '';
    QRL_U011.Caption    := '';
    QRL_U012.Caption    := '';
    QRL_U013.Caption    := '';
    qrl_jisa.Caption    := '';
end;

procedure TfrmLD4P723.QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
var temp,tempa :String;
   i, j : integer;
begin
   with frmLD4Q72 do
   begin
      // �ڷᰡ ��������� ó��
      if (grdMaster.Rows = 0) or (grdMaster.Rows = UV_iCount) then
      begin
         MoreData := False;
         exit;
      end;

      inc(UV_iCount);
      UP_FieldClear;

      QRL_gmsa.Caption := Copy(grdMaster.Cell[6, UV_iCount],1,4) + '-' + Copy(grdMaster.Cell[6, UV_iCount],5,2) + '-' + Copy(grdMaster.Cell[6, UV_iCount],7,2);
      QRL_date.Caption := QRL_gmsa.Caption;
      QRL_name.Caption := grdMaster.Cell[4, UV_iCount];
      QRL_birth.Caption := Copy(grdMaster.Cell[3, UV_iCount],1,8) + '******';
      QRL_samp.Caption := grdMaster.Cell[2, UV_iCount];

      if Trim(UV_vGumbul[UV_iCount-1]) <> '' then
      begin
         qry_GumBul.Close;
         qry_GumBul.ParamByName('Cod_jisa').AsString := Copy(cmbJisa.Text,1,2);
         qry_GumBul.ParamByName('num_id').AsString   := UV_vNum_id[UV_iCount-1];
         qry_GumBul.ParamByName('dat_gmgn').AsString := grdMaster.Cell[6, UV_iCount];
         qry_GumBul.Open;
         if qry_GumBul.RecordCount > 0 then
         begin
            while not qry_GumBul.Eof do
            begin
              if (qry_GumBul.FieldByName('gubn_bul').AsString = '08') then
              begin
                 QRL_jungdo8.Caption := 'Y';
                 QRL_jungdo8.Caption := QRL_jungdo8.Caption + ' / ' + Copy(qry_GumBul.FieldByName('dat_solve').AsString,1,4) + '-' +
                                        Copy(qry_GumBul.FieldByName('dat_solve').AsString,5,2) + '-' + Copy(qry_GumBul.FieldByName('dat_solve').AsString,7,2);
              end;
              qry_GumBul.Next;
            end;
         end;
         qry_GumBul.Close;
      end;

      for i := 0 to Round(length(vArr_HM)-1) do
      begin
         for j := 0 to QRBand1.ControlCount - 1 do
         begin
            if 'QRL_HM' + IntToStr(i+1) = TQRLabel(QRBand1.Controls[j]).Name then
               TQRLabel(QRBand1.Controls[j]).Caption := vArr_HM[i][1];
            if ('QRL_' + vArr_HM[i][0] = TQRLabel(QRBand1.Controls[j]).Name) and
               (pos(vArr_HM[i][0], UV_vCod_HM[UV_iCount-1]) > 0) and
               (Trim(Copy(UV_vGulkwa[UV_iCount-1], StrToInt(frmLD4Q72.vArr_HM[i][2]), 6)) <> '') then
               TQRLabel(QRBand1.Controls[j]).Caption := Trim(Copy(UV_vGulkwa[UV_iCount-1], StrToInt(frmLD4Q72.vArr_HM[i][2]), 6)) + ' ' + vArr_HM[i][3];
         end;
      end;

      with qryJisa do
      begin
         Close;
         ParamByName('cod_jisa').AsString := copy(GV_sUserId,1,2);
         FetchAll;
         Open;
         if RecordCount > 0 then
         begin
            if copy(GV_sUserId,1,2) = '15' then qrl_jisa.Left := 381
            else qrl_jisa.Left := 507;
            qrl_jisa.Caption := '������� : ' + FieldByname('desc_hlbw').AsString;
         end;
         Close;
      end;

      if UV_iCount <= grdMaster.Rows then
         MoreData := True
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4P723.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   iNo := 0;
   UV_iCount := 0;
end;

end.

