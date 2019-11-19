//==============================================================================
// 프로그램 설명 : 세포검사 결과출력 form
// 수정일자      : 2019.03.21
// 수정자        : 박수지
// 수정내용      :
// 참고사항      :
//==============================================================================
unit ld1p132;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls, Db, DBTables, jpeg;

type
  TfrmLD1P132 = class(TForm)
    QuickRep: TQuickRep;
    qryGumgin: TQuery;
    QRBand1: TQRBand;
    qrl_Doctor: TQRLabel;
    qrl_sex1: TQRLabel;
    qrl_Result: TQRLabel;
    qrl_age: TQRLabel;
    desc_buwi: TQRLabel;
    qrl_jisa: TQRLabel;
    qrl_name: TQRLabel;
    qrl_birth: TQRLabel;
    qrl_date: TQRLabel;
    qryCell: TQuery;
    qrl_caption: TQRLabel;
    Img_JSY: TQRImage;
    QRLabel1: TQRLabel;
    Img_PJK: TQRImage;
    Img_LSN: TQRImage;
    QRLabel3: TQRLabel;
    QRL_Damdang: TQRLabel;
    qrl_detect_date: TQRLabel;
    QRL_datgmgn: TQRLabel;
    QRL_Addr: TQRLabel;
    QRL_Tel: TQRLabel;
    qrl_Sokun: TQRMemo;
    qryJangbi: TQuery;
    qryCa_Request: TQuery;
    qrl_date2: TQRLabel;
    qrl_date3: TQRLabel;
    qrl_detect_date2: TQRLabel;
    qrl_detect_date3: TQRLabel;
    QRL_datgmgn2: TQRLabel;
    QRL_datgmgn3: TQRLabel;
    qrl_Result2: TQRLabel;
    qrl_Result3: TQRLabel;
    qrl_birth2: TQRLabel;
    qrl_sex2: TQRLabel;
    qrySawon_sign: TQuery;
    QRL_KiKwan: TQRLabel;
    qrl_dat_time: TQRLabel;
    qrl_dat_time_R: TQRLabel;
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }
    procedure UP_FieldClear(frm : TWinControl);

  public
    { Public declarations }
  end;

var
  frmLD1P132: TfrmLD1P132;

implementation

uses LD1P13, ZZComm, ZZfunc, MdmsFunc;

{$R *.DFM}

procedure TfrmLD1P132.UP_FieldClear(frm : TWinControl);
var i : Integer;
begin
   // 종속된 Control를 찾아서 해당 Property를 Clear
  for i := 0 to frm.ControlCount - 1 do
   begin
      if frm.Controls[i].ClassType = TQRLabel then
         TQRLabel(frm.Controls[i]).Caption := ''
      else if frm.Controls[i].ClassType = TQRMemo then
         TQRMemo(frm.Controls[i]).Lines.Clear;
   end;
end;

procedure TfrmLD1P132.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var iAge, count : Integer;
    sMan : string;
    //이미지 관련.
   Jpeg  : TJPEGImage;
   MS    : TMemoryStream;
begin
   UP_FieldClear(QRBand1);
   count := 0;

   with qryGumgin do
   begin
      GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);

      //성명
      if frmLD1P13.Ck_Desc_dc.Checked then
      begin
         qrl_name.Caption  := FieldByname('desc_name').AsString + ' (' + FieldByname('desc_dc').AsString + ')';
         if FieldByname('gubn_neawon').AsString = '2' then
            qrl_name.Caption  := qrl_name.Caption + '-출장'
         else
            qrl_name.Caption  := qrl_name.Caption + '-내원';
      end
      else
         qrl_name.Caption  := FieldByname('desc_name').AsString;
      //주민번호
      qrl_birth.Caption  := copy(FieldByname('num_jumin').AsString,1,6);
      qrl_birth2.Caption := copy(FieldByname('num_jumin').AsString,7,1) + '******';
      //담당의
      if frmLD1P13.cmbDoctor.Text <> '' then
      begin
         QRLabel3.Caption    := '담  당  의 :';
         QRL_Damdang.Caption := frmLD1P13.cmbDoctor.Text;
      end
      else
         QRL_Damdang.Caption := '';

      //검진일자
      qrl_date.Caption   := copy(FieldByname('dat_gmgn').AsString,4,1);
      qrl_date2.Caption  := copy(FieldByname('dat_gmgn').AsString,5,2);
      qrl_date3.Caption  := copy(FieldByname('dat_gmgn').AsString,7,2);

      //검체채취일
      qrl_detect_date.Caption  := copy(FieldByname('dat_gmgn').AsString,4,1);
      qrl_detect_date2.Caption := copy(FieldByname('dat_gmgn').AsString,5,2);
      qrl_detect_date3.Caption := copy(FieldByname('dat_gmgn').AsString,7,2);

      qryCell.Active := False;
      qryCell.ParamByName('cod_jisa').AsString   := FieldByname('cod_jisa').AsString;
      qryCell.ParamByName('num_id').AsString     := FieldByname('num_id').AsString;
      qryCell.ParamByName('dat_gmgn').AsString   := FieldByname('dat_gmgn').AsString;
      qryCell.ParamByName('cod_hm').AsString     := FieldByname('cod_hm').AsString;

      qryCell.Active := True;
      if qryCell.IsEmpty = False then
      begin
         while not qryCell.Eof do
         begin
               //검진센터
               qrl_jisa.Caption := '[' + qryCell.FieldByname('cod_jisa').AsString + ']' +
                                    qryCell.FieldByname('desc_jisa').AsString;
               //검체번호
               desc_buwi.Caption   := qryCell.FieldByName('desc_buwi').AsString;
               //나이
               qrl_age.Caption    := IntToStr(iAge);
               //성별
               qrl_sex1.Caption := '';
               qrl_sex2.Caption := '';

               if sMan = 'M' then qrl_sex1.Caption := 'v'
               else if sMan = 'F' then qrl_sex2.Caption := 'v';

               desc_buwi.Caption   := qryCell.FieldByName('desc_buwi').AsString;

               //검체접수일
               QRL_datgmgn.Caption  := copy(qryCell.FieldByName('dat_last').AsString,4,1);
               QRL_datgmgn2.Caption := copy(qryCell.FieldByName('dat_last').AsString,5,2);
               QRL_datgmgn3.Caption := copy(qryCell.FieldByName('dat_last').AsString,7,2);
               qrl_dat_time.Caption := (qryCell.FieldByName('dat_time').AsString);


               qrl_Result.Caption  :=copy(qryCell.FieldByName('dat_result').AsString,4,1);
               qrl_Result2.Caption :=copy(qryCell.FieldByName('dat_result').AsString,5,2);
               qrl_Result3.Caption :=copy(qryCell.FieldByName('dat_result').AsString,7,2);
               qrl_dat_time_R.Caption := (qryCell.FieldByName('dat_time_R').AsString);

               qrl_Doctor.Caption := 'signature of Pathologist ' + qryCell.FieldByName('DOCTOR').AsString + ' ' +
                                     '( ' + qryCell.FieldByName('DOCTOR_NO').AsString + ' )';

               if (qryCell.FieldByName('cod_hm').AsString= 'P001') or
                  (qryCell.FieldByName('cod_hm').AsString= 'P003') then
                  qrl_caption.Caption := 'Gy Cytology'

               else if (qryCell.FieldByName('cod_hm').AsString= 'B009') then
                  qrl_caption.Caption := 'Gy Cytology(LBC)'

               else if (qryCell.FieldByName('cod_hm').AsString= 'B001') or
                       (qryCell.FieldByName('cod_hm').AsString= 'B007') then
                  qrl_caption.Caption := 'Tissue Pathology'

               else if (qryCell.FieldByName('cod_hm').AsString= 'P010') then
                  qrl_caption.Caption := 'Fine Needle Aspiration (Thyroid)'

               else if (qryCell.FieldByName('cod_hm').AsString= 'P011') then
                  qrl_caption.Caption := 'Fine Needle Aspiration (Breast)'

               else if (qryCell.FieldByName('cod_hm').AsString= 'P006') then
                  qrl_caption.Caption := 'Fine Needle Aspiration'

               else if (qryCell.FieldByName('cod_hm').AsString= 'P004') then
                  qrl_caption.Caption := 'Non-Gy Cytology';

               if (qryCell.FieldByName('cod_hm').AsString= 'B001') or
                  (qryCell.FieldByName('cod_hm').AsString= 'B007') then
               begin
                  with qryJangbi do
                  begin
                     // Filter 제거
                     Filtered := False;
                     qryJangbi.Active := False;
                     qryJangbi.ParamByName('cod_jisa').AsString := qryCell.FieldByname('cod_jisa').AsString;
                     qryJangbi.ParamByName('num_id').AsString   := qryCell.FieldByname('num_id').AsString;
                     qryJangbi.ParamByName('dat_gmgn').AsString := qryCell.FieldByname('dat_gmgn').AsString;
                     qryJangbi.Active := True;
                     Filtered := True;
                  
                     if qryCell.FieldByName('cod_hm').AsString= 'B001' then        //위내시경
                     begin
                        qryJangbi.Filter := ' cod_jangbi = ''JJ34'' ';
                        if qryJangbi.RecordCount > 0 then
                        begin
                           QRLabel3.Caption    := '담  당  의 :';
                           QRL_Damdang.Caption := qryJangbi.FieldByName('DOCTOR').AsString;
                        end
                        else
                        begin
                           qryJangbi.Filter := ' cod_jangbi = ''JJB9'' ';
                           if qryJangbi.RecordCount > 0 then
                           begin
                              QRLabel3.Caption    := '담  당  의 :';
                              QRL_Damdang.Caption := qryJangbi.FieldByName('DOCTOR').AsString;
                           end;
                        end;
                     end
                     else if qryCell.FieldByName('cod_hm').AsString= 'B007' then   //대장내시경
                     begin
                        qryJangbi.Filter := ' cod_jangbi = ''JJ83'' ';
                        if qryJangbi.RecordCount > 0 then
                        begin
                           QRLabel3.Caption    := '담  당  의 :';
                           QRL_Damdang.Caption := qryJangbi.FieldByName('DOCTOR').AsString;
                        end
                        else
                        begin
                           qryJangbi.Filter := ' cod_jangbi = ''JJ86'' ';
                           if qryJangbi.RecordCount > 0 then
                           begin
                              QRLabel3.Caption    := '담  당  의 :';
                              QRL_Damdang.Caption := qryJangbi.FieldByName('DOCTOR').AsString;
                           end;
                        end;
                     end;
                  end;
               end;

               if (qryCell.FieldByName('cod_hm').AsString= 'P001') or
                  (qryCell.FieldByName('cod_hm').AsString= 'P003') then
               begin
                  with qryCa_Request do
                  begin
                     qryCa_Request.Active := False;
                     qryCa_Request.ParamByName('cod_jisa').AsString := qryCell.FieldByname('cod_jisa').AsString;
                     qryCa_Request.ParamByName('num_id').AsString   := qryCell.FieldByname('num_id').AsString;
                     qryCa_Request.ParamByName('dat_gmgn').AsString := qryCell.FieldByname('dat_gmgn').AsString;
                     qryCa_Request.Active := True;
                     if qryCa_Request.RecordCount > 0 then
                     begin
                        QRLabel3.Caption    := '담  당  의 :';
                        QRL_Damdang.Caption := qryCa_Request.FieldByName('DOCTOR').AsString;
                     end;
                  end;
               end;

{               qrl_Sokun.Caption := qryCell.FieldByName('DESC_SOKUN1').AsString +
                                     qryCell.FieldByName('DESC_SOKUN2').AsString +
                                     qryCell.FieldByName('DESC_SOKUN3').AsString +
                                     qryCell.FieldByName('DESC_SOKUN4').AsString +
                                     qryCell.FieldByName('DESC_SOKUN5').AsString;  }
               qrl_Sokun.Lines.Add(qryCell.FieldByName('DESC_SOKUN1').AsString +
                                     qryCell.FieldByName('DESC_SOKUN2').AsString +
                                     qryCell.FieldByName('DESC_SOKUN3').AsString +
                                     qryCell.FieldByName('DESC_SOKUN4').AsString +
                                     qryCell.FieldByName('DESC_SOKUN5').AsString);

               Img_JSY.Enabled := False;
               Img_PJK.Enabled := False;
               Img_LSN.Enabled := False;

               if pos('조수연',qrl_Doctor.Caption) > 0 then
               begin
                  //Sign
                  Img_JSY.Picture := nil;
                  Img_JSY.Enabled := False;

                  qrySawon_sign.Close;
                  qrySawon_sign.ParamByName('cod_sawon').AsString := qryCell.FieldByname('cod_doct').AsString;
                  qrySawon_sign.Open;
                  if (qrySawon_sign.IsEmpty = False) and (Trim(qrySawon_sign.FieldByName('image_sign').AsString) <> '') then
                  begin
                    try
                       Img_JSY.Enabled := True;
                       MS := TMemoryStream.Create;
                       (qrySawon_sign.FieldByName('image_sign') As TBlobField).SaveToStream(MS);
                       MS.Position:=0;
                       try
                          Jpeg := TJPEGImage.Create;
                          Jpeg.LoadFromStream(MS);
                          Img_JSY.Picture.Assign(Jpeg);
                          Jpeg.Free;
                       except  //Bmp
                          Img_JSY.Picture.Assign(qrySawon_sign.FieldByName('image_sign') As TBlobField);
                       end;
                    finally
                       MS.Free;
                    end;
                  end
                  else
                  begin
                    Img_JSY.Picture := nil;
                    Img_JSY.Enabled := False;
                  end;

                  qrySawon_sign.Close;
               end
               else if pos('박진규',qrl_Doctor.Caption) > 0 then
               begin
                  Img_PJK.Enabled := True;
               end
               else if pos('이시내',qrl_Doctor.Caption) > 0 then
               begin
                  Img_LSN.Enabled := True;
               end
               else
               begin
                  Img_JSY.Enabled := False;
                  Img_PJK.Enabled := False;
                  Img_LSN.Enabled := False;
               end;
               qryCell.Next;
         end;
      end;
      QRL_KiKwan.Caption := '검사기관 : 11372958';
      QRL_Addr.Caption := '주소 : 서울시 종로구 당주동 100번지 세종B/D';
      QRL_Tel.Caption  := '전화 : 02-3702-9171';
   end;
end;

end.
