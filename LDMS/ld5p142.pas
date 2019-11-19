//==============================================================================
// 프로그램 설명 : 건강검진 검체검사 결과지
// 시스템        : LDMS
// 수정일자      : 2016.4.15
// 수정자        : 박대용
//==============================================================================
unit LD5P142;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, jpeg, QuickRpt, ExtCtrls, Db, DBTables, StdCtrls, math;

  const GawngJu_Code = 'SE42SEA6SE92SE93SE96';

type
  TfrmLD5P142 = class(TForm)
    qryGumgin1: TQuery;
    qryHm: TQuery;
    QuickRep: TQuickRep;
    qrySawon_sign: TQuery;
    QRBand1: TQRBand;
    QRImage1: TQRImage;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRShape1: TQRShape;
    QRShape5: TQRShape;
    QRShape2: TQRShape;
    QRShape6: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape7: TQRShape;
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
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRLabel1: TQRLabel;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape30: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape44: TQRShape;
    QRShape45: TQRShape;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRShape48: TQRShape;
    QRShape49: TQRShape;
    QRShape50: TQRShape;
    QRShape51: TQRShape;
    QRShape52: TQRShape;
    QRShape53: TQRShape;
    QRShape54: TQRShape;
    QRShape55: TQRShape;
    QRShape57: TQRShape;
    QRShape58: TQRShape;
    QRShape59: TQRShape;
    QRShape60: TQRShape;
    QRShape61: TQRShape;
    QRShape27: TQRShape;
    QRShape62: TQRShape;
    QRShape63: TQRShape;
    QRShape64: TQRShape;
    QRShape65: TQRShape;
    QRShape66: TQRShape;
    QRShape67: TQRShape;
    QRShape69: TQRShape;
    QRShape70: TQRShape;
    QRShape21: TQRShape;
    QRShape68: TQRShape;
    QRShape56: TQRShape;
    QRLabel5: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
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
    QRShape74: TQRShape;
    QRShape75: TQRShape;
    QRShape76: TQRShape;
    QRShape77: TQRShape;
    QRShape78: TQRShape;
    QRShape79: TQRShape;
    QRShape80: TQRShape;
    QRShape81: TQRShape;
    QRShape82: TQRShape;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel49: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel45: TQRLabel;
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
    QRShape83: TQRShape;
    QRShape84: TQRShape;
    QRLabel78: TQRLabel;
    QRLabel79: TQRLabel;
    QRLabel80: TQRLabel;
    QRLabel81: TQRLabel;
    QRLabel82: TQRLabel;
    QRLabel86: TQRLabel;
    QRLabel87: TQRLabel;
    QRLabel88: TQRLabel;
    QRLabel92: TQRLabel;
    QRLabel94: TQRLabel;
    QRLabel95: TQRLabel;
    QRLabel97: TQRLabel;
    QRLabel99: TQRLabel;
    QRLabel100: TQRLabel;
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
    QRShape85: TQRShape;
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
    QRLabel133: TQRLabel;
    QRLabel134: TQRLabel;
    QRLabel135: TQRLabel;
    QRLabel136: TQRLabel;
    QRLabel137: TQRLabel;
    QRLabel130: TQRLabel;
    QRLabel131: TQRLabel;
    QRLabel132: TQRLabel;
    QRLabel138: TQRLabel;
    QRLabel140: TQRLabel;
    QRLabel144: TQRLabel;
    QRLabel139: TQRLabel;
    qryGumgin: TQuery;
    qryNhic: TQuery;
    QRL_Name: TQRLabel;
    QRL_Jumin: TQRLabel;
    QRL_Numbohm: TQRLabel;
    QRLabel148: TQRLabel;
    QRL_Dat_gmgn: TQRLabel;
    QRL_Dat_bunju: TQRLabel;
    qryGulkwa: TQuery;
    QRL_H002: TQRLabel;
    QRL_C032: TQRLabel;
    QRL_C025: TQRLabel;
    QRL_C026: TQRLabel;
    QRL_C028: TQRLabel;
    QRL_C009: TQRLabel;
    QRL_C010: TQRLabel;
    QRL_C011: TQRLabel;
    QRL_C042: TQRLabel;
    QRL_S007: TQRLabel;
    QRL_S099: TQRLabel;
    QRL_S091: TQRLabel;
    QRL_S016: TQRLabel;
    QRL_Y004: TQRLabel;
    QRL_TT03: TQRLabel;
    QRL_TT01: TQRLabel;
    QRL_TT02: TQRLabel;
    qryHgulkwa_chk: TQuery;
    qryProfile: TQuery;
    QRL_NumSamp1: TQRLabel;
    QRL_NumSamp2: TQRLabel;
    QRL_Telnum: TQRLabel;
    QRL_doctor: TQRLabel;
    QRL_Kiho: TQRLabel;
    QRL_Kikwan: TQRLabel;
    QRL_Dat_Tobo: TQRLabel;
    QRL_Dat_gum: TQRLabel;
    QRLabel149: TQRLabel;
    QRL_GFR: TQRLabel;
    QRL_C027: TQRLabel;
    QRL_C074: TQRLabel;
    QRLabel90: TQRLabel;
    QRLabel91: TQRLabel;
    QRLabel93: TQRLabel;
    qryGulkwa_hm: TQuery;
    QRLabel96: TQRLabel;
    QRLabel98: TQRLabel;
    QRLabel145: TQRLabel;
    QRLabel146: TQRLabel;
    QRLabel150: TQRLabel;
    QRLabel151: TQRLabel;
    QRL_S008: TQRLabel;
    QRLabel101: TQRLabel;
    QRLabel141: TQRLabel;
    QRL_S033: TQRLabel;
    QRLabel142: TQRLabel;
    QRShape71: TQRShape;
    QRShape72: TQRShape;
    QRShape73: TQRShape;
    QRLabel83: TQRLabel;
    QRLabel84: TQRLabel;
    QRLabel85: TQRLabel;
    QRI_HMY: TQRImage;
    QRImage4: TQRImage;
    QRLabel89: TQRLabel;
    QRL_gum: TQRLabel;
    QRLabel147: TQRLabel;
    QRI_KAR: TQRImage;
    QRI_CUH: TQRImage;
    QRI_JGU: TQRImage;
    QRI_SHJ: TQRImage;
    QRI_UHN: TQRImage;
    QRI_HDR: TQRImage;
    QRI_CME: TQRImage;
    QRI_JSM: TQRImage;
    QRI_KYJ: TQRImage;
    QRShape86: TQRShape;
    QRShape87: TQRShape;
    QRLabel143: TQRLabel;
    QRL_gum2: TQRLabel;
    QRI_JMR: TQRImage;
    QRI_HMY2: TQRImage;
    QRI_LCS: TQRImage;
    QRI_HSI: TQRImage;
    QRI_YHJ: TQRImage;
    QRI_KSH: TQRImage;
    QRLabel152: TQRLabel;
    QRLabel153: TQRLabel;
    QRLabel154: TQRLabel;
    QRLabel155: TQRLabel;
    QRLabel156: TQRLabel;
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRepNeedData(Sender: TObject;
      var MoreData: Boolean);  
    procedure UP_ClearImage(Sender: TObject);
  private
     UV_sUrine_Char, UV_sDat_gmgn : string;
     UV_iCount : Integer;
     UV_vCod_hm : Variant;

     procedure UP_Clear;

     function UF_Check(Value : String): String;

    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD5P142: TfrmLD5P142;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;

implementation

uses ZZComm, ZZfunc, MdmsFunc, LD5P14;

{$R *.DFM}

procedure TfrmLD5P142.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
  inherited;
  UV_iCount := 1;
  qryGumgin.First;
end;

procedure TfrmLD5P142.UP_Clear;
begin
   QRL_Jumin.Caption        := '';
   QRL_Name.Caption         := '';
   QRL_Dat_gmgn.Caption     := '';
   QRL_Dat_bunju.Caption    := '';
   QRL_Numbohm.Caption      := '';
   QRL_H002.Caption         := '';
   QRL_C032.Caption         := '';
   QRL_C025.Caption         := '';
   QRL_C026.Caption         := '';
   QRL_C028.Caption         := '';
   QRL_C009.Caption         := '';
   QRL_C010.Caption         := '';
   QRL_C074.Caption         := '';
   QRL_C011.Caption         := '';
   QRL_C042.Caption         := '';
   QRL_S007.Caption         := '';
   QRL_S008.Caption         := '';
   QRL_S099.Caption         := '';
   QRL_S091.Caption         := '';
   QRL_S016.Caption         := '';
   QRL_S033.Caption         := '';
   QRL_Y004.Caption         := '';
   QRL_TT03.Caption         := '';
   QRL_TT01.Caption         := '';
   QRL_TT02.Caption         := '';
   QRL_NumSamp1.Caption     := '';
   QRL_NumSamp2.Caption     := '';
   QRL_Kikwan.Caption       := '';
   QRL_Kiho.Caption         := '';
   QRL_doctor.Caption       := '';
   QRL_Telnum.Caption       := '';
   QRL_Dat_gum.Caption      := '';
   QRL_GFR.Caption          := '';
   QRL_C027.Caption         := '';
   QRL_Dat_Tobo.Caption     := '';
end;

function  TfrmLD5P142.UF_Check(Value : String): String;
begin
   Result := '';

   if Value <> '0' then Result := Value;
end;

procedure TfrmLD5P142.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
Begin
  With qryGumgin Do
  Begin
    If Eof Then
    Begin
      MoreData := False;
      Exit;
    End;

    If UV_iCount = 1 Then
    Begin
      MoreData := True;
      Inc(UV_iCount);
    End
    else If UV_iCount <= RecordCount Then
    Begin
      MoreData := True;
      Inc(UV_iCount);
      Next;
    End
    Else
      Moredata := False;
   end;
end;

procedure TfrmLD5P142.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
   iRet, i, iAge, iMatCnt, iRow, iTemp, iCnt : Integer;
   vCod_chuga : Variant;
   sHmCode, sName, sMan, sGubn, sTemp,
   sCode1, sCode2, sCode3, sCode4,
   sSelect, sWhere, sOrderby, sHangmok_Iist, sTime, sTime_H, sTime_M : string;
   eRslt : Real;

   //이미지 관련.
   Jpeg  : TJPEGImage;
   MS    : TMemoryStream;

begin
   UP_Clear;
   {
   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

   //항목불러오기...
   if not qryHm.Active then qryHm.Active := True;
    }
   sTime_H := '';
   sTime_M := '';
   sTime   := '';
   with qryGumgin do
   begin
      QRL_Jumin.Caption         :=  copy(FieldByName('num_jumin').AsString,1,6) +' - ' + copy(FieldByName('num_jumin').AsString,7,7);
      QRL_Dat_gmgn.Caption      :=  copy(FieldByName('dat_gmgn').AsString,1,4) + '. ' + copy(FieldByName('dat_gmgn').AsString,5,2) + '. ' + copy(FieldByName('dat_gmgn').AsString,7,2) +'. ';
      QRL_Dat_gum.Caption       :=  copy(FieldByName('dat_gmgn').AsString,1,4) + '. ' + copy(FieldByName('dat_gmgn').AsString,5,2) + '. ' + copy(FieldByName('dat_gmgn').AsString,7,2) +'. ';
      QRL_Name.Caption          := FieldByName('desc_name').AsString;
      sHangmok_Iist             := FieldByName('cod_chuga').AsString;
      QRL_NumSamp1.Caption      := FieldByName('num_samp').AsString;
      QRL_NumSamp2.Caption      := FieldByName('num_samp').AsString;if FieldByName('dat_gmgn').AsString > '20131231' then
      begin 
          QRL_Dat_gum.Caption       :=  copy(FieldByName('Today_insert').AsString,1,4) + '. ' + copy(FieldByName('Today_insert').AsString,5,2) + '. ' + copy(FieldByName('Today_insert').AsString,7,2) +'. ';
          sTime := copy(FieldByName('Today_insert').AsString,15,6);
          sTime_H := copy(sTime,1,pos(':', sTime)-1);
          sTime   := copy(sTime,pos(':', sTime)+1,3 );
          sTime_M := copy(sTime,1,pos(':', sTime)-1);
          if copy(FieldByName('Today_insert').AsString,12,2) = '전' then
             QRL_Dat_gum.Caption    := QRL_Dat_gum.Caption + sTime_H
          else if copy(FieldByName('Today_insert').AsString,12,2) = '후' then
             QRL_Dat_gum.Caption    := QRL_Dat_gum.Caption + IntToStr(Strtoint(sTime_H)+12);
          QRL_Dat_gum.Caption       := QRL_Dat_gum.Caption + ':' + sTime_M;
      end;
      if  (qryGumgin.FieldByName('gubn_life').AsString = '1') and
          (qryGumgin.FieldByName('gubn_lfyh').AsString = '1') then
      begin
          if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
             (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') then
              sHangmok_Iist := sHangmok_Iist + 'S099S091'
          else
              sHangmok_Iist := sHangmok_Iist + 'S007S008'
      end;

      if  (qryGumgin.FieldByName('gubn_life').AsString = '1') and
          ((qryGumgin.FieldByName('gubn_lfyh').AsString = '4') or
           (qryGumgin.FieldByName('gubn_lfyh').AsString = '5') or
            (qryGumgin.FieldByName('gubn_lfyh').AsString = '6')) then
      begin
              sHangmok_Iist := sHangmok_Iist + 'S033';
      end;

      GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);

      if FieldByName('cod_jisa').AsString  = '11' then
      begin
           QRL_Kikwan.Caption   := '한국의학연구소 강남의원';
           QRL_Kiho.Caption     := '11366338';
           QRL_doctor.Caption   := '정우진';
           QRL_Telnum.Caption   := '02-3496-3253';
      end
      else if FieldByName('cod_jisa').AsString  = '12' then
      begin
           if  FieldByName('Dat_gmgn').AsString <= '20170627' then
           begin
              QRL_Kikwan.Caption   := '케이엠아이 여의도의원';
              QRL_Kiho.Caption     := '11359994';
              QRL_doctor.Caption   := '주택소';
              QRL_Telnum.Caption   := '02-368-8206';
           end
           else
           begin
              QRL_Kikwan.Caption   := '케이엠아이 여의도의원';
              QRL_Kiho.Caption     := '11359994';
              QRL_doctor.Caption   := '이광설';
              QRL_Telnum.Caption   := '02-368-8206';
           end;
      end
      else if FieldByName('cod_jisa').AsString  = '43' then
      begin
           QRL_Kikwan.Caption   := '(재)KMI 경기의원';
           QRL_Kiho.Caption     := '31331254';
           if FieldByName('dat_gmgn').AsString >= '20160901' then
              QRL_doctor.Caption   := '김성훈'
           else
              QRL_doctor.Caption   := '허충';
           QRL_Telnum.Caption   := '031-231-0251';
      end
      else if FieldByName('cod_jisa').AsString  = '51' then
      begin
           QRL_Kikwan.Caption   := '한국의학연구소 광주의원 광주분사무소';
           QRL_Kiho.Caption     := '36332712';
           QRL_doctor.Caption   := '김연선';
           QRL_Telnum.Caption   := '062-602-2143';
      end
      else if FieldByName('cod_jisa').AsString  = '61' then
      begin
           QRL_Kikwan.Caption   := '재단법인 한국의학연구소 부산의원';
           QRL_Kiho.Caption     := '21343233';
           if copy(FieldByName('dat_gmgn').AsString,1,6) < '201512' then
               QRL_doctor.Caption   := '김성희'
           else
               QRL_doctor.Caption   := '이창희';
           QRL_Telnum.Caption   := '051-810-1500';
      end
      else if FieldByName('cod_jisa').AsString  = '71' then
      begin
           QRL_Kikwan.Caption   := '재) 한국의학연구소  대구분사무소 부설의원';
           QRL_Kiho.Caption     := '37315676';
           QRL_doctor.Caption   := '서준원' ;
           QRL_Telnum.Caption   := '053-430-5000';
      end;

      with qryProfile do
      begin
         Active := False;
         ParamByName('cod_pf').AsString    := qryGumgin.FieldByName('cod_hul').AsString;
         Active := True;
         if  0 < RecordCount then
         begin
             while not qryprofile.eof do
             begin
                sHangmok_Iist := sHangmok_Iist +  FieldByName('cod_hm').AsString;
                next;
             end;
         end;
         Active := False;
      end;

      with qryGulkwa do
      begin
          Active := False;
          ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
          ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
          ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
          Active := True;
          if Recordcount > 0 then
          begin
             QRL_Dat_bunju.Caption      :=  copy(FieldByName('Dat_bunju').AsString,1,4) + '. ' + copy(FieldByName('Dat_bunju').AsString,5,2) + '. ' + copy(FieldByName('Dat_bunju').AsString,7,2) +'. ';

             with qryHgulkwa_chk do
             begin
                Active := False;
                ParamByName('dat_bunju').AsString    := qryGulkwa.FieldByName('dat_bunju').AsString;
                Active := True;

                if Recordcount > 0 then
                begin
                    if FieldByName('gubn_05').AsString = 'Y' then
                    begin
                         QRL_Dat_Tobo.Caption      :=  copy(FieldByName('dat_update05').AsString,1,4) + '. ' + copy(FieldByName('dat_update05').AsString,5,2) + '. ' + copy(FieldByName('dat_update05').AsString,7,2) +'. ';
                    end;
                end;
             end;
          end;
          Active := False;
      end;    {
          if pos('H002', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''02''  ';
               QRL_H002.Caption := copy(FieldByName('desc_glkwa1').AsString,7,6);
          end;
          if pos('C032', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C032.Caption := copy(FieldByName('desc_glkwa1').AsString,157,6);
          end;
          if pos('C074', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C074.Caption := copy(FieldByName('desc_glkwa2').AsString,131,6);
          end;
          if pos('C025', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C025.Caption := copy(FieldByName('desc_glkwa1').AsString,121,6);
          end;
          if pos('C026', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C026.Caption := copy(FieldByName('desc_glkwa1').AsString,127,6);
          end;
          if pos('C027', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C027.Caption := copy(FieldByName('desc_glkwa1').AsString,133,6);
          end;
          if pos('C028', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C028.Caption := copy(FieldByName('desc_glkwa1').AsString,139,6);
          end;
          if pos('C009', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C009.Caption := copy(FieldByName('desc_glkwa1').AsString,49,6);
          end;
          if pos('C010', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C010.Caption := copy(FieldByName('desc_glkwa1').AsString,55,6);
          end;
          if pos('C011', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C011.Caption := copy(FieldByName('desc_glkwa1').AsString,61,6);
          end;
          if pos('C087', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_GFR.Caption := copy(FieldByName('desc_glkwa3').AsString,15,6);
          end;
          if pos('C042', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''01''  ';
               QRL_C042.Caption := copy(FieldByName('desc_glkwa1').AsString,187,6);
               // 삭제구간
               if (QRL_C042.Caption <> '') and (QRL_C042.Caption <> '0') and (QRL_C042.Caption <> '0.0') and (pos('C087', sHangmok_Iist) > 0) then
               begin
                    if (Copy(QRL_Jumin.Caption ,10,1) = '1') or
                       (Copy(QRL_Jumin.Caption ,10,1) = '3') or
                       (Copy(QRL_Jumin.Caption ,10,1) = '5') or
                       (Copy(QRL_Jumin.Caption ,10,1) = '7') then
                    begin
                       eRslt := 186 * Power(strToFloat(QRL_C042.Caption),-1.154) * Power(iAge, -0.203);
                    end
                    else if (Copy(QRL_Jumin.Caption ,10,1) = '2') or
                            (Copy(QRL_Jumin.Caption ,10,1) = '4') or
                            (Copy(QRL_Jumin.Caption ,10,1) = '6') or
                            (Copy(QRL_Jumin.Caption ,10,1) = '8') then
                    begin
                       eRslt := 186 * Power(strToFloat(QRL_C042.Caption),-1.154) * Power(iAge, -0.203) * 0.742 ;
                    end;
                    QRL_GFR.Caption := copy(FloatToStr(eRslt),1,10);
               end;//삭제구간
          end;
          if pos('S007', sHangmok_Iist) > 0 then
          begin
               qryGulkwa.Filter := ' gubn_part =''07''  ';
               QRL_S007.Caption := copy(FieldByName('desc_glkwa1').AsString,31,6);
          end;}
      if (qryGumgin.FieldByName('gubn_life').AsString = '1') and
         (qryGumgin.FieldByName('gubn_lfyh').AsString = '1') then
      begin
           with qryGulkwa_hm do
           begin
                Active := False;
                ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
                ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
                ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
                ParamByName('gubn_part').AsString  := '07';
                Active := True;

                if Recordcount > 0 then
                begin
                     if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') and
                        (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') then
                     begin
                          if (pos('S099', sHangmok_Iist)) > 0 then
                          begin
                               QRL_S099.Caption := copy(FieldByName('desc_glkwa2').AsString,29,6);
                          end;
                          if pos('S091', sHangmok_Iist) > 0 then
                          begin
                               QRL_S091.Caption := copy(FieldByName('desc_glkwa1').AsString,181,6);
                          end;
                     end
                     else
                     begin
                          if (pos('S007', sHangmok_Iist)) > 0 then
                          begin
                               QRL_S007.Caption := copy(FieldByName('desc_glkwa1').AsString,31,6);
                          end;
                          if pos('S008', sHangmok_Iist) > 0 then
                          begin
                               QRL_S008.Caption := copy(FieldByName('desc_glkwa1').AsString,37,6);
                          end;
                     end;
                end;
                Active := False;
           end;
      end;

      
      if (qryGumgin.FieldByName('gubn_life').AsString = '1') and
         ((qryGumgin.FieldByName('gubn_lfyh').AsString = '4') or
          (qryGumgin.FieldByName('gubn_lfyh').AsString = '5') or
          (qryGumgin.FieldByName('gubn_lfyh').AsString = '6') )then
      begin
           with qryGulkwa_hm do
           begin
                Active := False;
                ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
                ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
                ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
                ParamByName('gubn_part').AsString  := '05';
                Active := True;

                if Recordcount > 0 then
                begin
                    if (pos('S033', sHangmok_Iist)) > 0 then
                    begin
                         QRL_S033.Caption := copy(FieldByName('desc_glkwa3').AsString,57,6);
                    end;
                end;
                Active := False;
           end;
      end;

      if (pos('S016', sHangmok_Iist) > 0) and
         (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0) then
      begin
           with qryGulkwa_hm do
           begin
                Active := False;
                ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
                ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
                ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
                ParamByName('gubn_part').AsString  := '05';
                Active := True;

                if Recordcount > 0 then
                begin
                     QRL_S016.Caption := copy(FieldByName('desc_glkwa1').AsString,7,6);
                end;
                Active := False;
           end;
      end;
      if (pos('Y004', sHangmok_Iist) > 0) and
         (pos('대장',qryGumgin.FieldByName('desc_yhname').AsString) > 0) and
         (qryGumgin.FieldByName('dat_gmgn').AsString < '20160101')   then
      begin 
          with qryGulkwa_hm do
          begin
               Active := False;
               ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
               ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
               ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
               ParamByName('gubn_part').AsString  := '03';
               Active := True;

               if Recordcount > 0 then
               begin
                     QRL_Y004.Caption := copy(FieldByName('desc_glkwa1').AsString,85,6);
                end;
                Active := False;
          end;
      end;

      if (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0) then
      begin
           if (pos('TT03', sHangmok_Iist) > 0) and
           (qryGumgin.FieldByName('dat_gmgn').AsString >'20160101') then
           begin
                with qryGulkwa_hm do
                begin
                     Active := False;
                     ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
                     ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
                     ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
                     ParamByName('gubn_part').AsString  := '05';
                     Active := True;

                     if Recordcount > 0 then
                     begin
                          QRL_TT03.Caption := copy(FieldByName('desc_glkwa3').AsString,51,6);
                     end;
                     Active := False;
                end;
           end;{
           if pos('TT01', sHangmok_Iist) > 0 then
           begin
                qryGulkwa.Filter := ' gubn_part =''07''  ';
                QRL_TT01.Caption := copy(FieldByName('desc_glkwa1').AsString,85,6);
           end;}
           if (pos('TT02', sHangmok_Iist) > 0) and
           (qryGumgin.FieldByName('dat_gmgn').AsString <'20160101') then
           begin
                with qryGulkwa_hm do
                begin
                     Active := False;
                     ParamByName('num_id').AsString    := qryGumgin.FieldByName('num_id').AsString;
                     ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
                     ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
                     ParamByName('gubn_part').AsString  := '07';
                     Active := True;

                     if Recordcount > 0 then
                     begin
                         if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2014') or
                            (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2015') then
                                    QRL_TT02.Caption := copy(FieldByName('desc_glkwa1').AsString,169,6)
                         else if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2011') or
                                 (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2012') or
                                 (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2013') then
                                     QRL_TT03.Caption := copy(FieldByName('desc_glkwa1').AsString,169,6);
                     end;
                     Active := False;
                end;
           end;
      end;

      with qryNhic do
      begin
          Active := False;
          ParamByName('num_jumin').AsString := qryGumgin.FieldByName('num_jumin').AsString;
          ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
          Active := True;
          QRL_Numbohm.Caption      := FieldByName('num_bohm').AsString;
          Active := False;
      end;

      if  (QRL_Dat_bunju.Caption <> '') and
          ((QRL_S099.Caption <> '') or
          (QRL_S091.Caption <> '') or
          (QRL_S016.Caption <> '') or
          (QRL_Y004.Caption <> '') or
          (QRL_TT02.Caption <> '') or
          (QRL_TT03.Caption <> '') or
          (QRL_S007.Caption <> '') or
          (QRL_S008.Caption <> '') or
          (QRL_S033.Caption <> '')) then
             PrintBand := True
      else
             PrintBand := False;

      UP_ClearImage(self);

      if (((pos('S099', sHangmok_Iist) >0) or (pos('S091', sHangmok_Iist) >0)) and
          ((qryGumgin.FieldByName('gubn_life').AsString = '1') and
           (qryGumgin.FieldByName('gubn_lfyh').AsString = '1')) and
          (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
          (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012')) or
          (((pos('S007', sHangmok_Iist) >0) or (pos('S008', sHangmok_Iist) >0)) and
          ((qryGumgin.FieldByName('gubn_life').AsString = '1') and
           (qryGumgin.FieldByName('gubn_lfyh').AsString = '1')) and
          ((copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2011') or
           (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) = '2012'))) or
         ((pos('S016', sHangmok_Iist) >0)  and
          (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0)) or
         ((copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) >= '2017') and
          (qryGumgin.FieldByName('gubn_life').AsString = '1') and
          ((qryGumgin.FieldByName('gubn_lfyh').AsString = '4') or
           (qryGumgin.FieldByName('gubn_lfyh').AsString = '5') or
           (qryGumgin.FieldByName('gubn_lfyh').AsString = '6')) and
          (pos('S033', sHangmok_Iist) >0))  then
      begin
           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_gum.Caption    := '최민애';
                QRI_CME.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_gum.Caption    := '한두례';
                QRI_HDR.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_gum.Caption    := '한두례';
                QRI_HDR.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HDR.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') AND (FieldByName('dat_gmgn').AsString <= '20180304') then
           begin
                QRL_gum.Caption    := '김용진';
                QRI_KYJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2018' then
           begin
                QRL_gum.Caption    := '손혜정';
                QRI_SHJ.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20190101') then
           begin
                QRL_gum.Caption    := '손혜정';
                QRI_SHJ.Enabled      := True;
           end
      end
      else if ((pos('TT02', sHangmok_Iist) >0) or
               (pos('TT03', sHangmok_Iist) >0)) and
              (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0)  then
      begin 
           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_gum.Caption    := '최민애';
                QRI_CME.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_gum.Caption    := '한두례';
                QRI_HDR.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_gum.Caption    := '엄하니';
                QRI_UHN.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_gum.Caption    := '손혜정';
                QRI_SHJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2016'  then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') AND (FieldByName('dat_gmgn').AsString <= '20180304') then
           begin
                QRL_gum.Caption    := '김용진';
                QRI_KYJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2018' then
           begin
                QRL_gum.Caption    := '손혜정';
                QRI_SHJ.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20190101') then
           begin
                QRL_gum.Caption    := '손혜정';
                QRI_SHJ.Enabled      := True;
           end
      end
      else if (pos('Y004', sHangmok_Iist) >0)  and
              (pos('대장',qryGumgin.FieldByName('desc_yhname').AsString) > 0)then
      begin
           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_gum.Caption    := '정광우';
                QRI_JGU.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_gum.Caption    := '최은희';
                QRI_CUH.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_gum.Caption    := '김아름';
                QRI_KAR.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_gum.Caption    := '한두례';
                QRI_HDR.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_gum.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString <= '20170301')  then
           begin
                QRL_gum.Caption    := '정선미';
                QRI_JSM.Enabled      := True;
           end
           //센터검사로 넘어가서 없어짐. 박연숙 팀장.
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2018' then
            begin
                QRL_gum.Caption    := '';
           end
           else if (FieldByName('dat_gmgn').AsString >= '20190101') then
            begin
                QRL_gum.Caption    := '';
           end;
      end;
      //20170824수지
      if qryGumgin.FieldByName('cod_jisa').AsString  = '11' then
      begin
                QRL_gum2.Caption     := '전미라';
                QRI_JMR.Enabled      := True;
      end
      else if qryGumgin.FieldByName('cod_jisa').AsString  = '12' then
      begin
                QRL_gum2.Caption    := '한미영';
                QRI_HMY2.Enabled      := True;
      end
      else  if qryGumgin.FieldByName('cod_jisa').AsString  = '43' then
      begin
                QRL_gum2.Caption    := '이청심';
                QRI_LCS.Enabled      := True;
      end
      else if qryGumgin.FieldByName('cod_jisa').AsString  = '51' then
      begin
                QRL_gum2.Caption    := '한송이';
                QRI_HSI.Enabled      := True;
      end
      else if qryGumgin.FieldByName('cod_jisa').AsString  = '61' then
      begin
                QRL_gum2.Caption    := '유희짐';
                QRI_YHJ.Enabled      := True;
      end
      else if qryGumgin.FieldByName('cod_jisa').AsString  = '71' then
      begin
                QRL_gum2.Caption    := '김상희';
                QRI_KSH.Enabled      := True;
      end;
   end;
end;


procedure TfrmLD5P142.UP_ClearImage(Sender: TObject);
begin
  inherited;
     QRI_CME.Enabled      := False;
     QRI_HMY.Enabled      := False;
     QRI_HDR.Enabled      := False;
     QRI_UHN.Enabled      := False;
     QRI_SHJ.Enabled      := False;
     QRI_JGU.Enabled      := False;
     QRI_CUH.Enabled      := False;
     QRI_KAR.Enabled      := False;
     QRI_JSM.Enabled      := False;
     QRI_KYJ.Enabled      := False;

     QRI_JMR.Enabled      := False;
     QRI_HMY2.Enabled     := False;
     QRI_LCS.Enabled      := False;
     QRI_HSI.Enabled      := False;
     QRI_YHJ.Enabled      := False;
     QRI_KSH.Enabled      := False;
end;


end.
