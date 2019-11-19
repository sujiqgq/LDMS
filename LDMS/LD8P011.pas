//==============================================================================
// 프로그램 설명 : 혈액학 결과출력(출력 Form)
// 시스템        : 통합검진
// 수정일자      : 1999.10.30
// 수정자        : 허정남
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8P011;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD8P011 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRL_Date: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRS_Page: TQRSysData;
    QRLabel12: TQRLabel;
    qrlFrom: TQRLabel;
    QRL_1: TQRLabel;
    QRL_2: TQRLabel;
    QRL_3: TQRLabel;
    QRL_4: TQRLabel;
    QRL_8: TQRLabel;
    QRL_5: TQRLabel;
    QRL_6: TQRLabel;
    QRL_7: TQRLabel;
    QRL_12: TQRLabel;
    QRL_9: TQRLabel;
    QRL_10: TQRLabel;
    QRL_11: TQRLabel;
    QRL_13: TQRLabel;
    QRL_14: TQRLabel;
    QRL_27: TQRLabel;
    QRL_28: TQRLabel;
    QRL_26: TQRLabel;
    QRL_25: TQRLabel;
    QRL_24: TQRLabel;
    QRL_23: TQRLabel;
    QRL_22: TQRLabel;
    QRL_21: TQRLabel;
    QRL_20: TQRLabel;
    QRL_19: TQRLabel;
    QRL_18: TQRLabel;
    QRL_17: TQRLabel;
    QRL_16: TQRLabel;
    QRL_15: TQRLabel;
    QRL_41: TQRLabel;
    QRL_42: TQRLabel;
    QRL_40: TQRLabel;
    QRL_39: TQRLabel;
    QRL_38: TQRLabel;
    QRL_37: TQRLabel;
    QRL_36: TQRLabel;
    QRL_35: TQRLabel;
    QRL_34: TQRLabel;
    QRL_33: TQRLabel;
    QRL_32: TQRLabel;
    QRL_31: TQRLabel;
    QRL_30: TQRLabel;
    QRL_29: TQRLabel;
    QryHangmok: TQuery;
    QryGulkwa: TQuery;
    QRL_43: TQRLabel;
    QRL_44: TQRLabel;
    QRL_45: TQRLabel;
    QRL_46: TQRLabel;
    QRL_47: TQRLabel;
    QRL_48: TQRLabel;
    QRL_49: TQRLabel;
    QRL_50: TQRLabel;
    QRL_51: TQRLabel;
    QRL_52: TQRLabel;
    QRL_53: TQRLabel;
    QRL_54: TQRLabel;
    QRL_55: TQRLabel;
    QRL_56: TQRLabel;
    QRL_TITLE: TQRLabel;
    QRBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRL_D1: TQRLabel;
    QRL_D15: TQRLabel;
    QRL_D29: TQRLabel;
    QRL_D43: TQRLabel;
    QRL_D44: TQRLabel;
    QRL_D30: TQRLabel;
    QRL_D16: TQRLabel;
    QRL_D2: TQRLabel;
    QRL_D3: TQRLabel;
    QRL_D17: TQRLabel;
    QRL_D31: TQRLabel;
    QRL_D45: TQRLabel;
    QRL_D46: TQRLabel;
    QRL_D32: TQRLabel;
    QRL_D18: TQRLabel;
    QRL_D4: TQRLabel;
    QRL_D5: TQRLabel;
    QRL_D19: TQRLabel;
    QRL_D33: TQRLabel;
    QRL_D47: TQRLabel;
    QRL_D48: TQRLabel;
    QRL_D34: TQRLabel;
    QRL_D20: TQRLabel;
    QRL_D6: TQRLabel;
    QRL_D7: TQRLabel;
    QRL_D21: TQRLabel;
    QRL_D35: TQRLabel;
    QRL_D49: TQRLabel;
    QRL_D50: TQRLabel;
    QRL_D36: TQRLabel;
    QRL_D22: TQRLabel;
    QRL_D8: TQRLabel;
    QRL_D9: TQRLabel;
    QRL_D23: TQRLabel;
    QRL_D37: TQRLabel;
    QRL_D51: TQRLabel;
    QRL_D52: TQRLabel;
    QRL_D38: TQRLabel;
    QRL_D24: TQRLabel;
    QRL_D10: TQRLabel;
    QRL_D11: TQRLabel;
    QRL_D25: TQRLabel;
    QRL_D39: TQRLabel;
    QRL_D53: TQRLabel;
    QRL_D54: TQRLabel;
    QRL_D40: TQRLabel;
    QRL_D26: TQRLabel;
    QRL_D12: TQRLabel;
    QRL_D13: TQRLabel;
    QRL_D27: TQRLabel;
    QRL_D41: TQRLabel;
    QRL_D55: TQRLabel;
    QRL_D56: TQRLabel;
    QRL_D42: TQRLabel;
    QRL_D28: TQRLabel;
    QRL_D14: TQRLabel;
    DataSource1: TDataSource;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryTkgum: TQuery;
    QRLabel3: TQRLabel;
    QRDBText3: TQRDBText;
    procedure FormCreate(Sender: TObject);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
  private
    { Private declarations }
    UV_vCod_hm : Variant;
    UV_iCount : Integer;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Tk_Hangmok(sDate: String; iYh : Integer; iJumin : String): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    procedure UP_FieldClear(frm : TWinControl);

  public
    { Public declarations }
  end;

var
  frmLD8P011: TfrmLD8P011;

implementation

{$R *.DFM}

uses ZZComm, ZZfunc, MdmsFunc, LD8P01;


procedure TfrmLD8P011.FormCreate(Sender: TObject);
var i : integer;
   pos, sSelect, sWhere, sOrderBy : string;
begin
  inherited;
    with qryHangmok do
   begin
       Active := False;
       ParamByName('ST_Date').AsString  := frmLD8P01.mskSt_Date.Text;

       if  frmLD8P01.rbGubn_part.itemindex = 0 then
           ParamByName('Gubn_part').AsString  := '01' //생화학
       else if  frmLD8P01.rbGubn_part.itemindex = 1 then
           ParamByName('Gubn_part').AsString  := '02' //혈액학
       else if  frmLD8P01.rbGubn_part.itemindex = 2 then
           ParamByName('Gubn_part').AsString  := '03' //urine
       else if  frmLD8P01.rbGubn_part.itemindex = 3 then
           ParamByName('Gubn_part').AsString  := '04' //Ria
       else if  frmLD8P01.rbGubn_part.itemindex = 4 then
           ParamByName('Gubn_part').AsString  := '05' //EIA
       else if  frmLD8P01.rbGubn_part.itemindex = 5 then
           ParamByName('Gubn_part').AsString  := '07'; //혈청학
       Active := True;

       if  RecordCount > 0 then
       begin
          for i := 0 to  RecordCount -1  do
          begin
             pos := FieldbyName('Pos_glkwa').AsString;
             case  StrToInt(pos) of
              01 :  QRL_1.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              02 :  QRL_2.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              03 :  QRL_3.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              04 :  QRL_4.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              05 :  QRL_5.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              06 :  QRL_6.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              07 :  QRL_7.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              08 :  QRL_8.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              09 :  QRL_9.Caption  := copy(FieldByname('Desc_eng').AsString,1,6);
              10 :  QRL_10.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              11 :  QRL_11.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              12 :  QRL_12.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              13 :  QRL_13.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              14 :  QRL_14.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              15 :  QRL_15.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              16 :  QRL_16.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              17 :  QRL_17.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              18 :  QRL_18.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              19 :  QRL_19.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              20 :  QRL_20.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              21 :  QRL_21.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              22 :  QRL_22.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              23 :  QRL_23.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              24 :  QRL_24.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              25 :  QRL_25.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              26 :  QRL_26.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              27 :  QRL_27.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              28 :  QRL_28.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              29 :  QRL_29.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              30 :  QRL_30.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              31 :  QRL_31.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              32 :  QRL_32.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              33 :  QRL_33.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              34 :  QRL_34.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              35 :  QRL_35.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              36 :  QRL_36.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              37 :  QRL_37.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              38 :  QRL_38.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              39 :  QRL_39.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              40 :  QRL_40.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              41 :  QRL_41.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              42 :  QRL_42.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              43 :  QRL_43.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              44 :  QRL_44.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              45 :  QRL_45.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              46 :  QRL_46.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              47 :  QRL_47.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              48 :  QRL_48.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              49 :  QRL_49.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              50 :  QRL_50.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              51 :  QRL_51.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              52 :  QRL_52.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              53 :  QRL_53.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              54 :  QRL_54.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              55 :  QRL_55.Caption := copy(FieldByname('Desc_eng').AsString,1,6);
              56 :  QRL_56.Caption := copy(FieldByname('Desc_eng').AsString,1,6);

               end;
              next;
           end;
       end;
   end;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   with qryGulkwa do
   begin
      Active := False;
      sSelect := ' SELECT A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.desc_glkwa1, a.desc_glkwa2, a.desc_glkwa3, ' +
                 '        a.gubn_part, SubString(B.desc_name,1,10) As Desc_name, B.num_jumin, B.dat_gmgn, ' +
                 '        B.cod_jangbi, B.cod_hul, B.cod_cancer, B.cod_chuga, B.gubn_nosin, B.gubn_nsyh, ' +
                 '        B.gubn_bogun, B.gubn_bgyh, B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, B.gubn_life, ' +
                 '        B.gubn_lfyh, B.gubn_gong, B.gubn_goyh, B.gubn_tkgm, B.gubn_hulgum, B.num_samp ' +
                 ' FROM   gulkwa A INNER JOIN gumgin B ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ';

      sWhere := ' WHERE A.cod_bjjs = ''' + Copy(frmLD8P01.cmbJisa.Text, 1, 2) + '''';
      sWhere := sWhere + ' and  a.dat_bunju = ''' + frmLD8P01.mskST_Date.Text + '''';
      if frmLD8P01.mskST_Date.Text = '20081025' then
        sWhere := sWhere + ' and  (A.cod_bjjs = ''15'' and a.num_bunju <> ''1924'' )';
      case frmLD8P01.rbGubn_part.itemindex of
         0 : sWhere := sWhere + ' and  a.gubn_part = ''01''';
         1 : sWhere := sWhere + ' and  a.gubn_part = ''02''';
         2 : sWhere := sWhere + ' and  a.gubn_part = ''03''';
         3 : sWhere := sWhere + ' and  a.gubn_part = ''04''';
         4 : sWhere := sWhere + ' and  a.gubn_part = ''05''';
         5 : sWhere := sWhere + ' and  a.gubn_part = ''07''';
      end;
      if frmLD8P01.CB_Hulgum.Checked then sWhere := sWhere + ' and  b.gubn_hulgum in (''3'')'
      else                                sWhere := sWhere + ' and  b.gubn_hulgum in (''1'',''2'',''3'')';

      if frmLD8P01.Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(frmLD8P01.mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju >= ''' + frmLD8P01.mskFrom.Text + '''';
         if Trim(frmLD8P01.mskTo.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju <= ''' + frmLD8P01.mskTo.Text + '''';

         sOrderBy := ' ORDER BY A.num_bunju ';
      end
      else
      begin
         if Trim(frmLD8P01.MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp >= ''' + frmLD8P01.MEdt_SampS.Text + '''';
         if Trim(frmLD8P01.MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp <= ''' + frmLD8P01.MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;
      GP_query_log(qryGulkwa, 'LD8P011', 'Q', 'N');
   end;
end;

procedure TfrmLD8P011.DataSource1DataChange(Sender: TObject;
  Field: TField);

var
     iStart, iEnd, i, index, temp, iAge, iSil, iRet, temp2, temp3 : integer;
     pos, gul, C_Hm, cod_name, sMan, sGubn, G_Hangmok  : string;
     C_Check : Boolean;
     vCod_chuga, vCod_totyh : Variant;
     eLow, eHigh, T_gulkwa : Extended;
     
begin

  inherited;
  UP_FieldClear(QRBand1);
  UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

  // 주민번호를 이용하여 성별과 나이를 구함
  GP_JuminToAge(QryGulkwa.FieldByname('num_jumin').AsString,QryGulkwa.FieldByname('Dat_gmgn').AsString, sMan, iAge);

  // 성별과 나이를 기준으로 Column명을 추출
  sGubn := '';
  if iAge < 10 then sGubn := '1'
  else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
  else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
  else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
  else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
  else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
  else sGubn := '7';

  if sMan = 'F' then sGubn := sGubn + 'f';

   i := 0;  iStart := 0;  iEnd := 0;
   eLow := 0; eHigh := 0; index := 0; T_gulkwa := 0;

      //검사항목축출
      if QryGulkwa.FieldByName('cod_hul').AsString <> '' then
      begin
         qryPf_hm.Active := False;
         qryPf_hm.ParamByName('COD_PF').AsString := QryGulkwa.FieldByName('cod_hul').AsString;
         qryPf_hm.Active := True;

         if qryPf_hm.RecordCount > 0 then
         begin
            VarArrayRedim(UV_vCod_hm, qryPf_hm.RecordCount-1);
            while not qryPf_hm.Eof do
            begin
               UV_vCod_hm[index] := qryPf_hm.FieldByName('COD_HM').AsString;
               Inc(index);
               qryPf_hm.Next;
            end;
         end;
      end;

      if QryGulkwa.FieldByName('cod_cancer').AsString <> '' then
      begin
         qryPf_hm.Active := False;
         qryPf_hm.ParamByName('COD_PF').AsString := QryGulkwa.FieldByName('cod_cancer').AsString;
         qryPf_hm.Active := True;

         if qryPf_hm.RecordCount > 0 then
         begin
            VarArrayRedim(UV_vCod_hm, index+qryPf_hm.RecordCount-1);
            while not qryPf_hm.Eof do
            begin
               UV_vCod_hm[index] := qryPf_hm.FieldByName('COD_HM').AsString;
               Inc(index);
               qryPf_hm.Next;
            end;
         end;
      end;

      if QryGulkwa.FieldByName('cod_jangbi').AsString <> '' then
      begin
         qryPf_hm.Active := False;
         qryPf_hm.ParamByName('COD_PF').AsString := QryGulkwa.FieldByName('cod_jangbi').AsString;
         qryPf_hm.Active := True;

         if qryPf_hm.RecordCount > 0 then
         begin
            VarArrayRedim(UV_vCod_hm, index+qryPf_hm.RecordCount-1);
            while not qryPf_hm.Eof do
            begin
               if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               begin
                 UV_vCod_hm[index] := qryPf_hm.FieldByName('COD_HM').AsString;
                 Inc(index);
               end;
               qryPf_hm.Next;
            end;
         end;
      end;

      iRet := GF_MulToSingle(QryGulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
      VarArrayRedim(UV_vCod_hm, index+iRet);
      for temp2 := 0 to iRet-1 do
      begin
        if copy(vCod_chuga[temp2],1,2) <> 'JJ' then
        begin
          UV_vCod_hm[index] := vCod_chuga[temp2];
          Inc(index);
        end;
      end;

      // 4. 노신, 성인병, 간염에 대한 검사항목 추출
      cod_name := '';

      // 노신유형Display
      if QryGulkwa.FieldByName('gubn_nosin').AsString = '1' then
         cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', QryGulkwa.FieldByName('gubn_nsyh').AsInteger);

      //성인병유형Display
      if QryGulkwa.FieldByName('gubn_adult').AsString = '1' then
         cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', QryGulkwa.FieldByName('gubn_adyh').AsInteger);

      //간염유형Display
      if QryGulkwa.FieldByName('gubn_agab').AsString = '1' then
         cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', QryGulkwa.FieldByName('gubn_agyh').AsInteger);

      // 생애전환기유형Display
      if QryGulkwa.FieldByName('gubn_life').AsString = '1' then
         cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', QryGulkwa.FieldByName('gubn_lfyh').AsInteger);

      //특검유형Display
      if QryGulkwa.FieldByName('gubn_tkgm').AsString <> '' then
         cod_name := cod_name + UF_Tk_Hangmok(QryGulkwa.FieldByName('Dat_gmgn').AsString,
                                              QryGulkwa.FieldByName('gubn_tkgm').AsInteger,
                                              QryGulkwa.FieldByName('Num_Jumin').AsString);
      iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
      VarArrayRedim(UV_vCod_hm, index+iRet);
      for temp3 := 0 to iRet-1 do
      begin
        if copy(vCod_totyh[temp3],1,2) <> 'JJ' then
        begin
          UV_vCod_hm[index] := vCod_totyh[temp3];
          Inc(index);
        end;
      end;

  qryHangmok.first;
  for i := 0 to  qryHangmok.RecordCount -1  do
   begin
      C_Check   := False;
      G_Hangmok := qryHangmok.FieldbyName('gubn_hm').AsString;
      C_Hm      := qryHangmok.FieldbyName('cod_hm').AsString;
      eLow      := qryHangmok.FieldByName('IMS_LOW'   + sGubn).AsFloat;
      eHigh     := qryHangmok.FieldByName('IMS_HIGH'  + sGubn).AsFloat;
      iStart    := qryHangmok.FieldByName('Pos_Start').AsInteger;
      iEnd      := qryHangmok.FieldByName('Pos_End').AsInteger;
      pos       := qryHangmok.FieldbyName('Pos_glkwa').AsString;
      gul       := qryGulkwa.FieldByName('Desc_glkwa1').AsString +
                 qryGulkwa.FieldByName('Desc_glkwa2').AsString +
                 qryGulkwa.FieldByName('Desc_glkwa3').AsString;

      for temp := 0 to index - 1 do
      begin
        if  C_Hm = UV_vCod_hm[temp] then C_Check := True;
      end;

      if C_Check = True then
      begin
        Case   StrToInt(pos)   of
         1 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D1.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D1.Color := clYellow;
                     end else
                     begin
                       QRL_D1.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D1.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D1.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D1.Caption) = '양성' then
                        QRL_D1.Color := clYellow
                     else
                        QRL_D1.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D1.Caption  := '****';
                  QRL_D1.Color := clWhite;
               end;
             end;
         2 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D2.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D2.Color := clYellow;
                     end else
                     begin
                       QRL_D2.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D2.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D2.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D2.Caption) = '양성' then
                        QRL_D2.Color := clYellow
                     else
                        QRL_D2.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D2.Caption  := '****';
                  QRL_D2.Color := clWhite;
               end;
             end;
         3 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D3.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D3.Color := clYellow;
                     end else
                     begin
                       QRL_D3.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D3.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D3.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D3.Caption) = '양성' then
                        QRL_D3.Color := clYellow
                     else
                        QRL_D3.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D3.Caption  := '****';
                  QRL_D3.Color := clWhite;
               end;
             end;
         4 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D4.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D4.Color := clYellow;
                     end else
                     begin
                       QRL_D4.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D4.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D4.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D4.Caption) = '양성' then
                        QRL_D4.Color := clYellow
                     else
                        QRL_D4.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D4.Caption  := '****';
                  QRL_D4.Color := clWhite;
               end;
             end;
         5 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D5.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D5.Color := clYellow;
                     end else
                     begin
                       QRL_D5.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D5.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D5.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D5.Caption) = '양성' then
                        QRL_D5.Color := clYellow
                     else
                        QRL_D5.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D5.Caption  := '****';
                  QRL_D5.Color := clWhite;
               end;
             end;
         6 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D6.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D6.Color := clYellow;
                     end else
                     begin
                       QRL_D6.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D6.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D6.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D6.Caption) = '양성' then
                        QRL_D6.Color := clYellow
                     else
                        QRL_D6.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D6.Caption  := '****';
                  QRL_D6.Color := clWhite;
               end;
             end;
         7 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D7.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D7.Color := clYellow;
                     end else
                     begin
                       QRL_D7.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D7.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D7.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D7.Caption) = '양성' then
                        QRL_D7.Color := clYellow
                     else
                        QRL_D7.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D7.Caption  := '****';
                  QRL_D7.Color := clWhite;
               end;
             end;
         8 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D8.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D8.Color := clYellow;
                     end else
                     begin
                       QRL_D8.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D8.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D8.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D8.Caption) = '양성' then
                        QRL_D8.Color := clYellow
                     else
                        QRL_D8.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D8.Caption  := '****';
                  QRL_D8.Color := clWhite;
               end;
             end;
         9 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D9.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D9.Color := clYellow;
                     end else
                     begin
                       QRL_D9.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D9.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D9.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D9.Caption) = '양성' then
                        QRL_D9.Color := clYellow
                     else
                        QRL_D9.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D9.Caption  := '****';
                  QRL_D9.Color := clWhite;
               end;
             end;
        10 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D10.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D10.Color := clYellow;
                     end else
                     begin
                       QRL_D10.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D10.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D10.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D10.Caption) = '양성' then
                        QRL_D10.Color := clYellow
                     else
                        QRL_D10.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D10.Caption  := '****';
                  QRL_D10.Color := clWhite;
               end;
             end;
        11 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D11.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D11.Color := clYellow;
                     end else
                     begin
                       QRL_D11.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D11.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D11.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D11.Caption) = '양성' then
                        QRL_D11.Color := clYellow
                     else
                        QRL_D11.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D11.Caption  := '****';
                  QRL_D11.Color := clWhite;
               end;
             end;
        12 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D12.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D12.Color := clYellow;
                     end else
                     begin
                       QRL_D12.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D12.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D12.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D12.Caption) = '양성' then
                        QRL_D12.Color := clYellow
                     else
                        QRL_D12.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D12.Caption  := '****';
                  QRL_D12.Color := clWhite;
               end;
             end;
        13 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D13.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D13.Color := clYellow;
                     end else
                     begin
                       QRL_D13.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D13.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D13.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D13.Caption) = '양성' then
                        QRL_D13.Color := clYellow
                     else
                        QRL_D13.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D13.Caption  := '****';
                  QRL_D13.Color := clWhite;
               end;
             end;
        14 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D14.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D14.Color := clYellow;
                     end else
                     begin
                       QRL_D14.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D14.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D14.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D14.Caption) = '양성' then
                        QRL_D14.Color := clYellow
                     else
                        QRL_D14.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D14.Caption  := '****';
                  QRL_D14.Color := clWhite;
               end;
             end;
        15 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D15.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D15.Color := clYellow;
                     end else
                     begin
                       QRL_D15.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D15.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D15.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D15.Caption) = '양성' then
                        QRL_D15.Color := clYellow
                     else
                        QRL_D15.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D15.Caption  := '****';
                  QRL_D15.Color := clWhite;
               end;
             end;
        16 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D16.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D16.Color := clYellow;
                     end else
                     begin
                       QRL_D16.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D16.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D16.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D16.Caption) = '양성' then
                        QRL_D16.Color := clYellow
                     else
                        QRL_D16.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D16.Caption  := '****';
                  QRL_D16.Color := clWhite;
               end;
             end;
        17 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D17.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D17.Color := clYellow;
                     end else
                     begin
                       QRL_D17.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D17.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D17.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D17.Caption) = '양성' then
                        QRL_D17.Color := clYellow
                     else
                        QRL_D17.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D17.Caption  := '****';
                  QRL_D17.Color := clWhite;
               end;
             end;
        18 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D18.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D18.Color := clYellow;
                     end else
                     begin
                       QRL_D18.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D18.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D18.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D18.Caption) = '양성' then
                        QRL_D18.Color := clYellow
                     else
                        QRL_D18.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D18.Caption  := '****';
                  QRL_D18.Color := clWhite;
               end;
             end;
        19 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D19.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D19.Color := clYellow;
                     end else
                     begin
                       QRL_D19.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D19.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D19.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D19.Caption) = '양성' then
                        QRL_D19.Color := clYellow
                     else
                        QRL_D19.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D19.Caption  := '****';
                  QRL_D19.Color := clWhite;
               end;
             end;
        20 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D20.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D20.Color := clYellow;
                     end else
                     begin
                       QRL_D20.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D20.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D20.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D20.Caption) = '양성' then
                        QRL_D20.Color := clYellow
                     else
                        QRL_D20.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D20.Caption  := '****';
                  QRL_D20.Color := clWhite;
               end;
             end;
        21 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D21.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D21.Color := clYellow;
                     end else
                     begin
                       QRL_D21.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D21.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D21.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D21.Caption) = '양성' then
                        QRL_D21.Color := clYellow
                     else
                        QRL_D21.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D21.Caption  := '****';
                  QRL_D21.Color := clWhite;
               end;
             end;
        22 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D22.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D22.Color := clYellow;
                     end else
                     begin
                       QRL_D22.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D22.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D2.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D2.Caption) = '양성' then
                        QRL_D2.Color := clYellow
                     else
                        QRL_D2.Color := clWhite;
                  end;
                  begin
                     QRL_D22.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if QRL_D22.Caption = '양성' then
                        QRL_D22.Color := clYellow
                     else
                        QRL_D22.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D22.Caption  := '****';
                  QRL_D22.Color := clWhite;
               end;
             end;
        23 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D23.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D23.Color := clYellow;
                     end else
                     begin
                       QRL_D23.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D23.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D23.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D23.Caption) = '양성' then
                        QRL_D23.Color := clYellow
                     else
                        QRL_D23.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D23.Caption  := '****';
                  QRL_D23.Color := clWhite;
               end;
             end;
        24 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D24.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D24.Color := clYellow;
                     end else
                     begin
                       QRL_D24.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D24.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D24.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D24.Caption) = '양성' then
                        QRL_D24.Color := clYellow
                     else
                        QRL_D24.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D24.Caption  := '****';
                  QRL_D24.Color := clWhite;
               end;
             end;
        25 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D25.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D25.Color := clYellow;
                     end else
                     begin
                       QRL_D25.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D25.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D25.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D25.Caption) = '양성' then
                        QRL_D25.Color := clYellow
                     else
                        QRL_D25.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D25.Caption  := '****';
                  QRL_D25.Color := clWhite;
               end;
             end;
        26 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D26.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D26.Color := clYellow;
                     end else
                     begin
                       QRL_D26.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D26.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D26.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D26.Caption) = '양성' then
                        QRL_D26.Color := clYellow
                     else
                        QRL_D26.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D26.Caption  := '****';
                  QRL_D26.Color := clWhite;
               end;
             end;
        27 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D27.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D27.Color := clYellow;
                     end else
                     begin
                       QRL_D27.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D27.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D27.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D27.Caption) = '양성' then
                        QRL_D27.Color := clYellow
                     else
                        QRL_D27.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D27.Caption  := '****';
                  QRL_D27.Color := clWhite;
               end;
             end;
        28 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D28.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D28.Color := clYellow;
                     end else
                     begin
                       QRL_D28.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D28.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D28.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D28.Caption) = '양성' then
                        QRL_D28.Color := clYellow
                     else
                        QRL_D28.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D28.Caption  := '****';
                  QRL_D28.Color := clWhite;
               end;
             end;
        29 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D29.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D29.Color := clYellow;
                     end else
                     begin
                       QRL_D29.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D29.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D29.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D29.Caption) = '양성' then
                        QRL_D29.Color := clYellow
                     else
                        QRL_D29.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D29.Caption  := '****';
                  QRL_D29.Color := clWhite;
               end;
             end;
        30 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D30.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D30.Color := clYellow;
                     end else
                     begin
                       QRL_D30.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D30.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D30.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D30.Caption) = '양성' then
                        QRL_D30.Color := clYellow
                     else
                        QRL_D30.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D30.Caption  := '****';
                  QRL_D30.Color := clWhite;
               end;
             end;
        31 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D31.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D31.Color := clYellow;
                     end else
                     begin
                       QRL_D31.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D31.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D31.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D31.Caption) = '양성' then
                        QRL_D31.Color := clYellow
                     else
                        QRL_D31.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D31.Caption  := '****';
                  QRL_D31.Color := clWhite;
               end;
             end;
        32 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D32.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D32.Color := clYellow;
                     end else
                     begin
                       QRL_D32.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D32.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D32.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D32.Caption) = '양성' then
                        QRL_D32.Color := clYellow
                     else
                        QRL_D32.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D32.Caption  := '****';
                  QRL_D32.Color := clWhite;
               end;
             end;
        33 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D33.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D33.Color := clYellow;
                     end else
                     begin
                       QRL_D33.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D33.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D33.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D33.Caption) = '양성' then
                        QRL_D33.Color := clYellow
                     else
                        QRL_D33.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D33.Caption  := '****';
                  QRL_D33.Color := clWhite;
               end;
             end;
        34 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D34.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D34.Color := clYellow;
                     end else
                     begin
                       QRL_D34.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D34.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D34.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D34.Caption) = '양성' then
                        QRL_D34.Color := clYellow
                     else
                        QRL_D34.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D34.Caption  := '****';
                  QRL_D34.Color := clWhite;
               end;
             end;
        35 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D35.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D35.Color := clYellow;
                     end else
                     begin
                       QRL_D35.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D35.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D35.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D35.Caption) = '양성' then
                        QRL_D35.Color := clYellow
                     else
                        QRL_D35.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D35.Caption  := '****';
                  QRL_D35.Color := clWhite;
               end;
             end;
        36 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D36.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D36.Color := clYellow;
                     end else
                     begin
                       QRL_D36.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D36.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D36.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D36.Caption) = '양성' then
                        QRL_D36.Color := clYellow
                     else
                        QRL_D36.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D36.Caption  := '****';
                  QRL_D36.Color := clWhite;
               end;
             end;
        37 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D37.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D37.Color := clYellow;
                     end else
                     begin
                       QRL_D37.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D37.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D37.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D37.Caption) = '양성' then
                        QRL_D37.Color := clYellow
                     else
                        QRL_D37.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D37.Caption  := '****';
                  QRL_D37.Color := clWhite;
               end;
             end;
        38 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D38.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D38.Color := clYellow;
                     end else
                     begin
                       QRL_D38.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D38.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D38.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D38.Caption) = '양성' then
                        QRL_D38.Color := clYellow
                     else
                        QRL_D38.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D38.Caption  := '****';
                  QRL_D38.Color := clWhite;
               end;
             end;
        39 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D39.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D39.Color := clYellow;
                     end else
                     begin
                       QRL_D39.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D39.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D39.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D39.Caption) = '양성' then
                        QRL_D39.Color := clYellow
                     else
                        QRL_D39.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D39.Caption  := '****';
                  QRL_D39.Color := clWhite;
               end;
             end;
        40 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D40.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D40.Color := clYellow;
                     end else
                     begin
                       QRL_D40.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D40.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D40.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D40.Caption) = '양성' then
                        QRL_D40.Color := clYellow
                     else
                        QRL_D40.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D40.Caption  := '****';
                  QRL_D40.Color := clWhite;
               end;
             end;
        41 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D41.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D41.Color := clYellow;
                     end else
                     begin
                       QRL_D41.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D41.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D41.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D41.Caption) = '양성' then
                        QRL_D41.Color := clYellow
                     else
                        QRL_D41.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D41.Caption  := '****';
                  QRL_D41.Color := clWhite;
               end;
             end;
        42 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D42.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D42.Color := clYellow;
                     end else
                     begin
                       QRL_D42.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D42.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D42.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D42.Caption) = '양성' then
                        QRL_D42.Color := clYellow
                     else
                        QRL_D42.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D42.Caption  := '****';
                  QRL_D42.Color := clWhite;
               end;
             end;
        43 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D43.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D43.Color := clYellow;
                     end else
                     begin
                       QRL_D43.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D43.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D43.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D43.Caption) = '양성' then
                        QRL_D43.Color := clYellow
                     else
                        QRL_D43.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D43.Caption  := '****';
                  QRL_D43.Color := clWhite;
               end;
             end;
        44 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D44.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D44.Color := clYellow;
                     end else
                     begin
                       QRL_D44.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D44.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D44.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D44.Caption) = '양성' then
                        QRL_D44.Color := clYellow
                     else
                        QRL_D44.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D44.Caption  := '****';
                  QRL_D44.Color := clWhite;
               end;
             end;
        45 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D45.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D45.Color := clYellow;
                     end else
                     begin
                       QRL_D45.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D45.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D45.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D45.Caption) = '양성' then
                        QRL_D45.Color := clYellow
                     else
                        QRL_D45.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D45.Caption  := '****';
                  QRL_D45.Color := clWhite;
               end;
             end;
        46 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D46.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D46.Color := clYellow;
                     end else
                     begin
                       QRL_D46.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D46.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D46.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D46.Caption) = '양성' then
                        QRL_D46.Color := clYellow
                     else
                        QRL_D46.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D46.Caption  := '****';
                  QRL_D46.Color := clWhite;
               end;
             end;
        47 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D47.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D47.Color := clYellow;
                     end else
                     begin
                       QRL_D47.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D47.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D47.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D47.Caption) = '양성' then
                        QRL_D47.Color := clYellow
                     else
                        QRL_D47.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D47.Caption  := '****';
                  QRL_D47.Color := clWhite;
               end;
             end;
        48 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D48.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D48.Color := clYellow;
                     end else
                     begin
                       QRL_D48.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D48.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D48.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D48.Caption) = '양성' then
                        QRL_D48.Color := clYellow
                     else
                        QRL_D48.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D48.Caption  := '****';
                  QRL_D48.Color := clWhite;
               end;
             end;
        49 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D49.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D49.Color := clYellow;
                     end else
                     begin
                       QRL_D49.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D49.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D49.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D49.Caption) = '양성' then
                        QRL_D49.Color := clYellow
                     else
                        QRL_D49.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D49.Caption  := '****';
                  QRL_D49.Color := clWhite;
               end;
             end;
        50 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D50.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D50.Color := clYellow;
                     end else
                     begin
                       QRL_D50.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D50.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D50.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D50.Caption) = '양성' then
                        QRL_D50.Color := clYellow
                     else
                        QRL_D50.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D50.Caption  := '****';
                  QRL_D50.Color := clWhite;
               end;
             end;
        51 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D51.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D51.Color := clYellow;
                     end else
                     begin
                       QRL_D51.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D51.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D51.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D51.Caption) = '양성' then
                        QRL_D51.Color := clYellow
                     else
                        QRL_D51.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D51.Caption  := '****';
                  QRL_D51.Color := clWhite;
               end;
             end;
        52 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D52.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D52.Color := clYellow;
                     end else
                     begin
                       QRL_D52.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D52.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D52.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D52.Caption) = '양성' then
                        QRL_D52.Color := clYellow
                     else
                        QRL_D52.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D52.Caption  := '****';
                  QRL_D52.Color := clWhite;
               end;
             end;
        53 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D53.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D53.Color := clYellow;
                     end else
                     begin
                       QRL_D53.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D53.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D53.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D53.Caption) = '양성' then
                        QRL_D53.Color := clYellow
                     else
                        QRL_D53.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D53.Caption  := '****';
                  QRL_D53.Color := clWhite;
               end;
             end;
        54 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D54.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D54.Color := clYellow;
                     end else
                     begin
                       QRL_D54.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D54.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D54.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D54.Caption) = '양성' then
                        QRL_D54.Color := clYellow
                     else
                        QRL_D54.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D54.Caption  := '****';
                  QRL_D54.Color := clWhite;
               end;
             end;
        55 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D55.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D55.Color := clYellow;
                     end else
                     begin
                       QRL_D55.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D55.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D55.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D55.Caption) = '양성' then
                        QRL_D55.Color := clYellow
                     else
                        QRL_D55.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D55.Caption  := '****';
                  QRL_D55.Color := clWhite;
               end;
             end;
        56 : begin
               if Trim(copy(gul, iStart, iEnd - iStart + 1)) <> '' then
               begin
                  if G_Hangmok = '1' then
                  begin
                     T_gulkwa := StrToFloat(copy(gul, iStart, iEnd - iStart + 1));
                     if (T_gulkwa < eLOW) or (T_gulkwa > eHigh) then
                     begin
                       QRL_D56.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D56.Color := clYellow;
                     end else
                     begin
                       QRL_D56.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                       QRL_D56.Color := clWhite;
                     end;
                  end else
                  begin
                     QRL_D56.Caption  := copy(gul, iStart, iEnd - iStart + 1);
                     if trim(QRL_D56.Caption) = '양성' then
                        QRL_D56.Color := clYellow
                     else
                        QRL_D56.Color := clWhite;
                  end;
               end
               else
               begin
                  QRL_D56.Caption  := '****';
                  QRL_D56.Color := clWhite;
               end;
             end;
        end; //case end;
      end; //if end;
   qryHangmok.next;
   end;// for end;
end;

function  TfrmLD8P011.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryNo_hangmok do
   begin
      Active := False;
      ParamByName('dat_apply').AsString   := sDate;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('gubn_yh').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

function  TfrmLD8P011.UF_Tk_Hangmok(sDate : String; iYh : Integer; iJumin : String): String;
Var
   i, J, K : Integer;
begin
   Result := '';
   with qrytkgum do
   begin
      Active := False;
      ParamByName('In_Num_jumin').AsString   := iJumin;
      ParamByName('In_Gubn_Cha').AsInteger    := iYh;
      ParamByName('In_Dat_Gmgn').AsString    := sDate;
      Active := True;
      J := Round(Length(FieldByName('Cod_prf').AsString) / 4);
      For I := 0 To J - 1 do
         With qrypf_Hm Do
            Begin
               Close;
               ParamByName('Cod_pf').AsString := Copy(qrytkgum.FieldByName('Cod_prf').AsString, I*4 + 1, 4);
               Open;
               For K := 1 To RecordCount do
                   Begin
                      Result := Result + FieldByName('Cod_hm').AsString;
                      Next;
                   End;
               Close;
            End;
      Result := result + FieldByName('Cod_Chuga').AsString;
      Active := False;
   end;
end;

function  TfrmLD8P011.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_id').AsString      := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD8P011.UP_FieldClear(frm : TWinControl);
var i : Integer;
begin
   // 종속된 Control를 찾아서 해당 Property를 Clear
   for i := 0 to frm.ControlCount - 1 do
   begin
      if frm.Controls[i].ClassType = TQRLabel then
         TQRLabel(frm.Controls[i]).Caption := ''
   end;
end;

end.
