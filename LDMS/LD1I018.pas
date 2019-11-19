// 수정일자      : 2014.04.28
// 수정자        : 곽윤설
// 수정내용      : 생화학 일괄 자동 계산 기능 추가
// 참고사항      : 새로개발.....
//=============================================================================
// 수정일자      : 2014.06.09
// 수정자        : 곽윤설
// 수정내용      : BUN/Cr비율 0.0값 들어가 있을 시 공백으로 처리
// 참고사항      : 연구소, 수원, 대구를 제외한 나머지 지사 인터페이스에서 들어감
//=============================================================================
unit LD1I018;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORDialog, StdCtrls, Buttons, ExtCtrls, Grids_ts, TSGrid, Db, DBTables,
  DBCtrls, Mask, DateEdit, Gauges, math;

type
  TfrmLD1I018 = class(TfrmDialog)
    Panel36: TPanel;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    Panel1: TPanel;
    mskDate: TDateEdit;
    qryGlkwa: TQuery;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryU_Glkwa: TQuery;
    Gauge1: TGauge;
    Label1: TLabel;
    Label3: TLabel;
    cmbJisa: TComboBox;
    qryTkgum: TQuery;
    gubn_bunju: TRadioButton;
    gubn_samp: TRadioButton;
    Panel2: TPanel;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    Label2: TLabel;
    mskTo: TMaskEdit;
    Label4: TLabel;
    Edit1: TEdit;
    Panel3: TPanel;
    Label5: TLabel;
    SampFrom: TMaskEdit;
    Label6: TLabel;
    SampTo: TMaskEdit;
    Label7: TLabel;
    Edit2: TEdit;
    Label8: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure UpClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vDesc_gmgn, UV_vCod_gmgn, UV_vCod_jangbi, UV_vCod_hul, UV_vCod_cancer,
    UV_vCod_chuga, UV_vSuga_dchy, UV_vGubn_nosin, UV_vGubn_tkgm, UV_vGubn_bogun,
    UV_vGubn_adult, UV_vGubn_agab, UV_vGubn_gong, UV_vGubn_nscg, UV_vGubn_tkcg,
    UV_vGubn_bgcg, UV_vGubn_adcg, UV_vGubn_agcg, UV_vGubn_gocg,  UV_vDat_ft,
    UV_vGubn_man, UV_vNum_seq : Variant;

    // 추가항목 변수
    UV_vCod_chuga_d : Variant;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)

    function  UF_CodeCheck(sCode, sTcode : String) : Boolean;
    function  UF_Code(sHul, sCancer, sJangbi, sChuga : String) : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    function  UF_Gulkwa(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
    function  UF_Gulkwa_Diff(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
    function  UF_Tk_Hangmok(sDate: String; iYh : Integer; iJumin : String): String;

  public
    { Public declarations }
  end;

var
  frmLD1I018: TfrmLD1I018;
  UV_iRows : Integer;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I01;

{$R *.DFM}


procedure TfrmLD1I018.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbJisa.Enabled := True;
      sJisa := Copy(GV_sUserId,1,2);
   end
   else                                              
   begin                                             
      cmbJisa.Enabled := False;                      
      sJisa := GV_sJisa;                             
   end;                                              
                                                     
   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

end;

procedure TfrmLD1I018.UpClick(Sender: TObject);
var i, index, IndexTemp, iAge : Integer;
    eRslt, e03 : extended;
    e42 : string;
    sTcode, sbGulkwa, sMan, sGubn : String;
    sCheck : String;
    sSelect, sWhere, sOrderby : string;
begin

   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   index := 1; IndexTemp := 1;

   sSelect := ''; sWhere := ''; sOrderby := '';
   with qryGlkwa do
   begin
      Filter := '';
      Active := False;
      sSelect := ' SELECT A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.desc_glkwa1, A.desc_glkwa2, A.desc_glkwa3, A.gubn_part, ' +
                 '        B.num_jumin, B.desc_name, B.dat_gmgn, B.cod_dc, B.cod_hul, B.cod_cancer, B.cod_chuga, B.cod_jangbi, ' +
                 '        B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh, B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, ' +
                 '        B.gubn_gong, B.gubn_goyh, B.cod_jisa, B.gubn_tkgm, B.gubn_life, B.gubn_lfyh, B.num_samp, B.gubn_hulgum ' +
                 ' FROM   gulkwa A INNER JOIN gumgin B ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ';

      sWhere := ' WHERE A.cod_bjjs  = ''' + copy(cmbJisa.Text,1,2) + '''';

      sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';

      sWhere := sWhere + ' AND A.gubn_part = ''01''';
      sWhere := sWhere + ' AND (B.gubn_hulgum = ''1'' or B.gubn_hulgum = ''3'')';

      if      (gubn_bunju.checked = true) and (mskFrom.Text <> '')  then qryGlkwa.Filter := 'num_bunju >= ''' + mskFrom.Text + ''' AND num_bunju <= ''' + mskTo.Text + ''''
      else if (gubn_samp.checked  = true) and (SampFrom.Text <> '') then qryGlkwa.Filter := 'num_samp >= ''' + SampFrom.Text + ''' AND num_samp <= ''' + SampTo.Text + '''';

      SQL.Clear;
      SQL.Add(sSelect + sWhere);
      Active := True;

      Gauge1.MaxValue := RecordCount;
      Gauge1.Progress := 1;

      if qryGlkwa.RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I018', 'Q', 'Y');
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            if      gubn_bunju.checked = true then Edit1.text:= FieldByName('Num_Bunju').AsString
            else if gubn_samp.checked  = true then Edit2.text:= FieldByName('Num_samp').AsString;
            Application.ProcessMessages;

//---------------------------------------------------------------------------
            // 주민번호를 이용하여 성별과 나이를 구함
            GP_JuminToAge(FieldByName('Num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

            if iAge < 10 then sGubn := '1'
            else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
            else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
            else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
            else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
            else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
            else sGubn := '7';

            UV_fGulkwa := '';
            UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
            sbGulkwa := UV_fGulkwa;

            sTcode := UF_Code(FieldByName('cod_hul').AsString, FieldByName('cod_cancer').AsString, FieldByName('cod_jangbi').AsString, FieldByName('cod_chuga').AsString);

            // 4. 항목코드에 대한 검사항목 추출
            if FieldByName('gubn_nosin').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '1', FieldByName('gubn_nsyh').AsInteger)

            else if FieldByName('gubn_adult').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '4', FieldByName('gubn_adyh').AsInteger)
            else if FieldByName('gubn_adult').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger)
            else if FieldByName('gubn_gong').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '6',FieldByName('gubn_goyh').AsInteger)
            else if FieldByName('gubn_gong').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger);

            if FieldByName('gubn_bogun').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '3',FieldByName('gubn_bgyh').AsInteger)
            else if FieldByName('gubn_bogun').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '3', FieldByName('gubn_bgyh').AsInteger);
            
            if FieldByName('gubn_agab').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '5',FieldByName('gubn_agyh').AsInteger)
            else if FieldByName('gubn_agab').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger);
            
            if FieldByName('gubn_life').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '7',FieldByName('gubn_lfyh').AsInteger);

            //특검유형Display
            if FieldByName('gubn_tkgm').AsString <> '' then
               sTCode := sTCode + UF_Tk_Hangmok(FieldByName('Dat_gmgn').AsString,
                                                FieldByName('gubn_tkgm').AsInteger,
                                                FieldByName('Num_Jumin').AsString);


//실제 검사자만 일괄정리 가능하게...
//GFR코드만 남기고 나머지 전부 삭제 ..2018.04.28 수지
//---------------------------------------------------------------------------
            eRslt := 0;
            e42   := '';
            if Trim(Copy(UV_fGulkwa,  187,  6)) <> '' then
               e42 := (Trim(Copy(UV_fGulkwa, 187,  6)));


            if (Pos('C042', sTcode) > 0)  and (e42 <> '')then
            begin
               if (Pos('C087', sTcode) > 0) then
               begin
                   eRslt := 0;
                   if      (Copy(FieldByName('Num_Jumin').AsString,7,1) = '1') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '3') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '5') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '7') then
                           eRslt := 175 * Power(StrToFloat(e42),-1.154) * Power(iAge, -0.203)
                   else if (Copy(FieldByName('Num_Jumin').AsString,7,1) = '2') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '4') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '6') or
                           (Copy(FieldByName('Num_Jumin').AsString,7,1) = '8') then
                           eRslt := 175 * Power(StrToFloat(e42),-1.154) * Power(iAge, -0.203) * 0.742 ;

                  eRslt := eRslt * 10;
                  eRslt := trunc(eRslt + 0.5);
                  eRslt := eRslt / 10;
                  if Pos('.', FloatToStr(eRslt)) = 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 415, 420, (FloatToStr(eRslt)+'.0'))
                  else UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 415, 420, FloatToStr(eRslt));
               end;
            end;


            if Trim(sbGulkwa) <> Trim(UV_fGulkwa) then
            begin
               qryU_Glkwa.ParamByName('COD_BJJS').AsString   := FieldByName('cod_bjjs').AsString;
               qryU_Glkwa.ParamByName('NUM_ID').AsString     := FieldByName('num_id').AsString;
               qryU_Glkwa.ParamByName('DAT_BUNJU').AsString  := FieldByName('dat_bunju').AsString;
               qryU_Glkwa.ParamByName('NUM_BUNJU').AsString  := FieldByName('num_bunju').AsString;
               qryU_Glkwa.ParamByName('GUBN_PART').AsString  := FieldByName('gubn_part').AsString;
               UV_fGulkwa1 := '';
               UV_fGulkwa2 := '';
               UV_fGulkwa3 := '';
               GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
               qryU_Glkwa.ParamByName('DESC_GLKWA1').AsString := UV_fGulkwa1;
               qryU_Glkwa.ParamByName('DESC_GLKWA2').AsString := UV_fGulkwa2;
               qryU_Glkwa.ParamByName('DESC_GLKWA3').AsString := UV_fGulkwa3;

               qryU_Glkwa.Execsql;
               GP_query_log(qryU_Glkwa, 'LD1I011', 'Q', 'Y');
            end;
            if Index > 11     then Index := 1;
            if IndexTemp > 10 then IndexTemp := 1;

            Next;
         end; //for

         // Table과 Disconnect
         Active := False;
      end // if
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
   end; //with

end;

function TfrmLD1I018.UF_CodeCheck(sCode, sTcode : String) : Boolean;
var i, ii : Integer;
begin
   Result := False;
   ii := 1;
   for i := 1 to Round(Length(sTcode)/4) do
   begin
      ii := i * 4 - 3;
      if sCode = Copy(sTcode, ii, 4) then
      begin
         Result := True;
         Break;
      end;
   end;
end;
function TfrmLD1I018.UF_Code(sHul, sCancer, sJangbi, sChuga : String) : String;
var i, ii : Integer;
    sTcode : String;
begin
   Result := '';
   sTcode := '';
   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if sHul <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sHul;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if sCancer <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sCancer;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if sJangbi <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sJangbi;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTcode := sTcode + sChuga;

   Result := sTcode;
end;
function  TfrmLD1I018.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

function  TfrmLD1I018.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

function  TfrmLD1I018.UF_Gulkwa(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
begin
   Result := sGulkwa;
   if Trim(Copy(sGulkwa, iStart, iEnd-iStart)) = '' then
      sGulkwa := GF_RPad(Copy(sGulkwa, 1, iStart-1), iStart-1, ' ')
               + GF_RPad(sGul, iEnd - iStart + 1, ' ')
               + Copy(sGulkwa, iEnd+1, Length(sGulkwa));
   Result := sGulkwa;
end;

function  TfrmLD1I018.UF_Gulkwa_Diff(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
begin
   Result := sGulkwa;
   if Trim(Copy(sGulkwa, iStart, iEnd-iStart)) = '0.0' then
      sGulkwa := GF_RPad(Copy(sGulkwa, 1, iStart-1), iStart-1, ' ')
               + GF_RPad(sGul, iEnd - iStart + 1, ' ')
               + Copy(sGulkwa, iEnd+1, Length(sGulkwa));
   Result := sGulkwa;
end;

procedure TfrmLD1I018.btnCancelClick(Sender: TObject);
begin
  inherited;
   close;
end;

function  TfrmLD1I018.UF_Tk_Hangmok(sDate : String; iYh : Integer; iJumin : String): String;
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

end.
