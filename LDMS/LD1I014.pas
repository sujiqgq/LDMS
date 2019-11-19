//==============================================================================
// 프로그램 설명 : 검진접수(단체품의서)관리
// 시스템        : 통합검진
// 수정일자      : 1999.06.03
// 수정자        : 김은석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 검진접수(단체품의서)관리
// 시스템        : 통합검진
// 수정일자      : 2011.03.10
// 수정자        : 송재호
// 수정내용      : 전체체크/해제 버튼 추가(한의 재단혈청학팀00110)
// 참고사항      :
//==============================================================================
unit LD1I014;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORDialog, StdCtrls, Buttons, ExtCtrls, Grids_ts, TSGrid, Db, DBTables,
  DBCtrls, Mask, DateEdit, Gauges;

type
  TfrmLD1I014 = class(TfrmDialog)
    Panel36: TPanel;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    Gauge1: TGauge;
    Panel1: TPanel;
    mskDate: TDateEdit;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    qryGlkwa: TQuery;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryU_Glkwa: TQuery;
    qryPrev: TQuery;
    Label3: TLabel;
    cmbJisa: TComboBox;
    qryTkgum: TQuery;
    GroupBox1: TGroupBox;
    ckbS008_1: TCheckBox;
    GroupBox2: TGroupBox;
    ckbS001: TCheckBox;
    ckbS037: TCheckBox;
    ckbS003: TCheckBox;
    ckbS007: TCheckBox;
    ckbS034: TCheckBox;
    ckbS036: TCheckBox;
    ckbTt01: TCheckBox;
    ckbS035: TCheckBox;
    ckbS071: TCheckBox;
    ckbS004: TCheckBox;
    ckbS008: TCheckBox;
    btnALL: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure UpClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnAllClick(Sender: TObject);
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
    function  UF_Tk_Hangmok(sDate: String; iYh : Integer; iJumin : String): String;

  public
    { Public declarations }
  end;

var
  frmLD1I014: TfrmLD1I014;
  UV_iRows : Integer;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I01;

{$R *.DFM}


procedure TfrmLD1I014.FormCreate(Sender: TObject);
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

procedure TfrmLD1I014.UpClick(Sender: TObject);
var i, index : Integer;
    sTcode, sbGulkwa : String;
begin

   // 조회조건 Check
   //if not UF_Condition then exit;

   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   index := 1;
   with qryGlkwa do
   begin
      Active := False;
      ParamByName('cod_bjjs').AsString  := copy(cmbJisa.Text,1,2);
      ParamByName('dat_bunju').AsString := mskDate.Text;
      ParamByName('sbunju').AsString    := mskFrom.Text;
      ParamByName('ebunju').AsString    := mskTo.Text;
      Active := True;
      Gauge1.MaxValue := RecordCount;
      Gauge1.Progress := 1;
      if RecordCount > 0 then
      begin
        GP_query_log(qryGlkwa, 'LD1I014', 'Q', 'Y');
        for i := 0 to RecordCount - 1 do
        begin
            Gauge1.Progress := Gauge1.Progress + 1;

            UV_fGulkwa := '';
            UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
            sbGulkwa := UV_fGulkwa;
            sTcode := UF_Code(FieldByName('cod_hul').AsString, FieldByName('cod_cancer').AsString, FieldByName('cod_jangbi').AsString, FieldByName('cod_chuga').AsString);

            // 4. 항목코드에 대한 검사항목 추출
            if FieldByName('gubn_nosin').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '1', FieldByName('gubn_nsyh').AsInteger);
//          노신재검Table사용無
{            else if FieldByName('gubn_nosin').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger)
}
            if FieldByName('gubn_adult').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '4', FieldByName('gubn_adyh').AsInteger);
{            else if FieldByName('gubn_adult').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger)
}
            if FieldByName('gubn_gong').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '6',FieldByName('gubn_goyh').AsInteger);
{            else if FieldByName('gubn_gong').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger);
}
            if FieldByName('gubn_bogun').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '3',FieldByName('gubn_bgyh').AsInteger);
{            else if FieldByName('gubn_bogun').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '3', FieldByName('gubn_bgyh').AsInteger);
}
            if FieldByName('gubn_agab').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '5',FieldByName('gubn_agyh').AsInteger);
{            else if FieldByName('gubn_agab').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger);
}
            if FieldByName('gubn_life').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '7',FieldByName('gubn_lfyh').AsInteger);

      //특검유형Display
            if FieldByName('gubn_tkgm').AsString <> '' then
               sTCode := sTCode + UF_Tk_Hangmok(FieldByName('Dat_gmgn').AsString,
                                              FieldByName('gubn_tkgm').AsInteger,
                                              FieldByName('Num_Jumin').AsString);
            if ckbS001.Checked then
            begin
               if Pos('S001', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 7, 12, '음성');
            end;
            if ckbS003.Checked then
            begin
               if Pos('S003', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '음성');
            end;
            if ckbS007.Checked then
            begin
               if Pos('S007', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 31, 36, '음성');
            end;
            if ckbS008.Checked then
            begin
               if Pos('S008', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 37, 42, '음성');
            end;
            if ckbS008_1.Checked then
            begin
               if Pos('S008', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 37, 42, '양성');
            end;
            if ckbS034.Checked then
            begin
               if Pos('S034', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 43, 48, '음성');
            end;
            if ckbS035.Checked then
            begin
               if Pos('S035', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 49, 54, '음성');
            end;
            if ckbS036.Checked then
            begin
               if Pos('S036', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 55, 60, '음성');
            end;
            if ckbS037.Checked then
            begin
               if Pos('S037', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '음성');
            end;
            if ckbTt01.Checked then
            begin
               if Pos('TT01', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 85, 90, '음성');
            end;
            if ckbS071.Checked then
            begin
               if Pos('S071', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 127, 132, '음성');
            end;
            if ckbS004.Checked then
            begin
               if Pos('S004', sTcode) > 0 then
                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa,  25,  30, '음성');
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
               GP_query_log(qryU_Glkwa, 'LD1I014', 'Q', 'Y');
            end;
            Index := Index + 1;
            if Index > 5 then
               Index := 1;
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

function TfrmLD1I014.UF_CodeCheck(sCode, sTcode : String) : Boolean;
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
function TfrmLD1I014.UF_Code(sHul, sCancer, sJangbi, sChuga : String) : String;
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
function  TfrmLD1I014.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I014.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I014.UF_Gulkwa(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
begin
   Result := sGulkwa;
   if Trim(Copy(sGulkwa, iStart, iEnd-iStart)) = '' then
      sGulkwa := GF_RPad(Copy(sGulkwa, 1, iStart-1), iStart-1, ' ')
               + GF_RPad(sGul, iEnd - iStart + 1, ' ')
               + Copy(sGulkwa, iEnd+1, Length(sGulkwa));
   Result := sGulkwa;
end;

procedure TfrmLD1I014.btnCancelClick(Sender: TObject);
begin
  inherited;
   Close;
end;

function  TfrmLD1I014.UF_Tk_Hangmok(sDate : String; iYh : Integer; iJumin : String): String;
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

procedure TfrmLD1I014.btnAllClick(Sender: TObject);
begin
  inherited;

  if Sender = btnALL then
  begin
     if (ckbS003.Checked = True) AND (ckbS034.Checked = True) AND (ckbS001.Checked = True) AND (ckbS004.Checked = True) AND
        (ckbS007.Checked = True) AND (ckbS036.Checked = True) AND (ckbS037.Checked = True) AND (ckbS008.Checked = True) AND
        (ckbS035.Checked = True) AND (ckbTt01.Checked = True) AND (ckbS071.Checked = True) then
     begin
        ckbS003.Checked := False;
        ckbS034.Checked := False;
        ckbS001.Checked := False;
        ckbS004.Checked := False;
        ckbS007.Checked := False;
        ckbS036.Checked := False;
        ckbS037.Checked := False;
        ckbS008.Checked := False;
        ckbS035.Checked := False;
        ckbTt01.Checked := False;
        ckbS071.Checked := False;
        btnALL.Caption := '전체선택';
     end
     else
     begin
        ckbS003.Checked := True;
        ckbS034.Checked := True;
        ckbS001.Checked := True;
        ckbS004.Checked := True;
        ckbS007.Checked := True;
        ckbS036.Checked := True;
        ckbS037.Checked := True;
        ckbS008.Checked := True;
        ckbS035.Checked := True;
        ckbTt01.Checked := True;
        ckbS071.Checked := True;
        btnALL.Caption := '전체해제';
     end;
  end;
end;

end.
