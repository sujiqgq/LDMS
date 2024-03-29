//==============================================================================
// 프로그램 설명 : LDL 결과 비교 LIST
// 시스템        : 분석관리
// 수정일자      : 2015.09.16
// 수정자        : 곽윤설
// 참고사항      : 한의 재단진단검사의학팀1500734  (LD4Q56참고)
//==============================================================================
// 수정일자      : 2015.09.19
// 수정자        : 곽윤설
// 참고사항      : 분주일자 범위로 조회, FormatFloat적용
//==============================================================================
unit ld4q70;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q70 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate_From: TSpeedButton;
    qryGumgin: TQuery;
    Gride: TGauge;
    mskDate_From: TDateEdit;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryProfile: TQuery;
    RG_Order: TRadioGroup;
    RB_bunju: TRadioButton;
    RB_samp: TRadioButton;
    Label1: TLabel;
    mskDate_To: TDateEdit;
    btnBdate_To: TSpeedButton;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_bunju, UV_vNum_bunju, UV_vNum_samp, UV_vDesc_name,
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q70: TfrmLD4Q70;
  UV_fGulkwa_01,  UV_fGulkwa1_01,  UV_fGulkwa2_01,  UV_fGulkwa3_01 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;//, LD4Q561, LD4Q562;

{$R *.DFM}

procedure TfrmLD4Q70.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_v001       := VarArrayCreate([0, 0], varOleStr);
      UV_v002       := VarArrayCreate([0, 0], varOleStr);
      UV_v003       := VarArrayCreate([0, 0], varOleStr);
      UV_v004       := VarArrayCreate([0, 0], varOleStr);
      UV_v005       := VarArrayCreate([0, 0], varOleStr);
      UV_v006       := VarArrayCreate([0, 0], varOleStr);
      UV_v007       := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_v001,       iLength);
      VarArrayRedim(UV_v002,       iLength);
      VarArrayRedim(UV_v003,       iLength);
      VarArrayRedim(UV_v004,       iLength);
      VarArrayRedim(UV_v005,       iLength);
      VarArrayRedim(UV_v006,       iLength);
      VarArrayRedim(UV_v007,       iLength);
   end;
end;

function TfrmLD4Q70.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (mskDate_From.Text = '') or (mskDate_To.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q70.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
//   UP_GridSet;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q70.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자료를 화면에 조회
   case DataCol of
       1 : Value := UV_vNo[DataRow-1];
       2 : Value := UV_vDat_bunju[DataRow-1];
       3 : Value := UV_vNum_bunju[DataRow-1];
       4 : Value := UV_vNum_samp[DataRow-1];
       5 : Value := UV_vDesc_name[DataRow-1];
       6 : Value := UV_v001[DataRow-1];
       7 : Value := UV_v002[DataRow-1];
       8 : Value := UV_v003[DataRow-1];
       9 : Value := UV_v004[DataRow-1];
      10 : Value := UV_v005[DataRow-1];
      11 : Value := UV_v006[DataRow-1];
      12 : Value := UV_v007[DataRow-1];
   end;
end;

procedure TfrmLD4Q70.btnQueryClick(Sender: TObject);
var index, DataCol, I, e25, e26, e28 : Integer;
    eRslt : Real;
    sSelect, sWhere, sOrderBy, sHangmokList: String;
    chk_nosin,chk_ss : Boolean;
begin
   inherited;
    sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;
//   UP_GridSet;
   // Grid 초기화
   grdMaster.Rows := 0;
   with qryGumgin do
   begin
      // SQL문 생성
      Close;
      sSelect := ' Select A.Dat_gmgn, A.desc_name, B.cod_jisa, B.dat_bunju, B.num_id, B.num_bunju, B.DESC_GLKWA1, B.DESC_GLKWA2, B.DESC_GLKWA3,   ' +
                 '        A.num_jumin, A.num_samp, A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.gubn_hulgum, ' +
                 '        A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga    ';
      sSelect := sSelect + ' From Gumgin A with(nolock) inner join Gulkwa B with(nolock)                    ';
      sSelect := sSelect + ' On a.cod_jisa = B.cod_jisa and a.num_id = B.num_id and a.dat_gmgn = B.dat_gmgn ';

      sWhere := ' WHERE B.Dat_Bunju >= ''' + mskDate_From.Text + ''' ';
      sWhere := sWhere + 'AND  B.Dat_Bunju <= ''' + mskDate_To.Text + ''' ';
      sWhere := sWhere + 'AND  B.Gubn_Part = ''01'' ';

      if RB_bunju.Checked then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else if RB_samp.Checked then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;

      if copy(cmbJisa.Text,1,2) = '15' then
      begin
         sWhere := sWhere + ' AND (A.gubn_hulgum = ''1'' or (A.gubn_hulgum = ''3'' and A.gubn_nosin <> ''1'' and A.gubn_adult <> ''1'' and A.gubn_life <> ''1'')) ';
      end
     {else if copy(cmbJisa.Text,1,2) <> '43' then
      begin
         sWhere := sWhere + ' AND B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' AND (A.gubn_hulgum = ''2'' or (A.gubn_hulgum = ''3'' and (A.gubn_nosin = ''1'' or A.gubn_adult = ''1'' or A.gubn_life = ''1''))) ';
      end }
      else
      begin
        {sWhere := sWhere + ' and substring(A.cod_jangbi,1,2) <> ''SS''';
        sWhere := sWhere + ' and substring(A.cod_jangbi,1,2) <> ''GS''';
        sWhere := sWhere + ' and substring(A.cod_hul,1,2) <> ''SS''';
        sWhere := sWhere + ' and substring(A.cod_hul,1,2) <> ''GS'''; }

        sWhere := sWhere + ' and B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' AND ((A.gubn_nosin = ''2'') OR  (A.gubn_hulgum = ''2'') OR  (A.gubn_hulgum = ''3''))';
      end;

      if RG_Order.ItemIndex = 0 then sOrderBy := ' ORDER BY B.dat_bunju, B.num_bunju '
      else sOrderBy := ' ORDER BY  B.dat_bunju, B.cod_jisa, A.num_samp ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;


      if RecordCount > 0 then
      begin

         GP_query_log(qryGumgin, 'LD4Q70', 'Q', 'N');

         For I := 1 to RecordCount do
         begin
            Gride.Progress := I;
            sHangmokList := '';
            sHangmokList := UF_hangmokList;
            e25 := 0; e26 := 0; e28 := 0;
            chk_nosin := false;  chk_ss:=false;
         //순수채용 제외
         //=====================================================================================================================================================//
         if (qryGumgin.FieldByName('gubn_nosin').AsString = '1') or (qryGumgin.FieldByName('gubn_nosin').AsString = '2') or
            (qryGumgin.FieldByName('gubn_adult').AsString = '1') or (qryGumgin.FieldByName('gubn_adult').AsString = '2') or
            (qryGumgin.FieldByName('gubn_life').AsString = '1')  or (qryGumgin.FieldByName('gubn_life').AsString = '2') then chk_nosin := true;

         if (Copy(qryGumgin.FieldByName('cod_jangbi').AsString, 1,2) = 'SS') or (Copy(qryGumgin.FieldByName('cod_jangbi').AsString, 1,2) = 'SG') or
            (Copy(qryGumgin.FieldByName('cod_hul').AsString, 1,2) = 'SS') or (Copy(qryGumgin.FieldByName('cod_hul').AsString, 1,2) = 'SG') then chk_ss := true;
         //=====================================================================================================================================================//
         if ((chk_nosin = true) or (chk_ss = false) or ((chk_nosin = true) and (chk_ss = true)) or ((chk_nosin = false) and (chk_ss = false)) or ((chk_nosin = false) and (chk_ss = true))) then
         begin
            UV_fGulkwa_01 := '';
            UV_fGulkwa1_01 := qryGumgin.FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2_01 := qryGumgin.FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3_01 := qryGumgin.FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01, UV_fGulkwa_01);

            if ((pos('C025',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 121, 6))<>'') and
                (StrToInt(Trim(Copy(UV_fGulkwa_01, 121,  6)))>= 300)) OR
               ((pos('C028',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 139, 6))<>'') and
                (StrToInt(Trim(Copy(UV_fGulkwa_01, 139,  6)))>= 400)) then
            begin
               UP_VarMemSet(index);
               UV_vNo[index]       := IntToStr(Index+1);
               UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
               UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
               UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
               UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
               if   (pos('C025',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  121,  6))='') then
                    UV_v001[index]     := '*'
               else if Trim(Copy(UV_fGulkwa_01,  121,  6)) <> '' then
               begin
                    UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  121,  6));
                    e25 := StrToInt(Trim(Copy(UV_fGulkwa_01,  121,  6)));
               end;
               if   (pos('C026',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  127,  6))='') then
                    UV_v002[index]     := '*'
               else if Trim(Copy(UV_fGulkwa_01,  127,  6)) <> '' then
               begin
                    UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  127,  6));
                    e26 := StrToInt(Trim(Copy(UV_fGulkwa_01,  127,  6)));
               end;
               if   (pos('C027',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  133,  6))='') then
                    UV_v003[index]     := '*'
               else if Trim(Copy(UV_fGulkwa_01,  133,  6)) <> '' then
               begin
                    UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  133,  6));
               end;
               if   (pos('C028',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  139,  6))='') then
                    UV_v004[index]     := '*'
               else if Trim(Copy(UV_fGulkwa_01,  139,  6)) <> '' then
               begin
                    UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  139,  6));
                    e28 := StrToInt(Trim(Copy(UV_fGulkwa_01,  139,  6)));
               end;
               if (e25 > 0) and (e26 > 0) and (e28 > 0) then
               begin
                  eRslt := e25 - ((e28/5) + e26);
                  eRslt := eRslt * 10;
                  eRslt := trunc(eRslt + 0.5);
                  eRslt := eRslt / 10;
                  UV_v005[index] := FormatFloat('###.0', eRslt);
               end;
               if UV_v003[index] <> '' then
               begin
                  if UV_v005[index] = UV_v003[index] then UV_v006[index] := 'O'
                  else UV_v006[index] := 'X';
               end;

               if   (pos('C074',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  331,  6)) <> '') then
                    UV_v007[index]     := Trim(Copy(UV_fGulkwa_01,  331,  6));

               Inc(index);
            end;

         end;//begin of while
         next;
      end;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGumgin

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q70.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate_From then GF_CalendarClick(mskDate_From)
   else if Sender = btnBdate_To then GF_CalendarClick(mskDate_To)
   else if RB_bunju.Checked then
   begin
      mskFrom.Enabled    := True;
      mskTo.Enabled      := True;
      MEdt_SampS.Enabled := False;
      MEdt_SampE.Enabled := False;
   end
   else if RB_samp.Checked then
   begin
      mskFrom.Enabled    := False;
      mskTo.Enabled      := False;
      MEdt_SampS.Enabled := True;
      MEdt_SampE.Enabled := True;
   end;

end;

function TfrmLD4Q70.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryGumgin.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if qryGumgin.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if qryGumgin.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qryGumgin.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryGumgin.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '1', qryGumgin.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGumgin.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '4', qryGumgin.FieldByName('Gubn_adyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '7', qryGumgin.FieldByName('Gubn_lfyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '3',qryGumgin.FieldByName('Gubn_bgyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '5',qryGumgin.FieldByName('Gubn_agyh').AsInteger);

   if (qryGumgin.FieldByName('Gubn_tkgm').AsString = '1') or (qryGumgin.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGumgin.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   result := sTemp;
end;

function TfrmLD4Q70.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q70.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;
end.
