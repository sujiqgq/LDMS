//==============================================================================
// 프로그램 설명 : 간염 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2004.05.13
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 프로그램 설명 : C형 간염 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2006.04.20
// 수정자        : 김승철
//==============================================================================
// 프로그램 설명 : 간염(A형,C형)전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2010.11.03
// 수정내용      : A형 간염 전결과 비교 분주현황추가
// 수정자        : 정용수
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2010.12.03
// 수정내용      : S019 전결과 비교 분주현황추가
// 수정자        : 정용수
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2011.12.28
// 수정내용      : 이름추가
// 수정자        : 송재호
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2012.06.07
// 수정내용      : S010 추가
// 수정자        : 김희석
//==============================================================================
// 수정일자      : 2014.04.30
// 수정내용      : HAVAB,TOTAL -> SE31범위(18.0-20.0)&SE02종전 있는 경우만 조회
// 수정자        : 곽윤설
//==============================================================================
// 수정일자      : 2014.05.13    [본원-한두례]
// 수정내용      : HBsAg/Ab 조회시 => 숫자값(s091/s099)만 있을때도 조회
// 수정자        : 곽윤설
//==============================================================================

unit LD4Q09;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q09 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryGulkwa: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    chk_gulkwaD: TCheckBox;
    GroupBox1: TGroupBox;
    RBtn_SE02: TRadioButton;
    RBtn_S016: TRadioButton;
    RBtn_S019: TRadioButton;
    RBtn_S010: TRadioButton;
    RBtn_S021: TRadioButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    RBtn_S018: TRadioButton;
    RBtn_SE01: TRadioButton;
    RBtn_S011: TRadioButton;
    RBtn_S012: TRadioButton;
    RBtn_S034: TRadioButton;
    RBtn_S036: TRadioButton;
    RBtn_S018S019: TRadioButton;
    RBtn_SE01SE02: TRadioButton;
    Rbtn_S011S012: TRadioButton;
    Rbtn_S034S036: TRadioButton;
    Rbtn_S007S008: TRadioButton;
    Rbtn_SE02SE31: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Rbtn_SE31SE02: TRadioButton;
    qry_hangmok: TQuery;
    Rbtn_S033: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);    

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSBun, UV_vName, UV_vhm, UV_vPhm,UV_vhm_1,UV_vPhm_1,UV_vRack,UV_vNum_samp : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q09: TfrmLD4Q09;


implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q091;

{$R *.DFM}

procedure TfrmLD4Q09.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo      := VarArrayCreate([0, 0], varOleStr);
      UV_vSBun    := VarArrayCreate([0, 0], varOleStr);
      UV_vName    := VarArrayCreate([0, 0], varOleStr);
      UV_vhm      := VarArrayCreate([0, 0], varOleStr);
      UV_vhm_1    := VarArrayCreate([0, 0], varOleStr);
      UV_vPhm     := VarArrayCreate([0, 0], varOleStr);
      UV_vPhm_1   := VarArrayCreate([0, 0], varOleStr);
      UV_vRack    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vSBun,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vhm,   iLength);
      VarArrayRedim(UV_vPhm,  iLength);
      VarArrayRedim(UV_vhm_1,   iLength);
      VarArrayRedim(UV_vPhm_1,  iLength);
      VarArrayRedim(UV_vRack,  iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);      
   end;
end;

function TfrmLD4Q09.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q09.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
      Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD4Q09.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   if (RBtn_S018S019.Checked = true) or (RBtn_SE01SE02.checked = true) or (Rbtn_S011S012.Checked = true) or
      (Rbtn_S034S036.Checked = true) or (Rbtn_S007S008.Checked = true) or (Rbtn_SE31SE02.Checked = true)  then
   begin
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRack[DataRow-1];
      3 : Value := UV_vNum_samp[DataRow-1];
      4 : Value := UV_vSBun[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vPhm[DataRow-1];
      7 : Value := UV_vPhm_1[DataRow-1];
      8 : Value := UV_vhm[DataRow-1];
      9 : Value := UV_vhm_1[DataRow-1];
   end;
  { if (DataCol = 6) and (UV_vhm[DataRow-1] <> UV_vPhm[DataRow-1]) and (UV_vhm[DataRow-1] <> '') then
       grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
   if (DataCol = 7) and (UV_vPhm[DataRow-1] = '양성') then
      grdMaster.CellColor[DataCol, DataRow] := $0099F7F9;
   }
  end
  else
     // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRack[DataRow-1];
      3 : Value := UV_vNum_samp[DataRow-1];
      4 : Value := UV_vSBun[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vPhm[DataRow-1];
      7 : Value := UV_vhm[DataRow-1];
      8 : Value := UV_vPhm_1[DataRow-1];
      9 : Value := UV_vhm_1[DataRow-1];
   end;
end;

procedure TfrmLD4Q09.btnQueryClick(Sender: TObject);
var index, I, rowCount, iRet, iTemp, iAge  : Integer;
    sSelect, sWhere, sOrderBy : String;
    eLow, eHigh : Extended;
    sJangbi_List, sHul_List, sGubn, sMan, sUnit : String;
    temp_S099, temp_PS099, temp_S091, temp_PS091, p_gulkwa : String;
    temp_check, temp_Pcheck : Boolean;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;
      sSelect := 'select K.desc_rackno, a.num_samp,a.cod_jisa, a.num_id, a.desc_name, a.dat_gmgn, a.cod_jangbi, a.cod_hul, a.cod_chuga, a.num_jumin, b.desc_glkwa1, b.desc_glkwa2, b.desc_glkwa3, b.num_bunju From Gumgin a ';
      sSelect := sSelect + ' left outer join gulkwa b ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sSelect := sSelect + ' left  join bunju K ';
      sSelect := sSelect + ' on b.cod_bjjs = K.cod_bjjs ';
      sSelect := sSelect + ' and b.num_bunju = K.num_bunju ';
      sSelect := sSelect + ' and b.dat_bunju = K.dat_bunju ';

      sWhere := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';
      if  (RBtn_S034.Checked =True) or (RBtn_S036.Checked =True) or (Rbtn_S034S036.Checked =True) or (Rbtn_S007S008.Checked =True) then     sWhere := sWhere + ' And  b.Gubn_Part = ''' + '07' + ''' '
      else sWhere := sWhere + ' And  b.Gubn_Part = ''' + '05' + ''' ';

      sWhere := sWhere + ' And b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju <= ''' + bunjuno2.Text + '''';  
         sOrderBy := ' ORDER BY G.num_bunju ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), G.num_bunju '
      end;

      sOrderBy := ' ORDER BY b.Num_Bunju ';
      SQL.Clear;
    
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q09', 'Q', 'N');
         For I := 1 to RecordCount do
         begin
            application.ProcessMessages;
            UP_VarMemSet(index);

            // 주민번호를 이용하여 성별과 나이를 구함
            GP_JuminToAge(qryGulkwa.FieldByname('num_jumin').AsString,qryGulkwa.FieldByname('Dat_gmgn').AsString, sMan, iAge);

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

            eLow   := 0;
            eHigh  := 0;
            sUnit  := '';

            // 항목(혈액프로파일)추출..
            sJangbi_List := ''; sHul_List := '';
            if (qryGulkwa.FieldByName('cod_hul').AsString <> '') or (qryGulkwa.FieldByName('cod_jangbi').AsString <> '')then
            begin
               qryPf_hm.Active := False;
               if (qryGulkwa.FieldByName('cod_hul').AsString <> '') then
                  qryPf_hm.ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString
               else qryPf_hm.ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //추가검사항목..
            iRet := GF_MulToSingle(qryGulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            with qryPGulkwa do
            begin
               close;
               ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
               ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
               ParamByName('dat_bunju').AsString := edtBDate.text;
               if (RBtn_S034.Checked = True) or (RBtn_S036.Checked = True) or
                  (Rbtn_S034S036.Checked =True) or (Rbtn_S007S008.Checked =True) Then
               begin
               ParamByName('Gubn_part').AsString := '07';
               end
               else ParamByName('Gubn_part').AsString := '05';


               open;

               if RecordCount >= 0 then
               /////////////////////////////////////////////////////////////////
               begin
                 if chk_gulkwaD.Checked = true then
                 begin
                 if (RBtn_s016.Checked) and (pos('S016', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,7,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S021.Checked) and (pos('S021', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S019.Checked) and (pos('S019', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S018.Checked) and (pos('S018', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_SE02.Checked) and (pos('SE02', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,53,6));       //현
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6));      //전
                           Inc(index);
                           Gride.Progress := I;

                  end
                  else if (RBtn_SE01.Checked) and (pos('SE01', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,47,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,47,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,47,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,47,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S010.Checked) and (pos('S010', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S011.Checked) and (pos('S011', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S012.Checked) and (pos('S012', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S034.Checked) and (pos('S034', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S036.Checked) and (pos('S036', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_S033.Checked) and (pos('S033', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,57,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,57,6))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,57,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,57,6));
                           Inc(index);
                           Gride.Progress := I;
                  end
                  ///////////
                   //////////////
                  else if (RBtn_S018S019.Checked) and
                           (((pos('S018', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6)))) or
                           ((pos('S019', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6))))) then
                           begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                           UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                           UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6));

                           Inc(index);
                           Gride.Progress := I;
                  end

                  ////////////////////////////////
                  else if (RBtn_SE01SE02.Checked) and
                   (((pos('SE01', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,47,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,47,6)))) or
                   ((pos('SE02', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6))))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;


                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;


                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,47,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,47,6));

                           UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,53,6));
                           UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,53,6));

                           Inc(index);
                           Gride.Progress := I;
                  end
                  ////////////////////////////////
                  else if (RBtn_S011S012.Checked) and
                  (((pos('S011', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6)))) or
                  ((pos('S012', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))))) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;


                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;


                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6));

                           UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                           UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));

                           Inc(index);
                           Gride.Progress := I;
                  end
                  ////////////////////////////////
                  else if (RBtn_S034S036.Checked) and
                  (((pos('S034', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6)))) or
                   ((pos('S036', sHul_List) > 0) and (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6)) <> trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6))))) Then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;


                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;


                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));

                           UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                           UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6));

                           Inc(index);
                           Gride.Progress := I;
                  end
                  ////////////////////////////////
                  else if Rbtn_SE31SE02.Checked then
                  begin
                        qry_hangmok.Active := False;
                        qry_hangmok.ParamByName('cod_hm').AsString    := 'SE31';
                        qry_hangmok.ParamByName('dat_gmgn').AsString := qryGulkwa.FieldByName('dat_gmgn').AsString;
                        qry_hangmok.Active := True;
                        if qry_hangmok.RecordCount > 0 then
                        begin
                            eLow   := qry_hangmok.FieldByName('IMS_LOW'   + sGubn).AsFloat;
                            eHigh  := qry_hangmok.FieldByName('IMS_HIGH'  + sGubn).AsFloat;
                        end;

                        qry_hangmok.Active := False;
                        qry_hangmok.ParamByName('cod_hm').AsString    := 'SE02';
                        qry_hangmok.ParamByName('dat_gmgn').AsString := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                        qry_hangmok.Active := True;
                        if qry_hangmok.RecordCount > 0 then
                        begin
                            sUnit  := qry_hangmok.FIeldByName('Desc_unit').AsString;
                        end;
                        if(((pos('SE31', sHul_List) > 0) AND   //조건1
                            ((trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <> '' ) and
                             (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6))) >= eLow ) and
                             (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6))) <= eHigh )
                            )
                           ) AND
                           (((trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> '' ) and
                             (trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> sUnit ) //양성, 약양성
                            )
                           )
                          ) OR
                          (((pos('SE31', sHul_List) > 0) and  //조건2
                            ((trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <> '' )and
                             ((StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) ) < eLow )or
                              (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) ) > eHigh )
                             )
                            )
                           ) AND
                           (((trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> '' )and
                             (trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) = sUnit ) //음성
                            )
                           )
                          )then
                        begin
                            UV_vNo[index]    := intTostr(index+1);
                            UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                            UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                            UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                            UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;

                            UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                            if trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <> '' then
                            begin
                            UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                            end
                            else if trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> '' then
                            begin
                               UV_vhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6));
                            end;

                            Inc(index);
                            Gride.Progress := I;
                        end;
                  end;

                 end
                 else if chk_gulkwaD.Checked = false then
                 begin ///////////////////////////////    20140513~ 곽윤설
                    temp_S099 := ''; temp_PS099 := ''; temp_S091 := ''; temp_PS091 := ''; p_gulkwa :=''; temp_check := false; temp_Pcheck := false;
                    if (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString, 29,6)) <> '') then
                    begin
                      if      (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) <= 1 ) then temp_S099 := '음성'
                      else if (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) > 1 ) and
                              (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) < 50 ) then temp_S099 := '약양성'
                      else if (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) >= 50) then temp_S099 := '양성';
                    end;
                    if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '') then
                    begin
                      if      (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) <= 1 ) then temp_PS099 := '음성'
                      else if (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) > 1 ) and
                              (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) < 50 ) then temp_PS099 := '약양성'
                      else if (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) >= 50) then temp_PS099 := '양성';
                    end;

                    if (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString, 181,6)) <> '') then
                    begin
                      if      (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) <= 10 ) then temp_S091 := '음성'
                      else if (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) > 10 ) and
                              (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) < 30 ) then temp_S091 := '약양성'
                      else if (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) >= 30) then temp_S091 := '양성';
                    end;
                    if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '') then
                    begin
                      if      (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) <= 10 ) then temp_PS091 := '음성'
                      else if (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) > 10 ) and
                              (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) < 30 ) then temp_PS091 := '약양성'
                      else if (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) >= 30) then temp_PS091 := '양성';
                    end;

                    if (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> '') then temp_check := true;
                    while not qryPGulkwa.Eof do
                    begin
                    if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> '') and (temp_Pcheck = false) then
                     begin
                      temp_Pcheck := true;
                      p_gulkwa := (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)));
                     end;
                     qryPGulkwa.next;
                    end; //20161008 수지

                    if (RBtn_S007S008.Checked) AND
                       ((pos('S007', sHul_List) > 0) or (pos('S099', sHul_List) > 0) or (pos('S091', sHul_List) > 0) or (pos('S008', sHul_List) > 0)) then
                     begin
                      qryPGulkwa.First;
                      if  ((temp_check = true) and (temp_Pcheck = true) and
                           (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> P_gulkwa)) OR
                          ((temp_PS099 <> '') and (temp_check = true) and
                           (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> temp_PS099)) OR
                          ((temp_S099 <> '') and (temp_Pcheck = true) and (temp_S099 <>  p_gulkwa)) OR
                          ((temp_S099 <> '') and (temp_PS099 <> '') and (temp_S099 <> temp_PS099)) then
                       begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                           UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;

                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) +' '+ trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6));
                          qryPGulkwa.First;
                           while not qryPGulkwa.Eof do
                           begin
                           if ((trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> '') or (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,29,6)) <> ''))
                               and (UV_vPhm[index] = '') and ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> p_gulkwa)) then
                             begin
                              UV_vPhm[index] := p_gulkwa +' '+ trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) + ('('+ (qryPGulkwa.fieldByName('dat_bunju').AsString) + ')')  ;
                             end;
                             qryPGulkwa.next;
                           end;

                           UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6))+' '+ trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6));
                           qryPGulkwa.First;
                           while not qryPGulkwa.Eof do
                           begin
                            if ((trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) <> '') or (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> ''))
                               and (UV_vPhm_1[index] = '') then
                             begin
                              UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6))+' '+ trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6))
                                                  + ('('+ (qryPGulkwa.fieldByName('dat_bunju').AsString) + ')') ;
                             end;
                            qryPGulkwa.next;
                           end;

                           Inc(index);
                           Gride.Progress := I;
                       end;
                    end   //~20140513 곽윤설

                    else if (RBtn_S016.Checked) and (pos('S016', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S021.Checked) and (pos('S021', sHul_List) > 0) then
                    begin
                             qryPGulkwa.First;
                             UV_vPhm[index] := '';
                             //최근 결과 중 값이 공란이 아닌것만 끌고 오기 ..20170418

                             while not qryPGulkwa.Eof do
                             begin
                              if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6)) <> '') and (UV_vPhm[index] = '') then
                              begin
                                 UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                              end;
                             qryPGulkwa.next;
                             end;

                             if (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6))='음성') and
                                (UV_vPhm[index]='양성') then
                             begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));

                          //   UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                             Inc(index);
                             Gride.Progress := I;
                             end;
                    end
                    else if (RBtn_S019.Checked) and (pos('S019', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S018.Checked) and (pos('S018', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_SE02.Checked) and (pos('SE02', sHul_List) > 0) then
                    begin
                             if  (trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,53,6))='음성') and
                                 (trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) ='양성') then
                             begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,53,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6));
                             Inc(index);
                             Gride.Progress := I;
                             end;
                    end
                    else if (RBtn_SE01.Checked) and (pos('SE01', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,47,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,47,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S010.Checked) and (pos('S010', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S011.Checked) and (pos('S011', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S012.Checked) and (pos('S012', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S034.Checked) and (pos('S034', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S036.Checked) and (pos('S036', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_S033.Checked) and (pos('S033', sHul_List) > 0) then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,57,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,57,6));
                             Inc(index);
                             Gride.Progress := I;
                    end
                    /////////////////////////////////////////////////////////
                     //////////////
                    else if (RBtn_S018S019.Checked) and ((pos('S018', sHul_List) > 0) or (pos('S019', sHul_List) > 0)) Then
                             begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;


                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,73,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,73,6));

                             UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,79,6));
                             UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,79,6));

                             Inc(index);
                             Gride.Progress := I;
                    end

                    ////////////////////////////////
                    else if (RBtn_SE01SE02.Checked) and ((pos('SE01', sHul_List) > 0) or (pos('SE02', sHul_List) > 0)) Then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;

                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;


                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,47,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,47,6));

                             UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,53,6));
                             UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,53,6));

                             Inc(index);
                             Gride.Progress := I;
                    end
                    ////////////////////////////////
                    else if (RBtn_S011S012.Checked) and ((pos('S011', sHul_List) > 0) or (pos('S012', sHul_List) > 0)) Then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                             //최근 결과 중 값이 공란이 아닌것만 끌고 오기 ..20161007 수지(S011, S012)
                             while not qryPGulkwa.Eof do
                             begin
                              if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) <> '') and (UV_vPhm[index] = '') then
                              begin
                                 UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) + ('('+ (qryPGulkwa.fieldByName('dat_bunju').AsString) + ')');
                              end;
                             qryPGulkwa.next;
                             end;

                             UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));

                             qryPGulkwa.First;

                             while not qryPGulkwa.Eof do
                             BEGIN
                              if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> '') and (UV_vPhm_1[index] = '') then
                              begin
                               UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) + ('('+ (qryPGulkwa.fieldByName('dat_bunju').AsString) + ')');
                              end;
                             qryPGulkwa.next;
                             END;

                             Inc(index);
                             Gride.Progress := I;
                    end
                    ////////////////////////////////
                    else if (RBtn_S034S036.Checked) and ((pos('S034', sHul_List) > 0) or (pos('S036', sHul_List) > 0)) Then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;

                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,43,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,43,6));

                             UV_vhm_1[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,55,6));
                             UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6));

                             Inc(index);
                             Gride.Progress := I;
                    end
                    else if (RBtn_SE02SE31.Checked) and (pos('SE31', sHul_List) > 0) Then
                    begin

                         if ((trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <> '') and
                             ((StrToFloat(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) >= 18.0) and
                              (StrToFloat(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <= 20.0))) AND
                            (Copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)<>'')  then
                         begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;

                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                             UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)); //HAV lgM(전결과)

                             Inc(index);
                             Gride.Progress := I;
                         end;
                    end
                    else if (RBtn_SE31SE02.Checked) and (pos('SE31', sHul_List) > 0) Then
                    begin
                             UV_vNo[index]    := intTostr(index+1);
                             UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                             UV_vRack[index]  := qryGulkwa.FieldByName('desc_rackno').AsString;
                             UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;

                             qryPGulkwa.First;
                             UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                             if trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6)) <> '' then
                             begin
                             UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                             end;

                             if trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6)) <> '' then
                             begin
                                UV_vPhm_1[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,53,6));
                             end;

                             Inc(index);
                             Gride.Progress := I;
                    end;

                    /////////////////////////////////////
                  end;
               end;
               ///////////////////////////////////////////////////////////////
            end;
            next;
         end;
         Gride.Progress := RecordCount;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGulkwa

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q09.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);

   if Sender = RBtn_S016 then
   Begin
         grdMaster.Col[6].Heading := 'HCV Ab(전)';
         grdMaster.Col[7].Heading := 'HCV Ab(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S021 then
   Begin
         grdMaster.Col[6].Heading := 'H.pylori(전)';
         grdMaster.Col[7].Heading := 'H.pylori(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S019 then
   Begin
         grdMaster.Col[6].Heading := 'Rubella IgG(전)';
         grdMaster.Col[7].Heading := 'Rubella IgG(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';

   end
   else if Sender = RBtn_S018 then
   Begin
         grdMaster.Col[6].Heading := 'Rubella IgM(전)';
         grdMaster.Col[7].Heading := 'Rubella IgM(현)';;
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';

   end
   else if Sender = RBtn_SE02 then
   Begin
         grdMaster.Col[6].Heading := 'HAV IgG(전)';
         grdMaster.Col[7].Heading := 'HAV IgG(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_SE01 then
   Begin
         grdMaster.Col[6].Heading := 'HAV IgM(전)';
         grdMaster.Col[7].Heading := 'HAV IgM(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S010 then
   Begin
         grdMaster.Col[6].Heading := 'HBc Ab(전)';
         grdMaster.Col[7].Heading := 'HBc Ab(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S011 then
   Begin
         grdMaster.Col[6].Heading := 'HBe Ag(전)';
         grdMaster.Col[7].Heading := 'HBe Ag(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S012 then
   Begin
         grdMaster.Col[6].Heading := 'HBe Ab(전)';
         grdMaster.Col[7].Heading := 'HBe Ab(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S034    then
   Begin
         grdMaster.Col[6].Heading := 'RPR 정밀(전)';
         grdMaster.Col[7].Heading := 'RPR 정밀(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';
   end
   else if Sender = RBtn_S036    then
   Begin
         grdMaster.Col[6].Heading := 'TPLA 정밀(전)';
         grdMaster.Col[7].Heading := 'TPLA 정밀(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';

   end
   else if Sender = RBtn_S033    then
   Begin
         grdMaster.Col[6].Heading := 'anti-HCV(전)';
         grdMaster.Col[7].Heading := 'anti-HCV(현)';
         grdMaster.Col[8].Heading := '';
         grdMaster.Col[9].Heading := '';

   end
   else if Sender = RBtn_S018S019    then
   Begin
         grdMaster.Col[6].Heading := 'Rubella IgM(전)';
         grdMaster.Col[7].Heading := 'Rubella igG(전)';
         grdMaster.Col[8].Heading := 'Rubella IgM(현)';
         grdMaster.Col[9].Heading := 'Rubella igG(현)';
   end
   else if Sender = RBtn_SE01SE02    then
   Begin
         grdMaster.Col[6].Heading := 'HAV IgM(전)';
         grdMaster.Col[7].Heading := 'HAV IgG(전)';
         grdMaster.Col[8].Heading := 'HAV IgM(현)';
         grdMaster.Col[9].Heading := 'HAV IgG(현)';
   end
   else if Sender = Rbtn_S011S012    then
   Begin
         grdMaster.Col[6].Heading := 'HBe Ag(전)';
         grdMaster.Col[7].Heading := 'HBe Ab(전)';
         grdMaster.Col[8].Heading := 'HBe Ag(현)';
         grdMaster.Col[9].Heading := 'HBe Ab(현)';
   end
   else if Sender = Rbtn_S034S036    then
   Begin
         grdMaster.Col[6].Heading := 'RPR 정밀(전)';
         grdMaster.Col[7].Heading := 'TPLA 정밀(전)';
         grdMaster.Col[8].Heading := 'RPR 정밀(현)';
         grdMaster.Col[9].Heading := 'TPLA 정밀(현)';

   end
   else if Sender = Rbtn_S007S008    then
   Begin
         grdMaster.Col[6].Heading := 'HBS Ag(전)';
         grdMaster.Col[7].Heading := 'HBS Ab(전)';
         grdMaster.Col[8].Heading := 'HBS Ag(현)';
         grdMaster.Col[9].Heading := 'HBS Ab(현)';

   end
   else if Sender = Rbtn_SE02SE31  then
   Begin
         grdMaster.Col[6].Heading := 'HAVAB,TOTAL(전)';
         grdMaster.Col[7].Heading := 'HAVAB,TOTAL(현)'; 
         grdMaster.Col[8].Heading := 'HAV IgG(전)';
         grdMaster.Col[9].Heading := '';

   end
   else if Sender = Rbtn_SE31SE02 then
   Begin
         grdMaster.Col[6].Heading := 'HAVAB,TOTAL(전)';
         grdMaster.Col[7].Heading := 'HAV IgG(전)';
         grdMaster.Col[8].Heading := 'HAVAB,TOTAL(현)';
         grdMaster.Col[9].Heading := '';
   end;


end;



procedure TfrmLD4Q09.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q091 := TfrmLD4Q091.Create(Self);
     frmLD4Q091.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q091 := TfrmLD4Q091.Create(Self);
     frmLD4Q091.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q09.SBut_ExcelClick(Sender: TObject);
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


procedure TfrmLD4Q09.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
    end.

