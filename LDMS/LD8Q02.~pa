//==============================================================================
// 프로그램 설명 : 검사항목별분주리스트
// 시스템        : 통합검진
// 수정일자      : 2005.12.23
// 수정자        : 유동구
// 수정내용      : 특검항목 체크...
// 참고사항      : [최지혜]프로파일과 간염체크로 S007,S008중복되는 부분 수정
//==============================================================================
// 수정일자      : 2009.04.24
// 수정자        : 송재호
// 수정내용      : 혈액학 H001, H008, H011 제외.
//==============================================================================
// 수정일자      : 2010.04.15
// 수정자        : 송재호
// 수정내용      : 생화확 조회시 C022,C023,C034,C078,C080,C040,C082만 조회하는 버튼 생성
//==============================================================================
// 수정일자      : 2010.04.27
// 수정자        : 송재호
// 수정내용      : C040으로 인해 qry_gulkwa에 gubn_part2 추가하여 적용
//==============================================================================
// 수정일자      : 2011.03.15
// 수정자        : 김희석
// 수정내용      : u026  제외
//==============================================================================
// 수정일자      : 2011.04.28
// 수정자        : 송재호
// 수정내용      : 혈청학(s003,s004,s001,s071)만 순서별로 출력->composite사용못해 우선 따로 출력
//                 할수있도록 개발(차후 composite사용해서 한번에 출력토록 수정예정)
//==============================================================================
// 수정일자      : 2012.04.06
// 수정자        : 유동구
// 수정내용      : 자체검진 내용 출력으로 변경.. / 혈액학(ALL) 추가..
//==============================================================================
// 수정일자      : 2012.06.11
// 수정자        : 김희석
// 수정내용      : T6 ∩T7 , T6,T7 제외 
//==============================================================================
// 수정일자      : 2014.04.07
// 수정자        : 곽윤설
// 수정내용      : 엑셀 SBut_Excel(ComObj) 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.06.12
// 수정자        : 곽윤설
// 수정내용      : 생화학-특검 조회코드 변경
//                   기존 : C022,C023,C030,C034,C040,C056,C057,C058,C077,C078,C080,C082,C084,C085,C088,C089)
//                    -> 변경 : C022,C034,C080,C082,C033,C078,C074,GL01,GL02,GL03,GL04,GL05,GL06
// 참고사항      : [본원-문지혜]
//==============================================================================
// 수정일자      : 2015.04.02
// 수정자        : 곽윤설
// 수정내용      : 파트구분 : 08 일반혈액기타 -> 사용하지 않는 코드 삭제 요청
// 참고사항      : [본원-최정미 // 한의 재단진단검사의학팀1500269]
//==============================================================================
// 수정일자      : 2015.07.11
// 수정자        : 곽윤설
// 수정내용      : 파트구분 : 03 Urine -> 요화학  && U017,U046만 조회되도록
// 참고사항      : [본원-최정미 // 한의 재단진단검사의학팀1500269]
//==============================================================================
unit LD8Q02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD8Q02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    qry_gumgin: TQuery;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    msk_Bgmgn: TDateEdit;
    Com_Part: TComboBox;
    Label1: TLabel;
    qry_Hangmok: TQuery;
    qry_gulkwa: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    qryProfile: TQuery;
    Gride: TGauge;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    Label4: TLabel;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label6: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Ck_select: TCheckBox;
    QRCompositeReport1: TQRCompositeReport;
    btnPrint_S003: TSpeedButton;
    btnPrint_S004: TSpeedButton;
    btnPrint_S001: TSpeedButton;
    btnPrint_S071: TSpeedButton;
    Label3: TLabel;
    Ck_Name: TCheckBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    CB_Urine: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure Com_PartChange(Sender: TObject);
    procedure QRCompositeReport1AddReports(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
//    procedure QRCompositeReport1AddReports(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vCod_hangmok, UV_vDesc_hangmok, UV_vNum_bunju : Variant;
    sT006, sT007, ssT006, ssT007, sT006T007 : String;
    iRecordCount : integer;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    procedure UP_HangmokCount(iCount : integer);
    procedure UP_BunjuInsert(cHangmok, nBunju : String; nCount : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
   // procedure UP_T006T007;
  public
    { Public declarations }
  end;

var
  frmLD8Q02: TfrmLD8Q02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q021, LD8Q022, LD8Q023, LD8Q024, LD8Q025;

{$R *.DFM}

procedure TfrmLD8Q02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vCod_hangmok  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hangmok := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju    := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vCod_hangmok,  iLength);
      VarArrayRedim(UV_vDesc_hangmok, iLength);
      VarArrayRedim(UV_vNum_bunju,    iLength);
   end;
end;

function TfrmLD8Q02.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if msk_Bgmgn.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD8Q02.FormCreate(Sender: TObject);
begin
   inherited;
   sT006 := ''; sT007 := ''; ssT006 := ''; ssT007 := ''; sT006T007 := '';
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Com_Part.ItemIndex := 4;

   Cmb_gubn.ItemIndex := 0;
   msk_Bgmgn.Text := dateTostr(date);
end;

procedure TfrmLD8Q02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_hangmok[DataRow-1];
      2 : Value := UV_vDesc_hangmok[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
   end;
end;

procedure TfrmLD8Q02.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qry_Hangmok do
   begin
      // SQL문 생성
      Active := False;
      sSelect := 'select cod_hm, desc_kor, gubn_part' +
                 '  from Hangmok';
      if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
      begin
         sWhere :=  ' where gubn_part = ''01'' ';
         sWhere :=  sWhere + ' and cod_hm in (''C022'',''C034'',''C080'',''C082'',''C033'',''C078'',''C074'',''GL01'',''GL02'',''GL03'',''GL04'',''GL05'',''GL06'') ';
      end
      else if copy(Com_Part.text,1,2) = '02' then
      begin
         sWhere :=  ' where gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
         sWhere :=  sWhere + ' and cod_hm in (''H027'',''H028'',''H031'',''H036'',''H038'',''H039'')';
      end
      else if copy(Com_Part.text,1,2) = '03' then
      begin
         sWhere :=  ' where gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
         if CB_Urine.Checked then
         begin
            sWhere :=sWhere+  ' and (cod_hm in (''U017'',''U046''))'; //20150711
         end
         else
         begin
            sWhere :=sWhere+  '  and (cod_hm!=''U011'' or dat_apply!= ''19900101'')';
            sWhere :=sWhere+  '  and (cod_hm!=''U012'' or dat_apply!=''19900101'')';
            sWhere :=sWhere+  ' and (cod_hm!=''U013'' or dat_apply!=''19900101'')';
           sWhere :=  sWhere + ' and cod_hm not in (''U057'',''U001'',''U002'',''U003'',''U004'',''U005'',''U006'',''U007'',''U008'',''U009'',''U010'')';
         end;
      end
      else if copy(Com_Part.text,1,2) = '04' then
      begin
         sWhere :=  ' where (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' AND (cod_hm!=''E002'' or dat_apply!=''19900101'')) OR (cod_hm in (''TT01'',''TT02'') and dat_apply = ''20140127'')';
         //sWhere :=  sWhere + ' and (cod_hm!=''E002'' or dat_apply!=''19900101'')';
      end
      else if copy(Com_Part.text,1,2) = '08' then
      begin
         sWhere :=  ' where gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
         sWhere :=  sWhere + ' and cod_hm not in (''U026'',''E021'',''S055'',''S066'',''C036'',''E012'',''E013'',''E014'',''E016'',''T012'',''E015'',''C018'',''E010'',''S023'',''S029'',''S038'',''S040'',''S053'',''S054'',''SE38'',''T015'',''C038'',''C054'',''S075'',''C016'',''S039'',''SE16'')';
         sWhere :=  sWhere + ' and ysno_hidden = ''N'' ';
      end
      else if copy(Com_Part.text,1,2) = '10' then
      begin
         sWhere :=  ' where gubn_part = ''07''';
         sWhere :=  sWhere + ' and cod_hm in (''S003'',''S004'',''S001'',''S071'')';
      end
      else if copy(Com_Part.text,1,2) = '11' then
         sWhere :=  ' where gubn_part = ''02'''
      else if copy(Com_Part.text,1,2) = '12' then
      begin
         sWhere :=  ' where gubn_part = ''05''';
         sWhere :=  sWhere + ' and cod_hm in (''S079'',''S090'')';

      end
      else if copy(Com_Part.text,1,2) = '05' then
      begin
         sWhere :=  ' where gubn_part = ''05''';
         sWhere :=  sWhere + ' and cod_hm  not in (''C031'',''C035'',''C071'',''SE07'',''SE08'',''SE26'',''SE27'',''SE28'',''SE29'',''SE30'',''T008'',''T025'',''T029'',''T030'',''T031'',''T034'')';
      end
      else if copy(Com_Part.text,1,2) = '07' then
      begin
         sWhere :=  ' where gubn_part = ''07''';
         sWhere :=  sWhere + ' and cod_hm  not in (''EE01'',''S002'',''S004'',''S035'',''S037'',''S071'',''S080'',''S082'',''S083'',''S084'',''S093'',''S098'',''TT01'',''TT02'')';
      end
      else
         sWhere :=  ' where gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
      sOrderBy := 'group by cod_hm,desc_kor,gubn_part  ORDER BY cod_hm';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin

         // 자료의 갯수만큼 Variant 변수의 Memory 재할당
         // UP_VarMemSet(RecordCount - 1);
         // UP_VarMemSet(RecordCount + 4);
         UP_VarMemSet(RecordCount);

         // DataSet의 값을 Variant변수로 이동
         while not Eof do
         begin
            UV_vCod_hangmok[index]  := FieldByName('cod_hm').AsString;
            UV_vDesc_hangmok[index] := FieldByName('desc_kor').AsString;
            UV_vNum_bunju[index]    := '';
            Inc(index);
            Next;
         end;

         UP_HangmokCount(RecordCount - 1);
         Active := False;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   //grdMaster.Rows := index + 3;
  // if copy(Com_Part.text,1,2) = '05' then grdMaster.Rows := index + 3
  // else                                   grdMaster.Rows := index;
     grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD8Q02.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(msk_Bgmgn);
end;

procedure TfrmLD8Q02.UP_HangmokCount(iCount : integer);
var
   vSQL_Text, sSelect, sWhere, sOrderBy, sGroupBy : String;
   iRet, i, j, k, l : Integer;
   vCod_chuga : Variant;
   ysno_S007, ysno_S008,ysno_S091, sCode : String;
begin
   inherited;
   with qry_gumgin do
   begin
      // SQL문 생성
      Active := False;
      sSelect := 'select G.cod_bjjs,   G.num_bunju,  G.dat_bunju, ' +
                 '       M.desc_name,  M.dat_gmgn,   M.Cod_jangbi, M.cod_hul,   M.Cod_cancer, M.cod_chuga, M.num_samp, ' +
                 '       M.gubn_nosin, M.gubn_nsyh,  M.gubn_tkgm, M.gubn_bogun, M.gubn_bgyh, M.gubn_adult, M.gubn_adyh, M.gubn_life, M.gubn_lfyh, ' +
                 '       M.gubn_agab,  M.gubn_agyh,  M.num_jumin, M.dat_gmgn,   M.cod_jisa,  M.Gubn_hulgum, ' +
                 '       T.cod_prf,    T.cod_chuga As Tcod_chuga ' +
                 '  from gumgin M with(nolock) left outer join gulkwa G with(nolock) on G.num_id = M.num_id and G.dat_gmgn = M.dat_gmgn ' +
                 '                left outer join Tkgum T with(nolock) ON G.cod_jisa = T.cod_jisa and M.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and M.gubn_tkgm = T.gubn_cha';
      sWhere := ' where G.cod_bjjs = ''' + copy(cmbJisa.text,1,2) + '''';
      sWhere := sWhere + '   and G.dat_bunju = ''' + msk_Bgmgn.text + '''';

      if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
         sWhere := sWhere + '   and G.gubn_part = ''01'' '
      else
      begin
         if copy(Com_Part.text,1,2) = '11' then sWhere := sWhere + '   and G.gubn_part = ''02'''
         else if copy(Com_Part.text,1,2) = '10' then sWhere := sWhere + '   and G.gubn_part = ''07'''
         else if copy(Com_Part.text,1,2) = '12' then sWhere := sWhere + '   and G.gubn_part = ''05'''
         else if copy(Com_Part.text,1,2) = '04' then sWhere := sWhere + '   and (G.gubn_part = ''04'' OR G.gubn_part = ''07'')'
         else                                   sWhere := sWhere + '   and G.gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
      end;

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then sWhere := sWhere + ' AND G.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text)   <> '' then sWhere := sWhere + ' AND G.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY CAST(G.num_bunju AS INT)';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then  sWhere := sWhere + ' AND M.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(mskTo.Text)   <> '' then  sWhere := sWhere + ' AND M.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(M.num_samp AS INT) ';
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qry_gumgin, 'LD8Q02', 'Q', 'N');
         Gride.MaxValue := RecordCount;
         while not Eof do
         begin
            Gride.Progress := Gride.Progress + 1;
            ysno_S007 := 'N';
            ysno_S008 := 'N';
            ysno_S091 := 'N';            
            //---- 장비 검사(신신)..
            qry_gulkwa.Active := false;
            qry_gulkwa.ParamByName('cod_pf').AsString    :=  FieldByName('cod_Jangbi').AsString;
            qry_gulkwa.ParamByName('gubn_part').AsString :=  copy(Com_Part.text,1,2);

            if copy(Com_Part.text,1,2) = '11' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '02'
            else if copy(Com_Part.text,1,2) = '10' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
            else if copy(Com_Part.text,1,2) = '12' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '05'
            else if copy(Com_Part.text,1,2) = '04' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
            else                                   qry_gulkwa.ParamByName('gubn_part2').AsString :=  copy(Com_Part.text,1,2);

            qry_gulkwa.Active := True;
            if qry_gulkwa.RecordCount > 0 then
            begin
               sCode := '';
               while not qry_gulkwa.Eof do
               begin
                  sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                  if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                      (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                      (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                      (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                      (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                     (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                  begin
                     if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end
                  else
                  begin
                     if Ck_Name.Checked then
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                     end
                     else
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                     end;
                  end;

                  qry_gulkwa.Next;
               end;
            end;

            //---- 혈액 검사..
            qry_gulkwa.Active := false;
            qry_gulkwa.ParamByName('cod_pf').AsString    :=  FieldByName('cod_hul').AsString;
            qry_gulkwa.ParamByName('gubn_part').AsString :=  copy(Com_Part.text,1,2);
            if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
               qry_gulkwa.ParamByName('gubn_part2').AsString := '08'
            else
            begin
              if copy(Com_Part.text,1,2) = '11' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '02'
              else if copy(Com_Part.text,1,2) = '10' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
              else if copy(Com_Part.text,1,2) = '12' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '05'
              else if copy(Com_Part.text,1,2) = '04' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
              else                                   qry_gulkwa.ParamByName('gubn_part2').AsString :=  copy(Com_Part.text,1,2);
            end;

            qry_gulkwa.Active := True;
            if qry_gulkwa.RecordCount > 0 then
            begin
               while not qry_gulkwa.Eof do
               begin
                  if qry_gulkwa.FieldByname('cod_hm').AsString = 'S007' then ysno_S007 := 'Y';
                  if qry_gulkwa.FieldByname('cod_hm').AsString = 'S008' then ysno_S008 := 'Y';
                  if qry_gulkwa.FieldByname('cod_hm').AsString = 'S091' then ysno_S091 := 'Y';                  

                  sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                  if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                      (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                      (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                      (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                      (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                     (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                  begin
                     if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end
                  else
                  begin
                     if Ck_Name.Checked then
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                     end
                     else
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                     end;
                  end;

                  qry_gulkwa.Next;
               end;
            end;

            //---- Cancer 검사..
            qry_gulkwa.Active := false;
            qry_gulkwa.ParamByName('cod_pf').AsString    :=  FieldByName('cod_Cancer').AsString;
            qry_gulkwa.ParamByName('gubn_part').AsString :=  copy(Com_Part.text,1,2);
            if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
               qry_gulkwa.ParamByName('gubn_part2').AsString := '08'
            else
            begin
              if copy(Com_Part.text,1,2) = '11' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '02'
              else if copy(Com_Part.text,1,2) = '10' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
              else if copy(Com_Part.text,1,2) = '12' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '05'
              else if copy(Com_Part.text,1,2) = '04' then qry_gulkwa.ParamByName('gubn_part2').AsString :=  '07'
              else                                   qry_gulkwa.ParamByName('gubn_part2').AsString :=  copy(Com_Part.text,1,2);
            end;

            qry_gulkwa.Active := True;
            if qry_gulkwa.RecordCount > 0 then
            begin
               while not qry_gulkwa.Eof do
               begin
                  sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                  if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                      (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                      (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                      (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                      (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                     (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                  begin
                     if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end
                  else
                  begin
                     if Ck_Name.Checked then
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                     end
                     else
                     begin
                        if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                        else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                     end;
                  end;

                  qry_gulkwa.Next;
               end;
            end;

            //---- 추가 검사..
            for i := 1 to Length(FieldByName('cod_chuga').AsString) div 4 do
            begin
               qry_Hangmok.Active := false;
               vSQL_Text :=              ' select cod_hm, desc_kor From Hangmok ';
               vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(FieldByName('cod_chuga').AsString, (i * 4) - 3, 4) + '''';
               if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                  vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
               else
               begin
                  if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                  else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                  else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                  else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                  else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
               end;
               qry_Hangmok.SQL.Clear;
               qry_Hangmok.SQL.Add(vSQL_Text);
               qry_Hangmok.Active := True;

               if qry_Hangmok.RecordCount > 0 then
               begin
                  if ((qry_hangmok.FieldByName('cod_hm').AsString = 'S007') and (ysno_S007 = 'Y')) or
                     ((qry_hangmok.FieldByName('cod_hm').AsString = 'S008') and (ysno_S008 = 'Y')) or
                     ((qry_hangmok.FieldByName('cod_hm').AsString = 'S091') and (ysno_S091 = 'Y')) then
                  else
                  begin
                     if qry_hangmok.FieldByname('cod_hm').AsString = 'S007' then ysno_S007 := 'Y';
                     if qry_hangmok.FieldByname('cod_hm').AsString = 'S008' then ysno_S008 := 'Y';
                     if qry_hangmok.FieldByname('cod_hm').AsString = 'S091' then ysno_S091 := 'Y';                     

                     sCode := qry_hangmok.FieldByName('cod_hm').AsString;
                     if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                         (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                         (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                         (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                         (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                        (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                     begin
                        if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end
                     else
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end;
               end;
            end;

            // 노신유형Display
            if FieldByName('gubn_nosin').AsString = '1' then
            begin
               for i := 1 to Length(UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger)) div 4 do
               begin
                  qry_Hangmok.Active := false;
                  vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
                  vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger), (i * 4) - 3, 4) + '''';
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                     else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                     else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                     else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                     else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
                  end;

                  qry_Hangmok.SQL.Clear;
                  qry_Hangmok.SQL.Add(vSQL_Text);
                  qry_Hangmok.Active := True;

                  if qry_Hangmok.RecordCount > 0 then
                  begin
                     sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                     if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                         (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                         (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                         (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                         (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                        (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                     begin
                        if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end
                     else
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end;
               end;
            end;

            //성인병유형Display
            if FieldByName('gubn_adult').AsString = '1' then
            begin
               for i := 1 to Length(UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger)) div 4 do
               begin
                  qry_Hangmok.Active := false;
                  vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
                  vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger), (i * 4) - 3, 4) + '''';
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                     else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                     else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                     else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                     else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
                  end;


                  qry_Hangmok.SQL.Clear;
                  qry_Hangmok.SQL.Add(vSQL_Text);
                  qry_Hangmok.Active := True;

                  if qry_Hangmok.RecordCount > 0 then
                  begin
                     sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                     if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                         (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                         (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                         (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                         (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                        (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                     begin
                        if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end
                     else
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end;
               end;
            end;

            //간염유형Display
            if FieldByName('gubn_agab').AsString = '1' then
            begin
               for i := 1 to Length(UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger)) div 4 do
               begin
                  qry_Hangmok.Active := false;
                  vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
                  vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger), (i * 4) - 3, 4) + '''';
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                     else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                     else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                     else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                     else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
                  end;


                  qry_Hangmok.SQL.Clear;
                  qry_Hangmok.SQL.Add(vSQL_Text);
                  qry_Hangmok.Active := True;

                  if qry_Hangmok.RecordCount > 0 then
                  begin
                     if ((qry_hangmok.FieldByName('cod_hm').AsString = 'S007') and (ysno_S007 = 'Y')) or
                        ((qry_hangmok.FieldByName('cod_hm').AsString = 'S008') and (ysno_S008 = 'Y')) or
                        ((qry_hangmok.FieldByName('cod_hm').AsString = 'S091') and (ysno_S091 = 'Y')) then
                     else
                     begin
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S007' then ysno_S007 := 'Y';
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S008' then ysno_S008 := 'Y';
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S091' then ysno_S091 := 'Y';                        

                        sCode := qry_Hangmok.FieldByName('cod_hm').AsString;
                        if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                            (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                            (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                            (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                            (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                           (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                        begin
                           if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                           begin
                              if Ck_Name.Checked then
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                              end
                              else
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                              end;
                           end;
                        end
                        else
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end;
                  end;
               end;
            end;

            // 생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
            begin
               for i := 1 to Length(UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger)) div 4 do
               begin
                  qry_Hangmok.Active := false;
                  vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
                  vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger), (i * 4) - 3, 4) + '''';
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                     else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                     else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                     else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                     else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
                  end;


                  qry_Hangmok.SQL.Clear;
                  qry_Hangmok.SQL.Add(vSQL_Text);
                  qry_Hangmok.Active := True;

                  if qry_Hangmok.RecordCount > 0 then
                  begin
                     if ((qry_hangmok.FieldByName('cod_hm').AsString = 'S007') and (ysno_S007 = 'Y')) or
                        ((qry_hangmok.FieldByName('cod_hm').AsString = 'S008') and (ysno_S008 = 'Y')) or
                        ((qry_hangmok.FieldByName('cod_hm').AsString = 'S091') and (ysno_S091 = 'Y')) then
                     else
                     begin
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S007' then ysno_S007 := 'Y';
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S008' then ysno_S008 := 'Y';
                        if qry_hangmok.FieldByname('cod_hm').AsString = 'S091' then ysno_S091 := 'Y';                        
                        sCode := qry_Hangmok.FieldByName('cod_hm').AsString;
                        if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                            (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                            (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                            (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                            (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                           (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                        begin
                           if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                           begin
                              if Ck_Name.Checked then
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                              end
                              else
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                              end;
                           end;
                        end
                        else
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end;
                  end;
               end;
            end;

            if (FieldByName('gubn_tkgm').AsString = '1') or
               (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               //---- 특검Profile...
               iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_chuga);
               for i := 0 to iRet - 1 do
               begin
                  qry_gulkwa.Active := false;
                  qry_gulkwa.ParamByName('cod_pf').AsString    :=  vCod_chuga[i];
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     qry_gulkwa.ParamByName('gubn_part2').AsString := '08'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then qry_gulkwa.ParamByName('gubn_part2').AsString := '02'
                     else if copy(Com_Part.text,1,2) = '10' then qry_gulkwa.ParamByName('gubn_part2').AsString := '07'
                     else if copy(Com_Part.text,1,2) = '12' then qry_gulkwa.ParamByName('gubn_part2').AsString := '05'
                     else if copy(Com_Part.text,1,2) = '04' then qry_gulkwa.ParamByName('gubn_part2').AsString := '07'
                     else                                   qry_gulkwa.ParamByName('gubn_part2').AsString := copy(Com_Part.text,1,2);
                  end;

                  qry_gulkwa.Active := True;
                  if qry_gulkwa.RecordCount > 0 then
                  begin
                     while not qry_gulkwa.Eof do
                     begin
                        sCode := qry_gulkwa.FieldByName('cod_hm').AsString;
                        if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                            (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                            (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                            (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                            (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025')) and
                           (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11') then
                        begin
                           if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                           begin
                              if Ck_Name.Checked then
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                              end
                              else
                              begin
                                 if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                                 else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                              end;
                           end;
                        end
                        else
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;

                        qry_gulkwa.Next;
                     end;
                  end;
               end; // end of for[temp];

                  //---- 특검추가 검사..
               for i := 1 to Length(qry_gumgin.FieldByName('Tcod_chuga').AsString) div 4 do
               begin
                  qry_Hangmok.Active := false;
                  vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
                  vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(FieldByName('Tcod_chuga').AsString, (i * 4) - 3, 4) + '''';
                  if (copy(Com_Part.text,1,2) = '01') AND (Ck_select.Checked = True) then
                     vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''08'')'
                  else
                  begin
                     if copy(Com_Part.text,1,2) = '11' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''02'''
                     else if copy(Com_Part.text,1,2) = '10' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''07'''
                     else if copy(Com_Part.text,1,2) = '12' then vSQL_Text := vSQL_Text +  '   and gubn_part = ''05'''
                     else if copy(Com_Part.text,1,2) = '04' then vSQL_Text := vSQL_Text +  '   and (gubn_part = ''' + copy(Com_Part.text,1,2) + ''' OR gubn_part = ''07'')'
                     else                                   vSQL_Text := vSQL_Text +  '   and gubn_part = ''' + copy(Com_Part.text,1,2) + '''';
                  end;

                  qry_Hangmok.SQL.Clear;
                  qry_Hangmok.SQL.Add(vSQL_Text);
                  qry_Hangmok.Active := True;

                  if qry_Hangmok.RecordCount > 0 then
                  begin
                     sCode := qry_Hangmok.FieldByName('cod_hm').AsString;
                     if ((sCode = 'H001') or (sCode = 'H002') or (sCode = 'H003') or (sCode = 'H004') or (sCode = 'H005') or
                         (sCode = 'H006') or (sCode = 'H007') or (sCode = 'H008') or (sCode = 'H009') or (sCode = 'H010') or
                         (sCode = 'H011') or (sCode = 'H012') or (sCode = 'H013') or (sCode = 'H014') or (sCode = 'H015') or
                         (sCode = 'H016') or (sCode = 'H017') or (sCode = 'H018') or (sCode = 'H019') or (sCode = 'H020') or
                         (sCode = 'H021') or (sCode = 'H022') or (sCode = 'H023') or (sCode = 'H024') or (sCode = 'H025') and
                        (qry_gumgin.FieldByName('Gubn_hulgum').AsString = '3') and (copy(Com_Part.text,1,2) = '11')) then
                     begin
                        if qry_gumgin.FieldByName('cod_jisa').AsString = copy(GV_sUserId,1,2) then
                        begin
                           if Ck_Name.Checked then
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                           end
                           else
                           begin
                              if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                              else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                           end;
                        end;
                     end
                     else
                     begin
                        if Ck_Name.Checked then
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString + '(' + qry_gumgin.FieldByName('desc_name').AsString + ')', iCount);
                        end
                        else
                        begin
                           if Cmb_gubn.Text = '분주번호' then UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_bunju').AsString, iCount)
                           else                               UP_BunjuInsert(sCode, qry_gumgin.FieldByName('num_samp').AsString, iCount);
                        end;
                     end;
                  end;
               end;
            end;

            Next;
         end; // end of while[qry_gumgin];
      end; // end of if[qry_gumgin];

      iRecordCount := iCount;
    {  for j:= 0 to iCount do
      begin
         if UV_vCod_hangmok[j] = 'T006'      then sT006 := UV_vNum_bunju[j]
         else if UV_vCod_hangmok[j] = 'T007' then sT007 := UV_vNum_bunju[j];
      end;
      while (trim(copy(sT006,k,5)) <> '') and (trim(copy(sT007,l,5)) <> '') do
      begin
         if strToint(trim(copy(sT006,k,5))) > strToint(trim(copy(sT007,l,5))) then
         begin
            ssT007 := ssT007 + copy(sT007,l,5);
            l := l + 5;
         end
         else if strToint(trim(copy(sT006,k,5))) < strToint(trim(copy(sT007,l,5))) then
         begin
            ssT006 := ssT006 + copy(sT006,k,5);
            k := k + 5;
         end
         else if strToint(trim(copy(sT006,k,5))) = strToint(trim(copy(sT007,l,5))) then
         begin
            sT006T007 := sT006T007 + copy(sT006,k,5);
            l := l + 5;
            k := k + 5;
         end;
      end;
      ssT006 := ssT006 + copy(sT006,k,length(sT006) - k);
      ssT007 := ssT007 + copy(sT007,l,length(sT007) - l);
      UP_T006T007;}
   end; // end of with[qry_gumgin];
end;
     {
procedure TfrmLD8Q02.UP_T006T007;
begin
   if sT006T007 <> '' then
   begin
      UP_VarMemSet(iRecordCount + 4);
      UV_vCod_hangmok[iRecordCount+1]  := 'T6 ∩T7';
      UV_vDesc_hangmok[iRecordCount+1] := 'CA19-9 ∩CA125II';
      UV_vNum_bunju[iRecordCount+1]    := sT006T007;
      UV_vCod_hangmok[iRecordCount+2]  := 'T6(only)';
      UV_vDesc_hangmok[iRecordCount+2] := 'CA19-9(only)';
      UV_vNum_bunju[iRecordCount+2]    := ssT006;
      UV_vCod_hangmok[iRecordCount+3]  := 'T7(only)';
      UV_vDesc_hangmok[iRecordCount+3] := 'CA125II(only)';
      UV_vNum_bunju[iRecordCount+3]    := ssT007;
   end;
end;        }

procedure TfrmLD8Q02.UP_BunjuInsert(cHangmok, nBunju : String; nCount : Integer);
var
   vTemp : String;
   i : Integer;
begin
   for i:= 0 to nCount do
   begin
      if (UV_vCod_hangmok[i] = cHangmok) and (Trim(nBunju) <> '') then
      begin
         if pos(nBunju, UV_vNum_bunju[i]) = 0 then
         begin
            vTemp := UV_vNum_bunju[i];
            vTemp := vTemp + nBunju + ' ';
            UV_vNum_bunju[i] := vTemp;
         end;
      end;
   end;
end;

function  TfrmLD8Q02.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD8Q02.btnPrintClick(Sender: TObject);
begin
  inherited;
  if Sender = btnPrint then
  begin
     {if copy(Com_Part.Text,1,2) = '07' then
     begin
        if RBtn_preveiw.Checked = True then
        begin
           frmLD8Q021 := TfrmLD8Q021.Create(Self);
           QRCompositeReport1.Preview;
        end
        else
        begin
           frmLD8Q021 := TfrmLD8Q021.Create(Self);
           QRCompositeReport1.Print;
        end;
     end
     else
     begin }
        if RBtn_preveiw.Checked = True then
        begin
           frmLD8Q021 := TfrmLD8Q021.Create(Self);
           frmLD8Q021.QuickRep.Preview;
        end
        else
        begin
           frmLD8Q021 := TfrmLD8Q021.Create(Self);
           frmLD8Q021.QuickRep.Print;
        end;
     //end;
  end
  else if Sender = btnPrint_S003 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD8Q022 := TfrmLD8Q022.Create(Self);
        frmLD8Q022.QuickRep.Preview;
     end
     else
     begin
        frmLD8Q022 := TfrmLD8Q022.Create(Self);
        frmLD8Q022.QuickRep.Print;
     end;
  end
  else if Sender = btnPrint_S004 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD8Q023 := TfrmLD8Q023.Create(Self);
        frmLD8Q023.QuickRep.Preview;
     end
     else
     begin
        frmLD8Q023 := TfrmLD8Q023.Create(Self);
        frmLD8Q023.QuickRep.Print;
     end;
  end
  else if Sender = btnPrint_S001 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD8Q024 := TfrmLD8Q024.Create(Self);
        frmLD8Q024.QuickRep.Preview;
     end
     else
     begin
        frmLD8Q024 := TfrmLD8Q024.Create(Self);
        frmLD8Q024.QuickRep.Print;
     end;
  end
  else if Sender = btnPrint_S071 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD8Q025 := TfrmLD8Q025.Create(Self);
        frmLD8Q025.QuickRep.Preview;
     end
     else
     begin
        frmLD8Q025 := TfrmLD8Q025.Create(Self);
        frmLD8Q025.QuickRep.Print;
     end;
  end;
end;

procedure TfrmLD8Q02.Com_PartChange(Sender: TObject);
begin
  inherited;
  if Com_Part.ItemIndex = 0 then
  begin
     Ck_select.Visible := True;
     Ck_select.Checked := False;
     btnPrint_S003.Visible := False;
     btnPrint_S004.Visible := False;
     btnPrint_S001.Visible := False;
     btnPrint_S071.Visible := False;
     Label3.Visible := False;
     CB_Urine.Checked := False;
     CB_Urine.Visible := False;

  end
  else if (Com_Part.ItemIndex = 2) and (copy(GV_sUserId,1,2) = '15') then
  begin
     Ck_select.Visible := False;
     Ck_select.Checked := False;
     btnPrint_S003.Visible := False;
     btnPrint_S004.Visible := False;
     btnPrint_S001.Visible := False;
     btnPrint_S071.Visible := False;
     Label3.Visible := False;
     CB_Urine.Checked := True;
     CB_Urine.Visible := True;
  end
  else if Com_Part.ItemIndex = 6 then
  begin
     Ck_select.Visible := False;
     Ck_select.Checked := False;
     btnPrint_S003.Visible := True;
     btnPrint_S004.Visible := True;
     btnPrint_S001.Visible := True;
     btnPrint_S071.Visible := True;
     Label3.Visible := True;
     CB_Urine.Checked := False;
     CB_Urine.Visible := False;
  end
  else
  begin
     Ck_select.Visible := False;
     Ck_select.Checked := False;
     btnPrint_S003.Visible := False;
     btnPrint_S004.Visible := False;
     btnPrint_S001.Visible := False;
     btnPrint_S071.Visible := False;
     Label3.Visible := False;
     CB_Urine.Checked := False;
     CB_Urine.Visible := False;
  end;

  if Com_Part.ItemIndex = 11 then
  begin
     Ck_Name.Visible := True;
     Ck_Name.Checked := True;
  end
  else
  begin
     Ck_Name.Visible := False;
     Ck_Name.Checked := False;
  end;
end;

procedure TfrmLD8Q02.QRCompositeReport1AddReports(Sender: TObject);
begin
  inherited;
  with QRCompositeReport1 do
   begin
      Reports.Add(frmLD8Q022.QuickRep);
      Reports.Add(frmLD8Q023.QuickRep);
      Reports.Add(frmLD8Q024.QuickRep);
      Reports.Add(frmLD8Q025.QuickRep);
   end;
end;

procedure TfrmLD8Q02.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
