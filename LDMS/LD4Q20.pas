//==============================================================================
// 프로그램 설명 : 특수검진 분석 비용 내역[명단]
// 시스템        : 통합검진
// 수정일자      : 2016.08.11
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD4Q20;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, Gauges, ComObj;

type
  TfrmLD4Q20 = class(TfrmSingle)
    qryBunju: TQuery;
    pnlCond: TPanel;
    Label2: TLabel;
    msksDate: TDateEdit;
    btnsDate: TSpeedButton;
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryPf_hm: TQuery;
    qryHangmok: TQuery;
    Label3: TLabel;
    Com_Part: TComboBox;
    CheckBox: TCheckBox;
    btneDate: TSpeedButton;
    mskeDate: TDateEdit;
    Label4: TLabel;
    qryDanga: TQuery;
    ChkBox_Fire: TCheckBox;
    Label5: TLabel;
    Comb_Hm: TComboBox;
    Panel2: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel3: TPanel;
    Gauge: TGauge;
    pnl_Process: TPanel;
    grdMaster: TtsGrid;
    Grd_List: TtsGrid;
    Panel5: TPanel;
    SBtn_print1: TSpeedButton;
    RBtn_preveiw1: TRadioButton;
    RBtn_print1: TRadioButton;
    SBtn_Excel1: TSpeedButton;
    Gauge2: TGauge;
    Label6: TLabel;
    Panel4: TPanel;
    SBtn_print2: TSpeedButton;
    SBtn_Excel2: TSpeedButton;
    Gauge1: TGauge;
    Label7: TLabel;
    RBtn_preveiw2: TRadioButton;
    RBtn_print2: TRadioButton;
    Btn_ListQuery: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Btn_ListQueryClick(Sender: TObject);
    procedure SBtn_print1Click(Sender: TObject);
    procedure Grd_ListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure SBtn_print2Click(Sender: TObject);
    procedure SBtn_Excel1Click(Sender: TObject);
    procedure SBtn_Excel2Click(Sender: TObject);
  private
    { Private declarations }
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010, UV_v011, UV_v012 : Variant;

    UV_v201, UV_v202, UV_v203, UV_v204, UV_v205, UV_v206, UV_v207 : Variant;
    UV_vJisa, UV_vHmName, UV_vCode, UV_vSuga, UV_vCnt : Variant;

    // SQL문
    UV_sBasicSQL : String;

    //특수검진 비용 내역 명단(1)
    //--------------------------------------------------------------------------
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Grid_Display(iCnt : integer; sYear, sNum_bunju, sCod_bjjs, sDat_bunju, sDesc_jisa, sNum_samp, sDesc_dc, sCod_dc, sDesc_name, sNum_jumin, sCod_hanmok, sChkList : string);

    function UF_Condition : Boolean;
    function UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function UF_Profile_Hm(sHangmok_Lt, sCod_prf, sTk_Cod_Chuga : string) : string;
    //==========================================================================

    //특수검진 비용 내역서(2)
    //--------------------------------------------------------------------------
    procedure UP_VarMemSet_2(iLength : Integer);
    procedure UP_VarMemSet_C(iLength : Integer);

    procedure UP_Clear(iTemp : integer);
    procedure UP_Hangmok_Cnt(iOld, iNew : integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
    //==========================================================================
  public
     UV_sCod_jisa : String;  
    { Public declarations }
  end;

var
  frmLD4Q20: TfrmLD4Q20;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q201, LD4Q202;

{$R *.DFM}

procedure TfrmLD4Q20.UP_Clear(iTemp : integer);
begin
   UV_vJisa[iTemp]   := '';
   UV_vHmName[iTemp] := '';
   UV_vCode[iTemp]   := '';
   UV_vSuga[iTemp]   := '';
   UV_vCnt[iTemp]    := 0;
end;

procedure TfrmLD4Q20.UP_VarMemSet_C(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vJisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vHmName := VarArrayCreate([0, 0], varOleStr);
      UV_vCode   := VarArrayCreate([0, 0], varOleStr);
      UV_vSuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vCnt    := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vJisa,   iLength);
      VarArrayRedim(UV_vHmName, iLength);
      VarArrayRedim(UV_vCode,   iLength);
      VarArrayRedim(UV_vSuga,   iLength);
      VarArrayRedim(UV_vCnt,    iLength);
   end;
end;

procedure TfrmLD4Q20.UP_VarMemSet_2(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_v201 := VarArrayCreate([0, 0], varOleStr);
      UV_v202 := VarArrayCreate([0, 0], varOleStr);
      UV_v203 := VarArrayCreate([0, 0], varOleStr);
      UV_v204 := VarArrayCreate([0, 0], varOleStr);
      UV_v205 := VarArrayCreate([0, 0], varOleStr);
      UV_v206 := VarArrayCreate([0, 0], varOleStr);
      UV_v207 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v201, iLength);
      VarArrayRedim(UV_v202, iLength);
      VarArrayRedim(UV_v203, iLength);
      VarArrayRedim(UV_v204, iLength);
      VarArrayRedim(UV_v205, iLength);
      VarArrayRedim(UV_v206, iLength);
      VarArrayRedim(UV_v207, iLength);
   end;
end;

procedure TfrmLD4Q20.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001, iLength);
      VarArrayRedim(UV_v002, iLength);
      VarArrayRedim(UV_v003, iLength);
      VarArrayRedim(UV_v004, iLength);
      VarArrayRedim(UV_v005, iLength);
      VarArrayRedim(UV_v006, iLength);
      VarArrayRedim(UV_v007, iLength);
      VarArrayRedim(UV_v008, iLength);
      VarArrayRedim(UV_v009, iLength);
      VarArrayRedim(UV_v010, iLength);
      VarArrayRedim(UV_v011, iLength);
      VarArrayRedim(UV_v012, iLength);
   end;
end;

function TfrmLD4Q20.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (msksDate.Text = '') or (mskeDate.Text = '') then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD4Q20.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_VarMemSet(0);
   UP_VarMemSet_2(0);
   UP_VarMemSet_C(0);

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   GP_ComboJisa(cmbCOD_JISA);

   Com_Part.ItemIndex := 8;
   Comb_Hm.ItemIndex  := 0;

   PageControl1.ActivePage := TabSheet1;
   
   // SQL문을 저장
   UV_sBasicSQL := qryBunju.SQL.Text;
end;

procedure TfrmLD4Q20.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := GF_DateFormat(UV_v002[DataRow-1]);
      3 : Value := UV_v003[DataRow-1];
      4 : Value := UV_v004[DataRow-1];
      5 : Value := UV_v005[DataRow-1];
      6 : Value := UV_v006[DataRow-1];
      7 : Value := UV_v007[DataRow-1];
      8 : Value := UV_v008[DataRow-1];
      9 : Value := UV_v009[DataRow-1];
     10 : Value := UV_v010[DataRow-1];
     11 : Value := UV_v011[DataRow-1];
     12 : Value := FormatFloat('#,##0',UV_v012[DataRow-1]);
   end;

   if grdMaster.Rows = DataRow then grdMaster.RowColor[DataRow] := $00D6E9F3;
end;

procedure TfrmLD4Q20.btnQueryClick(Sender: TObject);
var
    iCnt, iIndex, iRet : Integer;

    sWhere, yh_name, sHangmokList, sHangmokList_Cut, sCod_Hm, sChkList : String;

    vCod_chuga : Variant;

    eSum : Extended;

    bCheck, bError : boolean;
begin
   inherited;
   iIndex := 0;
   // 조회조건 Check
   if not UF_Condition then exit; 

   // Grid & Chart 초기화
   UP_VarMemSet(0);
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   Gauge.MinValue := 0;
   Gauge.Progress := 0;
   pnl_Process.Caption := '0 / 0';

   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sWhere := ' WHERE B.dat_bunju >= ''' + msksDate.Text + '''';
      sWhere := sWhere + ' AND B.dat_bunju <= ''' + mskeDate.Text + '''';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND B.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then sWhere := sWhere + ' AND G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end;

      //검사파트...
      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';

      //배치전제외...
      if CheckBox.Checked then sWhere := sWhere + ' AND T.cod_prf Not Like ''%TC77%'' ';

      //소방서만 조회...
      if ChkBox_Fire.Checked then sWhere := sWhere + ' AND (T.Cod_Prf Like ''%TCA8%'' or T.Cod_Prf Like ''%TCB2%'') ';

      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';


      sWhere := sWhere + ' ORDER BY B.cod_bjjs, B.dat_bunju, B.num_bunju';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q20', 'Q', 'N');
         Gauge.MaxValue := RecordCount;

         if qryDanga.Active = False then qryDanga.Active := True;

         // DataSet의 값을 Variant변수로 이동
         while not Eof do
         begin
            Gauge.Progress := Gauge.Progress + 1;
            pnl_Process.Caption := IntToStr(Gauge.Progress) + ' / ' + IntToStr(Gauge.MaxValue);
            application.ProcessMessages;

            //특수검진 항목리스트..
            sHangmokList := '';

            //장비코드 추출..
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_jangbi').AsString,'');
            //혈액코드 추출..
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_hul').AsString,'');
            //종양코드 추출..
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_Cancer').AsString,'');

            //추가코드 추출..---------------------------------------------------
            iRet := GF_MulToSingle(FieldByName('Cod_chuga').AsString, 4, vCod_chuga);
            for iCnt := 0 to iRet - 1 do
            begin
               qryHangmok.Active := False;
               qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[iCnt];
               qryHangmok.Active := True;
               if qryHangmok.RecordCount > 0 then
               begin
                  if (Pos(qryHangmok.FieldByName('COD_HM').AsString, sHangmokList) = 0) and
                     (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     sHangmokList := sHangmokList + qryHangmok.FieldByName('COD_HM').AsString + '^';
                  end;
               end; //end of if(qryHangmok)
            end; //end of for(Cod_chuga)

            //특검코드 추출..---------------------------------------------------
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               //특검 프로파일 추출...
               iRet := GF_MulToSingle(FieldByName('TK_cod_prf').AsString, 4, vCod_chuga);
               for iCnt := 0 to iRet - 1 do
               begin
                  sHangmokList := UF_Profile_Hm(sHangmokList, vCod_chuga[iCnt],FieldByName('TK_cod_chuga').AsString);
               end;

               //특검 추가코드 추출..
               iRet := GF_MulToSingle(FieldByName('TK_cod_chuga').AsString, 4, vCod_chuga);
               for iCnt := 0 to iRet - 1 do
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[iCnt];
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if (Pos(qryHangmok.FieldByName('COD_HM').AsString, sHangmokList) = 0) and
                        (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                     begin
                        sHangmokList := sHangmokList + qryHangmok.FieldByName('COD_HM').AsString + '^';
                     end;
                  end; //end of if(qryHangmok)
               end; //end of for(Cod_chuga)
            end;

            // 4. 노신, 성인병, 간염, 생애전환기에 대한 검사항목 추출--------------------------
            yh_name := '';
            if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            if FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            if FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(yh_name, 4, vCod_chuga);
            for iCnt := 0 to iRet - 1 do
            begin
               qryHangmok.Active := False;
               qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[iCnt];
               qryHangmok.Active := True;
               if qryHangmok.RecordCount > 0 then
               begin
                  if (Pos(qryHangmok.FieldByName('COD_HM').AsString, sHangmokList) = 0) and
                     (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     sHangmokList := sHangmokList + qryHangmok.FieldByName('COD_HM').AsString + '^';
                  end;
               end; //end of if(qryHangmok)
            end;

            //요중삼염화초산 3개 중 2개 이상일 경우에는 한개만 가격책정...
            //------------------------------------------------------------------
            sChkList := 'NNN';
            if pos('SE92',sHangmokList) > 0 then sChkList := 'Y' + copy(sChkList,2,2);
            if pos('SE93',sHangmokList) > 0 then sChkList := copy(sChkList,1,1) + 'Y' + copy(sChkList,3,1);
            if pos('SE96',sHangmokList) > 0 then sChkList := copy(sChkList,1,2) + 'Y';
            //==================================================================

            //특수검진 항목 Display...
            sHangmokList_Cut := sHangmokList;
            while pos('^',sHangmokList_Cut) > 0 do
            begin
               UP_VarMemSet(iIndex);
               sCod_Hm := copy(sHangmokList_Cut,1,pos('^',sHangmokList_Cut)-1);
               sHangmokList_Cut := copy(sHangmokList_Cut,pos('^',sHangmokList_Cut)+1,length(sHangmokList_Cut));

               //광주센터는 5개 항목만 표기...
               //---------------------------------------------------------------
               bCheck := True;
               if FieldByName('cod_jisa').AsString = '51' then
               begin
                  if (sCod_Hm = 'SE42') or (sCod_Hm = 'SE92') or (sCod_Hm = 'SE93') or (sCod_Hm = 'SE96') or (sCod_Hm = 'SEA6') then bCheck := True
                  else                                                                                                               bCheck := False;
               end;
               //===============================================================

               //선택된 항목만 조회 가능하게 변경
               if (bCheck) and (copy(Comb_Hm.Text,1,4) <> '0000') then
               begin
                  //요중삼염화초산일 경우에는 3가지 모두 표기...
                  if (copy(Comb_Hm.Text,1,4) = '9999') and ((sCod_Hm = 'SE92') or (sCod_Hm = 'SE93') or (sCod_Hm = 'SE96')) then  bCheck := True
                  //그외 항목은 선택한 항목과 동일 시 표기..
                  else if (copy(Comb_Hm.Text,1,4) = sCod_Hm) then                                                                 bCheck := True
                  //아니면 표기 금지..
                  else                                                                                                            bCheck := False;
               end;

               if bCheck then
               begin
                  UP_Grid_Display(iIndex,
                                  FieldByName('num_year').AsString,
                                  FieldByName('num_bunju').AsString,
                                  FieldByName('cod_bjjs').AsString,
                                  FieldByName('dat_bunju').AsString,
                                  FieldByName('desc_jisa').AsString,
                                  FieldByName('Num_samp').AsString,
                                  FieldByName('desc_dc').AsString,
                                  FieldByName('cod_dc').AsString,
                                  FieldByName('desc_name').AsString,
                                  copy(FieldByName('num_jumin').AsString,1,7) + '******',
                                  sCod_Hm,
                                  sChkList);
                  Inc(iIndex);
               end;
            end;

            Next;
         end;

         eSum := 0;
         for iCnt := 0 to iIndex - 1 do
         begin
            eSum := eSum + StrToFloat(UV_v012[iCnt]);
         end;

         //합계...
         //---------------------------------------------------------------------
         UP_VarMemSet(iIndex);

         UV_v001[iIndex] := '합계';
         UV_v002[iIndex] := '';
         UV_v003[iIndex] := '';
         UV_v004[iIndex] := '';
         UV_v005[iIndex] := '';
         UV_v006[iIndex] := '';
         UV_v007[iIndex] := '';
         UV_v008[iIndex] := '';
         UV_v009[iIndex] := '';
         UV_v011[iIndex] := '';
         UV_v010[iIndex] := '';
         UV_v012[iIndex] := eSum;

         Inc(iIndex);
         //=====================================================================

         qryBunju.Active := False;
         qryDanga.Active := False;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := iIndex;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

function TfrmLD4Q20.UF_Profile_Hm(sHangmok_Lt, sCod_prf, sTk_Cod_Chuga : string) : string;
var
   sHangmok : string;
begin
   sHangmok := sHangmok_Lt;

   qryPf_hm.Active := False;
   qryPf_hm.ParamByName('cod_pf').AsString := sCod_prf;
   qryPf_hm.Active := True;

   if qryPf_hm.RecordCount > 0 then
   begin
      while not qryPf_hm.Eof do
      begin
         //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
         //---------------------------------------------------------------------
         if (sCod_prf = 'TC41')             and
            (pos('JJXE',sTk_Cod_Chuga) > 0) and
            (qryPf_hm.FieldByName('COD_HM').AsString = 'SE42') then
         begin
            // 혈중카복시(SE42) Skip...
         end
         else if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, sHangmok) = 0) and
                 (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
         begin
            sHangmok := sHangmok + qryPf_hm.FieldByName('COD_HM').AsString + '^';
         end;
         //=====================================================================

         qryPf_hm.next;
      end;   // while(qryPf_hm) end;
   end;      //if(qryPf_hm) end;

   Result := sHangmok;
end;

procedure TfrmLD4Q20.UP_Grid_Display(iCnt : integer; sYear, sNum_bunju, sCod_bjjs, sDat_bunju, sDesc_jisa, sNum_samp, sDesc_dc, sCod_dc, sDesc_name, sNum_jumin, sCod_hanmok, sChkList : string);
var
   sSex : string;
begin
   sSex := '';
   case StrToInt(copy(sNum_jumin,7,1)) of
      1,3,5,7,9 : sSex := 'M';
      2,4,6,8   : sSex := 'F';
   end;

   UV_v001[iCnt] := sCod_bjjs;
   UV_v002[iCnt] := sDat_bunju;
   UV_v003[iCnt] := sNum_bunju;
   UV_v004[iCnt] := sDesc_jisa;
   UV_v005[iCnt] := sNum_samp;
   UV_v006[iCnt] := sDesc_dc;
   UV_v007[iCnt] := sCod_dc;
   UV_v008[iCnt] := sDesc_name;
   UV_v009[iCnt] := copy(sNum_jumin,1,6) + '(' + sSex + ')';
   UV_v011[iCnt] := sCod_hanmok;
   qryDanga.Filter := '';
   qryDanga.Filter := ' cod_hm = ''' + sCod_hanmok + ''' ';
   if qryDanga.RecordCount > 0 then
   begin
      UV_v010[iCnt] := qryDanga.FieldByName('desc_kor').AsString;
      if length(sYear) = 4 then UV_v012[iCnt] := qryDanga.FieldByName('검사금액_' + sYear).AsString
      else
      begin
         showMessage('☆☆[기초(년도) 미등록자]☆☆'+ #13#10 +
                     '센터     : ' + sDesc_jisa  + #13#10 +
                     '성명     : ' + sDesc_name  + #13#10 +
                     '분주일   : ' + sDat_bunju  + #13#10 +
                     '분주번호 : ' + sNum_bunju  + #13#10 +
                     '생년월일 : ' + copy(sNum_jumin,1,6) + '(' + sSex + ')');
         UV_v012[iCnt] := '-1';
      end;
   end;

   if      (sChkList = 'YYY') and ((sCod_hanmok = 'SE93') or (sCod_hanmok = 'SE96')) then UV_v012[iCnt] := '0'
   else if (sChkList = 'NYY') and  (sCod_hanmok = 'SE96') then                            UV_v012[iCnt] := '0'
   else if (sChkList = 'YNY') and  (sCod_hanmok = 'SE96') then                            UV_v012[iCnt] := '0'
   else if (sChkList = 'YYN') and  (sCod_hanmok = 'SE93') then                            UV_v012[iCnt] := '0';
end;

procedure TfrmLD4Q20.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 없을경우의 처리
   if NewRow = 0 then exit;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // Grid Focus
   grdMaster.SetFocus;
end;

procedure TfrmLD4Q20.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnsDate then GF_CalendarClick(msksDate);
   if Sender = btneDate then GF_CalendarClick(mskeDate);
end;
procedure TfrmLD4Q20.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = msksDate then UP_Click(btnsDate);
   if Sender = mskeDate then UP_Click(btneDate);
end;

function  TfrmLD4Q20.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q20.Btn_ListQueryClick(Sender: TObject);
var
   iIndex, iCnt, iRow, iSum_Hm, iSum_Suga : integer;
   bCheck : Boolean;
begin
   inherited;

   if grdMaster.Rows < 2 then
   begin
      showmessage('특수검진 분석 비용 명단을 먼저 조회 하신 후 생성하시기 바랍니다.');
      exit;
   end;

   iIndex := 0;
   UP_VarMemSet_2(0);
   UP_VarMemSet_C(0);

   Grd_List.Rows := 0;
   for iCnt := 0 to grdMaster.Rows - 2 do
   begin
      if iCnt = 0 then begin UP_VarMemSet_C(iIndex); UP_Clear(iIndex); UP_Hangmok_Cnt(iCnt, iIndex); end
      else
      begin
         //기존데이터가 있는지 체크...
         bCheck := True;

         for iRow := 0 to iIndex do
         begin
            if (UV_vJisa[iRow] = Trim(UV_v004[iCnt])) and (UV_vCode[iRow] = Trim(UV_v011[iCnt])) and (UV_vSuga[iRow] = Trim(UV_v012[iCnt])) then
            begin
               bCheck := False;

               UP_Hangmok_Cnt(iCnt, iRow);
               break;
            end;
         end;

         if bCheck then
         begin
            Inc(iIndex);

            UP_VarMemSet_C(iIndex);
            UP_Clear(iIndex);
            UP_Hangmok_Cnt(iCnt, iIndex);
         end;
      end;
   end;

   iSum_Hm := 0; iSum_Suga := 0;
   for iCnt := 0 to iIndex do
   begin
      UP_VarMemSet_2(iCnt);

      UV_v201[iCnt] := iCnt + 1;
      UV_v202[iCnt] := Trim(UV_vJisa[iCnt]);
      UV_v203[iCnt] := Trim(UV_vHmName[iCnt]);
      UV_v204[iCnt] := Trim(UV_vCode[iCnt]);
      UV_v205[iCnt] := FormatFloat('#,##0',UV_vSuga[iCnt]);
      UV_v206[iCnt] := StrToInt(UV_vCnt[iCnt]);
      UV_v207[iCnt] := StrToInt(UV_vSuga[iCnt]) * StrToInt(UV_vCnt[iCnt]);

      iSum_Hm   := iSum_Hm   + StrToInt(UV_v206[iCnt]);
      iSum_Suga := iSum_Suga + StrToInt(UV_v207[iCnt]);
   end;

   //검진센터로 정렬..
   UP_QuickSort('D', UV_v202, 0, iIndex);

   for iCnt := 0 to iIndex do UV_v201[iCnt] := iCnt + 1;


   //합계...
   //---------------------------------------------------------------------
   Inc(iIndex);
   UP_VarMemSet_2(iIndex);

   UV_v201[iIndex] := '합계';
   UV_v202[iIndex] := '';
   UV_v203[iIndex] := '';
   UV_v204[iIndex] := '';
   UV_v205[iIndex] := '';
   UV_v206[iIndex] := iSum_Hm;
   UV_v207[iIndex] := iSum_Suga;
   //=====================================================================

   Grd_List.Rows := iIndex + 1;
end;

procedure TfrmLD4Q20.UP_Hangmok_Cnt(iOld, iNew : integer);
begin
   UV_vJisa[iNew]   := Trim(UV_v004[iOld]);
   UV_vHmName[iNew] := Trim(UV_v010[iOld]);
   UV_vCode[iNew]   := Trim(UV_v011[iOld]);
   UV_vSuga[iNew]   := Trim(UV_v012[iOld]);
   UV_vCnt[iNew]    := UV_vCnt[iNew] + 1;
end;

procedure TfrmLD4Q20.SBtn_print1Click(Sender: TObject);
begin
   inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q201 := TfrmLD4Q201.Create(Self);
   if RBtn_preveiw1.Checked then frmLD4Q201.QuickRep.Preview;
   if RBtn_print1.Checked   then frmLD4Q201.QuickRep.Print;
end;

procedure TfrmLD4Q20.Grd_ListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v201[DataRow-1];
      2 : Value := UV_v202[DataRow-1];
      3 : Value := UV_v203[DataRow-1];
      4 : Value := UV_v204[DataRow-1];
      5 : Value := UV_v205[DataRow-1];
      6 : Value := FormatFloat('#,##0',UV_v206[DataRow-1]);
      7 : Value := FormatFloat('#,##0',UV_v207[DataRow-1]);
   end;

   if Grd_List.Rows = DataRow then Grd_List.RowColor[DataRow] := $00D6E9F3;
end;

procedure TfrmLD4Q20.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];

   //--------------------------------------------------------------------------
   // 내림차순
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_v201[Lo]; UV_v201[Lo] := UV_v201[Hi]; UV_v201[Hi] := T;
            T := UV_v202[Lo]; UV_v202[Lo] := UV_v202[Hi]; UV_v202[Hi] := T;
            T := UV_v203[Lo]; UV_v203[Lo] := UV_v203[Hi]; UV_v203[Hi] := T;
            T := UV_v204[Lo]; UV_v204[Lo] := UV_v204[Hi]; UV_v204[Hi] := T;
            T := UV_v205[Lo]; UV_v205[Lo] := UV_v205[Hi]; UV_v205[Hi] := T;
            T := UV_v206[Lo]; UV_v206[Lo] := UV_v206[Hi]; UV_v206[Hi] := T;
            T := UV_v207[Lo]; UV_v207[Lo] := UV_v207[Hi]; UV_v207[Hi] := T;

            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // 오름차순
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_v201[Lo]; UV_v201[Lo] := UV_v201[Hi]; UV_v201[Hi] := T;
            T := UV_v202[Lo]; UV_v202[Lo] := UV_v202[Hi]; UV_v202[Hi] := T;
            T := UV_v203[Lo]; UV_v203[Lo] := UV_v203[Hi]; UV_v203[Hi] := T;
            T := UV_v204[Lo]; UV_v204[Lo] := UV_v204[Hi]; UV_v204[Hi] := T;
            T := UV_v205[Lo]; UV_v205[Lo] := UV_v205[Hi]; UV_v205[Hi] := T;
            T := UV_v206[Lo]; UV_v206[Lo] := UV_v206[Hi]; UV_v206[Hi] := T;
            T := UV_v207[Lo]; UV_v207[Lo] := UV_v207[Hi]; UV_v207[Hi] := T;

            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD4Q20.SBtn_print2Click(Sender: TObject);
begin
   inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q202 := TfrmLD4Q202.Create(Self);
   if RBtn_preveiw2.Checked then frmLD4Q202.QuickRep.Preview;
   if RBtn_print2.Checked   then frmLD4Q202.QuickRep.Print;
end;

procedure TfrmLD4Q20.SBtn_Excel1Click(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
  inherited;
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

procedure TfrmLD4Q20.SBtn_Excel2Click(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
  inherited;
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

   Gauge2.MaxValue := Grd_List.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, Grd_List.Rows + 1, 0, Grd_List.Cols], varOleStr);

      I := 0;
      for Row := 0 to Grd_List.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to Grd_List.Cols - 1 do
               ArrV3[Row, Col] := Grd_List.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to Grd_List.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := Grd_List.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[Grd_List.Rows + 1, Grd_List.Cols]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
