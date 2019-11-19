//==============================================================================
// 프로그램 설명 : 특수건강진단 분석 시료 외부 의뢰서
// 시스템        : 통합검진
// 수정일자      : 2016.09.01
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q27;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, Gauges, ComObj;

type
  TfrmLD4Q27 = class(TfrmSingle)
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
    ChkBox_Fire: TCheckBox;
    Label5: TLabel;
    Comb_Hm: TComboBox;
    Panel2: TPanel;
    Panel3: TPanel;
    Gauge: TGauge;
    pnl_Process: TPanel;
    grdMaster: TtsGrid;
    Panel5: TPanel;
    SBtn_Excel1: TSpeedButton;
    Gauge2: TGauge;
    RBtn_preveiw: TRadioButton;
    RBtn_print: TRadioButton;
    Edt_Skikwan: TEdit;
    Panel4: TPanel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Edt_Sdangdam: TEdit;
    Edt_STel: TEdit;
    Panel6: TPanel;
    Label9: TLabel;
    Edt_Ekikwan: TEdit;
    Label10: TLabel;
    Edt_EManager: TEdit;
    Edt_ETel: TEdit;
    Label11: TLabel;
    Label12: TLabel;
    Edt_Edangdam: TEdit;
    Label13: TLabel;
    DEdt_EDate: TDateEdit;
    btn_EDate: TSpeedButton;
    qryDanga: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SBtn_Excel1Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010, UV_v011 : Variant;

    // SQL문
    UV_sBasicSQL : String;

    //특수검진 비용 내역 명단(1)
    //--------------------------------------------------------------------------
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Grid_Display(iCnt : integer; sNum_bunju, sDat_bunju, sDesc_jisa, sDesc_dc, sDesc_name, sNum_jumin, sCod_hanmok : string);
    procedure UP_Clear(iTemp : integer);

    function UF_Condition : Boolean;
    function UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function UF_Profile_Hm(sHangmok_Lt, sCod_prf : string) : string;
    //==========================================================================
  public
     UV_sCod_jisa : String;
    { Public declarations }
  end;

var
  frmLD4Q27: TfrmLD4Q27;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q271;

{$R *.DFM}

procedure TfrmLD4Q27.UP_Clear(iTemp : integer);
begin
   UV_v001[iTemp] := '';
   UV_v002[iTemp] := '';
   UV_v003[iTemp] := '';
   UV_v004[iTemp] := '';
   UV_v005[iTemp] := '';
   UV_v006[iTemp] := '';
   UV_v007[iTemp] := '';
   UV_v008[iTemp] := '';
   UV_v009[iTemp] := '';
   UV_v010[iTemp] := '';
   UV_v010[iTemp] := '';
end;

procedure TfrmLD4Q27.UP_VarMemSet(iLength : Integer);
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
   end;
end;

function TfrmLD4Q27.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (msksDate.Text = '') or (mskeDate.Text = '') then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD4Q27.FormCreate(Sender: TObject);
begin
   inherited;
   // Grid의 환경설정 & Variant변수 Memory할당
   UP_VarMemSet(0);

   msksDate.Text   := GV_sToday;
   mskeDate.Text   := GV_sToday;
   DEdt_EDate.Text := GV_sToday;

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   GP_ComboJisa(cmbCOD_JISA);

   Com_Part.ItemIndex := 8;
   Comb_Hm.ItemIndex  := 0;

   // SQL문을 저장
   UV_sBasicSQL := qryBunju.SQL.Text;
end;

procedure TfrmLD4Q27.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := UV_v003[DataRow-1];
      4 : Value := GF_DateFormat(UV_v004[DataRow-1]);
      5 : if ByteType(copy(UV_v005[DataRow-1],1,2),2) = mbLeadByte then Value := copy(UV_v005[DataRow-1],1,3) + '*****'
          else                                                          Value := copy(UV_v005[DataRow-1],1,2) + '*****';
      6 : Value := UV_v006[DataRow-1];
      7 : Value := UV_v007[DataRow-1];
      8 : Value := UV_v008[DataRow-1];
      9 : Value := UV_v009[DataRow-1];
     10 : Value := UV_v010[DataRow-1];
     11 : Value := UV_v011[DataRow-1];
   end;
end;

procedure TfrmLD4Q27.btnQueryClick(Sender: TObject);
var
    iCnt, iIndex, iRet : Integer;
    sWhere, yh_name, sHangmokList, sHangmokList_Cut, sCod_Hm, sChkList : String;
    vCod_chuga : Variant;
    bCheck : boolean;
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

   //본원 박소연 선임요청
   If Copy(Comb_Hm.Text,1,4) = 'SE61' then  grdMaster.Col[10].Heading := '결과(mg/g crea.)'+ #13 +'(≤ 1.000)'
   else if Copy(Comb_Hm.Text,1,4) = 'SE85' then  grdMaster.Col[10].Heading := '결과(㎍/L)' + #13 + '(일반인기준 ≤ 3.4' + #13 +'노출기준 ≤ 200.0)'
   else  grdMaster.Col[10].Heading := '결과';

   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sWhere := ' WHERE B.dat_bunju >= ''' + msksDate.Text + '''';
      sWhere := sWhere + ' AND B.dat_bunju <= ''' + mskeDate.Text + '''';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND B.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then sWhere := sWhere + ' AND G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + ''''
         else                                 sWhere := sWhere + ' AND G.cod_jisa <> ''51'' ';
      end;

      //검사파트...
      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';

      //배치전제외...
      if CheckBox.Checked then sWhere := sWhere + ' AND T.cod_prf Not Like ''%TC77%'' ';

      //소방서만 조회...
      if ChkBox_Fire.Checked then sWhere := sWhere + ' AND (T.Cod_Prf Like ''%TCA8%'' or T.Cod_Prf Like ''%TCB2%'') ';

      sWhere := sWhere + ' ORDER BY B.cod_bjjs, B.dat_bunju, B.num_bunju';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q27', 'Q', 'N');
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
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_jangbi').AsString);
            //혈액코드 추출..
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_hul').AsString);
            //종양코드 추출..
            sHangmokList := UF_Profile_Hm(sHangmokList,FieldByName('Cod_Cancer').AsString);

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
                  sHangmokList := UF_Profile_Hm(sHangmokList, vCod_chuga[iCnt]);
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

            //특수검진 항목 Display...
            sHangmokList_Cut := sHangmokList;
            while pos('^',sHangmokList_Cut) > 0 do
            begin
               UP_VarMemSet(iIndex);
               sCod_Hm := copy(sHangmokList_Cut,1,pos('^',sHangmokList_Cut)-1);
               sHangmokList_Cut := copy(sHangmokList_Cut,pos('^',sHangmokList_Cut)+1,length(sHangmokList_Cut));

               //선택된 항목만 조회 가능하게 변경
               if (copy(Comb_Hm.Text,1,4) = sCod_Hm) then bCheck := True   //그외 항목은 선택한 항목과 동일 시 표기..
               else                                       bCheck := False; //아니면 표기 금지..

               if bCheck then
               begin
                  UP_Clear(iIndex);

                  UP_Grid_Display(iIndex,
                                  FieldByName('num_bunju').AsString,
                                  FieldByName('dat_bunju').AsString,
                                  FieldByName('desc_jisa').AsString,
                                  FieldByName('desc_dc').AsString,
                                  FieldByName('desc_name').AsString,
                                  copy(FieldByName('num_jumin').AsString,1,7) + '******',
                                  sCod_Hm);
                  Inc(iIndex);
               end;
            end;

            Next;
         end;
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

function TfrmLD4Q27.UF_Profile_Hm(sHangmok_Lt, sCod_prf : string) : string;
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
         if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, sHangmok) = 0) and
            (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
         begin
            sHangmok := sHangmok + qryPf_hm.FieldByName('COD_HM').AsString + '^';
         end;

         qryPf_hm.next;
      end;   // while(qryPf_hm) end;
   end;      //if(qryPf_hm) end;

   Result := sHangmok;
end;

procedure TfrmLD4Q27.UP_Grid_Display(iCnt : integer; sNum_bunju, sDat_bunju, sDesc_jisa, sDesc_dc, sDesc_name, sNum_jumin, sCod_hanmok : string);
var
   sSex : string;
begin
   sSex := '';
   case StrToInt(copy(sNum_jumin,7,1)) of
      1,3,5,7,9 : sSex := 'M';
      2,4,6,8   : sSex := 'F';
   end;

   qryDanga.Filter := '';
   qryDanga.Filter := ' cod_hm = ''' + sCod_hanmok + ''' ';
   if qryDanga.RecordCount > 0 then UV_v008[iCnt] := qryDanga.FieldByName('desc_kor').AsString;

   UV_v001[iCnt] := IntToStr(iCnt + 1);
   UV_v002[iCnt] := sDesc_jisa;
   UV_v003[iCnt] := sNum_bunju;
   UV_v004[iCnt] := sDat_bunju;
   UV_v005[iCnt] := sDesc_dc;
   UV_v006[iCnt] := sDesc_name;
   UV_v007[iCnt] := copy(sNum_jumin,1,6) + '(' + sSex + ')';

   UV_v009[iCnt] := '정상';
   UV_v010[iCnt] := '';
   UV_v010[iCnt] := '';
end;

procedure TfrmLD4Q27.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD4Q27.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnsDate  then GF_CalendarClick(msksDate);
   if Sender = btneDate  then GF_CalendarClick(mskeDate);
   if Sender = btn_EDate then GF_CalendarClick(DEdt_EDate);
end;
procedure TfrmLD4Q27.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = msksDate   then UP_Click(btnsDate);
   if Sender = mskeDate   then UP_Click(btneDate);
   if Sender = DEdt_EDate then UP_Click(btn_EDate);
end;

function  TfrmLD4Q27.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q27.SBtn_Excel1Click(Sender: TObject);
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

procedure TfrmLD4Q27.btnPrintClick(Sender: TObject);
begin
   inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q271 := TfrmLD4Q271.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q271.QuickRep.Preview;
   if RBtn_print.Checked   then frmLD4Q271.QuickRep.Print;
end;

end.
