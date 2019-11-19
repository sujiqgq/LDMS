//==============================================================================
// 프로그램 설명 : frmLD8Q08-바코드 출력관리(분주번호)
// 시스템        : 분석관리
// 개발일자      : 2012.11.08
// 개발자        : 유동구
//==============================================================================
unit LD8Q08;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD8Q08 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Gride: TGauge;
    Btn_NamePrint: TBitBtn;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label5: TLabel;
    cmbbjjs: TComboBox;
    mskSt_date: TDateEdit;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    cmbJisa: TComboBox;
    Chk_01: TCheckBox;
    Chk_07: TCheckBox;
    Chk_05: TCheckBox;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    qryGulkwa: TQuery;
    qryHangmokList: TQuery;
    qryNo_hangmok: TQuery;
    Chk_03: TCheckBox;
    RGrp_barcode: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vDat_bunju, UV_vNum_bunju, UV_vCod_jisa, UV_vDesc_name, UV_vDat_gmgn, UV_vNum_Samp, UV_vSex, UV_vBirth,
    UV_vGubn_01, UV_vGubn_05, UV_vGubn_07, UV_vGubn_03 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    Function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD8Q08: TfrmLD8Q08;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q081;

{$R *.DFM}

procedure TfrmLD8Q08.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate);

end;

procedure TfrmLD8Q08.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;
         grdMaster.Refresh;
      end;
   end;
end;

procedure TfrmLD8Q08.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vDat_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Samp   := VarArrayCreate([0, 0], varOleStr);
      UV_vSex        := VarArrayCreate([0, 0], varOleStr);
      UV_vBirth      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_01    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_05    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_07    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_03    := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vCod_jisa,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vNum_Samp,  iLength);
      VarArrayRedim(UV_vSex,       iLength);
      VarArrayRedim(UV_vBirth,     iLength);
      VarArrayRedim(UV_vGubn_01,   iLength);
      VarArrayRedim(UV_vGubn_05,   iLength);
      VarArrayRedim(UV_vGubn_07,   iLength);
      VarArrayRedim(UV_vGubn_03,   iLength);
   end;
end;

function TfrmLD8Q08.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if cmbbjjs.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD8Q08.FormCreate(Sender: TObject);
var
   sJisa : String;
begin
   inherited;
   GP_ComboJisa(cmbJisa);
   GP_ComboJisa(cmbBjjs);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbBjjs.Enabled := True;
      sJisa := copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbBjjs.Enabled := False;
      sJisa := GV_sJisa;
   end;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, '00',  2);
   GP_ComboDisplay(cmbBjjs, sJisa, 2);

   Cmb_gubn.ItemIndex := 0;

   // 오늘일자 설정
   mskST_Date.Text := GV_sToday;

   UP_VarMemSet(0);
end;

procedure TfrmLD8Q08.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vDat_bunju[DataRow-1];
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vCod_jisa [DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := UV_vDat_gmgn[DataRow-1];
      6 : Value := UV_vNum_Samp[DataRow-1];
      7 : Value := UV_vSex[DataRow-1];
      8 : Value := UV_vBirth[DataRow-1];
      9 : Value := UV_vGubn_01[DataRow-1];
     10 : Value := UV_vGubn_05[DataRow-1];
     11 : Value := UV_vGubn_07[DataRow-1];
     12 : Value := UV_vGubn_03[DataRow-1];
   end;
end;

procedure TfrmLD8Q08.btnQueryClick(Sender: TObject);
var
   sWhere, sSelect, sOrderBy, sWhere1, sWhere2, sEtc_Hangmok_hm, sALL_HangmokList  : string;
   index, i  : integer;
   bCheck : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryGulkwa do
   begin
      // SQL문 생성
      Active := False;
      sSelect := ' SELECT distinct B.dat_bunju, B.num_bunju, ' +
                 '        P.Cod_jisa, P.dat_gmgn, P.desc_name, P.num_samp, P.num_jumin, P.cod_jangbi AS cod_jangbi_G, P.cod_hul AS cod_hul_G, P.cod_cancer AS cod_cancer_G, P.cod_chuga AS cod_chuga_G, ' +
                 '        P.Gubn_Nosin, P.Gubn_nsyh, P.Gubn_adult, P.Gubn_adyh, P.Gubn_adult, P.Gubn_agyh, P.Gubn_agab, P.Gubn_agyh, P.Gubn_life, P.Gubn_lfyh, ' +
                 '        T.cod_prf AS cod_prf_T, T.cod_chuga AS cod_chuga_T ';

      sSelect := sSelect + ' FROM Gulkwa B INNER JOIN gumgin P ON B.cod_jisa = P.cod_jisa and B.num_id = P.num_id and B.dat_gmgn = P.dat_gmgn and (';

      if Chk_01.Checked = True then sSelect := sSelect + ' B.gubn_part = ''01'' or';
      if Chk_03.Checked = True then sSelect := sSelect + ' B.gubn_part = ''03'' or';
      if Chk_05.Checked = True then sSelect := sSelect + ' B.gubn_part = ''05'' or';
      if Chk_07.Checked = True then sSelect := sSelect + ' B.gubn_part = ''07'' or';
//      if (Chk_05.Checked = True) or (Chk_07.Checked = True) then sSelect := sSelect + ' B.gubn_part = ''07'' or';

      sSelect := copy(sSelect,1,length(sSelect) - 2);
      sSelect := sSelect + ') left outer join Tkgum T ON P.Cod_jisa = T.Cod_jisa AND P.Num_jumin = T.Num_jumin AND P.Dat_gmgn = T.Dat_gmgn AND P.Gubn_Tkgm = T.Gubn_cha ';

      sWhere := ' WHERE B.cod_bjjs = ''' + copy(cmbBjjs.Text,1,2) + '''';
      sWhere := sWhere + ' and B.dat_bunju = ''' + mskSt_date.Text + '''';

      if copy(cmbJisa.Text,1,2) <> '00' then
         sWhere := sWhere + ' and B.cod_jisa = ''' + copy(cmbJisa.Text,1,2) + '''';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND P.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND P.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;

      sOrderBy := ' ORDER BY B.dat_bunju, B.num_bunju ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      Gride.Progress := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD8Q08', 'Q', 'N');
         Gride.MaxValue := qryGulkwa.RecordCount;

         while not Eof do
         begin
            Gride.Progress := Gride.Progress + 1;
            application.ProcessMessages;
            
            bCheck := False;

            UP_VarMemSet(index);

            UV_vDat_bunju[index] := qryGulkwa.FieldByName('dat_bunju').AsString;
            UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
            UV_vCod_jisa[index]  := qryGulkwa.FieldByName('cod_jisa').AsString;
            UV_vDesc_name[index] := qryGulkwa.FieldByName('desc_name').AsString;
            UV_vDat_gmgn[index]  := qryGulkwa.FieldByName('dat_gmgn').AsString;
            UV_vNum_Samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
            UV_vBirth[index]     := copy(qryGulkwa.FieldByName('num_jumin').AsString,1,6);

            //생년월일
            case StrToInt(copy(qryGulkwa.FieldByName('num_jumin').AsString,7,1)) of
               1,3,5,7 : UV_vSex[index] := 'M';
               2,4,6,8 : UV_vSex[index] := 'F';
            end;

            //항목계산...
            //------------------------------------------------------------------------
            sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOrderBy := '';
            sEtc_Hangmok_hm := ''; sALL_HangmokList := '';
            
            if Trim(qryGulkwa.FieldByName('cod_prf_T').AsString) <> '' then
            begin
               sWhere1 := Trim(qryGulkwa.FieldByName('Cod_Jangbi_G').AsString) + ''',''' +
                          Trim(qryGulkwa.FieldByName('Cod_Hul_G').AsString)    + ''',''' +
                          Trim(qryGulkwa.FieldByName('Cod_Cancer_G').AsString) + ''',''';
               For i := 1 to (length(qryGulkwa.FieldByName('cod_prf_T').AsString) div 4) do
               begin
                  if i = (length(Trim(qryGulkwa.FieldByName('cod_prf_T').AsString)) div 4) then sWhere1 := sWhere1 + copy(Trim(qryGulkwa.FieldByName('cod_prf_T').AsString), (i*4)-3, 4)
                  else                                                                          sWhere1 := sWhere1 + copy(Trim(qryGulkwa.FieldByName('cod_prf_T').AsString), (i*4)-3, 4) + ''',''';
               end;
            end
            else
            begin
               sWhere1 := Trim(qryGulkwa.FieldByName('Cod_Jangbi_G').AsString) + ''',''' +
                          Trim(qryGulkwa.FieldByName('Cod_Hul_G').AsString)    + ''',''' +
                          Trim(qryGulkwa.FieldByName('Cod_Cancer_G').AsString) + ''',''';
            end;
      
            sEtc_Hangmok_hm := qryGulkwa.FieldByName('COD_CHUGA_G').AsString;
            // 노신유형Display
            if Trim(qryGulkwa.FieldByName('Gubn_Nosin').AsString) = '1' then
               sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '1', StrToInt(copy(qryGulkwa.FieldByName('Gubn_nsyh').AsString,1,1)));
            //성인병유형Display
            if Trim(qryGulkwa.FieldByName('Gubn_adult').AsString) = '1' then
               sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '4', StrToInt(copy(qryGulkwa.FieldByName('Gubn_adyh').AsString,1,1)));
            //간염유형Display
            if Trim(qryGulkwa.FieldByName('Gubn_agab').AsString) = '1' then
               sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '5', StrToInt(copy(qryGulkwa.FieldByName('Gubn_agyh').AsString,1,1)));
            //생애전환기유형Display
            if Trim(qryGulkwa.FieldByName('Gubn_life').AsString) = '1' then
               sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '7', StrToInt(copy(qryGulkwa.FieldByName('Gubn_lfyh').AsString,1,1)));
      
            if Trim(qryGulkwa.FieldByName('COD_CHUGA_T').AsString) <> '' then
            begin
               sWhere2 := sWhere2 + ''',''';
               For i := 1 to (length(qryGulkwa.FieldByName('COD_CHUGA_T').AsString) div 4) do
               begin
                  if (i = (length(qryGulkwa.FieldByName('COD_CHUGA_T').AsString) div 4)) and (Trim(sEtc_Hangmok_hm) = '') then sWhere2 := sWhere2 + copy(qryGulkwa.FieldByName('COD_CHUGA_T').AsString, (i*4)-3, 4)
                  else                                                                                                         sWhere2 := sWhere2 + copy(qryGulkwa.FieldByName('COD_CHUGA_T').AsString, (i*4)-3, 4) + ''',''';
               end;
            end;
      
            if Trim(sEtc_Hangmok_hm) <> '' then
            begin
               sWhere2 := sWhere2 + ''',''';
               For i := 1 to (length(sEtc_Hangmok_hm) div 4) do
               begin
                  if i = (length(sEtc_Hangmok_hm) div 4) then sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4)
                  else                                        sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
               end;
            end;
      
            with qryHangmokList do
            begin
               qryHangmokList.Active := False;
      
               sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + qryGulkwa.FieldByName('dat_gmgn').AsString + '''';
               sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
               sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part  ) UNION ( ';
               sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
               sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
               sSelect := sSelect + '   AND dat_apply <= ''' + qryGulkwa.FieldByName('dat_gmgn').AsString + ''' )';
      
               sOrderBy := ' ORDER BY H.gubn_part Desc, P.cod_hm ';
      
               qryHangmokList.SQL.Clear;
               qryHangmokList.SQL.Add(sSelect + sOrderBy);
               qryHangmokList.Active := True;
      
               if qryHangmokList.RecordCount > 0 then
               begin
                  while not Eof do
                  begin
                     sALL_HangmokList := sALL_HangmokList + qryHangmokList.FieldByName('cod_hm').AsString;

                     if (chk_01.Checked = True) and
                        (qryHangmokList.FieldByName('gubn_part').AsString = '01') then
                     begin
                        UV_vGubn_01[index] := 'O';
                        bCheck := True;
                     end;

                     if (chk_05.Checked = True) and
                        ((qryHangmokList.FieldByName('gubn_part').AsString = '05') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'S007') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'S008') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'TT01') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'TT02') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'S091')) then
                     begin
                        UV_vGubn_05[index] := 'O';
                        bCheck := True;
                     end;

                     if (chk_07.Checked = True) and
                        ((qryHangmokList.FieldByName('cod_hm').AsString = 'S034') or
                         (qryHangmokList.FieldByName('cod_hm').AsString = 'S003') or
                         (qryHangmokList.FieldByName('cod_hm').AsString = 'S080') or
                         (qryHangmokList.FieldByName('cod_hm').AsString = 'S036')) then
                     begin
                        UV_vGubn_07[index] := 'O';
                        bCheck := True;
                     end;

                     if (chk_03.Checked = True) and
                        ((qryHangmokList.FieldByName('cod_hm').AsString  = 'U024') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U029') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U030') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U031') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U032') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U044') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U051') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U053') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U054') or
                         (qryHangmokList.FieldByName('cod_hm').AsString  = 'U055')) then
                     begin
                        UV_vGubn_03[index] := 'O';
                        bCheck := True;
                     end;

                     qryHangmokList.Next;
                  end;
               end;
            end;
            //------------------------------------------------------------------------
            if bCheck = True then Inc(index);

            Next;
         end;


      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD8Q08.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date);
end;

procedure TfrmLD8Q08.btnPrintClick(Sender: TObject);
begin
  inherited;

  frmLD8Q081 := TfrmLD8Q081.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD8Q081.QuickRep.Preview
  else                                frmLD8Q081.QuickRep.Print;

end;

function TfrmLD8Q08.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.
