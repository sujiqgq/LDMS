//==============================================================================
// 프로그램 설명 : 항목리스트조회(특검샘플)
// 시스템        : 통합검진
// 개발일자      : 2014.05.12
// 개발자        : 백승현
// 참고사항      :
//==============================================================================
// 개발일자      : 2014.05.19
// 개발자        : 곽윤설
// 수정사항      : 특검유형 조회 추가
// 참고사항      : [부산-유희짐]
//==============================================================================
// 개발일자      : 2014.06.13
// 개발자        : 곽윤설
// 수정사항      : 5.특검URIN 조회시 특검프로파일 TC77 제외
// 참고사항      : [수원-이청심]
//==============================================================================
// 개발일자      : 2014.10.04
// 개발자        : 곽윤설
// 수정사항      : 6.소변세포검사 - JJFG 추가
// 참고사항      : [본원-유다현]
//==============================================================================
// 개발일자      : 2015.03.06
// 개발자        : 곽윤설
// 수정사항      : 특검 프로파일 모두 읽어오도록 오류 수정
// 참고사항      : 
//==============================================================================
// 개발일자      : 2015.06.16
// 개발자        : 곽윤설
// 수정사항      : SEA6(Met Hb)추가
// 참고사항      : 한의 부산진단검사의학팀1500377
//==============================================================================
// 개발일자      : 2015.06.27
// 개발자        : 곽윤설
// 수정사항      : 요 10종 추가
// 참고사항      : 한의 재단특검출장파트1500104
//==============================================================================
// 개발일자      : 2017.02.15
// 개발자        : 유동구
// 수정사항      : 프로그램 개편(배치전일 경우 지표물질항목 검사 제외)
// 참고사항      : 기존에 있었지만 코드로 정해져 있어서 출력되는 부분 발생(자동으로 지표물질 찾아서 확인하게 적용)
//==============================================================================
unit ld5q08;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD5Q08 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    cmbJisa: TComboBox;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    qryProfileG: TQuery;
    qryNo_hangmok: TQuery;
    Bevel2: TBevel;
    Btn_BacordPnt: TBitBtn;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    qryCheck: TQuery;
    Cmb_gubn: TComboBox;
    Panel3: TPanel;
    Rbtn_total: TRadioButton;
    Rbtn_each: TRadioButton;
    Label4: TLabel;
    sampleno1: TEdit;
    sampleno2: TEdit;
    qryProfile: TQuery;
    Panel27: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    qryComm12: TQuery;
    qryHangmok: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure Btn_BacordPntClick(Sender: TObject);
    procedure Rbtn_eachClick(Sender: TObject);
    procedure Rbtn_totalClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUNJU, UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn, UV_vHangmok, UV_vtkgum_yh : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD5Q08: TfrmLD5Q08;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q081, LD5Q082;

{$R *.DFM}

procedure TfrmLD5Q08.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD5Q08.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD5Q08.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo       := VarArrayCreate([0, 0], varOleStr);
      UV_vName     := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin    := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa     := VarArrayCreate([0, 0], varOleStr);
      UV_vBUNJU    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc  := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vHangmok  := VarArrayCreate([0, 0], varOleStr);
      UV_vtkgum_yh  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,       iLength);
      VarArrayRedim(UV_vName,     iLength);
      VarArrayRedim(UV_vJumin,    iLength);
      VarArrayRedim(UV_vJisa,     iLength);
      VarArrayRedim(UV_vBUNJU,    iLength);
      VarArrayRedim(UV_vDesc_dc,  iLength);
      VarArrayRedim(UV_vSampNo,   iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
      VarArrayRedim(UV_vHangmok,  iLength);
      VarArrayRedim(UV_vtkgum_yh,  iLength);
   end;
end;

function TfrmLD5Q08.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') or
      ((Rbtn_total.Checked = False) and (Rbtn_each.Checked = False) ) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q08.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Rbtn_total.Checked := True;
end;

procedure TfrmLD5Q08.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vBUNJU[DataRow-1];
      4 : Value := UV_vSampNo[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vName[DataRow-1];
      7 : Value := UV_vDesc_dc[DataRow-1];
      8 : Value := UV_vDat_gmgn[DataRow-1];
      9 : Value := UV_vHangmok[DataRow-1];
     10 : Value := UV_vtkgum_yh[DataRow-1];
   end;
end;

procedure TfrmLD5Q08.btnQueryClick(Sender: TObject);
var
    sSQlText, sCheck_List, sCheck_desc_List, sHul_List, sDesc_profile : String;
    iCnt, iRet, iTemp, iRow : Integer;
    bChk_TC77 : Boolean;
    vCod_chuga : Variant;
begin
   inherited;
   sSQlText := '';
   iRow := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);


   sSQlText := ' SELECT dbo.F_GET_DANCHE_FIND(G.cod_dc) Desc_dc, B.num_bunju, T.cod_prf, T.cod_chuga As Tcod_chuga,    ' + #13 +
              '        G.desc_name,  G.num_jumin,  G.dat_gmgn,   G.cod_jisa,   G.gubn_nosin, G.gubn_adult, G.gubn_life, ' + #13 +
              '        G.Gubn_bogun, G.Gubn_agab,  G.Gubn_tkgm,  G.cod_dc,     G.Gubn_nsyh,  G.Gubn_adyh,  G.Gubn_lfyh, ' + #13 +
              '        G.Gubn_bgyh,  G.Gubn_agyh,  G.cod_hul,    G.cod_jangbi, G.cod_cancer, G.cod_chuga,  G.num_samp   ' + #13 +
              ' FROM Gumgin G with(nolock) left outer join Bunju B with(nolock) ON G.num_id = B.num_id and G.dat_gmgn= B.dat_gmgn            ' + #13 +
              '                            left outer join Tkgum T with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ' + #13 +
              ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ' + #13 +
              '   AND G.Dat_gmgn = ''' + edtBDate.Text          + ''' ' + #13;

   if Trim(edtDc.Text) <> '' then sSQlText := sSQlText + '  AND G.cod_dc = ''' + edtDc.Text + '''';

   If (BunjuNo1.Text <> '') and (BunjuNo2.Text <> '') Then  sSQlText := sSQlText  + ' And B.num_bunju >= ''' + BunjuNo1.Text + '''' +
                                                                                    ' And B.num_bunju <= ''' + BunjuNo2.Text + '''';
   if (sampleno1.Text <> '') and (sampleno2.Text <> '') Then sSQlText := sSQlText + ' And G.Num_Samp >= ''' + sampleno1.Text + '''' +
                                                                                    ' And G.Num_Samp <= ''' + sampleno2.Text + '''';

   sSQlText := sSQlText + ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      SQL.Clear;
      SQL.Text := sSQlText;
      Open;

      GP_query_log(qryBunju, 'LD5Q08', 'Q', 'N');

      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         for iCnt := 1 to RecordCount do
         begin
            Gride.Progress := iCnt;
//            application.ProcessMessage;

            sHul_List := UF_hangmokList;

            sCheck_List := '';
            sCheck_desc_List := '';
            sDesc_profile := '';
            bChk_TC77 := False;

            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               if pos('TC77',FieldByName('cod_prf').AsString) > 0 then bChk_TC77 := True;

               iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_chuga);
               for iTemp := 0 to iRet - 1 do
               begin
                 qryProfile.Active := False;
                 qryProfile.ParamByName('cod_pf').AsString := vCod_chuga[iTemp];
                 qryProfile.Active := True;
                 if qryProfile.RecordCount > 0 then sDesc_profile := sDesc_profile + qryProfile.FieldByName('desc_pf').AsString + ',';
               end;
            end;

            //항목 Check
            sSQlText := '';
            if Rbtn_total.Checked then
            begin
              sSQlText := sSQlText + ' Select * from hangmok with(nolock)';
              sSQlText := sSQlText + ' where cod_hm IN (''SEA2'', ''SE87'', ''SE42'', ''P012'', ''SE85'', ' +
                                     '                  ''SE89'', ''SE90'', ''SE91'', ''SE92'', ''SE93'', ' +
                                     '                  ''SE96'', ''SE97'', ''SE98'', ''SEA1'', ''JJFG'', ' +
                                     '                  ''SEA6'', ''H029'', ''SE85'', ''SE61'')';
            end
            else if Rbtn_each.Checked then
            begin
               sSQlText  := sSQlText + ' Select * from hangmok with(nolock) ';

               if      Cmb_gubn.ItemIndex = 0 then sSQlText := sSQlText + ' where cod_hm = ''SEA2'' '
               else if Cmb_gubn.ItemIndex = 1 then sSQlText := sSQlText + ' where cod_hm = ''SE87'' '
               else if Cmb_gubn.ItemIndex = 2 then sSQlText := sSQlText + ' where cod_hm = ''SE42'' '
               else if Cmb_gubn.ItemIndex = 3 then sSQlText := sSQlText + ' where cod_hm = ''P012'' '
               else if Cmb_gubn.ItemIndex = 4 then sSQlText := sSQlText + ' where cod_hm in (''SE85'',''SE89'',''SE90'',''SE91'',''SE92'',''SE93'',''SE96'',''SE97'',''SE98'',''SEA1'',''SE61'') '
               else if Cmb_gubn.ItemIndex = 5 then sSQlText := sSQlText + ' where cod_hm = ''JJFG'' '
               else if Cmb_gubn.ItemIndex = 6 then sSQlText := sSQlText + ' where cod_hm = ''SEA6'' '
               else if Cmb_gubn.ItemIndex = 7 then sSQlText := sSQlText + ' where cod_hm = ''H029'' '
               else if Cmb_gubn.ItemIndex = 8 then sSQlText := sSQlText + ' where cod_hm = ''SE85'' '
               else if Cmb_gubn.ItemIndex = 9 then sSQlText := sSQlText + ' where cod_hm = ''SE61'' ';
            end;

            sSQlText := sSQlText + ' and gubn_sahu = ''Y'' and ysno_hidden =''N'' ';
            sSQlText := sSQlText + ' order by cod_hm ';

            With qryCheck Do
            Begin
               Close;
               SQL.Clear;
               SQL.Text := sSQlText;
               Open;

               if Cmb_gubn.ItemIndex = 4 then
               begin
                  while not Eof do
                  begin
                     if bChk_TC77 then
                     begin
                        qryComm12.Active := False;
                        qryComm12.ParamByName('cod_hm').AsString := FieldByName('cod_hm').AsString;
                        qryComm12.Active := True;

                        if qryComm12.IsEmpty then
                        begin
                           if (pos(FieldByName('cod_hm').AsString,sHul_List) > 0) then
                           begin
                             sCheck_List      := sCheck_List + FieldByName('cod_hm').AsString + ', ';
                             sCheck_desc_List := sCheck_desc_List + qryComm12.FieldByName('검사항목명_한').AsString + ', ';
                           end;
                        end;
                     end
                     else
                     begin
                        qryHangmok.Active := False;
                        qryHangmok.ParamByName('cod_hm').AsString := FieldByName('cod_hm').AsString;
                        qryHangmok.Active := True;
                        if (pos(FieldByName('cod_hm').AsString,sHul_List) > 0) then
                        begin
                          sCheck_List      := sCheck_List + FieldByName('cod_hm').AsString + ', ';
                          sCheck_desc_List := sCheck_desc_List + qryHangmok.FieldByName('desc_kor').AsString + ', ';
                        end;
                     end;

                     Next;
                  end;

                  Close;
            end
            else
            begin
               while not Eof do
                  begin
                     if bChk_TC77 then
                     begin
                        qryComm12.Active := False;
                        qryComm12.ParamByName('cod_hm').AsString := FieldByName('cod_hm').AsString;
                        qryComm12.Active := True;

                        if qryComm12.IsEmpty then
                        begin
                           if (pos(FieldByName('cod_hm').AsString,sHul_List) > 0) then sCheck_List := sCheck_List + FieldByName('cod_hm').AsString + ', ';
                        end;
                     end
                     else
                     begin
                        if (pos(FieldByName('cod_hm').AsString,sHul_List) > 0) then sCheck_List := sCheck_List + FieldByName('cod_hm').AsString + ', ';
                     end;

                     Next;
                  end;

                  Close;
               End;
            end;

            if sCheck_List <> '' then
            begin
               UP_VarMemSet(iRow);

               UV_vNo[iRow]       := IntToStr(iRow+1);
               UV_VBUNJU[iRow]    := qryBunju.FieldByName('num_bunju').AsString;
               UV_vName[iRow]     := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[iRow]    := (Copy(qryBunju.FieldByName('num_jumin').AsString,1,6)) +'-'+ (Copy(qryBunju.FieldByName('num_jumin').AsString,7,1))
                                      + '******';
               UV_vJisa[iRow]     := qryBunju.FieldByName('cod_jisa').AsString;
               UV_vDesc_dc[iRow]  := qryBunju.FieldByName('desc_dc').AsString;
               UV_vSampNo[iRow]   := qryBunju.FieldByName('num_samp').AsString;
               UV_vDat_gmgn[iRow] := qryBunju.FieldByName('Dat_gmgn').AsString;
               UV_vHangmok[iRow]  := COPY(sCheck_List,1,length(sCheck_List) - 2);

               if Cmb_gubn.ItemIndex = 4 then  UV_vtkgum_yh[iRow] := COPY(sCheck_desc_List,1,length(sCheck_desc_List) - 2) + #13#10 + #13#10 +'(특검유형:'+ sDesc_profile +')'
               else UV_vtkgum_yh[iRow] := sDesc_profile;

               Inc(iRow);
            end;
            Next;
         end;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := iRow;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD5Q08.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;
end;

procedure TfrmLD5Q08.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
  if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD5Q08.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD5Q081 := TfrmLD5Q081.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q081.QuickRep.Preview
  else                                frmLD5Q081.QuickRep.Print;
end;

procedure TfrmLD5Q08.FormActivate(Sender: TObject);
begin
   inherited;
  // edtBdate.SetFocus;
end;

function TfrmLD5Q08.UF_hangmokList : String;
var
   sTemp, sHmCode, sTkgum_pf, sCod_pf :string;
   vCod_chuga : Variant;
   iRet, i : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
                  sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      if Trim(qryBunju.FieldByName('cod_prf').AsString) <> '' then
      begin
          sTkgum_pf := qryBunju.FieldByName('cod_prf').AsString;
          while length(sTkgum_pf) >= 4 do
          begin
              sCod_pf   := copy(sTkgum_pf, 1, 4);
              sTkgum_pf := copy(sTkgum_pf,5,length(sTkgum_pf));
              qryProfileG.Active := False;
              qryProfileG.ParamByName('cod_pf').AsString := sCod_pf;
              qryProfileG.Active := True;
              if qryProfileG.RecordCount > 0 then
              begin
                 while not qryProfileG.Eof do
                 begin
                    //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                    //----------------------------------------------------------
                    if (pos('TC41', qryBunju.FieldByName('cod_prf').AsString)    > 0) and
                       (pos('JJXE', qryBunju.FieldByName('Tcod_chuga').AsString) > 0) and
                       (qryProfileG.FieldByName('cod_hm').AsString = 'SE42') then
                    begin
                       // 혈중카복시(SE42) Skip...
                    end
                    else if pos(qryProfileG.FieldByName('cod_hm').AsString, sHmCode) = 0 then
                       sHmCode := sHmCode + qryProfileG.FieldByName('cod_hm').AsString;
                    //==========================================================

                    qryProfileG.Next;
                 end;
              end;

              qryProfileG.Active := False;

              sHmCode := sHmCode + qryBunju.FieldByName('Tcod_chuga').AsString;
          end; //end of while
      end; //end of if
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if      (copy(vCod_chuga[i],1,4) = 'JJFG')                                   then sTemp := sTemp + vCod_chuga[i]
         else if (copy(vCod_chuga[i],1,2) <> 'JJ') and (pos(vCod_chuga[i],sTemp) = 0) then sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   
   result := sTemp;
end;

function TfrmLD5Q08.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q08.Btn_BacordPntClick(Sender: TObject);
begin
  inherited;
  frmLD5Q082 := TfrmLD5Q082.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q082.QuickRep.Preview
  else                                frmLD5Q082.QuickRep.Print;
end;

procedure TfrmLD5Q08.Rbtn_eachClick(Sender: TObject);
begin
  Cmb_gubn.Visible := True;
  Cmb_gubn.ItemIndex := 0;
end;

procedure TfrmLD5Q08.Rbtn_totalClick(Sender: TObject);
begin
  Cmb_gubn.Visible := False;
end;

end.
