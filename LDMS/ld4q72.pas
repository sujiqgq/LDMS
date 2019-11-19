//==============================================================================
// 프로그램 설명 : 생화학, 혈액학, Urine 결과지 출력
// 시스템        : 통합검진
// 수정일자      : 2015.11.18
// 수정자        : 곽윤설
// 수정내용      : 신규개발...
// 참고사항      : 한의 부산진단검사의학팀1500975 / 부산 유희짐 책임
//==============================================================================

unit LD4Q72;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;
type
  TfrmLD4Q72 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Gride: TGauge;
    qryTkgum: TQuery;
    qry_Profile: TQuery;
    qryNo_hangmok: TQuery;
    RB_jumin: TRadioButton;
    RB_Samp: TRadioButton;
    Msk_jumin: TMaskEdit;
    RG_part: TRadioGroup;
    GroupBox1: TGroupBox;
    CB_bul01: TCheckBox;
    CB_bul02: TCheckBox;
    CB_bul03: TCheckBox;
    GroupBox2: TGroupBox;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    Panel23: TPanel;
    Panel24: TPanel;
    Panel25: TPanel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit15: TEdit;
    Edit16: TEdit;
    Edit17: TEdit;
    Edit18: TEdit;
    Edit19: TEdit;
    Edit20: TEdit;
    Edit21: TEdit;
    Edit22: TEdit;
    pnl1: TPanel;
    pnl2: TPanel;
    pnl3: TPanel;
    pnl4: TPanel;
    pnl5: TPanel;
    pnl6: TPanel;
    pnl7: TPanel;
    pnl8: TPanel;
    pnl9: TPanel;
    pnl10: TPanel;
    pnl11: TPanel;
    pnl12: TPanel;
    pnl13: TPanel;
    pnl14: TPanel;
    pnl17: TPanel;
    pnl16: TPanel;
    pnl15: TPanel;
    pnl20: TPanel;
    pnl19: TPanel;
    pnl18: TPanel;
    pnl22: TPanel;
    pnl21: TPanel;
    qryHangmok: TQuery;
    qryGumBul: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_VarArr(iPart : Integer);
    procedure UP_GB_Clear;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
     UV_vGulkwa, UV_vNum_id, UV_vCod_HM, UV_vGumbul : Variant;
     vArr_HM : array of array of Variant;
  end;

var
  frmLD4Q72: TfrmLD4Q72;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4P721, LD4P722, LD4P723;

{$R *.DFM}

procedure TfrmLD4Q72.UP_VarMemSet(iLength : Integer);
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
      UV_vNum_id := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_HM := VarArrayCreate([0, 0], varOleStr);
      UV_vGumbul := VarArrayCreate([0, 0], varOleStr);
      UV_vGulkwa := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_vGulkwa, iLength);
      VarArrayRedim(UV_vNum_id, iLength);
      VarArrayRedim(UV_vCod_HM, iLength);
      VarArrayRedim(UV_vGumbul, iLength);
   end;
end;

function TfrmLD4Q72.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q72.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   edtBDate.Text := GV_sToday;

end;

procedure TfrmLD4Q72.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := UV_v003[DataRow-1];
      4 : Value := UV_v004[DataRow-1];
      5 : Value := UV_v005[DataRow-1];
      6 : Value := UV_v006[DataRow-1];
   end;
end;

procedure TfrmLD4Q72.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sTemp, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   index := 0;

   //배열 초기화
   vArr_HM := nil;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   Gride.Progress := 0;

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT G.cod_jisa, G.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, D.desc_dc, G.desc_name, G.num_jumin, B.num_bunju,       ' +
                 '        G.Cod_jangbi, G.Cod_cancer, G.cod_chuga, G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm, G.gubn_bogun, G.gubn_bgyh,        ' +
                 '        G.cod_hul, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, G.gubn_agab, G.gubn_agyh, T.cod_prf,               ' +
                 '        T.cod_chuga As Tcod_chuga, B.desc_glkwa1, B.desc_glkwa2, B.desc_glkwa3                                             ' +
                 ' FROM gumgin G with(nolock)                                                                                                ' +
                 ' INNER JOIN Gulkwa B with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn           ' +
                 ' LEFT OUTER JOIN Danche D with(nolock) ON D.cod_dc = G.cod_dc                                                              ' +
                 ' LEFT OUTER JOIN Tkgum T with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn ' +
                 '                                         and G.gubn_tkgm = T.gubn_cha                                                      ';

      sWhere := ' Where G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And G.Dat_gmgn = ''' + edtBDate.Text + ''' ';

      if RB_jumin.Checked then
      begin
         sWhere := sWhere + ' AND G.num_jumin = ''' + StringReplace(Msk_jumin.Text, '-', '', [rfReplaceAll, rfIgnoreCase]) + ''' ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + ''' ';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + ''' ';
      end;

      if      RG_part.ItemIndex = 0 then sWhere := sWhere + ' AND B.gubn_part = ''01'' '
      else if RG_part.ItemIndex = 1 then sWhere := sWhere + ' AND B.gubn_part = ''02'' '
      else if (RG_part.ItemIndex = 2) or (RG_part.ItemIndex = 3) then sWhere := sWhere + ' AND B.gubn_part = ''03'' ';

      sOrderBy := ' ORDER BY B.cod_jisa, CAST(G.num_samp AS INT) ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q72', 'Q', 'N');
         UP_VarArr(RG_part.ItemIndex);
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            sTemp := UF_HangmokCheck;

            UP_VarMemSet(Index);

            UV_fGulkwa1 := FieldByName('desc_glkwa1').AsString;
            UV_fGulkwa2 := FieldByName('desc_glkwa2').AsString;
            UV_fGulkwa3 := FieldByName('desc_glkwa3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

            UV_v001[Index] := IntToStr(Index+1);
            UV_v002[Index] := FieldByName('num_samp').AsString;
            UV_v003[Index] := FormatFloat('######-#######',GF_StrToNum(FieldByName('num_jumin').AsString));
            UV_v004[Index] := FieldByName('desc_name').AsString;
            UV_v005[Index] := FieldByName('desc_dc').AsString;
            UV_v006[Index] := FieldByName('dat_gmgn').AsString;
            UV_vGulkwa[Index] := sGlkwa;
            UV_vNum_id[Index] := FieldByName('Num_id').AsString;
            UV_vCod_HM[Index] := sTemp;
            UV_vGumbul[Index] := '';

            qryGumBul.Close;
            qryGumBul.ParamByName('Cod_jisa').AsString := FieldByName('Cod_jisa').AsString;
            qryGumBul.ParamByName('num_id').AsString   := FieldByName('Num_id').AsString;
            qryGumBul.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
            qryGumBul.Open;
            if qryGumBul.RecordCount > 0 then
            begin
               while not qryGumBul.Eof do
               begin
                 UV_vGumbul[Index] := UV_vGumbul[Index] + qryGumBul.FieldByName('gubn_bul').AsString;
                 qryGumBul.Next;
               end;
            end;
            qryGumBul.Close;

            INC(Index);

            Next;
         end;
         // Data건수 표시
         GP_SetDataCnt(grdMaster);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_Profile do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q72.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate)
   else if Sender = RB_jumin then
   begin
      Msk_jumin.Enabled := True;
      MEdt_SampS.Enabled := False;
      MEdt_SampE.Enabled := False;
   end
   else if Sender = RB_Samp then
   begin
      Msk_jumin.Enabled := False;
      MEdt_SampS.Enabled := True;
      MEdt_SampE.Enabled := True;
   end;
end;

procedure TfrmLD4Q72.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RG_part.ItemIndex = 0 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD4P721 := TfrmLD4P721.Create(Self);
        frmLD4P721.QuickRep.Preview;
    end
     else
     begin
        frmLD4P721 := TfrmLD4P721.Create(Self);
        frmLD4P721.QuickRep.Print;
     end;
  end
  else if RG_part.ItemIndex = 1 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD4P722 := TfrmLD4P722.Create(Self);
        frmLD4P722.QuickRep.Preview;
     end
     else
     begin
        frmLD4P722 := TfrmLD4P722.Create(Self);
        frmLD4P722.QuickRep.Print;
     end;
  end
  else if RG_part.ItemIndex = 2 then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD4P723 := TfrmLD4P723.Create(Self);
        frmLD4P723.QuickRep.Preview;
     end
     else
     begin
        frmLD4P723 := TfrmLD4P723.Create(Self);
        frmLD4P723.QuickRep.Print;
     end;
  end;
end;

function TfrmLD4Q72.UF_HangmokCheck : String;
var sHangmokList, vSQL_Text : string;
    iRet, i : Integer;
    vCod_chuga : Variant;
begin
   Result := ''; sHangmokList := ''; vSQL_Text := '';  
   i := 0;
   
   //---- 장비 검사(신신)..
   qry_Profile.Active := false;
   qry_Profile.ParamByName('cod_pf').AsString := qryBunju.FieldByName('cod_Jangbi').AsString;
   qry_Profile.Active := True;
   if qry_Profile.RecordCount > 0 then
   begin
      while not qry_Profile.Eof do
      begin
         if pos(qry_Profile.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_Profile.FieldByName('cod_hm').AsString;
         qry_Profile.Next;
      end;
   end;

   //---- 혈액 검사..
   qry_Profile.Active := false;
   qry_Profile.ParamByName('cod_pf').AsString :=  qryBunju.FieldByName('cod_hul').AsString;
   qry_Profile.Active := True;
   if qry_Profile.RecordCount > 0 then
   begin
      while not qry_Profile.Eof do
      begin
         if pos(qry_Profile.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_Profile.FieldByName('cod_hm').AsString;
         qry_Profile.Next;
      end;
   end;

   //---- Cancer 검사..
   qry_Profile.Active := false;
   qry_Profile.ParamByName('cod_pf').AsString :=  qryBunju.FieldByName('cod_Cancer').AsString;
   qry_Profile.Active := True;
   if qry_Profile.RecordCount > 0 then
   begin
      while not qry_Profile.Eof do
      begin
         if pos(qry_Profile.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_Profile.FieldByName('cod_hm').AsString;
         qry_Profile.Next;
      end;
   end;

   //---- 추가 검사..
   for i := 1 to Length(qryBunju.FieldByName('cod_chuga').AsString) div 4 do
   begin
      if (copy(copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4),1,2) <> 'JJ') and
         (pos(copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4), sHangmokList) <= 0) then
          sHangmokList := sHangmokList + copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4);
   end;

   // 노신유형Display
   if qryBunju.FieldByName('gubn_nosin').AsString = '1' then
   begin
      sHangmokList := sHangmokList + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qryBunju.FieldByName('gubn_nsyh').AsInteger);
   end;

   //성인병유형Display
   if qryBunju.FieldByName('gubn_adult').AsString = '1' then
   begin
      sHangmokList := sHangmokList + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qryBunju.FieldByName('gubn_adyh').AsInteger);
   end;

   //간염유형Display
   if qryBunju.FieldByName('gubn_agab').AsString = '1' then
   begin
      sHangmokList := sHangmokList + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qryBunju.FieldByName('gubn_adyh').AsInteger);
   end;
   // 생애전환기유형Display
   if qryBunju.FieldByName('gubn_life').AsString = '1' then
   begin
      sHangmokList := sHangmokList + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qryBunju.FieldByName('gubn_lfyh').AsInteger);
   end;

   if (qryBunju.FieldByName('gubn_tkgm').AsString = '1') or
      (qryBunju.FieldByName('gubn_tkgm').AsString = '2') then
   begin
      //---- 특검Profile...
      iRet := GF_MulToSingle(qryBunju.FieldByName('cod_prf').AsString, 4, vCod_chuga);
      for i := 0 to iRet - 1 do
      begin
         qry_Profile.Active := false;
         qry_Profile.ParamByName('cod_pf').AsString    :=  vCod_chuga[i];
         qry_Profile.Active := True;
         if qry_Profile.RecordCount > 0 then
         begin
            while not qry_Profile.Eof do
            begin
               if pos(qry_Profile.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
                  sHangmokList := sHangmokList + qry_Profile.FieldByName('cod_hm').AsString;

               qry_Profile.Next;
            end;
         end;
      end; // end of for[temp];

      //---- 특검추가 검사..
      for i := 1 to Length(qryBunju.FieldByName('Tcod_chuga').AsString) div 4 do
      begin
         if copy(copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i * 4) - 3, 4),1,1) = 'U' then
         begin
            if pos(copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i * 4) - 3, 4), sHangmokList) <= 0 then
               sHangmokList := sHangmokList + copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i * 4) - 3, 4);
         end;
      end;
   end;

   Result := sHangmokList;
end;

function  TfrmLD4Q72.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryNo_hangmok do
   begin
      Active := False;
      ParamByName('dat_apply').AsString := sDate;
      ParamByName('gubn_code').AsString := sGubun;
      ParamByName('gubn_yh').AsInteger  := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD4Q72.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i, j : Integer;
begin
  inherited;
   //초기화
   UP_GB_Clear;
   if NewRow = 0 then exit;

   //Row 변경 시..
   for i := 0 to Round(length(vArr_HM)-1) do
   begin
      if pos(vArr_HM[i][0], UV_vCod_HM[NewRow-1]) > 0 then
      begin
          for j := 0 to GroupBox2.ControlCount - 1 do
          begin
             if 'pnl' + IntToStr(i+1) = TPanel(GroupBox2.Controls[j]).Name then
                TPanel(GroupBox2.Controls[j]).Caption := vArr_HM[i][1];
             if 'Edit' + IntToStr(i+1) = TEdit(GroupBox2.Controls[j]).Name then
                TEdit(GroupBox2.Controls[j]).Text := Trim(Copy(UV_vGulkwa[NewRow-1], StrToInt(vArr_HM[i][2]), 6));
          end;
      end;
   end;
   if ((pos('01', UV_vGumbul[NewRow-1]) > 0) and (RG_part.ItemIndex = 0)) or
      ((pos('03', UV_vGumbul[NewRow-1]) > 0) and (RG_part.ItemIndex = 1)) then CB_bul01.Checked := True;
   if (pos('02', UV_vGumbul[NewRow-1]) > 0) then CB_bul02.Checked := True;
   if (pos('08', UV_vGumbul[NewRow-1]) > 0) then CB_bul03.Checked := True;

end;

procedure TfrmLD4Q72.UP_VarArr(iPart : Integer);
var sHangmok : String;
    i : Integer;
begin
  inherited;
  case iPart of
    0 : sHangmok := 'C009C010C011C025C026C027C028C032C042';
    1 : sHangmok := 'H001H002H003H004H005H006H007H008H009H010H011H012H013H014H015H016H017H018H019H020H021H022';
    2 : sHangmok := 'U001U002U003U004U005U006U007U008U009U010U011U012U013';
    3 : sHangmok := 'Y004';
  end;

  setLength(vArr_HM, Round(length(sHangmok)/4));

  with qryHangmok do
  begin
     Close;
     Filtered := False;
     Open;
     Filtered := True;
     for i := 0 to Round(length(sHangmok)/4) do
     begin
        Filter := ' Cod_hm = ''' + Trim(Copy(sHangmok, (i*4)+1, 4)) + ''' ';
        if RecordCount > 0 then
        begin
           setlength(vArr_HM[i], 4);
           vArr_HM[i][0] := FieldByName('Cod_hm').AsString;
           vArr_HM[i][1] := FieldByName('desc_kor').AsString;
           vArr_HM[i][2] := FieldByName('pos_start').AsString;
           if (FieldByName('desc_unit').AsString <> '음성') and
              (FieldByName('desc_unit').AsString <> '양성') then vArr_HM[i][3] := FieldByName('desc_unit').AsString
           else vArr_HM[i][3] := '';
        end;
     end;
     Filtered := False;
     Close;
  end;
end;

procedure TfrmLD4Q72.UP_GB_Clear;
var i : Integer;
begin
   for i := 0 to GroupBox2.ControlCount - 1 do
   begin
      if pos('pnl', TPanel(GroupBox2.Controls[i]).Name) > 0 then TPanel(GroupBox2.Controls[i]).Caption := '';
      if pos('Edit', TEdit(GroupBox2.Controls[i]).Name) > 0 then TEdit(GroupBox2.Controls[i]).Text := '';
   end;

   if RG_part.ItemIndex = 0 then
   begin
      CB_bul01.Caption := '검체혼탁';
      CB_bul02.Caption := '검체용혈';
      CB_bul03.Caption := '보류';
      CB_bul01.Visible := True;
      CB_bul02.Visible := True;
   end
   else if RG_part.ItemIndex = 1 then
   begin
      CB_bul01.Caption := 'CBC CLOT';
      CB_bul02.Caption := '검체용혈';
      CB_bul03.Caption := '보류';
      CB_bul01.Visible := True;
      CB_bul02.Visible := True;
   end
   else
   begin
      CB_bul01.Caption := '';
      CB_bul02.Caption := '';
      CB_bul03.Caption := '보류';
      CB_bul01.Visible := False;
      CB_bul02.Visible := False;
   end;
end;

end.
