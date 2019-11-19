//==============================================================================
// 프로그램 설명 : 검사항목결과조회 (뇨화학)
// 시스템        : 통합검진
// 수정일자      : 2015.07.06
// 수정자        : 곽윤설
// 수정내용      : 신규
// 참고사항      : 한의 재단진단검사의학팀1500548
//==============================================================================

unit LD4Q68;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q68 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    Gride: TGauge;
    qryNo_hangmok: TQuery;
    qryHangmokList: TQuery;
    RGrp_sort: TRadioGroup;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    RB_bunju: TRadioButton;
    RB_samp: TRadioButton;
    mskTo_bunju: TMaskEdit;
    Label5: TLabel;
    mskFrom_bunju: TMaskEdit;
    mskFrom_samp: TMaskEdit;
    Label1: TLabel;
    mskTo_samp: TMaskEdit;
    qryHangmok: TQuery;
    GroupBox1: TGroupBox;
    CB01: TComboBox;
    ck_plus: TCheckBox;
    Ck_hold: TCheckBox;
    ck_nothing: TCheckBox;
    ck_exist: TCheckBox;
    Label3: TLabel;
    CB_sub: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSamp, UV_vBunju, UV_vName, UV_vHM : Variant;
    vArrHM : Array of Array of String;
    iArr : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
    procedure UF_GmsaHM;
  public
    { Public declarations }
  end;

var
  frmLD4Q68: TfrmLD4Q68;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q681;

{$R *.DFM}

procedure TfrmLD4Q68.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo       := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp     := VarArrayCreate([0, 0], varOleStr);
      UV_vBunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vName     := VarArrayCreate([0, 0], varOleStr);
      UV_vHM       := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo       , iLength);
      VarArrayRedim(UV_vSamp     , iLength);
      VarArrayRedim(UV_vBunju    , iLength);
      VarArrayRedim(UV_vName     , iLength);
      VarArrayRedim(UV_vHM       , iLength);
   end;
end;

function TfrmLD4Q68.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

   if (Ck_Plus.Checked = False) and (Ck_nothing.Checked = False) and
      (Ck_Hold.Checked = False) and (Ck_Exist.Checked = False)   then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

end;

procedure TfrmLD4Q68.FormCreate(Sender: TObject);
begin
   inherited;

   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   CB01.ItemIndex := 0;
   edtBDate.Text := DateToStr(date);
end;

procedure TfrmLD4Q68.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];    
      2 : Value := UV_vSamp[DataRow-1];
      3 : Value := UV_vBunju[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vHM[DataRow-1];
   end;
end;

procedure TfrmLD4Q68.btnQueryClick(Sender: TObject);
var Index, i : Integer;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3,
    sSelect, sWhere, sOrderBy, sHangmok, sResult : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   index := 0;

   // 조회조건 Check
   if not UF_Condition then Exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
     // SQL문 생성
     Close;
     sSelect := ' SELECT B.cod_bjjs, B.dat_bunju, B.num_bunju, G.dat_gmgn, B.desc_rackNo, G.num_jumin, G.num_samp, G.desc_name,      ' +
                '        G.num_id, G.Cod_jangbi, G.Cod_cancer, G.cod_chuga, G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm, G.cod_hul,       ' +
                '        G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, G.gubn_agab, G.gubn_agyh,   ' +
                '        K.DESC_GLKWA1, K.DESC_GLKWA2, K.DESC_GLKWA3, T.cod_prf, T.cod_chuga As Tcod_chuga FROM Bunju B WITH(NOLOCK) ' +
                ' Inner JOIN Gulkwa K WITH(NOLOCK) ON B.cod_bjjs = K.cod_bjjs and B.num_id = K.num_id and B.dat_bunju = K.dat_bunju  ' +
                '                                     and B.num_bunju = K.num_bunju                                                  ' +
                ' LEFT OUTER JOIN Gumgin G WITH(NOLOCK) ON B.cod_jisa=G.cod_jisa and B.num_id=G.num_id and B.dat_gmgn=G.dat_gmgn     ' +
                ' LEFT OUTER JOIN Tkgum T  WITH(NOLOCK) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and                 ' +
                '                                          G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha                      ';
     sWhere :=  ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
     sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + ''' ';
     sWhere := sWhere + ' And K.gubn_part = ''03'' ';

     if RB_bunju.Checked then
     begin
       if Trim(mskFrom_bunju.Text) <> '' then
          sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom_bunju.Text + ''' ';
       if Trim(mskTo_bunju.Text) <> '' then
          sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo_bunju.Text + ''' ';
     end
     else
     begin
       if Trim(mskFrom_samp.Text) <> '' then
          sWhere := sWhere + ' AND G.num_samp >= ''' + mskFrom_samp.Text + ''' ';
       if Trim(mskTo_samp.Text) <> '' then
          sWhere := sWhere + ' AND G.num_samp <= ''' + mskTo_samp.Text + ''' ';
     end;

     case RGrp_sort.ItemIndex of
       0 : sOrderBy := ' ORDER BY CAST(B.num_bunju AS INT) ';
       1 : sOrderBy := ' ORDER BY CAST(G.num_samp AS INT) ';
     end;

     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrderBy);
     Open;
     Gride.Progress := 0;
     Gride.MaxValue := RecordCount;

     if RecordCount > 0 then
     begin
       GP_query_log(qryBunju, 'LD4Q68', 'Q', 'N');
       UF_GmsaHM;
       while not qryBunju.EOF do
       begin
         Gride.Progress := Gride.Progress + 1;
         application.ProcessMessages;
         sResult := '';
         sHangmok := '';
         sHangmok := UF_HangmokCheck;
         UV_fGulkwa := '';
         UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
         UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
         UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

         for i:=0 to iArr-1 do
         begin
             if pos(vArrHM[i][0], sHangmok) > 0 then
             begin
                if ((ck_plus.Checked) and ((Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) = '양성') or
                                           (Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) = '약양성'))) OR
                   ((ck_hold.Checked) and (Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) = '보류')) OR
                   ((ck_nothing.Checked) and (Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) = '')) OR
                   ((ck_exist.Checked) and (Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) <> '')) then
                 begin
                    sResult := sResult + vArrHM[i][0] + ':' + Trim(copy(UV_fGulkwa, StrToInt(vArrHM[i][1]), 6)) + ',';
                 end;
             end;
         end;
         if Trim(sResult) <> '' then
         begin
            UP_VarMemSet(Index+1);
            UV_vNo[Index]    := IntToStr(Index+1);
            UV_vSamp[Index]  := FieldByName('num_samp').AsString;
            UV_vBunju[Index] := FieldByName('num_bunju').AsString;
            UV_vName[Index]  := FieldByName('desc_name').AsString;
            UV_vHM[Index]    := copy(sResult,1,length(sResult)-1);
            INC(Index);
         end;
         Next;
       end;
     end
     else
     begin
       // 자료가 없음을 표시
       GF_MsgBox('Information', 'NODATA');
       exit;
     end;  // if RecordCount > 0 then
     Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q68.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate)
   else if Sender = CB01 then
   begin
      if CB01.Itemindex = 3 then
      begin
         CB_Sub.Visible := True;
         CB_Sub.ItemIndex := 0;
      end
      else CB_Sub.Visible := False;
   end;
end;

procedure TfrmLD4Q68.btnPrintClick(Sender: TObject);
begin
  inherited;
     frmLD4Q681 := TfrmLD4Q681.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q681.QuickRep.Preview
     else                                frmLD4Q681.QuickRep.Print;
end;

function TfrmLD4Q68.UF_HangmokCheck : String;
var sALL_HangmokList, sHul_List, sEtc_Hangmok_hm, sSelect, sWhere1, sWhere2, sOderby : string;
    i : Integer;
begin
   Result := ''; sALL_HangmokList := ''; sHul_List := ''; sEtc_Hangmok_hm := '';
   sSelect := ''; sWhere1 := ''; sWhere2 := ''; sOderby := '';
   i := 0;

//------------------------------------------------------------------------------
//검사항목 나열...
//------------------------------------------------------------------------------
   sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOderby := '';

   //장비검사(항목나열)리스트 추출...
   //------------------------------------------------------------------------
   if Trim(qryBunju.FieldByName('cod_prf').AsString) <> '' then
   begin
     sWhere1 := Trim(qryBunju.FieldByName('Cod_Jangbi').AsString) + ''',''' +
                Trim(qryBunju.FieldByName('Cod_Hul').AsString)    + ''',''' +
                Trim(qryBunju.FieldByName('Cod_Cancer').AsString) + ''',''';
     For i := 1 to (length(qryBunju.FieldByName('cod_prf').AsString) div 4) do
     begin
       if i = (length(Trim(qryBunju.FieldByName('cod_prf').AsString)) div 4) then
         sWhere1 := sWhere1 + copy(Trim(qryBunju.FieldByName('cod_prf').AsString), (i*4)-3, 4)
       else
         sWhere1 := sWhere1 + copy(Trim(qryBunju.FieldByName('cod_prf').AsString), (i*4)-3, 4) + ''',''';
     end;
   end
   else
   begin
     sWhere1 := Trim(qryBunju.FieldByName('Cod_Jangbi').AsString) + ''',''' +
                Trim(qryBunju.FieldByName('Cod_Hul').AsString)    + ''',''' +
                Trim(qryBunju.FieldByName('Cod_Cancer').AsString) + ''',''';
   end;

   sEtc_Hangmok_hm := sEtc_Hangmok_hm + qryBunju.FieldByName('COD_CHUGA').AsString;

   // 노신유형Display
   if Trim(qryBunju.FieldByName('Gubn_Nosin').AsString) = '1' then
     sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '1', StrToInt(copy(qryBunju.FieldByName('Gubn_nsyh').AsString,1,1)));
   //성인병유형Display
   if Trim(qryBunju.FieldByName('Gubn_adult').AsString) = '1' then
     sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '4', StrToInt(copy(qryBunju.FieldByName('Gubn_adyh').AsString,1,1)));
   //간염유형Display
   if Trim(qryBunju.FieldByName('Gubn_agab').AsString) = '1' then
     sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '5', StrToInt(copy(qryBunju.FieldByName('Gubn_agyh').AsString,1,1)));
   //생애전환기유형Display
   if Trim(qryBunju.FieldByName('Gubn_life').AsString) = '1' then
     sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '7', StrToInt(copy(qryBunju.FieldByName('Gubn_lfyh').AsString,1,1)));

   if Trim(qryBunju.FieldByName('Tcod_chuga').AsString) <> '' then
   begin
     sWhere2 := sWhere2 + ''',''';
     For i := 1 to (length(qryBunju.FieldByName('Tcod_chuga').AsString) div 4) do
     begin
       if (i = (length(qryBunju.FieldByName('Tcod_chuga').AsString) div 4)) and (Trim(sEtc_Hangmok_hm) = '') then
         sWhere2 := sWhere2 + copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i*4)-3, 4)
       else
         sWhere2 := sWhere2 + copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i*4)-3, 4) + ''',''';
     end;
   end;

   if Trim(sEtc_Hangmok_hm) <> '' then
   begin
     sWhere2 := sWhere2 + ''',''';
     For i := 1 to (length(sEtc_Hangmok_hm) div 4) do
     begin
       if i = (length(sEtc_Hangmok_hm) div 4) then
         sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4)
       else
         sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
     end;
   end;

   with qryHangmokList do
   begin
     qryHangmokList.Active := False;

     sSelect := '  ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P WITH(NOLOCK) ' +
                '    LEFT OUTER JOIN hangmok H WITH(NOLOCK) ON P.cod_hm = H.cod_hm AND       ' +
                '    H.dat_apply <= ''' + qryBunju.FieldByName('dat_gmgn').AsString + '''    ' +
                '    WHERE P.cod_pf IN (''' + sWhere1 + ''')                                 ' +
                '    GROUP BY P.cod_hm, H.desc_kor, H.gubn_part )                            ' +
                ' UNION                                                                      ' +
                '  ( SELECT cod_hm, desc_kor, gubn_part  FROM hangmok WITH(NOLOCK)           ' +
                '    WHERE cod_hm IN (''' + sWhere2 + ''')                                   ' +
                '    AND dat_apply <= ''' + qryBunju.FieldByName('dat_gmgn').AsString + ''') ' +
                ' ORDER BY H.gubn_part Desc, P.cod_hm                                        ';

     qryHangmokList.SQL.Clear;
     qryHangmokList.SQL.Add(sSelect + sOderby);
     qryHangmokList.Active := True;

     if qryHangmokList.RecordCount > 0 then
     begin
       while not Eof do
       begin
         sALL_HangmokList := sALL_HangmokList + qryHangmokList.FieldByName('cod_hm').AsString;

         if qryHangmokList.FieldByName('gubn_part').AsInteger < 10 then
           sHul_List := sHul_List + qryHangmokList.FieldByName('cod_hm').AsString;

         qryHangmokList.Next;
       end;
     end;
   end;

   Result := sHul_List;
end;

function  TfrmLD4Q68.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      Result := Result + FieldByName('desc_jangbi').AsString;
    end;
    Active := False;
  end;
end;


procedure  TfrmLD4Q68.UF_GmsaHM;
var sGmsaHM : String;
    i : Integer;
begin
   sGmsaHM := '';

   if CB01.ItemIndex = 0 then sGmsaHM := 'U033'
   else if CB01.ItemIndex = 1 then sGmsaHM := 'Y004'
   else if CB01.ItemIndex = 2 then sGmsaHM := 'U017U046'
   else if (CB01.ItemIndex = 3) and (CB_Sub.ItemIndex = 0) then sGmsaHM := 'U024U032'
   else if (CB01.ItemIndex = 3) and (CB_Sub.ItemIndex = 1) then sGmsaHM := 'U029'
   else if (CB01.ItemIndex = 3) and (CB_Sub.ItemIndex = 2) then sGmsaHM := 'U030U044U051'
   else if (CB01.ItemIndex = 3) and (CB_Sub.ItemIndex = 3) then sGmsaHM := 'U031'
   else if (CB01.ItemIndex = 4) then sGmsaHM := 'Y010';
   iArr := Round(Length(sGmsaHM)/4);
   setLength(vArrHM, iArr);

   qryHangmok.Close;
   qryHangmok.ParamByName('dat_apply').AsString := edtBDate.Text;
   qryHangmok.Filtered := False;
   qryHangmok.Open;
   qryHangmok.Filtered := True;
   for i:=0 to iArr-1 do
   begin
      qryHangmok.Filter := ' cod_hm = ''' + copy(sGmsaHM, (i*4)+1, 4) + ''' ';
      setLength(vArrHM[i], 2);
      vArrHM[i][0] := qryHangmok.FieldByName('cod_hm').AsString;
      vArrHM[i][1] := qryHangmok.FieldByName('pos_start').AsString;
   end;
   qryHangmok.Filtered := False;
   qryHangmok.Close;
end;

end.
