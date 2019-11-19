//==============================================================================
// 프로그램 설명 : 항목별검사리스트(마약,외주,특검)
// 시스템        : 통합검진
// 수정일자      : 2015.06.23
// 수정자        : 곽윤설
// 수정내용      : 신규
// 참고사항      : 한의 재단진단검사의학팀1500075
//==============================================================================
// 수정일자      : 2015.10.14
// 수정자        : 곽윤설
// 수정내용      : 씨젠+특검항목 - U063,U064,U065 추가
// 참고사항      : 한의 재단진단검사의학팀1500856
//==============================================================================
// 수정일자      : 2016.01.12
// 수정자        : 박수지
// 수정내용      : 외주검사2 (U066,U067,U068,U069,U070,U071) 추가
// 참고사항      :  한의 재단진단검사의학팀1600023
//==============================================================================
// 수정일자      : 2017.02.24
// 수정자        : 박수지
// 수정내용      : 외주검사2 (U073) 추가
// 참고사항      :  한의 재단진단검사의학팀1700100
//==============================================================================
// 수정일자      : 2017.03.23
// 수정자        : 박수지
// 수정내용      : 외주검사(u020, u028, u043)삭제
// 참고사항      :  한의 재단진단검사의학팀1700154
//==============================================================================
// 수정일자      : 2018.08.27
// 수정자        : 박수지
// 수정내용      : 외주검사(u053, u054)추가
// 참고사항      : 한의 재단진단검사의학팀장 1800070
//==============================================================================
unit LD4Q67;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q67 = class(TfrmSingle)
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
    qryHangmok: TQuery;
    Label_iArr: TLabel;
    CB_Nicotine: TRadioButton;
    CB_Tkgum: TRadioButton;
    CB_Narcotic: TRadioButton;
    CB_Sub: TRadioButton;
    CB_Sub2: TRadioButton;
    RB_bunju: TRadioButton;
    mskFrom_bunju: TMaskEdit;
    Label5: TLabel;
    mskTo_bunju: TMaskEdit;
    RB_samp: TRadioButton;
    mskFrom_samp: TMaskEdit;
    Label1: TLabel;
    mskTo_samp: TMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSamp, UV_vBunju, UV_vName, UV_vBirth, UV_vDat_gmgn, UV_vHM : Variant;
    sNicotine, sTkgum, sNeodine, sNarcotic, sSub, sSub2 : String;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    vArrHM : array of String;
    iArr : Integer;
  end;

var
  frmLD4Q67: TfrmLD4Q67;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q671, LD4Q672, LD4Q673, LD4Q674, LD4Q675, LD4Q676 ;

{$R *.DFM}

procedure TfrmLD4Q67.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo       := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp     := VarArrayCreate([0, 0], varOleStr);
      UV_vBunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vName     := VarArrayCreate([0, 0], varOleStr);
      UV_vBirth    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vHM       := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo       , iLength);
      VarArrayRedim(UV_vSamp     , iLength);
      VarArrayRedim(UV_vBunju    , iLength);
      VarArrayRedim(UV_vName     , iLength);
      VarArrayRedim(UV_vBirth    , iLength);
      VarArrayRedim(UV_vDat_gmgn , iLength);
      VarArrayRedim(UV_vHM       , iLength);
   end;
end;

function TfrmLD4Q67.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

   if (CB_Nicotine.Checked = False) and (CB_Tkgum.Checked = False)    and
      (CB_Narcotic.Checked = False) and (CB_Sub.Checked = False) and (CB_Sub2.Checked = False) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

end;

procedure TfrmLD4Q67.FormCreate(Sender: TObject);
begin
   inherited;

   sNicotine := 'U033';
   sTkgum    := 'U059U060U061U063U064U065';
   sNarcotic := 'U024U029U030U031U032U044U051';
   sSub      := 'U023U026U037U045U047U050U062';
   sSub2     := 'U066U067U068U069U070U071U072U073U074U053U054';

   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   if copy(GV_sUserId,1,2) ='12' then
   begin
     Label2.caption := '검진일자';
     Label7.caption := '지    사 : ';
   end;
   edtBDate.Text := GV_sToday;
end;

procedure TfrmLD4Q67.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSamp[DataRow-1];
      3 : Value := UV_vBunju[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vBirth[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vHM[DataRow-1];
   end;
end;

procedure TfrmLD4Q67.btnQueryClick(Sender: TObject);
var Index, i, iArr : Integer;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3,
    sSelect, sWhere, sOrderBy, sCheck, sHangmok, sCkeck_HM : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sCheck := ''; sCkeck_HM := '';
   index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   if CB_Nicotine.Checked then  sCkeck_HM := sCkeck_HM + sNicotine;
   if CB_Tkgum.Checked    then  sCkeck_HM := sCkeck_HM + sTkgum;
   if CB_Narcotic.Checked then  sCkeck_HM := sCkeck_HM + sNarcotic;
   if CB_Sub.Checked      then  sCkeck_HM := sCkeck_HM + sSub;
   if CB_Sub2.Checked     then  sCkeck_HM := sCkeck_HM + sSub2;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
     // SQL문 생성
     Close;
     if copy(cmbJisa.Text,1,2) ='12' then
     begin
     sSelect := ' SELECT  G.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.desc_name,             ' +
                '         G.num_jumin, G.Cod_jangbi, G.Cod_cancer, G.cod_chuga, G.gubn_neawon, G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm,  ' +
                '         G.gubn_bogun,  G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, G.gubn_agab, G.gubn_agyh,    ' +
                '         G.cod_hul, T.cod_prf, T.cod_chuga As Tcod_chuga FROM gumgin G WITH(NOLOCK)                                    ' +
                ' LEFT OUTER JOIN Tkgum T  WITH(NOLOCK) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and                    ' +
                '                                          G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha                         ' +
                ' LEFT OUTER JOIN BUNJU B  WITH(NOLOCK) ON G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and                          ' +
                '                                          G.dat_gmgn = B.dat_gmgn                                                      ' ;
     end
     else
     begin
     sSelect := ' SELECT B.cod_bjjs, B.dat_bunju, B.num_bunju, G.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.desc_name,             ' +
                '        G.num_jumin, G.Cod_jangbi, G.Cod_cancer, G.cod_chuga, G.gubn_neawon, G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm,  ' +
                '        G.gubn_bogun,  G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, G.gubn_agab, G.gubn_agyh,    ' +
                '        G.cod_hul, T.cod_prf, T.cod_chuga As Tcod_chuga FROM Bunju B WITH(NOLOCK)                                     ' +
                ' LEFT OUTER JOIN Gumgin G WITH(NOLOCK) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                ' LEFT OUTER JOIN Tkgum T  WITH(NOLOCK) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and                   ' +
                '                                          G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha                        ';
     end;
     //여의도 지사에서 검진일로 변경, 분주지사->지사로 요청 ..한미영
     if copy(cmbJisa.Text,1,2) ='12' then sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''''
     else sWhere  :=  ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';

     if copy(cmbJisa.text, 1, 2) = '12' then
     sWhere := sWhere + ' And G.Dat_Gmgn = ''' + edtBDate.Text + ''' '
     else  sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + ''' ';

     if sCheck <> '' then
        sWhere := sWhere + ' And B.cod_jisa in (' + copy(sCheck,1, Length(sCheck)-1) + ') ';

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
       0 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(B.num_bunju AS INT) ';
       1 : sOrderBy := ' ORDER BY CAST(G.num_samp AS INT)';
     end;

     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrderBy);
     Open;
     Gride.Progress := 0;
     Gride.MaxValue := RecordCount;
     if RecordCount > 0 then
     begin
       GP_query_log(qryBunju, 'LD4Q67', 'Q', 'N');

       iArr := Round(Length(sCkeck_HM)/4);
       SetLength(vArrHM, iArr);

       for i:=0 to iArr-1 do
         vArrHM[i] := copy(sCkeck_HM, (i*4)+1, 4);

       while not qryBunju.EOF do
       begin
         Gride.Progress := Gride.Progress + 1;
         application.ProcessMessages;

         sHangmok := '';
         sHangmok := UF_HangmokCheck;

         for i:=0 to iArr-1 do
         begin
           if pos(vArrHM[i], sHangmok) > 0 then
           begin
             UP_VarMemSet(Index);
             UV_vNo[Index]       := IntToStr(Index+1);
             UV_vSamp[Index]     := qryBunju.FieldByName('num_samp').AsString;
             if (copy(cmbJisa.Text,1,2) ='12') and  (edtBDate.text = GV_sToday) then
             UV_vBunju[Index]    := ''
             else UV_vBunju[Index]    := qryBunju.FieldByName('num_bunju').AsString;
             UV_vName[Index]     := qryBunju.FieldByName('desc_name').AsString;
             UV_vBirth[Index]    := Copy(qryBunju.FieldByName('num_jumin').AsString,1,6) + '-' +
                                    Copy(qryBunju.FieldByName('num_jumin').AsString,7,1);
             UV_vDat_gmgn[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
             UV_vHM[Index]       := vArrHM[i];
             INC(Index);
           end;
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

procedure TfrmLD4Q67.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q67.btnPrintClick(Sender: TObject);
begin
  inherited;
  if cb_nicotine.checked then
  begin
     frmLD4Q671 := TfrmLD4Q671.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q671.QuickRep.Preview
     else                                frmLD4Q671.QuickRep.Print;
  end
  else if cb_narcotic.checked then
  begin
     frmLD4Q673 := TfrmLD4Q673.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q673.QuickRep.Preview
     else                                frmLD4Q673.QuickRep.Print;
  end
  else if cb_tkgum.checked then
  begin
     frmLD4Q674 := TfrmLD4Q674.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q674.QuickRep.Preview
     else                                frmLD4Q674.QuickRep.Print;
  end
  else if cb_sub.checked then
  begin
     frmLD4Q675 := TfrmLD4Q675.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q675.QuickRep.Preview
     else                                frmLD4Q675.QuickRep.Print;
  end
  else if cb_sub2.checked then
  begin
     frmLD4Q676 := TfrmLD4Q676.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD4Q676.QuickRep.Preview
     else                                frmLD4Q676.QuickRep.Print;
  end;
end;


function TfrmLD4Q67.UF_HangmokCheck : String;
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

   sEtc_Hangmok_hm := qryBunju.FieldByName('COD_CHUGA').AsString;

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

function  TfrmLD4Q67.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.
