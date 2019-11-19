//==============================================================================
// 프로그램 설명 : 출장 요화학10종 조회
// 시스템        : 통합검진
// 수정일자      : 2015.10.01
// 수정자        : 곽윤설
// 수정내용      : 신규개발...
// 참고사항      : 한의 재단진단검사의학팀1500796 / 한두례 책임
//==============================================================================

unit ld4q71;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q71 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Gride: TGauge;
    qryTkgum: TQuery;
    qry_gulkwa: TQuery;
    qry_Hangmok: TQuery;
    qryNo_hangmok: TQuery;
    CB_tk: TCheckBox;
    RadioGroup1: TRadioGroup;
    CB_U033: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005 : Variant;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;    
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q71: TfrmLD4Q71;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q711;

{$R *.DFM}

procedure TfrmLD4Q71.UP_VarMemSet(iLength : Integer);
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
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,    iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
   end;
end;

function TfrmLD4Q71.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q71.FormCreate(Sender: TObject);
begin
   inherited;

   if  GV_sUserId = '150913' then  //본원 -> 여의도 출장으로 옮김
   begin
   if dmComm.DBURINE.Connected = False then  dmComm.DBURINE.Connected := True;
   end;

   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   edtBDate.Text := GV_sToday;
end;

procedure TfrmLD4Q71.grdMasterCellLoaded(Sender: TObject; DataCol,
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
   end;
end;

procedure TfrmLD4Q71.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sTemp : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   index := 0;

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
      sSelect := ' SELECT G.cod_jisa, G.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.desc_name, G.num_jumin, B.num_bunju, ' +
                 '        G.Cod_jangbi, G.cod_hul, G.Cod_cancer, G.cod_chuga, G.gubn_neawon, G.gubn_nosin, G.gubn_nsyh,     ' +
                 '        G.gubn_tkgm, G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh,      ' +
                 '        G.gubn_agab, G.gubn_agyh, T.cod_prf, T.cod_chuga As Tcod_chuga FROM gumgin G with(nolock)         ' +
                 ' LEFT OUTER JOIN Bunju B with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                 ' LEFT OUTER JOIN Tkgum T with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';

      sWhere := ' Where G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And G.Dat_gmgn = ''' + edtBDate.Text + ''' ';

      if RadioGroup1.ItemIndex = 1 then
         sWhere := sWhere + ' And G.gubn_neawon = ''2'' ';

      if Trim(MEdt_SampS.Text) <> '' then
         sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + ''' ';
      if Trim(MEdt_SampE.Text) <> '' then
         sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + ''' ';

      sOrderBy := ' ORDER BY B.cod_jisa, CAST(G.num_samp AS INT) ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q71', 'Q', 'N');
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            sTemp := UF_HangmokCheck;

            if (pos('U001', sTemp) > 0) and (pos('U002', sTemp) > 0) and (pos('U003', sTemp) > 0) and (pos('U004', sTemp) > 0) and
               (pos('U005', sTemp) > 0) and (pos('U006', sTemp) > 0) and (pos('U007', sTemp) > 0) and (pos('U008', sTemp) > 0) and
               (pos('U009', sTemp) > 0) and (pos('U010', sTemp) > 0) then
            begin

               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v003[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v004[Index] := qryBunju.FieldByName('desc_name').AsString;
               if (CB_U033.Checked) and (pos('U033', sTemp) > 0) then UV_v005[Index] := 'U033'
               else UV_v005[Index] := '';
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

procedure TfrmLD4Q71.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q71.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q711 := TfrmLD4Q711.Create(Self);
     frmLD4Q711.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q711 := TfrmLD4Q711.Create(Self);
     frmLD4Q711.QuickRep.Print;
  end;
end;

function TfrmLD4Q71.UF_HangmokCheck : String;
var sHangmokList, vSQL_Text, NoHangmokList : string;
    iRet, i : Integer;
    vCod_chuga : Variant;
begin
   Result := ''; sHangmokList := ''; vSQL_Text := '';  NoHangmokList := '';
   i := 0; iRet := 0;
   
   //---- 장비 검사(신신)..
   qry_gulkwa.Active := false;
   qry_gulkwa.ParamByName('cod_pf').AsString := qryBunju.FieldByName('cod_Jangbi').AsString;
   qry_gulkwa.Active := True;
   if qry_gulkwa.RecordCount > 0 then
   begin
      while not qry_gulkwa.Eof do
      begin
         if pos(qry_gulkwa.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_gulkwa.FieldByName('cod_hm').AsString;
         qry_gulkwa.Next;
      end;
   end;

   //---- 혈액 검사..
   qry_gulkwa.Active := false;
   qry_gulkwa.ParamByName('cod_pf').AsString    :=  qryBunju.FieldByName('cod_hul').AsString;
   qry_gulkwa.Active := True;
   if qry_gulkwa.RecordCount > 0 then
   begin
      while not qry_gulkwa.Eof do
      begin
         if pos(qry_gulkwa.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_gulkwa.FieldByName('cod_hm').AsString;
         qry_gulkwa.Next;
      end;
   end;

   //---- Cancer 검사..
   qry_gulkwa.Active := false;
   qry_gulkwa.ParamByName('cod_pf').AsString    :=  qryBunju.FieldByName('cod_Cancer').AsString;
   qry_gulkwa.Active := True;
   if qry_gulkwa.RecordCount > 0 then
   begin
      while not qry_gulkwa.Eof do
      begin
         if pos(qry_gulkwa.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
            sHangmokList := sHangmokList + qry_gulkwa.FieldByName('cod_hm').AsString;
         qry_gulkwa.Next;
      end;
   end;

   //---- 추가 검사..
   for i := 1 to Length(qryBunju.FieldByName('cod_chuga').AsString) div 4 do
   begin
      if copy(copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4),1,1) = 'U' then
      begin
         if pos(copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4), sHangmokList) <= 0 then
            sHangmokList := sHangmokList + copy(qryBunju.FieldByName('cod_chuga').AsString, (i * 4) - 3, 4);
      end;
   end;

   // 노신유형Display
   if qryBunju.FieldByName('gubn_nosin').AsString = '1' then
   begin
      NoHangmokList := UF_No_Hangmok(copy(GV_sToday,1,4), '1', qryBunju.FieldByName('gubn_nsyh').AsInteger);
      for i := 1 to Length(NoHangmokList) div 4 do
      begin
         qry_Hangmok.Active := false;
         vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
         vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(NoHangmokList, (i * 4) - 3, 4) + '''';
         vSQL_Text := vSQL_Text +  '   and gubn_part = ''03''';

         qry_Hangmok.SQL.Clear;
         qry_Hangmok.SQL.Add(vSQL_Text);
         qry_Hangmok.Active := True;

         if qry_Hangmok.RecordCount > 0 then
         begin
            if pos(qry_Hangmok.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
               sHangmokList := sHangmokList + qry_Hangmok.FieldByName('cod_hm').AsString;
         end;
      end;
   end;

   //성인병유형Display
   if qryBunju.FieldByName('gubn_adult').AsString = '1' then
   begin
      NoHangmokList := UF_No_Hangmok(copy(GV_sToday,1,4), '4', qryBunju.FieldByName('gubn_adyh').AsInteger);
      for i := 1 to Length(NoHangmokList) div 4 do
      begin
         qry_Hangmok.Active := false;
         vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
         vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(NoHangmokList, (i * 4) - 3, 4) + '''';
         vSQL_Text := vSQL_Text +  '   and gubn_part = ''03''';

         qry_Hangmok.SQL.Clear;
         qry_Hangmok.SQL.Add(vSQL_Text);
         qry_Hangmok.Active := True;

         if qry_Hangmok.RecordCount > 0 then
         begin
            if pos(qry_Hangmok.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
               sHangmokList := sHangmokList + qry_Hangmok.FieldByName('cod_hm').AsString;
         end;
      end;
   end;

   //간염유형Display
   if qryBunju.FieldByName('gubn_agab').AsString = '1' then
   begin
      NoHangmokList := UF_No_Hangmok(copy(GV_sToday,1,4), '5', qryBunju.FieldByName('gubn_adyh').AsInteger);
      for i := 1 to Length(NoHangmokList) div 4 do
      begin
         qry_Hangmok.Active := false;
         vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
         vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(NoHangmokList, (i * 4) - 3, 4) + '''';
         vSQL_Text := vSQL_Text +  '   and gubn_part = ''03''';

         qry_Hangmok.SQL.Clear;
         qry_Hangmok.SQL.Add(vSQL_Text);
         qry_Hangmok.Active := True;

         if qry_Hangmok.RecordCount > 0 then
         begin
            if pos(qry_Hangmok.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
               sHangmokList := sHangmokList + qry_Hangmok.FieldByName('cod_hm').AsString;
         end;
      end;
   end;
   // 생애전환기유형Display
   if qryBunju.FieldByName('gubn_life').AsString = '1' then
   begin
      NoHangmokList := UF_No_Hangmok(copy(GV_sToday,1,4), '7', qryBunju.FieldByName('gubn_lfyh').AsInteger);
      for i := 1 to Length(NoHangmokList) div 4 do
      begin
         qry_Hangmok.Active := false;
         vSQL_Text :=              'select cod_hm, desc_kor From Hangmok';
         vSQL_Text := vSQL_Text +  ' where cod_hm    = ''' + copy(NoHangmokList, (i * 4) - 3, 4) + '''';
         vSQL_Text := vSQL_Text +  '   and gubn_part = ''03''';

         qry_Hangmok.SQL.Clear;
         qry_Hangmok.SQL.Add(vSQL_Text);
         qry_Hangmok.Active := True;

         if qry_Hangmok.RecordCount > 0 then
         begin
            if pos(qry_Hangmok.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
               sHangmokList := sHangmokList + qry_Hangmok.FieldByName('cod_hm').AsString;
         end;
      end;
   end;

   if (CB_tk.Checked) and
      ((qryBunju.FieldByName('gubn_tkgm').AsString = '1') or
       (qryBunju.FieldByName('gubn_tkgm').AsString = '2')) then
   begin
      //---- 특검Profile...
      iRet := GF_MulToSingle(qryBunju.FieldByName('cod_prf').AsString, 4, vCod_chuga);
      for i := 0 to iRet - 1 do
      begin
         qry_gulkwa.Active := false;
         qry_gulkwa.ParamByName('cod_pf').AsString    :=  vCod_chuga[i];
         qry_gulkwa.Active := True;
         if qry_gulkwa.RecordCount > 0 then
         begin
            while not qry_gulkwa.Eof do
            begin
               if pos(qry_gulkwa.FieldByName('cod_hm').AsString, sHangmokList) <= 0 then
                  sHangmokList := sHangmokList + qry_gulkwa.FieldByName('cod_hm').AsString;

               qry_gulkwa.Next;
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

function  TfrmLD4Q71.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q71.UP_Change(Sender: TObject);
begin
  inherited;
  MEdt_SampS.Text := '';
  MEdt_SampE.Text := '';
end;

end.
