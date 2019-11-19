//==============================================================================
// 프로그램 설명 : 혈액형분주현황
// 시스템        : 통합검진
// 수정일자      : 2002.08.22
// 수정자        : 최지혜
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.07.09
// 수정자        : 곽윤설
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q14;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q14 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    qryURINE: TQuery;
    Chk_15: TCheckBox;
    Chk_12: TCheckBox;
    Chk_11: TCheckBox;
    Chk_43: TCheckBox;
    Chk_71: TCheckBox;
    Chk_61: TCheckBox;
    Chk_51: TCheckBox;
    Gride: TGauge;
    Chk_45: TCheckBox;
    Chk_52: TCheckBox;
    Chk_34: TCheckBox;
    Chk_41: TCheckBox;
    Chk_Etc: TCheckBox;
    qryTkgum: TQuery;
    qry_gulkwa: TQuery;
    qry_Hangmok: TQuery;
    qryNo_hangmok: TQuery;
    CheckBox: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007 : Variant;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;    
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q14: TfrmLD4Q14;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q141;

{$R *.DFM}

procedure TfrmLD4Q14.UP_VarMemSet(iLength : Integer);
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
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,    iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
   end;
end;

function TfrmLD4Q14.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q14.FormCreate(Sender: TObject);
begin
   inherited;

   if dmComm.DBURINE.Connected = False then  dmComm.DBURINE.Connected := True;

   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Cmb_gubn.ItemIndex := 0;
   edtBDate.Text := GV_sToday;

   if      copy(GV_sUserId,1,2)='15' then Chk_15.Checked := True
   else if copy(GV_sUserId,1,2)='12' then Chk_12.Checked := True
   else if copy(GV_sUserId,1,2)='11' then Chk_11.Checked := True
   else if copy(GV_sUserId,1,2)='43' then Chk_43.Checked := True
   else if copy(GV_sUserId,1,2)='71' then Chk_71.Checked := True
   else if copy(GV_sUserId,1,2)='61' then Chk_61.Checked := True
   else if copy(GV_sUserId,1,2)='51' then Chk_51.Checked := True;
end;

procedure TfrmLD4Q14.grdMasterCellLoaded(Sender: TObject; DataCol,
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
      7 : Value := UV_v007[DataRow-1];
   end;
end;

procedure TfrmLD4Q14.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sHangmok, sTemp : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sCheck := '';  sHangmok := ''; 
   index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select B.cod_bjjs, B.num_id, B.dat_bunju, B.num_bunju, B.cod_jisa, B.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.cod_jisa, G.desc_name, G.num_jumin, ' +
                 '       G.Cod_jangbi, G.cod_hul,   G.Cod_cancer, G.cod_chuga, G.gubn_neawon, ' +
                 '       G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm, G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, ' +
                 '       G.gubn_agab,  G.gubn_agyh, G.num_jumin, T.cod_prf,    T.cod_chuga As Tcod_chuga ' +
                 ' FROM gulkwa B LEFT OUTER JOIN gumgin G ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Tkgum T ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';
      if Copy(cmbJisa.Text,1,2) = '15' then  sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' '
      else                                   sWhere := ' Where B.Cod_bjjs in (''15'',''' + Copy(cmbJisa.Text,1,2) + ''')';

      sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
      sWhere := sWhere + ' And B.gubn_part = ''03''';

      if Chk_15.Checked then sCheck := sCheck + '''15'',';
      if Chk_12.Checked then sCheck := sCheck + '''12'',';
      if Chk_11.Checked then sCheck := sCheck + '''11'',';
      if Chk_43.Checked then sCheck := sCheck + '''43'',';
      if Chk_61.Checked then sCheck := sCheck + '''61'',';
      if Chk_51.Checked then sCheck := sCheck + '''51'',';
      if Chk_71.Checked then sCheck := sCheck + '''71'',';
      if Chk_45.Checked then sCheck := sCheck + '''45'',';
      if Chk_52.Checked then sCheck := sCheck + '''52'',';
      if Chk_34.Checked then sCheck := sCheck + '''34'',';
      if Chk_41.Checked then sCheck := sCheck + '''41'',';
      if Chk_Etc.Checked then sCheck := sCheck + '''62'',''80'',''80'',';

      if sCheck <> '' then
         sWhere := sWhere + ' And B.cod_jisa in (' + copy(sCheck,1, Length(sCheck)-1) + ')';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY B.cod_jisa, CAST(B.num_bunju AS INT)';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY B.cod_jisa, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q14', 'Q', 'N');
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            sTemp := UF_HangmokCheck;

            if (pos('U011', sTemp) > 0) or
               (pos('U012', sTemp) > 0) or
               (length(sTemp) > 20) or
               ((qryBunju.FieldByName('cod_jisa').AsString = '15') and (qryBunju.FieldByName('gubn_neawon').AsString = '2') and (sTemp <> 'U017')) then
            begin
               qryURINE.Close;
               sSelect := ''; sWhere := ''; sOrderBy := '';
               sSelect := ' select * FROM result ';

               sWhere := ' Where gmgnDate = ''' + qryBunju.FieldByName('dat_gmgn').AsString + '''';
               sWhere := sWhere + ' and sampNumber  = ''' + qryBunju.FieldByName('num_samp').AsString + '''';

               qryURINE.SQL.Clear;
               qryURINE.SQL.Add(sSelect + sWhere);
               qryURINE.Open;

               if (pos('U001', sTemp) > 0) and (pos('U002', sTemp) > 0) and (pos('U003', sTemp) > 0) and (pos('U004', sTemp) > 0) and
                  (pos('U005', sTemp) > 0) and (pos('U006', sTemp) > 0) and (pos('U007', sTemp) > 0) and (pos('U008', sTemp) > 0) and
                  (pos('U009', sTemp) > 0) and (pos('U010', sTemp) > 0) then
               begin
                  if qryURINE.RecordCount <= 0 then
                  begin
                      UP_VarMemSet(Index);
                      UV_v001[Index] := IntToStr(Index+1);
                      UV_v002[Index] := qryBunju.FieldByName('cod_jisa').AsString;
                      UV_v003[Index] := qryBunju.FieldByName('num_samp').AsString;
                      UV_v004[Index] := qryBunju.FieldByName('num_bunju').AsString;
                      UV_v005[Index] := qryBunju.FieldByName('desc_name').AsString;
                      UV_v006[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
                      UV_v007[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

                      INC(Index);
                  end;
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

procedure TfrmLD4Q14.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
   if Sender = CheckBox then
   begin
      if CheckBox.Checked then
      begin
         CheckBox.Caption := 'UnCheck All';
         Chk_15.Checked := True; Chk_12.Checked := True;
         Chk_11.Checked := True; Chk_43.Checked := True;
         Chk_71.Checked := True; Chk_61.Checked := True;
         Chk_51.Checked := True; Chk_45.Checked := True;
         Chk_52.Checked := True; Chk_34.Checked := True;
         Chk_41.Checked := True; Chk_Etc.Checked := True;
      end
      else
      begin
         CheckBox.Caption := 'Check All';
         Chk_15.Checked := False; Chk_12.Checked  := False;
         Chk_11.Checked := False; Chk_43.Checked  := False;
         Chk_71.Checked := False; Chk_61.Checked  := False;
         Chk_51.Checked := False; Chk_45.Checked  := False;
         Chk_52.Checked := False; Chk_34.Checked  := False;
         Chk_41.Checked := False; Chk_Etc.Checked := False;
      end;
   end;
end;

procedure TfrmLD4Q14.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q141 := TfrmLD4Q141.Create(Self);
     frmLD4Q141.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q141 := TfrmLD4Q141.Create(Self);
     frmLD4Q141.QuickRep.Print;
  end;
end;

function TfrmLD4Q14.UF_HangmokCheck : String;
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

   if (qryBunju.FieldByName('gubn_tkgm').AsString = '1') or
      (qryBunju.FieldByName('gubn_tkgm').AsString = '2') then
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

function  TfrmLD4Q14.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.
