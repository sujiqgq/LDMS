//==============================================================================
// 프로그램 설명 : 당일 응급 B렉 조회 리스트
// 시스템        : 통합검진
// 수정일자      : 2014.11.15
// 수정자        : 곽윤설
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD4Q63;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q63 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtDate: TDateEdit;
    SampNo1: TEdit;
    SampNo2: TEdit;
    qryGumgin: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    ComB_ShFloor: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryPf_hm2: TQuery;
    Panel2: TPanel;
    rb_nea: TRadioButton;
    rb_chul: TRadioButton;
    rb_All: TRadioButton;
    Label4: TLabel;
    qryPart0507_01: TQuery;
    qryProfile: TQuery;
    qryPart0507_02: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSamp, UV_vName, UV_vJumin, UV_vNeawon, UV_vFloor, Uv_vGubn : Variant;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
  end;

var
  frmLD4Q63: TfrmLD4Q63;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q631;

{$R *.DFM}

procedure TfrmLD4Q63.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo     := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp   := VarArrayCreate([0, 0], varOleStr);
      UV_vName   := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vNeawon := VarArrayCreate([0, 0], varOleStr);
      UV_vFloor  := VarArrayCreate([0, 0], varOleStr);
      Uv_vGubn   := VarArrayCreate([0, 0], varOleStr);
  end           
   else         
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vSamp,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vJumin,  iLength);
      VarArrayRedim(UV_vNeawon, iLength);
      VarArrayRedim(UV_vFloor,  iLength);
      VarArrayRedim(Uv_vGubn,   iLength);
   end;
end;

function TfrmLD4Q63.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q63.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   edtDate.Text := GV_sToday;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   GP_ComboFloor_Center(ComB_ShFloor);

   // 현재일자를 설정
   edtDate.Text := GV_sToday;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q63.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSamp[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vJumin[DataRow-1];
      5 : Value := UV_vNeawon[DataRow-1];
      6 : Value := UV_vFloor[DataRow-1];
      7 : Value := Uv_vGubn[DataRow-1];
   end;
end;

procedure TfrmLD4Q63.btnQueryClick(Sender: TObject);
var index, I : Integer;
    sSelect, sWhere, sOrderBy : String;
    sHul_List, sGubn_Neawon, sGubn_Floor : String;
    bunju_no : Integer;
    vCheck_01, vCheck_02, vCheck_03, sAll_Check : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGumgin do
   begin
      // SQL문 생성
      Close;

      sSelect := ' SELECT G.num_jumin, G.num_id, G.desc_name, G.dat_gmgn, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, G.num_samp, ' +
                 '        G.gubn_neawon, G.gubn_jinch, G.gubn_hulgum, G.gubn_nosin, G.gubn_adult, G.gubn_agab, G.gubn_life, G.Gubn_bogun, G.gubn_nsyh,  ' +
                 '        G.gubn_adyh, G.gubn_agyh, G.gubn_lfyh, G.gubn_tkgm, G.gubn_neawon, T.cod_prf, T.cod_chuga As Tcod_chuga         ' +
                 ' FROM Gumgin G with(nolock)                                                                                             ' +
                 ' left outer join Tkgum T with(nolock)                                                                                   ' +
                 ' ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha      ';
      sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND G.Dat_gmgn  = ''' + edtDate.Text + ''' ';
      sWhere  := sWhere + ' AND G.Num_samp  <> ''' + '' + ''' ';

      If SampNo1.Text <> '' Then
         sWhere := sWhere + ' And G.Num_samp >= ''' + SampNo1.Text + '''' +
                            ' And G.Num_samp <= ''' + SampNo2.Text + '''';

      IF ComB_ShFloor.Text <> '' then
         sWhere := sWhere + ' And G.gubn_jinch = ''' + copy(ComB_ShFloor.Text,1,2) + '''';

      if rb_nea.Checked then
         sWhere := sWhere + ' And G.gubn_neawon in (''1'',''3'',''4'')'
      else if rb_chul.Checked then
         sWhere := sWhere + ' And G.gubn_neawon = ''2''';
      sOrderBy := ' ORDER BY G.Num_samp';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD4Q63', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;

            sHul_List := UF_hangmokList;
            sAll_Check := False; vCheck_01 := False; vCheck_02 := False;
            //혈액Check
            If  qryGumgin.FieldByName('Cod_Hul').AsString <> '' Then
            With qryProfile do
               Begin
                  Close;
                  ParamByName('Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Hul').AsString;
                  Open;
                  If FieldByName('Gubn_Gmsa').AsString > '30' Then
                     sAll_Check := True;
               End
            //장비Check
            Else If  qryGumgin.FieldByName('Cod_Jangbi').AsString <> '' Then
            With qryProfile do
               Begin
                  Close;
                  ParamByName('Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Jangbi').AsString;
                  Open;
                  If FieldByName('Gubn_Gmsa').AsString > '30' Then
                     sAll_Check := True;
               End
            //종양Check
            Else If FieldByName('Cod_Cancer').AsString <> '' Then
               Begin
                  sAll_Check := True;
               End
            //노신
           Else if       (FieldByName('gubn_nosin').AsString >= '1')
                 or (FieldByName('gubn_tkgm').AsString >= '1')
                 or (FieldByName('gubn_bogun').AsString >= '1')
                 or (FieldByName('gubn_adult').AsString >= '1')
                 or (FieldByName('gubn_agab').AsString >= '1')
                 or (FieldByName('gubn_life').AsString >= '1') then
                 sAll_Check := True
          //추가코드 확인
            Else if       (Trim(FieldByName('cod_chuga').AsString) <> '')
                 and (FieldByName('gubn_nosin').AsString < '1')
                 and (FieldByName('gubn_tkgm').AsString  < '1')
                 and (FieldByName('gubn_bogun').AsString < '1')
                 and (FieldByName('gubn_adult').AsString < '1')
                 and (FieldByName('gubn_agab').AsString  < '1')
                 and (FieldByName('gubn_life').AsString  < '1') Then
                 sAll_Check := True;

            If sAll_Check Then
            begin
               With qryPart0507_01 do
               Begin
                  Close;
                  Open;
                  while not Eof do
                  begin
                     If pos(FieldByName('cod_hm').AsString, sHul_List) > 0 Then
                        vCheck_01 := True;
                     next;
                  end;
               End;
               With qryPart0507_02 do
               Begin
                  Close;
                  Open;
                  while not Eof do
                  begin
                     If pos(FieldByName('cod_hm').AsString, sHul_List) > 0 Then
                        vCheck_02 := True;
                     next;
                  end;
               End;
            end;

            if (vCheck_01) and (vCheck_02) Then
            begin
                UP_VarMemSet(Index);
                UV_vNo[Index]     := IntToStr(Index+1);
                UV_vSamp[Index]   := qryGumgin.FieldByName('num_samp').AsString;
                UV_vName[Index]   := qryGumgin.FieldByName('desc_name').AsString;
                UV_vJumin[Index]  := copy(qryGumgin.FieldByName('num_jumin').AsString,1,6)+'-'+copy(qryGumgin.FieldByName('num_jumin').AsString,7,13);//+'******';
                case qryGumgin.FieldByName('gubn_neawon').AsInteger of
                1 : begin
                       sGubn_Neawon := '내원(단체)';
                    end;
                2 : begin
                       sGubn_Neawon := '출장';
                    end;
                3 : begin
                       sGubn_Neawon := '내원';
                    end;
                4 : begin
                       sGubn_Neawon := '내원(진료(외래))';
                    end;
                5 : begin
                       sGubn_Neawon := '대체';
                    end;
                end;
                case qryGumgin.FieldByName('gubn_jinch').AsInteger of
                0 : begin
                       sGubn_Floor := '구분없음';
                    end;
                1 : begin
                       sGubn_Floor := '본원[종검-2F]';
                    end;
                2 : begin
                       sGubn_Floor := '본원[일반-3F]';
                    end;
                3 : begin
                       sGubn_Floor := '강남[종검-7F]';
                    end;
                4 : begin
                       sGubn_Floor := '강남[일반-6F]';
                    end;
                5 : begin
                       sGubn_Floor := '여의도[VIP-18F]';
                    end;
                6 : begin
                       sGubn_Floor := '여의도[종검-17F]';
                    end;
                7 : begin
                       sGubn_Floor := '여의도[종검-18F]';
                    end;
                8 : begin
                       sGubn_Floor := '여의도[일반-15F]';
                    end;
                9 : begin
                       sGubn_Floor := '강남[VIP-8F]';
                    end;
                10 : begin
                       sGubn_Floor := '본원[신신-4F]';
                    end;
                end;
                UV_vNeawon[Index] := sGubn_Neawon;
                UV_vFloor[Index]  := sGubn_Floor;
                inc(index);
            end;
            Next;
         End;
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

procedure TfrmLD4Q63.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtDate);
end;

procedure TfrmLD4Q63.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q631 := TfrmLD4Q631.Create(Self);
     frmLD4Q631.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q631 := TfrmLD4Q631.Create(Self);
     frmLD4Q631.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q63.FormActivate(Sender: TObject);
begin
   inherited;
   edtdate.SetFocus;
end;

function  TfrmLD4Q63.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function TfrmLD4Q63.UF_hangmokList : String;
var
    index, I, j, iRet, iTemp, num : Integer;
    sHmCode : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
begin
     // 항목(장비프로파일)추출..
     sHmCode := '';
     if qryGumgin.FieldByName('cod_jangbi').AsString <> '' then
     begin
        qryPf_hm.Active := False;
        qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_jangbi').AsString;
        qryPf_hm.Active := True;
        if qryPf_hm.RecordCount > 0 then
        begin
           while not qryPf_hm.Eof do
           begin
              if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                 sHmCode := sHmCode + qryPf_hm.FieldByName('COD_HM').AsString;
              qryPf_hm.Next;
           end;
        end;
     end;

     // 항목(혈액프로파일)추출..
     if qryGumgin.FieldByName('cod_hul').AsString <> '' then
     begin
        qryPf_hm.Active := False;
        qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_hul').AsString;
        qryPf_hm.Active := True;
        if qryPf_hm.RecordCount > 0 then
        begin
           while not qryPf_hm.Eof do
           begin
              if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                 sHmCode := sHmCode + qryPf_hm.FieldByName('COD_HM').AsString;
              qryPf_hm.Next;
           end;
        end;
     end;

     // 항목(종양프로파일)추출..
     if qryGumgin.FieldByName('cod_cancer').AsString <> '' then
     begin
        qryPf_hm.Active := False;
        qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_cancer').AsString;
        qryPf_hm.Active := True;
        if qryPf_hm.RecordCount > 0 then
        begin
           while not qryPf_hm.Eof do
           begin
              if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                 sHmCode := sHmCode + qryPf_hm.FieldByName('COD_HM').AsString;
              qryPf_hm.Next;
           end;
        end;
     end;

     //추가검사항목..
     iRet := GF_MulToSingle(qryGumgin.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
     for iTemp := 0 to iRet-1 do
     begin
        if copy(vCod_chuga[iTemp],1,2) <> 'JJ' then
           sHmCode := sHmCode + vCod_chuga[iTemp];
     end;

     if qryGumgin.FieldByName('gubn_tkgm').AsString <> '' then
     begin
        //---- 특검항목(Profile) 추출...
        iRet := GF_MulToSingle(qryGumgin.FieldByName('cod_prf').AsString, 4, vCod_profile);
        for j := 0 to iRet-1 do
        begin
           if Trim(vCod_profile[j]) <> '' then
           begin
              qryPf_hm2.Active := False;
              qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[j];
              qryPf_hm2.Active := True;
              if qryPf_hm2.RecordCount > 0 then
              begin
                 while not qryPf_hm2.Eof do
                 begin
                    check_tk := True;
                    for num := 1 to round(Length(sHmCode)/4) do
                    begin
                       if copy(sHmCode, (num * 4) - 3,4) =
                          qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                    end;
                    if check_tk = True then
                    begin
                       if copy(qryPf_hm2.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                          sHmCode := sHmCode + qryPf_hm2.FieldByName('COD_HM').AsString;
                    end;
                    qryPf_hm2.Next;
                 end; // end of while(항목 처리)
              end; // end of if
           end; //end of if
        end; //end of for

        //(특검)추가검사항목..
        iRet := GF_MulToSingle(qryGumgin.FieldByName('Tcod_chuga').AsString, 4, vCod_chuga);
        for iTemp := 0 to iRet-1 do
        begin
           if copy(vCod_chuga[iTemp],1,2) <> 'JJ' then
              sHmCode := sHmCode + vCod_chuga[iTemp];
        end;
     end;

     // 노신유형Display
     if qryGumgin.FieldByName('gubn_nosin').AsString = '1' then
        sHmCode := sHmCode + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qryGumgin.FieldByName('gubn_nsyh').AsInteger);

     //성인병유형Display
     if qryGumgin.FieldByName('gubn_adult').AsString = '1' then
        sHmCode := sHmCode + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qryGumgin.FieldByName('gubn_adyh').AsInteger);

     //간염유형Display
     if qryGumgin.FieldByName('gubn_agab').AsString = '1' then
        sHmCode := sHmCode + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qryGumgin.FieldByName('gubn_agyh').AsInteger);

     //생애전환기병유형Display
     if qryGumgin.FieldByName('gubn_life').AsString = '1' then
        sHmCode := sHmCode + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qryGumgin.FieldByName('gubn_lfyh').AsInteger);

     Result := sHmCode;
end;

procedure TfrmLD4Q63.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
  inherited;
//
end;

end.
