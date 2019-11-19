//==============================================================================
// 프로그램 설명 : 검사관리
// 시스템        : 분석관리
// 개발일자      : 2015.02.06
// 개발자        : 구수정
// 참고사항      :
//==============================================================================
unit LD9Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD9Q01 = class(TfrmSingle)
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
    Label1: TLabel;
    RadioGroup1: TRadioGroup;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    Bevel2: TBevel;
    RBtn_Bunju: TRadioButton;
    RBtn_Gmgn: TRadioButton;
    Btn_BacordPnt: TBitBtn;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    qryProfile: TQuery;
    Cmb_gubn: TComboBox;
    Cmb_Hm: TComboBox;
    Label3: TLabel;
    qry_LdmsHM_GR: TQuery;
    qry_LdmsHm: TQuery;
    Label2: TLabel;
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
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUNJU, UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn : Variant;
    HM_Count : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD9Q01: TfrmLD9Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD9Q011, LD9Q012;

{$R *.DFM}

procedure TfrmLD9Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD9Q01.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //수정
      end;
   end;
end;

procedure TfrmLD9Q01.UP_VarMemSet(iLength : Integer);
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
   end;
end;

function TfrmLD9Q01.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') or (Cmb_Hm.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD9Q01.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);

   // 오늘일자로 기본설정
   edtBDate.Text := GV_sToday;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   //검사코드 리스트
   qry_LdmsHM_GR.Active := False;
   qry_LdmsHM_GR.ParamByName('PROGRAM_ID').AsString := 'LD9Q01';
   qry_LdmsHM_GR.Active := True;
   if qry_LdmsHM_GR.RecordCount > 0 then
   begin
      while not qry_LdmsHM_GR.Eof do
      begin
          Cmb_Hm.Items.Add(qry_LdmsHM_GR.FieldByName('group_hm').AsString);
          qry_LdmsHM_GR.Next;
      end;
   end;

   //분주번호로 조회
   Cmb_gubn.ItemIndex := 0;
   Cmb_Hm.ItemIndex := 0;
   bunjuno1.MaxLength := 6;
   bunjuno2.MaxLength := 6;

end;

procedure TfrmLD9Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
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
   end;
end;

procedure TfrmLD9Q01.btnQueryClick(Sender: TObject);
var index, I, J, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
    vArr_code : array of Variant;
    bCheck : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ' WHERE '; sOrderBy := '';
   Index := 0; J := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
      // SQL문 생성
      Close;

      //분주일자 기준..
      if RBtn_Bunju.Checked then
      begin
         sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, B.desc_rackno, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, Gubn_nsyh, ' +
                    '        Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp           ' +
                    ' FROM Gulkwa K with(nolock) left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn                                             ' +
                    '                            left outer join Danche D ON G.cod_dc = D.cod_dc                                                                                                     ' +
                    '                            left outer join Bunju B  ON K.cod_jisa = B.cod_jisa and K.num_id = B.num_id and K.dat_gmgn = B.dat_gmgn                                             ' +
                    '                            left outer join Tkgum T  ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha          ';

         case RadioGroup1.ItemIndex of
            0 : sWhere := sWhere + '  K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
            1 : sWhere := sWhere + '  G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
         end;

         sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

         if Cmb_gubn.ItemIndex = 0 then
         begin
            if Trim(bunjuno1.Text) <> '' then
               sWhere := sWhere + ' AND G.num_samp >= ''' + bunjuno1.Text + ''' ';
            if Trim(bunjuno2.Text) <> '' then
               sWhere := sWhere + ' AND G.num_samp <= ''' + bunjuno2.Text + ''' ';
         end
         else if Cmb_gubn.ItemIndex = 1 then
         begin
            if Trim(bunjuno1.Text) <> '' then
               sWhere := sWhere + ' AND K.num_bunju >= ''' + bunjuno1.Text + ''' ';
            if Trim(bunjuno2.Text) <> '' then
               sWhere := sWhere + ' AND K.num_bunju <= ''' + bunjuno2.Text + ''' ';
         end
         else if Cmb_gubn.ItemIndex = 2 then
         begin
            if Trim(bunjuno1.Text) <> '' then
               sWhere := sWhere + ' AND B.desc_rackno like ''' + copy(bunjuno1.Text,1,1) + ':' + copy(bunjuno1.Text,2,2) + '%'' ';
         end;

         sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, B.desc_rackno, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                     '          Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga, G.num_samp   ';

         case RadioGroup1.ItemIndex of
            0 : sOrderBy := ' ORDER BY K.Num_Bunju ';
            1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         end;
      end
      //검진일자 기준..
      else if RBtn_Gmgn.Checked then
      begin
         sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, Gubn_nsyh, ' +
                    '        Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp         ' +
                    ' FROM Gumgin G left outer join Danche D ON G.cod_dc = D.cod_dc                                                                                                   ' +
                    '               left outer join Tkgum T ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha         ';

         sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
         sWhere  := sWhere + ' AND G.Dat_gmgn = ''' + edtBDate.Text + '''';

         if Cmb_gubn.ItemIndex = 1 then
         begin
            if Trim(bunjuno1.Text) <> '' then
               sWhere := sWhere + ' AND K.num_bunju >= ''' + bunjuno1.Text + ''' ';
            if Trim(bunjuno2.Text) <> '' then
               sWhere := sWhere + ' AND K.num_bunju <= ''' + bunjuno2.Text + ''' ';
         end;

         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

         sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                     '          Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga, G.num_samp ';

         case RadioGroup1.ItemIndex of
            0 : sOrderBy := ' ';
            1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         end;
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD9Q01', 'Q', 'N');

         //검사코드 추출
         qry_LdmsHm.Active := False;
         qry_LdmsHm.ParamByName('PROGRAM_ID').AsString := 'LD9Q01';
         qry_LdmsHm.ParamByName('GROUP_HM').AsString   := Cmb_Hm.Text;
         qry_LdmsHm.Active := True;
         if qry_LdmsHm.RecordCount > 0 then
         begin
            HM_Count := qry_LdmsHm.RecordCount;
            SetLength(vArr_code, HM_Count);
            while not qry_LdmsHm.Eof do
            begin
                vArr_code[J] := qry_LdmsHm.FieldByName('cod_hm').AsString;
                inc(J);
                qry_LdmsHm.Next;
            end;
         end;

         for I := 1 to RecordCount do
         begin
            Gride.Progress := I;
            bCheck := False;
            sHul_List := UF_hangmokList;

            for J := 0 to HM_Count-1 do
               if pos(vArr_code[J], sHul_List) > 0 then bCheck := True;

            if bCheck then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]    := IntToStr(Index+1);
               if RBtn_Gmgn.Checked then UV_vBUNJU[Index] :=  ''
               else UV_vBUNJU[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               UV_vName[Index]  := qryBunju.FieldByName('desc_name').AsString;
               case StrToInt(copy(qryBunju.FieldByName('num_jumin').AsString,7,1)) of
                  1,3,5,7,9 : UV_vJumin[Index] := 'M/' + copy(qryBunju.FieldByName('num_jumin').AsString,1,6);
                  2,4,6,8   : UV_vJumin[Index] := 'F/' + copy(qryBunju.FieldByName('num_jumin').AsString,1,6);
               end;
               UV_vJisa[Index]     := qryBunju.FieldByName('cod_jisa').AsString;
               UV_vDesc_dc[Index]  := qryBunju.FieldByName('desc_dc').AsString;
               UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
               UV_vDat_gmgn[Index] := qryBunju.FieldByName('Dat_gmgn').AsString;

               Inc(Index);
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
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD9Q01.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end
   else if Sender = Cmb_gubn then
   begin
      bunjuno1.text    := '';
      bunjuno2.text    := '';
      bunjuno2.Visible := True;
      Label5.Visible   := True;
      if Cmb_gubn.text = '분주번호' then
      begin
         bunjuno1.MaxLength := 4;
         bunjuno2.MaxLength := 4;
      end
      else if Cmb_gubn.text = '샘플번호' then
      begin
         bunjuno1.MaxLength := 6;
         bunjuno2.MaxLength := 6;
      end
      else
      begin
         bunjuno2.Visible   := False;
         Label5.Visible     := False;
         bunjuno1.MaxLength := 3;
      end;
   end;

end;

procedure TfrmLD9Q01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end
   else if Sender = btnBdate then GF_CalendarClick(edtBDate)
   else if Sender = RBtn_Bunju then Label2.Caption := '분주지사'
   else if Sender = RBtn_Gmgn then  Label2.Caption := '검진지사';
end;

procedure TfrmLD9Q01.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD9Q011 := TfrmLD9Q011.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD9Q011.QuickRep.Preview
  else                                frmLD9Q011.QuickRep.Print;
end;

procedure TfrmLD9Q01.Btn_BacordPntClick(Sender: TObject);
begin
  inherited;
  if HM_Count = 1 then
  begin
     frmLD9Q012 := TfrmLD9Q012.Create(Self);
     if RBtn_preveiw.Checked = True then frmLD9Q012.QuickRep.Preview
     else                                frmLD9Q012.QuickRep.Print;
  end
  else       Showmessage('하나의 검사항목에 대해서만 출력 가능합니다.');
end;


procedure TfrmLD9Q01.FormActivate(Sender: TObject);
begin
   inherited;
   edtBdate.SetFocus;
end;

function TfrmLD9Q01.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
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
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
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
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
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
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
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
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   result := sTemp;
end;

function TfrmLD9Q01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
