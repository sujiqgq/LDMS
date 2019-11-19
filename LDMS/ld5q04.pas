//==============================================================================
// 프로그램 설명 : [WList]항목리스트조회(ESR,뇨침사,혈액형)
// 시스템        : 분석관리
// 개발일자      : 2012.02.06
// 개발자        : 유동구
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
unit LD5Q04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD5Q04 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
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
    CB_H025: TCheckBox;
    CB_H023: TCheckBox;
    CB_U011: TCheckBox;
    qryCheck: TQuery;
    BitBtn1: TBitBtn;
    qryProfile: TQuery;
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
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUNJU, UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn, UV_vHangmok : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD5Q04: TfrmLD5Q04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q041, LD5Q042, LD5Q043, LD5Q044;

{$R *.DFM}

procedure TfrmLD5Q04.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD5Q04.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD5Q04.UP_VarMemSet(iLength : Integer);
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
   end;
end;

function TfrmLD5Q04.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q04.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD5Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
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
   end;
end;

procedure TfrmLD5Q04.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sCheck_List, sHul_List : String;
    vCod_chuga : Variant;
    bCkeck, sHulCheck : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

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
         sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, ' +
                    ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                    ' FROM Gulkwa K left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                    '               left outer join Danche D ON G.cod_dc = D.cod_dc ' +
                    '               left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';

         sWhere := ' Where K.Cod_bjjs in (''15'',''' + Copy(cmbJisa.Text,1,2) + ''')';
         sWhere := sWhere + ' and K.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
         sWhere := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' +
                            ' AND (K.Gubn_Part  = ''02'' OR K.Gubn_part = ''03'')';
         sWhere := sWhere + ' and G.gubn_hulgum <> ''1''';

         If BunjuNo1.Text <> '' Then
            sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                               ' And K.Num_Bunju <= ''' + BunjuNo2.Text + '''';
         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

         sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, ' +
                     ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga, G.num_samp ';

         case RadioGroup1.ItemIndex of
            0 : sOrderBy := ' ORDER BY K.Num_Bunju';
            1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         end;
      end
      //검진일자 기준..
      else if RBtn_Gmgn.Checked then
      begin
         sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, ' +
                    ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh,  G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                    ' FROM Gumgin G left outer join Danche D ON G.cod_dc = D.cod_dc ' +
                    '      left outer join Tkgum T ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';

         sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
         sWhere  := sWhere + ' AND G.Dat_gmgn = ''' + edtBDate.Text + '''';
         If BunjuNo1.Text <> '' Then
            sWhere := sWhere + ' And G.num_samp >= ''' + BunjuNo1.Text + '''' +
                               ' And G.num_samp <= ''' + BunjuNo2.Text + '''';
         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

         sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, ' +
                     ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga, G.num_samp ';

         sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
      end;


      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q32', 'Q', 'N');
         for I := 1 to RecordCount do
         begin
            sCheck_List := '';
            bCkeck := False;
            sHulCheck := False;
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

            //단체 Check
            With qryCheck Do
            Begin
               Close;
               ParamByName('In_Cod_dc').AsString := qryBunju.FieldByname('Cod_dc').AsString;
               ParamByName('num_year').AsString  := copy(qryBunju.FieldByname('DaT_Gmgn').AsString,1,4);
               Open;

               If Recordcount > 0 Then
                  If (qryBunju.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                     (qryBunju.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                     begin
                     bCkeck := True;
                     sHulCheck := True;
                     end;
               Close;
            End;

            //[2008.02.29] 검사추가에 있으면 검사들어감..
            if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
               (pos('H024',FieldByName('cod_chuga').AsString) > 0) then bCkeck := True;

            //어린이집
            If (FieldByname('Cod_hul').Asstring = 'B501') Or
               (FieldByname('Cod_hul').Asstring = 'B502') Or
               (FieldByname('Cod_hul').Asstring = '9023') Or
               (FieldByname('Cod_hul').Asstring = 'B507') Or
               (FieldByname('Cod_hul').Asstring = 'B508') Then bCkeck := True;


            If (Copy(FieldByname('Cod_Jangbi').Asstring,1,2) <> '') Then bCkeck := True;

            if RBtn_Bunju.Checked then
            begin
               If qryBunju.FieldByName('Num_Bunju').AsInteger > 4999  Then bCkeck := True;
            end;

            if (pos('H025',sHul_List) > 0) and (CB_H025.Checked) then sCheck_List := sCheck_List + 'ESR(H025),';
            if (pos('U011',sHul_List) > 0) and (CB_U011.Checked) then sCheck_List := sCheck_List + '뇨침사 WBC(U011),';
            if (pos('U012',sHul_List) > 0) and (CB_U011.Checked) then sCheck_List := sCheck_List + '뇨침사 RBC(U012),';
            if (pos('U013',sHul_List) > 0) and (CB_U011.Checked) then sCheck_List := sCheck_List + '상피세포(U013),';

            if (pos('H023',sHul_List) > 0) and (bCkeck) and (CB_H023.Checked) then sCheck_List := sCheck_List + 'ABO형(H023),';
            if (pos('H024',sHul_List) > 0) and (bCkeck) and (CB_H023.Checked) then sCheck_List := sCheck_List + 'Rh형(H024),';

            if sCheck_List <> '' then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]       := IntToStr(Index+1);
               if RBtn_Bunju.Checked then UV_VBUNJU[Index] := qryBunju.FieldByName('Num_Bunju').AsString
               else                       UV_VBUNJU[Index] := '-';
               UV_vName[Index]     := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[Index]    := qryBunju.FieldByName('num_jumin').AsString;
               UV_vJisa[Index]     := qryBunju.FieldByName('cod_jisa').AsString;
               UV_vDesc_dc[Index]  := qryBunju.FieldByName('desc_dc').AsString;
               UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
               UV_vDat_gmgn[Index] := qryBunju.FieldByName('Dat_gmgn').AsString;
               UV_vHangmok[Index]  := sCheck_List;

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

procedure TfrmLD5Q04.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD5Q04.UP_Click(Sender: TObject);
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
   else if Sender = RBtn_Bunju then
   begin
      Label2.Caption := '분 주 일 :';
      Label7.Caption := '분주지사 :';
      Label3.Caption := '분주번호 :';
      bunjuno1.MaxLength := 4;
      bunjuno2.MaxLength := 4;
      CB_H023.Enabled := True;
   end
   else if Sender = RBtn_Gmgn then
   begin
      Label2.Caption := '검 진 일 :';
      Label7.Caption := '검진지사 :';
      Label3.Caption := '샘플번호 :';
      bunjuno1.MaxLength := 6;
      bunjuno2.MaxLength := 6;
      CB_H023.Enabled := False;
   end;
end;

procedure TfrmLD5Q04.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD5Q041 := TfrmLD5Q041.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q041.QuickRep.Preview
  else                                frmLD5Q041.QuickRep.Print;
end;

procedure TfrmLD5Q04.FormActivate(Sender: TObject);
begin
   inherited;
   edtBdate.SetFocus;
end;

function TfrmLD5Q04.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
   sHulCheck : Boolean ;
begin
   result := ''; sTemp := '';

   //단체 Check
   sHulCheck := False;
   With qryCheck Do
   Begin
      Close;
      ParamByName('In_Cod_dc').AsString := qryBunju.FieldByname('Cod_dc').AsString;
      ParamByName('num_year').AsString  := copy(qryBunju.FieldByname('DaT_Gmgn').AsString,1,4);
      Open;

      If Recordcount > 0 Then
         If (qryBunju.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
            (qryBunju.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
            begin
            sHulCheck := True;
            end;
      Close;
   End;

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if (CB_H023.checked = True) and (qryBunju.FieldByName('cod_hul').AsString <> '') and (sHulCheck = True) then     //혈액형 검사 일경우 단체체크
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

      if (qryBunju.FieldByName('cod_hul').AsString <> '') and ((CB_H025.checked = True) or (CB_U011.checked = true))then
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

function TfrmLD5Q04.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q04.Btn_BacordPntClick(Sender: TObject);
begin
  inherited;
  frmLD5Q042 := TfrmLD5Q042.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q042.QuickRep.Preview
  else                                frmLD5Q042.QuickRep.Print;
end;

procedure TfrmLD5Q04.BitBtn1Click(Sender: TObject);
begin
  inherited;
  if frmLD5Q04.CB_H025.Checked or
     frmLD5Q04.CB_H023.Checked then
  begin
     frmLD5Q043 := TfrmLD5Q043.Create(Self);

     if RBtn_preveiw.Checked = True then frmLD5Q043.QuickRep.Preview
     else                                frmLD5Q043.QuickRep.Print;
  end
  else if frmLD5Q04.CB_U011.Checked then
  begin
     frmLD5Q044 := TfrmLD5Q044.Create(Self);

     if RBtn_preveiw.Checked = True then frmLD5Q044.QuickRep.Preview
     else                                frmLD5Q044.QuickRep.Print;
  end;
end;

end.
