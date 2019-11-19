//================================================================================
// 수정일자      : 2015.06.18
// 수정자        : 곽윤설
// 수정내용      : 신규
// 참고사항      : 한의 재단진단검사의학팀1500075 (본원-한두례, LD4Q36 참고 개발 요청)
//==============================================================================

unit LD4Q66;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q66 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label4: TLabel;
    Cmb_gubn: TComboBox;
    qryJHangmokList: TQuery;
    group_radio: TGroupBox;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Edit_RackS: TMaskEdit;
    Edit_RackE: TMaskEdit;
    qryProfile: TQuery;
    qryHangmok: TQuery;
    RG_01: TRadioGroup;
    Chk_All: TCheckBox;
    Chk_U001: TCheckBox;
    Chk_Y004: TCheckBox;
    Chk_U017: TCheckBox;
    Chk_U024: TCheckBox;
    Chk_U033: TCheckBox;
    Chk_Sub: TCheckBox;
    Chk_Not_Used: TCheckBox;
    Chk_Y010: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure UP_SetInitGrid;
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vBUN, UV_vSampNo : Variant;
    vArr_HM : array of Variant;
    sU001, sY004, sU017, sU024, sU033, sSub, sY010, sNot, sSelf, sHM_List : String;
    J : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    iArr : Integer;
    vArr_HulHM : array of array of String;
  end;

var
  frmLD4Q66: TfrmLD4Q66;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q661;

{$R *.DFM}

procedure TfrmLD4Q66.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q66.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q66.UP_VarMemSet(iLength : Integer);
var J : Integer;
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo     := VarArrayCreate([0, 0], varOleStr);
      UV_vName   := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN    := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo := VarArrayCreate([0, 0], varOleStr);
      for J := 0 to iArr-1 do
         vArr_HM[J] := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vBUN,    iLength);
      VarArrayRedim(UV_vSampNo, iLength);
      for J := 0 to iArr-1 do
         VarArrayRedim(vArr_HM[J], iLength);
   end;
end;

function TfrmLD4Q66.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q66.FormCreate(Sender: TObject);
begin
   inherited;
   sU001 := 'U001U002U003U004U005U006U007U008U009U010U011U012U013';
   sY004 := 'Y004';
   sU017 := 'U017U046';
   sU024 := 'U024U029U030U031U032U044U051';
   sU033 := 'U033';
   sY010 := 'Y010';
   sSub  := 'U020U023U026U028U037U043U045U047U050';
   sNot  := 'C066JJB3T003U014U015U016U018U027U035U036U038U039' +
            'U040U041U042U052U053U054U055U056U057U058Y007Y009';
   sSelf := 'U001U002U003U004U005U006U007U008U009U010U011U012U013';
            
   UP_SetInitGrid;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q66.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
       1 : Value := UV_vNo[DataRow-1];
       2 : Value := UV_vSampNo[DataRow-1];
       3 : Value := UV_vBUN[DataRow-1];
       4 : Value := UV_vName[DataRow-1];
   end;
   if  DataCol >= 5 then
   begin
      Value := vArr_HM[DataCol-5][DataRow-1];
      grdMaster.Col[DataCol].Heading := vArr_HulHM[DataCol-5][0];
   end;
end;

procedure TfrmLD4Q66.btnQueryClick(Sender: TObject);
var index, I, J, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy, sHul_List,
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List,
    UV_fGulkwa_07, UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07  : String;
    vCod_chuga : Variant;
    Check, Check_HM :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_SetInitGrid;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  B.desc_rackno, B.num_Bunju, G.num_jumin, G.num_samp, G.desc_name, G.num_id, G.dat_gmgn,           ' +
                 '         G.cod_jisa, G.gubn_nosin, G.gubn_adult,G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,       ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul,         ' +
                 '         G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga FROM bunju B with(nolock)         ' +
                 ' Join gumgin G with(nolock) on B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                 ' Join Gulkwa K with(nolock) on B.cod_bjjs = K.cod_bjjs and B.num_id = K.num_id and B.dat_bunju=K.dat_bunju ' +
                 ' Left Outer Join Tkgum T with (nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and        ' +
                 ' G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha                                                      ';

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
      sWhere  := sWhere + ' AND K.gubn_part  = ''03'' ';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
      end
      else if Cmb_gubn.Text = '샘플번호' then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else    // rack No.
      begin
         if Trim(Edit_RackS.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno >= ''' + Edit_RackS.Text + '''';
         if Trim(Edit_RackE.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno <= ''' + Edit_RackE.Text + '''';
      end;

      if RG_01.ItemIndex = 0 then
         sOrderBy := ' ORDER BY G.cod_jisa, G.num_samp '
      else if RG_01.ItemIndex = 1 then
         sOrderBy := ' ORDER BY B.num_bunju ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q66', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            sHul_List := '';
            Check := False;
            Check_HM := False;

            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;

            sHul_List := UF_hangmokList;

            For J := 0 to iArr-1 do
            begin
               if pos(vArr_HulHM[J][0], sHul_List) > 0 then Check_HM := True;
               if Check_HM then Break;
            end;

            if Check_HM then
            begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                  UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                  for J := 0 to iArr-1 do
                  begin
                     if (pos(vArr_HulHM[J][0], sHul_List) > 0) and (Trim(Copy(UV_fGulkwa, StrToInt(vArr_HulHM[J][1]),  6)) = '') then
                     begin
                        if pos(vArr_HulHM[J][0], sSelf) > 0 then
                        begin
                           if qryBunju.FieldByName('Cod_jisa').AsString = '15' then
                           begin
                              vArr_HM[J][Index] := '결과無';
                              vArr_HulHM[J][2] := IntToStr(StrToInt(vArr_HulHM[J][2]) + 1);
                              Check := True;
                           end;
                        end
                        else
                        begin
                           vArr_HM[J][Index] := '결과無';
                           vArr_HulHM[J][2] := IntToStr(StrToInt(vArr_HulHM[J][2]) + 1);
                           Check := True;
                        end;
                     end;
                  end;

                  if Check = True then
                  begin
                     UP_VarMemSet(Index+1);
                     UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                     UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                     UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                     UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
                     Inc(Index);
                  end;

               qryGulkwa.Close;
               end;          // end of qryGulkwa.RecordCount
            end;             // end of if(Check_HM)
            qryBunju.Next;
         End;                // end of qryBunju.Eof
         qryBunju.Close;     // Table과 Disconnect
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
         exit;
      end; // end of if qryBunju.RecordCount
   end;    // end of with(qry_gulkwa)

   // Grid에 자료를 할당
   UP_VarMemSet(Index+1);
   UV_vNo[Index]     := '';
   UV_VBUN[Index]    := '';
   UV_vName[Index]   := '총  합  계';
   UV_vSampNo[Index] := '';
   for J := 0 to iArr-1 do
      vArr_HM[J][Index] := vArr_HulHM[J][2];
   Inc(index);

   grdMaster.cols := 4 + iArr;
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q66.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate)
   else if Sender = Chk_All then
   begin
      if Chk_All.Checked then
      begin
         Chk_U001.Enabled := False;
         Chk_Y004.Enabled := False;
         Chk_U017.Enabled := False;
         Chk_U024.Enabled := False;
         Chk_U033.Enabled := False;
         Chk_U001.Checked := False;
         Chk_Y004.Checked := False;
         Chk_U017.Checked := False;
         Chk_U024.Checked := False;
         Chk_U033.Checked := False;
         Chk_Y010.Checked := False;
      end
      else
      begin
         Chk_U001.Enabled := True;
         Chk_Y004.Enabled := True;
         Chk_U017.Enabled := True;
         Chk_U024.Enabled := True;
         Chk_U033.Enabled := True;
         Chk_Y010.Enabled := True;
      end;
   end;
end;

procedure TfrmLD4Q66.btnPrintClick(Sender: TObject);
begin
  inherited;

  frmLD4Q661 := TfrmLD4Q661.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD4Q661.QuickRep.Preview
  else                                frmLD4Q661.QuickRep.Print;

end;

procedure TfrmLD4Q66.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;
end;

function TfrmLD4Q66.UF_hangmokList : String;
var
   i ,k, j : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
       sHmCode :string;
   vCod_chuga : Variant;
   iRet : integer;
begin
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

   for k:=0 to  Round(Length(sTemp)/4)-1 do
   begin
   result:=result+Copy(sTemp, (k+1)*4-3, 4)+'.';
   end;

end;

function TfrmLD4Q66.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q66.UP_SetInitGrid;
var J : Integer;
begin
   //배열초기화
   vArr_HM := nil;
   vArr_HulHM := nil;
   sHM_List := '';

   if Chk_All.Checked  then sHM_List := sHM_List + sU001 + sY004 + sU017 + sU024 + sU033 +sY010;
   if Chk_U001.Checked then sHM_List := sHM_List + sU001;
   if Chk_Y004.Checked then sHM_List := sHM_List + sY004;
   if Chk_U017.Checked then sHM_List := sHM_List + sU017;
   if Chk_U024.Checked then sHM_List := sHM_List + sU024;
   if Chk_U033.Checked then sHM_List := sHM_List + sU033;
   if Chk_Y010.Checked then sHM_List := sHM_List + sY010;
   if Chk_Sub.Checked  then sHM_List := sHM_List + sSub;
   if Chk_Not_Used.Checked then sHM_List := sHM_List + sNot;

   iArr := Round(length(sHM_List)/4);
   SetLength(vArr_HM, iArr);
   SetLength(vArr_HulHM, iArr);

   qryHangmok.Active := False;
   qryHangmok.Filtered := False;
   qryHangmok.ParamByName('dat_apply').AsString := edtBDate.Text;
   qryHangmok.Active := True;
   if qryHangmok.RecordCount > 0 then
   begin
      qryHangmok.Filtered := True;
      for J := 0 to iArr-1 do
      begin
         qryHangmok.Filter := ' cod_hm = ''' + Copy(sHM_List, (J*4)+1, 4) + ''' ';
         SetLength(vArr_HulHM[J], 3);
         vArr_HulHM[J][0] := Copy(sHM_List, (J*4)+1, 4); //qryHangmok.FieldByName('cod_hm').AsString;
         vArr_HulHM[J][1] := qryHangmok.FieldByName('pos_start').AsString;
         vArr_HulHM[J][2] := '0'
      end;
      qryHangmok.Filtered := False;
   end;
   qryHangmok.Active := False;
end;

end.
