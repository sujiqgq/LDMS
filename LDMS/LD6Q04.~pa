//==============================================================================
// 프로그램 설명 : 혈액학 분주현황 (Absolute Count)
// 시스템        : 통합검진
// 수정일자      : 2015.04.27
// 수정자        : 곽윤설
// 수정내용      :
// 참고사항      : [한의 본원진료팀1500211/본원-김소영]
//==============================================================================
// 수정일자      : 2016.03.21
// 수정자        : 박수지
// 수정내용      : UV_V006 없을시 결과 값이 안들어가 프로그램 에러
// 참고사항      : 본원 김소영 의사
//==============================================================================
// 수정일자      : 2016.03.24
// 수정자        : 박수지
// 수정내용      : 성별 추가
// 참고사항      : 한의 본원진료팀1600220
//==============================================================================


unit LD6Q04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD6Q04 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    qryProfile: TQuery;
    Gauge2: TGauge;
    SBut_Excel: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);

    procedure btnPrintClick(Sender: TObject);
     procedure MouseWheelHandler(var Message: TMessage); override;

    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008,
    UV_v009, UV_v010, UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016, UV_vSex : Variant;
    sMan : String;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD6Q04: TfrmLD6Q04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;//, LD6Q041;

{$R *.DFM}

procedure TfrmLD6Q04.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD6Q04.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD6Q04.UP_VarMemSet(iLength : Integer);
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
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
      UV_v013 := VarArrayCreate([0, 0], varOleStr);
      UV_v014 := VarArrayCreate([0, 0], varOleStr);
      UV_v015 := VarArrayCreate([0, 0], varOleStr);
      UV_v016 := VarArrayCreate([0, 0], varOleStr);
      UV_vSex := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,  iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
      VarArrayRedim(UV_v008,  iLength);
      VarArrayRedim(UV_v009,  iLength);
      VarArrayRedim(UV_v010,  iLength);
      VarArrayRedim(UV_v011,  iLength);
      VarArrayRedim(UV_v012,  iLength);
      VarArrayRedim(UV_v013,  iLength);
      VarArrayRedim(UV_v014,  iLength);
      VarArrayRedim(UV_v015,  iLength);
      VarArrayRedim(UV_v016,  iLength);
      VarArrayRedim(UV_vSex,  iLength);
   end;
end;

function TfrmLD6Q04.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD6Q04.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD6Q04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1  : Value := UV_v001[DataRow-1];
      2  : Value := UV_v002[DataRow-1];
      3  : Value := UV_vSex[DataRow-1];
      4  : Value := UV_v003[DataRow-1];
      5  : Value := UV_v004[DataRow-1];
      6  : Value := UV_v005[DataRow-1];
      7  : Value := UV_v006[DataRow-1];
      8  : Value := UV_v007[DataRow-1];
      9  : Value := UV_v008[DataRow-1];
      10 : Value := UV_v009[DataRow-1];
      11 : Value := UV_v010[DataRow-1];
      12 : Value := UV_v011[DataRow-1];
      13 : Value := UV_v012[DataRow-1];
      14 : Value := UV_v013[DataRow-1];
      15 : Value := UV_v014[DataRow-1];
      16 : Value := UV_v015[DataRow-1];
      17 : Value := UV_v016[DataRow-1];
   end;
end;

procedure TfrmLD6Q04.btnQueryClick(Sender: TObject);
var index, i,viRet, iTemp, num : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
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

      sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, G.dat_bunju, G.num_bunju, G.gubn_part, G.DESC_GLKWA1, G.DESC_GLKWA2, G.DESC_GLKWA3,  ' +
                 '        A.num_samp, A.desc_name, A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life,  A.Gubn_lfyh, A.num_jumin, ' +
                 '        A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga   ';

      sSelect := sSelect + ' From gulkwa G with(nolock) left outer join Gumgin A with(nolock) ';
      sSelect := sSelect + ' On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';

      sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                         ' AND  G.Gubn_Part = ''02'' ';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY CAST(A.num_samp AS INT)';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(A.num_samp AS INT)'
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD6Q04', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

            if (pos('H001',sHul_List) > 0) or (pos('H002',sHul_List) > 0) or (pos('H008',sHul_List) > 0) or
               (pos('H011',sHul_List) > 0) or (pos('H012',sHul_List) > 0) or (pos('H014',sHul_List) > 0) or
               (pos('H015',sHul_List) > 0) or (pos('H016',sHul_List) > 0) or (pos('H017',sHul_List) > 0) then
            begin
               UP_VarMemSet(Index+1);

               UV_v001[Index] := FieldByName('dat_bunju').AsString;
               UV_v002[Index] := FieldByName('num_bunju').AsString;
               case StrToInt(Copy(FieldByName('num_jumin').AsString, 7, 1)) of
               1,3,5,7,9 : sMan := 'M';
               2,4,6,8   : sMan := 'F';
               end;
               UV_vSex[Index] := sMan;
               UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

               if pos('H001',sHul_List) > 0 then
                  UV_v003[Index] := Trim(Copy(UV_fGulkwa,  1,  6));
               if pos('H002',sHul_List) > 0 then
                  UV_v004[Index] := Trim(Copy(UV_fGulkwa,  7,  6));
               if pos('H008',sHul_List) > 0 then
                  UV_v005[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
               if pos('H011',sHul_List) > 0 then
                  UV_v006[Index] := Trim(Copy(UV_fGulkwa, 61,  6));
               if pos('H012',sHul_List) > 0 then
               begin
                  UV_v007[Index] := Trim(Copy(UV_fGulkwa, 67,  6));
                  if (UV_v006[Index] <> '') and (UV_v007[Index] <> '') then
                     UV_v008[Index] := StrToFloat(UV_v007[Index]) * StrToFloat(UV_v006[Index]);
               end;
               if pos('H014',sHul_List) > 0 then
               begin
                  UV_v009[Index] := Trim(Copy(UV_fGulkwa, 79,  6));
                  if (UV_v009[Index] <> '') and (UV_v007[Index] <> '') and (UV_v006[Index] <> '') then
                     UV_v010[Index] := StrToFloat(UV_v009[Index]) * StrToFloat(UV_v006[Index]);
               end;
               if pos('H015',sHul_List) > 0 then
               begin
                  UV_v011[Index] := Trim(Copy(UV_fGulkwa, 85,  6));
                  if (UV_v011[Index] <> '') and (UV_v007[Index] <> '') and (UV_v006[Index] <> '')  then
                     UV_v012[Index] := StrToFloat(UV_v011[Index]) * StrToFloat(UV_v006[Index]);
               end;
               if pos('H016',sHul_List) > 0 then
               begin
                  UV_v013[Index] := Trim(Copy(UV_fGulkwa, 91,  6));
                  if (UV_v013[Index] <> '') and (UV_v007[Index] <> '') and (UV_v006[Index] <> '')  then
                     UV_v014[Index] := StrToFloat(UV_v013[Index]) * StrToFloat(UV_v006[Index]);
               end;
               if pos('H017',sHul_List) > 0 then
               begin
                  UV_v015[Index] := Trim(Copy(UV_fGulkwa, 97,  6));
                  if (UV_v015[Index] <> '') and (UV_v007[Index] <> '') and (UV_v006[Index] <> '')  then
                     UV_v016[Index] := StrToFloat(UV_v015[Index]) * StrToFloat(UV_v006[Index]);
               end;
               Inc(Index);
            end;
            Next;
         End;

         // Table과 Disconnect
         Close;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD6Q04.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;


end;

procedure TfrmLD6Q04.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD6Q04.btnPrintClick(Sender: TObject);
begin
  inherited;      {
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q371 := TfrmLD4Q371.Create(Self);
     frmLD4Q371.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q371 := TfrmLD4Q371.Create(Self);
     frmLD4Q371.QuickRep.Print;
  end;     }
end;

function TfrmLD6Q04.UF_hangmokList : String;
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

function TfrmLD6Q04.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD6Q04.FormActivate(Sender: TObject);
begin
  inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   Cmb_gubn.ItemIndex := 0;
   mskDate.SetFocus;
end;

procedure TfrmLD6Q04.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;

      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
