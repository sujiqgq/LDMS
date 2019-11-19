//==============================================================================
// 프로그램 설명 : 생물학적 노출지표검사 대상자조회
// 시스템        : 통합검진
// 개발일자      : 2017.08.08
// 개발자        : 유동구
// 참고사항      : 신규개발(한의 재단특검행정파트1701202)
//==============================================================================
unit LD5Q15;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD5Q15 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    cmbJisa: TComboBox;
    btnSdate: TSpeedButton;
    edtSDate: TDateEdit;
    qryGumgin: TQuery;
    Gride: TGauge;
    qryHangmokList: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Panel27: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    qryProfile: TQuery;
    Cmb_gubn: TComboBox;
    Panel3: TPanel;
    Panel6: TPanel;
    Chk_Time: TCheckBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Panel7: TPanel;
    CmbBox_Sort: TComboBox;
    CmbBD: TComboBox;
    Panel8: TPanel;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    Panel9: TPanel;
    RBtn_ALL: TRadioButton;
    RBtn_Nae: TRadioButton;
    RBtn_Chul: TRadioButton;
    Panel10: TPanel;
    Chk_Jumin: TCheckBox;
    btnEdate: TSpeedButton;
    edtEDate: TDateEdit;
    Label1: TLabel;
    Chk_saler: TCheckBox;
    Chk_Bogo: TCheckBox;
    Chk_sabun: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure SBut_ExcelClick(Sender: TObject);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn, UV_vMatter, UV_vHangmok, UV_vTemp, UV_vTime, UV_vDesc_jisa, UV_vDesc_Saler, UV_vNum_sabun : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_hangmokList(var sCod_HmList, sDesc_HmList : String);

    function UF_SqlSet(shm : string) : string;
    function UF_Condition : Boolean;
    function UF_HmCheck(sCod_Hm : String) : Boolean;    
  public
    { Public declarations }
  end;

var
  frmLD5Q15: TfrmLD5Q15;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q151;

{$R *.DFM}

function TfrmLD5Q15.UF_SqlSet(shm : string) : string;
var
   sSqlText : string;
   iRet, iCnt : integer;
   vCod_chuga : Variant;
begin
   Result := '';
   sSqlText := '';

   iRet := GF_MulToSingle(shm, 4, vCod_chuga);
   for iCnt := 0 to iRet - 1 do
   begin
      if Trim(sSqlText) = '' then sSqlText := sSqlText + '''' + vCod_chuga[iCnt] + ''''
      else                        sSqlText := sSqlText + ',''' + vCod_chuga[iCnt] + '''';
   end;
   sSqlText := ' AND H.cod_hm IN ( ' + sSqlText + ')';

   Result := sSqlText;
end;

procedure TfrmLD5Q15.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);
end;

procedure TfrmLD5Q15.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then
   begin
      if (ActiveControl is Ttsgrid) then
      begin
         if      Message.wParam > 0 then keybd_event(VK_UP,   VK_UP,   0, 0)
         else if Message.wParam < 0 then keybd_event(VK_DOWN, VK_DOWN, 0, 0);

         Application.ProcessMessages;
         grdMaster.Refresh;
      end;
   end;
end;

procedure TfrmLD5Q15.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo       := VarArrayCreate([0, 0], varOleStr);
      UV_vName     := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin    := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vMatter   := VarArrayCreate([0, 0], varOleStr);
      UV_vHangmok  := VarArrayCreate([0, 0], varOleStr);
      UV_vTemp     := VarArrayCreate([0, 0], varOleStr);
      UV_vTime     := VarArrayCreate([0, 0], varOleStr);

      UV_vDesc_jisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Saler := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Sabun  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,       iLength);
      VarArrayRedim(UV_vName,     iLength);
      VarArrayRedim(UV_vJumin,    iLength);
      VarArrayRedim(UV_vSampNo,   iLength);
      VarArrayRedim(UV_vDesc_dc,  iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
      VarArrayRedim(UV_vMatter,   iLength);
      VarArrayRedim(UV_vHangmok,  iLength);
      VarArrayRedim(UV_vTemp,     iLength);
      VarArrayRedim(UV_vTime,     iLength);

      VarArrayRedim(UV_vDesc_jisa,  iLength);
      VarArrayRedim(UV_vDesc_Saler, iLength);
      VarArrayRedim(UV_vNum_sabun,  iLength);
   end;
end;

function TfrmLD5Q15.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtSDate.Text = '') or (edtEDate.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q15.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Cmb_gubn.ItemIndex := 0;
   CmbBox_Sort.ItemIndex := 0;

   // 층구분으로 출력가능
   CmbBD.Enabled := True;
   CmbBD.ItemIndex := GP_SawonCheck(CmbBD, GV_sUserId);
end;

procedure TfrmLD5Q15.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vName[DataRow-1];
      3 : Value := UV_vJumin[DataRow-1];
      4 : Value := UV_vSampNo[DataRow-1];
      5 : Value := UV_vDesc_dc[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vMatter[DataRow-1];
      8 : Value := UV_vHangmok[DataRow-1];
      9 : Value := UV_vTemp[DataRow-1];
     10 : Value := UV_vTime[DataRow-1];
     11 : Value := UV_vDesc_jisa[DataRow-1];
     12 : Value := UV_vDesc_Saler[DataRow-1];
     13 : Value := UV_vNum_Sabun[DataRow-1];
   end;
end;

procedure TfrmLD5Q15.btnQueryClick(Sender: TObject);
var
   sSQlText, sCod_HmList, sDesc_HmList, sHangmokList, sDesc_profile : String;
   iCnt, iRet, iRow, iTemp : Integer;
   bChk_TC77 : Boolean;
   vCod_chuga : Variant;   
begin
   inherited;
   sSQlText := '';
   iRow := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);


   //센터 표기
   if copy(cmbJisa.Text,1,2) = '00' then grdMaster.Col[11].Visible := True
   else                                  grdMaster.Col[11].Visible := False;

   //영업사원 표기
   if Chk_saler.Checked then grdMaster.Col[12].Visible := True
   else                      grdMaster.Col[12].Visible := False;
   //사번표기
   if Chk_sabun.Checked then grdMaster.Col[13].Visible := True
   else                      grdMaster.Col[13].Visible := False;



   sSQlText := ' SELECT dbo.F_GET_JISA_FIND(G.cod_jisa) Desc_jisa, dbo.F_GET_SAWON_FIND(G.cod_saler) Desc_saler, ' + #13 +
               '        D.Desc_dc, G.cod_jisa, G.num_id,  T.cod_prf, ' + #13 +
               '        G.desc_name,  G.num_jumin,  G.dat_gmgn, G.Gubn_tkgm,  G.cod_dc, G.num_samp, G.Today_insert, G.num_sabun   ' + #13 +
               ' FROM Gumgin G with(nolock) INNER JOIN Tkgum  T WITH (NOLOCK) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ' + #13 +
               '                            INNER JOIN Danche D WITH (NOLOCK) ON G.cod_dc = D.cod_dc ' + #13 +
               ' WHERE 0 = 0 ';
   if copy(cmbJisa.Text,1,2) <> '00' then sSQlText := sSQlText + ' AND G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ' + #13;

   sSQlText := sSQlText + '   AND G.Dat_gmgn >= ''' + edtSDate.Text + ''' ' + #13;
   sSQlText := sSQlText + '   AND G.Dat_gmgn <= ''' + edtEDate.Text + ''' ' + #13;
   sSQlText := sSQlText + '   AND G.gubn_tkgm in (''1'',''2'') ' + #13;

   if Trim(edtDc.Text) <> '' then sSQlText := sSQlText + '  AND G.cod_dc = ''' + edtDc.Text + ''' ' + #13;

   if cmbBD.itemindex > 0 then sSQlText := sSQlText + '  AND G.gubn_jinch = ''' + trim(copy(cmbBD.Text,1,2)) + ''' ';

   //[2017.10.14-유동구]한의 재단특검행정파트1701515의거 내원/출장구분 추가
   if RBtn_Nae.Checked  then sSQlText := sSQlText + '  AND G.gubn_Neawon <> ''2'' ';
   if RBtn_Chul.Checked then sSQlText := sSQlText + '  AND G.gubn_Neawon =  ''2'' ';

   if Chk_Bogo.Checked then
   begin
      sSQlText := sSQlText + '  AND T.gubn_bogo = ''1'' ';
      sSQlText := sSQlText + '  AND (D.Gubn_Nodong <> ''Y'' Or D.Gubn_Nodong IS NULL) ';
   end;

   case CmbBox_Sort.ItemIndex of
      0 : sSQlText := sSQlText + ' ORDER BY G.Desc_name, G.num_jumin ';
      1 : sSQlText := sSQlText + ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
   end;

   with qryGumgin do
   begin
      // SQL문 생성
      Close;
      SQL.Clear;
      SQL.Text := sSQlText;
      Open;

      GP_query_log(qryGumgin, 'LD5Q15', 'Q', 'N');

      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         for iCnt := 1 to RecordCount do
         begin
            Gride.Progress := iCnt;
            application.ProcessMessages;

            sCod_HmList  := '';
            sDesc_HmList := '';
            UP_hangmokList(sCod_HmList, sDesc_HmList);

            sHangmokList  := '';
            sDesc_profile := '';
            bChk_TC77 := False;

            if (Trim(sCod_HmList) <> '') then
            begin
               sHangmokList := UF_SqlSet(sCod_HmList);

               if pos('TC77',FieldByName('cod_prf').AsString) > 0 then bChk_TC77 := True
               else
               begin
                  iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_chuga);
                  for iTemp := 0 to iRet - 1 do
                  begin
                     qryProfile.Active := False;
                     qryProfile.SQL.Clear;
                     qryProfile.SQL.Text := ' SELECT H.cod_pf, G.desc_pf FROM profile_hm H WITH (NOLOCK) INNER JOIN profile G WITH (NOLOCK) ON H.cod_pf = G.cod_pf ' +
                                            ' WHERE H.cod_pf = ''' + vCod_chuga[iTemp] + ''' ' + sHangmokList;
                     qryProfile.Active := True;

                     if qryProfile.RecordCount > 0 then
                     begin
                        if Trim(sDesc_profile) = '' then sDesc_profile := sDesc_profile + qryProfile.FieldByName('desc_pf').AsString
                        else                             sDesc_profile := sDesc_profile + ',' + #13 + qryProfile.FieldByName('desc_pf').AsString;
                     end;
                     qryProfile.Active := False;
                  end;
               end;

               if not bChk_TC77 then
               begin
                  UP_VarMemSet(iRow);

                  //No
                  UV_vNo[iRow]       := IntToStr(iRow+1);
                  //성명
                  UV_vName[iRow]     := qryGumgin.FieldByName('desc_name').AsString;

                  //생년월일
                  if Chk_Jumin.Checked then UV_vJumin[iRow] := Copy(qryGumgin.FieldByName('num_jumin').AsString,1,6) +'-'+ Copy(qryGumgin.FieldByName('num_jumin').AsString,7,7)
                  else                      UV_vJumin[iRow] := (Copy(qryGumgin.FieldByName('num_jumin').AsString,1,6)) +'-'+ (Copy(qryGumgin.FieldByName('num_jumin').AsString,7,1)) + '******';

                  //샘플번호
                  UV_vSampNo[iRow]   := qryGumgin.FieldByName('num_samp').AsString;
                  //단체명
                  UV_vDesc_dc[iRow]  := qryGumgin.FieldByName('desc_dc').AsString;
                  //검진일자
                  UV_vDat_gmgn[iRow] := GF_DateFormat(qryGumgin.FieldByName('Dat_gmgn').AsString);
                  //유해물질
                  UV_vMatter[iRow]   := sDesc_profile;
                  //노출지표 검사항목
                  UV_vHangmok[iRow]  := sDesc_HmList;
                  //채취시간
                  UV_vTemp[iRow]     := '문서첨부';
                  //수거시간
                  if Chk_Time.Checked then UV_vTime[iRow] := '문서첨부'
                  else                     UV_vTime[iRow] := copy(qryGumgin.FieldByName('Today_insert').AsString,1,4) + '-' +
                                                             copy(qryGumgin.FieldByName('Today_insert').AsString,5,2) + '-' +
                                                             copy(qryGumgin.FieldByName('Today_insert').AsString,7,20);

                  //센터 표기
                  if copy(cmbJisa.Text,1,2) = '00' then UV_vDesc_jisa[iRow] := qryGumgin.FieldByName('Desc_jisa').AsString
                  else                                  UV_vDesc_jisa[iRow] := '';
               
                  //영업사원 표기
                  if Chk_saler.Checked then UV_vDesc_Saler[iRow] := qryGumgin.FieldByName('Desc_saler').AsString
                  else                      UV_vDesc_Saler[iRow] := '';

                  //사번 표기
                  if Chk_sabun.Checked then UV_vNum_Sabun[iRow]  := qryGumgin.FieldByName('Num_sabun').AsString
                  else                      UV_vNum_Sabun[iRow] := '';

                  Inc(iRow);
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
      end;
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := iRow;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD5Q15.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;
end;

procedure TfrmLD5Q15.UP_Click(Sender: TObject);
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
  else if Sender = btnSdate then GF_CalendarClick(edtSDate)
  else if Sender = btnEdate then GF_CalendarClick(edtEDate);
end;

procedure TfrmLD5Q15.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD5Q151 := TfrmLD5Q151.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q151.QuickRep.Preview
  else                                frmLD5Q151.QuickRep.Print;
end;

procedure TfrmLD5Q15.UP_hangmokList(var sCod_HmList, sDesc_HmList : String);
begin
   sCod_HmList  := '';
   sDesc_HmList := '';
   with qryHangmokList do
   begin
      Active := False;
      Sql.Clear;
      Sql.Text := ' SELECT DISTINCT G.cod_hm, C12.HM_BIOEXPFLAG, C12.검사항목명_한, H.Desc_Kor ' +
                  ' FROM TF_Get_HangmokList(''' + qryGumgin.FieldByName('cod_jisa').AsString + ''', '''
                                                + qryGumgin.FieldByName('Num_id').AsString   + ''', '''
                                                + qryGumgin.FieldByName('Dat_gmgn').AsString + ''', ''3'') G ' +
                  ' LEFT OUTER JOIN HEMS..Comm12 C12 WITH (NOLOCK) ON G.Cod_hm = C12.MDMS항목코드 ' +
                  ' LEFT OUTER JOIN hangmok      H   WITH (NOLOCK) ON G.Cod_hm = H.cod_hm ' +
                  ' WHERE C12.HM_BIOEXPFLAG = ''1'' ';
      Active := True;

      if RecordCount > 0 then
      begin
         while Not Eof do
         begin
            if UF_HmCheck(FieldByName('cod_hm').AsString) then
            begin
               if sCod_HmList = '' then
               begin
                  sCod_HmList  := sCod_HmList  + FieldByName('cod_hm').AsString;
                  sDesc_HmList := sDesc_HmList + FieldByName('Desc_Kor').AsString;
               end
               else
               begin
                  sCod_HmList  := sCod_HmList  + FieldByName('cod_hm').AsString;
                  sDesc_HmList := sDesc_HmList + ',' + #13 + FieldByName('Desc_Kor').AsString;
               end;
            end;

            Next;
         end;
      end;
   end;
end;

function TfrmLD5Q15.UF_HmCheck(sCod_Hm : String) : Boolean;
begin
   Result := False;

   case Cmb_gubn.ItemIndex of
      0 : Result := True;
      1 : if (sCod_Hm = 'SEA2') then Result := True;
      2 : if (sCod_Hm = 'SE87') then Result := True;
      3 : if (sCod_Hm = 'SE42') then Result := True;
      4 : if (sCod_Hm = 'P012') then Result := True;
      5 : if (sCod_Hm = 'SE85') or
             (sCod_Hm = 'SE89') or
             (sCod_Hm = 'SE90') or
             (sCod_Hm = 'SE91') or
             (sCod_Hm = 'SE92') or
             (sCod_Hm = 'SE93') or
             (sCod_Hm = 'SE96') or
             (sCod_Hm = 'SE97') or
             (sCod_Hm = 'SE98') or
             (sCod_Hm = 'SEA1') or
             (sCod_Hm = 'SE61') then Result := True;
      6 : if (sCod_Hm = 'JJFG') then Result := True;
      7 : if (sCod_Hm = 'SEA6') then Result := True;
      8 : if (sCod_Hm = 'H029') then Result := True;
      9 : if (sCod_Hm = 'SE85') then Result := True;
     10 : if (sCod_Hm = 'SE61') then Result := True;
   end;
end;

procedure TfrmLD5Q15.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col, Col_Cnt : Integer;
  sTemp, sTemp2 : String;
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
            Col_Cnt := 0;
            for Col := 0 to grdMaster.Cols - 1 do
            begin
               //[2019.01.29-유동구]체크박스 되어 있는 내용만 엑셀변환
               if grdMaster.Col[Col + 1].Visible then
               begin
                  ArrV3[Row, Col_Cnt] := grdMaster.Col[Col + 1].Heading;
                  inc(Col_Cnt);
               end;
            end;
            //for Col := 0 to grdMaster.Cols - 1 do
            //   ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            Col_Cnt := 0;
            for Col := 0 to grdMaster.Cols - 1 do
            begin
               Application.ProcessMessages;

               //[2018.09.15-유동구]체크박스 되어 있는 내용만 엑셀변환(한의 재단특검행정파트1800947 의거)
               if grdMaster.Col[Col+1].Visible then
               begin
                  Application.ProcessMessages;
                  ArrV3[Row, Col_Cnt] := grdMaster.cell[Col + 1, Row];
                  inc(Col_Cnt);
               end;
            end;

            //for Col := 0 to grdMaster.Cols do
            //begin
            //   Application.ProcessMessages;
            //   ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            //end;
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

procedure TfrmLD5Q15.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   if      DataCol =  9 then UV_vTemp[DataRow-1] := grdMaster.CurrentCell.Value  //채취시간
   else if DataCol = 10 then UV_vTime[DataRow-1] := grdMaster.CurrentCell.Value; //수거시간
end;

end.
