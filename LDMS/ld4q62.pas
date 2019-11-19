//==============================================================================
// 프로그램 설명 : U017 && Y004 검진자 LIST
// 시스템        : 통합검진
// 개발일자      : 2014.10.20
// 개발자        : 곽윤설
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.01.23
// 수정자        : 곽윤설
// 수정내용      : 분주기간 조회 -> 분주일자 조회로 수정
// 참고사항      : 요청자 [전문의-김소영]
//==============================================================================
unit LD4Q62;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt,ComObj;

type
  TfrmLD4Q62 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    btnBdate_From: TSpeedButton;
    Label5: TLabel;
    edtBDate_From: TDateEdit;
    qryGumgin: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    RadioGroup1: TRadioGroup;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    edtBDate_To: TDateEdit;
    btnBdate_To: TSpeedButton;
    Label1: TLabel;
    RB_Bunju: TRadioButton;
    RB_Samp: TRadioButton;
    Label3: TLabel;
    BunjuNo1: TMaskEdit;
    BunjuNo2: TMaskEdit;
    SampNo1: TMaskEdit;
    SampNo2: TMaskEdit;
    Label4: TLabel;
    CB_HM: TComboBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Label9: TLabel;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vJisa, UV_vDat_gmgn, UV_vDat_bunju, UV_vBUNJU, UV_vSampNo,
    UV_vName, UV_vMan, UV_vAge, UV_vU017, UV_vY004 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q62: TfrmLD4Q62;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc ;

{$R *.DFM}

procedure TfrmLD4Q62.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q62.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then
   begin
      if (ActiveControl is Ttsgrid) then
      begin
         if Message.wParam > 0 then
            keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
            keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;
         grdMaster.Refresh;
      end;
   end;
end;

procedure TfrmLD4Q62.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vBUNJU     := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vMan       := VarArrayCreate([0, 0], varOleStr);
      UV_vAge       := VarArrayCreate([0, 0], varOleStr);
      UV_vU017      := VarArrayCreate([0, 0], varOleStr);
      UV_vY004      := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo       , iLength);
      VarArrayRedim(UV_vJisa     , iLength);
      VarArrayRedim(UV_vDat_gmgn , iLength);
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vBUNJU    , iLength);
      VarArrayRedim(UV_vSampNo   , iLength);
      VarArrayRedim(UV_vName     , iLength);
      VarArrayRedim(UV_vMan      , iLength);
      VarArrayRedim(UV_vAge      , iLength);
      VarArrayRedim(UV_vU017     , iLength);
      VarArrayRedim(UV_vY004     , iLength);
   end;
end;

function TfrmLD4Q62.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate_From.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q62.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   //기본항목설정
   CB_HM.ItemIndex := 0;
end;

procedure TfrmLD4Q62.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo       [DataRow-1];
      2 : Value := UV_vJisa     [DataRow-1];
      3 : Value := UV_vDat_gmgn [DataRow-1];
      4 : Value := UV_vDat_bunju[DataRow-1];
      5 : Value := UV_vBUNJU    [DataRow-1];
      6 : Value := UV_vSampNo   [DataRow-1];
      7 : Value := UV_vName     [DataRow-1];
      8 : Value := UV_vMan      [DataRow-1];
      9 : Value := UV_vAge      [DataRow-1];
     10 : Value := UV_vU017     [DataRow-1];
     11 : Value := UV_vY004     [DataRow-1];
   end;
end;

procedure TfrmLD4Q62.btnQueryClick(Sender: TObject);
var index, I, iAge : Integer;
    sSelect, sWhere, sOrderBy, sHangmokList, sMan : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryGumgin do
   begin
      // SQL문 생성
      Close;

      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, G.cod_jisa, G.num_samp, G.cod_chuga, G.cod_hul, G.cod_jangbi, G.cod_cancer, ' +
                 '        G.gubn_nosin, G.gubn_nsyh, G.Gubn_adult, G.Gubn_adyh, G.Gubn_life, G.Gubn_lfyh, G.Gubn_bogun, G.Gubn_bgyh, G.Gubn_agab,     ' +
                 '        G.Gubn_agyh, G.Gubn_tkgm, B.cod_bjjs, B.Dat_Bunju, B.num_Bunju, B.DESC_GLKWA1, B.DESC_GLKWA2, B.DESC_GLKWA3                 ' +
                 ' FROM Gumgin G with(nolock) left outer join gulkwa B with(nolock)                                                                   ' +
                 '        ON G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and G.dat_gmgn = B.dat_gmgn                                              ' ;

      sWhere := ' WHERE B.Dat_Bunju  = ''' + edtBDate_From.Text + ''' ';
//                 + '   AND B.Dat_Bunju  <= ''' + edtBDate_To.Text + ''' ';

      if RB_Bunju.Checked then
      begin
         If Trim(BunjuNo1.Text) <> '' Then
            sWhere := sWhere + ' And B.Num_Bunju >= ''' + BunjuNo1.Text + ''' ' +
                               ' And B.Num_Bunju <= ''' + BunjuNo2.Text + ''' ';
      end
      else
      begin
         If Trim(SampNo1.Text) <> '' Then
            sWhere := sWhere + ' And G.Num_Samp >= ''' + SampNo1.Text + ''' ' +
                               ' And G.Num_Samp <= ''' + SampNo2.Text + ''' ';
      end;

      if CB_HM.ItemIndex = 0 then
         sWhere := sWhere + ' And B.gubn_part = ''03'' ';

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY G.dat_gmgn, CAST(G.num_samp AS INT) ';
         1 : sOrderBy := ' ORDER BY G.dat_gmgn, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Gride.Progress := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD4Q62', 'Q', 'N');
         for I := 1 to RecordCount do
         begin
            Gride.Progress := I;

            sHangmokList := '';
            sHangmokList := UF_hangmokList;

//            if (qryGumgin.FieldByName('gubn_part').AsString = '03') then
//            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryGumgin.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryGumgin.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryGumgin.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
//            end;

            if (CB_HM.ItemIndex = 0)  then
            begin
               if ((pos('U017',sHangmokList) > 0)) or ((pos('Y004',sHangmokList) > 0)) AND ((Length(qryGumgin.FieldByName('num_samp').AsString) = 6))  then
               begin
                  UP_VarMemSet(Index+1);
                  GP_JuminToAge(qryGumgin.FieldByName('num_jumin').AsString, qryGumgin.FieldByName('Dat_gmgn').AsString, sMan, iAge);

                  UV_vNo        [Index] := IntToStr(Index+1);
                  UV_vJisa      [Index] := qryGumgin.FieldByName('cod_jisa').AsString;
                  UV_vDat_gmgn  [Index] := qryGumgin.FieldByName('Dat_gmgn').AsString;
                  UV_vDat_bunju [Index] := qryGumgin.FieldByName('Dat_bunju').AsString;
                  UV_vBUNJU     [Index] := qryGumgin.FieldByName('num_bunju').AsString;
                  UV_vSampNo    [Index] := qryGumgin.FieldByName('num_samp').AsString;
                  UV_vName      [Index] := qryGumgin.FieldByName('desc_name').AsString;
                  UV_vMan       [Index] := sMan;
                  UV_vAge       [Index] := iAge;
                  UV_vU017      [Index] := Trim(Copy(UV_fGulkwa,  115,  6));
                  UV_vY004      [Index] := Trim(Copy(UV_fGulkwa,  85,  6));
                  Inc(Index);
               end;
               Next;
            end;
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

procedure TfrmLD4Q62.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate_From then GF_CalendarClick(edtBDate_From);
//   else if Sender = btnBdate_To then GF_CalendarClick(edtBDate_To);

   if RB_Bunju.Checked then
   begin
      BunjuNo1.Enabled := True;
      BunjuNo2.Enabled := True;
      SampNo1.Enabled := False;
      SampNo2.Enabled := False;
   end
   else if RB_Samp.Checked then
   begin
      BunjuNo1.Enabled := False;
      BunjuNo2.Enabled := False;
      SampNo1.Enabled := True;
      SampNo2.Enabled := True;
   end
end;


function TfrmLD4Q62.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryGumgin.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_hul').AsString;
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
      if qryGumgin.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_cancer').AsString;
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
      if qryGumgin.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryGumgin.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryGumgin.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '1', qryGumgin.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGumgin.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '4', qryGumgin.FieldByName('Gubn_adyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '7', qryGumgin.FieldByName('Gubn_lfyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '3',qryGumgin.FieldByName('Gubn_bgyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '5',qryGumgin.FieldByName('Gubn_agyh').AsInteger);

   if (qryGumgin.FieldByName('Gubn_tkgm').AsString = '1') or (qryGumgin.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGumgin.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('Dat_gmgn').AsString;
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

function TfrmLD4Q62.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q62.FormActivate(Sender: TObject);
begin
   inherited;
   edtBdate_From.SetFocus;
end;

procedure TfrmLD4Q62.SBut_ExcelClick(Sender: TObject);
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
