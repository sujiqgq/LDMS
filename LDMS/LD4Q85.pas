//==============================================================================
// 프로그램 설명 : 검사항목분주현황(생화학전용)
// 시스템        : 통합검진
// 수정일자      : 2010.04.13
// 수정자        : 송재호
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.11.03
// 수정자        : 곽윤설
// 수정내용      : 출력 양식 변경
// 참고사항      : [본원-최은희]
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.05.29
// 수정자        : 곽윤설
// 수정내용      : TPLA(S036) 항목추가.
// 참고사항      : 한의 재단진단검사의학팀1500435
//==============================================================================
// 수정일자      : 2015.06.17
// 수정자        : 곽윤설
// 수정내용      : C022 아포지단백A, C023 아포지단백 B, C080 LP(a) 제외 요청
// 참고사항      : 한의 재단진단검사의학팀1500238
//==============================================================================
// 수정일자      : 2015.07.10
// 수정자        : 곽윤설
// 수정내용      : C022 아포지단백A, C023 아포지단백 B, C080 LP(a) 조회 가능하도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.07.15
// 수정자        : 곽윤설
// 수정내용      : C074 본원 검사만 조회되도록 (cod_jisa='15' or gubn_hulgum='3')
// 참고사항      : 본원-엄하니
//==============================================================================
// 수정일자      : 2016.03.29
// 수정자        : 박수지
// 수정내용      : CO78항목코드 추가
// 참고사항      : 한의 재단진단검사의학팀 1600348
//==============================================================================
// 수정일자      : 2018.04.25
// 수정자        : 박수지
// 수정내용      : CO78항목 본원 수검자만 (cod_jisa='15')
// 참고사항      : 한의 재단일반파트 1800324
//==============================================================================
unit LD4Q85;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD4Q85 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    RB_02: TRadioButton;
    RB_01: TRadioButton;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Label4: TLabel;
    RadioGroup1: TRadioGroup;
    qryProfile: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
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
    procedure QRCompositeReportAddReports(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vBUN, UV_vName, UV_v001, UV_v002, UV_v003, UV_v004  : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q85: TfrmLD4Q85;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q851;

{$R *.DFM}

procedure TfrmLD4Q85.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q85.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q85.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_v001   := VarArrayCreate([0, 0], varOleStr);
      UV_v002   := VarArrayCreate([0, 0], varOleStr);
      UV_v003   := VarArrayCreate([0, 0], varOleStr);
      UV_v004   := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,   iLength);
      VarArrayRedim(UV_vName, iLength);
      VarArrayRedim(UV_vBUN,  iLength);
      VarArrayRedim(UV_v001,  iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
   end;
end;

function TfrmLD4Q85.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q85.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // 오늘일자로 기본설정
   mskDate.Text := GV_sToday;
end;

procedure TfrmLD4Q85.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;
   begin
   // 자룔를 화면에 조회
   case DataCol of
      1  : value := UV_vNO[DataRow-1];
      2  : Value := UV_vBUN[DataRow-1];  //분주번호 or 샘플번호
      3  : Value := UV_vName[DataRow-1]; //성명
      4  : Value := UV_v001[DataRow-1];
      5  : Value := UV_v002[DataRow-1];
      6  : Value := UV_v003[DataRow-1];
      7  : Value := UV_v004[DataRow-1];
   end;
    grdMaster.Col[1].Heading   := 'NO.';
    grdMaster.Col[2].Heading   := '샘플번호';
    grdMaster.Col[3].Heading   := '성명';
    grdMaster.Col[4].Heading   := 'CA 15-3(여자)';
    grdMaster.Col[5].Heading   := 'Ferritin';
    grdMaster.Col[6].Heading   := 'T4';
    grdMaster.Col[7].Heading   := 'CYFRA';
   end;
end;

procedure TfrmLD4Q85.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    Sum_v001, Sum_v002, Sum_v003, Sum_v004 : Integer;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Sum_v001 := 0; Sum_v002 := 0; Sum_v003 := 0; Sum_v004 := 0;  Index := 0;
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  D.Desc_dc, B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp, G.desc_name, G.num_id, G.dat_gmgn, B.dat_bunju, G.cod_jisa, G.gubn_nosin, G.gubn_adult,  ' +
                 '         G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul,          ' +
                 '         G.cod_cancer, G.cod_chuga, G.gubn_hulgum, T.cod_prf, T.cod_chuga As Tcod_chuga FROM bunju B with(nolock)                                                ' +
                 '         Inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn                                         ' +
                 '         Left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ' +
                 '         Left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc                                                                                            ' ;

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ';
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + mskDate.Text + ''' and (Gubn_Part  = ''01''))';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if RB_01.Checked then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY B.Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q85', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

           if (pos('T037',sHul_List) > 0) or (pos('C049',sHul_List) > 0) or
              (pos('E002',sHul_List) > 0) or (pos('E040',sHul_List) > 0) then
           begin
              UP_VarMemSet(Index+1);

              UV_vNo[Index]  := Index+1;

              if RB_01.Checked then
                 UV_vBUN[Index] := qryBunju.FieldByName('num_samp').AsString
              else
                 UV_vBUN[Index] := qryBunju.FieldByName('num_bunju').AsString;
              UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;


              if pos('T037', sHul_List) > 0 then
              begin
                 UV_v001[Index] := '*';
                 Sum_v001 := Sum_v001 + 1;
              end
              else                               UV_v001[Index] := '';          //Sum_v001

              if pos('C049', sHul_List) > 0 then
              begin
                 UV_v002[Index] := '*';
                 Sum_v002 := Sum_v002 + 1;
              end
              else                               UV_v002[Index] := '';          //Sum_v002

              if pos('E002', sHul_List) > 0 then
              begin
                 UV_v003[Index] := '*';
                 Sum_v003 := Sum_v003 + 1;
              end
              else                               UV_v003[Index] := '';          //Sum_v003

              if pos('E040', sHul_List) > 0 then
              begin
                 UV_v004[Index] := '*';
                 Sum_v004 := Sum_v004 + 1;
              end
              else                               UV_v004[Index] := '';          //Sum_v004

              Inc(Index);
           end;
      Next;
      end;
         // Table과 Disconnect
         Close;

         UV_vNo[Index]   := '총 합 계';
         UV_vBUN[Index]  := '';
         UV_vName[Index] := '';
         UV_v001[Index]  := Sum_v001;
         UV_v002[Index]  := Sum_v002;
         UV_v003[Index]  := Sum_v003;
         UV_v004[Index]  := Sum_v004;
         Inc(Index);
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

procedure TfrmLD4Q85.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

   if RB_01.Checked then
   begin
       grdMaster.Col[2].Heading   := '샘플번호';
       MEdt_SampS.Enabled := True;
       MEdt_SampE.Enabled := True;
       mskFrom.Enabled := False;
       mskTo.Enabled := False;
   end
   else
   begin
       grdMaster.Col[2].Heading   := '분주번호';
       MEdt_SampS.Enabled := False;
       MEdt_SampE.Enabled := False;
       mskFrom.Enabled := True;
       mskTo.Enabled := True;
   end;

end;

procedure TfrmLD4Q85.UP_Click(Sender: TObject);
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
   else
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD4Q85.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True   then QRCompositeReport.Preview
  else                                  QRCompositeReport.Print;


end;

procedure TfrmLD4Q85.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   mskDate.SetFocus;
end;

procedure TfrmLD4Q85.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q851 := TfrmLD4Q851.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q851.QuickRep);
  end;
end;

function TfrmLD4Q85.UF_hangmokList : String;
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

function TfrmLD4Q85.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q85.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.

