//==============================================================================
// 프로그램 설명 : 심근경색 코드관련 혈액소겨능ㄹ 위한 자료조회 프로그램
// 시스템        : 통합검진
// 수정일자      : 20160526
// 수정자        : 박수지
// 수정내용      : 한의 재단검진사업관리팀징1600044
//==============================================================================
unit LD4Q75;

interface
                                                
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,
  QuickRpt;

type
  TfrmLD4Q75 = class(TfrmSingle)
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
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    qryProfile: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryGumgin: TQuery;
    qryHul: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_Bun, UV_vBun, UV_vC020, UV_vC022, UV_vC023, UV_vC024, UV_vC078, UV_vC080, UV_vS001, UV_vAge, UV_vSex : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q75: TfrmLD4Q75;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q801;

{$R *.DFM}

procedure TfrmLD4Q75.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q75.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q75.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Bun := VarArrayCreate([0, 0], varOleStr);
      UV_vBun     := VarArrayCreate([0, 0], varOleStr);
      UV_vC020    := VarArrayCreate([0, 0], varOleStr);
      UV_vC022    := VarArrayCreate([0, 0], varOleStr);
      UV_vC023    := VarArrayCreate([0, 0], varOleStr);
      UV_vC024    := VarArrayCreate([0, 0], varOleStr);
      UV_vC078    := VarArrayCreate([0, 0], varOleStr);
      UV_vC080    := VarArrayCreate([0, 0], varOleStr);
      UV_vS001    := VarArrayCreate([0, 0], varOleStr);
      UV_vSex     := VarArrayCreate([0, 0], varOleStr);
      UV_vAge     := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo     ,    iLength);
      VarArrayRedim(UV_vDat_Bun,  iLength);
      VarArrayRedim(UV_vBun    ,   iLength);
      VarArrayRedim(UV_vC020   ,  iLength);
      VarArrayRedim(UV_vC022   ,  iLength);
      VarArrayRedim(UV_vC023   ,  iLength);
      VarArrayRedim(UV_vC024   ,  iLength);
      VarArrayRedim(UV_vC078   ,  iLength);
      VarArrayRedim(UV_vC080   ,  iLength);
      VarArrayRedim(UV_vS001   ,  iLength);
      VarArrayRedim(UV_vSex    ,  iLength);
      VarArrayRedim(UV_vAge    ,  iLength);

   end;
end;

function TfrmLD4Q75.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q75.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q75.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDat_Bun[DataRow-1];
      3 : Value := UV_vBun[DataRow-1];
      4 : Value := UV_vC020[DataRow-1];
      5 : Value := UV_vC022[DataRow-1];
      6 : Value := UV_vC023[DataRow-1];
      7 : Value := UV_vC024[DataRow-1];
      8 : Value := UV_vC078[DataRow-1];
      9 : Value := UV_vC080[DataRow-1];
     10 : Value := UV_vS001[DataRow-1];
     11 : Value := UV_vAge[DataRow -1];
     12 : Value := UV_vSex[DataRow -1];
   end;
end;    

procedure TfrmLD4Q75.btnQueryClick(Sender: TObject);
var index, I, iRet, iTempm, iAge : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy: String;
    sJangbi_List, sHul_List, sSex : String;
    vCod_chuga : Variant;
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
      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, G.cod_jisa, G.num_samp, G.cod_chuga, G.cod_hul, G.cod_jangbi, G.cod_cancer,      ' +
                 '        G.gubn_nosin, G.gubn_nsyh, G.Gubn_adult, G.Gubn_adyh, G.Gubn_life, G.Gubn_lfyh, G.Gubn_bogun, G.Gubn_bgyh, G.Gubn_agab, g.num_id,' +
                 '        G.Gubn_agyh, G.Gubn_tkgm, B.cod_bjjs, B.Dat_Bunju, B.num_Bunju   ' +
                 ' FROM Gumgin G with(nolock) Left outer JOIN BUNJU B with(nolock)         ' +
                 '        ON G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and G.dat_gmgn = B.dat_gmgn                                                   ' ;
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ' ;
     sWhere  := sWhere + ' and ((G.cod_chuga LIKE ''%C020%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C020''))) AND  '+
                          '  (G.cod_chuga LIKE ''%C022%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C022''))) AND  '+
                          '  (G.cod_chuga LIKE ''%C023%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C023''))) AND  '+
                          '  (G.cod_chuga LIKE ''%C024%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C024''))) AND  '+
                          '  (G.cod_chuga LIKE ''%C078%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C078''))) AND  '+
                          '  (G.cod_chuga LIKE ''%C080%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''C080''))) AND  '+
                          '  (G.cod_chuga LIKE ''%S001%'' OR (G.cod_hul in (select cod_pf from profile_hm with(nolock) where cod_hm=''S001''))))' ;   



      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q38', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := index + 1;
               UV_VDat_Bun[Index] := qryBunju.FieldByName('Dat_Bunju').AsString;
               UV_VBun[Index] := qryBunju.FieldByName('num_bunju').AsString;
               GP_JuminToAge(qryBunju.FieldByname('num_jumin').AsString,qryBunju.FieldByname('Dat_gmgn').AsString, sSex, iAge);
               UV_vAge[Index] := iAge;
               case StrToInt(Copy(FieldByName('Num_jumin').AsString, 7, 1)) of
                  1,3,5,7,9 : sSex := '남';
                  2,4,6,8   : sSex := '여';
               end;
                UV_vSex[Index] := sSex;

               // 혈액결과...
                qryHul.Active := False;
                qryHul.ParamByName('num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
                qryHul.ParamByName('Dat_gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
                qryHul.Active := True;
                if qryHul.RecordCount > 0 then
                begin
                   while not qryHul.Eof do
                   begin
                      UV_fGulkwa := '';
                      UV_fGulkwa1 := qryHul.FieldByName('DESC_GLKWA1').AsString;
                      UV_fGulkwa2 := qryHul.FieldByName('DESC_GLKWA2').AsString;
                      UV_fGulkwa3 := qryHul.FieldByName('DESC_GLKWA3').AsString;
                      GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                      Case qryHul.FieldByName('gubn_part').AsInteger of
                         1 : begin
                                    if Trim(Copy(UV_fGulkwa,  103,  6)) <> '' then
                                       UV_vC022[Index] := Trim(Copy(UV_fGulkwa,  103,  6))
                                    else UV_vC022[Index] := '결과 미입력';

                                    if Trim(Copy(UV_fGulkwa,  109,  6)) <> '' then
                                       UV_vC023[Index] := Trim(Copy(UV_fGulkwa,  109,  6))
                                    else UV_vC023[Index] := '결과 미입력';

                                    if Trim(Copy(UV_fGulkwa,  115,  6)) <> '' then
                                       UV_vC024[Index] := Trim(Copy(UV_fGulkwa,  115,  6))
                                    else UV_vC024[Index] := '결과 미입력';

                                    if Trim(Copy(UV_fGulkwa,  355,  6)) <> '' then
                                       UV_vC078[Index] := Trim(Copy(UV_fGulkwa,  355,  6))
                                    else UV_vC078[Index] := '결과 미입력';

                                    if Trim(Copy(UV_fGulkwa,  367,  6)) <> '' then
                                       UV_vC080[Index] := Trim(Copy(UV_fGulkwa,  367,  6))
                                    else UV_vC080[Index] := '결과 미입력';
                             end;
                         5 : begin
                                    if Trim(Copy(UV_fGulkwa,  67,  6)) <> '' then
                                       UV_vC020[Index] := Trim(Copy(UV_fGulkwa,  67,  6))
                                    else UV_vC020[Index] := '결과 미입력';
                             end;
                         7 : begin
                                    if Trim(Copy(UV_fGulkwa,  7,  6)) <> '' then
                                       UV_vS001[Index] := Trim(Copy(UV_fGulkwa,  7,  6))
                                    else UV_vS001[Index] := '결과 미입력';
                             end;
                   end; // end of Case[Gubn_Part];
                   qryHul.Next;
                end; // end of while[qryHul];
                end; // end of if[qryHul];
            qryBunju.Next;
            Inc(Index);
            END;
         end;
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

procedure TfrmLD4Q75.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q75.UP_Click(Sender: TObject);
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


procedure TfrmLD4Q75.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   mskDate.SetFocus;
end;

procedure TfrmLD4Q75.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q801 := TfrmLD4Q801.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q801.QuickRep);
  end;
end;

function TfrmLD4Q75.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k: integer;
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

function TfrmLD4Q75.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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


procedure TfrmLD4Q75.SBut_ExcelClick(Sender: TObject);
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
