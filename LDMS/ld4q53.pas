//==============================================================================
// 프로그램 설명 : 요침사양성리스트
// 시스템        : 통합검진
// 수정일자      : 2013.10.31
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.10
// 수정자        : 곽윤설
// 수정내용      : 나이, 성별 추가
// 참고사항      : [전문의-김소영]
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.07.11
// 수정자        : 곽윤설
// 수정내용      : 샘플번호 추가
// 참고사항      : 한두례 책임 요청
//==============================================================================
unit LD4Q53;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,
  QuickRpt;

type
  TfrmLD4Q53 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    QRCompositeReport: TQRCompositeReport;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryProfile: TQuery;
    RB_bunju: TRadioButton;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    RB_samp: TRadioButton;
    sampNo1: TEdit;
    Label1: TLabel;
    SampNo2: TEdit;
    GR_sort: TRadioGroup;
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
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
     procedure SBut_ExcelClick(Sender: TObject);    
    
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vBUN, UV_vSAMP, UV_vDat_bunju, UV_vAge, UV_vGen, UV_vGulkwa :Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q53: TfrmLD4Q53;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q531;

{$R *.DFM}

procedure TfrmLD4Q53.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;



end;

procedure TfrmLD4Q53.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q53.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN       := VarArrayCreate([0, 0], varOleStr);
      UV_vSAMP      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vAge       := VarArrayCreate([0, 0], varOleStr);
      UV_vGen       := VarArrayCreate([0, 0], varOleStr);
      UV_vGulkwa    := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vSAMP,  iLength);
      VarArrayRedim(UV_vDat_bunju,iLength);
      VarArrayRedim(UV_vAge,   iLength);
      VarArrayRedim(UV_vGen,   iLength);
      VarArrayRedim(UV_vGulkwa,iLength);
   end;
end;

function TfrmLD4Q53.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check

end;

procedure TfrmLD4Q53.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   if copy(GV_sUserId, 1, 2) = '12' then
   begin
     Label7.caption := '지   사 : ';
     RB_samp.Checked;
   end;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD4Q53.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDat_bunju[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vSAMP[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vAge[DataRow-1];
      7 : Value := UV_vGen[DataRow-1];
      8 : Value := UV_vGulkwa[DataRow-1];
   end;
end;

procedure TfrmLD4Q53.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,U011,U012,U013 :Integer;
    iAge :Integer;
    sSelect, sWhere, sGroupBy, sOrderBy :String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3,  sHul_List :String;
    sMan :String;
    vCod_chuga : Variant;
    Check :Boolean;
begin
   inherited;   sSelect := ''; sWhere := ''; sOrderBy := '';
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
      sSelect := ' SELECT G.desc_name, G.num_id, G.num_jumin, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_agab,  ' +
                 '        G.Gubn_tkgm, G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.num_samp, ' +
                 '        Gubn_bogun, G.cod_cancer, G.cod_chuga, K.desc_glkwa1, K.desc_glkwa2, K.desc_glkwa3, K.dat_bunju, K.num_Bunju,      ' +
                 '        T.cod_prf, T.cod_chuga As Tcod_chuga FROM Gulkwa K with(nolock)                                                    ' +
                 ' left outer join Gumgin G with(nolock) ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn      ' +
                 ' left outer join Tkgum T with(nolock) ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn ';

     if copy(GV_sUserId,1,2) = '12' then
     sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''''
     else sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
     sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' +
                            ' AND  K.Gubn_Part  = ''03''';

     if RB_bunju.Checked then
     begin
         sWhere := sWhere + ' AND k.num_bunju >= ''' + BunjuNo1.Text + '''';
         sWhere := sWhere + ' AND k.num_bunju <= ''' + BunjuNo2.Text + '''';
     end
     else
     begin
         sWhere := sWhere + ' AND G.num_samp >= ''' + sampNo1.Text + '''';
         sWhere := sWhere + ' AND G.num_samp <= ''' + sampNo2.Text + '''';
     end;

     if GR_sort.ItemIndex = 0 then sOrderBy := ' ORDER BY G.cod_jisa, K.num_bunju '
     else sOrderBy := ' ORDER BY G.cod_jisa, G.num_samp ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q53', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            sHul_List := '';
            sHul_List := UF_hangmokList;

            if (pos('U011', sHul_List) > 0) or  (pos('U012', sHul_List) > 0) or (pos('U013', sHul_List) > 0)   then
            begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString   := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.ParamByName('gubn_part').AsString := '03';

               qryGulkwa.open;
               

               if qryGulkwa.RecordCount > 0 then
               begin
                  while not qryGulkwa.Eof do
                  begin
                         Check := False;

                          UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                          UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                          UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                          GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                         if  (((Trim(Copy(UV_fGulkwa, 67,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 67,  6))<> '') and (pos('U011', sHul_List) > 0)) or
                             (((Trim(Copy(UV_fGulkwa, 73,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 73,  6))<> '') and (pos('U012', sHul_List) > 0)) or
                             (((Trim(Copy(UV_fGulkwa, 79,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 79,  6))<> '') and (pos('U013', sHul_List) > 0))then
                             begin
                             UV_VDat_bunju[Index] := qryBunju.FieldByName('Dat_bunju').AsString;
                             UV_VBUN[Index]       := qryGulkwa.FieldByName('Num_Bunju').AsString;
                             UV_vSAMP[Index]      := qryGulkwa.FieldByName('num_samp').AsString;
                             UV_vName[Index]      := qryGulkwa.FieldByName('desc_name').AsString;
                             // 주민번호를 이용하여 성별과 나이를 구함
                             GP_JuminToAge(qryBunju.FieldByName('num_jumin').AsString,qryBunju.FieldByName('dat_gmgn').AsString, sMan, iAge);
                             UV_vAge[Index]      := iAge;
                             UV_vGen[Index]      := sMan;

                             if ((Trim(Copy(UV_fGulkwa, 67,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 67,  6))<> '')then
                                UV_vGulkwa[Index]:=+UV_vGulkwa[Index]+' WBC('+ Trim(Copy(UV_fGulkwa, 67,  6))+')';
                             if ((Trim(Copy(UV_fGulkwa, 73,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 73,  6))<> '') then
                                UV_vGulkwa[Index]:=UV_vGulkwa[Index]+' RBC('+ Trim(Copy(UV_fGulkwa, 73,  6))+')';
                             if ((Trim(Copy(UV_fGulkwa, 79,  6))) <>'0-2' ) and (Trim(Copy(UV_fGulkwa, 79,  6))<> '') then
                                UV_vGulkwa[Index]:=UV_vGulkwa[Index]+' EP('+ Trim(Copy(UV_fGulkwa, 79,  6))+')';

                             Check := True;
                        end;
                     if Check = True then
                     begin
                     UP_VarMemSet(Index+1);
                     UV_vNo[Index]  := IntToStr(Index+1);
                     Inc(Index);
                     end;
                  qryGulkwa.Next;
                  end;
                  qryGulkwa.Close;
               End;
            end;
            qryBunju.Next;
         End;
         // Table과 Disconnect
         qryBunju.Close;
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

procedure TfrmLD4Q53.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;



end;

procedure TfrmLD4Q53.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q53.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q531 := TfrmLD4Q531.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q531.QuickRep.Preview
  else                                frmLD4Q531.QuickRep.Print;

end;

procedure TfrmLD4Q53.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;

end;

procedure TfrmLD4Q53.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;

  with QRCompositeReport do
  begin

  end;
end;

function TfrmLD4Q53.UF_hangmokList : String;
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

function TfrmLD4Q53.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q53.SBut_ExcelClick(Sender: TObject);
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
