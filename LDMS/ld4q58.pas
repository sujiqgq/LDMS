//==============================================================================
// 프로그램 설명 : 뇨화학 외주항목조회
// 시스템        : 통합검진
// 수정일자      : 2014.04.29
// 수정자        : 곽윤설
// 수정내용      : 새로개발
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.03.13
// 수정자        : 곽윤설
// 수정내용      : U026추가
// 참고사항      : 본원-한두례
//==============================================================================
// 수정일자      : 2015.04.14
// 수정자        : 곽윤설
// 수정내용      : U043삭제, U062추가
// 참고사항      : 연세대 공동협약 연구사업 [본원-박연숙]
//==============================================================================
unit LD4Q58;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,
  QuickRpt;

type
  TfrmLD4Q58 = class(TfrmSingle)
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
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    QRCompositeReport: TQRCompositeReport;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    qryProfile: TQuery;
    RB_samp: TRadioButton;
    sampNo1: TEdit;
    Label1: TLabel;
    SampNo2: TEdit;
    RB_bunju: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo,   UV_vSAMP, UV_vBUN,  UV_vName, UV_vJumin,
    UV_vU020, UV_vU023, UV_vU026, UV_vU028, UV_vU037, UV_vU062, UV_vU045, UV_vU047 :Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q58: TfrmLD4Q58;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q581;

{$R *.DFM}

procedure TfrmLD4Q58.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;



end;

procedure TfrmLD4Q58.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q58.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vSAMP      := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN       := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin     := VarArrayCreate([0, 0], varOleStr);
      UV_vU020      := VarArrayCreate([0, 0], varOleStr);
      UV_vU023      := VarArrayCreate([0, 0], varOleStr);
      UV_vU026      := VarArrayCreate([0, 0], varOleStr);
      UV_vU028      := VarArrayCreate([0, 0], varOleStr);
      UV_vU037      := VarArrayCreate([0, 0], varOleStr);
      UV_vU062      := VarArrayCreate([0, 0], varOleStr);
      UV_vU045      := VarArrayCreate([0, 0], varOleStr);
      UV_vU047      := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vSAMP,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vU020,  iLength);
      VarArrayRedim(UV_vU023,  iLength);
      VarArrayRedim(UV_vU026,  iLength);
      VarArrayRedim(UV_vU028,  iLength);
      VarArrayRedim(UV_vU037,  iLength);
      VarArrayRedim(UV_vU062,  iLength);
      VarArrayRedim(UV_vU045,  iLength);
      VarArrayRedim(UV_vU047,  iLength);
   end;
end;

function TfrmLD4Q58.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check

end;

procedure TfrmLD4Q58.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   if copy(GV_sUserId,1,2) ='12' then
   begin
    Label7.caption := '지    사  : ';
    Label2.caption := '검진일';
   end;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD4Q58.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1  : Value := UV_vNo[DataRow-1];
      2  : Value := UV_vSAMP[DataRow-1];
      3  : Value := UV_vBUN[DataRow-1];
      4  : Value := UV_vName[DataRow-1];
      5  : Value := UV_vJumin[DataRow-1];
      6  : Value := UV_vU020[DataRow-1];
      7  : Value := UV_vU023[DataRow-1];
      8  : Value := UV_vU026[DataRow-1];
      9  : Value := UV_vU028[DataRow-1];
      10 : Value := UV_vU037[DataRow-1];
      11 : Value := UV_vU062[DataRow-1];
      12 : Value := UV_vU045[DataRow-1];
      13 : Value := UV_vU047[DataRow-1];
   end;
end;

procedure TfrmLD4Q58.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    temp_U020, temp_U023, temp_U026, temp_U028, temp_U037, temp_U062, temp_U045, temp_U047 : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3,  sHul_List :String;
    vCod_chuga : Variant;
    Check :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0; temp_U020 := 0; temp_U023 := 0; temp_U026 := 0; temp_U028 := 0;
   temp_U037 := 0; temp_U062 := 0; temp_U045 := 0; temp_U047 := 0;

   // 조회조건 Check
   if not UF_Condition then exit;
   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT G.desc_name, G.num_id, G.num_jumin, G.num_samp, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life,   ' +
                 '        G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, B.dat_bunju,        ' +
                 '        B.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp  ' +
                 ' FROM Bunju B with(nolock)                                                                                                 ' +
                 ' left outer join Gumgin G with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn      ' +
                 ' left outer join Tkgum T with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn ';

     //여의도 지사에서 검진일로 변경, 분주지사->지사로 요청 ..한미영
      if copy(cmbJisa.Text,1,2) ='12' then sWhere  := ' WHERE B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''''
      else sWhere  :=  ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      if copy(cmbJisa.Text,1,2) ='12' then sWhere  := sWhere + ' AND B.Dat_Gmgn  = ''' + edtBDate.Text + ''' '
      else sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';

      if RB_bunju.Checked then
      begin
         sWhere := sWhere + ' AND B.num_bunju >= ''' + BunjuNo1.Text + '''';
         sWhere := sWhere + ' AND B.num_bunju <= ''' + BunjuNo2.Text + '''';
      end
      else
      begin
         sWhere := sWhere + ' AND G.num_samp >= ''' + sampNo1.Text + '''';
         sWhere := sWhere + ' AND G.num_samp <= ''' + sampNo2.Text + '''';
      end;

      sOrderBy := ' ORDER BY G.cod_jisa, B.num_bunju ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q58', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            sHul_List := '';
            sHul_List := UF_hangmokList;

            if (pos('U020', sHul_List) > 0) or (pos('U023', sHul_List) > 0) or (pos('U028', sHul_List) > 0) or
               (pos('U037', sHul_List) > 0) or (pos('U062', sHul_List) > 0) or (pos('U045', sHul_List) > 0) or
               (pos('U047', sHul_List) > 0) or (pos('U026', sHul_List) > 0) then
            begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;

               if qryGulkwa.RecordCount > 0 then
               begin
                  Check := False;
                  UV_vSAMP[Index]  := qryBunju.FieldByName('num_samp').AsString;
                  UV_VBUN[Index]   := qryBunju.FieldByName('Num_Bunju').AsString;
                  UV_vName[Index]  := qryBunju.FieldByName('desc_name').AsString;
                  UV_vJumin[Index] := Copy(qryBunju.FieldByName('Num_jumin').AsString, 1, 6);

                  if   (pos ('U020', sHul_List) > 0) then
                  begin
                     UV_vU020[Index]  := '*';
                     temp_U020 := temp_U020 + 1;
                     Check := True;
                  end;
                  if   (pos ('U023', sHul_List) > 0) then
                  begin
                     UV_vU023[Index]  := '*';
                     temp_U023 := temp_U023 + 1;
                     Check := True;
                  end;
                  if   (pos ('U026', sHul_List) > 0) then
                  begin
                     UV_vU026[Index]  := '*';
                     temp_U026 := temp_U026 + 1;
                     Check := True;
                  end;
                  if   (pos ('U028', sHul_List) > 0) then
                  begin
                     UV_vU028[Index]  := '*';
                     temp_U028 := temp_U028 + 1;
                     Check := True;
                  end;
                  if  (pos ('U037', sHul_List) > 0) then
                  begin
                     UV_vU037[Index]  := '*';
                     temp_U037 := temp_U037 + 1;
                     Check := True;
                  end;
                  if  (pos ('U062', sHul_List) > 0) then
                  begin
                     UV_vU062[Index]  := '*';
                     temp_U062 := temp_U062 + 1;
                     Check := True;
                  end;
                  if  (pos ('U045', sHul_List) > 0) then
                  begin
                     UV_vU045[Index]  := '*';
                     temp_U045 := temp_U045 + 1;
                     Check := True;
                  end;
                  if  (pos ('U047', sHul_List) > 0) then
                  begin
                     UV_vU047[Index]  := '*';
                     temp_U047 := temp_U047 + 1;
                     Check := True;
                  end;
                  if Check = True then
                  begin
                     UP_VarMemSet(Index+1);
                     UV_vNo[Index]  := IntToStr(Index+1);
                     Inc(Index);
                  end;
                  qryGulkwa.Close;
               End;
            end;
            qryBunju.Next;
         End;
         // Table과 Disconnect
         qryBunju.Close;

         UV_vNo[Index]  := '';
         UV_vSAMP[Index]  := '';
         UV_VBUN[Index]   := '';
         UV_vName[Index]  := '';
         UV_vJumin[Index] := '총합계';
         UV_vU020[Index]  := temp_U020;
         UV_vU023[Index]  := temp_U023;
         UV_vU026[Index]  := temp_U026;
         UV_vU028[Index]  := temp_U028;
         UV_vU037[Index]  := temp_U037;
         UV_vU062[Index]  := temp_U062;
         UV_vU045[Index]  := temp_U045;
         UV_vU047[Index]  := temp_U047;
         Inc(Index);
         
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

procedure TfrmLD4Q58.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q58.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q581 := TfrmLD4Q581.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q581.QuickRep.Preview
  else                                frmLD4Q581.QuickRep.Print;
end;

procedure TfrmLD4Q58.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

procedure TfrmLD4Q58.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  with QRCompositeReport do
  begin

  end;
end;

function TfrmLD4Q58.UF_hangmokList : String;
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

function TfrmLD4Q58.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
 