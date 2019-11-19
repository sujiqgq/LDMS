//==============================================================================
// 프로그램 설명 : CKMS 문진에 만성C형 간염자 체크 리스트 출력
// 시스템        : 통합검진
// 수정일자      : 2017.03.24
// 수정자        : 박수지
// 수정내용      : 새로 개발
// 참고사항      :
//==============================================================================
unit LD4Q18;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q18 = class(TfrmSingle)
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
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    RadioGroup1: TRadioGroup;
    cmbHM: TComboBox;
    Label6: TLabel;
    Cmb_gubn: TComboBox;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    MEdt_SampE: TMaskEdit;
    Label8: TLabel;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryBunju: TQuery;
    QRCompositeReport: TQRCompositeReport;
    qryTkgum: TQuery;
    qryPGulkwa: TQuery;
    qry_hangmok: TQuery;
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
    procedure QRCompositeReportAddReports(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, Uv_vRack, UV_vSampNo, UV_vBUN, UV_vName, UV_vhangmok, UV_vhangmok1, UV_vJumin, UV_vJisa, UV_vDesc_dc : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q18: TfrmLD4Q18;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q181;

{$R *.DFM}

procedure TfrmLD4Q18.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q18.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q18.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vRack  := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vhangmok := VarArrayCreate([0, 0], varOleStr);
      UV_vhangmok1 := VarArrayCreate([0, 0], varOleStr);

      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa    := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vRack,  iLength);
      VarArrayRedim(UV_vSampNo, iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vhangmok,  iLength);
      VarArrayRedim(UV_vhangmok1,  iLength);

      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
   end;
end;

function TfrmLD4Q18.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') OR (cmbHM.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q18.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   cmbHM.ItemIndex := 0;
end;

procedure TfrmLD4Q18.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRack[DataRow-1];
      3 : Value := UV_vSampNo[DataRow-1];
      4 : Value := UV_vBUN[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vhangmok1[DataRow-1];

   end;

end;

procedure TfrmLD4Q18.btnQueryClick(Sender: TObject);
var index : Integer; //, I, iRet, iTemp : Integer;
    sSelect, sWhere, sOrderBy, test : String; //sGroupBy,
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sHul_List ,UV_fGulkwa_05, UV_fGulkwa1_05, UV_fGulkwa2_05, UV_fGulkwa3_05: String;      // sJangbi_List,
//    vCod_chuga : Variant;
    Check :Boolean;
begin
   inherited;   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   grdMaster.Col[6].Heading := '만성C형-문진체크여부';

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT B.desc_rackno,G.desc_name, G.num_id, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, K.desc_glkwa1, K.desc_glkwa2, K.desc_glkwa3,' +
                 ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp, M.CMun7_YN, M.Mun7_3, M.Cmun7 ' +
                 ' FROM Gulkwa K left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Bunju  B ON K.cod_jisa = B.cod_jisa and K.num_id = B.num_id and K.dat_gmgn = B.dat_gmgn ' +
                 '               left outer join Danche D ON G.cod_dc = D.cod_dc ' +
                 '               left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha' +
                 '               left outer join Tot_munjin2010 M ON G.cod_jisa = M.cod_jisa and G.dat_gmgn = M.dat_gmgn and G.num_id = M.num_id ';
      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';

      sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + '''  AND ( K.Gubn_Part  = ''05'')';


      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(BunjuNo1.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju >= ''' + BunjuNo1.Text + '''';
         if Trim(BunjuNo2.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju <= ''' + BunjuNo2.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         1 : sOrderBy := ' ORDER BY B.Desc_rackno';
         2 : sOrderBy := ' ORDER BY K.Num_Bunju';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q18', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            sHul_List := '';
            sHul_List := UF_hangmokList;

            if ((cmbHM.Text = 'S016') AND (pos('S016', sHul_List) > 0)) then
            begin
               qryPGulkwa.close;
               qryPGulkwa.ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
               qryPGulkwa.ParamByName('num_id').AsString   := qryBunju.FieldByName('num_id').AsString;
               qryPGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryPGulkwa.ParamByName('gubn_part').AsString := '05';

               qryPGulkwa.open;

               if qryPGulkwa.RecordCount > 0 then
               begin
                  while not qryPGulkwa.Eof do
                  begin
                     Check := False;
                     if  qryPGulkwa.FieldByName('gubn_part').AsString ='05' then
                     begin
                          UV_fGulkwa1_05 := qryPGulkwa.FieldByName('DESC_GLKWA1').AsString;
                          UV_fGulkwa2_05 := qryPGulkwa.FieldByName('DESC_GLKWA2').AsString;
                          UV_fGulkwa3_05 := qryPGulkwa.FieldByName('DESC_GLKWA3').AsString;
                          GF_HulGulkwa('2', UV_fGulkwa1_05, UV_fGulkwa2_05, UV_fGulkwa3_05, UV_fGulkwa_05);
                     end;

                     UV_VRack[Index]   := qryBunju.FieldByName('Desc_rackno').AsString;
                     UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
                     UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;


                     UV_vName[Index]   := qryPGulkwa.FieldByName('desc_name').AsString;

                     if (cmbHM.Text = 'S016')then
                     begin
                             if Copy(FieldByName('CMun7').AsString, 3, 1) = '2' then
                             begin
                             UV_vhangmok1[Index] := 'Y';
                             Check := True;
                             end;
                     end;


                     if Check = True then
                     begin
                     UP_VarMemSet(Index+1);
                     UV_vNo[Index]  := IntToStr(Index+1);
                     Inc(Index);
                     end;
                     qryPGulkwa.Next;
                  end;
                  qryPGulkwa.Close;
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

procedure TfrmLD4Q18.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q18.UP_Click(Sender: TObject);
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
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q18.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q181 := TfrmLD4Q181.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q181.QuickRep.Preview
  else                                frmLD4Q181.QuickRep.Print;

end;

procedure TfrmLD4Q18.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
   Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD4Q18.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;

  with QRCompositeReport do
  begin

  end;
end;

function TfrmLD4Q18.UF_hangmokList : String;
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

function TfrmLD4Q18.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
