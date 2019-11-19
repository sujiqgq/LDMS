//==============================================================================
// 개발일자      : 2015.06.27
// 개발자        : 곽윤설
// 수정사항      : 요 10종 추가
// 참고사항      : 한의 재단특검출장파트1500104
//==============================================================================
unit LD5Q12;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD5Q12 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    RadioGroup1: TRadioGroup;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    Btn_BacordPnt: TBitBtn;
    qryCheck: TQuery;
    Label4: TLabel;
    qryProfile: TQuery;
    Panel3: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    RB_BUNJU: TRadioButton;
    RB_SAMP: TRadioButton;
    BunjuNo1: TMaskEdit;
    BunjuNo2: TMaskEdit;
    SampleNo1: TMaskEdit;
    SampleNo2: TMaskEdit;
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
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUNJU, UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn, UV_vHangmok, UV_vtkgum_yh : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD5Q12: TfrmLD5Q12;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q121, LD5Q122;

{$R *.DFM}

procedure TfrmLD5Q12.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD5Q12.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;
         grdMaster.Refresh;
      end;
   end;
end;

procedure TfrmLD5Q12.UP_VarMemSet(iLength : Integer);
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
      UV_vtkgum_yh  := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_vtkgum_yh,  iLength);
   end;
end;

function TfrmLD5Q12.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q12.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD5Q12.grdMasterCellLoaded(Sender: TObject; DataCol,
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
     10 : Value := UV_vtkgum_yh[DataRow-1];
   end;
end;

procedure TfrmLD5Q12.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp, Temp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sCheck_List, sHul_List : String;
    vDesc_profile, chuga : String;
    vCod_chuga : Variant;
    bCkeck : Boolean;
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

      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, B.num_bunju, G.gubn_nosin, G.gubn_adult, G.gubn_life,   ' +
                 '        G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, G.cod_hul,  ' +
                 '        G.cod_jangbi, G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp FROM Gumgin G with(nolock) ' +
                 ' left outer join Bunju B with(nolock) ON G.num_id= B.num_id and G.dat_gmgn= B.dat_gmgn                                       ' +
                 ' left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc                                                                ' +
                 ' left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn  ' +
                 '                                          and G.gubn_tkgm = T.gubn_cha                                                       ';

      sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND G.Dat_gmgn = ''' + edtBDate.Text + '''';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

      If RB_BUNJU.Checked then
      begin
         if Trim(BunjuNo1.Text) <> '' then
            sWhere := sWhere + ' And B.num_bunju >= ''' + BunjuNo1.Text + ''' ';
         if Trim(BunjuNo2.Text) <> '' then
            sWhere := sWhere + ' And B.num_bunju <= ''' + BunjuNo2.Text + ''' ';
      end
      else
      begin
         if Trim(sampleno1.Text) <> '' then
            sWhere := sWhere + ' And G.Num_Samp >= ''' + sampleno1.Text + ''' ';
         if Trim(sampleno2.Text) <> '' then
            sWhere := sWhere + ' And G.Num_Samp <= ''' + sampleno2.Text + ''' ';
      end;

      sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, B.num_bunju, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.cod_dc, ' +
                  ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga, G.num_samp ';

      sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      GP_query_log(qryBunju, 'LD5Q12', 'Q', 'N');

      Open;
      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         for I := 1 to RecordCount do
         begin
            sCheck_List := '';
            Gride.Progress := I;
            sHul_List := UF_hangmokList;

            if (pos('U001',sHul_List) > 0) and (pos('U002',sHul_List) > 0) and
               (pos('U003',sHul_List) > 0) and (pos('U004',sHul_List) > 0) and
               (pos('U005',sHul_List) > 0) and (pos('U006',sHul_List) > 0) and
               (pos('U007',sHul_List) > 0) and (pos('U008',sHul_List) > 0) and
               (pos('U009',sHul_List) > 0) and (pos('U010',sHul_List) > 0) then
            begin
               sCheck_List := sCheck_List + 'U001, U002, U003, U004, U005, U006, U007, U008, U009, U010';
            end;

            chuga := ''; vDesc_profile := '';
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
               qryTkgum.Active := True;
               if qryTkgum.RecordCount > 0 then
               begin
                  chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
                  iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_chuga);
                  for Temp := 0 to iRet - 1 do
                  begin
                    qryProfile.Active := False;
                    qryProfile.ParamByName('cod_pf').AsString := vCod_chuga[Temp];
                    qryProfile.Active := True;
                    if qryProfile.RecordCount > 0 then
                    begin
                       vDesc_profile := vDesc_profile + qryProfile.FieldByName('desc_pf').AsString + ',';
                    end;
                  end;        //for(qryTkgum) end;
               end;           //if(qryTkgum) end;
            end;

            if sCheck_List <> '' then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]       := IntToStr(Index+1);
               UV_VBUNJU[Index]    := qryBunju.FieldByName('num_bunju').AsString;
               UV_vName[Index]     := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[Index]    := (Copy(qryBunju.FieldByName('num_jumin').AsString,1,6)) +'-'+ (Copy(qryBunju.FieldByName('num_jumin').AsString,7,1))
                                      + '******';
               UV_vJisa[Index]     := qryBunju.FieldByName('cod_jisa').AsString;
               UV_vDesc_dc[Index]  := qryBunju.FieldByName('desc_dc').AsString;
               UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
               UV_vDat_gmgn[Index] := qryBunju.FieldByName('Dat_gmgn').AsString;
               UV_vHangmok[Index]  := sCheck_List;
               UV_vtkgum_yh[Index] := vDesc_profile;
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

procedure TfrmLD5Q12.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;
end;

procedure TfrmLD5Q12.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
  if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD5Q12.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD5Q121 := TfrmLD5Q121.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q121.QuickRep.Preview
  else                                frmLD5Q121.QuickRep.Print;
end;

procedure TfrmLD5Q12.FormActivate(Sender: TObject);
begin
   inherited;
  // edtBdate.SetFocus;
end;

function TfrmLD5Q12.UF_hangmokList : String;
var
   sTemp, sHmCode, sTkgum_pf, sCod_pf :string;
   vCod_chuga : Variant;
   iRet, i : integer;
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
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
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
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
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
               if (copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ') and (pos(FieldByName('COD_HM').AsString,sTemp) = 0) then
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
         while not qryTkgum.Eof do
         begin
            if Trim(qryTkgum.FieldByName('cod_prf').AsString) <> '' then
            begin
                sTkgum_pf := qryTkgum.FieldByName('cod_prf').AsString;
                while length(sTkgum_pf) >= 4 do
                begin
                    sCod_pf := copy(sTkgum_pf, 1, 4);
                    sTkgum_pf := copy(sTkgum_pf,5,length(sTkgum_pf));
                    qryProfileG.Active := False;
                    qryProfileG.ParamByName('cod_pf').AsString := sCod_pf;
                    qryProfileG.Active := True;
                    if qryProfileG.RecordCount > 0 then
                    begin
                       while not qryProfileG.Eof do
                       begin
                          if pos(qryProfileG.FieldByName('cod_hm').AsString, sHmCode) = 0 then
                             sHmCode := sHmCode + qryProfileG.FieldByName('cod_hm').AsString;
                          qryProfileG.Next;
                       end;
                    end;
                    qryProfileG.Active := False;
                    sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
                end; //end of while
            end; //end of if
            qryTkgum.Next;
         end; //end of while
      end; //end of qryTkgum recordcount
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if (copy(vCod_chuga[i],1,2) <> 'JJ') and (pos(vCod_chuga[i],sTemp) = 0) then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   result := sTemp;
end;

function TfrmLD5Q12.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q12.Btn_BacordPntClick(Sender: TObject);
begin
  inherited;
  frmLD5Q122 := TfrmLD5Q122.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD5Q122.QuickRep.Preview
  else                                frmLD5Q122.QuickRep.Print;
end;



end.
