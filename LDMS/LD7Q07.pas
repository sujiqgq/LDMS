//==============================================================================
// 수정일자      : 201601226
// 수정자        : 박수지
// 수정내용      : 한의 재단진단면역학팀1600061
// 참고사항      :
//==============================================================================

unit LD7Q07;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD7Q07 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
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
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    Label5: TLabel;
    qryProfile: TQuery;
    mskDate_start: TDateEdit;
    Label8: TLabel;
    mskDate_end: TDateEdit;
    btnBdate: TSpeedButton;
    btnEdate: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN, 
    UV_vSE89,UV_vSE90, UV_vSampNo : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD7Q07: TfrmLD7Q07;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD7Q071;

{$R *.DFM}

procedure TfrmLD7Q07.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD7Q07.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD7Q07.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vSE89  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE90  := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vJisa,   iLength);
      VarArrayRedim(UV_vBUN,    iLength);
      VarArrayRedim(UV_vSE89,   iLength);
      VarArrayRedim(UV_vSE90,   iLength);
      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD7Q07.UF_Condition : Boolean;
begin
   Result := True;
   // 조회조건중 필수항목이 입력되었는지 Check
   if (mskDate_start.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end
   else if (mskDate_end.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD7Q07.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD7Q07.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSampNo[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vSE89[DataRow-1];
      6 : Value := UV_vSE90[DataRow-1];
   end;
   if (Trim(UV_vSE89[DataRow - 1]) <> '') and (Trim(UV_vSE89[DataRow - 1]) <> '결과미입력')  then
   begin
        if (strtOfloat(UV_vSE89[DataRow - 1]) <> 0) and  (strtOfloat(UV_vSE89[DataRow - 1]) >= 2.5)  then grdMaster.CellColor[5,DataRow] := clYellow;
   end;
   if (Trim(UV_vSE90[DataRow - 1]) <> '') and (Trim(UV_vSE90[DataRow - 1]) <> '결과미입력') then
   begin
        if (strtOfloat(UV_vSE90[DataRow - 1]) <> 0) and  (strtOfloat(UV_vSE90[DataRow - 1]) >= 1.5)  then grdMaster.CellColor[6,DataRow] := clYellow;
   end;
end;

procedure TfrmLD7Q07.btnQueryClick(Sender: TObject);
var index, I : integer;
    sSelect, sWhere, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    UV_fGulkwa_07, UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07  : String;
    vCod_chuga : Variant;
    Check_05 :Boolean;
    Check_07 :Boolean;
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
      sSelect := ' SELECT  B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B with(nolock) inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  >= ''' + mskDate_start.Text + '''  AND B.Dat_Bunju  <= ''' + mskDate_end.Text + ''' ' ;
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND (Gubn_Part  = ''09''))';

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
      end;
     // sOrderBy := ' ORDER BY B.Desc_Rackno ';



      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD7Q07', 'Q', 'N');
         while not qryBunju.Eof do
         Begin

            sHul_List := '';

            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;

            sHul_List := UF_hangmokList;
            if ((Pos('SE89' ,sHul_List)> 0) OR (Pos('SE90' ,sHul_List)> 0)) Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_sbunju').AsString := mskDate_start.text;
               qryGulkwa.ParamByName('dat_ebunju').AsString := mskDate_end.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         UV_vSE89[Index] :='';
                         UV_vSE90[Index] :='';

                         if  qryGulkwa.FieldByName('gubn_part').AsString  ='09'  then
                         begin
                             UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                             UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                             UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                             GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE89', sHul_List) > 0)  then
                         begin
                               if Trim(Copy(UV_fGulkwa, 343,  6)) <> '' then
                               begin
                                  UV_vSE89[Index] := Trim(Copy(UV_fGulkwa, 343,  6));
                               end
                               else UV_vSE89[Index] := '결과미입력';
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE90', sHul_List) > 0)  then
                         begin
                               if Trim(Copy(UV_fGulkwa, 349,  6)) <> '' then
                               begin
                                  UV_vSE90[Index] := Trim(Copy(UV_fGulkwa, 349,  6));
                               end
                               else UV_vSE90[Index] := '결과미입력';
                          end;
                   qryGulkwa.Next;
                   end;

                         UV_vNo[Index]     := index +1;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
                         UP_VarMemSet(Index+1);
                         inc(index);

              END;
               qryGulkwa.Close;
              end;
           qryBunju.Next;
         End;

         // Table과 Disconnect
         qryBunju.Close;
         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vSampNo[Index]  := '';
         Inc(index);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecOrdCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index-1;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD7Q07.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(mskDate_start)
   else if Sender = btnEdate then GF_CalendarClick(mskDate_end);
end;


procedure TfrmLD7Q07.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   mskDate_start.SetFocus;


end;

function TfrmLD7Q07.UF_hangmokList : String;
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

function TfrmLD7Q07.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD7Q07.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD7Q071 := TfrmLD7Q071.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD7Q071.QuickRep.Preview
  else                                frmLD7Q071.QuickRep.Print;

end;

end.
