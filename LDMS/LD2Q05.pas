//==============================================================================
// 프로그램 설명 : RIA 외주항목 List (분주실 전용)
// 시스템        : 통합검진
// 수정일자      : 2007.11.26
// 수정자        : 김승철
// 참고사항      :
//==============================================================================


unit LD2Q05;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD2Q05 = class(TfrmSingle)
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
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
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
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN,
    UV_vE001, UV_vE002, UV_vE003, UV_vE015, UV_vE040, UV_vC044, UV_vC049, UV_vT001, UV_vT002, UV_vT011, UV_vTT01,
    UV_vDesc_dc   : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD2Q05: TfrmLD2Q05;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q041, LD4Q042, LD2Q041, LD2Q042,
  LD2Q051;

{$R *.DFM}

procedure TfrmLD2Q05.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q05.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD2Q05.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vE001  := VarArrayCreate([0, 0], varOleStr);
      UV_vE002  := VarArrayCreate([0, 0], varOleStr);
      UV_vE003  := VarArrayCreate([0, 0], varOleStr);
      UV_vE015  := VarArrayCreate([0, 0], varOleStr);
      UV_vE040  := VarArrayCreate([0, 0], varOleStr);
      UV_vC044  := VarArrayCreate([0, 0], varOleStr);
      UV_vC049  := VarArrayCreate([0, 0], varOleStr);
      UV_vT001  := VarArrayCreate([0, 0], varOleStr);
      UV_vT002  := VarArrayCreate([0, 0], varOleStr);
      UV_vT011  := VarArrayCreate([0, 0], varOleStr);
      UV_vTT01  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vE001,  iLength);
      VarArrayRedim(UV_vE002,  iLength);
      VarArrayRedim(UV_vE003,  iLength);
      VarArrayRedim(UV_vE015,  iLength);
      VarArrayRedim(UV_vE040,  iLength);
      VarArrayRedim(UV_vC044,  iLength);
      VarArrayRedim(UV_vC049,  iLength);
      VarArrayRedim(UV_vT001,  iLength);
      VarArrayRedim(UV_vT002,  iLength);
      VarArrayRedim(UV_vT011,  iLength);
      VarArrayRedim(UV_vTT01,  iLength);
      VarArrayRedim(UV_vDesc_dc, iLength);
   end;
end;

function TfrmLD2Q05.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q05.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD2Q05.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vE001[DataRow-1];
      7 : Value := UV_vE002[DataRow-1];
      8 : Value := UV_vE003[DataRow-1];
      9 : Value := UV_vE015[DataRow-1];
     10 : Value := UV_vE040[DataRow-1];
     11 : Value := UV_vC044[DataRow-1];
     12 : Value := UV_vC049[DataRow-1];
     13 : Value := UV_vT001[DataRow-1];
     14 : Value := UV_vT002[DataRow-1];
     15 : Value := UV_vT011[DataRow-1];
     16 : Value := UV_vTT01[DataRow-1];
     17 : Value := UV_vDesc_dc[DataRow-1];
   end;
end;

procedure TfrmLD2Q05.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vE001, Sum_vE002, Sum_vE003, Sum_vE015, Sum_vE040, Sum_vC044, Sum_vC049, Sum_vT001, Sum_vT002, Sum_vT011, Sum_vTT01 : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   Sum_vE001 := 0; 
   Sum_vE002 := 0; 
   Sum_vE003 := 0; 
   Sum_vE015 := 0; 
   Sum_vE040 := 0; 
   Sum_vC044 := 0; 
   Sum_vC049 := 0; 
   Sum_vT001 := 0; 
   Sum_vT002 := 0;
   Sum_vT011 := 0;
   Sum_vTT01 := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT G.desc_name, G.num_jumin, D.desc_dc, G.cod_jisa, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga ' +
                 ' FROM Gulkwa K left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Danche D ON G.cod_dc = D.cod_dc ' +
                 '               left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha' +
                 '               left outer join profile_hm H ON G.cod_hul = H.cod_pf ';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' +
                          ' AND (K.Gubn_Part  = ''' + '04' + ''') ';
      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And K.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

       sWhere := sWhere + ' AND ((H.cod_hm = ''E001'') or (H.cod_hm = ''E002'') or (H.cod_hm = ''E003'')';
       sWhere := sWhere + ' or (H.cod_hm = ''E015'') or (H.cod_hm = ''E040'') or (H.cod_hm = ''C044'')';
       sWhere := sWhere + ' or (H.cod_hm = ''C049'') or (H.cod_hm = ''T001'') or (H.cod_hm = ''T002'')';
       sWhere := sWhere + ' or (H.cod_hm = ''T011'') or (H.cod_hm = ''TT01'') ';

       sWhere := sWhere + ' or (G.cod_chuga like ''%E001%'') or (G.cod_chuga like ''%E002%'') or (G.cod_chuga like ''%E003%'')';
       sWhere := sWhere + ' or (G.cod_chuga like ''%E015%'') or (G.cod_chuga like ''%E040%'') or (G.cod_chuga like ''%C044%'')';
       sWhere := sWhere + ' or (G.cod_chuga like ''%C049%'') or (G.cod_chuga like ''%T001%'') or (G.cod_chuga like ''%T002%'')';
       sWhere := sWhere + ' or (G.cod_chuga like ''%T011%'') or (G.cod_chuga like ''%TT01%'')) ';

      sGroupBy := ' group by G.desc_name, G.num_jumin, D.desc_dc, G.cod_jisa, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga ';
      sOrderBy := ' ORDER BY K.Num_Bunju';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q05', 'Q', 'N');
         For I := 1 to RecordCount do
//         while not qryBunju.Eof do
         Begin
            Gride.Progress := I;
            UP_VarMemSet(Index+1);
            UV_vNo[Index]  := IntToStr(Index+1);
            UV_VBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
            UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;
            UV_vJumin[Index] := qryBunju.FieldByName('num_jumin').AsString;
            UV_vJisa[Index] := qryBunju.FieldByName('cod_jisa').AsString;
            UV_vDesc_dc[Index] := qryBunju.FieldByName('desc_dc').AsString;
            
            // 항목(장비프로파일)추출..
            sJangbi_List := ''; sHul_List := '';
            if qryBunju.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            // 항목(혈액프로파일)추출..
            if qryBunju.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;
         
            // 항목(종양프로파일)추출..
            if qryBunju.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
                           sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //추가검사항목..
            iRet := GF_MulToSingle(qryBunju.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            //(특검)추가검사항목..
            iRet := GF_MulToSingle(qryBunju.FieldByName('Tcod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) = 'JJ' then
               begin
                  if vCod_chuga[iTemp] = 'JJ10' then
                     sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                  else
                     sJangbi_List := sJangbi_List + vCod_chuga[iTemp];
               end
               else
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            if pos('E001', sHul_List) > 0 then
            begin
               UV_vE001[Index] := '';
               Sum_vE001 := Sum_vE001 + 1;
            end
            else                               UV_vE001[Index] := '*';          //Sum_vE001
            if pos('E002', sHul_List) > 0 then
            begin
               UV_vE002[Index] := '';
               Sum_vE002 := Sum_vE002 + 1;
            end
            else                               UV_vE002[Index] := '*';          //Sum_vE002
            if pos('E003', sHul_List) > 0 then
            begin
               UV_vE003[Index] := '';
               Sum_vE003 := Sum_vE003 + 1;
            end
            else                               UV_vE003[Index] := '*';          //Sum_vE003
            if pos('E015', sHul_List) > 0 then
            begin
               UV_vE015[Index] := '';
               Sum_vE015 := Sum_vE015 + 1;
            end
            else                               UV_vE015[Index] := '*';          //Sum_vE015
            if pos('E040', sHul_List) > 0 then
            begin
               UV_vE040[Index] := '';
               Sum_vE040 := Sum_vE040 + 1;
            end
            else                               UV_vE040[Index] := '*';          //Sum_vE040
            if pos('C044', sHul_List) > 0 then
            begin
               UV_vC044[Index] := '';
               Sum_vC044 := Sum_vC044 + 1;
            end
            else                               UV_vC044[Index] := '*';          //Sum_vC044
            if pos('C049', sHul_List) > 0 then
            begin
               UV_vC049[Index] := '';
               Sum_vC049 := Sum_vC049  + 1;
            end
            else                               UV_vC049[Index] := '*';          //Sum_vC049
            if pos('T001', sHul_List) > 0 then
            begin
               UV_vT001[Index] := '';
               Sum_vT001 := Sum_vT001 + 1;
            end
            else                               UV_vT001[Index] := '*';          //Sum_vT001
            if pos('T002', sHul_List) > 0 then
            begin
               UV_vT002[Index] := '';
               Sum_vT002 := Sum_vT002 + 1;
            end
            else                               UV_vT002[Index] := '*';          //Sum_vT002
            if pos('T011', sHul_List) > 0 then
            begin
               UV_vT011[Index] := '';
               Sum_vT011 := Sum_vT011 + 1;
            end
            else                               UV_vT011[Index] := '*';          //Sum_vT011
            if pos('TT01', sHul_List) > 0 then
            begin
               UV_vTT01[Index] := '';
               Sum_vTT01 := Sum_vTT01 + 1;
            end
            else                               UV_vTT01[Index] := '*';          //Sum_vTT01
            Inc(Index);
            Next;
         End;

         // Table과 Disconnect
         Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '총  합  계';
         UV_vE001[Index]   := Sum_vE001;
         UV_vE002[Index]   := Sum_vE002;
         UV_vE003[Index]   := Sum_vE003;
         UV_vE015[Index]   := Sum_vE015;
         UV_vE040[Index]   := Sum_vE040;
         UV_vC044[Index]   := Sum_vC044;
         UV_vC049[Index]   := Sum_vC049;
         UV_vT001[Index]   := Sum_vT001;
         UV_vT002[Index]   := Sum_vT002;
         UV_vT011[Index]   := Sum_vT011;
         UV_vTT01[Index]   := Sum_vTT01;
         UV_vDesc_dc[Index] := '';
         Inc(index);
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

procedure TfrmLD2Q05.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD2Q05.UP_Click(Sender: TObject);
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

procedure TfrmLD2Q05.btnPrintClick(Sender: TObject);
begin
  inherited;
//  if RBtn_preveiw.Checked = True then QRCompositeReport.Preview
//  else                                QRCompositeReport.Print;

   frmLD2Q051 := TfrmLD2Q051.Create(Self); 
   if RBtn_preveiw.Checked = True then
    Try
        frmLD2Q051.QuickRep.Preview
    Finally
    frmLD2Q051.Free
    End
   Else
   begin
      if GF_MsgBox('Confirmation', '출력하시겠습니까?') then
      begin
         Try
            frmLD2Q051.QuickRep.Print;
         Finally
            frmLD2Q051.Free
         End;
      end else
   end;


end;

procedure TfrmLD2Q05.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

procedure TfrmLD2Q05.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD2Q051 := TfrmLD2Q051.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD2Q051.QuickRep);
  end;
end;

end.
