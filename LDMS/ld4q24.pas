//==============================================================================
// 프로그램 설명 : 혈액형분주비교현황
// 시스템        : 통합검진
// 수정일자      : 2011.01.25
// 수정자        : 송재호
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 혈액형분주비교현황
// 시스템        : 통합검진
// 수정일자      : 2011.02.10
// 수정자        : 송재호
// 수정내용      : 전결과 있을 경우만 조회 되도록 수정
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 혈액형분주비교현황
// 시스템        : 분석관리
// 수정일자      : 2012.05.08
// 수정자        : 유동구
// 수정내용      : 검진센터 선택 가능토록..
// 참고사항      : 한의 수원진단검사의학팀1200148
//==============================================================================
unit LD4Q24;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q24 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryCheck: TQuery;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    qrytkgum: TQuery;
    qryProfile: TQuery;
    Gride: TGauge;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    qryPGulkwa: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSamp, UV_vBunju, UV_vH023, UV_vH024, UV_vH023p, UV_vH024p, UV_vName : Variant;
    sHangmok : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    Procedure Hangmok_Set;
  public
    { Public declarations }
  end;

var
  frmLD4Q24: TfrmLD4Q24;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q021;

{$R *.DFM}

procedure TfrmLD4Q24.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp  := VarArrayCreate([0, 0], varOleStr);
      UV_vBunju := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vH023  := VarArrayCreate([0, 0], varOleStr);
      UV_vH024  := VarArrayCreate([0, 0], varOleStr);
      UV_vH023p := VarArrayCreate([0, 0], varOleStr);
      UV_vH024p := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vSamp,  iLength);
      VarArrayRedim(UV_vBunju, iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vH023,  iLength);
      VarArrayRedim(UV_vH024,  iLength);
      VarArrayRedim(UV_vH023p, iLength);
      VarArrayRedim(UV_vH024p, iLength);
   end;
end;

function TfrmLD4Q24.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q24.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 0;
   edtBDate.Text := GV_sToday;
end;

procedure TfrmLD4Q24.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSamp[DataRow-1];
      3 : Value := UV_vBunju[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vH023[DataRow-1];
      6 : Value := UV_vH023p[DataRow-1];
      7 : Value := UV_vH024[DataRow-1];
      8 : Value := UV_vH024p[DataRow-1];
   end;

   if (DataCol = 5) and (UV_vH023[DataRow-1] <> UV_vH023p[DataRow-1]) then
       grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
   if (DataCol = 7) and (UV_vH024[DataRow-1] <> UV_vH024p[DataRow-1]) then
       grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
end;

procedure TfrmLD4Q24.btnQueryClick(Sender: TObject);
var index, Index2, I : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sH023p, sH024p : String;
    sHulCheck : Boolean ;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   index := 0; Index2 := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select B.cod_bjjs, B.num_id, B.dat_bunju, B.num_bunju, B.cod_jisa, B.dat_gmgn, ' +
                 ' G.num_jumin, G.num_samp, G.cod_dc, G.desc_name, G.cod_hul, G.cod_jangbi, G.cod_cancer, G.gubn_tkgm, G.cod_chuga, G.gubn_hulgum, ' +
                 ' A.desc_glkwa1 ' +
                 ' FROM bunju B LEFT OUTER JOIN gumgin G ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                 ' LEFT OUTER JOIN gulkwa A ON B.cod_jisa = A.cod_jisa and B.num_id = A.num_id and B.dat_gmgn = A.dat_gmgn and A.gubn_part = ''02''';

      if Copy(cmbJisa.Text,1,2) = '15' then
      begin
         sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
         sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
      end
      else
      begin
         sWhere := ' Where B.Cod_bjjs in (''15'',''' + Copy(cmbJisa.Text,1,2) + ''')';
         sWhere := sWhere + ' and B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
         sWhere := sWhere + ' and B.Dat_Bunju = ''' + edtBDate.Text + '''';
         sWhere := sWhere + ' and G.gubn_hulgum <> ''1''';
      end;

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY CAST(B.num_bunju AS INT)';
      end
      else
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if qryBunju.RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q24', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            sHangmok := '';

            //단체 Check
            With qryCheck Do
            Begin
               Close;
               ParamByName('In_Cod_dc').AsString := qryBunju.FieldByname('Cod_dc').AsString;
               ParamByName('num_year').AsString  := copy(qryBunju.FieldByname('DaT_Gmgn').AsString,1,4);
               Open;
               sCheck := '';
               sHulCheck := False;
               If Recordcount > 0 Then
                  If (qryBunju.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                     (qryBunju.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        begin
                        sCheck := '99';
                        sHulCheck := true;
                        end;
               Close;
            End;

            //[2008.02.29] 검사추가에 있으면 검사들어감..
            if (pos('H023',qryBunju.FieldByName('cod_chuga').AsString) > 0) or
               (pos('H024',qryBunju.FieldByName('cod_chuga').AsString) > 0) then
               sCheck := '99';

            //어린이집
            If (FieldByname('Cod_hul').Asstring = 'B501') Or
               (FieldByname('Cod_hul').Asstring = 'B502') Or
               (FieldByname('Cod_hul').Asstring = '9023') Or
               (FieldByname('Cod_hul').Asstring = 'B507') Or
               (FieldByname('Cod_hul').Asstring = 'B508') Then
               sCheck := '99';

            If (Copy(FieldByname('Cod_Jangbi').Asstring,1,2) <> '') Then sCheck := '99';
            If qryBunju.FieldByName('Num_Bunju').AsInteger > 4999 Then sCheck := '99';

            //본원 분주시 지방센터 자체검진 제외...
            if (Copy(cmbJisa.Text,1,2) = '15') and (edtBDate.Text > '20120301') then
            begin
              if (qryBunju.FieldByName('Cod_jisa').AsString    <> '15') and
                 (qryBunju.FieldByName('gubn_hulgum').AsString =   '3') then sCheck := '';
            end;

            //===============================================================
            If sCheck = '99' Then Hangmok_Set;
            //===============================================================

            If (Pos('H023',sHangmok) > 0) or (Pos('H024',sHangmok) > 0) Then
            begin
               sH023p := ''; sH024p := '';
               sSelect := ''; sWhere := ''; sOrderBy := '';
               with qryPGulkwa do
               begin
                  close;
                  ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
                  ParamByName('num_jumin').AsString := qryBunju.FieldByName('num_jumin').AsString;
                  ParamByName('dat_bunju').AsString := edtBDate.text;
                  open;
                  if RecordCount > 0 then
                  begin
                     sH023p := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,133,6));
                     sH024p := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,139,6));
                  end;
                  close;
               end;

               if (sH023p <> '') or (sH024p <> '') then
               begin
                  UP_VarMemSet(Index);

                  UV_vNo[Index]    := IntToStr(Index+1);
                  UV_vSamp[Index]  := qryBunju.FieldByName('Num_Samp').AsString;
                  UV_vBunju[Index] := qryBunju.FieldByName('Num_bunju').AsString;
                  UV_vName[Index]  := qryBunju.FieldByName('desc_name').AsString;
                  If Pos('H023',sHangmok) > 0 then
                  begin
                     UV_vH023[Index]  := trim(copy(qryBunju.FieldByName('desc_glkwa1').AsString,133,6));
                     UV_vH023p[Index] := sH023p;
                  end;
                  if Pos('H024',sHangmok) > 0 then
                  begin
                     UV_vH024[Index]  := trim(copy(qryBunju.FieldByName('desc_glkwa1').AsString,139,6));
                     UV_vH024p[Index] := sH024p;
                  end;
                  Inc(Index);
               end;
            end;

            Gride.Progress := I;
            qryBunju.Next;
         end;
         Gride.Progress := RecordCount;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q24.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q24.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q021 := TfrmLD4Q021.Create(Self);
     frmLD4Q021.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q021 := TfrmLD4Q021.Create(Self);
     frmLD4Q021.QuickRep.Print;
  end;
end;

Procedure TfrmLD4Q24.Hangmok_Set;
Var I, J, K : Integer;
    sHulCheck : Boolean ;
Begin
      Inherited;
//PRofile Reading

      //단체 Check
      With qryCheck Do
      Begin
         Close;
         ParamByName('In_Cod_dc').AsString := qryBunju.FieldByname('Cod_dc').AsString;
         ParamByName('num_year').AsString  := copy(qryBunju.FieldByname('DaT_Gmgn').AsString,1,4);
         Open;
         If Recordcount > 0 Then
            If (qryBunju.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
               (qryBunju.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
               sHulCheck := True;
         Close;
      End;


      With qryProfile do
         Begin
            Close;
//장비
            ParamByName('In_Cod_Pf').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//혈액
            if sHulCheck = true then
            begin
            ParamByName('In_Cod_Pf').AsString := qryBunju.FieldByName('Cod_Hul').AsString;
            Open;
            sHulCheck := False;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
            end;
//Cancer
            ParamByName('In_Cod_Pf').AsString := qryBunju.FieldByName('Cod_Cancer').AsString;
            Open;
            For I := 1 To Recordcount Do
                Begin
                   sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                   Next;
                End;
             Close;
         End;
//추가코드
         sHangmok := sHangmok + qryBunju.FieldByName('Cod_chuga').AsString;
//특검
         If qryBunju.FieldByName('Gubn_tkgm').AsString <> '' Then
            With qryTkgum Do
               Begin
                  Close;
                  ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByname('Cod_jisa').AsString;
                  ParamByName('In_Dat_gmgn').AsString := qryBunju.FieldByname('DaT_gmgn').AsString;
                  ParamByName('In_NuM_Jumin').AsString := qryBunju.FieldByname('NuM_Jumin').AsString;
                  Open;
                  If RecordCount > 0 Then
                     Begin
                        I := Length(FieldByName('Cod_prf').AsString);
                        J := Round(I / 4);
                        For K := 0 To J - 1 Do
                           Begin
                              With qryProfile do
//특검 Profile
                                 Begin
                                    Close;
                                    ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4);
                                    Open;
                                    For I := 1 To Recordcount Do
                                       Begin
                                          sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                                          Next;
                                       End;
                                    Close;
                                 End;
                           End;
                      End;
                   sHangmok := sHangmok + FieldByName('Cod_chuga').AsString;
                   Close;
                End;
End;

end.
