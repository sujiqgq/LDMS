//==============================================================================
// 프로그램 설명 : 혈액형분주현황
// 시스템        : 통합검진
// 수정일자      : 2002.08.22
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 혈액형분주현황
// 시스템        : 통합검진
// 수정일자      : 2012.03.13
// 수정자        : 송재호
// 수정내용      : 검진센터 선택 가능토록..
// 참고사항      : 한의 재단혈액학팀1200022
//==============================================================================
// 프로그램 설명 : 혈액형분주현황
// 수정일자      : 2013.06.13
// 수정자        : 김희석
// 수정내용      : 검진일자 조회 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.23
// 수정자        : 곽윤설
// 수정내용      : 혈액프로파일 모두 읽어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2016.10.31
// 수정자        : 곽윤설
// 수정내용      : 장비프로파일 안에 h023, h024 조회, 혈액프로파일은 단체체크 확인 후 조회
// 참고사항      :
//==============================================================================
// 수정일자      : 2016.11.16
// 수정자        : 박수지
// 수정내용      : 검진일 조회시에도 샘플번호, 분주 번호로 조회 가능하도록 수정
// 참고사항      : 혈액학 손혜정선임.
//==============================================================================
unit LD4Q02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    qrygumgin: TQuery;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryCheck: TQuery;
    qryBunju: TQuery;
    qrytkgum: TQuery;
    qryProfile: TQuery;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Chk_15: TCheckBox;
    Chk_12: TCheckBox;
    Chk_11: TCheckBox;
    Chk_43: TCheckBox;
    Chk_71: TCheckBox;
    Chk_61: TCheckBox;
    Chk_51: TCheckBox;
    CheckBox: TCheckBox;
    Label4: TLabel;
    grd_New: TtsGrid;
    mskSt_date: TDateEdit;
    btnGmgnF: TSpeedButton;
    Gride: TGauge;
    Chk_bunju: TRadioButton;
    chk_gum: TRadioButton;
    GroupBox1: TGroupBox;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Cmb_gubn2: TComboBox;
    Label1: TLabel;
    MEdt_SampS2: TMaskEdit;
    MEdt_SampE2: TMaskEdit;
    Label2: TLabel;
    mskFrom2: TMaskEdit;
    Label6: TLabel;
    mskTo2: TMaskEdit;
    Label8: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure Ck_NewClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vH023, UV_vH024, UV_vName, UV_vNum_samp, UV_vNum_Bunju ,UV_vDat_gmgn, UV_vGulkwa : Variant;
    sHangmok : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    Procedure Hangmok_Set;
  public
    { Public declarations }
  end;

var
  frmLD4Q02: TfrmLD4Q02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q021, LD4Q022;

{$R *.DFM}

procedure TfrmLD4Q02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vH023  := VarArrayCreate([0, 0], varOleStr);
      UV_vH024  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vGulkwa   := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vH023,  iLength);
      VarArrayRedim(UV_vH024,  iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vNum_Bunju, iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
      VarArrayRedim(UV_vGulkwa,   iLength);
   end;
end;

function TfrmLD4Q02.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') and (mskST_date.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q02.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 1;
   Cmb_gubn2.ItemIndex := 1;
   edtBDate.Text := GV_sToday;
   mskST_date.Text := GV_sToday;

   if copy(GV_sUserId,1,2)='15' then
   begin
      Label4.Visible  := True;
      Chk_15.Visible  := True;
      Chk_12.Visible  := True;
      Chk_11.Visible  := True;
      Chk_43.Visible  := True;
      Chk_71.Visible  := True;
      Chk_61.Visible  := True;
      Chk_51.Visible  := True;
      CheckBox.Visible := True;
   end;

   grdMaster.Visible := false;
   grd_New.Visible   := True;
end;

procedure TfrmLD4Q02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;


      // 자룔를 화면에 조회
      case DataCol of
         1 : Value := UV_vNo[DataRow-1];
         2 : Value := UV_vNum_samp[DataRow-1];
         3 : Value := UV_vName[DataRow-1];
         4 : Value := UV_vNum_Bunju[DataRow-1];
         5 : Value := UV_vGulkwa[DataRow-1];
         8 : Value := UV_vDat_gmgn[DataRow-1];
      end;

end;

procedure TfrmLD4Q02.btnQueryClick(Sender: TObject);
var index, I : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sJisa : String;
    sHulCheck : Boolean ;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';  sJisa := '';
   index := 0; 

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      if Chk_bunju.Checked= true then
      begin
           sSelect := ' select B.desc_rackno, B.cod_bjjs, B.num_id, B.dat_bunju, B.num_bunju, B.cod_jisa, B.dat_gmgn, G.num_jumin, G.num_samp, G.gubn_hulgum, A.desc_glkwa1 ' +
                      ' FROM bunju B with(nolock) LEFT OUTER JOIN gumgin G with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn      ' +
                      ' Left Outer Join gulkwa A On B.cod_jisa = A.cod_jisa and B.num_id = A.num_id and B.dat_gmgn = A.dat_gmgn ';

           //[2012.02.22]자체검진 처리시 추가 내용...
           if (edtBDate.Text > '20120301') then
           begin
              if Copy(cmbJisa.Text,1,2) = '15' then
              begin
                 sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
                 sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
                 sWhere := sWhere + ' AND A.gubn_part = ''02''';

                 if Chk_15.Checked then sJisa := sJisa + '''15'',';
                 if Chk_12.Checked then sJisa := sJisa + '''12'',';
                 if Chk_11.Checked then sJisa := sJisa + '''11'',';
                 if Chk_43.Checked then sJisa := sJisa + '''43'',';
                 if Chk_71.Checked then sJisa := sJisa + '''71'',';
                 if Chk_61.Checked then sJisa := sJisa + '''61'',';
                 if Chk_51.Checked then sJisa := sJisa + '''51'',';

                 if sJisa <> '' then
                    sWhere := sWhere + ' And B.Cod_jisa in (' + copy(sJisa,1, Length(sJisa)-1) + ')'
                 else
                 begin
                    showmessage('검진센터를 선택해주세요!');
                    exit;
                 end;
              end
              else
              begin
                 sWhere := ' Where B.Cod_bjjs in (''15'',''' + Copy(cmbJisa.Text,1,2) + ''')';
                 sWhere := sWhere + ' and B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
                 sWhere := sWhere + ' and B.Dat_Bunju = ''' + edtBDate.Text + '''';
                 sWhere := sWhere + ' and G.gubn_hulgum <> ''1''';
                 sWhere := sWhere + ' AND A.gubn_part = ''02''';
              end;
           end
           else
           begin
              sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
              sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
              sWhere := sWhere + ' AND A.gubn_part = ''02''';
           end;
           sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
           sWhere := sWhere + ' AND A.gubn_part = ''02''';

           if Cmb_gubn.Text = '분주번호' then
           begin
              if Trim(mskFrom.Text) <> '' then
                 sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
              if Trim(mskTo.Text) <> '' then
                 sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

              sOrderBy := ' ORDER BY CAST(B.num_bunju AS INT), G.num_samp ';
           end
           else if Cmb_gubn.Text = '샘플번호' then
           begin
              if Trim(MEdt_SampS.Text) <> '' then
                 sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
              if Trim(MEdt_SampE.Text) <> '' then
                 sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';

              sOrderBy := ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju '
           end;

      end
      else if Chk_gum.Checked = true then
      begin
           sSelect := ' select B.desc_rackno, B.cod_bjjs, g.num_id, B.dat_bunju, B.num_bunju, g.cod_jisa, g.dat_gmgn, G.num_jumin, G.num_samp, G.gubn_hulgum, A.desc_glkwa1 ' +
                      ' FROM gumgin G with(nolock) LEFT OUTER JOIN bunju B with(nolock) ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                      ' Left Outer Join gulkwa A On B.cod_jisa = A.cod_jisa and B.num_id = A.num_id and B.dat_gmgn = A.dat_gmgn ';
                 sWhere := ' Where G.DAt_gmgn = ''' + mskST_date.Text + ''' ';
                // sWhere := sWhere + ' AND A.gubn_part = ''02''';

                 if Chk_15.Checked then sJisa := sJisa + '''15'',';
                 if Chk_12.Checked then sJisa := sJisa + '''12'',';
                 if Chk_11.Checked then sJisa := sJisa + '''11'',';
                 if Chk_43.Checked then sJisa := sJisa + '''43'',';
                 if Chk_71.Checked then sJisa := sJisa + '''71'',';
                 if Chk_61.Checked then sJisa := sJisa + '''61'',';
                 if Chk_51.Checked then sJisa := sJisa + '''51'',';

                 if sJisa <> '' then
                    sWhere := sWhere + ' and   G.Cod_jisa in (' + copy(sJisa,1, Length(sJisa)-1) + ')'
                 else
                 begin
                    showmessage('검진센터를 선택해주세요!');
                    exit;
                 end;
                 if Cmb_gubn2.Text = '분주번호' then
                 begin
                    if Trim(mskFrom2.Text) <> '' then
                       sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom2.Text + '''';
                    if Trim(mskTo2.Text) <> '' then
                       sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo2.Text + '''';

                    sOrderBy := ' ORDER BY CAST(B.num_bunju AS INT), G.num_samp ';
                 end
                 else if Cmb_gubn2.Text = '샘플번호' then
                 begin
                    if Trim(MEdt_SampS2.Text) <> '' then
                       sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS2.Text + '''';
                    if Trim(MEdt_SampE2.Text) <> '' then
                       sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE2.Text + '''';

                    sOrderBy := ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju '
                 end;
                  if Chk_gum.Checked = false then sWhere := sWhere + ' AND A.gubn_part = ''02''';

           sOrderBy := ' ORDER BY  G.dat_gmgn, G.num_samp'
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q02', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            //검진Check
            With qryGumgin Do
            Begin
               Close;
               ParamByName('In_Cod_Jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
               ParamByName('In_Dat_gmgn').AsString := qryBunju.FieldByName('DaT_Gmgn').AsString;
               ParamByName('In_NuM_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
               Open;
               sHangmok := '';

               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGumgin.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGumgin.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  sHulCheck := False;
                  If Recordcount > 0 Then
                     If (qryGumgin.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGumgin.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        begin
                        sCheck := '99';
                        sHulCheck := true;
                        end;
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',qryGumgin.FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',qryGumgin.FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';
 //20140523 곽윤설
           //    if (FieldByname('Cod_hul').Asstring <> '') then sCheck := '99';
{
               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then
                  sCheck := '99';
}


               If (Copy(FieldByname('Cod_Jangbi').Asstring,1,2) <> '') Then sCheck := '99';

               If (Chk_bunju.Checked =true) and (qryBunju.FieldByName('Num_Bunju').AsInteger > 4999)  Then sCheck := '99';

               //본원 분주시 지방센터 자체검진 제외...
               if (Copy(cmbJisa.Text,1,2) = '15') and (edtBDate.Text > '20120301') then
               begin
                 if (qryBunju.FieldByName('Cod_jisa').AsString    <> '15') and
                    (qryBunju.FieldByName('gubn_hulgum').AsString =   '3') then sCheck := '';
               end;

               //===============================================================
               If sCheck = '99' Then Hangmok_Set;
               //===============================================================

               If Pos('H023',sHangmok) > 0 Then
               Begin
                  UP_VarMemSet(Index);
                  if Cmb_gubn.Text = '분주번호' then
                     UV_VH023[Index] := qryBunju.FieldByName('Num_Bunju').AsString
                  else
                     UV_VH023[Index] := qryBunju.FieldByName('Num_Samp').AsString;

                   UV_vNo[Index]   := IntToStr(Index+1);
                   UV_vName[Index] := qryGumgin.FieldByName('desc_name').AsString;

                   if Pos('H024',sHangmok) > 0 Then
                     if Cmb_gubn.Text = '분주번호' then
                        UV_VH024[Index] := qryBunju.FieldByName('Num_Bunju').AsString
                     else
                        UV_VH024[Index] := qryBunju.FieldByName('Num_Samp').AsString;

                   UV_vNum_Samp[Index]  := qryBunju.FieldByName('Num_Samp').AsString;
                   UV_vNum_Bunju[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
                   UV_vDat_gmgn[Index] := qryBunju.FieldByName('Dat_gmgn').AsString;
                   UV_vGulkwa[Index]   := copy(qryBunju.FieldByName('desc_glkwa1').AsString,133,6);
                   Inc(Index);
               End
               else if Pos('H024',sHangmok) > 0 Then
               Begin
                  if Pos('H023',sHangmok) < 1 Then
                     UP_VarMemSet(Index);

                  if Cmb_gubn.Text = '분주번호' then
                     UV_VH024[Index] := qryBunju.FieldByName('Num_Bunju').AsString
                  else
                     UV_VH024[Index] := qryBunju.FieldByName('Num_Samp').AsString;
                  UV_vNo[Index]   := IntToStr(Index+1);
                  UV_vName[Index] := qryGumgin.FieldByName('desc_name').AsString;

                  UV_vNum_Samp[Index]  := qryBunju.FieldByName('Num_Samp').AsString;
                  UV_vNum_Bunju[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
                  UV_vDat_gmgn[Index] := qryBunju.FieldByName('Dat_gmgn').AsString;
                  
                  Inc(Index);
               End;
            End;

            Gride.Progress := i;
            Next;
         End;
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
   grd_New.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q02.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
   if Sender = btnGmgnF then GF_CalendarClick(mskST_date);
   if Sender = CheckBox then
   begin
      if CheckBox.Checked then
      begin
         CheckBox.Caption := 'UnCheck All';
         Chk_15.Checked := True; Chk_12.Checked := True;
         Chk_11.Checked := True; Chk_43.Checked := True;
         Chk_71.Checked := True; Chk_61.Checked := True;
         Chk_51.Checked := True;
      end
      else
      begin
         CheckBox.Caption := 'Check All';
         Chk_15.Checked := False; Chk_12.Checked  := False;
         Chk_11.Checked := False; Chk_43.Checked  := False;
         Chk_71.Checked := False; Chk_61.Checked  := False;
         Chk_51.Checked := False; 
      end;
   end;
   if Sender = Chk_gum then
   begin
      edtBDate.Enabled    :=False;
      edtBDate.ReadOnly   :=True;
      MEdt_SampS.ReadOnly :=True;
      MEdt_SampE.ReadOnly :=True;
      mskFrom.ReadOnly    :=True;
      mskTo.ReadOnly      :=True;
      mskST_date.ReadOnly :=False;
      mskST_date.Enabled  :=True;
   end;
   if Sender = Chk_bunju then
   begin
      edtBDate.Enabled    :=True;
      edtBDate.ReadOnly   :=False;
      MEdt_SampS.ReadOnly :=False;
      MEdt_SampE.ReadOnly :=False;
      mskFrom.ReadOnly    :=False;
      mskTo.ReadOnly      :=False;
      mskST_date.ReadOnly :=True;
      mskST_date.Enabled  :=False;
   end;


end;

procedure TfrmLD4Q02.btnPrintClick(Sender: TObject);
begin
  inherited;


     if RBtn_preveiw.Checked = True then
     begin
        frmLD4Q022 := TfrmLD4Q022.Create(Self);
        frmLD4Q022.QuickRep.Preview;
     end
     else
     begin
        frmLD4Q022:= TfrmLD4Q022.Create(Self);
        frmLD4Q022.QuickRep.Print;
     end;

end;

Procedure TfrmLD4Q02.Hangmok_Set;
Var I, J, K : Integer;
    sHulCheck : Boolean ;
Begin
      Inherited;
//PRofile Reading
      //단체 Check
      With qryCheck Do
      Begin
         Close;
         ParamByName('In_Cod_dc').AsString := qryGumgin.FieldByname('Cod_dc').AsString;
         ParamByName('num_year').AsString  := copy(qryGumgin.FieldByname('DaT_Gmgn').AsString,1,4);
         Open;
         sHulCheck := False;
         If Recordcount > 0 Then
            If (qryGumgin.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
               (qryGumgin.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
               begin
               sHulCheck := true;
               end;
         Close;
      End;

      With qryProfile do
         Begin
            Close;
//장비
            ParamByName('In_Cod_Pf').AsString := qryGumgin.FieldByName('Cod_jangbi').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
//혈액
            if sHulCheck = true then  // 혈액 프로파일 안에 있는 혈액형 검사는 혈액형 단체 체크 후 풀기 20161025 수지
            begin
            ParamByName('In_Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Hul').AsString;
            Open;
            For I := 1 To Recordcount Do
               Begin
                  sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                  Next;
               End;
            Close;
            end;
//Cancer
            ParamByName('In_Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Cancer').AsString;
            Open;
            For I := 1 To Recordcount Do
                Begin
                   sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                   Next;
                End;
             Close;
         End;
//추가코드
         sHangmok := sHangmok + qryGumgin.FieldByName('Cod_chuga').AsString;
//특검
         If qryGumgin.FieldByName('Gubn_tkgm').AsString <> '' Then
            With qryTkgum Do
               Begin
                  Close;
                  ParamByName('In_Cod_jisa').AsString := qryGumgin.FieldByname('Cod_jisa').AsString;
                  ParamByName('In_Dat_gmgn').AsString := qryGumgin.FieldByname('DaT_gmgn').AsString;
                  ParamByName('In_NuM_Jumin').AsString := qryGumgin.FieldByname('NuM_Jumin').AsString;
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

procedure TfrmLD4Q02.Ck_NewClick(Sender: TObject);
begin
  inherited;

      grdMaster.Visible := False;
      grd_New.Visible   := True;


end;


end.
