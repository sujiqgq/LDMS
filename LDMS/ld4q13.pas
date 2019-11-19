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
unit LD4Q13;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q13 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    qrygumgin: TQuery;
    qrytkgum: TQuery;
    qryProfile: TQuery;
    qryCheck: TQuery;
    Label1: TLabel;
    btnPrint_2: TSpeedButton;
    qryProfile_H025: TQuery;
    Chk_15: TCheckBox;
    Chk_12: TCheckBox;
    Chk_11: TCheckBox;
    Chk_43: TCheckBox;
    Chk_71: TCheckBox;
    Chk_61: TCheckBox;
    Chk_51: TCheckBox;
    Chk_45: TCheckBox;
    Chk_52: TCheckBox;
    Chk_34: TCheckBox;
    Chk_41: TCheckBox;
    Chk_Etc: TCheckBox;
    CheckBox: TCheckBox;
    Label4: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnPrint_2Click(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007 : Variant;
    sHangmok : String;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    Procedure Hangmok_Set;
    Procedure Hangmok_Set2;

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q13: TfrmLD4Q13;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q131, LD4Q132;

{$R *.DFM}

procedure TfrmLD4Q13.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,    iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
   end;
end;

function TfrmLD4Q13.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q13.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 1;
   edtBDate.Text := GV_sToday;

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
      Chk_45.Visible  := True;
      Chk_52.Visible  := True;
      Chk_34.Visible  := True;
      Chk_41.Visible  := True;
      Chk_Etc.Visible := True;
      CheckBox.Visible := True;
   end;
end;

procedure TfrmLD4Q13.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := UV_v003[DataRow-1];
      4 : Value := UV_v004[DataRow-1];
      5 : Value := UV_v005[DataRow-1];
      6 : Value := UV_v006[DataRow-1];
      7 : Value := UV_v007[DataRow-1];
   end;
end;

procedure TfrmLD4Q13.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sJisa : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sJisa := '';
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
      sSelect := ' select B.cod_bjjs, B.num_id, B.dat_bunju, B.num_bunju, B.cod_jisa, B.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.cod_jisa, G.desc_name, G.num_jumin ' +
                 ' FROM gulkwa B LEFT OUTER JOIN gumgin G ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ';

      sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
      sWhere := sWhere + ' And B.gubn_part = ''02''';

      if copy(GV_sUserId,1,2)='15' then
      begin
         if Chk_15.Checked then sJisa := sJisa + '''15'',';
         if Chk_12.Checked then sJisa := sJisa + '''12'',';
         if Chk_11.Checked then sJisa := sJisa + '''11'',';
         if Chk_43.Checked then sJisa := sJisa + '''43'',';
         if Chk_71.Checked then sJisa := sJisa + '''71'',';
         if Chk_61.Checked then sJisa := sJisa + '''61'',';
         if Chk_51.Checked then sJisa := sJisa + '''51'',';
         if Chk_45.Checked then sJisa := sJisa + '''45'',';
         if Chk_52.Checked then sJisa := sJisa + '''52'',';
         if Chk_34.Checked then sJisa := sJisa + '''34'',';
         if Chk_41.Checked then sJisa := sJisa + '''41'',';
         if Chk_Etc.Checked then sJisa := sJisa + '''62'',''80'',''80'',';
         
         if sJisa <> '' then
            sWhere := sWhere + ' And B.Cod_jisa in (' + copy(sJisa,1, Length(sJisa)-1) + ')'
         else
         begin
            showmessage('검진센터를 선택해주세요!');
            exit;
         end;
      end;

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY B.cod_jisa, CAST(B.num_bunju AS INT)';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY B.cod_jisa, CAST(G.num_samp AS INT) '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q13', 'Q', 'N');
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            UP_VarMemSet(Index);

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
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGumgin.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGumgin.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

//[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',qryGumgin.FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',qryGumgin.FieldByName('cod_chuga').AsString) > 0) then
                  sCheck := '99';

//어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then
                  sCheck := '99';
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then
                  sCheck := '99';
               If (sCheck = '99') Or (qryBunju.FieldByName('Num_Bunju').AsInteger > 4999) Then
                  Hangmok_Set;
            end;

            UV_v001[Index] := IntToStr(Index+1);
            UV_v002[Index] := qryBunju.FieldByName('cod_jisa').AsString;
            UV_v003[Index] := qryBunju.FieldByName('num_samp').AsString;
            UV_v004[Index] := qryBunju.FieldByName('num_bunju').AsString;
            If (Pos('H023',sHangmok) > 0) or (Pos('H024',sHangmok) > 0) Then
               UV_v005[Index] := '★' + qryBunju.FieldByName('desc_name').AsString
            else
               UV_v005[Index] := qryBunju.FieldByName('desc_name').AsString;

            Hangmok_Set2;
            if (Pos('H025',sHangmok) > 0) Then
               UV_v005[Index] := UV_v005[Index] + '※';

            UV_v006[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
            UV_v007[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

            INC(Index);
            Next;
         end;
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

procedure TfrmLD4Q13.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate)
   else if Sender = CheckBox then
   begin
      if CheckBox.Checked then
      begin
         CheckBox.Caption := 'UnCheck All';
         Chk_15.Checked := True; Chk_12.Checked := True;
         Chk_11.Checked := True; Chk_43.Checked := True;
         Chk_71.Checked := True; Chk_61.Checked := True;
         Chk_51.Checked := True; Chk_45.Checked := True;
         Chk_52.Checked := True; Chk_34.Checked := True;
         Chk_41.Checked := True; Chk_Etc.Checked := True;
      end
      else
      begin
         CheckBox.Caption := 'Check All';
         Chk_15.Checked := False; Chk_12.Checked  := False;
         Chk_11.Checked := False; Chk_43.Checked  := False;
         Chk_71.Checked := False; Chk_61.Checked  := False;
         Chk_51.Checked := False; Chk_45.Checked  := False;
         Chk_52.Checked := False; Chk_34.Checked  := False;
         Chk_41.Checked := False; Chk_Etc.Checked := False;
      end;
   end;
end;

procedure TfrmLD4Q13.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q131 := TfrmLD4Q131.Create(Self);
     frmLD4Q131.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q131 := TfrmLD4Q131.Create(Self);
     frmLD4Q131.QuickRep.Print;
  end;
end;

Procedure TfrmLD4Q13.Hangmok_Set;
Var I, J, K : Integer;
Begin
   Inherited;
//PRofile Reading
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
      ParamByName('In_Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Hul').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
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
   begin
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
//특검 Profile
               With qryProfile do
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
   end;
End;

Procedure TfrmLD4Q13.Hangmok_Set2;
Var I, J, K : Integer;
Begin
   Inherited;
//PRofile Reading
   With qryProfile_H025 do
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
      ParamByName('In_Cod_Pf').AsString := qryGumgin.FieldByName('Cod_Hul').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
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
   begin
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
//특검 Profile
               With qryProfile_H025 do
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
   end;
End;

procedure TfrmLD4Q13.btnPrint_2Click(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q132 := TfrmLD4Q132.Create(Self);
     frmLD4Q132.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q132 := TfrmLD4Q132.Create(Self);
     frmLD4Q132.QuickRep.Print;
  end;
end;

end.
