//==============================================================================
// 프로그램 설명 : 항목별검사리스트
// 시스템        : 통합검진
// 수정일자      : 2009.09.01
// 수정자        : 유동구
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 수정일자      : 2011.03.15
// 수정자        : 김희석
// 수정내용      : U026 추가
//==============================================================================
// 수정일자      : 2011.06.29
// 수정자        : 송재호
// 수정내용      : 인터페이스 장비의 접수 위치 나오도록 수정
//==============================================================================
// 수정일자      : 2011.06.30
// 수정자        : 송재호
// 수정내용      : 샘플번호로 조회하는 기능 없애고 검진센터별로 조회할 수 있도록 수정.
//==============================================================================
// 수정일자      : 2012.03.13
// 수정자        : 김희석
// 수정내용      : U047 추가
//==============================================================================
// 수정일자      : 2012.11.27
// 수정자        : 유동구
// 수정내용      : U051 추가
//==============================================================================
unit LD4Q25;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q25 = class(TfrmSingle)
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
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Chk_U020: TCheckBox;
    Chk_U023: TCheckBox;
    Chk_U024: TCheckBox;
    Chk_U028: TCheckBox;
    Chk_U029: TCheckBox;
    Chk_U030: TCheckBox;
    Chk_U031: TCheckBox;
    Gride: TGauge;
    Chk_U032: TCheckBox;
    Chk_U033: TCheckBox;
    Chk_U037: TCheckBox;
    Chk_U043: TCheckBox;
    qryNo_hangmok: TQuery;
    Chk_U044: TCheckBox;
    Chk_U045: TCheckBox;
    Chk_U026: TCheckBox;
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
    Bevel2: TBevel;
    Chk_U047: TCheckBox;
    RGrp_sort: TRadioGroup;
    Chk_U051: TCheckBox;
    qryHangmokList: TQuery;
    Chk_U017: TCheckBox;
    Button1: TButton;
    Chk_U053_5: TCheckBox;
    Chk_U057: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btn_formtecPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008 : Variant;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;    
    function  UF_HangmokCheck : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q25: TfrmLD4Q25;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q251,LD4Q252;

{$R *.DFM}

procedure TfrmLD4Q25.UP_VarMemSet(iLength : Integer);
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
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_v008,  iLength);
   end;
end;

function TfrmLD4Q25.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q25.FormCreate(Sender: TObject);
begin
   inherited;

   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   edtBDate.Text := dateTostr(date);
end;

procedure TfrmLD4Q25.grdMasterCellLoaded(Sender: TObject; DataCol,
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
      8 : Value := UV_v008[DataRow-1];
   end;
end;

procedure TfrmLD4Q25.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sHangmok : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sCheck := '';  sHangmok := '';
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
      sSelect := ' select B.cod_bjjs, B.num_id, B.dat_bunju, B.num_bunju, B.cod_jisa, B.dat_gmgn, G.num_jumin, G.num_id, G.num_samp, G.cod_jisa, G.desc_name, G.num_jumin, ' +
                 '       G.Cod_jangbi, G.cod_hul,   G.Cod_cancer, G.cod_chuga, G.gubn_neawon, ' +
                 '       G.gubn_nosin, G.gubn_nsyh, G.gubn_tkgm, G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_life, G.gubn_lfyh, ' +
                 '       G.gubn_agab,  G.gubn_agyh, G.num_jumin, T.cod_prf,    T.cod_chuga As Tcod_chuga ' +
                 ' FROM gulkwa B LEFT OUTER JOIN gumgin G ON B.cod_jisa = G.cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Tkgum T ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';
      sWhere := ' Where B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And B.Dat_Bunju = ''' + edtBDate.Text + '''';
      if Chk_U026.Checked then
          sWhere := sWhere + ' And B.gubn_part = ''08'''
      else
          sWhere := sWhere + ' And B.gubn_part = ''03''';

      if Chk_15.Checked then sCheck := sCheck + '''15'',';
      if Chk_12.Checked then sCheck := sCheck + '''12'',';
      if Chk_11.Checked then sCheck := sCheck + '''11'',';
      if Chk_43.Checked then sCheck := sCheck + '''43'',';
      if Chk_71.Checked then sCheck := sCheck + '''71'',';
      if Chk_61.Checked then sCheck := sCheck + '''61'',';
      if Chk_51.Checked then sCheck := sCheck + '''51'',';
      if Chk_45.Checked then sCheck := sCheck + '''45'',';
      if Chk_52.Checked then sCheck := sCheck + '''52'',';
      if Chk_34.Checked then sCheck := sCheck + '''34'',';
      if Chk_41.Checked then sCheck := sCheck + '''41'',';
      if Chk_Etc.Checked then sCheck := sCheck + '''62'',''80'',''80'',';

      if sCheck <> '' then
         sWhere := sWhere + ' And B.cod_jisa in (' + copy(sCheck,1, Length(sCheck)-1) + ')';

      if Trim(mskFrom.Text) <> '' then
         sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
      if Trim(mskTo.Text) <> '' then
         sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

      case RGrp_sort.ItemIndex of
         0 : sOrderBy := ' ORDER BY B.cod_jisa, CAST(B.num_bunju AS INT)';
         1 : sOrderBy := ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q25', 'Q', 'N');
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            application.ProcessMessages;
            
            sHangmok := '';
            sHangmok := UF_HangmokCheck;
            if (Chk_U026.Checked) and (pos('U026', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U026';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U020.Checked) and (pos('U020', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U020';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U023.Checked) and (pos('U023', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U023';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U024.Checked) and (pos('U024', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U024';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U028.Checked) and (pos('U028', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U028';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U029.Checked) and (pos('U029', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U029';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U030.Checked) and (pos('U030', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U030';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;


            if (Chk_U031.Checked) and (pos('U031', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U031';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U032.Checked) and (pos('U032', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U032';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U033.Checked) and (pos('U033', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U033';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U037.Checked) and (pos('U037', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U037';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;
            if (Chk_U017.Checked) and (pos('U017', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U017';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;
            if (Chk_U043.Checked) and (pos('U043', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U043';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U044.Checked) and (pos('U044', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U044';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U045.Checked) and (pos('U045', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U045';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;

            if (Chk_U047.Checked) and (pos('U047', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U047';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;
            if (Chk_U057.Checked) and (pos('U057', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U057';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;
            if (Chk_U051.Checked) and (pos('U051', sHangmok) > 0) then
            begin
               UP_VarMemSet(Index);
               UV_v001[Index] := IntToStr(Index+1);
               UV_v002[Index] := 'U051';
               UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
               UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
               UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);

               INC(Index);
            end;
            if (Chk_U053_5.Checked) then
               begin
               if (pos('U053', sHangmok) > 0) then
                  begin
                  UP_VarMemSet(Index);
                  UV_v001[Index] := IntToStr(Index+1);
                  UV_v002[Index] := 'U053';
                  UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
                  UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
                  UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
                  UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
                  UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
                  UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
                  INC(Index);
               end;
               if (pos('U054', sHangmok) > 0) then
                  begin
                  UP_VarMemSet(Index);
                  UV_v001[Index] := IntToStr(Index+1);
                  UV_v002[Index] := 'U054';
                  UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
                  UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
                  UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
                  UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
                  UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
                  UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
                  INC(Index);
               end;
               if (pos('U055', sHangmok) > 0) then
                  begin
                  UP_VarMemSet(Index);
                  UV_v001[Index] := IntToStr(Index+1);
                  UV_v002[Index] := 'U055';
                  UV_v003[Index] := qryBunju.FieldByName('cod_jisa').AsString;
                  UV_v004[Index] := qryBunju.FieldByName('num_samp').AsString;
                  UV_v005[Index] := qryBunju.FieldByName('num_bunju').AsString;
                  UV_v006[Index] := qryBunju.FieldByName('desc_name').AsString;
                  UV_v007[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
                  UV_v008[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
                  INC(Index);
               end;
            end;
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

procedure TfrmLD4Q25.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);

   if Sender = CheckBox then
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

procedure TfrmLD4Q25.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q251 := TfrmLD4Q251.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q251.QuickRep.Preview
  else                                frmLD4Q251.QuickRep.Print;
end;
procedure TfrmLD4Q25.btn_formtecPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q252 := TfrmLD4Q252.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q252.QuickRep.Preview
  else                                frmLD4Q252.QuickRep.Print;
end;

function TfrmLD4Q25.UF_HangmokCheck : String;
var sALL_HangmokList, sHul_List, sEtc_Hangmok_hm,
    sSelect, sWhere1, sWhere2, sOderby : string;
    i : Integer;
begin
   Result := ''; sALL_HangmokList := ''; sHul_List := ''; sEtc_Hangmok_hm := '';
   sSelect := ''; sWhere1 := ''; sWhere2 := ''; sOderby := '';
   i := 0;

//------------------------------------------------------------------------------
//검사항목 나열...
//------------------------------------------------------------------------------
   sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOderby := '';

   //장비검사(항목나열)리스트 추출...
   //------------------------------------------------------------------------
   if Trim(qryBunju.FieldByName('cod_prf').AsString) <> '' then
   begin
      sWhere1 := Trim(qryBunju.FieldByName('Cod_Jangbi').AsString) + ''',''' +
                 Trim(qryBunju.FieldByName('Cod_Hul').AsString)    + ''',''' +
                 Trim(qryBunju.FieldByName('Cod_Cancer').AsString) + ''',''';
      For i := 1 to (length(qryBunju.FieldByName('cod_prf').AsString) div 4) do
      begin
         if i = (length(Trim(qryBunju.FieldByName('cod_prf').AsString)) div 4) then sWhere1 := sWhere1 + copy(Trim(qryBunju.FieldByName('cod_prf').AsString), (i*4)-3, 4)
         else                                                                       sWhere1 := sWhere1 + copy(Trim(qryBunju.FieldByName('cod_prf').AsString), (i*4)-3, 4) + ''',''';
      end;
   end
   else
   begin
      sWhere1 := Trim(qryBunju.FieldByName('Cod_Jangbi').AsString) + ''',''' +
                 Trim(qryBunju.FieldByName('Cod_Hul').AsString)    + ''',''' +
                 Trim(qryBunju.FieldByName('Cod_Cancer').AsString) + ''',''';
   end;

   sEtc_Hangmok_hm := qryBunju.FieldByName('COD_CHUGA').AsString;

   // 노신유형Display
   if Trim(qryBunju.FieldByName('Gubn_Nosin').AsString) = '1' then
      sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '1', StrToInt(copy(qryBunju.FieldByName('Gubn_nsyh').AsString,1,1)));
   //성인병유형Display
   if Trim(qryBunju.FieldByName('Gubn_adult').AsString) = '1' then
      sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '4', StrToInt(copy(qryBunju.FieldByName('Gubn_adyh').AsString,1,1)));
   //간염유형Display
   if Trim(qryBunju.FieldByName('Gubn_agab').AsString) = '1' then
      sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '5', StrToInt(copy(qryBunju.FieldByName('Gubn_agyh').AsString,1,1)));
   //생애전환기유형Display
   if Trim(qryBunju.FieldByName('Gubn_life').AsString) = '1' then
      sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '7', StrToInt(copy(qryBunju.FieldByName('Gubn_lfyh').AsString,1,1)));

   if Trim(qryBunju.FieldByName('Tcod_chuga').AsString) <> '' then
   begin
      sWhere2 := sWhere2 + ''',''';
      For i := 1 to (length(qryBunju.FieldByName('Tcod_chuga').AsString) div 4) do
      begin
         if (i = (length(qryBunju.FieldByName('Tcod_chuga').AsString) div 4)) and (Trim(sEtc_Hangmok_hm) = '') then sWhere2 := sWhere2 + copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i*4)-3, 4)
         else                                                                                                       sWhere2 := sWhere2 + copy(qryBunju.FieldByName('Tcod_chuga').AsString, (i*4)-3, 4) + ''',''';
      end;
   end;

   if Trim(sEtc_Hangmok_hm) <> '' then
   begin
      sWhere2 := sWhere2 + ''',''';
      For i := 1 to (length(sEtc_Hangmok_hm) div 4) do
      begin
         if i = (length(sEtc_Hangmok_hm) div 4) then sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4)
         else                                        sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
      end;
   end;

   with qryHangmokList do
   begin
      qryHangmokList.Active := False;

      sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + qryBunju.FieldByName('dat_gmgn').AsString + '''';
      sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
      sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part  ) UNION ( ';
      sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
      sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
      sSelect := sSelect + '   AND dat_apply <= ''' + qryBunju.FieldByName('dat_gmgn').AsString + ''' )';

      sOderby := ' ORDER BY H.gubn_part Desc, P.cod_hm ';

      qryHangmokList.SQL.Clear;
      qryHangmokList.SQL.Add(sSelect + sOderby);
      qryHangmokList.Active := True;

      if qryHangmokList.RecordCount > 0 then
      begin
         while not Eof do
         begin
            sALL_HangmokList := sALL_HangmokList + qryHangmokList.FieldByName('cod_hm').AsString;

            if qryHangmokList.FieldByName('gubn_part').AsInteger < 10 then
               sHul_List := sHul_List + qryHangmokList.FieldByName('cod_hm').AsString;

            qryHangmokList.Next;
         end;
      end;
   end;

   Result := sHul_List;
end;

function  TfrmLD4Q25.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

end.
