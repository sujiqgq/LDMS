//==============================================================================
// 프로그램 설명 : 혈액관리대장(연세대[대사증후군])
// 시스템        : 통합검진
// 수정일자      : 2007.02.22
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD7Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD7Q01 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryGulkwa: TQuery;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    CBox_gubn: TComboBox;
    Label3: TLabel;
    RadioGroup: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_vV001, UV_vV002, UV_vV003, UV_vV004, UV_vV005,
    UV_vV006, UV_vV007, UV_vV008, UV_vV009, UV_vV010,
    UV_vV011, UV_vV012 : Variant;

    UV_sCod_jisa : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }

  end;

var
  frmLD7Q01: TfrmLD7Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD7Q011;

{$R *.DFM}

procedure TfrmLD7Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vV001 := VarArrayCreate([0, 0], varOleStr);
      UV_vV002 := VarArrayCreate([0, 0], varOleStr);
      UV_vV003 := VarArrayCreate([0, 0], varOleStr);
      UV_vV004 := VarArrayCreate([0, 0], varOleStr);
      UV_vV005 := VarArrayCreate([0, 0], varOleStr);
      UV_vV006 := VarArrayCreate([0, 0], varOleStr);
      UV_vV007 := VarArrayCreate([0, 0], varOleStr);
      UV_vV008 := VarArrayCreate([0, 0], varOleStr);
      UV_vV009 := VarArrayCreate([0, 0], varOleStr);
      UV_vV010 := VarArrayCreate([0, 0], varOleStr);
      UV_vV011 := VarArrayCreate([0, 0], varOleStr);
      UV_vV012 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vV001, iLength);
      VarArrayRedim(UV_vV002, iLength);
      VarArrayRedim(UV_vV003, iLength);
      VarArrayRedim(UV_vV004, iLength);
      VarArrayRedim(UV_vV005, iLength);
      VarArrayRedim(UV_vV006, iLength);
      VarArrayRedim(UV_vV007, iLength);
      VarArrayRedim(UV_vV008, iLength);
      VarArrayRedim(UV_vV009, iLength);
      VarArrayRedim(UV_vV010, iLength);
      VarArrayRedim(UV_vV011, iLength);
      VarArrayRedim(UV_vV012, iLength);
   end;
end;

function TfrmLD7Q01.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD7Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_VarMemSet(0);

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   if UV_sCod_jisa = '15' then
   begin
      Label1.Caption := '지 사:';
      GP_ComboJisa(cmbCOD_JISA);
      GP_ComboDisplay(cmbCOD_JISA, copy(GV_sUserId,1,2), 2);
   end
   else
   begin
      Label1.Caption := '분주소:';
      cmbCod_jisa.Items.Add('연 구 소');
      cmbCod_jisa.Items.Add('지    사');
      cmbCod_jisa.ItemIndex := 1;
   end;
   
   CBox_gubn.ItemIndex := 0;
end;

procedure TfrmLD7Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vV001[DataRow-1];
      2 : Value := GF_DateFormat(UV_vV002[DataRow-1]);
      3 : Value := UV_vV003[DataRow-1];
      4 : Value := UV_vV004[DataRow-1];
      5 : Value := UV_vV005[DataRow-1];
      6 : Value := UV_vV006[DataRow-1];
      7 : Value := UV_vV007[DataRow-1];
      8 : Value := GF_DateFormat(UV_vV008[DataRow-1]);
      9 : Value := UV_vV009[DataRow-1];
     10 : Value := UV_vV010[DataRow-1];
     11 : Value := UV_vV011[DataRow-1];
     12 : Value := UV_vV012[DataRow-1];
   end;
end;

procedure TfrmLD7Q01.btnQueryClick(Sender: TObject);
var
   sSelect, sWhere, sOrderBy : String;
   index : Integer;

begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid & Chart 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sSelect := ''; sWhere := ''; sOrderBy := '';

   with qryGulkwa do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;
      sSelect := ' SELECT G.num_id, G.dat_gmgn, G.desc_name, J.desc_jisa As Desc_bjjs, G.num_samp, G.num_tel, G.desc_addr, G.cod_chuga, ' +
                 '        B.dat_bunju, B.num_bunju, K.desc_jisa, ' +
                 '        SUBSTRING(G.num_jumin,1,2) + ''.'' + SUBSTRING(G.num_jumin,3,2) + ''.'' + SUBSTRING(G.num_jumin,5,2) AS BIRTH ' +
                 ' FROM bunju B LEFT OUTER JOIN gumgin G ON B.cod_jisa = G.cod_jisa AND B.num_id = G.num_id AND B.dat_gmgn = G.dat_gmgn ' +
                 '      LEFT OUTER JOIN jisa    J  ON B.cod_bjjs = J.cod_jisa ' +
                 '      LEFT OUTER JOIN jisa    K  ON B.cod_jisa = K.cod_jisa ';

      if RadioButton1.Checked then sWhere := ' WHERE B.cod_bjjs = ''' + copy(cmbCOD_JISA.Text,1,2) + ''''
      else                         sWhere := ' WHERE B.cod_bjjs = ''' + UV_sCod_jisa + '''';

      sWhere := sWhere + ' AND B.dat_gmgn = ''' + mskDate.Text + '''';

      case CBox_gubn.ItemIndex of
         0 : sWhere := sWhere + ' AND (G.cod_chuga like ''%DI01%'' or G.cod_chuga like ''%DI02%'' or G.cod_chuga like ''%DI03%'')';
         1 : sWhere := sWhere + ' AND (G.cod_chuga like ''%DI01%'' or G.cod_chuga like ''%DI02%'')';
         2 : sWhere := sWhere + ' AND G.cod_chuga like ''%DI03%''';
      end;

      case RadioGroup.ItemIndex of
         0 : sOrderBy := ' ORDER BY K.desc_jisa, B.dat_bunju, B.num_bunju';
         1 : sOrderBy := ' ORDER BY K.desc_jisa, B.dat_bunju, CAST(G.num_samp AS INT)';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD7Q01', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);
         // DataSet의 값을 Variant변수로 이동
         while Not qryGulkwa.EOF do
         begin
            UV_vV001[index] := FieldByName('num_id').AsString;
            UV_vV002[index] := FieldByName('dat_gmgn').AsString;
            UV_vV003[index] := FieldByName('desc_name').AsString;
            UV_vV004[index] := FieldByName('BIRTH').AsString;
            UV_vV005[index] := FieldByName('num_tel').AsString;
            UV_vV006[index] := FieldByName('desc_addr').AsString;
            UV_vV007[index] := FieldByName('desc_bjjs').AsString + '[' +  FieldByName('desc_jisa').AsString + ']';
            UV_vV008[index] := FieldByName('dat_bunju').AsString;
            UV_vV009[index] := FieldByName('num_bunju').AsString;
            UV_vV010[index] := FieldByName('num_samp').AsString;
            if pos('DI01',FieldByName('cod_chuga').AsString) > 0 then
               UV_vV011[index] := '일반'
            else if pos('DI02',FieldByName('cod_chuga').AsString) > 0 then
               UV_vV011[index] := '종검'
            else if pos('DI03',FieldByName('cod_chuga').AsString) > 0 then
               UV_vV011[index] := '재검'
            else UV_vV011[index] := '';
            UV_vV012[index] := '';

            Next;
            Inc(index);
         end;

         // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
         Active := False;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD7Q01.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 없을경우의 처리
   if NewRow = 0 then exit;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // Grid Focus
   grdMaster.SetFocus;
end;

procedure TfrmLD7Q01.UP_Click(Sender: TObject);
begin
   inherited;

   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD7Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD7Q01.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD7Q011 := TfrmLD7Q011.Create(Self);
   frmLD7Q011.QuickRep.Preview;
//   frmLD7Q011.QuickRep.Print;
end;

end.
