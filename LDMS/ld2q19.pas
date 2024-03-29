//==============================================================================
// 프로그램 설명 : 정선미 선임 요청 프로그램 (항목 비교)
// 시스템        : 통합검진
// 수정일자      : 2018.11.13
// 수정자        : 박수지
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD2Q19;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD2Q19 = class(TfrmSingle)
    qryBunju: TQuery;
    QRCompositeReport: TQRCompositeReport;
    qryGulkwa: TQuery;
    qryBunju2: TQuery;
    qryGulkwa2: TQuery;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    qryBunju3: TQuery;
    pnlCond: TPanel;
    Label7: TLabel;
    btnBdate: TSpeedButton;
    cmbJisa: TComboBox;
    mskDate: TDateEdit;
    mskbunju1: TMaskEdit;
    Panel2: TPanel;
    Label5: TLabel;
    tsGrid3: TtsGrid;
    Label1: TLabel;
    tsGrid4: TtsGrid;
    mskDate2: TDateEdit;
    btnBdate2: TSpeedButton;
    mskbunju2: TMaskEdit;
    btn_Pro: TBitBtn;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    tsGrid1: TtsGrid;
    tsGrid2: TtsGrid;
    procedure FormCreate(Sender: TObject);
    procedure tsGrid1CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure btn_ProClick(Sender: TObject);
    procedure tsGrid2CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);
    procedure tsGrid3CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);
    procedure tsGrid4CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);
    procedure mskbunju2Exit(Sender: TObject);
    procedure mskbunju1Exit(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vcod_hm, UV_vdesc_hm, UV_vcod_hm2, UV_vdesc_hm2, UV_vcod_hm3, UV_vdesc_hm3, UV_vcod_hm4, UV_vdesc_hm4: Variant;
      // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_VarMemSet2(iLength : Integer);
    procedure UP_VarMemSet3(iLength : Integer);
    procedure UP_VarMemSet4(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD2Q19: TfrmLD2Q19;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD2Q19.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD2Q19.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         tsGrid3.refresh;
         tsGrid4.refresh;
      end;
   end;
end;

procedure TfrmLD2Q19.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vcod_hm           := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_hm          := VarArrayCreate([0, 0], varOleStr);
  end
  else
  begin
     // 이미 생성된 변수들의 크기를 조정
     VarArrayRedim(UV_vcod_hm,     iLength);
     VarArrayRedim(UV_vdesc_hm,    iLength);
  end;
end;
procedure TfrmLD2Q19.UP_VarMemSet2(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vcod_hm2          := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_hm2         := VarArrayCreate([0, 0], varOleStr);
  end
  else
  begin
     // 이미 생성된 변수들의 크기를 조정
     VarArrayRedim(UV_vcod_hm2,    iLength);
     VarArrayRedim(UV_vdesc_hm2,   iLength);
  end;
end;
procedure TfrmLD2Q19.UP_VarMemSet3(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vcod_hm3          := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_hm3         := VarArrayCreate([0, 0], varOleStr);
  end
  else
  begin
     // 이미 생성된 변수들의 크기를 조정
     VarArrayRedim(UV_vcod_hm3,    iLength);
     VarArrayRedim(UV_vdesc_hm3,   iLength);
  end;
end;
procedure TfrmLD2Q19.UP_VarMemSet4(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vcod_hm4          := VarArrayCreate([0, 0], varOleStr);
      UV_vdesc_hm4         := VarArrayCreate([0, 0], varOleStr);
  end
  else
  begin
     // 이미 생성된 변수들의 크기를 조정
     VarArrayRedim(UV_vcod_hm4,    iLength);
     VarArrayRedim(UV_vdesc_hm4,   iLength);
  end;
end;



function TfrmLD2Q19.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q19.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;
procedure TfrmLD2Q19.tsGrid1CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
     // 자룔를 화면에 조회
     case DataCol of
        1 : Value := UV_vcod_hm[DataRow-1];
        2 : Value := UV_vdesc_hm[DataRow-1];
     end;
end;
procedure TfrmLD2Q19.tsGrid2CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vcod_hm2[DataRow-1];
      2 : Value := UV_vdesc_hm2[DataRow-1];
   end;
end;
procedure TfrmLD2Q19.tsGrid3CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
  // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vcod_hm3[DataRow-1];
      2 : Value := UV_vdesc_hm3[DataRow-1];
   end;
end;
procedure TfrmLD2Q19.tsGrid4CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
  // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vcod_hm4[DataRow-1];
      2 : Value := UV_vdesc_hm4[DataRow-1];
   end;
end;

procedure TfrmLD2Q19.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate  then GF_CalendarClick(mskDate);
   if Sender = btnBdate2 then GF_CalendarClick(mskDate2);

end;

procedure TfrmLD2Q19.btn_ProClick(Sender: TObject);
var index3, i, j, index4 : Integer;
    chk_jung : Boolean;
begin
  inherited;
  Index3 := 0;
  Index4 := 0;
  tsGrid3.Rows := 0;
  UP_VarMemSet3(0);
  tsGrid4.Rows := 0;
  UP_VarMemSet4(0);
  chk_jung := false;

  for j := 1 to tsGrid2.Rows  do
  begin
     for i := 1 to tsGrid1.Rows  do
     begin
        if tsGrid2.Cell[1,j] = tsGrid1.Cell[1,i] then
        begin
            UP_VarMemSet3(Index3+1);
            begin
            UV_vcod_hm3[Index3]   := tsGrid2.Cell[1,j] ;
            UV_vdesc_hm3[Index3]  := tsGrid2.Cell[2,j] ;
            Inc(index3);
            chk_jung := true;
            end;
        end;
     end;
     if chk_jung = false then
     begin
        UP_VarMemSet4(Index4+1);
        begin
        UV_vcod_hm4[Index4]   := tsGrid2.Cell[1,j] ;
        UV_vdesc_hm4[Index4]  := tsGrid2.Cell[2,j] ;
        Inc(index4);
        chk_jung := False;
        end;
     end;
     chk_jung := false;
  end;

   tsGrid3.Rows := index3;
   tsGrid4.Rows := index4;
end;

procedure TfrmLD2Q19.mskbunju2Exit(Sender: TObject);
var index2, index3, i, iRet, iTemp, num , j, k: Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
begin
  inherited;
  sSelect := ''; sWhere := ''; sOrderBy := '';
  Index2 := 0;

  // Grid 초기화
  tsGrid2.Rows := 0;
  UP_VarMemSet2(0);
  with qryGulkwa do
   begin
     Close;
     sSelect := ' select G.dat_gmgn, G.num_id, G.cod_Jisa, G.desc_name, G.num_jumin from bunju B' ;
     sSelect := sSelect+ ' Join gumgin G on B.cod_jisa = G.cod_jisa and B.dat_gmgn = G.dat_gmgn and B.num_id = G.num_id' ;
     sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
     sWhere  := sWhere + ' AND B.Dat_Bunju = ''' + mskDate.Text + ''' ';
     sWhere  := sWhere + ' AND B.num_bunju = ''' + mskbunju1.Text + ''' ';

     qryGulkwa.SQL.Clear;
     qryGulkwa.SQL.Add(sSelect + sWhere + sOrderBy);
     qryGulkwa.Open;
   end;

   with qryGulkwa2 do
   begin
     Close;
     sSelect := ' select G.dat_gmgn, G.num_id, G.cod_Jisa, G.desc_name, G.num_jumin from bunju B' ;
     sSelect := sSelect + ' Join gumgin G on B.cod_jisa = G.cod_jisa and B.dat_gmgn = G.dat_gmgn and B.num_id = G.num_id' ;
     sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
     sWhere  := sWhere + ' AND B.Dat_Bunju = ''' + mskDate2.Text + ''' ';
     sWhere  := sWhere + ' AND B.num_bunju = ''' + mskbunju2.Text + ''' ';

     qryGulkwa2.SQL.Clear;
     qryGulkwa2.SQL.Add(sSelect + sWhere + sOrderBy);
     qryGulkwa2.Open;
   end;

   if (qryGulkwa.FieldByName('num_id').AsString) = (qryGulkwa2.FieldByName('num_id').AsString) then
   begin
      Label2.Caption := qryGulkwa.FieldByName('desc_name').AsString;
      Label3.Caption := copy(qryGulkwa.FieldByName('num_jumin').AsString,1,6)+'-'+copy(qryGulkwa.FieldByName('num_jumin').AsString,7,7);
   end
   else
   begin
      GF_MsgBox('Warning', '수검자명이 같지 않습니다.');
      exit;
   end;


   with qryBunju2 do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select distinct H.cod_hm, H.desc_kor from hangmok H JOIN TF_Get_HangmokList(''' + qryGulkwa2.FieldByName('cod_jisa').AsString + ''','''
                                                                 + qryGulkwa2.FieldByName('num_id').AsString   + ''','''
                                                                 + qryGulkwa2.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                 + ' ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
      sWhere :=   ' WHERE (h.gubn_part <> ''06'') and (h.gubn_part <> ''11'') and (h.gubn_part <> ''12'') and (h.gubn_part <> ''13'') and (h.gubn_part <> ''14'')';

      qryBunju2.SQL.Clear;
      qryBunju2.SQL.Add(sSelect + sWhere + sOrderBy);
      qryBunju2.Open;

      if qryBunju2.RecordCount > 0 then
      begin
         For I := 1 to qryBunju2.RecordCount do
         Begin
               UP_VarMemSet2(Index2+1);
               begin
               UV_vcod_hm2[Index2]  := qryBunju2.FieldByName('cod_hm').AsString;
               UV_vdesc_hm2[Index2]  := qryBunju2.FieldByName('desc_kor').AsString;
               Inc(Index2);
               end;
            Next;
         End;
         // Table과 Disconnect
         Close;
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
   tsGrid2.Rows := index2;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD2Q19.mskbunju1Exit(Sender: TObject);
var index, i, iRet, iTemp, num , j, k: Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
begin
  inherited;
  sSelect := ''; sWhere := ''; sOrderBy := '';
  Index := 0;
     // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   tsGrid1.Rows := 0;
   UP_VarMemSet(0);

   with qryGulkwa do
   begin
     Close;
     sSelect := ' select G.dat_gmgn, G.num_id, G.cod_Jisa, G.desc_name, G.num_jumin from bunju B' ;
     sSelect := sSelect+ ' Join gumgin G on B.cod_jisa = G.cod_jisa and B.dat_gmgn = G.dat_gmgn and B.num_id = G.num_id' ;
     sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
     sWhere  := sWhere + ' AND B.Dat_Bunju = ''' + mskDate.Text + ''' ';
     sWhere  := sWhere + ' AND B.num_bunju = ''' + mskbunju1.Text + ''' ';

     qryGulkwa.SQL.Clear;
     qryGulkwa.SQL.Add(sSelect + sWhere + sOrderBy);
     qryGulkwa.Open;
   end;

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select distinct H.cod_hm, H.desc_kor from hangmok H JOIN TF_Get_HangmokList(''' + qryGulkwa.FieldByName('cod_jisa').AsString + ''','''
                                                                 + qryGulkwa.FieldByName('num_id').AsString   + ''','''
                                                                 + qryGulkwa.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                 + ' ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
      sWhere :=   ' WHERE (h.gubn_part <> ''06'') and (h.gubn_part <> ''11'') and (h.gubn_part <> ''12'') and (h.gubn_part <> ''13'') and (h.gubn_part <> ''14'')';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      if qryBunju.RecordCount > 0 then
      begin
         For I := 1 to qryBunju.RecordCount do
         Begin
               UP_VarMemSet(Index+1);
               begin
               UV_vcod_hm[Index]  := qryBunju.FieldByName('cod_hm').AsString;
               UV_vdesc_hm[Index]  := qryBunju.FieldByName('desc_kor').AsString;
               Inc(Index);
               end;
            Next;
         End;
         // Table과 Disconnect
         Close;
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
   tsGrid1.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

end;

end.
