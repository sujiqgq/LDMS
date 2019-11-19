//==============================================================================
// 프로그램 설명 : 조직외주관리
// 시스템        : 분석관리
// 수정일자      : 2009.06.04
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q22;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit ;

type
  TfrmLD4Q22 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    cmbJisa: TComboBox;
    mskFrom: TDateEdit;
    rbDate: TRadioButton;
    grdMaster: TtsGrid;
    btnsDate: TSpeedButton;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    Label4: TLabel;
    ComB_ShEndo: TComboBox;
    qryU_Cell: TQuery;
    Chk_ALL: TCheckBox;
    Chk_order: TCheckBox;
    qryCellChk: TQuery;
    Label2: TLabel;
    CmbOrder: TComboBox;

    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure UP_Click(Sender: TObject);
    procedure up_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Chk_ALLClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vGubn_check, UV_vDat_gmgn, UV_vDesc_name, UV_vNum_jumin, UV_vAge, UV_vSex,
    UV_vChart_no, UV_vEmp1, UV_vEmp2, UV_vDat_gmgn2, UV_vCod_jisa, UV_vNum_id, UV_vCod_cell, UV_vCod_hm : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q22: TfrmLD4Q22;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD4Q22.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vGubn_check := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vAge        := VarArrayCreate([0, 0], varOleStr);
      UV_vSex        := VarArrayCreate([0, 0], varOleStr);
      UV_vChart_no   := VarArrayCreate([0, 0], varOleStr);
      UV_vEmp1       := VarArrayCreate([0, 0], varOleStr);
      UV_vEmp2       := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn2  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vGubn_check, iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vAge,        iLength);
      VarArrayRedim(UV_vSex,        iLength);
      VarArrayRedim(UV_vChart_no,   iLength);
      VarArrayRedim(UV_vEmp1,       iLength);
      VarArrayRedim(UV_vEmp2,       iLength);
      VarArrayRedim(UV_vDat_gmgn2,  iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vCod_cell,   iLength);
      VarArrayRedim(UV_vCod_hm,     iLength);
   end;
end;


function TfrmLD4Q22.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if mskFROM.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;



procedure TfrmLD4Q22.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   // 환경설정
   UP_VarMemSet(0);

   // 일자(From, To)를 오늘일자로 설정
   mskFrom.Text := GV_sToday;

   ComB_ShEndo.ItemIndex := 0;
end;


procedure TfrmLD4Q22.btnQueryClick(Sender: TObject);
var index, iAge : Integer;
    sSelect, sWhere, sOrderBy, sMan : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sWhere  := ''; sSelect := ''; sOrderBy := '';
   with qryCell do
   begin
      Active := False;

      // SQL문 생성
      sSelect := ' SELECT C.cod_hm, C.cod_cell, C.cod_jisa, C.num_id, C.dat_gmgn, C.cod_cell, gubn_order, G.desc_name, G.num_jumin' +
                 ' FROM cell C INNER JOIN gumgin G ' +
                 ' ON C.cod_jisa = G.cod_jisa and C.num_id = G.num_id and C.dat_gmgn = G.dat_gmgn ';
      if pnlJisa.Visible then
      begin
         if Copy(cmbJisa.Text, 1, 2) <> '00' then
            sWhere := ' WHERE C.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
      end
      else
         sWhere := ' WHERE C.cod_jisa = ''' + copy(GV_sUserId,1,2) + '''';

      sWhere := sWhere +  ' AND C.dat_gmgn = ''' + mskFROM.Text + '''';

      if ComB_ShEndo.ItemIndex = 0 then sWhere := sWhere + ' and SUBSTRING(C.cod_cell,1,1) = ''J'''
      else                              sWhere := sWhere + ' and SUBSTRING(C.cod_cell,1,1) = ''P''';

      if Chk_order.Checked then sWhere := sWhere + ' and C.gubn_order = ''Y''';

      if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' order by G.cod_jisa, desc_name'
      else if CmbOrder.ItemIndex = 2 then
         sOrderBy := ' order by G.num_jumin'
      else
         sOrderBy := ' order by G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryCell, 'LD4Q22', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);

         // 조건에 맞는 데이터 변수에 입력
         while not Eof do
         begin
            iAge := 0; sMan := '';
            GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

            if FieldByName('gubn_order').AsString = 'Y' then UV_vGubn_check[index]   := '1'
            else                                             UV_vGubn_check[index]   := '0';

            UV_vDat_gmgn[index]  := FieldByName('dat_gmgn').AsInteger;
            UV_vDesc_name[index] := FieldByName('desc_name').AsString;
            UV_vNum_jumin[index] := FieldByName('num_jumin').AsString;
            UV_vCod_jisa[index]  := FieldByName('cod_jisa').AsString;
            UV_vNum_id[index]    := FieldByName('Num_id').AsString;
            UV_vCod_cell[index]  := FieldByName('Cod_cell').AsString;
            UV_vCod_hm[index]    := FieldByName('Cod_hm').AsString;

            UV_vAge[index]       := iAge;
            UV_vSex[index]       := sMan;

            UV_vChart_no[index]  := '';
            UV_vEmp1[index]      := '';
            UV_vEmp2[index]      := '';
            UV_vDat_gmgn2[index] := FieldByName('dat_gmgn').AsString;

            Next;
            Inc(index);
         end;
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
      // Table과 Disconnect
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD4Q22.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD4Q22.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := StrToint(UV_vGubn_check[DataRow-1]);
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vCod_hm[DataRow-1];
      4 : Value := UV_vDat_gmgn[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vAge[DataRow-1];
      8 : Value := UV_vSex[DataRow-1];
      9 : Value := UV_vChart_no[DataRow-1];
     10 : Value := UV_vEmp1[DataRow-1];
     11 : Value := UV_vEmp2[DataRow-1];
     12 : Value := UV_vDat_gmgn2[DataRow-1];
   end;
end;


procedure TfrmLD4Q22.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];

end;


procedure TfrmLD4Q22.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnsDate then
      GF_CalendarClick(mskFrom);
end;


procedure TfrmLD4Q22.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskFrom        then UP_Click(btnsDate);
end;

procedure TfrmLD4Q22.Chk_ALLClick(Sender: TObject);
var iTemp,I : Integer;
begin
   inherited;
   iTemp := grdMaster.Rows;
   if Chk_ALL.Checked then
      for I := 0 to grdMaster.Rows - 1 do UV_vGubn_check[I] := '1'
   else
      for I := 0 to grdMaster.Rows - 1 do UV_vGubn_check[I] := '0';

   grdMaster.RowS := 0;
   grdMaster.RowS := iTemp;
end;

procedure TfrmLD4Q22.btnSaveClick(Sender: TObject);
var i : Integer;
    bCheck : Boolean;
    sTemp : string;
begin
   inherited;
   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;


   for I := 0 to grdMaster.rows -1 do
   begin
      bCheck := True;

      qryCellChk.Active := False;
      qryCellChk.ParamByName('cod_jisa').AsString := UV_vCod_jisa[I];
      qryCellChk.ParamByName('num_id').AsString   := UV_vNum_id[I];
      qryCellChk.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[I];
      qryCellChk.ParamByName('cod_cell').AsString := UV_vCod_cell[I];
      qryCellChk.Active := True;

      if qryCellChk.RecordCount > 0 then
      begin
          if Trim(qryCellChk.FieldByName('gubn_order').AsString) = 'Y' then sTemp :='1'
          else                                                              sTemp :='0';

          if UV_vGubn_check[I] = sTemp then bCheck := False;
      end
      else bCheck := False;

      //값이 변화 있으면 저장..
      if bCheck then
      begin
         // Check 확인
         if UV_vGubn_check[I]   = '1' then
         begin
            with qryU_Cell do
            begin
              ParamByName('gubn_order').AsString := 'Y';
              ParamByName('cod_jisa').AsString   := UV_vCod_jisa[I];
              ParamByName('num_id').AsString     := UV_vNum_id[I];
              ParamByName('dat_gmgn').AsString   := UV_vDat_gmgn[I];
              ParamByName('cod_cell').AsString   := UV_vCod_cell[I];
              ParamByName('dat_update').AsString := GV_sToday;
              ParamByName('cod_update').AsString := GV_sUserId;

              ExecSql;
            end;
         end
         else if UV_vGubn_check[I] = '0' then
         begin
            with qryU_Cell do
            begin
              ParamByName('gubn_order').AsString := '';
              ParamByName('cod_jisa').AsString   := UV_vCod_jisa[I];
              ParamByName('num_id').AsString     := UV_vNum_id[I];
              ParamByName('dat_gmgn').Asinteger  := UV_vDat_gmgn[I];
              ParamByName('cod_cell').AsString   := UV_vCod_cell[I];              
              ParamByName('dat_update').AsString := GV_sToday;
              ParamByName('cod_update').AsString := GV_sUserId;

              ExecSql;
            end;
         end;
      end;
   end;
   ShowMessage ('저장되었습니다.');
end;

end.


