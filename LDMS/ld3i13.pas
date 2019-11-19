//==============================================================================
// 프로그램 설명 : 혈액소견 코드변경
// 시스템        : 통합검진
// 수정일자      : 2016.08.26
// 수정자        : 박수지
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD3I13;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD3I13 = class(TfrmSingle)
    pnlCond: TPanel;
    Label7: TLabel;
    qryBunju: TQuery;
    Gride: TGauge;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    qryU_Sokun: TQuery;
    qryI_Sokun: TQuery;
    qrySokun: TQuery;
    qryBunju2: TQuery;
    qrySokun2: TQuery;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    qryBunju3: TQuery;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label26: TLabel;
    grdMaster: TtsGrid;
    tsGrid1: TtsGrid;
    cmbJisa: TComboBox;
    mskDate: TDateEdit;
    edt_sokun: TEdit;
    edt_sokun2: TEdit;
    edt_sokun3: TEdit;
    btn_Start: TBitBtn;
    Btn_Insert: TBitBtn;
    MEMO1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure tsGrid1CellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure Btn_StartClick(Sender: TObject);
    procedure Btn_InsertClick(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vGulkwa, UV_vNum_Bunju, UV_vcod_bjjs, UV_vIndex, UV_vCod_sokun, UV_vNum_id, UV_vDat_gmgn, UV_vCod_jisa  : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_VarMemSet2(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD3I13: TfrmLD3I13;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD6Q021;

{$R *.DFM}

procedure TfrmLD3I13.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD3I13.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD3I13.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo          := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vIndex       := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sokun   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vNum_Bunju,  iLength);
      VarArrayRedim(UV_vCod_bjjs,  iLength);
      VarArrayRedim(UV_vIndex,  iLength);
      VarArrayRedim(UV_vCod_sokun,  iLength);
      VarArrayRedim(UV_vNum_id,  iLength);
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vCod_jisa,  iLength);
   end;
end;
procedure TfrmLD3I13.UP_VarMemSet2(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo          := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vGulkwa      := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vNum_Bunju,  iLength);
      VarArrayRedim(UV_vGulkwa,  iLength);
   end;
end;


function TfrmLD3I13.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD3I13.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   UP_VarMemSet2(0);
   memo1.Clear;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;
                                              
procedure TfrmLD3I13.tsGrid1CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_id[DataRow-1];
      3 : Value := UV_vNum_Bunju[DataRow-1];
      4 : Value := UV_vCod_sokun[DataRow-1];
      5 : Value := UV_vCod_bjjs[DataRow-1];
      6 : Value := UV_vIndex[DataRow-1];
      7 : Value := UV_vDat_gmgn[DataRow-1];
      8 : Value := UV_vCod_jisa[DataRow-1];
   end;
end;

procedure TfrmLD3I13.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;
procedure TfrmLD3I13.btn_StartClick(Sender: TObject);
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
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      if Memo1.Text <> '' then
      begin
           sSelect := ' select A.cod_jisa, A.num_id, A.dat_gmgn, A.num_samp, A.desc_name,A.num_jumin,H.cod_sokun,H.num_bunju, H.num_seq, H.cod_bjjs, H.cod_sokun';
           sSelect := sSelect + ' From hul_sokun H   left outer join Gumgin A ';
           sSelect := sSelect + ' On H.cod_jisa = A.cod_jisa and a.num_id = H.num_id and a.dat_gmgn = h.dat_gmgn ';
           sWhere := ' WHERE H.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
           sWhere := sWhere + ' AND  H.Dat_Bunju = ''' + mskDate.Text + ''' ';
           sWhere := sWhere + ' AND  H.cod_sokun like  ''' + edt_sokun.Text +'%'+ '''';
           sOrderBy := ' ORDER BY CAST(H.num_bunju AS INT)';
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD3I13', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
               UP_VarMemSet(Index+1);
               IF GF_LPad(memo1.lines.strings[k],4,'0') = qryBunju.FieldByName('num_bunju').AsString Then
               begin
               UV_vNo[Index]  := IntToStr(Index+1);
               UV_vNum_id[Index]  := qryBunju.FieldByName('num_id').AsString;
               UV_vNum_Bunju[Index]  := qryBunju.FieldByName('num_bunju').AsString;
               UV_vCod_bjjs[Index] := qryBunju.FieldByName('cod_bjjs').AsString;
               UV_vIndex[Index] := qryBunju.FieldByName('num_seq').AsString;
               UV_vCod_sokun[Index] := qryBunju.FieldByName('cod_sokun').AsString;
               UV_vDat_gmgn[Index] := qryBunju.FieldByName('dat_gmgn').AsString;
               UV_vCod_jisa[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               Inc(k);
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

procedure TfrmLD3I13.Btn_InsertClick(Sender: TObject);
var index, i, iRet, iTemp, num , j, k, icur : Integer;
    Chk_qryU_Sokun : Boolean;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
begin
  inherited;
  sSelect := ''; sWhere := ''; sOrderBy := '';
  Index := 0;
  // 조회조건 Check
   if not UF_Condition then exit;
   // 현재위치 추출
   iCur := tsGrid1.Rows;
   // Grid 초기화
   grdMaster.Rows := 0;
   chk_qryU_Sokun := False;


   if not GF_MsgBox('Confirmation', 'Transfer를 시작 하시겠습니까.?') then exit;

   With qrySokun Do
     Begin
      Active := False;
      ParamByName('cod_sokun').AsString  := edt_sokun2.Text;
      Active := True;
     END;

   With qrySokun2 Do
     Begin
      Active := False;
      ParamByName('cod_sokun').AsString  := edt_sokun3.Text;
      Active := True;
     END;

   If edt_sokun2.text <> '' then
   Begin
   for i := 1 to tsGrid1.Rows do
      begin
       With qryU_Sokun Do
         if qrySokun.fieldByName('COD_SOKUN').AsString <> '' then
         Begin
         ParamByName('cod_sokun').AsString   := qrySokun.fieldByName('COD_SOKUN').AsString;
         ParamByName('desc_sokun').AsString  := qrySokun.fieldByName('desc_sbsg').AsString;
         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         ParamByName('gubn_pan').AsString    := qrySokun.fieldByName('gubn_pnjg').AsString;
         ParamByName('Cod_bjjs').AsString    := Copy(cmbJisa.Text,1,2);
         ParamByName('dat_bunju').AsString   := mskDate.Text;
         ParamByName('Num_bunju').AsString   := (tsGrid1.Cell[3,i]);
         ParamByName('gubn_hsok').AsString   := '1';
         ParamByName('num_seq').AsString     := (tsGrid1.Cell[6,i]);
         Execsql;
         GP_query_log(qryU_Sokun, 'LD1I13', 'U', 'Y');
         end // end of for;
         else
         begin
          GF_MsgBox('Warning', '변경할 소견코드를 잘못 입력했습니다.');
          exit;
         end;
      end;
      chk_qryU_Sokun := True;
   end
   Else if  edt_sokun3.text = '' then
      begin
         GF_MsgBox('Warning', '변경할 소견 코드를 입력해주세요!');
         exit;
      end;
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   for i := 1 to iCur do
    begin
     with qryBunju3 do
     begin
      // SQL문 생성
      Close;
      begin
           sSelect := 'select COUNT(num_id) as cnt, H.num_id, H.cod_bjjs, H.dat_bunju, H.dat_gmgn, H.cod_jisa, H.num_bunju ';
           sSelect := sSelect + ' from hul_sokun H';
           sWhere := ' WHERE Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
           sWhere := sWhere + ' AND  Dat_Bunju = ''' + mskDate.Text + ''' ';
           sWhere := sWhere + ' AND  gubn_hsok = ''1''';
           sWHERE := sWhere + ' AND  num_id = ''' + tsGrid1.Cell[2, i]+ '''';
           sGroupBy := ' GROUP BY num_id, cod_bjjs, dat_bunju, dat_gmgn, cod_jisa, num_bunju' ;
           sOrderBy := ' ORDER BY CAST(num_bunju AS INT)';
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;
      Gride.MaxValue := iCur;

      If (edt_sokun3.text <> '') then
      begin
       With qryI_Sokun Do
         Begin
         ParamByName('cod_bjjs').AsString    := Copy(cmbJisa.Text,1,2);
         ParamByName('num_id').AsString      := (tsGrid1.Cell[2,i]);
         ParamByName('dat_bunju').AsString   := mskDate.Text;
         ParamByName('num_bunju').AsString   := (tsGrid1.Cell[3,i]);
         ParamByName('cod_sokun').AsString   := qrySokun2.fieldByName('COD_SOKUN').AsString;
         ParamByName('cod_jisa').AsString    := (tsGrid1.Cell[8,i]);
         ParamByName('cod_doct').AsString    := '150005';
         ParamByName('dat_gmgn').AsString    := (tsGrid1.Cell[7,i]);
         ParamByName('desc_sokun').AsString  := qrySokun2.fieldByName('desc_sbsg').AsString;
         ParamByName('gubn_hsok').AsString   := '1';
         ParamByName('num_seq').AsInteger    := qryBunju3.fieldByName('cnt').asInteger + 1 ;
         ParamByName('DAT_INSERT').AsString  := GV_sToday;
         ParamByName('COD_INSERT').AsString  := GV_sUserId;
         ParamByName('GUBN_PAN').AsString    := qrySokun2.fieldByName('gubn_pnjg').AsString;
         Execsql;
         GP_query_log(qryI_Sokun, 'LD1I13', 'I', 'Y');
         end;
      end;
      end;
      chk_qryU_Sokun := True;
   end;

   //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   if chk_qryU_Sokun = True then
   Begin
   for i := 1 to iCur do
   with qryBunju2 do
   begin
      // SQL문 생성
      Close;
          begin
            sSelect := ' select COUNT(A.num_id) as cnt, A.dat_gmgn, H.num_bunju, H.cod_bjjs ';
            sSelect := sSelect + ' From hul_sokun H   left outer join Gumgin A ';
            sSelect := sSelect + ' On H.cod_jisa = A.cod_jisa and a.num_id = H.num_id and a.dat_gmgn = h.dat_gmgn ';
            sWhere := ' WHERE H.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
            sWhere := sWhere + ' AND  H.Dat_Bunju = ''' + mskDate.Text + ''' ';
            if (edt_sokun2.text <> '') and (edt_sokun3.text <> '') then
            sWhere := sWhere + ' AND  (H.cod_sokun like  ''' + edt_sokun2.Text +'%'+ ''' or H.cod_sokun like  ''' + edt_sokun3.Text +'%'+ ''') '
            else if  (edt_sokun2.text = '') and (edt_sokun3.text <> '') then
            sWhere := sWhere + ' AND  H.cod_sokun like  ''' + edt_sokun3.Text +'%'+ ''''
            else
            sWhere := sWhere + ' AND  H.cod_sokun like  ''' + edt_sokun2.Text +'%'+ '''';
            sWHERE := sWhere + ' AND  A.num_id = ''' + tsGrid1.Cell[2, i]+ '''';
            sGroupBy := ' GROUP BY a.num_id, cod_bjjs, dat_bunju, a.dat_gmgn, a.Cod_jisa, num_bunju ';
            sOrderBy := ' ORDER BY CAST(H.num_bunju AS INT)';
          end;
           SQL.Clear;
           SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
           Open;
           Gride.MaxValue := icur;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD3I13', 'Q', 'N');
         Begin
            Gride.Progress := icur;
               UP_VarMemSet2(Index);
               begin
               UV_vNo[Index]  := IntToStr(Index+1);
               UV_vNum_Bunju[Index]  := qryBunju2.FieldByName('num_bunju').AsString;
               UV_vGulkwa[Index]  := qryBunju2.FieldByName('cnt').AsString;
               Inc(Index);
               end ;
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
   grdMaster.Rows := index;
   //Active := False;
End;

end;

procedure TfrmLD3I13.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
     // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_Bunju[DataRow-1];
      3 : Value := UV_vGulkwa[DataRow-1];
   end;
end;

end.
 