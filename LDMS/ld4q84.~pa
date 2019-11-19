//==============================================================================
// 프로그램 설명 : 세포조직 바코드 리딩 프로그램
// 시스템        : LDMS
// 수정일자      : 2019.03.18
// 수정자        : 박수지
// 수정내용      :
// 참고사항      : 바코드형식(년월일+샘플번호)
//==============================================================================
unit LD4Q84;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, ComObj,
  Gauges;

type
  TfrmLD4Q84 = class(TfrmSingle)
    pnlCond: TPanel;
    Panel4: TPanel;
    Panel2: TPanel;
    cmbGubn: TComboBox;
    Panel5: TPanel;
    Panel13: TPanel;
    barcode: TMaskEdit;
    Panel3: TPanel;
    edtBDate: TDateEdit;
    btnBdate: TSpeedButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel12: TPanel;
    Panel11: TPanel;
    Panel10: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    CK_Max: TCheckBox;
    Edt_Max: TEdit;
    grdMaster: TtsGrid;
    qryCa_desc_buwi_Max: TQuery;
    qryCell: TQuery;
    qryU_Cell: TQuery;
    qryBunju: TQuery;
    qryGumgin: TQuery;
    qry_Cell: TQuery;
    qrycell_B: TQuery;
    cmbJisa: TComboBox;
    procedure btnInsertClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnPInsertClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure barcodeExit(Sender: TObject);
    procedure btnBdateClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure CK_MaxClick(Sender: TObject);
    procedure edtBDateExit(Sender: TObject);
    procedure cmbGubnChange(Sender: TObject);
    private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vNo, UV_vCod_jisa, UV_vDat_gmgn, UV_vNum_samp, UV_vDesc_name,
    UV_vNum_jumin, UV_vSample, UV_vRegist, UV_vDesc_buwi: Variant;
    vCheck_B001 : Boolean;
    sJisa, sNum_id, sDat_gmgn : String;
    iCount : Integer;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
  public
    { Public declarations }
    // Display(저장 유무)...
    UV_sDisplay : String;
  end;

var
  frmLD4Q84: TfrmLD4Q84;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD4Q84.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vNo             := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa       := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn       := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp       := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin      := VarArrayCreate([0, 0], varOleStr);
      UV_vSample         := VarArrayCreate([0, 0], varOleStr);
      UV_vRegist         := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi      := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vNO,             iLength);
      VarArrayRedim(UV_vCod_jisa,       iLength);
      VarArrayRedim(UV_vDat_gmgn,       iLength);
      VarArrayRedim(UV_vNum_samp,       iLength);
      VarArrayRedim(UV_vDesc_name,      iLength);
      VarArrayRedim(UV_vNum_jumin,      iLength);
      VarArrayRedim(UV_vSample,         iLength);
      VarArrayRedim(UV_vRegist,         iLength);
      VarArrayRedim(UV_vDesc_buwi,      iLength);
   end;
end;

function TfrmLD4Q84.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (barcode.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;

end;

function TfrmLD4Q84.UF_Save : Boolean;
var
   i, iCol_sum, iMax, index : Integer;
   sNow, sBuwi, sCod_hm : string;
begin
   Result := False;
   UV_bEdit := True;
   // 작업중인지 Check
   if not UV_bEdit then exit;
   iMax := StrToInt(Edt_Max.Text);

   iCol_sum := 0;
   try
      begin
         with qryU_Cell do
         begin
            qryU_Cell.ParamByName('dat_last').AsString        := GV_sToday;
            qryU_Cell.ParamByName('dat_time').AsString        := formatdatetime('HH:NN:SS', now);
            qryU_Cell.ParamByName('cod_update').AsString      := GV_sUserId;
            qryU_Cell.ParamByName('dat_update').AsString      := GV_sToday;
            qryU_Cell.ParamByName('cod_jisa').AsString        := Copy(barcode.Text,7,2);
            qryU_Cell.ParamByName('num_id').AsString          := sNum_id;
            qryU_Cell.ParamByName('dat_gmgn').AsString        := sDat_gmgn;

            if vCheck_B001= true then
            begin
              if      cmbGubn.Text = 'B001' then
              begin
                   qryU_Cell.ParamByName('cod_hm').AsString := 'B005';
                   sCod_hm := 'B005';
              end
              else if cmbGubn.Text = 'B005' then
              begin
                    qryU_Cell.ParamByName('cod_hm').AsString := 'B001';
                    sCod_hm := 'B001';
              end
              else qryU_Cell.ParamByName('cod_hm').AsString   := cmbGubn.text;
              vCheck_B001 := false;
            end
            else qryU_Cell.ParamByName('cod_hm').AsString          := cmbGubn.text;

               if (Edt_Max.text <> '') and (cmbGubn.text = 'B009') then
               begin
                  qryU_Cell.ParamByName('desc_buwi').Asstring := 'L' + Copy(sDat_gmgn, 3, 2) + '-' +GF_LPad(IntToStr(iMax), 6, '0');
                  sBuwi := 'L' + Copy(sDat_gmgn, 3, 2) + '-' +GF_LPad(IntToStr(iMax), 6, '0');

               end
               else if (Edt_Max.text <> '') and ((cmbGubn.text = 'P003') or (cmbGubn.text = 'P012') or (cmbGubn.text = 'P013')) then
               begin
                  qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(sDat_gmgn, 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                  sBuwi := 'P' + Copy(sDat_gmgn, 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');

               end
               else if (Edt_Max.text <> '')  and ((cmbGubn.text = 'B001') or (cmbGubn.text = 'B005') or (cmbGubn.text = 'B007') or (cmbGubn.text = 'P009')) then
               begin

                  qryU_Cell.ParamByName('desc_buwi').Asstring := 'T' + Copy(sDat_gmgn, 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                  sBuwi := 'T' + Copy(sDat_gmgn, 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
               end
               else if (Edt_Max.text = '') then GF_MsgBox('Warning', '접수 no.시작 값 부터 부여해주세요.');
         end;
            qryU_Cell.Execsql;
            GP_query_log(qryU_Cell, 'LD4Q84', 'U', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;
   // Grid에 자료를 반영
   with grdMaster do
   begin

      UP_VarMemSet(Rows);
      i := Rows;
      index := 0;
      // 현재 Field의 자료를 Grid관련 Variant변수에 할당
      UV_vNo[i]                    := index + 1;
      UV_vCod_jisa [i]             := Copy(barcode.Text, 7, 2);
      UV_vDat_gmgn[i]              := Copy(GV_sToday,1,2) + Copy(barcode.Text,1,6);
      UV_vNum_samp[i]              := copy(barcode.Text,7,6);
      qryGumgin.Active := False;

      qryGumgin.ParamByName('cod_Jisa').AsString := Copy(barcode.Text, 7, 2);
      qryGumgin.ParamByName('dat_gmgn').AsString := Copy(GV_sToday,1,2) + Copy(barcode.Text,1,6);
      qryGumgin.ParamByName('num_samp').AsString := copy(barcode.Text,7,6);
      qryGumgin.Active := True;
      UV_vDesc_name[i]             := qryGumgin.fieldbyName('desc_name').asString;
      UV_vNum_jumin [i]            := copy(qryGumgin.fieldbyName('num_Jumin').asString,1,6);
      if sCod_hm <> '' then UV_vSample[i]  := sCod_hm
      else UV_vSample[i]           := cmbGubn.Text;
      UV_vRegist[i]                := GV_sToday + ' '+ formatdatetime('HH:NN:SS', now);
      UV_vDesc_buwi[i]             := sBuwi;
      sCod_hm := '';
      
      Rows := Rows + 1;
      CurrentDataRow := Rows;
      TopRow := Rows;

      // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;

      // 자료 Display
      grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLD4Q84.btnInsertClick(Sender: TObject);
begin
   inherited;

   // Insert Mode & Msg
   UP_SetMode('INSERT');
end;

procedure TfrmLD4Q84.btnSaveClick(Sender: TObject);
var bCheck : boolean;
begin
   inherited;

   UF_Save;
end;

procedure TfrmLD4Q84.FormCreate(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy : String;
    sJisa : string;
begin
   inherited;

   // 환경설정
   UP_VarMemSet(0);
   // Grid 초기화
   grdMaster.Rows := 0;

   GP_ComboJisa(cmbJisa);
   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      sJisa := '15';
   end
   else
   begin
      sJisa := GV_sJisa;
   end;
   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);
   edtBDate.Text := GV_sToday;

   //각 센터 분주 현황 체크 창 생성 (당일 자로 분주확인)
   index := 0;
   with qryBunju do
   begin
     qryBunju.Active := False;
     qryBunju.ParamByName('dat_bunju').AsString := GV_sToday;
     qryBunju.Active := True;

       if qryBunju.RecordCount > 0 then
       begin
          UP_VarMemSet(qryBunju.RecordCount-1);
          while Not Eof do
          begin
             if qryBunju.FieldByName('cod_jisa').AsString = '15' then  Panel8.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '12' then Panel9.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '11' then Panel10.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '43' then Panel12.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '51' then Panel11.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '71' then Panel14.Color := $001ACEF4
             else if qryBunju.FieldByName('cod_jisa').AsString = '61' then Panel15.Color := $001ACEF4;

             Next;
             Inc(index);
          end;
       end
       else
       begin
          GF_MsgBox('Information', 'NODATA');
       end;
       qryBunju.Active := False;
       qryBunju.Close;
   end;
end;

procedure TfrmLD4Q84.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
var
   index : Integer;
   sBalkubNo : String;
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vDat_gmgn[DataRow-1];
      4 : Value := UV_vNum_samp[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : Value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vSample[DataRow-1];
      8 : Value := UV_vRegist[DataRow-1];
      9 : Value := UV_vDesc_buwi[DataRow-1];

   end;
end;

procedure TfrmLD4Q84.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
   inherited;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

   // Query 수행 & 배열에 저장
   index := 0;
   with qry_Cell do
   begin
     qry_Cell.Active := False;
     sSelect := 'SELECT B.dat_bunju, G.num_samp, G.desc_name, G.num_jumin, C.* from cell C ';
     sSelect := sSelect + 'join bunju B on C.Cod_jisa = B.Cod_jisa and C.num_id = B.num_id and C.dat_gmgn = B.dat_gmgn ';
     sSelect := sSelect + 'Join gumgin G on C.cod_jisa = G.cod_jisa and C.dat_gmgn = G.dat_gmgn and C.num_id = G.num_id ' ;
     sWhere := 'WHERE ';
      if Copy(cmbJisa.Text, 1, 2) <> '00' then
              sWhere := sWhere + ' G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';
     sWhere := sWhere + ' B.dat_bunju = ''' + edtBDate.text + ''' AND ';
     sWhere := sWhere + ' C.cod_hm = ''' + cmbGubn.text + ''' AND ';
     sWhere := sWhere + ' C.desc_buwi <> '''' ';
     sOrderBy := 'order by C.desc_buwi';
     SQL.Clear;
     SQL.Add(sSelect + sWhere + sOrderBy);
     qry_Cell.Active := True;
     
     qry_Cell.Open;
     qry_Cell.Last;
     qry_Cell.First;
     qry_cell.sql.text;

     GP_query_log(qry_Cell, 'LD4Q84', 'Q', 'Y');
     if qry_Cell.RecordCount > 0 then
     begin
        UP_VarMemSet(qry_Cell.RecordCount-1);

        while Not Eof do
        begin
           UP_VarMemSet(Index+1);
           UV_vNo[index]               := index + 1;
           UV_vCod_jisa[index]         := qry_Cell.FieldByName('cod_jisa').AsString;
           UV_vNum_samp[index]         := qry_Cell.FieldByName('num_samp').AsString;
           UV_vDesc_name[index]        := qry_Cell.FieldByName('desc_name').AsString;
           UV_vNum_jumin[index]        := copy(qry_Cell.FieldByName('num_jumin').AsString, 1, 6) + '-' + copy(qry_Cell.FieldByName('num_jumin').AsString, 7, 1);
           UV_vSample[index]           := qry_Cell.FieldByName('cod_hm').AsString;
           UV_vRegist[index]           := qry_Cell.FieldByName('dat_last').AsString + ' ' + qry_Cell.FieldByName('dat_time').AsString;
           UV_vDat_gmgn[index]         := qry_Cell.FieldByName('dat_gmgn').AsString;
           UV_vDesc_buwi[index]        := qry_Cell.FieldByName('desc_buwi').AsString;
           Next;
           Inc(index);
        end;
     end
     else
     begin
        GF_MsgBox('Information', 'NODATA');
     end;

      // Table과 Disconnect
      qry_Cell.Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   iCount:= 0;
end;
{
procedure TfrmCK2I19.btnDeleteClick(Sender: TObject);
begin
   inherited;

   // Delete Confirm Message
   if (grdMaster.Rows = 0) or (grdMaster.VisibleRows < 1) then exit;
   if not GF_MsgBox('Confirmation', 'DELETE') then exit;

   // Delete 작업수행
   try
      with qryD_Meal_collect do
      begin
         ParamByName('cod_jisa').AsString := copy(cmbCOD_JISA.Text,1,2);
         ParamByName('year').AsString     := edtNUM_YEAR.Text;

         ExecSql;
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end;

   // 화면에서 안보이게함
   grdMaster.RowVisible[grdMaster.CurrentDataRow] := False;
   if grdMaster.VisibleRows < 1 then GP_FieldClear(pnlDetail);

   // Delete Mode & Msg
   UP_SetMode('DELETE');
end;
}
procedure TfrmLD4Q84.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;


   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD4Q84.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;
   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD4Q84.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var
  sBalkubNo : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;

procedure TfrmLD4Q84.btnPInsertClick(Sender: TObject);
begin
   inherited;

   // 작업절차
   // 1. 사용자가 작업중이라면 저장이 입력상태
   // 2. 사용자가 작업중이아니라면 바로 입력상태
   if UV_bEdit then
   begin
      if UF_Save then
         btnInsert.Click;
   end
   else
      btnInsert.Click;
end;

procedure TfrmLD4Q84.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  
begin
   inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q84.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
end;

procedure TfrmLD4Q84.barcodeExit(Sender: TObject);
begin
   inherited;

   with qryGumgin do
   begin
   qryGumgin.Active := False;
   qryGumgin.ParamByName('cod_Jisa').AsString := Copy(barcode.Text, 7, 2);
   qryGumgin.ParamByName('dat_gmgn').AsString := Copy(GV_sToday,1,2) + Copy(barcode.Text,1,6);
   qryGumgin.ParamByName('num_samp').AsString := copy(barcode.Text,7,7);
   qryGumgin.Active := True;

   sJisa     := qryGumgin.fieldbyName('cod_jisa').asString;
   sNum_id   := qryGumgin.fieldbyName('num_id').asString;
   sDat_gmgn := qryGumgin.fieldbyName('dat_gmgn').asString;
   qryGumgin.Close;

  if length(barcode.text) = 12 then
  begin
     with qryCell do
     begin
     qryCell.Active := False;
     qryCell.ParamByName('cod_jisa').AsString := sJisa;
     qryCell.ParamByName('num_id').AsString   := sNum_id;
     qryCell.ParamByName('dat_gmgn').AsString := sDat_gmgn;
     qryCell.ParamByName('cod_hm').AsString := cmbGubn.Text;
     qryCell.Active := True;
     qryCell.Open;
     qryCell.Last;
     qryCell.First;

     if qryCell.RecordCount > 0 then
     begin
       if (qryCell.fieldbyname('desc_buwi').asString <> '' )  then
       begin
          if (cmbGubn.text <> 'B001') and (cmbGubn.text <> 'B005') then
          begin
               if MessageDlg('이미 접수되어있습니다. (NO.'+copy(barcode.Text,7,6)+')접수를 진행하시겠습니까? 기존 접수no.는 삭제됩니다.', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit
               else
               begin
               //수정상태...(brow mode)
               UV_iStatus:= 2;
               UF_Save;
               CK_MaxClick(sender);
               end;
          end
          else GF_MsgBox('Warning', '접수되어 있는 '+ cmbGubn.text + '코드가 있습니다. 접수 번호 삭제 후 다시 시도해주세요.');
       end
       else
       begin
          if (cmbGubn.text = 'B001') OR (cmbGubn.text = 'B005') then
          begin
               if (cmbGubn.text = 'B001') then
               begin
                    vCheck_B001 := False;
                    with qryCell_B do
                    begin
                         qryCell_B.Active := False;
                         qryCell_B.ParamByName('cod_jisa').AsString := sJisa;
                         qryCell_B.ParamByName('num_id').AsString   := sNum_id;
                         qryCell_B.ParamByName('dat_gmgn').AsString := sDat_gmgn;
                         qryCell_B.ParamByName('cod_hm').AsString := 'B005';
                         qryCell_B.Active := True;
                         qryCell_B.Open;
                         qryCell_B.Last;
                         qryCell_B.First;
                         IF qryCell_B.RecordCount > 0 then
                         begin
                              UV_iStatus:= 2;
                              if MessageDlg('B005코드가 있습니다. 두 항목 모두 동일한 접수 번호로 저장합니다. ', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit
                              else
                              begin
                                   vCheck_B001 := True;
                                   UF_Save;
                              end;
                         end;
                    end;
               end
               else if (cmbGubn.text = 'B005') then
               begin
                    vCheck_B001 := False;
                    with qryCell_B do
                    begin
                         qryCell_B.Active := False;
                         qryCell_B.ParamByName('cod_jisa').AsString := sJisa;
                         qryCell_B.ParamByName('num_id').AsString   := sNum_id;
                         qryCell_B.ParamByName('dat_gmgn').AsString := sDat_gmgn;
                         qryCell_B.ParamByName('cod_hm').AsString := 'B001';
                         qryCell_B.Active := True;
                         qryCell_B.Open;
                         qryCell_B.Last;
                         qryCell_B.First;

                         IF qryCell_B.RecordCount > 0 then
                         begin
                              UV_iStatus:= 2;
                              if MessageDlg('B001코드가 있습니다. 두 항목 모두 동일한 접수 번호로 저장합니다. ', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit
                              else
                              begin
                                   vCheck_B001 := True;
                                   UF_Save;
                              end;
                         end;
                    end;
               end;

          end;
         //수정상태...(brow mode)
          UV_iStatus:= 2;
          UF_Save;
          barcode.text := '';
          barcode.setfocus;
          CK_MaxClick(sender);
       end;
     end
     else
     begin
           GF_MsgBox('Warning', '검진자정보가 없습니다.');
           barcode.text := '';
           barcode.setfocus;
           dmComm.database.Rollback;
           Exit;
     end;
     end;
     qryCell.Active := False;
  end
  else
  begin
     if length(barcode.text) > 0 then
     begin
        GF_MsgBox('Warning', '바코드 길이(20Byte)가 맞지 않습니다.');
        barcode.setfocus;
        barcode.text := '';
        dmComm.database.Rollback;
     end;
  end;
  end;

  qryGumgin.Active := False;
end;
{
procedure TfrmCK3Q06.cmbRelationChange(Sender: TObject);
begin
  inherited;
  if cmbRelation.ItemIndex = 0 then
  begin
        frmCK3Q061  := TFrmCK3Q061.Create(self);
        GP_ComboDisplay(frmCK3Q061.cmbJisa, copy(GV_sUserId,1,2), 2);
  end
  else if cmbRelation.ItemIndex = 1 then
  begin
        frmCK3Q062  := TFrmCK3Q062.Create(self);
        GP_Combojisa(frmCK3Q062.cmbJisa);
  end;
end;
}
procedure TfrmLD4Q84.btnBdateClick(Sender: TObject);
begin
  inherited;
  if Sender = btnBdate then GF_CalendarClick(edtBDate);
  CK_Max.checked := False;
  Edt_Max.Enabled := False;
  Edt_Max.Text    := '';
  Edt_Max.Color   := $00E0E0E0;
end;

procedure TfrmLD4Q84.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

procedure TfrmLD4Q84.CK_MaxClick(Sender: TObject);
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
   if CK_Max.Checked = True then
   begin
      Edt_Max.Enabled := False;
      Edt_Max.Color   := clWindow;

     //'L'로 시작하는 건
     if (cmbGubn.Text = 'B009') then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';
           sSelect := sSelect + ' WHERE desc_buwi > ''L'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''000000''';
           sSelect := sSelect + ' AND   desc_buwi < ''L'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''999999''';
           sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''L'' + ''' + copy(edtBDate.text,3,2) + '''';
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'T'로 시작하는건
     else if (cmbGubn.Text = 'B001') or
             (cmbGubn.Text = 'B005') or
             (cmbGubn.Text = 'B007') or
             (cmbGubn.Text = 'P009') then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';
           sSelect := sSelect + ' WHERE desc_buwi > ''T'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''000000''';
           sSelect := sSelect + ' AND   desc_buwi < ''T'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''999999''';
           sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''T'' + ''' + copy(edtBDate.text,3,2) + '''';
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'P'로 시작하는건
     else if (cmbGubn.Text = 'P003') or (cmbGubn.Text = 'P012') or (cmbGubn.Text = 'P013') then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';

           sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''000000''';
           sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + copy(edtBDate.text,3,2) + ''' + ''-'' + ''999999''';
           sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + copy(edtBDate.text,3,2) + '''';

           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end;

     Edt_Max.Text := intToStr(qryCa_desc_buwi_Max.FieldByName('RESULT').Asinteger + 1);
     qryCa_desc_buwi_Max.Active := False;
   end;
end;

procedure TfrmLD4Q84.edtBDateExit(Sender: TObject);
begin
  inherited;
  CK_Max.checked := false;
  Edt_Max.Enabled := False;
  Edt_Max.Text    := '';
  Edt_Max.Color   := $00E0E0E0;
end;

procedure TfrmLD4Q84.cmbGubnChange(Sender: TObject);
begin
  inherited;
  CK_Max.checked := false;
  Edt_Max.Enabled := False;
  Edt_Max.Text    := '';
  Edt_Max.Color   := $00E0E0E0;
end;

end.
