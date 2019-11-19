//==============================================================================
// 프로그램 설명 : Ferritin 결과 리스트 (여)
// 시스템        : 통합검진
// 수정일자      : 2014.06.16
// 수정자        : 곽윤설
// 수정내용      : Ferritin항목에 대해 여자 수검자의 결과를 나이와 함께 조회
// 참고사항      : [전문의-김소영]
//==============================================================================


unit LD4Q82;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q82 = class(TfrmSingle)
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
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);

    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vNum_samp, UV_vSBun, UV_vName, UV_vage, UV_vcod_Jangbi, UV_vJangbi_sokun : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q82: TfrmLD4Q82;


implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc ;

{$R *.DFM}

procedure TfrmLD4Q82.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo               := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp         := VarArrayCreate([0, 0], varOleStr);
      UV_vSBun             := VarArrayCreate([0, 0], varOleStr);
      UV_vName             := VarArrayCreate([0, 0], varOleStr);
      UV_vcod_Jangbi       := VarArrayCreate([0, 0], varOleStr);
      UV_vJangbi_sokun     := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vSBun,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vcod_Jangbi,   iLength);
      VarArrayRedim(UV_vJangbi_sokun,  iLength);
   end;
end;

function TfrmLD4Q82.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q82.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
      Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD4Q82.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vSBun[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vCod_jangbi[DataRow-1];
      6 : Value := UV_vJangbi_sokun[DataRow-1];
   end;
end;

procedure TfrmLD4Q82.btnQueryClick(Sender: TObject);
var index, I, rowCount, iRet, iTemp, iAge : Integer;
    sSelect, sWhere, sOrderBy : String;
    sJangbi_List, sHul_List, sMan : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;
      sSelect := 'Select  a.cod_jisa, a.num_id, a.dat_gmgn, a. num_jumin, a.num_samp, a.desc_name, a.cod_hul, a.cod_chuga, b.num_bunju, j.desc_sokun, j.desc_sbsg, j.cod_jangbi';
      sSelect := sSelect + ' From Gumgin a with(nolock) left outer join bunju b with(nolock) on a.cod_jisa = b.cod_jisa and a.num_id = b.num_id and a.dat_gmgn = b.dat_gmgn ';
      sSelect := sSelect + ' left outer join jangbi j with(nolock) on a.cod_jisa = j.cod_jisa and a.num_id = j.num_id and a.dat_gmgn = j.dat_gmgn ';

      sWhere := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';

       sWhere := sWhere + ' And b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju <= ''' + bunjuno2.Text + '''';
         sOrderBy := ' ORDER BY G.num_bunju ';
      end
      else
      begin
         if (Trim(MEdt_SampS.Text) <> '') and (Trim(MEdt_SampE.Text) = '') then
            sWhere := sWhere + ' AND A.num_samp = ''' + MEdt_SampS.Text + '''';

         if (Trim(MEdt_SampS.Text) <> '') and (Trim(MEdt_SampE.Text) <> '') then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';
         IF Cmb_gubn.Text ='샘플번호' then  sOrderBy := ' ORDER BY a.num_samp '
         else sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), b.num_bunju ' ;
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q41', 'Q', 'N');
         For I := 1 to RecordCount do
         begin
            UP_VarMemSet(index);

            // 항목(혈액프로파일)추출..
            sJangbi_List := ''; sHul_List := '';
            if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) = 'JJ' then
                     begin
                       sJangbi_List := sJangbi_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     end
                     else
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //추가검사항목..
            iRet := GF_MulToSingle(qryGulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
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

            if RecordCount >= 0 then
            begin
              if (qryGulkwa.FieldByName('cod_jangbi').AsString = 'JJ38') or (qryGulkwa.FieldByName('cod_jangbi').AsString = 'JJ27') or
                 (qryGulkwa.FieldByName('cod_jangbi').AsString = 'JJ41') or (qryGulkwa.FieldByName('cod_jangbi').AsString = 'JJA0') then
              begin
                if (qryGulkwa.FieldByName('desc_sokun').AsString <> '') or (qryGulkwa.FieldByName('desc_sbsg').AsString <> '') then
                begin
                   UV_vNo[index]           := intTostr(index+1);
                   UV_vNum_samp[index]     := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vSBun[index]         := qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vName[index]         := qryGulkwa.FieldByName('desc_name').AsString;
                   if qryGulkwa.FieldByName('cod_jangbi').AsString ='JJ38' then  UV_vcod_Jangbi[index]   := '갑상선'
                   else if qryGulkwa.FieldByName('cod_jangbi').AsString ='JJ27' then  UV_vcod_Jangbi[index]   := '전립선'
                   else if (qryGulkwa.FieldByName('cod_jangbi').AsString ='JJ41') or (qryGulkwa.FieldByName('cod_jangbi').AsString ='JJA0') then  UV_vcod_Jangbi[index]   := '질초음파';

                   UV_vJangbi_sokun[index] := qryGulkwa.FieldByName('desc_sokun').AsString + '/ ' + qryGulkwa.FieldByName('desc_sbsg').AsString;

                   Inc(index);
                   Gride.Progress := I;
                end;
            end;
            next;
            end;
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
   end; // qryGulkwa

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q82.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
   if Cmb_gubn.Text = '샘플번호' then
   begin
       bunjuno1.Enabled := False;
       bunjuno2.Enabled := False;
       bunjuno1.Text := '0001';
       bunjuno2.Text := '9999';
       MEdt_SampS.Enabled := True;
       MEdt_SampE.Enabled := True;
   end
   else if Cmb_gubn.Text = '분주번호' then
   begin
       bunjuno1.Enabled := True;
       bunjuno2.Enabled := True;
       MEdt_SampS.Enabled := False;
       MEdt_SampE.Enabled := False;
       MEdt_SampS.Text := '';
       MEdt_SampE.Text := '';
   end;
end;

procedure TfrmLD4Q82.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
    procedure TfrmLD4Q82.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.

