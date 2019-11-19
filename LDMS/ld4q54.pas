//==============================================================================
// 프로그램 설명 : 간염 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2004.05.13
// 수정자        : 최지혜
// 수정내용      : 
// 참고사항      :
//==============================================================================
// 프로그램 설명 : C형 간염 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2006.04.20
// 수정자        : 김승철
//==============================================================================
// 프로그램 설명 : 간염(A형,C형)전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2010.11.03
// 수정내용      : A형 간염 전결과 비교 분주현황추가
// 수정자        : 정용수
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2010.12.03
// 수정내용      : S019 전결과 비교 분주현황추가
// 수정자        : 정용수
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2011.12.28
// 수정내용      : 이름추가
// 수정자        : 송재호
//==============================================================================
// 프로그램 설명 : 간염(A형,C형), S019 전결과 비교 분주현황
// 시스템        : 통합검진
// 수정일자      : 2012.06.07
// 수정내용      : S010 추가
// 수정자        : 김희석
//==============================================================================


unit LD4Q54;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q54 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    pnlCond: TPanel;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    cmbJisa: TComboBox;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    GroupBox1: TGroupBox;
    RBtn_H002: TRadioButton;
    RBtn_C027: TRadioButton;
    RBtn_C028: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vSBun, UV_vName, UV_vhm, UV_vPhm,UV_vPDat_bunju,UV_vMan : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q54: TfrmLD4Q54;


implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q541;

{$R *.DFM}

procedure TfrmLD4Q54.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo      := VarArrayCreate([0, 0], varOleStr);
      UV_vSBun    := VarArrayCreate([0, 0], varOleStr);
      UV_vName    := VarArrayCreate([0, 0], varOleStr);
      UV_vhm      := VarArrayCreate([0, 0], varOleStr);
      UV_vPhm     := VarArrayCreate([0, 0], varOleStr);
      UV_vPDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vMan:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,     iLength);
      VarArrayRedim(UV_vSBun,   iLength);
      VarArrayRedim(UV_vName,   iLength);
      VarArrayRedim(UV_vhm,   iLength);
      VarArrayRedim(UV_vPhm,  iLength);
      VarArrayRedim(UV_vMan,  iLength);
      VarArrayRedim(UV_vPDat_bunju,  iLength);
   end;
end;

function TfrmLD4Q54.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q54.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD4Q54.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSBun[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vhm[DataRow-1];
      5 : Value := UV_vPhm[DataRow-1];
      6 : Value := UV_vPDat_bunju[DataRow-1]; 
      7 : Value := UV_vMan[DataRow-1];
   end;
end;
procedure TfrmLD4Q54.btnQueryClick(Sender: TObject);
var index, I, rowCount, iRet, iTemp,iAge : Integer;
    sSelect, sWhere, sOrderBy,sMan : String;
    sJangbi_List, sHul_List,hangmok_temp : String;
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
      sSelect := 'select K.desc_rackno,a.num_jumin, a.cod_jisa, a.num_id, a.desc_name, a.cod_jangbi, a.dat_gmgn, a.cod_hul, a.cod_chuga, b.desc_glkwa1, b.desc_glkwa2, b.desc_glkwa3, b.num_bunju From Gumgin a ';
      sSelect := sSelect + ' left outer join gulkwa b ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sSelect := sSelect + ' left  join bunju K ';
      sSelect := sSelect + ' on a.cod_jisa = K.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = K.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = K.dat_gmgn ';

      sWhere := 'Where b.Dat_Bunju = ''' + edtBDate.Text + ''' ';
      if  (RBtn_H002.Checked =True)  then     sWhere := sWhere + ' And  b.Gubn_Part = ''' + '02' + ''' '
      else sWhere := sWhere + ' And  b.Gubn_Part = ''' + '01' + ''' ';

      //20170816 수지  //센터마다 상이하여 보류 .. 김소영
      {sWhere := sWhere + ' and substring(a.cod_jangbi,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(a.cod_jangbi,1,2) <> ''GS''';
      sWhere := sWhere + ' and substring(a.cod_hul,1,2) <> ''SS''';
      sWhere := sWhere + ' and substring(a.cod_hul,1,2) <> ''GS''';   }

      sWhere := sWhere + ' And b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';

       if Trim(bunjuno1.Text) <> '' then
          sWhere := sWhere + ' AND b.num_bunju >= ''' + bunjuno1.Text + '''';
       if Trim(bunjuno2.Text) <> '' then
          sWhere := sWhere + ' AND b.num_bunju <= ''' + bunjuno2.Text + '''';
       sOrderBy := ' ORDER BY G.num_bunju ';




      sOrderBy := ' ORDER BY b.Num_Bunju ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q54', 'Q', 'N');
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
                        if qryPf_hm.FieldByName('COD_HM').AsString = 'JJ10' then
                           sJangbi_List := sJangbi_List + 'JJ20JJ21JJ22JJ23JJ24'
                        else
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

            GP_JuminToAge(qryGulkwa.FieldByname('num_jumin').AsString,qryGulkwa.FieldByname('dat_gmgn').AsString, sMan, iAge);

            with qryPGulkwa do
            begin
               close;
               ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
               ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
               ParamByName('dat_bunju').AsString := edtBDate.text;
               if (RBtn_H002.Checked = True) Then  ParamByName('Gubn_part').AsString := '02'
               else ParamByName('Gubn_part').AsString := '01';
               open;
               
               if RecordCount >= 0 then
               /////////////////////////////////////////////////////////////////
               begin
                   ////////////////////////////////
                  if (RBtn_H002.Checked) and (pos('H002', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           UV_vPhm[index] :=trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='남'
                           else if sMan='F' then UV_vMan[index] :='여';
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_C027.Checked) and (pos('C027', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,133,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,133,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='남'
                           else if sMan='F' then UV_vMan[index] :='여';
                           Inc(index);
                           Gride.Progress := I;
                  end
                  else if (RBtn_C028.Checked) and (pos('C028', sHul_List) > 0) then
                  begin
                           UV_vNo[index]    := intTostr(index+1);
                           UV_vSBun[index]  := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vName[index]  := qryGulkwa.FieldByName('desc_name').AsString;
                           UV_vhm[index]  := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,139,6));
                           UV_vPhm[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,139,6));
                           UV_vPDat_bunju[index] :=trim(qryPGulkwa.FieldByName('Dat_bunju').AsString);
                           if sMan='M' then UV_vMan[index] :='남'
                           else if sMan='F' then UV_vMan[index] :='여';
                           Inc(index);
                           Gride.Progress := I;
                  end;

               end;
               ///////////////////////////////////////////////////////////////
            end;
            next;
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

procedure TfrmLD4Q54.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);

   if Sender = RBtn_H002 then
   Begin
         grdMaster.Col[4].Heading := '혈색소(현)';
         grdMaster.Col[5].Heading := '혈색소(전)';

   end
   else if Sender = RBtn_C027 then
   Begin
         grdMaster.Col[4].Heading := 'LDL(현)';
         grdMaster.Col[5].Heading := 'LDL(전)';

   end
   else if Sender = RBtn_C028 then
   Begin
         grdMaster.Col[4].Heading := '중성지방(현)';
         grdMaster.Col[5].Heading := '중성지방(전)';

   end;



end;



procedure TfrmLD4Q54.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q541 := TfrmLD4Q541.Create(Self);
     frmLD4Q541.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q541 := TfrmLD4Q541.Create(Self);
     frmLD4Q541.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q54.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
procedure TfrmLD4Q54.SBut_ExcelClick(Sender: TObject);
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

