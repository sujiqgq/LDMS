//==============================================================================
// 프로그램 설명 : S021 분주현황
// 시스템        : 통합검진
// 수정일자      : 2009.11.10
// 수정자        : 송재호
// 수정내용      : 검사법 추가로 S021 추가(홀수 플레이트만 검사항목 기재)
//==============================================================================
// 프로그램 설명 : S021 분주현황
// 시스템        : 통합검진
// 수정일자      : 2012.01.11
// 수정자        : 송재호
// 수정내용      : S019, S021 분주현황으로 프로그램명 변경 후 S019 추가,
//                 Grid에 성명 추가, S019 List 및 폼텍 출력 추가(한의 재단EIA팀1100108)
//==============================================================================
unit LD4Q23;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q23 = class(TfrmSingle)
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnSdate: TSpeedButton;
    edtSDate: TDateEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    grdS021: TtsGrid;
    QRCompositeReport: TQRCompositeReport;
    btnEdate: TSpeedButton;
    edtEDate: TDateEdit;
    Label1: TLabel;
    Cmb_gubn: TComboBox;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label12: TLabel;
    Label3: TLabel;
    mskFrom: TMaskEdit;
    Label5: TLabel;
    mskTo: TMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure grdS021CellLoaded(Sender: TObject; DataCol, DataRow: Integer;
      var Value: Variant);





  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_Bunju, UV_vNum_Bunju,UV_vNum_samp, UV_vName, UV_vE068, UV_vCheck,UV_vRack : Variant;
    UV_iCount : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q23: TfrmLD4Q23;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q234;

{$R *.DFM}

procedure TfrmLD4Q23.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vE068      := VarArrayCreate([0, 0], varOleStr);

  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vDat_Bunju, iLength);
      VarArrayRedim(UV_vNum_Bunju, iLength);
      VarArrayRedim(UV_vNum_Samp,  iLength);
      VarArrayRedim(UV_vName,      iLength);
      VarArrayRedim(UV_vE068,      iLength);
   end;
end;


function TfrmLD4Q23.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtSDate.Text = '') OR (edtEDate.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
   // 하루가 아닐경우 분주번호 지정 못함
   if edtSDate.Text <> edtEDate.Text then
   begin
      if (mskFrom.Text <> '0001') OR (mskTo.Text <> '9999') then
      begin
         ShowMessage('이틀 이상 조회시 분주번호 지정 불가!!!');
         Result := False;
      end;
   end;
end;

procedure TfrmLD4Q23.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   Cmb_gubn.ItemIndex := 0;
   UV_iCount := 0;
end;

procedure TfrmLD4Q23.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdS021.Rows := 0;
   UP_VarMemSet(0);
   
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT G.num_samp,B.desc_RackNo,K.dat_bunju, K.num_Bunju, G.desc_name, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga ' +
                 ' FROM Gulkwa K left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '               left outer join Bunju B ON K.cod_jisa = B.cod_jisa and K.num_id = B.num_id and K.dat_gmgn = B.dat_gmgn ' +
                 '               left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';
      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND K.Dat_Bunju  >= ''' + edtSDate.Text + ''' ' +
                          ' AND K.Dat_Bunju  <= ''' + edtEDate.Text + ''' ' +
                          ' AND K.Gubn_Part  = ''' + '05' + ''' ';
    

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND K.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND K.num_bunju <= ''' + mskTo.Text + '''';
            sOrderBy := ' ORDER BY K.dat_bunju, K.Num_Bunju';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
         sOrderBy := ' ORDER BY K.dat_bunju, K.Num_Bunju';
      end;


      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q23', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;

            // 항목(장비프로파일)추출..
            sJangbi_List := ''; sHul_List := '';
            if qryBunju.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_jangbi').AsString;
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

            // 항목(혈액프로파일)추출..
            if qryBunju.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
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

            // 항목(종양프로파일)추출..
            if qryBunju.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_cancer').AsString;
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
            iRet := GF_MulToSingle(qryBunju.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
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

            //(특검)추가검사항목..
            iRet := GF_MulToSingle(qryBunju.FieldByName('Tcod_chuga').AsString, 4, vCod_chuga);
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



            if (pos('E068', sHul_List) > 0) then
            begin
               UP_VarMemSet(Index);
               
               UV_vNo[Index]  := IntToStr(Index + 1);
               UV_vDat_Bunju[Index] := qryBunju.FieldByName('dat_bunju').AsString;
               UV_vNum_samp[Index] := qryBunju.FieldByName('Num_samp').AsString;
               UV_vNum_Bunju[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               UV_vName[Index]      := qryBunju.FieldByName('desc_name').AsString;

               if pos('E068', sHul_List) > 0 then UV_vE068[Index] := '*'
               else                               UV_vE068[Index] := '';

               Inc(Index);
            end;

            UV_iCount := Index;
            Next;
         End;
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
   grdS021.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q23.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnSdate then GF_CalendarClick(edtSDate);
   if Sender = btnEdate then GF_CalendarClick(edtEDate);
end;

procedure TfrmLD4Q23.btnPrintClick(Sender: TObject);
begin
  inherited;

  frmLD4Q234 := TfrmLD4Q234.Create(Self);

  if RBtn_preveiw.Checked = True then
     frmLD4Q234.QuickRep.Preview
  else
     frmLD4Q234.QuickRep.Print;




end;

procedure TfrmLD4Q23.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   edtSdate.SetFocus;
end;

procedure TfrmLD4Q23.grdS021CellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDat_Bunju[DataRow-1];
      3 : Value := UV_vNum_Samp[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vNum_bunju[DataRow-1];
      6 : Value := UV_vE068[DataRow-1];      
     
   end;

end;





end.
