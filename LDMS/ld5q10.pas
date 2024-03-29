//==============================================================================
// 프로그램 설명 : 생화학 결과 미입력 LIST
// 시스템        : 통합검진
// 수정일자      : 2015.02.02
// 수정자        : 곽윤설
// 수정내용      :   
// 참고사항      : 한의 재단진단검사의학팀1500065
//==============================================================================
// 수정일자      : 2015.03.16
// 수정자        : 곽윤설
// 수정내용      : 성명추가
// 참고사항      : 본원-문지혜
//==============================================================================
// 수정일자      : 2015.03.20
// 수정자        : 곽윤설
// 수정내용      : 오류수정 - 특검조회 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.06.16
// 수정자        : 곽윤설
// 수정내용      : 항목추가 (S001, S003, S004, S036, S034, C080, C022, C023)
// 참고사항      : 한의 재단진단검사의학팀1500435
//==============================================================================
// 수정일자      : 2016.03.29
// 수정자        : 박수지
// 수정내용      : 항목추가 C078
// 참고사항      : 한의 재단진단검사의학팀1600348
//==============================================================================
// 수정일자      : 2016.07.09
// 수정자        : 박수지
// 수정내용      : 항목추가 C084 , C085
// 참고사항      : 한의 재단진단검사의학팀1600611
//==============================================================================
// 수정일자      : 2018.04.25
// 수정자        : 박수지
// 수정내용      : 항목추가 C090
// 참고사항      : 한의 재단일반파트 1800324
//==============================================================================
unit ld5q10;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt,ComObj;

type
  TfrmLD5Q10 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    qry_gumgin: TQuery;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    msk_Bgmgn: TDateEdit;
    qry_Hangmok: TQuery;
    qry_gulkwa: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Gride: TGauge;
    mskFrom_Bunju: TMaskEdit;
    Label6: TLabel;
    mskTo_Bunju: TMaskEdit;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    GroupBox1: TGroupBox;
    CB_A1c: TCheckBox;
    CB_P: TCheckBox;
    CB_D: TCheckBox;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    RB_samp: TRadioButton;
    RB_RackNo: TRadioButton;
    RB_bunju: TRadioButton;
    mskFrom_samp: TMaskEdit;
    mskTo_samp: TMaskEdit;
    Label4: TLabel;
    msk_RackNo: TMaskEdit;
    RG_order: TRadioGroup;
    qryTkgum: TQuery;
    qryProfile: TQuery;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure Change(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vHangmok, UV_vNum_samp, UV_vNum_bunju, UV_vRackNo, UV_vName : Variant;
    iRecordCount, iArr : integer;
    vArr_HM : array of array of String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
  end;

var
  frmLD5Q10: TfrmLD5Q10;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q101;

{$R *.DFM}

procedure TfrmLD5Q10.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vHangmok   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vRackNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vHangmok,   iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vRackNo,    iLength);
      VarArrayRedim(UV_vName,      iLength);
   end;
end;

function TfrmLD5Q10.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if msk_Bgmgn.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q10.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   msk_Bgmgn.Text := GV_sToday;
   mskFrom_Bunju.Text := '0001';
   mskTo_Bunju.Text := '9999';
end;

procedure TfrmLD5Q10.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNum_bunju[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vRackNo[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vHangmok[DataRow-1];
   end;
end;

procedure TfrmLD5Q10.btnQueryClick(Sender: TObject);
var index, i, j, iRet : Integer;
    sSelect, sWhere, sOrderBy, sCod_hm, sHangmok, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa_01, UV_fGulkwa_07 : String;
    vCod_chuga : Variant;
    vcheck : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   sHangmok := '';
                                                                                                                                  ;
   if CB_D.checked then
   begin
        sHangmok := sHangmok + 'C001C002C003C004C005C006C007C009C010C011C012C013C017C019C021' +
                    'C025C026C027C028C029C030C032C033C039C041C042C043C045C046C047C048C049C050C051' +
                    'C052C056C057C058C087GL01GL02GL03GL04GL05GL06S003S004S034';
   end;
   if CB_P.checked then
   begin
        sHangmok := sHangmok + 'C022C023C024C034C074C078C080C082C084C085C090S036';
   end;
   if CB_A1c.checked then
   begin
        sHangmok := sHangmok + 'C083';
   end;
   iArr := Round(length(sHangmok)/4);

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qry_gumgin do
   begin
      // SQL문 생성
      Close;
      sSelect := 'SELECT a.cod_jisa, a.num_id, a.dat_gmgn, a.Desc_name, a.num_samp, a.gubn_nsyh, a.gubn_adyh, a.gubn_agyh, a.gubn_lfyh, a.cod_hul, a.cod_jangbi, a.Gubn_tkgm,     ' +
                 'a.cod_cancer, a.cod_chuga, a.gubn_nosin, a.gubn_adult, a.Gubn_bogun, a.Gubn_bgyh, a.gubn_agab, a.gubn_life, b.cod_bjjs, b.num_bunju, b.desc_rackno, a.Num_jumin ' +
                 'From Gumgin a with(nolock) inner join bunju b with(nolock) on a.cod_jisa = b.cod_jisa and a.num_id = b.num_id and a.dat_gmgn = b.dat_gmgn                       ';

      sWhere := ' WHERE b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere := sWhere + ' AND  b.Dat_Bunju = ''' + msk_Bgmgn.Text + ''' ';

      if RB_bunju.Checked then
      begin
         if Trim(mskFrom_Bunju.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju >= ''' + mskFrom_Bunju.Text + ''' ';
         if Trim(mskTo_Bunju.Text) <> '' then
            sWhere := sWhere + ' AND b.num_bunju <= ''' + mskTo_Bunju.Text + ''' ';
      end
      else if RB_Samp.Checked then
      begin
         if Trim(mskFrom_Samp.Text) <> '' then
            sWhere := sWhere + ' AND a.num_Samp >= ''' + mskFrom_Samp.Text + ''' ';
         if Trim(mskTo_Samp.Text) <> '' then
            sWhere := sWhere + ' AND a.num_Samp <= ''' + mskTo_Samp.Text + ''' ';
      end
      else if RB_RackNo.Checked then
      begin
         if Trim(msk_RackNo.Text) <> '' then
            sWhere := sWhere + ' AND b.desc_RackNo LIKE ''' + msk_RackNo.Text + '%'' ';
      end;

      if   RG_Order.ItemIndex = 0 then sOrderBy := ' ORDER BY b.num_bunju '
      else sOrderBy := ' ORDER BY a.cod_jisa, a.num_samp  ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qry_gumgin, 'LD5Q10', 'Q', 'N');

         //20150616
         SetLength(vArr_HM, iArr);
         qry_hangmok.Active := False;
         qry_hangmok.Filtered := False;
         qry_hangmok.Active := True;
         for j := 0 to iArr-1 do
         begin
             SetLength(vArr_HM[j], 4);
             qry_hangmok.Filtered := True;
             qry_hangmok.Filter := ' cod_hm = ''' + copy(sHangmok,(j*4)+1,4) + ''' ';
             if qry_hangmok.RecordCount > 0 then
             begin
                 vArr_HM[j][0] := qry_hangmok.FieldByName('cod_hm').AsString;
                 vArr_HM[j][1] := qry_hangmok.FieldByName('gubn_part').AsString;
                 vArr_HM[j][2] := qry_hangmok.FieldByName('desc_kor').AsString;
                 vArr_HM[j][3] := qry_hangmok.FieldByName('pos_Start').AsString;
             end;
         end;
         qry_hangmok.Filtered := False;
         qry_hangmok.Active := False;

         For i := 1 to RecordCount do
         begin
             Gride.Progress := i;
             UP_VarMemSet(index);

             //검사항목추출
             sCod_hm := '';
             sCod_hm := UF_hangmokList;

             vcheck := False;
             for j := 0 to iArr-1 do
             begin
                 if pos(vArr_HM[j][0], sCod_hm) > 0 then vcheck := True;
                 if vcheck then break;
             end;

             if vcheck then
             begin
                qry_gulkwa.Active := False;
                qry_gulkwa.ParamByName('cod_bjjs').AsString   := Copy(cmbJisa.Text,1,2);
                qry_gulkwa.ParamByName('num_id').AsString     := FieldByName('num_id').AsString;
                qry_gulkwa.ParamByName('dat_gmgn').AsString   := FieldByName('dat_gmgn').AsString;
                qry_gulkwa.Active := True;
                if qry_gulkwa.RecordCount > 0 then
                begin
                   while not qry_gulkwa.eof do
                   begin
                       if qry_gulkwa.FieldByName('gubn_part').AsString = '01' then
                       begin
                          UV_fGulkwa1 := qry_gulkwa.FieldByName('DESC_GLKWA1').AsString;
                          UV_fGulkwa2 := qry_gulkwa.FieldByName('DESC_GLKWA2').AsString;
                          UV_fGulkwa3 := qry_gulkwa.FieldByName('DESC_GLKWA3').AsString;
                          GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa_01);
                       end
                       else if qry_gulkwa.FieldByName('gubn_part').AsString = '07' then
                       begin
                          UV_fGulkwa1 := qry_gulkwa.FieldByName('DESC_GLKWA1').AsString;
                          UV_fGulkwa2 := qry_gulkwa.FieldByName('DESC_GLKWA2').AsString;
                          UV_fGulkwa3 := qry_gulkwa.FieldByName('DESC_GLKWA3').AsString;
                          GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa_07);
                       end;
                       qry_gulkwa.next;
                   end;

                   For j := 0 to iArr-1 do
                   begin
                       if vArr_HM[j][1] = '01' then
                       begin
                          if (pos(vArr_HM[j][0], sCod_hm) > 0) AND (trim(copy(UV_fGulkwa_01, StrToInt(vArr_HM[j][3]), 6)) = '') then
                              UV_vHangmok[index] := UV_vHangmok[index] + vArr_HM[j][2] + '(' + vArr_HM[j][0] + '), ';
                       end
                       else if vArr_HM[j][1] = '07' then
                       begin
                          if (pos(vArr_HM[j][0], sCod_hm) > 0) AND (trim(copy(UV_fGulkwa_07, StrToInt(vArr_HM[j][3]), 6)) = '') then
                              UV_vHangmok[index] := UV_vHangmok[index] + vArr_HM[j][2] + '(' + vArr_HM[j][0] + '), ';
                       end;
                   end;
                end;
             end;
             qry_gulkwa.Active := False;

             if UV_vHangmok[index] <> '' then
             begin
                UV_vNum_bunju[index] := FieldByName('num_bunju').AsString;
                UV_vNum_samp[index]  := FieldByName('num_samp').AsString;
                UV_vRackNo[index]    := FieldByName('desc_rackno').AsString;
                UV_vName[index]      := FieldByName('desc_name').AsString;
                UV_vHangmok[index]   := copy(UV_vHangmok[index],1,length(UV_vHangmok[index])-2);
                Inc(index);
             end;

             next;
         end;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      vArr_HM := nil;
      Close;
   end; // qryGulkwa

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD5Q10.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(msk_Bgmgn);
end;

function  TfrmLD5Q10.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q10.btnPrintClick(Sender: TObject);
begin
  inherited;
  if Sender = btnPrint then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD5Q101 := TfrmLD5Q101.Create(Self);
        frmLD5Q101.QuickRep.Preview;
        frmLD5Q101.free;
     end
     else
     begin
        frmLD5Q101 := TfrmLD5Q101.Create(Self);
        frmLD5Q101.QuickRep.Print;
     end;
  end;
end;


procedure TfrmLD5Q10.SBut_ExcelClick(Sender: TObject);
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


procedure TfrmLD5Q10.Change(Sender: TObject);
begin
  inherited;
  if RB_bunju.Checked then
  begin
     mskFrom_Bunju.Enabled  := True;
     mskTo_Bunju.Enabled    := True;
     mskFrom_Samp.Enabled   := False;
     mskTo_Samp.Enabled     := False;
     msk_RackNo.Enabled := False;
  end
  else if RB_Samp.Checked then
  begin
     mskFrom_Bunju.Enabled  := False;
     mskTo_Bunju.Enabled    := False;
     mskFrom_Samp.Enabled   := True;
     mskTo_Samp.Enabled     := True;
     msk_RackNo.Enabled := False;
  end
  else if RB_Rackno.Checked then
  begin
     mskFrom_Bunju.Enabled  := False;
     mskTo_Bunju.Enabled    := False;
     mskFrom_Samp.Enabled   := False;
     mskTo_Samp.Enabled     := False;
     msk_RackNo.Enabled     := True;
  end
end;

function TfrmLD5Q10.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qry_gumgin.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qry_gumgin.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if qry_gumgin.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qry_gumgin.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if qry_gumgin.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qry_gumgin.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qry_gumgin.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qry_gumgin.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qry_gumgin.FieldByName('Dat_gmgn').AsString, '1', qry_gumgin.FieldByName('Gubn_nsyh').AsInteger)
   else if qry_gumgin.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qry_gumgin.FieldByName('Dat_gmgn').AsString, '4', qry_gumgin.FieldByName('Gubn_adyh').AsInteger);

   if qry_gumgin.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qry_gumgin.FieldByName('Dat_gmgn').AsString, '7', qry_gumgin.FieldByName('Gubn_lfyh').AsInteger);

   if qry_gumgin.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qry_gumgin.FieldByName('Dat_gmgn').AsString, '3',qry_gumgin.FieldByName('Gubn_bgyh').AsInteger);

   if qry_gumgin.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qry_gumgin.FieldByName('Dat_gmgn').AsString, '5',qry_gumgin.FieldByName('Gubn_agyh').AsInteger);

   if (qry_gumgin.FieldByName('Gubn_tkgm').AsString = '1') or (qry_gumgin.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qry_gumgin.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qry_gumgin.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qry_gumgin.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   result := sTemp;
end;

end.
