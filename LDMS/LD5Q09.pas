//==============================================================================
// 프로그램 설명 : 항목리스트조회(LDL-효소법) List
// 시스템        : 분석관리
// 수정일자      : 2014.07.18
// 수정자        : 곽윤설
// 참고사항      : 신규개발
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
unit LD5Q09;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges, ComObj;

type
  TfrmLD5Q09 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    qryGulkwa: TQuery;
    Gride: TGauge;
    mskDate: TDateEdit;
    Cmb_gubn: TComboBox;
    Label9: TLabel;
    Label12: TLabel;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_gmgn, UV_vNum_samp, UV_vDesc_name, UV_vNum_Jumin, UV_C074 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD5Q09: TfrmLD5Q09;
  UV_fGulkwa_01, UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q561, LD4Q562;

{$R *.DFM}

procedure TfrmLD5Q09.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Jumin := VarArrayCreate([0, 0], varOleStr);
      UV_C074       := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vNum_Jumin, iLength);
      VarArrayRedim(UV_C074,      iLength);
   end;
end;

function TfrmLD5Q09.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD5Q09.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
//   UP_GridSet;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 1;
end;

procedure TfrmLD5Q09.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자료를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vdat_gmgn[DataRow-1];
      3 : Value := UV_vNum_samp[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := UV_vNum_Jumin[DataRow-1];
      6 : Value := UV_C074[DataRow-1];
   end;
end;

procedure TfrmLD5Q09.btnQueryClick(Sender: TObject);
var index,DataCol : Integer;
    sSelect, sWhere, sOrderBy, sHangmokList: String;
begin
   inherited;
    sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;
//   UP_GridSet;
   // Grid 초기화
   grdMaster.Rows := 0;
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;       // A : Gumgin, G : Gulkwa
      sSelect := ' Select A.num_jumin, A.Dat_gmgn, A.desc_name, G.cod_jisa, G.dat_bunju, G.num_id, A.num_samp, G.desc_glkwa1, G.desc_glkwa2, G.DESC_GLKWA3, G.num_bunju, G.gubn_part, ' +
                 '        A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.num_jumin, ' +
                 '        A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga ';

      sSelect := sSelect + ' From gulkwa G with(nolock) left outer join Gumgin A with(nolock)';
      sSelect := sSelect + '                         On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';

      sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''' +
                '       AND  A.Dat_Gmgn = ''' + mskDate.Text + ''' ';
      sWhere := sWhere + 'AND  (G.Gubn_Part = ''01'')';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY G.num_bunju ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY A.num_samp '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD5Q09', 'Q', 'N');
         while Not Eof do
         begin
            Gride.Progress := Gride.Progress + 1;

            sHangmokList := '';
            sHangmokList := UF_hangmokList;

            if (qryGulkwa.FieldByName('gubn_part').AsString = '01') then   //생화학 결과
            begin
               UV_fGulkwa_01 := '';
               UV_fGulkwa1_01 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2_01 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3_01 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01, UV_fGulkwa_01);
            end;

            if (pos('C074',sHangmokList) > 0) then
            begin
               UP_VarMemSet(index);
               UV_vNo[index]       := IntToStr(Index+1);
               UV_vDat_gmgn[index] := qryGulkwa.FieldByName('dat_gmgn').AsString;
               UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
               UV_vNum_Jumin[index]:= copy(qryGulkwa.FieldByName('num_jumin').AsString,1,6) + '-' + copy(qryGulkwa.FieldByName('num_jumin').AsString,7,13);
               UV_C074[index]      := Trim(Copy(UV_fGulkwa_01,  331,  6));

               Inc(index);
            end;
            next;
         end//begin of while
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

procedure TfrmLD5Q09.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);
   begin
      if (Cmb_gubn.ItemIndex = 0) then
      begin
         mskFrom.Enabled := True;
         mskTo.Enabled := True;
         MEdt_SampS.Enabled := False;
         MEdt_SampE.Enabled := False;
      end
      else if (Cmb_gubn.ItemIndex = 1) then
      begin
         mskFrom.Enabled := False;
         mskTo.Enabled := False;
         MEdt_SampS.Enabled := True;
         MEdt_SampE.Enabled := True;
      end;
   end;
end;


function TfrmLD5Q09.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
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
      if qryGulkwa.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_cancer').AsString;
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
      if qryGulkwa.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryGulkwa.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryGulkwa.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '1', qryGulkwa.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGulkwa.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '4', qryGulkwa.FieldByName('Gubn_adyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '7', qryGulkwa.FieldByName('Gubn_lfyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '3',qryGulkwa.FieldByName('Gubn_bgyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '5',qryGulkwa.FieldByName('Gubn_agyh').AsInteger);

   if (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '1') or (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGulkwa.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGulkwa.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGulkwa.FieldByName('Dat_gmgn').AsString;
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

function TfrmLD5Q09.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

procedure TfrmLD5Q09.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then //미리보기
  begin
        frmLD4Q561 := TfrmLD4Q561.Create(Self);
        frmLD4Q561.QuickRep.Preview;
  end
  else   //프린트
  begin
        frmLD4Q561 := TfrmLD4Q561.Create(Self);
        frmLD4Q561.QuickRep.Print;
  end;
end;

end.
