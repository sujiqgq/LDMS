//==============================================================================
// 프로그램 설명 : 분주별 결과 미입력LIST(urin)
// 시스템        : 통합검진
// 수정일자      : 20130902
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================

unit LD4Q50;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q50 = class(TfrmSingle)
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
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    Gride: TGauge;
    Label4: TLabel;
    mskFrom: TMaskEdit;
    Label6: TLabel;
    mskTo: TMaskEdit;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vHangmok, UV_vNum_bunju : Variant;
    iRecordCount : integer;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
  public
    { Public declarations }
  end;

var
  frmLD4Q50: TfrmLD4Q50;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q501;

{$R *.DFM}

procedure TfrmLD4Q50.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vHangmok   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vHangmok,   iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
   end;
end;

function TfrmLD4Q50.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if msk_Bgmgn.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q50.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   msk_Bgmgn.Text := GV_sToday;
end;

procedure TfrmLD4Q50.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNum_bunju[DataRow-1];
      2 : Value := UV_vHangmok[DataRow-1];
   end;
end;

procedure TfrmLD4Q50.btnQueryClick(Sender: TObject);
var index, i,z, k,j, iRet : Integer;
    sSelect, sWhere, sOrderBy, sCod_hm, cod_name,sTemp : String;
    vCod_chuga, vCod_totyh : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; 
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qry_gumgin do
   begin
      // SQL문 생성

      Close;

      sSelect := ' SELECT  B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B inner join gumgin G on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';  


      sWhere := sWhere + ' AND  b.Dat_Bunju = ''' + msk_Bgmgn.Text + ''' ';

      if Trim(mskFrom.Text) <> '' then
         sWhere := sWhere + ' AND b.num_bunju >= ''' + mskFrom.Text + '''';
      if Trim(mskTo.Text) <> '' then
         sWhere := sWhere + ' AND b.num_bunju <= ''' + mskTo.Text + '''';

      sOrderBy := ' ORDER BY b.num_bunju ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qry_gumgin, 'LD4Q50', 'Q', 'N');
         For i := 1 to RecordCount do
         begin
            UP_VarMemSet(index);
            sCod_hm :='';
            sCod_hm := UF_hangmokList;



            if (pos('U001', sCod_hm) > 0) OR (pos('U011', sCod_hm) > 0) OR (pos('U044', sCod_hm) > 0) OR (pos('U034', sCod_hm) > 0) OR  (pos('U045', sCod_hm) > 0) OR
               (pos('U002', sCod_hm) > 0) OR (pos('U012', sCod_hm) > 0) OR (pos('U046', sCod_hm) > 0) OR (pos('U035', sCod_hm) > 0) OR  (pos('U047', sCod_hm) > 0) OR
               (pos('U003', sCod_hm) > 0) OR (pos('U013', sCod_hm) > 0) OR (pos('U051', sCod_hm) > 0) OR (pos('U036', sCod_hm) > 0) OR  (pos('U048', sCod_hm) > 0) OR
               (pos('U004', sCod_hm) > 0) OR (pos('U017', sCod_hm) > 0) OR (pos('Y004', sCod_hm) > 0) OR (pos('U037', sCod_hm) > 0) OR  (pos('U049', sCod_hm) > 0) OR
               (pos('U005', sCod_hm) > 0) OR (pos('U024', sCod_hm) > 0) OR (pos('U019', sCod_hm) > 0) OR (pos('U038', sCod_hm) > 0) OR  (pos('U050', sCod_hm) > 0) OR
               (pos('U006', sCod_hm) > 0) OR (pos('U029', sCod_hm) > 0) OR (pos('U020', sCod_hm) > 0) OR (pos('U039', sCod_hm) > 0) OR  (pos('U052', sCod_hm) > 0) OR
               (pos('U007', sCod_hm) > 0) OR (pos('U030', sCod_hm) > 0) OR (pos('U023', sCod_hm) > 0) OR (pos('U040', sCod_hm) > 0) OR  (pos('Y005', sCod_hm) > 0) OR
               (pos('U008', sCod_hm) > 0) OR (pos('U031', sCod_hm) > 0) OR (pos('U026', sCod_hm) > 0) OR (pos('U041', sCod_hm) > 0) OR  (pos('Y008', sCod_hm) > 0) OR
               (pos('U009', sCod_hm) > 0) OR (pos('U032', sCod_hm) > 0) OR (pos('U027', sCod_hm) > 0) OR (pos('U042', sCod_hm) > 0) OR  (pos('U053', sCod_hm) > 0) OR
               (pos('U010', sCod_hm) > 0) OR (pos('U033', sCod_hm) > 0) OR (pos('U028', sCod_hm) > 0) OR (pos('U043', sCod_hm) > 0) OR  (pos('U054', sCod_hm) > 0) OR
               (pos('U055', sCod_hm) > 0) OR (pos('U056', sCod_hm) > 0) Then
               begin
                  qry_gulkwa.Active := False;
                  qry_gulkwa.ParamByName('num_id').AsString     := FieldByName('num_id').AsString;
                  qry_gulkwa.ParamByName('dat_gmgn').AsString   := FieldByName('dat_gmgn').AsString;
                  qry_gulkwa.Active := True;
                  if qry_gulkwa.RecordCount > 0 then
                  begin
                      while not qry_gulkwa.Eof do
                      begin
                      UV_fGulkwa := '';
                      UV_fGulkwa1 := qry_gulkwa.FieldByName('DESC_GLKWA1').AsString;
                      UV_fGulkwa2 := qry_gulkwa.FieldByName('DESC_GLKWA2').AsString;
                      UV_fGulkwa3 := qry_gulkwa.FieldByName('DESC_GLKWA3').AsString;
                      GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                      Case qry_gulkwa.FieldByName('gubn_part').AsInteger of
                           3 : begin
                             if (qry_gumgin.FieldByName('cod_jisa').AsString='15') then
                             begin
                             if (pos('U001', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa,7,  6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '백혈구(뇨검사), ';
                             if (pos('U002', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 13, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Nitrite(뇨검사), ';
                             if (pos('U003', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 19, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'PH (뇨검사), ';
                             if (pos('U004', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 25, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '단백(뇨검사), ';
                             if (pos('U005', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 31, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '당(뇨검사), ';
                             if (pos('U006', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 37, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '케톤(뇨검사), ';
                             if (pos('U007', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 43, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '유로빌리노겐(뇨), ';
                             if (pos('U008', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 49, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '빌리루빈(뇨검사), ';
                             if (pos('U009', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 55, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '적혈구(뇨검사), ';
                             if (pos('U010', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 61, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '요비중, ';
                             if (pos('U011', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 67, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '뇨침사 WBC, ';
                             if (pos('U012', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 73, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '뇨침사 RBC, ';
                             if (pos('U013', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 79, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '상피세포, ';
                             end;
                             if (pos('Y004', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 85, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '분변잠혈, ';
                             if (pos('Y005', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 91, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '충란검사, ';
                             if (pos('U017', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 115, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '분변배양, ';
                             if (pos('U019', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 127, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '항산균염색(직접도말), ';
                             if (pos('U020', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 133, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '뇨단백정량(24시간), ';
                             if (pos('U023', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 139, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Microalbumin(정량), ';
                             if (pos('U024', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 145, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Amphetamine, ';
                             if (pos('U028', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 163, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Creatinine, ';
                             if (pos('U027', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 169, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '알부민, ';
                             if (pos('U029', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 175, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Cannabinoids, ';
                             if (pos('U030', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 181, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Morphine, ';
                             if (pos('U031', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 187, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Cocaine, ';
                             if (pos('U032', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 193, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'MethAmphetamine, ';
                             if (pos('U033', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 199, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Nicotine, ';
                             if (pos('U034', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 205, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '결핵균 PCR (MTB-PCR), ';
                             if (pos('U035', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 211, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '결핵균 배양, ';
                             if (pos('U036', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 217, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'VMA정량, ';
                             if (pos('U037', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 223, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Phencyclidine(PCP), ';
                             if (pos('U038', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 229, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'stool culture, ';
                             if (pos('Y008', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 235, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'PROTOZOA(STOOL), ';
                             if (pos('U039', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 247, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'aldosterone(24 Urin), ';
                             if (pos('U040', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 253, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'NTx urine, ';
                             if (pos('U041', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 259, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'NTx 24h urine, ';
                             if (pos('U042', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 259, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'NTx 24h urine, ';
                             if (pos('U043', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 271, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Creatinine(부산센타), ';
                             if (pos('U044', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 277, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Heroin  정성, ';
                             if (pos('U045', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 283, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'MDMA(Ecstasy), ';
                             if (pos('U046', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 289, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '세균성 이질, ';
                             if (pos('U047', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 295, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '요중페놀(벤젠), ';
                             if (pos('U048', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 301, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'E.Coli O157 culture, ';
                             if (pos('U049', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 307, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Culture & Sensitivit, ';
                             if (pos('U050', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 313, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '마약 정밀, ';
                             if (pos('U051', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 319, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + '아편(OPI), ';
                             if (pos('U052', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 331, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'CCR, ';
                             if (pos('U053', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 337, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Barbiturate, ';
                             if (pos('U054', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 343, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Benzodiazepine, ';
                             if (pos('U055', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 349, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'Methadone, ';
                             if (pos('U056', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 355, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'drug profiling test, ';

                             end;
                          8 : begin
                             if (pos('U026', sCod_hm) > 0) AND (Trim(Copy(UV_fGulkwa, 535, 6)) = '') then
                                UV_vHangmok[index] := UV_vHangmok[index] + 'TBPE, ';
                             end;
                      end;
                  qry_gulkwa.Next;
                  end;
               end;
            end;
            qry_gulkwa.Active := False;

            if UV_vHangmok[index] <> '' then
            begin
               UV_vNum_bunju[index] := FieldByName('num_bunju').AsString+' ['+FieldByName('num_samp').AsString+']';
               Inc(index);
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

function TfrmLD4Q50.UF_hangmokList : String;
var
   i , j, k : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
       sHmCode :string;
   vCod_chuga : Variant;
   iRet : integer;
begin
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

   for k:=0 to  Round(Length(sTemp)/4)-1 do
   begin
   result:=result+Copy(sTemp, (k+1)*4-3, 4)+'.';
   end;

end;

procedure TfrmLD4Q50.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(msk_Bgmgn);
end;

function  TfrmLD4Q50.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q50.btnPrintClick(Sender: TObject);
begin
  inherited;
  if Sender = btnPrint then
  begin
     if RBtn_preveiw.Checked = True then
     begin
        frmLD4Q501 := TfrmLD4Q501.Create(Self);
        frmLD4Q501.QuickRep.Preview;
     end
     else
     begin
        frmLD4Q501 := TfrmLD4Q501.Create(Self);
        frmLD4Q501.QuickRep.Print;
     end;
  end;
end;

end.
