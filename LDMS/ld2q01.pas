//==============================================================================
// 프로그램 설명 : 분주 list 현황
// 시스템        : 통합검진
// 수정일자      : 2005.10.04
// 수정자        : 김승철
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 분주 list 현황
// 시스템        : 통합검진
// 수정일자      : 2011.11.09
// 수정자        : 김희석
// 수정내용      :  주민번호 뒷자리 * 표시
//==============================================================================
// 프로그램 설명 : 분주 list 현황
// 시스템        : 통합검진
// 수정일자      : 2014.03.25
// 수정자        : 곽윤설
// 수정내용      : 조회구분 - U059, U060, U061 추가
//==============================================================================
// 프로그램 설명 : 분주 list 현황
// 시스템        : 통합검진
// 수정일자      : 2014.04.15
// 수정자        : 곽윤설
// 수정내용      : 조회구분 -  (1) M301, M302, M401, M402, M500 -> 메디젠휴먼케어(주) 코드 추가　
//                             (2) S501, S502, S503 -> SK케미칼 코드 추가
//==============================================================================
// 수정일자      : 2014.04.25
// 수정자        : 곽윤설
// 수정내용      : 조회구분 -  (1) 메디젠휴먼케어(주) 코드 추가 -> + M300
//                             (2) SK케미칼 코드 추가 -> + S504, S505, S506, S507, S508, S509, S510, S511
//==============================================================================
// 수정일자      : 2014.05.26
// 수정자        : 곽윤설
// 수정내용      : 유전자 검사 - 검진테이블
//==============================================================================
// 수정일자      : 2014.06.02
// 수정자        : 곽윤설
// 수정내용      : 유전자 검사 - M000 : 연구소 조회시 강남,여의도 제외
//                             - S000 : 각 지사별로 조회
//==============================================================================
// 수정일자      : 2014.07.21
// 수정자        : 곽윤설
// 수정내용      : 유전자 검사  M000~ : 총합계 출력
//==============================================================================
// 수정일자      : 2015.04.15
// 수정자        : 곽윤설
// 수정내용      : 리스트 상 항목 추가 : MM01, MF01, MC01, M501
//                 리스트 상 항목 삭제 : M500(미사용코드)
//==============================================================================
// 수정일자      : 2015.04.23
// 수정자        : 곽윤설
// 수정내용      : Y005,Y008,U017,U019,U046,U058,U035에 U015,U016추가
// 참고사항      : 한의 재단진단검사의학팀1500322
//==============================================================================
// 수정일자      : 2015.05.18
// 수정자        : 곽윤설
// 수정내용      : M403 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.10.14
// 수정내용      : 조회구분 - U059,U060,U061 -> U059,U060,U061,U063,U064,U065 변경
// 참고사항      : [한의 재단진단검사의학팀1500856] U063,U064,U065 신규코드추가
//==============================================================================
// 수정일자      : 2015.10.21
// 수정내용      : 리스트 상 항목 추가 : M500(미사용코드) →부산센터만 일시적으로 조회되도록 요청
// 참고사항      : [한의 부산진단검사의학팀1500921]
//==============================================================================
// 수정일자      : 2015.10.22
// 수정내용      : 메디젠 유전자 검사 : M500,MC3A,MC3B,MC4A,MM3B,MF4D,MC06,MC12,MC04,MF04,MM13,MF15,MM48,MF52 추가
// 참고사항      : [한의 재단진단검사의학팀1500879]
//==============================================================================
// 수정일자      : 2016.02.20
// 수정내용      : 메디젠 유전자 검사 : M304추가
// 참고사항      : [한의 재단진단검사의학팀1600187]
//==============================================================================
// 수정일자      : 2016.03.17
// 수정내용      : C078,E069,C022,C023,C0080 분주List 현황 _김소영 선생님
// 참고사항      : 본원진료팀 1600205
//==============================================================================
// 수정일자      : 2016.05.28
// 수정내용      : 케어링크  분주List 현황
// 참고사항      : 재단진단검사의학팀 1600510
//==============================================================================
// 수정일자      : 2017.06.13
// 수정내용      : MG01-06 유전자 검사 추가
// 참고사항      : 광주 나미영 책임 요청(광주진단검사의학팀1700512)
//==============================================================================
unit LD2Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD2Q01 = class(TfrmSingle)
    qryBunju: TQuery;
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    Label10: TLabel;
    mskTo: TMaskEdit;
    Label3: TLabel;
    edtDc: TEdit;
    btnDc: TSpeedButton;
    edtD_dc: TEdit;
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    RG_Query: TComboBox;
    Label4: TLabel;
    Ck_dept: TCheckBox;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label5: TLabel;
    MEdt_SampE: TMaskEdit;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    CB_01: TCheckBox;
    qryDNA: TQuery;
    PopupMenu1: TPopupMenu;
    Label6: TLabel;
    mskFDate: TDateEdit;
    btnDate_To: TSpeedButton;
    procedure FormCreate
    (Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    UV_sCod_jisa : String;
    // SQL문
    UV_sBasicSQL : String;

    //메디젠,SK 유전자, 케어링크
    Arr_DNA : Array of Array of String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    procedure UP_Clear(Temp : Integer);
    function  UP_Gumgin_YH(sJisa, sID, sHul, sJangbi, sCancer, sHulgum, sNosin, sTkgm, sAdult, sAgab, sGong, sLife : String;
                           iNosin, iTkgm, iAdult, iAgab, iGong, iLife : Integer) : String;
  public
    { Public declarations }
    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_vNum_bunju, UV_vDesc_dc, UV_vDesc_name, UV_vNum_jumin, UV_vCod_jisa,
    UV_vCod_jangbi, UV_vCod_hul, UV_vCod_Cancer, UV_vCod_chuga, UV_vDat_gmgn,
    UV_vCod_etc, UV_vNum_samp : Variant;

  end;

var
  frmLD2Q01: TfrmLD2Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q011, LD2Q012;

{$R *.DFM}

procedure TfrmLD2Q01.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];

   //--------------------------------------------------------------------------
   // 내림차순
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]  := UV_vNum_bunju[Hi];  UV_vNum_bunju[Hi] := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]   := UV_vNum_samp[Hi];   UV_vNum_samp[Hi]  := T;
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]    := UV_vDesc_dc[Hi];    UV_vDesc_dc[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]  := UV_vDesc_name[Hi];  UV_vDesc_name[Hi] := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]  := UV_vNum_jumin[Hi];  UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];    UV_vDat_gmgn[Lo]   := UV_vDat_gmgn[Hi];   UV_vDat_gmgn[Hi]  := T;
            T := UV_vCod_jisa[Lo];    UV_vCod_jisa[Lo]   := UV_vCod_jisa[Hi];   UV_vCod_jisa[Hi]  := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo] := UV_vCod_jangbi[Hi]; UV_vCod_jangbi[Hi]:= T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]    := UV_vCod_hul[Hi];    UV_vCod_hul[Hi]   := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo] := UV_vCod_Cancer[Hi]; UV_vCod_Cancer[Hi]:= T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]  := UV_vCod_chuga[Hi];  UV_vCod_chuga[Hi] := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]    := UV_vCod_etc[Hi];    UV_vCod_etc[Hi]   := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // 오름차순
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]  := UV_vNum_bunju[Hi];  UV_vNum_bunju[Hi] := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]   := UV_vNum_samp[Hi];   UV_vNum_samp[Hi]  := T;            
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]    := UV_vDesc_dc[Hi];    UV_vDesc_dc[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]  := UV_vDesc_name[Hi];  UV_vDesc_name[Hi] := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]  := UV_vNum_jumin[Hi];  UV_vNum_jumin[Hi] := T;
            T := UV_vDat_gmgn[Lo];    UV_vDat_gmgn[Lo]   := UV_vDat_gmgn[Hi];   UV_vDat_gmgn[Hi]  := T;
            T := UV_vCod_jisa[Lo];    UV_vCod_jisa[Lo]   := UV_vCod_jisa[Hi];   UV_vCod_jisa[Hi]  := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo] := UV_vCod_jangbi[Hi]; UV_vCod_jangbi[Hi]:= T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]    := UV_vCod_hul[Hi];    UV_vCod_hul[Hi]   := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo] := UV_vCod_Cancer[Hi]; UV_vCod_Cancer[Hi]:= T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]  := UV_vCod_chuga[Hi];  UV_vCod_chuga[Hi] := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]    := UV_vCod_etc[Hi];    UV_vCod_etc[Hi]   := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD2Q01.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading  := '분주번호';
      Col[2].Heading  := '샘플번호';
      Col[3].Heading  := '성  명';
      Col[4].Heading  := '주민번호';
      Col[5].Heading  := '검진일자';
      Col[6].Heading  := '단 체 명';
      Col[7].Heading  := '지사';
      Col[8].Heading  := '장비코드';
      Col[9].Heading  := '혈액코드';
      Col[10].Heading  := '종양코드';
      Col[11].Heading := '추가코드';
      Col[12].Heading := '추가검사';

      // Column Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[6].Alignment := taCenter;
      Col[7].Alignment := taCenter;
      Col[8].Alignment := taCenter;
      Col[9].Alignment := taCenter;

      for i := 1 to Cols do
         Col[i].SortPicture := spDown;
   end;

   IF  GV_sJisa <> '15' THEN
   Begin
      RG_Query.Items.Add('C074검사자 조회');
   end;

end;

procedure TfrmLD2Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Cancer := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_etc    := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNum_bunju,  iLength);
      VarArrayRedim(UV_vNum_samp,   iLength);
      VarArrayRedim(UV_vDesc_dc,    iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vCod_jangbi, iLength);
      VarArrayRedim(UV_vCod_hul,    iLength);
      VarArrayRedim(UV_vCod_Cancer, iLength);
      VarArrayRedim(UV_vCod_chuga,  iLength);
      VarArrayRedim(UV_vCod_etc,    iLength);
   end;
end;

function TfrmLD2Q01.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD2Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;
   UP_VarMemSet(0);

   mskDate.Text := Trim(GV_sToday);

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   //20170405 부산센터 유희짐 요청.. 부산만 적용
   if (GV_sJisa = '61') or (GV_sJisa = '71') then
   begin
   Label6.Visible := True;
   mskFDate.Visible := True;
   btnDate_To.Visible := True;
   end;
   if (GV_sJisa <> '61') and (GV_sJisa <> '71') then
   begin
   Label12.Top := Label12.Top - 20;
   Label9.Top  := Label9.Top - 20;
   Label5.Top  := Label5.Top - 20;
   Label10.Top := Label10.Top - 20;
   MEdt_SampS.Top  := MEdt_SampS.Top - 20;
   MEdt_SampE.Top  := MEdt_SampE.Top - 20;
   mskFrom.Top     := mskFrom.Top - 20;
   mskTo.Top       := mskTo.Top - 20;
   end;
   if UV_sCod_jisa = '15' then
   begin
      Label1.Caption := '지 사:';
      GP_ComboJisa(cmbCOD_JISA);
      CB_01.visible := False;
   end
   else
   begin
      Label1.Caption := '분주소:';
      cmbCod_jisa.Items.Add('연 구 소');   //0
      cmbCod_jisa.Items.Add('지    사');   //1
      cmbCod_jisa.ItemIndex := 1;
      CB_01.visible := False;
   end;

   // SQL문을 저장
   UV_sBasicSQL := qryBunju.SQL.Text;
   Cmb_gubn.ItemIndex  := 0;
   RG_Query.ItemIndex := 0;
   Cmb_gubn.ItemIndex := 0;

end;

procedure TfrmLD2Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNum_bunju[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := UV_vNum_jumin[DataRow-1];
      5 : Value := UV_vDat_gmgn[DataRow-1];
      6 : Value := UV_vDesc_dc[DataRow-1];
      7 : Value := UV_vCod_jisa[DataRow-1];
      8 : Value := UV_vCod_jangbi[DataRow-1];
      9 : Value := UV_vCod_hul[DataRow-1];
     10 : Value := UV_vCod_Cancer[DataRow-1];
     11 : Value := UV_vCod_chuga[DataRow-1];
     12 : Value := UV_vCod_etc[DataRow-1];
   end;

   if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex =6) or (RG_Query.ItemIndex =7) or
      (RG_Query.ItemIndex = 8) or (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 11) or (RG_Query.ItemIndex = 12) or (RG_Query.ItemIndex = 13) then  //16.03.15 E068를 당일로 확인 가능하도록.. 이청심
   begin
      grdMaster.Col[1].Heading := 'No.';    //20140415 곽윤설 - 분주번호->No, col[6~10]->안보이게
      grdMaster.Col[6].Width := 0;
      grdMaster.Col[8].Width := 0;
      grdMaster.Col[9].Width := 0;
      grdMaster.Col[10].Width := 0;
      grdMaster.Col[12].Width := 0;
   end
   else
   begin
      grdMaster.Col[1].Heading := '분주번호';
      grdMaster.Col[6].Width := 50;
      grdMaster.Col[7].Width := 70;
      grdMaster.Col[8].Width := 70;
      grdMaster.Col[9].Width := 70;
      grdMaster.Col[10].Width := 80;
      grdMaster.Col[12].Width := 70;
   end;

end;

procedure TfrmLD2Q01.btnQueryClick(Sender: TObject);
var i, j, k, index, iRet, temp : Integer;
    sWhere, sOrder, chuga, tk_chuga, chuga_hul, chuga_print, sLike : String;
    vCod_chuga : Variant;
    flag_01, flag_02, flag_03, flag_04, flag_05, flag_06 : Integer;
begin
   inherited;

   //유전자 검사항목
   Arr_DNA := nil;  sLike := '';
   if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 6) or (RG_Query.ItemIndex = 11) then
   begin
      with qryDNA do
      begin
         Active := False;
         if RG_Query.ItemIndex = 3 then
           ParamByName('sDNA').AsString := 'M'
         else if RG_Query.ItemIndex = 6 then
           ParamByName('sDNA').AsString := 'C'  //케어링크 유정자 160528 수지
         else if RG_Query.ItemIndex = 4 then
           ParamByName('sDNA').AsString := 'S'
         else ParamByName('sDNA').AsString := 'D';

         Active := True;
         if RecordCount > 0 then
         begin
            k:=0;
            SetLength(Arr_DNA, RecordCount);
            while not Eof do
            begin
               SetLength(Arr_DNA[k], 2);
               Arr_DNA[k][0] := FieldByName('Cod_hm').AsString;
               Arr_DNA[k][1] := '0';
               inc(k);
               Next;
            end;
         end;
         Active := False;
      end;
   end;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid & Chart 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;

   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      if (GV_sUserId <> '150005') and
         ((RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) OR (RG_Query.ItemIndex = 6)  OR (RG_Query.ItemIndex = 7) OR
         (RG_Query.ItemIndex = 8) or (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 11) or (RG_Query.ItemIndex = 12) OR
         (RG_Query.ItemIndex = 13)) then
         begin
         if (GV_sJisa <> '61') and (GV_sJisa <> '71') THEN sWhere := '  WHERE B.dat_gmgn = ''' + mskDate.Text + ''' '
         else sWhere := '  WHERE B.dat_gmgn >= ''' + mskDate.Text + ''' ' +
                        '  AND B.Dat_gmgn  <= ''' + mskFDate.Text + ''' ';
         end
      else
          begin
          if (GV_sJisa <> '61') and (GV_sJisa <> '71') THEN sWhere := '  WHERE A.dat_bunju = ''' + mskDate.Text + ''' '
          else sWhere :=  '  WHERE A.dat_bunju >= ''' + mskDate.Text + ''' ' +
                          '  AND A.dat_bunju  <= ''' + mskFDate.Text + ''' ';
          end;

      if Trim(mskFrom.Text) <> '' then
      begin
         sWhere := sWhere + ' AND A.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND A.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
      end;

      if Copy(cmbCod_jisa.Items[cmbCod_jisa.ItemIndex],1,2) <> '00' then
      begin
      if UV_sCod_jisa = '15' then
      begin
         if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 6)  OR (RG_Query.ItemIndex = 7) or
            (RG_Query.ItemIndex = 8) or (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 11) or (RG_Query.ItemIndex = 12) or
            (RG_Query.ItemIndex = 13) then
         begin
            if RG_Query.ItemIndex = 3 then
            begin
               if (CB_01.checked = True) then
                 sWhere := sWhere + ' AND (B.cod_jisa <> ''11'' AND B.cod_jisa <> ''12'') '
               else
                 sWhere := sWhere + ' AND B.cod_jisa <> '' '' ';
            end
            else if Trim(cmbCod_jisa.Text) <> '' then
                sWhere := sWhere + ' AND B.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + ''''
         end
         else sWhere := sWhere + ' AND A.cod_bjjs = ''' + UV_sCod_jisa + '''';
      end
      else
      begin
         if cmbCOD_JISA.ItemIndex = 0 then   //15~
         begin
            sWhere := sWhere + ' AND A.cod_bjjs = ''15''';
            sWhere := sWhere + ' AND A.cod_jisa = ''' + UV_sCod_jisa + '''';
         end                                              //15
         else if cmbCOD_JISA.ItemIndex = 1 then  //11, 12, 43, 61, 51, 71
         begin
            if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 6) OR (RG_Query.ItemIndex = 7) or
               (RG_Query.ItemIndex = 8) or (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 11) or (RG_Query.ItemIndex = 12) or
               (RG_Query.ItemIndex = 13 ) then
            begin
               sWhere := sWhere + ' AND B.cod_jisa = ''' + UV_sCod_jisa + '''';
            end
            else sWhere := sWhere + ' AND A.cod_bjjs = ''' + UV_sCod_jisa + '''';
         end
         else
         begin
            if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 6) OR (RG_Query.ItemIndex = 7) or
               (RG_Query.ItemIndex = 8) or (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 11) or (RG_Query.ItemIndex = 12) or
               (RG_Query.ItemIndex = 13) then
            begin
               sWhere := sWhere + ' AND B.cod_jisa = ''' + UV_sCod_jisa + '''';
            end
            else sWhere := sWhere + ' AND A.cod_bjjs = ''' + UV_sCod_jisa + '''';
         end;
      end;
      end;
      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
            sOrder := ' order by A.num_bunju ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp   >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
         sOrder := ' order by B.cod_jisa, B.num_Samp';
      end;

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + ' AND B.cod_dc = ''' + edtDc.Text + '''';

      if RG_Query.ItemIndex = 1 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''U046'' or cod_hm = ''Y005'' or cod_hm = ''Y008'' or cod_hm = ''U017'' or cod_hm = ''U019'' or cod_hm = ''U058'' or cod_hm = ''U035'' or cod_hm = ''U015'' or cod_hm = ''U016''))) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''U046'' or cod_hm = ''Y005'' or cod_hm = ''Y008'' or cod_hm = ''U017'' or cod_hm = ''U019'' or cod_hm = ''U058'' or cod_hm = ''U035'' or cod_hm = ''U015'' or cod_hm = ''U016''))) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''U046'' or cod_hm = ''Y005'' or cod_hm = ''Y008'' or cod_hm = ''U017'' or cod_hm = ''U019'' or cod_hm = ''U058'' or cod_hm = ''U035'' or cod_hm = ''U015'' or cod_hm = ''U016''))) ';
           sWhere := sWhere + ' or (  cod_chuga like ''%U046%''  or cod_chuga like ''%Y005%'' or cod_chuga like ''%U017%''  or cod_chuga like ''%Y008%''  or cod_chuga like ''%U019%'' or cod_chuga like ''%U058%'' or cod_chuga like ''%U035%'' or cod_chuga like ''%U015%'' or cod_chuga like ''%U016%'') ';
           sWhere := sWhere + ' or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
     { else if RG_Query.ItemIndex = 2 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T010'' or cod_hm = ''T011'' ))) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T010'' or cod_hm = ''T011'' ))) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T010'' or cod_hm = ''T011'' ))) ';
           sWhere := sWhere + ' or ( cod_chuga like ''%T010%'' or cod_chuga like ''%T011%'' ) ';
           sWhere := sWhere + ' or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 3 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T007'' or cod_hm = ''T008'' ))) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T007'' or cod_hm = ''T008'' ))) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''T007'' or cod_hm = ''T008'' ))) ';
           sWhere := sWhere + ' or ( cod_chuga like ''%T007%'' or cod_chuga like ''%T008%'' ) ';
           sWhere := sWhere + ' or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 4 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where   cod_hm = ''Y004'' )) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where   cod_hm = ''Y004'' )) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where   cod_hm = ''Y004'')) ';
           sWhere := sWhere + ' or  cod_chuga like ''%Y004%'' ';
           sWhere := sWhere + ' or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 5 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''S094'' or cod_hm = ''S095'' or cod_hm = ''S096'' or cod_hm = ''S097''  ))) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''S094'' or cod_hm = ''S095'' or cod_hm = ''S096'' or cod_hm = ''S097'' ))) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where (  cod_hm = ''S094'' or cod_hm = ''S095'' or cod_hm = ''S096'' or cod_hm = ''S097''))) ';
           sWhere := sWhere + ' or (  cod_chuga like ''%S094%''  or cod_chuga like ''%S095%'' or cod_chuga like ''%S096%''  or cod_chuga like ''%S097%''   )) ';
      end    }
      else if RG_Query.ItemIndex = 2 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where  cod_hm = ''U060'' or cod_hm = ''U061'')) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where  cod_hm = ''U060'' or cod_hm = ''U061'')) ';
           sWhere := sWhere + ' or B.cod_cancer in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where   cod_hm = ''U060'' or cod_hm = ''U061'')) ';
           sWhere := sWhere + ' or cod_chuga like ''%U060%'' or cod_chuga like ''%U061%'' ';
           sWhere := sWhere + ' or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 6) or (RG_Query.ItemIndex = 11) then
      begin
         for k := 0 to length(Arr_DNA)-1 do
         begin
             sLike := sLike + 'cod_chuga like ''%' + Arr_DNA[k][0] + '%'' or ';
         end;
         sWhere := sWhere  + ' And ( ' + copy(sLike,1,length(sLike)-3) + ') ';
      end
      else if RG_Query.ItemIndex = 5 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''E068'' )) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''E068'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''E068'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%E068%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 7 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''RE01'' )) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''RE01'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''RE01'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%RE01%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 8 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''T049'' )) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''T049'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''T049'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%T049%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 9 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where (cod_hm = ''MG01'' or cod_hm = ''MG02'' or cod_hm = ''MG03'' or cod_hm = ''MG04'' or cod_hm = ''MG05'' or cod_hm = ''MG06'' or cod_hm = ''MG07'' or cod_hm = ''MG08'' or cod_hm = ''MG09'' or cod_hm = ''MG10'' or cod_hm = ''MG11''))) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where (cod_hm = ''MG01'' or cod_hm = ''MG02'' or cod_hm = ''MG03'' or cod_hm = ''MG04'' or cod_hm = ''MG05'' or cod_hm = ''MG06'' or cod_hm = ''MG07'' or cod_hm = ''MG08'' or cod_hm = ''MG09'' or cod_hm = ''MG10'' or cod_hm = ''MG11'') )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where (cod_hm = ''MG01'' or cod_hm = ''MG02'' or cod_hm = ''MG03'' or cod_hm = ''MG04'' or cod_hm = ''MG05'' or cod_hm = ''MG06'' or cod_hm = ''MG07'' or cod_hm = ''MG08'' or cod_hm = ''MG09'' or cod_hm = ''MG10'' or cod_hm = ''MG11''))) ';
            sWhere := sWhere + 'or  (cod_chuga like ''%MG01%'' OR cod_chuga like ''%MG02%'' OR cod_chuga like ''%MG03%'' OR cod_chuga like ''%MG04%'' OR cod_chuga like ''%MG05%'' OR cod_chuga like ''%MG06%'' OR cod_chuga like ''%MG07%'' OR cod_chuga like ''%MG08%'' OR cod_chuga like ''%MG09%''  ';
            sWhere := sWhere + 'OR cod_chuga like ''%MG10%'' OR cod_chuga like ''%MG11%'')';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 10 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MB01'' or cod_hm = ''MB02'')) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MB01'' or cod_hm = ''MB02'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MB01'' or cod_hm = ''MB02'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%MB01%'' OR cod_chuga like ''%MB02%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 12 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MH01'' )) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MH01'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''MH01'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%MH01%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end
      else if RG_Query.ItemIndex = 13 then
      begin
            sWhere := sWhere + 'And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''C074'' )) ';
            sWhere := sWhere + 'or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''C074'' )) ';
            sWhere := sWhere + 'or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
            sWhere := sWhere + '(select cod_pf from profile_hm where   cod_hm = ''C074'')) ';
            sWhere := sWhere + 'or  cod_chuga like ''%C074%'' ';
            sWhere := sWhere + 'or ( B.gubn_tkgm = ''1'' or B.gubn_tkgm = ''2''))';
      end;


      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrder);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q01', 'Q', 'N');
         UP_VarMemSet(RecordCount);
         // DataSet의 값을 Variant변수로 이동
         for i := 0 to RecordCount - 1 do
         begin
            chuga := '';   tk_chuga := '';   chuga_hul := '';  chuga_print := '';
            UP_Clear(Index);
////////////////////////////////////
            if (RG_Query.ItemIndex = 0) then
            begin
               UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
               UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
               if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                  UV_vDesc_name[index] := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
               UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' +
                                        Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
               UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
               UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
               UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
               UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
               UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

               iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
               for j := 0 to iRet-1 do
               begin
                  if (copy(vCod_chuga[j],1,2) <> 'JJ') and (copy(vCod_chuga[j],1,3) <> 'DRU') then
                     chuga_hul := chuga_hul + vCod_chuga[j];
               end;

               if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
               begin
                  qryTkgum.Active := False;
                  qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
                  qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryTkgum.Active := True;
                  if qryTkgum.RecordCount > 0 then
                  begin
//                     chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
                     chuga := chuga_hul + '(특검: ';                            //07.01.12 철. 장비항목 뺌.
                     iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_chuga);
                     for Temp := 0 to iRet - 1 do
                     begin
                       qryPf_hm.Active := False;
                       qryPf_hm.ParamByName('cod_pf').AsString := vCod_chuga[Temp];
                       qryPf_hm.Active := True;
                       if qryPf_hm.RecordCount > 0 then
                       begin
                          while not qryPf_hm.Eof do
                          begin
                             if Pos(qryPf_hm.FieldByName('COD_HM').AsString, tk_chuga) = 0 then
                                tk_chuga := tk_chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                             qryPf_hm.next;
                          end;   // while(qryPf_hm) end;
                       end;      //if(qryPf_hm) end;
                     end;        //for(qryTkgum) end;
                  end;           //if(qryTkgum) end;
                  tk_chuga  := tk_chuga + qryTkgum.FieldByName('Cod_chuga').AsString;
                  UV_vCod_chuga[index] := chuga + tk_chuga + ')';
               end
               else
                  UV_vCod_chuga[index]  := chuga_hul;                               //07.01.12 철. 장비항목 뺌.

               UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                       FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                       FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                       FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                       FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
               Inc(index);
            end
////////////////////////////////////
            else if (RG_Query.ItemIndex = 1) or (RG_Query.ItemIndex = 2) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 7) OR (RG_Query.ItemIndex = 8)
                 OR (RG_Query.ItemIndex = 9) or (RG_Query.ItemIndex = 10) or (RG_Query.ItemIndex = 12) OR (RG_Query.ItemIndex = 13)  then
            begin
               //혈액코드 항목 조회
               if (FieldByName('cod_hul').AsString <> '') then
               begin
                  qryPf_hm.Active := False;
                  qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('cod_hul').AsString;
                  qryPf_hm.Active := True;
                  if qryPf_hm.RecordCount > 0 then
                  begin
                     while not qryPf_hm.Eof do
                     begin
                        if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                           chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                        qryPf_hm.next;
                     end;   // while(qryPf_hm) end;
                  end;
               end;
               //종양코드 항목 조회
               if (FieldByName('cod_cancer').AsString <> '') then
               begin
                  qryPf_hm.Active := False;
                  qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('cod_cancer').AsString;
                  qryPf_hm.Active := True;
                  if qryPf_hm.RecordCount > 0 then
                  begin
                     while not qryPf_hm.Eof do
                     begin
                        if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                           chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                        qryPf_hm.next;
                     end;   // while(qryPf_hm) end;
                  end;
               end;
               chuga := chuga + FieldByName('cod_chuga').AsString;

               //특검코드 항목 조회
               if ((FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2')) then
               begin
                  qryTkgum.Active := False;
                  qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
                  qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryTkgum.Active := True;
                  if qryTkgum.RecordCount > 0 then
                  begin
                     tk_chuga := tk_chuga + '(특검: ';                            //07.01.12 철. 장비항목 뺌.
                     iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, vCod_chuga);
                     for Temp := 0 to iRet - 1 do
                     begin
                       qryPf_hm.Active := False;
                       qryPf_hm.ParamByName('cod_pf').AsString := vCod_chuga[Temp];
                       qryPf_hm.Active := True;
                       if qryPf_hm.RecordCount > 0 then
                       begin
                          while not qryPf_hm.Eof do
                          begin
                             if Pos(qryPf_hm.FieldByName('COD_HM').AsString, tk_chuga) = 0 then
                                tk_chuga := tk_chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                             qryPf_hm.next;
                          end;   // while(qryPf_hm) end;
                       end;      //if(qryPf_hm) end;
                     end;        //for(qryTkgum) end;
                  end;           //if(qryTkgum) end;
                  tk_chuga := tk_chuga + qryTkgum.FieldByName('Cod_chuga').AsString;
                  chuga    := chuga + tk_chuga + ')';
                  UV_vCod_chuga[index] := chuga;
               end
               else
                  UV_vCod_chuga[index]  := chuga;


               if (RG_Query.ItemIndex = 1) and
                  ((Pos('Y005', chuga) > 0) or (Pos('Y008', chuga) > 0) or (Pos('U035', chuga) > 0) or
                   (Pos('U017', chuga) > 0) or (Pos('U019', chuga) > 0) or (Pos('U046', chuga) > 0) or (Pos('U058', chuga) > 0) or
                   (Pos('U015', chuga) > 0) or (Pos('U016', chuga) > 0)) then
               begin
                  UV_vNum_bunju[index]  := IntToStr(Index+1) + '/' + FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;


                  if Pos('Y005', chuga) > 0 then
                  chuga_print := chuga_print + 'Y005';
                  if Pos('Y008', chuga) > 0 then
                  chuga_print := chuga_print + 'Y008';
                  if Pos('U017', chuga) > 0 then
                  chuga_print := chuga_print + 'U017';
                  if Pos('U046', chuga) > 0 then
                  chuga_print := chuga_print + 'U046';
                  if Pos('U019', chuga) > 0 then
                  chuga_print := chuga_print + 'U019';
                  if Pos('U058', chuga) > 0 then
                  chuga_print := chuga_print + 'U058';
                  if Pos('U035', chuga) > 0 then
                  chuga_print := chuga_print + 'U035';
                  if Pos('U015', chuga) > 0 then
                  chuga_print := chuga_print + 'U015';
                  if Pos('U016', chuga) > 0 then
                  chuga_print := chuga_print + 'U016';


                  UV_vCod_chuga[index] := chuga_print;
                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
          {     else if (RG_Query.ItemIndex = 2) and
                  ((Pos('T010', chuga) > 0) or (Pos('T011', chuga) > 0)) then
               begin
                  UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
                  for j := 0 to iRet-1 do
                  begin
                     if (copy(vCod_chuga[j],1,2) <> 'JJ') and (copy(vCod_chuga[j],1,3) <> 'DRU') then
                        chuga_hul := chuga_hul + vCod_chuga[j];
                  end;

                  if ((FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2')) then
                  begin
                     chuga := chuga_hul + '(특검: ';                            //07.01.12 철. 장비항목 뺌.
                     UV_vCod_chuga[index]  := chuga + tk_chuga + ')';
                  end
                  else
                     UV_vCod_chuga[index]  := chuga_hul;                               //07.01.12 철. 장비항목 뺌.

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end

               else if (RG_Query.ItemIndex = 3) and
                  ((Pos('T007', chuga) > 0) or (Pos('T008', chuga) > 0)) then
               begin
                  UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;
                  iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);

                  for j := 0 to iRet-1 do
                  begin
                     if (copy(vCod_chuga[j],1,2) <> 'JJ') and (copy(vCod_chuga[j],1,3) <> 'DRU') then
                        chuga_hul := chuga_hul + vCod_chuga[j];
                  end;

                  if ((FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2')) then
                  begin
                     chuga := chuga_hul + '(특검: ';                            //07.01.12 철. 장비항목 뺌.
                     UV_vCod_chuga[index]  := chuga + tk_chuga + ')';
                  end
                  else
                     UV_vCod_chuga[index]  := chuga_hul;                               //07.01.12 철. 장비항목 뺌.

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 4) and (Pos('Y004', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('Y004', chuga) > 0 then
                  chuga_print := chuga_print + 'Y004';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end

               else if (RG_Query.ItemIndex = 5)  then
               begin
                  UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('S094', chuga) > 0 then
                  chuga_print := chuga_print + 'S094';
                  if Pos('S095', chuga) > 0 then
                  chuga_print := chuga_print + 'S095';
                  if Pos('S096', chuga) > 0 then
                  chuga_print := chuga_print + 'S096';
                  if Pos('S097', chuga) > 0 then
                  chuga_print := chuga_print + 'S097';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end  }
               else if (RG_Query.ItemIndex = 2) and
                       ((Pos('U060', chuga) > 0) or (Pos('U061', chuga) > 0))  then
               begin
                  UV_vNum_bunju[index]  := FieldByName('Num_bunju').AsString;
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('U060', chuga) > 0 then
                  chuga_print := chuga_print + 'U060';
                  if Pos('U061', chuga) > 0 then
                  chuga_print := chuga_print + 'U061';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 5) and (Pos('E068', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('E068', chuga) > 0 then
                  chuga_print := chuga_print + 'E068';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 7) and (Pos('RE01', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('RE01', chuga) > 0 then
                  chuga_print := chuga_print + 'RE01';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 8) and (Pos('T049', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('T049', chuga) > 0 then
                  chuga_print := chuga_print + 'T049';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 9)
                   and ((Pos('MG01', chuga) > 0) OR (Pos('MG02', chuga) > 0) OR (Pos('MG03', chuga) > 0) or
                        (Pos('MG04', chuga) > 0) OR (Pos('MG05', chuga) > 0) OR (Pos('MG06', chuga) > 0) or
                        (Pos('MG07', chuga) > 0) OR (Pos('MG08', chuga) > 0) OR (Pos('MG09', chuga) > 0) or
                        (Pos('MG10', chuga) > 0) OR (Pos('MG11', chuga) > 0))  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('MG01', chuga) > 0 then
                  chuga_print := chuga_print + 'MG01';
                  if Pos('MG02', chuga) > 0 then
                  chuga_print := chuga_print + 'MG02';
                  if Pos('MG03', chuga) > 0 then
                  chuga_print := chuga_print + 'MG03';
                  if Pos('MG04', chuga) > 0 then
                  chuga_print := chuga_print + 'MG04';
                  if Pos('MG05', chuga) > 0 then
                  chuga_print := chuga_print + 'MG05';
                  if Pos('MG06', chuga) > 0 then
                  chuga_print := chuga_print + 'MG06';
                  if Pos('MG07', chuga) > 0 then
                  chuga_print := chuga_print + 'MG07';
                  if Pos('MG08', chuga) > 0 then
                  chuga_print := chuga_print + 'MG08';
                  if Pos('MG09', chuga) > 0 then
                  chuga_print := chuga_print + 'MG09';
                  if Pos('MG10', chuga) > 0 then
                  chuga_print := chuga_print + 'MG10';
                  if Pos('MG11', chuga) > 0 then
                  chuga_print := chuga_print + 'MG11';


                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 10)
                   and ((Pos('MB01', chuga) > 0) OR (Pos('MB02', chuga) > 0)) then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('MB01', chuga) > 0 then
                  chuga_print := chuga_print + 'MB01';
                  if Pos('MB02', chuga) > 0 then
                  chuga_print := chuga_print + 'MB02';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 12) and (Pos('MH01', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('MH01', chuga) > 0 then
                  chuga_print := chuga_print + 'MH01';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
               else if (RG_Query.ItemIndex = 13) and (Pos('C074', chuga) > 0)  then
               begin
                  UV_vNum_bunju[index]  := index + 1; //분주번호 대신 No. Count 증가
                  UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
                  UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
                  UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
                  if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                     UV_vDesc_name[index]  := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
                  UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
                  UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
                  UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_Cancer[index] := FieldByName('Cod_Cancer').AsString;

                  if Pos('C074', chuga) > 0 then
                  chuga_print := chuga_print + 'C074';

                  UV_vCod_chuga[index] := chuga_print;

                  UV_vCod_etc[index] := UP_Gumgin_YH(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, FieldByName('Cod_hul').AsString, FieldByName('Cod_jangbi').AsString,
                                          FieldByName('Cod_Cancer').AsString, FieldByName('gubn_hulgum').AsString, FieldByName('gubn_nosin').AsString, FieldByName('gubn_tkgm').AsString,
                                          FieldByName('gubn_adult').AsString, FieldByName('gubn_agab').AsString, FieldByName('gubn_gong').AsString, FieldByName('gubn_life').AsString,
                                          FieldByName('gubn_nsyh').AsInteger, FieldByName('gubn_tkyh').AsInteger, FieldByName('gubn_adyh').AsInteger, FieldByName('gubn_agyh').AsInteger,
                                          FieldByName('gubn_goyh').AsInteger, FieldByName('gubn_lfyh').AsInteger);
                  Inc(index);
               end
            end
////////////////////////////////////
            else if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 6) or (RG_Query.ItemIndex = 11) then
            begin
               UV_vNum_bunju[index]  := index + 1;// 20140415 곽윤설 - 분주번호 대신 No. Count 증가
               UV_vNum_samp[index]   := FieldByName('Num_samp').AsString;
               UV_vDesc_name[index]  := FieldByName('Desc_name').AsString;
               if (Ck_dept.Checked) and (FieldByName('desc_dept').AsString <> '') then
                  UV_vDesc_name[index] := UV_vDesc_name[index] + '-' + Copy(FieldByName('desc_dept').AsString,1,10);
               UV_vNum_jumin[index]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' +
                                        Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
               UV_vDat_gmgn[index]   := FieldByName('Dat_gmgn').AsString;
               UV_vCod_jisa[index]   := FieldByName('Cod_jisa').AsString;
               chuga := FieldByName('cod_chuga').AsString;

               for k:=0 to length(Arr_DNA)-1 do
               begin
                 if Pos(Arr_DNA[k][0], chuga) > 0 then
                 begin
                    chuga_print := chuga_print + Arr_DNA[k][0];
                    Arr_DNA[k][1] := IntToStr(StrToInt(Arr_DNA[k][1]) + 1);
                 end;
               end;
               UV_vCod_chuga[index] := chuga_print;

               Inc(index);
            end;
            Next;
         end;
         // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
         Active := False;
         if (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 6) or (RG_Query.ItemIndex = 11) then
         begin
            flag_01 := 0; flag_02 := 0; flag_03 := 0; flag_04 := 0; flag_05 := 0; flag_06 := 0;

            UV_vDesc_dc[Index]   := '';
            UV_vCod_jangbi[Index]:= '';
            UV_vCod_hul[Index]   := '';
            UV_vCod_Cancer[Index]:= '';
            UV_vCod_etc[Index]   := '';
            UV_vNum_samp[Index]  := '';
            UV_vDesc_name[Index] := '';
            UV_vNum_jumin[Index] := '';
            UV_vDat_gmgn[Index]  := '';
            UV_vCod_jisa[Index]  := '';
            UV_vCod_chuga[Index] := '';
            UV_vNum_bunju[Index] := '총합계';

            for k:=0 to length(Arr_DNA)-1 do
            begin
               if Arr_DNA[k][1] <> '0' then
               begin
                  if (flag_06 = 4) then //or (k = length(Arr_DNA)-1)
                  begin
                     flag_01 := 0; flag_02 := 0; flag_03 := 0; flag_04 := 0; flag_05 := 0; flag_06 := 0;
                     Inc(index);
                     UP_VarMemSet(index);
                  end;

                  if flag_01 = 0 then
                  begin
                     UV_vNum_samp[Index]  := Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ')';
                     inc(flag_01);
                  end
                  else if flag_02 = 0 then
                  begin
                     UV_vDesc_name[Index] := Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ')';
                     inc(flag_02);
                  end
                  else if flag_03 = 0 then
                  begin
                     UV_vNum_jumin[Index] := Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ')';
                     inc(flag_03);
                  end
                  else if flag_04 < 2 then
                  begin
                     UV_vDat_gmgn[Index]  := UV_vDat_gmgn[Index] +
                                             Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ') ';
                     inc(flag_04);
                  end
                  else if flag_05 = 0 then
                  begin
                     UV_vCod_jisa[Index]  := Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ')';
                     inc(flag_05);
                  end
                  else if flag_06 < 4 then
                  begin
                     UV_vCod_chuga[Index] := UV_vCod_chuga[Index] +
                                             Arr_DNA[k][0] + '(' + Arr_DNA[k][1] + ') ';
                     inc(flag_06);
                  end;
               end;
            end;
            Inc(index);
         end;

      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;
   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD2Q01.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 없을경우의 처리
   if NewRow = 0 then exit;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // Grid Focus
   grdMaster.SetFocus;
end;

procedure TfrmLD2Q01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   //20140526
   if  (GV_sUserId <> '150005') and
       ((RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 6) or (RG_Query.ItemIndex = 7) OR (RG_Query.ItemIndex = 8) OR (RG_Query.ItemIndex = 9) OR (RG_Query.ItemIndex = 10) OR (RG_Query.ItemIndex = 11) OR (RG_Query.ItemIndex = 12) OR (RG_Query.ItemIndex = 13)) then
   begin
       Label2.Caption := '검 진 일 자 : ';
       if RG_Query.ItemIndex = 3 then
       begin
          Cmb_gubn.Text := '샘플번호';
          if UV_sCod_jisa = '15' then CB_01.Visible := True
          else CB_01.Visible := False;
       end;
   end
   else
   begin
       Label2.Caption := '분 주 일 자 : ';
       CB_01.Visible := False;
       if RG_Query.ItemIndex = 1 then
       begin
          Cmb_gubn.Text := '분주번호';
       end
   end;

   if Cmb_gubn.Text = '분주번호' then
   begin
       MEdt_SampS.Enabled := False;
       MEdt_SampE.Enabled := False;
       MEdt_SampS.Text := '';
       MEdt_SampE.Text := '';
       mskFrom.Enabled := True;
       mskTo.Enabled := True;
   end
   else
   begin
       MEdt_SampS.Enabled := True;
       MEdt_SampE.Enabled := True;
       mskFrom.Enabled := False;
       mskTo.Enabled := False;
       mskFrom.Text := '';
       mskTo.Text := '';
   end;

   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtD_dc.Text := sName;
      end;
   end
   else 
   if Sender = btnDate then GF_CalendarClick(mskDate);

   if Sender = btnDate_To then GF_CalendarClick(mskFDate);

end;

procedure TfrmLD2Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate then UP_Click(btnDate)
   else if Sender = mskFDate then UP_Click(btnDate_To)
   else if Sender = edtDc   then UP_Click(btnDC);
end;

procedure TfrmLD2Q01.grdMasterHeadingClick(Sender: TObject;
  DataCol: Integer);
var iCnt : Integer;
    sOrder : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if grdMaster.Rows = 0 then exit;

   iCnt := grdMaster.Rows;

   if grdMaster.Col[DataCol].SortPicture = spUp then
   begin
      grdMaster.Col[DataCol].SortPicture := spDown;
      sOrder := 'A';
   end
   else
   begin
      grdMaster.Col[DataCol].SortPicture := spUp;
      sOrder := 'D';
   end;

   case DataCol of
      1 : UP_QuickSort(sOrder, UV_vNum_bunju,  0, iCnt-1);
      2 : UP_QuickSort(sOrder, UV_vNum_samp,   0, iCnt-1);
      3 : UP_QuickSort(sOrder, UV_vDesc_name,  0, iCnt-1);
      4 : UP_QuickSort(sOrder, UV_vNum_jumin,  0, iCnt-1);
      5 : UP_QuickSort(sOrder, UV_vDat_gmgn,   0, iCnt-1);
      6 : UP_QuickSort(sOrder, UV_vDesc_dc,    0, iCnt-1);
      7 : UP_QuickSort(sOrder, UV_vCod_jisa,   0, iCnt-1);
      8 : UP_QuickSort(sOrder, UV_vCod_jangbi, 0, iCnt-1);
      9 : UP_QuickSort(sOrder, UV_vCod_hul,    0, iCnt-1);
     10 : UP_QuickSort(sOrder, UV_vCod_Cancer, 0, iCnt-1);
     11 : UP_QuickSort(sOrder, UV_vCod_chuga,  0, iCnt-1);
     12 : UP_QuickSort(sOrder, UV_vCod_etc,    0, iCnt-1);

      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;
end;

procedure TfrmLD2Q01.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   if (RG_Query.ItemIndex = 1) or (RG_Query.ItemIndex = 2) or
      (RG_Query.ItemIndex = 3) or (RG_Query.ItemIndex = 4) or
      (RG_Query.ItemIndex = 5) or (RG_Query.ItemIndex = 6) or
      (RG_Query.ItemIndex = 7) or (RG_Query.ItemIndex = 8) or
      (RG_Query.ItemIndex = 10) OR (RG_Query.ItemIndex = 11) or
      (RG_Query.ItemIndex = 12) OR (RG_Query.ItemIndex = 13) then //곽윤설 20140415 - 조회구분 3,4은 LD2Q012 양식으로 출력
       begin
       frmLD2Q012 := TfrmLD2Q012.Create(Self);
       if RBtn_preveiw.Checked = True  then  frmLD2Q012.QuickRep.Preview
       else                                  frmLD2Q012.QuickRep.Print;
       frmLD2Q012.free;
   end
   else
      begin
      frmLD2Q011 := TfrmLD2Q011.Create(Self);
      if RBtn_preveiw.Checked = True  then  frmLD2Q011.QuickRep.Preview
      else                                  frmLD2Q011.QuickRep.Print;
      frmLD2Q011.free;
      end;

end;

function  TfrmLD2Q01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         Result := FieldByName('desc_yhname').AsString;
      end;
      Active := False;
   end;
end;

function  TfrmLD2Q01.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_id').AsString      := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_yhname').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD2Q01.UP_Clear(Temp : Integer);
begin
    UV_vNum_bunju[Temp]  := ''; UV_vDesc_name[Temp] := ''; UV_vNum_jumin[Temp]  := '';
    UV_vDat_gmgn[Temp]   := ''; UV_vDesc_dc[Temp]   := ''; UV_vCod_jisa[Temp]   := '';
    UV_vCod_jangbi[Temp] := ''; UV_vCod_hul[Temp]   := ''; UV_vCod_Cancer[Temp] := '';
    UV_vCod_chuga[Temp]  := ''; UV_vCod_etc[Temp]   := ''; UV_vNum_samp[Temp]   := '';
end;

function  TfrmLD2Q01.UP_Gumgin_YH(sJisa, sID, sHul, sJangbi, sCancer, sHulgum, sNosin, sTkgm, sAdult, sAgab, sGong, sLife : String;
                                  iNosin, iTkgm, iAdult, iAgab, iGong, iLife : Integer) : String;
begin
   Result := '';
   // 노신유형Display
   if sNosin = '1' then
   begin
      if (sHulgum = '3') and (copy(sJangbi,1,2) = '') and
         (sHul = '') and (sCancer = '') then
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', iNosin) + ':(CBC), '
      else
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', iNosin) + ', ';
   end
   else if sNosin = '2' then
      Result := Result + '노신재검' + ', ';

   //특검유형Display
   if sTkgm = '1' then
      Result := Result + '특검' + ', '
   else if sTkgm = '2' then
      Result := Result + '특검재검';

   //성인병유형Display
   if sAdult = '1' then
   begin
      if (sHulgum = '3') and (copy(sJangbi,1,2) = '') and
         (sHul = '') and (sCancer = '') then
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', iAdult) + ':(CBC), '
      else
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', iAdult) + ', ';
   end
   else if sAdult = '2' then
      Result := Result + UF_Jae_Hangmok(sJisa, sID, '1', iAdult) + ', ';

   //간염유형Display
   if sAgab = '1' then
      Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', iAgab)
   else if sAgab = '2' then
      Result := Result + UF_Jae_Hangmok(sJisa, sID, '1', iAgab);

   //공교의보유형Display
   if sGong = '1' then
      Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', iGong)
   else if sGong = '2' then
      Result := Result + UF_Jae_Hangmok(sJisa, sID, '1', iGong);

   // 생애전환기유형Display
   if sLife = '1' then
   begin
      if (sHulgum = '3') and (copy(sJangbi,1,2) = '') and
         (sHul = '') and (sCancer = '') then
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', iLife) + ':(CBC), '
      else
         Result := Result + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', iLife) + ', ';
   end
   else if sLife = '2' then
      Result := Result + '생애전환기재검' + ', ';

end;

end.
