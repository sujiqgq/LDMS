//==============================================================================
// 프로그램 설명    : 분주 list 현황
// 시스템           : 분석관리프로그램
// 수정일자         : 2006.05.17
// [수정자]수정내용 : 추가..
// 참고사항         :
//==============================================================================
// 수정일자         : 2014.04.21
// [수정자]수정내용 : [곽윤설], [센터]+[연구소+센터]RIA - 혈검 1일때도 조회가능
// 참고사항         : (요청자 : 부산-유희짐)
//==============================================================================
// 수정일자         : 2014.04.29
// [수정자]수정내용 : [곽윤설], 1. 4-[연구소]+[연구소+센터]RIA 생성
//                              2. 3-[연구소]+[연구소+센터](위탁)RIA로 리스트 조회시 S016 검사건 조회
// 참고사항         : (요청자 : 부산-유희짐)
//                    HCV Ab(S016) 검사는 EIA법으로 하나 검체를 챙기는 팀이 RIA검사파트여서 바코드가 RIA로 나오기도 함.
//==============================================================================
// 수정일자  : 2014.05.21
// 수 정 자  : 곽윤설                                                                                            
// 수정내용  : S019, C044, C049, E001, E002, E003, E015, E040, T002, T006,
//             T007, T009, T011, T037 모두 EIA이나, RIA로 조회 가능하도록 (TT01, TT02, S016과 같은이유..)
// 참고사항  : [요청자 : 부산-유희짐]
//==============================================================================
// 수정일자  : 2015.04.20
// 수 정 자  : 곽윤설
// 수정내용  : 노신,성인병 재검 & C032검사 있을때 센터자체 공단 검사
//==============================================================================
// 수정일자  : 2016.02.17
// 수 정 자  : 박수지
// 수정내용  : 한의 본원종합검사3파트 1600068
//==============================================================================
unit LD5Q17;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, ORSingle;

type
  TfrmLD5Q17 = class(TfrmSingle)
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
    Label4: TLabel;
    cmbHulgum: TComboBox;
    Label5: TLabel;
    Com_BD: TComboBox;
    RBtn_Preview: TRadioButton;
    RBtn_Print: TRadioButton;
    qryJHangmokList: TQuery;
    qryHmCheck: TQuery;
    qryHmCheck2: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
  private
    { Private declarations }

    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_vGubn_hulgum, UV_vCod_bjjs, UV_vNum_bunju,  UV_vDesc_name, UV_vNum_jumin,
    UV_vNum_samp,    UV_vGubn_sex, UV_vCod_jangbi, UV_vCod_hul,   UV_vCod_Cancer,
    UV_vCod_chuga,   UV_vCod_etc,  UV_vDesc_dc : Variant;

    UV_sCod_jisa : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);

    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_No_Hangmok_Hm(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;

  public
    { Public declarations }
  end;

var
  frmLD5Q17: TfrmLD5Q17;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q071;

{$R *.DFM}


procedure TfrmLD5Q17.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
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
            T := UV_vGubn_hulgum[Lo]; UV_vGubn_hulgum[Lo] := UV_vGubn_hulgum[Hi]; UV_vGubn_hulgum[Hi] := T;
            T := UV_vCod_bjjs[Lo];    UV_vCod_bjjs[Lo]    := UV_vCod_bjjs[Hi];    UV_vCod_bjjs[Hi]    := T;
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];   UV_vNum_bunju[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];   UV_vDesc_name[Hi]   := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];   UV_vNum_jumin[Hi]   := T;
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];     UV_vDesc_dc[Hi]     := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]    := UV_vNum_samp[Hi];    UV_vNum_samp[Hi]    := T;
            T := UV_vGubn_sex[Lo];    UV_vGubn_sex[Lo]    := UV_vGubn_sex[Hi];    UV_vGubn_sex[Hi]    := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo]  := UV_vCod_jangbi[Hi];  UV_vCod_jangbi[Hi]  := T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]     := UV_vCod_hul[Hi];     UV_vCod_hul[Hi]     := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo]  := UV_vCod_Cancer[Hi];  UV_vCod_Cancer[Hi]  := T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]   := UV_vCod_chuga[Hi];   UV_vCod_chuga[Hi]   := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]     := UV_vCod_etc[Hi];     UV_vCod_etc[Hi]     := T;
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
            T := UV_vGubn_hulgum[Lo]; UV_vGubn_hulgum[Lo] := UV_vGubn_hulgum[Hi]; UV_vGubn_hulgum[Hi] := T;
            T := UV_vCod_bjjs[Lo];    UV_vCod_bjjs[Lo]    := UV_vCod_bjjs[Hi];    UV_vCod_bjjs[Hi]    := T;
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];   UV_vNum_bunju[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];   UV_vDesc_name[Hi]   := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];   UV_vNum_jumin[Hi]   := T;
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];     UV_vDesc_dc[Hi]     := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]    := UV_vNum_samp[Hi];    UV_vNum_samp[Hi]    := T;
            T := UV_vGubn_sex[Lo];    UV_vGubn_sex[Lo]    := UV_vGubn_sex[Hi];    UV_vGubn_sex[Hi]    := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo]  := UV_vCod_jangbi[Hi];  UV_vCod_jangbi[Hi]  := T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]     := UV_vCod_hul[Hi];     UV_vCod_hul[Hi]     := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo]  := UV_vCod_Cancer[Hi];  UV_vCod_Cancer[Hi]  := T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]   := UV_vCod_chuga[Hi];   UV_vCod_chuga[Hi]   := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]     := UV_vCod_etc[Hi];     UV_vCod_etc[Hi]     := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD5Q17.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vGubn_hulgum := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_sex    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_etc     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vGubn_hulgum, iLength);
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);
      VarArrayRedim(UV_vGubn_sex,    iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_Cancer,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vCod_etc,     iLength);
      VarArrayRedim(UV_vDesc_dc,     iLength);
   end;
end;

function TfrmLD5Q17.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD5Q17.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_VarMemSet(0);

   Com_BD.ItemIndex := GP_SawonCheck(Com_BD, GV_sUserId);

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   if UV_sCod_jisa = '15' then
   begin
      Label1.Caption := '지 사:';
      GP_ComboJisa(cmbCOD_JISA);
   end
   else
   begin
      Label1.Caption := '분주소:';
      cmbCod_jisa.Items.Add('연 구 소');
      cmbCod_jisa.Items.Add('지    사');
      cmbCod_jisa.ItemIndex := 1;
   end;


   cmbHulgum.ItemIndex := 0;
   Com_BD.ItemIndex  := 0;
end;

procedure TfrmLD5Q17.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vGubn_hulgum[DataRow-1];
      2 : Value := UV_vCod_bjjs[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      6 : Value := UV_vDesc_dc[DataRow-1];
      7 : Value := UV_vNum_samp[DataRow-1];
      8 : Value := UV_vGubn_sex[DataRow-1];
      9 : Value := UV_vCod_jangbi[DataRow-1];
     10 : Value := UV_vCod_hul[DataRow-1];
     11 : Value := UV_vCod_Cancer[DataRow-1];
     12 : Value := UV_vCod_chuga[DataRow-1];
     13 : Value := UV_vCod_etc[DataRow-1];
   end;
end;

procedure TfrmLD5Q17.btnQueryClick(Sender: TObject);
var index, iRet, temp, iAge : Integer;
    sSelect, sWhere, sOrder, yh_name, chuga, sMan,sEtc_Hangmok_hm,sTk_Hangmok_Pf,sTk_Hangmok_hm ,sSelect1,sWhere1,sWhere2,sOderby1,sTotal_HangmokList: String;
    vCod_chuga : Variant;
     i : Integer;
    vCheck_01, vCheck_08, vCheck_01_01, vCheck_08_01, vCheck_01_02,vCheck_04,vCheck_01_03, vCheck_01_jaegum, chk_01, chk_02 : Boolean;
begin
   inherited;
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid & Chart 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sSelect := ''; sWhere := ''; sOrder := '';
   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;
      if (cmbHulgum.ItemIndex = 0) or (cmbHulgum.ItemIndex = 1) or (cmbHulgum.ItemIndex = 2) then
      begin
         sSelect := ' SELECT G.dat_gmgn,  G.desc_name, G.num_jumin, G.num_id, G.cod_jangbi, G.num_samp, G.cod_jisa, ' +
                    ' G.cod_hul, G.cod_cancer, G.cod_chuga, G.gubn_nosin, G.gubn_nsyh, G.gubn_hulgum, ' +
                    ' G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_agab, G.gubn_agyh, ' +
                    ' G.gubn_gong, G.gubn_goyh, G.gubn_life, G.gubn_lfyh, G.gubn_tkgm, B.dat_bunju, B.num_bunju, B.cod_jisa, B.cod_bjjs, D.desc_dc '+
                    ' FROM gumgin G LEFT OUTER JOIN BUNJU B ' +
                    ' ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ' +
                    ' LEFT OUTER JOIN danche D ON G.cod_dc = D.cod_dc';

         sWhere := ' WHERE G.cod_jisa = ''' + UV_sCod_jisa + '''';
         sWhere := sWhere + ' AND G.dat_gmgn = ''' + mskDate.Text + '''';
         case cmbHulgum.ItemIndex of
            0 : sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';
            1 : sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';
            2 : sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';
         end;

         if Trim(Copy(Com_BD.Text, 1, 2)) <> '0' then
            sWhere := sWhere + ' AND G.gubn_jinch = ''' + Trim(Copy(Com_BD.Text, 1, 2)) + '''';
         if Trim(mskFrom.Text) <> '' then
         begin
            sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
            sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
         end;

         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + ' AND B.cod_dc = ''' + edtDc.Text + '''';
      end;

      sOrder := ' order by G.cod_jisa, G.dat_gmgn, CAST(G.num_samp AS INT) ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      if cmbHulgum.ItemIndex = 0 then
      begin
         if RecordCount > 0 then
         begin
           UP_VarMemSet(RecordCount-1);
           // DataSet의 값을 Variant변수로 이동
           while not Eof do
           begin

             with qryHmCheck do
             begin
                qryHmCheck.Active := False;
                Chk_01 := false;
                qryHmCheck.Sql.Clear;
                qryHmCheck.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryBunju.FieldByName('cod_jisa').AsString + ''','''
                                                                           + qryBunju.FieldByName('num_id').AsString   + ''','''
                                                                           + qryBunju.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                           + 'JOIN hangmok H ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
                qryHmCheck.Sql.Text := Sql.Text + ' WHERE (gubn_part =''03'') and (TF_Get_HangmokList.Cod_hm IN (''U001'', ''U002'', ''U003'', ''U004'', ''U005'', ''U006'', ''U007'', ''U008'', ''U009'', ''U010'',''U011'',''U012'',''U013''))';
                qryHmCheck.Active := True;
             end;
             if (qryHmCheck.RecordCount > 0) then Chk_01 := true;


             if chk_01 = True then
             begin
                yh_name := '';   chuga := '';
                GP_JuminToAge(FieldByName('Num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

                     if FieldByName('Gubn_hulgum').AsString = '1' then
                          UV_vGubn_hulgum[index] := '연구소'
                     else if FieldByName('Gubn_hulgum').AsString = '2' then
                          UV_vGubn_hulgum[index] := '지사'
                     else if FieldByName('Gubn_hulgum').AsString = '3' then
                     begin
                          if (FieldByName('gubn_nosin').AsString = '1') or
                             (FieldByName('gubn_nosin').AsString = '2') or
                             (FieldByName('gubn_adult').AsString = '1') or
                             (FieldByName('gubn_adult').AsString = '2') or
                             (FieldByName('gubn_life').AsString  = '1') or
                             (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '연구소+지사(공)'
                          else                                               UV_vGubn_hulgum[index] := '연구소+지사';
                     end;
                     UV_vCod_bjjs[index]    := FieldByName('Cod_bjjs').AsString;
                     UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
                     UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
                     UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
                     UV_vDesc_dc[index]     := FieldByName('Desc_dc').AsString;
                     UV_vNum_samp[index]    := FieldByName('Num_samp').AsString;
                     UV_vGubn_sex[index]    := sMan;
                     UV_vCod_jangbi[index]  := FieldByName('Cod_jangbi').AsString;
                     UV_vCod_hul[index]     := FieldByName('Cod_hul').AsString;
                     UV_vCod_Cancer[index]  := FieldByName('Cod_Cancer').AsString;

                     if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
                     begin
                     chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
                     qryTkgum.Active := False;
                     qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                     qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
                     qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                     qryTkgum.Active := True;
                     if qryTkgum.RecordCount > 0 then
                     begin
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
                                         if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                                         chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                                         qryPf_hm.next;
                                    end;   // while(qryPf_hm) end;
                               end;      //if(qryPf_hm) end;
                          end;        //for(qryTkgum) end;
                     end;           //if(qryTkgum) end;
                     UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
                     end
                     else UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;

                     // 노신유형Display
                     if FieldByName('gubn_nosin').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
                     else if FieldByName('gubn_nosin').AsString = '2' then
                       yh_name := yh_name + '노신재검';

                     //특검유형Display
                     if FieldByName('gubn_tkgm').AsString = '1' then
                       yh_name := yh_name + '특검'
                     else if FieldByName('gubn_tkgm').AsString = '2' then
                       yh_name := yh_name + '특검재검';

                     //성인병유형Display
                     if FieldByName('gubn_adult').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
                     else if FieldByName('gubn_adult').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger) + ', ';

                     //간염유형Display
                     if FieldByName('gubn_agab').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
                     else if FieldByName('gubn_agab').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger) + ', ';

                     //공교의보유형Display
                     if FieldByName('gubn_gong').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
                     else if FieldByName('gubn_gong').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger) + ', ';

                     //생애전환기유형Display
                     if FieldByName('gubn_life').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger)
                     else if FieldByName('gubn_life').AsString = '2' then
                       yh_name := yh_name + '생애재검';

                     UV_vCod_etc[index] := yh_name;
                     Inc(index);
             end;
             Next;

           end;
           // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
           Active := False;
         end;
      end
      else if cmbHulgum.ItemIndex = 1 then
      begin
         if RecordCount > 0 then
         begin
           UP_VarMemSet(RecordCount-1);
           // DataSet의 값을 Variant변수로 이동
           while not Eof do
           begin

             with qryHmCheck do
             begin
                qryHmCheck.Active := False;
                Chk_01 := false;
                qryHmCheck.Sql.Clear;
                qryHmCheck.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryBunju.FieldByName('cod_jisa').AsString + ''','''
                                                                           + qryBunju.FieldByName('num_id').AsString   + ''','''
                                                                           + qryBunju.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                           + 'JOIN hangmok H ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
                qryHmCheck.Sql.Text := Sql.Text + ' WHERE (gubn_part =''02'') and (TF_Get_HangmokList.Cod_hm IN (''H001'', ''H002'', ''H003'', ''H004'', ''H005'', ''H006'', ''H007'', ''H008'', ''H009'', ''H010''';
                qryHmCheck.Sql.Text := Sql.Text + ' ,''H011'',''H012'',''H013'',''H014'',''H015'',''H016'',''H017'',''H018'',''H019'',''H020'',''H021'',''H022'',''H023'',''H024'',''H025'',''H026'',''H035'',''H036''))';
                qryHmCheck.Active := True;
             end;
             if (qryHmCheck.RecordCount > 0) then Chk_01 := true;


             if chk_01 = True then
             begin
                yh_name := '';   chuga := '';
                GP_JuminToAge(FieldByName('Num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

                     if FieldByName('Gubn_hulgum').AsString = '1' then
                          UV_vGubn_hulgum[index] := '연구소'
                     else if FieldByName('Gubn_hulgum').AsString = '2' then
                          UV_vGubn_hulgum[index] := '지사'
                     else if FieldByName('Gubn_hulgum').AsString = '3' then
                     begin
                          if (FieldByName('gubn_nosin').AsString = '1') or
                             (FieldByName('gubn_nosin').AsString = '2') or
                             (FieldByName('gubn_adult').AsString = '1') or
                             (FieldByName('gubn_adult').AsString = '2') or
                             (FieldByName('gubn_life').AsString  = '1') or
                             (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '연구소+지사(공)'
                          else                                               UV_vGubn_hulgum[index] := '연구소+지사';
                     end;
                     UV_vCod_bjjs[index]    := FieldByName('Cod_bjjs').AsString;
                     UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
                     UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
                     UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
                     UV_vDesc_dc[index]     := FieldByName('Desc_dc').AsString;
                     UV_vNum_samp[index]    := FieldByName('Num_samp').AsString;
                     UV_vGubn_sex[index]    := sMan;
                     UV_vCod_jangbi[index]  := FieldByName('Cod_jangbi').AsString;
                     UV_vCod_hul[index]     := FieldByName('Cod_hul').AsString;
                     UV_vCod_Cancer[index]  := FieldByName('Cod_Cancer').AsString;

                     if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
                     begin
                     chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
                     qryTkgum.Active := False;
                     qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                     qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
                     qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                     qryTkgum.Active := True;
                     if qryTkgum.RecordCount > 0 then
                     begin
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
                                         if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                                         chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                                         qryPf_hm.next;
                                    end;   // while(qryPf_hm) end;
                               end;      //if(qryPf_hm) end;
                          end;        //for(qryTkgum) end;
                     end;           //if(qryTkgum) end;
                     UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
                     end
                     else UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;

                     // 노신유형Display
                     if FieldByName('gubn_nosin').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
                     else if FieldByName('gubn_nosin').AsString = '2' then
                       yh_name := yh_name + '노신재검';

                     //특검유형Display
                     if FieldByName('gubn_tkgm').AsString = '1' then
                       yh_name := yh_name + '특검'
                     else if FieldByName('gubn_tkgm').AsString = '2' then
                       yh_name := yh_name + '특검재검';

                     //성인병유형Display
                     if FieldByName('gubn_adult').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
                     else if FieldByName('gubn_adult').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger) + ', ';

                     //간염유형Display
                     if FieldByName('gubn_agab').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
                     else if FieldByName('gubn_agab').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger) + ', ';

                     //공교의보유형Display
                     if FieldByName('gubn_gong').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
                     else if FieldByName('gubn_gong').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger) + ', ';

                     //생애전환기유형Display
                     if FieldByName('gubn_life').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger)
                     else if FieldByName('gubn_life').AsString = '2' then
                       yh_name := yh_name + '생애재검';

                     UV_vCod_etc[index] := yh_name;
                     Inc(index);
             end;
             Next;

           end;
           // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
           Active := False;
         end;
      end
      else if cmbHulgum.ItemIndex = 2 then
      begin
         if RecordCount > 0 then
         begin
           UP_VarMemSet(RecordCount-1);
           // DataSet의 값을 Variant변수로 이동
           while not Eof do
           begin

             with qryHmCheck do
             begin
                qryHmCheck.Active := False;
                Chk_01 := false;
                qryHmCheck.Sql.Clear;
                qryHmCheck.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryBunju.FieldByName('cod_jisa').AsString + ''','''
                                                                           + qryBunju.FieldByName('num_id').AsString   + ''','''
                                                                           + qryBunju.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                           + 'JOIN hangmok H ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
                qryHmCheck.Sql.Text := Sql.Text + ' WHERE (gubn_part =''01'') and (TF_Get_HangmokList.Cod_hm IN (''C009'', ''C010'', ''C011'', ''C025'', ''C026'', ''C027'', ''C028'', ''C032'', ''C042'', ''C074''))';
                qryHmCheck.Active := True;
             end;
             if (qryHmCheck.RecordCount > 0) then Chk_01 := true;


             if chk_01 = True then
             begin
                yh_name := '';   chuga := '';
                GP_JuminToAge(FieldByName('Num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

                     if FieldByName('Gubn_hulgum').AsString = '1' then
                          UV_vGubn_hulgum[index] := '연구소'
                     else if FieldByName('Gubn_hulgum').AsString = '2' then
                          UV_vGubn_hulgum[index] := '지사'
                     else if FieldByName('Gubn_hulgum').AsString = '3' then
                     begin
                          if (FieldByName('gubn_nosin').AsString = '1') or
                             (FieldByName('gubn_nosin').AsString = '2') or
                             (FieldByName('gubn_adult').AsString = '1') or
                             (FieldByName('gubn_adult').AsString = '2') or
                             (FieldByName('gubn_life').AsString  = '1') or
                             (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '연구소+지사(공)'
                          else                                               UV_vGubn_hulgum[index] := '연구소+지사';
                     end;
                     UV_vCod_bjjs[index]    := FieldByName('Cod_bjjs').AsString;
                     UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
                     UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
                     UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
                     UV_vDesc_dc[index]     := FieldByName('Desc_dc').AsString;
                     UV_vNum_samp[index]    := FieldByName('Num_samp').AsString;
                     UV_vGubn_sex[index]    := sMan;
                     UV_vCod_jangbi[index]  := FieldByName('Cod_jangbi').AsString;
                     UV_vCod_hul[index]     := FieldByName('Cod_hul').AsString;
                     UV_vCod_Cancer[index]  := FieldByName('Cod_Cancer').AsString;

                     if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
                     begin
                     chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
                     qryTkgum.Active := False;
                     qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                     qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
                     qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                     qryTkgum.Active := True;
                     if qryTkgum.RecordCount > 0 then
                     begin
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
                                         if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                                         chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                                         qryPf_hm.next;
                                    end;   // while(qryPf_hm) end;
                               end;      //if(qryPf_hm) end;
                          end;        //for(qryTkgum) end;
                     end;           //if(qryTkgum) end;
                     UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
                     end
                     else UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;

                     // 노신유형Display
                     if FieldByName('gubn_nosin').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
                     else if FieldByName('gubn_nosin').AsString = '2' then
                       yh_name := yh_name + '노신재검';

                     //특검유형Display
                     if FieldByName('gubn_tkgm').AsString = '1' then
                       yh_name := yh_name + '특검'
                     else if FieldByName('gubn_tkgm').AsString = '2' then
                       yh_name := yh_name + '특검재검';

                     //성인병유형Display
                     if FieldByName('gubn_adult').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
                     else if FieldByName('gubn_adult').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger) + ', ';

                     //간염유형Display
                     if FieldByName('gubn_agab').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
                     else if FieldByName('gubn_agab').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger) + ', ';

                     //공교의보유형Display
                     if FieldByName('gubn_gong').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
                     else if FieldByName('gubn_gong').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger) + ', ';

                     //생애전환기유형Display
                     if FieldByName('gubn_life').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger)
                     else if FieldByName('gubn_life').AsString = '2' then
                       yh_name := yh_name + '생애재검';

                     UV_vCod_etc[index] := yh_name;
                     Inc(index);
             end;
             Next;

           end;
           // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
           Active := False;
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

procedure TfrmLD5Q17.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD5Q17.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

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
end;

procedure TfrmLD5Q17.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate then UP_Click(btnDate)
   else if Sender = edtDc   then UP_Click(btnDC);
end;

procedure TfrmLD5Q17.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD5Q071 := TfrmLD5Q071.Create(Self);

   if RBtn_Print.Checked then frmLD5Q071.QuickRep.Print
   else                       frmLD5Q071.QuickRep.Preview;

end;
function  TfrmLD5Q17.UF_No_Hangmok_Hm(sDate, sGubun : String; iYh : Integer): String;
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
         if FieldByName('desc_joo').AsString <> '' then
            Result :=  Result + FieldByName('desc_joo').AsString;
      end;
      Active := False;
   end;
end;

function  TfrmLD5Q17.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD5Q17.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q17.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtD_dc.Text := sName;
   end;

end;

procedure TfrmLD5Q17.grdMasterHeadingClick(Sender: TObject;
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
      1 : UP_QuickSort(sOrder, UV_vGubn_hulgum, 0, iCnt-1);
      2 : UP_QuickSort(sOrder, UV_vCod_bjjs,    0, iCnt-1);
      3 : UP_QuickSort(sOrder, UV_vNum_bunju,   0, iCnt-1);
      4 : UP_QuickSort(sOrder, UV_vDesc_name,   0, iCnt-1);
      5 : UP_QuickSort(sOrder, UV_vNum_jumin,   0, iCnt-1);
      6 : UP_QuickSort(sOrder, UV_vDesc_dc,     0, iCnt-1);
      7 : UP_QuickSort(sOrder, UV_vNum_samp,    0, iCnt-1);
      8 : UP_QuickSort(sOrder, UV_vGubn_sex,    0, iCnt-1);
      9 : UP_QuickSort(sOrder, UV_vCod_jangbi,  0, iCnt-1);
     10 : UP_QuickSort(sOrder, UV_vCod_hul,     0, iCnt-1);
     11 : UP_QuickSort(sOrder, UV_vCod_Cancer,  0, iCnt-1);
     12 : UP_QuickSort(sOrder, UV_vCod_chuga,   0, iCnt-1);
     13 : UP_QuickSort(sOrder, UV_vCod_etc,     0, iCnt-1);
      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;
end;

end.
