//==============================================================================
// 프로그램 설명 : 분주 list 출력
// 시스템        : 통합검진
// 수정일자      : 2005.10.01
// 수정자        : 유동구
// 수정내용      : 출장, 순수노신인 경우 Unrin(4종) 별표 체크...
// 참고사항      :
//==============================================================================
// 수정일자      : 2006.09.22
// 수정자        : 김승철
// 수정내용      : 조회구분(Y004,Y005,U017,U019) 조회기능추가.
// 참고사항      :
//==============================================================================
// 수정일자      : 2010.10.08
// 수정자        : 김희석
// 수정내용      : 검지일자별 정렬
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 분주 list 현황
// 시스템        : 통합검진
// 수정일자      : 2011.11.09
// 수정자        : 김희석
// 수정내용      :  주민번호 뒷자리 * 표시
//==============================================================================
unit LD2Q02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD2Q02 = class(TfrmSingle)
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
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    Label3: TLabel;
    RG_Query: TComboBox;
    Label4: TLabel;
    RD_print: TRadioButton;
    RD_preview: TRadioButton;
    Label5: TLabel;
    gumgin_sort: TCheckBox;
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
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    
  public
    { Public declarations }
    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_vNum_bunju, UV_vDesc_name, UV_vNum_jumin, UV_vCod_jisa, UV_vNum_samp,
    UV_vCod_jangbi, UV_vCod_hul, UV_vCod_Cancer, UV_vCod_chuga, UV_vDat_gmgn,
    UV_vCod_etc : Variant;
  end;

var
  frmLD2Q02: TfrmLD2Q02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q011, LD2Q021;

{$R *.DFM}


procedure TfrmLD2Q02.UP_GridSet;
var i : Integer;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      // Row갯수 초기화
      Rows := 0;

      // Column Title
      Col[1].Heading := '지 사';
      Col[2].Heading := '분주번호';
      Col[3].Heading := '의뢰자';
      Col[4].Heading := '주민번호';
      Col[5].Heading := 'M/F';
      Col[6].Heading := '검진일자';
      Col[7].Heading := '샘플';
      Col[8].Heading := '혈액';
      Col[9].Heading := '종양';
      Col[10].Heading := '장비';
      Col[11].Heading := '추가';
      Col[12].Heading := '기타검사';
   end;

end;

procedure TfrmLD2Q02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp   := VarArrayCreate([0, 0], varOleStr);
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

function TfrmLD2Q02.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '분주일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD2Q02.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_GridSet;
   UP_VarMemSet(0);

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

   // SQL문을 저장
   UV_sBasicSQL := qryBunju.SQL.Text;
   RG_Query.ItemIndex := 0;
end;

procedure TfrmLD2Q02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_jisa[DataRow-1];
      2 : Value := UV_vNum_bunju[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      5 : begin
             if Copy(UV_vNum_jumin[DataRow-1], 7, 1) = '1' then
                Value := '남'
             else
                Value := '여';
          end;
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vNum_samp[DataRow-1];
      8 : Value := UV_vCod_hul[DataRow-1];
      9 : Value := UV_vCod_Cancer[DataRow-1];
      10: Value := UV_vCod_jangbi[DataRow-1];
      11 : Value := UV_vCod_chuga[DataRow-1];
      12 : Value := UV_vCod_etc[DataRow-1];
   end;
end;

procedure TfrmLD2Q02.btnQueryClick(Sender: TObject);
var i, j, index, iRet, temp : Integer;
    sWhere, yh_name, chuga, sUrin10, chuga_hul : String;
    vCod_chuga : Variant;
begin
   inherited;

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

      sWhere := '  WHERE A.dat_bunju = ''' + mskDate.Text + '''';

      if Trim(mskFrom.Text) <> '' then
      begin
         sWhere := sWhere + ' AND A.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND A.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
      end;

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND A.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then
            sWhere := sWhere + ' AND A.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end
      else
      begin
         if cmbCOD_JISA.ItemIndex = 0 then
         begin
            sWhere := sWhere + ' AND A.cod_bjjs = ''15''';
            sWhere := sWhere + ' AND A.cod_jisa = ''' + UV_sCod_jisa + '''';
         end
         else if cmbCOD_JISA.ItemIndex = 1 then
            sWhere := sWhere + ' AND A.cod_bjjs = ''' + UV_sCod_jisa + ''''
         else
            sWhere := sWhere + ' AND A.cod_jisa = ''' + UV_sCod_jisa + '''';
      end;
      if RG_Query.ItemIndex = 1 then
      begin
           sWhere := sWhere + ' And (B.cod_jangbi in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''Y004'' or cod_hm = ''Y005'' or cod_hm = ''U017'' or cod_hm = ''U019'' ))) ';
           sWhere := sWhere + ' or B.cod_hul in ( select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''Y004'' or cod_hm = ''Y005'' or cod_hm = ''U017'' or cod_hm = ''U019'' ))) ';
           sWhere := sWhere + ' or B.cod_cancer in (select cod_pf from profile where cod_pf in ';
           sWhere := sWhere + ' (select cod_pf from profile_hm where ( cod_hm = ''Y004'' or cod_hm = ''Y005'' or cod_hm = ''U017'' or cod_hm = ''U019'' ))) ';
           sWhere := sWhere + ' or ( cod_chuga like ''%Y004%'' or cod_chuga like ''%Y005%'' or cod_chuga like ''%U017%'' or cod_chuga like ''%U019%'' )) ';
      end;
      if gumgin_sort.checked then
      begin
      sWhere := sWhere + ' ORDER BY B.dat_gmgn, A.num_bunju ';
      end
      else sWhere := sWhere + ' ORDER BY A.cod_jisa, A.num_bunju ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q02', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);
         // DataSet의 값을 Variant변수로 이동
         for i := 0 to RecordCount - 1 do
         begin
            yh_name := '';
            UV_vNum_bunju[i]  := FieldByName('Num_bunju').AsString;
            UV_vNum_samp[i]   := FieldByName('Num_samp').AsString;
            UV_vNum_jumin[i]  := Trim(Copy(FieldByName('Num_Jumin').AsString,1,6)) + '' + Trim(Copy(FieldByName('Num_Jumin').AsString,7,1)+'******');
            UV_vDat_gmgn[i]   := FieldByName('Dat_gmgn').AsString;
            UV_vCod_jisa[i]   := FieldByName('Cod_jisa').AsString;
            UV_vCod_jangbi[i] := FieldByName('Cod_jangbi').AsString;
            UV_vCod_hul[i]    := FieldByName('Cod_hul').AsString;
            UV_vCod_Cancer[i] := FieldByName('Cod_Cancer').AsString;

            // 본사-출장만 4종 검사를 함...
            if (FieldByName('Cod_jisa').AsString = '12') or
               (FieldByName('Cod_jisa').AsString = '15') then sUrin10 := 'N'
            else                                              sUrin10 := 'Y';

            if ((FieldByName('gubn_neawon').AsString = '2') and ((FieldByName('gubn_nosin').AsString = '1') or (FieldByName('gubn_adult').AsString = '1') or (FieldByName('gubn_Tkgm').AsString  = '1') or (FieldByName('gubn_life').AsString = '1'))) or
               ((FieldByName('gubn_jinch').AsString  = '2') and ((FieldByName('gubn_nosin').AsString = '1') or (FieldByName('gubn_adult').AsString = '1') or (FieldByName('gubn_life').AsString = '1')) and
                (FieldByName('dat_gmgn').AsString >= '20070626') and (FieldByName('dat_gmgn').AsString <= '20070628')) then
            begin
               //장비코드...
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_jangbi').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if qryPf_hm.FieldByName('COD_HM').AsString = 'U001' then
                        sUrin10 := 'Y';
                     qryPf_hm.next;
                  end;   // while(qryPf_hm) end;
               end;      //if(qryPf_hm) end;

               //혈액코드...
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if qryPf_hm.FieldByName('COD_HM').AsString = 'U001' then
                        sUrin10 := 'Y';
                     qryPf_hm.next;
                  end;   // while(qryPf_hm) end;
               end;      //if(qryPf_hm) end;

               //Cancer코드...
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_Cancer').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if qryPf_hm.FieldByName('COD_HM').AsString = 'U001' then
                        sUrin10 := 'Y';
                     qryPf_hm.next;
                  end;   // while(qryPf_hm) end;
               end;      //if(qryPf_hm) end;

               //추가검사
               if pos('U001',FieldByName('Cod_chuga').AsString) > 0 then
                  sUrin10 := 'Y';

               if sUrin10 = 'N' then
               begin
                  UV_vDesc_name[i]  := FieldByName('Desc_name').AsString + '★';
               end
               else UV_vDesc_name[i]  := FieldByName('Desc_name').AsString;
            end
            else UV_vDesc_name[i]  := FieldByName('Desc_name').AsString;


            chuga := '';  chuga_hul := '';

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
                  chuga := chuga_hul + '(특검: ';                               //07.01.12 철. 장비항목 뺌.
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
               UV_vCod_chuga[i]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
            end
            else
               UV_vCod_chuga[i]  := chuga_hul;                                  //07.01.12 철. 장비항목 뺌.
//               UV_vCod_chuga[i]  := FieldByName('Cod_chuga').AsString;

            // 노신유형Display
            if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
           else if FieldByName('gubn_nosin').AsString = '2' then
               yh_name := yh_name + '노신재검' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) }+ ', ';

            //특검유형Display
            if FieldByName('gubn_tkgm').AsString = '1' then
               yh_name := yh_name + '특검'{UF_No_Hangmok(copy(mskDate.Text,1,4), '2', FieldByName('gubn_tkyh').AsInteger)} + ', '
            else if FieldByName('gubn_tkgm').AsString = '2' then
               yh_name := yh_name + '특검재검'{UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_tkyh').AsInteger) + ', '};

            //성인병유형Display
            if FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
            else if FieldByName('gubn_adult').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_adyh').AsInteger) + ', ';

            //간염유형Display
            if FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
            else if FieldByName('gubn_agab').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_agyh').AsInteger) ;

            //공교의보유형Display
            if FieldByName('gubn_gong').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
            else if FieldByName('gubn_gong').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_goyh').AsInteger);

            // 생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger) + ', '
           else if FieldByName('gubn_life').AsString = '2' then
               yh_name := yh_name + '생애재검' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) }+ ', ';

            UV_vCod_etc[i]    := yh_name;

            Next;
            Inc(index);
         end;
         // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
         Active := False;
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

procedure TfrmLD2Q02.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD2Q02.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD2Q02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD2Q02.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD2Q021 := TfrmLD2Q021.Create(Self);
   if RD_print.Checked then
      frmLD2Q021.QuickRep.Print
   else if RD_preview.Checked then
      frmLD2Q021.QuickRep.Preview;

end;

function  TfrmLD2Q02.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD2Q02.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2Q02.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

end.
