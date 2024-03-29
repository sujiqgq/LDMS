//==============================================================================
// 프로그램 설명 : [WorkList]분주관리 대장
// 시스템        : 통합검진
// 수정일자      : 2007.03.10
// 수정자        : 유동구
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD5Q01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD5Q01 = class(TfrmSingle)
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
    qryHangmok: TQuery;
    CBox_gubn: TCheckBox;
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
  private
    { Private declarations }
    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    function  UF_HangmokList(sTemp : String): String;
  public
    { Public declarations }
    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010,
    UV_v011, UV_v012, UV_v013 : Variant;
  end;

var
  frmLD5Q01: TfrmLD5Q01;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q011;

{$R *.DFM}

procedure TfrmLD5Q01.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
      UV_v013 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001, iLength);
      VarArrayRedim(UV_v002, iLength);
      VarArrayRedim(UV_v003, iLength);
      VarArrayRedim(UV_v004, iLength);
      VarArrayRedim(UV_v005, iLength);
      VarArrayRedim(UV_v006, iLength);
      VarArrayRedim(UV_v007, iLength);
      VarArrayRedim(UV_v008, iLength);
      VarArrayRedim(UV_v009, iLength);
      VarArrayRedim(UV_v010, iLength);
      VarArrayRedim(UV_v011, iLength);
      VarArrayRedim(UV_v012, iLength);
      VarArrayRedim(UV_v013, iLength);
   end;
end;

function TfrmLD5Q01.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '검진일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD5Q01.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
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
end;

procedure TfrmLD5Q01.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := GF_DateFormat(UV_v003[DataRow-1]);
      4 : Value := UV_v004[DataRow-1];
      5 : Value := UV_v005[DataRow-1];
      6 : Value := GF_JuminFormat(UV_v006[DataRow-1]);
      7 : Value := UV_v007[DataRow-1];
      8 : Value := UV_v008[DataRow-1];
      9 : Value := UV_v009[DataRow-1];
     10 : Value := UV_v010[DataRow-1];
     11 : Value := UV_v011[DataRow-1];
     12 : Value := UV_v012[DataRow-1];
     13 : Value := UV_v013[DataRow-1];
   end;
end;

procedure TfrmLD5Q01.btnQueryClick(Sender: TObject);
var i, index, iRet, temp : Integer;
    sWhere, yh_name, chuga, sHangmok  : String;
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

      sWhere := '  WHERE G.dat_gmgn = ''' + mskDate.Text + '''';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND B.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then
            sWhere := sWhere + ' AND G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end
      else
      begin
         if cmbCOD_JISA.ItemIndex = 0 then
            sWhere := sWhere + ' AND G.cod_jisa = ''' + UV_sCod_jisa + ''''
         else if cmbCOD_JISA.ItemIndex = 1 then
            sWhere := sWhere + ' AND G.cod_jisa = ''' + UV_sCod_jisa + '''';
      end;
      
      if CBox_gubn.Checked = False then
         sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';

      if Trim(mskFrom.Text) <> '' then
      begin
         sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
      end;

      sWhere := sWhere + ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD5Q01', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);
         // DataSet의 값을 Variant변수로 이동
         for i := 0 to RecordCount - 1 do
         begin
            yh_name := '';
            UV_v001[i] := FieldByName('Num_samp').AsString;
            UV_v002[i] := FieldByName('cod_bjjs').AsString;
            UV_v003[i] := FieldByName('dat_bunju').AsString;
            UV_v004[i] := FieldByName('num_bunju').AsString;
            UV_v005[i] := FieldByName('desc_name').AsString;
            UV_v006[i] := FieldByName('num_jumin').AsString;
            case StrToInt(copy(FieldByName('num_jumin').AsString,7,1)) of
               1,3,5,7,9 :  UV_v007[i] := '남';
               2,4,6,8,0 :  UV_v007[i] := '여';
               else UV_v007[i] := copy(FieldByName('num_jumin').AsString,7,1);
            end;
            UV_v008[i] := FieldByName('Cod_jangbi').AsString;
            UV_v009[i] := FieldByName('Cod_hul').AsString;
            UV_v010[i] := FieldByName('Cod_Cancer').AsString;

            chuga := ''; sHangmok := '';
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
               qryTkgum.Active := True;
               if qryTkgum.RecordCount > 0 then
               begin
                  sHangmok := FieldByName('Cod_chuga').AsString;
                  chuga := FieldByName('Cod_chuga').AsString + '(특검: ';
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
                          begin
                             chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                             if Pos('SE',qryPf_hm.FieldByName('COD_HM').AsString) > 0 then
                                sHangmok := sHangmok + qryPf_hm.FieldByName('COD_HM').AsString;
                          end;
                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
               end;           //if(qryTkgum) end;
               UV_v011[i]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
               UV_v013[i]  := UF_HangmokList(sHangmok);
            end
            else
            begin
               UV_v011[i]  := FieldByName('Cod_chuga').AsString;
               UV_v013[i]  := UF_HangmokList(FieldByName('Cod_chuga').AsString);
            end;

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

            // 생애전환기유형Display
            if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger) + ', '
            else if FieldByName('gubn_life').AsString = '2' then
               yh_name := yh_name + '생애전환기재검' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) }+ ', ';

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

            UV_v012[i] := yh_name;

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

procedure TfrmLD5Q01.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD5Q01.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD5Q01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD5Q01.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD5Q011 := TfrmLD5Q011.Create(Self);
//   frmLD5Q011.QuickRep.Preview;
   frmLD5Q011.QuickRep.Print;

end;

function  TfrmLD5Q01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD5Q01.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

function  TfrmLD5Q01.UF_HangmokList(sTemp : String): String;
var
   i : Integer;
   sHangmok_name : string;
begin
   Result := ''; sHangmok_name := '';
   for i := 1 to length(sTemp) div 4 do
   begin
      if (pos('DR',copy(sTemp, (i * 4) - 3, 4)) = 0) and
         (pos('JJ',copy(sTemp, (i * 4) - 3, 4)) = 0) and
         (pos('SE80',copy(sTemp, (i * 4) - 3, 4)) = 0) then
      begin
         qryHangmok.Active := False;
         qryHangmok.ParamByName('cod_hm').AsString := copy(sTemp, (i * 4) - 3, 4);
         qryHangmok.Active := True;
         if qryHangmok.RecordCount > 0 then
            sHangmok_name := sHangmok_name + qryHangmok.FieldByName('desc_kor').AsString + ', ';
      end;
   end;
   Result := copy(sHangmok_name,1,length(sHangmok_name) - 2);
end;

end.
