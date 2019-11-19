//==============================================================================
// 프로그램 설명 : [WorkList]특수검진 검사항목 대장
// 시스템        : 통합검진
// 개발일자      : 2008.01.09
// 개발자        : 유동구
//==============================================================================
// 수정일자      : 2017.08.08
// 수정자        : 유동구
// 수정사항      : (한의 재단특수검사파트1700283)
//                 1.검진센터 위치 변경(분주번호 앞으로 이동)
//                 2.분주센터, 성별 표기 삭제
//                 3.유해물질 추가(검사항목List 앞으로 생성)
//                 4.조회 시 혈액샘플No 순으로 정렬
//==============================================================================
unit LD4Q12;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, Gauges, ComObj;

type
  TfrmLD4Q12 = class(TfrmSingle)
    qryBunju: TQuery;
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    mskFrom: TMaskEdit;
    Label10: TLabel;
    mskTo: TMaskEdit;
    cmbCOD_JISA: TComboBox;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    qryHangmok: TQuery;
    Com_Part: TComboBox;
    Gauge1: TGauge;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Panel27: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Pnl_Jisa: TPanel;
    Panel5: TPanel;
    RBtn_print: TRadioButton;
    RBtn_preveiw: TRadioButton;
    Panel6: TPanel;
    chk_stool: TCheckBox;
    CheckBox: TCheckBox;
    Panel4: TPanel;
    Chk_JJXE: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009 : Variant;

    { Private declarations }
    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
  end;

var
  frmLD4Q12: TfrmLD4Q12;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q121;

{$R *.DFM}

procedure TfrmLD4Q12.UP_VarMemSet(iLength : Integer);
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
   end;
end;

function TfrmLD4Q12.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '검진일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD4Q12.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid의 환경설정 & Variant변수 Memory할당
   UP_VarMemSet(0);

   // Login 지사가 '00'이면 '15'로 대체
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   if UV_sCod_jisa = '15' then
   begin
      Pnl_Jisa.Caption := '센터';
      GP_ComboJisa(cmbCOD_JISA);
   end
   else
   begin
      Pnl_Jisa.Caption := '분주소:';
      cmbCod_jisa.Items.Add('연 구 소');
      cmbCod_jisa.Items.Add('지    사');
      cmbCod_jisa.ItemIndex := 1;      
   end;

   Com_Part.ItemIndex := 8;

   // SQL문을 저장
   UV_sBasicSQL := qryBunju.SQL.Text;
end;

procedure TfrmLD4Q12.grdMasterCellLoaded(Sender: TObject; DataCol,
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
      6 : Value := UV_v006[DataRow-1];
      7 : Value := GF_JuminFormat(UV_v007[DataRow-1]);
      8 : Value := UV_v008[DataRow-1];
      9 : Value := UV_v009[DataRow-1];
   end;
end;

procedure TfrmLD4Q12.btnQueryClick(Sender: TObject);
var i, index, iRet, temp : Integer;
    sWhere, yh_name, chuga, sHangmok, sMatter  : String;
    vCod_chuga : Variant;

    bSE42_Check : Boolean;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid & Chart 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   Gauge1.MinValue := 0;
   Gauge1.Progress := 0;
   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sWhere := '  WHERE B.dat_bunju = ''' + mskDate.Text + ''' ';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND B.cod_bjjs = ''' + UV_sCod_jisa + ''' ';
         if Trim(cmbCod_jisa.Text) <> '' then sWhere := sWhere + ' AND G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + ''' '
         else                                 sWhere := sWhere + ' AND G.cod_jisa <> ''51'' ';
      end
      else
      begin
         if      cmbCOD_JISA.ItemIndex = 0 then sWhere := sWhere + ' AND G.cod_jisa = ''' + UV_sCod_jisa + ''' '
         else if cmbCOD_JISA.ItemIndex = 1 then sWhere := sWhere + ' AND G.cod_jisa = ''' + UV_sCod_jisa + ''' ';
      end;

      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';

      if Trim(mskFrom.Text) <> '' then
      begin
         sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + ''' ';
         sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text),   4, '0') + ''' ';
      end;

      sWhere := sWhere + ' ORDER BY CAST(G.num_samp AS INT), B.num_bunju ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q12', 'Q', 'N');

         UP_VarMemSet(RecordCount-1);
         Gauge1.MaxValue := RecordCount;

         // DataSet의 값을 Variant변수로 이동
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            application.ProcessMessages;

            yh_name := '';

            //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 체크...
            bSE42_Check := False;

            //[1]검진센터
            UV_v001[index] := FieldByName('desc_jisa').AsString;
            //[2]분주번호
            UV_v002[index]  := FieldByName('num_bunju').AsString;
            //[3]분주일자
            UV_v003[index]  := FieldByName('dat_bunju').AsString;
            //[4]혈액샘플 No.
            UV_v004[index]  := FieldByName('Num_samp').AsString;
            //[5]단체명
            UV_v005[index] := FieldByName('desc_dc').AsString;
            //[6]성명
            UV_v006[index]  := FieldByName('desc_name').AsString;
            //[7]주민번호
            UV_v007[index] := copy(FieldByName('num_jumin').AsString,1,7) + '******';

            //장비코드 추출..---------------------------------------------------
            chuga := ''; sHangmok := '';  sMatter := '';
            qryPf_hm.Active := False;
            qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_jangbi').AsString;
            qryPf_hm.Active := True;
            if qryPf_hm.RecordCount > 0 then
            begin
               while not qryPf_hm.Eof do
               begin
                  if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                     (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     if Trim(chuga) = '' then chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString
                     else                     chuga := chuga + ',' + #13 + qryPf_hm.FieldByName('COD_HM').AsString;

                     if Trim(sHangmok) = '' then sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString
                     else                        sHangmok := sHangmok + ',' + #13 + qryPf_hm.FieldByName('Desc_Kor').AsString;
                  end;
                  qryPf_hm.next;
               end;   // while(qryPf_hm) end;
            end;      //if(qryPf_hm) end;
            //혈액코드 추출..---------------------------------------------------
            qryPf_hm.Active := False;
            qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_hul').AsString;
            qryPf_hm.Active := True;
            if qryPf_hm.RecordCount > 0 then
            begin
               while not qryPf_hm.Eof do
               begin
                  if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                     (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     if Trim(chuga) = '' then chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString
                     else                     chuga := chuga + ',' + #13 + qryPf_hm.FieldByName('COD_HM').AsString;

                     if Trim(sHangmok) = '' then sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString
                     else                        sHangmok := sHangmok + ',' + #13 + qryPf_hm.FieldByName('Desc_Kor').AsString;
                  end;
                  qryPf_hm.next;
               end;   // while(qryPf_hm) end;
            end;      //if(qryPf_hm) end;
            //종양코드 추출..---------------------------------------------------
            qryPf_hm.Active := False;
            qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_Cancer').AsString;
            qryPf_hm.Active := True;
            if qryPf_hm.RecordCount > 0 then
            begin
               while not qryPf_hm.Eof do
               begin
                  if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                     (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     if Trim(chuga) = '' then chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString
                     else                     chuga := chuga + ',' + #13 + qryPf_hm.FieldByName('COD_HM').AsString;

                     if Trim(sHangmok) = '' then sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString
                     else                        sHangmok := sHangmok + ',' + #13 + qryPf_hm.FieldByName('Desc_Kor').AsString;
                  end;
                  qryPf_hm.next;
               end;   // while(qryPf_hm) end;
            end;      //if(qryPf_hm) end;
            //추가코드 추출..---------------------------------------------------
            iRet := GF_MulToSingle(FieldByName('Cod_chuga').AsString, 4, vCod_chuga);
            for Temp := 0 to iRet - 1 do
            begin
               qryHangmok.Active := False;
               qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[Temp];
               qryHangmok.Active := True;
               if qryHangmok.RecordCount > 0 then
               begin
                  if (Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga) = 0) and
                     (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     if Trim(chuga) = '' then chuga := chuga + qryHangmok.FieldByName('COD_HM').AsString
                     else                     chuga := chuga + ',' + #13 + qryHangmok.FieldByName('COD_HM').AsString;

                     if Trim(sHangmok) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                        sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end; //end of if(qryHangmok)
            end; //end of for(Cod_chuga)
            //특검코드 추출..---------------------------------------------------
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
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
                          //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                          //----------------------------------------------------
                          if (vCod_chuga[Temp] = 'TC41') and
                             (pos('JJXE',qryTkgum.FieldByName('Cod_chuga').AsString) > 0) and
                             (qryPf_hm.FieldByName('COD_HM').AsString = 'SE42') then
                          begin
                             // 혈중카복시(SE42) 검사 Skip...
                             bSE42_Check := True;

                             if Chk_JJXE.Checked then
                             begin
                                //유해물질리스트...
                                if trim(sMatter) = '' then  sMatter := sMatter + '일산화탄소(CO)'
                                else
                                begin
                                   if pos(qryPf_hm.FieldByName('desc_pf').AsString, sMatter) <= 0 then
                                      sMatter := sMatter + ',' + #13 + '일산화탄소(CO)';
                                end;
     
                                if Trim(chuga) = '' then chuga := chuga + 'JJXE'
                                else                     chuga := chuga + ',' + #13 + 'JJXE';
     
                                if Trim(sHangmok) = '' then sHangmok := sHangmok + '호기중일산화탄소농도'
                                else                        sHangmok := sHangmok + ',' + #13 + '호기중일산화탄소농도';
                             end;
                          end
                          else
                          begin
                             if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                                (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                             begin
                                //유해물질리스트...
                                if trim(sMatter) = '' then  sMatter := sMatter + qryPf_hm.FieldByName('desc_pf').AsString
                                else
                                begin
                                   if pos(qryPf_hm.FieldByName('desc_pf').AsString, sMatter) <= 0 then
                                      sMatter := sMatter + ',' + #13 +qryPf_hm.FieldByName('desc_pf').AsString;
                                end;

                                if Trim(chuga) = '' then chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString
                                else                     chuga := chuga + ',' + #13 + qryPf_hm.FieldByName('COD_HM').AsString;

                                if Trim(sHangmok) = '' then sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString
                                else                        sHangmok := sHangmok + ',' + #13 + qryPf_hm.FieldByName('Desc_Kor').AsString;
                             end;
                          end;
                          //====================================================

                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
                  //추가코드 추출..---------------------------------------------------
                  iRet := GF_MulToSingle(qryTkgum.FieldByName('Cod_chuga').AsString, 4, vCod_chuga);
                  for Temp := 0 to iRet - 1 do
                  begin
                     qryHangmok.Active := False;
                     qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[Temp];
                     qryHangmok.Active := True;
                     if qryHangmok.RecordCount > 0 then
                     begin
                        if (Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga) = 0) and
                           (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                        begin
                           if Trim(chuga) = '' then chuga := chuga + qryHangmok.FieldByName('COD_HM').AsString
                           else                     chuga := chuga + ',' + #13 + qryHangmok.FieldByName('COD_HM').AsString;

                           if Trim(sHangmok) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                           else                        sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                        end;
                     end; //end of if(qryHangmok)
                  end; //end of for(Cod_chuga)
               end;           //if(qryTkgum) end;
            end;

            //[8]유해물질리스트...
            UV_v008[index] := sMatter;

            // 4. 노신, 성인병, 간염, 생애전환기에 대한 검사항목 추출--------------------------
            yh_name := '';
            if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            if FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            if FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(yh_name, 4, vCod_chuga);
            for Temp := 0 to iRet - 1 do
            begin
               qryHangmok.Active := False;
               qryHangmok.ParamByName('cod_hm').AsString := vCod_chuga[Temp];
               qryHangmok.Active := True;
               if qryHangmok.RecordCount > 0 then
               begin
                  if (Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga) = 0) and
                     (qryHangmok.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                  begin
                     if Trim(chuga) = '' then chuga := chuga + qryHangmok.FieldByName('COD_HM').AsString
                     else                     chuga := chuga + ',' + #13 + qryHangmok.FieldByName('COD_HM').AsString;

                     if Trim(sHangmok) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                        sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end; //end of if(qryHangmok)
            end;

            if FieldByName('gubn_hulgum').AsString = '2' then
            begin
               if (Com_Part.ItemIndex = 2) and (chk_stool.checked = true) then
               begin
                   if pos('Y004', chuga) > 0 then
                   begin
                       UV_v009[index] := sHangmok;
                       Next;
                       Inc(index);
                   end
                   else
                       Next;
               end
               else
               begin
                   UV_v009[index] := sHangmok;
                   Next;
                   Inc(index);
               end;
            end
            else if (FieldByName('gubn_hulgum').AsString = '3') and (copy(Com_Part.Text,1,2) <> '09') then
            begin
               sHangmok := '';
               if pos('C009', chuga) > 0 then
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := 'C009';
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if Trim(chuga) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                     sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end;

               if pos('C010', chuga) > 0 then
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := 'C010';
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if Trim(chuga) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                     sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end;

               if pos('C011', chuga) > 0 then
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := 'C011';
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if Trim(chuga) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                     sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end;

               if pos('C025', chuga) > 0 then
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := 'C025';
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if Trim(chuga) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                     sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end;

               if pos('C032', chuga) > 0 then
               begin
                  qryHangmok.Active := False;
                  qryHangmok.ParamByName('cod_hm').AsString := 'C032';
                  qryHangmok.Active := True;
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if Trim(chuga) = '' then sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString
                     else                     sHangmok := sHangmok + ',' + #13 + qryHangmok.FieldByName('Desc_Kor').AsString;
                  end;
               end;
               UV_v009[index] := sHangmok;
               //재검+TC77 이면 리스트 띄우기. 신진주 수석
               if (CheckBox.Checked) and ((qryBunju.fieldByName('gubn_tkgm').AsString <> '2') and (pos('TC77',FieldByName('cod_prf').AsString) > 0)) then
               begin
               // 그냥 빈공간...
               end
               else
                  Inc(index);

               Next;
            end
            else
            begin
               UV_v009[index] := sHangmok;


               if Chk_JJXE.Checked then
               begin
                  if bSE42_Check then Inc(index);
               end
               else
               begin
                  //[2018.07.19-유동구]일산화탄소(TC41)과 호기중일산화탄소농도(JJXE) 있을 경우 혈중카복시(SE42) 제외...
                  //위 조건에 카복시만 있는 경우 빈 값으로 출력되는 부분 제외..
                  if ((bSE42_Check) and (sHangmok = '')) or
                     (CheckBox.Checked) and ((qryBunju.fieldByName('gubn_tkgm').AsString <> '2') and (pos('TC77',FieldByName('cod_prf').AsString) > 0)) then
                  begin
                     // 그냥 빈공간...
                  end
                  else
                     Inc(index);
               end;

               Next;
            end;

            //Next;
            //Inc(index);
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

procedure TfrmLD4Q12.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD4Q12.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);

   if chk_stool.checked = true then Com_Part.ItemIndex := 2;
   if Com_Part.ItemIndex <> 2  then chk_stool.checked := false;
end;
procedure TfrmLD4Q12.UP_Change(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Com_Part.ItemIndex <> 2 then
      chk_stool.checked := false;
end;

procedure TfrmLD4Q12.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD4Q12.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q121 := TfrmLD4Q121.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q121.QuickRep.Preview;
   if RBtn_print.Checked   then frmLD4Q121.QuickRep.Print;

end;

function  TfrmLD4Q12.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q12.SBut_ExcelClick(Sender: TObject);
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
