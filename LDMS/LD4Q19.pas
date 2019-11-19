//==============================================================================
// 프로그램 설명 : [WorkList]특수검진 검사항목 대장
// 시스템        : 통합검진
// 수정일자      : 2008.01.09
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.07.14
// 수정자        : 곽윤설
// 수정내용      : 요10종유린 추가
// 참고사항      : 한의 재단특검출장파트1500104
//==============================================================================

unit LD4Q19;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, Gauges;

type
  TfrmLD4Q19 = class(TfrmSingle)
    qryBunju: TQuery;
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
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
    chk_JJFH: TCheckBox;
    chk_JJFG: TCheckBox;
    chk_H029: TCheckBox;
    chk_H031: TCheckBox;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RBtn_print: TRadioButton;
    RB_bunju: TRadioButton;
    Label3: TLabel;
    MaskEdit1: TMaskEdit;
    MaskEdit2: TMaskEdit;
    RadioButton2: TRadioButton;
    chk_U010: TCheckBox;
    Gauge1: TGauge;
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
  public
    { Public declarations }
    // Grid와 연관된 Variant 변수 선언(Report에서 참조하므로 Public에 선언)
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_vADD1, UV_vADD2 : Variant;
  end;

var
  frmLD4Q19: TfrmLD4Q19;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q191;

{$R *.DFM}

procedure TfrmLD4Q19.UP_VarMemSet(iLength : Integer);
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
      UV_vADD1 := VarArrayCreate([0, 0], varOleStr);
      UV_vADD2 := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_vADD1, iLength);
      VarArrayRedim(UV_vADD2, iLength);
   end;
end;

function TfrmLD4Q19.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '검진일자를 입력해야 합니다.');
      Result := False;
   end;
end;

procedure TfrmLD4Q19.FormCreate(Sender: TObject);
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

procedure TfrmLD4Q19.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v004[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := GF_DateFormat(UV_v003[DataRow-1]);
      4 : Value := UV_vADD1[DataRow-1];
      5 : Value := UV_v001[DataRow-1];
      6 : Value := UV_vADD2[DataRow-1];
      7 : Value := UV_v005[DataRow-1];
      8 : Value := GF_JuminFormat(UV_v006[DataRow-1]);
      9 : Value := UV_v007[DataRow-1];
     10 : Value := UV_v008[DataRow-1];
   end;
end;

procedure TfrmLD4Q19.btnQueryClick(Sender: TObject);
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
   Gauge1.MinValue := 0;
   Gauge1.Progress := 0;
   with qryBunju do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      if UV_sCod_jisa = '15' then
      begin
         sWhere := ' WHERE B.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then
            sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end
      else
      begin
         if cmbCOD_JISA.ItemIndex = 0 then
            sWhere := ' WHERE G.cod_jisa = ''' + UV_sCod_jisa + ''''
         else if cmbCOD_JISA.ItemIndex = 1 then
            sWhere := ' WHERE G.cod_jisa = ''' + UV_sCod_jisa + '''';
      end;

      sWhere := sWhere + ' AND B.dat_bunju = ''' + mskDate.Text + '''';

      if RB_bunju.Checked then
      begin
         sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '9') + '''';
      end
      else
      begin
         sWhere := sWhere + ' AND G.num_samp >= ''' + GF_LPad(Trim(mskFrom.Text), 6, '0') + '''';
         sWhere := sWhere + ' AND G.num_samp <= ''' + GF_LPad(Trim(mskTo.Text), 6, '9') + '''';
      end;

      sWhere := sWhere + ' AND G.gubn_tkgm in (''1'',''2'')';

      sWhere := sWhere + ' ORDER BY B.num_bunju';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q19', 'Q', 'N');
         Gauge1.MaxValue := RecordCount;

         // DataSet의 값을 Variant변수로 이동
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            application.ProcessMessages;

            //장비코드 추출..---------------------------------------------------
            chuga := ''; sHangmok := '';
            qryPf_hm.Active := False;
            qryPf_hm.ParamByName('cod_pf').AsString := FieldByName('Cod_jangbi').AsString;
            qryPf_hm.Active := True;
            if qryPf_hm.RecordCount > 0 then
            begin
               while not qryPf_hm.Eof do
               begin
                  if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                  begin
                     chuga    := chuga    + qryPf_hm.FieldByName('COD_HM').AsString + '   ';
                     sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString + '   ';
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
                  if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                  begin
                     chuga    := chuga    + qryPf_hm.FieldByName('COD_HM').AsString + '   ';
                     sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString + '   ';
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
                  if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                  begin
                     chuga    := chuga    + qryPf_hm.FieldByName('COD_HM').AsString + '   ';
                     sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString + '   ';
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
                  if Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga) = 0 then
                  begin
                     chuga    := chuga    + qryHangmok.FieldByName('COD_HM').AsString + '   ';
                     sHangmok := sHangmok + qryHangmok.FieldByName('Desc_Kor').AsString + '   ';
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
                          if Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0 then
                          begin
                             chuga    := chuga    + qryPf_hm.FieldByName('COD_HM').AsString + '   ';
                             sHangmok := sHangmok + qryPf_hm.FieldByName('Desc_Kor').AsString + '   ';
                          end;
                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
               end;           //if(qryTkgum) end;
               chuga := chuga + qryTkgum.FieldByName('COD_CHUGA').AsString + '   ';
            end;

            if ((chk_JJFH.Checked = True) and (pos('JJFH',chuga) > 0)) or
               ((chk_JJFG.Checked = True) and (pos('JJFG',chuga) > 0)) or
               ((chk_H029.Checked = True) and (pos('H029',chuga) > 0)) or
               ((chk_H031.Checked = True) and (pos('H031',chuga) > 0)) or
               ((chk_U010.Checked = True) and (pos('U001',chuga) > 0) and (pos('U002',chuga) > 0) and (pos('U003',chuga) > 0) and
                (pos('U004',chuga) > 0) and (pos('U005',chuga) > 0) and (pos('U006',chuga) > 0) and (pos('U007',chuga) > 0) and
                (pos('U008',chuga) > 0) and (pos('U009',chuga) > 0) and (pos('U010',chuga) > 0)) then
            begin
               UP_VarMemSet(index);
               yh_name := '';
               UV_vADD1[index] := FieldByName('desc_jisa').AsString;
               UV_vADD2[index] := FieldByName('desc_dc').AsString;
               UV_v001[index] := FieldByName('Num_samp').AsString;
               UV_v002[index] := FieldByName('cod_bjjs').AsString;
               UV_v003[index] := FieldByName('dat_bunju').AsString;
               UV_v004[index] := FieldByName('num_bunju').AsString;
               UV_v005[index] := FieldByName('desc_name').AsString;
//               UV_v006[index] := FieldByName('num_jumin').AsString;
               UV_v006[index] := copy(FieldByName('num_jumin').AsString,1,7) + '******';
               case StrToInt(copy(FieldByName('num_jumin').AsString,7,1)) of
                  1,3,5,7,9 :  UV_v007[index] := '남';
                  2,4,6,8,0 :  UV_v007[index] := '여';
                  else UV_v007[index] := copy(FieldByName('num_jumin').AsString,7,1);
               end;

               sHangmok := '';
               if (chk_JJFH.Checked = True) and (pos('JJFH',chuga) > 0) then
                  sHangmok := sHangmok + '객담세포검사 ';
               if (chk_JJFG.Checked = True) and (pos('JJFG',chuga) > 0) then
                  sHangmok := sHangmok + '소변세포병리 ';
               if (chk_H029.Checked = True) and (pos('H029',chuga) > 0) then
                  sHangmok := sHangmok + '망상적혈구 ';
               if (chk_H031.Checked = True) and (pos('H031',chuga) > 0) then
                  sHangmok := sHangmok + '메트헤모글로불린 ';
               if (chk_U010.Checked = True) and (pos('U001',chuga) > 0) and (pos('U002',chuga) > 0) and
                  (pos('U003',chuga) > 0) and (pos('U004',chuga) > 0) and (pos('U005',chuga) > 0) and
                  (pos('U006',chuga) > 0) and (pos('U007',chuga) > 0) and (pos('U008',chuga) > 0) and
                  (pos('U009',chuga) > 0) and (pos('U010',chuga) > 0) then
                  sHangmok := sHangmok + '요10종유린 ';

               UV_v008[index] := copy(sHangmok,1,length(sHangmok)-1);
               Inc(index);
            end;

            Next;
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

procedure TfrmLD4Q19.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD4Q19.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);

end;
procedure TfrmLD4Q19.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD4Q19.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q191 := TfrmLD4Q191.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q191.QuickRep.Preview;
   if RBtn_print.Checked   then frmLD4Q191.QuickRep.Print;

end;

function  TfrmLD4Q19.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.
