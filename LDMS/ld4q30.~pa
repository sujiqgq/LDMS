//==============================================================================
// 프로그램 설명 : 조직결과 리스트
// 시스템        : 분석관리
// 수정일자      : 2011.02.
// 수정자        : 김승철
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.09
// 수정자        : 곽윤설
// 수정내용      : B012추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.06.15
// 수정자        : 곽윤설
// 수정내용      : B009 추가
// 참고사항      : 한의 재단병리팀장1500019
//==============================================================================
unit LD4Q30;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges ;

type
  TfrmLD4Q30 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    cmbJisa: TComboBox;
    mskFrom: TDateEdit;
    rbDate: TRadioButton;
    grdMaster: TtsGrid;
    btnsDate: TSpeedButton;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    Label4: TLabel;
    ComB_ShEndo: TComboBox;
    Label2: TLabel;
    CmbOrder: TComboBox;
    rbPan_Date: TRadioButton;
    mskPandate: TDateEdit;
    btn_Pandate: TSpeedButton;
    qryCa_desc_buwi_Max: TQuery;
    qryU_Cell: TQuery;
    Edte_No: TEdit;
    Label17: TLabel;
    Edts_No: TEdit;
    Rd_No: TRadioButton;
    qryCell_Old: TQuery;
    Gauge1: TGauge;
    Label_count: TLabel;
    Label3: TLabel;
    Ck_print: TCheckBox;
    qryJangbi: TQuery;
    qryCa_request: TQuery;
    qryCommon: TQuery;
    Ck_YSNO_SOKUN: TCheckBox;
    Label13: TLabel;
    Cmb_Class: TComboBox;
    Label5: TLabel;
    qryClass: TQuery;
    Ck_Class_High: TCheckBox;
    Panel2: TPanel;
    Rd_print1: TRadioButton;
    Rd_print2: TRadioButton;
    Label6: TLabel;
    Label7: TLabel;
    Cmb_Class2: TComboBox;
    Cmb_Class3: TComboBox;
    qryClass_cell_2: TQuery;
    qryClass_cell_3: TQuery;
    qryCommon2: TQuery;
    qryCommon3: TQuery;
    chk_jindan: TCheckBox;
    Ck_YSNO_SOKUN2: TCheckBox;
    chk_d_number: TCheckBox;
    Label8: TLabel;
    Label9: TLabel;
    Cmb_Class4: TComboBox;
    Cmb_Class5: TComboBox;
    qryClass_cell_4: TQuery;
    qryClass_cell_5: TQuery;
    rbLast: TRadioButton;
    mskST_date: TDateEdit;
    btnSt_date: TSpeedButton;
    Label10: TLabel;
    mskED_date: TDateEdit;
    btnED_date: TSpeedButton;

    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure UP_Click(Sender: TObject);
    procedure up_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure btnPrintClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    procedure chk_d_numberClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vGubn_check, UV_vDat_gmgn, UV_vDesc_name, UV_vNum_jumin, UV_vAge, UV_vSex,
    UV_vCod_jisa, UV_vNum_id, UV_vCod_cell, UV_vCod_hm, UV_vDesc_buwi, UV_vDesc_sokun,
    UV_vDat_gmgn_Old, UV_vDesc_buwi_Old, UV_vGum_Form_Old, UV_vDat_result, UV_vDoct_name, UV_vJindan_Class, UV_vNum_tel : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Class(sData : string) : String;

  public
    { Public declarations }
  end;

var
  frmLD4Q30: TfrmLD4Q30;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q301, LD4Q302;

{$R *.DFM}

procedure TfrmLD4Q30.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vGubn_check := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vAge        := VarArrayCreate([0, 0], varOleStr);
      UV_vSex        := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn_Old  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi_Old := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form_Old  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_result    := VarArrayCreate([0, 0], varOleStr);
      UV_vDoct_name     := VarArrayCreate([0, 0], varOleStr);
      UV_vJindan_Class  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_tel       := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vGubn_check, iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vAge,        iLength);
      VarArrayRedim(UV_vSex,        iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vCod_cell,   iLength);
      VarArrayRedim(UV_vCod_hm,     iLength);
      VarArrayRedim(UV_vDesc_buwi,  iLength);
      VarArrayRedim(UV_vDesc_sokun, iLength);
      VarArrayRedim(UV_vDat_gmgn_Old,  iLength);
      VarArrayRedim(UV_vDesc_buwi_Old, iLength);
      VarArrayRedim(UV_vGum_Form_Old,  iLength);
      VarArrayRedim(UV_vDat_result,    iLength);
      VarArrayRedim(UV_vDoct_name,     iLength);
      VarArrayRedim(UV_vJindan_Class,  iLength);
      VarArrayRedim(UV_vNum_tel,       iLength);
   end;
end;


function TfrmLD4Q30.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if mskFROM.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;



procedure TfrmLD4Q30.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   // 환경설정
   UP_VarMemSet(0);

   //판독확인용 체크시 분석부만 판정중 결과를 포함하여 확인할수 있도록 적용,, 20170424 본원이유경
   if ((GV_sUserID = '150007') or (GV_sUserID = '150681') or (GV_sUserID = '150644') or (GV_sUserID = '151021') or
       (GV_sUserID = '151022') or (GV_sUserID = '150879') or (GV_sUserID = '151030') or (GV_sUserID = '151026')) then  rd_print1.Visible := True;

   // 일자(From, To)를 오늘일자로 설정
   mskFrom.Text    := GV_sToday;
   mskPandate.Text := GV_sToday;
   mskST_date.Text := GV_sToday;
   mskED_date.Text  := GV_sToday;

   with qryClass do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                             FieldByName('DESC_CODE').AsString);
         Next;
      end;
      Close;
   end;
   with qryClass_cell_2 do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class2.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                              FieldByName('DESC_CODE').AsString);
         Next;
      end;
      Close;
   end;
   with qryClass_cell_3 do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class3.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                              FieldByName('DESC_CODE').AsString);
         Next;
      end;
      Close;
   end;
   with qryClass_cell_4 do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class4.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                              FieldByName('DESC_CODE').AsString);
         Next;
      end;
      Close;
   end;
   with qryClass_cell_5 do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class5.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                              FieldByName('DESC_CODE').AsString);
         Next;
      end;
      Close;
   end;

   Cmb_Class.ItemIndex    := 0;
   Cmb_Class2.ItemIndex   := 0;
   Cmb_Class3.ItemIndex   := 0;
   Cmb_Class4.ItemIndex   := 0;
   Cmb_Class5.ItemIndex   := 0;

   ComB_ShEndo.ItemIndex  := 0;
end;


procedure TfrmLD4Q30.btnQueryClick(Sender: TObject);
var index, iAge : Integer;
    sSelect, sWhere, sOrderBy, sMan, sSelect2, sWhere2 : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sWhere  := ''; sSelect := ''; sOrderBy := '';
   
   with qryCell do
   begin
      Active := False;
      // SQL문 생성

       sSelect := ' SELECT C.cod_hm, C.cod_bjjs, C.cod_cell, C.cod_jisa, C.num_id, C.dat_gmgn, C.cod_cell, gubn_order, C.YSNO_SOKUN, C.gubn_class, C.gubn_P012, ' +
                  ' G.desc_name, G.num_jumin, G.num_tel, C.Desc_buwi, C.Desc_sokun1, C.Desc_sokun2, C.Desc_sokun3, C.Desc_sokun4, C.Desc_sokun5, C.dat_result, ' +
                  ' J.Can1_Bung3_Jindan, J.Can2_Bung4_Jindan FROM cell C INNER JOIN gumgin G ' +
                  ' ON C.cod_jisa = G.cod_jisa and C.num_id = G.num_id and C.dat_gmgn = G.dat_gmgn ' +
                  ' LEFT OUTER JOIN Ca_Cancer2018 J ' +
                  ' ON C.cod_jisa = J.cod_jisa and C.num_id = J.num_id and C.dat_gmgn = J.dat_gmgn ';

      //2018.11.16 최윤선선임 요청.. 접수no.중복으로 접수되는 건들에 대해 확인이 가능하도록 해달라고 요청함.
      if  chk_d_number.Checked then
      begin
          sWhere := '  WHERE desc_buwi in(select desc_buwi  from cell where C.desc_buwi <> '''' group by desc_buwi HAVING count(*) > 1 ) and ';
      end
      else if pnlJisa.Visible then
      begin
         if Copy(cmbJisa.Text, 1, 2) <> '00' then
            sWhere := ' WHERE C.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And '
         else
            sWhere := ' WHERE ';
      end
      else
         sWhere := ' WHERE C.cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And ';

      if rbDate.Checked then
         sWhere := sWhere +  ' C.dat_gmgn = ''' + mskFROM.Text + ''' And '
      else if rbPan_Date.Checked then
         sWhere := sWhere +  ' C.dat_result = ''' + mskPandate.Text + ''' And '
      else
      if rbLast.Checked then
      begin
         sWhere := sWhere +  ' C.dat_last >= ''' + mskST_date.Text + '''';
         sWhere := sWhere +  ' AND C.dat_last <= ''' + mskED_date.Text + ''' and ';
      end
      else if Rd_No.Checked then
      begin
         if (EDTS_NO.Text <> '') and (EDTE_NO.Text <> '') then
            sWhere := sWhere + '  C.desc_buwi >= ''' + Edts_No.Text + ''''
                             + '  AND C.desc_buwi <= ''' + Edte_No.Text + ''' And '
         else
         begin
            showmessage('접수 No 를 입력해 주세요.');
            Edts_No.SetFocus;
            exit;
         end;
      end;

      if ComB_ShEndo.ItemIndex = 1 then
         sWhere := sWhere + ' SUBSTRING(C.cod_cell,1,1) = ''P'' '
      else if ComB_ShEndo.ItemIndex = 2 then
         sWhere := sWhere + ' SUBSTRING(C.cod_cell,1,1) = ''J'' '
      else if ComB_ShEndo.ItemIndex = 3 then
         sWhere := sWhere + ' C.cod_hm = ''B001'' '
      else if ComB_ShEndo.ItemIndex = 4 then
         sWhere := sWhere + ' C.cod_hm = ''B005'' '
      else if ComB_ShEndo.ItemIndex = 5 then
         sWhere := sWhere + ' C.cod_hm = ''B007'' '
      else if ComB_ShEndo.ItemIndex = 6 then
         sWhere := sWhere + ' C.cod_hm = ''B008'' '
      else if ComB_ShEndo.ItemIndex = 7 then
         sWhere := sWhere + ' C.cod_hm = ''B012'' '
      else if ComB_ShEndo.ItemIndex = 8 then
         sWhere := sWhere + ' C.cod_hm = ''P001'' '
      else if ComB_ShEndo.ItemIndex = 9 then
         sWhere := sWhere + ' C.cod_hm = ''P003'' '
      else if ComB_ShEndo.ItemIndex = 10 then
         sWhere := sWhere + ' C.cod_hm = ''P009'' '
      else if ComB_ShEndo.ItemIndex = 11 then
         sWhere := sWhere + ' C.cod_hm = ''P010'' '
      else if ComB_ShEndo.ItemIndex = 12 then
         sWhere := sWhere + ' C.cod_hm = ''P011'' '
      else if ComB_ShEndo.ItemIndex = 13 then
         sWhere := sWhere + ' C.cod_hm = ''P006'' '
      else if ComB_ShEndo.ItemIndex = 14 then
         sWhere := sWhere + ' C.cod_hm = ''B009'' '
      else if ComB_ShEndo.ItemIndex = 15 then
         sWhere := sWhere + ' C.cod_hm = ''P012'' '
      else if ComB_ShEndo.ItemIndex = 16 then
         sWhere := sWhere + ' C.cod_hm = ''P013'' '
      else
         sWhere := sWhere + ' C.cod_hm <> '''' ';

      if Ck_YSNO_SOKUN.Checked then
         sWhere := sWhere + ' And C.YSNO_SOKUN <> ''N'' ';
      if Ck_YSNO_SOKUN2.Checked then
         sWhere := sWhere + ' And C.YSNO_SOKUN = ''N'' ';
      if chk_jindan.Checked then
         sWhere := sWhere + ' And C.check_jindan = ''Y'' ';

      if Rd_print2.Checked then
         sWhere := sWhere + ' And (C.YSNO_SOKUN <> ''N'' AND C.YSNO_SOKUN <> ''Y'') ';

      if (copy(Cmb_Class.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class.Text, 1, 2)) > 0) then
         sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,1) + ''''
      else if copy(Cmb_Class.Text,1,1) <> '0' then
         sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,2) + '''';

      if (copy(Cmb_Class3.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class3.Text, 1, 2)) > 0) then
         sWhere := sWhere + '  AND J.Can1_bung3_Jindan = ''' + copy(Cmb_Class3.Text,1,1) + ''''
      else if copy(Cmb_Class3.Text,1,1) <> '0' then
         sWhere := sWhere + '  AND J.Can1_bung3_Jindan = ''' + copy(Cmb_Class3.Text,1,2) + '''';

      if (copy(Cmb_Class2.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class2.Text, 1, 2)) > 0) then
         sWhere := sWhere + '  AND J.Can2_bung4_Jindan = ''' + copy(Cmb_Class2.Text,1,1) + ''''
      else if copy(Cmb_Class2.Text,1,1) <> '0' then
         sWhere := sWhere + '  AND J.Can2_bung4_Jindan = ''' + copy(Cmb_Class2.Text,1,2) + '''';

      if (copy(Cmb_Class4.Text,1,2) <> '00') and (Pos('-', Copy(Cmb_Class4.Text, 1, 2)) > 0) then
         sWhere := sWhere + '  AND C.GUBN_P012 = ''' + copy(Cmb_Class4.Text,1,1) + ''''
      else if copy(Cmb_Class4.Text,1,2) <> '00' then
         sWhere := sWhere + '  AND C.GUBN_P012 = ''' + copy(Cmb_Class4.Text,1,2) + '''';

      if (copy(Cmb_Class5.Text,1,2) <> '00') and (Pos('-', Copy(Cmb_Class5.Text, 1, 2)) > 0) then
         sWhere := sWhere + '  AND C.GUBN_P012 = ''' + copy(Cmb_Class5.Text,1,1) + ''''
      else if copy(Cmb_Class5.Text,1,2) <> '00' then
         sWhere := sWhere + '  AND C.GUBN_P012 = ''' + copy(Cmb_Class5.Text,1,2) + '''';

      if Ck_Class_High.checked = true then
         sWhere := sWhere + ' and C.gubn_class in (''9'',''10'',''11'',''12'',''13'',''14'')' ;

      if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' order by G.cod_jisa, desc_name'
      else if CmbOrder.ItemIndex = 2 then
         sOrderBy := ' order by G.num_jumin'
      else if CmbOrder.ItemIndex = 3 then
         sOrderBy := ' order by C.Desc_buwi'
      else
         sOrderBy := ' order by G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

//      Gauge1.Progress := 0;
      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD4Q30', 'Q', 'N');
//         Gauge1.MinValue := 0;
//         Gauge1.MaxValue := qryCell.RecordCount;
         // 조건에 맞는 데이터 변수에 입력
         while not Eof do
         begin
//            Gauge1.Progress     := Gauge1.Progress + 1;
//            Label_count.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
//            Application.ProcessMessages;

            iAge := 0; sMan := '';
            GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
            UP_VarMemSet(index);

            if FieldByName('gubn_order').AsString = 'Y' then UV_vGubn_check[index]   := '1'
            else                                             UV_vGubn_check[index]   := '0';

            UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
            UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
            UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
            UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
            UV_vNum_id[index]     := FieldByName('Num_id').AsString;
            UV_vCod_cell[index]   := FieldByName('Cod_cell').AsString;
            UV_vCod_hm[index]     := FieldByName('Cod_hm').AsString;
            UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
            if FieldByName('YSNO_SOKUN').AsString = 'C' then
               UV_vDat_result[index] := FieldByName('Dat_result').AsString
            else UV_vDat_result[index] := '';
            UV_vDesc_sokun[index] := FieldByName('Desc_sokun1').AsString +
                                     FieldByName('Desc_sokun2').AsString +
                                     FieldByName('Desc_sokun3').AsString +
                                     FieldByName('Desc_sokun4').AsString +
                                     FieldByName('Desc_sokun5').AsString;
            UV_vAge[index]        := iAge;
            UV_vSex[index]        := sMan;
            UV_vNum_tel[index]    := FieldByName('num_tel').AsString;

            //P012
            if (UV_vCod_hm[index] = 'P012') then
            begin
                 if qryCell.FieldByName('GUBN_P012').AsString = '01' then
                    UV_vJindan_Class[index] := '01-검체불량(부적합)'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '02' then
                    UV_vJindan_Class[index] := '02-classⅰ(정상)'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '03' then
                    UV_vJindan_Class[index] := '03-classⅱ(정상-반응성)'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '04' then
                    UV_vJindan_Class[index] := '04-classⅲ(비정형세포)'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '05' then
                    UV_vJindan_Class[index] := '05-classⅳ(암의심)'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '06' then
                    UV_vJindan_Class[index] := '06-classⅴ(암)'
              end
              else if (UV_vCod_hm[index] = 'P013') then
              begin
                 if qryCell.FieldByName('GUBN_P012').AsString = '01' then
                    UV_vJindan_Class[index] := '01-검체불량'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '02' then
                    UV_vJindan_Class[index] := '02-음성ⅰ'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '03' then
                    UV_vJindan_Class[index] := '03-음성ⅱ'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '04' then
                    UV_vJindan_Class[index] := '04-의양성 ⅲ'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '05' then
                    UV_vJindan_Class[index] := '05-양성ⅳ'
                 else if qryCell.FieldByName('GUBN_P012').AsString = '06' then
                    UV_vJindan_Class[index] := '06-양성ⅴ'
            end;
            sWhere2  := ''; sSelect2 := '';
            if (UV_vCod_hm[index] = 'B001') or (UV_vCod_hm[index] = 'B007') then
            begin
               with qryJangbi do
               begin
                  Active := False;
                  sSelect2 := 'select S.desc_name, C.Can1_Bung3_Jindan, C.Can2_Bung4_Jindan, C.Can1_Bung3_Etc from jangbi J left outer join sawon S ' +
                              ' On J.cod_doct = S.cod_sawon ' +
                              ' left outer join ca_cancer2018 C ' +
                              ' On J.cod_jisa = C.cod_jisa ' +
                              ' and J.num_id = C.num_id ' +
                              ' and J.dat_gmgn = C.dat_gmgn ';
                  sWhere2  := ' where J.cod_jisa = ''' + UV_vCod_jisa[index] + ''' ';
                  sWhere2  := sWhere2 + ' and J.num_id = ''' + UV_vNum_id[index] + ''' ';
                  sWhere2  := sWhere2 + ' and J.dat_gmgn = ''' + UV_vDat_gmgn[index] + ''' ';
                  sWhere2  := sWhere2 + ' and J.desc_sokun <> '''' ';

                  if UV_vCod_hm[index] = 'B001' then
                     sWhere2  := sWhere2 + ' and J.cod_jangbi in (''JJ34'', ''JJB9'') '
                  else if UV_vCod_hm[index] = 'B007' then
                     sWhere2  := sWhere2 + ' and J.cod_jangbi in (''JJ83'', ''JJ86'') ';

                  SQL.Clear;
                  SQL.Add(sSelect2 + sWhere2);
                  Active := True;

                  UV_vDoct_name[index]  := qryJangbi.FieldByName('desc_name').AsString;

                  //위조직진단
                  if (UV_vCod_hm[index] = 'B001') then
                  begin
                     if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '1' then
                        UV_vJindan_Class[index] := '1.이상소견없음'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '21' then
                        UV_vJindan_Class[index] := '2-1.위염'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '22' then
                        UV_vJindan_Class[index] := '2-2.위축성위염/장상피화생'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '3' then
                        UV_vJindan_Class[index] := '3.염증성 또는 증식성 변변'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '4' then
                        UV_vJindan_Class[index] := '4.저도샘종 혹은 이형성'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '5' then
                        UV_vJindan_Class[index] := '5.고도샘종 혹은 이형성'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '6' then
                        UV_vJindan_Class[index] := '6.암의심'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '7' then
                        UV_vJindan_Class[index] := '7.암'
                     else if qryJangbi.FieldByName('Can1_Bung3_Jindan').AsString = '8' then
                        UV_vJindan_Class[index] := '8.기타';
                  end
                  //대장조직진단
                  else if (UV_vCod_hm[index] = 'B007') then
                  begin
                     if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '1' then
                        UV_vJindan_Class[index] := '1.정상'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '2' then
                        UV_vJindan_Class[index] := '2.염증성 또는 증식성 병변'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '3' then
                        UV_vJindan_Class[index] := '3.저도선종 혹은 이형성'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '4' then
                        UV_vJindan_Class[index] := '4.고도선종 혹은 이형성'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '5' then
                        UV_vJindan_Class[index] := '5.암의심'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '6' then
                        UV_vJindan_Class[index] := '6.암'
                     else if qryJangbi.FieldByName('Can2_Bung4_Jindan').AsString = '7' then
                        UV_vJindan_Class[index] := '7.기타';
                  end;

               end;
            end
            else if (UV_vCod_hm[index] = 'P003') or (UV_vCod_hm[index] = 'B009') then
            begin
               with qryCa_request do
               begin
                  Active := False;
                  sSelect2 := 'select S.desc_name from ca_request J left outer join sawon S ' +
                              ' On J.cod_doctor = S.cod_sawon ';
                  sWhere2  := ' where J.cod_jisa = ''' + UV_vCod_jisa[index] + ''' ';
                  sWhere2  := sWhere2 + ' and J.num_id = ''' + UV_vNum_id[index] + ''' ';
                  sWhere2  := sWhere2 + ' and J.dat_gmgn = ''' + UV_vDat_gmgn[index] + ''' ';

                  SQL.Clear;
                  SQL.Add(sSelect2 + sWhere2);
                  Active := True;

                  UV_vDoct_name[index]  := qryCa_request.FieldByName('desc_name').AsString;
                  UV_vJindan_Class[index] := UF_Class(qryCell.FieldByName('gubn_class').AsString);
               end;
            end
            else UV_vDoct_name[index]  := '';


            with qryCell_Old do
            begin
               qryCell_Old.Close;
               ParamByName('cod_bjjs').Asstring := qryCell.FieldByName('cod_bjjs').AsString;
               ParamByName('Num_id').Asstring   := qryCell.FieldByName('Num_id').AsString;
               ParamByName('dat_gmgn').Asstring := qryCell.FieldByName('dat_gmgn').AsString;
               ParamByName('cod_hm').Asstring   := qryCell.FieldByName('cod_hm').AsString;

               qryCell_Old.Open;
               if qryCell_Old.IsEmpty = False then
               begin
                  while not qryCell_Old.Eof do
                  begin
                     if UV_vDat_gmgn_Old[index] <> '' then
                        UV_vDat_gmgn_Old[index]  := UV_vDat_gmgn_Old[index] + ', ' + qryCell_Old.FieldByName('dat_gmgn').AsString
                     else UV_vDat_gmgn_Old[index]  := qryCell_Old.FieldByName('dat_gmgn').AsString;

                     if UV_vDesc_buwi_Old[index] <> '' then
                        UV_vDesc_buwi_Old[index] := UV_vDesc_buwi_Old[index] + ', ' + qryCell_Old.FieldByName('Desc_buwi').AsString
                     else UV_vDesc_buwi_Old[index] := qryCell_Old.FieldByName('Desc_buwi').AsString;

                     if UV_vGum_Form_Old[index] <> '' then
                        UV_vGum_Form_Old[index] := UV_vGum_Form_Old[index] + ', ';
                     if qryCell_Old.FieldByName('Gum_Form1').AsString = 'Y' then
                        UV_vGum_Form_Old[index]  := UV_vGum_Form_Old[index] + '음성'
                     else if qryCell_Old.FieldByName('Gum_Form2').AsString = 'Y' then
                        UV_vGum_Form_Old[index]  := UV_vGum_Form_Old[index] + '상피세포이상'
                     else if qryCell_Old.FieldByName('Gum_Form3').AsString = 'Y' then
                     begin
                        UV_vGum_Form_Old[index]  := UV_vGum_Form_Old[index] + '기타(자궁내막세포출현 등)';
                        if qryCell_Old.FieldByName('Gum_Form3_Etc').AsString <> '' then
                           UV_vGum_Form_Old[index]  := UV_vGum_Form_Old[index] + ' :' + qryCell_Old.FieldByName('Gum_Form3_Etc').AsString;
                     end;
                     Next;
                  end;
               end;
               qryCell_Old.Close;
            end;    

            Next;
            Inc(index);
         end;
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
      // Table과 Disconnect
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD4Q30.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD4Q30.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
//      1 : Value := StrToint(UV_vGubn_check[DataRow-1]);
      1 : Value := UV_vCod_jisa[DataRow-1];
      2 : Value := UV_vCod_hm[DataRow-1];
      3 : Value := UV_vDat_gmgn[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : value := UV_vNum_id[DataRow-1];
      6 : value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vAge[DataRow-1];
      8 : Value := UV_vSex[DataRow-1];
      9 : Value := UV_vDoct_name[DataRow-1];
     10 : Value := UV_vDesc_buwi[DataRow-1];
     11 : Value := UV_vDat_result[DataRow-1];
     12 : Value := UV_vJindan_Class[DataRow-1];
     13 : Value := UV_vDesc_sokun[DataRow-1];
     14 : Value := UV_vDat_gmgn_Old[DataRow-1];
     15 : Value := UV_vDesc_buwi_Old[DataRow-1];
     16 : Value := UV_vGum_Form_Old[DataRow-1];
     17 : Value := UV_vNum_tel[DataRow-1];
   end;
end;


procedure TfrmLD4Q30.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];

end;


procedure TfrmLD4Q30.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnsDate then
      GF_CalendarClick(mskFrom);
   if Sender = btn_Pandate then
      GF_CalendarClick(mskPandate);
   if Sender = btnSt_date then
      GF_CalendarClick(mskST_date);
   if Sender = btnEd_date then
      GF_CalendarClick(mskED_date);
end;


procedure TfrmLD4Q30.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskFrom        then UP_Click(btnsDate);
end;

procedure TfrmLD4Q30.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   if Rd_print1.Checked then
   begin
      frmLD4Q301 := TfrmLD4Q301.Create(Self);
      frmLD4Q301.QuickRep1.Preview;
   end
   else if Rd_print2.Checked then
   begin
      frmLD4Q302 := TfrmLD4Q302.Create(Self);
      frmLD4Q302.QuickRep1.Preview;
   end
end;

procedure TfrmLD4Q30.cmbRelationChange(Sender: TObject);
var i, j : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
 { if cmbRelation.ItemIndex = 1 then
  begin
     if Application.MessageBox('접수No. 가 없는 건에 대해서 접수No. 를 일괄부여 하시겠습니까?', '확인', MB_YESNO+MB_ICONquestion) = IDNO then
        Exit;

     for i := 0 to grdMaster.Rows - 1 do
     begin
        // 자료가 없을경우의 처리
        if grdMaster.Rows = 0 then
        begin
           exit;
        end;

        if (trim(UV_vDesc_buwi[i]) = '') and
           (UV_vCod_hm[i] <> 'B001') and
           (UV_vCod_hm[i] <> 'B007') and
           (UV_vCod_hm[i] <> 'B008') then
        begin
           qryU_Cell.ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[i];
           qryU_Cell.ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[i];
           qryU_Cell.ParamByName('Num_jumin').AsString  := UV_vNum_jumin[i];
           qryU_Cell.ParamByName('DAT_update').AsString := GV_sToday;
           qryU_Cell.ParamByName('COD_update').AsString := GV_sUserId;

           //'C'로 시작하는건
           qryCa_desc_buwi_Max.Active := False;
           if (UV_vCod_hm[i] = 'P001') or
              (UV_vCod_hm[i] = 'P003') then
           begin
              qryCa_desc_buwi_Max.ParamByName('desc_buwi_s').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000000';
              qryCa_desc_buwi_Max.ParamByName('desc_buwi_e').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '999999';
           end
           //'NG'로 시작하는건
           else if (UV_vCod_hm[i] = 'P005') or
                   (UV_vCod_hm[i] = 'P009') or
                   (UV_vCod_hm[i] = 'P010') or
                   (UV_vCod_hm[i] = 'P011') or
                   (UV_vCod_hm[i] = 'P012') then
           begin
              qryCa_desc_buwi_Max.ParamByName('desc_buwi_s').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000000';
              qryCa_desc_buwi_Max.ParamByName('desc_buwi_e').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '999999';
           end;
           qryCa_desc_buwi_Max.Active := True;

           if Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) = '' then
           begin
              if (UV_vCod_hm[i] = 'P001') or (UV_vCod_hm[i] = 'P003') then
                 qryU_Cell.ParamByName('desc_buwi').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001'
              else if (UV_vCod_hm[i] = 'P005') or
                      (UV_vCod_hm[i] = 'P009') or
                      (UV_vCod_hm[i] = 'P010') or
                      (UV_vCod_hm[i] = 'P011') or
                      (UV_vCod_hm[i] = 'P012') then
                 qryU_Cell.ParamByName('desc_buwi').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
           end
           else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                   (UV_vCod_hm[i] = 'P001') or (UV_vCod_hm[i] = 'P003') then
              qryU_Cell.ParamByName('desc_buwi').Asstring := Copy(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString, 1, 4) + GF_LPad(IntToStr(StrToInt(Copy(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString, 5, 6)) + 1), 6, '0')
           else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                   (UV_vCod_hm[i] = 'P005') or (UV_vCod_hm[i] = 'P009') or
                   (UV_vCod_hm[i] = 'P010') or (UV_vCod_hm[i] = 'P011') or (UV_vCod_hm[i] = 'P012') then
              qryU_Cell.ParamByName('desc_buwi').Asstring := Copy(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString, 1, 5) + GF_LPad(IntToStr(StrToInt(Copy(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString, 6, 6)) + 1), 6, '0');

           qryU_Cell.ExecSQL;
           // Table disconnect
           qryCa_desc_buwi_Max.Active := False;
        end;
     end;
     btnQuery.Click;
//     grdMaster.Repaint;
  end;     }

       if ComB_ShEndo.ItemIndex = 3 then
       begin
          Cmb_Class2.Enabled := False;
          Cmb_Class2.itemindex := 0;
          Cmb_Class3.Enabled := True;
          Cmb_Class4.Enabled := False;
          Cmb_Class4.itemindex := 0;
          Cmb_Class5.Enabled := False;
          Cmb_Class5.itemindex := 0;
       end
       else if ComB_ShEndo.ItemIndex = 5 then
       begin
          Cmb_Class2.Enabled := True;
          Cmb_Class3.Enabled := False;
          Cmb_Class3.itemindex := 0;
          Cmb_Class4.Enabled := False;
          Cmb_Class4.itemindex := 0;
          Cmb_Class5.Enabled := False;
          Cmb_Class5.itemindex := 0;
       end
       else if ComB_ShEndo.ItemIndex = 15 then
       begin
          Cmb_Class2.Enabled := False;
          Cmb_Class2.itemindex := 0;
          Cmb_Class3.Enabled := False;
          Cmb_Class3.itemindex := 0;
          Cmb_Class4.Enabled := True;
          Cmb_Class5.Enabled := False;
          Cmb_Class5.itemindex := 0;
       end
       else if ComB_ShEndo.ItemIndex = 16 then
       begin
          Cmb_Class2.Enabled := False;
          Cmb_Class2.itemindex := 0;
          Cmb_Class3.Enabled := False;
          Cmb_Class3.itemindex := 0;
          Cmb_Class4.Enabled := False;
          Cmb_Class4.itemindex := 0;
          Cmb_Class5.Enabled := True;
       end;
end;

function TfrmLD4Q30.UF_Class(sData : String) : String;
begin
   with qryCommon do
   begin
      Active := False;
      ParamByName('cod_detail').AsString := sData;
      Active := True;
      Result := qryCommon.FieldByName('desc_code').AsString;
   end;
   with qryCommon2 do
   begin
      Active := False;
      ParamByName('cod_detail').AsString := sData;
      Active := True;
      Result := qryCommon.FieldByName('desc_code').AsString;
   end;
   with qryCommon3 do
   begin
      Active := False;
      ParamByName('cod_detail').AsString := sData;
      Active := True;
      Result := qryCommon.FieldByName('desc_code').AsString;
   end;
end;

procedure TfrmLD4Q30.chk_d_numberClick(Sender: TObject);
begin
  inherited;
  if  chk_d_number.checked = true then
  begin
     rbDate.Enabled := False;
     rbPan_Date.Enabled := False;
     rbLast.Enabled := False;
  end
  else
  begin
     rbDate.Enabled := True;
     rbPan_Date.Enabled := True;
     rbPan_Date.Enabled := True;
  end;
end;

end.


