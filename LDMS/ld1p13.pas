//==============================================================================
// 프로그램 설명 : 조직검사 결과출력 form
// 수정일자      : 2015.08.01
// 수정자        : 곽윤설
// 수정내용      : 신규 개발
// 참고사항      : 한의 재단병리팀1500161 [병리팀 - 박예진 요청]
//==============================================================================
// 수정일자      : 2016.03.021
// 수정자        : 박수지
// 수정내용      : b009추가
// 참고사항      : 한의 재단병리팀1600026[병리팀 - 박예진 요청]
//==============================================================================
unit ld1p13;


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD1P13 = class(TfrmPrint)
    Label4: TLabel;
    msksDate: TDateEdit;
    btnsDate: TSpeedButton;
    Label10: TLabel;
    mskeDate: TDateEdit;
    btneDate: TSpeedButton;
    Label6: TLabel;
    edtsName: TEdit;
    Label7: TLabel;
    edteName: TEdit;
    Label9: TLabel;
    cmbJisa: TComboBox;
    Label8: TLabel;
    R_Date: TDateEdit;
    RDate: TSpeedButton;
    Label11: TLabel;
    sorting: TComboBox;
    Label12: TLabel;
    gubun: TComboBox;
    Ck_Class: TCheckBox;
    CheckBox: TCheckBox;
    jun_gum: TCheckBox;
    Label16: TLabel;
    Edts_No: TEdit;
    Label17: TLabel;
    Edte_No: TEdit;
    Label19: TLabel;
    Date_gmgn: TDateEdit;
    Dategmgn: TSpeedButton;
    cmbDoctor: TComboBox;
    qryDoctor: TQuery;
    Label20: TLabel;
    Label21: TLabel;
    Ck_Desc_dc: TCheckBox;
    Ck_Doctor: TCheckBox;
    Ck_Class_High: TCheckBox;
    qryCommon: TQuery;
    Chk_New: TCheckBox;
    Label13: TLabel;
    Cmb_Class: TComboBox;
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure Ck_DoctorClick(Sender: TObject);
  private
    { Private declarations }

    UV_sCod_jisa : String;

    // SQL문
    UV_sBasicSQL : String;

    function  UF_Condition : Boolean;
    function  UF_Query : Boolean;
    function  UF_Query2 : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD1P13: TfrmLD1P13;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LDAI01, LD1P132;

{$R *.DFM}

procedure TfrmLD1P13.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if      Sender = msksDate then UP_Click(btnsDate)
   else if Sender = mskeDate then UP_Click(btneDate);

end;

procedure TfrmLD1P13.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = gubun then
   begin
      if (gubun.ItemIndex = 6) or (gubun.ItemIndex = 8) then
      begin
         Label13.Enabled   := True;
         Cmb_Class.Enabled := True;
      end
      else
      begin
         Label13.Enabled   := false;
         Cmb_Class.Enabled := false;
      end;
   end;

end;

procedure TfrmLD1P13.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnsDate then GF_CalendarClick(msksDate)
   else
   if Sender = btneDate then GF_CalendarClick(mskeDate)
   else
   if Sender = RDate then GF_CalendarClick(R_Date)
   else
   if Sender = Dategmgn then GF_CalendarClick(Date_gmgn);
end;

procedure TfrmLD1P13.btnPreviewClick(Sender: TObject);
begin
  inherited;


  if Chk_New.Checked = False then
  begin
    frmLD1P132 := TfrmLD1P132.Create(Self);

     // SQL문을 저장
     UV_sBasicSQL := frmLD1P132.qryGumgin.SQL.Text;

     if UF_Query then
        frmLD1P132.QuickRep.Preview
     else
        ShowMessage('자료가 존재하지 않습니다');
  end
  else if Chk_New.Checked = True then
  begin
    frmLD1P132 := TfrmLD1P132.Create(Self);

     // SQL문을 저장
     UV_sBasicSQL := frmLD1P132.qryGumgin.SQL.Text;

     if UF_Query2 then
        frmLD1P132.QuickRep.Preview
     else
        ShowMessage('자료가 존재하지 않습니다');
  end;

end;

function  TfrmLD1P13.UF_Query : Boolean;
var sWhere : String;
begin

   Result:= False;

   // 조회조건 Check
   if not UF_Condition then exit;

   with frmLD1P132.qryGumgin do
   begin
      Active := False;
      sWhere := ' Where  G.dat_gmgn >= ''' + msksDate.Text + ''''
              + ' AND G.dat_gmgn <= ''' + mskeDate.Text + '''';

      if copy(cmbJisa.Items[cmbJisa.itemindex],1,2) <> '00' then
         sWhere := sWhere +' AND G.Cod_jisa  = ''' + copy(cmbJisa.Items[cmbJisa.itemindex],1,2) + '''';

      case gubun.itemindex of
        0 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''B001'',''B007'',''B008'', ''B010'', ''B012'', ''P001'',''P003'',''P008'',''P009'',''P010'',''P011'',''P012'',''B009'')';
            end;
        1 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B001''';
            end;
        2 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B007''';
            end;
        3 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B008''';
            end;
        4 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''B001'',''B007'')';
            end;
        5 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B010''';
            end;
        6 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B012''';
            end;
        7 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''P001'',''P003'', ''P004'',''P009'',''P010'',''P011'')';
              if (copy(Cmb_Class.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class.Text, 1, 2)) > 0) then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,1) + ''''
              else if copy(Cmb_Class.Text,1,1) <> '0' then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,2) + '''';
            end;
        8 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P001''';
            end;
        9 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P003''';
              if (copy(Cmb_Class.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class.Text, 1, 2)) > 0) then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,1) + ''''
              else if copy(Cmb_Class.Text,1,1) <> '0' then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,2) + '''';
            end;
       10 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P004''';
            end;
       11 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P006''';
            end;
       12 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P009''';
            end;
       13 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P010''';
            end;
       14 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P011''';
            end;
       15 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P012''';
            end;
       16 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B009''';
            end;
       17 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P013''';
            end;
      end;

      if CheckBox.Checked then
      begin
         if Trim(msksDate.Text) >= '20070701' then
         begin
            sWhere := sWhere + '  AND C.ysno_sokun = ''Y''';
         end
         else
         begin
            if (Trim(msksDate.Text) <= '20070701') and (Trim(mskeDate.Text) >= '20070701') then
            begin
               ShowMessage('2007년 7월 이후 자료는 따로 조회하십시요..' + #13#10 +
                           'ex) 2007.01.01 ~ 2007.06.31과 2007.07.01 ~ 2007.12.31)');
               exit;
            end;
         end;
      end
      else
      begin
         if Trim(msksDate.Text) >= '20070701' then
         begin
            sWhere := sWhere + '  AND C.ysno_sokun = ''C''';
         end
         else
         begin
            if (Trim(msksDate.Text) <= '20070701') and (Trim(mskeDate.Text) >= '20070701') then
            begin
               ShowMessage('2007년 7월 이후 자료는 따로 조회하십시요..' + #13#10 +
                           'ex) 2007.01.01 ~ 2007.06.31과 2007.07.01 ~ 2007.12.31)');
               exit;
            end;
         end;
      end;

      if Trim(edtsName.Text) <> '' then
      begin
         sWhere := sWhere + '  AND G.desc_name >= ''' + edtsName.Text + ''''
                          + '  AND G.desc_name <= ''' + edteName.Text + '''';
      end;

      if Trim(Edts_No.Text) <> '' then
      begin
         sWhere := sWhere + '  AND C.desc_buwi >= ''' + Edts_No.Text + ''''
                          + '  AND C.desc_buwi <= ''' + Edte_No.Text + '''';
      end;

      sWhere := sWhere + '  AND C.cod_doct <> ''''';

      // 준장비만 출력가능
      if jun_gum.Checked = True then
         sWhere := sWhere + ' AND G.cod_jangbi = '''' AND G.cod_hul <> '''' AND G.cod_chuga like ''%JJ%''';

      if (R_Date.text <> '') then
      begin
         sWhere := sWhere + '  AND C.dat_result = ''' + R_Date.Text + ''''
      end;

      if Ck_Class.checked = true then
      begin
         sWhere := sWhere + ' and C.cod_hm = ''P003'''
                          + ' and C.cod_pan <> '''''
                          + ' and not (C.gubn_class = ''1'' and C.desc_sokun1 = '''')'
      end;

      if Ck_Class_High.checked = true then
      begin
         sWhere := sWhere + ' and C.gubn_class in (''9'',''10'',''11'',''12'',''13'',''14'')' ;
      end;

      sWhere := sWhere + '  GROUP BY G.num_id, G.desc_name, G.dat_gmgn, G.cod_jisa, G.num_jumin, G.cod_dc, G.desc_dept, C.cod_hm, D.desc_dc, G.gubn_neawon';

      case sorting.itemindex of
        0 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_name';
            end;
        1 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_dept, G.desc_name';
            end;
        2 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_dept, G.desc_name';
            end;
        3 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_name';
            end;
        4 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.dat_gmgn, G.desc_name';
            end;
        5 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.desc_name';
            end;
        6 : begin
              sWhere := sWhere + '  ORDER BY G.dat_gmgn, G.desc_name';
            end;
        7 : begin
              sWhere := sWhere + '  ORDER BY G.desc_name';
            end;
      end;

      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);
      GP_query_log(frmLD1P132.qryGumgin, 'LD1I12', 'Q', 'N');
      Active := True;

      if frmLD1P132.qryGumgin.IsEmpty = False then
      begin
         if rbAll.Checked then
         begin
            frmLD1P132.QuickRep.PrinterSettings.FirstPage := 0;
            frmLD1P132.QuickRep.PrinterSettings.LastPage  := 0;
         end
         else if rbPart.Checked then
         begin
            frmLD1P132.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
            frmLD1P132.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
         end;
         Result := True;
      end
      else
      begin
         Result := False;
         exit;
      end;
   end;

end;

function TfrmLD1P13.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (Trim(msksDate.Text) = '') or (Trim(msksDate.Text) = '') then
   begin
      GF_MsgBox('Warning', '검진일자를 입력하여야 합니다');
      Result := False;
   end;
end;

procedure TfrmLD1P13.FormCreate(Sender: TObject);
var  sJisa : string;
begin
  inherited;

   GP_ComboJisa(cmbJisa);
   cmbJisa.Enabled := True;
   sJisa := copy(GV_sUserId,1,2);

   with qryCommon do
   begin
      Open;
      while not Eof do
      begin
         Cmb_Class.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                             FieldByName('DESC_CODE').AsString);
         Next;
      end;
   end;

   Cmb_Class.ItemIndex := 0;
   sorting.itemindex   := 0;
   gubun.itemindex     := 0;

   if (GV_sJisa <> '') then
   begin
      R_Date.Enabled := True;
      RDate.Enabled  := True;
   end;

   msksDate.Text := frmLDAI01.mskDate.Text;
   mskeDate.Text := frmLDAI01.mskDate.Text;
   CmbJIsa.Text := frmLDAI01.grdMaster.Cell[1,frmLDAI01.grdMaster.CurrentDataRow];
   GP_ComboSeek(cmbJisa);
   edtsName.Text := frmLDAI01.grdMaster.Cell[3,frmLDAI01.grdMaster.CurrentDataRow];
   edteName.Text := frmLDAI01.grdMaster.Cell[3,frmLDAI01.grdMaster.CurrentDataRow];
   Edts_No.Text  := frmLDAI01.edtDESC_BUWI.Text;
   Edte_No.Text  := frmLDAI01.edtDESC_BUWI.Text;
   gubun.Text    := frmLDAI01.grdMaster.Cell[6,frmLDAI01.grdMaster.CurrentDataRow];
   GP_ComboSeek(gubun);
   R_Date.Text   := frmLDAI01.mskDat_Result.Text;

   With  qryDoctor  do
   begin
      Close;
      cmbDoctor.Items.Clear;
      If (Copy(cmbJisa.Text,1,2) = '88') Then
         ParamByname('Cod_jisa').AsString := '15'
      Else If (Copy(cmbJisa.Text,1,2) = '99') Then                         //06.11.10 철. 전체 지사일 경우 접속자 해당 지사 의사선생님으로.
         ParamByname('Cod_jisa').AsString := copy(GV_sUserId,1,2)
      Else
         If (Copy(GV_sUserId,1,2) = '00') Then
            ParamByname('Cod_jisa').AsString := copy(GV_sUserId,1,2)
         Else ParamByname('Cod_jisa').AsString := Copy(cmbJisa.Text,1,2);
      Open;
      if  RecordCount > 0 then
      begin
        while  not Eof do
        begin
           cmbDoctor.items.Add(FieldByname('Desc_name').AsString);
           Next;
        End;
      end;
      Close;
   End;
end;

procedure TfrmLD1P13.btnPrintClick(Sender: TObject);
begin
  inherited;
  if Chk_New.Checked = False then
  begin
    frmLD1P132 := TfrmLD1P132.Create(Self);

     // SQL문을 저장
     UV_sBasicSQL := frmLD1P132.qryGumgin.SQL.Text;

     if UF_Query then
        frmLD1P132.QuickRep.print
     else
        ShowMessage('자료가 존재하지 않습니다');
  end
  else if Chk_New.Checked = True then
  begin
    frmLD1P132 := TfrmLD1P132.Create(Self);

     // SQL문을 저장
     UV_sBasicSQL := frmLD1P132.qryGumgin.SQL.Text;

     if UF_Query2 then
        frmLD1P132.QuickRep.print
     else
        ShowMessage('자료가 존재하지 않습니다');
  end;

end;

procedure TfrmLD1P13.btnCloseClick(Sender: TObject);
begin
  inherited;
  close;
end;

procedure TfrmLD1P13.Ck_DoctorClick(Sender: TObject);
begin
  inherited;
  if Ck_Doctor.Checked = true then
     cmbDoctor.Enabled := True
  else cmbDoctor.Enabled := False;
end;

function TfrmLD1P13.UF_Query2: Boolean;
var sWhere : String;
begin

   Result:= False;

   // 조회조건 Check
   if not UF_Condition then exit;

   with frmLD1P132.qryGumgin do
   begin
      Active := False;
      sWhere := ' Where  G.dat_gmgn >= ''' + msksDate.Text + ''''
              + ' AND G.dat_gmgn <= ''' + mskeDate.Text + '''';

      if copy(cmbJisa.Items[cmbJisa.itemindex],1,2) <> '00' then
         sWhere := sWhere +' AND G.Cod_jisa  = ''' + copy(cmbJisa.Items[cmbJisa.itemindex],1,2) + '''';

      case gubun.itemindex of
        0 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''B001'',''B007'',''B008'', ''B010'', ''B012'', ''P001'',''P003'',''P008'',''P009'',''P010'',''P011'',''P012'',''B009'')';
            end;
        1 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B001''';
            end;
        2 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B007''';
            end;
        3 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B008''';
            end;
        4 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''B001'',''B007'')';
            end;
        5 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B010''';
            end;
        6 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B012''';
            end;
        7 : begin
              sWhere := sWhere + '  AND C.cod_hm in (''P001'',''P003'', ''P004'',''P009'',''P010'',''P011'')';
              if (copy(Cmb_Class.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class.Text, 1, 2)) > 0) then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,1) + ''''
              else if copy(Cmb_Class.Text,1,1) <> '0' then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,2) + '''';
            end;
        8 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P001''';
            end;
        9 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P003''';
              if (copy(Cmb_Class.Text,1,1) <> '0') and (Pos('-', Copy(Cmb_Class.Text, 1, 2)) > 0) then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,1) + ''''
              else if copy(Cmb_Class.Text,1,1) <> '0' then
                 sWhere := sWhere + '  AND C.GUBN_CLASS = ''' + copy(Cmb_Class.Text,1,2) + '''';
            end;
       10 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P004''';
            end;
       11 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P006''';
            end;
       12 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P009''';
            end;
       13 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P010''';
            end;
       14 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P011''';
            end;
       15 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P012''';
            end;
       16 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''B009''';
            end;
       17 : begin
              sWhere := sWhere + '  AND C.cod_hm = ''P013''';
            end;
      end;

      if CheckBox.Checked then
      begin
         if Trim(msksDate.Text) >= '20070701' then
         begin
            sWhere := sWhere + '  AND C.ysno_sokun = ''Y''';
         end
         else
         begin
            if (Trim(msksDate.Text) <= '20070701') and (Trim(mskeDate.Text) >= '20070701') then
            begin
               ShowMessage('2007년 7월 이후 자료는 따로 조회하십시요..' + #13#10 +
                           'ex) 2007.01.01 ~ 2007.06.31과 2007.07.01 ~ 2007.12.31)');
               exit;
            end;
         end;
      end
      else
      begin
         if Trim(msksDate.Text) >= '20070701' then
         begin
            sWhere := sWhere + '  AND C.ysno_sokun = ''C''';
         end
         else
         begin
            if (Trim(msksDate.Text) <= '20070701') and (Trim(mskeDate.Text) >= '20070701') then
            begin
               ShowMessage('2007년 7월 이후 자료는 따로 조회하십시요..' + #13#10 +
                           'ex) 2007.01.01 ~ 2007.06.31과 2007.07.01 ~ 2007.12.31)');
               exit;
            end;
         end;
      end;

      if Trim(edtsName.Text) <> '' then
      begin
         sWhere := sWhere + '  AND G.desc_name >= ''' + edtsName.Text + ''''
                          + '  AND G.desc_name <= ''' + edteName.Text + '''';
      end;

      if Trim(Edts_No.Text) <> '' then
      begin
         sWhere := sWhere + '  AND C.desc_buwi >= ''' + Edts_No.Text + ''''
                          + '  AND C.desc_buwi <= ''' + Edte_No.Text + '''';
      end;

      sWhere := sWhere + '  AND C.cod_doct <> ''''';

      // 준장비만 출력가능
      if jun_gum.Checked = True then
         sWhere := sWhere + ' AND G.cod_jangbi = '''' AND G.cod_hul <> '''' AND G.cod_chuga like ''%JJ%''';

      if (R_Date.text <> '') then
      begin
         sWhere := sWhere + '  AND C.dat_result = ''' + R_Date.Text + ''''
      end;

      if Ck_Class.checked = true then
      begin
         sWhere := sWhere + ' and C.cod_hm = ''P003'''
                          + ' and C.cod_pan <> '''''
                          + ' and not (C.gubn_class = ''1'' and C.desc_sokun1 = '''')'
      end;

      if Ck_Class_High.checked = true then
      begin
         sWhere := sWhere + ' and C.gubn_class in (''9'',''10'',''11'',''12'',''13'',''14'')' ;
      end;

      sWhere := sWhere + '  GROUP BY G.num_id, G.desc_name, G.dat_gmgn, G.cod_jisa, G.num_jumin, G.cod_dc, G.desc_dept, C.cod_hm, D.desc_dc, G.gubn_neawon';

      case sorting.itemindex of
        0 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_name';
            end;
        1 : begin
              sWhere := sWhere + '  ORDER BY G.cod_dc, G.desc_dept, G.desc_name';
            end;
        2 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_dept, G.desc_name';
            end;
        3 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.cod_dc, G.desc_name';
            end;
        4 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.dat_gmgn, G.desc_name';
            end;
        5 : begin
              sWhere := sWhere + '  ORDER BY G.cod_jisa, G.desc_name';
            end;
        6 : begin
              sWhere := sWhere + '  ORDER BY G.dat_gmgn, G.desc_name';
            end;
        7 : begin
              sWhere := sWhere + '  ORDER BY G.desc_name';
            end;
      end;

      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);
      GP_query_log(frmLD1P132.qryGumgin, 'LD1I12', 'Q', 'N');
      Active := True;

      if frmLD1P132.qryGumgin.IsEmpty = False then
      begin
         if rbAll.Checked then
         begin
            frmLD1P132.QuickRep.PrinterSettings.FirstPage := 0;
            frmLD1P132.QuickRep.PrinterSettings.LastPage  := 0;
         end
         else if rbPart.Checked then
         begin
            frmLD1P132.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
            frmLD1P132.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
         end;
         Result := True;
      end
      else
      begin
         Result := False;
         exit;
      end;
   end;
end;

end.
 