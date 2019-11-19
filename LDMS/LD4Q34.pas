//==============================================================================
// ���α׷� ���� : ����˻� ����Ʈ
// �ý���        : �м�����
// ��������      : 2011.03.
// ������        : ���ö
// ��������      :                      
// ��������      :
//==============================================================================
// ��������      : 2014.05.14
// ������        : ������
// ��������      : ����No���۰�
// ��������      : [���� - �����]
//==============================================================================
// ��������      : 2014.06.12
// ������        : ������
// ��������      : B009�ڵ� �߰� & select�� ����
// ��������      : [���� - �����]
//==============================================================================
// ��������      : 2014.07.07
// ������        : ������
// ��������      : �������� ��ȸ �߰�
// ��������      : [���� - �����]
//==============================================================================
unit ld4q34;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges ;

type
  TfrmLD4Q34 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    cmbJisa: TComboBox;
    mskSDate: TDateEdit;
    grdMaster: TtsGrid;
    btnSDate: TSpeedButton;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    Label2: TLabel;
    CmbOrder: TComboBox;
    qryCa_desc_buwi_Max: TQuery;
    qryU_Cell: TQuery;
    Label17: TLabel;
    mskEDate: TDateEdit;
    btnEDate: TSpeedButton;
    qryPf_hm: TQuery;
    qryPf_hm2: TQuery;
    Gauge1: TGauge;
    Label4: TLabel;
    Rd_No: TRadioButton;
    Edts_No: TEdit;
    Label5: TLabel;
    Edte_No: TEdit;
    Rd_Date: TRadioButton;
    Label3: TLabel;
    Cmb_Gubun: TComboBox;
    qryCell_Old: TQuery;
    Rd_Gmgn: TRadioButton;
    mskS_gmgn: TDateEdit;
    btnS_gmgn: TSpeedButton;
    Label6: TLabel;
    mskE_gmgn: TDateEdit;
    btnE_gmgn: TSpeedButton;
    Label7: TLabel;
    Cmb_level: TComboBox;

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
    function  UF_SawonCk : Boolean;
    procedure Cmb_GubunChange(Sender: TObject);

  private
    { Private declarations }
    // Grid�� ������ ������
    UV_vCount, UV_vGubn_check, UV_vDat_gmgn, UV_vDesc_name, UV_vNum_jumin, UV_vAge, UV_vSex,
    UV_vCod_jisa, UV_vNum_id, UV_vCod_cell, UV_vCod_hm, UV_vDesc_buwi, UV_vDesc_sokun, UV_vDoctor, UV_vDesc_dc, UV_vGUBN_P012 : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_DSawonCk : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q34: TfrmLD4Q34;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q341;

{$R *.DFM}

procedure TfrmLD4Q34.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster.Refresh; //�߰�
      end;
   end;
end;

procedure TfrmLD4Q34.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCount      := VarArrayCreate([0, 0], varOleStr);
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
      UV_vDoctor     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc    := VarArrayCreate([0, 0], varOleStr);
      UV_vGUBN_P012  := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCount,      iLength);
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
      VarArrayRedim(UV_vDoctor,   iLength);
      VarArrayRedim(UV_vDesc_dc,    iLength);
      VarArrayRedim(UV_vGUBN_P012,  iLength);
   end;
end;


function TfrmLD4Q34.UF_Condition : Boolean;
begin
   Result := True;

   // �ʼ��Է� ��ȸ���� Check
   if Rd_Date.Checked then
   begin
      if (mskSDate.Text = '') or (mskEDate.Text = '') then
      begin
         GF_MsgBox('Warning', 'CONDITION');
         Result := False;
      end;
   end
   else if Rd_Gmgn.Checked then
   begin
      if (mskS_gmgn.Text = '') or (mskE_gmgn.Text = '') then
      begin
         GF_MsgBox('Warning', 'CONDITION');
         Result := False;
      end;
   end;
end;

function  TfrmLD4Q34.UF_SawonCk : Boolean;
var sSelect, sWhere : String;
begin
   Result := False;
   with dmComm.qryShare do
   begin
      Active := False;
      SQL.Clear;

      // SQL�� ����
      sSelect := 'SELECT cod_sawon, gubn_dept FROM sawon';
      sWhere  := ' WHERE cod_sawon = ''' + GV_sUserId + '''';

      SQL.Add(sSelect + sWhere);
      Active := True;

      if RecordCount > 0 then
      begin
         if copy(FieldByName('cod_sawon').AsString,1,2) = '15' then
         begin
            if (FieldByName('gubn_dept').AsString = '12') or
               (FieldByName('gubn_dept').AsString = '04') then
               Result := True
            else
               Result := False;
         end;
      end;
      // Table Disconnect
      Active := False;
   end;
end;

procedure TfrmLD4Q34.FormCreate(Sender: TObject);
begin
   inherited;
   // ������Ѱ���
//   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboJisa(cmbJisa);
   if UF_SawonCk then
      pnlJisa.Visible := True
   else
      pnlJisa.Visible := False;

   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   if copy(GV_sUserId,1,2) = '15' then //20140514 ������~
   begin
      cmbRelation.Visible := True;
      labRelation.Visible := True;
   end
   else
   begin
      cmbRelation.Visible := False;
      labRelation.Visible := False;
   end;

   // ȯ�漳��
   UP_VarMemSet(0);

   // ����(From, To)�� �������ڷ� ����
   mskSDate.Text    := GV_sToday;
   mskEDate.Text    := GV_sToday;
   mskS_gmgn.Text    := GV_sToday;
   mskE_gmgn.Text    := GV_sToday;
   CmbOrder.ItemIndex := 0;

   Cmb_Gubun.ItemIndex := 0;

   Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';

   if UF_SawonCk then
   begin
      labRelation.Visible := True;
      cmbRelation.Visible := True;
   end
   else
   begin
      labRelation.Visible := False;
      cmbRelation.Visible := False;   
   end;
end;


procedure TfrmLD4Q34.btnQueryClick(Sender: TObject);
var i, index, iAge, iRet, num : Integer;
    sSelect, sWhere, sOrderBy, sMan, sCod_hm, sSelect2, sWhere2 : String;
    vCod_profile : Variant;
    check_tk : boolean;
begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   index := 0;
   sWhere  := ''; sSelect := ''; sOrderBy := '';
   with qryCell do
   begin
      Active := False;

      // SQL�� ����
      sSelect := ' select G.cod_jisa, G.dat_gmgn, G.num_id, G.num_jumin, G.desc_name, G.num_samp, C.desc_buwi, C.cod_hm, G.gubn_jinch,';
      sSelect := sSelect + ' G.cod_jangbi, G.cod_cancer, G.cod_hul, G.cod_chuga, T.cod_prf, T.cod_chuga as TkGum_chuga, D.desc_dc, S.desc_name as doctor';
      sSelect := sSelect + ' from gumgin G with(nolock) Left outer Join Cell C with(nolock)';
      if Cmb_Gubun.ItemIndex = 0 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and (C.cod_hm = ''B001'' or C.cod_hm = ''B005'' or C.cod_hm = ''B007'' or C.cod_hm = ''P009'') '
      else if Cmb_Gubun.ItemIndex = 1 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''B001'' '
      else if Cmb_Gubun.ItemIndex = 2 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''B005'' '
      else if Cmb_Gubun.ItemIndex = 3 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''B007'' '
      else if Cmb_Gubun.ItemIndex = 4 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''P009'' '
      else
      begin
         showmessage('�˻籸���� ���õǾ����� �ʽ��ϴ�.');
         Exit;
      end;

      sSelect := sSelect + ' Left outer Join ca_request J with(nolock) On G.cod_jisa = J.cod_jisa and G.num_id = J.num_id and G.dat_gmgn = J.dat_gmgn ';
      sSelect := sSelect + ' Left outer join tkgum T with(nolock) On G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn ';
      sSelect := sSelect + ' Left outer join sawon S with(nolock) On J.cod_doctor = S.cod_sawon ';
      sSelect := sSelect + ' Inner join Danche D with(nolock) On G.cod_dc = D.cod_dc ';

      if pnlJisa.Visible then
      begin
         if Copy(cmbJisa.Text, 1, 2) <> '00' then
            sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And '
         else
            sWhere := ' WHERE ';
      end
      else
         sWhere := ' WHERE G.cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And ';

      if Rd_Date.Checked then
      begin
         sWhere := sWhere + ' G.dat_insert >= ''' + mskSDate.Text + '''';
         sWhere := sWhere + ' AND G.dat_insert <= ''' + mskEDate.Text + '''';
      end
      else if Rd_gmgn.Checked then  //20140707 ������
      begin
         sWhere := sWhere + ' G.dat_gmgn >= ''' + mskS_gmgn.Text + '''';
         sWhere := sWhere + ' AND G.dat_gmgn <= ''' + mskE_gmgn.Text + '''';
      end
      else if Rd_No.Checked then
      begin
         if (EDTS_NO.Text <> '') and (EDTE_NO.Text <> '') then
            sWhere := sWhere + '  C.desc_buwi >= ''' + Edts_No.Text + ''''
                             + '  AND C.desc_buwi <= ''' + Edte_No.Text + ''''
         else
         begin
            showmessage('���� No �� �Է��� �ּ���.');
            Edts_No.SetFocus;
            exit;
         end;
      end;
      //�������� �߰� ��û .. 2018.09.07 ����
      if Cmb_level.ItemIndex = 0 then
         sWhere := sWhere + 'and C.cell_level = ''1'' '
      else if Cmb_level.ItemIndex = 1 then
         sWhere := sWhere + 'and C.cell_level = ''2'' '
      else if Cmb_level.ItemIndex = 2 then
         sWhere := sWhere + 'and C.cell_level = ''3'' '
      else if Cmb_level.ItemIndex = 3 then
         sWhere := sWhere + 'and C.cell_level = ''4'' '
      else if Cmb_level.ItemIndex = 4 then
         sWhere := sWhere + 'and C.cell_level = ''5'' '
      else if Cmb_level.ItemIndex = 5 then
         sWhere := sWhere + 'and C.cell_level = ''6'' '
      else if Cmb_level.ItemIndex = 6 then
         sWhere := sWhere + 'and C.cell_level = ''7'' ';

      if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' Order by G.cod_jisa, G.desc_name'
      else if CmbOrder.ItemIndex = 2 then
         sOrderBy := ' Order by G.num_jumin'
      else if CmbOrder.ItemIndex = 3 then
         sOrderBy := ' Order by C.Desc_buwi'
      else if CmbOrder.ItemIndex = 4 then
         sOrderBy := ' Order by G.cod_jisa, G.gubn_jinch, G.num_samp'
      else if CmbOrder.ItemIndex = 5 then
         sOrderBy := ' Order by G.cod_jisa, G.dat_gmgn, G.num_samp'
      else if CmbOrder.ItemIndex = 6 then
         sOrderBy := ' Order by D.desc_dc, G.desc_name'
      else
         sOrderBy := ' Order by G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      Gauge1.Progress := 0;
      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD4Q34', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         // ���ǿ� �´� ������ ������ �Է�
         while not Eof do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            Label4.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;
            //�˻��׸�����
            sCod_hm := '';
            if FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //---- Ư���׸� ����...
            iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_profile);
            for i := 0 to iRet-1 do
            begin
               if Trim(vCod_profile[i]) <> '' then
               begin
                  qryPf_hm2.Active := False;
                  qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[i];
                  qryPf_hm2.Active := True;
                  if qryPf_hm2.RecordCount > 0 then
                  begin
                     while not qryPf_hm2.Eof do
                     begin
                        check_tk := True;
                        for num := 1 to round(Length(sCod_hm)/4) do
                        begin
                           if copy(sCod_hm, (num * 4) - 3,4) =
                              qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                        end;
                        if check_tk = True then
                           sCod_hm := sCod_hm + qryPf_hm2.FieldByName('COD_HM').AsString;
                        qryPf_hm2.Next;
                     end; // end of while(�׸� ó��)
                  end; // end of if
               end; //end of if
            end; //end of for
            sCod_hm := sCod_hm + FieldByName('TkGum_chuga').AsString;
            sCod_hm := sCod_hm + FieldByName('cod_chuga').AsString;

            if ((pos('B001',sCod_hm) > 0) or (pos('B005',sCod_hm) > 0) or (pos('B007',sCod_hm) > 0) or (pos('P009',sCod_hm) > 0)) and
                (Cmb_Gubun.ItemIndex = 0) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;

               Inc(index);
            end
            else if (pos('B001',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 1) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               Inc(index);
            end
            else if (pos('B005',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 2) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               Inc(index);
            end
            else if (pos('B007',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 3) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               if FieldByName('Desc_buwi').AsString = '' then
               UV_vDesc_buwi[index]  := ''
               else UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString + '/' + qryCell.FieldByName('num_id').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               Inc(index);
            end
            else if (pos('P009',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 4) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               if FieldByName('Desc_buwi').AsString = '' then
               UV_vDesc_buwi[index]  := ''
               else UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString + '/' + qryCell.FieldByName('num_id').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               Inc(index);
            end;
            Next;
         end;
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
      // Table�� Disconnect
      Active := False;
   end;

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD4Q34.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // �ڷᰡ �����ϴ��� Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data�Ǽ� ǥ��
   GP_SetDataCnt(grdMaster);

   // ��ȸ Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD4Q34.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vCount[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vCod_hm[DataRow-1];
      4 : Value := UV_vDat_gmgn[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vAge[DataRow-1];
      8 : Value := UV_vSex[DataRow-1];
      9 : Value := UV_vDoctor[DataRow-1];
     10 : Value := UV_vDesc_buwi[DataRow-1];
     11 : Value := UV_vDesc_dc[DataRow-1];
   end;
end;


procedure TfrmLD4Q34.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];

end;


procedure TfrmLD4Q34.UP_Click(Sender: TObject);
begin
   inherited;
   if Rd_Date.Checked = True then
   begin
      mskSDate.Enabled := True;
      mskEDate.Enabled := True;
      mskS_gmgn.Enabled := False;
      mskE_gmgn.Enabled := False;
      btnSDate.Enabled := True;
      btnEDate.Enabled := True;
      btnS_gmgn.Enabled := False;
      btnE_gmgn.Enabled := False;
   end
   else if Rd_Gmgn.Checked = True then
   begin
      mskSDate.Enabled := False;
      mskEDate.Enabled := False;
      mskS_gmgn.Enabled := True;
      mskE_gmgn.Enabled := True;
      btnSDate.Enabled := False;
      btnEDate.Enabled := False;
      btnS_gmgn.Enabled := True;
      btnE_gmgn.Enabled := True;
   end;

   if Sender = btnSDate then
      GF_CalendarClick(mskSDate);
   if Sender = btnEDate then
      GF_CalendarClick(mskEDate);
   if Sender = btnS_gmgn then
      GF_CalendarClick(mskS_gmgn);
   if Sender = btnE_gmgn then
      GF_CalendarClick(mskE_gmgn);
end;


procedure TfrmLD4Q34.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskSDate then UP_Click(btnSDate);
   if Sender = mskEDate then UP_Click(btnEDate);
   if Sender = mskS_gmgn then UP_Click(btnS_gmgn);
   if Sender = mskE_gmgn then UP_Click(btnE_gmgn);
end;

procedure TfrmLD4Q34.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q341 := TfrmLD4Q341.Create(Self);
   frmLD4Q341.QuickRep1.Preview;
end;

{procedure TfrmLD4Q34.cmbRelationChange(Sender: TObject);
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
  if cmbRelation.ItemIndex = 1 then
  begin
     if Application.MessageBox('����No. �� ���� �ǿ� ���ؼ� ����No. �� �ϰ��ο� �Ͻðڽ��ϱ�?', 'Ȯ��', MB_YESNO+MB_ICONquestion) = IDNO then
        Exit;

     iMax := StrToInt(Edt_Max.Text);

     for i := 0 to grdMaster.Rows - 1 do
     begin
        // �ڷᰡ ��������� ó��
        if grdMaster.Rows = 0 then
        begin
           exit;
        end;

        if (trim(UV_vDesc_buwi[i]) = '') then
        begin
           qryU_Cell.ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[i];
           qryU_Cell.ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[i];
           qryU_Cell.ParamByName('Num_jumin').AsString  := UV_vNum_jumin[i];
           //����(p012),����(p013),b009������ 2018�� 5��15�� ������..
           if (UV_vCod_hm[i] = 'B009') or (UV_vCod_hm[i] = 'P012') or (UV_vCod_hm[i] = 'P013') then
           qryU_Cell.ParamByName('dat_time').AsString   := formatdatetime('HH:NN:SS', now)
           else qryU_Cell.ParamByName('dat_time').AsString   := '';

           qryU_Cell.ParamByName('DAT_update').AsString := GV_sToday;
           qryU_Cell.ParamByName('COD_update').AsString := GV_sUserId;

           if CK_Max.Checked then    //������ 20140514~
           begin
              if (Trim(Edt_Max.Text) <> '') and
                 (UV_vCod_hm[i] = 'P012') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'S' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end
              else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'P013') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'U' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end
              else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'B005') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
              end
              else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'B009') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end;
              iMax := iMax + 1;
           end              //~������ 20140514
           else
           begin
             //P012 - 'S'�� ����
             //P013 - 'U'�� ����
             //B005 - 'T'�� ����
             //B009 - 'L'�� ����
             qryCa_desc_buwi_Max.Active := False;
             if (UV_vCod_hm[i] = 'P012') or
                (UV_vCod_hm[i] = 'P013') or
                (UV_vCod_hm[i] = 'B005') or
                (UV_vCod_hm[i] = 'B009') then  //������ 20140514
             begin
                with qryCa_desc_buwi_Max do
                begin
                   qryCa_desc_buwi_Max.Active := False;

                   if Cmb_Gubun.ItemIndex = 0 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''S'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''S'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''S'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end
                   else
                   if Cmb_Gubun.ItemIndex = 1 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''U'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''U'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''U'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end
                   else
                   if Cmb_Gubun.ItemIndex = 2 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end
                   else
                   if Cmb_Gubun.ItemIndex = 3 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end;
                   qryCa_desc_buwi_Max.SQL.Clear;
                   qryCa_desc_buwi_Max.SQL.Add(sSelect);
                   qryCa_desc_buwi_Max.Active := True;
                end;
             end;

             if Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) = '' then
             begin
                if (UV_vCod_hm[i] = 'P012') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'S' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end
                else if (UV_vCod_hm[i] = 'P013') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'U' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end
                else if (UV_vCod_hm[i] = 'B005') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                end
                else if (UV_vCod_hm[i] = 'B009') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end;
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'P012') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'S' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'P013') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'U' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'B005') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'B009') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end;
             // Table disconnect
             qryCa_desc_buwi_Max.Active := False;
           end; //end of 'CK_Max.Checked'
           qryU_Cell.ExecSQL;
           GP_query_log(qryU_Cell, 'LD4Q34', 'Q', 'Y');
        end;
     end;
     btnQuery.Click;
//     grdMaster.Repaint;
  end;
end;
}
function  TfrmLD4Q34.UF_DSawonCk : Boolean;
var sSelect, sWhere : String;
begin
   Result := False;
   with dmComm.qryShare do
   begin
      Active := False;
      SQL.Clear;

      // SQL�� ����
      sSelect := 'SELECT cod_sawon, gubn_dept FROM sawon';
      sWhere  := ' WHERE cod_sawon = ''' + GV_sUserId + '''';

      SQL.Add(sSelect + sWhere);
      Active := True;

      if RecordCount > 0 then
      begin
         if copy(FieldByName('cod_sawon').AsString,1,2) = '15' then
         begin
            if Trim(FieldByName('gubn_dept').AsString) = '' then
               Result := False
            else
               case FieldByName('gubn_dept').AsInteger of
                  12 : Result := True;
                  else
                     Result := False;
               end;
         end;
      end;
      // Table Disconnect
      Active := False;
   end;
end;

procedure TfrmLD4Q34.Cmb_GubunChange(Sender: TObject);
begin
  inherited;
   if Cmb_Gubun.ItemIndex = 0 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end
   else if Cmb_Gubun.ItemIndex = 1 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end
   else if Cmb_Gubun.ItemIndex = 2 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end
   else if Cmb_Gubun.ItemIndex = 3 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end
   else if Cmb_Gubun.ItemIndex = 4 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end;
end;
{
procedure TfrmLD4Q34.CK_MaxClick(Sender: TObject);    //20140514 ������~
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
   if CK_Max.Checked = True then
   begin
      Edt_Max.Enabled := False;
      Edt_Max.Color   := clWindow;
     //'S'�� ����
     if (Cmb_Gubun.ItemIndex = 0) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''S'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''S'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''S'' + ''' + Copy(mskSDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''S'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''S'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''S'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'U'�� �����ϴ°�
     else if (Cmb_Gubun.ItemIndex = 1) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''U'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''U'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''U'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''U'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''U'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''U'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'T'�� �����ϴ°�
     else if (Cmb_Gubun.ItemIndex = 2) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''T'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''T'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''T'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''T'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'L'�� �����ϴ°�
     else if (Cmb_Gubun.ItemIndex = 3) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''L'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''L'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''L'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''L'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end;
     Edt_Max.Text := intToStr(qryCa_desc_buwi_Max.FieldByName('RESULT').Asinteger + 1);
     qryCa_desc_buwi_Max.Active := False;
   end
   else if CK_Max.Checked = False then
   begin
      Edt_Max.Enabled := False;
      Edt_Max.Text    := '';
      Edt_Max.Color   := $00E0E0E0;
   end;
end;  //~20140514 ������
}
end.

