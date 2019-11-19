//==============================================================================
// ���α׷� ���� : [WorkList]Ư������ �˻��׸� ����
// �ý���        : ���հ���
// ��������      : 2008.01.09
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
// ��������      : 2015.07.14
// ������        : ������
// ��������      : ��10������ �߰�
// �������      : ���� ���Ư��������Ʈ1500104
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

    // SQL��
    UV_sBasicSQL : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
    // Grid�� ������ Variant ���� ����(Report���� �����ϹǷ� Public�� ����)
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
      // Variant������ ����ϱ� ���ؼ� ����
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
      // �̹� ������ �������� ũ�⸦ ����
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

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD4Q19.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_VarMemSet(0);

   // Login ���簡 '00'�̸� '15'�� ��ü
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   if UV_sCod_jisa = '15' then
   begin
      Label1.Caption := '�� ��:';
      GP_ComboJisa(cmbCOD_JISA);
   end
   else
   begin
      Label1.Caption := '���ּ�:';
      cmbCod_jisa.Items.Add('�� �� ��');
      cmbCod_jisa.Items.Add('��    ��');
      cmbCod_jisa.ItemIndex := 1;      
   end;

   // SQL���� ����
   UV_sBasicSQL := qryBunju.SQL.Text;
end;

procedure TfrmLD4Q19.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
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

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid & Chart �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   index := 0;
   Gauge1.MinValue := 0;
   Gauge1.Progress := 0;
   with qryBunju do
   begin
      // SQL���� ������ �ڷḦ Query
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

         // DataSet�� ���� Variant������ �̵�
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            application.ProcessMessages;

            //����ڵ� ����..---------------------------------------------------
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

            //�����ڵ� ����..---------------------------------------------------
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

            //�����ڵ� ����..---------------------------------------------------
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

            //�߰��ڵ� ����..---------------------------------------------------
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

            //Ư���ڵ� ����..---------------------------------------------------
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
                  1,3,5,7,9 :  UV_v007[index] := '��';
                  2,4,6,8,0 :  UV_v007[index] := '��';
                  else UV_v007[index] := copy(FieldByName('num_jumin').AsString,7,1);
               end;

               sHangmok := '';
               if (chk_JJFH.Checked = True) and (pos('JJFH',chuga) > 0) then
                  sHangmok := sHangmok + '���㼼���˻� ';
               if (chk_JJFG.Checked = True) and (pos('JJFG',chuga) > 0) then
                  sHangmok := sHangmok + '�Һ��������� ';
               if (chk_H029.Checked = True) and (pos('H029',chuga) > 0) then
                  sHangmok := sHangmok + '���������� ';
               if (chk_H031.Checked = True) and (pos('H031',chuga) > 0) then
                  sHangmok := sHangmok + '��Ʈ���۷κҸ� ';
               if (chk_U010.Checked = True) and (pos('U001',chuga) > 0) and (pos('U002',chuga) > 0) and
                  (pos('U003',chuga) > 0) and (pos('U004',chuga) > 0) and (pos('U005',chuga) > 0) and
                  (pos('U006',chuga) > 0) and (pos('U007',chuga) > 0) and (pos('U008',chuga) > 0) and
                  (pos('U009',chuga) > 0) and (pos('U010',chuga) > 0) then
                  sHangmok := sHangmok + '��10������ ';

               UV_v008[index] := copy(sHangmok,1,length(sHangmok)-1);
               Inc(index);
            end;

            Next;
         end;
         // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
         Active := False;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
   end;

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');

end;

procedure TfrmLD4Q19.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // �ڷᰡ ��������� ó��
   if NewRow = 0 then exit;

   // Data�Ǽ� ǥ��
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
