//==============================================================================
// ���α׷� ���� : ���� list ��Ȳ
// �ý���        : ���հ���
// ��������      : 2006.08.17
// ������        : ������
// ��������      :
// �������      :
//==============================================================================
unit LD2Q03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit;

type
  TfrmLD2Q03 = class(TfrmSingle)
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
    setting: TRadioGroup;
    CkBox_01: TCheckBox;
    qryHangmok: TQuery;
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

    // SQL��
    UV_sBasicSQL : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;

  public
    { Public declarations }
    // Grid�� ������ Variant ���� ����(Report���� �����ϹǷ� Public�� ����)
    UV_vNum_bunju, UV_vNum,
    UV_vCod_jangbi, UV_vCod_hul, UV_vCod_Cancer, UV_vCod_chuga,
    UV_vCod_etc, UV_vGubn_hulgum : Variant;
  end;

var
  frmLD2Q03: TfrmLD2Q03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q011, LD2Q021, LD2Q031;

{$R *.DFM}


procedure TfrmLD2Q03.UP_GridSet;
var i : Integer;
begin
   // Grid�� ȯ�� ����
   with grdMaster do
   begin
      // Row���� �ʱ�ȭ
      Rows := 0;

      // Column Title
      Col[1].Heading := '����';
      Col[2].Heading := '����';
      Col[3].Heading := '���';
      Col[4].Heading := '�߰��ڵ�';
      Col[5].Heading := '�߰��˻�';
      Col[6].Heading := '���ֹ�ȣ';
      Col[7].Heading := '�ο�';
      Col[8].Heading := '��������';

      // Column Alignment
      Col[7].Alignment := taRightJustify;
   end;

end;

procedure TfrmLD2Q03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum         := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_etc     := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_hulgum := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vNum,         iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_Cancer,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vCod_etc,     iLength);
      VarArrayRedim(UV_vGubn_hulgum, iLength);
   end;
end;

function TfrmLD2Q03.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD2Q03.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_GridSet;
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

procedure TfrmLD2Q03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vCod_hul[DataRow-1];
      2 : Value := UV_vCod_Cancer[DataRow-1];
      3 : Value := UV_vCod_jangbi[DataRow-1];
      4 : Value := UV_vCod_chuga[DataRow-1];
      5 : Value := UV_vCod_etc[DataRow-1];
      6 : Value := UV_vNum_bunju[DataRow-1];
      7 : Value := UV_vNum[DataRow-1];
      8 : Value := UV_vGubn_hulgum[DataRow-1];
   end;
end;

procedure TfrmLD2Q03.btnQueryClick(Sender: TObject);
var i, j, index, iNum, iRet, temp : Integer;
    sWhere, sStart, sEnd, sHul, sjangbi, sCancer, sChuga, sEtc, yh_name, chuga, shulgum, chuga_hul : String;
    vCod_chuga : Variant;

begin
   inherited;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid & Chart �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   index := 0;
   with qryBunju do
   begin
      // SQL���� ������ �ڷḦ Query
      Active := False;
      sWhere := '';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := ' WHERE A.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then
            sWhere := sWhere + ' AND A.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end
      else
      begin
         if cmbCOD_JISA.ItemIndex = 0 then
         begin
            sWhere := ' WHERE A.cod_bjjs = ''15''';
            sWhere := sWhere + ' AND A.cod_jisa = ''' + UV_sCod_jisa + '''';
         end
         else if cmbCOD_JISA.ItemIndex = 1 then
            sWhere := ' WHERE A.cod_bjjs = ''' + UV_sCod_jisa + ''''
         else
            sWhere := ' WHERE A.cod_jisa = ''' + UV_sCod_jisa + '''';
      end;

      sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';

      if Trim(mskFrom.Text) <> '' then
      begin
         sWhere := sWhere + ' AND A.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
         sWhere := sWhere + ' AND A.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
      end;

      if CkBox_01.Checked then
         sWhere := sWhere + ' AND A.gubn_part = ''01''';

      sWhere := sWhere + ' GROUP BY A.num_bunju, B.num_id, B.cod_jisa, B.dat_gmgn, B.num_jumin, ' +
                         ' B.cod_jangbi, B.cod_hul, B.cod_cancer, B.cod_chuga, ' +
                         ' B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh, B.gubn_life, B.gubn_lfyh, ' +
                         ' B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, ' +
                         ' B.gubn_gong, B.gubn_goyh, B.gubn_tkgm, B.gubn_hulgum ';

      sWhere := sWhere + ' ORDER BY A.num_bunju ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere);

      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q03', 'Q', 'N');
         UP_VarMemSet(RecordCount-1);

         qryHangmok.Active := False;
         qryHangmok.Active := True;

         for i := 0 to RecordCount - 1 do
         begin
            chuga := ''; chuga_hul := '';
            //�߰��ڵ� ����..---------------------------------------------------
            iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for j := 0 to iRet-1 do
            begin
               if CkBox_01.Checked then
               begin
                  qryHangmok.Filter := 'Cod_hm = ''' + vCod_chuga[j] + '''';
                  if qryHangmok.RecordCount > 0 then
                  begin
                     if (Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga_hul) = 0) and
                        (qryHangmok.FieldByName('GUBN_PART').AsString = '01') then
                        chuga_hul := chuga_hul + qryHangmok.FieldByName('COD_HM').AsString;
                  end;
               end
               else
               begin
                  if (copy(vCod_chuga[j],1,2) <> 'JJ') and (copy(vCod_chuga[j],1,3) <> 'DRU') then
                     chuga_hul := chuga_hul + vCod_chuga[j];
               end;
            end;

            yh_name := '';
            // �������Display
            if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
            else if FieldByName('gubn_nosin').AsString = '2' then
               yh_name := yh_name + '������' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) } + ', ';

            //Ư������Display
            if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
            begin
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := FieldByName('Num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
               qryTkgum.Active := True;
               if qryTkgum.RecordCount > 0 then
               begin
                  chuga := chuga_hul + '(Ư��: ';
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
                          if CkBox_01.Checked then
                          begin
                             if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                                (qryPf_hm.FieldByName('gubn_part').AsString = '01') and
                                (copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                                chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                          end
                          else
                          begin
                             if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                                (copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                                chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                          end;

                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;

                  iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
                  for j := 0 to iRet-1 do
                  begin
                     if CkBox_01.Checked then
                     begin
                        qryHangmok.Filter := 'Cod_hm = ''' + vCod_chuga[j] + '''';
                        if qryHangmok.RecordCount > 0 then
                        begin
                           if (Pos(qryHangmok.FieldByName('COD_HM').AsString, chuga) = 0) and
                              (qryHangmok.FieldByName('GUBN_PART').AsString = '01') then
                              chuga := chuga + qryHangmok.FieldByName('COD_HM').AsString;
                        end;
                     end
                     else
                     begin
                        if (copy(vCod_chuga[j],1,2) <> 'JJ') and (copy(vCod_chuga[j],1,3) <> 'DRU') then
                           chuga := chuga + vCod_chuga[j];
                     end;
                  end;
                  chuga := chuga + ')';                  
               end; //if (qryTkgum) end;
            end
            else
               chuga  := chuga_hul;

            if FieldByName('gubn_tkgm').AsString = '1' then
               yh_name := yh_name + 'Ư��' {UF_No_Hangmok(copy(mskDate.Text,1,4), '2', FieldByName('gubn_tkyh').AsInteger)} + ', '
            else if FieldByName('gubn_tkgm').AsString = '2' then
               yh_name := yh_name + 'Ư�����' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_tkyh').AsInteger) + ', '};

            //���κ�����Display
            if FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
            else if FieldByName('gubn_adult').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_adyh').AsInteger) + ', ';

            // ������ȯ������Display
            if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger) + ', '
            else if FieldByName('gubn_life').AsString = '2' then
               yh_name := yh_name + '������ȯ�����' {UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger) } + ', ';

            //��������Display
            if FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
            else if FieldByName('gubn_agab').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_agyh').AsInteger);

            //�����Ǻ�����Display
{            if FieldByName('gubn_gong').AsString = '1' then    07.05.17 ö. Ư���ϳ��� �ʿ����.
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
            else if FieldByName('gubn_gong').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_goyh').AsInteger);
}
            if i = 0 then
            begin
               sStart  := FieldByName('Num_bunju').AsString;
               sEnd    := FieldByName('Num_bunju').AsString;
               sHul    := FieldByName('cod_hul').AsString;
               sCancer := FieldByName('cod_Cancer').AsString;
               sJangbi := FieldByName('cod_jangbi').AsString;
               shulgum := FieldByName('gubn_hulgum').AsString;
               sChuga  := chuga;
               sEtc    := yh_name;
               iNum    := 1;
            end
            else
            begin
               if (sHul    = FieldByName('cod_hul').AsString) and
                  (sCancer = FieldByName('cod_cancer').AsString) and
                  (sJangbi = FieldByName('cod_jangbi').AsString) and
                  (shulgum = FieldByName('gubn_hulgum').AsString) and
                  (sChuga  = chuga) and
                  (sEtc    = yh_name) then
               begin
                  sEnd    := FieldByName('Num_bunju').AsString;
                  sHul    := FieldByName('cod_hul').AsString;
                  sCancer := FieldByName('cod_Cancer').AsString;
                  sJangbi := FieldByName('cod_jangbi').AsString;
                  shulgum := FieldByName('gubn_hulgum').AsString;
                  sChuga  := chuga;
                  sEtc    := yh_name;
                  iNum    := iNum + 1;
               end
               else
               begin
                  UV_vNum_bunju[index]   := sStart + ' - ' + sEnd;
                  UV_vNum[index]         := IntToStr(iNum);
                  UV_vCod_jangbi[index]  := sJangbi;
                  UV_vCod_hul[index]     := sHul;
                  UV_vCod_Cancer[index]  := sCancer;
                  UV_vCod_chuga[index]   := sChuga;
                  UV_vCod_etc[index]     := sEtc;
                  if shulgum = '3' then
                     UV_vGubn_hulgum[index] := '������+����'
                  else
                     UV_vGubn_hulgum[index] := '';
                  Inc(index);

                  sStart  := FieldByName('Num_bunju').AsString;
                  sEnd    := FieldByName('Num_bunju').AsString;
                  sHul    := FieldByName('cod_hul').AsString;
                  sCancer := FieldByName('cod_Cancer').AsString;
                  sJangbi := FieldByName('cod_jangbi').AsString;
                  shulgum := FieldByName('gubn_hulgum').AsString;
                  sChuga  := chuga;
                  sEtc    := yh_name;
                  iNum    := 1;
               end;
            end;

            if i = (RecordCount - 1) then
            begin
               UV_vNum_bunju[index]  := sStart + ' - ' + sEnd;
               UV_vNum[index]        := IntToStr(iNum);
               UV_vCod_jangbi[index] := sJangbi;
               UV_vCod_hul[index]    := sHul;
               UV_vCod_Cancer[index] := sCancer;
               UV_vCod_chuga[index]  := sChuga;
               UV_vCod_etc[index]    := sEtc;
               if shulgum = '3' then
                  UV_vGubn_hulgum[index] := '������+����'
               else
                  UV_vGubn_hulgum[index] := '';
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

procedure TfrmLD2Q03.grdMasterRowChanged(Sender: TObject; OldRow,
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

function  TfrmLD2Q03.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD2Q03.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2Q03.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD2Q03.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD2Q03.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD2Q031 := TfrmLD2Q031.Create(Self);
   case setting.ItemIndex of
      0 : frmLD2Q031.QuickRep.print;
      1 : frmLD2Q031.QuickRep.Preview;
   end;
end;

procedure TfrmLD2Q03.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;


end;

end.
