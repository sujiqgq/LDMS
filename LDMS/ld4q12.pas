//==============================================================================
// ���α׷� ���� : [WorkList]Ư������ �˻��׸� ����
// �ý���        : ���հ���
// ��������      : 2008.01.09
// ������        : ������
//==============================================================================
// ��������      : 2017.08.08
// ������        : ������
// ��������      : (���� ���Ư���˻���Ʈ1700283)
//                 1.�������� ��ġ ����(���ֹ�ȣ ������ �̵�)
//                 2.���ּ���, ���� ǥ�� ����
//                 3.���ع��� �߰�(�˻��׸�List ������ ����)
//                 4.��ȸ �� ���׻���No ������ ����
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
    // Grid�� ������ Variant ���� ����
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009 : Variant;

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
      // Variant������ ����ϱ� ���ؼ� ����
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
      // �̹� ������ �������� ũ�⸦ ����
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

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD4Q12.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_VarMemSet(0);

   // Login ���簡 '00'�̸� '15'�� ��ü
   if GV_sJisa = '00' then UV_sCod_jisa := copy(GV_sUserId,1,2)
   else                    UV_sCod_jisa := GV_sJisa;

   if UV_sCod_jisa = '15' then
   begin
      Pnl_Jisa.Caption := '����';
      GP_ComboJisa(cmbCOD_JISA);
   end
   else
   begin
      Pnl_Jisa.Caption := '���ּ�:';
      cmbCod_jisa.Items.Add('�� �� ��');
      cmbCod_jisa.Items.Add('��    ��');
      cmbCod_jisa.ItemIndex := 1;      
   end;

   Com_Part.ItemIndex := 8;

   // SQL���� ����
   UV_sBasicSQL := qryBunju.SQL.Text;
end;

procedure TfrmLD4Q12.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
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

         // DataSet�� ���� Variant������ �̵�
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            application.ProcessMessages;

            yh_name := '';

            //[2018.07.19-������]�ϻ�ȭź��(TC41)�� ȣ�����ϻ�ȭź�ҳ�(JJXE) ���� ��� ����ī����(SE42) üũ...
            bSE42_Check := False;

            //[1]��������
            UV_v001[index] := FieldByName('desc_jisa').AsString;
            //[2]���ֹ�ȣ
            UV_v002[index]  := FieldByName('num_bunju').AsString;
            //[3]��������
            UV_v003[index]  := FieldByName('dat_bunju').AsString;
            //[4]���׻��� No.
            UV_v004[index]  := FieldByName('Num_samp').AsString;
            //[5]��ü��
            UV_v005[index] := FieldByName('desc_dc').AsString;
            //[6]����
            UV_v006[index]  := FieldByName('desc_name').AsString;
            //[7]�ֹι�ȣ
            UV_v007[index] := copy(FieldByName('num_jumin').AsString,1,7) + '******';

            //����ڵ� ����..---------------------------------------------------
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
            //�����ڵ� ����..---------------------------------------------------
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
            //�����ڵ� ����..---------------------------------------------------
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
            //�߰��ڵ� ����..---------------------------------------------------
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
                          //[2018.07.19-������]�ϻ�ȭź��(TC41)�� ȣ�����ϻ�ȭź�ҳ�(JJXE) ���� ��� ����ī����(SE42) ����...
                          //----------------------------------------------------
                          if (vCod_chuga[Temp] = 'TC41') and
                             (pos('JJXE',qryTkgum.FieldByName('Cod_chuga').AsString) > 0) and
                             (qryPf_hm.FieldByName('COD_HM').AsString = 'SE42') then
                          begin
                             // ����ī����(SE42) �˻� Skip...
                             bSE42_Check := True;

                             if Chk_JJXE.Checked then
                             begin
                                //���ع�������Ʈ...
                                if trim(sMatter) = '' then  sMatter := sMatter + '�ϻ�ȭź��(CO)'
                                else
                                begin
                                   if pos(qryPf_hm.FieldByName('desc_pf').AsString, sMatter) <= 0 then
                                      sMatter := sMatter + ',' + #13 + '�ϻ�ȭź��(CO)';
                                end;
     
                                if Trim(chuga) = '' then chuga := chuga + 'JJXE'
                                else                     chuga := chuga + ',' + #13 + 'JJXE';
     
                                if Trim(sHangmok) = '' then sHangmok := sHangmok + 'ȣ�����ϻ�ȭź�ҳ�'
                                else                        sHangmok := sHangmok + ',' + #13 + 'ȣ�����ϻ�ȭź�ҳ�';
                             end;
                          end
                          else
                          begin
                             if (Pos(qryPf_hm.FieldByName('COD_HM').AsString, chuga) = 0) and
                                (qryPf_hm.FieldByName('GUBN_PART').AsString = copy(Com_Part.Text,1,2)) then
                             begin
                                //���ع�������Ʈ...
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
                  //�߰��ڵ� ����..---------------------------------------------------
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

            //[8]���ع�������Ʈ...
            UV_v008[index] := sMatter;

            // 4. ���, ���κ�, ����, ������ȯ�⿡ ���� �˻��׸� ����--------------------------
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
               //���+TC77 �̸� ����Ʈ ����. ������ ����
               if (CheckBox.Checked) and ((qryBunju.fieldByName('gubn_tkgm').AsString <> '2') and (pos('TC77',FieldByName('cod_prf').AsString) > 0)) then
               begin
               // �׳� �����...
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
                  //[2018.07.19-������]�ϻ�ȭź��(TC41)�� ȣ�����ϻ�ȭź�ҳ�(JJXE) ���� ��� ����ī����(SE42) ����...
                  //�� ���ǿ� ī���ø� �ִ� ��� �� ������ ��µǴ� �κ� ����..
                  if ((bSE42_Check) and (sHangmok = '')) or
                     (CheckBox.Checked) and ((qryBunju.fieldByName('gubn_tkgm').AsString <> '2') and (pos('TC77',FieldByName('cod_prf').AsString) > 0)) then
                  begin
                     // �׳� �����...
                  end
                  else
                     Inc(index);
               end;

               Next;
            end;

            //Next;
            //Inc(index);
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

procedure TfrmLD4Q12.grdMasterRowChanged(Sender: TObject; OldRow,
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
      Application.MessageBox('Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.',
                             '����', MB_OK or MB_ICONERROR);
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