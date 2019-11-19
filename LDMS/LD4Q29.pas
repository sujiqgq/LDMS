//==============================================================================
// ���α׷� ���� : UA Stick �˻縮��Ʈ
// �ý���        : ���հ���
// ��������      : 2011.06.22
// ������        : ����ȣ
// ��������      :
// �������      :
//==============================================================================
// ���α׷� ���� : UA Stick �˻縮��Ʈ
// �ý���        : ���հ���
// ��������      : 2011.07.06
// ������        : ����ȣ
// ��������      :
// �������      : 10���� ���� �� �ǵ���.. ��ħ��(U011~U013)�� ���ԵǸ� ���� �˻��̹Ƿ� ����Ʈ�� �� ��������..
//==============================================================================
// ���α׷� ���� : UA Stick �˻縮��Ʈ
// �ý���        : ���հ���
// ��������      : 2015.03.12
// ������        : ������
// ��������      : �ֹι�ȣ -> �������/����
// �������      : ����-�ѵη�
//==============================================================================
// ���α׷� ���� : UA Stick �˻縮��Ʈ
// �ý���        : ���հ���
// ��������      : 2017.02.15
// ������        : �ڼ���
// ��������      : UA stick �˻縮��Ʈ�� 2��+ħ��(U011, U012, U013)�� �ִ�(��:�ѱ��װ����ֻ��ä��) �׸� ��ȸ �����ϵ���
// �������      : �λ� �����
//==============================================================================
unit LD4Q29;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q29 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtDate: TDateEdit;
    SampNo1: TEdit;
    SampNo2: TEdit;
    qryGumgin: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    ComB_ShFloor: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryPf_hm2: TQuery;
    Panel2: TPanel;
    rb_nea: TRadioButton;
    rb_chul: TRadioButton;
    rb_All: TRadioButton;
    Label4: TLabel;
    Chk_four: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vDate, UV_vSamp, UV_vName, UV_vJumin, UV_vU003, UV_vU004, UV_vU005, UV_vU009,Uv_vGubn : Variant;


    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
  end;

var
  frmLD4Q29: TfrmLD4Q29;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q291;

{$R *.DFM}

procedure TfrmLD4Q29.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vDate  := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp  := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vU003  := VarArrayCreate([0, 0], varOleStr);
      UV_vU004  := VarArrayCreate([0, 0], varOleStr);
      UV_vU005  := VarArrayCreate([0, 0], varOleStr);
      UV_vU009  := VarArrayCreate([0, 0], varOleStr);
      Uv_vGubn  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,   iLength);
      VarArrayRedim(UV_vDate,  iLength);
      VarArrayRedim(UV_vSamp, iLength);
      VarArrayRedim(UV_vName, iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vU003, iLength);
      VarArrayRedim(UV_vU004, iLength);
      VarArrayRedim(UV_vU005, iLength);
      VarArrayRedim(UV_vU009, iLength);
      VarArrayRedim(Uv_vGubn, iLength);
   end;
end;

function TfrmLD4Q29.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q29.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   edtDate.Text := GV_sToday;
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);
   GP_ComboFloor_Center(ComB_ShFloor);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q29.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDate[DataRow-1];
      3 : Value := UV_vSamp[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vU003[DataRow-1];
      7 : Value := UV_vU004[DataRow-1];
      8 : Value := UV_vU005[DataRow-1];
      9 : Value := UV_vU009[DataRow-1];
      10 : Value := Uv_vGubn[DataRow-1];
   end;
end;

procedure TfrmLD4Q29.btnQueryClick(Sender: TObject);
var index, I, j, iRet, iTemp, num : Integer;
    sSelect, sWhere, sOrderBy : String;
    sHul_List : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryGumgin do
   begin
      // SQL�� ����
      Close;
      sSelect := ' SELECT G.num_jumin, G.desc_name, G.dat_gmgn, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, G.num_samp, G.gubn_jinch, G.gubn_nosin, ' +
                 ' G.gubn_adult, G.gubn_agab, G.gubn_life, G.gubn_nsyh, G.gubn_adyh, G.gubn_agyh, G.gubn_lfyh, G.gubn_tkgm, G.gubn_neawon, T.cod_prf, T.cod_chuga As Tcod_chuga ' +
                 ' FROM Gumgin G ' +
                 ' left outer join Tkgum T ' +
                 ' ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';
      sWhere  := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND G.Dat_gmgn  = ''' + edtDate.Text + ''' ';

      If SampNo1.Text <> '' Then
         sWhere := sWhere + ' And G.Num_samp >= ''' + SampNo1.Text + '''' +
                            ' And G.Num_samp <= ''' + SampNo2.Text + '''';
      IF ComB_ShFloor.Text <> '' then
         sWhere := sWhere + ' And G.gubn_jinch = ''' + copy(ComB_ShFloor.Text,1,2) + '''';

      if rb_nea.Checked then
         sWhere := sWhere + ' And G.gubn_neawon in (''1'',''3'',''4'')'
      else if rb_chul.Checked then
         sWhere := sWhere + ' And G.gubn_neawon = ''2''';
      sOrderBy := ' ORDER BY G.Num_samp';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD4Q29', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;

            // �׸�(�����������)����..
            sHul_List := '';
            if qryGumgin.FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            // �׸�(������������)����..
            if qryGumgin.FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            // �׸�(������������)����..
            if qryGumgin.FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if copy(qryPf_hm.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                        sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //�߰��˻��׸�..
            iRet := GF_MulToSingle(qryGumgin.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for iTemp := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[iTemp],1,2) <> 'JJ' then
                  sHul_List := sHul_List + vCod_chuga[iTemp];
            end;

            if qryGumgin.FieldByName('gubn_tkgm').AsString <> '' then
            begin
               //---- Ư���׸�(Profile) ����...
               iRet := GF_MulToSingle(qryGumgin.FieldByName('cod_prf').AsString, 4, vCod_profile);
               for j := 0 to iRet-1 do
               begin
                  if Trim(vCod_profile[j]) <> '' then
                  begin
                     qryPf_hm2.Active := False;
                     qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[j];
                     qryPf_hm2.Active := True;
                     if qryPf_hm2.RecordCount > 0 then
                     begin
                        while not qryPf_hm2.Eof do
                        begin
                           check_tk := True;
                           for num := 1 to round(Length(sHul_List)/4) do
                           begin
                              if copy(sHul_List, (num * 4) - 3,4) =
                                 qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                           end;
                           if check_tk = True then
                           begin
                              if copy(qryPf_hm2.FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
                                 sHul_List := sHul_List + qryPf_hm2.FieldByName('COD_HM').AsString;
                           end;
                           qryPf_hm2.Next;
                        end; // end of while(�׸� ó��)
                     end; // end of if
                  end; //end of if
               end; //end of for
               
               //(Ư��)�߰��˻��׸�..
               iRet := GF_MulToSingle(qryGumgin.FieldByName('Tcod_chuga').AsString, 4, vCod_chuga);
               for iTemp := 0 to iRet-1 do
               begin
                  if copy(vCod_chuga[iTemp],1,2) <> 'JJ' then
                     sHul_List := sHul_List + vCod_chuga[iTemp];
               end;
            end;

            // �������Display
            if FieldByName('gubn_nosin').AsString = '1' then
               sHul_List := sHul_List + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qryGumgin.FieldByName('gubn_nsyh').AsInteger);
            
            //���κ�����Display
            if FieldByName('gubn_adult').AsString = '1' then
               sHul_List := sHul_List + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qryGumgin.FieldByName('gubn_adyh').AsInteger);
            
            //��������Display
            if FieldByName('gubn_agab').AsString = '1' then
               sHul_List := sHul_List + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qryGumgin.FieldByName('gubn_agyh').AsInteger);
            
            //������ȯ�⺴����Display
            if FieldByName('gubn_life').AsString = '1' then
               sHul_List := sHul_List + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qryGumgin.FieldByName('gubn_lfyh').AsInteger);
            if Chk_four.Checked =true then
            begin
            if (pos('U003', sHul_List) > 0) and (pos('U004', sHul_List) > 0) and
               (pos('U005', sHul_List) > 0) and (pos('U009', sHul_List) > 0) then
               begin
            UP_VarMemSet(Index);
            UV_vNo[Index] := IntToStr(Index+1);
            UV_vDate[Index] := qryGumgin.FieldByName('dat_gmgn').AsString;
            UV_vSamp[Index] := qryGumgin.FieldByName('num_samp').AsString;
            UV_vName[Index] := qryGumgin.FieldByName('desc_name').AsString;
            UV_vJumin[Index] := copy(qryGumgin.FieldByName('num_jumin').AsString,1,6)+'/';
            case StrToInt(copy(qryGumgin.FieldByName('num_jumin').AsString,7,1)) of
               1,3,5,7,9 : UV_vJumin[Index] := UV_vJumin[Index] + 'M';
               2,4,6,8   : UV_vJumin[Index] := UV_vJumin[Index] + 'F';
            end ;
            if pos('U003', sHul_List) > 0 then UV_vU003[Index] := ''
            else                               UV_vU003[Index] := '*';
            if pos('U004', sHul_List) > 0 then UV_vU004[Index] := ''
            else                               UV_vU004[Index] := '*';
            if pos('U005', sHul_List) > 0 then UV_vU005[Index] := ''
            else                               UV_vU005[Index] := '*';
            if pos('U009', sHul_List) > 0 then UV_vU009[Index] := ''
            else                               UV_vU009[Index] := '*';
       //     if (pos('U011', sHul_List) > 0) or (pos('U012', sHul_List) > 0) or (pos('U013', sHul_List) > 0) then   Uv_vGubn:='ħ��'
       //     else                                                                                                   Uv_vGubn:='    ';
            if (pos('U006', sHul_List) <= 0) AND ((pos('U011', sHul_List) <= 0) OR (pos('U012', sHul_List) <= 0) OR (pos('U013', sHul_List) <= 0)) AND
               ((UV_vU003[Index] = '') and  (UV_vU004[Index] = '') and (UV_vU005[Index] = '') and (UV_vU009[Index] = ''))  then
               Inc(Index);
            end;
            end
            else if  Chk_four.Checked=false then
            begin
                 UP_VarMemSet(Index);
                 UV_vNo[Index] := IntToStr(Index+1);
//                 if qryGumgin.FieldByName('num_samp').AsString = '151098' then   showmessage('������');
                 UV_vDate[Index] := qryGumgin.FieldByName('dat_gmgn').AsString;
                 UV_vSamp[Index] := qryGumgin.FieldByName('num_samp').AsString;
                 UV_vName[Index] := qryGumgin.FieldByName('desc_name').AsString;
                 UV_vJumin[Index] := copy(qryGumgin.FieldByName('num_jumin').AsString,1,6)+'/';
                 case StrToInt(copy(qryGumgin.FieldByName('num_jumin').AsString,7,1)) of
                    1,3,5,7,9 : UV_vJumin[Index] := UV_vJumin[Index] + 'M';
                    2,4,6,8   : UV_vJumin[Index] := UV_vJumin[Index] + 'F';
                 end;
                 if pos('U003', sHul_List) > 0 then UV_vU003[Index] := ''
                 else                               UV_vU003[Index] := '*';
                 if pos('U004', sHul_List) > 0 then UV_vU004[Index] := ''
                 else                               UV_vU004[Index] := '*';
                 if pos('U005', sHul_List) > 0 then UV_vU005[Index] := ''
                 else                               UV_vU005[Index] := '*';
                 if pos('U009', sHul_List) > 0 then UV_vU009[Index] := ''
                 else                               UV_vU009[Index] := '*';
                 if (pos('U003', sHul_List) > 0) or (pos('U004', sHul_List) > 0) or (pos('U005', sHul_List) > 0) or (pos('U009', sHul_List) > 0) then
                 begin
                      if (pos('U011', sHul_List) > 0) or (pos('U012', sHul_List) > 0) or (pos('U013', sHul_List) > 0) then   Uv_vGubn[index]:='ħ��'
                      else                                                                                                   Uv_vGubn[index]:='    ';
                 end;

                 if  copy(GV_sUserId,1,2) ='61' then
                 begin
                   if  (pos('U006', sHul_List) <= 0) AND
                    //((pos('U011', sHul_List) <= 0) OR (pos('U012', sHul_List) <= 0) OR (pos('U013', sHul_List) <= 0)) AND  ..20170215 ����( ����� ����(�λ�)-��û)
                      ((UV_vU003[Index] = '') OR (UV_vU004[Index] = '') OR (UV_vU005[Index] = '') OR (UV_vU009[Index] = ''))  then
                        Inc(Index);
                 end

                 else if (pos('U006', sHul_List) <= 0) AND ((pos('U011', sHul_List) <= 0) OR (pos('U012', sHul_List) <= 0) OR (pos('U013', sHul_List) <= 0)) and
                        ((UV_vU003[Index] = '') OR (UV_vU004[Index] = '') OR (UV_vU005[Index] = '') OR (UV_vU009[Index] = ''))  then
                         Inc(Index);

            end;
            Next;
         End;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q29.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtDate);
end;

procedure TfrmLD4Q29.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q291 := TfrmLD4Q291.Create(Self);
     frmLD4Q291.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q291 := TfrmLD4Q291.Create(Self);
     frmLD4Q291.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q29.FormActivate(Sender: TObject);
begin
   inherited;
   edtdate.SetFocus;
end;

function  TfrmLD4Q29.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

end.
 