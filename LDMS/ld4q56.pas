//==============================================================================
// ���α׷� ���� : Fe&Hemoglobin, HbA1C&Glucoge ��� �� LIST
// �ý���        : �м�����
// ��������      : 2014.04.24
// ������        : ������
// ��������      : �ű԰���
//==============================================================================
// ��������      : 2014.06.21
// ������        : ������
// ��������      : C043, C003, C007, CO46, C047, C048 ��ȸ
//                 - ��� �� 0 ������ ���� ���̶���Ʈ
//                 - ����� �Է� �ȵ� ����� '*'�� ǥ��
// ��������      : [����-������]
//==============================================================================
// ��������      : 2014.07.07
// ������        : ������
// ��������      : C045 ���ڰ� ������ ���̶���Ʈ
// ��������      : [����-������]
//==============================================================================
// ��������      : 2014.11.18
// ������        : ������
// ��������      : ���׸� 3�� - ö ������ ��ȸ
// ��������      : [����-��ҿ�]
//==============================================================================
// ��������      : 2014.12.11
// ������        : ������
// ��������      : Ư�� ���� �������� ���� ��������� ����
// ��������      :
//==============================================================================
// ��������      : 2015.03.12
// ������        : ������
// ��������      : ���׸�2�� C034�ڵ� ��ȸ �߰� & ��»�� ����.
// ��������      : ����-������
//==============================================================================
// ��������      : 2015.03.13
// ������        : ������
// ��������      : ������Է�-����, ������-'*'
// ��������      : ����-������
//==============================================================================
// ��������      : 2015.06.17
// ������        : ������
// ��������      : ���׸�2 -> C001, C002 �߰���û
// ��������      : ���� ������ܰ˻�������1500297
//==============================================================================
// ��������      : 2015.09.16
// ������        : ������
// ��������      : LD ��� ��
// ��������      : ���� ������ܰ˻�������1500734
//==============================================================================
// ��������      : 2015.09.19
// ������        : ������
// ��������      : FormatFloat����
//==============================================================================
// ��������      : 2016.06.18
// ������        : �ڼ���
// ��������      : 9.LPa(C080), LIPASE(C082), Homocystein(C078) �߰�
// ��������      : ���� ������ܰ˻�������1600563
//==============================================================================
// ��������      : 2016.07.13
// ������        : �ڼ���
// ��������      : C005 �߰�
// ��������      : ���� ����������1600412
//==============================================================================
unit LD4Q56;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q56 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    qryGumgin: TQuery;
    Gride: TGauge;
    mskDate: TDateEdit;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    gubn_Comb: TComboBox;
    Label1: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryPGulkwa: TQuery;
    qryGulkwa02: TQuery;
    qryProfile: TQuery;
    RG_Order: TRadioGroup;
    RB_bunju: TRadioButton;
    RB_Samp: TRadioButton;
    qryPgulkwa05: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure UP_SetInitGrid;
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vDat_bunju, UV_vNum_bunju, UV_vNum_samp, UV_vDesc_name, UV_v001, UV_v002,
    UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010, UV_v011: Variant;
    CHK_C078 : Boolean;
    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q56: TfrmLD4Q56;
  UV_fGulkwa_01,  UV_fGulkwa1_01,  UV_fGulkwa2_01,  UV_fGulkwa3_01,
  UV_fPGulkwa_01, UV_fPGulkwa1_01, UV_fPGulkwa2_01, UV_fPGulkwa3_01,
  UV_fGulkwa_02,  UV_fGulkwa1_02,  UV_fGulkwa2_02,  UV_fGulkwa3_02 ,
  UV_fPGulkwa_05, UV_fGulkwa1_05,  UV_fGulkwa2_05,  UV_fGulkwa3_05  : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;//, LD4Q561, LD4Q562;

{$R *.DFM}

procedure TfrmLD4Q56.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_v001       := VarArrayCreate([0, 0], varOleStr);
      UV_v002       := VarArrayCreate([0, 0], varOleStr);
      UV_v003       := VarArrayCreate([0, 0], varOleStr);
      UV_v004       := VarArrayCreate([0, 0], varOleStr);
      UV_v005       := VarArrayCreate([0, 0], varOleStr);
      UV_v006       := VarArrayCreate([0, 0], varOleStr);
      UV_v007       := VarArrayCreate([0, 0], varOleStr);
      UV_v008       := VarArrayCreate([0, 0], varOleStr);
      UV_v009       := VarArrayCreate([0, 0], varOleStr);
      UV_v010       := VarArrayCreate([0, 0], varOleStr);
      UV_v011       := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_v001,       iLength);
      VarArrayRedim(UV_v002,       iLength);
      VarArrayRedim(UV_v003,       iLength);
      VarArrayRedim(UV_v004,       iLength);
      VarArrayRedim(UV_v005,       iLength);
      VarArrayRedim(UV_v006,       iLength);
      VarArrayRedim(UV_v007,       iLength);
      VarArrayRedim(UV_v008,       iLength);
      VarArrayRedim(UV_v009,       iLength);
      VarArrayRedim(UV_v010,       iLength);
      VarArrayRedim(UV_v011,       iLength);
   end;
end;

function TfrmLD4Q56.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q56.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
//   UP_GridSet;
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);
   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   gubn_comb.ItemIndex := 0;
   UP_SetInitGrid;
end;

procedure TfrmLD4Q56.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // �ڷḦ ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vNo[DataRow-1];
       2 : Value := UV_vDat_bunju[DataRow-1];
       3 : Value := UV_vNum_bunju[DataRow-1];
       4 : Value := UV_vNum_samp[DataRow-1];
       5 : Value := UV_vDesc_name[DataRow-1];
       6 : Value := UV_v001[DataRow-1];
       7 : Value := UV_v002[DataRow-1];
       8 : Value := UV_v003[DataRow-1];
       9 : Value := UV_v004[DataRow-1];
      10 : Value := UV_v005[DataRow-1];
      11 : Value := UV_v006[DataRow-1];
      12 : Value := UV_v007[DataRow-1];
      13 : Value := UV_v008[DataRow-1];
      14 : Value := UV_v009[DataRow-1];
      15 : Value := UV_v010[DataRow-1];
      16 : Value := UV_v011[DataRow-1];
   end;

   if gubn_Comb.ItemIndex = 1 then
   begin
     if (DataCol = 6) and (trim(Value) <> '') and (trim(Value) <> '*') and ((Value <= 4.3) or (Value >= 6.5)) then
        grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end
   {else if gubn_Comb.ItemIndex = 2 then
   begin
     if ((DataCol =  6) or (DataCol =  7) or (DataCol =  8) or (DataCol = 10) or
         (DataCol = 11) or (DataCol = 12) or (DataCol = 13)) AND
        (trim(Value) <> '') and (Value <> '*') and (Value <= 0.0) then
        grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end}
   else if gubn_Comb.ItemIndex = 2 then
   begin
       if (trim(Value) <> '') and (Value <> '*') then
       begin
           if ((DataCol = 6) and ((Value <= 6.0) or (Value >= 8.3))) or
              ((DataCol = 7) and ((Value <= 3.5) or (Value >= 5.2))) or
              ((DataCol = 8) and ((Value <= 1.5) or (Value >= 3.0))) or
              ((DataCol = 9) and ((Value <= 0.2) or (Value >= 0.7))) or
              ((DataCol = 10) and ((Value <= 10.0) or (Value >= 20.0))) or
              ((DataCol = 11) and ((Value <= 60) or (Value >= 175))) or
              ((DataCol = 12) and ((Value <= 60) or (Value >= 175))) or
              ((DataCol = 13) and ((Value <= 250) or (Value >= 400))) or
              ((DataCol = 14) and ((Value <= 20) or (Value >= 55))) or
              ((DataCol = 15) and ((Value <= 70) or (Value >= 310))) or
              ((DataCol = 16) and ((Value <= 0.2) or (Value >= 1.2)))
              then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
       end;
   end
   else if gubn_Comb.ItemIndex = 5 then
   begin
       if (trim(Value) <> '') and (Value <> '*') then
       begin
           if ((DataCol = 6) and ((Value <= 134) or (Value >= 146))) or
              ((DataCol = 7) and ((Value <= 3.4) or (Value >= 5.6))) or
              ((DataCol = 8) and ((Value <= 97)  or (Value >= 111))) or
              ((DataCol = 9) and ((Value <= 1.5) or (Value >= 2.7))) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
       end;
   end
   else if gubn_Comb.ItemIndex = 7 then
   begin
        if (UV_v001[Datarow -1] <> UV_v002[Datarow -1]) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end
   else if gubn_Comb.ItemIndex = 8 then
   begin
       if (trim(Value) <> '') and (Value <> '*') then
       begin
           if ((DataCol = 6) and ((Value <= 0) or (Value >= 30))) or
              ((DataCol = 7) and ((Value <= 13) or (Value >= 60))) or
              ((DataCol = 8) and ((Value <= 5)  or (Value >= 15))) or
              ((DataCol = 9) and ((Value <= 251)  or (Value >= 300))) or
              ((DataCol = 10) and ((Value <= 2199)  or (Value >= 10000))) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
       end;
   end
   else if gubn_Comb.ItemIndex = 9 then
   begin
        if (UV_v001[Datarow -1] = UV_v002[Datarow -1]) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end
   else if gubn_Comb.ItemIndex = 10 then
   begin
        if (UV_v002[Datarow -1] = '����') then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end
   else if gubn_Comb.ItemIndex = 12 then
   begin
        if (UV_v001[Datarow -1] <> UV_v002[Datarow -1]) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end;

end;

procedure TfrmLD4Q56.btnQueryClick(Sender: TObject);
var index, DataCol, I, e25, e26, e28 : Integer;
    eRslt : Real;
    chk_C078 : boolean;
    sSelect, sWhere, sOrderBy, sHangmokList: String;
begin
   inherited;
    sSelect := ''; sWhere := ''; sOrderBy := '';
    chk_C078 := false;
   // ��ȸ���� Check
   if not UF_Condition then exit;
//   UP_GridSet;
   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   with qryGumgin do
   begin
      // SQL�� ����
      Close;
      sSelect := ' Select A.Dat_gmgn, A.desc_name, B.cod_jisa, B.dat_bunju, B.num_id, B.num_bunju, B.DESC_GLKWA1, B.DESC_GLKWA2, B.DESC_GLKWA3,   ' +
                 '        A.num_jumin, A.num_samp, A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.gubn_hulgum, ' +
                 '        A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga    ';

      sSelect := sSelect + ' From Gumgin A with(nolock) inner join Gulkwa B with(nolock)                    ';
      sSelect := sSelect + ' On a.cod_jisa = B.cod_jisa and a.num_id = B.num_id and a.dat_gmgn = B.dat_gmgn ';

      sWhere := ' WHERE B.Dat_Bunju = ''' + mskDate.Text + ''' ';

      sWhere := sWhere + 'AND  B.Gubn_Part = ''01'' ';

      if RB_bunju.Checked then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else if RB_Samp.Checked then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;

      if (gubn_Comb.ItemIndex  = 7) and (copy(cmbJisa.Text,1,2) <> '15') then
      begin
         sWhere := sWhere + ' AND B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' AND (A.gubn_hulgum = ''2'' or A.gubn_hulgum = ''3'') ';
      end
      else if gubn_Comb.ItemIndex <> 3 then //20150916
      begin
         sWhere := sWhere + ' AND B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      end
      else
      begin
         if copy(cmbJisa.Text,1,2) = '15' then
         begin
            sWhere := sWhere + ' AND (A.gubn_hulgum = ''1'' or (A.gubn_hulgum = ''3'' and A.gubn_nosin <> ''1'' and A.gubn_adult <> ''1'' and A.gubn_life <> ''1'')) ';
         end
         else
         begin
            sWhere := sWhere + ' AND B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' AND (A.gubn_hulgum = ''2'' or A.gubn_hulgum = ''3'')';
         end
      end;

      if RG_Order.ItemIndex = 0 then sOrderBy := ' ORDER BY B.num_bunju '
      else sOrderBy := ' ORDER BY  B.cod_jisa, A.num_samp ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD4Q56', 'Q', 'N');
         For I := 1 to RecordCount do
         begin

            Gride.Progress := I;
            UP_SetInitGrid;
            sHangmokList := '';
            UV_fPGulkwa_01 := '';
            sHangmokList := UF_hangmokList;
            e25 := 0; e26 := 0; e28 := 0; eRslt := 0;

            UV_fGulkwa_01 := '';
            UV_fGulkwa1_01 := qryGumgin.FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2_01 := qryGumgin.FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3_01 := qryGumgin.FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01, UV_fGulkwa_01);

            if gubn_Comb.ItemIndex = 0 then  //������ ��Ʈ ���
            begin
               with qryGulkwa02 do
               begin
                  qryGulkwa02.close;
                  qryGulkwa02.ParamByName('dat_gmgn').AsString := qryGumgin.FieldByName('dat_gmgn').AsString;
                  qryGulkwa02.ParamByName('num_id').AsString   := qryGumgin.FieldByName('num_id').AsString;
                  qryGulkwa02.open;

                  if qryGulkwa02.RecordCount > 0 then
                  begin
                     UV_fGulkwa_02 := '';
                     UV_fGulkwa1_02 := qryGulkwa02.FieldByName('DESC_GLKWA1').AsString;
                     UV_fGulkwa2_02 := qryGulkwa02.FieldByName('DESC_GLKWA2').AsString;
                     UV_fGulkwa3_02 := qryGulkwa02.FieldByName('DESC_GLKWA3').AsString;
                     GF_HulGulkwa('2', UV_fGulkwa1_02, UV_fGulkwa2_02, UV_fGulkwa3_02, UV_fGulkwa_02);
                  end;
               end
            end
            else if (gubn_Comb.ItemIndex = 2) OR (gubn_Comb.ItemIndex = 12) then  //��ȭ�� ��Ʈ ���� ���
            begin
               with qryPGulkwa do
               begin
                  qryPGulkwa.close;
                  qryPGulkwa.ParamByName('dat_gmgn').AsString := qryGumgin.FieldByName('dat_gmgn').AsString;
                  qryPGulkwa.ParamByName('num_id').AsString   := qryGumgin.FieldByName('num_id').AsString;
                  qryPGulkwa.open;

                  if qryPGulkwa.RecordCount > 0 then
                  begin
                     UV_fPGulkwa_01 := '';
                     UV_fPGulkwa1_01 := qryPGulkwa.FieldByName('DESC_GLKWA1').AsString;
                     UV_fPGulkwa2_01 := qryPGulkwa.FieldByName('DESC_GLKWA2').AsString;
                     UV_fPGulkwa3_01 := qryPGulkwa.FieldByName('DESC_GLKWA3').AsString;
                     GF_HulGulkwa('2', UV_fPGulkwa1_01, UV_fPGulkwa2_01, UV_fPGulkwa3_01, UV_fPGulkwa_01);
                  end;
               end
            end
            else if gubn_Comb.ItemIndex = 10 then  //EIA ��Ʈ ���� ���  ..20180425 ����
            begin
               with qryPGulkwa05 do
               begin
                  qryPGulkwa05.close;
                  qryPGulkwa05.ParamByName('dat_gmgn').AsString := qryGumgin.FieldByName('dat_gmgn').AsString;
                  qryPGulkwa05.ParamByName('num_id').AsString   := qryGumgin.FieldByName('num_id').AsString;
                  qryPGulkwa05.open;

                  if qryPGulkwa05.RecordCount > 0 then
                  begin
                     UV_fPGulkwa_05 := '';
                     UV_fGulkwa1_05 := qryPGulkwa05.FieldByName('DESC_GLKWA1').AsString;
                     UV_fGulkwa2_05 := qryPGulkwa05.FieldByName('DESC_GLKWA2').AsString;
                     UV_fGulkwa3_05 := qryPGulkwa05.FieldByName('DESC_GLKWA3').AsString;
                     GF_HulGulkwa('2', UV_fGulkwa1_05, UV_fGulkwa2_05, UV_fGulkwa3_05, UV_fPGulkwa_05);
                  end;
               end
            end;

            if (gubn_Comb.ItemIndex = 0) then
            begin
                if (pos('C045',sHangmokList) > 0) and (pos('H002',sHangmokList) > 0) and (UV_fGulkwa_02 <> '') then
                begin
                   if (Trim(Copy(UV_fGulkwa_01,  199,  6)) <> '') and
                      ((StrToFloat(Trim(Copy(UV_fGulkwa_01,  199,  6))) <= 20) or
                       (StrToFloat(Trim(Copy(UV_fGulkwa_01,  199,  6))) >= 320)) then
                   begin
                       UP_VarMemSet(index);
                       UV_vNo[index]       := IntToStr(Index+1);
                       UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                       UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                       UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                       UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                       UV_v001[index] := Trim(Copy(UV_fGulkwa_01, 199,  6));
                       UV_v002[index] := Trim(Copy(UV_fGulkwa_02,   7,  6));
                       Inc(index);
                   end;
                   UV_fGulkwa_02 := '';
                end;
            end
            else if (gubn_Comb.ItemIndex = 1) then
            begin
                if (pos('C083',sHangmokList) > 0) OR (pos('C034',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   IF (pos('C083',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  385,  6))='') THEN
                       UV_v001[index] := '*'
                   ELSE  UV_v001[index] := Trim(Copy(UV_fGulkwa_01,  385,  6));
                   IF (pos('C032',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  157,  6))='') THEN
                       UV_v002[index] := '*'
                   ELSE  UV_v002[index] := Trim(Copy(UV_fGulkwa_01,  157,  6));
                   IF (pos('C034',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  169,  6))='') THEN
                       UV_v003[index] := '*'
                   ELSE  UV_v003[index] := Trim(Copy(UV_fGulkwa_01,  169,  6));
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 2) then  //20140621 ������          C043, C003, C007, C045, C046, C047, C048 +C001,C002
            begin
                if (pos('C001',sHangmokList) > 0) or (pos('C002',sHangmokList) > 0) or
                   (pos('C003',sHangmokList) > 0) or (pos('C007',sHangmokList) > 0) or (pos('C043',sHangmokList) > 0) or
                   (pos('C046',sHangmokList) > 0) or (pos('C047',sHangmokList) > 0) or (pos('C048',sHangmokList) > 0) or
                   (pos('C005',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if   (pos('C001',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,   1,  6))='') then
                        UV_v001[index]     := '*'
                   else UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,   1,  6));
                   if   (pos('C002',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,   7,  6))='') then
                        UV_v002[index]     := '*'
                   else UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,   7,  6));

                   if   (pos('C003',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  13,  6))='') then
                        UV_v003[index]     := '*'
                   else UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  13,  6));

                   if   (pos('C007',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  37,  6))='') then
                        UV_v004[index]     := '*'
                   else UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  37,  6));

                   if   (pos('C043',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  193,  6))='') then
                        UV_v005[index]     := '*'
                   else UV_v005[index]     := Trim(Copy(UV_fGulkwa_01,  193,  6));

                   if   UV_fPGulkwa_01 <> '' then
                        UV_v006[index]     := Trim(Copy(UV_fPGulkwa_01,  199,  6));

                   if   (pos('C045',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  199,  6))='') then //���� Fe
                        UV_v007[index]     := '*'
                   else UV_v007[index]     := Trim(Copy(UV_fGulkwa_01,  199,  6));

                   if   (pos('C046',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  205,  6))='') then
                        UV_v008[index]     := '*'
                   else UV_v008[index]     := Trim(Copy(UV_fGulkwa_01,  205,  6));

                   if   (pos('C047',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  211,  6))='') then
                        UV_v009[index]     := '*'
                   else UV_v009[index]     := Trim(Copy(UV_fGulkwa_01,  211,  6));

                   if   (pos('C048',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  217,  6))='') then
                        UV_v010[index]     := '*'
                   else UV_v010[index]     := Trim(Copy(UV_fGulkwa_01,  217,  6));

                   if   (pos('C005',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  25,  6))='') then
                        UV_v011[index]     := '*'
                   else UV_v011[index]     := Trim(Copy(UV_fGulkwa_01,  25,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 3) then //20150313
            begin
                if (pos('C027',sHangmokList)>0) and
                   (pos('C025',sHangmokList)>0) and
                   (pos('C028',sHangmokList)>0) then
                begin
                if ((pos('C025',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 121, 6))<>'') and
                    (StrToInt(Trim(Copy(UV_fGulkwa_01, 121,  6)))>= 300)) OR
                   ((pos('C028',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 139, 6))<>'') and
                    (StrToInt(Trim(Copy(UV_fGulkwa_01, 139,  6)))>= 400)) then
                   begin
                       UP_VarMemSet(index);
                       UV_vNo[index]       := IntToStr(Index+1);
                       UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                       UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                       UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                       UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                       if   (pos('C025',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  121,  6))='') then
                            UV_v001[index]     := '*'
                       else if Trim(Copy(UV_fGulkwa_01,  121,  6)) <> '' then
                       begin
                            UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  121,  6));
                            e25 := StrToInt(Trim(Copy(UV_fGulkwa_01,  121,  6)));
                       end;
                       if   (pos('C026',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  127,  6))='') then
                            UV_v002[index]     := '*'
                       else if Trim(Copy(UV_fGulkwa_01,  127,  6)) <> '' then
                       begin
                            UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  127,  6));
                            e26 := StrToInt(Trim(Copy(UV_fGulkwa_01,  127,  6)));
                       end;
                       if   (pos('C027',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  133,  6))='') then
                            UV_v003[index]     := '*'
                       else if Trim(Copy(UV_fGulkwa_01,  133,  6)) <> '' then
                       begin
                            UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  133,  6));
                       end;
                       if   (pos('C028',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  139,  6))='') then
                            UV_v004[index]     := '*'
                       else if Trim(Copy(UV_fGulkwa_01,  139,  6)) <> '' then
                       begin
                            UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  139,  6));
                            e28 := StrToInt(Trim(Copy(UV_fGulkwa_01,  139,  6)));
                       end;
                       if (e25 > 0) and (e26 > 0) and (e28 > 0) then
                       begin
                           eRslt := e25 - ((e28/5) + e26);
                           eRslt := trunc(eRslt + 0.5);
                           UV_v005[index] :=  eRslt;
                       end;
                       if UV_v003[index] <> '' then
                       begin
                          if UV_v005[index] = UV_v003[index] then UV_v006[index] := 'O'
                          else UV_v006[index] := 'X';
                       end;

                   Inc(index);
                   end;
                end;
            end
            else if (gubn_Comb.ItemIndex = 4) then //20150617
            begin
                if ((pos('C009',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 49, 6))<>'') and
                    ((StrToInt(Trim(Copy(UV_fGulkwa_01, 49,  6)))>= 50) or (StrToInt(Trim(Copy(UV_fGulkwa_01, 49,  6)))<= 300))) OR
                   ((pos('C010',sHangmokList)>0) and (Trim(Copy(UV_fGulkwa_01, 55, 6))<>'') and
                    ((StrToInt(Trim(Copy(UV_fGulkwa_01, 55,  6)))>= 50) or (StrToInt(Trim(Copy(UV_fGulkwa_01, 55,  6)))<= 300))) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if   (pos('C009',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  49,  6))='') then
                        UV_v001[index]     := '*'
                   else UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  49,  6));
                   if   (pos('C010',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  55,  6))='') then
                        UV_v002[index]     := '*'
                   else UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  55,  6));
                   if   (pos('C011',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  61,  6))='') then
                        UV_v003[index]     := '*'
                   else UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  61,  6));
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 5) then //20150617
            begin
                if (pos('C050',sHangmokList) > 0) or (pos('C051',sHangmokList) > 0) or
                   (pos('C052',sHangmokList) > 0) or (pos('C058',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if   (pos('C050',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  223,  6))='') then
                        UV_v001[index]     := '*'
                   else UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  223,  6));
                   if   (pos('C051',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  229,  6))='') then
                        UV_v002[index]     := '*'
                   else UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  229,  6));
                   if   (pos('C052',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  235,  6))='') then
                        UV_v003[index]     := '*'
                   else UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  235,  6));
                   if   (pos('C058',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  259,  6))='') then
                        UV_v004[index]     := '*'
                   else UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  259,  6));
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 6) then //20150617
            begin
                if ((pos('C005',sHangmokList)> 0) and (Trim(Copy(UV_fGulkwa_01, 25, 6)) <> '') and
                    ((StrToFloat(Trim(Copy(UV_fGulkwa_01, 25,  6))) = 0.0))) OR
                   ((pos('C006',sHangmokList)> 0) and (Trim(Copy(UV_fGulkwa_01, 31, 6)) <> '') and
                    ((StrToFloat(Trim(Copy(UV_fGulkwa_01, 31,  6))) = 0.0))) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if   (pos('C005',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  25,  6))='') then
                        UV_v001[index]     := '*'
                   else UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  25,  6));
                   if   (pos('C006',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  31,  6))='') then
                        UV_v002[index]     := '*'
                   else UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  31,  6));
                   if   (pos('C007',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  37,  6))='') then
                        UV_v003[index]     := '*'
                   else UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  37,  6));
                   Inc(index);
                end;
            end
            else if gubn_Comb.ItemIndex = 7 then
            begin
                if (pos('C074',sHangmokList) > 0)  then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   //20170826 ����
                   IF (pos('C074',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  331,  6)) = '') THEN
                       UV_v001[index] := '*'
                   ELSE  UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  331,  6));
                   IF (pos('C027',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  133,  6))='') THEN
                       UV_v002[index] := '*'
                   ELSE  UV_v002[index] := Trim(Copy(UV_fGulkwa_01,  133,  6));

                   //����
                   if   (pos('C025',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  121,  6)) <> '') then
                          e25 := StrToInt(Trim(Copy(UV_fGulkwa_01,  121,  6)));
                   if   (pos('C026',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  127,  6)) <> '') then
                          e26 := StrToInt(Trim(Copy(UV_fGulkwa_01,  127,  6)));
                   if   (pos('C028',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  139,  6)) <> '') then
                          e28 := StrToInt(Trim(Copy(UV_fGulkwa_01,  139,  6)));
                   if (e25 > 0) and (e26 > 0) and (e28 > 0) then
                   begin
                      { eRslt := e25 - ((e28/5) + e26);
                       eRslt := eRslt * 10;
                       eRslt := trunc(eRslt + 0.5);
                       eRslt := eRslt / 10;
                       UV_v003[index] := FormatFloat('###.0', eRslt);}
                       //2018.01.14 ���� ���� ���������� ���� ..����
                       eRslt := e25 - ((e28/5) + e26);
                       eRslt := trunc(eRslt + 0.5);
                       UV_v003[index] := eRslt;
                   end;


                   Inc(index);
                end;
            end
            else if gubn_Comb.ItemIndex = 8 then //20160618
            begin
                if (pos('C080',sHangmokList) > 0) or (pos('C082',sHangmokList) > 0) or
                   (pos('C078',sHangmokList) > 0) or (pos('C084',sHangmokList) > 0) or
                   (pos('C085',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if   (pos('C080',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  367,  6)) = '')  then
                        UV_v001[index]     := '*'
                   else UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  367,  6));
                   if   (pos('C082',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  379,  6)) = '') then
                        UV_v002[index]     := '*'
                   else UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  379,  6));
                   if   (pos('C078',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  355,  6)) = '') then
                        UV_v003[index]     := '*'
                   else UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  355,  6));
                   if   (pos('C084',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  391,  6)) = '') then
                        UV_v004[index]     := '*'
                   else UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  391,  6));
                   if   (pos('C085',sHangmokList) > 0) and (Trim(Copy(UV_fGulkwa_01,  397,  6)) = '') then
                        UV_v005[index]     := '*'
                   else UV_v005[index]     := Trim(Copy(UV_fGulkwa_01,  397,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 9) then
            begin
                if (pos('C078',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   IF (pos('C078',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  355,  6))='') THEN
                       UV_v001[index] := '*'
                   ELSE  UV_v001[index] := Trim(Copy(UV_fGulkwa_01,  355,  6));
                   IF (pos('C017',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  85,  6))='') THEN
                       UV_v002[index] := '*'
                   ELSE  UV_v002[index] := Trim(Copy(UV_fGulkwa_01,  85,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 10) then
            begin
                if (pos('C090',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   IF ((Trim(Copy(UV_fPGulkwa_05,  43,  6)) = '�缺') or (Trim(Copy(UV_fPGulkwa_05,  43,  6)) = '��缺')) THEN
                   begin
                       UV_v001[index] := Trim(Copy(UV_fPGulkwa_05,  43,  6));
                       UV_v002[index] := Trim(Copy(UV_fGulkwa_01,  433,  6));
                       Inc(index);
                   end;


                end;
            end
            else if (gubn_Comb.ItemIndex = 11) then
            begin
                if (pos('C090',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   IF (pos('C090',sHangmokList) > 0) AND (Trim(Copy(UV_fGulkwa_01,  433,  6)) = '�缺') THEN
                   begin
                       UV_v001[index] := Trim(Copy(UV_fGulkwa_01,  433,  6));
                       Inc(index);
                   end;
                end;
            end
            else if (gubn_Comb.ItemIndex = 12) then
            begin
                if (pos('C090',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGumgin.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGumgin.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGumgin.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGumgin.FieldByName('desc_name').AsString;
                   if ((Trim(Copy(UV_fPGulkwa_01,  433,  6)) <> '')) THEN
                   begin
                      UV_v001[index] := Trim(Copy(UV_fGulkwa_01,  433,  6));
                      UV_v002[index] := Trim(Copy(UV_fPGulkwa_01,  433,  6));

                      Inc(index);
                   end;
                end;
            end;
            next;
         end//begin of while
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGumgin

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q56.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate)
   else if Sender = RB_Bunju then
   begin
      mskFrom.Enabled    := True;
      mskTo.Enabled      := True;
      MEdt_SampS.Enabled := False;
      MEdt_SampE.Enabled := False;
   end
   else if Sender = RB_Samp then
   begin
      mskFrom.Enabled    := False;
      mskTo.Enabled      := False;
      MEdt_SampS.Enabled := True;
      MEdt_SampE.Enabled := True;
   end;
end;

procedure TfrmLD4Q56.UP_SetInitGrid;
begin
   inherited;
      grdMaster.Col[8].Width   := 0;
      grdMaster.Col[9].Width   := 0;
      grdMaster.Col[10].Width  := 0;
      grdMaster.Col[11].Width  := 0;
      grdMaster.Col[12].Width  := 0;
      grdMaster.Col[13].Width  := 0;
      grdMaster.Col[14].Width  := 0;
      grdMaster.Col[15].Width  := 0;
      grdMaster.Col[16].Width  := 0;

      if (gubn_Comb.ItemIndex = 0) then
      begin
         grdMaster.Col[6].Heading := 'Fe';
         grdMaster.Col[7].Heading := 'Hemoglobin';
      end
      else if (gubn_Comb.ItemIndex = 1) then
      begin
         grdMaster.Col[6].Heading := 'HbA1C';
         grdMaster.Col[7].Heading := 'Glucose';
         grdMaster.Col[8].Heading := 'Fructosamine';
         grdMaster.Col[8].Width   := 100;
      end
      else if (gubn_Comb.ItemIndex = 2) then
      begin
         grdMaster.Col[6].Heading := '�Ѵܹ�';
         grdMaster.Col[7].Heading := '�˺ι�';
         grdMaster.Col[8].Heading := 'Globulin';
         grdMaster.Col[9].Heading := 'Indiract billirubin';
         grdMaster.Col[10].Heading:= 'BUN/Cr';
         grdMaster.Col[11].Heading:= 'Fe(����)';
         grdMaster.Col[12].Heading:= 'Fe';
         grdMaster.Col[13].Heading:= 'TIBC';
         grdMaster.Col[14].Heading:= 'ö��ȭ��';
         grdMaster.Col[15].Heading:= 'UIBC';
         grdMaster.Col[16].Heading:= '�Ѻ������';
         grdMaster.Col[8].Width   := 60;
         grdMaster.Col[9].Width   := 90;
         grdMaster.Col[10].Width  := 60;
         grdMaster.Col[11].Width  := 60;
         grdMaster.Col[12].Width  := 60;
         grdMaster.Col[13].Width  := 50;
         grdMaster.Col[14].Width  := 60;
         grdMaster.Col[15].Width  := 50;
         grdMaster.Col[16].Width  := 60;
      end
      else if (gubn_Comb.ItemIndex = 3) then
      begin
         grdMaster.Col[6].Heading  := '�ݷ����׷�';
         grdMaster.Col[7].Heading  := 'HDL-�ݷ����׷�';
         grdMaster.Col[8].Heading  := 'LDL-�ݷ����׷�';
         grdMaster.Col[9].Heading  := '�߼�����';
         grdMaster.Col[10].Heading := '����';
         grdMaster.Col[11].Heading := '��ġ����';
         grdMaster.Col[8].Width   := 100;
         grdMaster.Col[9].Width   := 60;
         grdMaster.Col[10].Width  := 60;
         grdMaster.Col[11].Width  := 60;
      end
      else if (gubn_Comb.ItemIndex = 4) then
      begin
         grdMaster.Col[6].Heading := 'AST(SGOT)';
         grdMaster.Col[7].Heading := 'ALT(SGPT)';
         grdMaster.Col[8].Heading := 'r-GTP';
         grdMaster.Col[8].Width   := 60;
      end
      else if (gubn_Comb.ItemIndex = 5) then
      begin
         grdMaster.Col[6].Heading := 'Na+(��Ʈ��)';
         grdMaster.Col[7].Heading := 'K+(Į��)';
         grdMaster.Col[8].Heading := 'Cl(����)';
         grdMaster.Col[9].Heading := 'Mg';
         grdMaster.Col[8].Width   := 60;
         grdMaster.Col[9].Width   := 60;
      end
      else if (gubn_Comb.ItemIndex = 6) then
      begin
         grdMaster.Col[6].Heading := '�Ѻ������';
         grdMaster.Col[7].Heading := '�����������';
         grdMaster.Col[8].Heading := '�����������';
         grdMaster.Col[8].Width   := 100;
      end
      else if (gubn_Comb.ItemIndex = 7) then
      begin
         grdMaster.Col[6].Heading := 'LDL-ȿ�ҹ�';
         grdMaster.Col[7].Heading := 'LDL-�ݷ����׷�';
         grdMaster.Col[8].Heading := 'LDL-�ݷ����׷� ����';
         grdMaster.Col[8].Width   := 100;
      end
      else if (gubn_Comb.ItemIndex = 8) then
      begin
         grdMaster.Col[6].Heading  := 'LPa(C080)';
         grdMaster.Col[7].Heading  := 'LIPASE(C082)';
         grdMaster.Col[8].Heading  := 'Homocystein(C078)';
         grdMaster.Col[9].Heading  := 'd-ROMs(C084)';
         grdMaster.Col[10].Heading := 'BAP(C085)';
         grdMaster.Col[8].Width    := 60;
         grdMaster.Col[9].Width    := 60;
         grdMaster.Col[10].Width   := 60;
      end
      else if (gubn_Comb.ItemIndex = 9) then
      begin
         grdMaster.Col[6].Heading := 'Homocysteine';
         grdMaster.Col[7].Heading := '���';
      end
      else if (gubn_Comb.ItemIndex = 10) then
      begin
         grdMaster.Col[6].Heading := 'H.pylori(S021)';
         grdMaster.Col[7].Heading := 'H.pylori(C090)';
      end
      else if (gubn_Comb.ItemIndex = 11) then
      begin
         grdMaster.Col[6].Heading := 'H.pylori(C090)';
         grdMaster.Col[7].Width := 0;
      end
      else if (gubn_Comb.ItemIndex = 12) then
      begin
         grdMaster.Col[6].Heading := 'H.pylori(C090)';
         grdMaster.Col[7].Heading := '���� H.pylori(C090)';
      end;
end;

function TfrmLD4Q56.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. ���׿� ���� �˻��׸� ����
      if qryGumgin.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. ���翡 ���� �˻��׸� ����
      if qryGumgin.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. ��� ���� �˻��׸� ����
      if qryGumgin.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGumgin.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. �߰��ڵ忡 ���� �˻��׸� ����
   sTemp := sTemp + qryGumgin.FieldByName('Cod_chuga').AsString;;

   // 4. �׸��ڵ忡 ���� �˻��׸� ����
   sHmCode := '';
   if qryGumgin.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '1', qryGumgin.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGumgin.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '4', qryGumgin.FieldByName('Gubn_adyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '7', qryGumgin.FieldByName('Gubn_lfyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '3',qryGumgin.FieldByName('Gubn_bgyh').AsInteger);

   if qryGumgin.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGumgin.FieldByName('Dat_gmgn').AsString, '5',qryGumgin.FieldByName('Gubn_agyh').AsInteger);

   if (qryGumgin.FieldByName('Gubn_tkgm').AsString = '1') or (qryGumgin.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGumgin.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;
   result := sTemp;
end;

function TfrmLD4Q56.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q56.SBut_ExcelClick(Sender: TObject);
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