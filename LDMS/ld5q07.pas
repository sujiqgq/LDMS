//==============================================================================
// ���α׷� ����    : ���� list ��Ȳ
// �ý���           : �м��������α׷�
// ��������         : 2006.05.17
// [������]�������� : �߰�..
// �������         :
//==============================================================================
// ��������         : 2014.04.21
// [������]�������� : [������], [����]+[������+����]RIA - ���� 1�϶��� ��ȸ����
// �������         : (��û�� : �λ�-������)
//==============================================================================
// ��������         : 2014.04.29
// [������]�������� : [������], 1. 4-[������]+[������+����]RIA ����
//                              2. 3-[������]+[������+����](��Ź)RIA�� ����Ʈ ��ȸ�� S016 �˻�� ��ȸ
// �������         : (��û�� : �λ�-������)
//                    HCV Ab(S016) �˻�� EIA������ �ϳ� ��ü�� ì��� ���� RIA�˻���Ʈ���� ���ڵ尡 RIA�� �����⵵ ��.
//==============================================================================
// ��������  : 2014.05.21
// �� �� ��  : ������
// ��������  : S019, C044, C049, E001, E002, E003, E015, E040, T002, T006,
//             T007, T009, T011, T037 ��� EIA�̳�, RIA�� ��ȸ �����ϵ��� (TT01, TT02, S016�� ��������..)
// �������  : [��û�� : �λ�-������]
//==============================================================================
// ��������  : 2015.04.20
// �� �� ��  : ������
// ��������  : ���,���κ� ��� & C032�˻� ������ ������ü ���� �˻�
//==============================================================================
// ��������  : 2016.02.17
// �� �� ��  : �ڼ���
// ��������  : ���� �������հ˻�3��Ʈ 1600068
//==============================================================================
unit LD5Q07;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit, ORSingle;

type
  TfrmLD5Q07 = class(TfrmSingle)
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
    Label3: TLabel;
    edtDc: TEdit;
    btnDc: TSpeedButton;
    edtD_dc: TEdit;
    cmbCOD_JISA: TComboBox;
    Label1: TLabel;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    Label4: TLabel;
    cmbHulgum: TComboBox;
    Label5: TLabel;
    Com_BD: TComboBox;
    RBtn_Preview: TRadioButton;
    RBtn_Print: TRadioButton;
    qryJHangmokList: TQuery;
    qryHmCheck: TQuery;
    qryHmCheck2: TQuery;
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
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
  private
    { Private declarations }

    // Grid�� ������ Variant ���� ����(Report���� �����ϹǷ� Public�� ����)
    UV_vGubn_hulgum, UV_vCod_bjjs, UV_vNum_bunju,  UV_vDesc_name, UV_vNum_jumin,
    UV_vNum_samp,    UV_vGubn_sex, UV_vCod_jangbi, UV_vCod_hul,   UV_vCod_Cancer,
    UV_vCod_chuga,   UV_vCod_etc,  UV_vDesc_dc : Variant;

    UV_sCod_jisa : String;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);

    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_No_Hangmok_Hm(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;

  public
    { Public declarations }
  end;

var
  frmLD5Q07: TfrmLD5Q07;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD5Q071;

{$R *.DFM}


procedure TfrmLD5Q07.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];
   //--------------------------------------------------------------------------
   // ��������
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vGubn_hulgum[Lo]; UV_vGubn_hulgum[Lo] := UV_vGubn_hulgum[Hi]; UV_vGubn_hulgum[Hi] := T;
            T := UV_vCod_bjjs[Lo];    UV_vCod_bjjs[Lo]    := UV_vCod_bjjs[Hi];    UV_vCod_bjjs[Hi]    := T;
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];   UV_vNum_bunju[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];   UV_vDesc_name[Hi]   := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];   UV_vNum_jumin[Hi]   := T;
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];     UV_vDesc_dc[Hi]     := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]    := UV_vNum_samp[Hi];    UV_vNum_samp[Hi]    := T;
            T := UV_vGubn_sex[Lo];    UV_vGubn_sex[Lo]    := UV_vGubn_sex[Hi];    UV_vGubn_sex[Hi]    := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo]  := UV_vCod_jangbi[Hi];  UV_vCod_jangbi[Hi]  := T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]     := UV_vCod_hul[Hi];     UV_vCod_hul[Hi]     := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo]  := UV_vCod_Cancer[Hi];  UV_vCod_Cancer[Hi]  := T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]   := UV_vCod_chuga[Hi];   UV_vCod_chuga[Hi]   := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]     := UV_vCod_etc[Hi];     UV_vCod_etc[Hi]     := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // ��������
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vGubn_hulgum[Lo]; UV_vGubn_hulgum[Lo] := UV_vGubn_hulgum[Hi]; UV_vGubn_hulgum[Hi] := T;
            T := UV_vCod_bjjs[Lo];    UV_vCod_bjjs[Lo]    := UV_vCod_bjjs[Hi];    UV_vCod_bjjs[Hi]    := T;
            T := UV_vNum_bunju[Lo];   UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];   UV_vNum_bunju[Hi]   := T;
            T := UV_vDesc_name[Lo];   UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];   UV_vDesc_name[Hi]   := T;
            T := UV_vNum_jumin[Lo];   UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];   UV_vNum_jumin[Hi]   := T;
            T := UV_vDesc_dc[Lo];     UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];     UV_vDesc_dc[Hi]     := T;
            T := UV_vNum_samp[Lo];    UV_vNum_samp[Lo]    := UV_vNum_samp[Hi];    UV_vNum_samp[Hi]    := T;
            T := UV_vGubn_sex[Lo];    UV_vGubn_sex[Lo]    := UV_vGubn_sex[Hi];    UV_vGubn_sex[Hi]    := T;
            T := UV_vCod_jangbi[Lo];  UV_vCod_jangbi[Lo]  := UV_vCod_jangbi[Hi];  UV_vCod_jangbi[Hi]  := T;
            T := UV_vCod_hul[Lo];     UV_vCod_hul[Lo]     := UV_vCod_hul[Hi];     UV_vCod_hul[Hi]     := T;
            T := UV_vCod_Cancer[Lo];  UV_vCod_Cancer[Lo]  := UV_vCod_Cancer[Hi];  UV_vCod_Cancer[Hi]  := T;
            T := UV_vCod_chuga[Lo];   UV_vCod_chuga[Lo]   := UV_vCod_chuga[Hi];   UV_vCod_chuga[Hi]   := T;
            T := UV_vCod_etc[Lo];     UV_vCod_etc[Lo]     := UV_vCod_etc[Hi];     UV_vCod_etc[Hi]     := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD5Q07.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vGubn_hulgum := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_sex    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_etc     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vGubn_hulgum, iLength);
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);
      VarArrayRedim(UV_vGubn_sex,    iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_Cancer,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vCod_etc,     iLength);
      VarArrayRedim(UV_vDesc_dc,     iLength);
   end;
end;

function TfrmLD5Q07.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', '�������ڸ� �Է��ؾ� �մϴ�.');
      Result := False;
   end;
end;

procedure TfrmLD5Q07.FormCreate(Sender: TObject);
begin
   inherited;

   // Grid�� ȯ�漳�� & Variant���� Memory�Ҵ�
   UP_VarMemSet(0);

   Com_BD.ItemIndex := GP_SawonCheck(Com_BD, GV_sUserId);

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


   cmbHulgum.ItemIndex := 0;
   Com_BD.ItemIndex  := 0;
end;

procedure TfrmLD5Q07.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vGubn_hulgum[DataRow-1];
      2 : Value := UV_vCod_bjjs[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      6 : Value := UV_vDesc_dc[DataRow-1];
      7 : Value := UV_vNum_samp[DataRow-1];
      8 : Value := UV_vGubn_sex[DataRow-1];
      9 : Value := UV_vCod_jangbi[DataRow-1];
     10 : Value := UV_vCod_hul[DataRow-1];
     11 : Value := UV_vCod_Cancer[DataRow-1];
     12 : Value := UV_vCod_chuga[DataRow-1];
     13 : Value := UV_vCod_etc[DataRow-1];
   end;
end;

procedure TfrmLD5Q07.btnQueryClick(Sender: TObject);
var index, iRet, temp, iAge : Integer;
    sSelect, sWhere, sOrder, yh_name, chuga, sMan,sEtc_Hangmok_hm,sTk_Hangmok_Pf,sTk_Hangmok_hm ,sSelect1,sWhere1,sWhere2,sOderby1,sTotal_HangmokList: String;
    vCod_chuga : Variant;
     i : Integer;
    vCheck_01, vCheck_08, vCheck_01_01, vCheck_08_01, vCheck_01_02,vCheck_04,vCheck_01_03, vCheck_01_jaegum, chk_01, chk_02 : Boolean;
begin
   inherited;
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid & Chart �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   index := 0;
   sSelect := ''; sWhere := ''; sOrder := '';
   with qryBunju do
   begin
      // SQL���� ������ �ڷḦ Query
      Active := False;
      if (cmbHulgum.ItemIndex = 0) or (cmbHulgum.ItemIndex = 2) or (cmbHulgum.ItemIndex = 5) then
      begin
         sSelect := ' SELECT G.dat_gmgn,  G.desc_name, G.num_jumin, G.num_id, G.cod_jangbi, G.num_samp, G.cod_jisa, ' +
                    ' G.cod_hul, G.cod_cancer, G.cod_chuga, G.gubn_nosin, G.gubn_nsyh, G.gubn_hulgum, ' +
                    ' G.gubn_bogun, G.gubn_bgyh, G.gubn_adult, G.gubn_adyh, G.gubn_agab, G.gubn_agyh, ' +
                    ' G.gubn_gong, G.gubn_goyh, G.gubn_life, G.gubn_lfyh, G.gubn_tkgm, B.dat_bunju, B.num_bunju, B.cod_jisa, B.cod_bjjs, D.desc_dc '+
                    ' FROM gumgin G LEFT OUTER JOIN BUNJU B ' +
                    ' ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ' +
                    ' LEFT OUTER JOIN danche D ON G.cod_dc = D.cod_dc';

         sWhere := ' WHERE G.cod_jisa = ''' + UV_sCod_jisa + '''';
         sWhere := sWhere + ' AND G.dat_gmgn = ''' + mskDate.Text + '''';
         case cmbHulgum.ItemIndex of
            0 : sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';
            2 : sWhere := sWhere + ' AND G.gubn_hulgum in (''1'',''3'')';
            5 : sWhere := sWhere + ' AND G.gubn_hulgum in (''1'')';
         end;

         if Trim(Copy(Com_BD.Text, 1, 2)) <> '0' then
            sWhere := sWhere + ' AND G.gubn_jinch = ''' + Trim(Copy(Com_BD.Text, 1, 2)) + '''';
         if Trim(mskFrom.Text) <> '' then
         begin
            sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
            sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
         end;

         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + ' AND B.cod_dc = ''' + edtDc.Text + '''';
      end
      else if (cmbHulgum.ItemIndex = 1) or (cmbHulgum.ItemIndex = 3) or (cmbHulgum.ItemIndex = 4) then
      begin
         sSelect := 'select D.Desc_dc,g.*,B.num_bunju,B.cod_bjjs,B.dat_bunju,t.cod_prf,t.cod_chuga as Tcod_chuga  ' +
                    ' from  gumgin g  left outer join bunju b on g.Cod_jisa=b.Cod_jisa and g.num_id=b.num_id and  g.dat_gmgn=b.dat_gmgn  ' +
                    ' left outer join tkgum T on g.Cod_jisa=t.Cod_jisa and g.num_jumin=t.num_jumin and  g.dat_gmgn=t.dat_gmgn and g.gubn_tkgm =t.gubn_cha '+
                    ' LEFT OUTER JOIN danche D ON G.cod_dc = D.cod_dc';
         sWhere := ' WHERE G.cod_jisa = ''' + UV_sCod_jisa + '''';
         sWhere := sWhere + ' AND G.dat_gmgn = ''' + mskDate.Text + '''';
         case cmbHulgum.ItemIndex of   //2014.04.29~ ������
              1 : sWhere := sWhere + ' AND G.gubn_hulgum in (''2'',''3'')';
              4 : sWhere := sWhere + ' AND G.gubn_hulgum in (''1'',''3'')';
         end;                         //~2014.04.29
         if Trim(Copy(Com_BD.Text, 1, 2)) <> '0' then
            sWhere := sWhere + ' AND G.gubn_jinch = ''' + Trim(Copy(Com_BD.Text, 1, 2)) + '''';
         if Trim(mskFrom.Text) <> '' then
         begin
            sWhere := sWhere + ' AND B.num_bunju >= ''' + GF_LPad(Trim(mskFrom.Text), 4, '0') + '''';
            sWhere := sWhere + ' AND B.num_bunju <= ''' + GF_LPad(Trim(mskTo.Text), 4, '0') + '''';
         end;
         if Trim(edtDc.Text) <> '' then
            sWhere := sWhere + ' AND G.cod_dc = ''' + edtDc.Text + '''';
      end;
      sOrder := ' order by G.cod_jisa, G.dat_gmgn, CAST(G.num_samp AS INT) ';
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);
      Active := True;

      //if  (cmbHulgum.ItemIndex = 0) or (cmbHulgum.ItemIndex = 2) or (cmbHulgum.ItemIndex = 5) then
      if  (cmbHulgum.ItemIndex = 0) or (cmbHulgum.ItemIndex = 5) then
      begin
         if RecordCount > 0 then
         begin
           UP_VarMemSet(RecordCount-1);
           // DataSet�� ���� Variant������ �̵�
           while not Eof do
           begin
             yh_name := '';   chuga := '';
             GP_JuminToAge(FieldByName('Num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);

             if FieldByName('Gubn_hulgum').AsString = '1' then
               UV_vGubn_hulgum[index] := '������'
             else if FieldByName('Gubn_hulgum').AsString = '2' then
               UV_vGubn_hulgum[index] := '����'
             else if FieldByName('Gubn_hulgum').AsString = '3' then
             begin
               if (FieldByName('gubn_nosin').AsString = '1') or
                  (FieldByName('gubn_nosin').AsString = '2') or
                  (FieldByName('gubn_adult').AsString = '1') or
                  (FieldByName('gubn_adult').AsString = '2') or
                  (FieldByName('gubn_life').AsString  = '1') or
                  (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '������+����(��)'
               else                                               UV_vGubn_hulgum[index] := '������+����';
             end;
             UV_vCod_bjjs[index]    := FieldByName('Cod_bjjs').AsString;
             UV_vNum_bunju[index]   := FieldByName('Num_bunju').AsString;
             UV_vDesc_name[index]   := FieldByName('Desc_name').AsString;
             UV_vNum_jumin[index]   := FieldByName('Num_jumin').AsString;
             UV_vDesc_dc[index]     := FieldByName('Desc_dc').AsString;
             UV_vNum_samp[index]    := FieldByName('Num_samp').AsString;
             UV_vGubn_sex[index]    := sMan;
             UV_vCod_jangbi[index]  := FieldByName('Cod_jangbi').AsString;
             UV_vCod_hul[index]     := FieldByName('Cod_hul').AsString;
             UV_vCod_Cancer[index]  := FieldByName('Cod_Cancer').AsString;

             if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
             begin
               chuga := FieldByName('Cod_chuga').AsString + '(Ư��: ';
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
                             chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
               end;           //if(qryTkgum) end;
               UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
             end
             else UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;

             // �������Display
             if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
             else if FieldByName('gubn_nosin').AsString = '2' then
               yh_name := yh_name + '������';

             //Ư������Display
             if FieldByName('gubn_tkgm').AsString = '1' then
               yh_name := yh_name + 'Ư��'
             else if FieldByName('gubn_tkgm').AsString = '2' then
               yh_name := yh_name + 'Ư�����';

             //���κ�����Display
             if FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
             else if FieldByName('gubn_adult').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger) + ', ';

             //��������Display
             if FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
             else if FieldByName('gubn_agab').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger) + ', ';

             //�����Ǻ�����Display
             if FieldByName('gubn_gong').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
             else if FieldByName('gubn_gong').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger) + ', ';

             //������ȯ������Display
             if FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger)
             else if FieldByName('gubn_life').AsString = '2' then
               yh_name := yh_name + '�������';

             UV_vCod_etc[index] := yh_name;

             Next;
             Inc(index);
           end;
           // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
           Active := False;
         end;
      end
      //20180302 ����ä�� ��ȭ�� ��ü9�� -> ��ü�˻� ��ȯ
      else if  (cmbHulgum.ItemIndex = 2)  then
      begin
      if qryBunju.RecordCount > 0 then
      begin
      UP_VarMemSet(qryBunju.RecordCount-1);

      while not qryBunju.Eof do
      begin
           // DataSet�� ���� Variant������ �̵�
             with qryHmCheck do
             begin
                qryHmCheck.Active := False;
                Chk_01 := false;
                qryHmCheck.Sql.Clear;
                qryHmCheck.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryBunju.FieldByName('cod_jisa').AsString + ''','''
                                                                           + qryBunju.FieldByName('num_id').AsString   + ''','''
                                                                           + qryBunju.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                           + 'JOIN hangmok H ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
                qryHmCheck.Sql.Text := Sql.Text + ' WHERE (gubn_part =''04'') or (gubn_part =''05'') or (gubn_part =''07'')';
                qryHmCheck.Active := True;
             end;
             if (qryHmCheck.RecordCount > 0) then Chk_01 := true;

             with qryHmCheck2 do
             begin
                qryHmCheck2.Active := False;
                Chk_02 := false;
                qryHmCheck2.Sql.Clear;
                qryHmCheck2.Sql.Text := ' SELECT DISTINCT * FROM TF_Get_HangmokList(''' + qryBunju.FieldByName('cod_jisa').AsString + ''','''
                                                                           + qryBunju.FieldByName('num_id').AsString   + ''','''
                                                                           + qryBunju.FieldByName('dat_gmgn').AsString + ''',''1'') '
                                                                           + 'JOIN hangmok H ON H.cod_hm =  TF_Get_HangmokList.Cod_hm ';
                qryHmCheck2.Sql.Text := Sql.Text + ' WHERE (gubn_part =''01'') and (TF_Get_HangmokList.Cod_hm not IN (''C009'', ''C010'', ''C011'', ''C025'', ''C026'', ''C027'', ''C028'', ''C032'', ''C042'', ''C074'', ''C033''))';
                qryHmCheck2.Active := True;
             end;
             if (qryHmCheck2.RecordCount > 0) then Chk_02 := true;

             if ((chk_01 = true) and (Chk_02 = true)) or (chk_01 = true) or (chk_02 = true)  then
             begin
             yh_name := '';   chuga := '';
             GP_JuminToAge(qryBunju.FieldByName('Num_jumin').AsString,qryBunju.FieldByName('dat_gmgn').AsString, sMan, iAge);

             if qryBunju.FieldByName('Gubn_hulgum').AsString = '1' then
               UV_vGubn_hulgum[index] := '������'
             else if qryBunju.FieldByName('Gubn_hulgum').AsString = '2' then
               UV_vGubn_hulgum[index] := '����'
             else if qryBunju.FieldByName('Gubn_hulgum').AsString = '3' then
             begin
               if (qryBunju.FieldByName('gubn_nosin').AsString = '1') or
                  (qryBunju.FieldByName('gubn_nosin').AsString = '2') or
                  (qryBunju.FieldByName('gubn_adult').AsString = '1') or
                  (qryBunju.FieldByName('gubn_adult').AsString = '2') or
                  (qryBunju.FieldByName('gubn_life').AsString  = '1') or
                  (qryBunju.FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '������+����(��)'
               else                                               UV_vGubn_hulgum[index] := '������+����';
             end;
             UV_vCod_bjjs[index]    := qryBunju.FieldByName('Cod_bjjs').AsString;
             UV_vNum_bunju[index]   := qryBunju.FieldByName('Num_bunju').AsString;
             UV_vDesc_name[index]   := qryBunju.FieldByName('Desc_name').AsString;
             UV_vNum_jumin[index]   := qryBunju.FieldByName('Num_jumin').AsString;
             UV_vDesc_dc[index]     := qryBunju.FieldByName('Desc_dc').AsString;
             UV_vNum_samp[index]    := qryBunju.FieldByName('Num_samp').AsString;
             UV_vGubn_sex[index]    := sMan;
             UV_vCod_jangbi[index]  := qryBunju.FieldByName('Cod_jangbi').AsString;
             UV_vCod_hul[index]     := qryBunju.FieldByName('Cod_hul').AsString;
             UV_vCod_Cancer[index]  := qryBunju.FieldByName('Cod_Cancer').AsString;

             if (qryBunju.FieldByName('gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('gubn_tkgm').AsString = '2') then
             begin
               chuga := FieldByName('Cod_chuga').AsString + '(Ư��: ';
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString := qryBunju.FieldByName('cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString := qryBunju.FieldByName('dat_gmgn').AsString;
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
                             chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                          qryPf_hm.next;
                       end;   // while(qryPf_hm) end;
                    end;      //if(qryPf_hm) end;
                  end;        //for(qryTkgum) end;
               end;           //if(qryTkgum) end;
               UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
             end
             else UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;

             // �������Display
             if FieldByName('gubn_nosin').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', FieldByName('gubn_nsyh').AsInteger) + ', '
             else if FieldByName('gubn_nosin').AsString = '2' then
               yh_name := yh_name + '������';

             //Ư������Display
             if qryBunju.FieldByName('gubn_tkgm').AsString = '1' then
               yh_name := yh_name + 'Ư��'
             else if qryBunju.FieldByName('gubn_tkgm').AsString = '2' then
               yh_name := yh_name + 'Ư�����';

             //���κ�����Display
             if qryBunju.FieldByName('gubn_adult').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', FieldByName('gubn_adyh').AsInteger) + ', '
             else if qryBunju.FieldByName('gubn_adult').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_adyh').AsInteger) + ', ';

             //��������Display
             if qryBunju.FieldByName('gubn_agab').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', FieldByName('gubn_agyh').AsInteger)
             else if qryBunju.FieldByName('gubn_agab').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger) + ', ';

             //�����Ǻ�����Display
             if qryBunju.FieldByName('gubn_gong').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', FieldByName('gubn_goyh').AsInteger)
             else if qryBunju.FieldByName('gubn_gong').AsString = '2' then
               yh_name := yh_name + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_goyh').AsInteger) + ', ';

             //������ȯ������Display
             if qryBunju.FieldByName('gubn_life').AsString = '1' then
               yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', FieldByName('gubn_lfyh').AsInteger)
             else if qryBunju.FieldByName('gubn_life').AsString = '2' then
               yh_name := yh_name + '�������';

             UV_vCod_etc[index] := yh_name;

             Inc(index);
           end;
             Next;
           end;

           // Table�� Disconnect -> ������ �۾����� Variant������ �̿��ؼ� ����
           Active := False;
         end;
      end

      else if (cmbHulgum.ItemIndex = 1) or (cmbHulgum.ItemIndex = 4) then
      begin
        if qrybunju.RecordCount >0 then
        begin
           while  not qrybunju.Eof do
           begin
                sEtc_Hangmok_hm := '';
                sTk_Hangmok_Pf  := '';
                sTk_Hangmok_hm  := '';
                sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOderby1 := '';
                vCheck_01_03:=false; vCheck_01_jaegum := False;
                //------------------------------------------------------------------------
                //�˻��׸� ALL Display...
                //------------------------------------------------------------------------
                //1.Ư�������׸�üũ...
                if (qrybunju.FieldByName('gubn_tkgm').AsString <> '') and (qrybunju.FieldByName('gubn_tkgm').AsString <> '0')then
                begin
                     sTk_Hangmok_Pf := qrybunju.FieldByName('cod_prf').AsString;
                     sTk_Hangmok_hm := qrybunju.FieldByName('Tcod_chuga').AsString;
                end;
                if FieldByName('gubn_nosin').AsString = '1' then
                   sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok_Hm(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
                //���κ�����Display
                if FieldByName('gubn_adult').AsString = '1' then
                   sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok_Hm(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
                //��������Display
                if FieldByName('gubn_agab').AsString = '1' then
                   sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok_Hm(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
                //������ȯ������Display
                if FieldByName('gubn_life').AsString = '1' then
                   sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok_Hm(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

                //�߰��˻��׸�
                sEtc_Hangmok_hm := sEtc_Hangmok_hm + qrybunju.FieldByName('cod_chuga').AsString;

                //�������� ���Ǹ����...
                if Trim(sTk_Hangmok_Pf) <> '' then
                begin
                   sWhere1 := qrybunju.FieldByName('cod_jangbi').AsString + ''',''' + qrybunju.FieldByName('cod_hul').AsString + ''',''' + qrybunju.FieldByName('cod_cancer').AsString + ''',''';
                   For i := 1 to (length(sTk_Hangmok_Pf) div 4) do
                   begin
                      if i = (length(sTk_Hangmok_Pf) div 4) then sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4)
                      else                                       sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4) + ''',''';
                   end;
                end
                else
                   sWhere1 := qrybunju.FieldByName('cod_jangbi').AsString + ''',''' + qrybunju.FieldByName('cod_hul').AsString + ''',''' + qrybunju.FieldByName('cod_cancer').AsString + ''',''';

                if Trim(sTk_Hangmok_hm) <> '' then
                begin
                   sWhere2 := sWhere2 + ''',''';
                   For i := 1 to (length(sTk_Hangmok_hm) div 4) do
                   begin
                      if (i = (length(sTk_Hangmok_hm) div 4)) and (Trim(sEtc_Hangmok_hm) = '') then sWhere2 := sWhere2 + copy(sTk_Hangmok_hm, (i*4)-3, 4)
                      else                                                                          sWhere2 := sWhere2 + copy(sTk_Hangmok_hm, (i*4)-3, 4) + ''',''';
                   end;
                end;

                if Trim(sEtc_Hangmok_hm) <> '' then
                begin
                   sWhere2 := sWhere2 + ''',''';
                   For i := 1 to (length(sEtc_Hangmok_hm) div 4) do
                   begin
                      if i = (length(sEtc_Hangmok_hm) div 4) then sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4)
                      else                                           sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
                   end;
                end;

                sTotal_HangmokList := '';
                with qryJHangmokList do
                begin
                   sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + qrybunju.FieldByName('dat_gmgn').AsString+ '''';
                   sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
                   sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part ) UNION ( ';
                   sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
                   sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
                   sSelect := sSelect + '   AND dat_apply <= ''' +  qrybunju.FieldByName('dat_gmgn').AsString + ''' )';
                   sOderby1 := ' ORDER BY H.gubn_part Desc, P.cod_hm ';
                   qryJHangmokList.Active := False;
                   qryJHangmokList.SQL.Clear;
                   qryJHangmokList.SQL.Add(sSelect + sOderby1);
                   qryJHangmokList.Active := True;

                   if qryJHangmokList.RecordCount > 0 then
                   begin
                      while not Eof do
                      begin
                         sTotal_HangmokList := sTotal_HangmokList + FieldByName('cod_hm').AsString;
                         //��ȭ��/���ܸ鿪 ����...
                         if (FieldByName('gubn_part').AsString  = '01')   then
                         begin
                              if ((FieldByName('cod_hm').AsString = 'C001') or
                                  (FieldByName('cod_hm').AsString = 'C002') or
                                  (FieldByName('cod_hm').AsString = 'C003') or
                                  (FieldByName('cod_hm').AsString = 'C004') or
                                  (FieldByName('cod_hm').AsString = 'C005') or
                                  (FieldByName('cod_hm').AsString = 'C006') or
                                  (FieldByName('cod_hm').AsString = 'C007') or
                                  (FieldByName('cod_hm').AsString = 'C009') or
                                  (FieldByName('cod_hm').AsString = 'C010') or
                                  (FieldByName('cod_hm').AsString = 'C011') or
                                  (FieldByName('cod_hm').AsString = 'C012') or
                                  (FieldByName('cod_hm').AsString = 'C013') or
                                  (FieldByName('cod_hm').AsString = 'C017') or
                                  (FieldByName('cod_hm').AsString = 'C019') or
                                  (FieldByName('cod_hm').AsString = 'C021') or
                                  (FieldByName('cod_hm').AsString = 'C025') or
                                  (FieldByName('cod_hm').AsString = 'C026') or
                                  (FieldByName('cod_hm').AsString = 'C027') or
                                  (FieldByName('cod_hm').AsString = 'C028') or
                                  (FieldByName('cod_hm').AsString = 'C029') or
                                  (FieldByName('cod_hm').AsString = 'C032') or
                                  (FieldByName('cod_hm').AsString = 'C033') or
                                  (FieldByName('cod_hm').AsString = 'C041') or
                                  (FieldByName('cod_hm').AsString = 'C042') or
                                  (FieldByName('cod_hm').AsString = 'C043')) and
                                 (mskDate.Text >='20130729') and (mskDate.Text <'20140308')then vCheck_01_03 := True
                              else if (FieldByName('cod_hm').AsString = 'C032') and //20150420 ���,���κ� ��� & C032
                                      (mskDate.Text >='20150410')then vCheck_01_jaegum := True;
                         end;
                         Next;
                      end;
                   end;
                end;
                if ((FieldByName('gubn_nosin').AsString = '1') or
                    (FieldByName('gubn_adult').AsString = '1') or
                    (FieldByName('gubn_life').AsString  = '1')) and
                   (mskDate.Text >='20140308') then
                    vCheck_01_03 := True;

                if // (qrybunju.FieldByName('gubn_hulgum').AsString='2')   or   2014.04.29 ������
                   (qrybunju.FieldByName('gubn_hulgum').AsString='3') and
                   (qrybunju.FieldByName('cod_jisa').AsString <> '15') and
                   ((vCheck_01_03=true) or (vCheck_01_jaegum=True)) and
                   (Copy(qrybunju.FieldByName('Cod_jangbi').AsString,1,2) <> 'SS') and
                   (Copy(qrybunju.FieldByName('Cod_Hul').AsString,1,2) <> 'SS') and
                   (Copy(qrybunju.FieldByName('Cod_jangbi').AsString,1,2) <> 'GS') and
                   (Copy(qrybunju.FieldByName('Cod_Hul').AsString,1,2) <> 'GS') then
                begin
                   yh_name := '';   chuga := '';
                   UP_VarMemSet(RecordCount-1);
                   GP_JuminToAge(qrybunju.FieldByName('Num_jumin').AsString,qrybunju.FieldByName('dat_gmgn').AsString, sMan, iAge);


                   if FieldByName('Gubn_hulgum').AsString = '1' then
                      UV_vGubn_hulgum[index] := '������'
                   else if FieldByName('Gubn_hulgum').AsString = '2' then
                      UV_vGubn_hulgum[index] := '����'
                   else if FieldByName('Gubn_hulgum').AsString = '3' then
                   begin
                        if (FieldByName('gubn_nosin').AsString = '1') or
                           (FieldByName('gubn_nosin').AsString = '2') or
                           (FieldByName('gubn_adult').AsString = '1') or
                           (FieldByName('gubn_adult').AsString = '2') or
                           (FieldByName('gubn_life').AsString  = '1') or
                           (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '������+����(��)'
                       else UV_vGubn_hulgum[index] := '������+����';
                   end;
                   UV_vCod_bjjs[index]    := qrybunju.FieldByName('Cod_bjjs').AsString;
                   UV_vNum_bunju[index]   := qrybunju.FieldByName('Num_bunju').AsString;
                   UV_vDesc_name[index]   := qrybunju.FieldByName('Desc_name').AsString;
                   UV_vNum_jumin[index]   := qrybunju.FieldByName('Num_jumin').AsString;
                   UV_vDesc_dc[index]     := qrybunju.FieldByName('Desc_dc').AsString;
                   UV_vNum_samp[index]    := qrybunju.FieldByName('Num_samp').AsString;
                   UV_vGubn_sex[index]    := sMan;
                   UV_vCod_jangbi[index]  := qrybunju.FieldByName('Cod_jangbi').AsString;
                   UV_vCod_hul[index]     := qrybunju.FieldByName('Cod_hul').AsString;
                   UV_vCod_Cancer[index]  := qrybunju.FieldByName('Cod_Cancer').AsString;

                   if (qrybunju.FieldByName('gubn_tkgm').AsString = '1') or (qrybunju.FieldByName('gubn_tkgm').AsString = '2') then
                   begin
                      chuga := qrybunju.FieldByName('Cod_chuga').AsString + '(Ư��: ';
                      qryTkgum.Active := False;
                      qryTkgum.ParamByName('cod_jisa').AsString := qrybunju.FieldByName('cod_jisa').AsString;
                      qryTkgum.ParamByName('num_jumin').AsString := qrybunju.FieldByName('Num_jumin').AsString;
                      qryTkgum.ParamByName('dat_gmgn').AsString := qrybunju.FieldByName('dat_gmgn').AsString;
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
                                     chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                                  qryPf_hm.next;
                               end;   // while(qryPf_hm) end;
                           end;      //if(qryPf_hm) end;
                         end;        //for(qryTkgum) end;
                      end;           //if(qryTkgum) end;
                      UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
                   end
                   else UV_vCod_chuga[index]  := qrybunju.FieldByName('Cod_chuga').AsString;

                    // �������Display
                   if qrybunju.FieldByName('gubn_nosin').AsString = '1' then
                      yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', qrybunju.FieldByName('gubn_nsyh').AsInteger) + ', '
                   else if qrybunju.FieldByName('gubn_nosin').AsString = '2' then
                      yh_name := yh_name + '������';

                    //Ư������Display
                   if qrybunju.FieldByName('gubn_tkgm').AsString = '1' then
                      yh_name := yh_name + 'Ư��'
                   else if qrybunju.FieldByName('gubn_tkgm').AsString = '2' then
                      yh_name := yh_name + 'Ư�����';

                    //���κ�����Display
                   if qrybunju.FieldByName('gubn_adult').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', qrybunju.FieldByName('gubn_adyh').AsInteger) + ', '
                   else if qrybunju.FieldByName('gubn_adult').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '4', qrybunju.FieldByName('gubn_adyh').AsInteger) + ', ';

                    //��������Display
                   if qrybunju.FieldByName('gubn_agab').AsString = '1' then
                       yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', qrybunju.FieldByName('gubn_agyh').AsInteger)
                   else if qrybunju.FieldByName('gubn_agab').AsString = '2' then
                       yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '5', qrybunju.FieldByName('gubn_agyh').AsInteger) + ', ';

                    //�����Ǻ�����Display
                   if qrybunju.FieldByName('gubn_gong').AsString = '1' then
                      yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', qrybunju.FieldByName('gubn_goyh').AsInteger)
                   else if qrybunju.FieldByName('gubn_gong').AsString = '2' then
                      yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '6', qrybunju.FieldByName('gubn_goyh').AsInteger) + ', ';

                    //������ȯ������Display
                   if qrybunju.FieldByName('gubn_life').AsString = '1' then
                      yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', qrybunju.FieldByName('gubn_lfyh').AsInteger)
                   else if qrybunju.FieldByName('gubn_life').AsString = '2' then
                     yh_name := yh_name + '�������';
                   UV_vCod_etc[index] := yh_name;
                   Inc(index);
                end;
                qrybunju.next;
            end;
         end;
      end


      else if (cmbHulgum.ItemIndex =3)  then
      begin
         if qrybunju.RecordCount >0 then
         begin
            while  not qrybunju.Eof do
            begin
               sEtc_Hangmok_hm := '';
               sTk_Hangmok_Pf  := '';
               sTk_Hangmok_hm  := '';
               sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOderby1 := '';
               vCheck_01   := False;  vCheck_08 := False;
               vCheck_01_01 := False; vCheck_08_01 := False; vCheck_01_02 := False;
               vCheck_04 :=false;
               //------------------------------------------------------------------------
               //�˻��׸� ALL Display...
               //------------------------------------------------------------------------
               //1.Ư�������׸�üũ...
               if (qrybunju.FieldByName('gubn_tkgm').AsString <> '') and (qrybunju.FieldByName('gubn_tkgm').AsString <> '0')then
               begin
                    sTk_Hangmok_Pf := qrybunju.FieldByName('cod_prf').AsString;
                    sTk_Hangmok_hm := qrybunju.FieldByName('Tcod_chuga').AsString;
               end;


               //�߰��˻��׸�
               sEtc_Hangmok_hm := sEtc_Hangmok_hm + qrybunju.FieldByName('cod_chuga').AsString;

               //�������� ���Ǹ����...
               if Trim(sTk_Hangmok_Pf) <> '' then
               begin
                  sWhere1 := qrybunju.FieldByName('cod_jangbi').AsString + ''',''' + qrybunju.FieldByName('cod_hul').AsString + ''',''' + qrybunju.FieldByName('cod_cancer').AsString + ''',''';
                  For i := 1 to (length(sTk_Hangmok_Pf) div 4) do
                  begin
                     if i = (length(sTk_Hangmok_Pf) div 4) then sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4)
                     else                                       sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4) + ''',''';
                  end;
               end
               else
               begin
                  sWhere1 := qrybunju.FieldByName('cod_jangbi').AsString + ''',''' + qrybunju.FieldByName('cod_hul').AsString + ''',''' + qrybunju.FieldByName('cod_cancer').AsString + ''',''';
               end;

               if Trim(sTk_Hangmok_hm) <> '' then
               begin
                  sWhere2 := sWhere2 + ''',''';
                  For i := 1 to (length(sTk_Hangmok_hm) div 4) do
                  begin
                     if (i = (length(sTk_Hangmok_hm) div 4)) and (Trim(sEtc_Hangmok_hm) = '') then sWhere2 := sWhere2 + copy(sTk_Hangmok_hm, (i*4)-3, 4)
                     else                                                                          sWhere2 := sWhere2 + copy(sTk_Hangmok_hm, (i*4)-3, 4) + ''',''';
                  end;
               end;

               if Trim(sEtc_Hangmok_hm) <> '' then
               begin
                  sWhere2 := sWhere2 + ''',''';
                  For i := 1 to (length(sEtc_Hangmok_hm) div 4) do
                  begin
                     if i = (length(sEtc_Hangmok_hm) div 4) then sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4)
                     else                                           sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
                  end;
               end;

               sTotal_HangmokList := '';
               with qryJHangmokList do
               begin
                  sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + qrybunju.FieldByName('dat_gmgn').AsString+ '''';
                  sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
                  sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part ) UNION ( ';
                  sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
                  sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
                  sSelect := sSelect + '   AND dat_apply <= ''' +  qrybunju.FieldByName('dat_gmgn').AsString + ''' )';
                  sOderby1 := ' ORDER BY H.gubn_part Desc, P.cod_hm ';
                  qryJHangmokList.Active := False;
                  qryJHangmokList.SQL.Clear;
                  qryJHangmokList.SQL.Add(sSelect + sOderby1);
                  qryJHangmokList.Active := True;

                  if qryJHangmokList.RecordCount > 0 then
                  begin
                     while not Eof do
                     begin
                        sTotal_HangmokList := sTotal_HangmokList + FieldByName('cod_hm').AsString;
                         //RIA ����...

                        if (FieldByName('gubn_part').AsString  = '04') then vCheck_04 := True;
                        if (FieldByName('cod_hm').AsString = 'TT01') or (FieldByName('cod_hm').AsString = 'TT02') or
                           (FieldByName('cod_hm').AsString = 'S016') or (FieldByName('cod_hm').AsString = 'S019') or
                           (FieldByName('cod_hm').AsString = 'C004') or (FieldByName('cod_hm').AsString = 'C049') or
                           (FieldByName('cod_hm').AsString = 'E001') or (FieldByName('cod_hm').AsString = 'E002') or
                           (FieldByName('cod_hm').AsString = 'E003') or (FieldByName('cod_hm').AsString = 'E015') or
                           (FieldByName('cod_hm').AsString = 'E040') or (FieldByName('cod_hm').AsString = 'T002') or
                           (FieldByName('cod_hm').AsString = 'T006') or (FieldByName('cod_hm').AsString = 'T007') or
                           (FieldByName('cod_hm').AsString = 'T009') or (FieldByName('cod_hm').AsString = 'T011') or
                           (FieldByName('cod_hm').AsString = 'T037') then vCheck_04 := True;
                        //===============================================================
                        Next;
                     end;
                  end;
               end;
               if ((qrybunju.FieldByName('gubn_hulgum').AsString='1') or
                   (qrybunju.FieldByName('gubn_hulgum').AsString='3')) and  //������ 20140421 - ���� 1 �϶��� ��ȸ
                  (qrybunju.FieldByName('cod_jisa').AsString <> '15') then
               begin
                   if vCheck_04 then
                   begin
                       yh_name := '';   chuga := '';
                       UP_VarMemSet(RecordCount-1);
                       GP_JuminToAge(qrybunju.FieldByName('Num_jumin').AsString,qrybunju.FieldByName('dat_gmgn').AsString, sMan, iAge);

                       if FieldByName('Gubn_hulgum').AsString = '1' then     //20140421 ��.
                          UV_vGubn_hulgum[index] := '������'
                       else if FieldByName('Gubn_hulgum').AsString = '3' then
                       begin
                          if (FieldByName('gubn_nosin').AsString = '1') or
                             (FieldByName('gubn_nosin').AsString = '2') or
                             (FieldByName('gubn_adult').AsString = '1') or
                             (FieldByName('gubn_adult').AsString = '2') or
                             (FieldByName('gubn_life').AsString  = '1') or
                             (FieldByName('gubn_life').AsString  = '2') then UV_vGubn_hulgum[index] := '������+����(��)'
                          else                                                UV_vGubn_hulgum[index] := '������+����';
                       end;
                       UV_vCod_bjjs[index]    := qrybunju.FieldByName('Cod_bjjs').AsString;
                       UV_vNum_bunju[index]   := qrybunju.FieldByName('Num_bunju').AsString;
                       UV_vDesc_name[index]   := qrybunju.FieldByName('Desc_name').AsString;
                       UV_vNum_jumin[index]   := qrybunju.FieldByName('Num_jumin').AsString;
                       UV_vDesc_dc[index]     := qrybunju.FieldByName('Desc_dc').AsString;
                       UV_vNum_samp[index]    := qrybunju.FieldByName('Num_samp').AsString;
                       UV_vGubn_sex[index]    := sMan;
                       UV_vCod_jangbi[index]  := qrybunju.FieldByName('Cod_jangbi').AsString;
                       UV_vCod_hul[index]     := qrybunju.FieldByName('Cod_hul').AsString;
                       UV_vCod_Cancer[index]  := qrybunju.FieldByName('Cod_Cancer').AsString;

                       if (qrybunju.FieldByName('gubn_tkgm').AsString = '1') or (qrybunju.FieldByName('gubn_tkgm').AsString = '2') then
                       begin
                          chuga := qrybunju.FieldByName('Cod_chuga').AsString + '(Ư��: ';
                          qryTkgum.Active := False;
                          qryTkgum.ParamByName('cod_jisa').AsString := qrybunju.FieldByName('cod_jisa').AsString;
                          qryTkgum.ParamByName('num_jumin').AsString := qrybunju.FieldByName('Num_jumin').AsString;
                          qryTkgum.ParamByName('dat_gmgn').AsString := qrybunju.FieldByName('dat_gmgn').AsString;
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
                                        chuga := chuga + qryPf_hm.FieldByName('COD_HM').AsString;
                                     qryPf_hm.next;
                                  end;   // while(qryPf_hm) end;
                               end;      //if(qryPf_hm) end;
                             end;        //for(qryTkgum) end;
                          end;           //if(qryTkgum) end;
                          UV_vCod_chuga[index]  := chuga + qryTkgum.FieldByName('Cod_chuga').AsString + ')';
                       end
                       else UV_vCod_chuga[index]  := qrybunju.FieldByName('Cod_chuga').AsString;

                       // �������Display
                       if qrybunju.FieldByName('gubn_nosin').AsString = '1' then
                          yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '1', qrybunju.FieldByName('gubn_nsyh').AsInteger) + ', '
                       else if qrybunju.FieldByName('gubn_nosin').AsString = '2' then
                          yh_name := yh_name + '������';

                       //Ư������Display
                       if qrybunju.FieldByName('gubn_tkgm').AsString = '1' then
                          yh_name := yh_name + 'Ư��'
                       else if qrybunju.FieldByName('gubn_tkgm').AsString = '2' then
                          yh_name := yh_name + 'Ư�����';

                       //���κ�����Display
                       if qrybunju.FieldByName('gubn_adult').AsString = '1' then
                          yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '4', qrybunju.FieldByName('gubn_adyh').AsInteger) + ', '
                       else if qrybunju.FieldByName('gubn_adult').AsString = '2' then
                          yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '4', qrybunju.FieldByName('gubn_adyh').AsInteger) + ', ';

                       //��������Display
                       if qrybunju.FieldByName('gubn_agab').AsString = '1' then
                          yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '5', qrybunju.FieldByName('gubn_agyh').AsInteger)
                       else if qrybunju.FieldByName('gubn_agab').AsString = '2' then
                          yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '5', qrybunju.FieldByName('gubn_agyh').AsInteger) + ', ';

                       //�����Ǻ�����Display
                       if qrybunju.FieldByName('gubn_gong').AsString = '1' then
                          yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '6', qrybunju.FieldByName('gubn_goyh').AsInteger)
                       else if qrybunju.FieldByName('gubn_gong').AsString = '2' then
                          yh_name := yh_name + UF_Jae_Hangmok(qrybunju.FieldByName('cod_jisa').AsString, qrybunju.FieldByName('num_id').AsString, '6', qrybunju.FieldByName('gubn_goyh').AsInteger) + ', ';

                       //������ȯ������Display
                       if qrybunju.FieldByName('gubn_life').AsString = '1' then
                          yh_name := yh_name + UF_No_Hangmok(copy(mskDate.Text,1,4), '7', qrybunju.FieldByName('gubn_lfyh').AsInteger)
                       else if qrybunju.FieldByName('gubn_life').AsString = '2' then
                          yh_name := yh_name + '�������';

                       UV_vCod_etc[index] := yh_name;
                       Inc(index);
                   end;
               end;
               qrybunju.next;
            end;
         end;
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

procedure TfrmLD5Q07.grdMasterRowChanged(Sender: TObject; OldRow,
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

procedure TfrmLD5Q07.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtD_dc.Text := sName;
      end;
   end
   else
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD5Q07.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate then UP_Click(btnDate)
   else if Sender = edtDc   then UP_Click(btnDC);
end;

procedure TfrmLD5Q07.btnPrintClick(Sender: TObject);
begin
   inherited;

   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD5Q071 := TfrmLD5Q071.Create(Self);

   if RBtn_Print.Checked then frmLD5Q071.QuickRep.Print
   else                       frmLD5Q071.QuickRep.Preview;

end;
function  TfrmLD5Q07.UF_No_Hangmok_Hm(sDate, sGubun : String; iYh : Integer): String;
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
         if FieldByName('desc_joo').AsString <> '' then
            Result :=  Result + FieldByName('desc_joo').AsString;
      end;
      Active := False;
   end;
end;

function  TfrmLD5Q07.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD5Q07.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD5Q07.UP_Change(Sender: TObject);
var sName : String;
begin
   inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtD_dc.Text := sName;
   end;

end;

procedure TfrmLD5Q07.grdMasterHeadingClick(Sender: TObject;
  DataCol: Integer);
var iCnt : Integer;
    sOrder : String;
begin
   inherited;
   // �ڷᰡ �����ϴ��� Check
   if grdMaster.Rows = 0 then exit;

   iCnt := grdMaster.Rows;

   if grdMaster.Col[DataCol].SortPicture = spUp then
   begin
      grdMaster.Col[DataCol].SortPicture := spDown;
      sOrder := 'A';
   end
   else
   begin
      grdMaster.Col[DataCol].SortPicture := spUp;
      sOrder := 'D';
   end;

   case DataCol of
      1 : UP_QuickSort(sOrder, UV_vGubn_hulgum, 0, iCnt-1);
      2 : UP_QuickSort(sOrder, UV_vCod_bjjs,    0, iCnt-1);
      3 : UP_QuickSort(sOrder, UV_vNum_bunju,   0, iCnt-1);
      4 : UP_QuickSort(sOrder, UV_vDesc_name,   0, iCnt-1);
      5 : UP_QuickSort(sOrder, UV_vNum_jumin,   0, iCnt-1);
      6 : UP_QuickSort(sOrder, UV_vDesc_dc,     0, iCnt-1);
      7 : UP_QuickSort(sOrder, UV_vNum_samp,    0, iCnt-1);
      8 : UP_QuickSort(sOrder, UV_vGubn_sex,    0, iCnt-1);
      9 : UP_QuickSort(sOrder, UV_vCod_jangbi,  0, iCnt-1);
     10 : UP_QuickSort(sOrder, UV_vCod_hul,     0, iCnt-1);
     11 : UP_QuickSort(sOrder, UV_vCod_Cancer,  0, iCnt-1);
     12 : UP_QuickSort(sOrder, UV_vCod_chuga,   0, iCnt-1);
     13 : UP_QuickSort(sOrder, UV_vCod_etc,     0, iCnt-1);
      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;
end;

end.
