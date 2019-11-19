//==============================================================================
// ���α׷� ���� : ���ɾ�2014
// �ý���        : ���հ���
// ��������      : 2014.03.19
// ������        : ������
// ��������      : ���Ӱ� ����..
// �������      :
//==============================================================================
// ��������      : 2014.04.26
// ������        : ������
// ��������      : T040�߰� �� �ڵ庰 ��ȸ �� ���� ���
// �������      :
//==============================================================================
// ��������      : 2014.12.27
// ������        : ������
// ��������      : T009, T012 ����
// �������      : ����-������
//==============================================================================
// ��������      : 2015.01.21
// ������        : ������
// ��������      : 000917 ��ü�ڵ忡 ���ؼ� T009�ڵ� ��ȸ Ȱ��ȭ
// �������      : ����-�ڿ���
//==============================================================================
// ��������      : 2015.02.12
// ������        : ������
// ��������      :  �ش� ��ü�ڵ忡 ���� T009 �ڵ� ��ȸ �����ϵ��� ����
//                  432080,432082,432871,432083,432081,432079
// �������      : [����-������]
//==============================================================================
// ��������      : 2015.03.10
// ������        : ������
// ��������      : �ش� ��ü�ڵ忡 ���� T012 �ڵ� ��ȸ �����ϵ��� ����
//                 153899, 001880, 000890, 153910
// �������      : [����-������]
//==============================================================================
// ��������      : 2015.04.07
// ������        : ������
// ��������      : �ش� ��ü�ڵ忡 ���� T012 �ڵ� ��ȸ �����ϵ��� ����
//                 122991, 614023, 516023   + 2015.04.07 510155, 513141 �߰�
// �������      : [����-������]            + 2015.04.13 712177 �߰�
//==============================================================================
// ��������      : 2015.04.15
// ������        : ������
// ��������      : T012 - 116681 �߰�
//                 T009 - 430513,431062,431314,433141,430616,431186 �߰�
// �������      : [����-������]
//==============================================================================
// ��������      : ����  T012�߰�
// ������        : ������
// ��������      : 111034(2015.04.16)/712219(2015.04.18)/712075(2015.04.20)
//                 121565(2015.04.22)/153445(2015.04.23)/152010(2015.04.27) -> 512010(2015.05.18)����
//                 121497,122254,153604(2015.05.21)/510155->510115(2015.05.26)����/516201(2015.06.11)
//                 120096(2015.07.29)/112368,516142(2015.08.14)/159928(2015.08.17)/150127,611783(2015.08.27)
//                 712875,614454(2015.09.16)
// �������      : [����-������]
//==============================================================================
// ��������      : 20160203
// ������        : �ڼ���
// ��������      : 122801(cod_dc) ���� T009 �ڵ� ��ȸ �����ϵ��� ����
// �������      : ���� ������ܰ˻�������1600109
//==============================================================================
// ��������      : 20160302
// ������        : �ڼ���
// ��������      : 433359(cod_dc) ���� T012 �ڵ�  ��ȸ �����ϵ��� ����
// �������      : ���� ������ܰ˻�������1600232
//==============================================================================
// ��������      : 20160319
// ������        : �ڼ���
// ��������      : 510155 51210 512022 513141 516023 516142 ���� T012 �ڵ�  ��ȸ �����ϵ��� ����
// �������      : ���� ������ܰ˻�������1600325
//==============================================================================
// ��������      : 20160324
// ������        : �ڼ���
// ��������      : 510155 -> 510115�� ����(���� ��� ���� ��û), 516201�� T012,��ȸ �����ϵ��� ����
//                 712075 712465 712856�� t012 ��ȸ �����ϵ��� ����
//==============================================================================
// ��������      : 20160713
// ������        : �ڼ���
// ��������      : C088,C089 �׸� ����
//==============================================================================

unit LD2Q10;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD2Q10 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    RadioGroup1: TRadioGroup;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    Btn_NamePrint: TBitBtn;
    RBtn_1: TRadioButton;
    RBtn_2: TRadioButton;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure Btn_NamePrintClick(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure QRCompositeReportAddReports(Sender: TObject);

  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vDat_bunju, UV_vBUN, UV_vSampNo, UV_vName, UV_vJumin, UV_vCod_hm, UV_vDesc_hm : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_Bjjs : String;
  end;

var
  frmLD2Q10: TfrmLD2Q10;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q101, LD2Q103;

{$R *.DFM}

procedure TfrmLD2Q10.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q10.MouseWheelHandler(var Message: TMessage);  //����
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //����
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster.Refresh; //����
      end;
   end;
end;

procedure TfrmLD2Q10.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju:= VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo:= VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hm  := VarArrayCreate([0, 0], varOleStr);
      
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vDat_bunju,iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vSampNo,iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vCod_hm  ,iLength);
      VarArrayRedim(UV_vDesc_hm  ,iLength);
    
   end;
end;

function TfrmLD2Q10.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q10.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   RadioGroup1.ItemIndex:=2;
   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD2Q10.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDat_bunju[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vSampNo[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vJumin[DataRow-1];
      7 : Value := UV_vCod_hm   [DataRow-1];
      8 : Value := UV_vDesc_hm  [DataRow-1];

   end;
end;

procedure TfrmLD2Q10.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    temp_S019, temp_S018,
    temp_C070, //temp_C088, temp_C089, //20160713 ���������� ��û
    temp_T040, temp_T009, temp_T012 : Integer;

    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   temp_S019:=0; temp_S018:=0; temp_C070:=0;
   temp_T040:=0; temp_T009:=0; temp_T012:=0;
   Index := 0;
   UV_Bjjs := copy(cmbJisa.Text,1,2);

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
      Close;
      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, D.desc_dc, G.cod_jisa,G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, Gubn_nsyh, Gubn_adyh, ' +
                 ' Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.dat_bunju, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp          ' +
                 ' FROM bunju K with(nolock) left outer join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn                                         ' +
                 ' left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc                                                                                                              ' +
                 ' left outer join Tkgum T with(nolock) ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha                    ';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ' ;
      sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' ;

      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + ''' ' +
                            ' And K.Num_Bunju <= ''' + BunjuNo2.Text + ''' ';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + ''' ';

      sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                  ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.dat_bunju, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga, G.num_samp ';
      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY K.Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT)';
         2 : sOrderBy := ' ORDER BY CAST(G.num_samp AS INT),G.cod_jisa,K.num_Bunju,G.desc_name';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then

      begin
         GP_query_log(qryBunju, 'LD2Q10', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;  //�����׸� ��

           if (pos('S019',sHul_List) > 0) or (pos('S018',sHul_List) > 0) or (pos('C070',sHul_List) > 0) or
              (pos('T040',sHul_List) > 0) or (pos('T009',sHul_List) > 0) or (pos('T012',sHul_List) > 0) then
           begin
               if (pos('S019',sHul_List) > 0)then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'S019';
                    UV_vDesc_hm[Index]    := 'Rubella lgG';
                    temp_S019 := temp_S019 + 1;

                    Inc(Index);
               end;
               if (pos('T040',sHul_List) > 0)then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'T040';
                    UV_vDesc_hm[Index]    := 'HBc IgM';
                    temp_T040 := temp_T040 + 1;

                    Inc(Index);
               end;
               if (pos('S018',sHul_List) > 0)then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'S018';
                    UV_vDesc_hm[Index]    := 'Rebella lgM';
                    temp_S018 := temp_S018 + 1;

                    Inc(Index);
               end;
               if (pos('C070',sHul_List) > 0)then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'C070';
                    UV_vDesc_hm[Index]    := 'Alcohol';
                    temp_C070 := temp_C070 + 1;

                    Inc(Index);
               end;
               if (pos('T009',sHul_List) > 0) then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'T009';
                    UV_vDesc_hm[Index]    := 'SCC';
                    temp_T009 := temp_T009 + 1;

                    Inc(Index);
              end;
              if (pos('T012',sHul_List) > 0) then
              begin
                   UP_VarMemSet(Index+1);
                   UV_vNo[Index]  := Index+1;
                   UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                   UV_vBUN[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                   UV_vSampNo[Index]   := qryBunju.FieldByName('num_samp').AsString;
                   UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                   UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                   UV_vCod_hm[Index]     := 'T012';
                   UV_vDesc_hm[Index]    := 'NSE';
                   temp_T012 := temp_T012 + 1;

                   Inc(Index);
               end;
           END;
           Next;
         End;

         // Table�� Disconnect
         Close;

         UV_vNo[Index]        := '�հ�';
         UV_vDat_bunju[Index] :=   'HBc IgM:'+IntToStr(temp_T040);
         UV_vBUN[Index]       :=   'Rebella lgM:'+IntToStr(temp_S018);
         UV_vSampNo[Index]    :=   'Rubella lgG:'+IntToStr(temp_S019);
         UV_vName[Index]      :=   'Alcohol:'+IntToStr(temp_C070);

         if temp_T009 <> 0 then
              UV_vJumin[Index]   := 'SCC:'+IntToStr(temp_T009);
         if temp_T012 <> 0 then
         begin
              if UV_vDesc_hm[Index] <> '' then UV_vJumin[Index] := UV_vJumin[Index] +',';
              UV_vJumin[Index]   := UV_vJumin[Index] + 'NSE:'+IntToStr(temp_T012)
         end;
         Inc(index);

      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD2Q10.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD2Q10.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end
   else
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD2Q10.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD2Q101 := TfrmLD2Q101.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD2Q101.QuickRep.Preview
  else                                frmLD2Q101.QuickRep.Print;
end;

procedure TfrmLD2Q10.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

procedure TfrmLD2Q10.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;

  with QRCompositeReport do
  begin

  end;
end;

procedure TfrmLD2Q10.Btn_NamePrintClick(Sender: TObject);
begin
   inherited;
   if RBtn_2.Checked then
   begin
      frmLD2Q103 := TfrmLD2Q103.Create(Self);
      if RBtn_preveiw.Checked then frmLD2Q103.QuickRep.Preview
      else                         frmLD2Q103.QuickRep.Print;
   end;
end;

function TfrmLD2Q10.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. ���׿� ���� �˻��׸� ����
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
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
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
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
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. �׸��ڵ忡 ���� �˻��׸� ����
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
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

function TfrmLD2Q10.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
