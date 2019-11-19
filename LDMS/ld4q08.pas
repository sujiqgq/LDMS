//==============================================================================
// ���α׷� ���� : ���� ����� �� ������Ȳ
// �ý���        : ���հ���
// ��������      : 2004.05.13
// ������        : ������
// ��������      : 
// ��������      :
//==============================================================================
// ���α׷� ���� : qryPGulkwa ����� cod_bjjs �� �������� �߰�.
// ��������      : 2011.05.18
// ������        : ���ö
//==============================================================================
// ���α׷� ���� : ���� ����ִ°͸� ��ȸ�ϴ� ���ǿ��� �˻��׸�üũ�ؼ� ��ȸ�ϴ� �������� recording
// ��������      : 2011.10.19
// ������        : ����ȣ
//==============================================================================
// ���α׷� ���� : Ag,Ab�� ��� �缺�� ��� üũ��ư �߰�
// ��������      : 2011.10.20
// ������        : ����ȣ
//==============================================================================
// ���α׷� ���� : ���Ű�� ���ذ� �������� �������� �Ͽ� ��ȸ�� �� �ֵ��� ����
// ��������      : 2011.11.12
// ������        : ����ȣ
//==============================================================================
// ��������      : 2014.05.09
// ������        : ������      [��û�� : ����-�ѵη�]
// ��������      : 1. �����׸���� , Ag.Ab(��) ��� �缺�� ���  => ���ڰ�(s091/s099)�� �Բ� ��ȸ
//                 2. �����׸���� , All => ���ڰ�(s091/s099)�� �Բ� ��ȸ
//                 ==> S007/S099 , S008/S091 ��� �ϳ��� ���� ������� ��ȸ �����ϵ���
//==============================================================================
// ��������      : 2014.09.11
// ������        : ������      [��û�� : ����-�ѵη�]
// ��������      : S011, S012 ����
//==============================================================================

unit LD4Q08;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q08 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPGulkwa: TQuery;
    mskDate: TDateEdit;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    Panel2: TPanel;
    chk_all: TRadioButton;
    chk_yang: TRadioButton;
    Cmb_time: TComboBox;
    qryPPGulkwa: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    RG_print: TRadioGroup;
    RB_bunju: TRadioButton;
    RB_samp: TRadioButton;
    CB_HM: TComboBox;
    Label1: TLabel;
    qryPGulkwa_S011: TQuery;
    qryPPGulkwa_S011: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vNum_Rack, UV_vNum_samp, UV_vNum_bunju, UV_vS007, UV_vS008, UV_vpS007, UV_vpS008, UV_vDesc_name, UV_vPdat_gmgn : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
  end;

var
  frmLD4Q08: TfrmLD4Q08;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q081;

{$R *.DFM}
procedure TfrmLD4Q08.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin 
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster.Refresh; //����
      end;
   end;
end;

function  TfrmLD4Q08.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q08.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Rack  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_vS007      := VarArrayCreate([0, 0], varOleStr);
      UV_vS008      := VarArrayCreate([0, 0], varOleStr);
      UV_vpS007     := VarArrayCreate([0, 0], varOleStr);
      UV_vpS008     := VarArrayCreate([0, 0], varOleStr);
      UV_vPdat_gmgn := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vNum_Rack,        iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_vS007,      iLength);
      VarArrayRedim(UV_vS008,      iLength);
      VarArrayRedim(UV_vpS007,     iLength);
      VarArrayRedim(UV_vpS008,     iLength);
      VarArrayRedim(UV_vPdat_gmgn,     iLength);
   end;
end;

function TfrmLD4Q08.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q08.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);

   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_time.ItemIndex := 0;
   RG_print.ItemIndex := 0;
   CB_HM.ItemIndex := 0;
end;

procedure TfrmLD4Q08.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_Rack [DataRow-1];
      3 : Value := UV_vNum_samp[DataRow-1];
      4 : Value := UV_vNum_bunju[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : Value := UV_vpS007[DataRow-1];
      7 : Value := UV_vpS008[DataRow-1];
      8 : Value := UV_vS007[DataRow-1];
      9 : Value := UV_vS008[DataRow-1];
     10 : Value := UV_vPdat_gmgn[DataRow-1];
   end;
    if (DataCol = 6) and ((UV_vpS007[DataRow-1] = '�缺') or (UV_vpS007[DataRow-1] = '��缺')) then
   begin
      grdMaster.CellColor[DataCol, DataRow] := clYellow;
   end;
   if (DataCol = 7) and ((UV_vpS008[DataRow-1] = '�缺') or (UV_vpS008[DataRow-1] = '��缺')) then
   begin
      grdMaster.CellColor[DataCol, DataRow] := clYellow;
   end;

   if (DataCol = 8) and (UV_vS007[DataRow-1] <> UV_vpS007[DataRow-1]) then
   begin
      if UV_vS007[DataRow-1] = '' then
      begin
         grdMaster.CellColor[DataCol, DataRow] := clWindow;
         if UV_vpS007[DataRow-1] = '�缺' then grdMaster.CellColor[DataCol+1, DataRow] := $00FFC280;
      end
      else
         grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
   end;

   if (DataCol = 9) and (UV_vS008[DataRow-1] <> UV_vpS008[DataRow-1]) then
   begin
      if UV_vS008[DataRow-1] = '' then
      begin
         grdMaster.CellColor[DataCol, DataRow] := clWindow;
         if UV_vpS008[DataRow-1] = '�缺' then grdMaster.CellColor[DataCol, DataRow] := $00FFC280;
      end
      else
         grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
   end;
   if CB_HM.ItemIndex = 1 then grdMaster.Col[10].width := 0;
end;

procedure TfrmLD4Q08.btnQueryClick(Sender: TObject);
var index, I, j, iRet, rowCount : Integer;
    sSelect, sWhere, sOrderBy, cod_name, sCod_hm, spWhere : String;
    temp_S007, temp_S008 : Boolean;
    vCod_chuga, vCod_totyh : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; spWhere := '';
   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGulkwa do
   begin
      // SQL�� ����
      Close;
      sSelect := 'select a.cod_jisa, a.num_id, a.dat_gmgn,a.Desc_name, a.num_samp, a.gubn_nsyh, a.gubn_adyh, a.gubn_agyh, a.gubn_lfyh, a.cod_hul, ';
      sSelect := sSelect + ' a.cod_jangbi, a.cod_cancer, a.cod_chuga, a.gubn_nosin, a.gubn_adult, a.gubn_agab, a.gubn_life, ';
      sSelect := sSelect + ' b.desc_glkwa1, b.desc_glkwa2, b.num_bunju, C.desc_RackNo From Gumgin a with(nolock) ';
      sSelect := sSelect + ' left outer join gulkwa b with(nolock) ';
      sSelect := sSelect + ' on a.cod_jisa = b.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = b.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = b.dat_gmgn ';
      sSelect := sSelect + ' left outer join Bunju C with(nolock) ';
      sSelect := sSelect + ' on a.cod_jisa = C.cod_jisa ';
      sSelect := sSelect + ' and a.num_id = C.num_id ';
      sSelect := sSelect + ' and a.dat_gmgn = C.dat_gmgn ';

      sWhere := ' WHERE b.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' AND  b.Dat_Bunju = ''' + mskDate.Text + ''' ';

      if CB_HM.ItemIndex = 0 then
         sWhere := sWhere + ' AND  b.Gubn_Part = ''' + '07' + ''' '
      else if CB_HM.ItemIndex = 1 then
         sWhere := sWhere + ' AND  b.Gubn_Part = ''' + '05' + ''' ';


      if RB_bunju.Checked then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY B.num_bunju ';
      end
      else if RB_samp.Checked then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), B.num_bunju '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;
      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q08', 'Q', 'N');
         For I := 1 to RecordCount do
         begin
            UP_VarMemSet(index);

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
            iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for j := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_chuga[j];

            // ���, ���κ�, ������ ���� �˻��׸� ����
            cod_name := '';
            // �������Display
            if FieldByName('gubn_nosin').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
            //���κ�����Display
            if FieldByName('gubn_adult').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
            //��������Display
            if FieldByName('gubn_agab').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
            //������ȯ������Display
            if FieldByName('gubn_life').AsString = '1' then
               cod_name := cod_name + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);

            iRet := GF_MulToSingle(cod_name, 4, vCod_totyh);
            for j := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_totyh[j];


            if CB_HM.ItemIndex = 0 then
            begin
               if (pos('S007', sCod_hm) > 0) OR (pos('S008', sCod_hm) > 0) OR (pos('S099', sCod_hm) > 0) OR (pos('S091', sCod_hm) > 0) then
               begin
                  if Cmb_time.Text = '���Ű������' then
                  begin
                     with qryPGulkwa do
                     begin
                        close;
                        ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                        ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                        ParamByName('dat_bunju').AsString := mskDate.text;
                        open;
                        if RecordCount > 0 then
                        begin
                           if chk_all.Checked = True then
                           begin
                              UV_vNo[index]        := intTostr(index+1);
                              UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                              UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                              UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                              UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                              UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                              if trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '' then    //20140509
                                 UV_vS007[index] := UV_vS007[index] + ' ' + trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6));
                              UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                              if trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '' then
                                 UV_vS008[index]  := UV_vS008[index] + ' ' + trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6));  //20140509

                              if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='') or
                                 (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='����')  then
                              begin
                                  with qryPPGulkwa do
                                  begin
                                     close;
                                     ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                     ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                     ParamByName('dat_bunju').AsString := qryPGulkwa.FieldByName('dat_bunju').AsString;
                                     open;
                                     if RecordCount > 0 then
                                     begin
                                        UV_vpS007[index] := trim(copy(qryPPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))+' :'+qryPPGulkwa.FieldByName('dat_bunju').AsString;
                                        UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                        if trim(copy(qryPPGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '' then    //20140509
                                           UV_vpS007[index] := UV_vpS007[index] + ' ' + trim(copy(qryPPGulkwa.FieldByName('desc_glkwa2').AsString,29,6));
                                     end;
                                  end;
                              end
                              else
                              begin
                                  UV_vpS007[index]     := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                                  UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                                  if trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '' then    //20140509
                                     UV_vpS007[index] := UV_vpS007[index] + ' ' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6));
                              end;

                              UV_vpS008[index]     := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                              if trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '' then
                                 UV_vpS008[index]  := UV_vpS008[index] + ' ' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6));  //20140509

                              Inc(index);
                           end
                           else  //���� ��� �缺
                           begin
                              if ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '�缺') or
                                  (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '��缺')) AND
                                 ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) = '�缺') or
                                  (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) = '��缺')) then
                              begin
                                 UV_vNo[index]        := intTostr(index+1);
                                 UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                                 UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                                 UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                                 UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                                 UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                                 UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                                 if (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='') or
                                    (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='����')  then
                                 begin
                                    with qryPPGulkwa do
                                    begin
                                       close;
                                       ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                       ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                       ParamByName('dat_bunju').AsString := qryPGulkwa.FieldByName('dat_bunju').AsString;
                                       open;
                                       if RecordCount > 0 then
                                       begin
                                          UV_vPdat_gmgn[index] := (qryPPGulkwa.FieldByName('dat_gmgn').AsString);
                                          UV_vpS007[index] := trim(copy(qryPPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))+' :'+ qryPPGulkwa.FieldByName('dat_bunju').AsString;
                                       end;
                                    end;
                                 end
                                 else
                                 UV_vpS007[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                                 UV_vPdat_gmgn[index] := (qryPGulkwa.FieldByName('dat_gmgn').AsString);
                                 UV_vpS008[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                                 Inc(index);
                              end;
                           end;
                           Gride.Progress := I;
                        end;
                     end;
                  end
                  else if Cmb_time.Text = '�����׸����' then
                  begin
                     if chk_all.Checked = True then
                     begin
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                        UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                        UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                        with qryPGulkwa do
                        begin
                           close;
                           ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                           ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                           ParamByName('dat_bunju').AsString := mskDate.text;
                           open;
                           if RecordCount > 0 then
                           begin
                              if  (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='') or
                                  (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='����')  then
                              begin
                                  with qryPPGulkwa do
                                  begin
                                       close;
                                       ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                       ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                       ParamByName('dat_bunju').AsString := qryPGulkwa.FieldByName('dat_bunju').AsString;
                                       open;
                                       if RecordCount > 0 then
                                       begin
                                          UV_vpS007[index] := trim(copy(qryPPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))+' :'+ qryPPGulkwa.FieldByName('dat_bunju').AsString;
                                          UV_vPdat_gmgn[index] := (qryPPGulkwa.FieldByName('dat_gmgn').AsString);
                                          UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                       end;
                                  end;
                              end
                              else
                              UV_vpS007[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                              UV_vPdat_gmgn[index] := (qryPGulkwa.FieldByName('dat_gmgn').AsString);
                              UV_vpS008[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                           end;
                        end;
                        Inc(index);
                     end
                     else //��� �缺
                     begin
                        temp_S007 := false;
                        temp_S008 := false;

                        if (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '�缺') or
                           (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '��缺')then temp_S007 := true;
                        if (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) = '�缺') or
                           (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) = '��缺') then temp_S008 := true;

                        if ((temp_S007 = true) and
                            ((temp_S008 = true) or ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '') and
                                                    (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) > 10)))) OR
                           ((temp_S008 = true) and ((trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '') and
                                                    (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) > 1))) OR
                           (((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '') and (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6))) > 10)) and
                            ((trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '') and (StrToFloat(trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6))) > 1)))then
                        begin
                           UV_vNo[index]        := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                           UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                           UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) + ' ' + trim(copy(qryGulkwa.FieldByName('desc_glkwa2').AsString,29,6)); //20140509
                           UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) + ' ' + trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6));  //20140509
                           with qryPGulkwa do
                           begin
                              close;
                              ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                              ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                              ParamByName('dat_bunju').AsString := mskDate.text;
                              open;
                              if RecordCount > 0 then
                              begin
                                 if  (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='') or
                                      (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) ='����')  then
                                 begin
                                    with qryPPGulkwa do
                                    begin
                                      close;
                                      ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                      ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                      ParamByName('dat_bunju').AsString := qryPGulkwa.FieldByName('dat_bunju').AsString;
                                      open;
                                      if RecordCount > 0 then
                                      begin
                                         UV_vpS007[index] := trim(copy(qryPPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))+' :'+ qryPPGulkwa.FieldByName('dat_bunju').AsString;
                                         UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                         if trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <>'' then
                                            UV_vpS007[index] := UV_vpS007[index] + ':' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6));   //20140509
                                      end;
                                    end;
                                 end;
                              end
                              else
                              begin
                                UV_vpS007[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                                UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                                if trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> '' then
                                   UV_vpS007[index] := UV_vpS007[index]+ ':' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6));   //20140509
                              end;
                              UV_vpS007[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                              UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;

                              if trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6)) <> ''then
                                 UV_vpS007[index] := UV_vpS007[index] + ':' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6));   //20140509

                              UV_vpS008[index] := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));

                              if trim(copy(qryPGulkwa.FieldByName('desc_glkwa2').AsString,29,6)) <> '' then
                                 UV_vpS008[index] := UV_vpS008[index] + ':' + trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6));   //20140509
                           end;
                           Inc(index);
                        end;
                     end;
                     Gride.Progress := I;
                  end;
               end;
               next;
            end
            else if CB_HM.ItemIndex = 1 then  //S011,S012 //20140910 ������
            begin
               if (pos('S011', sCod_hm) > 0) OR (pos('S012', sCod_hm) > 0) then
               begin
                  if Cmb_time.Text = '���Ű������' then
                  begin
                     with qryPGulkwa_S011 do
                     begin
                        close;
                        ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                        ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                        ParamByName('dat_bunju').AsString := mskDate.text;
                        open;
                        if RecordCount > 0 then
                        begin
                           if chk_all.Checked = True then
                           begin
                              UV_vNo[index]        := intTostr(index+1);
                              UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                              UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                              UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                              UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                              UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                              UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));

                              if (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='') or
                                 (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='����')  then
                              begin
                                  with qryPPGulkwa_S011 do
                                  begin
                                     close;
                                     ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                     ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                     ParamByName('dat_bunju').AsString := qryPGulkwa_S011.FieldByName('dat_bunju').AsString;
                                     open;
                                     if RecordCount > 0 then
                                        UV_vpS007[index]      := trim(copy(qryPPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                        UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                  end;
                              end
                              else  UV_vpS007[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                              UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                              UV_vpS008[index]     := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,31,6));
                              Inc(index);
                           end
                           else if chk_yang.Checked then
                           begin
                              if ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) = '�缺') or
                                  (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) = '��缺')) AND
                                 ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '�缺') or
                                  (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '��缺')) then
                              begin
                                 UV_vNo[index]        := intTostr(index+1);
                                 UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                                 UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                                 UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                                 UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                                 UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                                 UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));

                                 if (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='') or
                                    (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='����')  then
                                 begin
                                    with qryPPGulkwa_S011 do
                                    begin
                                       close;
                                       ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                       ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                       ParamByName('dat_bunju').AsString := qryPGulkwa_S011.FieldByName('dat_bunju').AsString;
                                       open;
                                       if RecordCount > 0 then
                                          UV_vpS007[index] := trim(copy(qryPPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                          UV_vpDat_gmgn[index] := qryPPGulkwa_S011.FieldByName('dat_gmgn').AsString;
                                    end;
                                 end
                                 else
                                 UV_vpS007[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                 UV_vpDat_gmgn[index] := qryPGulkwa_S011.FieldByName('dat_gmgn').AsString;
                                 UV_vpS008[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,31,6));
                                 Inc(index);
                              end;
                           end;
                           Gride.Progress := I;
                        end;
                     end;
                  end
                  else if Cmb_time.Text = '�����׸����' then
                  begin
                     if chk_all.Checked = True then
                     begin
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                        UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                        UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                        with qryPGulkwa_S011 do
                        begin
                           close;
                           ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                           ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                           ParamByName('dat_bunju').AsString := mskDate.text;
                           open;
                           if RecordCount > 0 then
                           begin
                              if  (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='') or
                                  (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='����')  then
                              begin
                                  with qryPPGulkwa_S011 do
                                  begin
                                       close;
                                       ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                       ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                       ParamByName('dat_bunju').AsString := qryPGulkwa_S011.FieldByName('dat_bunju').AsString;
                                       open;
                                       if RecordCount > 0 then
                                          UV_vpS007[index] := trim(copy(qryPPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                          UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                  end;
                              end
                              else     UV_vpS007[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                              UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                              UV_vpS008[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,31,6));
                           end;
                        end;
                        Inc(index);
                     end
                     else if  chk_yang.Checked then
                     begin
                        if ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) = '��缺') or
                            (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) = '�缺')) AND
                           ((trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '��缺') or
                            (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '�缺')) then
                        begin
                           UV_vNo[index]        := intTostr(index+1);
                           UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                           UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                           UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                           UV_vNum_Rack[index]  := qryGulkwa.FieldByName('Desc_RackNo').AsString;
                           UV_vS007[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,25,6));
                           UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));

                           with qryPGulkwa_S011 do
                           begin
                              close;
                              ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                              ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                              ParamByName('dat_bunju').AsString := mskDate.text;
                              open;
                              if RecordCount > 0 then
                              begin
                                 if  (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='') or
                                     (trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6)) ='����')  then
                                 begin
                                    with qryPPGulkwa_S011 do
                                    begin
                                         close;
                                         ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                                         ParamByName('num_id').AsString    := qryGulkwa.FieldByName('num_id').AsString;
                                         ParamByName('dat_bunju').AsString := qryPGulkwa_S011.FieldByName('dat_bunju').AsString;
                                         open;
                                         if RecordCount > 0 then
                                            UV_vpS007[index] := trim(copy(qryPPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                            UV_vpDat_gmgn[index] := qryPPGulkwa.FieldByName('dat_gmgn').AsString;
                                    end;
                                 end
                                 else    UV_vpS007[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,25,6));
                                 UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_gmgn').AsString;
                                 UV_vpS008[index] := trim(copy(qryPGulkwa_S011.FieldByName('desc_glkwa1').AsString,31,6));
                              end;
                           end;
                           Inc(index);
                        end;
                     end;
                  end; // ����/����
               end; //S011, S012
               Gride.Progress := I;
               next;
            end; // CB_HM
         end;
         Gride.Progress := RecordCount;
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGulkwa

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q08.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);

   if RB_bunju.Checked then
   begin
      mskFrom.Enabled := True;
      mskTo.Enabled := True;
      MEdt_SampS.Enabled := False;
      MEdt_SampE.Enabled := False;
   end
   else if RB_samp.Checked then
   begin
      mskFrom.Enabled := False;
      mskTo.Enabled := False;
      MEdt_SampS.Enabled := True;
      MEdt_SampE.Enabled := True;
   end;
end;

procedure TfrmLD4Q08.btnPrintClick(Sender: TObject);
begin
  inherited;

  if RG_print.ItemIndex = 0 then
  begin
     frmLD4Q081 := TfrmLD4Q081.Create(Self);
     frmLD4Q081.QuickRep.Preview;
  end
  else if RG_print.ItemIndex = 1 then
  begin
     frmLD4Q081 := TfrmLD4Q081.Create(Self);
     frmLD4Q081.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q08.SBut_ExcelClick(Sender: TObject);
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