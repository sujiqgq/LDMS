//==============================================================================
// ���α׷� ���� : ���װ��üũ insert �κ� �߰�
// �ý���        : ���հ���
// ��������      : 2004.10.25
// ������        : ������
// ��������      :
// �������      : 
//==============================================================================
// ��������      : 2006.02.04
// ������        : ���ö
// ��������      : ��ü�ΰ�� ���ְ� �ȵ��ư��Բ� ���� (gubn_neawon <> '5')
//==============================================================================
// ��������      : 2006.05.17
// ������        : ������
// ��������      : ������ü �����Ҷ� ������� �����ϱ�..
//==============================================================================
// ��������      : 2006.06.29
// ������        : ������
// ��������      : ���'CV,CM' = 5000, 'SS,GS' = 9000 �������� �˻籸������ ��������..
//==============================================================================
// ��������      : 2007.10.26
// ������        : ���ö
// ��������      : cell �ڸ��� �������� 9999->99999 �� ����. (�⵵�� ���ڸ��� ������)
//==============================================================================
// ��������      : 2007.11.29
// ������        : ������
// ��������      : �����ڵ尡 ���� �߰��ڵ常 ������ ���ֹ�ȣ �����ȵǰ�...
//==============================================================================
// ��������      : 2008.10.30
// ������        : ���ö
// ��������      : 1000���� �ߺ��Ǵ� ���ֹ�ȣ �߻����� ���Ͽ� �˻籸�� 10,20 �� ���� ���ֹ�ȣ�� ���.
//                 ���漾���� ��� ��ü������ ���ֵ��ư��°� ����.
//==============================================================================
// ��������      : 2011.12.21
// ������        : ������
// ��������      : [����]��ü�˻� �ǽð��� �߰�...(qryHgulkwa_EtcChk)
//==============================================================================
// ��������      : 2013.5.21
// ������        : ����
// ��������      : Rack no  �߰�
//==============================================================================
// ��������      : 2014.04.16
// ������        : ������
// ��������      : �߱����� F2, H2 �׸� ���� Rack no-> 'D'�� �߰�
//==============================================================================
// ��������      : 2014.04.24
// ������        : ������
// ��������      : ���� ����� ������ �ƴϱ⿡ ��������� ���޿��� ���ܵ� �� �ֵ��� ������û
//                (���� �������� 4610)
//==============================================================================
// ��������      : 2014.09.29
// ������        : ������
// ��������      : TT01, TT02, C077, C083 ����ȣ �ο� ����
//==============================================================================
// ��������      : 2014.10.01
// ������        : ������
// ��������      : T006 ����ȣ �ο� ����
//==============================================================================
// ��������      : 2015.04.09~11
// ������        : ������
// ��������      : ���,���κ� ���&C032 ������ B�� �ο� �ȵǵ��� ������û
//                 [���� �������ܰ˻�������1500248 ���� - �������� ��û]
//==============================================================================
// ��������      : 2015.09.19
// ������        : ������
// ��������      : �߱����� �˻翡 CV05 �������� �߰�
//                 [���� ��ܱ۷ι��ｺ�ɾ���1500032 - �������� ��û]
//==============================================================================
//==============================================================================
// ��������      : 2015.12.30
// ������        : ��â��
// ��������      : 7000-8500 �̻��̸� 8000����� �̵�
//==============================================================================
// ��������      : 2017-10-12
// ������        : �ڼ���
// ��������      : 2017-10-16���� ���� ���� ��ȭ�� ��ü�˻� ����
//==============================================================================

unit LD2I01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges,
  Spin, jpeg;

type
  TfrmLD2I01 = class(TfrmSingle)
    qryGumgin: TQuery;
    pnlMaster: TPanel;
    Panel2: TPanel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    Gauge: TGauge;
    Animate1: TAnimate;
    Animate2: TAnimate;
    Label1: TLabel;
    Panel3: TPanel;
    labStatus: TLabel;
    qryProfile: TQuery;
    qryU_Gumgin: TQuery;
    qryBunju: TQuery;
    qryI_Bunju: TQuery;
    qryI_Gulkwa: TQuery;
    qryI_Cell: TQuery;
    qryHangmok: TQuery;
    qrySeq: TQuery;
    qryTot_Hangmok_List: TQuery;
    qryNo_hangmok: TQuery;
    qryBjValue: TQuery;
    qry_cell: TQuery;
    qryHgulkwa_chk: TQuery;
    qryIHgulkwa_chk: TQuery;
    GroupBox1: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Image1: TImage;
    SpeedButton1: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    SpinEdit3: TSpinEdit;
    SpinEdit4: TSpinEdit;
    SpinEdit5: TSpinEdit;
    SpinEdit2: TSpinEdit;
    SpinEdit1: TSpinEdit;
    Timer1: TTimer;
    Timer2: TTimer;
    qryHgulkwa_EtcChk: TQuery;
    qryIHgulkwa_EtcChk: TQuery;
    qryI_Allergy: TQuery;
    qry_Allergy: TQuery;
    Gauge1: TGauge;
    Label2: TLabel;
    qry_Bunju_Rack: TQuery;
    Qry_uRackNo: TQuery;
    qryJHangmokList: TQuery;
    Chk_Rack: TCheckBox;
    qryI_SMS: TQuery;
    qryBunju_Count: TQuery;
    Panel4: TPanel;
    btnStart: TBitBtn;
    Label3: TLabel;
    qryI_Allergy62: TQuery;
    qryI_Allergy107: TQuery;
    procedure UP_Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure UP_Change(Sender: TObject);
  private
    { Private declarations }
    function  UF_BunjuNo(iBunjuNo, iBunjuNoe : Integer) : Integer;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

    procedure Gubn_rack;
    procedure UP_num_rack(sGubn :String; sNum_bunju: integer; sDat_bunju : string);

  public
    // SQL �ӽú���
    UV_sBasicSQL : String;
    // �����ڵ�
    UV_sJisa : String;
    num_ARack, num_ACol, num_Arow, num_BRack, num_Brow, num_BCol, num_CRack, num_CCol, num_Crow,
    num_DRack, num_DCol, num_Drow :Integer;
    { Public declarations }
  end;

var
  frmLD2I01: TfrmLD2I01;

implementation

uses ZZFunc, ZZMenu, ZZComm;

{$R *.DFM}

procedure TfrmLD2I01.UP_Click(Sender: TObject);
begin
   inherited;

   // Button Click�� Event Procedure�� ������ �� ������ Sender�� ����
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD2I01.FormCreate(Sender: TObject);
begin
   inherited;

   // �������ڸ� ����
   mskDate.Text := GV_sToday;

   // SQL�� ����
   UV_sBasicSQL := qryGumgin.SQL.Text;

   SpinEdit3.Value := StrToInt(formatdatetime('YYYY', now));
   SpinEdit4.Value := StrToInt(formatdatetime('MM', now));
   SpinEdit5.Value := StrToInt(formatdatetime('DD', now));
   SpinEdit1.Value := StrToInt(formatdatetime('HH', now));
   SpinEdit2.Value := StrToInt(formatdatetime('NN', now));

   If (GV_sJisa = '15') Or (GV_sJisa = '00') Then Chk_Rack.Visible := True
   Else                                           Chk_Rack.Visible := False;;
end;

procedure TfrmLD2I01.btnStartClick(Sender: TObject);
var i, iCod_chuga, bunju_no, Jbunju_no, iRet : Integer;
    sSelect, sWhere, sOrderBy, sBunCheck, sJangCk, sCode : String;
    iBun10, iBun20, iBun30, iBun60, iBun70, iBun50, iBun90 : Integer;
    ijBun10, ijBun20, ijBun30, ijBun60, ijBun70, ijBun50, ijBun90 : Integer;
    vPart, vCod_chuga, vCod_hm: Variant;
    hul_chk : Boolean;

    iTemp : Integer;
begin
   inherited;

   //rack �ʱ�ȭ
   num_ARack:=1; num_ACol :=1; num_Arow :=1;
   num_BRack:=1; num_Brow :=1; num_BCol :=1;
   num_CRack:=1; num_Crow :=1; num_CCol :=1;
   num_DRack:=1; num_Drow :=1; num_DCol :=1;

   //10.12.09.ö. ������ֽ� dat_insert�� ���� �������ڸ� ����(Server System Date)
   with dmComm.qrySysDate do
   begin
      if not Active then Active := True;
      GV_sToday := FieldByName('SYSDATE').AsString;
      // Table disconnect
      Active := False;
   end;

   // ��ü(00)�ϰ�� ���縦 �������� �۾�
   if GV_sJisa = '00' then UV_sJisa := Copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   If Chk_Rack.Checked Then
   Begin
      Gubn_Rack;
      Showmessage('�� Rack ��ȣ�� �ο� �Ϸ� �Ͽ����ϴ�.~~~');
      Exit;
   End;

   //���ֹ�ȣ���Ⱚ
   if UV_sJisa = '15' then
   begin
   // �������簡 ������ ���...
      for i := 1 to 7 do
      begin
         with qryBjValue do
         begin
            Active := False;
            ParamByName('cod_bjjs').AsString  := UV_sJisa;
            ParamByName('dat_bunju').AsString := mskDate.Text;
            case i of
               1 : begin ParamByName('snum').AsString := '0';    ParamByName('enum').AsString := '3999'; end;
               3 : begin ParamByName('snum').AsString := '4000'; ParamByName('enum').AsString := '4999'; end;
               4 : begin ParamByName('snum').AsString := '5000'; ParamByName('enum').AsString := '5999'; end;
               5 : begin ParamByName('snum').AsString := '6000'; ParamByName('enum').AsString := '6999'; end;
               6 : begin ParamByName('snum').AsString := '7000'; ParamByName('enum').AsString := '8499'; end;
               7 : begin ParamByName('snum').AsString := '8500'; ParamByName('enum').AsString := '9999'; end;
            end;
            Active := True;

            case i of
               1:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun10 := FieldByName('bjvalue').AsInteger + 1 else iBun10 := 1;    end;
               3:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun30 := FieldByName('bjvalue').AsInteger + 1 else iBun30 := 4001; end;
               4:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun50 := FieldByName('bjvalue').AsInteger + 1 else iBun50 := 5001; end;
               5:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun60 := FieldByName('bjvalue').AsInteger + 1 else iBun60 := 6001; end;
               6:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun70 := FieldByName('bjvalue').AsInteger + 1 else iBun70 := 7001; end;
               7:  begin if Trim(FieldByName('bjvalue').AsString) <> '' then iBun90 := FieldByName('bjvalue').AsInteger + 1 else iBun90 := 8501; end;
            end;
          end;
      end;
   end
   else
   // �������簡 ������ ���...
   begin
      for i := 1 to 7 do
      begin
         with qryBjValue do
         begin
            Active := False;
            ParamByName('cod_bjjs').AsString  := UV_sJisa;
            ParamByName('dat_bunju').AsString := mskDate.Text;
            case i of
               1 : begin ParamByName('snum').AsString := '0';    ParamByName('enum').AsString := '0999'; end;
               2 : begin ParamByName('snum').AsString := '1000'; ParamByName('enum').AsString := '3999'; end;
               3 : begin ParamByName('snum').AsString := '4000'; ParamByName('enum').AsString := '4999'; end;
               4 : begin ParamByName('snum').AsString := '5000'; ParamByName('enum').AsString := '5999'; end;
               5 : begin ParamByName('snum').AsString := '6000'; ParamByName('enum').AsString := '6999'; end;
               6 : begin ParamByName('snum').AsString := '7000'; ParamByName('enum').AsString := '8499'; end;
               7 : begin ParamByName('snum').AsString := '8500'; ParamByName('enum').AsString := '9999'; end;
            end;
            Active := True;

            case i of
               1: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun10 := FieldByName('bjvalue').AsInteger + 1 else ijBun10 := 1;    end;
               2: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun20 := FieldByName('bjvalue').AsInteger + 1 else ijBun20 := 1001; end;
               3: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun30 := FieldByName('bjvalue').AsInteger + 1 else ijBun30 := 4001; end;
               4: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun50 := FieldByName('bjvalue').AsInteger + 1 else ijBun50 := 5001; end;
               5: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun60 := FieldByName('bjvalue').AsInteger + 1 else ijBun60 := 6001; end;
               6: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun70 := FieldByName('bjvalue').AsInteger + 1 else ijBun70 := 7001; end;
               7: begin if Trim(FieldByName('bjvalue').AsString) <> '' then ijBun90 := FieldByName('bjvalue').AsInteger + 1 else ijBun90 := 8501; end;
            end;
          end;
      end;
   end;

   // ���ڰ� ���õǾ����� Check
   if mskDate.Text  = '' then
   begin
      GF_MsgBox('Warning', '�۾����ڴ� �ݵ�� �Է��ؾ� �մϴ�.');
      mskDate.SetFocus;
      exit;
   end;

   if Label6.Caption = 'Ÿ�̸� ���۵� ���Դϴ�..'  then
   begin
      if not GF_MsgBox('Confirmation', '���׺��ָ� �����Ͻðڽ��ϱ� ?') then exit;
   end;

   with qryHangmok do
   begin
      Active := False;
      ParamByName('dat_apply').Asstring := mskDate.Text;
      Active := True;
   end;

   // ������ֽ� ������ ���ֿ��� Check.
   if UV_SJisa <> '15' then
   begin
      qryBunju.Active := False;
      qryBunju.ParamByName('cod_bjjs').AsString := '15';
      qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
      qryBunju.Active := True;
      if qryBunju.RecordCount = 0 then
      begin
         GF_MsgBox('Warning', '���� ���ֽǿ� �����ϼ���...');
         qryBunju.Active := False;
         Exit;
      end;
   end;

   // ���ֿ��� Check.
   qryBunju.Active := False;
   if GV_sJisa = '00' then qryBunju.ParamByName('cod_bjjs').AsString := Copy(GV_sUserId,1,2)
   else                    qryBunju.ParamByName('cod_bjjs').AsString := GV_sJisa;
   qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
   qryBunju.Active := True;
   if qryBunju.RecordCount > 0 then
   begin
      GF_MsgBox('Warning', '�����۾��� �� �����Դϴ�.');
      qryBunju.Active := False;
      Exit;
   end;

   // DB �۾�
   dmComm.database.StartTransaction;
   try
      // Query ���� & �迭�� ����
      with qryGumgin do
      begin
         Active := False;

         if UV_sJisa = '15' then
         begin
            sWhere := ' WHERE G.ysno_bunju = ''N'' ';
            sWhere := sWhere + ' AND G.dat_gmgn < ''' + mskDate.Text + ''' ';      //����
            sWhere := sWhere + ' AND G.gubn_neawon <> ''5'' ';
            sOrderBy := ' ORDER BY G.cod_jisa, G.cod_hul, G.num_samp';
         end
         else
         begin
            sWhere := ' WHERE G.cod_jisa = ''' + UV_sJisa + ''' AND G.ysno_bunju = ''N'' ';
            sWhere := sWhere + ' AND G.dat_gmgn < ''' + mskDate.Text + ''' ';      //����
            sWhere := sWhere + ' AND G.gubn_neawon <> ''5'' ';
            sOrderBy := ' ORDER BY G.cod_jisa, G.cod_hul, G.num_samp';
         end;

         SQL.Clear;
         SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
         Active := True;

         // Animate Start
         Animate1.Active := True;
         Animate2.Active := True;
         Gauge.Progress  := 0;
         Gauge.MaxValue  := RecordCount;


         GP_query_log(qryGumgin, 'LD2I01', 'Q', 'N');
         // Routine...
         while not EOF do
         begin
            Gauge.Progress := Gauge.Progress + 1;

            labStatus.Caption := 'C:'   + qryGumgin.FieldByName('cod_jisa').AsString +
                                 ' ID:' + qryGumgin.FieldByName('num_id').AsString   +
                                 ' D:'  + qryGumgin.FieldByName('Dat_gmgn').AsString +
                                 ' N:'  + qryGumgin.FieldByName('Desc_name').AsString;
            Application.ProcessMessages;

            bunju_no  := 0;
            Jbunju_no := 0;
            sBunCheck := '';

            // Part Setting
            vPart   := VarArrayCreate([0, 0], varOleStr);
            vCod_hm := VarArrayCreate([0, 0], varOleStr);

            VarArrayRedim(vPart, 10);

            //20140530 ���̺� �̻��� ����..
            for i:=1 to 10 do vPart[i] := '';

            if UV_sJisa = '15' then
            begin
               if (FieldByName('gubn_hulgum').AsString <> '1') and
                  (FieldByName('gubn_hulgum').AsString <> '3') and
                  (FieldByName('cod_jisa').AsString <> '15') then
               begin
                  Next;
                  Continue;
               end;
            end
            else
            begin
               if (UV_sJisa <> FieldByName('cod_jisa').AsString) or
                  (FieldByName('gubn_hulgum').AsString = '1') or
                  (FieldByName('gubn_hulgum').AsString = '3') then
               begin
                  Next;
                  Continue;
               end;
            end;

            //���ֹ�ȣ ����
            if Trim(FieldByName('cod_hul').AsString) <> '' then
            begin
               qryProfile.Active := False;
               qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_hul').AsString;
               qryProfile.Active := True;
               if qryProfile.RecordCount > 0 then
               begin
                  // ���� -> ���ֹ�ȣ ����...
                  if UV_sJisa = '15' then
                  begin
                     bunju_no := qryProfile.FieldByName('gubn_gmsa').AsInteger;

                     //ä��
                     if Copy(FieldByName('cod_hul').AsString,1,2) = 'SS' then bunju_no := 90;

                     //��ü���� ó��...
                     if ((qryGumgin.FieldByName('cod_dc').AsString) = '617097') or ((qryGumgin.FieldByName('cod_dc').AsString) = '617098') then bunju_no := 50;

                     //����ó��..
                     if Trim(FieldByName('HUL_CK').AsString) = 'Y' then bunju_no := 30;

                     case bunju_no of
                        10, 20 : begin sBunCheck := IntToStr(iBun10); iBun10 := iBun10 + 1; end;
                            30 : begin sBunCheck := IntToStr(iBun30); iBun30 := iBun30 + 1; end;
                            50 : begin sBunCheck := IntToStr(iBun50); iBun50 := iBun50 + 1; end;
                            60 : begin sBunCheck := IntToStr(iBun60); iBun60 := iBun60 + 1; end;
                            70 : begin
                                    //20151230-��â��
                                    If iBun70 = 8500 Then Begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; End
                                    Else                  Begin sBunCheck := IntToStr(iBun70); iBun70 := iBun70 + 1; End;
                                 end;
                            90 : begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; end;
                     end;
                  end
                  else
                  // ���� -> ���ֹ�ȣ ����...
                  begin
                     case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                        10 : begin sBunCheck := IntToStr(ijBun10); ijBun10 := ijBun10 + 1; end;
                        20 : begin sBunCheck := IntToStr(ijBun20); ijBun20 := ijBun20 + 1; end;
                        30 : begin sBunCheck := IntToStr(ijBun30); ijBun30 := ijBun30 + 1; end;
                        50 : begin sBunCheck := IntToStr(ijBun50); ijBun50 := ijBun50 + 1; end;
                        60 : begin sBunCheck := IntToStr(ijBun60); ijBun60 := ijBun60 + 1; end;
                        70 : begin
                                //20151230-��â��
                                If ijBun70 = 8500 Then Begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; End
                                Else                   Begin sBunCheck := IntToStr(ijBun70); ijBun70 := ijBun70 + 1; End;
                             end;
                        90 : begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; end;
                     end;
                  end;

                  // ���� Table Insert.
                  sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                  qryI_Bunju.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                  qryI_Bunju.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                  qryI_Bunju.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                  qryI_Bunju.ParamByName('num_bunju').Asstring  := sBunCheck;
                  qryI_Bunju.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                  qryI_Bunju.ParamByName('gubn_gmsa').Asstring  := qryProfile.FieldByName('gubn_gmsa').AsString;
                  qryI_Bunju.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                  qryI_Bunju.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                  qryI_Bunju.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                  qryI_Bunju.ExecSQL;
                  GP_query_log(qryI_Bunju, 'LD2I01', 'Q', 'Y');
              end;
            end
            //����ȣ ����
            else if (Trim(FieldByName('cod_jangbi').AsString) <> '') then
            begin
               qryProfile.Active := False;
               qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_jangbi').AsString;
               qryProfile.Active := True;

               if qryProfile.RecordCount > 0 then
               begin
                  if UV_sJisa = '15' then
                  begin
                     Jbunju_no := qryProfile.FieldByName('gubn_gmsa').AsInteger;

                     if (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'SS') or
                        (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'GS') then
                     begin
                        Jbunju_no := 90;
                     end;

                     //����ó��..
                     if Trim(FieldByName('HUL_CK').AsString) = 'Y' then Jbunju_no := 30;

                     case Jbunju_no of
                        10, 20 : begin sBunCheck := IntToStr(iBun10); iBun10 := iBun10 + 1; end;
                            30 : begin sBunCheck := IntToStr(iBun30); iBun30 := iBun30 + 1; end;
                            50 : begin sBunCheck := IntToStr(iBun50); iBun50 := iBun50 + 1; end;
                            60 : begin sBunCheck := IntToStr(iBun60); iBun60 := iBun60 + 1; end;
                            70 : begin
                                    //20151230-��â��
                                    If iBun70 = 8500 Then Begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; End
                                    Else                  Begin sBunCheck := IntToStr(iBun70); iBun70 := iBun70 + 1; End;
                                 end;
                            90 : begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; end;
                     end;
                  end
                  else
                  begin
                     case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                        10 : begin sBunCheck := IntToStr(ijBun10); ijBun10 := ijBun10 + 1; end;
                        20 : begin sBunCheck := IntToStr(ijBun20); ijBun20 := ijBun20 + 1; end;
                        30 : begin sBunCheck := IntToStr(ijBun30); ijBun30 := ijBun30 + 1; end;
                        50 : begin sBunCheck := IntToStr(ijBun50); ijBun50 := ijBun50 + 1; end;
                        60 : begin sBunCheck := IntToStr(ijBun60); ijBun60 := ijBun60 + 1; end;
                        70 : begin
                                //20151230-��â��
                                If ijBun70 = 8500 Then Begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; End
                                Else                   Begin sBunCheck := IntToStr(ijBun70); ijBun70 := ijBun70 + 1; End;
                             end;
                        90 : begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; end;
                     end;
                  end;

                  // ���� Table Insert.
                  sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                  qryI_Bunju.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                  qryI_Bunju.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                  qryI_Bunju.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                  qryI_Bunju.ParamByName('num_bunju').Asstring  := sBunCheck;
                  qryI_Bunju.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                  qryI_Bunju.ParamByName('gubn_gmsa').Asstring  := qryProfile.FieldByName('gubn_gmsa').AsString;
                  qryI_Bunju.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                  qryI_Bunju.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                  qryI_Bunju.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                  qryI_Bunju.ExecSQL;
                  GP_query_log(qryI_Bunju, 'LD2I01', 'Q', 'Y');
               end;
            end
            //�����ڵ� Ȯ��
            else if Trim(FieldByName('cod_cancer').AsString) <> '' then
            begin
               qryProfile.Active := False;
               qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_cancer').AsString;
               qryProfile.Active := True;
               if qryProfile.RecordCount > 0 then
               begin
                  if sBunCheck = '' then
                  begin
                     if UV_sJisa = '15' then
                     begin
                        case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                        60 : begin
                                sBunCheck := IntToStr(iBun60);
                                iBun60 := iBun60 + 1;
                             end;
                        end;
                     end
                     else
                     begin
                        case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                        60 : begin
                                sBunCheck := IntToStr(ijBun60);
                                ijBun60 := ijBun60 + 1;
                             end;
                        end;
                     end;

                     // ���� Table Insert.
                     sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                     qryI_Bunju.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                     qryI_Bunju.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                     qryI_Bunju.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                     qryI_Bunju.ParamByName('num_bunju').Asstring  := sBunCheck;
                     qryI_Bunju.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                     qryI_Bunju.ParamByName('gubn_gmsa').Asstring  := qryProfile.FieldByName('gubn_gmsa').AsString;
                     qryI_Bunju.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                     qryI_Bunju.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                     qryI_Bunju.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                     qryI_Bunju.ExecSQL;
                     GP_query_log(qryI_Bunju, 'LD2I01', 'Q', 'Y');
                  end;
               end;
            end
            else if (FieldByName('gubn_nosin').AsString >= '1')
                 or (FieldByName('gubn_tkgm').AsString  >= '1')
                 or (FieldByName('gubn_bogun').AsString >= '1')
                 or (FieldByName('gubn_adult').AsString >= '1')
                 or (FieldByName('gubn_agab').AsString  >= '1')
                 or (FieldByName('gubn_life').AsString  >= '1') then
            begin
               if sBunCheck = '' then
               begin
                  if UV_sJisa = '15' then
                  begin
                     //20151230-��â��
                     If iBun70 = 8500 Then Begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; End
                     Else                  Begin sBunCheck := IntToStr(iBun70); iBun70 := iBun70 + 1; End;
                  end
                  else
                  begin
                     //20151230-��â��
                     If ijBun70 = 8500 Then Begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; End
                     Else                   Begin sBunCheck := IntToStr(ijBun70); ijBun70 := ijBun70 + 1; End;
                  end;

                  // ���� Table Insert.
                  sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                  qryI_Bunju.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                  qryI_Bunju.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                  qryI_Bunju.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                  qryI_Bunju.ParamByName('num_bunju').Asstring  := sBunCheck;
                  qryI_Bunju.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                  qryI_Bunju.ParamByName('gubn_gmsa').Asstring  := '80';
                  qryI_Bunju.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                  qryI_Bunju.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                  qryI_Bunju.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                  qryI_Bunju.ExecSQL;
                  GP_query_log(qryI_Bunju, 'LD2I01', 'Q', 'Y');
               end;
            end
            //�߰��ڵ� Ȯ��
            else if (Trim(FieldByName('cod_chuga').AsString) <> '')
                 and ((FieldByName('gubn_nosin').AsString < '1')
                 and (FieldByName('gubn_tkgm').AsString   < '1')
                 and (FieldByName('gubn_bogun').AsString  < '1')
                 and (FieldByName('gubn_adult').AsString  < '1')
                 and (FieldByName('gubn_agab').AsString   < '1')
                 and (FieldByName('gubn_life').AsString   < '1')) then
            begin
               hul_chk := false;
               iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
               VarArrayRedim(vCod_hm, iRet);
               for i := 0 to iRet-1 do
               begin
                    vCod_hm[i] := vCod_chuga[i];

                    // �˻��׸�
                    qryHangmok.Filter := 'COD_HM = ''' + vCod_hm[i] + '''' +
                                       ' AND DAT_APPLY <= ''' + FieldByName('DAT_GMGN').AsString + '''';

                    // ���װ˻����� Check
                    if qryHangmok.FieldByName('GUBN_PART').AsString < '10' then begin hul_chk := true; break; end
                    else                                                              hul_chk := false;
               end;

               if hul_chk = true then
               begin
                  if sBunCheck = '' then
                  begin
                     if UV_sJisa = '15' then
                     begin
                        //20151230-��â��
                        If iBun70 = 8500 Then Begin sBunCheck := IntToStr(iBun90); iBun90 := iBun90 + 1; End
                        Else                  Begin sBunCheck := IntToStr(iBun70); iBun70 := iBun70 + 1; End;
                     end
                     else
                     begin
                        //20151230-��â��
                        If ijBun70 = 8500 Then Begin sBunCheck := IntToStr(ijBun90); ijBun90 := ijBun90 + 1; End
                        Else                   Begin sBunCheck := IntToStr(ijBun70); ijBun70 := ijBun70 + 1; End;
                     end;

                     // ���� Table Insert.
                     sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                     qryI_Bunju.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                     qryI_Bunju.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                     qryI_Bunju.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                     qryI_Bunju.ParamByName('num_bunju').Asstring  := sBunCheck;
                     qryI_Bunju.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                     qryI_Bunju.ParamByName('gubn_gmsa').Asstring  := '80';
                     qryI_Bunju.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                     qryI_Bunju.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                     qryI_Bunju.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                     qryI_Bunju.ExecSQL;
                     GP_query_log(qryI_Bunju, 'LD2I01', 'Q', 'Y');
                  end;
               end;
            end;

            //�˻��׸����� ��� ��Ʈ �����ϱ�..
            //------------------------------------------------------------------
            qryTot_Hangmok_List.Active := False;

            qryTot_Hangmok_List.ParamByName('Cod_jisa').AsString  := FieldByName('Cod_jisa').AsString;
            qryTot_Hangmok_List.ParamByName('num_id').AsString    := FieldByName('num_id').AsString;
            qryTot_Hangmok_List.ParamByName('Dat_gmgn').AsString  := FieldByName('Dat_gmgn').AsString;
            qryTot_Hangmok_List.ParamByName('Mode').AsString      := '1';

            qryTot_Hangmok_List.Active := True;

            if qryTot_Hangmok_List.RecordCount > 0 then
            begin
               while not qryTot_Hangmok_List.Eof do
               begin
                  sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');

                  qryHangmok.Filter := '';
                  qryHangmok.Filter := ' COD_HM = ''' + qryTot_Hangmok_List.FieldByName('cod_hm').AsString + ''' AND DAT_APPLY <= ''' + FieldByName('DAT_GMGN').AsString + '''';
                  if qryHangmok.RecordCount > 0 then
                  begin
                     qryI_Gulkwa.ParamByName('cod_bjjs').Asstring    := UV_sJisa;
                     qryI_Gulkwa.ParamByName('num_id').Asstring      := FieldByName('num_id').AsString;
                     qryI_Gulkwa.ParamByName('dat_bunju').Asstring   := mskDate.Text;
                     qryI_Gulkwa.ParamByName('num_bunju').Asstring   := sBunCheck;
                     qryI_Gulkwa.ParamByName('dat_gmgn').Asstring    := FieldByName('dat_gmgn').AsString;
                     qryI_Gulkwa.ParamByName('cod_jisa').Asstring    := FieldByName('cod_jisa').AsString;
                     qryI_Gulkwa.ParamByName('gubn_part').Asstring   := 'N';
                     qryI_Gulkwa.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                     qryI_Gulkwa.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                     Case qryHangmok.FieldByName('gubn_part').AsInteger of
                        01 : if Trim(vPart[1]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '01'; vPart[1] := 'Y'; end;
                        02 : if Trim(vPart[2]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '02'; vPart[2] := 'Y'; end;
                        03 : if Trim(vPart[3]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '03'; vPart[3] := 'Y'; end;
                        04 : if Trim(vPart[4]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '04'; vPart[4] := 'Y'; end;
                        05 : if Trim(vPart[5]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '05'; vPart[5] := 'Y'; end;
                        07 : if Trim(vPart[7]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '07'; vPart[7] := 'Y'; end;
                        08 : if Trim(vPart[8]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '08'; vPart[8] := 'Y'; end;
                        09 : if Trim(vPart[9]) = '' then begin qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '09'; vPart[9] := 'Y'; end;
                        06 : begin
                                qry_cell.Active := False;
                                qry_cell.ParamByName('num_id').AsString   := FieldByName('num_id').AsString;
                                qry_cell.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                                qry_cell.ParamByName('cod_hm').AsString   := qryHangmok.FieldByName('cod_hm').AsString;
                                qry_cell.Active := True;
                                if qry_cell.RecordCount = 0 then
                                begin
                                   qryI_Cell.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                                   qryI_Cell.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                                   qryI_Cell.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                                   qryI_Cell.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                                   qryI_Cell.ParamByName('cod_hm').Asstring     := qryHangmok.FieldByName('cod_hm').AsString;
                                   qryI_Cell.ParamByName('ysno_sokun').Asstring := 'N';
                                   qryI_Cell.ParamByName('DAT_INSERT').AsString := GV_sToday;
                                   qryI_Cell.ParamByName('COD_INSERT').AsString := GV_sUserId;

                                   qrySeq.Active := False;
                                   if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                   begin
                                      qrySeq.ParamByName('COD_CELLS').AsString := 'J' + Copy(mskDate.Text, 3, 4) + '000000';
                                      qrySeq.ParamByName('COD_CELLE').AsString := 'J' + Copy(mskDate.Text, 3, 4) + '999999';
                                   end
                                   else
                                   begin
                                      qrySeq.ParamByName('COD_CELLS').AsString := 'P' + Copy(mskDate.Text, 3, 4) + '000000';
                                      qrySeq.ParamByName('COD_CELLE').AsString := 'P' + Copy(mskDate.Text, 3, 4) + '999999';
                                   end;
                                   qrySeq.Active := True;

                                   if Trim(qrySeq.FieldByName('RESULT').AsString) = '' then
                                   begin
                                      if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                         qryI_Cell.ParamByName('cod_cell').Asstring := 'J' + Copy(mskDate.Text, 3, 4) + '000001'
                                      else
                                         qryI_Cell.ParamByName('cod_cell').Asstring := 'P' + Copy(mskDate.Text, 3, 4) + '000001';
                                   end
                                   else
                                      qryI_Cell.ParamByName('cod_cell').Asstring := Copy(qrySeq.FieldByName('RESULT').AsString, 1, 5) +  GF_LPad(IntToStr(StrToInt(Copy(qrySeq.FieldByName('RESULT').AsString, 6, 6)) + 1), 6, '0');

                                   qryI_Cell.ExecSQL;
                                   GP_query_log(qryI_Cell, 'LD2I01', 'Q', 'Y');
                                   // Table disconnect
                                   qrySeq.Active := False;
                                end;

                                qry_cell.Active := False;
                             end;
                     end;

                     if qryI_Gulkwa.ParamByName('gubn_part').Asstring  <> 'N' then
                     begin
                        qryI_Gulkwa.ExecSQL;
                        GP_query_log(qryI_Gulkwa, 'LD2I01', 'Q', 'Y');
                     end;

                     {if ((qryHangmok.FieldByName('cod_hm').AsString = 'S079') or
                         (qryHangmok.FieldByName('cod_hm').AsString = 'S090')) then
                     begin
                        qry_Allergy.Active := False;
                        qry_Allergy.ParamByName('num_id').AsString   := FieldByName('num_id').AsString;
                        qry_Allergy.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                        qry_Allergy.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                        qry_Allergy.ParamByName('cod_hm').AsString   := qryHangmok.FieldByName('cod_hm').AsString;
                        qry_Allergy.Active := True;
                        if qry_Allergy.RecordCount = 0 then
                        begin
                           qryI_Allergy.ParamByName('cod_bjjs').Asstring    := UV_sJisa;
                           qryI_Allergy.ParamByName('num_id').Asstring      := FieldByName('num_id').AsString;
                           qryI_Allergy.ParamByName('cod_jisa').Asstring    := FieldByName('cod_jisa').AsString;
                           qryI_Allergy.ParamByName('dat_gmgn').Asstring    := FieldByName('dat_gmgn').AsString;
                           qryI_Allergy.ParamByName('cod_hm').Asstring      := qryHangmok.FieldByName('cod_hm').AsString;
                           qryI_Allergy.ParamByName('dat_bunju').Asstring   := mskDate.Text;
                           qryI_Allergy.ParamByName('num_bunju').Asstring   := sBunCheck;
                           qryI_Allergy.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                           qryI_Allergy.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                           qryI_Allergy.ExecSQL;
                           GP_query_log(qryI_Allergy, 'LD2I01', 'Q', 'Y');
                        end; // end of if[qryI_Allergy];

                        qryI_Allergy.Active := False;
                     end;

                     if ((qryHangmok.FieldByName('cod_hm').AsString = 'S101') or
                         (qryHangmok.FieldByName('cod_hm').AsString = 'S102')) then
                     begin
                        qry_Allergy.Active := False;
                        qry_Allergy.ParamByName('num_id').AsString   := FieldByName('num_id').AsString;
                        qry_Allergy.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                        qry_Allergy.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                        qry_Allergy.ParamByName('cod_hm').AsString   := qryHangmok.FieldByName('cod_hm').AsString;
                        qry_Allergy.Active := True;
                        if qry_Allergy.RecordCount = 0 then
                        begin
                           qryI_Allergy62.ParamByName('cod_bjjs').Asstring    := UV_sJisa;
                           qryI_Allergy62.ParamByName('num_id').Asstring      := FieldByName('num_id').AsString;
                           qryI_Allergy62.ParamByName('cod_jisa').Asstring    := FieldByName('cod_jisa').AsString;
                           qryI_Allergy62.ParamByName('dat_gmgn').Asstring    := FieldByName('dat_gmgn').AsString;
                           qryI_Allergy62.ParamByName('cod_hm').Asstring      := qryHangmok.FieldByName('cod_hm').AsString;
                           qryI_Allergy62.ParamByName('dat_bunju').Asstring   := mskDate.Text;
                           qryI_Allergy62.ParamByName('num_bunju').Asstring   := sBunCheck;
                           qryI_Allergy62.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                           qryI_Allergy62.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                           qryI_Allergy62.ExecSQL;
                           GP_query_log(qryI_Allergy62, 'LD2I01', 'Q', 'Y');
                        end; // end of if[qryI_Allergy];

                        qryI_Allergy62.Active := False;
                     end;
                     }
                     //2018.05.28 ..����
                     if (qryHangmok.FieldByName('cod_hm').AsString = 'S104')  and (mskDate.Text >= '20180602') then  
                     begin
                        qry_Allergy.Active := False;
                        qry_Allergy.ParamByName('num_id').AsString   := FieldByName('num_id').AsString;
                        qry_Allergy.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                        qry_Allergy.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                        qry_Allergy.ParamByName('cod_hm').AsString   := qryHangmok.FieldByName('cod_hm').AsString;
                        qry_Allergy.Active := True;
                        if qry_Allergy.RecordCount = 0 then
                        begin
                           qryI_Allergy107.ParamByName('cod_bjjs').Asstring    := UV_sJisa;
                           qryI_Allergy107.ParamByName('num_id').Asstring      := FieldByName('num_id').AsString;
                           qryI_Allergy107.ParamByName('cod_jisa').Asstring    := FieldByName('cod_jisa').AsString;
                           qryI_Allergy107.ParamByName('dat_gmgn').Asstring    := FieldByName('dat_gmgn').AsString;
                           qryI_Allergy107.ParamByName('cod_hm').Asstring      := qryHangmok.FieldByName('cod_hm').AsString;
                           qryI_Allergy107.ParamByName('dat_bunju').Asstring   := mskDate.Text;
                           qryI_Allergy107.ParamByName('num_bunju').Asstring   := sBunCheck;
                           qryI_Allergy107.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                           qryI_Allergy107.ParamByName('COD_INSERT').AsString  := GV_sUserId;

                           qryI_Allergy107.ExecSQL;
                           GP_query_log(qryI_Allergy107, 'LD2I01', 'Q', 'Y');
                        end; // end of if[qryI_Allergy];

                        qryI_Allergy107.Active := False;
                     end;
                  end;
                  qryTot_Hangmok_List.Next;
               end;
            end;
            //==================================================================

            //���ϰ��� ������ ���ֿ��� Check.
            qryU_Gumgin.ParamByName('ysno_bunju').AsString := 'Y';
            qryU_Gumgin.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
            qryU_Gumgin.ParamByName('num_id').AsString := FieldByName('num_id').AsString;
            qryU_Gumgin.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
            qryU_Gumgin.ExecSQL;
            GP_query_log(qryU_Gumgin, 'LD2I01', 'Q', 'Y');

            //Hgulkwa_chk ���װ���Ϸ� üũ ���̺� insert
            qryHgulkwa_chk.Active := false;
            qryHgulkwa_chk.ParamByName('cod_bjjs').AsString := UV_sJisa;
            qryHgulkwa_chk.ParamByName('dat_bunju').AsString := mskDate.Text;
            qryHgulkwa_chk.Active := true;
            if qryHgulkwa_chk.Recordcount <= 0 then
            begin
               qryIHgulkwa_chk.Active := false;
               qryIHgulkwa_chk.ParamByName('cod_bjjs').AsSTring := UV_sJisa;
               qryIHgulkwa_chk.ParamByName('dat_bunju').AsString := mskDate.text;
               qryIHgulkwa_chk.ExecSQL;
               GP_query_log(qryIHgulkwa_chk, 'LD2I01', 'Q', 'Y');
            end;

            //Hgulkwa_EtcChk [������üó��]���װ���Ϸ� üũ ���̺� insert
            qryHgulkwa_EtcChk.Active := false;
            qryHgulkwa_EtcChk.ParamByName('cod_bjjs').AsString  := UV_sJisa;
            qryHgulkwa_EtcChk.ParamByName('dat_bunju').AsString := mskDate.Text;
            qryHgulkwa_EtcChk.Active := true;
            if qryHgulkwa_EtcChk.Recordcount <= 0 then
            begin
               qryIHgulkwa_EtcChk.Active := false;
               qryIHgulkwa_EtcChk.ParamByName('cod_bjjs').AsSTring  := UV_sJisa;
               qryIHgulkwa_EtcChk.ParamByName('dat_bunju').AsString := mskDate.text;
               qryIHgulkwa_EtcChk.ExecSQL;
               GP_query_log(qryIHgulkwa_EtcChk, 'LD2I01', 'Q', 'Y');
            end;

            Next;
         end;
      end;
   except
      on E : EDBEngineError do
      begin
         dmComm.database.Rollback;
         //���� ���н�, �޼�������  ..20161206
         IF copy(GV_sUserId,1,2) = '15' then
         Begin
            with qryI_SMS do
            begin
               ParamByName('USER_ID').AsString   := GV_sUserId;
               ParamByName('SMS_MSG').AsString   := '�����۾� ����..T^T';
               ParamByName('CALLBACK').AsString  := '0237029000';
               ParamByName('DEST_INFO').AsString := '';

               GP_query_log(qryI_SMS, 'LD2I01', 'I', 'Y');
            end;
         end;
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end;

   // database commit
   dmComm.database.Commit;
   //���� �Ϸ��, �޼��� ����
   {IF copy(GV_sUserId,1,2) = '15' then
   begin
   qryBunju_Count.Active := False;
   qryBunju_Count.ParamByName('cod_bjjs').AsString := '15';
   qryBunju_Count.ParamByName('dat_bunju').AsString := mskDate.Text;
   qryBunju_Count.Active := True;
   with qryI_SMS do
      begin
         //ParamByName('USER_ID').AsString   := GV_sUserId;
         //ParamByName('SMS_MSG').AsString   := copy(mskDate.Text,1,4)+'�� ' + copy(mskDate.Text,5,2)+'�� '+ copy(mskDate.Text,7,2)+'�� ' +
                                              '�����۾� '+ qryBunju_Count.fieldByName('count').AsString + '�� �Ϸ�Ǿ����ϴ�.';
         ParamByName('CALLBACK').AsString  := '0237029000';
         //ParamByName('DEST_INFO').AsString := '�̻���^01030425204|�ڼ���^01076766504|������^01098710333|������^01093412315';
         //ParamByName('DEST_INFO').AsString := '�ڼ���^01076766504';
         //ExecSql;

         //GP_query_log(qryI_SMS, 'LD2I01', 'I', 'Y');
      end;
   end;
    }
   // Animate End
   Animate1.Active := False;
   Animate2.Active := False;

   Gubn_rack;
   GF_MsgBox('Information', '�۾��� ���������� ����Ǿ����ϴ�.');
end;

procedure TfrmLD2I01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // F2���� Check
   if Key <> GC_Refer then exit;

   // Hot-Key�� ����Ͽ� Button�� ���� ȿ���� ��
   if Sender = mskDate then UP_Click(btnDate);
end;
function  TfrmLD2I01.UF_BunjuNo(iBunjuNo, iBunjuNoe : Integer) : Integer;
begin
   Result := iBunjuNo;

   qryBunju.Filter := ' NUM_BUNJU >= ''' + IntToStr(iBunjuNo)+ ''' and NUM_BUNJU <= ''' + IntToStr(iBunjuNoe)+ '''';
   with qryBunju do
   begin
      while not Eof do
      begin
         if FieldByName('cod_jisa').AsString <> UV_sJisa then
         begin
            Result := StrToInt(FieldByName('num_bunju').AsString);
            exit;
         end
         else
            Next;
      end;
   end;

end;

function  TfrmLD2I01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2I01.SpeedButton1Click(Sender: TObject);
begin
   inherited;
   //�۾����� Ȯ��
   if mskDate.text <= GV_sToday then
   begin
      showmessage('�� �۾����ڸ� �ٲټ���....');
      mskDate.SetFocus;
      exit;
   end;

   //�������� Ȯ��
   if (SpinEdit3.Text + GF_LPad(SpinEdit4.Text,2,'0') +  GF_LPad(SpinEdit5.Text,2,'0')) <= GV_sToday then
      begin
      showmessage('�� �������ڸ� �ٲټ���.');
      SpinEdit5.SetFocus;
      exit;
   end;
   //
   if Label6.Caption = 'Ÿ�̸� ���۵� ���Դϴ�..'  then
   begin
      if not GF_MsgBox('Confirmation', 'Ÿ�̸Ӹ� ����Ͻðڽ��ϱ�??') then exit;
   end
   else
   begin
      if not GF_MsgBox('Confirmation', 'Ÿ�̸Ӹ� �����Ͻðڽ��ϱ�??') then exit;
   end;

   if Label6.Caption = 'Ÿ�̸� ���۵� ���Դϴ�..' then
   begin
      Label6.Caption := 'Ÿ�̸� �۵��� �����մϴ�.';
      if SpinEdit1.Value = 24 then
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Am ' + IntToStr(SpinEdit1.Value - 12) + ':' + SpinEdit2.Text
      else if SpinEdit1.Value >= 12 then
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Pm ' + IntToStr(SpinEdit1.Value - 12) + ':' + SpinEdit2.Text
      else
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Am ' + SpinEdit1.Text + ':' + SpinEdit2.Text;

      // �������ڸ� ����(Server System Date)
      with dmComm.qrySysDate do
      begin
         if not Active then Active := True;
         GV_sToday := FieldByName('SYSDATE').AsString;
         GV_sPrintToday := Copy(GV_sToday, 1, 4) + '/' +
                           Copy(GV_sToday, 5, 2) + '/' +
                           Copy(GV_sToday, 7, 2);

         // Table disconnect
         Active := False;
      end;

      Timer1.Interval := 1000;
      Timer2.Interval := 1000;
   end
   else
   begin
      Label6.Caption  := 'Ÿ�̸� ���۵� ���Դϴ�..';
      Label7.Caption  := '���� ���� ��������..';
      Label14.Caption := '���� �ð��� ǥ��˴ϴ�..';
      Timer1.Interval := 0;
      Timer2.Interval := 0;
   end;

end;

procedure TfrmLD2I01.Timer1Timer(Sender: TObject);
begin
   inherited;
   if Trim(formatdatetime('YYYY-MM-DD Am/Pm H:N', now)) = Trim(Label7.Caption) then
   begin
      Timer1.Interval := 0;
      Timer2.Interval := 0;
      Label14.Caption := '�۾������߽��ϴ�..!!';

      btnStartClick(self);
   end;
end;

procedure TfrmLD2I01.Timer2Timer(Sender: TObject);
begin
   inherited;
   Label14.Caption := formatdatetime('YYYY-MM-DD Am/Pm HH:NN:SS', now);
end;

procedure TfrmLD2I01.UP_Change(Sender: TObject);
begin
  inherited;
  if SpinEdit2.Text = '00' then SpinEdit2.Text := '0';
end;

procedure TfrmLD2I01.UP_num_rack(sGubn :String; sNum_bunju: integer; sDat_bunju : string);
var
  sRack_num,gubn_rack:String;
begin

     if (sGubn='A')  then
     begin
         if sNum_bunju < 4000 then
         begin
            case num_Arow of
                 1 : gubn_rack := 'A';
                 2 : gubn_rack := 'B';
                 3 : gubn_rack := 'C';
                 4 : gubn_rack := 'D';
                 5 : gubn_rack := 'E';
            end;
            sRack_num:='A:'+GF_LPad(inttostr(num_ARack), 2, '0')+'-'+gubn_rack+GF_LPad(inttostr(num_ACol), 2, '0');
            if (num_ACol mod 20 = 0) then
            begin
               num_ACol:=0;
               num_Arow:=num_Arow+1;
               if (num_Arow mod 6 = 0) then
               begin
                  num_Arow:=1;
                  num_ARack:=num_ARack+1;
               end;
            end;
            inc(num_ACol);
         end
         else if sNum_bunju >= 4000 Then
         begin
            case num_Brow of
                 1 : gubn_rack := 'A';
                 2 : gubn_rack := 'B';
                 3 : gubn_rack := 'C';
                 4 : gubn_rack := 'D';
                 5 : gubn_rack := 'E';
            end;
            sRack_num:='B:'+ GF_LPad(inttostr(num_BRack), 2, '0')+'-'+gubn_rack+GF_LPad(inttostr(num_BCol), 2, '0');
            if (num_BCol mod 20 = 0) then
            begin
               num_BCol:=0;
               num_Brow:=num_Brow+1;
               if (num_Brow mod 6 = 0) then
               begin
                  num_Brow:=1;
                  num_BRack:=num_BRack+1;
               end;
            end;
            inc(num_BCol);
         end;
     end
     else if (sGubn='C') then
     begin
         case num_Crow of
              1 : gubn_rack := 'A';
              2 : gubn_rack := 'B';
              3 : gubn_rack := 'C';
              4 : gubn_rack := 'D';
              5 : gubn_rack := 'E';
         end;
         sRack_num:='C:'+GF_LPad(inttostr(num_CRack), 2, '0')+'-'+gubn_rack+GF_LPad(inttostr(num_CCol), 2, '0');
         if (num_CCol mod 20 = 0) then
             begin
             num_CCol:=0;
             num_Crow:=num_Crow+1;
             if (num_Crow mod 6 = 0) then
                 begin
                 num_Crow:=1;
                 num_CRack:=num_CRack+1;
                 end;
             end;
         inc(num_CCol);
     end//;
     else // 20140416~ D��
     begin
         case num_Drow of
              1 : gubn_rack := 'A';
              2 : gubn_rack := 'B';
              3 : gubn_rack := 'C';
              4 : gubn_rack := 'D';
              5 : gubn_rack := 'E';
         end;
         sRack_num:='D:'+GF_LPad(inttostr(num_DRack), 2, '0')+'-'+gubn_rack+GF_LPad(inttostr(num_DCol), 2, '0');
         if (num_DCol mod 20 = 0) then
             begin
             num_DCol:=0;
             num_Drow:=num_Drow+1;
             if (num_Drow mod 6 = 0) then
                 begin
                 num_Drow:=1;
                 num_DRack:=num_DRack+1;
                 end;
             end;
         inc(num_DCol);
     end;// ~20140416
     Qry_uRackNo.ParamByName('Desc_Rackno').AsString := sRack_num;
     Qry_uRackNo.ParamByName('cod_bjjs').AsString    := UV_sjisa;
     Qry_uRackNo.ParamByName('num_bunju').AsString   := GF_LPad(inttoStr(sNum_bunju), 4, '0');
     Qry_uRackNo.ParamByName('dat_bunju').AsString   := sDat_bunju;
     Qry_uRackNo.ExecSQL;
     GP_query_log(Qry_uRackNo, 'LD2I01', 'U', 'Y');
end;

procedure TfrmLD2I01.Gubn_rack;
var
   i : Integer;
   vCheck_01,vCheck_08, vCheck_01_01, vCheck_08_01, vCheck_01_02, vCheck_China_01, vCheck_China_02, vCheck_C032 : Boolean;  //20140416
   sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
begin
   inherited;
      i := 1;
      qry_Bunju_Rack.Active := false;
      qry_Bunju_Rack.ParamByName('cod_bjjs').asString  := UV_sJisa;
      qry_Bunju_Rack.ParamByName('dat_bunju').asString := mskDate.text;
      qry_Bunju_Rack.Active := true;

      if qry_Bunju_Rack.RecordCount >0 then
      begin
         Gauge1.Progress  := 0;
         Gauge1.MaxValue  := qry_Bunju_Rack.RecordCount;
         while  not qry_Bunju_Rack.Eof do
         begin
              sEtc_Hangmok_hm := '';
              sTk_Hangmok_Pf  := '';
              sTk_Hangmok_hm  := '';
              sSelect := ''; sWhere1 := '';  sWhere2 := ''; sOderby := '';
              vCheck_01   := False;  vCheck_08 := False;  vCheck_China_01 := False;//20140416  vCheck_China_01 := False;
              vCheck_01_01 := False; vCheck_08_01 := False; vCheck_01_02 := False; vCheck_China_02 := False; //20140424  vCheck_China_02 := False;
              vCheck_C032 := False;
              Gauge1.Progress  := Gauge1.Progress+1;
              application.ProcessMessages;
              //------------------------------------------------------------------------
              //�˻��׸� ALL Display...
              //------------------------------------------------------------------------
              //1.Ư�������׸�üũ...
              if (qry_Bunju_Rack.FieldByName('gubn_tkgm').AsString <> '') and (qry_Bunju_Rack.FieldByName('gubn_tkgm').AsString <> '0')then
              begin
                   sTk_Hangmok_Pf := qry_Bunju_Rack.FieldByName('cod_prf').AsString;
                   sTk_Hangmok_hm := qry_Bunju_Rack.FieldByName('Tcod_chuga').AsString;
              end;

              //2.�������Display
              if qry_Bunju_Rack.FieldByName('gubn_nosin').AsString  = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qry_Bunju_Rack.FieldByName('gubn_nsyh').AsInteger);
              //3.���κ�����Display
              if qry_Bunju_Rack.FieldByName('gubn_adult').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qry_Bunju_Rack.FieldByName('gubn_adyh').AsInteger);
              //4.��������Display
              if qry_Bunju_Rack.FieldByName('gubn_agab').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qry_Bunju_Rack.FieldByName('gubn_agyh').AsInteger);
              //5.������ȯ������Display
              if qry_Bunju_Rack.FieldByName('gubn_life').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qry_Bunju_Rack.FieldByName('gubn_lfyh').AsInteger);

              //�߰��˻��׸�
              sEtc_Hangmok_hm := sEtc_Hangmok_hm + qry_Bunju_Rack.FieldByName('cod_chuga').AsString;

              //�������� ���Ǹ����...
              if Trim(sTk_Hangmok_Pf) <> '' then
              begin
                 sWhere1 := qry_Bunju_Rack.FieldByName('cod_jangbi').AsString + ''',''' + qry_Bunju_Rack.FieldByName('cod_hul').AsString + ''',''' + qry_Bunju_Rack.FieldByName('cod_cancer').AsString + ''',''';
                 For i := 1 to (length(sTk_Hangmok_Pf) div 4) do
                 begin
                    if i = (length(sTk_Hangmok_Pf) div 4) then sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4)
                    else                                       sWhere1 := sWhere1 + copy(sTk_Hangmok_Pf, (i*4)-3, 4) + ''',''';
                 end;
              end
              else
              begin
                 sWhere1 := qry_Bunju_Rack.FieldByName('cod_jangbi').AsString + ''',''' + qry_Bunju_Rack.FieldByName('cod_hul').AsString + ''',''' + qry_Bunju_Rack.FieldByName('cod_cancer').AsString + ''',''';
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
                    else                                        sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
                 end;
              end;

              //20140416~
              if (qry_Bunju_Rack.FieldByName('cod_dc').AsString = '156940') or
                 (qry_Bunju_Rack.FieldByName('cod_dc').AsString = '156944') or
                 (qry_Bunju_Rack.FieldByName('cod_hul').AsString = 'CV05') or
                 (qry_Bunju_Rack.FieldByName('cod_jangbi').AsString = 'CV05') then vCheck_China_01 := True;
              if (qry_Bunju_Rack.FieldByName('cod_hul').AsString = '4610') then vCheck_China_02 := True;

              sTotal_HangmokList := ''; vCheck_01 := False; vCheck_01_02 := False;
              with qryJHangmokList do
              begin
                 sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + qry_Bunju_Rack.FieldByName('dat_gmgn').AsString+ '''';
                 sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
                 sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part ) UNION ( ';
                 sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
                 sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
                 sSelect := sSelect + '   AND dat_apply <= ''' +  qry_Bunju_Rack.FieldByName('dat_gmgn').AsString + ''' )';
                 sOderby := ' ORDER BY H.gubn_part Desc, P.cod_hm ';
                 qryJHangmokList.Active := False;
                 qryJHangmokList.SQL.Clear;
                 qryJHangmokList.SQL.Add(sSelect + sOderby);
                 qryJHangmokList.Active := True;

                 if qryJHangmokList.RecordCount > 0 then
                 begin
                    while not Eof do
                    begin
                       sTotal_HangmokList := sTotal_HangmokList + FieldByName('cod_hm').AsString;
                       //��ȭ��/���ܸ鿪 ����...

                       if (FieldByName('gubn_part').AsString  = '01') or
                          (FieldByName('gubn_part').AsString  = '05') or
                          (FieldByName('gubn_part').AsString  = '07') then
                       begin

                           //20150411 ������. ���,���κ� ���&C032 ������ B�� �ο� �ȵǵ��� [����-������]
                           if mskDate.Text < '20171016' then
                           begin
                               if (qry_Bunju_Rack.FieldByName('gubn_nosin').AsString = '2') or
                                  (qry_Bunju_Rack.FieldByName('gubn_adult').AsString = '2') or
                                  (qry_Bunju_Rack.FieldByName('gubn_life').AsString  = '2') then
                               begin
                                  if (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString <> '43') and
                                        (FieldByName('cod_hm').AsString <> 'S016') and
                                        (FieldByName('cod_hm').AsString <> 'S021') and
                                        (FieldByName('cod_hm').AsString <> 'E068') and
                                        (FieldByName('cod_hm').AsString <> 'T006') and
                                        (FieldByName('cod_hm').AsString <> 'T007') and
                                        (FieldByName('cod_hm').AsString <> 'SE02') and
                                        (FieldByName('cod_hm').AsString <> 'TT02') and
                                        (FieldByName('cod_hm').AsString <> 'T009') and
                                        (FieldByName('cod_hm').AsString <> 'TT01') and
                                        (FieldByName('cod_hm').AsString <> 'C032') and
                                        (FieldByName('cod_hm').AsString <> 'C077') and
                                        (FieldByName('cod_hm').AsString <> 'C083') and
                                        (FieldByName('cod_hm').AsString <> 'P007') Then

                                        vCheck_C032 := True

                                  else if  (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString = '43') and
                                              (FieldByName('cod_hm').AsString <> 'C009') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C010') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C011') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C025') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C026') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C027') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C028') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C032') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C042') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C074') And  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'S016') and
                                              (FieldByName('cod_hm').AsString <> 'S021') and
                                              (FieldByName('cod_hm').AsString <> 'E068') and
                                              (FieldByName('cod_hm').AsString <> 'T006') and
                                              (FieldByName('cod_hm').AsString <> 'T007') and
                                              (FieldByName('cod_hm').AsString <> 'SE02') and
                                              (FieldByName('cod_hm').AsString <> 'TT02') and
                                              (FieldByName('cod_hm').AsString <> 'T009') and
                                              (FieldByName('cod_hm').AsString <> 'TT01') and
                                              (FieldByName('cod_hm').AsString <> 'C077') and
                                              (FieldByName('cod_hm').AsString <> 'C083') and
                                              (FieldByName('cod_hm').AsString <> 'P007') Then

                                              vCheck_C032 := True;
                               end;
                           end
                           else if mskDate.Text >= '20171016' then
                           begin
                               if (qry_Bunju_Rack.FieldByName('gubn_nosin').AsString = '2') or
                                  (qry_Bunju_Rack.FieldByName('gubn_adult').AsString = '2') or
                                  (qry_Bunju_Rack.FieldByName('gubn_life').AsString  = '2') then
                               begin
                                  if (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString = '15') and
                                       (FieldByName('cod_hm').AsString <> 'S016') and  
                                        (FieldByName('cod_hm').AsString <> 'S021') and
                                        (FieldByName('cod_hm').AsString <> 'E068') and
                                        (FieldByName('cod_hm').AsString <> 'T006') and
                                        (FieldByName('cod_hm').AsString <> 'T007') and
                                        (FieldByName('cod_hm').AsString <> 'SE02') and
                                        (FieldByName('cod_hm').AsString <> 'TT02') and
                                        (FieldByName('cod_hm').AsString <> 'T009') and
                                        (FieldByName('cod_hm').AsString <> 'TT01') and
                                        (FieldByName('cod_hm').AsString <> 'C032') and
                                        (FieldByName('cod_hm').AsString <> 'C077') and
                                        (FieldByName('cod_hm').AsString <> 'C083') and
                                        (FieldByName('cod_hm').AsString <> 'P007') Then

                                        vCheck_C032 := True

                                  else if  (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString <> '15') and
                                              (FieldByName('cod_hm').AsString <> 'C009') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C010') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C011') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C025') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C026') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C027') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C028') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C032') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C042') and  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'C074') And  /// �߰�
                                              (FieldByName('cod_hm').AsString <> 'S016') and
                                              (FieldByName('cod_hm').AsString <> 'S021') and
                                              (FieldByName('cod_hm').AsString <> 'E068') and
                                              (FieldByName('cod_hm').AsString <> 'T006') and
                                              (FieldByName('cod_hm').AsString <> 'T007') and
                                              (FieldByName('cod_hm').AsString <> 'SE02') and
                                              (FieldByName('cod_hm').AsString <> 'TT02') and
                                              (FieldByName('cod_hm').AsString <> 'T009') and
                                              (FieldByName('cod_hm').AsString <> 'TT01') and
                                              (FieldByName('cod_hm').AsString <> 'C077') and
                                              (FieldByName('cod_hm').AsString <> 'C083') and
                                              (FieldByName('cod_hm').AsString <> 'P007') Then

                                              vCheck_C032 := True;
                               end;
                           end;
                           
                           if mskDate.Text < '20140930' then
                           begin
                              if (FieldByName('cod_hm').AsString <> 'S016') and
                                 (FieldByName('cod_hm').AsString <> 'S021') and
                                 (FieldByName('cod_hm').AsString <> 'E068') and
                                 (FieldByName('cod_hm').AsString <> 'T006') and
                                 (FieldByName('cod_hm').AsString <> 'T007') and
                                 (FieldByName('cod_hm').AsString <> 'SE02') and
                                 (FieldByName('cod_hm').AsString <> 'TT02') and
                                 (FieldByName('cod_hm').AsString <> 'T009') Then vCheck_01 := True;
                           end
                           else if mskDate.Text < '20171016' then  // 20140929 ������
                           begin
                              //-------------���� ��ü�˻� �и� �۾�
                              If (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString <> '43') or
                                 (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'SS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'SS') or
                                 (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'GS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'GS') Then
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                 End
                              Else
                              If (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString = '43') Then   //// �߰�
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'C009') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C010') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C011') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C025') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C026') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C027') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C028') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C032') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C042') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C074') And  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                       End
                           end
                           ////////////////////-------------2015.12.28 ��â��
                           else if mskDate.Text < '20180302' then  // 20171012 �ڼ���
                           begin
                              //-------------���� ��ü�˻� �۾�
                              //ä�� + ��� -> ���� ��ü, ä�븸 -> ������ �˻�
                              If  (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString = '15') or
                                 (((Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1) <> '1') and (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1) <> '1') and
                                   (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1)  <> '1') and (Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1) <> '2') and
                                   (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1) <> '2') and (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1)  <> '2')) and
                                  ((Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'SS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'SS') or
                                   (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'GS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'GS'))) Then
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                 End
                              Else
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'C009') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C010') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C011') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C025') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C026') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C027') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C028') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C032') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C042') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C074') And  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                 End;
                           end
                           else if mskDate.Text >= '20180302' then  // 2018.02.21 �ڼ���
                           begin
                              //ä�� ���� ���� ��ȭ�� ��ü9���� ��ü�˻� ���� _ 2018.03.02���� ����
                              If  (qry_Bunju_Rack.FieldByName('Cod_Jisa').AsString = '15') then
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                 End
                              Else
                                 Begin
                                    if (FieldByName('cod_hm').AsString <> 'C009') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C010') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C011') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C025') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C026') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C027') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C028') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C032') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C042') and  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'C074') And  /// �߰�
                                       (FieldByName('cod_hm').AsString <> 'S016') and
                                       (FieldByName('cod_hm').AsString <> 'S021') and
                                       (FieldByName('cod_hm').AsString <> 'E068') and
                                       (FieldByName('cod_hm').AsString <> 'T006') and
                                       (FieldByName('cod_hm').AsString <> 'T007') and
                                       (FieldByName('cod_hm').AsString <> 'SE02') and
                                       (FieldByName('cod_hm').AsString <> 'TT02') and
                                       (FieldByName('cod_hm').AsString <> 'T009') and
                                       (FieldByName('cod_hm').AsString <> 'TT01') and  //TT01TT02C077C083 ����
                                       (FieldByName('cod_hm').AsString <> 'C077') and
                                       (FieldByName('cod_hm').AsString <> 'C083') and
                                       (FieldByName('cod_hm').AsString <> 'P007') Then vCheck_01 := True;   //��ȭ��,�鿪 ����Ʈ
                                 End;
                           end;
                           //=======================================================================================================

                           if  mskDate.Text < '20140310' then
                           begin
                              if (FieldByName('cod_hm').AsString <> 'C001') and
                                 (FieldByName('cod_hm').AsString <> 'C002') and
                                 (FieldByName('cod_hm').AsString <> 'C003') and
                                 (FieldByName('cod_hm').AsString <> 'C004') and
                                 (FieldByName('cod_hm').AsString <> 'C005') and
                                 (FieldByName('cod_hm').AsString <> 'C006') and
                                 (FieldByName('cod_hm').AsString <> 'C007') and
                                 (FieldByName('cod_hm').AsString <> 'C009') and
                                 (FieldByName('cod_hm').AsString <> 'C010') and
                                 (FieldByName('cod_hm').AsString <> 'C011') and
                                 (FieldByName('cod_hm').AsString <> 'C012') and
                                 (FieldByName('cod_hm').AsString <> 'C013') and
                                 (FieldByName('cod_hm').AsString <> 'C017') and
                                 (FieldByName('cod_hm').AsString <> 'C019') and
                                 (FieldByName('cod_hm').AsString <> 'C021') and
                                 (FieldByName('cod_hm').AsString <> 'C025') and
                                 (FieldByName('cod_hm').AsString <> 'C026') and
                                 (FieldByName('cod_hm').AsString <> 'C027') and
                                 (FieldByName('cod_hm').AsString <> 'C028') and
                                 (FieldByName('cod_hm').AsString <> 'C029') and
                                 (FieldByName('cod_hm').AsString <> 'C032') and
                                 (FieldByName('cod_hm').AsString <> 'C033') and
                                 (FieldByName('cod_hm').AsString <> 'C041') and
                                 (FieldByName('cod_hm').AsString <> 'C042') and
                                 (FieldByName('cod_hm').AsString <> 'C043') and
                                 (FieldByName('cod_hm').AsString <> 'C077') and
                                 (FieldByName('cod_hm').AsString <> 'C083') and
                                 (FieldByName('cod_hm').AsString <> 'T006') and
                                 (FieldByName('cod_hm').AsString <> 'P007') and
                                 (FieldByName('cod_hm').AsString <> 'T007') and
                                 (FieldByName('cod_hm').AsString <> 'S016') and
                                 (FieldByName('cod_hm').AsString <> 'S021') Then  vCheck_01_02 := True;
                           end
                           else if  (mskDate.Text >= '20140310') and (mskDate.Text < '20140930' ) then
                           begin
                              if (FieldByName('cod_hm').AsString <> 'C009') and
                                 (FieldByName('cod_hm').AsString <> 'C010') and
                                 (FieldByName('cod_hm').AsString <> 'C011') and
                                 (FieldByName('cod_hm').AsString <> 'C025') and
                                 (FieldByName('cod_hm').AsString <> 'C026') and
                                 (FieldByName('cod_hm').AsString <> 'C027') and
                                 (FieldByName('cod_hm').AsString <> 'C028') and
                                 (FieldByName('cod_hm').AsString <> 'C032') and
                                 (FieldByName('cod_hm').AsString <> 'C042') and
                                 (FieldByName('cod_hm').AsString <> 'S016') and
                                 (FieldByName('cod_hm').AsString <> 'TT01') and
                                 (FieldByName('cod_hm').AsString <> 'TT02') and
                                 (FieldByName('cod_hm').AsString <> 'C077') Then  vCheck_01_02 := True;
                           end
                           else if  mskDate.Text >= '20140930' then   //20140929 ������
                           begin
                              if (FieldByName('cod_hm').AsString <> 'C009') and
                                 (FieldByName('cod_hm').AsString <> 'C010') and
                                 (FieldByName('cod_hm').AsString <> 'C011') and
                                 (FieldByName('cod_hm').AsString <> 'C025') and
                                 (FieldByName('cod_hm').AsString <> 'C026') and
                                 (FieldByName('cod_hm').AsString <> 'C027') and
                                 (FieldByName('cod_hm').AsString <> 'C028') and
                                 (FieldByName('cod_hm').AsString <> 'C032') and
                                 (FieldByName('cod_hm').AsString <> 'C042') and
                                 (FieldByName('cod_hm').AsString <> 'S016') and
                                 (FieldByName('cod_hm').AsString <> 'TT01') and
                                 (FieldByName('cod_hm').AsString <> 'TT02') and
                                 (FieldByName('cod_hm').AsString <> 'C077') and   //C083 �߰�
                                 (FieldByName('cod_hm').AsString <> 'C083') and   //T006 �߰�
                                 (FieldByName('cod_hm').AsString <> 'T006') and   //T007 �߰�
                                 (FieldByName('cod_hm').AsString <> 'T007') Then  vCheck_01_02 := True;
                           end;
                       end;
                       if (FieldByName('gubn_part').AsString  = '08') then
                       begin
                           vCheck_08 := True;
                           if (FieldByName('cod_hm').AsString <> 'E015') and
                              (FieldByName('cod_hm').AsString <> 'U026') then vCheck_08_01 := True;
                       end;
                       //===============================================================
                       Next;
                    end;
                 end;
              end;

              if vCheck_China_01 = False then//20140416
              begin//20140416
                if  qry_Bunju_Rack.FieldByName('cod_jisa').AsString = '15' then
                begin
                    if (vCheck_01) then UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                end
                else
                begin
                   case qry_Bunju_Rack.FieldByName('gubn_hulgum').AsInteger of
                        1,2 : begin
                                  if  (vCheck_01) then
                                      UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                              end;
                          3 : begin
                                  if  mskDate.Text < '20140310' then
                                  begin
                                       if ((Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1)='1')  or
                                           (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1)='1')  or
                                           (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1) ='1')) then
                                       begin
                                          if (vCheck_01_02) and (vCheck_01) then
                                             UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end
                                       else
                                       if (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) <> 'SS') and (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) <> 'SS') and
                                          (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) <> 'GS') and (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) <> 'GS') then
                                       begin
                                           if (vCheck_01_02) and (vCheck_01) then
                                              UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end
                                       else
                                       begin
                                           if (vCheck_01) then
                                              UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end;
                                  end
                                  else
                                  begin
                                       if (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'SS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'SS') or
                                          (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) = 'GS') or (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) = 'GS') then
                                       begin
                                            if (vCheck_01) then
                                               UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end
                                       else
                                       if ((Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1)='1') or
                                           (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1)='1') or
                                           (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1) ='1')) then
                                       begin
                                            if (vCheck_01_02) and (vCheck_01) then
                                               UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end
                                       else
                                       if (Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1)='2') or
                                          (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1)='2') or
                                          (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1)='2')  then
                                       begin
                                            //20150409 ������. ���,���κ� ���&C032 ������ B�� �ο� �ȵǵ��� [����-������]
                                            if (vCheck_C032) then
                                               UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end
                                       else
                                       begin
                                            if vCheck_01 then
                                               UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                                       end;
                                  end;
                              end;
                   end;
                end;
                if (vCheck_08 = True) and (vCheck_01 = False) and (vCheck_08_01) then
                   UP_num_rack('C',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
              end  //20140416~
              else if (vCheck_China_01 = True) and (vCheck_China_02 = False) then
                      UP_num_rack('D',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString)//~20140416
              else if (vCheck_China_01 = True) and (vCheck_China_02 = True) then
                      UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);//~20140424
              qry_Bunju_Rack.next;
         end;
      end;
end;

end.
