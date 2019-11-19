//==============================================================================
// ���α׷� ���� : ���װ˻� ���ϰ��� ���ֹ�ȣ �߰� �ο�..(���縸..)
// �ý���        : ���հ���
// ��������      : 2009.05.18
// ������        : ����ȣ
// ��������      :
// �������      : 
//==============================================================================

unit LD2I04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges;

type
  TfrmLD2I04 = class(TfrmSingle)
    qryGumgin: TQuery;
    pnlMaster: TPanel;
    Panel2: TPanel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    btnStart: TBitBtn;
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
    qryProfileG: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryBjValue: TQuery;
    qryTkgum: TQuery;
    qryHcondition: TQuery;
    qry_cell: TQuery;
    qryHgulkwa_chk: TQuery;
    qryIHgulkwa_chk: TQuery;
    Panel4: TPanel;
    edtDc: TEdit;
    btnDc: TSpeedButton;
    edtDESC_DC: TEdit;
    qryDache: TQuery;
    procedure UP_Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
  private
    { Private declarations }
    function  UF_BunjuNo(iBunjuNo, iBunjuNoe : Integer) : Integer;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
  public
    // SQL �ӽú���
    UV_sBasicSQL : String;
    // �����ڵ�
    UV_sJisa : String;

    { Public declarations }
  end;

var
  frmLD2I04: TfrmLD2I04;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD2I04.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   // Button Click�� Event Procedure�� ������ �� ������ Sender�� ����
   if Sender = btnDate then GF_CalendarClick(mskDate)
   else if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end;
end;

procedure TfrmLD2I04.FormCreate(Sender: TObject);
begin
   inherited;

   // �������ڸ� ����
   mskDate.Text := GV_sToday;

   // SQL�� ����
   UV_sBasicSQL := qryGumgin.SQL.Text;
end;

procedure TfrmLD2I04.btnStartClick(Sender: TObject);
var i, iCod_chuga, bunju_no, Jbunju_no, iRet : Integer;
    sWhere, sOrderBy, sBunCheck, sJangCk, sCode : String;
    iBun10, iBun20, iBun30, iBun60, iBun70, iBun50, iBun90 : Integer;
    ijBun10, ijBun20, ijBun30, ijBun60, ijBun70, ijBun50, ijBun90 : Integer;
    vPart, vCod_chuga, vCod_hm: Variant;
    hul_chk : Boolean;

begin
   inherited;
   // ��ü(00)�ϰ�� ���縦 �������� �۾�
   if GV_sJisa = '00' then UV_sJisa := Copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   //���ֹ�ȣ���Ⱚ
{   if UV_sJisa = '15' then
   begin
   // �������簡 ������ ���...
      showmessage('������ �߰��� ���ֹ�ȣ�� �ο��� �� �����ϴ�.');
      exit;
   end
   else
   // �������簡 ������ ���...     }
   begin
      for i := 1 to 7 do
      begin
         with qryBjValue do
         begin
            Active := False;
            ParamByName('cod_bjjs').AsString  := UV_sJisa;
            ParamByName('dat_bunju').AsString := mskDate.Text;
            case i of
            1 : begin
                   ParamByName('snum').AsString      := '0';
                   ParamByName('enum').AsString      := '0999';
                end;
            2 : begin
                   ParamByName('snum').AsString      := '1000';
                   ParamByName('enum').AsString      := '3999';
                end;
            3 : begin
                   ParamByName('snum').AsString      := '4000';
                   ParamByName('enum').AsString      := '4999';
                end;
            4 : begin
                   ParamByName('snum').AsString      := '5000';
                   ParamByName('enum').AsString      := '5999';
                end;
            5 : begin
                   ParamByName('snum').AsString      := '6000';
                   ParamByName('enum').AsString      := '6999';
                end;
            6 : begin
                   ParamByName('snum').AsString      := '7000';
                   ParamByName('enum').AsString      := '8499';
                end;
            7 : begin
                   ParamByName('snum').AsString      := '8500';
                   ParamByName('enum').AsString      := '9999';
                end;
            end;
            Active := True;
            case i of
            1:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun10 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun10 := 1;
                end;
            2:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun20 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun20 := 1001;
                end;
            3:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun30 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun30 := 4001;
                end;
            4:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun50 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun50 := 5001;
                end;
            5:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun60 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun60 := 6001;
                end;
            6:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun70 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun70 := 7001;
                end;
            7:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     ijBun90 := FieldByName('bjvalue').AsInteger + 1
                  else
                     ijBun90 := 8501;
                end;
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

   // ��ü�ڵ尡 ���õǾ����� Check
   if edtDC.text = '' then
   begin
      GF_MsgBox('Warning', '��ü�ڵ�� �ݵ�� �Է��ؾ� �մϴ�.');
      edtDC.SetFocus;
      exit;
   end;

   if not GF_MsgBox('Confirmation', '���׺��ָ� �����Ͻðڽ��ϱ� ?') then exit;

   with qryHangmok do
   begin
      Active := False;
      ParamByName('dat_apply').Asstring := mskDate.Text;
      Active := True;
   end;

   // ������ֽ� �߰����ֹ�ȣ ���� check.
//   if UV_SJisa <> '15' then
   begin
      if mskDate.Text <> GV_sToday then
      begin
         GF_MsgBox('Warning', '���� ������ �߰����ָ� �����մϴ�.');
         exit;
      end
      else
      begin
         qryBunju.Active := False;
         qryBunju.ParamByName('cod_bjjs').AsString  := UV_sJisa;
         qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
         qryBunju.Active := True;
         if qryBunju.RecordCount = 0 then
         begin
            GF_MsgBox('Warning', '�߰����ָ� �����մϴ�.');
            qryBunju.Active := False;
            Exit;
         end;
      end;
   end;

   {// ���ֿ��� Check.
   qryBunju.Active := False;
   qryBunju.ParamByName('cod_bjjs').AsString := GV_sJisa;
   qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
   qryBunju.Active := True;
   if qryBunju.RecordCount < 0 then
   begin
      GF_MsgBox('Warning', '�����۾��� �� �����Դϴ�.');
      qryBunju.Active := False;
      Exit;
   end;}

   // DB �۾�
   dmComm.database.StartTransaction;
   try
      // Query ���� & �迭�� ����
      with qryGumgin do
      begin
         Active := False;

//         if UV_sJisa <> '15' then
         begin
            sWhere := ' WHERE cod_jisa = ''' + UV_sJisa + ''' AND ysno_bunju = ''N''';
            sWhere := sWhere + ' AND dat_gmgn = ''' + mskDate.Text + '''';      // ���ϸ� ��ȸ.
            sWhere := sWhere + ' AND gubn_neawon <> ''5''';
            if edtDC.text <> '' then
               sWhere := sWhere + ' AND cod_dc = ''' + edtDC.text + '''';
            sOrderBy := ' ORDER BY cod_jisa, cod_hul, num_samp';
         end;

//            sWhere := sWhere + ' AND num_id = ''1502047575''';

         SQL.Clear;
         SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
         Active := True;

         if qryGumgin.RecordCount > 0 then
         begin
            GP_query_log(qryGumgin, 'LD2I04', 'Q', 'N');
            // Animate Start
            Animate1.Active := True;
            Animate2.Active := True;
            Gauge.Progress  := 0;
            Gauge.MaxValue  := RecordCount;
            
            // Routine...
            while not EOF do
            begin
               Gauge.Progress := Gauge.Progress + 1;
            
               labStatus.Caption := FieldByName('desc_name').AsString + '- ' + FieldByName('num_id').AsString;
               Application.ProcessMessages;
            
               bunju_no  := 0;
               Jbunju_no  := 0;
               sBunCheck := '';
            
               // Part Setting
               vPart := VarArrayCreate([0, 0], varOleStr);
               vCod_hm := VarArrayCreate([0, 0], varOleStr);
            
               VarArrayRedim(vPart, 10);

               if UV_sJisa = '15' then
               begin
                  if (UV_sJisa <> FieldByName('cod_jisa').AsString) then
                  begin
                     Next;
                     Continue;
                  end;
               end
               else if UV_sJisa <> '15' then
               begin
                  if (UV_sJisa <> FieldByName('cod_jisa').AsString) or
                     (FieldByName('gubn_hulgum').AsString = '1') or
                     (FieldByName('gubn_hulgum').AsString = '3') then
                  begin
                     Next;
                     Continue;
                  end;
               end;
            
              { if Trim(FieldByName('cod_jangbi').AsString) <> '' then
               begin
                  qryProfile.Active := False;
                  qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_jangbi').AsString;
                  qryProfile.Active := True;
                  sJangck := 'F';
                  while not qryProfile.Eof do
                  begin
                     qryHangmok.Filter := ' cod_hm = ''' + qryProfile.FieldByName('cod_hm').AsString + '''';
                     if qryhangmok.RecordCount > 0 then
                     begin
                        if (qryHangmok.FieldByName('gubn_part').AsString > '00') and (qryHangmok.FieldByName('gubn_part').AsString < '10') then
                        begin
                           sJangCk := 'T';
                           Break;
                        end;
                     end;
                     qryProfile.Next;
                  end;
               end;   }
            
               //���ֹ�ȣ ����
               if Trim(FieldByName('cod_hul').AsString) <> '' then
               begin
                  qryProfile.Active := False;
                  qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_hul').AsString;
                  qryProfile.Active := True;
                  if qryProfile.RecordCount > 0 then
                  begin
//                     if UV_sJisa <> '15' then
                     // ���� -> ���ֹ�ȣ ����...

                     bunju_no := qryProfile.FieldByName('gubn_gmsa').AsInteger;

                     if Copy(FieldByName('cod_hul').AsString,1,2) = 'SS' then
                     begin
                        bunju_no := 90;
                     end;

                     begin
                        case bunju_no of
                        10 : begin
//                             ijBun10 := UF_BunjuNo(ijBun10, 1000);
                             sBunCheck := IntToStr(ijBun10);
                             ijBun10 := ijBun10 + 1;
                             end;
                        20 : begin
//                             ijBun20 := UF_BunjuNo(ijBun20, 4000);
                             sBunCheck := IntToStr(ijBun20);
                             ijBun20 := ijBun20 + 1;
                             end;
                        30 : begin
//                             ijBun30 := UF_BunjuNo(ijBun30, 5000);
                             sBunCheck := IntToStr(ijBun30);
                             ijBun30 := ijBun30 + 1;
                             end;
                        50 : begin
//                             ijBun50 := UF_BunjuNo(ijBun50, 6000);
                             sBunCheck := IntToStr(ijBun50);
                             ijBun50 := ijBun50 + 1;
                             end;
                        60 : begin
//                             ijBun60 := UF_BunjuNo(ijBun60, 7000);
                             sBunCheck := IntToStr(ijBun60);
                             ijBun60 := ijBun60 + 1;
                             end;
                        70 : begin
//                             ijBun70 := UF_BunjuNo(ijBun70, 8999);
                             sBunCheck := IntToStr(ijBun70);
                             ijBun70 := ijBun70 + 1;
                             end;
                        90 : begin
//                             ijBun90 := UF_BunjuNo(ijBun90, 9999);
                             sBunCheck := IntToStr(ijBun90);
                             ijBun90 := ijBun90 + 1;
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
                     GP_query_log(qryI_Bunju, 'LD2I04', 'Q', 'Y');
                 end;
               end
               //����ȣ ����
//               else if (Trim(FieldByName('cod_jangbi').AsString) <> '') and (sJangCk = 'T') then
               else if (Trim(FieldByName('cod_jangbi').AsString) <> '') then
               begin
                  qryProfile.Active := False;
                  qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_jangbi').AsString;
                  qryProfile.Active := True;
            
                //  if sJangck = 'T' then
                //  begin
                     if qryProfile.RecordCount > 0 then
                     begin
//                        if UV_sJisa <> '15' then
                        Jbunju_no := qryProfile.FieldByName('gubn_gmsa').AsInteger;
                        
                        if (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'SS') or
                           (Copy(FieldByName('cod_jangbi').AsString,1,2) = 'GS') then
                        begin
                           Jbunju_no := 90;
                        end;

                        begin
                           case Jbunju_no of
                           10 : begin
//                                ijBun10 := UF_BunjuNo(ijBun10, 1000);
                                sBunCheck := IntToStr(ijBun10);
                                ijBun10 := ijBun10 + 1;
                                end;
                           20 : begin
//                                ijBun20 := UF_BunjuNo(ijBun20, 4000);
                                sBunCheck := IntToStr(ijBun20);
                                ijBun20 := ijBun20 + 1;
                                end;
                           30 : begin
//                                ijBun30 := UF_BunjuNo(ijBun30, 5000);
                                sBunCheck := IntToStr(ijBun30);
                                ijBun30 := ijBun30 + 1;
                                end;
                           50 : begin
//                                ijBun50 := UF_BunjuNo(ijBun50, 6000);
                                sBunCheck := IntToStr(ijBun50);
                                ijBun50 := ijBun50 + 1;
                                end;
                           60 : begin
//                                ijBun60 := UF_BunjuNo(ijBun60, 7000);
                                sBunCheck := IntToStr(ijBun60);
                                ijBun60 := ijBun60 + 1;
                                end;
                           70 : begin
//                                ijBun70 := UF_BunjuNo(ijBun70, 8999);
                                sBunCheck := IntToStr(ijBun70);
                                ijBun70 := ijBun70 + 1;
                                end;
                           90 : begin
//                                ijBun90 := UF_BunjuNo(ijBun90, 9999);
                                sBunCheck := IntToStr(ijBun90);
                                ijBun90 := ijBun90 + 1;
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
                        GP_query_log(qryI_Bunju, 'LD2I04', 'Q', 'Y');
            
                     end;
                  //end;
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
//                        if UV_sJisa <> '15' then
                        begin
                           case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                           60 : begin
//                                ijBun60 := UF_BunjuNo(ijBun60, 7000);
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
                        GP_query_log(qryI_Bunju, 'LD2I04', 'Q', 'Y');
                     end;
                  end;
               end
               else if (FieldByName('gubn_nosin').AsString >= '1')
                    or (FieldByName('gubn_tkgm').AsString >= '1')
                    or (FieldByName('gubn_bogun').AsString >= '1')
                    or (FieldByName('gubn_adult').AsString >= '1')
                    or (FieldByName('gubn_agab').AsString >= '1')
                    or (FieldByName('gubn_life').AsString >= '1') then
               begin
                  if sBunCheck = '' then
                  begin
//                     if UV_sJisa <> '15' then
                     begin
//                        ijBun70 := UF_BunjuNo(ijBun70, 9999);
                        sBunCheck := IntToStr(ijBun70);
                        ijBun70 := ijBun70 + 1;
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
                     GP_query_log(qryI_Bunju, 'LD2I04', 'Q', 'Y');
                  end;
               end
               //�߰��ڵ� Ȯ��
               else if (Trim(FieldByName('cod_chuga').AsString) <> '')
                    and ((FieldByName('gubn_nosin').AsString < '1')
                    and (FieldByName('gubn_tkgm').AsString <'1')
                    and (FieldByName('gubn_bogun').AsString < '1')
                    and (FieldByName('gubn_adult').AsString < '1')
                    and (FieldByName('gubn_agab').AsString <'1')
                    and (FieldByName('gubn_life').AsString < '1')) then
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
                       if qryHangmok.FieldByName('GUBN_PART').AsString < '10' then
                       begin
                            hul_chk := true;
                            break;
                       end
                       else
                            hul_chk := false;
                  end;
            
                  if hul_chk = true then
                  begin
                       if sBunCheck = '' then
                       begin
//                            if UV_sJisa <> '15' then
                            begin
//                               ijBun70 := UF_BunjuNo(ijBun70, 9999);
                                 sBunCheck := IntToStr(ijBun70);
                                 ijBun70 := ijBun70 + 1;
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
                            GP_query_log(qryI_Bunju, 'LD2I04', 'Q', 'Y');
                       end;
                  end;
               end;
            
            
               qryProfileG.Active := False;
               qryProfileG.ParamByName('cod_pf1').AsString  := FieldByName('cod_hul').AsString;
               qryProfileG.ParamByName('cod_pf2').AsString  := FieldByName('cod_jangbi').AsString;
               qryProfileG.ParamByName('cod_pf3').AsString  := FieldByName('cod_cancer').AsString;
               qryProfileG.ParamByName('cod_pf4').AsString  := '';
               qryProfileG.ParamByName('cod_pf5').AsString  := '';
               qryProfileG.ParamByName('cod_pf6').AsString  := '';
               qryProfileG.ParamByName('cod_pf7').AsString  := '';
               qryProfileG.ParamByName('cod_pf8').AsString  := '';
               qryProfileG.ParamByName('cod_pf9').AsString  := '';
               qryProfileG.ParamByName('cod_pf10').AsString := '';
               qryProfileG.ParamByName('cod_pf11').AsString := '';
               qryProfileG.ParamByName('cod_pf12').AsString := '';
               qryProfileG.ParamByName('cod_pf13').AsString := '';
               qryProfileG.ParamByName('cod_pf14').AsString := '';
               qryProfileG.ParamByName('cod_pf15').AsString := '';
               qryProfileG.ParamByName('cod_pf16').AsString := '';
               qryProfileG.ParamByName('cod_pf17').AsString := '';
               qryProfileG.ParamByName('cod_pf18').AsString := '';
               qryProfileG.ParamByName('cod_pf19').AsString := '';
               qryProfileG.ParamByName('cod_pf20').AsString := '';
               qryProfileG.ParamByName('cod_pf21').AsString := '';
               qryProfileG.ParamByName('cod_pf22').AsString := '';
               qryProfileG.ParamByName('cod_pf23').AsString := '';
               qryProfileG.ParamByName('cod_pf24').AsString := '';
               qryProfileG.Active := True;
               if qryProfileG.RecordCount > 0 then
               begin
                  while not qryProfileG.Eof do
                  begin
                     sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
            
                     qryHangmok.Filter := ' COD_HM = ''' + qryProfileG.FieldByName('cod_hm').AsString + '''' +
                                          ' AND DAT_APPLY <= ''' + FieldByName('DAT_GMGN').AsString + '''';
            
                     if qryHangmok.RecordCount > 0 then
                     begin
                        qryI_Gulkwa.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                        qryI_Gulkwa.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                        qryI_Gulkwa.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                        qryI_Gulkwa.ParamByName('num_bunju').Asstring  := sBunCheck;
                        qryI_Gulkwa.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                        qryI_Gulkwa.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                        qryI_Gulkwa.ParamByName('gubn_part').Asstring  := 'N';
                        qryI_Gulkwa.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                        qryI_Gulkwa.ParamByName('COD_INSERT').AsString  := GV_sUserId;
            
                        Case qryHangmok.FieldByName('gubn_part').AsInteger of
                        01 : begin
                                if Trim(vPart[1]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '01';
                                   vPart[1] := 'Y';
                                end;
                             end;
                        02 : begin
                                if Trim(vPart[2]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '02';
                                   vPart[2] := 'Y';
                                end;
                             end;
                        03 : begin
                                if Trim(vPart[3]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '03';
                                   vPart[3] := 'Y';
                                end;
                             end;
                        04 : begin
                                if Trim(vPart[4]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '04';
                                   vPart[4] := 'Y';
                                end;
                             end;
                        05 : begin
                                if Trim(vPart[5]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '05';
                                   vPart[5] := 'Y';
                                end;
                             end;
                        08 : begin
                                if Trim(vPart[8]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '08';
                                   vPart[8] := 'Y';
                                end;
                             end;
                        07 : begin
                                if Trim(vPart[7]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '07';
                                   vPart[7] := 'Y';
                                end;
                             end;
                        09 : begin
                                if Trim(vPart[9]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '09';
                                   vPart[9] := 'Y';
                                end;
                             end;
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
                                   qryI_Cell.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                                   qryI_Cell.ParamByName('COD_INSERT').AsString  := GV_sUserId;
            
                                   qrySeq.Active := False;
                                   if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                   begin
                                      qrySeq.ParamByName('COD_CELLS').AsString := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00000';
                                      qrySeq.ParamByName('COD_CELLE').AsString := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '99999';
                                   end
                                   else
                                   begin
                                      qrySeq.ParamByName('COD_CELLS').AsString := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00000';
                                      qrySeq.ParamByName('COD_CELLE').AsString := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '99999';
                                   end;
                                   qrySeq.Active := True;
            
                                   if Trim(qrySeq.FieldByName('RESULT').AsString) = '' then
                                   begin
                                      if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                         qryI_Cell.ParamByName('cod_cell').Asstring := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00001'
                                      else
                                         qryI_Cell.ParamByName('cod_cell').Asstring := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00001';
                                   end
                                   else
                                      qryI_Cell.ParamByName('cod_cell').Asstring := Copy(qrySeq.FieldByName('RESULT').AsString, 1, 6) +  GF_LPad(IntToStr(StrToInt(Copy(qrySeq.FieldByName('RESULT').AsString, 7, 5)) + 1), 5, '0');
            
                                   qryI_Cell.ExecSQL;
                                   GP_query_log(qryI_Cell, 'LD2I04', 'Q', 'Y');
                                   // Table disconnect
                                   qrySeq.Active := False;
                                end;
                                qry_cell.Active := False;
                             end;
                        end;
                        if qryI_Gulkwa.ParamByName('gubn_part').Asstring  <> 'N' then
                        begin
                           qryI_Gulkwa.ExecSQL;
                           GP_query_log(qryI_Gulkwa, 'LD2I04', 'Q', 'Y');
                        end;
                     end;
                     qryProfileG.Next;
                  end;
               end;
            
               sCode := FieldByName('COD_CHUGA').AsString;
               if FieldByName('gubn_nosin').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '1', FieldByName('gubn_nsyh').AsInteger);
               // ���_��� : ������̺��� ������.
               {else if FieldByName('gubn_nosin').AsString = '2' then
               begin
                  sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '1', FieldByName('gubn_nsyh').AsInteger)
               end;  }
            
               if (FieldByName('gubn_tkgm').AsString = '1') or (FieldByName('gubn_tkgm').AsString = '2') then
               begin
                  qryTkgum.Active := False;
                  qryTkgum.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                  qryTkgum.ParamByName('num_jumin').AsString := FieldByName('num_jumin').AsString;
                  qryTkgum.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                  qryTkgum.Active := True;
                  if qryTkgum.RecordCount > 0 then
                  begin
                     qryProfileG.Active := False;
                        qryProfileG.ParamByName('cod_pf1').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  1, 4);
                        qryProfileG.ParamByName('cod_pf2').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  5, 4);
                        qryProfileG.ParamByName('cod_pf3').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  9, 4);
                        qryProfileG.ParamByName('cod_pf4').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 13, 4);
                        qryProfileG.ParamByName('cod_pf5').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 17, 4);
                        qryProfileG.ParamByName('cod_pf6').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 21, 4);
                        qryProfileG.ParamByName('cod_pf7').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 25, 4);
                        qryProfileG.ParamByName('cod_pf8').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 29, 4);
                        qryProfileG.ParamByName('cod_pf9').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 33, 4);
                        qryProfileG.ParamByName('cod_pf10').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 37, 4);
                        qryProfileG.ParamByName('cod_pf11').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 41, 4);
                        qryProfileG.ParamByName('cod_pf12').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 45, 4);
                        qryProfileG.ParamByName('cod_pf13').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 49, 4);
                        qryProfileG.ParamByName('cod_pf14').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 53, 4);
                        qryProfileG.ParamByName('cod_pf15').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 57, 4);
                        qryProfileG.ParamByName('cod_pf16').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 61, 4);
                        qryProfileG.ParamByName('cod_pf17').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 65, 4);
                        qryProfileG.ParamByName('cod_pf18').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 69, 4);
                        qryProfileG.ParamByName('cod_pf19').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 73, 4);
                        qryProfileG.ParamByName('cod_pf20').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 77, 4);
                        qryProfileG.ParamByName('cod_pf21').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 81, 4);
                        qryProfileG.ParamByName('cod_pf22').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 85, 4);
                        qryProfileG.ParamByName('cod_pf23').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 89, 4);
                        qryProfileG.ParamByName('cod_pf24').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 93, 4);
                     qryProfileG.Active := True;
                     if qryProfileG.RecordCount > 0 then
                     begin
                        while not qryProfileG.Eof do
                        begin
                           sCode := sCode + qryProfileG.FieldByName('cod_hm').AsString;
                           qryProfileG.Next;
                        end;
                     end;
                     qryProfileG.Active := False;
                     sCode := sCode + qryTkgum.FieldByName('cod_chuga').AsString;
                  end;
                  qryTkgum.Active := False;
               end;
            
               if FieldByName('gubn_bogun').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '3', FieldByName('gubn_bgyh').AsInteger);
{               else if FieldByName('gubn_bogun').AsString = '2' then
               begin
                  sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '3', FieldByName('gubn_nsyh').AsInteger)
               end;
}           
               if FieldByName('gubn_adult').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '4', FieldByName('gubn_adyh').AsInteger);
{               else if FieldByName('gubn_adult').AsString = '2' then
               begin
                  sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_nsyh').AsInteger)
               end;
}           
               if FieldByName('gubn_agab').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '5', FieldByName('gubn_agyh').AsInteger);
{               else if FieldByName('gubn_agab').AsString = '2' then
               begin
                  sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_nsyh').AsInteger)
               end;
}           
               if FieldByName('gubn_gong').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '6', FieldByName('gubn_goyh').AsInteger);
{               else if FieldByName('gubn_gong').AsString = '2' then
               begin
                  sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '6', FieldByName('gubn_nsyh').AsInteger)
               end;
}           
               if FieldByName('gubn_life').AsString = '1' then
                  sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '7', FieldByName('gubn_lfyh').AsInteger);
            
               if Trim(sCode) <> '' then
               begin
                  iCod_chuga := GF_MulToSingle(sCode, 4, vCod_chuga);
                  for i := 0 to iCod_chuga - 1 do
                  begin
                     sBuncheck := GF_LPad(Trim(sBuncheck), 4, '0');
                     qryHangmok.Filter := ' COD_HM = ''' + vCod_chuga[i] + '''' +
                                          ' AND DAT_APPLY <= ''' + FieldByName('DAT_GMGN').AsString + '''';
            
                     if qryHangmok.RecordCount > 0 then
                     begin
                        qryI_Gulkwa.ParamByName('cod_bjjs').Asstring   := UV_sJisa;
                        qryI_Gulkwa.ParamByName('num_id').Asstring     := FieldByName('num_id').AsString;
                        qryI_Gulkwa.ParamByName('dat_bunju').Asstring  := mskDate.Text;
                        qryI_Gulkwa.ParamByName('num_bunju').Asstring  := sBunCheck;
                        qryI_Gulkwa.ParamByName('dat_gmgn').Asstring   := FieldByName('dat_gmgn').AsString;
                        qryI_Gulkwa.ParamByName('cod_jisa').Asstring   := FieldByName('cod_jisa').AsString;
                        qryI_Gulkwa.ParamByName('gubn_part').Asstring  := 'N';
                        qryI_Gulkwa.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                        qryI_Gulkwa.ParamByName('COD_INSERT').AsString  := GV_sUserId;
            
                        Case qryHangmok.FieldByName('gubn_part').AsInteger of
                        01 : begin
                                if Trim(vPart[1]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '01';
                                   vPart[1] := 'Y';
                                end;
                             end;
                        02 : begin
                                if Trim(vPart[2]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '02';
                                   vPart[2] := 'Y';
                                end;
                             end;
                        03 : begin
                                if Trim(vPart[3]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '03';
                                   vPart[3] := 'Y';
                                end;
                             end;
                        04 : begin
                                if Trim(vPart[4]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '04';
                                   vPart[4] := 'Y';
                                end;
                             end;
                        05 : begin
                                if Trim(vPart[5]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '05';
                                   vPart[5] := 'Y';
                                end;
                             end;
                        08 : begin
                                if Trim(vPart[8]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '08';
                                   vPart[8] := 'Y';
                                end;
                             end;
                        07 : begin
                                if Trim(vPart[7]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '07';
                                   vPart[7] := 'Y';
                                end;
                             end;
                        09 : begin
                                if Trim(vPart[9]) = '' then
                                begin
                                   qryI_Gulkwa.ParamByName('gubn_part').Asstring  := '09';
                                   vPart[9] := 'Y';
                                end;
                             end;
                        06 : begin
                                qry_cell.Active := False;
                                qry_cell.ParamByName('num_id').AsString   := FieldByName('num_id').AsString;
                                qry_cell.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                                qry_cell.ParamByName('cod_hm').AsString   := qryHangmok.FieldByName('cod_hm').AsString;
                                qry_cell.Active := True;
                                if qry_cell.RecordCount = 0 then
                                begin
                                    qryI_Cell.ParamByName('cod_bjjs').Asstring    := UV_sJisa;
                                    qryI_Cell.ParamByName('num_id').Asstring      := FieldByName('num_id').AsString;
                                    qryI_Cell.ParamByName('dat_gmgn').Asstring    := FieldByName('dat_gmgn').AsString;
                                    qryI_Cell.ParamByName('cod_jisa').Asstring    := FieldByName('cod_jisa').AsString;
                                    qryI_Cell.ParamByName('cod_hm').Asstring      := qryHangmok.FieldByName('cod_hm').AsString;
                                    qryI_Cell.ParamByName('ysno_sokun').Asstring  := 'N';
                                    qryI_Cell.ParamByName('DAT_INSERT').AsString  := GV_sToday;
                                    qryI_Cell.ParamByName('COD_INSERT').AsString  := GV_sUserId;
            
                                    qrySeq.Active := False;
                                    if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                    begin
                                       qrySeq.ParamByName('COD_CELLS').AsString := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00000';
                                       qrySeq.ParamByName('COD_CELLE').AsString := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '99999';
                                    end
                                    else
                                    begin
                                       qrySeq.ParamByName('COD_CELLS').AsString := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00000';
                                       qrySeq.ParamByName('COD_CELLE').AsString := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '99999';
                                    end;
                                    qrySeq.Active := True;
            
                                    if Trim(qrySeq.FieldByName('RESULT').AsString) = '' then
                                    begin
                                       if copy(qryHangmok.FieldByName('cod_hm').AsString, 1, 1) = 'B' then
                                          qryI_Cell.ParamByName('cod_cell').Asstring := 'J' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00001'
                                       else
                                          qryI_Cell.ParamByName('cod_cell').Asstring := 'P' + UV_sJisa + Copy(mskDate.Text, 4, 3) + '00001';
                                    end
                                    else
                                       qryI_Cell.ParamByName('cod_cell').Asstring := Copy(qrySeq.FieldByName('RESULT').AsString, 1, 6) +  GF_LPad(IntToStr(StrToInt(Copy(qrySeq.FieldByName('RESULT').AsString, 7, 5)) + 1), 5, '0');
                                    qryI_Cell.ExecSQL;
                                    GP_query_log(qryI_Cell, 'LD2I04', 'Q', 'Y');

                                    // Table disconnect
                                    qrySeq.Active := False;
                                end; // end of if[qry_cell];
                                qry_cell.Active := False;
                             end;
                        end;
                        if qryI_Gulkwa.ParamByName('gubn_part').Asstring  <> 'N' then
                        begin
                           qryI_Gulkwa.ExecSQL;
                           GP_query_log(qryI_Gulkwa, 'LD2I04', 'Q', 'Y');
                        end;
                     end;
                  end;
               end;
            
               //���ϰ��� ������ ���ֿ��� Check.
               qryU_Gumgin.ParamByName('ysno_bunju').AsString := 'Y';
               qryU_Gumgin.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
               qryU_Gumgin.ParamByName('num_id').AsString := FieldByName('num_id').AsString;
               qryU_Gumgin.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
               qryU_Gumgin.ExecSQL;
               GP_query_log(qryU_Gumgin, 'LD2I04', 'Q', 'Y');
            
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
                  GP_query_log(qryIHgulkwa_chk, 'LD2I04', 'Q', 'Y');
               end;
            
               Next;
            end;
         end
         else
         begin
            // �ڷᰡ ������ ǥ��
            GF_MsgBox('Information', 'NODATA');
            // database commit
            dmComm.database.Commit;
            exit;
         end;
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;

   // database commit
   dmComm.database.Commit;

   // Animate End
   Animate1.Active := False;
   Animate2.Active := False;
   Gauge.Progress  := 100;

   GF_MsgBox('Information', '�۾��� ���������� ����Ǿ����ϴ�.');
end;

procedure TfrmLD2I04.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // F2���� Check
   if Key <> GC_Refer then exit;

   // Hot-Key�� ����Ͽ� Button�� ���� ȿ���� ��
   if Sender = mskDate then UP_Click(btnDate)
   else if Sender = edtDc   then UP_Click(btnDC);
end;
function  TfrmLD2I04.UF_BunjuNo(iBunjuNo, iBunjuNoe : Integer) : Integer;
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

function  TfrmLD2I04.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD2I04.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_jumin').AsString   := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
         Result := FieldByName('desc_hul').AsString;
      Active := False;
   end;
end;

procedure TfrmLD2I04.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

  if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;
end;

end.
