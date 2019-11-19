//==============================================================================
// 프로그램 설명 : 혈액결과체크 insert 부분 추가
// 시스템        : 통합검진
// 수정일자      : 2004.10.25
// 수정자        : 최지혜
// 수정내용      :
// 참고사항      : 
//==============================================================================
// 수정일자      : 2006.02.04
// 수정자        : 김승철
// 수정내용      : 대체인경우 분주가 안돌아가게끔 수정 (gubn_neawon <> '5')
//==============================================================================
// 수정일자      : 2006.05.17
// 수정자        : 유동구
// 수정내용      : 지사자체 분주할때 순서대로 분주하기..
//==============================================================================
// 수정일자      : 2006.06.29
// 수정자        : 유동구
// 수정내용      : 장비'CV,CM' = 5000, 'SS,GS' = 9000 나머지는 검사구분으로 나뉘어짐..
//==============================================================================
// 수정일자      : 2007.10.26
// 수정자        : 김승철
// 수정내용      : cell 자릿수 부족으로 9999->99999 로 변경. (년도를 한자리로 변경함)
//==============================================================================
// 수정일자      : 2007.11.29
// 수정자        : 구수정
// 수정내용      : 혈액코드가 없는 추가코드만 있을때 분주번호 생성안되게...
//==============================================================================
// 수정일자      : 2008.10.30
// 수정자        : 김승철
// 수정내용      : 1000번때 중복되는 분주번호 발생으로 인하여 검사구분 10,20 은 같은 분주번호대 사용.
//                 지방센터의 경우 대체접수도 분주돌아가는거 막음.
//==============================================================================
// 수정일자      : 2011.12.21
// 수정자        : 유동구
// 수정내용      : [센터]자체검사 실시관련 추가...(qryHgulkwa_EtcChk)
//==============================================================================
// 수정일자      : 2013.5.21
// 수정자        : 김희석
// 수정내용      : Rack no  추가 
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
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    procedure Gubn_rack;
    procedure UP_num_rack(sGubn :String; sNum_bunju: integer; sDat_bunju : string);

  public
    // SQL 임시변수
    UV_sBasicSQL : String;
    // 지사코드
    UV_sJisa : String;
    num_ARack,num_ACol, num_Arow, num_BRack,  num_Brow, num_BCol,num_CRack,num_CCol, num_Crow :Integer;
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

   // Button Click시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD2I01.FormCreate(Sender: TObject);
begin
   inherited;

   // 현재일자를 설정
   mskDate.Text := GV_sToday;

   // SQL문 저장
   UV_sBasicSQL := qryGumgin.SQL.Text;

   SpinEdit3.Value := StrToInt(formatdatetime('YYYY', now));
   SpinEdit4.Value := StrToInt(formatdatetime('MM', now));
   SpinEdit5.Value := StrToInt(formatdatetime('DD', now));
   SpinEdit1.Value := StrToInt(formatdatetime('HH', now));
   SpinEdit2.Value := StrToInt(formatdatetime('NN', now));   
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
   //rack 초기화
   num_ARack:=1;
   num_ACol :=1;
   num_Arow :=1;
   num_BRack:=1;
   num_Brow :=1;
   num_BCol :=1;
   num_CRack:=1;
   num_Crow :=1;
   num_CCol :=1;
   //10.12.09.철. 예약분주시 dat_insert가 폼의 오늘일자를 구함(Server System Date)
   with dmComm.qrySysDate do
   begin
      if not Active then Active := True;
      GV_sToday := FieldByName('SYSDATE').AsString;
      // Table disconnect
      Active := False;
   end;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := Copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   //분주번호조기값
   if UV_sJisa = '15' then
   begin
   // 분주지사가 본사인 경우...
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
                   ParamByName('enum').AsString      := '3999';
                end;
    {        1 : begin
                   ParamByName('snum').AsString      := '0';
                   ParamByName('enum').AsString      := '999';
                end;
            2 : begin
                   ParamByName('snum').AsString      := '1000';
                   ParamByName('enum').AsString      := '3999';
                end;    }
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
                     iBun10 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun10 := 1;
                end;
        {    2:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun20 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun20 := 1001;
                end;      }
            3:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun30 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun30 := 4001;
                end;
            4:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun50 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun50 := 5001;
                end;
            5:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun60 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun60 := 6001;
                end;
            6:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun70 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun70 := 7001;
                end;
            7:  begin
                  if Trim(FieldByName('bjvalue').AsString) <> '' then
                     iBun90 := FieldByName('bjvalue').AsInteger + 1
                  else
                     iBun90 := 8501;
                end;
            end;
          end;
      end;
   end
   else
   // 분주지사가 지사인 경우...
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

   // 일자가 선택되었는지 Check
   if mskDate.Text  = '' then
   begin
      GF_MsgBox('Warning', '작업일자는 반드시 입력해야 합니다.');
      mskDate.SetFocus;
      exit;
   end;

   if Label6.Caption = '타이머 미작동 중입니다..'  then
   begin
      if not GF_MsgBox('Confirmation', '혈액분주를 시작하시겠습니까 ?') then exit;
   end;

   with qryHangmok do
   begin
      Active := False;
      ParamByName('dat_apply').Asstring := mskDate.Text;
      Active := True;
   end;

   // 지사분주시 본사의 분주여부 Check.
   if UV_SJisa <> '15' then
   begin
      qryBunju.Active := False;
      qryBunju.ParamByName('cod_bjjs').AsString := '15';
      qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
      qryBunju.Active := True;
      if qryBunju.RecordCount = 0 then
      begin
         GF_MsgBox('Warning', '본사 분주실에 문의하세요...');
         qryBunju.Active := False;
         Exit;
      end;
   end;

   // 분주여부 Check.
   qryBunju.Active := False;
   if GV_sJisa = '00' then
      qryBunju.ParamByName('cod_bjjs').AsString := Copy(GV_sUserId,1,2)
   else
      qryBunju.ParamByName('cod_bjjs').AsString := GV_sJisa;
   qryBunju.ParamByName('dat_bunju').AsString := mskDate.Text;
   qryBunju.Active := True;
   if qryBunju.RecordCount < 0 then
   begin
      GF_MsgBox('Warning', '분주작업이 된 일자입니다.');
      qryBunju.Active := False;
      Exit;
   end;

   
   // DB 작업
   dmComm.database.StartTransaction;
   try
      // Query 수행 & 배열에 저장
      with qryGumgin do
      begin
         Active := False;

         if UV_sJisa = '15' then
         begin
            sWhere := ' WHERE ysno_bunju = ''N''';
            sWhere := sWhere + ' AND dat_gmgn < ''' + mskDate.Text + '''';      //수정
            sWhere := sWhere + ' AND gubn_neawon <> ''5'' ';
            sOrderBy := ' ORDER BY cod_jisa, cod_hul, num_samp';
         end
         else
         begin
            sWhere := ' WHERE cod_jisa = ''' + UV_sJisa + ''' AND ysno_bunju = ''N''';
            sWhere := sWhere + ' AND dat_gmgn < ''' + mskDate.Text + '''';      //수정
            sWhere := sWhere + ' AND gubn_neawon <> ''5'' ';
            sOrderBy := ' ORDER BY cod_jisa, cod_hul, num_samp';
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

            labStatus.Caption := FieldByName('num_id').AsString;
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

            //분주번호 추출
            if Trim(FieldByName('cod_hul').AsString) <> '' then
            begin
               qryProfile.Active := False;
               qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_hul').AsString;
               qryProfile.Active := True;
               if qryProfile.RecordCount > 0 then
               begin
                  // 본사 -> 분주번호 추출...
                  if UV_sJisa = '15' then
                  begin
                     bunju_no := qryProfile.FieldByName('gubn_gmsa').AsInteger;

                     if Copy(FieldByName('cod_hul').AsString,1,2) = 'SS' then
                     begin
                        bunju_no := 90;
                     end;

                     qryHcondition.Active := False;
                     qryHcondition.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                     qryHcondition.ParamByName('num_id').AsString := FieldByName('num_id').AsString;
                     qryHcondition.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                     qryHcondition.Active := True;
                     if qryHcondition.RecordCount > 0 then
                     begin
                       if Trim(qryHcondition.FieldByName('HUL_CK').AsString) = 'Y' then
                          bunju_no := 30;
                     end;

                     case bunju_no of
                     10, 20 : begin
                          sBunCheck := IntToStr(iBun10);
                          iBun10 := iBun10 + 1;
                          end;
             {        20 : begin
                          sBunCheck := IntToStr(iBun20);
                          iBun20 := iBun20 + 1;
                          end;  }
                     30 : begin
                          sBunCheck := IntToStr(iBun30);
                          iBun30 := iBun30 + 1;
                          end;
                     50 : begin
                          sBunCheck := IntToStr(iBun50);
                          iBun50 := iBun50 + 1;
                          end;
                     60 : begin
                          sBunCheck := IntToStr(iBun60);
                          iBun60 := iBun60 + 1;
                          end;
                     70 : begin
                          sBunCheck := IntToStr(iBun70);
                          iBun70 := iBun70 + 1;
                          end;
                     90 : begin
                          sBunCheck := IntToStr(iBun90);
                          iBun90 := iBun90 + 1;
                          end;
                     end;
                  end
                  else
                  // 지사 -> 분주번호 추출...
                  begin
                     case qryProfile.FieldByName('gubn_gmsa').AsInteger of
                     10 : begin
//                          ijBun10 := UF_BunjuNo(ijBun10, 1000);
                          sBunCheck := IntToStr(ijBun10);
                          ijBun10 := ijBun10 + 1;
                          end;
                     20 : begin
//                          ijBun20 := UF_BunjuNo(ijBun20, 4000);
                          sBunCheck := IntToStr(ijBun20);
                          ijBun20 := ijBun20 + 1;
                          end;
                     30 : begin
//                          ijBun30 := UF_BunjuNo(ijBun30, 5000);
                          sBunCheck := IntToStr(ijBun30);
                          ijBun30 := ijBun30 + 1;
                          end;
                     50 : begin
//                          ijBun50 := UF_BunjuNo(ijBun50, 6000);
                          sBunCheck := IntToStr(ijBun50);
                          ijBun50 := ijBun50 + 1;
                          end;
                     60 : begin
//                          ijBun60 := UF_BunjuNo(ijBun60, 7000);
                          sBunCheck := IntToStr(ijBun60);
                          ijBun60 := ijBun60 + 1;
                          end;
                     70 : begin
//                          ijBun70 := UF_BunjuNo(ijBun70, 8999);
                          sBunCheck := IntToStr(ijBun70);
                          ijBun70 := ijBun70 + 1;
                          end;
                     90 : begin
//                          ijBun90 := UF_BunjuNo(ijBun90, 9999);
                          sBunCheck := IntToStr(ijBun90);
                          ijBun90 := ijBun90 + 1;
                          end;
                     end;
                  end;

                  // 분주 Table Insert.
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
            //장비번호 추출
//            else if (Trim(FieldByName('cod_jangbi').AsString) <> '') and (sJangCk = 'T') then
            else if (Trim(FieldByName('cod_jangbi').AsString) <> '') then
            begin
               qryProfile.Active := False;
               qryProfile.ParamByName('cod_pf').AsString := FieldByName('cod_jangbi').AsString;
               qryProfile.Active := True;

             //  if sJangck = 'T' then
             //  begin
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
                        // Song. 장비Profile에도 응급생성되도록..
                        qryHcondition.Active := False;
                        qryHcondition.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                        qryHcondition.ParamByName('num_id').AsString := FieldByName('num_id').AsString;
                        qryHcondition.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                        qryHcondition.Active := True;
                        if qryHcondition.RecordCount > 0 then
                        begin
                          if Trim(qryHcondition.FieldByName('HUL_CK').AsString) = 'Y' then
                             Jbunju_no := 30;
                        end;

                        case Jbunju_no of
                        10, 20 : begin
                             sBunCheck := IntToStr(iBun10);
                             iBun10 := iBun10 + 1;
                             end;
                {        20 : begin
                             sBunCheck := IntToStr(iBun20);
                             iBun20 := iBun20 + 1;
                             end;       }
                        30 : begin
                             sBunCheck := IntToStr(iBun30);
                             iBun30 := iBun30 + 1;
                             end;
                        50 : begin
                             sBunCheck := IntToStr(iBun50);
                             iBun50 := iBun50 + 1;
                             end;
                        60 : begin
                             sBunCheck := IntToStr(iBun60);
                             iBun60 := iBun60 + 1;
                             end;
                        70 : begin
                             sBunCheck := IntToStr(iBun70);
                             iBun70 := iBun70 + 1;
                             end;
                        90 : begin
                             sBunCheck := IntToStr(iBun90);
                             iBun90 := iBun90 + 1;
                             end;
                        end;
                     end
                     else
                     begin
                        case qryProfile.FieldByName('gubn_gmsa').AsInteger of
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

                     // 분주 Table Insert.
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
               //end;
            end
            //종양코드 확인
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
//                             ijBun60 := UF_BunjuNo(ijBun60, 7000);
                             sBunCheck := IntToStr(ijBun60);
                             ijBun60 := ijBun60 + 1;
                             end;
                        end;
                     end;

                     // 분주 Table Insert.
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
                 or (FieldByName('gubn_tkgm').AsString >= '1')
                 or (FieldByName('gubn_bogun').AsString >= '1')
                 or (FieldByName('gubn_adult').AsString >= '1')
                 or (FieldByName('gubn_agab').AsString >= '1')
                 or (FieldByName('gubn_life').AsString >= '1') then
            begin
               if sBunCheck = '' then
               begin
                  if UV_sJisa = '15' then
                  begin
                     sBunCheck := IntToStr(iBun70);
                     iBun70 := iBun70 + 1;
                  end
                  else
                  begin
//                     ijBun70 := UF_BunjuNo(ijBun70, 9999);
                     sBunCheck := IntToStr(ijBun70);
                     ijBun70 := ijBun70 + 1;
                  end;

                  // 분주 Table Insert.

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
            //추가코드 확인
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

                    // 검사항목
                    qryHangmok.Filter := 'COD_HM = ''' + vCod_hm[i] + '''' +
                                       ' AND DAT_APPLY <= ''' + FieldByName('DAT_GMGN').AsString + '''';

                    // 혈액검사인지 Check
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
                         if UV_sJisa = '15' then
                         begin
                              sBunCheck := IntToStr(iBun70);
                              iBun70 := iBun70 + 1;
                         end
                         else
                         begin
//                            ijBun70 := UF_BunjuNo(ijBun70, 9999);
                              sBunCheck := IntToStr(ijBun70);
                              ijBun70 := ijBun70 + 1;
                         end;

                         // 분주 Table Insert.

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


            qryProfileG.Active := False;
            sSelect  := '';
            sWhere   := '';
            sOrderby := '';

            sSelect := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';

            sWhere := sWhere + ' A.cod_pf = ''' + FieldByName('cod_hul').AsString    + ''' or ';
            sWhere := sWhere + ' A.cod_pf = ''' + FieldByName('cod_jangbi').AsString + ''' or ';
            sWhere := sWhere + ' A.cod_pf = ''' + FieldByName('cod_cancer').AsString + '''';

            sOrderby := ' ) GROUP BY B.cod_hm ';

            qryProfileG.SQL.Clear;
            qryProfileG.SQL.Add(sSelect + sWhere + sOrderby);
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

                     if ((qryHangmok.FieldByName('cod_hm').AsString = 'S079') or
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
                            // Table disconnect

                        end; // end of if[qryI_Allergy];
                        qryI_Allergy.Active := False;
                     end;
                  end;
                  qryProfileG.Next;
               end;
            end;

            sCode := FieldByName('COD_CHUGA').AsString;
            if FieldByName('gubn_nosin').AsString = '1' then
               sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '1', FieldByName('gubn_nsyh').AsInteger);
            // 노신_재검 : 재검테이블을 사용안함.
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
                  //------------------------------------------------------------------------------
                  if Trim(qryTkgum.FieldByName('cod_prf').AsString) <> '' then
                  begin
                     qryProfileG.Active := False;
                     sSelect  := '';
                     sWhere   := '';
                     sOrderby := '';

                     sSelect := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';
                     for iTemp := 1 to Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 do
                     begin
                        if iTemp = Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 then
                           sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''''
                        else
                           sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''' or ';
                     end;
                     sOrderby := ' ) GROUP BY B.cod_hm ';

                     qryProfileG.SQL.Clear;
                     qryProfileG.SQL.Add(sSelect + sWhere + sOrderby);
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
                  end;
                  //------------------------------------------------------------------------------
                  sCode := sCode + qryTkgum.FieldByName('cod_chuga').AsString;
               end;

               qryTkgum.Active := False;
            end;

            if FieldByName('gubn_bogun').AsString = '1' then
               sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '3', FieldByName('gubn_bgyh').AsInteger);
{            else if FieldByName('gubn_bogun').AsString = '2' then
            begin
               sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '3', FieldByName('gubn_nsyh').AsInteger)
            end;
}
            if FieldByName('gubn_adult').AsString = '1' then
               sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '4', FieldByName('gubn_adyh').AsInteger);
{            else if FieldByName('gubn_adult').AsString = '2' then
            begin
               sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '4', FieldByName('gubn_nsyh').AsInteger)
            end;
}
            if FieldByName('gubn_agab').AsString = '1' then
               sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '5', FieldByName('gubn_agyh').AsInteger);
{            else if FieldByName('gubn_agab').AsString = '2' then
            begin
               sCode := sCode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_nsyh').AsInteger)
            end;
}
            if FieldByName('gubn_gong').AsString = '1' then
               sCode := sCode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '6', FieldByName('gubn_goyh').AsInteger);
{            else if FieldByName('gubn_gong').AsString = '2' then
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
                             end; // end of if[qry_cell];
                             qry_cell.Active := False;
                          end;
                     end;
                     if qryI_Gulkwa.ParamByName('gubn_part').Asstring  <> 'N' then
                     begin
                        qryI_Gulkwa.ExecSQL;
                        GP_query_log(qryI_Gulkwa, 'LD2I01', 'Q', 'Y');
                     end;

                     if ((qryHangmok.FieldByName('cod_hm').AsString = 'S079') or
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
                            // Table disconnect

                        end; // end of if[qryI_Allergy];
                        qryI_Allergy.Active := False;
                     end;
                  end;
               end;
            end;

            //일일검진 내역에 분주여부 Check.
            qryU_Gumgin.ParamByName('ysno_bunju').AsString := 'Y';
            qryU_Gumgin.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
            qryU_Gumgin.ParamByName('num_id').AsString := FieldByName('num_id').AsString;
            qryU_Gumgin.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
            qryU_Gumgin.ExecSQL;
            GP_query_log(qryU_Gumgin, 'LD2I01', 'Q', 'Y');

            //Hgulkwa_chk 혈액결과완료 체크 테이블에 insert
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

            //Hgulkwa_EtcChk [센터자체처리]혈액결과완료 체크 테이블에 insert
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
   Gubn_rack;
   GF_MsgBox('Information', '작업이 정상적으로 수행되었습니다.');

end;

procedure TfrmLD2I01.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // F2인지 Check
   if Key <> GC_Refer then exit;

   // Hot-Key를 사용하여 Button을 누른 효과를 냄
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

function  TfrmLD2I01.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2I01.SpeedButton1Click(Sender: TObject);
begin
   inherited;
   if Label6.Caption = '타이머 미작동 중입니다..'  then
   begin
      if not GF_MsgBox('Confirmation', '타이머를 사용하시겠습니까??') then exit;
   end
   else
   begin
      if not GF_MsgBox('Confirmation', '타이머를 중지하시겠습니까??') then exit;
   end;

   if Label6.Caption = '타이머 미작동 중입니다..' then
   begin
      Label6.Caption := '타이머 작동을 시작합니다.';
      if SpinEdit1.Value = 24 then
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Am ' + IntToStr(SpinEdit1.Value - 12) + ':' + SpinEdit2.Text
      else if SpinEdit1.Value >= 12 then
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Pm ' + IntToStr(SpinEdit1.Value - 12) + ':' + SpinEdit2.Text
      else
         Label7.Caption := SpinEdit3.Text + '-' + GF_LPad(SpinEdit4.Text,2,'0') + '-' + GF_LPad(SpinEdit5.Text,2,'0') +
                           ' Am ' + SpinEdit1.Text + ':' + SpinEdit2.Text;

      // 오늘일자를 구함(Server System Date)
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
      Label6.Caption  := '타이머 미작동 중입니다..';
      Label7.Caption  := '설정 값이 없습나다..';
      Label14.Caption := '현재 시간을 표기됩니다..';
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
      Label14.Caption := '작업시작했습니다..!!';

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
var sRack_num,gubn_rack:String;
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
         else if sNum_bunju>=4000 Then
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
     else
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
     end;
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
   vCheck_01,vCheck_08, vCheck_01_01, vCheck_08_01, vCheck_01_02 : Boolean;
   sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
begin
   inherited;
      i := 1;
      qry_Bunju_Rack.Active := false;
      qry_Bunju_Rack.ParamByName('cod_bjjs').asString :=UV_sJisa;
      qry_Bunju_Rack.ParamByName('dat_bunju').asString :=mskDate.text;
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
              vCheck_01   := False;  vCheck_08 := False;
              vCheck_01_01 := False; vCheck_08_01 := False; vCheck_01_02 := False;
              Gauge1.Progress  := Gauge1.Progress+1;
              application.ProcessMessages;
              //------------------------------------------------------------------------
              //검사항목 ALL Display...
              //------------------------------------------------------------------------
              //1.특수검진항목체크...
              if (qry_Bunju_Rack.FieldByName('gubn_tkgm').AsString <> '') and (qry_Bunju_Rack.FieldByName('gubn_tkgm').AsString <> '0')then
              begin
                   sTk_Hangmok_Pf := qry_Bunju_Rack.FieldByName('cod_prf').AsString;
                   sTk_Hangmok_hm := qry_Bunju_Rack.FieldByName('Tcod_chuga').AsString;
              end;
              
              //2.노신유형Display
              if qry_Bunju_Rack.FieldByName('gubn_nosin').AsString  = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '1', qry_Bunju_Rack.FieldByName('gubn_nsyh').AsInteger);
              //3.성인병유형Display
              if qry_Bunju_Rack.FieldByName('gubn_adult').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '4', qry_Bunju_Rack.FieldByName('gubn_adyh').AsInteger);
              //4.간염유형Display
              if qry_Bunju_Rack.FieldByName('gubn_agab').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '5', qry_Bunju_Rack.FieldByName('gubn_agyh').AsInteger);
              //5.생애전환기유형Display
              if qry_Bunju_Rack.FieldByName('gubn_life').AsString = '1' then
                 sEtc_Hangmok_hm := sEtc_Hangmok_hm + UF_No_Hangmok(copy(GV_sToday,1,4), '7', qry_Bunju_Rack.FieldByName('gubn_lfyh').AsInteger);
              
              //추가검사항목
              sEtc_Hangmok_hm := sEtc_Hangmok_hm + qry_Bunju_Rack.FieldByName('cod_chuga').AsString;

              //프로파일 조건만들기...
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
                    else                                           sWhere2 := sWhere2 + copy(sEtc_Hangmok_hm, (i*4)-3, 4) + ''',''';
                 end;
              end;

              sTotal_HangmokList := '';
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
                       //생화학/진단면역 샘플...
                       if (FieldByName('gubn_part').AsString  = '01') or
                          (FieldByName('gubn_part').AsString  = '05') or
                          (FieldByName('gubn_part').AsString  = '07') then
                          begin
                           if (FieldByName('cod_hm').AsString <> 'S016') and
                              (FieldByName('cod_hm').AsString <> 'S021') and
                              (FieldByName('cod_hm').AsString <> 'E068') and
                              (FieldByName('cod_hm').AsString <> 'T006') and
                              (FieldByName('cod_hm').AsString <> 'T007') and
                              (FieldByName('cod_hm').AsString <> 'SE02') Then
                              vCheck_01 := True;

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
                           end;
                           if (FieldByName('gubn_part').AsString  = '08') then
                              begin
                           vCheck_08 := True;
                           if FieldByName('cod_hm').AsString <> 'E015' then vCheck_08_01 := True;
                       end;
                       //===============================================================
                       Next;
                    end;
                 end;
              end;
              
              if  qry_Bunju_Rack.FieldByName('cod_jisa').AsString = '15' then
              begin
                  if (vCheck_01) then UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
              end
              else
              begin
                 case qry_Bunju_Rack.FieldByName('gubn_hulgum').AsInteger of
                      1,2 : begin
                          if  (vCheck_01) then UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                          end;
                      3 : begin
                           if ((Copy(qry_Bunju_Rack.FieldByName('gubn_nosin').AsString,1,1)='1') or
                              (Copy(qry_Bunju_Rack.FieldByName('gubn_adult').AsString,1,1)='1') or
                              (Copy(qry_Bunju_Rack.FieldByName('gubn_life').AsString,1,1)='1')) then
                           begin
                              if (vCheck_01_02) and (vCheck_01) then  UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                           end
                           else if (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) <> 'SS') and (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) <> 'SS') and
                              (Copy(qry_Bunju_Rack.FieldByName('cod_hul').AsString,1,2) <> 'GS')    and   (Copy(qry_Bunju_Rack.FieldByName('cod_jangbi').AsString,1,2) <> 'GS') then
                           begin
                               if (vCheck_01_02) and (vCheck_01) then  UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                           end
                           else
                           begin
                               if (vCheck_01) then UP_num_rack('A',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
                           end;
                        end;
                 end;
              end;
              if (vCheck_08 = True) and (vCheck_01 = False) and (vCheck_08_01) then UP_num_rack('C',qry_Bunju_Rack.FieldByName('num_bunju').AsInteger,qry_Bunju_Rack.FieldByName('dat_bunju').AsString);
              qry_Bunju_Rack.next;
         end;
      end;
end;




end.
