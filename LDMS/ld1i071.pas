//==============================================================================
// 프로그램 설명 : SMS 문자메세지
// 시스템        : 통합검진
// 수정일자      : 2005.02.01
// 수정자        : 변재권
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2007.04.08
// 수정자        : 유동구
// 수정자        : 단체별 추가...
//==============================================================================
// 수정일자      : 2014.06.20
// 수정자        : 곽윤설
// 수정내용      : B009추가 [병리-김원연]
//==============================================================================
// 수정일자      : 2014.12.09
// 수정자        : 곽윤설
// 수정내용      : B012추가 [조직검사(재판독)]
//==============================================================================
unit LD1I071;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit ;

type
  TfrmLD1I071 = class(TfrmSingle)
    pnlCond: TPanel;
    grdMaster: TtsGrid;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    qryU_Cell: TQuery;
    Label2: TLabel;
    cmbGubn: TComboBox;
    rbDate: TRadioButton;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    rbJumin: TRadioButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    bitSend: TBitBtn;
    Cbox_check: TCheckBox;
    qryU_Tkgum_P012: TQuery;
    qryU_Tkgum_P013: TQuery;
    cmbJisa: TComboBox;
    Label1: TLabel;
    Label3: TLabel;
    CmbOrder: TComboBox;

    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure rbDateClick(Sender: TObject);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure UP_Click(Sender: TObject);
    procedure up_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Cbox_checkClick(Sender: TObject);
    procedure bitSendClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs, UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vCod_jisa,
    UV_vDat_gmgn, UV_vCod_cell, UV_vCod_hm, UV_vCod_doct, UV_vDat_result,
    UV_vYsno_sokun, UV_vGubn_check, UV_vGubn_P012, UV_vDesc_buwi   : Variant;

    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD1I071: TfrmLD1I071;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD1I071.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_result := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_sokun := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_check := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_P012  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi  := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vCod_cell,   iLength);
      VarArrayRedim(UV_vCod_hm,     iLength);
      VarArrayRedim(UV_vCod_doct,   iLength);
      VarArrayRedim(UV_vDat_result, iLength);
      VarArrayRedim(UV_vYsno_sokun, iLength);
      VarArrayRedim(UV_vGubn_check, iLength);
      VarArrayRedim(UV_vGubn_P012,  iLength);
      VarArrayRedim(UV_vDesc_buwi,  iLength);
   end;
end;


function TfrmLD1I071.UF_Condition : Boolean;
var bError : boolean;
begin
   Result := True;
   bError := False;

   // 필수입력 조회조건 Check
   if rbDate.Checked then
   begin
      if mskDate.Text = '' then bError := True;
   end
   else
   begin
      if mskJumin.Text = '' then bError := True;
   end;

   // 오류일 경우 Message 표시
   if bError then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;



procedure TfrmLD1I071.FormCreate(Sender: TObject);
begin
   inherited;
   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   // 환경설정
   UP_VarMemSet(0);

   // SQL문 저장
   UV_sBasicSQL := qryCell.SQL.Text;

   mskDate.Text := GV_sToday;
   cmbGubn.ItemIndex := 0;
   cmbJisa.ItemIndex := 0;
   CmbOrder.ItemIndex := 0;

end;


procedure TfrmLD1I071.btnQueryClick(Sender: TObject);
var  index : Integer;
    sWhere, sOrderBy,
    tTime : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryCell do
   begin
      Active := False;

      // SQL문 생성
      sWhere := 'WHERE ';

      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
         sWhere := sWhere + ' A.DAT_RESULT = ''' + mskDate.Text + '''';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' B.num_jumin = ''' + mskJumin.Text + '''';
      end;

      sWhere := sWhere + ' AND A.ysno_sokun = ''Y''';

      // 검사구분에 따라서 Where 조건추가
      if cmbGubn.ItemIndex = 1 then
         sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''P'''
      else if cmbGubn.ItemIndex = 2 then
         sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''J'''
      else if cmbGubn.ItemIndex = 3 then
         sWhere := sWhere + ' AND cod_hm = ''B001'''
      else if cmbGubn.ItemIndex = 4 then
         sWhere := sWhere + ' AND cod_hm = ''B007'''
      else if cmbGubn.ItemIndex = 5 then
         sWhere := sWhere + ' AND cod_hm = ''B008'''
      else if cmbGubn.ItemIndex = 6 then
         sWhere := sWhere + ' AND cod_hm = ''B009'''
      else if cmbGubn.ItemIndex = 7 then
         sWhere := sWhere + ' AND cod_hm = ''B012'''
      else if cmbGubn.ItemIndex = 8 then
         sWhere := sWhere + ' AND cod_hm = ''P001'''
      else if cmbGubn.ItemIndex = 9 then
         sWhere := sWhere + ' AND cod_hm = ''P003'''
      else if cmbGubn.ItemIndex = 10 then
         sWhere := sWhere + ' AND cod_hm = ''P009'''
      else if cmbGubn.ItemIndex = 11 then
         sWhere := sWhere + ' AND cod_hm = ''P010'''
      else if cmbGubn.ItemIndex = 12 then
         sWhere := sWhere + ' AND cod_hm = ''P011'''
      else if cmbGubn.ItemIndex = 13 then
         sWhere := sWhere + ' AND cod_hm = ''P006'''
      else if cmbGubn.ItemIndex = 14 then
         sWhere := sWhere + ' AND cod_hm = ''P012'''
      else if cmbGubn.ItemIndex = 15 then
         sWhere := sWhere + ' AND cod_hm = ''P013''';

      if CmbOrder.ItemIndex = 0 then
         sOrderBy := ' ORDER BY A.desc_buwi '
      else if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' ORDER BY A.dat_gmgn, A.cod_jisa, B.desc_name ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
      Active := True;

      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD1I071', 'Q', 'N'); 
         // 조건에 맞는 데이터 변수에 입력
         while not qryCell.Eof do
         begin
            UP_VarMemSet(index);
            UV_vGubn_check[index]  := '0';
            UV_vDesc_buwi[index]   := FieldByName('desc_buwi').AsString;
            UV_vCod_bjjs[index]    := FieldByName('COD_BJJS').AsString;
            UV_vNum_id[index]      := FieldByName('NUM_ID').AsString;
            UV_vNum_jumin[index]   := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]   := FieldByName('DESC_NAME').AsString;
            UV_vCod_jisa[index]    := FieldByName('COD_JISA').AsString;
            UV_vDat_gmgn[index]    := FieldByName('DAT_GMGN').AsString;
            UV_vCod_cell[index]    := FieldByName('COD_CELL').AsString;
            UV_vCod_hm[index]      := FieldByName('COD_HM').AsString;
            UV_vCod_doct[index]    := FieldByName('COD_DOCT').AsString + '-' + FieldByName('Doc_name').AsString;
            UV_vDat_result[index]  := FieldByName('DAT_RESULT').AsString;
            UV_vYsno_sokun[index]  := FieldByName('Ysno_sokun').AsString;
            UV_vGubn_P012[index]   := FieldByName('Gubn_P012').AsString;
            Next;
            Inc(index);
         end;

         // Table과 Disconnect
         Active := False;
      end
      else
      begin
//         GP_FieldClear(pnlDetail);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD1I071.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD1I071.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      // 찾기기능 활성화
      btnFind.Enabled   := True;

      mskDate.Color     := $00E6FFFE;
      mskDate.Enabled   := True;
      btnDate.Enabled   := True;
      mskJumin.Color    := clWindow;
      mskJumin.Enabled  := False;
      btnJumin.Enabled  := False;
      mskDate.SetFocus;
   end
   else if Sender = rbJumin then
   begin
      // 찾기기능 비활성화
      btnFind.Enabled   := False;

      mskDate.Color     := clWindow;
      mskDate.Enabled   := False;
      btnDate.Enabled   := False;
      mskJumin.Color    := $00E6FFFE;
      mskJumin.Enabled  := True;
      btnJumin.Enabled  := True;
      mskJumin.SetFocus;
   end;
end;


procedure TfrmLD1I071.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1  : Value := StrToint(UV_vGubn_check[DataRow-1]);
      2  : Value := UV_vDesc_buwi[DataRow-1];
      3  : Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
      4  : begin
              if UV_vCod_hm[DataRow-1] = 'B001' then value := 'B001-조직검사(내시경)'
              else if UV_vCod_hm[DataRow-1] = 'B007' then value := 'B007-조직검사(대장내시경)'
              else if UV_vCod_hm[DataRow-1] = 'B008' then value := 'B008-액상세포병리(폐암)'
              else if UV_vCod_hm[DataRow-1] = 'B009' then value := 'B009-액상세포병리(자궁)'
              else if UV_vCod_hm[DataRow-1] = 'B010' then value := 'B010-조직검사(유방)'
              else if UV_vCod_hm[DataRow-1] = 'B012' then value := 'B012-조직검자(재판독)'
              else if UV_vCod_hm[DataRow-1] = 'P001' then value := 'P001-세포학검사(성인병)'
              else if UV_vCod_hm[DataRow-1] = 'P003' then value := 'P003-세포학검사(일반)'
              else if UV_vCod_hm[DataRow-1] = 'P009' then value := 'P009-조직검사(자궁경부)'
              else if UV_vCod_hm[DataRow-1] = 'P010' then value := 'P010-세침천자(갑상선)'
              else if UV_vCod_hm[DataRow-1] = 'P011' then value := 'P011-세침천자(유방)'
              else if UV_vCod_hm[DataRow-1] = 'P012' then value := 'P012-객담검사'
              else if copy(UV_vCod_cell[DataRow-1],1,1) = 'J' then value := '조직'
              else if copy(UV_vCod_cell[DataRow-1],1,1) = 'P' then value := 'PAP';
           end;
      5  : Value := GF_DateFormat(UV_vDat_result[DataRow-1]);
      6  : Value := UV_vDesc_name[DataRow-1];
      7  : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      8  : Value := UV_vCod_doct[DataRow-1];
   end;
end;


procedure TfrmLD1I071.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];
end;


procedure TfrmLD1I071.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   if Sender = btnJumin then
   begin
      if GF_InjoukClick(Self, sCode, sName) then
      begin
         mskJumin.Text := sCode;
         edtName.Text  := sName;
      end;
   end
   else if Sender = btnDate       then GF_CalendarClick(mskDate);
end;


procedure TfrmLD1I071.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // Reference Key가 아니면 종료
   if Key <> GC_Refer then exit;

   if      Sender = mskDate       then UP_Click(btnDate)
   else if Sender = mskJumin      then UP_Click(btnJumin);
end;


procedure TfrmLD1I071.Cbox_checkClick(Sender: TObject);
var  iTemp, i : Integer;
begin
   inherited;
   iTemp := grdMaster.Rows;
   if Cbox_check.Checked then
   begin
     Cbox_check.Caption  := '해   제';
     for I := 0 to grdMaster.Rows - 1 do
      UV_vGubn_check[I] := '1';
      grdMaster.RowS := 0;
      grdMaster.RowS := iTemp;
   end // end of if[btnall(전체선택)];
   else
   if Cbox_check.Checked = False then
   begin
     Cbox_check.Caption  := '전체선택';
     for I := 0 to grdMaster.Rows - 1 do
       UV_vGubn_check[I] := '0';
       grdMaster.RowS := 0;
       grdMaster.RowS := iTemp;
   end; // end of if[btnall(해제)];
end;


procedure TfrmLD1I071.bitSendClick(Sender: TObject);
var sCount, i : integer;
begin
   inherited;
   sCount := 0 ;
   // 메세지 전송 (em_tran 테이블에 Insert)
   for I := 0 to grdMaster.rows -1 do
   begin
      if UV_vGubn_check[I] = '1' then
      begin
         sCount := sCount + 1 ;
         with qryU_Cell do
         begin
            ParamByName('cod_bjjs').AsString   := UV_vCod_bjjs[I];
            ParamByName('num_id').AsString     := UV_vNum_id[I];
            ParamByName('cod_cell').AsString   := UV_vCod_cell[I];
            ParamByName('ysno_sokun').AsString := 'C';
            ParamByName('dat_time_R').AsString := formatdatetime('HH:NN:SS', now);
            ParamByName('dat_update').AsString := GV_sToday;
            ParamByName('cod_update').AsString := GV_sUserId;
            ExecSql;
            GP_query_log(qryU_Cell, 'LD1I071', 'Q', 'Y');
         end;

         if UV_vCod_hm[I] = 'P012' then
         begin
            with qryU_Tkgum_P012 do
            begin
               ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[I];
               ParamByName('Num_jumin').AsString  := UV_vNum_jumin[I];
               ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[I];

               if   UV_vGubn_P012[I]    = '01' then ParamByName('Desc_PBIO').AsString := '검체불량(부적합)'
               else if UV_vGubn_P012[I] = '02' then ParamByName('Desc_PBIO').AsString := 'classⅰ(정상)'
               else if UV_vGubn_P012[I] = '03' then ParamByName('Desc_PBIO').AsString := 'classⅱ(정상-반응성)'
               else if UV_vGubn_P012[I] = '04' then ParamByName('Desc_PBIO').AsString := 'classⅲ(비정형세포)'
               else if UV_vGubn_P012[I] = '05' then ParamByName('Desc_PBIO').AsString := 'classⅳ(암의심)'
               else if UV_vGubn_P012[I] = '06' then ParamByName('Desc_PBIO').AsString := 'classⅴ(암)';

               ParamByName('dat_update').AsString := GV_sToday;
               ParamByName('cod_update').AsString := GV_sUserId;
               ExecSql;
               GP_query_log(qryU_Tkgum_P012, 'LD1I071', 'Q', 'Y');
            end;
         end;

         if UV_vCod_hm[I] = 'P013' then
         begin
            with qryU_Tkgum_P013 do
            begin
               ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[I];
               ParamByName('Num_jumin').AsString  := UV_vNum_jumin[I];
               ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[I];

               if      UV_vGubn_P012[I] = '01' then ParamByName('Desc_UBIO').AsString := '검체불량'
               else if UV_vGubn_P012[I] = '02' then ParamByName('Desc_UBIO').AsString := '음성ⅰ'
               else if UV_vGubn_P012[I] = '03' then ParamByName('Desc_UBIO').AsString := '음성ⅱ'
               else if UV_vGubn_P012[I] = '04' then ParamByName('Desc_UBIO').AsString := '의양성ⅲ'
               else if UV_vGubn_P012[I] = '05' then ParamByName('Desc_UBIO').AsString := '양성ⅳ'
               else if UV_vGubn_P012[I] = '06' then ParamByName('Desc_UBIO').AsString := '양성ⅴ';

               ParamByName('dat_update').AsString := GV_sToday;
               ParamByName('cod_update').AsString := GV_sUserId;
               ExecSql;
               GP_query_log(qryU_Tkgum_P013, 'LD1I071', 'Q', 'Y');
            end;
         end;

      end;
   end;

   if sCount = 0 then ShowMessage ('판독완료된 건수가 없습니다.')
   else               ShowMessage('총 ' + inttostr(sCount) + '건의 판독완료 하였습니다.');
end;





end.
