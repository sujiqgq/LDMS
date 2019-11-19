//==============================================================================
// 프로그램 설명 : U057 && 유전자 검진자 LIST
// 시스템        : 통합검진
// 개발일자      : 2014.09.16
// 개발자        : 곽윤설
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.05.07
// 수정자        : 곽윤설
// 수정내용      : 항목 추가 (MM01, MF01, MC01, M501) / 항목 삭제 (M500)
// 참고사항      : 본원-최정미
//==============================================================================
// 수정일자      : 2015.06.03
// 수정자        : 곽윤설
// 수정내용      : 항목 추가 (M403)
// 참고사항      : 본원-최정미
//==============================================================================
// 수정일자      : 2015.09.04
// 수정자        : 곽윤설
// 수정내용      : 메디전 검사항목 수가변동
// 참고사항      : 본원-최정미
//==============================================================================
// 수정일자      : 2015.10.22
// 수정내용      : 메디젠 유전자 검사 : M500,MC3A,MC3B,MC4A,MM3B,MF4D,MC06,MC12,MC04,MF04,MM13,MF15,MM48,MF52 추가
// 참고사항      : [한의 재단진단검사의학팀1500879]
//==============================================================================
// 수정일자      : 2016.05.06
// 수정내용      : E068 항목 추가
// 참고사항      : 본원 진단 최은영
//==============================================================================
// 수정일자      : 2017.05.20
// 수정내용      : RE01 항목 추가
// 참고사항      : 한의재단일반파트1700253
//==============================================================================
unit LD2Q11;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt,ComObj;

type
  TfrmLD2Q11 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    btnBdate_From: TSpeedButton;
    Label5: TLabel;
    edtBDate_From: TDateEdit;
    qryGumgin: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    RadioGroup1: TRadioGroup;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    edtBDate_To: TDateEdit;
    btnBdate_To: TSpeedButton;
    Label1: TLabel;
    edtDc: TEdit;
    btnDc: TSpeedButton;
    edtD_dc: TEdit;
    Label2: TLabel;
    Panel3: TPanel;
    RBtn_Gmgn: TRadioButton;
    RBtn_Bunju: TRadioButton;
    RB_Bunju: TRadioButton;
    RB_Samp: TRadioButton;
    Label3: TLabel;
    BunjuNo1: TMaskEdit;
    BunjuNo2: TMaskEdit;
    SampNo1: TMaskEdit;
    SampNo2: TMaskEdit;
    Label4: TLabel;
    CB_HM: TComboBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    Label6: TLabel;
    CB_15: TCheckBox;
    CB_12: TCheckBox;
    CB_11: TCheckBox;
    CB_43: TCheckBox;
    CB_51: TCheckBox;
    CB_71: TCheckBox;
    CB_61: TCheckBox;
    Label8: TLabel;
    qryDNA: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin,  UV_vSampNo, UV_vDesc_dc, UV_vDat_gmgn,
    UV_vcod_gmsa, UV_vDanga : Variant;

    //메디젠,SK 유전자
    Arr_DNA : Array of Array of String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD2Q11: TfrmLD2Q11;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q321 ;

{$R *.DFM}

procedure TfrmLD2Q11.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q11.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then
   begin
      if (ActiveControl is Ttsgrid) then
      begin
         if Message.wParam > 0 then
            keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
            keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;
         grdMaster.Refresh;
      end;
   end;
end;

procedure TfrmLD2Q11.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc   := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_gmsa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDanga     := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vName,      iLength);
      VarArrayRedim(UV_vJumin,     iLength);
      VarArrayRedim(UV_vDesc_dc,   iLength);
      VarArrayRedim(UV_vSampNo,    iLength);
      VarArrayRedim(UV_vDat_gmgn,  iLength);
      VarArrayRedim(UV_vCod_gmsa,  iLength);
      VarArrayRedim(UV_vDanga,     iLength);
   end;
end;

function TfrmLD2Q11.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate_From.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end
   else if edtBDate_To.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q11.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   //기본항목설정
   CB_HM.ItemIndex := 0;
end;

procedure TfrmLD2Q11.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDesc_dc[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vJumin[DataRow-1];
      5 : Value := UV_vDat_gmgn[DataRow-1];
      6 : Value := UV_vSampNo[DataRow-1];
      7 : Value := UV_vCod_gmsa[DataRow-1];
      8 : Value := UV_vDanga[DataRow-1];
   end;
end;

procedure TfrmLD2Q11.btnQueryClick(Sender: TObject);
var index, I, k, danga, tot_danga : Integer;
    sSelect, sWhere, sOrderBy : String;
    chuga, vCod_chuga, sJisa, sLike : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;  tot_danga := 0; sJisa := ''; sLike := '';


   //유전자 검사항목
   Arr_DNA := nil;
   if (CB_HM.ItemIndex = 1) or (CB_HM.ItemIndex = 2) or (CB_HM.ItemIndex = 4) or (CB_HM.ItemIndex = 6) or (CB_HM.ItemIndex = 7) or (CB_HM.ItemIndex = 8)then
   begin
      with qryDNA do
      begin
         Active := False;
         if CB_HM.ItemIndex = 1 then
           ParamByName('sDNA').AsString := 'M'
         else if CB_HM.ItemIndex = 2 then
           ParamByName('sDNA').AsString := 'S'
         else if CB_HM.ItemIndex = 4 then
           ParamByName('sDNA').AsString := 'C'
         else if CB_HM.ItemIndex = 6 then
           ParamByName('sDNA').AsString := 'B'
         else if CB_HM.ItemIndex = 7 then
           ParamByName('sDNA').AsString := 'D'
         else if CB_HM.ItemIndex = 8 then
           ParamByName('sDNA').AsString := 'A';

         Active := True;
         if RecordCount > 0 then
         begin
            k:=0;
            SetLength(Arr_DNA, RecordCount);
            while not Eof do
            begin
               SetLength(Arr_DNA[k], 3);
               Arr_DNA[k][0] := FieldByName('Cod_hm').AsString;
               Arr_DNA[k][1] := FieldByName('Suga').AsString;
               Arr_DNA[k][2] := '0';
               inc(k);
               Next;
            end;
         end;
         Active := False;
      end;
   end;


   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryGumgin do
   begin
      // SQL문 생성
      Close;

      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, D.desc_dc, G.cod_jisa, G.cod_chuga, G.num_samp, B.cod_bjjs, B.Dat_Bunju, B.num_Bunju' +
                 ' FROM Gumgin G with(nolock) left outer join bunju B with(nolock) ON G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and G.dat_gmgn = B.dat_gmgn ' +
                 '                            left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc ';

      sWhere := ' WHERE ';

      if CB_15.Checked  then  sJisa := sJisa + '''15'',';
      if CB_12.Checked  then  sJisa := sJisa + '''12'',';
      if CB_11.Checked  then  sJisa := sJisa + '''11'',';
      if CB_43.Checked  then  sJisa := sJisa + '''43'',';
      if CB_51.Checked  then  sJisa := sJisa + '''51'',';
      if CB_61.Checked  then  sJisa := sJisa + '''61'',';
      if CB_71.Checked  then  sJisa := sJisa + '''71'',';

      if sJisa <> '' then
         sWhere  := sWhere + ' G.cod_jisa in (' + copy(sJisa,1,length(sJisa)-1) + ') AND ';

      if RBtn_Gmgn.Checked then
      begin
         sWhere  := sWhere + ' G.Dat_gmgn  >= ''' + edtBDate_From.Text + ''' ' +
                             ' AND G.Dat_gmgn  <= ''' + edtBDate_To.Text + ''' ';
      end
      else
      begin
         sWhere  := sWhere + ' B.Dat_Bunju  >= ''' + edtBDate_From.Text + ''' ' +
                             ' AND B.Dat_Bunju  <= ''' + edtBDate_To.Text + ''' ';
      end;

      if RB_Bunju.Checked then
      begin
         If Trim(BunjuNo1.Text) <> '' Then
            sWhere := sWhere + ' And B.Num_Bunju >= ''' + BunjuNo1.Text + ''' ' +
                               ' And B.Num_Bunju <= ''' + BunjuNo2.Text + ''' ';
      end
      else
      begin
         If Trim(SampNo1.Text) <> '' Then
            sWhere := sWhere + ' And G.Num_Samp >= ''' + SampNo1.Text + ''' ' +
                               ' And G.Num_Samp <= ''' + SampNo2.Text + ''' ';
      end;

      if Trim(edtDc.Text) <> '' then
        sWhere := sWhere + ' AND G.cod_dc = ''' + edtDc.Text + ''' ';

      if  CB_HM.ItemIndex = 0 then
      begin
         sWhere := sWhere + ' And (G.cod_chuga like ''%U057%'') ';
      end
      else if CB_HM.ItemIndex = 3 then
      begin
         sWhere := sWhere + ' And (G.cod_chuga like ''%E068%'') ';
      end
      else if CB_HM.ItemIndex = 5 then
      begin
         sWhere := sWhere + ' And (G.cod_chuga like ''%RE01%'') ';
      end
      else if (CB_HM.ItemIndex = 1) or (CB_HM.ItemIndex = 2) or (CB_HM.ItemIndex = 4) or (CB_HM.ItemIndex = 6) or (CB_HM.ItemIndex = 7) or (CB_HM.ItemIndex = 8)then
      begin
         for k := 0 to length(Arr_DNA)-1 do
         begin
             sLike := sLike + 'G.cod_chuga like ''%' + Arr_DNA[k][0] + '%'' or ';
         end;
         sWhere := sWhere  + ' And ( ' + copy(sLike,1,length(sLike)-3) + ') ';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY B.Num_Bunju, G.cod_jisa ';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         2 : sOrderBy := ' ORDER BY G.dat_gmgn, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Gride.Progress := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD2Q11', 'Q', 'N');
         for I := 1 to RecordCount do
         begin
            Gride.Progress := I;

            chuga := ''; vCod_chuga := ''; danga := 0;
            chuga := qryGumgin.FieldByName('cod_chuga').AsString;

            if CB_HM.ItemIndex = 0 then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]        := IntToStr(Index+1);
               UV_vDat_gmgn[Index]  := qryGumgin.FieldByName('Dat_gmgn').AsString;
               UV_vSampNo[Index]    := qryGumgin.FieldByName('num_samp').AsString;
               UV_vJumin[Index]     := qryGumgin.FieldByName('num_jumin').AsString;
               UV_vName[Index]      := qryGumgin.FieldByName('desc_name').AsString;
               UV_vDesc_dc[Index]   := qryGumgin.FieldByName('desc_dc').AsString;
               UV_vCod_gmsa[Index]  := 'U057';
               danga := danga + 100000;
               UV_vDanga[Index]     := IntToStr(danga);
               tot_danga := tot_danga + danga;

               Inc(Index);
            end
            else if CB_HM.ItemIndex = 3 then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]        := IntToStr(Index+1);
               UV_vDat_gmgn[Index]  := qryGumgin.FieldByName('Dat_gmgn').AsString;
               UV_vSampNo[Index]    := qryGumgin.FieldByName('num_samp').AsString;
               UV_vJumin[Index]     := qryGumgin.FieldByName('num_jumin').AsString;
               UV_vName[Index]      := qryGumgin.FieldByName('desc_name').AsString;
               UV_vDesc_dc[Index]   := qryGumgin.FieldByName('desc_dc').AsString;
               UV_vCod_gmsa[Index]  := 'E068';
               danga := danga + 27000;
               UV_vDanga[Index]     := IntToStr(danga);
               tot_danga := tot_danga + danga;

               Inc(Index);
            end
            else if CB_HM.ItemIndex = 5 then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]        := IntToStr(Index+1);
               UV_vDat_gmgn[Index]  := qryGumgin.FieldByName('Dat_gmgn').AsString;
               UV_vSampNo[Index]    := qryGumgin.FieldByName('num_samp').AsString;
               UV_vJumin[Index]     := qryGumgin.FieldByName('num_jumin').AsString;
               UV_vName[Index]      := qryGumgin.FieldByName('desc_name').AsString;
               UV_vDesc_dc[Index]   := qryGumgin.FieldByName('desc_dc').AsString;
               UV_vCod_gmsa[Index]  := 'RE01';
               danga := danga + 24000;
               UV_vDanga[Index]     := IntToStr(danga);
               tot_danga := tot_danga + danga;

               Inc(Index);
            end
            else //if CB_HM.ItemIndex = 1 then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]        := IntToStr(Index+1);
               UV_vDat_gmgn[Index]  := qryGumgin.FieldByName('Dat_gmgn').AsString;
               UV_vSampNo[Index]    := qryGumgin.FieldByName('num_samp').AsString;
               UV_vJumin[Index]     := qryGumgin.FieldByName('num_jumin').AsString;
               UV_vName[Index]      := qryGumgin.FieldByName('desc_name').AsString;
               UV_vDesc_dc[Index]   := qryGumgin.FieldByName('desc_dc').AsString;

               for k:=0 to length(Arr_DNA)-1 do
               begin
                 if Pos(Arr_DNA[k][0], chuga) > 0 then
                 begin
                    vCod_chuga := vCod_chuga + Arr_DNA[k][0];
                    danga := danga + StrToInt(Arr_DNA[k][1]);
                    Arr_DNA[k][2] := IntToStr(StrToInt(Arr_DNA[k][2]) + 1);
                 end;
               end;
               UV_vCod_gmsa[Index] := vCod_chuga;
               UV_vDanga[Index]    := IntToStr(danga);
               tot_danga := tot_danga + danga;

               Inc(Index);
            end;
            Next;
         end;
         UV_vNo[Index]       := '';
         UV_vDat_gmgn[Index] := '';
         UV_vSampNo[Index]   := '';
         UV_vJumin[Index]    := '';
         UV_vName[Index]     := '';
         UV_vDesc_dc[Index]  := '';
         UV_vCod_gmsa[Index] := '총합계';
         UV_vDanga[Index]    := IntToStr(tot_danga);

         Inc(Index);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD2Q11.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
      end;
   end
   else if Sender = btnBdate_From then GF_CalendarClick(edtBDate_From)
   else if Sender = btnBdate_To then GF_CalendarClick(edtBDate_To);

   if RB_Bunju.Checked then
   begin
      BunjuNo1.Enabled := True;
      BunjuNo2.Enabled := True;
      SampNo1.Enabled := False;
      SampNo2.Enabled := False;
   end
   else if RB_Samp.Checked then
   begin
      BunjuNo1.Enabled := False;
      BunjuNo2.Enabled := False;
      SampNo1.Enabled := True;
      SampNo2.Enabled := True;
   end
end;

procedure TfrmLD2Q11.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q321 := TfrmLD4Q321.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q321.QuickRep.Preview
  else                                frmLD4Q321.QuickRep.Print;
end;

procedure TfrmLD2Q11.FormActivate(Sender: TObject);
begin
   inherited;
   edtBdate_From.SetFocus;
end;

procedure TfrmLD2Q11.SBut_ExcelClick(Sender: TObject);
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
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
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
