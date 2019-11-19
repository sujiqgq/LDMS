//==============================================================================
// 프로그램 설명 : 검사항목분주현황
// 시스템        : 분석관리
// 수정일자      : 2015.02.06
// 수정자        : 구수정
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.02.13
// 수정자        : 곽윤설
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD9Q02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD9Q02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Label4: TLabel;
    RadioGroup1: TRadioGroup;
    qryProfile: TQuery;
    Label3: TLabel;
    Cmb_Hm: TComboBox;
    qry_LdmsHM_GR: TQuery;
    qry_LdmsHm: TQuery;
    Edt_To: TEdit;
    Label5: TLabel;
    Edt_From: TEdit;
    Cmb_qry: TComboBox;
    RG_jisa: TRadioGroup;
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
    procedure UP_SetInitGrid;
    procedure QRCompositeReportAddReports(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vBUN, UV_vRackNo, UV_vSamp, UV_vName, UV_vBirth : Variant;
    vArr_HM : array of Variant;
    HM_Count : Integer;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD9Q02: TfrmLD9Q02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD9Q021;

{$R *.DFM}

procedure TfrmLD9Q02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD9Q02.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //수정
      end;
   end;
end;

procedure TfrmLD9Q02.UP_VarMemSet(iLength : Integer);
var
   i : Integer;
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo     := VarArrayCreate([0, 0], varOleStr);
      UV_vName   := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN    := VarArrayCreate([0, 0], varOleStr);
      UV_vSamp   := VarArrayCreate([0, 0], varOleStr);
      UV_vRackNo := VarArrayCreate([0, 0], varOleStr);
      UV_vBirth  := VarArrayCreate([0, 0], varOleStr);
      for i := 0 to HM_Count-1 do
      begin
         vArr_HM[i] := VarArrayCreate([0, 0], varOleStr);
      end;
   end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vSamp,  iLength);
      VarArrayRedim(UV_vRackNo,iLength);
      VarArrayRedim(UV_vBirth, iLength);
      for i := 0 to  HM_Count -1 do
      begin
         VarArrayRedim(vArr_HM[i],  iLength);
      end;     
   end;
end;

function TfrmLD9Q02.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD9Q02.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

   // 오늘일자로 기본설정
   mskDate.Text := GV_sToday;

   qry_LdmsHM_GR.Active := False;
   qry_LdmsHM_GR.ParamByName('PROGRAM_ID').AsString := 'LD9Q02';
   qry_LdmsHM_GR.Active := True;
   if qry_LdmsHM_GR.RecordCount > 0 then
   begin
      while not qry_LdmsHM_GR.Eof do
      begin
          Cmb_Hm.Items.Add(qry_LdmsHM_GR.FieldByName('GROUP_HM').AsString);
          qry_LdmsHM_GR.Next;
      end;
   end;
   Cmb_Hm.ItemIndex   := 0;
   Cmb_qry.ItemIndex  := 0;
   Edt_From.MaxLength := 6;
   Edt_To.MaxLength   := 6;
   UP_SetInitGrid;
end;

procedure TfrmLD9Q02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
//var
//   i : Integer;
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 :  Value := UV_vNo[DataRow-1];   //No
      2 :  Value := UV_vRackno[DataRow-1];  //RackNo
      3 :  Value := UV_vSamp[DataRow-1];  //샘플번호
      4 :  Value := UV_vBUN[DataRow-1];  //분주번호
      5 :  Value := UV_vBirth[DataRow-1]; //생년월일
      6 :  Value := UV_vName[DataRow-1]; //성명
   end;

   if DataCol >= 7 then
   begin
      Value := vArr_HM[DataCol-7][DataRow-1];
   end;

end;

procedure TfrmLD9Q02.btnQueryClick(Sender: TObject);
var index, I, J : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sHul_List : String;
    vArr_code : array of Variant;
    vSum_tot : array of Integer;
    bTemp : Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ' WHERE '; sOrderBy := '';
   Index := 0; J := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   //배열길이
   SetLength(vArr_HM, HM_Count);   //항목필드
   SetLength(vArr_code, HM_Count); //검사코드
   SetLength(vSum_tot, HM_Count);  //검사 총합계

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);

   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  D.Desc_dc, B.desc_rackno, B.num_Bunju, G.num_jumin, G.num_samp, G.desc_name, G.num_id, G.dat_gmgn, B.dat_bunju, G.cod_jisa, G.gubn_nosin, G.gubn_adult, ' +
                 '         G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul,          ' +
                 '         G.cod_cancer, G.cod_chuga, T.cod_prf, T.cod_chuga As Tcod_chuga FROM bunju B with(nolock)                                                               ' +
                 '         Inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn                                         ' +
                 '         Left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ' +
                 '         Left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc                                                                                            ' ;

      case RG_Jisa.ItemIndex of
         0 : sWhere := sWhere + ' B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
         1 : sWhere := sWhere + ' G.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      end;

      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if Cmb_qry.ItemIndex = 0 then
      begin
         if Trim(Edt_From.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + Edt_From.Text + ''' ';
         if Trim(Edt_To.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + Edt_To.Text + ''' ';
      end
      else if Cmb_qry.ItemIndex = 1 then
      begin
         if Trim(Edt_From.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + Edt_From.Text + ''' ';
         if Trim(Edt_To.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + Edt_To.Text + ''' ';
      end
      else
      begin
         if Trim(Edt_From.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno like ''' + copy(Edt_From.Text,1,1) + ':' + copy(Edt_From.Text,2,2) + '%'' ';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY B.Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD9Q02', 'Q', 'N');

         qry_LdmsHm.Active := False;
         qry_LdmsHm.ParamByName('PROGRAM_ID').AsString     := 'LD9Q02';
         qry_LdmsHm.ParamByName('GROUP_HM').AsString       := Cmb_Hm.Text;
         qry_LdmsHm.Active := True;
         if qry_LdmsHm.RecordCount > 0 then
         begin
              while not qry_LdmsHm.Eof do
              begin
                  vArr_code[J] := qry_LdmsHm.FieldByName('cod_hm').AsString;
                  inc(J);
                  qry_LdmsHm.Next;
              end;
         end;

         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;
            bTemp := False;
            for J := 0 to HM_Count-1 do
            begin
               if pos(vArr_code[J], sHul_List) > 0 then
               begin
                   bTemp := True;
                   vArr_HM[J][index] := '*';
                   vSum_tot[J] := vSum_tot[J] + 1;
               end
               else vArr_HM[J][index] := '';
            end;
            if bTemp = True then
            begin
                UP_VarMemSet(Index+1);
                UV_vNo[Index]    := index+1;
                UV_vRackNo[Index]:= qryBunju.FieldByName('desc_rackno').AsString;
                UV_vSamp[Index]  := qryBunju.FieldByName('num_samp').AsString;
                UV_vBUN[Index]   := qryBunju.FieldByName('num_bunju').AsString;
                UV_vName[Index]  := qryBunju.FieldByName('desc_name').AsString;
                case StrToInt(copy(qryBunju.FieldByName('num_jumin').AsString,7,1)) of
                    1,3,5,7,9 : UV_vBirth[Index] := 'M/' + copy(qryBunju.FieldByName('num_jumin').AsString,1,6);
                    2,4,6,8   : UV_vBirth[Index] := 'F/' + copy(qryBunju.FieldByName('num_jumin').AsString,1,6);
                end;

                Inc(Index);
            end;

            Next;
         End;
         // Table과 Disconnect
         Close;

         UV_vNo[Index]    :='';
         UV_vRackNo[Index]:='';
         UV_vSamp[Index]  :='';
         UV_vBUN[Index]   :='';
         UV_vBirth[Index] :='';
         UV_vName[Index]  :='총합계';

         for i := 0 to HM_Count-1 do
         begin
             vArr_HM[i][index] := vSum_tot[i];
         end;

         Inc(Index);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD9Q02.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;
   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end
   else if Sender = Cmb_qry then
   begin
       Edt_From.Text      := '';
       Edt_To.Text        := '';
       Edt_To.Visible     := True;
       Label5.Visible     := True;
       if Cmb_qry.ItemIndex = 0 then
       begin
          Edt_From.MaxLength := 6;
          Edt_To.MaxLength   := 6;
       end
       else if Cmb_qry.ItemIndex = 1 then
       begin
          Edt_From.MaxLength := 4;
          Edt_To.MaxLength   := 4;
       end
       else
       begin
          Edt_From.MaxLength := 3;
          Edt_To.Visible     := False;
          Label5.Visible     := False;
       end;
   end;
end;

procedure TfrmLD9Q02.UP_Click(Sender: TObject);
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
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD9Q02.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD9Q021 := TfrmLD9Q021.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD9Q021.QuickRep.Preview
  else                                frmLD9Q021.QuickRep.Print;

end;

procedure TfrmLD9Q02.FormActivate(Sender: TObject);
begin
   inherited;
   mskDate.SetFocus;
end;

procedure TfrmLD9Q02.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD9Q021 := TfrmLD9Q021.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD9Q021.QuickRep);
  end;
end;

function TfrmLD9Q02.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
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

      // 2. 종양에 대한 검사항목 추출
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

      // 3. 장비에 대한 검사항목 추출
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

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
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

function TfrmLD9Q02.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD9Q02.UP_SetInitGrid;
var
   i : Integer;
begin

      grdMaster.RowS := 0;
      qry_LdmsHm.Active := False;
      qry_LdmsHm.ParamByName('PROGRAM_ID').AsString     := 'LD9Q02';
      qry_LdmsHm.ParamByName('GROUP_HM').AsString       := Cmb_Hm.Text;
      qry_LdmsHm.Active := True;
      if qry_LdmsHm.RecordCount > 0 then
      begin
         HM_Count := qry_LdmsHm.RecordCount;
         grdMaster.cols := HM_Count + 6;

         grdMaster.Col[1].Heading   := 'No.';
         grdMaster.Col[2].Heading   := 'Rack No.';
         grdMaster.Col[3].Heading   := '샘플번호';
         grdMaster.Col[4].Heading   := '분주번호';
         grdMaster.Col[5].Heading   := '성별/생년월일';
         grdMaster.Col[6].Heading   := '성명';

         for i := 7 to grdMaster.cols do
         begin
            grdMaster.Col[i].Heading := qry_LdmsHm.FieldByName('desc_kor').AsString;
            qry_LdmsHm.Next;
         end;
      end;
      qry_LdmsHm.Active := False;

end;


end.
