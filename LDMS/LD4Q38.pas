//==============================================================================
// 프로그램 설명 : EIA분주현황
// 시스템        : 통합검진
// 수정일자      : 2011.5.27
// 수정자        : 김희석
// 수정내용      :
//==============================================================================
// 프로그램 설명 : EIA분주현황
// 시스템        : 통합검진
// 수정일자      : 2014.04.07
// 수정자        : 곽윤설
// 수정내용      : T043 제거
//==============================================================================
// 수정일자      : 2014.05.03
// 수정자        : 곽윤설
// 수정내용      : T040 제거
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.08.21
// 수정자        : 곽윤설
// 수정내용      : 20140901 분주일자부터 T042조회 안되도록..
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2016.06.07
// 수정자        : 박수지
// 수정내용      : 각 항목별 체크박스 생성..
// 참고사항      : 본원 진검 한미영
//==============================================================================


unit LD4Q38;

interface
                                                
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,
  QuickRpt;

type
  TfrmLD4Q38 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    qryProfile: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    group_radio: TGroupBox;
    Chk_ALL: TRadioButton;
    Chk_TT03: TRadioButton;
    Chk_T041: TRadioButton;
    Chk_S011: TRadioButton;
    GroupBox1: TGroupBox;
    Chk_T042: TRadioButton;
    CHK_S012: TRadioButton;
    Chk_S010: TRadioButton;
    Chk_S033: TRadioButton;
    Btn_NamePrint: TBitBtn;
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
    procedure QRCompositeReportAddReports(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure Btn_NamePrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vS010, UV_vT041,UV_vT042,UV_vS011,UV_vS012, UV_vTT03, UV_vS033, UV_vBUN, UV_vDesc_dc,
    UV_vDat_gmgn, UV_vNum_samp, UV_vSex : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;    
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q38: TfrmLD4Q38;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q381, LD4Q382;

{$R *.DFM}

procedure TfrmLD4Q38.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q38.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q38.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vS010  := VarArrayCreate([0, 0], varOleStr);
      UV_vT041  := VarArrayCreate([0, 0], varOleStr);
      UV_vT042  := VarArrayCreate([0, 0], varOleStr);
      UV_vS011  := VarArrayCreate([0, 0], varOleStr);
      UV_vS012  := VarArrayCreate([0, 0], varOleStr);
      UV_vTT03  := VarArrayCreate([0, 0], varOleStr);
      UV_vS033  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vSex      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vBUN,  iLength);
      VarArrayRedim(UV_vS010, iLength);
      VarArrayRedim(UV_vT041,   iLength);
      VarArrayRedim(UV_vT042,  iLength);
      VarArrayRedim(UV_vS011,  iLength);
      VarArrayRedim(UV_vS012,  iLength);
      VarArrayRedim(UV_vTT03,  iLength);
      VarArrayRedim(UV_vS033,  iLength);
      VarArrayRedim(UV_vDesc_dc, iLength);
      VarArrayRedim(UV_vDat_gmgn, iLength);
      VarArrayRedim(UV_vSex, iLength);
      VarArrayRedim(UV_vNum_samp, iLength);
   end;
end;

function TfrmLD4Q38.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q38.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
      Cmb_gubn.ItemIndex := 1;

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q38.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;


   if Chk_all.Checked =true  then
   begin
   grdMaster.Col[5].visible  := True;
   grdMaster.Col[6].visible  := True;
   grdMaster.Col[7].visible  := True;
   grdMaster.Col[8].visible  := True;
      begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vBUN[DataRow-1];
            3 : Value := UV_vName[DataRow-1];
            4 : Value := UV_vS010[DataRow-1];
            5 : Value := UV_vT041[DataRow-1];
            6 : Value := UV_vS011[DataRow-1];
            7 : Value := UV_vS012[DataRow-1];
            8 : Value := UV_vTT03[DataRow-1];
            9 : Value := UV_vT042[DataRow-1];
           10 : Value := UV_vS033[DataRow-1];
           11 : Value := UV_vDat_gmgn[DataRow-1];
           12 : Value := UV_vNum_samp[DataRow-1];
           13 : Value := UV_vSex[DataRow-1];
           14 : Value := UV_vDesc_dc[DataRow-1];
         end;
         grdMaster.Col[4].heading   := 'HBcAb(S010)';
         grdMaster.Col[5].heading   := 'Insulin(T041)';
         grdMaster.Col[6].heading   := 'HBeAg(S011)';
         grdMaster.Col[7].heading   := 'Hbe Ab(S012)';
         grdMaster.Col[8].heading   := 'AFP(TT03)';
         grdMaster.Col[9].heading   := 'Free T3(T042)';
         grdMaster.Col[10].heading  := 'anti-HCV(S033)';
         grdMaster.Col[11].visible := false;
         grdMaster.Col[12].visible := false;
         grdMaster.Col[13].visible := false;
         grdMaster.Col[14].visible := false;


      end;
   end
   else if Chk_S010.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vS010[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'HBcAb(S010)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_T041.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vT041[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'Insulin(T041)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_S011.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vS011[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'HBeAg(S011)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_S012.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vS012[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'Hbe Ab(S012)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_TT03.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vTT03[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'AFP(TT03)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_T042.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vT042[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];

            end;
            grdMaster.Col[4].heading   := 'Free T3(T042)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end
   else if Chk_S033.Checked =true  then
      begin
         case DataCol of
               1 : Value := UV_vNo[DataRow-1];
               2 : Value := UV_vBUN[DataRow-1];
               3 : Value := UV_vName[DataRow-1];
               4 : Value := UV_vS033[DataRow-1];
               5 : Value := UV_vDat_gmgn[DataRow-1];
               6 : Value := UV_vNum_samp[DataRow-1];
               7 : Value := UV_vSex[DataRow-1];
               8 : Value := UV_vDesc_dc[DataRow-1];
            end;
            grdMaster.Col[4].heading   := 'anti-HCV(S033)';
            grdMaster.Col[5].heading   := '';
            grdMaster.Col[6].heading   := '';
            grdMaster.Col[7].heading   := '';
            grdMaster.Col[8].heading   := '';
            grdMaster.Col[9].heading   := '';
            grdMaster.Col[10].heading  := '';
            grdMaster.Col[5].visible := false;
            grdMaster.Col[6].visible := false;
            grdMaster.Col[7].visible := false;
            grdMaster.Col[8].visible := false;
      end;

   if mskDate.Text >= '20140901' then
   begin
      grdMaster.Col[9].Width := 0;
   end
   else grdMaster.Col[9].Width := 64;
end;

procedure TfrmLD4Q38.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vS010, Sum_vT041, Sum_vT042, Sum_vTT03, Sum_vS011,Sum_vS012, Sum_vS033 : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List, sJisa : String;
    vCod_chuga : Variant;
    Check,  Check_S010 :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;
   Sum_vS010:=0;
   Sum_vT041:=0;
   Sum_vT042:=0;
   Sum_vS011:=0;
   Sum_vS012:=0;
   Sum_vTT03:=0;
   Sum_vS033:=0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;

      sSelect := ' SELECT  D.Desc_dc, B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, B.dat_bunju, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, D.desc_dc,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B with(nolock) inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha'+
                 '          left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc';

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + mskDate.Text + ''' ';
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + mskDate.Text + ''' and (Gubn_Part  = ''05'' OR Gubn_part = ''07''))';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND A.cod_dc = ''' + edtDc.Text + '''';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q38', 'Q', 'N');
         For I := 1 to RecordCount do
          Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;
            check := FALSE;
           if (pos('S010',sHul_List) > 0) or (pos('T041',sHul_List) > 0) or ((pos('T042',sHul_List) > 0) and (qryBunju.FieldByName('dat_Bunju').AsString < '20140901')) or
               (pos('S011',sHul_List) > 0) or (pos('S012',sHul_List) > 0) or (pos('TT03',sHul_List) > 0) or (pos('S033',sHul_List) > 0)  then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := qryBunju.FieldByName('desc_rackno').AsString;
//               UV_vBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               if Cmb_gubn.ItemIndex = 0 then
                  UV_VBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString
               else if  Cmb_gubn.ItemIndex = 1 then
               begin
                    UV_VBUN[Index] := qryBunju.FieldByName('num_samp').AsString;
                    grdMaster.Col[2].Heading :='샘플번호';
               end;
               UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
               UV_vDesc_dc[Index]    := qryBunju.FieldByName('desc_dc').AsString;
               UV_vNum_samp[Index]   := qryBunju.FieldByName('Num_samp').AsString;
               UV_vDat_gmgn[Index]   := qryBunju.FieldByName('dat_gmgn').AsString;
               case StrToInt(copy(qryBunju.FieldByName('num_jumin').AsString,7,1)) of
                  1,3,5,7 : UV_vSex[Index] := copy(qryBunju.FieldByName('num_jumin').AsString,3,4) + 'M' + copy(qryBunju.FieldByName('num_jumin').AsString,1,2);
                  2,4,6,8 : UV_vSex[Index] := copy(qryBunju.FieldByName('num_jumin').AsString,3,4) + 'F' + copy(qryBunju.FieldByName('num_jumin').AsString,1,2);
               end;

               if (pos('S010', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_S010.Checked=true))   then
               begin
                  UV_vS010[Index] := '';
                  Sum_vS010 := Sum_vS010 + 1;
                  Check := True;
               end
               else                               UV_vS010[Index] := '*';          //Sum_vS010
               if (pos('T041', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_T041.Checked=true))  then
               begin
                  UV_vT041[Index] := '';
                  Sum_vT041 := Sum_vT041 + 1;
                  Check := True;
               end
               else                               UV_vT041[Index] := '*';          //Sum_vT041
              if (pos('T042', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_T042.Checked=true))  then
               begin
                  UV_vT042[Index] := '';
                  Sum_vT042 := Sum_vT042 + 1;
                  Check := True;
               end
               else                               UV_vT042[Index] := '*';          //Sum_vT042
               if (pos('S011', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_S011.Checked=true))  then
               begin
                  UV_vS011[Index] := '';
                  Sum_vS011 := Sum_vS011 + 1;
                  Check := True;
               end
               else                               UV_vS011[Index] := '*';          //Sum_vS011
               if (pos('S012', sHul_List) > 0)  and ((Chk_all.Checked=true) or (Chk_S012.Checked=true))  then
               begin
                  UV_vS012[Index] := '';
                  Sum_vS012 := Sum_vS012 + 1;
                  Check := True;
               end
               else                               UV_vS012[Index] := '*';          //Sum_vS012
               if (pos('TT03', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_TT03.Checked=true))  then
               begin
                  UV_vTT03[Index] := '';
                  Sum_vTT03 := Sum_vTT03 + 1;
                  Check := True;
               end
               else                               UV_vTT03[Index] := '*';          //Sum_vTT03
               if (pos('S033', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_S033.Checked=true))  then
               begin
                  UV_vS033[Index] := '';
                  Sum_vS033 := Sum_vS033 + 1;
                  Check := True;
               end
               else                               UV_vS033[Index] := '*';          //Sum_vS033

               if Check = True then
                   begin
                        UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                        if Cmb_gubn.ItemIndex = 0 then
                           UV_VBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString
                        else if  Cmb_gubn.ItemIndex = 1 then
                        begin
                             UV_VBUN[Index] := qryBunju.FieldByName('num_samp').AsString;
                             grdMaster.Col[2].Heading :='샘플번호';
                        end;
                        UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                        UP_VarMemSet(Index+1);
                        inc(index);
                   end;
             end;
             Next;


         End;

         // Table과 Disconnect
         Close;

         UV_vNo[Index]     := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '총 합 계';
         UV_vS010[Index]   := Sum_vS010;
         UV_vT041[Index]   := Sum_vT041;
         UV_vT042[Index]   := Sum_vT042;
         UV_vS011[Index]   := Sum_vS011;
         UV_vS012[Index]   := Sum_vS012;
         UV_vTT03[Index]   := Sum_vTT03;
         UV_vS033[Index]   := Sum_vS033;

         UV_vDesc_dc[Index] := '';
         Inc(index);
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q38.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q38.UP_Click(Sender: TObject);
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

procedure TfrmLD4Q38.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then QRCompositeReport.Preview
  else                                QRCompositeReport.Print;

end;

procedure TfrmLD4Q38.FormActivate(Sender: TObject);
begin
   inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';
   Cmb_gubn.ItemIndex := 1;
   mskDate.SetFocus;
end;

procedure TfrmLD4Q38.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q381 := TfrmLD4Q381.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q381.QuickRep);
  end;
end;

function TfrmLD4Q38.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k: integer;
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

function TfrmLD4Q38.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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


procedure TfrmLD4Q38.SBut_ExcelClick(Sender: TObject);
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
procedure TfrmLD4Q38.Btn_NamePrintClick(Sender: TObject);
begin
  inherited;
     frmLD4Q382 := TfrmLD4Q382.Create(Self);
   if RBtn_preveiw.Checked then frmLD4Q382.QuickRep.Preview
   else                         frmLD4Q382.QuickRep.Print;
end;

end.
