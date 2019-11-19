// 수정일자      : 2017.05.29
// 수정자        : 박수지
// 수정내용      : 최은영선임 요청
// 참고사항      : 한의 재단일반파트1700262
//==============================================================================
{
Estogen(E021) 
Calcitonin(E016)
Insulin Ab(C038) 
NSE(T012) 
Estradiol(E051) 
CA125(T030) 
FSH(E052) 
LH(E053) 
Testosterone(E056)
Myoglobin(S085) 
Nicotine 정량(T039) 
SCC(T009)
충란(Y005)
Protozoa(Y008) 
PCP(U037) 
요중페놀(U047) 
BNP(C081) 
PL(C065) 
Alcohol(C070) 
Troponin-1(C072)}
unit LD2Q18;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD2Q18 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryGulkwa: TQuery;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    RadioGroup1: TRadioGroup;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    ck_nothing: TCheckBox;
    Label4: TLabel;
    cmbHM: TComboBox;
    Label6: TLabel;
    Ck_Exist: TCheckBox;
    qry_hangmok: TQuery;
    Cmb_gubn: TComboBox;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    MEdt_SampE: TMaskEdit;
    Label8: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryPGulkwa: TQuery;
    Gride: TGauge;
    Ck_PPlus: TCheckBox;
    qryProfile: TQuery;
    RadioButton1: TRadioButton;
    RadioButton3: TRadioButton;
    cmbjisa: TComboBox;
    Label9: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN, UV_vS007, UV_vSampNo, UV_vDesc_dc, UV_vPS007, UV_vSex, UV_vAge : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD2Q18: TfrmLD2Q18;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q601;

{$R *.DFM}

procedure TfrmLD2Q18.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q18.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD2Q18.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo      := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN     := VarArrayCreate([0, 0], varOleStr);
      UV_vName    := VarArrayCreate([0, 0], varOleStr);
      UV_vS007    := VarArrayCreate([0, 0], varOleStr);
      UV_vPS007   := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vAge     := VarArrayCreate([0, 0], varOleStr);
      UV_vSex     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);

  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,      iLength);
      VarArrayRedim(UV_vName,    iLength);
      VarArrayRedim(UV_vJumin,   iLength);
      VarArrayRedim(UV_vJisa,    iLength);
      VarArrayRedim(UV_vBUN,     iLength);
      VarArrayRedim(UV_vS007,    iLength);
      VarArrayRedim(UV_vPS007,   iLength);
      VarArrayRedim(UV_vDesc_dc, iLength);
      VarArrayRedim(UV_vAge,     iLength);
      VarArrayRedim(UV_vSex,     iLength);
      VarArrayRedim(UV_vSampNo,  iLength);
   end;
end;

function TfrmLD2Q18.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') OR (cmbHM.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q18.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   //분주일자 오늘일자로 기본 설정
   edtBDate.Text := GV_sToday;

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   SBut_Excel.Enabled := TRUE;
   cmbHM.ItemIndex := 0;
   grdMaster.Col[8].Visible := False;

end;

procedure TfrmLD2Q18.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vSampNo[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vAge[DataRow-1];
      6 : Value := UV_vSex[DataRow-1];
      7 : Value := UV_vS007[DataRow-1];
      8 : Value := UV_vPS007[DataRow-1];
   end;

end;

procedure TfrmLD2Q18.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp, iAge : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy, sMan : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List, Sexs, UV_fPGulkwa, UV_fPGulkwa1, UV_fPGulkwa2, UV_fPGulkwa3 : String;
    vCod_chuga, sGubn_part : Variant;
    Check :Boolean;
begin
   inherited;   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   if (ck_nothing.Checked = False) AND (Ck_Exist.Checked = False) AND (Ck_PPlus.Checked = False) then
   begin
      ShowMessage('결과유형을 하나 이상 선택하세요.');
      exit;
   end;

   //항목명
   if cmbHM.ItemIndex = 0 then grdMaster.Col[7].Heading := 'Estogen(E021)'
   else if cmbHM.ItemIndex = 1 then grdMaster.Col[7].Heading := 'Calcitonin(E016)'
   else if cmbHM.ItemIndex = 2 then grdMaster.Col[7].Heading := 'Insulin Ab(C038)'
   else if cmbHM.ItemIndex = 3 then grdMaster.Col[7].Heading := 'NSE(T012)'
   else if cmbHM.ItemIndex = 4 then grdMaster.Col[7].Heading := 'Estradiol(E051)'
   else if cmbHM.ItemIndex = 5 then grdMaster.Col[7].Heading := 'CA125(T030)'
   else if cmbHM.ItemIndex = 6 then grdMaster.Col[7].Heading := 'FSH(E052)'
   else if cmbHM.ItemIndex = 7 then grdMaster.Col[7].Heading := 'LH(E053)'
   else if cmbHM.ItemIndex = 8 then grdMaster.Col[7].Heading := 'Testosterone(E056)'
   else if cmbHM.ItemIndex = 9  then grdMaster.Col[7].Heading := 'Myoglobin(S085)'
   else if cmbHM.ItemIndex = 10 then grdMaster.Col[7].Heading := 'Nicotine 정량(T039)'
   else if cmbHM.ItemIndex = 11 then grdMaster.Col[7].Heading := 'SCC(T009)'
   else if cmbHM.ItemIndex = 12 then grdMaster.Col[7].Heading := '충란(Y005)'
   else if cmbHM.ItemIndex = 13 then grdMaster.Col[7].Heading := 'Protozoa(Y008)'
   else if cmbHM.ItemIndex = 14 then grdMaster.Col[7].Heading := 'PCP(U037)'
   else if cmbHM.ItemIndex = 15 then grdMaster.Col[7].Heading := '요중페놀(U047)'
   else if cmbHM.ItemIndex = 16 then grdMaster.Col[7].Heading := 'BNP(C081)'
   else if cmbHM.ItemIndex = 17 then grdMaster.Col[7].Heading := 'PL(C065)'
   else if cmbHM.ItemIndex = 18 then grdMaster.Col[7].Heading := 'Alcohol(C070)'
   else if cmbHM.ItemIndex = 19 then grdMaster.Col[7].Heading := 'Troponin-1(C072)'
   else if cmbHM.ItemIndex = 20 then grdMaster.Col[7].Heading := 'Myoglubin'
   else if cmbHM.ItemIndex = 21 then grdMaster.Col[7].Heading := 'Rubella IgM,Rubella IgG'
   else if cmbHM.ItemIndex = 22 then grdMaster.Col[7].Heading := 'osteocalcin'
   else if cmbHM.ItemIndex = 23 then grdMaster.Col[7].Heading := 'PT,aPTT'
   else if cmbHM.ItemIndex = 24 then grdMaster.Col[7].Heading := 'Vitamin B12'
   else if cmbHM.ItemIndex = 25 then grdMaster.Col[7].Heading := 'B2-MG'
   else if cmbHM.ItemIndex = 26 then grdMaster.Col[7].Heading := '이온화칼슘'
   else if cmbHM.ItemIndex = 27 then grdMaster.Col[7].Heading := 'Filaria/Malaria';

   //grdMaster.Col[7].Heading := cmbHM.text;

   //종전양성 체크시- 20141031 곽윤설
   if Ck_PPlus.Checked then grdMaster.Col[8].Visible := True
   else grdMaster.Col[8].Visible := False;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  B.num_Bunju, G.desc_name, G.num_id, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, K.desc_glkwa1, K.desc_glkwa2, K.desc_glkwa3,' +
                 ' Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                 ' FROM Gulkwa K with(nolock) left outer join Gumgin G with(nolock) ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '                            left outer join Bunju B with(nolock) ON K.num_id = B.num_id and K.dat_gmgn = B.dat_gmgn' +
                 '                            left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc ' +
                 '                            left outer join Tkgum T  with(nolock) ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ' +
                   ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ';

      //생화학
      if (cmbHM.Text = 'C065') or (cmbHM.Text = 'C070') OR  (cmbHM.Text = 'C072') OR  (cmbHM.Text = 'C081') then
         sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''01'') '
      //혈액학
      else if (cmbHM.Text ='H027/H028') OR (cmbHM.Text ='H038/H039') then
         sWhere  := sWhere + ' AND ( K.Gubn_Part = ''02'') '
      //요화학
      else if (cmbHM.Text = 'U037') or (cmbHM.Text = 'U047') or (cmbHM.Text = 'Y005') or (cmbHM.Text = 'Y008') then
         sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''03'') '
      //RIA
      else if (cmbHM.Text = 'C044') then
         sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''04'') '
      //EIA
      else if (cmbHM.Text = 'T030') or (cmbHM.Text = 'T039') or (cmbHM.Text = 'E051') or (cmbHM.Text = 'E052') or (cmbHM.Text = 'E053') or (cmbHM.Text = 'E056') or
         (cmbHM.Text = 'S085') or (cmbHM.Text = 'T009')  or (cmbHM.Text = 'S018/S019') or (cmbHM.Text = 'S085')  or (cmbHM.Text = 'T033')  then
        sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''05'') '
      else if (cmbHM.Text = 'T012') or (cmbHM.Text = 'C038') or (cmbHM.Text = 'E016') or (cmbHM.Text = 'E021') or (cmbHM.Text = 'S054') OR (cmbHM.Text = 'C054') then
        sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''08'')  ';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';




      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(BunjuNo1.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju >= ''' + BunjuNo1.Text + '''';
         if Trim(BunjuNo2.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju <= ''' + BunjuNo2.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         1 : sOrderBy := ' ORDER BY K.Num_Bunju';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD2Q18', 'Q', 'N');
         Gride.Progress := 0;
         while not qryGulkwa.Eof do
         Begin
            Sexs := ''; Check := False;
            Gride.Progress := Gride.Progress + 1;
            sHul_List := ''; sGubn_part := '';
            sHul_List := UF_hangmokList;
            UV_fPGulkwa := '';
            if ((cmbHM.Text = 'E021') and (pos('E021', sHul_List) > 0)) or
               ((cmbHM.Text = 'E016') and (pos('E016', sHul_List) > 0)) or
               ((cmbHM.Text = 'C038') and (pos('C038', sHul_List) > 0)) or
               ((cmbHM.Text = 'T012') and (pos('T012', sHul_List) > 0)) or
               ((cmbHM.Text = 'E051') and (pos('E051', sHul_List) > 0)) or
               ((cmbHM.Text = 'T030') and (pos('T030', sHul_List) > 0)) or
               ((cmbHM.Text = 'E052') and (pos('E052', sHul_List) > 0)) or
               ((cmbHM.Text = 'E053') and (pos('E053', sHul_List) > 0)) or
               ((cmbHM.Text = 'E056') and (pos('E056', sHul_List) > 0)) or
               ((cmbHM.Text = 'S085') and (pos('S085', sHul_List) > 0)) or
               ((cmbHM.Text = 'T039') and (pos('T039', sHul_List) > 0)) or
               ((cmbHM.Text = 'T009') and (pos('T009', sHul_List) > 0)) or
               ((cmbHM.Text = 'Y005') and (pos('Y005', sHul_List) > 0)) or
               ((cmbHM.Text = 'Y008') and (pos('Y008', sHul_List) > 0)) or
               ((cmbHM.Text = 'U037') and (pos('U037', sHul_List) > 0)) or
               ((cmbHM.Text = 'U047') and (pos('U047', sHul_List) > 0)) or
               ((cmbHM.Text = 'C081') and (pos('C081', sHul_List) > 0)) or
               ((cmbHM.Text = 'C065') and (pos('C065', sHul_List) > 0)) or
               ((cmbHM.Text = 'C070') and (pos('C070', sHul_List) > 0)) or
               ((cmbHM.Text = 'C072') and (pos('C072', sHul_List) > 0)) or
               ((cmbHM.Text = 'H027/H028') and ((pos('H027', sHul_List) > 0) or (pos('H028', sHul_List) > 0))) or
               ((cmbHM.Text = 'H038/H039') and ((pos('H038', sHul_List) > 0) or (pos('H039', sHul_List) > 0))) or
               ((cmbHM.Text = 'S018/S019') and ((pos('S018', sHul_List) > 0) or (pos('S019', sHul_List) > 0))) or
               ((cmbHM.Text = 'S085') and (pos('S085', sHul_List) > 0)) or
               ((cmbHM.Text = 'T033') and (pos('T033', sHul_List) > 0)) or
               ((cmbHM.Text = 'S054') and (pos('S054', sHul_List) > 0)) or
               ((cmbHM.Text = 'C044') and (pos('C044', sHul_List) > 0)) or
               ((cmbHM.Text = 'C054') and (pos('C054', sHul_List) > 0))  then
            begin
                UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                if Ck_PPlus.Checked = True then                                 //20141031 곽윤설 ~
                begin
                        if (cmbHM.Text = 'C065') or (cmbHM.Text = 'C070') OR  (cmbHM.Text = 'C072') OR  (cmbHM.Text = 'C081')   then  sGubn_part := '01'
                   Else if (cmbHM.Text = 'H027/H028') OR (cmbHM.Text = 'H038/H039') then sGubn_part := '02'
                   Else if (cmbHM.Text = 'U037') or (cmbHM.Text = 'U047') or (cmbHM.Text = 'Y005') or (cmbHM.Text = 'Y008') then   sGubn_part := '03'
                   Else if (cmbHM.Text = 'C044') then   sGubn_part := '04'
                   Else if (cmbHM.Text = 'T030') or (cmbHM.Text = 'T039') or (cmbHM.Text = 'E051') or (cmbHM.Text = 'E052') or (cmbHM.Text = 'E053') or (cmbHM.Text = 'E056') or
                           (cmbHM.Text = 'S085') or (cmbHM.Text = 'T009') or (cmbHM.Text = 'S018/S019') or (cmbHM.Text = 'S085') or (cmbHM.Text = 'T033')  then sGubn_part := '05'
                   Else if  (cmbHM.Text = 'T012') or (cmbHM.Text = 'C038') or (cmbHM.Text = 'E016') or (cmbHM.Text = 'E021') or (cmbHM.Text = 'S054') or (cmbHM.Text = 'C054') then sGubn_part := '08';

                   qryPGulkwa.Close;
                   with qryPGulkwa do
                   begin
                      Active := False;
                      ParamByName('dat_bunju').AsString   := edtBDate.Text;
                      ParamByName('num_id').AsString      := qryGulkwa.FieldByName('num_id').AsString;
                      ParamByName('gubn_part').AsString   := sGubn_part;
                      Active := True;
                      if RecordCount > 0 then
                      begin
                         UV_fPGulkwa1 := qryPGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fPGulkwa2 := qryPGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fPGulkwa3 := qryPGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fPGulkwa1, UV_fPGulkwa2, UV_fPGulkwa3, UV_fPGulkwa);
                      end;
                   end;
                   qryPGulkwa.Close;
                end;                                                            //~ 20141031 곽윤설

                // 주민번호를 이용하여 성별과 나이를 구함
                GP_JuminToAge(qryGulkwa.FieldByName('num_jumin').AsString ,edtBDate.Text, sMan, iAge);


                if (cmbHM.Text = 'C070') AND (pos('C070', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 313,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 313,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 313,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C038') AND (pos('C038', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 37,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 37,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 37,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C072') AND (pos('C072', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 319,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 319,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 319,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 319,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 319,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 319,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 319,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C081') AND (pos('C081', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 373,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 373,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 373,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 373,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 373,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 373,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 373,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C065') AND (pos('C065', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 277,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 277,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 277,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 277,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 277,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 277,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 277,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'U037') AND (pos('U037', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 223,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 223,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 223,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 223,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 223,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'U047') AND (pos('U047', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 295,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 295,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 295,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 295,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 295,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 295,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 295,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'Y005') AND (pos('Y005', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 91,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 91,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 91,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 91,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 91,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 91,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 91,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'Y008') AND (pos('Y008', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 235,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 235,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 235,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 235,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 235,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 235,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 235,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'T030') AND (pos('T030', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 205,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 205,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 205,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 205,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 205,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 205,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 205,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'T039') AND (pos('T039', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 313,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 313,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 313,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E051') AND (pos('E051', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 259,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 259,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 259,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 259,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 259,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 259,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 259,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E052') AND (pos('E052', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 265,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 265,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 265,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 265,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 265,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 265,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 265,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E053') AND (pos('E053', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 271,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 271,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 271,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 271,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 271,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 271,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 271,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E056') AND (pos('E056', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 289,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 289,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 289,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 289,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 289,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 289,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 289,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'S085') AND (pos('S085', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 175,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 175,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 175,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 175,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 175,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 175,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 175,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'T009') AND (pos('T009', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 49,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 49,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 49,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 49,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 49,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 49,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 49,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'T012') AND (pos('T012', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 505,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 505,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 505,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 505,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 505,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 505,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 505,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C083') AND (pos('C083', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 37,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 37,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 37,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E016') AND (pos('E016', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 157,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 157,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 157,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 157,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 157,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 157,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 157,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'E021') AND (pos('E021', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 187,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 187,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 187,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 187,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 187,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 187,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 187,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'H027/H028') AND ((pos('H027', sHul_List) > 0) or (pos('H028', sHul_List) > 0)) then
                begin
                   if (ck_nothing.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 151,  6)) = '') or (Trim(Copy(UV_fGulkwa, 157,  6)) = ''))  then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 151,  6))+ '/' + Trim(Copy(UV_fGulkwa, 157,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 151,  6)) <> '') or (Trim(Copy(UV_fGulkwa, 157,  6)) <> '')) then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 151,  6))+ '/' + Trim(Copy(UV_fGulkwa, 157,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND ((Trim(Copy(UV_fPGulkwa, 151,  6)) <> '') or (Trim(Copy(UV_fPGulkwa, 157,  6)) <> '')) then
                   begin
                         UV_vPS007[Index]  := Trim(Copy(UV_fPGulkwa, 151,  6))+ '/' + Trim(Copy(UV_fPGulkwa, 157,  6));
                         UV_vS007[Index]   := Trim(Copy(UV_fGulkwa, 151,  6))+ '/' + Trim(Copy(UV_fGulkwa, 157,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'H038/H039') AND ((pos('H038', sHul_List) > 0) or (pos('H039', sHul_List) > 0)) then
                begin
                   if (ck_nothing.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 217,  6)) = '') or (Trim(Copy(UV_fGulkwa, 223,  6)) = ''))  then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 217,  6))+ '/' + Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 217,  6)) <> '') or (Trim(Copy(UV_fGulkwa, 223,  6)) <> '')) then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 217,  6))+ '/' + Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND ((Trim(Copy(UV_fPGulkwa, 217,  6)) <> '') or (Trim(Copy(UV_fPGulkwa, 223,  6)) <> '')) then
                   begin
                         UV_vPS007[Index]  := Trim(Copy(UV_fPGulkwa, 217,  6))+ '/' + Trim(Copy(UV_fPGulkwa, 223,  6));
                         UV_vS007[Index]   := Trim(Copy(UV_fGulkwa, 217,  6))+ '/' + Trim(Copy(UV_fGulkwa, 223,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'S018/S019') AND ((pos('S018', sHul_List) > 0) or (pos('S019', sHul_List) > 0)) then
                begin
                   if (ck_nothing.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 79,  6)) = '') or (Trim(Copy(UV_fGulkwa, 79,  6)) = ''))  then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 73,  6))+ '/' + Trim(Copy(UV_fGulkwa, 79,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 73,  6)) <> '') or (Trim(Copy(UV_fGulkwa, 79,  6)) <> '')) then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 73,  6))+ '/' + Trim(Copy(UV_fGulkwa, 79,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND ((Trim(Copy(UV_fPGulkwa, 73,  6)) <> '') or (Trim(Copy(UV_fPGulkwa, 79,  6)) <> '')) then
                   begin
                         UV_vS007[Index]   := Trim(Copy(UV_fGulkwa, 73,  6))+ '/' + Trim(Copy(UV_fGulkwa, 79,  6));
                         UV_vPS007[Index]  := Trim(Copy(UV_fPGulkwa, 73,  6))+ '/' + Trim(Copy(UV_fPGulkwa, 79,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C044') AND (pos('C044', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 1,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 1,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 1,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 1,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 1,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 1,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 1,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'C054') AND (pos('C054', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 49,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 49,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 49,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 49,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 49,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 49,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 49,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'S085') AND (pos('S085', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 175,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 175,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 175,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 175,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 175,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 175,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 175,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'T033') AND (pos('T033', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 223,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 223,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 223,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 223,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 223,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 223,  6));
                         Check := True;
                   end;
                end
                else if (cmbHM.Text = 'S054') AND (pos('S054', sHul_List) > 0) then
                begin
                   if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 469,  6)) = '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 469,  6));
                      Check := True;
                   end;
                   if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 469,  6)) <> '') then
                   begin
                      UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 469,  6));
                      Check := True;
                   end;
                   if (Ck_PPlus.Checked = True) AND (Trim(Copy(UV_fPGulkwa, 469,  6)) <> '') then
                   begin
                         UV_vS007[Index]  := Trim(Copy(UV_fGulkwa, 469,  6));
                         UV_vPS007[Index] := Trim(Copy(UV_fPGulkwa, 469,  6));
                         Check := True;
                   end;
                end;

                if Check = True then
                begin
                   UV_VBUN[Index]    := qryGulkwa.FieldByName('Num_Bunju').AsString;
                   UV_vName[Index]   := qryGulkwa.FieldByName('desc_name').AsString;
                   UV_vSampNo[Index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vNo[Index]     := IntToStr(Index+1);
                   UV_vSex[Index]    := sMan;
                   UV_vAge[Index]    := iAge;
                   Inc(Index);
                end
            end;
            UP_VarMemSet(Index+1);
            qryGulkwa.Next;
         End;
         // Table과 Disconnect
         qryGulkwa.Close;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster
   .Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;


procedure TfrmLD2Q18.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

   if Cmb_gubn.ItemIndex = 1 then
   begin
      MEdt_SampS.Enabled := True;
      MEdt_SampE.Enabled := True;
      bunjuno1.Enabled := False;
      bunjuno2.Enabled := False;
   end
   else
   begin
      MEdt_SampS.Enabled := False;
      MEdt_SampE.Enabled := False;
      bunjuno1.Enabled := True;
      bunjuno2.Enabled := True;
   end;
end;

procedure TfrmLD2Q18.UP_Click(Sender: TObject);
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

procedure TfrmLD2Q18.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
   Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD2Q18.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  frmLD4Q601 := TfrmLD4Q601.Create(Self);

  with QRCompositeReport do
  begin
    Reports.Add(frmLD4Q601.QuickRep);
  end;
end;

function TfrmLD2Q18.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryGulkwa.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('cod_hul').AsString;
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
      if qryGulkwa.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_cancer').AsString;
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
      if qryGulkwa.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryGulkwa.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryGulkwa.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryGulkwa.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '1', qryGulkwa.FieldByName('Gubn_nsyh').AsInteger)
   else if qryGulkwa.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '4', qryGulkwa.FieldByName('Gubn_adyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '7', qryGulkwa.FieldByName('Gubn_lfyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '3',qryGulkwa.FieldByName('Gubn_bgyh').AsInteger);

   if qryGulkwa.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryGulkwa.FieldByName('Dat_gmgn').AsString, '5',qryGulkwa.FieldByName('Gubn_agyh').AsInteger);

   if (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '1') or (qryGulkwa.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryGulkwa.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryGulkwa.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryGulkwa.FieldByName('Dat_gmgn').AsString;
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

function TfrmLD2Q18.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2Q18.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;

end;

end.
