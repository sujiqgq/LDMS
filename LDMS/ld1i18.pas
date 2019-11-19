//==============================================================================
// 프로그램 설명 : 알레르기결과등록
// 수정일자      : 2012.04.29
// 수정자        : 김승철
// 참고사항      : 추가.
//==============================================================================

unit LD1I18;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I18 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlDetail: TPanel;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlName1: TPanel;
    pnlCond: TPanel;
    rbDate: TRadioButton;
    rbJumin: TRadioButton;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    Panel47: TPanel;
    Panel17: TPanel;
    pnlNum1: TPanel;
    pnlNum2: TPanel;
    pnlName2: TPanel;
    edtValue2: TEdit;
    pnlNum3: TPanel;
    pnlName3: TPanel;
    edtValue3: TEdit;
    pnlNum4: TPanel;
    pnlName4: TPanel;
    edtValue4: TEdit;
    pnlNum5: TPanel;
    pnlName5: TPanel;
    edtValue5: TEdit;
    pnlNum6: TPanel;
    pnlName6: TPanel;
    edtValue6: TEdit;
    pnlNum7: TPanel;
    pnlName7: TPanel;
    edtValue7: TEdit;
    pnlNum8: TPanel;
    pnlName8: TPanel;
    edtValue8: TEdit;
    pnlNum9: TPanel;
    pnlName9: TPanel;
    edtValue9: TEdit;
    pnlNum10: TPanel;
    pnlName10: TPanel;
    edtValue10: TEdit;
    pnlNum11: TPanel;
    pnlName11: TPanel;
    edtValue11: TEdit;
    pnlNum12: TPanel;
    pnlName12: TPanel;
    edtValue12: TEdit;
    pnlNum13: TPanel;
    pnlName13: TPanel;
    edtValue13: TEdit;
    pnlNum14: TPanel;
    pnlName14: TPanel;
    edtValue14: TEdit;
    pnlNum15: TPanel;
    pnlName15: TPanel;
    edtValue15: TEdit;
    Panel34: TPanel;
    pnlName16: TPanel;
    pnlNum16: TPanel;
    pnlNum17: TPanel;
    pnlName17: TPanel;
    edtValue17: TEdit;
    pnlNum18: TPanel;
    pnlName18: TPanel;
    edtValue18: TEdit;
    pnlNum19: TPanel;
    pnlName19: TPanel;
    edtValue19: TEdit;
    pnlNum20: TPanel;
    pnlName20: TPanel;
    edtValue20: TEdit;
    pnlNum21: TPanel;
    pnlName21: TPanel;
    edtValue21: TEdit;
    pnlNum22: TPanel;
    pnlName22: TPanel;
    edtValue22: TEdit;
    pnlNum23: TPanel;
    pnlName23: TPanel;
    edtValue23: TEdit;
    pnlNum24: TPanel;
    pnlName24: TPanel;
    edtValue24: TEdit;
    pnlNum25: TPanel;
    pnlName25: TPanel;
    edtValue25: TEdit;
    pnlNum26: TPanel;
    pnlName26: TPanel;
    edtValue26: TEdit;
    pnlNum27: TPanel;
    pnlName27: TPanel;
    edtValue27: TEdit;
    pnlNum28: TPanel;
    pnlName28: TPanel;
    edtValue28: TEdit;
    pnlNum29: TPanel;
    pnlName29: TPanel;
    edtValue29: TEdit;
    pnlNum30: TPanel;
    pnlName30: TPanel;
    edtValue30: TEdit;
    edtValue31: TEdit;
    pnlName31: TPanel;
    pnlNum31: TPanel;
    pnlNum32: TPanel;
    pnlName32: TPanel;
    edtValue1: TEdit;
    pnlNum33: TPanel;
    pnlName33: TPanel;
    edtValue33: TEdit;
    pnlNum34: TPanel;
    pnlName34: TPanel;
    edtValue34: TEdit;
    pnlNum35: TPanel;
    pnlName35: TPanel;
    edtValue35: TEdit;
    pnlNum36: TPanel;
    pnlName36: TPanel;
    edtValue36: TEdit;
    pnlNum37: TPanel;
    pnlName37: TPanel;
    edtValue37: TEdit;
    pnlNum38: TPanel;
    pnlName38: TPanel;
    edtValue38: TEdit;
    pnlNum39: TPanel;
    pnlName39: TPanel;
    edtValue39: TEdit;
    pnlNum40: TPanel;
    pnlName40: TPanel;
    edtValue40: TEdit;
    pnlNum41: TPanel;
    pnlName41: TPanel;
    edtValue41: TEdit;
    pnlNum42: TPanel;
    pnlName42: TPanel;
    edtValue42: TEdit;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    btnPrevNum: TBitBtn;
    btnNextNum: TBitBtn;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    Cmb_gubn: TComboBox;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
    Bevel2: TBevel;
    edtValue32: TEdit;
    edtValue16: TEdit;
    GroupBox1: TGroupBox;
    Rd_Eng: TRadioButton;
    Rd_Kor: TRadioButton;
    qryAllergy: TQuery;
    Label2: TLabel;
    qryProfile: TQuery;
    qryTkgum: TQuery;
    qryDanche: TQuery;
    Panel_MAST: TPanel;
    qryAllergy_glkwa: TQuery;
    qryU_Allergy: TQuery;
    Panel2: TPanel;
    PnlTop: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label11: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    mskDAT_BUNJU: TDateEdit;
    edtNUM_BUNJU: TEdit;
    mskNUM_JUMIN: TMaskEdit;
    edtDESC_NAME: TEdit;
    mskDAT_GMGN: TDateEdit;
    EdtDesc_dc: TEdit;
    EdtNum_samp: TEdit;
    EdtDesc_dept: TEdit;
    Edt_jisa: TEdit;
    Panel4: TPanel;
    cmbCOD_SAWON: TComboBox;
    Panel5: TPanel;
    cmbCOD_DOCT: TComboBox;
    Panel3: TPanel;
    mskDAT_RESULT: TDateEdit;
    btnDAT_RESULT: TSpeedButton;
    Edt_Class32: TEdit;
    Edt_Class2: TEdit;
    Edt_Class3: TEdit;
    Edt_Class4: TEdit;
    Edt_Class5: TEdit;
    Edt_Class6: TEdit;
    Edt_Class7: TEdit;
    Edt_Class8: TEdit;
    Edt_Class9: TEdit;
    Edt_Class10: TEdit;
    Edt_Class11: TEdit;
    Edt_Class12: TEdit;
    Edt_Class13: TEdit;
    Edt_Class14: TEdit;
    Edt_Class15: TEdit;
    Edt_Class16: TEdit;
    Edt_Class17: TEdit;
    Edt_Class18: TEdit;
    Edt_Class19: TEdit;
    Edt_Class20: TEdit;
    Edt_Class21: TEdit;
    Edt_Class22: TEdit;
    Edt_Class23: TEdit;
    Edt_Class24: TEdit;
    Edt_Class25: TEdit;
    Edt_Class26: TEdit;
    Edt_Class27: TEdit;
    Edt_Class28: TEdit;
    Edt_Class29: TEdit;
    Edt_Class30: TEdit;
    Edt_Class31: TEdit;
    Edt_Class1: TEdit;
    Edt_Class33: TEdit;
    Edt_Class34: TEdit;
    Edt_Class35: TEdit;
    Edt_Class36: TEdit;
    Edt_Class37: TEdit;
    Edt_Class38: TEdit;
    Edt_Class39: TEdit;
    Edt_Class40: TEdit;
    Edt_Class41: TEdit;
    Edt_Class42: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Edt_Class43: TEdit;
    edtValue43: TEdit;
    pnlName43: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    pnlName44: TPanel;
    edtValue44: TEdit;
    Edt_Class44: TEdit;
    Edt_Class45: TEdit;
    edtValue45: TEdit;
    pnlName45: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    pnlName46: TPanel;
    edtValue46: TEdit;
    Edt_Class46: TEdit;
    Edt_Class47: TEdit;
    edtValue47: TEdit;
    pnlName47: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    pnlName48: TPanel;
    edtValue48: TEdit;
    Edt_Class48: TEdit;
    Edt_Class49: TEdit;
    edtValue49: TEdit;
    pnlName49: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    pnlName50: TPanel;
    edtValue50: TEdit;
    Edt_Class50: TEdit;
    Edt_Class51: TEdit;
    edtValue51: TEdit;
    pnlName51: TPanel;
    Panel24: TPanel;
    Panel25: TPanel;
    pnlName52: TPanel;
    edtValue53: TEdit;
    edtValue52: TEdit;
    Edt_Class52: TEdit;
    Edt_Class53: TEdit;
    Edt_Class54: TEdit;
    pnlName54: TPanel;
    edtValue54: TEdit;
    pnlName53: TPanel;
    Panel29: TPanel;
    Panel30: TPanel;
    Panel31: TPanel;
    pnlName55: TPanel;
    edtValue55: TEdit;
    Edt_Class55: TEdit;
    Edt_Class56: TEdit;
    edtValue56: TEdit;
    pnlName56: TPanel;
    Panel35: TPanel;
    Panel36: TPanel;
    pnlName57: TPanel;
    edtValue57: TEdit;
    Edt_Class57: TEdit;
    Edt_Class58: TEdit;
    Edt_Class59: TEdit;
    Edt_Class60: TEdit;
    Edt_Class61: TEdit;
    Edt_Class62: TEdit;
    edtValue62: TEdit;
    edtValue61: TEdit;
    edtValue60: TEdit;
    edtValue59: TEdit;
    edtValue58: TEdit;
    pnlName58: TPanel;
    pnlName59: TPanel;
    pnlName60: TPanel;
    pnlName61: TPanel;
    pnlName62: TPanel;
    Panel43: TPanel;
    Panel44: TPanel;
    Panel45: TPanel;
    Panel46: TPanel;
    Panel48: TPanel;
    Panel49: TPanel;

    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure btnQueryClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure rbDateClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure Rd_EngClick(Sender: TObject);
    procedure UP_MoveNum(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPSaveClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
  
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,  UV_vNum_id,   UV_vDat_bunju,  UV_vNum_bunju,  UV_vNum_jumin,
    UV_vDesc_name, UV_vDat_gmgn, UV_vCod_jisa,   UV_vCod_dc,     UV_vDesc_dept,
    UV_vNum_Sample, UV_vCod_hm  : Variant;

    sHangmok : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    function  UF_Condition : Boolean;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_GridSet;

  public
    { Public declarations }
  end;

var
  frmLD1I18: TfrmLD1I18;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD1I18.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //추가
      end;
   end;
end;

procedure TfrmLD1I18.UP_MoveNum(Sender: TObject);
begin
   inherited;

   if Sender = btnPrevNum then
   begin
      if grdMaster.CurrentDataRow > 1 then
         grdMaster.CurrentDataRow := grdMaster.CurrentDataRow - 1
      else
         exit;
   end
   else if Sender = btnNextNum then
   begin
      if grdMaster.CurrentDataRow < grdMaster.Rows then
         grdMaster.CurrentDataRow := grdMaster.CurrentDataRow + 1
      else
         exit;
   end;

   grdMaster.TopRow := grdMaster.CurrentDataRow;
end;

procedure TfrmLD1I18.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if (rbDate.Checked = True) or (rbBarcode.Checked = True) then
      begin
         Col[1].Heading := '분주번호';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';
         Col[4].Heading := '항목코드';
      end
      else if rbJumin.Checked = True then
      begin
         Col[1].Heading := '검진일자';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';
         Col[4].Heading := '항목코드';
      end;
   end;
end;

procedure TfrmLD1I18.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid 환경설정
   UP_GridSet;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryAllergy do
   begin
      Active := False;

      // SQL문을 생성

      sSelect := 'SELECT A.Cod_hm, A.COD_BJJS, A.NUM_ID, A.DAT_BUNJU, A.NUM_BUNJU, ';
      sSelect := sSelect + ' B.num_jumin, B.desc_name, B.cod_dc, B.dat_gmgn, B.gubn_hulgum, B.num_samp, B.desc_dept, ';
      sSelect := sSelect + ' B.cod_hul, B.cod_cancer, B.cod_chuga, B.cod_jangbi, ';
      sSelect := sSelect + ' B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh, ';
      sSelect := sSelect + ' B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, ';
      sSelect := sSelect + ' B.gubn_gong, B.gubn_goyh, B.cod_jisa , B.gubn_tkgm, ';
      sSelect := sSelect + ' B.Num_samp, B.Gubn_life, B.Gubn_lfyh ';
      sSelect := sSelect + ' FROM Allergy_62 A INNER JOIN gumgin B ';
      sSelect := sSelect + ' ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ';

      sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';

      if cmbRelation.ItemIndex = 1 then
         sWhere := sWhere + ' AND A.cod_hm = ''S101'' '
      else if cmbRelation.ItemIndex = 2 then
         sWhere := sWhere + ' AND A.cod_hm = ''S102'' ';

      if rbDate.Checked then
      begin
         sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' AND B.num_jumin = ''' + mskJumin.Text + '''';
      end
      else if rbBarcode.Checked then
      begin
         sWhere := sWhere + ' AND B.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
         sWhere := sWhere + ' AND B.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';
      end;

      if Cmb_gubn.ItemIndex = 0 then
         sOrderBy := ' ORDER BY A.num_bunju '
      else if Cmb_gubn.ItemIndex = 1 then
         sOrderBy := ' ORDER BY CAST(B.num_samp AS INT)';

      qryAllergy.SQL.Clear;
      qryAllergy.SQL.Add(sSelect + sWhere + sOrderBy);
      qryAllergy.Active := True;

      if qryAllergy.RecordCount > 0 then
      begin
         GP_query_log(qryAllergy, 'LD1I16', 'Q', 'N');

         for i := 0 to qryAllergy.RecordCount - 1 do
         Begin
            UP_VarMemSet(Index);

            UV_vCod_hm[index]     := FieldByName('cod_hm').AsString;
            UV_vCod_bjjs[index]   := FieldByName('COD_BJJS').AsString;
            UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
            UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
            UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
            UV_vNum_Sample[index] := FieldByName('NUM_SAMP').AsString;
            UV_vCod_dc[index]     := FieldByName('Cod_dc').AsString;
            UV_vDesc_dept[index]  := FieldByName('Desc_dept').AsString;
            UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
            UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
            UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
            Inc(index);
            Next;
         end;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GP_FieldClear(PnlTop);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');     
end;

function TfrmLD1I18.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (pnlJisa.Visible and (cmbJisa.ItemIndex = -1)) or
      (rbDate.Checked and (mskDate.Text = '')) or
      (rbJumin.Checked and (mskJumin.Text = '')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;


procedure TfrmLD1I18.FormCreate(Sender: TObject);
begin
  inherited;
   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);   
   pnlJisa.Visible := True;   
   if Copy(GV_sUserId,1,2) = '15' then
      cmbCOD_DOCT.Text := '150005-김소영';
      
   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   //검사보고일
//   mskDAT_RESULT.Text := GV_sToday;
   //검사자
   GP_ComboSawonAll(cmbCOD_SAWON, 0);
   //전문의
   GP_ComboSawonAll(cmbCOD_DOCT, 1);

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   Cmb_gubn.ItemIndex := 0;
   cmbRelation.ItemIndex := 0;
end;

procedure TfrmLD1I18.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_dc     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dept  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Sample := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_hm,     iLength);
      VarArrayRedim(UV_vCod_bjjs,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vDat_bunju,  iLength);
      VarArrayRedim(UV_vNum_bunju,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vCod_dc,     iLength);
      VarArrayRedim(UV_vDesc_dept,  iLength);
      VarArrayRedim(UV_vNum_Sample, iLength);
   end;
end;

procedure TfrmLD1I18.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
  inherited;
   // 자료를 화면에 조회
   case DataCol of
      1 : begin
             if rbDate.Checked then
                Value := UV_vNum_bunju[DataRow-1]
             else if rbJumin.Checked then
                Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1])
             else if rbBarcode.Checked then
                Value := UV_vNum_bunju[DataRow-1];
          end;
      2 : begin
             Value := UV_vNum_Sample[DataRow-1];
          end;
      3 : begin
             Value := UV_vDesc_name[DataRow-1];
          end;
      4 : begin
             Value := UV_vCod_hm[DataRow-1];
          end;
   end;
end;

procedure TfrmLD1I18.rbDateClick(Sender: TObject);
begin
  inherited;

   if Sender = rbDate then
   begin
      mskDate.Color      := $00E6FFFE;
      mskDate.Enabled    := True; btnDate.Enabled    := True;
      mskFrom.Enabled    := True; mskTo.Enabled      := True;
      MEdt_SampS.Enabled := True; MEdt_SampE.Enabled := True;
      Cmb_gubn.Enabled   := True;

      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;  btnJumin.Enabled := False;

      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled := False;

      mskDate.SetFocus;
   end
   else if Sender = rbJumin then
   begin
      mskDate.Color      := clWindow;
      mskDate.Enabled    := False;  btnDate.Enabled    := False;
      mskFrom.Enabled    := False;  mskTo.Enabled      := False;
      MEdt_SampS.Enabled := False;  MEdt_SampE.Enabled := False;
      Cmb_gubn.Enabled   := False;

      mskJumin.Color   := $00E6FFFE;
      mskJumin.Enabled := True;    btnJumin.Enabled := True;

      MEdt_Barcode.Color   := clWindow;
      MEdt_Barcode.Enabled := False;

      mskJumin.SetFocus;
   end
   else if Sender = rbBarcode then
   begin
      mskDate.Color      := clWindow;
      mskDate.Enabled    := False;  btnDate.Enabled  := False;
      mskFrom.Enabled    := False;  mskTo.Enabled    := False;
      MEdt_SampS.Enabled := False;  MEdt_SampE.Enabled := False;
      Cmb_gubn.Enabled   := False;

      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;    btnJumin.Enabled := False;

      MEdt_Barcode.Color   := $00E6FFFE;
      MEdt_Barcode.Enabled := True;

      MEdt_Barcode.SetFocus;
   end;
end;

procedure TfrmLD1I18.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i : integer;
begin
  inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;
   GP_FieldClear(pnlDetail);
   GP_FieldClear(PnlTop);

   //검사보고일
   mskDAT_RESULT.Text := UV_vDat_bunju[NewRow-1];
   //검사자
   GP_ComboDisplay(cmbCOD_SAWON, GV_sUserId, 6);
   //전문의
   if Copy(GV_sUserId,1,2) = '15' then
      cmbCOD_DOCT.Text := '150005-김소영';

   // Grid의 Row가 변경되면 Detail에 자료를 할당
   mskNUM_JUMIN.Text := UV_vNum_jumin[NewRow-1];
   edtDESC_NAME.Text := UV_vDesc_name[NewRow-1];
   mskDAT_GMGN.Text  := UV_vDat_gmgn[NewRow-1];
   mskDAT_BUNJU.Text := UV_vDat_bunju[NewRow-1];
   edtNUM_BUNJU.Text := UV_vNum_bunju[NewRow-1];
   Edt_jisa.Text     := UV_vCod_jisa[NewRow-1];
   EdtDesc_dept.Text := UV_vDesc_dept[NewRow-1];
   EdtNum_samp.Text  := UV_vNum_Sample[NewRow-1];

   //단체명
   with qryDanche do
   begin
      Active := False;
      ParamByName('COD_dc').AsString := UV_vCod_dc[NewRow-1];
      Active := True;
      if RecordCount > 0 then
         EdtDesc_dc.Text := FieldByName('desc_dc').AsString;
      Active := False;
   end;

   if UV_vCod_hm[NewRow-1] = 'S101' then
   begin
      Panel_MAST.Caption := 'MAST 흡입성 62종';
      Panel_MAST.Font.Color := clblue;
      Rd_EngClick(Rd_Kor);
   end
   else if UV_vCod_hm[NewRow-1] = 'S102' then
   begin
      Panel_MAST.Caption := 'MAST 식이성 62종';
      Panel_MAST.Font.Color := clgreen;
      Rd_EngClick(Rd_Kor);
   end;

   with qryAllergy_glkwa do
   begin
      Active := False;
      ParamByName('cod_bjjs').AsString  := UV_vCod_bjjs[NewRow-1];
      ParamByName('num_id').AsString    := UV_vnum_id[NewRow-1];
      ParamByName('dat_bunju').AsString := UV_vdat_bunju[NewRow-1];
      ParamByName('cod_hm').AsString    := UV_vcod_hm[NewRow-1];
      Active := True;
      if RecordCount > 0 then
      begin
         GP_query_log(qryAllergy_glkwa, 'LD1I16', 'Q', 'N');
         GP_ComboDisplay(cmbCOD_DOCT, FieldByName('cod_doct').AsString,  6);

         edtValue1.Text  := FieldByName('Alle_01').AsString;
         Edt_Class1.Text := FieldByName('Alle_01_Class').AsString;

         edtValue2.Text  := FieldByName('Alle_02').AsString;
         Edt_Class2.Text := FieldByName('Alle_02_Class').AsString;
         edtValue3.Text  := FieldByName('Alle_03').AsString;
         Edt_Class3.Text := FieldByName('Alle_03_Class').AsString;
         edtValue4.Text  := FieldByName('Alle_04').AsString;
         Edt_Class4.Text := FieldByName('Alle_04_Class').AsString;
         edtValue5.Text  := FieldByName('Alle_05').AsString;
         Edt_Class5.Text := FieldByName('Alle_05_Class').AsString;
         edtValue6.Text  := FieldByName('Alle_06').AsString;
         Edt_Class6.Text := FieldByName('Alle_06_Class').AsString;
         edtValue7.Text  := FieldByName('Alle_07').AsString;
         Edt_Class7.Text := FieldByName('Alle_07_Class').AsString;
         edtValue8.Text  := FieldByName('Alle_08').AsString;
         Edt_Class8.Text := FieldByName('Alle_08_Class').AsString;
         edtValue9.Text  := FieldByName('Alle_09').AsString;
         Edt_Class9.Text := FieldByName('Alle_09_Class').AsString;
         edtValue10.Text := FieldByName('Alle_10').AsString;
         Edt_Class10.Text := FieldByName('Alle_10_Class').AsString;
         edtValue11.Text := FieldByName('Alle_11').AsString;
         Edt_Class11.Text := FieldByName('Alle_11_Class').AsString;
         edtValue12.Text := FieldByName('Alle_12').AsString;
         Edt_Class12.Text := FieldByName('Alle_12_Class').AsString;
         edtValue13.Text := FieldByName('Alle_13').AsString;
         Edt_Class13.Text := FieldByName('Alle_13_Class').AsString;
         edtValue14.Text := FieldByName('Alle_14').AsString;
         Edt_Class14.Text := FieldByName('Alle_14_Class').AsString;
         edtValue15.Text := FieldByName('Alle_15').AsString;
         Edt_Class15.Text := FieldByName('Alle_15_Class').AsString;
         edtValue16.Text := FieldByName('Alle_16').AsString;
         Edt_Class16.Text := FieldByName('Alle_16_Class').AsString;
         edtValue17.Text := FieldByName('Alle_17').AsString;
         Edt_Class17.Text := FieldByName('Alle_17_Class').AsString;
         edtValue18.Text := FieldByName('Alle_18').AsString;
         Edt_Class18.Text := FieldByName('Alle_18_Class').AsString;
         edtValue19.Text := FieldByName('Alle_19').AsString;
         Edt_Class19.Text := FieldByName('Alle_19_Class').AsString;
         edtValue20.Text := FieldByName('Alle_20').AsString;
         Edt_Class20.Text := FieldByName('Alle_20_Class').AsString;
         edtValue21.Text := FieldByName('Alle_21').AsString;
         Edt_Class21.Text := FieldByName('Alle_21_Class').AsString;
         edtValue22.Text := FieldByName('Alle_22').AsString;
         Edt_Class22.Text := FieldByName('Alle_22_Class').AsString;
         edtValue23.Text := FieldByName('Alle_23').AsString;
         Edt_Class23.Text := FieldByName('Alle_23_Class').AsString;
         edtValue24.Text := FieldByName('Alle_24').AsString;
         Edt_Class24.Text := FieldByName('Alle_24_Class').AsString;
         edtValue25.Text := FieldByName('Alle_25').AsString;
         Edt_Class25.Text := FieldByName('Alle_25_Class').AsString;
         edtValue26.Text := FieldByName('Alle_26').AsString;
         Edt_Class26.Text := FieldByName('Alle_26_Class').AsString;
         edtValue27.Text := FieldByName('Alle_27').AsString;
         Edt_Class27.Text := FieldByName('Alle_27_Class').AsString;
         edtValue28.Text := FieldByName('Alle_28').AsString;
         Edt_Class28.Text := FieldByName('Alle_28_Class').AsString;
         edtValue29.Text := FieldByName('Alle_29').AsString;
         Edt_Class29.Text := FieldByName('Alle_29_Class').AsString;
         edtValue30.Text := FieldByName('Alle_30').AsString;
         Edt_Class30.Text := FieldByName('Alle_30_Class').AsString;
         edtValue31.Text := FieldByName('Alle_31').AsString;
         Edt_Class31.Text := FieldByName('Alle_31_Class').AsString;
         edtValue32.Text := FieldByName('Alle_32').AsString;
         Edt_Class32.Text := FieldByName('Alle_32_Class').AsString;
         edtValue33.Text := FieldByName('Alle_33').AsString;
         Edt_Class33.Text := FieldByName('Alle_33_Class').AsString;
         edtValue34.Text := FieldByName('Alle_34').AsString;
         Edt_Class34.Text := FieldByName('Alle_34_Class').AsString;
         edtValue35.Text := FieldByName('Alle_35').AsString;
         Edt_Class35.Text := FieldByName('Alle_35_Class').AsString;
         edtValue36.Text := FieldByName('Alle_36').AsString;
         Edt_Class36.Text := FieldByName('Alle_36_Class').AsString;
         edtValue37.Text := FieldByName('Alle_37').AsString;
         Edt_Class37.Text := FieldByName('Alle_37_Class').AsString;
         edtValue38.Text := FieldByName('Alle_38').AsString;
         Edt_Class38.Text := FieldByName('Alle_38_Class').AsString;
         edtValue39.Text := FieldByName('Alle_39').AsString;
         Edt_Class39.Text := FieldByName('Alle_39_Class').AsString;
         edtValue40.Text := FieldByName('Alle_40').AsString;
         Edt_Class40.Text := FieldByName('Alle_40_Class').AsString;
         edtValue41.Text := FieldByName('Alle_41').AsString;
         Edt_Class41.Text := FieldByName('Alle_41_Class').AsString;
         edtValue42.Text := FieldByName('Alle_42').AsString;
         Edt_Class42.Text := FieldByName('Alle_42_Class').AsString;
         edtValue43.Text := FieldByName('Alle_43').AsString;
         Edt_Class43.Text := FieldByName('Alle_43_Class').AsString;
         edtValue44.Text := FieldByName('Alle_44').AsString;
         Edt_Class44.Text := FieldByName('Alle_44_Class').AsString;
         edtValue45.Text := FieldByName('Alle_45').AsString;
         Edt_Class45.Text := FieldByName('Alle_45_Class').AsString;
         edtValue46.Text := FieldByName('Alle_46').AsString;
         Edt_Class46.Text := FieldByName('Alle_46_Class').AsString;
         edtValue47.Text := FieldByName('Alle_47').AsString;
         Edt_Class47.Text := FieldByName('Alle_47_Class').AsString;
         edtValue48.Text := FieldByName('Alle_48').AsString;
         Edt_Class48.Text := FieldByName('Alle_48_Class').AsString;
         edtValue49.Text := FieldByName('Alle_49').AsString;
         Edt_Class49.Text := FieldByName('Alle_49_Class').AsString;
         edtValue50.Text := FieldByName('Alle_50').AsString;
         Edt_Class50.Text := FieldByName('Alle_50_Class').AsString;
         edtValue51.Text := FieldByName('Alle_51').AsString;
         Edt_Class51.Text := FieldByName('Alle_51_Class').AsString;
         edtValue52.Text := FieldByName('Alle_52').AsString;
         Edt_Class52.Text := FieldByName('Alle_52_Class').AsString;
         edtValue53.Text := FieldByName('Alle_53').AsString;
         Edt_Class53.Text := FieldByName('Alle_53_Class').AsString;
         edtValue54.Text := FieldByName('Alle_54').AsString;
         Edt_Class54.Text := FieldByName('Alle_54_Class').AsString;
         edtValue55.Text := FieldByName('Alle_55').AsString;
         Edt_Class55.Text := FieldByName('Alle_55_Class').AsString;
         edtValue56.Text := FieldByName('Alle_56').AsString;
         Edt_Class56.Text := FieldByName('Alle_56_Class').AsString;
         edtValue57.Text := FieldByName('Alle_57').AsString;
         Edt_Class57.Text := FieldByName('Alle_57_Class').AsString;
         edtValue58.Text := FieldByName('Alle_58').AsString;
         Edt_Class58.Text := FieldByName('Alle_58_Class').AsString;
         edtValue59.Text := FieldByName('Alle_59').AsString;
         Edt_Class59.Text := FieldByName('Alle_59_Class').AsString;
         edtValue60.Text := FieldByName('Alle_60').AsString;
         Edt_Class60.Text := FieldByName('Alle_60_Class').AsString;
         edtValue61.Text := FieldByName('Alle_61').AsString;
         Edt_Class61.Text := FieldByName('Alle_61_Class').AsString;
         edtValue62.Text := FieldByName('Alle_62').AsString;
         Edt_Class62.Text := FieldByName('Alle_62_Class').AsString;
      end;
      Active := False;
   end;

   if edtValue32.Text = '>100' then
      edtValue32.Color := $00FACDF4
   else
      edtValue32.Color := clWindow;

   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if Pos('edtValue', TEdit(pnlDetail.Controls[i]).Name) > 0 then
      begin
         if ((TEdit(pnlDetail.Controls[i]).Name) <> 'edtValue32') then
         begin
            if (strTofloat(TEdit(pnlDetail.Controls[i]).text) >= 0.70) then
            begin
               TEdit(pnlDetail.Controls[i]).Color  := $00FACDF4;
            end
            else
               TEdit(pnlDetail.Controls[i]).Color  := clWindow;
         end;
      end;
   end;
end;

procedure TfrmLD1I18.Rd_EngClick(Sender: TObject);
begin
  inherited;

  if Rd_Kor.Checked then
  begin
     if Panel_MAST.Caption = 'MAST 흡입성 62종' then
     begin
        pnlName1.Caption  := '수중다리가루진드기';    pnlName22.Caption := '소나무';           pnlName43.Caption := '게';
        pnlName2.Caption  := '말';                    pnlName23.Caption := '삼나무';           pnlName44.Caption := '새우';
        pnlName3.Caption  := '기니피그';              pnlName24.Caption := '아카시아';         pnlName45.Caption := '감자';
        pnlName4.Caption  := '양';                    pnlName25.Caption := '편백나무';         pnlName46.Caption := '사과';
        pnlName5.Caption  := '토끼';                  pnlName26.Caption := '불란서 국화';      pnlName47.Caption := '카카오';
        pnlName6.Caption  := '햄스터';                pnlName27.Caption := '민들레';           pnlName48.Caption := '복숭아';
        pnlName7.Caption  := '개암나무';              pnlName28.Caption := '창질경이';         pnlName49.Caption := '고등어';
        pnlName8.Caption  := '향기풀';                pnlName29.Caption := '명아주과풀';       pnlName50.Caption := 'Bromelain';
        pnlName9.Caption  := '우산잔디';              pnlName30.Caption := '미역취 국화';      pnlName51.Caption := '호밀풀';
        pnlName10.Caption := '오리새';                pnlName31.Caption := '털비름';           pnlName52.Caption := '집먼지';
        pnlName11.Caption := '큰조아재비';            pnlName32.Caption := '총 IgE';           pnlName53.Caption := '바퀴벌레';
        pnlName12.Caption := '갈대';                  pnlName33.Caption := '진드기(Dp)';       pnlName54.Caption := '곰팡이류';
        pnlName13.Caption := '외겨이삭';              pnlName34.Caption := '진드기(Df)';       pnlName55.Caption := '곰팡이류';
        pnlName14.Caption := '꿀벌';                  pnlName35.Caption := '저장진드기';       pnlName56.Caption := '곰팡이류';
        pnlName15.Caption := '말벌';                  pnlName36.Caption := '고양이';           pnlName57.Caption := '오리나무';
        pnlName16.Caption := '라텍스';                pnlName37.Caption := '개';               pnlName58.Caption := '자작나무';
        pnlName17.Caption := '곰팡이류';              pnlName38.Caption := '계란흰자';         pnlName59.Caption := '참나무';
        pnlName18.Caption := '플라타너스';            pnlName39.Caption := '우유';             pnlName60.Caption := '돼지풀';
        pnlName19.Caption := '수양버들';              pnlName40.Caption := '옥수수';           pnlName61.Caption := '쑥(머그워트)';
        pnlName20.Caption := '포플라';                pnlName41.Caption := '참깨';             pnlName62.Caption := '환삼덩굴';
        pnlName21.Caption := '물푸레';                pnlName42.Caption := '콩';


     end
     else if Panel_MAST.Caption = 'MAST 식이성 62종' then                 
     begin
        pnlName1.Caption  := '돼지고기';              pnlName22.Caption := '대구';             pnlName43.Caption := '게'; 
        pnlName2.Caption  := '소고기';                pnlName23.Caption := '홍합';             pnlName44.Caption := '새우'; 
        pnlName3.Caption  := '치즈';                  pnlName24.Caption := '참치';             pnlName45.Caption := '감자'; 
        pnlName4.Caption  := '닭고기';                pnlName25.Caption := '연어';             pnlName46.Caption := '사과'; 
        pnlName5.Caption  := '번데기';                pnlName26.Caption := '조개';             pnlName47.Caption := '카카오'; 
        pnlName6.Caption  := '토마토';                pnlName27.Caption := '오징어';           pnlName48.Caption := '복숭아';
        pnlName7.Caption  := '키위';                  pnlName28.Caption := '멸치';             pnlName49.Caption := '고등어'; 
        pnlName8.Caption  := '망고';                  pnlName29.Caption := '효모';             pnlName50.Caption := 'Bromelain'; 
        pnlName9.Caption  := '바나나';                pnlName30.Caption := '버섯';             pnlName51.Caption := '호밀풀'; 
        pnlName10.Caption := '감귤류';                pnlName31.Caption := '칸디다곰팡이';     pnlName52.Caption := '집먼지'; 
        pnlName11.Caption := '땅콩';                  pnlName32.Caption := '총 IgE';           pnlName53.Caption := '바퀴벌레'; 
        pnlName12.Caption := '호두';                  pnlName33.Caption := '진드기(Dp)';       pnlName54.Caption := '곰팡이류'; 
        pnlName13.Caption := '밤';                    pnlName34.Caption := '진드기(Df)';       pnlName55.Caption := '곰팡이류'; 
        pnlName14.Caption := '밀가루';                pnlName35.Caption := '저장진드기';       pnlName56.Caption := '곰팡이류'; 
        pnlName15.Caption := '보리';                  pnlName36.Caption := '고양이';           pnlName57.Caption := '오리나무'; 
        pnlName16.Caption := '쌀';                    pnlName37.Caption := '개';               pnlName58.Caption := '자작나무'; 
        pnlName17.Caption := '메밀';                  pnlName38.Caption := '계란흰자';         pnlName59.Caption := '참나무'; 
        pnlName18.Caption := '마늘';                  pnlName39.Caption := '우유';             pnlName60.Caption := '돼지풀';
        pnlName19.Caption := '양파';                  pnlName40.Caption := '옥수수';           pnlName61.Caption := '쑥(머그워트)';
        pnlName20.Caption := '셀러리';                pnlName41.Caption := '참깨';             pnlName62.Caption := '환삼덩굴'; 
        pnlName21.Caption := '오이';                  pnlName42.Caption := '콩';
     end;
  end
  else if Rd_Eng.Checked then
  begin
     if Panel_MAST.Caption = 'MAST 흡입성 62종' then
     begin
        pnlName1.Caption  := 'Acarus siro';                  pnlName22.Caption := 'Pine';                   pnlName43.Caption := 'Crab';
        pnlName2.Caption  := 'Horse';                        pnlName23.Caption := 'Japanese cedar';         pnlName44.Caption := 'Shrimp';
        pnlName3.Caption  := 'Guinea pig';                   pnlName24.Caption := 'Acacia';                 pnlName45.Caption := 'Potato';
        pnlName4.Caption  := 'Sheep';                        pnlName25.Caption := 'Hinoki cypress';         pnlName46.Caption := 'Apple';
        pnlName5.Caption  := 'Rabbit';                       pnlName26.Caption := 'Oxeye daisy';            pnlName47.Caption := 'Cacao';
        pnlName6.Caption  := 'Hamster';                      pnlName27.Caption := 'Dandelion';              pnlName48.Caption := 'Peach';
        pnlName7.Caption  := 'Hazel';                        pnlName28.Caption := 'English plantain';       pnlName49.Caption := 'Mackerel';
        pnlName8.Caption  := 'Sweet vernal grass';           pnlName29.Caption := 'Russian thistle';        pnlName50.Caption := 'CCD';
        pnlName9.Caption  := 'Bermuda grass';                pnlName30.Caption := 'Goldenrod';              pnlName51.Caption := 'Rye pollens';
        pnlName10.Caption := 'Orchard grass';                pnlName31.Caption := 'Pigweed';                pnlName52.Caption := 'House dust';
        pnlName11.Caption := 'Timothy grass';                pnlName32.Caption := 'Total IgE';              pnlName53.Caption := 'Cockroach';
        pnlName12.Caption := 'Reed';                         pnlName33.Caption := 'D. pteronyssinus';       pnlName54.Caption := 'Cladosporium herbarum';
        pnlName13.Caption := 'Redtop, bent grass';           pnlName34.Caption := 'D. farinae';             pnlName55.Caption := 'Aspergillus fumigatus';
        pnlName14.Caption := 'Honey bee';                    pnlName35.Caption := 'Storage mite';           pnlName56.Caption := 'Alternaria alternata';
        pnlName15.Caption := 'Yellow jacket';                pnlName36.Caption := 'Cat';                    pnlName57.Caption := 'Alder';
        pnlName16.Caption := 'Latex';                        pnlName37.Caption := 'Dog';                    pnlName58.Caption := 'Birch';
        pnlName17.Caption := 'Penicillium notatum';          pnlName38.Caption := 'Egg white';              pnlName59.Caption := 'Oak white';
        pnlName18.Caption := 'Sycamore mix';                 pnlName39.Caption := 'Milk';                   pnlName60.Caption := 'Ragweed, short';
        pnlName19.Caption := 'Sallow willow';                pnlName40.Caption := 'Maize';                  pnlName61.Caption := 'Mugwort';
        pnlName20.Caption := 'Poplar mix';                   pnlName41.Caption := 'Sesame';                 pnlName62.Caption := 'Japanese hop';
        pnlName21.Caption := 'Ash mix';                      pnlName42.Caption := 'Soy bean';
     end
     else if Panel_MAST.Caption = 'MAST 식이성 62종' then
     begin
        pnlName1.Caption  := 'Pork';                         pnlName22.Caption := 'Codfish';                pnlName43.Caption := 'Crab';
        pnlName2.Caption  := 'Beef';                         pnlName23.Caption := 'Mussel';                 pnlName44.Caption := 'Shrimp';
        pnlName3.Caption  := 'Cheddar cheese';               pnlName24.Caption := 'Tuna';                   pnlName45.Caption := 'Potato';
        pnlName4.Caption  := 'Chicken';                      pnlName25.Caption := 'Salmon';                 pnlName46.Caption := 'Apple';
        pnlName5.Caption  := 'Pupa, silk cocoon';            pnlName26.Caption := 'Clam';                   pnlName47.Caption := 'Cacao';
        pnlName6.Caption  := 'Tomato';                       pnlName27.Caption := 'Squid';                  pnlName48.Caption := 'Peach';
        pnlName7.Caption  := 'Kiwi';                         pnlName28.Caption := 'Anchovy';                pnlName49.Caption := 'Mackerel';
        pnlName8.Caption  := 'Mango';                        pnlName29.Caption := 'Yeast,bakers';           pnlName50.Caption := 'CCD';
        pnlName9.Caption  := 'Banana';                       pnlName30.Caption := 'Mushroom';               pnlName51.Caption := 'Rye pollens';
        pnlName10.Caption := 'Citrus mix';                   pnlName31.Caption := 'Candida albicans';       pnlName52.Caption := 'House dust';
        pnlName11.Caption := 'Peanut';                       pnlName32.Caption := 'Total IgE';              pnlName53.Caption := 'Cockroach';
        pnlName12.Caption := 'Walnut';                       pnlName33.Caption := 'D. pteronyssinus';       pnlName54.Caption := 'Cladosporium herbarum';
        pnlName13.Caption := 'Chestnut';                     pnlName34.Caption := 'D. farinae';             pnlName55.Caption := 'Aspergillus fumigatus';
        pnlName14.Caption := 'Wheat flour';                  pnlName35.Caption := 'Storage mite';           pnlName56.Caption := 'Alternaria alternata';
        pnlName15.Caption := 'Barley meal';                  pnlName36.Caption := 'Cat';                    pnlName57.Caption := 'Alder';
        pnlName16.Caption := 'Rice';                         pnlName37.Caption := 'Dog';                    pnlName58.Caption := 'Birch';
        pnlName17.Caption := 'Buck-wheat';                   pnlName38.Caption := 'Egg white';              pnlName59.Caption := 'Oak white';
        pnlName18.Caption := 'Garlic';                       pnlName39.Caption := 'Milk';                   pnlName60.Caption := 'Ragweed, short';
        pnlName19.Caption := 'Onion';                        pnlName40.Caption := 'Maize';                  pnlName61.Caption := 'Mugwort';
        pnlName20.Caption := 'Celery';                       pnlName41.Caption := 'Sesame';                 pnlName62.Caption := 'Japanese hop';
        pnlName21.Caption := 'Cucumber';                     pnlName42.Caption := 'Soy bean';
     end;
  end;
end;

procedure TfrmLD1I18.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;                    
                                 
   if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;                             
                                 
procedure TfrmLD1I18.btnPSaveClick(Sender: TObject);
var i : Integer;                 
begin                            
  inherited;                     
                                 
   // 현재위치 추출              
   i := grdMaster.CurrentDataRow - 1;
                                 
   // DB 작업
   dmComm.database.StartTransaction;
   try
      with qryU_Allergy do
      begin
         Active := False;
         ParamByName('cod_bjjs').AsString  := UV_vCod_bjjs[i];
         ParamByName('num_id').AsString    := UV_vnum_id[i];
         ParamByName('dat_bunju').AsString := mskDAT_BUNJU.Text;
         ParamByName('cod_hm').AsString    := UV_vcod_hm[i];
         ParamByName('Alle_01').AsString   := edtValue1.Text;         ParamByName('Alle_01_class').AsString   := Edt_Class1.Text;
         ParamByName('Alle_02').AsString   := edtValue2.Text;         ParamByName('Alle_02_class').AsString   := Edt_Class2.Text;
         ParamByName('Alle_03').AsString   := edtValue3.Text;         ParamByName('Alle_03_class').AsString   := Edt_Class3.Text;
         ParamByName('Alle_04').AsString   := edtValue4.Text;         ParamByName('Alle_04_class').AsString   := Edt_Class4.Text;
         ParamByName('Alle_05').AsString   := edtValue5.Text;         ParamByName('Alle_05_class').AsString   := Edt_Class5.Text;
         ParamByName('Alle_06').AsString   := edtValue6.Text;         ParamByName('Alle_06_class').AsString   := Edt_Class6.Text;
         ParamByName('Alle_07').AsString   := edtValue7.Text;         ParamByName('Alle_07_class').AsString   := Edt_Class7.Text;
         ParamByName('Alle_08').AsString   := edtValue8.Text;         ParamByName('Alle_08_class').AsString   := Edt_Class8.Text;
         ParamByName('Alle_09').AsString   := edtValue9.Text;         ParamByName('Alle_09_class').AsString   := Edt_Class9.Text;
         ParamByName('Alle_10').AsString   := edtValue10.Text;        ParamByName('Alle_10_class').AsString   := Edt_Class10.Text;
         ParamByName('Alle_11').AsString   := edtValue11.Text;        ParamByName('Alle_11_class').AsString   := Edt_Class11.Text;
         ParamByName('Alle_12').AsString   := edtValue12.Text;        ParamByName('Alle_12_class').AsString   := Edt_Class12.Text;
         ParamByName('Alle_13').AsString   := edtValue13.Text;        ParamByName('Alle_13_class').AsString   := Edt_Class13.Text;
         ParamByName('Alle_14').AsString   := edtValue14.Text;        ParamByName('Alle_14_class').AsString   := Edt_Class14.Text;
         ParamByName('Alle_15').AsString   := edtValue15.Text;        ParamByName('Alle_15_class').AsString   := Edt_Class15.Text;
         ParamByName('Alle_16').AsString   := edtValue16.Text;        ParamByName('Alle_16_class').AsString   := Edt_Class16.Text;
         ParamByName('Alle_17').AsString   := edtValue17.Text;        ParamByName('Alle_17_class').AsString   := Edt_Class17.Text;
         ParamByName('Alle_18').AsString   := edtValue18.Text;        ParamByName('Alle_18_class').AsString   := Edt_Class18.Text;
         ParamByName('Alle_19').AsString   := edtValue19.Text;        ParamByName('Alle_19_class').AsString   := Edt_Class19.Text;
         ParamByName('Alle_20').AsString   := edtValue20.Text;        ParamByName('Alle_20_class').AsString   := Edt_Class20.Text;
         ParamByName('Alle_21').AsString   := edtValue21.Text;        ParamByName('Alle_21_class').AsString   := Edt_Class21.Text;
         ParamByName('Alle_22').AsString   := edtValue22.Text;        ParamByName('Alle_22_class').AsString   := Edt_Class22.Text;
         ParamByName('Alle_23').AsString   := edtValue23.Text;        ParamByName('Alle_23_class').AsString   := Edt_Class23.Text;
         ParamByName('Alle_24').AsString   := edtValue24.Text;        ParamByName('Alle_24_class').AsString   := Edt_Class24.Text;
         ParamByName('Alle_25').AsString   := edtValue25.Text;        ParamByName('Alle_25_class').AsString   := Edt_Class25.Text;
         ParamByName('Alle_26').AsString   := edtValue26.Text;        ParamByName('Alle_26_class').AsString   := Edt_Class26.Text;
         ParamByName('Alle_27').AsString   := edtValue27.Text;        ParamByName('Alle_27_class').AsString   := Edt_Class27.Text;
         ParamByName('Alle_28').AsString   := edtValue28.Text;        ParamByName('Alle_28_class').AsString   := Edt_Class28.Text;
         ParamByName('Alle_29').AsString   := edtValue29.Text;        ParamByName('Alle_29_class').AsString   := Edt_Class29.Text;
         ParamByName('Alle_30').AsString   := edtValue30.Text;        ParamByName('Alle_30_class').AsString   := Edt_Class30.Text;
         ParamByName('Alle_31').AsString   := edtValue31.Text;        ParamByName('Alle_31_class').AsString   := Edt_Class31.Text;
         ParamByName('Alle_32').AsString   := edtValue32.Text;        ParamByName('Alle_32_class').AsString   := Edt_Class32.Text;
         ParamByName('Alle_33').AsString   := edtValue33.Text;        ParamByName('Alle_33_class').AsString   := Edt_Class33.Text;
         ParamByName('Alle_34').AsString   := edtValue34.Text;        ParamByName('Alle_34_class').AsString   := Edt_Class34.Text;
         ParamByName('Alle_35').AsString   := edtValue35.Text;        ParamByName('Alle_35_class').AsString   := Edt_Class35.Text;
         ParamByName('Alle_36').AsString   := edtValue36.Text;        ParamByName('Alle_36_class').AsString   := Edt_Class36.Text;
         ParamByName('Alle_37').AsString   := edtValue37.Text;        ParamByName('Alle_37_class').AsString   := Edt_Class37.Text;
         ParamByName('Alle_38').AsString   := edtValue38.Text;        ParamByName('Alle_38_class').AsString   := Edt_Class38.Text;
         ParamByName('Alle_39').AsString   := edtValue39.Text;        ParamByName('Alle_39_class').AsString   := Edt_Class39.Text;
         ParamByName('Alle_40').AsString   := edtValue40.Text;        ParamByName('Alle_40_class').AsString   := Edt_Class40.Text;
         ParamByName('Alle_41').AsString   := edtValue41.Text;        ParamByName('Alle_41_class').AsString   := Edt_Class41.Text;
         ParamByName('Alle_42').AsString   := edtValue42.Text;        ParamByName('Alle_42_class').AsString   := Edt_Class42.Text;
         ParamByName('Alle_43').AsString   := edtValue43.Text;        ParamByName('Alle_43_class').AsString   := Edt_Class43.Text;
         ParamByName('Alle_44').AsString   := edtValue44.Text;        ParamByName('Alle_44_class').AsString   := Edt_Class44.Text;
         ParamByName('Alle_45').AsString   := edtValue45.Text;        ParamByName('Alle_45_class').AsString   := Edt_Class45.Text;
         ParamByName('Alle_46').AsString   := edtValue46.Text;        ParamByName('Alle_46_class').AsString   := Edt_Class46.Text;
         ParamByName('Alle_47').AsString   := edtValue47.Text;        ParamByName('Alle_47_class').AsString   := Edt_Class47.Text;
         ParamByName('Alle_48').AsString   := edtValue48.Text;        ParamByName('Alle_48_class').AsString   := Edt_Class48.Text;
         ParamByName('Alle_49').AsString   := edtValue49.Text;        ParamByName('Alle_49_class').AsString   := Edt_Class49.Text;
         ParamByName('Alle_50').AsString   := edtValue50.Text;        ParamByName('Alle_50_class').AsString   := Edt_Class50.Text;
         ParamByName('Alle_51').AsString   := edtValue51.Text;        ParamByName('Alle_51_class').AsString   := Edt_Class51.Text;
         ParamByName('Alle_52').AsString   := edtValue52.Text;        ParamByName('Alle_52_class').AsString   := Edt_Class52.Text;
         ParamByName('Alle_53').AsString   := edtValue53.Text;        ParamByName('Alle_53_class').AsString   := Edt_Class53.Text;
         ParamByName('Alle_54').AsString   := edtValue54.Text;        ParamByName('Alle_54_class').AsString   := Edt_Class54.Text;
         ParamByName('Alle_55').AsString   := edtValue55.Text;        ParamByName('Alle_55_class').AsString   := Edt_Class55.Text;
         ParamByName('Alle_56').AsString   := edtValue56.Text;        ParamByName('Alle_56_class').AsString   := Edt_Class56.Text;
         ParamByName('Alle_57').AsString   := edtValue57.Text;        ParamByName('Alle_57_class').AsString   := Edt_Class57.Text;
         ParamByName('Alle_58').AsString   := edtValue58.Text;        ParamByName('Alle_58_class').AsString   := Edt_Class58.Text;
         ParamByName('Alle_59').AsString   := edtValue59.Text;        ParamByName('Alle_59_class').AsString   := Edt_Class59.Text;
         ParamByName('Alle_60').AsString   := edtValue60.Text;        ParamByName('Alle_60_class').AsString   := Edt_Class60.Text;
         ParamByName('Alle_61').AsString   := edtValue61.Text;        ParamByName('Alle_61_class').AsString   := Edt_Class61.Text;
         ParamByName('Alle_62').AsString   := edtValue62.Text;        ParamByName('Alle_62_class').AsString   := Edt_Class62.Text;
         ParamByName('dat_report').AsString := mskDAT_RESULT.Text;
         ParamByName('cod_tester').AsString := Copy(cmbCOD_SAWON.Text,1,6);
         ParamByName('cod_doct').AsString   := Copy(cmbCOD_DOCT.Text,1,6);
         ParamByName('dat_update').AsString := GV_sToday;
         ParamByName('cod_update').AsString := GV_sUserId;
         Execsql;
         GP_query_log(qryU_Allergy, 'LD1I16', 'Q', 'Y');
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

   // Save Mode & Msg
   UP_SetMode('SAVE');

   UP_MoveNum(btnNextNum);

end;

procedure TfrmLD1I18.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate            then GF_CalendarClick(mskDate)
   else if Sender = btnDAT_RESULT then GF_CalendarClick(mskDAT_RESULT);
end;


end.


















































































