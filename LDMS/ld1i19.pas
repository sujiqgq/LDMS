//==============================================================================
// 프로그램 설명 : 알레르기결과등록
// 수정일자      : 2012.04.29
// 수정자        : 김승철
// 참고사항      : 추가.
//==============================================================================

unit LD1I19;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I19 = class(TfrmSingle)
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
    Panel6: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    Panel17: TPanel;
    pnlName63: TPanel;
    edtValue63: TEdit;
    Panel23: TPanel;
    pnlName64: TPanel;
    edtValue64: TEdit;
    Panel27: TPanel;
    pnlName65: TPanel;
    edtValue65: TEdit;
    Edt_Class63: TEdit;
    Edt_Class64: TEdit;
    Edt_Class65: TEdit;
    Edt_Class66: TEdit;
    edtValue66: TEdit;
    pnlName66: TPanel;
    Panel33: TPanel;
    Panel37: TPanel;
    pnlName67: TPanel;
    edtValue67: TEdit;
    Edt_Class67: TEdit;
    Edt_Class68: TEdit;
    edtValue68: TEdit;
    pnlName68: TPanel;
    Panel40: TPanel;
    Panel41: TPanel;
    pnlName69: TPanel;
    edtValue69: TEdit;
    Edt_Class69: TEdit;
    pnlName70: TPanel;
    Panel26: TPanel;
    Panel28: TPanel;
    pnlName71: TPanel;
    edtValue71: TEdit;
    Panel38: TPanel;
    pnlName72: TPanel;
    edtValue72: TEdit;
    Panel42: TPanel;
    pnlName73: TPanel;
    edtValue73: TEdit;
    Panel51: TPanel;
    pnlName74: TPanel;
    edtValue74: TEdit;
    Panel53: TPanel;
    pnlName75: TPanel;
    edtValue75: TEdit;
    Panel55: TPanel;
    pnlName76: TPanel;
    edtValue76: TEdit;
    Panel57: TPanel;
    pnlName77: TPanel;
    edtValue77: TEdit;
    Panel59: TPanel;
    pnlName78: TPanel;
    edtValue78: TEdit;
    Panel61: TPanel;
    pnlName79: TPanel;
    edtValue79: TEdit;
    Panel63: TPanel;
    pnlName80: TPanel;
    edtValue80: TEdit;
    Panel65: TPanel;
    pnlName81: TPanel;
    edtValue81: TEdit;
    Panel67: TPanel;
    pnlName82: TPanel;
    edtValue82: TEdit;
    Panel69: TPanel;
    pnlName83: TPanel;
    edtValue83: TEdit;
    Panel71: TPanel;
    pnlName84: TPanel;
    edtValue84: TEdit;
    pnlName85: TPanel;
    Panel74: TPanel;
    Panel75: TPanel;
    pnlName86: TPanel;
    edtValue86: TEdit;
    Panel77: TPanel;
    pnlName87: TPanel;
    edtValue87: TEdit;
    Panel79: TPanel;
    pnlName88: TPanel;
    edtValue88: TEdit;
    Panel81: TPanel;
    pnlName89: TPanel;
    edtValue89: TEdit;
    Panel83: TPanel;
    pnlName90: TPanel;
    edtValue90: TEdit;
    edtValue70: TEdit;
    edtValue85: TEdit;
    Edt_Class71: TEdit;
    Edt_Class72: TEdit;
    Edt_Class73: TEdit;
    Edt_Class74: TEdit;
    Edt_Class75: TEdit;
    Edt_Class76: TEdit;
    Edt_Class77: TEdit;
    Edt_Class78: TEdit;
    Edt_Class79: TEdit;
    Edt_Class80: TEdit;
    Edt_Class81: TEdit;
    Edt_Class82: TEdit;
    Edt_Class83: TEdit;
    Edt_Class84: TEdit;
    Edt_Class85: TEdit;
    Edt_Class86: TEdit;
    Edt_Class87: TEdit;
    Edt_Class88: TEdit;
    Edt_Class89: TEdit;
    Edt_Class90: TEdit;
    Edt_Class70: TEdit;
    Panel86: TPanel;
    pnlName91: TPanel;
    edtValue91: TEdit;
    Edt_Class91: TEdit;
    Panel22: TPanel;

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

  const UC_First = 1;
        UC_Last  = 2;
var
  frmLD1I19: TfrmLD1I19;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD1I19.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD1I19.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I19.UP_GridSet;
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
      end
      else if rbJumin.Checked = True then
      begin
         Col[1].Heading := '검진일자';
         Col[2].Heading := '샘플번호';
         Col[3].Heading := '성명';
      end;
   end;
end;

procedure TfrmLD1I19.btnQueryClick(Sender: TObject);
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
      sSelect := sSelect + ' FROM Allergy_107 A INNER JOIN gumgin B ';
      sSelect := sSelect + ' ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ';

      sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';
      sWhere := sWhere + ' AND A.cod_hm = ''S104'' ';

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

function TfrmLD1I19.UF_Condition : Boolean;
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


procedure TfrmLD1I19.FormCreate(Sender: TObject);
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
   Rd_EngClick(Rd_Kor);
   Cmb_gubn.ItemIndex := 0;
   cmbRelation.ItemIndex := 0;
end;

procedure TfrmLD1I19.UP_VarMemSet(iLength : Integer);
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

procedure TfrmLD1I19.grdMasterCellLoaded(Sender: TObject; DataCol,
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
   end;
end;

procedure TfrmLD1I19.rbDateClick(Sender: TObject);
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

procedure TfrmLD1I19.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var i : integer;
    a : String;
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

   if Rd_Kor.checked then Rd_EngClick(Rd_Kor);

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
         edtValue63.Text := FieldByName('Alle_63').AsString;
         Edt_Class63.Text := FieldByName('Alle_63_Class').AsString;
         edtValue64.Text := FieldByName('Alle_64').AsString;
         Edt_Class64.Text := FieldByName('Alle_64_Class').AsString;
         edtValue65.Text := FieldByName('Alle_65').AsString;
         Edt_Class65.Text := FieldByName('Alle_65_Class').AsString;
         edtValue66.Text := FieldByName('Alle_66').AsString;
         Edt_Class66.Text := FieldByName('Alle_66_Class').AsString;
         edtValue67.Text := FieldByName('Alle_67').AsString;
         Edt_Class67.Text := FieldByName('Alle_67_Class').AsString;
         edtValue68.Text := FieldByName('Alle_68').AsString;
         Edt_Class68.Text := FieldByName('Alle_68_Class').AsString;
         edtValue69.Text := FieldByName('Alle_69').AsString;
         Edt_Class69.Text := FieldByName('Alle_69_Class').AsString;
         edtValue70.Text := FieldByName('Alle_70').AsString;
         Edt_Class70.Text := FieldByName('Alle_70_Class').AsString;
         edtValue71.Text := FieldByName('Alle_71').AsString;
         Edt_Class71.Text := FieldByName('Alle_71_Class').AsString;
         edtValue72.Text := FieldByName('Alle_72').AsString;
         Edt_Class72.Text := FieldByName('Alle_72_Class').AsString;
         edtValue73.Text := FieldByName('Alle_73').AsString;
         Edt_Class73.Text := FieldByName('Alle_73_Class').AsString;
         edtValue74.Text := FieldByName('Alle_74').AsString;
         Edt_Class74.Text := FieldByName('Alle_74_Class').AsString;
         edtValue75.Text := FieldByName('Alle_75').AsString;
         Edt_Class75.Text := FieldByName('Alle_75_Class').AsString;
         edtValue76.Text := FieldByName('Alle_76').AsString;
         Edt_Class76.Text := FieldByName('Alle_76_Class').AsString;
         edtValue77.Text := FieldByName('Alle_77').AsString;
         Edt_Class77.Text := FieldByName('Alle_77_Class').AsString;
         edtValue78.Text := FieldByName('Alle_78').AsString;
         Edt_Class78.Text := FieldByName('Alle_78_Class').AsString;
         edtValue79.Text := FieldByName('Alle_79').AsString;
         Edt_Class79.Text := FieldByName('Alle_79_Class').AsString;
         edtValue80.Text := FieldByName('Alle_80').AsString;
         Edt_Class80.Text := FieldByName('Alle_80_Class').AsString;
         edtValue81.Text := FieldByName('Alle_81').AsString;
         Edt_Class81.Text := FieldByName('Alle_81_Class').AsString;
         edtValue82.Text := FieldByName('Alle_82').AsString;
         Edt_Class82.Text := FieldByName('Alle_82_Class').AsString;
         edtValue83.Text := FieldByName('Alle_83').AsString;
         Edt_Class83.Text := FieldByName('Alle_83_Class').AsString;
         edtValue84.Text := FieldByName('Alle_84').AsString;
         Edt_Class84.Text := FieldByName('Alle_84_Class').AsString;
         edtValue85.Text := FieldByName('Alle_85').AsString;
         Edt_Class85.Text := FieldByName('Alle_85_Class').AsString;
         edtValue86.Text := FieldByName('Alle_86').AsString;
         Edt_Class86.Text := FieldByName('Alle_86_Class').AsString;
         edtValue87.Text := FieldByName('Alle_87').AsString;
         Edt_Class87.Text := FieldByName('Alle_87_Class').AsString;
         edtValue88.Text := FieldByName('Alle_88').AsString;
         Edt_Class88.Text := FieldByName('Alle_88_Class').AsString;
         edtValue89.Text := FieldByName('Alle_89').AsString;
         Edt_Class89.Text := FieldByName('Alle_89_Class').AsString;
         edtValue90.Text := FieldByName('Alle_90').AsString;
         Edt_Class90.Text := FieldByName('Alle_90_Class').AsString;
         edtValue91.Text := FieldByName('Alle_91').AsString;
         Edt_Class91.Text := FieldByName('Alle_91_Class').AsString;
      end;
      Active := False;
   end;

   if Edt_Class1.Text =  'P' then
      Edt_Class1.Color := $00FACDF4
   else
      edtValue1.Color := clWindow;

   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if Pos('edtValue', TEdit(pnlDetail.Controls[i]).Name) > 0 then
      begin
         if ((TEdit(pnlDetail.Controls[i]).Name) <> 'edtValue1') then
         begin
            if (TEdit(pnlDetail.Controls[i]).text) = '≥100.0' then  TEdit(pnlDetail.Controls[i]).Color  := $00FACDF4
            else if  (TEdit(pnlDetail.Controls[i]).text) = '<0.15' then  TEdit(pnlDetail.Controls[i]).Color  := clWindow
            else if (strTofloat(TEdit(pnlDetail.Controls[i]).text) >= 0.70) then
            begin
               TEdit(pnlDetail.Controls[i]).Color  := $00FACDF4;
            end
            else
               TEdit(pnlDetail.Controls[i]).Color  := clWindow;
         end;
      end;
   end;

end;

procedure TfrmLD1I19.Rd_EngClick(Sender: TObject);
begin
  inherited;

  if Rd_Kor.Checked then
  begin
     pnlName1.Caption  := '총IgE';
     pnlName2.Caption  := '집먼지';
     pnlName3.Caption  := '진드기(Dp)';
     pnlName4.Caption  := '진드기(Df)';
     pnlName5.Caption  := '저장진드기(a)';
     pnlName6.Caption  := '저장진드기(T)';
     pnlName7.Caption  := '페니실리움';
     pnlName8.Caption  := '클라도스포리움';
     pnlName9.Caption  := '아스퍼질러스';
     pnlName10.Caption := '칸디다';
     pnlName11.Caption := '알터나리아';
     pnlName12.Caption := '효모';
     pnlName13.Caption := '고양이털';
     pnlName14.Caption := '말';
     pnlName15.Caption := '개털';
     pnlName16.Caption := '기니피그';
     pnlName17.Caption := '양모';
     pnlName18.Caption := '토끼';
     pnlName19.Caption := '햄스터';
     pnlName20.Caption := '생쥐/쥐';
     pnlName21.Caption := '계란흰자';
     pnlName22.Caption := '닭고기';
     pnlName23.Caption := '우유';
     pnlName24.Caption := '치즈';
     pnlName25.Caption := '밀가루';
     pnlName26.Caption := '보리';
     pnlName27.Caption := '옥수수';
     pnlName28.Caption := '쌀';
     pnlName29.Caption := '참깨';
     pnlName30.Caption := '메밀';
     pnlName31.Caption := '땅콩';
     pnlName32.Caption := '콩';
     pnlName33.Caption := '헤이즐넛';
     pnlName34.Caption := '카카오';
     pnlName35.Caption := '호두';
     pnlName36.Caption := '밤';
     pnlName37.Caption := '아몬드/잣/해바라기씨';
     pnlName38.Caption := '대구';
     pnlName39.Caption := '게';
     pnlName40.Caption := '새우';
     pnlName41.Caption := '고등어';
     pnlName42.Caption := '장어';
     pnlName43.Caption := '홍합/굴/조개/가리비';
     pnlName44.Caption := '참치/연어';
     pnlName45.Caption := '랍스터/오징어';
     pnlName46.Caption := '가자미/멸치/명태 ';
     pnlName47.Caption := '토마토';
     pnlName48.Caption := '당근';
     pnlName49.Caption := '감귤류';
     pnlName50.Caption := '감자';
     pnlName51.Caption := '딸기';
     pnlName52.Caption := '사과';
     pnlName53.Caption := '샐러리';
     pnlName54.Caption := '복숭아';
     pnlName55.Caption := '오이';
     pnlName56.Caption := '마늘/양파';
     pnlName57.Caption := '키위/망고/바나나';
     pnlName58.Caption := '돼지고기';
     pnlName59.Caption := '소고기';
     pnlName60.Caption := '양고기';
     pnlName61.Caption := '오리나무';
     pnlName62.Caption := '자작나무';
     pnlName63.Caption := '개암나무';
     pnlName64.Caption := '참나무';
     pnlName65.Caption := '올리브';
     pnlName66.Caption := '플라터너스';
     pnlName67.Caption := '버드나무';
     pnlName68.Caption := '미루나무';
     pnlName69.Caption := '물푸레나무';
     pnlName70.Caption := '백송';
     pnlName71.Caption := '삼나무';
     pnlName72.Caption := '아카시아';
     pnlName73.Caption := '돼지풀';
     pnlName74.Caption := '쑥 꽃가루';
     pnlName75.Caption := '블란서국화';
     pnlName76.Caption := '민들레';
     pnlName77.Caption := '참질경이';
     pnlName78.Caption := '명아주과풀';
     pnlName79.Caption := '미역취국화';
     pnlName80.Caption := '털비름';
     pnlName81.Caption := '환삼덩굴';
     pnlName82.Caption := '우산잔디';
     pnlName83.Caption := '큰조아재비';
     pnlName84.Caption := '호밀꽃가루';
     pnlName85.Caption := '향기풀/오리새/';
     Panel22.Caption := '갈대/외겨이삭';
     pnlName86.Caption := '꿀벌독';
     pnlName87.Caption := '말벌독';
     pnlName88.Caption := '바퀴벌레';
     pnlName89.Caption := '라텍스';
     pnlName90.Caption := '번데기';
     pnlName91.Caption := 'CCD항원';

  end
  else if Rd_Eng.Checked then
  begin
     pnlName1.Caption  := 'Total IgE';
     pnlName2.Caption  := 'House dust';
     pnlName3.Caption  := 'D. pteronyssinus';
     pnlName4.Caption  := 'D. farinae';
     pnlName5.Caption  := 'Acarus siro';
     pnlName6.Caption  := 'Tyrophagus putrescentiae';
     pnlName7.Caption  := 'Penicillium notatum';
     pnlName8.Caption  := 'Cladosporium herbaum';
     pnlName9.Caption  := 'Aspergillus fumigatus';
     pnlName10.Caption := 'Candida albicans';
     pnlName11.Caption := 'Alternaria alterrnata';
     pnlName12.Caption := 'Yeast,baker^s';
     pnlName13.Caption := 'Cat epithelium & dander';
     pnlName14.Caption := 'Horse';
     pnlName15.Caption := 'Dog dander';
     pnlName16.Caption := 'Guniea pig';
     pnlName17.Caption := 'Wool, sheep';
     pnlName18.Caption := 'Rabbit';
     pnlName19.Caption := 'Hamster';
     pnlName20.Caption := 'Mouse/Rat';
     pnlName21.Caption := 'Egg white';
     pnlName22.Caption := 'Chicken';
     pnlName23.Caption := 'Milk';
     pnlName24.Caption := 'Chicken';
     pnlName25.Caption := 'wheat';
     pnlName26.Caption := 'Barley';
     pnlName27.Caption := 'Corn';
     pnlName28.Caption := 'Rice';
     pnlName29.Caption := 'Sesame';
     pnlName30.Caption := 'Buckwheat';
     pnlName31.Caption := 'Peanut';
     pnlName32.Caption := 'Soy bean';
     pnlName33.Caption := 'Hazel nut';
     pnlName34.Caption := 'Cacao';
     pnlName35.Caption := 'Walnut';
     pnlName36.Caption := 'Sweet chestnut';
     pnlName37.Caption := 'Almond/Pine nut/Sunflower';
     pnlName38.Caption := 'Codfish';
     pnlName39.Caption := 'Crab';
     pnlName40.Caption := 'Shrimp';
     pnlName41.Caption := 'Mackerel';
     pnlName42.Caption := 'Eel';
     pnlName43.Caption := 'Blue mussel/Oyster/Clam/Scallop';
     pnlName44.Caption := 'Tuna/Salmon';
     pnlName45.Caption := 'Lobster/Pacific squid';
     pnlName46.Caption := 'Plaice/Anchovy/Alaska pollock';
     pnlName47.Caption := 'Tomato';
     pnlName48.Caption := 'Carrot';
     pnlName49.Caption := 'Citrus mix';
     pnlName50.Caption := 'Potato';
     pnlName51.Caption := 'Strawberry';
     pnlName52.Caption := 'Apple';
     pnlName53.Caption := 'Celery';
     pnlName54.Caption := 'Peach';
     pnlName55.Caption := 'Cucumber';
     pnlName56.Caption := 'Galic/Onion';
     pnlName57.Caption := 'Kiwi/Mango/Banana';
     pnlName58.Caption := 'Pork';
     pnlName59.Caption := 'Beef';
     pnlName60.Caption := 'Lamb meat';
     pnlName61.Caption := 'Alder';
     pnlName62.Caption := 'Birch';
     pnlName63.Caption := 'Hazel';
     pnlName64.Caption := 'Oak';
     pnlName65.Caption := 'Olive';
     pnlName66.Caption := 'Maple leaf sycamore';
     pnlName67.Caption := 'Willow';
     pnlName68.Caption := 'Cottonwood';
     pnlName69.Caption := 'White ash';
     pnlName70.Caption := 'White pine';
     pnlName71.Caption := 'Japanese cedar';
     pnlName72.Caption := 'Acacia';
     pnlName73.Caption := 'Commom ragweed';
     pnlName74.Caption := 'Mugwort';
     pnlName75.Caption := 'Ox-eye daisy';
     pnlName76.Caption := 'Dandelion';
     pnlName77.Caption := 'Plantain';
     pnlName78.Caption := 'Russian thistle';
     pnlName79.Caption := 'Goldenrod';
     pnlName80.Caption := 'Commom pigweed';
     pnlName81.Caption := 'Japanese hop';
     pnlName82.Caption := 'Bermuda grass';
     pnlName83.Caption := 'Timothy grass';
     pnlName84.Caption := 'Cultivated rye';
     pnlName85.Caption := 'Sweet vernal grass/Orchard grass/Common reed/Bent grass';
     pnlName86.Caption := 'Bee venom';
     pnlName87.Caption := 'Wasp venom';
     pnlName88.Caption := 'Cockroach';
     pnlName89.Caption := 'Hevea latex';
     pnlName90.Caption := 'Silkworm, pupa';
     pnlName91.Caption := 'CCD';
  end;

end;

procedure TfrmLD1I19.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;
                                 
procedure TfrmLD1I19.btnPSaveClick(Sender: TObject);
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
         ParamByName('Alle_63').AsString   := edtValue63.Text;        ParamByName('Alle_63_class').AsString   := Edt_Class63.Text;
         ParamByName('Alle_64').AsString   := edtValue64.Text;        ParamByName('Alle_64_class').AsString   := Edt_Class64.Text;
         ParamByName('Alle_65').AsString   := edtValue65.Text;        ParamByName('Alle_65_class').AsString   := Edt_Class65.Text;
         ParamByName('Alle_66').AsString   := edtValue66.Text;        ParamByName('Alle_66_class').AsString   := Edt_Class66.Text;
         ParamByName('Alle_67').AsString   := edtValue67.Text;        ParamByName('Alle_67_class').AsString   := Edt_Class67.Text;
         ParamByName('Alle_68').AsString   := edtValue68.Text;        ParamByName('Alle_68_class').AsString   := Edt_Class68.Text;
         ParamByName('Alle_69').AsString   := edtValue69.Text;        ParamByName('Alle_69_class').AsString   := Edt_Class69.Text;
         ParamByName('Alle_70').AsString   := edtValue70.Text;        ParamByName('Alle_70_class').AsString   := Edt_Class70.Text;
         ParamByName('Alle_71').AsString   := edtValue71.Text;        ParamByName('Alle_71_class').AsString   := Edt_Class71.Text;
         ParamByName('Alle_72').AsString   := edtValue72.Text;        ParamByName('Alle_72_class').AsString   := Edt_Class72.Text;
         ParamByName('Alle_73').AsString   := edtValue73.Text;        ParamByName('Alle_73_class').AsString   := Edt_Class73.Text;
         ParamByName('Alle_74').AsString   := edtValue74.Text;        ParamByName('Alle_74_class').AsString   := Edt_Class74.Text;
         ParamByName('Alle_75').AsString   := edtValue75.Text;        ParamByName('Alle_75_class').AsString   := Edt_Class75.Text;
         ParamByName('Alle_76').AsString   := edtValue76.Text;        ParamByName('Alle_76_class').AsString   := Edt_Class76.Text;
         ParamByName('Alle_77').AsString   := edtValue77.Text;        ParamByName('Alle_77_class').AsString   := Edt_Class77.Text;
         ParamByName('Alle_78').AsString   := edtValue78.Text;        ParamByName('Alle_78_class').AsString   := Edt_Class78.Text;
         ParamByName('Alle_79').AsString   := edtValue79.Text;        ParamByName('Alle_79_class').AsString   := Edt_Class79.Text;
         ParamByName('Alle_80').AsString   := edtValue80.Text;        ParamByName('Alle_80_class').AsString   := Edt_Class80.Text;
         ParamByName('Alle_81').AsString   := edtValue81.Text;        ParamByName('Alle_81_class').AsString   := Edt_Class81.Text;
         ParamByName('Alle_82').AsString   := edtValue82.Text;        ParamByName('Alle_82_class').AsString   := Edt_Class82.Text;
         ParamByName('Alle_83').AsString   := edtValue83.Text;        ParamByName('Alle_83_class').AsString   := Edt_Class83.Text;
         ParamByName('Alle_84').AsString   := edtValue84.Text;        ParamByName('Alle_84_class').AsString   := Edt_Class84.Text;
         ParamByName('Alle_85').AsString   := edtValue85.Text;        ParamByName('Alle_85_class').AsString   := Edt_Class85.Text;
         ParamByName('Alle_86').AsString   := edtValue86.Text;        ParamByName('Alle_86_class').AsString   := Edt_Class86.Text;
         ParamByName('Alle_87').AsString   := edtValue87.Text;        ParamByName('Alle_87_class').AsString   := Edt_Class87.Text;
         ParamByName('Alle_88').AsString   := edtValue88.Text;        ParamByName('Alle_88_class').AsString   := Edt_Class88.Text;
         ParamByName('Alle_89').AsString   := edtValue89.Text;        ParamByName('Alle_89_class').AsString   := Edt_Class89.Text;
         ParamByName('Alle_90').AsString   := edtValue90.Text;        ParamByName('Alle_90_class').AsString   := Edt_Class90.Text;
         ParamByName('Alle_91').AsString   := edtValue91.Text;        ParamByName('Alle_91_class').AsString   := Edt_Class91.Text;

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

procedure TfrmLD1I19.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate            then GF_CalendarClick(mskDate)
   else if Sender = btnDAT_RESULT then GF_CalendarClick(mskDAT_RESULT);
end;

end.


















































































