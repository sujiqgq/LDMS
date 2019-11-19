//==============================================================================
// 프로그램 설명 : 알레르기결과등록
// 수정일자      : 2012.04.29
// 수정자        : 김승철
// 참고사항      : 추가.
//==============================================================================

unit LD1I16;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I16 = class(TfrmSingle)
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
    edtValue32: TEdit;
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
    edtValue1: TEdit;
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
    Edt_Class1: TEdit;
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
    Edt_Class32: TEdit;
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
  frmLD1I16: TfrmLD1I16;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD1I16.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD1I16.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I16.UP_GridSet;
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

procedure TfrmLD1I16.btnQueryClick(Sender: TObject);
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
      sSelect := sSelect + ' FROM Allergy A INNER JOIN gumgin B ';
      sSelect := sSelect + ' ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn ';

      sWhere := 'WHERE A.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + '''';

      if cmbRelation.ItemIndex = 1 then
         sWhere := sWhere + ' AND A.cod_hm = ''S079'' '
      else if cmbRelation.ItemIndex = 2 then
         sWhere := sWhere + ' AND A.cod_hm = ''S090'' ';

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

function TfrmLD1I16.UF_Condition : Boolean;
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


procedure TfrmLD1I16.FormCreate(Sender: TObject);
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

procedure TfrmLD1I16.UP_VarMemSet(iLength : Integer);
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

procedure TfrmLD1I16.grdMasterCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD1I16.rbDateClick(Sender: TObject);
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

procedure TfrmLD1I16.grdMasterRowChanged(Sender: TObject; OldRow,
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

   if UV_vCod_hm[NewRow-1] = 'S079' then
   begin
      Panel_MAST.Caption := 'MAST 흡입성 42종';
      Panel_MAST.Font.Color := clblue;
      Rd_EngClick(Rd_Kor);
   end
   else if UV_vCod_hm[NewRow-1] = 'S090' then
   begin
      Panel_MAST.Caption := 'MAST 식이성 42종';
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
      end;
      Active := False;
   end;

   if edtValue1.Text = '>100' then
      edtValue1.Color := $00FACDF4
   else
      edtValue1.Color := clWindow;

   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if Pos('edtValue', TEdit(pnlDetail.Controls[i]).Name) > 0 then
      begin
         if ((TEdit(pnlDetail.Controls[i]).Name) <> 'edtValue1') then
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

procedure TfrmLD1I16.Rd_EngClick(Sender: TObject);
begin
  inherited;

  if Rd_Kor.Checked then
  begin
     if Panel_MAST.Caption = 'MAST 흡입성 42종' then
     begin
        pnlName1.Caption  := '총 IgE';             pnlName22.Caption := '털비름';
        pnlName2.Caption  := '콩';                 pnlName23.Caption := '명아주과풀';
        pnlName3.Caption  := '우유';               pnlName24.Caption := '민들레';
        pnlName4.Caption  := '계란흰자';           pnlName25.Caption := '쑥';
        pnlName5.Caption  := '게';                 pnlName26.Caption := '돼지풀';
        pnlName6.Caption  := '새우';               pnlName27.Caption := '곰팡이류';
        pnlName7.Caption  := '복숭아';             pnlName28.Caption := '곰팡이류';
        pnlName8.Caption  := '아카시아';           pnlName29.Caption := '곰팡이류';
        pnlName9.Caption  := '물푸레나무';         pnlName30.Caption := '곰팡이류';
        pnlName10.Caption := '자작나무';           pnlName31.Caption := '고양이';
        pnlName11.Caption := '수양버들';           pnlName32.Caption := '개';
        pnlName12.Caption := '개암나무';           pnlName33.Caption := '바퀴벌레';
        pnlName13.Caption := '일본삼나무';         pnlName34.Caption := '집먼지';
        pnlName14.Caption := '참나무';             pnlName35.Caption := '진드기';
        pnlName15.Caption := '포플라';             pnlName36.Caption := '진드기';
        pnlName16.Caption := '플라타너스';         pnlName37.Caption := '향기풀';
        pnlName17.Caption := '우산잔디';           pnlName38.Caption := '갈대';
        pnlName18.Caption := '오리새';             pnlName39.Caption := '소나무';
        pnlName19.Caption := '큰조아재비';         pnlName40.Caption := '국화';
        pnlName20.Caption := '호밀풀';             pnlName41.Caption := '환삼덩굴';
        pnlName21.Caption := '미역취[국화]';       pnlName42.Caption := '고등어';
     end
     else if Panel_MAST.Caption = 'MAST 식이성 42종' then
     begin
        pnlName1.Caption  := '총 IgE';             pnlName22.Caption := '효모';
        pnlName2.Caption  := '콩';                 pnlName23.Caption := '자작나무';
        pnlName3.Caption  := '우유';               pnlName24.Caption := '참나무';
        pnlName4.Caption  := '체다치즈';           pnlName25.Caption := '호밀풀';
        pnlName5.Caption  := '계란흰자';           pnlName26.Caption := '쑥';
        pnlName6.Caption  := '게';                 pnlName27.Caption := '돼지풀';
        pnlName7.Caption  := '새우';               pnlName28.Caption := '곰팡이류';
        pnlName8.Caption  := '참치';               pnlName29.Caption := '곰팡이류';
        pnlName9.Caption  := '대구';               pnlName30.Caption := '곰팡이류';
        pnlName10.Caption := '연어';               pnlName31.Caption := '고양이';
        pnlName11.Caption := '돼지고기';           pnlName32.Caption := '개';
        pnlName12.Caption := '닭고기';             pnlName33.Caption := '바퀴벌레';
        pnlName13.Caption := '소고기';             pnlName34.Caption := '집먼지';
        pnlName14.Caption := '레몬,오렌지';        pnlName35.Caption := '진드기';
        pnlName15.Caption := '복숭아';             pnlName36.Caption := '진드기';
        pnlName16.Caption := '밀';                 pnlName37.Caption := '메밀';
        pnlName17.Caption := '쌀';                 pnlName38.Caption := '토마토';
        pnlName18.Caption := '보리';               pnlName39.Caption := '칸디다곰팡이';
        pnlName19.Caption := '마늘';               pnlName40.Caption := '수중다리 가루 진드기';
        pnlName20.Caption := '양파';               pnlName41.Caption := '환삼덩굴';
        pnlName21.Caption := '땅콩';               pnlName42.Caption := '고등어';
     end;
  end
  else if Rd_Eng.Checked then
  begin
     if Panel_MAST.Caption = 'MAST 흡입성 42종' then
     begin
        pnlName1.Caption  := 'Total IgE';          pnlName22.Caption := 'Pigweed Mix';
        pnlName2.Caption  := 'Soya beans';         pnlName23.Caption := 'Russian Thistle';
        pnlName3.Caption  := 'Milk';               pnlName24.Caption := 'Dandelion';
        pnlName4.Caption  := 'Egg White';          pnlName25.Caption := 'Mugwort';
        pnlName5.Caption  := 'Crab';               pnlName26.Caption := 'Ragweed. Short';
        pnlName6.Caption  := 'Shrimp';             pnlName27.Caption := 'Alternaria';
        pnlName7.Caption  := 'Peach';              pnlName28.Caption := 'Aspergillus';
        pnlName8.Caption  := 'Acacia';             pnlName29.Caption := 'Cladosporium';
        pnlName9.Caption  := 'Ash Mix';            pnlName30.Caption := 'Penicillium';
        pnlName10.Caption := 'Birch-Alder Mix';    pnlName31.Caption := 'Cat';
        pnlName11.Caption := 'Sallow. Willow';     pnlName32.Caption := 'Dog';
        pnlName12.Caption := 'Hazelnut';           pnlName33.Caption := 'Cockroach Mix';
        pnlName13.Caption := 'Cedar. Japan';       pnlName34.Caption := 'Housedust';
        pnlName14.Caption := 'Oak. White';         pnlName35.Caption := 'D.farinae';
        pnlName15.Caption := 'Poplar Mix';         pnlName36.Caption := 'D.pteronyssinus';
        pnlName16.Caption := 'Sycamore Mix';       pnlName37.Caption := 'Sweet vernal grass';
        pnlName17.Caption := 'Bermuda Grass';      pnlName38.Caption := 'Reed';
        pnlName18.Caption := 'Orchard grass';      pnlName39.Caption := 'Pine';
        pnlName19.Caption := 'Timothy Grass';      pnlName40.Caption := 'Ox-eye-daisy';
        pnlName20.Caption := 'Rye pollens';        pnlName41.Caption := 'Japanese hop';
        pnlName21.Caption := 'Goldenrod';          pnlName42.Caption := 'Mackerel';
     end
     else if Panel_MAST.Caption = 'MAST 식이성 42종' then
     begin
        pnlName1.Caption  := 'Total IgE';          pnlName22.Caption := 'Yeast. Bakers';
        pnlName2.Caption  := 'Soya beans';         pnlName23.Caption := 'Birch-Alder Mix';
        pnlName3.Caption  := 'Milk';               pnlName24.Caption := 'Oak White';
        pnlName4.Caption  := 'Cheddar Cheese';     pnlName25.Caption := 'Rye pollens';
        pnlName5.Caption  := 'Egg White';          pnlName26.Caption := 'Mugwort';
        pnlName6.Caption  := 'Crab';               pnlName27.Caption := 'Ragweed. Short';
        pnlName7.Caption  := 'Shrimp';             pnlName28.Caption := 'Alternaria';
        pnlName8.Caption  := 'Tuna';               pnlName29.Caption := 'Aspergillus';
        pnlName9.Caption  := 'Codfish';            pnlName30.Caption := 'Cladosporium';
        pnlName10.Caption := 'Salmon';             pnlName31.Caption := 'Cat';
        pnlName11.Caption := 'Pork';               pnlName32.Caption := 'Dog';
        pnlName12.Caption := 'Chicken';            pnlName33.Caption := 'Cockroach Mix';
        pnlName13.Caption := 'Beef';               pnlName34.Caption := 'Housedust';
        pnlName14.Caption := 'Citrus Mix';         pnlName35.Caption := 'Derm. farinae';
        pnlName15.Caption := 'Peach';              pnlName36.Caption := 'Derm. pteronyssinus';
        pnlName16.Caption := 'Wheat flour';        pnlName37.Caption := 'Buckwheat meal';
        pnlName17.Caption := 'Rice';               pnlName38.Caption := 'Tomatoe';
        pnlName18.Caption := 'Barley meal';        pnlName39.Caption := 'Candida albicans';
        pnlName19.Caption := 'Garlic';             pnlName40.Caption := 'Acarus siro';
        pnlName20.Caption := 'Onion';              pnlName41.Caption := 'Japanese hop';
        pnlName21.Caption := 'Peanut';             pnlName42.Caption := 'Mackerel';
     end;
  end;
end;

procedure TfrmLD1I16.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I16.btnPSaveClick(Sender: TObject);
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

procedure TfrmLD1I16.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDate            then GF_CalendarClick(mskDate)
   else if Sender = btnDAT_RESULT then GF_CalendarClick(mskDAT_RESULT);
end;

end.



