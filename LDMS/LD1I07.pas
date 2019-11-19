//==============================================================================
// 프로그램 설명 : 2007년도 조직실 결과등록(총센터)
// 시스템        : MDMS
// 수정일자      : 2007.09.06
// 수정자        : 유동구
// 수정내용      : 추가...
//==============================================================================
// 수정일자      : 2008.11.10
// 수정자        : 김승철
// 수정내용      : Class항목 추가관련 수정.
//==============================================================================
unit LD1I07;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I07 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    pnlDetail: TPanel;
    rbDate: TRadioButton;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    Panel2: TPanel;
    pnlJisa: TPanel;
    Label2: TLabel;
    cmbJisa: TComboBox;
    cmbGUBN_JONGRU: TComboBox;
    Panel7: TPanel;
    mskDAT_LAST: TDateEdit;
    btnDAT_LAST: TSpeedButton;
    rbJumin: TRadioButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    btnDate: TSpeedButton;
    mskDate: TDateEdit;
    cmbGubn: TComboBox;
    Panel3: TPanel;
    mskDAT_RESULT: TDateEdit;
    btnDAT_RESULT: TSpeedButton;
    Panel4: TPanel;
    cmbCOD_SAWON: TComboBox;
    Panel5: TPanel;
    cmbCOD_DOCT: TComboBox;
    retDESC_KGBL: TRichEdit;
    cmbGUBN_CLASS: TComboBox;
    Label1: TLabel;
    ckMiSokun: TCheckBox;
    Panel6: TPanel;
    GroupBox1: TGroupBox;
    ckYSNO_VIRUS1: TCheckBox;
    ckYSNO_VIRUS2: TCheckBox;
    ckYSNO_VIRUS3: TCheckBox;
    ckYSNO_VIRUS4: TCheckBox;
    ckYSNO_VIRUS5: TCheckBox;
    Label3: TLabel;
    edtDESC_VIRUS: TEdit;
    GroupBox2: TGroupBox;
    ckYSNO_CLIN1: TCheckBox;
    ckYSNO_CLIN3: TCheckBox;
    ckYSNO_CLIN5: TCheckBox;
    ckYSNO_CLIN2: TCheckBox;
    ckYSNO_CLIN4: TCheckBox;
    ckYSNO_CLIN6: TCheckBox;
    ckYSNO_CLIN7: TCheckBox;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Label4: TLabel;
    edtDESC_JONGRU: TEdit;
    edtDESC_BANG: TEdit;
    edtDESC_BUWI: TEdit;
    edtCod_sokun: TEdit;
    retDesc_Sokun: TMemo;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    ckYsno_Virus7: TCheckBox;
    ckYsno_Virus6: TCheckBox;
    edtDesc_Gum1: TCheckBox;
    edtDesc_Gum2: TCheckBox;
    Label5: TLabel;
    edtGum_Remark: TEdit;
    Label6: TLabel;
    edtGum_Chk1: TCheckBox;
    edtGum_Chk2: TCheckBox;
    edtGum_Form1: TCheckBox;
    edtGum_Form2: TCheckBox;
    edtGum_Form3: TCheckBox;
    edtSang_Form1: TCheckBox;
    edtSang_Form2: TCheckBox;
    edtSang_Form4: TCheckBox;
    edtSang_Form3: TCheckBox;
    edtSun_Form1: TCheckBox;
    edtSun_Form2: TCheckBox;
    Label7: TLabel;
    edtSun_Form4: TEdit;
    edtSun_Form3: TCheckBox;
    qrySokun: TQuery;
    qryCell: TQuery;
    qryGmgn: TQuery;
    qryU_Cell: TQuery;
    qryD_Cell: TQuery;
    GroupBox7: TGroupBox;
    edtCod_Pan1: TRadioButton;
    edtCod_Pan2: TRadioButton;
    edtCod_Pan3: TRadioButton;
    edtCod_Pan4: TRadioButton;
    edtCod_Pan5: TRadioButton;
    edtDesc_Pan: TEdit;
    qryUser_priv: TQuery;
    qryCommon: TQuery;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_Click(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Exit(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure edtCod_sokunExit(Sender: TObject);
    procedure edtCod_Pan2Click(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure cmbRelationChange(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,  UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vCod_cell,
    UV_vCod_hm, UV_vGubn_jongru, UV_vDesc_jongru, UV_vGubn_class, UV_vYsno_virus1,
    UV_vYsno_virus2, UV_vYsno_virus3, UV_vYsno_virus4, UV_vYsno_virus5, UV_vYsno_Virus6, UV_vYsno_Virus7,
    UV_vDesc_virus, UV_vYsno_clin1, UV_vYsno_clin2, UV_vYsno_clin3, UV_vYsno_clin4,
    UV_vYsno_clin5, UV_vYsno_clin6, UV_vYsno_clin7, UV_vDesc_bang, UV_vDesc_buwi,
    UV_vCod_sawon, UV_vCod_doct, UV_vDat_last, UV_vDat_result, UV_vDesc_kgbl,
    UV_vDesc_sokun, UV_vCod_jisa, UV_vDesc_Gum, UV_vGum_Chk, UV_vGum_Remark,
    UV_vGum_Form1, UV_vGum_Form2, UV_vGum_Form3,
    UV_vSang_Form1, UV_vSang_Form2, UV_vSang_Form3, UV_vSang_Form4,
    UV_vSun_Form1, UV_vSun_Form2, UV_vSun_Form3, UV_vSun_Form4, UV_vCod_Pan, UV_vDesc_Pan, UV_vYsno_sokun : Variant;

    // Select문
    UV_sBasicSQL : String;

    // 작업 지사코드
    UV_sCod_jisa : String;
    UV_iCod_doct : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD1I07: TfrmLD1I07;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     LD1I07F, LD1I071;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD1I07.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD1I07.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '주민번호';
         Col[3].Heading   := '이  름';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taLeftJustify;
         Col[3].Color     := clWindow;
         Col[3].Alignment := taLeftJustify;
      end
      else
      begin
         Col[1].Heading   := '지사코드';
         Col[2].Heading   := '검진일자';
         Col[2].Heading   := 'None';
         Col[1].Alignment := taCenter;
         Col[2].Alignment := taCenter;
         Col[3].Color     := clBtnFace;
      end;

      Col[4].Heading := '조직코드';
      Col[5].Heading := '항목코드';
      Col[6].Heading := '분주지사';
      Col[7].Heading := '판독유무';

      // Set Alignment
      Col[4].Alignment := taCenter;
      Col[5].Alignment := taCenter;
      Col[6].Alignment := taCenter;
      Col[7].Alignment := taCenter;
   end;
end;

procedure TfrmLD1I07.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_jongru := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_jongru := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_class  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus1 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus2 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus3 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus4 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus5 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus6 := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_virus7 := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_virus  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin1  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin2  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin3  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin4  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin5  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin6  := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_clin7  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Gum    := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Chk     := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Remark  := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form1   := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form2   := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_Form3   := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form1  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form2  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form3  := VarArrayCreate([0, 0], varOleStr);
      UV_vSang_Form4  := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form1   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form2   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form3   := VarArrayCreate([0, 0], varOleStr);
      UV_vSun_Form4   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_bang   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sawon   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_last    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_result  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_kgbl   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_Pan     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pan    := VarArrayCreate([0, 0], varOleStr);
      UV_vYsno_sokun  := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vCod_cell,    iLength);
      VarArrayRedim(UV_vCod_hm,      iLength);
      VarArrayRedim(UV_vGubn_jongru, iLength);
      VarArrayRedim(UV_vDesc_jongru, iLength);
      VarArrayRedim(UV_vGubn_class,  iLength);
      VarArrayRedim(UV_vYsno_virus1, iLength);
      VarArrayRedim(UV_vYsno_virus2, iLength);
      VarArrayRedim(UV_vYsno_virus3, iLength);
      VarArrayRedim(UV_vYsno_virus4, iLength);
      VarArrayRedim(UV_vYsno_virus5, iLength);
      VarArrayRedim(UV_vYsno_virus6, iLength);
      VarArrayRedim(UV_vYsno_virus7, iLength);
      VarArrayRedim(UV_vDesc_virus,  iLength);
      VarArrayRedim(UV_vYsno_clin1,  iLength);
      VarArrayRedim(UV_vYsno_clin2,  iLength);
      VarArrayRedim(UV_vYsno_clin3,  iLength);
      VarArrayRedim(UV_vYsno_clin4,  iLength);
      VarArrayRedim(UV_vYsno_clin5,  iLength);
      VarArrayRedim(UV_vYsno_clin6,  iLength);
      VarArrayRedim(UV_vYsno_clin7,  iLength);
      VarArrayRedim(UV_vDesc_Gum,    iLength);
      VarArrayRedim(UV_vGum_Chk,     iLength);
      VarArrayRedim(UV_vGum_Remark,  iLength);
      VarArrayRedim(UV_vGum_Form1,   iLength);
      VarArrayRedim(UV_vGum_Form2,   iLength);
      VarArrayRedim(UV_vGum_Form3,   iLength);
      VarArrayRedim(UV_vSang_Form1,  iLength);
      VarArrayRedim(UV_vSang_Form2,  iLength);
      VarArrayRedim(UV_vSang_Form3,  iLength);
      VarArrayRedim(UV_vSang_Form4,  iLength);
      VarArrayRedim(UV_vSun_Form1,   iLength);
      VarArrayRedim(UV_vSun_Form2,   iLength);
      VarArrayRedim(UV_vSun_Form3,   iLength);
      VarArrayRedim(UV_vSun_Form4,   iLength);
      VarArrayRedim(UV_vDesc_bang,   iLength);
      VarArrayRedim(UV_vDesc_buwi,   iLength);
      VarArrayRedim(UV_vCod_sawon,   iLength);
      VarArrayRedim(UV_vCod_doct,    iLength);
      VarArrayRedim(UV_vDat_last,    iLength);
      VarArrayRedim(UV_vDat_result,  iLength);
      VarArrayRedim(UV_vDesc_kgbl,   iLength);
      VarArrayRedim(UV_vDesc_sokun,  iLength);
      VarArrayRedim(UV_vDesc_Pan,    iLength);
      VarArrayRedim(UV_vCod_Pan,     iLength);
      VarArrayRedim(UV_vYsno_sokun,  iLength);
   end;
end;

function TfrmLD1I07.UF_Condition : Boolean;
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

function TfrmLD1I07.UF_Save : Boolean;
var i : Integer;
begin
   Result := False;

   // 작업중인지 Check
   //if not UV_bEdit then exit;

   // Validation Check
   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   try
      with qryU_Cell do
      begin
         ParamByName('GUBN_JONGRU').AsString := Copy(cmbGUBN_JONGRU.Text, 1, 1);
         ParamByName('DESC_JONGRU').AsString := edtDESC_JONGRU.Text;
         if Pos('-', Copy(cmbGUBN_CLASS.Text, 1, 2)) > 0 then
            ParamByName('GUBN_CLASS').AsString  := Copy(cmbGUBN_CLASS.Text, 1, 1)
         else
            ParamByName('GUBN_CLASS').AsString  := Copy(cmbGUBN_CLASS.Text, 1, 2);
         ParamByName('YSNO_VIRUS1').AsString := GF_B2YN(ckYSNO_VIRUS1.Checked);
         ParamByName('YSNO_VIRUS2').AsString := GF_B2YN(ckYSNO_VIRUS2.Checked);
         ParamByName('YSNO_VIRUS3').AsString := GF_B2YN(ckYSNO_VIRUS3.Checked);
         ParamByName('YSNO_VIRUS4').AsString := GF_B2YN(ckYSNO_VIRUS4.Checked);
         ParamByName('YSNO_VIRUS5').AsString := GF_B2YN(ckYSNO_VIRUS5.Checked);
         ParamByName('YSNO_VIRUS6').AsString := GF_B2YN(ckYSNO_VIRUS6.Checked);
         ParamByName('YSNO_VIRUS7').AsString := GF_B2YN(ckYSNO_VIRUS7.Checked);
         ParamByName('DESC_VIRUS').AsString  := edtDESC_VIRUS.Text;
         ParamByName('YSNO_CLIN1').AsString  := GF_B2YN(ckYSNO_CLIN1.Checked);
         ParamByName('YSNO_CLIN2').AsString  := GF_B2YN(ckYSNO_CLIN2.Checked);
         ParamByName('YSNO_CLIN3').AsString  := GF_B2YN(ckYSNO_CLIN3.Checked);
         ParamByName('YSNO_CLIN4').AsString  := GF_B2YN(ckYSNO_CLIN4.Checked);
         ParamByName('YSNO_CLIN5').AsString  := GF_B2YN(ckYSNO_CLIN5.Checked);
         ParamByName('YSNO_CLIN6').AsString  := GF_B2YN(ckYSNO_CLIN6.Checked);
         ParamByName('YSNO_CLIN7').AsString  := GF_B2YN(ckYSNO_CLIN7.Checked);
         If edtDesc_Gum1.Checked Then
            ParamByName('Desc_Gum').AsString  := '1'
         Else If edtDEsc_Gum2.Checked Then
            ParamByName('Desc_Gum').AsString  := '2'
         Else
            ParamByName('Desc_Gum').AsString  := '';
         If edtGum_Chk1.Checked Then
            ParamByName('Gum_Chk').AsString  := '1'
         Else If edtGum_Chk2.Checked Then
            ParamByName('Gum_Chk').AsString  := '2'
         Else
            ParamByName('Gum_Chk').AsString  := '';
         ParamByName('Gum_Remark').AsString  := edtGum_Remark.Text;
         ParamByName('Gum_Form1').AsString   := GF_B2YN(edtGum_Form1.Checked);
         ParamByName('Gum_Form2').AsString   := GF_B2YN(edtGum_Form2.Checked);
         ParamByName('Gum_Form3').AsString   := GF_B2YN(edtGum_Form3.Checked);
         ParamByName('Sang_Form1').AsString  := GF_B2YN(edtSang_Form1.Checked);
         ParamByName('Sang_Form2').AsString  := GF_B2YN(edtSang_Form2.Checked);
         ParamByName('Sang_Form3').AsString  := GF_B2YN(edtSang_Form3.Checked);
         ParamByName('Sang_Form4').AsString  := GF_B2YN(edtSang_Form4.Checked);
         ParamByName('Sun_Form1').AsString   := GF_B2YN(edtSun_Form1.Checked);
         ParamByName('Sun_Form2').AsString   := GF_B2YN(edtSun_Form2.Checked);
         ParamByName('Sun_Form3').AsString   := GF_B2YN(edtSun_Form3.Checked);
         ParamByName('Sun_Form4').AsString   := edtSun_Form4.Text;
         ParamByName('DESC_BANG').AsString   := edtDESC_BANG.Text;
         ParamByName('DESC_BUWI').AsString   := edtDESC_BUWI.Text;
         ParamByName('COD_SAWON').AsString   := Copy(cmbCOD_SAWON.Text, 1, 6);
         ParamByName('COD_DOCT').AsString    := Copy(cmbCOD_DOCT.Text, 1, 6);
         ParamByName('DAT_LAST').AsString    := mskDAT_LAST.Text;
         ParamByName('DAT_RESULT').AsString  := mskDAT_RESULT.Text;
         ParamByName('DESC_KGBL').AsString   := retDESC_KGBL.Text;
         ParamByName('DESC_SOKUN1').AsString  := copy(retDESC_SOKUN.Text,1,250);
         ParamByName('DESC_SOKUN2').AsString  := copy(retDESC_SOKUN.Text,251,250);
         ParamByName('DESC_SOKUN3').AsString  := copy(retDESC_SOKUN.Text,501,250);
         ParamByName('DESC_SOKUN4').AsString  := copy(retDESC_SOKUN.Text,751,250);
         ParamByName('DESC_SOKUN5').AsMemo    := copy(retDESC_SOKUN.Text,1001,length(retDESC_SOKUN.Text) - 1000);
         If edtCod_Pan1.Checked Then
            ParamByName('Cod_Pan').AsString := '1'
         Else If edtCod_Pan2.Checked Then
            ParamByName('Cod_Pan').AsString := '2'
         Else If edtCod_Pan3.Checked Then
            ParamByName('Cod_Pan').AsString := '3'
         Else If edtCod_Pan4.Checked Then
            ParamByName('Cod_Pan').AsString := '4'
         Else If edtCod_Pan5.Checked Then
            ParamByName('Cod_Pan').AsString := '5'
         Else
            ParamByName('Cod_Pan').AsString := '';
         If (edtCod_Pan2.Checked) Or (edtCod_Pan5.Checked)Then
            ParamByName('Desc_Pan').AsString := Trim(edtDesc_Pan.Text)
         Else
            ParamByName('Desc_Pan').AsString := '';

         // 결과일자에 따라서 소견유무 표시 (Y:판독중, C:판독완료, N: 미판독)
         if Trim(mskDAT_RESULT.Text) <> '' then
            ParamByName('YSNO_SOKUN').AsString  := 'Y'
         else
            ParamByName('YSNO_SOKUN').AsString  := 'N';

         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         ParamByName('COD_BJJS').AsString    := UV_vCod_bjjs[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('COD_CELL').AsString    := UV_vCod_cell[i];

         Execsql;
         GP_query_log(qryU_Cell, 'LD1I07', 'Q', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end;

   // Grid에 자료를 반영
   with grdMaster do
   begin
      // 현재 Field의 자료를 Grid관련 Variant변수에 할당
      UV_vGubn_jongru[i] := Copy(cmbGUBN_JONGRU.Text, 1, 1);
      UV_vDesc_jongru[i] := edtDESC_JONGRU.Text;
      if Pos('-', Copy(cmbGUBN_CLASS.Text, 1, 2)) > 0 then
         UV_vGubn_class[i]  := Copy(cmbGUBN_CLASS.Text, 1, 1)
      else
         UV_vGubn_class[i]  := Copy(cmbGUBN_CLASS.Text, 1, 2);
      UV_vYsno_virus1[i] := GF_B2YN(ckYSNO_VIRUS1.Checked);
      UV_vYsno_virus2[i] := GF_B2YN(ckYSNO_VIRUS2.Checked);
      UV_vYsno_virus3[i] := GF_B2YN(ckYSNO_VIRUS3.Checked);
      UV_vYsno_virus4[i] := GF_B2YN(ckYSNO_VIRUS4.Checked);
      UV_vYsno_virus5[i] := GF_B2YN(ckYSNO_VIRUS5.Checked);
      UV_vYsno_virus6[i] := GF_B2YN(ckYSNO_VIRUS6.Checked);
      UV_vYsno_virus7[i] := GF_B2YN(ckYSNO_VIRUS7.Checked);
      UV_vDesc_virus[i]  := edtDESC_VIRUS.Text;
      UV_vYsno_clin1[i]  := GF_B2YN(ckYSNO_CLIN1.Checked);
      UV_vYsno_clin2[i]  := GF_B2YN(ckYSNO_CLIN2.Checked);
      UV_vYsno_clin3[i]  := GF_B2YN(ckYSNO_CLIN3.Checked);
      UV_vYsno_clin4[i]  := GF_B2YN(ckYSNO_CLIN4.Checked);
      UV_vYsno_clin5[i]  := GF_B2YN(ckYSNO_CLIN5.Checked);
      UV_vYsno_clin6[i]  := GF_B2YN(ckYSNO_CLIN6.Checked);
      UV_vYsno_clin7[i]  := GF_B2YN(ckYSNO_CLIN7.Checked);
      UV_vGum_Form1[i]   := GF_B2YN(edtGum_Form1.Checked);
      UV_vGum_Form2[i]   := GF_B2YN(edtGum_Form2.Checked);
      UV_vGum_Form3[i]   := GF_B2YN(edtGum_Form3.Checked);
      UV_vSang_Form1[i]  := GF_B2YN(edtSang_Form1.Checked);
      UV_vSang_Form2[i]  := GF_B2YN(edtSang_Form2.Checked);
      UV_vSang_Form3[i]  := GF_B2YN(edtSang_Form3.Checked);
      UV_vSang_Form4[i]  := GF_B2YN(edtSang_Form4.Checked);
      UV_vSun_Form1[i]   := GF_B2YN(edtSun_Form1.Checked);
      UV_vSun_Form2[i]   := GF_B2YN(edtSun_Form2.Checked);
      UV_vSun_Form3[i]   := GF_B2YN(edtSun_Form3.Checked);
      UV_vSun_Form4[i]   := edtSun_Form4.Text;
      If edtDesc_Gum1.Checked Then
         UV_vDesc_Gum[i] := '1'
      Else If edtDesc_Gum2.Checked Then
         UV_vDesc_Gum[i] := '2';
      If edtGum_Chk1.Checked Then
         UV_vGum_Chk[i] := '1'
      Else If edtGum_Chk2.Checked Then
         UV_vGum_Chk[i] := '2';
      UV_vGum_Remark[i] := edtGum_Remark.Text;
      UV_vDesc_bang[i]   := edtDESC_BANG.Text;
      UV_vDesc_buwi[i]   := edtDESC_BUWI.Text;
      UV_vCod_sawon[i]   := Copy(cmbCOD_SAWON.Text, 1, 6);
      UV_vCod_doct[i]    := Copy(cmbCOD_DOCT.Text, 1, 6);
      UV_vDat_last[i]    := mskDAT_LAST.Text;
      UV_vDat_result[i]  := mskDAT_RESULT.Text;
      UV_vDesc_kgbl[i]   := retDESC_KGBL.Text;
      UV_vDesc_sokun[i]  := retDESC_SOKUN.Text;
      If edtCod_Pan1.Checked Then
         UV_vCod_Pan[i] := '1'
      Else If edtCod_Pan2.Checked Then
         UV_VCod_Pan[i] := '2'
      Else If edtCod_Pan3.Checked Then
         UV_VCod_Pan[i] := '3'
      Else If edtCod_Pan4.Checked Then
         UV_VCod_Pan[i] := '4'
      Else If edtCod_Pan5.Checked Then
         UV_VCod_Pan[i] := '5';
      UV_vDesc_Pan[i] := Trim(edtDesc_Pan.Text);

      if Trim(mskDAT_RESULT.Text) <> '' then
         UV_vYsno_sokun[i] := 'Y'
      else
         UV_vYsno_sokun[i] := 'N';
      // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;
   end;

   UV_iCod_doct := cmbCOD_DOCT.ItemIndex;
   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLD1I07.btnSaveClick(Sender: TObject);
begin
   inherited;
   If edtCod_Pan2.Checked Then
      If Trim(edtDesc_Pan.Text) = '' Then
         Begin
            Showmessage('☞ 의심 개월수를 입력하세요...');
            Exit;
         End;
   UF_Save;
end;

procedure TfrmLD1I07.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);

   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   pnlJisa.Visible := True;                                                     //06.11.24 철. 지사선택 보이게끔.

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   // Combo Populate
   GP_ComboFill(cmbGUBN_JONGRU, 'GCJR');
//   GP_ComboFill(cmbGUBN_CLASS,  'CLAS');
   with qryCommon do
   begin
//      Close;
      Open;
      while not Eof do
      begin
         cmbGUBN_CLASS.Items.add(FieldByName('COD_DETAIL').AsString + '-' +
                             FieldByName('DESC_CODE').AsString);
         Next;
      end;
   end;
   GP_ComboSawonAll(cmbCOD_SAWON, 0);
   GP_ComboSawonAll(cmbCOD_DOCT, 1);

   // 분석부 조직실 전체...
   if (GV_sUserID = '150365') or (GV_sUserID = '410015') then cmbRelation.Enabled := True
//   if GV_sUserID = '150143' then cmbRelation.Enabled := True
   else                          cmbRelation.Enabled := False;

   // SQL문 저장
   UV_sBasicSQL := qryCell.SQL.Text;

   UV_iCod_doct := -1;
   cmbGubn.ItemIndex := 0;
   mskDate.Text := GV_sToday;
end;

procedure TfrmLD1I07.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_Jisa[DataRow-1];
      2 : begin
             if rbDate.Checked  then Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
             if rbJumin.Checked then Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
          end;
      3 : begin
             if rbDate.Checked  then Value := UV_vDesc_name[DataRow-1];
             if rbJumin.Checked then Value := '';
          end;
      4 : Value := UV_vCod_cell[DataRow-1];
      5 : Value := UV_vCod_hm[DataRow-1];
      6 : Value := UV_vCod_bjjs[DataRow-1];
      7 : begin
             if UV_vDat_gmgn[DataRow-1] >= '20070701' then
             begin
                if      Trim(UV_vYsno_sokun[DataRow-1]) = 'C' then Value := '판독완료'
                else if Trim(UV_vYsno_sokun[DataRow-1]) = 'Y' then Value := '*판독중*'
                else Value := '미판독';
             end
             else
             begin
                if (Trim(UV_vYsno_sokun[DataRow-1]) = 'Y') or
                   (Trim(UV_vYsno_sokun[DataRow-1]) = 'C') then Value := '판독완료'
                else
                begin
                   if UV_vDat_result[DataRow-1] <> '' then Value := '판독완료'
                   else                                    Value := '미판독';
                end;
             end;
          end;
   end;

   // 조직검사일 경우 색깔 변경...
   if (Copy(UV_vCod_cell[DataRow-1],1,1) = 'J') then
      grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;
end;

procedure TfrmLD1I07.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderby : String;
begin
   inherited;
   UV_bEdit := True;
   btnPsave.Enabled   := True;
   btnPcancel.Enabled := True;

   with qryUser_priv do                              //06.11.24 철. 조회지사가 로그인한 아이디의 지사와 일치할 경우 저장할 수 있음.
   begin
      close;
      ParamByName('cod_sawon').AsString := GV_sUserId;
      open;
      if recordcount > 0 then
      begin
         if (copy(cmbJisa.Text,1,2) = copy(GV_sUserID,1,2)) and (FieldBYName('ysno_read').AsString = 'N') then
            btnPSave.Enabled := True
         else
            btnPSave.Enabled := False;
      end;
   end;

   // 분석부 조직실 전체...
   if (GV_sUserID = '150365') or (GV_sUserID = '150281') or
      (GV_sUserID = '150332') or (GV_sUserID = '120028') or
      (GV_sUserID = '120094') or (GV_sUserID = '410015') then btnPSave.Enabled := True
   else                          btnPSave.Enabled := False;
   
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid Setting
   UP_GridSet;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryCell do
   begin
      Active := False;

      sWhere := 'WHERE ';

      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' A.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
         if ckMiSokun.Checked then
         begin
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''' +
                               ' AND A.cod_doct is Null';
         end
         else
            sWhere := sWhere + ' A.dat_gmgn = ''' + mskDate.Text + '''';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' B.num_jumin = ''' + mskJumin.Text + '''';
      end;

      // 검사구분에 따라서 Where 조건추가
      if cmbGubn.ItemIndex = 1 then
         sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''P'''
      else if cmbGubn.ItemIndex = 2 then
         sWhere := sWhere + ' AND SUBSTRING(cod_cell,1,1) = ''J''';

      sOrderBy := ' ORDER BY A.dat_gmgn, A.cod_jisa, B.desc_name ';
      SQL.Clear;
      SQL.Add(UV_sBasicSQL + sWhere + sOrderBy);
      Active := True;

      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD1I07', 'Q', 'N');
         while Not Eof do
         begin
            UP_VarMemSet(index);
            UV_vCod_bjjs[index]    := FieldByName('COD_BJJS').AsString;
            UV_vNum_id[index]      := FieldByName('NUM_ID').AsString;
            UV_vNum_jumin[index]   := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]   := FieldByName('DESC_NAME').AsString;
            UV_vCod_jisa[index]    := FieldByName('COD_JISA').AsString;
            UV_vDat_gmgn[index]    := FieldByName('DAT_GMGN').AsString;
            UV_vCod_cell[index]    := FieldByName('COD_CELL').AsString;
            UV_vCod_hm[index]      := FieldByName('COD_HM').AsString;
            UV_vGubn_jongru[index] := FieldByName('GUBN_JONGRU').AsString;
            UV_vDesc_jongru[index] := FieldByName('DESC_JONGRU').AsString;
            UV_vGubn_class[index]  := FieldByName('GUBN_CLASS').AsString;
            UV_vYsno_virus1[index] := FieldByName('YSNO_VIRUS1').AsString;
            UV_vYsno_virus2[index] := FieldByName('YSNO_VIRUS2').AsString;
            UV_vYsno_virus3[index] := FieldByName('YSNO_VIRUS3').AsString;
            UV_vYsno_virus4[index] := FieldByName('YSNO_VIRUS4').AsString;
            UV_vYsno_virus5[index] := FieldByName('YSNO_VIRUS5').AsString;
            UV_vYsno_virus6[index] := FieldByName('YSNO_VIRUS6').AsString;
            UV_vYsno_virus7[index] := FieldByName('YSNO_VIRUS7').AsString;
            UV_vDesc_virus[index]  := FieldByName('DESC_VIRUS').AsString;
            UV_vYsno_clin1[index]  := FieldByName('YSNO_CLIN1').AsString;
            UV_vYsno_clin2[index]  := FieldByName('YSNO_CLIN2').AsString;
            UV_vYsno_clin3[index]  := FieldByName('YSNO_CLIN3').AsString;
            UV_vYsno_clin4[index]  := FieldByName('YSNO_CLIN4').AsString;
            UV_vYsno_clin5[index]  := FieldByName('YSNO_CLIN5').AsString;
            UV_vYsno_clin6[index]  := FieldByName('YSNO_CLIN6').AsString;
            UV_vYsno_clin7[index]  := FieldByName('YSNO_CLIN7').AsString;
            UV_vDesc_Gum[Index]    := FieldByName('Desc_Gum').AsString;
            UV_vGum_Chk[Index]     := FieldByName('Gum_Chk').AsString;
            UV_vGum_Remark[Index]  := FieldByName('Gum_Remark').AsString;
            UV_vGum_Form1[Index]   := FieldByName('Gum_Form1').AsString;
            UV_vGum_Form2[Index]   := FieldByName('Gum_Form2').AsString;
            UV_vGum_Form3[Index]   := FieldByName('Gum_Form3').AsString;
            UV_vSang_Form1[Index]  := FieldByName('Sang_Form1').AsString;
            UV_vSang_Form2[Index]  := FieldByName('Sang_Form2').AsString;
            UV_vSang_Form3[Index]  := FieldByName('Sang_Form3').AsString;
            UV_vSang_Form4[Index]  := FieldByName('Sang_Form4').AsString;
            UV_vSun_Form1[Index]   := FieldByName('Sun_Form1').AsString;
            UV_vSun_Form2[Index]   := FieldByName('Sun_Form2').AsString;
            UV_vSun_Form3[Index]   := FieldByName('Sun_Form3').AsString;
            UV_vSun_Form4[Index]   := FieldByName('Sun_Form4').AsString;
            UV_vDesc_bang[index]   := FieldByName('DESC_BANG').AsString;
            UV_vDesc_buwi[index]   := FieldByName('DESC_BUWI').AsString;
            UV_vCod_sawon[index]   := FieldByName('COD_SAWON').AsString;
            UV_vCod_doct[index]    := FieldByName('COD_DOCT').AsString;
            UV_vCod_Pan[index]     := FieldByName('COD_PAN').AsString;
            UV_vDESC_Pan[index]    := FieldByName('DESC_PAN').AsString;
            UV_vDat_last[index]    := FieldByName('DAT_LAST').AsString;
            UV_vDat_result[index]  := FieldByName('DAT_RESULT').AsString;
            UV_vDesc_kgbl[index]   := FieldByName('DESC_KGBL').AsString;
            UV_vDesc_sokun[index]  := FieldByName('DESC_SOKUN1').AsString +
                                      FieldByName('DESC_SOKUN2').AsString +
                                      FieldByName('DESC_SOKUN3').AsString +
                                      FieldByName('DESC_SOKUN4').AsString +
                                      FieldByName('DESC_SOKUN5').AsString;

            UV_vYsno_sokun[index]  := FieldByName('Ysno_sokun').AsString;

            Next;
            Inc(index);
         end;

         // Table과 Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD1I07.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // 취소시
   // 1.입력이면 화면 Clear
   // 2.수정이면 Re-query
   if UV_iStatus = GC_Insert then
   begin
      GP_FieldClear(pnlDetail);
   end;

   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD1I07.UP_Change(Sender: TObject);
begin
   inherited;
   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I07.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;
   edtGum_Chk1.checked := false;
   edtGum_Chk2.checked := false;

   // Grid의 Row가 변경되면 Detail에 자료를 할당
   GP_ComboDisplay(cmbGUBN_JONGRU, UV_vGubn_jongru[NewRow-1], 1);
   edtDESC_JONGRU.Text   := UV_vDesc_jongru[NewRow-1];
   if length(UV_vGubn_class[NewRow-1]) > 1 then
      GP_ComboDisplay(cmbGUBN_CLASS, UV_vGubn_class[NewRow-1], 2)
   else
      GP_ComboDisplay(cmbGUBN_CLASS, UV_vGubn_class[NewRow-1], 1);
   ckYSNO_VIRUS1.Checked := GF_YN2B(UV_vYsno_virus1[NewRow-1]);
   ckYSNO_VIRUS2.Checked := GF_YN2B(UV_vYsno_virus2[NewRow-1]);
   ckYSNO_VIRUS3.Checked := GF_YN2B(UV_vYsno_virus3[NewRow-1]);
   ckYSNO_VIRUS4.Checked := GF_YN2B(UV_vYsno_virus4[NewRow-1]);
   ckYSNO_VIRUS5.Checked := GF_YN2B(UV_vYsno_virus5[NewRow-1]);
   If Trim(UV_vYsno_virus6[NewRow-1]) = '' Then
      ckYSno_virus6.Checked := True
   Else
      ckYSNO_VIRUS6.Checked := GF_YN2B(UV_vYsno_virus6[NewRow-1]);
   ckYSNO_VIRUS7.Checked := GF_YN2B(UV_vYsno_virus7[NewRow-1]);
   edtDESC_VIRUS.Text    := UV_vDesc_virus[NewRow-1];
   ckYSNO_CLIN1.Checked  := GF_YN2B(UV_vYsno_clin1[NewRow-1]);
   ckYSNO_CLIN2.Checked  := GF_YN2B(UV_vYsno_clin2[NewRow-1]);
   ckYSNO_CLIN3.Checked  := GF_YN2B(UV_vYsno_clin3[NewRow-1]);
   ckYSNO_CLIN4.Checked  := GF_YN2B(UV_vYsno_clin4[NewRow-1]);
   ckYSNO_CLIN5.Checked  := GF_YN2B(UV_vYsno_clin5[NewRow-1]);
   ckYSNO_CLIN6.Checked  := GF_YN2B(UV_vYsno_clin6[NewRow-1]);
   ckYSNO_CLIN7.Checked  := GF_YN2B(UV_vYsno_clin7[NewRow-1]);
   edtDESC_BANG.Text     := UV_vDesc_bang[NewRow-1];
   edtDESC_BUWI.Text     := UV_vDesc_buwi[NewRow-1];
   if Trim(UV_vCod_sawon[NewRow-1]) = '' then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         cmbCOD_SAWON.Font.Color := clRed;
         GP_ComboDisplay(cmbCOD_SAWON, '150365', 6);
      end
      else cmbCOD_SAWON.ItemIndex := -1;

//      GP_ComboDisplay(cmbCOD_SAWON, '150284', 6)
//      GP_ComboDisplay(cmbCOD_SAWON, '150136', 6)
   end
   else
   begin
      cmbCOD_SAWON.Font.Color := clBlack;
      GP_ComboDisplay(cmbCOD_SAWON, UV_vCod_sawon[NewRow-1], 6);
   end;
//kmi 2003-08-09 의사선생 Fixed
   if Trim(UV_vCod_Doct[NewRow-1]) = '' then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         cmbCOD_Doct.Font.Color := clRed;
         GP_ComboDisplay(cmbCOD_Doct, '150007', 6);
      end
      else cmbCOD_Doct.ItemIndex := -1;
   end
   else
   begin
      cmbCOD_Doct.Font.Color := clBlack;
      GP_ComboDisplay(cmbCOD_Doct, UV_vCod_Doct[NewRow-1], 6);
   end;
//
//   GP_ComboDisplay(cmbCOD_DOCT,  UV_vCod_doct[NewRow-1], 6);
   If UV_vDesc_Gum[NewRow-1] = '1' Then
      Begin
         edtDesc_Gum1.Checked := true;
         edtDesc_Gum2.Checked := False;
      End
   Else If UV_vDesc_Gum[NewRow-1] = '2' Then
      Begin
         edtDesc_Gum1.Checked := False;
         edtDesc_Gum2.Checked := True;
      End
   Else
      Begin
         edtDesc_Gum1.Checked := True;
         edtDesc_Gum2.Checked := False;
      End;
   If UV_vGum_Chk[NewRow-1] = '1' Then
      Begin
         edtGum_Chk1.Checked := true;
         edtGum_Chk2.Checked := False;
      End
   Else If UV_vGum_Chk[NewRow-1] = '2' Then
      Begin
         edtGum_Chk1.Checked := False;
         edtGum_Chk2.Checked := True;
      End
   Else;
      Begin
         edtGum_Chk1.Checked := True;
         edtGum_Chk2.Checked := False;
      End;

   edtGum_Remark.Text    := UV_vGum_Remark[NewRow-1];
   If Trim(UV_vGum_Form1[NewRow-1]) = '' Then
      edtGum_Form1.Checked := True
   Else
      edtGum_Form1.Checked  := GF_YN2B(UV_vGum_Form1[NewRow-1]);
   edtGum_Form2.Checked  := GF_YN2B(UV_vGum_Form2[NewRow-1]);
   edtGum_Form3.Checked  := GF_YN2B(UV_vGum_Form3[NewRow-1]);
   edtSang_Form1.Checked := GF_YN2B(UV_vSang_Form1[NewRow-1]);
   edtSang_Form2.Checked := GF_YN2B(UV_vSang_Form2[NewRow-1]);
   edtSang_Form3.Checked := GF_YN2B(UV_vSang_Form3[NewRow-1]);
   edtSang_Form4.Checked := GF_YN2B(UV_vSang_Form4[NewRow-1]);
   edtSun_Form1.Checked  := GF_YN2B(UV_vSun_Form1[NewRow-1]);
   edtSun_Form2.Checked  := GF_YN2B(UV_vSun_Form2[NewRow-1]);
   edtSun_Form3.Checked  := GF_YN2B(UV_vSun_Form3[NewRow-1]);
   edtSun_Form4.Text     := UV_vSun_Form4[NewRow-1];
   mskDAT_LAST.Text      := UV_vDat_last[NewRow-1];
   if Trim(UV_vDat_result[NewRow-1]) = '' Then
   begin
      if copy(GV_sUserId,1,2) = '15' then
      begin
         mskDat_Result.Font.Color := clRed;
         mskDat_Result.Text       := GV_sToday;
      end
      else mskDat_Result.Text := '';
   end
   else
   begin

      mskDat_Result.Font.Color := clBlack;
      mskDAT_RESULT.Text       := UV_vDat_result[NewRow-1];
   end;
   retDESC_KGBL.Text     := UV_vDesc_kgbl[NewRow-1];
   retDESC_SOKUN.Text    := UV_vDesc_sokun[NewRow-1];
   edtCod_Pan1.Checked := False;
   edtCod_Pan2.Checked := False;
   edtCod_Pan3.Checked := False;
   edtCod_Pan4.Checked := False;
   edtCod_Pan5.Checked := False;
   If UV_vCod_Pan[NewRow-1] = '1' Then
      edtCod_Pan1.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '2' Then
      edtCod_Pan2.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '3' Then
      edtCod_Pan3.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '4' Then
      edtCod_Pan4.Checked := True
   Else If UV_vCod_Pan[NewRow-1] = '5' Then
      edtCod_Pan5.Checked := True;
   If  (UV_vCod_Pan[NewRow-1] = '2') Or (UV_vCod_Pan[NewRow-1] = '5') Then
      edtDesc_Pan.Text := UV_vDesc_Pan[NewRow-1]
   Else
      edtDesc_Pan.Text := '';

   if cmbCod_doct.ItemIndex = -1 then
      cmbCod_doct.ItemIndex := UV_iCod_doct;
   // Grid Focus
   grdMaster.SetFocus;
                                                 
   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD1I07.UP_Click(Sender: TObject);
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
   else if Sender = btnDate       then GF_CalendarClick(mskDate)
   else if Sender = btnDAT_LAST   then GF_CalendarClick(mskDAT_LAST)
   else if Sender = btnDAT_RESULT then GF_CalendarClick(mskDAT_RESULT);
end;

procedure TfrmLD1I07.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      // 찾기기능 활성화
      btnFind.Enabled   := True;

      mskDate.Color     := $00E6FFFE;
      mskDate.Enabled   := True;
      btnDate.Enabled   := True;
      ckMiSokun.Enabled := True;
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
      ckMiSokun.Enabled := False;
      mskJumin.Color    := $00E6FFFE;
      mskJumin.Enabled  := True;
      btnJumin.Enabled  := True;
      mskJumin.SetFocus;
   end;
end;

procedure TfrmLD1I07.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // Reference Key가 아니면 종료
   if Key <> GC_Refer then exit;

   if      Sender = mskDate       then UP_Click(btnDate)
   else if Sender = mskJumin      then UP_Click(btnJumin)
   else if Sender = mskDAT_LAST   then UP_Click(btnDAT_LAST)
   else if Sender = mskDAT_RESULT then UP_Click(btnDAT_RESULT);
end;

procedure TfrmLD1I07.UP_Exit(Sender: TObject);
begin
   inherited;

   if Sender = mskJumin then
   begin
      // 주민번호의 입력오류일 경우 처리
      if Length(mskJumin.Text) < 13 then exit;

      with qryGmgn do
      begin
         Active := False;
         ParamByName('NUM_JUMIN').AsString := mskJumin.Text;
         Active := True;

         if RecordCount > 0 then
            edtName.Text := FieldByName('DESC_NAME').AsString
         else
         begin
            GF_MsgBox('Warning', '해당 주민번호는 검진자가 아닙니다.');
            edtName.Text := '';
            mskJumin.SetFocus;
         end;
      end;
   end;
end;

procedure TfrmLD1I07.btnFindClick(Sender: TObject);
begin
   inherited;

   frmLD1I07F := TfrmLD1I07F.Create(Self);
   frmLD1I07F.ShowModal;
end;

procedure TfrmLD1I07.btnDeleteClick(Sender: TObject);
begin
   inherited;

   // Delete Confirm Message
   if (grdMaster.Rows = 0) or (grdMaster.VisibleRows < 1) then exit;
   if not GF_MsgBox('Confirmation', '해당 조직검사 자료를 삭제하면 복구가 불가능합니다.' + Chr(13) +
                                    '정말로 삭제하시겠습니까 ?') then exit;

   // Delete 작업수행
   try
      with qryD_Cell do
      begin
         ParamByName('COD_BJJS').AsString := UV_vCod_bjjs[grdMaster.CurrentDataRow-1];
         ParamByName('NUM_ID').AsString   := UV_vNum_id[grdMaster.CurrentDataRow-1];
         ParamByName('COD_CELL').AsString := UV_vCod_cell[grdMaster.CurrentDataRow-1];

         ExecSql;
         GP_query_log(qryD_Cell, 'LD1I07', 'Q', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end; 

   // 화면에서 안보이게함
   grdMaster.RowVisible[grdMaster.CurrentDataRow] := False;
   if grdMaster.VisibleRows < 1 then GP_FieldClear(pnlDetail);

   // Delete Mode & Msg
   UP_SetMode('DELETE');
end;

procedure TfrmLD1I07.edtCod_sokunExit(Sender: TObject);
begin
  inherited;
   qrySokun.Active := false;
   qrySokun.ParamByname('Cod_sokun').AsString := edtCod_sokun.text;
   qrySokun.Active := true;

    if  qrySokun.RecordCount > 0 then
        retDesc_sokun.lines.add(qrySokun.FieldByName('desc_sbsg').AsString +
                                qrySokun.FieldByName('desc_sbsg1').AsString +
                                qrySokun.FieldByName('desc_sbsg2').AsString )
    else
         GF_MsgBox('Information', ' 없는 소견 코드입니다. ' + Chr(13) +
                                  ' 다시한번 확인하십시오');
    edtCod_sokun.Text := '';
end;

procedure TfrmLD1I07.edtCod_Pan2Click(Sender: TObject);
begin
     inherited;
     If edtCod_Pan2.Checked Then
        Begin
           edtDesc_Pan.Visible := True;
           edtDesc_Pan.MaxLength := 2;
        End
     Else If edtCod_Pan5.Checked Then
        Begin
           edtDesc_Pan.Visible := True;
           edtDesc_Pan.MaxLength := 0;
        End
     Else edtDesc_Pan.Visible := False;
end;

procedure TfrmLD1I07.cmbRelationChange(Sender: TObject);
begin
   inherited;
   if cmbRelation.ItemIndex = 0 then
   begin
      Application.CreateForm(TfrmLD1I071, frmLD1I071);
      frmLD1I071.Show;
   end;
end;

end.
