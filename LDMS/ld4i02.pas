//==============================================================================
// 프로그램 설명 : 세포병리검사의뢰서
// 시스템        : LDMS
// 수정일자      : 2010.05.12
// 수정자        : 김승철
// 수정내용      :
//==============================================================================
// 수정일자      : 2014.07.28
// 수정자        : 곽윤설
// 수정내용      : 중복 조회 수정
//==============================================================================
// 수정일자      : 2014.11.04
// 수정자        : 곽윤설
// 수정내용      : 담당의사 고정 기능 추가
// 참고사항      : 광주-이화성 [한의 광주종합검사파트1400898]
//==============================================================================
// 수정일자      : 2015.06.15
// 수정자        : 곽윤설
// 수정내용      : B009 추가
// 참고사항      : 한의 재단병리팀장1500019
//==============================================================================

unit LD4I02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, ORPrint;

type
  TfrmLD4I02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    pnlDetail: TPanel;
    rbDate: TRadioButton;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlJisa: TPanel;
    Label2: TLabel;
    cmbJisa: TComboBox;
    rbJumin: TRadioButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
    btnDate: TSpeedButton;
    mskDate: TDateEdit;
    cmbGubn: TComboBox;
    Label1: TLabel;
    desc_sokun: TMemo;
    qryGumgin: TQuery;
    Panel2: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    desc_buwi: TEdit;
    desc_jisa: TEdit;
    dat_gmgn: TDateEdit;
    desc_name: TEdit;
    age_sex: TEdit;
    desc_dc: TEdit;
    dat_take: TDateEdit;
    btnDat_Take: TSpeedButton;
    Panel14: TPanel;
    GE_Conventional: TCheckBox;
    GE_Liquid: TCheckBox;
    Panel15: TPanel;
    Label3: TLabel;
    GE_Lmp: TCheckBox;
    GE_Lmp_text: TEdit;
    GE_Menopause: TCheckBox;
    GE_Menopause_text: TEdit;
    GE_Pregnancy: TCheckBox;
    GE_Pregnancy_text: TEdit;
    GE_IUD: TCheckBox;
    GE_Erosion: TCheckBox;
    GE_Hormone: TCheckBox;
    GE_Radiotherapy: TCheckBox;
    num_jumin: TMaskEdit;
    Panel3: TPanel;
    Panel4: TPanel;
    NGE_Sputum: TCheckBox;
    NGE_Pleural: TCheckBox;
    NGE_Ascitic: TCheckBox;
    NGE_Joint: TCheckBox;
    NGE_Urine: TCheckBox;
    NGE_CSF: TCheckBox;
    NGE_Other: TCheckBox;
    NGE_Cell_block: TCheckBox;
    Panel7: TPanel;
    Panel16: TPanel;
    FNA_Thyroid: TCheckBox;
    FNA_Lymph: TCheckBox;
    FNA_Breast: TCheckBox;
    FNA_Other: TCheckBox;
    FNA_Cell_block: TCheckBox;
    Panel18: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    qryI_Ca_Request: TQuery;
    qryCa_Request: TQuery;
    qryU_Ca_Request: TQuery;
    qryUser_priv: TQuery;
    qryGmgn: TQuery;
    gum_type1: TRadioButton;
    gum_type2: TRadioButton;
    gum_type3: TRadioButton;
    gum_type4: TRadioButton;
    Label6: TLabel;
    Cmb_Order: TComboBox;
    Panel17: TPanel;
    Panel19: TPanel;
    cmbCOD_DOCT: TComboBox;
    dat_human: TEdit;
    Panel20: TPanel;
    Virus_Y: TRadioButton;
    Virus_N: TRadioButton;
    Panel21: TPanel;
    Edt_Num_id: TEdit;
    Panel22: TPanel;
    Dat_Ask: TDateEdit;
    btnDat_Ask: TSpeedButton;
    sPrint: TBitBtn;
    DuplexUterus_Y: TRadioButton;
    DuplexUterus_N: TRadioButton;
    CB_fix_Doc: TCheckBox;
    SpeedButton1: TSpeedButton;
    qryD_Ca_Request: TQuery;
    rbCell: TRadioButton;
    mskcell: TMaskEdit;
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
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure sPrintClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_jisa,  UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vCod_hm, UV_vDuplexUterus_YN,
    UV_vDesc_buwi, UV_vdat_human, UV_vVirus_YN, UV_vcod_doctor,
    UV_vDat_take,  UV_vDat_Ask,   UV_vDesc_jisa,   UV_vDesc_dc,   UV_vGum_type,   UV_vDesc_sokun,
    UV_vGE_Conventional,   UV_vGE_Liquid,   UV_vGE_Lmp,   UV_vGE_Lmp_text,
    UV_vGE_Menopause,   UV_vGE_Menopause_text,   UV_vGE_Pregnancy,   UV_vGE_Pregnancy_text,
    UV_vGE_IUD,   UV_vGE_Erosion,   UV_vGE_Hormone,   UV_vGE_Radiotherapy,   UV_vNGE_Sputum,
    UV_vNGE_Pleural,   UV_vNGE_Ascitic,   UV_vNGE_Joint,   UV_vNGE_Urine,   UV_vNGE_CSF,
    UV_vNGE_Other,   UV_vNGE_Cell_block,   UV_vFNA_Thyroid,   UV_vFNA_Lymph,
    UV_vFNA_Breast,   UV_vFNA_Other,   UV_vFNA_Cell_block : Variant;

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
    sFix_doctor : String;
  end;

var
  frmLD4I02: TfrmLD4I02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     LD1I07F, LD1I071, LD4I021;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD4I02.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD4I02.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      Col[1].Heading   := '지사코드';
      Col[2].Heading   := '주민번호';
      Col[3].Heading   := '이  름';
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taLeftJustify;
      Col[3].Color     := clWindow;
      Col[3].Alignment := taLeftJustify;
      Col[4].Heading := '항목코드';

      // Set Alignment
      Col[4].Alignment := taCenter;
   end;
end;

procedure TfrmLD4I02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_jisa          := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id            := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin         := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name         := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn          := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi         := VarArrayCreate([0, 0], varOleStr);
      UV_vdat_human         := VarArrayCreate([0, 0], varOleStr);
      UV_vVirus_YN          := VarArrayCreate([0, 0], varOleStr);
      UV_vcod_doctor        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_take          := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_Ask           := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_jisa         := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc           := VarArrayCreate([0, 0], varOleStr);
      UV_vGum_type          := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun        := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm            := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Conventional   := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Liquid         := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Lmp            := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Lmp_text       := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Menopause      := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Menopause_text := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Pregnancy      := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Pregnancy_text := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_IUD            := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Erosion        := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Hormone        := VarArrayCreate([0, 0], varOleStr);
      UV_vGE_Radiotherapy   := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Sputum        := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Pleural       := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Ascitic       := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Joint         := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Urine         := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_CSF           := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Other         := VarArrayCreate([0, 0], varOleStr);
      UV_vNGE_Cell_block    := VarArrayCreate([0, 0], varOleStr);
      UV_vFNA_Thyroid       := VarArrayCreate([0, 0], varOleStr);
      UV_vFNA_Lymph         := VarArrayCreate([0, 0], varOleStr);
      UV_vFNA_Breast        := VarArrayCreate([0, 0], varOleStr);
      UV_vFNA_Other         := VarArrayCreate([0, 0], varOleStr);
      UV_vFNA_Cell_block    := VarArrayCreate([0, 0], varOleStr);
      UV_vDuplexUterus_YN   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_jisa,           iLength);
      VarArrayRedim(UV_vNum_id,             iLength);
      VarArrayRedim(UV_vNum_jumin,          iLength);
      VarArrayRedim(UV_vDesc_name,          iLength);
      VarArrayRedim(UV_vDat_gmgn,           iLength);
      VarArrayRedim(UV_vDesc_buwi,          iLength);
      VarArrayRedim(UV_vdat_human,          iLength);
      VarArrayRedim(UV_vVirus_YN,           iLength);
      VarArrayRedim(UV_vcod_doctor,         iLength);
      VarArrayRedim(UV_vDat_take,           iLength);
      VarArrayRedim(UV_vDat_Ask,            iLength);
      VarArrayRedim(UV_vDesc_jisa,          iLength);
      VarArrayRedim(UV_vDesc_dc,            iLength);
      VarArrayRedim(UV_vGum_type,           iLength);
      VarArrayRedim(UV_vDesc_sokun,         iLength);
      VarArrayRedim(UV_vCod_hm,             iLength);
      VarArrayRedim(UV_vGE_Conventional,    iLength);
      VarArrayRedim(UV_vGE_Liquid,          iLength);
      VarArrayRedim(UV_vGE_Lmp,             iLength);
      VarArrayRedim(UV_vGE_Lmp_text,        iLength);
      VarArrayRedim(UV_vGE_Menopause,       iLength);
      VarArrayRedim(UV_vGE_Menopause_text,  iLength);
      VarArrayRedim(UV_vGE_Pregnancy,       iLength);
      VarArrayRedim(UV_vGE_Pregnancy_text,  iLength);
      VarArrayRedim(UV_vGE_IUD,             iLength);
      VarArrayRedim(UV_vGE_Erosion,         iLength);
      VarArrayRedim(UV_vGE_Hormone,         iLength);
      VarArrayRedim(UV_vGE_Radiotherapy,    iLength);
      VarArrayRedim(UV_vNGE_Sputum,         iLength);
      VarArrayRedim(UV_vNGE_Pleural,        iLength);
      VarArrayRedim(UV_vNGE_Ascitic,        iLength);
      VarArrayRedim(UV_vNGE_Joint,          iLength);
      VarArrayRedim(UV_vNGE_Urine,          iLength);
      VarArrayRedim(UV_vNGE_CSF,            iLength);
      VarArrayRedim(UV_vNGE_Other,          iLength);
      VarArrayRedim(UV_vNGE_Cell_block,     iLength);
      VarArrayRedim(UV_vFNA_Thyroid,        iLength);
      VarArrayRedim(UV_vFNA_Lymph,          iLength);
      VarArrayRedim(UV_vFNA_Breast,         iLength);
      VarArrayRedim(UV_vFNA_Other,          iLength);
      VarArrayRedim(UV_vFNA_Cell_block,     iLength);
      VarArrayRedim(UV_vDuplexUterus_YN,    iLength);
   end;
end;

function TfrmLD4I02.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (pnlJisa.Visible and (cmbJisa.ItemIndex = -1)) or
      (rbDate.Checked and (mskDate.Text = '')) or
      (rbJumin.Checked and (mskJumin.Text = '')) or
      (rbCell.Checked and (mskCell.Text = '')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

function TfrmLD4I02.UF_Save : Boolean;
var i : Integer;
begin
   Result := False;

   // Validation Check
   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   try
      with qryCa_Request do
      begin
         Close;
         ParamByName('cod_jisa').AsString    := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('dat_gmgn').AsString    := UV_vDat_gmgn[i];
         Open;
         if qryCa_Request.IsEmpty = False then
         begin
            with qryU_Ca_Request do
            begin
               ParamByName('cod_jisa').AsString             := UV_vCod_jisa[i];
               ParamByName('NUM_ID').AsString               := UV_vNum_id[i];
               ParamByName('dat_gmgn').AsString             := dat_gmgn.Text;
               ParamByName('dat_human').AsString            := dat_human.Text;
               if Virus_Y.Checked then
                  ParamByName('Virus_YN').AsString          := 'Y'
               else if Virus_N.Checked then
                  ParamByName('Virus_YN').AsString          := 'N'
               else ParamByName('Virus_YN').AsString        := '';
               ParamByName('cod_doctor').AsString           := Copy(cmbCOD_DOCT.Text, 1, 6);
               ParamByName('dat_take').AsString             := dat_take.Text;
               ParamByName('dat_ask').AsString              := Dat_Ask.Text;

               if gum_type1.Checked then
                  ParamByName('gum_type').AsString          := '1'
               else if gum_type2.Checked then
                  ParamByName('gum_type').AsString          := '2'
               else if gum_type3.Checked then               
                  ParamByName('gum_type').AsString          := '3'
               else if gum_type4.Checked then               
                  ParamByName('gum_type').AsString          := '4'
               else ParamByName('gum_type').AsString        := '';
               ParamByName('desc_sokun').AsString           := desc_sokun.Text;
               if GE_Conventional.Checked then
                  ParamByName('GE_Conventional').AsString   := 'Y'
               ELSE ParamByName('GE_Conventional').AsString := 'N';
               if GE_Liquid.Checked then
                  ParamByName('GE_Liquid').AsString         := 'Y'
               ELSE ParamByName('GE_Liquid').AsString       := 'N';
               if GE_Lmp.Checked then
                  ParamByName('GE_Lmp').AsString            := 'Y'
               ELSE ParamByName('GE_Lmp').AsString          := 'N';
               ParamByName('GE_Lmp_text').AsString          := GE_Lmp_text.Text;
               if GE_Menopause.Checked then
                  ParamByName('GE_Menopause').AsString      := 'Y'
               ELSE ParamByName('GE_Menopause').AsString    := 'N';
               ParamByName('GE_Menopause_text').AsString    := GE_Menopause_text.Text;
               if GE_Pregnancy.Checked then
                  ParamByName('GE_Pregnancy').AsString      := 'Y'
               ELSE ParamByName('GE_Pregnancy').AsString    := 'N';
               ParamByName('GE_Pregnancy_text').AsString    := GE_Pregnancy_text.Text;
               if GE_IUD.Checked then
                  ParamByName('GE_IUD').AsString            := 'Y'
               ELSE ParamByName('GE_IUD').AsString          := 'N';
               if GE_Erosion.Checked then
                  ParamByName('GE_Erosion').AsString        := 'Y'
               ELSE ParamByName('GE_Erosion').AsString      := 'N';
               if GE_Hormone.Checked then
                  ParamByName('GE_Hormone').AsString        := 'Y'
               ELSE ParamByName('GE_Hormone').AsString      := 'N';
               if GE_Radiotherapy.Checked then
                  ParamByName('GE_Radiotherapy').AsString   := 'Y'
               ELSE ParamByName('GE_Radiotherapy').AsString := 'N';
               if NGE_Sputum.Checked then
                  ParamByName('NGE_Sputum').AsString        := 'Y'
               ELSE ParamByName('NGE_Sputum').AsString      := 'N';
               if NGE_Pleural.Checked then
                  ParamByName('NGE_Pleural').AsString       := 'Y'
               ELSE ParamByName('NGE_Pleural').AsString     := 'N';
               if NGE_Ascitic.Checked then
                  ParamByName('NGE_Ascitic').AsString       := 'Y'
               ELSE ParamByName('NGE_Ascitic').AsString     := 'N';
               if NGE_Joint.Checked then
                  ParamByName('NGE_Joint').AsString         := 'Y'
               ELSE ParamByName('NGE_Joint').AsString       := 'N';
               if NGE_Urine.Checked then
                  ParamByName('NGE_Urine').AsString         := 'Y'
               ELSE ParamByName('NGE_Urine').AsString       := 'N';
               if NGE_CSF.Checked then
                  ParamByName('NGE_CSF').AsString           := 'Y'
               ELSE ParamByName('NGE_CSF').AsString         := 'N';
               if NGE_Other.Checked then
                  ParamByName('NGE_Other').AsString         := 'Y'
               ELSE ParamByName('NGE_Other').AsString       := 'N';
               if NGE_Cell_block.Checked then
                  ParamByName('NGE_Cell_block').AsString    := 'Y'
               ELSE ParamByName('NGE_Cell_block').AsString  := 'N';
               if FNA_Thyroid.Checked then
                  ParamByName('FNA_Thyroid').AsString       := 'Y'
               ELSE ParamByName('FNA_Thyroid').AsString     := 'N';
               if FNA_Lymph.Checked then
                  ParamByName('FNA_Lymph').AsString         := 'Y'
               ELSE ParamByName('FNA_Lymph').AsString       := 'N';
               if FNA_Breast.Checked then
                  ParamByName('FNA_Breast').AsString        := 'Y'
               ELSE ParamByName('FNA_Breast').AsString      := 'N';
               if FNA_Other.Checked then
                  ParamByName('FNA_Other').AsString         := 'Y'
               ELSE ParamByName('FNA_Other').AsString       := 'N';
               if FNA_Cell_block.Checked then
                  ParamByName('FNA_Cell_block').AsString    := 'Y'
               ELSE ParamByName('FNA_Cell_block').AsString  := 'N';
               ParamByName('DAT_UPDATE').AsString           := GV_sToday;
               ParamByName('COD_UPDATE').AsString           := GV_sUserId;
               IF DuplexUterus_Y.Checked then
                  ParamByName('DuplexUterus_YN').AsString   := 'Y'
               ELSE ParamByName('DuplexUterus_YN').AsString := 'N';
               Execsql;
               GP_query_log(qryU_Ca_Request, 'LD4I02', 'U', 'Y');
            end;
         end
         else
         begin
            with qryI_Ca_Request do
            begin
               ParamByName('cod_jisa').AsString             := UV_vCod_jisa[i];
               ParamByName('NUM_ID').AsString               := UV_vNum_id[i];
               ParamByName('dat_gmgn').AsString             := dat_gmgn.Text;
               ParamByName('dat_human').AsString            := dat_human.Text;
               if Virus_Y.Checked then
                  ParamByName('Virus_YN').AsString          := 'Y'
               else if Virus_N.Checked then
                  ParamByName('Virus_YN').AsString          := 'N'
               else ParamByName('Virus_YN').AsString        := '';
               ParamByName('cod_doctor').AsString           := Copy(cmbCOD_DOCT.Text, 1, 6);
               ParamByName('dat_take').AsString             := dat_take.Text;
               ParamByName('dat_ask').AsString              := Dat_Ask.Text;

               if gum_type1.Checked then                    
                  ParamByName('gum_type').AsString          := '1'
               else if gum_type2.Checked then               
                  ParamByName('gum_type').AsString          := '2'
               else if gum_type3.Checked then               
                  ParamByName('gum_type').AsString          := '3'
               else if gum_type4.Checked then               
                  ParamByName('gum_type').AsString          := '4'
               else ParamByName('gum_type').AsString        := '';
               ParamByName('desc_sokun').AsString           := desc_sokun.Text;
               if GE_Conventional.Checked then
                  ParamByName('GE_Conventional').AsString   := 'Y'
               ELSE ParamByName('GE_Conventional').AsString := 'N';
               if GE_Liquid.Checked then
                  ParamByName('GE_Liquid').AsString         := 'Y'
               ELSE ParamByName('GE_Liquid').AsString       := 'N';
               if GE_Lmp.Checked then
                  ParamByName('GE_Lmp').AsString            := 'Y'
               ELSE ParamByName('GE_Lmp').AsString          := 'N';
               ParamByName('GE_Lmp_text').AsString          := GE_Lmp_text.Text;
               if GE_Menopause.Checked then
                  ParamByName('GE_Menopause').AsString      := 'Y'
               ELSE ParamByName('GE_Menopause').AsString    := 'N';
               ParamByName('GE_Menopause_text').AsString    := GE_Menopause_text.Text;
               if GE_Pregnancy.Checked then
                  ParamByName('GE_Pregnancy').AsString      := 'Y'
               ELSE ParamByName('GE_Pregnancy').AsString    := 'N';
               ParamByName('GE_Pregnancy_text').AsString    := GE_Pregnancy_text.Text;
               if GE_IUD.Checked then
                  ParamByName('GE_IUD').AsString            := 'Y'
               ELSE ParamByName('GE_IUD').AsString          := 'N';
               if GE_Erosion.Checked then
                  ParamByName('GE_Erosion').AsString        := 'Y'
               ELSE ParamByName('GE_Erosion').AsString      := 'N';
               if GE_Hormone.Checked then
                  ParamByName('GE_Hormone').AsString        := 'Y'
               ELSE ParamByName('GE_Hormone').AsString      := 'N';
               if GE_Radiotherapy.Checked then
                  ParamByName('GE_Radiotherapy').AsString   := 'Y'
               ELSE ParamByName('GE_Radiotherapy').AsString := 'N';
               if NGE_Sputum.Checked then
                  ParamByName('NGE_Sputum').AsString        := 'Y'
               ELSE ParamByName('NGE_Sputum').AsString      := 'N';
               if NGE_Pleural.Checked then
                  ParamByName('NGE_Pleural').AsString       := 'Y'
               ELSE ParamByName('NGE_Pleural').AsString     := 'N';
               if NGE_Ascitic.Checked then
                  ParamByName('NGE_Ascitic').AsString       := 'Y'
               ELSE ParamByName('NGE_Ascitic').AsString     := 'N';
               if NGE_Joint.Checked then
                  ParamByName('NGE_Joint').AsString         := 'Y'
               ELSE ParamByName('NGE_Joint').AsString       := 'N';
               if NGE_Urine.Checked then
                  ParamByName('NGE_Urine').AsString         := 'Y'
               ELSE ParamByName('NGE_Urine').AsString       := 'N';
               if NGE_CSF.Checked then
                  ParamByName('NGE_CSF').AsString           := 'Y'
               ELSE ParamByName('NGE_CSF').AsString         := 'N';
               if NGE_Other.Checked then
                  ParamByName('NGE_Other').AsString         := 'Y'
               ELSE ParamByName('NGE_Other').AsString       := 'N';
               if NGE_Cell_block.Checked then
                  ParamByName('NGE_Cell_block').AsString    := 'Y'
               ELSE ParamByName('NGE_Cell_block').AsString  := 'N';
               if FNA_Thyroid.Checked then
                  ParamByName('FNA_Thyroid').AsString       := 'Y'
               ELSE ParamByName('FNA_Thyroid').AsString     := 'N';
               if FNA_Lymph.Checked then
                  ParamByName('FNA_Lymph').AsString         := 'Y'
               ELSE ParamByName('FNA_Lymph').AsString       := 'N';
               if FNA_Breast.Checked then
                  ParamByName('FNA_Breast').AsString        := 'Y'
               ELSE ParamByName('FNA_Breast').AsString      := 'N';
               if FNA_Other.Checked then
                  ParamByName('FNA_Other').AsString         := 'Y'
               ELSE ParamByName('FNA_Other').AsString       := 'N';
               if FNA_Cell_block.Checked then
                  ParamByName('FNA_Cell_block').AsString    := 'Y'
               ELSE ParamByName('FNA_Cell_block').AsString  := 'N';
               ParamByName('DAT_insert').AsString           := GV_sToday;
               ParamByName('COD_insert').AsString           := GV_sUserId;
               IF DuplexUterus_Y.Checked then
                  ParamByName('DuplexUterus_YN').AsString   := 'Y'
               ELSE ParamByName('DuplexUterus_YN').AsString := 'N';
               Execsql;
               GP_query_log(qryI_Ca_Request, 'LD4I02', 'I', 'Y');
            end;
         end;
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
      UV_vdat_human[i]         := dat_human.Text;
      if Virus_Y.Checked then
         UV_vVirus_YN[i]  := 'Y'
      else if Virus_N.Checked then
         UV_vVirus_YN[i]  := 'N'
      else UV_vVirus_YN[i]  := '';
      UV_vcod_doctor[i]        := Copy(cmbCOD_DOCT.Text, 1, 6);
      UV_vDat_take[i]          := dat_take.Text;
      UV_vDat_Ask[i]           := Dat_Ask.Text;
      if gum_type1.Checked then
         UV_vGum_type[i] := '1'
      else if gum_type2.Checked then
         UV_vGum_type[i] := '2'
      else if gum_type3.Checked then
         UV_vGum_type[i] := '3'
      else if gum_type4.Checked then
         UV_vGum_type[i] := '4';
      UV_vDesc_sokun[i]           := desc_sokun.Text;
      if GE_Conventional.Checked then
         UV_vGE_Conventional[i]   := 'Y'
      else UV_vGE_Conventional[i] := 'N';
      if GE_Liquid.Checked then
         UV_vGE_Liquid[i]         := 'Y'
      else UV_vGE_Liquid[i]       := 'N';
      if GE_Lmp.Checked then
         UV_vGE_Lmp[i]            := 'Y'
      else UV_vGE_Lmp[i]          := 'N';
      UV_vGE_Lmp_text[i]          := GE_Lmp_text.Text;
      if GE_Menopause.Checked then
         UV_vGE_Menopause[i]      := 'Y'
      else UV_vGE_Menopause[i]    := 'N';
      UV_vGE_Menopause_text[i]    := GE_Menopause_text.Text;
      if GE_Pregnancy.Checked then
         UV_vGE_Pregnancy[i]      := 'Y'
      else UV_vGE_Pregnancy[i]    := 'N';
      UV_vGE_Pregnancy_text[i]    := GE_Pregnancy_text.Text;
      if GE_IUD.Checked then
         UV_vGE_IUD[i]            := 'Y'
      else UV_vGE_IUD[i]          := 'N';
      if GE_Erosion.Checked then
         UV_vGE_Erosion[i]        := 'Y'
      else UV_vGE_Erosion[i]      := 'N';
      if GE_Hormone.Checked then
         UV_vGE_Hormone[i]        := 'Y'
      else UV_vGE_Hormone[i]      := 'N';
      if GE_Radiotherapy.Checked then
         UV_vGE_Radiotherapy[i]   := 'Y'
      else UV_vGE_Radiotherapy[i] := 'N';
      if NGE_Sputum.Checked then
         UV_vNGE_Sputum[i]        := 'Y'
      else UV_vNGE_Sputum[i]      := 'N';
      if NGE_Pleural.Checked then
         UV_vNGE_Pleural[i]       := 'Y'
      else UV_vNGE_Pleural[i]     := 'N';
      if NGE_Ascitic.Checked then
         UV_vNGE_Ascitic[i]       := 'Y'
      else UV_vNGE_Ascitic[i]     := 'N';
      if NGE_Joint.Checked then
         UV_vNGE_Joint[i]         := 'Y'
      else UV_vNGE_Joint[i]       := 'N';
      if NGE_Urine.Checked then
         UV_vNGE_Urine[i]         := 'Y'
      else UV_vNGE_Urine[i]       := 'N';
      if NGE_CSF.Checked then
         UV_vNGE_CSF[i]           := 'Y'
      else UV_vNGE_CSF[i]         := 'N';
      if NGE_Other.Checked then
         UV_vNGE_Other[i]         := 'Y'
      else UV_vNGE_Other[i]       := 'N';
      if NGE_Cell_block.Checked then
         UV_vNGE_Cell_block[i]    := 'Y'
      else UV_vNGE_Cell_block[i]  := 'N';
      if FNA_Thyroid.Checked then
         UV_vFNA_Thyroid[i]       := 'Y'
      else UV_vFNA_Thyroid[i]     := 'N';
      if FNA_Lymph.Checked then
         UV_vFNA_Lymph[i]         := 'Y'
      else UV_vFNA_Lymph[i]       := 'N';
      if FNA_Breast.Checked then
         UV_vFNA_Breast[i]        := 'Y'
      else UV_vFNA_Breast[i]      := 'N';
      if FNA_Other.Checked then
         UV_vFNA_Other[i]         := 'Y'
      else UV_vFNA_Other[i]       := 'N';
      if FNA_Cell_block.Checked then
         UV_vFNA_Cell_block[i]    := 'Y'
      else UV_vFNA_Cell_block[i]  := 'N';
      // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;
   end;
   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLD4I02.btnSaveClick(Sender: TObject);
begin
   inherited;
   If dat_take.Text = '' Then
   Begin
      Showmessage('☞ 검체채취일을 입력하세요.');
      Exit;
   End;
   UF_Save;
end;

procedure TfrmLD4I02.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);

   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   pnlJisa.Visible := True;                                                     //06.11.24 철. 지사선택 보이게끔.
   GP_ComboSawonAll(cmbCOD_DOCT, 1);

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   cmbRelation.Enabled := False;

   UV_iCod_doct := -1;
   cmbGubn.ItemIndex := 1;
   mskDate.Text := GV_sToday;
   Cmb_Order.ItemIndex := 0;
   dat_human.Text  := '황인종';
end;

procedure TfrmLD4I02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_Jisa[DataRow-1];
      2 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      3 : Value := UV_vDesc_name[DataRow-1];
      4 : Value := UV_vCod_hm[DataRow-1];
   end;

   // 조직검사일 경우 색깔 변경...
{   if (Copy(UV_vCod_cell[DataRow-1],1,1) = 'J') then
      grdMaster.CellColor[DataCol, DataRow] := $00F9D9F5;      }
end;

procedure TfrmLD4I02.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrderby : String;
begin
   inherited;
   UV_bEdit := True;
   btnPsave.Enabled   := True;
   btnPcancel.Enabled := True;
   sFix_Doctor := ''; CB_fix_Doc.Checked := False;
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

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid Setting
   UP_GridSet;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryGumgin do
   begin
      Active := False;

      sSelect := ' SELECT G.*, R.*, C.cod_hm, J.desc_jisa, D.desc_dc, C.desc_buwi FROM gumgin G left outer join ca_request R ';
      sSelect := sSelect + ' On G.cod_jisa = R.cod_jisa ';
      sSelect := sSelect + ' and G.num_id = R.num_id ';
      sSelect := sSelect + ' and G.dat_gmgn = R.dat_gmgn ';
      sSelect := sSelect + ' left outer join Cell C ';
      sSelect := sSelect + ' On G.cod_jisa = C.cod_jisa ';
      sSelect := sSelect + ' and G.num_id = C.num_id ';
      sSelect := sSelect + ' and G.dat_gmgn = C.dat_gmgn ';
      //PAP
      if cmbGubn.ItemIndex = 1 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P001'' or C.cod_hm = ''P003'' or C.cod_hm = ''P006'' or C.cod_hm = ''P009'' or C.cod_hm = ''P010'' or C.cod_hm = ''P011'' ) ';
      end
      //조직
      else if cmbGubn.ItemIndex = 2 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''B009'') ';
      end
      else if cmbGubn.ItemIndex = 3 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P012'') ';
      end
      else if cmbGubn.ItemIndex = 4 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P013'') ';
      end
      else
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P001'' or C.cod_hm = ''P003'' or C.cod_hm = ''P006'' or C.cod_hm = ''P009'' or C.cod_hm = ''P010'' or C.cod_hm = ''P011'' or C.cod_hm = ''P012'' or C.cod_hm = ''P013'' ';
         sSelect := sSelect + ' or C.cod_hm = ''B001'' or C.cod_hm = ''B007'' or C.cod_hm = ''B008'' or C.cod_hm = ''B009'') ';
      end;
      sSelect := sSelect + ' left outer join profile_hm P ';
      sSelect := sSelect + ' On  G.cod_jangbi = P.cod_pf ';
      //PAP
      if cmbGubn.ItemIndex = 1 then
      begin
         sSelect := sSelect + ' and (P.cod_hm = ''P001'' or P.cod_hm = ''P003'' or P.cod_hm = ''P006'' or P.cod_hm = ''P009'' or P.cod_hm = ''P010'' or P.cod_hm = ''P011'' ) ';
      end
      //조직
      else if cmbGubn.ItemIndex = 2 then
      begin
         sSelect := sSelect + ' and (P.cod_hm = ''B009'') ';
      end
      else if cmbGubn.ItemIndex = 3 then
      begin
         sSelect := sSelect + ' and (P.cod_hm = ''P012'') ';
      end
      else if cmbGubn.ItemIndex = 4 then
      begin
         sSelect := sSelect + ' and (P.cod_hm = ''P013'') ';
      end
      else
      begin
         sSelect := sSelect + ' and (P.cod_hm = ''P001'' or P.cod_hm = ''P003'' or P.cod_hm = ''P006'' or P.cod_hm = ''P009'' or P.cod_hm = ''P010'' or P.cod_hm = ''P011'' or P.cod_hm = ''P012'' or P.cod_hm = ''P013'' ';
         sSelect := sSelect + ' or P.cod_hm = ''B001'' or P.cod_hm = ''B007'' or P.cod_hm = ''B008'' or P.cod_hm = ''B009'' ) ';
      end;                                                                                          
      sSelect := sSelect + ' inner join Jisa J ';
      sSelect := sSelect + ' On G.cod_jisa = J.cod_jisa ';
      sSelect := sSelect + ' inner join Danche D ';
      sSelect := sSelect + ' On G.cod_dc = D.cod_dc ';

      sWhere := 'WHERE ';
      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
            sWhere := sWhere + ' G.dat_gmgn = ''' + mskDate.Text + ''' AND ';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' G.num_jumin = ''' + mskJumin.Text + ''' And ';
      end
      else if rbCell.Checked then
      begin
         sWhere := sWhere + ' C.desc_buwi = ''' + mskCell.Text + ''' And ';
      end;
      //PAP
      if cmbGubn.ItemIndex = 1 then
      begin
         sWhere := sWhere + ' ( P.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P001%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P003%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P006%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P009%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P010%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P011%'' ) ';
      end
      else if cmbGubn.ItemIndex = 2 then
      begin
         sWhere := sWhere + ' ( P.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%B009%'' ) ';
      end
      else if cmbGubn.ItemIndex = 3 then
      begin
        { sWhere := sWhere + ' ( P.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P012%'' ) ';    }
          sWhere := sWhere + '  C.cod_hm = ''P012''  ';
      end
      else if cmbGubn.ItemIndex = 4 then
      begin
         {sWhere := sWhere + ' ( P.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P013%'' ) ';  } //p012, p013은 셀테이블이 생성되면 리스트업 ..20170428 본원 최윤선
         sWhere := sWhere + '  C.cod_hm = ''P013''  ';
      end
      else
      begin
         sWhere := sWhere + ' ( P.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P001%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P003%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P006%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P009%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P010%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P011%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P012%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P013%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%B001%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%B007%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%B008%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%B009%'' )';
      end;

      if Cmb_Order.ItemIndex = 1 then
         sOrderBy := ' ORDER BY G.dat_gmgn, G.cod_jisa, G.num_samp'
      else if Cmb_Order.ItemIndex = 2 then
         sOrderBy := ' ORDER BY G.dat_gmgn, G.cod_jisa, G.num_jumin'
      else
         sOrderBy := ' ORDER BY G.dat_gmgn, G.cod_jisa, G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      if qryGumgin.IsEmpty = False then
      begin
         GP_query_log(qryGumgin, 'LD4I02', 'Q', 'N');
         while Not Eof do
         begin
            UP_VarMemSet(index);
            UV_vCod_jisa[index]          := FieldByName('COD_jisa').AsString;
            UV_vNum_id[index]            := FieldByName('NUM_ID').AsString;
            UV_vNum_jumin[index]         := FieldByName('NUM_JUMIN').AsString;
            UV_vDesc_name[index]         := FieldByName('DESC_NAME').AsString;
            UV_vDat_gmgn[index]          := FieldByName('dat_gmgn').AsString;
            UV_vDesc_buwi[index]         := FieldByName('Desc_buwi').AsString;
            UV_vdat_human[index]         := FieldByName('dat_human').AsString;
            UV_vVirus_YN[index]          := FieldByName('Virus_YN').AsString;
            UV_vcod_doctor[index]        := FieldByName('cod_doctor').AsString;
            UV_vDat_take[index]          := FieldByName('Dat_take').AsString;
            UV_vDat_Ask[index]           := FieldByName('Dat_ask').AsString;
            UV_vDesc_jisa[index]         := FieldByName('Desc_jisa').AsString;
            UV_vDesc_dc[index]           := FieldByName('Desc_dc').AsString;
            UV_vGum_type[index]          := FieldByName('Gum_type').AsString;
            UV_vDesc_sokun[index]        := FieldByName('Desc_sokun').AsString;
            UV_vGE_Conventional[index]   := FieldByName('GE_Conventional').AsString;
            UV_vGE_Liquid[index]         := FieldByName('GE_Liquid').AsString;
            UV_vGE_Lmp[index]            := FieldByName('GE_Lmp').AsString;
            UV_vGE_Lmp_text[index]       := FieldByName('GE_Lmp_text').AsString;
            UV_vGE_Menopause[index]      := FieldByName('GE_Menopause').AsString;
            UV_vGE_Menopause_text[index] := FieldByName('GE_Menopause_text').AsString;
            UV_vGE_Pregnancy[index]      := FieldByName('GE_Pregnancy').AsString;
            UV_vGE_Pregnancy_text[index] := FieldByName('GE_Pregnancy_text').AsString;
            UV_vGE_IUD[index]            := FieldByName('GE_IUD').AsString;
            UV_vGE_Erosion[index]        := FieldByName('GE_Erosion').AsString;
            UV_vGE_Hormone[index]        := FieldByName('GE_Hormone').AsString;
            UV_vGE_Radiotherapy[index]   := FieldByName('GE_Radiotherapy').AsString;
            UV_vNGE_Sputum[index]        := FieldByName('NGE_Sputum').AsString;
            UV_vNGE_Pleural[index]       := FieldByName('NGE_Pleural').AsString;
            UV_vNGE_Ascitic[index]       := FieldByName('NGE_Ascitic').AsString;
            UV_vNGE_Joint[index]         := FieldByName('NGE_Joint').AsString;
            UV_vNGE_Urine[index]         := FieldByName('NGE_Urine').AsString;
            UV_vNGE_CSF[index]           := FieldByName('NGE_CSF').AsString;
            UV_vNGE_Other[index]         := FieldByName('NGE_Other').AsString;
            UV_vNGE_Cell_block[index]    := FieldByName('NGE_Cell_block').AsString;
            UV_vFNA_Thyroid[index]       := FieldByName('FNA_Thyroid').AsString;
            UV_vFNA_Lymph[index]         := FieldByName('FNA_Lymph').AsString;
            UV_vFNA_Breast[index]        := FieldByName('FNA_Breast').AsString;
            UV_vFNA_Other[index]         := FieldByName('FNA_Other').AsString;
            UV_vFNA_Cell_block[index]    := FieldByName('FNA_Cell_block').AsString;
            UV_vDuplexUterus_YN[index]   := FieldByName('DuplexUterus_YN').AsString;

            UV_vCod_hm[index]            := FieldByName('COD_hm').AsString;
            {                                                                   //2014.07.28 곽윤설
            if ((pos( 'P001', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P001')) then
               UV_vCod_hm[index]   := 'P001'
            else if ((pos( 'P003', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P003')) then
               UV_vCod_hm[index]   := 'P003'
            else if ((pos( 'P006', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P006')) then
               UV_vCod_hm[index]   := 'P006'
            else if ((pos( 'P009', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P009')) then
               UV_vCod_hm[index]   := 'P009'
            else if ((pos( 'P010', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P010') ) then
               UV_vCod_hm[index]   := 'P010'
            else if ((pos( 'P011', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'P011') ) then
               UV_vCod_hm[index]   := 'P011'
            else if ((pos( 'B001', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'B001') ) then
               UV_vCod_hm[index]   := 'B001'
            else if ((pos( 'B007', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'B007') ) then
               UV_vCod_hm[index]   := 'B007'
            else if ((pos( 'B008', FieldByName('cod_chuga').AsString ) > 0)
                or (FieldByName('cod_hm').AsString = 'B008') ) then
               UV_vCod_hm[index]   := 'B008';
            }
            Next;
            Inc(index);
         end;

         // Table과 Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GP_FieldClear(Panel2);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4I02.btnCancelClick(Sender: TObject);
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

procedure TfrmLD4I02.UP_Change(Sender: TObject);
begin
   inherited;
   // Edit Flag Check
   UV_bEdit := True;
   if CB_fix_Doc.Checked then
      sFix_doctor := cmbCOD_DOCT.Text
   else sFix_doctor := '';
end;

procedure TfrmLD4I02.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var sMan : String;
    iAge : integer;
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   Desc_buwi.Text    := UV_vDesc_buwi[NewRow-1];
   desc_name.Text    := UV_vdesc_name[NewRow-1];
   num_jumin.Text    := UV_vnum_jumin[NewRow-1];
   Edt_Num_id.Text   := UV_vNum_id[NewRow-1];
   GP_JuminToAge(UV_vnum_jumin[NewRow-1],UV_vdat_gmgn[NewRow-1], sMan, iAge);
   age_sex.Text      := intToStr(iAge) + ' / ' + sMan;
   desc_jisa.Text    := UV_vdesc_jisa[NewRow-1];
   desc_dc.Text      := UV_vdesc_dc[NewRow-1];
   dat_gmgn.Text     := UV_vdat_gmgn[NewRow-1];
   GP_JuminToAge(num_jumin.Text,dat_gmgn.Text, sMan, iAge);
   
   if UV_vdat_human[NewRow-1] <> '' then
      dat_human.Text  := UV_vdat_human[NewRow-1]
   else if UV_vdat_human[NewRow-1] = '' then
      dat_human.Text  := '황인종';
   if UV_vDuplexUterus_YN[NewRow-1] = 'Y' then
      DuplexUterus_Y.Checked := True
   else if UV_vDuplexUterus_YN[NewRow-1] = 'N' then
      DuplexUterus_N.Checked := True
   else
      DuplexUterus_N.Checked := True;


   if UV_vVirus_YN[NewRow-1] = 'Y' then
      Virus_Y.Checked := True
   else if UV_vVirus_YN[NewRow-1] = 'N' then
      Virus_N.Checked := True
   else
      Virus_N.Checked := True;

   if UV_vcod_doctor[NewRow-1] <> '' then                                       //2014.11.04 곽윤설 ~
   begin
      GP_ComboDisplay(cmbCOD_Doct, UV_vcod_doctor[NewRow-1], 6);
      cmbCOD_DOCT.Font.Color := clGrayText;
   end
   else
   begin
      if CB_fix_Doc.Checked then
         GP_ComboDisplay(cmbCOD_Doct, copy(sFix_doctor,1,6), 6)
      else
         GP_ComboDisplay(cmbCOD_Doct, GV_sUserId, 6);
      cmbCOD_DOCT.Font.Color := clWindowText;
   end;
   cmbCOD_DOCT.setfocus;
   desc_sokun.setfocus;                                                         //~ 2014.11.04 곽윤설

   //저장된 값이 없는경우 검진일자표기
   if UV_vdat_take[NewRow-1] <> '' then
      dat_take.Text     := UV_vdat_take[NewRow-1]
   else
      dat_take.Text     := dat_gmgn.Text;

   if UV_vdat_ask[NewRow-1] <> '' then
      Dat_Ask.Text      := UV_vdat_ask[NewRow-1]
   else
      Dat_Ask.Text      := dat_gmgn.Text;

   gum_type1.Checked := False;
   gum_type2.Checked := False;
   gum_type3.Checked := False;
   gum_type4.Checked := False;
   if UV_vGum_type[NewRow-1] = '1' then
      gum_type1.Checked := True
   else if UV_vGum_type[NewRow-1] = '2' then
      gum_type2.Checked := True
   else if UV_vGum_type[NewRow-1] = '3' then
      gum_type3.Checked := True
   else if UV_vGum_type[NewRow-1] = '4' then
      gum_type4.Checked := True;         
   desc_sokun.Text   := UV_vDesc_sokun[NewRow-1];
   if UV_vGE_Conventional[NewRow-1] = 'Y' then
      GE_Conventional.Checked   := True
   else GE_Conventional.Checked := False;
   if UV_vGE_Conventional[NewRow-1] = 'Y' then
      GE_Conventional.Checked   := True
   else GE_Conventional.Checked := False;
   if UV_vGE_Liquid[NewRow-1] = 'Y' then
      GE_Liquid.Checked   := True
   else GE_Liquid.Checked := False;
   if UV_vGE_Lmp[NewRow-1] = 'Y' then
      GE_Lmp.Checked   := True
   else GE_Lmp.Checked := False;
   GE_Lmp_text.Text    := UV_vGE_Lmp_text[NewRow-1];
   if UV_vGE_Menopause[NewRow-1] = 'Y' then
      GE_Menopause.Checked   := True
   else GE_Menopause.Checked := False;
   GE_Menopause_text.Text    := UV_vGE_Menopause_text[NewRow-1];
   if UV_vGE_Pregnancy[NewRow-1] = 'Y' then
      GE_Pregnancy.Checked   := True
   else GE_Pregnancy.Checked := False;
   GE_Pregnancy_text.Text    := UV_vGE_Pregnancy_text[NewRow-1];
   if UV_vGE_IUD[NewRow-1] = 'Y' then
      GE_IUD.Checked   := True
   else GE_IUD.Checked := False;
   if UV_vGE_Erosion[NewRow-1] = 'Y' then
      GE_Erosion.Checked   := True
   else GE_Erosion.Checked := False;
   if UV_vGE_Hormone[NewRow-1] = 'Y' then
      GE_Hormone.Checked   := True
   else GE_Hormone.Checked := False;
   if UV_vGE_Radiotherapy[NewRow-1] = 'Y' then
      GE_Radiotherapy.Checked   := True
   else GE_Radiotherapy.Checked := False;
   if UV_vNGE_Sputum[NewRow-1] = 'Y' then
      NGE_Sputum.Checked   := True
   else NGE_Sputum.Checked := False;
   if UV_vNGE_Pleural[NewRow-1] = 'Y' then
      NGE_Pleural.Checked   := True
   else NGE_Pleural.Checked := False;
   if UV_vNGE_Ascitic[NewRow-1] = 'Y' then
      NGE_Ascitic.Checked   := True
   else NGE_Ascitic.Checked := False;

   if UV_vNGE_Joint[NewRow-1] = 'Y' then
      NGE_Joint.Checked   := True
   else NGE_Joint.Checked := False;
   if UV_vNGE_Urine[NewRow-1] = 'Y' then
      NGE_Urine.Checked   := True
   else NGE_Urine.Checked := False;
   if UV_vNGE_CSF[NewRow-1] = 'Y' then
      NGE_CSF.Checked   := True
   else NGE_CSF.Checked := False;
   if UV_vNGE_Other[NewRow-1] = 'Y' then
      NGE_Other.Checked   := True
   else NGE_Other.Checked := False;
   if UV_vNGE_Cell_block[NewRow-1] = 'Y' then
      NGE_Cell_block.Checked   := True
   else NGE_Cell_block.Checked := False;
   if UV_vFNA_Thyroid[NewRow-1] = 'Y' then
      FNA_Thyroid.Checked   := True
   else FNA_Thyroid.Checked := False;
   if UV_vFNA_Lymph[NewRow-1] = 'Y' then
      FNA_Lymph.Checked   := True
   else FNA_Lymph.Checked := False;
   if UV_vFNA_Breast[NewRow-1] = 'Y' then
      FNA_Breast.Checked   := True
   else FNA_Breast.Checked := False;
   if UV_vFNA_Other[NewRow-1] = 'Y' then
      FNA_Other.Checked   := True
   else FNA_Other.Checked := False;
   if UV_vFNA_Cell_block[NewRow-1] = 'Y' then
      FNA_Cell_block.Checked   := True
   else FNA_Cell_block.Checked := False;
   if UV_vDuplexUterus_YN[NewRow-1] = 'Y' then
      DuplexUterus_Y.Checked   := True

   else DuplexUterus_N.Checked := True;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD4I02.UP_Click(Sender: TObject);
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
   else if Sender = btnDat_Take   then GF_CalendarClick(dat_take)
   else if Sender = btnDat_Ask    then GF_CalendarClick(Dat_Ask);
end;

procedure TfrmLD4I02.rbDateClick(Sender: TObject);
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

procedure TfrmLD4I02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // Reference Key가 아니면 종료
   if Key <> GC_Refer then exit;

   if      Sender = mskDate    then UP_Click(btnDate)
   else if Sender = mskJumin   then UP_Click(btnJumin)
   else if Sender = dat_take   then UP_Click(btnDat_Take)
   else if Sender = Dat_Ask    then UP_Click(btnDat_Ask);
end;

procedure TfrmLD4I02.UP_Exit(Sender: TObject);
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

procedure TfrmLD4I02.sPrintClick(Sender: TObject);
begin
  inherited;
  Try
     Begin
        frmLD4I021 := TfrmLD4I021.Create(Self);
        frmLD4I021.QuickRep.Preview;
     End
  Finally
        frmLD4I021.Free
  End;

end;

procedure TfrmLD4I02.btnDeleteClick(Sender: TObject);
begin
  inherited;
  // Delete Confirm Message
   if (grdMaster.Rows = 0) or (grdMaster.VisibleRows < 1) then exit;
   if not GF_MsgBox('Confirmation', '해당 자료를  삭제하면 복구가 불가능합니다.' + Chr(13) +
                                    '정말로 삭제하시겠습니까 ?') then exit;

   // Delete 작업수행
   try
      with qryD_Ca_Request do
      begin
         ParamByName('COD_JISA').AsString := UV_vCod_jisa[grdMaster.CurrentDataRow-1];
         ParamByName('NUM_ID').AsString   := UV_vNum_id[grdMaster.CurrentDataRow-1];         
         ParamByName('DAT_GMGN').AsString := dat_gmgn.Text;

         ExecSql;
         GP_query_log(qryD_Ca_Request, 'LD4I02', 'Q', 'Y');
      end;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         exit;
      end;
   end;
   {
   // 화면에서 안보이게함
   grdMaster.RowVisible[grdMaster.CurrentDataRow] := False;
   if grdMaster.VisibleRows < 1 then GP_FieldClear(pnlDetail);
   }
   // Delete Mode & Msg
   UP_SetMode('DELETE');
end;

end.
