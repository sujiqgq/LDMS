//==============================================================================
// 프로그램 설명 : PB Smear결과관리
// 시스템        : LDMS
// 수정일자      : 2011.09.14
// 수정자        : 구수정
// 수정내용      :
//==============================================================================
// 수정일자      : 2012.05.04
// 수정자        : 구수정
// 수정내용      :  H018~H022 0일경우 공란으로 혈액결과에 업데이트
//==============================================================================

unit LD4I03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD4I03 = class(TfrmSingle)
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
    qryPBS: TQuery;
    Panel2: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    dat_gmgn: TDateEdit;
    desc_name: TEdit;
    age_sex: TEdit;
    Panel_RBC: TPanel;
    Panel15: TPanel;
    Label3: TLabel;
    qryI_PBS: TQuery;
    qry_PBS: TQuery;
    qryU_PBS: TQuery;
    qryUser_priv: TQuery;
    qryGmgn: TQuery;
    Panel5: TPanel;
    Num_bunju: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    RBC_inc: TEdit;
    Panel_WBC: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Panel4: TPanel;
    WBC_num: TEdit;
    Label18: TLabel;
    WBC_bla: TEdit;
    Label19: TLabel;
    WBC_pro: TEdit;
    Label20: TLabel;
    WBC_mye: TEdit;
    Label21: TLabel;
    WBC_met: TEdit;
    Label22: TLabel;
    WBC_ban: TEdit;
    Label23: TLabel;
    WBC_seg: TEdit;
    Label24: TLabel;
    WBC_lym: TEdit;
    Label25: TLabel;
    WBC_mon: TEdit;
    Label26: TLabel;
    WBC_eos: TEdit;
    Label27: TLabel;
    WBC_bas: TEdit;
    Label28: TLabel;
    WBC_imm: TEdit;
    Label29: TLabel;
    WBC_rbc: TEdit;
    Panel_PLT: TPanel;
    Label35: TLabel;
    Panel8: TPanel;
    PLT_size: TEdit;
    Label31: TLabel;
    PLT_num: TEdit;
    Panel12: TPanel;
    Panel16: TPanel;
    Opinion: TMemo;
    btn_sign: TBitBtn;
    Panel6: TPanel;
    Dat_bunju: TDateEdit;
    qryNo_hangmok: TQuery;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    Qry_CK_PBS: TQuery;
    Qry_Sign: TQuery;
    QryI_Sign: TQuery;
    Panel3: TPanel;
    RBC_size_1: TRadioButton;
    RBC_size_2: TRadioButton;
    RBC_size_3: TRadioButton;
    Panel7: TPanel;
    RBC_stain_1: TRadioButton;
    RBC_stain_2: TRadioButton;
    RBC_stain_3: TRadioButton;
    RBC_stain_4: TRadioButton;
    Panel13: TPanel;
    RBC_ani_1: TRadioButton;
    RBC_ani_2: TRadioButton;
    RBC_ani_3: TRadioButton;
    RBC_ani_4: TRadioButton;
    Panel14: TPanel;
    RBC_poi_1: TRadioButton;
    RBC_poi_2: TRadioButton;
    RBC_poi_3: TRadioButton;
    RBC_poi_4: TRadioButton;
    Panel17: TPanel;
    WBC_gra_1: TRadioButton;
    WBC_gra_2: TRadioButton;
    WBC_gra_3: TRadioButton;
    WBC_gra_4: TRadioButton;
    Panel18: TPanel;
    WBC_vac_1: TRadioButton;
    WBC_vac_2: TRadioButton;
    WBC_vac_3: TRadioButton;
    WBC_vac_4: TRadioButton;
    Panel19: TPanel;
    WBC_doh_1: TRadioButton;
    WBC_doh_2: TRadioButton;
    Panel21: TPanel;
    WBC_seg_1: TRadioButton;
    WBC_seg_2: TRadioButton;
    WBC_seg_3: TRadioButton;
    BitBtn1: TBitBtn;
    qryU_glkwa: TQuery;
    qryHul: TQuery;
    Panel20: TPanel;
    cmbCOD_DOCT: TComboBox;
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
    procedure btn_signClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_bjjs,  UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn,
    UV_vnum_bunju, UV_vDat_bunju, UV_vCod_jisa,
    UV_vRBC_size,  UV_vRBC_stainability, UV_vRBC_anisocytosis,
    UV_vRBC_poikilocytosis, UV_vRBC_inclusions,
    UV_vWBC_number, UV_vWBC_toxic_granules, UV_vWBC_toxic_vacuoles,
    UV_vWBC_toxic_dohlebody, UV_vWBC_segmentation,
    UV_vWBC_diff_blast, UV_vWBC_diff_promyelo, UV_vWBC_diff_myelo,
    UV_vWBC_diff_meta, UV_vWBC_diff_band, UV_vWBC_diff_seg,
    UV_vWBC_diff_lym, UV_vWBC_diff_mono, UV_vWBC_diff_eos,
    UV_vWBC_diff_baso, UV_vWBC_diff_immaturecell, UV_vWBC_diff_nRBC,
    UV_vPLT_number, UV_vPLT_size, UV_vCod_doct,
    UV_vOpinion : Variant;

    // Select문
    UV_sBasicSQL : String;

    // 작업 지사코드
    UV_sCod_jisa : String;
    UV_iCod_doct : Integer;
    UV_vstatus   : Boolean;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : Boolean;
    function  UF_Save : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4I03: TfrmLD4I03;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

//마우스휠 적용
procedure TfrmLD4I03.MouseWheelHandler(var Message: TMessage);
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

procedure TfrmLD4I03.UP_GridSet;
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

   end;
end;

procedure TfrmLD4I03.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_bjjs               := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id                 := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin              := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name              := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn               := VarArrayCreate([0, 0], varOleStr);
      UV_vnum_bunju              := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju              := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa               := VarArrayCreate([0, 0], varOleStr);
      UV_vRBC_size               := VarArrayCreate([0, 0], varOleStr);
      UV_vRBC_stainability       := VarArrayCreate([0, 0], varOleStr);
      UV_vRBC_anisocytosis       := VarArrayCreate([0, 0], varOleStr);
      UV_vRBC_poikilocytosis     := VarArrayCreate([0, 0], varOleStr);
      UV_vRBC_inclusions         := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_number             := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_toxic_granules     := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_toxic_vacuoles     := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_toxic_dohlebody    := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_segmentation       := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_blast         := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_promyelo      := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_myelo         := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_meta          := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_band          := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_seg           := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_lym           := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_mono          := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_eos           := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_baso          := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_immaturecell  := VarArrayCreate([0, 0], varOleStr);
      UV_vWBC_diff_nRBC          := VarArrayCreate([0, 0], varOleStr);
      UV_vPLT_number             := VarArrayCreate([0, 0], varOleStr);
      UV_vPLT_size               := VarArrayCreate([0, 0], varOleStr);
      UV_vOpinion                := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doct               := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_bjjs,              iLength);
      VarArrayRedim(UV_vNum_id,                iLength);
      VarArrayRedim(UV_vNum_jumin,             iLength);
      VarArrayRedim(UV_vDesc_name,             iLength);
      VarArrayRedim(UV_vDat_gmgn,              iLength);
      VarArrayRedim(UV_vnum_bunju,             iLength);
      VarArrayRedim(UV_vDat_bunju,             iLength);
      VarArrayRedim(UV_vCod_jisa,              iLength);
      VarArrayRedim(UV_vRBC_size,              iLength);
      VarArrayRedim(UV_vRBC_stainability,      iLength);
      VarArrayRedim(UV_vRBC_anisocytosis,      iLength);
      VarArrayRedim(UV_vRBC_poikilocytosis,    iLength);
      VarArrayRedim(UV_vRBC_inclusions,        iLength);
      VarArrayRedim(UV_vWBC_number,            iLength);
      VarArrayRedim(UV_vWBC_toxic_granules,    iLength);
      VarArrayRedim(UV_vWBC_toxic_vacuoles,    iLength);
      VarArrayRedim(UV_vWBC_toxic_dohlebody,   iLength);
      VarArrayRedim(UV_vWBC_segmentation,      iLength);
      VarArrayRedim(UV_vWBC_diff_blast,        iLength);
      VarArrayRedim(UV_vWBC_diff_promyelo,     iLength);
      VarArrayRedim(UV_vWBC_diff_myelo,        iLength);
      VarArrayRedim(UV_vWBC_diff_meta,         iLength);
      VarArrayRedim(UV_vWBC_diff_band,         iLength);
      VarArrayRedim(UV_vWBC_diff_seg,          iLength);
      VarArrayRedim(UV_vWBC_diff_lym,          iLength);
      VarArrayRedim(UV_vWBC_diff_mono,         iLength);
      VarArrayRedim(UV_vWBC_diff_eos,          iLength);
      VarArrayRedim(UV_vWBC_diff_baso,         iLength);
      VarArrayRedim(UV_vWBC_diff_immaturecell, iLength);
      VarArrayRedim(UV_vWBC_diff_nRBC,         iLength);
      VarArrayRedim(UV_vPLT_number,            iLength);
      VarArrayRedim(UV_vPLT_size,              iLength);
      VarArrayRedim(UV_vOpinion,               iLength);
      VarArrayRedim(UV_vCod_doct,              iLength);
   end;
end;

function TfrmLD4I03.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if (rbDate.Checked and (mskDate.Text = '')) or
      (rbJumin.Checked and (mskJumin.Text = '')) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

function TfrmLD4I03.UF_Save : Boolean;
var i : Integer;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3  : String;
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
      with qry_PBS do
      begin
         Close;
         ParamByName('cod_bjjs').AsString     := UV_vCod_bjjs[i];
         ParamByName('num_id').AsString       := UV_vNum_id[i];
         ParamByName('dat_bunju').AsString    := UV_vDat_bunju[i];
         ParamByName('num_bunju').AsString    := UV_vNum_bunju[i];
         Open;
         if qry_PBS.IsEmpty = False then
         begin
            with qryU_PBS do
            begin
               ParamByName('cod_bjjs').AsString     := UV_vCod_bjjs[i];
               ParamByName('num_id').AsString       := UV_vNum_id[i];
               ParamByName('dat_bunju').AsString    := UV_vDat_bunju[i];
               ParamByName('num_bunju').AsString    := UV_vNum_bunju[i];

               if RBC_size_1.Checked = true then
                  ParamByName('RBC_size').AsString := 'marcrocytic'
               else if RBC_size_2.Checked = true then
                  ParamByName('RBC_size').AsString := 'normocytic'
               else if RBC_size_3.Checked = true then
                  ParamByName('RBC_size').AsString := 'microcytic'
               else
                  ParamByName('RBC_size').AsString := '';

               if RBC_stain_1.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'Hypochromic'
               else if RBC_stain_2.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'normochromic'
               else if RBC_stain_3.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'hyperchromic'
               else if RBC_stain_4.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'bichromatic'
               else
                  ParamByName('RBC_Stainability').AsString := '';

               if RBC_ani_1.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '0'
               else if RBC_ani_2.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '1'
               else if RBC_ani_3.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '2'
               else if RBC_ani_4.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '3'
               else
                  ParamByName('RBC_Anisocytosis').AsString := '';

               if RBC_poi_1.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '0'
               else if RBC_poi_2.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '1'
               else if RBC_poi_3.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '2'
               else if RBC_poi_4.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '3'
               else
                  ParamByName('RBC_Poikilocytosis').AsString := '';

               ParamByName('RBC_Inclusions').AsString := RBC_inc.Text;
               ParamByName('WBC_number').AsString     := WBC_num.Text;

               if WBC_gra_1.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '0'
               else if WBC_gra_2.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '1'
               else if WBC_gra_3.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '2'
               else if WBC_gra_4.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '3'
               else
                  ParamByName('WBC_toxic_granules').AsString := '';

               if WBC_vac_1.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '0'
               else if WBC_vac_2.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '1'
               else if WBC_vac_3.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '2'
               else if WBC_vac_4.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '3'
               else
                  ParamByName('WBC_toxic_vacuoles').AsString := '';

               if WBC_doh_1.Checked = true then
                  ParamByName('WBC_toxic_Dohlebody').AsString := '0'
               else if WBC_doh_2.Checked = true then
                  ParamByName('WBC_toxic_Dohlebody').AsString := '1'
               else
                  ParamByName('WBC_toxic_Dohlebody').AsString := '';

               if WBC_seg_1.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'hyposegmented'
               else if WBC_seg_2.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'normosegmented'
               else if WBC_seg_3.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'hypersemented'
               else
                  ParamByName('WBC_Segmentation').AsString := '';

               ParamByName('WBC_diff_blast').AsString        := WBC_bla.Text;
               ParamByName('WBC_diff_promyelo').AsString     := WBC_pro.Text;
               ParamByName('WBC_diff_myelo').AsString        := WBC_mye.Text;
               ParamByName('WBC_diff_meta').AsString         := WBC_met.Text;
               ParamByName('WBC_diff_band').AsString         := WBC_ban.Text;
               ParamByName('WBC_diff_seg').AsString          := WBC_seg.Text;
               ParamByName('WBC_diff_lym').AsString          := WBC_lym.Text;
               ParamByName('WBC_diff_mono').AsString         := WBC_mon.Text;
               ParamByName('WBC_diff_eos').AsString          := WBC_eos.Text;
               ParamByName('WBC_diff_baso').AsString         := WBC_bas.Text;
               ParamByName('WBC_diff_immaturecell').AsString := WBC_imm.Text;
               ParamByName('WBC_diff_nRBC').AsString         := WBC_rbc.Text;
               ParamByName('PLT_number').AsString            := PLT_num.Text;
               ParamByName('PLT_size').AsString              := PLT_size.Text;
               ParamByName('Opinion').AsMemo                 := Opinion.Text;
               ParamByName('COD_DOCT').AsString              := Copy(cmbCOD_DOCT.Text, 1, 6);
               ParamByName('COD_UPDATE').AsString            := GV_sUserId;
               Execsql;
               GP_query_log(qryU_PBS, 'LD4I03', 'Q', 'Y');
            end;
         end
         else
         begin
            with qryI_PBS do
            begin
               ParamByName('cod_bjjs').AsString     := UV_vCod_bjjs[i];
               ParamByName('num_id').AsString       := UV_vNum_id[i];
               ParamByName('dat_bunju').AsString    := UV_vDat_bunju[i];
               ParamByName('num_bunju').AsString    := UV_vNum_bunju[i];
               ParamByName('cod_jisa').AsString     := UV_vcod_jisa[i];
               ParamByName('dat_gmgn').AsString     := UV_vdat_gmgn[i];

               if RBC_size_1.Checked = true then
                  ParamByName('RBC_size').AsString := 'marcrocytic'
               else if RBC_size_2.Checked = true then
                  ParamByName('RBC_size').AsString := 'normocytic'
               else if RBC_size_3.Checked = true then
                  ParamByName('RBC_size').AsString := 'microcytic'
               else
                  ParamByName('RBC_size').AsString := '';

               if RBC_stain_1.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'Hypochromic'
               else if RBC_stain_2.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'normochromic'
               else if RBC_stain_3.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'hyperchromic'
               else if RBC_stain_4.Checked = true then
                  ParamByName('RBC_Stainability').AsString := 'bichromatic'
               else
                  ParamByName('RBC_Stainability').AsString := '';

               if RBC_ani_1.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '0'
               else if RBC_ani_2.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '1'
               else if RBC_ani_3.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '2'
               else if RBC_ani_4.Checked = true then
                  ParamByName('RBC_Anisocytosis').AsString := '3'
               else
                  ParamByName('RBC_Anisocytosis').AsString := '';

               if RBC_poi_1.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '0'
               else if RBC_poi_2.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '1'
               else if RBC_poi_3.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '2'
               else if RBC_poi_4.Checked = true then
                  ParamByName('RBC_Poikilocytosis').AsString := '3'
               else
                  ParamByName('RBC_Poikilocytosis').AsString := '';

               ParamByName('RBC_Inclusions').AsString := RBC_inc.Text;
               ParamByName('WBC_number').AsString     := WBC_num.Text;

               if WBC_gra_1.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '0'
               else if WBC_gra_2.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '1'
               else if WBC_gra_3.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '2'
               else if WBC_gra_4.Checked = true then
                  ParamByName('WBC_toxic_granules').AsString := '3'
               else
                  ParamByName('WBC_toxic_granules').AsString := '';

               if WBC_vac_1.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '0'
               else if WBC_vac_2.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '1'
               else if WBC_vac_3.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '2'
               else if WBC_vac_4.Checked = true then
                  ParamByName('WBC_toxic_vacuoles').AsString := '3'
               else
                  ParamByName('WBC_toxic_vacuoles').AsString := '';

               if WBC_doh_1.Checked = true then
                  ParamByName('WBC_toxic_Dohlebody').AsString := '0'
               else if WBC_doh_2.Checked = true then
                  ParamByName('WBC_toxic_Dohlebody').AsString := '1'
               else
                  ParamByName('WBC_toxic_Dohlebody').AsString := '';

               if WBC_seg_1.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'hyposegmented'
               else if WBC_seg_2.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'normosegmented'
               else if WBC_seg_3.Checked = true then
                  ParamByName('WBC_Segmentation').AsString := 'hypersemented'
               else
                  ParamByName('WBC_Segmentation').AsString := '';

               ParamByName('WBC_diff_blast').AsString        := WBC_bla.Text;
               ParamByName('WBC_diff_promyelo').AsString     := WBC_pro.Text;
               ParamByName('WBC_diff_myelo').AsString        := WBC_mye.Text;
               ParamByName('WBC_diff_meta').AsString         := WBC_met.Text;
               ParamByName('WBC_diff_band').AsString         := WBC_ban.Text;
               ParamByName('WBC_diff_seg').AsString          := WBC_seg.Text;
               ParamByName('WBC_diff_lym').AsString          := WBC_lym.Text;
               ParamByName('WBC_diff_mono').AsString         := WBC_mon.Text;
               ParamByName('WBC_diff_eos').AsString          := WBC_eos.Text;
               ParamByName('WBC_diff_baso').AsString         := WBC_bas.Text;
               ParamByName('WBC_diff_immaturecell').AsString := WBC_imm.Text;
               ParamByName('WBC_diff_nRBC').AsString         := WBC_rbc.Text;
               ParamByName('PLT_number').AsString            := PLT_num.Text;
               ParamByName('PLT_size').AsString              := PLT_size.Text;
               ParamByName('Opinion').AsMemo                 := Opinion.Text;
               ParamByName('COD_DOCT').AsString              := Copy(cmbCOD_DOCT.Text, 1, 6);
               ParamByName('COD_INSERT').AsString            := GV_sUserId;
               Execsql;
               GP_query_log(qryI_PBS, 'LD4I03', 'Q', 'Y');
            end;
         end;

         if UV_vstatus = false then
         begin
             // 혈액결과...업데이트
             qryHul.Active := False;
             qryHul.ParamByName('cod_bjjs').AsString     := UV_vCod_bjjs[i];
             qryHul.ParamByName('num_id').AsString       := UV_vNum_id[i];
             qryHul.ParamByName('dat_bunju').AsString    := UV_vDat_bunju[i];
             qryHul.ParamByName('num_bunju').AsString    := UV_vNum_bunju[i];
             qryHul.ParamByName('cod_jisa').AsString     := UV_vcod_jisa[i];
             qryHul.ParamByName('dat_gmgn').AsString     := UV_vdat_gmgn[i];
             qryHul.Active := True;
             if qryHul.RecordCount > 0 then
             begin
                 UV_fGulkwa := '';
                 UV_fGulkwa1 := qryHul.FieldByName('DESC_GLKWA1').AsString;
                 UV_fGulkwa2 := qryHul.FieldByName('DESC_GLKWA2').AsString;
                 UV_fGulkwa3 := qryHul.FieldByName('DESC_GLKWA3').AsString;
                 GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                 UV_fGulkwa := copy(UV_fGulkwa, 1, 67-1) +
                               GF_RPad(WBC_seg.Text, 6, ' ') +        //H012
                               GF_RPad(WBC_ban.Text, 6, ' ') +        //H013
                               GF_RPad(WBC_lym.Text, 6, ' ') +        //H014
                               GF_RPad(WBC_mon.Text, 6, ' ') +        //H015
                               GF_RPad(WBC_eos.Text, 6, ' ') +        //H016
                               GF_RPad(WBC_bas.Text, 6, ' ') +        //H017
                               GF_RPad(WBC_met.Text, 6, ' ') +        //H018
                               GF_RPad(WBC_mye.Text, 6, ' ') +        //H019
                               GF_RPad(WBC_pro.Text, 6, ' ') +        //H020
                               GF_RPad(WBC_bla.Text, 6, ' ') +        //H021
                               GF_RPad(WBC_rbc.Text, 6, ' ') +        //H022
                               copy(UV_fGulkwa, 133, length(UV_fGulkwa));
                 GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                 with qryU_glkwa do
                 begin
                    ParamByName('cod_bjjs').AsString     := UV_vCod_bjjs[i];
                    ParamByName('num_id').AsString       := UV_vNum_id[i];
                    ParamByName('dat_bunju').AsString    := UV_vDat_bunju[i];
                    ParamByName('num_bunju').AsString    := UV_vNum_bunju[i];
                    ParamByName('cod_jisa').AsString     := UV_vcod_jisa[i];
                    ParamByName('dat_gmgn').AsString     := UV_vdat_gmgn[i];
                    ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                    ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                    ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                    ParamByName('COD_update').AsString   := Copy(cmbCOD_DOCT.Text, 1, 6);
                    ParamByName('dat_update').AsString   := GV_sToday;
                    Execsql;
                    GP_query_log(qryU_glkwa, 'LD4I03', 'Q', 'Y');
                 end;
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

   //Grid에 자료를 반영
   with grdMaster do
   begin
      if RBC_size_1.Checked = true then
         UV_vRBC_size[i] := 'marcrocytic'
      else if RBC_size_2.Checked = true then
         UV_vRBC_size[i] := 'normocytic'
      else if RBC_size_3.Checked = true then
         UV_vRBC_size[i] := 'microcytic'
      else
         UV_vRBC_size[i] := '';

      if RBC_stain_1.Checked = true then
         UV_vRBC_Stainability[i] := 'Hypochromic'
      else if RBC_stain_2.Checked = true then
         UV_vRBC_Stainability[i] := 'normochromic'
      else if RBC_stain_3.Checked = true then
         UV_vRBC_Stainability[i] := 'hyperchromic'
      else if RBC_stain_4.Checked = true then
         UV_vRBC_Stainability[i] := 'bichromatic'
      else
         UV_vRBC_Stainability[i] := '';

      if RBC_ani_1.Checked = true then
         UV_vRBC_Anisocytosis[i] := '0'
      else if RBC_ani_2.Checked = true then
         UV_vRBC_Anisocytosis[i] := '1'
      else if RBC_ani_3.Checked = true then
         UV_vRBC_Anisocytosis[i] := '2'
      else if RBC_ani_4.Checked = true then
         UV_vRBC_Anisocytosis[i] := '3'
      else
         UV_vRBC_Anisocytosis[i] := '';

      if RBC_poi_1.Checked = true then
         UV_vRBC_Poikilocytosis[i] := '0'
      else if RBC_poi_2.Checked = true then
         UV_vRBC_Poikilocytosis[i] := '1'
      else if RBC_poi_3.Checked = true then
         UV_vRBC_Poikilocytosis[i] := '2'
      else if RBC_poi_4.Checked = true then
         UV_vRBC_Poikilocytosis[i] := '3'
      else
         UV_vRBC_Poikilocytosis[i] := '';

      UV_vRBC_Inclusions[i] := RBC_inc.Text;
      UV_vWBC_number[i] := WBC_num.Text;

      if WBC_gra_1.Checked = true then
         UV_vWBC_toxic_granules[i] := '0'
      else if WBC_gra_2.Checked = true then
         UV_vWBC_toxic_granules[i] := '1'
      else if WBC_gra_3.Checked = true then
         UV_vWBC_toxic_granules[i] := '2'
      else if WBC_gra_4.Checked = true then
         UV_vWBC_toxic_granules[i] := '3'
      else
         UV_vWBC_toxic_granules[i] := '';

      if WBC_vac_1.Checked = true then
         UV_vWBC_toxic_vacuoles[i] := '0'
      else if WBC_vac_2.Checked = true then
         UV_vWBC_toxic_vacuoles[i] := '1'
      else if WBC_vac_3.Checked = true then
         UV_vWBC_toxic_vacuoles[i] := '2'
      else if WBC_vac_4.Checked = true then
         UV_vWBC_toxic_vacuoles[i] := '3'
      else
         UV_vWBC_toxic_vacuoles[i] := '';

      if WBC_doh_1.Checked = true then
         UV_vWBC_toxic_Dohlebody[i] := '0'
      else if WBC_doh_2.Checked = true then
         UV_vWBC_toxic_Dohlebody[i] := '1'
      else
         UV_vWBC_toxic_Dohlebody[i] := '';

      if WBC_seg_1.Checked = true then
         UV_vWBC_Segmentation[i] := 'hyposegmented'
      else if WBC_seg_2.Checked = true then
         UV_vWBC_Segmentation[i] := 'normosegmented'
      else if WBC_seg_3.Checked = true then
         UV_vWBC_Segmentation[i] := 'hypersemented'
      else
         UV_vWBC_Segmentation[i] := '';

      UV_vWBC_diff_blast[i] := WBC_bla.Text;
      UV_vWBC_diff_promyelo[i] := WBC_pro.Text;
      UV_vWBC_diff_myelo[i] := WBC_mye.Text;
      UV_vWBC_diff_meta[i] := WBC_met.Text;
      UV_vWBC_diff_band[i] := WBC_ban.Text;
      UV_vWBC_diff_seg[i] := WBC_seg.Text;
      UV_vWBC_diff_lym[i] := WBC_lym.Text;
      UV_vWBC_diff_mono[i] := WBC_mon.Text;
      UV_vWBC_diff_eos[i] := WBC_eos.Text;
      UV_vWBC_diff_baso[i] := WBC_bas.Text;
      UV_vWBC_diff_immaturecell[i] := WBC_imm.Text;
      UV_vWBC_diff_nRBC[i] := WBC_rbc.Text;
      UV_vPLT_number[i] := PLT_num.Text;
      UV_vPLT_size[i] := PLT_size.Text;
      UV_vOpinion[i] := Opinion.Text;
      UV_vCod_doct[i]    := Copy(cmbCOD_DOCT.Text, 1, 6);
   // Grid Repaint하여 Cellloaded를 강제적으로 발생
      grdMaster.Repaint;
   end;

   UV_iCod_doct := cmbCOD_DOCT.ItemIndex;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;


procedure TfrmLD4I03.btnSaveClick(Sender: TObject);
begin
   inherited;
   UF_Save;
end;

procedure TfrmLD4I03.FormCreate(Sender: TObject);
begin
   inherited;

   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   pnlJisa.Visible := True;

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   UV_iCod_doct := -1;
   mskDate.Text := GV_sToday;
   GP_ComboSawonAll(cmbCOD_DOCT, 1);
   if GV_sUserId = '150005' then
       GP_ComboDisplay(cmbCOD_Doct, '150005', 6)
   else if GV_sUserId = '150783' then
       GP_ComboDisplay(cmbCOD_Doct, '150783', 6);


end;

procedure TfrmLD4I03.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCod_Jisa[DataRow-1];
      2 : Value := GF_JuminFormat(UV_vNum_jumin[DataRow-1]);
      3 : Value := UV_vDesc_name[DataRow-1];
   end;

end;

procedure TfrmLD4I03.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sSelect, sWhere, sOrderby : String;
begin
   inherited;
   UV_bEdit := True;
   btnPsave.Enabled   := True;
   btnPcancel.Enabled := True;
   Opinion.readonly := false;

   {with qryUser_priv do
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
   end;   }

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid Setting
   UP_GridSet;

   // Query 수행 & 배열에 저장
   index := 0;
   with qryPBS do
   begin
      Close;

      sSelect := '           SELECT g.num_jumin, g.desc_name, g.cod_jangbi, g.cod_hul, g.cod_cancer, g.cod_chuga, g.gubn_nosin, g.gubn_nsyh, g.dat_gmgn, g.cod_jisa, g.num_id, g.num_samp,                ';
      sSelect := sSelect + '          g.gubn_adult, g.gubn_adyh, g.gubn_bogun, g.gubn_bgyh, g.gubn_agab, g.gubn_agyh, g.gubn_life, g.gubn_lfyh, g.gubn_tkgm,  ';
      sSelect := sSelect + '          b.cod_bjjs, b.num_bunju, b.dat_bunju,                                                                                   ';
      sSelect := sSelect + '          p.RBC_size, p.RBC_stainability, p.RBC_anisocytosis, p.RBC_poikilocytosis, p.RBC_inclusions,                             ';
      sSelect := sSelect + '          p.WBC_number, p.WBC_toxic_granules, p.WBC_toxic_vacuoles, p.WBC_toxic_dohlebody,        ';
      sSelect := sSelect + '          p.WBC_segmentation, p.WBC_diff_blast, p.WBC_diff_promyelo, p.WBC_diff_myelo,                                            ';
      sSelect := sSelect + '          p.WBC_diff_meta, p.WBC_diff_band, p.WBC_diff_seg, p.WBC_diff_lym, p.WBC_diff_mono,                                      ';
      sSelect := sSelect + '          p.WBC_diff_eos, p.WBC_diff_baso, p.WBC_diff_immaturecell, p.WBC_diff_nRBC,                           ';
      sSelect := sSelect + '          p.PLT_number, p.PLT_size, p.Opinion,p.cod_doct                                                                        ';
      sSelect := sSelect + ' from gumgin g inner join bunju B                                                                                                 ';
      sSelect := sSelect + ' 		on g.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn                                        ';
      sSelect := sSelect + ' 		   left outer join PBSmear P                                                                                          ';
      sSelect := sSelect + ' 	        on b.cod_bjjs = p.cod_bjjs and b.num_id = p.num_id and b.dat_bunju = p.dat_bunju and b.num_bunju = p.num_bunju        ';

      sWhere := 'WHERE ';
      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := sWhere + ' G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND ';

      if rbDate.Checked then
      begin
            sWhere := sWhere + ' G.dat_gmgn = ''' + mskDate.Text + ''' ';
      end
      else if rbJumin.Checked then
      begin
         sWhere := sWhere + ' G.num_jumin = ''' + mskJumin.Text + ''' ';
      end;

      sOrderBy := ' ORDER BY g.num_samp';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      if qryPBS.IsEmpty = False then
      begin
         GP_query_log(qryPBS, 'LD4I03', 'Q', 'N');
         while Not Eof do
         begin
            if UF_hangmokList = true then
            begin
                 UP_VarMemSet(index);
                 UV_vCod_bjjs[index]              := FieldByName('Cod_bjjs').AsString;
                 UV_vNum_id[index]                := FieldByName('Num_id').AsString;
                 UV_vNum_jumin[index]             := FieldByName('Num_jumin').AsString;
                 UV_vDesc_name[index]             := FieldByName('Desc_name').AsString;
                 UV_vDat_gmgn[index]              := FieldByName('Dat_gmgn').AsString;
                 UV_vnum_bunju[index]             := FieldByName('num_bunju').AsString;
                 UV_vDat_bunju[index]             := FieldByName('Dat_bunju').AsString;
                 UV_vCod_jisa[index]              := FieldByName('Cod_jisa').AsString;
                 UV_vRBC_size[index]              := FieldByName('RBC_size').AsString;
                 UV_vRBC_stainability[index]      := FieldByName('RBC_stainability').AsString;
                 UV_vRBC_anisocytosis[index]      := FieldByName('RBC_anisocytosis').AsString;
                 UV_vRBC_poikilocytosis[index]    := FieldByName('RBC_poikilocytosis').AsString;
                 UV_vRBC_inclusions[index]        := FieldByName('RBC_inclusions').AsString;
                 UV_vWBC_number[index]            := FieldByName('WBC_number').AsString;
                 UV_vWBC_toxic_granules[index]    := FieldByName('WBC_toxic_granules').AsString;
                 UV_vWBC_toxic_vacuoles[index]    := FieldByName('WBC_toxic_vacuoles').AsString;
                 UV_vWBC_toxic_dohlebody[index]   := FieldByName('WBC_toxic_dohlebody').AsString;
                 UV_vWBC_segmentation[index]      := FieldByName('WBC_segmentation').AsString;
                 UV_vWBC_diff_blast[index]        := FieldByName('WBC_diff_blast').AsString;
                 UV_vWBC_diff_promyelo[index]     := FieldByName('WBC_diff_promyelo').AsString;
                 UV_vWBC_diff_myelo[index]        := FieldByName('WBC_diff_myelo').AsString;
                 UV_vWBC_diff_meta[index]         := FieldByName('WBC_diff_meta').AsString;
                 UV_vWBC_diff_band[index]         := FieldByName('WBC_diff_band').AsString;
                 UV_vWBC_diff_seg[index]          := FieldByName('WBC_diff_seg').AsString;
                 UV_vWBC_diff_lym[index]          := FieldByName('WBC_diff_lym').AsString;
                 UV_vWBC_diff_mono[index]         := FieldByName('WBC_diff_mono').AsString;
                 UV_vWBC_diff_eos[index]          := FieldByName('WBC_diff_eos').AsString;
                 UV_vWBC_diff_baso[index]         := FieldByName('WBC_diff_baso').AsString;
                 UV_vWBC_diff_immaturecell[index] := FieldByName('WBC_diff_immaturecell').AsString;
                 UV_vWBC_diff_nRBC[index]         := FieldByName('WBC_diff_nRBC').AsString;
                 UV_vPLT_number[index]            := FieldByName('PLT_number').AsString;
                 UV_vPLT_size[index]              := FieldByName('PLT_size').AsString;
                 UV_vOpinion[index]               := FieldByName('Opinion').AsString;
                 UV_vCod_doct[index]              := FieldByName('Cod_doct').AsString;
                 Inc(index);
            end;
            Next;
         end;

         // Table과 Disconnect
         Active := False;
      end
      else
      begin
         GP_FieldClear(pnlDetail);
         GP_FieldClear(Panel2);
         GP_FieldClear(Panel_RBC);
         GP_FieldClear(Panel_WBC);
         GP_FieldClear(Panel_PLT);
         GF_MsgBox('Information', 'NODATA');
      end;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4I03.btnCancelClick(Sender: TObject);
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
      GP_FieldClear(Panel_RBC);
      GP_FieldClear(Panel_WBC);
      GP_FieldClear(Panel_PLT);
   end;

   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD4I03.UP_Change(Sender: TObject);
begin
   inherited;
   // Edit Flag Check
   UV_bEdit := True;

end;

procedure TfrmLD4I03.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var sMan : String;
    iAge : integer;
begin
   inherited;
   UV_vstatus := false;
   Opinion.readonly := false;
   GP_FieldClear(pnlDetail);
   GP_FieldClear(Panel_RBC);
   GP_FieldClear(Panel_WBC);
   GP_FieldClear(Panel_PLT);

   RBC_size_1.Checked  := false;
   RBC_size_2.Checked  := false;
   RBC_size_3.Checked  := false;

   RBC_stain_1.Checked := false;
   RBC_stain_2.Checked := false;
   RBC_stain_3.Checked := false;
   RBC_stain_4.Checked := false;

   RBC_ani_1.Checked   := false;
   RBC_ani_2.Checked   := false;
   RBC_ani_3.Checked   := false;
   RBC_ani_4.Checked   := false;

   RBC_poi_1.Checked   := false;
   RBC_poi_2.Checked   := false;
   RBC_poi_3.Checked   := false;
   RBC_poi_4.Checked   := false;

   WBC_gra_1.Checked   := false;
   WBC_gra_2.Checked   := false;
   WBC_gra_3.Checked   := false;
   WBC_gra_4.Checked   := false;

   WBC_vac_1.Checked   := false;
   WBC_vac_2.Checked   := false;
   WBC_vac_3.Checked   := false;
   WBC_vac_4.Checked   := false;

   WBC_doh_1.Checked   := false;
   WBC_doh_2.Checked   := false;

   WBC_seg_1.Checked   := false;
   WBC_seg_2.Checked   := false;
   WBC_seg_3.Checked   := false;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   desc_name.Text    := UV_vdesc_name[NewRow-1];
   GP_JuminToAge(UV_vnum_jumin[NewRow-1],UV_vdat_gmgn[NewRow-1], sMan, iAge);
   age_sex.Text      := intToStr(iAge) + ' / ' + sMan;
   num_bunju.Text    := UV_vnum_bunju[NewRow-1];
   dat_gmgn.Text     := UV_vdat_gmgn[NewRow-1];
   dat_bunju.Text    := UV_vdat_bunju[NewRow-1];

   if UV_vRBC_size[NewRow-1] = 'marcrocytic' then
      RBC_size_1.Checked  := true
   else if UV_vRBC_size[NewRow-1] = 'normocytic' then
      RBC_size_2.Checked  := true
   else if UV_vRBC_size[NewRow-1] = 'microcytic' then
      RBC_size_3.Checked  := true;

   if UV_vRBC_Stainability[NewRow-1] = 'Hypochromic' then
      RBC_stain_1.Checked  := true
   else if UV_vRBC_Stainability[NewRow-1] = 'normochromic' then
      RBC_stain_2.Checked  := true
   else if UV_vRBC_Stainability[NewRow-1] = 'hyperchromic' then
      RBC_stain_3.Checked  := true
   else if UV_vRBC_Stainability[NewRow-1] = 'bichromatic' then
      RBC_stain_4.Checked  := true;

   if UV_vRBC_Anisocytosis[NewRow-1] = '0' then
      RBC_ani_1.Checked  := true
   else if UV_vRBC_Anisocytosis[NewRow-1] = '1' then
      RBC_ani_2.Checked  := true
   else if UV_vRBC_Anisocytosis[NewRow-1] = '2' then
      RBC_ani_3.Checked  := true
   else if UV_vRBC_Anisocytosis[NewRow-1] = '3' then
      RBC_ani_4.Checked  := true;

   if UV_vRBC_Poikilocytosis[NewRow-1] = '0' then
      RBC_poi_1.Checked  := true
   else if UV_vRBC_Poikilocytosis[NewRow-1] = '1' then
      RBC_poi_2.Checked  := true
   else if UV_vRBC_Poikilocytosis[NewRow-1] = '2' then
      RBC_poi_3.Checked  := true
   else if UV_vRBC_Poikilocytosis[NewRow-1] = '3' then
      RBC_poi_4.Checked  := true;

   RBC_inc.Text := UV_vRBC_Inclusions[NewRow-1];

   WBC_num.Text := UV_vWBC_number[NewRow-1];

   if UV_vWBC_toxic_granules[NewRow-1] = '0' then
      WBC_gra_1.Checked  := true
   else if UV_vWBC_toxic_granules[NewRow-1] = '1' then
      WBC_gra_2.Checked  := true
   else if UV_vWBC_toxic_granules[NewRow-1] = '2' then
      WBC_gra_3.Checked  := true
   else if UV_vWBC_toxic_granules[NewRow-1] = '3' then
      WBC_gra_4.Checked  := true;

   if UV_vWBC_toxic_vacuoles[NewRow-1] = '0' then
      WBC_vac_1.Checked  := true
   else if UV_vWBC_toxic_vacuoles[NewRow-1] = '1' then
      WBC_vac_2.Checked  := true
   else if UV_vWBC_toxic_vacuoles[NewRow-1] = '2' then
      WBC_vac_3.Checked  := true
   else if UV_vWBC_toxic_vacuoles[NewRow-1] = '3' then
      WBC_vac_4.Checked  := true;

   if UV_vWBC_toxic_Dohlebody[NewRow-1] = '0' then
      WBC_doh_1.Checked  := true
   else if UV_vWBC_toxic_Dohlebody[NewRow-1] = '1' then
      WBC_doh_2.Checked  := true;

   if UV_vWBC_Segmentation[NewRow-1] = 'hyposegmented' then
      WBC_seg_1.Checked  := true
   else if UV_vWBC_Segmentation[NewRow-1] = 'normosegmented' then
      WBC_seg_2.Checked  := true
   else if UV_vWBC_Segmentation[NewRow-1] = 'hypersemented' then
      WBC_seg_3.Checked  := true;

   WBC_bla.Text := UV_vWBC_diff_blast[NewRow-1];
   WBC_pro.Text := UV_vWBC_diff_promyelo[NewRow-1];
   WBC_mye.Text := UV_vWBC_diff_myelo[NewRow-1];
   WBC_met.Text := UV_vWBC_diff_meta[NewRow-1];
   WBC_ban.Text := UV_vWBC_diff_band[NewRow-1];
   WBC_seg.Text := UV_vWBC_diff_seg[NewRow-1];
   WBC_lym.Text := UV_vWBC_diff_lym[NewRow-1];
   WBC_mon.Text := UV_vWBC_diff_mono[NewRow-1];
   WBC_eos.Text := UV_vWBC_diff_eos[NewRow-1];
   WBC_bas.Text := UV_vWBC_diff_baso[NewRow-1];
   WBC_imm.Text := UV_vWBC_diff_immaturecell[NewRow-1];
   WBC_rbc.Text := UV_vWBC_diff_nRBC[NewRow-1];

   PLT_num.Text    := UV_vPLT_number[NewRow-1];
   PLT_size.Text   := UV_vPLT_size[NewRow-1];

   Opinion.Text := UV_vOpinion[NewRow-1];
   if Trim(UV_vCod_Doct[NewRow-1]) = '' then
   begin
     if GV_sUserId = '150005' then
          GP_ComboDisplay(cmbCOD_Doct, '150005', 6)
     else if GV_sUserId = '150783' then
          GP_ComboDisplay(cmbCOD_Doct, '150783', 6)
     else cmbCOD_Doct.ItemIndex := -1;
   end
   else
       GP_ComboDisplay(cmbCOD_Doct, UV_vCod_Doct[NewRow-1], 6);

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD4I03.UP_Click(Sender: TObject);
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

procedure TfrmLD4I03.rbDateClick(Sender: TObject);
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

procedure TfrmLD4I03.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   // Reference Key가 아니면 종료
   if Key <> GC_Refer then exit;

   if      Sender = mskDate    then UP_Click(btnDate)
   else if Sender = mskJumin   then UP_Click(btnJumin);
end;

procedure TfrmLD4I03.UP_Exit(Sender: TObject);
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
function TfrmLD4I03.UF_hangmokList : Boolean;
var sHmCode :string;
begin
   result := false; sHmCode := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryPBS.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPBS.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sHmCode := sHmCode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if qryPBS.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPBS.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sHmCode := sHmCode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if qryPBS.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPBS.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sHmCode := sHmCode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sHmCode := sHmCode + qryPBS.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   if qryPBS.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPBS.FieldByName('Dat_gmgn').AsString, '1', qryPBS.FieldByName('Gubn_nsyh').AsInteger)
   else if qryPBS.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPBS.FieldByName('Dat_gmgn').AsString, '4', qryPBS.FieldByName('Gubn_adyh').AsInteger);

   if qryPBS.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPBS.FieldByName('Dat_gmgn').AsString, '7', qryPBS.FieldByName('Gubn_lfyh').AsInteger);

   if qryPBS.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPBS.FieldByName('Dat_gmgn').AsString, '3',qryPBS.FieldByName('Gubn_bgyh').AsInteger);

   if qryPBS.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPBS.FieldByName('Dat_gmgn').AsString, '5',qryPBS.FieldByName('Gubn_agyh').AsInteger);

   if (qryPBS.FieldByName('Gubn_tkgm').AsString = '1') or (qryPBS.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryPBS.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryPBS.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryPBS.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         qryProfileG.Active := False;
         qryProfileG.ParamByName('cod_pf1').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  1, 4);
         qryProfileG.ParamByName('cod_pf2').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  5, 4);
         qryProfileG.ParamByName('cod_pf3').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString,  9, 4);
         qryProfileG.ParamByName('cod_pf4').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 13, 4);
         qryProfileG.ParamByName('cod_pf5').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 17, 4);
         qryProfileG.ParamByName('cod_pf6').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 21, 4);
         qryProfileG.ParamByName('cod_pf7').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 25, 4);
         qryProfileG.ParamByName('cod_pf8').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 29, 4);
         qryProfileG.ParamByName('cod_pf9').AsString  := Copy(qryTkgum.FieldByName('cod_prf').AsString, 33, 4);
         qryProfileG.ParamByName('cod_pf10').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 37, 4);
         qryProfileG.ParamByName('cod_pf11').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 41, 4);
         qryProfileG.ParamByName('cod_pf12').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 45, 4);
         qryProfileG.ParamByName('cod_pf13').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 49, 4);
         qryProfileG.ParamByName('cod_pf14').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 53, 4);
         qryProfileG.ParamByName('cod_pf15').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 57, 4);
         qryProfileG.ParamByName('cod_pf16').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 61, 4);
         qryProfileG.ParamByName('cod_pf17').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 65, 4);
         qryProfileG.ParamByName('cod_pf18').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 69, 4);
         qryProfileG.ParamByName('cod_pf19').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 73, 4);
         qryProfileG.ParamByName('cod_pf20').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 77, 4);
         qryProfileG.ParamByName('cod_pf21').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 81, 4);
         qryProfileG.ParamByName('cod_pf22').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 85, 4);
         qryProfileG.ParamByName('cod_pf23').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 89, 4);
         qryProfileG.ParamByName('cod_pf24').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 93, 4);
         qryProfileG.ParamByName('cod_pf25').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 97, 4);
         qryProfileG.Active := True;
         if qryProfileG.RecordCount > 0 then
         begin
            while not qryProfileG.Eof do
            begin
               sHmCode := sHmCode + qryProfileG.FieldByName('cod_hm').AsString;
               qryProfileG.Next;
            end;
         end;
         qryProfileG.Active := False;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if pos('H036',sHmCode) > 0 then
      result := True;
end;

function TfrmLD4I03.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4I03.btn_signClick(Sender: TObject);
Var
   index, Max_Num : integer;
begin
  inherited;
  index := grdMaster.CurrentDataRow - 1;
      If b_sgCAppSet = False Then              //인증로그인이 안되어 있으면
      begin
         KMIConnect();
         GF_PubCertifyLoad();
      end;

{      if not UV_bEdit then
      begin
         showmessage('변경된 내용이 없습니다.');
         exit;
      end; }

      If (b_sgCAppSet = True) and (Index >= 0) Then
      Begin
         //저장  프로시져 호출
         UF_Save;
         with Qry_CK_PBS do
         begin
            Close;
            ParamByName('Cod_bjjs').AsString := UV_vCod_bjjs[Index];
            ParamByName('Num_id').AsString   := UV_vNum_id[index];
            ParamByName('Dat_bunju').AsString := UV_vDat_bunju[index];
            ParamByName('num_bunju').AsString := UV_vnum_bunju[index];

            Open;
            if IsEmpty = False then
            begin
               with Qry_Sign do
               begin
                  Close;
                  ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString     := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[index];
                  ParamByName('Program_id').AsString := 'LD4I03';
                  open;
                  Max_Num := Qry_Sign.FieldByName('seq_no').AsInteger;
               end;

               GV_sGulkwa := 'cod_bjjs[' + FieldByName('cod_bjjs').AsString + ']' +
                             'num_id[' + FieldByName('num_id').AsString + ']' +
                             'dat_bunju[' + FieldByName('dat_bunju').AsString + ']' +
                             'num_bunju[' + FieldByName('num_bunju').AsString + ']' +
                             'cod_jisa[' + FieldByName('cod_jisa').AsString + ']' +
                             'dat_gmgn[' + FieldByName('dat_gmgn').AsString + ']' +
                             'RBC_size[' + FieldByName('RBC_size').AsString + ']' +
                             'RBC_stainability[' + FieldByName('RBC_stainability').AsString + ']' +
                             'RBC_anisocytosis[' + FieldByName('RBC_anisocytosis').AsString + ']' +
                             'RBC_poikilocytosis[' + FieldByName('RBC_poikilocytosis').AsString + ']' +
                             'RBC_inclusions[' + FieldByName('RBC_inclusions').AsString + ']' +
                             'WBC_number[' + FieldByName('WBC_number').AsString + ']' +
                             'WBC_toxic_granules[' + FieldByName('WBC_toxic_granules').AsString + ']' +
                             'WBC_toxic_vacuoles[' + FieldByName('WBC_toxic_vacuoles').AsString + ']' +
                             'WBC_toxic_dohlebody[' + FieldByName('WBC_toxic_dohlebody').AsString + ']' +
                             'WBC_segmentation[' + FieldByName('WBC_segmentation').AsString + ']' +
                             'WBC_diff_blast[' + FieldByName('WBC_diff_blast').AsString + ']' +
                             'WBC_diff_promyelo[' + FieldByName('WBC_diff_promyelo').AsString + ']' +
                             'WBC_diff_myelo[' + FieldByName('WBC_diff_myelo').AsString + ']' +
                             'WBC_diff_meta[' + FieldByName('WBC_diff_meta').AsString + ']' +
                             'WBC_diff_band[' + FieldByName('WBC_diff_band').AsString + ']' +
                             'WBC_diff_seg[' + FieldByName('WBC_diff_seg').AsString + ']' +
                             'WBC_diff_lym[' + FieldByName('WBC_diff_lym').AsString + ']' +
                             'WBC_diff_mono[' + FieldByName('WBC_diff_mono').AsString + ']' +
                             'WBC_diff_eos[' + FieldByName('WBC_diff_eos').AsString + ']' +
                             'WBC_diff_baso[' + FieldByName('WBC_diff_baso').AsString + ']' +
                             'WBC_diff_immaturecell[' + FieldByName('WBC_diff_immaturecell').AsString + ']' +
                             'WBC_diff_nRBC[' + FieldByName('WBC_diff_nRBC').AsString + ']' +
                             'PLT_number[' + FieldByName('PLT_number').AsString + ']' +
                             'PLT_size[' + FieldByName('PLT_size').AsString + ']' +
                             'Opinion[' + FieldByName('Opinion').AsString + ']' +
                             'dat_insert[' + FieldByName('dat_insert').AsString + ']' +
                             'cod_insert[' + FieldByName('cod_insert').AsString + ']' +
                             'dat_update[' + FieldByName('dat_update').AsString + ']' +
                             'cod_update[' + FieldByName('cod_update').AsString + ']' ;

               GF_PubSignCertify;

               with QryI_Sign do
               begin
                  ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[Index];
                  ParamByName('Num_id').AsString     := UV_vNum_id[index];
                  ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[index];
                  ParamByName('Program_id').AsString := 'LD4I03';
                  ParamByName('Seq_no').Asinteger    := Max_Num + 1;
                  ParamByName('Result').AsMemo       := GV_sGulkwa;
                  ParamByName('Sign_Key').AsMemo   := gv_strSignedMsg;
                  ParamByName('cod_sawon').AsString  := GV_sUserID;
                  Execsql;
                  GP_query_log(QryI_Sign, 'LD4I03', 'Q', 'Y');
               end;
            end
            else
            begin
               ShowMessage('서명저장할 내용이 없습니다.');
               Exit;
            end;
         end;
      end
      Else If (b_sgCAppSet = False) then //처음 로그인 할때 인증서 로그인을 안했을경우
      Begin
           ShowMessage('서명 저장을 다시 하시기 바랍니다.');
           Exit;
      end
      Else If Index < 0 then
      Begin
           ShowMessage('조회된 내용이 없습니다. 조회후 서명저장하시기 바랍니다.');
           Exit;
      end;

      //저장이 완료된후에는 글로벌 변수 초기화
      gv_strSignDate  := '';
      gv_strSignedMsg := '';
      GV_sGulkwa      := '';

end;

procedure TfrmLD4I03.BitBtn1Click(Sender: TObject);
begin
  inherited;
   UV_vstatus := true;

   GP_FieldClear(pnlDetail);
   GP_FieldClear(Panel_RBC);
   GP_FieldClear(Panel_WBC);
   GP_FieldClear(Panel_PLT);

   RBC_size_1.Checked  := false;
   RBC_size_2.Checked  := false;
   RBC_size_3.Checked  := false;

   RBC_stain_1.Checked := false;
   RBC_stain_2.Checked := false;
   RBC_stain_3.Checked := false;
   RBC_stain_4.Checked := false;

   RBC_ani_1.Checked   := false;
   RBC_ani_2.Checked   := false;
   RBC_ani_3.Checked   := false;
   RBC_ani_4.Checked   := false;

   RBC_poi_1.Checked   := false;
   RBC_poi_2.Checked   := false;
   RBC_poi_3.Checked   := false;
   RBC_poi_4.Checked   := false;

   WBC_gra_1.Checked   := false;
   WBC_gra_2.Checked   := false;
   WBC_gra_3.Checked   := false;
   WBC_gra_4.Checked   := false;

   WBC_vac_1.Checked   := false;
   WBC_vac_2.Checked   := false;
   WBC_vac_3.Checked   := false;
   WBC_vac_4.Checked   := false;

   WBC_doh_1.Checked   := false;
   WBC_doh_2.Checked   := false;

   WBC_seg_1.Checked   := false;
   WBC_seg_2.Checked   := false;
   WBC_seg_3.Checked   := false;
   
    Opinion.Text := '재검결과  WBC 항목이 참고치내에 있으므로 PB 슬라이드 제작 및 검경을 하지 않았습니다.';
    Opinion.readonly := true;
    cmbCOD_DOCT.Text := '';
end;

end.
