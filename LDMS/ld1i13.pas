//==============================================================================
// 프로그램 설명 : [센터]자체검진 혈액학결과등록
// 시스템        : LDMS
// 수정일자      : 2011.12.14
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2013.5.9
// 수정자        : 김희석
// 수정내용      : 키, 몸무게,임신및 생리
//=============================================================================
// 수정일자      : 2013.6.15
// 수정자        : 김승철
// 수정내용      : 임신중 안나오는 부분 수정
//=============================================================================
// 수정일자      : 2014.07.26
// 수정자        : 곽윤설
// 수정내용      : 혈액형 검사자 조회시, 혈액 프로파일 모두 확인
//=============================================================================
// 수정일자      : 2014.11.13
// 수정자        : 곽윤설
// 수정내용      : 결과값 변환하는 엑셀버튼 생성
// 참고사항      : [강남 - 전미라]
//=============================================================================
unit LD1I13;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges, ComObj;

type
  TfrmLD1I13 = class(TfrmSingle)
    qryGlkwa: TQuery;
    grdMaster: TtsGrid;
    qryU_Glkwa: TQuery;
    pnlDetail: TPanel;
    edtValue1: TEdit;
    btnPSave: TBitBtn;
    btnPCancel: TBitBtn;
    pnlName1: TPanel;
    pnlCond: TPanel;
    rbJumin: TRadioButton;
    btnDate: TSpeedButton;
    mskJumin: TMaskEdit;
    btnJumin: TSpeedButton;
    edtName: TEdit;
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
    pnlNum9: TPanel;
    pnlName9: TPanel;
    edtValue9: TEdit;
    Panel33: TPanel;
    Panel34: TPanel;
    mskDAT_BUNJU: TDateEdit;
    edtNUM_BUNJU: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    mskNUM_JUMIN: TMaskEdit;
    Label6: TLabel;
    edtDESC_NAME: TEdit;
    qryHangmok: TQuery;
    qryGmgn: TQuery;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    edtFind: TEdit;
    Label2: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    Label11: TLabel;
    mskDAT_GMGN: TDateEdit;
    btnNextNum: TBitBtn;
    qryU_Bunju: TQuery;
    valChange: TValEdit;
    btnPrevNum: TBitBtn;
    RG_sort: TRadioGroup;
    qryPart02: TQuery;
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
    edtPValue1: TEdit;
    edtPValue2: TEdit;
    edtPValue3: TEdit;
    edtPValue4: TEdit;
    edtPValue5: TEdit;
    edtPValue6: TEdit;
    edtPValue7: TEdit;
    edtPValue8: TEdit;
    edtPValue9: TEdit;
    qryPrev: TQuery;
    qryinjouk: TQuery;
    Label7: TLabel;
    Edt_smp: TEdit;
    Label12: TLabel;
    Edt_ID: TEdit;
    pnlNum10: TPanel;
    pnlName10: TPanel;
    edtValue10: TEdit;
    edtPValue10: TEdit;
    pnlNum11: TPanel;
    pnlName11: TPanel;
    edtValue11: TEdit;
    edtPValue11: TEdit;
    edtValue12: TEdit;
    pnlName12: TPanel;
    pnlNum12: TPanel;
    pnlNum13: TPanel;
    pnlName13: TPanel;
    edtValue13: TEdit;
    pnlNum14: TPanel;
    pnlName14: TPanel;
    edtValue14: TEdit;
    pnlNum15: TPanel;
    pnlName15: TPanel;
    edtValue15: TEdit;
    pnlNum16: TPanel;
    pnlName16: TPanel;
    edtValue16: TEdit;
    pnlNum17: TPanel;
    pnlName17: TPanel;
    edtValue17: TEdit;
    edtPValue12: TEdit;
    edtPValue13: TEdit;
    edtPValue14: TEdit;
    edtPValue15: TEdit;
    edtPValue16: TEdit;
    edtPValue17: TEdit;
    Panel18: TPanel;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryPf_hm: TQuery;
    Label13: TLabel;
    Edt_bjjs: TEdit;
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
    edtPValue18: TEdit;
    edtPValue19: TEdit;
    edtPValue20: TEdit;
    edtPValue21: TEdit;
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
    edtPValue22: TEdit;
    edtPValue23: TEdit;
    edtPValue24: TEdit;
    edtPValue25: TEdit;
    edtValue26: TEdit;
    edtPValue26: TEdit;
    pnlName26: TPanel;
    pnlNum26: TPanel;
    Label15: TLabel;
    Cmb_gubn: TComboBox;
    rbBarcode: TRadioButton;
    MEdt_Barcode: TMaskEdit;
    qryProfile_H025: TQuery;
    qryProfile: TQuery;
    qryCheck: TQuery;
    Label16: TLabel;
    mskPDAT_BUNJU: TDateEdit;
    Label17: TLabel;
    edtPNUM_BUNJU: TEdit;
    Label18: TLabel;
    Panel2: TPanel;
    Panel3: TPanel;
    P_1: TPanel;
    ckGUBN_JUNGDO1: TCheckBox;
    mskDAT_SOLVE1: TDateEdit;
    mskDAT_SOLVE2: TDateEdit;
    ckGUBN_JUNGDO2: TCheckBox;
    P_2: TPanel;
    P_3: TPanel;
    ckGUBN_JUNGDO3: TCheckBox;
    mskDAT_SOLVE3: TDateEdit;
    mskDAT_SOLVE8: TDateEdit;
    ckGUBN_JUNGDO8: TCheckBox;
    P_8: TPanel;
    P_9: TPanel;
    ckGUBN_JUNGDO9: TCheckBox;
    mskDAT_SOLVE9: TDateEdit;
    P_6: TPanel;
    ckGUBN_JUNGDO6: TCheckBox;
    mskDAT_SOLVE6: TDateEdit;
    mskDAT_SOLVE5: TDateEdit;
    ckGUBN_JUNGDO5: TCheckBox;
    P_5: TPanel;
    P_4: TPanel;
    ckGUBN_JUNGDO4: TCheckBox;
    mskDAT_SOLVE4: TDateEdit;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel30: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    RGrp_preGulkwa: TRadioGroup;
    qryinjouk1: TQuery;
    qryGum_bul: TQuery;
    qrykicho: TQuery;
    qryGumgin_Check: TQuery;
    Label21: TLabel;
    edt_sinjang: TEdit;
    Label22: TLabel;
    edt_chejung: TEdit;
    Label23: TLabel;
    edt_gumgin_Check: TEdit;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    mskDate: TDateEdit;
    mskNum: TMaskEdit;
    rbDate: TRadioButton;
    Label14: TLabel;
    MEdt_SampE: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label19: TLabel;
    EdtFind2: TEdit;
    chk_CRM: TCheckBox;
    qryU_check_CRM: TQuery;
    qryI_Check_CRM: TQuery;
    qry_Check_CRM: TQuery;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure btnPInsertClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Exit(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure edtFindExit(Sender: TObject);
    procedure mskNumExit(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure UP_MoveNum(Sender: TObject);
    procedure edtValue1Change(Sender: TObject);
    procedure Cmb_gubnChange(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure EdtFind2Exit(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_jisa, UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vNum_samp,
    UV_vCod_hul,    UV_vCod_cancer, UV_vCod_jangbi, UV_vCod_chuga,
    UV_vGubn_nosin, UV_vGubn_nsyh, UV_vGubn_adult, UV_vGubn_adyh,
    UV_vGubn_life,  UV_vGubn_lfyh, UV_vGubn_bogun, UV_vGubn_bgyh,
    UV_vGubn_agab,  UV_vGubn_agyh, UV_vGubn_tkgm,
    UV_vCod_bjjs, UV_vDat_bunju, UV_vNum_bunju, UV_vDesc_glkwa, UV_vDesc_Pglkwa, UV_vPart02, UV_vABO_chk,UV_vGumgin_check,

    UV_vH001, UV_vH002, UV_vH003, UV_vH004, UV_vH005, UV_vH006, UV_vH007,
    UV_vH008, UV_vH009, UV_vH010, UV_vH011, UV_vH012, UV_vH013, UV_vH014,
    UV_vH015, UV_vH016, UV_vH017, UV_vH018, UV_vH019, UV_vH020, UV_vH021,
    UV_vH022, UV_vH023, UV_vH024, UV_vH025, UV_vH035 : Variant;

    iCntCheck : Integer;

    // 항목코드, 시작위치, 마지막위치, 값
    UV_sValue : array[1..26] of String;
    UV_sGubn  : array[1..26] of String;
    UV_sUnit  : array[1..26] of String;

    // part구분
    sHangmok : String;

    // 인적쿼리저장
    UV_vQuery : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
    procedure UP_EditDisplay(Sender : TObject);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    Procedure Hangmok_Set_H023;
    Procedure Hangmok_Set_H025;

    procedure UP_DChange(Sender: TObject);    
  public
    { Public declarations }
  end;

var
  frmLD1I13: TfrmLD1I13;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I016;

{$R *.DFM}

Procedure TfrmLD1I13.Hangmok_Set_H023;
Var I, J, K : Integer;
Begin
   Inherited;
//PRofile Reading
   With qryProfile do
   Begin
      Close;
      //장비
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_jangbi').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
      //혈액
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Hul').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
      //Cancer
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Cancer').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
   End;
   //추가코드
   sHangmok := sHangmok + qryGlkwa.FieldByName('Cod_chuga').AsString;
   //특검
   If qryGlkwa.FieldByName('Gubn_tkgm').AsString <> '' Then
   begin
      With qryTkgum Do
      Begin
         Close;
         ParamByName('Cod_jisa').AsString  := qryGlkwa.FieldByname('Cod_jisa').AsString;
         ParamByName('Dat_gmgn').AsString  := qryGlkwa.FieldByname('DaT_gmgn').AsString;
         ParamByName('NuM_Jumin').AsString := qryGlkwa.FieldByname('NuM_Jumin').AsString;
         Open;
         If RecordCount > 0 Then
         Begin
            I := Length(FieldByName('Cod_prf').AsString);
            J := Round(I / 4);
            For K := 0 To J - 1 Do
            Begin
               With qryProfile do
               //특검 Profile
               Begin
                  Close;
                  ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4);
                  Open;
                  For I := 1 To Recordcount Do
                  Begin
                     sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                     Next;
                  End;
                  Close;
               End;
            End;
         End;
         sHangmok := sHangmok + FieldByName('Cod_chuga').AsString;
         Close;
      End;
   end;
End;

Procedure TfrmLD1I13.Hangmok_Set_H025;
Var I, J, K : Integer;
Begin
   Inherited;
   //PRofile Reading
   With qryProfile_H025 do
   Begin
      Close;
      //장비
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_jangbi').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
      //혈액
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Hul').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
      //Cancer
      ParamByName('In_Cod_Pf').AsString := qryGlkwa.FieldByName('Cod_Cancer').AsString;
      Open;
      For I := 1 To Recordcount Do
      Begin
         sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
         Next;
      End;
      Close;
   End;
   //추가코드
   sHangmok := sHangmok + qryGlkwa.FieldByName('Cod_chuga').AsString;
   //특검
   If qryGlkwa.FieldByName('Gubn_tkgm').AsString <> '' Then
   begin
      With qryTkgum Do
      Begin
         Close;
         ParamByName('Cod_jisa').AsString  := qryGlkwa.FieldByname('Cod_jisa').AsString;
         ParamByName('Dat_gmgn').AsString  := qryGlkwa.FieldByname('DaT_gmgn').AsString;
         ParamByName('NuM_Jumin').AsString := qryGlkwa.FieldByname('NuM_Jumin').AsString;
         Open;
         If RecordCount > 0 Then
         Begin
            I := Length(FieldByName('Cod_prf').AsString);
            J := Round(I / 4);
            For K := 0 To J - 1 Do
            Begin
               With qryProfile_H025 do
               //특검 Profile
               Begin
                  Close;
                  ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,J * 4 + 1, 4);
                  Open;
                  For I := 1 To Recordcount Do
                  Begin
                     sHangmok := sHangmok + FieldByName('Cod_hm').AsString;
                     Next;
                  End;
                  Close;
               End;
            End;
         End;

         sHangmok := sHangmok + FieldByName('Cod_chuga').AsString;
         Close;
      End;
   end;
End;

function  TfrmLD1I13.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I13.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_id').AsString      := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD1I13.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;
      // Column Title
      if rbDate.Checked then
      begin
         Col[1].Heading := '분주번호';
         Col[2].Heading := '혈액샘플No';
         Col[3].Heading := '검진자명';
      end
      else if rbJumin.Checked then
      begin
         Col[1].Heading := '검진일자';
         Col[2].Heading := '혈액샘플No';
         Col[3].Heading := '검진자명';
      end;
      // Set Alignment
      Col[1].Alignment := taCenter;
      Col[2].Alignment := taCenter;
      Col[3].Alignment := taCenter;
   end;
end;

procedure TfrmLD1I13.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);

      UV_vCod_hul     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cancer  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jangbi  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_chuga   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nosin  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_nsyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adult  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_adyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_life   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_lfyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bogun  := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_bgyh   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agab   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_agyh   := VarArrayCreate([0, 0], varOleStr);

      UV_vGubn_tkgm   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_glkwa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pglkwa := VarArrayCreate([0, 0], varOleStr);
      UV_vPart02      := VarArrayCreate([0, 0], varOleStr);

      UV_vH001        := VarArrayCreate([0, 0], varOleStr);
      UV_vH002        := VarArrayCreate([0, 0], varOleStr);
      UV_vH003        := VarArrayCreate([0, 0], varOleStr);
      UV_vH004        := VarArrayCreate([0, 0], varOleStr);
      UV_vH005        := VarArrayCreate([0, 0], varOleStr);
      UV_vH006        := VarArrayCreate([0, 0], varOleStr);
      UV_vH007        := VarArrayCreate([0, 0], varOleStr);
      UV_vH008        := VarArrayCreate([0, 0], varOleStr);
      UV_vH009        := VarArrayCreate([0, 0], varOleStr);
      UV_vH010        := VarArrayCreate([0, 0], varOleStr);
      UV_vH011        := VarArrayCreate([0, 0], varOleStr);
      UV_vH012        := VarArrayCreate([0, 0], varOleStr);
      UV_vH013        := VarArrayCreate([0, 0], varOleStr);
      UV_vH014        := VarArrayCreate([0, 0], varOleStr);
      UV_vH015        := VarArrayCreate([0, 0], varOleStr);
      UV_vH016        := VarArrayCreate([0, 0], varOleStr);
      UV_vH017        := VarArrayCreate([0, 0], varOleStr);
      UV_vH018        := VarArrayCreate([0, 0], varOleStr);
      UV_vH019        := VarArrayCreate([0, 0], varOleStr);
      UV_vH020        := VarArrayCreate([0, 0], varOleStr);
      UV_vH021        := VarArrayCreate([0, 0], varOleStr);
      UV_vH022        := VarArrayCreate([0, 0], varOleStr);
      UV_vH023        := VarArrayCreate([0, 0], varOleStr);
      UV_vH024        := VarArrayCreate([0, 0], varOleStr);
      UV_vH025        := VarArrayCreate([0, 0], varOleStr);
      UV_vH035        := VarArrayCreate([0, 0], varOleStr);
      UV_vGumgin_check := VarArrayCreate([0, 0], varOleStr);
      UV_vABO_chk     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);

      VarArrayRedim(UV_vCod_hul,     iLength);
      VarArrayRedim(UV_vCod_cancer,  iLength);
      VarArrayRedim(UV_vCod_jangbi,  iLength);
      VarArrayRedim(UV_vCod_chuga,   iLength);
      VarArrayRedim(UV_vGubn_nosin,  iLength);
      VarArrayRedim(UV_vGubn_nsyh,   iLength);
      VarArrayRedim(UV_vGubn_adult,  iLength);
      VarArrayRedim(UV_vGubn_adyh,   iLength);
      VarArrayRedim(UV_vGubn_life,   iLength);
      VarArrayRedim(UV_vGubn_lfyh,   iLength);
      VarArrayRedim(UV_vGubn_bogun,  iLength);
      VarArrayRedim(UV_vGubn_bgyh,   iLength);
      VarArrayRedim(UV_vGubn_agab,   iLength);
      VarArrayRedim(UV_vGubn_agyh,   iLength);

      VarArrayRedim(UV_vGubn_tkgm,   iLength);
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vDat_bunju,   iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_glkwa,  iLength);
      VarArrayRedim(UV_vDesc_Pglkwa, iLength);
      VarArrayRedim(UV_vPart02,      iLength);

      VarArrayRedim(UV_vH001,        iLength);
      VarArrayRedim(UV_vH002,        iLength);
      VarArrayRedim(UV_vH003,        iLength);
      VarArrayRedim(UV_vH004,        iLength);
      VarArrayRedim(UV_vH005,        iLength);
      VarArrayRedim(UV_vH006,        iLength);
      VarArrayRedim(UV_vH007,        iLength);
      VarArrayRedim(UV_vH008,        iLength);
      VarArrayRedim(UV_vH009,        iLength);
      VarArrayRedim(UV_vH010,        iLength);
      VarArrayRedim(UV_vH011,        iLength);
      VarArrayRedim(UV_vH012,        iLength);
      VarArrayRedim(UV_vH013,        iLength);
      VarArrayRedim(UV_vH014,        iLength);
      VarArrayRedim(UV_vH015,        iLength);
      VarArrayRedim(UV_vH016,        iLength);
      VarArrayRedim(UV_vH017,        iLength);
      VarArrayRedim(UV_vH018,        iLength);
      VarArrayRedim(UV_vH019,        iLength);
      VarArrayRedim(UV_vH020,        iLength);
      VarArrayRedim(UV_vH021,        iLength);
      VarArrayRedim(UV_vH022,        iLength);
      VarArrayRedim(UV_vH023,        iLength);
      VarArrayRedim(UV_vH024,        iLength);
      VarArrayRedim(UV_vH025,        iLength);
      VarArrayRedim(UV_vH035,        iLength);
      VarArrayRedim(UV_vGumgin_check, iLength);
      VarArrayRedim(UV_vABO_chk,     iLength);
   end;
end;

function TfrmLD1I13.UF_Condition : Boolean;
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

function TfrmLD1I13.UF_Save : Boolean;
var i, iTemp : Integer;
    Part02 : String;
begin
   Result := False;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   //CRM 체크 버튼, 작업 여부와 상관없이 체크 되면 저장 되도록 요청
   with qry_Check_CRM do
   begin
      Close;
      qry_Check_CRM.ParamByName('COD_jisa').AsString := UV_vCod_jisa[i];
      qry_Check_CRM.ParamByName('DAT_gmgn').AsString := UV_vdat_gmgn[i];
      qry_Check_CRM.ParamByName('NUM_ID').AsString   := UV_vNum_id[i];
      Open;
      if qry_Check_CRM.RecordCount > 0 then
      Begin
         if chk_CRM.Checked = false then
         begin
         with qryU_Check_CRM do
             begin
                ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
                ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
                ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
                ParamByName('YSNO_CRM').AsString    := 'N';
                ParamByName('DAT_update').AsString  := GV_sToday;
                ParamByName('COD_update').AsString  := GV_sUserId;
                Execsql;

                GP_query_log(qryU_Check_CRM, 'LD1I13', 'U', 'Y');
             end;
          end
          else
             with qryU_Check_CRM do
             begin
                ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
                ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
                ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
                ParamByName('YSNO_CRM').AsString    := 'Y';
                ParamByName('DAT_update').AsString  := GV_sToday;
                ParamByName('COD_update').AsString  := GV_sUserId;
                Execsql;

                GP_query_log(qryU_Check_CRM, 'LD1I13', 'U', 'Y');
             end;
      end
      else
      if chk_CRM.Checked = TRUE then
      begin
      with qryI_Check_CRM do
          begin
             ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
             ParamByName('DAT_gmgn').AsString    := UV_vdat_gmgn[i];
             ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
             ParamByName('YSNO_CRM').AsString    := 'Y';
             ParamByName('DAT_INSERT').AsString  := GV_sToday;
             ParamByName('COD_INSERT').AsString  := GV_sUserId;
             ParamByName('DAT_update').AsString  := '';
             ParamByName('COD_update').AsString  := '';
             Execsql;

             GP_query_log(qryI_Check_CRM, 'LD1I13', 'I', 'Y');
          end;
      end;
   end;
   // 작업중인지 Check
   if not UV_bEdit then exit;

   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   with qryPart02 do
   begin
      Close;
      qryPart02.ParamByName('dat_bunju').AsString := mskDAT_BUNJU.Text;
      qryPart02.ParamByName('num_bunju').AsString := edtNUM_BUNJU.Text;
      Open;

      if qryPart02.RecordCount > 0 then
      begin
         Part02 := '';
         if FieldByName('desc_glkwa1').AsString <> '' then
         begin
            for iTemp := 1 to 200 - Length(FieldByName('desc_glkwa1').AsString) do
                Part02 := Part02 + ' ';
            Part02 := FieldByName('desc_glkwa1').AsString + Part02;
         end
         else
         begin
            for iTemp := 1 to 2 do
                Part02 := Part02 + '                                                                                                    ';
         end;
      end;
   end;

   //H001	적혈구수	02	01 	1	6
   Part02 := GF_RPad(edtValue1.text, 6, ' ') + Copy(Part02, 7, Length(Part02));
   //H002	혈색소  	02	02 	7	12
   Part02 := Copy(Part02, 1,  6)   + GF_RPad(edtValue2.text, 6, ' ') + Copy(Part02,  13, Length(Part02));
   //H003	Hematocrit	02	03 	13	18
   Part02 := Copy(Part02, 1, 12)   + GF_RPad(edtValue3.text, 6, ' ') + Copy(Part02,  19, Length(Part02));
   //H004	MCV	        02	04 	19	24
   Part02 := Copy(Part02, 1, 18)   + GF_RPad(edtValue4.text, 6, ' ') + Copy(Part02,  25, Length(Part02));
   //H005	MCH	        02	05 	25	30
   Part02 := Copy(Part02, 1, 24)   + GF_RPad(edtValue5.text, 6, ' ') + Copy(Part02,  31, Length(Part02));
   //H006	MCHC	        02	06 	31	36
   Part02 := Copy(Part02, 1, 30)   + GF_RPad(edtValue6.text, 6, ' ') + Copy(Part02,  37, Length(Part02));
   //H007	RDW	        02	07 	37	42
   Part02 := Copy(Part02, 1, 36)   + GF_RPad(edtValue7.text, 6, ' ') + Copy(Part02,  43, Length(Part02));
   //H008	혈소판수	02	08 	43	48
   Part02 := Copy(Part02, 1, 42)   + GF_RPad(edtValue8.text, 6, ' ') + Copy(Part02,  49, Length(Part02));
   //H009	MPV	        02	09 	49	54
   Part02 := Copy(Part02, 1, 48)   + GF_RPad(edtValue9.text, 6, ' ') + Copy(Part02,  55, Length(Part02));
   //H010	PDW	        02	10 	55	60
   Part02 := Copy(Part02, 1, 54 )  + GF_RPad(edtValue10.text, 6, ' ') + Copy(Part02,  61, Length(Part02));
   //H011	백혈구수	02	11 	61	66
   Part02 := Copy(Part02, 1, 60)   + GF_RPad(edtValue11.text, 6, ' ') + Copy(Part02,  67, Length(Part02));
   //H012	분획호중구	02	12 	67	72
   Part02 := Copy(Part02, 1, 66)   + GF_RPad(edtValue12.text, 6, ' ') + Copy(Part02,  73, Length(Part02));
   //H013	봉상호중구	02	13 	73	78
   Part02 := Copy(Part02, 1, 72)   + GF_RPad(edtValue13.text, 6, ' ') + Copy(Part02,  79, Length(Part02));
   //H014	임파구	        02	14 	79	84
   Part02 := Copy(Part02, 1, 78)   + GF_RPad(edtValue14.text, 6, ' ') + Copy(Part02,  85, Length(Part02));
   //H015	단핵구	        02	15 	85	90
   Part02 := Copy(Part02, 1, 84)   + GF_RPad(edtValue15.text, 6, ' ') + Copy(Part02,  91, Length(Part02));
   //H016	호산구	        02	16 	91	96
   Part02 := Copy(Part02, 1, 90)   + GF_RPad(edtValue16.text, 6, ' ') + Copy(Part02,  97, Length(Part02));
   //H017	호염기성구	02	17 	97	102
   Part02 := Copy(Part02, 1, 96)   + GF_RPad(edtValue17.text, 6, ' ') + Copy(Part02, 103, Length(Part02));
   //H018	후골수구	2	18	103	108
   Part02 := Copy(Part02, 1, 102)  + GF_RPad(edtValue18.text, 6, ' ') + Copy(Part02, 109, Length(Part02));
   //H019	골수구	        2	19	109	114
   Part02 := Copy(Part02, 1, 108)  + GF_RPad(edtValue19.text, 6, ' ') + Copy(Part02, 115, Length(Part02));
   //H020	전골수세포	2	20	115	120
   Part02 := Copy(Part02, 1, 114)  + GF_RPad(edtValue20.text, 6, ' ') + Copy(Part02, 121, Length(Part02));
   //H021	골수아세포	2	21	121	126
   Part02 := Copy(Part02, 1, 120)  + GF_RPad(edtValue21.text, 6, ' ') + Copy(Part02, 127, Length(Part02));
   //H022	유핵적혈구	2	22	127	132
   Part02 := Copy(Part02, 1, 126)  + GF_RPad(edtValue22.text, 6, ' ') + Copy(Part02, 133, Length(Part02));
   //H023	ABO형	        2	23	133	138
   Part02 := Copy(Part02, 1, 132)  + GF_RPad(edtValue23.text, 6, ' ') + Copy(Part02, 139, Length(Part02));
   //H024	Rh형	        2	24	139	144
   Part02 := Copy(Part02, 1, 138)  + GF_RPad(edtValue24.text, 6, ' ') + Copy(Part02, 145, Length(Part02));
   //H025	ESR	        2	25	145	150
   Part02 := Copy(Part02, 1, 144)  + GF_RPad(edtValue25.text, 6, ' ') + Copy(Part02, 151, Length(Part02));
   //H035	PCT	        2	31	175	180
   Part02 := Copy(Part02, 1, 174)  + GF_RPad(edtValue26.text, 6, ' ') + Copy(Part02, 181, Length(Part02));



   Part02 := 'A' + Part02;
   Part02 := Trim(Part02);
   Part02 := Copy(Part02, 2, Length(Part02)-1);

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   // DB 작업
   dmComm.database.StartTransaction;
   try
      with qryU_Glkwa do
      begin
         ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString   := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString   := edtNUM_BUNJU.Text;
         ParamByName('GUBN_PART').AsString   := '02';
         ParamByName('desc_glkwa1').AsString := Part02;
         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         Execsql;

         GP_query_log(qryU_Glkwa, 'LD1I13', 'Q', 'Y');
      end;

      with qryU_Bunju do
      begin
         ParamByName('COD_jisa').AsString   := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString  := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString  := edtNUM_BUNJU.Text;
         Execsql;

         GP_query_log(qryU_Bunju, 'LD1I13', 'Q', 'Y');
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

   // Grid에 자료를 반영
   with grdMaster do
   begin
      // 현재 Field의 자료를 Grid관련 Variant변수에 할당
      UV_vPart02[i] := Part02;
      UV_vH001[i]   := Trim(copy(Part02,   1, 6));
      UV_vH002[i]   := Trim(copy(Part02,   7, 6));
      UV_vH003[i]   := Trim(copy(Part02,  13, 6));
      UV_vH004[i]   := Trim(copy(Part02,  19, 6));
      UV_vH005[i]   := Trim(copy(Part02,  25, 6));
      UV_vH006[i]   := Trim(copy(Part02,  31, 6));
      UV_vH007[i]   := Trim(copy(Part02,  37, 6));
      UV_vH008[i]   := Trim(copy(Part02,  43, 6));
      UV_vH009[i]   := Trim(copy(Part02,  49, 6));
      UV_vH010[i]   := Trim(copy(Part02,  55, 6));
      UV_vH011[i]   := Trim(copy(Part02,  61, 6));
      UV_vH012[i]   := Trim(copy(Part02,  67, 6));
      UV_vH013[i]   := Trim(copy(Part02,  73, 6));
      UV_vH014[i]   := Trim(copy(Part02,  79, 6));
      UV_vH015[i]   := Trim(copy(Part02,  85, 6));
      UV_vH016[i]   := Trim(copy(Part02,  91, 6));
      UV_vH017[i]   := Trim(copy(Part02,  97, 6));
      UV_vH018[i]   := Trim(copy(Part02, 103, 6));
      UV_vH019[i]   := Trim(copy(Part02, 109, 6));
      UV_vH020[i]   := Trim(copy(Part02, 115, 6));
      UV_vH021[i]   := Trim(copy(Part02, 121, 6));
      UV_vH022[i]   := Trim(copy(Part02, 127, 6));
      UV_vH023[i]   := Trim(copy(Part02, 133, 6));
      UV_vH024[i]   := Trim(copy(Part02, 139, 6));
      UV_vH025[i]   := Trim(copy(Part02, 145, 6));
      UV_vH035[i]   := Trim(copy(Part02, 175, 6));
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD1I13.btnSaveClick(Sender: TObject);
begin
   inherited;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I13.FormCreate(Sender: TObject);
var i : integer;
begin
   inherited;

   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   UV_vQuery := qryinjouk.SQL.Text;
   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   UV_sBasicSQL := qryGlkwa.SQL.Text;
end;

procedure TfrmLD1I13.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : begin
             if rbDate.Checked then
                Value := UV_vNum_bunju[DataRow-1]
             else if rbJumin.Checked then
                Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
          end;
      2 : Value := UV_vNum_Samp[DataRow-1];
      3 : Value := UV_vDesc_name[DataRow-1];
   end;
end;

procedure TfrmLD1I13.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy, sCheck : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;
   // Grid 초기화
   grdMaster.Rows := 0;
   // Grid 환경설정
   UP_GridSet;

   // 분주번호에 0을 채워줌
   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> ''   then mskTo.Text   := GF_LPad(Trim(mskTo.Text), 4, '0');

   // Query 수행 & 배열에 저장
   index := 0;
   with qryGlkwa do
   begin
      Active := False;
      if GV_sJisa = '00' then
         sWhere := ' B.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''''
      else
         sWhere := ' B.cod_jisa = ''' + Copy(GV_sUserId, 1, 2) + '''';

      //연구소+센터만 조회
      sWhere := sWhere + ' and B.gubn_hulgum = ''3'' ';

      if rbDate.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND A.dat_bunju = ''' + mskDate.Text + '''';

         if Cmb_gubn.Text = '분주번호' then
         begin
            if Trim(mskFrom.Text) <> '' then  sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
            if Trim(mskTo.Text) <> '' then    sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
         end
         else if Cmb_gubn.Text = '샘플번호' then
         begin
            if Trim(MEdt_SampS.Text) <> '' then sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
            if Trim(MEdt_SampE.Text) <> '' then sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
         end
         else
         begin
            if Trim(MEdt_SampS.Text) <> '' then sWhere := sWhere + ' AND B.num_samp >= ''' + MEdt_SampS.Text + '''';
            if Trim(MEdt_SampE.Text) <> '' then sWhere := sWhere + ' AND B.num_samp <= ''' + MEdt_SampE.Text + '''';
         end;

         case RG_sort.ItemIndex of
            0 : sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';  //20140503
            1 : sOrderBy := ' ORDER BY A.num_bunju, A.gubn_part ';
            else
               sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
         end
      end
      else if rbJumin.Checked then
      begin
         sWhere := ' WHERE ' + sWhere + 'AND B.num_jumin = ''' + mskJumin.Text + '''';
         sOrderBy := ' ORDER BY B.dat_gmgn, A.gubn_part ';
      end
      else if rbBarcode.Checked then
      begin
         sWhere := sWhere + ' AND B.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
         sWhere := sWhere + ' AND B.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';

         sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
      end;

      sWhere := sWhere + ' and A.gubn_part = ''02'' ';

      ParamByName('sWhere').Asmemo     :=  sWhere;
      ParamByName('sOrderBy').AsString :=  sOrderBy;
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I13', 'Q', 'Y');
         if Cmb_gubn.Text = '혈액형검사자' then
         begin
            while not qryGlkwa.Eof do
            begin
               sHangmok := '';

//[2012.03.07]혈액형 검사자만 보이게..
//------------------------------------------------------------------------------
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then sCheck := '99';
               {
               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then sCheck := '99';
               }
               If FieldByname('Cod_hul').Asstring <> '' Then sCheck := '99';  //20140726 곽윤설
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then sCheck := '99';

               If (sCheck = '99') Or (FieldByName('Num_Bunju').AsInteger > 4999) Then Hangmok_Set_H023;
//------------------------------------------------------------------------------

               If (Pos('H023',sHangmok) > 0) or (Pos('H024',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_jisa[index]   := FieldByName('COD_jisa').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vCod_bjjs[index]   := FieldByName('cod_bjjs').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vNum_samp[index]   := FieldByName('NUM_SAMP').AsString;

                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_cancer[index] := FieldByName('Cod_cancer').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;
                  UV_vGubn_nosin[index] := FieldByName('Gubn_nosin').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('Gubn_nsyh').AsString;
                  UV_vGubn_adult[index] := FieldByName('Gubn_adult').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('Gubn_adyh').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  UV_vGubn_bogun[index] := FieldByName('Gubn_bogun').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('Gubn_bgyh').AsString;
                  UV_vGubn_agab[index]  := FieldByName('Gubn_agab').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('Gubn_agyh').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('Gubn_tkgm').AsString;

                  UV_vH001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   1, 6));
                  UV_vH002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   7, 6));
                  UV_vH003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  13, 6));
                  UV_vH004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  19, 6));
                  UV_vH005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  25, 6));
                  UV_vH006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  31, 6));
                  UV_vH007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  37, 6));
                  UV_vH008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  43, 6));
                  UV_vH009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  49, 6));
                  UV_vH010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  55, 6));
                  UV_vH011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  61, 6));
                  UV_vH012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  67, 6));
                  UV_vH013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  73, 6));
                  UV_vH014[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  79, 6));
                  UV_vH015[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  85, 6));
                  UV_vH016[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  91, 6));
                  UV_vH017[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  97, 6));
                  UV_vH018[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 103, 6));
                  UV_vH019[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 109, 6));
                  UV_vH020[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 115, 6));
                  UV_vH021[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 121, 6));
                  UV_vH022[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 127, 6));
                  UV_vH023[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 133, 6));
                  UV_vH024[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 139, 6));
                  UV_vH025[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 145, 6));
                  UV_vH035[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 175, 6));
                  UV_vPart02[index]     := FieldByName('desc_glkwa1').AsString;

                  UV_vABO_chk[index]    := 'Y';
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';

                   qryGumgin_Check.Active := False;
                   qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                   qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                   qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                   qryGumgin_Check.Active := True;
                   if qryGumgin_Check.RecordCount > 0 then
                   begin
                        if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                           UV_vGumgin_check[index]   :='생리중';

                   end;
                  Inc(index);
               end;

               Next;
            end;
         end
         else if Cmb_gubn.Text = 'ESR검사자' then
         begin
            while not qryGlkwa.Eof do
            begin
               sHangmok := '';
//[2012.03.07]혈액형 검사자만 보이게..
//------------------------------------------------------------------------------
               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;

               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then sCheck := '99';

               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then sCheck := '99';


               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then sCheck := '99';

               If FieldByName('Num_Bunju').AsInteger > 4999 Then sCheck := '99';
//------------------------------------------------------------------------------
               Hangmok_Set_H025;

               If (Pos('H025',sHangmok) > 0) Then
               begin
                  UP_VarMemSet(Index);

                  UV_vCod_jisa[index]   := FieldByName('COD_jisa').AsString;
                  UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
                  UV_vCod_bjjs[index]   := FieldByName('cod_bjjs').AsString;
                  UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
                  UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
                  UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
                  UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
                  UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
                  UV_vNum_samp[index]   := FieldByName('NUM_SAMP').AsString;

                  UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
                  UV_vCod_cancer[index] := FieldByName('Cod_cancer').AsString;
                  UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
                  UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;
                  UV_vGubn_nosin[index] := FieldByName('Gubn_nosin').AsString;
                  UV_vGubn_nsyh[index]  := FieldByName('Gubn_nsyh').AsString;
                  UV_vGubn_adult[index] := FieldByName('Gubn_adult').AsString;
                  UV_vGubn_adyh[index]  := FieldByName('Gubn_adyh').AsString;
                  UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
                  UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
                  UV_vGubn_bogun[index] := FieldByName('Gubn_bogun').AsString;
                  UV_vGubn_bgyh[index]  := FieldByName('Gubn_bgyh').AsString;
                  UV_vGubn_agab[index]  := FieldByName('Gubn_agab').AsString;
                  UV_vGubn_agyh[index]  := FieldByName('Gubn_agyh').AsString;
                  UV_vGubn_tkgm[index]  := FieldByName('Gubn_tkgm').AsString;

                  UV_vH001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   1, 6));
                  UV_vH002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   7, 6));
                  UV_vH003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  13, 6));
                  UV_vH004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  19, 6));
                  UV_vH005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  25, 6));
                  UV_vH006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  31, 6));
                  UV_vH007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  37, 6));
                  UV_vH008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  43, 6));
                  UV_vH009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  49, 6));
                  UV_vH010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  55, 6));
                  UV_vH011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  61, 6));
                  UV_vH012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  67, 6));
                  UV_vH013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  73, 6));
                  UV_vH014[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  79, 6));
                  UV_vH015[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  85, 6));
                  UV_vH016[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  91, 6));
                  UV_vH017[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  97, 6));
                  UV_vH018[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 103, 6));
                  UV_vH019[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 109, 6));
                  UV_vH020[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 115, 6));
                  UV_vH021[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 121, 6));
                  UV_vH022[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 127, 6));
                  UV_vH023[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 133, 6));
                  UV_vH024[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 139, 6));
                  UV_vH025[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 145, 6));
                  UV_vH035[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 175, 6));
                  UV_vPart02[index]     := FieldByName('desc_glkwa1').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성';

                   qryGumgin_Check.Active := False;
                   qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
                   qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
                   qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
                   qryGumgin_Check.Active := True;
                   if qryGumgin_Check.RecordCount > 0 then
                   begin
                        if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                           UV_vGumgin_check[index]   :='생리중';

                   end;
                  if (sCheck = '99') then UV_vABO_chk[index] := 'Y'
                  else                    UV_vABO_chk[index] := 'N';

                  Inc(index);
               end;

               Next;
            end;
         end
         else
         begin
            UP_VarMemSet(RecordCount - 1);

            while not qryGlkwa.Eof do
            begin
//[2012.03.07]혈액형 검사자만 보이게..
//------------------------------------------------------------------------------

               //단체 Check
               With qryCheck Do
               Begin
                  Close;
                  ParamByName('In_Cod_dc').AsString := qryGlkwa.FieldByname('Cod_dc').AsString;
                  ParamByName('num_year').AsString  := copy(qryGlkwa.FieldByname('DaT_Gmgn').AsString,1,4);
                  Open;
                  sCheck := '';
                  If Recordcount > 0 Then
                     If (qryGlkwa.FieldByName('DaT_Gmgn').AsString >= FieldByname('Dat_Start').AsString) And
                        (qryGlkwa.FieldByName('DaT_Gmgn').AsString <= FieldByname('Dat_End').AsString) Then
                        sCheck := '99';
                  Close;
               End;
             
               //[2008.02.29] 검사추가에 있으면 검사들어감..
               if (pos('H023',FieldByName('cod_chuga').AsString) > 0) or
                  (pos('H024',FieldByName('cod_chuga').AsString) > 0) then sCheck := '99';
             
               //어린이집
               If (FieldByname('Cod_hul').Asstring = 'B501') Or
                  (FieldByname('Cod_hul').Asstring = 'B502') Or
                  (FieldByname('Cod_hul').Asstring = '9023') Or
                  (FieldByname('Cod_hul').Asstring = 'B507') Or
                  (FieldByname('Cod_hul').Asstring = 'B508') Then sCheck := '99';
             
               If Copy(FieldByname('Cod_Jangbi').Asstring,1,2) = 'SS' Then sCheck := '99';
             
               If FieldByName('Num_Bunju').AsInteger > 4999 Then sCheck := '99';
//------------------------------------------------------------------------------

               UV_vCod_jisa[index]   := FieldByName('COD_jisa').AsString;
               UV_vNum_id[index]     := FieldByName('NUM_ID').AsString;
               UV_vCod_bjjs[index]   := FieldByName('cod_bjjs').AsString;
               UV_vDat_bunju[index]  := FieldByName('DAT_BUNJU').AsString;
               UV_vNum_bunju[index]  := FieldByName('NUM_BUNJU').AsString;
               UV_vNum_jumin[index]  := FieldByName('NUM_JUMIN').AsString;
               UV_vDesc_name[index]  := FieldByName('DESC_NAME').AsString;
               UV_vDat_gmgn[index]   := FieldByName('DAT_GMGN').AsString;
               UV_vNum_samp[index]   := FieldByName('NUM_SAMP').AsString;

               UV_vCod_hul[index]    := FieldByName('Cod_hul').AsString;
               UV_vCod_cancer[index] := FieldByName('Cod_cancer').AsString;
               UV_vCod_jangbi[index] := FieldByName('Cod_jangbi').AsString;
               UV_vCod_chuga[index]  := FieldByName('Cod_chuga').AsString;
               UV_vGubn_nosin[index] := FieldByName('Gubn_nosin').AsString;
               UV_vGubn_nsyh[index]  := FieldByName('Gubn_nsyh').AsString;
               UV_vGubn_adult[index] := FieldByName('Gubn_adult').AsString;
               UV_vGubn_adyh[index]  := FieldByName('Gubn_adyh').AsString;
               UV_vGubn_life[index]  := FieldByName('Gubn_life').AsString;
               UV_vGubn_lfyh[index]  := FieldByName('Gubn_lfyh').AsString;
               UV_vGubn_bogun[index] := FieldByName('Gubn_bogun').AsString;
               UV_vGubn_bgyh[index]  := FieldByName('Gubn_bgyh').AsString;
               UV_vGubn_agab[index]  := FieldByName('Gubn_agab').AsString;
               UV_vGubn_agyh[index]  := FieldByName('Gubn_agyh').AsString;
               UV_vGubn_tkgm[index]  := FieldByName('Gubn_tkgm').AsString;

               UV_vH001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   1, 6));
               UV_vH002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,   7, 6));
               UV_vH003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  13, 6));
               UV_vH004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  19, 6));
               UV_vH005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  25, 6));
               UV_vH006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  31, 6));
               UV_vH007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  37, 6));
               UV_vH008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  43, 6));
               UV_vH009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  49, 6));
               UV_vH010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  55, 6));
               UV_vH011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  61, 6));
               UV_vH012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  67, 6));
               UV_vH013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  73, 6));
               UV_vH014[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  79, 6));
               UV_vH015[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  85, 6));
               UV_vH016[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  91, 6));
               UV_vH017[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  97, 6));
               UV_vH018[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 103, 6));
               UV_vH019[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 109, 6));
               UV_vH020[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 115, 6));
               UV_vH021[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 121, 6));
               UV_vH022[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 127, 6));
               UV_vH023[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 133, 6));
               UV_vH024[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 139, 6));
               UV_vH025[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 145, 6));
               UV_vH035[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 175, 6));
               UV_vPart02[index]     := FieldByName('desc_glkwa1').AsString;
               if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
               else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성'
               else  UV_vGumgin_check[index]   := '';
               qryGumgin_Check.Active := False;
               qryGumgin_Check.ParamByName('cod_jisa').AsString := FieldByName('cod_jisa').AsString;
               qryGumgin_Check.ParamByName('NUM_JUMIN').AsString   := FieldByName('NUM_JUMIN').AsString;
               qryGumgin_Check.ParamByName('dat_gmgn').AsString := FieldByName('dat_gmgn').AsString;
               qryGumgin_Check.Active := True;
               if qryGumgin_Check.RecordCount > 0 then
               begin
                    if qryGumgin_Check.FieldByName('Period').AsString ='Y' then
                       UV_vGumgin_check[index]   :='생리중';
               end;
               if (sCheck = '99') then UV_vABO_chk[index] := 'Y'
               else                    UV_vABO_chk[index] := 'N';


               Next;
               Inc(index);
            end;
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

procedure TfrmLD1I13.btnCancelClick(Sender: TObject);
begin
   inherited;

   // 작업중인지 Check & Cancel Confirm Message
   if not UV_bEdit then exit;
   if not GF_MsgBox('Confirmation', 'CANCEL') then exit;

   // 취소시
   // 1.입력이면 화면 Clear                                        
   // 2.수정이면 Re-query
   if UV_iStatus = GC_Insert then
      GP_FieldClear(pnlDetail);

   // 자료 Display
   grdMasterRowChanged(grdMaster, grdMaster.CurrentDataRow, grdMaster.CurrentDataRow);

   // Cancel Mode & Msg
   UP_SetMode('CANCEL');
end;

procedure TfrmLD1I13.UP_Change(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    eRslt, e1, e2, e3 : Extended;
    cod_HM  : array[1..26] of String;
begin
   inherited;

   cod_HM[1]  := 'H001';
   cod_HM[2]  := 'H002';
   cod_HM[3]  := 'H003';
   cod_HM[4]  := 'H004';
   cod_HM[5]  := 'H005';
   cod_HM[6]  := 'H006';
   cod_HM[7]  := 'H007';
   cod_HM[8]  := 'H008';
   cod_HM[9]  := 'H009';
   cod_HM[10] := 'H010';
   cod_HM[11] := 'H011';
   cod_HM[12] := 'H012';
   cod_HM[13] := 'H013';
   cod_HM[14] := 'H014';
   cod_HM[15] := 'H015';
   cod_HM[16] := 'H016';
   cod_HM[17] := 'H017';
   cod_HM[18] := 'H018';
   cod_HM[19] := 'H019';
   cod_HM[20] := 'H020';
   cod_HM[21] := 'H021';
   cod_HM[22] := 'H022';
   cod_HM[23] := 'H023';
   cod_HM[24] := 'H024';
   cod_HM[25] := 'H025';
   cod_HM[26] := 'H035';

   // 주민번호를 이용하여 성별과 나이를 구함
   GP_JuminToAge(mskNUM_JUMIN.Text,mskDAT_GMGN.Text, sMan, iAge);
   if iAge < 10 then sGubn := '1'
   else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
   else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
   else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
   else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
   else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
   else sGubn := '7';
   if sMan = 'F' then sGubn := sGubn + 'f';


   
   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   with qryHangmok do
   begin
      close;
      ParamByName('dat_apply').AsString := mskDAT_GMGN.Text;
      ParamByName('cod_hm').AsString    := cod_HM[StrToInt(Copy(TEdit(Sender).Name, 9, 2))];
      open;
      ims_Low   := FieldByName('IMS_LOW'   + sGubn).AsFloat;
      ims_High  := FieldByName('IMS_HIGH'  + sGubn).AsFloat;
   end;

   if (Trim(TEdit(Sender).Text) <> '') and (TEdit(Sender).Color <> clBtnFace) then
   begin
      if qryHangmok.FieldByName('gubn_hm').AsString = '1' then
      begin
         if (GF_StrToFloat(Trim(TEdit(Sender).Text)) < ims_Low) or (GF_StrToFloat(Trim(TEdit(Sender).Text)) > ims_High) then
            TEdit(Sender).Color     := $00FACDF4
         else
            TEdit(Sender).Color     := clWindow;
      end
      else if qryHangmok.FieldByName('gubn_hm').AsString = '2' then
      begin
         if TEdit(Sender).text = qryHangmok.FieldByName('DESC_UNIT').AsString then
            TEdit(Sender).Color     := clWindow
         else
            TEdit(Sender).Color     := $00FACDF4;
      end;
   end;

   eRslt := 0;

   // 2. 혈액학
   if (Sender = edtValue1) or(Sender = edtValue2) or (Sender = edtValue3) then
   begin
      e1 := GF_StrToFloat(edtValue1.Text);
      e2 := GF_StrToFloat(edtValue2.Text);
      e3 := GF_StrToFloat(edtValue3.Text);

      if Sender = edtValue1 then
      begin
         // H004 = (H003 * 10) / H001
         if e1 <> 0 then eRslt := (e3 * 10) / e1;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');

         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue4.Text := sValue;

         // H005 = (H002 * 10) / H001
         if e1 <> 0 then eRslt := (e2 * 10) / e1;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');
         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue5.Text := sValue;
      end
      else if Sender = edtValue2 then
      begin
         // H005 = (H002 * 10) / H001
         if e1 <> 0 then eRslt := (e2 * 10) / e1;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');
         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue5.Text := sValue;

         // H006 = (H002 * 10) / H003
         if e3 <> 0 then eRslt := (e2 * 10) / e3;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');
         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue6.Text := sValue;
      end
      else if Sender = edtValue3 then
      begin
         // H004 = (H003 * 10) / H001
         if e1 <> 0 then eRslt := (e3 * 10) / e1;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');
         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue4.Text := sValue;

         // H006 = (H002 * 100) / H003
         if e3 <> 0 then eRslt := (e2 * 100) / e3;
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');
         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         edtValue6.Text := sValue;
      end;
   end;


   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I13.UP_DChange(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    cod_HM  : array[1..26] of String;
begin
   inherited;

   cod_HM[1]  := 'H001';
   cod_HM[2]  := 'H002';
   cod_HM[3]  := 'H003';
   cod_HM[4]  := 'H004';
   cod_HM[5]  := 'H005';
   cod_HM[6]  := 'H006';
   cod_HM[7]  := 'H007';
   cod_HM[8]  := 'H008';
   cod_HM[9]  := 'H009';
   cod_HM[10] := 'H010';
   cod_HM[11] := 'H011';
   cod_HM[12] := 'H012';
   cod_HM[13] := 'H013';
   cod_HM[14] := 'H014';
   cod_HM[15] := 'H015';
   cod_HM[16] := 'H016';
   cod_HM[17] := 'H017';
   cod_HM[18] := 'H018';
   cod_HM[19] := 'H019';
   cod_HM[20] := 'H020';
   cod_HM[21] := 'H021';
   cod_HM[22] := 'H022';
   cod_HM[23] := 'H023';
   cod_HM[24] := 'H024';
   cod_HM[25] := 'H025';
   cod_HM[26] := 'H035';

   // 주민번호를 이용하여 성별과 나이를 구함
   GP_JuminToAge(mskNUM_JUMIN.Text,mskDAT_GMGN.Text,sMan, iAge);
   if iAge < 10 then sGubn := '1'
   else if (iAge >= 10) and (iAge <= 19) then sGubn := '2'
   else if (iAge >= 20) and (iAge <= 29) then sGubn := '3'
   else if (iAge >= 30) and (iAge <= 39) then sGubn := '4'
   else if (iAge >= 40) and (iAge <= 49) then sGubn := '5'
   else if (iAge >= 50) and (iAge <= 59) then sGubn := '6'
   else sGubn := '7';
   if sMan = 'F' then sGubn := sGubn + 'f';

   // Edit Change시 Event Procedure를 공유한 후 구분을 Sender로 수행
   with qryHangmok do
   begin
      close;
      ParamByName('dat_apply').AsString := mskDAT_GMGN.Text;
      ParamByName('cod_hm').AsString    := cod_HM[StrToInt(Copy(TEdit(Sender).Name, 9, 2))];
      open;
      ims_Low   := FieldByName('IMS_LOW'   + sGubn).AsFloat;
      ims_High  := FieldByName('IMS_HIGH'  + sGubn).AsFloat;
   end;

   if (Trim(TEdit(Sender).Text) <> '') and (TEdit(Sender).Color <> clBtnFace) then
   begin
      if qryHangmok.FieldByName('gubn_hm').AsString = '1' then
      begin
         if (GF_StrToFloat(Trim(TEdit(Sender).Text)) < ims_Low) or (GF_StrToFloat(Trim(TEdit(Sender).Text)) > ims_High) then
            TEdit(Sender).Color     := $00FACDF4
         else
            TEdit(Sender).Color     := clWindow;
      end
      else if qryHangmok.FieldByName('gubn_hm').AsString = '2' then
      begin
         if TEdit(Sender).text = qryHangmok.FieldByName('DESC_UNIT').AsString then
            TEdit(Sender).Color     := clWindow
         else
            TEdit(Sender).Color     := $00FACDF4;
      end;
   end;
end;

procedure TfrmLD1I13.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var index, iRet, i, iTemp : Integer;
    vCod_chuga : Variant;
    sHmCode, sHmList_02, sSelect, sWhere, sOrderby : String;
begin
   inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   iCntCheck := 0;
   sHmList_02 := '';

   // Grid의 Row가 변경되면 Detail에 자료를 할당
   mskNUM_JUMIN.Text := UV_vNum_jumin[NewRow-1];
   edtDESC_NAME.Text := UV_vDesc_name[NewRow-1];
   mskDAT_GMGN.Text  := UV_vDat_gmgn[NewRow-1];

   Edt_bjjs.Text     := UV_vcod_bjjs[NewRow-1];
   mskDAT_BUNJU.Text := UV_vDat_bunju[NewRow-1];
   edtNUM_BUNJU.Text := UV_vNum_bunju[NewRow-1];
   Edt_smp.Text      := UV_vNum_samp[NewRow-1];
   Edt_ID.Text       := UV_vNum_id[NewRow-1];
   edt_gumgin_Check.Text   := UV_vGumgin_check[NewRow-1];
   qrykicho.Active := False;
   qrykicho.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qrykicho.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qrykicho.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qrykicho.Active := True;
   if qrykicho.RecordCount > 0 then
   begin
        Edt_Sinjang.text:= qrykicho.FieldByname('Sinjang').Asstring;
        Edt_chejung.Text:= qrykicho.FieldByname('chejung').Asstring;
   end;
   qry_Check_CRM.Active := False;
   qry_Check_CRM.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qry_Check_CRM.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qry_Check_CRM.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qry_Check_CRM.Active := True;
   if qry_Check_CRM.FieldByName('ysno_crm').AsString = 'Y' then chk_CRM.checked := true
   else chk_CRM.checked := false;

   //항목찾기
   //---------------------------------------------------------------------------
   sSelect := '';
   with qryPf_hm do
   begin
      Active := False;

      sSelect := ' SELECT P.cod_hm FROM profile_hm P INNER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.gubn_part = ''02'' and H.dat_apply <= ''' + mskDAT_GMGN.Text + '''';
      sSelect := sSelect + ' WHERE P.cod_pf IN (''' + UV_vCod_hul[NewRow-1] + ''',''' + UV_vCod_cancer[NewRow-1] + ''',''' + UV_vCod_jangbi[NewRow-1] + ''')';
      sSelect := sSelect + ' GROUP BY P.cod_hm ';

      SQL.Clear;
      SQL.Add(sSelect);
      Active := True;

      if RecordCount > 0 then
      begin
         while not Eof do
         begin
            sHmList_02 := sHmList_02 + FieldByName('COD_HM').AsString;
            Next;
         end;
      end;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   iRet := GF_MulToSingle(UV_vCod_chuga[NewRow-1], 4, vCod_chuga);
   for i := 0 to iRet-1 do
   begin
      sHmList_02 := sHmList_02 + vCod_chuga[i];
   end;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if UV_vGubn_nosin[NewRow-1] = '1' then
      sHmCode := UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '1', UV_vGubn_nsyh[NewRow-1])
   else if UV_vGubn_adult[NewRow-1] = '1' then
      sHmCode := UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '4', UV_vGubn_adyh[NewRow-1]);

   if UV_vGubn_life[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '7', UV_vGubn_lfyh[NewRow-1]);

   if UV_vGubn_bogun[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '3',UV_vGubn_bgyh[NewRow-1])
   else if UV_vGubn_bogun[NewRow-1] = '2' then
      sHmCode := sHmCode + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '3', UV_vGubn_bgyh[NewRow-1]);

   if UV_vGubn_agab[NewRow-1] = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(UV_vDat_gmgn[NewRow-1], '5',UV_vGubn_agyh[NewRow-1])
   else if UV_vGubn_agab[NewRow-1] = '2' then
      sHmCode := sHmCode + UF_Jae_Hangmok(UV_vCod_jisa[NewRow-1], UV_vNum_id[NewRow-1], '5', UV_vGubn_agyh[NewRow-1]);

   if (UV_vGubn_tkgm[NewRow-1] = '1') or (UV_vGubn_tkgm[NewRow-1] = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := UV_vCod_jisa[NewRow-1];
      qryTkgum.ParamByName('num_jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryTkgum.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
//[2012.03.15]특검항목 자유롭게..
//------------------------------------------------------------------------------
         if Trim(qryTkgum.FieldByName('cod_prf').AsString) <> '' then
         begin
            qryProfileG.Active := False;

            sSelect := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';
            for iTemp := 1 to Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 do
            begin
               if iTemp = Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 then
                  sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''''
               else
                  sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''' or ';
            end;
            sOrderby := ' ) and B.cod_pf <> ''SE81'' GROUP BY B.cod_hm ';
        
            qryProfileG.SQL.Clear;
            qryProfileG.SQL.Add(sSelect + sWhere + sOrderby);
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
         end;
//------------------------------------------------------------------------------
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         sHmList_02 := sHmList_02 + vCod_chuga[i];
      end;
   end;
   //--------------------------------------------------------------------------

   if pos('H001',sHmList_02) > 0 then edtValue1.Color  := clWindow
   else                               edtValue1.Color  := clBtnFace;
   if pos('H002',sHmList_02) > 0 then edtValue2.Color  := clWindow
   else                               edtValue2.Color  := clBtnFace;
   if pos('H003',sHmList_02) > 0 then edtValue3.Color  := clWindow
   else                               edtValue3.Color  := clBtnFace;
   if pos('H004',sHmList_02) > 0 then edtValue4.Color  := clWindow
   else                               edtValue4.Color  := clBtnFace;
   if pos('H005',sHmList_02) > 0 then edtValue5.Color  := clWindow
   else                               edtValue5.Color  := clBtnFace;
   if pos('H006',sHmList_02) > 0 then edtValue6.Color  := clWindow
   else                               edtValue6.Color  := clBtnFace;
   if pos('H007',sHmList_02) > 0 then edtValue7.Color  := clWindow
   else                               edtValue7.Color  := clBtnFace;
   if pos('H008',sHmList_02) > 0 then edtValue8.Color  := clWindow
   else                               edtValue8.Color  := clBtnFace;
   if pos('H009',sHmList_02) > 0 then edtValue9.Color  := clWindow
   else                               edtValue9.Color  := clBtnFace;
   if pos('H010',sHmList_02) > 0 then edtValue10.Color := clWindow
   else                               edtValue10.Color := clBtnFace;
   if pos('H011',sHmList_02) > 0 then edtValue11.Color := clWindow
   else                               edtValue11.Color := clBtnFace;
   if pos('H012',sHmList_02) > 0 then edtValue12.Color := clWindow
   else                               edtValue12.Color := clBtnFace;
   if pos('H013',sHmList_02) > 0 then edtValue13.Color := clWindow
   else                               edtValue13.Color := clBtnFace;
   if pos('H014',sHmList_02) > 0 then edtValue14.Color := clWindow
   else                               edtValue14.Color := clBtnFace;
   if pos('H015',sHmList_02) > 0 then edtValue15.Color := clWindow
   else                               edtValue15.Color := clBtnFace;
   if pos('H016',sHmList_02) > 0 then edtValue16.Color := clWindow
   else                               edtValue16.Color := clBtnFace;
   if pos('H017',sHmList_02) > 0 then edtValue17.Color := clWindow
   else                               edtValue17.Color := clBtnFace;
   if pos('H018',sHmList_02) > 0 then edtValue18.Color := clWindow
   else                               edtValue18.Color := clBtnFace;
   if pos('H019',sHmList_02) > 0 then edtValue19.Color := clWindow
   else                               edtValue19.Color := clBtnFace;
   if pos('H020',sHmList_02) > 0 then edtValue20.Color := clWindow
   else                               edtValue20.Color := clBtnFace;
   if pos('H021',sHmList_02) > 0 then edtValue21.Color := clWindow
   else                               edtValue21.Color := clBtnFace;
   if pos('H022',sHmList_02) > 0 then edtValue22.Color := clWindow
   else                               edtValue22.Color := clBtnFace;
   if pos('H023',sHmList_02) > 0 then edtValue23.Color := clWindow
   else                               edtValue23.Color := clBtnFace;
   if pos('H024',sHmList_02) > 0 then edtValue24.Color := clWindow
   else                               edtValue24.Color := clBtnFace;
   if pos('H025',sHmList_02) > 0 then edtValue25.Color := clWindow
   else                               edtValue25.Color := clBtnFace;
   if pos('H035',sHmList_02) > 0 then edtValue26.Color := clWindow
   else                               edtValue26.Color := clBtnFace;


   if pos('H001',sHmList_02) > 0 then edtValue1.Enabled  := True
   else                               edtValue1.Enabled  := False;
   if pos('H002',sHmList_02) > 0 then edtValue2.Enabled  := True
   else                               edtValue2.Enabled  := False;
   if pos('H003',sHmList_02) > 0 then edtValue3.Enabled  := True
   else                               edtValue3.Enabled  := False;
   if pos('H004',sHmList_02) > 0 then edtValue4.Enabled  := True
   else                               edtValue4.Enabled  := False;
   if pos('H005',sHmList_02) > 0 then edtValue5.Enabled  := True
   else                               edtValue5.Enabled  := False;
   if pos('H006',sHmList_02) > 0 then edtValue6.Enabled  := True
   else                               edtValue6.Enabled  := False;
   if pos('H007',sHmList_02) > 0 then edtValue7.Enabled  := True
   else                               edtValue7.Enabled  := False;
   if pos('H008',sHmList_02) > 0 then edtValue8.Enabled  := True
   else                               edtValue8.Enabled  := False;
   if pos('H009',sHmList_02) > 0 then edtValue9.Enabled  := True
   else                               edtValue9.Enabled  := False;
   if pos('H010',sHmList_02) > 0 then edtValue10.Enabled := True
   else                               edtValue10.Enabled := False;
   if pos('H011',sHmList_02) > 0 then edtValue11.Enabled := True
   else                               edtValue11.Enabled := False;
   if pos('H012',sHmList_02) > 0 then edtValue12.Enabled := True
   else                               edtValue12.Enabled := False;
   if pos('H013',sHmList_02) > 0 then edtValue13.Enabled := True
   else                               edtValue13.Enabled := False;
   if pos('H014',sHmList_02) > 0 then edtValue14.Enabled := True
   else                               edtValue14.Enabled := False;
   if pos('H015',sHmList_02) > 0 then edtValue15.Enabled := True
   else                               edtValue15.Enabled := False;
   if pos('H016',sHmList_02) > 0 then edtValue16.Enabled := True
   else                               edtValue16.Enabled := False;
   if pos('H017',sHmList_02) > 0 then edtValue17.Enabled := True
   else                               edtValue17.Enabled := False;
   if pos('H018',sHmList_02) > 0 then edtValue18.Enabled := True
   else                               edtValue18.Enabled := False;
   if pos('H019',sHmList_02) > 0 then edtValue19.Enabled := True
   else                               edtValue19.Enabled := False;
   if pos('H020',sHmList_02) > 0 then edtValue20.Enabled := True
   else                               edtValue20.Enabled := False;
   if pos('H021',sHmList_02) > 0 then edtValue21.Enabled := True
   else                               edtValue21.Enabled := False;
   if pos('H022',sHmList_02) > 0 then edtValue22.Enabled := True
   else                               edtValue22.Enabled := False;
   if pos('H023',sHmList_02) > 0 then edtValue23.Enabled := True
   else                               edtValue23.Enabled := False;
   if pos('H024',sHmList_02) > 0 then edtValue24.Enabled := True
   else                               edtValue24.Enabled := False;
   if pos('H025',sHmList_02) > 0 then edtValue25.Enabled := True
   else                               edtValue25.Enabled := False;
   if pos('H035',sHmList_02) > 0 then edtValue26.Enabled := True
   else                               edtValue26.Enabled := False;

   edtValue1.Text  := UV_vH001[NewRow-1];
   edtValue2.Text  := UV_vH002[NewRow-1];
   edtValue3.Text  := UV_vH003[NewRow-1];
   edtValue4.Text  := UV_vH004[NewRow-1];
   edtValue5.Text  := UV_vH005[NewRow-1];
   edtValue6.Text  := UV_vH006[NewRow-1];
   edtValue7.Text  := UV_vH007[NewRow-1];
   edtValue8.Text  := UV_vH008[NewRow-1];
   edtValue9.Text  := UV_vH009[NewRow-1];
   edtValue10.Text := UV_vH010[NewRow-1];
   edtValue11.Text := UV_vH011[NewRow-1];
   edtValue12.Text := UV_vH012[NewRow-1];
   edtValue13.Text := UV_vH013[NewRow-1];
   edtValue14.Text := UV_vH014[NewRow-1];
   edtValue15.Text := UV_vH015[NewRow-1];
   edtValue16.Text := UV_vH016[NewRow-1];
   edtValue17.Text := UV_vH017[NewRow-1];
   edtValue18.Text := UV_vH018[NewRow-1];
   edtValue19.Text := UV_vH019[NewRow-1];
   edtValue20.Text := UV_vH020[NewRow-1];
   edtValue21.Text := UV_vH021[NewRow-1];
   edtValue22.Text := UV_vH022[NewRow-1];
   edtValue23.Text := UV_vH023[NewRow-1];
   edtValue24.Text := UV_vH024[NewRow-1];
   edtValue25.Text := UV_vH025[NewRow-1];
   edtValue26.Text := UV_vH035[NewRow-1];

   mskPDAT_BUNJU.Text := '';
   edtPNUM_BUNJU.Text := '';

   edtPValue1.Text  := '';
   edtPValue2.Text  := '';
   edtPValue3.Text  := '';
   edtPValue4.Text  := '';
   edtPValue5.Text  := '';
   edtPValue6.Text  := '';
   edtPValue7.Text  := '';
   edtPValue8.Text  := '';
   edtPValue9.Text  := '';
   edtPValue10.Text := '';
   edtPValue11.Text := '';
   edtPValue12.Text := '';
   edtPValue13.Text := '';
   edtPValue14.Text := '';
   edtPValue15.Text := '';
   edtPValue16.Text := '';
   edtPValue17.Text := '';
   edtPValue18.Text := '';
   edtPValue19.Text := '';
   edtPValue20.Text := '';
   edtPValue21.Text := '';
   edtPValue22.Text := '';
   edtPValue23.Text := '';
   edtPValue24.Text := '';
   edtPValue25.Text := '';
   edtPValue26.Text := '';



//[2012.03.15]종전결과 선택사항으로..
//------------------------------------------------------------------------------
   case RGrp_preGulkwa.ItemIndex of
      0 : begin
             qryinjouk.Active := False;
             qryinjouk.ParamByName('NUM_JUMIN').AsString := UV_vNum_jumin[NewRow-1];
             qryinjouk.ParamByName('DAT_GMGN').AsString  := UV_vDat_gmgn[NewRow-1];
             qryinjouk.Active := True;
             if qryinjouk.RecordCount > 0 then
             begin
                qryPrev.Active := False;
                qryPrev.Filter := '';
                qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
                qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
                qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
                qryPrev.ParamByName('GUBN_PART').AsString := '02';
                qryPrev.Active := True;
                if qryPrev.RecordCount > 0 then
                begin
                   mskPDAT_BUNJU.Text := qryPrev.FieldByName('dat_bunju').AsString;
                   edtPNUM_BUNJU.Text := qryPrev.FieldByName('num_bunju').AsString;
      
                   UV_fGulkwa := '';
                   UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
      
                   UV_vDesc_Pglkwa[index] := UV_fGulkwa;
      
                   //H001	적혈구수	02	01 	1	6
                   if Trim(Copy(UV_fGulkwa, 1,  6)) <> '' then edtPValue1.Text := Trim(Copy(UV_fGulkwa,   1,  6))
                   else                                        edtPValue1.Text := '';
                   //H002	혈색소  	02	02 	7	12
                   if Trim(Copy(UV_fGulkwa, 7,  6)) <> '' then edtPValue2.Text := Trim(Copy(UV_fGulkwa,   7,  6))
                   else                                        edtPValue2.Text := '';
                   //H003	Hematocrit	02	03 	13	18
                   if Trim(Copy(UV_fGulkwa, 13,  6)) <> '' then edtPValue3.Text := Trim(Copy(UV_fGulkwa, 13,  6))
                   else                                         edtPValue3.Text := '';
                   //H004	MCV	        02	04 	19	24
                   if Trim(Copy(UV_fGulkwa, 19,  6)) <> '' then edtPValue4.Text := Trim(Copy(UV_fGulkwa, 19,  6))
                   else                                         edtPValue4.Text := '';
                   //H005	MCH	        02	05 	25	30
                   if Trim(Copy(UV_fGulkwa, 25,  6)) <> '' then edtPValue5.Text := Trim(Copy(UV_fGulkwa, 25,  6))
                   else                                         edtPValue5.Text := '';
                   //H006	MCHC	        02	06 	31	36
                   if Trim(Copy(UV_fGulkwa, 31,  6)) <> '' then edtPValue6.Text := Trim(Copy(UV_fGulkwa, 31,  6))
                   else                                         edtPValue6.Text := '';
                   //H007	RDW	        02	07 	37	42
                   if Trim(Copy(UV_fGulkwa, 37,  6)) <> '' then edtPValue7.Text := Trim(Copy(UV_fGulkwa, 37,  6))
                   else                                         edtPValue7.Text := '';
                   //H008	혈소판수	02	08 	43	48
                   if Trim(Copy(UV_fGulkwa, 43,  6)) <> '' then edtPValue8.Text := Trim(Copy(UV_fGulkwa, 43,  6))
                   else                                         edtPValue8.Text := '';
                   //H009	MPV	        02	09 	49	54
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then edtPValue9.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else                                         edtPValue9.Text := '';
                   //H010	PDW	        02	10 	55	60
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else                                         edtPValue10.Text := '';
                   //H011	백혈구수	02	11 	61	66
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else                                         edtPValue11.Text := '';
                   //H012	분획호중구	02	12 	67	72
                   if Trim(Copy(UV_fGulkwa, 67,  6)) <> '' then edtPValue12.Text := Trim(Copy(UV_fGulkwa, 67,  6))
                   else                                         edtPValue12.Text := '';
                   //H013	봉상호중구	02	13 	73	78
                   if Trim(Copy(UV_fGulkwa, 73,  6)) <> '' then edtPValue13.Text := Trim(Copy(UV_fGulkwa, 73,  6))
                   else                                         edtPValue13.Text := '';
                   //H014	임파구	        02	14 	79	84
                   if Trim(Copy(UV_fGulkwa, 79,  6)) <> '' then edtPValue14.Text := Trim(Copy(UV_fGulkwa, 79,  6))
                   else                                         edtPValue14.Text := '';
                   //H015	단핵구	        02	15 	85	90
                   if Trim(Copy(UV_fGulkwa, 85,  6)) <> '' then edtPValue15.Text := Trim(Copy(UV_fGulkwa, 85,  6))
                   else                                         edtPValue15.Text := '';
                   //H016	호산구	        02	16 	91	96
                   if Trim(Copy(UV_fGulkwa, 91,  6)) <> '' then edtPValue16.Text := Trim(Copy(UV_fGulkwa, 91,  6))
                   else                                         edtPValue16.Text := '';
                   //H017	호염기성구	02	17 	97	102
                   if Trim(Copy(UV_fGulkwa, 97,  6)) <> '' then edtPValue17.Text := Trim(Copy(UV_fGulkwa, 97,  6))
                   else                                         edtPValue17.Text := '';
                   //H018	후골수구	2	18	103	108
                   if Trim(Copy(UV_fGulkwa, 103,  6)) <> '' then edtPValue18.Text := Trim(Copy(UV_fGulkwa, 103,  6))
                   else                                          edtPValue18.Text := '';
                   //H019	골수구	        2	19	109	114
                   if Trim(Copy(UV_fGulkwa, 109,  6)) <> '' then edtPValue19.Text := Trim(Copy(UV_fGulkwa, 109,  6))
                   else                                          edtPValue19.Text := '';
                   //H020	전골수세포	2	20	115	120
                   if Trim(Copy(UV_fGulkwa, 115,  6)) <> '' then edtPValue20.Text := Trim(Copy(UV_fGulkwa, 115,  6))
                   else                                          edtPValue20.Text := '';
                   //H021	골수아세포	2	21	121	126
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then edtPValue21.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else                                          edtPValue21.Text := '';
                   //H022	유핵적혈구	2	22	127	132
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then edtPValue22.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else                                          edtPValue22.Text := '';
                   //H023	ABO형	        2	23	133	138
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then edtPValue23.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else                                          edtPValue23.Text := '';
                   //H024	Rh형	        2	24	139	144
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then edtPValue24.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else                                          edtPValue24.Text := '';
                   //H025	ESR	        2	25	145	150
                   if Trim(Copy(UV_fGulkwa, 145,  6)) <> '' then edtPValue25.Text := Trim(Copy(UV_fGulkwa, 145,  6))
                   else                                          edtPValue25.Text := '';
                   //H035	PCT	        2	31	175	180
                   if Trim(Copy(UV_fGulkwa, 175,  6)) <> '' then edtPValue26.Text := Trim(Copy(UV_fGulkwa, 175,  6))
                   else                                          edtPValue26.Text := '';
                end;
                qryPrev.Active := False;
             end;
             qryinjouk.Active := False;
          end;
      1 : begin
             qryinjouk1.Active := False;
      
             sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, B.cod_bjjs, B.dat_bunju, B.num_bunju, B.gubn_part ' +
                        ' FROM gumgin G INNER JOIN gulkwa b ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';
             sWhere  := ' WHERE G.num_jumin = ''' + UV_vNum_jumin[NewRow-1] + '''';
             sWhere  := sWhere + ' AND G.dat_gmgn < ''' + UV_vDat_gmgn[NewRow-1] + '''';
             sWhere  := sWhere + ' AND B.gubn_part = ''02''';
             sOrderBy := ' order by G.dat_gmgn DESC ';

             qryinjouk1.SQL.Clear;
             qryinjouk1.SQL.Add(sSelect + sWhere + sOrderBy);
             qryinjouk1.Active := True;
             if qryinjouk1.RecordCount > 0 then
             begin
                qryPrev.Active := False;
                qryPrev.Filter := '';
                qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk1.FieldByName('cod_bjjs').AsString;
                qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk1.FieldByName('num_id').AsString;
                qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk1.FieldByName('dat_gmgn').AsString;
                qryPrev.ParamByName('GUBN_PART').AsString := '02';
                qryPrev.Active := True;
                if qryPrev.RecordCount > 0 then
                begin
                   mskPDAT_BUNJU.Text := qryPrev.FieldByName('dat_bunju').AsString;
                   edtPNUM_BUNJU.Text := qryPrev.FieldByName('num_bunju').AsString;
      
                   UV_fGulkwa := '';
                   UV_fGulkwa1 := qryPrev.FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := qryPrev.FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := qryPrev.FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
      
                   UV_vDesc_Pglkwa[index] := UV_fGulkwa;
      
                   //H001	적혈구수	02	01 	1	6
                   if Trim(Copy(UV_fGulkwa, 1,  6)) <> '' then edtPValue1.Text := Trim(Copy(UV_fGulkwa,   1,  6))
                   else                                        edtPValue1.Text := '';
                   //H002	혈색소  	02	02 	7	12
                   if Trim(Copy(UV_fGulkwa, 7,  6)) <> '' then edtPValue2.Text := Trim(Copy(UV_fGulkwa,   7,  6))
                   else                                        edtPValue2.Text := '';
                   //H003	Hematocrit	02	03 	13	18
                   if Trim(Copy(UV_fGulkwa, 13,  6)) <> '' then edtPValue3.Text := Trim(Copy(UV_fGulkwa, 13,  6))
                   else                                         edtPValue3.Text := '';
                   //H004	MCV	        02	04 	19	24
                   if Trim(Copy(UV_fGulkwa, 19,  6)) <> '' then edtPValue4.Text := Trim(Copy(UV_fGulkwa, 19,  6))
                   else                                         edtPValue4.Text := '';
                   //H005	MCH	        02	05 	25	30
                   if Trim(Copy(UV_fGulkwa, 25,  6)) <> '' then edtPValue5.Text := Trim(Copy(UV_fGulkwa, 25,  6))
                   else                                         edtPValue5.Text := '';
                   //H006	MCHC	        02	06 	31	36
                   if Trim(Copy(UV_fGulkwa, 31,  6)) <> '' then edtPValue6.Text := Trim(Copy(UV_fGulkwa, 31,  6))
                   else                                         edtPValue6.Text := '';
                   //H007	RDW	        02	07 	37	42
                   if Trim(Copy(UV_fGulkwa, 37,  6)) <> '' then edtPValue7.Text := Trim(Copy(UV_fGulkwa, 37,  6))
                   else                                         edtPValue7.Text := '';
                   //H008	혈소판수	02	08 	43	48
                   if Trim(Copy(UV_fGulkwa, 43,  6)) <> '' then edtPValue8.Text := Trim(Copy(UV_fGulkwa, 43,  6))
                   else                                         edtPValue8.Text := '';
                   //H009	MPV	        02	09 	49	54
                   if Trim(Copy(UV_fGulkwa, 49,  6)) <> '' then edtPValue9.Text := Trim(Copy(UV_fGulkwa, 49,  6))
                   else                                         edtPValue9.Text := '';
                   //H010	PDW	        02	10 	55	60
                   if Trim(Copy(UV_fGulkwa, 55,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 55,  6))
                   else                                         edtPValue10.Text := '';
                   //H011	백혈구수	02	11 	61	66
                   if Trim(Copy(UV_fGulkwa, 61,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 61,  6))
                   else                                         edtPValue11.Text := '';
                   //H012	분획호중구	02	12 	67	72
                   if Trim(Copy(UV_fGulkwa, 67,  6)) <> '' then edtPValue12.Text := Trim(Copy(UV_fGulkwa, 67,  6))
                   else                                         edtPValue12.Text := '';
                   //H013	봉상호중구	02	13 	73	78
                   if Trim(Copy(UV_fGulkwa, 73,  6)) <> '' then edtPValue13.Text := Trim(Copy(UV_fGulkwa, 73,  6))
                   else                                         edtPValue13.Text := '';
                   //H014	임파구	        02	14 	79	84
                   if Trim(Copy(UV_fGulkwa, 79,  6)) <> '' then edtPValue14.Text := Trim(Copy(UV_fGulkwa, 79,  6))
                   else                                         edtPValue14.Text := '';
                   //H015	단핵구	        02	15 	85	90
                   if Trim(Copy(UV_fGulkwa, 85,  6)) <> '' then edtPValue15.Text := Trim(Copy(UV_fGulkwa, 85,  6))
                   else                                         edtPValue15.Text := '';
                   //H016	호산구	        02	16 	91	96
                   if Trim(Copy(UV_fGulkwa, 91,  6)) <> '' then edtPValue16.Text := Trim(Copy(UV_fGulkwa, 91,  6))
                   else                                         edtPValue16.Text := '';
                   //H017	호염기성구	02	17 	97	102
                   if Trim(Copy(UV_fGulkwa, 97,  6)) <> '' then edtPValue17.Text := Trim(Copy(UV_fGulkwa, 97,  6))
                   else                                         edtPValue17.Text := '';
                   //H018	후골수구	2	18	103	108
                   if Trim(Copy(UV_fGulkwa, 103,  6)) <> '' then edtPValue18.Text := Trim(Copy(UV_fGulkwa, 103,  6))
                   else                                          edtPValue18.Text := '';
                   //H019	골수구	        2	19	109	114
                   if Trim(Copy(UV_fGulkwa, 109,  6)) <> '' then edtPValue19.Text := Trim(Copy(UV_fGulkwa, 109,  6))
                   else                                          edtPValue19.Text := '';
                   //H020	전골수세포	2	20	115	120
                   if Trim(Copy(UV_fGulkwa, 115,  6)) <> '' then edtPValue20.Text := Trim(Copy(UV_fGulkwa, 115,  6))
                   else                                          edtPValue20.Text := '';
                   //H021	골수아세포	2	21	121	126
                   if Trim(Copy(UV_fGulkwa, 121,  6)) <> '' then edtPValue21.Text := Trim(Copy(UV_fGulkwa, 121,  6))
                   else                                          edtPValue21.Text := '';
                   //H022	유핵적혈구	2	22	127	132
                   if Trim(Copy(UV_fGulkwa, 127,  6)) <> '' then edtPValue22.Text := Trim(Copy(UV_fGulkwa, 127,  6))
                   else                                          edtPValue22.Text := '';
                   //H023	ABO형	        2	23	133	138
                   if Trim(Copy(UV_fGulkwa, 133,  6)) <> '' then edtPValue23.Text := Trim(Copy(UV_fGulkwa, 133,  6))
                   else                                          edtPValue23.Text := '';
                   //H024	Rh형	        2	24	139	144
                   if Trim(Copy(UV_fGulkwa, 139,  6)) <> '' then edtPValue24.Text := Trim(Copy(UV_fGulkwa, 139,  6))
                   else                                          edtPValue24.Text := '';
                   //H025	ESR	        2	25	145	150
                   if Trim(Copy(UV_fGulkwa, 145,  6)) <> '' then edtPValue25.Text := Trim(Copy(UV_fGulkwa, 145,  6))
                   else                                          edtPValue25.Text := '';
                   //H035	PCT	        2	31	175	180
                   if Trim(Copy(UV_fGulkwa, 175,  6)) <> '' then edtPValue26.Text := Trim(Copy(UV_fGulkwa, 175,  6))
                   else                                          edtPValue26.Text := '';
                end;
                qryPrev.Active := False;
             end;
             qryinjouk.Active := False;
          end;
   end;

//[2012.003.15]검체불량관리 추가..
//------------------------------------------------------------------------------
   ckGUBN_JUNGDO1.Checked := False;
   ckGUBN_JUNGDO2.Checked := False;
   ckGUBN_JUNGDO3.Checked := False;
   ckGUBN_JUNGDO4.Checked := False;
   ckGUBN_JUNGDO5.Checked := False;
   ckGUBN_JUNGDO6.Checked := False;
   ckGUBN_JUNGDO8.Checked := False;
   ckGUBN_JUNGDO9.Checked := False;
   mskDAT_SOLVE1.Text := '';
   mskDAT_SOLVE2.Text := '';
   mskDAT_SOLVE3.Text := '';
   mskDAT_SOLVE4.Text := '';
   mskDAT_SOLVE5.Text := '';
   mskDAT_SOLVE6.Text := '';
   mskDAT_SOLVE8.Text := '';
   mskDAT_SOLVE9.Text := '';

   qryGum_bul.Active := False;
   qryGum_bul.ParamByName('cod_jisa').AsString := UV_vCod_jisa[NewRow-1];
   qryGum_bul.ParamByName('num_id').AsString   := UV_vNum_id[NewRow-1];
   qryGum_bul.ParamByName('dat_gmgn').AsString := UV_vDat_gmgn[NewRow-1];
   qryGum_bul.Active := True;
   if qryGum_bul.RecordCount > 0 then
   begin
      while not qryGum_bul.Eof  do
      begin
         if qryGum_bul.FieldByName('gubn_bul').AsString = '01' then
         begin
            ckGUBN_JUNGDO1.Checked := True;
            mskDAT_SOLVE1.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '02' then
         begin
            ckGUBN_JUNGDO2.Checked := True;
            mskDAT_SOLVE2.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '03' then
         begin
            ckGUBN_JUNGDO3.Checked := True;
            mskDAT_SOLVE3.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '04' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '05' then
         begin
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '06' then
         begin
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '07' then
         begin
            ckGUBN_JUNGDO4.Checked := True;
            mskDAT_SOLVE4.Text := qryGum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO5.Checked := True;
            mskDAT_SOLVE5.Text := qryGum_bul.FieldByName('dat_solve').AsString;
            ckGUBN_JUNGDO6.Checked := True;
            mskDAT_SOLVE6.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '08' then
         begin
            ckGUBN_JUNGDO8.Checked := True;
            mskDAT_SOLVE8.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;
         if qryGum_bul.FieldByName('gubn_bul').AsString = '09' then
         begin
            ckGUBN_JUNGDO9.Checked := True;
            mskDAT_SOLVE9.Text := qryGum_bul.FieldByName('dat_solve').AsString;
         end;

         qryGum_bul.Next;
      end;
   end;
//------------------------------------------------------------------------------

   // 값변형 체크
   UP_DChange(edtValue1);
   UP_DChange(edtValue2);
   UP_DChange(edtValue3);
   UP_DChange(edtValue4);
   UP_DChange(edtValue5);
   UP_DChange(edtValue6);
   UP_DChange(edtValue7);
   UP_DChange(edtValue8);
   UP_DChange(edtValue9);
   UP_DChange(edtValue10);
   UP_DChange(edtValue11);
   UP_DChange(edtValue12);
   UP_DChange(edtValue13);
   UP_DChange(edtValue14);
   UP_DChange(edtValue15);
   UP_DChange(edtValue16);
   UP_DChange(edtValue17);
   UP_DChange(edtValue18);
   UP_DChange(edtValue19);
   UP_DChange(edtValue20);
   UP_DChange(edtValue21);
   UP_DChange(edtValue22);
   UP_DChange(edtValue23);
   UP_DChange(edtValue24);
   UP_DChange(edtValue25);
   UP_DChange(edtValue26);

   if (UV_vABO_chk[NewRow-1] = 'Y') and (edtValue23.Enabled = True) then pnlNum23.Color := clWhite
   else                                                                  pnlNum23.Color := $00FFFF80;

   if (UV_vABO_chk[NewRow-1] = 'Y') and (edtValue24.Enabled = True) then pnlNum24.Color := clWhite
   else                                                                  pnlNum24.Color := $00FFFF80;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD1I13.btnPInsertClick(Sender: TObject);
begin
   inherited;

   // 작업절차
   // 1. 사용자가 작업중이라면 저장이 입력상태
   // 2. 사용자가 작업중이아니라면 바로 입력상태
   if UV_bEdit then
   begin
      if UF_Save then
         btnInsert.Click;
   end
   else
      btnInsert.Click;
end;

procedure TfrmLD1I13.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;

   // Button Click시 Event Procedure를 공유한 후 구분을 Sender로 수행
   if Sender = btnJumin then
   begin
      if GF_InjoukClick(Self, sCode, sName) then
      begin
         mskJumin.Text := sCode;
         edtName.Text  := sName;
      end;   
   end
   else if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD1I13.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I13.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD1I13.UP_Exit(Sender: TObject);
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

procedure TfrmLD1I13.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      mskDate.Color       := $00E6FFFE;
      mskDate.Enabled     := True;
      btnDate.Enabled     := True;
      mskFrom.Enabled     := True;
      mskTo.Enabled       := True;
      mskJumin.Color      := clWindow;
      mskJumin.Enabled    := False;
      btnJumin.Enabled    := False;
      mskDate.SetFocus;

      // 찾기 Enable
     // edtFind.Enabled  := True;
      edtFind2.Enabled   := True;

      mskDate.SetFocus;
   end
   else if Sender = rbJumin then
   begin
      mskDate.Color       := clWindow;
      mskDate.Enabled     := False;
      btnDate.Enabled     := False;
      mskFrom.Enabled     := False;
      mskTo.Enabled       := False;
      mskJumin.Color      := $00E6FFFE;
      mskJumin.Enabled    := True;
      btnJumin.Enabled    := True;
      mskJumin.SetFocus;

      // 찾기 Disable
    //  edtFind.Enabled  := False;
      edtFind2.Enabled  := False;
   end;
end;

procedure TfrmLD1I13.edtFindExit(Sender: TObject);
begin
   inherited;

   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0);
end;

procedure TfrmLD1I13.mskNumExit(Sender: TObject);
var i, iTemp : Integer;
    sName : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if mskNum.Text = '' then exit;

   // 위치를 추출
   iTemp := StrToInt(mskNum.Text);
   if iTemp > 45 then iTemp := iTemp - 45;

   // 해당위치로 Focus이동
   for i := 0 to pnlDetail.ControlCount - 1 do
   begin
      if pnlDetail.Controls[i].ClassType = TEdit then
      begin
         sName := TEdit(pnlDetail.Controls[i]).Name;
         if ('edtValue' + IntToStr(iTemp)) = sName then
         begin
            if TEdit(pnlDetail.Controls[i]).Enabled then
               TEdit(pnlDetail.Controls[i]).SetFocus
            else
               mskNum.SetFocus;

            break;
         end;
      end;
   end;
end;

procedure TfrmLD1I13.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key = VK_F7 then
   begin
      if edtFind2.Enabled then edtFind2.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);

   if Key = VK_F3 then
   begin
      if edtFind.Enabled then edtFind.SetFocus;
   end
   else if Key = VK_F4 then mskNum.SetFocus
   else if Key = VK_F5 then UP_MoveNum(btnPrevNum)
   else if Key = VK_F6 then UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I13.UP_EditDisplay(Sender : TObject);
begin
   if Sender.ClassType <> TEdit then exit;

   // 자료가 선택되었는지 Check
   if TEdit(Sender).Text = '~' then TEdit(Sender).Text := '음성'
   else if TEdit(Sender).Text = '!' then TEdit(Sender).Text := '양성'
   else if TEdit(Sender).Text = '@' then TEdit(Sender).Text := '약양성'
   else if TEdit(Sender).Text = '#' then TEdit(Sender).Text := '경계역'
   else if TEdit(Sender).Text = '$' then TEdit(Sender).Text := '보류'
   else if TEdit(Sender).Text = '%' then TEdit(Sender).Text := 'Rh+'
   else if TEdit(Sender).Text = '^' then TEdit(Sender).Text := 'Rh-'
   else if TEdit(Sender).Text = '&' then TEdit(Sender).Text := 'A'
   else if TEdit(Sender).Text = '*' then TEdit(Sender).Text := 'B'
   else if TEdit(Sender).Text = '(' then TEdit(Sender).Text := 'O'
   else if TEdit(Sender).Text = ')' then TEdit(Sender).Text := 'AB';

   UP_Change(Sender);
end;

procedure TfrmLD1I13.edtValue1Change(Sender: TObject);
begin
  inherited;
   //해당하는 값을 넣어줌.
   if Copy(TEdit(Sender).Name, 1, 8)  = 'edtValue' then
      UP_EditDisplay(Sender);
end;

procedure TfrmLD1I13.Cmb_gubnChange(Sender: TObject);
begin
  inherited;
  if Cmb_gubn.ItemIndex = 0 then
  begin
     Label2.Caption := '분주번호(F3)';
     edtFind.MaxLength := 4;
     mskFrom.Enabled := True;
     mskTo.Enabled := True;
     MEdt_SampS.Enabled := False;
     MEdt_SampE.Enabled := False;
  end
  else if Cmb_gubn.ItemIndex = 1 then
  begin
     Label2.Caption := '샘플번호(F3)';
     edtFind.MaxLength := 6;
     mskFrom.Enabled := False;
     mskTo.Enabled := False;
     MEdt_SampS.Enabled := True;
     MEdt_SampE.Enabled := True;
  end
  else if Cmb_gubn.ItemIndex = 2 then
  begin
     Label2.Caption := '샘플번호(F3)';
     edtFind.MaxLength := 6;
     mskFrom.Enabled := False;
     mskTo.Enabled := False;
     MEdt_SampS.Enabled := True;
     MEdt_SampE.Enabled := True;
  end
  else if Cmb_gubn.ItemIndex = 3 then
  begin
     Label2.Caption := '샘플번호(F3)';
     edtFind.MaxLength := 6;
     mskFrom.Enabled := False;
     mskTo.Enabled := False;
     MEdt_SampS.Enabled := True;
     MEdt_SampE.Enabled := True;
  end;
end;

procedure TfrmLD1I13.cmbRelationChange(Sender: TObject);
begin
   inherited;
   frmLD1I016 := TFrmLD1I016.Create(self);
   frmLD1I016.ShowModal;
end;

procedure TfrmLD1I13.SBut_ExcelClick(Sender: TObject);       //2014.11.13 곽윤설
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
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, 18], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows do
      begin
         if Row = 0 then
         begin
             ArrV3[Row,  0] := '검체번호';
             ArrV3[Row,  1] := '성명';
             ArrV3[Row,  2] := 'RBC';
             ArrV3[Row,  3] := 'HGB';
             ArrV3[Row,  4] := 'HCT';
             ArrV3[Row,  5] := 'MCV';
             ArrV3[Row,  6] := 'MCH';
             ArrV3[Row,  7] := 'MCHC';
             ArrV3[Row,  8] := 'RDW-CV';
             ArrV3[Row,  9] := 'PLT';
             ArrV3[Row, 10] := 'MPV';
             ArrV3[Row, 11] := 'PDW';
             ArrV3[Row, 12] := 'WBC';
             ArrV3[Row, 13] := 'NEUT%';
             ArrV3[Row, 14] := 'LYMPH%';
             ArrV3[Row, 15] := 'MONO%';
             ArrV3[Row, 16] := 'EO%';
             ArrV3[Row, 17] := 'BASO%';
         end
         else
         begin
            // 자료 Display               //old //new
            grdMasterRowChanged(grdMaster, Row, Row);
            Application.ProcessMessages;
            ArrV3[Row,  0] := copy(mskDAT_GMGN.Text,3,8) + Edt_smp.Text;
            ArrV3[Row,  1] := edtDESC_NAME.Text;
            ArrV3[Row,  2] := edtValue1.Text;
            ArrV3[Row,  3] := edtValue2.Text;
            ArrV3[Row,  4] := edtValue3.Text;
            ArrV3[Row,  5] := edtValue4.Text;
            ArrV3[Row,  6] := edtValue5.Text;
            ArrV3[Row,  7] := edtValue6.Text;
            ArrV3[Row,  8] := edtValue7.Text;
            ArrV3[Row,  9] := edtValue8.Text;
            ArrV3[Row, 10] := edtValue9.Text;
            ArrV3[Row, 11] := edtValue10.Text;
            ArrV3[Row, 12] := edtValue11.Text;
            ArrV3[Row, 13] := edtValue12.Text;
            ArrV3[Row, 14] := edtValue14.Text;
            ArrV3[Row, 15] := edtValue15.Text;
            ArrV3[Row, 16] := edtValue16.Text;
            ArrV3[Row, 17] := edtValue17.Text;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, 18]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

procedure TfrmLD1I13.EdtFind2Exit(Sender: TObject);
begin
  inherited;
  // 자료가 존재하는지 Check
  if Length(edtFind2.Text) < edtFind2.MaxLength then exit;

  // 찾기수행
  GF_Find(grdMaster, edtFind2.Text, 2, 1, 1, 0);
end;

end.
