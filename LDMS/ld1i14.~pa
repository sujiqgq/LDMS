//==============================================================================
// 프로그램 설명 : [센터]자체검진 Urin결과등록
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
unit LD1I14;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit;

type
  TfrmLD1I14 = class(TfrmSingle)
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
    qryPart03: TQuery;
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
    pnlNum12: TPanel;
    pnlName12: TPanel;
    edtValue12: TEdit;
    pnlNum10: TPanel;
    pnlName10: TPanel;
    edtValue10: TEdit;
    pnlNum11: TPanel;
    pnlName11: TPanel;
    edtValue11: TEdit;
    edtPValue10: TEdit;
    edtPValue11: TEdit;
    edtPValue12: TEdit;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    qryProfileG: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    Label13: TLabel;
    Edt_bjjs: TEdit;
    Panel2: TPanel;
    Panel3: TPanel;
    edtPValue13: TEdit;
    edtValue13: TEdit;
    Label15: TLabel;
    Label14: TLabel;
    edtPValue2: TEdit;
    MEdt_Barcode: TMaskEdit;
    rbBarcode: TRadioButton;
    Cmb_gubn: TComboBox;
    qryProfile: TQuery;
    Label16: TLabel;
    mskPDAT_BUNJU: TDateEdit;
    Label17: TLabel;
    edtPNUM_BUNJU: TEdit;
    Label21: TLabel;
    edt_sinjang: TEdit;
    Label22: TLabel;
    edt_chejung: TEdit;
    Label23: TLabel;
    edt_gumgin_Check: TEdit;
    qryGumgin_Check: TQuery;
    qrykicho: TQuery;
    Panel4: TPanel;
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
    qryGum_bul: TQuery;
    PnlNum14: TPanel;
    PnlName14: TPanel;
    edtValue14: TEdit;
    edtPValue14: TEdit;
    Label19: TLabel;
    EdtFind2: TEdit;
    mskNum: TMaskEdit;
    rbDate: TRadioButton;
    RG_sort: TRadioGroup;
    mskDate: TDateEdit;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
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
    procedure edtFind2Exit(Sender: TObject);
    procedure mskNumExit(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure UP_MoveNum(Sender: TObject);
    procedure edtValue1Change(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCod_jisa, UV_vNum_id, UV_vNum_jumin, UV_vDesc_name, UV_vDat_gmgn, UV_vNum_samp,
    UV_vCod_hul,    UV_vCod_cancer, UV_vCod_jangbi, UV_vCod_chuga,
    UV_vGubn_nosin, UV_vGubn_nsyh, UV_vGubn_adult, UV_vGubn_adyh,
    UV_vGubn_life,  UV_vGubn_lfyh, UV_vGubn_bogun, UV_vGubn_bgyh,
    UV_vGubn_agab,  UV_vGubn_agyh, UV_vGubn_tkgm,
    UV_vCod_bjjs, UV_vDat_bunju, UV_vNum_bunju, UV_vDesc_glkwa, UV_vDesc_Pglkwa, UV_vPart03,UV_vGumgin_check,
    UV_vU001, UV_vU002, UV_vU003, UV_vU004, UV_vU005, UV_vU006,
    UV_vU007, UV_vU008, UV_vU009, UV_vU010, UV_vU011, UV_vU012,
    UV_vU013, UV_vY004 : Variant;

    iCntCheck : Integer;

    // 항목코드, 시작위치, 마지막위치, 값
    UV_sValue : array[1..14] of String;
    UV_sGubn  : array[1..14] of String;
    UV_sUnit  : array[1..14] of String;

    // part구분
    UV_sGubn_part, sHangmok : String;

    // SQL문
    UV_sBasicSQL : String;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_GridSet;
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_EditDisplay(Sender : TObject);

    function  UF_Condition : Boolean;
    function  UF_Save : Boolean;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    Procedure Hangmok_Set_U011;
    procedure UP_DChange(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmLD1I14: TfrmLD1I14;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I017;

{$R *.DFM}

Procedure TfrmLD1I14.Hangmok_Set_U011;

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

function  TfrmLD1I14.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I14.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD1I14.UP_GridSet;
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

procedure TfrmLD1I14.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCod_jisa    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_bjjs    := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_glkwa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_Pglkwa := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp    := VarArrayCreate([0, 0], varOleStr);
      UV_vPart03      := VarArrayCreate([0, 0], varOleStr);

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

      UV_vU001        := VarArrayCreate([0, 0], varOleStr);
      UV_vU002        := VarArrayCreate([0, 0], varOleStr);
      UV_vU003        := VarArrayCreate([0, 0], varOleStr);
      UV_vU004        := VarArrayCreate([0, 0], varOleStr);
      UV_vU005        := VarArrayCreate([0, 0], varOleStr);
      UV_vU006        := VarArrayCreate([0, 0], varOleStr);
      UV_vU007        := VarArrayCreate([0, 0], varOleStr);
      UV_vU008        := VarArrayCreate([0, 0], varOleStr);
      UV_vU009        := VarArrayCreate([0, 0], varOleStr);
      UV_vU010        := VarArrayCreate([0, 0], varOleStr);
      UV_vU011        := VarArrayCreate([0, 0], varOleStr);
      UV_vU012        := VarArrayCreate([0, 0], varOleStr);
      UV_vU013        := VarArrayCreate([0, 0], varOleStr);
      UV_vY004        := VarArrayCreate([0, 0], varOleStr);
      UV_vGumgin_check := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCod_jisa,    iLength);
      VarArrayRedim(UV_vNum_id,      iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vCod_bjjs,    iLength);
      VarArrayRedim(UV_vDat_bunju,   iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_glkwa,  iLength);
      VarArrayRedim(UV_vDesc_Pglkwa, iLength);
      VarArrayRedim(UV_vNum_samp,    iLength);
      VarArrayRedim(UV_vPart03,      iLength);

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

      VarArrayRedim(UV_vU001,        iLength);
      VarArrayRedim(UV_vU002,        iLength);
      VarArrayRedim(UV_vU003,        iLength);
      VarArrayRedim(UV_vU004,        iLength);
      VarArrayRedim(UV_vU005,        iLength);
      VarArrayRedim(UV_vU006,        iLength);
      VarArrayRedim(UV_vU007,        iLength);
      VarArrayRedim(UV_vU008,        iLength);
      VarArrayRedim(UV_vU009,        iLength);
      VarArrayRedim(UV_vU010,        iLength);
      VarArrayRedim(UV_vU011,        iLength);
      VarArrayRedim(UV_vU012,        iLength);
      VarArrayRedim(UV_vU013,        iLength);
      VarArrayRedim(UV_vY004,        iLength);
      VarArrayRedim(UV_vGumgin_check, iLength);
   end;
end;

function TfrmLD1I14.UF_Condition : Boolean;
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

function TfrmLD1I14.UF_Save : Boolean;
var i, iTemp : Integer;
    Part03 : String;
begin
   Result := False;

   // 작업중인지 Check
   if not UV_bEdit then exit;

   // 1. Not Null Check
   if not GF_NotNullCheck(pnlDetail) then exit;

   // Save Confirm message
   if not GF_MsgBox('Confirmation', 'SAVE') then exit;

   // 현재위치 추출
   i := grdMaster.CurrentDataRow - 1;

   with qryPart03 do
   begin
      Close;
      qryPart03.ParamByName('dat_bunju').AsString := mskDAT_BUNJU.Text;
      qryPart03.ParamByName('num_bunju').AsString := edtNUM_BUNJU.Text;
      Open;

      if qryPart03.RecordCount > 0 then
      begin
         Part03 := '';
         if FieldByName('desc_glkwa1').AsString <> '' then
         begin
            for iTemp := 1 to 200 - Length(FieldByName('desc_glkwa1').AsString) do
                Part03 := Part03 + ' ';
            Part03 := FieldByName('desc_glkwa1').AsString + Part03;
         end
         else
         begin
            for iTemp := 1 to 2 do
                Part03 := Part03 + '                                                                                                    ';
         end;
      end;
   end;

   // U001	백혈구(뇨검사)	        03	02 	7	12
   Part03 := Copy(Part03, 1,  6)  + GF_RPad(edtValue1.text,  6, ' ') + Copy(Part03, 13, Length(Part03));
   // U002	Nitrite(뇨검사)	        03	03 	13	18
   Part03 := Copy(Part03, 1, 12)  + GF_RPad(edtValue2.text,  6, ' ') + Copy(Part03, 19, Length(Part03));
   // U003	PH (뇨검사)	        03	04 	19	24
   Part03 := Copy(Part03, 1, 18)  + GF_RPad(edtValue3.text,  6, ' ') + Copy(Part03, 25, Length(Part03));
   // U004	단백(뇨검사)	        03	05 	25    	30
   Part03 := Copy(Part03, 1, 24)  + GF_RPad(edtValue4.text,  6, ' ') + Copy(Part03, 31, Length(Part03));
   // U005	당(뇨검사)	        03	06 	31   	36
   Part03 := Copy(Part03, 1, 30)  + GF_RPad(edtValue5.text,  6, ' ') + Copy(Part03, 37, Length(Part03));
   // U006	케톤(뇨검사)	        03	07 	37   	42
   Part03 := Copy(Part03, 1, 36)  + GF_RPad(edtValue6.text,  6, ' ') + Copy(Part03, 43, Length(Part03));
   // U007	유로빌리노겐(뇨)	03	08 	43  	48
   Part03 := Copy(Part03, 1, 42)  + GF_RPad(edtValue7.text,  6, ' ') + Copy(Part03, 49, Length(Part03));
   // U008	빌리루빈(뇨검사)	03	09 	49  	54
   Part03 := Copy(Part03, 1, 48)  + GF_RPad(edtValue8.text,  6, ' ') + Copy(Part03, 55, Length(Part03));
   // U009	적혈구(뇨검사)	        03	10 	55  	60
   Part03 := Copy(Part03, 1, 54)  + GF_RPad(edtValue9.text,  6, ' ') + Copy(Part03, 61, Length(Part03));
   // U010	요비중	                03	11 	61	66
   Part03 := Copy(Part03, 1, 60)  + GF_RPad(edtValue10.text, 6, ' ') + Copy(Part03, 67, Length(Part03));
   // U011	뇨침사 WBC             	03	12 	67	72
   Part03 := Copy(Part03, 1, 66)  + GF_RPad(edtValue11.text, 6, ' ') + Copy(Part03, 73, Length(Part03));
   // U012	뇨침사 RBC            	03	13 	73	78
   Part03 := Copy(Part03, 1, 72)  + GF_RPad(edtValue12.text, 6, ' ') + Copy(Part03, 79, Length(Part03));
   // U013	상피세포            	03	14 	79	84
   Part03 := Copy(Part03, 1, 78)  + GF_RPad(edtValue13.text, 6, ' ') + Copy(Part03, 85, Length(Part03));
   // Y004	분별잠혈            	03	16      85      90
   Part03 := Copy(Part03, 1, 84)  + GF_RPad(edtValue14.text, 6, ' ') + Copy(Part03, 91, Length(Part03));


   Part03:= 'A' + Part03;
   Part03:= Trim(Part03);
   Part03:= Copy(Part03, 2, Length(Part03)-1);

   // DB 작업
   dmComm.database.StartTransaction;
   try
      with qryU_Glkwa do
      begin
         ParamByName('COD_jisa').AsString    := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString      := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString   := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString   := edtNUM_BUNJU.Text;
         ParamByName('GUBN_PART').AsString   := '03';
         ParamByName('desc_glkwa1').AsString := Part03;
         ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         ParamByName('COD_UPDATE').AsString  := GV_sUserId;
         Execsql;

         GP_query_log(qryU_Glkwa, 'LD1I14', 'U', 'Y');
      end;

      with qryU_Bunju do
      begin
         ParamByName('COD_jisa').AsString   := UV_vCod_jisa[i];
         ParamByName('NUM_ID').AsString     := UV_vNum_id[i];
         ParamByName('DAT_BUNJU').AsString  := mskDAT_BUNJU.Text;
         ParamByName('NUM_BUNJU').AsString  := edtNUM_BUNJU.Text;
         Execsql;

         GP_query_log(qryU_Bunju, 'LD1I14', 'U', 'Y');
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
      UV_vPart03[i] := Part03;

      UV_vU001[i]   := Trim(copy(Part03,  7, 6));
      UV_vU002[i]   := Trim(copy(Part03, 13, 6));
      UV_vU003[i]   := Trim(copy(Part03, 19, 6));
      UV_vU004[i]   := Trim(copy(Part03, 25, 6));
      UV_vU005[i]   := Trim(copy(Part03, 31, 6));
      UV_vU006[i]   := Trim(copy(Part03, 37, 6));
      UV_vU007[i]   := Trim(copy(Part03, 43, 6));
      UV_vU008[i]   := Trim(copy(Part03, 49, 6));
      UV_vU009[i]   := Trim(copy(Part03, 55, 6));
      UV_vU010[i]   := Trim(copy(Part03, 61, 6));
      UV_vU011[i]   := Trim(copy(Part03, 67, 6));
      UV_vU012[i]   := Trim(copy(Part03, 73, 6));
      UV_vU013[i]   := Trim(copy(Part03, 79, 6));
      UV_vY004[i]   := Trim(copy(Part03, 85, 6));
   end;

   // Save Mode & Msg
   UP_SetMode('SAVE');

   Result := True;
end;

procedure TfrmLD1I14.btnSaveClick(Sender: TObject);
begin
   inherited;

   if UF_Save then
      UP_MoveNum(btnNextNum);
end;

procedure TfrmLD1I14.FormCreate(Sender: TObject);
var i : integer;
begin
   inherited;

   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);

   // 환경설정
   UP_GridSet;
   UP_VarMemSet(0);

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   UV_sBasicSQL := qryGlkwa.SQL.Text;
end;

procedure TfrmLD1I14.grdMasterCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD1I14.btnQueryClick(Sender: TObject);
var i, index : Integer;
    sWhere, sOrderBy : String;
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
            0 : sOrderBy := ' ORDER BY CAST(B.num_samp AS INT), A.num_bunju, A.gubn_part ';
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

      sWhere := sWhere + ' and A.gubn_part = ''03'' ';

      ParamByName('sWhere').Asmemo     :=  sWhere;
      ParamByName('sOrderBy').AsString :=  sOrderBy;
      Active := True;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I14', 'Q', 'Y');
         if Cmb_gubn.Text = '뇨침사검사자' then
         begin
            while not qryGlkwa.Eof do
            begin
               sHangmok := '';
               Hangmok_Set_U011;

               If (Pos('U011',sHangmok) > 0) or (Pos('U012',sHangmok) > 0) Then
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

                  UV_vU001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  7, 6));
                  UV_vU002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 13, 6));
                  UV_vU003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 19, 6));
                  UV_vU004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 25, 6));
                  UV_vU005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 31, 6));
                  UV_vU006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 37, 6));
                  UV_vU007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 43, 6));
                  UV_vU008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 49, 6));
                  UV_vU009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 55, 6));
                  UV_vU010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 61, 6));
                  UV_vU011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 67, 6));
                  UV_vU012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 73, 6));
                  UV_vU013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 79, 6));
                  UV_vY004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 85, 6));
                  UV_vPart03[index]     := FieldByName('desc_glkwa1').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성'
                  else UV_vGumgin_check[index]   := ' ';
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
         else if Cmb_gubn.Text = 'Y004검사자' then
         begin
            while not qryGlkwa.Eof do
            begin
               sHangmok := '';
               Hangmok_Set_U011;

               If Pos('Y004',sHangmok) > 0 Then
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

                  UV_vU001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  7, 6));
                  UV_vU002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 13, 6));
                  UV_vU003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 19, 6));
                  UV_vU004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 25, 6));
                  UV_vU005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 31, 6));
                  UV_vU006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 37, 6));
                  UV_vU007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 43, 6));
                  UV_vU008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 49, 6));
                  UV_vU009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 55, 6));
                  UV_vU010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 61, 6));
                  UV_vU011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 67, 6));
                  UV_vU012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 73, 6));
                  UV_vU013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 79, 6));
                  UV_vY004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 85, 6));
                  UV_vPart03[index]     := FieldByName('desc_glkwa1').AsString;
                  if FieldByName('gubn_preg').AsString ='Y' then
                     UV_vGumgin_check[index]    := '임신중'
                  else if FieldByName('gubn_preg').AsString ='P' then
                     UV_vGumgin_check[index]   := '임신가능성'
                  else UV_vGumgin_check[index]   := ' ';
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

         else
         begin
            UP_VarMemSet(RecordCount - 1);

            while not qryGlkwa.Eof do
            begin
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

               UV_vU001[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString,  7, 6));
               UV_vU002[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 13, 6));
               UV_vU003[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 19, 6));
               UV_vU004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 25, 6));
               UV_vU005[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 31, 6));
               UV_vU006[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 37, 6));
               UV_vU007[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 43, 6));
               UV_vU008[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 49, 6));
               UV_vU009[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 55, 6));
               UV_vU010[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 61, 6));
               UV_vU011[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 67, 6));
               UV_vU012[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 73, 6));
               UV_vU013[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 79, 6));
               UV_vY004[index]       := Trim(copy(FieldByName('desc_glkwa1').AsString, 85, 6));

               UV_vPart03[index]     := FieldByName('desc_glkwa1').AsString;
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

procedure TfrmLD1I14.btnCancelClick(Sender: TObject);
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

procedure TfrmLD1I14.UP_Change(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    cod_HM  : array[1..14] of String;
begin
   inherited;

   cod_HM[1]  := 'U001';
   cod_HM[2]  := 'U002';
   cod_HM[3]  := 'U003';
   cod_HM[4]  := 'U004';
   cod_HM[5]  := 'U005';
   cod_HM[6]  := 'U006';
   cod_HM[7]  := 'U007';
   cod_HM[8]  := 'U008';
   cod_HM[9]  := 'U009';
   cod_HM[10] := 'U010';
   cod_HM[11] := 'U011';
   cod_HM[12] := 'U012';
   cod_HM[13] := 'U013';
   cod_HM[14] := 'Y004';

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

   // Edit Flag Check
   UV_bEdit := True;
end;

procedure TfrmLD1I14.UP_DChange(Sender: TObject);
var index, iPos, iAge : Integer;
    sCode, sName, sMan, sGubn, sValue : String;
    ims_Low, ims_High : Extended;
    cod_HM  : array[1..14] of String;
begin
   inherited;

   cod_HM[1]  := 'U001';
   cod_HM[2]  := 'U002';
   cod_HM[3]  := 'U003';
   cod_HM[4]  := 'U004';
   cod_HM[5]  := 'U005';
   cod_HM[6]  := 'U006';
   cod_HM[7]  := 'U007';
   cod_HM[8]  := 'U008';
   cod_HM[9]  := 'U009';
   cod_HM[10] := 'U010';
   cod_HM[11] := 'U011';
   cod_HM[12] := 'U012';
   cod_HM[13] := 'U013';
   cod_HM[14] := 'Y004';

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
end;

procedure TfrmLD1I14.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
var iRet, i : Integer;
    vCod_chuga : Variant;
    sHmCode, sHmList_03, sSelect : String;
begin
   inherited;
   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   iCntCheck := 0;
   sHmList_03 := '';

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
   //항목찾기
   //---------------------------------------------------------------------------
   sSelect := '';
   with qryPf_hm do
   begin
      Active := False;

      sSelect := ' SELECT P.cod_hm FROM profile_hm P INNER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.gubn_part = ''03'' and H.dat_apply <= ''' + mskDAT_GMGN.Text + '''';
      sSelect := sSelect + ' WHERE P.cod_pf IN (''' + UV_vCod_hul[NewRow-1] + ''',''' + UV_vCod_cancer[NewRow-1] + ''',''' + UV_vCod_jangbi[NewRow-1] + ''')';
      sSelect := sSelect + ' GROUP BY P.cod_hm ';

      SQL.Clear;
      SQL.Add(sSelect);
      Active := True;

      if RecordCount > 0 then
      begin
         while not Eof do
         begin
            sHmList_03 := sHmList_03 + FieldByName('COD_HM').AsString;
            Next;
         end;
      end;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   iRet := GF_MulToSingle(UV_vCod_chuga[NewRow-1], 4, vCod_chuga);
   for i := 0 to iRet-1 do
   begin
      sHmList_03 := sHmList_03 + vCod_chuga[i];
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

   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         sHmList_03 := sHmList_03 + vCod_chuga[i];
      end;
   end;
   //--------------------------------------------------------------------------
   if pos('U001',sHmList_03) > 0 then edtValue1.Color := clWindow
   else                               edtValue1.Color := clBtnFace;
   if pos('U002',sHmList_03) > 0 then edtValue2.Color := clWindow
   else                               edtValue2.Color := clBtnFace;
   if pos('U003',sHmList_03) > 0 then edtValue3.Color := clWindow
   else                               edtValue3.Color := clBtnFace;
   if pos('U004',sHmList_03) > 0 then edtValue4.Color := clWindow
   else                               edtValue4.Color := clBtnFace;
   if pos('U005',sHmList_03) > 0 then edtValue5.Color := clWindow
   else                               edtValue5.Color := clBtnFace;
   if pos('U006',sHmList_03) > 0 then edtValue6.Color := clWindow
   else                               edtValue6.Color := clBtnFace;
   if pos('U007',sHmList_03) > 0 then edtValue7.Color := clWindow
   else                               edtValue7.Color := clBtnFace;
   if pos('U008',sHmList_03) > 0 then edtValue8.Color := clWindow
   else                               edtValue8.Color := clBtnFace;
   if pos('U009',sHmList_03) > 0 then edtValue9.Color := clWindow
   else                               edtValue9.Color := clBtnFace;
   if pos('U010',sHmList_03) > 0 then edtValue10.Color := clWindow
   else                               edtValue10.Color := clBtnFace;
   if pos('U011',sHmList_03) > 0 then edtValue11.Color := clWindow
   else                               edtValue11.Color := clBtnFace;
   if pos('U012',sHmList_03) > 0 then edtValue12.Color := clWindow
   else                               edtValue12.Color := clBtnFace;
   if pos('U013',sHmList_03) > 0 then edtValue13.Color := clWindow
   else                               edtValue13.Color := clBtnFace;
   if pos('Y004',sHmList_03) > 0 then edtValue14.Color := clWindow
   else                               edtValue14.Color := clBtnFace;

   if pos('U001',sHmList_03) > 0 then edtValue1.Enabled := True
   else                               edtValue1.Enabled := False;
   if pos('U002',sHmList_03) > 0 then edtValue2.Enabled := True
   else                               edtValue2.Enabled := False;
   if pos('U003',sHmList_03) > 0 then edtValue3.Enabled := True
   else                               edtValue3.Enabled := False;
   if pos('U004',sHmList_03) > 0 then edtValue4.Enabled := True
   else                               edtValue4.Enabled := False;
   if pos('U005',sHmList_03) > 0 then edtValue5.Enabled := True
   else                               edtValue5.Enabled := False;
   if pos('U006',sHmList_03) > 0 then edtValue6.Enabled := True
   else                               edtValue6.Enabled := False;
   if pos('U007',sHmList_03) > 0 then edtValue7.Enabled := True
   else                               edtValue7.Enabled := False;
   if pos('U008',sHmList_03) > 0 then edtValue8.Enabled := True
   else                               edtValue8.Enabled := False;
   if pos('U009',sHmList_03) > 0 then edtValue9.Enabled := True
   else                               edtValue9.Enabled := False;
   if pos('U010',sHmList_03) > 0 then edtValue10.Enabled := True
   else                               edtValue10.Enabled := False;
   if pos('U011',sHmList_03) > 0 then edtValue11.Enabled := True
   else                               edtValue11.Enabled := False;
   if pos('U012',sHmList_03) > 0 then edtValue12.Enabled := True
   else                               edtValue12.Enabled := False;
   if pos('U013',sHmList_03) > 0 then edtValue13.Enabled := True
   else                               edtValue13.Enabled := False;
   if pos('Y004',sHmList_03) > 0 then edtValue14.Enabled := True
   else                               edtValue14.Enabled := False;

   edtValue1.Text    := UV_vU001[NewRow-1];
   edtValue2.Text    := UV_vU002[NewRow-1];
   edtValue3.Text    := UV_vU003[NewRow-1];
   edtValue4.Text    := UV_vU004[NewRow-1];
   edtValue5.Text    := UV_vU005[NewRow-1];
   edtValue6.Text    := UV_vU006[NewRow-1];
   edtValue7.Text    := UV_vU007[NewRow-1];
   edtValue8.Text    := UV_vU008[NewRow-1];
   edtValue9.Text    := UV_vU009[NewRow-1];
   edtValue10.Text   := UV_vU010[NewRow-1];
   edtValue11.Text   := UV_vU011[NewRow-1];
   edtValue12.Text   := UV_vU012[NewRow-1];
   edtValue13.Text   := UV_vU013[NewRow-1];
   edtValue14.Text   := UV_vY004[NewRow-1];

   mskPDAT_BUNJU.Text := '';
   edtPNUM_BUNJU.Text := '';
   with qryinjouk do
   begin
      qryinjouk.Active := False;
      qryinjouk.ParamByName('num_Jumin').AsString := UV_vNum_jumin[NewRow-1];
      qryinjouk.ParamByName('dat_gmgn').AsString  := UV_vDat_gmgn[NewRow-1];
      qryinjouk.Active := True;
      if RecordCount > 0 then
      begin
         qryPrev.Active := False;
         qryPrev.Filter := '';
         qryPrev.ParamByName('COD_BJJS').AsString := qryinjouk.FieldByName('cod_bjjs').AsString;
         qryPrev.ParamByName('NUM_ID').AsString   := qryinjouk.FieldByName('num_id').AsString;
         qryPrev.ParamByName('DAT_GMGN').AsString := qryinjouk.FieldByName('dat_gmgn').AsString;
         qryPrev.ParamByName('GUBN_PART').AsString := '03';
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

            UV_vDesc_Pglkwa[NewRow-1] := UV_fGulkwa;

            // U001	백혈구(뇨검사)	        03	02 	7	12
            if Trim(Copy(UV_fGulkwa,   7,  6)) <> '' then edtPValue1.Text  := Trim(Copy(UV_fGulkwa,  7,  6))
            else                                          edtPValue1.Text  := '';
            // U002	Nitrite(뇨검사)	        03	03 	13	18
            if Trim(Copy(UV_fGulkwa,  13,  6)) <> '' then edtPValue2.Text  := Trim(Copy(UV_fGulkwa, 13,  6))
            else                                          edtPValue2.Text  := '';
            // U003	PH (뇨검사)	        03	04 	19	24
            if Trim(Copy(UV_fGulkwa,  19,  6)) <> '' then edtPValue3.Text  := Trim(Copy(UV_fGulkwa, 19,  6))
            else                                          edtPValue3.Text  := '';
            // U004	단백(뇨검사)	        03	05 	25    	30
            if Trim(Copy(UV_fGulkwa,  25,  6)) <> '' then edtPValue4.Text  := Trim(Copy(UV_fGulkwa, 25,  6))
            else                                          edtPValue4.Text  := '';
            // U005	당(뇨검사)	        03	06 	31   	36
            if Trim(Copy(UV_fGulkwa,  31,  6)) <> '' then edtPValue5.Text  := Trim(Copy(UV_fGulkwa, 31,  6))
            else                                          edtPValue5.Text  := '';
            // U006	케톤(뇨검사)	        03	07 	37   	42
            if Trim(Copy(UV_fGulkwa,  37,  6)) <> '' then edtPValue6.Text  := Trim(Copy(UV_fGulkwa, 37,  6))
            else                                          edtPValue6.Text  := '';
            // U007	유로빌리노겐(뇨)	03	08 	43  	48
            if Trim(Copy(UV_fGulkwa,  43,  6)) <> '' then edtPValue7.Text  := Trim(Copy(UV_fGulkwa, 43,  6))
            else                                          edtPValue7.Text  := '';
            // U008	빌리루빈(뇨검사)	03	09 	49  	54
            if Trim(Copy(UV_fGulkwa,  49,  6)) <> '' then edtPValue8.Text  := Trim(Copy(UV_fGulkwa, 49,  6))
            else                                          edtPValue8.Text  := '';
            // U009	적혈구(뇨검사)	        03	10 	55  	60
            if Trim(Copy(UV_fGulkwa,  55,  6)) <> '' then edtPValue9.Text  := Trim(Copy(UV_fGulkwa, 55,  6))
            else                                          edtPValue9.Text  := '';
            // U010	요비중	                03	11 	61	66
            if Trim(Copy(UV_fGulkwa,  61,  6)) <> '' then edtPValue10.Text := Trim(Copy(UV_fGulkwa, 61,  6))
            else                                          edtPValue10.Text := '';
            // U011	뇨침사 WBC             	03	12 	67	72
            if Trim(Copy(UV_fGulkwa,  67,  6)) <> '' then edtPValue11.Text := Trim(Copy(UV_fGulkwa, 67,  6))
            else                                          edtPValue11.Text := '';
            // U012	뇨침사 RBC            	03	13 	73	78
            if Trim(Copy(UV_fGulkwa,  73,  6)) <> '' then edtPValue12.Text := Trim(Copy(UV_fGulkwa, 73,  6))
            else                                          edtPValue12.Text := '';
            // Y004	분별잠혈            	03	14 	85	90
            if Trim(Copy(UV_fGulkwa,  85,  6)) <> '' then edtPValue14.Text := Trim(Copy(UV_fGulkwa, 85,  6))
            else

         end;
         qryPrev.Active := False;
      end
      else
      begin
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
         edtPValue14.Text := '';
      end;
      qryinjouk.Active := False;
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

   //Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY')
end;

procedure TfrmLD1I14.btnPInsertClick(Sender: TObject);
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

procedure TfrmLD1I14.UP_Click(Sender: TObject);
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

procedure TfrmLD1I14.UP_MoveNum(Sender: TObject);
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

procedure TfrmLD1I14.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate  then UP_Click(btnDate)
   else if Sender = mskJumin then Up_Click(btnJumin);
end;

procedure TfrmLD1I14.UP_Exit(Sender: TObject);
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

procedure TfrmLD1I14.rbDateClick(Sender: TObject);
begin
   inherited;

   if Sender = rbDate then
   begin
      mskDate.Color    := $00E6FFFE;
      mskDate.Enabled  := True;
      btnDate.Enabled  := True;
      mskFrom.Enabled  := True;
      mskTo.Enabled    := True;
      mskJumin.Color   := clWindow;
      mskJumin.Enabled := False;
      btnJumin.Enabled := False;
      mskDate.SetFocus;

      // 찾기 Enable
      // edtFind.Enabled    := True;
      edtFind2.Enabled   := True;
   end
   else if Sender = rbJumin then
   begin
      mskDate.Color    := clWindow;
      mskDate.Enabled  := False;
      btnDate.Enabled  := False;
      mskFrom.Enabled  := False;
      mskTo.Enabled    := False;
      mskJumin.Color   := $00E6FFFE;
      mskJumin.Enabled := True;
      btnJumin.Enabled := True;
      mskJumin.SetFocus;

      // 찾기 Disable
     // edtFind.Enabled   := False;
      edtFind2.Enabled  := False;
   end;
end;

procedure TfrmLD1I14.edtFindExit(Sender: TObject);
begin
   inherited;
   // 자료가 존재하는지 Check
   if Length(edtFind.Text) < edtFind.MaxLength then exit;

   // 찾기수행
   GF_Find(grdMaster, edtFind.Text, 1, 1, 1, 0);
end;

procedure TfrmLD1I14.edtFind2Exit(Sender: TObject);
begin
   inherited;
 // 자료가 존재하는지 Check
  if Length(edtFind2.Text) < edtFind2.MaxLength then exit;

  // 찾기수행
  GF_Find(grdMaster, edtFind2.Text, 2, 1, 1, 0);
end;

procedure TfrmLD1I14.mskNumExit(Sender: TObject);
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

procedure TfrmLD1I14.FormKeyUp(Sender: TObject; var Key: Word;
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

procedure TfrmLD1I14.UP_EditDisplay(Sender : TObject);
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

procedure TfrmLD1I14.edtValue1Change(Sender: TObject);
begin
  inherited;
   //해당하는 값을 넣어줌.
   if Copy(TEdit(Sender).Name, 1, 8)  = 'edtValue' then
      UP_EditDisplay(Sender);
end;

procedure TfrmLD1I14.cmbRelationChange(Sender: TObject);
begin
   inherited;
   frmLD1I017 := TFrmLD1I017.Create(self);
   frmLD1I017.ShowModal;
end;



end.
