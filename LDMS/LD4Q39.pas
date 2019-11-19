//==============================================================================
// 프로그램 설명 : 검사항목 결과 List(혈청학전용)
// 시스템        : 통합검진
// 수정일자      : 2011.09.05
// 수정자        : 송재호
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 검사항목 결과 List(혈청학전용)
// 시스템        : 통합검진
// 수정일자      : 2011.09.07
// 수정자        : 송재호
// 수정내용      : 결과유형 추가 및 출력 생성
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 검사항목 결과 List(혈청학전용)
// 시스템        : 통합검진
// 수정일자      : 2011.12.06                                
// 수정자        : 송재호
// 수정내용      : 검사항목에 S034 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2013.01.08
// 수정자        : 김희석
// 수정내용      : S001,S003,s036 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2013.7.23
// 수정자        : 김희석
// 수정내용      : 한의 재단진단면역학팀1300097
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.28
// 수정자        : 곽윤설
// 수정내용      : 엑셀버튼 추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.03
// 수정자        : 곽윤설
// 수정내용      : E069 -> 50이상
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.05.12
// 수정자        : 곽윤설
// 수정내용      : 1. S099 : S007S099 둘중 하나라도 있을 경우 조회가능
//                 2. S091 : S008S091              "
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.05.15
// 수정자        : 곽윤설
// 수정내용      : 검사항목 - TT01TT02 항목에서 제거, S016 추가
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.05.29
// 수정자        : 곽윤설
// 수정내용      : 검사항목 - S021 추가
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.09.17
// 수정자        : 곽윤설
// 수정내용      : 검사항목 - T039 추가
// 참고사항      : [본부-한두례]
//==============================================================================
// 수정일자      : 2014.09.22
// 수정자        : 곽윤설
// 수정내용      : 검사항목 - TT01, TT02 삭제
// 참고사항      : 
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.02.11
// 수정자        : 곽윤설
// 수정내용      : S004추가  (S003과 함께 조회되도록)
// 참고사항      :
//==============================================================================
// 수정일자      : 2015.04.09
// 수정자        : 곽윤설
// 수정내용      : S010, S011, S012, S020, S070, S090, SE01, SE31, T041, T042, T043  코드추가
// 참고사항      : [한의 재단진단검사의학팀1500304 : 본원-한미영]
//==============================================================================
// 수정일자      :  2015.08.10
// 수정자        : 박수지
// 수정내용      : C020, S033, SE03 코드추가
// 참고사항      : 본원-한미영
//==============================================================================
unit LD4Q39;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD4Q39 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    RadioGroup1: TRadioGroup;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    ck_Plus: TCheckBox;
    ck_LPlus: TCheckBox;
    ck_hold: TCheckBox;
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
    Chk_S091: TCheckBox;
    Chk_E069: TCheckBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryProfile: TQuery;
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
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure cmbHMChange(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN, UV_vS007, UV_vSampNo, UV_vDesc_dc,Uv_vRack : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q39: TfrmLD4Q39;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q391;

{$R *.DFM}

procedure TfrmLD4Q39.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q39.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q39.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vS007  := VarArrayCreate([0, 0], varOleStr);
      UV_vRack  := VarArrayCreate([0, 0], varOleStr);

      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vS007,  iLength);
      VarArrayRedim(UV_vRack,  iLength);

      VarArrayRedim(UV_vDesc_dc, iLength);
      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD4Q39.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (edtBDate.Text = '') OR (cmbHM.Text = '') then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q39.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   cmbHM.ItemIndex := 0;
end;

procedure TfrmLD4Q39.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vRack[DataRow-1];
      3 : Value := UV_vSampNo[DataRow-1];
      4 : Value := UV_vBUN[DataRow-1];
      5 : Value := UV_vName[DataRow-1];
      6 : Value := UV_vS007[DataRow-1];
   end;

   if (DataCol =  6) and (UV_vS007[DataRow-1] = '결과無') then grdMaster.CellColor[DataCol, DataRow] := $00FFA2A2;
end;

procedure TfrmLD4Q39.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    temp_S099, temp_S091 : String;
    vCod_chuga : Variant;
    Check :Boolean;
begin
   inherited;   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;
   if ((cmbHM.Text  <>  'S007S099')       and (cmbHM.Text  <>  'S007S099(비교)') and
       (cmbHM.Text  <>  'S008S091(비교)'))  then
   begin
      if (ck_Plus.Checked = False)  AND (ck_LPlus.Checked = False)   AND
         (ck_hold.Checked = False)  AND (ck_nothing.Checked = False) AND
         (Ck_Exist.Checked = False) AND (Chk_S091.Checked = False)   AND
         (Chk_E069.Checked = False) then
      begin
         ShowMessage('결과유형을 하나 이상 선택하세요.');
         exit;
      end;
   end;
   if      (cmbHM.Text = 'S099 (HBs Ag)') then grdMaster.Col[6].Heading := 'S007 또는 S099'
   else if (cmbHM.Text  =  'S091 (HBs Ab)') then grdMaster.Col[6].Heading := 'S008 또는 S091'
   else grdMaster.Col[6].Heading := cmbHM.text;
   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   Gride.Progress := 0;
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  B.desc_rackno,G.desc_name, G.num_id, G.num_jumin, G.dat_gmgn, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, K.desc_glkwa1, K.desc_glkwa2, K.desc_glkwa3, ' +
                 ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                 ' FROM Gulkwa K with(nolock) left outer join Gumgin G with(nolock) ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '                            left outer join Bunju  B with(nolock) ON K.cod_jisa = B.cod_jisa and K.num_id = B.num_id and K.dat_gmgn = B.dat_gmgn ' +
                 '                            left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc ' +
                 '                            left outer join Tkgum T  with(nolock) ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ' +
                 ' AND K.Dat_Bunju  = ''' + edtBDate.Text + ''' ' ;

      if (cmbHM.Text = 'E069') or (cmbHM.Text = 'E068') or (cmbHM.Text = 'SE31') or (cmbHM.Text = 'S016') or (cmbHM.Text = 'S021') or (cmbHM.Text = 'T039') or
         (cmbHM.Text = 'S010') or (cmbHM.Text = 'S011') or (cmbHM.Text = 'S012') or (cmbHM.Text = 'S020') or (cmbHM.Text = 'SE01') or
         (cmbHM.Text = 'T041') or (cmbHM.Text = 'T042') or (cmbHM.Text = 'T043') or (cmbHM.Text = 'C020') or (cmbHM.Text = 'S033') or
         (cmbHM.Text = 'SE03') then
          sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''05'') '
      else
          sWhere  := sWhere + ' AND ( K.Gubn_Part  = ''07'') ';

      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + ''' ';

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(BunjuNo1.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju >= ''' + BunjuNo1.Text + ''' ';
         if Trim(BunjuNo2.Text) <> '' then
            sWhere := sWhere + ' AND k.num_bunju <= ''' + BunjuNo2.Text + ''' ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp >= ''' + MEdt_SampS.Text + ''' ';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND g.num_samp <= ''' + MEdt_SampE.Text + ''' ';
      end;

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         1 : sOrderBy := ' ORDER BY B.Desc_rackno ';
         2 : sOrderBy := ' ORDER BY K.Num_Bunju ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q39', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            sHul_List := '';
            sHul_List := UF_hangmokList;

            if ((cmbHM.Text = 'S007') AND (pos('S007', sHul_List) > 0)) OR
               ((cmbHM.Text = 'S091 (HBs Ab)') AND ((pos('S008', sHul_List) > 0) or (pos('S091', sHul_List) > 0))) OR
               ((cmbHM.Text = 'S099 (HBs Ag)') AND ((pos('S099', sHul_List) > 0) or (pos('S007', sHul_List) > 0))) OR
               ((cmbHM.Text = 'S034') AND (pos('S034', sHul_List) > 0)) or
               ((cmbHM.Text = 'S001') AND (pos('S001', sHul_List) > 0)) OR
               ((cmbHM.Text = 'S003S004') AND ((pos('S003', sHul_List) > 0) or (pos('S004', sHul_List) > 0))) or //20150211
               ((cmbHM.Text = 'S003') AND (pos('S003', sHul_List) > 0)) or
               ((cmbHM.Text = 'S036') AND (pos('S036', sHul_List) > 0)) OR
               ((cmbHM.Text = 'E069') AND (pos('E069', sHul_List) > 0)) or
               ((cmbHM.Text = 'E068') AND (pos('E068', sHul_List) > 0)) or
               ((cmbHM.Text = 'SE31') AND (pos('SE31', sHul_List) > 0)) or
               ((cmbHM.Text = 'S016') AND (pos('S016', sHul_List) > 0)) or
               ((cmbHM.Text = 'S021') AND (pos('S021', sHul_List) > 0)) or
               ((cmbHM.Text = 'T039') AND (pos('T039', sHul_List) > 0)) or
               ((cmbHM.Text = 'S007S099') AND ((pos('S099', sHul_List) > 0) or (pos('S007', sHul_List) > 0))) or
               ((cmbHM.Text = 'S007S099(비교)') and  ((pos('S007', sHul_List) > 0) and (pos('S099', sHul_List) > 0))) or
               ((cmbHM.Text = 'S008S091(비교)') and  ((pos('S008', sHul_List) > 0) and (pos('S091', sHul_List) > 0))) or
               ((cmbHM.Text = 'S010') AND (pos('S010', sHul_List) > 0)) or
               ((cmbHM.Text = 'S011') AND (pos('S011', sHul_List) > 0)) or
               ((cmbHM.Text = 'S012') AND (pos('S012', sHul_List) > 0)) or
               ((cmbHM.Text = 'S020') AND (pos('S020', sHul_List) > 0)) or
               ((cmbHM.Text = 'S070') AND (pos('S070', sHul_List) > 0)) or
               ((cmbHM.Text = 'S090') AND (pos('S090', sHul_List) > 0)) or
               ((cmbHM.Text = 'SE01') AND (pos('SE01', sHul_List) > 0)) or
               ((cmbHM.Text = 'T041') AND (pos('T041', sHul_List) > 0)) or
               ((cmbHM.Text = 'T042') AND (pos('T042', sHul_List) > 0)) or
               ((cmbHM.Text = 'S004') AND (pos('S004', sHul_List) > 0)) or
               ((cmbHM.Text = 'T043') AND (pos('T043', sHul_List) > 0)) or
               ((cmbHM.Text = 'C020') AND (pos('C020', sHul_List) > 0)) or
               ((cmbHM.Text = 'S033') AND (pos('S033', sHul_List) > 0)) or
               ((cmbHM.Text = 'SE03') AND (pos('SE03', sHul_List) > 0))
            then
            begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString   := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString := qryBunju.FieldByName('dat_gmgn').AsString;
               if (cmbHM.Text = 'E069') or (cmbHM.Text = 'E068') or (cmbHM.Text = 'SE31') or (cmbHM.Text = 'S016') or (cmbHM.Text = 'S021') or
                  (cmbHM.Text = 'T039') or (cmbHM.Text = 'S010') or (cmbHM.Text = 'S011') or (cmbHM.Text = 'S012') or
                  (cmbHM.Text = 'S020') or (cmbHM.Text = 'SE01') or (cmbHM.Text = 'T041') or (cmbHM.Text = 'SE03') or
                  (cmbHM.Text = 'T042') or (cmbHM.Text = 'T043') or (cmbHM.Text = 'C020') or (cmbHM.Text = 'S033')
               then  qryGulkwa.ParamByName('gubn_part').AsString := '05'
               else  qryGulkwa.ParamByName('gubn_part').AsString := '07';

               qryGulkwa.open;

               if qryGulkwa.RecordCount > 0 then
               begin
                  while not qryGulkwa.Eof do
                  begin
                     Check := False;

                     UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                     UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                     UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                     GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                     UV_VBUN[Index]    := qryGulkwa.FieldByName('Num_Bunju').AsString;
                     UV_VRack[Index]   := qryBunju.FieldByName('Desc_rackno').AsString;
                     UV_vName[Index]   := qryGulkwa.FieldByName('desc_name').AsString;
                     UV_vSampNo[Index] := qryGulkwa.FieldByName('num_samp').AsString;

                     if (cmbHM.Text = 'S007S099') AND (pos('S007', sHul_List) > 0) and (pos('S099', sHul_List) > 0)then
                     begin
                        if ((Trim(Copy(UV_fGulkwa, 31,  6)) = '약양성') and ((StrToFloat(Copy(UV_fGulkwa, 229,  6)) > 1.0) or (StrToFloat(Copy(UV_fGulkwa, 229,  6)) < 50.0))) or //((StrToFloat(Copy(UV_fGulkwa, 229,  6)) < 1.0) or (StrToFloat(Copy(UV_fGulkwa, 229,  6)) >9.99)))
                           ((Trim(Copy(UV_fGulkwa, 31,  6)) = '양성')   and (StrToFloat(Copy(UV_fGulkwa, 229,  6)) >= 50)) or
                           ((Trim(Copy(UV_fGulkwa, 31,  6)) = '음성')   and (StrToFloat(Copy(UV_fGulkwa, 229,  6)) <= 1))  then
                        begin
                           UV_vS007[Index] := 'S007:'+Trim(Copy(UV_fGulkwa, 31,  6))+'S099:'+Trim(Copy(UV_fGulkwa, 229,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'S007S099(비교)') AND (pos('S007', sHul_List) > 0) and (pos('S099', sHul_List) > 0)then
                     begin
                        UV_vS007[Index] := 'S007:'+Trim(Copy(UV_fGulkwa, 31,  6))+' S099:'+Trim(Copy(UV_fGulkwa, 229,  6));
                        Check := True;
                     end
                     else if (cmbHM.Text = 'S008S091(비교)') AND (pos('S008', sHul_List) > 0) and (pos('S091', sHul_List) > 0)then
                     begin
                        UV_vS007[Index] := 'S008:'+Trim(Copy(UV_fGulkwa, 37,  6))+' S091:'+Trim(Copy(UV_fGulkwa, 181,  6));
                        Check := True;
                     end
                     //20140512~ 곽윤설
                     else if (cmbHM.Text = 'S099 (HBs Ag)') AND ((pos('S007', sHul_List) > 0) or (pos('S099', sHul_List) > 0)) then
                     begin
                        if (ck_Plus.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 31,  6)) = '양성') or
                                                         ((Trim(Copy(UV_fGulkwa, 229,  6)) <> '') and (StrToFloat(Trim(Copy(UV_fGulkwa, 229,  6))) >= 50.0))) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6)) +' '+ Trim(Copy(UV_fGulkwa, 229,  6)) ;
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 31,  6)) = '약양성') or ((Trim(Copy(UV_fGulkwa, 229,  6)) <> '') and
                                                          ((StrToFloat(Trim(Copy(UV_fGulkwa, 229,  6))) > 1.0) and (StrToFloat(Trim(Copy(UV_fGulkwa, 229,  6))) < 50.0)))) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6))+' '+ Trim(Copy(UV_fGulkwa, 229,  6)) ;
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 31,  6)) = '보류') or (Trim(Copy(UV_fGulkwa, 229,  6)) = '보류')) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6))+' '+ Trim(Copy(UV_fGulkwa, 229,  6)) ;
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 31,  6)) = '') and (Trim(Copy(UV_fGulkwa, 229,  6)) = '')) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6)) +' '+ Trim(Copy(UV_fGulkwa, 229,  6)) ;
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 31,  6)) <> '') or (Trim(Copy(UV_fGulkwa, 229,  6)) <> ''))then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6)) +' '+ Trim(Copy(UV_fGulkwa, 229,  6)) ;
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'S091 (HBs Ab)') AND ((pos('S008', sHul_List) > 0) or (pos('S091', sHul_List) > 0)) then
                     begin
                        if (ck_Plus.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 37,  6)) = '양성') or
                                                         ((Trim(Copy(UV_fGulkwa, 181,  6)) <> '') and (StrToFloat(Trim(Copy(UV_fGulkwa, 181,  6))) >= 30.0))) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6)) + ' ' + Trim(Copy(UV_fGulkwa, 181,  6)) ;
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 37,  6)) = '약양성') or ((Trim(Copy(UV_fGulkwa, 181,  6)) <> '') and
                                                          ((StrToFloat(Trim(Copy(UV_fGulkwa, 181,  6))) > 10.0) and (StrToFloat(Trim(Copy(UV_fGulkwa, 181,  6))) < 30.0)))) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6)) + ' ' + Trim(Copy(UV_fGulkwa, 181,  6)) ;
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 37,  6)) = '보류') or (Trim(Copy(UV_fGulkwa, 181,  6)) = '보류')) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6)) + ' ' + Trim(Copy(UV_fGulkwa, 181,  6)) ;
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 37,  6)) = '') and (Trim(Copy(UV_fGulkwa, 181,  6)) = '')) then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6)) + ' ' + Trim(Copy(UV_fGulkwa, 181,  6)) ;
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND ((Trim(Copy(UV_fGulkwa, 37,  6)) <> '') or (Trim(Copy(UV_fGulkwa, 181,  6)) <> ''))then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6)) + ' ' + Trim(Copy(UV_fGulkwa, 181,  6)) ;
                           Check := True;
                        end;
                        if (Chk_S091.Checked = True) AND (Trim(Copy(UV_fGulkwa, 181,  6)) <> '') then
                        begin
                           if (StrToFloat(Trim(Copy(UV_fGulkwa, 181,  6))) >= 1000) then
                           begin
                              UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 181,  6));
                              Check := True;
                           end;
                        end;
                     end
                     else if (cmbHM.Text = 'S034') AND (pos('S034', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'S001') AND (pos('S001', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND  (Trim(Copy(UV_fGulkwa, 7,  6)) <> '') then
                        begin
                           if (StrToFloat(Copy(UV_fGulkwa, 7,  6)) > 0.3 ) then
                           begin
                              UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                              Check := True;
                           end
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'S003') AND (pos('S003', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 19,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 19,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 19,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 19,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 19,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 19,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 19,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 19,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 19,  6));
                           Check := True;
                        end;
                     end
                     /////////////////////////////////////////////////////////////
                     else if (cmbHM.Text = 'S003S004') AND ((pos('S003', sHul_List) > 0) or (pos('S004', sHul_List) > 0)) then
                     begin    //20150211
                        if (ck_Plus.Checked = True) then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 19,  6)) = '양성') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S003:' + Trim(Copy(UV_fGulkwa, 19,  6)) + ' ';
                               Check := True;
                            end;
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) = '양성') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S004:' + Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                        if (ck_LPlus.Checked = True) then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 19,  6)) = '약양성') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S003:' + Trim(Copy(UV_fGulkwa, 19,  6)) + ' ';
                               Check := True;
                            end;
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) = '약양성') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S004:' + Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                        if (ck_hold.Checked = True) then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 19,  6)) = '보류') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S003:' + Trim(Copy(UV_fGulkwa, 19,  6)) + ' ';
                               Check := True;
                            end;
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) = '보류') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S004:' + Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                        if (ck_nothing.Checked = True)  then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                            begin
                                UV_vS007[Index] := UV_vS007[Index] + 'S003:' + Trim(Copy(UV_fGulkwa, 19,  6)) + ' ';
                                Check := True;
                            end;
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S004:' + Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                        if (Ck_Exist.Checked = True) then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 19,  6)) <> '') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S003:' + Trim(Copy(UV_fGulkwa, 19,  6)) + ' ';
                               Check := True;
                            end;
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) <> '') then
                            begin
                               UV_vS007[Index] := UV_vS007[Index] + 'S004:' + Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                     end
                     /////////////////////////////////////////////////////////////
                     else if (cmbHM.Text = 'S004') AND  (pos('S004', sHul_List) > 0) then
                     begin    //20160314
                        if (Ck_Exist.Checked = True) then
                        begin
                            if (Trim(Copy(UV_fGulkwa, 25,  6)) = '음성') OR (Trim(Copy(UV_fGulkwa, 25,  6)) = '양성') then
                            begin
                               UV_vS007[Index] :='결과 오류 : '+ Trim(Copy(UV_fGulkwa, 25,  6));
                               Check := True;
                            end;
                        end;
                     end
                     /////////////////////////////////////////////////////////////
                     else if (cmbHM.Text = 'S036') AND (pos('S036', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 55,  6)) <> '') then
                        begin
                           if (StrToFloat(Trim(Copy(UV_fGulkwa, 55,  6))) > 9.9) then
                           begin
                              UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                              Check := True;
                           end;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 55,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 55,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 55,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 55,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 55,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'E069') AND (pos('E069', sHul_List) > 0) then
                     begin
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 433,  6)) = '') then
                        begin
                          UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 433,  6));
                          Check := True;
                        end;
                        if (Chk_E069.Checked = True) AND (Trim(Copy(UV_fGulkwa, 433,  6)) <> '') then
                        begin
                          if (StrToFloat(Trim(Copy(UV_fGulkwa, 433,  6))) >= 50) then //20140503 곽윤설 65이상->50이상 변경
                          begin
                             UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 433,  6));
                             Check := True;
                          end;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 433,  6)) <> '') then
                        begin
                          UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 433,  6));
                          Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'E068') AND (pos('E068', sHul_List) > 0) then
                     begin
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 427,  6)) = '') then
                        begin
                          UV_vS007[Index] := '결과 無';
                          Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 427,  6)) <> '') then
                        begin
                         if (StrToFloat(Trim(Copy(UV_fGulkwa, 427,  6))) = 39) then
                          begin
                             UV_vS007[Index] := '40미만';
                             Check := True;
                          end
                          else if (Trim(Copy(UV_fGulkwa, 427,  6)) <> '') then 
                          begin
                          UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 427,  6));
                          Check := True;
                          end;
                        end;
                     end
                     else if (cmbHM.Text = 'S016') AND (pos('S016', sHul_List) > 0) then      //20140515 곽윤설
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 7,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'S021') AND (pos('S021', sHul_List) > 0) then    // 20140529
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 43,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 43,  6));
                           Check := True;
                        end;
                     end
                     else if (cmbHM.Text = 'SE31') AND (pos('SE31', sHul_List) > 0) then
                     begin
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 445,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 445,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 445,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 445,  6));
                           Check := True;
                        end;
                        if (ck_Plus.Checked = True) AND  (Trim(Copy(UV_fGulkwa, 445,  6)) <> '') then
                        begin
                           if  (StrToFloat(Copy(UV_fGulkwa, 445,  6)) >=20) then
                           begin
                               UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 445,  6));
                               Check := True;
                           end;
                        end;
                     end
                     else if (cmbHM.Text = 'T039') AND (pos('T039', sHul_List) > 0) then  //20140917
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) <> '') AND
                           (StrToFloat(Trim(Copy(UV_fGulkwa, 313,  6))) >= 25.0) then //흡연
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) <> '') AND
                           (StrToFloat(Trim(Copy(UV_fGulkwa, 313,  6))) < 25) then //비흡연
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 313,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 313,  6));
                           Check := True;
                        end;
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
                     end
                     ///////////////////////////////////////////// 20150409 곽윤설. 한미영 책임 요청
                     else if (cmbHM.Text = 'S010') AND (pos('S010', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 37,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 37,  6));
                           Check := True;
                        end;
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
                     end /////////////////////////////////////////////
                     else if (cmbHM.Text = 'S011') AND (pos('S011', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 25,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 25,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 25,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 25,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 25,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 25,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 25,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 25,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 25,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'S012') AND (pos('S012', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 31,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 31,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 31,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 31,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 31,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 31,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'S020') AND (pos('S020', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 1,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 1,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 1,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 1,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 1,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 1,  6));
                           Check := True;
                        end;
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
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'S070') AND (pos('S070', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 121,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 121,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 121,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 121,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 121,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 121,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 121,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 121,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 121,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 121,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'S090') AND (pos('S090', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 367,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 367,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 367,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 367,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 367,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 367,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'SE01') AND (pos('SE01', sHul_List) > 0) then
                     begin
                        if (ck_Plus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 247,  6)) = '양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 247,  6));
                           Check := True;
                        end;
                        if (ck_LPlus.Checked = True) AND (Trim(Copy(UV_fGulkwa, 247,  6)) = '약양성') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 247,  6));
                           Check := True;
                        end;
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 247,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 247,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 247,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 247,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 247,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 247,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'T041') AND (pos('T041', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 349,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 349,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 349,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 349,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 349,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 349,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'T042') AND (pos('T042', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 355,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 355,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 355,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 355,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 355,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 355,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'T043') AND (pos('T043', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 361,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 361,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 361,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 361,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 361,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 361,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'C020') AND (pos('C020', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 67,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 67,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 67,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 67,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 67,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'S033') AND (pos('S033', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 457,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 457,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 457,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 457,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 457,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 457,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     else if (cmbHM.Text = 'SE03') AND (pos('SE03', sHul_List) > 0) then
                     begin
                        if (ck_hold.Checked = True) AND (Trim(Copy(UV_fGulkwa, 463,  6)) = '보류') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 463,  6));
                           Check := True;
                        end;
                        if (ck_nothing.Checked = True) AND (Trim(Copy(UV_fGulkwa, 463,  6)) = '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 463,  6));
                           Check := True;
                        end;
                        if (Ck_Exist.Checked = True) AND (Trim(Copy(UV_fGulkwa, 463,  6)) <> '') then
                        begin
                           UV_vS007[Index] := Trim(Copy(UV_fGulkwa, 463,  6));
                           Check := True;
                        end;
                     end/////////////////////////////////////////////
                     ;

                     if Check = True then
                     begin
                        UP_VarMemSet(Index+1);
                        UV_vNo[Index]  := IntToStr(Index+1);
                        Inc(Index);
                     end;
                     qryGulkwa.Next;
                  end;
                  qryGulkwa.Close;
               End;
            end;
             {  if Check = True then
               begin
                  UP_VarMemSet(Index+1);
                  UV_vNo[Index]  := IntToStr(Index+1);

                  Inc(Index);
               end;}
            qryBunju.Next;
         End;

         // Table과 Disconnect
         qryBunju.Close;
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

procedure TfrmLD4Q39.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD4Q39.UP_Click(Sender: TObject);
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

procedure TfrmLD4Q39.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q391 := TfrmLD4Q391.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q391.QuickRep.Preview
  else                                frmLD4Q391.QuickRep.Print;

end;

procedure TfrmLD4Q39.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
   Cmb_gubn.ItemIndex:=1;
end;

procedure TfrmLD4Q39.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;

  with QRCompositeReport do
  begin

  end;
end;

function TfrmLD4Q39.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
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

function TfrmLD4Q39.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q39.cmbHMChange(Sender: TObject);
begin
  inherited;
  if (cmbHM.Text = 'S091') then  Chk_S091.Visible := true
  else
  begin
       Chk_S091.Visible :=False;
       Chk_S091.checked :=False;
  end;
  if (cmbHM.Text = 'E069') then  Chk_E069.Visible := true
  else
  begin
       Chk_E069.Visible :=False;
       Chk_E069.checked :=False;
  end;
  if (cmbHM.Text = 'T039') then
  begin
     Ck_Plus.Caption := '흡연';
     Ck_LPlus.Caption := '비흡연';
  end
  else
  begin
     Ck_Plus.Caption := '양성';
     Ck_LPlus.Caption := '약양성';
  end;
end;

procedure TfrmLD4Q39.SBut_ExcelClick(Sender: TObject);
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

end.
