//==============================================================================
// 프로그램 설명 : 생화학 자동계산 결과 LIST
// 시스템        : 분석관리
// 수정일자      : 2014.07.08
// 수정자        : 곽윤설
// 참고사항      : 신규개발 [본원 전문의-김소영]
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================
// 수정일자      : 2018.04.30
// 수정자        : 박수지
// 수정내용      : 검사 항목 계산값 부분 추가 (LIS 불안정...)
// 참고사항      :
//==============================================================================
// 수정일자      : 2018.08.14
// 수정자        : 박수지
// 수정내용      : 입력값과 계산값이 다르면 노란색으로 표시됨.
//                 미납일 경우 입력값은 공란, 계산값은 0.0으로 되어 노란색으로 표시되어야함.
// 참고사항      : 김아름 (생화학)
//==============================================================================

unit LD4Q59;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,math;

type
  TfrmLD4Q59 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    qryGulkwa: TQuery;
    Gride: TGauge;
    mskDate: TDateEdit;
    Cmb_gubn: TComboBox;
    Label9: TLabel;
    Label12: TLabel;
    mskFrom: TMaskEdit;
    MEdt_SampS: TMaskEdit;
    Label10: TLabel;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    gubn_Comb: TComboBox;
    Label1: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryProfile: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_bunju, UV_vNum_bunju, UV_vNum_samp, UV_vDesc_name, UV_v001, UV_v002,
    UV_v003, UV_v004, UV_v005, UV_v006, UV_v007 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q59: TfrmLD4Q59;
  UV_fGulkwa_01, UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q591, LD4Q592;

{$R *.DFM}

procedure TfrmLD4Q59.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_v001      := VarArrayCreate([0, 0], varOleStr);
      UV_v002      := VarArrayCreate([0, 0], varOleStr);
      UV_v003      := VarArrayCreate([0, 0], varOleStr);
      UV_v004      := VarArrayCreate([0, 0], varOleStr);
      UV_v005      := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vDat_bunju, iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vDesc_name, iLength);
      VarArrayRedim(UV_v001,      iLength);
      VarArrayRedim(UV_v002,      iLength);
      VarArrayRedim(UV_v003,      iLength);
      VarArrayRedim(UV_v004,      iLength);
      VarArrayRedim(UV_v005,      iLength);
   end;
end;

function TfrmLD4Q59.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q59.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
//   UP_GridSet;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 0;
   gubn_Comb.ItemIndex := 0;
   grdMaster.Col[9].Width := 0;
   grdMaster.Col[10].Width := 0;

end;

procedure TfrmLD4Q59.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자료를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vDat_bunju[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
      4 : Value := UV_vNum_samp[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : Value := UV_v001[DataRow-1];
      7 : Value := UV_v002[DataRow-1];
      8 : Value := UV_v003[DataRow-1];
      9 : Value := UV_v004[DataRow-1];
      10 : Value := UV_v005[DataRow-1];
   end;
   if (gubn_Comb.ItemIndex = 5) or (gubn_Comb.ItemIndex = 6) or (gubn_Comb.ItemIndex = 7) or (gubn_Comb.ItemIndex = 8) or (gubn_Comb.ItemIndex = 9) or
      (gubn_Comb.ItemIndex = 10) or (gubn_Comb.ItemIndex = 11) or (gubn_Comb.ItemIndex = 12) or (gubn_Comb.ItemIndex = 13)then
   begin
        if (UV_v001[Datarow -1] <> UV_v002[Datarow -1]) then
             grdMaster.CellColor[DataCol,DataRow] := clYellow;
   end;
   if (gubn_Comb.ItemIndex = 14)then
   begin
       if (UV_v001[Datarow -1] <> '') and  (UV_v002[Datarow -1] <> '') then
       begin
            if (strtoint(UV_v001[Datarow -1]) > strtOInt(UV_v002[Datarow -1])) then
                grdMaster.CellColor[DataCol,DataRow] := clYellow;
       end;
   end;
end;

procedure TfrmLD4Q59.btnQueryClick(Sender: TObject);
var index,DataCol : Integer;
    sSelect, sWhere, sOrderBy, sHangmokList: String;
    e1, e2, e3, e4, eRslt : Extended;
    eRslt_2 : Integer;
begin
   inherited;
    sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;
//   UP_GridSet;
   // Grid 초기화
   grdMaster.Rows := 0;
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;
      sSelect := ' Select A.num_jumin, A.Dat_gmgn, A.desc_name, G.cod_jisa, G.dat_bunju, G.num_id, A.num_samp, G.desc_glkwa1, G.desc_glkwa2, ' +
                 '        G.DESC_GLKWA3, G.num_bunju, G.gubn_part, A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.num_jumin, ' +
                 '        A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga ';

      sSelect := sSelect + ' From gulkwa G with(nolock) left outer join Gumgin A with(nolock)';
      sSelect := sSelect + '                         On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';

      if copy(GV_sUserId,1,2) <> '15' then
      begin
           Label7.caption := '검진지사';
           sWhere := ' WHERE A.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''' +
                     '       AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                     ' AND  G.Gubn_Part = ''01'' ' ;
      end
      else
      sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''' +
                '       AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                ' AND  G.Gubn_Part = ''01'' ' ;

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND G.num_bunju <= ''' + mskTo.Text + '''';

         sOrderBy := ' ORDER BY G.num_bunju ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND A.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY A.num_samp '
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q59', 'Q', 'N');
         while Not Eof do
         begin
            Gride.Progress := Gride.Progress + 1;

            sHangmokList := '';
            sHangmokList := UF_hangmokList;

            e1 := 0; e2 := 0; e3 := 0; e4 := 0; eRslt := 0;

            if (qryGulkwa.FieldByName('gubn_part').AsString = '01') then
            begin
               UV_fGulkwa_01 := '';
               UV_fGulkwa1_01 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2_01 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3_01 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1_01, UV_fGulkwa2_01, UV_fGulkwa3_01, UV_fGulkwa_01);
            end;

            if (gubn_Comb.ItemIndex = 0) then
            begin
                if (pos('C022',sHangmokList) > 0) and (pos('C023',sHangmokList) > 0) and (pos('C024',sHangmokList) > 0) then
                begin
                    UP_VarMemSet(index);
                    UV_vNo[index]       := IntToStr(Index+1);
                    UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                    UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                    UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                    UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                    UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  103,  6));
                    UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  109,  6));
                    UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  115,  6));

                    Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 1) then
            begin
                if (pos('C025',sHangmokList) > 0) and (pos('C026',sHangmokList) > 0) and (pos('C027',sHangmokList) > 0) and
                   (pos('C028',sHangmokList) > 0) and (pos('C029',sHangmokList) > 0)then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  121,  6));
                   UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  127,  6));
                   UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  133,  6));
                   UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  139,  6));
                   UV_v005[index]     := Trim(Copy(UV_fGulkwa_01,  151,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 2) then
            begin
                if (pos('C045',sHangmokList) > 0) and (pos('C046',sHangmokList) > 0) and (pos('C047',sHangmokList) > 0) and
                   (pos('C048',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  199,  6));
                   UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  205,  6));
                   UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  211,  6));
                   UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  217,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 3) then
            begin
                if (pos('C001',sHangmokList) > 0) and (pos('C002',sHangmokList) > 0) and (pos('C003',sHangmokList) > 0) and
                   (pos('C004',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  1,  6));
                   UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  7,  6));
                   UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  13,  6));
                   UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  19,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 4) then
            begin
                if (pos('C041',sHangmokList) > 0) and (pos('C042',sHangmokList) > 0) and (pos('C043',sHangmokList) > 0) and
                   (pos('C087',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  181,  6));
                   UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  187,  6));
                   UV_v003[index]     := Trim(Copy(UV_fGulkwa_01,  193,  6));
                   UV_v004[index]     := Trim(Copy(UV_fGulkwa_01,  415,  6));

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 5) then
            begin
                if (pos('C001',sHangmokList) > 0) and (pos('C002',sHangmokList) > 0) and (pos('C003',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  13,  6)); //C003 저장값

                   if (Trim(Copy(UV_fGulkwa_01,  1,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,  7,  6)) <> '') then
                   begin
                   e1 := StrtoFloat(Trim(Copy(UV_fGulkwa_01,  1,  6)));
                   e2 := StrtoFloat(Trim(Copy(UV_fGulkwa_01,  7,  6)));
                   end;
                   eRslt :=  e1 - e2;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt; //C003 계산값

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 6) then
            begin
                if (pos('C004',sHangmokList) > 0) and (pos('C002',sHangmokList) > 0) and (pos('C003',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  19,  6));    //입력값
                   if (Trim(Copy(UV_fGulkwa_01,   7,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,   13,  6)) <> '') then
                   begin
                   e1 := Strtofloat(Trim(Copy(UV_fGulkwa_01,   7,  6)));
                   e2 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  13,  6)));
                   eRslt := (e1 / e2) * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;

                     if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                     else UV_v002[index]     := eRslt;   //계산값
                   end
                   else
                   begin
                     UV_v001[index] := '';
                     UV_v002[index] := 0.0 ;
                   end;
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 7) then
            begin
                if (pos('C007',sHangmokList) > 0) and (pos('C005',sHangmokList) > 0) and (pos('C006',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   if (Trim(Copy(UV_fGulkwa_01,   37,  6)) <> '') then  UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,   37,  6));
                   if (Trim(Copy(UV_fGulkwa_01,   25,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,   31,  6)) <> '') then
                   begin
                   e1 := Strtofloat(Trim(Copy(UV_fGulkwa_01,   25,  6)));
                   e2 := Strtofloat(Trim(Copy(UV_fGulkwa_01,   31,  6)));
                   end;
                   eRslt := e1 -  e2;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt;

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 8) then
            begin
                if (pos('C023',sHangmokList) > 0) and (pos('C024',sHangmokList) > 0) and (pos('C022',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  115,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  109,  6)) <> '') and  (Trim(Copy(UV_fGulkwa_01,  103,  6)) <> '') then
                   begin
                   e1 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  109,  6))));
                   e2 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  103,  6))));

                   eRslt := e1 / e2;
                   eRslt := eRslt * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;
                   end
                   else
                   begin
                     UV_v001[index] := '';
                     UV_v002[index] := 0.0 ;
                   end;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt;
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 9) then
            begin
                if (pos('C027',sHangmokList) > 0) and (pos('C025',sHangmokList) > 0) and (pos('C026',sHangmokList) > 0) and
                   (pos('C028',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  133,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  121,  6)) <>'') and (Trim(Copy(UV_fGulkwa_01,  127,  6)) <>'') and (Trim(Copy(UV_fGulkwa_01,  139,  6))<>'') then
                   begin
                   e1 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  121,  6)));
                   e2 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  127,  6)));
                   e3 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  139,  6)));
                   end ;

                   eRslt := e1 - e2 - (e3 / 5)  ;
                   eRslt := trunc(eRslt + 0.5);
                  { eRslt := eRslt * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;
                   if Pos('.', FloatToStr(eRslt)) = 0 then  UV_v002[index] := FloatToStr(eRslt) + '.0' }
                   UV_v002[index]     := eRslt;

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 10) then
            begin
                if (pos('C025',sHangmokList) > 0) and (pos('C026',sHangmokList) > 0) and (pos('C029',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  151,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  151,  6)) <> '') and  (Trim(Copy(UV_fGulkwa_01,  121,  6)) <> '') then
                   begin
                   e1 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  127,  6)));
                   e2 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  121,  6)));

                   eRslt := (e1 / e2) * 100;
                   eRslt := eRslt * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt;
                   end
                   else
                   begin
                   UV_v001[index] := '';
                   UV_v002[index] := 0.0 ;
                   end;
                   
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 11) then
            begin
                if (pos('C041',sHangmokList) > 0) and (pos('C042',sHangmokList) > 0) and (pos('C043',sHangmokList) > 0)  then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  193,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  181,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,  187,  6)) <> '') then
                   begin
                   e1 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  181,  6))));
                   e2 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  187,  6))));

                   eRslt :=  e1/ e2;
                   eRslt := eRslt * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt;
                   end
                   else
                   begin
                     UV_v001[index] := '';
                     UV_v002[index] := 0.0 ;
                   end;

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 12) then
            begin
                if (pos('C046',sHangmokList) > 0) and (pos('C045',sHangmokList) > 0) and (pos('C048',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  205,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  199,  6)) <> '') and  (Trim(Copy(UV_fGulkwa_01,  217,  6)) <> '') then
                   begin
                   e1 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  199,  6)));
                   e2 := Strtofloat(Trim(Copy(UV_fGulkwa_01,  217,  6)));
                   end;
                   eRslt :=  e1 + e2 ;
                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt ;

                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 13) then
            begin
                if (pos('C047',sHangmokList) > 0) and (pos('C045',sHangmokList) > 0) and (pos('C048',sHangmokList) > 0) then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v001[index]     := Trim(Copy(UV_fGulkwa_01,  211,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  199,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,  199,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,  217,  6)) <> '') then
                   begin
                   e1 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  199,  6))));
                   e2 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  199,  6))));
                   e3 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  217,  6))));

                   eRslt := (e1 / (e2 + e3)) * 100;
                   eRslt := eRslt * 10;
                   eRslt := trunc(eRslt + 0.5);
                   eRslt := eRslt / 10;

                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v002[index] := FloatToStr(eRslt) + '.0'
                   else UV_v002[index]     := eRslt;
                   end
                   else
                   begin
                     UV_v001[index] := '';
                     UV_v002[index] := 0.0 ;
                   end;
                   Inc(index);
                end;
            end
            else if (gubn_Comb.ItemIndex = 14) then
            begin
                if (pos('C074',sHangmokList) > 0) and (pos('C026',sHangmokList) > 0) and  (pos('C025',sHangmokList) > 0)then
                begin
                   UP_VarMemSet(index);
                   UV_vNo[index]       := IntToStr(Index+1);
                   UV_vDat_bunju[index]:= qryGulkwa.FieldByName('dat_bunju').AsString;
                   UV_vNum_bunju[index]:= qryGulkwa.FieldByName('num_bunju').AsString;
                   UV_vNum_samp[index] := qryGulkwa.FieldByName('num_samp').AsString;
                   UV_vDesc_name[index]:= qryGulkwa.FieldByName('desc_name').AsString;
                   UV_v002[index]     := Trim(Copy(UV_fGulkwa_01,  121,  6));
                   if (Trim(Copy(UV_fGulkwa_01,  331,  6)) <> '') and (Trim(Copy(UV_fGulkwa_01,  127,  6)) <> '') then
                   begin
                   e1 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  331,  6)))); //c074
                   e2 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  127,  6)))); //c026
                   //e3 := Strtofloat((Trim(Copy(UV_fGulkwa_01,  139,  6)))); //c028

                   eRslt_2 := ceil(e1 + e2) ;
                   {eRslt_2 := eRslt_2 * 10;
                   eRslt_2 := trunc(eRslt_2 + 0.5);
                   eRslt_2 := eRslt_2 / 10;

                   if Pos('.', FloatToStr(eRslt)) = 0 then UV_v001[index] := FloatToStr(eRslt) + '.0'
                   else }UV_v001[index]     := eRslt_2;
                   end;
                   Inc(index);
                end;
            end;
            next;
         end//begin of while
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end; // qryGulkwa

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q59.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);
   begin
      if (Cmb_gubn.ItemIndex = 0) then
      begin
         mskFrom.Enabled := True;
         mskTo.Enabled := True;
         MEdt_SampS.Enabled := False;
         MEdt_SampE.Enabled := False;
      end
      else if (Cmb_gubn.ItemIndex = 1) then
      begin
         mskFrom.Enabled := False;
         mskTo.Enabled := False;
         MEdt_SampS.Enabled := True;
         MEdt_SampE.Enabled := True;
      end;

      if (gubn_Comb.ItemIndex = 0) then
      begin
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading := 'APO A-I';
         grdMaster.Col[7].Heading := 'APO- B';
         grdMaster.Col[8].Heading := 'B/A-I 비율';
      end
      else if (gubn_Comb.ItemIndex = 1) then
      begin
         grdMaster.Col[9].Width := 80;
         grdMaster.Col[10].Width := 80;
         grdMaster.Col[6].Heading := '콜레스테롤';
         grdMaster.Col[7].Heading := 'HDL-콜레스테롤';
         grdMaster.Col[8].Heading :=  'LDL-콜레스테롤';
         grdMaster.Col[9].Heading :=  '중성지방';
         grdMaster.Col[10].Heading :=  'Cardiac Risk Index';
      end
      else if (gubn_Comb.ItemIndex = 2) then
      begin
         grdMaster.Col[9].Width := 80;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'Fe';
         grdMaster.Col[7].Heading :=  'TIBC';
         grdMaster.Col[8].Heading :=  '철포화율';
         grdMaster.Col[9].Heading :=  'UIBC';
      end
      else if (gubn_Comb.ItemIndex = 3) then
      begin
         grdMaster.Col[9].Width := 80;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  '총단백';
         grdMaster.Col[7].Heading :=  '알부민';
         grdMaster.Col[8].Heading :=  '글로부린';
         grdMaster.Col[9].Heading :=  'A/G 비율';
      end
      else if (gubn_Comb.ItemIndex = 4) then
      begin
         grdMaster.Col[9].Width := 80;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  '뇨소질소';
         grdMaster.Col[7].Heading :=  '크레아티닌';
         grdMaster.Col[8].Heading :=  'BUN/Cr 비율';
         grdMaster.Col[9].Heading :=  'GFR';
      end
      else if (gubn_Comb.ItemIndex = 5) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C003(글로불린) 입력 값';
         grdMaster.Col[7].Heading :=  'C003(글로불린) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 6) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C004(A/G비율) 입력 값';
         grdMaster.Col[7].Heading :=  'C004(A/G비율) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 7) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C007(간접빌리루빈) 입력 값';
         grdMaster.Col[7].Heading :=  'C007(간접빌리루빈) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 8) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C024(B/A -1비율) 입력 값';
         grdMaster.Col[7].Heading :=  'C024(B/A -1비율) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 9) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C027(LDL콜레스테롤) 입력 값';
         grdMaster.Col[7].Heading :=  'C027(LDL콜레스테롤) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 10) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C029(cardiac risk index) 입력 값';
         grdMaster.Col[7].Heading :=  'C029(cardiac risk index) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 11) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C043(BUN/Cr 비) 입력 값';
         grdMaster.Col[7].Heading :=  'C043(BUN/Cr 비) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 12) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C046(TIBC) 입력 값';
         grdMaster.Col[7].Heading :=  'C046(TIBC) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 13) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C047(철포화율) 입력 값';
         grdMaster.Col[7].Heading :=  'C047(철포화율) 계산 값 ';
      end
      else if (gubn_Comb.ItemIndex = 14) then
      begin
         grdMaster.Col[8].Width := 0;
         grdMaster.Col[9].Width := 0;
         grdMaster.Col[10].Width := 0;
         grdMaster.Col[6].Heading :=  'C074+C026';
         grdMaster.Col[7].Heading :=  'C025 ';
      end;
   end;
end;


function TfrmLD4Q59.UF_hangmokList : String;
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

function TfrmLD4Q59.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q59.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then //미리보기
  begin
     if (gubn_Comb.ItemIndex <> 1) then
     begin
        frmLD4Q591 := TfrmLD4Q591.Create(Self);
        frmLD4Q591.QuickRep.Preview;
     end
     else
     begin
        frmLD4Q592 := TfrmLD4Q592.Create(Self);
        frmLD4Q592.QuickRep.Preview;
     end;
  end
  else   //프린트
  begin
     if (gubn_Comb.ItemIndex <> 1) then
     begin
        frmLD4Q591 := TfrmLD4Q591.Create(Self);
        frmLD4Q591.QuickRep.Print;
     end
     else
     begin
        frmLD4Q592 := TfrmLD4Q592.Create(Self);
        frmLD4Q592.QuickRep.Preview;
     end;
  end;
end;

procedure TfrmLD4Q59.SBut_ExcelClick(Sender: TObject);
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
