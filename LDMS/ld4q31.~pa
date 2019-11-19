//==============================================================================
// 프로그램 설명 : 세포병리검사의뢰서 리스트
// 시스템        : 분석관리
// 수정일자      : 2010.05.13
// 수정자        : 김승철
// 수정내용      :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 세포병리검사의뢰서 리스트
// 시스템        : 분석관리
// 수정일자      : 2017.07.20
// 수정자        : 박수지
// 수정내용      : 조직 분주 돌릴때 L.M.P 입력 되도록
// 참고사항      :
//==============================================================================
unit LD4Q31;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges, ComObj ;

type
  TfrmLD4Q31 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    cmbJisa: TComboBox;
    mskFrom: TDateEdit;
    rbDate: TRadioButton;
    grdMaster: TtsGrid;
    btnsDate: TSpeedButton;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    Label4: TLabel;
    ComB_ShEndo: TComboBox;
    Label2: TLabel;
    CmbOrder: TComboBox;
    rbPan_Date: TRadioButton;
    mskPandate: TDateEdit;
    btn_Pandate: TSpeedButton;
    qryCa_desc_buwi_Max: TQuery;
    qryU_Cell: TQuery;
    Edte_No: TEdit;
    Label17: TLabel;
    Edts_No: TEdit;
    Rd_No: TRadioButton;
    rd_GetDate: TRadioButton;
    msk_Getdate: TDateEdit;
    btn_Getdate: TSpeedButton;
    CK_Max: TCheckBox;
    Edt_Max: TEdit;
    Gauge2: TGauge;
    SBut_Excel: TSpeedButton;
    chk_jumin: TCheckBox;
    qryDoctor: TQuery;

    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure UP_Click(Sender: TObject);
    procedure up_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure btnPrintClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    procedure CK_MaxClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCount, UV_vGubn_check, UV_vDat_gmgn, UV_vDesc_name, UV_vNum_jumin, UV_vAge, UV_vSex,
    UV_vCod_jisa, UV_vNum_id, UV_vCod_cell, UV_vCod_hm, UV_vDesc_buwi, UV_vDesc_sokun,
    UV_vDat_take, UV_vDesc_limsang, UV_vPatient_limsang, UV_vRequest_Hangmok, UV_vCod_doctor : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q31: TfrmLD4Q31;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q301, LD4Q311;

{$R *.DFM}

procedure TfrmLD4Q31.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCount      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_check := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vAge        := VarArrayCreate([0, 0], varOleStr);
      UV_vSex        := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_take   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_limsang     := VarArrayCreate([0, 0], varOleStr);
      UV_vPatient_limsang  := VarArrayCreate([0, 0], varOleStr);
      UV_vRequest_Hangmok  := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_doctor       := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCount,      iLength);
      VarArrayRedim(UV_vGubn_check, iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vAge,        iLength);
      VarArrayRedim(UV_vSex,        iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vCod_cell,   iLength);
      VarArrayRedim(UV_vCod_hm,     iLength);
      VarArrayRedim(UV_vDesc_buwi,  iLength);
      VarArrayRedim(UV_vDesc_sokun, iLength);
      VarArrayRedim(UV_vDat_take,   iLength);
      VarArrayRedim(UV_vDesc_limsang,    iLength);
      VarArrayRedim(UV_vPatient_limsang, iLength);
      VarArrayRedim(UV_vRequest_Hangmok, iLength);
      VarArrayRedim(UV_vCod_Doctor,      iLength);
   end;
end;


function TfrmLD4Q31.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if mskFROM.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;



procedure TfrmLD4Q31.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbJisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   if copy(GV_sUserId,1,2) = '15' then
   begin
      cmbRelation.Visible := True;
      labRelation.Visible := True;
      CK_Max.Visible      := True;
      Edt_Max.Visible     := True;
   end
   else
   begin
      cmbRelation.Visible := False;
      labRelation.Visible := False;
      CK_Max.Visible      := False;
      Edt_Max.Visible     := False;
   end;
   Edt_Max.Color  := $00E0E0E0;

   // 환경설정
   UP_VarMemSet(0);

   // 일자(From, To)를 오늘일자로 설정
   mskFrom.Text     := GV_sToday;
   mskPandate.Text  := GV_sToday;
   msk_Getdate.Text := GV_sToday;

   ComB_ShEndo.ItemIndex := 1;
   CmbOrder.ItemIndex    := 1;
end;


procedure TfrmLD4Q31.btnQueryClick(Sender: TObject);
var index, iAge : Integer;
    sSelect, sWhere, sOrderBy, sMan, sName : String;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sWhere  := ''; sSelect := ''; sOrderBy := '';
   with qryCell do
   begin
      Active := False;

      // SQL문 생성
      sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn, G.cod_chuga, H.cod_hm, C.cod_hm as Cell_cod_hm, ' +
                 ' G.desc_name, G.num_jumin, C.Desc_buwi, C.Desc_sokun1, C.Desc_sokun2, C.Desc_sokun3, C.Desc_sokun4, C.Desc_sokun5, ' +
                 ' R.*, R.desc_sokun as Desc_limsang ' +
                 ' FROM Ca_Request R Inner join gumgin G ' +
                 ' On R.cod_jisa = G.cod_jisa ' +
                 ' and R.num_id = G.num_id ' +
                 ' and R.dat_gmgn = G.dat_gmgn ' +
                 ' Left outer Join Cell C ' +
                 ' ON R.cod_jisa = C.cod_jisa and R.num_id = C.num_id and R.dat_gmgn = C.dat_gmgn ';
      if ComB_ShEndo.ItemIndex = 1 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P001'' or C.cod_hm = ''P003'' or C.cod_hm = ''P006'' or C.cod_hm = ''P009'' or C.cod_hm = ''P010'' or C.cod_hm = ''P011'') ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P001'' or H.cod_hm = ''P003'' or H.cod_hm = ''P006'' or H.cod_hm = ''P009'' or H.cod_hm = ''P010'' or H.cod_hm = ''P011'') ';
      end
      else if ComB_ShEndo.ItemIndex = 2 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P001'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P001'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 3 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P003'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P003'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 4 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P006'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P006'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 5 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P009'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P009'' ) ';
      end

      else if ComB_ShEndo.ItemIndex = 6 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P010'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P010'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 7 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P011'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P011'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 8 then
      begin
         sSelect := sSelect + ' and (C.cod_hm = ''P012'' ) ';
         sSelect := sSelect + ' left outer join profile_hm H ';
         sSelect := sSelect + ' On  G.cod_jangbi = H.cod_pf ';
         sSelect := sSelect + ' and (H.cod_hm = ''P012'' ) ';
      end;

      if Copy(cmbJisa.Text, 1, 2) <> '00' then
         sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' AND '
      else
         sWhere := ' WHERE ';

      if rbDate.Checked then
         sWhere := sWhere +  ' G.dat_gmgn = ''' + mskFROM.Text + ''''
      else if rbPan_Date.Checked then
         sWhere := sWhere +  ' C.dat_result = ''' + mskPandate.Text + ''''
      else if rd_GetDate.Checked then
         sWhere := sWhere +  ' R.dat_take = ''' + msk_Getdate.Text + ''''
      else if Rd_No.Checked then
      begin
         if (EDTS_NO.Text <> '') and (EDTE_NO.Text <> '') then
            sWhere := sWhere + '  C.desc_buwi >= ''' + Edts_No.Text + ''''
                             + '  AND C.desc_buwi <= ''' + Edte_No.Text + ''''
         else
         begin
            showmessage('접수 No 를 입력해 주세요.');
            Edts_No.SetFocus;
            exit;
         end;
      end;

      if ComB_ShEndo.ItemIndex = 1 then
      begin
         sWhere := sWhere + ' AND ( H.cod_hm <> '''' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P001%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P003%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P006%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P009%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P010%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P011%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P012%'' ) ';
      end
      else if ComB_ShEndo.ItemIndex = 2 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P001'' or G.cod_chuga like ''%P001%'')'
      else if ComB_ShEndo.ItemIndex = 3 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P003'' or G.cod_chuga like ''%P003%'')'
      else if ComB_ShEndo.ItemIndex = 4 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P006'' or G.cod_chuga like ''%P006%'')'
      else if ComB_ShEndo.ItemIndex = 5 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P009'' or G.cod_chuga like ''%P009%'')'
      else if ComB_ShEndo.ItemIndex = 6 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P010'' or G.cod_chuga like ''%P010%'')'
      else if ComB_ShEndo.ItemIndex = 7 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P011'' or G.cod_chuga like ''%P011%'')'
      else if ComB_ShEndo.ItemIndex = 8 then
         sWhere := sWhere + ' AND (H.cod_hm = ''P012'' or G.cod_chuga like ''%P012%'')'
      else
      begin
         sWhere := sWhere + ' AND ((H.cod_hm = ''P001'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P003'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P006'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P009'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P010'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P011'' ';
         sWhere := sWhere + ' or H.cod_hm = ''P012'') ';
         sWhere := sWhere + ' or (G.cod_chuga like ''%P001%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P003%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P006%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P009%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P010%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P011%'' ';
         sWhere := sWhere + ' or G.cod_chuga like ''%P012%'') ) ';
      end;

      if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' order by G.cod_jisa, desc_name'
      else if CmbOrder.ItemIndex = 2 then
         sOrderBy := ' order by G.num_jumin'
      else if CmbOrder.ItemIndex = 3 then
         sOrderBy := ' order by C.Desc_buwi'
      else if CmbOrder.ItemIndex = 4 then
         sOrderBy := ' order by G.cod_jisa, G.gubn_jinch, R.dat_intime'
      else
         sOrderBy := ' order by G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

//      if RecordCount > 0 then
      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD4Q31', 'Q', 'N');
//         UP_VarMemSet(RecordCount-1);

         // 조건에 맞는 데이터 변수에 입력
         while not Eof do
         begin
            iAge := 0; sMan := '';
            GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
            UP_VarMemSet(index);

            UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
            UV_vDesc_name[index]  := FieldByName('desc_name').AsString;

            //2018.04.19 이유경
            if chk_jumin.checked then  UV_vNum_jumin[index]  := copy(FieldByName('num_jumin').AsString, 1, 6) + '-' + copy(FieldByName('num_jumin').AsString, 7, 1) + '******'
            else UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;

            UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
            UV_vNum_id[index]     := FieldByName('Num_id').AsString;
//            UV_vCod_cell[index]   := FieldByName('Cod_cell').AsString;

            if FieldByName('Cell_cod_hm').AsString <> '' then
            begin
               if (FieldByName('Cell_cod_hm').AsString = 'P001') then
                  UV_vCod_hm[index]   := 'P001'
               else if (FieldByName('Cell_cod_hm').AsString = 'P003') then
                  UV_vCod_hm[index]   := 'P003'
               else if (FieldByName('Cell_cod_hm').AsString = 'P006') then
                  UV_vCod_hm[index]   := 'P006'
               else if (FieldByName('Cell_cod_hm').AsString = 'P009') then
                  UV_vCod_hm[index]   := 'P009'
               else if (FieldByName('Cell_cod_hm').AsString = 'P010') then
                  UV_vCod_hm[index]   := 'P010'
               else if (FieldByName('Cell_cod_hm').AsString = 'P011') then
                  UV_vCod_hm[index]   := 'P011';
            end
            else
            begin
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
                  UV_vCod_hm[index]   := 'P011';
            end;

            UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
            //9.체취일
            UV_vDat_take[index]   := FieldByName('Dat_take').AsString;
            //10.임상소견
            UV_vDesc_limsang[index] := FieldByName('Desc_limsang').AsString;
            //11.환자의 임상정보
            if FieldByName('GE_Lmp').AsString = 'Y' then
               UV_vPatient_limsang[index] := 'LMP : ' + FieldByName('GE_Lmp_text').AsString;
            if FieldByName('GE_Menopause').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Menopause : ' + FieldByName('GE_Menopause_text').AsString;
            end;
            if FieldByName('GE_Pregnancy').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Pregnancy : ' + FieldByName('GE_Pregnancy_text').AsString;
            end;
            if FieldByName('GE_IUD').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Pill or IUD';
            end;
            if FieldByName('GE_Erosion').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Erosion ( -, +, ++, +++)';
            end;
            if FieldByName('GE_Hormone').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Hormone therapy';
            end;
            if FieldByName('GE_Radiotherapy').AsString = 'Y' then
            begin
               if trim(UV_vPatient_limsang[index]) <> '' then
                  UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + ', ';
               UV_vPatient_limsang[index] := UV_vPatient_limsang[index] + 'Radiotherapy';
            end;

            //12.의뢰항목
            if (FieldByName('GE_Conventional').AsString = 'Y') or
               (FieldByName('GE_Liquid').AsString = 'Y') then
               UV_vRequest_Hangmok[index] := 'GE';

            if (FieldByName('NGE_Sputum').AsString  = 'Y') or
               (FieldByName('NGE_Pleural').AsString = 'Y') or
               (FieldByName('NGE_Ascitic').AsString = 'Y') or
               (FieldByName('NGE_Joint').AsString   = 'Y') or
               (FieldByName('NGE_Urine').AsString   = 'Y') or
               (FieldByName('NGE_CSF').AsString     = 'Y') or
               (FieldByName('NGE_Other').AsString   = 'Y') or
               (FieldByName('NGE_Cell_block').AsString = 'Y') then
            begin
               if trim(UV_vRequest_Hangmok[index]) <> '' then
                  UV_vRequest_Hangmok[index] := UV_vRequest_Hangmok[index] + ', ';
               UV_vRequest_Hangmok[index] := UV_vRequest_Hangmok[index] + 'NGE';
            end;

            if (FieldByName('FNA_Thyroid').AsString    = 'Y') or
               (FieldByName('FNA_Lymph').AsString      = 'Y') or
               (FieldByName('FNA_Breast').AsString     = 'Y') or
               (FieldByName('FNA_Other').AsString      = 'Y') or
               (FieldByName('FNA_Cell_block').AsString = 'Y') then
            begin
               if trim(UV_vRequest_Hangmok[index]) <> '' then
                  UV_vRequest_Hangmok[index] := UV_vRequest_Hangmok[index] + ', ';
               UV_vRequest_Hangmok[index] := UV_vRequest_Hangmok[index] + 'FNA';
            end;

            //담당의 추가
            with  qryDoctor do
            begin
             Active := False;
             ParamByName('COD_sawon').AsString := qryCell.FieldByName('cod_doctor').AsString;
             Active := True;

             UV_vCod_doctor[index] := qryDoctor.FieldByName('cod_sawon').AsString + '-' + qryDoctor.FieldByName('desc_name').AsString ;
            end;
            qryDoctor.Active := False;
            UV_vAge[index]        := iAge;
            UV_vSex[index]        := sMan;
            UV_vCount[index]      := index + 1;
            Next;
            Inc(index);
         end;
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
      // Table과 Disconnect
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD4Q31.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD4Q31.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCount[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vCod_hm[DataRow-1];
      4 : Value := UV_vDat_gmgn[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vAge[DataRow-1];
      8 : Value := UV_vSex[DataRow-1];
      9 : Value := UV_vDesc_buwi[DataRow-1];
     10 : Value := UV_vDat_take[DataRow-1];
     11 : Value := UV_vDesc_limsang[DataRow-1];
     12 : Value := UV_vPatient_limsang[DataRow-1];
     13 : Value := UV_vRequest_Hangmok[DataRow-1];
     14 : Value := UV_vCod_Doctor[DataRow-1];
   end;
end;


procedure TfrmLD4Q31.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];

end;


procedure TfrmLD4Q31.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnsDate then
      GF_CalendarClick(mskFrom);
   if Sender = btn_Pandate then
      GF_CalendarClick(mskPandate);
   if Sender = btn_Getdate then
      GF_CalendarClick(msk_Getdate);
end;


procedure TfrmLD4Q31.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskFrom        then UP_Click(btnsDate);
end;

procedure TfrmLD4Q31.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q311 := TfrmLD4Q311.Create(Self);
   frmLD4Q311.QuickRep1.Preview;
end;

procedure TfrmLD4Q31.cmbRelationChange(Sender: TObject);
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
  if cmbRelation.ItemIndex = 1 then
  begin
     if Application.MessageBox('접수No. 가 없는 건에 대해서 접수No. 를 일괄부여 하시겠습니까?', '확인', MB_YESNO+MB_ICONquestion) = IDNO then
        Exit;

     iMax := StrToInt(Edt_Max.Text);

     for i := 0 to grdMaster.Rows - 1 do
     begin
        // 자료가 없을경우의 처리
        if grdMaster.Rows = 0 then
        begin
           exit;
        end;

        if (trim(UV_vDesc_buwi[i]) = '') and
           (UV_vCod_hm[i] <> 'B001') and
           (UV_vCod_hm[i] <> 'B007') and
           (UV_vCod_hm[i] <> 'B008') then
        begin
           qryU_Cell.ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[i];
           qryU_Cell.ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[i];
           qryU_Cell.ParamByName('Num_jumin').AsString  := UV_vNum_jumin[i];
           qryU_Cell.ParamByName('DAT_update').AsString := GV_sToday;
           qryU_Cell.ParamByName('COD_update').AsString := GV_sUserId;
           //P003만 적용 최윤선 2018년 5월 15일
           if (UV_vCod_hm[i] = 'P003') then qryU_Cell.ParamByName('dat_time').AsString   := formatdatetime('HH:NN:SS', now)
           else qryU_Cell.ParamByName('dat_time').AsString   := '';

           if CK_Max.Checked then
           begin
              if (Trim(Edt_Max.Text) <> '') and (UV_vCod_hm[i] = 'P001') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end
              else if (Trim(Edt_Max.Text) <> '') and
                      ((UV_vCod_hm[i] = 'P005') or (UV_vCod_hm[i] = 'P009') or
                      (UV_vCod_hm[i] = 'P010') or (UV_vCod_hm[i] = 'P011')) then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'NG' + Copy(UV_vDat_gmgn[i], 4, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
              end
              else if (Trim(Edt_Max.Text) <> '') and (UV_vCod_hm[i] = 'P003') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end;

              iMax := iMax + 1;
           end
           else
           begin
              //'C'로 시작하는건
              qryCa_desc_buwi_Max.Active := False;
              if (UV_vCod_hm[i] = 'P001') then
              begin
                 with qryCa_desc_buwi_Max do
                 begin
                    qryCa_desc_buwi_Max.Active := False;

                    sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ' +
                               ' WHERE desc_buwi > ''C'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                               ' AND   desc_buwi < ''C'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                               ' and substring(desc_buwi,1,3) = ''C'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + '''';
                    qryCa_desc_buwi_Max.SQL.Clear;
                    qryCa_desc_buwi_Max.SQL.Add(sSelect);
                    qryCa_desc_buwi_Max.Active := True;
                 end;
//                 qryCa_desc_buwi_Max.ParamByName('desc_buwi_s').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000000';
//                 qryCa_desc_buwi_Max.ParamByName('desc_buwi_e').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '999999';
              end
              //'NG'로 시작하는건
              else if (UV_vCod_hm[i] = 'P005') or
                      (UV_vCod_hm[i] = 'P009') or
                      (UV_vCod_hm[i] = 'P010') or
                      (UV_vCod_hm[i] = 'P011') then
              begin
                 with qryCa_desc_buwi_Max do
                 begin
                    qryCa_desc_buwi_Max.Active := False;

                    sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,6,6))) result FROM  cell ' +
                               ' WHERE desc_buwi > ''NG'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                               ' AND   desc_buwi < ''NG'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                               ' and substring(desc_buwi,1,4) = ''NG'' + ''' + Copy(UV_vDat_gmgn[i], 4, 2) + '''';
                    qryCa_desc_buwi_Max.SQL.Clear;
                    qryCa_desc_buwi_Max.SQL.Add(sSelect);
                    qryCa_desc_buwi_Max.Active := True;
                 end;
//                 qryCa_desc_buwi_Max.ParamByName('desc_buwi_s').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000000';
//                 qryCa_desc_buwi_Max.ParamByName('desc_buwi_e').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '999999';
              end
              else if (UV_vCod_hm[i] = 'P003') then
              begin
                 with qryCa_desc_buwi_Max do
                 begin
                    qryCa_desc_buwi_Max.Active := False;

                    sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ' +
                               ' WHERE desc_buwi > ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                               ' AND   desc_buwi < ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                               ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + '''';
                    qryCa_desc_buwi_Max.SQL.Clear;
                    qryCa_desc_buwi_Max.SQL.Add(sSelect);
                    qryCa_desc_buwi_Max.Active := True;
                 end;
              end;

              if Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) = '' then
              begin
                 if (UV_vCod_hm[i] = 'P001') then
                 begin
                    qryU_Cell.ParamByName('desc_buwi').AsString := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                    qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                    qryU_Cell.ParamByName('dat_last').AsString    := GV_sToday;
                 end
                 else if (UV_vCod_hm[i] = 'P005') or
                         (UV_vCod_hm[i] = 'P009') or
                         (UV_vCod_hm[i] = 'P010') or
                         (UV_vCod_hm[i] = 'P011') then
                 begin
                    qryU_Cell.ParamByName('desc_buwi').AsString := 'NG' + Copy(UV_vDat_gmgn[i], 4, 2) + '-' + '000001';
                    qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 end
                 else if (UV_vCod_hm[i] = 'P003') then
                 begin
                    qryU_Cell.ParamByName('desc_buwi').AsString := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                    qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                    qryU_Cell.ParamByName('dat_last').AsString    := GV_sToday;
                 end;
              end
              else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                      (UV_vCod_hm[i] = 'P001') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'C' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString    := GV_sToday;
              end
              else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                      ((UV_vCod_hm[i] = 'P005') or (UV_vCod_hm[i] = 'P009') or
                      (UV_vCod_hm[i] = 'P010') or (UV_vCod_hm[i] = 'P011')) then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'NG' + Copy(UV_vDat_gmgn[i], 4, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
              end
              else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                      (UV_vCod_hm[i] = 'P003') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString    := GV_sToday;
              end;
              qryCa_desc_buwi_Max.Active := False;
           end;

           qryU_Cell.ExecSQL;
           GP_query_log(qryU_Cell, 'LD4Q31', 'Q', 'Y');
        end;
     end;
     btnQuery.Click;
//     grdMaster.Repaint;
  end;
end;

procedure TfrmLD4Q31.CK_MaxClick(Sender: TObject);
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
   if CK_Max.Checked = True then
   begin
      Edt_Max.Enabled := False;
      Edt_Max.Color   := clWindow;

     if (ComB_ShEndo.ItemIndex = 2) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';
           if rd_GetDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''C'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''C'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''C'' + ''' + Copy(msk_Getdate.Text, 3, 2) + '''';
           end
           else if rbDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''C'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''C'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''C'' + ''' + Copy(mskFrom.Text, 3, 2) + '''';
           end
           else if rbPan_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''C'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''C'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''C'' + ''' + Copy(mskPandate.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'NG'로 시작하는건
     else if (ComB_ShEndo.ItemIndex = 4) or
             (ComB_ShEndo.ItemIndex = 5) or
             (ComB_ShEndo.ItemIndex = 6) or
             (ComB_ShEndo.ItemIndex = 7) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';
           if rd_GetDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''NG'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''NG'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''NG'' + ''' + Copy(msk_Getdate.Text, 3, 2) + '''';
           end
           else if rbDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''NG'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''NG'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''NG'' + ''' + Copy(mskFrom.Text, 3, 2) + '''';
           end
           else if rbPan_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''NG'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''NG'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''NG'' + ''' + Copy(mskPandate.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'NG'로 시작하는건
     else if (ComB_ShEndo.ItemIndex = 3) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell ';
           if rd_GetDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(msk_Getdate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(msk_Getdate.Text, 3, 2) + '''';
           end
           else if rbDate.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskFrom.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskFrom.Text, 3, 2) + '''';
           end
           else if rbPan_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + '''';
           end
           else if Rd_No.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskPandate.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end;
     Edt_Max.Text := intToStr(qryCa_desc_buwi_Max.FieldByName('RESULT').Asinteger + 1);
     qryCa_desc_buwi_Max.Active := False;
   end
   else if CK_Max.Checked = False then
   begin
      Edt_Max.Enabled := False;
      Edt_Max.Text    := '';
      Edt_Max.Color   := $00E0E0E0;
   end;
end;

procedure TfrmLD4Q31.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.


