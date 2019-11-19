//==============================================================================
// 프로그램 설명 : S008,S091 전결과 비교LIST
// 시스템        : 분석관리
// 수정일자      : 2012.11.20
// 수정자        : 김희석
// 참고사항      : 추가..
//==============================================================================
unit LD4Q43;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj;

type
  TfrmLD4Q43 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    qryGulkwa: TQuery;
    Gride: TGauge;
    qryPGulkwa: TQuery;
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
    Label1: TLabel;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryProfileG: TQuery;
    qryPf_hm: TQuery;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    gubn_Combo: TComboBox;
    Label3: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
      procedure UP_GridSet;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vNum_samp, UV_vNum_bunju, UV_vpDat_gmgn,UV_vpNum_bunju, UV_vpS008, UV_vS008,UV_vDesc_name : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_hangmokList : String;
    function  UF_PhangmokList : String;
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q43: TfrmLD4Q43;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q431;

{$R *.DFM}

procedure TfrmLD4Q43.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju := VarArrayCreate([0, 0], varOleStr);
      UV_vpDat_gmgn := VarArrayCreate([0, 0], varOleStr);
      UV_vS008      := VarArrayCreate([0, 0], varOleStr);
      UV_vpS008      := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name := VarArrayCreate([0, 0], varOleStr);
      UV_vpNum_bunju:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,        iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vNum_bunju, iLength);
      VarArrayRedim(UV_vpDat_gmgn, iLength);
      VarArrayRedim(UV_vS008,      iLength);
      VarArrayRedim(UV_vpS008,      iLength);
      VarArrayRedim(UV_vDesc_name,      iLength);
      VarArrayRedim(UV_vpNum_bunju,      iLength);

   end;
end;

function TfrmLD4Q43.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q43.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   UP_GridSet;
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   gubn_Combo.ItemIndex :=0;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 0;
end;

procedure TfrmLD4Q43.UP_GridSet;
begin
   // Grid의 환경 설정
   with grdMaster do
   begin
      Rows := 0;

      // Column Title
      if  (gubn_Combo.ItemIndex =1 ) then
      begin
         Col[7].Heading := 'S008(전)';
         Col[8].Heading := 'S008(현)';

      end
      else if (gubn_Combo.ItemIndex =2 ) then
      begin
         Col[7].Heading := 'S007(전)';
         Col[8].Heading := 'S007(현)';
      end
      else if (gubn_Combo.ItemIndex =3 ) then
      begin
         Col[7].Heading := 'S091(전)';
         Col[8].Heading := 'S091(현)';
      end
      else if (gubn_Combo.ItemIndex =4 ) then
      begin
         Col[7].Heading := 'SE31(전)';
         Col[8].Heading := 'SE31(현)';
      end
      else if (gubn_Combo.ItemIndex =5 ) then
      begin
         Col[7].Heading := '임신';
         Col[8].Heading := '';
      end
      else if (gubn_Combo.ItemIndex =6 ) then
      begin
         Col[7].Heading := 'C047(현)';
         Col[8].Heading := '';
      end
      else if (gubn_Combo.ItemIndex =7 ) then
      begin
         Col[7].Heading := 'H011(현)';
         Col[8].Heading := '';
      end;

   end;
end;

procedure TfrmLD4Q43.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vNum_bunju[DataRow-1];
      4 : Value := UV_vDesc_name[DataRow-1];
      5 : Value := UV_vpDat_gmgn[DataRow-1];
      6 : Value := UV_vpNum_bunju[DataRow-1];
      7 : Value := UV_vpS008[DataRow-1];
      8 : Value := UV_vS008[DataRow-1];
   end;


end;

procedure TfrmLD4Q43.btnQueryClick(Sender: TObject);
var index,iAGe : Integer;
    sSelect, sWhere, sOrderBy, sHangmokList, spHangmokList ,Temp,pTemp,sMan: String;
    bCheck : Boolean;
    Temp_s091,Temp_ps091 : Extended;
begin
   inherited;
    sSelect := ''; sWhere := ''; sOrderBy := '';
   // 조회조건 Check
   if not UF_Condition then exit;
   UP_GridSet;
   // Grid 초기화
   grdMaster.Rows := 0;
   with qryGulkwa do
   begin
      // SQL문 생성
      Close;
      if gubn_Combo.ItemIndex <> 5 then
          begin
          sSelect := ' select A.num_jumin,A.Dat_gmgn,A.desc_name,G.cod_jisa, G.num_id, G.dat_gmgn, A.num_samp, G.desc_glkwa1,G.desc_glkwa2,G.DESC_GLKWA3, G.num_bunju, G.gubn_part, ' +
                     ' A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.num_jumin, ' +
                     ' A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga ';
          
          sSelect := sSelect + ' From gulkwa G   left outer join Gumgin A ';
          sSelect := sSelect + ' On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';

          sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';

          if  gubn_Combo.ItemIndex = 4 then
          begin
          sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                             ' AND  G.Gubn_Part = ''05''';
          end
          else if  (gubn_Combo.ItemIndex = 6) or (gubn_Combo.ItemIndex = 9) then
          begin
          sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                             ' AND  G.Gubn_Part = ''01''';
          end
          else if  (gubn_Combo.ItemIndex = 7) or (gubn_Combo.ItemIndex = 8) then
          begin
          sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                             ' AND  G.Gubn_Part = ''02''';
          end
          else
          begin
          sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                             ' AND  G.Gubn_Part = ''07''';
          end;

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
          
             sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), G.num_bunju '
          end;
      end

      else if gubn_Combo.ItemIndex =5 then
          begin
          sSelect := ' select A.Gubn_preg,A.desc_name,G.cod_jisa, G.num_id, G.dat_gmgn, A.num_samp,  G.num_bunju,  ' +
                     ' A.gubn_nosin, A.gubn_nsyh, A.Gubn_adult, A.Gubn_adyh, A.Gubn_life, A.Gubn_lfyh, A.num_jumin, ' +
                     ' A.Gubn_bogun, A.Gubn_bgyh, A.Gubn_agab, A.Gubn_agyh, A.Gubn_tkgm, A.cod_hul, A.cod_jangbi, A.cod_cancer, A.cod_chuga ';
          
          sSelect := sSelect + ' From bunju G   left outer join Gumgin A ';
          sSelect := sSelect + ' On a.cod_jisa = G.cod_jisa and a.num_id = G.num_id and a.dat_gmgn = G.dat_gmgn ';
          
          sWhere := ' WHERE G.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
          sWhere := sWhere + ' AND  G.Dat_Bunju = ''' + mskDate.Text + ''' ' +
                             ' AND  A.Gubn_preg <> ''N''';
          
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
          
             sOrderBy := ' ORDER BY CAST(A.num_samp AS INT), G.num_bunju '
          end;
      end;          

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGulkwa, 'LD4Q43', 'Q', 'N');
         while Not Eof do
         begin
            Gride.Progress := Gride.Progress + 1;

            sHangmokList := '';
            sHangmokList := UF_hangmokList;
            temp:='';
            pTemp:='';
            Temp_pS091:=0;
            Temp_S091:=0;


            if gubn_Combo.ItemIndex =5 then
            begin
               UP_VarMemSet(index);
               UV_vNo[index]        := intTostr(index+1);
               UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
               UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;

               if qryGulkwa.FieldByName('Gubn_preg').AsString ='Y' then
                  UV_vpS008[index] :='임신'
               else if qryGulkwa.FieldByName('Gubn_preg').AsString ='P' then
                  UV_vpS008[index] :='임신가능성(유)'
               else UV_vpS008[index] :='';
               Inc(index);
            end
          else if (pos('C047',sHangmokList) > 0) and (gubn_Combo.ItemIndex =6) then
          begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

               temp:= Trim(Copy(UV_fGulkwa,  211,  6));
               if temp <> '' then
                  Temp_s091:=StrTofloat(Temp);
               if Temp_s091 >=90 then
               begin
               UP_VarMemSet(index);
               UV_vNo[index]        := intTostr(index+1);
               UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
               UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
               UV_vpS008[index] :=Trim(Copy(UV_fGulkwa,  211,  6));

               Inc(index);
               end
            end
            else if (pos('H001',sHangmokList) > 0) and (gubn_Combo.ItemIndex =8) then
            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
               GP_JuminToAge(qryGulkwa.FieldByName('Num_Jumin').AsString,qryGulkwa.FieldByName('dat_gmgn').AsString, sMan, iAge);
               temp:= Trim(Copy(UV_fGulkwa,  1,  6));
               if temp <> '' then
                  Temp_s091:=StrTofloat(Temp);
               if ((Temp_s091 >=7.8)and (sMan='M')) or
                  ((Temp_s091 >=7.2)and (sMan='F')) Then
               begin
               UP_VarMemSet(index);
               UV_vNo[index]        := intTostr(index+1);
               UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
               UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
               UV_vpS008[index] :=Trim(Copy(UV_fGulkwa,  1,  6));

               Inc(index);
               end
            end
            else if (pos('C083',sHangmokList) > 0) and (gubn_Combo.ItemIndex =9) then
            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

               temp:= Trim(Copy(UV_fGulkwa,  385,  6));
               if temp <> '' then
                  Temp_s091:=StrTofloat(Temp);
               if (Temp_s091 <=3.4) and (Temp_s091 >=0.001)  then
               begin
               UP_VarMemSet(index);
               UV_vNo[index]        := intTostr(index+1);
               UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
               UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
               UV_vpS008[index] :=Trim(Copy(UV_fGulkwa,  385,  6));

               Inc(index);
               end
            end
            else if (pos('H011',sHangmokList) > 0) and (gubn_Combo.ItemIndex =7) then
            begin
               UV_fGulkwa := '';
               UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
               UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
               UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
               GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

               temp:= Trim(Copy(UV_fGulkwa,  61,  6));
               if temp <> '' then
                  Temp_s091:=StrTofloat(Temp);
               if Temp_s091 >=20 then
               begin
               UP_VarMemSet(index);
               UV_vNo[index]        := intTostr(index+1);
               UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
               UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
               UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
               UV_vpS008[index] :=Trim(Copy(UV_fGulkwa,  61,  6));

               Inc(index);
               end
            end
            else if (pos('S008',sHangmokList) > 0) and (gubn_Combo.ItemIndex =1) then
            begin
               qryPGulkwa.close;
               qryPGulkwa.ParamByName('num_id').AsString   := qryGulkwa.FieldByName('num_id').AsString;
               qryPGulkwa.ParamByName('dat_bunju').AsString := mskDate.text;
               qryPGulkwa.ParamByName('gubn_part').Asinteger :=7;
               qryPGulkwa.open;

               if qryPGulkwa.RecordCount > 0 then
               begin
                  bCheck := True;
                  while not qryPGulkwa.Eof do
                  begin
                     spHangmokList := '';
                     spHangmokList := UF_PhangmokList;
                     if (pos('S008',spHangmokList) > 0) and (bCheck = True) then
                     begin
                     if ((trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6))='양성') or
                         (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6))='약양성'))    and
                        (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6))='음성') and (bCheck) then
                        begin
                        bCheck := False;
                        UP_VarMemSet(index);
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_bunju').AsString;
                        UV_vpNum_bunju[index]:= qryPGulkwa.FieldByName('Num_bunju').AsString;
                        UV_vpS008[index]     := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                        UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,37,6));
                        Inc(index);
                        end;
                     end;
                     qryPGulkwa.Next;
                  end;
                  qryPGulkwa.close;
               end;
            end
            else if (pos('SE31',sHangmokList) > 0) and  (gubn_Combo.ItemIndex =4) then
            begin
               qryPGulkwa.close;
               qryPGulkwa.ParamByName('num_id').AsString   := qryGulkwa.FieldByName('num_id').AsString;
               qryPGulkwa.ParamByName('dat_bunju').AsString := mskDate.text;
               qryPGulkwa.ParamByName('gubn_part').Asinteger :=5;
               qryPGulkwa.open;

               if qryPGulkwa.RecordCount > 0 then
               begin
                  bCheck := True;
                  while not qryPGulkwa.Eof do
                  begin
                     spHangmokList := '';
                     spHangmokList := UF_PhangmokList;
                     if (pos('SE31',spHangmokList) > 0) and (bCheck = True) then
                     begin  
                     if (strToint(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6))>=20) and
                        (strToint(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6))<20) and (bCheck) then
                        begin
                        bCheck := False;
                        UP_VarMemSet(index);
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_bunju').AsString;
                        UV_vpNum_bunju[index]:= qryPGulkwa.FieldByName('Num_bunju').AsString;
                        UV_vpS008[index]     := trim(copy(qryPGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                        UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa3').AsString,45,6));
                        Inc(index);
                        end;
                     end;
                     qryPGulkwa.Next;
                  end;
                  qryPGulkwa.close;
               end;
            end

            else if (pos('S007',sHangmokList) > 0) and  (gubn_Combo.ItemIndex =2) then
            begin
               qryPGulkwa.close;
               qryPGulkwa.ParamByName('num_id').AsString   := qryGulkwa.FieldByName('num_id').AsString;
               qryPGulkwa.ParamByName('dat_bunju').AsString := mskDate.text;
               qryPGulkwa.ParamByName('gubn_part').Asinteger :=7;
               qryPGulkwa.open;
               if qryPGulkwa.RecordCount > 0 then
               begin
                  bCheck := True;
                  while not qryPGulkwa.Eof do
                  begin
                     spHangmokList := '';
                     spHangmokList := UF_PhangmokList;
                     if (pos('S008',spHangmokList) > 0) and (bCheck = True) then
                     begin
                     if ((trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))='양성') or
                        (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6))='약양성')) and
                        (trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6))='음성') and (bCheck) then
                        begin
                        bCheck := False;
                        UP_VarMemSet(index);
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_bunju').AsString;
                        UV_vpNum_bunju[index]:= qryPGulkwa.FieldByName('Num_bunju').AsString;
                        UV_vpS008[index]     := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                        UV_vS008[index]      := trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,31,6));
                        Inc(index);
                        end;
                     end;
                     qryPGulkwa.Next;
                  end;
                  qryPGulkwa.close;
               end;
            end

            else if (pos('S091',sHangmokList) > 0) and  (gubn_Combo.ItemIndex =3)  then
            begin
               qryPGulkwa.close;
               qryPGulkwa.ParamByName('num_id').AsString   := qryGulkwa.FieldByName('num_id').AsString;
               qryPGulkwa.ParamByName('dat_bunju').AsString := mskDate.text;
               qryPGulkwa.ParamByName('gubn_part').Asinteger :=7;
               qryPGulkwa.open;

               if qryPGulkwa.RecordCount > 0 then
               begin
                  bCheck := True;
                  while not qryPGulkwa.Eof do
                  begin

                     spHangmokList := '';
                     spHangmokList := UF_PhangmokList;

                     if (pos('S091',spHangmokList) > 0) and (bCheck = True)   then
                     begin
                     Temp:= trim(copy(qryGulkwa.FieldByName('desc_glkwa1').AsString,181,6));
                     pTemp:=trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,181,6));
                     if temp <> '' then
                        Temp_S091:=StrToFloat(Temp);
                     if pTemp <> '' then
                        Temp_pS091:=StrToFloat(pTemp);


                     if (Temp_pS091>=10)and ((Temp_S091<10) and (Temp_S091>0)) and (bCheck) then
                     begin
                        bCheck := False;
                        UP_VarMemSet(index);
                        UV_vNo[index]        := intTostr(index+1);
                        UV_vNum_samp[index]  := qryGulkwa.FieldByName('num_samp').AsString;
                        UV_vNum_bunju[index] := qryGulkwa.FieldByName('num_bunju').AsString;
                        UV_vpDat_gmgn[index] := qryPGulkwa.FieldByName('dat_bunju').AsString;
                        UV_vDesc_name[index] := qryGulkwa.FieldByName('Desc_name').AsString;
                        UV_vpNum_bunju[index]:= qryPGulkwa.FieldByName('Num_bunju').AsString;                                             
                        UV_vpS008[index]     := pTemp;
                        UV_vS008[index]      := Temp;
                        Inc(index);
                     end;
                     end;
                     qryPGulkwa.Next;
                  end;

                  qryPGulkwa.close;
               end;
            end;
            next;
         end;
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

procedure TfrmLD4Q43.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;


function TfrmLD4Q43.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i : integer;
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
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;


   result := sTemp;
end;

function TfrmLD4Q43.UF_PhangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryPGulkwa.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPGulkwa.FieldByName('cod_hul').AsString;
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
      if qryPGulkwa.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPGulkwa.FieldByName('Cod_cancer').AsString;
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
      if qryPGulkwa.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryPGulkwa.FieldByName('Cod_jangbi').AsString;
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
   sTemp := sTemp + qryPGulkwa.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryPGulkwa.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryPGulkwa.FieldByName('Dat_gmgn').AsString, '1', qryPGulkwa.FieldByName('Gubn_nsyh').AsInteger)
   else if qryPGulkwa.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryPGulkwa.FieldByName('Dat_gmgn').AsString, '4', qryPGulkwa.FieldByName('Gubn_adyh').AsInteger);

   if qryPGulkwa.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPGulkwa.FieldByName('Dat_gmgn').AsString, '7', qryPGulkwa.FieldByName('Gubn_lfyh').AsInteger);

   if qryPGulkwa.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPGulkwa.FieldByName('Dat_gmgn').AsString, '3',qryPGulkwa.FieldByName('Gubn_bgyh').AsInteger);

   if qryPGulkwa.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryPGulkwa.FieldByName('Dat_gmgn').AsString, '5',qryPGulkwa.FieldByName('Gubn_agyh').AsInteger);

   if (qryPGulkwa.FieldByName('Gubn_tkgm').AsString = '1') or (qryPGulkwa.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryPGulkwa.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryPGulkwa.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryPGulkwa.FieldByName('Dat_gmgn').AsString;
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
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;  

   result := sTemp;
end;

function TfrmLD4Q43.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD4Q43.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
    begin
     frmLD4Q431 := TfrmLD4Q431.Create(Self);
     frmLD4Q431.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q431 := TfrmLD4Q431.Create(Self);
     frmLD4Q431.QuickRep.Print;
  end;

end;

procedure TfrmLD4Q43.SBut_ExcelClick(Sender: TObject);
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
