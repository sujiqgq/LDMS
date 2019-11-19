//==============================================================================
// 프로그램 설명 : 생화확 ,진단면역 분주 리스트
// 시스템        : 통합검진
// 수정일자      : 2013.5.13
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD4Q44;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges;

type
  TfrmLD4Q44 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    MEdt_SampS: TMaskEdit;
    Label12: TLabel;
    Cmb_gubn: TComboBox;
    mskFrom: TMaskEdit;
    mskTo: TMaskEdit;
    qrygumgin: TQuery;
    qrytkgum: TQuery;
    qryProfile: TQuery;
    qryCheck: TQuery;
    qryProfile_H025: TQuery;
    qryProfile_Check: TQuery;
    qryProfileG: TQuery;
    qryHangmok: TQuery;
    qryNo_hangmok: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    //procedure btnPrint_2Click(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, vCod_chuga : Variant;

    sCode, sHangmok : String;

    sChk_Hm : Boolean;


    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    Procedure Hangmok_Check_Rtn;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q44: TfrmLD4Q44;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q441;

{$R *.DFM}

procedure TfrmLD4Q44.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,    iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
   end;
end;

function TfrmLD4Q44.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q44.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
   Cmb_gubn.ItemIndex := 1;
   edtBDate.Text := dateTostr(date);


end;

procedure TfrmLD4Q44.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_v001[DataRow-1];
      2 : Value := UV_v002[DataRow-1];
      3 : Value := UV_v003[DataRow-1];
      4 : Value := UV_v004[DataRow-1];
      5 : Value := UV_v005[DataRow-1];
      6 : Value := UV_v006[DataRow-1];
      7 : Value := UV_v007[DataRow-1];
   end;
end;

procedure TfrmLD4Q44.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy, sCheck, sJisa : String;

begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := ''; sJisa := '';
   index := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   //[2018.06.23-유동구]항목 조회(필터로 변경)...
   qryHangmok.Active := False;
   qryHangmok.Filter := '';
   qryHangmok.ParamByName('Dat_Apply').AsString := edtBDate.Text;
   qryHangmok.Active := True;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := 'Select B.*,g.num_jumin,g.desc_name,g.num_samp, G.Cod_Jangbi, G.Cod_Hul, G.Cod_Cancer, G.Cod_Chuga,        ' +
                 'G.Gubn_Nosin, G.Gubn_Life, G.Gubn_Adult, G.Gubn_Tkgm, G.Gubn_Bogun, G.Gubn_Agab, G.Gubn_Gong,             ' +
                 'G.Gubn_NsYh,  G.Gubn_LfYh, G.Gubn_AdYh, G.Gubn_BgYh,  G.Gubn_AgYh, G.Gubn_GoYh   ' +
                 'From Bunju B inner join Gumgin g on b.cod_jisa=g.cod_jisa and b.num_id=g.num_id and b.dat_gmgn=g.dat_gmgn ' ;


      sWhere := ' Where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere := sWhere + ' And b.Dat_Bunju = ''' + edtBDate.Text + '''';
      sWhere := sWhere + ' And b.desc_rackno <> ''''' ;

      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
          sOrderBy := ' ORDER BY B.desc_rackno ';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';

         sOrderBy := ' ORDER BY B.desc_rackno ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;
      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q13', 'Q', 'N');
         while not qryBunju.EOF do
         begin
            Gride.Progress := Gride.Progress + 1;
            application.ProcessMessages;

            sChk_Hm := False;

            If FieldByName('Cod_Jisa').AsString = '15' Then
               sChk_Hm := True
            else Hangmok_Check_Rtn; //2018.03.02 채용 자체검사 시행 .. 수지
               {Begin
                if ((qryBunju.FieldByName('gubn_life').AsString <> '1')  AND (qryBunju.FieldByName('gubn_life').AsString <> '2') AND
                    (qryBunju.FieldByName('gubn_adult').AsString <> '1') AND (qryBunju.FieldByName('gubn_adult').AsString <> '2') AND
                    (qryBunju.FieldByName('gubn_nosin').AsString <> '1') AND (qryBunju.FieldByName('gubn_nosin').AsString <> '2')) AND
                   ((Pos('SS', FieldByName('Cod_Jangbi').AsString)> 0) Or (Pos('SG', FieldByName('Cod_Jangbi').AsString)> 0) Or
                    (Pos('SS', FieldByName('Cod_Hul').AsString)> 0 )   Or (Pos('SG', FieldByName('Cod_Hul').AsString)> 0)) Then sChk_Hm := True
                Else  Hangmok_Check_Rtn;
                end;}
                  
            If sChk_Hm = True Then
               Begin
                  UP_VarMemSet(Index);
                  UV_v001[Index] := qryBunju.FieldByName('desc_Rackno').AsString;
                  UV_v002[Index] := qryBunju.FieldByName('cod_jisa').AsString;
                  UV_v003[Index] := qryBunju.FieldByName('num_samp').AsString;
                  UV_v004[Index] := qryBunju.FieldByName('num_bunju').AsString;
                  UV_v005[Index] := qryBunju.FieldByName('Desc_name').AsString;
                  UV_v006[Index] := GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
                  UV_v007[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
                  INC(Index);
               End;
            Next;
         end;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
      Close;
   end;  // with qry_gulkwa do

   qryHangmok.Active := False;
   
   // Grid에 자료를 할당
   grdMaster.Rows := index;
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q44.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate)
   
end;

procedure TfrmLD4Q44.btnPrintClick(Sender: TObject);
begin
  inherited;
  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q441 := TfrmLD4Q441.Create(Self);
     frmLD4Q441.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q441 := TfrmLD4Q441.Create(Self);
     frmLD4Q441.QuickRep.Print;
  end;
end;

procedure TfrmLD4Q44.Hangmok_Check_Rtn;
Var
   sSelect, sWhere, sOrderby : String;
   iTemp, i, iCod_chuga : Integer;
Begin
   sCode := '';
   With qryProfile_Check Do
      Begin
         Active := False;
         ParamByName('Cod_Jangbi').AsString := qryBunju.FieldByName('Cod_Jangbi').AsString;
         ParamByName('Cod_Hul').AsString    := qryBunju.FieldByName('Cod_Hul').AsString;
         ParamByName('Cod_Cancer').AsString := qryBunju.FieldByName('Cod_Cancer').AsString;
         Active := True;
         While Not Eof do
           Begin
              sCode := sCode + FieldByName('Cod_Hm').AsString;
              Next;
           End;
         Active := False;
      End;
      sCode := sCode + qryBunju.FieldByName('COD_CHUGA').AsString;
      if qryBunju.FieldByName('gubn_nosin').AsString = '1' then
         sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '1', qryBunju.FieldByName('gubn_nsyh').AsInteger);
      if (qryBunju.FieldByName('gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('gubn_tkgm').AsString = '2') then
         begin
            qryTkgum.Active := False;
            qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('cod_jisa').AsString;
            qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('num_jumin').AsString;
            qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
            qryTkgum.Active := True;
            if qryTkgum.RecordCount > 0 then
            begin
               //------------------------------------------------------------------------------
               if Trim(qryTkgum.FieldByName('cod_prf').AsString) <> '' then
               begin
                  qryProfileG.Active := False;
                  sSelect  := '';
                  sWhere   := '';
                  sOrderby := '';

                  sSelect := ' SELECT B.cod_hm FROM profile A INNER JOIN profile_hm B ON A.cod_pf = B.cod_pf WHERE ( ';
                  for iTemp := 1 to Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 do
                  begin
                     if iTemp = Length(qryTkgum.FieldByName('cod_prf').AsString) div 4 then
                        sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''''
                     else
                        sWhere := sWhere + ' A.cod_pf = ''' + Copy(qryTkgum.FieldByName('cod_prf').AsString, (iTemp*4)-3, 4) + ''' or ';
                  end;
                  sOrderby := ' ) GROUP BY B.cod_hm ';

                  qryProfileG.SQL.Clear;
                  qryProfileG.SQL.Add(sSelect + sWhere + sOrderby);
                  qryProfileG.Active := True;

                  if qryProfileG.RecordCount > 0 then
                  begin
                     while not qryProfileG.Eof do
                     begin
                        sCode := sCode + qryProfileG.FieldByName('cod_hm').AsString;
                        qryProfileG.Next;
                     end;
                  end;
                  qryProfileG.Active := False;
               end;
               //------------------------------------------------------------------------------
               sCode := sCode + qryTkgum.FieldByName('cod_chuga').AsString;
            end;
            qryTkgum.Active := False;
         end;
         if qryBunju.FieldByName('gubn_bogun').AsString = '1' then
            sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '3', qryBunju.FieldByName('gubn_bgyh').AsInteger);
         if qryBunju.FieldByName('gubn_adult').AsString = '1' then
            sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '4', qryBunju.FieldByName('gubn_adyh').AsInteger);
         if qryBunju.FieldByName('gubn_agab').AsString = '1' then
            sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '5', qryBunju.FieldByName('gubn_agyh').AsInteger);
         if qryBunju.FieldByName('gubn_gong').AsString = '1' then
            sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '6', qryBunju.FieldByName('gubn_goyh').AsInteger);
         if qryBunju.FieldByName('gubn_life').AsString = '1' then
            sCode := sCode + UF_No_Hangmok(qryBunju.FieldByName('dat_gmgn').AsString, '7', qryBunju.FieldByName('gubn_lfyh').AsInteger);
         if Trim(sCode) <> '' then
            begin
               iCod_chuga := GF_MulToSingle(sCode, 4, vCod_chuga);
               for i := 0 to iCod_chuga - 1 do
               begin
                  qryHangmok.Filter := '';
                  qryHangmok.Filter := ' COD_HM = ''' + vCod_chuga[i] + ''' ';

//                  With qryHangmok Do
//                     Begin
//                        Active := False;
//                        ParamByName('Cod_Hm').AsString    := vCod_chuga[i];
//                        ParamByName('Dat_Apply').AsString := qryBunju.FieldByName('Dat_Gmgn').AsString;
//                         Active := True;

                  If (vCod_chuga[i] <> 'C009') And
                     (vCod_chuga[i] <> 'C010') And
                     (vCod_chuga[i] <> 'C011') And
                     (vCod_chuga[i] <> 'C025') And
                     (vCod_chuga[i] <> 'C026') And
                     (vCod_chuga[i] <> 'C027') And
                     (vCod_chuga[i] <> 'C028') And
                     (vCod_chuga[i] <> 'C032') And
                     (vCod_chuga[i] <> 'C042') And
                     (vCod_chuga[i] <> 'C074') AND
                     (vCod_chuga[i] <> 'C083') AND
                     (qryHangmok.RecordCount > 0) Then sChk_Hm := True;

                  If (vCod_chuga[i]  = 'S007') OR
                     (vCod_chuga[i]  = 'S001') OR
                     (vCod_chuga[i]  = 'S008') AND
                     (qryHangmok.RecordCount > 0) Then sChk_Hm := True;
//                     End;
               End;
            End;
End;

function  TfrmLD4Q44.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         if FieldByName('desc_joo').AsString <> '' then
            Result :=  Result + FieldByName('desc_joo').AsString;
      end;
      Active := False;
   end;
end;

end.

