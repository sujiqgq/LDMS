// 수정일자      : 2010.11.28
// 수정자        : 김희석
// 수정내용      : 특검추가
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.30  ~ 2014.05.13  [본원-박연숙]
// 수정자        : 곽윤설
// 수정내용      : 지사별로 조회
// 참고사항      : 센터접수 + '연구소+센터' + 노신,생애,성인 -> 센터검진 (생화학9종)
//==============================================================================
// 수정일자      : 2014.06.20     --> X 삭제
// 수정자        : 곽윤설
// 수정내용      : 생화학 C074 센터 자체검사
// 참고사항      : [본원-최은희]
//==============================================================================

unit LD8Q07;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit,  ComObj, Gauges;

type
  TfrmLD8Q07 = class(TfrmSingle)
    pnlCond: TPanel;
    mskST_date: TDateEdit;
    btnSt_date: TSpeedButton;
    mskEd_date: TDateEdit;
    btnEd_date: TSpeedButton;
    Label10: TLabel;
    Label11: TLabel;
    grdMaster: TtsGrid;
    Label1: TLabel;
    cmbPart: TComboBox;
    qryHangmok: TQuery;
    qryGumgin: TQuery;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    ValEdit1: TValEdit;
    ValEdit2: TValEdit;
    ValEdit3: TValEdit;
    qryProfile: TQuery;
    qryNo_hangmok: TQuery;
    Label5: TLabel;
    Label6: TLabel;
    qryTkgum: TQuery;
    Label7: TLabel;
    Label8: TLabel;
    Bevel2: TBevel;
    Label9: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    QryProfile1: TQuery;
    Label14: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    cmbJisa: TComboBox;
    Label19: TLabel;
    Label20: TLabel;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
 private
     { Private declarations }
    UV_vPart, UV_vCode, UV_vDesc_hm,
    UV_v11, UV_v12, UV_v15, UV_v43, UV_v51, UV_v61, UV_v71, UV_vTot : Variant;

    procedure UP_VarMemSet(iLength : Integer);
    Function  HangmokFind( vTemp , vHangmok : String) : String;
    procedure UP_Insert_2012(iTemp : Integer; sHangmok, sGubn_nosin, sGubn_adult, sGubn_life, sGubn_hulgum, sCod_jisa : String);
    procedure UP_Insert_2013(iTemp : Integer; sHangmok, sCod_jangbi, sCod_hul, sGubn_hulgum, sCod_jisa : String);
  public
    { Public declarations }
  end;


var
  frmLD8Q07 : TfrmLD8Q07;

implementation

uses ZZFunc, ZZComm, MdmsFunc, LD8Q071;

{$R *.DFM}

procedure TfrmLD8Q07.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vPart    := VarArrayCreate([0, 0], varOleStr);
      UV_vCode    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hm := VarArrayCreate([0, 0], varOleStr);
      UV_v11      := VarArrayCreate([0, 0], varOleStr);
      UV_v12      := VarArrayCreate([0, 0], varOleStr);
      UV_v15      := VarArrayCreate([0, 0], varOleStr);
      UV_v43      := VarArrayCreate([0, 0], varOleStr);
      UV_v51      := VarArrayCreate([0, 0], varOleStr);
      UV_v61      := VarArrayCreate([0, 0], varOleStr);
      UV_v71      := VarArrayCreate([0, 0], varOleStr);
      UV_vTot     := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vPart,    iLength);
      VarArrayRedim(UV_vCode,    iLength);
      VarArrayRedim(UV_vDesc_hm, iLength);
      VarArrayRedim(UV_v11,      iLength);
      VarArrayRedim(UV_v12,      iLength);
      VarArrayRedim(UV_v15,      iLength);
      VarArrayRedim(UV_v43,      iLength);
      VarArrayRedim(UV_v51,      iLength);
      VarArrayRedim(UV_v61,      iLength);
      VarArrayRedim(UV_v71,      iLength);
      VarArrayRedim(UV_vTot,     iLength);
   end;
end;
procedure TfrmLD8Q07.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   case DataCol of
      1 : Value := UV_vPart[DataRow-1];
      2 : Value := UV_vCode[DataRow - 1];
      3 : Value := UV_vDesc_hm[DataRow-1];
      4 : Value := UV_v15[DataRow-1];
      5 : Value := UV_v12[DataRow-1];
      6 : Value := UV_v11[DataRow-1];
      7 : Value := UV_v43[DataRow-1];
      8 : Value := UV_v71[DataRow-1];
      9 : Value := UV_v61[DataRow-1];
     10 : Value := UV_v51[DataRow-1];
     11 : Value := UV_vTot[DataRow-1];
   end;
end;

procedure TfrmLD8Q07.FormCreate(Sender: TObject);
begin
   inherited;
   cmbPart.ItemIndex := 0;
   cmbJisa.ItemIndex := 0 ; //20140430 곽윤설
   // Grid 설정 & Memory Allocation
   UP_VarMemSet(0);
   // Grid 초기화
   grdMaster.Rows := 0;
end;

procedure TfrmLD8Q07.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnST_Date then GF_CalendarClick(mskST_Date);
   if Sender = btnED_Date then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD8Q07.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if       Sender = mskST_Date       then UP_Click(mskST_Date)
   else if  Sender = mskED_Date       then UP_Click(mskED_Date);
end;

procedure TfrmLD8Q07.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;
   Try
      Begin
           frmLD8Q071 := TfrmLD8Q071.Create(Self);
           frmLD8Q071.QuickRep.Preview;
      End
   Finally
           frmLD8Q071.Free
   End;
end;

procedure TfrmLD8Q07.btnQueryClick(Sender: TObject);
Var
   SqlWhere, SqlFrom, SqlSelect, SqlOrder : String;
   sSelect, sOderby, sWhere1, sWhere2 : String;

   Index, sPos, I, J, K : Integer;
   wk_11, wk_12, wk_15, wk_33, wk_34, wk_35, wk_41, wk_43, wk_45, wk_46, wk_rec,
   wk_47, wk_51, wk_52, wk_61, wk_63, wk_71, wk_72, wk_gita, wk_tot,iRet,Temp: Integer;
   sValue, sCode, sHangmok : String;
   cod_tk : Variant;
begin
   inherited;
   ValEdit3.Value := 0;
   ValEdit2.Value := 0;
   Index := 0;
   GrdMaster.Rows := 0;
   With qryGumgin do    //분주 조회
   Begin
      Close;

      SqlSelect := ' SELECT G.Cod_jisa, G.dat_gmgn, G.Cod_hul, G.Cod_Cancer, G.Cod_jangbi, G.Cod_chuga, G.Gubn_nosin, G.Gubn_nsyh, ' +
                   ' G.Gubn_adult, G.Gubn_adyh, G.Gubn_agab, G.Gubn_agyh, G.Gubn_life, G.Gubn_lfyh, G.Gubn_hulgum, B.cod_bjjs, B.dat_bunju, B.num_bunju, G.num_jumin, G.gubn_tkgm ';
      SqlFrom   := ' FROM Gumgin G with(nolock) INNER JOIN Bunju B ON G.cod_jisa = B.cod_jisa AND G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn ';

      SqlWhere := ' where B.Dat_bunju >= ''' + mskST_Date.Text + '''';
      SqlWhere := SqlWhere + ' AND B.Dat_bunju <= ''' + mskEd_Date.Text + '''';
      SqlWhere := SqlWhere + ' AND G.ysno_bunju = ''Y''';
      SqlWhere := SqlWhere + ' AND G.gubn_neawon <> ''5''';
{
      if (copy(GV_sUserId,1,2) = '15') then SqlWhere  := 'WHERE (B.cod_bjjs = ''' + copy(GV_sUserId,1,2) + ''')'
      else                                  SqlWhere  := 'WHERE ((B.cod_bjjs = ''' + copy(GV_sUserId,1,2) + ''') or ((B.cod_jisa = ''' + copy(GV_sUserId,1,2) + ''') and (B.cod_bjjs = ''15'') and (G.Gubn_hulgum = ''3''))) ';
}
     if Copy(cmbJisa.Text,1,2) <> '00' then
     begin
         SqlWhere  := SqlWhere + ' and (B.cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''') '
     end;

      Sql.Clear;
      Sql.Add(SqlSelect + SqlFrom + SqlWhere);
      Open;

      wk_rec := 0;

      Valedit1.Value := RecordCount;
      Valedit1.Refresh;

      with qryHangmok Do
      begin
         Close;
         SqlSelect := 'Select Gubn_part, Cod_hm, Desc_kor, Pos_Start ';
         SqlFrom   := 'From Hangmok ';
         if Copy(cmbPart.Text,1,2) = '99' then SqlWhere := 'Where (Gubn_part = ''01'' or Gubn_part = ''02'' or Gubn_part = ''03'' or Gubn_part = ''04'' or Gubn_part = ''05'' or Gubn_part = ''06'' or Gubn_part = ''07'' or Gubn_part = ''08'')'
         else                                  SqlWhere := 'Where Gubn_Part = ''' + Copy(cmbPart.Text,1,2) + '''';
         SqlWhere := SqlWhere + ' Group By Gubn_part, Cod_hm, Desc_kor, Pos_Start';
         SqlORder  := ' Order By Gubn_Part, Cod_hm';

         Sql.Clear;
         Sql.Add(SqlSelect + SqlFrom + SqlWhere + SqlOrder);
         GP_query_log(qryGumgin, 'LD8Q07', 'Q', 'N');
         Open;

         ValEdit2.Value := RecordCount;
         ValEdit2.Refresh;

         UP_VarMemSet(RecordCount-1);

         while Not Eof Do
         begin
            If      FieldByName('Gubn_Part').AsString = '01' Then UV_vPart[Index] := '생화학'
            Else If FieldByname('Gubn_Part').AsString = '02' Then UV_vPart[Index] := '혈액학'
            Else If FieldByname('Gubn_Part').AsString = '03' Then UV_vPart[Index] := 'URIN'
            Else If FieldByname('Gubn_Part').AsString = '04' Then UV_vPart[Index] := 'RIA'
            Else If FieldByname('Gubn_Part').AsString = '05' Then UV_vPart[Index] := 'EIA'
            Else If FieldByname('Gubn_Part').AsString = '06' Then UV_vPart[Index] := '조직학'
            Else If FieldByname('Gubn_Part').AsString = '07' Then UV_vPart[Index] := '혈청학'
            Else If FieldByname('Gubn_Part').AsString = '08' Then UV_vPart[Index] := '혈액기타'
            Else If FieldByname('Gubn_Part').AsString = '09' Then UV_vPart[Index] := '특검';

            UV_vCode[Index]    := FieldByname('Cod_hm').AsString;
            sCode := FieldByname('Cod_hm').AsString;
            UV_vDesc_hm[Index] := FieldByname('Desc_kor').AsString;

            Next;
            Inc(index);
         end; // end of while(qryHangmok);
         Close;
      end;

      K := Index;
      for Index := 0 To K - 1 Do
      begin
         Uv_v15[Index]   := 0;
         Uv_v12[Index]   := 0;
         Uv_v11[Index]   := 0;
         Uv_v43[Index]   := 0;
         Uv_v51[Index]   := 0;
         Uv_v61[Index]   := 0;
         Uv_v71[Index]   := 0;
         Uv_vtot[Index]  := 0;
      end; // end of for(Index);

      While Not Eof Do
      Begin
         sValue := '';
         sSelect := ''; sOderby := ''; sWhere1 := ''; sWhere2 := '';
         with QryProfile1 Do
         Begin
            QryProfile1.Active := False;

            sWhere1 := QryGumgin.FieldByName('Cod_jangbi').AsString + ''',''' + QryGumgin.FieldByName('Cod_hul').AsString + ''',''' + QryGumgin.FieldByName('Cod_Cancer').AsString;

            if Trim(QryGumgin.FieldByName('Cod_chuga').AsString) <> '' then
            begin
               For i := 1 to (length(QryGumgin.FieldByName('Cod_chuga').AsString) div 4) do
               begin
                  if (i = (length(QryGumgin.FieldByName('Cod_chuga').AsString) div 4)) then
                     sWhere2 := sWhere2 + copy(QryGumgin.FieldByName('Cod_chuga').AsString, (i*4)-3, 4)
                  else
                     sWhere2 := sWhere2 + copy(QryGumgin.FieldByName('Cod_chuga').AsString, (i*4)-3, 4) + ''',''';
               end;
            end;

            sSelect := ' ( SELECT P.cod_hm, H.desc_kor, H.gubn_part FROM profile_hm P LEFT OUTER JOIN hangmok H ON P.cod_hm = H.cod_hm AND H.dat_apply <= ''' + mskST_date.Text + '''';
            sSelect := sSelect + ' WHERE P.cod_pf IN (''' + sWhere1 + ''')';
            sSelect := sSelect + ' GROUP BY P.cod_hm, H.desc_kor, H.gubn_part ) UNION ( ';
            sSelect := sSelect + ' SELECT cod_hm, desc_kor, gubn_part  FROM hangmok ';
            sSelect := sSelect + ' WHERE cod_hm IN (''' + sWhere2 + ''')';
            sSelect := sSelect + '   AND dat_apply <= ''' + mskST_date.Text + ''' )';
            sOderby := ' ORDER BY H.gubn_part Desc, P.cod_hm ';

            QryProfile1.SQL.Clear;
            QryProfile1.SQL.Add(sSelect + sOderby);
            QryProfile1.Active := True;

            while Not Eof Do
            begin
               sValue := sValue + QryProfile1.FieldByName('Cod_hm').AsString;
               Next;
            end;
            Close;
         end;

         sValue := sValue;

         If ((FieldByName('gubn_Tkgm').AsString = '1') or (FieldByName('gubn_Tkgm').AsString = '2')) Then
         begin
            qryTkgum.Active := False;
            qryTkgum.ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
            qryTkgum.ParamByName('num_jumin').AsString := qryGumgin.FieldByName('num_jumin').AsString;
            qryTkgum.ParamByName('dat_gmgn').AsString  := qryGumgin.FieldByName('dat_gmgn').AsString;
            qryTkgum.Active := True;

            if qryTkgum.RecordCount > 0 then
            begin
               iRet := GF_MulToSingle(qryTkgum.FieldByName('cod_prf').AsString, 4, cod_tk);

               for Temp := 0 to iRet - 1 do
               begin
                  qryProfile.Active := False;
                  qryProfile.ParamByName('In_Cod_pf').AsString := cod_tk[Temp];
                  qryProfile.Active := True;

                  if qryProfile.RecordCount > 0 then
                  begin
                     while not qryProfile.Eof do
                     begin
                        if pos(qryProfile.FieldByName('Cod_hm').AsString, sValue) < 1 then
                           sValue := sValue + qryProfile.FieldByName('Cod_hm').AsString;

                        qryProfile.next;
                     end;   // while(qryPf_hm) end;
                  end;      //if(qryPf_hm) end;
               end;        //for(qryTkgum) end;

               if qryTkgum.FieldByName('cod_chuga').AsString <> '' then
               begin
                  sHangmok := '';
                  for Temp := 1 to (Length(qryTkgum.FieldByName('cod_chuga').AsString) div 4) do
                  begin
                     sHangmok := copy(qryTkgum.FieldByName('cod_chuga').AsString, (Temp * 4)-3, 4);
                     if pos(sHangmok, sValue) < 1 then sValue := sValue + sHangmok;
                  end;
               end;
            end; //if(qryTkgum) end;
         end;

         // 노신[일반검진]...
         If FieldByName('Gubn_nosin').AsString = '1' Then
         begin
            With qryNo_hangmok Do
            Begin
               Close;
               ParamByName('In_Gubn_Code').AsString := '1';
               ParamByName('In_Gubn_yh').AsString   := QryGumgin.FieldByName('Gubn_nsyh').AsString;
               ParamByName('In_Dat_Apply').AsString := Copy(QryGumgin.FieldByName('dat_gmgn').AsString,1,4);

               Open;
               sValue := HangmokFind(sValue, Trim(FieldByName('Desc_Hul').AsString) +
                                                      Trim(FieldByname('Desc_joo').AsString));
               Close;
            End;
         end;

         // 성인병...
         If FieldByName('Gubn_adult').AsString = '1' Then
         begin
            With qryNo_hangmok Do
            Begin
               Close;
               ParamByName('In_Gubn_Code').AsString := '4';
               ParamByName('In_Gubn_yh').AsString := QryGumgin.FieldByName('Gubn_adyh').AsString;
               ParamByName('In_Dat_Apply').AsString := Copy(QryGumgin.FieldByName('dat_gmgn').AsString,1,4);
               Open;
               sValue := HangmokFind(sValue, Trim(FieldByName('Desc_Hul').AsString) +
                                             Trim(FieldByname('Desc_joo').AsString));
               Close;
            End;
         end;

         // 간염...
         If FieldByName('Gubn_agab').AsString = '1' Then
         begin
            With qryNo_hangmok Do
            Begin
               Close;
               ParamByName('In_Gubn_Code').AsString := '5';
               ParamByName('In_Gubn_yh').AsString := QryGumgin.FieldByName('Gubn_agyh').AsString;
               ParamByName('In_Dat_Apply').AsString := Copy(QryGumgin.FieldByName('dat_gmgn').AsString,1,4);
               Open;
               sValue := HangmokFind(sValue, Trim(FieldByName('Desc_Hul').AsString) + Trim(FieldByname('Desc_joo').AsString));
               Close;
            End;
         end;

         // 생애전환기...
         If FieldByName('gubn_life').AsString = '1' Then
         begin
            With qryNo_hangmok Do
            Begin
               Close;
               ParamByName('In_Gubn_Code').AsString := '7';
               ParamByName('In_Gubn_yh').AsString := QryGumgin.FieldByName('gubn_lfyh').AsString;
               ParamByName('In_Dat_Apply').AsString := Copy(QryGumgin.FieldByName('dat_gmgn').AsString,1,4);
               Open;
               sValue := HangmokFind(sValue, Trim(FieldByName('Desc_Hul').AsString) + Trim(FieldByname('Desc_joo').AsString));
               Close;
            End;
         end;

         // 각각 항목 Count...
         wk_rec := wk_rec + 1;
         ValEdit3.Value  := wk_rec;
         application.ProcessMessages;

         I := Round(Length(sValue) / 4);
         for J := 0 to I - 1 do
            for Index := 0 to K - 1 do
            begin
               sCode := Copy(sValue, J * 4 + 1, 4);
               if sCode = UV_vCode[Index] Then
               begin
                  if (FieldByName('Cod_bjjs').AsString = '15') Then  //곽윤설 20140430
                  begin
                     Uv_vtot[Index] := Uv_vtot[Index] + 1;
                     if      (QryGumgin.FieldByName('cod_jisa').AsString <> '12') and (QryGumgin.FieldByName('dat_gmgn').AsString < '20130805') then
                             UP_Insert_2012(Index, sCode, QryGumgin.FieldByName('gubn_nosin').AsString, QryGumgin.FieldByName('Gubn_adult').AsString,
                                                          QryGumgin.FieldByName('gubn_life').AsString,  QryGumgin.FieldByName('Gubn_hulgum').AsString,
                                                          QryGumgin.FieldByName('cod_jisa').AsString)
                     else if (QryGumgin.FieldByName('cod_jisa').AsString = '12') and (QryGumgin.FieldByName('dat_gmgn').AsString < '20130729') then
                             UP_Insert_2012(Index, sCode, QryGumgin.FieldByName('gubn_nosin').AsString, QryGumgin.FieldByName('Gubn_adult').AsString,
                                                          QryGumgin.FieldByName('gubn_life').AsString,  QryGumgin.FieldByName('Gubn_hulgum').AsString,
                                                          QryGumgin.FieldByName('cod_jisa').AsString)
                     else if (QryGumgin.FieldByName('dat_gmgn').AsString > '20140331') then //20140430 곽윤설   //20140704
                             UP_Insert_2012(Index, sCode, QryGumgin.FieldByName('gubn_nosin').AsString, QryGumgin.FieldByName('Gubn_adult').AsString,
                                                          QryGumgin.FieldByName('gubn_life').AsString,  QryGumgin.FieldByName('Gubn_hulgum').AsString,
                                                          QryGumgin.FieldByName('cod_jisa').AsString)
                     else    UP_Insert_2013(Index, sCode, QryGumgin.FieldByName('cod_jangbi').AsString,  QryGumgin.FieldByName('cod_hul').AsString,
                                                          QryGumgin.FieldByName('Gubn_hulgum').AsString, QryGumgin.FieldByName('cod_jisa').AsString)
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '12') AND ((Copy(cmbJisa.Text,1,2)='12') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v12[Index] := Uv_v12[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '11') AND ((Copy(cmbJisa.Text,1,2)='11') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v11[Index] := Uv_v11[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '43') AND ((Copy(cmbJisa.Text,1,2)='43') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v43[Index] := Uv_v43[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '71') AND ((Copy(cmbJisa.Text,1,2)='71') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v71[Index] := Uv_v71[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '61') AND ((Copy(cmbJisa.Text,1,2)='61') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v61[Index] := Uv_v61[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (FieldByName('Cod_bjjs').AsString = '51') AND ((Copy(cmbJisa.Text,1,2)='51') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v51[Index] := Uv_v51[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end;
               End;
            end;
         Next;
      End;
      Close;
   end; // end of With(qryGumgin);
   GrdMaster.Rows := index;
   GrdMaster.SetFocus;
end;

Function TfrmLD8Q07.HangmokFind( vTemp , vHangmok : String) : String;
var
   iRet, i, Temp1 : Integer;
   vCod_totyh : Variant;
   ckeck1 : Boolean;
   sHangmokFind, sHangmok : String;
begin
   ckeck1 := False;
   sHangmokFind := ''; sHangmok := ''; Result := '';
   iRet   := GF_MulToSingle(vHangmok, 4, vCod_totyh);
   for i := 0 to iRet-1 do
   begin
      if copy(vCod_totyh[i],1,2) <> 'JJ' then
      begin
         ckeck1 := True;
         for temp1 := 1 to (Length(vTemp) div 4) do
         begin
            sHangmok := copy(vTemp, (temp1 * 4)-3, 4);
            if sHangmok = vCod_totyh[i] then ckeck1 := False;
         end;
         if ckeck1 = True then sHangmokFind := sHangmokFind + vCod_totyh[i];
      end;
   end;
   Result := vTemp + sHangmokFind;
end;

procedure TfrmLD8Q07.UP_Insert_2012(iTemp : Integer; sHangmok, sGubn_nosin, sGubn_adult, sGubn_life, sGubn_hulgum, sCod_Jisa : String);
begin
   //[2012.04.03]본원센터 통계시 필요....
   //------------------------------------------------------------
   //20160817 수원센터 자체 검사 진행으로 노신에 상관없이 생화학 9종은 수원으로 카운트.. 수지
   if ((sHangmok = 'C009') or (sHangmok = 'C010') or (sHangmok = 'C011') or (sHangmok = 'C025') or (sHangmok = 'C026') or
       (sHangmok = 'C027') or (sHangmok = 'C028') or (sHangmok = 'C032') or (sHangmok = 'C042')) and
      (QryGumgin.FieldByName('cod_jisa').AsString = '43') and (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1;
   end
   else if ((sHangmok = 'C009') or (sHangmok = 'C010') or (sHangmok = 'C011') or (sHangmok = 'C025') or (sHangmok = 'C026') or
       (sHangmok = 'C027') or (sHangmok = 'C028') or (sHangmok = 'C032') or (sHangmok = 'C042')) and
      ((sGubn_nosin = '1') or (sGubn_adult = '1') or (sGubn_life = '1')) and
      (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else if ((sHangmok = 'H001') or (sHangmok = 'H002') or (sHangmok = 'H003') or (sHangmok = 'H004') or (sHangmok = 'H005') or
       (sHangmok = 'H006') or (sHangmok = 'H007') or (sHangmok = 'H008') or (sHangmok = 'H009') or (sHangmok = 'H010') or
       (sHangmok = 'H011') or (sHangmok = 'H012') or (sHangmok = 'H013') or (sHangmok = 'H014') or (sHangmok = 'H015') or
       (sHangmok = 'H016') or (sHangmok = 'H017') or (sHangmok = 'H018') or (sHangmok = 'H019') or (sHangmok = 'H020') or
       (sHangmok = 'H021') or (sHangmok = 'H022') or (sHangmok = 'H023') or (sHangmok = 'H024') or (sHangmok = 'H025') or (sHangmok = 'H035')) and
      (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
      else if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else if ((sHangmok = 'U001') or (sHangmok = 'U002') or (sHangmok = 'U003') or (sHangmok = 'U004') or (sHangmok = 'U005') or
            (sHangmok = 'U006') or (sHangmok = 'U007') or (sHangmok = 'U008') or (sHangmok = 'U009') or (sHangmok = 'U010') or
            (sHangmok = 'U011') or (sHangmok = 'U012') or (sHangmok = 'U013') or (sHangmok = 'U053') or (sHangmok = 'U054') or
            (sHangmok = 'U055') or (sHangmok = 'Y004')) and
           (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
      else if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else
   begin
      Uv_v15[iTemp]   := Uv_v15[iTemp] + 1;
   end;
end;

procedure TfrmLD8Q07.UP_Insert_2013(iTemp : Integer; sHangmok, sCod_jangbi, sCod_hul, sGubn_hulgum, sCod_Jisa : String);
begin
   //[2013.08.02]본원센터 통계시 필요....
   //------------------------------------------------------------
   if ((sHangmok = 'C001') or (sHangmok = 'C002') or (sHangmok = 'C003') or (sHangmok = 'C004') or (sHangmok = 'C005') or
       (sHangmok = 'C006') or (sHangmok = 'C007') or (sHangmok = 'C009') or (sHangmok = 'C010') or (sHangmok = 'C011') or
       (sHangmok = 'C012') or (sHangmok = 'C013') or (sHangmok = 'C017') or (sHangmok = 'C019') or (sHangmok = 'C021') or
       (sHangmok = 'C025') or (sHangmok = 'C026') or (sHangmok = 'C027') or (sHangmok = 'C028') or (sHangmok = 'C029') or
       (sHangmok = 'C032') or (sHangmok = 'C033') or (sHangmok = 'C041') or (sHangmok = 'C042') or (sHangmok = 'C043')) and
       (QryGumgin.FieldByName('cod_jisa').AsString = '43') and (sGubn_hulgum = '3') then
   begin
      if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1 ;
   end
   else if ((sHangmok = 'C001') or (sHangmok = 'C002') or (sHangmok = 'C003') or (sHangmok = 'C004') or (sHangmok = 'C005') or
       (sHangmok = 'C006') or (sHangmok = 'C007') or (sHangmok = 'C009') or (sHangmok = 'C010') or (sHangmok = 'C011') or
       (sHangmok = 'C012') or (sHangmok = 'C013') or (sHangmok = 'C017') or (sHangmok = 'C019') or (sHangmok = 'C021') or
       (sHangmok = 'C025') or (sHangmok = 'C026') or (sHangmok = 'C027') or (sHangmok = 'C028') or (sHangmok = 'C029') or
       (sHangmok = 'C032') or (sHangmok = 'C033') or (sHangmok = 'C041') or (sHangmok = 'C042') or (sHangmok = 'C043')) and
      (copy(sCod_jangbi,1,2) <> 'SS') and (copy(sCod_hul,1,2) <> 'SS') and
      (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
    //  else if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else if ((sHangmok = 'H001') or (sHangmok = 'H002') or (sHangmok = 'H003') or (sHangmok = 'H004') or (sHangmok = 'H005') or
            (sHangmok = 'H006') or (sHangmok = 'H007') or (sHangmok = 'H008') or (sHangmok = 'H009') or (sHangmok = 'H010') or
            (sHangmok = 'H011') or (sHangmok = 'H012') or (sHangmok = 'H013') or (sHangmok = 'H014') or (sHangmok = 'H015') or
            (sHangmok = 'H016') or (sHangmok = 'H017') or (sHangmok = 'H018') or (sHangmok = 'H019') or (sHangmok = 'H020') or
            (sHangmok = 'H021') or (sHangmok = 'H022') or (sHangmok = 'H023') or (sHangmok = 'H024') or (sHangmok = 'H025') or (sHangmok = 'H035')) and
           (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
      else if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else if ((sHangmok = 'U001') or (sHangmok = 'U002') or (sHangmok = 'U003') or (sHangmok = 'U004') or (sHangmok = 'U005') or
            (sHangmok = 'U006') or (sHangmok = 'U007') or (sHangmok = 'U008') or (sHangmok = 'U009') or (sHangmok = 'U010') or
            (sHangmok = 'U011') or (sHangmok = 'U012') or (sHangmok = 'U013') or (sHangmok = 'U053') or (sHangmok = 'U054') or
            (sHangmok = 'U055') or (sHangmok = 'Y004')) and
           (sGubn_hulgum = '3') then
   begin
      if      sCod_jisa = '15' Then Uv_v15[iTemp]   := Uv_v15[iTemp] + 1
      else if sCod_jisa = '12' Then Uv_v12[iTemp]   := Uv_v12[iTemp] + 1
      else if sCod_jisa = '11' Then Uv_v11[iTemp]   := Uv_v11[iTemp] + 1
      else if sCod_jisa = '43' Then Uv_v43[iTemp]   := Uv_v43[iTemp] + 1
      else if sCod_jisa = '71' Then Uv_v71[iTemp]   := Uv_v71[iTemp] + 1
      else if sCod_jisa = '61' Then Uv_v61[iTemp]   := Uv_v61[iTemp] + 1
      else if sCod_jisa = '51' Then Uv_v51[iTemp]   := Uv_v51[iTemp] + 1;
   end
   else
   begin
      Uv_v15[iTemp]   := Uv_v15[iTemp] + 1;
   end;
end;

procedure TfrmLD8Q07.SBut_ExcelClick(Sender: TObject);
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
