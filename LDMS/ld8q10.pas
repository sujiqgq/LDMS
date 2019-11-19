// 수정일자      : 2018-04-21
// 수정자        : 박수지
// 수정내용      : 질관리 관련 프로그램 생성요청
// 참고사항      :
//==============================================================================


unit LD8Q10;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit,  ComObj, Gauges;

type
  TfrmLD8Q10 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryHangmok: TQuery;
    qryGumgin: TQuery;
    qryProfile: TQuery;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    QryProfile1: TQuery;
    Panel2: TPanel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Bevel3: TBevel;
    Label25: TLabel;
    Label26: TLabel;
    Gauge2: TGauge;
    Bevel4: TBevel;
    Bevel5: TBevel;
    cmbJisa: TComboBox;
    StaticText2: TStaticText;
    cmbPart: TComboBox;
    ValEdit1: TValEdit;
    ValEdit3: TValEdit;
    ValEdit2: TValEdit;
    SBut_Excel: TSpeedButton;
    mskST_date: TDateEdit;
    mskEd_date: TDateEdit;
    btnSt_date: TSpeedButton;
    btnEd_date: TSpeedButton;
    QRY_TEMP: TQuery;
    Edts_No: TEdit;
    Label17: TLabel;
    Edte_No: TEdit;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnPrintClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
 private
     { Private declarations }
    UV_vPart, UV_vCode, UV_vDesc_hm,
    UV_v11, UV_v12, UV_v15, UV_v43, UV_v51, UV_v61, UV_v71, UV_vTot : Variant;

    procedure UP_VarMemSet(iLength : Integer);
    Function  HangmokFind( vTemp , vHangmok : String) : String;
    //procedure UP_Insert_2012(iTemp : Integer; sHangmok, sGubn_nosin, sGubn_adult, sGubn_life, sGubn_hulgum, sCod_jisa : String);
   // procedure UP_Insert_2013(iTemp : Integer; sHangmok, sCod_jangbi, sCod_hul, sGubn_hulgum, sCod_jisa : String);
  public
    { Public declarations }
  end;


var
  frmLD8Q10 : TfrmLD8Q10;

implementation

uses ZZFunc, ZZComm, MdmsFunc, LD8Q101, LD8Q11;

{$R *.DFM}

procedure TfrmLD8Q10.UP_VarMemSet(iLength : Integer);
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
procedure TfrmLD8Q10.grdMasterCellLoaded(Sender: TObject; DataCol,
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

procedure TfrmLD8Q10.FormCreate(Sender: TObject);
begin
   inherited;
   cmbPart.ItemIndex := 0;
   cmbJisa.ItemIndex := 0 ; //20140430 곽윤설
   // Grid 설정 & Memory Allocation
   UP_VarMemSet(0);
   // Grid 초기화
   grdMaster.Rows := 0;
end;

procedure TfrmLD8Q10.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnST_Date then GF_CalendarClick(mskST_Date);
   if Sender = btnED_Date then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD8Q10.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if       Sender = mskST_Date       then UP_Click(mskST_Date)
   else if  Sender = mskED_Date       then UP_Click(mskED_Date);
end;

procedure TfrmLD8Q10.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;
   Try
      Begin
           frmLD8Q101 := TfrmLD8Q101.Create(Self);
           frmLD8Q101.QuickRep.Preview;
      End
   Finally
           frmLD8Q101.Free
   End;
end;

procedure TfrmLD8Q10.btnQueryClick(Sender: TObject);
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

      SqlSelect := ' SELECT G.Cod_jisa, G.dat_gmgn, G.Cod_hul, G.Cod_Cancer, G.Cod_jangbi, G.Cod_chuga, G.Gubn_nosin, G.Gubn_nsyh, G.num_Id, C.ysno_sokun, C.cod_hm, C.dat_result, C.dat_last, ' +
                   ' G.Gubn_adult, G.Gubn_adyh, G.Gubn_agab, G.Gubn_agyh, G.Gubn_life, G.Gubn_lfyh, G.Gubn_hulgum, B.cod_bjjs, B.dat_bunju, B.num_bunju, G.num_jumin, G.gubn_tkgm ';
      SqlFrom   := ' FROM CELL C with(nolock) INNER JOIN Bunju B ON C.cod_jisa = B.cod_jisa AND C.num_id = B.num_id AND C.dat_gmgn = B.dat_gmgn ' ;
      SqlFrom   := SqlFrom + ' INNER JOIN GUMGIN G ON C.cod_jisa = G.cod_jisa AND C.num_id = G.num_id AND C.dat_gmgn = G.dat_gmgn ' ;
      if RadioButton1.checked then
      begin
      SqlWhere := SqlWhere + ' where C.Dat_last >= ''' + mskSt_Date.Text + '''';
      SqlWhere := SqlWhere + ' AND C.Dat_last <= ''' + mskEd_Date.Text + '''';
      end
      else
      begin
      SqlWhere := SqlWhere + ' where C.desc_buwi >= ''' + Edts_No.Text + '''';
      SqlWhere := SqlWhere + ' AND C.desc_buwi <= ''' + Edte_No.Text + '''';
      end;
      SqlWhere := SqlWhere + ' AND c.ysno_sokun <= ''C''';
      SqlWhere := SqlWhere + ' AND c.desc_buwi <> ''''';
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
              if Copy(cmbPart.Text,1,2) = '00' then SqlWhere := 'Where Gubn_part = ''06'' '
         else if Copy(cmbPart.Text,1,2) = '01' then SqlWhere := 'Where cod_hm = ''B001'' '
         else if Copy(cmbPart.Text,1,2) = '02' then SqlWhere := 'Where cod_hm = ''B005'' '
         else if Copy(cmbPart.Text,1,2) = '03' then SqlWhere := 'Where cod_hm = ''B007'' '
         else if Copy(cmbPart.Text,1,2) = '04' then SqlWhere := 'Where cod_hm = ''B012'' '
         else if Copy(cmbPart.Text,1,2) = '05' then SqlWhere := 'Where cod_hm = ''P009'' '
         else if Copy(cmbPart.Text,1,2) = '06' then SqlWhere := 'Where cod_hm = ''B008'' '
         else if Copy(cmbPart.Text,1,2) = '07' then SqlWhere := 'Where cod_hm = ''B009'' '
         else if Copy(cmbPart.Text,1,2) = '08' then SqlWhere := 'Where cod_hm = ''P003'' '
         else if Copy(cmbPart.Text,1,2) = '09' then SqlWhere := 'Where cod_hm = ''P006'' '
         else if Copy(cmbPart.Text,1,2) = '10' then SqlWhere := 'Where cod_hm = ''P010'' '
         else if Copy(cmbPart.Text,1,2) = '11' then SqlWhere := 'Where cod_hm = ''P011'' '
         else if Copy(cmbPart.Text,1,2) = '12' then SqlWhere := 'Where cod_hm = ''P012'' '
         else if Copy(cmbPart.Text,1,2) = '13' then SqlWhere := 'Where cod_hm = ''P013'' '
         else if Copy(cmbPart.Text,1,2) = '14' then sqlWhere := 'Where cod_hm = ''P007'' '
         else if Copy(cmbPart.Text,1,2) = '15' then sqlWhere := 'Where cod_hm = ''P008'' ';

         SqlWhere := SqlWhere + ' Group By Gubn_part, Cod_hm, Desc_kor, Pos_Start';
         SqlORder  := ' Order By Gubn_Part, Cod_hm';

         Sql.Clear;
         Sql.Add(SqlSelect + SqlFrom + SqlWhere + SqlOrder);
         GP_query_log(qryGumgin, 'LD8Q10', 'Q', 'N');
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

           QRY_TEMP.Active := False;
           QRY_TEMP.SQL.Clear;
           QRY_TEMP.SQL.Text := ' SELECT * FROM dbo.TF_Get_HangmokList('''+qryGumgin.FieldByName('cod_jisa').AsString +''','''+ qryGumgin.FieldByName('num_id').AsString+''','''+ qryGumgin.FieldByName('dat_gmgn').AsString+''',''1'') ' +
                                ' WHERE COD_HM =''B001'' OR COD_HM =''B005'' OR COD_HM =''B007'' OR COD_HM =''B009'' OR COD_HM =''P003'' OR COD_HM =''P009'' OR COD_HM =''P010'' OR COD_HM =''P012'' OR COD_HM =''P013'' OR COD_HM =''P007''  OR COD_HM =''P008''';
           QRY_TEMP.Active := True;

           while Not QRY_TEMP.Eof Do
           begin
               sValue := sValue + QRY_TEMP.FieldByName('Cod_hm').AsString;
               QRY_TEMP.Next;
            end;
            QRY_TEMP.Close;

         // 각각 항목 Count...
         wk_rec := wk_rec + 1;
         ValEdit3.Value  := wk_rec;
         application.ProcessMessages;

            for Index := 0 to K - 1 do
            begin
               sCode := Copy(sValue, J * 4 + 1, 4);

               if (qryGumgin.FieldByName('cod_hm').AsString = UV_vCode[Index]) Then
               begin
                  if (qryGumgin.FieldByName('Cod_jisa').AsString = '15') Then
                  begin
                     Uv_v15[Index] := Uv_v15[Index] + 1;
                     Uv_vtot[Index] := Uv_vtot[Index] + 1;
                    { if (QryGumgin.FieldByName('dat_gmgn').AsString > '20140331') then
                             UP_Insert_2012(Index, sCode, QryGumgin.FieldByName('gubn_nosin').AsString, QryGumgin.FieldByName('Gubn_adult').AsString,
                                                          QryGumgin.FieldByName('gubn_life').AsString,  QryGumgin.FieldByName('Gubn_hulgum').AsString,
                                                          QryGumgin.FieldByName('cod_jisa').AsString)
                     else    UP_Insert_2013(Index, sCode, QryGumgin.FieldByName('cod_jangbi').AsString,  QryGumgin.FieldByName('cod_hul').AsString,
                                                          QryGumgin.FieldByName('Gubn_hulgum').AsString, QryGumgin.FieldByName('cod_jisa').AsString)}
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '12') AND ((Copy(cmbJisa.Text,1,2)='12') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v12[Index] := Uv_v12[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '11') AND ((Copy(cmbJisa.Text,1,2)='11') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v11[Index] := Uv_v11[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '43') AND ((Copy(cmbJisa.Text,1,2)='43') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v43[Index] := Uv_v43[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '71') AND ((Copy(cmbJisa.Text,1,2)='71') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v71[Index] := Uv_v71[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '61') AND ((Copy(cmbJisa.Text,1,2)='61') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v61[Index] := Uv_v61[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '51') AND ((Copy(cmbJisa.Text,1,2)='51') or (Copy(cmbJisa.Text,1,2)='00')) Then
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

Function TfrmLD8Q10.HangmokFind( vTemp , vHangmok : String) : String;
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

procedure TfrmLD8Q10.SBut_ExcelClick(Sender: TObject);
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
procedure TfrmLD8Q10.cmbRelationChange(Sender: TObject);
begin
  inherited;
   if cmbRelation.ItemIndex = 0 then
   begin
      Application.CreateForm(TfrmLD8Q11, frmLD8Q11);
      frmLD8Q11.Show;
   end
end;

end.
