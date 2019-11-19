// 수정일자      : 2018-04-21
// 수정자        : 박수지
// 수정내용      : 질관리 관련 프로그램 생성요청
// 참고사항      :
//==============================================================================


unit LD8Q11;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ValEdit,  ComObj, Gauges;

type
  TfrmLD8Q11 = class(TfrmSingle)
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
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit15: TEdit;
    Edit16: TEdit;
    Edit17: TEdit;
    Edit18: TEdit;
    Edit19: TEdit;
    Edit20: TEdit;
    Edit21: TEdit;
    Edit22: TEdit;
    Edit23: TEdit;
    Edit24: TEdit;
    Edit25: TEdit;
    Edit26: TEdit;
    Edit27: TEdit;
    Edit28: TEdit;
    Edit29: TEdit;
    Edit30: TEdit;
    Edit31: TEdit;
    Edit32: TEdit;
    Edit33: TEdit;
    Edit34: TEdit;
    Edit35: TEdit;
    Edit36: TEdit;
    Edit37: TEdit;
    Edit38: TEdit;
    Edit39: TEdit;
    Edit40: TEdit;
    Edit41: TEdit;
    Edit42: TEdit;
    Edit43: TEdit;
    Edit44: TEdit;
    Edit45: TEdit;
    Edit46: TEdit;
    Edit47: TEdit;
    Edit48: TEdit;
    Edit49: TEdit;
    Edit50: TEdit;
    Edit51: TEdit;
    Edit52: TEdit;
    Edit53: TEdit;
    Edit54: TEdit;
    Edit55: TEdit;
    Edit56: TEdit;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    Panel23: TPanel;
    Panel24: TPanel;
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
    UV_vPart, UV_vCode, UV_vLevel, UV_vDesc_hm,
    UV_v11, UV_v12, UV_v15, UV_v43, UV_v51, UV_v61, UV_v71, UV_vTot, UV_vTot_L,
    UV_v11_L, UV_v12_L, UV_v15_L, UV_v43_L, UV_v51_L, UV_v61_L, UV_v71_L : Variant;

    procedure UP_VarMemSet(iLength : Integer);
    Function  HangmokFind( vTemp , vHangmok : String) : String;
    //procedure UP_Insert_2012(iTemp : Integer; sHangmok, sGubn_nosin, sGubn_adult, sGubn_life, sGubn_hulgum, sCod_jisa : String);
   // procedure UP_Insert_2013(iTemp : Integer; sHangmok, sCod_jangbi, sCod_hul, sGubn_hulgum, sCod_jisa : String);
  public
    { Public declarations }
  end;


var
  frmLD8Q11 : TfrmLD8Q11;

implementation

uses ZZFunc, ZZComm, MdmsFunc, LD8Q111;

{$R *.DFM}

procedure TfrmLD8Q11.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vPart    := VarArrayCreate([0, 0], varOleStr);
      UV_vCode    := VarArrayCreate([0, 0], varOleStr);
      UV_vLevel   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hm := VarArrayCreate([0, 0], varOleStr);
      UV_v11      := VarArrayCreate([0, 0], varOleStr);
      UV_v11_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v12      := VarArrayCreate([0, 0], varOleStr);
      UV_v12_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v15      := VarArrayCreate([0, 0], varOleStr);
      UV_v15_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v43      := VarArrayCreate([0, 0], varOleStr);
      UV_v43_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v51      := VarArrayCreate([0, 0], varOleStr);
      UV_v51_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v61      := VarArrayCreate([0, 0], varOleStr);
      UV_v61_L    := VarArrayCreate([0, 0], varOleStr);
      UV_v71      := VarArrayCreate([0, 0], varOleStr);
      UV_v71_L    := VarArrayCreate([0, 0], varOleStr);
      UV_vTot     := VarArrayCreate([0, 0], varOleStr);
      UV_vTot_L   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vPart,    iLength);
      VarArrayRedim(UV_vCode,    iLength);
      VarArrayRedim(UV_vLevel,    iLength);
      VarArrayRedim(UV_vDesc_hm, iLength);
      VarArrayRedim(UV_v11,      iLength);
      VarArrayRedim(UV_v11_L,    iLength);
      VarArrayRedim(UV_v12,      iLength);
      VarArrayRedim(UV_v12_L,    iLength);
      VarArrayRedim(UV_v15,      iLength);
      VarArrayRedim(UV_v15_L,    iLength);
      VarArrayRedim(UV_v43,      iLength);
      VarArrayRedim(UV_v43_L,    iLength);
      VarArrayRedim(UV_v51,      iLength);
      VarArrayRedim(UV_v51_L,    iLength);
      VarArrayRedim(UV_v61,      iLength);
      VarArrayRedim(UV_v61_L,    iLength);
      VarArrayRedim(UV_v71,      iLength);
      VarArrayRedim(UV_v71_L,    iLength);
      VarArrayRedim(UV_vTot,     iLength);
      VarArrayRedim(UV_vTot_L,   iLength);
   end;
end;
procedure TfrmLD8Q11.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vPart[DataRow-1];
      2 : Value := UV_vCode[DataRow - 1];
      3 : Value := UV_vDesc_hm[DataRow-1];
      4 : Value := UV_v15_L[DataRow-1];
      5 : Value := UV_v12_L[DataRow-1];
      6 : Value := UV_v11_L[DataRow-1];
      7 : Value := UV_v43_L[DataRow-1];
      8 : Value := UV_v71_L[DataRow-1];
      9 : Value := UV_v61_L[DataRow-1];
     10 : Value := UV_v51_L[DataRow-1];
     11 : Value := UV_vTot_L[DataRow-1];
   end;
end;

procedure TfrmLD8Q11.FormCreate(Sender: TObject);
begin
   inherited;
   cmbPart.ItemIndex := 0;
   cmbJisa.ItemIndex := 0 ; //20140430 곽윤설
   // Grid 설정 & Memory Allocation
   UP_VarMemSet(0);
   // Grid 초기화
   grdMaster.Rows := 0;
end;

procedure TfrmLD8Q11.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnST_Date then GF_CalendarClick(mskST_Date);
   if Sender = btnED_Date then GF_CalendarClick(mskED_Date);
end;

procedure TfrmLD8Q11.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if       Sender = mskST_Date       then UP_Click(mskST_Date)
   else if  Sender = mskED_Date       then UP_Click(mskED_Date);
end;

procedure TfrmLD8Q11.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;
   Try
      Begin
           frmLD8Q111 := TfrmLD8Q111.Create(Self);
           frmLD8Q111.QuickRep.Preview;
      End
   Finally
           frmLD8Q111.Free
   End;
end;

procedure TfrmLD8Q11.btnQueryClick(Sender: TObject);
Var
   SqlWhere, SqlFrom, SqlSelect, SqlOrder : String;
   sSelect, sOderby, sWhere1, sWhere2 : String;

   Index, sPos, I, J, K : Integer;
   wk_11, wk_12, wk_15, wk_33, wk_34, wk_35, wk_41, wk_43, wk_45, wk_46, wk_rec,
   wk_47, wk_51, wk_52, wk_61, wk_63, wk_71, wk_72, wk_gita, wk_tot,iRet,Temp: Integer;

   //레벨표기 .. 2018.09.06
   b1_level_15_a, b1_level_15_b, b1_level_15_c1, b1_level_15_c2, b1_level_15_d1, b1_level_15_d2, b1_level_15_d3,
   b7_level_15_a, b7_level_15_b, b7_level_15_c1, b7_level_15_c2, b7_level_15_d1, b7_level_15_d2, b7_level_15_d3,
   P9_level_15_a, P9_level_15_b, P9_level_15_c1, P9_level_15_c2, P9_level_15_d1, P9_level_15_d2, P9_level_15_d3,

   tot_level_a, tot_level_b, tot_level_c1, tot_level_c2, tot_level_d1, tot_level_d2, tot_level_d3,

   b1_level_12_a, b1_level_12_b, b1_level_12_c1, b1_level_12_c2, b1_level_12_d1, b1_level_12_d2, b1_level_12_d3,
   b7_level_12_a, b7_level_12_b, b7_level_12_c1, b7_level_12_c2, b7_level_12_d1, b7_level_12_d2, b7_level_12_d3,
   P9_level_12_a, P9_level_12_b, P9_level_12_c1, P9_level_12_c2, P9_level_12_d1, P9_level_12_d2, P9_level_12_d3,
   //tot_level_12_a, tot_level_12_b, tot_level_12_c1, tot_level_12_c2, tot_level_12_d1, tot_level_12_d2, tot_level_12_d3,

   b1_level_11_a, b1_level_11_b, b1_level_11_c1, b1_level_11_c2, b1_level_11_d1, b1_level_11_d2, b1_level_11_d3,
   b7_level_11_a, b7_level_11_b, b7_level_11_c1, b7_level_11_c2, b7_level_11_d1, b7_level_11_d2, b7_level_11_d3,
   P9_level_11_a, P9_level_11_b, P9_level_11_c1, P9_level_11_c2, P9_level_11_d1, P9_level_11_d2, P9_level_11_d3,
   //tot_level_11_a, tot_level_11_b, tot_level_11_c1, tot_level_11_c2, tot_level_11_d1, tot_level_11_d2, tot_level_11_d3,

   b1_level_43_a, b1_level_43_b, b1_level_43_c1, b1_level_43_c2, b1_level_43_d1, b1_level_43_d2, b1_level_43_d3,
   b7_level_43_a, b7_level_43_b, b7_level_43_c1, b7_level_43_c2, b7_level_43_d1, b7_level_43_d2, b7_level_43_d3,
   P9_level_43_a, P9_level_43_b, P9_level_43_c1, P9_level_43_c2, P9_level_43_d1, P9_level_43_d2, P9_level_43_d3,
   //tot_level_43_a, tot_level_43_b, tot_level_43_c1, tot_level_43_c2, tot_level_43_d1, tot_level_43_d2, tot_level_43_d3,

   b1_level_51_a, b1_level_51_b, b1_level_51_c1, b1_level_51_c2, b1_level_51_d1, b1_level_51_d2, b1_level_51_d3,
   b7_level_51_a, b7_level_51_b, b7_level_51_c1, b7_level_51_c2, b7_level_51_d1, b7_level_51_d2, b7_level_51_d3,
   P9_level_51_a, P9_level_51_b, P9_level_51_c1, P9_level_51_c2, P9_level_51_d1, P9_level_51_d2, P9_level_51_d3,
   //tot_level_51_a, tot_level_51_b, tot_level_51_c1, tot_level_51_c2, tot_level_51_d1, tot_level_51_d2, tot_level_51_d3,

   b1_level_61_a, b1_level_61_b, b1_level_61_c1, b1_level_61_c2, b1_level_61_d1, b1_level_61_d2, b1_level_61_d3,
   b7_level_61_a, b7_level_61_b, b7_level_61_c1, b7_level_61_c2, b7_level_61_d1, b7_level_61_d2, b7_level_61_d3,
   P9_level_61_a, P9_level_61_b, P9_level_61_c1, P9_level_61_c2, P9_level_61_d1, P9_level_61_d2, P9_level_61_d3,
   //tot_level_61_a, tot_level_61_b, tot_level_61_c1, tot_level_61_c2, tot_level_61_d1, tot_level_61_d2, tot_level_61_d3,

   b1_level_71_a, b1_level_71_b, b1_level_71_c1, b1_level_71_c2, b1_level_71_d1, b1_level_71_d2, b1_level_71_d3,
   b7_level_71_a, b7_level_71_b, b7_level_71_c1, b7_level_71_c2, b7_level_71_d1, b7_level_71_d2, b7_level_71_d3,
   P9_level_71_a, P9_level_71_b, P9_level_71_c1, P9_level_71_c2, P9_level_71_d1, P9_level_71_d2, P9_level_71_d3 : Integer;
   //tot_level_71_a, tot_level_71_b, tot_level_71_c1, tot_level_71_c2, tot_level_71_d1, tot_level_71_d2, tot_level_71_d3
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

      SqlSelect := ' SELECT G.Cod_jisa, G.dat_gmgn, G.Cod_hul, G.Cod_Cancer, G.Cod_jangbi, G.Cod_chuga, G.Gubn_nosin, G.Gubn_nsyh, G.num_Id, C.ysno_sokun, C.cod_hm, C.dat_result, C.cell_level, ' +
                   ' G.Gubn_adult, G.Gubn_adyh, G.Gubn_agab, G.Gubn_agyh, G.Gubn_life, G.Gubn_lfyh, G.Gubn_hulgum, B.cod_bjjs, B.dat_bunju, B.num_bunju, G.num_jumin, G.gubn_tkgm ';
      SqlFrom   := ' FROM CELL C with(nolock) INNER JOIN Bunju B ON C.cod_jisa = B.cod_jisa AND C.num_id = B.num_id AND C.dat_gmgn = B.dat_gmgn ' ;
      SqlFrom   := SqlFrom + ' INNER JOIN GUMGIN G ON C.cod_jisa = G.cod_jisa AND C.num_id = G.num_id AND C.dat_gmgn = G.dat_gmgn ' ;
      if RadioButton1.checked then
      begin
      SqlWhere := SqlWhere + ' where C.Dat_result >= ''' + mskSt_Date.Text + '''';
      SqlWhere := SqlWhere + ' AND C.Dat_result <= ''' + mskEd_Date.Text + '''';
      end
      else
      begin
      SqlWhere := SqlWhere + ' where C.desc_buwi >= ''' + Edts_No.Text + '''';
      SqlWhere := SqlWhere + ' AND C.desc_buwi <= ''' + Edte_No.Text + '''';
      end;
      SqlWhere := SqlWhere + ' AND c.ysno_sokun = ''C''';
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
              if Copy(cmbPart.Text,1,2) = '00' then  SqlWhere := 'Where cod_hm = ''B001'' OR cod_hm = ''B007'' OR  cod_hm = ''P009'' ';

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
         Uv_vtot_L[Index]:= 0;
      end; // end of for(Index);


      While Not Eof Do
      Begin
         sValue := '';
         sSelect := ''; sOderby := ''; sWhere1 := ''; sWhere2 := '';

           QRY_TEMP.Active := False;
           QRY_TEMP.SQL.Clear;
           QRY_TEMP.SQL.Text := ' SELECT * FROM dbo.TF_Get_HangmokList('''+qryGumgin.FieldByName('cod_jisa').AsString +''','''+ qryGumgin.FieldByName('num_id').AsString+''','''+ qryGumgin.FieldByName('dat_gmgn').AsString+''',''1'') ' +
                                ' WHERE COD_HM =''B001'' OR COD_HM =''B005'' OR COD_HM =''B007'' OR COD_HM =''B009'' OR COD_HM =''P003'' OR COD_HM =''P009'' OR COD_HM =''P010'' OR COD_HM =''P012'' OR COD_HM =''P013''  ';
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

                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_15_a  := b1_level_15_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_15_b  := b1_level_15_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_15_c1 := b1_level_15_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_15_c2 := b1_level_15_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_15_d1 := b1_level_15_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_15_d2 := b1_level_15_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_15_d3 := b1_level_15_d3 + 1
                            else  ;
                            Uv_v15_L[Index] := Uv_v15[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_15_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_15_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_15_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_15_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_15_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_15_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_15_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_15_a  := b7_level_15_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_15_b  := b7_level_15_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_15_c1 := b7_level_15_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_15_c2 := b7_level_15_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_15_d1 := b7_level_15_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_15_d2 := b7_level_15_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_15_d3 := b7_level_15_d3 + 1
                            else  ;
                            Uv_v15_L[Index] := Uv_v15[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_15_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_15_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_15_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_15_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_15_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_15_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_15_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_15_a  := p9_level_15_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_15_b  := p9_level_15_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_15_c1 := p9_level_15_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_15_c2 := p9_level_15_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_15_d1 := p9_level_15_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_15_d2 := p9_level_15_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_15_d3 := p9_level_15_d3 + 1
                            else  ;
                            Uv_v15_L[Index] := Uv_v15[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_15_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_15_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_15_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_15_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_15_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_15_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_15_d3);
                        end
                     end;

                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '12') AND ((Copy(cmbJisa.Text,1,2)='12') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v12[Index] := Uv_v12[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;

                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_12_a  := b1_level_12_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_12_b  := b1_level_12_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_12_c1 := b1_level_12_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_12_c2 := b1_level_12_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_12_d1 := b1_level_12_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_12_d2 := b1_level_12_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_12_d3 := b1_level_12_d3 + 1
                            else  ;
                            Uv_v12_L[Index] := Uv_v12[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_12_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_12_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_12_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_12_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_12_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_12_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_12_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_12_a  := b7_level_12_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_12_b  := b7_level_12_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_12_c1 := b7_level_12_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_12_c2 := b7_level_12_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_12_d1 := b7_level_12_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_12_d2 := b7_level_12_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_12_d3 := b7_level_12_d3 + 1
                            else  ;
                            Uv_v12_L[Index] := Uv_v12[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_12_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_12_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_12_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_12_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_12_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_12_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_12_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_12_a  := p9_level_12_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_12_b  := p9_level_12_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_12_c1 := p9_level_12_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_12_c2 := p9_level_12_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_12_d1 := p9_level_12_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_12_d2 := p9_level_12_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_12_d3 := p9_level_12_d3 + 1
                            else  ;
                            Uv_v12_L[Index] := Uv_v12[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_12_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_12_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_12_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_12_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_12_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_12_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_12_d3);
                        end
                     end;

                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '11') AND ((Copy(cmbJisa.Text,1,2)='11') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v11[Index] := Uv_v11[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;

                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_11_a  := b1_level_11_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_11_b  := b1_level_11_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_11_c1 := b1_level_11_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_11_c2 := b1_level_11_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_11_d1 := b1_level_11_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_11_d2 := b1_level_11_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_11_d3 := b1_level_11_d3 + 1
                            else  ;
                            Uv_v11_L[Index] := Uv_v11[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_11_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_11_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_11_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_11_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_11_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_11_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_11_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_11_a  := b7_level_11_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_11_b  := b7_level_11_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_11_c1 := b7_level_11_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_11_c2 := b7_level_11_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_11_d1 := b7_level_11_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_11_d2 := b7_level_11_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_11_d3 := b7_level_11_d3 + 1
                            else  ;
                            Uv_v11_L[Index] := Uv_v11[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_11_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_11_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_11_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_11_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_11_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_11_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_11_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_11_a  := p9_level_11_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_11_b  := p9_level_11_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_11_c1 := p9_level_11_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_11_c2 := p9_level_11_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_11_d1 := p9_level_11_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_11_d2 := p9_level_11_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_11_d3 := p9_level_11_d3 + 1
                            else  ;
                            Uv_v11_L[Index] := Uv_v11[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_11_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_11_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_11_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_11_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_11_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_11_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_11_d3);
                        end
                     end;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '43') AND ((Copy(cmbJisa.Text,1,2)='43') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v43[Index] := Uv_v43[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;
                                                           
                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin                                 
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_43_a  := b1_level_43_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_43_b  := b1_level_43_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_43_c1 := b1_level_43_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_43_c2 := b1_level_43_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_43_d1 := b1_level_43_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_43_d2 := b1_level_43_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_43_d3 := b1_level_43_d3 + 1
                            else  ;
                            Uv_v43_L[Index] := Uv_v43[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_43_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_43_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_43_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_43_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_43_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_43_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_43_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_43_a  := b7_level_43_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_43_b  := b7_level_43_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_43_c1 := b7_level_43_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_43_c2 := b7_level_43_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_43_d1 := b7_level_43_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_43_d2 := b7_level_43_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_43_d3 := b7_level_43_d3 + 1
                            else  ;
                            Uv_v43_L[Index] := Uv_v43[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_43_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_43_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_43_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_43_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_43_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_43_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_43_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_43_a  := p9_level_43_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_43_b  := p9_level_43_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_43_c1 := p9_level_43_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_43_c2 := p9_level_43_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_43_d1 := p9_level_43_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_43_d2 := p9_level_43_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_43_d3 := p9_level_43_d3 + 1
                            else  ;
                            Uv_v43_L[Index] := Uv_v43[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_43_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_43_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_43_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_43_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_43_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_43_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_43_d3);
                        end
                     end;

                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '71') AND ((Copy(cmbJisa.Text,1,2)='71') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v71[Index] := Uv_v71[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;

                    if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_71_a  := b1_level_71_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_71_b  := b1_level_71_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_71_c1 := b1_level_71_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_71_c2 := b1_level_71_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_71_d1 := b1_level_71_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_71_d2 := b1_level_71_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_71_d3 := b1_level_71_d3 + 1
                            else  ;
                            Uv_v71_L[Index] := Uv_v71[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_71_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_71_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_71_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_71_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_71_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_71_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_71_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_71_a  := b7_level_71_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_71_b  := b7_level_71_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_71_c1 := b7_level_71_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_71_c2 := b7_level_71_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_71_d1 := b7_level_71_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_71_d2 := b7_level_71_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_71_d3 := b7_level_71_d3 + 1
                            else  ;
                            Uv_v71_L[Index] := Uv_v71[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_71_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_71_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_71_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_71_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_71_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_71_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_71_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_71_a  := p9_level_71_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_71_b  := p9_level_71_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_71_c1 := p9_level_71_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_71_c2 := p9_level_71_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_71_d1 := p9_level_71_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_71_d2 := p9_level_71_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_71_d3 := p9_level_71_d3 + 1
                            else  ;
                            Uv_v71_L[Index] := Uv_v71[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_71_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_71_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_71_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_71_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_71_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_71_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_71_d3);
                        end
                     end;

                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '61') AND ((Copy(cmbJisa.Text,1,2)='61') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v61[Index] := Uv_v61[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;

                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                     begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_61_a  := b1_level_61_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_61_b  := b1_level_61_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_61_c1 := b1_level_61_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_61_c2 := b1_level_61_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_61_d1 := b1_level_61_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_61_d2 := b1_level_61_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_61_d3 := b1_level_61_d3 + 1
                            else  ;
                            Uv_v61_L[Index] := Uv_v61[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_61_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_61_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_61_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_61_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_61_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_61_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_61_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_61_a  := b7_level_61_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_61_b  := b7_level_61_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_61_c1 := b7_level_61_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_61_c2 := b7_level_61_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_61_d1 := b7_level_61_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_61_d2 := b7_level_61_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_61_d3 := b7_level_61_d3 + 1
                            else  ;
                            Uv_v61_L[Index] := Uv_v61[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_61_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_61_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_61_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_61_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_61_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_61_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_61_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_61_a  := p9_level_61_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_61_b  := p9_level_61_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_61_c1 := p9_level_61_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_61_c2 := p9_level_61_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_61_d1 := p9_level_61_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_61_d2 := p9_level_61_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_61_d3 := p9_level_61_d3 + 1
                            else  ;
                            Uv_v61_L[Index] := Uv_v61[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_61_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_61_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_61_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_61_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_61_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_61_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_61_d3);
                        end
                     end;
                  end
                  else if (qryGumgin.FieldByName('Cod_jisa').AsString = '51') AND ((Copy(cmbJisa.Text,1,2)='51') or (Copy(cmbJisa.Text,1,2)='00')) Then
                  begin
                          Uv_v51[Index] := Uv_v51[Index] + 1;
                          Uv_vtot[Index] := Uv_vtot[Index] + 1;

                     if (qryGumgin.FieldByName('Cell_level').AsString <> '') then
                      begin
                        if UV_vCode[Index] = 'B001' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b1_level_51_a  := b1_level_51_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b1_level_51_b  := b1_level_51_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b1_level_51_c1 := b1_level_51_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b1_level_51_c2 := b1_level_51_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b1_level_51_d1 := b1_level_51_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b1_level_51_d2 := b1_level_51_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b1_level_51_d3 := b1_level_51_d3 + 1
                            else  ;
                            Uv_v51_L[Index] := Uv_v51[Index] + #13#10 + 'A : ' + INTTOSTR(b1_level_51_a) + #13#10 + 'B : ' + INTTOSTR(b1_level_51_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b1_level_51_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b1_level_51_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b1_level_51_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b1_level_51_d2) + #13#10 + 'D3 : ' + INTTOSTR(b1_level_51_d3);
                        end
                        else if UV_vCode[Index] = 'B007' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then b7_level_51_a  := b7_level_51_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then b7_level_51_b  := b7_level_51_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then b7_level_51_c1 := b7_level_51_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then b7_level_51_c2 := b7_level_51_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then b7_level_51_d1 := b7_level_51_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then b7_level_51_d2 := b7_level_51_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then b7_level_51_d3 := b7_level_51_d3 + 1
                            else  ;
                            Uv_v51_L[Index] := Uv_v51[Index] + #13#10 + 'A : ' + INTTOSTR(b7_level_51_a) + #13#10 + 'B : ' + INTTOSTR(b7_level_51_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(b7_level_51_c1)+ #13#10 + 'C2 : ' + INTTOSTR(b7_level_51_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(b7_level_51_d1)+ #13#10 + 'D2 : ' + INTTOSTR(b7_level_51_d2) + #13#10 + 'D3 : ' + INTTOSTR(b7_level_51_d3);
                        end
                        else if UV_vCode[Index] = 'P009' then
                        begin
                                 if qryGumgin.FieldByName('Cell_level').AsString = '1' then P9_level_51_a  := p9_level_51_a + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '2' then p9_level_51_b  := p9_level_51_b + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '3' then p9_level_51_c1 := p9_level_51_c1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '4' then p9_level_51_c2 := p9_level_51_c2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '5' then p9_level_51_d1 := p9_level_51_d1 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '6' then p9_level_51_d2 := p9_level_51_d2 + 1
                            else if qryGumgin.FieldByName('Cell_level').AsString = '7' then p9_level_51_d3 := p9_level_51_d3 + 1
                            else  ;
                            Uv_v51_L[Index] := Uv_v51[Index] + #13#10 + 'A : ' + INTTOSTR(P9_level_51_a) + #13#10 + 'B : ' + INTTOSTR(P9_level_51_b) +#13#10 +
                                               'C1 : ' + INTTOSTR(P9_level_51_c1)+ #13#10 + 'C2 : ' + INTTOSTR(P9_level_51_c2) + #13#10 +
                                               'D1 : ' + INTTOSTR(P9_level_51_d1)+ #13#10 + 'D2 : ' + INTTOSTR(P9_level_51_d2) + #13#10 + 'D3 : ' + INTTOSTR(P9_level_51_d3);
                        end
                     end;
                  end;

                  Uv_vtot_L[0] := Uv_vtot[0] + #13#10 +
                                      'A : ' + INTTOSTR(b1_level_15_a  + b1_level_11_a  + b1_level_12_a  + b1_level_43_a  + b1_level_51_a  + b1_level_61_a  + b1_level_71_a)  + #13#10 +
                                      'B : ' + INTTOSTR(b1_level_15_b  + b1_level_11_b  + b1_level_12_b  + b1_level_43_b  + b1_level_51_b  + b1_level_61_b  + b1_level_71_b)  + #13#10 +
                                     'C1 : ' + INTTOSTR(b1_level_15_c1 + b1_level_11_c1 + b1_level_12_c1 + b1_level_43_c1 + b1_level_51_c1 + b1_level_61_c1 + b1_level_71_c1) + #13#10 +
                                     'C2 : ' + INTTOSTR(b1_level_15_c2 + b1_level_11_c2 + b1_level_12_c2 + b1_level_43_c2 + b1_level_51_c2 + b1_level_61_c2 + b1_level_71_c2) + #13#10 +
                                     'D1 : ' + INTTOSTR(b1_level_15_d1 + b1_level_11_d1 + b1_level_12_d1 + b1_level_43_d1 + b1_level_51_d1 + b1_level_61_d1 + b1_level_71_d1) + #13#10 +
                                     'D2 : ' + INTTOSTR(b1_level_15_d2 + b1_level_11_d2 + b1_level_12_d2 + b1_level_43_d2 + b1_level_51_d2 + b1_level_61_d2 + b1_level_71_d2) + #13#10 +
                                     'D3 : ' + INTTOSTR(b1_level_15_d3 + b1_level_11_d3 + b1_level_12_d3 + b1_level_43_d3 + b1_level_51_d3 + b1_level_61_d3 + b1_level_71_d3) ;
                  Uv_vtot_L[1] := Uv_vtot[1] + #13#10 +
                                      'A : ' + INTTOSTR(b7_level_15_a  + b7_level_11_a  + b7_level_12_a  + b7_level_43_a  + b7_level_51_a  + b7_level_61_a  + b7_level_71_a)  + #13#10 +
                                      'B : ' + INTTOSTR(b7_level_15_b  + b7_level_11_b  + b7_level_12_b  + b7_level_43_b  + b7_level_51_b  + b7_level_61_b  + b7_level_71_b)  + #13#10 +
                                     'C1 : ' + INTTOSTR(b7_level_15_c1 + b7_level_11_c1 + b7_level_12_c1 + b7_level_43_c1 + b7_level_51_c1 + b7_level_61_c1 + b7_level_71_c1) + #13#10 +
                                     'C2 : ' + INTTOSTR(b7_level_15_c2 + b7_level_11_c2 + b7_level_12_c2 + b7_level_43_c2 + b7_level_51_c2 + b7_level_61_c2 + b7_level_71_c2) + #13#10 +
                                     'D1 : ' + INTTOSTR(b7_level_15_d1 + b7_level_11_d1 + b7_level_12_d1 + b7_level_43_d1 + b7_level_51_d1 + b7_level_61_d1 + b7_level_71_d1) + #13#10 +
                                     'D2 : ' + INTTOSTR(b7_level_15_d2 + b7_level_11_d2 + b7_level_12_d2 + b7_level_43_d2 + b7_level_51_d2 + b7_level_61_d2 + b7_level_71_d2) + #13#10 +
                                     'D3 : ' + INTTOSTR(b7_level_15_d3 + b7_level_11_d3 + b7_level_12_d3 + b7_level_43_d3 + b7_level_51_d3 + b7_level_61_d3 + b7_level_71_d3) ;
                  Uv_vtot_L[2] := Uv_vtot[2] + #13#10 +
                                      'A : ' + INTTOSTR(p9_level_15_a  + p9_level_11_a  + p9_level_12_a  + p9_level_43_a  + p9_level_51_a  + p9_level_61_a  + p9_level_71_a)  + #13#10 +
                                      'B : ' + INTTOSTR(p9_level_15_b  + p9_level_11_b  + p9_level_12_b  + p9_level_43_b  + p9_level_51_b  + p9_level_61_b  + p9_level_71_b)  + #13#10 +
                                     'C1 : ' + INTTOSTR(p9_level_15_c1 + p9_level_11_c1 + p9_level_12_c1 + p9_level_43_c1 + p9_level_51_c1 + p9_level_61_c1 + p9_level_71_c1) + #13#10 +
                                     'C2 : ' + INTTOSTR(p9_level_15_c2 + p9_level_11_c2 + p9_level_12_c2 + p9_level_43_c2 + p9_level_51_c2 + p9_level_61_c2 + p9_level_71_c2) + #13#10 +
                                     'D1 : ' + INTTOSTR(p9_level_15_d1 + p9_level_11_d1 + p9_level_12_d1 + p9_level_43_d1 + p9_level_51_d1 + p9_level_61_d1 + p9_level_71_d1) + #13#10 +
                                     'D2 : ' + INTTOSTR(p9_level_15_d2 + p9_level_11_d2 + p9_level_12_d2 + p9_level_43_d2 + p9_level_51_d2 + p9_level_61_d2 + p9_level_71_d2) + #13#10 +
                                     'D3 : ' + INTTOSTR(p9_level_15_d3 + p9_level_11_d3 + p9_level_12_d3 + p9_level_43_d3 + p9_level_51_d3 + p9_level_61_d3 + p9_level_71_d3) ;


               End;
               Edit1.Text :=  INTTOSTR(b1_level_15_a  + b1_level_11_a  + b1_level_12_a  + b1_level_43_a  + b1_level_51_a  + b1_level_61_a  + b1_level_71_a +
                                       b7_level_15_a  + b7_level_11_a  + b7_level_12_a  + b7_level_43_a  + b7_level_51_a  + b7_level_61_a  + b7_level_71_a +
                                       p9_level_15_a  + p9_level_11_a  + p9_level_12_a  + p9_level_43_a  + p9_level_51_a  + p9_level_61_a  + p9_level_71_a );

               Edit2.Text :=  INTTOSTR(b1_level_15_b  + b1_level_11_b  + b1_level_12_b  + b1_level_43_b  + b1_level_51_b  + b1_level_61_b  + b1_level_71_b +
                                       b7_level_15_b  + b7_level_11_b  + b7_level_12_b  + b7_level_43_b  + b7_level_51_b  + b7_level_61_b  + b7_level_71_b +
                                       p9_level_15_b  + p9_level_11_b  + p9_level_12_b  + p9_level_43_b  + p9_level_51_b  + p9_level_61_b  + p9_level_71_b );

               Edit3.Text :=  INTTOSTR(b1_level_15_c1  + b1_level_11_c1 + b1_level_12_c1 + b1_level_43_c1 + b1_level_51_c1 + b1_level_61_c1 + b1_level_71_c1 +
                                       b7_level_15_c1  + b7_level_11_c1 + b7_level_12_c1 + b7_level_43_c1 + b7_level_51_c1 + b7_level_61_c1 + b7_level_71_c1 +
                                       p9_level_15_c1  + p9_level_11_c1 + p9_level_12_c1 + p9_level_43_c1 + p9_level_51_c1 + p9_level_61_c1 + p9_level_71_c1 );

               Edit4.Text :=  INTTOSTR(b1_level_15_c2  + b1_level_11_c2  + b1_level_12_c2  + b1_level_43_c2  + b1_level_51_c2  + b1_level_61_c2  + b1_level_71_c2 +
                                       b7_level_15_c2  + b7_level_11_c2  + b7_level_12_c2  + b7_level_43_c2  + b7_level_51_c2  + b7_level_61_c2  + b7_level_71_c2 +
                                       p9_level_15_c2  + p9_level_11_c2  + p9_level_12_c2  + p9_level_43_c2  + p9_level_51_c2  + p9_level_61_c2  + p9_level_71_c2 );

               Edit5.Text :=  INTTOSTR(b1_level_15_d1  + b1_level_11_d1  + b1_level_12_d1  + b1_level_43_d1  + b1_level_51_d1  + b1_level_61_d1  + b1_level_71_d1 +
                                       b7_level_15_d1  + b7_level_11_d1  + b7_level_12_d1  + b7_level_43_d1  + b7_level_51_d1  + b7_level_61_d1  + b7_level_71_d1 +
                                       p9_level_15_d1  + p9_level_11_d1  + p9_level_12_d1  + p9_level_43_d1  + p9_level_51_d1  + p9_level_61_d1  + p9_level_71_d1 );

               Edit6.Text :=  INTTOSTR(b1_level_15_d2  + b1_level_11_d2  + b1_level_12_d2  + b1_level_43_d2  + b1_level_51_d2  + b1_level_61_d2  + b1_level_71_d2 +
                                       b7_level_15_d2  + b7_level_11_d2  + b7_level_12_d2  + b7_level_43_d2  + b7_level_51_d2  + b7_level_61_d2  + b7_level_71_d2 +
                                       p9_level_15_d2  + p9_level_11_d2  + p9_level_12_d2  + p9_level_43_d2  + p9_level_51_d2  + p9_level_61_d2  + p9_level_71_d2 );

               Edit7.Text :=  INTTOSTR(b1_level_15_d3  + b1_level_11_d3  + b1_level_12_d3  + b1_level_43_d3  + b1_level_51_d3  + b1_level_61_d3  + b1_level_71_d3 +
                                       b7_level_15_d3  + b7_level_11_d3  + b7_level_12_d3  + b7_level_43_d3  + b7_level_51_d3  + b7_level_61_d3  + b7_level_71_d3 +
                                       p9_level_15_d3  + p9_level_11_d3  + p9_level_12_d3  + p9_level_43_d3  + p9_level_51_d3  + p9_level_61_d3  + p9_level_71_d3 );


               Edit8.Text  := INTTOSTR(b1_level_15_a + b7_level_15_a + p9_level_15_a);
               Edit9.Text  := INTTOSTR(b1_level_15_b + b7_level_15_b + p9_level_15_b);
               Edit10.Text := INTTOSTR(b1_level_15_c1 + b7_level_15_c1 + p9_level_15_c1);
               Edit11.Text := INTTOSTR(b1_level_15_c2 + b7_level_15_c2 + p9_level_15_c2);
               Edit12.Text := INTTOSTR(b1_level_15_d1 + b7_level_15_d1 + p9_level_15_d1);
               Edit13.Text := INTTOSTR(b1_level_15_d2 + b7_level_15_d2 + p9_level_15_d2);
               Edit14.Text := INTTOSTR(b1_level_15_d3 + b7_level_15_d3 + p9_level_15_d3);

               Edit15.Text := INTTOSTR(b1_level_12_a  + b7_level_12_a + p9_level_12_a );
               Edit16.Text := INTTOSTR(b1_level_12_b  + b7_level_12_b + p9_level_12_b );
               Edit17.Text := INTTOSTR(b1_level_12_c1 + b7_level_12_c1+ p9_level_12_c1);
               Edit18.Text := INTTOSTR(b1_level_12_c2 + b7_level_12_c2+ p9_level_12_c2);
               Edit19.Text := INTTOSTR(b1_level_12_d1 + b7_level_12_d1+ p9_level_12_d1);
               Edit20.Text := INTTOSTR(b1_level_12_d2 + b7_level_12_d2+ p9_level_12_d2);
               Edit21.Text := INTTOSTR(b1_level_12_d3 + b7_level_12_d3+ p9_level_12_d3);

               Edit22.Text := INTTOSTR(b1_level_11_a  + b7_level_11_a  + p9_level_11_a );
               Edit23.Text := INTTOSTR(b1_level_11_b  + b7_level_11_b  + p9_level_11_b );
               Edit24.Text := INTTOSTR(b1_level_11_c1 + b7_level_11_c1 + p9_level_11_c1);
               Edit25.Text := INTTOSTR(b1_level_11_c2 + b7_level_11_c2 + p9_level_11_c2);
               Edit26.Text := INTTOSTR(b1_level_11_d1 + b7_level_11_d1 + p9_level_11_d1);
               Edit27.Text := INTTOSTR(b1_level_11_d2 + b7_level_11_d2 + p9_level_11_d2);
               Edit28.Text := INTTOSTR(b1_level_11_d3 + b7_level_11_d3 + p9_level_11_d3);

               Edit29.Text := INTTOSTR(b1_level_43_a  + b7_level_43_a  + p9_level_43_a );
               Edit30.Text := INTTOSTR(b1_level_43_b  + b7_level_43_b  + p9_level_43_b );
               Edit31.Text := INTTOSTR(b1_level_43_c1 + b7_level_43_c1 + p9_level_43_c1);
               Edit32.Text := INTTOSTR(b1_level_43_c2 + b7_level_43_c2 + p9_level_43_c2);
               Edit33.Text := INTTOSTR(b1_level_43_d1 + b7_level_43_d1 + p9_level_43_d1);
               Edit34.Text := INTTOSTR(b1_level_43_d2 + b7_level_43_d2 + p9_level_43_d2);
               Edit35.Text := INTTOSTR(b1_level_43_d3 + b7_level_43_d3 + p9_level_43_d3);

               Edit36.Text := INTTOSTR(b1_level_71_a  + b7_level_71_a  + p9_level_71_a );
               Edit37.Text := INTTOSTR(b1_level_71_b  + b7_level_71_b  + p9_level_71_b );
               Edit38.Text := INTTOSTR(b1_level_71_c1 + b7_level_71_c1 + p9_level_71_c1);
               Edit39.Text := INTTOSTR(b1_level_71_c2 + b7_level_71_c2 + p9_level_71_c2);
               Edit40.Text := INTTOSTR(b1_level_71_d1 + b7_level_71_d1 + p9_level_71_d1);
               Edit41.Text := INTTOSTR(b1_level_71_d2 + b7_level_71_d2 + p9_level_71_d2);
               Edit42.Text := INTTOSTR(b1_level_71_d3 + b7_level_71_d3 + p9_level_71_d3);

               Edit43.Text := INTTOSTR(b1_level_61_a  + b7_level_61_a  + p9_level_61_a );
               Edit44.Text := INTTOSTR(b1_level_61_b  + b7_level_61_b  + p9_level_61_b );
               Edit45.Text := INTTOSTR(b1_level_61_c1 + b7_level_61_c1 + p9_level_61_c1);
               Edit46.Text := INTTOSTR(b1_level_61_c2 + b7_level_61_c2 + p9_level_61_c2);
               Edit47.Text := INTTOSTR(b1_level_61_d1 + b7_level_61_d1 + p9_level_61_d1);
               Edit48.Text := INTTOSTR(b1_level_61_d2 + b7_level_61_d2 + p9_level_61_d2);
               Edit49.Text := INTTOSTR(b1_level_61_d3 + b7_level_61_d3 + p9_level_61_d3);

               Edit50.Text := INTTOSTR(b1_level_51_a  + b7_level_51_a  + p9_level_51_a );
               Edit51.Text := INTTOSTR(b1_level_51_b  + b7_level_51_b  + p9_level_51_b );
               Edit52.Text := INTTOSTR(b1_level_51_c1 + b7_level_51_c1 + p9_level_51_c1);
               Edit53.Text := INTTOSTR(b1_level_51_c2 + b7_level_51_c2 + p9_level_51_c2);
               Edit54.Text := INTTOSTR(b1_level_51_d1 + b7_level_51_d1 + p9_level_51_d1);
               Edit55.Text := INTTOSTR(b1_level_51_d2 + b7_level_51_d2 + p9_level_51_d2);
               Edit56.Text := INTTOSTR(b1_level_51_d3 + b7_level_51_d3 + p9_level_51_d3);


             end;
         Next;
      End;

      Close;
   end; // end of With(qryGumgin);

   GrdMaster.Rows := index;
   GrdMaster.SetFocus;
end;

Function TfrmLD8Q11.HangmokFind( vTemp , vHangmok : String) : String;
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

{procedure TfrmLD8Q10.UP_Insert_2012(iTemp : Integer; sHangmok, sGubn_nosin, sGubn_adult, sGubn_life, sGubn_hulgum, sCod_Jisa : String);
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

procedure TfrmLD8Q10.UP_Insert_2013(iTemp : Integer; sHangmok, sCod_jangbi, sCod_hul, sGubn_hulgum, sCod_Jisa : String);
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
    }
procedure TfrmLD8Q11.SBut_ExcelClick(Sender: TObject);
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
