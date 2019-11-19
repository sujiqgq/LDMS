//==============================================================================
// ���α׷� ���� : �����׸� ���װ�� �Է�
// �ý���        : ���հ���
// ��������      : 2015.11.25
// ������        : ������
// ��������      : ���װ�� ���ε�...
// �������      : ���� ������ܰ˻�������1500972//2015.12 �Ѵ޸� ���..
//==============================================================================

unit LD3I10;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;

type
  TfrmLD3I10 = class(TfrmSingle)
    OpenDialog: TOpenDialog;
    GroupBox1: TGroupBox;
    gkPie: TGauge;
    pnlJisa: TPanel;
    Btn_Insert: TBitBtn;
    pnlMaster: TPanel;
    btnDESC_DEPT: TSpeedButton;
    btnStart: TBitBtn;
    Panel5: TPanel;
    edtFile: TEdit;
    Panel3: TPanel;
    Panel6: TPanel;
    qrdSPList: TtsGrid;
    qryGulkwa: TQuery;
    qryU_Gulkwa: TQuery;
    TG_gumgin: TStringGrid;
    qrdSList: TtsGrid;
    Panel2: TPanel;
    qryProfile: TQuery;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
  procedure FormCreate(Sender: TObject);
    procedure btnDESC_DEPTClick(Sender: TObject);
    procedure Btn_InsertClick(Sender: TObject);
    procedure qrdSListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure qrdSPListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnStartClick(Sender: TObject);

 private
    { Private declarations }
    UV_sJisa : String;
    Excel, Worksheet  : Variant;

    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005, UV_vS006 : Variant;
    UV_vSP001, UV_vSP002, UV_vSP003, UV_vSP004, UV_vSP005, UV_vSP006 : Variant;
    
    procedure UP_SVarMemSet(iLength : Integer);
    procedure UP_SPVarMemSet(iLength : Integer);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    { Public declarations }
    UV_sValue  : array[1..8] of String;
  end;


var
  frmLD3I10 : TfrmLD3I10;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I10.FormCreate(Sender: TObject);
begin
   inherited;

   // ��ü(00)�ϰ�� ���縦 �������� �۾�
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I10.UP_SVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vS001 := VarArrayCreate([0, 0], varOleStr);
      UV_vS002 := VarArrayCreate([0, 0], varOleStr);
      UV_vS003 := VarArrayCreate([0, 0], varOleStr);
      UV_vS004 := VarArrayCreate([0, 0], varOleStr);
      UV_vS005 := VarArrayCreate([0, 0], varOleStr);
      UV_vS006 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vS001, iLength);
      VarArrayRedim(UV_vS002, iLength);
      VarArrayRedim(UV_vS003, iLength);
      VarArrayRedim(UV_vS004, iLength);
      VarArrayRedim(UV_vS005, iLength);
      VarArrayRedim(UV_vS006, iLength);
   end;
end;
procedure TfrmLD3I10.qrdSListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vS001[DataRow - 1];
      2 : Value := UV_vS002[DataRow - 1];
      3 : Value := UV_vS003[DataRow - 1];
      4 : Value := UV_vS004[DataRow - 1];
      5 : Value := UV_vS005[DataRow - 1];
      6 : Value := UV_vS006[DataRow - 1];
   end;
end;
procedure TfrmLD3I10.UP_SPVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vSP001 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP002 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP003 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP004 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP005 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP006 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vSP001, iLength);
      VarArrayRedim(UV_vSP002, iLength);
      VarArrayRedim(UV_vSP003, iLength);
      VarArrayRedim(UV_vSP004, iLength);
      VarArrayRedim(UV_vSP005, iLength);
      VarArrayRedim(UV_vSP006, iLength);
   end;
end;
procedure TfrmLD3I10.qrdSPListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vSP001[DataRow - 1];
      2 : Value := UV_vSP002[DataRow - 1];
      3 : Value := UV_vSP003[DataRow - 1];
      4 : Value := UV_vSP004[DataRow - 1];
      5 : Value := UV_vSP005[DataRow - 1];
      6 : Value := UV_vSP006[DataRow - 1];
   end;
end;

procedure TfrmLD3I10.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I10.btnStartClick(Sender: TObject);
var
  icol, sCol, iRow, sRow : Integer;
  sLine, FileName, sSheetName : String;
begin
   Screen.Cursor := crHourGlass;
   sCol := 0;  sRow := 0;
   try
      Excel:= CreateOleObject('Excel.Application');
      Excel.Workbooks.Open(OpenDialog.FileName);
      if Trim(sSheetName) = '' Then
        Worksheet := Excel.Workbooks[1].WorkSheetS[Excel.WorkBooks[1].Worksheets[1].Name]
      else
        Worksheet := Excel.Workbooks[1].WorkSheetS[SSheetName];
   except
      MessageDlg('Excel�� ���������� ������� �ʾҽ��ϴ�.' + Chr(13) +
                         'ȯ�漳���� �ٽ� �Ͻñ� �ٶ��ϴ�.',mtError,[mbOk], 0);
   end;
   For iRow := 1 to 65000 do
   begin
      if WorkSheet.Cells[iRow,1].Text <> '' then
         Inc(sRow)
      else
         Break;
   end;
   For iCol := 1 to 20 do
   begin
      if WorkSheet.Cells[1,iCol].Text <> '' then
         Inc(sCol)
      else
         Break;
   end;

   TG_gumgin.RowCount := sRow;
   TG_gumgin.ColCount := sCol;
   
   For iRow := 0 to sRow do
   begin
      For iCol := 0 to sCol do
         TG_gumgin.Cells[iCol,iRow] := Trim(WorkSheet.Cells[iRow + 1,iCol + 1].Text);
   end;

   if not VarIsEmpty(Excel) then
     Excel.Quit;
   Screen.Cursor := crDefault;
end;


procedure TfrmLD3I10.Btn_InsertClick(Sender: TObject);
var
    sGlkwa, sCod_HM, sSuchi : String;
    I, iTemp, sIndex, spIndex: Integer;
begin
   inherited;
   sIndex := 0;
   spIndex := 0;
   if TG_gumgin.RowCount = 0 then
   begin
      showmessage('�����Ͱ� �������� �ʽ��ϴ�.');
      exit;
   end;
   if not GF_MsgBox('Confirmation', 'Transfer�� ���� �Ͻðڽ��ϱ�.?') then exit;
   gkPie.MaxValue := TG_gumgin.RowCount -1;
   gkPie.Progress := 0;

   for I := 1 to TG_gumgin.RowCount -1 do
   begin
      gkPie.Progress := gkPie.Progress + 1;
      sSuchi := '';
      // DB �۾�
      try
         // �ʱ�����.
         with qryGulkwa do
         begin
              Active := False;
              ParamByName('cod_bjjs').AsString   := '15';
              qryGulkwa.ParamByName('dat_bunju').AsString  := TG_gumgin.Cells[0,I];
              qryGulkwa.ParamByName('num_bunju').AsString  := TG_gumgin.Cells[3,I];
              qryGulkwa.ParamByName('desc_name').AsString  := TG_gumgin.Cells[6,I];
              Active := True;
              if RecordCount > 0 then
              begin
                   sCod_HM := UF_hangmokList;

                   UP_SVarMemSet(sIndex);

                   UV_vS001[sIndex] := TG_gumgin.Cells[0,I];
                   UV_vS002[sIndex] := TG_gumgin.Cells[3,I];
                   UV_vS003[sIndex] := TG_gumgin.Cells[4,I] + '-' + TG_gumgin.Cells[5,I];
                   UV_vS004[sIndex] := TG_gumgin.Cells[6,I];
                   UV_vS005[sIndex] := TG_gumgin.Cells[7,I];
                   UV_vS006[sIndex] := TG_gumgin.Cells[10,I];

                   UV_fGulkwa := '';
                   UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
                   UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
                   UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
                   GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                   sGlkwa := UV_fGulkwa;
                   //Active := False;

                   if sGlkwa = '' then
                   begin
                        for iTemp := 1 to 10 do
                            begin
                                 sGlkwa := sGlkwa + '                                                            ';
                            end; // end of for;
                   end; // end of if

                   if (pos('Positive(', Trim(TG_gumgin.Cells[10,I])) > 0) or (pos('Negative(', Trim(TG_gumgin.Cells[10,I])) > 0) then
                   begin
                       if (pos('Positive(', Trim(TG_gumgin.Cells[10,I])) > 0) then
                          TG_gumgin.Cells[10,I] := StringReplace(Trim(TG_gumgin.Cells[10,I]), 'Positive(', '',[rfReplaceAll, rfIgnoreCase])
                       else if (pos('Negative(', Trim(TG_gumgin.Cells[10,I])) > 0) then
                          TG_gumgin.Cells[10,I] := StringReplace(Trim(TG_gumgin.Cells[10,I]), 'Negative(', '',[rfReplaceAll, rfIgnoreCase]);

                       TG_gumgin.Cells[10,I] := StringReplace(Trim(TG_gumgin.Cells[10,I]), ')', '',[rfReplaceAll, rfIgnoreCase]);
                   end;

                   // 1.S008
                   if pos('S008', sCod_HM) > 0 then
                   begin
                        if StrToFloat(Trim(TG_gumgin.Cells[10,I])) < 10.0 then sSuchi := '����'
                        else if (StrToFloat(Trim(TG_gumgin.Cells[10,I])) >= 10.0) and
                                (StrToFloat(Trim(TG_gumgin.Cells[10,I])) <= 30.0) then sSuchi := '��缺'
                        else if StrToFloat(Trim(TG_gumgin.Cells[10,I])) > 30.0 then sSuchi := '�缺';

                        sGlkwa := Copy(sGlkwa, 1, 36) + GF_RPad(sSuchi, 6, ' ') +
                                  Copy(sGlkwa, 43, Length(sGlkwa));
                   end;
                   // 2.S091
                   if pos('S091', sCod_HM) > 0 then
                        sGlkwa := Copy(sGlkwa, 1, 180) + GF_RPad(Trim(TG_gumgin.Cells[10,I]), 6, ' ') +
                                  Copy(sGlkwa, 187, Length(sGlkwa));

                   sGlkwa := 'A' + sGlkwa;
                   sGlkwa := Trim(sGlkwa);
                   sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

                   UV_fGulkwa1 := '';
                   UV_fGulkwa2 := '';
                   UV_fGulkwa3 := '';
                   GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

                   qryU_Gulkwa.ParamByName('COD_BJJS').AsString     := '15';
                   qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString    := TG_gumgin.Cells[0,I];
                   qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString    := TG_gumgin.Cells[3,I];
                   qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString  := UV_fGulkwa1;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString  := UV_fGulkwa2;
                   qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString  := UV_fGulkwa3;
                   qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString   := GV_sToday;
                   qryU_Gulkwa.ParamByName('COD_UPDATE').AsString   := GV_sUserId;
                   qryU_Gulkwa.Execsql;
                   GP_query_log(qryU_Gulkwa, 'LD3I10', 'Q', 'Y');
                   Inc(sIndex);

              end  //end od if;
              else
              begin
                   UP_SPVarMemSet(spIndex);

                   UV_vSP001[spIndex] := TG_gumgin.Cells[0,I];
                   UV_vSP002[spIndex] := TG_gumgin.Cells[3,I];
                   UV_vSP003[spIndex] := TG_gumgin.Cells[4,I] + '-' + TG_gumgin.Cells[5,I];
                   UV_vSP004[spIndex] := TG_gumgin.Cells[6,I];
                   UV_vSP005[spIndex] := TG_gumgin.Cells[7,I];
                   UV_vSP006[spIndex] := TG_gumgin.Cells[10,I];

                   Inc(spIndex);
              end; // end of else;
         end; // end of with;
         except
               on E : EDBEngineError do
                  begin
                       GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
                       exit;
                  end; // end of except;
         end; // end of try;

   end; // end of for;

   qrdSList.Rows  := sIndex;
   qrdSPList.Rows  := spIndex;

end;

function TfrmLD3I10.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, j, k : integer;
begin
   result := ''; sTemp := '';

   with qryPf_hm do
   begin
      // 1. ���׿� ���� �˻��׸� ����
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

      // 2. ���翡 ���� �˻��׸� ����
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

      // 3. ��� ���� �˻��׸� ����
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

   // 3. �߰��ڵ忡 ���� �˻��׸� ����
   sTemp := sTemp + qryGulkwa.FieldByName('Cod_chuga').AsString;;

   // 4. �׸��ڵ忡 ���� �˻��׸� ����
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

function  TfrmLD3I10.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

end.






