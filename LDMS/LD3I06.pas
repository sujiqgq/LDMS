//==============================================================================
// 프로그램 설명 : 수원지사 혈액결과 입력
// 시스템        : 통합검진
// 수정일자      : 2005.11.29
// 수정자        : 유동구
// 수정내용      : 수원지사 분석장비 업로드...
// 참고사항      :
//==============================================================================
unit LD3I06;



interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;



type
  TfrmLD3I06 = class(TfrmSingle)
    OpenDialog: TOpenDialog;
    grdmaster: TtsGrid;
    GroupBox1: TGroupBox;
    gkPie: TGauge;
    pnlJisa: TPanel;
    Btn_Insert: TBitBtn;
    pnlMaster: TPanel;
    btnDate: TSpeedButton;
    btnDESC_DEPT: TSpeedButton;
    Panel2: TPanel;
    mskDate: TDateEdit;
    btnStart: TBitBtn;
    Panel5: TPanel;
    edtFile: TEdit;
    Panel4: TPanel;
    valNum_bunju: TValEdit;
    RGrp_Hulgum: TRadioGroup;
    Panel3: TPanel;
    Panel6: TPanel;
    qrdSList: TtsGrid;
    qryGumgin: TQuery;
    qryGulkwa: TQuery;
    qryU_Gulkwa: TQuery;
    valChange: TValEdit;
  procedure FormCreate(Sender: TObject);
  procedure grdMasterCellLoaded(Sender: TObject; DataCol,
            DataRow: Integer; var Value: Variant);
    procedure btnDESC_DEPTClick(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure Btn_InsertClick(Sender: TObject);
    procedure qrdSListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);

 private
    { Private declarations }
    UV_sJisa : String;
    
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010, UV_v011, UV_v012, UV_v013 : Variant;

    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005, UV_vS006, UV_vS007, UV_vS008, UV_vS009, UV_vS010, UV_vS011, UV_vS012, UV_vS013 : Variant;

    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_SVarMemSet(iLength : Integer);

  public
    { Public declarations }
    UV_sValue  : array[1..8] of String;

  end;


var
  frmLD3I06 : TfrmLD3I06;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}
  

procedure TfrmLD3I06.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;
end;

procedure TfrmLD3I06.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
      UV_v013 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_v001, iLength);
      VarArrayRedim(UV_v002, iLength);
      VarArrayRedim(UV_v003, iLength);
      VarArrayRedim(UV_v004, iLength);
      VarArrayRedim(UV_v005, iLength);
      VarArrayRedim(UV_v006, iLength);
      VarArrayRedim(UV_v007, iLength);
      VarArrayRedim(UV_v008, iLength);
      VarArrayRedim(UV_v009, iLength);
      VarArrayRedim(UV_v010, iLength);
      VarArrayRedim(UV_v011, iLength);
      VarArrayRedim(UV_v012, iLength);
      VarArrayRedim(UV_v013, iLength);
   end;
end;

procedure TfrmLD3I06.UP_SVarMemSet(iLength : Integer);
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
      UV_vS007 := VarArrayCreate([0, 0], varOleStr);
      UV_vS008 := VarArrayCreate([0, 0], varOleStr);
      UV_vS009 := VarArrayCreate([0, 0], varOleStr);
      UV_vS010 := VarArrayCreate([0, 0], varOleStr);
      UV_vS011 := VarArrayCreate([0, 0], varOleStr);
      UV_vS012 := VarArrayCreate([0, 0], varOleStr);
      UV_vS013 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vS001, iLength);
      VarArrayRedim(UV_vS002, iLength);
      VarArrayRedim(UV_vS003, iLength);
      VarArrayRedim(UV_vS004, iLength);
      VarArrayRedim(UV_vS005, iLength);
      VarArrayRedim(UV_vS006, iLength);
      VarArrayRedim(UV_vS007, iLength);
      VarArrayRedim(UV_vS008, iLength);
      VarArrayRedim(UV_vS009, iLength);
      VarArrayRedim(UV_vS010, iLength);
      VarArrayRedim(UV_vS011, iLength);
      VarArrayRedim(UV_vS012, iLength);
      VarArrayRedim(UV_vS013, iLength);
   end;
end;

procedure TfrmLD3I06.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_v001[DataRow - 1];
      2 : Value := UV_v002[DataRow - 1];
      3 : Value := UV_v003[DataRow - 1];
      4 : Value := UV_v004[DataRow - 1];
      5 : Value := UV_v005[DataRow - 1];
      6 : Value := UV_v006[DataRow - 1];
      7 : Value := UV_v007[DataRow - 1];
      8 : Value := UV_v008[DataRow - 1];
      9 : Value := UV_v009[DataRow - 1];
     10 : Value := UV_v010[DataRow - 1];
     11 : Value := UV_v011[DataRow - 1];
     12 : Value := UV_v012[DataRow - 1];
     13 : Value := UV_v013[DataRow - 1];
   end;
end;

procedure TfrmLD3I06.btnDESC_DEPTClick(Sender: TObject);
begin
   inherited;
   if OpenDialog.Execute = False then
      exit;
   edtFile.Text := OpenDialog.FileName;
end;

procedure TfrmLD3I06.btnStartClick(Sender: TObject);
var
   fFile : Textfile;
   S, Temp, Te_temp, sDate, sJumin : string;
   i , k, index, Sindex : integer;
   bDataChk : Boolean;

begin
   inherited;
   grdMaster.Rows := 0;  qrdSList.RowS := 0;
   Index := 0; Sindex := 0;
   i:= 0; k := 0;
   bDataChk := True;

   if not GF_MsgBox('Confirmation', '불러오기를 시작 하시겠습니까.?') then exit;

   try
      // file open
      assignfile(fFile, OpenDialog.FileName);
      reset(fFile);

      // file 끝부분 확인
      while not EOF(fFile) do
      begin
         // 한줄씩 나눔.
         ReadLn(fFile, S);

         Temp := copy(S,1,length(S));

         // type이 N 일경우 만 입력
         if copy(Temp,1,1) = 'N' then
         begin
             Temp[pos(';',Temp)] := '+';
             // 검진일자..
             sDate :=  copy(Trim(copy(Temp,3,  pos(';',Temp)-3)),1,8);
             k:= pos(';',Temp);
             Temp[pos(';',Temp)] := '+';
             //주민번호..
             sJumin :=  Trim(copy(Temp,k+1,pos(';',Temp)-(k+1)));

             qryGumgin.Active := False;
             qryGumgin.ParamByName('cod_jisa').AsString    := UV_sJisa;
             qryGumgin.ParamByName('num_jumin').AsString   := sJumin;
             qryGumgin.ParamByName('dat_gmgn').AsString    := sDate;
             case RGrp_Hulgum.ItemIndex of
              0 : qryGumgin.ParamByName('gubn_hulgum').AsString := '2';
              1 : qryGumgin.ParamByName('gubn_hulgum').AsString := '3';
             end;
             qryGumgin.Active := True;

             if (qryGumgin.RecordCount > 0) and (pos('REJECT',Temp) = 0) and (pos('*INFO*',Temp) = 0) then
             begin
                // 분주가 존재하는 데이터...
                UP_VarMemSet(index);
                bDataChk := True;
                UV_v001[index] := ''; UV_v002[index] := ''; UV_v003[index] := ''; UV_v004[index] := ''; UV_v005[index] := '';
                UV_v006[index] := ''; UV_v007[index] := ''; UV_v008[index] := ''; UV_v009[index] := ''; UV_v010[index] := '';
                UV_v011[index] := ''; UV_v012[index] := ''; UV_v013[index] := '';

                UV_v001[index] := qryGumgin.FieldByName('cod_bjjs').AsString;
                UV_v002[index] := qryGumgin.FieldByName('dat_bunju').AsString;
                UV_v003[index] := qryGumgin.FieldByName('num_bunju').AsString;
                UV_v004[index] := sDate;
                UV_v005[index] := sJumin;
             end
             else
             begin
                // 분주가 존재하는 데이터...
                bDataChk := False;
                UP_SVarMemSet(Sindex);
                UV_vS001[Sindex] := ''; UV_vS002[Sindex] := ''; UV_vS003[Sindex] := ''; UV_vS004[Sindex] := ''; UV_vS005[Sindex] := '';
                UV_vS006[Sindex] := ''; UV_vS007[Sindex] := ''; UV_vS008[Sindex] := ''; UV_vS009[Sindex] := ''; UV_vS010[Sindex] := '';
                UV_vS011[Sindex] := ''; UV_vS012[Sindex] := ''; UV_vS013[Sindex] := '';

                UV_vS001[Sindex] := '';
                UV_vS002[Sindex] := '';
                UV_vS003[Sindex] := '';
                UV_vS004[Sindex] := sDate;
                UV_vS005[Sindex] := sJumin;

             end;
             qryGumgin.Active := False;

             // GOT,GPT,CHO,GLU,GGT 입력
             if pos('GOT',TEMP) <> 0 then
             begin
                Te_temp := copy(Temp,POS('GOT',TEMP),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v006[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS006[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('GPT',Temp) <> 0 then
             begin
                Te_temp := copy(Temp,POS('GPT',Temp),length(temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v007[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS007[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('CHO',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('CHO',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v008[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS008[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('GLU',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('GLU',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v009[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS009[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('GGT',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('GGT',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v010[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS010[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('HDL',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('HDL',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v011[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS011[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('TG',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('TG',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v012[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1)
                   else
                      UV_vS012[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) -1);
                end;
             end;

             if pos('CREA',Temp) <> 0 then
             begin
                Te_temp := copy(TEMP,POS('CREA',Temp),length(Temp));
                k := pos(';',Te_temp);
                Te_temp[POS(';',Te_temp)] := '+';
                if (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> 'REJECT') and (copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)) <> '*INFO*') then
                begin
                   if bDataChk then
                      UV_v013[index]   :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) +1)
                   else
                      UV_vS013[Sindex] :=  copy(copy(Te_temp,k+1,POS(';',Te_temp)-(k+1)), 1, pos('.', copy(Te_temp,k+1,POS(';',Te_temp)-(k+1))) +1);
                end;
             end;
         end
         else
         begin
            // type 이 N이 아닌경우 증가된 index값 감소
            if bDataChk then index := index -1
            else Sindex := Sindex -1;
         end;

         // 다음 줄을 찾기위해
         i:=pos(#10,s);
         s[pos(#10,s)]:='+';
         if bDataChk then inc(index)
         else             inc(Sindex);
      end;

      closefile(fFile);
   except
      on E : EDBEngineError do
      begin
         closefile(fFile);
         showmessage('다음 파일을 열수 없습니다.');
      end;
   end;

   // display
   grdMaster.Rows := Index;
   qrdSList.RowS  := Sindex;
end;

procedure TfrmLD3I06.Btn_InsertClick(Sender: TObject);
var
    sGlkwa, sValue : String;
    I, iTemp : Integer;
    eRslt, e25, e26, e28 : Extended;
begin
   inherited;
   if grdmaster.Rows = 0 then
   begin
      showmessage('데이터가 존재하지 않습니다.');
      exit;
   end;
   if not GF_MsgBox('Confirmation', 'Transfer를 시작 하시겠습니까.?') then exit;
   gkPie.MaxValue := grdmaster.Rows;
   gkPie.Progress := 0;
   
   for I := 1 to grdmaster.Rows do
   begin
      gkPie.Progress := gkPie.Progress + 1;
      // 초기결과값.
      with qryGulkwa do
      begin
         Active := False;
         ParamByName('cod_bjjs').AsString  := grdmaster.Cell[1,I];
         ParamByName('dat_bunju').AsString := grdmaster.Cell[2,I];
         ParamByName('num_bunju').AsString := grdmaster.Cell[3,I];
         ParamByName('gubn_part').AsString := '01';
         Active := True;
         if RecordCount > 0 then
         begin
            GP_query_log(qryGulkwa, 'LD3I06', 'Q', 'Y');
            UV_fGulkwa := '';
            UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
            sGlkwa := UV_fGulkwa;
            Active := False;
            if sGlkwa = '' then
            begin
               for iTemp := 1 to 10 do
               begin
                  sGlkwa := sGlkwa + '                                                            ';
               end; // end of for;
            end; // end of if
         end; // end of if;
      end; // end of with;

      // GOT..
      sGlkwa := Copy(sGlkwa, 1, 48) + GF_RPad(grdmaster.Cell[ 6,I], 6, ' ') + Copy(sGlkwa, 55, Length(sGlkwa));

      // GPT..
      sGlkwa := Copy(sGlkwa, 1, 54) + GF_RPad(grdmaster.Cell[ 7,I], 6, ' ') + Copy(sGlkwa, 61, Length(sGlkwa));

      // CHO..
      sGlkwa := Copy(sGlkwa, 1, 120) + GF_RPad(grdmaster.Cell[8,I], 6, ' ') + Copy(sGlkwa, 127, Length(sGlkwa));

      // GLU..
      sGlkwa := Copy(sGlkwa, 1, 156) + GF_RPad(grdmaster.Cell[9,I], 6, ' ') + Copy(sGlkwa, 163, Length(sGlkwa));

      // GGT..
      sGlkwa := Copy(sGlkwa, 1, 60) + GF_RPad(grdmaster.Cell[10,I], 6, ' ') + Copy(sGlkwa, 67, Length(sGlkwa));

      // HDL-CHO..
      sGlkwa := Copy(sGlkwa, 1, 126) + GF_RPad(grdmaster.Cell[11,I], 6, ' ') + Copy(sGlkwa, 133, Length(sGlkwa));

      // TG..
      sGlkwa := Copy(sGlkwa, 1, 138) + GF_RPad(grdmaster.Cell[12,I], 6, ' ') + Copy(sGlkwa, 145, Length(sGlkwa));

      // CREA..
      sGlkwa := Copy(sGlkwa, 1, 186) + GF_RPad(grdmaster.Cell[13,I], 6, ' ') + Copy(sGlkwa, 193, Length(sGlkwa));

      if (grdmaster.Cell[8,I] <> '') and (grdmaster.Cell[11,I] <> '') and (grdmaster.Cell[12,I] <> '') then
      begin
         e25 := GF_StrToFloat(grdmaster.Cell[8,I]);
         e26 := GF_StrToFloat(grdmaster.Cell[11,I]);
         e28 := GF_StrToFloat(grdmaster.Cell[12,I]);

         eRslt := e25 - ((e28/5) + e26);
         if Pos('.', FloatToStr(eRslt)) = 0 then eRslt := StrToFloat(FloatToStr(eRslt) + '.0');                      //2007.02.27 철.

         valChange.Scale := 1;
         valChange.Value := eRslt;
         sValue := valChange.Text;

         // LDL..
         sGlkwa := Copy(sGlkwa, 1, 132) + GF_RPad(sValue, 6, ' ') + Copy(sGlkwa, 139, Length(sGlkwa));
      end;

      sGlkwa := 'A' + sGlkwa;
      sGlkwa := Trim(sGlkwa);
      sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

      // DB 작업
      try
         UV_fGulkwa1 := '';
         UV_fGulkwa2 := '';
         UV_fGulkwa3 := '';
         GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

         qryU_Gulkwa.ParamByName('COD_BJJS').AsString    := grdmaster.Cell[1,I];
         qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString   := grdmaster.Cell[2,I];
         qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString   := grdmaster.Cell[3,I];
         qryU_Gulkwa.ParamByName('GUBN_PART').AsString   := '01';
         qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString := UV_fGulkwa1;
         qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString := UV_fGulkwa2;
         qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString := UV_fGulkwa3;
         qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString  := GV_sToday;
         qryU_Gulkwa.ParamByName('COD_UPDATE').AsString  := GV_sUserId;

         qryU_Gulkwa.Execsql;
         GP_query_log(qryU_Gulkwa, 'LD3I06', 'Q', 'Y');
      except
         on E : EDBEngineError do
         begin
            GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
            exit;
         end; // end of except;
      end; // end of try;
   end; // end of for;
end;

procedure TfrmLD3I06.qrdSListCellLoaded(Sender: TObject; DataCol,
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
      7 : Value := UV_vS007[DataRow - 1];
      8 : Value := UV_vS008[DataRow - 1];
      9 : Value := UV_vS009[DataRow - 1];
     10 : Value := UV_vS010[DataRow - 1];
     11 : Value := UV_vS011[DataRow - 1];
     12 : Value := UV_vS012[DataRow - 1];
     13 : Value := UV_vS013[DataRow - 1];
   end;
end;

end.






