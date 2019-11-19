//==============================================================================
// 프로그램 설명 : [광주]Urin - Miditron M (Interface & Transfer)
// 시스템        : 통합검진
// 수정일자      : 2007.03.09
// 수정자        : 유동구
// 수정내용      : 
// 참고사항      :
//==============================================================================
unit LD5I02;

interface

uses
  Windows, Messages,  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges, CPort, CPortCtl;



type
  TfrmLD5I02 = class(TfrmSingle)
    grdmaster: TtsGrid;
    pnlMaster: TPanel;
    btnStart: TBitBtn;
    Panel3: TPanel;
    ComPort: TComPort;
    btnSetting: TBitBtn;
    Panel2: TPanel;
    Panel6: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Cb_Port: TComboBox;
    Bevel2: TBevel;
    ComLed1: TComLed;
    Label2: TLabel;
    ComLed2: TComLed;
    Label3: TLabel;
    ComLed3: TComLed;
    Label4: TLabel;
    ComLed4: TComLed;
    ComLed5: TComLed;
    ComLed6: TComLed;
    Label5: TLabel;
    Label1: TLabel;
    Label6: TLabel;
    ComLed7: TComLed;
    Label7: TLabel;
    qrdSList: TtsGrid;
    qrdSPList: TtsGrid;
    Panel7: TPanel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    Animate1: TAnimate;
    Label_cmd: TLabel;
    Gauge: TGauge;
    Image1: TImage;
    Panel8: TPanel;
    Animate2: TAnimate;
    qryHangmok: TQuery;
    qryGumgin: TQuery;
    qryGulkwa: TQuery;
    qryU_Gulkwa: TQuery;
    qryTkgum: TQuery;
    qryPf_hm: TQuery;
    qryProfileG: TQuery;
    qryNo_hangmok: TQuery;
    SBtn_Start: TSpeedButton;
    Panel9: TPanel;
    Edt_No: TEdit;
  procedure FormCreate(Sender: TObject);
  procedure grdMasterCellLoaded(Sender: TObject; DataCol,
            DataRow: Integer; var Value: Variant);
    procedure btnStartClick(Sender: TObject);
    procedure qrdSListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnSettingClick(Sender: TObject);
    procedure ComPortAfterClose(Sender: TObject);
    procedure ComPortAfterOpen(Sender: TObject);
    procedure ComPortRxChar(Sender: TObject; Count: Integer);
    procedure Cb_PortChange(Sender: TObject);
    procedure qrdSPListCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure SBtn_StartClick(Sender: TObject);

 private
    { Private declarations }
    UV_sJisa, UV_List : String;
    UV_index : Integer;
    
    //자료Display
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010,
    UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016 : Variant;

    //등록된 검진자
    UV_vS001, UV_vS002, UV_vS003, UV_vS004, UV_vS005, UV_vS006, UV_vS007, UV_vS008, UV_vS009, UV_vS010 : Variant;

    //등록 확인이 필요한 검진자
    UV_vSP001, UV_vSP002, UV_vSP003: Variant;

    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_SVarMemSet(iLength : Integer);
    procedure UP_SPVarMemSet(iLength : Integer);
    procedure UP_Clear;
    procedure UP_SClear(iTemp : Integer);
    procedure UP_SPClear(iTemp : Integer);

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  public
    function  UF_HagmokSetting(iSIndex, sil : Integer; var sValue : String; var iStart, iEnd : Integer) : String;
    { Public declarations }

  end;


var
  frmLD5I02 : TfrmLD5I02;

implementation

uses ZZFunc, ZZMenu, MdmsFunc, ZZComm;

{$R *.DFM}


procedure TfrmLD5I02.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0); UP_SVarMemSet(0); UP_SPVarMemSet(0);
   
   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

        if ComPort.Port = 'COM1' then Cb_Port.ItemIndex := 0
   else if ComPort.Port = 'COM2' then Cb_Port.ItemIndex := 1
   else if ComPort.Port = 'COM3' then Cb_Port.ItemIndex := 2
   else if ComPort.Port = 'COM4' then Cb_Port.ItemIndex := 3
   else if ComPort.Port = 'COM5' then Cb_Port.ItemIndex := 4;
end;

procedure TfrmLD5I02.UP_VarMemSet(iLength : Integer);
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
      UV_v014 := VarArrayCreate([0, 0], varOleStr);
      UV_v015 := VarArrayCreate([0, 0], varOleStr);
      UV_v016 := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_v014, iLength);
      VarArrayRedim(UV_v015, iLength);
      VarArrayRedim(UV_v016, iLength);
   end;
end;

procedure TfrmLD5I02.UP_SVarMemSet(iLength : Integer);
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
   end;
end;

procedure TfrmLD5I02.UP_SPVarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vSP001 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP002 := VarArrayCreate([0, 0], varOleStr);
      UV_vSP003 := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vSP001, iLength);
      VarArrayRedim(UV_vSP002, iLength);
      VarArrayRedim(UV_vSP003, iLength);
   end;
end;

procedure TfrmLD5I02.grdMasterCellLoaded(Sender: TObject; DataCol,
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
     14 : Value := UV_v014[DataRow - 1];
     15 : Value := UV_v015[DataRow - 1];
     16 : Value := UV_v016[DataRow - 1];
   end;
end;

procedure TfrmLD5I02.btnStartClick(Sender: TObject);
begin
   inherited;
   if mskDate.Text = '' then
   begin
      ShowMessage('검진일자를 입력하여 주십시요!!');
      exit;
   end;

   if ComPort.Connected then
   begin
      if grdmaster.Rows > 0 then SBtn_Start.Enabled := True;
      ComPort.Close;
   end
   else
   begin
      SBtn_Start.Enabled := False;
      grdMaster.Rows := 0;
      UV_List := ''; UV_index := 0;
      ComPort.Open;
   end;
end;

procedure TfrmLD5I02.qrdSListCellLoaded(Sender: TObject; DataCol,
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
   end;
end;

procedure TfrmLD5I02.btnSettingClick(Sender: TObject);
begin
   ComPort.ShowSetupDialog;
end;

procedure TfrmLD5I02.ComPortAfterClose(Sender: TObject);
begin
   if btnStart <> nil then
     btnStart.Caption := 'Open';
end;

procedure TfrmLD5I02.ComPortAfterOpen(Sender: TObject);
begin
  btnStart.Caption := 'Close';
end;

procedure TfrmLD5I02.ComPortRxChar(Sender: TObject; Count: Integer);
var
  Str, sList, sTemp : String;
begin
   sList := '';

   ComPort.ReadStr(Str, Count);

   UV_List := UV_List + Str;

   while pos(#3#2,UV_List) > 0 do
   begin
      sList := copy(UV_List,1,pos(#3#2,UV_List) - 5);
      UV_List  := sList + copy(UV_List,pos(#3#2,UV_List) + 4, length(UV_List));
   end;

   if (Pos(#3,UV_List) > 0) then
   begin
      UP_VarMemSet(UV_index);
      UP_Clear;

      sList := Trim(UV_List);

      //Seq No..
      UV_v001[UV_index] := IntToStr(UV_index + 1);
      //검진일자 획득..
      UV_v002[UV_index] := mskDate.Text;
      //CASSPOS No..
      UV_v003[UV_index] := Trim(copy(sList,Pos('CASSPOS',sList) + 8,7));
      //혈액샘플No. 획득..
      if pos('-',Trim(copy(sList,Pos('ID1',sList)+4,11))) > 0 then
         UV_v004[UV_index] := '-----------'
      else
         UV_v004[UV_index] := IntToStr(StrToInt(Trim(copy(sList,Pos('ID1',sList)+4,11))) + StrToInt(Edt_No.Text));

      //------------------------------------------------------------------------
      //백혈구 획득..
      if  Pos('WBC',sList) > 0 then
         UV_v005[UV_index] := Trim(copy(sList,Pos('WBC',sList) + 3,6));
      //적혈구 획득..
      if  Pos('RBC',sList) > 0 then
         UV_v006[UV_index] := Trim(copy(sList,Pos('RBC',sList) + 3,6));
      //헤모글로빈 획득..
      if  Pos('HGB',sList) > 0 then
         UV_v007[UV_index] := Trim(copy(sList,Pos('HGB',sList) + 3,6));
      //혈구용적치 획득..
      if  Pos('HCT',sList) > 0 then
         UV_v008[UV_index] := Trim(copy(sList,Pos('HCT',sList) + 3,6));
      //MCV 획득..
      if  Pos('MCV',sList) > 0 then
         UV_v009[UV_index] := Trim(copy(sList,Pos('MCV',sList) + 3,6));
      //MCH 획득..
      if  Pos('MCH',sList) > 0 then
         UV_v010[UV_index] := Trim(copy(sList,Pos('MCH',sList) + 3,6));
      //MCHC 획득..
      if  Pos('MCHC',sList) > 0 then
         UV_v011[UV_index] := Trim(copy(sList,Pos('MCHC',sList) + 4,6));
      //RDW 획득..
      if  Pos('RDW',sList) > 0 then
         UV_v012[UV_index] := Trim(copy(sList,Pos('RDW',sList) + 3,6));
      //PLT 획득..
      if  Pos('PLT',sList) > 0 then
         UV_v013[UV_index] := Trim(copy(sList,Pos('PLT',sList) + 3,6));
      //PCT 획득..
      if  Pos('PCT',sList) > 0 then
         UV_v014[UV_index] := Trim(copy(sList,Pos('PCT',sList) + 3,6));
      //MPV 획득..
      if  Pos('MPV',sList) > 0 then
         UV_v015[UV_index] := Trim(copy(sList,Pos('MPV',sList) + 3,6));
      //PDW 획득..
      if  Pos('PDW',sList) > 0 then
         UV_v016[UV_index] := Trim(copy(sList,Pos('PDW',sList) + 3,6));

      Inc(UV_index);

      grdMaster.Rows := 0;
      grdMaster.Rows := UV_index;

      UV_List := '';
   end;
end;

procedure TfrmLD5I02.UP_Clear;
begin
   UV_v001[UV_index] := ''; UV_v002[UV_index] := ''; UV_v003[UV_index] := ''; UV_v004[UV_index] := ''; UV_v005[UV_index] := '';
   UV_v006[UV_index] := ''; UV_v007[UV_index] := ''; UV_v008[UV_index] := ''; UV_v009[UV_index] := ''; UV_v010[UV_index] := '';
   UV_v011[UV_index] := ''; UV_v012[UV_index] := ''; UV_v013[UV_index] := ''; UV_v014[UV_index] := ''; UV_v015[UV_index] := '';
   UV_v016[UV_index] := '';
end;

procedure TfrmLD5I02.UP_SClear(iTemp : Integer);
begin
   UV_vS001[iTemp] := ''; UV_vS002[iTemp] := ''; UV_vS003[iTemp] := ''; UV_vS004[iTemp] := ''; UV_vS005[iTemp] := '';
   UV_vS006[iTemp] := ''; UV_vS007[iTemp] := ''; UV_vS008[iTemp] := ''; UV_vS009[iTemp] := ''; UV_vS010[iTemp] := '';
end;

procedure TfrmLD5I02.UP_SPClear(iTemp : Integer);
begin
   UV_vSP001[iTemp] := ''; UV_vSP002[iTemp] := ''; UV_vSP003[iTemp] := '';
end;

procedure TfrmLD5I02.Cb_PortChange(Sender: TObject);
begin
   inherited;
   case Cb_Port.ItemIndex of
     0 : ComPort.Port := 'COM1';
     1 : ComPort.Port := 'COM2';
     2 : ComPort.Port := 'COM3';
     3 : ComPort.Port := 'COM4';
     4 : ComPort.Port := 'COM5';
   end;
end;

procedure TfrmLD5I02.qrdSPListCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_vSP001[DataRow - 1];
      2 : Value := UV_vSP002[DataRow - 1];
      3 : Value := UV_vSP003[DataRow - 1];
   end;
end;

procedure TfrmLD5I02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;

   if Key <> GC_Refer then exit;

   if      Sender = mskDate then UP_Click(btnDate);
end;

procedure TfrmLD5I02.UP_Click(Sender: TObject);
begin
  inherited;
  if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD5I02.SBtn_StartClick(Sender: TObject);
var
   iRet, iIndex, i, iCnt, iStart, iEnd, sIndex, spIndex : Integer;
   vCod_chuga : Variant;
   sCod_hm, sHmCode, sGlkwa, sCode, sValue : String;
   bCheck : Boolean;
begin
   inherited;
   if not GF_MsgBox('Confirmation', '저장을 시작 하시겠습니까.?') then exit;
   
   Animate1.Active := True;
   Animate2.Active := True;

   if grdMaster.Rows = 0 then
   begin
      showMessage('업데이트 자료가 존재하지 않습니다.');
      exit;
   end;

   sIndex := 0; spIndex := 0;
   Gauge.MaxValue := grdMaster.Rows;
   Gauge.Progress := 0;

   qryHangmok.Active := False;
   qryHangmok.ParamByName('gubn_part').AsString := '02';
   qryHangmok.Active := True;
   
   for iIndex := 1 to grdMaster.Rows do
   begin
      bCheck := True;
      Gauge.Progress := Gauge.Progress + 1;
      Label_cmd.Caption := '작업내용 : ' + grdmaster.Cell[2,iIndex] + '[' +
                                           grdmaster.Cell[4,iIndex] + ']';
      Application.ProcessMessages;

      sCod_hm := '';

      with qryGumgin do
      begin
         Active := False;
         ParamByName('cod_jisa').AsString  := copy(GV_sUserId,1,2);
         ParamByName('dat_gmgn').AsString  := grdmaster.Cell[2,iIndex];
         ParamByName('num_samp').AsString  := grdmaster.Cell[4,iIndex];
         Active := True;

         if RecordCount > 0 then
         begin
            GP_query_log(qryGumgin, 'LD5I02', 'Q', 'Y');
            // 1. 혈액에 대한 검사항목 추출
            if Trim(FieldByName('Cod_hul').AsString) <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := Trim(FieldByName('Cod_hul').AsString);
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) <= 0 then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end; // end of while[qryPf_hm];
               end; // end of if[qryPf_hm];
            end; // end of if[Cod_hul];

            // 2. 종양에 대한 검사항목 추출
            if Trim(FieldByName('Cod_cancer').AsString) <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := Trim(FieldByName('Cod_cancer').AsString);
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) <= 0 then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end; // end of while[qryPf_hm];
               end; // end of if[qryPf_hm];
            end; // end of if[Cod_cancer];

            // 3. 장비에 대한 검사항목 추출
            if Trim(FieldByName('Cod_jangbi').AsString) <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('COD_PF').AsString := Trim(FieldByName('Cod_jangbi').AsString);
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     if pos(qryPf_hm.FieldByName('COD_HM').AsString, sCod_hm) <= 0 then
                        sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end; // end of while[qryPf_hm];
               end; // end of if[qryPf_hm];
            end; // end of if[Cod_jangbi];

            // 3. 추가코드에 대한 검사항목 추출
            iRet := GF_MulToSingle(Trim(FieldByName('Cod_chuga').AsString), 4, vCod_chuga);
            for i := 0 to iRet-1 do
            begin
               if copy(vCod_chuga[i],1,1) = 'H' then
               begin
                  if pos(vCod_chuga[i], sCod_hm) <= 0 then
                     sCod_hm := sCod_hm + vCod_chuga[i];
               end;
            end; // end of for[i];

            // 4. 항목코드에 대한 검사항목 추출
            sHmCode := '';
            if Trim(FieldByName('Gubn_nosin').AsString) = '1' then
               sHmCode := sHmCode + UF_No_Hangmok(FieldByName('Dat_gmgn').AsString, '1', FieldByName('Gubn_nsyh').AsInteger);

            if Trim(FieldByName('Gubn_adult').AsString) = '1' then
               sHmCode := sHmCode + UF_No_Hangmok(FieldByName('Dat_gmgn').AsString, '4', FieldByName('Gubn_adyh').AsInteger);

            if Trim(FieldByName('Gubn_bogun').AsString)  = '1' then
               sHmCode := sHmCode + UF_No_Hangmok(FieldByName('Dat_gmgn').AsString, '3', FieldByName('Gubn_bgyh').AsInteger);

            if Trim(FieldByName('Gubn_agab').AsString) = '1' then
               sHmCode := sHmCode + UF_No_Hangmok(FieldByName('Dat_gmgn').AsString, '5', FieldByName('Gubn_agyh').AsInteger);

            if (Trim(FieldByName('Gubn_tkgm').AsString) = '1') or
               (Trim(FieldByName('Gubn_tkgm').AsString) = '2') then
            begin
               qryTkgum.Active := False;
               qryTkgum.ParamByName('cod_jisa').AsString  := FieldByName('Cod_jisa').AsString;
               qryTkgum.ParamByName('num_jumin').AsString := FieldByName('num_jumin').AsString;
               qryTkgum.ParamByName('dat_gmgn').AsString  := FieldByName('Dat_gmgn').AsString;
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
                  if copy(vCod_chuga[i],1,1) = 'U' then
                  begin
                     if pos(vCod_chuga[i], sCod_hm) <= 0 then
                        sCod_hm := sCod_hm + vCod_chuga[i];
                  end;
               end; // end of for[i];
            end; // end of if[sHmCode];
         end
         else bCheck := False; // end of if[RecordCount];
      end; // end of with[qryGumgin];

      if bCheck then
      begin
         // 업로드 가능한 검진자...
         qryGulkwa.Active := False;
         qryGulkwa.ParamByName('cod_bjjs').AsString  := qryGumgin.FieldByName('cod_bjjs').AsString;
         qryGulkwa.ParamByName('dat_bunju').AsString := qryGumgin.FieldByName('dat_bunju').AsString;
         qryGulkwa.ParamByName('num_bunju').AsString := qryGumgin.FieldByName('num_bunju').AsString;
         qryGulkwa.ParamByName('gubn_part').AsString := '02';
         qryGulkwa.Active := True;
         if qryGulkwa.RecordCount > 0 then
         begin
            UV_fGulkwa := '';
            UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
            sGlkwa := UV_fGulkwa;
            qryGulkwa.Active := False;
            if sGlkwa = '' then
            begin
               for iCnt := 1 to 10 do
               begin
                  sGlkwa := sGlkwa + '                                                            ';
               end;
            end;

            for iCnt := 1 to 12 do
            begin
               sCode := UF_HagmokSetting(iIndex, iCnt, sValue, iStart, iEnd);
               if (pos(Trim(sCode), sCod_hm) > 0) and (Trim(sCode) <> '') then
               begin
                  sGlkwa := Copy(sGlkwa, 1, iStart-1) + GF_RPad(sValue, iEnd - iStart + 1, ' ')
                          + Copy(sGlkwa, iEnd+1, Length(sGlkwa));
               end;
            end;
         end
         else bCheck := False; // end of if[qryGulkwa];

         if bCheck then
         begin
            sGlkwa := 'A' + sGlkwa;
            sGlkwa := Trim(sGlkwa);
            sGlkwa := Copy(sGlkwa, 2, Length(sGlkwa)-1);

            // DB 작업
            try
               UV_fGulkwa1 := '';
               UV_fGulkwa2 := '';
               UV_fGulkwa3 := '';
               GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sGlkwa);

               qryU_Gulkwa.ParamByName('COD_BJJS').AsString    := qryGumgin.FieldByName('cod_bjjs').AsString;
               qryU_Gulkwa.ParamByName('DAT_BUNJU').AsString   := qryGumgin.FieldByName('dat_bunju').AsString;
               qryU_Gulkwa.ParamByName('NUM_BUNJU').AsString   := qryGumgin.FieldByName('num_bunju').AsString;
               qryU_Gulkwa.ParamByName('GUBN_PART').AsString   := '02';
               qryU_Gulkwa.ParamByName('DESC_GLKWA1').AsString := UV_fGulkwa1;
               qryU_Gulkwa.ParamByName('DESC_GLKWA2').AsString := UV_fGulkwa2;
               qryU_Gulkwa.ParamByName('DESC_GLKWA3').AsString := UV_fGulkwa3;
               qryU_Gulkwa.ParamByName('DAT_UPDATE').AsString  := GV_sToday;
               qryU_Gulkwa.ParamByName('COD_UPDATE').AsString  := GV_sUserId;

               qryU_Gulkwa.Execsql;
               GP_query_log(qryU_Gulkwa, 'LD5I02', 'Q', 'Y');
            except
               on E : EDBEngineError do
               begin
                  GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
                  exit;
               end;
            end;

            UP_SVarMemSet(sIndex);
            UP_SClear(sIndex);

            UV_vS001[sIndex] := grdmaster.Cell[1,iIndex];
            UV_vS002[sIndex] := qryGumgin.FieldByName('dat_gmgn').AsString;
            UV_vS003[sIndex] := qryGumgin.FieldByName('num_samp').AsString;
            UV_vS004[sIndex] := qryGumgin.FieldByName('desc_name').AsString;
            UV_vS005[sIndex] := qryGumgin.FieldByName('num_jumin').AsString;
            UV_vS006[sIndex] := qryGumgin.FieldByName('desc_bjjs').AsString;
            UV_vS007[sIndex] := qryGumgin.FieldByName('dat_bunju').AsString;
            UV_vS008[sIndex] := qryGumgin.FieldByName('num_bunju').AsString;
            UV_vS009[sIndex] := qryGumgin.FieldByName('desc_dc').AsString;
            UV_vS010[sIndex] := qryGumgin.FieldByName('desc_dept').AsString;

            Inc(sIndex);
         end
         else // 업로드 불가능한 검진자...
         begin
            UP_SPVarMemSet(spIndex);
            UP_SPClear(spIndex);

            UV_vSP001[spIndex] := grdmaster.Cell[1,iIndex];
            UV_vSP002[spIndex] := grdmaster.Cell[2,iIndex];
            UV_vSP003[spIndex] := grdmaster.Cell[3,iIndex];

            Inc(spIndex);
         end;
      end
      else // 업로드 불가능한 검진자...
      begin
         UP_SPVarMemSet(spIndex);
         UP_SPClear(spIndex);

         UV_vSP001[spIndex] := grdmaster.Cell[1,iIndex];
         UV_vSP002[spIndex] := grdmaster.Cell[2,iIndex];
         UV_vSP003[spIndex] := grdmaster.Cell[3,iIndex];

         Inc(spIndex);
      end;
   end; // end of for[grdMaster];

   qrdSList.Rows  := sIndex;
   qrdSPList.Rows := spIndex;

   Animate1.Active := False;
   Animate2.Active := False;
   showMessage('업데이트가 완료되었습니다!!');
end;

function  TfrmLD5I02.UF_HagmokSetting(iSIndex, sil : Integer; var sValue : String; var iStart, iEnd : Integer) : String;
var sCode, sTemp : String;
    i, iLength : Integer;
begin
   Result := '';
   case sil of
      // 적혈구수..
      1  : begin
              sValue := grdmaster.Cell[6,iSIndex];
              sCode  := 'H001';
           end;
      // 혈색소..
      2  : begin
              sValue := grdmaster.Cell[7,iSIndex];
              sCode := 'H002';
           end;
      // Hematocrit..
      3  : begin
              sValue := grdmaster.Cell[8,iSIndex];
              sCode := 'H003';
           end;
      // MCV..
      4  : begin
              sValue := grdmaster.Cell[9,iSIndex];
              sCode := 'H004';
           end;
      // MCH..
      5  : begin
              sValue := grdmaster.Cell[10,iSIndex];
              sCode := 'H005';
           end;
      // MCHC..
      6  : begin
              sValue := grdmaster.Cell[11,iSIndex];
              sCode := 'H006';
           end;
      // RDW..
      7  : begin
              sValue := grdmaster.Cell[12,iSIndex];
              sCode := 'H007';
           end;
      // 혈소판수..
      8  : begin
              sValue := grdmaster.Cell[13,iSIndex];
              sCode := 'H008';
           end;
      // MPV..
      9  : begin
              sValue := grdmaster.Cell[15,iSIndex];
              sCode := 'H009';
           end;
      // PDW..
      10 : begin
              sValue := grdmaster.Cell[16,iSIndex];
              sCode := 'H010';
           end;
      // 백혈구수..
      11 : begin
              sValue := grdmaster.Cell[5,iSIndex];
              sCode := 'H011';
           end;

      // PCT..
      12 : begin
              sValue := grdmaster.Cell[14,iSIndex];
              sCode := 'H035';
           end;

   else
      sCode := '';  sValue := '';
   end;

   with qryHangmok do
   begin
      Filtered := True;
      Filter := ' cod_hm = ''' + sCode + '''';
      if RecordCount = 0 then   exit;
      iStart := FieldByName('pos_start').Asinteger;
      iEnd   := FieldByName('pos_end').Asinteger;
   end;

   sTemp := sValue;
   iLength := Length(sValue);
   sValue := '';
   for i := 1 to iLength do
   begin
      if ((Copy(sTemp, i, 1) >= '0') and (Copy(sTemp, i, 1) <= '9')) or (Copy(sTemp, i, 1) = '.') then
        sValue := sValue + Copy(sTemp, i, 1);
   end;

   Result := sCode;
end;

function  TfrmLD5I02.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

end.






