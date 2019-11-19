//==============================================================================
// 프로그램 설명 : 혈액결과통계
// 시스템        : 분석관리
// 작성일자      : 2008.10.02
// 작성자        : 구수정
//==============================================================================


unit LD8Q06;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Calendar;

type
  TfrmLD8Q06 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    mskTo: TDateEdit;
    qry_gulkwa: TQuery;
    btnsDate: TSpeedButton;
    Label2: TLabel;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    qryCell: TQuery;
    qry_gulkwa2: TQuery;
    Label3: TLabel;
    mskFrom: TDateEdit;
    btneDate: TSpeedButton;
    Panel2: TPanel;
    grdMaster: TtsGrid;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel29: TPanel;
    Panel30: TPanel;
    Panel31: TPanel;
    Panel32: TPanel;
    Panel33: TPanel;
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
    Panel25: TPanel;
    Panel26: TPanel;
    Panel34: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Panel35: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure cmbJisaChange(Sender: TObject);
    procedure UP_clear;
    procedure btnPrintClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vM20, UV_vF20, UV_vM30, UV_vF30,
    UV_vM40, UV_vF40, UV_vM50, UV_vF50,
    UV_vM60, UV_vF60, UV_vM70, UV_vF70,
    UV_vMSum, UV_vFSum  : Variant;

    UV_sBasicSQL: String;
    UV_bCkEdit, UV_bCkEditCnt : Boolean;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    // 달력관련한 함수 및 변수..
    UV_sCalDate : String;
    function  UF_SawonCk : Boolean;
    { Public declarations }
  end;
const
    NumberOfRows = 10;
    NumberOfCols = 14;
var
  frmLD8Q06: TfrmLD8Q06;
  TextArray: Array [1..NumberOfCols,1..NumberOfRows] of Variant;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD8Q061;

{$R *.DFM}

procedure TfrmLD8Q06.MouseWheelHandler(var Message: TMessage);  //수정
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //수정
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //수정
      end;
   end;
end;

procedure TfrmLD8Q06.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vM20   := VarArrayCreate([0, 0], varOleStr);
      UV_vF20   := VarArrayCreate([0, 0], varOleStr);
      UV_vM30   := VarArrayCreate([0, 0], varOleStr);
      UV_vF30   := VarArrayCreate([0, 0], varOleStr);
      UV_vM40   := VarArrayCreate([0, 0], varOleStr);
      UV_vF40   := VarArrayCreate([0, 0], varOleStr);
      UV_vM50   := VarArrayCreate([0, 0], varOleStr);
      UV_vF50   := VarArrayCreate([0, 0], varOleStr);
      UV_vM60   := VarArrayCreate([0, 0], varOleStr);
      UV_vF60   := VarArrayCreate([0, 0], varOleStr);
      UV_vM70   := VarArrayCreate([0, 0], varOleStr);
      UV_vF70   := VarArrayCreate([0, 0], varOleStr);
      UV_vMSum  := VarArrayCreate([0, 0], varOleStr);
      UV_vFSum  := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      VarArrayRedim(UV_vM20,    iLength);
      VarArrayRedim(UV_vF20,    iLength);
      VarArrayRedim(UV_vM30,    iLength);
      VarArrayRedim(UV_vF30,    iLength);
      VarArrayRedim(UV_vM40,    iLength);
      VarArrayRedim(UV_vF40,    iLength);
      VarArrayRedim(UV_vM50,    iLength);
      VarArrayRedim(UV_vF50,    iLength);
      VarArrayRedim(UV_vM60,    iLength);
      VarArrayRedim(UV_vF60,    iLength);
      VarArrayRedim(UV_vM70,    iLength);
      VarArrayRedim(UV_vF70,    iLength);
      VarArrayRedim(UV_vMSum,   iLength);
      VarArrayRedim(UV_vFSum,   iLength);
   end;
end;

procedure TfrmLD8Q06.UP_clear;
var I : Integer;
begin
   for I := 0 to 10 do
   begin
      UV_vM20[I]  := '0';
      UV_vF20[I]  := '0';
      UV_vM30[I]  := '0';
      UV_vF30[I]  := '0';
      UV_vM40[I]  := '0';
      UV_vF40[I]  := '0';
      UV_vM50[I]  := '0';
      UV_vF50[I]  := '0';
      UV_vM60[I]  := '0';
      UV_vF60[I]  := '0';
      UV_vM70[I]  := '0';
      UV_vF70[I]  := '0';
      UV_vMSum[I] := '0';
      UV_vFSum[I] := '0';
   end;
end;


function TfrmLD8Q06.UF_Condition : Boolean;
var bError : Boolean;
begin
   Result := True;
   bError := False;

   // 지사가 입력되었는지 Check
   if (GV_sJisa = '00') and (cmbJisa.ItemIndex = -1) then
   begin
      GF_MsgBox('Warning', '지사를 반드시 선택하십시요.');
      cmbJisa.SetFocus;
      Result := False;
      exit;
   end;


   // 오류일 경우 Message 표시
   if bError then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;


procedure TfrmLD8Q06.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);
   grdMaster.Rows := 0;

   // 환경설정
   UP_VarMemSet(0);
   UV_bCkEdit := False;

   // SQL문 저장
   UV_sBasicSQL := qry_gulkwa.SQL.Text;

end;

procedure TfrmLD8Q06.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   // 자룔를 화면에 조회
   case DataCol of
      1  :  Value := UV_vM20[DataRow-1];
      2  :  Value := UV_vF20[DataRow-1];
      3  :  Value := UV_vM30[DataRow-1];
      4  :  Value := UV_vF30[DataRow-1];
      5  :  Value := UV_vM40[DataRow-1];
      6  :  Value := UV_vF40[DataRow-1];
      7  :  Value := UV_vM50[DataRow-1];
      8  :  Value := UV_vF50[DataRow-1];
      9  :  Value := UV_vM60[DataRow-1];
      10 :  Value := UV_vF60[DataRow-1];
      11 :  Value := UV_vM70[DataRow-1];
      12 :  Value := UV_vF70[DataRow-1];
      13 :  Value := UV_vMSum[DataRow-1];
      14 :  Value := UV_vFSum[DataRow-1];
   end;
end;

procedure TfrmLD8Q06.btnQueryClick(Sender: TObject);
var i, index, iRet,  iTemp, iAge,
    vCnt_C009, vCnt_C010, vCnt_C011, vCnt_C025, vCnt_H002_M, vCnt_H002_F,
    vCnt_U004, vCnt_U009, vCnt_S007  : Integer;
    sWhere, sGroupBy, sOrderBy, sHul_List, vSex : String;
    vCod_chuga : Variant;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 10;
   UP_VarMemSet(10);
   UP_clear;

   // Query 수행 & 배열에 저장

   index := 0;
   with qry_gulkwa do
   begin
        Active := False;

        // SQL문 생성
        sWhere := '';
        sGroupBy := '';
        sHul_List := '';

        {if Copy(cmbJisa.Text, 1, 2) <> '15' then
           sWhere := ' WHERE (G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' or G.cod_jisa = ''15'') '
        else
            sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''''; }

        sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + '''';

        sWhere := sWhere +  ' AND G.dat_gmgn >= ''' + mskTo.Text + '''';
        sWhere := sWhere +  ' AND G.dat_gmgn <= ''' + mskFrom.Text + '''';

        sWhere := sWhere +  ' AND K.num_bunju >= ''0000'' ';
        sWhere := sWhere +  ' AND K.num_bunju <= ''9999'' ';
        //sWhere := sWhere +  ' AND G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' ';
        {if Copy(cmbJisa.Text, 1, 2) <> '15' then
           sWhere := sWhere +  ' AND (G.gubn_hulgum = ''2'' or G.gubn_hulgum = ''3'') '
        else
            sWhere := sWhere +  ' AND G.gubn_hulgum = ''1''';    }

        sGroupBy := ' GROUP BY G.dat_gmgn, K.cod_bjjs, G.num_id, K.dat_bunju, K.num_bunju, G.num_jumin';

        sOrderBy := ' ORDER BY K.dat_bunju, K.num_bunju' ;

        SQL.Clear;
        SQL.Add(UV_sBasicSQL + sWhere + sGroupBy + sOrderBy);

        Active := True;

        if qry_gulkwa.RecordCount > 0 then
        begin
             GP_query_log(qry_gulkwa, 'LD8Q06', 'Q', 'N');
             for i := 0 to RecordCount - 1 do
             begin
                  case StrToInt(Copy(qry_gulkwa.FieldByName('num_jumin').AsString, 7, 1)) of
                     1,3,5,7,9 : vSex := 'M';
                     2,4,6,8   : vSex := 'F';
                  end;
                  GP_JuminToAge(qry_gulkwa.FieldByName('num_jumin').AsString,qry_gulkwa.FieldByName('Dat_gmgn').AsString, vSex, iAge);
                  //qrlAge.Caption := '(' + IntToStr(iAge) + ')';

                  with qryPGulkwa do
                  begin
                       Active := False;
                       ParamByName('cod_bjjs').AsString  := qry_Gulkwa.FieldByName('cod_bjjs').AsString;
                       ParamByName('num_id').AsString    := qry_Gulkwa.FieldByName('num_id').AsString;
                       ParamByName('dat_gmgn').AsString  := qry_Gulkwa.FieldByName('dat_gmgn').AsString;
                       Active := True;

                       if qryPGulkwa.RecordCount > 0 then
                       begin
                          vCnt_C009 := 0; vCnt_C010 := 0; vCnt_C011 := 0; vCnt_C025 := 0; vCnt_H002_M := 0; vCnt_H002_F := 0;
                          vCnt_U004 := 0; vCnt_U009 := 0; vCnt_S007 := 0;

                           filter := 'gubn_part = ''01''';
                           filtered := true;

                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,49,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,49,6))) >= 100) then
                           begin
                                vCnt_C009 := 1;
                           end;
                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6))) >= 100) then
                           begin
                                vCnt_C010 := 1;
                           end;
                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,61,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,61,6))) >= 100) then
                           begin
                                vCnt_C011 := 1;
                           end;
                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,121,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,121,6))) >= 240) then
                           begin
                                vCnt_C025 := 1;
                           end;

                           filter := 'gubn_part = ''02''';
                           filtered := true;

                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '02') and (vSex = 'M')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)) <> '')
                              and  (Trunc(StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)))) < 12) then
                           begin
                                vCnt_H002_M := 1;
                           end
                           else if (qryPGulkwa.FieldByName('gubn_part').AsString = '02') and (vSex = 'F')
                                and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)) <> '')
                                and  (Trunc(StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)))) < 10) then
                           begin
                                vCnt_H002_F := 1;
                           end;

                           filter := 'gubn_part = ''03''';
                           filtered := true;

                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '03')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,25,6))) >= 40) then
                           begin
                                vCnt_U004 := 1;
                           end;
                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '03')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6)) <> '')
                              and (StrToInt(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6))) >= 30) then
                           begin
                                vCnt_U009 := 1;
                           end;

                           filter := 'gubn_part = ''07''';
                           filtered := true;

                           if (qryPGulkwa.FieldByName('gubn_part').AsString = '07')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) <> '')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,31,6)) = '양성') then
                           begin
                                vCnt_S007 := 1;
                           end;
                           {if (qryPGulkwa.FieldByName('gubn_part').AsString = '07')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) <> '')
                              and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,37,6)) = '양성') then
                           begin
                                vCnt_S008 := 1;
                           end;   }

                       end;
                       filtered := False;
                  end;  //end of with
                  if  vSex = 'M' then
                  begin
                     case iAge of
                        20..30 : begin
                                    UV_vM20[0] := UV_vM20[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM20[1]     := UV_vM20[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM20[2]     := UV_vM20[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM20[3]     := UV_vM20[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM20[4]     := UV_vM20[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM20[5]     := UV_vM20[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM20[6]     := UV_vM20[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM20[7]     := UV_vM20[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM20[8]     := UV_vM20[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM20[9]    := UV_vM20[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM20[11]    := UV_vM20[11]    + vCnt_S008;   }
                                 end;
                        31..40 : begin
                                    UV_vM30[0] := UV_vM30[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM30[1]     := UV_vM30[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM30[2]     := UV_vM30[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM30[3]     := UV_vM30[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM30[4]     := UV_vM30[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM30[5]     := UV_vM30[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM30[6]     := UV_vM30[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM30[7]     := UV_vM30[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM30[8]     := UV_vM30[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM30[9]    := UV_vM30[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM30[11]    := UV_vM30[11]    + vCnt_S008; }
                                 end;
                        41..50 : begin
                                    UV_vM40[0] := UV_vM40[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM40[1]     := UV_vM40[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM40[2]     := UV_vM40[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM40[3]     := UV_vM40[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM40[4]     := UV_vM40[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM40[5]     := UV_vM40[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM40[6]     := UV_vM40[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM40[7]     := UV_vM40[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM40[8]     := UV_vM40[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM40[9]    := UV_vM40[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM40[11]    := UV_vM40[11]    + vCnt_S008;}
                                 end;
                        51..60 : begin
                                    UV_vM50[0] := UV_vM50[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM50[1]     := UV_vM50[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM50[2]     := UV_vM50[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM50[3]     := UV_vM50[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM50[4]     := UV_vM50[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM50[5]     := UV_vM50[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM50[6]     := UV_vM50[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM50[7]     := UV_vM50[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM50[8]     := UV_vM50[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM50[9]    := UV_vM50[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM50[11]    := UV_vM50[11]    + vCnt_S008; }
                                 end;
                        61..70 : begin
                                    UV_vM60[0] := UV_vM60[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM60[1]     := UV_vM60[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM60[2]     := UV_vM60[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM60[3]     := UV_vM60[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM60[4]     := UV_vM60[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM60[5]     := UV_vM60[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM60[6]     := UV_vM60[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM60[7]     := UV_vM60[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM60[8]     := UV_vM60[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM60[9]    := UV_vM60[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM60[11]    := UV_vM60[11]    + vCnt_S008;  }
                                 end;
                        71..100 : begin
                                    UV_vM70[0] := UV_vM70[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vM70[1]     := UV_vM70[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vM70[2]     := UV_vM70[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vM70[3]     := UV_vM70[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vM70[4]     := UV_vM70[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vM70[5]     := UV_vM70[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vM70[6]     := UV_vM70[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vM70[7]     := UV_vM70[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vM70[8]     := UV_vM70[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vM70[9]    := UV_vM70[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vM70[11]    := UV_vM70[11]    + vCnt_S008; }
                                 end;
                        end;
                        UV_vMSum[0] := FloatToStr(GF_StrToNum(UV_vM20[0]) + GF_StrToNum(UV_vM30[0]) +
                                                       GF_StrToNum(UV_vM40[0]) + GF_StrToNum(UV_vM50[0]) +
                                                       GF_StrToNum(UV_vM60[0]) + GF_StrToNum(UV_vM70[0]));
                        UV_vMSum[1] := FloatToStr(GF_StrToNum(UV_vM20[1]) + GF_StrToNum(UV_vM30[1]) +
                                                       GF_StrToNum(UV_vM40[1]) + GF_StrToNum(UV_vM50[1]) +
                                                       GF_StrToNum(UV_vM60[1]) + GF_StrToNum(UV_vM70[1]));
                        UV_vMSum[2] := FloatToStr(GF_StrToNum(UV_vM20[2]) + GF_StrToNum(UV_vM30[2]) +
                                                       GF_StrToNum(UV_vM40[2]) + GF_StrToNum(UV_vM50[2]) +
                                                       GF_StrToNum(UV_vM60[2]) + GF_StrToNum(UV_vM70[2]));
                        UV_vMSum[3] := FloatToStr(GF_StrToNum(UV_vM20[3]) + GF_StrToNum(UV_vM30[3]) +
                                                       GF_StrToNum(UV_vM40[3]) + GF_StrToNum(UV_vM50[3]) +
                                                       GF_StrToNum(UV_vM60[3]) + GF_StrToNum(UV_vM70[3]));
                        UV_vMSum[4] := FloatToStr(GF_StrToNum(UV_vM20[4]) + GF_StrToNum(UV_vM30[4]) +
                                                       GF_StrToNum(UV_vM40[4]) + GF_StrToNum(UV_vM50[4]) +
                                                       GF_StrToNum(UV_vM60[4]) + GF_StrToNum(UV_vM70[4]));
                        UV_vMSum[5] := FloatToStr(GF_StrToNum(UV_vM20[5]) + GF_StrToNum(UV_vM30[5]) +
                                                       GF_StrToNum(UV_vM40[5]) + GF_StrToNum(UV_vM50[5]) +
                                                       GF_StrToNum(UV_vM60[5]) + GF_StrToNum(UV_vM70[5]));
                        UV_vMSum[6] := FloatToStr(GF_StrToNum(UV_vM20[6]) + GF_StrToNum(UV_vM30[6]) +
                                                       GF_StrToNum(UV_vM40[6]) + GF_StrToNum(UV_vM50[6]) +
                                                       GF_StrToNum(UV_vM60[6]) + GF_StrToNum(UV_vM70[6]));
                        UV_vMSum[7] := FloatToStr(GF_StrToNum(UV_vM20[7]) + GF_StrToNum(UV_vM30[7]) +
                                                       GF_StrToNum(UV_vM40[7]) + GF_StrToNum(UV_vM50[7]) +
                                                       GF_StrToNum(UV_vM60[7]) + GF_StrToNum(UV_vM70[7]));
                        UV_vMSum[8] := FloatToStr(GF_StrToNum(UV_vM20[8]) + GF_StrToNum(UV_vM30[8]) +
                                                       GF_StrToNum(UV_vM40[8]) + GF_StrToNum(UV_vM50[8]) +
                                                       GF_StrToNum(UV_vM60[8]) + GF_StrToNum(UV_vM70[8]));
                        UV_vMSum[9] := FloatToStr(GF_StrToNum(UV_vM20[9]) + GF_StrToNum(UV_vM30[9]) +
                                                       GF_StrToNum(UV_vM40[9]) + GF_StrToNum(UV_vM50[9]) +
                                                       GF_StrToNum(UV_vM60[9]) + GF_StrToNum(UV_vM70[9]));
                        {UV_vMSum[11] := FloatToStr(GF_StrToNum(UV_vM20[11]) + GF_StrToNum(UV_vM30[11]) +
                                                       GF_StrToNum(UV_vM40[11]) + GF_StrToNum(UV_vM50[11]) +
                                                       GF_StrToNum(UV_vM60[11]) + GF_StrToNum(UV_vM70[11]));  }
                  end
                  else if  vSex = 'F' then
                  begin
                     case iAge of
                        20..30 : begin
                                    UV_vF20[0] := UV_vF20[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF20[1]     := UV_vF20[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF20[2]     := UV_vF20[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF20[3]     := UV_vF20[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF20[4]     := UV_vF20[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF20[5]     := UV_vF20[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF20[6]     := UV_vF20[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF20[7]     := UV_vF20[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF20[8]     := UV_vF20[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF20[9]    := UV_vF20[9]      + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF20[11]    := UV_vF20[11]    + vCnt_S008;   }
                                 end;
                        31..40 : begin
                                    UV_vF30[0] := UV_vF30[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF30[1]     := UV_vF30[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF30[2]     := UV_vF30[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF30[3]     := UV_vF30[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF30[4]     := UV_vF30[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF30[5]     := UV_vF30[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF30[6]     := UV_vF30[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF30[7]     := UV_vF30[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF30[8]     := UV_vF30[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF30[9]    := UV_vF30[9]    + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF30[11]    := UV_vF30[11]    + vCnt_S008;}
                                 end;
                        41..50 : begin
                                    UV_vF40[0] := UV_vF40[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF40[1]     := UV_vF40[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF40[2]     := UV_vF40[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF40[3]     := UV_vF40[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF40[4]     := UV_vF40[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF40[5]     := UV_vF40[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF40[6]     := UV_vF40[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF40[7]     := UV_vF40[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF40[8]     := UV_vF40[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF40[9]    := UV_vF40[9]    + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF40[11]    := UV_vF40[11]    + vCnt_S008; }
                                 end;
                        51..60 : begin
                                    UV_vF50[0] := UV_vF50[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF50[1]     := UV_vF50[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF50[2]     := UV_vF50[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF50[3]     := UV_vF50[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF50[4]     := UV_vF50[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF50[5]     := UV_vF50[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF50[6]     := UV_vF50[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF50[7]     := UV_vF50[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF50[8]     := UV_vF50[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF50[9]    := UV_vF50[9]    + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF50[11]    := UV_vF50[11]    + vCnt_S008; }
                                 end;
                        61..70 : begin
                                    UV_vF60[0] := UV_vF60[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF60[1]     := UV_vF60[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF60[2]     := UV_vF60[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF60[3]     := UV_vF60[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF60[4]     := UV_vF60[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF60[5]     := UV_vF60[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF60[6]     := UV_vF60[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF60[7]     := UV_vF60[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF60[8]     := UV_vF60[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF60[9]    := UV_vF60[9]    + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF60[11]    := UV_vF60[11]    + vCnt_S008; }
                                 end;
                        71..100 : begin
                                    UV_vF70[0] := UV_vF70[0] + 1;

                                    if vCnt_C009 = 1 then
                                      UV_vF70[1]     := UV_vF70[1]     + vCnt_C009
                                    else if vCnt_C010 = 1 then
                                      UV_vF70[2]     := UV_vF70[2]     + vCnt_C010
                                    else if vCnt_C011 = 1 then
                                      UV_vF70[3]     := UV_vF70[3]     + vCnt_C011
                                    else if vCnt_C025 = 1 then
                                      UV_vF70[4]     := UV_vF70[4]     + vCnt_C025
                                    else if vCnt_H002_M = 1 then
                                      UV_vF70[5]     := UV_vF70[5]     + vCnt_H002_M
                                    else if vCnt_H002_F = 1 then
                                      UV_vF70[6]     := UV_vF70[6]     + vCnt_H002_F
                                    else if vCnt_U004 = 1 then
                                      UV_vF70[7]     := UV_vF70[7]     + vCnt_U004
                                    else if vCnt_U009 = 1 then
                                      UV_vF70[8]     := UV_vF70[8]     + vCnt_U009
                                    else if vCnt_S007 = 1 then
                                      UV_vF70[9]    := UV_vF70[9]    + vCnt_S007;
                                    {else if vCnt_S008 = 1 then
                                      UV_vF70[11]    := UV_vF70[11]    + vCnt_S008; }
                                 end;
                        end;
                        UV_vFSum[0] := FloatToStr(GF_StrToNum(UV_vF20[0]) + GF_StrToNum(UV_vF30[0]) +
                                                       GF_StrToNum(UV_vF40[0]) + GF_StrToNum(UV_vF50[0]) +
                                                       GF_StrToNum(UV_vF60[0]) + GF_StrToNum(UV_vF70[0]));
                        UV_vFSum[1] := FloatToStr(GF_StrToNum(UV_vF20[1]) + GF_StrToNum(UV_vF30[1]) +
                                                       GF_StrToNum(UV_vF40[1]) + GF_StrToNum(UV_vF50[1]) +
                                                       GF_StrToNum(UV_vF60[1]) + GF_StrToNum(UV_vF70[1]));
                        UV_vFSum[2] := FloatToStr(GF_StrToNum(UV_vF20[2]) + GF_StrToNum(UV_vF30[2]) +
                                                       GF_StrToNum(UV_vF40[2]) + GF_StrToNum(UV_vF50[2]) +
                                                       GF_StrToNum(UV_vF60[2]) + GF_StrToNum(UV_vF70[2]));
                        UV_vFSum[3] := FloatToStr(GF_StrToNum(UV_vF20[3]) + GF_StrToNum(UV_vF30[3]) +
                                                       GF_StrToNum(UV_vF40[3]) + GF_StrToNum(UV_vF50[3]) +
                                                       GF_StrToNum(UV_vF60[3]) + GF_StrToNum(UV_vF70[3]));
                        UV_vFSum[4] := FloatToStr(GF_StrToNum(UV_vF20[4]) + GF_StrToNum(UV_vF30[4]) +
                                                       GF_StrToNum(UV_vF40[4]) + GF_StrToNum(UV_vF50[4]) +
                                                       GF_StrToNum(UV_vF60[4]) + GF_StrToNum(UV_vF70[4]));
                        UV_vFSum[5] := FloatToStr(GF_StrToNum(UV_vF20[5]) + GF_StrToNum(UV_vF30[5]) +
                                                       GF_StrToNum(UV_vF40[5]) + GF_StrToNum(UV_vF50[5]) +
                                                       GF_StrToNum(UV_vF60[5]) + GF_StrToNum(UV_vF70[5]));
                        UV_vFSum[6] := FloatToStr(GF_StrToNum(UV_vF20[6]) + GF_StrToNum(UV_vF30[6]) +
                                                       GF_StrToNum(UV_vF40[6]) + GF_StrToNum(UV_vF50[6]) +
                                                       GF_StrToNum(UV_vF60[6]) + GF_StrToNum(UV_vF70[6]));
                        UV_vFSum[7] := FloatToStr(GF_StrToNum(UV_vF20[7]) + GF_StrToNum(UV_vF30[7]) +
                                                       GF_StrToNum(UV_vF40[7]) + GF_StrToNum(UV_vF50[7]) +
                                                       GF_StrToNum(UV_vF60[7]) + GF_StrToNum(UV_vF70[7]));
                        UV_vFSum[8] := FloatToStr(GF_StrToNum(UV_vF20[8]) + GF_StrToNum(UV_vF30[8]) +
                                                       GF_StrToNum(UV_vF40[8]) + GF_StrToNum(UV_vF50[8]) +
                                                       GF_StrToNum(UV_vF60[8]) + GF_StrToNum(UV_vF70[8]));
                        UV_vFSum[9] := FloatToStr(GF_StrToNum(UV_vF20[9]) + GF_StrToNum(UV_vF30[9]) +
                                                       GF_StrToNum(UV_vF40[9]) + GF_StrToNum(UV_vF50[9]) +
                                                       GF_StrToNum(UV_vF60[9]) + GF_StrToNum(UV_vF70[9]));
                        {UV_vFSum[11] := FloatToStr(GF_StrToNum(UV_vF20[11]) + GF_StrToNum(UV_vF30[11]) +
                                                       GF_StrToNum(UV_vF40[11]) + GF_StrToNum(UV_vF50[11]) +
                                                       GF_StrToNum(UV_vF60[11]) + GF_StrToNum(UV_vF70[11])); }
                  end;
                  next;

             end;  //end of for

             // Table과 Disconnect
             //Active := False;
        end
        else
        begin
             GF_MsgBox('Information', 'NODATA');
             Exit;
        end;
   end;
   // Grid에 자료를 할당
   grdMaster.Rows := 10;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD8Q06.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;
   UV_bCkEdit := False; UV_bCkEditCnt := False;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   //같은 지사는 입력 및 수정 가능하게..
   cmbJisaChange(self);

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
   
   // 수정이 아닐때...
   if not UV_bEdit then
   begin
      UV_bCkEdit := True; UV_bCkEditCnt := True;
   end;
end;


procedure TfrmLD8Q06.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

end;

procedure TfrmLD8Q06.UP_Click(Sender: TObject);
begin
   inherited;
   //날짜 입력
    if Sender = btnsDate then
      GF_CalendarClick(mskTo);
    if Sender = btneDate then
      GF_CalendarClick(mskFrom);

end;


procedure TfrmLD8Q06.rbDateClick(Sender: TObject);
begin
   inherited;

   // 모두 No Visible
   mskFrom.Visible  := False; 
   btnsDate.Visible := False;


   
end;

function  TfrmLD8Q06.UF_SawonCk : Boolean;
var sSelect, sWhere : String;
begin
   Result := False;
   with dmComm.qryShare do
   begin
      Active := False;
      SQL.Clear;

      // SQL문 생성
      sSelect := 'SELECT cod_sawon, gubn_dept FROM sawon';
      sWhere  := ' WHERE cod_sawon = ''' + GV_sUserId + '''';

      SQL.Add(sSelect + sWhere);
      Active := True;

      if RecordCount > 0 then
      begin
         if copy(FieldByName('cod_sawon').AsString,1,2) = '15' then
         begin
            {if Trim(FieldByName('gubn_dept').AsString) = '' then
               Result := False
            else
               case FieldByName('gubn_dept').AsInteger of
                  7, 15, 25 : Result := True;
                  else
                     Result := False;
               end;   }
         end
         else if copy(FieldByName('cod_sawon').AsString,1,2) = '12' then
         begin
           { if Trim(FieldByName('gubn_dept').AsString) = '' then
               Result := False
            else
               case FieldByName('gubn_dept').AsInteger of
                  3, 4, 5, 6, 7, 11, 12, 14 : Result := True;
                  else
                     Result := False;
               end;  }
         end;
      end;
      // Table Disconnect
      Active := False;
   end;
end;
procedure TfrmLD8Q06.cmbJisaChange(Sender: TObject);
var jisa : string;
begin
   inherited;


end;
procedure TfrmLD8Q06.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD8Q061 := TfrmLD8Q061.Create(self);
  frmLD8Q061.QuickRep.preview;
  frmLD8Q061.Free;
end;

end.
