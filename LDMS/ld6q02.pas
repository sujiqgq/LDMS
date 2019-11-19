//==============================================================================
// 프로그램 설명 : 혈액소견리스트
// 시스템        : 통합검진
// 수정일자      : 2012.12.3
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD6Q02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;




type
  TfrmLD6Q02 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    QRCompositeReport: TQRCompositeReport;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    Panel2: TPanel;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    edt_sokun: TEdit;
    Label1: TLabel;
    Chk_Last: TCheckBox;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure UP_Click(Sender: TObject);

     procedure MouseWheelHandler(var Message: TMessage); override;

    procedure FormActivate(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vDat_gmgn, UV_vBunju,UV_vCod_sokun   : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);


    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD6Q02: TfrmLD6Q02;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD6Q021, LD3I13;

{$R *.DFM}

procedure TfrmLD6Q02.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;



end;

procedure TfrmLD6Q02.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD6Q02.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vBunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_sokun   := VarArrayCreate([0, 0], varOleStr);


  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vBunju,  iLength);
      VarArrayRedim(UV_vCod_sokun,  iLength);

   end;
end;

function TfrmLD6Q02.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD6Q02.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD6Q02.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vBunju[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vCod_sokun[DataRow-1];
   end;
end;

procedure TfrmLD6Q02.btnQueryClick(Sender: TObject);
var index, i, iRet, iTemp, num : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga, vCod_profile : Variant;
    check_tk : boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;
  

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;

      if Chk_last.Checked =false  then
      begin
           sSelect := ' select A.cod_jisa, A.num_id, A.dat_gmgn, A.num_samp, A.desc_name,A.num_jumin,H.cod_sokun,H.num_bunju';
           sSelect := sSelect + ' From hul_sokun H   left outer join Gumgin A ';
           sSelect := sSelect + ' On H.cod_jisa = A.cod_jisa and a.num_id = H.num_id and a.dat_gmgn = h.dat_gmgn ';
           sWhere := ' WHERE H.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
           sWhere := sWhere + ' AND  H.Dat_Bunju = ''' + mskDate.Text + ''' ';
           sWhere := sWhere + ' AND  H.cod_sokun like  ''' + edt_sokun.Text +'%'+ '''';


           if Trim(mskFrom.Text) <> '' then
              sWhere := sWhere + ' AND H.num_bunju >= ''' + mskFrom.Text + '''';
           if Trim(mskTo.Text) <> '' then
              sWhere := sWhere + ' AND H.num_bunju <= ''' + mskTo.Text + '''';
           sOrderBy := ' ORDER BY CAST(H.num_bunju AS INT)';

      end
      else if Chk_last.Checked =true then
      begin
           sSelect := ' select A.cod_jisa, A.num_id, A.dat_gmgn, A.num_samp, A.desc_name,A.num_jumin,H.num_bunju From hul_sokun H  left outer join Gumgin A';
           sSelect := sSelect + ' On H.cod_jisa = A.cod_jisa and a.num_id = H.num_id and a.dat_gmgn = h.dat_gmgn ';
           sWhere := ' WHERE H.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND  H.Dat_Bunju = ''' + mskDate.Text + '''  AND H.num_bunju >= ''' + mskFrom.Text + ''' AND H.num_bunju <= ''' + mskTo.Text + '''    AND ';
           sWhere := sWhere + ' A.Num_id not in (select distinct num_id from hul_sokun where  cod_sokun in (''NH1'', ''NH2'', ''NH3'', ''NH4'', ''NJG'', ''NJB'')  AND ';
           sWhere := sWhere + ' Dat_Bunju =  ''' + mskDate.Text + ''' and  Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''  AND num_bunju >= ''' + mskFrom.Text + ''' AND num_bunju <= ''' + mskTo.Text + '''   )';
           sWhere := sWhere + ' group by A.cod_jisa, A.num_id, A.dat_gmgn, A.num_samp, A.desc_name,A.num_jumin,H.num_bunju ';
           sOrderBy := ' ORDER BY CAST(A.num_samp AS INT)'
      end;


      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q37', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := IntToStr(Index+1);
               UV_vJisa[Index] := qryBunju.FieldByName('Cod_jisa').AsString;
               UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[Index] :=GF_JuminFormat(qryBunju.FieldByName('num_jumin').AsString);
               UV_vDat_gmgn[Index] := GF_DateFormat(qryBunju.FieldByName('dat_gmgn').AsString);
               UV_vBunju[Index] := qryBunju.FieldByName('num_Bunju').AsString;
               if Chk_Last.Checked = false  then UV_vCod_sokun[Index] := qryBunju.FieldByName('cod_sokun').AsString;
               Inc(Index);
            Next;
         End;

         // Table과 Disconnect
         Close;


      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;
   
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD6Q02.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD6Q02.FormActivate(Sender: TObject);
begin
  inherited;
   mskFrom.Text := '0001';
   mskTo.Text := '9999';

   mskDate.SetFocus;
end;

procedure TfrmLD6Q02.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD6Q021 := TfrmLD6Q021.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD6Q021.QuickRep.Preview
  else                                frmLD6Q021.QuickRep.Print;
end;

procedure TfrmLD6Q02.SBut_ExcelClick(Sender: TObject);
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

procedure TfrmLD6Q02.cmbRelationChange(Sender: TObject);
begin
  inherited;
  if (cmbRelation.ItemIndex = 0) AND ((GV_sUserId ='150005') OR (GV_sUserId ='151026')) then
   begin
      Application.CreateForm(TfrmLD3I13, frmLD3I13);
      frmLD3I13.Show;
   end;
end;

procedure TfrmLD6Q02.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If Key = VK_F12 Then SBut_ExcelClick(Self);
end;

end.
