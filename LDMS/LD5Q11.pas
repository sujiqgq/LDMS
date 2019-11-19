//==============================================================================
// 프로그램 설명 : 통합문진 조회
// 시스템        : 통합검진
// 수정일자      : 2015.03.27
// 수정자        : 곽윤설
// 수정내용      : 새로개발
// 참고사항      : [한의 본원진료팀1500177 - 김소영 전문의]
//==============================================================================

unit LD5Q11;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,ComObj,
  QuickRpt;

type
  TfrmLD5Q11 = class(TfrmSingle)
    grdMaster: TtsGrid;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    QRCompositeReport: TQRCompositeReport;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    qryProfile: TQuery;
    pnlCond: TPanel;
    Label7: TLabel;
    Label2: TLabel;
    btnBdate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    cmbJisa: TComboBox;
    edtBDate: TDateEdit;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    BunjuNo1: TMaskEdit;
    BunjuNo2: TMaskEdit;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008 :Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD5Q11: TfrmLD5Q11;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD5Q11.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD5Q11.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD5Q11.UP_VarMemSet(iLength : Integer);
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
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_v001,  iLength);
      VarArrayRedim(UV_v002,  iLength);
      VarArrayRedim(UV_v003,  iLength);
      VarArrayRedim(UV_v004,  iLength);
      VarArrayRedim(UV_v005,  iLength);
      VarArrayRedim(UV_v006,  iLength);
      VarArrayRedim(UV_v007,  iLength);
      VarArrayRedim(UV_v008,  iLength);
   end;
end;

function TfrmLD5Q11.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check

end;

procedure TfrmLD5Q11.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);

end;

procedure TfrmLD5Q11.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1  : Value := UV_v001[DataRow-1];
      2  : Value := UV_v002[DataRow-1];
      3  : Value := UV_v003[DataRow-1];
      4  : Value := UV_v004[DataRow-1];
      5  : Value := UV_v005[DataRow-1];
      6  : Value := UV_v006[DataRow-1];
      7  : Value := UV_v007[DataRow-1];
      8  : Value := UV_v008[DataRow-1];
   end;
end;

procedure TfrmLD5Q11.btnQueryClick(Sender: TObject);
var index: Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';

   // 조회조건 Check
   if not UF_Condition then exit;
   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT B.dat_bunju, B.num_Bunju, Mun1_Jindan3, Mun1_Drug3, Mun1_Jindan4, Mun1_Drug4, Mun1_Jindan5, Mun1_Drug5 ' +
                 ' FROM bunju B with(nolock) join tot_munjin2010 M ON B.cod_jisa = M.cod_jisa and B.num_id = M.num_id and B.dat_gmgn = M.dat_gmgn ';
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' ';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
      sWhere := sWhere + ' AND B.num_bunju >= ''' + BunjuNo1.Text + ''' ';
      sWhere := sWhere + ' AND B.num_bunju <= ''' + BunjuNo2.Text + ''' ';
      sOrderBy := ' ORDER BY B.cod_bjjs, B.num_bunju ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD5Q11', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            UP_VarMemSet(index);

            UV_v001[index] := FieldByName('dat_bunju').AsString;
            UV_v002[index] := FieldByName('num_bunju').AsString;
            if      FieldByName('Mun1_Jindan3').AsString = '0' then UV_v003[index] := 'N'
            else if FieldByName('Mun1_Jindan3').AsString = '1' then UV_v003[index] := 'Y';
            if      FieldByName('Mun1_Drug3').AsString = '0'   then UV_v004[index] := 'N'
            else if FieldByName('Mun1_Drug3').AsString = '1'   then UV_v004[index] := 'Y';
            if      FieldByName('Mun1_Jindan4').AsString = '0' then UV_v005[index] := 'N'
            else if FieldByName('Mun1_Jindan4').AsString = '1' then UV_v005[index] := 'Y';
            if      FieldByName('Mun1_Drug4').AsString = '0'   then UV_v006[index] := 'N'
            else if FieldByName('Mun1_Drug4').AsString = '1'   then UV_v006[index] := 'Y';
            if      FieldByName('Mun1_Jindan5').AsString = '0' then UV_v007[index] := 'N'
            else if FieldByName('Mun1_Jindan5').AsString = '1' then UV_v007[index] := 'Y';
            if      FieldByName('Mun1_Drug5').AsString ='0'    then UV_v008[index] := 'N'
            else if FieldByName('Mun1_Drug5').AsString ='1'    then UV_v008[index] := 'Y';
            inc(index);

            qryBunju.Next;
         End;
         // Table과 Disconnect
         qryBunju.Close;
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

procedure TfrmLD5Q11.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;



end;

procedure TfrmLD5Q11.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD5Q11.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;
procedure TfrmLD5Q11.SBut_ExcelClick(Sender: TObject);
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
