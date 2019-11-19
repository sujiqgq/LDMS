//==============================================================================
// 프로그램 설명 : Architect(누락) 조회
// 시스템        : 분석관리
// 수정일자      : 2013.05.23
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD4Q45;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q45 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    qryBunju: TQuery;
    Gride: TGauge;
    qryPf_hm: TQuery;
    Label2: TLabel;
    mskDate: TDateEdit;
    btnBdate: TSpeedButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    mskTo: TMaskEdit;
    Label10: TLabel;
    mskFrom: TMaskEdit;
    Label9: TLabel;
    qryProfileG: TQuery;
    qryNo_hangmok: TQuery;
    qryTkgum: TQuery;
    RadioButton2: TRadioButton;
    RBtn_preveiw: TRadioButton;
    RadioGroup: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure FormActivate(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vRackNo, UV_vNum_samp, UV_vName, UV_vNum_Bunju   : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q45: TfrmLD4Q45;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q451;

{$R *.DFM}

procedure TfrmLD4Q45.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;
end;

procedure TfrmLD4Q45.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q45.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vRackNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp  := VarArrayCreate([0, 0], varOleStr);
      UV_vName      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_Bunju := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vRackNo,    iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vName,      iLength);
      VarArrayRedim(UV_vNum_Bunju, iLength);
   end;
end;

function TfrmLD4Q45.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if mskDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q45.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);

   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q45.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vRackNo[DataRow-1];
      2 : Value := UV_vNum_samp[DataRow-1];
      3 : Value := UV_vName[DataRow-1];
      4 : Value := UV_vNum_Bunju[DataRow-1];
   end;
end;

procedure TfrmLD4Q45.btnQueryClick(Sender: TObject);
var index : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   grdMaster.Rows := 0;
   
   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   with qryBunju do
   begin
      // SQL문 생성
      Close;

      sSelect := ' select B.desc_RackNo, G.num_samp, G.desc_name, B.num_bunju ' +
                 ' from bunju B Inner Join gumgin G ON B.Cod_jisa = G.Cod_jisa and B.num_id = G.num_id and B.dat_gmgn = G.dat_gmgn ';

      sWhere := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere := sWhere + ' AND  B.Dat_Bunju = ''' + mskDate.Text + '''' +
                         ' AND  B.num_bunju not in ( select num_bunju from Bunju_Interface ' +
                         ' where cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''' +
                         ' and Dat_Bunju = ''' + mskDate.Text + '''' +
                         ' and gubn_machine like ''Architect C%'')';

      case RadioGroup.ItemIndex of
         0 : sWhere := sWhere + ' AND B.desc_RackNo <> ''''';
         1 : sWhere := sWhere + ' AND SUBSTRING(B.desc_RackNo,1,1) = ''A''';
         2 : sWhere := sWhere + ' AND SUBSTRING(B.desc_RackNo,1,1) = ''B''';
         3 : sWhere := sWhere + ' AND SUBSTRING(B.desc_RackNo,1,1) = ''C''';
      end;



      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(mskFrom.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
         if Trim(mskTo.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end;



      sOrderBy := ' ORDER BY B.desc_RackNo ';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;

      Gride.MaxValue := qryBunju.RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q45', 'Q', 'N');

         while not Eof do
         Begin
            Gride.Progress := Gride.Progress + 1;
            UP_VarMemSet(Index);

            UV_vRackNo[Index]    := qryBunju.FieldByName('desc_RackNo').AsString;
            UV_vNum_samp[Index]  := qryBunju.FieldByName('Num_samp').AsString;
            UV_vName[Index]      := qryBunju.FieldByName('desc_name').AsString;
            UV_vNum_Bunju[Index] := qryBunju.FieldByName('num_Bunju').AsString;

            Inc(Index);

            Next;
         end;
      end
      else
      begin
         // 자료가 없음을 표시
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
   end;  // with qry_gulkwa do

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q45.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnBdate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD4Q45.btnPrintClick(Sender: TObject);
begin
  inherited;

  if RBtn_preveiw.Checked = True then
  begin
     frmLD4Q451 := TfrmLD4Q451.Create(Self);
     frmLD4Q451.QuickRep.Preview;
  end
  else
  begin
     frmLD4Q451 := TfrmLD4Q451.Create(Self);
     frmLD4Q451.QuickRep.Print;
  end;

end;

procedure TfrmLD4Q45.FormActivate(Sender: TObject);
begin
  inherited;
   Cmb_gubn.ItemIndex := 1;
   mskDate.SetFocus;
end;

procedure TfrmLD4Q45.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;
   
   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;

end.
