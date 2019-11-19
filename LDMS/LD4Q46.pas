//==============================================================================
// 프로그램 설명 : 혈액,소변 미검 리스트   (접수기준)
// 시스템        : 통합검진
// 수정일자      : 2013.0607
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================

unit LD4Q46;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q46 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    cmbJisa: TComboBox;
    qry_gumgin: TQuery;
    Gride: TGauge;
    Label11: TLabel;
    mskST_date: TDateEdit;
    btnGmgnF: TSpeedButton;
    Label10: TLabel;
    mskEd_date: TDateEdit;
    btnGmgnT: TSpeedButton;
    Label12: TLabel;
    Rb_All: TRadioButton;
    Rb_urin: TRadioButton;
    Rb_blood: TRadioButton;
    Panel2: TPanel;
    RBtn_preveiw: TRadioButton;
    RadioButton1: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure UP_Click(Sender: TObject);


     procedure MouseWheelHandler(var Message: TMessage); override;

    procedure FormActivate(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vDat_gmgn, UV_vNum_samp,UV_vBunju,UV_vUrin,UV_vBlood   : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);


    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q46: TfrmLD4Q46;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q461;

{$R *.DFM}

procedure TfrmLD4Q46.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;



end;

procedure TfrmLD4Q46.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q46.UP_VarMemSet(iLength : Integer);
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
      UV_vNum_samp   := VarArrayCreate([0, 0], varOleStr);
      UV_vBunju      := VarArrayCreate([0, 0], varOleStr);
      UV_vUrin       := VarArrayCreate([0, 0], varOleStr);
      UV_vBlood      := VarArrayCreate([0, 0], varOleStr);


  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vNum_samp,  iLength);
      VarArrayRedim(UV_vBunju,  iLength);
      VarArrayRedim(UV_vUrin,  iLength);
      VarArrayRedim(UV_vBlood,  iLength);

   end;
end;

function TfrmLD4Q46.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (mskST_date.Text = '') or (mskEd_date.Text = '')  then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q46.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q46.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vNum_Samp[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vDat_gmgn[DataRow-1];
      7 : Value := UV_vBlood[DataRow-1];
      8 : Value := UV_vURin[DataRow-1];
   end;
end;

procedure TfrmLD4Q46.btnQueryClick(Sender: TObject);
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
   with qry_gumgin do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select A.cod_jisa, A.num_id, A.dat_gmgn, A.num_samp, A.desc_name,A.num_jumin,C.Ck_urin,C.Ck_Blood ';
      sSelect := sSelect + ' From Gumgin_check C   left outer join Gumgin A ';
      sSelect := sSelect + ' On C.cod_jisa = A.cod_jisa and a.num_id = C.num_id and a.dat_gmgn = C.dat_gmgn ';
      sWhere := ' WHERE A.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere := sWhere + '   AND A.dat_gmgn >= ''' + mskST_date.Text + '''';
      sWhere := sWhere + '   AND A.dat_gmgn <= ''' + mskEd_date.Text + '''';

      if rb_All.Checked =true then
      begin
          sWhere := sWhere + '   AND (C.CK_Urin=''Y'' or C.CK_Blood=''Y'')';

      end
      else if rb_Urin.Checked =true then
      begin
          sWhere := sWhere + '   AND C.CK_Urin=''Y'' ';
      end
      else if rb_Blood.Checked =true then
      begin
          sWhere := sWhere + '   AND C.CK_Blood=''Y'' ';
      end;



      sOrderBy := ' ORDER BY CAST(A.num_samp AS INT)';


      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;

      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qry_gumgin, 'LD4Q37', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := IntToStr(Index+1);
               UV_vJisa[Index] := qry_gumgin.FieldByName('Cod_jisa').AsString;
               UV_vName[Index] := qry_gumgin.FieldByName('desc_name').AsString;
               UV_vJumin[Index] :=GF_JuminFormat(qry_gumgin.FieldByName('num_jumin').AsString);
               UV_vDat_gmgn[Index] := GF_DateFormat(qry_gumgin.FieldByName('dat_gmgn').AsString);
               UV_vnum_samp[Index] := qry_gumgin.FieldByName('num_samp').AsString;
               if qry_gumgin.FieldByName('CK_Urin').AsString ='Y' then
               begin
                  UV_vUrin[Index]  :='미검'
               end
               else UV_vUrin[Index]  :=' ';
               if qry_gumgin.FieldByName('CK_blood').AsString ='Y'  then
               begin
                  UV_vBlood[Index]   :='미검'
               end
               else UV_vBlood[Index] :=' ';

               if Rb_urin.Checked =true then
                  UV_vBlood[Index] :=' ';
               if Rb_blood.Checked =true then
                  UV_vUrin[Index]  :=' ';

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



procedure TfrmLD4Q46.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnGmgnF then GF_CalendarClick(mskST_date);
   if Sender = btnGmgnT then GF_CalendarClick(mskEd_date);
end;

procedure TfrmLD4Q46.FormActivate(Sender: TObject);
begin
  inherited;


   mskST_date.SetFocus;
end;

procedure TfrmLD4Q46.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q461 := TfrmLD4Q461.Create(Self);

  if RBtn_preveiw.Checked = True then frmLD4Q461.QuickRep.Preview
  else                                frmLD4Q461.QuickRep.Print;
end;

end.
