//================================================================================
// 수정일자      : 2013.7.23
// 수정자        : 김희석
// 수정내용      : 한의 재단진단면역학팀1300097
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.29
// 수정자        : 곽윤설
// 수정내용      : U024,U029,U030,U031,U032,U044,U051 제거
// 참고사항      : U020,U023,U028,U037,U043,U045,U047 제거
//==============================================================================
unit LD4Q49;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q49 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    Label5: TLabel;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label4: TLabel;
    Cmb_gubn: TComboBox;
    Chk_Self: TRadioButton;
    Chk_outside: TRadioButton;
    qryJHangmokList: TQuery;
    group_radio: TGroupBox;
    qryPf_hm: TQuery;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN,
        UV_vU001,UV_vU002,UV_vU003,UV_vU004,UV_vU005,UV_vU006,UV_vU007,UV_vU008,UV_vU009,UV_vU010,UV_vU011,UV_vU012,UV_vU013,UV_vU017,
    UV_vU024,UV_vU029,UV_vU030,UV_vU031,UV_vU032,UV_vU033,UV_vU044,UV_vU046,UV_vU051,UV_vY004,
    UV_vU019,UV_vU020,UV_vU023,UV_vU027,UV_vU028,UV_vU034,UV_vU035,UV_vU036,UV_vU037,UV_vU038,UV_vU039,UV_vU040,UV_vU041,UV_vU042,UV_vU043,
    UV_vU045,UV_vU047,UV_vU048,UV_vU049,UV_vU050,UV_vU052,UV_vY005,UV_vY008,UV_vSampNo, UV_vDesc_dc : Variant;



    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q49: TfrmLD4Q49;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,LD4Q491;

{$R *.DFM}

procedure TfrmLD4Q49.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q49.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q49.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vU001  := VarArrayCreate([0, 0], varOleStr);
      UV_vU002  := VarArrayCreate([0, 0], varOleStr);
      UV_vU003  := VarArrayCreate([0, 0], varOleStr);
      UV_vU004  := VarArrayCreate([0, 0], varOleStr);
      UV_vU005  := VarArrayCreate([0, 0], varOleStr);
      UV_vU006  := VarArrayCreate([0, 0], varOleStr);
      UV_vU007  := VarArrayCreate([0, 0], varOleStr);
      UV_vU008  := VarArrayCreate([0, 0], varOleStr);
      UV_vU009  := VarArrayCreate([0, 0], varOleStr);
      UV_vU010  := VarArrayCreate([0, 0], varOleStr);
      UV_vU011  := VarArrayCreate([0, 0], varOleStr);
      UV_vU012  := VarArrayCreate([0, 0], varOleStr);
      UV_vU013  := VarArrayCreate([0, 0], varOleStr);
      UV_vU017  := VarArrayCreate([0, 0], varOleStr);
      UV_vU024  := VarArrayCreate([0, 0], varOleStr);
      UV_vU029  := VarArrayCreate([0, 0], varOleStr);
      UV_vU030  := VarArrayCreate([0, 0], varOleStr);
      UV_vU031  := VarArrayCreate([0, 0], varOleStr);
      UV_vU032  := VarArrayCreate([0, 0], varOleStr);
      UV_vU033  := VarArrayCreate([0, 0], varOleStr);
      UV_vU044  := VarArrayCreate([0, 0], varOleStr);
      UV_vU046  := VarArrayCreate([0, 0], varOleStr);
      UV_vU051  := VarArrayCreate([0, 0], varOleStr);
      UV_vY004  := VarArrayCreate([0, 0], varOleStr);

      UV_vU019  := VarArrayCreate([0, 0], varOleStr);
      UV_vU020  := VarArrayCreate([0, 0], varOleStr);
      UV_vU023  := VarArrayCreate([0, 0], varOleStr);
      UV_vU027  := VarArrayCreate([0, 0], varOleStr);
      UV_vU028  := VarArrayCreate([0, 0], varOleStr);
      UV_vU034  := VarArrayCreate([0, 0], varOleStr);
      UV_vU035  := VarArrayCreate([0, 0], varOleStr);
      UV_vU036  := VarArrayCreate([0, 0], varOleStr);
      UV_vU037  := VarArrayCreate([0, 0], varOleStr);
      UV_vU038  := VarArrayCreate([0, 0], varOleStr);
      UV_vU039  := VarArrayCreate([0, 0], varOleStr);
      UV_vU040  := VarArrayCreate([0, 0], varOleStr);
      UV_vU041  := VarArrayCreate([0, 0], varOleStr);
      UV_vU042  := VarArrayCreate([0, 0], varOleStr);
      UV_vU043  := VarArrayCreate([0, 0], varOleStr);
      UV_vU045  := VarArrayCreate([0, 0], varOleStr);
      UV_vU047  := VarArrayCreate([0, 0], varOleStr);
      UV_vU048  := VarArrayCreate([0, 0], varOleStr);
      UV_vU049  := VarArrayCreate([0, 0], varOleStr);
      UV_vU050  := VarArrayCreate([0, 0], varOleStr);
      UV_vU052  := VarArrayCreate([0, 0], varOleStr);
      UV_vY005  := VarArrayCreate([0, 0], varOleStr);
      UV_vY008  := VarArrayCreate([0, 0], varOleStr);


      UV_vSampNo:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);

      VarArrayRedim(UV_vU001,  iLength);
      VarArrayRedim(UV_vU002,  iLength);
      VarArrayRedim(UV_vU003,  iLength);
      VarArrayRedim(UV_vU004,  iLength);
      VarArrayRedim(UV_vU005,  iLength);
      VarArrayRedim(UV_vU006,  iLength);
      VarArrayRedim(UV_vU007,  iLength);
      VarArrayRedim(UV_vU008,  iLength);
      VarArrayRedim(UV_vU009,  iLength);
      VarArrayRedim(UV_vU010,  iLength);
      VarArrayRedim(UV_vU011,  iLength);
      VarArrayRedim(UV_vU012,  iLength);
      VarArrayRedim(UV_vU013,  iLength);
      VarArrayRedim(UV_vU017,  iLength);
      VarArrayRedim(UV_vU024,  iLength);
      VarArrayRedim(UV_vU029,  iLength);
      VarArrayRedim(UV_vU030,  iLength);
      VarArrayRedim(UV_vU031,  iLength);
      VarArrayRedim(UV_vU032,  iLength);
      VarArrayRedim(UV_vU033,  iLength);
      VarArrayRedim(UV_vU044,  iLength);
      VarArrayRedim(UV_vU046,  iLength);
      VarArrayRedim(UV_vU051,  iLength);
      VarArrayRedim(UV_vY004,  iLength);

      VarArrayRedim(UV_vU019,  iLength);
      VarArrayRedim(UV_vU020,  iLength);
      VarArrayRedim(UV_vU023,  iLength);
      VarArrayRedim(UV_vU027,  iLength);
      VarArrayRedim(UV_vU028,  iLength);
      VarArrayRedim(UV_vU034,  iLength);
      VarArrayRedim(UV_vU035,  iLength);
      VarArrayRedim(UV_vU036,  iLength);
      VarArrayRedim(UV_vU037,  iLength);
      VarArrayRedim(UV_vU038,  iLength);
      VarArrayRedim(UV_vU039,  iLength);
      VarArrayRedim(UV_vU040,  iLength);
      VarArrayRedim(UV_vU041,  iLength);
      VarArrayRedim(UV_vU042,  iLength);
      VarArrayRedim(UV_vU043,  iLength);
      VarArrayRedim(UV_vU045,  iLength);
      VarArrayRedim(UV_vU047,  iLength);
      VarArrayRedim(UV_vU048,  iLength);
      VarArrayRedim(UV_vU049,  iLength);
      VarArrayRedim(UV_vU050,  iLength);
      VarArrayRedim(UV_vU052,  iLength);
      VarArrayRedim(UV_vY005,  iLength);
      VarArrayRedim(UV_vY008,  iLength);



      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD4Q49.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q49.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q49.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;

   // 자룔를 화면에 조회

   if Chk_Self.Checked =true  then
   begin
        case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := UV_vU001[DataRow-1];
           6 : Value := UV_vU002[DataRow-1];
           7 : Value := UV_vU003[DataRow-1];
           8 : Value := UV_vU004[DataRow-1];
           9 : Value := UV_vU005[DataRow-1];
          10 : Value := UV_vU006[DataRow-1];
          11 : Value := UV_vU007[DataRow-1];
          12 : Value := UV_vU008[DataRow-1];
          13 : Value := UV_vU009[DataRow-1];
          14 : Value := UV_vU010[DataRow-1];
          15 : Value := UV_vU011[DataRow-1];
          16 : Value := UV_vU012[DataRow-1];
          17 : Value := UV_vU013[DataRow-1];
          18 : Value := UV_vU017[DataRow-1];
          19 : Value := UV_vU033[DataRow-1];
          20 : Value := UV_vU046[DataRow-1];
          21 : Value := UV_vY004[DataRow-1];
          end;
          grdMaster.Col[5].heading  :='U001';
          grdMaster.Col[6].heading  :='U002';
          grdMaster.Col[7].heading  :='U003';
          grdMaster.Col[8].heading  :='U004';
          grdMaster.Col[9].heading  :='U005';
          grdMaster.Col[10].heading :='U006';
          grdMaster.Col[11].heading :='U007';
          grdMaster.Col[12].heading :='U008';
          grdMaster.Col[13].heading :='U009';
          grdMaster.Col[14].heading :='U010';
          grdMaster.Col[15].heading :='U011';
          grdMaster.Col[16].heading :='U012';
          grdMaster.Col[17].heading :='U013';
          grdMaster.Col[18].heading :='U017';
          grdMaster.Col[19].heading :='U033';
          grdMaster.Col[20].heading :='U046';
          grdMaster.Col[21].heading :='Y004';
    end
    else if Chk_outside.Checked =true  then
    begin
        case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := UV_vU019[DataRow-1];
           6 : Value := UV_vU027[DataRow-1];
           7 : Value := UV_vU034[DataRow-1];
           8 : Value := UV_vU035[DataRow-1];
           9 : Value := UV_vU036[DataRow-1];
          10 : Value := UV_vU038[DataRow-1];
          11 : Value := UV_vU039[DataRow-1];
          12 : Value := UV_vU040[DataRow-1];
          13 : Value := UV_vU041[DataRow-1];
          14 : Value := UV_vU042[DataRow-1];
          15 : Value := UV_vU048[DataRow-1];
          16 : Value := UV_vU049[DataRow-1];
          17 : Value := UV_vU050[DataRow-1];
          18 : Value := UV_vU052[DataRow-1];
          19 : Value := UV_vY005[DataRow-1];
          20 : Value := UV_vY008[DataRow-1];
          end;

          grdMaster.Col[5].heading  :='U019';
          grdMaster.Col[6].heading  :='U027';
          grdMaster.Col[7].heading  :='U034';
          grdMaster.Col[8].heading  :='U035';
          grdMaster.Col[9].heading  :='U036';
          grdMaster.Col[10].heading :='U038';
          grdMaster.Col[11].heading :='U039';
          grdMaster.Col[12].heading :='U040';
          grdMaster.Col[13].heading :='U041';
          grdMaster.Col[14].heading :='U042';
          grdMaster.Col[15].heading :='U048';
          grdMaster.Col[16].heading :='U049';
          grdMaster.Col[17].heading :='U050';
          grdMaster.Col[18].heading :='U052';
          grdMaster.Col[19].heading :='Y005';
          grdMaster.Col[20].heading :='Y008';
          grdMaster.Col[21].heading :='';
    end;

end;

procedure TfrmLD4Q49.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vU001,Sum_vU002,Sum_vU003,Sum_vU004,Sum_vU005,Sum_vU006,Sum_vU007,Sum_vU008,Sum_vU009,Sum_vU010,
    Sum_vU011,Sum_vU012,Sum_vU013,Sum_vU017,Sum_vU033,Sum_vU046,Sum_vU051,Sum_vY004,
    Sum_vU019,Sum_vU020,Sum_vU023,Sum_vU027,Sum_vU034,Sum_vU035,Sum_vU036,Sum_vU038,Sum_vU039,Sum_vU040,Sum_vU041,Sum_vU042,
    Sum_vU045,Sum_vU048,Sum_vU049,Sum_vU050,Sum_vU052,Sum_vY005,Sum_vY008 : integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    UV_fGulkwa_07, UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07  : String;
    vCod_chuga : Variant;
    Check :Boolean;
     Check_05 :Boolean;
      Check_07 :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;



   Sum_vU001 :=0; Sum_vU002 :=0; Sum_vU003 :=0; Sum_vU004 :=0;
   Sum_vU005 :=0; Sum_vU006 :=0; Sum_vU007 :=0; Sum_vU008 :=0;
   Sum_vU009 :=0; Sum_vU010 :=0; Sum_vU011 :=0; Sum_vU012 :=0;
   Sum_vU013 :=0; Sum_vU017 :=0; Sum_vU033 :=0; Sum_vU046 :=0;
   Sum_vY004 :=0; Sum_vU019 :=0; Sum_vU027 :=0; Sum_vU034 :=0;
   Sum_vU035 :=0; Sum_vU036 :=0; Sum_vU038 :=0; Sum_vU039 :=0;
   Sum_vU040 :=0; Sum_vU041 :=0; Sum_vU042 :=0; Sum_vU048 :=0;
   Sum_vU049 :=0; Sum_vU050 :=0; Sum_vU052 :=0; Sum_vY005 :=0;
   Sum_vY008 :=0;


   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT  B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B inner join gumgin G on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha'+
                 '         left outer join gulkwa K  ON G.cod_jisa = k.cod_jisa and G.num_id = K.num_id and G.dat_gmgn = K.dat_gmgn ';

      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + '''  and gubn_part=''03''';



      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';   
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno ';



      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q49', 'Q', 'N');
         while not qryBunju.Eof do
         Begin

            sHul_List := '';

            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;
            
            sHul_List := UF_hangmokList;
            

            if ((Pos('U001' ,sHul_List)> 0) or (Pos('U008' ,sHul_List)> 0) or (Pos('U046' ,sHul_List)> 0) or
                (Pos('U002' ,sHul_List)> 0) or (Pos('U009' ,sHul_List)> 0) or (Pos('U051' ,sHul_List)> 0) or
                (Pos('U003' ,sHul_List)> 0) or (Pos('U010' ,sHul_List)> 0) or (Pos('Y004' ,sHul_List)> 0) or
                (Pos('U004' ,sHul_List)> 0) or (Pos('U011' ,sHul_List)> 0) or (Pos('U005' ,sHul_List)> 0) or
                (Pos('U012' ,sHul_List)> 0) or (Pos('U006' ,sHul_List)> 0) or (Pos('U013' ,sHul_List)> 0) or
                (Pos('U033' ,sHul_List)> 0) or (Pos('U007' ,sHul_List)> 0) or (Pos('U017' ,sHul_List)> 0)) and (Chk_Self.Checked=true)  Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         Check := False;
                         UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         UV_vNo[Index]     := index+1;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('U001', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                                  begin
                                  UV_vU001[Index] := '결과無';
                                  Sum_vU001 := Sum_vU001 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('U002', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 13,  6)) = '') then
                                  begin
                                  UV_vU002[Index] := '결과無';
                                  Sum_vU002 := Sum_vU002 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('U003', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                                  begin
                                  UV_vU003[Index] := '결과無';
                                  Sum_vU003 := Sum_vU003 + 1;
                                  Check := True;
                               end;
                          end;
                             ///////////////////////////////////////////////////////////
                          if (pos('U004', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                                  begin
                                  UV_vU004[Index] := '결과無';
                                  Sum_vU004 := Sum_vU004 + 1;
                                  Check := True;
                               end;
                          end;
                              ///////////////////////////////////////////////////////////
                          if (pos('U005', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 31,  6)) = '') then
                                  begin
                                  UV_vU005[Index] := '결과無';
                                  Sum_vU005 := Sum_vU005 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('U006', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                                  begin
                                  UV_vU006[Index] := '결과無';
                                  Sum_vU006 := Sum_vU006 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('U007', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 43,  6)) = '') then
                                  begin
                                  UV_vU007[Index] := '결과無';
                                  Sum_vU007 := Sum_vU007 + 1;
                                  Check := True;
                               end;
                          end;
                                 ///////////////////////////////////////////////////////////
                          if (pos('U008', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 49,  6)) = '') then
                                  begin
                                  UV_vU008[Index] := '결과無';
                                  Sum_vU008 := Sum_vU008 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('U009', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 55,  6)) = '') then
                                  begin
                                  UV_vU009[Index] := '결과無';
                                  Sum_vU009 := Sum_vU009 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('U010', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 61,  6)) = '') then
                                  begin
                                  UV_vU010[Index] := '결과無';
                                  Sum_vU010 := Sum_vU010 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('U011', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                                  begin
                                  UV_vU011[Index] := '결과無';
                                  Sum_vU011:= Sum_vU011 + 1;
                                  Check := True;
                               end;
                          end;
                                    ///////////////////////////////////////////////////////////
                          if (pos('U012', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 73,  6)) = '') then
                                  begin
                                  UV_vU012[Index] := '결과無';
                                  Sum_vU012:= Sum_vU012 + 1;
                                  Check := True;
                               end;
                          end;
                                     ///////////////////////////////////////////////////////////
                          if (pos('U013', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 79,  6)) = '') then
                                  begin
                                  UV_vU013[Index] := '결과無';
                                  Sum_vU013:= Sum_vU013 + 1;
                                  Check := True;
                               end;
                          end;
                                      ///////////////////////////////////////////////////////////
                          if (pos('U017', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 115,  6)) = '') then
                                  begin
                                  UV_vU017[Index] := '결과無';
                                  Sum_vU017:= Sum_vU017 + 1;
                                  Check := True;
                               end;
                          end;
                                        ///////////////////////////////////////////////////////////
                          if (pos('U033', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 199,  6)) = '') then
                                  begin
                                  UV_vU033[Index] := '결과無';
                                  Sum_vU033:= Sum_vU033 + 1;
                                  Check := True;
                               end;
                          end;
                                        ///////////////////////////////////////////////////////////
                          if (pos('U046', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 289,  6)) = '') then
                                  begin
                                  UV_vU046[Index] := '결과無';
                                  Sum_vU046:= Sum_vU046 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('Y004', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 85,  6)) = '') then
                                  begin
                                  UV_vY004[Index] := '결과無';
                                  Sum_vY004:= Sum_vY004 + 1;
                                  Check := True;
                               end;
                          end;

                          if Check = True then
                          begin
                               UP_VarMemSet(Index+1);
                               inc(index);
                          end;
                   qryGulkwa.Next;
                   end;
               End;
               qryGulkwa.Close;

              end
              else if ((Pos('U019' ,sHul_List)> 0) or (Pos('U041' ,sHul_List)> 0) or (Pos('U034' ,sHul_List)> 0) or
                       (Pos('U038' ,sHul_List)> 0) or (Pos('U042' ,sHul_List)> 0) or (Pos('U035' ,sHul_List)> 0) or
                       (Pos('U039' ,sHul_List)> 0) or (Pos('U027' ,sHul_List)> 0) or (Pos('U036' ,sHul_List)> 0) or
                       (Pos('U040' ,sHul_List)> 0) or (Pos('U048' ,sHul_List)> 0) or (Pos('U049' ,sHul_List)> 0) or
                       (Pos('U050' ,sHul_List)> 0) or (Pos('U052' ,sHul_List)> 0) or (Pos('Y005' ,sHul_List)> 0) or
                       (Pos('Y008' ,sHul_List)> 0))and (Chk_outside.Checked =True )Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin

                         Check := False;
                         UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                         UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                         UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                         GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

                         UV_vNo[Index]     := index+1;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('U019', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 127,  6)) = '') then
                                  begin
                                  UV_vU019[Index] := '결과無';
                                  Sum_vU019 := Sum_vU019 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('U027', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 169,  6)) = '') then
                                  begin
                                  UV_vU027[Index] := '결과無';
                                  Sum_vU027 := Sum_vU027 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////
                          if (pos('U034', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 205,  6)) = '') then
                                  begin
                                  UV_vU034[Index] := '결과無';
                                  Sum_vU034 := Sum_vU034 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('U035', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 211,  6)) = '') then
                                  begin
                                  UV_vU035[Index] := '결과無';
                                  Sum_vU035 := Sum_vU035 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('U036', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 217,  6)) = '') then
                                  begin
                                  UV_vU036[Index] := '결과無';
                                  Sum_vU036 := Sum_vU036 + 1;
                                  Check := True;
                               end;
                          end;
                             ///////////////////////////////////////////////////////////
                          if (pos('U038', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 229,  6)) = '') then
                                  begin
                                  UV_vU038[Index] := '결과無';
                                  Sum_vU038 := Sum_vU038 + 1;
                                  Check := True;
                               end;
                          end;
                              ///////////////////////////////////////////////////////////
                          if (pos('U039', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 247,  6)) = '') then
                                  begin
                                  UV_vU039[Index] := '결과無';
                                  Sum_vU039 := Sum_vU039 + 1;
                                  Check := True;
                               end;
                          end;
                               ///////////////////////////////////////////////////////////
                          if (pos('U040', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 253,  6)) = '') then
                                  begin
                                  UV_vU040[Index] := '결과無';
                                  Sum_vU040 := Sum_vU040 + 1;
                                  Check := True;
                               end;
                          end;
                                ///////////////////////////////////////////////////////////
                          if (pos('U041', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 259,  6)) = '') then
                                  begin
                                  UV_vU041[Index] := '결과無';
                                  Sum_vU041 := Sum_vU041 + 1;
                                  Check := True;
                               end;
                          end;
                                  ///////////////////////////////////////////////////////////
                          if (pos('U042', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 265,  6)) = '') then
                                  begin
                                  UV_vU042[Index] := '결과無';
                                  Sum_vU042 := Sum_vU042 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('U048', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 301,  6)) = '') then
                                  begin
                                  UV_vU048[Index] := '결과無';
                                  Sum_vU048 := Sum_vU048 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('U049', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 307,  6)) = '') then
                                  begin
                                  UV_vU049[Index] := '결과無';
                                  Sum_vU049 := Sum_vU049 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('U050', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 313,  6)) = '') then
                                  begin
                                  UV_vU050[Index] := '결과無';
                                  Sum_vU050 := Sum_vU050 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('Y005', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 91,  6)) = '') then
                                  begin
                                  UV_vY005[Index] := '결과無';
                                  Sum_vY005 := Sum_vY005 + 1;
                                  Check := True;
                               end;
                          end;
                                   ///////////////////////////////////////////////////////////
                          if (pos('Y008', sHul_List) > 0)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 235,  6)) = '') then
                                  begin
                                  UV_vY008[Index] := '결과無';
                                  Sum_vY008 := Sum_vY008 + 1;
                                  Check := True;
                               end;
                          end;

                          if Check = True then
                          begin
                               UP_VarMemSet(Index+1);
                               inc(index);
                          end;
                   qryGulkwa.Next;
                   end;
               End;
               qryGulkwa.Close;

              end;
           qryBunju.Next;
         End;


         // Table과 Disconnect
         qryBunju.Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '총  합  계';
         UV_vU001[Index]:= Sum_vU001;
         UV_vU002[Index]:= Sum_vU002;
         UV_vU003[Index]:= Sum_vU003;
         UV_vU004[Index]:= Sum_vU004;
         UV_vU005[Index]:= Sum_vU005;
         UV_vU006[Index]:= Sum_vU006;
         UV_vU007[Index]:= Sum_vU007;
         UV_vU008[Index]:= Sum_vU008;
         UV_vU009[Index]:= Sum_vU009;
         UV_vU010[Index]:= Sum_vU010;
         UV_vU011[Index]:= Sum_vU011;
         UV_vU012[Index]:= Sum_vU012;
         UV_vU013[Index]:= Sum_vU013;
         UV_vU017[Index]:= Sum_vU017;
         UV_vU033[Index]:= Sum_vU033;
         UV_vU046[Index]:= Sum_vU046;
         UV_vY004[Index]:= Sum_vY004;
         UV_vU019[Index]:= Sum_vU019;
         UV_vU027[Index]:= Sum_vU027;
         UV_vU034[Index]:= Sum_vU034;
         UV_vU035[Index]:= Sum_vU035;
         UV_vU036[Index]:= Sum_vU036;
         UV_vU038[Index]:= Sum_vU038;
         UV_vU039[Index]:= Sum_vU039;
         UV_vU040[Index]:= Sum_vU040;
         UV_vU041[Index]:= Sum_vU041;
         UV_vU042[Index]:= Sum_vU042;
         UV_vU048[Index]:= Sum_vU048;
         UV_vU049[Index]:= Sum_vU049;
         UV_vU050[Index]:= Sum_vU050;
         UV_vU052[Index]:= Sum_vU052;
         UV_vY005[Index]:= Sum_vY005;
         UV_vY008[Index]:= Sum_vY008;

         UV_vSampNo[Index]  := '';
         Inc(index);
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

procedure TfrmLD4Q49.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q49.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q491 := TfrmLD4Q491.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD4Q491.QuickRep.Preview
  else                                frmLD4Q491.QuickRep.Print;

end;

procedure TfrmLD4Q49.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;


end;

function TfrmLD4Q49.UF_hangmokList : String;
var
   i : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
       sHmCode :string;
   vCod_chuga : Variant;
   iRet : integer;
begin





   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
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

      // 2. 종양에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
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

      // 3. 장비에 대한 검사항목 추출
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
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

   // 3. 추가코드에 대한 검사항목 추출
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. 항목코드에 대한 검사항목 추출
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
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
         qryProfileG.ParamByName('cod_pf25').AsString := Copy(qryTkgum.FieldByName('cod_prf').AsString, 97, 4);
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
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;

   result := sTemp;
end;

function TfrmLD4Q49.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
