//==============================================================================
// 프로그램 설명 : 검사항목 결과오류(혈액학
// 시스템        : 통합검진
// 수정일자      : 20130625
// 수정자        : 김희석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.12.11
// 수정자        : 곽윤설
// 수정내용      : 특검 물질 갯수제한 없이 끌어오도록 수정
// 참고사항      :
//==============================================================================

unit LD4Q48;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD4Q48 = class(TfrmSingle)
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
    qryTkgum: TQuery;
    qryProfile: TQuery;
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
    UV_vH001,UV_vH002,UV_vH003,UV_vH004,UV_vH005,UV_vH006,UV_vH007,UV_vH008,UV_vH009,UV_vH010,UV_vH011,UV_vH012,UV_vH014,UV_vH015,UV_vH016,UV_vH017,UV_vH023,UV_vH024,UV_vH025,
    UV_vH027,UV_vH028,UV_vH029,UV_vH031,UV_vH035,UV_vH038,UV_vH039,

    UV_vSampNo, UV_vDesc_dc : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q48: TfrmLD4Q48;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q481;

{$R *.DFM}

procedure TfrmLD4Q48.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q48.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD4Q48.UP_VarMemSet(iLength : Integer);
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
      UV_vH001  := VarArrayCreate([0, 0], varOleStr);
      UV_vH002  := VarArrayCreate([0, 0], varOleStr);
      UV_vH003  := VarArrayCreate([0, 0], varOleStr);
      UV_vH004  := VarArrayCreate([0, 0], varOleStr);
      UV_vH005  := VarArrayCreate([0, 0], varOleStr);
      UV_vH006  := VarArrayCreate([0, 0], varOleStr);
      UV_vH007  := VarArrayCreate([0, 0], varOleStr);
      UV_vH008  := VarArrayCreate([0, 0], varOleStr);
      UV_vH009  := VarArrayCreate([0, 0], varOleStr);
      UV_vH010  := VarArrayCreate([0, 0], varOleStr);
      UV_vH011  := VarArrayCreate([0, 0], varOleStr);
      UV_vH012  := VarArrayCreate([0, 0], varOleStr);
      UV_vH014  := VarArrayCreate([0, 0], varOleStr);
      UV_vH015  := VarArrayCreate([0, 0], varOleStr);
      UV_vH016  := VarArrayCreate([0, 0], varOleStr);
      UV_vH017  := VarArrayCreate([0, 0], varOleStr);
      UV_vH023  := VarArrayCreate([0, 0], varOleStr);
      UV_vH024  := VarArrayCreate([0, 0], varOleStr);      
      UV_vH025  := VarArrayCreate([0, 0], varOleStr);
      UV_vH027  := VarArrayCreate([0, 0], varOleStr);
      UV_vH028  := VarArrayCreate([0, 0], varOleStr);
      UV_vH029  := VarArrayCreate([0, 0], varOleStr);
      UV_vH031  := VarArrayCreate([0, 0], varOleStr);
      UV_vH035  := VarArrayCreate([0, 0], varOleStr);
      UV_vH038  := VarArrayCreate([0, 0], varOleStr);
      UV_vH039  := VarArrayCreate([0, 0], varOleStr);


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

      VarArrayRedim(UV_vH001,  iLength);
      VarArrayRedim(UV_vH002,  iLength);
      VarArrayRedim(UV_vH003,  iLength);
      VarArrayRedim(UV_vH004,  iLength);
      VarArrayRedim(UV_vH005,  iLength);
      VarArrayRedim(UV_vH006,  iLength);
      VarArrayRedim(UV_vH007,  iLength);
      VarArrayRedim(UV_vH008,  iLength);
      VarArrayRedim(UV_vH009,  iLength);
      VarArrayRedim(UV_vH010,  iLength);
      VarArrayRedim(UV_vH011,  iLength);
      VarArrayRedim(UV_vH012,  iLength);
      VarArrayRedim(UV_vH014,  iLength);
      VarArrayRedim(UV_vH015,  iLength);
      VarArrayRedim(UV_vH016,  iLength);
      VarArrayRedim(UV_vH017,  iLength);
      VarArrayRedim(UV_vH023,  iLength);
      VarArrayRedim(UV_vH024,  iLength);      

      VarArrayRedim(UV_vH025,  iLength);

      VarArrayRedim(UV_vH027,  iLength);
      VarArrayRedim(UV_vH028,  iLength);
      VarArrayRedim(UV_vH029,  iLength);
      VarArrayRedim(UV_vH031,  iLength);
      VarArrayRedim(UV_vH035,  iLength);
      VarArrayRedim(UV_vH038,  iLength);
      VarArrayRedim(UV_vH039,  iLength);
      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD4Q48.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q48.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q48.grdMasterCellLoaded(Sender: TObject; DataCol,
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
           5 : Value := UV_vH001[DataRow-1];
           6 : Value := UV_vH002[DataRow-1];
           7 : Value := UV_vH003[DataRow-1];
           8 : Value := UV_vH004[DataRow-1];
           9 : Value := UV_vH005[DataRow-1];
          10 : Value := UV_vH006[DataRow-1];
          11 : Value := UV_vH007[DataRow-1];
          12 : Value := UV_vH008[DataRow-1];
          13 : Value := UV_vH009[DataRow-1];
          14 : Value := UV_vH010[DataRow-1];
          15 : Value := UV_vH011[DataRow-1];
          16 : Value := UV_vH012[DataRow-1];
          17 : Value := UV_vH014[DataRow-1];
          18 : Value := UV_vH015[DataRow-1];
          19 : Value := UV_vH016[DataRow-1];
          20 : Value := UV_vH017[DataRow-1];
          21 : Value := UV_vH023[DataRow-1];
          22 : Value := UV_vH024[DataRow-1];
          23 : Value := UV_vH025[DataRow-1];
          24 : Value := UV_vH029[DataRow-1];


          end;
          grdMaster.Col[5].heading   :='H001';
          grdMaster.Col[6].heading   :='H002';
          grdMaster.Col[7].heading   :='H003';
          grdMaster.Col[8].heading   :='H004';
          grdMaster.Col[9].heading   :='H005';
          grdMaster.Col[10].heading  :='H006';
          grdMaster.Col[11].heading  :='H007';
          grdMaster.Col[12].heading  :='H008';
          grdMaster.Col[13].heading  :='H009';
          grdMaster.Col[14].heading  :='H010';
          grdMaster.Col[15].heading  :='H011';
          grdMaster.Col[16].heading  :='H012';
          grdMaster.Col[17].heading  :='H014';
          grdMaster.Col[18].heading  :='H015';
          grdMaster.Col[19].heading  :='H016';
          grdMaster.Col[20].heading  :='H017';
          grdMaster.Col[21].heading  :='H023';
          grdMaster.Col[22].heading  :='H024';
          grdMaster.Col[23].heading  :='H025';
          grdMaster.Col[24].heading  :='H029';


    end
    else if Chk_outside.Checked =true  then
    begin
        case DataCol of
           1 : Value := UV_vNo[DataRow-1];
           2 : Value := UV_vSampNo[DataRow-1];
           3 : Value := UV_vBUN[DataRow-1];
           4 : Value := UV_vName[DataRow-1];
           5 : Value := UV_vH027[DataRow-1];
           6 : Value := UV_vH028[DataRow-1];
           7 : Value := UV_vH031[DataRow-1];
           8 : Value := UV_vH035[DataRow-1];
           9 : Value := UV_vH038[DataRow-1];
          10 : Value := UV_vH039[DataRow-1];

          end;
          for i:=5 to 24 do
          begin
               grdMaster.Col[i].heading :=' ';
          end;
          grdMaster.Col[5].heading  :='H027';
          grdMaster.Col[6].heading  :='H028';
          grdMaster.Col[7].heading  :='H031';
          grdMaster.Col[8].heading  :='H035';
          grdMaster.Col[9].heading  :='H038';
          grdMaster.Col[10].heading :='H039';


    end;




end;

procedure TfrmLD4Q48.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vH001,Sum_vH002,Sum_vH003,Sum_vH004,Sum_vH005,Sum_vH006,Sum_vH007,Sum_vH008,Sum_vH009,Sum_vH010,Sum_vH011,Sum_vH012,Sum_vH014,Sum_vH015,Sum_vH016,Sum_vH017,Sum_vH023,Sum_vH024,Sum_vH025,
    Sum_vH027,Sum_vH028,Sum_vH029,Sum_vH031,Sum_vH035,Sum_vH038,Sum_vH039 : integer;

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


   Sum_vH001  :=0;  Sum_vH002  :=0;  Sum_vH003  :=0;  Sum_vH004  :=0;  Sum_vH005  :=0;
   Sum_vH006  :=0;  Sum_vH007  :=0;  Sum_vH008  :=0;  Sum_vH009  :=0;  Sum_vH010  :=0;
   Sum_vH011  :=0;  Sum_vH012  :=0;  Sum_vH014  :=0;  Sum_vH015  :=0;  Sum_vH016  :=0;
   Sum_vH017  :=0;  Sum_vH023  :=0;  Sum_vH024  :=0;  Sum_vH025  :=0;

   Sum_vH027  :=0;  Sum_vH028  :=0;  Sum_vH029  :=0;  Sum_vH031  :=0;  Sum_vH035  :=0;
   Sum_vH038  :=0;  Sum_vH039  :=0;
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

      if Copy(cmbJisa.Text,1,2) ='15' then
      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''''
      else
      sWhere  := ' WHERE B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + '''  and K.gubn_part=''02''';


      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
         sOrderBy := ' ORDER BY B.num_bunju ';

      end
      else
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
         sOrderBy := ' ORDER BY G.num_samp ';   
      end;
      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q36', 'Q', 'N');
         while not qryBunju.Eof do                                                  
         Begin                                                                      
                                                                                    
            sHul_List := '';                                                        
                                                                                    
            Gride.Progress := Gride.Progress + 1;                                   
            Application.ProcessMessages;

            sHul_List := UF_hangmokList;

            if (Pos('H001' ,sHul_List)> 0) or (Pos('H009' ,sHul_List)> 0) or (Pos('H023' ,sHul_List)> 0) or (Pos('H039' ,sHul_List)> 0)or
               (Pos('H002' ,sHul_List)> 0) or (Pos('H010' ,sHul_List)> 0) or (Pos('H025' ,sHul_List)> 0) or (Pos('H024' ,sHul_List)> 0)or
               (Pos('H003' ,sHul_List)> 0) or (Pos('H011' ,sHul_List)> 0) or (Pos('H027' ,sHul_List)> 0) or
               (Pos('H004' ,sHul_List)> 0) or (Pos('H012' ,sHul_List)> 0) or (Pos('H028' ,sHul_List)> 0) or
               (Pos('H005' ,sHul_List)> 0) or (Pos('H014' ,sHul_List)> 0) or (Pos('H029' ,sHul_List)> 0) or
               (Pos('H006' ,sHul_List)> 0) or (Pos('H015' ,sHul_List)> 0) or (Pos('H031' ,sHul_List)> 0) or
               (Pos('H007' ,sHul_List)> 0) or (Pos('H016' ,sHul_List)> 0) or (Pos('H035' ,sHul_List)> 0) or
               (Pos('H008' ,sHul_List)> 0) or (Pos('H017' ,sHul_List)> 0) or (Pos('H038' ,sHul_List)> 0) Then
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


                         UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                         UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                         UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                         UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

                          ///////////////////////////////////////////////////////////
                          if (pos('H001', sHul_List) > 0) and (Chk_Self.Checked=true)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 1,  6)) = '') then
                                  begin
                                  UV_vH001[Index] := '결과無';
                                  Sum_vH001 := Sum_vH001 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('H002', sHul_List) > 0) and (Chk_Self.Checked=true)  then
                            begin
                                 if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                                    begin
                                    UV_vH002[Index] := '결과無';
                                    Sum_vH002 := Sum_vH002 + 1;
                                    Check := True;
                                 end;

                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('H003', sHul_List) > 0) and (Chk_Self.Checked=true)  then
                            begin
                                 if (Trim(Copy(UV_fGulkwa, 13,  6)) = '') then
                                    begin
                                    UV_vH003[Index] := '결과無';
                                    Sum_vH003 := Sum_vH003 + 1;
                                    Check := True;
                                 end;
                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('H004', sHul_List) > 0) and (Chk_Self.Checked=true) then
                         begin
                              if (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                                 begin
                                 UV_vH004[Index] := '결과無';
                                 Sum_vH004 := Sum_vH004 + 1;
                                 Check := True;
                              end;
                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('H005', sHul_List) > 0) and (Chk_Self.Checked=true) then
                         begin
                              if (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                                 begin
                                 UV_vH005[Index] := '결과無';
                                 Sum_vH005 := Sum_vH005 + 1;
                                 Check := True;
                              end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('H006', sHul_List) > 0) and (Chk_Self.Checked=true) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 31,  6)) = '') then
                                  begin
                                  UV_vH006[Index] := '결과無';
                                  Sum_vH006 := Sum_vH006 + 1;
                                  Check := True;
                                  end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('H007', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                                  begin
                                  UV_vH007[Index] := '결과無';
                                  Sum_vH007 := Sum_vH007 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('H008', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 43,  6)) = '') then
                                  begin
                                  UV_vH008[Index] := '결과無';
                                  Sum_vH008 := Sum_vH008 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('H009', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 49,  6)) = '') then
                                  begin
                                  UV_vH009[Index] := '결과無';
                                  Sum_vH009 := Sum_vH009 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('H010', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 55,  6)) = '') then
                                  begin
                                  UV_vH010[Index] := '결과無';
                                  Sum_vH010 := Sum_vH010 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('H011', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 61,  6)) = '') then
                                  begin
                                  UV_vH011[Index] := '결과無';
                                  Sum_vH011 := Sum_vH011 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('H012', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                                  begin
                                  UV_vH012[Index] := '결과無';
                                  Sum_vH012 := Sum_vH012 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('H014', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 79,  6)) = '') then
                                  begin
                                  UV_vH014[Index] := '결과無';
                                  Sum_vH014 := Sum_vH014 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H015', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 85,  6)) = '') then
                                  begin
                                  UV_vH015[Index] := '결과無';
                                  Sum_vH015 := Sum_vH015 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H016', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 91,  6)) = '') then
                                  begin
                                  UV_vH016[Index] := '결과無';
                                  Sum_vH016 := Sum_vH016 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H017', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 97,  6)) = '') then
                                  begin
                                  UV_vH017[Index] := '결과無';
                                  Sum_vH017 := Sum_vH017 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H023', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 133,  6)) = '') then
                                  begin
                                  UV_vH023[Index] := '결과無';
                                  Sum_vH023 := Sum_vH023 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H024', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 139,  6)) = '') then
                                  begin
                                  UV_vH024[Index] := '결과無';
                                  Sum_vH024 := Sum_vH024 + 1;
                                  Check := True;
                               end;
                          end;

                            ///////////////////////////////////////////////////////////
                          if (pos('H025', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 145,  6)) = '') then
                                  begin
                                  UV_vH025[Index] := '결과無';
                                  Sum_vH025 := Sum_vH025 + 1;
                                  Check := True;
                               end;
                          end;

                           ///////////////////////////////////////////////////////////외주 
                          if (pos('H027', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 151,  6)) = '') then
                                  begin
                                  UV_vH027[Index] := '결과無';
                                  Sum_vH027 := Sum_vH027 + 1;
                                  Check := True;
                               end;
                          end;

                            ///////////////////////////////////////////////////////////
                          if (pos('H028', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 157,  6)) = '') then
                                  begin
                                  UV_vH028[Index] := '결과無';
                                  Sum_vH028 := Sum_vH028 + 1;
                                  Check := True;
                               end;
                          end;
                            /////////////////////////////////////////////////////////// 자체검사로 변경
                          if (pos('H029', sHul_List) > 0) and (Chk_Self.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 163,  6)) = '') then
                                  begin
                                  UV_vH029[Index] := '결과無';
                                  Sum_vH029 := Sum_vH029 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H031', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 169,  6)) = '') then
                                  begin
                                  UV_vH029[Index] := '결과無';
                                  Sum_vH029 := Sum_vH029 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H035', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 175,  6)) = '') then
                                  begin
                                  UV_vH035[Index] := '결과無';
                                  Sum_vH035 := Sum_vH035 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H038', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 217,  6)) = '') then
                                  begin
                                  UV_vH038[Index] := '결과無';
                                  Sum_vH038 := Sum_vH038 + 1;
                                  Check := True;
                               end;
                          end;
                            ///////////////////////////////////////////////////////////
                          if (pos('H039', sHul_List) > 0) and (Chk_OutSide.Checked=true) Then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 223,  6)) = '') then
                                  begin
                                  UV_vH039[Index] := '결과無';
                                  Sum_vH039 := Sum_vH039 + 1;
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
             { if Check = True then
                 begin
                 UP_VarMemSet(Index+1);
                 inc(index);
              end  }

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
         UV_vH001[Index]:= Sum_vH001;              
         UV_vH002[Index]:= Sum_vH002;              
         UV_vH003[Index]:= Sum_vH003;              
         UV_vH004[Index]:= Sum_vH004;              
         UV_vH005[Index]:= Sum_vH005;              
         UV_vH006[Index]:= Sum_vH006;              
         UV_vH007[Index]:= Sum_vH007;              
         UV_vH008[Index]:= Sum_vH008;              
         UV_vH009[Index]:= Sum_vH009;              
         UV_vH010[Index]:= Sum_vH010;              
         UV_vH011[Index]:= Sum_vH011;              
         UV_vH012[Index]:= Sum_vH012;              
         UV_vH014[Index]:= Sum_vH014;              
         UV_vH015[Index]:= Sum_vH015;              
         UV_vH016[Index]:= Sum_vH016;              
         UV_vH017[Index]:= Sum_vH017;              
         UV_vH023[Index]:= Sum_vH023;
         UV_vH024[Index]:= Sum_vH024;                       
         UV_vH025[Index]:= Sum_vH025;
         UV_vH027[Index]:= Sum_vH027;              
         UV_vH028[Index]:= Sum_vH028;              
         UV_vH029[Index]:= Sum_vH029;              
         UV_vH031[Index]:= Sum_vH031;              
         UV_vH035[Index]:= Sum_vH035;              
         UV_vH038[Index]:= Sum_vH038;              
         UV_vH039[Index]:= Sum_vH039;

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

procedure TfrmLD4Q48.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;

   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;

procedure TfrmLD4Q48.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q481 := TfrmLD4Q481.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD4Q481.QuickRep.Preview
  else                                frmLD4Q481.QuickRep.Print;

end;

procedure TfrmLD4Q48.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;


end;

function TfrmLD4Q48.UF_hangmokList : String;
var
   i, j, k : integer;
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

function TfrmLD4Q48.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
