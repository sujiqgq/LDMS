//==============================================================================
// 수정일자      : 20181207
// 수정자        : 박수지
// 수정내용      : 신진주 수석 특검 항목 결과미입력 확인 프로그램 요청
// 참고사항      :
//==============================================================================
unit LD7Q06;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD7Q06 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryGulkwa: TQuery;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label4: TLabel;
    Cmb_gubn: TComboBox;
    Chk_ALL: TRadioButton;
    qryJHangmokList: TQuery;
    group_radio: TGroupBox;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    Label1: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Edit_RackS: TMaskEdit;
    Edit_RackE: TMaskEdit;
    qryProfile: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryNo_hangmok: TQuery;
    Label8: TLabel;
    cmbJisa: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN,
    UV_vSE42,UV_vSE61,UV_vSE75,UV_vSE85,UV_vSE87,UV_vSE89,UV_vSE90,UV_vSE92,UV_vSE93, UV_vSE96, UV_vSE97,
    UV_vSE98,UV_vSEA1,UV_vSEA2,UV_vSEA6,
    UV_vSampNo, UV_vDesc_dc : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD7Q06: TfrmLD7Q06;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD7Q06.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD7Q06.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD7Q06.UP_VarMemSet(iLength : Integer);
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
      UV_vSE42  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE61  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE75  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE85  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE87  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE89  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE90  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE92  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE93  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE96  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE97  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE98  := VarArrayCreate([0, 0], varOleStr);
      UV_vSEA1  := VarArrayCreate([0, 0], varOleStr);
      UV_vSEA2  := VarArrayCreate([0, 0], varOleStr);
      UV_vSEA6  := VarArrayCreate([0, 0], varOleStr);
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
      VarArrayRedim(UV_vSE42,   iLength);
      VarArrayRedim(UV_vSE61,  iLength);
      VarArrayRedim(UV_vSE75,  iLength);
      VarArrayRedim(UV_vSE85,  iLength);
      VarArrayRedim(UV_vSE87,  iLength);
      VarArrayRedim(UV_vSE89,  iLength);
      VarArrayRedim(UV_vSE90,  iLength);
      VarArrayRedim(UV_vSE92,  iLength);
      VarArrayRedim(UV_vSE93,  iLength);
      VarArrayRedim(UV_vSE96,  iLength);
      VarArrayRedim(UV_vSE97,  iLength);
      VarArrayRedim(UV_vSE98,  iLength);
      VarArrayRedim(UV_vSEA1,  iLength);
      VarArrayRedim(UV_vSEA2,  iLength);
      VarArrayRedim(UV_vSEA6,  iLength);


      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD7Q06.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD7Q06.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD7Q06.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;

   for i:=5 to 36 do
   begin
        grdMaster.Col[i].heading := ' ';
        grdMaster.Col[i].Width   := 45;
   end;

   // 자룔를 화면에 조회
   if Chk_all.Checked =true  then
   begin
      begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vSE42[DataRow-1];
            6 : Value := UV_vSE61[DataRow-1];
            7 : Value := UV_vSE75[DataRow-1];
            8 : Value := UV_vSE85[DataRow-1];
            9 : Value := UV_vSE87[DataRow-1];
           10 : Value := UV_vSE89[DataRow-1];
           11 : Value := UV_vSE90[DataRow-1];
           12 : Value := UV_vSE92[DataRow-1];
           13 : Value := UV_vSE93[DataRow-1];
           14 : Value := UV_vSE96[DataRow-1];
           15 : Value := UV_vSE97[DataRow-1];
           16 : Value := UV_vSE98[DataRow-1];
           17 : Value := UV_vSEA1[DataRow-1];
           18 : Value := UV_vSEA2[DataRow-1];
           19 : Value := UV_vSEA6[DataRow-1];
         end;
         grdMaster.Col[5].heading   :='SE42';
         grdMaster.Col[6].heading   :='SE61';
         grdMaster.Col[7].heading   :='SE75';
         grdMaster.Col[8].heading   :='SE85';
         grdMaster.Col[9].heading   :='SE87';
         grdMaster.Col[10].heading  :='SE89';
         grdMaster.Col[11].heading  :='SE90';
         grdMaster.Col[12].heading  :='SE92';
         grdMaster.Col[13].heading  :='SE93';
         grdMaster.Col[14].heading  :='SE96';
         grdMaster.Col[15].heading  :='SE97';
         grdMaster.Col[16].heading  :='SE98';
         grdMaster.Col[17].heading  :='SEA1';
         grdMaster.Col[18].heading  :='SEA2';
         grdMaster.Col[19].heading  :='SEA6';
      end;

 end;
 END;
procedure TfrmLD7Q06.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vSE42,Sum_vSE61,Sum_vSE75,Sum_vSE85,Sum_vSE87,Sum_vSE89,Sum_vSE90,Sum_vSE92,Sum_vSE93,Sum_vSE96,Sum_vSE97,Sum_vSE98,Sum_vSEA1,Sum_vSEA2,Sum_vSEA6: integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
    Check, chk_nosin, chk_ss :Boolean;
    Check_01 :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   Sum_vSE42 :=0; Sum_vSE61 :=0; Sum_vSE75 :=0; Sum_vSE85 :=0;
   Sum_vSE87 :=0; Sum_vSE89 :=0; Sum_vSE90 :=0; Sum_vSE92 :=0;
   Sum_vSE93 :=0; Sum_vSE96 :=0; Sum_vSE97 :=0; Sum_vSE98 :=0;
   Sum_vSEA1 :=0; Sum_vSEA2 :=0; Sum_vSEA6 :=0;

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
                 '         FROM bunju B with(nolock) inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';
      if cmbJisa.ItemIndex <> 0 then
      begin
        sWhere  := ' WHERE B.Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + '''';
        sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
        sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_jisa = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + edtBDate.Text + ''' and (Gubn_Part  = ''09''))';
      end
      else
      begin
        sWhere  := ' WHERE B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
        sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Dat_Bunju  = ''' + edtBDate.Text + ''' and (Gubn_Part  = ''09''))';
      end;
      
      sWhere  := sWhere + ' AND T.cod_prf not like ''%TC77%'' ';
      if Cmb_gubn.Text = '분주번호' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
      end
      else if Cmb_gubn.Text = '샘플번호' then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else    // rack No.
      begin
         if Trim(Edit_RackS.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno >= ''' + Edit_RackS.Text + '''';
         if Trim(Edit_RackE.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno <= ''' + Edit_RackE.Text + '''';
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno, B.num_bunju ';



      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD7Q06', 'Q', 'N');
         while not qryBunju.Eof do
         Begin
           Check_01:=false;

            sHul_List := '';
            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;

            sHul_List := UF_hangmokList;


            if (Pos('SE42' ,sHul_List)> 0) or (Pos('SE61' ,sHul_List)> 0) or (Pos('SE75' ,sHul_List)> 0) or (Pos('SE85' ,sHul_List)> 0) or
               (Pos('SE87' ,sHul_List)> 0) or (Pos('SE89' ,sHul_List)> 0) or (Pos('SE90' ,sHul_List)> 0) or (Pos('SE92' ,sHul_List)> 0) or
               (Pos('SE93' ,sHul_List)> 0) or (Pos('SE96' ,sHul_List)> 0) or (Pos('SE97' ,sHul_List)> 0) or (Pos('SE98' ,sHul_List)> 0) or
               (Pos('SEA1' ,sHul_List)> 0) or (Pos('SEA2' ,sHul_List)> 0) or (Pos('SEA6' ,sHul_List)> 0) Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               Check := False;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin
                         Check_01:=false;
                         if  qryGulkwa.FieldByName('gubn_part').AsString  ='09'  then
                         begin
                             UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                             UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                             UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                             GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                             Check_01:=true;
                         end;
                         if (pos('SE42', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                               begin
                                  UV_vse42[Index] := '결과無';
                                  Sum_vSE42 := Sum_vSE42 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE61', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 121,  6)) = '') then
                               begin
                                  UV_vSE61[Index] := '결과無';
                                  Sum_vSE61 := Sum_vSE61 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE75', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 199,  6)) = '') then
                               begin
                                  UV_vSE75[Index] := '결과無';
                                  Sum_vSE75 := Sum_vSE75 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE85', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 319,  6)) = '') then
                               begin
                                  UV_vSE85[Index] := '결과無';
                                  Sum_vSE85 := Sum_vSE85 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE87', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 331,  6)) = '') then
                               begin
                                  UV_vSE87[Index] := '결과無';
                                  Sum_vSE87 := Sum_vSE87 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE89', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 343,  6)) = '') then
                               begin
                                  UV_vSE89[Index] := '결과無';
                                  Sum_vSE89 := Sum_vSE89 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE90', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 349,  6)) = '') then
                               begin
                                  UV_vSE90[Index] := '결과無';
                                  Sum_vSE90 := Sum_vSE90 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE92', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 361,  6)) = '') then
                               begin
                                  UV_vSE92[Index] := '결과無';
                                  Sum_vSE92 := Sum_vSE92 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE93', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 367,  6)) = '') then
                               begin
                                  UV_vSE93[Index] := '결과無';
                                  Sum_vSE93 := Sum_vSE93 + 1;
                                  Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('SE96', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 223,  6)) = '') then
                               begin
                                  UV_vSE96[Index] := '결과無';
                                  Sum_vSE96 := Sum_vSE96 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                         if (pos('SE97', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 229,  6)) = '') then
                               begin
                                  UV_vSE97[Index] := '결과無';
                                  Sum_vSE97 := Sum_vSE97 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                         if (pos('SE98', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 385,  6)) = '') then
                               begin
                                  UV_vSE98[Index] := '결과無';
                                  Sum_vSE98 := Sum_vSE98 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                         if (pos('SEA1', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 397,  6)) = '') then
                               begin
                                  UV_vSEA1[Index] := '결과無';
                                  Sum_vSEA1 := Sum_vSEA1 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                         if (pos('SEA2', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 235,  6)) = '') then
                               begin
                                  UV_vSEA2[Index] := '결과無';
                                  Sum_vSEA2 := Sum_vSEA2 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                         if (pos('SEA6', sHul_List) > 0) and (Chk_all.Checked=true)  then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 415,  6)) = '') then
                               begin
                                  UV_vSEA6[Index] := '결과無';
                                  Sum_vSEA6 := Sum_vSEA6 + 1;
                                  Check := True;
                               end;
                          end;

                   qryGulkwa.Next;
                   end;
                   if Check = True then
                   begin
                        UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                        UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                        UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                        UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
                        UP_VarMemSet(Index+1);
                        inc(index);
                   end;
               End;
               qryGulkwa.Close;
              end;
             // end;
           qryBunju.Next;
         End;


         // Table과 Disconnect
         qryBunju.Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '총  합  계';
         UV_vSE42[Index]:= Sum_vSE42;
         UV_vSE61[Index]:= Sum_vSE61;
         UV_vSE75[Index]:= Sum_vSE75;
         UV_vSE85[Index]:= Sum_vSE85;
         UV_vSE87[Index]:= Sum_vSE87;
         UV_vSE89[Index]:= Sum_vSE89;
         UV_vSE90[Index]:= Sum_vSE90;
         UV_vSE92[Index]:= Sum_vSE92;
         UV_vSE93[Index]:= Sum_vSE93;
         UV_vSE96[Index]:= Sum_vSE96;
         UV_vSE97[Index]:= Sum_vSE97;
         UV_vSE98[Index]:= Sum_vSE98;
         UV_vSEA1[Index]:= Sum_vSEA1;
         UV_vSEA2[Index]:= Sum_vSEA2;
         UV_vSEA6[Index]:= Sum_vSEA6;
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

procedure TfrmLD7Q06.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;
procedure TfrmLD7Q06.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;


end;

function TfrmLD7Q06.UF_hangmokList : String;
var
   i ,k, j : integer;
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

   for k:=0 to  Round(Length(sTemp)/4)-1 do
   begin
   result:=result+Copy(sTemp, (k+1)*4-3, 4)+'.';
   end;

end;

function TfrmLD7Q06.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD7Q06.SBut_ExcelClick(Sender: TObject);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
