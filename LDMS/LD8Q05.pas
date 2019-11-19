//==============================================================================
// 프로그램 설명 : 이상내역조회
// 시스템        : 통합검진
// 작성일자      : 2007.04.18
// 작성자        : 구수정
//==============================================================================


unit LD8Q05;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Calendar;

type
  TfrmLD8Q05 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    Label1: TLabel;
    cmbJisa: TComboBox;
    mskFrom: TDateEdit;
    qry_gulkwa: TQuery;
    btnsDate: TSpeedButton;
    Label2: TLabel;
    grdMaster: TtsGrid;
    Gubn_hm1: TCheckBox;
    Gubn_hm2: TCheckBox;
    Gubn_hm3: TCheckBox;
    qryPGulkwa: TQuery;
    qryPf_hm: TQuery;
    qryCell: TQuery;
    qry_gulkwa2: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure rbDateClick(Sender: TObject);
    procedure grdMasterHeadingClick(Sender: TObject; DataCol: Integer);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure cmbJisaChange(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vDesc_jisa,  UV_vDat_gmgn,    UV_vDesc_name,   UV_vNum_jumin,    UV_vDesc_bjjs,
    UV_vDat_bunju,  UV_vNum_bunju,   UV_vDesc_dc,
    UV_vDesc_dept,  UV_vGubn_hm,     UV_vNum_glkwa  : Variant;

    UV_sBasicSQL: String;
    UV_bCkEdit, UV_bCkEditCnt : Boolean;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
    function  UF_Condition : Boolean;
  public
    // 달력관련한 함수 및 변수..
    UV_sCalDate : String;
    function  UF_SawonCk : Boolean;
    { Public declarations }
  end;
const
    NumberOfRows = 6;
    NumberOfCols = 7;
var
  frmLD8Q05: TfrmLD8Q05;
  TextArray: Array [1..NumberOfCols,1..NumberOfRows] of Variant;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD8Q05.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD8Q05.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vDesc_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn    := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_bjjs   := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dept   := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_glkwa   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vDesc_jisa,   iLength);
      VarArrayRedim(UV_vDat_gmgn,    iLength);
      VarArrayRedim(UV_vDesc_name,   iLength);
      VarArrayRedim(UV_vNum_jumin,   iLength);
      VarArrayRedim(UV_vDesc_bjjs,   iLength);
      VarArrayRedim(UV_vDat_bunju,   iLength);
      VarArrayRedim(UV_vNum_bunju,   iLength);
      VarArrayRedim(UV_vDesc_dc,     iLength);
      VarArrayRedim(UV_vDesc_dept,   iLength);
      VarArrayRedim(UV_vGubn_hm,     iLength);
      VarArrayRedim(UV_vNum_glkwa,   iLength);
   end;
end;

function TfrmLD8Q05.UF_Condition : Boolean;
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


procedure TfrmLD8Q05.FormCreate(Sender: TObject);
begin
   inherited;

   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboJisa(cmbJisa);
   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   // 환경설정
   UP_VarMemSet(0);
   UV_bCkEdit := False;

   // SQL문 저장
   UV_sBasicSQL := qry_gulkwa.SQL.Text;

   // 일자(From, To)를 오늘일자로 설정
   mskFrom.Text := GV_sToday;

end;

procedure TfrmLD8Q05.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   // 자룔를 화면에 조회

   case DataCol of
      1  : Value := UV_vDesc_jisa[DataRow-1];
      2  : Value := GF_DateFormat(UV_vDat_gmgn[DataRow-1]);
      3  : Value := UV_vDesc_name[DataRow-1];
      4  : Value := UV_vNum_jumin[DataRow-1];
      5  : Value := UV_vDesc_bjjs[DataRow-1];
      6  : Value := GF_DateFormat(UV_vDat_bunju[DataRow-1]);
      7  : Value := UV_vNum_bunju[DataRow-1];
      8  : Value := UV_vDesc_dc[DataRow-1];
      9  : Value := UV_vDesc_dept[DataRow-1];
      10 : Value := UV_vGubn_hm[DataRow-1];
      11 : Value := UV_vNum_glkwa[DataRow-1];
   end;
end;

procedure TfrmLD8Q05.btnQueryClick(Sender: TObject);
var i, index, iRet,  iTemp : Integer;
    sWhere, sGroupBy, sOrderBy, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장

   index := 0;
   with qry_gulkwa do
   begin
        Active := False;

        // SQL문 생성
        sWhere := '';
        sGroupBy := '';
        sHul_List := '';

        if Copy(cmbJisa.Text, 1, 2) <> '00' then
           sWhere := ' WHERE K.cod_bjjs = ''' + Copy(cmbJisa.Text, 1, 2) + ''''
        else
            sWhere := ' WHERE K.cod_bjjs <> ''''';

        sWhere := sWhere +  ' AND K.dat_bunju = ''' + mskFROM.Text + '''';

        sWhere := sWhere +  ' AND K.num_bunju >= ''0000'' ';
        sWhere := sWhere +  ' AND K.num_bunju <= ''9999'' ';

        sGroupBy := ' GROUP BY G.cod_jisa, G.dat_gmgn, G.num_jumin, G.desc_name,' +
                    '          K.cod_bjjs, K.dat_bunju, K.num_bunju, D.desc_dc,' +
                    '          G.desc_dept, G.num_id,' +
                    '          G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga';

        sOrderBy := ' ORDER BY K.dat_bunju, K.num_bunju' ;

        SQL.Clear;
        SQL.Add(UV_sBasicSQL + sWhere + sGroupBy + sOrderBy);

        Active := True;

        if qry_gulkwa.RecordCount > 0 then
        begin
             GP_query_log(qry_gulkwa, 'LD8Q05', 'Q', 'N');
             //UP_VarMemSet(RecordCount-1);
             for i := 0 to RecordCount - 1 do
             begin
                 { //UP_VarMemSet(index);
                  // 항목(장비프로파일)추출.. => SS 항목에 포함되는 경우도 있음.
                  sHul_List := '';
                  if qry_Gulkwa.FieldByName('cod_jangbi').AsString <> '' then
                  begin
                       qryPf_hm.Active := False;
                       qryPf_hm.ParamByName('COD_PF').AsString := qry_Gulkwa.FieldByName('cod_jangbi').AsString;
                       qryPf_hm.Active := True;
                       if qryPf_hm.RecordCount > 0 then
                       begin
                            while not qryPf_hm.Eof do
                            begin
                                 sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;
                                 qryPf_hm.Next;
                            end;
                       end;
                  end;

                  // 항목(혈액프로파일)추출..
//                sJangbi_List := ''; sHul_List := '';
                  if qry_Gulkwa.FieldByName('cod_hul').AsString <> '' then
                  begin
                       qryPf_hm.Active := False;
                       qryPf_hm.ParamByName('COD_PF').AsString := qry_Gulkwa.FieldByName('cod_hul').AsString;
                       qryPf_hm.Active := True;
                       if qryPf_hm.RecordCount > 0 then
                       begin
                            while not qryPf_hm.Eof do
                            begin
                                 sHul_List := sHul_List + qryPf_hm.FieldByName('COD_HM').AsString;

                                 qryPf_hm.Next;
                            end;
                       end;
                  end;

                  //추가검사항목..
                  iRet := GF_MulToSingle(qry_Gulkwa.FieldByName('cod_chuga').AsString, 4, vCod_chuga);
                  for iTemp := 0 to iRet-1 do
                  begin
                       sHul_List := sHul_List + vCod_chuga[iTemp];
                  end; }

                  if (Gubn_hm2.checked = true) or (Gubn_hm1.checked = true) then
                  begin
                  with qryPGulkwa do
                  begin
                       close;
                       ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                       ParamByName('num_id').AsString    := qry_Gulkwa.FieldByName('num_id').AsString;
                       ParamByName('dat_gmgn').AsString  := qry_Gulkwa.FieldByName('dat_gmgn').AsString;
                       open;

                       if RecordCount > 0 then
                       begin
                           if Gubn_hm1.checked = true then
                           begin
                                 filter := 'gubn_part = ''01''';
                                 //if (pos('C009', sHul_List) > 0) and (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                                 if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                                    and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,49,6)) <> '')
                                    and (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,49,6))) > 200) then
                                 begin
                                      UP_VarMemSet(index);

                                      UV_vDesc_jisa[index]   := qry_Gulkwa.FieldByName('cod_jisa').AsString;
                                      UV_vDat_gmgn[index]    := qry_Gulkwa.FieldByName('Dat_gmgn').AsString;
                                      UV_vDesc_name[index]   := qry_Gulkwa.FieldByName('Desc_name').AsString;
                                      UV_vNum_jumin[index]   := qry_Gulkwa.FieldByName('Num_jumin').AsString;
                                      UV_vDesc_bjjs[index]   := qry_Gulkwa.FieldByName('cod_bjjs').AsString;
                                      UV_vDat_bunju[index]   := qry_Gulkwa.FieldByName('dat_bunju').AsString;
                                      UV_vNum_bunju[index]   := qry_Gulkwa.FieldByName('num_bunju').AsString;
                                      UV_vDesc_dc[index]     := qry_Gulkwa.FieldByName('Desc_dc').AsString;
                                      UV_vDesc_dept[index]   := qry_Gulkwa.FieldByName('Desc_dept').AsString;
                                      UV_vGubn_hm[index]     := '간기능검사(AST)';
                                      UV_vNum_glkwa[index]   := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,49,6));

                                      Inc(index);
                                 end;
                                 //if (pos('C010', sHul_List) > 0) and (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                                 if (qryPGulkwa.FieldByName('gubn_part').AsString = '01')
                                    and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6)) <> '')
                                    and (StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6))) > 200) then
                                 begin
                                      UP_VarMemSet(index);

                                      UV_vDesc_jisa[index]   := qry_Gulkwa.FieldByName('cod_jisa').AsString;
                                      UV_vDat_gmgn[index]    := qry_Gulkwa.FieldByName('Dat_gmgn').AsString;
                                      UV_vDesc_name[index]   := qry_Gulkwa.FieldByName('Desc_name').AsString;
                                      UV_vNum_jumin[index]   := qry_Gulkwa.FieldByName('Num_jumin').AsString;
                                      UV_vDesc_bjjs[index]   := qry_Gulkwa.FieldByName('cod_bjjs').AsString;
                                      UV_vDat_bunju[index]   := qry_Gulkwa.FieldByName('dat_bunju').AsString;
                                      UV_vNum_bunju[index]   := qry_Gulkwa.FieldByName('num_bunju').AsString;
                                      UV_vDesc_dc[index]     := qry_Gulkwa.FieldByName('Desc_dc').AsString;
                                      UV_vDesc_dept[index]   := qry_Gulkwa.FieldByName('Desc_dept').AsString;
                                      UV_vGubn_hm[index]     := '간기능검사(ALT)';
                                      UV_vNum_glkwa[index]   := trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,55,6));

                                      Inc(index);
                                 end;
                              end;
                              if Gubn_hm2.checked = true then
                              begin
                                 filter := 'gubn_part = ''02''';
                                 //if (pos('H002', sHul_List) > 0) and (qryPGulkwa.FieldByName('gubn_part').AsString = '02')
                                 if (qryPGulkwa.FieldByName('gubn_part').AsString = '02')
                                    and (trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)) <> '')
                                    and  (Trunc(StrToFloat(trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6)))) < 8) then
                                 begin
                                      UP_VarMemSet(index);

                                      UV_vDesc_jisa[index]   := qry_Gulkwa.FieldByName('cod_jisa').AsString;
                                      UV_vDat_gmgn[index]    := qry_Gulkwa.FieldByName('Dat_gmgn').AsString;
                                      UV_vDesc_name[index]   := qry_Gulkwa.FieldByName('Desc_name').AsString;
                                      UV_vNum_jumin[index]   := qry_Gulkwa.FieldByName('Num_jumin').AsString;
                                      UV_vDesc_bjjs[index]   := qry_Gulkwa.FieldByName('cod_bjjs').AsString;
                                      UV_vDat_bunju[index]   := qry_Gulkwa.FieldByName('dat_bunju').AsString;
                                      UV_vNum_bunju[index]   := qry_Gulkwa.FieldByName('num_bunju').AsString;
                                      UV_vDesc_dc[index]     := qry_Gulkwa.FieldByName('Desc_dc').AsString;
                                      UV_vDesc_dept[index]   := qry_Gulkwa.FieldByName('Desc_dept').AsString;
                                      UV_vGubn_hm[index]     := '혈색소';
                                      UV_vNum_glkwa[index]   :=  trim(copy(qryPGulkwa.FieldByName('desc_glkwa1').AsString,7,6));

                                      Inc(index);
                                 end;
                              end;
                       end;
                  end;  //end of with
                  end;
                  if Gubn_hm3.checked = true then
                  begin
                       with qryCell do
                       begin
                            close;
                            ParamByName('cod_bjjs').AsString  := Copy(cmbJisa.Text,1,2);
                            ParamByName('num_id').AsString    := qry_Gulkwa.FieldByName('num_id').AsString;
                            ParamByName('dat_gmgn').AsString  := qry_Gulkwa.FieldByName('dat_gmgn').AsString;
                            open;
                            if RecordCount > 0 then
                            begin
                                 UP_VarMemSet(index);

                                 UV_vDesc_jisa[index]   := qry_Gulkwa.FieldByName('cod_jisa').AsString;
                                 UV_vDat_gmgn[index]    := qry_Gulkwa.FieldByName('Dat_gmgn').AsString;
                                 UV_vDesc_name[index]   := qry_Gulkwa.FieldByName('Desc_name').AsString;
                                 UV_vNum_jumin[index]   := qry_Gulkwa.FieldByName('Num_jumin').AsString;
                                 UV_vDesc_bjjs[index]   := qry_Gulkwa.FieldByName('cod_bjjs').AsString;
                                 UV_vDat_bunju[index]   := qry_Gulkwa.FieldByName('dat_bunju').AsString;
                                 UV_vNum_bunju[index]   := qry_Gulkwa.FieldByName('num_bunju').AsString;
                                 UV_vDesc_dc[index]     := qry_Gulkwa.FieldByName('Desc_dc').AsString;
                                 UV_vDesc_dept[index]   := qry_Gulkwa.FieldByName('Desc_dept').AsString;
                                 UV_vGubn_hm[index]     := '자궁경부암검사';
                                 UV_vNum_glkwa[index]   := '유형3(이형성세포)';

                                 Inc(index);
                            end;
                       end;  //end of with
                  end;

                  next;
             end;  //end of for

             // Table과 Disconnect
             Active := False;
        end
        else
        begin
             GF_MsgBox('Information', 'NODATA');
             Exit;
        end;
   end;
   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD8Q05.grdMasterRowChanged(Sender: TObject; OldRow,
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


procedure TfrmLD8Q05.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

end;

procedure TfrmLD8Q05.UP_Click(Sender: TObject);
begin
   inherited;
   //날짜 입력
    if Sender = btnsDate then
      GF_CalendarClick(mskFrom);

end;


procedure TfrmLD8Q05.rbDateClick(Sender: TObject);
begin
   inherited;

   // 모두 No Visible
   mskFrom.Visible  := False; 
   btnsDate.Visible := False;


   
end;

procedure TfrmLD8Q05.UP_QuickSort(sDivs: String; var A : Variant; l,h : Integer);
var Lo, Hi : Integer;
    Mid, T : String;
begin
   Lo := l;
   Hi := h;
   Mid := A[(Lo + Hi) div 2];

   //--------------------------------------------------------------------------
   // 내림차순
   //--------------------------------------------------------------------------
   if sDivs = 'D' then
   begin
      repeat
         while A[Lo] < Mid do Inc(Lo);
         while A[Hi] > Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vDesc_jisa[Lo];     UV_vDesc_jisa[Lo]   := UV_vDesc_jisa[Hi];     UV_vDesc_jisa[Hi]  := T;
            T := UV_vDat_gmgn[Lo];      UV_vDat_gmgn[Lo]    := UV_vDat_gmgn[Hi];      UV_vDat_gmgn[Hi]   := T;
            T := UV_vDesc_name[Lo];     UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];     UV_vDesc_name[Hi]  := T;
            T := UV_vNum_jumin[Lo];     UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];     UV_vNum_jumin[Hi]  := T;
            T := UV_vDesc_bjjs[Lo];     UV_vDesc_bjjs[Lo]   := UV_vDesc_bjjs[Hi];     UV_vDesc_bjjs[Hi]  := T;
            T := UV_vDat_bunju[Lo];     UV_vDat_bunju[Lo]   := UV_vDat_bunju[Hi];     UV_vDat_bunju[Hi]  := T;
            T := UV_vNum_bunju[Lo];     UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];     UV_vNum_bunju[Hi]  := T;
            T := UV_vDesc_dc[Lo];       UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];       UV_vDesc_dc[Hi]    := T;
            T := UV_vDesc_dept[Lo];     UV_vDesc_dept[Lo]   := UV_vDesc_dept[Hi];     UV_vDesc_dept[Hi]  := T;
            T := UV_vGubn_hm[Lo];       UV_vGubn_hm[Lo]     := UV_vGubn_hm[Hi];       UV_vGubn_hm[Hi]    := T;
            T := UV_vNum_glkwa[Lo];     UV_vNum_glkwa[Lo]   := UV_vNum_glkwa[Hi];     UV_vNum_glkwa[Hi]  := T;

            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;

      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end
   else
   //--------------------------------------------------------------------------
   // 오름차순
   //--------------------------------------------------------------------------
   begin
      repeat
         while A[Lo] > Mid do Inc(Lo);
         while A[Hi] < Mid do Dec(Hi);
         if Lo <= Hi then
         begin
            T := UV_vDesc_jisa[Lo];     UV_vDesc_jisa[Lo]   := UV_vDesc_jisa[Hi];     UV_vDesc_jisa[Hi]  := T;
            T := UV_vDat_gmgn[Lo];      UV_vDat_gmgn[Lo]    := UV_vDat_gmgn[Hi];      UV_vDat_gmgn[Hi]   := T;
            T := UV_vDesc_name[Lo];     UV_vDesc_name[Lo]   := UV_vDesc_name[Hi];     UV_vDesc_name[Hi]  := T;
            T := UV_vNum_jumin[Lo];     UV_vNum_jumin[Lo]   := UV_vNum_jumin[Hi];     UV_vNum_jumin[Hi]  := T;
            T := UV_vDesc_bjjs[Lo];     UV_vDesc_bjjs[Lo]   := UV_vDesc_bjjs[Hi];     UV_vDesc_bjjs[Hi]  := T;
            T := UV_vDat_bunju[Lo];     UV_vDat_bunju[Lo]   := UV_vDat_bunju[Hi];     UV_vDat_bunju[Hi]  := T;
            T := UV_vNum_bunju[Lo];     UV_vNum_bunju[Lo]   := UV_vNum_bunju[Hi];     UV_vNum_bunju[Hi]  := T;
            T := UV_vDesc_dc[Lo];       UV_vDesc_dc[Lo]     := UV_vDesc_dc[Hi];       UV_vDesc_dc[Hi]    := T;
            T := UV_vDesc_dept[Lo];     UV_vDesc_dept[Lo]   := UV_vDesc_dept[Hi];     UV_vDesc_dept[Hi]  := T;
            T := UV_vGubn_hm[Lo];       UV_vGubn_hm[Lo]     := UV_vGubn_hm[Hi];       UV_vGubn_hm[Hi]    := T;
            T := UV_vNum_glkwa[Lo];     UV_vNum_glkwa[Lo]   := UV_vNum_glkwa[Hi];     UV_vNum_glkwa[Hi]  := T;
            Inc(Lo);
            Dec(Hi);
         end;
      until Lo > Hi;
      if Hi > l then UP_QuickSort(sDivs, A, l, Hi);
      if Lo < h then UP_QuickSort(sDivs, A, Lo, h);
   end;
end;

procedure TfrmLD8Q05.grdMasterHeadingClick(Sender: TObject;
  DataCol: Integer);
var iCnt : Integer;
    sOrder : String;
begin
   inherited;

   // 자료가 존재하는지 Check
   if grdMaster.Rows = 0 then exit;

   iCnt := grdMaster.Rows;

   if grdMaster.Col[DataCol].SortPicture = spUp then
   begin
      grdMaster.Col[DataCol].SortPicture := spDown;
      sOrder := 'A';
   end
   else
   begin
      grdMaster.Col[DataCol].SortPicture := spUp;
      sOrder := 'D';
   end;

   case DataCol of
      1  : UP_QuickSort(sOrder, UV_vDesc_jisa,    0, iCnt-1);
      2  : UP_QuickSort(sOrder, UV_vDat_gmgn,     0, iCnt-1);
      3  : UP_QuickSort(sOrder, UV_vDesc_name,    0, iCnt-1);
      4  : UP_QuickSort(sOrder, UV_vNum_jumin,    0, iCnt-1);
      5  : UP_QuickSort(sOrder, UV_vDesc_bjjs,    0, iCnt-1);
      6  : UP_QuickSort(sOrder, UV_vDat_bunju,    0, iCnt-1);
      7  : UP_QuickSort(sOrder, UV_vNum_bunju,    0, iCnt-1);
      8  : UP_QuickSort(sOrder, UV_vDesc_dc,      0, iCnt-1);
      9  : UP_QuickSort(sOrder, UV_vDesc_dept,    0, iCnt-1);
      10 : UP_QuickSort(sOrder, UV_vGubn_hm,      0, iCnt-1);
      11 : UP_QuickSort(sOrder, UV_vNum_glkwa,    0, iCnt-1);


      else exit;
   end;

   grdMaster.Rows := 0;
   grdMaster.Rows := iCnt;
end;

function  TfrmLD8Q05.UF_SawonCk : Boolean;
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
procedure TfrmLD8Q05.cmbJisaChange(Sender: TObject);
var jisa : string;
begin
   inherited;

   
end;
end.
