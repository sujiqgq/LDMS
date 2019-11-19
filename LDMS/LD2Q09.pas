//==============================================================================
// 프로그램 설명 : 외주항목분주현황
// 시스템        : 통합검진
// 수정일자      : 2011.06.01
// 수정자        : 송재호
// 수정내용      : 녹십자용 새로 개발
// 참고사항      :
//==============================================================================
// 수정일자      : 2012.07.20
// 수정자        : 김희석
// 수정내용      : U019 추가 
//==============================================================================

unit LD2Q09;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt;

type
  TfrmLD2Q09 = class(TfrmSingle)
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
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    RadioGroup1: TRadioGroup;
    qryProfileG: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    qryHangmok: TQuery;
    GroupBox1: TGroupBox;
    Chk_t043: TCheckBox;
    Chk_SE01: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);

    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;


    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
  private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vDat_bunju,UV_vNum_samp,UV_vName,UV_vJumin,UV_vCod_hm,UV_vDesc_hm,UV_vNum_bunju,UV_vTemp: Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
    UV_Bjjs : String;
  end;

var
  frmLD2Q09: TfrmLD2Q09;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

procedure TfrmLD2Q09.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q09.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD2Q09.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vNo        := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_bunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_samp     := VarArrayCreate([0, 0], varOleStr);
      UV_vName         := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_bunju    := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin        := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm       := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_hm      := VarArrayCreate([0, 0], varOleStr);
      UV_vTemp         := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo       ,iLength);
      VarArrayRedim(UV_vDat_bunju,iLength);
      VarArrayRedim(UV_vNum_bunju ,iLength);
      VarArrayRedim(UV_vName     ,iLength);
      VarArrayRedim(UV_vNum_samp ,iLength);
      VarArrayRedim(UV_vJumin    ,iLength);
      VarArrayRedim(UV_vCod_hm   ,iLength);
      VarArrayRedim(UV_vDesc_hm  ,iLength);
      VarArrayRedim(UV_vTemp  ,iLength);


   end;
end;

function TfrmLD2Q09.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q09.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   RadioGroup1.ItemIndex:=2;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD2Q09.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value :=UV_vNo[DataRow-1];
      2 : Value :=UV_vDat_bunju[DataRow-1];
      3 : Value :=UV_vNum_bunju[DataRow-1];
      4 : Value :=UV_vNum_samp[DataRow-1];
      5 : Value :=UV_vName[DataRow-1];
      6 : Value :=UV_vJumin    [DataRow-1];
      7 : Value :=UV_vCod_hm   [DataRow-1];
      8 : Value :=UV_vDesc_hm  [DataRow-1];
      9 : Value :=UV_vTemp  [DataRow-1];
   end;
end;

procedure TfrmLD2Q09.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp : Integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;



   
   UV_Bjjs := copy(cmbJisa.Text,1,2);

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                 ' G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, K.num_Bunju, K.dat_bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                 ' FROM bunju K inner join Gumgin G ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '         left outer join Tkgum T ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND K.Dat_Bunju  = ''' + edtBDate.Text + '''' ;

      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And K.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

      sGroupBy := ' group by G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                  ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, K.dat_bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga, G.num_samp ';
      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY K.Num_Bunju';
         1 : sOrderBy := ' ORDER BY G.cod_jisa, CAST(G.num_samp AS INT) ';
         2 : sOrderBy := ' ORDER BY CAST(G.num_samp AS INT),G.cod_jisa,K.num_Bunju,G.desc_name  ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sGroupBy + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q09', 'Q', 'N');
         For I := 1 to RecordCount do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;  //접수항목 비교

            if  ((pos('T043',sHul_List) > 0) and (Chk_t043.checked=true)) or
                ((pos('SE01',sHul_List) > 0) and (Chk_SE01.checked=true))  then
            begin
               if (pos('T043',sHul_List) > 0) and (Chk_t043.checked=true) then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vNum_bunju[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vNum_samp[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'T043';
                    UV_vDesc_hm[Index]    := '25-OH Vitamin D';

                    Inc(Index);
               end;

               if (pos('SE01',sHul_List) > 0) and (Chk_SE01.checked=true) then
               begin
                    UP_VarMemSet(Index+1);
                    UV_vNo[Index]  := Index+1;
                    UV_vDat_bunju[Index]  := qryBunju.FieldByName('Dat_bunju').AsString;
                    UV_vNum_bunju[Index]  := qryBunju.FieldByName('Num_Bunju').AsString;
                    UV_vNum_samp[Index]   := qryBunju.FieldByName('num_samp').AsString;
                    UV_vName[Index]       := qryBunju.FieldByName('desc_name').AsString;
                    UV_vJumin[Index]      := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) + '******';
                    UV_vCod_hm[Index]     := 'SE01';
                    UV_vDesc_hm[Index]    := 'HAV IgM';

                    Inc(Index);
               end;
               

            End;
            Next;
         End;

         // Table과 Disconnect
         Close;

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

procedure TfrmLD2Q09.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD2Q09.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end
   else
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;



procedure TfrmLD2Q09.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

function TfrmLD2Q09.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i ,k: integer;
begin
   result := ''; sTemp := '';

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

   for k:=0 to  Round(Length(sTemp)/4)-1 do
   begin
   result:=result+Copy(sTemp, (k+1)*4-3, 4)+'.';
   end;


end;

function TfrmLD2Q09.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
