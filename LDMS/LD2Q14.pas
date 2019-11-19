//==============================================================================
// 수정일자      : 2016.03.29
// 수정자        : 박수지
// 수정내용      : 외주항목 분주현황(씨젠)
// 참고사항      : 한의 재단진단검사의학팀1600362
//==============================================================================
unit LD2Q14;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD2Q14 = class(TfrmSingle)
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
    RBtn_print: TRadioButton;
    Gride: TGauge;
    qryPf_hm: TQuery;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    edtDc: TEdit;
    Label1: TLabel;
    QRCompositeReport: TQRCompositeReport;
    RadioGroup1: TRadioGroup;
    Btn_NamePrint: TBitBtn;
    RBtn_1: TRadioButton;
    RBtn_2: TRadioButton;
    qryProfile: TQuery;
    qryTkgum: TQuery;
    qryNo_hangmok: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Change(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    procedure QRCompositeReportAddReports(Sender: TObject);
    procedure Btn_NamePrintClick(Sender: TObject);
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure SBut_ExcelClick(Sender: TObject);
    private
    { Private declarations }
    // Grid와 연관된 Variant 변수 선언
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vDesc_dc, UV_vSampNo, UV_vBUN,
    UV_vP007, UV_vP008, UV_vU059, UV_vU063, UV_vU064, UV_vU065 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD2Q14: TfrmLD2Q14;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q141;

{$R *.DFM}

procedure TfrmLD2Q14.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

   if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD2Q14.MouseWheelHandler(var Message: TMessage);  //수정
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

procedure TfrmLD2Q14.UP_VarMemSet(iLength : Integer);
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

      UV_vP007  := VarArrayCreate([0, 0], varOleStr);
      UV_vP008  := VarArrayCreate([0, 0], varOleStr);
      UV_vU059  := VarArrayCreate([0, 0], varOleStr);
      UV_vU063  := VarArrayCreate([0, 0], varOleStr);
      UV_vU064  := VarArrayCreate([0, 0], varOleStr);
      UV_vU065  := VarArrayCreate([0, 0], varOleStr);

      UV_vDesc_dc := VarArrayCreate([0, 0], varOleStr);
      UV_vSampNo  := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vP007,  iLength);
      VarArrayRedim(UV_vP008,  iLength);
      VarArrayRedim(UV_vU059,  iLength);
      VarArrayRedim(UV_vU063,  iLength);
      VarArrayRedim(UV_vU064,  iLength);
      VarArrayRedim(UV_vU065,  iLength);
      VarArrayRedim(UV_vDesc_dc, iLength);
      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD2Q14.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD2Q14.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // 지사가 전체일 경우 지사를 Combo에 보여줌
   GP_ComboJisa(cmbjisa);

   RadioGroup1.ItemIndex:=2;
   // Default로 본사(15)를 선택
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD2Q14.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vNo[DataRow-1];
      2 : Value := UV_vJisa[DataRow-1];
      3 : Value := UV_vBUN[DataRow-1];
      4 : Value := UV_vName[DataRow-1];
      5 : Value := UV_vJumin[DataRow-1];
      6 : Value := UV_vP007[DataRow-1];
      7 : Value := UV_vP008[DataRow-1];
      8 : Value := UV_vU059[DataRow-1];
      9 : Value := UV_vU063[DataRow-1];
     10 : Value := UV_vU064[DataRow-1];
     11 : Value := UV_vU065[DataRow-1];
     12: Value := UV_vDesc_dc[DataRow-1];
     13: Value := UV_vSampNo[DataRow-1];
   end;
end;

procedure TfrmLD2Q14.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,

    Sum_vP007, Sum_vP008, Sum_vU059, Sum_vU063, Sum_vU064, Sum_vU065 : Integer;

    sSelect, sWhere, sGroupBy, sOrderBy, sdate : String;
    sJangbi_List, sHul_List : String;
    vCod_chuga : Variant;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   Sum_vP007 := 0;  Sum_vP008 := 0;  Sum_vU059 := 0;
   Sum_vU063 := 0;  Sum_vU064 := 0;  Sum_vU065 := 0;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL문 생성
      Close;
      sSelect := ' select desc_name, num_jumin, dat_gmgn, cod_dc, desc_dc, cod_jisa, gubn_nosin, gubn_adult, gubn_life, Gubn_bogun, Gubn_agab, Gubn_tkgm,                ' +
                 ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, num_Bunju, cod_jangbi, cod_hul, cod_cancer, cod_chuga,  cod_prf, Tcod_chuga, num_samp  ' +
                 ' from                                                                                                                                          ' +
                 ' (                                                                                                                                             ' +
                 ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                 ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, K.num_Bunju, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,  T.cod_prf, T.cod_chuga As Tcod_chuga, G.num_samp ' +
                 ' FROM bunju K with(nolock) left outer join Gumgin G with(nolock) ON K.cod_jisa = G.cod_jisa and K.num_id = G.num_id and K.dat_gmgn = G.dat_gmgn ' +
                 '                           left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc ' +
                 '                           left outer join Tkgum T with(nolock) ON K.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and K.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';

      sWhere  := ' WHERE K.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND (K.Dat_Bunju  = ''' + edtBDate.Text + ''' )' ;//' AND (K.Dat_Bunju  = ''' + edtBDate.Text + ''' )' ;

      If BunjuNo1.Text <> '' Then
         sWhere := sWhere + ' And K.Num_Bunju >= ''' + BunjuNo1.Text + '''' +
                            ' And K.Num_Bunju <= ''' + BunjuNo2.Text + '''';
      if Trim(edtDc.Text) <> '' then
         sWhere := sWhere + '  AND G.cod_dc = ''' + edtDc.Text + '''';

      sWhere := sWhere +
                ' union                                                                                                                                                ' +
                '                                                                                                                                                      ' +
                ' SELECT G.desc_name, G.num_jumin, G.dat_gmgn, G.cod_dc, D.desc_dc, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm, ' +
                ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, '''', G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga, '''', '''', G.num_samp ' +
                ' FROM Gumgin G with(nolock) left outer join Danche D with(nolock) ON G.cod_dc = D.cod_dc ';

      sdate := GF_Date_decrease(StrToInt(copy(edtBDate.Text,1,4)), StrToInt(copy(edtBDate.Text,5,2)), StrToInt(copy(edtBDate.Text,7,2)),0);

      sWhere := sWhere + '  Where G.dat_gmgn = ''' + sdate + ''' ';
      sWhere := sWhere + '  and (G.gubn_hulgum = ''1'' or G.gubn_hulgum = ''3'')' ;
      sWhere := sWhere + '  and G.cod_chuga = ''DRUC'' ) A ' ;

      sWhere := sWhere + ' group by desc_name, num_jumin, dat_gmgn, cod_dc, desc_dc, cod_jisa, gubn_nosin, gubn_adult, gubn_life, Gubn_bogun, Gubn_agab, Gubn_tkgm, ' +
                         ' Gubn_nsyh, Gubn_adyh, Gubn_lfyh, Gubn_bgyh, Gubn_agyh, num_Bunju, cod_jangbi, cod_hul, cod_cancer, cod_chuga,  cod_prf, Tcod_chuga, num_samp  ';

      case RadioGroup1.ItemIndex of
         0 : sOrderBy := ' ORDER BY Num_Bunju';
         1 : sOrderBy := ' ORDER BY cod_jisa, CAST(num_samp AS INT) ';
         2 : sOrderBy := ' ORDER BY CAST(num_samp AS INT), cod_jisa, num_Bunju, desc_name  ';
      end;

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Open;
      Gride.MaxValue := RecordCount;

      if qryBunju.RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD2Q04', 'Q', 'N');
         For I := 1 to RecordCount do
//         while not qryBunju.Eof do
         Begin
            Gride.Progress := I;
            sHul_List := UF_hangmokList;
            if (pos('P007',sHul_List) > 0) or (pos('P008',sHul_List) > 0) or (pos('U059',sHul_List) > 0) or
               (pos('U063',sHul_List) > 0) or (pos('U064',sHul_List) > 0) or (pos('U065',sHul_List) > 0) Then
            begin
               UP_VarMemSet(Index+1);
               UV_vNo[Index]  := IntToStr(Index+1) + '/' + qryBunju.FieldByName('dat_gmgn').AsString;
               UV_VBUN[Index] := qryBunju.FieldByName('Num_Bunju').AsString;
               UV_vName[Index] := qryBunju.FieldByName('desc_name').AsString;
               UV_vJumin[Index] := copy(qryBunju.FieldByName('num_jumin').AsString,1,7) ;
               UV_vJisa[Index] := qryBunju.FieldByName('cod_jisa').AsString;
               UV_vDesc_dc[Index] := qryBunju.FieldByName('desc_dc').AsString;
               UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;

               if pos('P007', sHul_List) > 0 then
               begin
                  UV_vP007[Index] := '';
                  Sum_vP007 := Sum_vP007 + 1;
               end
               else                               UV_vP007[Index] := '*';          //Sum_vP007
               if pos('P008', sHul_List) > 0 then
               begin
                  UV_vP008[Index] := '';
                  Sum_vP008 := Sum_vP008 + 1;
               end
               else                               UV_vP008[Index] := '*';          //Sum_vP008
               if pos('U059', sHul_List) > 0 then
               begin
                  UV_vU059[Index] := '';
                  Sum_vU059 := Sum_vU059 + 1;
               end
               else                               UV_vU059[Index] := '*';          //Sum_vU059
               if pos('U063', sHul_List) > 0 then
               begin
                  UV_vU063[Index] := '';
                  Sum_vU063 := Sum_vU063 + 1;
               end
               else                               UV_vU063[Index] := '*';          //Sum_vU063
               if pos('U064', sHul_List) > 0 then
               begin
                  UV_vU064[Index] := '';
                  Sum_vU064 := Sum_vU064 + 1;
               end
               else                               UV_vU064[Index] := '*';          //Sum_vU064
               if pos('U065', sHul_List) > 0 then
               begin
                  UV_vU065[Index] := '';
                  Sum_vU065 := Sum_vU065 + 1;
               end
               else                               UV_vU065[Index] := '*';          //Sum_vU065
               Inc(Index);

            End;
            Next;
         End;

         // Table과 Disconnect
         Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '총  합  계';

         UV_vP007[Index] := Sum_vP007; UV_vP008[Index] := Sum_vP008; UV_vU059[Index] := Sum_vU059;
         UV_vU063[Index] := Sum_vU063; UV_vU064[Index] := Sum_vU064; UV_vU065[Index] := Sum_vU065;
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

procedure TfrmLD2Q14.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtDc then
   begin
      if GF_DancheEdit(edtDc, sName) then
         edtDesc_dc.Text := sName;
   end;

end;

procedure TfrmLD2Q14.UP_Click(Sender: TObject);
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

procedure TfrmLD2Q14.btnPrintClick(Sender: TObject);
begin
  inherited;

  frmLD2Q141 := TfrmLD2Q141.Create(Self);


  frmLD2Q141.QuickRep.Print;


end;

procedure TfrmLD2Q14.FormActivate(Sender: TObject);
begin
   inherited;
   BunjuNo1.Text := '0001';
   BunjuNo2.Text := '9999';
   edtBdate.SetFocus;
end;

procedure TfrmLD2Q14.QRCompositeReportAddReports(Sender: TObject);
begin
  inherited;
  with QRCompositeReport do
  begin
  end;
end;

procedure TfrmLD2Q14.Btn_NamePrintClick(Sender: TObject);
begin
   inherited;
  { if RBtn_1.Checked then
   begin
      frmLD2Q044 := TfrmLD2Q044.Create(Self);
      if RBtn_preveiw.Checked then frmLD2Q044.QuickRep.Preview
      else                         frmLD2Q044.QuickRep.Print;
   end
   else
   begin
      frmLD2Q045 := TfrmLD2Q045.Create(Self);
  //    frmLD2Q045.QuickRep.Preview;

      if RBtn_preveiw.Checked then frmLD2Q045.QuickRep.Preview
      else                         frmLD2Q045.QuickRep.Print;
   end; }
end;

function TfrmLD2Q14.UF_hangmokList : String;
var
   sTemp, sHmCode :string;
   vCod_chuga : Variant;
   iRet, i, J, K : integer;
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

function TfrmLD2Q14.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

procedure TfrmLD2Q14.SBut_ExcelClick(Sender: TObject);
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
