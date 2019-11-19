//==============================================================================
// 프로그램 설명 : 객담검사 리스트
// 시스템        : 분석관리
// 수정일자      : 2011.03.
// 수정자        : 김승철
// 수정내용      :                      
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.14
// 수정자        : 곽윤설
// 수정내용      : 접수No시작값
// 참고사항      : [본원 - 김원연]
//==============================================================================
// 수정일자      : 2014.06.12
// 수정자        : 곽윤설
// 수정내용      : B009코드 추가 & select문 수정
// 참고사항      : [본원 - 김원연]
//==============================================================================
// 수정일자      : 2014.07.07
// 수정자        : 곽윤설
// 수정내용      : 검진일자 조회 추가
// 참고사항      : [본원 - 김원연]
//==============================================================================
unit LD4Q33;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges ;

type
  TfrmLD4Q33 = class(TfrmSingle)
    pnlCond: TPanel;
    pnlJisa: TPanel;
    cmbJisa: TComboBox;
    mskSDate: TDateEdit;
    grdMaster: TtsGrid;
    btnSDate: TSpeedButton;
    qryCell: TQuery;
    btnSize: TSpeedButton;
    Label2: TLabel;
    CmbOrder: TComboBox;
    qryCa_desc_buwi_Max: TQuery;
    qryU_Cell: TQuery;
    Label17: TLabel;
    mskEDate: TDateEdit;
    btnEDate: TSpeedButton;
    qryPf_hm: TQuery;
    qryPf_hm2: TQuery;
    Gauge1: TGauge;
    Label4: TLabel;
    Rd_No: TRadioButton;
    Edts_No: TEdit;
    Label5: TLabel;
    Edte_No: TEdit;
    Rd_Date: TRadioButton;
    Label3: TLabel;
    Cmb_Gubun: TComboBox;
    qryCell_Old: TQuery;
    CK_Max: TCheckBox;
    Edt_Max: TEdit;
    Rd_Gmgn: TRadioButton;
    mskS_gmgn: TDateEdit;
    btnS_gmgn: TSpeedButton;
    Label6: TLabel;
    mskE_gmgn: TDateEdit;
    btnE_gmgn: TSpeedButton;
    chk_bunju: TCheckBox;

    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure grdMasterRowChanged(Sender: TObject; OldRow,
      NewRow: Integer);
    procedure grdMasterCellEdit(Sender: TObject; DataCol, DataRow: Integer;
      ByUser: Boolean);
    procedure UP_Click(Sender: TObject);
    procedure up_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure btnPrintClick(Sender: TObject);
    procedure cmbRelationChange(Sender: TObject);
    function  UF_SawonCk : Boolean;
    procedure Cmb_GubunChange(Sender: TObject);
    procedure CK_MaxClick(Sender: TObject);

  private
    { Private declarations }
    // Grid와 연관된 변수들
    UV_vCount, UV_vGubn_check, UV_vDat_gmgn, UV_vDesc_name, UV_vNum_jumin, UV_vAge, UV_vSex,
    UV_vCod_jisa, UV_vNum_id, UV_vCod_cell, UV_vCod_hm, UV_vDesc_buwi, UV_vDesc_sokun, UV_vDoctor, UV_vDesc_dc, UV_vGUBN_P012 : Variant;

    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_DSawonCk : Boolean;

  public
    { Public declarations }
  end;

var
  frmLD4Q33: TfrmLD4Q33;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q331;

{$R *.DFM}

procedure TfrmLD4Q33.MouseWheelHandler(var Message: TMessage);
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//추가
         grdMaster.Refresh; //추가
      end;
   end;
end;

procedure TfrmLD4Q33.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vCount      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn_check := VarArrayCreate([0, 0], varOleStr);
      UV_vDat_gmgn   := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_name  := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_jumin  := VarArrayCreate([0, 0], varOleStr);
      UV_vAge        := VarArrayCreate([0, 0], varOleStr);
      UV_vSex        := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_jisa   := VarArrayCreate([0, 0], varOleStr);
      UV_vNum_id     := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_cell   := VarArrayCreate([0, 0], varOleStr);
      UV_vCod_hm     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_buwi  := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_sokun := VarArrayCreate([0, 0], varOleStr);
      UV_vDoctor     := VarArrayCreate([0, 0], varOleStr);
      UV_vDesc_dc    := VarArrayCreate([0, 0], varOleStr);
      UV_vGUBN_P012  := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vCount,      iLength);
      VarArrayRedim(UV_vGubn_check, iLength);
      VarArrayRedim(UV_vDat_gmgn,   iLength);
      VarArrayRedim(UV_vDesc_name,  iLength);
      VarArrayRedim(UV_vNum_jumin,  iLength);
      VarArrayRedim(UV_vAge,        iLength);
      VarArrayRedim(UV_vSex,        iLength);
      VarArrayRedim(UV_vCod_jisa,   iLength);
      VarArrayRedim(UV_vNum_id,     iLength);
      VarArrayRedim(UV_vCod_cell,   iLength);
      VarArrayRedim(UV_vCod_hm,     iLength);
      VarArrayRedim(UV_vDesc_buwi,  iLength);
      VarArrayRedim(UV_vDesc_sokun, iLength);
      VarArrayRedim(UV_vDoctor,   iLength);
      VarArrayRedim(UV_vDesc_dc,    iLength);
      VarArrayRedim(UV_vGUBN_P012,  iLength);
   end;
end;


function TfrmLD4Q33.UF_Condition : Boolean;
begin
   Result := True;

   // 필수입력 조회조건 Check
   if Rd_Date.Checked then
   begin
      if (mskSDate.Text = '') or (mskEDate.Text = '') then
      begin
         GF_MsgBox('Warning', 'CONDITION');
         Result := False;
      end;
   end
   else if Rd_Gmgn.Checked then
   begin
      if (mskS_gmgn.Text = '') or (mskE_gmgn.Text = '') then
      begin
         GF_MsgBox('Warning', 'CONDITION');
         Result := False;
      end;
   end;
end;

function  TfrmLD4Q33.UF_SawonCk : Boolean;
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
            if (FieldByName('gubn_dept').AsString = '12') or
               (FieldByName('gubn_dept').AsString = '04') then
               Result := True
            else
               Result := False;
         end;
      end;
      // Table Disconnect
      Active := False;
   end;
end;

procedure TfrmLD4Q33.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사권한관리
//   GF_JisaSelect(pnlJisa, cmbJisa);
   GP_ComboJisa(cmbJisa);
   if UF_SawonCk then
      pnlJisa.Visible := True
   else
      pnlJisa.Visible := False;

   GP_ComboDisplay(cmbJisa, copy(GV_sUserId,1,2), 2);

   if copy(GV_sUserId,1,2) = '15' then //20140514 곽윤설~
   begin
      cmbRelation.Visible := True;
      labRelation.Visible := True;
      CK_Max.Visible      := True;
      Edt_Max.Visible     := True;
   end
   else
   begin
      cmbRelation.Visible := False;
      labRelation.Visible := False;
      CK_Max.Visible      := False;
      Edt_Max.Visible     := False;
   end;
   Edt_Max.Color  := $00E0E0E0;     //~20140514 곽윤설

   // 환경설정
   UP_VarMemSet(0);

   // 일자(From, To)를 오늘일자로 설정
   mskSDate.Text    := GV_sToday;
   mskEDate.Text    := GV_sToday;
   mskS_gmgn.Text    := GV_sToday;
   mskE_gmgn.Text    := GV_sToday;
   CmbOrder.ItemIndex := 0;

   Cmb_Gubun.ItemIndex := 0;

   Edts_No.Text := 'S' + Copy(GV_sToday,3,2) + '-';
   Edte_No.Text := 'S' + Copy(GV_sToday,3,2) + '-';

   if UF_SawonCk then
   begin
      labRelation.Visible := True;
      cmbRelation.Visible := True;
   end
   else
   begin
      labRelation.Visible := False;
      cmbRelation.Visible := False;   
   end;
end;


procedure TfrmLD4Q33.btnQueryClick(Sender: TObject);
var i, index, iAge, iRet, num : Integer;
    sSelect, sWhere, sOrderBy, sMan, sCod_hm, sSelect2, sWhere2 : String;
    vCod_profile : Variant;
    check_tk : boolean;
begin
   inherited;

   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   index := 0;
   sWhere  := ''; sSelect := ''; sOrderBy := '';
   with qryCell do
   begin
      Active := False;

      // SQL문 생성
      sSelect := ' select G.cod_jisa, G.dat_gmgn, G.num_id, G.num_jumin, G.desc_name, G.num_samp, C.desc_buwi, G.gubn_jinch,';
      sSelect := sSelect + ' G.cod_jangbi, G.cod_cancer, G.cod_hul, G.cod_chuga, T.cod_prf, T.cod_chuga as TkGum_chuga, D.desc_dc, S.desc_name as doctor';
      sSelect := sSelect + ' from gumgin G with(nolock) Left outer Join Cell C with(nolock)';
      if Cmb_Gubun.ItemIndex = 0 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''P012'' '
      else if Cmb_Gubun.ItemIndex = 1 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''P013'' '
     { else if Cmb_Gubun.ItemIndex = 2 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''B005'' ' }
      else if Cmb_Gubun.ItemIndex = 2 then
         sSelect := sSelect + ' ON G.cod_jisa = C.cod_jisa and G.num_id = C.num_id and G.dat_gmgn = C.dat_gmgn and C.cod_hm = ''B009'' '
      else
      begin
         showmessage('검사구분이 선택되어있지 않습니다.');
         Exit;
      end;
      sSelect := sSelect + ' Left outer Join ca_request J with(nolock) On G.cod_jisa = J.cod_jisa and G.num_id = J.num_id and G.dat_gmgn = J.dat_gmgn ';
      sSelect := sSelect + ' Left outer join tkgum T with(nolock) On G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn ';
      sSelect := sSelect + ' Left outer join sawon S with(nolock) On J.cod_doctor = S.cod_sawon ';
      sSelect := sSelect + ' Inner join Danche D with(nolock) On G.cod_dc = D.cod_dc ';
      //분주전 조회 시, 시간 오래걸림.. 체크박스 생성
      if chk_bunju.Checked then
      begin
         if pnlJisa.Visible then
         begin
            if Copy(cmbJisa.Text, 1, 2) <> '00' then
                 sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And '
            else
                 sWhere := ' WHERE ';
         end
         else
              sWhere := ' WHERE G.cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And ';
      end
      else
      begin
           if pnlJisa.Visible then
           begin
              if Copy(cmbJisa.Text, 1, 2) <> '00' then
                 sWhere := ' WHERE G.cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And C.cod_hm = ''' + Copy(Cmb_Gubun.text,3,4) + ''' and '
              else
                 sWhere := ' WHERE  C.cod_hm = ''' + Copy(Cmb_Gubun.text,3,4) + ''' and  ';
           end
           else
              sWhere := ' WHERE G.cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And C.cod_hm = ''' + Copy(Cmb_Gubun.text,3,4) + ''' and';
      end;
           
      if Rd_Date.Checked then
      begin
         sWhere := sWhere + ' G.dat_insert >= ''' + mskSDate.Text + '''';
         sWhere := sWhere + ' AND G.dat_insert <= ''' + mskEDate.Text + '''';
      end
      else if Rd_gmgn.Checked then  //20140707 곽윤설
      begin
         sWhere := sWhere + ' G.dat_gmgn >= ''' + mskS_gmgn.Text + '''';
         sWhere := sWhere + ' AND G.dat_gmgn <= ''' + mskE_gmgn.Text + '''';
      end
      else if Rd_No.Checked then
      begin
         if (EDTS_NO.Text <> '') and (EDTE_NO.Text <> '') then
            sWhere := sWhere + '  C.desc_buwi >= ''' + Edts_No.Text + ''''
                             + '  AND C.desc_buwi <= ''' + Edte_No.Text + ''''
         else
         begin
            showmessage('접수 No 를 입력해 주세요.');
            Edts_No.SetFocus;
            exit;
         end;
      end;

      if CmbOrder.ItemIndex = 1 then
         sOrderBy := ' Order by G.cod_jisa, G.desc_name'
      else if CmbOrder.ItemIndex = 2 then
         sOrderBy := ' Order by G.num_jumin'
      else if CmbOrder.ItemIndex = 3 then
         sOrderBy := ' Order by C.Desc_buwi'
      else if CmbOrder.ItemIndex = 4 then
         sOrderBy := ' Order by G.cod_jisa, G.gubn_jinch, G.num_samp'
      else if CmbOrder.ItemIndex = 5 then
         sOrderBy := ' Order by G.cod_jisa, G.dat_gmgn, G.num_samp'
      else if CmbOrder.ItemIndex = 6 then
         sOrderBy := ' Order by D.desc_dc, G.desc_name'
      else
         sOrderBy := ' Order by G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);
      Active := True;

      Gauge1.Progress := 0;
      if qryCell.IsEmpty = False then
      begin
         GP_query_log(qryCell, 'LD4Q33', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         // 조건에 맞는 데이터 변수에 입력
         while not Eof do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            Label4.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;
            //검사항목축출
            sCod_hm := '';
            if FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_cancer').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_cancer').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            if FieldByName('cod_jangbi').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_jangbi').AsString;
               qryPf_hm.Active := True;

               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            //---- 특검항목 추출...
            iRet := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_profile);
            for i := 0 to iRet-1 do
            begin
               if Trim(vCod_profile[i]) <> '' then
               begin
                  qryPf_hm2.Active := False;
                  qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[i];
                  qryPf_hm2.Active := True;
                  if qryPf_hm2.RecordCount > 0 then
                  begin
                     while not qryPf_hm2.Eof do
                     begin
                        check_tk := True;
                        for num := 1 to round(Length(sCod_hm)/4) do
                        begin
                           if copy(sCod_hm, (num * 4) - 3,4) =
                              qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                        end;
                        if check_tk = True then
                           sCod_hm := sCod_hm + qryPf_hm2.FieldByName('COD_HM').AsString;
                        qryPf_hm2.Next;
                     end; // end of while(항목 처리)
                  end; // end of if
               end; //end of if
            end; //end of for
            sCod_hm := sCod_hm + FieldByName('TkGum_chuga').AsString;
            sCod_hm := sCod_hm + FieldByName('cod_chuga').AsString;

            if (pos('P012',sCod_hm) > 0) and
               (Cmb_Gubun.ItemIndex = 0) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := 'P012';
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               
               sSelect := ''; sWhere  := ''; sOrderBy := '';
               with qryCell_Old do
               begin
                  qryCell_Old.Active := False;
                  sSelect := ' select top(1) dat_gmgn, GUBN_P012 from Cell ';
                  if pnlJisa.Visible then
                  begin
                     if Copy(cmbJisa.Text, 1, 2) <> '00' then
                        sWhere := ' WHERE cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And '
                     else
                        sWhere := ' WHERE ';
                  end
                  else
                     sWhere := ' WHERE cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And ';
                  sWhere := sWhere + ' num_id = ''' + qryCell.FieldByName('num_id').AsString + ''' ';
                  sWhere := sWhere + ' And dat_gmgn < ''' + qryCell.FieldByName('dat_gmgn').AsString + ''' ';
                  sWhere := sWhere + ' And cod_hm = ''P012'' ';
                  sOrderBy := ' order by dat_gmgn DESC ';
                  SQL.Clear;
                  SQL.Add(sSelect + sWhere + sOrderBy);
                  qryCell_Old.Active := True;
                  if qryCell_Old.IsEmpty = False then
                  begin
                     if qryCell_Old.FieldByName('GUBN_P012').AsString = '01' then
                        UV_vGUBN_P012[index] := '01-검체불량' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '02' then
                        UV_vGUBN_P012[index] := '02-음성ⅰ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '03' then
                        UV_vGUBN_P012[index] := '03-음성ⅱ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '04' then
                        UV_vGUBN_P012[index] := '04-의양성ⅳ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '05' then
                        UV_vGUBN_P012[index] := '05-양성ⅳ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '06' then
                        UV_vGUBN_P012[index] := '06-양성ⅴ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']';
                  end;
                  qryCell_Old.Active := False;
               end;
               Inc(index);
            end
            else if (pos('P013',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 1) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := 'P013';
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               with qryCell_Old do
               begin
                  qryCell_Old.Active := False;
                  sSelect := ' select top(1) dat_gmgn, GUBN_P012 from Cell ';
                  if pnlJisa.Visible then
                  begin
                     if Copy(cmbJisa.Text, 1, 2) <> '00' then
                        sWhere := ' WHERE cod_jisa = ''' + Copy(cmbJisa.Text, 1, 2) + ''' And '
                     else
                        sWhere := ' WHERE ';
                  end
                  else
                     sWhere := ' WHERE cod_jisa = ''' + copy(GV_sUserId,1,2) + ''' And ';
                  sWhere := sWhere + ' num_id = ''' + qryCell.FieldByName('num_id').AsString + ''' ';
                  sWhere := sWhere + ' And dat_gmgn < ''' + qryCell.FieldByName('dat_gmgn').AsString + ''' ';
                  sWhere := sWhere + ' And cod_hm = ''P013'' ';
                  sOrderBy := ' order by dat_gmgn DESC ';
                  SQL.Clear;
                  SQL.Add(sSelect + sWhere + sOrderBy);
                  qryCell_Old.Active := True;
                  if qryCell_Old.IsEmpty = False then
                  begin
                     if qryCell_Old.FieldByName('GUBN_P012').AsString = '01' then
                        UV_vGUBN_P012[index] := '01-검체불량' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '02' then
                        UV_vGUBN_P012[index] := '02-음성ⅰ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '03' then
                        UV_vGUBN_P012[index] := '03-음성ⅱ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '04' then
                        UV_vGUBN_P012[index] := '04-의양성ⅳ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '05' then
                        UV_vGUBN_P012[index] := '05-양성ⅳ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']'
                     else if qryCell_Old.FieldByName('GUBN_P012').AsString = '06' then
                        UV_vGUBN_P012[index] := '06-양성ⅴ' + '[' + qryCell_Old.FieldByName('dat_gmgn').AsString + ']';
                  end;
                  qryCell_Old.Active := False;
               end;
               Inc(index);
            end
            {else if (pos('B005',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 2) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := 'B005';
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               UV_vGUBN_P012[index]  := '';
               Inc(index);
            end        }
            else if (pos('B009',sCod_hm) > 0) and
                    (Cmb_Gubun.ItemIndex = 2) then
            begin
               iAge := 0; sMan := '';
               GP_JuminToAge(FieldByName('num_jumin').AsString,FieldByName('Dat_gmgn').AsString, sMan, iAge);
               UP_VarMemSet(index);
               UV_vCount[index]      := index + 1;
               UV_vCod_jisa[index]   := FieldByName('cod_jisa').AsString;
               UV_vCod_hm[index]     := 'B009';
               UV_vDat_gmgn[index]   := FieldByName('dat_gmgn').AsInteger;
               UV_vDesc_name[index]  := FieldByName('desc_name').AsString;
               UV_vNum_jumin[index]  := FieldByName('num_jumin').AsString;
               UV_vAge[index]        := iAge;
               UV_vSex[index]        := sMan;
               UV_vDoctor[index]     := COPY(FieldByName('doctor').AsString,1,8);
               if FieldByName('Desc_buwi').AsString = '' then
               UV_vDesc_buwi[index]  := ''
               else UV_vDesc_buwi[index]  := FieldByName('Desc_buwi').AsString + '/' + qryCell.FieldByName('num_id').AsString;
               UV_vDesc_dc[index]    := FieldByName('Desc_dc').AsString;
               UV_vGUBN_P012[index]  := '';
               Inc(index);
            end;
            Next;
         end;
      end
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
      // Table과 Disconnect
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;



procedure TfrmLD4Q33.grdMasterRowChanged(Sender: TObject; OldRow,
  NewRow: Integer);
begin
   inherited;

   // 자료가 존재하는지 Check
   if NewRow = 0 then exit;

   // Grid Focus
   grdMaster.SetFocus;

   // Data건수 표시
   GP_SetDataCnt(grdMaster);

   // 조회 Mode
   UP_SetMode('QUERY');
end;


procedure TfrmLD4Q33.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vCount[DataRow-1];
      2 : Value := UV_vCod_jisa[DataRow-1];
      3 : Value := UV_vCod_hm[DataRow-1];
      4 : Value := UV_vDat_gmgn[DataRow-1];
      5 : Value := UV_vDesc_name[DataRow-1];
      6 : value := UV_vNum_jumin[DataRow-1];
      7 : Value := UV_vAge[DataRow-1];
      8 : Value := UV_vSex[DataRow-1];
      9 : Value := UV_vDoctor[DataRow-1];
     10 : Value := UV_vDesc_buwi[DataRow-1];
     11 : Value := UV_vDesc_dc[DataRow-1];
     12 : Value := UV_vGUBN_P012[DataRow-1];
   end;
end;


procedure TfrmLD4Q33.grdMasterCellEdit(Sender: TObject; DataCol,
  DataRow: Integer; ByUser: Boolean);
begin
   inherited;
   UV_vGubn_check[DataRow - 1] := grdmaster.Cell[1, DataRow];

end;


procedure TfrmLD4Q33.UP_Click(Sender: TObject);
begin
   inherited;
   if Rd_Date.Checked = True then
   begin
      mskSDate.Enabled := True;
      mskEDate.Enabled := True;
      mskS_gmgn.Enabled := False;
      mskE_gmgn.Enabled := False;
      btnSDate.Enabled := True;
      btnEDate.Enabled := True;
      btnS_gmgn.Enabled := False;
      btnE_gmgn.Enabled := False;
   end
   else if Rd_Gmgn.Checked = True then
   begin
      mskSDate.Enabled := False;
      mskEDate.Enabled := False;
      mskS_gmgn.Enabled := True;
      mskE_gmgn.Enabled := True;
      btnSDate.Enabled := False;
      btnEDate.Enabled := False;
      btnS_gmgn.Enabled := True;
      btnE_gmgn.Enabled := True;
   end;

   if Sender = btnSDate then
      GF_CalendarClick(mskSDate);
   if Sender = btnEDate then
      GF_CalendarClick(mskEDate);
   if Sender = btnS_gmgn then
      GF_CalendarClick(mskS_gmgn);
   if Sender = btnE_gmgn then
      GF_CalendarClick(mskE_gmgn);
end;


procedure TfrmLD4Q33.up_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;
   if Key <> GC_Refer then exit;

   if Sender = mskSDate then UP_Click(btnSDate);
   if Sender = mskEDate then UP_Click(btnEDate);
   if Sender = mskS_gmgn then UP_Click(btnS_gmgn);
   if Sender = mskE_gmgn then UP_Click(btnE_gmgn);
end;

procedure TfrmLD4Q33.btnPrintClick(Sender: TObject);
begin
  inherited;
   if not GF_MsgBox('Confirmation', 'PRINT') then exit;

   frmLD4Q331 := TfrmLD4Q331.Create(Self);
   frmLD4Q331.QuickRep1.Preview;
end;

procedure TfrmLD4Q33.cmbRelationChange(Sender: TObject);
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
  if cmbRelation.ItemIndex = 1 then
  begin
     if Application.MessageBox('접수No. 가 없는 건에 대해서 접수No. 를 일괄부여 하시겠습니까?', '확인', MB_YESNO+MB_ICONquestion) = IDNO then
        Exit;

     iMax := StrToInt(Edt_Max.Text);

     for i := 0 to grdMaster.Rows - 1 do
     begin
        // 자료가 없을경우의 처리
        if grdMaster.Rows = 0 then
        begin
           exit;
        end;

        if (trim(UV_vDesc_buwi[i]) = '') then
        begin
           qryU_Cell.ParamByName('Cod_jisa').AsString   := UV_vCod_jisa[i];
           qryU_Cell.ParamByName('Dat_gmgn').AsString   := UV_vDat_gmgn[i];
           qryU_Cell.ParamByName('Num_jumin').AsString  := UV_vNum_jumin[i];
           //객담(p012),유린(p013),b009만적용 2018년 5월15일 최윤선..
           if (UV_vCod_hm[i] = 'B009') or (UV_vCod_hm[i] = 'P012') or (UV_vCod_hm[i] = 'P013') then
           qryU_Cell.ParamByName('dat_time').AsString   := formatdatetime('HH:NN:SS', now)
           else qryU_Cell.ParamByName('dat_time').AsString   := '';

           qryU_Cell.ParamByName('DAT_update').AsString := GV_sToday;
           qryU_Cell.ParamByName('COD_update').AsString := GV_sUserId;

           if CK_Max.Checked then    //곽윤설 20140514~
           begin
              if (Trim(Edt_Max.Text) <> '') and
                 (UV_vCod_hm[i] = 'P012') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end
              else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'P013') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end
              {else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'B005') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
              end }
              else if (Trim(Edt_Max.Text) <> '') and
                      (UV_vCod_hm[i] = 'B009') then
              begin
                 qryU_Cell.ParamByName('desc_buwi').Asstring := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(iMax), 6, '0');
                 qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                 qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
              end;
              iMax := iMax + 1;
           end              //~곽윤설 20140514
           else
           begin
             //P012 - 'S'로 시작   -> 2019.02.01 'P' 로시작
             //P013 - 'U'로 시작   -> 2019.01.01 'P' 로시작
             //B005 - 'T'로 시작
             //B009 - 'L'로 시작
             qryCa_desc_buwi_Max.Active := False;
             if (UV_vCod_hm[i] = 'P012') or
                (UV_vCod_hm[i] = 'P013') or
              //  (UV_vCod_hm[i] = 'B005') or
                (UV_vCod_hm[i] = 'B009') then  //곽윤설 20140514
             begin
                with qryCa_desc_buwi_Max do
                begin
                   qryCa_desc_buwi_Max.Active := False;

                   if Cmb_Gubun.ItemIndex = 0 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end
                   else
                   if Cmb_Gubun.ItemIndex = 1 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end
                   {else
                   if Cmb_Gubun.ItemIndex = 2 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end  }
                   else
                   if Cmb_Gubun.ItemIndex = 2 then
                   begin
                      sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ' +
                                 ' WHERE desc_buwi > ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''000000''' +
                                 ' AND   desc_buwi < ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' + ''-'' + ''999999''' +
                                 ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(UV_vDat_gmgn[i], 3, 2) + ''' ';
                   end;
                   qryCa_desc_buwi_Max.SQL.Clear;
                   qryCa_desc_buwi_Max.SQL.Add(sSelect);
                   qryCa_desc_buwi_Max.Active := True;
                end;
             end;

             if Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) = '' then
             begin
                if (UV_vCod_hm[i] = 'P012') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end
                else if (UV_vCod_hm[i] = 'P013') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end
               { else if (UV_vCod_hm[i] = 'B005') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                end  }
                else if (UV_vCod_hm[i] = 'B009') then
                begin
                   qryU_Cell.ParamByName('desc_buwi').AsString := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-' + '000001';
                   qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                   qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
                end;
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'P012') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'P013') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'P' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end
           {  else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'B005') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'T' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
             end }
             else if (Trim(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) <> '') and
                     (UV_vCod_hm[i] = 'B009') then
             begin
                qryU_Cell.ParamByName('desc_buwi').Asstring := 'L' + Copy(UV_vDat_gmgn[i], 3, 2) + '-'  + GF_LPad(IntToStr(StrToInt(qryCa_desc_buwi_Max.FieldByName('RESULT').AsString) + 1), 6, '0');
                qryU_Cell.ParamByName('cod_hm').AsString    := UV_vCod_hm[i];
                qryU_Cell.ParamByName('dat_last').AsString  := GV_sToday;
             end;
             // Table disconnect
             qryCa_desc_buwi_Max.Active := False;
           end; //end of 'CK_Max.Checked'
           qryU_Cell.ExecSQL;
           GP_query_log(qryU_Cell, 'LD4Q33', 'Q', 'Y');
        end;
     end;
     btnQuery.Click;
//     grdMaster.Repaint;
  end;
end;

function  TfrmLD4Q33.UF_DSawonCk : Boolean;
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
            if Trim(FieldByName('gubn_dept').AsString) = '' then
               Result := False
            else
               case FieldByName('gubn_dept').AsInteger of
                  12 : Result := True;
                  else
                     Result := False;
               end;
         end;
      end;
      // Table Disconnect
      Active := False;
   end;
end;

procedure TfrmLD4Q33.Cmb_GubunChange(Sender: TObject);
begin
  inherited;
   if Cmb_Gubun.ItemIndex = 0 then
   begin
      Edts_No.Text := 'S' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'S' + Copy(GV_sToday,3,2) + '-';
   end
   else if Cmb_Gubun.ItemIndex = 1 then
   begin
      Edts_No.Text := 'U' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'U' + Copy(GV_sToday,3,2) + '-';
   end
  { else if Cmb_Gubun.ItemIndex = 2 then
   begin
      Edts_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'T' + Copy(GV_sToday,3,2) + '-';
   end  }
   else if Cmb_Gubun.ItemIndex = 2 then
   begin
      Edts_No.Text := 'L' + Copy(GV_sToday,3,2) + '-';
      Edte_No.Text := 'L' + Copy(GV_sToday,3,2) + '-';
   end;
end;

procedure TfrmLD4Q33.CK_MaxClick(Sender: TObject);    //20140514 곽윤설~
var i, j, iMax : Integer;
    sSelect, sWhere, sOrderBy : String;
begin
  inherited;
   if CK_Max.Checked = True then
   begin
      Edt_Max.Color   := clWindow;
     //'S'로 시작
     if (Cmb_Gubun.ItemIndex = 0) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskSDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
     //'U'로 시작하는건
     else if (Cmb_Gubun.ItemIndex = 1) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''P'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''P'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''P'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end
    { //'T'로 시작하는건
     else if (Cmb_Gubun.ItemIndex = 2) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''T'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''T'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''T'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''T'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''T'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end }
     //'L'로 시작하는건
     else if (Cmb_Gubun.ItemIndex = 2) then
     begin
        with qryCa_desc_buwi_Max do
        begin
           qryCa_desc_buwi_Max.Active := False;

           sSelect := ' SELECT MAX(convert(int, substring(desc_buwi,5,6))) result FROM  cell with(nolock) ';
           if Rd_Date.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''L'' + ''' + Copy(mskSDate.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''L'' + ''' + Copy(mskEDate.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(mskEDate.Text, 3, 2) + '''';
           end
           else if Rd_gmgn.Checked then
           begin
              sSelect := sSelect + ' WHERE desc_buwi > ''L'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + ''' + ''-'' + ''000000''';
              sSelect := sSelect + ' AND   desc_buwi < ''L'' + ''' + Copy(mskE_gmgn.Text, 3, 2) + ''' + ''-'' + ''999999''';
              sSelect := sSelect + ' and substring(desc_buwi,1,3) = ''L'' + ''' + Copy(mskS_gmgn.Text, 3, 2) + '''';
           end;
           qryCa_desc_buwi_Max.SQL.Clear;
           qryCa_desc_buwi_Max.SQL.Add(sSelect);
           qryCa_desc_buwi_Max.Active := True;
        end;
     end;
     Edt_Max.Text := intToStr(qryCa_desc_buwi_Max.FieldByName('RESULT').Asinteger + 1);
     qryCa_desc_buwi_Max.Active := False;
   end
   else if CK_Max.Checked = False then
   begin
      Edt_Max.Text    := '';
      Edt_Max.Color   := $00E0E0E0;
   end;
end;  //~20140514 곽윤설

end.


