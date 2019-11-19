//==============================================================================
// 프로그램 설명 : Barcode출력
// 시스템        : 통합검진
// 수정일자      : 2006.03.13
// 수정자        : 구수정
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2013.07.03
// 수정자        : 김희석
// 수정내용      : 바코드출력 추가
//==============================================================================
// 수정일자      : 2015.03.06
// 수정자        : 곽윤설
// 수정내용      : 렉번호 추가
// 참고사항      : 본원 - 엄하니
//==============================================================================
unit LD8P02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD8P02 = class(TfrmPrint)
    cmbbjjs: TComboBox;
    mskSt_date: TDateEdit;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    rbGubn_part: TRadioGroup;
    Bevel2: TBevel;
    Label8: TLabel;
    DEdt_date: TDateEdit;
    btnDate1: TSpeedButton;
    Panel1: TPanel;
    Label9: TLabel;
    Label10: TLabel;
    RBtn_gumgin: TRadioButton;
    RBtn_No: TRadioButton;
    ValEdit_Start: TValEdit;
    ValEdit_end: TValEdit;
    RBtn_kwangju: TRadioButton;
    Cmb_gubn: TComboBox;
    Label12: TLabel;
    mskFrom: TMaskEdit;
    Label13: TLabel;
    mskTo: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label14: TLabel;
    MEdt_SampS: TMaskEdit;
    Label15: TLabel;
    Label5: TLabel;
    cmbJisa: TComboBox;
    Many_Samp: TButton;
    MEdt_Barcode: TMaskEdit;
    Bevel3: TBevel;
    barcode: TLabel;
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure UP_ReportClick(Sender: TObject);
    procedure Many_SampClick(Sender: TObject);
    procedure Cmb_gubnChange(Sender: TObject);
    procedure MEdt_BarcodeExit(Sender: TObject);

  private
    { Private declarations }
     public


    // 지사코드
    UV_sCod_jisa : String;
    // SQL문 저장
   UV_sBasicSQL : String;

// Procedure UP_gulkwa;
 //procedure UP_Select;
    { Public declarations }
  end;

var
  frmLD8P02: TfrmLD8P02;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,MDmanySamp,
     LD8P021;

procedure TfrmLD8P02.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date)
   else if Sender = btnDate1 then GF_CalendarClick(DEdt_date);
end;

procedure TfrmLD8P02.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

      // 주민번호의 입력오류일 경우 처리



    if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate);



end;

procedure TfrmLD8P02.FormCreate(Sender: TObject);
var sJisa, sName : String;
begin
   inherited;
   GP_ComboJisa(cmbJisa);
   GP_ComboJisa(cmbBjjs);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbBjjs.Enabled := True;
      sJisa := copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbBjjs.Enabled := False;
      sJisa := GV_sJisa;
   end;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, '00',  2);
   GP_ComboDisplay(cmbBjjs, sJisa, 2);

   Cmb_gubn.ItemIndex := 0;
   
   // 오늘일자 설정
   mskST_Date.Text := GV_sToday;
   frmMDmanySamp := TfrmMDmanySamp.Create(Self);
   frmMDmanySamp.listSamp.items.clear;


end;

procedure TfrmLD8P02.UP_ReportClick(Sender: TObject);
var
   sWhere, sSelect, sOrderBy : string;
   index, i  : integer;
begin
   inherited;

   // 조회조건 Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD8P021 := TfrmLD8P021.Create(Self);
   
   frmLD8P021.UV_Check := False;
   if RBtn_kwangju.Checked then
   begin
      if rbAll.Checked then
      begin
         frmLD8P021.QuickRep.PrinterSettings.FirstPage := 0;
         frmLD8P021.QuickRep.PrinterSettings.LastPage  := 0;
      end
      else if rbPart.Checked then
      begin
         frmLD8P021.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
         frmLD8P021.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
      end;
      // 인쇄매수 설정
      frmLD8P021.QuickRep.PrinterSettings.Copies := spnCopy.Value;


      if RBtn_No.Checked then
      begin
         if (ValEdit_Start.Value <= 0) or (ValEdit_end.Value <= 0) then
         begin
            ShowMessage('입력이 올바르지 않습니다.');
            exit;
         end;
         frmLD8P021.UV_istart := ValEdit_Start.Value;
         frmLD8P021.UV_iend   := ValEdit_end.Value;
         frmLD8P021.UV_Date   := DEdt_date.Text;
         frmLD8P021.UV_Check := True;
      end
      else
      begin

      end;
      // Preview or Print
      if Sender = btnPreview then frmLD8P021.QuickRep.Preview
      else if Sender = btnPrint then frmLD8P021.QuickRep.Print;
   end
   else
   begin
      with frmLD8P021.qryGulkwa do
      begin
         Active := False;


         if Cmb_gubn.Text = '바코드' then
         begin
            sSelect := ' SELECT *,D.Desc_dc, A.desc_rackno ' +
                       ' FROM gumgin as p LEFT OUTER JOIN danche D  ON P.cod_dc = D.cod_dc' +
                       '                  LEFT OUTER JOIN Bunju  A ON p.cod_jisa = A.cod_jisa and p.dat_gmgn = A.dat_gmgn and p.num_id = A.num_id ';
            sWhere  := ' WHERE p.dat_gmgn = ''' + '20' + copy(MEdt_Barcode.Text,1,6) + '''';
            sWhere  := sWhere + ' AND p.num_samp = ''' + copy(MEdt_Barcode.Text,7,6) + '''';
         end

         else if Cmb_gubn.Text = '분주번호' then
         begin
            sSelect := ' SELECT D.desc_dc, P.desc_name, P.cod_dc, P.desc_dept, P.num_samp, P.num_jumin, P.dat_gmgn, B.num_bunju, P.cod_chuga, A.desc_rackno ' +
                       ' FROM Gulkwa B LEFT OUTER JOIN gumgin P ON B.cod_jisa = P.cod_jisa and B.num_id = P.num_id and B.dat_gmgn = P.dat_gmgn ' +
                       '               LEFT OUTER JOIN danche D ON P.cod_dc = D.cod_dc ' +
                       '               LEFT OUTER JOIN Bunju  A ON B.cod_bjjs = A.cod_bjjs and B.dat_Bunju = A.dat_bunju and B.num_id = A.num_id ';
            sWhere  := ' WHERE B.cod_bjjs = ''' + copy(cmbBjjs.Text,1,2) + '''';
            sWhere  := sWhere + ' and B.dat_bunju = ''' + mskSt_date.Text + '''';
            if copy(cmbJisa.Text,1,2) <> '00' then
               sWhere := sWhere + ' and B.cod_jisa = ''' + copy(cmbJisa.Text,1,2) + '''';
            case rbGubn_part.ItemIndex of
                 0 : sWhere := sWhere + ' and B.gubn_part = ''01''';
                 1 : sWhere := sWhere + ' and B.gubn_part = ''02''';
                 2 : sWhere := sWhere + ' and B.gubn_part = ''03''';
                 3 : sWhere := sWhere + ' and B.gubn_part = ''04''';
                 4 : sWhere := sWhere + ' and B.gubn_part = ''05''';
                 5 : sWhere := sWhere + ' and B.gubn_part = ''07''';
            end;
            if Trim(mskFrom.Text) <> '' then
               sWhere := sWhere + ' AND B.num_bunju >= ''' + mskFrom.Text + '''';
            if Trim(mskTo.Text) <> '' then
               sWhere := sWhere + ' AND B.num_bunju <= ''' + mskTo.Text + '''';

            sOrderBy := ' ORDER BY B.cod_jisa, CAST(B.num_bunju AS INT) DESC ';
         end
         else if Cmb_gubn.Text = '샘플번호' then
         begin
            sSelect := ' SELECT D.desc_dc, P.desc_name, P.cod_dc, P.desc_dept, P.num_samp, P.num_jumin, P.dat_gmgn, B.num_bunju, P.cod_chuga, A.desc_rackno ' +
                       ' FROM Gulkwa B LEFT OUTER JOIN gumgin P ON B.cod_jisa = P.cod_jisa and B.num_id = P.num_id and B.dat_gmgn = P.dat_gmgn ' +
                       '               LEFT OUTER JOIN danche D  ON P.cod_dc = D.cod_dc ' +
                       '               LEFT OUTER JOIN Bunju  A ON B.cod_bjjs = A.cod_bjjs and B.dat_Bunju = A.dat_bunju and B.num_id = A.num_id ';
            sWhere :=  ' WHERE B.cod_bjjs = ''' + copy(cmbBjjs.Text,1,2) + '''';
            sWhere := sWhere + ' and B.dat_bunju = ''' + mskSt_date.Text + '''';
            if copy(cmbJisa.Text,1,2) <> '00' then
               sWhere := sWhere + ' and B.cod_jisa = ''' + copy(cmbJisa.Text,1,2) + '''';
            case rbGubn_part.ItemIndex of
                 0 : sWhere := sWhere + ' and B.gubn_part = ''01''';
                 1 : sWhere := sWhere + ' and B.gubn_part = ''02''';
                 2 : sWhere := sWhere + ' and B.gubn_part = ''03''';
                 3 : sWhere := sWhere + ' and B.gubn_part = ''04''';
                 4 : sWhere := sWhere + ' and B.gubn_part = ''05''';
                 5 : sWhere := sWhere + ' and B.gubn_part = ''07''';
            end;
            if (Trim(MEdt_SampS.Text) <> '') and (Trim(MEdt_SampE.Text) = '') then
            begin
                 sWhere := sWhere + ' AND P.num_samp = ''' + MEdt_SampS.Text + '''';
            end
            else if (Trim(MEdt_SampS.Text) <> '') and (Trim(MEdt_SampE.Text) <> '') then
            begin
               sWhere := sWhere + ' AND P.num_samp >= ''' + MEdt_SampS.Text + '''';
               sWhere := sWhere + ' AND P.num_samp <= ''' + MEdt_SampE.Text + '''';
            end; 
            sOrderBy := ' ORDER BY B.cod_jisa, CAST(P.num_samp AS INT)  ';
         end
         else if (Cmb_gubn.Text = '샘플번호(다수)')  then
         begin
              sSelect := ' SELECT D.desc_dc, P.desc_name, P.cod_dc, P.desc_dept, P.num_samp, P.num_jumin, P.dat_gmgn, B.num_bunju, P.cod_chuga, A.desc_rackno ' +
                         ' FROM Gulkwa B LEFT OUTER JOIN gumgin P ON B.cod_jisa = P.cod_jisa and B.num_id = P.num_id and B.dat_gmgn = P.dat_gmgn ' +
                         '               LEFT OUTER JOIN danche D  ON P.cod_dc = D.cod_dc ' +
                         '               LEFT OUTER JOIN Bunju  A ON B.cod_bjjs = A.cod_bjjs and B.dat_Bunju = A.dat_bunju and B.num_id = A.num_id ';
              sWhere :=  ' WHERE B.cod_bjjs = ''' + copy(cmbBjjs.Text,1,2) + '''';
              sWhere := sWhere + ' and B.dat_bunju = ''' + mskSt_date.Text + '''';
              if copy(cmbJisa.Text,1,2) <> '00' then
                sWhere := sWhere + ' and B.cod_jisa = ''' + copy(cmbJisa.Text,1,2) + '''';
              case rbGubn_part.ItemIndex of
                 0 : sWhere := sWhere + ' and B.gubn_part = ''01''';
                 1 : sWhere := sWhere + ' and B.gubn_part = ''02''';
                 2 : sWhere := sWhere + ' and B.gubn_part = ''03''';
                 3 : sWhere := sWhere + ' and B.gubn_part = ''04''';
                 4 : sWhere := sWhere + ' and B.gubn_part = ''05''';
                 5 : sWhere := sWhere + ' and B.gubn_part = ''07''';
              end;
              if frmMDmanySamp.listSamp.items.count >0 then
              begin
                   sWhere := sWhere + ' AND ';
                   for i := 0 to frmMDmanySamp.listSamp.items.count-1 do
                   begin
                        frmMDmanysamp.listSamp.itemindex := i;
                        if i = 0 then
                           sWhere := sWhere + ' (P.Num_samp = ''' + copy(frmMDmanysamp.listSamp.items.strings[i],1,6) + ''' '
                        else
                            sWhere := sWhere + ' or P.num_samp = ''' + copy(frmMDmanySamp.listSamp.items.strings[i],1,6) + '''' ;
                   end;
                   sWhere := sWhere + ') ' ;
                   sOrderBy := ' ORDER BY B.cod_jisa, CAST(P.num_samp AS INT)  ';
             end
             else
             begin
                  ShowMessage('입력이 올바르지 않습니다.');
                  MEdt_SampS.setFocus;
                  exit;
             end;

         end;

         SQL.Clear;
         SQL.Add(sSelect + sWhere + sOrderBy);
         Active := True;
         if RecordCount > 0 then
         begin
            GP_query_log(frmLD8P021.qryGulkwa, 'LD8P02', 'Q', 'N');
            if rbAll.Checked then
            begin
               frmLD8P021.QuickRep.PrinterSettings.FirstPage := 0;
               frmLD8P021.QuickRep.PrinterSettings.LastPage  := 0;
            end
            else if rbPart.Checked then
            begin
               frmLD8P021.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
               frmLD8P021.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
            end;
            // 인쇄매수 설정
            frmLD8P021.QuickRep.PrinterSettings.Copies := spnCopy.Value;

            // Preview or Print
            if Sender = btnPreview then frmLD8P021.QuickRep.Preview
            else if Sender = btnPrint then frmLD8P021.QuickRep.Print;
         end;
      end; // end of with;
   end;

   frmMDmanySamp.ListSamp.Items.Clear;
   if Cmb_gubn.Text = '바코드' then
   begin
     MEdt_Barcode.text:='';
     MEdt_Barcode.SetFocus;
  end;
end;


procedure TfrmLD8P02.Many_SampClick(Sender: TObject);
begin
  inherited;
   frmMDmanySamp.ShowModal;
end;

procedure TfrmLD8P02.Cmb_gubnChange(Sender: TObject);
begin
  inherited;
  if Cmb_gubn.Text = '바코드' then
     begin
     MEdt_Barcode.visible:=True;
     barcode.visible:=True;
     MEdt_Barcode.SetFocus
     end
  else
  begin
       MEdt_Barcode.text:='';
       MEdt_Barcode.visible:=False;
       barcode.visible:=False;
  end;
end;

procedure TfrmLD8P02.MEdt_BarcodeExit(Sender: TObject);
begin
  inherited;
  if Sender = MEdt_Barcode then
   begin
      // 주민번호의 입력오류일 경우 처리
      if Length(MEdt_Barcode.Text) < 12 then exit;
      UP_ReportClick(btnPrint);
   end;
end;

end.
