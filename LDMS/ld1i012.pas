//==============================================================================
// 프로그램 설명 : urin 일괄 정리
// 시스템        : 분석관리
// 수정일자      : 1999.06.03
// 수정자        : 김은석
// 수정내용      :
// 참고사항      :
//==============================================================================
// 수정일자      : 2011.08.13
// 수정자        : 김희석
// 수정내용      :  U046 일괄 정리 추가
//==============================================================================
// 수정일자      : 2012.02.06
// 수정자        : 유동구
// 수정내용      : 수원센터 와 혈검 = '3' 일 경우 돌아가지 않게.. 자체에서 돌리기..
//==============================================================================
// 수정일자      : 2012.11.27
// 수정자        : 유동구
// 수정내용      : U051 추가..
//==============================================================================
// 수정일자      : 2013.04.23
// 수정자        : 김희석
// 수정내용      : U004.U005,U009
//==============================================================================
// 수정일자      : 2013.0607
// 수정자        : 김희석
// 수정내용      : 검진불량및 샘플번호 조회 추가 
//==============================================================================
// 수정일자      : 2015.04.23
// 수정자        : 곽윤설
// 수정내용      : 센터검사일 때 U004, U005, U009 사용할수 있도록 수정
// 참고사항      : 광주 - 한송이책임 요청
//==============================================================================
unit LD1I012;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORDialog, StdCtrls, Buttons, ExtCtrls, Grids_ts, TSGrid, Db, DBTables,
  DBCtrls, Mask, DateEdit, Gauges;

type
  TfrmLD1I012 = class(TfrmDialog)
    Panel36: TPanel;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    Panel1: TPanel;
    mskDate: TDateEdit;
    qryGlkwa: TQuery;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    qryJaegumhm: TQuery;
    qryU_Glkwa: TQuery;
    qryPrev: TQuery;
    Gauge1: TGauge;
    Label1: TLabel;
    Label3: TLabel;
    cmbJisa: TComboBox;
    qryTkgum: TQuery;
    GroupBox1: TGroupBox;
    ckbY004: TCheckBox;
    ckbY005: TCheckBox;
    ckbU019: TCheckBox;
    ckbU013: TCheckBox;
    ckbU012: TCheckBox;
    ckbU011: TCheckBox;
    ckbU033: TCheckBox;
    Label4: TLabel;
    Lab_Num: TLabel;
    ckbU017: TCheckBox;
    ckbU024: TCheckBox;
    ckbU037: TCheckBox;
    ckbU029: TCheckBox;
    ckbU030: TCheckBox;
    ckbU031: TCheckBox;
    ckbU032: TCheckBox;
    ckbU001: TCheckBox;
    ckbU044: TCheckBox;
    ckbU045: TCheckBox;
    ckbU046: TCheckBox;
    ckbU051: TCheckBox;
    ckbU004: TCheckBox;
    ckbU005: TCheckBox;
    ckbU009: TCheckBox;
    ckbU053: TCheckBox;
    ckbU054: TCheckBox;
    ckbU055: TCheckBox;
    Chk_bul: TCheckBox;
    ckbU003: TCheckBox;
    mskTo: TMaskEdit;
    Label2: TLabel;
    mskFrom: TMaskEdit;
    rb_bunju: TRadioButton;
    rb_samp: TRadioButton;
    MEdt_SampS: TMaskEdit;
    Label13: TLabel;
    MEdt_SampE: TMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure UpClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }

    function  UF_Code(sHul, sCancer, sJangbi, sChuga : String) : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    function  UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
    function  UF_Gulkwa(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
    function  UF_Tk_Hangmok(sDate: String; iYh : Integer; iJumin : String): String;

  public
    { Public declarations }
  end;

var
  frmLD1I012: TfrmLD1I012;
  UV_iRows : Integer;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD1I01;

{$R *.DFM}


procedure TfrmLD1I012.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;

   // 분주일자를 오늘일자로 설정
   mskDate.Text := GV_sToday;

   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbJisa.Enabled := True;
      sJisa := Copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbJisa.Enabled := False;
      sJisa := GV_sJisa;
   end;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);
   rb_samp.Checked :=true;

end;

procedure TfrmLD1I012.UpClick(Sender: TObject);
var i, index : Integer;
    sTcode, sbGulkwa,
    sSelect, sWhere, sOrderby : String;
begin
   inherited;
   // 조회조건 Check
   //if not UF_Condition then exit;

   if Trim(mskFrom.Text) <> '' then mskFrom.Text := GF_LPad(Trim(mskFrom.Text), 4, '0');
   if Trim(mskTo.Text) <> '' then mskTo.Text := GF_LPad(Trim(mskTo.Text), 4, '0');

   index := 1;
   sSelect := ''; sWhere := ''; sOrderby := '';
   with qryGlkwa do
   begin
      Active := False;
      sSelect := ' SELECT A.cod_bjjs, A.num_id, A.dat_bunju, A.num_bunju, A.cod_jisa, A.desc_glkwa1, A.desc_glkwa2, A.desc_glkwa3, A.gubn_part, ' +
                 '        B.num_jumin, B.desc_name, B.dat_gmgn, B.cod_hul, B.cod_cancer, B.cod_chuga, B.cod_jangbi, ' +
                 '        B.gubn_nosin, B.gubn_nsyh, B.gubn_bogun, B.gubn_bgyh, B.gubn_adult, B.gubn_adyh, B.gubn_agab, B.gubn_agyh, ' +
                 '        B.gubn_gong, B.gubn_goyh, B.cod_jisa, B.gubn_tkgm, B.gubn_life, B.gubn_lfyh, B.gubn_hulgum ' +
                 ' FROM   gulkwa A INNER JOIN gumgin B ON A.cod_jisa = B.cod_jisa AND A.num_id = B.num_id AND A.dat_gmgn = B.dat_gmgn '+
                 '        left join  gum_bul G ON A.cod_jisa = G.cod_jisa AND A.num_id = G.num_id AND A.dat_gmgn = G.dat_gmgn and G.gubn_bul=''03''';  

      if chk_bul.Checked= True then
      begin    
           sWhere := sWhere+ 'Where  A.cod_bjjs  = ''' + copy(cmbJisa.Text,1,2) + '''';
           sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';
           sWhere := sWhere + ' AND (G.Gubn_bul !=''03'' or G.gubn_bul='' '' or G.gubn_bul is null )';
      end
      else
      begin
          sWhere := ' WHERE A.cod_bjjs  = ''' + copy(cmbJisa.Text,1,2) + '''';
          sWhere := sWhere + ' AND A.dat_bunju = ''' + mskDate.Text + '''';
      end;



      if  rb_samp.Checked =true then
      begin
           sWhere := sWhere + ' AND B.num_Samp >= ''' + MEdt_SampS.Text + '''';
           sWhere := sWhere + ' AND B.num_Samp <= ''' + MEdt_SampE.Text + '''';
      end
      else if rb_Bunju.Checked =true then
      begin
           sWhere := sWhere + ' AND A.num_bunju >= ''' + mskFrom.Text + '''';
           sWhere := sWhere + ' AND A.num_bunju <= ''' + mskTo.Text + '''';
      end;
      sWhere := sWhere + ' AND A.gubn_part = ''03''';

      SQL.Clear;
      SQL.Add(sSelect + sWhere);
      Active := True;

      Gauge1.MaxValue := RecordCount;
      Gauge1.Progress := 1;

      if RecordCount > 0 then
      begin
         GP_query_log(qryGlkwa, 'LD1I012', 'Q', 'Y');
         for i := 0 to RecordCount - 1 do
         begin
            Gauge1.Progress := Gauge1.Progress + 1;
            Lab_Num.Caption := FieldByName('Num_Bunju').AsString;
            application.ProcessMessages;

            UV_fGulkwa := '';
            UV_fGulkwa1 := FieldByName('DESC_GLKWA1').AsString;
            UV_fGulkwa2 := FieldByName('DESC_GLKWA2').AsString;
            UV_fGulkwa3 := FieldByName('DESC_GLKWA3').AsString;
            GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);

            sbGulkwa := UV_fGulkwa;

            sTcode := UF_Code(FieldByName('cod_hul').AsString, FieldByName('cod_cancer').AsString, FieldByName('cod_jangbi').AsString, FieldByName('cod_chuga').AsString);

            // 4. 항목코드에 대한 검사항목 추출
            if FieldByName('gubn_nosin').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '1', FieldByName('gubn_nsyh').AsInteger);

            if FieldByName('gubn_adult').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '4', FieldByName('gubn_adyh').AsInteger);

            if FieldByName('gubn_bogun').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '3',FieldByName('gubn_bgyh').AsInteger);

            if FieldByName('gubn_agab').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '5',FieldByName('gubn_agyh').AsInteger)
            else if FieldByName('gubn_agab').AsString = '2' then
               sTcode := sTcode + UF_Jae_Hangmok(FieldByName('cod_jisa').AsString, FieldByName('num_id').AsString, '5', FieldByName('gubn_agyh').AsInteger);

            if FieldByName('gubn_life').AsString = '1' then
               sTcode := sTcode + UF_No_Hangmok(FieldByName('dat_gmgn').AsString, '7',FieldByName('gubn_lfyh').AsInteger);

            //특검유형Display
            if FieldByName('gubn_tkgm').AsString <> '' then
               sTCode := sTCode + UF_Tk_Hangmok(FieldByName('Dat_gmgn').AsString,
                                                FieldByName('gubn_tkgm').AsInteger,
                                                FieldByName('Num_Jumin').AsString);
            //ji 02.10.21
            if (ckbY004.checked) or (ckbY005.checked) or (ckbU011.checked) or (ckbU012.checked) or
               (ckbU013.checked) or (ckbU019.checked) or (ckbU017.checked) or (ckbU024.checked) or
               (ckbU029.checked) or (ckbU030.checked) or (ckbU031.checked) or (ckbU032.checked) or
               (ckbU033.checked) or (ckbU037.checked) or (ckbU001.checked) or (ckbU044.checked) or
               (ckbU045.checked) or (ckbU046.checked) or (ckbU051.checked) or (ckbU004.checked) or
               (ckbU005.checked) or (ckbU009.checked) or (ckbU053.checked) or (ckbU054.checked) or
               (ckbU055.checked)  then
               begin
               if FieldByName('gubn_hulgum').AsString = '1' then
               begin
                    if (FieldByName('gubn_nosin').AsString ='1') or (FieldByName('gubn_adult').AsString ='1') or
                       (FieldByName('gubn_life').AsString ='1')  or (copy(FieldByName('cod_jangbi').AsString,1,2)='SS') then
                        begin
                        if ckbU004.Checked then
                           begin
                           if Pos('U004', sTcode) > 0 then
                              UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 25, 30, '5');
                        end;
                        if ckbU005.Checked then
                        begin
                           if Pos('U005', sTcode) > 0 then
                              UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 31, 36, '50');
                        end;
                        if ckbU009.Checked then
                        begin
                           if Pos('U009', sTcode) > 0 then
                              UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 55, 60, '5');
                        end;
                    end

               end;
               if (FieldByName('gubn_hulgum').AsString <> '2')   then
               begin
                    if ckbU053.Checked then
                    begin
                         if Pos('U053', sTcode) > 0 then
                            UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 337, 342, '음성');
                    end;
                    if ckbU054.Checked then
                    begin
                         if Pos('U054', sTcode) > 0 then
                            UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 343, 348, '음성');
                    end;
                    if ckbU055.Checked then
                    begin
                         if Pos('U055', sTcode) > 0 then
                            UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 349, 354, '음성');
                    end;
               end;

               if FieldByName('gubn_hulgum').AsString <> '3' then
               begin
                  if ckbU001.checked then
                  begin
                     Case Index of
                     1 : begin
                            if Pos('U010', sTcode) > 0 then
                            begin
                               UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '1.005');
                               if Pos('U003', sTcode) > 0 then
                                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '6.0');
                            end;
                          end;
                     2 : begin
                            if Pos('U010', sTcode) > 0 then
                            begin
                               UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '1.010');
                               if Pos('U003', sTcode) > 0 then
                                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '7.5');
                            end;
                         end;
                     3 : begin
                            if Pos('U010', sTcode) > 0 then
                            begin
                               UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '1.015');
                               if Pos('U003', sTcode) > 0 then
                                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '6.5');
                            end;
                         end;
                     4 : begin
                            if Pos('U010', sTcode) > 0 then
                            begin
                               UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '1.020');
                               if Pos('U003', sTcode) > 0 then
                                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '8.0');
                            end;
                         end;
                     5 : begin
                            if Pos('U010', sTcode) > 0 then
                            begin
                               UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 61, 66, '1.025');
                               if Pos('U003', sTcode) > 0 then
                                  UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '7.0');
                            end;
                         end;
                     end;

                     if Pos('U001', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 7, 12, '10');
                     if Pos('U002', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 13, 18, '음성');
                     if Pos('U004', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 25, 30, '5');
                     if Pos('U005', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 31, 36, '50');
                     if Pos('U006', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 37, 42, '5');
                     if Pos('U007', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 43, 48, '0.9');
                     if Pos('U008', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 49, 54, '0.4');
                     if Pos('U009', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 55, 60, '5');
                  end;

                  if ckbU011.Checked then
                  begin
                     if Pos('U011', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 67, 72, '0-2');
                  end;
                  if ckbU012.Checked then
                  begin
                     if Pos('U012', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 73, 78, '0-2');
                  end;
                  if ckbU013.Checked then
                  begin
                     if Pos('U013', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 79, 84, '0-2');
                  end;
                  if ckbU004.Checked then //20150423 곽윤설
                     begin
                     if Pos('U004', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 25, 30, '5');
                  end;
                  if ckbU005.Checked then
                  begin
                     if Pos('U005', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 31, 36, '50');
                  end;
                  if ckbU009.Checked then
                  begin
                     if Pos('U009', sTcode) > 0 then
                        UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 55, 60, '5');
                  end;
               end; // end if[본원센터 제외]

               if ckbY004.Checked then
               begin
                  if Pos('Y004', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 85, 90, '음성');
               end;
               if ckbY005.Checked then
               begin
                  if Pos('Y005', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 91, 96, '음성');
               end;
               if ckbU017.Checked then
               begin
                  if Pos('U017', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 115, 120, '음성');
               end;
               if ckbU019.Checked then
               begin
                  if Pos('U019', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 127, 132, '음성');
               end;
               if ckbU024.Checked then
               begin
                  if Pos('U024', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 145, 150, '음성');
               end;
               if ckbU029.Checked then
               begin
                  if Pos('U029', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 175, 180, '음성');
               end;
               if ckbU030.Checked then
               begin
                  if Pos('U030', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 181, 186, '음성');
               end;
               if ckbU031.Checked then
               begin
                  if Pos('U031', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 187, 192, '음성');
               end;
               if ckbU032.Checked then
               begin
                  if Pos('U032', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 193, 198, '음성');
               end;
               if ckbU033.Checked then
               begin
                  if Pos('U033', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 199, 204, '음성');
               end;
               if ckbU037.Checked then
               begin
                  if Pos('U037', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 223, 228, '음성');
               end;

               if ckbU044.Checked then
               begin
                  if Pos('U044', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 277, 282, '음성');
               end;
               if ckbU045.Checked then
               begin
                  if Pos('U045', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 283, 288, '음성');
               end;
              if ckbU046.Checked then
               begin
                  if Pos('U046', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 289, 294, '음성');
               end;
               if ckbU051.Checked then
               begin
                  if Pos('U051', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 319, 324, '음성');
               end;
               if ckbU003.Checked then
               begin
                  if Pos('U003', sTcode) > 0 then
                     UV_fGulkwa := UF_Gulkwa(UV_fGulkwa, 19, 24, '5.0');
               end;
            end
            else
            begin
               Showmessage('한가지 이상 체크하시기 바랍니다!');
               exit;
            end;

            if Trim(sbGulkwa) <> Trim(UV_fGulkwa) then
            begin
               qryU_Glkwa.ParamByName('COD_BJJS').AsString   := FieldByName('cod_bjjs').AsString;
               qryU_Glkwa.ParamByName('NUM_ID').AsString     := FieldByName('num_id').AsString;
               qryU_Glkwa.ParamByName('DAT_BUNJU').AsString  := FieldByName('dat_bunju').AsString;
               qryU_Glkwa.ParamByName('NUM_BUNJU').AsString  := FieldByName('num_bunju').AsString;
               qryU_Glkwa.ParamByName('GUBN_PART').AsString  := FieldByName('gubn_part').AsString;
               UV_fGulkwa1 := '';
               UV_fGulkwa2 := '';
               UV_fGulkwa3 := '';
               GF_HulGulkwa('1', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
               qryU_Glkwa.ParamByName('DESC_GLKWA1').AsString := UV_fGulkwa1;
               qryU_Glkwa.ParamByName('DESC_GLKWA2').AsString := UV_fGulkwa2;
               qryU_Glkwa.ParamByName('DESC_GLKWA3').AsString := UV_fGulkwa3;

               qryU_Glkwa.ParamByName('dat_update').AsString  := GV_sToday;
               qryU_Glkwa.ParamByName('cod_update').AsString  := GV_sUserId;


               qryU_Glkwa.Execsql;
               GP_query_log(qryU_Glkwa, 'LD1I012', 'Q', 'Y');
            end;
            Index := Index + 1;

            if Index > 5 then Index := 1;

            Next;
         end; //for
         // Table과 Disconnect
         Active := False;
      end // if
      else
      begin
         GF_MsgBox('Information', 'NODATA');
      end;
   end; //with
   Showmessage('정상적으로 작업을 완료하였습니다.!');
end;

function TfrmLD1I012.UF_Code(sHul, sCancer, sJangbi, sChuga : String) : String;
var sTcode : String;
begin
   Result := '';
   sTcode := '';
   with qryPf_hm do
   begin
      // 1. 혈액에 대한 검사항목 추출
      if sHul <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sHul;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. 종양에 대한 검사항목 추출
      if sCancer <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sCancer;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. 장비에 대한 검사항목 추출
      if sJangbi <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := sJangbi;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               sTcode := sTcode + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
   end;

   // 3. 추가코드에 대한 검사항목 추출
   sTcode := sTcode + sChuga;

   Result := sTcode;
end;
function  TfrmLD1I012.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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

function  TfrmLD1I012.UF_Jae_Hangmok(sJisa, sJumin, sGubun : String; iYh : Integer): String;
begin
   Result := '';
   with qryJaegumhm do
   begin
      Active := False;
      ParamByName('cod_jisa').AsString    := sJisa;
      ParamByName('num_id').AsString      := sJumin;
      ParamByName('gubn_code').AsString   := sGubun;
      ParamByName('num_sil').AsInteger    := iYh;
      Active := True;
      if RecordCount > 0 then
      begin
         Result := FieldByName('desc_hul').AsString;
      end;
      Active := False;
   end;
end;

function  TfrmLD1I012.UF_Gulkwa(sGulkwa : String ; iStart, iEnd : Integer ; sGul : String): String;
begin
   Result := sGulkwa;
   if Trim(Copy(sGulkwa, iStart, iEnd-iStart)) = '' then
      sGulkwa := GF_RPad(Copy(sGulkwa, 1, iStart-1), iStart-1, ' ')
               + GF_RPad(sGul, iEnd - iStart + 1, ' ')
               + Copy(sGulkwa, iEnd+1, Length(sGulkwa));
   Result := sGulkwa;
end;

procedure TfrmLD1I012.btnCancelClick(Sender: TObject);
begin
  inherited;
   Close;
end;

function  TfrmLD1I012.UF_Tk_Hangmok(sDate : String; iYh : Integer; iJumin : String): String;
Var
   i, J, K : Integer;
begin
   Result := '';
   with qrytkgum do
   begin
      Active := False;
      ParamByName('In_Num_jumin').AsString   := iJumin;
      ParamByName('In_Gubn_Cha').AsInteger    := iYh;
      ParamByName('In_Dat_Gmgn').AsString    := sDate;
      Active := True;
      J := Round(Length(FieldByName('Cod_prf').AsString) / 4);
      For I := 0 To J - 1 do
         With qrypf_Hm Do
            Begin
               Close;
               ParamByName('Cod_pf').AsString := Copy(qrytkgum.FieldByName('Cod_prf').AsString, I*4 + 1, 4);
               Open;
               For K := 1 To RecordCount do
                   Begin
                      Result := Result + FieldByName('Cod_hm').AsString;
                      Next;
                   End;
               Close;
            End;
      Result := result + FieldByName('Cod_Chuga').AsString;
      Active := False;
   end;
end;

end.
