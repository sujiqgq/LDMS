//==============================================================================
// 프로그램 설명 : 특수건강진단 분석 비용 내역서
// 시스템        : 통합검진
// 수정일자      : 2008.01.09
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD4Q202;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD4Q202 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel12: TQRLabel;
    qrlGumDate: TQRLabel;
    QRLabel11: TQRLabel;
    QRBand2: TQRBand;
    QRLabel10: TQRLabel;
    QRLabel14: TQRLabel;
    QRBand1: TQRBand;
    QRS_Page: TQRSysData;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel9: TQRLabel;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel2: TQRLabel;
    QRL_Date: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape9: TQRShape;
    QRLabel15: TQRLabel;
    QRShape10: TQRShape;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape11: TQRShape;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel25: TQRLabel;
    QRShape12: TQRShape;
    QRShape14: TQRShape;
    QRShape16: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRL1: TQRLabel;
    QRL2: TQRLabel;
    QRShape_1: TQRShape;
    QRL3: TQRLabel;
    QRL6: TQRLabel;
    QRL7: TQRLabel;
    QRShape_2: TQRShape;
    QRShape_7: TQRShape;
    QRShape_8: TQRShape;
    QRShape_9: TQRShape;
    QRShape_11: TQRShape;
    QRBand3: TQRBand;
    qryTemp: TQuery;
    QRShape22: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    qrlBunjuDate: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRL4: TQRLabel;
    QRL5: TQRLabel;
    QRShape15: TQRShape;
    QRLSumSuga: TQRLabel;
    QRShape33: TQRShape;
    QRLabel37: TQRLabel;
    QRShape13: TQRShape;
    QRLSumCnt: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }
    UV_iCount : Integer;
    UV_sSumCnt, UV_sSumSuga : string;

    Procedure UP_Clear;

  public
    { Public declarations }
  end;

var
  frmLD4Q202: TfrmLD4Q202;

implementation

{$R *.DFM}

uses ZZfunc, LD4Q20;

Procedure TfrmLD4Q202.UP_Clear;
begin
   QRL1.Caption := '';
   QRL2.Caption := '';
   QRL3.Caption := '';
   QRL4.Caption := '';
   QRL5.Caption := '';
   QRL6.Caption := '';
   QRL7.Caption := '';
end;

procedure TfrmLD4Q202.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;

   with frmLD4Q20.Grd_List do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      UP_Clear;

      if UV_iCount < Rows then
      begin
         // 출력자료 전달
         QRL1.Caption := Cell[1, UV_iCount];
         QRL2.Caption := Cell[2, UV_iCount];
         QRL3.Caption := Cell[3, UV_iCount];
         QRL4.Caption := Cell[4, UV_iCount];
         QRL5.Caption := Cell[5, UV_iCount];
         QRL6.Caption := Cell[6, UV_iCount];
         QRL7.Caption := Cell[7, UV_iCount];
      end
      else if UV_iCount = Rows then
      begin
         UV_sSumCnt  := Cell[6, UV_iCount];
         UV_sSumSuga := Cell[7, UV_iCount];
      end;

      if UV_iCount < Rows  then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD4Q202.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var
  sSelect, sWhere : string;
begin
   inherited;
   UV_sSumCnt := '';  UV_sSumSuga := '';
   UV_iCount := 1;

   // 조회조건을 전달
   with frmLD4Q20 do
   begin
      qrlBunjuDate.Caption  := GF_DateFormat(msksDate.Text) + ' ~ ' +
                               GF_DateFormat(mskeDate.Text);
      QRL_Date.Caption := GV_sPrintToday;

      //검진일자 구하기...
      //------------------------------------------------------------------------
      sSelect := ' SELECT Min(G.dat_gmgn) Dat_Min, Max(G.dat_gmgn) Dat_Max ' +
                 ' FROM Gumgin G LEFT OUTER JOIN gulkwa B ON G.cod_jisa = B.cod_jisa and G.num_id = B.num_id and G.dat_gmgn = B.dat_gmgn ' +
                 '               LEFT OUTER JOIN tkgum  T ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha ';

      sWhere := ' WHERE B.dat_bunju >= ''' + msksDate.Text + '''';
      sWhere := sWhere + ' AND B.dat_bunju <= ''' + mskeDate.Text + '''';

      if UV_sCod_jisa = '15' then
      begin
         sWhere := sWhere + ' AND B.cod_bjjs = ''' + UV_sCod_jisa + '''';
         if Trim(cmbCod_jisa.Text) <> '' then sWhere := sWhere + ' AND G.cod_jisa = ''' + Copy(cmbCod_jisa.Text, 1, 2) + '''';
      end;

      //검사파트...
      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';

      //배치전제외...
      if CheckBox.Checked then sWhere := sWhere + ' AND T.cod_prf Not Like ''%TC77%'' ';

      //소방서만 조회...
      if ChkBox_Fire.Checked then sWhere := sWhere + ' AND (T.Cod_Prf Like ''%TCA8%'' or T.Cod_Prf Like ''%TCB2%'') ';

      sWhere := sWhere + ' AND B.gubn_part = ''' + copy(Com_Part.Text,1,2) + '''';


      qryTemp.Active := False;
      qryTemp.SQL.Clear;
      qryTemp.SQL.text := sSelect + sWhere;
      qryTemp.Active := True;

      qrlGumDate.Caption := GF_DateFormat(qryTemp.FieldByName('Dat_Min').AsString) + ' ~ ' +
                            GF_DateFormat(qryTemp.FieldByName('Dat_Max').AsString);
      //========================================================================

      //수탁의뢰기관
      //------------------------------------------------------------------------
      qryTemp.Active := False;
      qryTemp.Sql.Text := ' SELECT * FROM JISA WHERE COD_JISA = ''' + copy(cmbCOD_JISA.Text,1,2) + ''' ';
      qryTemp.Active := True;

      if qryTemp.RecordCount > 0 then
      begin
         if pos('(재)한국의학연구소',qryTemp.FieldByName('desc_hlbw').AsString) > 0 then QRLabel17.Caption := qryTemp.FieldByName('desc_hlbw').AsString
         else                                                                            QRLabel17.Caption := '(재)한국의학연구소 ' + qryTemp.FieldByName('desc_hlbw').AsString;
      end;
      //========================================================================
   end;

   //[2017.05.29-유동구]한의 재단특수검사파트1700223
   //---------------------------------------------------------------------------
   //변경전-(총괄자)김남재 / (담당자)양새하
   //변경후-(총괄자)조찬우 / (담당자)박소연
   //분석운영총괄자(2017년 2월 이후 변경)
   //[2017.10.26-유동구]총괄자 변경(한의 재단특수검사파트1700352)   
   if      frmLD4Q20.msksDate.Text >= '20170920' then QRLabel15.Caption := '김경연'
   else if frmLD4Q20.msksDate.Text >= '20170201' then QRLabel15.Caption := '조찬우'
   else                                               QRLabel15.Caption := '김남재';

   //분석 담당자(2017년 5월 이후 변경)
   if frmLD4Q20.msksDate.Text >= '20170501' then QRLabel34.Caption := '박소연'
   else                                          QRLabel34.Caption := '양새하';
   //===========================================================================
end;

procedure TfrmLD4Q202.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
   inherited;
   QRLSumCnt.Caption  := UV_sSumCnt;
   QRLSumSuga.Caption := UV_sSumSuga;
end;

end.
