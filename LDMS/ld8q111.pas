unit LD8Q111;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrintForm, ExtCtrls, QuickRpt, Qrctrls, Db, DBTables;

type
  TfrmLD8Q111 = class(TfrmPrintForm)
    TitleBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRS_Page: TQRSysData;
    QRBand1: TQRBand;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRL_STDATE: TQRLabel;
    QRL_EDDATE: TQRLabel;
    QRL_DASH: TQRLabel;
    QRLabel17: TQRLabel;
    DetailBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel3: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel16: TQRLabel;
    QRL_part: TQRLabel;
    QRL_code: TQRLabel;
    QRL_hm: TQRLabel;
    QRL_15: TQRLabel;
    QRL_12: TQRLabel;
    QRL_11: TQRLabel;
    QRL_43: TQRLabel;
    QRL_71: TQRLabel;
    QRL_61: TQRLabel;
    QRL_51: TQRLabel;
    QRL_sum: TQRLabel;
    QRLabel49: TQRLabel;
    QRL_S15: TQRLabel;
    QRL_S12: TQRLabel;
    QRL_S11: TQRLabel;
    QRL_S43: TQRLabel;
    QRL_S71: TQRLabel;
    QRL_S61: TQRLabel;
    QRL_S51: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel50: TQRLabel;
    QRLabel51: TQRLabel;
    QRLabel52: TQRLabel;
    QRLabel53: TQRLabel;
    QRLabel54: TQRLabel;
    QRLabel55: TQRLabel;
    QRLabel56: TQRLabel;
    QRLabel57: TQRLabel;
    QRLabel58: TQRLabel;
    QRLabel59: TQRLabel;
    QRLabel60: TQRLabel;
    QRLabel61: TQRLabel;
    QRLabel62: TQRLabel;
    QRLabel63: TQRLabel;
    QRLabel64: TQRLabel;
    QRLabel65: TQRLabel;
    QRLabel66: TQRLabel;
    QRLabel67: TQRLabel;
    QRLabel68: TQRLabel;
    QRLabel69: TQRLabel;
    QRLabel70: TQRLabel;
    QRLabel71: TQRLabel;
    QRLabel72: TQRLabel;
    QRLabel73: TQRLabel;
    QRLabel74: TQRLabel;
    QRLabel75: TQRLabel;
    QRLabel76: TQRLabel;
    QRLabel77: TQRLabel;
    QRLabel78: TQRLabel;
    QRLabel79: TQRLabel;
    QRLabel80: TQRLabel;
    QRLabel81: TQRLabel;
    QRLabel82: TQRLabel;
    QRLabel83: TQRLabel;

    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);    
  private
    { Private declarations }
    UV_iCount : Integer;
  public
    { Public declarations }
  end;

var
  frmLD8Q111: TfrmLD8Q111;

implementation

uses ZZFunc, ZZComm, MdmsFunc,LD8Q11;

{$R *.DFM}


procedure TfrmLD8Q111.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   inherited;
   with frmLD8Q11.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      QRL_part.Caption := Cell[1, UV_iCount];
      QRL_code.Caption := Cell[2, UV_iCount];
      QRL_hm.Caption   := Cell[3, UV_iCount];
      QRL_15.Caption   := Cell[4, UV_iCount];
      QRL_12.Caption   := Cell[5, UV_iCount];
      QRL_11.Caption   := Cell[6, UV_iCount];
      QRL_43.Caption   := Cell[7, UV_iCount];
      QRL_71.Caption   := Cell[8, UV_iCount];
      QRL_61.Caption   := Cell[9, UV_iCount];
      QRL_51.Caption   := Cell[10, UV_iCount];
      QRL_sum.Caption  := Cell[11, UV_iCount];


      if UV_iCount <= Rows then
      begin
         MoreData := True;
         Inc(UV_iCount);
      end
      else
         MoreData := False;
   end;

end;

procedure TfrmLD8Q111.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
Var
   UV_i : Integer;
begin
   inherited;


   UV_iCount := 1;
   //레벨총합
   QRL_S15.Caption := 'A : ' + frmLD8Q11.Edit1.Text;
   QRL_S12.Caption := 'B : ' + frmLD8Q11.Edit2.Text;
   QRL_S11.Caption := 'C1 : ' + frmLD8Q11.Edit3.Text;
   QRL_S43.Caption := 'C2 : ' + frmLD8Q11.Edit4.Text;
   QRL_S71.Caption := 'C3 : ' + frmLD8Q11.Edit5.Text;
   QRL_S61.Caption := 'D1 : ' + frmLD8Q11.Edit6.Text;
   QRL_S51.Caption := 'D2 : ' + frmLD8Q11.Edit7.Text;

   //15
   QRLabel14.caption := frmLD8Q11.Edit8.Text;
   QRLabel15.caption := frmLD8Q11.Edit9.Text;
   QRLabel19.caption := frmLD8Q11.Edit10.Text;
   QRLabel22.caption := frmLD8Q11.Edit11.Text;
   QRLabel23.caption := frmLD8Q11.Edit12.Text;
   QRLabel24.caption := frmLD8Q11.Edit13.Text;
   QRLabel25.caption := frmLD8Q11.Edit14.Text;
   //12
   QRLabel26.caption := frmLD8Q11.Edit15.Text;
   QRLabel27.caption := frmLD8Q11.Edit16.Text;
   QRLabel28.caption := frmLD8Q11.Edit17.Text;
   QRLabel29.caption := frmLD8Q11.Edit18.Text;
   QRLabel30.caption := frmLD8Q11.Edit19.Text;
   QRLabel31.caption := frmLD8Q11.Edit20.Text;
   QRLabel32.caption := frmLD8Q11.Edit21.Text;
   //11
   QRLabel33.caption := frmLD8Q11.Edit22.Text;
   QRLabel34.caption := frmLD8Q11.Edit23.Text;
   QRLabel35.caption := frmLD8Q11.Edit24.Text;
   QRLabel36.caption := frmLD8Q11.Edit25.Text;
   QRLabel37.caption := frmLD8Q11.Edit26.Text;
   QRLabel38.caption := frmLD8Q11.Edit27.Text;
   QRLabel39.caption := frmLD8Q11.Edit28.Text;
   //43
   QRLabel40.caption := frmLD8Q11.Edit29.Text;
   QRLabel41.caption := frmLD8Q11.Edit30.Text;
   QRLabel42.caption := frmLD8Q11.Edit31.Text;
   QRLabel43.caption := frmLD8Q11.Edit32.Text;
   QRLabel44.caption := frmLD8Q11.Edit33.Text;
   QRLabel45.caption := frmLD8Q11.Edit34.Text;
   QRLabel46.caption := frmLD8Q11.Edit35.Text;

   //71
   QRLabel47.caption := frmLD8Q11.Edit36.Text;
   QRLabel48.caption := frmLD8Q11.Edit37.Text;
   QRLabel50.caption := frmLD8Q11.Edit38.Text;
   QRLabel51.caption := frmLD8Q11.Edit39.Text;
   QRLabel52.caption := frmLD8Q11.Edit40.Text;
   QRLabel53.caption := frmLD8Q11.Edit41.Text;
   QRLabel54.caption := frmLD8Q11.Edit42.Text;

   //61
   QRLabel55.caption := frmLD8Q11.Edit43.Text;
   QRLabel56.caption := frmLD8Q11.Edit44.Text;
   QRLabel57.caption := frmLD8Q11.Edit45.Text;
   QRLabel58.caption := frmLD8Q11.Edit46.Text;
   QRLabel59.caption := frmLD8Q11.Edit47.Text;
   QRLabel60.caption := frmLD8Q11.Edit48.Text;
   QRLabel61.caption := frmLD8Q11.Edit49.Text;

   //51
   QRLabel62.caption := frmLD8Q11.Edit50.Text;
   QRLabel63.caption := frmLD8Q11.Edit51.Text;
   QRLabel64.caption := frmLD8Q11.Edit52.Text;
   QRLabel65.caption := frmLD8Q11.Edit53.Text;
   QRLabel66.caption := frmLD8Q11.Edit54.Text;
   QRLabel67.caption := frmLD8Q11.Edit55.Text;
   QRLabel68.caption := frmLD8Q11.Edit56.Text;


   // 조회조건을 전달
   if (frmLD8Q11.mskST_date.Text <> '') then
   begin
      QRL_STDATE.Caption   := Copy(frmLD8Q11.mskST_date.Text,1,4) + '/' +
                              Copy(frmLD8Q11.mskST_date.Text,5,2) + '/' +
                              Copy(frmLD8Q11.mskST_date.Text,7,2);
      QRL_EDDATE.Caption   := Copy(frmLD8Q11.mskED_date.Text,1,4) + '/' +
                              Copy(frmLD8Q11.mskED_date.Text,5,2) + '/' +
                              Copy(frmLD8Q11.mskED_date.Text,7,2);
      QRLabel17.Caption    := FormatDateTime('yyyy/mm/dd', Date);
   end
   else
   begin
     QRLabel20.Caption    := '(접수번호 :';
     QRL_STDATE.Caption   := frmLD8Q11.Edts_No.Text;
     QRL_EDDATE.Caption   := frmLD8Q11.Edte_No.Text;
     QRLabel17.Caption    := FormatDateTime('yyyy/mm/dd', Date);
   end;

end;

end.

