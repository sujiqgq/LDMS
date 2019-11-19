unit LD8P031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, Qrctrls, ExtCtrls, StdCtrls, Db, DBTables;

type
  TfrmLD8P031 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    ColumnHeaderBand1: TQRBand;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;

    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD8P031: TfrmLD8P031;

implementation

uses ZZFunc, MdmsFunc, LD8P03;
{$R *.DFM}

procedure TfrmLD8P031.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   with frmLD8P03.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      // 출력자료 전달
      UV_iCount := UV_iCount + 1;

      if UV_iCount <= Rows then
      begin

         QRLabel1.Caption  := Cell[5, UV_iCount];
                                          //20150121 곽윤설 [성명(생년) 출력 : 병리팀-박예진 요청]
         if Cell[3, UV_iCount] <> '' then QRLabel2.Caption  := '(' + Cell[ 3, UV_iCount] + ')'
         else                             QRLabel2.Caption  := '';
         if Cell[4, UV_iCount] <> '' then
         begin
         //a := (length(Cell[ 4, UV_iCount]));

         if (length(Cell[ 4, UV_iCount]) > 13)  then //15자 이상일 경우 라벨지 밀림 ..본원 최윤선 선임
           begin
             QRLabel2.Font.Size := 5;
             QRLabel2.Caption   := QRLabel2.Caption + Cell[ 4, UV_iCount]
           end
         else
           begin
             QRLabel2.Font.Size := 8;
             QRLabel2.Caption   := QRLabel2.Caption + Cell[ 4, UV_iCount];
           end;
         end;
         QRLabel3.Caption := '';
         if Cell[7, UV_iCount] <> '' then
         begin
              if Cell[7, UV_iCount] = 'P012' then QRLabel3.Caption := 'PapS'
              else if Cell[7, UV_iCount] = 'P013' then QRLabel3.Caption := 'PapU';
         end;
        If QRLabel3.Caption = '' then QRLabel3.Caption := Cell[6, UV_iCount];
         {
         if frmLD8P03.RBtn_HE.Checked then
            QRLabel3.Caption := 'H&E'
         else if frmLD8P03.RBtn_G.Checked then
            QRLabel3.Caption := 'G'
         else if frmLD8P03.RBtn_P.Checked then
            QRLabel3.Caption := 'PAP';
         }

         //  20141027 곽윤설 [라벨지 가로순으로 출력 : 병리팀-한영애 요청]
         if UV_iCount+4 <= Rows then
         begin
             QRLabel4.Caption  := Cell[5, UV_iCount+4];

                                                //20150121 곽윤설 [성명(생년) 출력 : 병리팀-박예진 요청]
             if Cell[3, UV_iCount+4] <> '' then QRLabel5.Caption  := '(' + Cell[ 3, UV_iCount+4] + ')'
             else                               QRLabel5.Caption  := '';
             if (length(Cell[ 4, UV_iCount+4]) > 13)  then //15자 이상일 경우 라벨지 밀림 ..본원 최윤선 선임
               begin
                 QRLabel5.Font.Size := 5;
                 QRLabel5.Caption   := QRLabel5.Caption + Cell[ 4, UV_iCount+4]
               end
             else
               begin
                 QRLabel5.Font.Size := 8;
                 QRLabel5.Caption   := QRLabel5.Caption + Cell[ 4, UV_iCount+4];
               end;
              QRLabel6.Caption := '';
              if Cell[7, UV_iCount+4] <> '' then
              begin
                   if Cell[7, UV_iCount+4] = 'P012' then QRLabel6.Caption := 'PapS'
                   else if Cell[7, UV_iCount+4] = 'P013' then QRLabel6.Caption := 'PapU';
              end;
              if QRLabel6.Caption ='' then QRLabel6.Caption := Cell[6, UV_iCount+4];
             {
             if frmLD8P03.RBtn_HE.Checked then
                QRLabel6.Caption := 'H&E'
             else if frmLD8P03.RBtn_G.Checked then
                QRLabel6.Caption := 'G'
             else if frmLD8P03.RBtn_P.Checked then
                QRLabel6.Caption := 'PAP';
             }
         end
         else
         begin
            QRLabel4.Caption  := '';
            QRLabel5.Caption  := '';
            QRLabel6.Caption  := '';
         end;

         if UV_iCount+8 <= Rows then
         begin
             QRLabel7.Caption  := Cell[5, UV_iCount+8];
                                                //20150121 곽윤설 [성명(생년) 출력 : 병리팀-박예진 요청]
             if Cell[3, UV_iCount+8] <> '' then QRLabel8.Caption  := '(' + Cell[ 3, UV_iCount+8] + ')'
             else                               QRLabel8.Caption  := '';
             if (length(Cell[ 4, UV_iCount+8]) > 13)  then //15자 이상일 경우 라벨지 밀림 ..본원 최윤선 선임
               begin
                 QRLabel8.Font.Size := 5;
                 QRLabel8.Caption   := QRLabel8.Caption + Cell[ 4, UV_iCount+8]
               end
             else
               begin
                 QRLabel8.Font.Size := 8;
                 QRLabel8.Caption   := QRLabel8.Caption + Cell[ 4, UV_iCount+8];
               end;
             QRLabel9.Caption := '';
             if Cell[7, UV_iCount+8] <> '' then
             begin
                  if Cell[7, UV_iCount+8] = 'P012' then QRLabel9.Caption := 'PapS'
                  else if Cell[7, UV_iCount+8] = 'P013' then QRLabel9.Caption := 'PapU';
             end;
             if QRLabel9.Caption = '' then QRLabel9.Caption := Cell[6, UV_iCount+8];

             {
             if frmLD8P03.RBtn_HE.Checked then
                QRLabel9.Caption := 'H&E'
             else if frmLD8P03.RBtn_G.Checked then
                QRLabel9.Caption := 'G'
             else if frmLD8P03.RBtn_P.Checked then
                QRLabel9.Caption := 'PAP';
             }
         end
         else
         begin
            QRLabel7.Caption  := '';
            QRLabel8.Caption  := '';
            QRLabel9.Caption  := '';
         end;
      end;

      if UV_iCount <= Rows then
      begin
         MoreData := True;
         if UV_iCount mod 4 = 0 then       //20141027 곽윤설
         begin
            UV_iCount := UV_iCount + 8;
         end;
      end
      else
         MoreData := False;
   end;
end;

procedure TfrmLD8P031.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 0;
end;

end.
