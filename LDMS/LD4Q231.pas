unit LD4Q231;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q231 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand2: TQRBand;
    QRShape13: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape2: TQRShape;
    QRShape1: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRL_A1: TQRLabel;
    QRL_B1: TQRLabel;
    QRL_C1: TQRLabel;
    QRL_D1: TQRLabel;
    QRL_E1: TQRLabel;
    QRL_F1: TQRLabel;
    QRL_G1: TQRLabel;
    QRL_H1: TQRLabel;
    QRL_A2: TQRLabel;
    QRL_B2: TQRLabel;
    QRL_C2: TQRLabel;
    QRL_D2: TQRLabel;
    QRL_E2: TQRLabel;
    QRL_F2: TQRLabel;
    QRL_A3: TQRLabel;
    QRL_A4: TQRLabel;
    QRL_A5: TQRLabel;
    QRL_A6: TQRLabel;
    QRL_A7: TQRLabel;
    QRL_A8: TQRLabel;
    QRL_A9: TQRLabel;
    QRL_A10: TQRLabel;
    QRL_A11: TQRLabel;
    QRL_A12: TQRLabel;
    QRL_B3: TQRLabel;
    QRL_B4: TQRLabel;
    QRL_B5: TQRLabel;
    QRL_B6: TQRLabel;
    QRL_B7: TQRLabel;
    QRL_B8: TQRLabel;
    QRL_B9: TQRLabel;
    QRL_B10: TQRLabel;
    QRL_B11: TQRLabel;
    QRL_B12: TQRLabel;
    QRL_C3: TQRLabel;
    QRL_C4: TQRLabel;
    QRL_C5: TQRLabel;
    QRL_C6: TQRLabel;
    QRL_C7: TQRLabel;
    QRL_C8: TQRLabel;
    QRL_C9: TQRLabel;
    QRL_C10: TQRLabel;
    QRL_C11: TQRLabel;
    QRL_C12: TQRLabel;
    QRL_D3: TQRLabel;
    QRL_E3: TQRLabel;
    QRL_G2: TQRLabel;
    QRL_H2: TQRLabel;
    QRL_F3: TQRLabel;
    QRL_G3: TQRLabel;
    QRL_H3: TQRLabel;
    QRL_D4: TQRLabel;
    QRL_E4: TQRLabel;
    QRL_F4: TQRLabel;
    QRL_G4: TQRLabel;
    QRL_H4: TQRLabel;
    QRL_D5: TQRLabel;
    QRL_E5: TQRLabel;
    QRL_F5: TQRLabel;
    QRL_G5: TQRLabel;
    QRL_H5: TQRLabel;
    QRL_D6: TQRLabel;
    QRL_D7: TQRLabel;
    QRL_E6: TQRLabel;
    QRL_E7: TQRLabel;
    QRL_F6: TQRLabel;
    QRL_F7: TQRLabel;
    QRL_G6: TQRLabel;
    QRL_G7: TQRLabel;
    QRL_H6: TQRLabel;
    QRL_H7: TQRLabel;
    QRL_D8: TQRLabel;
    QRL_E8: TQRLabel;
    QRL_F8: TQRLabel;
    QRL_G8: TQRLabel;
    QRL_H8: TQRLabel;
    QRL_D9: TQRLabel;
    QRL_E9: TQRLabel;
    QRL_F9: TQRLabel;
    QRL_G9: TQRLabel;
    QRL_H9: TQRLabel;
    QRL_F10: TQRLabel;
    QRL_F12: TQRLabel;
    QRL_E11: TQRLabel;
    QRL_D11: TQRLabel;
    QRL_G10: TQRLabel;
    QRL_E10: TQRLabel;
    QRL_D12: TQRLabel;
    QRL_E12: TQRLabel;
    QRL_H10: TQRLabel;
    QRL_H11: TQRLabel;
    QRL_H12: TQRLabel;
    QRL_G11: TQRLabel;
    QRL_D10: TQRLabel;
    QRL_F11: TQRLabel;
    QRL_G12: TQRLabel;
    QRLabel23: TQRLabel;
    Qrl_Date2: TQRLabel;
    Qrl_Place: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount, UV_iPage, UV_itemp, UV_iSpace : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4Q231: TfrmLD4Q231;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q23;

{$R *.DFM}

procedure TfrmLD4Q231.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var i, j : integer;
begin
   inherited;

   //i := 0; j := 0;
   with frmLD4Q23.grdS021 do
   begin
      // 자료가 없을경우의 처리
      if UV_iPage = 0 then
      begin
         MoreData := False;
         exit;
      end;



         QRL_A1.Font.Size := 11;
         QRL_B1.Font.Size := 11;
         QRL_C1.Font.Size := 11;
         QRL_D1.Font.Size := 11;
         QRL_E1.Font.Size := 11;
         QRL_E1.Font.Size := 11;
         QRL_F1.Font.Size := 11;
         QRL_G1.Font.Size := 11;
         QRL_H1.Font.Size := 11;
         QRL_A2.Font.Size := 11;
         QRL_B2.Font.Size := 11;
         QRL_C2.Font.Size := 11;

         QRL_A1.Caption := 'B';
         QRL_B1.Caption := '1';
         QRL_C1.Caption := '1';
         QRL_D1.Caption := '2';
         QRL_E1.Caption := '2';
         QRL_F1.Caption := '3';
         QRL_G1.Caption := '3';
         QRL_H1.Caption := '4';
         QRL_A2.Caption := '4';
         QRL_B2.Caption := '5';
         QRL_C2.Caption := '5';
         QRL_D2.Caption := Cell[4,(UV_iCount*85)-84+UV_iSpace];
         QRL_E2.Caption := Cell[4,(UV_iCount*85)-83+UV_iSpace];
         QRL_F2.Caption := Cell[4,(UV_iCount*85)-82+UV_iSpace];
         QRL_G2.Caption := Cell[4,(UV_iCount*85)-81+UV_iSpace];
         QRL_H2.Caption := Cell[4,(UV_iCount*85)-80+UV_iSpace];

         QRL_A3.Caption := Cell[4,(UV_iCount*85)-79+UV_iSpace];
         QRL_B3.Caption := Cell[4,(UV_iCount*85)-78+UV_iSpace];
         QRL_C3.Caption := Cell[4,(UV_iCount*85)-77+UV_iSpace];
         QRL_D3.Caption := Cell[4,(UV_iCount*85)-76+UV_iSpace];
         QRL_E3.Caption := Cell[4,(UV_iCount*85)-75+UV_iSpace];
         QRL_F3.Caption := Cell[4,(UV_iCount*85)-74+UV_iSpace];
         QRL_G3.Caption := Cell[4,(UV_iCount*85)-73+UV_iSpace];
         QRL_H3.Caption := Cell[4,(UV_iCount*85)-72+UV_iSpace];

         QRL_A4.Caption := Cell[4,(UV_iCount*85)-71+UV_iSpace];
         QRL_B4.Caption := Cell[4,(UV_iCount*85)-70+UV_iSpace];
         QRL_C4.Caption := Cell[4,(UV_iCount*85)-69+UV_iSpace];
         QRL_D4.Caption := Cell[4,(UV_iCount*85)-68+UV_iSpace];
         QRL_E4.Caption := Cell[4,(UV_iCount*85)-67+UV_iSpace];
         QRL_F4.Caption := Cell[4,(UV_iCount*85)-66+UV_iSpace];
         QRL_G4.Caption := Cell[4,(UV_iCount*85)-65+UV_iSpace];
         QRL_H4.Caption := Cell[4,(UV_iCount*85)-64+UV_iSpace];

         QRL_A5.Caption := Cell[4,(UV_iCount*85)-63+UV_iSpace];
         QRL_B5.Caption := Cell[4,(UV_iCount*85)-62+UV_iSpace];
         QRL_C5.Caption := Cell[4,(UV_iCount*85)-61+UV_iSpace];
         QRL_D5.Caption := Cell[4,(UV_iCount*85)-60+UV_iSpace];
         QRL_E5.Caption := Cell[4,(UV_iCount*85)-59+UV_iSpace];
         QRL_F5.Caption := Cell[4,(UV_iCount*85)-58+UV_iSpace];
         QRL_G5.Caption := Cell[4,(UV_iCount*85)-57+UV_iSpace];
         QRL_H5.Caption := Cell[4,(UV_iCount*85)-56+UV_iSpace];

         QRL_A6.Caption := Cell[4,(UV_iCount*85)-55+UV_iSpace];
         QRL_B6.Caption := Cell[4,(UV_iCount*85)-54+UV_iSpace];
         QRL_C6.Caption := Cell[4,(UV_iCount*85)-53+UV_iSpace];
         QRL_D6.Caption := Cell[4,(UV_iCount*85)-52+UV_iSpace];
         QRL_E6.Caption := Cell[4,(UV_iCount*85)-51+UV_iSpace];
         QRL_F6.Caption := Cell[4,(UV_iCount*85)-50+UV_iSpace];
         QRL_G6.Caption := Cell[4,(UV_iCount*85)-49+UV_iSpace];
         QRL_H6.Caption := Cell[4,(UV_iCount*85)-48+UV_iSpace];

         QRL_A7.Caption := Cell[4,(UV_iCount*85)-47+UV_iSpace];
         QRL_B7.Caption := Cell[4,(UV_iCount*85)-46+UV_iSpace];
         QRL_C7.Caption := Cell[4,(UV_iCount*85)-45+UV_iSpace];
         QRL_D7.Caption := Cell[4,(UV_iCount*85)-44+UV_iSpace];
         QRL_E7.Caption := Cell[4,(UV_iCount*85)-43+UV_iSpace];
         QRL_F7.Caption := Cell[4,(UV_iCount*85)-42+UV_iSpace];
         QRL_G7.Caption := Cell[4,(UV_iCount*85)-41+UV_iSpace];
         QRL_H7.Caption := Cell[4,(UV_iCount*85)-40+UV_iSpace];

         QRL_A8.Caption := Cell[4,(UV_iCount*85)-39+UV_iSpace];
         QRL_B8.Caption := Cell[4,(UV_iCount*85)-38+UV_iSpace];
         QRL_C8.Caption := Cell[4,(UV_iCount*85)-37+UV_iSpace];
         QRL_D8.Caption := Cell[4,(UV_iCount*85)-36+UV_iSpace];
         QRL_E8.Caption := Cell[4,(UV_iCount*85)-35+UV_iSpace];
         QRL_F8.Caption := Cell[4,(UV_iCount*85)-34+UV_iSpace];
         QRL_G8.Caption := Cell[4,(UV_iCount*85)-33+UV_iSpace];
         QRL_H8.Caption := Cell[4,(UV_iCount*85)-32+UV_iSpace];


         QRL_A9.Caption := Cell[4,(UV_iCount*85)-31+UV_iSpace];
         QRL_B9.Caption := Cell[4,(UV_iCount*85)-30+UV_iSpace];
         QRL_C9.Caption := Cell[4,(UV_iCount*85)-29+UV_iSpace];
         QRL_D9.Caption := Cell[4,(UV_iCount*85)-28+UV_iSpace];
         QRL_E9.Caption := Cell[4,(UV_iCount*85)-27+UV_iSpace];
         QRL_F9.Caption := Cell[4,(UV_iCount*85)-26+UV_iSpace];
         QRL_G9.Caption := Cell[4,(UV_iCount*85)-25+UV_iSpace];
         QRL_H9.Caption := Cell[4,(UV_iCount*85)-24+UV_iSpace];

         QRL_A10.Caption := Cell[4,(UV_iCount*85)-23+UV_iSpace];
         QRL_B10.Caption := Cell[4,(UV_iCount*85)-22+UV_iSpace];
         QRL_C10.Caption := Cell[4,(UV_iCount*85)-21+UV_iSpace];
         QRL_D10.Caption := Cell[4,(UV_iCount*85)-20+UV_iSpace];
         QRL_E10.Caption := Cell[4,(UV_iCount*85)-19+UV_iSpace];
         QRL_F10.Caption := Cell[4,(UV_iCount*85)-18+UV_iSpace];
         QRL_G10.Caption := Cell[4,(UV_iCount*85)-17+UV_iSpace];
         QRL_H10.Caption := Cell[4,(UV_iCount*85)-16+UV_iSpace];

         QRL_A11.Caption := Cell[4,(UV_iCount*85)-15+UV_iSpace];
         QRL_B11.Caption := Cell[4,(UV_iCount*85)-14+UV_iSpace];
         QRL_C11.Caption := Cell[4,(UV_iCount*85)-13+UV_iSpace];
         QRL_D11.Caption := Cell[4,(UV_iCount*85)-12+UV_iSpace];
         QRL_E11.Caption := Cell[4,(UV_iCount*85)-11+UV_iSpace];
         QRL_F11.Caption := Cell[4,(UV_iCount*85)-10+UV_iSpace];
         QRL_G11.Caption := Cell[4,(UV_iCount*85)-9+UV_iSpace];
         QRL_H11.Caption := Cell[4,(UV_iCount*85)-8+UV_iSpace];

         QRL_A12.Caption := Cell[4,(UV_iCount*85)-7+UV_iSpace];
         QRL_B12.Caption := Cell[4,(UV_iCount*85)-6+UV_iSpace];
         QRL_C12.Caption := Cell[4,(UV_iCount*85)-5+UV_iSpace];
         QRL_D12.Caption := Cell[4,(UV_iCount*85)-4+UV_iSpace];
         QRL_E12.Caption := Cell[4,(UV_iCount*85)-3+UV_iSpace];
         QRL_F12.Caption := Cell[4,(UV_iCount*85)-2+UV_iSpace];
         QRL_G12.Caption := Cell[4,(UV_iCount*85)-1+UV_iSpace];
         QRL_H12.Caption := Cell[4,UV_iCount*85+UV_iSpace];





//      end
 {     else
      begin
         QRL_A1.Font.Size := 7;
         QRL_B1.Font.Size := 7;
         QRL_C1.Font.Size := 7;
         QRL_D1.Font.Size := 7;
         QRL_E1.Font.Size := 7;

         // 5개 검사를 빼주고 계산
         QRL_A1.Caption := Cell[3,(UV_iCount*96)-95-UV_iSpace]; QRL_A2.Caption := Cell[3,(UV_iCount*96)-87-UV_iSpace];
         QRL_B1.Caption := Cell[3,(UV_iCount*96)-94-UV_iSpace]; QRL_B2.Caption := Cell[3,(UV_iCount*96)-86-UV_iSpace];
         QRL_C1.Caption := Cell[3,(UV_iCount*96)-93-UV_iSpace]; QRL_C2.Caption := Cell[3,(UV_iCount*96)-85-UV_iSpace];
         QRL_D1.Caption := Cell[3,(UV_iCount*96)-92-UV_iSpace]; QRL_D2.Caption := Cell[3,(UV_iCount*96)-84-UV_iSpace];
         QRL_E1.Caption := Cell[3,(UV_iCount*96)-91-UV_iSpace]; QRL_E2.Caption := Cell[3,(UV_iCount*96)-83-UV_iSpace];
         QRL_F1.Caption := Cell[3,(UV_iCount*96)-90-UV_iSpace]; QRL_F2.Caption := Cell[3,(UV_iCount*96)-82-UV_iSpace];
         QRL_G1.Caption := Cell[3,(UV_iCount*96)-89-UV_iSpace]; QRL_G2.Caption := Cell[3,(UV_iCount*96)-81-UV_iSpace];
         QRL_H1.Caption := Cell[3,(UV_iCount*96)-88-UV_iSpace]; QRL_H2.Caption := Cell[3,(UV_iCount*96)-80-UV_iSpace];

         QRL_A3.Caption := Cell[3,(UV_iCount*96)-79-UV_iSpace]; QRL_A4.Caption := Cell[3,(UV_iCount*96)-71-UV_iSpace];
         QRL_B3.Caption := Cell[3,(UV_iCount*96)-78-UV_iSpace]; QRL_B4.Caption := Cell[3,(UV_iCount*96)-70-UV_iSpace];
         QRL_C3.Caption := Cell[3,(UV_iCount*96)-77-UV_iSpace]; QRL_C4.Caption := Cell[3,(UV_iCount*96)-69-UV_iSpace];
         QRL_D3.Caption := Cell[3,(UV_iCount*96)-76-UV_iSpace]; QRL_D4.Caption := Cell[3,(UV_iCount*96)-68-UV_iSpace];
         QRL_E3.Caption := Cell[3,(UV_iCount*96)-75-UV_iSpace]; QRL_E4.Caption := Cell[3,(UV_iCount*96)-67-UV_iSpace];
         QRL_F3.Caption := Cell[3,(UV_iCount*96)-74-UV_iSpace]; QRL_F4.Caption := Cell[3,(UV_iCount*96)-66-UV_iSpace];
         QRL_G3.Caption := Cell[3,(UV_iCount*96)-73-UV_iSpace]; QRL_G4.Caption := Cell[3,(UV_iCount*96)-65-UV_iSpace];
         QRL_H3.Caption := Cell[3,(UV_iCount*96)-72-UV_iSpace]; QRL_H4.Caption := Cell[3,(UV_iCount*96)-64-UV_iSpace];

         QRL_A5.Caption := Cell[3,(UV_iCount*96)-63-UV_iSpace]; QRL_A6.Caption := Cell[3,(UV_iCount*96)-55-UV_iSpace];
         QRL_B5.Caption := Cell[3,(UV_iCount*96)-62-UV_iSpace]; QRL_B6.Caption := Cell[3,(UV_iCount*96)-54-UV_iSpace];
         QRL_C5.Caption := Cell[3,(UV_iCount*96)-61-UV_iSpace]; QRL_C6.Caption := Cell[3,(UV_iCount*96)-53-UV_iSpace];
         QRL_D5.Caption := Cell[3,(UV_iCount*96)-60-UV_iSpace]; QRL_D6.Caption := Cell[3,(UV_iCount*96)-52-UV_iSpace];
         QRL_E5.Caption := Cell[3,(UV_iCount*96)-59-UV_iSpace]; QRL_E6.Caption := Cell[3,(UV_iCount*96)-51-UV_iSpace];
         QRL_F5.Caption := Cell[3,(UV_iCount*96)-58-UV_iSpace]; QRL_F6.Caption := Cell[3,(UV_iCount*96)-50-UV_iSpace];
         QRL_G5.Caption := Cell[3,(UV_iCount*96)-57-UV_iSpace]; QRL_G6.Caption := Cell[3,(UV_iCount*96)-49-UV_iSpace];
         QRL_H5.Caption := Cell[3,(UV_iCount*96)-56-UV_iSpace]; QRL_H6.Caption := Cell[3,(UV_iCount*96)-48-UV_iSpace];

         QRL_A7.Caption := Cell[3,(UV_iCount*96)-47-UV_iSpace]; QRL_A8.Caption := Cell[3,(UV_iCount*96)-39-UV_iSpace];
         QRL_B7.Caption := Cell[3,(UV_iCount*96)-46-UV_iSpace]; QRL_B8.Caption := Cell[3,(UV_iCount*96)-38-UV_iSpace];
         QRL_C7.Caption := Cell[3,(UV_iCount*96)-45-UV_iSpace]; QRL_C8.Caption := Cell[3,(UV_iCount*96)-37-UV_iSpace];
         QRL_D7.Caption := Cell[3,(UV_iCount*96)-44-UV_iSpace]; QRL_D8.Caption := Cell[3,(UV_iCount*96)-36-UV_iSpace];
         QRL_E7.Caption := Cell[3,(UV_iCount*96)-43-UV_iSpace]; QRL_E8.Caption := Cell[3,(UV_iCount*96)-35-UV_iSpace];
         QRL_F7.Caption := Cell[3,(UV_iCount*96)-42-UV_iSpace]; QRL_F8.Caption := Cell[3,(UV_iCount*96)-34-UV_iSpace];
         QRL_G7.Caption := Cell[3,(UV_iCount*96)-41-UV_iSpace]; QRL_G8.Caption := Cell[3,(UV_iCount*96)-33-UV_iSpace];
         QRL_H7.Caption := Cell[3,(UV_iCount*96)-40-UV_iSpace]; QRL_H8.Caption := Cell[3,(UV_iCount*96)-32-UV_iSpace];

         QRL_A9.Caption := Cell[3,(UV_iCount*96)-31-UV_iSpace]; QRL_A10.Caption := Cell[3,(UV_iCount*96)-23-UV_iSpace];
         QRL_B9.Caption := Cell[3,(UV_iCount*96)-30-UV_iSpace]; QRL_B10.Caption := Cell[3,(UV_iCount*96)-22-UV_iSpace];
         QRL_C9.Caption := Cell[3,(UV_iCount*96)-29-UV_iSpace]; QRL_C10.Caption := Cell[3,(UV_iCount*96)-21-UV_iSpace];
         QRL_D9.Caption := Cell[3,(UV_iCount*96)-28-UV_iSpace]; QRL_D10.Caption := Cell[3,(UV_iCount*96)-20-UV_iSpace];
         QRL_E9.Caption := Cell[3,(UV_iCount*96)-27-UV_iSpace]; QRL_E10.Caption := Cell[3,(UV_iCount*96)-19-UV_iSpace];
         QRL_F9.Caption := Cell[3,(UV_iCount*96)-26-UV_iSpace]; QRL_F10.Caption := Cell[3,(UV_iCount*96)-18-UV_iSpace];
         QRL_G9.Caption := Cell[3,(UV_iCount*96)-25-UV_iSpace]; QRL_G10.Caption := Cell[3,(UV_iCount*96)-17-UV_iSpace];
         QRL_H9.Caption := Cell[3,(UV_iCount*96)-24-UV_iSpace]; QRL_H10.Caption := Cell[3,(UV_iCount*96)-16-UV_iSpace];

         QRL_A11.Caption := Cell[3,(UV_iCount*96)-15-UV_iSpace]; QRL_A12.Caption := Cell[3,(UV_iCount*96)-7-UV_iSpace];
         QRL_B11.Caption := Cell[3,(UV_iCount*96)-14-UV_iSpace]; QRL_B12.Caption := Cell[3,(UV_iCount*96)-6-UV_iSpace];
         QRL_C11.Caption := Cell[3,(UV_iCount*96)-13-UV_iSpace]; QRL_C12.Caption := Cell[3,(UV_iCount*96)-5-UV_iSpace];
         QRL_D11.Caption := Cell[3,(UV_iCount*96)-12-UV_iSpace]; QRL_D12.Caption := Cell[3,(UV_iCount*96)-4-UV_iSpace];
         QRL_E11.Caption := Cell[3,(UV_iCount*96)-11-UV_iSpace]; QRL_E12.Caption := Cell[3,(UV_iCount*96)-3-UV_iSpace];
         QRL_F11.Caption := Cell[3,(UV_iCount*96)-10-UV_iSpace]; QRL_F12.Caption := Cell[3,(UV_iCount*96)-2-UV_iSpace];
         QRL_G11.Caption := Cell[3,(UV_iCount*96)-9-UV_iSpace];  QRL_G12.Caption := Cell[3,(UV_iCount*96)-1-UV_iSpace];
         QRL_H11.Caption := Cell[3,(UV_iCount*96)-8-UV_iSpace];  QRL_H12.Caption := Cell[3,(UV_iCount*96)-UV_iSpace];
      end;
  }
      QRL_Place.Caption := intTostr(UV_iCount) + '. S021 H.P';

      if UV_iCount <= UV_iPage then
      begin
         UV_iCount := UV_iCount + 1;
         MoreData := True;
      end
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q231.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   if frmLD4Q23.edtSDate.Text = frmLD4Q23.edtEDate.Text then
   begin
      QRL_date.Caption := copy(frmLD4Q23.EdtSDate.Text,1,4) + '년'  +
                          copy(frmLD4Q23.EdtSDate.Text,5,2) + '월'  +
                          copy(frmLD4Q23.EdtSDate.Text,7,2) + '일';
      Qrl_Date2.Caption := copy(frmLD4Q23.EdtSDate.Text,5,2) + '/' + copy(frmLD4Q23.EdtSDate.Text,7,2);
   end
   else
   begin
      QRL_date.Caption := copy(frmLD4Q23.EdtSDate.Text,1,4) + '년'  +
                          copy(frmLD4Q23.EdtSDate.Text,5,2) + '월'  +
                          copy(frmLD4Q23.EdtSDate.Text,7,2) + '일'  +
                          ' ~ ' + copy(frmLD4Q23.EdtEDate.Text,7,2) + '일';
      Qrl_Date2.Caption := copy(frmLD4Q23.EdtSDate.Text,5,2) + '/' + copy(frmLD4Q23.EdtSDate.Text,7,2) +
                           '~' + copy(frmLD4Q23.EdtEDate.Text,7,2);
   end;
   UV_iCount := 1;
   UV_itemp  := 0;
   UV_iSpace := 0;
   UV_iPage  := 0;

   if (frmLD4Q23.grdS021.Rows > 0) AND (frmLD4Q23.grdS021.Rows <= 85) then
      UV_iPage := 1
   else if (frmLD4Q23.grdS021.Rows > 85) AND (frmLD4Q23.grdS021.Rows <= 170) then
      UV_iPage := 2
   else if frmLD4Q23.grdS021.Rows > 170 then
   begin
      UV_iPage := frmLD4Q23.grdS021.Rows div 170;
      UV_iPage := UV_iPage * 2;
      if frmLD4Q23.grdS021.Rows mod 170 <= 85 then
         UV_iPage := UV_iPage + 1
      else UV_iPage := UV_iPage + 2;
   end;
end;

end.
