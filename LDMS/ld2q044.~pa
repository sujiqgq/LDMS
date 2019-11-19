unit LD2Q044;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QBarcode, QuickRpt, Qrctrls, ExtCtrls;

type
  TfrmLD2Q044 = class(TForm)
    QuickRep: TQuickRep;
    QRBand1: TQRBand;
    QRL_Jisa: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_samp: TQRLabel;
    QRL_Bunju: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    UV_iCount : Integer;
    { Private declarations }
  public
    { Public declarations }
  end;

var  frmLD2Q044: TfrmLD2Q044;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD2Q04;

{$R *.DFM}

procedure TfrmLD2Q044.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
var
   sCheck : String;
begin
   inherited;
   sCheck := 'N';

   with frmLD2Q04.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;

      repeat
         if UV_iCount <= Rows then
         begin
            // 출력자료 전달
            if Cell[ 6,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[ 7,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[ 8,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[ 9,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[10,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[11,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[12,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[13,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[14,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[15,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[16,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[17,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[18,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[19,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[20,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[21,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[22,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[24,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[25,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[26,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[27,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[28,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[29,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[30,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[31,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[32,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[33,UV_iCount]  <> '*' then sCheck := 'Y';
//          if Cell[34,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[35,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[36,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[37,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[38,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[39,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[40,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[41,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[42,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[43,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[44,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[45,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[46,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[47,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[48,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[49,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[50,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[51,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[52,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[53,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[54,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[55,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[56,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[57,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[58,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[59,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[60,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[61,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[62,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[63,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[64,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[65,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[66,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[67,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[68,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[69,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[70,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[71,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[72,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[73,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[74,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[75,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[76,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[77,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[78,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[79,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[80,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[81,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[82,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[83,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[84,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[85,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[86,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[87,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[88,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[89,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[90,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[91,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[92,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[93,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[94,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[95,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[96,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[97,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[98,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[99,UV_iCount]  <> '*' then sCheck := 'Y';
            if Cell[100,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[101,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[102,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[103,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[104,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[105,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[106,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[107,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[108,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[109,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[110,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[112,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[113,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[114,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[115,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[116,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[117,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[118,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[119,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[120,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[121,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[122,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[123,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[124,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[125,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[126,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[127,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[128,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[129,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[130,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[131,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[132,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[133,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[134,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[135,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[136,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[137,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[138,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[139,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[140,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[141,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[142,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[143,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[144,UV_iCount] <> '*' then sCheck := 'Y';
            if Cell[145,UV_iCount] <> '*' then sCheck := 'Y';
         end
         else sCheck := 'Y';

         UV_iCount := UV_iCount + 1;
      until (sCheck = 'Y');

      UV_iCount := UV_iCount - 1;

      QRL_Jisa.Caption  := Cell[ 2,UV_iCount];
      QRL_Bunju.Caption := Cell[ 3,UV_iCount];
            if Length(Cell[ 4,UV_iCount]) >= 8 then
         QRL_Name.Caption  :=  copy(Cell[ 4,UV_iCount],1,8)
      else
         QRL_Name.Caption  := Cell[ 4,UV_iCount];
      QRL_samp.Caption  := '(' + Cell[147,UV_iCount] + ')';

      UV_iCount := UV_iCount + 1;
      if UV_iCount <= Rows then MoreData := True
      else MoreData := False;
   end;
end;

procedure TfrmLD2Q044.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   UV_iCount := 1;
end;

end.
