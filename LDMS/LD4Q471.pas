unit LD4Q471;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q471 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel5: TQRLabel;
    Qrl_Date: TQRLabel;
    QRBand3: TQRBand;
    QRLabel9: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel_101: TQRLabel;
    QRLabel_102: TQRLabel;
    QRLabel_103: TQRLabel;
    QRLabel_104: TQRLabel;
    QRLabel_105: TQRLabel;
    QRLabel_106: TQRLabel;
    QRLabel_107: TQRLabel;
    QRLabel_108: TQRLabel;
    QRLabel_109: TQRLabel;
    QRLabel_110: TQRLabel;
    QRLabel_111: TQRLabel;
    QRLabel_112: TQRLabel;
    QRLabel_113: TQRLabel;
    QRLabel_114: TQRLabel;
    QRLabel_115: TQRLabel;
    QRLabel_116: TQRLabel;
    QRLabel_201: TQRLabel;
    QRLabel_202: TQRLabel;
    QRLabel_203: TQRLabel;
    QRLabel_204: TQRLabel;
    QRLabel_205: TQRLabel;
    QRLabel_206: TQRLabel;
    QRLabel_207: TQRLabel;
    QRLabel_208: TQRLabel;
    QRLabel_209: TQRLabel;
    QRLabel_210: TQRLabel;
    QRLabel_211: TQRLabel;
    QRLabel_212: TQRLabel;
    QRLabel_213: TQRLabel;
    QRLabel_214: TQRLabel;
    QRLabel_215: TQRLabel;
    QRLabel_216: TQRLabel;
    QRLabel_301: TQRLabel;
    QRLabel_302: TQRLabel;
    QRLabel_303: TQRLabel;
    QRLabel_304: TQRLabel;
    QRLabel_307: TQRLabel;
    QRLabel_308: TQRLabel;
    QRLabel_309: TQRLabel;
    QRLabel_310: TQRLabel;
    QRLabel_311: TQRLabel;
    QRLabel_314: TQRLabel;
    QRLabel_315: TQRLabel;
    QRLabel_312: TQRLabel;
    QRLabel_313: TQRLabel;
    QRLabel_316: TQRLabel;
    QRLabel_305: TQRLabel;
    QRLabel_306: TQRLabel;
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
  frmLD4Q471: TfrmLD4Q471;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q47;

{$R *.DFM}

procedure TfrmLD4Q471.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i: integer;
begin
   inherited;
   if frmLD4Q47.Chk_special.Checked =True then
      begin
           QRLabel_101.Caption:= 'APOA';
           QRLabel_102.Caption:= 'APOB';
           QRLabel_103.Caption:= 'B/A';
           QRLabel_104.Caption:= 'Tich';
           QRLabel_105.Caption:= 'HDL';
           QRLabel_106.Caption:= 'LDL';
           QRLabel_107.Caption:= 'TG';
           QRLabel_108.Caption:= 'CRI';
           QRLabel_109.Caption:= 'Bun';
           QRLabel_110.Caption:= 'Crea';
           QRLabel_111.Caption:= '/C';
           QRLabel_112.Caption:= 'FE';
           QRLabel_113.Caption:= 'TIBC';
           QRLabel_114.Caption:= 'TIBCS';
           QRLabel_115.Caption:= 'UIBC';
           QRLabel_116.Caption:= 'Ca';

           QRLabel_301.Caption:= '[C022]';
           QRLabel_302.Caption:= '[C023]';
           QRLabel_303.Caption:= '[C024]';
           QRLabel_304.Caption:= '[C025]';
           QRLabel_305.Caption:= '[C026]';
           QRLabel_306.Caption:= '[C027]';
           QRLabel_307.Caption:= '[C028]';
           QRLabel_308.Caption:= '[C029]';
           QRLabel_309.Caption:= '[C041]';
           QRLabel_310.Caption:= '[C042]';
           QRLabel_311.Caption:= '[C043]';
           QRLabel_312.Caption:= '[C045]';
           QRLabel_313.Caption:= '[C046]';
           QRLabel_314.Caption:= '[C047]';
           QRLabel_315.Caption:= '[C048]';
           QRLabel_316.Caption:= '[C056]';
      end
      else if frmLD4Q47.Chk_part_01.Checked =True then
      begin
           QRLabel_101.Caption:= 'TP';
           QRLabel_102.Caption:= 'Alb';
           QRLabel_103.Caption:= 'Glo';
           QRLabel_104.Caption:= 'A/G';
           QRLabel_105.Caption:= 'TB';
           QRLabel_106.Caption:= 'DB';
           QRLabel_107.Caption:= 'IB';
           QRLabel_108.Caption:= 'AST';
           QRLabel_109.Caption:= 'ALT';
           QRLabel_110.Caption:= 'GTP';
           QRLabel_111.Caption:= 'LAP';
           QRLabel_112.Caption:= 'ALP';
           QRLabel_113.Caption:= 'UA';
           QRLabel_114.Caption:= 'CPK';
           QRLabel_115.Caption:= 'LDH';
           QRLabel_116.Caption:= 'B-Lipo';

           QRLabel_301.Caption:= '[C001]';
           QRLabel_302.Caption:= '[C002]';
           QRLabel_303.Caption:= '[C003]';
           QRLabel_304.Caption:= '[C004]';
           QRLabel_305.Caption:= '[C005]';
           QRLabel_306.Caption:= '[C006]';
           QRLabel_307.Caption:= '[C007]';
           QRLabel_308.Caption:= '[C009]';
           QRLabel_309.Caption:= '[C010]';
           QRLabel_310.Caption:= '[C011]';
           QRLabel_311.Caption:= '[C012]';
           QRLabel_312.Caption:= '[C013]';
           QRLabel_313.Caption:= '[C017]';
           QRLabel_314.Caption:= '[C019]';
           QRLabel_315.Caption:= '[C021]';
           QRLabel_316.Caption:= '[C030]';
      end
      else if frmLD4Q47.Chk_part_02.Checked =True then
      begin
           QRLabel_101.Caption:= 'Glu';
           QRLabel_102.Caption:= 'PP2';
           QRLabel_103.Caption:= 'FA';
           QRLabel_104.Caption:= 'AMY';
           QRLabel_105.Caption:= 'I.P';
           QRLabel_106.Caption:= 'Mg';
           QRLabel_107.Caption:= 'Homo';
           QRLabel_108.Caption:= 'Lp(a)';
           QRLabel_109.Caption:= 'Lip';
           QRLabel_110.Caption:= 'A1C';
           QRLabel_111.Caption:= 'LDL-C';
           QRLabel_112.Caption:= 'Na';
           QRLabel_113.Caption:= 'K ';
           QRLabel_114.Caption:= 'CI ';
           QRLabel_115.Caption:= '';
           QRLabel_116.Caption:= '';

           QRLabel_301.Caption:= '[C032]';
           QRLabel_302.Caption:= '[C033]';
           QRLabel_303.Caption:= '[C034]';
           QRLabel_304.Caption:= '[C039]';
           QRLabel_305.Caption:= '[C057]';
           QRLabel_306.Caption:= '[C058]';
           QRLabel_307.Caption:= '[C078]';
           QRLabel_308.Caption:= '[C080]';
           QRLabel_309.Caption:= '[C082]';
           QRLabel_310.Caption:= '[C083]';
           QRLabel_311.Caption:= '[C074]';
           QRLabel_312.Caption:= '[C050]';
           QRLabel_313.Caption:= '[C051]';
           QRLabel_314.Caption:= '[C052]';
           QRLabel_315.Caption:= '';
           QRLabel_316.Caption:= '';
      end;



   with frmLD4Q47.grdMaster do
   begin
      // 자료가 없을경우의 처리
      if Rows = 0 then
      begin
         MoreData := False;
         exit;
      end;
      Inc(UV_iCount);



      QRL_NO.caption     := Cell[1,UV_iCount];
      QRL_SAMPLE.Caption := Cell[2,UV_iCount];
      QRL_NAME.Caption   := Cell[4,UV_iCount];
      QRL_BUNJU.Caption  := Cell[3,UV_iCount];

      QRLabel_201.Caption   := Cell[5,UV_iCount];
      QRLabel_202.Caption   := Cell[6,UV_iCount];
      QRLabel_203.Caption   := Cell[7,UV_iCount];
      QRLabel_204.Caption   := Cell[8,UV_iCount];
      QRLabel_205.Caption   := Cell[9,UV_iCount];
      QRLabel_206.Caption   := Cell[10,UV_iCount];
      QRLabel_207.Caption   := Cell[11,UV_iCount];
      QRLabel_208.Caption   := Cell[12,UV_iCount];
      QRLabel_209.Caption   := Cell[13,UV_iCount];
      QRLabel_210.Caption   := Cell[14,UV_iCount];
      QRLabel_211.Caption   := Cell[15,UV_iCount];
      QRLabel_212.Caption   := Cell[16,UV_iCount];
      QRLabel_213.Caption   := Cell[17,UV_iCount];
      QRLabel_214.Caption   := Cell[18,UV_iCount];
      QRLabel_215.Caption   := Cell[19,UV_iCount];
      QRLabel_216.Caption   := Cell[20,UV_iCount];

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q471.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q47.edtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q47.edtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q47.edtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
