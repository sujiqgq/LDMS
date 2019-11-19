unit LD4Q361;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, QuickRpt, ExtCtrls;

type
  TfrmLD4Q361 = class(TForm)
    QuickRep: TQuickRep;
    QRBand2: TQRBand;
    QRL_NO: TQRLabel;
    QRL_BUNJU: TQRLabel;
    QRL_NAME: TQRLabel;
    QRL_SAMPLE: TQRLabel;
    QRShape2: TQRShape;
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
    QRShape1: TQRShape;
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
    QRLabel_117: TQRLabel;
    QRLabel_118: TQRLabel;
    QRLabel_119: TQRLabel;
    QRLabel_120: TQRLabel;
    QRLabel_121: TQRLabel;
    QRLabel_122: TQRLabel;
    QRLabel_123: TQRLabel;
    QRLabel_124: TQRLabel;
    QRLabel_125: TQRLabel;
    QRLabel_126: TQRLabel;
    QRLabel_127: TQRLabel;
    QRLabel_128: TQRLabel;
    QRLabel_129: TQRLabel;
    QRLabel_130: TQRLabel;
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
    QRLabel_217: TQRLabel;
    QRLabel_218: TQRLabel;
    QRLabel_219: TQRLabel;
    QRLabel_220: TQRLabel;
    QRLabel_221: TQRLabel;
    QRLabel_222: TQRLabel;
    QRLabel_223: TQRLabel;
    QRLabel_224: TQRLabel;
    QRLabel_225: TQRLabel;
    QRLabel_226: TQRLabel;
    QRLabel_227: TQRLabel;
    QRLabel_228: TQRLabel;
    QRLabel_229: TQRLabel;
    QRLabel_230: TQRLabel;
    QRLabel_131: TQRLabel;
    QRLabel_231: TQRLabel;
    QRLabel_132: TQRLabel;
    QRLabel_232: TQRLabel;
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
  frmLD4Q361: TfrmLD4Q361;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q36;

{$R *.DFM}

procedure TfrmLD4Q361.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
  var i: integer;
begin
   inherited;




   with frmLD4Q36.grdMaster do
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
      QRLabel_217.Caption   := Cell[21,UV_iCount];
      QRLabel_218.Caption   := Cell[22,UV_iCount];
      QRLabel_219.Caption   := Cell[23,UV_iCount];
      QRLabel_220.Caption   := Cell[24,UV_iCount];
      QRLabel_221.Caption   := Cell[25,UV_iCount];
      QRLabel_222.Caption   := Cell[26,UV_iCount];
      QRLabel_223.Caption   := Cell[27,UV_iCount];
      QRLabel_224.Caption   := Cell[28,UV_iCount];
      QRLabel_225.Caption   := Cell[29,UV_iCount];
      QRLabel_226.Caption   := Cell[30,UV_iCount];
      QRLabel_227.Caption   := Cell[31,UV_iCount];
      QRLabel_228.Caption   := Cell[32,UV_iCount];
      QRLabel_229.Caption   := Cell[33,UV_iCount];
      QRLabel_230.Caption   := Cell[34,UV_iCount];
      QRLabel_231.Caption   := Cell[35,UV_iCount];
      QRLabel_232.Caption   := Cell[36,UV_iCount];



      QRLabel_101.Caption   := frmLD4Q36.grdMaster.Col[5].heading;
      QRLabel_102.Caption   := frmLD4Q36.grdMaster.Col[6].heading;
      QRLabel_103.Caption   := frmLD4Q36.grdMaster.Col[7].heading;
      QRLabel_104.Caption   := frmLD4Q36.grdMaster.Col[8].heading;
      QRLabel_105.Caption   := frmLD4Q36.grdMaster.Col[9].heading;
      QRLabel_106.Caption   := frmLD4Q36.grdMaster.Col[10].heading;
      QRLabel_107.Caption   := frmLD4Q36.grdMaster.Col[11].heading;
      QRLabel_108.Caption   := frmLD4Q36.grdMaster.Col[12].heading;
      QRLabel_109.Caption   := frmLD4Q36.grdMaster.Col[13].heading;
      QRLabel_110.Caption   := frmLD4Q36.grdMaster.Col[14].heading;
      QRLabel_111.Caption   := frmLD4Q36.grdMaster.Col[15].heading;
      QRLabel_112.Caption   := frmLD4Q36.grdMaster.Col[16].heading;
      QRLabel_113.Caption   := frmLD4Q36.grdMaster.Col[17].heading;
      QRLabel_114.Caption   := frmLD4Q36.grdMaster.Col[18].heading;
      QRLabel_115.Caption   := frmLD4Q36.grdMaster.Col[19].heading;
      QRLabel_116.Caption   := frmLD4Q36.grdMaster.Col[20].heading;
      QRLabel_117.Caption   := frmLD4Q36.grdMaster.Col[21].heading;
      QRLabel_118.Caption   := frmLD4Q36.grdMaster.Col[22].heading;
      QRLabel_119.Caption   := frmLD4Q36.grdMaster.Col[23].heading;
      QRLabel_120.Caption   := frmLD4Q36.grdMaster.Col[24].heading;
      QRLabel_121.Caption   := frmLD4Q36.grdMaster.Col[25].heading;
      QRLabel_122.Caption   := frmLD4Q36.grdMaster.Col[26].heading;
      QRLabel_123.Caption   := frmLD4Q36.grdMaster.Col[27].heading;
      QRLabel_124.Caption   := frmLD4Q36.grdMaster.Col[28].heading;
      QRLabel_125.Caption   := frmLD4Q36.grdMaster.Col[29].heading;
      QRLabel_126.Caption   := frmLD4Q36.grdMaster.Col[30].heading;
      QRLabel_127.Caption   := frmLD4Q36.grdMaster.Col[31].heading;
      QRLabel_128.Caption   := frmLD4Q36.grdMaster.Col[32].heading;
      QRLabel_129.Caption   := frmLD4Q36.grdMaster.Col[33].heading;
      QRLabel_130.Caption   := frmLD4Q36.grdMaster.Col[34].heading;
      QRLabel_131.Caption   := frmLD4Q36.grdMaster.Col[35].heading;
      QRLabel_132.Caption   := frmLD4Q36.grdMaster.Col[36].heading;


      if frmLD4Q36.edtBDate.Text < '20140911' then
      begin
          if frmLD4Q36.Chk_ALL.Checked then //전체 - T039 안보이게
          begin
             QRLabel_121.Caption := '';
             QRLabel_221.Caption := '';
          end;
      end;

      if frmLD4Q36.edtBDate.Text >= '20140901' then
      begin
          if frmLD4Q36.CHK_sub.Checked then //외주항목 - T043 안보이게
          begin
             QRLabel_202.Caption := '';
             QRLabel_102.Caption := '';
          end
          else if frmLD4Q36.Chk_Arch.Checked then //스페셜 - T042 안보이게
          begin
             QRLabel_205.Caption := '';
             QRLabel_105.Caption := '';
          end
      end
      else
      begin
          if frmLD4Q36.Chk_ALL.Checked then //전체조회 - T043 안보이게
          begin
             QRLabel_221.Caption := '';
             QRLabel_121.Caption := '';
          end
          else if frmLD4Q36.Chk_C1_C5.Checked then //Routine - T042, T043 안보이게
          begin
             QRLabel_209.Caption := '';
             QRLabel_210.Caption := '';
             QRLabel_109.Caption := '';
             QRLabel_110.Caption := '';
          end;
      end;

      if UV_iCount <= Rows then
         MoreData := True
      else
         MoreData := False;
   end;

end;

procedure TfrmLD4Q361.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
   inherited;
   QRL_date.Caption := copy(frmLD4Q36.edtBDate.Text,1,4) + '년'  +
                       copy(frmLD4Q36.edtBDate.Text,5,2) + '월'  +
                       copy(frmLD4Q36.edtBDate.Text,7,2) + '일';
   UV_iCount := 0;
end;

end.
 