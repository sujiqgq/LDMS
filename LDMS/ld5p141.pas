//==============================================================================
// 프로그램 설명 : 건강검진 검체검사 의뢰서
// 시스템        : LDMS
// 수정일자      : 2016.4.15
// 수정자        : 박대용
//==============================================================================
unit LD5P141;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Qrctrls, jpeg, QuickRpt, ExtCtrls, Db, DBTables, StdCtrls;

  const GawngJu_Code = 'SE42SEA6SE92SE93SE96';

type
  TfrmLD5P141 = class(TForm)
    qryGumgin1: TQuery;
    qryHm: TQuery;
    QuickRep: TQuickRep;
    qrySawon_sign: TQuery;
    QRBand1: TQRBand;
    QRImage1: TQRImage;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRShape1: TQRShape;
    QRShape5: TQRShape;
    QRShape2: TQRShape;
    QRShape6: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape7: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRLabel1: TQRLabel;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRShape36: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRShape42: TQRShape;
    QRShape43: TQRShape;
    QRShape44: TQRShape;
    QRShape45: TQRShape;
    QRShape46: TQRShape;
    QRShape47: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRL_M: TQRLabel;
    QRL_H: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
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
    QRShape30: TQRShape;
    QRShape48: TQRShape;
    QRShape49: TQRShape;
    QRShape50: TQRShape;
    QRShape51: TQRShape;
    QRShape52: TQRShape;
    QRShape53: TQRShape;
    QRShape54: TQRShape;
    QRShape56: TQRShape;
    QRShape57: TQRShape;
    QRShape58: TQRShape;
    QRShape59: TQRShape;
    QRShape60: TQRShape;
    QRShape61: TQRShape;
    QRShape62: TQRShape;
    QRShape63: TQRShape;
    QRShape64: TQRShape;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel49: TQRLabel;
    QRLabel50: TQRLabel;
    QRLabel51: TQRLabel;
    QRI_CUY: TQRImage;
    QRI_CJM: TQRImage;
    QRLabel17: TQRLabel;
    QRLabel52: TQRLabel;
    QRLabel53: TQRLabel;
    gumche: TQRShape;
    qryGumgin: TQuery;
    QRL_Jumin: TQRLabel;
    QRL_Dat_gmgn_Y: TQRLabel;
    QRL_Dat_gmgn_M: TQRLabel;
    QRL_Dat_gmgn_D: TQRLabel;
    QRL_Dat_gmgn_H: TQRLabel;
    QRL_Dat_gmgn_Mi: TQRLabel;
    QRL_Dat_gmgn: TQRLabel;
    QRL_Name: TQRLabel;
    QRL_Numbohm: TQRLabel;
    qryNhic: TQuery;
    QRL_NumSamp1: TQRLabel;
    QRL_NumSamp2: TQRLabel;
    QRShape55: TQRShape;
    QRL_Chk1: TQRLabel;
    QRL_Chk4: TQRLabel;
    QRL_Chk3: TQRLabel;
    QRL_Chk2: TQRLabel;
    qryProfile: TQuery;
    QRL_Kikwan: TQRLabel;
    QRL_Telnum: TQRLabel;
    QRL_doctor: TQRLabel;
    QRL_Kiho: TQRLabel;
    QRL_Ingae: TQRLabel;
    QRL_Insu: TQRLabel;
    QRI_JGS: TQRImage;
    QRI_SJG: TQRImage;
    QRI_AHG: TQRImage;
    QRI_PHL: TQRImage;
    QRI_KMHa: TQRImage;
    QRI_KMH: TQRImage;
    QRI_JSH: TQRImage;
    QRI_YYJ: TQRImage;
    QRI_LMJ: TQRImage;
    QRI_SJH: TQRImage;
    QRI_NUB: TQRImage;
    QRI_PYU: TQRImage;
    QRI_LUJ: TQRImage;
    QRI_LCS: TQRImage;
    QRI_LJH: TQRImage;
    QRI_kym: TQRImage;
    QRI_CMJ: TQRImage;
    QRI_CGA: TQRImage;
    QRI_BMY: TQRImage;
    QRI_MJH: TQRImage;
    QRI_CMS: TQRImage;
    QRI_YWH: TQRImage;
    QRI_KHY: TQRImage;
    QRI_KAR: TQRImage;
    QRI_GYM: TQRImage;
    QRI_KYH: TQRImage;
    QRI_SSY: TQRImage;
    QRI_KYJ: TQRImage;
    QRL_Chk5: TQRLabel;
    QRI_LEEMJ: TQRImage;
    QRI_JSN: TQRImage;
    QRI_HSN: TQRImage;
    QRI_KHW: TQRImage;
    QRI_LKL: TQRImage;
    QRI_KGH: TQRImage;
    QRI_LMY: TQRImage;
    QRI_HSL: TQRImage;
    QRI_JSM: TQRImage;
    QRI_HMY: TQRImage;
    QRI_LSJ: TQRImage;
    QRI_PHY: TQRImage;
    QRI_CYJ: TQRImage;
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRepNeedData(Sender: TObject;
      var MoreData: Boolean);
    procedure UP_ClearImage(Sender: TObject);
  private
     UV_sUrine_Char, UV_sDat_gmgn : string;
     UV_iCount : Integer;
     UV_vCod_hm : Variant;

     procedure UP_Clear;

     function UF_Check(Value : String): String;

    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD5P141: TfrmLD5P141;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3 : String;
  UV_iIndex : Integer ;

implementation

uses ZZComm, ZZfunc, MdmsFunc, LD5P14;

{$R *.DFM}

procedure TfrmLD5P141.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
  inherited;
  UV_iCount := 1;
  qryGumgin.First;
end;

//procedure TfrmHE5P011.UP_Clear;
procedure TfrmLD5P141.UP_Clear;
begin
   QRL_Jumin.Caption        := '';
   QRL_Numbohm.Caption      := '';
   QRL_Name.Caption         := '';
   QRL_Dat_gmgn.Caption     := '';
   QRL_Dat_gmgn_Y.Caption   := '';
   QRL_Dat_gmgn_M.Caption   := '';
   QRL_Dat_gmgn_D.Caption   := '';
   QRL_Dat_gmgn_H.Caption   := '';
   QRL_Dat_gmgn_Mi.Caption  := '';
   QRL_NumSamp1.Caption     := '';
   QRL_NumSamp2.Caption     := '';
   QRL_Kikwan.Caption       := '';
   QRL_Kiho.Caption         := '';
   QRL_doctor.Caption       := '';
   QRL_Telnum.Caption       := '';
   QRL_Ingae.Caption        := '';
   QRL_Insu.Caption         := '';
   QRL_Chk1.Caption         := '';
   QRL_Chk2.Caption         := '';
   QRL_Chk3.Caption         := '';
   QRL_Chk4.Caption         := '';
   QRL_Chk5.Caption         := '';

end;
 procedure TfrmLD5P141.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
Begin
  With qryGumgin Do
  Begin
    If Eof Then
    Begin
      MoreData := False;
      Exit;
    End;

    If UV_iCount = 1 Then
    Begin
      MoreData := True;
      Inc(UV_iCount);
    End
    else If UV_iCount <= RecordCount Then
    Begin
      MoreData := True;
      Inc(UV_iCount);
      Next;
    End
    Else
      Moredata := False;
   end;
end;

function  TfrmLD5P141.UF_Check(Value : String): String;
begin
   Result := '';

   if Value <> '0' then Result := Value;
end;

procedure TfrmLD5P141.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
   iRet, i, iAge, iMatCnt, iRow, iTemp, iCnt : Integer;
   vCod_chuga : Variant;
   sHmCode, sName, sMan, sGubn, sTemp,
   sCode1, sCode2, sCode3, sCode4,
   sSelect, sWhere, sOrderby, sHangmok_Iist, sTime : string;

   //이미지 관련.
   Jpeg  : TJPEGImage;
   MS    : TMemoryStream;

begin
   UP_Clear;
   sTime := '';
   {

   UV_vCod_hm := VarArrayCreate([0, 0], varOleStr);

   //항목불러오기...
   if not qryHm.Active then qryHm.Active := True;
     }
   with qryGumgin do
   begin
      QRL_Jumin.Caption         :=  copy(FieldByName('num_jumin').AsString,1,6) +' - ' + copy(FieldByName('num_jumin').AsString,7,7);
      QRL_Dat_gmgn.Caption      :=  copy(FieldByName('dat_gmgn').AsString,1,4) + '. ' + copy(FieldByName('dat_gmgn').AsString,5,2) + '. ' + copy(FieldByName('dat_gmgn').AsString,7,2) +'. ';
      QRL_Dat_gmgn_Y.Caption    :=  copy(FieldByName('Today_insert').AsString,1,4);
      QRL_Dat_gmgn_M.Caption    :=  copy(FieldByName('Today_insert').AsString,5,2);
      QRL_Dat_gmgn_D.Caption    :=  copy(FieldByName('Today_insert').AsString,7,2);

      if FieldByName('dat_gmgn').AsString > '20131231' then
      begin
           QRL_Dat_gmgn_Y.Caption    :=  copy(FieldByName('Today_insert').AsString,1,4);
           QRL_Dat_gmgn_M.Caption    :=  copy(FieldByName('Today_insert').AsString,5,2);
           QRL_Dat_gmgn_D.Caption    :=  copy(FieldByName('Today_insert').AsString,7,2);
           sTime := copy(FieldByName('Today_insert').AsString,15,6);
           QRL_Dat_gmgn_H.Caption := copy(sTime,1,pos(':', sTime)-1);
           if copy(FieldByName('Today_insert').AsString,12,2) = '후' then
              QRL_Dat_gmgn_H.Caption :=  IntToStr(Strtoint(QRL_Dat_gmgn_H.Caption)+12);
           sTime   := copy(sTime,pos(':', sTime)+1,3 );
           QRL_Dat_gmgn_Mi.Caption := copy(sTime,1,pos(':', sTime)-1);
           QRL_H.Enabled := True;
           QRL_M.Enabled := True;
      end
      else
      begin
           QRL_Dat_gmgn_Y.Caption    := copy(FieldByName('dat_gmgn').AsString,1,4);
           QRL_Dat_gmgn_M.Caption    := copy(FieldByName('dat_gmgn').AsString,5,2);
           QRL_Dat_gmgn_D.Caption    := copy(FieldByName('dat_gmgn').AsString,7,2);
           QRL_H.Enabled := False;
           QRL_M.Enabled := False;
      end;
      QRL_Name.Caption          := FieldByName('desc_name').AsString;
      QRL_NumSamp1.Caption      := FieldByName('num_samp').AsString;
      QRL_NumSamp2.Caption      := FieldByName('num_samp').AsString;
      sHangmok_Iist             := FieldByName('cod_chuga').AsString;

      UP_ClearImage(self);

      if FieldByName('cod_jisa').AsString  = '11' then
      begin
           QRL_Kikwan.Caption   := '한국의학연구소 강남의원';
           QRL_Kiho.Caption     := '11366338';
           QRL_doctor.Caption   := '정우진';
           QRL_Telnum.Caption   := '02-3496-3253';

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '지경선';
                QRI_JGS.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '지경선';
                QRI_JGS.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '배민영';
                QRI_BMY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '배민영';
                QRI_BMY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '최근아';
                QRI_CGA.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_Ingae.Caption    := '최민정';
                QRI_CMJ.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301')  and
                   (FieldByName('dat_gmgn').AsString <= '20180401') then
           begin
                QRL_Ingae.Caption    := '김혜원';
                QRI_KHW.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170401') then
           begin
                QRL_Ingae.Caption    := '이수지';
                QRI_LSJ.Enabled      := True;
           end;
      end
      else if FieldByName('cod_jisa').AsString  = '12' then
      begin
           if FieldByName('dat_gmgn').AsString <= '20170627' then
           begin
              QRL_Kikwan.Caption   := '케이엠아이 여의도의원';
              QRL_Kiho.Caption     := '11359994';
              QRL_doctor.Caption   := '주택소';
              QRL_Telnum.Caption   := '02-368-8206';
           end
           else
           begin
              QRL_Kikwan.Caption   := '케이엠아이 여의도의원';
              QRL_Kiho.Caption     := '11359994';
              QRL_doctor.Caption   := '이광설';
              QRL_Telnum.Caption   := '02-368-8206';
           end;

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '김윤미';
                QRI_kym.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '이미정';
                QRI_LMJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '김윤미';
                QRI_kym.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '이미정';
                QRI_LMJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '이지현';
                QRI_LJH.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString <= '20180311')  then
           begin
                QRL_Ingae.Caption    := '정소나';
                QRI_JSN.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20180312')  then
           begin
                QRL_Ingae.Caption    := '한미영';
                QRI_HMY.Enabled      := True;
           end
      end
      else if FieldByName('cod_jisa').AsString  = '43' then
      begin
           QRL_Kikwan.Caption   := '(재)KMI 경기의원';
           QRL_Kiho.Caption     := '31331254';
           if FieldByName('dat_gmgn').AsString >= '20160901' then
              QRL_doctor.Caption   := '김성훈'
           else
              QRL_doctor.Caption   := '허충';
           QRL_Telnum.Caption   := '031-231-0251';

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '이청심';
                QRI_LCS.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '이은정';
                QRI_LUJ.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '박영언';
                QRI_PYU.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '이청심';
                QRI_LCS.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '노은빈';
                QRI_NUB.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_Ingae.Caption    := '서진희';
                QRI_SJH.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') and
                   (FieldByName('dat_gmgn').AsString <= '20180409') then
           begin
                QRL_Ingae.Caption    := '이갑례';
                QRI_LKL.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString <= '20180410') and
                   (FieldByName('dat_gmgn').AsString >= '20190308')  then
           begin
                QRL_Ingae.Caption    := '박혜영';
                QRI_PHY.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20190309')  then
           begin
                QRL_Ingae.Caption    := '최예지';
                QRI_CYJ.Enabled      := True;
           end;
      end
      else if FieldByName('cod_jisa').AsString  = '51' then
      begin
           QRL_Kikwan.Caption   := '한국의학연구소 광주의원 광주분사무소';
           QRL_Kiho.Caption     := '36332712';
           QRL_doctor.Caption   := '김연선';
           QRL_Telnum.Caption   := '062-602-2143';

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '허새나';
                QRI_HSN.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '허새나';
                QRI_HSN.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '정상현';
                QRI_JSH.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '김명화';
                QRI_KMH.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '허새나';
                QRI_HSN.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2016' then
           begin
                QRL_Ingae.Caption    := '허새나';
                QRI_HSN.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_Ingae.Caption    := '윤애진';
                QRI_YYJ.Enabled      := True;          //허새나 -> 윤애진으로 변경
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') then
           begin
                QRL_Ingae.Caption    := '한송이';
                QRI_HSL.Enabled      := True;
           end;
      end
      else if FieldByName('cod_jisa').AsString  = '61' then
      begin
           QRL_Kikwan.Caption   := '재단법인 한국의학연구소 부산의원';
           QRL_Kiho.Caption     := '21343233';
           if copy(FieldByName('dat_gmgn').AsString,1,6) < '201512' then
               QRL_doctor.Caption   := '김성희'
           else
               QRL_doctor.Caption   := '이창희';
           QRL_Telnum.Caption   := '051-810-1500';

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '김민하';
                QRI_KMHa.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '박혜림';
                QRI_PHL.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '안희경';
                QRI_AHG.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '송진경';
                QRI_SJG.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '송수영';
                QRI_SSY.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228')  then
           begin
                QRL_Ingae.Caption    := '김영하';
                QRI_KYH.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') then
           begin
                QRL_Ingae.Caption    := '이미영';
                QRI_LMY.Enabled      := True;
           end;
      end
      else if FieldByName('cod_jisa').AsString  = '71' then
      begin
           QRL_Kikwan.Caption   := '재) 한국의학연구소     대구분사무소 부설의원';
           QRL_Kiho.Caption     := '37315676';
           QRL_doctor.Caption   := '서준원';
           QRL_Telnum.Caption   := '053-430-5000';

           if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Ingae.Caption    := '강유미';
                QRI_GYM.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Ingae.Caption    := '김아람';
                QRI_KAR.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Ingae.Caption    := '김희영';
                QRI_KHY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Ingae.Caption    := '김희영';
                QRI_KHY.Enabled      := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Ingae.Caption    := '유원화';
                QRI_YWH.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20160101') and
                   (FieldByName('dat_gmgn').AsString <= '20170228') then
           begin
                QRL_Ingae.Caption    := '유원화';
                QRI_YWH.Enabled      := True;
           end
           else if (FieldByName('dat_gmgn').AsString >= '20170301') then
           begin
                QRL_Ingae.Caption    := '김지향';
                QRI_KGH.Enabled      := True;
           end;
      end;

      if copy(FieldByName('dat_gmgn').AsString,1,4) = '2011' then
           begin
                QRL_Insu.Caption    := '김용진';
                QRI_KYJ.Enabled     := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2012' then
           begin
                QRL_Insu.Caption    := '최민선'; 
                QRI_CMS.Enabled     := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2013' then
           begin
                QRL_Insu.Caption    := '문지혜';
                QRI_MJH.Enabled     := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2014' then
           begin
                QRL_Insu.Caption    := '최정미';
                QRI_CJM.Enabled     := True;
           end
           else if copy(FieldByName('dat_gmgn').AsString,1,4) = '2015' then
           begin
                QRL_Insu.Caption    := '최정미';
                QRI_CJM.Enabled     := True;
           end
           else if (copy(FieldByName('dat_gmgn').AsString,1,4) = '2016') or
                   (copy(FieldByName('dat_gmgn').AsString,1,4) = '2017') or
                   (FieldByName('dat_gmgn').AsString <= '20180304') then
           begin
                QRL_Insu.Caption    := '최은영';
                QRI_CUY.Enabled     := True;
           end
           else if (copy(FieldByName('dat_gmgn').AsString,1,4) = '2018') then
           begin
                QRL_Insu.Caption    := '정선미';
                QRI_JSM.Enabled     := True;
           end;


      with qryProfile do
      begin
         Active := False;
         ParamByName('cod_pf').AsString    := qryGumgin.FieldByName('cod_hul').AsString;
         Active := True;
         if  0 < RecordCount then
         begin
             while not qryprofile.eof do
             begin
                sHangmok_Iist := sHangmok_Iist +  FieldByName('cod_hm').AsString;
                next;
             end;
         end;
         Active := False;
      end;
      
      if  (qryGumgin.FieldByName('gubn_life').AsString = '1') and
          (qryGumgin.FieldByName('gubn_lfyh').AsString = '1') then
      begin
          if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
             (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') then
              sHangmok_Iist := sHangmok_Iist + 'S099S091'
          else
              sHangmok_Iist := sHangmok_Iist + 'S007S008'
      end;

      with qryNhic do
      begin
          Active := False;
          ParamByName('num_jumin').AsString := qryGumgin.FieldByName('num_jumin').AsString;
          ParamByName('cod_jisa').AsString  := qryGumgin.FieldByName('cod_jisa').AsString;
          Active := True;
          QRL_Numbohm.Caption      := FieldByName('num_bohm').AsString;
          Active := False;
      end;
      if  (qryGumgin.FieldByName('gubn_life').AsString = '1') and
          (qryGumgin.FieldByName('gubn_lfyh').AsString = '1')  then
      begin
           if (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2011') and
              (copy(qryGumgin.FieldByName('dat_gmgn').AsString,1,4) <> '2012') then
           begin
                if (pos('S099',sHangmok_Iist) > 0) then  //B형간염표면항원
                begin
                     QRL_Chk1.Caption  := 'V';
                end
                else
                begin
                     QRL_Chk1.Caption  := '';
                end;

                if (pos('S091',sHangmok_Iist) > 0) then  //B형간염표면항체
                begin
                     QRL_Chk2.Caption  := 'V';
                end
                else
                begin
                     QRL_Chk2.Caption  := '';
                end;
           end
           else
           begin
                if (pos('S007',sHangmok_Iist) > 0) then  //B형간염표면항원
                begin
                     QRL_Chk1.Caption  := 'V';
                end
                else
                begin
                     QRL_Chk1.Caption  := '';
                end;

                if (pos('S008',sHangmok_Iist) > 0) then  //B형간염표면항체
                begin
                     QRL_Chk2.Caption  := 'V';
                end
                else
                begin
                     QRL_Chk2.Caption  := '';
                end;
           end
      end;

      if (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0) and
         (qryGumgin.FieldByName('dat_gmgn').AsString < '20170101') and
         (pos('S016',sHangmok_Iist) > 0) then  //C형간염표면항체
      begin
           QRL_Chk3.Caption  := 'V';
      end
      else if (qryGumgin.FieldByName('dat_gmgn').AsString >= '20170101') and
              ((qryGumgin.FieldByName('gubn_lfyh').AsString = '4') or
               (qryGumgin.FieldByName('gubn_lfyh').AsString = '5') or
               (qryGumgin.FieldByName('gubn_lfyh').AsString = '6')) then
      begin   
           QRL_Chk3.Caption  := 'V';
      end
      else
      begin
           QRL_Chk3.Caption  := '';
      end;

      if (pos('간',qryGumgin.FieldByName('desc_yhname').AsString) > 0)  then  //혈청알파태아단백검사
      begin
           if ((pos('TT03',sHangmok_Iist) > 0) and (qryGumgin.FieldByName('dat_gmgn').AsString > '20160101')) or
              ((pos('TT02',sHangmok_Iist) > 0) and (qryGumgin.FieldByName('dat_gmgn').AsString < '20160101')) then
                    QRL_Chk4.Caption  := 'V'
           else
                    QRL_Chk4.Caption  := '';
      end
      else
      begin
           QRL_Chk4.Caption  := '';
      end;

      if (pos('Y004',sHangmok_Iist) > 0)and
         (pos('대장',qryGumgin.FieldByName('desc_yhname').AsString) > 0) and
         (qryGumgin.FieldByName('dat_gmgn').AsString < '20160101')  then  //분변잠혈검사
      begin
           QRL_Chk5.Caption  := 'V';
      end
      else
      begin
           QRL_Chk5.Caption  := '';
      end;

      if (QRL_Chk1.Caption <> '') or
         (QRL_Chk2.Caption <> '') or
         (QRL_Chk3.Caption <> '') or
         (QRL_Chk4.Caption <> '') or
         (QRL_Chk5.Caption <> '') then
              PrintBand := True
      else
              PrintBand := False;
   end;

end;

procedure TfrmLD5P141.UP_ClearImage(Sender: TObject);
begin
  inherited;
     QRI_JGS.Enabled      := False;
     QRI_JGS.Enabled      := False;
     QRI_BMY.Enabled      := False;
     QRI_BMY.Enabled      := False;
     QRI_CGA.Enabled      := False;
     QRI_CMJ.Enabled      := False;
     QRI_kym.Enabled      := False;
     QRI_kym.Enabled      := False;
     QRI_LMJ.Enabled      := False;
     QRI_LJH.Enabled      := False;
     QRI_LJH.Enabled      := False;
     QRI_LCS.Enabled      := False;
     QRI_LUJ.Enabled      := False;
     QRI_PYU.Enabled      := False;
     QRI_LCS.Enabled      := False;
     QRI_NUB.Enabled      := False;
     QRI_SJH.Enabled      := False;
     QRI_HSN.Enabled      := False;
     QRI_YYJ.Enabled      := False;
     QRI_YYJ.Enabled      := False;
     QRI_JSH.Enabled      := False;
     QRI_KMH.Enabled      := False;
     QRI_YYJ.Enabled      := False;
     QRI_YYJ.Enabled      := False;
     QRI_KMHa.Enabled     := False;
     QRI_PHL.Enabled      := False;
     QRI_AHG.Enabled      := False;
     QRI_SJG.Enabled      := False;
     QRI_SSY.Enabled      := False;
     QRI_KYH.Enabled      := False;
     QRI_GYM.Enabled      := False;
     QRI_KAR.Enabled      := False;
     QRI_KHY.Enabled      := False;
     QRI_KHY.Enabled      := False;
     QRI_YWH.Enabled      := False;
     QRI_YWH.Enabled      := False;
     QRI_KYJ.Enabled      := False;
     QRI_MJH.Enabled      := False;
     QRI_CJM.Enabled      := False;
     QRI_CJM.Enabled      := False;
     QRI_CUY.Enabled      := False; 
     QRI_CMS.Enabled      := False;
     QRI_JSN.Enabled      := False;
     QRI_KHW.Enabled      := False;
     QRI_LKL.Enabled      := False;
     QRI_KGH.Enabled      := False;
     QRI_LMY.Enabled      := False; 
     QRI_HSL.Enabled      := False;
end;

end.
