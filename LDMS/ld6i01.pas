//==============================================================================
// 프로그램 설명 : 자동소견작업
// 시스템        : 통합검진
// 수정일자      : 2006.03.14
// 수정자        : 유동구
// 수정내용      : 혈액검사 완료 체크...
// 참고사항      : [2006.01.25-CEA(T002)수정]
//                 [2006.02.22-TTT(C008)항목제외]                               
//                 [2006.02.22-케톤(U006)수치변경[20 ->10]]
//                 [2006.04.07-CEA(T002-82), AFP(TT01-88),CA19-9(T006-83),
//                             CA125(T007-84), N.S.E(T012-106),calcitonin(E016-107),HBcAbIgG(SE13-108)]
//                 [2006.05.02-혈액소견 HT1/HT2 제외]
//                 [2006.06.01-혈액소견 10개로 늘림]
//                 [2006.08.01-자동소견시 소견 입력된 검진자는 제외하고 자동소견됨]
//                 [2007.04.08-Aldolase(C015)소견추가]
//                 [2007.07.09-혈당120 -> 126추가]
//                              TSH(E003) 0.2 이하 - TSH2
//                              IP(C057)  5.5이상 - PH3
//                              Hb(H002)  여자  11.0-11.9  - IN06
//                                        남자 12.0-12.9   - IN06
//                 [2008.09.09-rubella IgG(음),   igM(음) - rig3
//                                     IgG(양),   igM(음) - rig2
//                                     IgG(약양), igM(음) - rig1
//                             여자일때 소견코드 UWB -> UWBC
//                                               UWN -> UWN1
//==============================================================================
// 수정일자      : 2006.05.08
// 수정자        : 김승철
// 수정내용      : 생애전환기 추가
//==============================================================================
// 수정일자      : 2009.07.04
// 수정자        : 김승철
// 수정내용      : SE01,SE02 추가
//==============================================================================
// 수정일자      : 2010.05.06
// 수정자        : 송재호
// 수정내용      : LDL콜레스테롤 기준치 50~160 -> 50~130 변경
//                중성지방 기준치 50~170 -> 50~150 변경
//                 HDL콜레스테롤 기준치 남자 37~58 -> 40~58, 여자 45~70 -> 50~70 변경
//                 LDL콜레스테롤(C027) 130이상 - LDLH
//                 HDL콜레스테롤(C026) 남자39이하 OR 여자49이하 - HDLL
//                 중성지방(C028) 150~199 - TH3
//                 중성지방(C028) 200~399 - TH
//                 중성지방(C028) 400이상 - TH5
//                 콜레스테롤(C025) 201이상, LDL콜레스테롤(C027) 50~129 - LDLN
//                 51세이상 여자, Fe(C045) 54이하, 철포화율(C047) 19이하, 혈색소(H002) 11.9이하 - IDPM
//                 51세이상 여자, Fe(C045) 54이하, 철포화율(C047) 19이하, 혈색소(H002) 12이상 16이하 - IDOM
//                 LDL콜레스테롤(C027)검사가 없을 경우에만 CH, CH1, CTL, CTH, CLH 소견이 나가도록..
//==============================================================================
// 수정일자      : 2010.05.14
// 수정자        : 유동구
// 수정내용      : Hul_sokun table : cod_pan 추가...
//==============================================================================
// 수정일자      : 2010.05.25
// 수정자        : 송재호
// 수정내용      : Fe(C045)관련 소견 적용범위를 남자 59이하, 여자 54이하로 변경.
//                 HDL콜레스테롤(C026) 100 이상일때 HDLH 추가.
//                 HDL콜레스테롤(C026) 정상미만 & cardiac risk index(C029) 검사안됐을때 HDLL 소견
//                 LAP 100 이상으로 수정
//                 혈당(C032) 100~125 & 당화단백(C034) 정상이상 & 당화혈색소(C083) 정상이상 일때 HBA1
//                 혈당(C032) 100~125 & 당화단백(C034) 정상이상 일때 DIC5
//                 혈당(C032) 100~125 & 당화혈색소(C083) 정상이상 일때 DIH5
//                 혈당(C032) 100~125 일때 DMN
//                 혈당(C032) 126~199 & 당화단백(C034) 정상이상 & 당화혈색소(C083) 정상이상 일때 HBA1
//                 혈당(C032) 126~199 & 당화단백(C034) 정상이상 일때 DIC
//                 혈당(C032) 126~199 & 당화혈색소(C083) 정상이상 일때 DIH
//                 혈당(C032) 126~199  일때 DIB
//                 혈당(C032) 200이상 & 당화단백(C034) 정상이상 일때 DIC4
//                 혈당(C032) 200이상 & 당화혈색소(C083) 정상이상 일때 DIH4
//                 혈당(C032) 200이상 일때 DIB4
//                 혈당(C032) 99이하 & 당화단백(C034) 정상이상 일때 DIO 
//                 혈당(C032) 99이하 & 당화혈색소(C083) 정상이상 일때 DIN
//                 혈당(C032) 59이하 일때 SL
//==============================================================================
// 수정일자      : 2010.06.30
// 수정자        : 송재호
// 수정내용      : 노신,생애,성인,간염 유형별 항목 함수로 끌고 오도록 수정(기존에는 항목 고정)
//                 특검항목도 끌고 오도록 추가
//==============================================================================
// 수정일자      : 2010.09.14
// 수정자        : 송재호
// 수정내용      : SE02(HAV lgG)가 음성일 경우 HAVB
//==============================================================================
// 수정일자      : 2010.12.22
// 수정자        : 송재호
// 수정내용      : nicotine 양성시 NCT
//==============================================================================
// 수정일자      : 2011.03.03
// 수정자        : 송재호
// 수정내용      : 기 존 : BET(C044 >= 2.6), BET1(2.1<=C044<=2.5)
//                변경후 : BET(C044 >= 6.1), BET1(3.1<=C044<=6.0)
//                 기 존 : THL (E002 <= 4.9), THH1
//(13.1 >= E002 <= 13.9), THH (E002 >= 14.0)
//                변경후 : THL (E002 <= 6.0), THH1 (11.9 >= E002 <= 13.0), THH (E002 >= 13.1)
//                 기 존 : CAD (T007 >= 37)
//                변경후 : CAD (T007 >= 35.1)
//==============================================================================
// 수정일자      : 2011.03.07
// 수정자        : 송재호
// 수정내용      : 기  존 : WBL1 (H011 >= 3.0) And (H011 <= 3.3)
//                 변경후 : WBL1 (H011 >= 3.1) And (H011 <= 3.9)
//                 기  존 : WBC (H011 >= 10.1)
//                 변경후 : WBC (H011 >= 11.1)
//                 추  가 : WBC1 (10.1 <=H011 >= 11.0)
//                 삭  제 : WBE (H011 >= 11.0 And H016 >= 8)
//==============================================================================
// 수정일자      : 2011.03.19
// 수정자        : 송재호
// 수정내용      : 기  존 : Cod = ANE (H002 <= 11.9) - 성인남자
//                 변경후 : Cod = ANE (H002 <= 12.9) - 성인남자
//                 기  존 : Cod = IN06(12.0 <= H002 <= 12.9) - 남자
//                 변경후 : Cod = IN06(13.0 <= H002 <= 13.9) - 남자
//                 기  존 : Cod = IDA (C045 <= 59 And H002 <= 12.0 남자) or (C045 <= 54 And H002 <= 11.7 여자) 51세이상 여자는 IDPM으로..
//                 변경후 : Cod = IDA (C045 <= 54 And H002 <= 11.7 여자) 51세이상 여자는 IDPM으로..
//                          Cod = IDA6(C045 <= 59 And H002 <= 13.9 남자)
//                 기  존 : Cod = WBL (H011 < 3.0)
//                 변경후 : Cod = WBL (H011 <= 3.0)
//                 기  존 : Cod = WBL1 (H011 >= 3.0) And (H011 <= 3.3)
//                 변경후 : Cod = WBL1 (H011 >= 3.1) And (H011 <= 3.3)
//==============================================================================
// 수정일자      : 2011.6.20
// 수정자        : 김희석
// 수정내용      : urin  참고치 변경
//==============================================================================
// 수정일자      : 2011.11.04
// 수정자        : 송재호
// 수정내용      : 소견코드 DIA는 사용안하고 126<=C032<=199 AND 1<=U005<=100일때 DIB 사용(한의 재단진단검사의학부서장1100009)
//                 소견코드 DID는 사용안하고 126<=C032<=199 And C034>=2.86 일때 DIC 사용(한의 재단진단검사의학부서장1100009 추가요청)
//==============================================================================
// 수정일자      : 2012.03.22
// 수정자        : 유동구
// 수정내용      : 기  존 : Cod = IN06(13.0 <= H002 <= 13.9) - 남자
//                 변경후 : Cod = IN06(13.0 <= H002 <= 13.4) - 남자
//==============================================================================
// 수정일자      : 2012.03.23
// 수정자        : 송재호
// 수정내용      : 기  존 : Cod = IDA6(13.9)
//                 변경후 : Cod = IDA6(13.4) -> 김소영 선생님 요청
//==============================================================================
// 수정일자      : 2012.03.28
// 수정자        : 송재호
// 수정내용      : 기  존 : 남자 - Cod = IDO (H002 >= 12.0)
//                          여자(50세미만) - Cod = IDO (H002 >= 11.8)
//                          여자(50세이상) - Cod = IDOM(H002 >= 11.8)
//                 변경후 : 남자 - Cod = IDO2(H002 >= 13.6)
//                          여자(50세미만) - Cod = IDO (H002 >= 12.1)
//                          여자(50세이상) - Cod = IDOM(H002 >= 12.1) -> 김소영 선생님 요청
//==============================================================================
// 수정일자      : 2012.06.27
// 수정자        : 김희석
// 수정내용      : HAV3혈액코드는 삭제.
//                 SE02 양성 이고 (C009 or C010)>40  일경우 HAVG
//                 SE02 양성 이고 (C009 or C010)<=40  일경우 HAVH
//==============================================================================
// 수정자        : 김희석
// 수정내용      :  추가
//                  TT01 양성이고  임신일경우 (접수)  TT02 > 15 일경우  AFG
//                  TT01 양성이고  TT02 > 15 일경우 AFP0
//                  C078 >15 이면  HOMO
//                  C003 >4.0 이면 GH
//                  H025 여자 >15 or 남자 >9  이면 ESR
//                  소견 HDLL 이면서 C025 <120 일경우  소견 코드 삭제 
//==============================================================================
// 수정일자      : 2012.07.23
// 수정자        : 김희석
// 수정내용      : H002 남자 =>18.1
//                      여자 =>16.1  일경우   HHB
//==============================================================================
// 수정일자      : 2012.08.14
// 수정자        : 김희석
// 수정내용      : C065 <149.9 ,C065 >232.1
//               : LDLN 0~129.9 
//==============================================================================
// 수정일자      : 2012.09.25
// 수정자        : 김희석
// 수정내용      : 한의 본원진료팀1200076
//==============================================================================
// 수정일자      : 2012.10.08
// 수정자        : 김희석
// 수정내용      : 한의 본원진료팀1200078
//==============================================================================
// 수정일자      : 20130409
// 수정자        : 김희석
// 수정내용      : 7000 이상 분주 일경우 u017 항목 소견만
//==============================================================================
//==============================================================================
// 수정일자      : 20140329
// 수정자        : 김창규
// 수정내용      : CA125는 여자만, PSA 는 남자만
// 수정내용      : CA125, T007 SKIP 제외
//==============================================================================
//==============================================================================
// 수정일자      : 20140408
// 수정자        : 김창규
// 수정내용      : S010=양성 에서 약양성 추가
// 수정내용      : 소견코드 : HBC8
//==============================================================================
//==============================================================================
// 수정일자      : 20140409
// 수정자        : 김창규
// 수정내용      : C083 5.7~6.4 And C032 <=99 --> DINH
// 수정내용      : DMN조건 + C083 5.7~6.4 ------> DMNH
//==============================================================================

//==============================================================================
// 수정일자      : 20140423
// 수정자        : 김창규
// 수정내용      : 단체소견Routine 추가(시립대,힐튼호텔)
//	C009		C010		         S007 or S099		S008 or S091
// HPD	<=60.0		45.1(34.1)~60.0		음성 or <=0.99		상관없음
//
// FL	60.1~100       	45.1(34.1)~60.0		음성 or <=0.99		상관없음
//	<=100		60.1~100.0		음성 or <=0.99		상관없음
//
// HPM	>=100.1		>=45.1(34.1)		음성 or <=0.99		상관없음
//	전범위		>=100.1		        음성 or <=0.99		상관없음
//
// AST	>=40.1(32.1)   	<=45.0(34.0)				        상관없음
//
// HBM	<=100.0		45.1(34.1)~100		양성 or >=1.0		음성
//
// HPB	전범위		>=100.1		        양성 or >=1.0		음성
// 	>=100.1		>=45.1(34.1)		양성 or >=1.0		음성
//
// HBC1	<=40.0(32.0)   	<=45.0(34.0)		양성 or >=1.0		음성
//
// HBC 	>=40.1(32.1)	<=45.0(34.0)		양성 or >=1.0		음성
// +AST
// -------- Cancer 정상 소견
// TT02 <=8.99 or TT01 음성 --> AFPD
// T002 <=5.0  --> CEAN
// T007 <=35.0 --> CA2N
// T006 <=37.0 --> CA9N
// T011 <=4.0  --> PSAN
// C044 <=3.0  --> BETN
// T037 <=28.0 --> CA5N
// E040 <=3.0  --> CYFN
// T005 <=6.9  --> CA7N
//
//==============================================================================

//==============================================================================
// 수정일자      : 20140425
// 수정자        : 김창규
// 수정내용      : 지질소견 변경
//	        C025(Tchol)     C027(LDLC)	C028(TG)
//      LDLL	                >=200.0		<=99.9
//      LDLQ	<=199.9		100~129.9
//      LDLN	>=200		100~129.9
//      LDLK			130~159.9
//      LDLO			160~189.9.
//      LDLP			>=190.0
//      THVH					>=1000.0
//      TH5					500.0~999.9
//      TH					150.0~499.9
//      TH3					<=149.9
//      ** 더불어 C025 참고치 오류가 발견되어 수정 원합니다. 종전 <=200을 <=199 로 수정해주세요.
//==============================================================================
// 수정일자      : 20140509
// 수정자        : 곽윤설
// 수정내용      : CH1 범위변경
//	        C025(Tchol)
//      CH1	>=240.0
//==============================================================================
//==============================================================================
// 수정일자      : 20140610
// 수정자        : 김창규
// 수정내용      : E056 소견 수정및 E049 소견 추가
//==============================================================================
//==============================================================================
// 수정일자      : 20140707
// 수정자        : 김창규
// 수정내용      : H015 >= 20.1 --> MONH
//==============================================================================
//==============================================================================
// 수정일자      : 20140722
// 수정자        : 김창규
// 수정내용      : Y004 = '음성' --> OBN
//==============================================================================
// 수정일자      : 20140911
// 수정자        : 곽윤설
// 수정내용      : 당뇨병 진단 또는 치료일 경우 DINH 소견 미적용
//==============================================================================
// 수정일자      : 20141219
// 수정자        : 곽윤설
// 수정내용      : C039 참고치 변경(125.01 -> 146.1), 분주일자 20141216부터 적용되도록 (관련 소견코드 : AMH, AMHR)
//==============================================================================
// 수정일자      : 20141230
// 수정자        : 곽윤설
// 수정내용      : E001 참고치 변경에 따른 소견 판정 수정
//                              E001(T3)
//                 THH1  >=2.11  -->  >=2.01
//                 THL1  <=0.59  -->  <=0.79
//                 TSH2  0.6~2.1 -->  0.8~2.0
//                 TSH   0.6~2.1 -->  0.8~2.0
//                 TSH7  0.6~2.1 -->  0.8~1.9
//                 TSHB  <=0.59  -->  <=0.79
//                       >=2.11       >=2.01
//                 THH   >=0.6   -->  >=0.8
//                       >=2.11       >=1.91
//                 THL   <=2.1   -->  <=1.91
//                       <=0.59       <=0.79
//==============================================================================
// 수정일자      : 2015.04.30
// 수정자        : 곽윤설
// 수정내용      : PSA 소견코드 - T011 >=4.1에서 >=4.01로 변경
// 수정내용      : CA125, T007 SKIP 제외
//==============================================================================
// 수정일자      : 2015.11.06
// 수정자        : 곽윤설
// 수정내용      : C054(sGulkwa[122]) 소견기준 변경
//                 ICAH   >= 1.351  -->  >= 1.321
//                 ICAL   <= 1.049  -->  <= 0.09  으로 변경
// 수정내용      : 김소영 전문의 쪽지로 요청.
//==============================================================================
// 수정일자      : 2015.12.31
// 수정자        : 박수지
// 수정내용      : AFPD(TT01,TT02,TT03), CA2N, CA5N, PSAN,CEAN,CA9N,CA7N,BETN,SCCN,CYFN
//                 _정상일시 해당항목에 해당코드 들어가게 요청
// 수정내용      : 김소영 전문의 - 한의 본원진료팀1500618
//==============================================================================
// 수정일자      : 2016.03.07
// 수정자        : 박수지
// 수정내용      : 자동소견 추가
//                  T012 <=16.3  --> NSEN
//                  T012 >+= 16.31 --> NSE
// 수정내용      : 김소영 전문의 - 한의 본원진료팀1600199
//==============================================================================
// 수정일자      : 2016.03.23
// 수정자        : 박수지
// 수정내용      : NH4 소견 수정 -김소영 선생님
//==============================================================================
// 수정일자      : 20160325
// 수정자        : 박수지
// 수정내용      : 호모시스테인 시약 변경에 따른 자동 소견 수정
//==============================================================================
// 수정일자      : 20160518
// 수정자        : 박수지
// 수정내용      : 당뇨관련 LDLC과 요단백 관련한 소견 추가
//                 LDLF LDLN LDLQ UPDS UPDM  UPH1 UPHR UPB UPBW 소견코드 추가
//==============================================================================
// 수정일자      : 20160602
// 수정자        : 박수지
// 수정내용      : 추가 ; C038 >=7.01 --> INAB   C020 <=6.30 --> CMBN
//                 수정 ; C020 >=6.31 --> CKMB
//                 C083코드 추가 171 -> 172
//==============================================================================
// 수정일자      : 20160622
// 수정자        : 박수지
// 수정내용      : 추가 ; CPN 자동 소견 추가
//==============================================================================
// 수정일자      : 20160630
// 수정자        : 박수지
// 수정내용      : CRP1, CRP3 -> 비교값 수정, CRP4, CRP5 소견 생성
//                 본원 김소영 전문의 요청
//==============================================================================
// 수정일자      : 2017.02.16
// 수정자        : 박수지
// 수정내용      :  Amylase(C039) >=109.1
//                 본원 김소영 전문의 요청
//==============================================================================
// 수정일자      : 2017.06.21
// 수정자        : 박수지
// 수정내용      :  pt소견 신규 생성및 소견 추가
//                 본원 김소영 전문의 요청
//==============================================================================
// 수정일자      : 2017.06.26
// 수정자        : 박수지
// 수정내용      :  S085, C065, E058 관련 소견 추가
//                 본원 김소영 전문의 1700284
//==============================================================================
// 수정일자      : 2017.07.25
// 수정자        : 박수지
// 수정내용      : HE4, HE4N, ICAH, ICAL, ADAH, ADAL 관련 소견 추가
//                 본원 김소영 전문의
//==============================================================================
// 수정일자      : 2017.11.24
// 수정자        : 박수지
// 수정내용      : PEPC, AHPM
//                 본원 김소영 전문의
//==============================================================================

unit LD6I01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, ValEdit, Gauges;

type
  TfrmLD6I01 = class(TfrmSingle)
    pnlMaster: TPanel;
    Panel2: TPanel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    btnStart: TBitBtn;
    ProBar: TGauge;
    Animate1: TAnimate;
    Animate2: TAnimate;
    Label1: TLabel;
    Panel3: TPanel;
    labStatus: TLabel;
    mskFrom: TMaskEdit;
    Label10: TLabel;
    mskTo: TMaskEdit;
    Panel4: TPanel;
    qryGumgin: TQuery;
    qryProfileHm: TQuery;
    qryHangmok: TQuery;                                                               
    qryGlkwa: TQuery;
    qrySokun: TQuery;
    qryInsert: TQuery;
    qryHul_sokun: TQuery;
    qryBunju: TQuery;
    qrykicho: TQuery;
    qryNo_hangmok: TQuery;
    cmbDoctor: TComboBox;
    Label2: TLabel;
    qryDoctor: TQuery;
    Label3: TLabel;
    qryPf_hm2: TQuery;
    QryTot_Munjin2018: TQuery;
    pnlCod_dc: TPanel;
    Label4: TLabel;
    edtDc: TEdit;
    edtDESC_DC: TEdit;
    btnDc: TSpeedButton;
    chkCod_Dc: TCheckBox;
    chk_Sex: TCheckBox;
    Ck_hulgulkwa: TCheckBox;
    qryHgulkwa_EtcChk: TQuery;

    procedure btnStartClick(Sender: TObject);
    procedure btnDateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnDcClick(Sender: TObject);
    procedure chkCod_DcClick(Sender: TObject);
  private
    { Private declarations }

    Sk_Value, TG_Check,gPreg : String;
    sGulkwa  : Array[0..182] Of String;
    sCheck   : Array[0..182] Of String;
    tSokun   : Array[0..182] Of String;
    sSokun   : Array[0..182] Of String;
    uSokun   : Array[0..182] Of String;
    tSokun_chuga: Array[0..182] Of String;
    biman    : Single;
    Hul_h, Hul_l : Integer;
    K, S, U,C, cSex, sAge, aSkip, dCheck  : Integer;
    Chk_Dang_Drug, Chk_Dang_Jindan : String;

    Procedure Sokun_Rtn;
    Procedure End_Sokun_Rtn;
    Procedure Insert_Update_sokun;
    Procedure Child_Sokun_Rtn;
    Procedure Chuga_Sokun_Rtn;
    Procedure NOT_C011_Sokun_Rtn;
    Procedure Danche_Sokun_Rtn;

    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

  public
    UV_sJisa : String;
    { Public declarations }
  end;

var
  frmLD6I01: TfrmLD6I01;
  UV_sHul_Part : String;
  
implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc;

{$R *.DFM}

function  TfrmLD6I01.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         //Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

Procedure TfrmLD6I01.NOT_C011_Sokun_Rtn;
Var
   Val1, Val2, Val3, Val4, val5, val6 : Extended;
   sVal1, sVal2, sVal3 : String;
   dCnt1, dCheck, I, J : Integer;

Begin
   //시립대
   If (sGulkwa[09] <> '') and
      (sGulkwa[10] <> '') Then
      Begin
         if (sGulkwa[09])<> '' then Val1 := StrToFloat(sGulkwa[09])
            else val1 :=0;
         if (sGulkwa[10])<> '' then Val2 := StrToFloat(sGulkwa[10])
            else val2 :=0;
         if (sGulkwa[55])<> '' then Val3 := StrToFloat(sGulkwa[55])
            else val3 :=0;
         if (sGulkwa[115])<> '' then Val4 := StrToFloat(sGulkwa[115])
            else val4 :=0;
         sVal1 := Trim(sGulkwa[70]);
         sVal2 := Trim(sGulkwa[71]);
         //HPD
         If (cSex  = 1)       And
            (val1  <= 60.0)   And
            ((val2 >= 45.1)   And (Val2 <= 60.0)) And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPD ';
         End;
         If (cSex  = 2)       And
            (val1  <= 60.0)   And
            ((val2 >= 34.1)   And (Val2 <= 60.0)) And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPD ';
         End;
         //FL
         If (cSex  = 1)       And
            ((val1 >= 60.1)   And (Val1 <= 100))  And
            ((val2 >= 45.1)   And (Val2 <= 60.0)) And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'FL  ';
         End;
         If (cSex  = 2)       And
            ((val1 >= 60.1)   And (Val1 <= 100))  And
            ((val2 >= 34.1)   And (Val2 <= 60.0)) And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'FL  ';
         End;
         If (val1 <= 100)     And
            ((val2 >= 60.1)   And (Val2 <= 100))  And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'FL  ';
         End;
         //HPM
         If (cSex  = 1)       And
            (val1 >= 100.1)   And
            (val2 >= 45.1)    And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPM ';
         End;
         If (cSex  = 2)       And
            (val1 >= 100.1)   And
            (val2 >= 34.1)    And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPM ';
         End;
         If (val2 >= 100.1)    And
            ((sVal1 = '음성') Or  (Val3 <= 0.99)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPM ';
         End;
         //AST
         If (cSex  = 1)       And
            (val1 >= 40.1)    And
            (val2 <= 45.0)    Then
         Begin
            K := K + 1;
            tSokun[K] := 'AST ';
         End;
         If (cSex  = 2)       And
            (val1 >= 32.1)    And
            (val2 <= 34.0)    Then
         Begin
            K := K + 1;
            tSokun[K] := 'AST ';
         End;
         //HBM
         If (cSex  = 1)       And
            (val1 <= 100.0)   And
            ((val2 >= 45.1)   And (Val2 <= 100))    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBM ';
         End;
         If (cSex  = 2)       And
            (val1 <= 100.0)   And
            ((val2 >= 34.1)   And (Val2 <= 100))    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBM ';
         End;
         //HPB
         If (cSex  = 1)       And
            (val2 >= 100.1)   And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPB ';
         End;
         If (cSex  = 1)       And
            (val1 >= 100.1)   And
            (val2 >= 45.1)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPB ';
         End;
         If (cSex  = 2)       And
            (val1 >= 100.1)   And
            (val2 >= 34.1)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HPB ';
         End;
         //HBC1
         If (cSex  = 1)       And
            (val1 <= 40.0)    And
            (val2 >= 45.0)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBC1';
         End;
         If (cSex  = 2)       And
            (val1 <= 32.0)   And
            (val2 <= 34.0)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBC1';
         End;
         //HBC + AST
         If (cSex  = 1)       And
            (val1 >= 40.1)    And
            (val2 <= 45.0)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBC ';
            K := K + 1;
            tSokun[K] := 'AST ';
         End;
         If (cSex  = 2)       And
            (val1 >= 32.1)    And
            (val2 <= 34.0)    And
            ((sVal1 = '양성') Or  (Val3 >= 1.0))    And
            ((sVal2 = '음성') Or  (Val4 < 10))      Then
         Begin
            K := K + 1;
            tSokun[K] := 'HBC ';
            K := K + 1;
            tSokun[K] := 'AST ';
         End;
   End;
   If (qryBunju.FieldByName('Dat_bunju').AsString < '20140428') And
      (sGulkwa[92] <> '') and (sGulkwa[97] <> '') Then
       Begin
           Val1 := StrToFloat(sGulkwa[92]);
           Val3 := StrToFloat(sGulkwa[97]);
           If (Val1 >= 30) and (Val3 >= 10) and (cSex = 1)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'UPB ';
              End
           else If (Val1 >= 30) and (Val3 >= 10) and  (cSex = 2)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'UPBW';
              End
           else  If (Val1 <= 20) and (Val3 >= 10) and (cSex = 1)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW ';
              End
           else  If (Val1 <= 20) and (Val3 >= 10) and (cSex = 2) and  (sAge <=50)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW5';
              End
            else  If (Val1 <= 20) and (Val3 >= 10) and (cSex = 2) and  (sAge >=51)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW6';
              End;
      End;
   If (qryBunju.FieldByName('Dat_bunju').AsString >= '20140428') And
      (sGulkwa[92] <> '') and (sGulkwa[97] <> '') Then
       Begin
           Val1 := StrToFloat(sGulkwa[92]);
           Val3 := StrToFloat(sGulkwa[97]);
           If (Val1 >= 10) and (Val3 >= 10) and (cSex = 1)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'UPB ';
              End
           else If (Val1 >= 10) and (Val3 >= 10) and  (cSex = 2)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'UPBW';
              End
           else If (Val1 <= 5) and (Val3 >= 10) and (cSex = 1)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW ';
              End
           else  If (Val1 <= 5) and (Val3 >= 10) and (cSex = 2) and  (sAge <=50)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW5';
              End
            else  If (Val1 <= 5) and (Val3 >= 10) and (cSex = 2) and  (sAge >=51)  then
              Begin
                 K := K + 1;
                 tSokun[K] := 'URW6';
              End;
      End;
End;

Procedure TfrmLD6I01.Danche_Sokun_Rtn;
Var
   Val1, Val2, Val3, Val4, val5, val6 : Extended;
   sVal1, sVal2, sVal3 : String;
   dCnt1, dCheck, I, J : Integer;

Begin
   //항목별 정상소견
   //AFPD
   If (sGulkwa[118] <> '') Then
      Begin
         if (sGulkwa[118]) <> '' then Val1 := StrToFloat(sGulkwa[118])
            else val1 :=0;
         If (val1  <= 8.99)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'AFPD';
         End;
      End;

   IF (sGulkwa[88] <> '') then
      BEGIN
        sVal1 := Trim(sGulkwa[88]);
         If sVal1 = '음성'  Then
         Begin
            K := K + 1;
            tSokun[K] := 'AFPD';
         End;
      END;

   //CEAN
   If (sGulkwa[82] <> '') Then
      Begin
         if (sGulkwa[82])<> '' then Val1 := StrToFloat(sGulkwa[82])
            else val1 :=0;
         If (val1  <= 5.0)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'CEAN';
         End;
      End;
   //CA2N
   If (sGulkwa[84] <> '') OR (sGulkwa[175] <> '')Then
      Begin
         if (sGulkwa[84])<> '' then Val1 := StrToFloat(sGulkwa[84])
            else val1 :=0;
         if (sGulkwa[175])<> '' then Val2 := StrToFloat(sGulkwa[175])
            else val2 :=0;
         If (((sGulkwa[84])<> '') and (val1  <= 35.0)) OR (((sGulkwa[175])<> '') and (val2 <= 35))  Then
         Begin
            K := K + 1;
            tSokun[K] := 'CA2N';
         End;
      End;
   //CA5N
   If (sGulkwa[142] <> '') Then
      Begin
         if (sGulkwa[142])<> '' then Val1 := StrToFloat(sGulkwa[142])
            else val1 :=0;
         If (val1 <= 28.0)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'CA5N';
         End;
      End;
   //CYFN
   If (sGulkwa[104] <> '') Then
      Begin
         if (sGulkwa[104])<> '' then Val1 := StrToFloat(sGulkwa[104])
            else val1 :=0;
         If (val1 <= 3.0)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'CYFN';
         End;
      End;
   //CA7N
   If (sGulkwa[139] <> '') Then
      Begin
         if (sGulkwa[139])<> '' then Val1 := StrToFloat(sGulkwa[139])
            else val1 :=0;
         If (val1 <= 6.9)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'CA7N';
         End;
      End;
   //CA9N
   If (sGulkwa[83] <> '') Then
      Begin
         if (sGulkwa[83])<> '' then Val1 := StrToFloat(sGulkwa[83])
            else val1 :=0;
         If (val1  <= 37.0)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'CA9N';
         End;
      End;
   //PSAN
   If (sGulkwa[87] <> '') or (sGulkwa[174] <> '')Then
      Begin
         if (sGulkwa[87])<> '' then Val1 := StrToFloat(sGulkwa[87])
            else val1 :=0;
         if (sGulkwa[174])<> '' then Val2 := StrToFloat(sGulkwa[174])
            else val2 :=0;
         If (((sGulkwa[87]) <> '') and (val1  <= 4.0)) or
            (((sGulkwa[174]) <> '') and (val2 <= 3.6)) Then
         Begin
            K := K + 1;
            tSokun[K] := 'PSAN';
         End;
      End;


   //BETN
   If (sGulkwa[32] <> '') Then
      Begin
         if (sGulkwa[32])<> '' then Val1 := StrToFloat(sGulkwa[32])
            else val1 :=0;
         If (val1 <= 3.0)   Then
         Begin
            K := K + 1;
            tSokun[K] := 'BETN';
         End;
      End;
End;

Procedure TfrmLD6I01.Sokun_Rtn;
Var
   Val1, Val2, Val3, Val4,val5,val6 : Extended;
   dCnt1, dCheck, I, J : Integer;

Begin

// 자동소견결과처리 Routine (Sgulkwa Table에서 Position(Index)를 참고하여 찾을것. //
// ??? == 주의 == ??? 처리순서는 소견 Sort 순서임
  For I := 0 To 182 Do
      tSokun[I] := '';
//      if qryBunju.FieldByName('num_bunju').Asinteger < 7000  then
  If True Then
      begin
//간장질환
//999.Cod = DDD1 (S007, S008 의 결과가 없을때(간염검사無)는 정상 결과처리 2003.01.03)
//    SE07,SE08 >> SE09,SE10 >> S007, S008 우선순위로 결정 2003.01.03

         If ((sGulkwa[70] <> '') and (sGulkwa[71] <> '')) Or
            ((sGulkwa[55] <> '') and (sGulkwa[115] <> '')) Then
         Begin
            if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
            if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
               else val2 :=0;
            If ((sGulkwa[70] = '음성') or  ((val1 > 0) and (val1 <1.0)))  and
               ((sGulkwa[71] = '양성') or  (val2 >= 30))  and
                ((sGulkwa[08] = '음성') or (sGulkwa[08] = '')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBAB';
               End
        End;


         If (sGulkwa[100] <> '') And (sGulkwa[101] <> '') Then
            Begin
               sGulkwa[70] := sGulkwa[100];
               sGulkwa[71] := sGulkwa[101];
            End
         Else If (sGulkwa[102] <> '') And (sGulkwa[103] <> '') Then
            Begin
               sGulkwa[70] := sGulkwa[102];
               sGulkwa[71] := sGulkwa[103];
            End
         Else If (sGulkwa[70] = '') And (sGulkwa[71] = '') Then
            Begin
               sGulkwa[70] := '음성';
               sGulkwa[71] := '양성';
            End;

//87.Cod = HCV1 (S016 = '양성) And (HPD or HPM)일때

           If  (sGulkwa[10] <> '') and   (sGulkwa[72] <> '') Then
            Begin
               if (sGulkwa[09])<> '' then Val1 := StrToFloat(sGulkwa[09])
               else val1 :=0;
               val2 := StrToFloat(sGulkwa[10]);
               if (sGulkwa[11])<> '' then val3 := StrToFloat(sGulkwa[11])
               else val3 :=0;
               If (sGulkwa[72] = '양성') And (Val2 >= 40.1) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val2 >= 45.1) And (cSex  = 1) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val2 >= 34.1) And (cSex  = 2)and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val1>= 40.1) and (Val2<=40.0) and (Val3>=50.1)  And (cSex  = 1) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val1>= 40.1) and (Val2<=45) and (Val3>=70.1)  And (cSex  = 1) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val1>= 40.1) and (Val2<=39.9) and (Val3>=40.1)  And (cSex  = 2) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
               else If (sGulkwa[72] = '양성') And (Val1>= 32.1) and (Val2<=34) and (Val3>=42.1)  And (cSex  = 2) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV1';
                  End
            End;
            If  (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
                (sGulkwa[11] <> '') And (sGulkwa[72] <> '') Then
                 Begin
                 Val1 := StrToFloat(sGulkwa[09]);
                 val2 := StrToFloat(sGulkwa[10]);
                 val3 := StrToFloat(sGulkwa[11]);
               If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 40.0) and (Val3 >= 50.1) And (cSex  = 1) and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV7';
                  End
               else If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 45.0) and (Val3 >= 70.1) And (cSex  = 1)  and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                    Begin
                     K := K + 1;
                     tSokun[K] := 'HCV7';
                   End
               else  If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 39.9) and (Val3 >= 40.1) And (cSex  = 2)   and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV7';
                  End
               else  If (sGulkwa[72] = '양성') And (Val1 <= 32) and (Val2 <= 34) and (Val3 >= 42.1) And (cSex  = 2)   and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV7';
                  End
            End;
            If  (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
                (sGulkwa[11] <> '') And (sGulkwa[72] <> '') Then
                 Begin
                 Val1 := StrToFloat(sGulkwa[09]);
                 val2 := StrToFloat(sGulkwa[10]);
                 val3 := StrToFloat(sGulkwa[11]);
                 If (sGulkwa[72] = '양성') And (Val1 >= 40.1) and (Val2 <= 45.0) and (Val3 <= 70) And (cSex  = 1)  and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                    Begin
                     K := K + 1;
                     tSokun[K] := 'HCVS';
                   End
               else  If (sGulkwa[72] = '양성') And (Val1 >= 32.1) and (Val2 <= 34) and (Val3 <= 42) And (cSex  = 2)   and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCVS';
                  End
            End;
            If  (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
                (sGulkwa[11] <> '') And (sGulkwa[72] <> '') Then
                 Begin
                 Val1 := StrToFloat(sGulkwa[09]);
                 val2 := StrToFloat(sGulkwa[10]);
                 val3 := StrToFloat(sGulkwa[11]);
               If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 40.0) and (Val3 >= 50.1) And (cSex  = 1) and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'GAH9';
                  End
               else If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 45) and (Val3 >= 70.1) And (cSex  = 1)  and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'GAH9';
                  End
               else  If (sGulkwa[72] = '양성') And (Val1 <= 40.0) and (Val2 <= 39.9) and (Val3 >= 40.1) And (cSex  = 2)  and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'GAH9';
                  End
               else  If (sGulkwa[72] = '양성') And (Val1 <= 32.0) and (Val2 <= 34) and (Val3 >= 42.1) And (cSex  = 2)   and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'GAH9';
                  End
            End;
           
{            If (sGulkwa[72] <> '') Then
               Begin
               If (sGulkwa[72] = '약양성')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HCV3';
                  End
            End;
 }

//88.Cod = HCV (S016 = '양성) And ((C009 <= 40) And (C010 <= 40) And (C011 <= 50))
         If (tSokun[K] <> 'HPB ') And
            (tSokun[K] <> 'HCV1') Then
            If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
               (sGulkwa[11] <> '') And (sGulkwa[72] <> '') Then
               Begin
                  Val1 := StrToFloat(sGulkwa[09]);
                  Val2 := StrToFloat(sGulkwa[10]);
                  Val3 := StrToFloat(sGulkwa[11]);
                  If (sGulkwa[72] = '양성') And (Val1 <= 40) And (cSex  = 1) and
                     (Val2 <= 40) And (Val3 <= 50) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'HCV ';
                     End
                  else If (sGulkwa[72] = '양성') And (Val1 <= 40) And (cSex  = 1) and
                     (Val2 <= 45) And (Val3 <= 70) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'HCV ';
                     End
                  else If (sGulkwa[72] = '양성') And (Val1 <= 40) And (cSex  = 2) and
                     (Val2 <= 39.9) And (Val3 <= 40) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'HCV ';
                     End
                  else If (sGulkwa[72] = '양성') And (Val1 <= 32) And (cSex  = 2) and
                     (Val2 <= 34) And (Val3 <= 42) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'HCV ';
                     End;

               End;


//84.Cod = HAB ((S007 = 양성 or S007 = 약양성) And S008 = 양성)
      If (sGulkwa[70] <> '') And ((sGulkwa[71] <> '') or (sGulkwa[115] <> '')) Then
         Begin
            if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
            if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
               else val2 :=0;
            If (((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성')) And
               ((sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성'))) Or
               ((Val1 >= 10) And (Val2 > 10)) Then
//            If (((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성')) or (Val1 >= 10)) And
//               (((sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성')) or (Val2 >= 10)) Then


               Begin
                  K := K + 1;
                  tSokun[K] := 'HAB ';
               End;
         End;
//20121011

        If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And (sGulkwa[11] <> '') Then
            Begin
               val1 :=0;
               val2 :=0;
               val3 :=0;
               val4 :=0;
               val5 :=0;
               Val1 := StrToFloat(sGulkwa[10]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[11]);
               if (sGulkwa[55])<> '' then Val4 := StrToFloat(sGulkwa[55])
               else val4 :=0;
               if (sGulkwa[115])<> '' then Val5 := StrToFloat(sGulkwa[115])
               else val5 :=0;

               If ((Val1>=45.1) and (Val1<=60)) and (Val2<=60) and (Val3<=85) and  (cSex = 1) and
                  ((((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))or ((Val4>0) and (Val4<=9.99))) or ((sGulkwa[70] ='')and (Val4=0)) or ((sGulkwa[71] ='')and (Val5 =0)))and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = '') )and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End
               else If ((Val1>=40) and (Val1<=59.9)) and (Val2<=59.9) and (Val3<=79.9) and (cSex = 2) and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End
               else If ((Val1>=34.1) and (Val1<=60)) and (Val2<=60) and (Val3<=85) and (cSex = 2) and
                       ((((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))or ((Val4>0) and (Val4<=9.99)))   or ((sGulkwa[70] ='')and (Val4=0)) or ((sGulkwa[71] ='')and (Val5 =0)))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))  and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End

               else If (Val1<=40) and ((val2>=40.1) and(Val2<=59.9)) and ((Val3>=50.1) and (Val3<=79.9)) and (cSex = 1) and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End
               else If (Val1<=45) and ((val2>=40.1) and(Val2<=60)) and ((Val3>=70.1) and (Val3<=85)) and (cSex = 1) and
                       ((((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))or ((Val4>0) and (Val4<=9.99)))  or ((sGulkwa[70] ='')and (Val4=0)) or ((sGulkwa[71] ='')and (Val5 =0))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End
               else If (Val1<=39.9) and ((val2>=40.1) and(Val2<=59.9)) and ((Val3>=40.1) and (Val3<=79.9)) and (cSex = 2) and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End
               else If (Val1<=34) and ((val2>=32.1) and(Val2<=60)) and ((Val3>=42.1) and (Val3<=85)) and (cSex = 2) and
                       ((((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))or ((Val4>0) and (Val4<=9.99)))  or ((sGulkwa[70] ='')and (Val4=0)) or ((sGulkwa[71] ='')and (Val5 =0)))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPD ';
                  End;



            End;

            //FLS
            If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '') Then
            Begin
               val1 :=0;
               val2 :=0;
               val3 :=0;
               Val1 := StrToFloat(sGulkwa[10]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[11]);

               if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
               else Val4 :=0;
               if (sGulkwa[55])<> '' then Val5 := StrToFloat(sGulkwa[55])
               else val5 :=0;
               If ((Val1>=45.1)and (Val1<=100)) and
                  (Val2<=100) and
                  (Val3>=200.1)  and
                  (cSex = 1)   and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성') or ((Val5>0)and(Val5<=9.99))) or ((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='') and (val4=0)))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FLS ';
                  End
                else If ((Val1>=34.1)and (Val1<=100)) and
                  (Val2<=100) and
                  (Val3>=200.1)  and
                  (cSex = 2)   and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or((Val5>0)and(Val5<=9.99))) or ((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='') and (val4=0)))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = ''))and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FLS ';
                  End
               else If (Val1<=45)    and
                       ((Val2>=40.1) and (Val2<=100)) and
                       (Val3>=200.1)   and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or ((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='') and (val4=0))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = ''))and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FLS ';
                  End
                 else If (Val1<=34)  and
                       ((Val2>=32.1) and (Val2<=100)) and
                       (Val3>=200.1)   and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or ((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='') and (val4=0)))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = ''))and ((qryBunju.FieldByName('Dat_bunju').AsString)>='20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FLS ';
                  End;
            End;
            ///FL
            If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '') Then
            Begin
               val1 :=0;
               val2 :=0;
               val3 :=0;
               Val1 := StrToFloat(sGulkwa[10]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[11]);
               if sGulkwa[115] <> '' then   Val4 := StrToFloat(sGulkwa[115])
               else Val4 :=0;
               if (sGulkwa[55])<> '' then Val5 := StrToFloat(sGulkwa[55])
               else val5 :=0;
               If ((Val1>=60)  and (Val1<=99.9)) and
                  (Val2<=99.9) and
                  (Val3<=199.9)and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End

               else If ((Val1>=60.1)  and (Val1<=100)) and
                  (Val2<=100) and
                  (Val3<=200)and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
               else If (Val1<=45) and
                  ((Val2>=40.1) and (Val2<=100)) and
                  ((Val3>=85.1) and (Val3<=200)) and  (cSex = 1)    and
                 (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If (Val1<=34) and
                  ((Val2>=40.1) and (Val2<=100)) and
                  ((Val3>=85.1) and (Val3<=200)) and  (cSex = 2)    and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
               else If ((Val1>=40.1) and (Val1<=59.9)) and
                       ((Val2>=60)   and (Val2<=99.9)) and
                       (Val3<=199.9) and (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
               else If ((Val1>=45.1) and (Val1<=60)) and
                       ((Val2>=60.1)   and (Val2<=100)) and
                       (Val3<=200) and (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
               else If ((Val1>=40)   and (Val1<=59.9)) and
                       ((Val2>=60)   and (Val2<=99.9)) and
                       (Val3<=199.9) and (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If ((Val1>=34.1)   and (Val1<=60)) and
                       ((Val2>=60.1)   and (Val2<=100)) and
                       (Val3<=200) and (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If ((Val1>=40.1) and (Val1<=59.9)) and
                      (Val2<=59.9)  and
                      ((Val3>=80)   and (Val3<=199.9)) and
                      (cSex = 1)    and
                      (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                      ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If ((Val1>=45.1) and (Val1<=60)) and
                      (Val2<=100)  and
                      ((Val3>=85.1)   and (Val3<=200)) and
                      (cSex = 1)    and
                      (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                      ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If ((Val1>=40) and (Val1<=59.9)) and
                      (Val2<=59.9)and
                      ((Val3>=80) and (Val3<=199.9))and
                      (cSex = 2)  and
                      (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                      ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
              else If ((Val1>=34.1) and (Val1<=60)) and
                      (Val2<=100)and
                      ((Val3>=85.1) and (Val3<=200))and
                      (cSex = 2)  and
                      (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                      ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=40)    and
                       ((Val2>=60)   and (Val2<=99.9)) and
                       ((Val3>=50.1) and (Val3<=199.9))and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=45)    and
                       ((Val2>=60.1) and (Val2<=100)) and
                       ((Val3>=70.1) and (Val3<=200)) and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=45)    and
                       ((Val2>=40.1) and (Val2<=60)) and
                       ((Val3>=85.1) and (Val3<=200)) and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=39.9)  and
                       ((Val2>=60)   and (Val2<=99.9)) and
                       ((Val3>=40.1) and (Val3<=199.9))and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=34)  and
                       ((Val2>=60.1)   and (Val2<=100)) and
                       ((Val3>=42.1) and (Val3<=200))and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=34)  and
                       ((Val2>=32.1)   and (Val2<=60)) and
                       ((Val3>=85.1) and (Val3<=200))and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End
             else If   (Val1<=40)    and
                       ((Val2>=40.1) and (Val2<=59.9)) and
                       ((Val3>=80.1) and (Val3<=199.9)) and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = ''))and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End

              else If  (Val1<=39.9)  and
                       ((Val2>=40.1) and (Val2<=59.9)) and
                       ((Val3>=80.1) and (Val3<=199.9)) and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FL  ';
                  End;

            End;

            If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '') Then
            Begin
               val1 :=0;
               val2 :=0;
               val3 :=0;
               Val1 := StrToFloat(sGulkwa[10]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[11]);
               if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
               else Val4 :=0;
               if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
               else Val5 :=0;


               If (Val1>=100) and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] ='')))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = '')) and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPM ';
                  End
               else If (Val1>=100.1) and
                  (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                  ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = '')) and
                  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPM ';
                  End
               else If (Val1<=40)  and
                       (Val2>=100) and
                       (Val3>=50.1)and
                       (cSex = 1)  and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = '')) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                        Begin
                        K := K + 1;
                        tSokun[K] := 'HPM ';
                  End
               else If (Val1<=45)  and
                       (Val2>=100.1) and
                       (Val3>=70.1)and
                       (cSex = 1)  and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = '')) and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                        Begin
                        K := K + 1;
                        tSokun[K] := 'HPM ';
                  End
               else If (Val1<=39.9)and
                       (Val2>=100) and
                       (Val3>=40.1)and
                       (cSex = 2)  and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End
               else If (Val1<=34)and
                       (Val2>=100.1) and
                       (Val3>=42.1)and
                       (cSex = 2)  and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End
               else If ((Val1>=40.1) and (Val1<=99.9)) and
                       (Val2>=100)   and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성')  or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End
               else If ((Val1>=45.1) and (Val1<=100)) and
                       (Val2>=100.1)   and
                       (cSex = 1)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End
               else If ((Val1>=40) and (Val1<=99.9)) and
                       (Val2>=100)   and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성')  or (sgulkwa[70] = '약양성')) or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End
               else If ((Val1>=34.1) and (Val1<=100)) and
                       (Val2>=100.1)   and
                       (cSex = 2)    and
                       (((sGulkwa[70] ='음성') or (sgulkwa[70] = '약양성')or ((Val5>0)and(Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val4=0))))  and
                       ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')   or (sGulkwa[72] = ''))and
                       ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                       Begin
                       K := K + 1;
                       tSokun[K] := 'HPM ';
                  End;
            End;



            //HBM
            If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '') Then
            Begin
               val1 :=0;
               val2 :=0;
               val3 :=0;
               Val1 := StrToFloat(sGulkwa[10]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[11]);
               if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
               else Val4 :=0;
               if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
               else Val5 :=0;

               If      ((Val1>=40.1)  and (Val1<=99.9)) and
                       (Val2<=99.9) and
                       (Val3<=199.9)and
                       (cSex = 1)  and
                       (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If ((Val1>=45.1)  and (Val1<=100)) and
                       (Val2<=100) and
                       (Val3<=200)and
                       (cSex = 1)  and
                       ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If ((Val1>=34.1)  and (Val1<=100)) and
                       (Val2<=100) and
                       (Val3<=200)and
                       (cSex = 2)  and
                       ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If ((Val1>=40.0)  and (Val1<=99.9)) and
                       (Val2<=99.9) and
                       (Val3<=199.9)and
                       (cSex = 2)  and
                       (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If (Val1<=40.0) and
                       ((Val2>=40.1)and (Val2<=99.9))  and
                       ((Val3>=50.1)and (Val3<=199.9)) and
                       (cSex = 1)   and
                       (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If (Val1<=45) and
                       ((Val2>=40.1)and (Val2<=100))  and
                       ((Val3>=70.1)and (Val3<=200)) and
                       (cSex = 1)   and
                       ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If (Val1<=34) and
                       ((Val2>=32.1)and (Val2<=100))  and
                       ((Val3>=42.1)and (Val3<=200)) and
                       (cSex = 2)   and
                       ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End
               else If (Val1<=39.9) and
                       ((Val2>=40.1)and (Val2<=99.9)) and
                       ((Val3>=40.1)and (Val3<=199.9))and
                       (cSex = 2)   and
                       (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End;
            End;
            
              ///sGulkwa[009] =C009 ,   sGulkwa[010] =C010 , sGulkwa[011] =C011 ,
         If (sGulkwa[009] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[011] <> '')  Then
         Begin
            Val1 := StrToFloat(sGulkwa[010]);
            Val2 := StrToFloat(sGulkwa[009]);
            Val3 := StrToFloat(sGulkwa[011]);
            if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
               else Val4 :=0;
            if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
               else Val5 :=0;

            If  ((Val1 >= 40.1) and (val1<=99.9 ))  and
                (val2 <=99.9)   and
                (Val3 >= 200)   and
                (cSex = 1)   and
                (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                   K := K + 1;
                   tSokun[K] := 'HBMG';
                End
            else If  ((Val1 >= 45.1) and (val1<=100 ))  and
                (val2 <=100)   and
                (Val3 >= 200.1)   and
                (cSex = 1)   and
                ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                   K := K + 1;
                   tSokun[K] := 'HBMG';
                End
             else If  ((Val1 >= 40) and (val1<=99.9 ))  and
                      (val2 <=99.9)   and
                      (Val3 >= 200)   and
                      (cSex = 2)   and
                      (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End
             else If  ((Val1 >= 34.1) and (val1<=100 ))  and
                      (val2 <=100)   and
                      (Val3 >= 200.1)   and
                      (cSex = 2)   and
                      ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End
             else If  (Val1 <= 40)    and
                      ((val2 >=40.1)  and (val2 <=99.9)) and
                      (Val3 >= 200)   and
                      (cSex = 1)   and
                      (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End
             else If  (Val1 <= 45)    and
                      ((val2 >=40.1)  and (val2 <=100)) and
                      (Val3 >= 200.1)   and
                      (cSex = 1)   and
                      ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End
             else If  (Val1 <= 39.9)    and
                      ((val2 >=40.1)  and (val2 <=99.9)) and
                      (Val3 >= 200)   and
                      (cSex = 2)   and
                      (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End
             else If  (Val1 <= 34)   and
                      ((val2 >=32.1)  and (val2 <=100)) and
                      (Val3 >= 20031)   and
                      (cSex = 2)   and
                      ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HBMG';
                      End;
        End;

        ////
        If (sGulkwa[009] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[011] <> '')  Then
        Begin
            Val1 := StrToFloat(sGulkwa[010]);
            Val2 := StrToFloat(sGulkwa[009]);
            Val3 := StrToFloat(sGulkwa[011]);
            if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
            else Val4 :=0;
            if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
            else Val5 :=0;

            If  (Val1 >= 100) and
                 (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HPB ';
            End
            else If  (Val1 >= 100.1) and
                 ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HPB ';
            End
            else If  ((Val1 >= 40.1) and (val1<=99.9 ))  and
                      (val2 >=100)    and
                      (cSex = 1)      and
                      (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
                   Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
             else If  ((Val1 >= 45.1) and (val1<=100 ))  and
                      (val2 >=100.1)    and
                      (cSex = 1)      and
                     ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
             else If  ((Val1 >= 40) and (val1<=99.9 ))  and
                      (val2 >=100)  and
                      (cSex = 2)    and
                      (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
             else If  ((Val1 >= 34.1) and (val1<=100 ))  and
                      (val2 >=100.1)  and
                      (cSex = 2)    and
                      ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
             else If  (Val1 <= 40)  and
                      (val2 >=100)  and
                      (Val3 >=50.1) and
                      (cSex = 1)    and
                      (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
            else If  (Val1 <= 45)  and
                      (val2 >=100.1)  and
                      (Val3 >=70.1) and
                      (cSex = 1)    and
                      ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
            else If   (Val1 <= 39.9) and
                      (val2 >=100)   and
                      (Val3 >=40.1)  and
                      (cSex = 2)     and
                      (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End
             else If   (Val1 <= 34) and
                      (val2 >=100.1)   and
                      (Val3 >=42.1)  and
                      (cSex = 2)     and
                      ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                      Begin
                           K := K + 1;
                           tSokun[K] := 'HPB ';
                      End;
        End;
        ////
         ////
        If (sGulkwa[009] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[011] <> '')  Then
        Begin
            Val1 := StrToFloat(sGulkwa[010]);
            Val2 := StrToFloat(sGulkwa[009]);
            Val3 := StrToFloat(sGulkwa[011]);
            if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
            else Val4 :=0;
            if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
            else Val5 :=0;

            If  (Val1 <= 40) and
                (Val2 <= 40) and
                (Val3 >= 50.1) and
                (cSex = 1)     and
                (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'GAH8';
            End
            else If  (Val1 <= 45) and
                     (Val2 <= 40) and
                     (Val3 >= 70.1) and
                     (cSex = 1)     and
                     ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'GAH8';
            End
            else If  (Val1 <= 39.9) and
                     (Val2 <= 40) and
                     (Val3 >= 40.1) and
                     (cSex = 2)     and
                     (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'GAH8';
            End
            else If  (Val1 <= 34) and
                     (Val2 <= 32) and
                     (Val3 >= 42.1) and
                     (cSex = 2)     and
                     ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'GAH8';
            End;
         End;
               ////
        If (sGulkwa[009] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[011] <> '')  Then
        Begin
            Val1 := StrToFloat(sGulkwa[010]);
            Val2 := StrToFloat(sGulkwa[009]);
            Val3 := StrToFloat(sGulkwa[011]);
            if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
            else Val4 :=0;
            if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
            else Val5 :=0;

            If  (Val1 <= 40) and
                (Val2 <= 40) and
                (Val3 >= 50.1) and
                (cSex = 1)     and
                (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HBC ';
            End
            else If  (Val1 <= 45) and
                     (Val2 <= 40) and
                     (Val3 >= 70.1) and
                     (cSex = 1)     and
                     ((sGulkwa[70] ='양성') or (Val5>=10)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HBC ';
            End
            else If  (Val1 <= 39.9) and
                     (Val2 <= 40) and
                     (Val3 >= 40.1) and
                     (cSex = 2)     and
                     (sGulkwa[70] ='양성')  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HBC ';
            End
            else If  (Val1 <= 34) and
                     (Val2 <= 32) and
                     (Val3 >= 42.1) and
                     (cSex = 2)     and
                     ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            Begin
                   K := K + 1;
                   tSokun[K] := 'HBC ';
            End;
        End;
        ////

        If (sGulkwa[009] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[011] <> '')  Then
        Begin
             Val1 := StrToFloat(sGulkwa[010]);
             Val2 := StrToFloat(sGulkwa[009]);
             Val3 := StrToFloat(sGulkwa[011]);

             if sGulkwa[115]<> '' then   Val4 := StrToFloat(sGulkwa[115])
             else Val4 :=0;
             if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
             else Val5 :=0;

             If  (Val1 <= 40) and
                 (Val2 <= 40) and
                 (Val3 <= 50) and
                 (cSex = 1)     and
                 (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
             Begin
                   K := K + 1;
                   tSokun[K] := 'HBC1';
             End
             else If  (Val1 <= 45) and
                 (Val2 <= 40) and
                 (Val3 <= 70) and
                 (cSex = 1)     and
                 ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
             Begin
                   K := K + 1;
                   tSokun[K] := 'HBC1';
             End
             else If  (Val1 <= 39.9) and
                 (Val2 <= 40) and
                 (Val3 <= 40) and
                 (cSex = 2)     and
                 (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
             Begin
                   K := K + 1;
                   tSokun[K] := 'HBC1';
             End
             else If  (Val1 <= 34) and
                 (Val2 <= 32) and
                 (Val3 <= 42) and
                 (cSex = 2)     and
                 ((sGulkwa[70] ='양성') or (Val5>=10))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
             Begin
                   K := K + 1;
                   tSokun[K] := 'HBC1';
             End;
        End;
        ////


        If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '')  Then
            Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val4 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);
                if (sGulkwa[05]) <> '' then   Val4 := StrToFloat(sGulkwa[05])
                else Val4 :=0;
                if sGulkwa[55]<> '' then   Val5 := StrToFloat(sGulkwa[55])
                else Val5 :=0;

                If ((Val1<=40)   and (Val2<=40)) and
                   ((Val3>=50.1) and (Val3<=60)) and
                   (cSex = 1)    and
                   (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                   (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                   ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH0';
                End
                else If ((Val1<=45)   and (Val2<=40)) and
                        ((Val3>=70.1) and (Val3<=85)) and
                        (cSex = 1)    and
                        (((val4>=0.01) and (Val4<=1.20)) or (Val4=0))  and
                         (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH0';
                End
                else If ((Val1<=40)   and (Val2<=40)) and
                        ((Val3>=50.1) and (Val3<=199.9)) and
                        (cSex = 1)    and
                        (Val4>=1.21)   and
                         (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH2';
                End
                else If ((Val1<=45)   and (Val2<=40)) and
                        ((Val3>=70.1) and (Val3<=110)) and
                        (cSex = 1)    and
                        (Val4>=1.21)   and
                         (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH2';
                End
                else If ((Val1<=39.9) and (Val2<=40)) and
                        ((Val3>=40.1) and (Val3<=50))and
                        (cSex = 2)    and
                        (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                         (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAH0';
                End
                else If ((Val1<=34) and (Val2<=32)) and
                        ((Val3>=42.1) and (Val3<=52))and
                        (cSex = 2)    and
                        (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAH0';
                End
                else If ((Val1<=39.9) and (Val2<=40)) and
                        ((Val3>=40.1) and (Val3<=199.9))and
                        (cSex = 2)    and
                        (Val4>=1.21)   and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAH2';
                End
                else If ((Val1<=34) and (Val2<=32)) and
                        ((Val3>=42.1) and (Val3<=110))and
                        (cSex = 2)    and
                        (Val4>=1.21)   and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5 >0) and (val5 <=9.99))) or ((sGulkwa[70] ='')and (val5 =0))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAH2';
                End;
            End;

            If (sGulkwa[10] <> '') And (sGulkwa[09] <> '') And (sGulkwa[11] <> '')  Then
            Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val4 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);
                if (sGulkwa[05]) <> '' then   Val4 := StrToFloat(sGulkwa[05])
                else Val4 :=0;
                if (sGulkwa[55])<> '' then Val5 := StrToFloat(sGulkwa[55])
                else val5 :=0;
                if (sGulkwa[115])<> '' then Val6 := StrToFloat(sGulkwa[115])
                else val6 :=0;

                If ((Val1<=40)   and (Val2<=40)) and
                   ((Val3>=60.1) and (Val3<=109.9)) and
                   (cSex = 1)    and
                   (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                   (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or((val5>0) and (Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val6=0))))
                   and ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH ';
                End
                else  If ((Val1<=45)   and (Val2<=40)) and
                   ((Val3>=85.1) and (Val3<=110)) and
                   (cSex = 1)    and
                   (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                   (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or ((val5>0) and (Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val6=0))))
                   and ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>='20130518')  Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH ';
                End
                else If ((Val1<=39.9) and (Val2<=40)) and
                       ((Val3>=50.1) and (Val3<=109.9)) and
                       (cSex = 2)    and
                       (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                       (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or ((val5>0) and (Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val6=0))))
                       and ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                       K := K + 1;
                       tSokun[K] := 'GAH ';
                End
                else If ((Val1<=34) and (Val2<=32)) and
                       ((Val3>=52.1) and (Val3<=110)) and
                       (cSex = 2)    and
                       (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                       (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or ((val5>0) and (Val5<=9.99))) or (((sGulkwa[70] ='') and (Val5=0))or ((sGulkwa[71] ='')and (val6=0))))
                       and ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                       K := K + 1;
                       tSokun[K] := 'GAH ';
                End;

             End;
             If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '') and (sGulkwa[11] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);
                if (sGulkwa[05]) <> '' then   Val4 := StrToFloat(sGulkwa[05])
                else Val4 :=0;
                if (sGulkwa[55])<> '' then  Val5 := StrToFloat(sGulkwa[55])
                else val5 :=0;
                if (sGulkwa[115])<> '' then Val6 := StrToFloat(sGulkwa[115])
                else val6 :=0;



                If  ((Val1 <= 45) and (Val2 <= 40))  and
                    ((Val3 >= 110.1) and (Val3 <= 200)) and
                    (cSex = 1)     and
                    ((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성') or ((sGulkwa[70] ='') and (Val5=0)) or ((val5>0) and(val5 <9.99))) and
                    ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and
                    ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'GAH1';
                End
                else If  ((Val1 <= 39.9) and (Val2 <= 40))  and
                    ((Val3 >= 110) and (Val3 <= 199.9)) and
                    (cSex = 2)     and
                    (((val4>=0.01) and (Val4<=1.20)) or (Val4=0)) and
                    ((sGulkwa[70] ='음성')  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')Then
                Begin
                     K := K + 1;
                     tSokun[K] := 'GAH1';
                End
                else If  ((Val1 <= 34) and (Val2 <= 32))  and
                    ((Val3 >= 110.1) and (Val3 <= 200)) and
                    (cSex = 2)     and
                    (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')) or((sGulkwa[70] ='') and (Val5=0)) or ((val5>0) and(val5 <9.99))) and
                    ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = '')) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                     K := K + 1;
                     tSokun[K] := 'GAH1';
                End
                else If  ((Val1 <= 40) and (Val2 <= 40)) and
                        (Val3 >= 200) and
                        (cSex = 1)    and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성'))  or ((sGulkwa[70] ='') or (sGulkwa[71] =''))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성')  or (sGulkwa[72] = ''))and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAHH';
                End
                else If  ((Val1 <= 45) and (Val2 <= 40)) and
                          (Val3 >= 200.1) and
                          (cSex = 1)    and
                          (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')  or ((val5>0) and (val5 <9.99)))  or (((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='')and (val6=0)))) and
                          ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                          ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAHH';
                End
                else If ((Val1 <= 39.9) and (Val2 <= 40)) and
                        (Val3 >= 200) and
                        (cSex = 2)    and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5>0) and (val5 <9.99)))  or (((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='')and (val6=0)))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAHH';
                End
                else If ((Val1 <= 34) and (Val2 <= 32)) and
                        (Val3 >= 200.1) and
                        (cSex = 2)    and
                        (((sGulkwa[70] ='음성') or (sGulkwa[70] ='약양성')or ((val5>0) and (val5 <9.99)))  or (((sGulkwa[70] ='')and (val5=0)) or ((sGulkwa[71] ='')and (val6=0)))) and
                        ((sGulkwa[72] = '음성') or (sGulkwa[72] = '약양성') or (sGulkwa[72] = ''))and
                        ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                        K := K + 1;
                        tSokun[K] := 'GAHH';
                End;
             End;

             //CPN
             IF (sGulkwa[128] <> '') and (sGulkwa[127] <> '')  and (sGulkwa[116] <> '') and (sGulkwa[129] <> '') and (sGulkwa[67] <> '') then
                Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val4 :=0;
                Val5 :=0;
                Val1 := StrToFloat(sGulkwa[128]);
                Val2 := StrToFloat(sGulkwa[127]);
                Val3 := StrToFloat(sGulkwa[116]);
                Val4 := StrToFloat(sGulkwa[129]);
                Val5 := StrToFloat(sGulkwa[67]);

                    if (((val1 >= 110) and (cSex = 1)) OR ((val1 >= 120) and (cSex = 2)))
                    and (((val2 <= 160) and (cSex = 1)) OR ((val2 <= 150) and (cSex = 2)))
                    and  (val3 <= 15) and (val4 <= 30) and (val5 <= 0.3) then
                    Begin
                      K := K + 1;
                      tSokun[K] := 'CPN ';
                    End
                end;



             If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '') and (sGulkwa[11] <> '')  Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val4 :=0;
                Val5 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);

                if (sGulkwa[17]) <> '' then   Val5 := StrToFloat(sGulkwa[17])
                else Val5 :=0;

                If  (Val1 <= 40)  and
                    (Val2 >= 40.1)and
                    (Val3 <= 50)  and
                    (Val5 <= 450) and
                    (cSex = 1)    and
                    (sGulkwa[70] ='음성') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else If  (Val1 <= 45)  and
                         (Val2 >= 40.1)and
                         (Val3 <= 70)  and
                         ((Val5 <= 260)or (Val5 = 0)) and
                         (cSex = 1)    and
                         ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else  If  (Val1 <= 39.9)  and
                          (Val2 >= 40.1)and
                          (Val3 <= 40)  and
                          (Val5 <= 450) and
                          (cSex = 2)    and
                          (sGulkwa[70] ='음성') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else  If  (Val1 <= 34)  and
                          (Val2 >= 32.1)and
                          (Val3 <= 42)  and
                          ((Val5 <= 260)or (Val5 = 0)) and
                          (cSex = 2)    and
                          ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
             End;
             If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '') and (sGulkwa[11] <> '') and (sGulkwa[17] <> '') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val4 :=0;
                Val5 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);
                Val5 := StrToFloat(sGulkwa[17]);


                If  (Val1 <= 45)  and
                    (Val2 >= 40.1)and
                    (Val3 <= 70)  and
                    (Val5 >= 260.1) and
                    (cSex = 1)      then
                Begin
                      K := K + 1;
                      tSokun[K] := 'ASL ';
                End
                else If  (Val1 <= 34)  and
                         (Val2 >= 32.1)and
                         (Val3 <= 42)  and
                         (Val5 >= 260.1) and
                         (cSex = 2)    Then
                Begin
                      K := K + 1;
                      tSokun[K] := 'ASL ';
                End
             End;


             If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '') and (sGulkwa[11] <> '') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val1 := StrToFloat(sGulkwa[10]);
                Val2 := StrToFloat(sGulkwa[09]);
                Val3 := StrToFloat(sGulkwa[11]);
                if (sGulkwa[55])<> '' then Val4 := StrToFloat(sGulkwa[55])
                else val4 :=0;

                If  (Val1 <= 40)  and
                    (Val2 >= 40.1)and
                    (Val3 <= 50)  and
                    (cSex = 1)    and
                    (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   then
                Begin
                      K := K + 1;
                      tSokun[K] := 'HBCS';
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else If  (Val1 <= 45)  and
                         (Val2 >= 40.1)and
                         (Val3 <= 70)  and
                         (cSex = 1)    and
                         ((sGulkwa[70] ='양성') or (val4>=10.0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'HBCS';
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else  If  (Val1 <= 39.9)  and
                          (Val2 >= 40.1)and
                          (Val3 <= 40)  and
                          (cSex = 2)    and
                          (sGulkwa[70] ='양성') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  then
                Begin
                      K := K + 1;
                      tSokun[K] := 'HBCS';
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
                else  If  (Val1 <= 34)  and
                          (Val2 >= 32.1)and
                          (Val3 <= 42)  and
                          (cSex = 2)    and
                          ((sGulkwa[70] ='양성') or (val4>=10.0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                Begin
                      K := K + 1;
                      tSokun[K] := 'HBCS';
                      K := K + 1;
                      tSokun[K] := 'AST ';
                End
             End;



             If (sGulkwa[130] <> '')  and (sGulkwa[42] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                Val1 := StrToFloat(sGulkwa[130]);
                if (sGulkwa[131]) <> '' then Val2 := StrToFloat(sGulkwa[131])
                else val2 :=0;
                Val3 := StrToFloat(sGulkwa[42]);
                If  ((Val1>=0.01)    and (Val1 <= 0.69))  and
                    (((Val2 >= 1.71) and (Val2 <= 3.71)) or (val2=0)) and
                    ((val3>=0.3) and (val3<=5.5))  and
                    ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124')     then
                Begin
                      K := K + 1;
                      tSokun[K] := 'TSH9';
                End
                else If  ((Val1>=0.01)    and (Val1 <= 0.69))  and
                         (((Val2 >= 2.3) and (Val2 <= 4.2)) or (val2=0)) and
                         ((val3>=0.3) and (val3<=5.5))  and
                         ((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125')     then
                Begin
                      K := K + 1;
                      tSokun[K] := 'TSH9';
                End
                else If (Val1 >= 1.861) and
                        (((Val2 >= 1.71) and (Val2 <= 3.71)) or (val2=0)) and
                        ((val3>=0.3) and (val3<=5.5)) and
                        ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124')  then
                Begin
                        K := K + 1;
                        tSokun[K] := 'TSH9';
                End
                else If (Val1 >= 1.861) and
                        (((Val2 >= 2.3) and (Val2 <= 4.2)) or (val2=0)) and
                        ((val3>=0.3) and (val3<=5.5)) and
                        (((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                         ((qryBunju.FieldByName('dat_bunju').AsString)<= '20180917')) then
                Begin
                        K := K + 1;
                        tSokun[K] := 'TSH9';
                End
                //2018.09.18
                else If   (Val1 > 1.76) and
                        (((Val2 >= 2.3) and (Val2 <= 4.2)) or (val2=0)) and
                         ((val3 >= 0.3) and (val3 <= 5.5)) and
                         ((qryBunju.FieldByName('dat_bunju').AsString)>= '20180918')  then
                Begin
                        K := K + 1;
                        tSokun[K] := 'TSH9';
                End
                else If   (Val1 < 0.84) and
                        (((Val2 >= 2.3) and (Val2 <= 4.2)) or (val2=0)) and
                         ((val3 >= 0.3) and (val3 <= 5.5)) and
                         ((qryBunju.FieldByName('dat_bunju').AsString)>= '20180918')  then
                Begin
                        K := K + 1;
                        tSokun[K] := 'TSH9';
                End

             End;

//수정 20140319 TSH2>TSH>THH
             If (sGulkwa[42] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[130]) <> '' then Val1 := StrToFloat(sGulkwa[130])
                else val1 :=0;
                if (sGulkwa[131]) <> '' then  Val2 := StrToFloat(sGulkwa[131])
                else val2 :=0;
                if (sGulkwa[42]) <> '' then  Val3 := StrToFloat(sGulkwa[42])
                else val3 :=0;
                if (sGulkwa[40]) <> '' then  Val4 := StrToFloat(sGulkwa[40])
                else val4 :=0;
                if (sGulkwa[41]) <> '' then  Val5 := StrToFloat(sGulkwa[41])
                else val5 :=0;

                If  (((Val1 >= 0.7) and (Val1<=1.86)) or (val1=0))  and
                    (((Val2 >= 1.71)and (Val2 <= 3.71)) or (Val2=0)) and
                    ((val3>=0.001)  and (val3<=0.299)) and
                    (((Val4 >= 0.6) and (Val4<=2.1)) or (val4=0))  and
                    (((Val5 >= 4.5) and (Val5<=12.0)) or (val5=0))  and
                    ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH2';
                End
                else If  (((Val1 >= 0.7) and (Val1<=1.86)) or (val1=0))  and
                         (((Val2 >= 2.3)and (Val2 <= 4.2)) or (Val2=0)) and
                         ((val3>=0.001)  and (val3<=0.299)) and
                         (((Val4 >= 0.6) and (Val4<=2.1)) or (val4=0))  and
                         (((Val5 >= 4.5) and (Val5<=12.0)) or (val5=0))  and
                         (((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                          ((qryBunju.FieldByName('dat_bunju').AsString) < '20150109')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH2';
                End
                else If  (((Val1 >= 0.7) and (Val1<=1.86))  or (val1=0))  and
                         (((Val2 >= 2.3) and (Val2 <= 4.2)) or (Val2=0))  and
                          ((val3>=0.001) and (val3<=0.299)) and
                         (((Val4 >= 0.8) and (Val4<=2.0))   or (val4=0))  and
                         (((Val5 >= 4.5) and (Val5<=12.0))  or (val5=0))  and
                         (((qryBunju.FieldByName('dat_bunju').AsString) >= '20150109') and
                          ((qryBunju.FieldByName('dat_bunju').AsString) <  '20180918'))then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH2';
                End
                //2018.09.18
                else If  (((Val1 >= 0.84)  and (Val1<=1.76))   or (val1=0))  and
                         (((Val2 >= 2.3)   and (Val2 <= 4.2))  or (Val2=0))  and
                           (val3 <= 0.299) and
                         (((Val4 >= 0.8)   and (Val4 <= 2.0))  or (val4=0))  and
                         (((Val5 >= 4.5)   and (Val5 <= 12.0)) or (val5=0))  and
                          ((qryBunju.FieldByName('dat_bunju').AsString) >= '20180918')then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH2';
                End


                else If  (((Val1 >= 0.7) and (Val1<=1.86)) or (val1=0))  and
                         (((Val2 >= 1.71)and (Val2 <= 3.71)) or (Val2=0)) and
                         (val3>=5.501)    and
                         (((Val4 >= 0.6) and (Val4<=2.1)) or (val4=0))  and
                         (((Val5 >= 4.5) and (Val5<=12.0)) or (val5=0)) and
                         ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH ';
                End
                else If  (((Val1 >= 0.7) and (Val1<=1.86)) or (val1=0))  and
                         (((Val2 >= 2.3)and (Val2 <= 4.2)) or (Val2=0)) and
                         (val3>=5.501)    and
                         (((Val4 >= 0.6) and (Val4<=2.1)) or (val4=0))  and
                         (((Val5 >= 4.5) and (Val5<=12.0)) or (val5=0)) and
                         (((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                         ((qryBunju.FieldByName('dat_bunju').AsString) < '20150109')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH ';
                End
                else If  (((Val1 >= 0.7) and (Val1<=1.86)) or (val1=0))  and
                         (((Val2 >= 2.3)and (Val2 <= 4.2)) or (Val2=0)) and
                         (val3>=5.501)    and
                         (((Val4 >= 0.8) and (Val4<=2.0)) or (val4=0))  and
                         (((Val5 >= 4.5) and (Val5<=12.0)) or (val5=0)) and
                         (((qryBunju.FieldByName('dat_bunju').AsString) >= '20150109') and
                          ((qryBunju.FieldByName('dat_bunju').AsString) <  '20180918'))then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH ';
                End
                //2018.09.18
                else If  (((Val1 >= 0.84) and (Val1 <= 1.76)) or (val1=0))  and
                         (((Val2 >= 2.3)  and (Val2 <= 4.2))  or (Val2=0))  and
                           (val3>=5.501)  and
                         (((Val4 >= 0.8)  and (Val4 <= 2.0))  or (val4=0))  and
                         (((Val5 >= 4.5)  and (Val5 <= 12.0)) or (val5=0)) and
                         ((qryBunju.FieldByName('dat_bunju').AsString) >= '20180918') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH ';
                End;
             End;

//20140319 - TSH2>TSH>THH

             If (tSokun[K] <> 'TSH2') And
                (tSokun[K] <> 'TSH ')  Then
             Begin
                   If  (sGulkwa[42] <> '') and ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124') Then
                   Begin
                      val1 :=0;
                      val2 :=0;
                      val3 :=0;
                      val4 :=0;
                      val5 :=0;
                      if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                      else val1 :=0;
                      if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                      else val2 :=0;
                      if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                      else val3 :=0;
                      if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                      else val4 :=0;
                      Val5 := StrToFloat(sGulkwa[42]);

                      If ((Val1>=0.6)  or (Val1=0)) and
                         ((Val2>=4.5)  or (Val2=0)) and
                         ((Val3>=1.71) or (Val3=0)) and
                         (Val4>=1.861) and
                         ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.6) or (Val1=0))  and
                              ((Val2>=4.5) or (Val2=0))  and
                              (Val3>=3.711) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.6) or (Val1=0))  and
                              (Val2>=12.01) and
                              ((Val3>=1.71) or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If (Val1>=2.11)   and
                              ((Val2>=4.5) or (Val2=0)) and
                              ((Val3>=1.71)or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End;
                   End
                   else If  (sGulkwa[42] <> '') and ((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                            ((qryBunju.FieldByName('dat_bunju').AsString) < '20150109') then
                   Begin
                      val1 :=0;
                      val2 :=0;
                      val3 :=0;
                      val4 :=0;
                      val5 :=0;
                      if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                      else val1 :=0;
                      if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                      else val2 :=0;
                      if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                      else val3 :=0;
                      if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                      else val4 :=0;
                      Val5 := StrToFloat(sGulkwa[42]);

                      If ((Val1>=0.6) or (Val1=0))  and
                         ((Val2>=4.5) or (Val2=0))  and
                         ((Val3>=2.3) or (Val3=0))  and
                         (Val4>=1.861) and
                         ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.6) or (Val1=0))  and
                              ((Val2>=4.5) or (Val2=0))  and
                              (Val3>=2.3) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.6) or (Val1=0))  and
                              (Val2>=12.01) and
                              ((Val3>=2.3) or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If (Val1>=2.11)   and
                              ((Val2>=4.5) or (Val2=0)) and
                              ((Val3>=2.3)or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End;
                   End
                   else If  (sGulkwa[42] <> '') and
                          (((qryBunju.FieldByName('dat_bunju').AsString)>= '20150109') and ((qryBunju.FieldByName('dat_bunju').AsString)<= '20180917')) then
                   Begin
                      val1 :=0;
                      val2 :=0;
                      val3 :=0;
                      val4 :=0;
                      val5 :=0;
                      if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                      else val1 :=0;
                      if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                      else val2 :=0;
                      if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                      else val3 :=0;
                      if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                      else val4 :=0;
                      Val5 := StrToFloat(sGulkwa[42]);

                      If ((Val1>=0.8) or (Val1=0))  and
                         ((Val2>=4.5) or (Val2=0))  and
                         ((Val3>=2.3) or (Val3=0))  and
                         (Val4>=1.861) and
                         ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.8) or (Val1=0))  and
                              ((Val2>=4.5) or (Val2=0))  and
                              (Val3>=2.3) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.8) or (Val1=0))  and
                              (Val2>=12.01) and
                              ((Val3>=2.3) or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If (Val1>=1.91)   and
                              ((Val2>=4.5) or (Val2=0)) and
                              ((Val3>=2.3)or (Val3=0)) and
                              ((Val4>=0.7) or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End;
                   End
                   //2018.09.18
                   else If  (sGulkwa[42] <> '') and
                           ((qryBunju.FieldByName('dat_bunju').AsString)>= '20180918') then
                   Begin
                      val1 :=0;
                      val2 :=0;
                      val3 :=0;
                      val4 :=0;
                      val5 :=0;
                      if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                      else val1 :=0;
                      if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                      else val2 :=0;
                      if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                      else val3 :=0;
                      if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                      else val4 :=0;
                      Val5 := StrToFloat(sGulkwa[42]);

                      If ((Val1>=0.8) or (Val1=0))  and
                         ((Val2>=4.5) or (Val2=0))  and
                         ((Val3>=2.3) or (Val3=0))  and
                          (Val4 >1.76) and
                         ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.8)  or (Val1=0))  and
                              ((Val2>=4.5)  or (Val2=0))  and
                               (Val3 > 4.2) and
                              ((Val4>=0.84) or (Val4=0)) and
                              ((Val5>=5.501)or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If ((Val1>=0.8)   or (Val1=0))  and
                               (Val2>=12.01) and
                              ((Val3>=2.3)   or (Val3=0)) and
                              ((Val4>=0.84)  or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End
                      else If (Val1> 2.0)   and
                              ((Val2>=4.5)   or (Val2=0)) and
                              ((Val3>=2.3)   or (Val3=0)) and
                              ((Val4>=0.84)  or (Val4=0)) and
                              ((Val5>=5.501) or (Val5<=0.299)) then
                      begin
                          K := K + 1;
                          tSokun[K] := 'THH ';
                      End;
                   End;

             End;

//End 20140319 -

             If (sGulkwa[42] <> '')  and ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                else val1 :=0;
                if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                else val2 :=0;
                if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                else val3 :=0;
                if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                else val4 :=0;
                Val5 := StrToFloat(sGulkwa[42]);

                If ((Val1<=2.1) or (Val1=0)) and
                   ((Val2<=12.0)or (Val2=0))and
                   ((Val3<=3.71)or (Val3=0))  and
                   ((Val4>=0.001) and(Val4<=0.69)) and
                   ((Val5>=5.501) or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.1) or (Val1=0)) and
                        ((Val2<=12.0) or (Val2=0))and
                        ((Val3>=0.001)and (Val3<=1.709)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.1) or (Val1=0)) and
                        ((Val2>=0.001) and (Val2<=4.49))  and
                        ((Val3<=3.71) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1>=0.001) and (Val1<=0.59))  and
                        ((Val2<=12.0) or (Val2=0)) and
                        ((Val3<=3.71) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End;
             End
             else If (sGulkwa[42] <> '')  and ((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                     ((qryBunju.FieldByName('dat_bunju').AsString) <'20150109') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                else val1 :=0;
                if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                else val2 :=0;
                if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                else val3 :=0;
                if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                else val4 :=0;
                Val5 := StrToFloat(sGulkwa[42]);

                If ((Val1<=2.1) or (Val1=0)) and
                   ((Val2<=12.0)or (Val2=0))and
                   ((Val3<=4.2)or (Val3=0))  and
                   ((Val4>=0.001) and(Val4<=0.69)) and
                   ((Val5>=5.501) or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.1) or (Val1=0)) and
                        ((Val2<=12.0) or (Val2=0))and
                        ((Val3>=0.001)and (Val3 < 2.3)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.1) or (Val1=0)) and
                        ((Val2>=0.001) and (Val2<=4.49))  and
                        ((Val3<=4.2) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1>=0.001) and (Val1<=0.59))  and
                        ((Val2<=12.0) or (Val2=0)) and
                        ((Val3<=4.2) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End;
             End
             else If (sGulkwa[42] <> '')  and
                   (((qryBunju.FieldByName('dat_bunju').AsString)>='20150109') and
                    ((qryBunju.FieldByName('dat_bunju').AsString)<='20180917'))Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                else val1 :=0;
                if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                else val2 :=0;
                if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                else val3 :=0;
                if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                else val4 :=0;
                Val5 := StrToFloat(sGulkwa[42]);

                If ((Val1<=1.9) or (Val1=0)) and
                   ((Val2<=12.0)or (Val2=0))and
                   ((Val3<=4.2) or (Val3=0))  and
                   ((Val4>=0.001) and(Val4<=0.69)) and
                   ((Val5>=5.501) or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=1.9) or (Val1=0)) and
                        ((Val2<=12.0) or (Val2=0))and
                        ((Val3>=0.001)and (Val3 < 2.3)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=1.9) or (Val1=0)) and
                        ((Val2>=0.001) and (Val2<=4.49))  and
                        ((Val3<=4.2) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1>=0.001) and (Val1<=0.79))  and
                        ((Val2<=12.0) or (Val2=0)) and
                        ((Val3<=4.2) or (Val3=0)) and
                        ((Val4<=1.86) or (Val4=0)) and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End;
             End
             //2018.09.18
             else If (sGulkwa[42] <> '')  and
                    ((qryBunju.FieldByName('dat_bunju').AsString)>='20180918') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[40]) <> ''  then Val1 := StrToFloat(sGulkwa[40])
                else val1 :=0;
                if (sGulkwa[41]) <> ''  then Val2 := StrToFloat(sGulkwa[41])
                else val2 :=0;
                if (sGulkwa[131]) <> '' then Val3 := StrToFloat(sGulkwa[131])
                else val3 :=0;
                if (sGulkwa[130]) <> '' then Val4 := StrToFloat(sGulkwa[130])
                else val4 :=0;
                Val5 := StrToFloat(sGulkwa[42]);

                If ((Val1<=2.0)  or (Val1=0)) and
                   ((Val2<=12.0) or (Val2=0))and
                   ((Val3<=4.2)  or (Val3=0))  and
                   ((Val4>=0.001)and(Val4<0.84)) and
                   ((Val5>=5.501) or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.0)  or  (Val1=0))  and
                        ((Val2<=12.0) or  (Val2=0))  and
                        ((Val3>=0.001)and (Val3<2.3))and
                        ((Val4<=1.76) or  (Val4=0))  and
                        ((Val5>=5.501)or  (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1<=2.0)   or (Val1=0))     and
                        ((Val2>=0.001) and(Val2<=4.49)) and
                        ((Val3<=4.2)   or (Val3=0))     and
                        ((Val4<=1.76)  or (Val4=0))     and
                        ((Val5>=5.501) or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End
                else If ((Val1>=0.001) and(Val1<0.8))and
                        ((Val2<=12.0)  or (Val2=0))    and
                        ((Val3<=4.2)   or (Val3=0))    and
                        ((Val4<=1.76) or  (Val4=0))    and
                        ((Val5>=5.501)or (Val5<=0.299)) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End;
             End;



             If (sGulkwa[131] <> '')  and (sGulkwa[42] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                if (sGulkwa[130]) <> '' then Val1 := StrToFloat(sGulkwa[130])
                else val1 :=0;
                Val2 := StrToFloat(sGulkwa[131]);
                Val3 := StrToFloat(sGulkwa[42]);

                If  (((Val1 >= 0.7)  and (Val1 <= 1.86)) or (Val1=0)) and
                    ((Val2 >= 3.711)  or ((Val2 >=0.001) and (Val2 <= 1.709)))and
                    ((val3>=0.3) and (val3<=5.5))        and
                    ((qryBunju.FieldByName('dat_bunju').AsString)<= '20140124') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHC';
                End
                else If  (((Val1 >= 0.7) and (Val1 <= 1.86)) or (Val1=0)) and
                         ((Val2 > 4.2)  or ((Val2 >=0.001) and (Val2 <2.3)))and
                         ((val3>=0.3) and (val3<=5.5))        and
                        (((qryBunju.FieldByName('dat_bunju').AsString)>= '20140125') and
                         ((qryBunju.FieldByName('dat_bunju').AsString)<= '20180917')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHC';
                End
                //2018.09.18
                else If  (((Val1 >= 0.84) and (Val1 <= 1.76)) or  (Val1=0)) and
                          ((Val2 >  4.2)  or ((Val2 >= 0.001) and (Val2 <2.3))) and
                          ((val3 >= 0.3)  and (val3 <= 5.5))        and
                          ((qryBunju.FieldByName('dat_bunju').AsString)>= '20180918') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHC';
                End;
             End;


             If (sGulkwa[42] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                if (sGulkwa[130]) <> '' then Val1 := StrToFloat(sGulkwa[130])
                else val1 :=0;
                if (sGulkwa[131]) <> '' then  Val2 := StrToFloat(sGulkwa[131])
                else val2 :=0;
                Val3 := StrToFloat(sGulkwa[42]);


                If  ((Val1 >= 0.001) and (Val1 <= 0.69))  and
                    (((Val2 >= 1.71)  and (Val2 <= 3.71)) or (Val2=0)) and
                    ((val3>=5.01) or (val3<=0.29)) and
                    ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121013') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'THL ';
                End

                else If  ((Val1 >= 0.001) and (Val1 <= 0.69)) and
                         ((Val2 >= 0.001) and (Val2 <= 1.709)) and
                         ((val3>=5.01) or  (val3<=0.29))  and
                         ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121013')then
                Begin
                         K := K + 1;
                         tSokun[K] := 'THL ';
                End

                else If (((Val1 <= 1.86) and (Val1 >= 0.7)) or (Val1=0)) and
                        ((Val2 >= 0.001) and (Val2 <= 1.709)) and
                        ((val3>=5.01) or  (val3<=0.29)) and
                        ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121013') then
                Begin
                         K := K + 1;
                         tSokun[K] := 'THL ';
                End;

             End;


             If ((sGulkwa[130] <> '') or (sGulkwa[131] <> '')  or  (sGulkwa[41] <> '') or (sGulkwa[40] <> '')) and (sGulkwa[42] ='') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;

                if (sGulkwa[41]) <> '' then Val1 := StrToFloat(sGulkwa[41])
                else val1 :=0;
                if (sGulkwa[40]) <> '' then  Val2 := StrToFloat(sGulkwa[40])
                else val2 :=0;
                if (sGulkwa[130]) <> '' then  Val3 := StrToFloat(sGulkwa[130])
                else val3 :=0;
                if (sGulkwa[131]) <> '' then  Val4 := StrToFloat(sGulkwa[131])
                else val4 :=0;

                If  ((Val1 >= 12.01) or
                     (Val2 >= 2.11)  or
                     (Val3 >= 1.861) or
                     (Val4 >= 3.711)) and ((qryBunju.FieldByName('Dat_bunju').AsString)<= '20140124') then
                begin
                    K := K + 1;
                    tSokun[K] := 'THH1';
                End
                else If  (((Val1 >=0.001) and (Val1 <= 4.49)) or
                          ((Val2>=0.001)  and (Val2 <= 0.59)) or
                          ((Val3>=0.001)  and (Val3 <= 0.69)) or
                          ((Val4>=0.001)  and (Val4<= 1.709))) and ((qryBunju.FieldByName('Dat_bunju').AsString)<= '20140124') Then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL1';
                End
                else If  ((Val1 >= 12.01) or
                           (Val2 >= 2.11) or
                           (Val3 >= 1.861)or
                           (Val4 > 4.2)) and
                          (((qryBunju.FieldByName('Dat_bunju').AsString)>= '20140125') and
                           ((qryBunju.FieldByName('Dat_bunju').AsString) < '20150109')) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THH1';
                End
                else If  (((Val1 >=0.001) and (Val1 <= 4.49)) or
                          ((Val2>=0.001) and  (Val2 <= 0.59)) or
                          ((Val3>=0.001) and  (Val3 <= 0.69)) or
                          ((Val4>=0.001) and  (Val4 < 2.3))) and
                         (((qryBunju.FieldByName('Dat_bunju').AsString)>= '20140125') and
                          ((qryBunju.FieldByName('Dat_bunju').AsString) < '20150109')) Then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL1';
                End
                else If  ((Val1 >= 12.01) or
                          (Val2 >= 2.01)  or
                          (Val3 >= 1.861) or
                          (Val4 > 4.2)) and
                         (((qryBunju.FieldByName('Dat_bunju').AsString) >= '20150109') and
                          ((qryBunju.FieldByName('Dat_bunju').AsString) <= '20180917')) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THH1';
                End
                else If  (((Val1 >=0.001) and (Val1 <= 4.49)) or
                          ((Val2>=0.001)  and (Val2 <= 0.79)) or
                          ((Val3>=0.001)  and (Val3 <= 0.69)) or
                          ((Val4>=0.001)  and (Val4 < 2.3))) and
                         (((qryBunju.FieldByName('Dat_bunju').AsString) >= '20150109') and
                          ((qryBunju.FieldByName('Dat_bunju').AsString) <= '20180917')) then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL1';
                End
                //2018.09.18
                else If  ((Val1 >= 12.01) or
                          (Val2 > 2.0)   or
                          (Val3 > 1.76)   or
                          (Val4 > 4.2))   and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20180918') then
                begin
                    K := K + 1;
                    tSokun[K] := 'THH1';
                End
                else If  (((Val1 >=0.001) and (Val1 <= 4.49)) or
                          ((Val2>=0.001)  and (Val2 < 0.8)) or
                          ((Val3>=0.001)  and (Val3 < 0.84)) or
                          ((Val4>=0.001)  and (Val4 < 2.3))) and
                          ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20180918') Then
                begin
                    K := K + 1;
                    tSokun[K] := 'THL1';
                End
             End;
             If (sGulkwa[140] <> '') and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422') Then
             Begin
                val1 :=0;
                Val1 := StrToFloat(sGulkwa[140]);
                If  (Val1  < 10) Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'VTDD';
                End
                else If  ((Val1  >= 10) and (Val1  <= 30))   Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'VTDI';
                End
                else If  (Val1  >100)   Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'VTDT';
                End
             End;

             If (sGulkwa[42] <> '') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[131]) <> '' then  Val1 := StrToFloat(sGulkwa[131]) //T042
                else val1 :=0;
                if (sGulkwa[130]) <> '' then Val2 := StrToFloat(sGulkwa[130])  //E015
                else val2 :=0;

                if (sGulkwa[42]) <> '' then  Val3 := StrToFloat(sGulkwa[42])   //E003
                else val3 :=0;
                if (sGulkwa[40]) <> '' then  Val4 := StrToFloat(sGulkwa[40])   //E001
                else val4 :=0;
                if (sGulkwa[41]) <> '' then  Val5 := StrToFloat(sGulkwa[41])  //E002
                else val5 :=0;

                If  (((Val2 >= 0.7) and (Val2<=1.86)) or (val2=0))  and
                    (((Val1 >= 1.71)and (Val1 <= 3.71)) or (Val1=0)) and
                    ((val3>=0.3)  and (val3<=5.5)) and
                    ((Val5 >= 12.01) or  ((Val5>=0.01) and  (val5<=4.49)))  and
                    (((Val4>=0.6) and  (Val4<=2.1)) or (Val4=0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)<='20140124')   then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH7';
                End
                else If  (((Val2 >= 0.7) and (Val2<=1.86))  or (val2=0)) and
                         (((Val1 >= 2.3) and (Val1 <= 4.2)) or (Val1=0)) and
                         ((val3>=0.3)    and (val3<=5.5))  and
                         ((Val5 >= 12.01) or ((Val5>=0.01) and  (val5<=4.49))) and
                         (((Val4>=0.6)   and (Val4<=2.1))  or (Val4=0)) and
                         (((qryBunju.FieldByName('Dat_bunju').AsString)>='20140125') and ((qryBunju.FieldByName('Dat_bunju').AsString)<'20150109'))   then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH7';
                End
                else If  (((Val2 >= 0.7) and (Val2<=1.86))  or (val2=0)) and
                         (((Val1 >= 2.3) and (Val1 <= 4.2)) or (Val1=0)) and
                         ((val3>=0.3)    and (val3<=5.5))  and
                         ((Val5 >= 12.01) or ((Val5>=0.01) and (val5<=4.49))) and
                         (((Val4>=0.8)   and (Val4<=1.9))  or (Val4=0)) and
                         (((qryBunju.FieldByName('Dat_bunju').AsString)>='20150109') and
                          ((qryBunju.FieldByName('Dat_bunju').AsString)>='20180917')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH7';
                End
                //2018.09.18
                else If  (((Val2 >= 0.84) and  (Val2 <= 1.76)) or  (val2=0)) and
                         (((Val1 >= 2.3)  and  (Val1 <= 4.2))  or  (Val1=0)) and
                          ((val3 >= 0.3)  and  (val3 <= 5.5))  and
                           (Val5 >= 12.01) and
                         (((Val4 >= 0.8)  and  (Val4 <=2.0))  or (Val4=0)) and
                         ((qryBunju.FieldByName('Dat_bunju').AsString) >= '20180918') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH7';
                End
                else If  (((Val2 >= 0.84) and  (Val2 <= 1.76)) or  (val2=0)) and
                         (((Val1 >= 2.3)  and  (Val1 <= 4.2))  or  (Val1=0)) and
                          ((val3 >= 0.3)  and  (val3 <= 5.5))  and
                          ((Val5 >= 0.01) and  (val5 <= 4.49)) and
                         (((Val4 >= 0.8)  and  (Val4 <=2.0))  or (Val4=0)) and
                         ((qryBunju.FieldByName('Dat_bunju').AsString) >= '20180918') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSH7';
                End;

             End;
             If (sGulkwa[42] <> '') Then
             Begin
                val1 :=0;                                                                           
                val2 :=0;
                val3 :=0;
                val4 :=0;
                val5 :=0;
                if (sGulkwa[131]) <> '' then  Val1 := StrToFloat(sGulkwa[131])   //T042
                else val1 :=0;
                if (sGulkwa[130]) <> '' then Val2 := StrToFloat(sGulkwa[130])    //E015
                else val2 :=0;

                if (sGulkwa[42]) <> '' then  Val3 := StrToFloat(sGulkwa[42])     //E003
                else val3 :=0;
                if (sGulkwa[40]) <> '' then  Val4 := StrToFloat(sGulkwa[40])     //E001
                else val4 :=0;
                if (sGulkwa[41]) <> '' then  Val5 := StrToFloat(sGulkwa[41])    //E002
                else val5 :=0;

                If  (((Val1 >= 1.71) and (Val1<=3.71))    or (val1=0)) and
                     (((Val2 >= 0.7)  and (Val2 <= 1.86)) or (Val2=0)) and
                     ((val3>=0.3)  and (val3<=5.5)) and
                     (((Val5 >= 4.5)  and (Val5<=12.0))   or (Val5=0)) and
                     (((Val4 >=0.01) and  (Val4<=0.59))   or( Val4>=2.11)) and  ((qryBunju.FieldByName('Dat_bunju').AsString)<='20140124') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHB';
                End
                else If  (((Val1 >= 2.3) and (Val1<=4.2)) or (val1=0)) and
                     (((Val2 >= 0.7)  and (Val2 <= 1.86)) or (Val2=0)) and
                     ((val3>=0.3)  and (val3<=5.5)) and
                     (((Val5 >= 4.5)  and (Val5<=12.0))   or (Val5=0)) and
                     (((Val4 >=0.01) and  (Val4<=0.59))   or( Val4>=2.11)) and
                     (((qryBunju.FieldByName('Dat_bunju').AsString)>='20140125') and ((qryBunju.FieldByName('Dat_bunju').AsString)<'20150109')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHB';
                End
                else If  (((Val1 >= 2.3) and (Val1<=4.2)) or (val1=0)) and
                     (((Val2 >= 0.7)  and (Val2 <= 1.86)) or (Val2=0)) and
                     ((val3>=0.3)  and (val3<=5.5)) and
                     (((Val5 >= 4.5)  and (Val5<=12.0))   or (Val5=0)) and
                     (((Val4 >=0.01) and  (Val4<=0.79))   or( Val4>=2.01)) and
                     (((qryBunju.FieldByName('Dat_bunju').AsString)>='20150109') and
                      ((qryBunju.FieldByName('Dat_bunju').AsString)<='20180917')) then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHB';
                End
                //2018.09.18
                else If  (((Val1 >= 2.3)  and (Val1 <= 4.2))  or (val1=0)) and
                         (((Val2 >= 0.84) and (Val2 <= 1.76)) or (Val2=0)) and
                          ((val3 >= 0.3)  and (val3 <= 5.5))  and
                         (((Val5 >= 4.5)  and (Val5 <= 12.0)) or (Val5=0)) and
                         (((Val4 >= 0.01) and (Val4 < 0.8)) or (Val4 > 2.0)) and
                          ((qryBunju.FieldByName('Dat_bunju').AsString)>='20180918') then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHB';
                End;

             End;

             If (sGulkwa[130] <> '') and (sGulkwa[131] <> '')  and (sGulkwa[42] <> '')Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                if (sGulkwa[130]) <> '' then Val1 := StrToFloat(sGulkwa[130])
                else val1 :=0;
                if (sGulkwa[131]) <> '' then  Val2 := StrToFloat(sGulkwa[131])
                else val2 :=0;
                if (sGulkwa[42]) <> '' then  Val3 := StrToFloat(sGulkwa[42])
                else val3 :=0;

                If  ((Val1 <= 0.69) or (Val1 >= 1.861)) and
                    ((Val2 >= 3.711)or (Val2 <= 1.709)) and
                    ((val3 >=0.3)  and (val3 <=5.5)) and ((qryBunju.FieldByName('Dat_bunju').AsString)<='20140124')then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHM';
                End
                else If  ((Val1 <= 0.69) or (Val1 >= 1.861)) and
                    ((Val2 > 4.2)or (Val2 < 2.3)) and
                    ((val3 >=0.3)  and (val3 <=5.5)) and
                    (((qryBunju.FieldByName('Dat_bunju').AsString)>= '20140125') and
                     ((qryBunju.FieldByName('Dat_bunju').AsString)<= '20180917'))then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHM';
                End
                //2018.09.18
                else If  ((Val1 <  0.84) or  (Val1 >  1.76)) and
                         ((Val2 >  4.2)  or  (Val2 <  2.3)) and
                         ((val3 >= 0.3)  and (val3 <= 5.5)) and
                         ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20180918')then
                Begin
                    K := K + 1;
                    tSokun[K] := 'TSHM';
                End

             End;

             If (sGulkwa[128] <> '') Then
             Begin
                val1 :=0;
                val2 :=0;
                val3 :=0;
                if (sGulkwa[128]) <> '' then Val1  := StrToFloat(sGulkwa[128])
                else val1 :=0;
                if (sGulkwa[127]) <> '' then  Val2 := StrToFloat(sGulkwa[127])
                else val2 :=0;
                if (sGulkwa[129]) <> '' then  Val3 := StrToFloat(sGulkwa[129])
                else val3 :=0;

                If  (Val1 >= 158.01) and
                    (((Val2 <= 98.0)  and (Val2>=68)) or (Val2 = 0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'APB1';
                End
                else If  (Val1 > 200) and (cSex = 1) and
                    (((Val2 <= 160)  and (Val2>=70)) or (Val2 = 0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'APB1';
                End
                else If  (Val1 > 220) and (cSex = 2) and
                    (((Val2 <= 150.0)  and (Val2>=60)) or (Val2 = 0)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                Begin
                    K := K + 1;
                    tSokun[K] := 'APB1';
                End;
             End;
          {   If (sGulkwa[128] <> '') Then
             Begin
                val1 :=0;
                val2 :=0;

                if (sGulkwa[128]) <> '' then Val1  := StrToFloat(sGulkwa[128])
                else val1 :=0;
                if (sGulkwa[127]) <> '' then  Val2 := StrToFloat(sGulkwa[127])
                else val2 :=0;
                If  (Val1  <= 109.9) and
                    ((Val2 <= 98.0)  and (Val2 >= 0))  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'APA ';
                    End
                else If  (Val1  < 104) and (cSex = 1) and
                    ((Val2 <= 133.0)  and (Val2 >= 0))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'APA ';
                    End
                 else If  (Val1  < 108) and (cSex = 2) and
                    ((Val2 <= 117.0)  and (Val2 >= 0))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'APA ';
                    End;
             End;     }

             If (sGulkwa[138] <> '') Then
             Begin
                val1 :=0;
                Val1 := StrToFloat(sGulkwa[138]);
                If  (Val1  >= 360.01) then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'TFRH';
                    End
                else If  (Val1  <= 199.9) then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'TFRL';
                    End

             End;

             If (sGulkwa[139] <> '') Then
             Begin
                val1 :=0;
                Val1 := StrToFloat(sGulkwa[139]);
                If  (Val1  >= 6.91) Then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'CA72';
                    End
             End;



              If (sGulkwa[140] <> '') and  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130422')Then
             Begin
                val1 :=0;
                Val1 := StrToFloat(sGulkwa[140]);
                If  (Val1  <= 9.39) Then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'VTD ';
                    End
                else If  (Val1  >= 52.41) and (cSex = 1)  Then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'VTD1';
                    End
                else If  (Val1  >= 59.11) and (cSex = 2)  Then
                    Begin
                    K := K + 1;
                    tSokun[K] := 'VTD1';
                    End
             End;

       
//84-1 .Cod = HAB2 (S007 = 약양성 And S008 = 약양성) 2003.01.03
      If (sGulkwa[70] <> '') And (sGulkwa[71] <> '') Then
         Begin
            if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
            if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
               else val2 :=0;
            If ((sGulkwa[70] = '약양성') or ((val1>=1.0 )and  (val1<=9.99)))And
               ((sGulkwa[71] = '약양성') or ((val2>=10.0 )and  (val2<=29.9)))Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HAB2';
               End;
         End;
//22.Cod = LDH (C021 >= 500) And (C009,C010 <= 40) And (C011 <= 50)
      If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
         (sGulkwa[11] <> '') And (sGulkwa[17] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[17]);
            If (Val4 >= 500) And (Val1 <= 40) And
               (Val2 <= 40)  And (Val3 <= 50) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LDH ';
               End
            else  If (Val4 >= 260.1) And (Val1 <= 40) And
               (Val2 <= 45)  And (Val3 <= 70) and (cSex = 1) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LDH ';
               End
            else  If (Val4 >= 350.1) And (Val1 <= 40) And
               (Val2 <= 45)  And (Val3 >= 70.1) and (cSex = 1) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LDH ';
               End
            else  If (Val4 >= 260.1) And (Val1 <= 32) And
               (Val2 <= 34)  And (Val3 <= 42)and (cSex = 2) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LDH ';
               End
            else  If (Val4 >= 350) And (Val1 <= 32) And
               (Val2 <= 34)  And (Val3 >= 42.1)and (cSex = 2) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LDH ';
               End;
         End;
//13.Cod = BIH (C005 >= 2.0) And C009,C010 <= 40 And C011 <= 50))
    {  If (sGulkwa[05] <> '') And (sGulkwa[09] <> '') And
         (sGulkwa[10] <> '') And (sGulkwa[11] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[05]);
            Val2 := StrToFloat(sGulkwa[09]);
            Val3 := StrToFloat(sGulkwa[10]);
            Val4 := StrToFloat(sGulkwa[11]);
            If (Val1 >= 2.01) And (Val4 <= 50) And
               (Val2 <= 40) And (Val3 <= 40)  and
               (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End
            else If (Val1 >= 2.01) And (Val4 <= 70) And
                    (Val3 <= 45)   and (Val2 <= 40) and
               (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End
            else If (Val1 >= 2.01) And (Val4 <= 40) And
               (Val2 <= 40) And (Val3 <= 39.9)  and
               (cSex = 2) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End
            else If (Val1 >= 2.01) And (Val4 <= 42) And
                    (Val3 <= 34)  and (Val2 <= 32) and
               (cSex = 2) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End;
         End;  }

         If (sGulkwa[05] <> '')And     (sGulkwa[10] <> '') And (sGulkwa[11] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[05]);
            Val3 := StrToFloat(sGulkwa[10]);
            Val4 := StrToFloat(sGulkwa[11]);
            If (Val1 >= 2.01) And (Val4 <= 70) And
                    (Val3 <= 45)   and
               (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End
            else  If (Val1 >= 2.01) And (Val4 <= 42) And
                    (Val3 <= 34)   and
               (cSex = 2) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BIH ';
               End;
         End;

//12.Cod = BIM ((1.4 >= C005 <= 1.9) And C009,C010 <= 40 And C011 <= 50))
//      If tSokun[K] <> 'BIH' Then
     {    If (sGulkwa[05] <> '') And (sGulkwa[09] <> '') And
            (sGulkwa[10] <> '') And (sGulkwa[11] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[05]);
               Val2 := StrToFloat(sGulkwa[09]);
               Val3 := StrToFloat(sGulkwa[10]);
               Val4 := StrToFloat(sGulkwa[11]);
               If (Val1 >= 1.21) And (Val1 <= 2.0) And
                  (Val2 <= 40) And (Val3 <= 40) And
                  (Val4 <= 50) and (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End
               else If (Val1 >= 1.21) And (Val1 <= 2.0) And
                   (Val3 <= 45) And  (Val2 <= 40) and
                  (Val4 <= 70) and (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End
               else  If (Val1 >= 1.21) And (Val1 <= 2.0) And
                  (Val2 <= 40) And (Val3 <= 39.9) And
                  (Val4 <= 40) and (cSex = 2)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End
                else  If (Val1 >= 1.21) And (Val1 <= 2.0) And
                   (Val3 <= 34) And  (Val2 <= 32) and
                  (Val4 <= 42) and (cSex = 2)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End;
            End;    }

          If (sGulkwa[05] <> '') And       (sGulkwa[10] <> '') And (sGulkwa[11] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[05]);
               Val3 := StrToFloat(sGulkwa[10]);
               Val4 := StrToFloat(sGulkwa[11]);
               If (Val1 >= 1.5) And (Val1 <= 2.0) And
                   (Val3 <= 45) And
                  (Val4 <= 70) and (cSex = 1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End
              else If (Val1 >= 1.5) And (Val1 <= 2.0) And
                   (Val3 <= 34) And  
                  (Val4 <= 42) and (cSex = 2)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BIM ';
                  End;
            End;


//2.Cod = AFP ((TT01=양성 또는 T001 >= 16)인 남자)
      If (sGulkwa[088] <> '') Then
         Begin
            If ((sGulkwa[088] = '양성') or (sGulkwa[088] = '약양성')) And (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFP ';
               End;
       End;

      If (sGulkwa[88] <> '') Then
         Begin
            If ((sGulkwa[088] = '양성') or (sGulkwa[088] = '약양성')) and  (( Gpreg = 'N') or ( Gpreg = 'P')) And (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFP ';
               End;
       End;
     
//2.Cod = AFP ((TT01=양성 또는 TT02>15.0  --AFP0   --20120704
      
       If (sGulkwa[118] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[118]);
            If ((cSex = 1) And (Val1 >=9)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20131224') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFP0';
               End
           else  If ((Gpreg <> 'Y') And (cSex = 2))And (Val1 >=9) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20131224') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFP0';
               End;
       End;


//1.Cod = AFPF ((TT01=양성 또는 T001 >= 16)인 여자이면서 45세이하)
      If (sGulkwa[88] <> '') Then
         Begin
            If ((sGulkwa[88] = '양성') or (sGulkwa[88] = '약양성')) And
               (cSex = 2) And ( Gpreg = 'Y') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFPF';
               End;
         End;



// 임신시 , TT02 양성일경우   AFG   --20120704
    {   If (sGulkwa[118] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[118]);
            If (Val1 >= 15.01) And (cSex = 2) And ( Gpreg <> 'N') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFG ';
               End;
         End;    }
// 임신시 , TT02 양성일경우   AFG   --20120704


          If (sGulkwa[118] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[118]);
            If (Val1 >=9) And (cSex = 2) And ( Gpreg = 'Y') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20131224')    Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFG ';
               End;
         End;
//  TT03 양성일경우   AFP0 , TT03 음성인 경우   --20151214 수지

         If (sGulkwa[171] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[171]);
            If (Val1 >=8.10) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20151226')    Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFP0';
               End;
         End;
         If (sGulkwa[171] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[171]);
            If (Val1 <= 8.09) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20151226')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AFPD';
               End;
         End;



//21.Cod = LAP (C012 >= 100) And (C009,C010 <= 40) And (C011 <= 50)
      If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
         (sGulkwa[11] <> '') And (sGulkwa[12] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[12]);
            If (Val4 >= 100) And (Val1 <= 40.0) And
               (Val2 <= 40.0) And (Val3 <= 50)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LAP ';
               End
           {else If (Val4 >= 60) And (Val1 <= 40.0) And
                 (Val2 <= 45.0) And (Val3 <= 50)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'LAP ';
            End}
            else If (Val2 <= 45) and (Val3<=70) and (Val4>=70.1) and (cSex = 1) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'LAP ';
            End
            else If (Val2 <= 34) and (Val3<=42) and (Val4>=70.1) and (cSex = 2) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'LAP ';
            End
            else If ((Val2 >= 45.1) or (Val3>=70.1)) and (Val4>=100.1) and (cSex = 1) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'LAP ';
            End
            else If ((Val2 >= 34.1) or (Val3>=42.1)) and (Val4>=100.1) and (cSex = 2) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'LAP ';
            End;

         End;
//20.Cod = CHE (C014 >= 8000) And (C009,C010 <= 40) And (C011 <= 50)
      If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And
         (sGulkwa[11] <> '') And (sGulkwa[14] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);                                                                         
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[14]);
            If (Val4 >= 8000) And (Val1 <= 40.0) And
               (Val2 <= 40.0) And (Val3 <= 50) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CHE ';
               End;
         End;

{김소영 삭제의뢰 2014.04.24
//101.Cod = GOT (간기능정상이고 C009 >= 41)
      If tSokun[1] = '' Then
         If sGulkwa[09] <> '' Then
            Begin
               Val1 := StrToFloat(sGulkwa[09]);
               If (Val1 >= 41) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'GOT ';
                  End;
            End;
}
//순환기질환
 // LDL콜레스테롤(C027 >= 130)
       If (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString < '20140428') Then
       Begin
          Val1 := StrToFloat(sGulkwa[20]);
          If (Val1 >= 130) Then
          Begin
             K := K + 1;
             tSokun[K] := 'LDLH';
          End;
       End;

// 분주일 2014.04.28 부터
 // COD = LDLP (C027 >= 190)
       If (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')  Then
       Begin
          Val1 := StrToFloat(sGulkwa[20]);
          If (Val1 >= 190) Then
          Begin
             K := K + 1;
             tSokun[K] := 'LDLP';
          End;
       End;
 // COD = LDLO (160 <= C027 <= 189.9)
       If (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')   Then
       Begin
          Val1 := StrToFloat(sGulkwa[20]);
          If (Val1 >= 160) And (Val1 <= 189.9) Then
          Begin
             K := K + 1;
             tSokun[K] := 'LDLO';
          End;
       End;
 // Cod = LDLK(130 <= C027 <= 159.9)
       If (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')  Then
       Begin
          Val1 := StrToFloat(sGulkwa[20]);
          If (Val1 >= 130) And (Val1 <= 159.9) Then
          Begin
             K := K + 1;
             tSokun[K] := 'LDLK';
          End;
       End;
// Cod = LDLN (C025 >= 200 And 100 <= C027 <= 129.9)
      If (sGulkwa[18] <> '') And (sGulkwa[20] <> '') And (sGulkwa[24] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')   Then
      Begin
         Val1 := StrToFloat(sGulkwa[18]);
         Val2 := StrToFloat(sGulkwa[20]);
         Val3 := StrToFloat(sGulkwa[24]);
         if (sGulkwa[37]) <> '' then  Val4 := StrToFloat(sGulkwa[37])
         else Val4 := 0;
         Chk_Dang_Drug := '';         //통합문진 당뇨 확인
         Chk_Dang_Jindan := '';
         With QryTot_Munjin2018 Do
         Begin
            Close;
            ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
            ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
            ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
            Open;
            If RecordCount > 0 Then
            Begin
               Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
               Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
            End;
            Close;
         End;
         If (Val1 >= 200) And (Val2 >= 100) And (Val2 <= 129.9) And (val3 <= 125.9) And ((val4 <= 6.4) or (val4 = 0))  And
            (Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1') Then
         Begin
            K := K + 1;
            tSokun[K] := 'LDLN';
         End;
      End;

// Cod = LDLQ (C025 <= 199.9 And 100 <= C027 <= 129.9)
      If (sGulkwa[18] <> '') And (sGulkwa[20] <> '') And (sGulkwa[24] <> '')  And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')   Then
      Begin
         Val1 := StrToFloat(sGulkwa[18]);
         Val2 := StrToFloat(sGulkwa[20]);
         Val3 := StrToFloat(sGulkwa[24]);
         IF (sGulkwa[37] <> '') then  Val4 := StrToFloat(sGulkwa[37])
         else Val4 := 0;
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                  Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;
         If (Val1 <= 199.9) And (Val2 >= 100) And (Val2 <= 129.9) and (val3 <= 125.9) and ((val4 <= 6.4) or (val4 = 0))  and
            (Chk_Dang_Drug <> '1') and (Chk_Dang_Jindan <> '1')  Then
         Begin
            K := K + 1;
            tSokun[K] := 'LDLQ';
         End;
      End;

// Cod = LDLL (C025 >= 200 And 0 <= C027 <= 99.9)
      If (sGulkwa[18] <> '') And (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427')   Then
      Begin
         Val1 := StrToFloat(sGulkwa[18]);
         Val2 := StrToFloat(sGulkwa[20]);
         If (Val1 >= 200) And (Val2 <= 99.9) Then
         Begin
            K := K + 1;
            tSokun[K] := 'LDLL';
         End;
      End;

 //30.Cod = CHL (C025 <= 119)
       If (sGulkwa[18] <> '') Then
       Begin
          Val1 := StrToFloat(sGulkwa[18]);
          If Val1 <= 119 Then
          Begin
             K := K + 1;
             tSokun[K] := 'CHL ';
          End;
       End;
// 중성지방 Cod = TH3(150 <= C028 <= 199), Cod = TH(200 <= C028 <= 399), Cod = TH5(400 <= C028)
      If (sGulkwa[21] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString < '20140428') Then
      Begin
         Val1 := StrToFloat(sGulkwa[21]);
         If (Val1 >= 150) AND (Val1 <= 199) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH3 ';
         End
         else if (Val1 >= 200) AND (Val1 <= 399) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH  ';
         End
         else if (Val1 >= 400) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH5 ';
         End;
      End;

// 중성지방 Cod = TH3(150 <= C028 <= 199), Cod = TH(200 <= C028 <= 499), Cod = TH5(500 <= C028 <= 999.9), Cod = TH5(1,000 <= C028)   <-2014.04.23 이후
      If (sGulkwa[21] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427') Then
      Begin
         Val1 := StrToFloat(sGulkwa[21]);
         If (Val1 >= 150) AND (Val1 <= 199) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH3 ';
         End
         else if (Val1 >= 200) AND (Val1 <= 499) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH  ';
         End
         else if (Val1 >= 500) AND (Val1 <= 999) Then
         Begin
            K := K + 1;
            tSokun[K] := 'TH5 ';
         End
         else if (Val1 >= 1000) Then
         Begin
            K := K + 1;
            tSokun[K] := 'THVH';
         End;
      End;

{ 김소영 수정요청 2014.04.28
// Cod = LDLN (C025 >= 201 And 0 <= C027 <= 129)
      If (sGulkwa[18] <> '') And (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString < '20140428') Then
      Begin
         Val1 := StrToFloat(sGulkwa[18]);
         Val2 := StrToFloat(sGulkwa[20]);
         If (Val1 >= 200.1) And (Val2 >= 0) And (Val2 <= 129.9) Then
         Begin
            K := K + 1;
            tSokun[K] := 'LDLN';
         End;
      End;
// Cod = LDLN (C025 >= 199.1 And 0 <= C027 <= 129)
      If (sGulkwa[18] <> '') And (sGulkwa[20] <> '') And (qryBunju.FieldByName('Dat_bunju').AsString > '20140427') Then
      Begin
         Val1 := StrToFloat(sGulkwa[18]);
         Val2 := StrToFloat(sGulkwa[20]);
         If (Val1 >= 199.1) And (Val2 >= 0) And (Val2 <= 129.9) Then
         Begin
            K := K + 1;
            tSokun[K] := 'LDLN';
         End;
      End;
}

// Cod = HCRP, S080>3.1
    {  If (sGulkwa[119] <> '')  Then
      Begin
         Val1 := StrToFloat(sGulkwa[119]);
         If (Val1 >= 3.1) Then
         Begin
            K := K + 1;
            tSokun[K] := 'HCRP';
         End;
      End;
           }
// Cod = PHL, C065<150,C065>232

      If (sGulkwa[120] <> '')  Then
      Begin
         Val1 := StrToFloat(sGulkwa[120]);
         If (Val1 > 232) or (Val1 < 150) Then
         Begin
            K := K + 1;
            tSokun[K] := 'PHL ';
         End;
      End;

//29.Cod = CTL (C025 >= 201 And C028 >= 150) And ((C030 > 521 남자) or (C030 > 651 여자)) AND C027없을경우
      If (sGulkwa[18] <> '') And (sGulkwa[21] <> '') And (sGulkwa[23] <> '') AND (sGulkwa[20] = '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[18]);
            Val2 := StrToFloat(sGulkwa[21]);
            Val3 := StrToFloat(sGulkwa[23]);
            If (Val1 >= 200.1) And (Val2 >= 200) And
               (((Val3 >= 521) And (cSex = 1)) Or ((Val3 >= 651) And (cSex = 2))) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CTL ';
               End;
         End;
//28.Cod = CTH (C025 >= 201 And C028 >= 150 AND C027없을경우)
  {    If tSokun[K] <> 'CTL ' Then
         If (sGulkwa[18] <> '') And (sGulkwa[21] <> '') AND (sGulkwa[20] = '') Then
             Begin
                Val1 := StrToFloat(sGulkwa[18]);
                Val2 := StrToFloat(sGulkwa[21]);
                If (Val1 >= 201) And (Val2 >= 200) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'CTH ';
                   End
                else If (Val1 >= 200.1) And (Val2 >= 150.1) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'CTH ';
                   End;
             End;    }
//99.Cod = CLH (C025 >= 201) And ((C030 > 521 남자) or (C030 > 651 여자)) AND C027없을경우
   {   If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'CTH ') Then
         If (sGulkwa[18] <> '') And (sGulkwa[23] <> '') AND (sGulkwa[20] = '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[18]);
               Val3 := StrToFloat(sGulkwa[23]);
               If (Val1 >= 201) And
                  (((Val3 >= 521) And (cSex = 1)) Or ((Val3 >= 651) And (cSex = 2))) Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'CLH ';
                   End;
            End;
            }
//53.Cod = TLH (C028 >= 150 And (S030 >= 521 남자 Or  C030 >= 651 여자)
     { If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'CTH ') And
         (tSokun[K] <> 'CLH ') Then
         If (sGulkwa[21] <> '') And (sGulkwa[23] <> '') Then
             Begin
                Val1 := StrToFloat(sGulkwa[21]);
                Val2 := StrToFloat(sGulkwa[23]);
                If (Val1 >= 200) Then
                   Begin
                      If (cSex = 1) And (Val2 >= 521) Then
                          Begin
                             K := K + 1;
                             tSokun[K] := 'TLH ';
                          End;
                      If (cSex = 2) And (Val2 >= 651) Then
                          Begin
                             K := K + 1;
                             tSokun[K] := 'TLH ';
                          End;
                   End;
             End;  }
//25.Cod = CH1 (C025 >= 241) AND C027없을경우
      If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'CLH ') And
         (tSokun[K] <> 'TLH ') Then
          If (sGulkwa[18] <> '') AND (sGulkwa[20] = '') Then
              Begin
                 Val1 := StrToFloat(sGulkwa[18]);
                 If Val1 >= 240 Then
                    Begin
                       K := K + 1;
                       tSokun[K] := 'CH1 ';
                    End;
               End;

//24.Cod = CH (201 >= C025 <= 240) AND C027없을경우
      If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'TLH ') And
         (tSokun[K] <> 'CLH ') And
         (tSokun[K] <> 'CH1 ') Then
         If (sGulkwa[18] <> '') AND (sGulkwa[20] = '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[18]);
               If (Val1 >= 200.1) And (Val1 <= 239.9) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'CH  ';
                  End;
               End;
      TG_Check := '';
//67.Cod = TH (C028 >= 150)
      If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'CTH ') And
         (tSokun[K] <> 'CLH ') And
         (tSokun[K] <> 'TLH ') And
         (tSokun[K] <> 'CH1 ') And
         (tSokun[K] <> 'CH  ') Then
          If (sGulkwa[21] <> '') Then
             Begin
                 Val1 := StrToFloat(sGulkwa[21]);
                 {If (Val1 >= 200) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TH  ';
                     End;}
                 If (Val1 >= 150) And (Val1 <= 199) Then
                     Begin
                        TG_Check := 'KKKK';
                     End;
             End;
//23.Cod = LH (C030 >= 521 And 남자) or (C030 > 651 여자)
  {    If (tSokun[K] <> 'CTL ') And
         (tSokun[K] <> 'CTH ') And
         (tSokun[K] <> 'CLH ') And
         (tSokun[K] <> 'TLH ') And
         (tSokun[K] <> 'CH1 ') And
         (tSokun[K] <> 'CH  ') And
         (tSokun[K] <> 'TH  ') Then
          If (sGulkwa[23] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[23]);
                  If ((Val1 >= 521) And (cSex = 1)) Or
                     ((Val1 >= 651) And (cSex = 2)) Then
                       Begin
                          K := K + 1;
                          tSokun[K] := 'LH  ';
                       End;
              End;  }
//26.Cod = CK (281 >= C019 <= 320 남자) Or (191 >= C019 <= 280 여자)
    {  If (sGulkwa[16] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[16]);
            If ((Val1 >= 281) And (Val1 <= 320) And (cSex = 1)) Or
               ((Val1 >= 191) And (Val1 <= 280) And (cSex = 2)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CK  ';
               End;
         End;   }
//27.Cod = CKH (C019 >= 321 남자) Or (C019 >= 281 여자)
      If (sGulkwa[16] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[16]);
            If ((Val1 >= 321) And (cSex = 1)) Or
               ((Val1 >= 281) And (cSex = 2)) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CKH ';
               End
            else If (Val1 >= 200.1) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CKH ';
               End;
         End;
// Cod = HDLH (C026 >= 100)
      If (sGulkwa[19] <> '') Then
      Begin
         Val1 := StrToFloat(sGulkwa[19]);
         if (sGulkwa[20]) <> '' then  Val2 := StrToFloat(sGulkwa[20])
         else  Val2 :=0;

         If (Val1 >= 100) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
         Begin
            K := K + 1;
            tSokun[K] := 'HDLH';
         End
         else If (Val1 >= 60.1) and ((Val2=0) or (Val2<=129.9)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
         Begin
            K := K + 1;
            tSokun[K] := 'HDLH';
         End;
      End;



// HDL콜레스테롤(남자 : C026 <= 39, 여자 : C026 <= 49)
      If  (tSokun[K] <> 'CHL ')  then
      begin
         If (sGulkwa[19] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[19]);
            if (sGulkwa[18]) <> '' then  Val2 := StrToFloat(sGulkwa[18])
            else  Val2 :=0;
            If (((Val1 <= 39.9) AND (cSex = 1)) OR ((Val1 <= 49.9) AND (cSex = 2)))
                and (Val2>=120) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HDLL';
            End
            else If (Val1 <= 39.9) and (Val2>=120)  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HDLL';
            End;
         End;
      end;
//Cod = CRL (((남자C026 <= 39)OR(여자C026 <= 49)) And C029 <= 29.9) c025 <=119  --- 20120704
     If (sGulkwa[19] <> '') And (sGulkwa[22] <> '') and (sGulkwa[18] <> '') Then
      Begin
         Val1 := StrToFloat(sGulkwa[19]);
         Val2 := StrToFloat(sGulkwa[22]);
         Val3 := StrToFloat(sGulkwa[18]);
         If (((Val1 <= 39.0) AND (cSex = 1)) OR ((Val1 <= 49.0) AND (cSex = 2))) And (Val2 <= 29.9) And (Val3 <= 119)Then
         Begin
            K := K + 1;
            tSokun[K] := ' ';
         End;
      End;

// 당뇨질환
//[20080709유동구]혈당 120 -> 126으로 수정..
//                     119 -> 125으로 수정..
// Cod = DIC4 (C032 >= 200 And C034 >= 2.86)

// Cod = DIH4 (C032 >= 200 And C083 >= 6.5)
      If tsokun[K] <> 'DIC4' Then
      begin
         If (sGulkwa[24] <> '') And (sGulkwa[37] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[37]);
            If (Val1 >= 200) And (Val2 >= 6.5) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
            Begin
               K := K + 1;
               tSokun[K] := 'DIH4';
            End
            else If (Val1 >= 200.1) And (Val2 >= 6.5) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
            Begin
               K := K + 1;
               tSokun[K] := 'DIH4';
            End;
         End;
      end;
// Cod = DIB4 (C032 >= 200 And C034 >= 2.86)
      If (tsokun[K] <> 'DIH4') AND (tsokun[K] <> 'DIC4') Then
      begin
         If (sGulkwa[24] <> '')  Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            if (sGulkwa[26]) <> '' then  Val2 := StrToFloat(sGulkwa[26])
            else val2 :=0;
            if (sGulkwa[37]) <> '' then  Val3 := StrToFloat(sGulkwa[37])
            else val3 :=0;
            If (Val1 >= 200) and ((Val2 <=285) or (Val2 =0)) and ((Val3 <=6.4) or (Val3 =0)) then
            Begin
               K := K + 1;
               tSokun[K] := 'DIB4';
            End;
         End;
      end;
//34.Cod = DID (126 <= C032 <= 199 And C034 >= 2.86 And U005 >= 70)
//34.Cod = DIC (126 <= C032 <= 199 And C034 >= 2.86) //DID를 DIC로 변경(한의 재단진단검사의학부서장1100009) - by.Song
    {  If (sGulkwa[24] <> '') And (sGulkwa[26] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[26]);
            If (Val1 >= 126) And (Val1 <= 199) And (Val2 >= 2.86) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'DIC ';
               End;
            if (sGulkwa[37] <> '') Then
            begin
               Val3 := StrToFloat(sGulkwa[37]);
               if (Val3 >= 6.5) Then
               begin
                  K := K + 1;
                  tSokun[K] := 'HBA1';
               end;
            end;
         End;
         }
// Cod = DIH (126 <= C032 <= 199 AND C083 > 6.5)
      If (tsokun[K] <> 'DIC ') Then
         If (sGulkwa[24] <> '') And (sGulkwa[37] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[24]);
               Val2 := StrToFloat(sGulkwa[37]);
               If (Val1 >= 126) And (Val1 <= 199) And (Val2 >= 6.5) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'DIH ';
                  End
               else If (Val1 >= 126) And (Val1 <= 200) And (Val2 >= 6.5) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'DIH ';
                  End;
            End;
{//30.Cod = DIB (126 <= C032 <= 199 And 1 <= U005 <= 100)
      If (tsokun[K] <> 'DIH ') And
         (tSokun[K] <> 'DIC ') Then
         If (sGulkwa[24] <> '') And (sGulkwa[93] <> '') Then
             Begin
                 Val1 := StrToFloat(sGulkwa[24]);
                 Val2 := StrToFloat(sGulkwa[93]);
                 If (Val1 >= 126) And (Val1 <= 199) And (Val2 >= 1) AND (Val2 >= 100)Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'DIB ';
                     End;
             End; }   // 김소영 선새님 요청 U005 제외

//30.Cod = DIB (126 <= C032 <= 199
      If (tsokun[K] <> 'DIH ') And
         (tSokun[K] <> 'DIC ') Then
         If (sGulkwa[24] <> '') Then
             Begin
                 Val1 := StrToFloat(sGulkwa[24]);
                 if sGulkwa[37] <> '' then Val2 := StrToFloat(sGulkwa[37])
                 else val2 :=0;
                 if sGulkwa[26] <> '' then Val3 := StrToFloat(sGulkwa[26])
                 else val3 :=0;
                 If (Val1 >= 126) And (Val1 <= 199) and  ((Val2 >=0) and (Val2<=6.4)) and ((Val3 >=0) and (Val3<=2.85)) And
                    ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'DIB ';
                     End
                 else If (Val1 >= 126) And (Val1 < 200) and  ((Val2 >=0) and (Val2<=6.4)) and ((Val3 >=0) and (Val3<=285)) And
                         ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'DIB ';
                     End;
             End;
// Cod = DIC5 (100 <= C032 <= 125 AND C034 >= 2.86)
    {  If (sGulkwa[24] <> '') AND (sGulkwa[26] <> '') Then
      Begin
         Val1 := StrToFloat(sGulkwa[24]);
         Val2 := StrToFloat(sGulkwa[26]);
         if (Val1 >= 100) And (Val1 <= 125) And (Val2 >= 2.86) Then
         begin
            K := K + 1;
            tSokun[K] := 'DIC5';
            if (sGulkwa[37] <> '') then
            begin
               Val3 := StrToFloat(sGulkwa[37]);
               if (Val3 >= 6.5) then
               begin
                  K := K + 1;
                  tSokun[K] := 'HBA1';
               end;
            end;
         end;
      end;  }
// Cod = DIH5 (100 <= C032 <= 125 AND C0834 >= 6.5)
      If (tsokun[K] <> 'DIC5') then
      begin
         If (sGulkwa[24] <> '') AND (sGulkwa[37] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[37]);
            if (Val1 >= 100) And (Val1 <= 125) AND (Val2 >= 6.5) Then
            begin
               K := K + 1;
               tSokun[K] := 'DIH5';
            end;
         end;
      end;
// Cod = DMN \(100 <= C032 <= 125)
     If (sGulkwa[24] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            if sGulkwa[37] <> '' then Val2 := StrToFloat(sGulkwa[37])
            else val2 :=0;
            if sGulkwa[26] <> '' then Val3 := StrToFloat(sGulkwa[26])
            else val3 :=0;

            if (Val1 >= 100) And (Val1 <= 125) and
               ((Val2 >=0) and (Val2<=6.4)) and ((Val3 >=0) and (Val3<=285))  and
               ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')     Then
            begin
               K := K + 1;
               tSokun[K] := 'DMN ';
            end
            else if (Val1 >= 100) And (Val1 <= 125) and
               ((Val2 >=5.7) and (Val2<=6.4)) and ((Val3 >=0) and (Val3<=285)) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
            begin
               K := K + 1;
               tSokun[K] := 'DMNH';
            end
            else if (Val1 >= 100) And (Val1 <= 125) and
               ((Val2 >=0) and (Val2 < 5.7)) and 
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            begin
               K := K + 1;
               tSokun[K] := 'DMN ';
            end;
         end;

// Cod = DIN (C032 <= 99 And C083 >= 6.5)
      If (sGulkwa[24] <> '') And (sGulkwa[37] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[37]);

            //20140913~ 곽윤설
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                  Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End; //~20140913 곽윤설

            If (Val1 <= 99.9) And (Val2 >= 6.5)  then
            Begin
               K := K + 1;
               tSokun[K] := 'DIN ';
            End
            Else If (Val1 <= 99.9) And
                    ((Val2 >= 5.7) And (Val2 <= 6.4))    And
                    ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) then //20140911 곽윤설
            Begin
               K := K + 1;
               tSokun[K] := 'DINH';
            End;
         End;
    If (sGulkwa[20] <> '')  Then
         Begin
            if sGulkwa[24] <> '' then  Val1 := StrToFloat(sGulkwa[24])
            else Val1 :=0;
            if sGulkwa[37] <> '' then  Val2 := StrToFloat(sGulkwa[37])
            else Val2 :=0;
            Val3 := StrToFloat(sGulkwa[20]);
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                  Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End; //~20140913 곽윤설
               If ((Val3>= 100) and (Val3 <=129.9)) and ((Val1 >= 126 ) Or (Val2 >= 6.5) or (Chk_Dang_Drug = '1') or  (Chk_Dang_Jindan = '1')) then
            Begin
               K := K + 1;
               tSokun[K] := 'LDLF';
            END;
         END;
//UPDS
 if (sGulkwa[92] <> '') Then
    Begin
       Val1 := StrToFloat(sGulkwa[92]);
       if sGulkwa[24] <> '' then  Val2 := StrToFloat(sGulkwa[24])
            else Val2 :=0;
       if sGulkwa[37] <> '' then  Val3 := StrToFloat(sGulkwa[37])
            else Val3 :=0;

            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                  Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End; //~20140913 곽윤설
       IF  (Val1 >= 10) AND ((Val2 >= 126) or (Val3 >= 6.5)) and ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) then
       Begin
         K := K + 1;
         tSokun[K] := 'UPDS';
       End;
 End;



// Cod = SL (C032 <= 59)
      If (sGulkwa[24] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            If (Val1 <= 40) then
            Begin
               K := K + 1;
               tSokun[K] := 'SLR ';
            End
            else If (Val1 >= 40.1) and  (Val1 <= 50) then
            Begin
               K := K + 1;
               tSokun[K] := 'SL  ';
            End
            else If (Val1 >= 50.1) and  (Val1 <= 60) then
            Begin
               K := K + 1;
               tSokun[K] := 'SL2 ';
            End;
         End;

//신장질환
// 2003년04월15일 정정됨(BCH2>BCH3>U03>BCH1>BCH)
//BCH2 (C041 >= 50 And C042 >= 6.0 And U004 >= 40 And U009 >= 30)
    {  If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') And (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            Val3 := StrToFloat(sGulkwa[92]);
            Val4 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 50) And (Val2 >= 6.0) And (Val3 >= 40) And (Val4 >= 30) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BCH2';
               End;
         End;   }
//BCH3 (C041 < 50 And C041 >= 30 And C042 >= 3.0 And C042 < 6.0 And U004 >= 40 And U009 >= 30)
 {     If (tsokun[K] <> 'BCH2') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') And (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
            Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            Val3 := StrToFloat(sGulkwa[92]);
            Val4 := StrToFloat(sGulkwa[97]);
            If ((Val1 < 50)  And (Val1 >= 30)) And
               ((Val2 < 6.0) And (Val2 >= 3.0)) And
               (Val3 >= 40) And (Val4 >= 30) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BCH3';
               End;
            End; }
{//U03 (C041  >= 30 And C042 >= 1.6  And U004 >= 40)
      If (tsokun[K] <> 'BCH2') And
         (tsokun[K] <> 'BCH3') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') And (sGulkwa[92] <> '') Then
            Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            Val3 := StrToFloat(sGulkwa[92]);
            If (Val1 >= 30) And (Val2 >= 1.6) And (Val3 >= 40) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'U03 ';
               End;
            End;   }
//BCH1 (C041  >= 25 And C041 <= 29 And C042 >= 1.6 And C042 <= 1.9)
      {If (tsokun[K] <> 'BCH2') And
         (tsokun[K] <> 'BCH3') And
         (tsokun[K] <> 'U03 ') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '')  Then
            Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            If ((Val1 >= 25) And (Val1 <= 29)) And
               ((Val2 >= 1.6) And (Val2 <= 1.9)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BCH1';
               End;
            End; }
{//BCH (C041  >= 30And C042 >= 3.0)
      If (tsokun[K] <> 'BCH2') And
         (tsokun[K] <> 'BCH3') And
         (tsokun[K] <> 'U03 ') And
         (tsokun[K] <> 'BCH1') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '')  Then
            Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            If (Val1 >= 30) And (Val2 >= 3.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BCH ';
               End;
            End; }

{2003년04월15일 삭제됨
//57.Cod = U03 (C041 >= 30 And C042 >= 2.0 And U004 >= 40)
      If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') And (sGulkwa[92] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[29]);
            Val2 := StrToFloat(sGulkwa[30]);
            Val3 := StrToFloat(sGulkwa[92]);
            If (Val1 >= 30) And (Val2 >= 2.0) And (Val3 >= 40)Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'U03 ';
               End;
         End;
//10.Cod = BCH (25 >= C041 <= 30) And (1.5 >= C042 <= 2)
      If (tsokun[K] <> 'U03 ') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[29]);
               Val2 := StrToFloat(sGulkwa[30]);
               If (Val1 >= 15) And (Val1 <= 30) And
                  (Val2 >= 1.5) And (Val2 <= 2.0) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BCH ';
                  End;
            End;
//11.Cod = BCH1 (C041 > 30) And (C042 > 2.0)
      If (tsokun[K] <> 'U03 ') Then
         If (sGulkwa[29] <> '') And (sGulkwa[30] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[29]);
               Val2 := StrToFloat(sGulkwa[30]);
               If (Val1 > 30) And (Val2 > 2) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BCH1';
                  End;
            End;
//14.Cod = BUN (C041 >= 30)
      If (tsokun[K] <> 'U03 ') Then
         If (sGulkwa[29] <> '') Then
            Begin
              Val1 := StrToFloat(sGulkwa[29]);
              If Val1 >= 30 Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BUN ';
                  End;
            End;
}

//요산
//69.Cod = U (C017 >= 9.0) and C042
      If (sGulkwa[15] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[15]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 9) and ((Val2=0) or (Val2 <=1.40)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'U   ';
               End;
         End;
//54.Cod = U1 (C017 >= 8.1 And C017 < 9.0)
      If (sGulkwa[15] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[15]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 8.1) And (Val1 < 9.0) and  ((Val2=0) or (Val2 <=1.40)) Then
                Begin
                  K := K + 1;
                  tSokun[K] := 'U1  ';
                End;
         End;
//54.Cod = U4 (C017 >= 8.1 And C017 < 9.0)
      If (sGulkwa[15] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[15]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 8.1)  and (Val2 >=1.41) Then
                Begin
                  K := K + 1;
                  tSokun[K] := 'U4  ';
                End;
         End;


//췌장
//7.Cod = AMH (C039 >= 180)   C042 <1.40 or 무
      If (sGulkwa[28] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >=180.01) and ((Val2=0) or (Val2<=1.40)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMH ';
               End;
         End;
       If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') And ((qryBunju.FieldByName('Dat_bunju').AsString) < '20141216') Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >=125.01) and ((Val2=0) or (Val2<=1.40)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMH ';
               End;
         End;
         If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20141216') And ((qryBunju.FieldByName('Dat_bunju').AsString) < '20170302')  Then
           Begin
              Val1 := StrToFloat(sGulkwa[28]);
              if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
              else Val2 :=0;
              If (Val1 >=146.01) and ((Val2=0) or (Val2<=1.40)) Then
                 Begin
                    K := K + 1;
                    tSokun[K] := 'AMH ';
                 End;
           End;
         If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20170302')  Then
           Begin
              Val1 := StrToFloat(sGulkwa[28]);
              if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
              else Val2 :=0;
              If (Val1 >=109.1) and ((Val2=0) or (Val2<=1.40)) Then
                 Begin
                    K := K + 1;
                    tSokun[K] := 'AMH ';
                 End;
           End;
//Cod = AMHR (C039 >= 180) C42 >1.401
      If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 180.01) and (Val2>=1.41) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMHR';
               End;
         End;
      If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20141216') Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 125.01) and (Val2>=1.41) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMHR';
               End;
         End;
      If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20141216') and  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20170302') Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 146.01) and (Val2>=1.41) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMHR';
               End;
         End;
      If (sGulkwa[28] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20170302') Then
         Begin
            Val1 := StrToFloat(sGulkwa[28]);
            if sGulkwa[30] <> '' then  Val2 := StrToFloat(sGulkwa[30])
            else Val2 :=0;
            If (Val1 >= 109.01) and (Val2>=1.41) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AMHR';
               End;
         End;

//갑상선기능
{//49.Cod = THH (E002 > 15.0)
      If (sGulkwa[41] <> '') Then
         Begin
            Val2 := StrToFloat(sGulkwa[41]);
            If (Val2 > 15.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'THH ';
               End;
         End; }
//50.Cod = THH1 (12.5 > E002 <= 15.0)
  {    If tSokun[K] <> 'THH ' Then
         If (sGulkwa[41] <> '') Then
             Begin
                Val2 := StrToFloat(sGulkwa[41]);
                If  (Val2 > 12.5) And (Val2 <= 15.0) Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'THH1';
                   End;
            End;     }
//51.Cod = THL (E002 < 4.7)
      If (tSokun[K] <> 'THH ') And
         (tSokun[K] <> 'THH1') Then
          If (sGulkwa[41] <> '') Then
              Begin
                  Val2 := StrToFloat(sGulkwa[41]);
                  If (Val2 <= 4.49) Then
                      Begin
                         K := K + 1;
                         tSokun[K] := 'THL ';
                      End;
              End;
{//52.Cod = TSH (E003 >= 5.1)
      If (tSokun[K] <> 'THH ') And
         (tSokun[K] <> 'THH1') And
         (tSokun[K] <> 'THL ') Then
          If (sGulkwa[42] <> '') Then
              Begin
                 Val1 := StrToFloat(sGulkwa[42]);
                 If (Val1 >= 5.1) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TSH ';
                     End;
              End;  }
//53.Cod = TSH (E003 <= 0.2)
    {  If (tSokun[K] <> 'THH ') And
         (tSokun[K] <> 'THH1') And
         (tSokun[K] <> 'THL ') And
         (tSokun[K] <> 'TSH ') Then
          If (sGulkwa[42] <> '') Then
              Begin
                 Val1 := StrToFloat(sGulkwa[42]);
                 If (Val1 <= 0.2) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TSH2';
                     End;
              End; }

//관절염
//44.Cod = RAH (S003 이 양성)
      If (sGulkwa[69] <> '') Then
         Begin
            If (sGulkwa[69] = '양성') or
               (sGulkwa[69] = '약양성')Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RAH ';
               End;
         End;

     { If (sGulkwa[136] <> '') Then
         Begin
            val1 := StrToFloat(sGulkwa[136]);
            If (val1>=1.2) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RAH ';
               End;
         End;        }
     {  If (sGulkwa[137] <> '') Then
         Begin
            val1 := StrToFloat(sGulkwa[137]);
            If (val1>=1.2) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP1';
               End;
         End; }

         If (sGulkwa[137] <> '') Then
         Begin
            val1 := StrToFloat(sGulkwa[137]);
            If (val1<=1.89)and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MGL ';
               End
            else If (val1<=1.59)and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MGL ';
               End
            else If (val1>=3.11) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MGH ';
               End
            else If (val1>=2.61) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MGH ';
               End;
         End;


         If (sGulkwa[59] <> '') Then
         Begin
            val1 := StrToFloat(sGulkwa[59]);
            If (val1>=3) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BAS ';
               End;

         End;

//매독
//66.Cod = VD1 (S034 = '양성')
      If (sGulkwa[79] <> '') Then
         Begin
            If (sGulkwa[79] = '양성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VD1 ';
               End;
            If (sGulkwa[79] = '약양성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VD2 ';
               End;
         End;
//혈액질환
//74.Cod = ANE2 (H002 <= 8.0 And H008 <= 70 And H011 <= 2.0)
      If (sGulkwa[44] <> '') And (sGulkwa[50] <> '') And (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[44]);
            Val2 := StrToFloat(sGulkwa[50]);
            Val3 := StrToFloat(sGulkwa[53]);
            If (Val1 <= 8) And (Val2 <= 70) And (Val3 <= 2.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'ANE2';
               End;
         End;
//38.Cod = IDA (C045 <= 59 And H002 <= 12.0 남자) or (C045 <= 54 And H002 <= 11.7 여자) 51세이상 여자는 IDPM으로..
    {  If (tSokun[K] <> 'ANE2') Then
         If (sGulkwa[33] <> '') And (sGulkwa[44] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[33]);
            Val2 := StrToFloat(sGulkwa[44]);
            If (Val1 <= 59) And (Val2 <= 13.4) And (cSex = 1) And  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') then
            begin
               K := K + 1;
               tSokun[K] := 'IDA6';
            end
            else If (Val1 <= 59) And (Val2 <= 12.9) And (cSex = 1) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004')  then
            begin
               K := K + 1;
               tSokun[K] := 'IDA6';
            end;
            

         End;   }
//36.Cod = FEL2 (C049 <= 8.9 남자, C049 <= 5.9 여자)
     { If (tSokun[K] <> 'ANE2') AND
         (tSokun[K] <> 'IDA ') AND
         (tSokun[K] <> 'IDA6') AND
         (tSokun[K] <> 'IDPM') Then
         If (sGulkwa[36] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[36]);
               If ((Val1 <= 8.9) And (cSex = 1)) Or
                  ((Val1 <= 5.9) And (cSex = 2)) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'FEL2';
                  End;
            End;}
//38-1.Cod = IDO2(C045 <= 59 And H002 >= 13.6 남자)
//     Cod = IDO (C045 <= 54 And H002 >= 12.1 여자) 51세이상 여자는 IDOM
     { If (tSokun[K] <> 'ANE2') AND
         (tSokun[K] <> 'IDA ') And
         (tSokun[K] <> 'IDA6') AND
         (tSokun[K] <> 'IDPM') And
         (tSokun[K] <> 'FEL2') Then
         If (sGulkwa[33] <> '') And (sGulkwa[44] <> '') Then
             Begin
                Val1 := StrToFloat(sGulkwa[33]);
                Val2 := StrToFloat(sGulkwa[44]);
               if ((Val1 <= 54) And (Val2 >= 12.1) And (cSex = 2) AND (sAge >= 51)) Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'IDOM';
                   End
             End; }
//8.Cod = ANE (H002 <= 11.8 이고 여자 또는 H002 <= 12.5인 남자)
//

          If (sGulkwa[44] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
          Begin
               if (sGulkwa[36]) <> '' then Val1 := StrToFloat(sGulkwa[36])//  C049
               else val1 :=0;
               if (sGulkwa[33]) <> '' then Val2 := StrToFloat(sGulkwa[33]) //  C045
               else Val2 :=0;
               Val3 := StrToFloat(sGulkwa[44]);     //H002                      
                if cSex = 2 then
                begin
                     if ((Val1=0) and (Val2=0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1=0) and (Val2>=55.0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=10) and (Val2>=55)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';                                                                                                         
                     End
                     else if ((Val1=0) and (Val2=0)) and  ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1=0) and (Val2>=55.0)) and ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and   ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=10) and (Val2>=55)) and ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else if ((Val1=0) and (Val2=0)) and  (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC ';
                     End
                     else  if ((Val1=0) and (Val2>=55.0)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and   (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=10) and (Val2>=55)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End;
                end
                else if cSex = 1 then
                begin
                     if ((Val1=0) and (Val2=0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End

                     else  if ((Val1>=17) and (Val2>=60)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End

                     else if ((Val1=0) and (Val2=0)) and  ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2>=60)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End

                     else if ((Val1=0) and (Val2=0)) and  (Val3<=7.99)  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and  (Val3<=7.99)  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and   (Val3<=7.99)  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2>=60)) and   (Val3<=7.99)  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End;


                 end;

             End;
          If (sGulkwa[44] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
          Begin
               if (sGulkwa[36]) <> '' then Val1 := StrToFloat(sGulkwa[36])   //c049
               else val1 :=0;
               if (sGulkwa[33]) <> '' then Val2 := StrToFloat(sGulkwa[33])  //c045
               else Val2 :=0;
               Val3 := StrToFloat(sGulkwa[44]);//h002
                if cSex = 2 then
                begin
                     if ((Val1=0) and (Val2=0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1=0) and (Val2>=50.0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=10) and (Val2>=50)) and  ((Val3>=11.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else if ((Val1=0) and (Val2=0)) and  ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1=0) and (Val2>=50.0)) and ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and   ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=10) and (Val2>=50)) and ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else if ((Val1=0) and (Val2=0)) and  (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC ';
                     End
                     else  if ((Val1=0) and (Val2>=50.0)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=10) and (Val2=0)) and   (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=10) and (Val2>=50)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End;
                end
                else if cSex = 1 then
                begin
                     if ((Val1=0) and (Val2=0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End

                     else  if ((Val1>=17) and (Val2>=60)) and  ((Val3>=12.0) and (Val3<=12.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End

                     else if ((Val1=0) and (Val2=0)) and  ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else  if ((Val1>=17) and (Val2>=60)) and   ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End


                     else if ((Val1=0) and (Val2=0)) and  (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC ';
                     End
                     else  if ((Val1=0) and (Val2>=60.0)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=17) and (Val2=0)) and   (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End
                     else  if ((Val1>=17) and (Val2>=60)) and (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End;
                 end;

             End;



        {     If (sGulkwa[44] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
          Begin
               if (sGulkwa[36]) <> '' then Val1 := StrToFloat(sGulkwa[36])
               else val1 :=0;
               if (sGulkwa[33]) <> '' then Val2 := StrToFloat(sGulkwa[33])
               else Val2 :=0;
               Val3 := StrToFloat(sGulkwa[44]);
               if cSex = 1 then
                begin
                     if  (((Val2>=65) or (Val2=0))  or ((Val1>=17) or (Val1=0))) and ((Val3>=12.0) and (Val3<=12.99))  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else if (((Val2>=65) or (Val2=0))  or ((Val1>=17) or (Val1=0))) and  ((Val3>=8.0) and (Val3<=11.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else if (((Val2>=65) or (Val2=0))  or ((Val1>=17) or (Val1=0))) and  (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End;

                end
                else if cSex = 2 then
                begin
                      if  (((Val2>=50) or (Val2=0))  or ((Val1>=10) or (Val1=0))) and ((Val3>=11.0) and (Val3<=11.99))  then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'IN06';
                     End
                     else if  (((Val2>=50) or (Val2=0))  or ((Val1>=10) or (Val1=0))) and ((Val3>=8.0) and (Val3<=10.99)) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANE ';
                     End
                     else if (((Val2>=50) or (Val2=0))  or ((Val1>=10) or (Val1=0))) and  (Val3<=7.99) then
                         Begin
                         K := K + 1;
                         tSokun[K] := 'ANEC';
                     End;
                end;
             End;

                }




   ///H002 남자 >=18.1 여자 >=16.1  ->HHB
        If (sGulkwa[44] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[44]);
            If (Val1 >= 18.1) and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HHB ';
               End
            else If (Val1 >= 16.1) and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HHB ';
               End;
         End;
// FEL2 + ANE = IDA
      //[2006.06.01-소견10개로 늘림]
      For I := 1 To 10 Do //20131026
         If tSokun[I] = 'FEL2' Then
            For J := 1 To 5 Do
               If tSokun[J] = 'ANE ' Then
                  Begin
                     tSokun[J] := '';
                     tsokun[I] := 'IDA';
                     K := K - 1;
                     Break;
                  End;

//70.Cod = WBC (H011 >= 11.1)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            If (Val1 >= 11.1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC ';
               End;


         End;
//70.Cod = WBC1 (10.1 <=H011 >= 11.0)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            If (Val1 >= 10.1) AND (Val1 <= 11.0) and  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End
            else If (Val1 >= 10.21) AND (Val1 <= 11.0) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End
            else If (Val1 >= 9.81) AND (Val1 <= 11.0) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End;
         End;
{//71.Cod = WBE (H011 >= 11.0 And H016 >= 8)
      If (sGulkwa[53] <> '') And (sGulkwa[58] <> '')Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            Val2 := StrToFloat(sGulkwa[58]);
            If (Val1 >= 11) And (Val2 >= 8) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBE ';
               End;
         End; }

//MONH (H015 >= 20.1)
      If (sGulkwa[57] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[57]);
            If Val1 >= 20.1 Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MONH';
               End;
         End;
//35.Cod = EOS (H016 >= 12)
      If (sGulkwa[58] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[58]);
            If Val1 >= 12 Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'EOS ';
               End;
         End;
//45.Cod = RH (H024 이 음성)
      If (sGulkwa[66] <> '') Then
         Begin
            If (sGulkwa[66] = 'Rh-') Then
                Begin
                   K := K + 1;
                   tSokun[K] := 'RH ';
                End;
         End;

//소변관리
// Cod = USH1(U005 >= 70 & C032 >= 100)
// Cod = USH1(U005 >= 250 & C032 >= 100) --20110620 김희석
// Cod = USH1(U005 >= 100 & C032 >= 100) --20140428 김창규
      If (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') And
         (sGulkwa[93] <> '') And (sGulkwa[24] <> '') Then
      begin
         Val1 := StrToFloat(sGulkwa[93]);
         Val2 := StrToFloat(sGulkwa[24]);
         If (Val1 >= 250) And (Val2 >= 100) Then
         begin
            K := K + 1;
            tSokun[K] := 'USH1';
         end;
      end;
      If (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') And
         (sGulkwa[93] <> '') And (sGulkwa[24] <> '') Then
      begin
         Val1 := StrToFloat(sGulkwa[93]);
         Val2 := StrToFloat(sGulkwa[24]);
         If (Val1 >= 100) And (Val2 >= 100) Then
         begin
            K := K + 1;
            tSokun[K] := 'USH1';
         end;
      end;
//65.Cod = URWP (U001 >= 50 And U004 >= 40 And U009 >= 30)
//65.Cod = URWP (U001 >= 25 And U004 >= 30 And U009 >= 10)--20110620 김희석
     { If (sGulkwa[89] <> '') And (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            Val2 := StrToFloat(sGulkwa[92]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 25) And (Val2 >= 30) And (Val3 >= 10)Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
               End;
         End;      }
//59.Cod = URW (U001 >= 50 And U009 >= 30)
//59.Cod = URW (U001 >= 25 And U009 >= 10)--20110620 김희석
    {  If tsokun[K] <> 'URWP' Then
         If (sGulkwa[89] <> '') And (sGulkwa[97] <> '') Then
             Begin
                Val1 := StrToFloat(sGulkwa[89]);
                Val2 := StrToFloat(sGulkwa[97]);
                If (Val1 >= 25) And (Val2 >= 10) And
                   (sGulkwa[90] <> '양성')Then
                   Begin
                      K := K + 1;
                      //090409 철. 김소영선생님 요청.
                      if (sAge <= 50) and (cSex = 2) then
                         tSokun[K] := 'URW5'
                      else if (sAge >= 51) and (cSex = 2) then
                         tSokun[K] := 'URW6'
                      else
                         tSokun[K] := 'URW ';
                   End
             End;   }
{//58.Cod = UPB (U004 >= 40 And U009 >= 30)
//58.Cod = UPB (U004 >= 30 And U009 >= 10)  --20110620 김희석
      If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
         If (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[92]);
               Val2 := StrToFloat(sGulkwa[97]);
               If (Val1 >= 30) And (Val2 >= 10)Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'UPB ';
                   End;
            End;}
//60.Cod = UWN (U001 >= 50 And U002 = '양성')
//60.Cod = UWN (U001 >= 25 And U002 = '양성')     --20110620 김희석
      {If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
         If (sGulkwa[89] <> '') And (sGulkwa[90] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            If (Val1 >= 25) And (sGulkwa[90] = '양성')Then
            Begin
               K := K + 1;
               if (cSex = 2) And (sAge <= 55) then tSokun[K] := 'UWN1'
               else if   (sAge >= 55)         then tSokun[K] := 'UWN ';
            End
            else If (Val1 <= 10) And (sGulkwa[90] = '양성')Then
            Begin
               K := K + 1;
               tSokun[K] := 'NIH ';
            End;
         End; }
//56.Cod = UWB (U001 >= 50)
//56.Cod = UWB (U001 >= 25) --20110620 김희석
    {  If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'UWN1') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
         If Sgulkwa[89] <> '' Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            If (Val1 >= 25)  Then
            Begin
               K := K + 1;
               if cSex = 2 then tSokun[K] := 'UWBC'
               else             tSokun[K] := 'UWB ';
            End
         End;}
//62.Cod = URB (U009 >= 30고 여자이면서 49세이하))
//62.Cod = URB (U009 >= 10고 여자이면서 49세이하)) --20110620 김희석
    {  If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
//         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'UWBC') And
         (tsokun[K] <> 'UWB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
          If sGulkwa[97] <> '' Then
             Begin
                 Val1 := StrToFloat(sGulkwa[97]);
                 If (Val1 >= 10) And (cSex = 2) And (sAge <= 49)Then
                     Begin
                         K := K + 1;
                         tSokun[K] := 'URB ';
                     End
             End; }


//63.Cod = URM (U009 >= 30고 남자)
//63.Cod = URM (U009 >= 10고 남자) --20110620 김희석
   {   If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'UWN1') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'UWBC') And
         (tsokun[K] <> 'UWB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
          If (sGulkwa[97] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[97]);
                  If (Val1 >= 10) And (cSex = 1)Then
                     Begin
                         K := K + 1;
                         tSokun[K] := 'URM ';
                     End;
              End; }
//61.Cod = USH (C032 <= 110 And U005 >= 70)
//61.Cod = USH (C032 <= 100 And U005 >= 250) -20110620 김희석
//61.Cod = USH (C032 <= 100 And U005 >= 100) -20140428 김창규
      If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'UWN1') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'URM ') And
         (tsokun[K] <> 'UWBC') And
         (tsokun[K] <> 'UWB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
          If (sGulkwa[24] <> '') And (sGulkwa[93] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[24]);
                  Val2 := StrToFloat(sGulkwa[93]);
                  If (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') And
                     (Val1 <= 99.9) And (Val2 >= 250) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'USH ';
                      End;
                  If (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') And
                     (Val1 <= 99.9) And (Val2 >= 100) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'USH ';
                      End;
              End;
    {    If (tsokun[K] = 'UWB ') or
           (tsokun[K] = 'UWBC') Then
          If (sGulkwa[92] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[92]);
                  If (Val1 >= 30) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'UPH ';
                      End;
              End;
                     }


//64.Cod = UPH (U004 >= 40)
//64.Cod = UPH (U004 >= 30)  -20110620 김희석
    {  If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
//         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'URM ') And
         (tsokun[K] <> 'USH ') And
         (tsokun[K] <> 'UWBC') And
         (tsokun[K] <> 'UWB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
          If (sGulkwa[92] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[92]);
                  If (Val1 >= 30) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'UPH ';
                      End;
              End;
         If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
//         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'URM ') And
         (tsokun[K] <> 'USH ') And
         (tsokun[K] <> 'UWBC') And
         (tsokun[K] <> 'UWB ') And
         (tsokun[K] <> 'URW5') And
         (tsokun[K] <> 'URW6') Then
          If (sGulkwa[92] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[92]);
                  If (Val1 >= 30) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'UPH ';
                      End;
              End; }


//77.Cod = URB2 (U009 >= 30고  여자 이면서 50세이상 여자)
//77.Cod = URB2 (U009 >= 10고  여자 이면서 50세이상 여자) -- 김희석20110620
  {    dCheck := 0;
      For dCnt1 := 1 To 140 Do
          If (tSokun[dCnt1] = 'URW ') Or
             (tSokun[dCnt1] = 'UPB ') Or
             (tSokun[dCnt1] = 'URWP') Or
             (tSokun[dCnt1] = 'URW6') Then
             dCheck := dCheck + 1;
      If dCheck = 0 Then
         If sGulkwa[97] <> '' Then
            Begin
               Val1 := StrToFloat(sGulkwa[97]);
               //if tSokun[K] = 'URW6' then
               If (Val1 >= 10) And (cSex = 2) And (sAge >= 50) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'URB2';
                  End;
            End;  }
//115.Cod = NCT (nicotine이 양성일 경우)
      If (sGulkwa[39] <> '') Then
      Begin
         If (sGulkwa[39] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'NCT ';
            End;
      End;

//그외의 소견처리
  If  (sGulkwa[127] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   Then
         Begin               
            Val2:=0;
            Val3:=0;
            if (Trim(sGulkwa[127]) <> '') then
            begin
            Val2 := StrToFloat(sGulkwa[127]);
            end;
            if (sGulkwa[128] <> '') then
            begin
                Val3 := StrToFloat(sGulkwa[128]);
            end;

           If  (Val2 >= 98.1) and ((Val3>=0.01) and (Val3 <=109.9)) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB2';
               End
            else If  (Val2 <= 67.9)  and (Val3 >=158.1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB3';
               End;
          {  else If (Val2 >= 98.1) and  (((Val3 <=158) and (Val3 >=110)) or (sGulkwa[128] = ''))  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
               End;  }
        End;
    If  (sGulkwa[127] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
         Begin               
            Val2:=0;
            Val3:=0;
            if (Trim(sGulkwa[127]) <> '') then
            begin
            Val2 := StrToFloat(sGulkwa[127]);
            end;
            if (sGulkwa[128] <> '') then
            begin
                Val3 := StrToFloat(sGulkwa[128]);
            end;

           If  (Val2 >160) and ((Val3>=0.01) and (Val3 <110)) and (cSex = 1) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB2';
               End
           else If  (Val2 >150) and ((Val3>=0.01) and (Val3 <120)) and (cSex = 2) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB2';
               End
           else If  (Val2 < 70)  and (Val3 >202) and (cSex = 1)   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB3';
               End
           else If  (Val2 < 60)  and (Val3 >225) and (cSex = 2)   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB3';
               End;
           {else If (Val2 > 133) and  (((Val3 <=202) and (Val3 >=104)) or (sGulkwa[128] = '')) and (cSex = 1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
               End
           else If (Val2 > 117) and  (((Val3 <=225) and (Val3 >=108)) or (sGulkwa[128] = '')) and (cSex = 2)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
               End; }
        End;
    if (sGulkwa[128] <> '') and (sGulkwa[127] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
       begin
       Val1 := StrToFloat(sGulkwa[128]);
       Val2 := StrToFloat(sGulkwa[127]);

       if (Val1>=158.01) and (Val2>=98.1) then
       begin
           if (val1/val2) >=1 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB4';
           End;
           {else if (val1/val2) <=9.9 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
           End; }
       end;

    end;
      if (sGulkwa[128] <> '') and (sGulkwa[127] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
       begin
       Val1 := StrToFloat(sGulkwa[128]);
       Val2 := StrToFloat(sGulkwa[127]);

       if (Val1>200) and (Val2>160)and  (cSex = 1) then
       begin
           if (val1/val2) >=1 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB4';
           End;
           {else if (val1/val2) <=9.9 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
           End; }
       end
       else if (Val1>220) and (Val2>150)and  (cSex = 2) then
       begin
           if (val1/val2) >=1 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB4';
           End;
           {else if (val1/val2) <=9.9 then
           Begin
                  K := K + 1;
                  tSokun[K] := 'APB ';
           End;}
       end;
    end;




//4.Cod = AHP (C090 = 양성)
      If (sGulkwa[76] <> '') Then
         Begin
            If (sGulkwa[76] = '양성')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AHP ';
               End;
         End;
     If (sGulkwa[76] <> '') Then
         Begin
            If (sGulkwa[76] = '약양성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'AHPW';
               End;
         End;

//15.Cod = CAD (T007 >= 35.1)
//      If (sGulkwa[84] <> '') Then
      If ((sGulkwa[84] <> '') or (sGulkwa[175] <> '')) And (cSex = 2) Then    //2014.03.29 수정(Only 여자)
         Begin
            if (sGulkwa[84] <> '') then
            Val1 := StrToFloat(sGulkwa[84])
            else Val1 := 0;
            if (sGulkwa[175] <> '') then
            Val2 := StrToFloat(sGulkwa[175])
            else val2 := 0;

            If (Val1 >= 35.01) or (val2 >= 35.01)Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CAD ';
               End;
         End;

//16.Cod = CA91 (37.1 > T006 <= 45.0)
      If (sGulkwa[83] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[83]);
            If (Val1 >= 37.1) And (Val1 <= 45.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CA91';
               End;
         End;
//17.Cod = CA92 (T006 >= 45.1)
      If (sGulkwa[83] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[83]);
            If Val1 >= 45.1 Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CA92';
               End;
         End;
//18.Cod = CEA (T002 >= 10.1)
      If (sGulkwa[82] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[82]);
            If Val1 >= 10.1 Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CEA ';
               End;
         End;
//19.Cod = CEA1 ( 5.1 >= T002 >= 15.0)
      If (sGulkwa[82] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[82]);
            If (Val1 >= 5.1) And (Val1 <= 10.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CEA1';
               End;
         End;
//37.Cod = FRT (C049 >= 447 남자, C049 >= 304 여자)
      If (sGulkwa[36] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            If ((Val1 >= 400.1) and (Val1 <= 499.9))  And (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FRT0 ';
               End
           else If (Val1 >= 500)   And (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FRT ';
               End
           else If ((Val1 >= 90.1) and (Val1 <= 199.9))  And (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FRT0';
               End
           else If (Val1 >= 200)    And (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FRT ';
               End;
         End;
//39.Cod = MCA1 (T008 >= 29.0)
      If (sGulkwa[85] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[85]);
            If Val1 >= 29 Then
               Begin
                  K := K + 1;
                 tSokun[K] := 'MCA1';
               End;
         End;
//40.Cod = PLH (H008 >= 520)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            If (Val1 >= 520) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End
           else If (Val1 >= 400.1) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End
           else If (Val1 >= 420.1) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End;
         End;
//43.Cod = PSA (T011 >= 4.1 Or T010 >= 4.1)  => (T011 >= 4.01 // 2015.04.30 곽윤설)
//      If (sGulkwa[87] <> '') Or (sGulkwa[105] <> '') Then
      If ((sGulkwa[87] <> '') Or (sGulkwa[105] <> '') OR (sGulkwa[174] <> '')) And (cSex = 1) Then    //2014.03.29 수정 (Only 남자)
         Begin
            If (sGulkwa[87] <> '') Then
               Val1 := StrToFloat(sGulkwa[87])
            Else
               Val1 := 0;
            If (sGulkwa[105] <> '') Then
               Val2 := StrToFloat(sGulkwa[105])
            Else
               Val2 := 0;
            IF (sGulkwa[174] <> '') Then
               Val3 := StrToFloat(sGulkwa[174])
            Else
               Val3 := 0;
            If (Val1 >= 4.01) Or (Val2 >= 4.1) OR (Val3 >= 3.61) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PSA ';
               End;
         End;
      if (sGulkwa[74] <> '') and (sGulkwa[73] <> '') Then
      begin
         if (sGulkwa[73] = '음성') and (sGulkwa[74] = '음성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG3';
         end
         else if (sGulkwa[73] = '음성') and (sGulkwa[74] = '양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG2';
         end
         else if (sGulkwa[73] = '음성') and (sGulkwa[74] = '약양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG1';
         end
         else if ((sGulkwa[73] = '양성') or (sGulkwa[73] = '약양성')) and ((sGulkwa[74] = '양성') or (sGulkwa[74] = '약양성')) Then
         begin
            K := K + 1;
            tSokun[K] := 'RIGM';
         end
         else if (sGulkwa[73] = '양성') and (sGulkwa[74] = '음성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIM ';
         end
         else if (sGulkwa[73] = '약양성') and (sGulkwa[74] = '음성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIMW';
         end;
      end;
      if (sGulkwa[74] <> '') and (sGulkwa[73] = '') Then
      begin
         if (sGulkwa[74] = '약양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIGW';
         end
         else if (sGulkwa[74] = '음성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIGN';
         end
         else if (sGulkwa[74] = '양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG ';
         end;

      end;
      if (sGulkwa[74] = '') and (sGulkwa[73] <> '') Then
      begin
         if (sGulkwa[73] = '양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIM1';
         end
         else if (sGulkwa[73] = '약양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIM2';
         end;
      end;

//??.Cod = RIG1,RIG2,RIG3 [2008.09.09]
   {   if (sGulkwa[74] <> '') and (sGulkwa[73] <> '') Then
      begin
         if (sGulkwa[73] = '음성') and (sGulkwa[74] = '음성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG3';
         end
         else if (sGulkwa[73] = '음성') and (sGulkwa[74] = '양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG2';
         end
         else if (sGulkwa[73] = '음성') and (sGulkwa[74] = '약양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIG1';
         end
         else if (sGulkwa[73] = '양성') and (sGulkwa[74] = '양성')Then
         begin
            K := K + 1;
            tSokun[K] := 'RIGM';
         end
         else
         begin
            if (sGulkwa[73] <> '') Then
            begin
               if (sGulkwa[73] = '양성') Then
               begin
                  K := K + 1;
                  tSokun[K] := 'RIM ';
               end;
            end;
         end;
      end
      else
      begin
//46.Cod = RIG (S019 이 양성)
         if (sGulkwa[74] <> '') Then
         begin
            if (sGulkwa[74] = '양성') Then
            begin
               K := K + 1;
               tSokun[K] := 'RIG ';
            end;
         end;
//47.Cod = RIM (S018 이 양성)
         if (sGulkwa[73] <> '') Then
         begin
            if (sGulkwa[73] = '양성') Then
            begin
               K := K + 1;
               tSokun[K] := 'RIM ';
            end;
         end;
      end;
           }
//55.Cod = UK (U006 >= 10)

      If (sGulkwa[94] <> '') AND (sGulkwa[24] <> '')  Then
         Begin
            Val1 := StrToFloat(sGulkwa[94]);
            Val2 := StrToFloat(sGulkwa[24]);
            If (Val1 >= 10) AND (Val2 <= 199.9) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UK  ';
               End
           Else if (Val1 >= 10) AND (Val2 >= 200) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UKDM';
               End;
         End;

//73.Cod = OB1 (Y004 = 양성)
      If (sGulkwa[99] <> '') Then
         Begin
            If (sGulkwa[99] = '양성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'OB1 ';
               End;
         End;
//OBN (Y004 = 음성) 2014.07.22
      If (sGulkwa[99] <> '') Then
         Begin
            If (sGulkwa[99] = '음성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'OBN ';
               End;
         End;
//76.Cod = CRP (S001 = '양성')
      If (sGulkwa[67] <> '') Then
         Begin
            If (sGulkwa[67] = '양성') or (sGulkwa[67] = '약양성') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP ';
               End;
         End;


//89.Cod = HPX (C013 >= 350 And C005 >= 2.5)
      If (sGulkwa[05] <> '') And (sGulkwa[13] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[05]);
            Val2 := StrToFloat(sGulkwa[13]);
            If (Val2 >= 350) And (Val1 >= 2.5) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HPX ';
               End;
         End;
//100.Cod = BET(C044 >= 6.1), BET1(3.1<=C044<=6.0)
      If (sGulkwa[32] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[32]);
            If (Val1 > 2.37) And (Val1 <= 6.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BET1';
               End;
            If (Val1 >= 6.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BET ';
               End;
         End;
//xxx.Cod = CYFR (E040 >= 3.1)
      If (sGulkwa[104] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[104]);
            If (Val1 >= 3.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CYFR';
               End;
         End;
//xxx.Cod = Aldolase (C015 >= 3.1)  ==>  C015 >= 7.7 (외주항목으로 기본값이 0~7.6 임)
      If (sGulkwa[109] <> '') Then
      Begin
         Val1 := StrToFloat(sGulkwa[109]);
         If (Val1 >= 7.7) Then
         Begin
            K := K + 1;
            tSokun[K] := 'ALDL';
         End;
      End;
//xxx.Cod = IP (C057 >= 5.5)
    {  If (sGulkwa[110] <> '') Then
      Begin
         Val1 := StrToFloat(sGulkwa[110]);
         If (Val1 >= 5.5) Then
         Begin
            K := K + 1;
            tSokun[K] := 'PH3 ';
         End;
      End;  }




{//111. A형간염. Cod = HAV (S014 음성, S015 양성)
      If (sGulkwa[111] <> '') and (sGulkwa[112] <> '') Then
      Begin
         If (sGulkwa[111] = '음성') and
            (sGulkwa[112] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV ';
            End
         else If (sGulkwa[111] = '음성') and
            (sGulkwa[112] = '양성') and
            (sGulkwa[114] = '약양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVC';
            End
         Else If (sGulkwa[111] = '양성') and
            (sGulkwa[112] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV2';
            End
         Else If (sGulkwa[111] = '양성') and
            (sGulkwa[112] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV4';
            End
         Else If (sGulkwa[111] = '음성') and
            (sGulkwa[112] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
            End;
      End; }
      If (sGulkwa[182] <> '') Then
      Begin

          If (sGulkwa[182] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'M2P ';
            End
         else If (sGulkwa[182] = '약양성')  Then
            Begin
               K := K + 1;
               tSokun[K] := 'M2WP';
            End
         else If (sGulkwa[182] = '음성')  Then
            Begin
               K := K + 1;
               tSokun[K] := 'M2N ';
            End;
      End;

//111.se02 양성 se01 X   C009,C010 > 40 HAVG   ,C009,C010 <= 40 HAVH
      If (sGulkwa[114] <> '') and (sGulkwa[113] = '')Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;


          If (sGulkwa[114] = '양성') and  (cSex = 1) and ((Val2 >=45.1) or ((Val2 <=45) and (Val1 >=40.1)and (Val3 >=70.1)))  Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVG';
            End
         else If (sGulkwa[114] = '양성') and  (cSex = 2) and ((Val2 >=34.1) or ((Val2 <=34) and (Val1 >=32.1)and (Val3 >=42.1)))    Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVG';
            End

      End;
      If (sGulkwa[168] <> '') and (sGulkwa[113] = '')Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
          if  (sGulkwa[168]<> '')  Then  Val4 := StrToFloat(sGulkwa[168])
          else  Val4 :=0;


          If (Val4 >=20) and  (cSex = 1) and ((Val2 >=45.1) or ((Val2 <=45) and (Val1 >=40.1)and (Val3 >=70.1)))   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVL';
            End
         else If (Val4 >=20) and  (cSex = 2) and ((Val2 >=34.1) or ((Val2 <=34) and (Val1 >=32.1)and (Val3 >=42.1)))    Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVL';
            End

      End;


      If (sGulkwa[09]  <> '') and (sGulkwa[10]  <> '') and (sGulkwa[113] = '') Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
         If (Val2 <=45) and (Val1 <=40)  and (cSex = 1) and  (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
         else If (Val2 <=34) and (Val1 <=32)  and (cSex = 2) and  (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
         else If (Val2 <=45) and (Val1 >=40.1) and (Val3 <=70) and (cSex = 1) and    (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
          else If (Val2 <=34) and (Val1 >=32.1) and (Val3 <=42) and (cSex = 2) and   (sGulkwa[114] = '양성')  and 
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End;
      End;

      If  (sGulkwa[70] <> '')  Then
      Begin
         if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
             else val1 :=0;
         If (sGulkwa[70] = '약양성') or ((Val1>=1.0) and (val1<=9.99))then
            Begin
               K := K + 1;
               tSokun[K] := 'HBCW';
            End
      end;
      If  (sGulkwa[72] <> '')  Then
      Begin
         If (sGulkwa[72] = '약양성') then
            Begin
               K := K + 1;
               tSokun[K] := 'HCVW';
            End
      end;


//09.07.04. A형간염. Cod = HAV (SE01 음성, SE02 양성)
      If  (sGulkwa[114] <> '')  Then
      Begin
         If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV ';
            End
         Else If (sGulkwa[113] = '양성') and
            (sGulkwa[114] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV2';
            End
         Else If (sGulkwa[113] = '양성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV4';
            End
         Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
            End
           Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
            End
            Else If (sGulkwa[113] = '') and
            (sGulkwa[114] = '약양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVW';
            End
            Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '약양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVC';
            End;
      End;                
      If (trim(sGulkwa[113]) = '') and (sGulkwa[114] <> '') then
         begin
          if (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVB';
            End;
      end;
 //Cod = C003 >4.1)   ---20120704

      If (sGulkwa[03] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[03]);
            If (Val1 > 4.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'GH ';
               End;
      End;
  //Cod = C078 >15.0)  ---20120704
   {   If (sGulkwa[164] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[164]);
            If (Val1 > 15.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HOMO ';
               End;
         End;  }
  //Cod = H025 >15 여자 ,H025 >9 남자 )
    If (sGulkwa[117] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[117]);
            If (((Val1 > 9) And (cSex = 1)) Or
               ((Val1 > 15)And (cSex = 2))) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20131101') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'ESR ';
               End
            else If (((Val1 >= 15.1) And (cSex = 1)) Or
               ((Val1 >= 20.1)And (cSex = 2))) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20131101') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'ESR ';
               End;
         End;
//20121008
      If (sGulkwa[91] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[91]);
            If  (Val1 <= 4.9) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPHL';
               End
            else if (Val1 >= 8.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPHH';
               End;
         End;
         //
       If (sGulkwa[110] <> '') Then
         Begin
            Val2 := StrToFloat(sGulkwa[110]);
            if (sGulkwa[121] <> '') then    Val1 := StrToFloat(sGulkwa[121])
            else Val1 :=0;
            If  (((Val1 >= 8.5) and (Val1 <= 10.5)) or  (Val1=0))  and  (Val2 >= 4.6)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PH3 ';
               End
           else If  (((Val1 >= 8.5) and (Val1 <= 10.5)) or  (Val1=0))  and  (Val2 <= 2.49)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PL ';
               End;
         end;
         //
        If (sGulkwa[121] <> '')   Then
         Begin
            if (sGulkwa[110] <> '') then    Val2 := StrToFloat(sGulkwa[110])
            else val2:=0;
            if (sGulkwa[121] <> '') then    Val1 := StrToFloat(sGulkwa[121])
            else Val1 :=0;

          If  (val1>=10.51) and (((val2>=2.5) and (val2<=4.5)) or (val2=0)) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CAH';
               End
           else If (val1<=8.49) and (((val2>=2.5) and (val2<=4.5)) or (val2=0)) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CAL';
               End;
        end;
         //
        If (sGulkwa[121] <> '')  and (sGulkwa[110] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[121]);
            Val2 := StrToFloat(sGulkwa[110]);
            If  (Val1 >= 10.51) and (Val2 <= 2.49) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLCH';
               End
           else If (Val1 <= 8.49) and (Val2 >= 4.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PHCL';
               End
           else If (Val1 <= 8.49) and (Val2 <= 2.49)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLCL';
               End
           else If (Val1 >= 10.51) and (Val2 >= 4.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PHCH';
               End;
        end;
        //
        If (sGulkwa[08] <> '') and (sGulkwa[70] <> '')  Then
         Begin
            if  sGulkwa[115] <> '' then val1 := StrToFloat(sGulkwa[115])
            else val1 :=0;
            if (sGulkwa[55])<> '' then Val2 := StrToFloat(sGulkwa[55])
               else val2 :=0;

            If ((sGulkwa[70] = '음성') or ((val2>0) and (val2<1.0))) and
               (((sGulkwa[71] = '약양성') or (sGulkwa[71] = '양성'))or  (val1>=10))
                and ((sGulkwa[08] = '양성') or (sGulkwa[08] = '약양성'))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBC8';
               End
           else If ((sGulkwa[08] = '양성') or (sGulkwa[08] = '약양성')) and
               ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (Val2>=1.0)) and
               ((sGulkwa[71] = '음성') or ((val1>0) and (val1<=9.9))) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBC6';
               End
           else If ((sGulkwa[08] = '양성') or (sGulkwa[08] = '약양성')) and
               ((sGulkwa[70] = '음성') or ((val2>0) and (val2<1.0))) and
               ((sGulkwa[71] = '음성') or ((val1>0) and (val1<=9.9))) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBCA';
               End
       End;
       If (sGulkwa[70] <> '') and (sGulkwa[71] <> '')  Then
         Begin
           if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
           if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
               else val2 :=0;
           If ((sGulkwa[70] = '음성')  or ((val1>0) and (val1<1.0)))  and
              (((sGulkwa[71] = '약양성')) or (val2>=10) and (val2<30)) and
              ((sGulkwa[08] = '음성') or (sGulkwa[08] = '')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HABW';
               End
       End;


       If (sGulkwa[132] <> '') and (sGulkwa[70] <> '')  Then
         Begin
            if  sGulkwa[115] <> '' then val1 := StrToFloat(sGulkwa[115])
            else val1 :=0;
            if (sGulkwa[55])<> '' then Val2 := StrToFloat(sGulkwa[55])
            else val2 :=0;
            If ((sGulkwa[132] = '양성') or (sGulkwa[132] = '약양성'))    and
               (((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성')) or(Val2>=1.0) )and
               ((sGulkwa[71] = '음성') or ((val1>0) and (val1<=0.99))) and
               ((sGulkwa[133] = '음성') or (sGulkwa[133] = '')) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBE1';
               End
       End;
       If (sGulkwa[132] <> '') and (sGulkwa[71] <> '') and (sGulkwa[133] <> '')  Then
         Begin
            if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
            If ((sGulkwa[132] = '양성') or (sGulkwa[132] = '약양성')) and
               ((sGulkwa[70]  = '양성') or (sGulkwa[70]  = '약양성')  or (val1>=1.0)) and
               ((sGulkwa[133] = '양성') or (sGulkwa[133] = '약양성')) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBE4';
               End
       End;                             

       If (sGulkwa[132] <> '') and (sGulkwa[70] <> '') and (sGulkwa[133] <> '') Then
         Begin
            if  sGulkwa[115] <> '' then val1 := StrToFloat(sGulkwa[115])
            else val1 :=0;
            if (sGulkwa[55])<> '' then Val2 := StrToFloat(sGulkwa[55])
            else val2 :=0;
            If (sGulkwa[132] = '음성') and
               ((sGulkwa[133] = '양성') or (sGulkwa[133] = '약양성'))   and
               ((sGulkwa[70] = '양성')  or (sGulkwa[70] = '약양성') or (val2>=1.0)) and
               ((sGulkwa[71] = '음성') or (val1<=0.99)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBE7';
               End
           else  If (sGulkwa[132] = '음성') and
               ((sGulkwa[133] = '양성') or (sGulkwa[133] = '약양성')) and
               ((sGulkwa[70]   = '음성') or ((val2>0) and (val2<1.0)))  and
               ((sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성')  or (val1>=10.0)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBE5';
               End
           else  If (sGulkwa[132] = '음성') and
               ((sGulkwa[133] = '양성') or (sGulkwa[133] = '약양성')) and
               ((sGulkwa[70]   = '음성') or ((val2>0) and (val2<1.0)))  and
               ((sGulkwa[71] = '음성')   or (val1<=0.99)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HBE6';
               End
       End;
       If (sGulkwa[134] <> '') Then
         Begin 
            If (sGulkwa[134] = '양성') and
               ((sGulkwa[135] = '음성') or (sGulkwa[135] = ''))  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VZVG';
               End

       End;
       If (sGulkwa[134] <> '') Then
         Begin
            If (sGulkwa[134] = '약양성') and
               ((sGulkwa[135] = '음성') or (sGulkwa[135] = ''))  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VZVC';
               End

       End;
       If (sGulkwa[135] <> '') Then
         Begin
            If (sGulkwa[134] = '') and
               ((sGulkwa[135] = '양성') or (sGulkwa[135] = '약양성'))  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VZVM';
               End

       End;    


       If (sGulkwa[29] <> '') and (sGulkwa[30] <> '')  Then
         Begin
            val1 := StrToFloat(sGulkwa[29]);
            val2 := StrToFloat(sGulkwa[30]);

            If (Val1>=30.1) and (Val2<=1.4) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BUN ';
               End
            else If ((Val1>=26) and (Val1<=30)) and (Val2<=1.4) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BUN2';
               End
            else If (Val1>=20.1) and (Val2>=1.401) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'BCH ';
               End   
       End;
       If (sGulkwa[30] <> '')  Then
          Begin
            if (sGulkwa[29]) <> '' then      val1 := StrToFloat(sGulkwa[29])
            else val1 :=0;
            val2 := StrToFloat(sGulkwa[30]);
            If (Val1<=20.1) and (Val2>=1.401) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRH ';
               End
       End;

        If (sGulkwa[123] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[123]);
            If ((Val1 <= 1.89) and (cSex = 1)) and  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130422')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INL ';
               End
            else If ((Val1 >= 15.98) and (cSex = 1)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130422')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INH ';
               End
            else If ((Val1 <= 1.79) and (cSex = 2)) and  ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130422')   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INL ';
               End
           else If ((Val1 >= 8.30) and (cSex = 2)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130422')   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INH ';
               End;
        End;
       If (sGulkwa[123] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[123]);
            If ((Val1 <= 2.59) and (cSex = 1)) and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INL ';
               End
            else If ((Val1 >= 24.9) and (cSex = 1))  and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INH ';
               End
            else If ((Val1 <= 2.59) and (cSex = 2) )and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INL ';
               End
           else If ((Val1 >= 24.9) and (cSex = 2)) and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'INH ';
               End;
        End;
        //

        //
  {      If (sGulkwa[124] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[124]);
            If (Val1 >= 0.72) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'TRNI';
               End;
        End;}



        //
      {   If (sGulkwa[129] <> '') Then
         Begin
            val1:=0;
            val2:=0;
            val3:=0; 
            if (sGulkwa[127] <> '' ) then
               begin
               Val1 := StrToFloat(sGulkwa[127]);
               end;
            if (sGulkwa[128] <> '' ) then
               begin
               Val2 := StrToFloat(sGulkwa[128]);
               end;
            Val3 := StrToFloat(sGulkwa[129]);
            If ((Val3 >= 30.1) and  (Val3 <= 40))  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LPA1';
               End;
            else If (Val3 >= 40.1)   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LPA ';
               End;
        End;   }
        //
        
        If (sGulkwa[096] <> '') or (sGulkwa[095] <> '') Then
         Begin
            if sGulkwa[096] <> '' then  Val1 := StrToFloat(sGulkwa[096]);
            if sGulkwa[095] <> '' Then  Val2 := StrToFloat(sGulkwa[095]);

            If (Val1 >= 1.0) or (Val2 >= 1.0)   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UB1 ';
               End;
        End;



        ///sGulkwa[011] =C011 ,   sGulkwa[010] =C010 , sGulkwa[033] =C045 ,
        If (sGulkwa[011] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[033] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[011]);
            Val2 := StrToFloat(sGulkwa[010]);
            Val3 := StrToFloat(sGulkwa[033]);

            If (((Val1 >= 41) and (cSex  = 2)) or ((Val1 >= 51) and (cSex  = 1)) or  (Val2 >= 41))
            and (Val3 >= 301) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FH  ';
               End
            else If ((((Val1 >= 42.1)or (Val2 >= 34.1))  and (cSex  = 2)) or (((Val1 >= 70.1) or (Val2 >= 45.1))  and (cSex  = 1)))
            and (Val3 >= 300.1) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FH  ';
               End;
        End;

  ///sGulkwa[011] =C009 ,   sGulkwa[010] =C010 , sGulkwa[033] =C045 ,
         If (sGulkwa[011] <> '') and  (sGulkwa[010] <> '') and  (sGulkwa[033] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[011]);
            Val2 := StrToFloat(sGulkwa[010]);
            Val3 := StrToFloat(sGulkwa[033]);

            If (((Val1 <= 40) and (cSex  = 2)) or ((Val1 <= 50) and (cSex  = 1)))  AND    (Val2 <= 40)
                and (Val3 >= 301)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FE  ';
               End
            else If (((Val1 <= 42) and (cSex  = 2) AND    (Val2 <= 34) ) or ((Val1 <= 70) and (cSex  = 1) AND    (Val2 <= 45)))  
                and (Val3 >= 300.1)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FE  ';
               End
        End;
        If (sGulkwa[169] <> '')  Then
        Begin
             Val1 := StrToFloat(sGulkwa[169]);
             If val1>=10.1 then
                Begin
                     K := K + 1;
                     tSokun[K] := 'ALCH';
                End;
        End;

        //

        //40.Cod = PLH1 (530>= H008 <= 599)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            If (Val1 <= 599) And (Val1 >= 530) And  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End
            else If (Val1 >= 365.1) And (Val1 <= 400) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End
            else If (Val1 >= 398.1) And (Val1 <= 420) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End;

         End;
   {      If (tSokun[K] <> 'ANE2') Then
         If (sGulkwa[33] <> '') And (sGulkwa[44] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[33]);
            Val2 := StrToFloat(sGulkwa[44]);
            If (Val1 <= 59) And (Val2 <= 13.4) And (cSex = 1) And  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') then
            begin
               K := K + 1;
               tSokun[K] := 'IDA6';
            end
            else If (Val1 <= 59) And (Val2 <= 12.9) And (cSex = 1) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004')  then
            begin
               K := K + 1;
               tSokun[K] := 'IDA6';
            end;

         End;   }
       {  If (tSokun[K] <> 'ANE2') AND
         (tSokun[K] <> 'IDA ') AND
         (tSokun[K] <> 'IDA6') AND
         (tSokun[K] <> 'IDPM') Then
          If (sGulkwa[44] <> '') Then
          Begin
             Val1 := StrToFloat(sGulkwa[44]);
             if sAge <= 4 Then
             begin
                if Val1 <= 10.9 Then
                begin
                   K := K + 1;
                   tSokun[K] := 'ANE ';
                end;
             end
             else if (sAge >= 5) and (sAge <= 9) Then
             begin
                if Val1 <= 11.4 Then
                begin
                   K := K + 1;
                   tSokun[K] := 'ANE ';
                end;
             end
             else if (sAge >= 10) and (sAge <= 14) Then
             begin
                if Val1 <= 11.9 Then
                begin
                   K := K + 1;
                   tSokun[K] := 'ANE ';
                end;
             end
             else if sAge >= 15 Then
             begin
                if cSex = 1 then
                begin
                   if (Val1 <= 12.9)and ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'ANE ';
                   end
                   else if (Val1 <= 11.9) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'ANE ';
                   end;

                   if ((Val1 >= 13.0) and (Val1 <= 13.4)) and  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'IN06';
                   end
                   else if (Val1 >= 12.0) and (Val1 <= 12.9) and ((qryBunju.FieldByName('Dat_Gmgn').AsString) >= '20121004') Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'IN06';
                   end;
                end
                else if cSex = 2 then
                begin
                   if Val1 <= 10.9 Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'ANE ';
                   end;
                   if (Val1 >= 11.0) and (Val1 <= 11.9) Then
                   begin
                      K := K + 1;
                      tSokun[K] := 'IN06';
                   end;
                end;
             End;
          End;    }



//70.Cod = WBC1 (10.1 <=H011 >= 11.0)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            If (Val1 >= 10.1) AND (Val1 <= 11.0) and  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End
            else If (Val1 >= 10.21) AND (Val1 <= 11.0) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End
            else If (Val1 >= 9.81) AND (Val1 <= 11.0) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC1';
               End;
         End;

//40.Cod = PLH (H008 >= 520)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            If (Val1 >= 520) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End
           else If (Val1 >= 400.1) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End
           else If (Val1 >= 420.1) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End;
         End;

         If (sGulkwa[44] <> '') and (cSex = 2) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;
            if  (sGulkwa[34]) <> '' then  Val3 := StrToFloat(sGulkwa[34])
            else val3:=0;
            Val4 := StrToFloat(sGulkwa[44]);
            If ((Val4 >= 8.0) and (Val4 <= 11.99)) and ((Val1 >= 0.001) and (Val1 <= 9.99))  and (sAge <= 50) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
               End
           else If ((Val4 >= 8.0) and (Val4 <= 11.99)) and ((Val1 >=10)or  (Val1 =0))  and ((val2>=0.01) and  (val2<=54.9)) and ((val3>=0.001)and (Val3<=14.9)) and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
           else If ((Val4 >= 8.0) and (Val4 <= 11.99))  and  ((Val1 >= 0.001) and (Val1 <= 9.99))  and (val2=0) and (Val3=0) and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
           else If ((Val4 >= 8.0) and (Val4 <= 10.99)) and (Val1 >=10.0) and  ((val2>=0.01) and  (val2<=54.9)) and ((Val3>=15.0) or (Val3=0)) and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
           else If ((Val4 >= 8.0) and (Val4 <= 10.99)) and (Val1 =0) and  ((val2>=0.01) and  (val2<=54.9)) and ((Val3>=15.0) or (Val3=0)) and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End;
         End;
          If (sGulkwa[44] <> '') and (cSex = 2) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;
            if  (sGulkwa[34]) <> '' then  Val3 := StrToFloat(sGulkwa[34])
            else val3:=0;
            Val4 := StrToFloat(sGulkwa[44]);
            If ((Val4 >= 8.0) and (Val4 <= 11.99)) and ((Val1 >= 0.001) and (Val1 <= 9.99))  and (sAge <= 50) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
               End
           else If ((Val4 >= 11.0) and (Val4 <= 11.9)) and ((val2>=0.01) and  (val2<=49.9)) and ((val3>=0.001)and (Val3<=14.9)) and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
          { else If ((Val4 >= 8.0) and (Val4 <= 10.99)) and   ((val2>=0.01) and  (val2<=49.9)) and (Val3>=15.0)  and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End }
           else If ((Val4 >= 8.0) and (Val4 <= 10.9)) and ((val2>=0.01) and  (val2<=49.9)) and   (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End;
         End;
       { If (sGulkwa[44] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>='20130518')  then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;
            if  (sGulkwa[34]) <> '' then  Val3 := StrToFloat(sGulkwa[34])
            else val3:=0;
            Val4 := StrToFloat(sGulkwa[44]);

           If (cSex = 1) and((Val4 >= 8.0) and (Val4 <= 11.99)) and ((Val1 >=10)or  (Val1 =0))  and ((val2>=0.01) and  (val2<=64.9)) and
              ((val3>=0.001)and (Val3<=14.9)) and (sAge <= 50) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
           else If (cSex = 2) and ((Val4 >= 8.0) and (Val4 <= 11.99)) and ((Val1 >=10)or  (Val1 =0))  and ((val2>=0.01) and  (val2<=49.9)) and
                   ((val3>=0.001)and (Val3<=14.9)) and (sAge <= 50)  then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End

           else If (cSex = 1) and ((Val4 >= 8.0) and (Val4 <= 10.99)) and ((Val1 >=10.0) or (Val1 =0)) and  ((val2>=0.01) and  (val2<=64.9)) and (Val3>=15.0)  and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End
           else If (cSex = 2) and ((Val4 >= 8.0) and (Val4 <= 10.99)) and ((Val1 >=10.0) or (Val1 =0)) and  ((val2>=0.01) and  (val2<=49.9)) and (Val3>=15.0)  and (sAge <= 50) then
                Begin
                  K := K + 1;
                  tSokun[K] := 'IDA1';
                End;

         End;  }
         If (sGulkwa[33] <> '')  and  (sGulkwa[44] <> '') then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            Val2 := StrToFloat(sGulkwa[33]);
            if  (sGulkwa[34]) <> '' then  Val3 := StrToFloat(sGulkwa[34])
            else val3:=0;
            Val4 := StrToFloat(sGulkwa[44]);
            If (cSex = 2) and (sAge <= 50) and
               ((Val1>=10) or (Val1=0)) and
               (Val2<=54.9) and
               ((Val3>=15.0) or (Val3=0))  and
               ((Val4>=11.0) and (Val4<=11.99)) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA ';
               End
            else If (cSex = 2) and (sAge <= 50) and
               (Val2<=49.9) and
                ((Val1>=10) or (Val1=0)) and
                ((Val3>=15.0) or (Val3=0))  and
               ((Val4>=11.0) and (Val4<=11.99)) and
               ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA ';
               End
         End;
         If  (sGulkwa[44] <> '')  and (cSex = 1) then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else Val2:=0;
            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=0.001)  and  (val1<=16.99)) and (Val4<=7.9) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDAG';
               End
            else If ((val2>=0.001)  and  (val2<=59.9)) and (Val4<=7.9)  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDAG';
               End
            else If ((val1>=17) or (val1=0)) and
                    ((val2>=0.001) and (Val2<=59.9)) and
                    (Val4<=7.99) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDAG';
               End;

         End;
         If  (sGulkwa[44] <> '') and   (cSex = 2) and (sAge <= 50)  then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else Val2:=0;
            if (sGulkwa[34]) <> '' then   Val3 := StrToFloat(sGulkwa[34])
            else Val3:=0;

            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=0.001) and (val1<=9.99)) and
               (Val4<=7.99)  and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA8';
               End
            else If ((val1>=10) or (val1=0))and
                 ((Val2>=0.001) and (val2<=54.9))and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  and
                 (Val4<=7.99) then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'IDA8';
            End
            else If ((Val2>=0.001) and (val2<=49.9))and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  and
                 (Val4<=7.99) then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'IDA8';
            End
            else If ((Val1>=0.001) and (val1<=9.9))and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  and
                 (Val4<=7.99) then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'IDA8';
            End;
         End;

         If  (sGulkwa[44] <> '') and  (cSex = 1)    then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else Val2:=0;

            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=0.001) and (val1<=16.99)) and
               ((Val4>=8.0)   and (Val4<=12.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDA6';
               End
           else If ((val2>=0.001) and (val2<=59.9))  and
                ((Val4>=8.0) and (Val4<=12.9)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'IDA6';
            End
            else If ((val1>=0.001) and (val1<=16.9)) and
                ((Val4>=8.0) and (Val4<=12.9)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                 K := K + 1;
                 tSokun[K] := 'IDA6';
            End;
         End;


         If  (sGulkwa[36] <> '' ) and (sGulkwa[44] <> '') and   (cSex = 2) and (sAge <=50)   then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            if sGulkwa[33] <> '' then    Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;
            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=9.99) and
               ((val2>=55) or (Val2=0)) and
               (Val4>=12)  And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FEL2';
               End
            else If (val1<=9.99) and
               ((val2>=50) or (Val2=0)) and
               (Val4>=12)  And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FEL2';
               End;
         End;
         If  (sGulkwa[36] <> '' ) and (sGulkwa[44] <> '')  and (cSex = 2) and (sAge >=51)   then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            if sGulkwa[33]<> '' then    Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;

            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=9.99) and
               ((val2>=55) or (Val2=0)) and
               (Val4>=12) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FELM';
               End
            else If (val1<=9.99) and
               ((val2>=50) or (Val2=0)) and
               (Val4>=12) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FELM';
               End;
         End;
         If  (sGulkwa[36] <> '' ) and (sGulkwa[44] <> '')  and (cSex = 1)    then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            if sGulkwa[33]<> '' then    Val2 := StrToFloat(sGulkwa[33])
            else val2:=0;

            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=16.99) and
               ((val2>=60) or (Val2=0)) and
               (Val4>=13) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FEL4';
               End
            else If (val1<=16.99) and
               ((val2>=60) or (Val2=0)) and
               (Val4>=13) And ((qryBunju.FieldByName('Dat_bunju').AsString) >= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FEL4';
               End;
         End;

         If  (sGulkwa[44] <> '') and (sGulkwa[33] <> '')and (cSex = 2) and (sAge <=50)  then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=10) or (val1=0)) and
               (val2<=54.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO ';
               End
            else If ((val1>=10) or (val1=0)) and
               (val2<=49.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO ';
               End;
         End;

        If  (sGulkwa[44] <> '') and (sGulkwa[33] <> '')and  (cSex = 1)  then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;      
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=17) or (val1=0)) and
               (val2<=59.9) and
               (Val4>=13) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO2';
               End
            else If ((val1>=17) or (val1=0)) and
               (val2<=59.9) and
               (Val4>=13) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO2';
               End;
         End;

         If  (sGulkwa[44] <> '') and (sGulkwa[33] <> '') and (cSex = 2) and (sAge >=51)  then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;   
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=10) or (val1=0)) and
               (val2<=54.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDOM';
               End
            else If ((val1>=10) or (val1=0)) and
               (val2<=49.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDOM';
               End;
         End;

         If (sGulkwa[36] <> '') and (sGulkwa[44] <> '') and (sGulkwa[33] <> '') and (cSex = 2) and (sAge <=50)  then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=9.99) and
               (val2<=54.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO3';
               End
            else If (val1<=9.99) and
               (val2<=49.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO3';
               End;
         End;

         If (sGulkwa[36] <> '') and (sGulkwa[44] <> '') and (sGulkwa[33] <> '') and (cSex = 1) then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=16.99) and
               (val2<=59.9) and
               (Val4>=13) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO4';
               End
            else If (val1<=16.99) and
               (val2<=59.9) and
               (Val4>=13) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDO4';
               End;
         End;

         If (sGulkwa[36] <> '') and (sGulkwa[44] <> '') and (sGulkwa[33] <> '') and (cSex = 2) and (sAge >=51)then
         Begin
            Val1 := StrToFloat(sGulkwa[36]);
            Val2 := StrToFloat(sGulkwa[33]);

            Val4 := StrToFloat(sGulkwa[44]);
            If (val1<=9.99) and
               (val2<=54.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDON';
               End
            else If (val1<=9.99) and
               (val2<=49.9) and
               (Val4>=12) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDON';
               End;
         End;

         If  (sGulkwa[44] <> '') and   (cSex = 2) and (sAge >=51)then
         Begin
            if (sGulkwa[36]) <> '' then   Val1 := StrToFloat(sGulkwa[36])
            else val1:=0;
            if (sGulkwa[33]) <> '' then   Val2 := StrToFloat(sGulkwa[33])
            else Val2:=0;         
   
            Val4 := StrToFloat(sGulkwa[44]);
            If ((val1>=0.01) and (val1<=9.99)) and
               ((Val4>=8.0) and (Val4<=11.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'IDPM';
               End
            else If ((val1>=10.0) or (val1=0)) and
                    ((Val2>=0.001)  and (Val2<=54.9))  and
                    ((Val4>=8.0) and (Val4<=11.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDPM';
            End
           { else If ((val1>=10.0) or (val1=0)) and
                    ((Val2>=0.001)  and (Val2<=49.9))  and
                    ((Val4>=8.0) and (Val4<=11.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDPM';
            End}
            else If ((val1>=0.001) and  (val1<=9.9))  and
                    ((Val4>=8.0) and (Val4<=11.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDPM';
            End
             else If ((Val2>=0.001)  and (Val2<=49.9)) and
                    ((Val4>=8.0) and (Val4<=11.99)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDPM';
            End
            else If ((val1>=0.001) and  (val1<=9.9)) and
                    ((Val4>=0.01) and (Val4<=7.9)) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDAH';
            End
             else If ((Val2>=0.001)  and (Val2<=49.9)) and
                    ((Val4>=0.01) and (Val4<=7.9))  and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') then
                 Begin
                      K := K + 1;
                      tSokun[K] := 'IDAH';
            End;



           { else  If ((val1>=0.01) and (val1<=9.99)) and
                     (Val4<=7.99) then
                  Begin
                       K := K + 1;
                       tSokun[K] := 'IDAH';
            End
            else  If ((val1>=10.0) or  (val1=0)) and
                     ((Val2>=0.001)  and (Val2<=54.9))  and
                     (Val4<=7.99) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') then
                  Begin
                       K := K + 1;
                       tSokun[K] := 'IDAH';
            End
             else  If ((val1>=10.0) or  (val1=0)) and
                     ((Val2>=0.001)  and (Val2<=49.9))  and
                     (Val4<=7.99) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')then
                  Begin
                       K := K + 1;
                       tSokun[K] := 'IDAH';
            End }
         End; 
         //UPDM
         If (sGulkwa[92] <> '') and (qryBunju.FieldByname('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;
            If ((Val1 >= 10) And (Val1 <= 30)) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1'))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
            else  If (Val1 >= 100) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 <= 1.40) and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End;
         End;
         //////////UPDM
          If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;
            If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 1)  and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 2)  and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
            else  If (Val1 >= 10) and (Val2 <=10) and (Val3 <= 5) and   (sGulkwa[90] = '양성')  and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1)  and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)  and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
        End;


         If (sGulkwa[92] <> '') and (qryBunju.FieldByname('Dat_Bunju').AsString < '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            If (Val1 >= 30) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = ''))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               End
            else  If (Val1 >= 30) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 >= 1.401) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH1';
               End;
         End;
         If (sGulkwa[92] <> '') and (qryBunju.FieldByname('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89]<> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97]<> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            if sGulkwa[24]<> '' then  Val5 := StrToFloat(sGulkwa[24])
            else  val5:=0;
            if sGulkwa[37]<> '' then  Val6 := StrToFloat(sGulkwa[37])
            else  val6:=0;
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;

            If ((Val1 >= 10) And (Val1 <= 30)) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) AND (Val5<=125.9) and ((Val6<=6.4) OR (Val6 = 0)) and  ((Chk_Dang_Drug <> '1')AND(Chk_Dang_Jindan <> '1'))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               End
            else  If (Val1 >= 100) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 <= 1.40) AND (Val5<=125.9) and ((Val6<=6.5) OR (Val6 = 0)) AND
               ((Chk_Dang_Drug <> '1')AND(Chk_Dang_Jindan <> '1')) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH1';
               End;
         End;
         If (sGulkwa[92] <> '') And (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;
            If (Val1 >= 10) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1'))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End
            else  If (Val1 >= 10) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 >= 1.401) and ((Chk_Dang_Drug = '1')or(Chk_Dang_Jindan ='1')) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPDM';
               End;
         End;

         If (sGulkwa[92] <> '') and (qryBunju.FieldByname('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
             If (Val1 >= 10) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 >= 1.41) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPHR';
               End;
         End;


         If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
            (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 <= 20) and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWB ';
               End
            else  If (Val1 <= 20) and (Val2 >= 25) and (Val3 <= 5) and  (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWBC';
               End
            else  If (Val1 <= 20) and (Val2 <= 10) and (Val3 <= 5) and  (sGulkwa[90] = '양성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End;
         End;
         If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
            (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 <= 5) and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWB ';
               End
            else  If (Val1 <= 5) and (Val2 >= 25) and (Val3 <= 5) and  (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWBC';
               End
            else  If (Val1 <= 5) and (Val2 <= 10) and (Val3 <= 5) and  (sGulkwa[90] = '양성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End;
         End;

        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 30) and (Val2 <= 10) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPB ';
               End
            else If (Val1 >= 30) and (Val2 <= 10) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPBW';
               End
            else  If (Val1 <= 20) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW ';
               End
            else  If (Val1 <= 20) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2) and  (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW5';
               End
             else  If (Val1 <= 20) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2) and  (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW6';
               End;
        End;
        Chk_Dang_Drug := '';         //통합문진 당뇨 확인
        Chk_Dang_Jindan := '';
        With QryTot_Munjin2018 Do
        Begin
           Close;
           ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
           ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
           ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
           Open;
           If RecordCount > 0 Then
           Begin
              Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
              Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
           End;
           Close;
        End;
        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 10) and (Val2 <= 10) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPB ';
               End
            else If (Val1 >= 10) and (Val2 <= 10) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPBW';
               End
            else  If (Val1 <= 5) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW ';
               End
            else  If (Val1 <= 5) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2) and  (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW5';
               End
             else  If (Val1 <= 5) and (Val2 >= 25) and (Val3 >= 10) and  (sGulkwa[90] = '음성') and (cSex = 2) and  (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW6';
               End;
        End;

        If (sGulkwa[92] <> '') and  (sGulkwa[97] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then Val2 := StrToFloat(sGulkwa[89])
            else val2:=0;
            Val3 := StrToFloat(sGulkwa[97]);

            If (Val1 <= 20) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and   ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URM ';
               End
            else  If (Val1 <= 20) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and   ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (cSex = 2) and  (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB ';
               End
            else  If (Val1 <= 20) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and  ((sGulkwa[90] = '음성') or (sGulkwa[90] = '') )and (cSex = 2) and  (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB2';
               End
        end;
        If (sGulkwa[92] <> '') and  (sGulkwa[97] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then Val2 := StrToFloat(sGulkwa[89])
            else val2:=0;
            Val3 := StrToFloat(sGulkwa[97]);

            If (Val1 <= 5) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and   ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URM ';
               End
            else  If (Val1 <= 5) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and   ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (cSex = 2) and  (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB ';
               End
            else  If (Val1 <= 5) and ((Val2 <= 10) or (Val2 = 0)) and (Val3 >= 10) and  ((sGulkwa[90] = '음성') or (sGulkwa[90] = '') )and (cSex = 2) and  (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB2';
               End
        end;


        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 30) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                  K := K + 1;
                  tSokun[K] := 'UWB ';
               End
            else  If (Val1 >= 30) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                  K := K + 1;
                  tSokun[K] := 'UWBC';
               End
             else  If (Val1 <= 20) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URM ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 <= 20) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 2) and (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 <= 20) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 2) and (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB2';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 >= 30) and (Val2 <=10) and (Val3 <= 5) and   (sGulkwa[90] = '양성')    then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 >= 30) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '음성') and (cSex = 1)    then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
               End
            else  If (Val1 >= 30) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW7';
               End
           else  If (Val1 <= 20) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 <= 20) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)  and (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW5';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 <= 20) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)  and (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW6';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 30) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPB ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 30) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                  K := K + 1;
                  tSokun[K] := 'UWN ';
               End
          else  If (Val1 >= 30) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                  K := K + 1;
                  tSokun[K] := 'UWN1';
               End

          else  If (Val1 >= 30) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 30) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW7';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
        End;
         If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and  (sGulkwa[24] <> '')  and
            (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            Val4 := StrToFloat(sGulkwa[24]);
            if (sGulkwa[37] <> '') then Val5 := StrToFloat(sGulkwa[37])
            else Val5 := 0;

            If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (Val4 <=125.9) and ((Val5 <= 6.4) or (Val5 = 0)) and
            ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               end

            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (Val4 <=125.9) and ((Val5 <= 6.4) or (Val5 = 0))  and
            ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';

               End
             else  If (Val1 >= 10) and (Val2 <=10) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (Val4 <=125.9) and ((Val5 <= 6.4) or (Val5 = 0))
                  and  ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1'))  then
                Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
                End
             else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1) and (Val4 <=125.9) and ((Val5 <= 6.4) or (Val5 = 0))
                  and  ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               END
            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)  and (Val4 <=125.9) and ((Val5 <= 6.4) or (Val5 = 0))
                 and  ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1')) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               END
        End;
        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
            Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWB ';
               End
            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWBC';
               End
             else  If (Val1 <= 5) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URM ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 <= 5) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 2) and (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 <= 5) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 2) and (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URB2';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 >= 10) and (Val2 <=10) and (Val3 <= 5) and   (sGulkwa[90] = '양성')    then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '음성') and (cSex = 1)    then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
               End
            else  If (Val1 >= 10) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '음성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW7';
               End
           else  If (Val1 <= 5) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW ';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 <= 5) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)  and (sAge <=50)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW5';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 <= 5) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)  and (sAge >=51)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW6';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 10) and (Val2 <=10) and (Val3 >= 10) and   (sGulkwa[90] = '양성')   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN ';
               End
          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN1';
               End

          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
          else  If (Val1 >= 10) and (Val2 >=25) and (Val3 >= 10) and   (sGulkwa[90] = '양성')  and (cSex = 2)   then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URW7';
                  K := K + 1;
                  tSokun[K] := 'NIH ';
               End
        End;


        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 <= 20)and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN ';
               End
            else If (Val1 <= 20)and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN1';
               End;
         End;
        If (sGulkwa[92] <> '') and (sGulkwa[89] <> '') and (sGulkwa[97] <> '') and (sGulkwa[90] <> '') And
           (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            Val2 := StrToFloat(sGulkwa[89]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 <= 5)and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 1)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN ';
               End
            else If (Val1 <= 5)and (Val2 >= 25) and (Val3 <= 5) and   (sGulkwa[90] = '양성') and (cSex = 2)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UWN1';
               End;
         End;

         If (sGulkwa[141] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[141]);
            If (Val1 >= 60.01) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'LIP ';
               End;
         End;


         If (sGulkwa[142] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[142]);
            If (Val1 >= 28.01) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CA15';
               End;
         End;

          If (sGulkwa[02] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[02]);
            If (Val1 <= 2.99) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'TPL ';
               End;
         End;
         If (sGulkwa[143] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[143]);
            If (Val1 >= 115.01) and (sAge <=39) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NBNP';
               End
            else  If (Val1 >= 172.01) and ((sAge >=40) and (sAge <=49)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NBNP';
               End
            else  If (Val1 >= 263.01) and ((sAge >=50) and (sAge <=59)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NBNP';
               End
             else  If (Val1 >= 738.01) and (sAge >=60)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NBNP';
               End
         End;

         If (sGulkwa[106] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[106]);
            If (Val1 > 16.3)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NSE ';
               End
            Else if (Val1 <= 16.3) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NSEN';
               End ;
         End;

         if (sGulkwa[148] <> '') Then
         begin
             Val1 := StrToFloat(sGulkwa[148]);
             if ((Val1 >=15.201) or (Val1<=4.039)) and (cSex = 1) then
             begin
                 K := K + 1;
                 tSokun[K] := 'PRO1';
             end
             else if ((Val1 >=23.301) or (Val1<=4.789)) and (cSex = 2) then
             begin
                 K := K + 1;
                 tSokun[K] := 'PRO1';
             end;
         end;

 ///////////////////////////////////////////////////////////////////////            
         if (cSex = 1) and (sGulkwa[166] <> '') Then
         Begin
             Val1 := StrToFloat(sGulkwa[166]);

             if (val1 < 40) or (val1 >115) then
             begin
                 K := K + 1;
                 tSokun[K] := 'HOR3';
             end;
         End;
         if (cSex = 1) and  (sGulkwa[144] <> '') Then
         Begin
             Val2 := StrToFloat(sGulkwa[144]);

             if (val2 <5.65) or (val2 >44.19) then
             begin
                 K := K + 1;
                 tSokun[K] := 'HOR3';
             end;
         End;
         if (cSex = 1) and (sGulkwa[145] <> '') Then
         Begin
             Val3 := StrToFloat(sGulkwa[145]);

             if (val3 <1.50) or (val3 >12.4) then
             begin
                 K := K + 1;
                tSokun[K] := 'HOR3';
             end;
         End;
         if (cSex = 1) and (sGulkwa[146] <> '')  Then
         Begin
             Val4 := StrToFloat(sGulkwa[146]);

             if (val4 <1.7) or (val4 >8.6) then
             begin
                 K := K + 1;
                 tSokun[K] := 'HOR3';
             end;
         End;
         if (cSex = 1) and (sGulkwa[147] <> '') Then
         Begin
             Val5 := StrToFloat(sGulkwa[147]);

             if  (val5 <0.2) or (val5>1.4)then
             begin
                 K := K + 1;
                 tSokun[K] := 'HOR3';
             end;
         End;
 //////////////////////////////////////////////////////////////////////
       {  if (sGulkwa[125] <> '') Then
         begin
              Val1 := StrToFloat(sGulkwa[125]);
              if val1>=0.161 then
              begin
                 K := K + 1;
                 tSokun[K] := 'TRNI';
             end;
         end;     }
{         if (sGulkwa[152] <> '') or (sGulkwa[153] <> '')   or (sGulkwa[154] <> '') or
            (sGulkwa[155] <> '') Then
             begin
                 K := K + 1;
                 tSokun[K] := 'HBDN';
         end;}
         if (sGulkwa[153] <> '')   or (sGulkwa[154] <> '') or   (sGulkwa[155] <> '') Then
             begin
                 K := K + 1;
                 tSokun[K] := 'HBDN';
         end;  



         if (sGulkwa[53] <> '') and (sGulkwa[50] <> '')   Then
             begin
              Val1 := StrToFloat(sGulkwa[53]);
              Val2 := StrToFloat(sGulkwa[50]);
              if ((val1<=3.79) and (val2<=139.9))and (cSex = 1)  then
              begin
                 K := K + 1;
                 tSokun[K] := 'WPL1';
             end
             else if ((val1<=3.69) and (val2<=137.9))and (cSex = 2)  then
              begin
                 K := K + 1;
                 tSokun[K] := 'WPL1';
             end;
         end;



         //72.Cod = WBL (H011 <= 3.0)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            if sGulkwa[50] <> '' then  Val2 := StrToFloat(sGulkwa[50])
            else Val2 :=0;

             If (Val1 <= 2.99) and ((Val2>=140) or (val2=0))   and (cSex = 1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBL ';
               End
             else If (Val1 <= 2.99) and ((Val2>=138) or (val2=0)) and (cSex = 2)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBL ';
               End;
         End;
//72-1.Cod = WBL1 (H011 >= 3.1) And (H011 <= 3.3)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            if sGulkwa[50] <> '' then  Val2 := StrToFloat(sGulkwa[50])
            else Val2 :=0;

            If (Val1 >= 3.0) And (Val1 <= 3.79) And ((Val2>=140) or (val2=0)) and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBL1';
               End
            else If (Val1 >= 3.0) And (Val1 <= 3.69) And ((Val2>=138) or (val2=0))  and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBL1';
               End;
         End;

         If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            if sGulkwa[53] <> ''  then Val2 := StrToFloat(sGulkwa[53])
            else Val2:=0;
            If (Val1 >= 120) and (Val1 <= 139.9) and  ((Val2>=3.80) or (Val2=0)) and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLL1';
            End
            else If (Val1 >= 120) and (Val1 <= 137.9) and  ((Val2>=3.70) or (Val2=0)) and (cSex = 2)Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLL1';
            End;
         End;
         //41.Cod = PLL (H008 <= 139)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            if sGulkwa[53] <> ''  then Val2 := StrToFloat(sGulkwa[53])
            else Val2:=0;
             If (Val1 <= 119.9) and  ((Val2>=3.80) or (Val2=0)) and (cSex = 1)   Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLL ';
            End
            else If (Val1 <= 119.9) and  ((Val2>=3.70) or (Val2=0)) and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLL ';
            End;
         End;
      If (sGulkwa[86] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[86]);
            If (Val1 >= 1.491) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20131018')  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'SCC ';
               End
            else if  (Val1 >= 2.01) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20131018') then
            begin
                K := K + 1;
                tSokun[K] := 'SCC ';
            end;
            If  Val1 <= 2.0  then
            begin
                K := K + 1;
                tSokun[K] := 'SCCN';
            end;
         End;
    if (sGulkwa[24] <> '') AND (sGulkwa[26] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[26]);
            if (Val1 <= 99.9) And (Val2 >= 2.851) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
            begin
               K := K + 1;
               tSokun[K] := 'DIO ';
            end
            else if (Val1 <= 99.9) And (Val2 >= 285.1) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            begin
               K := K + 1;
               tSokun[K] := 'DIO ';
            end;
        end;
    if (sGulkwa[24] <> '') AND (sGulkwa[26] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[26]);
            if ((Val1 >= 100) and (Val1 <= 125.9)) And (Val2 >= 2.851) and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')  Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC5';
            end
            else if ((Val1 >= 100) and (Val1 <= 125.9)) And (Val2 >= 285.1) and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')  Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC5';
            end;
        end;
     if (sGulkwa[24] <> '') AND (sGulkwa[26] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[26]);
            if ((Val1 >= 126) and (Val1 <= 199.9)) And (Val2 >= 2.851) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC ';
            end
            else if ((Val1 >= 126) and (Val1 <= 200)) And (Val2 >= 285.1) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC ';
            end;
        end;
     if (sGulkwa[24] <> '') AND (sGulkwa[26] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[26]);
            if (Val1 >= 200) And (Val2 >= 2.851) And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC4';
            end
            else if (Val1 >= 200.1) And (Val2 >= 285.1) And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            begin
               K := K + 1;
               tSokun[K] := 'DIC4';
            end;
        end;
      //C020 >=5.01    ..20160806 자체검사로 인한 참고치 변경
      if (sGulkwa[38] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[38]);
            if (Val1 >= 5.01) Then
            begin
               K := K + 1;
               tSokun[K] := 'CKMB';
            end;
        end;
      //C020 <=5.00
        if (sGulkwa[38] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[38]);
            if (Val1 <= 5.00) Then
            begin
               K := K + 1;
               tSokun[K] := 'CMBN';
            end;
        end;
        //C038 >=7.01
        if (sGulkwa[172] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[172]);
            if (Val1 > 7) Then
            begin
               K := K + 1;
               tSokun[K] := 'INAB';
            end;
        end;
        if (sGulkwa[09] <> '') and (sGulkwa[10] <> '') and
           (sGulkwa[11] <> '') and (sGulkwa[13] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[13]);
            if (Val1 <= 40) and (Val2 <= 40) and (Val3 <= 50) and
               ((Val4 >=300.1) and (Val4 <= 500.0)) and (cSex = 1) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end
            else if (Val1 <= 39.9) and (Val2 <= 39.9) and (Val3 <= 49.9) and
               ((Val4 >=300.1) and (Val4 <= 500.0)) and (cSex = 2) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end;
        end;
        if (sGulkwa[09] <> '') and (sGulkwa[10] <> '') and
           (sGulkwa[11] <> '') and (sGulkwa[13] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[13]);
            if (Val1 <= 40) and (Val2 <= 45) and (Val3 <= 70) and
               ((Val4 >=150.1) and (Val4 <= 250.0)) and (cSex = 1) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end
            else if (Val1 <= 32) and (Val2 <= 34) and (Val3 <= 42) and
                ((Val4 >=150.1)  and (Val4 <= 250)) and (cSex = 2) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end;
        end;
        if (sGulkwa[09] <> '') and (sGulkwa[10] <> '') and
           (sGulkwa[11] <> '') and (sGulkwa[13] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518')Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[13]);
            if (Val1 <= 40) and (Val2 <= 40) and (Val3 <= 50) and
               (Val4 >=500.1)and (cSex = 1) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH1';
            end
            else if (Val1 <= 39.9) and (Val2 <= 39.9) and (Val3 <= 49.9) and
               (Val4 >=500.1) and (cSex = 2) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH1';
            end;
        end;
        if (sGulkwa[09] <> '') and (sGulkwa[10] <> '') and
           (sGulkwa[11] <> '') and (sGulkwa[13] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518')Then
         Begin
            Val1 := StrToFloat(sGulkwa[09]);
            Val2 := StrToFloat(sGulkwa[10]);
            Val3 := StrToFloat(sGulkwa[11]);
            Val4 := StrToFloat(sGulkwa[13]);
            if (Val1 <= 40) and (Val2 <= 45) and (Val3 <= 70) and
               (Val4 >=250.1)and (cSex = 1) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH1';
            end
            else if (Val1 <= 32) and (Val2 <= 34) and (Val3 <= 42) and
               (Val4 >=250.1) and (cSex = 2) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH1';
            end;
        end;

     {   if  (sGulkwa[13] <> '') Then
         Begin
            if (sGulkwa[09]) <> '' then
                Val1 := StrToFloat(sGulkwa[09])
            else val1 :=0;
            if (sGulkwa[10]) <> '' then
                Val2 := StrToFloat(sGulkwa[10])
            else val2 :=0;
            if (sGulkwa[11]) <> '' then
               Val3 := StrToFloat(sGulkwa[11])
            else val3:=0;
            Val4 := StrToFloat(sGulkwa[13]);


            if ((Val1 >= 40.1) or (Val2 >= 40.1) or (Val3 >= 50.1)) and
               (Val4 >=500.1)  and (cSex = 1) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end
            else if ((Val1 >= 40.0) or (Val2 >= 40.0) or (Val3 >= 50.0)) and
               (Val4 >=500.1) and (cSex = 2) and  (sAge >=20)   Then
            begin
               K := K + 1;
               tSokun[K] := 'APH ';
            end;

        end; }
        if  (sGulkwa[13] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)< '20130518') Then
         Begin
            Val1 := StrToFloat(sGulkwa[13]);
            if ((Val1 >= 300.1) and (sAge <=19)) Then
            begin
               K := K + 1;
               tSokun[K] := 'ALP1';
            end;
        end;
        if  (sGulkwa[13] <> '') and ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
         Begin
            Val1 := StrToFloat(sGulkwa[13]);
            if ((Val1 >= 150.1) and (sAge <=19)) Then
            begin
               K := K + 1;
               tSokun[K] := 'ALP1';
            end;
        end;
        if  (sGulkwa[156] <> '') or (sGulkwa[157] <> '') or
            (sGulkwa[158] <> '') or (sGulkwa[159] <> '') Then
         Begin
            if sGulkwa[156] <> '' then  Val1 := StrToFloat(sGulkwa[156])
            else val1:=0;
            if sGulkwa[157] <> '' then  Val2 := StrToFloat(sGulkwa[157])
            else val2:=0;
            if sGulkwa[158] <> '' then  Val3 := StrToFloat(sGulkwa[158])
            else val3:=0;
            if sGulkwa[159] <> '' then  Val4 := StrToFloat(sGulkwa[159])
            else val4:=0;
            if (((val1>0) and (val1<=134.9)) or (val1>=145.1)) or
               (((val2>0) and (val2<=3.49))  or (val2>=5.51))  or
               (((val3>0) and (val3<=97.9))  or (val3>=110.1)) or
               (((val4>0) and (val4<=21.9))  or (val4>=29.1))  Then
            begin
               K := K + 1;
               tSokun[K] := 'NACL';
            end;
        end;



        if  (sGulkwa[25] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[25]);
            if (Val1 >= 2.351) and  (cSex = 1) then
            begin
               K := K + 1;
               tSokun[K] := 'UCR ';
            end
            else if (Val1 >= 1.571) and  (cSex = 2) then
            begin
               K := K + 1;
               tSokun[K] := 'UCR ';
            end;
        end;
        if  (sGulkwa[160] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[160]);
            if (Val1 >= 30.01) Then
            begin
               K := K + 1;
               tSokun[K] := 'MIAL';
            end;
        end;
        if  (sGulkwa[80] <> '음성') and  (sGulkwa[80] <> '') and
            (((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121201') and ((qryBunju.FieldByName('Dat_Gmgn').AsString)<= '20121209')) Then
        Begin
             K := K + 1;
             tSokun[K] := 'TPHA';
        end;

        if  (sGulkwa[80] <> '') and ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121210')  Then
        Begin
             Val1 := StrToFloat(sGulkwa[80]);
             If Val1>=10 then
                Begin
                K := K + 1;
                tSokun[K] := 'TPHA';
             end;
        end;
        
        If (sGulkwa[67] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[67]);
            if (sGulkwa[53] <> '') then
            Val2 := StrToFloat(sGulkwa[53])
            else
            Val2 := 0 ;

            If (Val1>=0.31) and (Val1 <= 0.99) and
               (((Val2<=10.2) and (cSex  = 1)) or
               ((Val2<=9.8) and (cSex  = 2))) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP1';
               End;
         End;
         If (sGulkwa[67] <> '') and (sGulkwa[53] <> '')  Then
         Begin
            Val1 := StrToFloat(sGulkwa[67]);
            Val2 := StrToFloat(sGulkwa[53]);
            If (Val1>=0.31) and
               (((Val2>10.2) and (cSex  = 1)) or
               ((Val2>9.8) and (cSex  = 2))) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP3';
               End;
         End;
         //20160803 CRP6 ..광주 사업장 문의로 김소영 요청
         If (sGulkwa[67] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[67]);
            if (sGulkwa[53] <> '') then
            Val2 := StrToFloat(sGulkwa[53])
            else
            Val2 := 0 ;

            If (Val1>=0.31) and
               (((Val2<=10.2) and (cSex  = 1)) or
               ((Val2<=9.8) and (cSex  = 2))) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP6';
               End;
         End;
         //crp4 2016.06.30  수지
         If (sGulkwa[67] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[67]);
            if (sGulkwa[53] <> '') then
            Val2 := StrToFloat(sGulkwa[53])
            else
            Val2 := 0 ;

            If (Val1>=1) and (Val1 <= 2.99) and
               (((Val2<=10.2) and (cSex  = 1)) or
               ((Val2<=9.8) and (cSex  = 2))) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP4';
               End;
         End;
         //crp5 2016.06.30  수지
         If (sGulkwa[67] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[67]);
            if (sGulkwa[53] <> '') then
            Val2 := StrToFloat(sGulkwa[53])
            else
            Val2 := 0 ;

            If (Val1>=3) and
               (((Val2<=10.2) and (cSex  = 1)) or
               ((Val2<=9.8) and (cSex  = 2))) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CRP5';
               End;
         End;
         //PT 2017.06.21 수지
         If (sGulkwa[176] <> '') OR (sGulkwa[177] <> '') Then
         Begin
            if (sGulkwa[176] <> '') then
            Val1 := StrToFloat(sGulkwa[176])
            else
            Val2 := 0 ;

            if (sGulkwa[177] <> '') then
            Val2 := StrToFloat(sGulkwa[177])
            else
            Val2 := 0 ;

            If (Val1<11.9) OR (Val1>15.3) OR (Val2<29.6) OR (Val2>45.2) then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PT  ';
               End;
         End;

        //S085          ..20170626 수지
        If (sGulkwa[178] <> '') then
        Begin
             if (sGulkwa[178] <> '') then
             val1 := StrToFloat(sGulkwa[178])
             else val1 := 0;

             if (Val1 > 70) then
             begin
             K := K + 1;
             tSokun[K] := 'MYGH';
             end;
        End;

         //E058
         If (sGulkwa[179] <> '') Then
         Begin
            if (sGulkwa[179] <> '') then
            Val1 := StrToFloat(sGulkwa[179])
            else
            Val2 := 0 ;


            If (Val1>18.4)  then
               Begin
                  K := K + 1;
                  tSokun[K] := 'CORH'
               End
            ELSE IF (Val1<6.02)  then
              Begin
                  K := K + 1;
                  tSokun[K] := 'CORL'
               End;
         End;


        if  (sGulkwa[161] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[161]);
            if (Val1 >= 1.51) Then
            begin
               K := K + 1;
               tSokun[K] := 'TSHR';
            end;
        end;
        if  (sGulkwa[124] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[124]);
            if (Val1 >= 0.161) Then
            begin
               K := K + 1;
               tSokun[K] := 'TRNI';
            end;
        end;
        //
        If (sGulkwa[125] <> '')  or (sGulkwa[126] <> '')Then
         Begin
             VAl1:=0;
             val2:=0;
            if   (sGulkwa[125] <> '') then    Val1 := StrToFloat(sGulkwa[125]);
            if   (sGulkwa[126] <> '') then    Val2 := StrToFloat(sGulkwa[126]);
            If (Val1 >= 34.1) or (Val2 >= 34.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'MIAB';
               End;
        End;

        If (sGulkwa[140] <> '') and  ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130422') Then
         Begin
            Val1 := StrToFloat(sGulkwa[140]);  
            If (Val1 > 30) and  (Val1 <= 100) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'VTDS';
               End;
        End;
        If (sGulkwa[119] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[119]);
            If (Val1 > 25.0) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'NCTS';
               End;
        End;
        If (sGulkwa[24] <> '') and (sGulkwa[37] <> '') and (sGulkwa[44] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[24]);
            Val2 := StrToFloat(sGulkwa[37]);
            Val3 := StrToFloat(sGulkwa[44]);
            If (Val1 >= 126) and (Val2 <= 6.4) and (Val3 <= 10) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'GHHL';
               End;
        End;
        If (sGulkwa[46] <> '') and (sGulkwa[43] <> '') and (sGulkwa[44] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[46]);
            Val2 := StrToFloat(sGulkwa[43]);
            Val3 := StrToFloat(sGulkwa[44]);
            If (Val1 <= 80) and (Val2 >= 6.0) and (Val3 <= 12.9) and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'THAL';
               End
            else If (Val1 <= 75) and (Val2 >= 5.5) and (Val3 <= 11.9) and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'THAL';
               End;
        End;

        If (sGulkwa[163] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[163]);
            If (Val1 >= 60)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'GFR1';
               End
            else If (Val1 >= 30) and (Val1 <= 59) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'GFR3';
               End
           else If (Val1 <= 29) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'GFR5';
               End;
        End;
        //E056 소견
        If (sGulkwa[149] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString) < '20140609') Then
         Begin
            Val1 := StrToFloat(sGulkwa[149]);
            If ((Val1 >= 2.8) and (Val1 <= 8.0) and (cSex = 1)) or
               ((Val1 >= 0.06) and (Val1 <= 0.82) and (cSex = 2)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'TTT ';
               End
            else  If ((Val1 >= 8.01)  and (cSex = 1)) or
                     ((Val1 >= 0.821)  and (cSex = 2)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'TTTH';
               End
            else  If ((Val1 <= 2.79)  and (cSex = 1)) or
                     ((Val1 <= 0.059)  and (cSex = 2)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'TTTL';
               End
        End;
        If (sGulkwa[149] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString) >= '20140609') Then
         Begin
            Val1 := StrToFloat(sGulkwa[149]);
            //남자-50세미만
            If (cSex = 1) And (sAge < 50) Then
               Begin
                  If (Val1 > 8.36) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTH';
                     End
                  Else If (Val1 >= 2.49) And (Val1 <= 8.36) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTT ';
                     End
                  Else If (Val1 < 2.49) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTL';
                     End;
               End;
            //남자-50세이상
            If (cSex = 1) And (sAge >= 50) Then
               Begin
                  If (Val1 > 7.40) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTH';
                     End
                  Else If (Val1 >= 1.93) And (Val1 <= 7.40) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTT ';
                     End
                  Else If (Val1 < 1.93) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTL';
                     End;
               End;
            //여자-50세미만
            If (cSex = 2) And (sAge < 50) Then
               Begin
                  If (Val1 > 0.48) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTH';
                     End
                  Else If (Val1 >= 0.08) And (Val1 <= 0.48) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTT ';
                     End
                  Else If (Val1 < 0.08) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTL';
                     End;
               End;
            //여자-50세이상
            If (cSex = 2) And (sAge >= 50) Then
               Begin
                  If (Val1 > 0.41) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTH';
                     End
                  Else If (Val1 >= 0.03) And (Val1 <= 0.41) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTT ';
                     End
                  Else If (Val1 < 0.03) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'TTTL';
                     End;
               End;
        End;
        //E049 소견 2014-06-10 추가
        If (sGulkwa[170] <> '') And ((qryBunju.FieldByName('Dat_bunju').AsString) >= '20140609') Then
         Begin
            Val1 := StrToFloat(sGulkwa[170]);
            //남자-15세미만
            If (cSex = 1) And (sAge < 15) Then
               Begin
                  If (Val1 > 1.80) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 0.13) And (Val1 <= 1.80) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 0.13) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //남자-20세~39세
            If (cSex = 1) And (sAge >= 15) And (sAge <= 39) Then
               Begin
                  If (Val1 > 40.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 5.40) And (Val1 <= 40.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 5.40) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //남자-40세~59세
            If (cSex = 1) And (sAge >= 40) And (sAge <= 59) Then
               Begin
                  If (Val1 > 25.70) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 3.60) And (Val1 <= 25.70) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 3.60) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //남자-60세이상
            If (cSex = 1) And (sAge >= 60) Then
               Begin
                  If (Val1 > 28.80) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 1.50) And (Val1 <= 28.80) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 1.50) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //여자-15세이하
            If (cSex = 2) And (sAge < 15) Then
               Begin
                  If (Val1 > 2.70) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 0.13) And (Val1 <= 2.70) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 0.13) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //여자-20세~39세
            If (cSex = 2) And (sAge >= 15) And (sAge <= 39) Then
               Begin
                  If (Val1 > 4.60) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 0.13) And (Val1 <= 4.60) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 0.13) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //여자-40세~59세
            If (cSex = 2) And (sAge >= 40) And (sAge <= 59) Then
               Begin
                  If (Val1 > 4.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 0.13) And (Val1 <= 4.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 0.13) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
            //여자-60세이상
            If (cSex = 2) And (sAge >= 60) Then
               Begin
                  If (Val1 > 5.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTH';
                     End
                  Else If (Val1 >= 0.13) And (Val1 <= 5.00) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTN';
                     End
                  Else If (Val1 < 0.13) Then
                     Begin
                        K := K + 1;
                        tSokun[K] := 'FTTL';
                     End;
               End;
        End;


        If (sGulkwa[44] <> '') and (sGulkwa[165] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[44]);
            Val2 := StrToFloat(sGulkwa[165]);
            If (cSex = 1) and (Val1 >= 13.1) and (Val1 <= 18.0) and (Val2>=2.51) and (Val2 <= 4.0)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RET ';
               End
            else  If (cSex = 2) and (Val1 >= 12.1) and (Val1 <= 16.0) and (Val2>=2.51) and (Val2 <= 4.0)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RET ';
               End
            else  If (cSex = 1) and (Val1 <= 12.9) and (Val2 >= 2.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RET3';
               End
            else  If (cSex = 2) and (Val1 <= 11.9) and (Val2 >= 2.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PET3';
               End
            else  If (cSex = 1) and (Val1 >= 13.1) and (Val1 <= 18.0)  and (Val2 >= 4.1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RET4';
               End
             else  If (cSex = 2) and (Val1 >= 12.1) and (Val1 <= 16.0)  and (Val2 >= 4.1)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PET4';
               End
            else  If (cSex = 1) and (Val1 >= 18.1)  and (Val2 >= 2.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RET4';
               End
            else  If (cSex = 2) and (Val1 >= 16.1)  and (Val2 >= 2.51)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PET4';
               End
        End;
      If (sGulkwa[09]  <> '') and (sGulkwa[10]  <> '') and (sGulkwa[113] = '') Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
          if  (sGulkwa[168]<> '')  Then  Val4 := StrToFloat(sGulkwa[168])
          else  Val4 :=0;
         If (Val2 <=45) and (Val1 <=40)  and (cSex = 1) and  (val4>=20)  Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
         else If (Val2 <=34) and (Val1 <=32)  and (cSex = 2) and  (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
         else If (Val2 <=45) and (Val1 >=40.1) and (Val3 <=70) and (cSex = 1) and   (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
          else If (Val2 <=34) and (Val1 >=32.1) and (Val3 <=42) and (cSex = 2) and   (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End;
      End;
      /////2017.07.25 수지
      IF (sGulkwa[180] <> '') Then
      begin
          if  (sGulkwa[180] <> '') then  Val1 := StrToFloat(sGulkwa[180])
          else  Val1 :=0;
          if (sAge <= 39) and  (Val1 > 60.5) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4 ';
          end
          Else if ((sAge >= 40) and (sAge <= 49)) and (Val1 > 76.2) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4 ';
          end
          Else if ((sAge >= 50) and (sAge <= 59)) and (Val1 > 74.3) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4 ';
          end
          Else if ((sAge >= 60) and (sAge <=69)) and (Val1 > 82.9) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4 ';
          end
          Else if (sAge >= 70) and (Val1 > 104) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4 ';
          end;
      end;
      IF (sGulkwa[180] <> '') Then
      begin
          if  (sGulkwa[180] <> '') then  Val1 := StrToFloat(sGulkwa[180])
          else  Val1 :=0;
          if (sAge <= 39) and  (Val1 <= 60.5) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4N';
          end
          Else if ((sAge >= 40) and (sAge <= 49)) and (Val1 <= 76.2) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4N';
          end
          Else if ((sAge >= 50) and (sAge <= 59)) and (Val1 <= 74.3) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4N';
          end
          Else if ((sAge >= 60) and (sAge <= 69)) and (Val1 <= 82.9) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4N';
          end
          Else if (sAge >= 70) and (Val1 <= 104) then
          begin
            K := K + 1;
            tSokun[K] := 'HE4N';
          end;
      end;
      if   (sGulkwa[122] <> '')  Then
      begin
           if  (sGulkwa[122]<> '')  Then  Val1 := StrToFloat(sGulkwa[122])
           else  Val1 :=0;

           If (Val1 > 1.32)  Then
           Begin
               K := K + 1;
               tSokun[K] := 'ICAH';
           End
           Else if (Val1 < 1.0)  Then
           Begin
               K := K + 1;
               tSokun[K] := 'ICAL';
           End;
      END;
      if   (sGulkwa[181] <> '')  Then
      begin
           if  (sGulkwa[181]<> '')  Then  Val1 := StrToFloat(sGulkwa[181])
           else  Val1 :=0;

           If (Val1 > 19)  Then
           Begin
               K := K + 1;
               tSokun[K] := 'ADAH';
           End
           Else if (Val1 < 8)  Then
           Begin
               K := K + 1;
               tSokun[K] := 'ADAL';
           End;
      END;
      //////////////////////////
      if   (sGulkwa[168] <> '')  Then
      begin
           if  (sGulkwa[168]<> '')  Then  Val1 := StrToFloat(sGulkwa[168])
           else  Val1 :=0;

           If (Val1>=20) and (sGulkwa[113] = '양성')   Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV4';
           End
           else If (Val1>=20) and (sGulkwa[113] = '음성')    Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV ';
           End
           else If (Val1< 20) and ((sGulkwa[113] = '음성') or (sGulkwa[113] = ''))    Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
           End
      end;

        If (sGulkwa[165] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[165]);
            If (Val1 <= 0.49)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PET1';
               End;
        End;
        If (sGulkwa[167] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[167]);
            If (Val1 >= 0.9) and (Val1 <= 1.09) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HPYW';
               End
            else  If (Val1 >= 1.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HPYP';
               End; 
        End;


        If (sGulkwa[152] <> '') Then
         Begin
            If (sGulkwa[152] = '양성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FCP ';
               End
            else  If (sGulkwa[152] = '음성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FCN ';
               End;
        End;
      end
      else if  qryBunju.FieldByName('Num_Bunju').asInteger>= 7000 then
      begin
          If (sGulkwa[152] <> '') Then
          Begin
            If (sGulkwa[152] = '양성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FCP ';
               End
            else  If (sGulkwa[152] = '음성') then
               Begin
                  K := K + 1;
                  tSokun[K] := 'FCN ';
               End;
        End;

      end;



End;

procedure TfrmLD6I01.btnStartClick(Sender: TObject);
Var
   i, sCnt, tAge : Integer;
   Val1, Val2 : Extended;
   cValue, sTemp, sDate : String;
   dCnt1, dCnt2, dCnt3 : Integer;
   dStr1, dStr2, dStr3,sokun_temp : String;


   sPos, ePos, Size1, Size2, Size3,iAge,cnt : Integer;
   Low1, Low2, Low3, Low4, Low5, Low6, Low7 : Single;
   High1, High2, High3, High4, High5, High6, High7 : Single;
   Low1f, Low2f, Low3f, Low4f, Low5f, Low6f, Low7f : Single;
   High1f, High2f, High3f, High4f, High5f, High6f, High7f : Single;
   vTemp : Single;
   Jung,sMan, Ok_Start : String;
   GfGulkwa, GfGulkwa1, GfGulkwa2, GfGulkwa3 : String;
   sSkip, Wk_Sw1 : Integer;
   check_tk : Boolean;
   vCod_profile : Variant;
begin
   inherited;
   Animate1.Active := False;
   Animate2.Active := False;
   ProBar.Progress := 0;

// 일자가 선택되었는지 Check
   if mskDate.Text  = '' then
   begin
      GF_MsgBox('Warning', '분주일자는 반드시 입력해야 합니다.');
      mskDate.SetFocus;
      exit;
   end;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   //지사일 경우 7000번대 이상으로는 돌아가지 못하게함. [7000번대 이상인 분주에 대해 본원에서 처리X]
   if (GV_sJisa <> '15') then
   begin
      if (mskFrom.Text = '') or (copy(mskFrom.Text,1,1) = '0') or
         (copy(mskFrom.Text,1,1) = '1') or (copy(mskFrom.Text,1,1) = '2') or (copy(mskFrom.Text,1,1) = '3') or
         (copy(mskFrom.Text,1,1) = '4') or (copy(mskFrom.Text,1,1) = '5') or (copy(mskFrom.Text,1,1) = '6') then
      begin
         GF_MsgBox('Warning', '7000번 이하는 소견작업을 할 수 없습니다.');
         mskDate.SetFocus;
         exit;
      end;
   end;

   // 전체(00)일경우 본사를 기준으로 작업
   if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
   else                    UV_sJisa := GV_sJisa;

   //혈액결과 완료 처리 무시 ..응급으로 소견 나올 경우 _박연숙, 김소영
   if Ck_hulgulkwa.Checked = True then
   begin
   end
   else if GV_sJisa = '51' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_51').AsString = 'N') or
                 (FieldBYName('chk_02_51').AsString = 'N') or
                 (FieldBYName('chk_03_51').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   else if GV_sJisa = '11' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_11').AsString = 'N') or
                 (FieldBYName('chk_02_11').AsString = 'N') or
                 (FieldBYName('chk_03_11').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   else if GV_sJisa = '12' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_12').AsString = 'N') or
                 (FieldBYName('chk_02_12').AsString = 'N') or
                 (FieldBYName('chk_03_12').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   else if GV_sJisa = '43' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_43').AsString = 'N') or
                 (FieldBYName('chk_02_43').AsString = 'N') or
                 (FieldBYName('chk_03_43').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   else if GV_sJisa = '61' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_61').AsString = 'N') or
                 (FieldBYName('chk_02_61').AsString = 'N') or
                 (FieldBYName('chk_03_61').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   else if GV_sJisa = '71' then
   Begin
   with qryHgulkwa_EtcChk do
   begin
      close;
      ParamByName('cod_bjjs').AsString   := '15';
      ParamByName('dat_bunju').AsString := mskDate.text;
      open;
      begin
         GP_query_log(qryHgulkwa_EtcChk, 'LD6I01', 'Q', 'Y');
         begin
              IF (FieldBYName('chk_01_71').AsString = 'N') or
                 (FieldBYName('chk_02_71').AsString = 'N') or
                 (FieldBYName('chk_03_71').AsString = 'N') then
                 begin
                    GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 + '혈액파트 검사중입니다.');
                    exit;
                 end;
         end;
      end;
   end;
   end
   // 각 혈액파트 완료체크...
   else if GF_HulGulkwaCheck(1, mskDate.Text, UV_sJisa, UV_sHul_Part) then
   begin
      GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 +
                '다음 혈액파트[' + UV_sHul_Part + ']가 검사중입니다.');
      mskDate.SetFocus;
      exit;
   end; // end of if[GF_HulGulkwaCheck];


   // 단체확인작업...
   if (chkCod_Dc.Checked) then
   begin
      If (Trim(edtDc.Text) = '') Or (Trim(edtDesc_Dc.Text) = '') Then
         Begin
            GF_MsgBox('Warning', '혈액 자동소견을 작업할 수 없습니다.' + #13#13 +
                      '☞  단체를 확인 하여 주십시요');
            mskDate.SetFocus;
            exit;
         End;
   end; // end of if 단체확인


   if not GF_MsgBox('Confirmation', '자동소견을 시작하시겠습니까 ?') then exit;

// 결과 Check항목(115종)

   Sk_Value := 'C001C002C003C004C005C006C007S010C009C010' +  //   1 ~  10
               'C011C012C013C014C017C019C021C025C026C027' +  //  11 ~  20
               'C028C029C030C032U028C034C035C039C041C042' +  //  21 ~  30
               'C043C044C045C047C048C049C083C020U033E001' +  //  31 ~  40
               'E002E003H001H002H003H004H005H006H007H008' +  //  41 ~  50
               'H009H010H011H012S099H014H015H016H017H018' +  //  51 ~  60
               'H019H020H021H022H023H024S001S002S003S007' +  //  61 ~  70
               'S008S016S018S019S020C090S030S031S034S036' +  //  71 ~  80
               'T001T002T006T007T008T009T011TT01U001U002' +  //  81 ~  90
               'U003U004U005U006U007U008U009U010Y004SE07' +  //  91 ~ 100
               'SE08SE09SE10E040T010T012E016SE13C015C057' +  // 101 ~ 110
               'S014S015SE01SE02S091C078H025TT02T039C065' +  // 111 ~ 120
               'C056C054T041C072E005E019C023C022C080E015' +  // 121~  130
               'T042S011S012E066E067S071C058C069T005T043' +  // 131~  140
               'C082T037C081E051E052E053E054E055E056E057' +  // 141~  150
               'T031U017S064S056S075C050C051C052C053U023' +  // 151~  160
               'E012E072C087E069H029E021U058SE31C070E049' +  // 161~  170
               'TT03C038S004T028T030H027H028S085E058S103' +  // 171~  180
               'C016Y010';                                       // 181~
   sDate := DateToStr(Date);
// 분주Query 시작
   With qryBunju Do
      Begin
         Close;
         ParamByName('In_Cod_bjjs').AsString := UV_sJisa;
         ParamByName('In_Dat_Bunju').AsString := mskDate.Text;
         If mskFrom.Text = '' Then
            ParamByName('St_Num_Bunju').AsString := '0000'
         Else
            ParamByName('St_Num_Bunju').AsString := mskFrom.Text;
         If mskTo.Text = '' Then
            Begin
               If ((Trim(edtDc.Text) <> '') And (chkCod_Dc.Checked)) or (GV_sJisa = '51') Then
                  ParamByName('Ed_Num_Bunju').AsString := '9999'
               Else
                  ParamByName('Ed_Num_Bunju').AsString := '6999';
            End
         Else
            ParamByName('Ed_Num_Bunju').AsString := mskTo.Text;
         Open;
         sCnt := 0;
         dStr1 := '';
         dStr2 := '';
         dStr3 := '';
         cnt :=0;
         Probar.MaxValue := RecordCount;
         Animate1.Active := True;
         Animate2.Active := True;
         GP_query_log(qryBunju, 'LD6I01', 'Q', 'Y');
// 검진Query 시작
         While Not Eof Do
            Begin
               With qryGumgin Do
                  Begin
                     Close;
                     ParamByName('In_Cod_Jisa').Asstring := qryBunju.FieldByName('Cod_jisa').AsString;
                     ParamByName('In_Num_Id').AsString   := qryBunju.FieldByName('Num_id').Asstring;
                     ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_Gmgn').AsString;
                     Open;
                     Ok_Start := 'Ok';
                     If (Trim(edtDc.Text) <> '') And (chkCod_Dc.Checked) Then
                        Begin
                          If qryGumgin.FieldByName('Cod_Dc').AsString = Trim(edtDc.Text) Then
                             Ok_Start := 'Ok'
                         Else
                             Ok_Start := 'No';
                      End;

// Profile_hm Query시작
                     cValue := '';
                     aSkip := 0;
                     LabStatus.Caption := qryBunju.FieldByName('Num_Bunju').Asstring;

                     For dCnt1 := 0 To 182 Do
                        Begin
                            tSokun[dCnt1] := '';
                            sSokun[dCnt1] := '';
                            uSokun[dCnt1] := '';
                            sCheck[dCnt1] := '';
                            tSokun_chuga[dCnt1] := '';
                        End;
                     For dCnt1 := 0 To 182 Do
                        sGulkwa[dCnt1] := '';
                     K := 0;
                     U := 0;
                     S := 0;
                     C := 0;
                     sSkip := 0;
                     If Ok_Start = 'Ok' Then
                     With qryProfileHm Do
//장비코드
                        Begin
                           Close;
                           ParamByName('In_Cod_pf').AsString := qryGumgin.FieldByname('Cod_jangbi').AsString;
                           Open;
                           While Not Eof Do
                              Begin
                                 If Copy(FieldByName('Cod_hm').AsString,1,2) <> 'JJ' Then
                                    cValue := cValue + FieldByName('Cod_hm').AsString;
                                 Next;
                              End;
                           Close;
//혈액코드
                           Close;
                           ParamByName('In_Cod_pf').AsString := qryGumgin.FieldByname('Cod_hul').AsString;
                           Open;
                           While Not Eof Do
                              Begin
                                 If Copy(FieldByName('Cod_hm').AsString,1,2) <> 'JJ' Then
                                    cValue := cValue + FieldByName('Cod_hm').AsString;
                                 Next;
                              End;
                           Close;
//Cancer코드
                           Close;
                           ParamByName('In_Cod_pf').AsString := qryGumgin.FieldByname('Cod_Cancer').AsString;
                           Open;
                           While Not Eof Do
                              Begin
                                 If Copy(FieldByName('Cod_hm').AsString,1,2) <> 'JJ' Then
                                    cValue := cValue + FieldByName('Cod_hm').AsString;
                                 Next;
                              End;
                           Close;
                        End;

                     //---- 특검항목 추출...
                     dCnt1 := GF_MulToSingle(FieldByName('cod_prf').AsString, 4, vCod_profile);
                     If Ok_Start = 'Ok' Then
                     for dCnt2 := 0 to dCnt1 - 1 do
                     begin
                        if Trim(vCod_profile[i]) <> '' then
                        begin
                           qryPf_hm2.Active := False;
                           qryPf_hm2.ParamByName('COD_PF').AsString := vCod_profile[i];
                           qryPf_hm2.Active := True;
                           if qryPf_hm2.RecordCount > 0 then
                           begin
                              while not qryPf_hm2.Eof do
                              begin
                                 check_tk := True;
                                 for dCnt3 := 1 to round(Length(cValue)/4) do
                                 begin
                                    if copy(cValue, (dCnt3 * 4) - 3,4) =
                                       qryPf_hm2.FieldByName('COD_HM').AsString then check_tk := False;
                                 end;
                                 if (check_tk = True) AND (Copy(qryPf_hm2.FieldByName('COD_HM').AsString,1,2) <> 'JJ') then
                                    cValue := cValue + qryPf_hm2.FieldByName('COD_HM').AsString;
                                 qryPf_hm2.Next;
                              end; // end of while(항목 처리)
                           end; // end of if
                        end; //end of if
                     end; //end of for
                     dCnt1 := Length(FieldByName('TkGum_chuga').AsString);
                     If Ok_Start = 'Ok' Then
                     For dCnt2 := 0 to dCnt1 - 1 Do
                        If Copy(FieldByName('TkGum_chuga').AsString,dCnt2 * 4 + 1,2) <> 'JJ' Then
                           cValue := cValue + Copy(FieldByName('TkGum_chuga').AsString,dCnt2 * 4 + 1,4);

                     // 노신유형Display
                     if FieldByName('gubn_nosin').AsString = '1' then
                        cValue := cValue + UF_No_Hangmok(copy(GV_sToday,1,4), '1', FieldByName('gubn_nsyh').AsInteger);
                     //성인병유형Display
                     if FieldByName('gubn_adult').AsString = '1' then
                        cValue := cValue + UF_No_Hangmok(copy(GV_sToday,1,4), '4', FieldByName('gubn_adyh').AsInteger);
                     //간염유형Display
                     if FieldByName('gubn_agab').AsString = '1' then
                        cValue := cValue + UF_No_Hangmok(copy(GV_sToday,1,4), '5', FieldByName('gubn_agyh').AsInteger);
                     //생애전환기유형Display
                     if FieldByName('gubn_life').AsString = '1' then
                        cValue := cValue + UF_No_Hangmok(copy(GV_sToday,1,4), '7', FieldByName('gubn_lfyh').AsInteger);
                     //추가코드
                     dCnt1 := Length(FieldByName('Cod_chuga').AsString);
                     If Ok_Start = 'Ok' Then
                     For dCnt2 := 0 to dCnt1 - 1 Do
                        If Copy(FieldByName('Cod_Chuga').AsString,dCnt2 * 4 + 1,2) <> 'JJ' Then
                           cValue := cValue + Copy(FieldByName('Cod_Chuga').AsString,dCnt2 * 4 + 1,4);

                     dStr1 := '';
                     dCnt1 := Round(Length(cValue) / 4);
                     If Ok_Start = 'Ok' Then
                     For dCnt3 := 0 To 182 Do
                        For dCnt2 := 0 To dCnt1 - 1 Do
                            If Copy(Sk_Value,dCnt3 * 4 + 1,4) = Copy(cValue,dCnt2 * 4 + 1, 4) Then
                               dStr1 := dStr1 + Copy(cValue,dCnt2 * 4 + 1, 4);
                     cValue := dStr1;
//항목의 임상자료 Query
                     dCnt1 := Round(Length(cValue) / 4);
                     If Ok_Start = 'Ok' Then
                     For dCnt2 := 0 To dCnt1 - 1 Do
                        With qryHangmok Do
                           Begin
                              Close;
                              ParamByName('In_Cod_hm').AsString := Copy(cValue,dCnt2 * 4 + 1,4);
                              Open;
                              If FieldByName('Gubn_hm').AsString = '2' Then
                                 Jung := FieldByname('Ims_Munja').AsString
                              Else
                                 Begin
                                    Low1 := FieldByName('Ims_low1').AsFloat;
                                    Low2 := FieldByName('Ims_low2').AsFloat;
                                    Low3 := FieldByName('Ims_low3').AsFloat;
                                    Low4 := FieldByName('Ims_low4').AsFloat;
                                    Low5 := FieldByName('Ims_low5').AsFloat;
                                    Low6 := FieldByName('Ims_low6').AsFloat;
                                    Low7 := FieldByName('Ims_low7').AsFloat;
                                    High1 := FieldByName('Ims_High1').AsFloat;
                                    High2 := FieldByName('Ims_High2').AsFloat;
                                    High3 := FieldByName('Ims_High3').AsFloat;
                                    High4 := FieldByName('Ims_High4').AsFloat;
                                    High5 := FieldByName('Ims_High5').AsFloat;
                                    High6 := FieldByName('Ims_High6').AsFloat;
                                    High7 := FieldByName('Ims_High7').AsFloat;
                                    Low1f := FieldByName('Ims_low1f').AsFloat;
                                    Low2f := FieldByName('Ims_low2f').AsFloat;
                                    Low3f := FieldByName('Ims_low3f').AsFloat;
                                    Low4f := FieldByName('Ims_low4f').AsFloat;
                                    Low5f := FieldByName('Ims_low5f').AsFloat;
                                    Low6f := FieldByName('Ims_low6f').AsFloat;
                                    Low7f := FieldByName('Ims_low7f').AsFloat;
                                    High1f := FieldByName('Ims_High1f').AsFloat;
                                    High2f := FieldByName('Ims_High2f').AsFloat;
                                    High3f := FieldByName('Ims_High3f').AsFloat;
                                    High4f := FieldByName('Ims_High4f').AsFloat;
                                    High5f := FieldByName('Ims_High5f').AsFloat;
                                    High6f := FieldByName('Ims_High6f').AsFloat;
                                    High7f := FieldByName('Ims_High7f').AsFloat;
                                 End;
                              sPos := FieldByName('Pos_Start').AsInteger;
                              ePos := FieldByname('Pos_End').AsInteger;
// 혈액결과찾기
                              With qryGlkwa Do
                                 Begin
                                    Close;
                                    ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
                                    ParamByName('In_Num_id').Asstring := qryBunju.FieldByname('Num_id').AsString;
                                    ParamByName('In_Dat_Bunju').Asstring := qryBunju.FieldByname('Dat_Bunju').AsString;
                                    ParamByName('In_Gubn_Part').AsString := qryHangmok.FieldByname('Gubn_Part').AsString;
                                    ParamByname('In_Dat_Gmgn').Asstring := qryBunju.FieldByName('Dat_Gmgn').AsString;
                                    Open;
                                    GfGulkwa1 := Copy(FieldByName('DESC_GLKWA1').AsString,1,200);
                                    GfGulkwa2 := Copy(FieldByName('DESC_GLKWA2').AsString,1,200);
                                    GfGulkwa3 := Copy(FieldByName('DESC_GLKWA3').AsString,1,200);
                                    GF_HulGulkwa('2', GfGulkwa1, GfGulkwa2, GfGulkwa3, GfGulkwa);
                                    sTemp := Trim(Copy(GfGulkwa,sPos,6));
// 분석항목 Sort
                                    cSex := StrToInt(Copy(qryGumgin.FieldByName('Num_Jumin').AsString,7,1));
//남여Check
                                    If (cSex = 1) Or (cSex = 3) Or (cSex = 5) Or (cSex = 7) Then
                                       cSex := 1
                                    else if (cSex = 2) Or (cSex = 4) Or (cSex = 6) Or (cSex = 8) Then
                                       cSex := 2;
                                    Gpreg :=qryGumgin.FieldByName('Gubn_preg').AsString;


                                    dCnt3 := 0;
                                    While dCnt3 <= 182 Do
                                       Begin
//정상,이상자 항목Check
                                          If Copy(Sk_Value,dCnt3 * 4 + 1,4) = Copy(cValue,dCnt2 * 4 + 1,4) Then
                                             Begin
                                                sGulkwa[dCnt3 + 1] := sTemp;
                                               { If Copy(qryGumgin.FieldByName('Num_jumin').AsString,1,2) < '10' Then
                                                   tAge := StrToInt('20' + Copy(qryGumgin.FieldByName('Num_jumin').AsString,1,2))
                                                Else
                                                   tAge := StrToInt('19' + Copy(qryGumgin.FieldByName('Num_jumin').AsString,1,2));}


                                                GP_JuminToAge(qryGumgin.FieldByname('num_jumin').AsString,qryGumgin.FieldByname('Dat_gmgn').AsString, sMan, iAge);
                                                sAge := iAge;

                                                If (sTemp = '') Then
                                                   If (Copy(cValue,dCnt2 * 4 + 1,4) >= 'U001') And
                                                      (Copy(cValue,dCnt2 * 4 + 1,4) <= 'U010') Then
                                                      If qryHangmok.FieldByName('Gubn_hm').AsString <> '2' Then
                                                         sTemp := FloatToStr(Low3)
                                                      Else
                                                         sTemp := Jung;


                                                If qryHangmok.FieldByName('Gubn_hm').AsString <> '2' Then
                                                begin
                                                   If (sTemp = '') Or (sTemp = ' ') or (pos(':',sTemp)>0) Then
                                                      begin
                                                      sSkip := 99999;
                                                      end
                                                   Else                                  vTemp := StrToFloat(sTemp);
//[2012.03.22] H018~H022 결과값 없어서 제외~
                                                   if (Copy(cValue,dCnt2 * 4 + 1,4) >= 'H018') And (Copy(cValue,dCnt2 * 4 + 1,4) <= 'H022') Then sSkip := 0;
                                                end;
                                                If (Pos('보류',sGulkwa[70]) > 0) Then
                                                   sSkip := 99999;
                                                If (Pos('보류',sGulkwa[dCnt3 + 1]) > 0) Then
                                                   sSkip := 99999;
                                                //[20060407-유동구]검사항목 결과 없을 때 보류..
                                                If (Trim(sGulkwa[82])  = '') and (pos('T002',cValue) > 0) and ((dCnt3 + 1) = 82) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[83])  = '') and (pos('T006',cValue) > 0) and ((dCnt3 + 1) = 83) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[84])  = '') and (pos('T007',cValue) > 0) and ((dCnt3 + 1) = 84) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[88])  = '') and (pos('TT01',cValue) > 0) and ((dCnt3 + 1) = 88) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[106]) = '') and (pos('T012',cValue) > 0) and ((dCnt3 + 1) = 106) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[107]) = '') and (pos('E016',cValue) > 0) and ((dCnt3 + 1) = 107) then
                                                   sSkip := 99999;
                                                If (Trim(sGulkwa[108]) = '') and (pos('SE13',cValue) > 0) and ((dCnt3 + 1) = 108) then
                                                   sSkip := 99999;

                                                If qryHangmok.FieldByName('Gubn_hm').AsString <> '2' Then
                                                   Case cSex Of
                                                   1 : Begin
                                                         If sAge > 59 Then
                                                            Begin
                                                              If (vTemp < Low7) or (vTemp > High7) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 49 Then
                                                            Begin
                                                              If (vTemp < Low6) or (vTemp > High6) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 39 Then
                                                            Begin
                                                              If (vTemp < Low5) or (vTemp > High5) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 29 Then
                                                            Begin
                                                              If (vTemp < Low4) or (vTemp > High4) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 19 Then
                                                            Begin
                                                              If (vTemp < Low3) or (vTemp > High3) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 9 Then
                                                            Begin
                                                              If (vTemp < Low2) or (vTemp > High2) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else
                                                            Begin
                                                              If (vTemp < Low1) or (vTemp > High1) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End;
                                                       End;
                                                   2 : Begin
                                                         If sAge > 59 Then
                                                            Begin
                                                              If (vTemp < Low7f) or (vTemp > High7f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 49 Then
                                                            Begin
                                                              If (vTemp < Low6f) or (vTemp > High6f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 39 Then
                                                            Begin
                                                              If (vTemp < Low5f) or (vTemp > High5f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 29 Then
                                                            Begin
                                                              If (vTemp < Low4f) or (vTemp > High4f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 19 Then
                                                            Begin
                                                              If (vTemp < Low3f) or (vTemp > High3f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else If sAge > 9 Then
                                                            Begin
                                                              If (vTemp < Low2f) or (vTemp > High2f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End
                                                          Else
                                                            Begin
                                                              If (vTemp < Low1f) or (vTemp > High1f) Then
                                                                  sCheck[dCnt3 + 1] := 'C'
                                                              Else
                                                                  sCheck[dCnt3 + 1] := ' ';
                                                            End;
                                                         End;
                                                  End
                                                Else
                                                  If Jung = '' Then
                                                     sCheck[dCnt3 + 1] := ' '
                                                  Else If sTemp <> Jung Then
                                                          sCheck[dCnt3 + 1] := 'C'
                                                       Else
                                                          sCheck[dCnt3 + 1] := ' ';
                                                dCnt3 := 182;/////   카운터
                                             End;
                                          dCnt3 := dCnt3 + 1;
                                       End;
                                    Close;
                                 End;
                           End;
//소견처리 Routine
                     If Ok_Start = 'Ok' Then
                     If ((sSkip <> 99999) And (Chk_Sex.Checked = False)) Or (Chk_Sex.Checked = True) Then
                        Begin
                           If sAge < 14 Then
                              Child_sokun_Rtn
                           Else
                              Sokun_Rtn;
                              Danche_Sokun_Rtn;
                              Chuga_Sokun_Rtn;
//C011이 없는 소견 처리(시립대,힐튼호텔) 2014.04.23
                           If Trim(sGulkwa[11]) = '' Then
                               NOT_C011_Sokun_Rtn;
//거래처별 소견 처리(시립대,힐튼호텔) 2014.04.23
                           If (Trim(edtDc.Text) <> '') And (chkCod_Dc.Checked) Then
                               Danche_Sokun_Rtn;
// 마지막소견처리
                           aSkip := 0;
                           Wk_Sw1 := 0;
                           For dCnt1 := 1 To 14 Do
                              If tSokun[dCnt1] <> '' Then
                                 aSkip := 99;
                           If (sGulkwa[71] = '') and (sGulkwa[115] = '') Then
                              If sGulkwa[70] = '음성' Then
                                 Begin
                                    sGulkwa[71] := '양성';
                                    Wk_sw1 := 99;
                                 End
                              Else
                                 If sGulkwa[70] = '양성' Then
                                    Begin
                                       sGulkwa[71] := '음성';
                                       wk_sw1 := 99;
                                    End;
//혈액소견이 없을때
                           If aSkip = 0 Then
                              Begin
//모든대상자포함 2014.04.24
//                              if qryBunju.FieldByname('Num_bunju').AsInteger < 7000 then
                              //C011 결과 있을때
                              if Trim(sGulkwa[11]) <> '' then
                              begin
                                 if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
                                 else val1 :=0;
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] <> '' )   then
                                     begin
                                         If (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (strTofloat(sGulkwa[011]) <= 70) And (cSex = 1) and
                                        ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                        ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))      then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH1 ';
                                     End
                                     else If (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                             (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2) and
                                             ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                             ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))    then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH1 ';
                                     End;
                                 end;
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     If (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (strTofloat(sGulkwa[011]) <= 70) And (cSex = 1)   and
                                        ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                        (sGulkwa[71] = '음성')  and
                                         ((sGulkwa[008]='음성') or (sGulkwa[008]=''))        then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH1 ';
                                     End
                                     else If (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                             (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2)     and
                                             ((sGulkwa[70] = '음성')  or ((val1>0) and (val1<1.0)))   and
                                             (sGulkwa[071] = '음성')   and
                                             ((sGulkwa[008]='음성') or (sGulkwa[008]=''))     then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH1 ';
                                     End;
                                 end;
                                 ////   NH2
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] <> '' ) then
                                     begin
                                     If ((strTofloat(sGulkwa[009]) >= 40.1) or (strTofloat(sGulkwa[010]) >= 45.1) or
                                        (strTofloat(sGulkwa[011]) >= 70.1)  or (sGulkwa[70] = '양성') or  (sGulkwa[70] = '약양성') or  (val1>=1.0) or
                                        ((sGulkwa[71] = '양성') or (sGulkwa[071] = '약양성')or (strTofloat(sGulkwa[115]) >=10)) or (sGulkwa[008] ='양성') ) and (cSex = 1) then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH2 ';
                                     End
                                     else If ((strTofloat(sGulkwa[009]) >= 32.1) or (strTofloat(sGulkwa[010]) >= 34.1) or
                                             (strTofloat(sGulkwa[011]) >= 42.1) or (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or  (val1>=1.0) or
                                             ((sGulkwa[71] = '양성') or (sGulkwa[071] = '약양성') or (strTofloat(sGulkwa[115]) >=10) ) or
                                             (sGulkwa[008] ='양성'))and  (cSex = 2)  then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH2 ';
                                     End;
                                 end;

                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     if ((strTofloat(sGulkwa[009]) >= 40.1) or (strTofloat(sGulkwa[010]) >= 45.1) or
                                        (strTofloat(sGulkwa[011]) >= 70.1) or (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or
                                        ((sGulkwa[70] = '약양성') or  (sGulkwa[70] = '양성')or (val1>=1.0))or (sGulkwa[008] ='양성'))  and
                                        (cSex = 1) then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH2 ';
                                     End
                                     else If ((strTofloat(sGulkwa[009]) >= 32.1) or (strTofloat(sGulkwa[010]) >= 34.1) and
                                             (strTofloat(sGulkwa[011]) >= 42.1) or (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or
                                             ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0)) or (sGulkwa[008] ='양성'))
                                             And (cSex = 2)    then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH2 ';
                                     End;
                                 end;
                               end;

                              //C011 결과 없을때
                              if Trim(sGulkwa[11]) = '' then
                              begin
                                 if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
                                 else val1 :=0;
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] <> '' )   then
                                     begin
                                         If (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (cSex = 1) and
                                        ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                        ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))      then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH1 ';
                                     End
                                     else If (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                             (cSex = 2) and
                                             ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                             ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))    then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH1 ';
                                     End;
                                 end;
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     If (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (cSex = 1)   and
                                        ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))   and
                                        (sGulkwa[71] = '음성')  and
                                         ((sGulkwa[008]='음성') or (sGulkwa[008]=''))        then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH1 ';
                                     End
                                     else If (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                             (cSex = 2)     and
                                             ((sGulkwa[70] = '음성')  or ((val1>0) and (val1<1.0)))   and
                                             (sGulkwa[071] = '음성')   and
                                             ((sGulkwa[008]='음성') or (sGulkwa[008]=''))     then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH1 ';
                                     End;
                                 end;
                                 ////   NH2
                                 IF  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] <> '' ) then
                                     begin
                                     If ((strTofloat(sGulkwa[009]) >= 40.1) or (strTofloat(sGulkwa[010]) >= 45.1) or
                                        (sGulkwa[70] = '양성') or  (sGulkwa[70] = '약양성') or  (val1>=1.0) or
                                        ((sGulkwa[71] = '양성') or (sGulkwa[071] = '약양성')or (strTofloat(sGulkwa[115]) >=10)) or (sGulkwa[008] ='양성') ) and (cSex = 1) then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH2 ';
                                     End
                                     else If ((strTofloat(sGulkwa[009]) >= 32.1) or (strTofloat(sGulkwa[010]) >= 34.1) or
                                             (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or  (val1>=1.0) or
                                             ((sGulkwa[71] = '양성') or (sGulkwa[071] = '약양성') or (strTofloat(sGulkwa[115]) >=10) ) or
                                             (sGulkwa[008] ='양성'))and  (cSex = 2)  then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH2 ';
                                     End;
                                 end;

                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     if ((strTofloat(sGulkwa[009]) >= 40.1) or (strTofloat(sGulkwa[010]) >= 45.1) or
                                        (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or
                                        ((sGulkwa[70] = '약양성') or  (sGulkwa[70] = '양성')or (val1>=1.0))or (sGulkwa[008] ='양성'))  and
                                        (cSex = 1) then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH2 ';
                                     End
                                     else If ((strTofloat(sGulkwa[009]) >= 32.1) or (strTofloat(sGulkwa[010]) >= 34.1) and
                                             (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or
                                             ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0)) or (sGulkwa[008] ='양성'))
                                             And (cSex = 2)    then
                                             Begin
                                             K := K + 1;
                                             tSokun[K] := 'NH2 ';
                                     End;
                                 end;
                               end;
                              End
                           Else
//혈액소견이 있을때
                              Begin
// 모든대상자 포함 2014.04.24
//                              if qryBunju.FieldByname('Num_bunju').AsInteger < 7000 then
                              //2017.06.27 수지 HORM소견 관련
                                If ((pos('E021',cValue) > 0) OR (pos('E051',cValue) > 0) OR (pos('E052',cValue) > 0) OR
                                    (pos('E053',cValue) > 0) OR (pos('E054',cValue) > 0)) and  (cSex = 2) THEN
                                Begin
                                     K := K + 1;
                                     tSokun[K] := 'HORM';
                                End;
                                //20171124.. 수지   S030 이나 S031이 들어가 있으면 무조건 결과와 상관 없이 PEPC
                                //S076 이 들어가 있으면 무조건 결과와 상관없이 AHPM
                                If (pos('S030',cValue) > 0) OR (pos('S031',cValue) > 0) THEN
                                Begin
                                     K := K + 1;
                                     tSokun[K] := 'PEPC';
                                End;
                                If (pos('S076',cValue) > 0)  THEN
                                Begin
                                     K := K + 1;
                                     tSokun[K] := 'AHPM';
                                End;
                              //C011 결과 있을때
                              if Trim(sGulkwa[11]) <> '' then
                              begin
                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] <> '' )  then
                                     begin
                                     if (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (strTofloat(sGulkwa[011]) <= 70) And (cSex = 1)                       and
                                        ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                         and  ((sGulkwa[008]='음성') or (sGulkwa[008]=''))        then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                         (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2)  and ( Gpreg <> 'Y')    and
                                          ((sGulkwa[71] = '음성')or (strTofloat(sGulkwa[115]) <=9.99)) and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                           and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))
                                           then
                                         Begin
                                         K := K + 1;
                                         tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                         (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2)  and ( Gpreg = 'Y')    and
                                         ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99))  and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0))) and
                                         ((sGulkwa[008]='음성') or (sGulkwa[008]=''))     then
                                         Begin
                                         K := K + 1;
                                         tSokun[K] := 'NJG ';
                                     End;
                                 end;

                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     if (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (strTofloat(sGulkwa[011]) <= 70) And (cSex = 1)                       and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                        and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))  then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                        (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2)   and ( Gpreg <> 'Y')    and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                        and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))             then
                                        Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                        (strTofloat(sGulkwa[011]) <= 42) And (cSex = 2)   and ( Gpreg = 'Y')    and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0))) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))  then
                                        Begin
                                        K := K + 1;
                                        tSokun[K] := 'NJG ';
                                     End;
                                 end;

                                 ////NH4
                                 sokun_Temp:='';
                                 For Cnt := 1 To 182 Do
                                 begin
                                      If tSokun[Cnt]<>'' Then
                                         sokun_Temp:=sokun_Temp+ '.'+tSokun[Cnt]
                                 end;

                                 For dCnt1 := 1 To 182 Do
                                 begin
                                     If ((tSokun[dCnt1] = 'GAH0') Or
                                         (tSokun[dCnt1] = 'GAH ') Or
                                         (tSokun[dCnt1] = 'GAH1') Or
                                         (tSokun[dCnt1] = 'GAH2') Or
                                         (tSokun[dCnt1] = 'GAHH') Or
                                         (tSokun[dCnt1] = 'HPD ') Or
                                         (tSokun[dCnt1] = 'HPM ') Or
                                         (tSokun[dCnt1] = 'FL  ') Or
                                         (tSokun[dCnt1] = 'FLS ')) and (pos('HBAB',sokun_temp) <=0) and (pos('HABW',sokun_temp) <=0) and
                                         ((pos('S008',cValue) > 0) and (sGulkwa[71] <> '양성')) Then
                                     Begin
                                          k := k + 1;
                                          tSokun[k] := 'HABN';
                                     End;
                                 End;


                                 if   (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] <> '' ) and (pos('NH3 ',sokun_temp) <=0)then
                                     begin
                                     if ((strTofloat(sGulkwa[010]) >= 45.1) or   (strTofloat(sGulkwa[011]) >= 70.1) or ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0))
                                         or ((sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') and (strTofloat(sGulkwa[115])>=10.0) )) And (cSex = 1)  then
                                        Begin

                                              K := K + 1;
                                              tSokun[K] := 'NH4 '
                                        end
                                     else if ((strTofloat(sGulkwa[010]) >= 34.1) or  (strTofloat(sGulkwa[011]) >= 42.1) or ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0))
                                             or ((sGulkwa[71] = '양성') or  (sGulkwa[71] = '약양성') and (strTofloat(sGulkwa[115])>=10.0))or (sGulkwa[008] ='양성') ) And (cSex = 2)  then
                                        Begin


                                              K := K + 1;
                                              tSokun[K] := 'NH4 ';
                                        End;
                                     end;

                                 if (sGulkwa[010] <> '' ) and (sGulkwa[011] <> '' ) and (sGulkwa[115] = '' ) and (pos('NH3 ',sokun_temp) <=0) then
                                     begin
                                     if ( (strTofloat(sGulkwa[010]) >= 45.1) or (strTofloat(sGulkwa[011]) >= 70.1)  or (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0)
                                        or (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or (sGulkwa[008] ='양성') ) And (cSex = 1)     then
                                        Begin

                                        K := K + 1;
                                        tSokun[K] := 'NH4 ';
                                     End
                                     else if ((strTofloat(sGulkwa[010]) >= 34.1) or
                                        (strTofloat(sGulkwa[011]) >= 42.1) or (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성')or  (val1>=1.0) or
                                        (sGulkwa[071] = '양성') or (sGulkwa[071] = '약양성')or (sGulkwa[008] ='양성') ) And (cSex = 2)      then
                                        Begin

                                        K := K + 1;
                                        tSokun[K] := 'NH4 ';
                                     End;
                                 end;
                                 ////AST 조건 + 임신X + (S008 or S012)
                                 If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '') and (sGulkwa[11] <> '')  and (sGulkwa[17] <> '')Then
                                 begin
                                     if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
                                     else val1 :=0;
                                     if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
                                     else val2 :=0;
                                     if  (sGulkwa[71]='음성') or((val2>0) and (val2<10.0)) then
                                         begin
                                         if (StrToFloat(sGulkwa[10])<=45)  and (StrToFloat(sGulkwa[09])>=40.1) and
                                            (StrToFloat(sGulkwa[11])<=70)  and
                                            (StrToFloat(sGulkwa[17])<=260) and
                                            ((sGulkwa[70] ='음성') or ((val1>0)and (val1<1.0))) and
                                            ((sGulkwa[008]='음성') or (sGulkwa[008]=''))       and
                                            (cSex = 1)  then
                                            begin
                                            K := K + 1;
                                            tSokun[K] := 'NJB ';
                                         end
                                         else if (StrToFloat(sGulkwa[10])<=34)  and (StrToFloat(sGulkwa[09])>=32.1) and
                                                 (StrToFloat(sGulkwa[11])<=42)    and
                                                 (StrToFloat(sGulkwa[17])<=260)   and
                                                 ((sGulkwa[70] ='음성') or ((val1>0) and (val1<1.0))) and
                                                 (cSex = 2) and
                                                 ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                                  begin
                                                   K := K + 1;
                                                   tSokun[K] := 'NJB ';
                                         end;
                                     end
                                     else if  (sGulkwa[115] <> '') then
                                         begin
                                         if (StrToFloat(sGulkwa[10])<=45)  and (StrToFloat(sGulkwa[09])>=40.1) and
                                            (StrToFloat(sGulkwa[11])<=70)  and
                                            (StrToFloat(sGulkwa[17])<=260) and
                                            ((sGulkwa[70] ='음성') or ((val1>0)and (val1<1.0))) and
                                            (cSex = 1) and
                                            ((strTofloat(sGulkwa[115]) <=9.99) or (sGulkwa[71]='음성'))  and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                            begin
                                            K := K + 1;
                                            tSokun[K] := 'NJB ';
                                         end
                                         else if (StrToFloat(sGulkwa[10])<=34) and (StrToFloat(sGulkwa[09])>=32.1) and
                                                 (StrToFloat(sGulkwa[11])<=42)   and
                                                 (StrToFloat(sGulkwa[17])<=260)   and
                                                 ((sGulkwa[70] ='음성') and ((val1>0)and (val1<1.0)))  and
                                                 (cSex = 2) and
                                                 ((strTofloat(sGulkwa[115]) <=9.99) or (sGulkwa[71]='음성')) and
                                                 ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                                  begin
                                                   K := K + 1;
                                                   tSokun[K] := 'NJB ';
                                         end;
                                     end;
                                 end;
                                 end;

                              //C011이 없을때
                              if Trim(sGulkwa[11]) = '' then
                              begin
                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] <> '' )  then
                                     begin
                                     if (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (cSex = 1)                       and
                                        ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99)) and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                         and  ((sGulkwa[008]='음성') or (sGulkwa[008]=''))        then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                         (cSex = 2)  and ( Gpreg <> 'Y')    and
                                          ((sGulkwa[71] = '음성')or (strTofloat(sGulkwa[115]) <=9.99)) and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                           and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))
                                           then
                                         Begin
                                         K := K + 1;
                                         tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                         (cSex = 2)  and ( Gpreg = 'Y')    and
                                         ((sGulkwa[71] = '음성') or (strTofloat(sGulkwa[115]) <=9.99))  and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0))) and
                                         ((sGulkwa[008]='음성') or (sGulkwa[008]=''))     then
                                         Begin
                                         K := K + 1;
                                         tSokun[K] := 'NJG ';
                                     End;
                                 end;

                                 if  (sGulkwa[009] <> '' ) and (sGulkwa[010] <> '' ) and (sGulkwa[115] = '' ) then
                                     begin
                                     if (strTofloat(sGulkwa[009]) <= 40) and (strTofloat(sGulkwa[010]) <= 45) and
                                        (cSex = 1)                       and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                        and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))  then
                                     Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                        (cSex = 2)   and ( Gpreg <> 'Y')    and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0)))
                                        and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))             then
                                        Begin
                                        K := K + 1;
                                        tSokun[K] := 'NH3 ';
                                     End
                                     else if (strTofloat(sGulkwa[009]) <= 32) and (strTofloat(sGulkwa[010]) <= 34) and
                                        (cSex = 2)   and ( Gpreg = 'Y')    and
                                        (sGulkwa[071] = '음성') and  ((sGulkwa[70] = '음성') or ((val1>0) and (val1<1.0))) and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))  then
                                        Begin
                                        K := K + 1;
                                        tSokun[K] := 'NJG ';
                                     End;
                                 end;
                                 ////NH4
                                 sokun_Temp:='';
                                 For Cnt := 1 To 182 Do
                                 begin
                                      If tSokun[Cnt]<>'' Then
                                         sokun_Temp:=sokun_Temp+ '.'+tSokun[Cnt]
                                 end;

                                 For dCnt1 := 1 To 182 Do
                                 begin
                                     If ((tSokun[dCnt1] = 'GAH0') Or
                                         (tSokun[dCnt1] = 'GAH ') Or
                                         (tSokun[dCnt1] = 'GAH1') Or
                                         (tSokun[dCnt1] = 'GAH2') Or
                                         (tSokun[dCnt1] = 'GAHH') Or
                                         (tSokun[dCnt1] = 'HPD ') Or
                                         (tSokun[dCnt1] = 'HPM ') Or
                                         (tSokun[dCnt1] = 'FL  ') Or
                                         (tSokun[dCnt1] = 'FLS ')) and (pos('HBAB',sokun_temp) <=0) and (pos('HABW',sokun_temp) <=0) and
                                         ((pos('S008',cValue) > 0) and (sGulkwa[71] <> '양성')) Then
                                     Begin
                                          k := k + 1;
                                          tSokun[k] := 'HABN';
                                     End;
                                 End;


                                 if   (sGulkwa[010] <> '' ) and (sGulkwa[115] <> '' ) and (pos('NH3 ',sokun_temp) <=0)then
                                     begin
                                     if ((strTofloat(sGulkwa[010]) >= 45.1) or   ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0))
                                         or ((sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') and (strTofloat(sGulkwa[115])>=10.0) )) And (cSex = 1)  then
                                        Begin

                                              K := K + 1;
                                              tSokun[K] := 'NH4 '
                                        end
                                     else if ((strTofloat(sGulkwa[010]) >= 34.1) or  ((sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0))
                                             or ((sGulkwa[71] = '양성') or  (sGulkwa[71] = '약양성') and (strTofloat(sGulkwa[115])>=10.0))or (sGulkwa[008] ='양성') ) And (cSex = 2)  then
                                        Begin


                                              K := K + 1;
                                              tSokun[K] := 'NH4 ';
                                        End;
                                     end;

                                 if (sGulkwa[010] <> '' ) and (sGulkwa[115] = '' ) and (pos('NH3 ',sokun_temp) <=0) then
                                     begin
                                     if ( (strTofloat(sGulkwa[010]) >= 45.1) or
                                        (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성') or (val1>=1.0)
                                        or (sGulkwa[71] = '양성') or (sGulkwa[71] = '약양성') or (sGulkwa[008] ='양성') ) And (cSex = 1)     then
                                        Begin

                                        K := K + 1;
                                        tSokun[K] := 'NH4 ';
                                     End
                                     else if ((strTofloat(sGulkwa[010]) >= 34.1) or
                                        (sGulkwa[70] = '양성') or (sGulkwa[70] = '약양성')or  (val1>=1.0) or
                                        (sGulkwa[071] = '양성') or (sGulkwa[071] = '약양성')or (sGulkwa[008] ='양성') ) And (cSex = 2)      then
                                        Begin

                                        K := K + 1;
                                        tSokun[K] := 'NH4 ';
                                     End;
                                 end;
                                 ////AST 조건 + 임신X + (S008 or S012)
                                 If (sGulkwa[09] <> '')  and (sGulkwa[10] <> '')  and (sGulkwa[17] <> '')Then
                                 begin
                                     if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
                                     else val1 :=0;
                                     if (sGulkwa[115])<> '' then Val2 := StrToFloat(sGulkwa[115])
                                     else val2 :=0;
                                     if  (sGulkwa[71]='음성') or((val2>0) and (val2<10.0)) then
                                         begin
                                         if (StrToFloat(sGulkwa[10])<=45)  and (StrToFloat(sGulkwa[09])>=40.1) and
                                            (StrToFloat(sGulkwa[17])<=260) and
                                            ((sGulkwa[70] ='음성') or ((val1>0)and (val1<1.0))) and
                                            ((sGulkwa[008]='음성') or (sGulkwa[008]=''))       and
                                            (cSex = 1)  then
                                            begin
                                            K := K + 1;
                                            tSokun[K] := 'NJB ';
                                         end
                                         else if (StrToFloat(sGulkwa[10])<=34)  and (StrToFloat(sGulkwa[09])>=32.1) and
                                                 (StrToFloat(sGulkwa[17])<=260)   and
                                                 ((sGulkwa[70] ='음성') or ((val1>0) and (val1<1.0))) and
                                                 (cSex = 2) and
                                                 ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                                  begin
                                                   K := K + 1;
                                                   tSokun[K] := 'NJB ';
                                         end;
                                     end
                                     else if  (sGulkwa[115] <> '') then
                                         begin
                                         if (StrToFloat(sGulkwa[10])<=45)  and (StrToFloat(sGulkwa[09])>=40.1) and
                                            (StrToFloat(sGulkwa[17])<=260) and
                                            ((sGulkwa[70] ='음성') or ((val1>0)and (val1<1.0))) and
                                            (cSex = 1) and
                                            ((strTofloat(sGulkwa[115]) <=9.99) or (sGulkwa[71]='음성'))  and ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                            begin
                                            K := K + 1;
                                            tSokun[K] := 'NJB ';
                                         end
                                         else if (StrToFloat(sGulkwa[10])<=34) and (StrToFloat(sGulkwa[09])>=32.1) and
                                                 (StrToFloat(sGulkwa[17])<=260)   and
                                                 ((sGulkwa[70] ='음성') and ((val1>0)and (val1<1.0)))  and
                                                 (cSex = 2) and
                                                 ((strTofloat(sGulkwa[115]) <=9.99) or (sGulkwa[71]='음성')) and
                                                 ((sGulkwa[008]='음성') or (sGulkwa[008]=''))then
                                                  begin
                                                   K := K + 1;
                                                   tSokun[K] := 'NJB ';
                                         end;
                                     end;
                                 end;
                                 end;


                              End;

                           If wk_sw1 = 99 Then sGulkwa[71] := '';
                           With qryKicho Do
                                Begin
                                     Close;
                                     ParamByName('In_Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
                                     ParamByName('In_Num_Id').AsString := qryGumgin.FieldByname('Num_id').Asstring;
                                     ParamByName('In_Dat_Gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
                                     Open;
                                     If RecordCount > 0 Then
                                        Begin
                                           Biman := FieldByname('Biman').AsFloat;
                                           Hul_h := FieldByname('Hul_h').ASInteger;
                                           Hul_l := FieldByname('Hul_l').AsInteger;
                                        End
                                     Else
                                        Begin
                                           Biman := 100.0;
                                           Hul_h := 0;
                                           Hul_l := 0;
                                        End;
                                     Close;
                                End;
                           {  //20140913 곽윤설 주석처리
                           Chk_Dang_Drug := '';         //통합문진 당뇨 확인
                           Chk_Dang_Jindan := '';
                           With QryTot_Munjin2010 Do
                           Begin
                              Close;
                              ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
                              ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
                              ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
                              Open;
                              If RecordCount > 0 Then
                              Begin
                                 Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                                 Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
                              End;
                              Close;
                           End;
                           }

                           End_Sokun_Rtn;
                           // Chuga_Sokun_Rtn;
                           //소견 Update Routine      Chuga_Sokun_Rtn;
                           Insert_Update_Sokun;
                           sCnt := sCnt + 1;
                           ProBar.Progress := sCnt;
                           Close;
                        End;
                  End;
               Next;
            End;
      End;
   ProBar.Progress := 100;
   ProBar.Enabled := False;
   Animate1.Active := False;
   Animate2.Active := False;
   GF_MsgBox('Information', '작업이 정상적으로 수행되었습니다.');
end;

Procedure TfrmLD6I01.End_Sokun_Rtn;
Var
   dcnt1, dCnt2, dCnt3, lcheck : Integer;

Begin
// 모든대상자 포함 2014.04.24
//      if qryBunju.FieldByName('num_bunju').Asinteger < 7000  then
      if True  then
      begin
//식이요법 처리
// 2004.07.05일 수정 => PREG가 나오면 다른 소견 X.
      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'AFG ') or (tSokun[dCnt1] = 'AFPF')Then
             dCheck := dCheck + 1;
      If (dCheck <> 0) Then
         Begin
            S := S + 1;
            sSokun[S] := 'PREG';
         End
         else
         begin



//간장질환식이요법
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'HPM ') or
                   (tSokun[dCnt1] = 'HPB ') Then
                   dCheck := dCheck + 1;
            If dCheck > 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'ALD1';
               End;
            If sSokun[S] <> 'ALD1' Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      IF (tSokun[dCnt1] = 'HO2 ') or
                         (tSokun[dCnt1] = 'FL  ')  or
                         (tsokun[dCnt1] = 'GAHH') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'FLD ';
                     End;
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      IF (tSokun[dCnt1] = 'GAH ') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'HD  ';
                     End;
               End;
            If (sSokun[S] <> 'ALD1') And (sSokun[S] <> 'FDL ') Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      If (tSokun[dCnt1] = 'HBM ') Or
                         (tsokun[dcnt1] = 'HCV1') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'ALD ';
                     End;
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      If (tSokun[dCnt1] = 'HPD ') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'HD  ';
                     End;
               End;
//당뇨질환식이요법
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'DIB ') or
                   (tSokun[dCnt1] = 'DIC ') Then
                   dCheck := dCheck + 1;
            If dCheck > 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'DD2 ';
               End;
            If sSokun[S] <>  'DD2 ' Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 14 Do
                      If tSokun[dCnt1] = 'DIB ' Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'DD1 ';
                     End;
               End;
//순환기질환식이요법
            dCheck := 0;
            lcheck := 0;
            For dCnt1 := 1 To 182 Do
            begin
                If (tSokun[dCnt1] = 'CTH ') or (tSokun[dCnt1] = 'CTL ') Then
                   dCheck := dCheck + 1
                else if (tSokun[dCnt1] = 'LDLN') or (tSokun[dCnt1] = 'LDLH') or (tSokun[dCnt1] = 'TH  ') Then
                   lcheck := lcheck + 1;
            end;
            If (dCheck > 0) OR (lcheck > 1) Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'CTHD';
               End;
            If sSokun[S] <> 'CTHD' Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      If (tSokun[dCnt1] = 'CH  ') or (tSokun[dCnt1] = 'CLH ') or (tSokun[dCnt1] = 'CH1 ') or
                         (tSokun[dCnt1] = 'LDLN') or (tSokun[dCnt1] = 'LDLH') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'CHD ';
                     End;
               End;
            If (sSokun[S] <> 'CTHD') And (sSokun[S] <> 'CHD ') Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      If (tSokun[dCnt1] = 'TH  ') Or
                         (tSokun[dCnt1] = 'TLH ') Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                        S := S + 1;
                        sSokun[S] := 'THD ';
                     End;
               End;
//혈액질환식이요법
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'IDA ') Or (tsokun[dCnt1] = 'IDO ') or (tsokun[dCnt1] = 'IDO2') or (tsokun[dCnt1] = 'IDA6') or
                   (tSokun[dCnt1] = 'IDPM') Or (tSokun[dCnt1] = 'IDOM') Then
                   dCheck := dCheck + 1;
            If dCheck > 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'IDA3';
               End;
            If sSokun[S] <> 'IDA3' Then
               Begin
                  dCheck := 0;
                  For dCnt1 := 1 To 182 Do
                      If tSokun[dCnt1] = 'ANE ' Then
                         dCheck := dCheck + 1;
                  If dCheck > 0 Then
                     Begin
                       S := S + 1;
                       sSokun[S] := 'ADD2';
                     End;
               End;
//기타식이요법
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'HBC ') Then
                   dCheck := dCheck + 1;
            If dCheck > 0 Then
               Begin
                   S := S + 1;
                   sSokun[S] := 'HD  ';
               End;
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'THH ') Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'THHD';
               End;
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'THL ') Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'THLD';
               End;
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'U1  ') Or
                   (tSokun[dCnt1] = 'U   ') Then
                   dCheck := dCheck + 1;
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (sSokun[dCnt1] = 'FLD ') Or
                   (sSokun[dCnt1] = 'CHD ') Or
                   (sSokun[dCnt1] = 'CTHD') Then
                   dCheck := 0;
            If dCheck <> 0 Then
               Begin
                  S := S + 1;
                  sSokun[S] := 'URA1';
               End;
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'BCH ') Or
                   (tSokun[dCnt1] = 'BCH1') or
                   (tSokun[dCnt1] = 'U03 ') Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                   S := S + 1;
                   sSokun[S] := 'RD  ';
               End;
            dCnt2 := 0;
            For dCnt1 := 1 To 182 Do
              If sSokun[dCnt1] <> '' Then
                 dCnt2 := 99;
            If dCnt2 = 0 Then
               If Biman <= 100 Then
                  Begin
                     S := S + 1;
                     sSokun[S] := 'SGD ';
                  End
               Else
                  Begin
                     S := S + 1;
                     sSokun[S] := 'HF  ';
                  End;
         end; // PREG가 나오면 다른 식이요법은 없음.
//운동 요법
      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'AFPF') or (tSokun[dCnt1] = 'AFG ')  Then
             dCheck := dCheck + 1;
      If (dCheck <> 0) Then
         Begin
            U := U + 1;
            uSokun[U] := 'EW  ';
            Exit;
         End;


      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'DIB ') or
             (tSokun[dCnt1] = 'DIC ') Then
             dCheck := dCheck + 1;
      If dCheck > 0 Then
         Begin
            U := U + 1;
            USokun[U] := 'ED2 ';
         End;
      If uSokun[U] <> 'ED2 ' Then
         Begin
            dCheck := 0;
            For dCnt1 := 1 To 13 Do
               If (tSokun[dCnt1] = 'DIB ') Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'ED1 ';
               End;
         End;
//      If uSokun[U] <> 'ED1 ' Then
//         Begin
            If (Biman >= 120) And (Biman <= 10000) Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'EX4 ';
                  If sAge < 60 Then
                    Begin
                       U := U + 1;
                       uSokun[U] := 'EX2 ';
                    End;
               End;
//         End;
//      If (uSokun[U] <> 'ED1 ') And (usokun[U] = 'EX4 ') Then
//         Begin
            If (Biman >= 110) And (Biman <= 119) Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'EX3 ';
                  If sAge < 60 Then
                    Begin
                       U := U + 1;
                       uSokun[U] := 'EX2 ';
                    End;
               End;
//         End;
      If (uSokun[U] <> 'ED1 ') And (usokun[U] = 'EX4 ') And
         (usokun[U] <> 'EX3 ') Then
         Begin
            dCheck := 0;
            For dCnt1 := 1 To 14 Do
                If tSokun[dCnt1] = 'HBC ' Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'EXG ';
               End;
         End;
//      If (uSokun[U] <> 'ED1 ') And (usokun[U] = 'EX4 ') And
//         (usokun[U] <> 'EX3 ') And (uSokun[U] = 'EXG ') Then
//         Begin
//            dCheck := 0;
//            For dCnt1 := 1 To 99 Do
//                If (tSokun[dCnt1] = 'CH  ') Or
//                   (tSokun[dCnt1] = 'CTH ') Or
//                   (tSokun[dCnt1] = 'CTL ') Or
//                   (tsokun[dCnt1] = 'TH  ') Then
//                   dCheck := dCheck + 1;
//            If dCheck <> 0 Then
//               Begin
//                  U := U + 1;
//                  uSokun[U] := 'EXM2';
//               End;
//         End;
      If sAge >= 60 Then
         Begin
            U := U + 1;
            uSokun[U] := 'EXM2';
         End;
      If (uSokun[U] <> 'ED1 ') And (usokun[U] = 'EX4 ') And
         (usokun[U] <> 'EX3 ') And (uSokun[U] = 'EXG ') And (uSokun[U] = 'EXM2') Then
         Begin
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
                If (tSokun[dCnt1] = 'GAH ') Then
                   dCheck := dCheck + 1;
            If dCheck <> 0 Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'EXT ';
               End;
         End;
      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'CH  ') Or
             (tSokun[dCnt1] = 'TH  ') Or
             (tSokun[dCnt1] = 'CTH ') Then
             dCheck := dCheck + 1;
      If (dCheck <> 0) And (Biman <= 109) Then
         Begin
            U := U + 1;
            uSokun[U] := 'HL1 ';
         End;

      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'CH  ') Or
             (tSokun[dCnt1] = 'TH  ') Or
             (tSokun[dCnt1] = 'CTH ') Then
             dCheck := dCheck + 1;
      If (dCheck <> 0) And (Biman <= 109) Then
         Begin
            dCheck := 0;
            For dCnt1 := 1 To 182 Do
               If (uSokun[dCnt1] = 'HL1 ') And
                  (uSokun[dCnt1] = 'EX2 ') Then
                   dCheck := dCheck + 1;
            If dCheck = 0 Then
               Begin
                  U := U + 1;
                  uSokun[U] := 'EX2 ';
               End;
         End;


      dCnt2 := 0;
      For dCnt1 := 1 To 182 Do
        If uSokun[dCnt1] <> '' Then
           dCnt2 := 99;
      If dCnt2 = 0 Then
         Begin
            U := U + 1;
            uSokun[U] := 'EW2 ';
         End;
      end;

End;
Procedure TfrmLD6I01.Chuga_Sokun_Rtn;
Var
   sbiman    : Single;
   shul_h, sHul_l : Integer;
   dcnt1: Integer;
   Val1 : Extended;
   Munjin_drug,Temp :String;
   sokun_Check :boolean;
Begin
// 모든대상자 포함 2014.04.24
//      if qryBunju.FieldByName('num_bunju').Asinteger < 7000  then
      if True  then
      begin
           Temp:='';
           For dCnt1 := 1 To 182 Do
           begin
                If tSokun[dCnt1]<>'' Then
                   Temp:=Temp+ '.'+tSokun[dCnt1]
           end;
           With qryKicho Do
           Begin
                Close;
                ParamByName('In_Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
                ParamByName('In_Num_Id').AsString := qryGumgin.FieldByname('Num_id').Asstring;
                ParamByName('In_Dat_Gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
                Open;
                If RecordCount > 0 Then
                   Begin
                     sBiman := FieldByname('Biman').AsFloat;
                     sHul_h := FieldByname('Hul_h').ASInteger;
                     sHul_l := FieldByname('Hul_l').AsInteger;
                   End

           End;
           if sGulkwa[127] <> '' then
           begin
                val1 := StrToFloat(sGulkwa[127]);
                if (pos('LDLP', Temp) <= 0) and
                   (pos('LDLO', Temp) <= 0) and
                   (pos('LDLK', Temp) <= 0) and
                   (pos('TH3 ', Temp) <= 0) and
                   (pos('TH  ', Temp) <= 0) and
                   (pos('TH5 ', Temp) <= 0) then
                Begin
                     if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBQ';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBR';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90)) and (sbiman< 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBC';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90)) and (sbiman>= 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBS';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h=0) and  (shul_l=0)) and (sbiman=0 )   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBM';
                     end;
                end
                else If   (pos('LDLP', Temp) > 0) or
                          (pos('LDLO', Temp) > 0) or
                          (pos('LDLK', Temp) > 0) or
                          (pos('TH3 ', Temp) > 0) or
                          (pos('TH  ', Temp) > 0) or
                          (pos('TH5 ', Temp) > 0) then
                Begin

                    if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBP';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBT';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90)) and (sbiman< 120)   then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBD';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90)) and (sbiman>= 120)  then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBU';
                     end
                     else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h=0) and (shul_l=0)) and (sbiman= 0)  then
                     begin
                          k := k + 1;
                          tSokun[k] := 'APBN';
                     end;
                End;
            end;
            if (sGulkwa[127] <> '') and (sGulkwa[20] = '') and (sGulkwa[21] ='') then
            begin
                 Val1 := StrToFloat(sGulkwa[127]);
                 if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APBV';
                      end
                 else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APBW';
                      end
                 else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APBX';
                      end
                  else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman> 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APBY';
                      end
                 else if  (((val1>=160.1) and(cSex = 1)) or ((val1>=150.1) and(cSex = 2)))   and ((shul_h=0) and  (shul_l=0)) and (sbiman=0)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APBZ';
                      end;

            end;
            if sGulkwa[116] <> '' then
            begin
                 Val1 := StrToFloat(sGulkwa[116]);

                 if (pos('LDLP', Temp) <= 0) and
                    (pos('LDLO', Temp) <= 0) and
                    (pos('LDLK', Temp) <= 0) and
                    (pos('TH3 ', Temp) <= 0) and
                    (pos('TH  ', Temp) <= 0) and
                    (pos('TH5 ', Temp) <= 0) then
                 Begin
                      if  (val1>=15.1)    and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMQ';
                      end
                      else if  (val1>=15.1)    and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMR';
                      end
                      else if  (val1>=15.1)    and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMC';
                      end
                      else if  (val1>=15.1)    and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMS';
                      end
                      else if  (val1>=15.1)  and ((shul_h=0) and  (shul_l=0)) and (sbiman=0 )   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMM';
                      end;
                 end
                 else If  (pos('LDLP', Temp) > 0) or
                          (pos('LDLO', Temp) > 0) or
                          (pos('LDLK', Temp) > 0) or
                          (pos('TH3 ', Temp) > 0) or
                          (pos('TH  ', Temp) > 0) or
                          (pos('TH5 ', Temp) > 0) then
                  Begin
                      if  (val1>=15.1) and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                       begin
                            k := k + 1;
                            tSokun[k] := 'HOMP';

                       end
                       else if (val1>=15.1)  and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                       begin
                            k := k + 1;
                            tSokun[k] := 'HOMT';

                       end
                       else if  (val1>=15.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                       begin
                            k := k + 1;
                            tSokun[k] := 'HOMD';

                       end
                       else if  (val1>=15.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman>= 120)  then
                       begin
                            k := k + 1;
                            tSokun[k] := 'HOMU';

                       end
                       else if  (val1>=15.1)   and ((shul_h=0) and (shul_l=0)) and (sbiman= 0)  then
                       begin
                            k := k + 1;
                            tSokun[k] := 'HOMN';

                       end;
                 end;
            end;
            if (sGulkwa[116] <> '') and (sGulkwa[20] = '') and (sGulkwa[21] ='') then
            begin
                 Val1 := StrToFloat(sGulkwa[116]);

                 if   (val1>=15.1)  and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMV';

                      end
                 else if   (val1>=15.1)   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMW';

                      end
                 else if   (val1>=15.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMX';

                      end
                  else if   (val1>=15.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90)) and (sbiman> 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMY';

                      end
                 else if   (val1>=15.1)  and ((shul_h=0) and  (shul_l=0)) and (sbiman=0)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'HOMZ';

                 end;
            end;
            if sGulkwa[128] <> '' then
            begin
                 Val1 := StrToFloat(sGulkwa[128]);
                 if (pos('LDLP', Temp) <= 0) and
                    (pos('LDLO', Temp) <= 0) and
                    (pos('LDLK', Temp) <= 0) and
                    (pos('TH3 ', Temp) <= 0) and
                    (pos('TH  ', Temp) <= 0) and
                    (pos('TH5 ', Temp) <= 0) then
                    Begin
                    if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                    begin
                         k := k + 1;
                         tSokun[k] := 'APAQ';

                    end
                    else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                    begin
                         k := k + 1;
                         tSokun[k] := 'APAR';

                    end
                    else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                    begin
                         k := k + 1;
                         tSokun[k] := 'APAC';

                    end
                    else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman>= 120)   then
                    begin
                         k := k + 1;
                         tSokun[k] := 'APAS';

                    end
                    else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))  and ((shul_h=0) and  (shul_l=0)) and (sbiman=0 )   then
                    begin
                         k := k + 1;
                         tSokun[k] := 'APAM';

                    end
                end
                else If (pos('LDLP', Temp) > 0) or
                        (pos('LDLO', Temp) > 0) or
                        (pos('LDLK', Temp) > 0) or
                        (pos('TH3 ', Temp) > 0) or
                        (pos('TH  ', Temp) > 0) or
                        (pos('TH5 ', Temp) > 0) then
                         Begin
                         if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                         begin
                             k := k + 1;
                             tSokun[k] := 'APAP';

                         end
                         else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                         begin
                              k := k + 1;
                              tSokun[k] := 'APAT';
                         end
                         else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                         begin
                              k := k + 1;
                              tSokun[k] := 'APAD';
                         end
                         else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman>= 120)  then
                         begin
                              k := k + 1;
                              tSokun[k] := 'APAU';
                         end
                         else if  (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))  and ((shul_h=0) and (shul_l=0)) and (sbiman= 0)  then
                         begin
                              k := k + 1;
                              tSokun[k] := 'APAN';
                         end;
                End;
            end;
            If (sGulkwa[173] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[173]);

               If (Val1 >= 15.01)  Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'RAHQ';
               End;
            End;
            if (sGulkwa[128] <> '') and (sGulkwa[20] = '') and (sGulkwa[21] ='') then
            begin
                 Val1 := StrToFloat(sGulkwa[128]);
                 if   (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APAV';
                      end
                 else if    (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APAW';
                      end
                 else if    (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))    and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APAX';
                      end
                  else if    (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman> 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APAY';
                      end
                 else if    (((val1<=109.9) and(cSex = 1)) or ((val1<=119.9) and(cSex = 2)))   and ((shul_h=0) and  (shul_l=0)) and (sbiman=0)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'APAZ';
                      end;
             end;

            if sGulkwa[129] <> '' then
            begin
                 Val1 := StrToFloat(sGulkwa[129]);

                 if (pos('LDLP', Temp) <= 0) and
                    (pos('LDLO', Temp) <= 0) and
                    (pos('LDLK', Temp) <= 0) and
                    (pos('TH3 ', Temp) <= 0) and
                    (pos('TH  ', Temp) <= 0) and
                    (pos('TH5 ', Temp) <= 0) then
                 Begin
                      if  (val1>=30.1)    and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAQ';
                      end
                      else if  (val1>=30.1)   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAR';
                      end
                      else if (val1>=30.1)    and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAC';
                      end
                      else if  (val1>=30.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAS';
                      end
                      else if  (val1>=30.1)  and ((shul_h=0) and  (shul_l=0)) and (sbiman=0 )   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAM';
                      end;
                 end

                 else If   (pos('LDLP', Temp) > 0) or
                           (pos('LDLO', Temp) > 0) or
                           (pos('LDLK', Temp) > 0) or
                           (pos('TH3 ', Temp) > 0) or
                           (pos('TH  ', Temp) > 0) or
                           (pos('TH5 ', Temp) > 0) then
                 Begin
                     if  (val1>=30.1)  and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAP';
                      end
                      else if  (val1>=30.1)   and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAT';
                      end
                      else if  (val1>=30.1)  and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAD';
                      end
                      else if  (val1>=30.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))   and (sbiman>= 120)  then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAU';
                      end
                      else if  (val1>=30.1)   and ((shul_h=0) and (shul_l=0)) and (sbiman= 0)  then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAN';
                      end;
                 End;
             end;
            if (sGulkwa[129] <> '') and (sGulkwa[20] = '') and (sGulkwa[21] ='') then
            begin
                 Val1 := StrToFloat(sGulkwa[129]);
                 if  (val1>=30.1) and ((shul_h>=130) or (shul_l>=90)) and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAV';
                      end
                 else if (val1>=30.1) and ((shul_h>=130) or (shul_l>=90)) and (sbiman>= 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAW';
                      end
                 else if   (val1>=30.1)   and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman< 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAX';
                      end
                 else if  (val1>=30.1)  and ((shul_h<130)and (shul_h>0) and (shul_l>0)and (shul_l<90))  and (sbiman> 120)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAY';
                      end
                 else if   (val1>=30.1)  and ((shul_h=0) and  (shul_l=0)) and (sbiman=0)   then
                      begin
                           k := k + 1;
                           tSokun[k] := 'LPAZ';
                           end;
                 end;
            end;



End;



Procedure TfrmLD6I01.Insert_Update_sokun;
Var
   rCnt, rCheck, rSw,dCnt1,Cnt : Integer;
   Val1 : Single;
   sValue,Munjin_drug : String;
Begin

//혈액소견Update
     sValue := '';
     If sAge < 14 Then
        Begin
           For rCnt := 1 To 182 Do
              Begin
                 sSokun[rCnt] := '';
                 usokun[rCnt] := '';
              End;
           //[2006.06.01-소견10개로 늘림]
           For rCnt := 1 To 10 Do //20131026
              Begin
                 sValue := sValue + tSokun[rCnt];
                 If tSokun[rCnt] = 'NH  ' Then
                    tsokun[rCnt] := 'NNNN'
                 Else If (tsokun[rCnt] = 'NNG ')or
                         (tSokun[rCnt] = 'NFF ')or
                         (tSokun[rCnt] = 'NMM ')Then
                         tSokun[rCnt] := 'NNNM'
                 Else If (tsokun[rCnt] = 'NJ  ')Then
                         tSokun[rCnt] := 'BBBB'
                 Else If (tsokun[rCnt] = 'NNM ')Then
                         tSokun[rCnt] := 'BBBN';
              End;
        End;
        Munjin_drug:='';
        With QryTot_Munjin2018 Do
              Begin
              Close;
              ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
              ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
              ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
              Open;
              If RecordCount > 0 Then
                 Begin
                    If (FieldByName('Mun1_Jindan4').AsString = '1') Or    //2014.08.05 약물복용 - > 복용 Or 진단으로
                       (FieldByname('Mun1_Drug4').AsString   = '1') Then
                    Munjin_drug := '1';
                 End;
              Close;
        End;
        if  Munjin_drug ='1' then
        begin
             For rCnt := 1 To 182 Do
             begin
             If (tSokun[rCnt] = 'DIB ') or
                (tSokun[rCnt] = 'DIB4') or
                (tSokun[rCnt] = 'DMN ') or
                (tSokun[rCnt] = 'DMNH') then
                Begin
                     tSokun[rCnt] := 'DIBT';
                     For Cnt := 1 To 182 Do
                     begin
                     if (tSokun[Cnt] = 'USH1') then
                         tSokun[Cnt] := 'USH2';
                     end;
                End
             else If (tSokun[rCnt] = 'DIH ') or
                     (tSokun[rCnt] = 'DIH4') or
                     (tSokun[rCnt] = 'DIH5') then    //2014.03.31 김소영 선생님 의뢰
                Begin
                     tSokun[rCnt] := 'DIHT';
                     For Cnt := 1 To 182 Do
                     begin
                     if (tSokun[Cnt] = 'USH1') then
                         tSokun[Cnt] := 'USH2';
                     end;
                End
             else If (tSokun[rCnt] = 'DIC ') or
                     (tSokun[rCnt] = 'DIC4') then
                Begin
                     tSokun[rCnt] := 'DICT';
                     For Cnt := 1 To 182 Do
                     begin
                     if (tSokun[Cnt] = 'USH1') then
                         tSokun[Cnt] := 'USH2';
                     end;
                End
             else If (tSokun[rCnt] = 'DIN ')  then
                Begin
                     tSokun[rCnt] := 'DINT';
                End
             else If (tSokun[rCnt] = 'DIO ')  then
                Begin
                     tSokun[rCnt] := 'DIOT';
                End;
             end;
        end;

        For Cnt := 1 To 182 Do
        begin
             If (tSokun[Cnt] = 'GAH0') Or
                (tSokun[Cnt] = 'GAH ') Or
                (tSokun[Cnt] = 'GAH1') Or
                (tSokun[Cnt] = 'GAH2') Or
                (tSokun[Cnt] = 'GAHH') Or
                (tSokun[Cnt] = 'HPD ') Or
                (tSokun[Cnt] = 'HPM ') Or
                (tSokun[Cnt] = 'FL  ') Or
                (tSokun[Cnt] = 'FLS ')  Then
                Begin
                For rCnt := 1 To 182 Do
                begin
                    if (tSokun[rCnt] = 'HABW') then
                       tSokun[rCnt] := 'HABL';
                End;
             end;
         End;


//간장질환의 마지막소견 변경
     sValue := '';
     //[2006.06.01-소견10개로 늘림]
     For rCnt := 1 To 10 Do          //20131026
         If (tSokun[rCnt] = 'HPB ') or
            (tSokun[rCnt] = 'HCV ') or
            (tSokun[rCnt] = 'HBM ') or
            (tSokun[rCnt] = 'HPM ') or
            (tSokun[rCnt] = 'FL  ') or
            (tSokun[rCnt] = 'HPD ') or
            (tSokun[rCnt] = 'GAHH') or
            (tSokun[rCnt] = 'GAH ') or
            (tSokun[rCnt] = 'HBC ') or
            (tSokun[rCnt] = 'HAB ') or
            (tSokun[rCnt] = 'HCV1') Then
            //[2006.06.01-소견10개로 늘림]
            For rCheck := rCnt to 10 Do    //20131026
                If tSokun[rCheck] = 'NJ  ' Then
                   tsokun[rCheck] := 'NNM '
                Else
                   If tsokun[rCheck] = 'GNHG' Then
                      tSokun[rCheck] := 'GNGG';
//소견 Convert Routine 처리(NNM->GNGG, NJ->GNHG, NH->NHM(남),NHF(여),NFF,NMM->NFM)
     sValue := '';
     //[2006.06.01-소견10개로 늘림]
     For rCnt := 1 To 11 Do
         If (tSokun[rCnt] = 'NNM ') or
            (tSokun[rCnt] = 'NJ  ') or
            (tSokun[rCnt] = 'NH  ') or
            (tSokun[rCnt] = 'NFF ') or
            (tSokun[rCnt] = 'NMM ') Then
             Begin
                rSw := 0;
{// C005 (1.3 >= C005 <= 1.4)
                If sGulkwa[05] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[05]);
                      If (Val1 >= 1.3) And (Val1 <= 1.4) Then
                         rSw := 99;
                   End;
}
// C009 (32 >= C009 <= 39 And 여자)
                If sGulkwa[09] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[09]);
                      If (Val1 >= 32) And (Val1 <= 39) And (cSex = 2) Then
                         rSw := 99;
                   End;
// C010 (32 >= C010 <= 39 And 여자)
                If sGulkwa[10] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[10]);
                      If (Val1 >= 32) And (Val1 <= 39) And (cSex = 2) Then
                         rSw := 99;
                   End;
// C011 (33 >= C011 <= 40 And 여자)
                If sGulkwa[11] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[11]);
                      If (Val1 >= 33) And (Val1 <= 40) And (cSex = 2) Then
                         rSw := 99;
                   End;
// C021 (451 >= C021 <= 499)
                If sGulkwa[17] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[17]);
                      If (Val1 >= 451) And (Val1 <= 499) Then
                         rSw := 99;
                   End;
// C028 (150 >= C028 <= 199)
                If sGulkwa[21] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[21]);
                      If (Val1 >= 150) And (Val1 <= 199) Then
                         rSw := 99;
                   End;
// C041 (21 >= C041 <= 29)
                If sGulkwa[29] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[29]);
                      If (Val1 >= 21) And (Val1 <= 29) Then
                         rSw := 99;
                   End;
// C045 (111 >= C045 <= 114)
                If sGulkwa[33] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[33]);
                      If (Val1 >= 111) And (Val1 <= 114) Then
                         rSw := 99;
                   End;
// H001 (3.89 <= H001)
                If sGulkwa[43] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[43]);
                      If Val1 <= 3.89 Then
                         rSw := 99;
                   End;
// H011 (3.1 >= H011 <= 3.9)
                If sGulkwa[53] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[53]);
                      If ((Val1 >= 3.1) And (Val1 <= 3.9)) Then
                         rSw := 99;
                   End;
// U006 (10 >= U006 <= 20)
// U006 (10 >= U006 <= 50)    --20110620 김희석
                If sGulkwa[94] <> '' Then
                   Begin
                      Val1 := StrtoFloat(sGulkwa[94]);
                      If (Val1 >= 10) And (Val1 <= 50) Then
                         rSw := 99;
                   End;
                If rSw <> 0 Then
                   Begin
                      If tSokun[rCnt] = 'NNM ' Then
                         tSokun[rCnt] := 'GNGG';
                      If tSokun[rCnt] = 'NJ  ' Then
                         tSokun[rCnt] := 'GNHG';
                      If tSokun[rCnt] = 'NH  ' Then
                         If cSex = 1 Then
                            tSokun[rCnt] := 'NHM '
                         Else
                            If cSex = 2 Then
                               tSokun[rCnt] := 'NHF ';
                      If (tSokun[rCnt] = 'NFF ') or
                         (tSokun[rCnt] = 'NMM ') Then
                         tSokun[rCnt] := 'NFM ';
                   End;
             End;
{
     If (Hul_h >= 160) Then
     Begin
        For rCnt := 1 To 99 Do
           If tSokun[rCnt] = '' Then
           Begin
              tSokun[rCnt] := 'HT2 ';
              Break;
           End;
     End
     Else If ((Hul_h >= 140) And (Hul_l >= 100)) Or
             ((Hul_h >= 151) And (Hul_l >= 110)) Then
     Begin
        For rCnt := 1 To 99 Do
           If tSokun[rCnt] = '' Then
           Begin
              tSokun[rCnt] := 'HT2 ';
              Break;
           End;
     End
     Else If (Hul_h >= 140) And (Hul_h <= 150) And
             (Hul_l >= 90) And (Hul_l <= 99) Then
     Begin
        For rCnt := 1 To 99 Do
           If tSokun[rCnt] = '' Then
           Begin
              tSokun[rCnt] := 'HT1 ';
              Break;
           End;
     End;
}
     sValue := '';
     //[2006.06.01-소견10개로 늘림]
     For rCnt := 1 To 30 Do   //20160321
      If tSokun[rCnt] <> '' Then
        With qryHul_Sokun Do
           Begin
              Close;
              ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
              ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
              ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
              ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
              ParamByName('In_Cod_sokun').AsString := tSokun[rCnt];
              ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
              ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
              ParamByName('In_Gubn_hsok').AsString := '1';
              Open;
              If RecordCount > 0 Then rCheck := 99
              Else                    rCheck := 00;
              Close;
              If rCheck = 00 Then
                 With qryInsert Do
                    Begin
                       Close;
                       ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
                       ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
                       ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
                       ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
                       ParamByName('In_Cod_sokun').AsString := tSokun[rCnt];
                       ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
                       ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
                       With qrySokun Do
                       Begin
                          Close;
                          ParamByName('In_Cod_sokun').AsString := tSokun[rCnt];
                          Open;

                          qryInsert.ParamByName('In_Gubn_pan').AsString   := FieldByName('GUBN_PNJG').AsString;

                          qryInsert.ParamByName('In_Desc_sokun').AsString := FieldByName('Desc_sbsg').AsString  +
                                                                             FieldByName('Desc_sbsg1').AsString +
                                                                             FieldByName('Desc_sbsg2').AsString  +
                                                                             FieldByName('Desc_sbsg3').AsString;
                          Close;
                       End;

                       ParamByName('In_Gubn_hsok').AsString := '1';

                       if cmbDoctor.Text <> '' then ParamByname('In_Cod_Doct').AsString := Copy(cmbDoctor.Text,1,6)
                       else                         ParamByname('In_Cod_Doct').AsString := '150005';

                       ParamByName('In_Num_Seq').AsInteger := rCnt;
                       ParamByName('In_Dat_Insert').AsString := GV_sToDay;
                       ParamByname('In_Cod_Insert').AsString := GV_sUserId;
                       ParamByname('In_Dat_Update').AsString := '';
                       ParamByName('In_Cod_Update').AsString := '';
                       ExecSql;
                       GP_query_log(qryInsert, 'LD6I01', 'Q', 'Y');
                    End;
                sValue := sValue + tsokun[rCnt] + ',';
           End;
//운동소견Update
     sValue := '';
     For rCnt := 1 To 2 Do
      If uSokun[rCnt] <> '' Then
        With qryHul_Sokun Do
           Begin
              Close;
              ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
              ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
              ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
              ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
              ParamByName('In_Cod_sokun').AsString := uSokun[rCnt];
              ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
              ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
              ParamByName('In_Gubn_hsok').AsString := '3';
              Open;
              If RecordCount > 0 Then rCheck := 99
              Else                    rCheck := 00;
              Close;
              If rCheck = 00 Then
                 With qryInsert Do
                    Begin
                       Close;
                       ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
                       ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
                       ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
                       ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
                       ParamByName('In_Cod_sokun').AsString := uSokun[rCnt];
                       ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
                       ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;

                       With qrySokun Do
                       Begin
                          Close;
                          ParamByName('In_Cod_sokun').AsString := uSokun[rCnt];
                          Open;

                          qryInsert.ParamByName('In_Gubn_pan').AsString   := FieldByName('GUBN_PNJG').AsString;

                          qryInsert.ParamByName('In_Desc_sokun').AsString := FieldByName('Desc_sbsg').AsString  +
                                                                             FieldByName('Desc_sbsg1').AsString +
                                                                             FieldByName('Desc_sbsg2').AsString  +
                                                                             FieldByName('Desc_sbsg3').AsString;
                          Close;
                       End;

                       ParamByName('In_Gubn_hsok').AsString := '3';

                       if cmbDoctor.Text <> '' then ParamByname('In_Cod_Doct').AsString := Copy(cmbDoctor.Text,1,6)
                       else                         ParamByname('In_Cod_Doct').AsString := '150005';

                       ParamByName('In_Num_Seq').AsInteger   := rCnt;
                       ParamByName('In_Dat_Insert').AsString := GV_sToDay;
                       ParamByname('In_Cod_Insert').AsString := GV_sUserId;;
                       ParamByname('In_Dat_Update').AsString := '';
                       ParamByName('In_Cod_Update').AsString := '';
                       ExecSql;
                    End;
                sValue := sValue + uSokun[rCnt] + ',';
           End;
//식이소견Update
     For rCnt := 1 To 3 Do
      If sSokun[rCnt] <> '' Then
        With qryHul_Sokun Do
           Begin
              Close;
              ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
              ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
              ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
              ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
              ParamByName('In_Cod_sokun').AsString := sSokun[rCnt];
              ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
              ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;
              ParamByName('In_Gubn_hsok').AsString := '2';
              Open;
              If RecordCount > 0 Then rCheck := 99
              Else                    rCheck := 00;
              Close;
              If rCheck = 00 Then
                 With qryInsert Do
                    Begin
                       Close;
                       ParamByName('In_Cod_bjjs').AsString := qryBunju.FieldByName('Cod_bjjs').AsString;
                       ParamByName('In_Num_id').AsString   := qryBunju.FieldByName('Num_id').AsString;
                       ParamByName('In_Dat_Bunju').AsString := qryBunju.FieldByName('Dat_Bunju').AsString;
                       ParamByName('In_Num_Bunju').AsString := qryBunju.FieldByName('Num_bunju').AsString;
                       ParamByName('In_Cod_sokun').AsString := sSokun[rCnt];
                       ParamByName('In_Cod_jisa').AsString := qryBunju.FieldByName('Cod_jisa').AsString;
                       ParamByName('In_Dat_Gmgn').AsString := qryBunju.FieldByName('Dat_gmgn').AsString;


                       With qrySokun Do
                       Begin
                          Close;
                          ParamByName('In_Cod_sokun').AsString := sSokun[rCnt];
                          Open;

                          qryInsert.ParamByName('In_Gubn_pan').AsString   := FieldByName('GUBN_PNJG').AsString;

                          qryInsert.ParamByName('In_Desc_sokun').AsString := FieldByName('Desc_sbsg').AsString  +
                                                                             FieldByName('Desc_sbsg1').AsString +
                                                                             FieldByName('Desc_sbsg2').AsString  +
                                                                             FieldByName('Desc_sbsg3').AsString;
                          Close;
                       End;

                       ParamByName('In_Gubn_hsok').AsString := '2';

                       if cmbDoctor.Text <> '' then ParamByname('In_Cod_Doct').AsString := Copy(cmbDoctor.Text,1,6)
                       else                         ParamByname('In_Cod_Doct').AsString := '150005';

                       ParamByName('In_Num_Seq').AsInteger   := rCnt;
                       ParamByName('In_Dat_Insert').AsString := GV_sToDay;
                       ParamByname('In_Cod_Insert').AsString := GV_sUserId;
                       ParamByname('In_Dat_Update').AsString := '';
                       ParamByName('In_Cod_Update').AsString := '';
                       ExecSql;
                    End;
               sValue := sValue + sSokun[rCnt] + ',';
           End;
End;

procedure TfrmLD6I01.btnDateClick(Sender: TObject);
begin
  inherited;
  GF_CalendarClick(mskDate);
end;

Procedure TfrmLD6I01.Child_Sokun_Rtn;//14세이하
Var
   Val1, Val2, Val3, Val4, Val5, Val6 : Single;
   dCnt1, dCheck : Integer;
Begin

//82.Cod = HPB ((C010 >= 100) OR C010 >= 100) And (S007 = 양성 or S007 = 약양성)
      If (sGulkwa[09] <> '') And (sGulkwa[10] <> '') And (sGulkwa[70] <> '') Then
         Begin
            Val2 := StrToFloat(sGulkwa[10]);
            if (sGulkwa[55])<> '' then Val3 := StrToFloat(sGulkwa[55])
            else val3 :=0;

            If (Val2 >= 100) And ((sgulkwa[70] = '양성') or (sgulkwa[70] = '약양성')or (val3>=1.0)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HPB ';
            End;
         End;
//81.Cod = HBM (HPD조건에서 (S007 = 양성 or S007 = 약양성))
      If (tSokun[K] <> 'HPB ') Then
         If (sGulkwa[10] <> '') Then
            Begin
               Val2 := StrToFloat(sGulkwa[10]);
               if (sGulkwa[55])<> '' then Val3 := StrToFloat(sGulkwa[55])
               else val3 :=0;
               If ((Val2 >= 41) And (Val2 <= 99.0)) And ((sgulkwa[70] = '양성') or (sgulkwa[70] = '약양성')or (val3>=1.0)) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBM ';
                  End;
            End;

            //여기
       if  (sGulkwa[13] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[13]);
            if ((Val1 >= 150.1) and (sAge <=19)) Then
            begin
               K := K + 1;
               tSokun[K] := 'ALP1';
            end;
        end;

      If (sGulkwa[92] <> '') And (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            If (Val1 >= 30) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = ''))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               End
            else  If (Val1 >= 30) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 >= 1.401) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH1';
               End;
         End;
   If (sGulkwa[92] <> '') And (sGulkwa[89] <> '') And (sGulkwa[97] <> '') And (sGulkwa[30] <> '') And (sGulkwa[24] <> '') And (sGulkwa[37] <> '') And
      (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
         Begin
            Val1 := StrToFloat(sGulkwa[92]);
            if sGulkwa[89] <> '' then  Val2 := StrToFloat(sGulkwa[89])
            else Val2:=0;
            if  sGulkwa[97] <> '' then Val3 := StrToFloat(sGulkwa[97])
            else val3:=0;
            if sGulkwa[30]<> '' then  Val4 := StrToFloat(sGulkwa[30])
            else  val4:=0;
            if sGulkwa[24]<> '' then  Val5 := StrToFloat(sGulkwa[24])
            else  val5:=0;
            if sGulkwa[37]<> '' then  Val6 := StrToFloat(sGulkwa[37])
            else  val6:=0;
            Chk_Dang_Drug := '';         //통합문진 당뇨 확인
            Chk_Dang_Jindan := '';
            With QryTot_Munjin2018 Do
            Begin
               Close;
               ParamByName('Cod_jisa').Asstring := qryGumgin.FieldByName('Cod_jisa').AsString;
               ParamByName('Num_id').AsString   := qryGumgin.FieldByname('Num_id').Asstring;
               ParamByName('Dat_gmgn').Asstring := qryGumgin.FieldByname('Dat_gmgn').AsString;
               Open;
               If RecordCount > 0 Then
               Begin
                  Chk_Dang_Drug := FieldByname('Mun1_Drug4').AsString;
                  Chk_Dang_Jindan := FieldByname('Mun1_Jindan4').AsString;
               End;
               Close;
            End;

            If (Val1 >= 10) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0)) and ((Val4 =0) or (Val4<=1.40)) and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Chk_Dang_Drug <> '1') and (Chk_Dang_Jindan <> '1') and (Val5 <=125.9) and ((Val6 <= 6.4) OR (Val6 = 0))then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH ';
               End
            else  If (Val1 >= 10) and ((Val2<=10) or (Val2=0)) and ((Val3<=5) or (Val3=0))  and
               ((sGulkwa[90] = '음성') or (sGulkwa[90] = '')) and (Val4 >= 1.401) and
               ((Chk_Dang_Drug <> '1') And (Chk_Dang_Jindan <> '1'))and (Val5 <=125.9) and ((Val6 <= 6.5) OR (Val6 = 0)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'UPH1';
               End;
         End;

      If (sGulkwa[114] <> '') and (sGulkwa[113] = '')Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;


         If (sGulkwa[114] = '양성') and  (cSex = 1) and ((Val2 >=45) or ((Val2 <45) and (Val1 >=40.1)and (Val3 >=70.1)))  Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVG';
            End
         else If (sGulkwa[114] = '양성') and  (cSex = 2) and ((Val2 >=34) or ((Val2 <34) and (Val1 >=32.1)and (Val3 >=42.1))) Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVG';
            End


      End;
      If (sGulkwa[09]  <> '') and (sGulkwa[10]  <> '') and (sGulkwa[113] = '') Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
         If (Val2 <=45) and (Val1 <=40)  and (cSex = 1) and  (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
         else If (Val2 <=34) and (Val1 <=32)  and (cSex = 2) and  (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
         else If (Val2 <=45) and (Val1 >=40.1) and (Val3 <=70) and (cSex = 1) and    (sGulkwa[114] = '양성')  and
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End
          else If (Val2 <=34) and (Val1 >=32.1) and (Val3 <=42) and (cSex = 2) and   (sGulkwa[114] = '양성')  and 
            ((qryBunju.FieldByName('Dat_bunju').AsString)>= '20130518') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVH';
            End;
      End;
      If (sGulkwa[168] <> '') and (sGulkwa[113] = '')Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
          if  (sGulkwa[168]<> '')  Then  Val4 := StrToFloat(sGulkwa[168])
          else  Val4 :=0;


          If (Val4 >=20) and  (cSex = 1) and ((Val2 >=45.1) or ((Val2 <=45) and (Val1 >=40.1)and (Val3 >=70.1)))   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVL';
            End
         else If (Val4 >=20) and  (cSex = 2) and ((Val2 >=34.1) or ((Val2 <=34) and (Val1 >=32.1)and (Val3 >=42.1)))    Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVL';
            End

      End;


    If (sGulkwa[09]  <> '') and (sGulkwa[10]  <> '') and (sGulkwa[113] = '') Then
      Begin

          if  (sGulkwa[09] <> '') then  Val1 := StrToFloat(sGulkwa[09])
          else  Val1 :=0;
          if  (sGulkwa[10]<> '')  Then  Val2 := StrToFloat(sGulkwa[10])
          else  Val2 :=0;
          if  (sGulkwa[11]<> '')  Then  Val3 := StrToFloat(sGulkwa[11])
          else  Val3 :=0;
          if  (sGulkwa[168]<> '')  Then  Val4 := StrToFloat(sGulkwa[168])
          else  Val4 :=0;
         If (Val2 <=45) and (Val1 <=40)  and (cSex = 1) and  (val4>=20)  Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
         else If (Val2 <=34) and (Val1 <=32)  and (cSex = 2) and  (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
         else If (Val2 <=45) and (Val1 >=40.1) and (Val3 <=70) and (cSex = 1) and   (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End
          else If (Val2 <=34) and (Val1 >=32.1) and (Val3 <=42) and (cSex = 2) and   (val4>=20)   Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVO';
            End;
      End;

      if   (sGulkwa[168] <> '')  Then
      begin
           if  (sGulkwa[168]<> '')  Then  Val1 := StrToFloat(sGulkwa[168])
           else  Val1 :=0;

           If (Val1>=20) and (sGulkwa[113] = '양성')   Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV4';
           End
           else If (Val1>=20) and (sGulkwa[113] = '음성')  Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV ';
           End
           else If (Val1< 20) and ((sGulkwa[113] = '음성') or (sGulkwa[113] = ''))   Then
           Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
           End
      end;

 If  (sGulkwa[114] <> '')  Then
      Begin
         If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV ';
            End
         Else If (sGulkwa[113] = '양성') and
            (sGulkwa[114] = '양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV2';
            End
         Else If (sGulkwa[113] = '양성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV4';
            End
         Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
            End
           Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAV5';
            End
            Else If (sGulkwa[113] = '') and
            (sGulkwa[114] = '약양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVW';
            End
            Else If (sGulkwa[113] = '음성') and
            (sGulkwa[114] = '약양성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVC';
            End;
      End;                
      If (trim(sGulkwa[113]) = '') and (sGulkwa[114] <> '') then
         begin
          if (sGulkwa[114] = '음성') Then
            Begin
               K := K + 1;
               tSokun[K] := 'HAVB';
            End;
      end;


//79.Cod = HPM (C010 >= 100,S007 = '음성')
      If (tSokun[K] <> 'HPB ') And
         (tsokun[K] <> 'HBM ') Then
         If (sGulkwa[10] <> '') Then
            Begin
               if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
               Val2 := StrToFloat(sGulkwa[10]);
               If (Val2 >= 100) And ((sGulkwa[70] = '음성')or ((val1>0) and (val1<1.0))) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HPM ';
                  End;
            End;
//78.Cod = HPD (41 >= C010 <= 99,S007 = '음성')
      If (tSokun[K] <> 'HPB ') And
         (tSokun[K] <> 'HBM ') And
         (tsokun[K] <> 'HPM ') Then
         If (sGulkwa[10] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[10]);
               if (sGulkwa[55])<> '' then Val2 := StrToFloat(sGulkwa[55])
               else val2 :=0;
              If ((Val1 >= 41) And (Val1 <= 99.0)) And
                 ((sGulkwa[70] = '음성')or ((val2>0) and (val2<1.0))) then
                   Begin
                       K := K + 1;
                       tSokun[K] := 'HPD ';
                   End;
            End;
//84.Cod = HAB ((S007 = '양성' or S007 = '약양성') And S008 = 양성)
      If (sGulkwa[70] <> '') And ((sGulkwa[71] <> '') or (sGulkwa[115] <> '')) Then
         Begin
            if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
            else val1 :=0;
            If ((sgulkwa[70] = '양성') or (sgulkwa[70] = '약양성')or (val1>=1.0)) And ((sGulkwa[71] = '양성') or (strTofloat(sGulkwa[115]) >= 10)) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'HAB ';
               End;
         End;
//83.Cod = HBC (C010 <= 40 And (S007 = '양성' or S007 = '약양성')
      If (tSokun[K] <> 'HAB ') Then
         If (sGulkwa[10] <> '') And (sGulkwa[70] <> '') Then
            Begin
               if (sGulkwa[55])<> '' then Val1 := StrToFloat(sGulkwa[55])
               else val1 :=0;
               Val2 := StrToFloat(sGulkwa[10]);
               If ((sgulkwa[70] = '양성') or (sgulkwa[70] = '약양성')or (val1>=1.0)) And (Val2 <= 40.0) then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'HBC ';
                  End;
            End;
//순환기질환

//25.Cod = CH1 (C025 >= 241)
       If (sGulkwa[18] <> '')  and (sGulkwa[20] = '')  Then
       Begin
          Val1 := StrToFloat(sGulkwa[18]);
          If Val1 >= 240 Then
          Begin
             K := K + 1;
             tSokun[K] := 'CH1 ';
          End;
       End;
//24.Cod = CH (201 >= C025 <= 240)
      If (tSokun[K] <> 'CH1 ') Then
         If (sGulkwa[18] <> '')  and (sGulkwa[20] = '')Then
            Begin
               Val1 := StrToFloat(sGulkwa[18]);
               If (Val1 >= 200.1) And (Val1 <= 239.9) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'CH  ';
                  End;
            End;

// 당뇨질환
//30.Cod = DIA (C032 >= 121 And U005 >= 70)
//30.Cod = DIA (C032 >= 121 And U005 >= 250)--20110620 김희석
//30.Cod = DIB (C032 >= 126 And U005 <= 100)--DIA는 사용안함. 20111104 송재호
      If (sGulkwa[24] <> '') And (sGulkwa[93] <> '') Then
          Begin
              Val1 := StrToFloat(sGulkwa[24]);
              Val2 := StrToFloat(sGulkwa[93]);
              If (Val1 >= 126) And (Val2 >= 1) AND (Val2 <= 100) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'DIB ';
                  End;
          End;
//31.Cod = DMN1 (11 >= C032 <= 140)
      If (tSokun[K] <> 'DIB ') Then
          If (sGulkwa[24] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[24]);
                  If (Val1 >= 100) And (Val1 <= 140) Then
                      Begin
                         K := K + 1;
                         tSokun[K] := 'DMN1';
                      End;
              End;
//신장질환
{//14.Cod = BUN (C041 >= 30)
      If (tsokun[K] <> 'U03 ') Then
         If (sGulkwa[29] <> '') Then
            Begin
              Val1 := StrToFloat(sGulkwa[29]);
              If Val1 >= 30 Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'BUN ';
                  End;
            End;  }

//혈액질환
//74.Cod = ANE2 (H002 <= 10.9 And H008 <= 70 And H011 <= 3.1)
      If (sGulkwa[44] <> '') And (sGulkwa[50] <> '') And (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[44]);
            Val2 := StrToFloat(sGulkwa[50]);
            Val3 := StrToFloat(sGulkwa[53]);
            If (Val1 <= 8) And (Val2 <= 70) And (Val3 <= 3.1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'ANE2';
               End;
         End;
//8.Cod = ANE (H002 <= 10.9)
      If (tSokun[K] <> 'ANE2') Then
          If (sGulkwa[44] <> '') Then
              Begin
                 Val1 := StrToFloat(sGulkwa[44]);
                 If (Val1 <= 10.9) and  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004')  Then
                    Begin
                       K := K + 1;
                       tSokun[K] := 'ANE ';
                    End;
              End;


//40.Cod = PLH (H008 >= 600)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            If (Val1 >= 600) and  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH ';
               End;
         End;
//40.Cod = PLH1 (530>= H008 <= 599)
      If (sGulkwa[50] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[50]);
            If (Val1 <= 599) And (Val1 >= 530) And  ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End
            else If (Val1 >= 365.1) And (Val1 <= 400) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 1) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End
            else If (Val1 >= 398.1) And (Val1 <= 420) And ((qryBunju.FieldByName('Dat_Gmgn').AsString)>= '20121004') and (cSex = 2) Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'PLH1';
               End;

         End;
//70.Cod = WBC (H011 >= 14.5)
      If (sGulkwa[53] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[53]);
            If (Val1 >= 14.5) and ((qryBunju.FieldByName('Dat_Gmgn').AsString)< '20121004') Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'WBC ';
               End;
         End;

//45.Cod = RH (H024 이 음성)
      If (sGulkwa[66] <> '') Then
         Begin
            If (sGulkwa[66] = 'Rh-') Then
                Begin
                   K := K + 1;
                   tSokun[K] := 'RH ';
                End;
         End;

//소변관리
//65.Cod = URWP (U001 >= 50 And U004 >= 40 And U009 >= 30)
//65.Cod = URWP (U001 >= 25 And U004 >= 30 And U009 >= 10) 20110620 김희석
    {  If (sGulkwa[89] <> '') And (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            Val2 := StrToFloat(sGulkwa[92]);
            Val3 := StrToFloat(sGulkwa[97]);
            If (Val1 >= 25) And (Val2 >= 30) And (Val3 >= 10)Then
               Begin
                  K := K + 1;
                  tSokun[K] := 'URWP';
               End;
         End;  }
//59.Cod = URW (U001 >= 50 And U009 >= 30)
//59.Cod = URW (U001 >= 25 And U009 >= 10)--20110620 김희석
      If tsokun[K] <> 'URWP' Then
         If (sGulkwa[89] <> '') And (sGulkwa[97] <> '') Then
             Begin
                Val1 := StrToFloat(sGulkwa[89]);
                Val2 := StrToFloat(sGulkwa[97]);
                If (Val1 >= 25) And (Val2 >= 10) And
                   (sGulkwa[90] <> '양성')Then
                   Begin
                      K := K + 1;
                      //090409 철. 김소영선생님 요청.
                      if (sAge <= 50) and (cSex = 2) then
                         tSokun[K] := 'URW5'
                      else if (sAge > 50) and (cSex = 2) then
                         tSokun[K] := 'URW6'
                      else
                         tSokun[K] := 'URW ';
                   End
             End;
//58.Cod = UPB (U004 >= 40 And U009 >= 30)
//58.Cod = UPB (U004 >= 30 And U009 >= 10) --20110620 김희석
   {   If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') Then
         If (sGulkwa[92] <> '') And (sGulkwa[97] <> '') Then
            Begin
               Val1 := StrToFloat(sGulkwa[92]);
               Val2 := StrToFloat(sGulkwa[97]);
               If (Val1 >= 30) And (Val2 >= 10)Then
                   Begin
                      K := K + 1;
                      tSokun[K] := 'UPB ';
                   End;
            End;  }
//60.Cod = UWN (U001 >= 50 And U002 = '양성')
//60.Cod = UWN (U001 >= 25 And U002 = '양성') --20110620 김희석
    {  If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') Then
         If (sGulkwa[89] <> '') And (sGulkwa[90] <> '') Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            If (Val1 >= 25) And (sGulkwa[90] = '양성')Then
            Begin
               K := K + 1;
               if cSex = 2 then tSokun[K] := 'UWN1'
               else             tSokun[K] := 'UWN ';
            End;
         End;}
//56.Cod = UWB (U001 >= 50)
//56.Cod = UWB (U001 >= 25)--20110620  김희석
    {  If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'UWN1') Then
         If Sgulkwa[89] <> '' Then
         Begin
            Val1 := StrToFloat(sGulkwa[89]);
            If (Val1 >= 25)  Then
            Begin
               K := K + 1;
               if cSex = 2 then tSokun[K] := 'UWBC'
               else             tSokun[K] := 'UWB ';
            End
         End;  }
//62.Cod = URB1 (U009 >= 50)

      If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'URB ') Then
          If sGulkwa[97] <> '' Then
             Begin
                 Val1 := StrToFloat(sGulkwa[97]);
                 If (Val1 >= 50) Then
                     Begin
                         K := K + 1;
                         tSokun[K] := 'URB1';
                     End
             End;
//61.Cod = USH (C032 <= 110 And U005 >= 70)
//61.Cod = USH (C032 <= 110 And U005 >= 250)  --20110620 김희석
      If (tsokun[K] <> 'URWP') And
         (tsokun[K] <> 'URW ') And
         (tsokun[K] <> 'UPB ') And
         (tsokun[K] <> 'UWN ') And
         (tsokun[K] <> 'UWN1') And
         (tsokun[K] <> 'URB ') And
         (tsokun[K] <> 'URM ') Then
          If (sGulkwa[24] <> '') And (sGulkwa[93] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[24]);
                  Val2 := StrToFloat(sGulkwa[93]);
                  If (Val1 <= 110) And (Val2 >= 250) And (qryBunju.FieldByName('Dat_Bunju').AsString < '20140428') Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'USH ';
                      End;
                  If (Val1 <= 110) And (Val2 >= 100) And (qryBunju.FieldByName('Dat_Bunju').AsString >= '20140428') Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'USH ';
                      End;
              End;
//64.Cod = UPH (U004 >= 40)
//64.Cod = UPH (U004 >= 40) --20110620 김희석

    {      If (sGulkwa[92] <> '') Then
              Begin
                  Val1 := StrToFloat(sGulkwa[92]);
                  If (Val1 >= 30) Then
                      Begin
                          K := K + 1;
                          tSokun[K] := 'UPH ';
                      End;
              End;  }



//77.Cod = URB2 (U009 = 30)
//77.Cod = URB2 (U009 = 10)  --20110620 김희석
      dCheck := 0;
      For dCnt1 := 1 To 182 Do
          If (tSokun[dCnt1] = 'URW ') Or
             (tSokun[dCnt1] = 'UPB ') Or
             (tSokun[dCnt1] = 'URWP') Or
             (tSokun[dCnt1] = 'URW6') Then
             dCheck := dCheck + 1;
      If dCheck = 0 Then
         If sGulkwa[97] <> '' Then
            Begin
               Val1 := StrToFloat(sGulkwa[97]);
               If (Val1 = 10.0) Then
                  Begin
                     K := K + 1;
                     tSokun[K] := 'URB2';
                  End;
            End;
End;

procedure TfrmLD6I01.FormCreate(Sender: TObject);
begin
  inherited;
     if GV_sJisa = '00' then UV_sJisa := copy(GV_sUserId,1,2)
     else                    UV_sJisa := GV_sJisa;
     pnlCod_Dc.Visible := False;
     chkCod_Dc.Checked := False;
     If UV_sJisa = '15' Then chkCod_Dc.Visible := True;
     cmbDoctor.Clear;
      if  (GV_sUserId ='150005') or (GV_sUserId ='151026') or (GV_sUserId ='150043') or (GV_sUserId ='430064') then  //중국관련 소견시, 응급으로 인한 혈액결과 완료 임시 체크
         begin
              Ck_hulgulkwa.Visible := True;
         end;
     if  GV_sUserId ='150783' then  //송민경 선생님 요청
         begin
              GP_ComboSawon(cmbDoctor, 1);
              GP_ComboDisplay(cmbDoctor,GV_sUserId, 6);
         end;
     With  qryDoctor  do
        Begin
            Close;
            ParamByname('In_Cod_jisa').AsString :=  UV_sJisa;
            Open;
            If RecordCount > 0 then
               Begin
                   While  Not  Eof do
                      Begin
                         cmbDoctor.items.Add(FieldByname('cod_sawon').AsString + '-' +
                                             FieldByName('Desc_name').AsString);
                         Next;
                      End;
               End;
            Close;
        End;
end;

procedure TfrmLD6I01.btnDcClick(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtDc.Text  := sCode;
         edtDesc_dc.Text := sName;
      end;
   end;
end;

procedure TfrmLD6I01.chkCod_DcClick(Sender: TObject);
begin
  inherited;
  If chkCod_Dc.Checked Then
       pnlCod_Dc.Visible := True
  Else pnlCod_Dc.Visible := False;
end;

end.

