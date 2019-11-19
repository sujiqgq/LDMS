//==============================================================================
// 프로그램 설명 : 외주항목 결과 오류 List
// 수정일자      :
// 수정자        :
// 참고사항      :
//==============================================================================
// 프로그램 설명 : 외주항목 결과 오류 List
// 수정일자      : 2014.03.27
// 수정자        : 곽윤설
// 수정내용      : 외주항목 추가로 인한 항목추가 (T009)
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.07
// 수정자        : 곽윤설
// 수정내용      : SE13제거 & 외주항목 추가로 인한 항목추가 (T043)
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.04.26
// 수정자        : 곽윤설
// 수정내용      : 항목추가 (T040)
// 참고사항      :
//==============================================================================
// 수정일자      : 2014.05.12
// 수정자        : 곽윤설
// 수정내용      : 항목변경 (S082->S098)
// 참고사항      : [본원-최정미]
//==============================================================================
// 수정일자      : 2014.05.16
// 수정자        : 곽윤설
// 수정내용      : 항목추가 Y005,Y008
// 참고사항      : [본원-최정미]
//==============================================================================
// 수정일자      : 2014.06.21
// 수정자        : 곽윤설
// 수정내용      : SE12->C038 / C071->S094 / T025->S095 코드변경
// 참고사항      : [본원-최정미]
//==============================================================================

unit LD8Q09;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, TeEngine, Series,
  TeeProcs, Chart, ComObj, ValEdit, Gauges;

type
  TfrmLD8Q09 = class(TfrmSingle)
    pnlCond: TPanel;
    mskST_date: TDateEdit;
    btnGmgnF: TSpeedButton;
    mskEd_date: TDateEdit;
    btnGmgnT: TSpeedButton;
    Label10: TLabel;
    Label11: TLabel;
    cmbJisa: TComboBox;
    Label12: TLabel;
    Label1: TLabel;
    edtCod_dc: TEdit;
    edtDesc_Dc: TEdit;
    btnDc: TSpeedButton;
    qryHul: TQuery;
    qryGumgin: TQuery;
    Gauge1: TGauge;
    grdMaster: TtsGrid;
    qryCell: TQuery;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    qryNsSokun: TQuery;
    qrySokun: TQuery;
    Rd_Tot: TRadioButton;
    Label3: TLabel;
    qryPf_hm: TQuery;
    qryNo_hangmok: TQuery;
    Rd_Sawon: TRadioButton;
    Ck_Jong: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
              DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnQuitClick(Sender: TObject);
    procedure UP_Change(Sender: TObject);
    procedure SBut_ExcelClick(Sender: TObject);
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;

 private
    { Private declarations }

    UV_v001, UV_v002, UV_v003, UV_v004, UV_v005, UV_v006, UV_v007, UV_v008, UV_v009, UV_v010,
    UV_v011, UV_v012, UV_v013, UV_v014, UV_v015, UV_v016, UV_v017, UV_v018, UV_v019, UV_v020,
    UV_v021, UV_v022, UV_v023, UV_v024, UV_v025, UV_v026, UV_v027, UV_v028, UV_v029, UV_v030,
    UV_v031, UV_v032, UV_v033, UV_v034, UV_v035, UV_v036, UV_v037, UV_v038, UV_v039, UV_v040,
    UV_v041, UV_v042, UV_v043, UV_v044, UV_v045, UV_v046, UV_v047, UV_v048, UV_v049, UV_v050,
    UV_v051, UV_v052, UV_v053, UV_v054, UV_v055, UV_v056, UV_v057, UV_v058, UV_v059, UV_v060,
    UV_v061, UV_v062, UV_v063, UV_v064, UV_v065, UV_v066, UV_v067, UV_v068, UV_v069, UV_v070,
    UV_v071, UV_v072, UV_v073, UV_v074, UV_v075, UV_v076, UV_v077, UV_v078, UV_v079, UV_v080,
    UV_v081, UV_v082, UV_v083, UV_v084, UV_v085, UV_v086, UV_v087, UV_v088, UV_v089, UV_v090,
    UV_v091, UV_v092, UV_v093, UV_v094, UV_v095, UV_v096, UV_v097, UV_v098, UV_v099, UV_v100,
    UV_v101, UV_v102, UV_v103, UV_v104, UV_v105, UV_v106, UV_v107, UV_v108, UV_v109, UV_v110,
    UV_v111, UV_v112, UV_v113, UV_v114, UV_v115, UV_v116, UV_v117, UV_v118, UV_v119, UV_v120,
    UV_v121, UV_v122, UV_v123: Variant;

    procedure UP_VarMemSet(iLength : Integer);
    procedure UP_Clear(Temp : Integer);

    function  UF_Condition : Boolean;
    function  UF_Kicho_Urin(sValue : Integer) : String;
    function  UF_Sokun(sValue : String) : String;
  public
    { Public declarations }
  end;


var
  frmLD8Q09 : TfrmLD8Q09;
  UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_sUrine_Char,UV_sDat_gmgn : String;

implementation

uses ZZFunc, ZZComm, MdmsFunc;


{$R *.DFM}

procedure TfrmLD8Q09.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_v001 := VarArrayCreate([0, 0], varOleStr);
      UV_v002 := VarArrayCreate([0, 0], varOleStr);
      UV_v003 := VarArrayCreate([0, 0], varOleStr);
      UV_v004 := VarArrayCreate([0, 0], varOleStr);
      UV_v005 := VarArrayCreate([0, 0], varOleStr);
      UV_v006 := VarArrayCreate([0, 0], varOleStr);
      UV_v007 := VarArrayCreate([0, 0], varOleStr);
      UV_v008 := VarArrayCreate([0, 0], varOleStr);
      UV_v009 := VarArrayCreate([0, 0], varOleStr);
      UV_v010 := VarArrayCreate([0, 0], varOleStr);
      UV_v011 := VarArrayCreate([0, 0], varOleStr);
      UV_v012 := VarArrayCreate([0, 0], varOleStr);
      UV_v013 := VarArrayCreate([0, 0], varOleStr);
      UV_v014 := VarArrayCreate([0, 0], varOleStr);
      UV_v015 := VarArrayCreate([0, 0], varOleStr);
      UV_v016 := VarArrayCreate([0, 0], varOleStr);
      UV_v017 := VarArrayCreate([0, 0], varOleStr);
      UV_v018 := VarArrayCreate([0, 0], varOleStr);
      UV_v019 := VarArrayCreate([0, 0], varOleStr);
      UV_v020 := VarArrayCreate([0, 0], varOleStr);
      UV_v021 := VarArrayCreate([0, 0], varOleStr);
      UV_v022 := VarArrayCreate([0, 0], varOleStr);
      UV_v023 := VarArrayCreate([0, 0], varOleStr);
      UV_v024 := VarArrayCreate([0, 0], varOleStr);
      UV_v025 := VarArrayCreate([0, 0], varOleStr);
      UV_v026 := VarArrayCreate([0, 0], varOleStr);
      UV_v027 := VarArrayCreate([0, 0], varOleStr);
      UV_v028 := VarArrayCreate([0, 0], varOleStr);
      UV_v029 := VarArrayCreate([0, 0], varOleStr);
      UV_v030 := VarArrayCreate([0, 0], varOleStr);
      UV_v031 := VarArrayCreate([0, 0], varOleStr);
      UV_v032 := VarArrayCreate([0, 0], varOleStr);
      UV_v033 := VarArrayCreate([0, 0], varOleStr);
      UV_v034 := VarArrayCreate([0, 0], varOleStr);
      UV_v035 := VarArrayCreate([0, 0], varOleStr);
      UV_v036 := VarArrayCreate([0, 0], varOleStr);
      UV_v037 := VarArrayCreate([0, 0], varOleStr);
      UV_v038 := VarArrayCreate([0, 0], varOleStr);
      UV_v039 := VarArrayCreate([0, 0], varOleStr);
      UV_v040 := VarArrayCreate([0, 0], varOleStr);
      UV_v041 := VarArrayCreate([0, 0], varOleStr);
      UV_v042 := VarArrayCreate([0, 0], varOleStr);
      UV_v043 := VarArrayCreate([0, 0], varOleStr);
      UV_v044 := VarArrayCreate([0, 0], varOleStr);
      UV_v045 := VarArrayCreate([0, 0], varOleStr);
      UV_v046 := VarArrayCreate([0, 0], varOleStr);
      UV_v047 := VarArrayCreate([0, 0], varOleStr);
      UV_v048 := VarArrayCreate([0, 0], varOleStr);
      UV_v049 := VarArrayCreate([0, 0], varOleStr);
      UV_v050 := VarArrayCreate([0, 0], varOleStr);
      UV_v051 := VarArrayCreate([0, 0], varOleStr);
      UV_v052 := VarArrayCreate([0, 0], varOleStr);
      UV_v053 := VarArrayCreate([0, 0], varOleStr);
      UV_v054 := VarArrayCreate([0, 0], varOleStr);
      UV_v055 := VarArrayCreate([0, 0], varOleStr);
      UV_v056 := VarArrayCreate([0, 0], varOleStr);
      UV_v057 := VarArrayCreate([0, 0], varOleStr);
      UV_v058 := VarArrayCreate([0, 0], varOleStr);
      UV_v059 := VarArrayCreate([0, 0], varOleStr);
      UV_v060 := VarArrayCreate([0, 0], varOleStr);
      UV_v061 := VarArrayCreate([0, 0], varOleStr);
      UV_v062 := VarArrayCreate([0, 0], varOleStr);
      UV_v063 := VarArrayCreate([0, 0], varOleStr);
      UV_v064 := VarArrayCreate([0, 0], varOleStr);
      UV_v065 := VarArrayCreate([0, 0], varOleStr);
      UV_v066 := VarArrayCreate([0, 0], varOleStr);
      UV_v067 := VarArrayCreate([0, 0], varOleStr);
      UV_v068 := VarArrayCreate([0, 0], varOleStr);
      UV_v069 := VarArrayCreate([0, 0], varOleStr);
      UV_v070 := VarArrayCreate([0, 0], varOleStr);
      UV_v071 := VarArrayCreate([0, 0], varOleStr);
      UV_v072 := VarArrayCreate([0, 0], varOleStr);
      UV_v073 := VarArrayCreate([0, 0], varOleStr);
      UV_v074 := VarArrayCreate([0, 0], varOleStr);
      UV_v075 := VarArrayCreate([0, 0], varOleStr);
      UV_v076 := VarArrayCreate([0, 0], varOleStr);
      UV_v077 := VarArrayCreate([0, 0], varOleStr);
      UV_v078 := VarArrayCreate([0, 0], varOleStr);
      UV_v079 := VarArrayCreate([0, 0], varOleStr);
      UV_v080 := VarArrayCreate([0, 0], varOleStr);
      UV_v081 := VarArrayCreate([0, 0], varOleStr);
      UV_v082 := VarArrayCreate([0, 0], varOleStr);
      UV_v083 := VarArrayCreate([0, 0], varOleStr);
      UV_v084 := VarArrayCreate([0, 0], varOleStr);
      UV_v085 := VarArrayCreate([0, 0], varOleStr);
      UV_v086 := VarArrayCreate([0, 0], varOleStr);
      UV_v087 := VarArrayCreate([0, 0], varOleStr);
      UV_v088 := VarArrayCreate([0, 0], varOleStr);
      UV_v089 := VarArrayCreate([0, 0], varOleStr);
      UV_v090 := VarArrayCreate([0, 0], varOleStr);
      UV_v091 := VarArrayCreate([0, 0], varOleStr);
      UV_v092 := VarArrayCreate([0, 0], varOleStr);
      UV_v093 := VarArrayCreate([0, 0], varOleStr);
      UV_v094 := VarArrayCreate([0, 0], varOleStr);
      UV_v095 := VarArrayCreate([0, 0], varOleStr);
      UV_v096 := VarArrayCreate([0, 0], varOleStr);
      UV_v097 := VarArrayCreate([0, 0], varOleStr);
      UV_v098 := VarArrayCreate([0, 0], varOleStr);
      UV_v099 := VarArrayCreate([0, 0], varOleStr);
      UV_v100 := VarArrayCreate([0, 0], varOleStr);
      UV_v101 := VarArrayCreate([0, 0], varOleStr);
      UV_v102 := VarArrayCreate([0, 0], varOleStr);
      UV_v103 := VarArrayCreate([0, 0], varOleStr);
      UV_v104 := VarArrayCreate([0, 0], varOleStr);
      UV_v105 := VarArrayCreate([0, 0], varOleStr);
      UV_v106 := VarArrayCreate([0, 0], varOleStr);
      UV_v107 := VarArrayCreate([0, 0], varOleStr);
      UV_v108 := VarArrayCreate([0, 0], varOleStr);
      UV_v109 := VarArrayCreate([0, 0], varOleStr);
      UV_v110 := VarArrayCreate([0, 0], varOleStr);
      UV_v111 := VarArrayCreate([0, 0], varOleStr);
      UV_v112 := VarArrayCreate([0, 0], varOleStr);
      UV_v113 := VarArrayCreate([0, 0], varOleStr);
      UV_v114 := VarArrayCreate([0, 0], varOleStr);
      UV_v115 := VarArrayCreate([0, 0], varOleStr);
      UV_v116 := VarArrayCreate([0, 0], varOleStr);
      UV_v117 := VarArrayCreate([0, 0], varOleStr);
      UV_v118 := VarArrayCreate([0, 0], varOleStr);
      UV_v119 := VarArrayCreate([0, 0], varOleStr);
      UV_v120 := VarArrayCreate([0, 0], varOleStr);
      UV_v121 := VarArrayCreate([0, 0], varOleStr);
      UV_v122 := VarArrayCreate([0, 0], varOleStr);
      UV_v123 := VarArrayCreate([0, 0], varOleStr);

   end
   else
   begin
      VarArrayRedim(UV_v001, iLength);
      VarArrayRedim(UV_v002, iLength);
      VarArrayRedim(UV_v003, iLength);
      VarArrayRedim(UV_v004, iLength);
      VarArrayRedim(UV_v005, iLength);
      VarArrayRedim(UV_v006, iLength);
      VarArrayRedim(UV_v007, iLength);
      VarArrayRedim(UV_v008, iLength);
      VarArrayRedim(UV_v009, iLength);
      VarArrayRedim(UV_v010, iLength);
      VarArrayRedim(UV_v011, iLength);
      VarArrayRedim(UV_v012, iLength);
      VarArrayRedim(UV_v013, iLength);
      VarArrayRedim(UV_v014, iLength);
      VarArrayRedim(UV_v015, iLength);
      VarArrayRedim(UV_v016, iLength);
      VarArrayRedim(UV_v017, iLength);
      VarArrayRedim(UV_v018, iLength);
      VarArrayRedim(UV_v019, iLength);
      VarArrayRedim(UV_v020, iLength);
      VarArrayRedim(UV_v021, iLength);
      VarArrayRedim(UV_v022, iLength);
      VarArrayRedim(UV_v023, iLength);
      VarArrayRedim(UV_v024, iLength);
      VarArrayRedim(UV_v025, iLength);
      VarArrayRedim(UV_v026, iLength);
      VarArrayRedim(UV_v027, iLength);
      VarArrayRedim(UV_v028, iLength);
      VarArrayRedim(UV_v029, iLength);
      VarArrayRedim(UV_v030, iLength);
      VarArrayRedim(UV_v031, iLength);
      VarArrayRedim(UV_v032, iLength);
      VarArrayRedim(UV_v033, iLength);
      VarArrayRedim(UV_v034, iLength);
      VarArrayRedim(UV_v035, iLength);
      VarArrayRedim(UV_v036, iLength);
      VarArrayRedim(UV_v037, iLength);
      VarArrayRedim(UV_v038, iLength);
      VarArrayRedim(UV_v039, iLength);
      VarArrayRedim(UV_v040, iLength);
      VarArrayRedim(UV_v041, iLength);
      VarArrayRedim(UV_v042, iLength);
      VarArrayRedim(UV_v043, iLength);
      VarArrayRedim(UV_v044, iLength);
      VarArrayRedim(UV_v045, iLength);
      VarArrayRedim(UV_v046, iLength);
      VarArrayRedim(UV_v047, iLength);
      VarArrayRedim(UV_v048, iLength);
      VarArrayRedim(UV_v049, iLength);
      VarArrayRedim(UV_v050, iLength);
      VarArrayRedim(UV_v051, iLength);
      VarArrayRedim(UV_v052, iLength);
      VarArrayRedim(UV_v053, iLength);
      VarArrayRedim(UV_v054, iLength);
      VarArrayRedim(UV_v055, iLength);
      VarArrayRedim(UV_v056, iLength);
      VarArrayRedim(UV_v057, iLength);
      VarArrayRedim(UV_v058, iLength);
      VarArrayRedim(UV_v059, iLength);
      VarArrayRedim(UV_v060, iLength);
      VarArrayRedim(UV_v061, iLength);
      VarArrayRedim(UV_v062, iLength);
      VarArrayRedim(UV_v063, iLength);
      VarArrayRedim(UV_v064, iLength);
      VarArrayRedim(UV_v065, iLength);
      VarArrayRedim(UV_v066, iLength);
      VarArrayRedim(UV_v067, iLength);
      VarArrayRedim(UV_v068, iLength);
      VarArrayRedim(UV_v069, iLength);
      VarArrayRedim(UV_v070, iLength);
      VarArrayRedim(UV_v071, iLength);
      VarArrayRedim(UV_v072, iLength);
      VarArrayRedim(UV_v073, iLength);
      VarArrayRedim(UV_v074, iLength);
      VarArrayRedim(UV_v075, iLength);
      VarArrayRedim(UV_v076, iLength);
      VarArrayRedim(UV_v077, iLength);
      VarArrayRedim(UV_v078, iLength);
      VarArrayRedim(UV_v079, iLength);
      VarArrayRedim(UV_v080, iLength);
      VarArrayRedim(UV_v081, iLength);
      VarArrayRedim(UV_v082, iLength);
      VarArrayRedim(UV_v083, iLength);
      VarArrayRedim(UV_v084, iLength);
      VarArrayRedim(UV_v085, iLength);
      VarArrayRedim(UV_v086, iLength);
      VarArrayRedim(UV_v087, iLength);
      VarArrayRedim(UV_v088, iLength);
      VarArrayRedim(UV_v089, iLength);
      VarArrayRedim(UV_v090, iLength);
      VarArrayRedim(UV_v091, iLength);
      VarArrayRedim(UV_v092, iLength);
      VarArrayRedim(UV_v093, iLength);
      VarArrayRedim(UV_v094, iLength);
      VarArrayRedim(UV_v095, iLength);
      VarArrayRedim(UV_v096, iLength);
      VarArrayRedim(UV_v097, iLength);
      VarArrayRedim(UV_v098, iLength);
      VarArrayRedim(UV_v099, iLength);
      VarArrayRedim(UV_v100, iLength);
      VarArrayRedim(UV_v101, iLength);
      VarArrayRedim(UV_v102, iLength);
      VarArrayRedim(UV_v103, iLength);
      VarArrayRedim(UV_v104, iLength);
      VarArrayRedim(UV_v105, iLength);
      VarArrayRedim(UV_v106, iLength);
      VarArrayRedim(UV_v107, iLength);
      VarArrayRedim(UV_v108, iLength);
      VarArrayRedim(UV_v109, iLength);
      VarArrayRedim(UV_v110, iLength);
      VarArrayRedim(UV_v111, iLength);
      VarArrayRedim(UV_v112, iLength);
      VarArrayRedim(UV_v113, iLength);
      VarArrayRedim(UV_v114, iLength);
      VarArrayRedim(UV_v115, iLength);
      VarArrayRedim(UV_v116, iLength);
      VarArrayRedim(UV_v117, iLength);
      VarArrayRedim(UV_v118, iLength);
      VarArrayRedim(UV_v119, iLength);
      VarArrayRedim(UV_v120, iLength);
      VarArrayRedim(UV_v121, iLength);
      VarArrayRedim(UV_v122, iLength);
      VarArrayRedim(UV_v123, iLength);
   end;
end;

procedure TfrmLD8Q09.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;
   case DataCol of
      1 : Value := UV_v001[DataRow - 1];
      2 : Value := UV_v002[DataRow - 1];
      3 : Value := UV_v003[DataRow - 1];
      4 : Value := UV_v004[DataRow - 1];
      5 : Value := UV_v005[DataRow - 1];
      6 : Value := UV_v006[DataRow - 1];
      7 : Value := UV_v007[DataRow - 1];
      8 : Value := UV_v008[DataRow - 1];
      9 : Value := UV_v009[DataRow - 1];
     10 : Value := UV_v010[DataRow - 1];
     11 : Value := UV_v011[DataRow - 1];
     12 : Value := UV_v012[DataRow - 1];
     13 : Value := UV_v013[DataRow - 1];
     14 : Value := UV_v014[DataRow - 1];
     15 : Value := UV_v015[DataRow - 1];
     16 : Value := UV_v016[DataRow - 1];
     17 : Value := UV_v017[DataRow - 1];
     18 : Value := UV_v018[DataRow - 1];
     19 : Value := UV_v019[DataRow - 1];
     20 : Value := UV_v020[DataRow - 1];
     21 : Value := UV_v021[DataRow - 1];
     22 : Value := UV_v022[DataRow - 1];
     23 : Value := UV_v023[DataRow - 1];
     24 : Value := UV_v024[DataRow - 1];
     25 : Value := UV_v025[DataRow - 1];
     26 : Value := UV_v026[DataRow - 1];
     27 : Value := UV_v027[DataRow - 1];
     28 : Value := UV_v028[DataRow - 1];
     29 : Value := UV_v029[DataRow - 1];
     30 : Value := UV_v030[DataRow - 1];
     31 : Value := UV_v031[DataRow - 1];
     32 : Value := UV_v032[DataRow - 1];
     33 : Value := UV_v033[DataRow - 1];
     34 : Value := UV_v034[DataRow - 1];
     35 : Value := UV_v035[DataRow - 1];
     36 : Value := UV_v036[DataRow - 1];
     37 : Value := UV_v037[DataRow - 1];
     38 : Value := UV_v038[DataRow - 1];
     39 : Value := UV_v039[DataRow - 1];
     40 : Value := UV_v040[DataRow - 1];
     41 : Value := UV_v041[DataRow - 1];
     42 : Value := UV_v042[DataRow - 1];
     43 : Value := UV_v043[DataRow - 1];
     44: Value := UV_v044[DataRow - 1];
     45: Value := UV_v045[DataRow - 1];
     46: Value := UV_v046[DataRow - 1];
     47: Value := UV_v047[DataRow - 1];
     48: Value := UV_v048[DataRow - 1];
     49: Value := UV_v049[DataRow - 1];
     50: Value := UV_v050[DataRow - 1];
     51: Value := UV_v051[DataRow - 1];
     52: Value := UV_v052[DataRow - 1];
     53: Value := UV_v053[DataRow - 1];
     54: Value := UV_v054[DataRow - 1];
     55: Value := UV_v055[DataRow - 1];
     56: Value := UV_v056[DataRow - 1];
     57: Value := UV_v057[DataRow - 1];
     58: Value := UV_v058[DataRow - 1];
     59: Value := UV_v059[DataRow - 1];
     60: Value := UV_v060[DataRow - 1];
     61: Value := UV_v061[DataRow - 1];
     62: Value := UV_v062[DataRow - 1];
     63: Value := UV_v063[DataRow - 1];
     64: Value := UV_v064[DataRow - 1];
     65: Value := UV_v065[DataRow - 1];
     66: Value := UV_v066[DataRow - 1];
     67: Value := UV_v067[DataRow - 1];
     68: Value := UV_v068[DataRow - 1];
     69: Value := UV_v069[DataRow - 1];
     70: Value := UV_v070[DataRow - 1];
     71: Value := UV_v071[DataRow - 1];
     72: Value := UV_v072[DataRow - 1];
     73: Value := UV_v073[DataRow - 1];
     74: Value := UV_v074[DataRow - 1];
     75: Value := UV_v075[DataRow - 1];
     76: Value := UV_v076[DataRow - 1];
     77: Value := UV_v077[DataRow - 1];
     78: Value := UV_v078[DataRow - 1];
     79: Value := UV_v079[DataRow - 1];
     80: Value := UV_v080[DataRow - 1];
     81: Value := UV_v081[DataRow - 1];
     82: Value := UV_v082[DataRow - 1];
     83: Value := UV_v083[DataRow - 1];
     84: Value := UV_v084[DataRow - 1];
     85: Value := UV_v085[DataRow - 1];
     86: Value := UV_v086[DataRow - 1];
     87: Value := UV_v087[DataRow - 1];
     88: Value := UV_v088[DataRow - 1];
     89: Value := UV_v089[DataRow - 1];
     90: Value := UV_v090[DataRow - 1];
     91: Value := UV_v091[DataRow - 1];
     92: Value := UV_v092[DataRow - 1];
     93: Value := UV_v093[DataRow - 1];
     94: Value := UV_v094[DataRow - 1];
     95: Value := UV_v095[DataRow - 1];
     96: Value := UV_v096[DataRow - 1];
     97: Value := UV_v097[DataRow - 1];
     98: Value := UV_v098[DataRow - 1];
     99: Value := UV_v099[DataRow - 1];
     100: Value := UV_v100[DataRow - 1];
     101: Value := UV_v101[DataRow - 1];
     102: Value := UV_v102[DataRow - 1];
     103: Value := UV_v103[DataRow - 1];
     104: Value := UV_v104[DataRow - 1];
     105: Value := UV_v105[DataRow - 1];
     106: Value := UV_v106[DataRow - 1];
     107: Value := UV_v107[DataRow - 1];
     108: Value := UV_v108[DataRow - 1];

     109: Value := UV_v109[DataRow - 1];
     110: Value := UV_v110[DataRow - 1];
     111: Value := UV_v111[DataRow - 1];
     112: Value := UV_v112[DataRow - 1];
     113: Value := UV_v113[DataRow - 1];
     114: Value := UV_v114[DataRow - 1];
     115: Value := UV_v115[DataRow - 1];
     116: Value := UV_v116[DataRow - 1];
     117: Value := UV_v117[DataRow - 1];
     118: Value := UV_v118[DataRow - 1];
     119: Value := UV_v119[DataRow - 1];
     120: Value := UV_v120[DataRow - 1];
     121: Value := UV_v121[DataRow - 1];
     122: Value := UV_v122[DataRow - 1];
     123: Value := UV_v123[DataRow - 1];
   end;
end;

function  TfrmLD8Q09.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
         Result := Result + FieldByName('desc_jangbi').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD8Q09.FormCreate(Sender: TObject);
var  sJisa : string;
begin
   inherited;


   UP_VarMemSet(0);
   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      sJisa := '15';
   end
   else
   begin
      sJisa := GV_sJisa;
   end;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);

   edtCod_dc.text := '';


end;

procedure TfrmLD8Q09.btnQueryClick(Sender: TObject);
var
   sSex, sSelect, sWhere, sOrder,
   sJJ12, sCod_hm, cod_name, nSokun ,Temp_gulkwa : String;
   Index, iTemp, iRet, i, sokun_count  : Integer;
   vCod_chuga, vCod_hm : Variant;
   Check : boolean;
   Sum_vE021,  Sum_vS055,   Sum_vT043,   Sum_vU034,   Sum_vE016,   Sum_vT012,   Sum_vC081,   Sum_vT038,   Sum_vC070,   Sum_vY009,
   Sum_vE019,  Sum_vSE16,   Sum_vC038,   Sum_vS093,   Sum_vC020,   Sum_vC064,   Sum_vT021,   Sum_vT022,   Sum_vT023,   Sum_vT024,
   Sum_vS094,  Sum_vE047,   Sum_vS076,   Sum_vS077,   Sum_vS078,   Sum_vE012,   Sum_vS085,   Sum_vSE26,   Sum_vSE27,   Sum_vSE28,
   Sum_vSE29,  Sum_vS095,   Sum_vT026,   Sum_vT027,   Sum_vT028,   Sum_vT032,   Sum_vS054,   Sum_vSE38,   Sum_vT015,   Sum_vS053,
   Sum_vU049,  Sum_vC075,   Sum_vH031,   Sum_vC076,   Sum_vE050,   Sum_vU039,   Sum_vU040,   Sum_vU041,   Sum_vE051,   Sum_vE052,
   Sum_vE053,  Sum_vE054,   Sum_vE055,   Sum_vE056,   Sum_vE057,   Sum_vE058,   Sum_vT005,   Sum_vS022,   Sum_vE049,   Sum_vS019,
   Sum_vT033,  Sum_vT035,   Sum_vT036,   Sum_vH027,   Sum_vH028,   Sum_vE014,   Sum_vE060,   Sum_vE061,   Sum_vE062,   Sum_vE063,
   Sum_vE064,  Sum_vE065,   Sum_vT044,   Sum_vT045,   Sum_vT046,   Sum_vT047,   Sum_vT048,   Sum_vE066,   Sum_vE067,   Sum_vT049,
   Sum_vS086,  Sum_vS087,   Sum_vS088,   Sum_vS089,   Sum_vS064,   Sum_vC065,   Sum_vC072,   Sum_vC079,   Sum_vS059,   Sum_vS060,
   Sum_vS061,  Sum_vS062,   Sum_vS063,   Sum_vS070,   Sum_vS098,   Sum_vS030,   Sum_vS031,   Sum_vS032,   Sum_vT018,   Sum_vE013,
   Sum_vU048,  Sum_vS092,   Sum_vT050,   Sum_vC053,   Sum_vC055,   Sum_vS018,   Sum_vT039,   Sum_vU019,   Sum_vC018,   Sum_vE010,
   Sum_vS023,  Sum_vS029,   Sum_vS038,   Sum_vS040,   Sum_vC036,   Sum_vT009,   Sum_vT040,   Sum_vY005,   Sum_vY008   :integer;



begin
   inherited;
   sSelect := '';  sWhere := '';  sOrder := '';
   sSex := '';
   Sum_vE021:=0;  Sum_vS055:=0;   Sum_vT043:=0;   Sum_vU034:=0;  Sum_vE016:=0;  Sum_vT012:=0;   Sum_vC081:=0;
   Sum_vE019:=0;  Sum_vSE16:=0;   Sum_vC038:=0;   Sum_vS093:=0;  Sum_vC020:=0;  Sum_vC064:=0;   Sum_vT021:=0;
   Sum_vS094:=0;  Sum_vE047:=0;   Sum_vS076:=0;   Sum_vS077:=0;  Sum_vS078:=0;  Sum_vE012:=0;   Sum_vS085:=0;
   Sum_vSE29:=0;  Sum_vS095:=0;   Sum_vT026:=0;   Sum_vT027:=0;  Sum_vT028:=0;  Sum_vT032:=0;   Sum_vS054:=0;
   Sum_vU049:=0;  Sum_vC075:=0;   Sum_vH031:=0;   Sum_vC076:=0;  Sum_vE050:=0;  Sum_vU039:=0;   Sum_vU040:=0;
   Sum_vE053:=0;  Sum_vE054:=0;   Sum_vE055:=0;   Sum_vE056:=0;  Sum_vE057:=0;  Sum_vE058:=0;   Sum_vT005:=0;
   Sum_vT033:=0;  Sum_vT035:=0;   Sum_vT036:=0;   Sum_vH027:=0;  Sum_vH028:=0;  Sum_vE014:=0;   Sum_vE060:=0;
   Sum_vE064:=0;  Sum_vE065:=0;   Sum_vT044:=0;   Sum_vT045:=0;  Sum_vT046:=0;  Sum_vT047:=0;   Sum_vT048:=0;
   Sum_vS086:=0;  Sum_vS087:=0;   Sum_vS088:=0;   Sum_vS089:=0;  Sum_vS064:=0;  Sum_vC065:=0;   Sum_vC072:=0;
   Sum_vS061:=0;  Sum_vS062:=0;   Sum_vS063:=0;   Sum_vS070:=0;  Sum_vS098:=0;  Sum_vS030:=0;   Sum_vS031:=0;
   Sum_vU048:=0;  Sum_vS092:=0;   Sum_vT050:=0;   Sum_vC053:=0;  Sum_vC055:=0;  Sum_vS018:=0;   Sum_vT039:=0;
   Sum_vS023:=0;  Sum_vS029:=0;   Sum_vS038:=0;   Sum_vS040:=0;  Sum_vC036:=0;  Sum_vT009:=0;   Sum_vT040:=0;
   Sum_vT038:=0;  Sum_vC070:=0;   Sum_vY009:=0;   Sum_vY005:=0;  Sum_vY008:=0;
   Sum_vT022:=0;  Sum_vT023:=0;   Sum_vT024:=0;
   Sum_vSE26:=0;  Sum_vSE27:=0;   Sum_vSE28:=0;
   Sum_vSE38:=0;  Sum_vT015:=0;   Sum_vS053:=0;
   Sum_vU041:=0;  Sum_vE051:=0;   Sum_vE052:=0;
   Sum_vS022:=0;  Sum_vE049:=0;   Sum_vS019:=0;
   Sum_vE061:=0;  Sum_vE062:=0;   Sum_vE063:=0;
   Sum_vE066:=0;  Sum_vE067:=0;   Sum_vT049:=0;
   Sum_vC079:=0;  Sum_vS059:=0;   Sum_vS060:=0;
   Sum_vS032:=0;  Sum_vT018:=0;   Sum_vE013:=0;
   Sum_vU019:=0;  Sum_vC018:=0;   Sum_vE010:=0;



   // 조회조건 Check
   if not UF_Condition then exit;

   // Grid 초기화
   grdMaster.Rows := 0;

   // Grid 환경설정
   Index := 0;
   with qryGumgin do
   begin
      // SQL문을 생성후 자료를 Query
      Active := False;

      sSelect := ' SELECT G.cod_jisa, G.num_id, G.dat_gmgn,  G.desc_name, G.num_jumin, G.desc_dept, G.cod_hul, G.cod_chuga, ' +
                 '        B.dat_bunju, B.num_bunju, G.num_samp '  +
                 ' FROM bunju  B inner join gumgin  G  ON G.num_id = B.num_id AND G.dat_gmgn = B.dat_gmgn '  ;

      if Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) <> '00' then
         sWhere := ' WHERE G.Cod_jisa = ''' + Copy(cmbJisa.Items[cmbJisa.ItemIndex],1,2) + '''';

      if edtCod_dc.Text <> '' then
         sWhere := sWhere + '   AND G.cod_dc = ''' + edtCod_dc.Text + '''';

      sWhere := sWhere + '   AND B.dat_bunju >= ''' + mskST_date.Text + '''';
      sWhere := sWhere + '   AND B.dat_bunju <= ''' + mskEd_date.Text + '''';

      if Rd_Sawon.Checked then
         sWhere := sWhere + ' AND G.ysno_folk <> ''Y'' ';

      if Ck_Jong.Checked then
         sWhere := sWhere + ' AND G.cod_jangbi <> '''' and G.cod_hul <> '''' ';

      sOrder := ' Order BY B.Dat_bunju, B.num_bunju, G.num_samp, G.desc_name';

      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrder);

      Active := True;

      Gauge1.Progress := 0;
      if qryGumgin.RecordCount > 0 then
      begin
         GP_query_log(qryGumgin, 'LD8Q09', 'Q', 'N');
         Gauge1.MinValue := 0;
         Gauge1.MaxValue := RecordCount;
         UP_VarMemSet(qryGumgin.RecordCount-1);
         // DataSet의 값을 Variant변수로 이동
         while not qryGumgin.Eof do
         begin
            UP_Clear(Index);
            Gauge1.Progress := Gauge1.Progress + 1;
            //Label2.Caption := IntToStr(Gauge1.Progress) + ' / ' + IntToStr(Gauge1.MaxValue);
            Application.ProcessMessages;

            Check := false;
            //검사항목축출
            sCod_hm := '';
            if FieldByName('cod_hul').AsString <> '' then
            begin
               qryPf_hm.Active := False;
               qryPf_hm.ParamByName('In_Cod_Pf').AsString := FieldByName('cod_hul').AsString;
               qryPf_hm.Active := True;
               if qryPf_hm.RecordCount > 0 then
               begin
                  while not qryPf_hm.Eof do
                  begin
                     sCod_hm := sCod_hm + qryPf_hm.FieldByName('COD_HM').AsString;
                     qryPf_hm.Next;
                  end;
               end;
            end;

            iRet := GF_MulToSingle(FieldByName('cod_chuga').AsString, 4, vCod_chuga);
            for i := 0 to iRet-1 do
               sCod_hm := sCod_hm + vCod_chuga[i];

            iTemp := GF_MulToSingle(sCod_hm, 4, vCod_hm);

            case StrToInt(Copy(FieldByName('Num_jumin').AsString, 7, 1)) of
               1,3,5,7,9 : sSex := '남';
               2,4,6,8   : sSex := '여';
            end;

            // 혈액결과...
            qryHul.Active := False;
            qryHul.ParamByName('num_id').AsString   := qryGumgin.FieldByName('Num_id').AsString;
            qryHul.ParamByName('Dat_gmgn').AsString := qryGumgin.FieldByName('Dat_gmgn').AsString;
            qryHul.Active := True;
            if qryHul.RecordCount > 0 then
            begin
               while not qryHul.Eof do
               begin
                  UV_fGulkwa := '';
                  UV_fGulkwa1 := qryHul.FieldByName('DESC_GLKWA1').AsString;
                  UV_fGulkwa2 := qryHul.FieldByName('DESC_GLKWA2').AsString;
                  UV_fGulkwa3 := qryHul.FieldByName('DESC_GLKWA3').AsString;
                  GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                  Case qryHul.FieldByName('gubn_part').AsInteger of
                     1 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'C053') or (vCod_hm[iTemp] = 'C055') or (vCod_hm[iTemp] = 'C064') or
                                (vCod_hm[iTemp] = 'C065') or (vCod_hm[iTemp] = 'C070') or (vCod_hm[iTemp] = 'C072') or
                                (vCod_hm[iTemp] = 'C075') or (vCod_hm[iTemp] = 'C076') or (vCod_hm[iTemp] = 'C079') or
                                (vCod_hm[iTemp] = 'C081') then

                             if (vCod_hm[iTemp] = 'C053') and (Trim(Copy(UV_fGulkwa, 241,   6)) = '') then
                             begin
                                  UV_v104[Index] := '결과 무';
                                  Sum_vC053 := Sum_vC053  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C055') and (Trim(Copy(UV_fGulkwa, 271,   6)) = '') then
                             begin
                                  UV_v105[Index] := '결과 무';
                                  Sum_vC055 := Sum_vC055  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'C064') and (Trim(Copy(UV_fGulkwa, 265,   6)) = '') then
                             begin
                                  UV_v020[Index] := '결과 무';
                                  Sum_vC064 := Sum_vC064  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C065') and (Trim(Copy(UV_fGulkwa, 277,   6)) = '') then
                             begin
                                  UV_v046[Index] := '결과 무';
                                  Sum_vC065 := Sum_vC065  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C070') and (Trim(Copy(UV_fGulkwa, 313,   6)) = '') then
                             begin
                                  UV_v013[Index] := '결과 무';
                                  Sum_vC070 := Sum_vC070  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'C072') and (Trim(Copy(UV_fGulkwa, 319,   6)) = '') then
                             begin
                                  UV_v047[Index] := '결과 무';
                                  Sum_vC072 := Sum_vC072  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C075') and (Trim(Copy(UV_fGulkwa, 337,   6)) = '') then
                             begin
                                  UV_v022[Index] := '결과 무';
                                  Sum_vC075 := Sum_vC075  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C076') and (Trim(Copy(UV_fGulkwa, 343,   6)) = '') then
                             begin
                                  UV_v024[Index] := '결과 무';
                                  Sum_vC076 := Sum_vC076  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C079') and (Trim(Copy(UV_fGulkwa, 361,   6)) = '') then
                             begin
                                  UV_v048[Index] := '결과 무';
                                  Sum_vC079 := Sum_vC079  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C081') and (Trim(Copy(UV_fGulkwa, 373,   6)) = '') then
                             begin
                                  UV_v011[Index] := '결과 무';
                                  Sum_vC081 := Sum_vC081  + 1;
                                  check := True;
                             end;

                           {  if (vCod_hm[iTemp] = 'C053') and (Trim(Copy(UV_fGulkwa, 241,   6)) <> '') then UV_v104[Index] := '.';
                             if (vCod_hm[iTemp] = 'C055') and (Trim(Copy(UV_fGulkwa, 271,   6)) <> '') then UV_v105[Index] := '.';
                             if (vCod_hm[iTemp] = 'C064') and (Trim(Copy(UV_fGulkwa, 265,   6)) <> '') then UV_v020[Index] := '.';
                             if (vCod_hm[iTemp] = 'C065') and (Trim(Copy(UV_fGulkwa, 277,   6)) <> '') then UV_v046[Index] := '.';
                             if (vCod_hm[iTemp] = 'C070') and (Trim(Copy(UV_fGulkwa, 313,   6)) <> '') then UV_v013[Index] := '.';
                             if (vCod_hm[iTemp] = 'C072') and (Trim(Copy(UV_fGulkwa, 319,   6)) <> '') then UV_v047[Index] := '.';
                             if (vCod_hm[iTemp] = 'C075') and (Trim(Copy(UV_fGulkwa, 337,   6)) <> '') then UV_v022[Index] := '.';
                             if (vCod_hm[iTemp] = 'C076') and (Trim(Copy(UV_fGulkwa, 343,   6)) <> '') then UV_v024[Index] := '.';
                             if (vCod_hm[iTemp] = 'C079') and (Trim(Copy(UV_fGulkwa, 361,   6)) <> '') then UV_v048[Index] := '.';
                             if (vCod_hm[iTemp] = 'C081') and (Trim(Copy(UV_fGulkwa, 373,   6)) <> '') then UV_v011[Index] := '.';    }
                           end;
                         end;
                     2 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'H027') or (vCod_hm[iTemp] = 'H028') or (vCod_hm[iTemp] = 'H031') then

                             if (vCod_hm[iTemp] = 'H027') and (Trim(Copy(UV_fGulkwa, 151,   6)) = '') then
                             begin
                                  UV_v084[Index] := '결과 무';
                                  Sum_vH027 := Sum_vH027  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'H028') and (Trim(Copy(UV_fGulkwa, 157,   6)) = '') then
                             begin
                                  UV_v085[Index] := '결과 무';
                                  Sum_vH028 := Sum_vH028  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'H031') and (Trim(Copy(UV_fGulkwa, 169,   6)) = '') then
                             begin
                                  UV_v023[Index] := '결과 무';
                                  Sum_vH031 := Sum_vH031  + 1;
                                  check := True;
                             end;

                            { if (vCod_hm[iTemp] = 'H027') and (Trim(Copy(UV_fGulkwa, 151,   6)) <> '') then UV_v084[Index] := '.';
                             if (vCod_hm[iTemp] = 'H028') and (Trim(Copy(UV_fGulkwa, 157,   6)) <> '') then UV_v085[Index] := '.';
                             if (vCod_hm[iTemp] = 'H031') and (Trim(Copy(UV_fGulkwa, 169,   6)) <> '') then UV_v023[Index] := '.';  }
                           end;
                         end;
                     3 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'U019') or (vCod_hm[iTemp] = 'U034') or (vCod_hm[iTemp] = 'U039') or
                                (vCod_hm[iTemp] = 'U040') or (vCod_hm[iTemp] = 'U041') or (vCod_hm[iTemp] = 'U048') or
                                (vCod_hm[iTemp] = 'U049') or (vCod_hm[iTemp] = 'Y009') or (vCod_hm[iTemp] = 'Y005') or
                                (vCod_hm[iTemp] = 'Y008') then

                             if (vCod_hm[iTemp] = 'U019') and (Trim(Copy(UV_fGulkwa, 127,   6)) = '') then
                             begin
                                  UV_v108[Index] := '결과 무';
                                  Sum_vU019 := Sum_vU019  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'U034') and (Trim(Copy(UV_fGulkwa, 205,   6)) = '') then
                             begin
                                  UV_v008[Index] := '결과 무';
                                  Sum_vU034 := Sum_vU034  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'U039') and (Trim(Copy(UV_fGulkwa, 247,   6)) = '') then
                             begin
                                  UV_v026[Index] := '결과 무';
                                  Sum_vU039 := Sum_vU039  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'U040') and (Trim(Copy(UV_fGulkwa, 253,   6)) = '') then
                             begin
                                  UV_v027[Index] := '결과 무';
                                  Sum_vU040 := Sum_vU040  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'U041') and (Trim(Copy(UV_fGulkwa, 259,   6)) = '') then
                             begin
                                  UV_v028[Index] := '결과 무';
                                  Sum_vU040 := Sum_vU040  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'U048') and (Trim(Copy(UV_fGulkwa, 301,   6)) = '') then
                             begin
                                  UV_v101[Index] := '결과 무';
                                  Sum_vU048 := Sum_vU048  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'U049') and (Trim(Copy(UV_fGulkwa, 307,   6)) = '') then
                             begin
                                  UV_v021[Index] := '결과 무';
                                  Sum_vU049 := Sum_vU049  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'Y009') and (Trim(Copy(UV_fGulkwa, 241,   6)) = '') then
                             begin
                                  UV_v014[Index] := '결과 무';
                                  Sum_vY009 := Sum_vY009  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'Y005') and (Trim(Copy(UV_fGulkwa, 91,   6)) = '') then
                             begin
                                  UV_v122[Index] := '결과 무';
                                  Sum_vY005 := Sum_vY005  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'Y008') and (Trim(Copy(UV_fGulkwa, 235,   6)) = '') then
                             begin
                                  UV_v123[Index] := '결과 무';
                                  Sum_vY008 := Sum_vY008  + 1;
                                  check := True;
                             end;

                           {  if (vCod_hm[iTemp] = 'U019') and (Trim(Copy(UV_fGulkwa, 127,   6)) <> '') then UV_v108[Index] := '.';
                             if (vCod_hm[iTemp] = 'U034') and (Trim(Copy(UV_fGulkwa, 205,   6)) <> '') then UV_v008[Index] := '.';
                             if (vCod_hm[iTemp] = 'U039') and (Trim(Copy(UV_fGulkwa, 247,   6)) <> '') then UV_v026[Index] := '.';
                             if (vCod_hm[iTemp] = 'U040') and (Trim(Copy(UV_fGulkwa, 253,   6)) <> '') then UV_v027[Index] := '.';
                             if (vCod_hm[iTemp] = 'U041') and (Trim(Copy(UV_fGulkwa, 259,   6)) <> '') then UV_v028[Index] := '.';
                             if (vCod_hm[iTemp] = 'U048') and (Trim(Copy(UV_fGulkwa, 301,   6)) <> '') then UV_v101[Index] := '.';
                             if (vCod_hm[iTemp] = 'U049') and (Trim(Copy(UV_fGulkwa, 307,   6)) <> '') then UV_v021[Index] := '.';
                             if (vCod_hm[iTemp] = 'Y009') and (Trim(Copy(UV_fGulkwa, 241,   6)) <> '') then UV_v014[Index] := '.'; }
                           end;                                                                 
                         end;                                                                   
                     4 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'E019') or (vCod_hm[iTemp] = 'E049') or (vCod_hm[iTemp] = 'E050') or
                                (vCod_hm[iTemp] = 'S030') or (vCod_hm[iTemp] = 'S031') or (vCod_hm[iTemp] = 'S032') or
                                (vCod_hm[iTemp] = 'T005') or (vCod_hm[iTemp] = 'T018') or
                                (vCod_hm[iTemp] = 'T021') or (vCod_hm[iTemp] = 'T022') or (vCod_hm[iTemp] = 'T023') or
                                (vCod_hm[iTemp] = 'T024') or (vCod_hm[iTemp] = 'T044') or (vCod_hm[iTemp] = 'T045') or
                                (vCod_hm[iTemp] = 'T046') or (vCod_hm[iTemp] = 'T047') or (vCod_hm[iTemp] = 'T048') or
                                (vCod_hm[iTemp] = 'T049') or (vCod_hm[iTemp] = 'T050') then

                             if (vCod_hm[iTemp] = 'E019') and (Trim(Copy(UV_fGulkwa, 109,   6)) = '') then
                             begin
                                  UV_v015[Index] := '결과 무';
                                  Sum_vE019 := Sum_vE019  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E049') and (Trim(Copy(UV_fGulkwa, 163,   6)) = '') then
                             begin
                                  UV_v039[Index] := '결과 무';
                                  Sum_vE049 := Sum_vE049  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E050') and (Trim(Copy(UV_fGulkwa, 169,   6)) = '') then
                             begin
                                  UV_v025[Index] := '결과 무';
                                  Sum_vE050 := Sum_vE050  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S030') and (Trim(Copy(UV_fGulkwa,  31,   6)) = '') then
                             begin
                                  UV_v056[Index] := '결과 무';
                                  Sum_vS030 := Sum_vS030  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S031') and (Trim(Copy(UV_fGulkwa,  43,   6)) = '') then
                             begin
                                  UV_v057[Index] := '결과 무';
                                  Sum_vS031 := Sum_vS031  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S032') and (Trim(Copy(UV_fGulkwa,  37,   6)) = '') then
                             begin
                                  UV_v058[Index] := '결과 무';
                                  Sum_vS032 := Sum_vS032  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T005') and (Trim(Copy(UV_fGulkwa,  55,   6)) = '') then
                             begin
                                  UV_v037[Index] := '결과 무';
                                  Sum_vT005 := Sum_vT005  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T018') and (Trim(Copy(UV_fGulkwa,  85,   6)) = '') then
                             begin
                                  UV_v059[Index] := '결과 무';
                                  Sum_vT005 := Sum_vT005  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'T021') and (Trim(Copy(UV_fGulkwa,  97,   6)) = '') then
                             begin
                                  UV_v061[Index] := '결과 무';
                                  Sum_vT021 := Sum_vT021  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'T022') and (Trim(Copy(UV_fGulkwa, 103,   6)) = '') then
                             begin
                                  UV_v062[Index] := '결과 무';
                                  Sum_vT022 := Sum_vT022  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'T023') and (Trim(Copy(UV_fGulkwa, 127,   6)) = '') then
                             begin
                                  UV_v063[Index] := '결과 무';
                                  Sum_vT023:= Sum_vT023 + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T024') and (Trim(Copy(UV_fGulkwa, 133,   6)) = '') then
                             begin
                                  UV_v064[Index] := '결과 무';
                                  Sum_vT024 := Sum_vT024  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T044') and (Trim(Copy(UV_fGulkwa, 181,   6)) = '') then
                             begin
                                  UV_v093[Index] := '결과 무';
                                  Sum_vT044 := Sum_vT044  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T045') and (Trim(Copy(UV_fGulkwa, 187,   6)) = '') then
                             begin
                                  UV_v094[Index] := '결과 무';
                                  Sum_vT045 := Sum_vT045  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T046') and (Trim(Copy(UV_fGulkwa, 193,   6)) = '') then
                             begin
                                  UV_v095[Index] := '결과 무';
                                  Sum_vT046 := Sum_vT046  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T047') and (Trim(Copy(UV_fGulkwa, 199,   6)) = '') then
                             begin
                                  UV_v096[Index] := '결과 무';
                                  Sum_vT021 := Sum_vT021  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T048') and (Trim(Copy(UV_fGulkwa, 205,   6)) = '') then
                             begin
                                  UV_v097[Index] := '결과 무';
                                  Sum_vT048 := Sum_vT048  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'T049') and (Trim(Copy(UV_fGulkwa, 211,   6)) = '') then
                             begin
                                  UV_v100[Index] := '결과 무';
                                  Sum_vT049 := Sum_vT049  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T050') and (Trim(Copy(UV_fGulkwa, 217,   6)) = '') then
                             begin
                                  UV_v103[Index] := '결과 무';
                                  Sum_vT050 := Sum_vT050  + 1;
                                  check := True;
                             end;
                           end;
                         end;
                     5 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'C020') or (vCod_hm[iTemp] = 'E047') or
                                (vCod_hm[iTemp] = 'E051') or (vCod_hm[iTemp] = 'E052') or (vCod_hm[iTemp] = 'E053') or
                                (vCod_hm[iTemp] = 'E054') or (vCod_hm[iTemp] = 'E055') or (vCod_hm[iTemp] = 'E056') or
                                (vCod_hm[iTemp] = 'E057') or (vCod_hm[iTemp] = 'E058') or (vCod_hm[iTemp] = 'E060') or
                                (vCod_hm[iTemp] = 'E061') or (vCod_hm[iTemp] = 'E062') or (vCod_hm[iTemp] = 'E063') or
                                (vCod_hm[iTemp] = 'E064') or (vCod_hm[iTemp] = 'E065') or (vCod_hm[iTemp] = 'E066') or
                                (vCod_hm[iTemp] = 'E067') or (vCod_hm[iTemp] = 'S018') or (vCod_hm[iTemp] = 'S019') or
                                (vCod_hm[iTemp] = 'S022') or (vCod_hm[iTemp] = 'S076') or (vCod_hm[iTemp] = 'S077') or
                                (vCod_hm[iTemp] = 'S078') or (vCod_hm[iTemp] = 'S085') or (vCod_hm[iTemp] = 'S086') or
                                (vCod_hm[iTemp] = 'S087') or (vCod_hm[iTemp] = 'S088') or (vCod_hm[iTemp] = 'S089') or
                                (vCod_hm[iTemp] = 'SE26') or (vCod_hm[iTemp] = 'SE27') or (vCod_hm[iTemp] = 'SE28') or
                                (vCod_hm[iTemp] = 'SE29') or (vCod_hm[iTemp] = 'T026') or
                                (vCod_hm[iTemp] = 'T027') or (vCod_hm[iTemp] = 'T028') or (vCod_hm[iTemp] = 'T032') or
                                (vCod_hm[iTemp] = 'T033') or (vCod_hm[iTemp] = 'T035') or (vCod_hm[iTemp] = 'T036') or
                                (vCod_hm[iTemp] = 'T038') or (vCod_hm[iTemp] = 'T039') or (vCod_hm[iTemp] = 'T009') or
                                (vCod_hm[iTemp] = 'T043') or (vCod_hm[iTemp] = 'T040')then

                             if (vCod_hm[iTemp] = 'C020') and (Trim(Copy(UV_fGulkwa,  67,   6)) = '') then
                             begin
                                  UV_v019[Index] := '결과 무';
                                  Sum_vC020 := Sum_vC020  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E047') and (Trim(Copy(UV_fGulkwa, 139,   6)) = '') then
                             begin
                                  UV_v066[Index] := '결과 무';
                                  Sum_vE047 := Sum_vE047  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'E051') and (Trim(Copy(UV_fGulkwa, 259,   6)) = '') then
                             begin
                                  UV_v029[Index] := '결과 무';
                                  Sum_vE051 := Sum_vE051  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E052') and (Trim(Copy(UV_fGulkwa, 265,   6)) = '') then
                             begin
                                  UV_v030[Index] := '결과 무';
                                  Sum_vE052 := Sum_vE052  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E053') and (Trim(Copy(UV_fGulkwa, 271,   6)) = '') then
                             begin
                                  UV_v031[Index] := '결과 무';
                                  Sum_vE053 := Sum_vE053  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'E054') and (Trim(Copy(UV_fGulkwa, 277,   6)) = '') then
                             begin
                                  UV_v032[Index] := '결과 무';
                                  Sum_vE054 := Sum_vE054  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E055') and (Trim(Copy(UV_fGulkwa, 283,   6)) = '') then
                             begin
                                  UV_v033[Index] := '결과 무';
                                  Sum_vE055 := Sum_vE055  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E056') and (Trim(Copy(UV_fGulkwa, 289,   6)) = '') then
                             begin
                                  UV_v034[Index] := '결과 무';
                                  Sum_vE056 := Sum_vE056  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E057') and (Trim(Copy(UV_fGulkwa, 295,   6)) = '') then
                             begin
                                  UV_v035[Index] := '결과 무';
                                  Sum_vE057 := Sum_vE057  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E058') and (Trim(Copy(UV_fGulkwa, 301,   6)) = '') then
                             begin
                                  UV_v036[Index] := '결과 무';
                                  Sum_vE058:= Sum_vE058+ 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E060') and (Trim(Copy(UV_fGulkwa, 379,   6)) = '') then
                             begin
                                  UV_v087[Index] := '결과 무';
                                  Sum_vE060 := Sum_vE060  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E061') and (Trim(Copy(UV_fGulkwa, 385,   6)) = '') then
                             begin
                                  UV_v088[Index] := '결과 무';
                                  Sum_vE061 := Sum_vE061  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E062') and (Trim(Copy(UV_fGulkwa, 391,   6)) = '') then
                             begin
                                  UV_v089[Index] := '결과 무';
                                  Sum_vE062 := Sum_vE062  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E063') and (Trim(Copy(UV_fGulkwa, 397,   6)) = '') then
                             begin
                                  UV_v090[Index] := '결과 무';
                                  Sum_vE063 := Sum_vE063  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E064') and (Trim(Copy(UV_fGulkwa, 403,   6)) = '') then
                             begin
                                  UV_v091[Index] := '결과 무';
                                  Sum_vE064 := Sum_vE064  + 1;
                                  check := True;
                             end;

                             if (vCod_hm[iTemp] = 'E065') and (Trim(Copy(UV_fGulkwa, 409,   6)) = '') then
                             begin
                                  UV_v092[Index] := '결과 무';
                                  Sum_vE065 := Sum_vE065  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E066') and (Trim(Copy(UV_fGulkwa, 415,   6)) = '') then
                             begin
                                  UV_v098[Index] := '결과 무';
                                  Sum_vE066 := Sum_vE066  + 1;
                                  check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E067') and (Trim(Copy(UV_fGulkwa, 421,   6)) = '') then
                             begin
                             UV_v099[Index] := '결과 무';
                             Sum_vE067 := Sum_vE067  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S018') and (Trim(Copy(UV_fGulkwa,  73,   6)) = '') then
                             begin
                             UV_v106[Index] := '결과 무';
                             Sum_vS018 := Sum_vS018  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S019') and (Trim(Copy(UV_fGulkwa,  79,   6)) = '') then
                             begin
                             UV_v040[Index] := '결과 무';
                             Sum_vS019 := Sum_vS019  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S022') and (Trim(Copy(UV_fGulkwa,  91,   6)) = '') then
                             begin
                             UV_v038[Index] := '결과 무';
                             Sum_vS022 := Sum_vS022  + 1;
                             check := True;
                             end;

                             if (vCod_hm[iTemp] = 'S076') and (Trim(Copy(UV_fGulkwa,  55,   6)) = '') then
                             begin
                             UV_v067[Index] := '결과 무';
                             Sum_vS076 := Sum_vS076  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S077') and (Trim(Copy(UV_fGulkwa, 121,   6)) = '') then
                             begin
                             UV_v068[Index] := '결과 무';
                             Sum_vS077 := Sum_vS077  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S078') and (Trim(Copy(UV_fGulkwa, 157,   6)) = '') then
                             begin
                             UV_v069[Index] := '결과 무';
                             Sum_vS078 := Sum_vS078  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S085') and (Trim(Copy(UV_fGulkwa, 175,   6)) = '') then
                             begin
                             UV_v071[Index] := '결과 무';
                             Sum_vS085 := Sum_vS085  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S086') and (Trim(Copy(UV_fGulkwa, 319,   6)) = '') then
                             begin
                             UV_v041[Index] := '결과 무';
                             Sum_vS086 := Sum_vS086  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S087') and (Trim(Copy(UV_fGulkwa, 325,   6)) = '') then
                             begin
                             UV_v042[Index] := '결과 무';
                             Sum_vS087 := Sum_vS087  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S088') and (Trim(Copy(UV_fGulkwa, 331,   6)) = '') then
                             begin
                             UV_v043[Index] := '결과 무';
                             Sum_vS088 := Sum_vS088  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S089') and (Trim(Copy(UV_fGulkwa, 337,   6)) = '') then
                             begin
                             UV_v044[Index] := '결과 무';
                             Sum_vS089 := Sum_vS089  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE26') and (Trim(Copy(UV_fGulkwa,  97,   6)) = '') then
                             begin
                             UV_v072[Index] := '결과 무';
                             Sum_vSE26 := Sum_vSE26  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE27') and (Trim(Copy(UV_fGulkwa, 145,   6)) = '') then
                             begin
                             UV_v073[Index] := '결과 무';
                             Sum_vSE27 := Sum_vSE27  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE28') and (Trim(Copy(UV_fGulkwa, 109,   6)) = '') then
                             begin
                             UV_v074[Index] := '결과 무';
                             Sum_vSE28 := Sum_vSE28  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE29') and (Trim(Copy(UV_fGulkwa, 151,   6)) = '') then
                             begin
                             UV_v075[Index] := '결과 무';
                             Sum_vSE29 := Sum_vSE29  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T026') and (Trim(Copy(UV_fGulkwa, 181,   6)) = '') then
                             begin
                             UV_v077[Index] := '결과 무';
                             Sum_vT026 := Sum_vT026  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T027') and (Trim(Copy(UV_fGulkwa, 187,   6)) = '') then
                             begin
                             UV_v078[Index] := '결과 무';
                             Sum_vT027 := Sum_vT027  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T028') and (Trim(Copy(UV_fGulkwa, 193,   6)) = '') then
                             begin
                             UV_v079[Index] := '결과 무';
                             Sum_vT028 := Sum_vT028  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T032') and (Trim(Copy(UV_fGulkwa, 217,   6)) = '') then
                             begin
                             UV_v080[Index] := '결과 무';
                             Sum_vT032 := Sum_vT032  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T033') and (Trim(Copy(UV_fGulkwa, 223,   6)) = '') then
                             begin
                             UV_v081[Index] := '결과 무';
                             Sum_vT033 := Sum_vT033  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T035') and (Trim(Copy(UV_fGulkwa, 235,   6)) = '') then
                             begin
                             UV_v082[Index] := '결과 무';
                             Sum_vT035 := Sum_vT035  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T036') and (Trim(Copy(UV_fGulkwa, 241,   6)) = '') then
                             begin
                             UV_v083[Index] := '결과 무';
                             Sum_vT036 := Sum_vT036  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T038') and (Trim(Copy(UV_fGulkwa, 307,   6)) = '') then
                             begin
                             UV_v012[Index] := '결과 무';
                             Sum_vT038 := Sum_vT038  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T039') and (Trim(Copy(UV_fGulkwa, 313,   6)) = '') then
                             begin
                             UV_v107[Index] := '결과 무';
                             Sum_vT039 := Sum_vT039  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T009') and (Trim(Copy(UV_fGulkwa,  49,   6)) = '') then
                             begin
                             UV_v120[Index] := '결과 무';
                             Sum_vT009 := Sum_vT009  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T043') and (Trim(Copy(UV_fGulkwa,  361, 366)) = '') then
                             begin
                             UV_v007[Index] := '결과 무';
                             Sum_vT043 := Sum_vT043  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T040') and (Trim(Copy(UV_fGulkwa,  343, 348)) = '') then
                             begin
                             UV_v121[Index] := '결과 무';
                             Sum_vT040 := Sum_vT040  + 1;
                             check := True;
                             end;
                           end;
                         end;
                     7 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'S059') or (vCod_hm[iTemp] = 'S060') or (vCod_hm[iTemp] = 'S061') or
                                (vCod_hm[iTemp] = 'S062') or (vCod_hm[iTemp] = 'S063') or (vCod_hm[iTemp] = 'S064') or
                                (vCod_hm[iTemp] = 'S070') or (vCod_hm[iTemp] = 'S098') or (vCod_hm[iTemp] = 'S092') or
                                (vCod_hm[iTemp] = 'S093') or (vCod_hm[iTemp] = 'S094') or (vCod_hm[iTemp] = 'S095') then

                             if (vCod_hm[iTemp] = 'S059') and (Trim(Copy(UV_fGulkwa,  67,   6)) = '') then
                             begin
                             UV_v049[Index] := '결과 무';
                             Sum_vS059 := Sum_vS059  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S060') and (Trim(Copy(UV_fGulkwa,  73,   6)) = '') then
                              begin
                             UV_v050[Index] := '결과 무';
                             Sum_vS060 := Sum_vS060  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S061') and (Trim(Copy(UV_fGulkwa,  79,   6)) = '') then
                             begin
                             UV_v051[Index] := '결과 무';
                             Sum_vS061 := Sum_vS061  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S062') and (Trim(Copy(UV_fGulkwa, 157,   6)) = '') then
                             begin
                             UV_v052[Index] := '결과 무';
                             Sum_vS062 := Sum_vS062  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S063') and (Trim(Copy(UV_fGulkwa, 163,   6)) = '') then
                             begin
                             UV_v053[Index] := '결과 무';
                             Sum_vS063 := Sum_vS063  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S064') and (Trim(Copy(UV_fGulkwa, 175,   6)) = '') then
                             begin
                             UV_v045[Index] := '결과 무';
                             Sum_vS064 := Sum_vS064  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S070') and (Trim(Copy(UV_fGulkwa, 121,   6)) = '') then
                             begin
                             UV_v054[Index] := '결과 무';
                             Sum_vS070 := Sum_vS070  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S098') and (Trim(Copy(UV_fGulkwa, 223,   6)) = '') then
                             begin
                             UV_v055[Index] := '결과 무';
                             Sum_vS098 := Sum_vS098  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S092') and (Trim(Copy(UV_fGulkwa, 187,   6)) = '') then
                             begin
                             UV_v102[Index] := '결과 무';
                             Sum_vS092 := Sum_vS092  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S093') and (Trim(Copy(UV_fGulkwa, 193,   6)) = '') then
                             begin
                             UV_v018[Index] := '결과 무';
                             Sum_vS093 := Sum_vS093  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S094') and (Trim(Copy(UV_fGulkwa, 199,   6)) = '') then
                             begin
                             UV_v065[Index] := '결과 무';
                             Sum_vS094 := Sum_vS094  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S095') and (Trim(Copy(UV_fGulkwa, 205,   6)) = '') then
                             begin
                             UV_v076[Index] := '결과 무';
                             Sum_vS095 := Sum_vS095  + 1;
                             check := True;
                             end;
                           end;
                         end;
                     8 : begin
                           for iTemp := 0 to iTemp -1 do
                           begin
                             if (vCod_hm[iTemp] = 'E012') or (vCod_hm[iTemp] = 'E013') or (vCod_hm[iTemp] = 'E014') or
                                (vCod_hm[iTemp] = 'E016') or (vCod_hm[iTemp] = 'E021') or (vCod_hm[iTemp] = 'S055') or
                                (vCod_hm[iTemp] = 'SE16') or (vCod_hm[iTemp] = 'T012') or (vCod_hm[iTemp] = 'C018') or
                                (vCod_hm[iTemp] = 'SE16') or (vCod_hm[iTemp] = 'T012') or (vCod_hm[iTemp] = 'C018') or
                                (vCod_hm[iTemp] = 'E010') or (vCod_hm[iTemp] = 'S023') or (vCod_hm[iTemp] = 'S029') or
                                (vCod_hm[iTemp] = 'S038') or (vCod_hm[iTemp] = 'S040') or (vCod_hm[iTemp] = 'S053') or
                                (vCod_hm[iTemp] = 'S054') or (vCod_hm[iTemp] = 'SE38') or (vCod_hm[iTemp] = 'T015') or
                                (vCod_hm[iTemp] = 'C036') or (vCod_hm[iTemp] = 'C038') 
                                then


                             if (vCod_hm[iTemp] = 'E012') and (Trim(Copy(UV_fGulkwa, 133,   6)) = '') then
                             begin
                             UV_v070[Index] := '결과 무';
                             Sum_vE012 := Sum_vE012  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E013') and (Trim(Copy(UV_fGulkwa, 139,   6)) = '') then
                             begin
                             UV_v060[Index] := '결과 무';
                             Sum_vE013 := Sum_vE013  + 1;
                             check := True;
                              end;
                             if (vCod_hm[iTemp] = 'E014') and (Trim(Copy(UV_fGulkwa, 145,   6)) = '') then
                             begin
                             UV_v086[Index] := '결과 무';
                             Sum_vE014 := Sum_vE014  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E016') and (Trim(Copy(UV_fGulkwa, 157,   6)) = '') then
                             begin
                             UV_v009[Index] := '결과 무';
                             Sum_vE016 := Sum_vE016  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E021') and (Trim(Copy(UV_fGulkwa, 187,   6)) = '') then
                             begin
                             UV_v005[Index] := '결과 무';
                             Sum_vE021 := Sum_vE021  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S055') and (Trim(Copy(UV_fGulkwa, 475,   6)) = '') then
                             begin
                             UV_v006[Index] := '결과 무';
                             Sum_vS055 := Sum_vS055  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE16') and (Trim(Copy(UV_fGulkwa, 175,   6)) = '') then
                             begin
                             UV_v016[Index] := '결과 무';
                             Sum_vSE16 := Sum_vSE16  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T012') and (Trim(Copy(UV_fGulkwa, 505,   6)) = '') then
                             begin
                             UV_v010[Index] := '결과 무';
                             Sum_vT012 := Sum_vT012  + 1;
                             check := True;
                             end;

                             if (vCod_hm[iTemp] = 'C018') and (Trim(Copy(UV_fGulkwa, 19,    6)) = '') then
                             begin
                             UV_v109[Index] := '결과 무';
                             Sum_vC018 := Sum_vC018  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'E010') and (Trim(Copy(UV_fGulkwa, 121,   6)) = '') then
                             begin
                             UV_v110[Index] := '결과 무';
                             Sum_vE010 := Sum_vE010  + 1;
                             end;
                             if (vCod_hm[iTemp] = 'S023') and (Trim(Copy(UV_fGulkwa, 565,   6)) = '') then
                             begin
                             UV_v111[Index] := '결과 무';
                             Sum_vS023 := Sum_vS023  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S029') and (Trim(Copy(UV_fGulkwa, 1,     6)) = '') then
                             begin
                             UV_v112[Index] := '결과 무';
                             Sum_vS029 := Sum_vS029  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S038') and (Trim(Copy(UV_fGulkwa, 373,   6)) = '') then
                             begin
                             UV_v113[Index] := '결과 무';
                             Sum_vS038 := Sum_vS038  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S040') and (Trim(Copy(UV_fGulkwa, 385,   6)) = '') then
                             begin
                             UV_v114[Index] := '결과 무';
                             Sum_vS040 := Sum_vS040  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S053') and (Trim(Copy(UV_fGulkwa, 463,   6)) = '') then
                             begin
                             UV_v115[Index] := '결과 무';
                             Sum_vS053 := Sum_vS053  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'S054') and (Trim(Copy(UV_fGulkwa, 469,   6)) = '') then
                             begin
                             UV_v116[Index] := '결과 무';
                             Sum_vS054 := Sum_vS054  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'SE38') and (Trim(Copy(UV_fGulkwa, 73,    6)) = '') then
                             begin
                             UV_v117[Index] := '결과 무';
                             Sum_vSE38 := Sum_vSE38  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'T015') and (Trim(Copy(UV_fGulkwa, 523,   6)) = '') then
                             begin
                             UV_v118[Index] := '결과 무';
                             Sum_vT015 := Sum_vT015  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C036') and (Trim(Copy(UV_fGulkwa, 25,   6)) = '') then
                             begin
                             UV_v119[Index] := '결과 무';
                             Sum_vC036 := Sum_vC036  + 1;
                             check := True;
                             end;
                             if (vCod_hm[iTemp] = 'C038') and (Trim(Copy(UV_fGulkwa, 37,   6)) = '') then
                             begin
                             UV_v017[Index] := '결과 무';
                             Sum_vC038 := Sum_vC038  + 1;
                             check := True;
                             end;
                           end;
                         end;     
                  end; // end of Case[Gubn_Part];
                  qryHul.Next;
               end; // end of while[qryHul];
            end; // end of if[qryHul];
            qryHul.Active := False;

            if check = True then
            begin
              UV_v001[Index]        := FieldByName('dat_bunju').AsString;
              UV_v002[Index]        := FieldByName('num_bunju').AsString;
              UV_v003[Index]        := FieldByName('num_samp').AsString;
              UV_v004[Index]        := FieldByName('desc_name').AsString;
              Inc(Index);
            end;



            Next;
         end;
            UV_v001[Index]        :=  '';           UV_v021[Index] := Sum_vU049;         UV_v041[Index] := Sum_vS086;
            UV_v002[Index]        :=  '';           UV_v022[Index] := Sum_vC075;         UV_v042[Index] := Sum_vS087;
            UV_v003[Index]        :=  '';           UV_v023[Index] := Sum_vH031;         UV_v043[Index] := Sum_vS088;
            UV_v004[Index]        := '합계:';       UV_v024[Index] := Sum_vC076;         UV_v044[Index] := Sum_vS089;
            UV_v005[Index] := Sum_vE021;            UV_v025[Index] := Sum_vE050;         UV_v045[Index] := Sum_vS064;
            UV_v006[Index] := Sum_vS055;            UV_v026[Index] := Sum_vU039;         UV_v046[Index] := Sum_vC065;
            UV_v007[Index] := Sum_vT043;            UV_v027[Index] := Sum_vU040;         UV_v047[Index] := Sum_vC072;
            UV_v008[Index] := Sum_vU034;            UV_v028[Index] := Sum_vU041;         UV_v048[Index] := Sum_vC079;
            UV_v009[Index] := Sum_vE016;            UV_v029[Index] := Sum_vE051;         UV_v049[Index] := Sum_vS059;
            UV_v010[Index] := Sum_vT012;            UV_v030[Index] := Sum_vE052;         UV_v050[Index] := Sum_vS060;
            UV_v011[Index] := Sum_vC081;            UV_v031[Index] := Sum_vE053;         UV_v051[Index] := Sum_vS061;
            UV_v012[Index] := Sum_vT038;            UV_v032[Index] := Sum_vE054;         UV_v052[Index] := Sum_vS062;
            UV_v013[Index] := Sum_vC070;            UV_v033[Index] := Sum_vE055;         UV_v053[Index] := Sum_vS063;
            UV_v014[Index] := Sum_vY009;            UV_v034[Index] := Sum_vE056;         UV_v054[Index] := Sum_vS070;
            UV_v015[Index] := Sum_vE019;            UV_v035[Index] := Sum_vE057;         UV_v055[Index] := Sum_vS098;
            UV_v016[Index] := Sum_vSE16;            UV_v036[Index] := Sum_vE058;         UV_v056[Index] := Sum_vS030;
            UV_v017[Index] := Sum_vC038;            UV_v037[Index] := Sum_vT005;         UV_v057[Index] := Sum_vS031;
            UV_v018[Index] := Sum_vS093;            UV_v038[Index] := Sum_vS022;         UV_v058[Index] := Sum_vS032;
            UV_v019[Index] := Sum_vC020;            UV_v039[Index] := Sum_vE049;         UV_v059[Index] := Sum_vT018;
            UV_v020[Index] := Sum_vC064;            UV_v040[Index] := Sum_vS019;         UV_v060[Index] := Sum_vE013;


            UV_v061[Index] := Sum_vT021;            UV_v081[Index] := Sum_vT033;         UV_v101[Index] := Sum_vU048;
            UV_v062[Index] := Sum_vT022;            UV_v082[Index] := Sum_vT035;         UV_v102[Index] := Sum_vS092;
            UV_v063[Index] := Sum_vT023;            UV_v083[Index] := Sum_vT036;         UV_v103[Index] := Sum_vT050;
            UV_v064[Index] := Sum_vT024;            UV_v084[Index] := Sum_vH027;         UV_v104[Index] := Sum_vC053;
            UV_v065[Index] := Sum_vS094;            UV_v085[Index] := Sum_vH028;         UV_v105[Index] := Sum_vC055;
            UV_v066[Index] := Sum_vE047;            UV_v086[Index] := Sum_vE014;         UV_v106[Index] := Sum_vS018;
            UV_v067[Index] := Sum_vS076;            UV_v087[Index] := Sum_vE060;         UV_v107[Index] := Sum_vT039;
            UV_v068[Index] := Sum_vS077;            UV_v088[Index] := Sum_vE061;         UV_v108[Index] := Sum_vU019;
            UV_v069[Index] := Sum_vS078;            UV_v089[Index] := Sum_vE062;         UV_v109[Index] := Sum_vC018;
            UV_v070[Index] := Sum_vE012;            UV_v090[Index] := Sum_vE063;         UV_v110[Index] := Sum_vE010;
            UV_v071[Index] := Sum_vS085;            UV_v091[Index] := Sum_vE064;         UV_v111[Index] := Sum_vS023;
            UV_v072[Index] := Sum_vSE26;            UV_v092[Index] := Sum_vE065;         UV_v112[Index] := Sum_vS029;
            UV_v073[Index] := Sum_vSE27;            UV_v093[Index] := Sum_vT044;         UV_v113[Index] := Sum_vS038;
            UV_v074[Index] := Sum_vSE28;            UV_v094[Index] := Sum_vT045;         UV_v114[Index] := Sum_vS040;
            UV_v075[Index] := Sum_vSE29;            UV_v095[Index] := Sum_vT046;         UV_v115[Index] := Sum_vS053;
            UV_v076[Index] := Sum_vS095;            UV_v096[Index] := Sum_vT047;         UV_v116[Index] := Sum_vS054;
            UV_v077[Index] := Sum_vT026;            UV_v097[Index] := Sum_vT048;         UV_v117[Index] := Sum_vSE38;
            UV_v078[Index] := Sum_vT027;            UV_v098[Index] := Sum_vE066;         UV_v118[Index] := Sum_vT015;
            UV_v079[Index] := Sum_vT028;            UV_v099[Index] := Sum_vE067;         UV_v119[Index] := Sum_vC036;
            UV_v080[Index] := Sum_vT032;            UV_v100[Index] := Sum_vT049;         UV_v120[Index] := Sum_vT009;
            UV_v121[Index] := Sum_vT040;            UV_v122[Index] := Sum_vY005;         UV_v123[Index] := Sum_vY008;
            Inc(Index);

      end
      else
      begin
         ShowMessage('조건에 맞는 자료가 존재하지 않습니다.');
      end;
      // Table과 Disconnect -> 이후의 작업들은 Variant변수를 이용해서 수행
      Active := False;
   end;

   // Grid에 자료를 할당
   grdMaster.Rows := Index;
   // Grid Focus
   grdMaster.SetFocus;
   // Data건수 표시
   GP_SetDataCnt(grdMaster);
   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

function TfrmLD8Q09.UF_sokun(sValue : String) : String;
begin
   inherited;
   Result := '';
   with qrySokun do
   begin
      Active := False;
      ParamByName('COD_SOKUN').AsString := sValue;
      Active := True;
      if RecordCount > 0 then
      begin                                                                                                                 
         Result := FieldByName('desc_sokun').AsString;
      end;
      Active := False;
   end;
end;

procedure TfrmLD8Q09.UP_Clear(Temp : Integer);
begin
    UV_v001[Temp] := ''; UV_v002[Temp] := ''; UV_v003[Temp] := ''; UV_v004[Temp] := ''; UV_v005[Temp] := '';
    UV_v006[Temp] := ''; UV_v007[Temp] := ''; UV_v008[Temp] := ''; UV_v009[Temp] := ''; UV_v010[Temp] := '';
    UV_v011[Temp] := ''; UV_v012[Temp] := ''; UV_v013[Temp] := ''; UV_v014[Temp] := ''; UV_v015[Temp] := '';
    UV_v016[Temp] := ''; UV_v017[Temp] := ''; UV_v018[Temp] := ''; UV_v019[Temp] := ''; UV_v020[Temp] := '';
    UV_v021[Temp] := ''; UV_v022[Temp] := ''; UV_v023[Temp] := ''; UV_v024[Temp] := ''; UV_v025[Temp] := '';
    UV_v026[Temp] := ''; UV_v027[Temp] := ''; UV_v028[Temp] := ''; UV_v029[Temp] := ''; UV_v030[Temp] := '';
    UV_v031[Temp] := ''; UV_v032[Temp] := ''; UV_v033[Temp] := ''; UV_v034[Temp] := ''; UV_v035[Temp] := '';
    UV_v036[Temp] := ''; UV_v037[Temp] := ''; UV_v038[Temp] := ''; UV_v039[Temp] := ''; UV_v040[Temp] := '';
    UV_v041[Temp] := ''; UV_v042[Temp] := ''; UV_v043[Temp] := ''; UV_v044[Temp] := ''; UV_v045[Temp] := '';
    UV_v046[Temp] := ''; UV_v047[Temp] := ''; UV_v048[Temp] := ''; UV_v049[Temp] := ''; UV_v050[Temp] := '';
    UV_v051[Temp] := ''; UV_v052[Temp] := ''; UV_v053[Temp] := ''; UV_v054[Temp] := ''; UV_v055[Temp] := '';
    UV_v056[Temp] := ''; UV_v057[Temp] := ''; UV_v058[Temp] := ''; UV_v059[Temp] := ''; UV_v060[Temp] := '';
    UV_v061[Temp] := ''; UV_v062[Temp] := ''; UV_v063[Temp] := ''; UV_v064[Temp] := ''; UV_v065[Temp] := '';
    UV_v066[Temp] := ''; UV_v067[Temp] := ''; UV_v068[Temp] := ''; UV_v069[Temp] := ''; UV_v070[Temp] := '';
    UV_v071[Temp] := ''; UV_v072[Temp] := ''; UV_v073[Temp] := ''; UV_v074[Temp] := ''; UV_v075[Temp] := '';
    UV_v076[Temp] := ''; UV_v077[Temp] := ''; UV_v078[Temp] := ''; UV_v079[Temp] := ''; UV_v080[Temp] := '';
    UV_v081[Temp] := ''; UV_v082[Temp] := ''; UV_v083[Temp] := ''; UV_v084[Temp] := ''; UV_v085[Temp] := '';
    UV_v086[Temp] := ''; UV_v087[Temp] := ''; UV_v088[Temp] := ''; UV_v089[Temp] := ''; UV_v090[Temp] := '';
    UV_v091[Temp] := ''; UV_v092[Temp] := ''; UV_v093[Temp] := ''; UV_v094[Temp] := ''; UV_v095[Temp] := '';
    UV_v096[Temp] := ''; UV_v097[Temp] := ''; UV_v098[Temp] := ''; UV_v099[Temp] := ''; UV_v100[Temp] := '';
    UV_v101[Temp] := ''; UV_v102[Temp] := ''; UV_v103[Temp] := ''; UV_v104[Temp] := ''; UV_v105[Temp] := '';
    UV_v106[Temp] := ''; UV_v107[Temp] := ''; UV_v108[Temp] := ''; UV_v109[Temp] := ''; UV_v110[Temp] := '';
    UV_v111[Temp] := ''; UV_v112[Temp] := ''; UV_v113[Temp] := ''; UV_v114[Temp] := ''; UV_v115[Temp] := '';
    UV_v116[Temp] := ''; UV_v117[Temp] := ''; UV_v118[Temp] := ''; UV_v119[Temp] := ''; UV_v120[Temp] := '';
    UV_v121[Temp] := ''; UV_v122[Temp] := ''; UV_v123[Temp] := '';

end;

procedure TfrmLD8Q09.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
   inherited;
   if Sender = btnDc then
   begin
      if GF_DancheClick(Self, sCode, sName) then
      begin
         edtCod_dc.Text  := sCode;
         edtDesc_Dc.Text := sName;
      end;
   end
   else if Sender = btnGmgnF then GF_CalendarClick(mskST_date)
   else if Sender = btnGmgnT then GF_CalendarClick(mskEd_date);
end;

function TfrmLD8Q09.UF_Condition : Boolean;
begin
   Result := True;

   // 조회조건중 필수항목이 입력되었는지 Check
   if (cmbJisa.ItemIndex = -1) or
      ((mskST_date.Text = '') and (mskEd_date.Text = '' )) then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

  // Urin functin
function  TfrmLD8Q09.UF_Kicho_Urin(sValue : Integer): String;
begin
   Result := '';
   case sValue of
      1 : Result := '음성';
      2 : Result := '±';
      3 : Result := '+1';
      4 : Result := '+2';
      5 : Result := '+3';
      6 : Result := '+4';
   end;
end;



procedure TfrmLD8Q09.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if      Sender = edtCod_dc  then UP_Click(btnDc)
   else if Sender = mskST_date then UP_Click(btnGmgnF)
   else if Sender = mskEd_date then UP_Click(btnGmgnT);

end;

procedure TfrmLD8Q09.btnQuitClick(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmLD8Q09.UP_Change(Sender: TObject);
var sName : String;
begin
  inherited;

   if Sender = edtCod_dc then
   begin
      if GF_DancheEdit(edtCod_dc, sName) then
         edtDesc_Dc.Text := sName;
   end;
end;

procedure TfrmLD8Q09.SBut_ExcelClick(Sender: TObject);
var
  XL,WorkBook: Variant;
  i : integer;
  ArrV3 : OleVariant;
  Row,Col : Integer;
begin
   inherited;
   Screen.Cursor:= crHourGlass;
   try
      XL:= CreateOleObject('Excel.Application');
   except
      Application.MessageBox('Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.',
                             '오류', MB_OK or MB_ICONERROR);
      Screen.Cursor  := crDefault;
      Exit;
   end;

   Gauge2.MaxValue := grdMaster.Rows;
   Gauge2.Progress := 1;
   Application.ProcessMessages;
   try
      WorkBook := XL.WorkBooks.Add;

      //Data import
      ArrV3 := VarArrayCreate([0, grdMaster.Rows + 1, 0, grdMaster.Cols], varOleStr);

      I := 0;
      for Row := 0 to grdMaster.Rows + 1 do
      begin
         if Row = 0 then
         begin
            for Col := 0 to grdMaster.Cols - 1 do
               ArrV3[Row, Col] := grdMaster.Col[Col + 1].Heading;
         end
         else
         begin
            for Col := 0 to grdMaster.Cols do
            begin
               Application.ProcessMessages;
               ArrV3[Row, Col] := grdMaster.cell[Col + 1, Row];
            end;
         end;
         Gauge2.Progress:= i;
         Application.ProcessMessages;
         Inc(I);
      end;
      XL.Range[XL.Cells[1, 1], XL.Cells[grdMaster.Rows + 1, grdMaster.Cols]].Value := ArrV3;


      XL.DisplayAlerts := False;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.

