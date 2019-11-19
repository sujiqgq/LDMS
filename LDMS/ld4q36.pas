//================================================================================
// ��������      : 2013.7.23
// ������        : ����
// ��������      : ���� ������ܸ鿪����1300097
// �������      :
//==============================================================================
// ��������      : 2014.03.26
// ������        : ������
// ��������      : �����׸� - T009�׸� ����
// �������      :
//==============================================================================
// ��������      : 2014.04.07
// ������        : ������
// ��������      : ��ü��ȸ & Special(C2):T043����, �����׸�:T043����, Rack.No�� ��ȸ����
// �������      :
//==============================================================================
// ��������      : 2014.05.03
// ������        : ������
// ��������      : ��ü��ȸ & Special(C2):T040����, �����׸�:T040����
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.07.11
// ������        : ������
// ��������      : �̻���ڵ� : SE02�߰�
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.08.20
// ������        : ������
// ��������      : 20140901 �������ں���
//                 T043 �����׸� -> Routine / T042 Special -> Routine
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.09.03
// ������        : ������
// ��������      : IMMULITE - T039�ڵ� �߰�
// �������      : [����-�ѵη�]
//==============================================================================
// ��������      : 2014.12.11
// ������        : ������
// ��������      : Ư�� ���� �������� ���� ��������� ����
// �������      :
//==============================================================================
// ��������      : 2015.02.11
// ������        : ������
// ��������      : S004 -> toshiba�� ��ȸ�����ϵ��� ����
// �������      :  UV_vS004
//==============================================================================
// ��������      : 2015.06.22
// ������        : ������
// ��������      : S079[�˷�����-���Լ� 42��]�� S090[�˷�����-���̼� 42��] ��ȸ
// �������      : ���� ������ܰ˻�������1500495 [�������ܰ˻�������-�ѹ̿�]
//==============================================================================
// ��������      : 2015.11.25
// ������        : ������
// ��������      : S008, S091 ���� ��ȸ ��� �߰�
// �������      : ���� ������ܰ˻�������1500972
//==============================================================================
// ��������      : 20160122
// ������        : �ڼ���
// ��������      : tt03 �׸� �߰�
// �������      : ���� ������ܰ˻������� 1600059
//==============================================================================
// ��������      : 20160609
// ������        : �ڼ���
// ��������      : Routine ���ý� �˻��׸� �׸� ��ȸ
// �������      : �ѹ̿� ���� ������ܰ˻������� 1600533
//==============================================================================
// ��������      : 20160712
// ������        : �ڼ���
// ��������      : c020,s033
// �������      : �ѹ̿� ���� ������ܰ˻������� 1600610, 1600614
//==============================================================================
unit LD4Q36;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, DdeMan, Gauges,
  QuickRpt, ComObj;

type
  TfrmLD4Q36 = class(TfrmSingle)
    grdMaster: TtsGrid;
    pnlCond: TPanel;
    Label7: TLabel;
    cmbJisa: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    btnBdate: TSpeedButton;
    edtBDate: TDateEdit;
    bunjuno1: TEdit;
    bunjuno2: TEdit;
    qryBunju: TQuery;
    RBtn_preveiw: TRadioButton;
    RadioButton2: TRadioButton;
    Gride: TGauge;
    qryNo_hangmok: TQuery;
    qryGulkwa: TQuery;
    MEdt_SampS: TMaskEdit;
    MEdt_SampE: TMaskEdit;
    Label13: TLabel;
    Label4: TLabel;
    Cmb_gubn: TComboBox;
    qryJHangmokList: TQuery;
    qryPf_hm: TQuery;
    qryTkgum: TQuery;
    Label1: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Edit_RackS: TMaskEdit;
    Edit_RackE: TMaskEdit;
    qryProfile: TQuery;
    Rout_01: TRadioButton;
    Rout_02: TRadioButton;
    Rout_03: TRadioButton;
    Rout_04: TRadioButton;
    Rout_05: TRadioButton;
    Rout_06: TRadioButton;
    Rout_07: TRadioButton;
    Rout_08: TRadioButton;
    Rout_09: TRadioButton;
    GroupRadio2: TGroupBox;
    Rout_10: TRadioButton;
    Rout_11: TRadioButton;
    Rout_12: TRadioButton;
    SBut_Excel: TSpeedButton;
    Gauge2: TGauge;
    group_radio: TGroupBox;
    Chk_ALL: TRadioButton;
    Chk_Arch: TRadioButton;
    CHK_Not_used: TRadioButton;
    CHK_sub: TRadioButton;
    CHK_S101: TRadioButton;
    CHK_S102: TRadioButton;
    Chk_C1_C5: TRadioButton;
    chk_S033: TRadioButton;
    Chk_NKCell: TRadioButton;
    CB02: TCheckBox;
    CHK_S104: TRadioButton;
    CB01: TCheckBox;
    Rout_13: TRadioButton;
    Chk_Toshiba: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure btnQueryClick(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UP_Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure MouseWheelHandler(var Message: TMessage); override;
    function  UF_hangmokList : String;
    function  UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
    procedure SBut_ExcelClick(Sender: TObject);
  private
    { Private declarations }
    // Grid�� ������ Variant ���� ����
    UV_vNo, UV_vName, UV_vJumin, UV_vJisa, UV_vBUN,
    UV_vS010,UV_vS011,UV_vS012,UV_vT009,UV_vT040,UV_vT041,UV_vT042,UV_vT043,UV_vS094,UV_vS095,UV_vS096,UV_vS097,
    UV_vS016,UV_vS019,UV_vS021,UV_vE068,UV_vS018,
    UV_vS007,UV_vS008,UV_vSE01,UV_vSE02,UV_vTT01,UV_vTT02,UV_vTT03, UV_vE069,UV_vS020,UV_vS091,
    UV_vS001,UV_vS003,UV_vS034,UV_vS036,UV_vS099,
    Uv_vS002,Uv_vS004,Uv_vS035,Uv_vS037,Uv_vS071,Uv_vS080,Uv_vS083,Uv_vS084,
    Uv_vS098,Uv_vC031,Uv_vC035,Uv_vSE07,Uv_vSE08,Uv_vSE30,Uv_vT008,Uv_vT029,Uv_vT030,Uv_vT031,Uv_vT034,Uv_vSE31,
    Uv_vT039,Uv_vS101,Uv_vS102,Uv_vS104,Uv_vC020,Uv_vS033,Uv_vSE03,

    UV_vSampNo, UV_vDesc_dc : Variant;

    // ���������� ����ϴ� �Լ�(�̸��� �����ϰ� ���)
    procedure UP_VarMemSet(iLength : Integer);
    function  UF_Condition : Boolean;
  public
    { Public declarations }
  end;

var
  frmLD4Q36: TfrmLD4Q36;

implementation

uses ZZFunc, ZZMenu, ZZComm, MdmsFunc, LD4Q361;

{$R *.DFM}

procedure TfrmLD4Q36.UP_KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
   if Key <> GC_Refer then exit;

 //if Sender = edtDc then UP_Click(btnDC);

end;

procedure TfrmLD4Q36.MouseWheelHandler(var Message: TMessage);  //����
begin
   if Message.msg = WM_MOUSEWHEEL then begin
      if (ActiveControl is Ttsgrid) then begin //����
         if Message.wParam > 0 then
         keybd_event(VK_UP, VK_UP, 0, 0)
         else if Message.wParam < 0 then
         keybd_event(VK_DOWN, VK_DOWN, 0, 0);
         Application.ProcessMessages;//�߰�
         grdMaster.Refresh; //����
      end;
   end;
end;

procedure TfrmLD4Q36.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant������ ����ϱ� ���ؼ� ����
      UV_vNo    := VarArrayCreate([0, 0], varOleStr);
      UV_vName  := VarArrayCreate([0, 0], varOleStr);
      UV_vJumin := VarArrayCreate([0, 0], varOleStr);
      UV_vJisa  := VarArrayCreate([0, 0], varOleStr);
      UV_vBUN   := VarArrayCreate([0, 0], varOleStr);
      UV_vS010  := VarArrayCreate([0, 0], varOleStr);
      UV_vS011  := VarArrayCreate([0, 0], varOleStr);
      UV_vS012  := VarArrayCreate([0, 0], varOleStr);
      UV_vT009  := VarArrayCreate([0, 0], varOleStr);
      UV_vT040  := VarArrayCreate([0, 0], varOleStr);
      UV_vT041  := VarArrayCreate([0, 0], varOleStr);
      UV_vT042  := VarArrayCreate([0, 0], varOleStr);
      UV_vT043  := VarArrayCreate([0, 0], varOleStr);
      UV_vS094  := VarArrayCreate([0, 0], varOleStr);
      UV_vS095  := VarArrayCreate([0, 0], varOleStr);
      UV_vS096  := VarArrayCreate([0, 0], varOleStr);
      UV_vS097  := VarArrayCreate([0, 0], varOleStr);
      UV_vS016  := VarArrayCreate([0, 0], varOleStr);
      UV_vS019  := VarArrayCreate([0, 0], varOleStr);
      UV_vS021  := VarArrayCreate([0, 0], varOleStr);
      UV_vE068  := VarArrayCreate([0, 0], varOleStr);
      UV_vS007  := VarArrayCreate([0, 0], varOleStr);
      UV_vS008  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE01  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE02  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE07  := VarArrayCreate([0, 0], varOleStr);
      UV_vSE08  := VarArrayCreate([0, 0], varOleStr);
      UV_vTT01  := VarArrayCreate([0, 0], varOleStr);
      UV_vTT02  := VarArrayCreate([0, 0], varOleStr);
      UV_vTT03  := VarArrayCreate([0, 0], varOleStr);
      UV_vE069  := VarArrayCreate([0, 0], varOleStr);
      UV_vS020  := VarArrayCreate([0, 0], varOleStr);
      UV_vS001  := VarArrayCreate([0, 0], varOleStr);
      UV_vS003  := VarArrayCreate([0, 0], varOleStr);
      UV_vS034  := VarArrayCreate([0, 0], varOleStr);
      UV_vS036  := VarArrayCreate([0, 0], varOleStr);
      UV_vS091  := VarArrayCreate([0, 0], varOleStr);
      UV_vS099  := VarArrayCreate([0, 0], varOleStr);
      UV_vS018  := VarArrayCreate([0, 0], varOleStr);
      UV_vT009  := VarArrayCreate([0, 0], varOleStr);
      Uv_vS002 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS004 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS035 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS037 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS071 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS080 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS083 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS084 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS098 := VarArrayCreate([0, 0], varOleStr);
      Uv_vC031 := VarArrayCreate([0, 0], varOleStr);
      Uv_vC035 := VarArrayCreate([0, 0], varOleStr);
      Uv_vSE07 := VarArrayCreate([0, 0], varOleStr);
      Uv_vSE08 := VarArrayCreate([0, 0], varOleStr);
      Uv_vSE30 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT008 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT029 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT030 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT031 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT034 := VarArrayCreate([0, 0], varOleStr);
      Uv_vSE31 := VarArrayCreate([0, 0], varOleStr);
      Uv_vT039 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS101 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS102 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS104 := VarArrayCreate([0, 0], varOleStr);
      Uv_vC020 := VarArrayCreate([0, 0], varOleStr);
      Uv_vS033 := VarArrayCreate([0, 0], varOleStr);
      Uv_vSE03 := VarArrayCreate([0, 0], varOleStr);

      UV_vSampNo:= VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // �̹� ������ �������� ũ�⸦ ����
      VarArrayRedim(UV_vNo,    iLength);
      VarArrayRedim(UV_vName,  iLength);
      VarArrayRedim(UV_vJumin, iLength);
      VarArrayRedim(UV_vJisa,  iLength);
      VarArrayRedim(UV_vBUN,   iLength);
      VarArrayRedim(UV_vT009,   iLength);
      VarArrayRedim(UV_vS010,  iLength);
      VarArrayRedim(UV_vS011,  iLength);
      VarArrayRedim(UV_vS012,  iLength);
      VarArrayRedim(UV_vT009,  iLength); //
      VarArrayRedim(UV_vT040,  iLength);
      VarArrayRedim(UV_vT041,  iLength);
      VarArrayRedim(UV_vT042,  iLength);
      VarArrayRedim(UV_vT043,  iLength);
      VarArrayRedim(UV_vS094,  iLength);
      VarArrayRedim(UV_vS095,  iLength);
      VarArrayRedim(UV_vS096,  iLength);
      VarArrayRedim(UV_vS097,  iLength);
      VarArrayRedim(UV_vS016,  iLength);
      VarArrayRedim(UV_vS019,  iLength); //
      VarArrayRedim(UV_vS021,  iLength);
      VarArrayRedim(UV_vE068,  iLength);
      VarArrayRedim(UV_vS007,  iLength);
      VarArrayRedim(UV_vS008,  iLength);
      VarArrayRedim(UV_vSE01,  iLength);
      VarArrayRedim(UV_vSE02,  iLength);
      VarArrayRedim(UV_vSE07,  iLength);
      VarArrayRedim(UV_vSE08,  iLength);
      VarArrayRedim(UV_vTT01,  iLength);
      VarArrayRedim(UV_vTT02,  iLength); //
      VarArrayRedim(UV_vTT03,  iLength);
      VarArrayRedim(UV_vE069,  iLength);
      VarArrayRedim(UV_vS020,  iLength);
      VarArrayRedim(UV_vS001,  iLength);
      VarArrayRedim(UV_vS003,  iLength);
      VarArrayRedim(UV_vS034,  iLength);
      VarArrayRedim(UV_vS036,  iLength);
      VarArrayRedim(UV_vS091,  iLength);
      VarArrayRedim(UV_vS099,  iLength);
      VarArrayRedim(UV_vS018,  iLength); //

      VarArrayRedim(Uv_vS002,  iLength);
      VarArrayRedim(Uv_vS004,  iLength);
      VarArrayRedim(Uv_vS035,  iLength);
      VarArrayRedim(Uv_vS037,  iLength);
      VarArrayRedim(Uv_vS071,  iLength);
      VarArrayRedim(Uv_vS080,  iLength);
      VarArrayRedim(Uv_vS083,  iLength);
      VarArrayRedim(Uv_vS084,  iLength);
      VarArrayRedim(Uv_vS098,  iLength);
      VarArrayRedim(Uv_vC031,  iLength);  //
      VarArrayRedim(Uv_vC035,  iLength);
      VarArrayRedim(Uv_vSE07,  iLength);
      VarArrayRedim(Uv_vSE08,  iLength);
      VarArrayRedim(Uv_vSE30,  iLength);
      VarArrayRedim(Uv_vT008,  iLength);
      VarArrayRedim(Uv_vT029,  iLength);
      VarArrayRedim(Uv_vT030,  iLength);
      VarArrayRedim(Uv_vT031,  iLength);
      VarArrayRedim(Uv_vT034,  iLength);
      VarArrayRedim(Uv_vSE31,  iLength);  //
      VarArrayRedim(Uv_vT039,  iLength);
      VarArrayRedim(Uv_vS101,  iLength);
      VarArrayRedim(Uv_vS102,  iLength);
      VarArrayRedim(Uv_vS104,  iLength);
      VarArrayRedim(Uv_vC020,  iLength);
      VarArrayRedim(Uv_vS033,  iLength);
      VarArrayRedim(Uv_vSE03,  iLength);
      VarArrayRedim(UV_vSampNo, iLength);
   end;
end;

function TfrmLD4Q36.UF_Condition : Boolean;
begin
   Result := True;

   // ��ȸ������ �ʼ��׸��� �ԷµǾ����� Check
   if edtBDate.Text = '' then
   begin
      GF_MsgBox('Warning', 'CONDITION');
      Result := False;
   end;
end;

procedure TfrmLD4Q36.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   // ���簡 ��ü�� ��� ���縦 Combo�� ������
   GP_ComboJisa(cmbjisa);
   // Default�� ����(15)�� ����
   GP_ComboDisplay(cmbjisa, copy(GV_sUserId,1,2), 2);
end;

procedure TfrmLD4Q36.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
  var i :integer;
begin
   inherited;

   for i:=5 to 36 do
   begin
        grdMaster.Col[i].heading := ' ';
        grdMaster.Col[i].Width   := 45;
   end;

   // �ڷ㸦 ȭ�鿡 ��ȸ
   if Chk_all.Checked =true  then
   begin
      if CB01.Checked then
      begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vC020[DataRow-1];
            6 : Value := UV_vS010[DataRow-1];
            7 : Value := UV_vS011[DataRow-1];
            8 : Value := UV_vS012[DataRow-1];
            9 : Value := UV_vT041[DataRow-1];
           10 : Value := UV_vT042[DataRow-1];
           11 : Value := UV_vS016[DataRow-1];
           12 : Value := UV_vS007[DataRow-1];
           13 : Value := UV_vSE01[DataRow-1];
           14 : Value := UV_vSE03[DataRow-1];
           15 : Value := UV_vSE31[DataRow-1];
           16 : Value := UV_vS020[DataRow-1];
           17 : Value := UV_vS001[DataRow-1];
           18 : Value := UV_vS003[DataRow-1];
           19 : Value := UV_vS033[DataRow-1];
           20 : Value := UV_vS034[DataRow-1];
           21 : Value := UV_vS036[DataRow-1];
           22 : Value := UV_vS099[DataRow-1];
           23 : Value := UV_vT043[DataRow-1];
           24 : Value := UV_vS004[DataRow-1];
           25 :Value := UV_vTT03[DataRow-1];
         end;
         grdMaster.Col[5].heading   :='C020';
         grdMaster.Col[6].heading   :='S010';
         grdMaster.Col[7].heading   :='S011';
         grdMaster.Col[8].heading   :='S012';
         grdMaster.Col[9].heading   :='T041';
         grdMaster.Col[10].heading  :='T042';
         grdMaster.Col[11].heading  :='S016';
         grdMaster.Col[12].heading  :='S007';
         grdMaster.Col[13].heading  :='SE01';
         grdMaster.Col[14].heading  :='SE03';
         grdMaster.Col[15].heading  :='SE31';
         grdMaster.Col[16].heading  :='S020';
         grdMaster.Col[17].heading  :='S001';
         grdMaster.Col[18].heading  :='S003';
         grdMaster.Col[19].heading  :='S033';
         grdMaster.Col[20].heading  :='S034';
         grdMaster.Col[21].heading  :='S036';
         grdMaster.Col[22].heading  :='S099';
         grdMaster.Col[23].heading  :='T043';
         grdMaster.Col[24].heading  :='S004';
         grdMaster.Col[25].heading  :='TT03';
      end
      else if CB02.Checked then
      begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vC020[DataRow-1];
            6 : Value := UV_vS010[DataRow-1];
            7 : Value := UV_vS011[DataRow-1];
            8 : Value := UV_vS012[DataRow-1];
            9 : Value := UV_vT041[DataRow-1];
           10 : Value := UV_vT042[DataRow-1];
           11 : Value := UV_vS016[DataRow-1];
           12 : Value := UV_vS007[DataRow-1];
           13 : Value := UV_vSE01[DataRow-1];
           14 : Value := UV_vSE03[DataRow-1];
           15 : Value := UV_vSE31[DataRow-1];
           16 : Value := UV_vS020[DataRow-1];
           17 : Value := UV_vS001[DataRow-1];
           18 : Value := UV_vS003[DataRow-1];
           19 : Value := UV_vS033[DataRow-1];
           20 : Value := UV_vS034[DataRow-1];
           21 : Value := UV_vS036[DataRow-1];
           22 : Value := UV_vS099[DataRow-1];
           23 : Value := UV_vT043[DataRow-1];
           24 : Value := UV_vTT03[DataRow-1];
           25 : Value := UV_vS008[DataRow-1];
           26 : Value := UV_vS091[DataRow-1];
           end;
         grdMaster.Col[5].heading   :='C020';
         grdMaster.Col[6].heading   :='S010';
         grdMaster.Col[7].heading   :='S011';
         grdMaster.Col[8].heading   :='S012';
         grdMaster.Col[9].heading   :='T041';
         grdMaster.Col[10].heading  :='T042';
         grdMaster.Col[11].heading  :='S016';
         grdMaster.Col[12].heading  :='S007';
         grdMaster.Col[13].heading  :='SE01';
         grdMaster.Col[14].heading  :='SE03';
         grdMaster.Col[15].heading  :='SE31';
         grdMaster.Col[16].heading  :='S020';
         grdMaster.Col[17].heading  :='S001';
         grdMaster.Col[18].heading  :='S003';
         grdMaster.Col[19].heading  :='S033';
         grdMaster.Col[20].heading  :='S034';
         grdMaster.Col[21].heading  :='S036';
         grdMaster.Col[22].heading  :='S099';
         grdMaster.Col[23].heading  :='T043';
         grdMaster.Col[24].heading  :='TT03';
         grdMaster.Col[25].heading  :='S008';
         grdMaster.Col[26].heading  :='S091';
      end
      else if (CB01.checked) and (CB02.checked) then
      begin
      case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vC020[DataRow-1];
            6 : Value := UV_vS010[DataRow-1];
            7 : Value := UV_vS011[DataRow-1];
            8 : Value := UV_vS012[DataRow-1];
            9 : Value := UV_vT041[DataRow-1];
           10 : Value := UV_vT042[DataRow-1];
           11 : Value := UV_vS016[DataRow-1];
           12 : Value := UV_vS007[DataRow-1];
           13 : Value := UV_vSE01[DataRow-1];
           14 : Value := UV_vSE03[DataRow-1];
           15 : Value := UV_vSE31[DataRow-1];
           16 : Value := UV_vS020[DataRow-1];
           17 : Value := UV_vS001[DataRow-1];
           18 : Value := UV_vS003[DataRow-1];
           19 : Value := UV_vS033[DataRow-1];
           20 : Value := UV_vS034[DataRow-1];
           21 : Value := UV_vS036[DataRow-1];
           22 : Value := UV_vS099[DataRow-1];
           23 : Value := UV_vT043[DataRow-1];
           24 : Value := UV_vTT03[DataRow-1];
           end;
         grdMaster.Col[5].heading   :='C020';
         grdMaster.Col[6].heading   :='S010';
         grdMaster.Col[7].heading   :='S011';
         grdMaster.Col[8].heading   :='S012';
         grdMaster.Col[9].heading   :='T041';
         grdMaster.Col[10].heading  :='T042';
         grdMaster.Col[11].heading  :='S016';
         grdMaster.Col[12].heading  :='S007';
         grdMaster.Col[13].heading  :='SE01';
         grdMaster.Col[14].heading  :='SE03';
         grdMaster.Col[15].heading  :='SE31';
         grdMaster.Col[16].heading  :='S020';
         grdMaster.Col[17].heading  :='S001';
         grdMaster.Col[18].heading  :='S003';
         grdMaster.Col[19].heading  :='S033';
         grdMaster.Col[20].heading  :='S034';
         grdMaster.Col[21].heading  :='S036';
         grdMaster.Col[22].heading  :='S099';
         grdMaster.Col[23].heading  :='T043';
         grdMaster.Col[24].heading  :='TT03';
      end
      else
      begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vC020[DataRow-1];
            6 : Value := UV_vS010[DataRow-1];
            7 : Value := UV_vS011[DataRow-1];
            8 : Value := UV_vS012[DataRow-1];
            9 : Value := UV_vT041[DataRow-1];
           10 : Value := UV_vT042[DataRow-1];
           11 : Value := UV_vS016[DataRow-1];
           12 : Value := UV_vS007[DataRow-1];
           13 : Value := UV_vSE01[DataRow-1];
           14 : Value := UV_vSE03[DataRow-1];
           15 : Value := UV_vSE31[DataRow-1];
           16 : Value := UV_vS020[DataRow-1];
           17 : Value := UV_vS001[DataRow-1];
           18 : Value := UV_vS003[DataRow-1];
           19 : Value := UV_vS033[DataRow-1];
           20 : Value := UV_vS034[DataRow-1];
           21 : Value := UV_vS036[DataRow-1];
           22 : Value := UV_vS099[DataRow-1];
           23 : Value := UV_vT043[DataRow-1];
           24 : Value := UV_vS004[DataRow-1];
           25 : Value := UV_vTT03[DataRow-1];
           26 : Value := UV_vS008[DataRow-1];
           27 : Value := UV_vS091[DataRow-1];
           end;
         grdMaster.Col[5].heading   :='C020';
         grdMaster.Col[6].heading   :='S010';
         grdMaster.Col[7].heading   :='S011';
         grdMaster.Col[8].heading   :='S012';
         grdMaster.Col[9].heading   :='T041';
         grdMaster.Col[10].heading  :='T042';
         grdMaster.Col[11].heading  :='S016';
         grdMaster.Col[12].heading  :='S007';
         grdMaster.Col[13].heading  :='SE01';
         grdMaster.Col[14].heading  :='SE03';
         grdMaster.Col[15].heading  :='SE31';
         grdMaster.Col[16].heading  :='S020';
         grdMaster.Col[17].heading  :='S001';
         grdMaster.Col[18].heading  :='S003';
         grdMaster.Col[19].heading  :='S033';
         grdMaster.Col[20].heading  :='S034';
         grdMaster.Col[21].heading  :='S036';
         grdMaster.Col[22].heading  :='S099';
         grdMaster.Col[23].heading  :='T043';
         grdMaster.Col[24].heading  :='S004';
         grdMaster.Col[25].heading  :='TT03';
         grdMaster.Col[26].heading  :='S008';
         grdMaster.Col[27].heading  :='S091';
      end;
   end
   else if Chk_Arch.Checked =true  then
   begin
      case DataCol of
         1 : Value := UV_vNo[DataRow-1];
         2 : Value := UV_vSampNo[DataRow-1];
         3 : Value := UV_vBUN[DataRow-1];
         4 : Value := UV_vName[DataRow-1];
         5 : Value := UV_vS010[DataRow-1];
         6 : Value := UV_vS011[DataRow-1];
         7 : Value := UV_vS012[DataRow-1];
         8 : Value := UV_vT041[DataRow-1];
         9 : Value := UV_vTT03[DataRow-1];
      end;
      grdMaster.Col[5].heading  :='S010';
      grdMaster.Col[6].heading  :='S011';
      grdMaster.Col[7].heading  :='S012';
      grdMaster.Col[8].heading  :='T041';
      grdMaster.Col[9].heading  :='TT03';
   end
   else if (Chk_C1_C5.Checked =true) and ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
   (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
   (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and (Rout_12.Checked = false) and (Rout_13.Checked = false)) then
   begin
      if CB01.Checked then
      begin
         case DataCol of
            1  : Value := UV_vNo[DataRow-1];
            2  : Value := UV_vSampNo[DataRow-1];
            3  : Value := UV_vBUN[DataRow-1];
            4  : Value := UV_vName[DataRow-1];
            5  : Value := UV_vC020[DataRow-1];
            6  : Value := UV_vSE01[DataRow-1];
            7  : Value := UV_vSE03[DataRow-1];
            8  : Value := UV_vSE31[DataRow-1];
            9  : Value := UV_vS020[DataRow-1];
            10 : Value := Uv_vS007[DataRow-1];
            11 : Value := Uv_vS008[DataRow-1];
            12 : Value := Uv_vS033[DataRow-1];
            13 : Value := Uv_vS091[DataRow-1];
            14 : Value := Uv_vS099[DataRow-1];
            15 : Value := Uv_vT042[DataRow-1];
            16 : Value := Uv_vT043[DataRow-1];
            17 : Value := Uv_vS016[DataRow-1];
         end;
          grdMaster.Col[5].heading  :='C020';
          grdMaster.Col[6].heading  :='SE01';
          grdMaster.Col[7].heading  :='SE03';
          grdMaster.Col[8].heading  :='SE31';
          grdMaster.Col[9].heading  :='S020';
          grdMaster.Col[10].heading :='S007';
          grdMaster.Col[11].heading :='S008';
          grdMaster.Col[12].heading :='S033';
          grdMaster.Col[13].heading :='S091';
          grdMaster.Col[14].heading :='S099';
          grdMaster.Col[15].heading :='T042';
          grdMaster.Col[16].heading :='T043';
          grdMaster.Col[17].heading :='S016';
      end
      else
        begin
         case DataCol of
            1 : Value := UV_vNo[DataRow-1];
            2 : Value := UV_vSampNo[DataRow-1];
            3 : Value := UV_vBUN[DataRow-1];
            4 : Value := UV_vName[DataRow-1];
            5 : Value := UV_vC020[DataRow-1];
            6 : Value := UV_vSE01[DataRow-1];
            7 : Value := UV_vSE03[DataRow-1];
            8 : Value := UV_vSE31[DataRow-1];
            9 : Value := UV_vS020[DataRow-1];
            10: Value := Uv_vS007[DataRow-1];
            11: Value := Uv_vS008[DataRow-1];
            12: Value := Uv_vS033[DataRow-1];
            13: Value := Uv_vS091[DataRow-1];
            14: Value := Uv_vS099[DataRow-1];
            15: Value := Uv_vT042[DataRow-1];
            16: Value := Uv_vT043[DataRow-1];
            17 : Value := Uv_vS016[DataRow-1];
         end;
         grdMaster.Col[5].heading  :='C020';
         grdMaster.Col[6].heading  :='SE01';
         grdMaster.Col[7].heading  :='SE03';
         grdMaster.Col[8].heading  :='SE31';
         grdMaster.Col[9].heading  :='S020';
         grdMaster.Col[10].heading  :='S007';
         grdMaster.Col[11].heading :='S008';
         grdMaster.Col[12].heading :='S033';
         grdMaster.Col[13].heading :='S091';
         grdMaster.Col[14].heading :='S099';
         grdMaster.Col[15].heading :='T042';
         grdMaster.Col[16].heading :='T043';
         grdMaster.Col[17].heading :='S016';
      end;
   end
   else if Chk_Toshiba.Checked =true  then
   begin
      if CB02.checked then
      begin
      case DataCol of
         1 : Value := UV_vNo[DataRow-1];
         2 : Value := UV_vSampNo[DataRow-1];
         3 : Value := UV_vBUN[DataRow-1];
         4 : Value := UV_vName[DataRow-1];
         5 : Value := UV_vS001[DataRow-1];
         6 : Value := UV_vS003[DataRow-1];
         7 : Value := UV_vS034[DataRow-1];
         8 : Value := UV_vS036[DataRow-1];
      end;
      grdMaster.Col[5].heading :='S001';
      grdMaster.Col[6].heading :='S003';
      grdMaster.Col[7].heading :='S034';
      grdMaster.Col[8].heading :='S036';
      end
      else
      begin
      case DataCol of
         1 : Value := UV_vNo[DataRow-1];
         2 : Value := UV_vSampNo[DataRow-1];
         3 : Value := UV_vBUN[DataRow-1];
         4 : Value := UV_vName[DataRow-1];
         5 : Value := UV_vS001[DataRow-1];
         6 : Value := UV_vS003[DataRow-1];
         7 : Value := UV_vS034[DataRow-1];
         8 : Value := UV_vS036[DataRow-1];
         9 : Value := UV_vS004[DataRow-1];
      end;
      grdMaster.Col[5].heading :='S001';
      grdMaster.Col[6].heading :='S003';
      grdMaster.Col[7].heading :='S034';
      grdMaster.Col[8].heading :='S036';
      grdMaster.Col[9].heading :='S004';
      end;
   end
   else if CHK_Not_used.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := Uv_vS002[DataRow-1];
         6 : Value  := Uv_vS035[DataRow-1];
         7 : Value  := Uv_vS037[DataRow-1];
         8 : Value  := Uv_vS071[DataRow-1];
         9 : Value  := Uv_vS080[DataRow-1];
         10 : Value := Uv_vS083[DataRow-1];
         11 : Value := Uv_vS084[DataRow-1];
         12 : Value := Uv_vS098[DataRow-1];
         13 : Value := Uv_vC031[DataRow-1];
         14 : Value := Uv_vC035[DataRow-1];
         15 : Value := Uv_vSE07[DataRow-1];
         16 : Value := Uv_vSE08[DataRow-1];
         17 : Value := Uv_vSE30[DataRow-1];
         18 : Value := Uv_vT008[DataRow-1];
         19 : Value := Uv_vT029[DataRow-1];
         20 : Value := Uv_vT030[DataRow-1];
         21 : Value := Uv_vT031[DataRow-1];
         22 : Value := Uv_vT034[DataRow-1];
         23 : Value := UV_vSE02[DataRow-1];
      end;
      grdMaster.Col[5].heading  :='S002';
      grdMaster.Col[6].heading  :='S035';
      grdMaster.Col[7].heading  :='S037';
      grdMaster.Col[8].heading  :='S071';
      grdMaster.Col[9].heading  :='S080';
      grdMaster.Col[10].heading :='S083';
      grdMaster.Col[11].heading :='S084';
      grdMaster.Col[12].heading :='S098';
      grdMaster.Col[13].heading :='C031';
      grdMaster.Col[14].heading :='C035';
      grdMaster.Col[15].heading :='SE07';
      grdMaster.Col[16].heading :='SE08';
      grdMaster.Col[17].heading :='SE30';
      grdMaster.Col[18].heading :='T008';
      grdMaster.Col[19].heading :='T029';
      grdMaster.Col[20].heading :='T030';
      grdMaster.Col[21].heading :='T031';
      grdMaster.Col[22].heading :='T034';
      grdMaster.Col[23].heading :='SE02';
   end
   else if CHK_sub.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := UV_vT040[DataRow-1];
      end;
      grdMaster.Col[5].heading   :='T040';
   end
   else if CHK_S101.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := UV_vS101[DataRow-1];
      end;
      grdMaster.Col[5].heading   :='S101';
   end
   else if CHK_S102.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := UV_vS102[DataRow-1];
      end;
      grdMaster.Col[5].heading   :='S102';
   end
   else if CHK_S104.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := UV_vS104[DataRow-1];
      end;
      grdMaster.Col[5].heading   :='S104';
   end
   else if Chk_NKCell.Checked =true  then
   begin
      case DataCol of
         1 : Value  := UV_vNo[DataRow-1];
         2 : Value  := UV_vSampNo[DataRow-1];
         3 : Value  := UV_vBUN[DataRow-1];
         4 : Value  := UV_vName[DataRow-1];
         5 : Value  := UV_vE068[DataRow-1];
      end;
      grdMaster.Col[5].heading   :='E068';
   end
   else if Rout_10.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vC020[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='C020';
   end
   else if Rout_01.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vSE01[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='SE01';
   end
   else if Rout_12.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vSE03[DataRow-1];
        end;
        grdMaster.Col[5].heading   := 'SE03';
   end
   else if Rout_02.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vSE31[DataRow-1];
        end;
        grdMaster.Col[5].heading   := 'SE31';
   end
   else if Rout_11.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS033[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S033';
   end
   else if Rout_03.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS020[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S020';
   end
   else if Rout_04.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS007[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S007';
   end
   else if Rout_05.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS008[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S008';
   end
   else if Rout_06.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS091[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S091';
   end
   else if Rout_07.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS099[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S099';
   end
   else if Rout_08.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vT042[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='T042';
   end
   else if Rout_09.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vT043[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='T043';
   end
   else if Rout_13.Checked = true then
   begin
        case DataCol of
           1 : Value  := UV_vNo[DataRow-1];
           2 : Value  := UV_vSampNo[DataRow-1];
           3 : Value  := UV_vBUN[DataRow-1];
           4 : Value  := UV_vName[DataRow-1];
           5 : Value  := UV_vS016[DataRow-1];
        end;
        grdMaster.Col[5].heading   :='S016';
   end;
 end;
procedure TfrmLD4Q36.btnQueryClick(Sender: TObject);
var index, I, iRet, iTemp,
    Sum_vS010,Sum_vS011,Sum_vS012,Sum_vT009,Sum_vT040,Sum_vT041,Sum_vT042,Sum_vT043,Sum_vS094,Sum_vS095,Sum_vS096,Sum_vS097,
    Sum_vS016,Sum_vS019,Sum_vS021,Sum_vE068,Sum_vS018,
    Sum_vS007,Sum_vS008,Sum_vSE01,Sum_vSE02,Sum_vSE07,Sum_vSE08,Sum_vTT01,Sum_vTT02, Sum_vTT03, Sum_vE069,Sum_vS020,Sum_vS091,Sum_vS099,
    Sum_vS001,Sum_vS003,Sum_vS034,Sum_vS036,Sum_vS101,Sum_vS102,Sum_vS104,
    Sum_vS002,Sum_vS004,Sum_vS035,Sum_vS037,Sum_vS071,Sum_vS080,Sum_vS083,Sum_vS084,Sum_vS098,Sum_vC031,
    Sum_vC035,Sum_vSE30,Sum_vT008,Sum_vT029,Sum_vT030,Sum_vT031,Sum_vT034,Sum_vSE31,Sum_vT039,Sum_vC020,Sum_vS033,Sum_vSE03 : integer;
    sSelect, sWhere, sGroupBy, sOrderBy : String;
    UV_fGulkwa, UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, sJangbi_List, sHul_List : String;
    UV_fGulkwa_07, UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07  : String;
    vCod_chuga : Variant;
    Check :Boolean;
    Check_05 :Boolean;
    Check_07 :Boolean;
begin
   inherited;
   sSelect := ''; sWhere := ''; sOrderBy := '';
   Index := 0;

   Sum_vS010 :=0; Sum_vS011 :=0; Sum_vS012 :=0; Sum_vT009 :=0;
   Sum_vT040 :=0; Sum_vT041 :=0; Sum_vT042 :=0; Sum_vT043 :=0;
   Sum_vS094 :=0; Sum_vS095 :=0; Sum_vS096 :=0; Sum_vS097 :=0;
   Sum_vS016 :=0; Sum_vS019 :=0; Sum_vS021 :=0; Sum_vE068 :=0;
   Sum_vS007 :=0; Sum_vS008 :=0; Sum_vSE01 :=0; Sum_vSE02 :=0;
   Sum_vSE08 :=0; Sum_vTT01 :=0; Sum_vTT02 :=0; Sum_vTT03 :=0;
   Sum_vE069 :=0; Sum_vS020 :=0; Sum_vS001 :=0; Sum_vS003 :=0;
   Sum_vS034 :=0; Sum_vS036 :=0; Sum_vS091 :=0; Sum_vS099 :=0;

   Sum_vS002:=0;
   Sum_vS004:=0;
   Sum_vS035:=0;
   Sum_vS037:=0;
   Sum_vS071:=0;
   Sum_vS080:=0;
   Sum_vS083:=0;
   Sum_vS084:=0;
   Sum_vS098:=0;
   Sum_vC031:=0;

   Sum_vC031:=0;
   Sum_vC035:=0;
   Sum_vSE07:=0;
   Sum_vSE08:=0;
   Sum_vSE30:=0;
   Sum_vT008:=0;
   Sum_vT029:=0;
   Sum_vT030:=0;
   Sum_vT031:=0;
   Sum_vT034:=0;
   Sum_vSE31:=0;
   Sum_vT039:=0;
   Sum_vS101:=0;
   Sum_vS102:=0;
   Sum_vS104:=0;
   Sum_vC020:=0;
   Sum_vS033:=0;
   Sum_vSE03:=0;

   // ��ȸ���� Check
   if not UF_Condition then exit;

   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;
   UP_VarMemSet(0);
   with qryBunju do
   begin
      // SQL�� ����
      Close;
      sSelect := ' SELECT  B.desc_rackno,B.num_Bunju, G.num_jumin, G.num_samp,G.desc_name, G.num_id, G.dat_gmgn, G.cod_jisa, G.gubn_nosin, G.gubn_adult, G.gubn_life, G.Gubn_bogun, G.Gubn_agab, G.Gubn_tkgm,  ' +
                 '         G.Gubn_nsyh, G.Gubn_adyh, G.Gubn_lfyh, G.Gubn_bgyh, G.Gubn_agyh, G.cod_jangbi, G.cod_hul, G.cod_cancer, G.cod_chuga,'+
                 '         T.cod_prf, T.cod_chuga As Tcod_chuga '+
                 '         FROM bunju B with(nolock) inner join gumgin G with(nolock) on G.cod_jisa = b.cod_jisa and g.num_id = b.num_id and g.dat_gmgn = b.dat_gmgn ' +
                 '         left outer join Tkgum T  with(nolock) ON G.cod_jisa = T.cod_jisa and G.num_jumin = T.num_jumin and G.dat_gmgn = T.dat_gmgn and G.gubn_tkgm = T.gubn_cha';




      sWhere  := ' WHERE B.Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + '''';
      sWhere  := sWhere + ' AND B.Dat_Bunju  = ''' + edtBDate.Text + ''' ';
      sWhere  := sWhere + ' and B.num_bunju in(select num_bunju from Gulkwa where Cod_bjjs = ''' + Copy(cmbJisa.Text,1,2) + ''' AND Dat_Bunju  = ''' + edtBDate.Text + ''' and (Gubn_Part  = ''05'' OR Gubn_part = ''07''))';

      if Cmb_gubn.Text = '���ֹ�ȣ' then
      begin
         if Trim(bunjuno1.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju >= ''' + bunjuno1.Text + '''';
         if Trim(bunjuno2.Text) <> '' then
            sWhere := sWhere + ' AND B.num_bunju <= ''' + bunjuno2.Text + '''';
      end
      else if Cmb_gubn.Text = '���ù�ȣ' then
      begin
         if Trim(MEdt_SampS.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp >= ''' + MEdt_SampS.Text + '''';
         if Trim(MEdt_SampE.Text) <> '' then
            sWhere := sWhere + ' AND G.num_samp <= ''' + MEdt_SampE.Text + '''';
      end
      else    // rack No.
      begin
         if Trim(Edit_RackS.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno >= ''' + Edit_RackS.Text + '''';
         if Trim(Edit_RackE.Text) <> '' then
            sWhere := sWhere + ' AND B.desc_rackno <= ''' + Edit_RackE.Text + '''';
      end;
      sOrderBy := ' ORDER BY B.Desc_Rackno ';



      SQL.Clear;
      SQL.Add(sSelect + sWhere + sOrderBy);

      Open;
      Gride.Progress := 0;
      Gride.MaxValue := RecordCount;
      Index := 0;

      if RecordCount > 0 then
      begin
         GP_query_log(qryBunju, 'LD4Q36', 'Q', 'N');
         while not qryBunju.Eof do
         Begin

            sHul_List := '';

            Gride.Progress := Gride.Progress + 1;
            Application.ProcessMessages;

            sHul_List := UF_hangmokList;

            
            if (Pos('S010' ,sHul_List)> 0) or (Pos('S011' ,sHul_List)> 0) or (Pos('S012' ,sHul_List)> 0) or (Pos('T009' ,sHul_List)> 0)or
               (Pos('T040' ,sHul_List)> 0) or (Pos('T041' ,sHul_List)> 0) or (Pos('T042' ,sHul_List)> 0) or (Pos('T043' ,sHul_List)> 0)or
               (Pos('S016' ,sHul_List)> 0) or (Pos('S019' ,sHul_List)> 0) or (Pos('S021' ,sHul_List)> 0) or (Pos('E068' ,sHul_List)> 0)or
               (Pos('S007' ,sHul_List)> 0) or (Pos('S008' ,sHul_List)> 0) or (Pos('SE01' ,sHul_List)> 0) or
               (Pos('SE07' ,sHul_List)> 0) or (Pos('SE08' ,sHul_List)> 0) or
               (Pos('S020' ,sHul_List)> 0) or (Pos('S001' ,sHul_List)> 0) or (Pos('S003' ,sHul_List)> 0)or
               (Pos('S034' ,sHul_List)> 0) or (Pos('S036' ,sHul_List)> 0) or (Pos('S091' ,sHul_List)> 0) or (Pos('S099' ,sHul_List)> 0)or

               (Pos('S002' ,sHul_List)> 0) or (Pos('S083' ,sHul_List)> 0) or (Pos('T034' ,sHul_List)> 0) or (Pos('SE31' ,sHul_List)> 0) or
               (Pos('S004' ,sHul_List)> 0) or (Pos('S084' ,sHul_List)> 0) or (Pos('SE30' ,sHul_List)> 0) or
               (Pos('S035' ,sHul_List)> 0) or (Pos('S098' ,sHul_List)> 0) or (Pos('T008' ,sHul_List)> 0) or
               (Pos('S037' ,sHul_List)> 0) or (Pos('C031' ,sHul_List)> 0) or (Pos('T029' ,sHul_List)> 0) or
               (Pos('S071' ,sHul_List)> 0) or (Pos('C035' ,sHul_List)> 0) or (Pos('T030' ,sHul_List)> 0) or
               (Pos('S080' ,sHul_List)> 0) or (Pos('SE07' ,sHul_List)> 0) or (Pos('T031' ,sHul_List)> 0) or (Pos('T039' ,sHul_List)> 0) or
               (Pos('S101' ,sHul_List)> 0) or (Pos('S102' ,sHul_List)> 0) or (Pos('S104' ,sHul_List)> 0) or (Pos('TT01' ,sHul_List)> 0) or (Pos('TT03' ,sHul_List)> 0) or
               (Pos('C020' ,sHul_List)> 0) or (Pos('S033' ,sHul_List)> 0) or (Pos('SE03' ,sHul_List)> 0)
               Then
               begin
               qryGulkwa.close;
               qryGulkwa.ParamByName('num_id').AsString    := qryBunju.FieldByName('num_id').AsString;
               qryGulkwa.ParamByName('dat_bunju').AsString := edtBDate.text;
               qryGulkwa.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('dat_gmgn').AsString;
               qryGulkwa.open;
               Check := False;
               if qryGulkwa.RecordCount > 0 then
               begin
                    while not qryGulkwa.Eof do
                    begin
                         Check_05:=false;
                         Check_07:=false;


                         if  qryGulkwa.FieldByName('gubn_part').AsString  ='05'  then
                         begin
                             UV_fGulkwa1 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                             UV_fGulkwa2 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                             UV_fGulkwa3 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                             GF_HulGulkwa('2', UV_fGulkwa1, UV_fGulkwa2, UV_fGulkwa3, UV_fGulkwa);
                             Check_05:=true;
                         end
                         else if  qryGulkwa.FieldByName('gubn_part').AsString  ='07' then
                         begin
                             UV_fGulkwa1_07 := qryGulkwa.FieldByName('DESC_GLKWA1').AsString;
                             UV_fGulkwa2_07 := qryGulkwa.FieldByName('DESC_GLKWA2').AsString;
                             UV_fGulkwa3_07 := qryGulkwa.FieldByName('DESC_GLKWA3').AsString;
                             GF_HulGulkwa('2', UV_fGulkwa1_07, UV_fGulkwa2_07, UV_fGulkwa3_07, UV_fGulkwa_07);
                             Check_07:=true;
                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('S010', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Arch.Checked=true)) and (Check_05 =True) then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 37,  6)) = '') then
                               begin
                                  UV_vS010[Index] := '�����';
                                  Sum_vS010 := Sum_vS010 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 37,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 37,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 37,  6)) <> '����')) then
                               begin
                                    UV_vS010[Index] := '' + Trim(Copy(UV_fGulkwa, 37,  6));
                                    Sum_vS010 := Sum_vS010 + 1;
                                    Check := True;
                               end;
                          end;
                         ///////////////////////////////////////////////////////////
                         if (pos('S011', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Arch.Checked=true)) and (Check_05 =True) then
                            begin
                               if (Trim(Copy(UV_fGulkwa, 25,  6)) = '') then
                               begin
                                    UV_vS011[Index] := '�����';
                                    Sum_vS011 := Sum_vS011 + 1;
                                    Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 25,  6)) <> '�缺') AND(Trim(Copy(UV_fGulkwa, 25,  6)) <> '��缺') AND  (Trim(Copy(UV_fGulkwa, 25,  6)) <> '����')) then
                               begin
                                      UV_vS011[Index] := '' + Trim(Copy(UV_fGulkwa, 25,  6));
                                      Sum_vS011 := Sum_vS011 + 1;
                                      Check := True;
                               end;
                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('S012', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Arch.Checked=true)) and (Check_05 =True) then
                            begin
                               if (Trim(Copy(UV_fGulkwa, 31,  6)) = '') then
                               begin
                                    UV_vS012[Index] := '�����';
                                    Sum_vS012 := Sum_vS012 + 1;
                                    Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 31,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 31,  6)) <> '��缺') AND(Trim(Copy(UV_fGulkwa, 31,  6)) <> '����')) then
                               begin
                                      UV_vS012[Index] := '' + Trim(Copy(UV_fGulkwa, 31,  6));
                                      Sum_vS012 := Sum_vS012 + 1;
                                      Check := True;
                               end;
                         end;
                         ///////////////////////////////////////////////////////////
                         if (pos('T040', sHul_List) > 0) and (CHK_sub.Checked=true) and (Check_05 =True) then
                         begin
                               if (Trim(Copy(UV_fGulkwa, 343,  6)) = '') then
                               begin
                                   UV_vT040[Index] := '�����';
                                   Sum_vT040 := Sum_vT040 + 1;
                                   Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 343,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 343,  6)) <> '��缺')AND (Trim(Copy(UV_fGulkwa, 343,  6)) <> '����')) then
                               begin
                                   UV_vT040[Index] := '' + Trim(Copy(UV_fGulkwa, 343,  6));
                                   Sum_vT040 := Sum_vT040 + 1;
                                   Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('T041', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Arch.Checked=true)) and (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 349,  6)) = '') then
                               begin
                                  UV_vT041[Index] := '�����';
                                  Sum_vT041 := Sum_vT041 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('T042', sHul_List) > 0) and (Check_05 =True) and
                             (((Chk_Arch.Checked=true) and (qryGulkwa.ParamByName('dat_bunju').AsString < '20140901')) or
                              ((Chk_C1_C5.Checked=true) and (qryGulkwa.ParamByName('dat_bunju').AsString >= '20140901')) or (Chk_all.Checked=true)) and
                              ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                               (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                               (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                               (Rout_12.Checked = false) and (Rout_13.Checked = false))  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 355,  6)) = '') then
                               begin
                                  UV_vT042[Index] := '�����';
                                  Sum_vT042 := Sum_vT042 + 1;
                                  Check := True;
                               end;
                          end;

                          if (pos('T042', sHul_List) > 0) and (Check_05 =True)  and  (Rout_08.checked =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 355,  6)) = '') then
                               begin
                                  UV_vT042[Index] := '�����';
                                  Sum_vT042 := Sum_vT042 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('T043', sHul_List) > 0) and (Check_05 =True) and
                             (((CHK_sub.Checked=true) and (qryGulkwa.ParamByName('dat_bunju').AsString < '20140901')) or
                              (((Chk_ALL.Checked=true) or (Chk_C1_C5.Checked=true)) and
                               ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                                (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                                (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                                (Rout_12.Checked = false) and (Rout_13.Checked = false) and (qryGulkwa.ParamByName('dat_bunju').AsString >= '20140901')))) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 361,  6)) = '') then
                               begin
                                  UV_vT043[Index] := '�����';
                                  Sum_vT043 := Sum_vT043 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T043', sHul_List) > 0) and (Check_05 =True) and (Rout_09.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 361,  6)) = '') then
                               begin
                                  UV_vT043[Index] := '�����';
                                  Sum_vT043 := Sum_vT043 + 1;
                                  Check := True;
                               end;
                          end;
                          //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                          if (pos('S007', sHul_List) > 0) and ((Chk_all.Checked=true) or  (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false)) and (Check_07 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 31,  6)) = '') then
                               begin
                                  UV_vS007[Index] := '�����' ;
                                  Sum_vS007 := Sum_vS007 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '����')) then
                               begin
                                    UV_vS007[Index] := Trim(Copy(UV_fGulkwa_07, 31,  6));
                                    Sum_vS007 := Sum_vS007 + 1;
                                    Check := True;
                               end;


                          end;
                          if (pos('S007', sHul_List) > 0) and (Rout_04.Checked = True) and (Check_07 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 31,  6)) = '') then
                               begin
                                  UV_vS007[Index] := '�����';
                                  Sum_vS007 := Sum_vS007 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa_07, 31,  6)) <> '����')) then
                               begin
                                    UV_vS007[Index] := Trim(Copy(UV_fGulkwa_07, 31,  6));
                                    Sum_vS007 := Sum_vS007 + 1;
                                    Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S008', sHul_List) > 0) and ((Chk_all.Checked=true)  or (Chk_C1_C5.Checked=true)) and
                            ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                             (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                             (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                             (Rout_12.Checked = false) and (Rout_13.Checked = false)) and
                             (Check_07 =True) and (CB01.Checked = False) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 37,  6)) = '') then
                               begin
                                  UV_vS008[Index] := '�����';
                                  Sum_vS008 := Sum_vS008 + 1;
                                  Check := True;
                               end;

                          end;
                          if (pos('S008', sHul_List) > 0) and (Rout_05.Checked = True) and (Check_07 =True)then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 37,  6)) = '') then
                               begin
                                  UV_vS008[Index] := '�����';
                                  Sum_vS008 := Sum_vS008 + 1;
                                  Check := True;
                               end;

                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('SE01', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false)) and (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 247,  6)) = '') then
                               begin
                                  UV_vSE01[Index] := '�����';
                                  Sum_vSE01 := Sum_vSE01 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 247,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 247,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 247,  6)) <> '����')) then
                               begin
                                    UV_vSE01[Index] := '' + Trim(Copy(UV_fGulkwa, 247,  6));
                                    Sum_vSE01 := Sum_vSE01 + 1;
                                    Check := True;
                                end;
                          end;
                          if (pos('SE01', sHul_List) > 0) and (Rout_01.Checked = True) and (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 247,  6)) = '') then
                               begin
                                  UV_vSE01[Index] := '�����';
                                  Sum_vSE01 := Sum_vSE01 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 247,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 247,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 247,  6)) <> '����')) then
                               begin
                                    UV_vSE01[Index] := '' + Trim(Copy(UV_fGulkwa, 247,  6));
                                    Sum_vSE01 := Sum_vSE01 + 1;
                                    Check := True;
                                end;
                          end;


                          ///////////////////////////////////////////////////////////
                          {if (pos('E069', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                           (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                           (Rout_08.Checked = false) and (Rout_09.Checked = false))and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 433,  6)) = '') then
                               begin
                                  UV_vE069[Index] := '�����';
                                  Sum_vE069 := Sum_vE069 + 1;
                                  Check := True;
                               end;
                          end;} //..ȣ��ý����� ����
                          ///////////////////////////////////////////////////////////
                          if (pos('SE31', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false)) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 445,  6)) = '') then
                               begin
                                  UV_vSE31[Index] := '�����';
                                  Sum_vSE31 := Sum_vSE31 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE31', sHul_List) > 0) and (Rout_02.Checked = True) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 445,  6)) = '') then
                               begin
                                  UV_vSE31[Index] := '�����';
                                  Sum_vSE31 := Sum_vSE31 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S020', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false) and  (Check_05 =True))  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 1,  6)) = '') then
                               begin
                                  UV_vS020[Index] := '�����';
                                  Sum_vS020 := Sum_vS020 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 1,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 1,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 1,  6)) <> '����')) then
                               begin
                                   UV_vS020[Index] := '' + Trim(Copy(UV_fGulkwa, 1,  6));
                                   Sum_vS020 := Sum_vS020 + 1;
                                   Check := True;
                               end;
                          end;
                          if (pos('S020', sHul_List) > 0) and (Rout_03.Checked = True) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 1,  6)) = '') then
                               begin
                                  UV_vS020[Index] := '�����';
                                  Sum_vS020 := Sum_vS020 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 1,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 1,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 1,  6)) <> '����')) then
                               begin
                                   UV_vS020[Index] := '' + Trim(Copy(UV_fGulkwa, 1,  6));
                                   Sum_vS020 := Sum_vS020 + 1;
                                   Check := True;
                               end;
                          end;
                          if (pos('S091', sHul_List) > 0) and ((Chk_all.Checked=true)  or  (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false))and
                             (Check_07 =True) and (CB01.Checked = False) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 181,  6)) = '') then
                               begin
                                  UV_vS091[Index] := '�����';
                                  Sum_vS091 := Sum_vS091 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S091', sHul_List) > 0) and  (Check_07 =True) and (Rout_06.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 181,  6)) = '') then
                               begin
                                  UV_vS091[Index] := '�����';
                                  Sum_vS091 := Sum_vS091 + 1;
                                  Check := True;
                               end;
                          end;

                          if (pos('C020', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false))and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                               begin
                                  UV_vC020[Index] := '�����';
                                  Sum_vC020 := Sum_vC020 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('C020', sHul_List) > 0) and  (Check_05 =True) and (Rout_10.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 67,  6)) = '') then
                               begin
                                  UV_vC020[Index] := '�����';
                                  Sum_vC020 := Sum_vC020 + 1;
                                  Check := True;
                               end;
                          end;

                          if (pos('S033', sHul_List) > 0) and ((Chk_all.Checked=true)  or  (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false))and
                            (Check_05 =True) and (CB01.Checked = False) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 457,  6)) = '') then
                               begin
                                  UV_vS033[Index] := '�����';
                                  Sum_vS033 := Sum_vS033 + 1;
                                  Check := True;                                         
                               end;
                          end;
                          if (pos('S033', sHul_List) > 0) and  (Check_05 =True) and (Rout_11.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 457,  6)) = '') then
                               begin
                                  UV_vS033[Index] := '�����';
                                  Sum_vS033 := Sum_vS033 + 1;
                                  Check := True;
                               end;
                          end;

                          if (pos('SE03', sHul_List) > 0) and ((Chk_all.Checked=true)  or  (Chk_C1_C5.Checked=true)) and
                           ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) and (Rout_13.Checked = false))and
                            (Check_05 =True) and (CB01.Checked = False) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 463,  6)) = '') then
                               begin
                                  UV_vSE03[Index] := '�����';
                                  Sum_vSE03 := Sum_vSE03 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE03', sHul_List) > 0) and  (Check_05 =True) and (Rout_12.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 463,  6)) = '') then
                               begin
                                  UV_vSE03[Index] := '�����';
                                  Sum_vSE03 := Sum_vSE03 + 1;
                                  Check := True;
                               end;
                          end;



                          if (pos('S099', sHul_List) > 0) and ((Chk_all.Checked=true) or  (Chk_C1_C5.Checked=true)) and ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                            (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                            (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                            (Rout_12.Checked = false) )and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 229,  6)) = '') then
                               begin
                                  UV_vS099[Index] := '�����';
                                  Sum_vS099 := Sum_vS099 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S016', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_C1_C5.Checked=true)) and (Check_05 =True) and
                            ((Rout_01.Checked = false) and (Rout_02.Checked = false) and (Rout_03.Checked = false) and
                             (Rout_04.Checked = false) and (Rout_05.Checked = false) and (Rout_06.Checked = false) and (Rout_07.Checked = false) and
                             (Rout_08.Checked = false) and (Rout_09.Checked = false) and (Rout_10.Checked = false) and (Rout_11.Checked = false) and
                             (Rout_12.Checked = false) and (Rout_13.Checked = false))  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                               begin
                                  UV_vS016[Index] := '�����';
                                  Sum_vS016 := Sum_vS016 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 31,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '����')) then
                               begin
                                    UV_vS016[Index] := '' + Trim(Copy(UV_fGulkwa, 7,  6));
                                    Sum_vS016 := Sum_vS016 + 1;
                                    Check := True;
                               end;
                          end;
                          if (pos('S016', sHul_List) > 0) and  (Check_05 =True) and (Rout_13.Checked = True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 7,  6)) = '') then
                               begin
                                  UV_vS016[Index] := '�����';
                                  Sum_vS016 := Sum_vS016 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 31,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 7,  6)) <> '����')) then
                               begin
                                    UV_vS016[Index] := '' + Trim(Copy(UV_fGulkwa, 7,  6));
                                    Sum_vS016 := Sum_vS016 + 1;
                                    Check := True;
                               end;
                          end;


                          if (pos('S099', sHul_List) > 0) and  (Rout_07.Checked = True) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 229,  6)) = '') then
                               begin
                                  UV_vS099[Index] := '�����';
                                  Sum_vS099 := Sum_vS099 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S001', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Toshiba.Checked=true)) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 7,  6)) = '') then
                               begin
                                  UV_vS001[Index] := '�����';
                                  Sum_vS001 := Sum_vS001 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S003', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Toshiba.Checked=true)) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 19,  6)) = '') then
                               begin
                                  UV_vS003[Index] := '�����';
                                  Sum_vS003 := Sum_vS003 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa_07, 19,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa_07, 19,  6)) <> '��缺') AND  (Trim(Copy(UV_fGulkwa_07, 19,  6)) <> '����')) then
                               begin
                                   UV_vS003[Index] := '' + Trim(Copy(UV_fGulkwa_07, 19,  6));
                                   Sum_vS003 := Sum_vS003 + 1;
                                   Check := True;
                               end;
                          end;

                          ///////////////////////////////////////////////////////////
                          if (pos('S034', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Toshiba.Checked=true)) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 43,  6)) = '') then
                               begin
                                  UV_vS034[Index] := '�����';
                                  Sum_vS034 := Sum_vS034 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa_07, 43,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa_07, 43,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa_07, 43,  6)) <> '����')) then
                               begin
                                   UV_vS034[Index] := '*' + Trim(Copy(UV_fGulkwa_07, 43,  6));
                                   Sum_vS034 := Sum_vS034 + 1;
                                   Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S036', sHul_List) > 0) and  ((Chk_all.Checked=true) or (Chk_Toshiba.Checked=true)) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 55,  6)) = '') then
                               begin
                                  UV_vS036[Index] := '�����';
                                  Sum_vS036 := Sum_vS036 + 1;
                                  Check := True;
                               end;
                          end;
                          //****//
                          ///////////////////////////////////////////////////////////
                          if (pos('S002', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 13,  6)) = '') then
                               begin
                                  UV_vS002[Index] := '�����';
                                  Sum_vS002 := Sum_vS002 + 1;
                                  Check := True;
                               end;
                          end;
                          {
                          if (pos('S004', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 25,  6)) = '') then
                               begin
                                  UV_vS004[Index] := '�����';
                                  Sum_vS004 := Sum_vS004 + 1;
                                  Check := True;
                               end;
                          end;
                          }
                          ///////////////////////////////////////////////////////////     20150211
                          if (pos('S004', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Toshiba.Checked=true)) and (Check_07 =True) and (CB02.Checked = False) then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 25,  6)) = '') then
                               begin
                                  UV_vS004[Index] := '�����';
                                  Sum_vS004 := Sum_vS004 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa_07, 25,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa_07, 25,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa_07, 25,  6)) <> '����')) then
                               begin
                                   UV_vS004[Index] := '*' + Trim(Copy(UV_fGulkwa_07, 25,  6));
                                   Sum_vS004 := Sum_vS004 + 1;
                                   Check := True;
                               end;
                          end;
                          ///////////////////////////////////////////////////////////
                          if (pos('S035', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 49,  6)) = '') then
                               begin
                                  UV_vS035[Index] := '�����';
                                  Sum_vS035 := Sum_vS035 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S037', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 61,  6)) = '') then
                               begin
                                  UV_vS037[Index] := '�����';
                                  Sum_vS037 := Sum_vS037 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S071', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 127,  6)) = '') then
                               begin
                                  UV_vS071[Index] := '�����';
                                  Sum_vS071 := Sum_vS071 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S080', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 91,  6)) = '') then
                               begin
                                  UV_vS080[Index] := '�����';
                                  Sum_vS080 := Sum_vS080 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S083', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 145,  6)) = '') then
                               begin
                                  UV_vS083[Index] := '�����';
                                  Sum_vS083 := Sum_vS083 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S084', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 151,  6)) = '') then
                               begin
                                  UV_vS084[Index] := '�����';
                                  Sum_vS084 := Sum_vS084 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S098', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_07 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa_07, 223,  6)) = '') then
                               begin
                                  UV_vS098[Index] := '�����';
                                  Sum_vS098 := Sum_vS098 + 1;
                                  Check := True;
                               end;
                          end;
                          ///////////
                          if (pos('C031', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 61,  6)) = '') then
                               begin
                                  UV_vC031[Index] := '�����';
                                  Sum_vC031 := Sum_vC031 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('C035', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 85,  6)) = '') then
                               begin
                                  UV_vC035[Index] := '�����';
                                  Sum_vC035 := Sum_vC035 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE07', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 13,  6)) = '') then
                               begin
                                  UV_vSE07[Index] := '�����';
                                  Sum_vSE07 := Sum_vSE07 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE08', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 19,  6)) = '') then
                               begin
                                  UV_vSE08[Index] := '�����';
                                  Sum_vSE08 := Sum_vSE08 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE30', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 439,  6)) = '') then
                               begin
                                  UV_vSE30[Index] := '�����';
                                  Sum_vSE30 := Sum_vSE30 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T008', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 103,  6)) = '') then
                               begin
                                  UV_vT008[Index] := '�����';
                                  Sum_vT008 := Sum_vT008 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T029', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 199,  6)) = '') then
                               begin
                                  UV_vT029[Index] := '�����';
                                  Sum_vT029 := Sum_vT029 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T030', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 205,  6)) = '') then
                               begin
                                  UV_vT030[Index] := '�����';
                                  Sum_vT030 := Sum_vT030 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T031', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 211,  6)) = '') then
                               begin
                                  UV_vT031[Index] := '�����';
                                  Sum_vT031 := Sum_vT031 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T034', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 229,  6)) = '') then
                               begin
                                  UV_vT034[Index] := '�����';
                                  Sum_vT034 := Sum_vT034 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('SE02', sHul_List) > 0) and (Chk_Not_used.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 253,  6)) = '') then
                               begin
                                  UV_vSE02[Index] := '�����';
                                  Sum_vSE02 := Sum_vSE02 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S101', sHul_List) > 0) and (Chk_S101.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 469,  6)) = '') then
                               begin
                                  UV_vS101[Index] := '�����';
                                  Sum_vS101 := Sum_vS101 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S102', sHul_List) > 0) and (Chk_S102.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 475,  6)) = '') then
                               begin
                                  UV_vS102[Index] := '�����';
                                  Sum_vS102 := Sum_vS102 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('S104', sHul_List) > 0) and (Chk_S104.Checked=true) and (Check_05 =True)  then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 487,  6)) = '') then
                               begin
                                  UV_vS104[Index] := '�����';
                                  Sum_vS104 := Sum_vS104 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('TT03', sHul_List) > 0) and ((Chk_all.Checked=true) or (Chk_Arch.Checked=true)) and (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 451,  6)) = '') then
                               begin
                                  UV_vTT03[Index] := '�����';
                                  Sum_vTT03 := Sum_vTT03 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('E068', sHul_List) > 0) and (Chk_NKCell.checked = true) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 427,  6)) = '') then
                               begin
                                  UV_vE068[Index] := '�����';
                                  Sum_vE068 := Sum_vE068 + 1;
                                  Check := True;
                               end;
                          end;
                           ///////////////////////////////////////////////////////////
                          if (pos('S021', sHul_List) > 0) and (Chk_all.Checked=true) and (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 43,  6)) = '') then
                               begin
                                  UV_vS021[Index] := '�����';
                                  Sum_vS021 := Sum_vS021 + 1;
                                  Check := True;
                               end
                               else if ((Trim(Copy(UV_fGulkwa, 43,  6)) <> '�缺') AND (Trim(Copy(UV_fGulkwa, 43,  6)) <> '��缺') AND (Trim(Copy(UV_fGulkwa, 43,  6)) <> '����')) then
                               begin
                                  UV_vS021[Index] := '' + Trim(Copy(UV_fGulkwa, 43,  6));
                                  Sum_vS021 := Sum_vS021 + 1;
                                  Check := True;
                               end;
                          end;
                          if (pos('T039', sHul_List) > 0) and (Chk_all.Checked=true) and (qryGulkwa.ParamByName('dat_bunju').AsString >= '20140911') and
                             (Check_05 =True) then
                          begin
                               if (Trim(Copy(UV_fGulkwa, 313,  6)) = '') then
                               begin
                                  UV_vT039[Index] := '�����';
                                  Sum_vT039 := Sum_vT039 + 1;
                                  Check := True;
                               end;
                          end;

                   qryGulkwa.Next;
                   end;
                   if Check = True then
                   begin
                        UV_vNo[Index]     := qryBunju.FieldByName('desc_RackNo').AsString;
                        UV_VBUN[Index]    := qryBunju.FieldByName('Num_Bunju').AsString;
                        UV_vName[Index]   := qryBunju.FieldByName('desc_name').AsString;
                        UV_vSampNo[Index] := qryBunju.FieldByName('num_samp').AsString;
                        UP_VarMemSet(Index+1);
                        inc(index);
                   end;
               End;
               qryGulkwa.Close;
              end;
           qryBunju.Next;
         End;


         // Table�� Disconnect
         qryBunju.Close;

         UV_vNo[Index]     := '';
         UV_vJisa[Index]   := '';
         UV_vBUN[Index]    := '';
         UV_vName[Index]   := '';
         UV_vJumin[Index]  := '��  ��  ��';
         UV_vS010[Index]:= Sum_vS010;
         UV_vS011[Index]:= Sum_vS011;
         UV_vS012[Index]:= Sum_vS012;
         UV_vT009[Index]:= Sum_vT009;
         UV_vT040[Index]:= Sum_vT040;
         UV_vT041[Index]:= Sum_vT041;
         UV_vT042[Index]:= Sum_vT042;
         UV_vT043[Index]:= Sum_vT043;
         UV_vS094[Index]:= Sum_vS094;
         UV_vS095[Index]:= Sum_vS095;
         UV_vS096[Index]:= Sum_vS096;
         UV_vS097[Index]:= Sum_vS097;
         UV_vS016[Index]:= Sum_vS016;
         UV_vS019[Index]:= Sum_vS019;
         UV_vS021[Index]:= Sum_vS021;
         UV_vS018[Index]:= Sum_vs018;
         UV_vS007[Index]:= Sum_vS007;
         UV_vS008[Index]:= Sum_vS008;
         UV_vSE01[Index]:= Sum_vSE01;
         UV_vSE02[Index]:= Sum_vSE02;
         UV_vSE07[Index]:= Sum_vSE07;
         UV_vSE08[Index]:= Sum_vSE08;
         UV_vTT01[Index]:= Sum_vTT01;
         UV_vTT02[Index]:= Sum_vTT02;
         UV_vTT03[Index]:= Sum_vTT03;
         UV_vE068[Index]:= Sum_vE068;
         UV_vE069[Index]:= Sum_vE069;
         UV_vS020[Index]:= Sum_vS020;
         UV_vS001[Index]:= Sum_vS001;
         UV_vS003[Index]:= Sum_vS003;
         UV_vS034[Index]:= Sum_vS034;
         UV_vS036[Index]:= Sum_vS036;
         UV_vS091[Index]:= Sum_vS091;
         UV_vS099[Index]:= Sum_vS099;


         Uv_vS002[Index]:= Sum_vS002;
         Uv_vS004[Index]:= Sum_vS004;
         Uv_vS035[Index]:= Sum_vS035;
         Uv_vS037[Index]:= Sum_vS037;
         Uv_vS071[Index]:= Sum_vS071;
         Uv_vS080[Index]:= Sum_vS080;
         Uv_vS083[Index]:= Sum_vS083;
         Uv_vS084[Index]:= Sum_vS084;
         Uv_vS098[Index]:= Sum_vS098;
         Uv_vC031[Index]:= Sum_vC031;
         Uv_vC035[Index]:= Sum_vC035;
         Uv_vSE07[Index]:= Sum_vSE07;
         Uv_vSE08[Index]:= Sum_vSE08;
         Uv_vSE30[Index]:= Sum_vSE30;
         Uv_vT008[Index]:= Sum_vT008;
         Uv_vT029[Index]:= Sum_vT029;
         Uv_vT030[Index]:= Sum_vT030;
         Uv_vT031[Index]:= Sum_vT031;
         Uv_vT034[Index]:= Sum_vT034;
         Uv_vSE31[Index]:= Sum_vSE31;
         Uv_vT039[Index]:= Sum_vT039;
         Uv_vS101[Index]:= Sum_vS101;
         Uv_vS102[Index]:= Sum_vS102;
         Uv_vS104[Index]:= Sum_vS104;
         Uv_vC020[Index]:= Sum_vC020;
         Uv_vS033[Index]:= Sum_vS033;
         Uv_vSE03[Index]:= Sum_vSE03;

         UV_vSampNo[Index]  := '';
         Inc(index);
      end
      else
      begin
         // �ڷᰡ ������ ǥ��
         GF_MsgBox('Information', 'NODATA');
         exit;
      end;  // if RecordCount > 0 then
//      Close;
   end;  // with qry_gulkwa do

   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;

   // Query Mode & Msg
   if Sender = btnQuery then UP_SetMode('QUERY');
end;

procedure TfrmLD4Q36.UP_Click(Sender: TObject);
var sCode, sName : String;
begin
  inherited;
   if Sender = Chk_C1_C5 then GroupRadio2.Visible := True
   ELSE
   begin
   GroupRadio2.Visible := False ;
   Rout_01.checked := False;
   Rout_02.checked := False;
   Rout_03.checked := False;
   Rout_04.checked := False;
   Rout_05.checked := False;
   Rout_06.checked := False;
   Rout_07.checked := False;
   Rout_08.checked := False;
   Rout_09.checked := False;
   Rout_10.checked := False;
   Rout_11.checked := False;
   Rout_12.checked := False;
   Rout_13.checked := False;
   end;
   if Sender = btnBdate then GF_CalendarClick(edtBDate);
end;
procedure TfrmLD4Q36.btnPrintClick(Sender: TObject);
begin
  inherited;
  frmLD4Q361 := TfrmLD4Q361.Create(Self);
  if RBtn_preveiw.Checked = True then frmLD4Q361.QuickRep.Preview
  else                                frmLD4Q361.QuickRep.Print;

end;
procedure TfrmLD4Q36.FormActivate(Sender: TObject);
begin
   inherited;
   Cmb_gubn.ItemIndex := 1;
   edtBdate.SetFocus;


end;

function TfrmLD4Q36.UF_hangmokList : String;
var
   i ,k, j : integer;
   sTemp, sSelect, sWhere1, sWhere2, sOderby, sEtc_Hangmok_hm, sTk_Hangmok_Pf, sTk_Hangmok_hm, sTotal_HangmokList : string;
       sHmCode :string;
   vCod_chuga : Variant;
   iRet : integer;
begin
   with qryPf_hm do
   begin
      // 1. ���׿� ���� �˻��׸� ����
      if qryBunju.FieldByName('cod_hul').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('cod_hul').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 2. ���翡 ���� �˻��׸� ����
      if qryBunju.FieldByName('Cod_cancer').AsString <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_cancer').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;

      // 3. ��� ���� �˻��׸� ����
      if qryBunju.FieldByName('Cod_jangbi').AsString  <> '' then
      begin
         Active := False;
         ParamByName('COD_PF').AsString := qryBunju.FieldByName('Cod_jangbi').AsString;
         Active := True;

         if RecordCount > 0 then
         begin
            while not Eof do
            begin
               if copy(FieldByName('COD_HM').AsString,1,2) <> 'JJ' then
               sTemp := sTemp + FieldByName('COD_HM').AsString;
               Next;
            end;
         end;
      end;
      Active := False;
   end;

   // 3. �߰��ڵ忡 ���� �˻��׸� ����
   sTemp := sTemp + qryBunju.FieldByName('Cod_chuga').AsString;;

   // 4. �׸��ڵ忡 ���� �˻��׸� ����
   sHmCode := '';
   if qryBunju.FieldByName('Gubn_nosin').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '1', qryBunju.FieldByName('Gubn_nsyh').AsInteger)
   else if qryBunju.FieldByName('Gubn_adult').AsString = '1' then
      sHmCode := UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '4', qryBunju.FieldByName('Gubn_adyh').AsInteger);

   if qryBunju.FieldByName('Gubn_life').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '7', qryBunju.FieldByName('Gubn_lfyh').AsInteger);

   if qryBunju.FieldByName('Gubn_bogun').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '3',qryBunju.FieldByName('Gubn_bgyh').AsInteger);

   if qryBunju.FieldByName('Gubn_agab').AsString = '1' then
      sHmCode := sHmCode + UF_No_Hangmok(qryBunju.FieldByName('Dat_gmgn').AsString, '5',qryBunju.FieldByName('Gubn_agyh').AsInteger);

   if (qryBunju.FieldByName('Gubn_tkgm').AsString = '1') or (qryBunju.FieldByName('Gubn_tkgm').AsString = '2') then
   begin
      qryTkgum.Active := False;
      qryTkgum.ParamByName('cod_jisa').AsString  := qryBunju.FieldByName('Cod_jisa').AsString;
      qryTkgum.ParamByName('num_jumin').AsString := qryBunju.FieldByName('Num_jumin').AsString;
      qryTkgum.ParamByName('dat_gmgn').AsString  := qryBunju.FieldByName('Dat_gmgn').AsString;
      qryTkgum.Active := True;
      if qryTkgum.RecordCount > 0 then
      begin
         I := Length(qryTkgum.FieldByName('Cod_prf').AsString);
         J := Round(I / 4);
         For K := 0 To J - 1 Do
         Begin
           With qryProfile do
           Begin
              Close;
              ParamByName('In_Cod_Pf').AsString := Copy(qryTkgum.FieldByName('Cod_Prf').AsString,K * 4 + 1, 4);
              Open;
              For I := 1 To Recordcount Do
              Begin
                 if pos(FieldByName('Cod_hm').AsString, sHmCode) = 0 then
                    sHmCode := sHmCode + FieldByName('Cod_hm').AsString;
                 Next;
              End;
              Close;
            end;
         end;
         sHmCode := sHmCode + qryTkgum.FieldByName('cod_chuga').AsString;
      end;
      qryTkgum.Active := False;
   end;
   if Trim(sHmCode) <> '' then
   begin
      iRet := GF_MulToSingle(sHmCode, 4, vCod_chuga);
      for i := 0 to iRet-1 do
      begin
         if copy(vCod_chuga[i],1,2) <> 'JJ' then
         sTemp := sTemp + vCod_chuga[i];
      end;
   end;

   for k:=0 to  Round(Length(sTemp)/4)-1 do
   begin
   result:=result+Copy(sTemp, (k+1)*4-3, 4)+'.';
   end;

end;

function TfrmLD4Q36.UF_No_Hangmok(sDate, sGubun : String; iYh : Integer): String;
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
      end;
      Active := False;
   end;
end;

procedure TfrmLD4Q36.SBut_ExcelClick(Sender: TObject);
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
      Application.MessageBox('Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.',
                             '����', MB_OK or MB_ICONERROR);
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


      XL.DisplayAlerts := True;
      XL.Visible:= True;
   finally
      Application.ProcessMessages;
      Screen.Cursor  := crDefault;
   end;
end;

end.
