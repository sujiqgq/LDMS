//==============================================================================
// 프로그램 설명 : 혈액샘플대장(연세대연구목적)
// 시스템        : 통합검진
// 수정일자      : 2007.01.25
// 수정자        : 유동구
// 수정내용      : 추가..
// 참고사항      :
//==============================================================================
unit LD7P01;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD7P01 = class(TfrmPrint)
    cmbJisa: TComboBox;
    mskSt_date: TDateEdit;
    btnDate: TSpeedButton;
    Label4: TLabel;
    Label6: TLabel;
    qrySaler: TQuery;
    Query1: TQuery;
    procedure UP_Click(Sender: TObject);
    procedure UP_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure UP_ReportClick(Sender: TObject);

  private
    { Private declarations }
     public


    // 지사코드
    UV_sCod_jisa : String;
    // SQL문 저장
   UV_sBasicSQL : String;

// Procedure UP_gulkwa;
 //procedure UP_Select;
    { Public declarations }
  end;

var
  frmLD7P01: TfrmLD7P01;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,
     LD7P011;

procedure TfrmLD7P01.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date);

end;

procedure TfrmLD7P01.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate)
end;

procedure TfrmLD7P01.FormCreate(Sender: TObject);
var sJisa, sName : String;
begin
   inherited;
   GP_ComboJisa(cmbJisa);

   // 사용자권한에 따라서 지사Combo 활성화 조절
   if GV_sJisa = '00' then
   begin
      cmbJisa.Enabled := True;
      sJisa := copy(GV_sUserId,1,2);
   end
   else
   begin
      cmbJisa.Enabled := False;
      sJisa := GV_sJisa;
   end;

   // Default로 현재지사를 설정
   GP_ComboDisplay(cmbJisa, sJisa, 2);


   // 오늘일자 설정
   mskST_Date.Text := GV_sToday;
end;

procedure TfrmLD7P01.UP_ReportClick(Sender: TObject);
var
   sWhere : string;
   index, i  : integer;
begin
   inherited;

   // 조회조건 Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD7P011 := TfrmLD7P011.Create(Self);

   with frmLD7P011.qryGulkwa do
   begin
      Active := False;
      ParambyName('cod_bjjs').AsString := copy(cmbJisa.Text,1,2);
      ParambyName('dat_gmgn').AsString := mskSt_date.Text;
      Active := True;
      if RecordCount > 0 then
      begin
         GP_query_log(frmLD7P011.qryGulkwa, 'LD7P01', 'Q', 'Y');
         if rbAll.Checked then
         begin
            frmLD7P011.QuickRep.PrinterSettings.FirstPage := 0;
            frmLD7P011.QuickRep.PrinterSettings.LastPage  := 0;
         end
         else if rbPart.Checked then
         begin
            frmLD7P011.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
            frmLD7P011.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
         end;
         // 인쇄매수 설정
         frmLD7P011.QuickRep.PrinterSettings.Copies := spnCopy.Value;

      // Preview or Print
      if Sender = btnPreview then frmLD7P011.QuickRep.Preview
      else if Sender = btnPrint then frmLD7P011.QuickRep.Print;

      end;
  end; // end of with;
  
end;


end.
