//==============================================================================
// 프로그램 설명 : 2015년연세대혈액대장(금연,호르몬)
// 시스템        : 통합검진
// 수정일자      : 2015.04.06
// 수정자        : 김창규
// 수정내용      : 추가..
// 참고사항      :
//==============================================================================
unit LD7P03;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables;

type
  TfrmLD7P03 = class(TfrmPrint)
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
  frmLD7P03: TfrmLD7P03;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,
     LD7P031;

procedure TfrmLD7P03.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskST_Date);

end;

procedure TfrmLD7P03.UP_KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
begin
   inherited;

   if Key <> GC_Refer then exit;

   if Sender = mskST_Date then UP_Click(btnDate)
end;

procedure TfrmLD7P03.FormCreate(Sender: TObject);
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

procedure TfrmLD7P03.UP_ReportClick(Sender: TObject);
var
   sWhere : string;
   index, i  : integer;
begin
   inherited;

   // 조회조건 Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD7P031 := TfrmLD7P031.Create(Self);

   with frmLD7P031.qryGulkwa do
   begin
      Active := False;
      ParambyName('dat_gmgn').AsString := mskSt_date.Text;
      Active := True;
      if RecordCount > 0 then
      begin
         GP_query_log(frmLD7P031.qryGulkwa, 'LD7P03', 'Q', 'Y');
         if rbAll.Checked then
         begin
            frmLD7P031.QuickRep.PrinterSettings.FirstPage := 0;
            frmLD7P031.QuickRep.PrinterSettings.LastPage  := 0;
         end
         else if rbPart.Checked then
         begin
            frmLD7P031.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
            frmLD7P031.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
         end;
         // 인쇄매수 설정
         frmLD7P031.QuickRep.PrinterSettings.Copies := spnCopy.Value;

      // Preview or Print
      if Sender = btnPreview then frmLD7P031.QuickRep.Preview
      else if Sender = btnPrint then frmLD7P031.QuickRep.Print;

      end;
  end; // end of with;
  
end;


end.
