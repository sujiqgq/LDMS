//==============================================================================
// 프로그램 설명 : 조직라벨지출력
// 시스템        : 통합검진
// 수정일자      : 2006.11.26
// 수정자        : 유동구
// 수정내용      :
// 참고사항      :
//==============================================================================
unit LD8P04;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORPrint, ExtCtrls, StdCtrls, Buttons, Spin, ValEdit, Mask, DateEdit, Db,
  DBTables, Grids_ts, TSGrid;

type
  TfrmLD8P04 = class(TfrmPrint)
    Label9: TLabel;
    RBtn_bio: TRadioButton;
    RBtn_pap: TRadioButton;
    ValEdit_Start: TValEdit;
    SpinEdit1: TSpinEdit;
    Label4: TLabel;
    grdMaster: TtsGrid;
    Btn_Insert: TBitBtn;
    Bevel2: TBevel;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    Label5: TLabel;
    ValEdit1: TValEdit;
    Label7: TLabel;
    ValEdit3: TValEdit;
    Label8: TLabel;
    ValEdit4: TValEdit;
    Btn_Delete: TBitBtn;
    RBtn_NG: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure UP_ReportClick(Sender: TObject);
    procedure Btn_InsertClick(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure ValEdit1Change(Sender: TObject);
    procedure ValEdit2Change(Sender: TObject);
    procedure Btn_DeleteClick(Sender: TObject);

  private
    { Private declarations }
    UV_vSeq, UV_vNum, UV_vGubn, UV_vPCaption : Variant;

    UV_Index : Integer;
    // 공통적으로 사용하는 함수(이름을 동일하게 사용)
    procedure UP_VarMemSet(iLength : Integer);

    public
    // 지사코드
    UV_sCod_jisa : String;

    // SQL문 저장
    UV_sBasicSQL : String;

    { Public declarations }
  end;

var
  frmLD8P04: TfrmLD8P04;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc, LD8P041;

procedure TfrmLD8P04.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      // Variant변수를 사용하기 위해서 생성
      UV_vSeq      := VarArrayCreate([0, 0], varOleStr);
      UV_vNum      := VarArrayCreate([0, 0], varOleStr);
      UV_vGubn     := VarArrayCreate([0, 0], varOleStr);
      UV_vPCaption := VarArrayCreate([0, 0], varOleStr);
  end
   else
   begin
      // 이미 생성된 변수들의 크기를 조정
      VarArrayRedim(UV_vSeq,      iLength);
      VarArrayRedim(UV_vNum,      iLength);
      VarArrayRedim(UV_vGubn,     iLength);
      VarArrayRedim(UV_vPCaption, iLength);
   end;
end;

procedure TfrmLD8P04.FormCreate(Sender: TObject);
begin
   inherited;
   UP_VarMemSet(0);
   grdMaster.Rows := 0;
   UV_Index := 0;
   SpinEdit1.Text := Copy(Gv_sToday,1,4);
end;

procedure TfrmLD8P04.UP_ReportClick(Sender: TObject);
var
   sWhere : string;
   index, i  : integer;
begin
   inherited;
   // 조회조건 Check
   if not GF_NotNullCheck(pnlPrint) then exit;

   // Report Form Create
   frmLD8P041 := TfrmLD8P041.Create(Self);

   if rbAll.Checked then
   begin
      frmLD8P041.QuickRep.PrinterSettings.FirstPage := 0;
      frmLD8P041.QuickRep.PrinterSettings.LastPage  := 0;
   end
   else if rbPart.Checked then
   begin
      frmLD8P041.QuickRep.PrinterSettings.FirstPage := StrToInt(valFrom.Text);
      frmLD8P041.QuickRep.PrinterSettings.LastPage  := StrToInt(valTo.Text);
   end;

   // 인쇄매수 설정
   frmLD8P041.QuickRep.PrinterSettings.Copies := spnCopy.Value;

   // Preview or Print
   if Sender = btnPreview then frmLD8P041.QuickRep.Preview
   else if Sender = btnPrint then frmLD8P041.QuickRep.Print;
end;


procedure TfrmLD8P04.Btn_InsertClick(Sender: TObject);
var
   i : Integer;
   sGubn :String;
begin
   inherited;
   if ValEdit_Start.Value = 0 then
   begin
      ShowMessage('샘플시작번호를 입력후 등록하여 주십시요..!');
      exit;
   end;

   if (ValEdit1.Value <> 0) and (ValEdit3.Value = 0) then
   begin
      for i := 0 to StrToInt(FloatToStr(ValEdit1.Value)) - 1 do
      begin
         UP_VarMemSet(UV_Index);
         UV_vSeq[UV_Index]      := UV_Index + 1;
         UV_vNum[UV_Index]      := ValEdit_Start.Value + StrToFloat(IntToStr(i));
         UV_vGubn[UV_Index]     := '';
         if RBtn_bio.Checked then
         begin
            UV_vPCaption[UV_Index] := 'T' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value + StrToFloat(IntToStr(i))),6,'0');
         end
         else if RBtn_pap.Checked then
         begin
            UV_vPCaption[UV_Index] := 'C' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value + StrToFloat(IntToStr(i))),6,'0');
         end
         else if RBtn_NG.Checked then
         begin
            UV_vPCaption[UV_Index] := 'NG' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value + StrToFloat(IntToStr(i))),6,'0');
         end;

         Inc(UV_Index);
      end;
   end
   else
   begin
      for i := 1 to StrToInt(FloatToStr(ValEdit3.Value)) do
      begin
         sGubn := '';
         UP_VarMemSet(UV_Index);
         UV_vSeq[UV_Index]      := UV_Index + 1;
         UV_vNum[UV_Index]      := ValEdit_Start.Value;
         case i of
            1 : sGubn := 'A';
            2 : sGubn := 'B';
            3 : sGubn := 'C';
            4 : sGubn := 'D';
            5 : sGubn := 'E';
            6 : sGubn := 'F';
            7 : sGubn := 'G';
            8 : sGubn := 'H';
            9 : sGubn := 'I';
           10 : sGubn := 'J';
           11 : sGubn := 'K';
           12 : sGubn := 'L';
           13 : sGubn := 'M';
           14 : sGubn := 'N';
           15 : sGubn := 'O';
           16 : sGubn := 'P';
           17 : sGubn := 'Q';
           18 : sGubn := 'R';
           19 : sGubn := 'S';
           20 : sGubn := 'T';
           21 : sGubn := 'U';
           22 : sGubn := 'V';
           23 : sGubn := 'W';
           24 : sGubn := 'X';
           25 : sGubn := 'Y';
           26 : sGubn := 'Z';
         end;

         UV_vGubn[UV_Index] := sGubn;

         if RBtn_bio.Checked then
         begin
            UV_vPCaption[UV_Index] := 'T' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value),6,'0');
         end
         else if RBtn_pap.Checked then
         begin
            UV_vPCaption[UV_Index] := 'C' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value),6,'0');
         end
         else if RBtn_NG.Checked then
         begin
            UV_vPCaption[UV_Index] := 'NG' + copy(IntToStr(SpinEdit1.Value),3,2) + '-' + GF_LPad(FloatToStr(ValEdit_Start.Value),6,'0');
         end;

         Inc(UV_Index);
      end;
   end;
   
   if (ValEdit3.Value > 0) then
      ValEdit_Start.Value := ValEdit_Start.Value + 1
   else if ValEdit1.Value > 0 then
      ValEdit_Start.Value := ValEdit_Start.Value + StrToFloat(IntToStr(i));

   ValEdit4.Value := StrToFloat(IntToStr(UV_Index));

   // Grid에 자료를 할당
   grdMaster.Rows := UV_Index;

   ValEdit1.Value := 0;
//   ValEdit2.Value := 0;
   ValEdit3.Value := 0;
end;

procedure TfrmLD8P04.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   inherited;

   // 자룔를 화면에 조회
   case DataCol of
      1 : Value := UV_vSeq[DataRow-1];
      2 : Value := UV_vNum[DataRow-1];
      3 : Value := UV_vGubn[DataRow-1];
      4 : Value := UV_vPCaption[DataRow-1];
   end;
end;

procedure TfrmLD8P04.ValEdit1Change(Sender: TObject);
begin
   inherited;
   if ValEdit1.Value >= 0 then
   begin
//      ValEdit2.Value := 0;
      ValEdit3.Value := 0;
   end;
end;

procedure TfrmLD8P04.ValEdit2Change(Sender: TObject);
begin
   inherited;
 {  if ValEdit2.Value >= 0 then
   begin
      ValEdit1.Value := 0;
   end;        }
end;

procedure TfrmLD8P04.Btn_DeleteClick(Sender: TObject);
var
   i, ii, j : Integer;
begin
   inherited;

   // 자료가 존재하는지 Check
   if grdMaster.Rows = 0 then exit;

   // Confirm Message
   if not GF_MsgBox('Confirmation', '샘플번호를 삭제하시겠습니까 ?') then exit;

   // 선택된 모든 항목을 삭제
   with grdMaster do
   begin
      i := 1;
      while Rows > 0 do
      begin
         if i > Rows then break;

         if RowSelected[i] then
         begin
            for j := (i-1) to rows-2 do
            begin
               UV_vSeq[j]      := UV_vSeq[j+1];
               UV_vNum[j]      := UV_vNum[j+1];
               UV_vGubn[j]     := UV_vGubn[j+1];
               UV_vPCaption[j] := UV_vPCaption[j+1];
            end;
            DeleteRows(i, i);
            UV_Index := UV_Index - 1;
            ValEdit4.Value := StrToFloat(IntToStr(UV_Index));

         end
         else
            Inc(i);
      end;
   end;
end;

end.
