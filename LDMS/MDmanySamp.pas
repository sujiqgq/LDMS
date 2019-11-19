unit MDmanySamp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Db, DBTables, ComCtrls;

type
  TfrmMDmanySamp = class(TForm)
    Label5: TLabel;
    edtSamp: TEdit;
    edtDesc_Name: TEdit;
    listSamp: TListBox;
    qrySamp: TQuery;
    BitBtn2: TBitBtn;
    btnSamp_add: TBitBtn;
    btnSamp_delete: TBitBtn;

    procedure edtSampKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtn2Click(Sender: TObject);
    procedure edtSampKeyPress(Sender: TObject; var Key: Char);
    procedure btnSamp_addClick(Sender: TObject);
    procedure btnSamp_deleteClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMDmanySamp: TfrmMDmanySamp;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc,LD8P02;


procedure TfrmMDmanySamp.edtSampKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  if  length(EdtSamp.Text)=6 then
  begin
  With qrySamp Do
       Begin
            Close;
            ParamByName('Dat_Gmgn').AsString := frmLD8P02.mskSt_date.Text;
            ParamByName('num_samp').AsString := edtSamp.Text;
            Open;
            edtDesc_Name.Text := FieldByName('desc_name').AsString;
            Close;
       End;
  end;
end;

procedure TfrmMDmanySamp.BitBtn2Click(Sender: TObject);
begin
   inherited;
   close;
end;

procedure TfrmMDmanySamp.btnSamp_deleteClick(Sender: TObject);
var i : Integer;
begin
   inherited;
   // Confirm Message
   if not GF_MsgBox('Confirmation', '샘플번호를 삭제 하시겠습니까  ?') then exit;

   i := listSamp.items.count;
   while i > 0 do
   begin
      i := i - 1;
      if ListSamp.selected[i] = true then
         ListSamp.Items.Delete(i);
   end;
end;

procedure TfrmMDmanySamp.btnSamp_addClick(Sender: TObject);
begin
   listSamp.items.add(edtSamp.text + ' - ' + edtDesc_Name.text);
   edtSamp.clear;
end;

procedure TfrmMDmanySamp.edtSampKeyPress(Sender: TObject; var Key: Char);
begin
   if (edtDesc_Name.text <> '') and (length(edtSamp.text) = 6) then
   begin
      listSamp.items.add(edtSamp.text + ' - ' + edtDesc_Name.text);
      edtSamp.Clear;
   end;
end;

end.

