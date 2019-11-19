//==============================================================================
// ���α׷� ���� : �����ǰ����� ã��
// �ý���        : ���հ���
// ��������      : 1999.06.08
// ������        : ��öȣ
// ��������      :
// �������      :
//==============================================================================
unit LD1I07F;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORFind, StdCtrls, Buttons, ExtCtrls, ValEdit, Mask, DateEdit;

type
  TfrmLD1I07F = class(TfrmFind)
    rbJumin: TRadioButton;
    rbName: TRadioButton;
    edtName: TEdit;
    mskJumin: TMaskEdit;
    rbJk: TRadioButton;
    edtJk: TEdit;
    rbHm: TRadioButton;
    edtHm: TEdit;
    procedure UP_Click(Sender: TObject);
    procedure rbJuminClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD1I07F: TfrmLD1I07F;

implementation

uses ZZFunc, LD1I07;

{$R *.DFM}

procedure TfrmLD1I07F.UP_Click(Sender: TObject);
var iStart, iCol, iType : Integer;
    sData : String;
begin
   inherited;
   // iType Paremeter
   // 1 : String
   // 2 : Date
   // 3 : Float
   // 4 : Jumin

   if rbJumin.Checked then
   begin
      sData := mskJumin.Text;
      iCol  := 1;
      iType := 4;
   end
   else if rbName.Checked then
   begin
      sData := edtName.Text;
      iCol  := 2;
      iType := 1;
   end
   else if rbJk.Checked then
   begin
      sData := edtJk.Text;
      iCol  := 3;
      iType := 1;
   end
   else if rbHm.Checked then
   begin
      sData := edtHm.Text;
      iCol  := 4;
      iType := 1;
   end;

   iStart := 1;
   if Sender = btnFindNext then iStart := frmLD1I07.grdMaster.CurrentDataRow+1;

   // ã���Լ� ȣ��
   GF_Find(frmLD1I07.grdMaster, sData, iCol, iStart, iType, cmbOption.ItemIndex);
end;

procedure TfrmLD1I07F.rbJuminClick(Sender: TObject);
begin
   inherited;

   if      Sender = rbJumin then mskJumin.SetFocus
   else if Sender = rbName  then edtName.SetFocus
   else if Sender = rbJk    then edtJk.SetFocus
   else if Sender = rbHm    then edtHm.SetFocus;
end;

end.
