unit LDAI021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Db, DBTables, ComCtrls, Grids_ts, TSGrid, ExtCtrls;

type
  TfrmLDAI021 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    edt_cod_group: TEdit;
    grdMaster: TtsGrid;
    qry_common_cellsokun: TQuery;
    Label1: TLabel;
    edt_select: TEdit;
    procedure FormShow(Sender: TObject);
    procedure UP_VarMemSet(iLength : Integer);
    procedure FormCreate(Sender: TObject);
    procedure grdMasterCellLoaded(Sender: TObject; DataCol,
      DataRow: Integer; var Value: Variant);
    procedure grdMasterClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    // Grid�� ������ ������
    UV_vdata : Variant;
  public
    { Public declarations }
  end;

var
  frmLDAI021: TfrmLDAI021;

implementation

{$R *.DFM}

uses ZZFunc, MdmsFunc, LDAI02;

procedure TfrmLDAI021.UP_VarMemSet(iLength : Integer);
begin
   // Variant Memory Allocation
   if iLength = 0 then
   begin
      UV_vData   := VarArrayCreate([0, 0], varOleStr);
   end
   else
   begin
      VarArrayRedim(UV_vData,     iLength);
   end;
end;

procedure TfrmLDAI021.FormShow(Sender: TObject);
var i, index : Integer;
begin
   // Grid �ʱ�ȭ
   grdMaster.Rows := 0;

   // Query ���� & �迭�� ����
   with qry_common_cellsokun do
   begin
      Active := False;
      ParamByName('cod_group').AsString := edt_cod_group.text;
      Active := True;

      if Recordcount > 0 then
      begin
         // DataSet�� ���� Variant������ �̵�
         index := 0;
         for i := 0 to RecordCount - 1 do
         begin
            UP_VarMemSet(RecordCount-1);
            if edt_cod_group.text = 'CEL4' then
            UV_vdata[index] := qry_common_cellsokun.fieldByName('desc_code').AsString + qry_common_cellsokun.fieldByName('etc_code').AsString
            else UV_vdata[index] := qry_common_cellsokun.fieldByName('desc_code').AsString;
            Next;
            Inc(index);
         end;
      end;
   end;
   // Grid�� �ڷḦ �Ҵ�
   grdMaster.Rows := index;
   grdMaster.Repaint;
end;

procedure TfrmLDAI021.FormCreate(Sender: TObject);
begin
   UP_VarMemSet(0);
end;

procedure TfrmLDAI021.grdMasterCellLoaded(Sender: TObject; DataCol,
  DataRow: Integer; var Value: Variant);
begin
   // �ڷ㸦 ȭ�鿡 ��ȸ
   case DataCol of
       1 : Value := UV_vdata[DataRow-1];
   end;
end;

procedure TfrmLDAI021.grdMasterClick(Sender: TObject);
var iPos : Integer;
begin
   iPos := grdMaster.CurrentDataRow-1;
   frmLDAI02.edt_data.text := UV_vdata[iPos];
   ModalResult := mrCancel;

end;

procedure TfrmLDAI021.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

end.

