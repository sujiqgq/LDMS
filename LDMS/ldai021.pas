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
    // Grid와 연관된 변수들
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
   // Grid 초기화
   grdMaster.Rows := 0;

   // Query 수행 & 배열에 저장
   with qry_common_cellsokun do
   begin
      Active := False;
      ParamByName('cod_group').AsString := edt_cod_group.text;
      Active := True;

      if Recordcount > 0 then
      begin
         // DataSet의 값을 Variant변수로 이동
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
   // Grid에 자료를 할당
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
   // 자룔를 화면에 조회
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

