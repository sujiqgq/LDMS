unit LD2I02;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ORSingle, Menus, Buttons, ComCtrls, ToolWin, Db, DBTables, Grids, DBGrids,
  Grids_ts, TSGrid, ExtCtrls, StdCtrls, Mask, DateEdit, MdmsFunc, Gauges,
  ValEdit;

type
  TfrmLD2I02 = class(TfrmSingle)
    pnlCond: TPanel;
    qry_Bunju: TQuery;
    Label1: TLabel;
    mskDate: TDateEdit;
    btnDate: TSpeedButton;
    pnlJisa: TPanel;
    Label3: TLabel;
    cmbJisa: TComboBox;
    Panel5: TPanel;
    edtSt_Bunju: TEdit;
    Label2: TLabel;
    edtED_Bunju: TEdit;
    qryD_bunju: TQuery;
    qryD_gulkwa: TQuery;
    qryU_Gumgin: TQuery;
    Panel2: TPanel;
    btnMibun: TBitBtn;
    btn_Close: TBitBtn;
    Auto_Proc: TGauge;
    btnCheck: TBitBtn;
    Label4: TLabel;
    cntBunju: TValEdit;
    procedure FormCreate(Sender: TObject);
    procedure UP_Click(Sender: TObject);
    procedure btnCheckClick(Sender: TObject);
    procedure edtSt_BunjuChange(Sender: TObject);
    procedure btnMibunClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }

  end;

var
  frmLD2I02: TfrmLD2I02;

implementation

uses ZZFunc, ZZMenu, ZZComm,  MDRefer1;

{$R *.DFM}

procedure TfrmLD2I02.FormCreate(Sender: TObject);
begin
   inherited;
   // 지사권한관리
   GF_JisaSelect(pnlJisa, cmbJisa);
   cmbJisa.Text := Copy(GV_sUserId,1,2);
   // 환경설정
end;

procedure TfrmLD2I02.UP_Click(Sender: TObject);
begin
   inherited;
   if Sender = btnDate then GF_CalendarClick(mskDate);
end;

procedure TfrmLD2I02.btnCheckClick(Sender: TObject);
begin
     inherited;
     If (Trim(mskDate.Text) = '') Or (Length(mskDate.Text) <> 8) Then
        Begin
           Showmessage('☞ 분주일자가 정확하지 않습니다..');
           mskDate.SetFocus;
           Exit;
        End;
     If (Trim(edtSt_Bunju.Text) = '') Or (Length(edtSt_Bunju.Text) <> 4) Then
        Begin
           Showmessage('☞ 시작 분주번호가 정확하지 않습니다..');
           edtSt_Bunju.SetFocus;
           Exit;
        End;
     If (Trim(edtED_Bunju.Text) = '') Or (Length(edtED_Bunju.Text) <> 4) Then
        Begin
           Showmessage('☞ 끝 분주번호가 정확하지 않습니다..');
           edtED_Bunju.SetFocus;
           Exit;
        End;
     With qry_Bunju Do
        Begin
           Active := False;
           ParamByName('Cod_bjjs').AsString := Copy(cmbJisa.Text,1,2);
           ParamByName('Dat_Bunju').AsString := mskDate.Text;
           ParamByName('St_Bunju').AsString  := edtSt_Bunju.Text;
           ParamByName('Ed_Bunju').AsString  := edtED_Bunju.Text;
           Active := True;
           cntBunju.Value := RecordCount;
           Auto_Proc.MaxValue := RecordCount * 3;
           If RecordCount > 0 then
           begin
              btnMibun.Enabled := True;
              GP_query_log(qry_Bunju, 'LD2I02', 'Q', 'Y');
           end;
        End;
     Auto_Proc.Progress := 0;
end;

procedure TfrmLD2I02.edtSt_BunjuChange(Sender: TObject);
begin
     inherited;
     btnMibun.Enabled := False;
end;

procedure TfrmLD2I02.btnMibunClick(Sender: TObject);
Var
   I, sPos : Integer;
begin
     inherited;
     sPos := 0;
  dmcomm.database.starttransaction;
  try
     With qry_Bunju Do
     For I := 1 To RecordCount do
        Begin
           With qryD_Gulkwa Do
              Begin
                 Close;
                 ParamByName('Cod_Jisa').AsString  := qry_Bunju.FieldByName('Cod_Jisa').AsString;
                 ParamByName('Num_id').AsString    := qry_Bunju.FieldByName('Num_Id').AsString;
                 ParamByName('Dat_Bunju').AsString := qry_Bunju.FieldByName('Dat_Bunju').AsString;
                 ParamByName('Num_Bunju').AsString := qry_Bunju.FieldByName('Num_Bunju').AsString;
                 ExecSql;
                 GP_query_log(qryD_Gulkwa, 'LD2I02', 'Q', 'Y');
                 sPos := sPos + 1;
                 Auto_Proc.Progress := sPos;
              End;
           With qryU_Gumgin Do
              Begin
                 Close;
                 ParamByName('Ysno_Bunju').AsString  := 'N';
                 ParamByName('Cod_Jisa').AsString    := qry_Bunju.FieldByName('Cod_Jisa').AsString;
                 ParamByName('Num_id').AsString      := qry_Bunju.FieldByName('Num_Id').AsString;
                 ParamByName('Dat_Gmgn').AsString    := qry_Bunju.FieldByName('Dat_gmgn').AsString;
                 ParamByName('Cod_Update').AsString  := GV_sUserId;
                 ParamByName('Dat_Update').AsString  := GV_sToDay;
                 ExecSql;
                 GP_query_log(qryU_Gumgin, 'LD2I02', 'Q', 'Y');
                 sPos := sPos + 1;
                 Auto_Proc.Progress := sPos;
              End;
           With qryD_Bunju Do
              Begin
                 Close;
                 ParamByName('Cod_Jisa').AsString  := qry_Bunju.FieldByName('Cod_Jisa').AsString;
                 ParamByName('Num_id').AsString    := qry_Bunju.FieldByName('Num_Id').AsString;
                 ParamByName('Dat_Bunju').AsString := qry_Bunju.FieldByName('Dat_Bunju').AsString;
                 ParamByName('Num_Bunju').AsString := qry_Bunju.FieldByName('Num_Bunju').AsString;
                 ExecSql;
                 GP_query_log(qryD_Bunju, 'LD2I02', 'Q', 'Y');
                 sPos := sPos + 1;
                 Auto_Proc.Progress := sPos;
              End;
           Next;
        End;
   except
      on E : EDBEngineError do
      begin
         GF_MsgBox('Error', IntToStr(E.Errors[1].NativeError));
         dmComm.database.Rollback;
         exit;
      end;
   end;
   // database commit
   dmComm.database.Commit;

     qry_Bunju.Active := False;
     Showmessage(' ☞ 미번작업이 완료되었습니다.');
     btnMibun.Enabled := False;

end;

end.
