
unit LD4I031;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Buttons, Db, DBTables, Grids_ts, TSGrid, ORDialog,
  ComCtrls;

type
  TfrmLD4I031 = class(TfrmDialog)
    pnlDetail: TPanel;
    Panel5: TPanel;
    PBSmear: TQuery;
    Panel15: TPanel;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Panel4: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Label17: TLabel;
    Label14: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Panel8: TPanel;
    Label31: TLabel;
    Label35: TLabel;
    Panel16: TPanel;
    cmbGmgn: TComboBox;
    Label1: TLabel;
    edt_num_jumin: TEdit;
    RBC_size: TLabel;
    RBC_stainability: TLabel;
    RBC_anisocytosis: TLabel;
    RBC_poikilocytosis: TLabel;
    RBC_inclusions: TLabel;
    WBC_number: TLabel;
    WBC_toxic_dohlebody: TLabel;
    WBC_toxic_vacuoles: TLabel;
    WBC_toxic_granules: TLabel;
    WBC_segmentation: TLabel;
    WBC_diff_blast: TLabel;
    WBC_diff_eos: TLabel;
    WBC_diff_promyelo: TLabel;
    WBC_diff_baso: TLabel;
    WBC_diff_myelo: TLabel;
    Label30: TLabel;
    Label33: TLabel;
    WBC_diff_meta: TLabel;
    WBC_diff_immaturecell: TLabel;
    WBC_diff_band: TLabel;
    WBC_diff_seg: TLabel;
    WBC_diff_nRBC: TLabel;
    WBC_diff_lym: TLabel;
    WBC_diff_mono: TLabel;
    PLT_number: TLabel;
    PLT_size: TLabel;
    Opinion: TMemo;
    edt_desc_name: TEdit;
    Label2: TLabel;
    Label6: TLabel;
    edt_cod_jisa: TEdit;
    btnCancel: TBitBtn;

    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  //  procedure FormCreate(Sender: TObject);
   
  private
    { Private declarations }
    SUBSOKUN, UV_SBSOKUN1, UV_SBSOKUN2, UV_SBSOKUN3, UV_SBSOKUN4 : String;
  public
    { Public declarations }
    UV_iCol, UV_iRow : Integer;
  end;

var
  frmLD4I031: TfrmLD4I031;

implementation

uses ZZFunc, ZZComm, MdmsFunc,LD4I03;

{$R *.DFM}


procedure TfrmLD4I031.FormCreate(Sender: TObject);
begin
   inherited;

   //Æû»ý¼º

end;


procedure TfrmLD4I031.FormShow(Sender: TObject);
begin
   with PBSmear do
   begin
      Close;
      PBSmear.ParamByName('NUM_jumin').AsString     := edt_num_jumin.Text;
      PBSmear.ParamByName('cod_jisa').AsString     := edt_cod_jisa.Text;
      PBSmear.ParamByName('Dat_gmgn').AsString   := cmbGmgn.Text;
      Open;
      if PBSmear.IsEmpty = False then
      begin
         RBC_size.caption :=FieldByName('RBC_size').AsString;
         RBC_stainability.caption :=FieldByName('RBC_stainability').AsString;
         if FieldByName('RBC_anisocytosis').AsString <> '' then
         begin
              case StrToInt(FieldByName('RBC_anisocytosis').AsString) of
                   0 : RBC_anisocytosis.caption := '-';
                   1 : RBC_anisocytosis.caption := '+';
                   2 : RBC_anisocytosis.caption := '++';
                   3 : RBC_anisocytosis.caption := '+++';
               else
                   RBC_anisocytosis.caption := '';
               end;
         end;
         if FieldByName('RBC_poikilocytosis').AsString <> '' then
         begin
              case StrToInt(FieldByName('RBC_poikilocytosis').AsString) of
                   0 : RBC_poikilocytosis.caption := '-';
                   1 : RBC_poikilocytosis.caption := '+';
                   2 : RBC_poikilocytosis.caption := '++';
                   3 : RBC_poikilocytosis.caption := '+++';
                   else
                   RBC_poikilocytosis.caption := '';
               end;
         end;
         RBC_inclusions.caption :=FieldByName('RBC_inclusions').AsString;
         WBC_number.caption     :=FieldByName('WBC_number').AsString;
         if FieldByName('WBC_toxic_granules').AsString <> '' then
         begin
              case StrToInt(FieldByName('WBC_toxic_granules').AsString) of
                   0 : WBC_toxic_granules.caption := '-';
                   1 : WBC_toxic_granules.caption := '+';
                   2 : WBC_toxic_granules.caption := '++';
                   3 : WBC_toxic_granules.caption := '+++';
                   else
                   WBC_toxic_granules.caption := '';
              end;
         end;
         if FieldByName('WBC_toxic_vacuoles').AsString <> '' then
         begin
              case StrToInt(FieldByName('WBC_toxic_vacuoles').AsString) of
                   0 : WBC_toxic_vacuoles.caption := '-';
                   1 : WBC_toxic_vacuoles.caption := '+';
                   2 : WBC_toxic_vacuoles.caption := '++';
                   3 : WBC_toxic_vacuoles.caption := '+++';
                   else
                   WBC_toxic_vacuoles.caption := '';
               end;
         end;
         if FieldByName('WBC_toxic_dohlebody').AsString <> '' then
         begin
              case StrToInt(FieldByName('WBC_toxic_dohlebody').AsString) of
                   0 : WBC_toxic_dohlebody.caption := '-';
                   1 : WBC_toxic_dohlebody.caption := '+';
                   else
                   WBC_toxic_dohlebody.caption := '';
              end;
         end;
         
         WBC_segmentation.caption        :=FieldByName('WBC_segmentation').AsString;
         WBC_diff_blast.caption          :=FieldByName('WBC_diff_blast').AsString;
         WBC_diff_promyelo.caption       :=FieldByName('WBC_diff_promyelo').AsString;
         WBC_diff_myelo.caption          :=FieldByName('WBC_diff_myelo').AsString;
         WBC_diff_meta.caption           :=FieldByName('WBC_diff_meta').AsString;
         WBC_diff_band.caption           :=FieldByName('WBC_diff_band').AsString;
         WBC_diff_seg.caption            :=FieldByName('WBC_diff_seg').AsString;
         WBC_diff_lym.caption            :=FieldByName('WBC_diff_lym').AsString;
         WBC_diff_mono.caption           :=FieldByName('WBC_diff_mono').AsString;
         WBC_diff_eos.caption            :=FieldByName('WBC_diff_eos').AsString;
         WBC_diff_baso.caption           :=FieldByName('WBC_diff_baso').AsString;
         WBC_diff_immaturecell.caption   :=FieldByName('WBC_diff_immaturecell').AsString;
         WBC_diff_nRBC.caption           :=FieldByName('WBC_diff_nRBC').AsString;
         PLT_number.caption              :=FieldByName('PLT_number').AsString;
         PLT_size.caption                :=FieldByName('PLT_size').AsString;
         Opinion.Text                    :=FieldByName('Opinion').AsString;
      end;
      Close;
   end;

end;
end.
