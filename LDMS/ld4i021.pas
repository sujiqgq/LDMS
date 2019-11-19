unit LD4I021;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, QuickRpt, Qrctrls, StdCtrls;

type
  TfrmLD4I021 = class(TForm)
    QuickRep: TQuickRep;
    PageFooterBand1: TQRBand;
    DetailBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape1: TQRShape;
    QRShape4: TQRShape;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRL_UnitNo: TQRLabel;
    QRL_jumin: TQRLabel;
    QRL_Race: TQRLabel;
    QRL_Center: TQRLabel;
    QRL_Datgmgn: TQRLabel;
    QRL_Y: TQRLabel;
    QRLabel15: TQRLabel;
    QRL_N: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape5: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRLabel21: TQRLabel;
    QRShape11: TQRShape;
    QRLabel22: TQRLabel;
    QRShape12: TQRShape;
    QRL_Name: TQRLabel;
    QRL_Age: TQRLabel;
    QRL_Numid: TQRLabel;
    QRL_Doctor: TQRLabel;
    QRL_DCname: TQRLabel;
    QRL_Datget: TQRLabel;
    QRL_Datreq: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRShape13: TQRShape;
    QRLabel13: TQRLabel;
    QRM_sokun: TQRMemo;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRL100: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel49: TQRLabel;
    QRLabel50: TQRLabel;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRL1: TQRLabel;
    QRL5: TQRLabel;
    QRL3: TQRLabel;
    QRL9: TQRLabel;
    QRL7: TQRLabel;
    QRL11: TQRLabel;
    QRL4: TQRLabel;
    QRL6: TQRLabel;
    QRL8: TQRLabel;
    QRL10: TQRLabel;
    QRL2: TQRLabel;
    QRL12: TQRLabel;
    QRL13: TQRLabel;
    QRL17: TQRLabel;
    QRL21: TQRLabel;
    QRL15: TQRLabel;
    QRL14: TQRLabel;
    QRL16: TQRLabel;
    QRL20: TQRLabel;
    QRL19: TQRLabel;
    QRL18: TQRLabel;
    QRL22: TQRLabel;
    QRL_3: TQRLabel;
    QRL_4: TQRLabel;
    QRL_5: TQRLabel;
    Label1: TLabel;
    QRLabel51: TQRLabel;
    QRL_a: TQRLabel;
    QRL_b: TQRLabel;
    QRL_C: TQRLabel;
    QRL_d: TQRLabel;
    QRLabel45: TQRLabel;
    procedure QuickRepNeedData(Sender: TObject; var MoreData: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLD4I021: TfrmLD4I021;

implementation


uses ZZFunc, ZZMenu, ZZComm, MdmsFunc,
     LD1I07F, LD1I071, LD4I02;


{$R *.DFM}

procedure TfrmLD4I021.QuickRepNeedData(Sender: TObject;
  var MoreData: Boolean);
begin
   QRL_UnitNo.Caption      := frmLD4I02.desc_buwi.Text;
   QRL_Name.Caption        := frmLD4I02.desc_name.Text;
   QRL_Age.Caption         := frmLD4I02.age_sex.Text;
   QRL_jumin.Caption       := frmLD4I02.num_jumin.Text;
   QRL_Numid.Caption       := frmLD4I02.Edt_Num_id.Text;
   QRL_Race.Caption        := frmLD4I02.dat_human.Text;
   if frmLD4I02.Virus_Y.Checked  = True then QRL_Y.caption := '¡Û'
   else if frmLD4I02.Virus_N.Checked  = True then QRL_N.caption := '¡Û';
   QRL_Doctor.Caption      := frmLD4I02.cmbCOD_DOCT.Text;
   QRL_Center.Caption      := frmLD4I02.desc_jisa.Text;
   QRL_DCname.Caption      := frmLD4I02.desc_dc.Text;
   QRL_Datgmgn.Caption     := frmLD4I02.dat_gmgn.Text;
   QRL_Datget.Caption      := frmLD4I02.dat_take.Text;
   QRL_Datreq.Caption      := frmLD4I02.Dat_Ask.Text;
   if frmLD4I02.gum_type1.Checked = True then QRL_a.caption := 'O'
   else if frmLD4I02.gum_type2.Checked = True then QRL_b.caption := 'O'
   else if frmLD4I02.gum_type3.Checked = True then QRL_c.caption := 'O'
   else if frmLD4I02.gum_type4.Checked = True then QRL_d.caption := 'O';
   QRM_sokun.Lines.add(frmLD4I02.desc_sokun.Lines.Text);
   if frmLD4I02.GE_Conventional.Checked = True then QRL1.caption  := '¡î';
   if frmLD4I02.GE_Liquid.Checked = True then QRL2.caption  := '¡î';
   if frmLD4I02.GE_Lmp.Checked = True then QRL3.caption  := '¡î';
   QRL_3.Caption := frmLD4I02.GE_Lmp_text.Text;
   if frmLD4I02.GE_Menopause.Checked = True then QRL4.caption  := '¡î';
   QRL_4.Caption := frmLD4I02.GE_Menopause_text.Text;
   if frmLD4I02.GE_Pregnancy.Checked = True then QRL5.caption  := '¡î';
   QRL_5.Caption := frmLD4I02.GE_Pregnancy_text.Text;
   if frmLD4I02.GE_IUD.Checked = True then QRL6.caption  := '¡î';
   if frmLD4I02.GE_Erosion.Checked = True then QRL7.caption  := '¡î';
   if frmLD4I02.GE_Hormone.Checked = True then QRL8.caption  := '¡î';
   if frmLD4I02.GE_Radiotherapy.Checked = True then QRL9.caption  := '¡î';
   if frmLD4I02.NGE_Sputum.Checked = True then QRL10.caption := '¡î';
   if frmLD4I02.NGE_Pleural.Checked = True then QRL11.caption := '¡î';
   if frmLD4I02.NGE_Ascitic.Checked = True then QRL12.caption := '¡î';
   if frmLD4I02.NGE_Joint.Checked = True then QRL13.caption := '¡î';
   if frmLD4I02.NGE_Urine.Checked = True then QRL14.caption := '¡î';
   if frmLD4I02.NGE_CSF.Checked = True then QRL15.caption := '¡î';
   if frmLD4I02.NGE_Other.Checked = True then QRL16.caption := '¡î';
   if frmLD4I02.NGE_Cell_block.Checked = True then QRL17.caption := '¡î';
   if frmLD4I02.FNA_Thyroid.Checked = True then QRL18.caption := '¡î';
   if frmLD4I02.FNA_Lymph.Checked = True then QRL19.caption := '¡î';
   if frmLD4I02.FNA_Breast.Checked = True then QRL20.caption := '¡î';
   if frmLD4I02.FNA_Other.Checked = True then QRL21.caption := '¡î';
   if frmLD4I02.FNA_Cell_block.Checked = True then QRL22.caption := '¡î';
end;

procedure TfrmLD4I021.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin

   QRL_UnitNo.Caption      := '';
   QRL_Name.Caption        := '';
   QRL_Age.Caption         := '';
   QRL_jumin.Caption       := '';
   QRL_Numid.Caption       := '';
   QRL_Race.Caption        := '';
   QRL_Y.Caption           := '';
   QRL_N.Caption           := '';
   QRL_Doctor.Caption      := '';
   QRL_Center.Caption      := '';
   QRL_DCname.Caption      := '';
   QRL_Datgmgn.Caption     := '';
   QRL_Datget.Caption      := '';
   QRL_Datreq.Caption      := '';
   QRL_a.Caption           :='';
   QRL_b.Caption           :='';
   QRL_c.Caption           :='';
   QRL_d.Caption           :='';
   QRM_sokun.Lines.Clear;
   QRL1.Caption            :='';
   QRL2.Caption            :='';
   QRL3.Caption            :='';
   QRL4.Caption            :='';
   QRL5.Caption            :='';
   QRL6.Caption            :='';
   QRL7.Caption            :='';
   QRL8.Caption            :='';
   QRL9.Caption            :='';
   QRL10.Caption           :='';
   QRL11.Caption           :='';
   QRL12.Caption           :='';
   QRL13.Caption           :='';
   QRL14.Caption           :='';
   QRL15.Caption           :='';
   QRL16.Caption           :='';
   QRL17.Caption           :='';
   QRL18.Caption           :='';
   QRL19.Caption           :='';
   QRL20.Caption           :='';
   QRL21.Caption           :='';
   QRL22.Caption           :='';
   QRL_3.Caption           :=''; 
   QRL_4.Caption           :=''; 
   QRL_5.Caption           :=''; 
end;

end.


