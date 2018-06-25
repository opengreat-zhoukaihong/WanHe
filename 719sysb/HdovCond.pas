unit HdovCond;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmHdovCond = class(TForm)
    GroupBox1: TGroupBox;
    rbFilter: TRadioButton;
    ckFilterTCH: TCheckBox;
    cbFilterTCH: TComboBox;
    rbOrder: TRadioButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    lbTch: TLabel;
    Label1: TLabel;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    cbOrderTch: TComboBox;
    edOrderTchQty: TEdit;
    cbOrderSdcch: TComboBox;
    edOrderSdcchQty: TEdit;
    rbFlunkRate: TRadioButton;
    rbHolost: TRadioButton;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmHdovCond: TfmHdovCond;

implementation

uses BscMain;

{$R *.DFM}

procedure TfmHdovCond.BitBtn1Click(Sender: TObject);
begin
  if rbFilter.Checked then
  begin
    fmBscMain.wFilterOrder := 'FILTER';
    if ckFilterTCH.Checked then
    begin
      fmBscMain.wTchCheck := 'Y';
      fmBscMain.wTchFlag := cbFilterTCH.Text;
      //fmBscMain.wTchQty := edFilterTchQty.Text;
    end
    else
      fmBscMain.wTchCheck := 'N';

  end
  else
  begin
    fmBscMain.wFilterOrder := 'ORDER';
    if rbHolost.Checked then
    begin
      fmBscMain.wTchCheck := 'Y';
      fmBscMain.wTchFlag := cbOrderTCH.Text;
      fmBscMain.wTchQty := edOrderTchQty.Text;
    end ;
    if rbFlunkRate.Checked then
    begin
      fmBscMain.wSdcchCheck := 'Y';
      fmBscMain.wSdcchFlag := cbOrderSDCCH.Text;
      fmBscMain.wTchQty := edOrderSDCCHQty.Text;
    end ;
  end;
end;

end.
