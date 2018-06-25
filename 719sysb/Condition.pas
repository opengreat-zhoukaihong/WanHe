unit Condition;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmCondition = class(TForm)
    GroupBox1: TGroupBox;
    rbFilter: TRadioButton;
    ckFilterTCH: TCheckBox;
    ckFilterSDCCH: TCheckBox;
    cbFilterTCH: TComboBox;
    edFilterTchQty: TEdit;
    lbTch: TLabel;
    cbFilterSDCCH: TComboBox;
    edFilterSdcchQty: TEdit;
    lbCch: TLabel;
    rbOrder: TRadioButton;
    ckOrderTch: TCheckBox;
    ckOrderSdcch: TCheckBox;
    cbOrderTch: TComboBox;
    edOrderTchQty: TEdit;
    cbOrderSdcch: TComboBox;
    edOrderSdcchQty: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label3: TLabel;
    Label4: TLabel;
    lbCch1: TLabel;
    lbCch2: TLabel;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmCondition: TfmCondition;

implementation

{$R *.DFM}
uses BscMain;

procedure TfmCondition.BitBtn1Click(Sender: TObject);
begin
  if rbFilter.Checked then
  begin
    fmBscMain.wFilterOrder := 'FILTER';
    if ckFilterTCH.Checked then
    begin
      fmBscMain.wTchCheck := 'Y';
      fmBscMain.wTchFlag := cbFilterTCH.Text;
      fmBscMain.wTchQty := edFilterTchQty.Text;
    end
    else
      fmBscMain.wTchCheck := 'N';
    if ckFilterSDCCH.Checked then
    begin
      fmBscMain.wSdcchCheck := 'Y';
      fmBscMain.wSdcchFlag := cbFilterSDCCH.Text;
      fmBscMain.wSdcchQty := edFilterSDCCHQty.Text;
    end
    else
      fmBscMain.wTchCheck := 'N';

  end
  else
  begin
    fmBscMain.wFilterOrder := 'ORDER';
    if ckOrderTCH.Checked then
    begin
      fmBscMain.wTchCheck := 'Y';
      fmBscMain.wTchFlag := cbOrderTCH.Text;
      fmBscMain.wTchQty := edOrderTchQty.Text;
    end
    else
      fmBscMain.wTchCheck := 'N';
    if ckOrderSDCCH.Checked then
    begin
      fmBscMain.wSdcchCheck := 'Y';
      fmBscMain.wSdcchFlag := cbOrderSDCCH.Text;
      fmBscMain.wSdcchQty := edOrderSDCCHQty.Text;
    end
    else
      fmBscMain.wSdcchCheck := 'N';
  end;
end;

procedure TfmCondition.BitBtn2Click(Sender: TObject);
begin
  fmBscMain.wFilterOrder := 'FILTER';
  fmBscMain.wTchCheck := 'Y';
  fmBscMain.wTchFlag := '>=';
  fmBscMain.wTchQty := '-9999';


  fmBscMain.wSdcchCheck := 'Y';
  fmBscMain.wSdcchFlag := '>=';
  fmBscMain.wSdcchQty := '-9999';



end;

end.
