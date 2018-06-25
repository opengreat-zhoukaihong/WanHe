unit legeng;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, jpeg;

type
  TfmLegeng = class(TForm)
    nbLegeng: TNotebook;
    Image3: TImage;
    Label11: TLabel;
    Label19: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label17: TLabel;
    Label15: TLabel;
    Image1: TImage;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label1: TLabel;
    Image2: TImage;
    Label4: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Image4: TImage;
    Label14: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    Label10: TLabel;
    Image5: TImage;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Image6: TImage;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Image7: TImage;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Image8: TImage;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Image9: TImage;
    Label38: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Image10: TImage;
    Label43: TLabel;
    Label39: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmLegeng: TfmLegeng;

implementation

{$R *.DFM}

procedure TfmLegeng.FormCreate(Sender: TObject);
begin
  MoveWindow(Handle, Screen.Width - Width, Screen.Height - Height, Width, Height, True);
end;

end.
