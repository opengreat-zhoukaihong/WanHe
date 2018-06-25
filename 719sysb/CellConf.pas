unit CellConf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmCellConf = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    edCellLength: TEdit;
    edCellAngle: TEdit;
    Label1: TLabel;
    Label2: TLabel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmCellConf: TfmCellConf;

implementation

{$R *.DFM}

end.
