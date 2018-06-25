unit Find;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmFind = class(TForm)
    cbBSName: TComboBox;
    Label1: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmFind: TfmFind;

implementation

{$R *.DFM}
uses BscMain;

procedure TfmFind.FormCreate(Sender: TObject);
var
  i, wRow : Integer;
begin
  wRow := oleMapInfo.eval('tableInfo(base,8)');
  oleMapInfo.do('fetch first from base');
  cbBsName.Text := oleMapInfo.eval('base.bs_name');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from base');
    cbBsName.Items.Add(oleMapInfo.eval('base.bs_name'));
  end;

end;

procedure TfmFind.BitBtn1Click(Sender: TObject);
begin
  oleMapInfo.do('select * from  base where bs_name = "' + cbBSName.Text +
                '" into tmp');
  oleMapInfo.do('fetch first from tmp');
  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  if gSelFlag <> 'CELL' then
    oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.2 Units "km"');
  oleMapInfo.do('Close table tmp');
end;

end.
