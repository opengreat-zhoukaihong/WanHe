unit test;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses Tabmap;
{$R *.DFM}

type
  TCreateObj = function(MapInfo: Variant; UserType: String): Boolean; stdcall;

var
  PFunc : TFarProc;
  Module : THandle;


procedure TForm1.Button1Click(Sender: TObject);
begin
//wServer := dmICData.dbICData.Params.Values['SERVER NAME'];
  Module := LoadLibrary('C:\My Documents\antcell\CellObj.dll');
  if Module > 32 then
  begin
    PFunc := GetProcAddress(Module, 'CreateObj');
    if  TCreateObj(PFunc)(oleMapInfo, 'NQI') then
      ShowMessage('dfgdg');

  end
  else
    ShowMessage('not find <ICLibrary.dll>');
  FreeLibrary(Module);
end;

end.
