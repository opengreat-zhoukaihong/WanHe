program P1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {fmCellObj},
  Unit2 in 'Unit2.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfmCellObj, fmCellObj);
  Application.Run;
end.
