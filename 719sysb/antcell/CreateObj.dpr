program CreateObj;

uses
  Forms,
  ObjDlg in 'ObjDlg.pas' {fmCellObj},
  ImpData in 'ImpData.pas',
  ColorDlg in 'ColorDlg.pas' {fmColorDlg};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfmCellObj, fmCellObj);
  Application.Run;
end.
