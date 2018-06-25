library CellObj;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  SysUtils,
  Classes,
  ComCtrls,
  comobj,
  OleCtnrs,
  Dialogs,
  ImpData in 'ImpData.pas',
  forms,
  ObjDlg in 'ObjDlg.pas' {fmCellObj},
  ColorDlg in 'ColorDlg.pas' {fmColorDlg};
  
function CreateObj(MapInfo : Variant; UserType: String) : Boolean; stdcall;
var
  wTableNum, i : Integer;
begin
  wUserType := UserType;
  oleMapInfo := MapInfo;
  wTableNum := oleMapInfo.eval('NumTables()');
  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('TableInfo(' + IntToStr(i)
         + ', 1)'))) = 'CELLOBJ' then
    begin
      oleMapInfo.do('Close table CellObj');
      Break;
    end;
  end;


  fmCellObj := TfmCellObj.Create(Application);
  try
    fmCellObj.ShowModal;
  finally
    fmCellObj.Free;
  end;
  Result := wCreated;
end;

function ColorDlg(MapInfo : Variant; UserType: String) : Boolean; stdcall;
var
  wTableName : String;
  i, wTableNum : Integer;
  wHasCellObj, wHasRxlevGrid, wHasRxqualGrid : Boolean;
begin
  wUserType := UserType;
  oleMapInfo := MapInfo;
  wHasCellObj := False;
  fmColorDlg := TfmColorDlg.Create(Application);
  try
    oleMapInfo.do('Dim RxlevObject as Object');
    wTableNum := oleMapInfo.eval('NumTables()');
    fmColorDlg.cbTestData.Items.Clear;
    for i := 1 to wTableNum do
    begin
      if UpperCase(Trim(oleMapInfo.eval('TableInfo(' + IntToStr(i)
         + ', 1)'))) = 'RXLEV_GRID' then
      begin
        wHasRxlevGrid := True;
        Continue;
      end;
      if UpperCase(Trim(oleMapInfo.eval('TableInfo(' + IntToStr(i)
         + ', 1)'))) = 'RXQUAL_GRID' then
      begin
        wHasRxqualGrid := True;
        Continue;
      end;
      if UpperCase(Trim(oleMapInfo.eval('TableInfo(' + IntToStr(i)
         + ', 1)'))) = 'CELLOBJ' then
        wHasCellObj := True;
      if UpperCase(Trim(oleMapInfo.eval('ColumnInfo(' + IntToStr(i)
         + ', COL1, 1)'))) = 'TIME' then
      begin
        fmColorDlg.cbTestData.Items.Add(Trim(oleMapInfo.eval('tableInfo('+
          IntToStr(i)+', 1)')));
        fmColorDlg.cbTestData.Text := Trim(oleMapInfo.eval('tableInfo('+
          IntToStr(i)+', 1)'));
      end;

    end;
    if wHasRxlevGrid then
      oleMapInfo.do('Close table Rxlev_grid');
    if wHasRxqualGrid then
      oleMapInfo.do('Close table Rxqual_grid');
    if (fmColorDlg.cbTestData.Items.Count >= 1) and wHasCellObj then
      fmColorDlg.ShowModal
    else
      if not wHasCellObj then
        ShowMessage('没有打开仿真层!')
      else
        ShowMessage('没有路测数据');
  finally
    fmColorDlg.Free;
    oleMapInfo.do('unDim RxlevObject');
    
  end;
  Result := wCreated;
end;

exports
  CreateObj,
  ColorDlg;

begin
end.
