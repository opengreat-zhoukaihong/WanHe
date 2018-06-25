unit ObjDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, DBTables, Db, ExtCtrls, DBCtrls,ComObj;

type
  TfmCellObj = class(TForm)
    btCancel: TButton;
    quCellObj: TQuery;
    Database1: TDatabase;
    taCell: TTable;
    taCellObj: TTable;
    bmCellObj: TBatchMove;
    quCell: TQuery;
    btOk: TButton;
    rgCellObj: TRadioGroup;
    ckBscObj: TCheckBox;
    ckMscObj: TCheckBox;
    procedure btCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btOkClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmCellObj: TfmCellObj;
  left1,
  right1,
  top1,
  bottom1,
  Max1:Double;
  wExePath, wUserType : String;
  oleMapInfo : Variant;
  wCreated : Boolean;

implementation
    uses ImpData, ColorDlg;
{$R *.DFM}


procedure TfmCellObj.btCancelClick(Sender: TObject);

begin
  wCreated := False;
  Close;
end;


procedure TfmCellObj.FormCreate(Sender: TObject);
var
  wLength, wPos, i : integer;
begin
  wExePath := UpperCase(Application.ExeName);
  //ShowMessage(wExePath);
  wLength := Length(wExePath);
  for i := 1 to wLength do
  begin
    if wExePath[i] = '\' then
      wPos := i;
  end;
  wExePath := Copy(wExePath, 1, wPos - 1);
  //wExePath := 'C:\My Documents\antcell';
  if wUserType = '' then
    oleMapInfo := CreateOleObject('MapInfo.Application');

end;

procedure TfmCellObj.btOkClick(Sender: TObject);
var
  wRowId, i : Integer;
procedure  inittable;
var
  ii : integer;
  st : string;
begin
  //wExePath := 'c:\cellarea';
  with taCellObj do
  begin
    try
      DataBaseName := wExePath + '\map';
      TableName := 'CellObj.dbf';
      if  active = true then
        Close;
      EmptyTable;
    except
    end;
  end;
  with quCell do
  begin
    DataBaseName := wExePath + '\map';
    //TableName := 'Cell.dbf';
   // Open;
  end;
  with taCell do
  begin
    DataBaseName := wExePath + '\map';
    TableName := 'Cell.dbf';
    //Open;
  end;

  bmCellObj.Execute;
  //exit;
  with quCellObj do
  begin
    if Active then
      Close;
    DataBaseName := wExePath + '\map';
    sql.clear;
    sql.Add('select * ');
    sql.Add('from cellObj ');
    sql.Add('order by LON ');



    open;
    First;
    left1 := FieldByName('LON').ASFloat - 0.1;
    Last;
    right1 := FieldByName('LON').ASFloat + 0.3;

    close;

    sql.clear;
    sql.Add('select LAT from CellObj  order by LAT');

    open;
    first;
    Bottom1 := FieldByName('LAT').AsFloat - 0.15;
    last;
    top1 := FieldByName('LAT').AsFloat;
    Close;
  end;
end;
begin
  oleMapInfo.do('set ProgressBars off');
  wCreated := True;
  oleMapInfo.do('Commit table Selection as "' + wExePath + '\map\Tmparea.tab"');
  if wUserType = '' then
  begin
    oleMapInfo.do('Open Table "' + wExePath + '\map\Cell.TAB" Interactive');
    //oleMapInfo.RunMenuCommand(102);
  end;
  //oleMapInfo.do('Export "Cell" Into "' + wExePath +
  //  '\map\cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('Select * from cell where basetype = "1" or basetype = "2" into tmp');
  oleMapInfo.do('Commit table tmp as "' + wExePath + '\map\MCellObj.tab"');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Open Table "' + wExePath + '\map\MCellObj.TAB" Interactive');
  oleMapInfo.do('update MCellObj set obj = CreateCircle(Lon, lat, 0.01)');
  oleMapInfo.do('Commit table MCellObj');
  oleMapInfo.do('Close table MCellObj');
  oleMapInfo.do('Set Style Pen MakePen(1, 4, RGB(100,100,100))');
  oleMapInfo.do('set Style Brush MakeBrush(1, RGB(255,255,255), -1)');


  try
    inittable;
    maincalc;
  except

    ShowMessage('操作失败!');
    wCreated := False;
    Close;
  end;
  if not wCreated then
    Exit;
  oleMapInfo.do('Register Table "' + wExePath + '\map\CellObj.DBF"  TYPE DBF Charset "WindowsSimpChinese" Into "' + wExePath + '\map\CellObj.TAB" ');
  oleMapInfo.do('Open Table "' + wExePath + '\map\CellObj.TAB" Interactive');
  oleMapInfo.do('Browse * From CellObj');

  //oleMapInfo.RunMenuCommand(102);
  //oleMapInfo.do('Open Table "' + wExePath + '\AllArea.TAB" Interactive');
  oleMapInfo.do('Browse * From CellObj');
  oleMapInfo.do('Create Map For CellObj CoordSys Earth Projection 1, 0');
  if rgCellObj.ItemIndex = 0 then
  begin
    if wUserType = '' then
      oleMapInfo.do('open table "' + wExePath + '\map\area.tab" Interactive');
   // ShowMessage(UpperCase(oleMapInfo.eval('SelectionInfo(1)')));
    {if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'AREA' then
      oleMapInfo.do('Commit table Selection as "'
                     + wExePath + '\map\TmpArea.Tab"')
    else}
    oleMapInfo.do('Commit table area as "' + wExePath + '\map\TmpArea.tab"');
    oleMapInfo.do('open table "' + wExePath + '\map\TmpArea.tab" Interactive');

    oleMapInfo.do('select * from TmpArea into tmp ' );
    oleMapInfo.do('objects combine data name = "AllArea"');
    oleMapInfo.do('Commit table TmpArea as "' + wExePath + '\map\Allarea.tab"');
    oleMapInfo.do('Open Table "' + wExePath + '\map\AllArea.TAB" Interactive');
    oleMapInfo.do('close table TmpArea');
    //oleMapInfo.do('close table tmp');
    oleMapInfo.do('Run Application "' + wExePath + '\map\CellObj.mbx"');
    oleMapInfo.do('Close Table AllArea');
  end
  else
    if rgCellObj.ItemIndex = 1 then
      oleMapInfo.do('Run Application "' + wExePath + '\map\CellObjNone.mbx"')
    else
    begin
      {oleMapInfo.do('Commit table selection as "' + wExePath + '\map\TmpArea.tab"');
      oleMapInfo.do('open table "' + wExePath + '\map\TmpArea.tab" Interactive');
      oleMapInfo.do('Set Map Layer Area Display Off');
      oleMapInfo.do('Set Map Layer TmpArea Editable On');

      oleMapInfo.do('select * from TmpArea into tmp ');
      oleMapInfo.do('objects combine ');

      }

      oleMapInfo.do('open table "' + wExePath + '\map\TmpArea.tab" Interactive');
      wRowId := oleMapInfo.eval('tableInfo(TmpArea,8)');
      oleMapInfo.do('fetch rec 1 from TmpArea');
        oleMapInfo.do('TmpObject = TmpArea.obj');
      for i := 2 to wRowId do
      begin
        oleMapInfo.do('fetch rec ' + IntToStr(i) +' from TmpArea');
        oleMapInfo.do('TmpObject = Combine( TmpObject ,TmpArea.obj)');
      end;
      //oleMapInfo.do('Commit table selection as "' + wExePath + '\map\TmpArea.tab"');
      oleMapInfo.do('Open Table "' + wExePath + '\map\AllArea.TAB" Interactive');
      oleMapInfo.do('update allarea set obj = TmpObject where RowId = 1');
      oleMapInfo.do('Commit table AllArea');
      //oleMapInfo.do('close table TmpArea');
      //oleMapInfo.do('close table tmp');
      oleMapInfo.do('Run Application "' + wExePath + '\map\CellObj.mbx"');
      oleMapInfo.do('Close Table AllArea');
    end;
  oleMapInfo.do('Close table CellObj');
  //ShowMessage('建立成功!');
  Close;
end;

procedure TfmCellObj.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if wUserType = '' then
    oleMapInfo := Unassigned;
end;

procedure TfmCellObj.FormDblClick(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(102);
  oleMapInfo.RunMenuCommand(102);
  Application.CreateForm(TfmColorDlg, fmColorDlg);
  try
    fmColorDlg.ShowModal;
  finally
    fmColorDlg.Free;
  end;
end;

end.






