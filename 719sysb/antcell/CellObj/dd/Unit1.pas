unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, DBTables, Db, ExtCtrls, DBCtrls,ComObj;

type
  TfmCellObj = class(TForm)
    Button1: TButton;
    quCellObj: TQuery;
    Database1: TDatabase;
    taCell: TTable;
    taCellObj: TTable;
    bmCellObj: TBatchMove;
    quCell: TQuery;
    Button3: TButton;
    RadioGroup1: TRadioGroup;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
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
  wExePath : String;
  oleMapInfo : Variant;
  
implementation
    uses unit2;
{$R *.DFM}


procedure TfmCellObj.Button1Click(Sender: TObject);
procedure  inittable;
var
  ii : integer;
  st : string;
begin
  wExePath := 'c:\cellarea';
  with taCellObj do
  begin
    try
      DataBaseName := wExePath;
      TableName := 'CellObj.dbf';
      close;
      if  active = true then
        Close;
      EmptyTable;
    except
    end;
  end;
  with quCell do
  begin
    DataBaseName := wExePath;
    //TableName := 'Cell.dbf';
  end;
  with taCell do
  begin
    DataBaseName := wExePath;
    TableName := 'Cell.dbf';
  end;
  bmCellObj.Execute;

  with quCellObj do
  begin
    if Active then
      Close;
    DataBaseName := wExePath;
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
  try
    inittable;
  except
  end;
  try
    maincalc;
  finally
    close;
  end;

end;


procedure TfmCellObj.FormCreate(Sender: TObject);
begin
  wExePath := 'C:\My Documents\antcell';
  oleMapInfo := CreateOleObject('MapInfo.Application');

end;

procedure TfmCellObj.Button3Click(Sender: TObject);
procedure  inittable;
var
  ii : integer;
  st : string;
begin
  wExePath := 'c:\cellarea';
  with taCellObj do
  begin
    try
      DataBaseName := wExePath;
      TableName := 'CellObj.dbf';
      close;
      if  active = true then
        Close;
      EmptyTable;
    except
    end;
  end;
  with quCell do
  begin
    DataBaseName := wExePath;
    //TableName := 'Cell.dbf';
  end;
  with taCell do
  begin
    DataBaseName := wExePath;
    TableName := 'Cell.dbf';
  end;
  bmCellObj.Execute;

  with quCellObj do
  begin
    if Active then
      Close;
    DataBaseName := wExePath;
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
  try
    inittable;
  except
  end;
  try
    maincalc;
  finally
    close;
  end;

  oleMapInfo.do('Register Table "' + wExePath + '\CellObj.DBF"  TYPE DBF Charset "WindowsSimpChinese" Into "' + wExePath + '\CellObj.TAB" ');
  oleMapInfo.do('Open Table "' + wExePath + '\CellObj.TAB" Interactive');
  oleMapInfo.do('Browse * From CellObj');

  //oleMapInfo.RunMenuCommand(102);
  oleMapInfo.do('Open Table "' + wExePath + '\AllArea.TAB" Interactive');
  oleMapInfo.do('Browse * From CellObj');
  oleMapInfo.do('Create Map For CellObj CoordSys Earth Projection 1, 0');
  oleMapInfo.do('Run Application "' + wExePath + '\CellObj.mbx"');
  oleMapInfo.do('Close table CellObj');
  
end;

procedure TfmCellObj.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  oleMapInfo := Unassigned;
end;

end.

