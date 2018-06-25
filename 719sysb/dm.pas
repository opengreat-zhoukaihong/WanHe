unit dm;

interface

uses
  SysUtils, Windows, Messages, Classes, Graphics, Controls, Forms,
  Dialogs, DBTables, DB,FileCtrl;
type
  TDBData = class(TDataModule)
    Table1: TTable;
    Database2: TDatabase;
    Database1: TDatabase;
    procedure DBDataCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DBData: TDBData;

implementation

{$R *.DFM}


procedure TDBData.DBDataCreate(Sender: TObject);
var
  aliList:TStringList;
  DataBasePath:string;
const
  Ali_Name  = 'CDD';
begin
  {DataBasePath:=ExtractFilePath(Application.ExeName)+'TEMP';
  if not DirectoryExists(DataBasePath) then
  try
    MkDir(DataBasePath);
  except
  end;
  Database1.DatabaseName:=DataBasePath;
  try
    if not Session.IsAlias(ALI_NAME) then
    begin
      Session.AddStandardAlias(ALI_NAME,DataBasePath,'Paradox');
      Session.SaveConfigFile;
      Database1.Connected:=False;
      Database1.AliasName:=ALIAS_NAME;
      Database1.Connected:=true;
    end;
  except
  end; //try...except...end }

   DataBasePath:=ExtractFilePath(Application.ExeName)+'DATA';
  if not DirectoryExists(DataBasePath) then
  try
    MkDir(DataBasePath);
  except
  end;
  Database2.DatabaseName:=DataBasePath;

  try
    if   Session.IsAlias(Ali_Name) then session.DeleteAlias (ALI_NAME);
    Session.SaveConfigFile;
    begin
      Session.AddStandardAlias(ALI_NAME,DataBasePath,'Paradox');
      Session.SaveConfigFile;
      Database2.Connected:=False;
      Database2.AliasName:=ALI_NAME;
      Database2.Connected:=true;
    end;
  except
  end;


  end;

end.
