unit ctr_globe;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, DBGrids, Tabs,CommCtrl,ComCtrls,FileCtrl;

procedure   ConvertCDD;              //小区参数转换
Function    get_WindowsTemp_Path:String;//获取系统路径
Procedure   CreateRLACPTable(PathToTable:String);//888*******************
Procedure   CreateRLBCPTable(PathToTable:String); //*****************

Procedure   CreateRLCAPTable(PathToTable:String); //888*****************
Procedure   CreateRLCFPTable(PathToTable:String);
Procedure   CreateRLCPPTable(PathToTable:String);
Procedure   CreateRLCXPTable(PathToTable:String);
Procedure   CreateRLCRPTable(PathToTable:String);

Procedure   CreateRLDCPTable(PathToTable:String);
Procedure   CreateRLDEPTable(PathTotable:String);
Procedure   CreateRLDTPTable(PathToTable:String);
Procedure   CreateRLDGPTable(PathTotable:String);

Procedure   CreateRLHPPTable(PathToTable:String);  //**************

Procedure   CreateRLIHPTable(PathToTable:String);
Procedure   CreateRLIMPTable(PathToTable:String);//****************
Procedure   CreateRLLAPTable(PathToTable:String);//888****************
Procedure   CreateRLLBPTable(PathToTable:String);//****************
Procedure   CreateRLLCPTable(PathToTable:String);//******************
Procedure   CreateRLLDPTable(PathToTable:String);
Procedure   CreateRLLFPTable(PathToTable:String);//****************
Procedure   CreateRLLHPTable(PathToTable:String);
Procedure   CreateRLLLPTable(PathToTable:String);//888****************
Procedure   CreateRLLOPTable(PathToTable:String);

Procedure   CreateRLLUPTable(PathToTable:String);
Procedure   CreateRLLPPTable(PathToTable:String);
Procedure   CreateRLLSPTable(PathToTable:String);//************

Procedure   CreateRLMFPTable(PathToTable:String);

Procedure   CreateRLNRPTable(PathToTable:String);

Procedure   CreateRLOLPTable(PathToTable:String);
Procedure   CreateRLOMPTable(PathToTable:String);//888**************
Procedure   CreateRLPCPTable(PathToTable:String);
Procedure   CreateRLPPPTable(PathToTable:String);//888**************
Procedure   CreateRLPRPTable(PathToTable:String);//888**************
Procedure   CreateRLSBPTable(PathToTable:String);
Procedure   CreateRLSCPTable(PathToTable:String);//888**************
Procedure   CreateRLSLPTable(PathToTable:String);//888**************
Procedure   CreateRLSMPTable(PathToTable:String);//888**************
Procedure   CreateRLSTPTable(PathToTable:String);//888**************
Procedure   CreateRLSSPTable(PathToTable:String);

Procedure   CreateRLTYPTable(PathToTable:String);//888**************
Procedure   CreateRLVLPTable(PathToTable:String);//888**************
Function    PickString(S:String;num:byte):String;
procedure   FirstAccess(SName,sdir:String);//预处理

function    findtable(kk:string):integer;
var
  BSCName:String;
  TName : String;//实用命令名 etc: RLBCP
  SFile:TStringList;//小区参数转换的时候存贮选择的文件名
  LineNum:Integer;

implementation

uses
	dm,IniFiles,PUnit;


Procedure ConvertCDD;              //小区参数转换
Var OpenDlg:TOpenDialog;
    i:Integer;
    Sp:String;
{Main}
begin
     OpenDlg := TOpenDialog.Create(nil);
     OpenDlg.FileName := '';
     OpenDlg.DefaultExt := '*.*';
     OpenDlg.Options := [ofHideReadOnly,ofAllowMultiSelect];
     OpenDlg.Filter  := '所有文件 (*.*)|*.*|数据文件 (*.LOG)|*.LOG';
     OpenDlg.Execute;

     if OpenDlg.FileName<>'' then
        begin
           Sp := get_WindowsTemp_Path;

           SFile := TStringList.Create;
           For i:= 0 to OpenDlg.Files.Count-1 do
           SFile.Add(OpenDlg.Files.Strings[I]);

           Fcdd_conv:= TFcdd_conv.Create(Application);
           Fcdd_conv.ShowModal;
           Fcdd_conv.Free;

           SFile.Free;
        end;//if OpenDialog1.FileName<>'' then
end;

Function get_WindowsTemp_Path:String;
var pc,pcc:string;
begin
  pc:=ExtractFilePath(Application.ExeName)+'Ctr';
  if not DirectoryExists(pc) then
  try
    MkDir(pc);
  except
  end;
  pcc:=pc+'\Temp';
  if not DirectoryExists(pcc) then
  try
    MkDir(pcc);
  except
  end;
  Result := String(pc);
end;



Function PickString(S:String;num:byte):String;
Var i:byte;snum:byte;
begin
   For i:=1 to num do
    begin
      S := Trimleft(S);
      if Pos(' ',S)<>0 then snum := Pos(' ',S)-1 else
      snum := length(s);

      Result := Copy(S,1,snum);
      S := Copy(S,Pos(' ',S),length(S));
    end;
end;

{**********************************************************************************}
Procedure CreateRLACPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLACP';
    TableType :=ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('ACSTATE', ftString, 6, False);
    FieldDefs.Add('SLEVEL', ftString, 4, False);
    FieldDefs.Add('STIME', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString,14  , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }

Procedure CreateRLBCPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLBCP';
    TableType :=ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('DBPSTATE', ftString, 12, False);
    FieldDefs.Add('SCTYPE', ftString, 8, False);
    FieldDefs.Add('SDCCHREG', ftString, 6, False);
    FieldDefs.Add('SSDESDL', ftString, 10 , False);
    FieldDefs.Add('REGINTDL', ftString, 10, False);
    FieldDefs.Add('SSLENDL', ftString, 10, False);
    FieldDefs.Add('LCOMPDL', ftString, 10, False);
    FieldDefs.Add('QDESDL', ftString, 10, False);
    FieldDefs.Add('QCOMPDL', ftString, 10, False);
    FieldDefs.Add('QLENDL', ftString, 10, False);
    FieldDefs.Add('BSPWRMIN', ftString,10 , False);
    FieldDefs.Add('RE_DATE', ftString,14  , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{**********************************************************************************}
Procedure CreateRLCAPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLCAP';
    TableType :=ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('ALG', ftString, 10, False);
    FieldDefs.Add('RE_DATE', ftString,14  , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
//************************************************************
Procedure CreateRLCFPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLCFP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);

    FieldDefs.Add('CHGR', ftString, 16, False);
    FieldDefs.Add('SCTYPE', ftString, 16, False);

    FieldDefs.Add('SDCCH', ftString, 12, False);
    FieldDefs.Add('SDCCHAC', ftString, 12, False);

    FieldDefs.Add('TN', ftString, 8, False);
    FieldDefs.Add('CCHPOS', ftString, 6, False);
    FieldDefs.Add('CBCH', ftString, 6, False);

    FieldDefs.Add('HSN', ftString, 10 , False);
    FieldDefs.Add('HOP', ftString, 10, False);
    FieldDefs.Add('DCHNO01', ftString, 4, False);
    FieldDefs.Add('DCHNO02', ftString, 4, False);
    FieldDefs.Add('DCHNO03', ftString, 4, False);
    FieldDefs.Add('DCHNO04', ftString, 4, False);
    FieldDefs.Add('DCHNO05', ftString, 4, False);
    FieldDefs.Add('DCHNO06', ftString, 4, False);
    FieldDefs.Add('DCHNO07', ftString, 4, False);
    FieldDefs.Add('DCHNO08', ftString, 4, False);
    FieldDefs.Add('DCHNO09', ftString, 4, False);
    FieldDefs.Add('DCHNO10', ftString, 4, False);
    FieldDefs.Add('DCHNO11', ftString, 4, False);
    FieldDefs.Add('DCHNO12', ftString, 4, False);
    FieldDefs.Add('DCHNO13', ftString, 4, False);
    FieldDefs.Add('DCHNO14', ftString, 4, False);
    FieldDefs.Add('DCHNO15', ftString, 4, False);
    FieldDefs.Add('DCHNO16', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString,14  , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLCPPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLCPP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('TYPE', ftString, 6, False);
    FieldDefs.Add('BSPWRB', ftString, 6, False);
    FieldDefs.Add('BSPWRT', ftString, 6, False);
    FieldDefs.Add('MSTXPWR', ftString, 6 , False);
    FieldDefs.Add('SCTYPE', ftString, 6 , False);
    FieldDefs.Add('RE_DATE', ftstring,14, False);
    FieldDefs.Add('DATA_CHANGE', ftstring, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLCXPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLCXP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('DTXD', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString,14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLCRPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLCRP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('BCCH', ftString, 4, False);
    FieldDefs.Add('CBCH', ftString, 4, False);
    FieldDefs.Add('SDCCH', ftString, 4, False);
    FieldDefs.Add('NOOFTCH', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString,14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLDCPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLDCP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('STATE', ftString, 6, False);
    FieldDefs.Add('EMERGPRL', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString,14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLDEPTable(PathToTable: String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLDEP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CGI', ftString, 30, False);
    FieldDefs.Add('BSIC', ftString, 6, False);
    FieldDefs.Add('BCCHNO', ftString, 6, False);
    FieldDefs.Add('AGBLK', ftString, 6, False);
    FieldDefs.Add('MFRMS', ftString, 6, False);
    FieldDefs.Add('BCCHTYPE', ftString, 12, False);
    FieldDefs.Add('TYPE', ftString, 6, False);
    FieldDefs.Add('FNOFFSET', ftString, 6, False);
    FieldDefs.Add('XRANGE', ftString, 6, False);
    FieldDefs.Add('CSYSTYPE', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString,14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLDGPTable(PathToTable: String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLDGP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CHGR', ftString, 8, False);
    FieldDefs.Add('SCTYPE', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftstring,255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLDTPTable(PathToTable: String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLDTP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('TSC', ftString, 8, False);
    FieldDefs.Add('SCTYPE', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

//******************************************
Procedure CreateRLHPPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLHPP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CHAP', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLIHPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLIHP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SCTYPE', ftString, 6, False);
    FieldDefs.Add('IHO', ftString, 4, False);
    FieldDefs.Add('MAXIHO', ftString, 4, False);
    FieldDefs.Add('TMAXIHO', ftString, 4, False);
    FieldDefs.Add('TIHO', ftString, 4, False);
    FieldDefs.Add('SSOFFSETUL', ftString, 4, False);
    FieldDefs.Add('SSOFFSETDL', ftString, 4, False);
    FieldDefs.Add('QOFFSETUL', ftString, 4, False);
    FieldDefs.Add('QOFFSETDL', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
//**********************************************
Procedure CreateRLIMPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLIMP';
    TableType :=  ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('ICMSTATE', ftString, 10, False);
    FieldDefs.Add('INTAVE', ftString, 10, False);
    FieldDefs.Add('LIMIT1', ftString, 4, False);
    FieldDefs.Add('LIMIT2', ftString, 4, False);
    FieldDefs.Add('LIMIT3', ftString, 4, False);
    FieldDefs.Add('LIMIT4', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLAPTable(PathToTable: String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLAP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('LAI', ftString, 16, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLBPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLBP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SYSTYPE', ftString, 4, False);
    FieldDefs.Add('TAAVELEN', ftString, 4, False);
    FieldDefs.Add('TINIT', ftString, 4, False);
    FieldDefs.Add('TALLOC', ftString, 4, False);
    FieldDefs.Add('TURGEN', ftString, 4, False);
    FieldDefs.Add('EVALTYPE', ftString, 4, False);
    FieldDefs.Add('TINITAW', ftString, 4, False);
    FieldDefs.Add('TALLOCAW', ftString, 4, False);
    FieldDefs.Add('ASSOC', ftString, 4, False);
    FieldDefs.Add('IBHOASS', ftString, 4, False);
    FieldDefs.Add('IBHOSICH', ftString, 4, False);
    FieldDefs.Add('IHOSICH', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
{    with IndexDefs do
    begin
    Clear;
    Add('', 'CELLID', [ixPrimary, ixUnique]);
    end;}
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLCPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLCP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CLSSTATE', ftString, 4, False);
    FieldDefs.Add('CLSLEVEL', ftString, 4, False);
    FieldDefs.Add('CLSACC', ftString, 4, False);
    FieldDefs.Add('HOCLSACC', ftString, 4, False);
    FieldDefs.Add('RHYST', ftString, 4, False);
    FieldDefs.Add('CLSRAMP', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    {with IndexDefs do
    begin
    Clear;
    Add('', 'CELLID', [ixPrimary, ixUnique]);
    end;}
    CreateTable;
    Free;
  end;
end;


{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLDPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLDP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('MAXTA', ftString, 4, False);
    FieldDefs.Add('RLINKUP', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLFPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLFP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SSEVALSD', ftString, 4, False);
    FieldDefs.Add('QEVALSD', ftString, 4, False);
    FieldDefs.Add('SSEVALSI', ftString, 4, False);
    FieldDefs.Add('QEVALSI', ftString, 4, False);
    FieldDefs.Add('SSLENSD', ftString, 4, False);
    FieldDefs.Add('QLENSD', ftString, 4, False);
    FieldDefs.Add('SSLENSI', ftString, 4, False);
    FieldDefs.Add('QLENSI', ftString, 4, False);
    FieldDefs.Add('SSRAMPSD', ftString, 4, False);
    FieldDefs.Add('SSRAMPSI', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLHPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLHP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('TYPE', ftString, 4, False);
    FieldDefs.Add('LEVEL', ftString, 10, False);
    FieldDefs.Add('LEVTHR', ftString, 4, False);
    FieldDefs.Add('LEVHYST', ftString, 4, False);
    FieldDefs.Add('PSSTEMP', ftString, 4, False);
    FieldDefs.Add('PTIMTEMP', ftString, 4, False);
    FieldDefs.Add('FASTMSREG', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString, 14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLLPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLLP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SCLD', ftString, 6, False);
    FieldDefs.Add('SCLDUL', ftString, 4, False);
    FieldDefs.Add('SCLDLL', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLOPTable(PathToTable:String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLOP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('BSPWR', ftString, 4, False);
    FieldDefs.Add('BSRXMIN', ftString, 4, False);
    FieldDefs.Add('BSRXSUFF', ftString, 4, False);
    FieldDefs.Add('MSRXMIN', ftString, 4, False);
    FieldDefs.Add('MSRXSUFF', ftString, 4, False);
    FieldDefs.Add('SCHO', ftString, 4, False);
    FieldDefs.Add('MISSNM', ftString, 4, False);
    FieldDefs.Add('AW', ftString, 4, False);
    FieldDefs.Add('SCTYPE', ftString, 4, False);
    FieldDefs.Add('BSTXPWR', ftString, 4, False);
    FieldDefs.Add('EXTPEN', ftString, 4, False);
    FieldDefs.Add('HYSTSEP', ftString, 4, False);
    FieldDefs.Add('RE_DATE',ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;


{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLPPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLPP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('PTIMHF', ftString, 4, False);
    FieldDefs.Add('PTIMBQ', ftString, 10, False);
    FieldDefs.Add('PTIMTA', ftString, 4, False);
    FieldDefs.Add('PSSHF', ftString, 4, False);
    FieldDefs.Add('PSSBQ', ftString, 4, False);
    FieldDefs.Add('PSSTA', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
//************************* Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLUPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLUP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SCTYPE', ftString, 10, False);
    FieldDefs.Add('QLIMUL', ftString, 4, False);
    FieldDefs.Add('QLIMDL', ftString, 10, False);
    FieldDefs.Add('TALIM', ftString, 4, False);
    FieldDefs.Add('CELLQ', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLLSPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLLSP';
    TableType :=ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('LSSTATE', ftString, 12, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLMFPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLMFP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('LISTTYPE', ftString, 10, False);
    FieldDefs.Add('LISTTYPE2', ftString, 10, False);
    FieldDefs.Add('MH0', ftString, 4, False);
    FieldDefs.Add('MH1', ftString, 4, False);
    FieldDefs.Add('MH2', ftString, 4, False);
    FieldDefs.Add('MH3', ftString, 4, False);
    FieldDefs.Add('MH4', ftString, 4, False);
    FieldDefs.Add('MH5', ftString, 4, False);
    FieldDefs.Add('MH6', ftString, 4, False);
    FieldDefs.Add('MH7', ftString, 4, False);
    FieldDefs.Add('MH8', ftString, 4, False);
    FieldDefs.Add('MH9', ftString, 4, False);
    FieldDefs.Add('MH10', ftString, 4, False);
    FieldDefs.Add('MH11', ftString, 4, False);
    FieldDefs.Add('MH12', ftString, 4, False);
    FieldDefs.Add('MH13', ftString, 4, False);
    FieldDefs.Add('MH14', ftString, 4, False);
    FieldDefs.Add('MH15', ftString, 4, False);
    FieldDefs.Add('MH20', ftString, 4, False);
    FieldDefs.Add('MH21', ftString, 4, False);
    FieldDefs.Add('MH22', ftString, 4, False);
    FieldDefs.Add('MH23', ftString, 4, False);
    FieldDefs.Add('MH24', ftString, 4, False);
    FieldDefs.Add('MH25', ftString, 4, False);
    FieldDefs.Add('MH26', ftString, 4, False);
    FieldDefs.Add('MH27', ftString, 4, False);
    FieldDefs.Add('MH28', ftString, 4, False);
    FieldDefs.Add('MH29', ftString, 4, False);
    FieldDefs.Add('MH30', ftString, 4, False);
    FieldDefs.Add('MH31', ftString, 4, False);
    FieldDefs.Add('MH32', ftString, 4, False);
    FieldDefs.Add('MH33', ftString, 4, False);
    FieldDefs.Add('MH34', ftString, 4, False);
    FieldDefs.Add('MH35', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLNRPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLNRP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CELLR', ftString, 10, False);
    FieldDefs.Add('DIR', ftString, 10, False);
    FieldDefs.Add('CAND', ftString, 4, False);
    FieldDefs.Add('CS', ftString, 4, False);
    FieldDefs.Add('KHYST', ftString, 4, False);
    FieldDefs.Add('KOFFSET', ftString, 4, False);
    FieldDefs.Add('LHYST', ftString, 4, False);
    FieldDefs.Add('LOFFSET', ftString, 4, False);
    FieldDefs.Add('TRHYST', ftString, 4, False);
    FieldDefs.Add('TROFFSET', ftString, 4, False);
    FieldDefs.Add('AWOFFSET', ftString, 4, False);
    FieldDefs.Add('BQOFFSET', ftString, 4, False);

    FieldDefs.Add('HIHYST', ftString, 4, False);
    FieldDefs.Add('LOHYST', ftString, 4, False);
    FieldDefs.Add('OFFSET', ftString, 4, False);
    FieldDefs.Add('RE_DATE',ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLOLPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLOLP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('LOL', ftString, 10, False);
    FieldDefs.Add('LOLHYST', ftString, 10, False);
    FieldDefs.Add('TAOl', ftString, 4, False);
    FieldDefs.Add('TAOLHYST', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
END;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLOMPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLOMP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('BSCMODE', ftString, 12, False);
    FieldDefs.Add('RE_DATE', ftString, 14 , False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;
 { ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLPCPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLPCP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('DMPSTATE', ftString, 10, False);
    FieldDefs.Add('SCTYPE', ftString, 10, False);
    FieldDefs.Add('SSDES', ftString, 10, False);
    FieldDefs.Add('SSLEN', ftString, 4, False);
    FieldDefs.Add('LCOMPUL', ftString, 4, False);
    FieldDefs.Add('INIDES', ftString, 4, False);
    FieldDefs.Add('PMARG', ftString, 4, False);
    FieldDefs.Add('INILEN', ftString, 4, False);
    FieldDefs.Add('QDESUL', ftString, 4, False);
    FieldDefs.Add('QLEN', ftString, 4, False);
    FieldDefs.Add('QCOMPUL', ftString, 4, False);
    FieldDefs.Add('REGINT', ftString, 4, False);
    FieldDefs.Add('DTXFUL', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

 { ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLPPPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLPPP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('PP', ftString, 10, False);
    FieldDefs.Add('PRL', ftString, 4, False);
    FieldDefs.Add('INAC', ftString, 4, False);
    FieldDefs.Add('PROBF', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

  { ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLPRPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLPRP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SCTYPE', ftString, 6, False);
    FieldDefs.Add('CHTYPE', ftString, 8, False);
    FieldDefs.Add('CHRATE', ftString, 4, False);
    FieldDefs.Add('PP', ftString, 12, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSBPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSBP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CB', ftString, 10, False);
    FieldDefs.Add('MAXRET', ftString, 10, False);
    FieldDefs.Add('TX', ftString, 4, False);
    FieldDefs.Add('ATT', ftString, 4, False);
    FieldDefs.Add('T3212', ftString, 4, False);
    FieldDefs.Add('CBQ', ftString, 4, False);
    FieldDefs.Add('CRO', ftString, 4, False);
    FieldDefs.Add('TO', ftString, 4, False);
    FieldDefs.Add('PT', ftString, 4, False);
    FieldDefs.Add('ECSC', ftString, 4, False);
    FieldDefs.Add('ACC', ftString, 8, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

  { ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSCPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSCP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('STATE', ftString, 6, False);
    FieldDefs.Add('STATSINT', ftString, 8, False);
    FieldDefs.Add('TIME', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSLPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSLP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SCTYPE', ftString, 6, False);
    FieldDefs.Add('ACTIVE', ftString, 6, False);
    FieldDefs.Add('CHTYPE', ftString, 6, False);
    FieldDefs.Add('CHRATE', ftString, 4, False);
    FieldDefs.Add('SPV', ftString, 4, False);
    FieldDefs.Add('LVA', ftString, 4, False);
    FieldDefs.Add('ACL', ftString, 4, False);
    FieldDefs.Add('NCH', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSMPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSMP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SIMSG', ftString, 4, False);
    FieldDefs.Add('MSGDIST', ftString, 6, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSSPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSSP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('ACCMIN', ftString, 10, False);
    FieldDefs.Add('CCHPWR', ftString, 10, False);
    FieldDefs.Add('CRH', ftString, 4, False);
    FieldDefs.Add('DTXU', ftString, 4, False);
    FieldDefs.Add('RLINKT', ftString, 4, False);
    FieldDefs.Add('NECI', ftString, 4, False);
    FieldDefs.Add('MBCR', ftString, 4, False);
    FieldDefs.Add('NCCPERM', ftString, 4, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLSTPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLSTP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('STATE', ftString, 8, False);
    FieldDefs.Add('CHGR', ftString, 4, False);
    FieldDefs.Add('CHSTATE', ftString, 8, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString, 255, False);
    CreateTable;
    Free;
  end;
end;
{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLTYPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLTYP';
    TableType := ttParadox;
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('GSYSTYPE', ftString, 12, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

{ ************************** Create Procedure for Table XuTemp ************************** }
Procedure CreateRLVLPTable(PathToTable : String);
begin
  with TTable.Create(Application) do
  begin
    Active := False;
    DatabaseName := PathToTable;
    TableName := 'RLVLP';
    TableType := ttParadox;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('CHTYPE', ftString, 6, False);
    FieldDefs.Add('ACL', ftString, 4, False);
    FieldDefs.Add('PL', ftString, 4, False);
    FieldDefs.Add('STATUS', ftString, 10, False);
    FieldDefs.Add('SUPSTATE', ftString, 8, False);
    FieldDefs.Add('TCH', ftString, 12, False);
    FieldDefs.Add('SDCCH', ftString, 12, False);
    FieldDefs.Add('RE_DATE', ftString, 14, False);
    FieldDefs.Add('DATA_CHANGE', ftString,255, False);
    CreateTable;
    Free;
  end;
end;

//*************************  FirstAccess***********************
Procedure FirstAccess(SName,sdir:String);//预处理
var F1,F2:TextFile;S1,TS:String;FirstRead:Boolean;
    spp:string;
begin
    LineNum := 0;
    spp:=sdir;
    FirstRead := True;
    AssignFile(F1,SName);
    Reset(F1);
    AssignFile(F2,Spp+'\'+'CTRTEMP.TXT');
    ReWrite(F2);
    While NOT EOF(F1) do
     begin
       Readln(F1,S1);

       if ((Pos('WO',S1)<>0) OR (Pos('EX-A',S1)<>0) OR (Pos('EX-B',S1)<>0)) and (Pos('TIME',S1)<>0) then
       begin
        Continue;
       end else
       begin
       Inc(LineNum,1);
       Writeln(F2,S1);
       end;

       if (Trim(S1)<>'') and (FirstRead) and (Trim(S1)<>'<' )then
        begin
        FirstRead := False;
        ts := Trim(S1);
        TName := Copy(ts,Pos('<',ts)+1,Pos(':',ts)-Pos('<',ts)-1);
        if TName='' then
        TName := Copy(ts,Pos('<',ts)+1,Pos(';',ts)-Pos('<',ts)-1);
        end;
     end;
    CloseFile(F1);
    CloseFile(F2);
    TName := trim(UpperCase(TName));
end;
//**********************************************************************8
{function findtable(kk:string):integer;
 Var list:TStringList;i:integer;
     tr:string;
 begin
   tr:=kk+'.DB';

      list := TStringList.Create;
      try
      Session.GetTableNames(DBData.Database2.DatabaseName,'',True,False,list);
      result:=0;
      For i:=0 to List.Count-1 do
      begin
       if tr=List[i] then
       begin
       result:=1;
       break;
       end;
      end;
      List.Clear;
   Finally
     List.Free;
   end;
end; }
function findtable(kk:string):integer;
var dir,found,tr:string;
begin
  tr:=kk+'.DB';
  dir:= get_windowsTemp_path;
  Found := FileSearch(tr,dir);
  if found<>''  then result:=1 else result:=0;
end;
end.
