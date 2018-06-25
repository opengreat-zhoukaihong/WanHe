unit punit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,
  DBTables, ExtCtrls, Grids, DBGrids, ComCtrls, Db;

type
  TFcdd_conv = class(TForm)
    Table1: TTable;
    Timer1: TTimer;
    BatchMove: TBatchMove;
    Table2: TTable;
    Query2: TQuery;
    Anm: TAnimate;
    Table3: TTable;
    Label1: TLabel;
    Query1: TQuery;
    StatusBar1: TStatusBar;
    ProgressBar: TProgressBar;
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    private
    Sp : String;

    Function  PickString2(FirstS,SubS,SecondS:String):String;
    Function  PickString(S:String;num:byte):String;

    Function  AccessRLACPFile:integer;
    Procedure  AccessRLACPRecord;

    Function  AccessRLBCPFile:integer;
    Procedure AccessRLBCPRecord;

    Function  AccessRLCFPFile:integer;
    Procedure AccessRLCFPRecord;
    Function  AccessRLCAPFile:integer;
    Procedure AccessRLCAPRecord;
    Function  AccessRLCPPFile:integer;
    Procedure AccessRLCPPRecord;

    Function  AccessRLCXPFile:integer;
    Procedure AccessRLCXPRecord;

    Function  AccessRLCRPFile:integer;
    Procedure AccessRLCRPRecord;

    Function  AccessRLDEPFile:integer;
    Procedure AccessRLDEPRecord;
    Function  AccessRLDCPFile:integer;
    Procedure AccessRLDCPRecord;
    Function  AccessRLDTPFile:integer;
    Procedure AccessRLDTPRecord;
    Function  AccessRLDGPFile:integer;
    Procedure AccessRLDGPRecord;

    Function  AccessRLHPPFile:integer;
    Procedure AccessRLHPPRecord;
    Function  AccessRLIHPFile:integer;
    Procedure AccessRLIHPRecord;
    Function  AccessRLIMPFile:integer;
    Procedure AccessRLIMPRecord;
    Function  AccessRLLAPFile:integer;
    Procedure AccessRLLAPRecord;
    Function  AccessRLLBPFile:integer;
    Procedure AccessRLLBPRecord;
    Function  AccessRLLCPFile:integer;
    Procedure AccessRLLCPRecord;
    Function  AccessRLLDPFile:integer;
    Procedure AccessRLLDPRecord;
    Function  AccessRLLHPFile:integer;
    Procedure AccessRLLHPRecord;
    Function  AccessRLLLPFile:integer;
    Procedure AccessRLLLPRecord;
    Function  AccessRLLFPFile:integer;
    Procedure AccessRLLFPRecord;
    Function  AccessRLLOPFile:integer;
    Procedure AccessRLLOPRecord;
    Function  AccessRLLPPFile:integer;
    Procedure AccessRLLPPRecord;
    Function  AccessRLLUPFile:integer;
    Procedure AccessRLLUPRecord;
    Function  AccessRLLSPFile:integer;
    Procedure AccessRLLSPRecord;

    Function  AccessRLMFPFile:integer;
    Procedure AccessRLMFPRecord;

    Function  AccessRLNRPFile:integer;
    Procedure AccessRLNRPRecord(TTs:String;version:integer);

    Function  AccessRLOLPFile:integer;
    Procedure AccessRLOLPRecord;

    Function  AccessRLOMPFile:integer;
    Procedure AccessRLOMPRecord;

    Function  AccessRLPCPFile:integer;
    Procedure AccessRLPCPRecord;
    Function  AccessRLPPPFile:integer;
    Procedure AccessRLPPPRecord;
    Function  AccessRLPRPFile:integer;
    Procedure AccessRLPRPRecord;
    Function  AccessRLSBPFile:integer;
    Procedure AccessRLSBPRecord;
    Function  AccessRLSCPFile:integer;
    Procedure AccessRLSCPRecord;
    Function  AccessRLSLPFile:integer;
    Procedure AccessRLSLPRecord;
    Function  AccessRLSMPFile:integer;
    Procedure AccessRLSMPRecord;
    Function  AccessRLSSPFile:integer;
    Procedure AccessRLSSPRecord;
    Function  AccessRLSTPFile:integer;
    Procedure AccessRLSTPRecord;

    Function  AccessRLTYPFile:integer;
    Procedure AccessRLTYPRecord;

    Function  AccessRLVLPFile:integer;
    Procedure AccessRLVLPRecord;
    Procedure Refreshtable;
    procedure ChangeFields ;
    procedure getChangeField;
    function  Checktable:integer;
    procedure nrptblChange;
    procedure smptblChange;
    procedure prptblChange;
    procedure ppptblChange;
    procedure slptblChange;
    procedure vlptblChange;
    procedure stp_dgptblChange;
    procedure OnecellidChange;
    procedure NonecellidChange;
    procedure moverecords;
  { Private declarations }
  public
    { Public declarations }

  end;

var
  Fcdd_conv: TFcdd_conv;
  F:TextFile;
 // progressBar:TprogressBar;
  CELLS:String;//RLNRP使用

implementation

uses  ctr_globe;

{$R *.DFM}


procedure TFcdd_conv.FormShow(Sender: TObject);
begin
Timer1.Enabled := True;
Table1.DatabaseName := Sp;
end;

Function TFcdd_conv.PickString2(FirstS,SubS,SecondS:String):String;
Var n1,n2:integer;
begin
     n1 := Pos(SubS,FirstS);
     n2 := n1+Length(SubS);
     Result := Copy(SecondS,n1,n2-n1+1);
end;

Function TFcdd_conv.PickString(S:String;num:byte):String;
Var i:byte;snum:byte;First:boolean;
begin
   First := False;
   For i:=1 to num do
    begin
      S := Trimleft(S);

      if First then begin Result := '';exit;end;

      if Pos(' ',S)<>0 then snum := Pos(' ',S)-1 else
      begin
      snum := length(s);
      First := True;
      end;

      Result := Copy(S,1,snum);
      S := Copy(S,snum+1,length(S));
    end;
end;

procedure TFcdd_conv.Button2Click(Sender: TObject);
begin
Close;
end;

procedure TFcdd_conv.Timer1Timer(Sender: TObject);
Var
    j,i,k,ok:integer;PNum:Integer;
    ssp,S1:string;
    F1:TextFile;
begin
      Timer1.Enabled := False;
      label1.caption:='';
      StatusBar1.Panels[1].Text := ' 共选择了 '+InttoStr(SFile.Count)+' 个文件';
      StatusBar1.Panels[0].Text := ' 正在查找BSCName信息！';

      bscname:='';
      For i:=0 to SFile.Count-1 do
        begin
          AssignFile(F1,SFile.Strings[i]);
          Reset(F1);
          While NOT EOF(F1) do
          begin
             Readln(F1,S1);
             if ((Pos('WO',S1)<>0) OR (Pos('EX-A',S1)<>0) OR (Pos('EX-B',S1)<>0)) and (Pos('TIME',S1)<>0) then
             begin
              BSCName := Copy(PickString(S1,2),1,Pos('/',PickString(S1,2))-1);
              break;
             end;
          end;
          CloseFile(F1);
          if  BSCName <>'' then break;
        end;

       if bscname='' then
       begin
         MessageDlg('无BSCName信息！请重新下载这些文件。',mtError,[mbOK],0) ;
         close;
         exit;
       end;

      SP:=  get_WindowsTemp_Path ;
      query2.DatabaseName:=sp;
      table2.DatabaseName:=sp;
      table3.DatabaseName:=sp;
      SP:= sp+'\temp';
      Table1.databaseName:=sp;
      query1.DatabaseName:=sp;
      
      progressBar.visible:=true;
      ssp:='';
      For i:=0 to SFile.Count-1 do
       begin
          ok:=0;
          progressBar.position:=0;
          StatusBar1.Panels[1].Text := ' 共选择了 '+InttoStr(SFile.Count)+' 个文件';
          label1.caption:='正在转换第'+InttoStr(i+1)+'个文件:'+SFile.Strings[i];
          FirstAccess(SFile.Strings[i],sp);
          progressBar.MAX:=2*LineNum;
          repaint;

          progressBar.position:=LineNum div 4;

           if TName = 'RLACP' then //------------------
           begin
            CreateRLACPTable(sp);
            Table1.TableName := 'RLACP';
            if AccessRLACPFile=0 then ssp:=ssp+' '+TName  else ok:=1;
            end;

          if TName = 'RLBCP' then //*******************88
           begin
            CreateRLBCPTable(sp);
            Table1.TableName := 'RLBCP';
            if AccessRLBCPFile=0 then ssp:=ssp+' '+TName  else ok:=1;
           end;

         if TName = 'RLCAP' then //------------------
           begin
            CreateRLCAPTable(sp);
            Table1.TableName := 'RLCAP';
            if AccessRLCAPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

          if TName = 'RLCFP' then
           begin
            CreateRLCFPTable(sp);
            Table1.TableName := 'RLCFP';
            if AccessRLCFPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLCPP' then
           begin
            CreateRLCPPTable(sp);
            Table1.TableName := 'RLCPP';
            if AccessRLCPPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLCXP' then
           begin
            CreateRLCXPTable(sp);
            Table1.TableName := 'RLCXP';
            if AccessRLCXPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLCRP' then
           begin
            CreateRLCRPTable(sp);
            Table1.TableName := 'RLCRP';
            if AccessRLCRPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLDCP' then
           begin
            CreateRLDCPTable(sp);
            Table1.TableName := 'RLDCP';
            if AccessRLDCPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLDGP' then
           begin
            CreateRLDGPTable(sp);
            Table1.TableName := 'RLDGP';
            if AccessRLDGPFile=0 then ssp:=ssp+' '+TName else ok:=1
           end;

           if TName = 'RLDTP' then
           begin
            CreateRLDTPTable(sp);
            Table1.TableName := 'RLDTP';
            if AccessRLDTPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;
           if TName = 'RLDEP' then
           begin
            CreateRLDEPTable(sp);
            Table1.TableName := 'RLDEP';
            if AccessRLDEPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLHPP' then    //*******************8
           begin
            CreateRLHPPTable(sp);
            Table1.TableName := 'RLHPP';
            if AccessRLHPPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

           if TName = 'RLIHP' then
           begin
            CreateRLIHPTable(sp);
            Table1.TableName := 'RLIHP';
            if AccessRLIHPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

          if TName = 'RLIMP' then    //*******************8
           begin
            CreateRLIMPTable(sp);
            Table1.TableName := 'RLIMP';
            if AccessRLIMPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

           if TName = 'RLLAP' then //------------------
           begin
            CreateRLLAPTable(sp);
            Table1.TableName := 'RLLAP';
            if AccessRLLAPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

          if TName = 'RLLBP' then    //*******************8
           begin
            CreateRLLBPTable(sp);
            Table1.TableName := 'RLLBP';
            if AccessRLLBPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

          if TName = 'RLLCP' then    //*******************8
           begin
            CreateRLLCPTable(sp);
            Table1.TableName := 'RLLCP';
            if AccessRLLCPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

          if TName = 'RLLDP' then
           begin
            CreateRLLDPTable(sp);
            Table1.TableName := 'RLLDP';
            if AccessRLLDPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLLFP' then    //*******************8
           begin
            CreateRLLFPTable(sp);
            Table1.TableName := 'RLLFP';
            if AccessRLLFPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

          if TName = 'RLLLP' then //------------------
           begin
            CreateRLLLPTable(sp);
            Table1.TableName := 'RLLLP';
            if AccessRLLLPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

          if TName = 'RLLSP' then    //*******************8
           begin
            CreateRLLSPTable(sp);
            Table1.TableName := 'RLLSP';;
            if AccessRLLSPFile=0 then ssp:=ssp+' '+TName else ok:=1;
          end;

           if TName = 'RLLHP' then
           begin
            CreateRLLHPTable(sp);
            Table1.TableName := 'RLLHP';
            if AccessRLLHPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLLOP' then
           begin
            CreateRLLOPTable(sp);
            Table1.TableName := 'RLLOP';
            if AccessRLLOPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLLPP' then
           begin
            CreateRLLPPTable(sp);
            Table1.TableName := 'RLLPP';
            if AccessRLLPPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

          if TName = 'RLLUP' then
           begin
            CreateRLLUPTable(sp);
            Table1.TableName := 'RLLUP';
            if AccessRLLUPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLMFP' then
           begin
            CreateRLMFPTable(sp);
            Table1.TableName := 'RLMFP';
            if AccessRLMFPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLNRP' then
           begin
            CreateRLNRPTable(sp);
            Table1.TableName := 'RLNRP';
            if AccessRLNRPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLPCP' then
           begin
            CreateRLPCPTable(sp);
            Table1.TableName := 'RLPCP';
            if AccessRLPCPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLPRP' then //------------------
           begin
            CreateRLPRPTable(sp);
            Table1.TableName := 'RLPRP';
            if  AccessRLPRPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

            if TName = 'RLPPP' then //------------------
           begin
            CreateRLPPPTable(sp);
            Table1.TableName := 'RLPPP';
            if AccessRLPPPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLOLP' then
           begin
            CreateRLOLPTable(sp);
            Table1.TableName := 'RLOLP';
            if AccessRLOLPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLOMP' then //------------------
           begin
            CreateRLOMPTable(sp);
            Table1.TableName := 'RLOMP';
            if AccessRLOMPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLSBP' then
           begin
            CreateRLSBPTable(sp);
            Table1.TableName := 'RLSBP';
            if AccessRLSBPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLSCP' then //------------------
           begin
            CreateRLSCPTable(sp);
            Table1.TableName := 'RLSCP';
            if AccessRLSCPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLSLP' then //------------------
           begin
            CreateRLSLPTable(sp);
            Table1.TableName := 'RLSLP';
            if AccessRLSLPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

          if TName = 'RLSMP' then //------------------
           begin
            CreateRLSMPTable(sp);
            Table1.TableName := 'RLSMP';
            if AccessRLSMPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLSTP' then //------------------
           begin
            CreateRLSTPTable(sp);
            Table1.TableName := 'RLSTP';
            if AccessRLSTPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

           if TName = 'RLSSP' then
           begin
            CreateRLSSPTable(sp);
            Table1.TableName := 'RLSSP';
            if AccessRLSSPFile=0 then ssp:=ssp+' '+TName else ok:=1;
           end;

           if TName = 'RLTYP' then //------------------
           begin
            CreateRLTYPTable(sp);
            Table1.TableName := 'RLTYP';
            if AccessRLTYPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

            if TName = 'RLVLP' then //------------------
            begin
            CreateRLVLPTable(sp);
            Table1.TableName := 'RLVLP';
            if AccessRLVLPFile=0 then ssp:=ssp+' '+TName else ok:=1;
            end;

            if ok=1 then refreshtable;
            DeleteFile(sp+'\'+TName+'.DB');
  End;//for i:=0 ...
  
  ProgressBar.visible:=false;
  StatusBar1.Panels[0].Text := ' 文件转换完毕！';

  if ssp<>'' then
  begin
   sp:=ssp;  ssp:=''; ok:=length(sp) div 36;
   for i:=0 to ok do
     ssp:=ssp+copy(sp,(i*36)+1,(i+1)*36)+char(13);
   ssp:='下列文件转换出错，请重新下载: '+char(13)+ssp;
   MessageDlg(ssp,mtInformation,[mbOK],0) ;
  end;
  Close;
end;

procedure TFcdd_conv.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
Anm.Active := False;
Table1.Close;
CanClose := True;
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLACPFile:integer;
Var ks,ts:String;
begin
      try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLACP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
            if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL ADAPTIVE LOGICAL CHANNEL',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     ACSTATE',ts)<>0 then AccessRLACPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
      Except
          result:=0;
          CloseFile(F);
          Table1.Close;
      end;
end;


Procedure TFcdd_conv.AccessRLACPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     ACSTATE  SLEVEL  STIME';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)   do
    begin
     progressBar.position:=progressBar.position+1;
      With Table1 do
      begin
        Append;
        FieldbyName('CELLID').AsString   := PickString(RS2,1);
        FieldbyName('BSCNAME').AsString  := BscName;
        FieldbyName('ACSTATE').AsString     := PickString2(RS,'ACSTATE',RS2);
        FieldbyName('SLEVEL').AsString    := PickString2(RS,'SLEVEL',RS2);
        FieldbyName('STIME').AsString    := PickString2(RS,'STIME',RS2);
        Post;
      end;
     Readln(F,RS2);
    end;//With Table1
end;


Function TFcdd_conv.AccessRLBCPFile:integer;
Var ks,ts:String;PNum:integer;
begin
     Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLBCP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
             Readln(F,ts);
            until Pos('DYNAMIC BTS POWER CONTROL CELL DATA',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     DBPSTATE',ts)<>0 then AccessRLBCPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    result:=0;
    Table1.Close;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLBCPRecord;
Var RS,RS2:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BSCName;
    FieldbyName('CELLID').AsString := PickString(RS,1);
    FieldbyName('DBPSTATE').AsString := PickString(RS,2);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('SCTYPE    SDCCHREG',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('SCTYPE').AsString    := PickString2(RS,'SCTYPE',RS2);
    FieldbyName('SDCCHREG').AsString  := PickString2(RS,'SDCCHREG',RS2);
    FieldbyName('SSDESDL').AsString  := PickString2(RS,'SSDESDL',RS2);
    FieldbyName('REGINTDL').AsString := PickString2(RS,'REGINTDL',RS2);
    FieldbyName('SSLENDL').AsString  := PickString2(RS,'SSLENDL',RS2);
    FieldbyName('LCOMPDL').AsString  := PickString2(RS,'LCOMPDL',RS2);
    FieldbyName('QDESDL').AsString   := PickString2(RS,'QDESDL',RS2);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('QCOMPDL  QLENDL',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('QCOMPDL').AsString := PickString2(RS,'QCOMPDL',RS2);
    FieldbyName('QLENDL').AsString := PickString2(RS,'QLENDL',RS2);
    if (Trim(PickString2(RS,'BSPWRMINP',RS2)) = '') and
       (Trim(PickString2(RS,'BSPWRMINN',RS2)) <> '') then
    FieldbyName('BSPWRMIN').AsString := '-'+PickString2(RS,'BSPWRMINN',RS2);

    if (Trim(PickString2(RS,'BSPWRMINP',RS2)) <> '') and
       (Trim(PickString2(RS,'BSPWRMINN',RS2)) = '') then
    FieldbyName('BSPWRMIN').AsString := PickString2(RS,'BSPWRMINP',RS2);

    end;
     progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1

end;

 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLCAPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLCAP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('BSC CIPHERING ALGORITHM DATA',ts)  <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('ALG',ts)<>0 then AccessRLCAPRecord;
             end; //While
            end; //if Not EOF(f);
            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLCAPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        ALG';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0) and (trim(RS2)<>'') do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('ALG').AsString     := PickString(RS2,1);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;

//****************************************************
Function TFcdd_conv.AccessRLCFPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos(TName,UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL CONFIGURATION FREQUENCY DATA',ts)<> 0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              progressBar.position:=progressBar.position+1;
              if Trim(ts)='CELL' then AccessRLCFPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    Table1.Close;
    result:=0;
end;
end;

Procedure TFcdd_conv.AccessRLCFPRecord;
Var RS,RS2,CS:String;
    SDCCH,TN,CBCH,HSN,HOP,DCHNO,SDCCHAC,CCHPOS:STRING[12];TNum:Byte;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    CS := PickString(RS,1);
    FieldbyName('CELLID').AsString := CS;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('CHGR   SCTYPE    SDCCH',RS)<>0 then
    begin
    Readln(F,RS2);
    SDCCH := PickString2(RS,'SDCCH',RS2);
    TN    := PickString2(RS,'TN',RS2);
    CBCH  := PickString2(RS,'CBCH',RS2);
    HSN   := PickString2(RS,'HSN',RS2);
    HOP   := PickString2(RS,'HOP',RS2);
    DCHNO := PickString2(RS,'DCHNO',RS2);
    SDCCHAC:=PickString2(RS,'SDCCHAC',RS2);
    CCHPOS:= PickString2(RS,'CCHPOS',RS2);

    end;
    progressBar.position:=progressBar.position+1;

    Readln(F,RS2);
    Tnum := 2;
    While (Trim(RS2)<>'') and (Pos('END',RS2)=0) do
     begin
     if Trim(PickString2(RS,'SDCCH',RS2))<>'' then
     SDCCH := PickString2(RS,'SDCCH',RS2);
     if Trim(PickString2(RS,'TN',RS2))<>'' then
     TN    := PickString2(RS,'TN',RS2);
     if Trim(PickString2(RS,'CBCH',RS2))<>'' then
     CBCH  := PickString2(RS,'CBCH',RS2);
     if Trim(PickString2(RS,'HSN',RS2))<>'' then
     HSN   := PickString2(RS,'HSN',RS2);
     if Trim(PickString2(RS,'HOP',RS2))<>'' then
     HOP   := PickString2(RS,'HOP',RS2);
     if Trim(PickString2(RS,'DCHNO',RS2))<>'' then
     DCHNO := PickString2(RS,'DCHNO',RS2);
     if Trim(PickString2(RS,'SDCCHAC',RS2))<>'' then
     SDCCHAC := PickString2(RS,'SDCCHAC',RS2);
     if Trim(PickString2(RS,'CCHPOS',RS2))<>'' then
     CCHPOS := PickString2(RS,'CCHPOS',RS2);

     if TNum<=9 then
     FieldbyName('DCHNO0'+InttoStr(TNum)).AsString  := DCHNO else
     FieldbyName('DCHNO'+InttoStr(TNum)).AsString  := DCHNO;
     Readln(F,RS2);
     progressBar.position:=progressBar.position+1;
     end;
     FieldbyName('SDCCH').AsString  := SDCCH;
     FieldbyName('TN').AsString     := TN;
     FieldbyName('CBCH').AsString   := CBCH;
     FieldbyName('HSN').AsString    := HSN;
     FieldbyName('HOP').AsString    := HOP;
     FieldbyName('DCHNO01').AsString  := DCHNO;
     FieldbyName('SDCCHAC').AsString := SDCCHAC;
     FieldbyName('CCHPOS').AsString  :=CCHPOS;
     Post;

    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    Exit;
    end;//With Table1
end;



Function TFcdd_conv.AccessRLCPPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
          result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;
           If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos(TName,UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL CONFIGURATION POWER DATA',ts)<>0;

            While Not Eof(F) do
             begin
              AccessRLCPPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    Table1.Close;
    result:=0;
end;
end;

Procedure TFcdd_conv.AccessRLCPPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '';RS2 := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('CELL    TYPE BSPWRB BSPWRT MSTXPWR SCTYPE',RS)<>0 then
    begin
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    progressBar.position:=progressBar.position+1;
    Append;
    FieldbyName('CELLID').AsString  := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('TYPE').AsString     := PickString2(RS,'TYPE',RS2);
    FieldbyName('BSPWRB').AsString   := PickString2(RS,'BSPWRB',RS2);
    FieldbyName('BSPWRT').AsString   := PickString2(RS,'BSPWRT',RS2);
    FieldbyName('MSTXPWR').AsString    := PickString2(RS,'MSTXPWR',RS2);
    FieldbyName('SCTYPE').AsString    := PickString2(RS,'SCTYPE',RS2);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;
end;

{---------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLCXPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

           If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if (Pos('RLCXP',UpperCase(ts))= 0) and (Pos('EXT',UpperCase(ts))= 0) then begin CloseFile(F);exit;end;

            ts := '';
            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            ks := trim(ts);
            until Pos('CELL CONFIGURATION DTX DOWNLINK DATA',ts)<> 0;

            PNum := 0;
            While Not Eof(F) do
             begin
              AccessRLCXPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLCXPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '';RS2 := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('CELL    DTXD',RS)<>0 then
    begin
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString  := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('DTXD').AsString    := PickString2(RS,'DTXD',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;
end;

//------------------------------------------------------------------------
Function TFcdd_conv.AccessRLCRPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if (Pos('RLCRP',UpperCase(ts))= 0) and (Pos('EXT',UpperCase(ts))= 0) then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL RESOURCES',ts)<>0;

            PNum := 0;
            While Not Eof(F) do
             begin
              AccessRLCRPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLCRPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '';RS2 := '';
    While (NOT EOF(F)) and (Trim(RS)='') do    Readln(F,RS);

    if Pos('CELL      BCCH  CBCH  SDCCH  NOOFTCH',RS)<>0 then
    begin
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString  := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('BCCH').AsString    :=Trim(PickString2(RS,'BCCH',RS2));
    FieldbyName('CBCH').AsString    := Trim(PickString2(RS,'CBCH',RS2));
    FieldbyName('SDCCH').AsString    := Trim(PickString2(RS,'SDCCH',RS2));
    FieldbyName('NOOFTCH').AsString    := Trim(PickString2(RS,'NOOFTCH',RS2));
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
    end;
end;
{----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLDEPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if (Pos('RLDEP',UpperCase(ts))= 0) then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL DESCRIPTION DATA',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     CGI',ts)<>0 then AccessRLDEPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLDEPRecord;
Var RS,RS2:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString  := PickString(RS,1);
    FieldbyName('CGI').AsString     := PickString(RS,2);
    FieldbyName('BSIC').AsString    := PickString(RS,3);
    FieldbyName('BCCHNO').AsString  := PickString(RS,4);
    FieldbyName('AGBLK').AsString   := PickString(RS,5);
    FieldbyName('MFRMS').AsString   := PickString(RS,6);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('TYPE     BCCHTYPE',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('TYPE').AsString    := PickString2(RS,'TYPE',RS2);
    FieldbyName('BCCHTYPE').AsString    := PickString2(RS,'BCCHTYPE',RS2);
    FieldbyName('FNOFFSET').AsString    := PickString2(RS,'FNOFFSET',RS2);
    FieldbyName('XRANGE').AsString      := PickString2(RS,'XRANGE',RS2);
    FieldbyName('CSYSTYPE').AsString    := PickString2(RS,'CSYSTYPE',RS2);
    end;
    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLDGPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;
           
          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLDGP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
            if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL CHANNEL GROUP DATA',ts) <> 0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('CELL    CHGR',ts)<>0 then AccessRLDGPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLDGPRecord;
Var RS,RS2,tt:String;
    TNum:Byte;
begin
    RS := '        CELL    CHGR   SCTYPE';
    Readln(F,RS2);

    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    if  trim(PickString2(RS,'CELL',RS2))<>'' then tt:=PickString(RS2,1);
    Append;
    FieldbyName('CELLID').AsString   := tt;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('CHGR').AsString     := PickString2(RS,'CHGR',RS2);
    FieldbyName('SCTYPE').AsString    := PickString2(RS,'SCTYPE',RS2);
    Post;
    Readln(F,RS2);
    progressBar.position:=progressBar.position+1;
    end;
    end;//With Table1
end;
 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLDTPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLDTP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
            if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL TRAINING SEQUENCE CODE DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     TSC',ts)<>0 then AccessRLDTPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    result:=0;
    closefile(F);
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLDTPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     TSC   SCTYPE';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('TSC').AsString     := PickString2(RS,'TSC',RS2);
    FieldbyName('SCTYPE').AsString    := PickString2(RS,'SCTYPE',RS2);
    Post;
    Readln(F,RS2);
    progressBar.position:=progressBar.position+1;
    end;
    end;//With Table1
end;
 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLDCPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

           If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLDCP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('BSC DIFFERENTIAL CHANNEL ALLOCATION DATA',ts)<>0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('STATE   EMERGPRL',ts)<>0 then AccessRLDCPRecord;
             end; //While
            end; //if Not EOF(f);
            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLDCPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        STATE   EMERGPRL';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('STATE').AsString     := PickString2(RS,'STATE',RS2);
    FieldbyName('EMERGPRL').AsString    := PickString2(RS,'EMERGPRL',RS2);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLHPPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;
           
          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLHPP',ts)= 0 then begin CloseFile(F);exit;end;

            Repeat
            if EOF(F) then break;
            Readln(F,ts);
            until Pos('CONNECTION OF CELL TO CHANNEL ALLOCATION PROFILE DATA',ts)<> 0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('CELL     CHAP',ts)<>0 then AccessRLHPPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    result:=0;
    CloseFile(F);
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLHPPRecord;
Var RS,RS2:String;
begin
    RS := '        CELL     CHAP';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    IF UpperCase(Trim(RS2))='NONE' then exit;
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('CHAP').AsString := PickString2(RS,'CHAP',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


{----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLIHPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLIHP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOCATING INTRACELL HANDOVER DATA',ts)<> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     SCTYPE',ts)<>0 then AccessRLIHPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLIHPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    TS := '        CELL     SCTYPE  IHO  MAXIHO  TMAXIHO  TIHO  SSOFFSETULP  SSOFFSETULN';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);
    FieldbyName('IHO').AsString := PickString2(TS,'IHO',RS);
    FieldbyName('SCTYPE').AsString := PickString2(TS,'SCTYPE',RS);
    FieldbyName('MAXIHO').AsString := PickString2(TS,'MAXIHO',RS);
    FieldbyName('TMAXIHO').AsString := PickString2(TS,'TMAXIHO',RS);
    FieldbyName('TIHO').AsString := PickString2(TS,'TIHO',RS);
    if (Trim(PickString2(TS,'SSOFFSETULP',RS2)) = '') and
       (Trim(PickString2(TS,'SSOFFSETULN',RS2)) <> '') then
    FieldbyName('SSOFFSETUL').AsString := '-'+PickString2(TS,'SSOFFSETULN',RS);
    if (Trim(PickString2(TS,'SSOFFSETULP',RS2)) <> '') and
       (Trim(PickString2(TS,'SSOFFSETULN',RS2)) = '') then
    FieldbyName('SSOFFSETUL').AsString := PickString2(TS,'SSOFFSETULP',RS);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('SSOFFSETDLP',RS)<>0 then
    begin
    Readln(F,RS2);
    if (Trim(PickString2(RS,'SSOFFSETDLP',RS2)) <> '') and
       (Trim(PickString2(RS,'SSOFFSETDLN',RS2)) = '') then
    FieldbyName('SSOFFSETDL').AsString := PickString2(RS,'SSOFFSETDLP',RS2);
    if (Trim(PickString2(RS,'SSOFFSETDLP',RS2)) = '') and
       (Trim(PickString2(RS,'SSOFFSETDLN',RS2)) <> '') then
    FieldbyName('SSOFFSETDL').AsString := '-'+PickString2(RS,'SSOFFSETDLN',RS2);

    if (Trim(PickString2(RS,'QOFFSETULP',RS2)) <> '') and
       (Trim(PickString2(RS,'QOFFSETULN',RS2)) = '') then
    FieldbyName('QOFFSETUL').AsString := PickString2(RS,'QOFFSETULP',RS2);
    if (Trim(PickString2(RS,'QOFFSETULP',RS2)) = '') and
       (Trim(PickString2(RS,'QOFFSETULN',RS2)) <> '') then
    FieldbyName('QOFFSETUL').AsString := '-'+PickString2(RS,'QOFFSETULN',RS2);

    if (Trim(PickString2(RS,'QOFFSETDLP',RS2)) <> '') and
       (Trim(PickString2(RS,'QOFFSETDLN',RS2)) = '') then
    FieldbyName('QOFFSETDL').AsString := PickString2(RS,'QOFFSETDLP',RS2);
    if (Trim(PickString2(RS,'QOFFSETULP',RS2)) = '') and
       (Trim(PickString2(RS,'QOFFSETULN',RS2)) <> '') then
    FieldbyName('QOFFSETDL').AsString := '-'+PickString2(RS,'QOFFSETULN',RS2);
    end;
    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLFPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLFP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOCATING FILTER DATA',ts) <> 0 ;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('CELL    SSEVALSD',ts)<>0 then AccessRLLFPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;

Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLLFPRecord;
Var RS,TS:String;
begin
    progressBar.position:=progressBar.position+1;
    RS := '';
    TS := '        CELL    SSEVALSD QEVALSD SSEVALSI QEVALSI';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);
    FieldbyName('SSEVALSD').AsString := PickString2(TS,'SSEVALSD',RS);
    FieldbyName('QEVALSD').AsString := PickString2(TS,'QEVALSD',RS);
    FieldbyName('SSEVALSI').AsString := PickString2(TS,'SSEVALSI',RS);
    FieldbyName('QEVALSI').AsString := PickString2(TS,'QEVALSI',RS);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('SSLENSD QLENSD',RS)<>0 then
    begin
    TS := '        SSLENSD QLENSD SSLENSI QLENSI SSRAMPSD SSRAMPSI';
    Readln(F,RS);
    FieldbyName('SSLENSD').AsString := PickString2(TS,'SSLENSD',RS);
    FieldbyName('QLENSD').AsString := PickString2(TS,'QLENSD',RS);
    FieldbyName('SSLENSI').AsString := PickString2(TS,'SSLENSI',RS);
    FieldbyName('QLENSI').AsString := PickString2(TS,'QLENSI',RS);
    FieldbyName('SSRAMPSD').AsString := PickString2(TS,'SSRAMPSD',RS);
    FieldbyName('SSRAMPSI').AsString := PickString2(TS,'SSRAMPSI',RS);
    end;

    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLLPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           result:=1;
           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLLP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('SUBCELL LOAD DISTRIBUTION DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     SCLD',ts)<>0 then AccessRLLLPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;


Procedure TFcdd_conv.AccessRLLLPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     SCLD  SCLDUL  SCLDLL';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
      With Table1 do
      begin
        Append;
        FieldbyName('CELLID').AsString   := PickString(RS2,1);
        FieldbyName('BSCNAME').AsString  := BscName;
        FieldbyName('SCLD').AsString     := PickString2(RS,'SCLD',RS2);
        FieldbyName('SCLDUL').AsString    := PickString2(RS,'SCLDUL',RS2);
        FieldbyName('SCLDLL').AsString    := PickString2(RS,'SCLDLL',RS2);
        Post;
      end;
     Readln(F,RS2);
     progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLHPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

           if  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLHP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOCATING HIERARCHICAL DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     TYPE',ts)<>0 then AccessRLLHPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLHPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     TYPE  LEVEL  LEVTHR  LEVHYST  PSSTEMP  PTIMTEMP';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('TYPE').AsString     := PickString2(RS,'TYPE',RS2);
    FieldbyName('LEVEL').AsString    := PickString2(RS,'LEVEL',RS2);
    FieldbyName('LEVTHR').AsString   := PickString2(RS,'LEVTHR',RS2);
    FieldbyName('LEVHYST').AsString  := PickString2(RS,'LEVHYST',RS2);
    FieldbyName('PSSTEMP').AsString  := PickString2(RS,'PSSTEMP',RS2);
    FieldbyName('PTIMTEMP').AsString := PickString2(RS,'PTIMTEMP',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLIMPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';
           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLIMP',ts)= 0 then begin CloseFile(F);exit;end;

            Repeat
              if EOF(F) then break;
              Readln(F,ts);
            until Pos('CELL IDLE CHANNEL MEASUREMENT DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL    ICMSTATE',ts)<>0 then AccessRLIMPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLIMPRecord;
Var RS,RS2:String;
begin
    RS := '         CELL    ICMSTATE  INTAVE  LIMIT1  LIMIT2  LIMIT3  LIMIT4';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('ICMSTATE').AsString := PickString2(RS,'ICMSTATE',RS2);
    FieldbyName('INTAVE').AsString   := PickString2(RS,'INTAVE',RS2);
    FieldbyName('LIMIT1').AsString   := PickString2(RS,'LIMIT1',RS2);
    FieldbyName('LIMIT2').AsString   := PickString2(RS,'LIMIT2',RS2);
    FieldbyName('LIMIT3').AsString   := PickString2(RS,'LIMIT3',RS2);
    FieldbyName('LIMIT4').AsString   := PickString2(RS,'LIMIT4',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLAPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLAP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('LOCATION AREA DATA',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('LAI',ts)<>0 then AccessRLLAPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLAPRecord;
Var RS,RS2,tt:String;
    i:integer;
begin
    Readln(F,RS);
    RS:=trim(RS);

    RS2:='';
    while (Not EOF(F)) and (trim(RS2)<>'CELL') do Readln(F,RS2);

    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0) and (trim(RS2)<>'') do
    begin
      progressBar.position:=progressBar.position+1;
      RS2:=trim(RS2);
      while RS2<>''do
      begin
       i:=pos(' ',RS2);
       if i=0 then i:=length(RS2);
       tt:=copy(RS2,1,i);
       With Table1 do
       begin
        Append;
        FieldbyName('BSCNAME').AsString := BscName;
        FieldbyName('LAI').AsString     := RS;
        FieldbyName('CELLID').AsString  := tt;
        Post;
       end;
       RS2:=trim(copy(RS2,i+1,length(RS2)-i));
      end;
      Readln(F,RS2);
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLBPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLBP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('BSC LOCATING DATA',ts) <>0 ;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('SYSTYPE',ts)<>0 then AccessRLLBPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
            progressBar.position:=lineNuM;
Except
    Table1.Close;
     result:=0;
     CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLBPRecord;
Var RS,RS2:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('SYSTYPE').AsString := PickString(RS,1);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('TAAVELEN TINIT',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('TAAVELEN').AsString    := PickString2(RS,'TAAVELEN',RS2);
    FieldbyName('TINIT').AsString       := PickString2(RS,'TINIT',RS2);
    FieldbyName('TALLOC').AsString      := PickString2(RS,'TALLOC',RS2);
    FieldbyName('TURGEN').AsString      := PickString2(RS,'TURGEN',RS2);
    FieldbyName('EVALTYPE').AsString    := PickString2(RS,'EVALTYPE',RS2);
    FieldbyName('TINITAW').AsString     := PickString2(RS,'TINITAW',RS2);
    FieldbyName('TALLOCAW').AsString    := PickString2(RS,'TALLOCAW',RS2);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('ASSOC IBHOASS',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('ASSOC').AsString := PickString2(RS,'ASSOC',RS2);
    FieldbyName('IBHOASS').AsString := PickString2(RS,'IBHOASS',RS2);
    FieldbyName('IBHOSICH').AsString := PickString2(RS,'IBHOSICH',RS2);
    FieldbyName('IHOSICH').AsString := PickString2(RS,'IHOSICH',RS2);
    end;

    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLCPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLCP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOAD SHARING DATA',ts)<> 0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     CLSSTATE',ts)<>0 then AccessRLLCPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLCPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     CLSSTATE  CLSLEVEL  CLSACC    HOCLSACC  RHYST     CLSRAMP';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('CLSSTATE').AsString := PickString2(RS,'CLSSTATE',RS2);
    FieldbyName('CLSLEVEL').AsString := PickString2(RS,'CLSLEVEL',RS2);
    FieldbyName('CLSACC').AsString   := PickString2(RS,'CLSACC',RS2);
    FieldbyName('HOCLSACC').AsString := PickString2(RS,'HOCLSACC',RS2);
    FieldbyName('RHYST').AsString    := PickString2(RS,'RHYST',RS2);
    FieldbyName('CLSRAMP').AsString  := PickString2(RS,'CLSRAMP',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLDPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLDP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL LOCATING DISCONNECT DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     MAXTA',ts)<>0 then AccessRLLDPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLDPRecord;
Var RS,RS2,Ks:String;
    TNum:Byte;
begin
    Repeat
    Readln(F,RS2);
    ks := trim(RS2);
    until ks<>'';

    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('MAXTA').AsString    := PickString(RS2,2);
    FieldbyName('RLINKUP').AsString := PickString(RS2,3);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLOPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLOP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOCATING DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     BSPWR',ts)<>0 then AccessRLLOPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLOPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    TS := '        CELL     BSPWR  BSRXMIN  BSRXSUFF  MSRXMIN  MSRXSUFF  SCHO  MISSNM  AW';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);
    FieldbyName('BSPWR').AsString := PickString2(TS,'BSPWR',RS);
    FieldbyName('BSRXMIN').AsString := PickString2(TS,'BSRXMIN',RS);
    FieldbyName('BSRXSUFF').AsString := PickString2(TS,'BSRXSUFF',RS);
    FieldbyName('MSRXMIN').AsString := PickString2(TS,'MSRXMIN',RS);
    FieldbyName('MSRXSUFF').AsString := PickString2(TS,'MSRXSUFF',RS);
    FieldbyName('SCHO').AsString := PickString2(TS,'SCHO',RS);
    FieldbyName('MISSNM').AsString := PickString2(TS,'MISSNM',RS);
    FieldbyName('AW').AsString := PickString(RS,9);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('SCTYPE   BSTXPWR',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('BSTXPWR').AsString := PickString2(RS,'BSTXPWR',RS2);
    FieldbyName('EXTPEN').AsString := PickString2(RS,'EXTPEN',RS2);
    end;
    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1
end;
{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLNRPFile:integer;
Var ks,ts:String;PNum,ver:integer;
begin
  Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

          If  Not EOF(F) then
            begin
            Repeat
              Readln(F,ts);
              ks := trim(ts);
            until ks<>'';
            if Pos('RLNRP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('NEIGHBOUR RELATION DATA',ts) <> 0 ;

            PNum := 0;  ver:=6;
            While (Not Eof(F)) and (PNum<2)do
             begin
              Readln(F,ts);
              if UpperCase(Copy(Trim(Ts),1,4))='CELLR' then  PNum := PNum+1;
              if Pos('HIHYST  LOHYST OFFSETP  OFFSETN',tS)<>0 then  ver:=7;
             end; //While
          end; //if Not EOF(f);
  finally
          CloseFile(F);
  end;

  try
          Table1.Open;
          Reset(F);
          If  Not EOF(F) then
            begin
            While Not Eof(F) do
             begin
              Readln(F,ts);
              if UpperCase(Copy(Trim(Ts),1,4))='CELL' then AccessRLNRPRecord(Ts,ver);
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLNRPRecord(TTs:String;version:integer);
Var RS,RS2,TS:String;
    i:integer;
begin
    if UpperCase(Trim(TTS))='CELL' then
      begin
       Readln(F,RS);
       CELLS := Trim(RS);
      end;
    i:=version;
    RS := '';
    TS := '        CELLR   DIR     CAND    CS';

    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('CELLR   DIR     CAND    CS',RS)<>0 then
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := CELLS;
    FieldbyName('CELLR').AsString := PickString(RS,1);
    FieldbyName('DIR').AsString := PickString(RS,2);
    FieldbyName('CAND').AsString := PickString(RS,3);
    FieldbyName('CS').AsString := PickString(RS,4);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('KHYST   KOFFSETP',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('KHYST').AsString := PickString2(RS,'KHYST',RS2);

    if (Trim(PickString2(RS,'KOFFSETP',RS2)) <> '') and
       (Trim(PickString2(RS,'KOFFSETN',RS2)) = '') then
    FieldbyName('KOFFSET').AsString := PickString2(RS,'KOFFSETP',RS2);
    if (Trim(PickString2(RS,'KOFFSETP',RS2)) = '') and
       (Trim(PickString2(RS,'KOFFSETN',RS2)) <> '') then
    FieldbyName('KOFFSET').AsString := '-'+PickString2(RS,'KOFFSETN',RS2);

    FieldbyName('LHYST').AsString := PickString2(RS,'LHYST',RS2);

    if (Trim(PickString2(RS,'LOFFSETP',RS2)) <> '') and
       (Trim(PickString2(RS,'LOFFSETN',RS2)) = '') then
    FieldbyName('LOFFSET').AsString := PickString2(RS,'LOFFSETP',RS2);
    if (Trim(PickString2(RS,'LOFFSETP',RS2)) = '') and
       (Trim(PickString2(RS,'LOFFSETN',RS2)) <> '') then
    FieldbyName('LOFFSET').AsString := '-'+PickString2(RS,'LOFFSETN',RS2);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('TRHYST  TROFFSETP',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('TRHYST').AsString := PickString2(RS,'TRHYST',RS2);

    if (Trim(PickString2(RS,'TROFFSETP',RS2)) <> '') and
       (Trim(PickString2(RS,'TROFFSETN',RS2)) = '') then
    FieldbyName('TROFFSET').AsString := PickString2(RS,'TROFFSETP',RS2);
    if (Trim(PickString2(RS,'TROFFSETP',RS2)) = '') and
       (Trim(PickString2(RS,'TROFFSETN',RS2)) <> '') then
    FieldbyName('TROFFSET').AsString := '-'+PickString2(RS,'TROFFSETN',RS2);

    FieldbyName('BQOFFSET').AsString := PickString2(RS,'BQOFFSET',RS2);
    FieldbyName('AWOFFSET').AsString := PickString2(RS,'AWOFFSET',RS2);
    end;

    if i=7 then
    begin
    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='')   do   Readln(F,RS);
    if Pos('HIHYST  LOHYST OFFSETP  OFFSETN',RS)<>0 then
    begin
    Readln(F,RS2);
    FieldbyName('HIHYST').AsString := PickString2(RS,'HIHYST',RS2);
    FieldbyName('LOHYST').AsString := PickString2(RS,'LOHYST',RS2);
    if (Trim(PickString2(RS,'OFFSETP',RS2)) <> '') and
       (Trim(PickString2(RS,'OFFSETN',RS2)) = '') then
    FieldbyName('OFFSET').AsString := PickString2(RS,'OFFSETP',RS2);
    if (Trim(PickString2(RS,'OFFSETP',RS2)) = '') and
       (Trim(PickString2(RS,'OFFSETN',RS2)) <> '') then
    FieldbyName('OFFSET').AsString := '-'+PickString2(RS,'OFFSETN',RS2);
    end;
    end;

    progressBar.position:=progressBar.position+1;
     Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLPPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLPP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL LOCATING PENALTY DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL    PTIMHF',ts)<>0 then AccessRLLPPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLPPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL    PTIMHF PTIMBQ PTIMTA PSSHF PSSBQ PSSTA';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('PTIMHF').AsString     := PickString2(RS,'PTIMHF',RS2);
    FieldbyName('PTIMBQ').AsString    := PickString2(RS,'PTIMBQ',RS2);
    FieldbyName('PTIMTA').AsString   := PickString2(RS,'PTIMTA',RS2);
    FieldbyName('PSSHF').AsString  := PickString2(RS,'PSSHF',RS2);
    FieldbyName('PSSBQ').AsString  := PickString2(RS,'PSSBQ',RS2);
    FieldbyName('PSSTA').AsString := PickString2(RS,'PSSTA',RS2);
    Post;
    Readln(F,RS2);
    progressBar.position:=progressBar.position+1;    
    end;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLSPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLSP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('BSC LOAD SHARING STATUS',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('LSSTATE',ts)<>0 then AccessRLLSPRecord;
             end; //While
            end; //if Not EOF(f);

            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLLSPRecord;
Var RS2:String;
    TNum:Byte;
begin
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('LSSTATE').AsString   := PickString(RS2,1);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
    progressBar.position:=progressBar.position+1;

end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLLUPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLLUP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL LOCATING URGENCY DATA',ts) <> 0 ;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL     SCTYPE',ts)<>0 then AccessRLLUPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    result:=0;
    closeFile(F);
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLLUPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        CELL     SCTYPE  QLIMUL  QLIMDL  TALIM  CELLQ';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('QLIMUL').AsString     := PickString2(RS,'QLIMUL',RS2);
    FieldbyName('QLIMDL').AsString    := PickString2(RS,'QLIMDL',RS2);
    FieldbyName('TALIM').AsString   := PickString2(RS,'TALIM',RS2);
    FieldbyName('CELLQ').AsString  := PickString2(RS,'CELLQ',RS2);
    Post;
    Readln(F,RS2);
    end;
    progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLOLPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

           If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLOLP',ts)= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL LOCATING OVERLAID SUBCELL DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL    LOL',ts)<>0 then AccessRLOLPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLOLPRecord;
Var RS,RS2:String;
begin
    RS := '        CELL    LOL LOLHYST TAOL TAOLHYST';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    IF UpperCase(Trim(RS2))='NONE' then exit;
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('LOL').AsString := PickString2(RS,'LOL',RS2);
    FieldbyName('TAOL').AsString   := PickString2(RS,'TAOL',RS2);
    FieldbyName('TAOLHYST').AsString   := PickString2(RS,'TAOLHYST',RS2);
    Post;
    Readln(F,RS2);
    end;
     progressBar.position:=progressBar.position+1;
    end;//With Table1
end;


 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLOMPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLOMP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('BSC BAND OPERATION MODE DATA',ts) <> 0 ;
           
            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('MODE',ts)<>0 then AccessRLOMPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
    result:=0;
    closeFile(F);
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLOMPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('BSCMODE').AsString     := PickString(RS2,1);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLPPPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLPPP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

           Repeat
             if EOF(F) then break;
            Readln(F,ts);
           until  Pos('BSC DIFFERENTIAL CHANNEL ALLOCATION',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('PP        PRL',ts)<>0 then AccessRLPPPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;


Procedure TFcdd_conv.AccessRLPPPRecord;
Var RS,RS2,tt,ttt:String;
    TNum:Byte;
begin
    RS := '        PP        PRL   INAC   PROBF';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
      With Table1 do
      begin
        Append;
        tt:= PickString2(RS,'PP',RS2);
        if trim(tt)<>'' then  ttt:=PickString(RS2,1);
        FieldbyName('PP').AsString   := ttt;
        FieldbyName('BSCNAME').AsString  := BscName;
        FieldbyName('PRL').AsString  := PickString2(RS,'PRL',RS2);
        FieldbyName('INAC').AsString    := PickString2(RS,'INAC',RS2);
        FieldbyName('PROBF').AsString    := PickString2(RS,'PROBF',RS2);
        Post;
      end;
     Readln(F,RS2);
     progressBar.position:=progressBar.position+1;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLPRPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLPRP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('BSC DIFFERENTIAL CHANNEL ALLOCATION',ts)<>0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('PP       CELL',ts)<>0 then AccessRLPRPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;


Procedure TFcdd_conv.AccessRLPRPRecord;
Var RS,RS2,tt,ttt,ff,fff:String;
    TNum:Byte;
    i:integer;
begin
    RS := '        PP       CELL     SCTYPE  CHTYPE  CHRATE';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
      progressBar.position:=progressBar.position+1;
      With Table1 do
      begin
        Append;
        i:=0;

        tt:= PickString2(RS,'PP',RS2);
        if trim(tt)<>'' then
        begin
          ttt:=PickString(RS2,1);
          i:=1;
        end;
        ff:=PickString2(RS,'CELL',RS2);
        if trim(ff)<>'' then
        begin
          if i=0 then  fff:=PickString(RS2,1);
          if i=1 then  fff:=PickString(RS2,2);
        end;

        FieldbyName('PP').AsString   := ttt;
        FieldbyName('CELLID').AsString   := fff;
        FieldbyName('BSCNAME').AsString  := BscName;
        FieldbyName('SCTYPE').AsString  := PickString2(RS,'SCTYPE',RS2);
        FieldbyName('CHTYPE').AsString    := PickString2(RS,'CHTYPE',RS2);
        FieldbyName('CHRATE').AsString    := PickString2(RS,'CHRATE',RS2);
        Post;
      end;
     Readln(F,RS2);
    end;//With Table1
end;

//----------------------------------------------------------------
Function TFcdd_conv.AccessRLSBPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLSBP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL SYSTEM INFORMATION BCCH DATA',ts) <> 0 ;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              if UpperCase(Trim(ts))='CELL' then AccessRLSBPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLSBPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('CB   MAXRET',RS)<>0 then
    begin
    TS := '        CB   MAXRET  TX  ATT  T3212  CBQ   CRO  TO  PT  ECSC';
    Readln(F,RS);
    FieldbyName('CB').AsString := PickString(RS,1);
    FieldbyName('MAXRET').AsString := PickString2(TS,'MAXRET',RS);
    FieldbyName('TX').AsString := PickString2(TS,'TX',RS);
    FieldbyName('ATT').AsString := PickString2(TS,'ATT',RS);
    FieldbyName('T3212').AsString := PickString2(TS,'T3212',RS);
    FieldbyName('CBQ').AsString := PickString(RS,6);
    FieldbyName('CRO').AsString := PickString2(TS,'CRO',RS);
    FieldbyName('TO').AsString := PickString2(TS,'TO',RS);
    FieldbyName('PT').AsString := PickString2(TS,'PT',RS);
    FieldbyName('ECSC').AsString := PickString2(TS,'ECSC',RS);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('ACC',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('ACC').AsString := PickString(RS,1);
    end;
    progressBar.position:=progressBar.position+1;

    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLSCPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLSCP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('BSC DIFFERENTIAL CHANNEL ALLOCATION STATISTICS COLLECTION DATA',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('STATE   STATSINT',ts)<>0 then AccessRLSCPRecord;
             end; //While
            end; //if Not EOF(f);
            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLSCPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        STATE   STATSINT    TIME';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('STATE').AsString     := PickString2(RS,'STATE',RS2);
    FieldbyName('STATSINT').AsString     := PickString2(RS,'STATSINT',RS2);
    FieldbyName('TIME').AsString     := PickString2(RS,'TIME',RS2);

    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLSLPFile:integer;
Var ks,ts:String;PNum,ver:integer;
begin
  try
          result:=1;
          AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           ts := '';
          Table1.Open;
          Reset(F);
          If  Not EOF(F) then
            begin
            Repeat
              Readln(F,ts);
              ks := trim(ts);
            until ks<>'';
            if Pos('RLSLP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
             until  Pos('CELL SUPERVISION OF LOGICAL CHANNEL AVAILABILITY DATA',ts) <> 0;

            While Not Eof(F)  do
             begin
              Readln(F,ts);
              if Trim(Ts)='CELL       SCTYPE' then AccessRLSLPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLSLPRecord;
Var RS,TS3,TS,TS1,TS2:String;
    i:integer;
begin

    TS1:='';
    While (NOT EOF(F)) and (TS1='')  do  Readln(F,TS1);

    TS := '        CELL       SCTYPE';

    RS:='';
    While (NOT EOF(F)) and (pos('ACTIVE     CHTYPE ',RS)=0) do  Readln(F,RS);

    TS2:='        ACTIVE     CHTYPE   CHRATE   SPV   LVA   ACL   NCH';
    Readln(F,RS);
    TS3:=PickString2(TS2,'ACTIVE',RS);
    While (NOT EOF(F))  and (Pos('END',RS)=0) AND (trim(RS)<>'')  do
    begin
       With Table1 do
         begin
         Append;
         FieldbyName('BSCNAME').AsString := BscName;
         FieldbyName('CELLID').AsString := PickString(TS1,1);
         FieldbyName('SCTYPE').AsString := PickString2(TS,'SCTYPE',TS1);
         FieldbyName('ACTIVE').AsString := TS3;
         FieldbyName('CHTYPE').AsString := PickString2(TS2,'CHTYPE',RS);
         FieldbyName('CHRATE').AsString := PickString2(TS2,'CHRATE',RS);
         FieldbyName('SPV').AsString := PickString2(TS2,'SPV',RS);
         FieldbyName('LVA').AsString := PickString2(TS2,'LVA',RS);
         FieldbyName('ACL').AsString := PickString2(TS2,'ACL',RS);
         FieldbyName('NCH').AsString := PickString2(TS2,'NCH',RS);
         Post;
         end;
       Readln(F,RS);
       progressBar.position:=progressBar.position+1;
    end;
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLSMPFile:integer;
Var ks,ts:String;PNum,ver:integer;
begin
  try
          result:=1;
          AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           ts := '';
          Table1.Open;
          Reset(F);
          If  Not EOF(F) then
            begin
            Repeat
              Readln(F,ts);
              ks := trim(ts);
            until ks<>'';
            if Pos('RLSMP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL SYSTEM INFORMATION BCCH MESSAGE DISTRIBUTION',UpperCase(ts))<> 0 ;

            PNum := 0;
            While Not Eof(F)  do
             begin
              Readln(F,ts);
              if Trim(Ts)='CELL     SIMSG  MSGDIST' then AccessRLSMPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLSMPRecord;
Var RS,TS,TS1:String;
begin
    Readln(F,RS);
    TS:=PickString(RS,1);
    TS1:='        CELL     SIMSG  MSGDIST';
    While (NOT EOF(F))  and (Pos('END',RS)=0) AND (trim(RS)<>'')  do
    begin
       With Table1 do
         begin
         Append;
         FieldbyName('BSCNAME').AsString := BscName;
         FieldbyName('CELLID').AsString := TS;
         FieldbyName('SIMSG').AsString := PickString2(TS1,'SIMSG',RS);
         FieldbyName('MSGDIST').AsString := PickString2(TS1,'MSGDIST',RS);
         Post;
         end;
       Readln(F,RS);
       progressBar.position:=progressBar.position+1;
    end;
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLSSPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLSSP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL SYSTEM INFORMATION SACCH AND BCCH DATA',ts) <> 0 ;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if UpperCase(Trim(ts))='CELL' then AccessRLSSPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLSSPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('ACCMIN  CCHPWR',RS)<>0 then
    begin
    TS := '        ACCMIN  CCHPWR  CRH  DTXU  RLINKT  NECI  MBCR';
    Readln(F,RS);
    FieldbyName('ACCMIN').AsString := PickString(RS,1);
    FieldbyName('CCHPWR').AsString := PickString2(TS,'CCHPWR',RS);
    FieldbyName('CRH').AsString := PickString2(TS,'CRH',RS);
    FieldbyName('DTXU').AsString := PickString2(TS,'DTXU',RS);
    FieldbyName('RLINKT').AsString := PickString2(TS,'RLINKT',RS);
    FieldbyName('NECI').AsString := PickString2(TS,'NECI',RS);
    FieldbyName('MBCR').AsString := PickString2(TS,'MBCR',RS);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('NCCPERM',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('NCCPERM').AsString := PickString(RS,1);
    end;

    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLSTPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLSTP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL STATUS',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('CELL    STATE',ts)<>0 then AccessRLSTPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLSTPRecord;
Var tt,RS,RS2,RS3:String;
    TNum:Byte;
    i:integer;
begin
    Rs3:='';
    i:=0;
    RS := '        CELL    STATE  CHGR STATE';
    Readln(F,RS2);
    tt:=trim(RS2);
    if length(tt)>length('CELL    STATE  CHGR') then i:=1;
    While (Not EOF(F)) and (Pos('END',RS2)=0)  and (trim(RS2)<>'')do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('CELLID').AsString   := PickString(RS2,1);
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('STATE').AsString     := PickString2(RS,'STATE',RS2);
    if i=1 then
    begin
     FieldbyName('CHGR').AsString     := PickString2(RS,'CHGR',RS2);
     FieldbyName('CHSTATE').AsString     := PickString2(RS,'STATE',RS2);
     post;
     Readln(F,RS3);
     Append;
     FieldbyName('CELLID').AsString   := PickString(RS2,1);
     FieldbyName('BSCNAME').AsString  := BscName;
     FieldbyName('STATE').AsString     := PickString2(RS,'STATE',RS2);
     FieldbyName('CHGR').AsString     := PickString2(RS,'CHGR',RS3);
     FieldbyName('CHSTATE').AsString     := PickString(RS3,2);
     progressBar.position:=progressBar.position+1;
    end;
    Post;
    Readln(F,RS2);
    progressBar.position:=progressBar.position+1;
    end;
    end;//With Table1
end;



{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLMFPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLMFP',ts)= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until Pos('CELL MEASUREMENT FREQUENCIES',ts) <> 0;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if UpperCase(Trim(ts))='CELL' then AccessRLMFPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
     result:=0;
     CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLMFPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    While (NOT EOF(F)) and (Trim(RS)='') do Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('LISTTYPE',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('LISTTYPE').AsString := PickString(RS,1);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('MBCCHNO',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('MH0').AsString := PickString(RS,1);
    FieldbyName('MH1').AsString := PickString(RS,2);
    FieldbyName('MH2').AsString := PickString(RS,3);
    FieldbyName('MH3').AsString := PickString(RS,4);
    FieldbyName('MH4').AsString := PickString(RS,5);
    FieldbyName('MH5').AsString := PickString(RS,6);
    FieldbyName('MH6').AsString := PickString(RS,7);
    FieldbyName('MH7').AsString := PickString(RS,8);
    FieldbyName('MH8').AsString := PickString(RS,9);
    FieldbyName('MH9').AsString := PickString(RS,10);
    FieldbyName('MH10').AsString := PickString(RS,11);
    FieldbyName('MH11').AsString := PickString(RS,12);
    FieldbyName('MH12').AsString := PickString(RS,13);
    FieldbyName('MH13').AsString := PickString(RS,14);
    FieldbyName('MH14').AsString := PickString(RS,15);
    FieldbyName('MH15').AsString := PickString(RS,16);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('LISTTYPE',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('LISTTYPE2').AsString := PickString(RS,1);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('MBCCHNO',RS)<>0 then
    begin
    Readln(F,RS);
    FieldbyName('MH20').AsString := PickString(RS,1);
    FieldbyName('MH21').AsString := PickString(RS,2);
    FieldbyName('MH22').AsString := PickString(RS,3);
    FieldbyName('MH23').AsString := PickString(RS,4);
    FieldbyName('MH24').AsString := PickString(RS,5);
    FieldbyName('MH25').AsString := PickString(RS,6);
    FieldbyName('MH26').AsString := PickString(RS,7);
    FieldbyName('MH27').AsString := PickString(RS,8);
    FieldbyName('MH28').AsString := PickString(RS,9);
    FieldbyName('MH29').AsString := PickString(RS,10);
    FieldbyName('MH30').AsString := PickString(RS,11);
    FieldbyName('MH31').AsString := PickString(RS,12);
    FieldbyName('MH32').AsString := PickString(RS,13);
    FieldbyName('MH33').AsString := PickString(RS,14);
    FieldbyName('MH34').AsString := PickString(RS,15);
    FieldbyName('MH35').AsString := PickString(RS,16);
    end;

    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;
    end;//With Table1
end;

{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLPCPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';
           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLPCP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('DYNAMIC MS POWER CONTROL CELL DATA',ts)<>0;

            PNum := 0;
            While Not Eof(F) do
             begin
              Readln(F,ts);
              Inc(PNum);
              if Pos('CELL     DMPSTATE',ts)<>0 then AccessRLPCPRecord;
              Application.ProcessMessages;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
    Table1.Close;
    result:=0;
    CloseFile(F);
end;
end;

Procedure TFcdd_conv.AccessRLPCPRecord;
Var RS,RS2,TS:String;
begin
    RS := '';
    TS := '        CELL     DMPSTATE';
    While (NOT EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString := BscName;
    FieldbyName('CELLID').AsString := PickString(RS,1);
    FieldbyName('DMPSTATE').AsString := PickString2(TS,'DMPSTATE',RS);

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('SCTYPE   SSDES',RS)<>0 then
    begin
    TS := '        SCTYPE   SSDES   SSLEN  LCOMPUL  INIDES  PMARG  INILEN  QDESUL  QLEN';
    Readln(F,RS);
    FieldbyName('SSDES').AsString := PickString2(TS,'SSDES',RS);
    FieldbyName('SSLEN').AsString := PickString2(TS,'SSLEN',RS);
    FieldbyName('LCOMPUL').AsString := PickString2(TS,'LCOMPUL',RS);
    FieldbyName('INIDES').AsString := PickString2(TS,'INIDES',RS);
    FieldbyName('PMARG').AsString := PickString2(TS,'PMARG',RS);
    FieldbyName('INILEN').AsString := PickString2(TS,'INILEN',RS);
    FieldbyName('QDESUL').AsString := PickString2(TS,'QDESUL',RS);
    FieldbyName('QLEN').AsString := PickString2(TS,'QLEN',RS);
    end;

    RS := '';
    While (NOT System.EOF(F)) and (Trim(RS)='') do
    Readln(F,RS);

    if Pos('QCOMPUL  REGINT',RS)<>0 then
    begin
    TS := '        QCOMPUL  REGINT  DTXFUL';
    Readln(F,RS);
    FieldbyName('QCOMPUL').AsString := PickString2(TS,'QCOMPUL',RS);
    FieldbyName('REGINT').AsString := PickString2(TS,'REGINT',RS);
    FieldbyName('DTXFUL').AsString := PickString2(TS,'DTXFUL',RS);
    end;
    progressBar.position:=progressBar.position+1;
    Readln(F,RS);
    if (Trim(RS)='') OR (Pos('END',RS)<>0) then
    begin
    Post;
    Exit;
    end;

    end;//With Table1
end;

 {-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLTYPFile:integer;
Var ks,ts:String;PNum:integer;
begin
Try
           result:=1;
           AssignFile(F,Sp+'\'+'CTRTEMP.TXT');
           Reset(F);
           ts := '';

           Table1.Open;

          If  Not EOF(F) then
            begin
            Repeat
            Readln(F,ts);
            ks := trim(ts);
            until ks<>'';
            if Pos('RLTYP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
             until  Pos('BSC SYSTEM TYPE DATA',ts) <> 0 ;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Pos('GSYSTYPE',ts)<>0 then AccessRLTYPRecord;
             end; //While
            end; //if Not EOF(f);
            progressBar.position:=lineNuM;
            CloseFile(F);
            Table1.Close;
Except
   CloseFile(F);
    result:=0;
    Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLTYPRecord;
Var RS,RS2:String;
    TNum:Byte;
begin
    RS := '        GSYSTYPE';
    Readln(F,RS2);
    While (Not EOF(F)) and (Pos('END',RS2)=0)  do
    begin
    With Table1 do
    begin
    Append;
    FieldbyName('BSCNAME').AsString  := BscName;
    FieldbyName('GSYSTYPE').AsString     := PickString2(RS,'GSYSTYPE',RS2);
    Post;
    Readln(F,RS2);
    end;
    end;//With Table1
end;


{-----------------------------------------------------------------------------}
Function TFcdd_conv.AccessRLVLPFile:integer;
Var ks,ts:String;PNum,ver:integer;
begin
  try
          result:=1;
          AssignFile(F,Sp+'/'+'CTRTEMP.TXT');
          ts := '';
          Table1.Open;
          Reset(F);
          If  Not EOF(F) then
            begin
            Repeat
              Readln(F,ts);
              ks := trim(ts);
            until ks<>'';
            if Pos('RLVLP',UpperCase(ts))= 0 then begin CloseFile(F);exit;end;

            Repeat
             if EOF(F) then break;
            Readln(F,ts);
            until  Pos('CELL SEIZURE SUPERVISION OF LOGICAL CHANNELS DATA',ts) <> 0 ;

            While Not Eof(F) do
             begin
              Readln(F,ts);
              if Trim(Ts)='CHTYPE  ACL     PL     STATUS    SUPSTATE' then AccessRLVLPRecord;
             end; //While
            end; //if Not EOF(f);
            CloseFile(F);
            Table1.Close;
Except
   CloseFile(F);
   result:=0;
   Table1.Close;
end;
end;

Procedure TFcdd_conv.AccessRLVLPRecord;
Var RS,TS3,TS,TS1,TS2:String;
    i:integer;
begin
    TS2:='';
    While (NOT EOF(F)) and (TS2='')  do  Readln(F,TS2);
    TS3:='';  Readln(F,TS3);

    TS := '        CHTYPE  ACL     PL     STATUS    SUPSTATE';

    Readln(F,RS);

    While (NOT EOF(F)) and (pos('SDCCH',RS)=0) do  Readln(F,RS);

    RS:='';
    While (NOT EOF(F)) and (RS='')  do  Readln(F,RS);

    While (NOT EOF(F))  and (Pos('END',RS)=0)   do
    begin
        TS1:=TS2;
        while TS1<>'' do
        begin
         With Table1 do
         begin
         Append;
         FieldbyName('BSCNAME').AsString := BscName;
         FieldbyName('CELLID').AsString := PickString(RS,1);
         FieldbyName('TCH').AsString := PickString(RS,2);
         FieldbyName('SDCCH').AsString := PickString(RS,3);
         FieldbyName('CHTYPE').AsString := PickString2(TS,'CHTYPE',TS1);
         FieldbyName('ACL').AsString := PickString2(TS,'ACL',TS1);
         FieldbyName('PL').AsString := PickString2(TS,'PL',TS1);
         FieldbyName('STATUS').AsString := PickString2(TS,'STATUS',TS1);
         FieldbyName('SUPSTATE').AsString := PickString2(TS,'SUPSTATE',TS1);
         Post;
         end;
         if TS1=TS3 then break;
         TS1 :=TS3;
        end;
      Readln(F,RS);
      progressBar.position:=progressBar.position+1;
     end;
end;


function TFcdd_conv.Checktable:integer;
VAR kg,sy,hh:string;
    i:integer;
begin

   result:=1;
   kg:=Lowercase(TNAME);
   try
        with query2 do
        begin
        close;
        sql.clear;
        sql.Add('delete  from '+kg+' where BSCNAME=:p ');
        sql.Add('and  RE_DATE=:q ');
        parambyname('p').ASstring:=BscName;
        parambyname('q').ASstring:='1999';
        execsql;
        close;
        end;

        hh:='1999';
         with query2 do
         begin
           Close;
           Sql.Clear;
           Sql.Add('UPDATE  '+TName+'  SET  RE_DATE="'+hh+'"');
           sql.Add('WHERE   BSCNAME=:p');
           parambyname('p').ASstring:=BscName;
           ExecSql;
           close;
         end;
   except
   result:=0;
   end;
end;

procedure TFcdd_conv.ChangeFields;
Var   ch:string ;
      i,j:integer;
      begin
         Ch:='';
         Table1.open;
         for i:=0  to Table1.Fieldcount-3 do CH:=ch+'&'+Table1.Fields[i].FieldName;
         Table1.close;
         progressBar.position:=progressBar.position+2;
         with query1 do
         begin
           Close;
           Sql.Clear;
           Sql.Add('UPDATE  '+TName+'  SET DATA_CHANGE="'+CH+'"');
           ExecSql;
           close;
         end;

         with query2 do
         begin
           close;
           sql.clear;
           sql.Add('select *  from '+TName+' where BSCNAME=:p ');
           parambyname('p').ASstring:=BSCNAME;
           execsql;
         end;
         progressBar.position:=progressBar.position+1;

        if pos(TNAME,'RLCAP,RLOMP,RLDCP,RLLBP,RLLSP,RLTYP,RLSTP,RLDGP,RLNRP,RLSMP,RLSLP,RLPRP,RLPPP,RLSLP,RLSCP,RLVLP')<>0 then
        begin
        if pos(TNAME,'RLCAP,RLOMP,RLDCP,RLTYP,RLLBP,RLLSP,RLSCP')<>0 then  NonecellidChange;
        if pos(TNAME,'RLSTP,RLDGP')            <>0 then  stp_dgptblChange;
        if TNAME='RLNRP' then nrptblChange;
        if TNAME='RLSMP' then smptblChange;
        if TNAME='RLSLP' then slptblChange;
        if TNAME='RLPRP' then prptblChange;
        if TNAME='RLPPP' then ppptblChange;
        if TNAME='RLVLP' then vlptblChange;
        end
        else OnecellidChange;

    end;

    
procedure TFcdd_conv.getChangeField;
Var   ch,fna:string ;
      i:integer;
begin
             Ch:='';
             for i:=1  to Table1.Fieldcount-3 do
             begin
               fna:=Table1.Fields[i].FieldName;
               if  trim(table1.fieldbyname(fna).AsString)<> trim(Query2.fieldbyname(fna).AsString)
                        then
                        begin
                        CH:=ch+'&'+Table1.Fields[i].FieldName;
                        end;
             end;

             with query1 do
             begin
               Close;
               Sql.Clear;
               Sql.Add('UPDATE  '+TName+'  SET  DATA_CHANGE="'+CH+'"');
               ExecSql;
               close;
             end;

end;



procedure TFcdd_conv.nrptblChange;
Var   cel,cer:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cer:=query2.Fieldbyname('CELLR').Asstring;
           cel:=query2.Fieldbyname('CELLID').Asstring;

           if  Table1.Locate('CELLID;CELLR',VarArrayOf([cel,cer]), []) then   getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;


procedure TFcdd_conv.stp_dgptblChange;
Var  cel,chgr:string ;
      i,k:integer;
      begin
         k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cel:=query2.Fieldbyname('CELLID').Asstring;
           chgr:=query2.Fieldbyname('CHGR').Asstring;

           if  Table1.Locate('CELLID;CHGR',VarArrayOf([cel,chgr]), []) then  getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;

procedure TFcdd_conv.smptblChange;
Var   cel,chgr:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cel:=query2.Fieldbyname('CELLID').Asstring;
           chgr:=query2.Fieldbyname('SIMSG').Asstring;

           if  Table1.Locate('CELLID;SIMSG',VarArrayOf([cel,chgr]), []) then  getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;

procedure TFcdd_conv.slptblChange;
Var   cel,ct,cr,acl:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cel:=query2.Fieldbyname('CELLID').Asstring;
           ct:=query2.Fieldbyname('CHTYPE').Asstring;
           cr:=query2.Fieldbyname('CHRATE').Asstring;
           acl:=query2.Fieldbyname('ACL').Asstring;

           if  Table1.Locate('CELLID;CHTYPE;CHRATE;ACL',VarArrayOf([cel,ct,cr,acl]), []) then  getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;

procedure TFcdd_conv.prptblChange;
Var  cel,ct,cr:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cel:=query2.Fieldbyname('CELLID').Asstring;
           ct:=query2.Fieldbyname('CHTYPE').Asstring;
           cr:=query2.Fieldbyname('CHRATE').Asstring;

           if  Table1.Locate('CELLID;CHTYPE;CHRATE',VarArrayOf([cel,ct,cr]), []) then  getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;


procedure TFcdd_conv.ppptblChange;
Var   pp,prl,inac,probf:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           PP:=query2.Fieldbyname('PP').Asstring;
           Prl:=query2.Fieldbyname('PRL').Asstring;
           inac:=query2.Fieldbyname('INAC').Asstring;
           probf:=query2.Fieldbyname('PROBF').Asstring;
           if  Table1.Locate('PP;PRL;INAC;PROBF',VarArrayOf([pp,prl,inac,probf]), []) then  getChangeField;
           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;



procedure TFcdd_conv.NonecellidChange;
var
      i:integer;
begin
        query2.Open;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           getChangeField;
           query2.next;
        end; //while .....do .........
     progressBar.position:=2*lineNuM;
     Table1.close;
     query2.close;
end;

procedure TFcdd_conv.OnecellidChange;
var    cel:string;
       i,j,k:integer;
begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;
        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i;
           cel:=query2.Fieldbyname('CELLID').Asstring;
           if  Table1.Locate('CELLID',cel, []) then  getChangeField;
           query2.next;
        end; //while .....do .........

     Table1.close;
     query2.close;
end;


procedure TFcdd_conv.vlptblChange;
Var  Ch, cel,ct,fna:string ;
      i,k:integer;
      begin
        k:=progressBar.Max-progressBar.position;
        query2.Open;
        i:=query2.recordcount;
        if i=0 then progressBar.position:=progressBar.MAX
          else i:=k div i;
        if i=0 then i:=1;
        Query2.first;
        Table1.open;

        while not query2.EOf do
        begin
           progressBar.position:=progressBar.position+i; 
           cel:=query2.Fieldbyname('CELLID').Asstring;
           ct:=query2.Fieldbyname('CHTYPE').Asstring;

           if  Table1.Locate('CELLID;CHTYPE',VarArrayOf([cel,ct]), []) then
           begin
             Ch:='';
             for i:=2  to Table1.Fieldcount-3 do
             begin
               fna:=Table1.Fields[i].FieldName;
               if  trim(table1.fieldbyname(fna).AsString)<> trim(Query2.fieldbyname(fna).AsString)
                        then
                        begin
                        CH:=ch+'&'+Table1.Fields[i].FieldName;
                        end;
             end;

             with query1 do
             begin
               Close;
               Sql.Clear;
               Sql.Add('UPDATE  '+TName+'  SET  DATA_CHANGE="'+CH+'"');
               ExecSql;
               close;
             end;
           end;

           query2.next;
        end; //while .....do .........
        table1.close;
        query2.close;
    end;

Procedure TFcdd_conv.Refreshtable;
var fa: string;
begin
      Fa:='2000';
      with query1 do
         begin
         Close;
         Sql.Clear;
         Sql.Add('UPDATE  '+TName+'  SET RE_DATE="'+Fa+'"');
         ExecSql;
         close;
         end;

       if findtable(TName)=0 then
        begin
         progressBar.position:=progressBar.MAX;
         BatchMove.Mode := batCopy;
         moverecords;
         with query2 do
         begin
         Close;
         Sql.Clear;
         Sql.Add('CREATE INDEX '+TName+'INDEX ON '+TName+'(BSCNAME,RE_DATE)');
         ExecSql;
         close;
         end;
        end
        else
        begin
        if   Checktable=1 then
         begin
         ChangeFields;
         BatchMove.Mode := batAppend ;
         moverecords;
         end;
        end;
end;



procedure Tfcdd_conv.moverecords;
  begin
      Table3.close;
      Table3.TableName := TNAME;
      BatchMove.Source := table1;
      BatchMove.Destination := table3;
      table2.Close;
      table1.Close;
      table3.Close;
      //BatchMove.
      BatchMove.Execute;
  end;


procedure TFcdd_conv.FormCreate(Sender: TObject);
  begin
    Anm.Active := True;
 end;

end.
