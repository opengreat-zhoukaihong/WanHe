unit detail;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Db, DBTables, Buttons, ExtCtrls;

type
  Tfdetail = class(TForm)
    tvcdd: TTreeView;
    l1: TLabel;
    Query1: TQuery;
    Query2: TQuery;
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Panel1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
  function space(space1:integer):string;
    { Public declarations }
  end;

var
  fdetail: Tfdetail;

implementation

uses DataList, BscData, BscMain;

{$R *.DFM}
function Tfdetail.space(space1:integer):string;
var  i:integer;
begin
  result:='';
  for i:=1 to space1 do
  begin
    result:=result+' ';
  end;
end;

procedure Tfdetail.FormShow(Sender: TObject);
var
  wTreeNode,wTreeNode1 : TTreeNode;
  i, j, wCount ,len1,MM,len2: Integer;
  ss,TT,s1,s2 : string;
begin
try
  oleMapInfo.do('fetch rec ' + IntToStr(1) +' from selection');

  ss := oleMapInfo.eval('selection.bs_no');
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
    TT:=QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

  wCount := tvCdd.Items.Count;
  i := 0;
  while i < wCount do
  begin
    wTreeNode := tvCdd.Items[i];
    if wTreeNode.Parent <> nil then
    begin
      if Pos('RLCFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCFP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;

      if Pos('RLCPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCPP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;

      if Pos('RLCXP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCXP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcxp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcxp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLDEP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDEP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldep where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldep where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLIHP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLIHP do
        begin
          for j := 2 to FieldCount -3 do
          begin

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlihp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlihp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLOP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLOP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllop where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllop where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
///////////////////////////.............................?????????//
      if Pos('RLMFP', wTreeNode.Text) > 0 then
      begin

       tvCdd.Items.AddChild(wTreeNode,'LISTTYPE       IDLE');
        with dmBscData.quRLMFP do
        begin
          for j := 4 to 19 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlmfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlmfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
        tvCdd.Items.AddChild(wTreeNode,'LISTTYPE       ACTIVE');
        with dmBscData.quRLMFP do
        begin
          for j := 20 to 35 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlmfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlmfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      if Pos('RLNRP', wTreeNode.Text) > 0 then
      begin
   with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select CELLid,bscname, CELLR, DIR, CAND, CS, KHYST, KOFFSET , LHYST ,LOFFSET ,TRHYST ,TROFFSET ,AWOFFSET ,BQOFFSET, HIHYST ,LOHYST, OFFSET,re_date,data_change   from rlnrp');
      sql.add('  where cellid=:cellid  and re_date=:qq');
      sql.add('  group by CELLid,bscname, CELLR, DIR, CAND, CS, KHYST, KOFFSET , LHYST ,LOFFSET ,TRHYST ,TROFFSET ,AWOFFSET ,BQOFFSET, HIHYST ,LOHYST, OFFSET,re_date,data_change   ');
      parambyname('cellid').asstring:=SS;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';

      open;
    end;
    MM:=1;
    query1.First;
   with query2 do
    begin
      close;
      sql.clear;
      sql.add('  select CELLid,bscname, CELLR, DIR, CAND, CS, KHYST, KOFFSET , LHYST ,LOFFSET ,TRHYST ,TROFFSET ,AWOFFSET ,BQOFFSET, HIHYST ,LOHYST, OFFSET,re_date,data_change   from rlnrp');
      sql.add('  where cellid=:cellid  and re_date=:qq');
      sql.add('  group by CELLid,bscname, CELLR, DIR, CAND, CS, KHYST, KOFFSET , LHYST ,LOFFSET ,TRHYST ,TROFFSET ,AWOFFSET ,BQOFFSET, HIHYST ,LOHYST, OFFSET,re_date,data_change   ');
      parambyname('cellid').asstring:=SS;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
   query2.First;
   WHILE NOT query1.Eof DO
   BEGIN
        tvCdd.Items.AddChild(wTreeNode,INTTOSTR(MM));
        with dmBscData.quRLNRP do
        begin
          for j := 2 to FieldCount-3 do
          begin

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(query1.fieldbyname(Fields[j].FieldName).asstring);
    len2:=length(trimleft(query1.fieldbyname(Fields[j].FieldName).asstring));
    if not QUERY2.Eof then
    s2:=space(25-len2)+trimleft(QUERY2.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName+s1+s2);

          end;
        end;
        MM:=MM+1;
   query1.Next;
   query2.Next;
   END;
      end;

      if Pos('RLSBP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSBP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlsbp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlsbp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLSSP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSSP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlssp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlssp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
//////////////////////////////////////////////////////////////////////
      if Pos('RLCRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCRP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcrp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcrp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLDCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDCP do
        begin
          for j := 1 to FieldCount -3 do
          begin
   with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldcp WHERE BSCNAME=:PP and re_date=:qq');
      parambyname('PP').asstring:=TT;
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
   with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldcp WHERE BSCNAME=:PP and re_date=:qq');
      parambyname('PP').asstring:=TT;
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
    //    QUERY1.Next;
    if  QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLDGP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDGP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldgp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldgp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLDTP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDTP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldtp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rldtp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLDP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLDP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlldp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlldp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLHP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLHP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllhp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllhp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLPP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLUP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLUP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllup where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllup where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLOLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLOLP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlolp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlolp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
////////////////////////////??????????????????????????????!!!!!!!!!!!!!!!!!!!
      if Pos('RLLAP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLAP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllap where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllap where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLCAP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCAP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcAp  where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlcAp  where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
//    QUERY1.Next;
    if QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLOMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLOMP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlOMp WHERE BSCNAME=:PP and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlOMp WHERE BSCNAME=:PP and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
//    QUERY1.Next;
    if  QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLTYP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLTYP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlTYp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlTYp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
//    QUERY1.Next;
    if  QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLBCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLBCP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlbcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlbcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLCP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLLP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlllp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlllp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLSP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLSP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlLSp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
//    QUERY1.Next;
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlLSp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
    if  QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLBP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLBP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlLBp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlLBp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
    //    QUERY1.Next;
    if  QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLLFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLFP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rllfp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLSTP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSTP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlstp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlstp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLACP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLACP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlacp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlacp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLSMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSMP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlsmp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlsmp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLIMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLIMP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlimp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlimp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLSLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSLP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlslp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlslp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLVLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLVLP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlvlp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlvlp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLHPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLHPP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlhpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlhpp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLPPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPPP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlPpp where bscname=:pp and re_date=:qq ');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlPpp where bscname=:pp and re_date=:qq ');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
//    QUERY1.Next;
    if not QUERY1.Eof then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLPRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPRP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlprp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlprp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
      if Pos('RLSCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSCP do
        begin
          for j := 1 to FieldCount -3 do
          begin
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlSCp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='2000';
      open;
    end;
    query1.First;
    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));
    with query1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlSCp where bscname=:pp and re_date=:qq');
      parambyname('pp').asstring:=tt;//'ZHAHZS3';
      parambyname('qq').asstring:='1999';
      open;
    end;
    query1.First;
//    QUERY1.Next;
    if QUERY1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
//..............
      if Pos('RLPCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPCP do
        begin
          for j := 2 to FieldCount -3 do
          begin
    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlpcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='2000';
      open;
    end;
    QUERY1.First;
//    L1.Caption:='小区：'+QUERY1.FIELDBYNAME('CELLID').ASSTRING+' 归属局：'+QUERY1.FIELDBYNAME('BSCNAME').ASSTRING;

    len1:=length(Fields[j].FieldName);
    s1:=space(15-len1)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring);
    len2:= length(trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring));

    with QUERY1 do
    begin
      close;
      sql.clear;
      sql.add('  select * from rlpcp where cellid=:cellid and re_date=:p');
      parambyname('cellid').asstring:=ss;//'ZHAHZS3';
      PARAMBYNAME('P').ASSTRING:='1999';
      open;
    end;
    QUERY1.First;
    if query1.RecordCount<>0 then
    s2:=space(25-len2)+trimleft(QUERY1.fieldbyname(Fields[j].FieldName).asstring)
    else s2:='';
    with tvCdd.Items do
       wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName+s1+s2);
          end;
        end;
      end;
     end;
    wCount := tvCdd.Items.Count;
    i := i + 1;
  end;
except

end;
end;



procedure Tfdetail.SpeedButton1Click(Sender: TObject);
begin
//11
//IF  tvCdd.Selected.Text='BSPWRB'  THEN
  if pos('BSPWRB',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH载频发射功率');
  END
//IF  tvCdd.Selected.Text='CGI'  THEN
ELSE  if pos('CGI',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('小区全球识别码');
  END
//IF  tvCdd.Selected.Text='BSIC'  THEN
ELSE  if pos('BSIC',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('基站识别码');
  END
//IF  tvCdd.Selected.Text='BCCHNO'  THEN
ELSE  if pos('BCCHNO',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH载波频率');
  END
//22
//IF  tvCdd.Selected.Text='BCCHTYPE'  THEN
ELSE  if pos('BCCHTYPE',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH组合类型');
  END
//IF  tvCdd.Selected.Text='AGBLK'  THEN
ELSE  if pos('AGBLK',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('接入允许保留块数');
  END
//IF  tvCdd.Selected.Text='MFRMS'  THEN
ELSE  if pos('MFRMS',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('寻呼复帧数');
  END
//IF  tvCdd.Selected.Text='FNOFFSET'  THEN
ELSE  if pos('FNOFFSET',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('帧偏置');
  END
//33
//IF  tvCdd.Selected.Text='MSTXPWR'  THEN
ELSE  if pos('MSTXPWR',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('移动站最大发射功率');
  END
//IF  tvCdd.Selected.Text='HOP'  THEN
//showmessage(tvCdd.Selected.Text);
ELSE  if pos('HOP',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('跳频状态');
  END
//IF  tvCdd.Selected.Text='HSN'  THEN
ELSE  if pos('HSN',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('跳频序列号');
  END
//IF  tvCdd.Selected.Text='SDCCH'  THEN
ELSE  if pos('SDCCH',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('SDCCH/8信道数');
  END
//44
//IF  tvCdd.Selected.Text='CBCH'  THEN
ELSE  if pos('CBCH',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区广播信道');
  END
//IF  tvCdd.Selected.Text='ACCMIN'  THEN
ELSE  if pos('ACCMIN',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最小接入电平');
  END
//IF  tvCdd.Selected.Text='CCHPWR'  THEN
ELSE  if pos('CCHPWR',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('控制信道最大发射功率');
  END
//IF  tvCdd.Selected.Text='CRH'  THEN
ELSE  if pos('CRH',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选滞后');
  END
//55
//IF  tvCdd.Selected.Text='NCCPERM'  THEN
ELSE  if pos('NCCPERM',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('允许的网络色码');
  END
//IF  tvCdd.Selected.Text='SIMSG和MSGDIST'  THEN
ELSE  if pos('SIMSG',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH系统消息开关');
  END
//IF  tvCdd.Selected.Text='CB'  THEN
ELSE  if pos('CB',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区接入禁止');
  END
//IF  tvCdd.Selected.Text='CBQ'  THEN
ELSE  if pos('CBQ',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区禁止限制');
  END
//66
//IF  tvCdd.Selected.Text='ACC'  THEN
ELSE  if pos('ACC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入控制等级');
  END
//IF  tvCdd.Selected.Text='MAXRET'  THEN
ELSE  if pos('MAXRET',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最大重发次数');
  END
//IF  tvCdd.Selected.Text='TX'  THEN
ELSE  if pos('TX',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('发送分布时隙数');
  END
//IF  tvCdd.Selected.Text='ATT'  THEN
ELSE  if pos('ATT',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('IMSI结合分离允许');
  END
//77
//IF  tvCdd.Selected.Text='T3212'  THEN
ELSE  if pos('T3212',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('周期位置更新定时器');
  END
//IF  tvCdd.Selected.Text='CRO TO PT'  THEN
ELSE  if pos('CRO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选偏置、临时偏置和惩罚时间');
  END
//IF  tvCdd.Selected.Text='EVALTYPE'  THEN
ELSE  if pos('EVALTYPE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('算法类型');
  END
//IF  tvCdd.Selected.Text='RLINKUP'  THEN
ELSE  if pos('RLINKUP',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行无线链路超时');
  END
//88
//IF  tvCdd.Selected.Text='RLINKT'  THEN
ELSE  if pos('RLINKT',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行无线链路超时');
  END
//IF  tvCdd.Selected.Text='NECI'  THEN
ELSE  if pos('NECI',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('新建原因指示');
  END
//IF  tvCdd.Selected.Text='DMPSTATE'  THEN
ELSE  if pos('DMPSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('MS动态功率控制状态');
  END
//IF  tvCdd.Selected.Text='DBPSTATE'  THEN
ELSE  if pos('DBPSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BTS动态功率控制状态');
  END
//99
//IF  tvCdd.Selected.Text='DTXD'  THEN
ELSE  if pos('DTXD',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行不连续发射');
  END
//IF  tvCdd.Selected.Text='DTXU'  THEN
ELSE  if pos('DTXU',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行不连续发射');
  END
//IF  tvCdd.Selected.Text='IHO'  THEN
ELSE  if pos('IHO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区内切换开关');
  END
//IF  tvCdd.Selected.Text='ASSOC'  THEN
ELSE  if pos('ASSOC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('指配其它小区允许');
  END
//10
//IF  tvCdd.Selected.Text='MBCCHNO'  THEN
ELSE  if pos('MBCCHNO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump(' BCCH频率表');
  END
//IF  tvCdd.Selected.Text='LISTTYPE'  THEN
ELSE  if pos('LISTTYPE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('频率表类型');
  END
//IF  tvCdd.Selected.Text='ICMSTATE'  THEN
ELSE  if pos('ICMSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道测量状态');
  END
//IF  tvCdd.Selected.Text='NOALLOC'  THEN
ELSE  if pos('NOALLOC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('信道分配开关');
  END
//////111
//IF  tvCdd.Selected.Text='INTAVE'  THEN
ELSE  if pos('INTAVE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道干扰电平平均周期');
  END
//IF  tvCdd.Selected.Text='LIMITn'  THEN
ELSE  if pos('LIMITn',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('干扰带边界');
  END
//IF  tvCdd.Selected.Text='MBCR'  THEN
ELSE  if pos('MBCR',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('多频段指示');
  END
//IF  tvCdd.Selected.Text='ECSC'  THEN
ELSE  if pos('ECSC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('CLASSMARK早送控制');
  END
ELSE
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入控制等级');
END;

end;

procedure Tfdetail.Panel1Click(Sender: TObject);
begin
//11
//IF  tvCdd.Selected.Text='BSPWRB'  THEN
  if pos('BSPWRB',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH载频发射功率');
  END;
//IF  tvCdd.Selected.Text='CGI'  THEN
  if pos('CGI',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('小区全球识别码');
  END;
//IF  tvCdd.Selected.Text='BSIC'  THEN
  if pos('BSIC',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('基站识别码');
  END;
//IF  tvCdd.Selected.Text='BCCHNO'  THEN
  if pos('BCCHNO',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH载波频率');
  END;
//22
//IF  tvCdd.Selected.Text='BCCHTYPE'  THEN
  if pos('BCCHTYPE',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('BCCH组合类型');
  END;
//IF  tvCdd.Selected.Text='AGBLK'  THEN
  if pos('AGBLK',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('接入允许保留块数');
  END;
//IF  tvCdd.Selected.Text='MFRMS'  THEN
  if pos('MFRMS',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('寻呼复帧数');
  END;
//IF  tvCdd.Selected.Text='FNOFFSET'  THEN
  if pos('FNOFFSET',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('帧偏置');
  END;
//33
//IF  tvCdd.Selected.Text='MSTXPWR'  THEN
  if pos('MSTXPWR',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('移动站最大发射功率');
  END;
//IF  tvCdd.Selected.Text='HOP'  THEN
//showmessage(tvCdd.Selected.Text);
  if pos('HOP',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('跳频状态');
  END;
//IF  tvCdd.Selected.Text='HSN'  THEN
  if pos('HSN',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('跳频序列号');
  END;
//IF  tvCdd.Selected.Text='SDCCH'  THEN
  if pos('SDCCH',tvCdd.Selected.Text)<>0 then
  BEGIN
    application.HelpFile:='CDD.hlp';
    application.HelpJump('SDCCH/8信道数');
  END;
//44
//IF  tvCdd.Selected.Text='CBCH'  THEN
  if pos('CBCH',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区广播信道');
  END;
//IF  tvCdd.Selected.Text='ACCMIN'  THEN
  if pos('ACCMIN',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最小接入电平');
  END;
//IF  tvCdd.Selected.Text='CCHPWR'  THEN
  if pos('CCHPWR',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('控制信道最大发射功率');
  END;
//IF  tvCdd.Selected.Text='CRH'  THEN
  if pos('CRH',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选滞后');
  END;
//55
//IF  tvCdd.Selected.Text='NCCPERM'  THEN
  if pos('NCCPERM',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('允许的网络色码');
  END;
//IF  tvCdd.Selected.Text='SIMSG和MSGDIST'  THEN
  if pos('SIMSG',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH系统消息开关');
  END;
//IF  tvCdd.Selected.Text='CB'  THEN
  if pos('CB',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区接入禁止');
  END;
//IF  tvCdd.Selected.Text='CBQ'  THEN
  if pos('CBQ',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区禁止限制');
  END;
//66
//IF  tvCdd.Selected.Text='ACC'  THEN
  if pos('ACC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入控制等级');
  END;
//IF  tvCdd.Selected.Text='MAXRET'  THEN
  if pos('MAXRET',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最大重发次数');
  END;
//IF  tvCdd.Selected.Text='TX'  THEN
  if pos('TX',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('发送分布时隙数');
  END;
//IF  tvCdd.Selected.Text='ATT'  THEN
  if pos('ATT',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('IMSI结合分离允许');
  END;
//77
//IF  tvCdd.Selected.Text='T3212'  THEN
  if pos('T3212',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('周期位置更新定时器');
  END;
//IF  tvCdd.Selected.Text='CRO TO PT'  THEN
  if pos('CRO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选偏置、临时偏置和惩罚时间');
  END;
//IF  tvCdd.Selected.Text='EVALTYPE'  THEN
  if pos('EVALTYPE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('算法类型');
  END;
//IF  tvCdd.Selected.Text='RLINKUP'  THEN
  if pos('RLINKUP',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行无线链路超时');
  END;
//88
//IF  tvCdd.Selected.Text='RLINKT'  THEN
  if pos('RLINKT',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行无线链路超时');
  END;
//IF  tvCdd.Selected.Text='NECI'  THEN
  if pos('NECI',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('新建原因指示');
  END;
//IF  tvCdd.Selected.Text='DMPSTATE'  THEN
  if pos('DMPSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('MS动态功率控制状态');
  END;
//IF  tvCdd.Selected.Text='DBPSTATE'  THEN
  if pos('DBPSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BTS动态功率控制状态');
  END;
//99
//IF  tvCdd.Selected.Text='DTXD'  THEN
  if pos('DTXD',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行不连续发射');
  END;
//IF  tvCdd.Selected.Text='DTXU'  THEN
  if pos('DTXU',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行不连续发射');
  END;
//IF  tvCdd.Selected.Text='IHO'  THEN
  if pos('IHO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区内切换开关');
  END;
//IF  tvCdd.Selected.Text='ASSOC'  THEN
  if pos('ASSOC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('指配其它小区允许');
  END;
//10
//IF  tvCdd.Selected.Text='MBCCHNO'  THEN
  if pos('MBCCHNO',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump(' BCCH频率表');
  END;
//IF  tvCdd.Selected.Text='LISTTYPE'  THEN
  if pos('LISTTYPE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('频率表类型');
  END;
//IF  tvCdd.Selected.Text='ICMSTATE'  THEN
  if pos('ICMSTATE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道测量状态');
  END;
//IF  tvCdd.Selected.Text='NOALLOC'  THEN
  if pos('NOALLOC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('信道分配开关');
  END;
//////111
//IF  tvCdd.Selected.Text='INTAVE'  THEN
  if pos('INTAVE',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道干扰电平平均周期');
  END;
//IF  tvCdd.Selected.Text='LIMITn'  THEN
  if pos('LIMITn',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('干扰带边界');
  END;
//IF  tvCdd.Selected.Text='MBCR'  THEN
  if pos('MBCR',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('多频段指示');
  END;
//IF  tvCdd.Selected.Text='ECSC'  THEN
  if pos('ECSC',tvCdd.Selected.Text)<>0 then
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('CLASSMARK早送控制');
  END;


end;

procedure Tfdetail.SpeedButton2Click(Sender: TObject);
begin
close;
end;

end.
