unit AllCdd;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, ExtCtrls, Db, DBTables, Menus;

type
  TfmAllCdd = class(TForm)
    tvCdd: TTreeView;
    gbCellField: TGroupBox;
    lbCddField: TListBox;
    Panel1: TPanel;
    Panel2: TPanel;
    paDataList: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Query2: TQuery;
    Query1: TQuery;
    Query3: TQuery;
    Table1: TTable;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Query4: TQuery;
    Table2: TTable;
    procedure tvCddDblClick(Sender: TObject);
    procedure lbCddFieldDblClick(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure paDataListClick(Sender: TObject);
    procedure Panel2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Panel3Click(Sender: TObject);
    procedure Panel4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
  Procedure Createtable1;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmAllCdd: TfmAllCdd;
  wTableName : String;
  wSql:STRING;
  all:integer;
implementation

uses  BscData, DataList ;

{$R *.DFM}
Procedure TfmAllCdd.Createtable1;
begin
{  with TTable.Create(Application) do
  begin
    Active := False;

    TableName := 'TEMP_RLSMP';
    TableType := ttDefault;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SIMSG1', ftString, 10, False);
    FieldDefs.Add('SIMSG7', ftString, 10, False);
    FieldDefs.Add('SIMSG8', ftString, 10, False);
    CreateTable;
    Free;
  end; }
end;

procedure TfmAllCdd.tvCddDblClick(Sender: TObject);
begin
all:=0;
  if tvCdd.Selected.Parent <> nil then
  begin
    if gbCellField.Caption <> Copy(tvCdd.Selected.Parent.Text,1,5) THEN
    BEGIN
      lbCddField.Items.Clear;
    END;
    if (gbCellField.Caption <> Copy(tvCdd.Selected.Parent.Text,1,5)) AND (   (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLDCP') AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLTYP')  AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLCAP')  AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLLBP')   AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLLSP')   AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLOMP')   AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLPPP')   AND  (Copy(tvCdd.Selected.Parent.Text,1,5)<>'RLSCP'))   then
    begin
      lbCddField.Items.Clear;
      lbCddField.Items.add('CELLID');
      gbCellField.Caption := Copy(tvCdd.Selected.Parent.Text, 1, 5);
      wTableName := gbCellField.Caption;
    end;
    if (gbCellField.Caption <> Copy(tvCdd.Selected.Parent.Text,1,5)) AND (   (Copy(tvCdd.Selected.Parent.Text,1,5)='RLDCP') or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLTYP')  or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLCAP')  or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLLBP')   or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLLSP')   or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLOMP')   or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLPPP')   or  (Copy(tvCdd.Selected.Parent.Text,1,5)='RLSCP'))   then
    begin
      lbCddField.Items.Clear;
      lbCddField.Items.add('BSCNAME');
      gbCellField.Caption := Copy(tvCdd.Selected.Parent.Text, 1, 5);
      wTableName := gbCellField.Caption;
    end;

    if lbCddField.Items.IndexOf(tvCdd.Selected.Text) < 0 then
    begin
      gbCellField.Caption := Copy(tvCdd.Selected.Parent.Text, 1, 5);
      wTableName := gbCellField.Caption;
      lbCddField.Items.Add(tvCdd.Selected.Text);
    end
  end;
end;

procedure TfmAllCdd.lbCddFieldDblClick(Sender: TObject);
begin
  if lbCddField.Items.Count > 0 then
  begin
    if lbCddField.ItemIndex >= 0 then
      lbCddField.Items.Delete(lbCddField.ItemIndex)
    else
      lbCddField.Items.Delete(0);
  end;
end;

procedure TfmAllCdd.FormPaint(Sender: TObject);
begin
 { paDataList.Width := Round(fmAllCdd.Width /3);
  panel2.Width := Round(fmAllCdd.Width /3);
  panel3.Width := Round(fmAllCdd.Width /3); }
end;

procedure TfmAllCdd.paDataListClick(Sender: TObject);
var
  i,j : Integer;
  tt:ttable;
  t1,t2,t3:integer;
begin
  fmDataList.Caption := wTableName;
  if wtablename<>'RLSMP' THEN
  BEGIN
  if all=1 then
  begin
  wsql:='';
  table2.TableName:=wTableName;
  table2.Active:=true;
  with table2 do
  begin
    for j := 0 to FieldCount -3 do
    begin
      if j < FieldCount -3 then
        wSql := wSql + Fields[j].FieldName + ','
      else
        wSql := wSql + Fields[j].FieldName;
    end;
  end;
  table2.Close;
  end  else
  begin
  if lbCddField.Items.Count > 0 then
  begin
    wSql := '';
    for i := 0 to  lbCddField.Items.Count - 1 do
    begin
      if i < lbCddField.Items.Count - 1 then
        wSql := wSql + lbCddField.Items.Strings[i] + ','
      else
        wSql := wSql + lbCddField.Items.Strings[i];
    end;
  end;
  end;
  with fmDataList.quDataList do
  begin
    if Active then
      Close;
    IF Pos('CELLID', WSQL) <>0 then
    BEGIN
      with sql do
      begin
        Clear;
        Add('select '+wsql+'  from ' + wTableName +' WHERE RE_DATE=:P order by cellid');
        PARAMBYNAME('P').ASSTRING:='2000';
      end;
      Open;
    END  ELSE
    BEGIN
      with sql do
      begin
        Clear;
        Add('select '+wsql+' from ' + wTableName +' WHERE RE_DATE=:P order by BSCNAME');
        PARAMBYNAME('P').ASSTRING:='2000';
      end;
      Open;
    END;
  end;
  fmDataList.Show;
  END ELSE
  BEGIN
  if lbCddField.Items.Count > 0 then
  begin
    wSql := '';
    for i := 0 to  lbCddField.Items.Count - 1 do
    begin
      if i < lbCddField.Items.Count - 1 then
        wSql := wSql + lbCddField.Items.Strings[i] + ','
      else
        wSql := wSql + lbCddField.Items.Strings[i];
    end;
  end;

  with QUERY1 do
  begin
    if Active then
      Close;
    IF Pos('CELLID', WSQL) <>0 then
    BEGIN
      with sql do
      begin
        Clear;
        Add('select * from ' + wTableName +' WHERE RE_DATE=:P order by cellid');
        PARAMBYNAME('P').ASSTRING:='2000';
      end;
      Open;
    END  ELSE
    BEGIN
      with sql do
      begin
        Clear;
        Add('select * from ' + wTableName +' WHERE RE_DATE=:P order by BSCNAME');
        PARAMBYNAME('P').ASSTRING:='2000';
      end;
      Open;
    END;
  end;
   with Query2 do
         begin
           Close;

           SQL.Clear;
           SQL.Add('delete from temp_RLSMP');
           execsql;
         end;

  TABLE1.open;
  table1.edit;

  t1:=query1.RecordCount;
  for t2:=1 to t1 do
  begin
    WHILE NOT QUERY1.Eof DO
    BEGIN
    table1.append;
    table1.fieldbyname('CELLID').asstring:=query1.fieldbyname('CELLID').asstring;
    table1.fieldbyname('BSCNAME').asstring:=query1.fieldbyname('BSCNAME').asstring;
    table1.fieldbyname('SIMSG1').asstring:=query1.fieldbyname('MSGDIST').asstring;
    QUERY1.Next;
    table1.fieldbyname('SIMSG7').asstring:=query1.fieldbyname('MSGDIST').asstring;
    QUERY1.Next;
    table1.fieldbyname('SIMSG8').asstring:=query1.fieldbyname('MSGDIST').asstring;
    QUERY1.Next;
    END;
  end;
  with fmDataList.quDataList do
         begin
           Close;

           SQL.Clear;
           SQL.Add('Select *');
           SQL.Add('From temp_RLSMP ');
           Open;
         end;
  fmDataList.Show;
  END;
end;

procedure TfmAllCdd.Panel2Click(Sender: TObject);
begin
  lbCddField.Items.Clear;
  gbCellField.Caption := '选择参数';
end;

procedure TfmAllCdd.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //fmDataList.Close;
  //fmDataList.free;
  Action := caHide;
end;

procedure TfmAllCdd.Panel3Click(Sender: TObject);
begin
//11
IF  tvCdd.Selected.Text='BSPWRB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH载频发射功率');
  END
ELSE IF  tvCdd.Selected.Text='CGI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区全球识别码');
  END
ELSE IF  tvCdd.Selected.Text='BSIC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('基站识别码');
  END
ELSE IF  tvCdd.Selected.Text='BCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH载波频率');
  END
//22
ELSE IF  tvCdd.Selected.Text='BCCHTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH组合类型');
  END
ELSE IF  tvCdd.Selected.Text='AGBLK'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入允许保留块数');
  END
ELSE IF  tvCdd.Selected.Text='MFRMS'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('寻呼复帧数');
  END
ELSE IF  tvCdd.Selected.Text='FNOFFSET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('帧偏置');
  END
//33
ELSE IF  tvCdd.Selected.Text='MSTXPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('移动站最大发射功率');
  END
ELSE IF  tvCdd.Selected.Text='HOP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('跳频状态');
  END
ELSE IF  tvCdd.Selected.Text='HSN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('跳频序列号');
  END
ELSE IF  tvCdd.Selected.Text='SDCCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('SDCCH/8信道数');
  END
//44
ELSE IF  tvCdd.Selected.Text='CBCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区广播信道');
  END
ELSE IF  tvCdd.Selected.Text='ACCMIN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最小接入电平');
  END
ELSE IF  tvCdd.Selected.Text='CCHPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('控制信道最大发射功率');
  END
ELSE IF  tvCdd.Selected.Text='CRH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选滞后');
  END
//55
ELSE IF  tvCdd.Selected.Text='NCCPERM'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('允许的网络色码');
  END
ELSE IF  tvCdd.Selected.Text='SIMSG和MSGDIST'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH系统消息开关');
  END
ELSE IF  tvCdd.Selected.Text='CB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区接入禁止');
  END
ELSE IF  tvCdd.Selected.Text='CBQ'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区禁止限制');
  END
//66
ELSE IF  tvCdd.Selected.Text='ACC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入控制等级');
  END
ELSE IF  tvCdd.Selected.Text='MAXRET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('最大重发次数');
  END
ELSE IF  tvCdd.Selected.Text='TX'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('发送分布时隙数');
  END
ELSE IF  tvCdd.Selected.Text='ATT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('IMSI结合分离允许');
  END
//77
ELSE IF  tvCdd.Selected.Text='T3212'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('周期位置更新定时器');
  END
ELSE IF  tvCdd.Selected.Text='CRO TO PT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区重选偏置、临时偏置和惩罚时间');
  END
ELSE IF  tvCdd.Selected.Text='EVALTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('算法类型');
  END
ELSE IF  tvCdd.Selected.Text='RLINKUP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行无线链路超时');
  END
//88
ELSE IF  tvCdd.Selected.Text='RLINKT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行无线链路超时');
  END
ELSE IF  tvCdd.Selected.Text='NECI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('新建原因指示');
  END
ELSE IF  tvCdd.Selected.Text='DMPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('MS动态功率控制状态');
  END
ELSE IF  tvCdd.Selected.Text='DBPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BTS动态功率控制状态');
  END
//99
ELSE IF  tvCdd.Selected.Text='DTXD'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('下行不连续发射');
  END
ELSE IF  tvCdd.Selected.Text='DTXU'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('上行不连续发射');
  END
ELSE IF  tvCdd.Selected.Text='IHO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('小区内切换开关');
  END
ELSE IF  tvCdd.Selected.Text='ASSOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('指配其它小区允许');
  END
//10
ELSE IF  tvCdd.Selected.Text='MBCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump(' BCCH频率表');
  END
ELSE IF  tvCdd.Selected.Text='LISTTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('频率表类型');
  END
ELSE IF  tvCdd.Selected.Text='ICMSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道测量状态');
  END
ELSE IF  tvCdd.Selected.Text='NOALLOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('信道分配开关');
  END
//////111
ELSE IF  tvCdd.Selected.Text='INTAVE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('空闲信道干扰电平平均周期');
  END
ELSE IF  tvCdd.Selected.Text='LIMITn'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('干扰带边界');
  END
ELSE IF  tvCdd.Selected.Text='MBCR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('多频段指示');
  END
ELSE IF  tvCdd.Selected.Text='ECSC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('CLASSMARK早送控制');
  END ELSE
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('接入控制等级');
END;
end;

procedure TfmAllCdd.Panel4Click(Sender: TObject);
begin
  close;
end;

procedure TfmAllCdd.FormShow(Sender: TObject);
var
  wTreeNode ,wTreeNode1,wTreeNode2: TTreeNode;
  i, j, wCount : Integer;
begin
  dmBscData.quRLCFP.Active:=true;
  Application.CreateForm(TfmDataList, fmDataList);
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
          for j := 0 to FieldCount -3 do
          begin
            with tvCdd.Items do
            wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName);//+'    '+fieldbyname(Fields[j].FieldName).asstring);
          end;
        end;
      end;
      if Pos('RLCPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCPP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            with tvCdd.Items do
            wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName);//+'    '+fieldbyname(Fields[j].FieldName).asstring);
          end;
        end;
      end;
      if Pos('RLCXP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCXP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            with tvCdd.Items do
            wTreeNode1:=AddChild(wTreeNode, Fields[j].FieldName);//+'    '+fieldbyname(Fields[j].FieldName).asstring);
          end;
        end;
      end;
      if Pos('RLDEP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDEP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLIHP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLIHP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLOP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLOP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLMFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLMFP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLNRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLNRP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSBP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSBP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSSP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSSP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
//////////////////////////////////////////////////////////////////////
      if Pos('RLCRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCRP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLDCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDCP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLDGP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDGP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLDTP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDTP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLDP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLDP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLHP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLHP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLPP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLUP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLUP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLOLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLOLP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLPCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPCP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
///////////////////////////////////////////////////////////////////////////
      if Pos('RLLAP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLAP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLCAP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCAP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLOMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLOMP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLTYP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLTYP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLBCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLBCP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLCP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLLP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLSP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLSP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
/////////////////////////////////////////////////////////////////////////
      if Pos('RLLBP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLBP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLFP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSTP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSTP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLACP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLACP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSMP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLIMP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLIMP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
//??????????????????????????????????????????????????????????????????????????
      if Pos('RLSLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSLP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
       // ++++++++++++++++++++++++++++++++++++++
      end;
      if Pos('RLVLP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLVLP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLHPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLHPP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLPPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPPP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLPRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLPRP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSCP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSCP do
        begin
          for j := 0 to FieldCount -3 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
    end;
    wCount := tvCdd.Items.Count;
    i := i + 1;
  end;
end;

procedure TfmAllCdd.N1Click(Sender: TObject);
var I,j:integer;
begin
  I:=TVCDD.Selected.Level;
  IF I=1  THEN
  begin
  wTableName:=tvCdd.Selected.Text;
//  showmessage(wtablename);
  lbCddField.Items.Clear;
  table2.TableName:=wTableName;
  table2.Active:=true;
  with table2 do
  begin
    for j := 0 to FieldCount -3 do
    begin
      lbCddField.Items.Add( Fields[j].FieldName);

    end;
  end;
  gbCellField.Caption := Copy(tvCdd.Selected.Parent.Text, 1, 5);
//  wTableName := gbCellField.Caption;
  all:=1;
  table2.Close;
end;
end;

procedure TfmAllCdd.FormActivate(Sender: TObject);
begin
all:=0;
end;

procedure TfmAllCdd.FormCreate(Sender: TObject);
begin
  MoveWindow(Handle, Screen.Width - Width, Screen.Height - Height, Width, Height, True);
end;

end.
