
unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Buttons, ImgList, ExtCtrls, Grids, DBGrids, Db,
  DBTables, DBCtrls, ToolWin, Menus;

type
  TForm1 = class(TForm)
    ImageList1: TImageList;
    DataSource1: TDataSource;
    QuyMsc: TQuery;
    Table1: TTable;
    QuyBsc: TQuery;
    QuyBase: TQuery;
    QuyCell: TQuery;
    Query1: TQuery;
    Table2: TTable;
    DBNavigator1: TDBNavigator;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    StatusBar1: TStatusBar;
    Panel1: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    TreeView1: TTreeView;
    StatusBar2: TStatusBar;
    StatusBar3: TStatusBar;
    DBGrid1: TDBGrid;
    BatchMove: TBatchMove;
    ImageList2: TImageList;
    BitBtn1: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TreeView1Expanded(Sender: TObject; Node: TTreeNode);
    procedure TreeView1Collapsed(Sender: TObject; Node: TTreeNode);
    procedure Table1AfterPost(DataSet: TDataSet);
    procedure DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
    procedure Table1AfterDelete(DataSet: TDataSet);
    procedure Table1BeforeDelete(DataSet: TDataSet);
    procedure Table1BeforeEdit(DataSet: TDataSet);
    procedure Table1AfterCancel(DataSet: TDataSet);
    procedure Table1BeforeInsert(DataSet: TDataSet);

  private
    { Private declarations }
    procedure realtblchange;
    procedure RealNodechange;
    procedure deleNoderecord;
    procedure gettemptbl;
    function  TreefindItem(NodeItem:TTreenode;NO:string):TTreeNode;
    procedure AddBscNode(MSC_no:string);
    procedure AddBaseNode(BSC_no:string);
    procedure AddcellNode(BS_no:string);
    procedure showMSC0(var Name:string);
    procedure showMSC(var Name:string);
    procedure showbsc(var Name:string);
    procedure showbase(var Name:string);
    procedure showcell(var Name:string);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  var AllNode,MscNode,BscNode,BaseNode,CellNode,deleNode:TtreeNode;
      TempName,NNO,state,place:string;
      dele:Word;
implementation

{$R *.DFM}


procedure TForm1.gettemptbl;
begin
 Table1.TableName := 'temp';
 BatchMove.Source := query1;
 BatchMove.Destination := table1;
 BatchMove.Mode := batcopy;
 BatchMove.Execute;
end;


function  TForm1.TreefindItem(NodeItem:TTreenode;NO:string):TTreeNode;
var parent:TTreeNode;
    n:Integer;
begin
 n:=NodeItem.level;
 try
 parent:=NodeItem.getFirstChild ;
 if (parent<>nil) and (parent.Text <> NO) then
 repeat
 parent:=parent.getNextSibling ;
 until (parent=nil) or (parent.Text = NO);
 except
  parent:=nil;
 end;
  result:=parent;
end;

procedure TForm1.FormCreate(Sender: TObject);
var    Msc_no:String;
begin
place:='珠海移动局';
 StatusBar2.Panels[0].Text :='  所有文件';
 with  treeview1.Items  do
 begin
   beginUpdate;
   AllNode:=add(nil,Place);
   allNode.ImageIndex :=allNode.level+1;
   allNode.selectedindex:=0;
   with quyMsc do
   begin
    close;
    sqL.clear;
    Sql.add('select MSC_NO,MSC_Name from MSC.dbf');
    execsql;
    open;
    first;

    while not Eof do
     begin
      Msc_no:=fieldbyName('MSC_NO').Asstring;
      MscNode:=AddChild(AllNode,fieldbyName('MSC_NO').Asstring);
      MscNode.ImageIndex :=MscNode.level+1;
      allNode.selectedindex:=0;
      AddBscNode(MSC_no);
      Next;
     end;
     EndUpDate;
     close;
    end;

 end;
  treeview1.Items[0].selected:=true;
  treeview1.Items[0].Expanded:=true;
end;


procedure TForm1.AddBscNode(MSC_no:string);
var Bsc_No:string;
begin
with  treeview1.Items  do
begin
   with quyBsc do
   begin
    close;
    sqL.clear;
    Sql.add('select BSC_NO,BSC_Name from bsc.dbf');
    Sql.add('Where Msc_No="'+MSC_no+'"');
    execsql;
    open;
    first;

    while not Eof do
     begin
      Bsc_no:=fieldbyName('BSC_NO').Asstring;
      BscNode:=AddChild(MSCNode,fieldbyName('BSC_NO').Asstring);
      AddBaseNode(BSC_no);
      BscNode.ImageIndex :=BscNode.level+1;
      allNode.selectedindex:=0;
      Next;
     end;
     close;
   end;
 end;
end;


procedure TForm1.AddBaseNode(BSC_no:string);
var Bs_No:string;
begin
with  treeview1.Items  do
begin
   with quyBase do
   begin
    close;
    sqL.clear;
    Sql.add('select Bs_NO,Bs_Name from base.dbf');
    Sql.add('Where Bsc_No="'+BSC_no+'"');
    execsql;
    open;
    first;
    while not Eof do
     begin
      Bs_no:=fieldbyName('Bs_NO').Asstring;
      BaseNode:=AddChild(BSCNode,BS_NO+'('+fieldbyName('BS_NAME').Asstring+')');
      baseNode.ImageIndex :=baseNode.level+1;
      allNode.selectedindex:=0;
      AddcellNode(bs_no);
      Next;
     end;
     close;
   end;
 end;

end;



procedure TForm1.AddcellNode(BS_no:string);
begin
with  treeview1.Items  do
begin
   with quycell do
   begin
    close;
    sqL.clear;
    Sql.add('select Cell_Name,BS_NO from cell.dbf');
    Sql.add('Where BAse_No="'+BS_no+'"');
    execsql;
    open;
    first;
    while not Eof do
     begin
      CellNode:=AddChild(BaseNode,fieldbyName('BS_NO').Asstring);
      cellNode.ImageIndex :=cellnode.level+1;
      cellNode.selectedindex:=0;
      Next;
     end;
     close;
   end;
 end;
end;


procedure TForm1.BitBtn1Click(Sender: TObject);
begin
close;
end;

procedure TForm1.TreeView1Change(Sender: TObject; Node: TTreeNode);
var  position: integer;
     selectNode:TTreeNode;
     text:string;
begin
 if  Node.Data=nil then
 begin
   selectNode:=TreeView1.Selected;
   text:= selectNode.text;
   position:=selectnode.level;
    case position of
      0  :Showmsc0(text);
      1  : Showmsc(text);
      2  : ShowBsc(text);
      3  : ShowBase(text);
      4  : Showcell(text);
     end;
 end;
end;

procedure TForm1.showMSC0(var Name:string);
begin
  with query1 do
  begin
    close;
    sqL.clear;
    Sql.add('select * from MSC.dbf');
    execsql;
  end;
  table1.Close;
  gettemptbl;
  table1.Open;
  TempName:='MSC.dbf';
end;

procedure TForm1.showMSC(var Name:string);
begin
  with query1 do
  begin
    close;
    sqL.clear;
    if   (TreeView1.Selected.Expanded) and (TreeView1.Selected.haschildren) then
    begin
    // Sql.add('select * from BSC.dbf');
     Sql.add('select BSC_NAME,BSC_NO,LON,LAT from BSC.dbf');
     Sql.add('Where Msc_No="'+TreeView1.Selected.text+'"');
     TempName:='BSC.dbf';
    end
    else
    begin
     Sql.add('select * from MSC.dbf');
     Sql.add('Where Msc_No="'+TreeView1.Selected.text+'"');
     TempName:='MSC.dbf';
    end;
    execsql;
  end;
  // query1.open;
  table1.Close;
  gettemptbl;
  table1.Open;
end;


procedure TForm1.showBSC(var Name:string);
begin
   with query1 do
   begin
    close;
    sqL.clear;
    if   (TreeView1.Selected.Expanded) and (TreeView1.Selected.haschildren)  then
    begin
    // Sql.add('select * from Base.dbf');
     Sql.add('select BS_NAME, BS_NO,BCCH_1,BCCH_2,BCCH_3,CI_1,CI_2,CI_3,BSIC_1,');
     Sql.add('BSIC_2,BSIC_3,BEARING_1,BEARING_2,BEARING_3,LAC,BSC__SYSGE,');
     Sql.add('BASE_TYPE,BTS_TYPE,POWER_TYPE,LON,LAT  from Base.dbf');
     Sql.add('Where Bsc_No="'+TreeView1.Selected.text+'"');
     TempName:='BASE.dbf';
    end
    else
    begin
   // Sql.add('select * from Bsc.dbf');
    Sql.add('select BSC_NAME,BSC_NO,LON,LAT from BSC.dbf');
    Sql.add('Where Bsc_No="'+TreeView1.Selected.text+'"');
    TempName:='BSC.dbf';
    end;
    execsql;
   end;

  // query1.open;
      table1.Close;
  gettemptbl;
  table1.Open;


end;

procedure TForm1.showBase(var Name:string);
var bs :string;
begin 
   with query1 do
   begin
    close;
    sqL.clear;
    if   (TreeView1.Selected.Expanded) and (TreeView1.Selected.haschildren)  then
    begin
     bs:=copy(TreeView1.Selected.Text,1,pos('(',TreeView1.Selected.Text)-1);
     //Sql.add('select * from Cell.dbf');
     Sql.add('select CELL_NAME,BS_NO,CI,ARFCN,BSIC,BEARING,LAC,NON_BCCH,DOWNTILT,');
     Sql.add('MAX_TX_BTS,MAX_TX_MS,LON,LAT,MICROCELL,NCELL1,NCELL2,NCELL3,');
     Sql.add('NCELL4,NCELL5,NCELL6,NCELL7,NCELL8,NCELL9,NCELL10,NCELL11,NCELL12,');
     Sql.add('NCELL13,NCELL14,NCELL15,NCELL16 from Cell.dbf');
     Sql.add('Where Base_No="'+BS+'"');
     TempName:='CELL.dbf';
    end
    else
    begin
    bs:=copy(TreeView1.Selected.Text,1,pos('(',TreeView1.Selected.Text)-1);
    //Sql.add('select * from Base.dbf');
    Sql.add('select BS_NAME, BS_NO,BCCH_1,BCCH_2,BCCH_3,CI_1,CI_2,CI_3,BSIC_1,');
    Sql.add('BSIC_2,BSIC_3,BEARING_1,BEARING_2,BEARING_3,LAC,BSC__SYSGE,');
     Sql.add('BASE_TYPE,BTS_TYPE,POWER_TYPE,LON,LAT  from Base.dbf');
    Sql.add('Where Bs_No="'+bs+'"');
    TempName:='BASE.dbf';
    end;
    execsql;
   end;
  // query1.open;
    table1.Close;
  gettemptbl;
  table1.Open;
end;

procedure TForm1.showcell(var Name:string);
var bs :string;
begin
  
   with query1 do
   begin
    close;
    sqL.clear;
    //Sql.add('select * from cell.dbf');
      Sql.add('select CELL_NAME,BS_NO,CI,ARFCN,BSIC,BEARING,LAC,NON_BCCH,DOWNTILT,');
     Sql.add('MAX_TX_BTS,MAX_TX_MS,LON,LAT,MICROCELL,NCELL1,NCELL2,NCELL3,');
     Sql.add('NCELL4,NCELL5,NCELL6,NCELL7,NCELL8,NCELL9,NCELL10,NCELL11,NCELL12,');
     Sql.add('NCELL13,NCELL14,NCELL15,NCELL16 from Cell.dbf');
    Sql.add('Where Bs_No="'+TreeView1.Selected.Text+'"');
    execsql;
   end;
   TempName:='CELL.dbf';
   table1.Close;
   gettemptbl;
   table1.Open;

end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 try
  table1.close;
  except
  end;
end;

procedure TForm1.TreeView1Expanded(Sender: TObject; Node: TTreeNode);
var position:integer;
    text :string;
begin
  if  (TreeView1.Selected<>nil) and (Node.text=TreeView1.Selected.text) then
  begin
    position:=TreeView1.Selected.level;
    case position of
      0  : Showmsc0(text);
      1  : Showmsc(text);
      2  : ShowBsc(text);
      3  : ShowBase(text);
      4  : Showcell(text);
     end;
  end;
end;

procedure TForm1.TreeView1Collapsed(Sender: TObject; Node: TTreeNode);
var Position:integer;
    text :string;
begin
 if  (TreeView1.Selected<>nil) and (Node.text=TreeView1.Selected.text) then
  begin
    position:=TreeView1.Selected.level;
     case position of
      0  : Showmsc0(text);
      1  : Showmsc(text);
      2  : ShowBsc(text);
      3  : ShowBase(text);
      4  : Showcell(text);
     end;
  end;
  end;

procedure TForm1.Table1AfterPost(DataSet: TDataSet);
var mn:string;
    node:TTreenode ;
begin

   if state='Edit' then
   begin
     Query1.ExecSQL ;
     if TempName='MSC.dbf' then  mn:=table1.FieldbyName('MSC_NO').Asstring;
     if TempName='BSC.dbf' then  mn:=table1.FieldbyName('BSC_NO').Asstring;
     if TempName='BASE.dbf' then  mn:=table1.FieldbyName('BS_NO').Asstring+'('+table1.FieldbyName('BS_NAME').Asstring+')';
     if TempName='CELL.dbf' then  mn:=table1.FieldbyName('BS_NO').Asstring;
     if  mn=NNO then realtblchange
          else
          begin
           if (TreeView1.Selected.Expanded)and (TreeView1.Selected.HasChildren ) then  node:=TreeView1.Selected
                 else node:=TreeView1.Selected.parent;
           realtblchange;
           delenode:=TreeFindItem(NOde,NNO);
           if delenode=nil then
           begin
             exit;
             showmessage('delenode=nil');
           end;
           node:=delenode.parent;
           NNO:=mn;
           if  TempName ='MSC.dbf'then
           begin
             MscNode:=treeview1.items.addchild(node,NNO);
             MscNode.ImageIndex :=MscNode.level+1;
             MscNode.selectedindex:=0;
             AddBscNode(NNO);
           end;
           if  TempName ='BSC.dbf'then
           begin
             BscNode:=treeview1.items.addchild(node,NNO);
             BscNode.ImageIndex :=BscNode.level+1;
             BscNode.selectedindex:=0;
             AddBaseNode(NNO);
           end;

            if  TempName ='BASE.dbf'then
           begin
             BaseNode:=treeview1.items.addchild(node,NNO);
             BaseNode.ImageIndex :=BaseNode.level+1;
             BaseNode.selectedindex:=0;
             AddCellNode(NNO);
           end;

           if  TempName ='CELL.dbf'then
           begin
             CellNode:=treeview1.items.addchild(node,NNO);
             CellNode.ImageIndex :=CellNode.level+1;
             CellNode.selectedindex:=0;
           end;
           deleNode.Delete ;
           
          end;
   end;
   
   if state='Append' then
   begin
      realtblchange;
      realNodeChange;
   end;
   state:='cancel';
end;




procedure TForm1.RealNodechange;
var i,n:integer;
    tblName,XNo:string;
begin
  
   tblName:=table2.TableName;

   if  TempName ='MSC.dbf'then
    begin
     XNO:=table1.fieldbyName('MSC_NO').Asstring;
     MscNode:=treeview1.items.addchild(allnode,XNO);
     MscNode.ImageIndex :=MscNode.level+1;
     MscNode.selectedindex:=0;
     AddBscNode(XNO);
    end;

    if  TempName ='BSC.dbf'then
    begin
     XNO:=table1.fieldbyName('BSC_NO').Asstring;
     if  TreeView1.Selected.Expanded or not (TreeView1.Selected.HasChildren) then
               BscNode:=treeview1.items.addchild(TreeView1.Selected,XNO)
            else  BscNode:=treeview1.items.addchild(TreeView1.Selected.parent,XNO);
     BscNode.ImageIndex :=BscNode.level+1;
     BscNode.selectedindex:=0;
     AddBaseNode(XNO);
    end;

    if  TempName ='BASE.dbf'then
    begin
     XNO:=table1.fieldbyName('BS_NO').Asstring;
     if  TreeView1.Selected.Expanded or not (TreeView1.Selected.HasChildren) then
               BaseNode:=treeview1.items.addchild(TreeView1.Selected,XNO+'('+table1.fieldbyName('BS_NAME').Asstring+')')
            else  BaseNode:=treeview1.items.addchild(TreeView1.Selected.parent,XNO+'('+table1.fieldbyName('BS_NAME').Asstring+')');
     BaseNode.ImageIndex :=BaseNode.level+1;
     BaseNode.selectedindex:=0;
     AddCellNode(XNO);
    end;

    if  TempName ='CELL.dbf'then
    begin
     XNO:=table1.fieldbyName('BS_NO').Asstring;
     if  (TreeView1.Selected.Expanded) or not (TreeView1.Selected.HasChildren)   then
               CellNode:=treeview1.items.addchild(TreeView1.Selected,XNO)
            else  CellNode:=treeview1.items.addchild(TreeView1.Selected.parent,XNO);
     CellNode.ImageIndex :=CellNode.level+1;
     CellNode.selectedindex:=0;
    end;
end;


procedure TForm1.realtblchange;
var i,n:integer;
    fid:string;
    Node:TtreeNode;
begin

    table2.TableName:=TempName;
    table2.open;
    table2.append;
    for i:=0 to table1.FieldCount-1  do
    begin
     fid:=table1.fields[i].FieldName;
     table2.fieldbyname(fid).Asstring:=table1.fieldbyname(fid).Asstring;
    end;
    if TempName='BSC.dbf' then
    begin
       if TreeView1.Selected.level=2 then
            table2.fieldbyname('MSC_NO').Asstring:=TreeView1.Selected.parent.Text
           else table2.fieldbyname('MSC_NO').Asstring:=TreeView1.Selected.Text;
    end;

    if TempName='BASE.dbf' then
    begin
       if TreeView1.Selected.level=3 then
       table2.fieldbyname('BSC_NO').Asstring:=TreeView1.Selected.parent.Text
         else table2.fieldbyname('BSC_NO').Asstring:=TreeView1.Selected.Text;
    end;

    if TempName='CELL.dbf' then
    begin
       if TreeView1.Selected.level=4 then
         begin
            table2.fieldbyname('BASE_NO').Asstring:=copy(TreeView1.Selected.parent.Text,1,pos('(',TreeView1.Selected.parent.Text)-1);
            node:=TreeView1.Selected.parent;
         end
         else
         begin
           node:=TreeView1.Selected;
           table2.fieldbyname('BASE_NO').Asstring:=copy(TreeView1.Selected.Text,1,pos('(',TreeView1.Selected.Text)-1);
         end;
       table2.fieldbyname('BSC_NO').Asstring:=Node.parent.Text;
       table2.fieldbyname('TIME').Asstring:= Datetostr(Date);
    end;
    
    table2.Post ;
    table2.Refresh ;
    table2.close;
end;




procedure TForm1.DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
var n:integer;
    SQLL:string;
begin
    n:= TreeView1.Selected.level;
   if (Button in [nbInsert]) and (n<4)and (not TreeView1.Selected.HasChildren)   then
   begin
     case n  of
      0 : begin
          SQLL:='select * from Msc.dbf where MSC_NO is null';
          TempName:='MSC.dbf';
          end;
      1 : begin
          SQLL:='select BSC_NAME,BSC_NO,LON,LAT from Bsc.dbf where BSC_NO is null';
          TempName:='BSC.dbf';
          end;
      2 : begin
          SQLL:='select BS_NAME,BS_NO,BCCH_1,BCCH_2,BCCH_3,CI_1,CI_2,CI_3,BSIC_1,';
          SQLL:=SQLL+'BSIC_2,BSIC_3,BEARING_1,BEARING_2,BEARING_3,LAC,BSC__SYSGE,';
          SQLL:=SQLL+'BASE_TYPE,BTS_TYPE,POWER_TYPE,LON,LAT';
          SQLL:=SQLL+' from Base.dbf where BS_NO is null';
          TempName:='BASE.dbf';
          end;
      3 : begin
          SQLL:='select CELL_NAME,BS_NO,CI,ARFCN,BSIC,BEARING,LAC,NON_BCCH,';
          SQLL:=SQLL+'DOWNTILT,MAX_TX_BTS,MAX_TX_MS,LON,LAT,MICROCELL,NCELL1,';
          SQLL:=SQLL+'NCELL2,NCELL3,NCELL4,NCELL5,NCELL6,NCELL7,NCELL8,NCELL9,';
          SQLL:=SQLL+'NCELL10, NCELL11,NCELL12,NCELL13,NCELL14,NCELL15,NCELL16';
          SQLL:=SQLL+' from Cell.dbf where BS_NO is null';
          TempName:='CELL.dbf';
          end;
     end;
      with query1 do
      begin
        close;
        sql.Clear;
        sql.add(SQLL);
        execsql;
      end;
     table1.Close;
     gettemptbl;
     table1.Open;
     table1.Insert ;
   end;
end;



procedure TForm1.Table1AfterDelete(DataSet: TDataSet);
var text,fid:string;
      n:integer;
begin
     State:='';
      if dele=mrCancel then
      begin
       
       table1.Close;
       Table1.TableName := 'temp';
       BatchMove.Source := query1;
       BatchMove.Destination := table1;
       BatchMove.Mode := batAppend;
       BatchMove.Execute;
       table1.Open;
      end
      else
      begin
          deleNode.Delete ;
          if table1.recordcount=0 then
          begin
           n:= TreeView1.Selected.level;
           case n of
            0  : Showmsc0(text);
            1  : Showmsc(text);
            2  : ShowBsc(text);
            3  : ShowBase(text);
           end;
          end;
      end;

end;

procedure TForm1.Table1BeforeDelete(DataSet: TDataSet);
var ssq,mn,mess,text,sq,mm:string;
    node:TTreeNode;
    n:integer;
begin
    State:='Delete';
    if TempName='CELL.dbf' then
     begin
      mn:=table1.FieldbyName('BS_NO').Asstring;
      ssq:='select CELL_NAME,BS_NO,CI,ARFCN,BSIC,BEARING,LAC,NON_BCCH,';
      ssq:=ssq+'DOWNTILT,MAX_TX_BTS,MAX_TX_MS,LON,LAT,MICROCELL,NCELL1,';
      ssq:=ssq+'NCELL2,NCELL3,NCELL4,NCELL5,NCELL6,NCELL7,NCELL8,NCELL9,';
      ssq:=ssq+'NCELL10, NCELL11,NCELL12,NCELL13,NCELL14,NCELL15,NCELL16';
      sq:=' where BS_NO="'+mn+'"';
     end;

     if TempName='BASE.dbf' then
     begin
      mm:=table1.FieldbyName('BS_NO').Asstring;
      mn:= mm+'('+table1.FieldbyName('BS_NAME').Asstring+')';
      ssq:='select BS_NAME,BS_NO,BCCH_1,BCCH_2,BCCH_3,CI_1,CI_2,CI_3,BSIC_1,';
      ssq:=ssq+'BSIC_2,BSIC_3,BEARING_1,BEARING_2,BEARING_3,LAC,BSC__SYSGE,';
      ssq:=ssq+'BASE_TYPE,BTS_TYPE,POWER_TYPE,LON,LAT ';
      sq:=' where BS_NO="'+mm+'"';
     end;

     if TempName='BSC.dbf' then
     begin
      mn:=table1.FieldbyName('BSC_NO').Asstring;
      ssq:='select BSC_NAME,BSC_NO,LON,LAT' ;
      sq:=' where BSC_NO="'+mn+'"';
     end;

     if TempName='MSC.dbf' then
     begin
     mn:=table1.FieldbyName('MSC_NO').Asstring;
     ssq:='Select * ' ;
     sq:=' where MSC_NO="'+mn+'"';
     end;

     NNO:=mn;

     mess:='你的确要删除记录：'+node.Text ;
     dele:= MessageDlg(mess,mtConfirmation,[mbYes, mbCancel],0);

     if dele=mrCancel then
     begin
         with query1 do
         begin
              close;
              sql.clear;
              sql.Add( ssq+ ' from '+TempName+sq);
              execsql;
         end;
     end
     else
     begin

         if ((TreeView1.Selected.Expanded)
              and (TreeView1.Selected.HasChildren ))
               or (TreeView1.Selected.level=0 )   then  node:=TreeView1.Selected
                else node:=TreeView1.Selected.parent;
         delenode:=TreeFindItem(NOde,mn);
         deleNoderecord;
         with query1 do
         begin
            close;
            sql.clear;
            sql.Add('delete  from '+TempName+sq);
            execsql;
         end;
     end
end;

procedure TForm1.deleNoderecord;
var Bs:string;
    thisNode:TTreeNode;
begin
   if TempName='MSC.dbf' then
   begin
      if delenode.haschildren then
      begin
         thisNode:=delenode.getFirstChild ;

         while (thisNode<>nil)do
         begin
           with query1 do
           begin
             close;
            sql.clear;
            sql.Add('delete  from BASE.dbf where BSC_NO="'+thisNode.Text+'"');
            execsql;
            close;
            sql.clear;
            sql.Add('delete  from CELL.dbf where BSC_NO="'+thisNode.Text+'"');
            execsql;
           end;
           try
             thisNode:=thisNode.getNextSibling ;
           except
             thisNode:=nil;
           end;
         end;//end while ..do..

         with query1 do
         begin
          close;
          sql.clear;
          sql.Add('delete  from BSC.dbf where MSC_NO="'+delenode.Text+'"');
          execsql;
         end;
      end; //end ..if delenode.haschildren then
   end;

   if TempName='BSC.dbf' then
    begin
      with query1 do
      begin
        close;
        sql.clear;
        sql.Add('delete  from BASE.dbf where BSC_NO="'+delenode.Text+'"');
        execsql;
        close;
        sql.clear;
        sql.Add('delete  from CELL.dbf where BSC_NO="'+delenode.Text+'"');
        execsql;
      end;
   end;

   if TempName='BASE.dbf' then
   begin
      bs:=copy(delenode.Text,1,pos('(',delenode.Text)-1);
      with query1 do
      begin
        close;
        sql.clear;
        sql.Add('delete  from BASE.dbf where BS_NO="'+bs+'"');
        execsql;
      end;
   end;

end;


procedure TForm1.Table1BeforeEdit(DataSet: TDataSet);
Var mn,ssq:string;
begin
 State:='Edit';

    if TempName='CELL.dbf' then
    begin
     MN:=table1.FieldbyName('BS_NO').Asstring;
     ssq:='  where BS_NO="'+MN+'"';
    end;
    if TempName='BASE.dbf' then
    begin
     MN:=table1.FieldbyName('BS_NO').Asstring;
     
     ssq:='  where BS_NO="'+MN+'"';
    end;
    if TempName='BSC.dbf' then
    begin
     MN:=table1.FieldbyName('BSC_NO').Asstring;
     ssq:='  where BSC_NO="'+MN+'"'
   end;
    if TempName='MSC.dbf' then
    begin
     MN:=table1.FieldbyName('MSC_NO').Asstring;
     ssq:='  where MSC_NO="'+MN+'"'
    end;

    with Query1 do
    begin
       close;
       Sql.clear;
       sql.Add('delete  from '+TempName+ssq);
    end;
    NNO:=MN;
    if TempName='BASE.dbf' then NNO:=NNO+'('+table1.FieldbyName('BS_NAME').Asstring+')';
end;


procedure TForm1.Table1AfterCancel(DataSet: TDataSet);
begin
   state:= 'Cancel';
end;

procedure TForm1.Table1BeforeInsert(DataSet: TDataSet);
begin
  state:='Append';
end;




end.
