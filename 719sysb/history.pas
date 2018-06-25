unit history;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, Grids, DBGrids, DBTables, ExtCtrls, Menus, StdCtrls, Buttons;

type
  Tfhistory = class(TForm)
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    Query1: TQuery;
    Panel1: TPanel;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    CB5: TComboBox;
    SpeedButton1: TSpeedButton;
    procedure N3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
  procedure cbx;
    { Public declarations }
  end;

var
  fhistory: Tfhistory;

implementation

uses  DataList,AllCdd, uqual_ta;

{$R *.DFM}
procedure Tfhistory.cbx;
var
  i:integer;
begin
  cb5.Items.Clear;
  with query1 do
  begin
    for i:=0 to FieldCount-1 do
    begin
      Cb5.Items.Add(Fields.Fields[i].FieldName);
    end;
  end;
  Cb5.Text:=fhistory.dbgrid1.selectedfield.FieldName;
end;

procedure Tfhistory.N3Click(Sender: TObject);
begin
fuqual_ta.showmodal;
end;

procedure Tfhistory.N1Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
//  showmessage(by);
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select '+wSql);
           SQL.Add('From  '+wTableName +' WHERE RE_DATE=:P ');
           sql.add('order by '+by);
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
end;

procedure Tfhistory.N2Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select '+wSql);
           SQL.Add('From  '+wTableName+'  WHERE RE_DATE=:P');
           sql.add('order by '+by+' desc');
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
end;


procedure Tfhistory.SpeedButton1Click(Sender: TObject);
begin
close;
end;

end.
