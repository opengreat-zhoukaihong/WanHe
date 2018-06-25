unit uqual_ta;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Grids, DBGrids, ExtCtrls, DBCtrls, TeEngine, Series,
  TeeProcs, Chart, DBChart, Buttons, StdCtrls;

type
  Tfuqual_ta = class(TForm)
    GroupBox1: TGroupBox;
    cb1: TComboBox;
    ComboBox2: TComboBox;
    Edit1: TEdit;
    rg1: TRadioGroup;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure FormActivate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure cbx;
    { Public declarations }
  end;

var
  fuqual_ta: Tfuqual_ta;

implementation
uses ALLCDD,history, DataList;
{$R *.DFM}

procedure tfuqual_ta.cbx;
var
  i:integer;
begin
  cb1.Items.Clear;
  with fmDataList.query1 do
  begin
    for i:=0 to FieldCount-1 do
    begin
      Cb1.Items.Add(Fields.Fields[i].FieldName);
    end;
  end;
  Cb1.Text:=fmDataList.dbgrid1.selectedfield.FieldName;
end;

procedure Tfuqual_ta.FormActivate(Sender: TObject);
begin
  cbx;
end;

procedure Tfuqual_ta.BitBtn2Click(Sender: TObject);
var
  str,by:string;
  dbcol:integer;
begin
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  str:=Cb1.Text+ComboBox2.Text+edit1.Text;
  by:=Cb1.Text;
         with fmDataList.query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('select  '+wSql+',DATA_CHANGE');
           sql.add('from   '+wTableName);
           sql.add('where RE_DATE=:P AND '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
  fmDataList.dbgrid2.SelectedIndex:=dbcol;
  fmDataList.dbgrid2.SetFocus;
  fmDataList.DBGRID2.Columns.Items[fmDataList.DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  except
  end;
end;

procedure Tfuqual_ta.SpeedButton1Click(Sender: TObject);
var
  str,by:string;
  dbcol:integer;
  selfield:tfield;
begin
  if wtablename<>'RLSMP' THEN
BEGIN
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  SelField:=fmDataList.query1.Fields.FindField(cb1.Text);
  if SelField.DataType=ftString then
    str:=Cb1.Text+ComboBox2.Text+''''+edit1.Text+''''
  else
    str:=Cb1.Text+ComboBox2.Text+edit1.Text;
    by:=Cb1.Text;
         with fmDataList.query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('select  '+wSql+',DATA_CHANGE');
           sql.add('from   '+wTableName);
           sql.add('where RE_DATE=:P AND '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
  fmDataList.dbgrid2.SelectedIndex:=dbcol;
  fmDataList.dbgrid2.SetFocus;
  fmDataList.DBGRID2.Columns.Items[fmDataList.DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  except
  end;
end else
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  SelField:=fmDataList.query1.Fields.FindField(cb1.Text);
  if SelField.DataType=ftString then
    str:=Cb1.Text+ComboBox2.Text+''''+edit1.Text+''''
  else
    str:=Cb1.Text+ComboBox2.Text+edit1.Text;
    by:=Cb1.Text;
         with fmDataList.query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('select * ');
           sql.add('from  temp_rlsmp  ');
           sql.add('where  '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
           open;
         end;
  fmDataList.dbgrid2.SelectedIndex:=dbcol;
  fmDataList.dbgrid2.SetFocus;
  fmDataList.DBGRID2.Columns.Items[fmDataList.DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  except
  end;
end;


procedure Tfuqual_ta.SpeedButton2Click(Sender: TObject);
begin
  close;
end;

end.
