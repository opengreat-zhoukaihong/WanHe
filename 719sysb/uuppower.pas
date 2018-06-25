unit uuppower;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, DBTables, DB;

type
  Tfuuppower = class(TForm)
    GroupBox1: TGroupBox;
    cb1: TComboBox;
    ComboBox2: TComboBox;
    Edit1: TEdit;
    rg1: TRadioGroup;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure cbx;
    { Public declarations }
  end;

var
  fuuppower: Tfuuppower;

implementation

uses DATALIST,ALLCDD;

{$R *.DFM}

procedure tfuuppower.cbx;
var
  i:integer;
begin
  cb1.Items.Clear;
  with fmDataList.quDataList do
  begin
    for i:=0 to FieldCount-1 do
    begin
      Cb1.Items.Add(Fields.Fields[i].FieldName);
    end;
  end;
  Cb1.Text:=fmDataList.dbgrid1.selectedfield.FieldName;
end;

procedure Tfuuppower.BitBtn1Click(Sender: TObject);
var
  str,by:string;
  dbcol:integer;
begin
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  str:=Cb1.Text+ComboBox2.Text+edit1.Text;
  by:=Cb1.Text;
         with fmDataList.quDataList do
         begin
           Close;

           SQL.Clear;
           SQL.Add('select  '+wSql);
           sql.add('from   '+wTableName);
           sql.add('where RE_DATE=:P AND  '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
           PARAMBYNAME('P').ASSTRING:='2000';
           open;
         end;
  fmDataList.dbgrid1.SelectedIndex:=dbcol;
  fmDataList.dbgrid1.SetFocus;
  except
  end;
end;

procedure Tfuuppower.FormActivate(Sender: TObject);
begin
  cbx;
end;

procedure Tfuuppower.SpeedButton1Click(Sender: TObject);
begin
CLOSE;
end;

procedure Tfuuppower.SpeedButton2Click(Sender: TObject);
var
  str,by:string;
  dbcol:integer;
  selfield:tfield;
begin
  if wtablename<>'RLSMP' THEN
BEGIN
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  SelField:=fmDataList.quDataList.Fields.FindField(cb1.Text);
  if SelField.DataType=ftString then
    str:=Cb1.Text+ComboBox2.Text+''''+edit1.Text+''''
  else
    str:=Cb1.Text+ComboBox2.Text+edit1.Text;
    by:=Cb1.Text;
         with fmDataList.quDataList do
         begin
           Close;
           SQL.Clear;
           SQL.Add('select  '+wSql);
           sql.add('from   '+wTableName);
           sql.add('where RE_DATE=:P AND  '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
           PARAMBYNAME('P').ASSTRING:='2000';
           open;
         end;
  fmDataList.dbgrid1.SelectedIndex:=dbcol;
  fmDataList.dbgrid1.SetFocus;
  except
  end;
end else
begin
  try
  dbcol:=fmDataList.dbgrid1.SelectedIndex;
  SelField:=fmDataList.quDataList.Fields.FindField(cb1.Text);
  if SelField.DataType=ftString then
    str:=Cb1.Text+ComboBox2.Text+''''+edit1.Text+''''
  else
    str:=Cb1.Text+ComboBox2.Text+edit1.Text;
    by:=Cb1.Text;
         with fmDataList.quDataList do
         begin
           Close;
           SQL.Clear;
           SQL.Add('select  * ');
           sql.add('from   temp_rlsmp');
           sql.add('where '+ str);
           if rg1.ItemIndex=0 then
           sql.add(' order by '+ by)
           else
           sql.add(' order by '+ by+ ' desc');
//           PARAMBYNAME('P').ASSTRING:='2000';
           open;
         end;
  fmDataList.dbgrid1.SelectedIndex:=dbcol;
  fmDataList.dbgrid1.SetFocus;
  except
  end;
end;

end;

end.


