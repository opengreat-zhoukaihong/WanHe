unit Trendline;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, ExtCtrls, TeeProcs, TeEngine, Chart, DBChart, TeeFunci,
  Series, StdCtrls, Buttons, ComCtrls, Grids, DBGrids;

type
  TfmTrendline = class(TForm)
    quTrendline: TQuery;
    dsTrendline: TDataSource;
    Panel1: TPanel;
    pnFoot: TPanel;
    Label7: TLabel;
    cbType: TComboBox;
    Label8: TLabel;
    cbParameter: TComboBox;
    dtpStartDate: TDateTimePicker;
    dtpEndDate: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    rgType: TRadioGroup;
    Panel2: TPanel;
    dchTrendline: TDBChart;
    cbTchParm: TComboBox;
    cbCchParm: TComboBox;
    srTrendline: TLineSeries;
    TeeFunction1: TSubtractTeeFunction;
    dgTrendLine: TDBGrid;
    Series1: TLineSeries;
    Series2: TLineSeries;
    Series3: TLineSeries;
    Series4: TLineSeries;
    Series5: TLineSeries;
    DataSource1: TDataSource;
    Query1: TQuery;
    DataSource2: TDataSource;
    Query2: TQuery;
    DataSource3: TDataSource;
    Query3: TQuery;
    DataSource4: TDataSource;
    Query4: TQuery;
    DataSource5: TDataSource;
    Query5: TQuery;
    procedure cbTypeChange(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure pnFootClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmTrendline: TfmTrendline;

implementation
uses BscMain, bscData;
{$R *.DFM}

procedure TfmTrendline.cbTypeChange(Sender: TObject);
begin
  if cbType.Text = 'TCH' then
  begin
    cbParameter.Items := cbTchParm.Items;
    cbParameter.Text := cbTchParm.Text;
  end
  else
  begin
    cbParameter.Items := cbCchParm.Items;
    cbParameter.Text := cbCchParm.Text;
  end;

end;

procedure TfmTrendline.BitBtn1Click(Sender: TObject);
var
  wType, wParm, wStartDate, wEndDate, wCellId  : String;
begin

  wType := cbType.Text;
  wParm := cbParameter.Text;
  Delete(wParm, 1, Pos(' ',wParm));
  wParm := UpperCase(Trim(wParm));
  wStartDate := FormatDateTime('yyyymmdd', dtpStartDate.DateTime);
  wEndDate := FormatDateTime('yyyymmdd', dtpEndDate.DateTime);
 // if srTrendline.Active then
 //   srTrendline.Active := False;
  srTrendline.YValues.ValueSource := '';
  if rgType.ItemIndex = 0 then
  begin
    dgTrendLine.Columns[0].FieldName := 'START_DATE';
    dgTrendLine.Columns[0].Title.Caption := '开始日期';
    srTrendline.XValues.ValueSource := 'START_DATE';
    srTrendline.XLabelsSource := 'START_DATE';
  end
  else
  begin
    dgTrendLine.Columns[0].FieldName := 'START_TIME';
    dgTrendLine.Columns[0].Title.Caption := '开始时间';
    srTrendline.XValues.ValueSource := 'START_TIME';
    srTrendline.XLabelsSource := 'START_TIME';
  end;

  if cbType.Text = 'SDCCH' then
  begin
    with quTrendline do
    begin
      if Active then
        Close;

      if gSelFlag = 'BSC' then
      begin
        with sql do
        begin
          Clear;
          Add(' select * from bsc_all_cch_file where start_date >= :Start_date ');
          Add(' and start_date <= :end_date and bsc_no = :bsc_no order by start_date');
        end;
        ParamByName('Start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('End_date').AsInteger := StrToInt(wEndDate);
        ParamByName('bsc_no').AsString := gSelName;
        srTrendline.YValues.ValueSource := wParm;
        Open;

      end
      else
      begin
        oleMapInfo.do('fetch rec 1 from selection');
        wCellId := oleMapInfo.eval('selection.bs_no');
        with sql do
        begin
          Clear;
          Add(' select * from all_cch_file where start_date >= :Start_date ');
          Add(' and start_date <= :end_date and cell_id = :cell_id order by start_date');
        end;
        ParamByName('Start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('End_date').AsInteger := StrToInt(wEndDate);
        ParamByName('Cell_id').AsString := wCellId;
        srTrendline.YValues.ValueSource := wParm;
        Open;

      end;
    end;
  end
  else
  begin
    //srTrendline.Title := cbParameter.Text;
    with quTrendline do
    begin
      if Active then
        Close;
      if gSelFlag = 'BSC' then
      begin
        with sql do
        begin
          Clear;
          Add(' select * from bsc_all_Tch_file where start_date >= :Start_date ');
          Add(' and start_date <= :end_date and bsc_no = :bsc_no order by start_date');
        end;
        ParamByName('Start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('End_date').AsInteger := StrToInt(wEndDate);
        ParamByName('bsc_no').AsString := gSelName;
        srTrendline.YValues.ValueSource := wParm;
        Open;

      end
      else
      begin
        oleMapInfo.do('fetch rec 1 from selection');
        wCellId := oleMapInfo.eval('selection.bs_no');
        with sql do
        begin
          Clear;
          Add(' select * from all_tch_file where start_date >= :Start_date ');
          Add(' and start_date <= :end_date and cell_id = :cell_id order by start_date');
        end;
        ParamByName('Start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('End_date').AsInteger := StrToInt(wEndDate);
        ParamByName('Cell_id').AsString := wCellId;
        srTrendline.YValues.ValueSource := wParm;
        Open;

      end;
    end;
  end;
  dgTrendLine.Columns[1].FieldName := wParm;
  dgTrendLine.Columns[1].Title.Caption :=
    Copy(cbParameter.Text, 1, Pos(' ', cbParameter.Text) - 1);
  if gSelFlag = 'BSC' then
     pnFoot.Caption :=  gSelName
  else
    pnFoot.Caption := oleMapInfo.eval('selection.cell_name');
 // srTrendline.Active := True;
  {
  With MySeries do
begin
 ParentChart:=DBChart1;
 DataSource:=Table1;
 XLabelsSource:='Name';
 YValues.ValueSource:= 'Amount';
 CheckDatasource;
end;}

end;

procedure TfmTrendline.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  with quTrendline do
    if Active then
      Close;
  
end;

procedure TfmTrendline.FormCreate(Sender: TObject);
begin
  if gSelFlag = 'BSC' then
    Caption := 'BSC' + Caption
  else
    Caption := '小区' + Caption;
end;

procedure TfmTrendline.pnFootClick(Sender: TObject);
begin
  quTrendline.Open;
  Query1.Open;
  Query2.Open;
  Query3.Open;
  Query4.Open;
  Query5.Open;
end;

end.
