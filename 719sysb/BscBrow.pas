unit BscBrow;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Grids, DBGrids, ComCtrls, Menus;

type
  TfmBscBrow = class(TForm)
    pcBscBrow: TPageControl;
    tsTch: TTabSheet;
    dbgTchBrowse: TDBGrid;
    tsSdcch: TTabSheet;
    dbgCchBrowse: TDBGrid;
    dsBscTchBrow: TDataSource;
    quBscTchBrow: TQuery;
    dsBscCchBrow: TDataSource;
    quBscCchBrow: TQuery;
    quBscTchBrowBSC_NO: TStringField;
    quBscTchBrowTCH: TStringField;
    quBscTchBrowU: TFloatField;
    quBscTchBrowCG: TFloatField;
    quBscTchBrowDR: TFloatField;
    quBscTchBrowF: TFloatField;
    quBscTchBrowTRAFFIC: TFloatField;
    quBscCchBrowBSC_NO: TStringField;
    quBscCchBrowCCH: TStringField;
    quBscCchBrowSU: TFloatField;
    quBscCchBrowSC: TFloatField;
    quBscCchBrowSDR: TFloatField;
    quBscCchBrowSF: TFloatField;
    quBscCchBrowTRAFFIC: TFloatField;
    pmBrowse: TPopupMenu;
    mmAscend: TMenuItem;
    mmDesc: TMenuItem;
    procedure pcBscBrowMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbgTchBrowseDblClick(Sender: TObject);
    procedure dbgCchBrowseDblClick(Sender: TObject);
    procedure dbgTchBrowseTitleClick(Column: TColumn);
    procedure dbgCchBrowseTitleClick(Column: TColumn);
    procedure mmAscendClick(Sender: TObject);
    procedure mmDescClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmBscBrow: TfmBscBrow;

implementation
uses BscMain;
{$R *.DFM}

procedure TfmBscBrow.pcBscBrowMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  wBsc : String;
begin
  if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'BSC' then
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(1) +' from selection');
    wBSC := oleMapInfo.eval('selection.BSC_NO');
    if fmBscBrow.quBScTchBrow.Active then
    begin
      fmBscBrow.quBScTchBrow.First;
      while (not fmBscBrow.quBScTchBrow.eof)  do
      begin
        if (wBsc = fmBscBrow.quBscTchBrow.FieldByName('bsc_no').AsString) then
          Break;
        fmBscBrow.quBscTchBrow.Next;
      end;
    end;
    if fmBscBrow.quBscCchBrow.Active then
    begin
      fmBscBrow.quBScCchBrow.First;
      while not fmBscBrow.quBScCchBrow.eof  do
      begin
        if (wBsc = fmBscBrow.quBscCchBrow.FieldByName('bsc_no').AsString) then
          Break;
        fmBscBrow.quBscCchBrow.Next;
      end;
    end;
  end;
end;

procedure TfmBscBrow.dbgTchBrowseDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from bsc where bsc_no = "' +
    quBscTchBrow.FieldByName('bsc_no').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');
end;

procedure TfmBscBrow.dbgCchBrowseDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from bsc where bsc_no = "' +
    quBscCchBrow.FieldByName('bsc_no').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');
end;

procedure TfmBscBrow.dbgTchBrowseTitleClick(Column: TColumn);
var
  wSql : String;
begin
  wSql := Trim(Column.FieldName);
  with quBscTchBrow do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select  * from bsc_tch_file');
    if mmAscend.Checked then
      Sql.Add('order by bsc_tch_file.' + wSql)
    else
      Sql.Add('order by bsc_tch_file.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;

end;

procedure TfmBscBrow.dbgCchBrowseTitleClick(Column: TColumn);
var
  wSql : String;
begin
  wSql := Trim(Column.FieldName);
  with quBscCchBrow do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select  * from bsc_cch_file');
    if mmAscend.Checked then
      Sql.Add('order by bsc_cch_file.' + wSql)
    else
      Sql.Add('order by bsc_cch_file.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;
end;

procedure TfmBscBrow.mmAscendClick(Sender: TObject);
begin
  mmAscend.Checked := True;
  mmDesc.Checked := False;
end;

procedure TfmBscBrow.mmDescClick(Sender: TObject);
begin
  mmAscend.Checked := False;
  mmDesc.Checked := True;
end;

procedure TfmBscBrow.FormCreate(Sender: TObject);
begin
  MoveWindow(Handle, Screen.Width - Width, Screen.Height - 150, Width, Height, True);
end;

end.
