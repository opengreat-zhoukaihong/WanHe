unit Browse;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, Db, DBTables, ComCtrls, Buttons, ToolWin, ExtCtrls,
  DBCtrls, StdCtrls, Wwdbigrd, Wwdbgrid, Wwdatsrc, Wwquery, Menus;

type
  TfmBrowse = class(TForm)
    pcBrowse: TPageControl;
    tsTch: TTabSheet;
    tsSdcch: TTabSheet;
    tsHover: TTabSheet;
    GroupBox1: TGroupBox;
    Splitter1: TSplitter;
    GroupBox2: TGroupBox;
    quAllSelCell: TQuery;
    quTchBrowse: TwwQuery;
    dsTchBrowse: TwwDataSource;
    dbgTchBrowse: TwwDBGrid;
    quTchBrowseCELL_ID: TStringField;
    quTchBrowseBSC_NO: TStringField;
    quTchBrowseTCH: TStringField;
    quTchBrowseCA: TFloatField;
    quTchBrowseCS: TFloatField;
    quTchBrowseU: TFloatField;
    quTchBrowseERPAC: TFloatField;
    quTchBrowseCG: TFloatField;
    quTchBrowseMH: TFloatField;
    quTchBrowseDR: TFloatField;
    quTchBrowseCH: TFloatField;
    quTchBrowseAC: TFloatField;
    quTchBrowseF: TFloatField;
    quTchBrowseTQA: TFloatField;
    quTchBrowseTSS4: TFloatField;
    quTchBrowseTHSI: TFloatField;
    quTchBrowseTHSE: TFloatField;
    quTchBrowseER_DR: TFloatField;
    quTchBrowseTG: TFloatField;
    quTchBrowseTRAFFIC: TFloatField;
    quTchBrowseSTANDARD: TFloatField;
    dsCchBrowse: TwwDataSource;
    quCchBrowse: TwwQuery;
    quCchBrowseCELL_ID: TStringField;
    quCchBrowseBSC_NO: TStringField;
    quCchBrowseCCH: TStringField;
    quCchBrowseRAF: TFloatField;
    quCchBrowseRAA: TFloatField;
    quCchBrowseRAS: TFloatField;
    quCchBrowseRAC: TFloatField;
    quCchBrowseSA: TFloatField;
    quCchBrowseSS: TFloatField;
    quCchBrowseSU: TFloatField;
    quCchBrowseSC: TFloatField;
    quCchBrowseSDR: TFloatField;
    quCchBrowseCH: TFloatField;
    quCchBrowseAC: TFloatField;
    quCchBrowseSF: TFloatField;
    quCchBrowseDQA: TFloatField;
    quCchBrowseDSS4: TFloatField;
    quCchBrowseTRAFFIC: TFloatField;
    dgCch: TwwDBGrid;
    dgHdove: TwwDBGrid;
    quHdove: TwwQuery;
    dsHdove: TwwDataSource;
    quHdoveCELL_ID_SE: TStringField;
    quHdoveCELL_ID_TA: TStringField;
    quHdoveHORTTOCH: TFloatField;
    quHdoveHOVERCNT: TFloatField;
    quHdoveHOVERSUC: TFloatField;
    quHdoveHOLOST: TFloatField;
    quHdoveFLUNK_RATE: TFloatField;
    dgHdovi: TwwDBGrid;
    quHdovi: TwwQuery;
    dsHdovi: TwwDataSource;
    quSumHdovi: TwwQuery;
    quSumHdove: TwwQuery;
    quSumHdoviSUM_HCNT: TFloatField;
    quSumHdoviSUM_HSUC: TFloatField;
    quSumHdoviSUM_HORT: TFloatField;
    quSumHdoveSUM_HCNT: TFloatField;
    quSumHdoveSUM_HSUC: TFloatField;
    quSumHdoveSUM_HORT: TFloatField;
    quAllCell: TQuery;
    quAllCellCELL_ID: TStringField;
    pmBrowse: TPopupMenu;
    mmAscend: TMenuItem;
    mmDesc: TMenuItem;
    quSumHdoviIn: TwwQuery;
    quSumHdoveIn: TwwQuery;
    quSumHdoviInSUM_HCNT: TFloatField;
    quSumHdoviInSUM_HSUC: TFloatField;
    quSumHdoviInSUM_HORT: TFloatField;
    quSumHdoveInSUM_HCNT: TFloatField;
    quSumHdoveInSUM_HSUC: TFloatField;
    quSumHdoveInSUM_HORT: TFloatField;
    N1: TMenuItem;
    N2: TMenuItem;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dbgTchBrowse1TitleClick(Column: TColumn);
    procedure pcBrowseMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure dbgTchBrowseDblClick(Sender: TObject);
    procedure dbgTchBrowseTitleButtonClick(Sender: TObject;
      AFieldName: String);
    procedure dgCchDblClick(Sender: TObject);
    procedure dgCchTitleButtonClick(Sender: TObject;
      AFieldName: String);
    procedure wwDBGrid1DblClick(Sender: TObject);
    procedure dgHdoveTitleButtonClick(Sender: TObject; AFieldName: String);
    procedure dgHdoviTitleButtonClick(Sender: TObject; AFieldName: String);
    procedure dgHdoviUpdateFooter(Sender: TObject);
    procedure dgHdoveUpdateFooter(Sender: TObject);
    procedure dgHdoviDblClick(Sender: TObject);
    procedure dgHdoveDblClick(Sender: TObject);
    procedure mmAscendClick(Sender: TObject);
    procedure mmDescClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmBrowse: TfmBrowse;

implementation

uses BscMain, BscData;

{$R *.DFM}

procedure TfmBrowse.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  with quTchBrowse do
    if Active then
      Close;
  with quCchBrowse do
    if Active then
      Close;
  with quHdovi do
    if Active then
      Close;
  with quHdove do
    if Active then
      Close;
  Action := caFree;
  fmBscMain.wBrowseShow := False;
end;

procedure TfmBrowse.dbgTchBrowse1TitleClick(Column: TColumn);
var
  wSql : String;
begin
  wSql := Trim(Column.FieldName);
  with quTchBrowse do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select  all_sel_cell.*, tch_file.* from all_sel_cell, tch_file');
    Sql.Add(' where all_sel_cell.cell_id = tch_file.cell_id ');
    Sql.Add('order by tch_file.' + wSql);
    //ShowMessage(sql.text);
    Open;
  end;

end;

procedure TfmBrowse.pcBrowseMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  wCell : String;
begin

  if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'CELL' then
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(1) +' from selection');
    wCell := oleMapInfo.eval('selection.bs_no');
    if quTchBrowse.Active then
      quTchBrowse.Locate('cell_id',UpperCase(wCell),[loPartialKey]);
    if quCchBrowse.Active then
      quCchBrowse.Locate('cell_id',UpperCase(wCell),[loPartialKey]);
  end;
end;

procedure TfmBrowse.FormCreate(Sender: TObject);
begin
  if gHover then
  begin
    pcBrowse.ActivePage := tsHover;
    tsTch.Visible := False;
    tsSdcch.Visible := False;
  end
  else
  begin
    pcBrowse.ActivePage := tsTch;
    tsHover.Visible := False;
  end;
  MoveWindow(Handle, Screen.Width - Width, Screen.Height - 150, Width, Height, True);
end;

procedure TfmBrowse.dbgTchBrowseDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from cell where bs_no = "' +
    quTchBrowse.FieldByName('cell_id').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.1 Units "km"');
end;

procedure TfmBrowse.dbgTchBrowseTitleButtonClick(Sender: TObject;
  AFieldName: String);
var
  wSql : String;
begin
  wSql := AFieldName;
  //ShowMessage(wSql);
  with quTchBrowse do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select  * from tch_file');
    //Sql.Add(' where all_sel_cell.cell_id = tch_file.cell_id ');
    if mmAscend.Checked then
      Sql.Add('order by tch_file.' + wSql )
    else
      Sql.Add('order by tch_file.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;

end;


procedure TfmBrowse.dgCchDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from cell where bs_no = "' +
    quCchBrowse.FieldByName('cell_id').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.1 Units "km"');
end;



procedure TfmBrowse.dgCchTitleButtonClick(Sender: TObject;
  AFieldName: String);
var
  wSql : String;
begin
  wSql := Trim(AFieldName);
  with quCchBrowse do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select  * from  Cch_file');
    //Sql.Add(' where all_sel_cell.cell_id = Cch_file.cell_id ');
    if mmAscend.Checked then
      Sql.Add('order by Cch_file.' + wSql)
    else
      Sql.Add('order by Cch_file.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;
end;

procedure TfmBrowse.wwDBGrid1DblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from cell where bs_no = "' +
    quCchBrowse.FieldByName('cell_id').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.1 Units "km"');
end;

procedure TfmBrowse.dgHdoveTitleButtonClick(Sender: TObject;
  AFieldName: String);
var
  wSql : String;
begin
  wSql := AFieldName;
  //ShowMessage(wSql);
  with quHdove do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select * from hdov_e_bsc ');
    //Sql.Add(' where all_sel_cell.cell_id = tch_file.cell_id ');
    if mmAscend.Checked then
      Sql.Add('order by hdov_e_bsc.' + wSql)
    else
      Sql.Add('order by hdov_e_bsc.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;
end;
procedure TfmBrowse.dgHdoviTitleButtonClick(Sender: TObject;
  AFieldName: String);
var
  wSql : String;
begin
  wSql := AFieldName;
  //ShowMessage(wSql);
  with quHdovi do
  begin
    if Active then
      Close;
    Sql.Clear;
    Sql.Add('select * from hdov_i_bsc ');
    //Sql.Add(' where all_sel_cell.cell_id = tch_file.cell_id ');
    if mmAscend.Checked then
      Sql.Add('order by hdov_i_bsc.' + wSql)
    else
      Sql.Add('order by hdov_i_bsc.' + wSql + ' desc');
    //ShowMessage(sql.text);
    Open;
  end;
end;

procedure TfmBrowse.dgHdoviUpdateFooter(Sender: TObject);
var
  wInt : Integer;
  wCell : String;
begin
  if gSelFlag  = 'CELL' then
  begin

    if gMultiCell.Count > 0 then
    begin
      wCell := gMultiCell.Strings[0];
      wInt := Pos(' ', wCell) - 1;
      wCell := Copy(wCell, 1, wInt);
      with quSumHdovi do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
        Open;
      end;
      with quSumHdoviIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
        Open;
      end;

      dgHdovi.ColumnByName('HOVERCNT').FooterValue :=
        quSumHdovi.FieldByName('sum_hcnt').AsString + '/' +
        quSumHdoviIn.FieldByName('sum_hcnt').AsString;
      dgHdovi.ColumnByName('HOVERSUC').FooterValue :=
        quSumHdovi.FieldByName('sum_hsuc').AsString + '/' +
        quSumHdoviIn.FieldByName('sum_hsuc').AsString;
      dgHdovi.ColumnByName('horttoch').FooterValue :=
        quSumHdovi.FieldByName('sum_hort').AsString + '/' +
        quSumHdoviIN.FieldByName('sum_hort').AsString;
      dgHdovi.ColumnByName('cell_id_se').FooterValue := wCell;
      dgHdovi.ColumnByName('cell_id_ta').FooterValue := '总计:';
      quSumHdovi.Close;
      quSumHdoviIn.Close;
    end
    else
    begin
      with quSumHdovi do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;
      with quSumHdove do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;
      with quSumHdoviIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;
      with quSumHdoveIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;
      dgHdovi.ColumnByName('HOVERCNT').FooterValue :=
        quSumHdovi.FieldByName('sum_hcnt').AsString + '/' +
        quSumHdoviIn.FieldByName('sum_hcnt').AsString;
      dgHdovi.ColumnByName('HOVERSUC').FooterValue :=
        quSumHdovi.FieldByName('sum_hsuc').AsString + '/' +
        quSumHdoviIn.FieldByName('sum_hsuc').AsString;
      dgHdovi.ColumnByName('horttoch').FooterValue :=
        quSumHdovi.FieldByName('sum_hort').AsString + '/' +
        quSumHdoviIN.FieldByName('sum_hort').AsString;
      dgHdovi.ColumnByName('cell_id_se').FooterValue := gSelName;
      dgHdovi.ColumnByName('cell_id_ta').FooterValue := '总计:';
      quSumHdovi.Close;
      quSumHdoviIn.Close;
    end;
  end;
end;

procedure TfmBrowse.dgHdoveUpdateFooter(Sender: TObject);
var
  wInt : Integer;
  wCell : String;
begin
  if gSelFlag = 'CELL' then
  begin
    if gMultiCell.Count > 0 then
    begin
      wCell := gMultiCell.Strings[0];
      wInt := Pos(' ', wCell) - 1;
      wCell := Copy(wCell, 1, wInt);
      with quSumHdove do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
        Open;
      end;
     {with quSumHdoveIN do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gMultiCell.Strings[0];
        Open;
      end; }
      with quSumHdoveIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
        Open;
      end;
     { with quSumHdoveIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gMultiCell.Strings[0];
        Open;
      end; }
      dgHdove.ColumnByName('HOVERCNT').FooterValue :=
        quSumHdove.FieldByName('sum_hcnt').AsString + '/' +
        quSumHdoveIn.FieldByName('sum_hcnt').AsString;
      dgHdove.ColumnByName('HOVERSUC').FooterValue :=
        quSumHdove.FieldByName('sum_hsuc').AsString + '/' +
        quSumHdoveIn.FieldByName('sum_hsuc').AsString;
      dgHdove.ColumnByName('horttoch').FooterValue :=
        quSumHdove.FieldByName('sum_hort').AsString + '/' +
        quSumHdoveIN.FieldByName('sum_hort').AsString;
      dgHdove.ColumnByName('cell_id_se').FooterValue := wCell;
      dgHdove.ColumnByName('cell_id_ta').FooterValue := '总计:' ;
      quSumHdove.Close;
      quSumHdoveIn.Close;
    end
    else
    begin
      with quSumHdove do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;


      with quSumHdoveIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := gSelName;
        Open;
      end;
      dgHdove.ColumnByName('HOVERCNT').FooterValue :=
        quSumHdove.FieldByName('sum_hcnt').AsString + '/' +
        quSumHdoveIn.FieldByName('sum_hcnt').AsString;
      dgHdove.ColumnByName('HOVERSUC').FooterValue :=
        quSumHdove.FieldByName('sum_hsuc').AsString + '/' +
        quSumHdoveIn.FieldByName('sum_hsuc').AsString;
      dgHdove.ColumnByName('horttoch').FooterValue :=
        quSumHdove.FieldByName('sum_hort').AsString + '/' +
        quSumHdoveIN.FieldByName('sum_hort').AsString;
      dgHdove.ColumnByName('cell_id_se').FooterValue := gSelName;
      dgHdove.ColumnByName('cell_id_ta').FooterValue := '总计:';
      quSumHdove.Close;
      quSumHdoveIn.Close;
    end;
  end;
end;

procedure TfmBrowse.dgHdoviDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from cell where bs_no = "' +
    quHdovi.FieldByName('cell_id_se').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.1 Units "km"');
end;

procedure TfmBrowse.dgHdoveDblClick(Sender: TObject);
begin
  oleMapInfo.do('select * from cell where bs_no = "' +
    quHdove.FieldByName('cell_id_se').AsString + '" into tmp');

  oleMapInfo.do('Set Map  Center (tmp.lon, tmp.lat)');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.1 Units "km"');
end;

procedure TfmBrowse.mmAscendClick(Sender: TObject);
begin
  mmAscend.Checked := True;
  mmDesc.Checked := False;
end;

procedure TfmBrowse.mmDescClick(Sender: TObject);
begin
  mmAscend.Checked := False;
  mmDesc.Checked := True;
end;

procedure TfmBrowse.N2Click(Sender: TObject);
begin
  Print;
end;

end.


