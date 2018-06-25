unit DataList;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Wwdbigrd, Wwdbgrid, Db, DBTables, Wwquery, Wwdatsrc;

type
  TfmDataList = class(TForm)
    quDataList: TwwQuery;
    wwDBGrid1: TwwDBGrid;
    dsDataList: TwwDataSource;
    quSelCell: TQuery;
    procedure wwDBGrid1DblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDataList: TfmDataList;
  //wTableName : String;
  
implementation

{$R *.DFM}
uses BscData, BscMain;

procedure TfmDataList.wwDBGrid1DblClick(Sender: TObject);
var
  i : Integer;
begin

  with dmBscData.quSelCell do

  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString :=
      quDataList.FieldByName('cellid').AsString;
    Open;
    if not IsEmpty then
    begin
      oleMapInfo.do('Set Map  Center (' + FieldByName('lon').AsString + ',' +
                     FieldByName('lat').AsString + ')');
    end;
  end;
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');
end;

end.
