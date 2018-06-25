unit FindNcell;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmFindNCell = class(TForm)
    edLength: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmFindNCell: TfmFindNCell;

implementation

{$R *.DFM}
uses bscMain;

procedure TfmFindNCell.BitBtn1Click(Sender: TObject);
var
  wLon1, wLat1, wBearing1, wLon2, wLat2, wBearing2, wLength : real;
  i, wRow, wNum : Integer;
function GetMultiNCell(V : TStrings) : String;
var
  i : Integer;
  wCellId : string;
begin
  Result := ' (';
  for i := 0 to V.Count - 1 do
  begin
    wCellId := Copy(V.Strings[i], 1, Pos(' ', V.Strings[i]) - 1);
    if i < V.Count -1 then
      Result := Result + 'Ncell.bs_no = "' + wCellId + '" or '
    else
      Result := Result + 'Ncell.bs_no = "' + wCellId + '" ) ';
  end;
end;
begin
  //oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');

  oleMapInfo.do('Set Style pen makepen(2 ,2, RGB(255,0,0))');

  oleMapInfo.do('Open table "' + gExePath + 'ncell.tab" Interactive');
  oleMapInfo.do('Add Column "ncell" (Lon_s Decimal (12, 6))From Cell Set To Lon Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Lon_t Decimal (12, 6))From Cell Set To Lon Where COL2 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Lat_s Decimal (12, 6))From Cell Set To Lat Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Lat_t Decimal (12, 6))From Cell Set To Lat Where COL2 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Bearing_s Decimal (12, 6))From Cell Set To Bearing Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Bearing_t Decimal (12, 6))From Cell Set To Bearing Where COL2 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Bsc_No_s Char(6))From Cell Set To Bsc_no Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell" (Bsc_No_t Char(6))From Cell Set To Bsc_no Where COL2 = COL2  Dynamic');
  if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('select * from ncell where bsc_no_s = "'
                 + gSelName + '" into  tmp')
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('select * from ncell  where bs_no = "'  + gSelName + '" into tmp')
      else
        oleMapInfo.do('select * from ncell  where ' + GetMultiNCell(gMultiCell)  + ' into tmp');
    end
    else
      oleMapInfo.do('select * from ncell into tmp');
  end;

  {if gSelFlag = 'BSC' then
    oleMapInfo.do('select * from ncell where bsc_no_s = "'
                 + gSelName + '" into  tmp')
  else
    oleMapInfo.do('select * from ncell into tmp'); }

  oleMapInfo.do('commit table tmp as "' + gExePath + 'ncell_length.tab"');
  oleMapInfo.do('Open table "' + gExePath + 'ncell_length.tab" Interactive');
  oleMapInfo.do('close table ncell');
  //oleMapInfo.do('close table tmp');


  oleMapInfo.do('Create Map For ncell_length CoordSys Earth Projection 1, 0');
  wRow := oleMapInfo.eval('tableinfo(ncell_length,8)');
  oleMapInfo.do('fetch first from ncell_length');
  wNum := 0;
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from ncell_length');
    wLon1 := oleMapInfo.eval('ncell_length.Lon_s');
    wLat1 := oleMapInfo.eval('ncell_length.Lat_s');
    wLon2 := oleMapInfo.eval('ncell_length.Lon_t');
    wLat2 := oleMapInfo.eval('ncell_length.Lat_t');
    //wRate := oleMapInfo.eval('Hdov_i_bsc.hovercnt') / wMaxQty ;
    wBearing1 := oleMapInfo.eval('ncell_length.Bearing_s');
    wBearing2 := oleMapInfo.eval('ncell_length.Bearing_t');
    wLength := sqrt(sqr((wLon1 - wLon2) * 103.1) + sqr((wLat1 - wLat2) * 111.2));
    if (wLength >= StrToFloat(edLength.Text)) and (wLon1 <> 0) and (wLat1 <> 0)
       and (wLon2 <> 0) and (wLat2 <> 0) then
    begin
      fmBscMain.DrawLine(wLon1, wLat1, wLon2, wLat2, gCellLength, wBearing1, wBearing2, oleMapInfo, i, 'ncell_length');
      wNum := wNum + 1;
    end;
  end;
  oleMapInfo.do('commit table ncell_length');
  oleMapInfo.do('add map auto layer ncell_length');
 // oleMapInfo.do('Set Map  Center (ncell_length.lon_s, ncell_length.lat_s)');
  oleMapInfo.do('Set Map Layer ncell_length Label Position Above Font ("Arial",0,10,0) ' +
                ' With bs_no+" -- "+ncell_id' +
                ' Auto On Visibility Zoom (0, 100) Units "km"');
  ShowMessage('共有 ' + IntToStr(wNum) + '对邻小区距离大于' + edLength.Text + 'km');
end;

end.
