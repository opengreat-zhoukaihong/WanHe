unit SelCdd1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, DBCtrls, Mask, dbtables, db, Buttons;

type
  TfmSelCdd = class(TForm)
    CELLBASE: TGroupBox;
    Panel1: TPanel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Panel3: TPanel;
    Label29: TLabel;
    edNcell1: TEdit;
    edNcell2: TEdit;
    edNcell3: TEdit;
    edNcell4: TEdit;
    edNcell5: TEdit;
    edNcell6: TEdit;
    edNcell7: TEdit;
    edNcell8: TEdit;
    edNcell9: TEdit;
    edNcell10: TEdit;
    edNcell11: TEdit;
    edNcell12: TEdit;
    edNcell13: TEdit;
    edNcell14: TEdit;
    edNcell15: TEdit;
    edNcell16: TEdit;
    edNcell17: TEdit;
    edNcell18: TEdit;
    edNcell19: TEdit;
    edNcell20: TEdit;
    Panel4: TPanel;
    Label31: TLabel;
    Panel2: TPanel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Panel5: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Log: TLabel;
    Lat: TLabel;
    IHO: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label30: TLabel;
    Panel7: TPanel;
    dedCellID: TDBEdit;
    dedBsc: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    DBEdit12: TDBEdit;
    DBEdit13: TDBEdit;
    DBEdit14: TDBEdit;
    DBEdit15: TDBEdit;
    DBEdit16: TDBEdit;
    DBEdit19: TDBEdit;
    DBEdit20: TDBEdit;
    DBEdit21: TDBEdit;
    DBEdit22: TDBEdit;
    DBEdit24: TDBEdit;
    DBEdit17: TDBEdit;
    DBEdit18: TDBEdit;
    DBEdit23: TDBEdit;
    DBEdit25: TDBEdit;
    DBEdit26: TDBEdit;
    DBEdit27: TDBEdit;
    DBEdit28: TDBEdit;
    DBEdit29: TDBEdit;
    DBEdit30: TDBEdit;
    DBEdit31: TDBEdit;
    DBEdit32: TDBEdit;
    DBEdit33: TDBEdit;
    DBEdit34: TDBEdit;
    DBEdit35: TDBEdit;
    DBEdit36: TDBEdit;
    DBEdit37: TDBEdit;
    DBEdit38: TDBEdit;
    DBEdit39: TDBEdit;
    DBEdit40: TDBEdit;
    DBEdit41: TDBEdit;
    DBEdit42: TDBEdit;
    DBEdit43: TDBEdit;
    DBEdit44: TDBEdit;
    DBEdit45: TDBEdit;
    DBEdit46: TDBEdit;
    DBEdit47: TDBEdit;
    DBEdit48: TDBEdit;
    DBEdit49: TDBEdit;
    DBEdit50: TDBEdit;
    DBEdit51: TDBEdit;
    DBEdit52: TDBEdit;
    DBEdit53: TDBEdit;
    DBEdit54: TDBEdit;
    DBEdit55: TDBEdit;
    DBEdit56: TDBEdit;
    DBEdit57: TDBEdit;
    DBEdit58: TDBEdit;
    DBEdit59: TDBEdit;
    DBEdit60: TDBEdit;
    DBEdit61: TDBEdit;
    DBEdit62: TDBEdit;
    DBEdit63: TDBEdit;
    DBEdit64: TDBEdit;
    DBEdit65: TDBEdit;
    DBEdit66: TDBEdit;
    DBEdit67: TDBEdit;
    edState: TEdit;
    edMsc: TEdit;
    Panel6: TPanel;
    Label9: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit68: TDBEdit;
    DBEdit69: TDBEdit;
    DBEdit70: TDBEdit;
    DBEdit71: TDBEdit;
    DBEdit72: TDBEdit;
    DBEdit73: TDBEdit;
    DBEdit74: TDBEdit;
    DBEdit75: TDBEdit;
    DBEdit76: TDBEdit;
    DBEdit77: TDBEdit;
    DBEdit78: TDBEdit;
    DBEdit79: TDBEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Query1: TQuery;
    procedure Button2Click(Sender: TObject);
    procedure dedCellIDChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dedBscChange(Sender: TObject);
    procedure dedCellIDExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSelCdd: TfmSelCdd;

implementation

{$R *.DFM}
uses bscData,  bscMain, detail;

procedure TfmSelCdd.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TfmSelCdd.dedCellIDChange(Sender: TObject);
var
  i, j : Integer;
begin

end;

procedure TfmSelCdd.FormCreate(Sender: TObject);
var
  wCell : String;
begin
wCell := UpperCase(oleMapInfo.eval('Selection.bs_no'));
try
  with dmBscData do
  begin
    with quCellID do
    begin

      wCell := UpperCase(oleMapInfo.eval('Selection.bs_no'));
      if not Active then
        Open;

      Locate('BS_NO',UpperCase(wCell),[loPartialKey]);

    end;
    with quRLCFP do
      if not Active then
        Open;
    with quRLCPP do
      if not Active then
        Open;
    with quRLCXP do
      if not Active then
        Open;
    with quRLDEP do
      if not Active then
        Open;
    with quRLIHP do
      if not Active then
        Open;
    with quRLLOP do
      if not Active then
        Open;
    with quRLMFP do
      if not Active then
        Open;
   {  with quRLNRP do
      if not Active then
        Open;  }
    with quRLSBP do
      if not Active then
        Open;
    with quRLSSP do
      if not Active then
        Open;
  end;

{  if not dmBscData.quCellID.Active then
    Exit; }
  with dmBscData.quRLNRP do
  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString := wCell;
    Open;
    First;
    if RecordCount = 0 then
      Exit;
    
    if not Eof then
    begin
      edNcell1.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
      edNcell2.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
       edNcell3.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell4.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell5.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell6.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell7.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell8.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell9.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell10.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell11.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell12.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell13.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell14.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell15.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell16.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell17.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell18.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell19.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell20.Text := FieldByName('CellR').AsString;
       Next;
    end;
  end;
{  with dmBscData.quRLMFP do
  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString := dmBscData.quCellID.FieldByName('bs_no').AsString;
    Open;
    First;
  end;}



except

end;

try
  if not dmBscData.quCellID.Active then
    Exit;
  with dmBscData.quRLNRP do
  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString := wCell;
    Open;
    First;
    if RecordCount = 0 then
      Exit;

    if not Eof then
    begin
      edNcell1.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
      edNcell2.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
       edNcell3.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell4.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell5.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell6.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell7.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell8.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell9.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell10.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell11.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell12.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell13.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell14.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell15.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell16.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell17.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell18.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell19.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell20.Text := FieldByName('CellR').AsString;
       Next;
    end;
  end;
  with dmBscData.quRLMFP do
  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString := dmBscData.quCellID.FieldByName('bs_no').AsString;
    Open;
    First;
  end;
except
end;






  //dedState.Text := 'ACTIVE';
end;

procedure TfmSelCdd.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 // oleMapInfo.do('Set Style pen makepen(2 ,59, RGB(0,0,255))') ;
  //oleMapInfo.RunMenuCommand(810);
   with dmBscData do
  begin
    with quCellID do
      if Active then
        Close;
    with quRLCFP do
      if Active then
        Close;
    with quRLCPP do
      if Active then
        Close;
    with quRLCXP do
      if Active then
        Close;
    with quRLDEP do
      if Active then
        Close;
    with quRLIHP do
      if Active then
        Close;
    with quRLLOP do
      if Active then
        Close;
    with quRLMFP do
      if Active then
        Close;
     with quRLNRP do
      if Active then
        Close;
    with quRLSBP do
      if Active then
        Close;
    with quRLSSP do
      if Active then
        Close;
  end;
  
end;

procedure TfmSelCdd.dedBscChange(Sender: TObject);
var
  wStr : String;
begin
  wStr := Copy(dedBsc.Text,5,1);
  if wStr = 'A' then
    edMsc.Text := 'MSC_A';
  if wStr = 'B' then
    edMsc.Text := 'MSC_B';
  if wStr = 'C' then
    edMsc.Text := 'MSC_C';
end;

procedure TfmSelCdd.dedCellIDExit(Sender: TObject);
begin
  with dmBscData.quCellID do
  begin
    Locate('BS_NO',UpperCase(dedCellID.Text),[loPartialKey]);
  end;
end;

procedure TfmSelCdd.Button1Click(Sender: TObject);

begin

  //fmSelCdd.close;
end;

procedure TfmSelCdd.SpeedButton1Click(Sender: TObject);
begin
   fdetail:=tfdetail.create(self);
   fdetail.showMODAL;
   fdetail.free;
end;

procedure TfmSelCdd.FormShow(Sender: TObject);
var
  wCell : String;
begin
{try
  with dmBscData do
  begin
    with quCellID do
    begin

      wCell := fmCddMain.edCell.Text;
      if not Active then
        Open;

      Locate('BS_NO',UpperCase(wCell),[loPartialKey]);

    end;
    with quRLCFP do
      if not Active then
        Open;
    with quRLCPP do
      if not Active then
        Open;
    with quRLCXP do
      if not Active then
        Open;
    with quRLDEP do
      if not Active then
        Open;
    with quRLIHP do
      if not Active then
        Open;
    with quRLLOP do
      if not Active then
        Open;
    with quRLMFP do
      if not Active then
        Open;
     with quRLNRP do
      if not Active then
        Open;
    with quRLSBP do
      if not Active then
        Open;
    with quRLSSP do
      if not Active then
        Open;
  end;
except

end; }

//try
{  if not dmBscData.quCellID.Active then
    Exit; }
  with query1 do
  begin
    if Active then
      Close;
      sql.clear;
      sql.add('  select cellid,cellr  from rlnrp  where cellid=:cellid   group by cellid,cellr ');
      wCell := UpperCase(oleMapInfo.eval('Selection.bs_no'));
      parambyname('cellid').asstring := wCell;//'ZHAHZS3';
      open;

//    ParamByName('bs_no').AsString :=UPPERCASE(fmCddMain.edCell.Text);
//    Open;
    First;
    if RecordCount = 0 then
      Exit;

    if not Eof then
    begin
      edNcell1.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
      edNcell2.Text := FieldByName('CellR').AsString;
      Next;
    end;
    if not Eof then
    begin
       edNcell3.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell4.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell5.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell6.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell7.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell8.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell9.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell10.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell11.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell12.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell13.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell14.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell15.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell16.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell17.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell18.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell19.Text := FieldByName('CellR').AsString;
       Next;
    end;
    if not Eof then
    begin
       edNcell20.Text := FieldByName('CellR').AsString;
       Next;
    end;
  end;
  with dmBscData.quRLMFP do
  begin
    if Active then
      Close;
    ParamByName('bs_no').AsString := dmBscData.quCellID.FieldByName('bs_no').AsString;
    Open;
    First;
  end;
//except
//end;
  //dedState.Text := 'ACTIVE';
  oleMapInfo.do('select * from cell where bs_no="' + wCell + '" into tmp');
  edNcell1.Text := oleMapInfo.eval('tmp.Ncell1');
  edNcell2.Text := oleMapInfo.eval('tmp.Ncell2');
  edNcell3.Text := oleMapInfo.eval('tmp.Ncell3');
  edNcell4.Text := oleMapInfo.eval('tmp.Ncell4');
  edNcell5.Text := oleMapInfo.eval('tmp.Ncell5');
  edNcell6.Text := oleMapInfo.eval('tmp.Ncell6');
  edNcell7.Text := oleMapInfo.eval('tmp.Ncell7');
  edNcell8.Text := oleMapInfo.eval('tmp.Ncell8');
  edNcell9.Text := oleMapInfo.eval('tmp.Ncell9');
  edNcell10.Text := oleMapInfo.eval('tmp.Ncell10');
  edNcell11.Text := oleMapInfo.eval('tmp.Ncell11');
  edNcell12.Text := oleMapInfo.eval('tmp.Ncell12');
  edNcell13.Text := oleMapInfo.eval('tmp.Ncell13');
  edNcell14.Text := oleMapInfo.eval('tmp.Ncell14');
  edNcell15.Text := oleMapInfo.eval('tmp.Ncell15');
  edNcell16.Text := oleMapInfo.eval('tmp.Ncell16');
  oleMapInfo.do('close table tmp');
end;

procedure TfmSelCdd.SpeedButton2Click(Sender: TObject);
begin
close;
end;

procedure TfmSelCdd.SpeedButton3Click(Sender: TObject);
var
  wLon1, wLat1, wBearing1,wLon2, wLat2,wBearing2 : real;
  i, wRow : Integer;
begin
  //oleMapInfo.do('close table ncell_tmp');
  //oleMapInfo.do('Set Map Layer cell Editable On');
  
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');

  oleMapInfo.do('Set Style pen makepen(1 , 26, RGB(0, 0,0))');

  oleMapInfo.do('Open table "' + gExePath + 'ncell.tab" Interactive');
  oleMapInfo.do('select * from ncell where bs_no = "' + dedCellID.Text + '" into tmp');
  oleMapInfo.do('commit table tmp as "' + gExePath + 'ncell_tmp.tab"');
  oleMapInfo.do('Open table "' + gExePath + 'ncell_tmp.tab" Interactive');

  oleMapInfo.do('Add Column "ncell_tmp" (Lon_s Decimal (12, 6))From Cell Set To Lon Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell_tmp" (Lon_t Decimal (12, 6))From Cell Set To Lon Where COL2 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell_tmp" (Lat_s Decimal (12, 6))From Cell Set To Lat Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell_tmp" (Lat_t Decimal (12, 6))From Cell Set To Lat Where COL2 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell_tmp" (Bearing_s Decimal (12, 6))From Cell Set To Bearing Where COL1 = COL2  Dynamic');
  oleMapInfo.do('Add Column "ncell_tmp" (Bearing_t Decimal (12, 6))From Cell Set To Bearing Where COL2 = COL2  Dynamic');

  oleMapInfo.do('Create Map For ncell_tmp CoordSys Earth Projection 1, 0');
  wRow := oleMapInfo.eval('tableinfo(ncell_tmp,8)');
  oleMapInfo.do('fetch first from ncell_tmp');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from ncell_tmp');
    wLon1 := oleMapInfo.eval('ncell_tmp.Lon_s');
    wLat1 := oleMapInfo.eval('ncell_tmp.Lat_s');
    wLon2 := oleMapInfo.eval('ncell_tmp.Lon_t');
    wLat2 := oleMapInfo.eval('ncell_tmp.Lat_t');
    //wRate := oleMapInfo.eval('Hdov_i_bsc.hovercnt') / wMaxQty ;
    wBearing1 := oleMapInfo.eval('ncell_tmp.Bearing_s');
    wBearing2 := oleMapInfo.eval('ncell_tmp.Bearing_t');
    fmBscMain.DrawLine(wLon1, wLat1, wLon2, wLat2, gCellLength, wBearing1, wBearing2, oleMapInfo, i, 'ncell_tmp');
  end;
  oleMapInfo.do('commit table ncell_tmp');
  oleMapInfo.do('add map auto layer ncell_tmp');
  oleMapInfo.do('Set Map  Center (ncell_tmp.lon_s, ncell_tmp.lat_s)');
 /////////////////////////////////////////
 {with dmBscData.quCellID do
  begin
    wLon1 :=  FieldByName('lon').AsFloat;
    wLat1 := FieldByName('lat').AsFloat;
    wBearing1 := FieldByName('bearing').AsFloat;
    oleMapInfo.do('Set Map  Center (' + FieldByName('lon').AsString + ',' +
                     FieldByName('lat').AsString + ')');
  end;
  with dmBscData.quRLNRP do
  begin
    First;
    while not eof do
    begin
      with dmBscData.quSelCell do
      begin
        if Active then
          Close;
        ParamByName('bs_no').AsString :=
          dmBscData.quCellID.FieldByName('Bs_no').AsString;
        Open;
        if not IsEmpty then
        begin
          wLon2 :=  FieldByName('lon').AsFloat;
          wLat2 := FieldByName('lat').AsFloat;
          wBearing2 := FieldByName('bearing').AsFloat;
          fmBscMain.DrawLine(wLon1, wLat1, wLon2, wLat2, 0.0025, wBearing1, wBearing2, oleMapInfo);
        end;
      end;
      Next;
    end;
  end;
  oleMapInfo.RunMenuCommand(610);   }
  fmSelCdd.close;
end;

end.
