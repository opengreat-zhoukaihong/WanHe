unit BscMain;

interface

uses
  Windows, SysUtils, Classes, Graphics, Forms, Controls, Menus,
  StdCtrls, Dialogs, Buttons, Messages, ExtCtrls, ComCtrls, StdActns,
  ActnList, ToolWin, ImgList, ComObj, db, DBTables,Wireless_TLB;

type
  TfmBscMain = class(TForm)
    odBscMain: TOpenDialog;
    alBscMain: TActionList;
    EditCut1: TEditCut;
    EditCopy1: TEditCopy;
    EditPaste1: TEditPaste;
    WindowCascade1: TWindowCascade;
    WindowTileHorizontal1: TWindowTileHorizontal;
    WindowArrangeAll1: TWindowArrange;
    WindowMinimizeAll1: TWindowMinimizeAll;
    HelpAbout1: TAction;
    WindowTileVertical1: TWindowTileVertical;
    ToolBar2: TToolBar;
    ToolButton3: TToolButton;
    ToolButton7: TToolButton;
    ilBscMain: TImageList;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    SpeedButton18: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    ToolButton4: TToolButton;
    SpeedButton5: TSpeedButton;
    odData: TOpenDialog;
    SpeedButton8: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    SpeedButton21: TSpeedButton;
    SpeedButton22: TSpeedButton;
    SpeedButton23: TSpeedButton;
    SpeedButton24: TSpeedButton;
    SpeedButton25: TSpeedButton;
    SpeedButton26: TSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    sbBscMain: TStatusBar;
    Label1: TLabel;
    Label2: TLabel;
    SpeedButton27: TSpeedButton;
    SpeedButton28: TSpeedButton;
    cbNext: TComboBox;
    SpeedButton29: TSpeedButton;
    SpeedButton30: TSpeedButton;
    tbBsc: TTable;
    mmBscMain: TMainMenu;
    mmSysSet: TMenuItem;
    mmOpenMap: TMenuItem;
    mmSelData: TMenuItem;
    mmDataConv: TMenuItem;
    MSC1: TMenuItem;
    BSC1: TMenuItem;
    CTR1: TMenuItem;
    CDD2: TMenuItem;
    N44: TMenuItem;
    N45: TMenuItem;
    N47: TMenuItem;
    N38: TMenuItem;
    mmLoadMap: TMenuItem;
    mmUpdateData: TMenuItem;
    N52: TMenuItem;
    mmResetCell: TMenuItem;
    mmResetBase: TMenuItem;
    DFDF1: TMenuItem;
    N48: TMenuItem;
    N35: TMenuItem;
    N49: TMenuItem;
    N15: TMenuItem;
    N19: TMenuItem;
    N32: TMenuItem;
    N33: TMenuItem;
    mmMapConf: TMenuItem;
    N2: TMenuItem;
    mmExit: TMenuItem;
    mmDailyRep: TMenuItem;
    mmTCH: TMenuItem;
    mmSDCCH: TMenuItem;
    N18: TMenuItem;
    N1: TMenuItem;
    N3: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N23: TMenuItem;
    N10: TMenuItem;
    N22: TMenuItem;
    N24: TMenuItem;
    N25: TMenuItem;
    N26: TMenuItem;
    N27: TMenuItem;
    N28: TMenuItem;
    N29: TMenuItem;
    N30: TMenuItem;
    N31: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N41: TMenuItem;
    N50: TMenuItem;
    N72: TMenuItem;
    CDD1: TMenuItem;
    mmAllCdd: TMenuItem;
    mmSelCdd: TMenuItem;
    CDD5: TMenuItem;
    N39: TMenuItem;
    N40: TMenuItem;
    N53: TMenuItem;
    mmAnalysis: TMenuItem;
    N17: TMenuItem;
    N73: TMenuItem;
    N4: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N16: TMenuItem;
    N5: TMenuItem;
    NAP1: TMenuItem;
    N6: TMenuItem;
    G1: TMenuItem;
    mmSu: TMenuItem;
    mmDr: TMenuItem;
    mmTraffic: TMenuItem;
    mmHover: TMenuItem;
    mmRaf: TMenuItem;
    mmSaSs: TMenuItem;
    mmChAc: TMenuItem;
    mmDqa: TMenuItem;
    mmDss: TMenuItem;
    mmMH: TMenuItem;
    N9: TMenuItem;
    mmTraShade: TMenuItem;
    mmDensityShade: TMenuItem;
    N37: TMenuItem;
    N36: TMenuItem;
    N46: TMenuItem;
    N61: TMenuItem;
    N62: TMenuItem;
    N63: TMenuItem;
    N65: TMenuItem;
    N66: TMenuItem;
    N67: TMenuItem;
    N68: TMenuItem;
    N69: TMenuItem;
    N70: TMenuItem;
    N71: TMenuItem;
    W1: TMenuItem;
    N34: TMenuItem;
    N54: TMenuItem;
    N43: TMenuItem;
    N51: TMenuItem;
    N55: TMenuItem;
    N56: TMenuItem;
    N57: TMenuItem;
    N58: TMenuItem;
    N59: TMenuItem;
    N60: TMenuItem;
    Help1: TMenuItem;
    N42: TMenuItem;
    HelpAboutItem: TMenuItem;
    N74: TMenuItem;
    N75: TMenuItem;
    N76: TMenuItem;
    N77: TMenuItem;
    N78: TMenuItem;
    N79: TMenuItem;
    N80: TMenuItem;
    N81: TMenuItem;
    N82: TMenuItem;
    N83: TMenuItem;
    N84: TMenuItem;
    N85: TMenuItem;
    N86: TMenuItem;
    N87: TMenuItem;
    N88: TMenuItem;
    N89: TMenuItem;
    N90: TMenuItem;
    BA1: TMenuItem;
    N91: TMenuItem;
    BSC2: TMenuItem;
    N92: TMenuItem;
    N93: TMenuItem;
    N94: TMenuItem;
    N95: TMenuItem;
    N96: TMenuItem;
    N97: TMenuItem;
    N98: TMenuItem;
    N99: TMenuItem;
    N100: TMenuItem;
    N101: TMenuItem;
    N102: TMenuItem;
    N103: TMenuItem;
    N104: TMenuItem;
    N105: TMenuItem;
    N106: TMenuItem;
    N108: TMenuItem;
    N109: TMenuItem;
    N110: TMenuItem;
    N64: TMenuItem;
    procedure FileNew1Execute(Sender: TObject);
    //procedure FileOpen1Execute(Sender: TObject);
    procedure HelpAbout1Execute(Sender: TObject);
    procedure FileExit1Execute(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton17Click(Sender: TObject);
    procedure SpeedButton16Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton12Click(Sender: TObject);
    procedure SpeedButton13Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton15Click(Sender: TObject);
    procedure SpeedButton18Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure mmTCHClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N21Click(Sender: TObject);
    procedure N25Click(Sender: TObject);
    procedure N26Click(Sender: TObject);
    procedure mmDrClick(Sender: TObject);
    procedure mmSaSsClick(Sender: TObject);
    procedure mmSuClick(Sender: TObject);
    procedure mmChClick(Sender: TObject);
    procedure mmTrafficClick(Sender: TObject);
    procedure mmDqaClick(Sender: TObject);
    procedure mmRafClick(Sender: TObject);
    procedure mmHoverClick(Sender: TObject);
    procedure mmSelDataClick(Sender: TObject);
    procedure mmOpenMapClick(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    //procedure mmCgShadeClick(Sender: TObject);
    procedure mmTraShadeClick(Sender: TObject);
    procedure MSC1Click(Sender: TObject);
    procedure mmSDCCHClick(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N27Click(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N31Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure SpeedButton21Click(Sender: TObject);
    procedure SpeedButton20Click(Sender: TObject);
    procedure SpeedButton19Click(Sender: TObject);
    procedure SpeedButton23Click(Sender: TObject);
    procedure SpeedButton24Click(Sender: TObject);
    procedure SpeedButton22Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure N32Click(Sender: TObject);
    procedure N33Click(Sender: TObject);
    procedure SpeedButton25Click(Sender: TObject);
    procedure SpeedButton26Click(Sender: TObject);
    procedure mmMapConfClick(Sender: TObject);
    procedure N46Click(Sender: TObject);
    procedure mmAllCddClick(Sender: TObject);
    procedure mmSelCddClick(Sender: TObject);
    procedure mmResetBaseClick(Sender: TObject);
    procedure mmResetCellClick(Sender: TObject);
    procedure SpeedButton27Click(Sender: TObject);
    procedure N53Click(Sender: TObject);
    procedure N39Click(Sender: TObject);
    procedure N40Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure N34Click(Sender: TObject);
    procedure cbNextChange(Sender: TObject);
    procedure SpeedButton29Click(Sender: TObject);
    procedure mmDssClick(Sender: TObject);
    procedure mmMHClick(Sender: TObject);
    procedure CDD2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure mmDensityShadeClick(Sender: TObject);
    procedure SpeedButton28Click(Sender: TObject);
    procedure mmLoadMapClick(Sender: TObject);
    procedure mmUpdateDataClick(Sender: TObject);
    procedure FormDblClick(Sender: TObject);
    procedure DFDF1Click(Sender: TObject);
    procedure N36Click(Sender: TObject);
    procedure N110Click(Sender: TObject);
    procedure N105Click(Sender: TObject);
    procedure N106Click(Sender: TObject);
    procedure N104Click(Sender: TObject);
    procedure N108Click(Sender: TObject);
    procedure N109Click(Sender: TObject);
    procedure N63Click(Sender: TObject);
    procedure N64Click(Sender: TObject);
    procedure N61Click(Sender: TObject);
  private
    { Private declarations }
    procedure TchCGShade;
    Procedure TchTrafficShade;
    procedure CreateRegion_4(X, Y, L, P, Bearing, Rate : Real; V : Variant; Flag : Integer; LR : String);
    procedure CreateMDIChild(const Name: string);

    procedure CreateRegion_5(X, Y, L, P, Bearing, Rate : Real; V : Variant; Flag : Integer);
    procedure CreateCircle(X, Y, L, P, Bearing, Rate, R : Real; V : Variant; Flag : Integer);
    procedure UpDateLine(Lon_1, Lat_1, Lon_2, Lat_2, L, P, Bearing_1, Bearing_2: Real; Rate: Real; FlunkRate : Real; V : Variant; RowId : Integer; TableName, CellSe, CellTa: string;  LostRate, Length_rate : Real );
    procedure UpDateCircle(X, Y, L, P, Bearing, Rate : Real; V : Variant; RowId, Flag : Integer; TableName : String);
    procedure UpdateCellObject;
    procedure UpdateASObject(X, Y, L, P, Bearing, R1, R2 : Real; V : Variant; RowId, Flag : Integer; TableName : String);
    function GetMultiCell(V : TStrings) : String;
    //procedure UpdateSDRObject;
  public
    { Public declarations }
    wFilterOrder, wTchCheck, wTchFlag, wTchQty, wSdcchCheck, wSdcchFlag, wSdcchQty,
    wCondition, wFristTchTable, wPriorTchTable, wFristCchTable, wPriorCchTable  : String;
    wBrowseShow : Boolean;
    procedure DrawLine(Lon_1, Lat_1, Lon_2, Lat_2, L,  Bearing_1, Bearing_2 : Real; V : Variant ; rowid : Integer; tablename : String );
    procedure CreateRegion_6(X, Y, L: Real; V : Variant; Flag : Integer);
    procedure CreateRegion_3(X, Y, L, D, Bearing, Rate : Real; V : Variant);
    Procedure ShowBsc;
    Procedure ShowCondition(Desc : String);
    Procedure ResetMap;
  end;

var
  fmBscMain: TfmBscMain;
  oleMapInfo : oleVariant;
  sWinHand, gMulCellName, gCompLayer : String;
  gDr, gHover, gRaf, gTraffic, gChAc, gDss,
  gSaSs, gSu, gDqa, gTchCG, gTchTraffic, gMh,
  gBscShowing, gCQT : Boolean;
  gSelDate,  gSelName, gSelFlag, gExePath, gCtrExePath : String;
  gMultiCell : TStrings;
  gMapNo : Integer;
  gCellAngle, gCellLength : Real;
  wWirClass: wClass;

implementation
{$R *.DFM}

uses  Childwin, Find, About, legeng, DataHist, BscData, Compare, MapConf,
  AllCdd, SelCdd, FindNcell, Trendline, Condition, CellConf, Browse,
  HdovCond, CQTDlg,dm,punit,ctr_globe;

type
  TCreateObj = function(MapInfo: Variant; UserType: String): Boolean; stdcall;
  TColorDlg = function(MapInfo: Variant; UserType: String): Boolean; stdcall;

var
  PFunc : TFarProc;
  Module : THandle;

  
procedure TfmBscMain.ResetMap;
begin
  if gMh then
  begin
    oleMapInfo.do('Close table mh_Shade  Interactive');
    
    //oleMapInfo.do('Close table TCh_dr  Interactive');
  end;
  if gDr then
  begin
    oleMapInfo.do('Close table CCh_Sdr  Interactive');
    oleMapInfo.do('Close table TCh_dr  Interactive');
  end;
  if gHover then
  begin
    oleMapInfo.do('Close table HDOV_I_BSC  Interactive');
    oleMapInfo.do('Close table HDOV_E_BSC  Interactive');
    oleMapInfo.do('Close table HDOV_I_61BSC  Interactive');
    oleMapInfo.do('Close table HDOV_E_61BSC  Interactive');
  end;
  if gTraffic then
  begin
    oleMapInfo.do('Close table CCh_Traffic  Interactive');
    oleMapInfo.do('Close table TCh_Traffic  Interactive');
  end;
  if gChAc then
  begin
    oleMapInfo.do('Close table CCh_ch  Interactive');
    oleMapInfo.do('Close table TCh_ch  Interactive');
  end;
  if gSaSs then
  begin
    oleMapInfo.do('Close table CCh_Sa  Interactive');
    oleMapInfo.do('Close table TCh_ca  Interactive');
  end;
  if gRaf then
  begin
    oleMapInfo.do('Close table raf_shade  Interactive');
  end;
  if gDqa then
  begin
    oleMapInfo.do('Close table CCh_dqa  Interactive');
    oleMapInfo.do('Close table TCh_tqa  Interactive');

  end;
  if gDss then
  begin
    oleMapInfo.do('Close table CCh_dss4  Interactive');
    oleMapInfo.do('Close table TCh_tss4  Interactive');
  end;
  if gRaf then
  begin
  end;
  if gSu then
  begin
    oleMapInfo.do('Close table CCh_Su  Interactive');
    oleMapInfo.do('Close table TCh_u  Interactive');
  end;
  if gTchCG then
  begin
     oleMapInfo.do('Close table cg_shade  Interactive');
  end;
  if gTchTraffic then
  begin
     oleMapInfo.do('Close table erpac_shade  Interactive');
  end;
  gTchCG := False;
  gTchTraffic := False;
  gDr := False;
  gHover := False;
  gChAc := False;
  gDqa := False;
  gSaSs := False;
  gRaf := False;
  gSu := False;
  gTraffic := False;

  fmBscMain.mmTraShade.Checked := False;
  //fmBscMain.mmCgShade.Checked := False;
  fmBscMain.mmdr.Checked := False;
  fmBscMain.mmHover.Checked := False;
  fmBscMain.mmChAc.Checked := False;
  fmBscMain.mmDqa.Checked := False;
  fmBscMain.mmSaSs.Checked := False;
  fmBscMain.mmRaf.Checked := False;
  fmBscMain.mmSu.Checked := False;
  fmBscMain.mmTraffic.Checked := False;
end;

Procedure  TfmBscMain.ShowCondition(Desc : String);
begin
  Application.CreateForm(TfmCondition, fmCondition);
  try
    fmCondition.ckFilterTCH.Caption := 'TCH ' + Desc;
    fmCondition.ckFilterSDCCH.Caption := 'SDCCH ' + Desc;
    fmCondition.ckOrderTch.Caption := 'TCH ' + Desc;

    fmCondition.ckOrderSdcch.Caption := 'SDCCH ' + Desc;
    if (Desc = '随机失败率') or (Desc = '平均通话时间')  then
    begin
      if  (Desc = '平均通话时间') then
        fmCondition.lbTch.Visible := False;
      fmCondition.ckFilterTCH.Caption :=  Desc;
      fmCondition.ckOrderTch.Caption :=  Desc;
      fmCondition.ckFilterSDCCH.Visible := False;
      fmCondition.ckOrderSdcch.Visible := False;
      fmCondition.edFilterSdcchQty.Visible := False;
      fmCondition.edOrderSdcchQty.Visible := False;
      fmCondition.cbFilterSDCCH.Visible := False;
      fmCondition.lbCch.Visible := False;
      fmCondition.lbCch1.Visible := False;
      fmCondition.lbCch2.Visible := False;
      fmCondition.edOrderSdcchQty.Visible := False;
      fmCondition.cbOrderSdcch.Visible := False;

    end;
    fmCondition.ShowModal;
  finally
    fmCondition.Free;
  end;
end;

Procedure TfmBscMain.ShowBsc;
var
  i, wRow, wTableNum, wBscLayerNum : Integer;
  wBscNo, wMsg : String;
  wLon, wLat: real;
begin
  if not gBscShowing then
  begin
    oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,50))');
    oleMapInfo.do('set style brush makebrush(64,rgb(0,255,50),rgb(0,255,50))');
    oleMapInfo.do('Open table "' + gExePath + 'bsc.tab" Interactive');
    wRow := oleMapInfo.eval('tableinfo(bsc,8)');
    for i := 1 to wRow do
    begin
      oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc');
      wBscNo := oleMapInfo.eval('bsc.bsc_no');
      oleMapInfo.do('select avg(lon), avg(lat) from cell where bsc_no ="' + wBscNo +'" into tmp');
      wLon := oleMapInfo.eval('tmp.col1');
      wLat := oleMapInfo.eval('tmp.col2');
      oleMapInfo.do('close table tmp');
      CreateRegion_6(wLon, wLat, 0.02, oleMapInfo, 2);
      oleMapInfo.do('update bsc set obj = TmpObject, lon =' +
              FloatToStr(wLon) + ', lat = ' +
              FloatToStr(wLat) + '  where rowid =' + IntToStr(i));
    end;
    oleMapInfo.do('commit table bsc');
    oleMapInfo.do('add map auto layer bsc');
    oleMapInfo.do('Set Map Layer bsc Label Position Below Font ("Arial",256,12,16777215,160) ' +
               'With Bsc_no  Auto On Offset 10  Visibility Zoom (0, 100) Units "km"');
    oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_file.tab" Interactive');
    oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_file.tab" Interactive');
    oleMapInfo.do('Set Map Layer cell Display Off');
    oleMapInfo.do('Set Map Layer base Display Off');
    oleMapInfo.do('set Map Layer street Display off');
  end;
 { wTableNum := oleMapInfo.eval('NumTables()');
  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'BSC' then
    begin
      wBscLayerNum := i;
      Break;
    end;
  end; }
 // wMsg := 'Set Map Order ' + IntToStr(wBscLayerNum) + ', 1';
  {for i := 1 to 3 do
    wMsg := wMsg + ' ,' + IntToStr(i); }
//  oleMapInfo.do(wMsg);
  gBscShowing := True;
end;

procedure TfmBscMain.CreateRegion_6(X, Y, L: Real; V : Variant; Flag : Integer);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6 : Real;
begin
  {X1 := X + L ;
  Y1 := Y ;
  X2 := X + L * Cos(60 * DEG_2_Rad);
  Y2 := Y - L * Sin(60 * DEG_2_Rad) * (103.1 / 111.2);
  X3 := X - L * Cos(60 * DEG_2_Rad);
  Y3 := Y - L * Sin(60 * DEG_2_Rad) * (103.1 / 111.2);
  X4 := X - L ;
  Y4 := Y ;
  X5 := X - L * Cos(60 * DEG_2_Rad);
  Y5 := Y + L * Sin(60 * DEG_2_Rad) * (103.1 / 111.2);
  X6 := X + L * Cos(60 * DEG_2_Rad);
  Y6 := Y + L * Sin(60 * DEG_2_Rad) * (103.1 / 111.2); }

  X1 := X + L / 2 ;
  Y1 := Y + L / 2 * (103.1 / 111.2);
  X2 := X + L;
  Y2 := Y + L / 2 * (103.1 / 111.2);
  X3 := X + L;
  Y3 := Y - L / 2 * (103.1 / 111.2);
  X4 := X - L ;
  Y4 := Y - L / 2 * (103.1 / 111.2);
  X5 := X - L;
  Y5 := Y + L / 2 * (103.1 / 111.2);
  X6 := X - L / 2;
  Y6 := Y + L / 2 * (103.1 / 111.2); 

  if Flag = 1 then
     V.do('create region  1 7 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ') ('
                    + FloatToStr(X5) + ',' + FloatToStr(Y5) + ') ('
                    + FloatToStr(X6) + ',' + FloatToStr(Y6) + ')')
  else
    V.do('create region into variable TmpObject 1 7 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ') ('
                    + FloatToStr(X5) + ',' + FloatToStr(Y5) + ') ('
                    + FloatToStr(X6) + ',' + FloatToStr(Y6) + ')');
end;



procedure TfmBscMain.DrawLine(Lon_1, Lat_1, Lon_2, Lat_2, L,  Bearing_1, Bearing_2 : Real; V : Variant; rowid : Integer; tablename : String );
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2 : real;
begin
  X1 := Lon_1 + 0.5 * L * Sin(Bearing_1 * DEG_2_RAD);
  Y1 := lat_1 + 0.5 * L * Cos(Bearing_1 * DEG_2_RAD) * (103.1 / 111.2);
  X2 := Lon_2 + 0.5 * L * Sin(Bearing_2 * DEG_2_RAD);
  Y2 := lat_2 + 0.5 * L * Cos(Bearing_2 * DEG_2_RAD) * (103.1 / 111.2);

 
   V.do('update ' + TableName + ' set obj = createLine (' + FloatToStr(X1) + ',' + FloatToStr(Y1)
       + ',' +  FloatToStr(X2) + ',' + FloatToStr(Y2)  + ') where rowid = ' + IntToStr(rowid));
end;


function TfmBscMain.GetMultiCell(V : TStrings) : String;
var
  i : Integer;
  wCellId : string;
begin
  Result := ' (';
  for i := 0 to V.Count - 1 do
  begin
    wCellId := Copy(V.Strings[i], 1, Pos(' ', V.Strings[i]) - 1);
    if i < V.Count -1 then
      Result := Result + 'cell.bs_no = "' + wCellId + '" or '
    else
      Result := Result + 'cell.bs_no = "' + wCellId + '" ) ';
  end;
end;

procedure TfmBscMain.TchCGShade;
begin
  oleMapInfo.do('commit table cell as "' + gExePath + 'cg_shade.tab"');
  oleMapInfo.do('open table "' + gExePath + 'cg_shade.tab"');
  oleMapInfo.do('add map auto layer cg_shade');
  oleMapInfo.do('Add Column "cg_shade" (Cg Decimal (5, 2))From Tch_file ' +
                 'Set To Cg Where COL2 = COL6  Dynamic');//
  {oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
                 ' cg_shade with Cg ignore 0 ranges apply all use color '
                + ' Brush (2,65280,16777215)  0.07: 0.13 '
                + 'Brush (2,65280,16777215) Pen (1,2,0) ,0.13: 0.19 '
                + 'Brush (2,2154496,16777215) Pen (1,2,0) ,0.19: 0.29 '
                + 'Brush (2,4243456,16777215) Pen (1,2,0) ,0.29: 0.33 '
                + 'Brush (2,5287936,16777215) Pen (1,2,0) ,0.33: 0.38 '
                + 'Brush (2,7376896,16777215) Pen (1,2,0) ,0.38: 1.32 '
                + 'Brush (2,9465856,16777215) Pen (1,2,0) ,1.32: 2.81 '
                + 'Brush (2,11554816,16777215) Pen (1,2,0) ,2.81: 3.62 '
                + 'Brush (2,12599296,16777215) Pen (1,2,0) ,3.62: 10.16 '
                + 'Brush (2,14688256,16777215) Pen (1,2,0) ,10.16: 10.16 '
                + 'Brush (2,16711680,16777215) Pen (1,2,0) default '
                + 'Brush (2,16777215,16777215) Pen (1,2,0)  # use 1 round 0.0001 '
                + 'inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
                ' layer prev display on shades on symbols off lines off ' +
                'count on title auto Font ("Arial",0,12,0) subtitle auto ' +
                'Font ("Arial",0,11,0) ascending off ranges ' +
                'Font ("Arial",0,11,0) auto display off ,auto display on ,' +
                'auto display on ,auto display on ,auto display on , ' +
                'auto display on ,auto display on ,auto display on , ' +
                'auto display on ,auto display on ,auto display on '); }

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
                 ' cg_shade with Cg ignore 0 ranges apply all use color ' +
                 ' Brush (2,65280,16777215)  0: 1 Brush (2,65280,16777215) Pen (1,2,0) ,1: 3 ' +
                 ' Brush (2,4243456,16777215) Pen (1,2,0) ,3: 6 Brush (2,8421376,16777215) ' +
                 ' Pen (1,2,0) ,6: 10 Brush (2,12599296,16777215) Pen (1,2,0) ,10: 100 ' +
                 ' Brush (2,16711680,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) ' +
                 ' Pen (1,2,0)  # use 0 round 0.1 inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))
               +  ' layer prev display on shades on symbols off lines off count on ' +
               ' title auto Font ("Arial",0,12,0) subtitle auto Font ("Arial",0,11,0) ' +
               '  ascending off ranges Font ("Arial",0,11,0) auto display off ,' +
               ' auto display on ,auto display on ,auto display on ,auto display on ,' +
               ' auto display on ');

//set map redraw off
  oleMapInfo.do('Set Map Layer cg_shade Label Position Above Font ("Arial",1,10,0) With cg+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  gTchCG := True;


end;

Procedure TfmBscMain.TchTrafficShade;
begin
  oleMapInfo.do('commit table cell as "' + gExePath + 'erpac_shade.tab"');
  oleMapInfo.do('open table "' + gExePath + 'erpac_shade.tab"');
  oleMapInfo.do('add map auto layer erpac_shade');
  oleMapInfo.do('Add Column "erpac_shade" (Erpac Decimal (8, 2))From Tch_file Set To Erpac Where COL2 = COL6  Dynamic');
  //oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' erpac_shade with Erpac ignore 0 ranges apply all use color ' +
             ' Brush (2,65280,16777215)  0: 0.2 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             '0.2: 0.5 Brush (2,5287936,16777215) Pen (1,2,0) ,0.5: 0.7 ' +
             'Brush (2,11554816,16777215) Pen (1,2,0) ,0.7: 1.0 ' +
             'Brush (2,16711680,16777215) Pen (1,2,0) default ' +
             'Brush (2,65280,16777215) Pen (1,2,0)  # use 0 round 0.01 inflect off ' +
             'Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
    ' layer prev display on shades on symbols ' +
    'off lines off count on title "每线话务量"  Font ("Arial",0,12,0) subtitle auto ' +
    'Font ("Arial",0,11,0) ascending off ranges Font ("Arial",0,11,0) auto ' +
    'display off ,auto display on ,auto display on ,auto display on ,auto ' +
    'display on');

  oleMapInfo.do('Set Map Layer erpac_shade Label Position Above Font ("Arial",1,10,0) With erpac Auto On Visibility Zoom (0, 6) Units "km"');
  gTchTraffic := True;
  oleMapInfo.do('select cell_id  from tch_file into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
end;


procedure TfmBscMain.CreateRegion_4(X, Y, L, P, Bearing, Rate : Real; V : Variant; Flag : Integer; LR : String);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2, X3, Y3, X4, Y4, Xm,Ym : Real;
begin
  X1 := X + L * Sin((Bearing - 0.5 * P) * DEG_2_Rad);
  Y1 := Y + L * Cos((Bearing - 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  X2 := X + L * Sin((Bearing - P / 6) * DEG_2_Rad);
  Y2 := Y + L * Cos((Bearing - P / 6) * DEG_2_Rad) * (103.1 / 111.2);
  X3 := X + L * Sin((Bearing + P / 6) * DEG_2_Rad);
  Y3 := Y + L * Cos((Bearing + P / 6) * DEG_2_Rad) * (103.1 / 111.2);
  X4 := X + L * Sin((Bearing + 0.5 * P) * DEG_2_Rad);
  Y4 := Y + L * Cos((Bearing + 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  Xm := (X2 + X3)/2;
  Ym := (Y2 + Y3)/2;

  //V.do('set style pen makepen(1,2, rgb(' +
  //               FloatToStr(255 * Rate) + ',' +  FloatToStr(255 * (1 - Rate))
  //               + ', 0))');
 
  if Flag = 1 then
  begin
    if LR = 'L' then
      V.do('create region  1 4 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(Xm) + ',' + FloatToStr(Ym) + ')')
    else
      V.do('create region  1 4 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(Xm) + ',' + FloatToStr(Ym) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')');
  end
  else
  begin
    if LR = 'L' then
      V.do('create region into variable TmpObject 1 4 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(Xm) + ',' + FloatToStr(Ym) + ')')
    else
      V.do('create region into variable TmpObject 1 4 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(Xm) + ',' + FloatToStr(Ym) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')');
  end;
end;


procedure TfmBscMain.UpdateASObject(X, Y, L, P, Bearing, R1, R2 : Real; V : Variant; RowId, Flag : Integer; TableName : String);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2 : real;
begin
  if Flag = 1 then
  begin
    X1 := X +  L * Sin((Bearing - 0.5 * P) * DEG_2_Rad);
    Y1 := Y +  L * Cos((Bearing - 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  end
  else
  begin
    if Flag = 2 then
    begin
      X1 := X + L * Sin((Bearing + 0.5 * P) * DEG_2_Rad);
      Y1 := Y + L * Cos((Bearing + 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
    end;
  end;
  V.do('update ' +  TableName +' set obj = createcircle (' + FloatToStr(X1) + ',' + FloatToStr(Y1) + ','
                  + FloatToStr(R1) + ') where rowid = ' + IntToStr(rowid));
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,0),rgb(0,255,0))');
  V.do('update ' +  TableName +' set obj = OverlayNodes(OBJ, createcircle (' +
                   FloatToStr(X1) + ',' + FloatToStr(Y1) + ',' +
                   FloatToStr(R2) + '))  where rowid = ' + IntToStr(rowid));
end;


procedure TfmBscMain.CreateCircle(X, Y, L, P, Bearing, Rate, R : Real; V : Variant; Flag : Integer);
begin


end;


procedure TfmBscMain.UpdateCellObject;
var
  i, wRow : Integer;
  wLon, wLat, wBearing, wRate : real;
begin
  //oleMapInfo.do('Set Map Layer cell Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cell,8)');
 // oleMapInfo.do('fetch first from cell');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cell');
    wLon := oleMapInfo.eval('cell.lon');
    wLat := oleMapInfo.eval('cell.lat');
    wBearing := oleMapInfo.eval('Cell.Bearing');
    wRate := 0;
    if Pos('5', oleMapInfo.eval('cell.bs_no')) > 0 then
      CreateRegion_5(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, 2)
    else
      CreateRegion_5(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2);
    if Pos('0', oleMapInfo.eval('cell.bs_no')) > 0 then
      oleMapInfo.do('TmpObject =  Combine(TmpObject, CreateCircle('
        + FloatToStr(wLon) + ',' + FloatToStr(wLat) + ',' + FloatToStr(20 * gCellLength)
        + '))');

    oleMapInfo.do('Update Cell set Obj = TmpObject where rowId = ' + IntToStr(i));
  end;
  oleMapInfo.do('commit table cell');
end;

procedure TfmBscMain.UpDateCircle(X, Y, L, P, Bearing, Rate : Real; V : Variant; RowId, Flag : Integer; TableName : String);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2 : real;
begin
  if Flag = 1 then
  begin
    X1 := X +  L * Sin((Bearing - 0.5 * P) * DEG_2_Rad);
    Y1 := Y +  L * Cos((Bearing - 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  end
  else
  begin
    if Flag = 2 then
    begin
      X1 := X + L * Sin((Bearing + 0.5 * P) * DEG_2_Rad);
      Y1 := Y + L * Cos((Bearing + 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
    end;
  end;
  {if Rate = 0 then
    oleMapInfo.do('Set Style pen makepen(1 ,59, RGB(0,0,0))')
  else
  begin
    if FlunkRate <= 0.05 then
      oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(0,0,255))')
    else
      if FlunkRate <= 0.1 then
        oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(255,0,255))')
      else
        oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(255,0,0))');
  end;   }
  V.do('update ' +  TableName +' set obj = createcircle (' + FloatToStr(X1) + ',' + FloatToStr(Y1) + ','
                  + FloatToStr(Rate) + ') where rowid = ' + IntToStr(rowid));
end;


procedure TfmBscMain.UpDateLine(Lon_1, Lat_1, Lon_2, Lat_2, L, P, Bearing_1, Bearing_2, Rate, FlunkRate : Real; V : Variant; rowid: Integer; TableName, CellSe, CellTa : string; LostRate, Length_Rate  : Real );
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2 : real;
begin
  if Pos('5', CellSe) > 0 then
  begin
    X1 := Lon_1 + L / 2 * Sin((Bearing_1) * DEG_2_Rad);
    Y1 := Lat_1 + L / 2 * Cos((Bearing_1) * DEG_2_Rad) * (103.1 / 111.2);
  end
  else
  begin
    X1 := Lon_1 + L * Sin((Bearing_1) * DEG_2_Rad);
    Y1 := Lat_1 + L * Cos((Bearing_1) * DEG_2_Rad) * (103.1 / 111.2);
  end;
  if Pos('5', CellTa) > 0 then
  begin
    X2 := Lon_2 + L / 2 * Sin((Bearing_2) * DEG_2_Rad);
    Y2 := Lat_2 + L / 2 * Cos((Bearing_2) * DEG_2_Rad) * (103.1 / 111.2);
  end
  else
  begin
    X2 := Lon_2 + L * Sin((Bearing_2) * DEG_2_Rad);
    Y2 := Lat_2 + L * Cos((Bearing_2) * DEG_2_Rad) * (103.1 / 111.2);
  end;

  X2 := X1 + ( X2 - X1 ) * Length_Rate;
  Y2 := Y1 + ( Y2 - Y1 ) * Length_Rate;
 { if (Rate = 0) then
    V.do('Set Style pen makepen(2 ,59, RGB(0,0,0))');
  if (Rate > 0) and (LostRate > 0.05 ) then
    V.do('Set Style pen makepen(' + FloatToStr(2 + 4 * Rate) + ',59, RGB(255,0,0))');
  if (Rate > 0) and (LostRate <= 0.05 ) and (FlunkRate > 0.10) then
    V.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(255,0,255))');
  if (Rate > 0) and (LostRate <= 0.05 ) and (FlunkRate <= 0.10) then
    V.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(0,0,255))');
  if (Rate > 0) and (LostRate <= 0.05 ) and (FlunkRate = 0) then
    V.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',59, RGB(0,0,255))'); }
  if (Rate = 0) then
  begin

    V.do('Set Style pen makepen(2 ,2, RGB(0,0,0))');

  end
  else
  begin
    if (LostRate > 0.05)  then
    begin
       V.do('Set Style pen makepen(' + FloatToStr(2 + 4 * Rate) + ',2, RGB(255,0,0))');
       if (Rate = 0) then
          V.do('Set Style pen makepen(2 ,2, RGB(0,0,0))');
    end
    else
    begin
      if (FlunkRate < 0.10)  then
      begin
        v.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',2, RGB(0,0,255))');
      end
      else
      begin
        v.do('Set Style pen makepen(' + FloatToStr(1 + 5 * Rate) + ',2, RGB(255,0,255))');
      end;
    end;
  end; 
  V.do('update ' + TableName + ' set obj = createLine (' + FloatToStr(X1) + ',' + FloatToStr(Y1)
       + ',' +  FloatToStr(X2) + ',' + FloatToStr(Y2)  + ') where rowid = ' + IntToStr(rowid));
end;

procedure TfmBscMain.CreateRegion_3(X, Y, L, D, Bearing, Rate : Real; V : Variant);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X0, Y0, X1, Y1, X2, Y2, X3, Y3 : Real;
begin
  Rate := - Rate;
  if Rate < -2 then
    Rate := -2;
  if Rate > 2 then
    Rate := 2;
  X0 := X + 0.5 * L * Sin(Bearing * DEG_2_RAD);
  Y0 := Y + 0.5 * L * Cos(Bearing * DEG_2_RAD);
  X1 := X0 - 0.5 * D * Rate;
  X2 := X0 + 0.5 * D * Rate;
  X3 := X0 ;
  Y1 := Y0;
  Y2 := Y0;
  Y3 := Y0 - Sin(gCellAngle * DEG_2_RAD) * D * Rate * (103.1 / 111.2);

  V.do('create region into variable TmpObject 1 3 ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ')');

end;

procedure TfmBscMain.CreateRegion_5(X, Y, L, P, Bearing, Rate : Real; V : Variant; Flag : Integer);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2, X3, Y3, X4, Y4 : Real;
begin
  X1 := X + L * Sin((Bearing - 0.5 * P) * DEG_2_Rad);
  Y1 := Y + L * Cos((Bearing - 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  X2 := X + L * Sin((Bearing - P / 6) * DEG_2_Rad);
  Y2 := Y + L * Cos((Bearing - P / 6) * DEG_2_Rad) * (103.1 / 111.2);
  X3 := X + L * Sin((Bearing + P / 6) * DEG_2_Rad);
  Y3 := Y + L * Cos((Bearing + P / 6) * DEG_2_Rad) * (103.1 / 111.2);
  X4 := X + L * Sin((Bearing + 0.5 * P) * DEG_2_Rad);
  Y4 := Y + L * Cos((Bearing + 0.5 * P) * DEG_2_Rad) * (103.1 / 111.2);
  V.do('set style pen makepen(1,2, rgb(' +
                 FloatToStr(255 * Rate) + ',' +  FloatToStr(255 * (1 - Rate))
                 + ', 0))');
  V.do('set style brush makebrush(82,rgb(' +
                 FloatToStr(255 * Rate) + ',' +  FloatToStr(255 * (1 - Rate))
                 + ', 0), rgb(255,255,255))');
  if Flag = 1 then
     V.do('create region  1 5 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')')
  else
  begin
    {if L = 0.5 * gCellLength then
    begin

      V.do('set style pen makepen(3, 2, rgb(' +
                 FloatToStr(255 * Rate) + ',' +  FloatToStr(255 * (1 - Rate))
                 + ', 0))');
      V.do('create pline into variable TmpObject 6 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ') ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ')');
      V.do('TmpObject = Combine(TmpObject, CreateLine('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ','
                    + FloatToStr(2 * X1 - X) + ',' + FloatToStr(2 * Y1 - Y) + '))');
      V.do('TmpObject = Combine(TmpObject, CreateLine('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ','
                    + FloatToStr(2 * X4 - X) + ',' + FloatToStr(2 * Y4 - Y) + '))');
    end
    else}
      V.do('create region into variable TmpObject 1 5 ('
                    + FloatToStr(X) + ',' + FloatToStr(Y) + ') ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')');

  end;
end;


procedure TfmBscMain.CreateMDIChild(const Name: string);
var
  Child: TfmMap;
begin
  { create a new MDI child window }
  Child := TfmMap.Create(Application);
  Child.Caption := Name;
  //if FileExists(Name) then Child.Memo1.Lines.LoadFromFile(Name);
end;

procedure TfmBscMain.FileNew1Execute(Sender: TObject);
begin
  //CreateMDIChild('NONAME' + IntToStr(MDIChildCount + 1));
end;

//procedure TfmBscMain.FileOpen1Execute(Sender: TObject);



procedure TfmBscMain.HelpAbout1Execute(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TfmBscMain.FileExit1Execute(Sender: TObject);
begin
  Close;
end;

procedure TfmBscMain.FormActivate(Sender: TObject);
var
  sWinHandle : String;
begin

  str(Handle, sWinHandle);
  oleMapInfo.Do('Set Application Window ' + sWinHandle);
  oleMapInfo.Do('Set Window Info Parent ' +  sWinHandle);
  oleMapInfo.Do('Set Window ruler Parent ' +  sWinHandle);
  oleMapInfo.Do('Set Window Legend Parent ' +  sWinHandle);
  oleMapInfo.Do('Set Window message Parent ' +  sWinHandle);
end;

procedure TfmBscMain.FormCreate(Sender: TObject);
var
  i : Integer;
begin
  //gMultiCell := TStrings.Create;
  wWirClass := CoWClass.Create;
  
  oleMapInfo := CreateOleObject('MapInfo.Application');
  oleMapInfo.do('Dim TmpObject as Object');
  oleMapInfo.do('Dim MyObj as Object');
  oleMapInfo.do('Dim EmptyObject as Object');

  with dmBscData.dbBscData do
  begin
    if not Connected then
      Connected := True;
  end;
  with dmBscData.tbBscControl do
  begin
    if not Active then
      Open;
    First;
    gCellAngle := FieldByName('cell_angle').AsFloat;
    gCellLength := FieldByName('cell_Length').AsFloat;
    if FieldByName('Control_key').AsString = '0' then
    begin
      mmDataConv.Enabled := False;
      mmUpdateData.Enabled := False;
      mmOpenMap.Enabled := False;
      mmMapConf.Enabled := False;
    end;
    Close;
  end;
  gCompLayer := '';
  gDr := False;
  gHover := False;
  gRaf := False;
  gTraffic := False;
  gChAc := False;
  gSaSs := False;
  gSu := False;
  gDqa := False;
  gTchCG := False;
  gTchTraffic := False;
  
end;

procedure TfmBscMain.SpeedButton3Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1705);
end;

procedure TfmBscMain.SpeedButton14Click(Sender: TObject);
begin
 
  oleMapInfo.RunMenuCommand(1706);
end;

procedure TfmBscMain.SpeedButton1Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1701);
end;

procedure TfmBscMain.SpeedButton17Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1710);
end;

procedure TfmBscMain.SpeedButton16Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1707);
end;

procedure TfmBscMain.SpeedButton10Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(309);
end;

procedure TfmBscMain.SpeedButton11Click(Sender: TObject);
var
  i : Integer;
begin
  {for i := 1 to 10 do
    oleMapInfo.do('set map layer ' + IntToStr(i) + ' selectable off');
  oleMapInfo.do('set map layer cell selectable on');}

  oleMapInfo.RunMenuCommand(1703);
end;

procedure TfmBscMain.SpeedButton12Click(Sender: TObject);
var
  i : Integer;
begin
{  for i := 1 to 10 do
    oleMapInfo.do('set map layer ' + IntToStr(i) + ' selectable off');
  oleMapInfo.do('set map layer cell selectable on'); }

  oleMapInfo.RunMenuCommand(1722);
end;

procedure TfmBscMain.SpeedButton13Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(801);
end;

procedure TfmBscMain.SpeedButton4Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1711);
end;

procedure TfmBscMain.SpeedButton2Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1702);
end;

procedure TfmBscMain.SpeedButton15Click(Sender: TObject);
begin
  Application.CreateForm(TfmFind, fmFind);
  try
    fmFind.ShowModal;
  finally
    fmFind.Free;
  end;
end;

procedure TfmBscMain.SpeedButton18Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1708);
end;

procedure TfmBscMain.SpeedButton9Click(Sender: TObject);
begin
  // mapinfo.runmenucommand 606
 // if gRaf or gMh or gTchTraffic then
    oleMapInfo.RunMenuCommand(606)
  //else
  //  if not fmLegeng.Showing then
   //   fmLegeng.Show;
end;

procedure TfmBscMain.mmTCHClick(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(102);
end;

procedure TfmBscMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  oleMapInfo.do('unDim TmpObject ');
  oleMapInfo.do('unDim EmptyObject ');
  oleMapInfo.do('unDim MyObj ');
  oleMapInfo := Unassigned;
end;

procedure TfmBscMain.N21Click(Sender: TObject);
begin
  oleMapInfo.do('Set Map Layer 0 Editable On');
end;

procedure TfmBscMain.N25Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(108);
end;

procedure TfmBscMain.N26Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(109);
end;

procedure TfmBscMain.mmDrClick(Sender: TObject);
var
  i, wRow, wNcellRow, wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty : real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gDr then
  begin
    if mmDr.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_sdr Display Off');
      oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_sdr Display Graphic');
      oleMapInfo.do('Set Map Layer Tch_dr Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmDr.Checked := not mmDr.Checked;
    exit;
  end;
  mmDr.Checked := not mmDr.Checked;
  oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';

  if  (gSelFlag <> 'CELL') then
    ShowCondition('掉话率')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.sdr ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into SdrTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into SdrTmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.sdr '
        else
          wCondition := ' order by cch_file.sdr desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into SdrTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0"  into sdrTmp');
    end;
   { oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" into SdrTmp');}
  end
  else  //cell
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into SdrTmp')
      else
      begin
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into SdrTmp');
      end;
    end
    else //all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.sdr ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition
                 + ' into SdrTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "" into SdrTmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.sdr '
          else
            wCondition := ' order by cch_file.sdr desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' 
                 + wCondition + ' into SdrTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "" into SdrTmp');
      end;

      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into SdrTmp');}
    end;
  end;

  oleMapInfo.do('commit table SdrTmp as "' + gExePath + 'cch_sdr.tab"');
  oleMapInfo.do('close table SdrTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_sdr.tab" Interactive');

  oleMapInfo.do('select Max(sdr) from cch_sdr into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Sdr, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_sdr');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_sdr where rowid = ' + IntToStr(i));
    end
    else 
    begin
      wLon := oleMapInfo.eval('cch_sdr.lon');
      wLat := oleMapInfo.eval('cch_sdr.lat');
      wBearing := oleMapInfo.eval('cch_sdr.Bearing');
      wRate := oleMapInfo.eval('cch_sdr.sdr') / wMaxQty;
      {oleMapInfo.do('set style brush makebrush(82,rgb(0,' + FloatToStr(255 * (1 - wRate)) + ','
                + FloatToStr(255 * (1 - wRate)) +'), rgb(255,255,255))');
      CreateRegion_4(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2, 'L');
      oleMapInfo.do('Update cch_sdr set obj = TmpObject where rowid = ' + IntToStr(i));
      }
      if wRate > 0 then
        wRate := 0.005 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_sdr.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_sdr')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_sdr');
    end;
  end;
  oleMapInfo.do('commit table cch_sdr');
  oleMapInfo.do('add map auto layer cch_sdr');
  oleMapInfo.do('Set Map Layer cch_sdr Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With sdr+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_sdr');
  wFristCchTable := 'CCH_SDR';
  wPriorCchTable := 'CCH_SDR';

  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and Tch_file.dr ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into drTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "" into drTmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by Tch_file.dr '
        else
          wCondition := ' order by Tch_file.dr desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into drTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "" into drTmp');
    end;

    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into drTmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into drTmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into drTmp');
    end
    else//all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.dr ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into drTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "" into drTmp');
        end
        else
        //order by
        begin
          if wTchCheck = 'Y' then
          begin
            if wTchFlag = '升序' then
              wCondition := ' order by Tch_file.dr '
            else
              wCondition := ' order by Tch_file.dr desc ' ;
            oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into drTmp');
          end
          else
            oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "" into drTmp');
      end;
      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into drTmp');}
    end;
  end;

  oleMapInfo.do('commit table drTmp as "' + gExePath + 'tch_dr.tab"');
  oleMapInfo.do('close table drTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_dr.tab" Interactive');

  oleMapInfo.do('select Max(dr) from tch_dr into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_dr, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_dr');
    if (wFilterOrder = 'ORDER') and (i > wCondNum) then
    begin
      oleMapInfo.do('Delete from Tch_dr  where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_dr.lon');
      wLat := oleMapInfo.eval('tch_dr.lat');
      wBearing := oleMapInfo.eval('tch_dr.Bearing');
      wRate := oleMapInfo.eval('tch_dr.dr') / wMaxQty;

      {oleMapInfo.do('set style brush makebrush(82,rgb(' +
                 FloatToStr(255 * (1- wRate)) +
                 ',' + FloatToStr(255 * (1- wRate)) +
                 ',0), rgb(255,255,255))');
      CreateRegion_4(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2, 'R');
      oleMapInfo.do('Update Tch_dr set obj = TmpObject where rowid = ' + IntToStr(i)); }

      if wRate > 0 then
        wRate := 0.005 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('Tch_dr.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_dr')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_dr');
    end;
  end;
  oleMapInfo.do('commit table tch_dr');
  oleMapInfo.do('add map auto layer tch_dr');
  oleMapInfo.do('Set Map Layer tch_dr Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With dr+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  //wFristTchTable := 'TCH_DR';
  //wPriorTchTable := 'TCH_DR';

  oleMapInfo.do('set map redraw on');
  gDr := True;
  //sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;

  oleMapInfo.do('select  cell_id from tch_dr into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_sdr into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');

  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 1;
end;

procedure TfmBscMain.mmSaSsClick(Sender: TObject);
var
  i, wRow , wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wPenWidth : real;
begin
 // fmBscMain.ResetMap;
 // fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gSaSs then
  begin
    if mmSaSs.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_SA Display Off');
      oleMapInfo.do('Set Map Layer tch_CA Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_SA Display Graphic');
      oleMapInfo.do('Set Map Layer tch_Ca Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmSaSs.Checked := not mmSaSs.Checked;
    exit;
  end;
  mmSaSs.Checked := not mmSaSs.Checked;
  if  (gSelFlag <> 'CELL') then
    ShowCondition('申请失败率')
  else
    wFilterOrder := '';
  oleMapInfo.do('set map redraw off');
   sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  oleMapInfo.do('Alter Table "cch_File" ( Add Ss_Fail_Rate Decimal(5,2) ) Interactive');
  oleMapInfo.do('Update cch_file Set Ss_Fail_Rate = 100*(Sa-Ss)/( sa + 0.0001)');
  oleMapInfo.do('Commit table cch_file');
  //oleMapInfo.do('Update cch_file Set Ss_Fail_Rate = 0 where sa = 0');


  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.Ss_Fail_Rate ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id, cch_file.SA, cch_file.Ss, Ss_Fail_Rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into SATmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id, cch_file.SA, cch_file.Ss, Ss_Fail_Rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into SATmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.ss_fail_rate '
        else
          wCondition := ' order by cch_file.ss_fail_rate desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id, cch_file.SA, cch_file.Ss, Ss_Fail_Rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into SATmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id, cch_file.SA, cch_file.Ss, Ss_Fail_Rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into SATmp');
    end;

    {oleMapInfo.do('Select CCh_file.Cell_id, cch_file.SA, cch_file.Ss, Ss_Fail_Rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into SATmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into SATmp')
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into SATmp')
    end
    else //all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.sdr ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, cch_file.ss_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into SATmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, cch_file.ss_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into SATmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.ss_fail_rate '
          else
            wCondition := ' order by cch_file.ss_fail_rate desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, cch_file.ss_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into SATmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, cch_file.ss_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into SATmp');
      end;
      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into SATmp');}
    end;
  end;

  oleMapInfo.do('commit table SATmp as "' + gExePath + 'cch_SA.tab"');
  oleMapInfo.do('close table SATmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_SA.tab" Interactive');

  oleMapInfo.do('Alter Table "cch_file" ( drop Ss_Fail_Rate ) Interactive');
  oleMapInfo.do('commit table cch_file');

  oleMapInfo.do('select Max(sa) from cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(ca) from Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_SA, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);

  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_SA');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_sa where rowid = ' + IntToStr(i));
    end
    else
    begin

      wLon := oleMapInfo.eval('cch_sA.lon');
      wLat := oleMapInfo.eval('cch_sA.lat');
      wBearing := oleMapInfo.eval('cch_SA.Bearing');
      wRate := oleMapInfo.eval('cch_SA.SA') / wMaxQty;
      if wRate > 0 then
        wRate := 0.01 + 0.12 * wRate;
      if oleMapInfo.eval('cch_sa.Sa') <> 0 then
        wPenWidth := 1 +  6 *(oleMapInfo.eval('cch_sa.Sa') - oleMapInfo.eval('cch_sa.Ss')) /
                     oleMapInfo.eval('cch_sa.Sa')
      else
        wPenWidth := 1;
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    //if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_sa.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_sa')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_sa');
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 2, 'cch_sa');
    end;
  end;
  oleMapInfo.do('commit table cch_SA');
  oleMapInfo.do('add map auto layer cch_SA');
  oleMapInfo.do('Set Map Layer cch_SA Label Position Above Font ("Arial",0,10,0) With sa+"("+Ss+")" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_sa');
  {//////////////////////////////////////////////////////////////////////
  oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Ss,cch_file.SA, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into SsTmp');
  oleMapInfo.do('commit table SSTmp as "' + gExePath + 'cch_Ss.tab"');
  oleMapInfo.do('close table SsTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_Ss.tab" Interactive');
  oleMapInfo.do('add map auto layer cch_Ss');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
 // oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,255),rgb(255,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Ss, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_Ss');
    wLon := oleMapInfo.eval('cch_ss.lon');
    wLat := oleMapInfo.eval('cch_ss.lat');
    wBearing := oleMapInfo.eval('cch_Ss.Bearing');
    wRate := oleMapInfo.eval('cch_Ss.Ss') /wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.10 * wRate;
    wPenWidth := 1 +  6 *(oleMapInfo.eval('cch_Traffic.Sa') - oleMapInfo.eval('cch_Traffic.Ss')) /
                     oleMapInfo.eval('cch_Traffic.Sa');

    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 2, 'cch_ss');
  end;
  oleMapInfo.do('commit table cch_Ss');
  oleMapInfo.do('Set Map Layer cch_Ss Label Position Above Font ("Arial",0,10,0) With sa+"("+Ss+")" Auto On Visibility Zoom (0, 6) Units "km"');
  }
  //////////////////////////////////////////////Tch
  oleMapInfo.do('Alter Table "Tch_File" ( Add ca_Fail_Rate Decimal(5,2) ) Interactive');
  oleMapInfo.do('Update Tch_file Set ca_Fail_Rate = 100*(ca - cs)/( ca + 0.0001)');
  oleMapInfo.do('Commit table Tch_file');
  if gSelFlag = 'BSC' then
  begin
    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into cATmp'); }
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.ca_fail_rate ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into cATmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs,tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into cATmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.ca_fail_rate '
        else
          wCondition := ' order by tch_file.ca_fail_rate desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs,tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into cATmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs,tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into cATmp');
    end;

  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs,tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into cATmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs,tch_file.ca_Fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into cATmp')
    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.ca_fail_rate ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, tch_fail.ca_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into cATmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, tch_fail.ca_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into cATmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.ca_fail_rate '
          else
            wCondition := ' order by tch_file.ca_fail_rate desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, tch_file.ca_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into cATmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, tch_fail.ca_fail_rate, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into cATmp');
      end;

      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into cATmp'); }
    end;
  end;

  oleMapInfo.do('commit table cATmp as "' + gExePath + 'tch_cA.tab"');
  oleMapInfo.do('close table cATmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_cA.tab" Interactive');
  oleMapInfo.do('Alter Table "Tch_file" ( drop ca_Fail_Rate ) Interactive');

  oleMapInfo.do('Commit table Tch_file');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_cA, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_cA');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from Tch_ca where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_ca.lon');
      wLat := oleMapInfo.eval('tch_ca.lat');
      wBearing := oleMapInfo.eval('tch_ca.Bearing');
      wRate := oleMapInfo.eval('tch_cA.ca')/ wMaxQty;
      if wRate > 0 then
        wRate := 0.01 + 0.12 * wRate;
      if oleMapInfo.eval('tch_ca.ca') <> 0 then
        wPenWidth := 1 +  6 *(oleMapInfo.eval('tch_ca.ca') - oleMapInfo.eval('tch_ca.cs')) /
                     oleMapInfo.eval('tch_ca.ca')
      else
        wPenWidth := 1;
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    //if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('tch_ca.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_ca')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_ca');
    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_ca');
    end;
  end;
  oleMapInfo.do('commit table Tch_cA');
  oleMapInfo.do('add map auto layer tch_cA');
  oleMapInfo.do('Set Map Layer Tch_cA Label Position Above Font ("Arial",0,10,0) With ca+"("+cs+")" Auto On Visibility Zoom (0, 6) Units "km"');
  //////////////////////////////////////////////////////////////////////
  {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cs,Tch_file.cA, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into csTmp');
  oleMapInfo.do('commit table cSTmp as "' + gExePath + 'tch_cs.tab"');
  oleMapInfo.do('close table csTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_cs.tab" Interactive');
  oleMapInfo.do('add map auto layer tch_cs');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
 // oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,255),rgb(255,255,255))');
  wRow := oleMapInfo.eval('tableinfo(tch_cs, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_cs');
    wLon := oleMapInfo.eval('tch_cs.lon');
    wLat := oleMapInfo.eval('tch_cs.lat');
    wBearing := oleMapInfo.eval('tch_cs.Bearing');
    wRate := oleMapInfo.eval('tch_cs.cs') /wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.10 * wRate;
    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_cs');
  end;
  oleMapInfo.do('commit table tch_cs');
  oleMapInfo.do('Set Map Layer tch_cs Label Position Above Font ("Arial",0,10,0) With ca+"("+cs+")" Auto On Visibility Zoom (0, 6) Units "km"');
  }

  oleMapInfo.do('set map redraw on');
  gSaSs := True;
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  oleMapInfo.do('select  cell_id from tch_ca into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_sa into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 4;
end;


procedure TfmBscMain.mmSuClick(Sender: TObject);
var
  i, wRow, wNcellRow, wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wMinQty : real;
begin
  //fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gSu then
  begin
    if mmSu.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_Su Display Off');
      oleMapInfo.do('Set Map Layer tch_U Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_su Display Graphic');
      oleMapInfo.do('Set Map Layer tch_u Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmSu.Checked := not mmSu.Checked;
    exit;
  end;
  mmSu.Checked := not mmSu.Checked;
  oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  if  (gSelFlag <> 'CELL') then
    ShowCondition('接通率')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.su ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into suTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into SuTmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.su '
        else
          wCondition := ' order by cch_file.su desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into SuTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0"  into suTmp');
    end;


    {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into suTmp');  }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into suTmp')
      else
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into suTmp');
    end
    else
    begin //all
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.sdr ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into SuTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into SuTmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.su '
          else
            wCondition := ' order by cch_file.su desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into SuTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0"  into suTmp');
      end;


      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into suTmp'); }
    end;
  end;

  oleMapInfo.do('commit table SuTmp as "' + gExePath + 'cch_su.tab"');
  oleMapInfo.do('close table SuTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_su.tab" Interactive');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Su, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);

  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_su');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_su where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('cch_su.lon');
      wLat := oleMapInfo.eval('cch_su.lat');
      wBearing := oleMapInfo.eval('cch_su.Bearing');
      wRate := oleMapInfo.eval('cch_su.su') / 100;
      if wRate > 0 then
        wRate := 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_su.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength / 2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_su')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_su');
    end;
  end;
  oleMapInfo.do('commit table cch_su');
  oleMapInfo.do('add map auto layer cch_su');
  oleMapInfo.do('Set Map Layer cch_su Label Position Above Font ("Arial",0,10,0) With Su+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_su');
  ////////////////////////////////////////////tch
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.u ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select tCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into uTmp');
      end
      else
        oleMapInfo.do('Select tCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into uTmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.u '
        else
          wCondition := ' order by tch_file.u desc ' ;
        oleMapInfo.do('Select tCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into uTmp');
      end
      else
        oleMapInfo.do('Select tCh_file.Cell_id,tch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0"  into uTmp');
    end;
    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into uTmp'); }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into uTmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + 'into uTmp')
    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.u ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into uTmp');
        end
        else
          oleMapInfo.do('Select tCh_file.Cell_id,tch_file.Su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into uTmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.u '
          else
            wCondition := ' order by tch_file.u desc ' ;
          oleMapInfo.do('Select tCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into uTmp');
        end
        else
          oleMapInfo.do('Select tCh_file.Cell_id,tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0"  into uTmp');
      end;


      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into uTmp'); }
    end;
  end;

  oleMapInfo.do('commit table uTmp as "' + gExePath + 'tch_u.tab"');
  oleMapInfo.do('close table uTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_u.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_u, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_u');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from Tch_u where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_u.lon');
      wLat := oleMapInfo.eval('tch_u.lat');
      wBearing := oleMapInfo.eval('tch_u.Bearing');
      wRate := oleMapInfo.eval('tch_u.u') / 100;
      if wRate > 0 then
        wRate :=  0.05 * wRate;
      if Pos('5', oleMapInfo.eval('Tch_u.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_u')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_u');
    end;
  end;
  oleMapInfo.do('commit table tch_u');
   oleMapInfo.do('add map auto layer tch_u');
  oleMapInfo.do('Set Map Layer tch_u Label Position Above Font ("Arial",0,10,0) With u+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  oleMapInfo.do('set map redraw on');
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  gSu := True;
  oleMapInfo.do('select  cell_id from tch_u into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_su into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  {with cbNext do
  begin
    Visible := True;
    Items.Clear;
    Items.Add('TCH 申请与分配');
    Items.Add('SDCCH 申请与分配');
  end;}
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 0;
end;

procedure TfmBscMain.mmChClick(Sender: TObject);
var
  i, wRow, wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty , wPenWidth: real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gChAc then
  begin
    if mmChAc.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_ch Display Off');
      oleMapInfo.do('Set Map Layer tch_ch Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_ch Display Graphic');
      oleMapInfo.do('Set Map Layer tch_ch Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmChAc.Checked := not mmChAc.Checked;
    exit;
  end;
  mmChAc.Checked := not mmChAc.Checked;
  oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  if  (gSelFlag <> 'CELL') then
    ShowCondition('信道损坏率')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.sf ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into chTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0"into chTmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.sf '
        else
          wCondition := ' order by cch_file.sf desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into chTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = ""into chTmp');
    end;
    {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into chTmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into chTmp')
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into chTmp')
    end
    else //all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.sf ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition +' into chTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into chTmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.sf '
          else
            wCondition := ' order by cch_file.sf desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into chTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into chTmp');
      end;
      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into chTmp');}
    end;
  end;


  oleMapInfo.do('commit table chTmp as "' + gExePath + 'cch_ch.tab"');
  oleMapInfo.do('close table chTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_ch.tab" Interactive');

  oleMapInfo.do('select Max(ch) from cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(ch) from Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_ch, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_ch');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_ch where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('cch_ch.lon');
      wLat := oleMapInfo.eval('cch_ch.lat');
      wBearing := oleMapInfo.eval('cch_ch.Bearing');
      wRate := oleMapInfo.eval('cch_ch.ch')/ wMaxQty;
      if wRate > 0 then
        wRate := 0.005 + 0.012 * wRate;
      wPenWidth := 1 + 6 * oleMapInfo.eval('cch_ch.sf') / 100;
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
   // if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_ch.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_ch')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_ch');

    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_ch');
    end;
   end;
  oleMapInfo.do('commit table cch_ch');
  oleMapInfo.do('add map auto layer cch_ch');
  oleMapInfo.do('Set Map Layer cch_ch Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+sf+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_ch');
  //////////////////////////////////////////////////////////////////////
 { oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ac,cch_file.ch,cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into acTmp');
  oleMapInfo.do('commit table acTmp as "' + gExePath + 'cch_ac.tab"');
  oleMapInfo.do('close table acTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_ac.tab" Interactive');
  oleMapInfo.do('add map auto layer cch_ac');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
 // oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,255),rgb(255,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_ac, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_ac');
    wLon := oleMapInfo.eval('cch_ac.lon');
    wLat := oleMapInfo.eval('cch_ac.lat');
    wBearing := oleMapInfo.eval('cch_ac.Bearing');
    wRate := oleMapInfo.eval('cch_ac.ac') /wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.10 * wRate;
    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_ac');
  end;
  oleMapInfo.do('commit table cch_ac');
  oleMapInfo.do('Set Map Layer cch_ac Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+sf+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
   }
  //////////////////////////////////////////////Tch
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.f ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into chTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into chTmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.f '
        else
          wCondition := ' order by tch_file.f desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into chTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into chTmp');
    end;
    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into chTmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into chTmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into chTmp')
    end
    else//all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.f ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into chTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into chTmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.f '
          else
            wCondition := ' order by tch_file.f desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into chTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into chTmp');
      end;

      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into chTmp');}
    end;
  end;

  oleMapInfo.do('commit table chTmp as "' + gExePath + 'tch_ch.tab"');
  oleMapInfo.do('close table chTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_ch.tab" Interactive');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_ch, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_ch');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from Tch_ch where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_ch.lon');
      wLat := oleMapInfo.eval('tch_ch.lat');
      wBearing := oleMapInfo.eval('tch_ch.Bearing');
      wRate := oleMapInfo.eval('tch_ch.ch')/ wMaxQty;
      if wRate > 0 then
        wRate := 0.005 + 0.012 * wRate;
      wPenWidth := 1 + 6 * oleMapInfo.eval('tch_ch.f');
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
   // if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('tch_ch.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_ch')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_ch');
    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_ch');
    end;
  end;
  oleMapInfo.do('commit table Tch_ch');
  oleMapInfo.do('add map auto layer tch_ch');

  oleMapInfo.do('Set Map Layer Tch_ch Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+f+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
  //////////////////////////////////////////////////////////////////////
  {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ac,Tch_file.ch,Tch_file.F, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into acTmp');
  oleMapInfo.do('commit table acTmp as "' + gExePath + 'tch_ac.tab"');
  oleMapInfo.do('close table acTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_ac.tab" Interactive');
  oleMapInfo.do('add map auto layer tch_ac');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
 // oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,255),rgb(255,255,255))');
  wRow := oleMapInfo.eval('tableinfo(tch_ac, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_ac');
    wLon := oleMapInfo.eval('tch_ac.lon');
    wLat := oleMapInfo.eval('tch_ac.lat');
    wBearing := oleMapInfo.eval('tch_ac.Bearing');
    wRate := oleMapInfo.eval('tch_ac.ac') /wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.10 * wRate;
    //UpdateASObject(wLon, wLat, 0.0025, gCellAngle, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_ac');
  end;
  oleMapInfo.do('commit table tch_ac');
  oleMapInfo.do('Set Map Layer tch_ac Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+f+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
    }
  oleMapInfo.do('set map redraw on');
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  gChAc := True;
  oleMapInfo.do('select  cell_id from tch_ch into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_ch into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 5;
end;

procedure TfmBscMain.mmTrafficClick(Sender: TObject);
var
  i, wRow, wNcellRow, wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wPenWidth : real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gTraffic then
  begin
    if mmTraffic.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_traffic Display Off');
      oleMapInfo.do('Set Map Layer Tch_traffic Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_traffic Display Graphic');
      oleMapInfo.do('Set Map Layer tch_traffic Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmTraffic.Checked := not mmTraffic.Checked;
    exit;
  end;
  mmTraffic.Checked := not mmTraffic.Checked;
  oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  if  (gSelFlag <> 'CELL') then
    ShowCondition('拥塞率')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.sc ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into TrafficTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TrafficTmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.sc '
        else
          wCondition := ' order by cch_file.sc desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition +' into TrafficTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TrafficTmp');
    end;

    {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into TrafficTmp'); }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into TrafficTmp')
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into TrafficTmp')
    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.sc ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition +' into TrafficTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into TrafficTmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.sc '
          else
            wCondition := ' order by cch_file.sc desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition  + ' into TrafficTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into TrafficTmp');
      end;
      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into TrafficTmp');}
    end;
  end;

  oleMapInfo.do('commit table TrafficTmp as "' + gExePath + 'cch_Traffic.tab"');
  oleMapInfo.do('close table TrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_Traffic.tab" Interactive');

  oleMapInfo.do('select Max(Traffic) from cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(Traffic) from Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Traffic, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_Traffic');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_Traffic where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('cch_Traffic.lon');
      wLat := oleMapInfo.eval('cch_Traffic.lat');
      wBearing := oleMapInfo.eval('cch_Traffic.Bearing');
      wRate := oleMapInfo.eval('cch_Traffic.Traffic') / wMaxQty;
      wPenWidth := 1 + oleMapInfo.eval('cch_Traffic.Sc');
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
      if wRate > 0 then
        wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_Traffic.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_Traffic')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_Traffic');
    end;
  end;
  oleMapInfo.do('commit table cch_Traffic');
  oleMapInfo.do('add map auto layer cch_Traffic');
  oleMapInfo.do('Set Map Layer cch_Traffic Label Position Above Font ("Arial",0,10,0) With Traffic+","+sc+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  if wRow > 0 then
    //oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_traffic')
  else
    ShowMessage('Sdcch没有数据符合条件!');
  ////tch
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.cg ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into TrafficTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TrafficTmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.cg '
        else
          wCondition := ' order by tch_file.cg desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"' + wCondition + ' into TrafficTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TrafficTmp');
    end;
    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into TrafficTmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into TrafficTmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into TrafficTmp')
    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.cg ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into TrafficTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into TrafficTmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.cg '
          else
            wCondition := ' order by tch_file.cg desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into TrafficTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into TrafficTmp');
      end;

      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into TrafficTmp');}
    end;
  end;

  oleMapInfo.do('commit table TrafficTmp as "' + gExePath + 'tch_Traffic.tab"');
  oleMapInfo.do('close table TrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_Traffic.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_Traffic, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_Traffic');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from tch_traffic where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_Traffic.lon');
      wLat := oleMapInfo.eval('tch_Traffic.lat');
      wBearing := oleMapInfo.eval('tch_Traffic.Bearing');
      wRate := oleMapInfo.eval('tch_Traffic.Traffic') / wMaxQty;
      wPenWidth := 1 + oleMapInfo.eval('tch_Traffic.cg');
      if wPenWidth > 7 then
        wPenWidth := 7;
      oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
      if wRate > 0 then
        wRate := 0.02 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('tch_Traffic.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_Traffic')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_Traffic');
    end;
  end;
  oleMapInfo.do('commit table tch_Traffic');
  oleMapInfo.do('add map auto layer tch_Traffic');
  oleMapInfo.do('Set Map Layer tch_Traffic Label Position Above Font ("Arial",0,10,0) With Traffic+","+cg+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  if wRow = 0 then
    ShowMessage('Tch没有数据符合条件!');
  oleMapInfo.do('set map redraw on');
  gTraffic := True;
  oleMapInfo.do('select  cell_id from tch_traffic into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_traffic into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  wPriorTchTable := 'Tch_traffic';
  wFristCchTable := 'Cch_traffic';
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 2;
end;

procedure TfmBscMain.mmDqaClick(Sender: TObject);
var
  i, wRow, wNcellRow , wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty : real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gDqa then
  begin
    if mmDqa.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_Dqa Display Off');
      oleMapInfo.do('Set Map Layer Tch_Tqa Display off');
      //oleMapInfo.do('Set Map Layer cch_Dss4 Display Off');
      //oleMapInfo.do('Set Map Layer Tch_Tss4 Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cch_Dqa Display Graphic');
      oleMapInfo.do('Set Map Layer Tch_Tqa Display Graphic');
      //oleMapInfo.do('Set Map Layer cch_Dss4 Display Graphic');
      //oleMapInfo.do('Set Map Layer Tch_Dss4 Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmDqa.Checked := not mmDqa.Checked;
    exit;
  end;
  mmDqa.Checked := not mmDqa.Checked;
  oleMapInfo.do('set map redraw off');
   sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  if  (gSelFlag <> 'CELL') then
    ShowCondition('质差断线')
  else
    wFilterOrder := '';

  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.dqa ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into DqaTmp'); 
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into DqaTmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.dqa '
        else
          wCondition := ' order by cch_file.dqa desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into DqaTmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into DqaTmp');
    end;
{
    oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into DqaTmp');  }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into DqaTmp')
      else
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into DqaTmp')


    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.dqa ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into DqaTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into DqaTmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.dqa '
          else
            wCondition := ' order by cch_file.dqa desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into DqaTmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into DqaTmp');
      end;

      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into DqaTmp');}
    end;
  end;

  oleMapInfo.do('commit table DqaTmp as "' + gExePath + 'cch_Dqa.tab"');
  oleMapInfo.do('close table DqaTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_Dqa.tab" Interactive');

  oleMapInfo.do('select Max(Dqa) from cch_Dqa into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Dqa, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);

  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_Dqa');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_dqa where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('cch_Dqa.lon');
      wLat := oleMapInfo.eval('cch_Dqa.lat');
      wBearing := oleMapInfo.eval('cch_Dqa.Bearing');
      try
        wRate := oleMapInfo.eval('cch_Dqa.Dqa') / wMaxQty;
      except
        wRate := 0;
      end;
      if wRate > 0 then
        wRate := 0.01 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_dqa.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_dqa')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_dqa');
    //oleMapInfo.do('set style brush makebrush(82,rgb(0,' + FloatToStr(255 * (1 - wRate)) + ','
     //           + FloatToStr(255 * (1 - wRate)) +'), rgb(255,255,255))');
    //CreateRegion_4(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2, 'L');
    //oleMapInfo.do('Update cch_dqa set obj = TmpObject where rowid = ' + IntToStr(i));
    end;
  end;
  oleMapInfo.do('commit table cch_Dqa');
  oleMapInfo.do('add map auto layer cch_Dqa');
  oleMapInfo.do('Set Map Layer cch_Dqa Label Position Above Font ("Arial",256,8,16777215,0) ' +
               ' With Dqa+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_dqa');
  //Tch
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.tqa ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into TqaTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TqaTmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.tqa '
        else
          wCondition := ' order by tch_file.tqa desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into TqaTmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into TqaTmp');
    end;


    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into TqaTmp'); }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into TqaTmp')
      else
         oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + 'into TqaTmp')
    end
    else
    begin  //
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.tqa ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' +  wCondition + ' into TqaTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into TqaTmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.tqa '
          else
            wCondition := ' order by tch_file.tqa desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' +  wCondition + ' into TqaTmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into TqaTmp');
      end;
      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into TqaTmp'); }
    end;
  end;

  oleMapInfo.do('commit table TqaTmp as "' + gExePath + 'tch_Tqa.tab"');
  oleMapInfo.do('close table TqaTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_Tqa.tab" Interactive');

  oleMapInfo.do('select Max(Tqa) from tch_Tqa into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_Tqa, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_Tqa');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from Tch_tqa where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_Tqa.lon');
      wLat := oleMapInfo.eval('tch_Tqa.lat');
      wBearing := oleMapInfo.eval('tch_Tqa.Bearing');
      //if oleMapInfo.eval('tch_tqa.tqa') > 0 then
      try
        wRate := oleMapInfo.eval('tch_tqa.tqa') / wMaxQty
      except
        wRate := 0;
      end;
      //else
        //wRate := 0;
      //wRate := oleMapInfo.eval('tch_Tqa.Tqa') / wMaxQty;
      if wRate > 0 then
        wRate := 0.01 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('Tch_tqa.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_tqa')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'Tch_tqa');
    {oleMapInfo.do('set style brush makebrush(82,rgb(' +

                  FloatToStr(255 * (1 - wRate)) + ',' +
                  FloatToStr(255 * (1 - wRate)) + ',0), rgb(255,255,255))');

    CreateRegion_4(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2, 'R');
    oleMapInfo.do('Update Tch_Tqa set obj = TmpObject where rowid = ' + IntToStr(i));}
    end;
  end;
  oleMapInfo.do('commit table tch_Tqa');
  oleMapInfo.do('add map auto layer tch_Tqa');
  oleMapInfo.do('Set Map Layer tch_Tqa Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With Tqa+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  //oleMapInfo.do('set map layer cell display off');

  oleMapInfo.do('set map redraw on');
  gDqa := True;
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  oleMapInfo.do('select  cell_id from tch_tqa into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_dqa into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  //oleMapInfo.do('Set Map Layer cell Display on');
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 6;
end;

procedure TfmBscMain.mmRafClick(Sender: TObject);
procedure raf_shade;
var
  wCondNum , wRow, i : Integer;
  wCellid : string;
  wLon , wLat, wBearing, wRate : Real;
begin
  if  (gSelFlag <> 'CELL') then
    ShowCondition('随机失败率')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.raf ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into Tmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by cch_file.raf '
        else
          wCondition := ' order by cch_file.raf desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into Tmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
    end;
    {oleMapInfo.do('Select * from cell where cell.bsc_no = "'
                 + gSelName + '" into tmp');  }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select * from cell where cell.bs_no = "'
                 + gSelName + '" into tmp')
      else
        oleMapInfo.do('Select * from cell where bs_no > "0"  and ' +
                 GetMultiCell(gMultiCell) + ' into tmp');

    end
    else //all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.raf ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into Tmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by cch_file.raf '
          else
            wCondition := ' order by cch_file.raf desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into Tmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.raf, cch_file.raa, cch_file.ras, cch_file.rac, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
      end;
      //oleMapInfo.do('Select * from cell into tmp');
    end;
  end;
  oleMapInfo.do('commit table tmp as "' + gExePath + 'raf_shade.tab"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('open table "' + gExePath + 'raf_shade.tab"');
  wRow := oleMapInfo.eval('tableinfo(raf_shade, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from raf_shade');
    wLon := oleMapInfo.eval('raf_shade.lon');
    wLat := oleMapInfo.eval('raf_shade.lat');
    wBearing := oleMapInfo.eval('raf_shade.Bearing');
    wRate := 0;
    CreateRegion_5(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2);
    oleMapInfo.do('Update raf_shade set Obj = TmpObject where rowId = ' + IntToStr(i));
  end;
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
  begin
    wCondNum := StrToInt(wTchQty);
    for i := wCondNum to wRow do
      oleMapInfo.do('delete from raf_shade  where rowid = ' + IntToStr(i) );
   // wRow := wCondNum;
  end;
  //wRow := oleMapInfo.eval('tableinfo(raf_shade, 8)');
  //oleMapInfo.do('fetch first from raf_shade');

  oleMapInfo.do('commit table raf_shade');
  oleMapInfo.do('add map auto layer raf_shade');
  {oleMapInfo.do('Add Column "raf_shade" (raf Decimal (8, 2))From cch_file Set To raf Where COL2 = COL6  Dynamic');
  oleMapInfo.do('Add Column "raf_shade" (raa Decimal (8, 2))From cch_file Set To raa Where COL2 = COL6  Dynamic');
  oleMapInfo.do('Add Column "raf_shade" (ras Decimal (8, 2))From cch_file Set To ras Where COL2 = COL6  Dynamic');
  oleMapInfo.do('Add Column "raf_shade" (rac Decimal (8, 2))From cch_file Set To rac Where COL2 = COL6  Dynamic');}
 { oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' raf_shade  with Arfcn ' +
                'ignore 0 ranges apply all use color Brush (2,16711680,16777215) '+
                '  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,25: 50 Brush ' +
                ' (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush (2,11554816,16777215) ' +
                ' Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) Pen (1,2,0) ' +
                ' default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.1 ' +
                ' inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
 }
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' raf_shade with Raf ignore 0 ranges apply all use color Brush ' +
             ' (2,65280,16777215)  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             ' 25: 50 Brush (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush ' +
             ' (2,11554816,16777215) Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) ' +
             ' Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 ' +
             ' round 0.1 inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
   ' layer prev display on shades on symbols off lines off count on title "随机接入" '  +
   ' Font ("Arial",0,12,0) subtitle auto Font ("Arial",0,11,0) ascending off ' +
   ' ranges Font ("Arial",0,11,0) auto display off ,auto display on ,auto ' +
   ' display on ,auto display on ,auto display on ');

  oleMapInfo.do('Set Map Layer Raf_shade Label Position Above Font ("Arial",1,10,0) With Raf+"%,"+raa+"%"+chr$(13)+ras+"%,"+Rac+"%" Auto On Visibility Zoom (0, 6) Units "km"');
 // oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer public');
  gRaf := True;
  oleMapInfo.do('select  cell_id from  raf_shade into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
end;
begin
//  fmBscMain.ResetMap;
  if gRaf then
  begin
    if mmRaf.Checked then
    begin
      oleMapInfo.do('close table raf_shade');
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Off');
      //oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on')   }
    end
    else
    begin
      raf_shade;
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Graphic');
     // oleMapInfo.do('Set Map Layer RafTmp Display Graphic');
      oleMapInfo.do('set map redraw on')}
    end;
    mmRaf.Checked := not mmRaf.Checked;
    exit;
  end;
  mmRaf.Checked := not mmRaf.Checked;
  //oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  raf_shade;

  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
end;

procedure TfmBscMain.mmHoverClick(Sender: TObject);
var
  wStr, wCellId, wCellSe, wCellTa : String;
  i, wRow, wCondNum : Integer;
  wLon, wLat, wBearing_s, wBearing_t, wLon_s, wLat_s, wLon_t, wLat_t, wRate, wFlunkRate, wHolost, wMaxQty : real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gHover then
  begin
    if mmHover.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Hdov_i_bsc Display Off');
      oleMapInfo.do('Set Map Layer Hdov_e_bsc Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Hdov_i_bsc Display Graphic');
      oleMapInfo.do('Set Map Layer Hdov_e_bsc Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmHover.Checked := not mmHover.Checked;
    exit;
  end;
  mmHover.Checked := not mmHover.Checked;
  {if  (gSelFlag <> 'CELL') then
  begin
    Application.CreateForm(TfmHdovCond, fmHdovcond);
    try
      fmHdovCond.ShowModal;
    finally
      fmHdovCond.Free;
    end;
  end
  else
    wFilterOrder := ''; }
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  oleMapInfo.do('Open table "' + gExePath + 'Hdov_e_61bsc.tab" Interactive');
  oleMapInfo.do('Open table "' + gExePath + 'Hdov_i_61bsc.tab" Interactive');
  oleMapInfo.do('Alter Table "Hdov_i_61Bsc" ( Add Flunk_Rate Decimal(5,2) ) Interactive');
  oleMapInfo.do('Update Hdov_i_61Bsc Set Flunk_Rate = 100 * horttoch / (hovercnt + 0.00001)');
  oleMapInfo.do('Commit table hdov_i_61bsc');
  //,FlunkRate
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Lon_s Decimal (12, 6))From Cell Set To Lon Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Lon_t Decimal (12, 6))From Cell Set To Lon Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Lat_s Decimal (12, 6))From Cell Set To Lat Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Lat_t Decimal (12, 6))From Cell Set To Lat Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Bearing_s Decimal (12, 6))From Cell Set To Bearing Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Bearing_t Decimal (12, 6))From Cell Set To Bearing Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Bsc_no_s Char (6))From Cell Set To Bsc_no Where COL5 = COL2 Dynamic');
  oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Bsc_no_t Char (6))From Cell Set To Bsc_no Where COL6 = COL2 Dynamic');
  //oleMapInfo.do('Add Column "Hdov_i_61Bsc" (Hovercnt_t Decimal (9, 0))From Hdov_i_61bsc Set To Bearing Where COL1 = COL2 and col2 = col1  Dynamic');
  //oleMapInfo.do('commit table hdov_i_61bsc as "' + gExePath + 'hdov_i_tmp.tab"');
  //oleMapInfo.do('close table hdov_i_bsc_tmp');
  //oleMapInfo.do('open table "' + gExePath + 'Hdov_i_tmp.tab" Interactive');



  oleMapInfo.do('set style pen makepen(1,1, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  if gSelFlag = 'BSC' then
  begin
    {if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and hdov_i_61bsc.holost ' + wTchFlag ;
        oleMapInfo.do('select * from hdov_i_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition + ' into hdov_i_bsc_tmp');
      end
      else
        oleMapInfo.do('select * from hdov_i_61bsc where bsc_no_s = "0" ' +
                     ' and lon_s > 0 and lat_t >0 into hdov_i_bsc_tmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by Hdov_i_61bsc.Flunk_rate '
        else
          wCondition := ' order by Hdov_i_61bsc.Flunk_rate desc' ;
        oleMapInfo.do('select * from hdov_i_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition  + ' into hdov_i_bsc_tmp');
      end;
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by Hdov_i_61bsc.Holost '
        else
          wCondition := ' order by Hdov_i_61bsc.holost desc' ;
        oleMapInfo.do('select * from hdov_i_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition  + ' into hdov_i_bsc_tmp');
      end;
    end;}

    oleMapInfo.do('select * from hdov_i_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 into hdov_i_bsc_tmp');
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
         oleMapInfo.do('select * from hdov_i_61bsc where (cell_id_se = "' +
                    gSelName + '" or cell_id_ta = "' +
                    gSelName + '") and lon_s > 0 and lat_t >0 into hdov_i_bsc_tmp')

      else
      begin
        wStr := ' (';
        for i := 0 to gMultiCell.Count -1 do
        begin
          wCellId := Copy(gMultiCell.Strings[i], 1, Pos(' ', gMultiCell.Strings[i]) - 1);
          if i < gMultiCell.Count -1 then
            wStr := wStr + 'cell_id_se = "' +
                    wCellId + '" or cell_id_ta = "'  + wCellId + '" or '
          else
            wStr :=  wStr + 'cell_id_se = "' + wCellId +
                            '" or cell_id_ta = "'  + wCellId + '" ) ';
        end;
        oleMapInfo.do('select * from hdov_i_61bsc where ' + wStr +
                     ' and lon_s > 0 and lat_t >0 into hdov_i_bsc_tmp');

      end;
    end
    else
    begin
      {if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and hdov_i_61bsc.holost ' + wTchFlag ;
          oleMapInfo.do('select * from hdov_i_61bsc where  lon_s > 0 and lat_t >0 '
                     + wCondition + ' into hdov_i_bsc_tmp');
        end
        else
          oleMapInfo.do('select * from hdov_i_61bsc where bsc_no_s = "0" ' +
                     ' and lon_s > 0 and lat_t >0 into hdov_i_bsc_tmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by Hdov_i_61bsc.Flunk_rate '
          else
            wCondition := ' order by Hdov_i_61bsc.Flunk_rate desc' ;
          oleMapInfo.do('select * from hdov_i_61bsc where lon_s > 0 and lat_t >0 '
          + wCondition  + ' into hdov_i_bsc_tmp');
        end;
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by Hdov_i_61bsc.Holost  '
          else
            wCondition := ' order by Hdov_i_61bsc.Holost  desc' ;
          oleMapInfo.do('select * from hdov_i_61bsc where lon_s > 0 and lat_t >0 '
                  + wCondition  + ' into hdov_i_bsc_tmp');
        end;
      end;}
      oleMapInfo.do('select * from hdov_i_61bsc where  lon_s > 0 and lat_t >0 '+
                    ' into hdov_i_bsc_tmp');
    end;
  end;
  //oleMapInfo.do('Close table hdov_i_tmp');
  oleMapInfo.do('commit table hdov_i_bsc_tmp as "' + gExePath + 'hdov_i_bsc.tab"');
  oleMapInfo.do('close table hdov_i_bsc_tmp');
  oleMapInfo.do('open table "' + gExePath + 'Hdov_i_bsc.tab" Interactive');
  oleMapInfo.do('Alter Table "Hdov_i_61Bsc" ( drop Flunk_Rate ) Interactive');
  oleMapInfo.do('commit table hdov_i_61bsc');
  oleMapInfo.do('Create Map For Hdov_i_bsc CoordSys Earth Projection 1, 0');
  oleMapInfo.do('select max(hovercnt) from hdov_i_bsc into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('close table tmp');
  {oleMapInfo.do('select max(hovercnt) from hdov_e_bsc into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('close table tmp');  }
  //oleMapInfo.do('drop map from Hdov_i_bsc');
  //oleMapInfo.do('create map from Hdov_i_Bsc');

  wRow := oleMapInfo.eval('tableinfo(Hdov_i_bsc,8)');
 { oleMapInfo.do('Alter Table "Hdov_i_bsc" ( Add Hdov_Total Decimal(9,0) ) Interactive');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Hdov_i_bsc');
  end; }
  oleMapInfo.do('fetch first from Hdov_i_bsc');
 { if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);}
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Hdov_i_bsc');
    {if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from hdov_i_bsc where rowid = ' + IntToStr(i));
    end
    else }
    begin
      wLon_s := oleMapInfo.eval('Hdov_i_bsc.Lon_s');
      wLat_s := oleMapInfo.eval('Hdov_i_bsc.Lat_s');
      wLon_t := oleMapInfo.eval('Hdov_i_bsc.Lon_t');
      wLat_t := oleMapInfo.eval('Hdov_i_bsc.Lat_t');
      wRate := oleMapInfo.eval('Hdov_i_bsc.hovercnt') / wMaxQty ;
      wBearing_s := oleMapInfo.eval('Hdov_i_bsc.Bearing_s');
      wBearing_t := oleMapInfo.eval('Hdov_i_bsc.Bearing_t');
      {if oleMapInfo.eval('Hdov_i_bsc.hovercnt') <> 0 then
        wFlunkRate :=  oleMapInfo.eval('Hdov_i_bsc.horttoch') /
                     oleMapInfo.eval('Hdov_i_bsc.hovercnt')
      else
        wFlunkRate := 0; }
      if oleMapInfo.eval('Hdov_i_bsc.Flunk_Rate') > 0 then
        wFlunkRate := oleMapInfo.eval('Hdov_i_bsc.Flunk_Rate') / 100
      else
        wFlunkRate := 0;
      if oleMapInfo.eval('Hdov_i_bsc.Holost') > 0 then
        wHolost := oleMapInfo.eval('Hdov_i_bsc.Holost')/100
      else
        wHolost := 0;
    {if wRate = 0 then
    begin
      oleMapInfo.do('Set Style pen makepen(2 ,59, RGB(0,0,0))');
    end
    else
    begin
      if (wHoLost > 0.05)  then
      begin
         oleMapInfo.do('Set Style pen makepen(' + FloatToStr(2 + 4 * wRate) + ',59, RGB(255,0,0))');
      end
      else
      begin
        if (wFlunkRate > 0.10)  then
        begin
          oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * wRate) + ',59, RGB(255,0,255))');
        end
        else
        begin
          oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * wRate) + ',59, RGB(0,0,255))');
        end;
      end;
    end;  }
      wCellSe := oleMapInfo.eval('Hdov_i_bsc.cell_id_se');
      wCellTa := oleMapInfo.eval('Hdov_i_bsc.cell_id_ta');
    //UpDateLine(Lon_1, Lat_1, Lon_2, Lat_2, L, P, Bearing, Rate : Real; V : Variant; rowid: Integer );
      if (wLon_t >  0) and (wLon_s > 0) and (wLat_s >  0) and (wLat_t > 0)  then
         UpDateLine(wLon_s, wLat_s, wLon_t, wLat_t, gCellLength, gCellAngle, wBearing_s, wBearing_t, wRate, wFlunkRate,  oleMapInfo, i,'Hdov_i_bsc', wCellSe, wCellTa, wHolost, 0.5)
      else
        oleMapInfo.do('update hdov_i_bsc  set obj = createCircle(0,0,0) where rowId =' + IntToStr(i));
    end;
  end;
  oleMapInfo.do('commit table Hdov_i_bsc');
  oleMapInfo.do('Add Map auto layer Hdov_i_bsc');
  oleMapInfo.do('Set Map Layer Hdov_i_bsc Label Position Center Font ("Arial",0,10,0) ' +
                ' With Hovercnt+"("+Hoversuc+","+Horttoch+")"+' +
                ' Chr$(13)+Holost+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer hdov_i_bsc');


  /////////////////////////////bsc --->bsc

  oleMapInfo.do('Alter Table "Hdov_e_61Bsc" ( Add Flunk_Rate Decimal(5,2) ) Interactive');
  oleMapInfo.do('Update Hdov_e_61Bsc Set Flunk_Rate = 100 * horttoch / (hovercnt + 0.00001)');
  oleMapInfo.do('Commit table hdov_e_61bsc');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Lon_s Decimal (12, 6))From Cell Set To Lon Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Lon_t Decimal (12, 6))From Cell Set To Lon Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Lat_s Decimal (12, 6))From Cell Set To Lat Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Lat_t Decimal (12, 6))From Cell Set To Lat Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Bearing_s Decimal (12, 6))From Cell Set To Bearing Where COL5 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Bearing_t Decimal (12, 6))From Cell Set To Bearing Where COL6 = COL2  Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Bsc_no_s Char (6))From Cell Set To Bsc_no Where COL5 = COL2 Dynamic');
  oleMapInfo.do('Add Column "Hdov_e_61Bsc" (Bsc_no_t Char (6))From Cell Set To Bsc_no Where COL6 = COL2 Dynamic');
  //oleMapInfo.
  //oleMapInfo.do('open table "' + gExePath + 'Hdov_e_tmp.tab" Interactive');

  if gSelFlag = 'BSC' then
  begin
    {if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and hdov_e_61bsc.holost ' + wTchFlag ;
        oleMapInfo.do('select * from hdov_e_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition + ' into hdov_e_bsc_tmp');
      end
      else
        oleMapInfo.do('select * from hdov_e_61bsc where bsc_no_s = "0" ' +
                     ' and lon_s > 0 and lat_t >0 into hdov_e_bsc_tmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by Hdov_e_61bsc.Flunk_rate '
        else
          wCondition := ' order by Hdov_e_61bsc.Flunk_rate desc' ;
        oleMapInfo.do('select * from hdov_e_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition  + ' into hdov_e_bsc_tmp');
      end;
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by Hdov_e_61bsc.Holost '
        else
          wCondition := ' order by Hdov_e_61bsc.holost desc' ;
        oleMapInfo.do('select * from hdov_e_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 ' + wCondition  + ' into hdov_e_bsc_tmp');
      end;
    end; }
    oleMapInfo.do('select * from hdov_e_61bsc where (bsc_no_s = "' +
                    gSelName + '"  or bsc_no_t = "' + gSelName +
                    '") and lon_s > 0 and lat_t >0 into hdov_e_bsc_tmp');
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('select * from hdov_e_61bsc where (cell_id_se = "' +
                    gSelName + '" or cell_Id_ta ="' +
                    gSelName + '") and lon_s > 0 and lat_t >0 into hdov_e_bsc_tmp')
      else
        oleMapInfo.do('select * from hdov_e_61bsc where ' +  wstr +
                    ' and lon_s > 0 and lat_t >0 into hdov_e_bsc_tmp') ;
    end
    else //all
    begin
      {if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and hdov_e_61bsc.holost ' + wTchFlag ;
          oleMapInfo.do('select * from hdov_e_61bsc where  lon_s > 0 and lat_t >0 '
                     + wCondition + ' into hdov_e_bsc_tmp');
        end
        else
          oleMapInfo.do('select * from hdov_e_61bsc where bsc_no_s = "0" ' +
                     ' and lon_s > 0 and lat_t >0 into hdov_e_bsc_tmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by Hdov_e_61bsc.Flunk_rate '
          else
            wCondition := ' order by Hdov_e_61bsc.Flunk_rate desc' ;
          oleMapInfo.do('select * from hdov_e_61bsc where lon_s > 0 and lat_t >0 '
          + wCondition  + ' into hdov_e_bsc_tmp');
        end;
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by Hdov_e_61bsc.Holost  '
          else
            wCondition := ' order by Hdov_e_61bsc.Holost  desc' ;
          oleMapInfo.do('select * from hdov_e_61bsc where lon_s > 0 and lat_t >0 '
                  + wCondition  + ' into hdov_e_bsc_tmp');
        end;
      end; }
      oleMapInfo.do('select * from hdov_e_61bsc where lon_s > 0 and lat_t >0 ' +
                    'into hdov_e_bsc_tmp');
    end;
  end;
  //oleMapInfo.do('close table hdov_e_tmp');
  oleMapInfo.do('commit table hdov_e_bsc_tmp as "' + gExePath + 'hdov_e_bsc.tab"');
  oleMapInfo.do('close table hdov_e_bsc_tmp');
  oleMapInfo.do('open table "' + gExePath + 'Hdov_e_bsc.tab" Interactive');
  oleMapInfo.do('Alter Table "Hdov_e_61Bsc" ( drop Flunk_Rate ) Interactive');
 // oleMapInfo.do('Update Hdov_e_61Bsc Set Flunk_Rate = 100 * horttoch / (hovercnt + 0.00001)');
  oleMapInfo.do('Commit table hdov_e_61bsc');
  oleMapInfo.do('Create Map For Hdov_E_bsc CoordSys Earth Projection 1, 0');
  wRow := oleMapInfo.eval('tableinfo(Hdov_e_bsc,8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);

  oleMapInfo.do('fetch first from Hdov_e_bsc');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Hdov_e_bsc');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from hdov_e_bsc where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon_s := oleMapInfo.eval('Hdov_e_bsc.Lon_s');
      wLat_s := oleMapInfo.eval('Hdov_e_bsc.Lat_s');
      wLon_t := oleMapInfo.eval('Hdov_e_bsc.Lon_t');
      wLat_t := oleMapInfo.eval('Hdov_e_bsc.Lat_t');
      wRate := oleMapInfo.eval('Hdov_e_bsc.hovercnt') / wMaxQty ;
   // if oleMapInfo.eval('Hdov_e_bsc.hovercnt') = 68 then
   //   ShowMessage('13');
      wBearing_s := oleMapInfo.eval('Hdov_e_bsc.Bearing_s');
      wBearing_t := oleMapInfo.eval('Hdov_e_bsc.Bearing_t');
      if oleMapInfo.eval('Hdov_e_bsc.hovercnt') <> 0 then
        wFlunkRate := oleMapInfo.eval('Hdov_e_bsc.Horttoch') /
                     oleMapInfo.eval('Hdov_e_bsc.hovercnt')
      else
        wFlunkRate := 0;
      if oleMapInfo.eval('Hdov_E_bsc.Holost') > 0 then
        wHolost := oleMapInfo.eval('Hdov_E_bsc.Holost')/100
      else
        wHolost := 0;
   { if wRate = 0 then
    begin
      oleMapInfo.do('Set Style pen makepen(2 ,59, RGB(0,0,0))');
    end
    else
    begin
      if (wHoLost > 0.05)  then
      begin
         oleMapInfo.do('Set Style pen makepen(' + FloatToStr(2 + 4 * wRate) + ',59, RGB(255,0,0))');
      end
      else
      begin
        if (wFlunkRate < 0.10)  then
        begin
          oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * wRate) + ',59, RGB(255,0,255))');
        end
        else
        begin
          oleMapInfo.do('Set Style pen makepen(' + FloatToStr(1 + 5 * wRate) + ',59, RGB(0,0,255))');
        end;
      end;
    end; }
      wCellSe := oleMapInfo.eval('Hdov_e_bsc.cell_id_se');
      wCellTa := oleMapInfo.eval('Hdov_e_bsc.cell_id_ta');
    //UpDateLine(Lon_1, Lat_1, Lon_2, Lat_2, L, P, Bearing, Rate : Real; V : Variant; rowid: Integer );
      if (wLon_t >  0) and (wLon_s > 0) and (wLat_s >  0) and (wLat_t > 0)  then
        UpDateLine(wLon_s, wLat_s, wLon_t, wLat_t, gCellLength, gCellAngle, wBearing_s, wBearing_t, wRate, wFlunkRate, oleMapInfo, i,'Hdov_E_bsc', wCellSe, wCellTa, wHolost, 0.5)
      else
        oleMapInfo.do('update hdov_e_bsc  set obj = createCircle(0,0,0) where rowId =' + IntToStr(i));
    end;
  end;

  oleMapInfo.do('commit table Hdov_e_bsc');
  oleMapInfo.do('Add Map auto layer Hdov_e_bsc');
  oleMapInfo.do('Set Map Layer Hdov_e_bsc Label Position Center Font ("Arial",0,8,0) ' +
                ' With Hovercnt+"("+Hoversuc+","+Horttoch+")"+' +
                ' Chr$(13)+Holost+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  
  gHover := True;
  ////Bsc }

  oleMapInfo.do('Export "hdov_i_bsc" Into "' + gExePath + 'hdov_i_bsc.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  //oleMapInfo.do('close table tmp');
  //oleMapInfo.do('select  cell_id from Cch_su into tmp');
  oleMapInfo.do('Export "hdov_e_bsc" Into "' + gExePath + 'hdov_e_bsc.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  //oleMapInfo.do('close table tmp');
  //oleMapInfo.do('select  cell_id from tch_u into tmp');
  //Alter Table "hdov_i_bsc" ( modify Cell_id_se Char(8) ) Interactive
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 3;
end;

procedure TfmBscMain.mmSelDataClick(Sender: TObject);
begin
  Application.CreateForm(TfmDataHist, fmDataHist);
  try
    fmDataHist.ShowModal;
    mmSelData.Enabled := True;
  finally
    fmDataHist.Free;
  end;
end;

procedure TfmBscMain.mmOpenMapClick(Sender: TObject);
var
  sWinHand : String;
  sMsg : String;
  i, wTableNum : Integer;
  wLon, wLat : Real;
begin
  with dmBscData.tbBscControl do
  begin
    if not Active then
      Open;
    First;
    if FieldByName('bsc_file_update').AsString = 'N' then
    begin
      ShowMessage('请更新数据，输入UPDATE_BSC_FILE!');
      Exit;
    end;
    if FieldByName('bsc_update').AsString = 'N' then
    begin
      ShowMessage('请更新数据，输入UPDATE_BSC_FILE!');
      Exit;
    end;
    if (FieldByName('Cell_update').AsString = 'N') or
       (FieldByName('BASE_update').AsString = 'N') then
    begin
      ShowMessage('请加载地图，输入CELL.TAB,BASE.TAB!');
      Exit;
    end;
    Close;
  end;
  CreateMDIChild('地图');
  oleMapInfo.do('Open table "' + gExePath + 'area.tab" Interactive');
  oleMapInfo.do('Map from Area');
  //oleMapInfo.do('Set Map  Center (113.325662, 22.279674)');
  //oleMapInfo.do('Set Map  Scale 1 Units "cm" For 3 Units "km"');
  oleMapInfo.do('Set Map Layer Area Label Position Above Font ("Arial",256,10,0,16777215) Auto On Visibility Zoom (50, 1000) Units "km"');
 // oleMapInfo.do('Set Map Layer Area Label Auto On');

  oleMapInfo.do('Open table "' + gExePath + 'base.tab" Interactive');
  oleMapInfo.do('Add Map auto layer base');
  oleMapInfo.do('Set Map Layer Base Label Position Above Font ("Arial",256,10,0,16777215) Auto On Visibility Zoom (0, 50) Units "km"');
  oleMapInfo.do('set style Symbol makeSymbol(33,9437256,8)');



  oleMapInfo.do('Open table "' + gExePath + 'water.tab" Interactive');
  oleMapInfo.do('Add Map auto layer water');

  oleMapInfo.do('Open table "' + gExePath + 'landmark.tab" Interactive');
  oleMapInfo.do('Add Map auto layer landmark');

  oleMapInfo.do('Open table "' + gExePath + 'mountain.tab" Interactive');
  oleMapInfo.do('Add Map auto layer mountain');

  oleMapInfo.do('Open table "' + gExePath + 'block.tab" Interactive');
  oleMapInfo.do('Add Map auto layer block');

 { oleMapInfo.do('Open table "' + gExePath + 'public.tab" Interactive');
  oleMapInfo.do('Add Map auto layer public');
   oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))
          + ' Zoom Entire Layer public'); }

  oleMapInfo.do('Open table "' + gExePath + 'vip.tab" Interactive');
  oleMapInfo.do('Add Map auto layer vip');

  oleMapInfo.do('Open table "' + gExePath + 'street.tab" Interactive');
  oleMapInfo.do('Add Map auto layer street');

  with dmBscData.tbBscControl do
  begin
    if not Active then
      Open;
    First;
    if FieldByName('cell_Obj_Flag').AsString = 'Y' then
    begin
      oleMapInfo.do('Open table "' + gExePath + 'CellObj.tab" Interactive');
      oleMapInfo.do('Add Map auto layer CellObj');
    end;
    Close;
  end;
  
  oleMapInfo.do('Open table "' + gExePath + 'cell.tab" Interactive');
  oleMapInfo.do('Add Map auto layer cell');
  oleMapInfo.do('Set Map Layer cell Label Font ("Arial",0,10,0) ' +
                'With Arfcn Auto On Visibility Zoom (0, 2) Units "km"');
     //  Set Map Layer 1 Label With Arfcn Auto On Visibility Zoom (0, 2) Units "km"
  //wTableNum := oleMapInfo.eval('NumTables()');


  oleMapInfo.do('set map layer AREA selectable off');


  oleMapInfo.do('set map layer cell selectable on');
  //UpdateCellObject;

  oleMapInfo.do('Open table "' + gExePath + 'All_CCH_FILE.tab" Interactive');
  oleMapInfo.do('Open table "' + gExePath + 'All_TCH_FILE.tab" Interactive');

  oleMapInfo.do('Open table "' + gExePath + 'CCH_FILE.tab" Interactive');
  oleMapInfo.do('Open table "' + gExePath + 'TCH_FILE.tab" Interactive');



  sbBscMain.Panels[1].Text := IntToStr(oleMapInfo.eval('cch_file.start_date'))
                              + ' [' + IntToStr(oleMapInfo.eval('cch_file.start_Time'))
                              + '-' + IntToStr(oleMapInfo.eval('cch_file.End_Time'))
                              + ']';
  mmSelDataClick(self);
  mmOpenMap.Enabled := False;
  //mmSelData.Enabled := False;
  //oleMapInfo.RunMenuCommand(102);
  gMapNo := 1;
end;

procedure TfmBscMain.SpeedButton7Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1709);
end;

procedure TfmBscMain.SpeedButton5Click(Sender: TObject);
begin
  Application.CreateForm(TfmCompare, fmCompare);
  try
    fmCompare.ShowModal;
  finally
    fmCompare.Free;
  end;
end;

{procedure TfmBscMain.mmCgShadeClick(Sender: TObject);
begin
  fmBscMain.ResetMap;
  if gTchCG then
  begin
    if mmCgShade.Checked then
    begin
      oleMapInfo.do('close table cg_shade');
     { oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cg_shade Display Off');
     // oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on') }
    {end
    else
    begin
      TchCGShade;
     { oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cg_shade Display Graphic');
      //oleMapInfo.do('Set Map Layer Tch_dr Display Graphic');
      oleMapInfo.do('set map redraw on')  }
    {end;
  end
  else
    TchCGShade;
  mmCgShade.Checked := not mmCgShade.Checked;
  if mmTraShade.Checked and mmCgShade.Checked then
  begin
    mmTraShadeClick(self);
  end;
end; }

procedure TfmBscMain.mmTraShadeClick(Sender: TObject);
begin
//  fmBscMain.ResetMap;
  if gTchTraffic then
  begin
    if mmTraShade.Checked then
    begin
      oleMapInfo.do('close table erpac_shade');
     { oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer traffic_shade Display Off');
     // oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on') }
    end
    else
    begin
      TchTrafficShade;
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer traffic_shade Display Graphic');
      //oleMapInfo.do('Set Map Layer Tch_dr Display Graphic');
      oleMapInfo.do('set map redraw on')   }
    end;
  end
  else
    TchTrafficShade;
  mmTraShade.Checked := not mmTraShade.Checked;
  if mmTraShade.Checked {and mmCgShade.Checked} then
  begin
    //mmCgShadeClick(self);
  end;
 //oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer public');
end;

procedure TfmBscMain.MSC1Click(Sender: TObject);
var
  oleExcel  : Variant;
  i, j, p: Integer;
  wStr, wStartDate, wStartTime, wEndTime, wBscNo, wTmpComp, wCTRALACC, wCNSCAN : String;
  wYear, wMonth, wDay: Word;
  wSumCTRALCC, wSumCNSCAn, wSumCount : Integer;
begin
  if not odData.Execute then
    exit;
  sbBscMain.Panels[0].Text := '正在输入文件' + odData.Files[0] + '...  ';
  oleExcel := CreateOleObject('Excel.Application');
  for p := 0 to odData.Files.Count - 1 do
  begin
    sbBscMain.Panels[0].Text := '正在输入文件' + odData.Files[p] + '...  ';
    wBscNo := odData.Files[p];
    while Pos('\',wBscNo) > 0 do
      Delete(wBscNo,1,Pos('\',wBscNo));
    wBscNo := Trim(wBscNo);



    //oleExcel.Caption:='万禾通信';
    if P > 0 then
      oleExcel.WorkBooks[1].Close;
    oleExcel.WorkBooks.Open(odData.Files[p]);
    oleExcel.WorkSheets[1].Activate;
    if Pos('Object', wBscNo) > 0 then
    begin
      //切换
      Delete(wBscNo, 1, 6);
      wStr := wBscNo;
      Delete(wBscNo, Length(wBscNo) - 12 , 15);
      wBscNo := UpperCase(wBscNo);
      wStr := Copy(wStr, Length(wStr) - 12 , 9);
      wEndTime := Copy(wStr, 6, 4);
      wStartTime := IntToStr(StrToInt(wEndTime) - 100);
      DecodeDate(Now, wYear, wMonth, wDay);
      //if wYear < 2009 then
      wStartDate := IntToStr(wYear) + Copy(wStr, 1, 4);
      oleExcel.WorkBooks[1].WorkSheets['NCELLREL'].Activate;
      wStr := oleExcel.Cells[1,1];
      if wStr <> 'CELL_ID_SERV' then
      begin
        ShowMessage(odData.Files[p] + '切换文件内容不对！');
        Continue;
        //Exit;
      end;
      with dmBscData.quAllHdovI do
      begin
        if Active then
          Close;
        ParamByName('start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
      //ParamByName('End_Time').AsString := wEndTime;
        ParamByName('Bsc_no').AsString := wBscNo;
        Open;
        if not IsEmpty then
        begin
          if MessageDlg('数据已存在, 是否覆盖?',
            mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            with dmBscData.quDelHdovI do
            begin
              ParamByName('start_date').AsInteger := StrToInt(wStartDate);
              ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
              //ParamByName('End_Time').AsString := wEndTime;
              ParamByName('Bsc_no').AsString := wBscNo;
              ExecSQL;
            end;
          end
          else
          begin
            Continue;
            //Exit;
          end;
        end;
        i := 2;
        wTmpComp := oleExcel.cells[i, 1] ;
        //showMessage(oleExcel.cells[i, 1]);
        while wTmpComp  <> '' do
        begin
          try
            Insert;
            dmBscData.quAllHdovISTART_TIME.AsString := wStartTime;
            dmBscData.quAllHdovIEND_TIME.AsString := wEndTime;
            dmBscData.quAllHdovISTART_DATE.AsString := wStartDate;
            dmBscData.quAllHdovIBSC_NO_SER.AsString :=  wBscNo;
            dmBscData.quAllHdovICELL_ID_SE.AsString := oleExcel.cells[i,1];
            dmBscData.quAllHdovICELL_ID_TA.AsString := oleExcel.cells[i,2];
            dmBscData.quAllHdovIHOVERCNT.AsString := oleExcel.cells[i,3];
            dmBscData.quAllHdovIHOVERSUC.AsString := oleExcel.cells[i,4];
            dmBscData.quAllHdovIHORTTOCH.AsString := oleExcel.cells[i,5];
            if dmBscData.quAllHdovIHOVERCNT.AsInteger > 0 then
               dmBscData.quAllHdovIHOLOST.AsFloat := (dmBscData.quAllHdovIHOVERCNT.AsInteger -
                                      dmBscData.quAllHdovIHORTTOCH .AsInteger -
                                      dmBscData.quAllHdovIHOVERSUC.AsInteger ) /
                                      dmBscData.quAllHdovIHOVERCNT.AsInteger * 100
            else
              dmBscData.quAllHdovIHOLOST.AsFloat := 0;
            Post;
          except
            ShowMessage(wStartDate + '(' + wStartTime +
              '-' + wEndTime + ') ' + wTmpComp + '小区间的切换数据有问题!');
          end;
          i := i + 1;
          wTmpComp := oleExcel.cells[i, 1] ;
        end;
        Close;
      end;
      oleExcel.WorkBooks[1].WorkSheets['NECELLREL'].Activate;
      i := 2;
    //////////////
      with dmBscData.quAllHdovE do
      begin
        if Active then
          Close;
        ParamByName('start_date').AsInteger := StrToInt(wStartDate);
        ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
      //ParamByName('End_Time').AsString := wEndTime;
        ParamByName('Bsc_no').AsString := wBscNo;
        Open;
        if not IsEmpty then
        begin
          if MessageDlg('数据已存在, 是否覆盖?',
            mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            with dmBscData.quDelHdovE do
            begin
              ParamByName('start_date').AsInteger := StrToInt(wStartDate);
              ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
              //ParamByName('End_Time').AsString := wEndTime;
              ParamByName('Bsc_no').AsString := wBscNo;
              ExecSQL;
            end;
          end
          else
          begin
            Continue;
            //Exit;
            //sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
          end;
        end;
        i := 2;
        wTmpComp := oleExcel.cells[i, 1] ;
        //showMessage(oleExcel.cells[i, 1]);
        while wTmpComp  <> '' do
        begin
          try
            Insert;
            dmBscData.quAllHdovESTART_TIME.AsString := wStartTime;
            dmBscData.quAllHdovEEND_TIME.AsString := wEndTime;
            dmBscData.quAllHdovESTART_DATE.AsString := wStartDate;
            dmBscData.quAllHdovEBSC_NO_SER.AsString :=  wBscNo;
            dmBscData.quAllHdovECELL_ID_SE.AsString := oleExcel.cells[i,1];
            dmBscData.quAllHdovECELL_ID_TA.AsString := oleExcel.cells[i,2];
            dmBscData.quAllHdovEHOVERCNT.AsString := oleExcel.cells[i,3];
            dmBscData.quAllHdovEHOVERSUC.AsString := oleExcel.cells[i,4];
            dmBscData.quAllHdovEHORTTOCH.AsString := oleExcel.cells[i,5];
            if dmBscData.quAllHdovEHOVERCNT.AsInteger > 0 then
              dmBscData.quAllHdovEHOLOST.AsFloat := (dmBscData.quAllHdovEHOVERCNT.AsInteger -
                                        dmBscData.quAllHdovEHORTTOCH .AsInteger -
                                        dmBscData.quAllHdovEHOVERSUC.AsInteger ) /
                                        dmBscData.quAllHdovEHOVERCNT.AsInteger * 100
            else
              dmBscData.quAllHdovEHOLOST.AsFloat := 0;
            Post;
          except
            ShowMessage(wStartDate + '(' + wStartTime +
            '-' + wEndTime + ') ' + wTmpComp + 'BSC间的数据有问题!');
          end;
          i := i + 1;
          wTmpComp := oleExcel.cells[i, 1] ;
        end;
        Close;
      end;
      //ShowMessage('BSC数据输入成功!');
      //sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
      Continue;
      //Exit;
    end;


    oleExcel.WorkBooks[1].WorkSheets['TCH'].Activate;
    Delete(wBscNo, Length(wBscNo) - 12 , 15);
    wBscNo := UpperCase(wBscNo);
    wStr := oleExcel.Cells[1,1];
    if Pos('Period', wStr) <= 0 then
    begin
     // ShowMessage('文件内容不对！');
      Continue;
     // Exit;
    end;
    Delete(wStr, 1, Pos(':',wStr));
    //wStr := Trim(wStr);
    wStartTime := Copy(Trim(wStr), 10, 4);
    wEndTime := Copy(Trim(wStr), 15, 4);
    wStr := Copy(Trim(wStr), 1, 8);
    wStartDate := Copy(wStr, 1, 2);
    if wStartDate = '99' then
      wStartDate := '1999'
    else
      wStartDate := '20' + wStartDate ;
    wStartDate := wStartDate + Copy(wStr, 4, 2) + Copy(wStr, 7, 2);
    //showMessage(wBscNo + ' || ' +  wStartDate + ' || ' + wStartTime + ' || ' + wEndTime);
    //showMessage(oleExcel.Cells[10,3]);

    if Trim(UpperCase(oleExcel.Cells[3,10])) <> 'CH' then
      j := 2
    else
      j := 0;
    i := 5;
    wSumCTRALCC := 0;
    wSumCNSCAn := 0;
    wSumCount := 0;
    ////////////////////////////////////////
    with dmBscData.quAllTchFile do
    begin
      if Active then
        Close;
      ParamByName('start_date').AsInteger := StrToInt(wStartDate);
      ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
      //ParamByName('End_Time').AsString := wEndTime;
      ParamByName('Bsc_no').AsString := wBscNo;
      Open;
      if not IsEmpty then
      begin
        if MessageDlg('数据已存在, 是否覆盖?',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          with dmBscData.quDelTch do
          begin
            ParamByName('start_date').AsInteger := StrToInt(wStartDate);
            ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
            //ParamByName('End_Time').AsString := wEndTime;
            ParamByName('Bsc_no').AsString := wBscNo;
            ExecSQL;
          end;
        end
        else
        begin
          Continue;
         // Exit;
        end;
      end;
      wSumCTRALCC := 0;
      wSumCNSCAn := 0;
      i := 5;
      wTmpComp := oleExcel.cells[i, 2] ;
        //showMessage(oleExcel.cells[i, 1]);
      while wTmpComp  <> '' do
      begin
        try
          dmBscData.quAllTchFile.Insert;
          dmBscData.quAllTchFileSTART_TIME.AsString := wStartTime;
          dmBscData.quAllTchFileEND_TIME.AsString := wEndTime;
          dmBscData.quAllTchFileSTART_DATE.AsString := wStartDate;
          dmBscData.quAllTchFileBSC_NO.AsString :=  wBscNo;
          dmBscData.quAllTchFileTCH.AsString := oleExcel.cells[i,1];
          dmBscData.quAllTchFileCELL_ID.AsString := oleExcel.cells[i,2];
          dmBscData.quAllTchFileCA.AsString := oleExcel.cells[i,3];
          dmBscData.quAllTchFileCS.AsString := oleExcel.cells[i,4];
          dmBscData.quAllTchFileU.AsString := oleExcel.cells[i,5];
          dmBscData.quAllTchFileERPAC.AsString := oleExcel.cells[i,6];
          dmBscData.quAllTchFileCG.AsString := oleExcel.cells[i,7];
          dmBscData.quAllTchFileMH.AsString := oleExcel.cells[i,8];
          dmBscData.quAllTchFileDR.AsString := oleExcel.cells[i,9];
          dmBscData.quAllTchFileCH.AsString := oleExcel.cells[i,10 + j];
          dmBscData.quAllTchFileAC.AsString := oleExcel.cells[i,11 + j];
          dmBscData.quAllTchFileF.AsString := oleExcel.cells[i,12 + j];
          dmBscData.quAllTchFileTQA.AsString := oleExcel.cells[i,13 + j];
          dmBscData.quAllTchFileTSS4.AsString := oleExcel.cells[i,14 + j];
          dmBscData.quAllTchFileTHSI.AsString := oleExcel.cells[i,15 + j];
          dmBscData.quAllTchFileTHSE.AsString := oleExcel.cells[i,16 + j];
          dmBscData.quAllTchFileER_DR.AsString := oleExcel.cells[i,17 + j];
          dmBscData.quAllTchFileTG.AsString := oleExcel.cells[i,18 + j];
          oleExcel.WorkBooks[1].WorkSheets['TCHData'].Activate;
          wCTRALACC := oleExcel.cells[i-2, 24 - j];
          wCNSCAN := oleExcel.cells[i-2, 25 - j];
          wSumCount := wSumCount + 1;
          if wCTRALACC <> '' then
            wSumCTRALCC := wSumCTRALCC + StrToInt(wCTRALACC);
          if wCNSCAN <> '' then
            wSumCNSCAn := wSumCNSCAn + StrToInt(wCNSCAN);
          //if j = 0 then
          //  dmBscData.quAllTchFileTRAFFIC.AsString := oleExcel.cells[i,19 + j]
          // else
          // begin

          if (wCTRALACC <> '') and (wCNSCAN <> '') and (wCNSCAN <> '0') then
            dmBscData.quAllTchFileTRAFFIC.AsFloat := StrToInt(wCTRALACC) / StrToInt(wCNSCAN)
          else
            dmBscData.quAllTchFileTRAFFIC.AsFloat := 0;

          oleExcel.WorkBooks[1].WorkSheets['TCH'].Activate;
          dmBscData.quAllTchFileSTANDARD.AsString := oleExcel.cells[i,20 + j];

          dmBscData.quAllTchFile.Post;
        except
          ShowMessage(wStartDate + '(' + wStartTime +
               '-' + wEndTime + ') ' + wTmpComp + 'TCH的数据有问题!');
        end;
        i := i + 1;
        wTmpComp := oleExcel.cells[i, 2] ;
      end;
      Close;
    end;
    /////////////////////////////////
    with dmBscData.quBscAllTch do
    begin
      if Active then
        Close;
      ParamByName('start_date').AsInteger := StrToInt(wStartDate);
      ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
      //ParamByName('End_Time').AsString := wEndTime;
      ParamByName('Bsc_no').AsString := wBscNo;
      Open;
      if not IsEmpty then
      begin
        if MessageDlg('数据已存在, 是否覆盖?',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          with dmBscData.quDelBscTch do
          begin
            ParamByName('start_date').AsInteger := StrToInt(wStartDate);
            ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
            //ParamByName('End_Time').AsString := wEndTime;
            ParamByName('Bsc_no').AsString := wBscNo;
            ExecSQL;
          end;
        end
        else
        begin
          Continue;
          //Exit;
        end;
      end;
      begin
        try
          dmBscData.quBscAllTch.Insert;
          dmBscData.quBscAllTchSTART_TIME.AsString := wStartTime;
          dmBscData.quBscAllTchEND_TIME.AsString := wEndTime;
          dmBscData.quBscAllTchSTART_DATE.AsString := wStartDate;
          dmBscData.quBscAllTchBSC_NO.AsString :=  wBscNo;
          dmBscData.quBscAllTchTCH.AsString := oleExcel.cells[4,1];
          //dmBscData.quBscAllTchCELL_ID.AsString := oleExcel.cells[i,2];
          dmBscData.quBscAllTchCA.AsString := oleExcel.cells[4,3];
          dmBscData.quBscAllTchCS.AsString := oleExcel.cells[4,4];
          dmBscData.quBscAllTchU.AsString := oleExcel.cells[4,5];
          dmBscData.quBscAllTchERPAC.AsString := oleExcel.cells[4,6];
          dmBscData.quBscAllTchCG.AsString := oleExcel.cells[4,7];
          dmBscData.quBscAllTchMH.AsString := oleExcel.cells[4,8];
          dmBscData.quBscAllTchDR.AsString := oleExcel.cells[4,9 ];
          dmBscData.quBscAllTchCH.AsString := oleExcel.cells[4,10 + j];
          dmBscData.quBscAllTchAC.AsString := oleExcel.cells[4,11 + j];
          dmBscData.quBscAllTchF.AsString := oleExcel.cells[4,12 + j];
          dmBscData.quBscAllTchTQA.AsString := oleExcel.cells[4,13 + j];
          dmBscData.quBscAllTchTSS4.AsString := oleExcel.cells[4,14 + j];
          dmBscData.quBscAllTchTHSI.AsString := oleExcel.cells[4,15 + j];
          dmBscData.quBscAllTchTHSE.AsString := oleExcel.cells[4,16 + j];
          dmBscData.quBscAllTchER_DR.AsString := oleExcel.cells[4,17 + j];
          dmBscData.quBscAllTchTG.AsString := oleExcel.cells[4,18 + j];
          if wSumCNSCAn > 0 then
            dmBscData.quBscAllTchTRAFFIC.AsFloat := wSumCTRALCC / ( wSumCNSCAn / wSumCount) ;
          dmBscData.quBscAllTchSTANDARD.AsString := oleExcel.cells[4,20 + j];
          dmBscData.quBscAllTch.Post;
        except
          ShowMessage(wStartDate + '(' + wStartTime +
            '-' + wEndTime + ') ' + wTmpComp + 'TCH的BSC总数据有问题!');
        end;
      end;
    end;

    wSumCTRALCC := 0;
    wSumCNSCAn := 0;
    wSumCount := 0;
    oleExcel.WorkBooks[1].WorkSheets['CCH'].Activate;
    //showMessage(oleExcel.Cells[10,3]);
    i := 5;
    ///////////////////////////////////////
    with dmBscData.quAllCchFile do
    begin
      if Active then
        Close;
      ParamByName('start_date').AsInteger := StrToInt(wStartDate);
      ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
     // ParamByName('End_Time').AsString := wEndTime;
      ParamByName('Bsc_no').AsString := wBscNo;
      Open;
      if not IsEmpty then
      begin
        if MessageDlg('数据已存在, 是否覆盖?',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          with dmBscData.quDelCch do
          begin
            ParamByName('start_date').AsInteger := StrToInt(wStartDate);
            ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
            //ParamByName('End_Time').AsString := wEndTime;
            ParamByName('Bsc_no').AsString := wBscNo;
            ExecSQL;
          end;
        end
        else
        begin
          Continue;
          //Exit;
        end;
      end;
      i := 5;
      wTmpComp := oleExcel.cells[i, 2] ;
     // showMessage(oleExcel.cells[i, 1]);
      while wTmpComp  <> '' do
      begin
        try
          dmBscData.quAllCchFile.Insert;
          dmBscData.quAllCchFileSTART_TIME.AsString := wStartTime;
          dmBscData.quAllCchFileEND_TIME.AsString := wEndTime;
          dmBscData.quAllCchFileSTART_DATE.AsString := wStartDate;
          dmBscData.quAllCchFileBSC_NO.AsString :=  wBscNo;
          dmBscData.quAllCchFileCCH.AsString := oleExcel.cells[i,1];
          dmBscData.quAllCchFileCELL_ID.AsString := oleExcel.cells[i,2];
          dmBscData.quAllCchFileRAF.AsString := oleExcel.cells[i,3];
          dmBscData.quAllCchFileRAA.AsString := oleExcel.cells[i,4];
          dmBscData.quAllCchFileRAS.AsString := oleExcel.cells[i,5];
          dmBscData.quAllCchFileRAC.AsString := oleExcel.cells[i,6];
          dmBscData.quAllCchFileSA.AsString := oleExcel.cells[i,7];
          dmBscData.quAllCchFileSS.AsString := oleExcel.cells[i,8];
          dmBscData.quAllCchFileSU.AsString := oleExcel.cells[i,9];
          dmBscData.quAllCchFileSC.AsString := oleExcel.cells[i,10];
          dmBscData.quAllCchFileSDR.AsString := oleExcel.cells[i,11];
          dmBscData.quAllCchFileCH.AsString := oleExcel.cells[i,12];
          dmBscData.quAllCchFileAC.AsString := oleExcel.cells[i,13];
          dmBscData.quAllCchFileSF.AsString := oleExcel.cells[i,14];
          dmBscData.quAllCchFileDQA.AsString := oleExcel.cells[i,15];
          dmBscData.quAllCchFileDSS4.AsString := oleExcel.cells[i,16];
          dmBscData.quAllCchFileSTANDARD.AsString := oleExcel.cells[i,17];
          oleExcel.WorkBooks[1].WorkSheets['CCHData'].Activate;
          wCTRALACC := oleExcel.cells[i-2,5];
          wCNSCAN := oleExcel.cells[i-2,6];
          wSumCount := wSumCount + 1;
          if wCTRALACC <> '' then
            wSumCTRALCC := wSumCTRALCC + StrToInt(wCTRALACC);
          if wCNSCAN <> '' then
            wSumCNSCAn := wSumCNSCAn + StrToInt(wCNSCAN);
          if (wCTRALACC <> '') and (wCNSCAN <> '') and (wCNSCAN <> '0') then
            dmBscData.quAllCchFileTRAFFIC.AsFloat := StrToInt(wCTRALACC) / StrToInt(wCNSCAN)
          else
            dmBscData.quAllCchFileTRAFFIC.AsFloat := 0;
          dmBscData.quAllCchFile.Post;
        except
          ShowMessage(wStartDate + '(' + wStartTime +
            '-' + wEndTime + ') ' + wTmpComp + '的CCH数据有问题!');
        end;

        oleExcel.WorkBooks[1].WorkSheets['CCH'].Activate;
        i := i + 1;
        wTmpComp := oleExcel.cells[i, 2] ;
      end;
      Close;
    end;
    /////////////////////////////////////////
    with dmBscData.quBscAllCch do
    begin
      if Active then
        Close;
      ParamByName('start_date').AsInteger := StrToInt(wStartDate);
      ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
     // ParamByName('End_Time').AsString := wEndTime;
      ParamByName('Bsc_no').AsString := wBscNo;
      Open;
      if not IsEmpty then
      begin
        if MessageDlg('数据已存在, 是否覆盖?',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          with dmBscData.quDelBscCch do
          begin
            ParamByName('start_date').AsInteger := StrToInt(wStartDate);
            ParamByName('start_Time').AsInteger := StrToInt(wStartTime);
            //ParamByName('End_Time').AsString := wEndTime;
            ParamByName('Bsc_no').AsString := wBscNo;
            ExecSQL;
          end;
        end
        else
        begin
          Continue;
          //Exit;
        end;
      end;

      begin
        try
          Insert;
          dmBscData.quBscAllCchSTART_TIME.AsString := wStartTime;
          dmBscData.quBscAllCchEND_TIME.AsString := wEndTime;
          dmBscData.quBscAllCchSTART_DATE.AsString := wStartDate;
          dmBscData.quBscAllCchBSC_NO.AsString :=  wBscNo;
          dmBscData.quBscAllCchCCH.AsString := oleExcel.cells[4,1];
          //dmBscData.quAllCchFileCELL_ID.AsString := oleExcel.cells[i,2];
          dmBscData.quBscAllCchRAF.AsString := oleExcel.cells[4,3];
          dmBscData.quBscAllCchRAA.AsString := oleExcel.cells[4,4];
          dmBscData.quBscAllCchRAS.AsString := oleExcel.cells[4,5];
          dmBscData.quBscAllCchRAC.AsString := oleExcel.cells[4,6];
          dmBscData.quBscAllCchSA.AsString := oleExcel.cells[4,7];
          dmBscData.quBscAllCchSS.AsString := oleExcel.cells[4,8];
          dmBscData.quBscAllCchSU.AsString := oleExcel.cells[4,9];
          dmBscData.quBscAllCchSC.AsString := oleExcel.cells[4,10];
          dmBscData.quBscAllCchSDR.AsString := oleExcel.cells[4,11];
          dmBscData.quBscAllCchCH.AsString := oleExcel.cells[4,12];
          dmBscData.quBscAllCchAC.AsString := oleExcel.cells[4,13];
          dmBscData.quBscAllCchSF.AsString := oleExcel.cells[4,14];
          dmBscData.quBscAllCchDQA.AsString := oleExcel.cells[4,15];
          dmBscData.quBscAllCchDSS4.AsString := oleExcel.cells[4,16];
          dmBscData.quBscAllCchSTANDARD.AsString := oleExcel.cells[4,17];
         { oleExcel.WorkBooks[1].WorkSheets['CCHData'].Activate;
          wCTRALACC := oleExcel.cells[i-2,5];
          wCNSCAN := oleExcel.cells[i-2,6];
          if (wCTRALACC <> '') and (wCNSCAN <> '') and (wCNSCAN <> '0') then
            dmBscData.quAllCchFileTRAFFIC.AsFloat := StrToInt(wCTRALACC) / StrToInt(wCNSCAN)
          }
          if wSumCNSCAn > 0 then
            dmBscData.quBScAllCchTRAFFIC.AsFloat := wSumCTRALCC / ( wSumCNSCAn / wSumCount );

          dmBscData.quBScAllCch.Post;
        except
          ShowMessage(wStartDate + '(' + wStartTime +
            '-' + wEndTime + ') ' + wTmpComp + '的CCH BSC总数据有问题!');
        end;

      end;
    end;
    with dmBscData.tbDataHist do
    begin
      //delete
      with dmBscData.quDelDataHist do
      begin
        ParamByName('Start_Date').AsInteger := strToInt(wStartDate);
        ParamByName('Start_Time').AsInteger := StrToInt(wStartTime);
        ExecSQL;
      end;
      if not Active then
        Open;
      Insert;
      FieldByName('Start_Date').AsString := wStartDate;
      FieldByName('Start_Time').AsString := wStartTime;
      FieldByName('End_time').AsString := wEndTime;
      try
        Post;
      except
        Cancel;
      end;
    end;

  end;

  oleExcel.quit;

  sbBscMain.Panels[0].Text := '';
  with dmBscData do
  begin
    quBScAllCch.Close;
    quAllCchFile.Close;
    quBscAllTch.Close;
    quAllTchFile.Close;
  end;
  ShowMessage('BSC数据输入成功!');
end;

procedure TfmBscMain.mmSDCCHClick(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(105);
end;

procedure TfmBscMain.N18Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(106);
end;

procedure TfmBscMain.N1Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(103);
end;

procedure TfmBscMain.N3Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(104);
end;

procedure TfmBscMain.N27Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(609);
end;

procedure TfmBscMain.N30Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(604);
end;

procedure TfmBscMain.N29Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(111);
end;

procedure TfmBscMain.N31Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(112);
end;

procedure TfmBscMain.N22Click(Sender: TObject);
begin
   oleMapInfo.do('Set Map Layer 0 Editable Off');
   
end;

procedure TfmBscMain.N23Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(809);
end;

procedure TfmBscMain.N10Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(810);
end;

procedure TfmBscMain.N13Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(114);
end;

procedure TfmBscMain.N14Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(115);
end;

procedure TfmBscMain.SpeedButton21Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1712);
end;

procedure TfmBscMain.SpeedButton20Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1713);
end;

procedure TfmBscMain.SpeedButton19Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1716);
end;

procedure TfmBscMain.SpeedButton23Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1714);
end;

procedure TfmBscMain.SpeedButton24Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1715);
end;

procedure TfmBscMain.SpeedButton22Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1717);
end;

procedure TfmBscMain.SpeedButton8Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(1718);
end;

procedure TfmBscMain.N15Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(501);
end;

procedure TfmBscMain.N19Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(503);
end;

procedure TfmBscMain.N32Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(502);
end;

procedure TfmBscMain.N33Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(504);
end;

procedure TfmBscMain.SpeedButton25Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(308);  
end;

procedure TfmBscMain.SpeedButton26Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(307);
end;

procedure TfmBscMain.mmMapConfClick(Sender: TObject);
begin
  Application.CreateForm(TfmMapConf, fmMapConf);
  try
    fmMapConf.ShowModal;
  finally
    fmMapConf.Free;
  end;
end;

procedure TfmBscMain.N46Click(Sender: TObject);
begin
  oleMapInfo.RunMenuCommand(610);
end;

procedure TfmBscMain.mmAllCddClick(Sender: TObject);
begin
  with dmBscData.dbCtrData do
  begin
    if not Connected then
    try
      Connected := True;
    except
      showMessage('没有连接CTR');
    end;
  end;
  Application.CreateForm(TfmAllCdd, fmAllCdd);
 // try
  fmAllCdd.Show;
 // finally
 //   fmAllCdd.Free;
 // end;
end;

procedure TfmBscMain.mmSelCddClick(Sender: TObject);
var
  i, wTableNum : Integer;

begin


  if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) <> 'CELL' then
  begin
    ShowMessage('请选取一个小区!');
    Exit;
  end;

  wTableNum := oleMapInfo.eval('NumTables()');

  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'NCELL_TMP' then
    begin
      oleMapInfo.do('Close table NCELL_TMP');
      break;
    end;
  end;
  with dmBscData.dbCtrData do
  begin
    if not Connected then
    try
      Connected := True;
    except
      showMessage('没有连接CTR');
    end;
  end;
  Application.CreateForm(TfmSelCdd, fmSelCdd);
  try
    fmSelCdd.ShowModal;
  finally
    fmSelCdd.Free;
  end;
end;

procedure TfmBscMain.mmResetBaseClick(Sender: TObject);
var
  i, wSumCount, wColor : Integer;
  wLon, wLat : Real;
  wSumBscNo, wBscNo : String;
 // wReset : Boolean;
begin
  with tbBsc do
  begin
    wSumBscNo := '';
    if not Active then
      Open;
    wSumCount := RecordCount;
    First;
    i := 0;
    while not eof do
    begin
      i := i + 1;

      wSumBscNo := wSumBscNo + '&' +  FieldByName('Bsc_no').AsString;
      Next;
    end;
  end;
  for i := 1 to oleMapInfo.eval('tableinfo(base,8)') do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from base');
    wLon := oleMapInfo.eval('base.lon');
    wLat := oleMapInfo.eval('base.lat');
    wBscNo := oleMapInfo.eval('base.bsc_no');
    //wColor := Pos(wBscNo, wSumBscNo);
    wColor := Round(16437256 * Pos(wBscNo, wSumBscNo) / Length(wSumBscNo));
    oleMapInfo.do('set style Symbol makeSymbol(33, ' + IntToStr(wColor)
                 + ',10)');
   { if oleMapInfo.eval('base.bsc_no') = 'BSC1B' then
       oleMapInfo.do('set style Symbol makeSymbol(33,6437256,10)');
    if oleMapInfo.eval('base.bsc_no') = 'BSC1C' then
       oleMapInfo.do('set style Symbol makeSymbol(33,14237256,10)');
    if oleMapInfo.eval('base.bsc_no') = 'BSC2A' then
       oleMapInfo.do('set style Symbol makeSymbol(33,16437256,10)');
    if oleMapInfo.eval('base.bsc_no') = 'BSC2B' then
       oleMapInfo.do('set style Symbol makeSymbol(33,737256,10)');
    if oleMapInfo.eval('base.bsc_no') = 'BSC2C' then
       oleMapInfo.do('set style Symbol makeSymbol(33,337256,10)'); }
   // wBearing := oleMapInfo.eval('tch_tss4.Bearing');
   // wRate := oleMapInfo.eval('tch_tss4.tss4') / 100;
    //if wRate > 0 then
    //  wRate := 0.01 + 0.05 * wRate;
    oleMapInfo.do('update base set obj = createPoint(' + FloatToStr(wLon) +
                  ',' + FloatToStr(wLat) + ') where rowid = ' + IntToStr(i));
    //UpDateCircle(wLon, wLat, 0.0025, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'tch_tss4');
  end;
  oleMapInfo.do('commit table base');
end;

procedure TfmBscMain.mmResetCellClick(Sender: TObject);
var
  wReset : Boolean;
begin
  wReset := False;
  Application.CreateForm(TfmCellConf, fmCellConf);
  try
    fmCellConf.edCellLength.Text := FloatToStr(gCellLength);
    fmCellConf.edCellAngle.Text := FloatToStr(gCellAngle);
    if (fmCellConf.ShowModal = mrOK) then
    begin
      gCellAngle := StrToFloat(Trim(fmCellConf.edCellAngle.Text));
      gCellLength := StrToFloat(Trim(fmCellConf.edCellLength.Text));
      with dmBscData.tbBscControl do
      begin
        if not Active then
          Open;
        First;
        Edit;
        FieldByName('Cell_Angle').AsFloat := gCellAngle;
        FieldByName('Cell_Length').AsFloat := gCellLength;
        Post;
        Close; 
      end;
      wReset := True;
    end;
  finally
    fmCellConf.Free;
  end;
  if not wReset then
    Exit;
  UpdateCellObject;
  oleMapInfo.do('commit table cell');
  try
    dmBscData.tbCellRemote.EmptyTable;
    with dmBscData.bmCell do
    begin
      Execute;
    end;
  except
    ShowMessage('ORACLE 数据连接失败!');
  end;

end;

procedure TfmBscMain.SpeedButton27Click(Sender: TObject);
begin
  if gCompLayer <> '' then
  begin
    oleMapInfo.do('close table ' + gCompLayer);
    gCompLayer := '';
  end;
end;

procedure TfmBscMain.N53Click(Sender: TObject);
begin
  Application.CreateForm(TfmFindNcell, fmFindNCell);
  try
    fmFindNcell.ShowModal;
  finally
    fmFindNCell.Free;
  end;
end;

procedure TfmBscMain.N39Click(Sender: TObject);
var
  wLon1, wLat1, wBearing1, wLon2, wLat2, wBearing2, wLength : real;
  i, wRow, wNum, wArfcn, wTableNum : Integer;
  wCell : String;
begin
  if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) <> 'CELL' then
  begin
    ShowMessage('请选取一个小区!');
    Exit;
  end;

  wTableNum := oleMapInfo.eval('NumTables()');
  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'SAME_ARFCN' then
      oleMapInfo.do('Close table same_arfcn');
  end;

  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');
  wNum := 0;
  oleMapInfo.do('Set Style pen makepen(1.2 ,5, RGB(255,0,0))');
  oleMapInfo.do('fetch rec  1 from selection');
  wLon1 := oleMapInfo.eval('selection.Lon');
  wLat1 := oleMapInfo.eval('selection.Lat');
  wBearing1 := oleMapInfo.eval('selection.Bearing');
  wCell := oleMapInfo.eval('selection.bs_no');
  wArfcn := oleMapInfo.eval('selection.arfcn');
  //oleMapInfo.do('Open table "' + gExePath + 'ncell.tab" Interactive');
  oleMapInfo.do('select * from cell where bs_no <> "' + wCell + '" and arfcn = '
                + IntTostr(wArfcn) + ' into tmp');
  oleMapInfo.do('commit table tmp as "' + gExePath + 'same_arfcn.tab"');
  oleMapInfo.do('Open table "' + gExePath + 'same_arfcn.tab" Interactive');
  oleMapInfo.do('close table tmp');

  oleMapInfo.do('Create Map For same_arfcn CoordSys Earth Projection 1, 0');
  wRow := oleMapInfo.eval('tableinfo(same_arfcn,8)');
  oleMapInfo.do('fetch first from same_arfcn');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from same_arfcn');

    wLon2 := oleMapInfo.eval('same_arfcn.Lon');
    wLat2 := oleMapInfo.eval('same_arfcn.Lat');
    //wRate := oleMapInfo.eval('Hdov_i_bsc.hovercnt') / wMaxQty ;
    wBearing2 := oleMapInfo.eval('same_arfcn.Bearing');
    if sqrt(sqr((wLon1 - wLon2) * 103.1) + sqr((wLat1 - wLat2) * 111.2)) <= 5 then
    begin
      fmBscMain.DrawLine(wLon1, wLat1, wLon2, wLat2, gCellLength, wBearing1, wBearing2, oleMapInfo, i, 'same_arfcn');
      wNum := wNum + 1;
    end
  end;
  oleMapInfo.do('commit table same_arfcn');
  oleMapInfo.do('add map auto layer same_arfcn');
  oleMapInfo.do('Set Map Layer same_arfcn Label Position Above Font ("Arial",256,10,208,16777215) ' +
                ' With arfcn ' +
                ' Auto On Visibility Zoom (0, 100) Units "km"');
  if wNum > 0 then
    ShowMessage('共有同频小区：' + IntToStr(wNum) + '个')
  else
    ShowMessage('没有同频小区!');
 // oleMapInfo.do('Set Map  Center (selection.lon, selection.lat)');

end;

procedure TfmBscMain.N40Click(Sender: TObject);
var
  wLon1, wLat1, wBearing1, wLon2, wLat2, wBearing2, wLength : real;
  i, wRow, wNum, wArfcn, wTableNum : Integer;
  wCell : String;
begin
  if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) <> 'CELL' then
  begin
    ShowMessage('请选取一个小区!');
    Exit;
  end;
  wTableNum := oleMapInfo.eval('NumTables()');
  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'NEIGHBOUR_ARFCN' then
      oleMapInfo.do('Close table neighbour_arfcn');
  end;
  oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.3 Units "km"');
  wNum := 0;
  oleMapInfo.do('Set Style pen makepen(2 ,59, RGB(255,0,255))');
  oleMapInfo.do('fetch rec 1 from selection');
  wLon1 := oleMapInfo.eval('selection.Lon');
  wLat1 := oleMapInfo.eval('selection.Lat');
  wBearing1 := oleMapInfo.eval('selection.Bearing');
  wCell := oleMapInfo.eval('selection.bs_no');
  wArfcn := oleMapInfo.eval('selection.arfcn');
  //oleMapInfo.do('Open table "' + gExePath + 'ncell.tab" Interactive');
  oleMapInfo.do('select * from cell where bs_no <> "' + wCell + '" and ( arfcn = '
                + IntTostr(wArfcn + 1) + ' or arfcn = ' + IntToStr(wArfcn + 1) + ') into tmp');
  oleMapInfo.do('commit table tmp as "' + gExePath + 'neighbour_arfcn.tab"');
  oleMapInfo.do('Open table "' + gExePath + 'neighbour_arfcn.tab" Interactive');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('Create Map For neighbour_arfcn CoordSys Earth Projection 1, 0');
  wRow := oleMapInfo.eval('tableinfo(neighbour_arfcn,8)');
  oleMapInfo.do('fetch first from neighbour_arfcn');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from neighbour_arfcn');

    wLon2 := oleMapInfo.eval('neighbour_arfcn.Lon');
    wLat2 := oleMapInfo.eval('neighbour_arfcn.Lat');
    //wRate := oleMapInfo.eval('Hdov_i_bsc.hovercnt') / wMaxQty ;
    wBearing2 := oleMapInfo.eval('neighbour_arfcn.Bearing');
    if sqrt(sqr((wLon1 - wLon2) * 103.1) + sqr((wLat1 - wLat2) * 111.2)) <= 5 then
    begin
      fmBscMain.DrawLine(wLon1, wLat1, wLon2, wLat2, gCellLength, wBearing1, wBearing2, oleMapInfo, i, 'neighbour_arfcn');
      wNum := wNum + 1;
    end
  end;
  oleMapInfo.do('commit table neighbour_arfcn');
  oleMapInfo.do('add map auto layer neighbour_arfcn');
  oleMapInfo.do('Set Map Layer neighbour_arfcn Label Position Above Font ("Arial",256,10,208,16777215) ' +
                ' With bs_no+"("+arfcn+")"' +
                ' Auto On Visibility Zoom (0, 100) Units "km"');
  if wNum > 0 then
    ShowMessage('共有邻频小区：' + IntToStr(wNum) + '个')
  else
    ShowMessage('没有邻频小区!');
 // oleMapInfo.do('Set Map  Center (selection.lon, selection.lat)');

end;


procedure TfmBscMain.SpeedButton6Click(Sender: TObject);
begin
  if gSelFlag <> 'BSC' then
  begin
    if UpperCase(oleMapInfo.eval('selectionInfo(1)')) <> 'CELL' then
    begin
      ShowMessage('请选取一个小区或BSC!');
      Exit;
    end;
  end;
  oleMapInfo.do('close table all_cch_file');
  oleMapInfo.do('close table all_Tch_file');
  Application.CreateForm(TfmTrendline, fmTrendline);
  try
    fmTrendline.ShowModal;
  finally
    fmTrendline.Free;
  end;
  oleMapInfo.do('Open table "' + gExePath + 'all_cch_file.tab" Interactive');
  oleMapInfo.do('Open table "' + gExePath + 'all_Tch_file.tab" Interactive');
end;

procedure TfmBscMain.N34Click(Sender: TObject);
var
  wStr : array [0..60] of char;
  wPath : String;
begin
  wPath := '' + gCtrExePath + 'ctrproject.exe';
  strpcopy(wStr, wPath);
  if WinExec(wStr, SW_SHOW) < 32 then
  begin
    ShowMessage('CTRProject.EXE不存在！');
  end;
end;

procedure TfmBscMain.cbNextChange(Sender: TObject);
begin
  
end;
{var
  i, wRow, wNcellRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wPenWidth : real;
begin

  CreateMDIChild(cbNext.Text);


  oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell ,' + wPriorCchTable
                 + ' where cch_file.cell_id = cell.bs_no and cell.bs_no = ' + wPriorCchTable + '.cell_id'
                 + ' into TrafficTmp ');


  oleMapInfo.do('commit table TrafficTmp as "' + gExePath + 'cch_Traffic_next.tab"');
  oleMapInfo.do('close table TrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_Traffic_next.tab" Interactive');

  oleMapInfo.do('select Max(Traffic) from cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(Traffic) from Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_Traffic_next, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_Traffic_next');
    wLon := oleMapInfo.eval('cch_Traffic_next.lon');
    wLat := oleMapInfo.eval('cch_Traffic_next.lat');
    wBearing := oleMapInfo.eval('cch_Traffic_next.Bearing');
    wRate := oleMapInfo.eval('cch_Traffic_next.Traffic') / wMaxQty;
    wPenWidth := 1 + oleMapInfo.eval('cch_Traffic_next.Sc');
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    if wRate > 0 then
      wRate := 0.02 + 0.05 * wRate;
    UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'cch_Traffic_next');
  end;
  oleMapInfo.do('commit table cch_Traffic_next');
  oleMapInfo.do('add map auto layer cch_Traffic_next');
  oleMapInfo.do('Set Map Layer cch_Traffic_next Label Position Above Font ("Arial",0,10,0) With Traffic+","+sc+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Remove Map Layer ' + wFristCchTable + ' Interactive ');
  ////tch

  oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell, ' +  wPriorTchTable
                 + ' where tch_file.cell_id = cell.bs_no  and  cell.bs_no = ' + wPriorTchTable + '.cell_id'
                 + ' into TrafficTmp');

  oleMapInfo.do('commit table TrafficTmp as "' + gExePath + 'tch_Traffic_next.tab"');
  oleMapInfo.do('close table TrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_Traffic_next.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_Traffic_next, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_Traffic_next');
    wLon := oleMapInfo.eval('tch_Traffic_next.lon');
    wLat := oleMapInfo.eval('tch_Traffic_next.lat');
    wBearing := oleMapInfo.eval('tch_Traffic_next.Bearing');
    wRate := oleMapInfo.eval('tch_Traffic_next.Traffic') / wMaxQty;
    wPenWidth := 1 + oleMapInfo.eval('tch_Traffic_next.cg');
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    if wRate > 0 then
      wRate := 0.02 + 0.05 * wRate;
    UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'tch_Traffic_next');
  end;
  oleMapInfo.do('commit table tch_Traffic_next');
  oleMapInfo.do('add map auto layer tch_Traffic_next');
  oleMapInfo.do('Set Map Layer tch_Traffic_next Label Position Above Font ("Arial",0,10,0) With Traffic+","+cg+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  oleMapInfo.do('set map redraw on');
  oleMapInfo.do('Remove Map Layer ' + wFristTchTable + ' Interactive ');
  //gTraffic := True;

  //showMessage(IntToStr(oleMapInfo.eval('numwindows()')));

end;  }

procedure TfmBscMain.SpeedButton29Click(Sender: TObject);
var
  wCell : String;
begin
  if not wBrowseShow then
  begin
    Application.CreateForm(TfmBrowse, fmBrowse);
    with fmBrowse do
    begin
      with quAllSelCell do
      begin
        Sql.Clear;
        Sql.Add('delete  from all_sel_cell ');
        ExecSQL;

        if not gHover then
        begin
          if gMh then
          begin
            Sql.Clear;
            Sql.Add('delete from cch_sel_cell');
            ExecSQL;
          end;
          if gRaf then
          begin
            Sql.Clear;
            Sql.Add('delete from tch_sel_cell');
            ExecSQL;
          end;
          if gCqt then
          begin
            Sql.Clear;
            Sql.Add('delete from cch_sel_cell');
            ExecSQL;
          end;

          Sql.Clear;
          Sql.Add('insert into tch_sel_cell select cell_id from cch_sel_cell');
          ExecSQL;


          Sql.Clear;
          Sql.Add('insert into All_sel_cell select cell_id from tch_sel_cell group by cell_id');
          ExecSQL;

        end;
      end;

      with  quHdovi do
      begin
        if not Active then
          Open;
        First;
        if gHover then
        begin
          with quAllSelCell do
          begin
            Sql.Clear;
            Sql.Add('delete  from tch_sel_cell ');
            ExecSQL;

          end;
          if not quAllCell.Active then
            quAllCell.Open;
          while not eof do
          begin
            quAllCell.Insert;
            quAllCell.FieldByName('cell_id').AsString :=
              FieldByName('Cell_id_se').AsString;
            quAllCell.Post;
            Next;
          end;
         // quAllCell.Close;
        end;
        First;
      end;

      with quHdove do
      begin
        if not Active then
          Open;
        First;
        if gHover then
        begin
          if not quAllCell.Active then
            quAllCell.Open;
          while not eof do
          begin
            quAllCell.Insert;
            quAllCell.FieldByName('cell_id').AsString :=
              FieldByName('Cell_id_se').AsString;
            quAllCell.Post;
            Next;
          end;
          quAllCell.Close;
          with quAllSelCell do
          begin
            Sql.Clear;
            Sql.Add('insert into All_sel_cell select cell_id from tch_sel_cell group by cell_id');
            ExecSQL;
          end;
        end;
        First;

      end;

      with quTchBrowse do
      begin
        if Active then
          Close;
        Open;
      end;
      with quCchBrowse do
      begin
        if Active then
          Close;
        Open;
      end;

      if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'CELL' then
      begin
        oleMapInfo.do('fetch rec ' + IntToStr(1) +' from selection');
        wCell := oleMapInfo.eval('selection.bs_no');
        if fmBrowse.quTchBrowse.Active then
          fmBrowse.quTchBrowse.Locate('cell_id',UpperCase(wCell),[loPartialKey]);
        if quCchBrowse.Active then
          fmBrowse.quCchBrowse.Locate('cell_id',UpperCase(wCell),[loPartialKey]);
      end;
      Show;
    end;

    wBrowseShow := True;
  end;
end;

procedure TfmBscMain.mmDssClick(Sender: TObject);
var
  i, wRow, wNcellRow , wCondNum : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty : real;
begin
//  fmBscMain.ResetMap;
  fmLegeng.Hide;
  if gDss then
  begin
    if mmDss.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      //oleMapInfo.do('Set Map Layer cch_Dqa Display Off');
      //oleMapInfo.do('Set Map Layer Tch_Tqa Display off');
      oleMapInfo.do('Set Map Layer cch_Dss4 Display Off');
      oleMapInfo.do('Set Map Layer Tch_Tss4 Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      //oleMapInfo.do('Set Map Layer cch_Dqa Display Graphic');
      //oleMapInfo.do('Set Map Layer Tch_Tqa Display Graphic');
      oleMapInfo.do('Set Map Layer cch_Dss4 Display Graphic');
      oleMapInfo.do('Set Map Layer Tch_Dss4 Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmDss.Checked := not mmDss.Checked;
    exit;
  end;

  mmDss.Checked := not mmDss.Checked;
  oleMapInfo.do('set map redraw off');
   sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  if  (gSelFlag <> 'CELL') then
    ShowCondition('弱信号断线')
  else
    wFilterOrder := '';

  //弱信号断线
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wSdcchCheck = 'Y' then
      begin
        wCondition := ' and cch_file.Dss4 ' + wSdcchFlag + ' ' + wSdcchQty;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into dss4Tmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into dss4Tmp');
    end
    else
    //order by
    begin
      if wSdcchCheck = 'Y' then
      begin
        if wSdcchFlag = '升序' then
          wCondition := ' order by cch_file.dss4 '
        else
          wCondition := ' order by cch_file.dss4 desc ' ;
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition  + ' into dss4Tmp');
      end
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into dss4Tmp');
    end;
    {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into dss4Tmp');  }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into dss4Tmp')
      else
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into dss4Tmp');
    end
    else
    begin   //all
      if wFilterOrder = 'FILTER' then
      begin
        if wSdcchCheck = 'Y' then
        begin
          wCondition := ' and cch_file.dss4 ' + wSdcchFlag + ' ' + wSdcchQty;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no ' + wCondition + ' into dss4Tmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into dss4Tmp');
      end
      else
      //order by
      begin
        if wSdcchCheck = 'Y' then
        begin
          if wSdcchFlag = '升序' then
            wCondition := ' order by cch_file.dss4 '
          else
            wCondition := ' order by cch_file.dss4 desc ' ;
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into dss4Tmp');
        end
        else
          oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = "0" into dss4Tmp');
      end;

      {oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into dss4Tmp');}
    end;
  end;


  oleMapInfo.do('commit table dss4Tmp as "' + gExePath + 'cch_dss4.tab"');
  oleMapInfo.do('close table dss4Tmp ');
  oleMapInfo.do('Open table "' + gExePath + 'cch_dss4.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(cch_dss4, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wSdcchQty)) then
    wCondNum := StrToInt(wSdcchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from cch_dss4');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from cch_dss4 where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('cch_dss4.lon');
      wLat := oleMapInfo.eval('cch_dss4.lat');
      wBearing := oleMapInfo.eval('cch_dss4.Bearing');
      wRate := oleMapInfo.eval('cch_dss4.dss4') / 100;
      if wRate > 0 then
        wRate := 0.005 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('cch_dss4.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'cch_dss4')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 1, 'cch_dss4');
    end;
  end;
  oleMapInfo.do('commit table cch_dss4');
  oleMapInfo.do('add map auto layer cch_dss4');
  oleMapInfo.do('Set Map Layer cch_dss4 Label Position Above Font ("Arial",256,8,16777215,0) ' +
               ' With dss4+"%" Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer cch_dss4');

  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and tch_file.tss4 ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" '  + wCondition + ' into tss4Tmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into tss4Tmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by tch_file.tss4 '
        else
          wCondition := ' order by tch_file.tss4 desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into tss4Tmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "0" into tss4Tmp');
    end;


    {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into tss4Tmp');}
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
         oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into tss4Tmp')
      else
         oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + 'into tss4Tmp');
    end
    else
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and Tch_file.Tss4 ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into tss4Tmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into tss4Tmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.Tss4 '
          else
            wCondition := ' order by tch_file.Tss4 desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition +  ' into tss4Tmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into tss4Tmp');
      end;

      {oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into tss4Tmp'); }
    end;
  end;

  oleMapInfo.do('commit table tss4Tmp as "' + gExePath + 'tch_tss4.tab"');
  oleMapInfo.do('close table tss4Tmp ');
  oleMapInfo.do('Open table "' + gExePath + 'tch_tss4.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(tch_tss4, 8)');
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
    wCondNum := StrToInt(wTchQty);
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from tch_tss4');
    if (i > wCondNum) and (wFilterOrder = 'ORDER') then
    begin
      oleMapInfo.do('Delete from Tch_Tss4 where rowid = ' + IntToStr(i));
    end
    else
    begin
      wLon := oleMapInfo.eval('tch_tss4.lon');
      wLat := oleMapInfo.eval('tch_tss4.lat');
      wBearing := oleMapInfo.eval('tch_tss4.Bearing');
      wRate := oleMapInfo.eval('tch_tss4.tss4') / 100;

      if wRate > 0 then
        wRate := 0.005 + 0.05 * wRate;
      if Pos('5', oleMapInfo.eval('tch_tss4.cell_id')) > 0 then
        UpDateCircle(wLon, wLat, gCellLength/2, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'tch_tss4')
      else
        UpDateCircle(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, i, 2, 'tch_tss4');
    end;
  end;
  oleMapInfo.do('commit table tch_tss4');
  oleMapInfo.do('add map auto layer tch_tss4');
  oleMapInfo.do('Set Map Layer tch_tss4 Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With tss4+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  oleMapInfo.do('set map redraw on');
  //gDqa := True;
  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
  oleMapInfo.do('select  cell_id from tch_Tss4 into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select  cell_id from Cch_dss4 into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Cch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  gDss := True;
  fmLegeng.Show;
  fmLegeng.nbLegeng.PageIndex := 7;
end;


procedure TfmBscMain.mmMHClick(Sender: TObject);
procedure mh_shade;
var
  wCondNum , wRow, i : Integer;
  wCellid : string;
  wLon , wLat, wBearing, wRate : Real;
begin
  if  (gSelFlag <> 'CELL') then
    ShowCondition('平均通话时间')
  else
    wFilterOrder := '';
  if gSelFlag = 'BSC' then
  begin
    if wFilterOrder = 'FILTER' then
    begin
      if wTchCheck = 'Y' then
      begin
        wCondition := ' and Tch_file.mh ' + wTchFlag + ' ' + wTchQty;
        oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into Tmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
    end
    else
    //order by
    begin
      if wTchCheck = 'Y' then
      begin
        if wTchFlag = '升序' then
          wCondition := ' order by Tch_file.mh '
        else
          wCondition := ' order by Tch_file.mh desc ' ;
        oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "'
                 + gSelName + '" ' + wCondition + ' into Tmp');
      end
      else
        oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.Bsc_no = "0" into Tmp');
    end;
    {oleMapInfo.do('Select * from cell where cell.bsc_no = "'
                 + gSelName + '" into tmp');  }
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select * from cell where cell.bs_no = "'
                 + gSelName + '" into tmp')
      else
        oleMapInfo.do('Select * from cell where bs_no > "0"  and ' +
                 GetMultiCell(gMultiCell) + ' into tmp');

    end
    else //all
    begin
      if wFilterOrder = 'FILTER' then
      begin
        if wTchCheck = 'Y' then
        begin
          wCondition := ' and tch_file.mh ' + wTchFlag + ' ' + wTchQty;
          oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no ' + wCondition + ' into Tmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0"  into Tmp');
      end
      else
      //order by
      begin
        if wTchCheck = 'Y' then
        begin
          if wTchFlag = '升序' then
            wCondition := ' order by tch_file.mh '
          else
            wCondition := ' order by tch_file.mh desc ' ;
          oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no  ' + wCondition + ' into Tmp');
        end
        else
          oleMapInfo.do('Select TCh_file.Cell_id, tch_file.mh, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = "0" into Tmp');
      end;
      //oleMapInfo.do('Select * from cell into tmp');
    end;
  end;
  if gSelFlag = 'CELL' then
  begin
    oleMapInfo.do('Add Column "hm_shade" (hm Decimal (8, 2))From cch_file Set To raf Where COL2 = COL6  Dynamic');
  end;

  oleMapInfo.do('commit table tmp as "' + gExePath + 'Mh_shade.tab"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('open table "' + gExePath + 'Mh_shade.tab"');
  wRow := oleMapInfo.eval('tableinfo(mh_shade, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from mh_shade');
    wLon := oleMapInfo.eval('mh_shade.lon');
    wLat := oleMapInfo.eval('mh_shade.lat');
    wBearing := oleMapInfo.eval('mh_shade.Bearing');
    wRate := 0;
    CreateRegion_5(wLon, wLat, gCellLength, gCellAngle, wBearing, wRate, oleMapInfo, 2);
    oleMapInfo.do('Update mh_shade set Obj = TmpObject where rowId = ' + IntToStr(i));
  end;
  if (wFilterOrder = 'ORDER') and (wRow > StrToInt(wTchQty)) then
  begin
    wCondNum := StrToInt(wTchQty);
    for i := wCondNum to wRow do
      oleMapInfo.do('delete from mh_shade  where rowid = ' + IntToStr(i) );
   // wRow := wCondNum;
  end;
  //wRow := oleMapInfo.eval('tableinfo(raf_shade, 8)');
  //oleMapInfo.do('fetch first from raf_shade');

  oleMapInfo.do('commit table mh_shade');
  oleMapInfo.do('add map auto layer mh_shade');

 { oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' raf_shade  with Arfcn ' +
                'ignore 0 ranges apply all use color Brush (2,16711680,16777215) '+
                '  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,25: 50 Brush ' +
                ' (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush (2,11554816,16777215) ' +
                ' Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) Pen (1,2,0) ' +
                ' default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.1 ' +
                ' inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
 }
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
        ' mh_shade with Mh bar Max Size 1 Units "cm" At Value 50 vary size ' +
        ' by "CONST" border Pen (1,2,0) Width 0.1 Units "cm"  position center ' +
        ' center style Brush (2,16752848,16777215)  # max 50 color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
        ' layer prev display on shades on symbols off lines off count off ' +
        ' title auto Font ("Arial",0,12,0) subtitle auto Font ("Arial",0,11,0) ' +
        '  ascending on ranges Font ("Arial",0,11,0)  auto display off ,auto display on ');
  {oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' mh_shade with mh ignore 0 ranges apply all use color Brush ' +
             ' (2,65280,16777215)  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             ' 25: 50 Brush (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush ' +
             ' (2,11554816,16777215) Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) ' +
             ' Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 ' +
             ' round 0.1 inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
   ' layer prev display on shades on symbols off lines off count on title auto '  +
   ' Font ("Arial",0,12,0) subtitle auto Font ("Arial",0,11,0) ascending off ' +
   ' ranges Font ("Arial",0,11,0) auto display off ,auto display on ,auto ' +
   ' display on ,auto display on ,auto display on '); }

  oleMapInfo.do('Set Map Layer mh_shade Label Position Above Font ("Arial",1,10,0) With mh Auto On Visibility Zoom (0, 6) Units "km"');
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer mh_shade');
  gMh := True;
  oleMapInfo.do('select  cell_id from  mh_shade into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
end;
begin
//  fmBscMain.ResetMap;
  if gMh then
  begin
    if mmMh.Checked then
    begin
      oleMapInfo.do('close table mh_shade');
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Off');
      //oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on')   }
    end
    else
    begin
      Mh_shade;
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Graphic');
     // oleMapInfo.do('Set Map Layer RafTmp Display Graphic');
      oleMapInfo.do('set map redraw on')}
    end;
    mmMh.Checked := not mmMh.Checked;
    exit;
  end;
  mmMh.Checked := not mmMh.Checked;
  //oleMapInfo.do('set map redraw off');
  sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  Mh_shade;

  sbBscMain.Panels[0].Text := 'BSC分析 -- ' + TMenuItem(Sender).Caption;
end;

procedure TfmBscMain.CDD2Click(Sender: TObject);
{var
  wStr : array [0..60] of char;
  wPath : String;}
begin
  {wPath := '' + gExePath + 'Cdd.exe';
  strpcopy(wStr, wPath);
  if WinExec(wStr, SW_SHOW) < 32 then
  begin
    ShowMessage('CDD.EXE不存在！');
  end; }
  ConvertCDD;
end;

procedure TfmBscMain.FormShow(Sender: TObject);
begin
//  if dmBscData.dbCtrData.AliasName = 'NqiData' then
        //sbBscMain.Panels[4].Text := '本地数据';
end;

procedure TfmBscMain.mmDensityShadeClick(Sender: TObject);
Procedure DensityShade;
begin
  oleMapInfo.do('commit table cell as "' + gExePath + 'Density_shade.tab"');
  oleMapInfo.do('open table "' + gExePath + 'Density_shade.tab"');
  oleMapInfo.do('add map auto layer Density_shade');
  oleMapInfo.do('Add Column "Density_shade" (Erpac Decimal (8, 2))From Tch_file Set To Erpac Where COL2 = COL6  Dynamic');
  //oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  {oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' Density_shade with Erpac ignore 0 ranges apply all use color ' +
             ' Brush (2,65280,16777215)  0: 0.199 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             '0.2: 0.499 Brush (2,5287936,16777215) Pen (1,2,0) ,0.5: 0.699 ' +
             'Brush (2,11554816,16777215) Pen (1,2,0) ,0.7: 100 ' +
             'Brush (2,16711680,16777215) Pen (1,2,0) default ' +
             'Brush (2,65280,16777215) Pen (1,2,0)  # use 0 round 0.01 inflect off ' +
             'Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
    ' layer prev display on shades on symbols ' +
    'off lines off count on title auto Font ("Arial",0,12,0) subtitle auto ' +
    'Font ("Arial",0,11,0) ascending off ranges Font ("Arial",0,11,0) auto ' +
    'display off ,auto display on ,auto display on ,auto display on ,auto ' +
    'display on'); }
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
               ' Density_shade with Erpac ranges apply all use all Brush ' +
               ' (2,16777215,16777215)  0: 0.2 Brush (2,16777215,16777215) ' +
               ' Pen (1,2,0) ,0.2: 0.8 Brush (2,65280,16777215) Pen (1,2,0) ' +
               ' ,0.8: 1 Brush (2,16711680,16777215) Pen (1,2,0) default Brush ' +
               ' (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.001 inflect' +
               ' off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
               ' layer prev display on shades on symbols off lines off count ' +
               ' on title "话务密度分析" Font ("Arial",0,14,0) subtitle "每线话务量" ' +
               ' Font ("Arial",0,11,0) ascending off ranges Font ("Arial",0,11,0)' +
               '  auto display off ,auto display on ,auto display on ,auto display off');

  oleMapInfo.do('Set Map Layer Density_shade Label Position Above Font ("Arial",1,10,0) With erpac Auto On Visibility Zoom (0, 6) Units "km"');
  gTchTraffic := True;
  oleMapInfo.do('commit table Density_shade');
  oleMapInfo.do('select cell_id  from tch_file into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'yTch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
end;
begin
  if gTchTraffic then
  begin
    if mmTraShade.Checked then
    begin
      oleMapInfo.do('close table Density_shade');

    end
    else
    begin
      DensityShade;
    end;
  end
  else
    DensityShade;
  mmDensityShade.Checked := not mmDensityShade.Checked;
  if mmDensityShade.Checked {and mmCgShade.Checked} then
  begin
    //mmCgShadeClick(self);
  end;
end;

procedure TfmBscMain.SpeedButton28Click(Sender: TObject);
begin
  Application.CreateForm(TfmDensityDlg, fmDensityDlg);
  try
    fmDensityDlg.ShowModal;
  finally
    fmDensityDlg.Free;
  end;
end;

procedure TfmBscMain.mmLoadMapClick(Sender: TObject);
var
  wTableNum, i : Integer;
  wTableName : String;
  wHasCell, wHasBase : Boolean;
begin
  wHasCell := False;
  wHasBase := False;
  oleMapInfo.RunMenuCommand(102);
  wTableNum := oleMapInfo.eval('NumTables()');
  for i := 1 to wTableNum do
  begin
    wTableName := UpperCase(oleMapInfo.eval('TableInfo(' + IntToStr(i) +  ', 1)'));
    oleMapInfo.do('Commit table ' + wTableName + ' as "' + gExePath + wTableName + '.tab"');
    if (wTableName = 'CELL') then
    begin
      wHasCell := True;
    end;
    if (wTableName = 'BASE') then
    begin
     // oleMapInfo.do('Close table BASE');
      wHasBase := True;
    end;
  end;
  oleMapInfo.RunMenuCommand(104);
  //oleMapInfo.do('Close table CELL');
  if wHasCell then
  begin
    oleMapInfo.do('Open table "' + gExePath + 'cell.tab" Interactive');
    oleMapInfo.do('Alter Table "Cell" ( add bsc_no Char(6) ) Interactive');
  end
  else
    ShowMessage('请加载CELL.TAB');
  if wHasBase then
  begin
    oleMapInfo.do('Open table "' + gExePath + 'BASE.tab" Interactive');
    oleMapInfo.do('Alter Table "BASE" ( add bsc_no Char(6) ) Interactive');
  end
  else
    ShowMessage('请加载BASE.TAB');
  if wHasCell or wHasBase then
    oleMapInfo.RunMenuCommand(104);
  ShowMessage('加载成功!');
  mmUpdateData.Enabled := True;
end;

procedure TfmBscMain.mmUpdateDataClick(Sender: TObject);
var
  oleExcel  : Variant;
  i, j, p, wRowNum : Integer;
  wCellID, wCellName, wBsc : String;
begin
  sbBscMain.Panels[0].Text := '更新小区...';
  if not odData.Execute then
    exit;
  oleExcel := CreateOleObject('Excel.Application');
  oleExcel.WorkBooks.Open(odData.FileName);
  oleExcel.WorkSheets[1].Activate;
  wCellId := oleExcel.cells[2, 1];
  dmBscData.quDelBscFile.ExecSql;
  i := 2;
  if not dmBscData.quBscFile.Active then
    dmBscData.quBscFile.Open;
  while wCellID <> '' do
  begin
    with dmBscData.quBscFile do
    begin;
      Insert;
      FieldByName('Cell_id').AsString := oleExcel.cells[i, 1];
      FieldByName('Cell_name').AsString := oleExcel.cells[i, 2];
      FieldByName('bsc_no').AsString := oleExcel.cells[i, 3];
      Post;
    end;

    i := i + 1;
    wCellId := oleExcel.cells[i, 1];

  end;
  dmBscData.quBscFile.Close;

  oleExcel.Quit;
  with dmBscData.quInsertBsc do
  begin
    Sql.Clear;
    Sql.Add('Delete from Bsc');
    ExecSql;
    Sql.Clear;
    Sql.Add('insert into bsc (bsc_no) select bsc_no from bsc_file group by bsc_no');
    ExecSql;
  end;

  oleMapInfo.do('Register Table "' +
     gExePath + 'Bsc_file.dbf"  TYPE DBF Charset "WindowsSimpChinese" Into "' +
     gExePath + 'Bsc_file.TAB"');
  oleMapInfo.do('Open Table "' + gExePath + 'Bsc_file.TAB" Interactive');
  oleMapInfo.do('Open Table "' + gExePath + 'Cell.TAB" Interactive');
  oleMapInfo.do('Open Table "' + gExePath + 'Base.TAB" Interactive');
  oleMapInfo.do('Add Column "Cell" (Bsc_no )From Bsc_file Set To Bsc_no Where COL2 = COL1');
  oleMapInfo.do('Add Column "base" (Bsc_no )From Bsc_file Set To Bsc_no Where COL2 = COL1');
  oleMapInfo.do('Commit table cell');
  oleMapInfo.do('Commit table base');
  oleMapInfo.do('Register Table "' +
     gExePath + 'Bsc.dbf"  TYPE DBF Charset "WindowsSimpChinese" Into "' +
     gExePath + 'Bsc.TAB"');
  oleMapInfo.do('Open Table "' + gExePath + 'Bsc.TAB" Interactive');
  oleMapInfo.do('Create Table "MAP0097" (Bsc_name Char(10),Msc_no Char(8),' +
        'Lon Decimal(12,6),Lat Decimal(12,6),Bsc_no Char(6)) file "'
        + gExePath + 'MAP0097.TMP" TYPE DBF Version 300');
  oleMapInfo.do('Create Map For MAP0097 CoordSys Earth Projection 1, 0');
  oleMapInfo.do('Set Table MAP0097 FastEdit On Undo Off');
  oleMapInfo.do('Insert Into MAP0097 (Bsc_name, Msc_no, Lon, Lat, Bsc_no) Select Bsc_name, Msc_no, Lon, Lat, Bsc_no From bsc');
  oleMapInfo.do('Commit Table MAP0097');
  oleMapInfo.do('Close table bsc');
  oleMapInfo.do('Commit Table MAP0097 as "' + gExePath + 'bsc.tab"');
  oleMapInfo.do('close table MAP0097');

  //oleMapInfo.do('commit table ');

  UpdateCellObject;
  oleMapInfo.do('commit table cell');
  oleMapInfo.do('Export "Cell" Into "' + gExePath +
     'Cell.DBF" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  mmResetBaseClick(Self);
  oleMapInfo.RunMenuCommand(104);
  with dmBscData.tbBscControl do
  begin
    if not Active then
      Open;
    First;
    Edit;
    FieldByName('Base_update').AsString := 'Y' ;
    FieldByName('Cell_update').AsString := 'Y' ;
    FieldByName('bsc_update').AsString := 'Y' ;
    FieldByName('bsc_file_update').AsString := 'Y' ;
    FieldByName('Control_key').AsString := '1' ;
    Post;
    Close;
  end;
  mmDataConv.Enabled := True;
  mmMapConf.Enabled  := True;
  mmOpenMap.Enabled := True;
  ShowMessage('输入成功');
  sbBscMain.Panels[0].Text := '';
end;

procedure TfmBscMain.FormDblClick(Sender: TObject);
begin
  mmAllCdd.Enabled := True;
  mmSelCdd.Enabled := True;
  mmOpenMap.Enabled := True;
  mmDataConv.Enabled := True;
end;

procedure TfmBscMain.DFDF1Click(Sender: TObject);
begin
  Module := LoadLibrary('CellObj.dll');
  if Module > 32 then
  begin
    PFunc := GetProcAddress(Module, 'CreateObj');
    if  TCreateObj(PFunc)(oleMapInfo, 'NQI') then
      ShowMessage('建立成功！');

  end
  else
    ShowMessage('not find <CellObj.dll>');
  FreeLibrary(Module);
  oleMapInfo.do('Open table "' + gExePath + 'CellObj.tab" Interactive');
  oleMapInfo.do('Add Map auto layer CellObj');
  oleMapInfo.do('Add Column "CellOBj" (Bsc_no )From Cell Set To Bsc_no Where COL2 = COL2');
  oleMapInfo.do('Commit table CellObj');
  with dmBscData.tbBscControl do
  begin
    if not Active then
      Open;
    First;
    Edit;
    FieldByName('cell_Obj_Flag').AsString := 'Y';
    Post;
    Close;
  end;
end;

procedure TfmBscMain.N36Click(Sender: TObject);
begin
  Module := LoadLibrary('CellObj.dll');
  if Module > 32 then
  begin
    PFunc := GetProcAddress(Module, 'ColorDlg');
    if  TColorDlg(PFunc)(oleMapInfo, 'NQI') then
      ShowMessage('建立成功！');

  end
  else
    ShowMessage('not find <CellObj.dll>');
  FreeLibrary(Module);

end;

procedure TfmBscMain.N110Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'ERPAC' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (Erpac Decimal (8, 2))From Tch_file Set To Erpac Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with Erpac ranges apply all use color Brush (2,13697000,16777215)  0: 0.20 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,0.20: 0.40 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,0.40: 0.60 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,0.60: 0.80 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,0.80: 1.00 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "话务密度分析" Font ("宋体",0,10,255) subtitle "单位：ERP/每线" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');
  oleMapInfo.do('Create Cartographic Legend From Window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Behind Frame From Layer cellobj');
  oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N105Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'DR' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (DR Decimal (8, 2))From Tch_file Set To DR Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with DR ranges apply all use color Brush (2,13697000,16777215)  0: 0.1 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,0.1: 0.5 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,0.5: 1 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,1: 5 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,5: 100 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "TCH掉话率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');

  wHasAdd := False;
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'SDR' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (SDR Decimal (8, 2))From Cch_file Set To SDR Where COL2 = COL6  Dynamic');


  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' cellobj with Sdr graduated 0.0:0 15:36 Symbol (34,53456,36) vary size by ' +
  ' "SQRT"  # color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades off symbols on lines off count off title ' +
  ' "CCH掉话率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  ' ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  '  ,auto display off ,auto display off ,auto display off ,auto display on , ' +
  ' auto display on ,auto display on    ');
  oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N106Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'CG' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (CG Decimal (8, 2))From Tch_file Set To CG Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with CG ranges apply all use color Brush (2,13697000,16777215)  0: 0.1 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,0.1: 0.5 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,0.5: 1 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,1: 2 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,2: 100 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "TCH拥塞率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');

  wHasAdd := False;
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'SC' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (SC Decimal (8, 2))From Cch_file Set To SC Where COL2 = COL6  Dynamic');


  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' cellobj with SC graduated 0.0:0 50:36 Symbol (34,53456,36) vary size by ' +
  ' "SQRT"  # color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades off symbols on lines off count off title ' +
  ' "CCH拥塞率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  ' ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  '  ,auto display off ,auto display off ,auto display off ,auto display on , ' +
  ' auto display on ,auto display on    ');
  oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N104Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'U' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (U Decimal (8, 2))From Tch_file Set To U Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with U ranges apply all use color Brush (2,13697000,16777215) 98 : 100 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,95: 98 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,90: 95 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,85: 90 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,0: 80 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "TCH接通率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');

  wHasAdd := False;
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'SU' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (SU Decimal (8, 2))From Cch_file Set To SU Where COL2 = COL6  Dynamic');


  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' cellobj with SU graduated 0.0:0 100:36 Symbol (34,53456,36) vary size by ' +
  ' "SQRT"  # color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades off symbols on lines off count off title ' +
  ' "CCH接通率分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  ' ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  '  ,auto display off ,auto display off ,auto display off ,auto display on , ' +
  ' auto display on ,auto display on    ');
   oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N108Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'F' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (F Decimal (8, 2))From Tch_file Set To F Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with F ranges apply all use color Brush (2,13697000,16777215)  0: 0.05 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,0.05: 0.1 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,0.1: 0.5 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,0.5: 1 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,1: 100 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "TCH信道损坏率" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');

  wHasAdd := False;
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'SF' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (SF Decimal (8, 2))From Cch_file Set To SF Where COL2 = COL6  Dynamic');


  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' cellobj with SF graduated 0.0:0 1:36 Symbol (34,53456,36) vary size by ' +
  ' "SQRT"  # color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades off symbols on lines off count off title ' +
  ' "CCH信道损坏率" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  ' ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  '  ,auto display off ,auto display off ,auto display off ,auto display on , ' +
  ' auto display on ,auto display on    ');
  oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N109Click(Sender: TObject);
var
  i , wColNum : Integer;
  wHasAdd : Boolean;
begin
  wHasAdd := False;
  wColNum := oleMapInfo.eval('tableInfo(cellobj,4)');
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'RAF' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (RAF Decimal (8, 2))From Cch_file Set To RAF Where COL2 = COL6  Dynamic');

  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
      ' cellobj with RAF ranges apply all use color Brush (2,13697000,16777215)  0: 0.1 ' +
      ' Brush (2,13697000,16777215) Pen (1,2,0) ,0.1: 0.5 ' +
      ' Brush (2,14745504,16777215) Pen (1,2,0) ,0.5: 1 ' +
      ' Brush (2,15794000,16777215) Pen (1,2,0) ,1: 2 ' +
      ' Brush (2,16760896,16777215) Pen (1,2,0) ,2: 100 ' +
      ' Brush (2,16744576,16777215) Pen (1,2,0) ' +
      ' default Brush (2,16777215,16777215) Pen (1,2,0)  ' +
      ' # use 1 round 0.01 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades on symbols off lines off count on title ' +
  '  "随机接入失败率" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  '  ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  ' ,auto display on ,auto display on ,auto display on ,auto display on ,' +
  ' auto display on ');

 { wHasAdd := False;
  for i := 1 to wColNum do
  begin
    if UpperCase(oleMapInfo.eval('ColumnInfo(cellobj, col' + IntToStr(i) + ', 1 )')) = 'SC' then
      wHasAdd := True;
  end;
  if not wHasAdd then
    oleMapInfo.do('Add Column "CellObj" (SC Decimal (8, 2))From Cch_file Set To SC Where COL2 = COL6  Dynamic');


  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' cellobj with SC graduated 0.0:0 20:36 Symbol (34,16711680,36) vary size by ' +
  ' "SQRT"  # color 0 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  ' layer prev display on shades off symbols on lines off count off title ' +
  ' "CCH拥塞分析" Font ("宋体",0,10,255) subtitle "单位：%" Font ' +
  ' ("宋体",0,10,255) ascending on ranges Font ("Arial",0,8,0) auto display off ' +
  '  ,auto display off ,auto display off ,auto display off ,auto display on , ' +
  ' auto display on ,auto display on    ');   }
  oleMapInfo.RunMenuCommand(606)
end;

procedure TfmBscMain.N63Click(Sender: TObject);
var
  wStr : WideString;
begin
  wStr := 'NQI' ;
  wWirClass.DataThematic(oleMapInfo,wStr);
end;

procedure TfmBscMain.N64Click(Sender: TObject);
var
  wStr : WideString;
begin
  wStr := 'NQI' ;
  wWirClass.DataLabel(oleMapInfo,wStr);
end;

procedure TfmBscMain.N61Click(Sender: TObject);
var
  wStr : WideString;
begin
  wStr := 'NQI' ;
  wWirClass.AutoDiagnose(oleMapInfo, wStr);
end;

end.
