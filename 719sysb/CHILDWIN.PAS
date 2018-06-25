unit Childwin;

interface

uses Windows, Classes, Graphics, SysUtils, Forms, Controls, StdCtrls,
  ComCtrls, ExtCtrls, Dialogs,  Menus, db, ImgList,
  OleCtnrs;

type
  TfmMap = class(TForm)
    puCell: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    puBsc: TPopupMenu;
    mmDr: TMenuItem;
    mmTraffic: TMenuItem;
    mmChAc: TMenuItem;
    mmSu: TMenuItem;
    mmTraShade: TMenuItem;
    mmShowBsc: TMenuItem;
    mmCloseBsc: TMenuItem;
    N16: TMenuItem;
    PopupMenu1: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    N4: TMenuItem;
    ImageList1: TImageList;
    Panel1: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Splitter1: TSplitter;
    tvCell: TTreeView;
    GroupBox1: TGroupBox;
    lbCell: TListBox;
    TabSheet2: TTabSheet;
    TreeView1: TTreeView;
    Splitter2: TSplitter;
    paMap: TPanel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure tvCellChange(Sender: TObject; Node: TTreeNode);
    procedure tvCellClick(Sender: TObject);
    procedure tvCellDblClick(Sender: TObject);
    procedure lbCellDblClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure mmDrClick(Sender: TObject);
    procedure mmShowBscClick(Sender: TObject);
    procedure mmCloseBscClick(Sender: TObject);
    procedure mmTraShadeClick(Sender: TObject);
    procedure mmTrafficClick(Sender: TObject);
    procedure mmChAcClick(Sender: TObject);
    //procedure mmSaSsClick(Sender: TObject);
    procedure mmSuClick(Sender: TObject);
    //procedure mmDqaClick(Sender: TObject);
    //procedure mmRafClick(Sender: TObject);
    //procedure mmCgShadeClick(Sender: TObject);
    procedure tvCellEnter(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure lbCellClick(Sender: TObject);
  private
    { Private declarations }
    procedure BscCreateRegion_4(X, Y, L, Rate : Real; V : Variant; Flag : Integer; TB : String);
    procedure BscUpDateCircle(X, Y, L, Rate : Real; V : Variant; RowId, Flag : Integer; TableName : String);
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}
uses BscMain, BscData, BscBrow, Browse, legeng;

var
  gBscDr, gBscHover, gBscRaf, gBscChAc,
  gBscSaSs, gBscSu, gBscDqa, gBscTchCG, gBscTchTraffic,
  gBscTraffic: Boolean;

procedure TfmMap.BscCreateRegion_4(X, Y, L, Rate : Real; V : Variant; Flag : Integer; TB : String);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6 : Real;
begin
  X1 := X + L ;
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
  Y6 := Y + L * Sin(60 * DEG_2_Rad) * (103.1 / 111.2);


  //V.do('set style pen makepen(1,2, rgb(' +
  //               FloatToStr(255 * Rate) + ',' +  FloatToStr(255 * (1 - Rate))
  //               + ', 0))');
 
  if Flag = 1 then
  begin
    if TB = 'T' then
      V.do('create region  1 4 ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X6) + ',' + FloatToStr(Y6) + ') ('
                    + FloatToStr(X5) + ',' + FloatToStr(Y5) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')')
    else
      V.do('create region  1 4 ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')');
  end
  else
  begin
    if TB = 'T' then
      V.do('create region into variable TmpObject 1 4 ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X6) + ',' + FloatToStr(Y6) + ') ('
                    + FloatToStr(X5) + ',' + FloatToStr(Y5) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')')
    else
      V.do('create region into variable TmpObject 1 4 ('
                    + FloatToStr(X1) + ',' + FloatToStr(Y1) + ') ('
                    + FloatToStr(X2) + ',' + FloatToStr(Y2) + ') ('
                    + FloatToStr(X3) + ',' + FloatToStr(Y3) + ') ('
                    + FloatToStr(X4) + ',' + FloatToStr(Y4) + ')');
  end;
end;


procedure TfmMap.BscUpDateCircle(X, Y, L, Rate : Real; V : Variant; RowId, Flag : Integer; TableName : String);
const
  DEG_2_RAD = 0.01745329252 ;
var
  X1, Y1, X2, Y2 : real;
begin
  if Flag = 1 then
  begin
    X1 := X + L ;
    Y1 := Y ;

  end
  else
  begin
    X1 := X - L ;
    Y1 := Y ;
  end;

  V.do('update ' +  TableName +' set obj = createcircle (' + FloatToStr(X1) + ',' + FloatToStr(Y1) + ','
                  + FloatToStr(Rate) + ') where rowid = ' + IntToStr(rowid));
end;


procedure TfmMap.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if gMapNo = 1 then
  begin
    oleMapInfo.RunMenuCommand(104);
    fmBscMain.mmOpenMap.Enabled := True;
    fmBscMain.mmMapConf.Enabled := True;
    fmBscMain.mmDataConv.Enabled := True;
    fmBscMain.mmResetBase.Enabled := False;
    fmBscMain.mmResetCell.Enabled := False;
    fmBscMain.cbNext.Visible := False;
  end;
  fmBscMain.ResetMap;
  Action := caFree;
  gMapNo := gMapNo - 1;
end;

procedure TfmMap.FormCreate(Sender: TObject);
var
  sMsg : String;
  wNode, wBscNode, wCellNode : TTreeNode;
  wText, wBscNo : String;
begin
  tvCell.Items.Clear;
  wNode := tvCell.Items.Add(nil, '全部');
  wNode.ImageIndex := 1;
  wNode.SelectedIndex := 4;
  wNode.StateIndex := 4;
  with dmBscData.quCell do
  begin
    if not Active then
      Open;
    First;
    wBscNo := Trim(FieldByName('Bsc_no').AsString);
    wBscNode := tvCell.Items.AddChild(wNode, wBscNo);
    wBscNode.ImageIndex := 2;
    wBscNode.SelectedIndex := 4;
    while not eof do
    begin
      if wBscNo <> Trim(FieldByName('Bsc_no').AsString) then
      begin
        wBscNo := Trim(FieldByName('Bsc_no').AsString);

        wBscNode := tvCell.Items.AddChild(wNode, wBscNo);
        wBscNode.ImageIndex := 2;
        wBscNode.SelectedIndex := 4;
        wBscNode.StateIndex := 4;
        wCellNode := tvCell.Items.AddChild(wBScNode,
          Trim(FieldByName('CELL_ID').AsString) + ' ' +
          Trim(FieldByName('Cell_Name').AsString));
        //wCellNode.ImageIndex := 3;
        wCellNode.SelectedIndex := 4;
      end
      else
      begin
    //  while wBscNo = Trim(FieldByName('Bsc_no').AsString) do
     // begin
          wCellNode := tvCell.Items.AddChild(wBScNode,
          Trim(FieldByName('CELL_ID').AsString) + ' ' +
          Trim(FieldByName('Cell_Name').AsString));
          //wCellNode.ImageIndex := 3;
          wCellNode.SelectedIndex := 4;
      end;

      Next;
    end;
    Close;
  end;
  //tvCell.Selected.Expand(True);
  Str(paMap.handle, sWinHand);
  //Str(handle, sWinHand);
  sMsg := 'Set Next Document Parent ' + sWinhand + 'Style 1';



  oleMapInfo.Do(sMsg);
  if gMapNo > 0 then
  begin
    Caption := '副图';
    oleMapInfo.do('Run Command WindowInfo(FrontWindow(),15)');
    gMapNo := gMapNo + 1;
    //WindowState := wsNormal;
  end;
  //oleMapInfo.RunMenuCommand(102);
  gMultiCell := lbCell.Items;
  fmBscMain.mmOpenMap.Enabled := False;
  fmBscMain.mmMapConf.Enabled := False;
  fmBscMain.mmDataConv.Enabled := False;
  fmBscMain.mmResetBase.Enabled := True;
  fmBscMain.mmResetCell.Enabled := True;
  Application.CreateForm(TfmLegeng, fmLegeng);
end;

procedure TfmMap.FormResize(Sender: TObject);
var
  wHandle : THandle;
  mapwinid : Integer;
  sMapwinid : String;
begin
  if oleMapInfo.eval('NumTables()') > 0 then
  begin
    mapWinID := oleMapInfo.Eval('FrontWindow()');
    sMapwinId := IntToStr(mapWinID);
    wHandle := oleMapInfo.eval('WindowInfo(' + sMapWinID + ', 12)');
    MoveWindow(wHandle,0,0,Width,Height,true);
  end;
end;

procedure TfmMap.tvCellChange(Sender: TObject; Node: TTreeNode);
begin
  if tvCell.Selected.Parent = nil then
  begin
    //ShowMessage('All');
    gSelName := tvCell.Selected.Text;
    gSelFlag := 'ALL';
  end
  else
  begin
    if (tvCell.Selected.HasChildren)  then
    begin
      gSelName := tvCell.Selected.Text;
      gSelFlag := 'BSC';
    end
    else
    begin
      gSelName := Copy(tvCell.Selected.Text, 1, Pos(' ',tvCell.Selected.Text) - 1);
      gSelFlag := 'CELL';
    end;
  end;
end;

procedure TfmMap.tvCellClick(Sender: TObject);
begin
  fmBscMain.cbNext.Visible := False;
  if gSelFlag = 'All' then
  begin
    oleMapInfo.do('Set Map  Center (113.325662, 22.279674)');
    oleMapInfo.do('Set Map  Scale 1 Units "cm" For 3 Units "km"');
  end;
  if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('select Bsc_no, avg(Lon), avg(Lat)  from cell where bsc_no = "' + gSelName + '" group by Bsc_no into tmp');
    oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.5 Units "km"');
    oleMapInfo.do('Set Map  Center (tmp.col2, tmp.col3)');
    oleMapInfo.do('Close table tmp');
  end;
  if gSelFlag = 'CELL' then
  begin
    oleMapInfo.do('select * from cell where bs_no = "' + gSelName + '" into tmp');
    oleMapInfo.do('Set Map  Scale 1 Units "cm" For 0.2 Units "km"');
    oleMapInfo.do('Set Map  Center (tmp.Lon, tmp.Lat)');
    oleMapInfo.do('Close table tmp');
  end;
  if gCqt then
  begin
    oleMapInfo.do('Close table Cqt_shade Interactive');
    gCqt := False;
  end;
  if gMh then
  begin
    oleMapInfo.do('Close table mh_Shade  Interactive');
    //oleMapInfo.do('Close table TCh_dr  Interactive');
    gMh := False;
    fmBscMain.mmMh.Checked := False;
  end;
  if fmBscMain.mmDensityShade.Checked then
  begin
    oleMapInfo.do('Close table Density_Shade  Interactive');
    fmBscMain.mmDensityShade.Checked := False;
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
 // fmBscMain.mmCgShade.Checked := False;
  fmBscMain.mmdr.Checked := False;
  fmBscMain.mmHover.Checked := False;
  fmBscMain.mmChAc.Checked := False;
  fmBscMain.mmDqa.Checked := False;
  fmBscMain.mmSaSs.Checked := False;
  fmBscMain.mmRaf.Checked := False;
  fmBscMain.mmSu.Checked := False;
  fmBscMain.mmTraffic.Checked := False;
 // wTableNum := oleMapInfo.eval('NumTables()');
 { for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_SDR' then
      oleMapInfo.do('Close table CCh_Sdr  Interactive');
  end;
  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_DR' then
      oleMapInfo.do('Close table TCh_dr  Interactive');
  end;
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_TRAFFIC' then
      oleMapInfo.do('Close table CCh_TRAFFIC  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_TRAFFIC' then
      oleMapInfo.do('Close table TCh_TRAFFIC  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_CH' then
      oleMapInfo.do('Close table CCh_CH  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_CH' then
      oleMapInfo.do('Close table TCh_CH  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_SA' then
      oleMapInfo.do('Close table CCh_SA  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_CA' then
      oleMapInfo.do('Close table TCh_CA  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_SU' then
      oleMapInfo.do('Close table CCh_SU  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_U' then
      oleMapInfo.do('Close table TCh_U  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'CCH_DQA' then
      oleMapInfo.do('Close table CCh_DQA  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'TCH_TQA' then
      oleMapInfo.do('Close table TCh_TQA  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'HDOV_I_BSC' then
      oleMapInfo.do('Close table HDOV_I_BSC  Interactive');
  for i := 1 to wTableNum do
    if UpperCase(Trim(oleMapInfo.eval('tableInfo('+ IntToStr(i)+', 1)'))) = 'HDOV_E_BSC' then
      oleMapInfo.do('Close table HDOV_E_BSC  Interactive');
    }
end;

procedure TfmMap.tvCellDblClick(Sender: TObject);
begin
  if gSelFlag = 'CELL' then
  begin
    if lbCell.Items.IndexOf(tvCell.Selected.Text) < 0 then
    begin
      lbCell.Items.Add(tvCell.Selected.Text);
      gMultiCell := lbCell.Items;
    end
  end;
end;

procedure TfmMap.lbCellDblClick(Sender: TObject);
begin
  if lbCell.Items.Count > 0 then
  begin
    if lbCell.ItemIndex >= 0 then
      lbCell.Items.Delete(lbCell.ItemIndex)
    else
      lbCell.Items.Delete(0);
    gMultiCell := lbCell.Items;
  end;
end;

procedure TfmMap.N1Click(Sender: TObject);
begin
  lbCell.Items.Clear;
  gMultiCell := lbCell.Items;
end;

procedure TfmMap.N3Click(Sender: TObject);
var
  wRow, i : Integer;
begin
  lbCell.Items.Clear;
  wRow := oleMapInfo.eval('tableinfo(selection,8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from selection');
    lbCell.Items.Add(oleMapInfo.eval('selection.bs_no') + ' ' +
                     oleMapInfo.eval('selection.cell_name'));
  end;
end;

procedure TfmMap.mmDrClick(Sender: TObject);
var
  i, wRow, wNcellRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty : real;
begin
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  if gBscDr then
  begin
    if mmDr.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_sdr Display Off');
      oleMapInfo.do('Set Map Layer bsc_tch_dr Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_sdr Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_Tch_dr Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmDr.Checked := not mmDr.Checked;
    exit;
  end;
  mmDr.Checked := not mmDr.Checked;
  oleMapInfo.do('set map redraw off');
  //sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.Sdr, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no and cell.Bsc_no = "'
                 + gSelName + '" into SdrTmp');
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '" into SdrTmp')
      else
        oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into SdrTmp');
    end
    else
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Sdr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into SdrTmp');
  end; }

  oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.Sdr, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into SdrTmp');

  oleMapInfo.do('commit table SdrTmp as "' + gExePath + 'bsc_cch_sdr.tab"');
  oleMapInfo.do('close table SdrTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_sdr.tab" Interactive');

  oleMapInfo.do('select Max(sdr) from bsc_cch_sdr into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_Sdr, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_sdr');
    wLon := oleMapInfo.eval('bsc_cch_sdr.lon');
    wLat := oleMapInfo.eval('bsc_cch_sdr.lat');
    //wBearing := oleMapInfo.eval('bsc_cch_sdr.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_sdr.sdr') / wMaxQty;
    //oleMapInfo.do('set style brush makebrush(82,rgb(0,' + FloatToStr(255 * (1 - wRate)) + ','
    //            + FloatToStr(255 * (1 - wRate)) +'), rgb(255,255,255))');
    //BscCreateRegion_4(wLon, wLat, 0.02, wRate, oleMapInfo, 2, 'T');
    //oleMapInfo.do('Update bsc_cch_sdr set obj = TmpObject where rowid = ' + IntToStr(i));
    if wRate > 0 then
      wRate := 0.05 + 0.5 * wRate;

    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'Bsc_cch_sdr');
  end;
  oleMapInfo.do('commit table bsc_cch_sdr');
  oleMapInfo.do('add map auto layer bsc_cch_sdr');
  oleMapInfo.do('Set Map Layer bsc_cch_sdr Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With sdr+"%" Auto On Visibility Zoom (0, 1000) Units "km"');


  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into drTmp');
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
    else
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.dr, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into drTmp');
  end;}
  oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.dr, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into drTmp');
  oleMapInfo.do('commit table drTmp as "' + gExePath + 'bsc_tch_dr.tab"');
  oleMapInfo.do('close table drTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_dr.tab" Interactive');

  oleMapInfo.do('select Max(dr) from bsc_tch_dr into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_dr, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_dr');
    wLon := oleMapInfo.eval('bsc_tch_dr.lon');
    wLat := oleMapInfo.eval('bsc_tch_dr.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_dr.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_dr.dr') / wMaxQty;

    {oleMapInfo.do('set style brush makebrush(82,rgb(' +
                 FloatToStr(255 * (1- wRate)) +
                 ',' + FloatToStr(255 * (1- wRate)) +
                 ',0), rgb(255,255,255))');
    BscCreateRegion_4(wLon, wLat, 0.02, wRate, oleMapInfo, 2, 'B');
    oleMapInfo.do('Update bsc_Tch_dr set obj = TmpObject where rowid = ' + IntToStr(i));
    }
    
    if wRate > 0 then
        wRate := 0.05 + 0.5 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 2, 'Bsc_tch_dr');
  end;
  oleMapInfo.do('commit table bsc_tch_dr');
  oleMapInfo.do('add map auto layer bsc_tch_dr');
  oleMapInfo.do('Set Map Layer bsc_tch_dr Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With dr+"%" Auto On Visibility Zoom (0, 1000) Units "km"');

  oleMapInfo.do('set map redraw on');
  gBscDr := True;
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC -- ' + TMenuItem(Sender).Caption;
end;

procedure TfmMap.mmShowBscClick(Sender: TObject);
begin
  fmBscMain.ShowBsc;
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  oleMapInfo.do('Set Map Layer area Label Auto Off');
  mmCloseBsc.Enabled := True;
  mmDr.Enabled := True;
  //mmHover.Enabled := True;
  //mmRaf.Enabled := True;
  mmTraffic.Enabled := True;
  mmChAc.Enabled := True;
  //mmSaSs.Enabled := True;
  mmSu.Enabled := True;
  //mmDqa.Enabled := True;
  mmTraShade.Enabled := True;
  //mmCgShade.Enabled := True;
  mmShowBsc.Enabled := False;
end;

procedure TfmMap.mmCloseBscClick(Sender: TObject);
begin
  if gBscShowing then
  begin
    oleMapInfo.do('Set Map Layer area Label Auto On');
    oleMapInfo.do('close table bsc');
    oleMapInfo.do('close table bsc_tch_file');
    oleMapInfo.do('close table bsc_cch_file');
    oleMapInfo.do('Set Map Layer cell Display Graphic');
    oleMapInfo.do('Set Map Layer base Display Graphic');
    oleMapInfo.do('Set Map Layer street Display Graphic');
    mmShowBsc.Enabled := True;
    mmCloseBsc.Enabled := False;
    mmDr.Enabled := False;
    //mmHover.Enabled := False;
    //mmRaf.Enabled := False;
    mmTraffic.Enabled := False;
    mmChAc.Enabled := False;
    //mmSaSs.Enabled := False;
    mmSu.Enabled := False;
    //mmDqa.Enabled := False;
    //mmCgShade.Enabled := False;
    mmTraShade.Enabled := False;
  end;

  if gBscDr then
  begin
    oleMapInfo.do('Close table bsc_CCh_Sdr  Interactive');
    oleMapInfo.do('Close table Bsc_TCh_dr  Interactive');
  end;
  {if gHover then
  begin
    oleMapInfo.do('Close table HDOV_I_BSC  Interactive');
    oleMapInfo.do('Close table HDOV_E_BSC  Interactive');
    oleMapInfo.do('Close table HDOV_I_61BSC  Interactive');
    oleMapInfo.do('Close table HDOV_E_61BSC  Interactive');
  end; }
  if gBscTraffic then
  begin
    oleMapInfo.do('Close table bsc_CCh_Traffic  Interactive');
    oleMapInfo.do('Close table bsc_TCh_Traffic  Interactive');
  end;
  if gBScChAc then
  begin
    oleMapInfo.do('Close table bsc_CCh_ch  Interactive');
    oleMapInfo.do('Close table bsc_TCh_ch  Interactive');
  end;
  if gBscSaSs then
  begin
    oleMapInfo.do('Close table bsc_CCh_Sa  Interactive');
    oleMapInfo.do('Close table bsc_TCh_ca  Interactive');
  end;
  if gBScRaf then
  begin
    oleMapInfo.do('Close table bsc_raf_shade  Interactive');
  end;
  if gBscDqa then
  begin
    oleMapInfo.do('Close table bsc_CCh_dqa  Interactive');
    oleMapInfo.do('Close table bsc_TCh_tqa  Interactive');
    oleMapInfo.do('Close table bsc_CCh_dss4  Interactive');
    oleMapInfo.do('Close table bsc_TCh_tss4  Interactive');
  end;
  if gBScRaf then
  begin
  end;
  if gBscSu then
  begin
    oleMapInfo.do('Close table bsc_CCh_Su  Interactive');
    oleMapInfo.do('Close table bsc_TCh_u  Interactive');
  end;
  if gbscTchCG then
  begin
     oleMapInfo.do('Close table bsc_cg_shade  Interactive');
  end;
  if gbscTchTraffic then
  begin
     oleMapInfo.do('Close table bsc_erpac_shade  Interactive');
  end;
  gBscTchCG := False;
  gBscTchTraffic := False;
  gBscDr := False;
  gBscHover := False;
  gBscChAc := False;
  gBscDqa := False;
  gBscSaSs := False;
  gBscRaf := False;
  gBscSu := False;
  gBscTraffic := False;

  mmTraShade.Checked := False;
  //mmCgShade.Checked := False;
  mmdr.Checked := False;
  //mmHover.Checked := False;
  mmChAc.Checked := False;
  //mmDqa.Checked := False;
  //mmSaSs.Checked := False;
  //mmRaf.Checked := False;
  mmSu.Checked := False;
  mmTraffic.Checked := False;
  gBscShowing := False;
  
end;

procedure TfmMap.mmTraShadeClick(Sender: TObject);
procedure BscTchTrafficShade;
begin
  oleMapInfo.do('commit table bsc as "' + gExePath + 'bsc_erpac_shade.tab"');
  oleMapInfo.do('open table "' + gExePath + 'bsc_erpac_shade.tab"');
  oleMapInfo.do('add map auto layer bsc_erpac_shade');
  oleMapInfo.do('Add Column "bsc_erpac_shade" (Erpac Decimal (8, 2))From bsc_Tch_file Set To Erpac Where COL5 = COL4  Dynamic');
  //oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' bsc_erpac_shade with Erpac ignore 0 ranges apply all use color ' +
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
    'display on');

  oleMapInfo.do('Set Map Layer bsc_erpac_shade Label Position Above Font ("Arial",256,10,65535,0) With erpac Overlap On Visibility Zoom (0, 100) Units "km"');
  oleMapInfo.do('Set Map Layer bsc_erpac_shade Label Auto On');
  gBscTchTraffic := True;
end;

begin
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  if gBscTchTraffic then
  begin
    if mmTraShade.Checked then
    begin
      oleMapInfo.do('close table Bsc_erpac_shade');
    end
    else
    begin
      BscTchTrafficShade;
    end;
  end
  else
    BscTchTrafficShade;
  mmTraShade.Checked := not mmTraShade.Checked;
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC -- ' + TMenuItem(Sender).Caption; 
end;

procedure TfmMap.mmTrafficClick(Sender: TObject);
var
  i, wRow, wNcellRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wPenWidth : real;
begin
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  if gBscTraffic then
  begin
    if mmTraffic.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Bsc_cch_traffic Display Off');
      oleMapInfo.do('Set Map Layer Bsc_Tch_traffic Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Bsc_cch_traffic Display Graphic');
      oleMapInfo.do('Set Map Layer Bsc_tch_traffic Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmTraffic.Checked := not mmTraffic.Checked;
    exit;
  end;
  mmTraffic.Checked := not mmTraffic.Checked;
  oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into TrafficTmp');
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
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Traffic,cch_file.sc, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into TrafficTmp');
  end;}
  oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.Traffic,bsc_cch_file.sc, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into BscTrafficTmp');
  oleMapInfo.do('commit table BscTrafficTmp as "' + gExePath + 'bsc_cch_Traffic.tab"');
  oleMapInfo.do('close table bscTrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_Traffic.tab" Interactive');

  oleMapInfo.do('select Max(Traffic) from bsc_cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(Traffic) from bsc_Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_Traffic, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_Traffic');
    wLon := oleMapInfo.eval('bsc_cch_Traffic.lon');
    wLat := oleMapInfo.eval('bsc_cch_Traffic.lat');
    //wBearing := oleMapInfo.eval('bsc_cch_Traffic.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_Traffic.Traffic') / wMaxQty;
    wPenWidth := 1 + oleMapInfo.eval('bsc_cch_Traffic.Sc');
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    if wRate > 0 then
      wRate := 0.02 + 0.5 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'bsc_cch_Traffic');

  end;
  oleMapInfo.do('commit table bsc_cch_Traffic');
  oleMapInfo.do('add map auto layer bsc_cch_Traffic');
  oleMapInfo.do('Set Map Layer bsc_cch_Traffic Label Position Above Font ("Arial",0,12,0) With Traffic+","+sc+"%" Auto On Visibility Zoom (0, 1000) Units "km"');

  ////tch
 { if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '" into TrafficTmp');
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
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Traffic,Tch_file.CG, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into TrafficTmp');
  end;}
  oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.Traffic,bsc_Tch_file.CG, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into bscTrafficTmp');
  oleMapInfo.do('commit table bscTrafficTmp as "' + gExePath + 'bsc_tch_Traffic.tab"');
  oleMapInfo.do('close table bscTrafficTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_Traffic.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  //oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_Traffic, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_Traffic');
    wLon := oleMapInfo.eval('bsc_tch_Traffic.lon');
    wLat := oleMapInfo.eval('bsc_tch_Traffic.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_Traffic.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_Traffic.Traffic') / wMaxQty;
    wPenWidth := 1 + oleMapInfo.eval('bsc_tch_Traffic.cg');
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    if wRate > 0 then
      wRate := 0.02 + 0.5 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 4, 'bsc_tch_Traffic');

  end;
  oleMapInfo.do('commit table bsc_tch_Traffic');
  oleMapInfo.do('add map auto layer bsc_tch_Traffic');
  oleMapInfo.do('Set Map Layer bsc_tch_Traffic Label Position Above Font ("Arial",0,12,0) With Traffic+","+cg+"%" Auto On Visibility Zoom (0, 1000) Units "km"');

  oleMapInfo.do('set map redraw on');
  gBscTraffic := True;
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC -- ' + TMenuItem(Sender).Caption;

end;

procedure TfmMap.mmChAcClick(Sender: TObject);
var
  i, wRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty , wPenWidth: real;
begin
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  if gBscChAc then
  begin
    if mmChAc.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_ch Display Off');
      oleMapInfo.do('Set Map Layer bsc_tch_ch Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_ch Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_tch_ch Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmChAc.Checked := not mmChAc.Checked;
    exit;
  end;
  mmChAc.Checked := not mmChAc.Checked;
  oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into chTmp');
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
    else
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.ch, cch_file.ac, cch_file.sf, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into chTmp');
  end;}

  oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.ch, bsc_cch_file.ac, bsc_cch_file.sf, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into bscchTmp');
  oleMapInfo.do('commit table bscchTmp as "' + gExePath + 'bsc_cch_ch.tab"');
  oleMapInfo.do('close table bscchTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_ch.tab" Interactive');

  oleMapInfo.do('select Max(ch) from bsc_cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(ch) from bsc_Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_ch, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_ch');
    wLon := oleMapInfo.eval('bsc_cch_ch.lon');
    wLat := oleMapInfo.eval('bsc_cch_ch.lat');
   // wBearing := oleMapInfo.eval('cch_ch.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_ch.ch')/ wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.6 * wRate;
    wPenWidth := 1 + 6 * oleMapInfo.eval('bsc_cch_ch.sf') / 100;
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
   // if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'bsc_cch_ch');

    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 2, 'cch_ch');
  end;
  oleMapInfo.do('commit table bsc_cch_ch');
  oleMapInfo.do('add map auto layer bsc_cch_ch');
  oleMapInfo.do('Set Map Layer bsc_cch_ch Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+sf+"%)" Auto On Visibility Zoom (0, 1000) Units "km"');
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
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 2, 'cch_ac');
  end;
  oleMapInfo.do('commit table cch_ac');
  oleMapInfo.do('Set Map Layer cch_ac Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+sf+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
   }
  //////////////////////////////////////////////Tch
 { if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.ch, bsc_tch_file.ac, bsc_tch_file.f, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no and cell.bsc_no = "'
                 + gSelName + '"into chTmp');
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
    else
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.ch, tch_file.ac, tch_file.f, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into chTmp');
  end; }

  oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.ch, bsc_tch_file.ac, bsc_tch_file.f, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into chTmp');
  oleMapInfo.do('commit table chTmp as "' + gExePath + 'bsc_tch_ch.tab"');
  oleMapInfo.do('close table chTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_ch.tab" Interactive');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_ch, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_ch');
    wLon := oleMapInfo.eval('bsc_tch_ch.lon');
    wLat := oleMapInfo.eval('bsc_tch_ch.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_ch.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_ch.ch')/ wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.6 * wRate;
     wPenWidth := 1 + 6 * oleMapInfo.eval('bsc_tch_ch.f');
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
   // if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 4, 'bsc_Tch_ch');
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 1, 'tch_ch');
  end;
  oleMapInfo.do('commit table bsc_Tch_ch');
  oleMapInfo.do('add map auto layer bsc_tch_ch');

  oleMapInfo.do('Set Map Layer bsc_Tch_ch Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+f+"%)" Auto On Visibility Zoom (0, 1000) Units "km"');
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
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 1, 'Tch_ac');
  end;
  oleMapInfo.do('commit table tch_ac');
  oleMapInfo.do('Set Map Layer tch_ac Label Position Above Font ("Arial",0,10,0) With ch+"("+ac+","+f+"%)" Auto On Visibility Zoom (0, 6) Units "km"');
    }
  oleMapInfo.do('set map redraw on');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC -- ' + TMenuItem(Sender).Caption;
  gBscChAc := True;
end;

{procedure TfmMap.mmSaSsClick(Sender: TObject);
var
  i, wRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wPenWidth : real;
begin
  if gBScSaSs then
  begin
    if mmSaSs.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_SA Display Off');
      oleMapInfo.do('Set Map Layer bsc_tch_CA Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_SA Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_tch_Ca Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmSaSs.Checked := not mmSaSs.Checked;
    exit;
  end;
  mmSaSs.Checked := not mmSaSs.Checked;
  oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.SA, bsc_cch_file.Ss, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no and cell.bsc_no = "'
                 + gSelName + '"into SATmp');
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
    else
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.SA, cch_file.Ss, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into SATmp');
  end;}

  {oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.SA, bsc_cch_file.Ss, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into SATmp');

  oleMapInfo.do('commit table SATmp as "' + gExePath + 'bsc_cch_SA.tab"');
  oleMapInfo.do('close table SATmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_SA.tab" Interactive');

  oleMapInfo.do('select Max(sa) from bsc_cch_file into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  oleMapInfo.do('select Max(ca) from bsc_Tch_file into tmp');
  if wMaxQty < oleMapInfo.eval('tmp.col1') then
    wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_SA, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_SA');
    wLon := oleMapInfo.eval('bsc_cch_sA.lon');
    wLat := oleMapInfo.eval('bsc_cch_sA.lat');
    //wBearing := oleMapInfo.eval('bsc_cch_SA.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_SA.SA')/ wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.6 * wRate;
    if oleMapInfo.eval('bsc_cch_sa.Sa') <> 0 then
      wPenWidth := 1 +  6 *(oleMapInfo.eval('bsc_cch_sa.Sa') - oleMapInfo.eval('bsc_cch_sa.Ss')) /
                     oleMapInfo.eval('bsc_cch_sa.Sa')
    else
      wPenWidth := 1;
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    //if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'bsc_cch_sa');
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 2, 'cch_sa');
  end;
  oleMapInfo.do('commit table bsc_cch_SA');
  oleMapInfo.do('add map auto layer bsc_cch_SA');
  oleMapInfo.do('Set Map Layer bsc_cch_SA Label Position Above Font ("Arial",0,10,0) With sa+"("+Ss+")" Auto On Visibility Zoom (0, 1000) Units "km"');
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
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into cATmp');
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into cATmp')
      else
        oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + ' into cATmp')
    end
    else
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.cA, tch_file.cs, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into cATmp');
  end;}

  {oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.cA, bsc_tch_file.cs, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into cATmp');
  oleMapInfo.do('commit table cATmp as "' + gExePath + 'bsc_tch_cA.tab"');
  oleMapInfo.do('close table cATmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_cA.tab" Interactive');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_cA, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_cA');
    wLon := oleMapInfo.eval('bsc_tch_ca.lon');
    wLat := oleMapInfo.eval('bsc_tch_ca.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_ca.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_cA.ca')/ wMaxQty;
    if wRate > 0 then
      wRate := 0.01 + 0.6 * wRate;
    if oleMapInfo.eval('bsc_tch_ca.ca') <> 0 then
      wPenWidth := 1 +  6 *(oleMapInfo.eval('bsc_tch_ca.ca') - oleMapInfo.eval('bsc_tch_ca.cs')) /
                     oleMapInfo.eval('bsc_tch_ca.ca')
    else
      wPenWidth := 1;
    if wPenWidth > 7 then
      wPenWidth := 7;
    oleMapInfo.do('set style pen makepen(' + FloatToStr(wPenWidth) + ',2, rgb(255,0,0))');
    //if wRate > 0 then
    //  wRate := 0.02 + 0.05 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 4, 'bsc_tch_ca');
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    //UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 1, 'tch_ca');
  end;
  oleMapInfo.do('commit table bsc_Tch_cA');
  oleMapInfo.do('add map auto layer bsc_tch_cA');
  oleMapInfo.do('Set Map Layer bsc_Tch_cA Label Position Above Font ("Arial",0,10,0) With ca+"("+cs+")" Auto On Visibility Zoom (0, 1000) Units "km"');
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
    //UpdateASObject(wLon, wLat, 0.0025, 60, wBearing, wR1, Wr2, oleMapInfo, i, 2, 'cch_SA_SS');
    UpDateCircle(wLon, wLat, 0.0025, 60, wBearing, wRate, oleMapInfo, i, 1, 'Tch_cs');
  end;
  oleMapInfo.do('commit table tch_cs');
  oleMapInfo.do('Set Map Layer tch_cs Label Position Above Font ("Arial",0,10,0) With ca+"("+cs+")" Auto On Visibility Zoom (0, 6) Units "km"');
  }

  {oleMapInfo.do('set map redraw on');
  gBscSaSs := True;
  fmBscMain.sbBscMain.Panels[0].Text := '';
end;   }

procedure TfmMap.mmSuClick(Sender: TObject);
var
  i, wRow, wNcellRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty, wMinQty : real;
begin
  oleMapInfo.do('Set Map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' Zoom Entire Layer bsc');
  if gBscSu then
  begin
    if mmSu.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Bsc_cch_Su Display Off');
      oleMapInfo.do('Set Map Layer Bsc_tch_U Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Bsc_cch_su Display Graphic');
      oleMapInfo.do('Set Map Layer Bsc_tch_u Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmSu.Checked := not mmSu.Checked;
    exit;
  end;
  mmSu.Checked := not mmSu.Checked;
  oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select Bsc_CCh_file.bsc_no,Bsc_cch_file.su, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from Bsc_cch_file, bsc '
                 + 'where Bsc_cch_file.bsc_no = bsc.bsc_no and cell.bsc_no = "'
                 + gSelName + '"into suTmp');
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
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.su, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into suTmp');
  end; }

  oleMapInfo.do('Select Bsc_CCh_file.bsc_no,Bsc_cch_file.su, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from Bsc_cch_file, bsc '
                 + 'where Bsc_cch_file.bsc_no = bsc.bsc_no and bsc.bsc_no into suTmp');

  oleMapInfo.do('commit table SuTmp as "' + gExePath + 'bsc_cch_su.tab"');
  oleMapInfo.do('close table SuTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_su.tab" Interactive');

  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(Bsc_cch_Su, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Bsc_cch_su');
    wLon := oleMapInfo.eval('Bsc_cch_su.lon');
    wLat := oleMapInfo.eval('Bsc_cch_su.lat');
   // wBearing := oleMapInfo.eval('Bsc_cch_su.Bearing');
    wRate := oleMapInfo.eval('Bsc_cch_su.su') / 100;
    if wRate > 0 then
      wRate := 0.6 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'Bsc_cch_su');
  end;
  oleMapInfo.do('commit table Bsc_cch_su');
  oleMapInfo.do('add map auto layer Bsc_cch_su');
  oleMapInfo.do('Set Map Layer Bsc_cch_su Label Position Above Font ("Arial",0,10,0) With Su+"%" Auto On Visibility Zoom (0, 1000) Units "km"');

  ////////////////////////////////////////////tch
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into uTmp');
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
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.u, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into uTmp');
  end;}
  oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.u, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into uTmp');
  oleMapInfo.do('commit table uTmp as "' + gExePath + 'bsc_tch_u.tab"');
  oleMapInfo.do('close table uTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_u.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_u, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_u');
    wLon := oleMapInfo.eval('bsc_tch_u.lon');
    wLat := oleMapInfo.eval('bsc_tch_u.lat');
   // wBearing := oleMapInfo.eval('bsc_tch_u.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_u.u') / 100;
    if wRate > 0 then
      wRate :=  0.6 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 4, 'bsc_tch_u');
  end;
  oleMapInfo.do('commit table bsc_tch_u');
   oleMapInfo.do('add map auto layer bsc_tch_u');
  oleMapInfo.do('Set Map Layer bsc_tch_u Label Position Above Font ("Arial",0,10,0) With u+"%" Auto On Visibility Zoom (0, 1000) Units "km"');

  oleMapInfo.do('set map redraw on');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC -- ' + TMenuItem(Sender).Caption;
  gBScSu := True;
end;

{procedure TfmMap.mmDqaClick(Sender: TObject);
var
  i, wRow, wNcellRow : Integer;
  wLon, wLat, wBearing,  wRate, wMaxQty : real;
begin
  if gBscDqa then
  begin
    if mmDqa.Checked then
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_Dqa Display Off');
      oleMapInfo.do('Set Map Layer bsc_Tch_Tqa Display off');
      oleMapInfo.do('Set Map Layer bsc_cch_Dss4 Display Off');
      oleMapInfo.do('Set Map Layer bsc_Tch_Tss4 Display off');
      oleMapInfo.do('set map redraw on')
    end
    else
    begin
      oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer bsc_cch_Dqa Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_Tch_Tqa Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_cch_Dss4 Display Graphic');
      oleMapInfo.do('Set Map Layer bsc_Tch_Dss4 Display Graphic');
      oleMapInfo.do('set map redraw on')
    end;
    mmDqa.Checked := not mmDqa.Checked;
    exit;
  end;
  mmDqa.Checked := not mmDqa.Checked;
  oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into DqaTmp');
  end
  else
  begin
    if gSelFlag = 'CELL' then
    begin
      if gMultiCell.Count = 0 then
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bs_no = "'
                 + gSelName + '"into DqaTmp')
      else
         oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and '
                 + GetMultiCell(gMultiCell) + 'into DqaTmp')


    end
    else
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.dqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into DqaTmp');
  end; }

  {oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.dqa, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into DqaTmp');
  oleMapInfo.do('commit table DqaTmp as "' + gExePath + 'bsc_cch_Dqa.tab"');
  oleMapInfo.do('close table DqaTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_Dqa.tab" Interactive');

  oleMapInfo.do('select Max(Dqa) from bsc_cch_Dqa into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_Dqa, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_Dqa');
    wLon := oleMapInfo.eval('bsc_cch_Dqa.lon');
    wLat := oleMapInfo.eval('bsc_cch_Dqa.lat');
   // wBearing := oleMapInfo.eval('bsc_cch_Dqa.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_Dqa.Dqa') / wMaxQty;

    oleMapInfo.do('set style brush makebrush(82,rgb(0,' + FloatToStr(255 * (1 - wRate)) + ','
                + FloatToStr(255 * (1 - wRate)) +'), rgb(255,255,255))');
    BscCreateRegion_4(wLon, wLat, 0.02,  wRate, oleMapInfo, 2, 'T');
    oleMapInfo.do('Update bsc_cch_dqa set obj = TmpObject where rowid = ' + IntToStr(i));
  end;
  oleMapInfo.do('commit table bsc_cch_Dqa');
  oleMapInfo.do('add map auto layer bsc_cch_Dqa');
  oleMapInfo.do('Set Map Layer bsc_cch_Dqa Label Position Above Font ("Arial",256,8,16777215,0) ' +
               ' With Dqa+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  //Tch
  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select bsc_TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into TqaTmp');
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
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.Tqa, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into TqaTmp');
  end;}

  {oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.Tqa, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into TqaTmp');
  oleMapInfo.do('commit table TqaTmp as "' + gExePath + 'bsc_tch_Tqa.tab"');
  oleMapInfo.do('close table TqaTmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_Tqa.tab" Interactive');

  oleMapInfo.do('select Max(Tqa) from bsc_tch_Tqa into tmp');
  wMaxQty := oleMapInfo.eval('tmp.col1');
  oleMapInfo.do('Close table tmp');
  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  //oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_Tqa, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_Tqa');
    wLon := oleMapInfo.eval('bsc_tch_Tqa.lon');
    wLat := oleMapInfo.eval('bsc_tch_Tqa.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_Tqa.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_Tqa.Tqa') / wMaxQty;

    oleMapInfo.do('set style brush makebrush(82,rgb(' +

                  FloatToStr(255 * (1 - wRate)) + ',' +
                  FloatToStr(255 * (1 - wRate)) + ',0), rgb(255,255,255))');

    BscCreateRegion_4(wLon, wLat, 0.02, wRate, oleMapInfo, 2, 'B');
    oleMapInfo.do('Update bsc_Tch_Tqa set obj = TmpObject where rowid = ' + IntToStr(i));
  end;
  oleMapInfo.do('commit table bsc_tch_Tqa');
  oleMapInfo.do('add map auto layer bsc_tch_Tqa');
  oleMapInfo.do('Set Map Layer bsc_tch_Tqa Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With Tqa+"%" Auto On Visibility Zoom (0,1000) Units "km"');
  oleMapInfo.do('set map layer cell display off');


  //弱信号断线
 { if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into dss4Tmp');
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
      oleMapInfo.do('Select CCh_file.Cell_id,cch_file.Dss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from cch_file, cell '
                 + 'where cch_file.cell_id = cell.bs_no into dss4Tmp');
  end; }

 { oleMapInfo.do('Select bsc_CCh_file.bsc_no,bsc_cch_file.Dss4, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_cch_file, bsc '
                 + 'where bsc_cch_file.bsc_no = bsc.bsc_no into dss4Tmp');

  oleMapInfo.do('commit table dss4Tmp as "' + gExePath + 'bsc_cch_dss4.tab"');
  oleMapInfo.do('close table dss4Tmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_cch_dss4.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,255,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,255,255),rgb(0,255,255))');
  wRow := oleMapInfo.eval('tableinfo(bsc_cch_dss4, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_cch_dss4');
    wLon := oleMapInfo.eval('bsc_cch_dss4.lon');
    wLat := oleMapInfo.eval('bsc_cch_dss4.lat');
    //wBearing := oleMapInfo.eval('bsc_cch_dss4.Bearing');
    wRate := oleMapInfo.eval('bsc_cch_dss4.dss4') / 100;
    if wRate > 0 then
      wRate := 0.02 + 0.6 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 1, 'bsc_cch_dss4');
  end;
  oleMapInfo.do('commit table bsc_cch_dss4');
  oleMapInfo.do('add map auto layer bsc_cch_dss4');
  oleMapInfo.do('Set Map Layer bsc_cch_dss4 Label Position Above Font ("Arial",256,8,16777215,0) ' +
               ' With dss4+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  {if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no and cell.bsc_no = "'
                 + gSelName + '"into tss4Tmp');
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
      oleMapInfo.do('Select TCh_file.Cell_id,Tch_file.tss4, Cell.cell_name, cell.Lon,'
                 + 'cell.Lat, cell.Bearing from tch_file, cell '
                 + 'where tch_file.cell_id = cell.bs_no into tss4Tmp');
  end;}

  {oleMapInfo.do('Select bsc_TCh_file.bsc_no,bsc_Tch_file.tss4, bsc.bsc_name, bsc.Lon,'
                 + 'bsc.Lat from bsc_tch_file, bsc '
                 + 'where bsc_tch_file.bsc_no = bsc.bsc_no into tss4Tmp');
  oleMapInfo.do('commit table tss4Tmp as "' + gExePath + 'bsc_tch_tss4.tab"');
  oleMapInfo.do('close table tss4Tmp ');
  oleMapInfo.do('Open table "' + gExePath + 'bsc_tch_tss4.tab" Interactive');


  //oleMapInfo.do('Set Map Layer 1 Editable On');
  oleMapInfo.do('set style pen makepen(1,2, rgb(255,255,0))');
  oleMapInfo.do('set style brush makebrush(64,rgb(255,255,0),rgb(255,255,0))');
  wRow := oleMapInfo.eval('tableinfo(bsc_tch_tss4, 8)');
  for i := 1 to wRow do
  begin
    oleMapInfo.do('fetch rec ' + IntToStr(i) +' from bsc_tch_tss4');
    wLon := oleMapInfo.eval('bsc_tch_tss4.lon');
    wLat := oleMapInfo.eval('bsc_tch_tss4.lat');
    //wBearing := oleMapInfo.eval('bsc_tch_tss4.Bearing');
    wRate := oleMapInfo.eval('bsc_tch_tss4.tss4') / 100;
    if wRate > 0 then
      wRate := 0.02 + 0.6 * wRate;
    BscUpDateCircle(wLon, wLat, 0.02, wRate, oleMapInfo, i, 4, 'bsc_tch_tss4');
  end;
  oleMapInfo.do('commit table bsc_tch_tss4');
  oleMapInfo.do('add map auto layer bsc_tch_tss4');
  oleMapInfo.do('Set Map Layer bsc_tch_tss4 Label Position Above Font ("Arial",256,8,16777215,0) ' +
                ' With tss4+"%" Auto On Visibility Zoom (0, 6) Units "km"');

  oleMapInfo.do('set map redraw on');
  gBscDqa := True;
  fmBscMain.sbBscMain.Panels[0].Text := '';
end;  }

{procedure TfmMap.mmRafClick(Sender: TObject);
procedure raf_shade;
begin
  if gSelFlag = 'BSC' then
  begin
    oleMapInfo.do('Select * from bsc where cell.bsc_no = "'
                 + gSelName + '" into tmp');
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
    else
      oleMapInfo.do('Select * from cell into tmp');
  end; }
  {oleMapInfo.do('Select * from bsc into tmp');
  oleMapInfo.do('commit table tmp as "' + gExePath + 'bsc_raf_shade.tab"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('open table "' + gExePath + 'bsc_raf_shade.tab"');
  oleMapInfo.do('add map auto layer bsc_raf_shade');
  oleMapInfo.do('Add Column "bsc_raf_shade" (raf Decimal (8, 2))From bsc_cch_file Set To raf Where COL5 = COL4  Dynamic');
  oleMapInfo.do('Add Column "bsc_raf_shade" (raa Decimal (8, 2))From bsc_cch_file Set To raa Where COL5 = COL4  Dynamic');
  oleMapInfo.do('Add Column "bsc_raf_shade" (ras Decimal (8, 2))From bsc_cch_file Set To ras Where COL5 = COL4  Dynamic');
  oleMapInfo.do('Add Column "bsc_raf_shade" (rac Decimal (8, 2))From bsc_cch_file Set To rac Where COL5 = COL4  Dynamic');
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + ' raf_shade  with Arfcn ' +
                'ignore 0 ranges apply all use color Brush (2,16711680,16777215) '+
                '  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,25: 50 Brush ' +
                ' (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush (2,11554816,16777215) ' +
                ' Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) Pen (1,2,0) ' +
                ' default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.1 ' +
                ' inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 #');
 }
 { oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' bsc_raf_shade with Raf ignore 0 ranges apply all use color Brush ' +
             ' (2,65280,16777215)  0: 25 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             ' 25: 50 Brush (2,5287936,16777215) Pen (1,2,0) ,50: 75 Brush ' +
             ' (2,11554816,16777215) Pen (1,2,0) ,75: 100 Brush (2,16711680,16777215) ' +
             ' Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 ' +
             ' round 0.1 inflect off Brush (2,16777215,16777215) at 2 by 0 color 1 # ');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
   ' layer prev display on shades on symbols off lines off count on title auto '  +
   ' Font ("Arial",0,12,0) subtitle auto Font ("Arial",0,11,0) ascending off ' +
   ' ranges Font ("Arial",0,11,0) auto display off ,auto display on ,auto ' +
   ' display on ,auto display on ,auto display on ');

  oleMapInfo.do('Set Map Layer bsc_Raf_shade Label Position Above Font ("Arial",1,12,0) With Raf+"%,"+raa+"%"+chr$(13)+ras+"%,"+Rac+"%" Auto On Visibility Zoom (0, 1000) Units "km"');
  gbscRaf := True;
end;
begin
  if gBScRaf then
  begin
    if mmRaf.Checked then
    begin
      oleMapInfo.do('close table bsc_raf_shade');
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Off');
      //oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on')   }
    {end
    else
    begin
      raf_shade;
      {oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer Raf_shade Display Graphic');
     // oleMapInfo.do('Set Map Layer RafTmp Display Graphic');
      oleMapInfo.do('set map redraw on')}
    {end;
    mmRaf.Checked := not mmRaf.Checked;
    exit;
  end;
  mmRaf.Checked := not mmRaf.Checked;
  //oleMapInfo.do('set map redraw off');
  fmBscMain.sbBscMain.Panels[0].Text := 'BSC' + TMenuItem(Sender).Caption  + '分析正在进行中...';
  raf_shade;

  fmBscMain.sbBscMain.Panels[0].Text := '';
end; }


{procedure TfmMap.mmCgShadeClick(Sender: TObject);
procedure TchCgShade;
begin
  oleMapInfo.do('commit table bsc as "' + gExePath + 'bsc_cg_shade.tab"');
  oleMapInfo.do('open table "' + gExePath + 'bsc_cg_shade.tab"');
  oleMapInfo.do('add map auto layer bsc_cg_shade');
  oleMapInfo.do('Add Column "bsc_cg_shade" (Cg Decimal (5, 2))From bsc_Tch_file ' +
                 'Set To Cg Where COL5 = COL4  Dynamic');//
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
                 ' bsc_cg_shade with Cg ignore 0 ranges apply all use color ' +
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
  oleMapInfo.do('Set Map Layer bsc_cg_shade Label Position Above Font ("Arial",1,10,0) With cg+"%" Auto On Visibility Zoom (0, 1000) Units "km"');
  gBscTchCG := True;

end;

begin
  if gBscTchCG then
  begin
    if mmCgShade.Checked then
    begin
      oleMapInfo.do('close table bsc_cg_shade');
     { oleMapInfo.do('set map redraw off');
      oleMapInfo.do('Set Map Layer cg_shade Display Off');
     // oleMapInfo.do('Set Map Layer tch_dr Display off');
      oleMapInfo.do('set map redraw on') }
   { end
    else
    begin
      TchCGShade; }
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

procedure TfmMap.tvCellEnter(Sender: TObject);
begin
  paMap.SetFocus;
end;

procedure TfmMap.N4Click(Sender: TObject);
var
  wBsc : String;
begin
  begin
    Application.CreateForm(TfmBScBrow, fmBscBrow);
    with fmBscBrow do
    begin
      with quBscTchBrow do
      begin
        if not Active then
          Open;
      end;
      with quBscCchBrow do
      begin
        if not Active then
          Open;
      end;

      if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'BSC' then
      begin
        oleMapInfo.do('fetch rec ' + IntToStr(1) +' from selection');
        wBSC := oleMapInfo.eval('selection.BSC_NO');
        if fmBscBrow.quBScTchBrow.Active then
          fmBscBrow.quBscTchBrow.Locate('Bsc_no',UpperCase(wBsc),[loPartialKey]);
        if fmBscBrow.quBscCchBrow.Active then
          fmBscBrow.quBscCchBrow.Locate('Bsc_no',UpperCase(wBsc),[loPartialKey]);
      end;
      Show;
    end;

    //wBrowseShow := True;
  end;
end;

procedure TfmMap.lbCellClick(Sender: TObject);
var
  wInt : Integer;
  wCell : String;
begin
  if lbCell.ItemIndex < 0 then
    Exit;
  if fmBscMain.wBrowseShow and gHover and (gSelFlag = 'CELL') then
  begin
    with fmBrowse do
    begin
      with quSumHdovi do
      begin
        if Active then
         Close;
        wInt := Pos(' ', lbCell.Items.Strings[lbCell.ItemIndex]) - 1;
        wCell := Trim(Copy(lbCell.Items.Strings[lbCell.ItemIndex], 1, wInt));
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

      with quSumHdove do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
        Open;
      end;

      with quSumHdoveIn do
      begin
        if Active then
          Close;
        ParamByName('Cell_id').AsString := wCell;
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
      dgHdove.ColumnByName('cell_id_se').FooterValue := wCell;
      dgHdove.ColumnByName('cell_id_ta').FooterValue := '总计:' ;
      quSumHdove.Close;
      quSumHdoveIn.Close;
    end;
  end;
end;

end.
