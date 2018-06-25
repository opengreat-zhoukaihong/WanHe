unit ColorDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Buttons;

type
  TfmColorDlg = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    rgType: TRadioGroup;
    Panel1: TPanel;
    rgCondition: TRadioGroup;
    rgAreaRange: TRadioGroup;
    cbZip: TCheckBox;
    edBcch: TEdit;
    edCi: TEdit;
    cbRate: TComboBox;
    cbTestData: TComboBox;
    Label1: TLabel;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmColorDlg: TfmColorDlg;

implementation
uses ObjDlg;

{$R *.DFM}

procedure TfmColorDlg.BitBtn2Click(Sender: TObject);
begin
  wCreated := False;
  Close;
end;

procedure TfmColorDlg.FormCreate(Sender: TObject);
var
  wLength, wPos, i : integer;
begin
  wExePath := UpperCase(Application.ExeName);
  //ShowMessage(wExePath);
  wLength := Length(wExePath);
  for i := 1 to wLength do
  begin
    if wExePath[i] = '\' then
      wPos := i;
  end;
  wExePath := Copy(wExePath, 1, wPos - 1);
  //wExePath := 'C:\My Documents\antcell';
  {if wUserType = '' then
    oleMapInfo := CreateOleObject('MapInfo.Application');}

end;

procedure TfmColorDlg.BitBtn1Click(Sender: TObject);
var
  wTableName, wGridName, wOrder : String;
  wCellObjPos, wGridPos, i , wTableNum, wRate, wRow: Integer;
begin
  wCreated := True;
  wTableName := cbTestData.Text;
  if cbZip.Checked then
  begin
    //oleMapInfo.do('Close table ' + wTableName );
    //oleMapInfo.do('
    wRate := Round(100 / (StrToFloat(Copy(cbRate.Text, 1, Length(cbRate.Text) -1))));
    wRow := oleMapInfo.eval('TableInfo(' + wTableName + ' , 8)');
    for i := 1 to wRow do
    begin
      if (i mod wRate) <> 0 then
        oleMapInfo.do('delete from ' + wTableName + ' where rowid = ' + IntToStr(i));
    end;
    oleMapInfo.do('Commit table ' +  wTableName );
  end;
  case rgType.ItemIndex of
  0 : begin
        if rgAreaRange.ItemIndex = 0 then
        begin
          if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'CELLOBJ' then
          begin
            oleMapInfo.do('Commit table Selection as "'
                     + wExePath + '\map\AreaRange.Tab"');
          end
          else
          begin
            oleMapInfo.do('select * from cellobj where ci in ' +
                   ' (select ci_serv from  ' + wTableName +
                   ' group by ci_serv) into tmp ');
            oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\AreaRange.Tab"');
            oleMapInfo.do('Close table tmp');
          end;
          oleMapInfo.do('Open table  "'
                     + wExePath + '\map\AreaRange.Tab"');
          oleMapInfo.do('set style symbol MakeSymbol(35,12632256,2)');


          case rgCondition.ItemIndex of
          -1: begin
                oleMapInfo.do('commit table ' + wTableName + ' as "' +
                     wExePath + '\map\Rxlev_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');

              {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) select lon, lat, 63 from AreaRange');
                oleMapInfo.do('update Rxlev_test set obj = createpoint(lon,lat)');
              end;
          0 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Bcch_serv = ' +  edBcch.Text + ' into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Rxlev_test.Tab"');
                oleMapInfo.do('Close table tmp');
                //oleMapInfo.do('commit table ' + wTableName + ' as "' +
                 //    wExePath + '\map\Rxlev_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //       ' select lon, lat, 63 from AreaRange where Bcch_serv = ' +  edBcch.Text );
              end;
          1 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Ci_serv = "' +  edBcch.Text + '" into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Rxlev_test.Tab"');
                oleMapInfo.do('Close table tmp');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //              ' select lon, lat, 63 from AreaRange where Ci = "' +  edBcch.Text + '"');
              end;
          end;

         // oleMapInfo.do('Open Table "' +  wExePath
         //              + '\map\Rxlev_test.TAB" Interactive');

          {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
          //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) select lon, lat, 63 from AreaRange');
                  //oleMapInfo.do('fetch rec 1 from Rxlev_test');
                  //oleMapInfo.do('RxlevObject = Rxlev_test.obj');

          oleMapInfo.do('Commit table Rxlev_test');
          oleMapInfo.do('create grid from Rxlev_test ' +
                    ' with Rxlev_s into "' + wExePath +
                    '\MAP\Rxlev_grid" CoordSys Earth Projection 1, 0 clipping ' +
                    ' table AreaRange inflect 5 at RGB(255, 0, 0) : 10 ' +
                    ' RGB(255, 255, 0) : 20 RGB(0, 255, 0) : 30 RGB(0, 255, 255) ' +
                    ' : 40 RGB(0, 0, 255) : 50 cell min 200 border 193 interpolate ' +
                    ' with "IDW" version "100" using 6 "BORDER":  "193" "CELL SIZE":' +
                    ' "0.00675" "EXPONENT":  "2" "MAX POINTS":  "25" "MIN POINTS": ' +
                    '  "1" "SEARCH RADIUS":  "193"');
          oleMapInfo.do('Open Table "' +  wExePath + '\map\Rxlev_grid.TAB" Interactive');

          oleMapInfo.do('add map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  + ' auto layer Rxlev_grid');
          oleMapInfo.do('set map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' layer prev contrast 50 brightness 50 grayscale off');
          oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + '  layer prev display on shades ' +
                        ' on symbols off lines off count off title "路测覆盖场强分布" Font ("宋体",0,9,0) ' +
                        '  subtitle "单位：-dBm" Font ("宋体",0,9,255) ascending off ranges Font ' +
                        ' ("宋体",0,9,0) "10  微弱" display on , "20  弱" display on ,' +
                        ' "30  一般" display on ,"40  中" display on ,"50  强" display on');
                 // oleMapInfo.do('Create Cartographic Legend From Window ' +
                 //   IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' Behind Frame From Layer 3 ');
          oleMapInfo.do('Close table AreaRange');
          oleMapInfo.do('Close table Rxlev_test');
          oleMapInfo.do('Set Map Window FrontWindow() Zoom Entire Layer Rxlev_grid' );
        end
        else
        begin

        end;
        wGridName := 'Rxlev_grid';
      end;
  1 : begin
        if rgAreaRange.ItemIndex = 0 then
        begin
          if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'CELLOBJ' then
          begin
            oleMapInfo.do('Commit table Selection as "'
                     + wExePath + '\map\AreaRange.Tab"');
          end
          else
          begin
            oleMapInfo.do('select * from cellobj where ci in ' +
                   ' (select ci_serv from  ' + wTableName +
                   ' group by ci_serv) into tmp ');
            oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\AreaRange.Tab"');
            oleMapInfo.do('Close table tmp');
          end;
          oleMapInfo.do('Open table  "'
                     + wExePath + '\map\AreaRange.Tab"');
          oleMapInfo.do('set style symbol MakeSymbol(35,12632256,2)');

          case rgCondition.ItemIndex of
          -1: begin
                oleMapInfo.do('commit table ' + wTableName + ' as "' +
                     wExePath + '\map\Rxqual_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxqual_test.TAB" Interactive');

              {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
               // oleMapInfo.do('insert into Rxqual_test (Lon, Lat, Rxlev_s) select lon, lat, 63 from AreaRange');
               // oleMapInfo.do('update Rxqual_test set obj = createpoint(lon,lat)');
              end;
          0 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Bcch_serv = ' +  edBcch.Text + ' into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Rxqual_test.Tab"');
                oleMapInfo.do('Close table tmp');
                //oleMapInfo.do('commit table ' + wTableName + ' as "' +
                 //    wExePath + '\map\Rxlev_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxqual_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //       ' select lon, lat, 63 from AreaRange where Bcch_serv = ' +  edBcch.Text );
              end;
          1 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Ci_serv = "' +  edBcch.Text + '" into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Rxqual_test.Tab"');
                oleMapInfo.do('Close table tmp');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxqual_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //              ' select lon, lat, 63 from AreaRange where Ci = "' +  edBcch.Text + '"');
              end;
          end;

         // oleMapInfo.do('Open Table "' +  wExePath
         //              + '\map\Rxlev_test.TAB" Interactive');

          {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
          //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) select lon, lat, 63 from AreaRange');
                  //oleMapInfo.do('fetch rec 1 from Rxlev_test');
                  //oleMapInfo.do('RxlevObject = Rxlev_test.obj');
          oleMapInfo.do('Alter Table "Rxqual_test" ( modify Rxqual_s Integer ) Interactive');
          oleMapInfo.do('Commit table Rxqual_test');
          oleMapInfo.do('create grid from Rxqual_test ' +
                    ' with Rxqual_s into "' + wExePath +
                    '\MAP\Rxqual_grid" CoordSys Earth Projection 1, 0 clipping ' +
                    ' table AreaRange inflect 5 at RGB(0, 0, 255) : 1 ' +
                    ' RGB(0, 255, 255) : 2 RGB(0, 255, 0) : 3 RGB(255, 255, 0) ' +
                    ' : 4 RGB(255, 0, 0) : 5 cell min 200 border 193 interpolate ' +
                    ' with "IDW" version "100" using 6 "BORDER":  "193" "CELL SIZE":' +
                    ' "0.00675" "EXPONENT":  "2" "MAX POINTS":  "25" "MIN POINTS": ' +
                    '  "1" "SEARCH RADIUS":  "193"');
          oleMapInfo.do('Open Table "' +  wExePath + '\map\Rxqual_grid.TAB" Interactive');

          oleMapInfo.do('add map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  + ' auto layer Rxqual_grid');
          oleMapInfo.do('set map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' layer prev contrast 50 brightness 50 grayscale off');
          oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + '  layer prev display on shades ' +
                        ' on symbols off lines off count off title "路测误码分布" Font ("宋体",0,9,0) ' +
                        ' subtitle "单位：RxQual" Font ("宋体",0,9,255) ascending off ranges Font ' +
                        ' ("宋体",0,9,0) auto display on ,auto display on ,' +
                        ' auto display on ,auto display on ,auto display on');
                 // oleMapInfo.do('Create Cartographic Legend From Window ' +
                 //   IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' Behind Frame From Layer 3 ');
          oleMapInfo.do('Close table AreaRange');
          oleMapInfo.do('Close table Rxqual_test');
          oleMapInfo.do('Set Map Window FrontWindow() Zoom Entire Layer Rxqual_grid' );
        end
        else
        begin

        end;
        wGridName := 'Rxqual_grid';
      end;
  2 : begin
        if rgAreaRange.ItemIndex = 0 then
        begin
          if UpperCase(oleMapInfo.eval('SelectionInfo(1)')) = 'CELLOBJ' then
          begin
            oleMapInfo.do('Commit table Selection as "'
                     + wExePath + '\map\AreaRange.Tab"');
          end
          else
          begin
            oleMapInfo.do('select * from cellobj where ci in ' +
                   ' (select ci_serv from  ' + wTableName +
                   ' group by ci_serv) into tmp ');
            oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\AreaRange.Tab"');
            oleMapInfo.do('Close table tmp');
          end;
          oleMapInfo.do('Open table  "'
                     + wExePath + '\map\AreaRange.Tab"');
          oleMapInfo.do('set style symbol MakeSymbol(35,12632256,2)');

          case rgCondition.ItemIndex of
          -1: begin
                oleMapInfo.do('commit table ' + wTableName + ' as "' +
                     wExePath + '\map\Ta_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Ta_test.TAB" Interactive');

              {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                oleMapInfo.do('insert into Ta_test (Lon, Lat, ta) select lon, lat, "0" from AreaRange');
                oleMapInfo.do('update Ta_test set obj = createpoint(lon,lat)');
              end;
          0 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Bcch_serv = ' +  edBcch.Text + ' into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Ta_test.Tab"');
                oleMapInfo.do('Close table tmp');
                //oleMapInfo.do('commit table ' + wTableName + ' as "' +
                 //    wExePath + '\map\Rxlev_test.tab"');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Ta_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //       ' select lon, lat, 63 from AreaRange where Bcch_serv = ' +  edBcch.Text );
              end;
          1 : begin
                oleMapInfo.do('select * from  ' + wTableName +
                   ' where Ci_serv = "' +  edBcch.Text + '" into tmp ');
                oleMapInfo.do('Commit table tmp as "'
                     + wExePath + '\map\Ta_test.Tab"');
                oleMapInfo.do('Close table tmp');
                oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Ta_test.TAB" Interactive');

                {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
                //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) ' +
                //              ' select lon, lat, 63 from AreaRange where Ci = "' +  edBcch.Text + '"');
              end;
          end;

         // oleMapInfo.do('Open Table "' +  wExePath
         //              + '\map\Rxlev_test.TAB" Interactive');

          {oleMapInfo.do('Open Table "' +  wExePath
                       + '\map\Rxlev_test.TAB" Interactive');}
          //oleMapInfo.do('insert into Rxlev_test (Lon, Lat, Rxlev_s) select lon, lat, 63 from AreaRange');
                  //oleMapInfo.do('fetch rec 1 from Rxlev_test');
                  //oleMapInfo.do('RxlevObject = Rxlev_test.obj');
          oleMapInfo.do('Commit table Ta_test');
          oleMapInfo.do('Alter Table "Ta_test" ( modify Ta Integer ) Interactive');
          oleMapInfo.do('Commit table Ta_test');
          oleMapInfo.do('create grid from Ta_test ' +
                    ' with ta into "' + wExePath +
                    '\MAP\ta_grid" CoordSys Earth Projection 1, 0 clipping ' +
                    ' table AreaRange inflect 5 at RGB(0, 0, 255) : 0 ' +
                    ' RGB(0, 255, 255) : 1 RGB(0, 255, 0) : 2 RGB(255, 255, 0) ' +
                    ' : 3 RGB(255, 0, 0) : 4 cell min 200 border 193 interpolate ' +
                    ' with "IDW" version "100" using 6 "BORDER":  "193" "CELL SIZE":' +
                    ' "0.00675" "EXPONENT":  "2" "MAX POINTS":  "25" "MIN POINTS": ' +
                    '  "1" "SEARCH RADIUS":  "193"');
          oleMapInfo.do('Open Table "' +  wExePath + '\map\Ta_grid.TAB" Interactive');
          oleMapInfo.do('add map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  + ' auto layer Ta_grid');
          oleMapInfo.do('set map window ' + IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' layer prev contrast 50 brightness 50 grayscale off');
          oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) + '  layer prev display on shades ' +
                        ' on symbols off lines off count off title "路测TA分布" Font ("Arial",0,9,0) ' +
                        ' subtitle "单位：公里" Font ("Arial",0,8,0) ascending off ranges Font ' +
                        ' ("Arial",0,11,0) auto display on ,auto display on ,' +
                        ' auto display on ,auto display on ,auto display on');
                 // oleMapInfo.do('Create Cartographic Legend From Window ' +
                 //   IntToStr(oleMapInfo.Eval('FrontWindow()'))  +  ' Behind Frame From Layer 3 ');
          oleMapInfo.do('Close table AreaRange');
          oleMapInfo.do('Close table Ta_test');
        end
        else
        begin

        end;
      end;
  end;
  wTableNum := oleMapInfo.eval('MapperInfo(FrontWindow(), 9)');

  for i := 1 to wTableNum do
  begin
    if UpperCase(Trim(oleMapInfo.eval('LayerInfo(FrontWindow(),' + IntToStr(i)
         + ', 1)'))) = UpperCase(wGridName) then
      wGridPos := i;
    if UpperCase(Trim(oleMapInfo.eval('LayerInfo(FrontWindow(),' + IntToStr(i)
         + ', 1)'))) = 'CELLOBJ' then
      wCellObjPos := i;
  end;
  for i := 1 to wCellObjPos do
  begin
    wOrder := wOrder +  IntToStr(i) + ',';
  end;
  oleMapInfo.do('set map order ' +  wOrder + IntToStr(wGridPos));
  wCreated := True;
  oleMapInfo.RunMenuCommand(304);
  Close;
end;

end.
