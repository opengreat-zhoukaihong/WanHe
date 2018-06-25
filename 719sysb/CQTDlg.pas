unit CQTDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfmDensityDlg = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label1: TLabel;
    edCount: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDensityDlg: TfmDensityDlg;

implementation
uses BscMain;
{$R *.DFM}

procedure TfmDensityDlg.BitBtn1Click(Sender: TObject);
var
  wErpac : real;
begin
  oleMapInfo.do('select cell_id, Erpac from tch_file order by erpac desc into tmp');
  oleMapInfo.do('fetch rec ' + edCount.Text +' from tmp');
  wErpac := oleMapInfo.eval('tmp.Erpac');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('select tch_file.cell_id, cell.cell_name, tch_file.Erpac from Cell, Tch_file where ' +
               ' cell.Bs_no = tch_file.cell_id and tch_file.erpac >= '
               + FloatToStr(wErpac) + ' into tmp');

  oleMapInfo.do('commit table tmp as "' + gExePath + 'CQT_shade.tab"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('open table "' + gExePath + 'CQT_shade.tab"');
  oleMapInfo.do('add map auto layer CQT_shade');
  //oleMapInfo.do('Add Column "CQT_shade" (Erpac Decimal (8, 2))From Tch_file Set To Erpac Where COL2 = COL6  Dynamic');
  //oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
  oleMapInfo.do('shade window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
             ' CQT_shade with Erpac ignore 0 ranges apply all use color ' +
             ' Brush (2,65280,16777215)  0: 0.2 Brush (2,65280,16777215) Pen (1,2,0) ,' +
             '0.2: 0.5 Brush (2,5287936,16777215) Pen (1,2,0) ,0.5: 0.7 ' +
             'Brush (2,11554816,16777215) Pen (1,2,0) ,0.7: 1.0 ' +
             'Brush (2,16711680,16777215) Pen (1,2,0) default ' +
             'Brush (2,65280,16777215) Pen (1,2,0)  # use 0 round 0.01 inflect off ' +
             'Brush (2,16777215,16777215) at 2 by 0 color 1 #');
  oleMapInfo.do('set legend window ' + IntToStr(oleMapInfo.Eval('FrontWindow()')) +
    ' layer prev display on shades on symbols ' +
    'off lines off count on title "每线话务最大的' + edCount.Text +
    '个小区" Font ("Arial",0,12,0) subtitle auto ' +
    'Font ("Arial",0,11,0) ascending off ranges Font ("Arial",0,11,0) auto ' +
    'display off ,auto display on ,auto display on ,auto display on ,auto ' +
    'display on');

  oleMapInfo.do('Set Map Layer CQT_shade Label Position Above Font ("Arial",1,10,0) With erpac Auto On Visibility Zoom (0, 6) Units "km"');
  gTchTraffic := True;
  oleMapInfo.do('select cell_id  from CQT_shade into tmp');
  oleMapInfo.do('Export "tmp" Into "' + gExePath + 'Tch_Sel_Cell.dbf" Type "DBF" Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table tmp');
  oleMapInfo.do('commit table Cqt_Shade');
  gCqt := True;
end;

end.
