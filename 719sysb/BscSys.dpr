program BscSys;
uses
  Forms,
  Dialogs,
  SysUtils,
  BscMain in 'BscMain.pas' {fmBscMain},
  CHILDWIN in 'CHILDWIN.PAS' {fmMap},
  About in 'about.pas' {AboutBox},
  legeng in 'legeng.pas' {fmLegeng},
  DataHist in 'DataHist.pas' {fmDataHist},
  BscData in 'BscData.pas' {dmBscData: TDataModule},
  Compare in 'Compare.pas' {fmCompare},
  Find in 'Find.pas' {fmFind},
  MapConf in 'MapConf.pas' {fmMapConf},
  FindNcell in 'FindNcell.pas' {fmFindNCell},
  Trendline in 'Trendline.pas' {fmTrendline},
  Condition in 'Condition.pas' {fmCondition},
  CellConf in 'CellConf.pas' {fmCellConf},
  Browse in 'Browse.pas' {fmBrowse},
  BscBrow in 'BscBrow.pas' {fmBscBrow},
  HdovCond in 'HdovCond.pas' {fmHdovCond},
  AllCdd in 'AllCdd.pas' {fmAllCdd},
  DataList in 'DataList.pas' {fmDataList},
  uuppower in 'uuppower.pas' {fuuppower},
  SelCdd in 'SelCdd.pas' {fmSelCdd},
  detail in 'detail.pas' {fdetail},
  CQTDlg in 'CQTDlg.pas' {fmDensityDlg},
  punit in 'punit.pas' {Fcdd_conv},
  dm in 'dm.pas' {DBData: TDataModule},
  ctr_globe in 'ctr_globe.pas',
  history in 'history.pas' {fhistory},
  uqual_ta in 'uqual_ta.pas' {fuqual_ta},
  AFG_TLB in '..\Program Files\Borland\Delphi4\Imports\AFG_TLB.pas',
  Wireless_TLB in '..\Program Files\Borland\Delphi4\Imports\Wireless_TLB.pas';

{$R *.RES}
var
  gCtrPath : String;

begin
  Application.Initialize;
  {if Now > 36586 then
    Application.Terminate; }
  gExePath := Application.ExeName;
  while gExePath[Length(gExePath)] <> '\' do
  begin
    Delete(gExePath, Length(gExePath), 1);
  end;
  gCtrExePath := gExePath;
  //gPath := gExePath ;
  gCtrPath := gExePath + 'Ctr';
  gExePath := gExePath + 'map\';
  //gExePath := 'c:\mapdata\';
  //ShowMessage(gExePath);
  Application.CreateForm(TdmBscData, dmBscData);
  with dmBscData.dbBscData do
  begin
    Params.Add('PATH=' +   Copy(gExePath, 1, Length(gExePath)-1));
    Connected := True;
  end;
  with dmBscData.dbCtrData do
  begin
    Params.Add('PATH=' +   gCtrPath);
    try
      Connected := True;
    except
      AliasName := 'NqiData';
      Connected := True;
    end;
    if Connected then
    begin
      Application.CreateForm(TfmBscMain, fmBscMain);
      //if AliasName = 'NqiData' then
        //fmBscMain.sbBscMain.Panels[4].Text := '本地数据';
      Application.CreateForm(TAboutBox, AboutBox);
      Application.Run;
    end
    else
    begin
      ShowMessage('连接NQIDATA数据库失败!');
      Application.Terminate;
    end;
  end;
end.
