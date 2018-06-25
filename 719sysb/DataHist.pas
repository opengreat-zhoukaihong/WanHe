unit DataHist;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Buttons;

type
  TfmDataHist = class(TForm)
    tvDataHist: TTreeView;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDataHist: TfmDataHist;

implementation

uses BscData, BscMain;

{$R *.DFM}

procedure TfmDataHist.FormCreate(Sender: TObject);
var
  wNode : TTreeNode;
  wDate : String;
begin
  with dmBscData.quDataHist do
  begin
    if not Active then
      Open;
    First;
    wDate := Trim(FieldByName('Start_date').AsString);
    wNode := tvDataHist.Items.Add(nil, wDate);
    while not eof do
    begin
      if wDate <> Trim(FieldByName('Start_date').AsString) then
      begin
        wDate := Trim(FieldByName('Start_date').AsString);
        wNode := tvDataHist.Items.Add(nil, wDate);
      end;
    //  while wDate = Trim(FieldByName('Start_date').AsString) do

       tvDataHist.Items.AddChild(wNode,
          Trim(FieldByName('Start_Time').AsString) + '~' +
          Trim(FieldByName('End_Time').AsString));

      Next;
    end;
    Close;
  end;
end;

procedure TfmDataHist.BitBtn1Click(Sender: TObject);
var
  wDate, wTime, wCellIdSe, wCellIdTa : String;
  i, wRow : Integer;
  wConfDate : TDateTime;
begin
  //for I := 0 to (tvDataHist.Selected.Count - 1) do
  //  ListBox1.Items.Add(tvDataHist.Selected.Item[I].Text);
  if (Not tvDataHist.Selected.HasChildren) and
  	(tvDataHist.Selected.Parent<>nil) then
  begin
    wDate := tvDataHist.Selected.Parent.Text;
    wConfDate := EncodeDate(StrToInt(Copy(wDate,1,4)),
                            StrToInt(Copy(wDate,5,2)),
                            StrToInt(Copy(wDate,7,2)));
    {if wConfDate > 36586 then
    begin
    
      ShowMessage('已过试用期限!');
      Application.Terminate;
    end; }
    wTime := Copy(tvDataHist.Selected.Text,1,4);
   { wRow := oleMapInfo.eval('tableinfo(cch_file, 8)');
    for i := 1 to wRow do
      oleMapInfo.do('delete from cch_file where rowid = ' + IntTostr(i));
    wRow := oleMapInfo.eval('tableinfo(Tch_file, 8)');
    for i := 1 to wRow do
      oleMapInfo.do('delete from Tch_file where rowid = ' + IntTostr(i));
    oleMapInfo.do('commit table cch_file');
    oleMapInfo.do('commit table tch_file'); }
    oleMapInfo.do('close table cch_file');
    oleMapInfo.do('close table tch_file');
    oleMapInfo.do('select * from all_cch_file where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into tmp_cch');
    oleMapInfo.do('commit table tmp_cch as "' + gExePath + 'cch_file.tab"');
    oleMapInfo.do('close table tmp_cch');
    oleMapInfo.do('select * from all_tch_file where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into tmp_tch');
    oleMapInfo.do('commit table tmp_Tch as "' + gExePath + 'Tch_file.tab"');
    oleMapInfo.do('close table tmp_Tch');
    oleMapInfo.do('open table "' + gExePath + 'CCH_FILE.tab"');
    oleMapInfo.do('Create Index On cch_file (Start_time)');
    oleMapInfo.do('Create Index On cch_file (Start_date)');
    oleMapInfo.do('Create Index On cch_file (Bsc_no)');
    oleMapInfo.do('Create Index On cch_file (Cell_id)');

    oleMapInfo.do('open table "' + gExePath + 'TCH_FILE.tab"');
    oleMapInfo.do('Create Index On Tch_file (Start_time)');
    oleMapInfo.do('Create Index On Tch_file (Start_date)');
    oleMapInfo.do('Create Index On Tch_file (Bsc_no)');
    oleMapInfo.do('Create Index On Tch_file (Cell_id)');
    fmBscMain.sbBscMain.Panels[1].Text := IntToStr(oleMapInfo.eval('cch_file.start_date'))
                              + ' [' + IntToStr(oleMapInfo.eval('cch_file.start_Time'))
                              + '-' + IntToStr(oleMapInfo.eval('cch_file.End_Time'))
                              + ']';
    //bsc
    oleMapInfo.do('open table "' + gExePath + 'bsc_all_cch_file.tab"');
    oleMapInfo.do('open table "' + gExePath + 'bsc_all_Tch_file.tab"');
    oleMapInfo.do('select * from bsc_all_cch_file where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into bsc_tmp_cch');
    oleMapInfo.do('commit table bsc_tmp_cch as "' + gExePath + 'bsc_cch_file.tab"');
    oleMapInfo.do('close table bsc_tmp_cch');
    oleMapInfo.do('select * from bsc_all_tch_file where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into bsc_tmp_tch');
    oleMapInfo.do('commit table bsc_tmp_Tch as "' + gExePath + 'bsc_Tch_file.tab"');
    oleMapInfo.do('close table bsc_tmp_Tch');
    oleMapInfo.do('open table "' + gExePath + 'bsc_CCH_FILE.tab"');
    oleMapInfo.do('Create Index On bsc_cch_file (Start_time)');
    oleMapInfo.do('Create Index On bsc_cch_file (Start_date)');
    oleMapInfo.do('Create Index On bsc_cch_file (Bsc_no)');
    //oleMapInfo.do('Create Index On cch_file (Cell_id)');

    oleMapInfo.do('open table "' + gExePath + 'bsc_TCH_FILE.tab"');
    oleMapInfo.do('Create Index On bsc_Tch_file (Start_time)');
    oleMapInfo.do('Create Index On bsc_Tch_file (Start_date)');
    oleMapInfo.do('Create Index On bsc_Tch_file (Bsc_no)');
    //oleMapInfo.do('Create Index On Tch_file (Cell_id)');
    oleMapInfo.do('close table bsc_all_cch_file');
    oleMapInfo.do('close table bsc_all_Tch_file');

    //hdov
    oleMapInfo.do('open table "' + gExePath + 'All_hdov_i_61bsc.tab"');
    oleMapInfo.do('open table "' + gExePath + 'All_hdov_e_61bsc.tab"');
    oleMapInfo.do('select * from all_hdov_i_61bsc where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into tmp_hdov_i_61bsc');
    oleMapInfo.do('commit table tmp_hdov_i_61bsc as "' + gExePath + 'hdov_i_61bsc.tab"');
    oleMapInfo.do('close table tmp_hdov_i_61bsc');
    oleMapInfo.do('select * from all_hdov_e_61bsc where start_date = ' +
                  wDate + ' and start_time = ' + wTime + ' into tmp_hdov_e_61bsc');
    oleMapInfo.do('commit table tmp_hdov_e_61bsc as "' + gExePath + 'hdov_e_61bsc.tab"');
    oleMapInfo.do('close table tmp_hdov_e_61bsc');
    oleMapInfo.do('open table "' + gExePath + 'hdov_i_61bsc.tab"');
    oleMapInfo.do('Create Index On hdov_i_61bsc (Start_time)');
    oleMapInfo.do('Create Index On hdov_i_61bsc (Start_date)');
    oleMapInfo.do('Create Index On hdov_i_61bsc (cell_id_se)');
    oleMapInfo.do('Create Index On hdov_i_61bsc (Cell_id_ta)');

    oleMapInfo.do('open table "' + gExePath + 'hdov_e_61bsc.tab"');
    oleMapInfo.do('Create Index On hdov_e_61bsc (Start_time)');
    oleMapInfo.do('Create Index On hdov_e_61bsc (Start_date)');
    oleMapInfo.do('Create Index On hdov_e_61bsc (cell_id_se)');
    oleMapInfo.do('Create Index On hdov_e_61bsc (Cell_id_ta)');

    {wRow := oleMapInfo.eval('tableinfo(Hdov_i_61bsc,8)');
    oleMapInfo.do('Alter Table "Hdov_i_61bsc" ( Add Hdov_Total Decimal(10,0) ) Interactive');
    oleMapInfo.do('Alter Table "Hdov_e_61bsc" ( Add Hdov_Total Decimal(10,0) ) Interactive');
    oleMapInfo.do(' dim wCell_id_se as string');
    oleMapInfo.do(' dim wCell_id_ta as string');
    oleMapInfo.do(' dim wHovercnt as integer');
    for i := 1 to wRow do
    begin
      oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Hdov_i_61bsc');
      oleMapInfo.do('wcell_id_se = hdov_i_61bsc.cell_id_se');
      oleMapInfo.do('wcell_id_ta = hdov_i_61bsc.cell_id_ta');
      oleMapInfo.do('select hovercnt from hdov_i_61bsc where cell_id_se = wcell_id_ta and cell_id_ta into tmp');
      oleMapInfo.do('wHovercnt = hdov_i_61bsc.hovercnt');
      oleMapInfo.do('wHovercnt = wHovercnt + tmp.hovercnt');
      oleMapInfo.do('close table tmp');
      oleMapInfo.do('update hdov_i_61bsc set Hdov_Total = wHovercnt where rowid = ' + IntToStr(i) );
    end;
    wRow := oleMapInfo.eval('tableinfo(Hdov_e_61bsc,8)');
    for i := 1 to wRow do
    begin
      oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Hdov_e_61bsc');
      oleMapInfo.do('wcell_id_se = hdov_e_61bsc.cell_id_se');
      oleMapInfo.do('wcell_id_ta = hdov_e_61bsc.cell_id_ta');
      oleMapInfo.do('select hovercnt from hdov_e_61bsc where cell_id_se = wcell_id_ta and cell_id_ta into tmp');
      oleMapInfo.do('wHovercnt =  hdov_e_61bsc.hovercnt');
      oleMapInfo.do('wHovercnt = tmp.hovercnt + wHovercnt');
      oleMapInfo.do('close table tmp');
      oleMapInfo.do('update hdov_e_61bsc set Hdov_Total = wHovercnt where rowid = ' + IntToStr(i) );
    end;
    oleMapInfo.do(' undim wCell_id_se ');
    oleMapInfo.do(' undim wCell_id_ta ');
    oleMapInfo.do(' undim wHovercnt ');
    oleMapInfo.do('commit table hdov_i_61bsc');
    oleMapInfo.do('commit table hdov_e_61bsc');}
    oleMapInfo.do('Export "hdov_i_61bsc" Into "' + gExePath + 'hdov_i_61bsc.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
    oleMapInfo.do('Export "hdov_e_61bsc" Into "' + gExePath + 'hdov_e_61bsc.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
    oleMapInfo.do('close table hdov_e_61bsc');
    oleMapInfo.do('close table hdov_i_61bsc');

  end
  else
  begin
    ShowMessage('请选择数据！');
    Abort;
  end;
  oleMapInfo.do('Export "Tch_file" Into "' + gExePath + 'Tch_file.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('Export "Cch_file" Into "' + gExePath + 'Cch_file.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('Export "Bsc_Tch_file" Into "' + gExePath + 'Bsc_Tch_file.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('Export "Bsc_Cch_file" Into "' + gExePath + 'Bsc_Cch_file.dbf" Type "DBF" ' +
                ' Overwrite CharSet "WindowsSimpChinese"');
  oleMapInfo.do('close table bsc_cch_file');
  oleMapInfo.do('close table bsc_Tch_file');

  oleMapInfo.RunMenuCommand(610);
end;

procedure TfmDataHist.Button1Click(Sender: TObject);
var
  I : Integer;
begin
  {If (Not TreeView1.Selected.HasChildren) and
  	(TreeView1.Selected.Parent<>nil) then
   begin
         tstext := TreeView1.Selected.Parent.Text;
         parentnode:=TreeView1.Selected.Parent; //node
         tntext := TreeView1.Selected.Text;
         childnode:=TreeView1.Selected;
  //tvDataHist.
  for I := 0 to (tvDataHist.Selected.Count - 1) do
    ListBox1.Items.Add(tvDataHist.Selected.Item[I].Text);    }
  ShowMessage(tvDataHist.Selected.Text);
end;


end.
