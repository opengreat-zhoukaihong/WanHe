unit Compare;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, ExtCtrls, Buttons, Mask, wwdbedit, Wwdbspin;

type
  TfmCompare = class(TForm)
    Panel1: TPanel;
    rgDateType: TRadioGroup;
    Panel2: TPanel;
    gbDateType: TGroupBox;
    gbMonthType: TGroupBox;
    dtpObjDate: TDateTimePicker;
    dtpCompDate: TDateTimePicker;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    wwDBSpinEdit1: TwwDBSpinEdit;
    ComboBox2: TComboBox;
    wwDBSpinEdit2: TwwDBSpinEdit;
    gbWeekType: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    Label4: TLabel;
    Label3: TLabel;
    cbType: TComboBox;
    cbParameter: TComboBox;
    Label7: TLabel;
    Label8: TLabel;
    cbTchParm: TComboBox;
    cbCchParm: TComboBox;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbTypeChange(Sender: TObject);
  private

    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmCompare: TfmCompare;
 

implementation

{$R *.DFM}
uses BscMain;




procedure TfmCompare.BitBtn1Click(Sender: TObject);
var
  wStr, wStartDate1, wStartDate2, wEndDate1, wEndDate2, wFieldName, wType : String;
  i, wRow : Integer;
  wLon, wLat, wBearing, wRate , wMaxQty: real;
begin
  if gCompLayer <> '' then
    oleMapInfo.do('close table ' + gCompLayer);
  oleMapInfo.do('set style pen makepen(1,2, rgb(0,0,255))');
  oleMapInfo.do('set style brush makebrush(64,rgb(0,0,255),rgb(0,0,255))');
  case rgDateType.ItemIndex of
  0: begin
       wStartDate1 := FormatDateTime('yyyymmdd', dtpObjDate.DateTime);
       wStartDate2 := FormatDateTime('yyyymmdd', dtpCompDate.DateTime);
       wEndDate1 := wStartDate1;
       wEndDate2 := wStartDate2;
     end;
  1: ShowMessage('1');
  2: ShowMessage('2');
  end;
  wType := Trim(cbType.Text);
  if wType = 'SDCCH' then
    wType := 'CCH';
  wFieldName := Trim(cbParameter.Text);
  Delete(wFieldName, 1, Pos(' ', wFieldName));
  wFieldName := Trim(wFieldName);
  fmBscMain.sbBscMain.Panels[0].Text := cbType.Text + Copy(cbParameter.Text, 1, Pos(' ',cbParameter.Text) -1 )
                    + '比较分析正在进行中...';
  if not gBscShowing then
  begin
      oleMapInfo.do('Select bsc_no, Cell_id, avg(' + wFieldName + ') from all_' + wType + '_file '
                  + 'where start_date >= ' + wStartDate1
                 + ' and  start_date <= ' + wEndDate1 + ' group by bsc_no, cell_id '
                 + ' into tmp1' );
      oleMapInfo.do('commit table tmp1 as "' + gExePath + 'comp_' + wType +'_' + wFieldName + '_tmp.tab"');
      oleMapInfo.do('close table tmp1');
      oleMapInfo.do('open table "' + gExePath + 'comp_' + wType +'_' + wFieldName + '_tmp.tab"');
      oleMapInfo.do('Alter Table "COMP_' + wType +'_' + wFieldName + '_tmp" ( modify _COL3 Decimal(8,2) ) Interactive');
      oleMapInfo.do('Alter Table "comp_' + wType +'_' + wFieldName + '_tmp" ( rename _COL3 obj_' + wFieldName + ' ) Interactive');

      oleMapInfo.do('Select bsc_no, Cell_id, avg(' + wFieldName + ') from all_' + wType + '_file '
                 + 'where start_date >= ' + wStartDate2
                 + ' and  start_date <= ' + wEndDate2 + ' group by bsc_no, cell_id '
                 + ' into tmp2' );
      oleMapInfo.do('Alter Table "comp_' + wType +'_' + wFieldName + '_tmp" ( add percent Float ) Interactive ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Comp_' + wFieldName +' Decimal (12, 2) )From tmp2 Set To col3 Where COL2 = COL2  Dynamic ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Lon Decimal (12, 6) )From cell Set To lon Where COL2 = COL2  Dynamic ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Lat Decimal (12, 6) )From cell Set To lat Where COL2 = COL2  Dynamic ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Bearing Decimal (12, 6) )From cell Set To bearing Where COL2 = COL2  Dynamic ');


      oleMapInfo.do('Create Map For comp_' + wType +'_' + wFieldName + '_tmp CoordSys Earth Projection 1, 0');

      wRow :=  oleMapInfo.eval('tableInfo(comp_' + wType +'_' + wFieldName + '_tmp,8)');
      for i := 1 to wRow do
      begin
        oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Comp_' + wType +'_' + wFieldName + '_tmp');
        wLon := oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Lon');
        wLat := oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Lat');
        wBearing := oleMapInfo.eval('comp_' + wType +'_' + wFieldName + '_tmp.Bearing');
        if  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName) <> 0 then
          wRate := (oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Obj_' + wFieldName) -
                  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName)) /
                  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName)
        else
          if oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Obj_' + wFieldName) > 0 then
            wRate := 1
          else
            wRate :=  0;

        if Pos('5', oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.cell_id')) > 0 then
          fmBscMain.CreateRegion_3(wLon, wLat, gCellLength / 2, gCellLength, wBearing, wRate, oleMapInfo)
        else
          fmBscMain.CreateRegion_3(wLon, wLat, gCellLength, gCellLength, wBearing, wRate, oleMapInfo);
        oleMapInfo.do('update comp_' + wType +'_' + wFieldName + '_tmp set obj = TmpObject,percent = ' +
                     Format('%5.2f',[100 * wRate]) + ' where rowid = ' + IntToStr(i));
      end;

      oleMapInfo.do('commit table comp_' + wType +'_' + wFieldName + '_tmp');
      oleMapInfo.do('add map auto layer comp_' + wType +'_' + wFieldName + '_tmp');
      oleMapInfo.do('Set Map Layer Comp_' + wType +'_' + wFieldName + '_tmp Label Position Above Font ("Arial",256,8,0,16777215) ' +
                ' With Percent+"%"+chr$(13)+"("+Obj_' + wFieldName + '+","+Comp_' + wFieldName  + '+")" Auto On Visibility Zoom (0, 6) Units "km"');
      oleMapInfo.do('close table tmp2');

      gCompLayer := 'Comp_' + wType +'_' + wFieldName + '_tmp';
  end
  else
  begin//bsc
      oleMapInfo.do('Open table "' + gExePath + 'bsc_all_' + wType + '_file.tab" Interactive');
      oleMapInfo.do('Select bsc_no,  avg(' + wFieldName + ') from bsc_all_' + wType + '_file '
                  + 'where start_date >= ' + wStartDate1
                 + ' and  start_date <= ' + wEndDate1 + ' group by bsc_no '
                 + ' into tmp1' );
      oleMapInfo.do('commit table tmp1 as "' + gExePath + 'comp_' + wType +'_' + wFieldName + '_tmp.tab"');
      oleMapInfo.do('close table tmp1');
      oleMapInfo.do('open table "' + gExePath + 'comp_' + wType +'_' + wFieldName + '_tmp.tab"');
      oleMapInfo.do('Alter Table "comp_' + wType +'_' + wFieldName + '_tmp" ( rename _COL2 obj_' + wFieldName + ' ) Interactive');

      oleMapInfo.do('Select bsc_no,  avg(' + wFieldName + ') from all_' + wType + '_file '
                 + 'where start_date >= ' + wStartDate2
                 + ' and  start_date <= ' + wEndDate2 + ' group by bsc_no '
                 + ' into tmp2' );
      oleMapInfo.do('Alter Table "comp_' + wType +'_' + wFieldName + '_tmp" ( add percent Float ) Interactive ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Comp_' + wFieldName +' Float )From tmp2 Set To col2 Where COL1 = COL1  Dynamic ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Lon Decimal (12, 6) )From bsc Set To lon Where COL1 = COL5  Dynamic ');
      oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Lat Decimal (12, 6) )From bsc Set To lat Where COL1 = COL5  Dynamic ');
     // oleMapInfo.do('Add Column "comp_' + wType +'_' + wFieldName + '_tmp" (Bearing Decimal (12, 6) )From cell Set To bearing Where COL1 = COL5  Dynamic ');


      oleMapInfo.do('Create Map For comp_' + wType +'_' + wFieldName + '_tmp CoordSys Earth Projection 1, 0');

      wRow :=  oleMapInfo.eval('tableInfo(comp_' + wType +'_' + wFieldName + '_tmp,8)');
      for i := 1 to wRow do
      begin
        oleMapInfo.do('fetch rec ' + IntToStr(i) +' from Comp_' + wType +'_' + wFieldName + '_tmp');
        wLon := oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Lon');
        wLat := oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Lat');
       // wBearing := oleMapInfo.eval('comp_' + wType +'_' + wFieldName + '_tmp.Bearing');
        if  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName) <> 0 then
          wRate := (oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Obj_' + wFieldName) -
                  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName)) /
                  oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Comp_' + wFieldName)
        else
          if oleMapInfo.eval('Comp_' + wType +'_' + wFieldName + '_tmp.Obj_' + wFieldName) > 0 then
            wRate := 1
          else
            wRate :=  0;


        fmBscMain.CreateRegion_3(wLon, wLat, 0.00025, 0.03, 180, wRate, oleMapInfo);
        oleMapInfo.do('update comp_' + wType +'_' + wFieldName + '_tmp set obj = TmpObject,percent = ' +
                     Format('%5.2f',[100*wRate]) + ' where rowid = ' + IntToStr(i));
      end;

      oleMapInfo.do('commit table comp_' + wType +'_' + wFieldName + '_tmp');
      oleMapInfo.do('add map auto layer comp_' + wType +'_' + wFieldName + '_tmp');
      oleMapInfo.do('Set Map Layer Comp_' + wType +'_' + wFieldName + '_tmp Label Position Above Font ("Arial",256,8,0,16777215) ' +
                ' With Percent+"%"+chr$(13)+"("+Obj_' + wFieldName + '+","+Comp_' + wFieldName  + '+")" Auto On Visibility Zoom (0, 50) Units "km"');
      oleMapInfo.do('close table tmp2');

      gCompLayer := 'Comp_' + wType +'_' + wFieldName + '_tmp';
      oleMapInfo.do('Close table bsc_all_' + wType + '_file');
  end;
  fmBscMain.sbBscMain.Panels[0].Text := '';
end;

procedure TfmCompare.FormShow(Sender: TObject);
begin
 {
  wCchStrings.Add('申请  CA');
  wCchStrings.Add('分配  CS');
  wCchStrings.Add('话音接通率  U');
  wCchStrings.Add('每线话务 ERPAC');
  wCchStrings.Add('拥塞率  CG');
  wCchStrings.Add('平均通话时间 MH');
  wCchStrings.Add('掉话率  DR');
  wCchStrings.Add('存在信道  CH');
  wCchStrings.Add('可用信道  AC');
  wCchStrings.Add('信道损坏率  F');
  wCchStrings.Add('质差断线  TQA');
  wCchStrings.Add('弱信号断线  TSS4');
  wCchStrings.Add('切换成功率  THIS');
  wCchStrings.Add('话务掉话比  ER_DR');
  wCchStrings.Add('时间拥塞率  TG');
  wCchStrings.Add('话务量 TRAFFIC');
  wTchString := TStrings.Create;
  wTchString := cbParameter.Items;}
end;

procedure TfmCompare.cbTypeChange(Sender: TObject);

begin
  if cbType.Text = 'TCH' then
  begin
    cbParameter.Items := cbTchParm.Items;
    cbParameter.Text := cbTchParm.Text;
  end
  else
  begin
    cbParameter.Items := cbCchParm.Items;
    cbParameter.Text := cbCchParm.Text;
  end;

end;

end.
