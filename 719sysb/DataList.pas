unit DataList;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids,  Db, DBTables,  DBGrids, Menus, StdCtrls, Buttons, ExtCtrls,
  ColordDBGrid;

type
  TfmDataList = class(TForm)
    quSelCell: TQuery;
    quDataList: TQuery;
    dsDataList: TDataSource;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    CB5: TComboBox;
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    SpeedButton3: TSpeedButton;
    DataSource1: TDataSource;
    Query1: TQuery;
    PopupMenu2: TPopupMenu;
    cb51: TComboBox;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    DBGrid1: TColordDBGrid;
    DBGrid2: TColordDBGrid;
    Button1: TButton;
    Query2: TQuery;
    Query11: TQuery;
    Query3: TQuery;
    Table1: TTable;
    procedure N3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure DBGrid2drawcoloreddbgrid(sender: TObject; field: TField;
      var color: TColor; var font: TFont);
    procedure FormShow(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
  Procedure Createtable2;
    { Private declarations }
  public
  procedure cbx;
  procedure cbx1;
    { Public declarations }
  end;

var
  fmDataList: TfmDataList;

implementation

{$R *.DFM}
uses BscData,uuppower,ALLCDD, history, uqual_ta;

Procedure TfmDataList.Createtable2;
begin
{  with TTable.Create(Application) do
  begin
    Active := False;

    TableName := 'TEMP1_RLSMP';
    TableType := ttDefault;
    FieldDefs.Add('CELLID', ftString, 12, False);
    FieldDefs.Add('BSCNAME', ftString, 16, False);
    FieldDefs.Add('SIMSG1', ftString, 10, False);
    FieldDefs.Add('SIMSG7', ftString, 10, False);
    FieldDefs.Add('SIMSG8', ftString, 10, False);
    CreateTable;
    Free;
  end; }
end;
procedure TfmDataList.cbx1;
var
  i:integer;
begin
  cb51.Items.Clear;
  with query1 do
  begin
    for i:=0 to FieldCount-1 do
    begin
      Cb51.Items.Add(Fields.Fields[i].FieldName);
    end;
  end;
  Cb51.Text:=dbgrid2.selectedfield.FieldName;
end;

procedure TfmDataList.cbx;
var
  i:integer;
begin
  cb5.Items.Clear;
  with fmDataList.quDataList do
  begin
    for i:=0 to FieldCount-1 do
    begin
      Cb5.Items.Add(Fields.Fields[i].FieldName);
    end;
  end;
  Cb5.Text:=fmDataList.dbgrid1.selectedfield.FieldName;
end;

procedure TfmDataList.N3Click(Sender: TObject);
begin
  try
    fuuppower:=tfuuppower.Create(self);
    fuuppower.showmodal;
  finally
    fuuppower.Free;
  end;
end;

procedure TfmDataList.N1Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
         with quDataList do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select '+wSql);
           SQL.Add('From  '+wTableName +' WHERE RE_DATE=:P');
           sql.add('order by '+by);
           PARAMBYNAME('P').ASSTRING:='2000';
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
  end else
  begin
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
         with quDataList do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select * ');
           SQL.Add('From  temp_RLSMP ');
           sql.add('order by '+by);
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
  end;
end;


procedure TfmDataList.N2Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
         with quDataList do
         begin
           Close;

           SQL.Clear;
           SQL.Add('Select '+wSql);
           SQL.Add('From  '+wTableName +' WHERE RE_DATE=:P ');
           sql.add('order by '+by+' desc');
           PARAMBYNAME('P').ASSTRING:='2000';
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
  end else
  begin
  cbx;
  dbcol:=dbgrid1.SelectedIndex;
  by:=Cb5.Text;
         with quDataList do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select * ');
           SQL.Add('From  temp_rlsmp  ');
           sql.add('order by '+by+' desc');
           open;
         end;
       dbgrid1.SelectedIndex:=dbcol;
       dbgrid1.SetFocus;
  end;
end;

procedure TfmDataList.SpeedButton2Click(Sender: TObject);
begin
  close;
end;

procedure TfmDataList.SpeedButton1Click(Sender: TObject);
begin
cbx;
//11
IF  Cb5.Text='BSPWRB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH��Ƶ���书��');
  END
ELSE IF  Cb5.Text='CGI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С��ȫ��ʶ����');
  END
ELSE IF  Cb5.Text='BSIC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��վʶ����');
  END
ELSE IF  Cb5.Text='BCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH�ز�Ƶ��');
  END
//22
ELSE IF  Cb5.Text='BCCHTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH�������');
  END
ELSE IF  Cb5.Text='AGBLK'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��������������');
  END
ELSE IF  Cb5.Text='MFRMS'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('Ѱ����֡��');
  END
ELSE IF  Cb5.Text='FNOFFSET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('֡ƫ��');
  END
//33
ELSE IF  Cb5.Text='MSTXPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�ƶ�վ����书��');
  END
ELSE IF  Cb5.Text='HOP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ״̬');
  END
ELSE IF  Cb5.Text='HSN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ���к�');
  END
ELSE IF  Cb5.Text='SDCCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('SDCCH/8�ŵ���');
  END
//44
ELSE IF  Cb5.Text='CBCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С���㲥�ŵ�');
  END
ELSE IF  Cb5.Text='ACCMIN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��С�����ƽ');
  END
ELSE IF  Cb5.Text='CCHPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ�����书��');
  END
ELSE IF  Cb5.Text='CRH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ѡ�ͺ�');
  END
//55
ELSE IF  Cb5.Text='NCCPERM'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���������ɫ��');
  END
ELSE IF  Cb5.Text='SIMSG��MSGDIST'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCHϵͳ��Ϣ����');
  END
ELSE IF  Cb5.Text='CB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С�������ֹ');
  END
ELSE IF  Cb5.Text='CBQ'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ֹ����');
  END
//66
ELSE IF  Cb5.Text='ACC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('������Ƶȼ�');
  END
ELSE IF  Cb5.Text='MAXRET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����ط�����');
  END
ELSE IF  Cb5.Text='TX'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���ͷֲ�ʱ϶��');
  END
ELSE IF  Cb5.Text='ATT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('IMSI��Ϸ�������');
  END
//77
ELSE IF  Cb5.Text='T3212'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����λ�ø��¶�ʱ��');
  END
ELSE IF  Cb5.Text='CRO TO PT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ѡƫ�á���ʱƫ�úͳͷ�ʱ��');
  END
ELSE IF  Cb5.Text='EVALTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�㷨����');
  END
ELSE IF  Cb5.Text='RLINKUP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����������·��ʱ');
  END
//88
ELSE IF  Cb5.Text='RLINKT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����������·��ʱ');
  END
ELSE IF  Cb5.Text='NECI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�½�ԭ��ָʾ');
  END
ELSE IF  Cb5.Text='DMPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('MS��̬���ʿ���״̬');
  END
ELSE IF  Cb5.Text='DBPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BTS��̬���ʿ���״̬');
  END
//99
ELSE IF  Cb5.Text='DTXD'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���в���������');
  END
ELSE IF  Cb5.Text='DTXU'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���в���������');
  END
ELSE IF  Cb5.Text='IHO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С�����л�����');
  END
ELSE IF  Cb5.Text='ASSOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('ָ������С������');
  END
//10
ELSE IF  Cb5.Text='MBCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump(' BCCHƵ�ʱ�');
  END
ELSE IF  Cb5.Text='LISTTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('Ƶ�ʱ�����');
  END
ELSE IF  Cb5.Text='ICMSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ�����״̬');
  END
ELSE IF  Cb5.Text='NOALLOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�ŵ����俪��');
  END
//////111
ELSE IF  Cb5.Text='INTAVE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ����ŵ�ƽƽ������');
  END
ELSE IF  Cb5.Text='LIMITn'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���Ŵ��߽�');
  END
ELSE IF  Cb5.Text='MBCR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ��ָʾ');
  END
ELSE IF  Cb5.Text='ECSC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('CLASSMARK���Ϳ���');
  END
ELSE
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('������Ƶȼ�');
END;
cbx1;
//11
IF  Cb51.Text='BSPWRB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH��Ƶ���书��');
  END
ELSE IF  Cb51.Text='CGI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С��ȫ��ʶ����');
  END
ELSE IF  Cb51.Text='BSIC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��վʶ����');
  END
ELSE IF  Cb51.Text='BCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH�ز�Ƶ��');
  END
//22
ELSE IF  Cb51.Text='BCCHTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCH�������');
  END
ELSE IF  Cb51.Text='AGBLK'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��������������');
  END
ELSE IF  Cb51.Text='MFRMS'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('Ѱ����֡��');
  END
ELSE IF  Cb51.Text='FNOFFSET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('֡ƫ��');
  END
//33
ELSE IF  Cb51.Text='MSTXPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�ƶ�վ����书��');
  END
ELSE IF  Cb51.Text='HOP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ״̬');
  END
ELSE IF  Cb51.Text='HSN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ���к�');
  END
ELSE IF  Cb51.Text='SDCCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('SDCCH/8�ŵ���');
  END
//44
ELSE IF  Cb51.Text='CBCH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С���㲥�ŵ�');
  END
ELSE IF  Cb51.Text='ACCMIN'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��С�����ƽ');
  END
ELSE IF  Cb51.Text='CCHPWR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ�����书��');
  END
ELSE IF  Cb51.Text='CRH'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ѡ�ͺ�');
  END
//55
ELSE IF  Cb51.Text='NCCPERM'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���������ɫ��');
  END
ELSE IF  Cb51.Text='SIMSG��MSGDIST'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BCCHϵͳ��Ϣ����');
  END
ELSE IF  Cb51.Text='CB'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С�������ֹ');
  END
ELSE IF  Cb51.Text='CBQ'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ֹ����');
  END
//66
ELSE IF  Cb51.Text='ACC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('������Ƶȼ�');
  END
ELSE IF  Cb51.Text='MAXRET'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����ط�����');
  END
ELSE IF  Cb51.Text='TX'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���ͷֲ�ʱ϶��');
  END
ELSE IF  Cb51.Text='ATT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('IMSI��Ϸ�������');
  END
//77
ELSE IF  Cb51.Text='T3212'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����λ�ø��¶�ʱ��');
  END
ELSE IF  Cb51.Text='CRO TO PT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С����ѡƫ�á���ʱƫ�úͳͷ�ʱ��');
  END
ELSE IF  Cb51.Text='EVALTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�㷨����');
  END
ELSE IF  Cb51.Text='RLINKUP'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����������·��ʱ');
  END
//88
ELSE IF  Cb51.Text='RLINKT'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('����������·��ʱ');
  END
ELSE IF  Cb51.Text='NECI'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�½�ԭ��ָʾ');
  END
ELSE IF  Cb51.Text='DMPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('MS��̬���ʿ���״̬');
  END
ELSE IF  Cb51.Text='DBPSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('BTS��̬���ʿ���״̬');
  END
//99
ELSE IF  Cb51.Text='DTXD'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���в���������');
  END
ELSE IF  Cb51.Text='DTXU'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���в���������');
  END
ELSE IF  Cb51.Text='IHO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('С�����л�����');
  END
ELSE IF  Cb51.Text='ASSOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('ָ������С������');
  END
//10
ELSE IF  Cb51.Text='MBCCHNO'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump(' BCCHƵ�ʱ�');
  END
ELSE IF  Cb51.Text='LISTTYPE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('Ƶ�ʱ�����');
  END
ELSE IF  Cb51.Text='ICMSTATE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ�����״̬');
  END
ELSE IF  Cb51.Text='NOALLOC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�ŵ����俪��');
  END
//////111
ELSE IF  Cb51.Text='INTAVE'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('�����ŵ����ŵ�ƽƽ������');
  END
ELSE IF  Cb51.Text='LIMITn'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('���Ŵ��߽�');
  END
ELSE IF  Cb51.Text='MBCR'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('��Ƶ��ָʾ');
  END
ELSE IF  Cb51.Text='ECSC'  THEN
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('CLASSMARK���Ϳ���');
  END
ELSE
BEGIN
  application.HelpFile:='CDD.hlp';
  application.HelpJump('������Ƶȼ�');
END;
end;

procedure TfmDataList.SpeedButton3Click(Sender: TObject);
VAR
  tt:ttable;
  t1,t2,t3:integer;
  X1,X2,X3:STRING;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
  button1.SetFocus;
  IF SPEEDBUTTON3.Caption='��ʷ����' then
  begin
    dbgrid1.Align:=altop;
    dbgrid1.Height:=90;
    dbgrid2.Align:=alclient;
    dbgrid2.Height:=90;
    with query1 do
    begin
      if Active then
        Close;

      IF Pos('CELLID', WSQL) <>0 then
      BEGIN
        with sql do
        begin
          Clear;
          Add('select ' + wSql +',DATA_CHANGE'+ ' from ' + wTableName +' WHERE RE_DATE=:P  order by cellid');
          PARAMBYNAME('P').ASSTRING:='1999';
        end;
        Open;
      END  ELSE
      BEGIN
        with sql do
        begin
          Clear;
          Add('select ' + wSql +',DATA_CHANGE'+' from ' + wTableName +' WHERE RE_DATE=:P order by BSCNAME');
          PARAMBYNAME('P').ASSTRING:='1999';
        end;
        Open;
      END;

    end;
    SPEEDBUTTON3.Caption:='����';
    DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;

  end  else
  IF SPEEDBUTTON3.Caption='����' then
  begin
    dbgrid1.Align:=altop;
    dbgrid1.Height:=180;
//    dbgrid2.Visible:=false;;
    dbgrid2.Height:=0;
    SPEEDBUTTON3.Caption:='��ʷ����';
  end;
  end else
  begin
  button1.SetFocus;
  IF SPEEDBUTTON3.Caption='��ʷ����' then
  begin
    dbgrid1.Align:=altop;
    dbgrid1.Height:=90;
    dbgrid2.Align:=alclient;
    dbgrid2.Height:=90;

    with query11 do
    begin
      if Active then
        Close;
  {    IF Pos('CELLID', WSQL) <>0 then
      BEGIN
        with sql do
        begin
          Clear;
          Add('select * '+ ' from ' + wTableName +' WHERE RE_DATE=:P  order by cellid');
          PARAMBYNAME('P').ASSTRING:='1999';
        end;
        Open;
      END  ELSE
      BEGIN     }
        with sql do
        begin
          Clear;
          Add('select * '+' from rlsmp' +' WHERE RE_DATE=:P order by BSCNAME');
          PARAMBYNAME('P').ASSTRING:='1999';
        end;
        Open;
   //   END;
   end;

   with Query2 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('delete from temp1_RLSMP');
           execsql;
         end;
  TABLE1.open;
  table1.edit;
  T1:=query11.RecordCount;
  for t2:=1 to t1 do
  begin
    WHILE NOT QUERY11.Eof DO
    BEGIN
    table1.append;
    table1.fieldbyname('CELLID').asstring:=query11.fieldbyname('CELLID').asstring;
    table1.fieldbyname('BSCNAME').asstring:=query11.fieldbyname('BSCNAME').asstring;
    table1.fieldbyname('SIMSG1').asstring:=query11.fieldbyname('MSGDIST').asstring;
    X1:=query11.fieldbyname('DATA_CHANGE').asstring;
    QUERY11.Next;
    table1.fieldbyname('SIMSG7').asstring:=query11.fieldbyname('MSGDIST').asstring;
    X2:=query11.fieldbyname('DATA_CHANGE').asstring;
    QUERY11.Next;
    table1.fieldbyname('SIMSG8').asstring:=query11.fieldbyname('MSGDIST').asstring;
    X3:=query11.fieldbyname('DATA_CHANGE').asstring;
    table1.fieldbyname('CHANGE').asstring:=X1+X2+X3;
    QUERY11.Next;
    END;
  end;

    with QUERY1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select *');
           SQL.Add('From temp1_RLSMP ');
           Open;
         end;
    SPEEDBUTTON3.Caption:='����';
    DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
 end  else
  IF SPEEDBUTTON3.Caption='����' then
  begin
  try
    dbgrid1.Align:=altop;
    dbgrid1.Height:=180;
    dbgrid2.Height:=0;
    SPEEDBUTTON3.Caption:='��ʷ����';
  except
  end;
  end;
  end;
  query11.Close;
  table1.Close;
end;

procedure TfmDataList.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  dbgrid2.Height:=0;
  dbgrid1.Align:=altop;
  dbgrid1.Height:=180;
  SPEEDBUTTON3.Caption:='��ʷ����';
end;

procedure TfmDataList.N4Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
  cbx1;
  dbcol:=dbgrid2.SelectedIndex;
  by:=Cb51.Text;
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select '+wSql+',DATA_CHANGE');
           SQL.Add('From  '+wTableName +' WHERE RE_DATE=:P ');
           sql.add('order by '+by);
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
       dbgrid2.SelectedIndex:=dbcol;
       dbgrid2.SetFocus;
       DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  end else
  begin
  cbx1;
  dbcol:=dbgrid2.SelectedIndex;
  by:=Cb51.Text;
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select  * ');
           SQL.Add('From  temp1_rlsmp ');
           sql.add('order by '+by);
           open;
         end;
       dbgrid2.SelectedIndex:=dbcol;
       dbgrid2.SetFocus;
       DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  end;
end;

procedure TfmDataList.N5Click(Sender: TObject);
var tt:ttable;
    t1,t2,t3,dbcol:integer;
    by:string;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
  cbx1;
  dbcol:=dbgrid2.SelectedIndex;
  by:=Cb51.Text;
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select '+wSql+',DATA_CHANGE');
           SQL.Add('From  '+wTableName+'  WHERE RE_DATE=:P');
           sql.add('order by '+by+' desc');
           PARAMBYNAME('P').ASSTRING:='1999';
           open;
         end;
       dbgrid2.SelectedIndex:=dbcol;
       dbgrid2.SetFocus;
       DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  end else
  begin
  cbx1;
  dbcol:=dbgrid2.SelectedIndex;
  by:=Cb51.Text;
         with query1 do
         begin
           Close;
           SQL.Clear;
           SQL.Add('Select * ');
           SQL.Add('From  temp1_rlsmp  ');
           sql.add('order by '+by+' desc');
           open;
         end;
       dbgrid2.SelectedIndex:=dbcol;
       dbgrid2.SetFocus;
       DBGRID2.Columns.Items[DBGRID2.Columns.Count-1].VISIBLE:=FALSE;
  end;
end;

procedure TfmDataList.N6Click(Sender: TObject);
begin
  fuqual_ta:=tfuqual_ta.Create(self);
  fuqual_ta.showmodal;
  fuqual_ta.Free;
end;

procedure TfmDataList.DBGrid2drawcoloreddbgrid(sender: TObject;
  field: TField; var color: TColor; var font: TFont);
var p,cex,bsx:STRING;
begin
  if wtablename<>'RLSMP' THEN
  BEGIN
    p:=QUERY1.findfield('DATA_CHANGE').asSTRING;
    if p<>''  then
    color:=clred;
  end else
  begin
    p:=QUERY1.findfield('CHANGE').asSTRING;
    if p<>''  then
    color:=clred;
  end;
end;

procedure TfmDataList.FormShow(Sender: TObject);
begin
  SPEEDBUTTON3.Caption:='��ʷ����';
  BUTTON1.SetFocus;
end;

procedure TfmDataList.DBGrid1CellClick(Column: TColumn);
begin
  if SPEEDBUTTON3.Caption<>'��ʷ����' then
  begin
  IF Pos('CELLID', WSQL) <>0 then
  BEGIN
    with QUERY1 do
    begin
       Locate('CELLID',qudatalist.fieldbyname('CELLID').asstring,[loPartialKey]);
    END;
  end else
  begin
    with QUERY1 do
    begin
       Locate('BSCNAME',qudatalist.fieldbyname('BSCNAME').asstring,[loPartialKey]);
    END;
  end;
  end;
end;

procedure TfmDataList.Button1Click(Sender: TObject);
begin
  CLOSE;
end;

procedure TfmDataList.FormCreate(Sender: TObject);
begin
  MoveWindow(Handle, Screen.Width - Width, Screen.Height - Height, Width, Height, True);
end;

end.
