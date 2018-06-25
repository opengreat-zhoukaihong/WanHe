unit AllCdd;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, ExtCtrls;

type
  TfmAllCdd = class(TForm)
    tvCdd: TTreeView;
    gbCellField: TGroupBox;
    lbCddField: TListBox;
    Panel1: TPanel;
    Panel2: TPanel;
    paDataList: TPanel;
    procedure tvCddDblClick(Sender: TObject);
    procedure lbCddFieldDblClick(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure paDataListClick(Sender: TObject);
    procedure Panel2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmAllCdd: TfmAllCdd;
  wTableName : String;

implementation

uses DataList, BscData ;

{$R *.DFM}

procedure TfmAllCdd.tvCddDblClick(Sender: TObject);
begin

  if tvCdd.Selected.Parent <> nil then
  begin
    if gbCellField.Caption <> Copy(tvCdd.Selected.Parent.Text,1,5) then
    begin
      lbCddField.Items.Clear;
      lbCddField.Items.add('CELLID');
      gbCellField.Caption := Copy(tvCdd.Selected.Parent.Text, 1, 5);
      wTableName := gbCellField.Caption;
    end;
    if lbCddField.Items.IndexOf(tvCdd.Selected.Text) < 0 then
    begin
      lbCddField.Items.Add(tvCdd.Selected.Text);
     // gMultiCell := lbCell.Items;
    end
  end;
end;

procedure TfmAllCdd.lbCddFieldDblClick(Sender: TObject);
begin
  if lbCddField.Items.Count > 0 then
  begin
    if lbCddField.ItemIndex >= 0 then
      lbCddField.Items.Delete(lbCddField.ItemIndex)
    else
      lbCddField.Items.Delete(0);
    //gMultiCell := lbCell.Items;
  end;
end;

procedure TfmAllCdd.FormPaint(Sender: TObject);
begin
  paDataList.Width := Round(fmAllCdd.Width /2);
end;

procedure TfmAllCdd.paDataListClick(Sender: TObject);
var
  wSql : String;
  i : Integer;
begin
//  wTableName := 'RLCFP';
  fmDataList.Caption := wTableName;
  if lbCddField.Items.Count > 0 then
  begin
    wSql := '';
    for i := 0 to  lbCddField.Items.Count - 1 do
    begin
      if i < lbCddField.Items.Count - 1 then
        wSql := wSql + lbCddField.Items.Strings[i] + ','
      else
        wSql := wSql + lbCddField.Items.Strings[i];
    end;
  end;
 
  with fmDataList.quDataList do
  begin
    if Active then
      Close;
    with sql do
    begin
      Clear;
      Add('select ' + wSql + ' from ' + wTableName );

    end;
    Open;
  end;
  fmDataList.Show;
end;

procedure TfmAllCdd.Panel2Click(Sender: TObject);
begin
  lbCddField.Items.Clear;
  gbCellField.Caption := 'Ñ¡Ôñ²ÎÊý';
end;

procedure TfmAllCdd.FormCreate(Sender: TObject);
var
  wTreeNode : TTreeNode;
  i, j, wCount : Integer;

begin
  Application.CreateForm(TfmDataList, fmDataList);
  for i := 0 to tvCdd.Items.Count - 1 do
  begin
    if i > tvCdd.Items.Count - 1 then
      Break;
    wTreeNode := tvCdd.Items[i];
    if wTreeNode.Parent = nil then
    begin
      //ShowMessage(wTreeNode.Text);
      if wTreeNode.HasChildren then
      begin
        while wTreeNode.Count > 0 do
          wTreeNode.Item[0].Delete;
      end;
    end;
  end;
  i := 0;
  while i < wCount do
  begin
    wTreeNode := tvCdd.Items[i];
    if wTreeNode.Parent = nil then
    begin
      if Pos('RLCFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCFP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLCPP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCPP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLCXP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLCXP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLDEP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLDEP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLIHP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLIHP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLLOP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLLOP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLMFP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLMFP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLNRP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLNRP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSBP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSBP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;
      if Pos('RLSSP', wTreeNode.Text) > 0 then
      begin
        with dmBscData.quRLSSP do
        begin
          for j := 0 to FieldCount -1 do
          begin
            tvCdd.Items.AddChild(wTreeNode, Fields[j].FieldName);
          end;
        end;
      end;

    end;
    wCount := tvCdd.Items.Count;
    i := i + 1;
  end;
end;

procedure TfmAllCdd.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  fmDataList.Close;
  fmDataList.free;

end;

end.
