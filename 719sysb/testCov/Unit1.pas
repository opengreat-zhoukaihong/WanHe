unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    ListBox1: TListBox;
    ListBox2: TListBox;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Button1Click(Sender: TObject);
var
  i : Integer;
function ConvertLine(oldStr : string): String;
var
	i:Integer;
  szStr : String[5];
  nAscii: char;
  V,Code:Integer;
  nByte:Integer;
  IsSpace:Bool;
  spacenum:integer;
  NewStr : String;
begin
  i := 1;
  result := '';
  nByte:=0;
  IsSpace:= false;
	while i<=Length(OldStr) do
  begin
     spacenum := 0;
  	while OldStr[i] = ' ' do
     begin
     	Inc(i);
        IsSpace := true;
        Inc(spacenum);
     end;

     if(spacenum>3) or
     not(
      (OldStr[i] in ['0'..'9'])
     or (OldStr[i] in ['A'..'Z'])
     or (OldStr[i] in ['a'..'z'])
     )
     then begin
     	result := '';
        exit;
     end;
     if (nByte<4) and IsSpace then     //For Lose Bytes,Add Space
     	NewStr := NewStr+Copy('      ',1,4-nByte);
     if (IsSpace) then
     begin
  	   nByte:=0;
     	IsSpace := false;
     end;
     if(OldStr[i+1]=' ')then
        Insert('0',OldStr,i+1);

     szStr := '$'+ Copy(OldStr,i,2);
     if(szStr='$20') then
     	nAscii := ' '//#7
     else	begin
     Val(szStr,V,Code);
     nAscii := ' ';
		if(Code=0) then
     	nAscii := Chr(V);
     if( nAscii< #32) then
     	nAscii := ' ';
     end;

     NewStr := NewStr+String(nAscii);
     Inc(i,2);
     Inc(nByte);
  end;
  spacenum := Length(NewStr) mod 4;
  if(spacenum<>0) then
  	NewStr := NewStr+copy('     ',1,4-spacenum);
  result := NewStr;
end;
begin
// Edit1.Text := ConvertLine('37380000 00000000 5A484231 42534337 2F332F30 362F3030 2F31332F')
  for i := 0 to listBox2.Items.count - 1 do
  begin
    listbox1.Items.add(ConvertLine(trim(listbox2.items.strings[i])));
  end;
end;

end.
