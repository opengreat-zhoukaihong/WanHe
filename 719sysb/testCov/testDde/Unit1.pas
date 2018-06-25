unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, DdeMan;

type
  TForm1 = class(TForm)
    DdeClientConv1: TDdeClientConv;
    DdeClientItem1: TDdeClientItem;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Memo1: TMemo;
    ListBox1: TListBox;
    DdeClientConv2: TDdeClientConv;
    DdeClientItem2: TDdeClientItem;
    OpenDialog1: TOpenDialog;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.BitBtn1Click(Sender: TObject);
var
  Cmd : array[0..255] of Char;
  Cmds : TStrings;
begin
  with DdeClientConv1 do
  begin
   // DdeService := 'WinFIOL';
   // DdeTopic := 'Main';
    SetLink('WinFIOL','Main');
    OpenLink;
    memo1.Text := StrPas(RequestData('Plugin')); //
    {,Information,Status,Output,Event,Login  //chan
      Information,Channels ,ActiveChannel ,Plugin
      Execute "[Quit]"
    }
    StrPCopy(Cmd, '[OpenChannel(C:\Program Files\Element Management\ChnFiles\New channel 0.chn)]');
   // C:\Program Files\Element Management\ChnFiles\New channel 0.chn

    ExecuteMacro(Cmd, False);
    memo1.Text := StrPas(RequestData('Channels'));
   { listbox1.items.Add('[Maximize]');
    listbox1.items.ADD('[Maximize]');
    listbox1.items.Add('[Maximize]');
    listbox1.items.Add('[Quit]');
    ExecuteMacroLines(listbox1.items, false); }
    CloseLink;
  end;
  {Execute "[Minimize]"


This command causes the WinFIOL main window to minimize.


Execute "[Maximize]"


This command causes the WinFIOL main window to maximize.


Execute "[Restore]"


This command restores the WinFIOL main window to its previous size and state.


Execute "[Activate]"


This command activates the WinFIOL main window. If WinFIOL is minimized, it will be restored to its previous size and state.


Execute "[OpenChannel(file)]"


This command opens a new channel by loading the specified channel file. If the channel file does not have a full path specification, the channel file should be present in the channel file directory (you can not define the channel file directory in WinFIOL). The client can check if a new channel is opened by requesting the currently opened channel with the "Channel" data item (see above).


Execute "[OpenChannelParam(parameters)]"


This command opens a new channel and initialises the channel properties as specified by parameters. The parameter string can be a concatenation of channel properties, separated by a comma. The list of possible properties are listed below:

	"protocol=<protocol>", where <protocol> can be "telnet", "rs232serial" or "modem"
	"protocol.telnet.address=<address>", where <address> is an IP address
	"protocol.telnet.port=<port>", where <port> is the IP port number
	"target=<target>, where <target> can be "iog3", "iog11", "iog20", "apg30", "md110" or "eripax"
	"browser=<browser>", where <browser> can be "alexremote", "alexlocal", "dynatext", "krswin" or "docview"

"browser=alexremote.book=<book>, where <book> is the documentation database
	"login.access=<access>, where <access> can be "none", "read only", "write only" or "read and write"
	"login.<prompt>=<response>", where <prompt> can be anything except comma and <response> can be anything except comma.
	"logout.access=<access>, where <access> can be "none", "read only", "write only" or "read and write"
	"logout.<prompt>=<response>", where <prompt> can be anything except comma and <response> can be anything except comma.

Example: "[OpenChannelParam(protocol=telnet,protocol.telnet.address=12.34.56.78,target=iog11)]"

The client can check if a new channel is opened by requesting the currently opened channel with the "Channel" data item (see above). This execute command is new since WinFIOL 5.1.}
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
var
  Cmd : array[0..255] of Char;
  Cmds : TStrings;
begin
  with DdeClientConv2 do
  begin
    SetLink('WinFIOL','Channel #1');
    //DdeTopic := 'Channel #1';
    OpenLink;
    memo1.Text := StrPas(RequestData('Plugin')); //
    {,Information,Status,Output,Event,Login  //chan
      Information,Channels ,ActiveChannel ,Plugin
      Execute "[Quit]"
    }
    if not OpenDialog1.Execute then
      exit;
    StrPCopy(Cmd, '[Transmit(' + OpenDialog1.FileName + ')]');
   // C:\Program Files\Element Management\ChnFiles\New channel 0.chn

    ExecuteMacro(Cmd, False);
    memo1.Text := StrPas(RequestData('Information'));
   { listbox1.items.Add('[Maximize]');
    listbox1.items.ADD('[Maximize]');
    listbox1.items.Add('[Maximize]');
    listbox1.items.Add('[Quit]');
    ExecuteMacroLines(listbox1.items, false); }
    CloseLink;
  end;
  {Execute commands


  The supported execute commands are: "[Close]", "[Minimize]", "[Maximize]",
 "[Restore]", "[Activate]", "[Connect]", "[Release]", "[Break]",
 "[Send(MML cmd)]", "[Transmit(file)]", "[RunMacro(file)]",
 "[Load(file)]", "[Lock(code)]", "[XModemSend(file)]",
  "[XModemReceive(file)]" and "[FindDoc(book,keyword,fault code,category)]".
   Any installed WinFIOL plug-ins may support additional execute commands.
    See further: DDE channel execute commands.  }
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  wStr : array [0..60] of char;
  wPath : String;
begin
{  wPath := 'F:\Program Files\Element Management\WinFIOL\winfiol.exe';
  strpcopy(wStr, wPath);
  if WinExec(wStr, SW_SHOW) < 32 then
  begin
    ShowMessage('CTRProject.EXE²»´æÔÚ£¡');  }
  //end;
end;


end.
