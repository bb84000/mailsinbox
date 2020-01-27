{****************************************************************************** }
{ settings1 - Modify settings form and record                                   }
{ for mailsinbox application                                                    }
{ bb - sdtp -january 2020                                                       }
{*******************************************************************************}

unit settings1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, laz2_DOM, laz2_XMLRead, laz2_XMLWrite, lazbbutils, registry, mailclients1, Types;

type

  // Define the classes in this Unit at the very start for clarity
  TFSettings = Class;          // This is a forward class definition

  TmailClient = record
    Name: string;
    Command: string;
    Parameters: string;
    Url: boolean;
    Defaut: boolean;
    Tag: boolean;
  end;

  // Settings record management
  TConfig = class
  private
    FOnChange: TNotifyEvent;
    FOnStateChange: TNotifyEvent;
    FOnQuitAlertChange: TNotifyEvent;
    FSavSizePos: Boolean;
    FWState: string;
    FLastUpdChk: Tdatetime;
    FNoChkNewVer: Boolean;
    FStartup: Boolean;
    FStartMini: Boolean;
    FMailClientMini: Boolean;
    FHideInTaskbar: Boolean;
    FRestNewMsg: Boolean;
    FSaveLogs: Boolean;
    FStartupCheck: Boolean;
    FSmallBtns: Boolean;
    FNotifications: Boolean;
    FNoCloseAlert: Boolean;
    FNoQuitAlert: Boolean;
    FDisplayAllAccMsgs: Boolean;
    FSoundFile: String;
    FLangStr: String;
    FMailClient: String;
    FMailClientName: string;
    FMailClientIsUrl: boolean;
    FRestart: Boolean;
    FAppName: String;
    FVersion: String;
    FLastFires: TStringList;
    FNextFires: TstringList;
    function SaveItem(iNode: TDOMNode; sname, svalue: string): TDOMNode;
  public

    constructor Create (AppName: string);
    procedure SetSavSizePos (b: Boolean);
    procedure SetWState (s: string);
    procedure SetLastUpdChk (dt: TDateTime);
    procedure SetNoChkNewVer (b: Boolean);
    procedure SetStartup (b: Boolean);
    procedure SetStartmini (b: Boolean);
    procedure SetMailClientMini(b: boolean);
    procedure SetHideInTaskbar(b: boolean);
    procedure SetRestNewMsg(b: boolean);
    procedure SetSaveLogs(b: boolean);
    procedure SetStartupCheck(b: boolean);
    procedure SetSmallBtns(b: boolean);
    procedure SetNotifications(b: boolean);
    procedure SetNoCloseAlert(b: boolean);
    procedure SetNoQuitAlert(b: boolean);
    procedure SetDisplayAllAccMsgs(b: boolean);
    procedure SetSoundFile(s: string);
    procedure SetLangStr (s: string);
    procedure SetMailClient(s: string);
    procedure SetMailClientName(s: string);
    procedure SetMailClientIsUrl(b: boolean);
    procedure SetVersion(s: string);
    procedure SetRestart(b: boolean);
    procedure SetLastFires(sl: TStringList);
    procedure SetNextFires(sl: TStringList);
    function saveXMLnode(iNode: TDOMNode): Boolean;
    function SaveToXMLfile(filename: string): Boolean;
    function LoadXMLNode(iNode: TDOMNode): Boolean;
    function LoadXMLFile(filename: string): Boolean;
  published
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property OnStateChange: TNotifyEvent read FOnStateChange write FOnStateChange;
    property OnQuitAlertChange: TNotifyEvent read FOnQuitAlertChange write FOnQuitAlertChange;
    property SavSizePos: Boolean read FSavSizePos write SetSavSizePos;
    property WState: string read FWState write SetWState;
    property LastUpdChk: Tdatetime read FLastUpdChk write SetLastUpdChk;
    property NoChkNewVer: Boolean read FNoChkNewVer write SetNoChkNewVer;
    property Startup: Boolean read FStartup write SetStartup;
    property StartMini: Boolean read FStartMini write SetStartMini;
    property MailClientMini: Boolean read FMailClientMini write SetMailClientMini;
    property HideInTaskbar: Boolean read FHideInTaskbar write SetHideInTaskbar;
    property RestNewMsg: Boolean read FRestNewMsg write SetRestNewMsg;
    property SaveLogs: Boolean read FSaveLogs write SetSaveLogs;
    property StartupCheck: boolean read FStartupCheck write SetStartupCheck;
    property SmallBtns: boolean read FSmallBtns write SetSmallBtns;
    property Notifications: boolean read FNotifications write SetNotifications;
    property NoCloseAlert: boolean read FNoCloseAlert write SetNoCloseAlert;
    property NoQuitAlert: boolean read FNoQuitAlert write SetNoQuitAlert;
    property DisplayAllAccMsgs: boolean read FDisplayAllAccMsgs write SetDisplayAllAccMsgs;
    property SoundFile: string read FSoundFile write SetSoundFile;
    property LangStr: String read FLangStr write SetLangStr;
    property MailClient: string read FMailClient write SetMailClient;
    property MailClientName: string read FMailClientName write SetMailClientName;
    property MailClientIsUrl: boolean read FMailClientIsUrl write SetMailClientIsUrl;
    property AppName: string read FAppName write FAppName;
    property Version: string read FVersion write SetVersion;
    property Restart: boolean read FRestart write SetRestart;
    // couple of account UID and unix date converted to strings.
    // separator is '|'
    property LastFires: TstringList read FLastFires write SetLastFires;
    property NextFires: TstringList read FNextFires write SetNextFires;
end;

  { TFSettings }

  TFSettings = class(TForm)
    BtnCancel: TBitBtn;
    BtnMailClient: TSpeedButton;
    BtnSoundFile: TSpeedButton;
    BtnOK: TBitBtn;
    CBMailClient: TComboBox;
    CBMailClientMini: TCheckBox;
    CBHideInTaskBar: TCheckBox;
    CBNoCloseAlert: TCheckBox;
    CBNoQuitAlert: TCheckBox;
    CBDisplayAllAccMsgs: TCheckBox;
    CBSmallBtns: TCheckBox;
    CBSaveLogs: TCheckBox;
    CBNotifications: TCheckBox;
    CBStartupCheck: TCheckBox;
    CBSavSizePos: TCheckBox;
    CBRestNewMsg: TCheckBox;
    CBLangue: TComboBox;
    CBStartup: TCheckBox;
    CBStartMini: TCheckBox;
    CBNoChkNewVer: TCheckBox;
    CBUrl: TCheckBox;
    ESoundFile: TEdit;
    GBSystem: TGroupBox;
    LSoundFile: TLabel;
    LLangue: TLabel;
    LMailClient: TLabel;
    Lstatus: TLabel;
    PnlButtons: TPanel;
    PnLstatus: TPanel;
    BtnPlaySound: TSpeedButton;
    procedure BtnMailClientClick(Sender: TObject);
    procedure CBMailClientChange(Sender: TObject);
    procedure CBMailClientDrawItem(Control: TWinControl; Index: Integer;
      ARect: TRect; State: TOwnerDrawState);
    procedure CBMailClientKeyPress(Sender: TObject; var Key: char);
    procedure CBStartupChange(Sender: TObject);
    procedure ESoundFileChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormWindowStateChange(Sender: TObject);
    procedure LLangueClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    first: boolean;
    function parseMimeApps(filename: string): string;
  public
    Settings: TConfig;
    GMailWeb, OutlookWeb, Win10Mail: string;
    MailClients: Array of TmailClient;
    function GetDefaultMailCllient: string;
    procedure GetMailClientNames(frst: boolean);
  end;

var
  FSettings: TFSettings;

  const
  // GMail URL
  GmailUrl='https://mail.google.com/mail/u/0/#inbox';
  // Outlook.com URL
  OutlookUrl='https://outlook.live.com/mail/0/inbox';
  // Integrated windows 10 mail client
  Win10MailCmd='"C:\Windows\explorer.exe" "shell:AppsFolder\microsoft.windowscommunicationsapps_8wekyb3d8bbwe!microsoft.windowslive.mail"';


implementation

uses mailsinbox1;

{$R *.lfm}

constructor TConfig.Create(AppName: string);
begin
  inherited Create;
  FAppName:= AppName;
end;

procedure TConfig.SetSavSizePos(b: Boolean);
begin
  if FSavSizePos=b then exit;
  FSavSizePos:= b;
  if Assigned(FOnStateChange) then FOnStateChange(Self);
end;

procedure TConfig.SetWState(s: string);
begin
  if FWState=s then exit;
  FWState:= s;
  if Assigned(FOnStateChange) then FOnStateChange(Self);
end;

procedure TConfig.SetLastUpdChk(dt: TDateTime);
begin
  if FLastUpdChk=dt then exit;
  FLastUpdChk:= dt;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetNoChkNewVer(b: Boolean);
begin
  if FNoChkNewVer=b then exit;
  FNoChkNewVer:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;


procedure TConfig.SetStartup (b: Boolean);
begin
  if FStartup=b then exit;
  FStartup:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetStartMini (b: Boolean);
begin
  if FStartMini=b then exit;
  FStartMini:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetMailClientMini (b: Boolean);
begin
  if FMailClientMini=b then exit;
  FMailClientMini:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetHideInTaskbar (b: Boolean);
begin
  if FHideInTaskbar=b then exit;
  FHideInTaskbar := b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetRestNewMsg (b: Boolean);
begin
  if FRestNewMsg=b then exit;
  FRestNewMsg:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetSaveLogs (b: boolean);
begin
  if FSaveLogs=b then exit;
  FSaveLogs:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetStartupCheck (b: boolean);
begin
  if FStartupCheck=b then exit;
  FStartupCheck:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetSmallBtns (b: boolean);
begin
  if FSmallBtns=b then exit;
  FSmallBtns:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetNotifications(b: boolean);
begin
  if FNotifications=b then exit;
  FNotifications:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetNoCloseAlert(b: boolean);
begin
  if FNoCloseAlert=b then exit;
  FNoCloseAlert:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetNoQuitAlert(b: boolean);
begin
  if FNoQuitAlert=b then exit;
  FNoQuitAlert:= b;
  if Assigned(FOnChange) then FOnChange(Self);
  if Assigned(OnQuitAlertChange) then FOnQuitAlertChange(self);
end;

procedure TConfig.SetDisplayAllAccMsgs(b: boolean);
begin
  if FDisplayAllAccMsgs=b then exit;
  FDisplayAllAccMsgs:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetSoundFile (s: string);
begin
  if FSoundFile=s then exit;
  FSoundFile:= s;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetLangStr (s: string);
begin
  if FLangStr=s then exit;
  FLangStr:= s;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetMailClient (s: string);
begin
  if FMailClient=s then exit;
  FMailClient:= s;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetMailClientName (s: string);
begin
  if FMailClientName=s then exit;
  FMailClientName:= s;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetMailClientIsUrl (b: boolean);
begin
  if FMailClientIsUrl=b then exit;
  FMailClientIsUrl:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetVersion(s:string);
begin
  if FVersion=s then exit;
  FVersion:= s;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetRestart(b:boolean);
begin
  if FRestart=b then exit;
  FRestart:= b;
  if Assigned(FOnChange) then FOnChange(Self);
end;


procedure TConfig.SetLastFires(sl: TStringList);
begin
  if FLastFires.text=sl.text then exit;
  FLastFires.text:= sl.text;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TConfig.SetNextFires(sl: TStringList);
begin
  if FNextFires.text=sl.text then exit;
  FNextFires.text:= sl.text;
  if Assigned(FOnChange) then FOnChange(Self);
end;


function TConfig.SaveItem(iNode: TDOMNode; sname, svalue: string): TDOMNode;
begin
  result:= iNode.OwnerDocument.CreateElement(sname);
  result.TextContent:= svalue;
end;

function TConfig.SaveXMLnode(iNode: TDOMNode): Boolean;
var
  j: integer;
  jNode: TDOMNode;
  ts: TStringList;
begin
  ts:= TstringList.Create;
  Try
    TDOMElement(iNode).SetAttribute ('version', FVersion);
    iNode.AppendChild(SaveItem(iNode, 'savsizepos', BoolToString(FSavSizePos)));
    iNode.AppendChild(SaveItem(iNode, 'wstate', FWState));
    iNode.AppendChild(SaveItem(iNode, 'lastupdchk', TimeDateToString(FLastUpdChk)));
    iNode.AppendChild(SaveItem(iNode, 'nochknewver', BoolToString(FNoChkNewVer)));
    iNode.AppendChild(SaveItem(iNode, 'startup', BoolToString(FStartup)));
    iNode.AppendChild(SaveItem(iNode, 'startmini',BoolToString(FStartMini)));
    iNode.AppendChild(SaveItem(iNode, 'mailclientmini',BoolToString(FMailClientMini)));
    iNode.AppendChild(SaveItem(iNode, 'hideintaskbar',BoolToString(FHideInTaskbar)));
    iNode.AppendChild(SaveItem(iNode, 'restnewmsg',BoolToString(FRestNewMsg)));
    iNode.AppendChild(SaveItem(iNode, 'savelogs',BoolToString(FSaveLogs)));
    iNode.AppendChild(SaveItem(iNode, 'startupcheck',BoolToString(FStartupCheck)));
    iNode.AppendChild(SaveItem(iNode, 'smallbtns',BoolToString(FSmallBtns)));
    iNode.AppendChild(SaveItem(iNode, 'notifications',BoolToString(FNotifications)));
    iNode.AppendChild(SaveItem(iNode, 'noclosealert', BoolToString(FNoCloseAlert)));
    iNode.AppendChild(SaveItem(iNode, 'noquitalert', BoolToString(FNoQuitAlert)));
    iNode.AppendChild(SaveItem(iNode, 'DisplayAllAccMsgs', BoolToString(FDisplayAllAccMsgs)));
    iNode.AppendChild(SaveItem(iNode, 'soundfile', FSoundFile));
    iNode.AppendChild(SaveItem(iNode, 'langstr', FLangStr));
    iNode.AppendChild(SaveItem(iNode, 'mailclient', FMailClient));
    iNode.AppendChild(SaveItem(iNode, 'mailclientname', FMailClientName));
    iNode.AppendChild(SaveItem(iNode, 'mailclientisurl', BoolToString(FMailClientIsUrl)));
    iNode.AppendChild(SaveItem(iNode, 'restart', BoolToString(FRestart)));
    jnode:= iNode.AppendChild(SaveItem(iNode, 'lastfires', ''));
    if FLastFires.count > 0 then
       for j:=0 to FLastFires.count-1 do jNode.AppendChild(SaveItem(jNode, 'lastfire', FLastFires.Strings[j]));  ;
    jnode:= iNode.AppendChild(SaveItem(iNode, 'nextfires', ''));
    if FNextFires.count > 0 then
       for j:=0 to FNextFires.count-1 do jNode.AppendChild(SaveItem(jNode, 'nextfire', FNextFires.Strings[j]));  ;
    Result:= true;
  except
    Result:= false;
  end;
  if assigned(ts) then ts.free;
end;

function TConfig.SaveToXMLfile(filename: string): Boolean;
var
  SettingsXML: TXMLDocument;
  RootNode, SettingsNode :TDOMNode;
begin
  result:= false;
  if FileExists(filename)then
  begin
    ReadXMLFile(SettingsXML, filename);
    RootNode := SettingsXML.DocumentElement;
  end else
  begin
    SettingsXML := TXMLDocument.Create;
    RootNode := SettingsXML.CreateElement(lowercase(FAppName));
    SettingsXML.Appendchild(RootNode);
  end;
  SettingsNode:= RootNode.FindNode('settings');
  if SettingsNode <> nil then RootNode.RemoveChild(SettingsNode);
  SettingsNode:= SettingsXML.CreateElement('settings');
  SaveXMLnode(SettingsNode);
  RootNode.Appendchild(SettingsNode);
  writeXMLFile(SettingsXML, filename);
  result:= true;
  if assigned(SettingsXML) then SettingsXML.free;
end;

function TConfig.LoadXMLNode(iNode: TDOMNode): Boolean;
var
  UpCaseSetting: string;
  subNode: TDOMNode;
  fireNode: TDomNode;
  s:string;
  ts: TStringList;
begin
  Result := false;
  ts:= TStringList.Create;
  if (iNode = nil) then exit;
  try
    UpCaseSetting:=UpperCase(iNode.Attributes.Item[0].NodeName);
    if UpCaseSetting='VERSION' then FVersion:= iNode.Attributes.Item[0].NodeValue;
    subNode := iNode.FirstChild;
    while subNode <> nil do
    try
      upCaseSetting:= UpperCase(subNode.NodeName);
      s:= subNode.TextContent;
      if upCaseSetting = 'SAVSIZEPOS' then FSavSizePos:= StringToBool(s);
      if upCaseSetting = 'WSTATE' then  FWState:= s;
      if upCaseSetting = 'LASTUPDCHK' then FLastUpdChk:= StringToTimeDate(s);
      if upCaseSetting = 'NOCHKNEWVER' then FNoChkNewVer:= StringToBool(s);
      if upCaseSetting = 'STARTUP' then FStartup:= StringToBool(s);
      if upCaseSetting = 'STARTMINI' then FStartMini:= StringToBool(s);
      if upCaseSetting = 'MAILCLIENTMINI' then FMailClientMini:= StringToBool(s);
      if upCaseSetting = 'HIDEINTASKBAR' then FHideInTaskbar:= StringToBool(s);
      if upCaseSetting = 'RESTNEWMSG' then FRestNewMsg:= StringToBool(s);
      if upCaseSetting = 'SAVELOGS' then FSaveLogs:= StringToBool(s);
      if upCaseSetting = 'STARTUPCHECK' then FStartupCheck:= StringToBool(s);
      if upCaseSetting = 'SMALLBTNS' then FSmallBtns:= StringToBool(s);
      if upCaseSetting = 'NOTIFICATIONS' then FNotifications:= StringToBool(s);
      if upCaseSetting = 'NOCLOSEALERT' then FNoCloseAlert:= StringToBool(s);
      if upCaseSetting = 'NOQUITALERT' then FNoQuitAlert:= StringToBool(s);
      if upCaseSetting = 'DISPLAYALLACCMSGS' then FDisplayAllAccMsgs:= StringToBool(s);
      if upCaseSetting = 'SOUNDFILE' then FSoundFile:= s;
      if upCaseSetting = 'LANGSTR' then FLangStr:= s;
      if upCaseSetting = 'MAILCLIENT' then FMailClient:= s;
      if upCaseSetting = 'MAILCLIENTNAME' then FMailClientName:= s;
      if upCaseSetting = 'MAILCLIENTISURL' then FMailClientIsUrl:= StringToBool(s);
      if upCaseSetting = 'RESTART' then FRestart:= StringToBool(s);
      // parse child nodes
      if upCaseSetting = 'LASTFIRES' then
      begin
        fireNode := subNode.FirstChild;
        while fireNode <> nil do
        try
          FLastFires.add(fireNode.TextContent);
        finally
          firenode:= firenode.NextSibling;
        end;
      end;
      if upCaseSetting = 'NEXTFIRES' then
      begin
        fireNode := subNode.FirstChild;
        while fireNode <> nil do
        try
          FNextFires.add(fireNode.TextContent);
        finally
          firenode:= firenode.NextSibling;
        end;
      end;
    finally
        subnode:= subnode.NextSibling;
    end;
    result:= true;
  except
    result:= false;
  end;
  if Assigned(ts) then ts.free;
end;

function TConfig.LoadXMLFile(filename: string): Boolean;
var
  SettingsXML: TXMLDocument;
  RootNode,SettingsNode : TDOMNode;
begin
  result:= false;
  if not FileExists(filename) then
  begin
    SaveToXMLfile(filename);
  end;
  ReadXMLFile(SettingsXML, filename);
  RootNode := SettingsXML.DocumentElement;
  SettingsNode:= RootNode.FindNode('settings');
  if SettingsNode= nil then exit;
  LoadXMLnode(SettingsNode);
  If assigned(SettingsNode) then SettingsNode.free;
  result:= true;
end;

{ TFSettings : Settings dialog }

procedure TFSettings.FormCreate(Sender: TObject);
begin
  GMailWeb:= 'GMail Web site';
  OutlookWeb:= 'Outlook Web site';
  Win10Mail:= 'Windows 10 mail';
  Settings:= TConfig.Create('progname');
  Settings.FLastFires:= TstringList.Create;
  Settings.FNextFires:= TstringList.Create;
end;

procedure TFSettings.FormActivate(Sender: TObject);
begin
  if first then
  begin
    GetDefaultMailCllient;
    first:= false;
    // Center buttons in case of width change
    BtnOK.Left:= (PnlButtons.ClientWidth-BtnOK.width*2-20) div 2;
    BtnCancel.Left:= BtnOK.Left+BtnOK.Width+20;
  end;
end;

procedure TFSettings.FormDestroy(Sender: TObject);
begin
  if assigned(Settings.FLastFires) then  Settings.FLastFires.free;
  if assigned(Settings.FNextFires) then  Settings.FNextFires.free;
  if assigned (Settings) then Settings.free;
end;

procedure TFSettings.FormWindowStateChange(Sender: TObject);
begin
  end;

procedure TFSettings.LLangueClick(Sender: TObject);
begin

end;

// Read only until last item (user item)
procedure TFSettings.Timer1Timer(Sender: TObject);
begin
end;


procedure TFSettings.CBStartupChange(Sender: TObject);
begin
  CBStartMini.Enabled:= CBStartup.Checked;
end;

procedure TFSettings.CBMailClientChange(Sender: TObject);
begin
   if CBMailClient.ItemIndex>=0 then CBUrl.checked:= MailClients[CBMailClient.ItemIndex].Url;
  // CBMailClientDrawItem();
end;

// default client is italic, selected client is bold
// combination of two is bold italic

procedure TFSettings.CBMailClientDrawItem(Control: TWinControl; Index: Integer;
  ARect: TRect; State: TOwnerDrawState);
begin
  CBMailClient.Canvas.FillRect(ARect);
  if MailClients[Index].Defaut=true then CBMailClient.Canvas.Font.Style := [fsItalic];
  if MailClients[Index].Tag=True then CBMailClient.Canvas.Font.Style := [fsbold];
  if (MailClients[Index].Defaut=true) and  (MailClients[Index].Tag=True) then
  CBMailClient.Canvas.Font.Style := [fsbold, fsItalic];
  CBMailClient.Canvas.TextOut(ARect.Left,ARect.Top,CBMailClient.Items[Index]);
end;

procedure TFSettings.BtnMailClientClick(Sender: TObject);
begin
   With FMailClientChoose do
   begin
     If ShowModal=mrOk then
     begin
       SetLength(MailClients, length(MailClients)+1);
       MailClients[length(MailClients)-1]:= Default(TmailClient);
       MailClients[length(MailClients)-1].Name:=EName.Text;
       MailClients[length(MailClients)-1].Command:=ECommand.Text ;
       MailClients[length(MailClients)-1].Url:=CBUrl.Checked;
       self.CBUrl.Checked:=CBUrl.Checked;
       CBMailClient.Items.Add(EName.Text);
       CBMailClient.ItemIndex:= length(MailClients)-1;

     end;
   end;
end;

procedure TFSettings.CBMailClientKeyPress(Sender: TObject; var Key: char);
begin
end;

procedure TFSettings.ESoundFileChange(Sender: TObject);
begin
  BtnPlaySound.Enabled:= not (length(ESoundFile.Text)=0);
end;

// Parse apps lists on linux machines

function TFSettings.parseMimeApps(filename: string): string;
var
  sl: TstringList;
  i: integer;
begin
  result:= '';
  sl:= TstringList.Create;
  sl.LoadFromFile(filename);
   //x-scheme-handler/mailto
  for i:= 0 to sl.Count-1 do
  begin
    if pos('x-scheme-handler/mailto=', sl.Strings[i])=1 then
    begin
     result:= Copy(sl.Strings[i], 25, length(sl.Strings[i])-25);
     break;
    end;
  end;
  sl.free;
end;


procedure TFSettings.GetMailClientNames(frst: boolean);
var
   FixedCount: integer;
begin
  // Add fixed clients
  FixedCount:= 2;
  if frst then SetLength(MailClients, FixedCount);
  MailClients[0]:= Default(TmailClient);
  MailClients[0].Name:=GMailWeb;
  MailClients[0].Command:=GmailUrl;
  MailClients[0].Url:=true;
  MailClients[1]:= Default(TmailClient);
  MailClients[1].Name:=OutlookWeb;
  MailClients[1].Command:=OutlookUrl;
  MailClients[1].Url:=true;
  {$IFDEF WINDOWS}
  //If Windows 10
  if FMailsInBox.OsInfo.VerMaj=10 then
  begin
    inc (FixedCount);
    if frst then SetLength(MailClients, FixedCount);
    MailClients[FixedCount-1]:= Default(TmailClient);
    MailClients[FixedCount-1].Name:=Win10Mail;
    MailClients[FixedCount-1].Command:=Win10MailCmd;
    MailClients[FixedCount-1].Url:=false;
  end;
  {$ENDIF}
end;

// Retrieve default mail client and all available mail clients

function TFSettings.GetDefaultMailCllient: string;
var
  Reg: Tregistry;
  MailKey, DefMailName, DefMailSubkey: string;
  MailSubkey: string;
  {$IFDEF Linux}
    A:TStringArray;
    HomeDir: string;
    AppListArr : array of string;
  {$ENDIF}
  SubKeyNames: TStringList;
  i: integer;
  FixedCount, clicount: integer;
  selfound: boolean;
begin
  result:= '';
  first:= false;
  selfound:= false;
  CliCount:=0;
  // Add fixed clients

  GetMailClientNames(true);
  {SetLength(MailClients, FixedCount);
  MailClients[0]:= Default(TmailClient);
  MailClients[0].Name:=GMailWeb;
  MailClients[0].Command:=GmailUrl;
  MailClients[0].Url:=true;
  MailClients[1]:= Default(TmailClient);
  MailClients[1].Name:=OutlookWeb;
  MailClients[1].Command:=OutlookUrl;
  MailClients[1].Url:=true; }
  {$IFDEF WINDOWS}
    //If Windows 10
  {  if FMailsInBox.OsInfo.VerMaj=10 then
    begin
      inc (FixedCount);
      SetLength(MailClients, FixedCount);
      MailClients[FixedCount-1]:= Default(TmailClient);
      MailClients[FixedCount-1].Name:=Win10Mail;
      MailClients[FixedCount-1].Command:=Win10MailCmd;
      MailClients[FixedCount-1].Url:=false;
    end; }
    FixedCount:= length(MailClients);
    Reg:= TRegistry.Create;
    SubKeyNames:= TStringList.Create;
    // D'abord dans HKCU, sinon dans HKLM
    Reg.RootKey := HKEY_CURRENT_USER;
    MailKey:= 'Software\Clients\Mail';
    if Reg.KeyExists(MailKey) then
    begin
      Reg.OpenKeyReadOnly(MailKey);
      DefMailName:= Reg.ReadString('');
      DefMailSubkey:= MailKey+'\'+DefMailName+'\shell\open\command';
      Reg.CloseKey;
      if not (Reg.KeyExists(DefMailSubkey)) then Reg.RootKey := HKEY_LOCAL_MACHINE;
    end else Reg.RootKey :=HKEY_LOCAL_MACHINE;
    if Reg.KeyExists(MailKey) then
    begin
      Reg.OpenKeyReadOnly(MailKey);
      Reg.GetKeyNames(SubKeyNames);
      CliCount:= FixedCount+SubKeyNames.Count;
      SetLength(MailClients, CliCount);
      DefMailName:= Reg.ReadString('');
      DefMailSubKey:= MailKey+'\'+DefMailName+'\shell\open\command';
      Reg.CloseKey;
      if Reg.KeyExists(DefMailSubkey) then  // Get default client
      begin
        Reg.OpenKeyReadOnly(DefMailSubkey);
        result:= Reg.ReadString('');
        Reg.CloseKey;
      end;
      for i:= 0 to SubKeyNames.Count-1 do
      begin
        MailClients[FixedCount+i]:= Default(TmailClient);
        MailClients[FixedCount+i].Name:= SubKeyNames[i];
        MailSubkey:= MailKey+'\'+SubKeyNames[i]+'\shell\open\command';
        Reg.OpenKeyReadOnly(MailSubkey);
        MailClients[FixedCount+i].Command:=  Reg.ReadString('');
        MailClients[FixedCount+i].Url:=false;
        Reg.CloseKey;
      end;
    end;
    SubKeyNames.Free;
    Reg.free;
  {$ENDIF}
  {$IFDEF Linux}
    // Linux return .desktop file, can be executed w/o path info
    FixedCount:= length(MailClients);
    HomeDir:= GetUserDir;
    Setlength(AppListArr, 9);
    AppListArr[0]:= HomeDir+'.local/share/applications/mimeapps.list';
    AppListArr[1]:= HomeDir+'.local/share/applications/mimeinfo.cache';
    AppListArr[2]:= HomeDir+'.local/share/applications/defaults.list';
    AppListArr[3]:= '/usr/local/share/applications/mimeapps.list';
    AppListArr[4]:= '/usr/local/share/applications/mimeinfo.cache';
    AppListArr[5]:= '/usr/local/share/applications/defaults.list';
    AppListArr[6]:= '/usr/share/applications/mimeapps.list';
    AppListArr[7]:= '/usr/share/applications/mimeinfo.cache';
    AppListArr[8]:= '/usr/share/applications/defaults.list';
    For i:= 0 to 8 do
    begin
      if FileExists(AppListArr[i]) then
      begin
        result:= parseMimeApps(AppListArr[i]);
        if length(result) > 0 then
        begin
          CliCount:= FixedCount+1;
          SetLength(MailClients, clicount);
          A:= result.split('.');             // remove .desktop extension
          if length(A[0])>0 then A[0][1]:= UpCase(A[0][1]);
          MailClients[FixedCount]:= Default(TmailClient);
          MailClients[FixedCount].Name:=A[0];
          MailClients[FixedCount].Command:=result;
          MailClients[FixedCount].Url:=false;
          MailClients[FixedCount].Defaut:=true;
          break;
        end;
      end;
    end;
  {$ENDIF}
  for i:= 0 to clicount-1 do
  begin
    if MailClients[i].Command=result then                // default client
    begin
      MailClients[i].Defaut:=true;
    end;
    if MailClients[i].Command=Settings.MailClient then   // Selected client
    begin
      MailClients[i].Tag:=true;
      MailClients[i].Url:=Settings.MailClientIsUrl;
      selfound:= true;
    end;
  end;
  // Now  add user client at the end if is not in list
  if not selfound then
  begin
    SetLength(MailClients, length(MailClients)+1);
    MailClients[length(MailClients)-1]:= Default(TmailClient);
    MailClients[length(MailClients)-1].Name:=Settings.MailClientName;
    MailClients[length(MailClients)-1].Command:=Settings.MailClient;
    MailClients[length(MailClients)-1].Url:=Settings.MailClientIsUrl;
    MailClients[length(MailClients)-1].Tag:=true;
  end;
end;




end.

