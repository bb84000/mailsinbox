{****************************************************************************** }
{ settings1 - Modify settings form and record                                                }
{ bb - sdtp - november 2019                                                     }
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
    FSavSizePos: Boolean;
    FWState: string;
    FLastUpdChk: Tdatetime;
    FNoChkNewVer: Boolean;
    FStartup: Boolean;
    FStartMini: Boolean;
    FMailClientMini: Boolean;
    FRestNewMsg: Boolean;
    FSaveLogs: Boolean;
    FStartupCheck: Boolean;
    FSmallBtns: Boolean;
    FSoundFile: String;
    FLangStr: String;
    FMailClient: String;
    FMailClientName: string;
    FMailClientIsUrl: boolean;
    FAppName: String;
    FVersion: String;

  public

    constructor Create (AppName: string);
    procedure SetSavSizePos (b: Boolean);
    procedure SetWState (s: string);
    procedure SetLastUpdChk (dt: TDateTime);
    procedure SetNoChkNewVer (b: Boolean);
    procedure SetStartup (b: Boolean);
    procedure SetStartmini (b: Boolean);
    procedure SetMailClientMini(b: boolean);
    procedure SetRestNewMsg(b: boolean);
    procedure SetSaveLogs(b: boolean);
    procedure SetStartupCheck(b: boolean);
    procedure SetSmallBtns(b: boolean);
    procedure SetSoundFile(s: string);
    procedure SetLangStr (s: string);
    procedure SetMailClient(s: string);
    procedure SetMailClientName(s: string);
    procedure SetMailClientIsUrl(b: boolean);
    procedure SetVersion(s: string);
    function SaveXMLnode(iNode: TDOMNode): Boolean;
    function SaveToXMLfile(filename: string): Boolean;
    function LoadXMLNode(iNode: TDOMNode): Boolean;
    function LoadXMLFile(filename: string): Boolean;
  published
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property OnStateChange: TNotifyEvent read FOnStateChange write FOnStateChange;
    property SavSizePos: Boolean read FSavSizePos write SetSavSizePos;
    property WState: string read FWState write SetWState;
    property LastUpdChk: Tdatetime read FLastUpdChk write SetLastUpdChk;
    property NoChkNewVer: Boolean read FNoChkNewVer write SetNoChkNewVer;
    property Startup: Boolean read FStartup write SetStartup;
    property StartMini: Boolean read FStartMini write SetStartMini;
    property MailClientMini: Boolean read FMailClientMini write SetMailClientMini;
    property RestNewMsg: Boolean read FRestNewMsg write SetRestNewMsg;
    property SaveLogs: Boolean read FSaveLogs write SetSaveLogs;
    property StartupCheck: boolean read FStartupCheck write SetStartupCheck;
    property SmallBtns: boolean read FSmallBtns write SetSmallBtns;
    property SoundFile: string read FSoundFile write SetSoundFile;
    property LangStr: String read FLangStr write SetLangStr;
    property MailClient: string read FMailClient write SetMailClient;
    property MailClientName: string read FMailClientName write SetMailClientName;
    property MailClientIsUrl: boolean read FMailClientIsUrl write SetMailClientIsUrl;
    property AppName: string read FAppName write FAppName;
    property Version: string read FVersion write SetVersion;
end;


  { TFSettings }

  TFSettings = class(TForm)
    BtnCancel: TBitBtn;
    BtnMailClient: TSpeedButton;
    BtnSoundFile: TSpeedButton;
    BtnOK: TBitBtn;
    CBMailClient: TComboBox;
    CBMailClientMini: TCheckBox;
    CBSmallBtns: TCheckBox;
    CBSaveLogs: TCheckBox;
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
    LStatus: TLabel;
    PnlButtons: TPanel;
    PnlStatus: TPanel;
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
    procedure Timer1Timer(Sender: TObject);
  private
    first: boolean;
    function parseMimeApps(filename: string): string;
  public
    Settings: TConfig;
    GMailWeb, OutlookWeb, Win10Mail: string;
    MailClients: Array of TmailClient;
    function GetDefaultMailCllient: string;
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
  if FSavSizePos <> b then
  begin
    FSavSizePos:= b;
    if Assigned(FOnStateChange) then FOnStateChange(Self);
  end;
end;

procedure TConfig.SetWState(s: string);
begin
   if FWState <> s then
   begin
     FWState:= s;
     if Assigned(FOnStateChange) then FOnStateChange(Self);
   end;
end;

procedure TConfig.SetLastUpdChk(dt: TDateTime);
begin
   if FLastUpdChk <> dt then
   begin
     FLastUpdChk:= dt;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;

procedure TConfig.SetNoChkNewVer(b: Boolean);
begin
   if FNoChkNewVer <> b then
   begin
     FNoChkNewVer:= b;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;


procedure TConfig.SetStartup (b: Boolean);
begin
   if FStartup <> b then
   begin
     FStartup:= b;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;

procedure TConfig.SetStartMini (b: Boolean);
begin
   if FStartMini <> b then
   begin
     FStartMini:= b;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;

procedure TConfig.SetMailClientMini (b: Boolean);
begin
   if FMailClientMini <> b then
   begin
     FMailClientMini:= b;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;

procedure TConfig.SetRestNewMsg (b: Boolean);
begin
   if FRestNewMsg <> b then
   begin
     FRestNewMsg:= b;
     if Assigned(FOnChange) then FOnChange(Self);
   end;
end;

procedure TConfig.SetSaveLogs (b: boolean);
begin
  if FSaveLogs <> b then
  begin
    FSaveLogs:= b;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetStartupCheck (b: boolean);
begin
  if FStartupCheck <> b then
  begin
    FStartupCheck:= b;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetSmallBtns (b: boolean);
begin
  if FSmallBtns <> b then
  begin
    FSmallBtns:= b;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetSoundFile (s: string);
begin
  if FSoundFile <> s then
  begin
    FSoundFile:= s;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetLangStr (s: string);
begin
  if FLangStr <> s then
  begin
    FLangStr:= s;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetMailClient (s: string);
begin
  if FMailClient <> s then
  begin
    FMailClient:= s;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetMailClientName (s: string);
begin
  if FMailClientName <> s then
  begin
    FMailClientName:= s;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetMailClientIsUrl (b: boolean);
begin
  if FMailClientIsUrl <> b then
  begin
    FMailClientIsUrl:= b;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;

procedure TConfig.SetVersion(s:string);
begin
  if FVersion <> s then
  begin
    FVersion:= s;
    if Assigned(FOnChange) then FOnChange(Self);
  end;
end;



function TConfig.SaveXMLnode(iNode: TDOMNode): Boolean;
begin
  Try
    TDOMElement(iNode).SetAttribute ('version', FVersion);
    TDOMElement(iNode).SetAttribute ('savsizepos', BoolToString(FSavSizePos));
    TDOMElement(iNode).SetAttribute ('wstate', FWState);
    TDOMElement(iNode).SetAttribute ('lastupdchk', DateTimeToString(FLastUpdChk));
    TDOMElement(iNode).SetAttribute ('nochknewver', BoolToString(FNoChkNewVer));
    TDOMElement(iNode).SetAttribute ('startup', BoolToString(FStartup));
    TDOMElement(iNode).SetAttribute ('startmini',BoolToString(FStartMini));
    TDOMElement(iNode).SetAttribute ('mailclientmini',BoolToString(FMailClientMini));
    TDOMElement(iNode).SetAttribute ('restnewmsg',BoolToString(FRestNewMsg));
    TDOMElement(iNode).SetAttribute ('savelogs',BoolToString(FSaveLogs));
    TDOMElement(iNode).SetAttribute ('startupcheck',BoolToString(FStartupCheck));
    TDOMElement(iNode).SetAttribute ('smallbtns',BoolToString(FSmallBtns));
    TDOMElement(iNode).SetAttribute ('soundfile', FSoundFile);
    TDOMElement(iNode).SetAttribute ('langstr', FLangStr);
    TDOMElement(iNode).SetAttribute ('mailclient', FMailClient);
    TDOMElement(iNode).SetAttribute ('mailclientname', FMailClientName);
    TDOMElement(iNode).SetAttribute ('mailclientisurl', BoolToString(FMailClientIsUrl));
    Result:= True;
  except
    result:= False;
  end;
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
  i: integer;
  UpCaseAttrib: string;
begin
  Result := false;
  if (iNode = nil) or (iNode.Attributes = nil) then exit;
  // Browse settings attributes
  for i:= 0 to iNode.Attributes.Length-1 do
  try
    UpCaseAttrib:=UpperCase(iNode.Attributes.Item[i].NodeName);
    if UpCaseAttrib='VERSION' then FVersion:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='SAVSIZEPOS' then FSavSizePos:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='WSTATE' then  FWState:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='LASTUPDCHK' then FLastUpdChk:= StringToDateTime(iNode.Attributes.Item[i].NodeValue,'dd/mm/yyyy hh:nn:ss');
    if UpCaseAttrib='NOCHKNEWVER' then FNoChkNewVer:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='STARTUP' then FStartup:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='STARTMINI' then FStartMini:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='MAILCLIENTMINI' then FMailClientMini:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='RESTNEWMSG' then FRestNewMsg:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='SAVELOGS' then FSaveLogs:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='STARTUPCHECK' then FStartupCheck:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='SMALLBTNS' then FSmallBtns:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    if UpCaseAttrib='SOUNDFILE' then FSoundFile:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='LANGSTR' then FLangStr:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='MAILCLIENT' then FMailClient:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='MAILCLIENTNAME' then FMailClientName:= iNode.Attributes.Item[i].NodeValue;
    if UpCaseAttrib='MAILCLIENTISURL' then FMailClientIsUrl:= StringToBool(iNode.Attributes.Item[i].NodeValue);
    result:= true;
  except
    Result:= False;
  end;
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
end;

procedure TFSettings.FormActivate(Sender: TObject);
begin
  if first then
  begin
    GetDefaultMailCllient;
    first:= false;
  end;
end;

procedure TFSettings.FormDestroy(Sender: TObject);
begin
  if assigned (Settings) then Settings.free;
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

// Retrieve default mail client and all available mail clients

function TFSettings.GetDefaultMailCllient: string;
var
  Reg: Tregistry;
  MailKey, DefMailName, DefMailSubkey: string;
  MailSubkey: string;
  A:TStringArray;
  HomeDir: string;
  AppListArr : array of string;
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
  FixedCount:= 2;
  SetLength(MailClients, FixedCount);
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
      SetLength(MailClients, FixedCount);
      MailClients[FixedCount-1]:= Default(TmailClient);
      MailClients[FixedCount-1].Name:=Win10Mail;
      MailClients[FixedCount-1].Command:=Win10MailCmd;
      MailClients[FixedCount-1].Url:=false;
    end;
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

