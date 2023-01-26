{******************************************************************************}
{ accounts1 unit                                                               }
{ Form and types for accounts management                                       }
{ bb - sdtp - january 2023                                                     }
{ key for imported passwords 14235                                             }
{ key for saved passwords 14236                                                }
{******************************************************************************}
unit accounts1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, ComCtrls,
  StdCtrls, ColorBox, Spin, Buttons, laz2_DOM , laz2_XMLRead, laz2_XMLWrite,
  lazbbutils, lazbbinifiles ;

type
  TChampsCompare = (cdcNone, cdcName, cdcIndex, cdcUserName, cdcServer, cdcAccountUID, cdcMessageDate, cdcMessageNum, cdcMessageSize, cdcMessageUIDL,
                        cdcMessageFrom, cdcMessageTo, cdcMessageSubject, cdcMessageContentType);
  TSortDirections = (sdAscend, sdDescend);
  TProtocols = (ptcNone, ptcPOP3, ptcIMAP);
  TSaveType = (selection, all);

  // Define the classes in this Unit at the very start for clarity
  TFAccounts = Class;          // This is a forward class definition

  // Mail record

  PMail = ^TMail;
  TMail = Record
    AccountName : string;
    AccountIndex: Integer;
    AccountUID: Integer;
    MessageNum: Integer;
    MessageSize: Integer;
    MessageUIDL : string;
    MessageFrom : String;
    FromAddress : String;
    MessageTo : String;
    ToAddress: String;
    MessageSubject : String;
    MessageDate : TDateTime;
    MessageContentType: String;
    MessageNew: Boolean;
    MessageDisplayed: boolean;
    MessageToDelete: Boolean;
  end;

  // Account mails list
  TMailsList = class(TList)
  private
    FOnChange: TNotifyEvent;
    FSortType: TChampsCompare;
    FSortDirection: TSortDirections;
    procedure SetSortType (Value: TChampsCompare);
    procedure SetSortDirection(Value: TSortDirections);
  public
    Duplicates : TDuplicates;
    procedure Delete (const i : Integer);
    procedure DeleteMulti (j, k : Integer);
    procedure Reset;
    procedure AddMail(Mail : TMail);
    procedure ModifyMail (const i: integer; Mail : TMail);
    procedure ModifyField (const i: integer; field: string; value: variant);
    function GetItem(const i: Integer): TMail;
    function FindUIDL(value: string): integer;
    procedure DoSort;
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property SortType : TChampsCompare read FSortType write SetSortType default cdcNone;
    property SortDirection: TSortDirections read FSortDirection write SetSortDirection;

  end;

  // Account record

  PAccount = ^TAccount;
  TAccount= record
     Enabled: Boolean;
     Name: String;
     Index: Integer;
     Server: String;
     Protocol: TProtocols;
     SSL: integer;           //0:None, 1: Implicit, 2: Required, 3 Explicit
     Port: integer;
     UserName: String;
     Password: String;
     SecureAuth: boolean;
     Color: TColor;
     MailClient: String;
     SoundFile: string;
     Interval: Integer;
     Mails: TMailsList;
     UIDLToDel: Array of string;
     MsgToDel: TstringList;
     LastFire: TDateTime;
     NextFire: TDateTime;
     Email: string;
     ReplyEmail: string;
     Error: Boolean;
     ErrorStr: String;
     Checking: Boolean;
     Tag: Boolean;
     UID: Integer;
  end;

  // Accounts list

  TAccountsList = class(TList)
  private
    FOnChange: TNotifyEvent;
    FSortType: TChampsCompare;
    FAppName: string;
    procedure SetSortType (Value: TChampsCompare);

    function SaveItem(iNode: TDOMNode; sname, svalue: string): TDOMNode;
  public
    Duplicates : TDuplicates;
    constructor Create(AppName: String);
    destructor Destroy; override;
    procedure AddAccount(Account : TAccount);
    procedure Delete (const i : Integer);
    procedure Reset;
    procedure DeleteMulti (j, k : Integer);
    procedure ModifyAccount (const i: integer; Account : TAccount);
    procedure ModifyField (const i: integer; field: string; value: variant);
    function GetItem(const i: Integer): TAccount;
    function GetItemByUID(uid: Integer): TAccount;
    function LoadXMLnode(iNode: TDOMNode): Boolean;
    function LoadXMLfile(filename: string): Boolean;
    function SaveToXMLnode(iNode: TDOMNode; typ: TSaveType= all): Boolean;
    function SaveToXMLfile(filename: string; typ: TSaveType= all): Boolean;
    function ImportOldXML(filename: string): Boolean;
    function StringToProtocol(s: string): TProtocols;
    function ProtocolToString(protocol: TProtocols): string;
    procedure DoSort;
    function charsum(s: string): integer;
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    Property SortType : TChampsCompare read FSortType write SetSortType default cdcNone;
    property AppName: string read FAppName write FAppName;
 end;

  { TFAccounts }

  TFAccounts = class(TForm)
    BtnCancel: TBitBtn;
    BtnOK: TBitBtn;
    CBEnabledAcc: TCheckBox;
    CBSecureAuth: TCheckBox;
    CBColorAcc: TColorBox;
    CBSSL: TComboBox;
    CBProtocol: TComboBox;
    CBShowPass: TCheckBox;
    EEmail: TEdit;
    EInterval: TSpinEdit;
    EName: TEdit;
    EPassword: TEdit;
    EPort: TEdit;
    EReplyEmail: TEdit;
    EServer: TEdit;
    ESoundFile: TEdit;
    EUserName: TEdit;
    LAccName: TLabel;
    LColor: TLabel;
    LEmail: TLabel;
    LHost: TLabel;
    LInterval: TLabel;
    LMin: TLabel;
    LPassword: TLabel;
    LPort: TLabel;
    LProtocol: TLabel;
    LReply: TLabel;
    LSoundFile: TLabel;
    LSSL: TLabel;
    LUserName: TLabel;
    PnlButtons: TPanel;
    BtnSoundFile: TSpeedButton;
    BtnPlaySound: TSpeedButton;
    procedure BtnOKClick(Sender: TObject);
    procedure CBShowPassClick(Sender: TObject);
    procedure ESoundFileChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private

  public
    Accounts: TAccountsList;
    procedure Translate(LngFile: TBbIniFile);
  end;

var
  FAccounts: TFAccounts;
  ClesTri: array[0..Ord(High(TChampsCompare))] of TChampsCompare;

implementation

{$R *.lfm}


// global functions

function stringCompare(Item1, Item2: String): Integer;
begin
   result := Comparestr(Item1, Item2);
end;

function NumericCompare(Item1, Item2: Double): Integer;
begin
  if Item1 > Item2 then result := 1
  else
  if Item1 = Item2 then result := 0
  else result := -1;
end;

function CompareMultiA(Item1, Item2: Pointer): Integer;
var
  Entry1, Entry2: PAccount;
  R, J: integer;
  ResComp: array[TChampsCompare] of integer;
begin
  Entry1:= PAccount(Item1);
  Entry2:= PAccount(Item2);
  ResComp[cdcNone]  := 0;
  ResComp[cdcName]  := StringCompare(Entry1^.Name, Entry2^.Name);
  ResComp[cdcIndex] := NumericCompare(Entry1^.Index, Entry2^.Index);
  ResComp[cdcUserName] := StringCompare(Entry1^.Name, Entry2^.Name);
  ResComp[cdcServer] := StringCompare(Entry1^.Name, Entry2^.Name);
  R := 0;
  for J := 0 to Ord(High(TChampsCompare)) do
  begin
    if ResComp[ClesTri[J]] <> 0 then
     begin
       R := ResComp[ClesTri[J]];
       break;
     end;
  end;
  result :=  R;
end;

function CompareMultiM(Item1, Item2: Pointer): Integer;
var
  Entry1, Entry2: PMail;
  R, J: integer;
  ResComp: array[TChampsCompare] of integer;
begin
  Entry1:= PMail(Item1);
  Entry2:= PMail(Item2);
  ResComp[cdcNone]  := 0;
  ResComp[cdcName]  := StringCompare(Entry1^.AccountName, Entry2^.AccountName);
  ResComp[cdcIndex] := NumericCompare(Entry1^.AccountIndex, Entry2^.AccountIndex);
  ResComp[cdcAccountUID] := NumericCompare(Entry1^.AccountUID, Entry2^.AccountUID);
  ResComp[cdcMessageDate] := NumericCompare(Entry1^.MessageDate, Entry2^.MessageDate);
  ResComp[cdcMessageFrom] := StringCompare(Entry1^.MessageFrom, Entry2^.MessageFrom);
  ResComp[cdcMessageTo] := StringCompare(Entry1^.MessageTo, Entry2^.MessageTo);
  ResComp[cdcMessageSubject] := StringCompare(Entry1^.MessageSubject, Entry2^.MessageSubject);
  ResComp[cdcMessageSize] := NumericCompare(Entry1^.MessageSize, Entry2^.MessageSize);
  R := 0;
  for J := 0 to Ord(High(TChampsCompare)) do
  begin
    if ResComp[ClesTri[J]] <> 0 then
     begin
       R := ResComp[ClesTri[J]];
       break;
     end;
  end;
  result :=  R;
end;

function CompareMultiDesc(Item1, Item2: Pointer): Integer;
begin
  result:=  -CompareMultiM(Item1, Item2);
end;

{ TMailsList }

procedure TMailsList.SetSortDirection(Value: TSortDirections);
begin
  FSortDirection:= Value;
  if FSortDirection = sdAscend then sort(@CompareMultiM) else sort(@comparemultidesc);
end;

procedure TMailsList.SetSortType (Value: TChampsCompare);
begin
  FSortType:= Value;
  if Assigned(FOnChange) then FOnChange(Self);
  DoSort;
end;

procedure TMailsList.Delete(const i: Integer);
begin
inherited delete(i);
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TMailsList.DeleteMulti(j, k : Integer);
var
  i : Integer;
begin
  // On commence par le dernier, ajouter des sécurités
  For i:= k downto j do
  begin
     inherited delete(i);
  end;
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TMailsList.Reset;
var
 i: Integer;
begin
 for i := 0 to Count-1 do
  if Items[i] <> nil then Items[i]:= nil;
 Clear;
end;

procedure TMailsList.AddMail(Mail : TMail);
var
 K: PMail;
begin
  new(K);
  K^:= Mail;
  add(K);
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TMailsList.ModifyMail (const i: integer; Mail : TMail);
begin
  TMail(Items[i]^):= Mail;

  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TMailsList.ModifyField (const i: integer; field: string; value: variant);
begin
  field:= Uppercase(field);
  if field = 'ACCOUNTNAME' then TMail(Items[i]^).AccountName:= value;
  if field = 'ACCOUNTINDEX' then TMail(Items[i]^).AccountIndex:= value;
  if field = 'ACCOUNTUID' then TMail(Items[i]^).AccountUID:= value;
  if field = 'MESSAGENUM' then TMail(Items[i]^).MessageNum:= value;
  if field = 'MESSAGESIZE' then TMail(Items[i]^).MessageSize:= value;
  if field = 'MESSAGEUIDL' then TMail(Items[i]^).MessageUIDL:= value;
  if field = 'MESSAGEFROM' then TMail(Items[i]^).MessageFrom:= value;
  if field = 'FROMADDRESS' then TMail(Items[i]^).FromAddress:= value;
  if field = 'MESSAGETO' then TMail(Items[i]^).MessageTo:= value;
  if field = 'TOADDRESS' then TMail(Items[i]^).ToAddress:= value;
  if field = 'MESSAGESUBJECT' then TMail(Items[i]^).MessageSubject:= value;
  if field = 'MESSAGEDATE' then TMail(Items[i]^).Messagedate:= value;
  if field = 'MESSAGECONTENTTYPE' then TMail(Items[i]^).MessageContentType:= value;
  if field = 'MESSAGENEW' then TMail(Items[i]^).MessageNew:= value;
  if field = 'MESSAGEDISPLAYED' then TMail(Items[i]^).MessageDisplayed:= value;
  if field = 'MESSAGETODELETE' then TMail(Items[i]^).MessageToDelete:= value;
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

function TMailsList.GetItem(const i: Integer): TMail;
begin
 Result := TMail(Items[i]^);
end;



function TMailsList.FindUIDL(value: String): integer;
var
 i: integer;
begin
  result:=-1;
  value:= uppercase(value);
  for i:= 0 to count-1 do
    if value= uppercase(TMail(Items[i]^).MessageUIDL) then
    begin
      result:= i;
      break;
    end;
end;

procedure TMailsList.DoSort;
begin
  if FSortType <> cdcNone then
  begin
    ClesTri[1] := FSortType;
    //ClesTri[2] := cdcName;
    //ClesTri[3] := cdcDur;
    sort(@comparemultiM);
  end;
end;


{ TAccounts }

procedure TAccountsList.SetSortType (Value: TChampsCompare);
begin
  FSortType:= Value;
  if Assigned(FOnChange) then FOnChange(Self);
  DoSort;
end;


constructor TAccountsList.Create( AppName: String);
begin
  inherited Create;
  FAppName:= AppName;
end;

destructor  TAccountsList.Destroy;
begin
  Reset;
  inherited Destroy;
end;

// Efface tous les comptes

procedure TAccountsList.Reset;
var
 i: Integer;
begin
 for i := 0 to Count-1 do
 begin
   if assigned(TAccount(Items[i]^).Mails) then TAccount(Items[i]^).Mails:= nil;//.Free ;
   //if assigned(TAccount(Items[i]^).UIDLList) then TAccount(Items[i]^).UIDLList.Free;
   if Items[i] <> nil then Items[i]:= nil;
 end;
 Clear;
end;

function TAccountsList.charsum(s: string): integer;
var
 i: integer;
begin
  result:=0;
  for i:=1 to length(s) do result:=result+ord(s[i]);
end;

procedure TAccountsList.AddAccount(Account : TAccount);
var
 K: PAccount;
begin
  new(K);
  if Account.UID=0 then Account.UID:= charsum(Account.Name+Account.Password+Account.Server);
  K^:= Account;
  // we create the mails list if not already created
  if not assigned(K^.Mails) then K^.Mails:= TMailsList.Create;
  //if not assigned(K^.UIDLList) then K^.UIDLList:= TStringList.Create;

  add(K);
  DoSort;
  K:= nil;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TAccountsList.Delete(const i: Integer);
begin
  inherited delete(i);
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TAccountsList.DeleteMulti(j, k : Integer);
var
  i : Integer;
begin
  // On commence par le dernier, ajouter des sécurités
  For i:= k downto j do
  begin
     inherited delete(i);
  end;
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TAccountsList.ModifyAccount (const i: integer; Account : TAccount);
begin
  TAccount(Items[i]^):= Account;
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TAccountsList.ModifyField (const i: integer; field: string; value: variant);
begin
  field:= Uppercase(field);
  if field = 'ENABLED' then TAccount(Items[i]^).Enabled:= value;
  if field = 'NAME' then TAccount(Items[i]^).Name := value;
  if field = 'INDEX' then TAccount(Items[i]^).Index:= value;
  if field = 'SERVER' then TAccount(Items[i]^).Server:= value;
  if field = 'PROTOCOL' then TAccount(Items[i]^).Protocol := value;
  if field = 'SSL' then TAccount(Items[i]^).SSL:= value;
  if field = 'PORT' then TAccount(Items[i]^).Port:= value;
  if field = 'USERNAME' then TAccount(Items[i]^).UserName:= value;
  if field = 'PASSWORD' then TAccount(Items[i]^).Password := value;
  if field = 'SECUREAUTH' then TAccount(Items[i]^).SecureAuth:= value;
  if field = 'COLOR' then TAccount(Items[i]^).Color:= value;
  if field = 'MAILCLIENT' then TAccount(Items[i]^).MailClient:= value;
  if field = 'SOUNDFILE' then TAccount(Items[i]^).SoundFile:= value;
  if field = 'INTERVAL' then TAccount(Items[i]^).Interval:= value;
  if field = 'EMAIL' then TAccount(Items[i]^).Email:= value;
  if field = 'REPLYEMAIL' then TAccount(Items[i]^).ReplyEmail:= value;
  if field = 'LASTFIRE' then TAccount(Items[i]^).LastFire:= value;
  if field = 'NEXTFIRE' then TAccount(Items[i]^).NextFire:= value;
  //if field = 'ERROR' then TAccount(Items[i]^).Error:= value;
  //if field = 'ERRORSTR' then TAccount(Items[i]^).ErrorStr:= value;
  if field = 'CHECKING' then TAccount(Items[i]^).Checking:= value;
  if field = 'TAG' then TAccount(Items[i]^).Tag:= value;
  if field = 'UID' then TAccount(Items[i]^).UID:= value;
  DoSort;
  if Assigned(FOnChange) then FOnChange(Self);
end;

function TAccountsList.GetItem(const i: Integer): TAccount;
begin
  Result := TAccount(Items[i]^);
end;

function TAccountsList.GetItemByUID(uid: Integer): TAccount;
var
  i: integer;
 begin
  Result:= Default(TAccount);
  Result.Index:=-1;
  for i:= 0 to count-1 do
  try
    if GetItem(i).UID=uid then
    begin
      result:= GetItem(i);
      break;
    end;
  except
  end;
end;

procedure TAccountsList.DoSort;
begin
  if (FSortType<>cdcNone) then
  begin
    ClesTri[1] := FSortType;
    //ClesTri[2] := cdcName;
    //ClesTri[3] := cdcDur;
    sort(@comparemultiA);
  end;
end;

function TAccountsList.StringToProtocol(s: string):TProtocols;
begin
  result:= ptcNone;
  if UpperCase(s)='POP3' then result:= ptcPOP3;
  if UpperCase(s)='IMAP' then result:= ptcIMAP;
end;

function TAccountsList.ProtocolToString(protocol: TProtocols): string;
begin
  result:= 'none';
  if protocol=ptcPOP3 then result:= 'pop3';
  if protocol=ptcIMAP then result:= 'imap';
end;

function TAccountsList.LoadXMLNode(iNode: TDOMNode): Boolean;
var
  chNode: TDOMNode;
  subNode: TDOMNode;
  s: string;
  upNodeName: string;
  A: TAccount;
const
  key=14236;
begin
  SortType:= TChampsCompare(StringToInt(TDOMElement(iNode).GetAttribute('sort')));
  chNode := iNode.FirstChild;
  while (chNode <> nil) and (UpperCase(chnode.NodeName)='ACCOUNT')  do
  begin
    Try
      A:= Default(TAccount);
      subNode:= chNode.FirstChild;
      while subNode <> nil do
      try
        upNodeName:= UpperCase(subNode.NodeName);
        s:= subNode.TextContent;
        if upNodeName = 'ENABLED' then A.Enabled:= StringToBool(s);
        if upNodeName = 'NAME' then A.Name := s;
        if upNodeName = 'INDEX' then A.Index:= StringToInt(s);
        if upNodeName = 'SERVER' then A.Server:= s;
        if upNodeName = 'PROTOCOL' then A.Protocol := StringToProtocol(s);
        if upNodeName = 'SSL' then A.SSL:= StringToInt(s);
        if upNodeName = 'PORT' then A.Port:= StringToInt(s);
        if upNodeName = 'USERNAME' then A.UserName:= s;
        if upNodeName = 'PASSWORD' then A.Password := StringDecrypt(s, key);
        if upNodeName = 'SECUREAUTH' then A.SecureAuth:= StringToBool(s);
        if upNodeName = 'COLOR' then A.Color:= StrToInt('$'+s);
        if upNodeName = 'MAILCLIENT' then A.MailClient:= s;
        if upNodeName = 'SOUNDFILE' then A.SoundFile:= s;
        if upNodeName = 'INTERVAL' then A.Interval:= StringToInt(s);
        if upNodeName = 'EMAIL' then A.Email:= s;
        if upNodeName = 'REPLYEMAIL' then A.ReplyEmail:= s;
        if upNodeName = 'LASTFIRE' then A.LastFire:= StringToTimeDate(s, 'dd/mm/yyyy hh:nn:ss');
        if upNodeName = 'NEXTFIRE' then A.NextFire:= StringToTimeDate(s, 'dd/mm/yyyy hh:nn:ss');
        if upNodeName = 'TAG' then A.Tag:= StringToBool(s);
        if upNodeName = 'UID' then A.UID:= StringToInt(s);
      finally
        subnode:= subnode.NextSibling;
      end;
      AddAccount(A);
    finally
      chNode := chNode.NextSibling;
    end;
  end;
  result:= true;
end;

function TAccountsList.LoadXMLFile(filename: string): Boolean;
var
  AccountsXML: TXMLDocument;
  RootNode,AccountsNode : TDOMNode;
begin
  result:= false;
  if not FileExists(filename) then
  begin
    SaveToXMLfile(filename);
  end;
  ReadXMLFile(AccountsXML, filename);
  RootNode := AccountsXML.DocumentElement;
  AccountsNode:= RootNode.FindNode('accounts');
  if AccountsNode= nil then exit;
  LoadXMLnode(AccountsNode);
  If assigned(AccountsNode) then AccountsNode.free;
  result:= true;
end;


function TAccountsList.ImportOldXML(filename: string): Boolean;
var
  AccountsXML: TXMLDocument;
  RootNode : TDOMNode;
  iNode:TDOMNode;
  i: integer;
  UpCaseAttrib: string;
  s: string;
  A: TAccount;
const
  key=14235;
begin
  result:= false;
  if not FileExists(filename) then exit;
  ReadXMLFile(AccountsXML, filename);
  RootNode := AccountsXML.DocumentElement;   // accounts are in root node "accounts"
  if RootNode.NodeName='accounts' then
  begin
    iNode := RootNode.FirstChild;
    i:= 0;
    while iNode <> nil do
      if iNode.Attributes<>nil then
      begin
        A:= Default(TAccount);
        for i:= 0 to iNode.Attributes.Length-1 do
        try
          UpCaseAttrib:=UpperCase(iNode.Attributes.Item[i].NodeName);
          s:= iNode.Attributes.Item[i].NodeValue;
          if UpCaseAttrib='ENABLED' then  A.Enabled:= StringToBool(s);
          if UpCaseAttrib='ACCOUNTNAME' then A.Name:= s;
          if UpCaseAttrib='ACCOUNTINDEX' then A.Index:= StringToInt(s);
          if UpCaseAttrib='HOST' then A.Server:= s;
          if UpCaseAttrib='PROTOCOL' then A.Protocol:= StringToProtocol(s);
          if UpCaseAttrib='SSLTYPE' then A.SSL:= StringToInt(s);
          if UpCaseAttrib='PORT' then A.Port:= StringToInt(s);
          if UpCaseAttrib='USERNAME' then A.UserName:= s;
          if UpCaseAttrib='PASSWORD' then A.Password := StringDecrypt(s, key);
          if UpCaseAttrib='AUTH' then A.SecureAuth:= StringToBool(s);
          if UpCaseAttrib='COLOR' then A.Color:= StrToInt('$'+s);
          if UpCaseAttrib='MAILCLIENT' then A.MailClient:= s;
          if UpCaseAttrib='SOUND' then A.SoundFile:= s;
          if UpCaseAttrib='INTERVAL' then A.Interval:= StringToInt(s);
          if UpCaseAttrib='EMAIL' then A.Email:= s;
          if UpCaseAttrib='REPLYEMAIL' then A.ReplyEmail:= s;
        except
        end;
        AddAccount(A);
        iNode:= iNode.NextSibling;
      end;
    result:= true;
  end;
end;


function TAccountsList.SaveItem(iNode: TDOMNode; sname, svalue: string): TDOMNode;
begin
  result:= iNode.OwnerDocument.CreateElement(sname);
  result.TextContent:= svalue;
end;

function TAccountsList.SaveToXMLnode(iNode: TDOMNode; typ: TSaveType= all): Boolean;
var
  i: Integer;
  ContNode: TDOMNode;
const
  key=14236;
begin
  Result:= True;
  If Count > 0 Then
   begin
     TDOMElement(iNode).SetAttribute('sort', IntToStr(Ord(SortType)));
     For i:= 0 to Count-1 do
     Try
       // Skip tagged contact if in selection typ
       if (typ=selection) and not(GetItem(i).Tag) then continue;
       // Reset tag to false when processed
       if GetItem(i).Tag then TAccount(Items[i]^).Tag:= false;
       ContNode := iNode.OwnerDocument.CreateElement('account');
       iNode.Appendchild(ContNode);
       ContNode.AppendChild(SaveItem(ContNode, 'enabled', BoolToString(GetItem(i).Enabled)));
       ContNode.AppendChild(SaveItem(ContNode, 'name', GetItem(i).Name));
       ContNode.AppendChild(SaveItem(ContNode, 'index', IntToStr(i)));
       ContNode.AppendChild(SaveItem(ContNode, 'server', GetItem(i).Server));
       ContNode.AppendChild(SaveItem(ContNode, 'protocol', ProtocolToString(GetItem(i).Protocol)));
       ContNode.AppendChild(SaveItem(ContNode, 'ssl', IntToStr(GetItem(i).ssl)));
       ContNode.AppendChild(SaveItem(ContNode, 'port', IntToStr(GetItem(i).Port)));
       ContNode.AppendChild(SaveItem(ContNode, 'username', GetItem(i).UserName));
       ContNode.AppendChild(SaveItem(ContNode, 'password', StringEncrypt(GetItem(i).Password, key)));
       ContNode.AppendChild(SaveItem(ContNode, 'secureauth', BoolToString(GetItem(i).SecureAuth)));
       ContNode.AppendChild(SaveItem(ContNode, 'color', IntToHex(GetItem(i).Color, 8)));
       ContNode.AppendChild(SaveItem(ContNode, 'mailclient', GetItem(i).MailClient));
       ContNode.AppendChild(SaveItem(ContNode, 'soundfile', GetItem(i).SoundFile));
       ContNode.AppendChild(SaveItem(ContNode, 'interval', IntToStr(GetItem(i).Interval)));
       ContNode.AppendChild(SaveItem(ContNode, 'email', GetItem(i).Email));
       ContNode.AppendChild(SaveItem(ContNode, 'replyemail', GetItem(i).ReplyEmail));
       ContNode.AppendChild(SaveItem(ContNode, 'lastfire', TimeDateToString(GetItem(i).LastFire, 'dd/mm/yyyy hh:nn:ss')));
       ContNode.AppendChild(SaveItem(ContNode, 'nextfire', TimeDateToString(GetItem(i).NextFire, 'dd/mm/yyyy hh:nn:ss')));
       ContNode.AppendChild(SaveItem(ContNode, 'tag', BoolToString(GetItem(i).Tag)));
       ContNode.AppendChild(SaveItem(ContNode, 'uid', IntToStr(GetItem(i).UID)));
     except
       Result:= False;
     end;
   end;
end;

function TAccountsList.SaveToXMLfile(filename: string; typ: TSaveType= all): Boolean;
var
  AccountsXML: TXMLDocument;
  RootNode, AccountsNode :TDOMNode;
begin
  result:= false;
  if FileExists(filename)then
  begin
    ReadXMLFile(AccountsXML, filename);
    RootNode := AccountsXML.DocumentElement;
  end else
  begin
    AccountsXML := TXMLDocument.Create;
    RootNode := AccountsXML.CreateElement(lowercase(FAppName));
    AccountsXML.Appendchild(RootNode);
  end;
  AccountsNode:= RootNode.FindNode('accounts');
  if AccountsNode <> nil then RootNode.RemoveChild(AccountsNode);
  AccountsNode:= AccountsXML.CreateElement('accounts');
  //if Count > 0 then                           // Prevents first creation of file !!!
  begin
    SaveToXMLnode(AccountsNode);
    RootNode.Appendchild(AccountsNode);
    writeXMLFile(AccountsXML, filename);
    result:= true;
  end;
  if assigned(AccountsXML) then AccountsXML.free;;
end;



{ TFAccounts }

procedure TFAccounts.FormCreate(Sender: TObject);
begin
  Accounts:= TAccountsList.Create('progname');
end;

procedure TFAccounts.CBShowPassClick(Sender: TObject);
begin
  if CBShowPass.checked then EPassword.PasswordChar:=#0
  else EPassword.PasswordChar:='*';
end;

procedure TFAccounts.BtnOKClick(Sender: TObject);
begin

end;

procedure TFAccounts.ESoundFileChange(Sender: TObject);
begin
  BtnPlaySound.Enabled:= not (length(ESoundFile.Text)=0);
end;

procedure TFAccounts.FormActivate(Sender: TObject);
begin
  // Center buttons in case of width change
  BtnOK.Left:= (PnlButtons.ClientWidth-BtnOK.width*2-20) div 2;
  BtnCancel.Left:= BtnOK.Left+BtnOK.Width+20;
end;

procedure TFAccounts.FormDestroy(Sender: TObject);
begin
  if assigned (accounts) then accounts.Destroy;
end;

procedure TFAccounts.Translate(LngFile: TBbIniFile);
var
  DefaultCaption: String;
begin
  if assigned (Lngfile) then
  with LngFile do
  begin
    BtnOK.Caption:= ReadString('Common', 'OKBtn', BtnOK.Caption);
    BtnCancel.Caption:= ReadString('Common', 'CancelBtn', BtnCancel.Caption);
    DefaultCaption:= ReadString('Common', 'DefaultCaption', '...');
    Caption:=Format(ReadString('FSettings','Caption','Préférences de %s'), [DefaultCaption]);
    LAccName.Caption:=ReadString('FAccounts', 'LAccName.Caption', LAccName.Caption);
    LHost.Caption:=ReadString('FAccounts', 'LHost.Caption',LHost.Caption);
    LProtocol.Caption:=ReadString('FAccounts', 'LProtocol.Caption', LProtocol.Caption);
    CBProtocol.Items[0]:=ReadString('FAccounts', 'CBProtocol.Items_0', CBProtocol.Items[0]);
    LUserName.Caption:=ReadString('FAccounts', 'LUserName.Caption', LUserName.Caption);
    LPassword.Caption:=ReadString('FAccounts', 'LPassword.Caption', LPassword.Caption);
    LEmail.Caption:=ReadString('FAccounts', 'LEmail.Caption', LEmail.Caption);
    LColor.Caption:=ReadString('FAccounts', 'LColor.Caption', LColor.Caption);
    LSoundFile.Caption:=ReadString('FAccounts', 'LSoundFile.Caption', LSoundFile.Caption);
    LSSL.Caption:=ReadString('FAccounts', 'LSSL.Caption', LSSL.Caption);
    CBSSL.Items[0]:=ReadString('FAccounts', 'CBSSL.Items_0', CBSSL.Items[0]);
    CBSSL.Items[1]:=ReadString('FAccounts', 'CBSSL.Items_1', CBSSL.Items[1]);
    CBSSL.Items[2]:=ReadString('FAccounts', 'CBSSL.Items_2', CBSSL.Items[2]);
    LPort.Caption:=ReadString('FAccounts', 'LPort.Caption', LPort.Caption);
    CBSecureAuth.Caption:=ReadString('FAccounts', 'CBSecureAuth.Caption', CBSecureAuth.Caption);
    LInterval.Caption:=ReadString('FAccounts', 'LInterval.Caption', LInterval.Caption);
    LMin.Caption:= ReadString('FAccounts', 'LMin.Caption', LMin.Caption);;
    CBShowPass.Caption:=ReadString('FAccounts', 'CBShowPass.Caption', CBShowPass.Caption);
    LReply.Caption:=ReadString('FAccounts', 'LReply.Caption', LReply.Caption);
    CBEnabledAcc.Caption:=ReadString('FAccounts', 'CBEnabledAcc.Caption', CBEnabledAcc.Caption);
    //BtnMailClient.Hint:=ReadString('FAccounts', 'BtnMailClient.Hint', BtnMailClient.Hint);
    BtnPlaySound.Hint:=ReadString('FAccounts', 'BtnPlaySound.Hint', BtnPlaySound.Hint);
    BtnSoundFile.Hint:=ReadString('FAccounts', 'BtnSoundFile.Hint', BtnSoundFile.Hint);
  end;
end;


end.

