{*******************************************************************************}
{ Impex1 unit - Import Outlook, Thunderbird and Mailattente mail accounts       }
{ for MailsInBox application                                                    }
{ bb - sdtp - january 2020                                                      }
{*******************************************************************************}

unit impex1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, Grids, accounts1, registry, lazbbinifiles, lazbbutils;

type
  // Main Impex form
  TFImpex = class(TForm)
    BtnCancel: TBitBtn;
    BtnOK: TBitBtn;
    CBAccType: TComboBox;
    EXMLAcc: TEdit;
    LAccTyp: TLabel;
    LBImpex: TListBox;
    LFileName: TLabel;
    ODImpex: TOpenDialog;
    Panel2: TPanel;
    BtnAccFile: TSpeedButton;
    SGImpex: TStringGrid;
    procedure BtnOKClick(Sender: TObject);
    procedure CBAccTypeChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BtnAccFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LBImpexSelectionChange(Sender: TObject; User: boolean);
  private
    Reg: Tregistry;
    IsOutlook: boolean;
    IsMailAttente: boolean;
    IsTBird: boolean;
    MailAttentePath: string;
    TBirdPath, TbirdProfilePath: string;
    procedure ImportOutlook;
    procedure ImportMailAttente(filename: string);
    procedure ImportTBird(filename: string);
    function BinToWideStr(a: array of word): widestring;
    function RegGetBinaryString(regkey: Tregistry; attr: string): string;
    function GetTbirdProfilePath(tbpath: string): string;
  public
    spassNotAvail: string;
    ImpAccounts: TAccountsList;
    MailattAccName, OutlAccName, TBirdAccName: string;
    xmlFilter, jsFilter: string;
    sBtnAccFileHint: string;
  end;

var
  FImpex: TFImpex;

  const
    OutlRegKey='Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676';
    CLSID_OlkMail='{ed475418-b0d6-11d2-8c3b-00104b2a6676}';

implementation

{$R *.lfm}

uses mailsinbox1;

{ TFImpex }

procedure TFImpex.FormCreate(Sender: TObject);

begin
  ImpAccounts:= TAccountsList.Create('impex');
  MailattAccName:= 'MailAttente accounts';
  OutlAccName:= 'Outlook 2007-2013 accounts';
  // Mailattente accounts file
  MailAttentePath:= FMailsInBox.UserAppsDataPath+PathDelim+'mailattente'+PathDelim+'accounts.xml';
  if FileExists(MailAttentePath) then IsMailAttente:= true;
  // Outlook accounts registry key
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.KeyExists(OutlRegKey) then IsOutlook:= true;
  // Thunderbird profiles path
  TBirdPath:= FMailsInBox.UserAppsDataPath+PathDelim+'thunderbird'+PathDelim;
  TbirdProfilePath:=GetTbirdProfilePath(TBirdPath);
  if DirectoryExists(TBirdPath) then
  begin
    IsTBird:= true;
    TbirdProfilePath:=GetTbirdProfilePath(TBirdPath);
  end;
end;

procedure TFImpex.FormActivate(Sender: TObject);
begin
  CBAccType.Clear;
  CBAccType.Items.Add(MailattAccName);
  if IsOutlook then CBAccType.Items.add(OutlAccName);
  if IsTBird then CBAccType.Items.add('Thunderbird');
  CBAccType.ItemIndex:= 0;
  CBAccTypeChange(Sender);
end;

// Get Thunderbird current profile path

function TFImpex.GetTbirdProfilePath(tbpath: string): string;
var
  ProfileIni: TBbInifile;
  Sections: TStringList;
  ProfilesCount, CurrentProfile: integer;
  i: integer;
  ssect, sdefpath, spath: string;
  ProfileVersion: integer;
  Paths: array of string;
  CurProfilePath: string;
begin
  ProfileIni:= TBbInifile.Create(tbpath+'profiles.ini');
  // Enumerate sections
  Sections:= TStringList.Create;
  ProfileIni.ReadSections(Sections);
  ProfilesCount:= 0;
  CurrentProfile:= 0;
  if Sections.Count > 0 then
  begin
    // First, search Install... section if exists (new profiles version)
    for i:= 0 to Sections.Count-1 do
    begin
      ssect:= Sections[i];
      if UpperCase(copy(ssect, 0, 7))='INSTALL' then
      begin
        sdefpath:= ProfileIni.ReadString (ssect, 'Default', '');
        ProfileVersion:= 2;
        Break;
      end else
      begin
        sdefpath:= '';
        ProfileVersion:= 1;
      end;
    end;
    // Now parse profiles
   for i:= 0 to Sections.Count-1 do
    begin
      ssect:= Sections[i];
      if copy(ssect, 0, 7)='Profile' then
      begin
        Inc(ProfilesCount);
        SetLength(Paths, ProfilesCount);
        spath:= IsAnsi2Utf8(ProfileIni.ReadString (ssect, 'Path', ''));
        // New profiles.ini version
        if (ProfileVersion=2) and (sdefpath=spath) then CurrentProfile:= ProfilesCount-1 ;
        spath:= StringReplace(spath, '/', PathDelim, [rfReplaceAll]);
        if Boolean(ProfileIni.ReadInteger(ssect, 'IsRelative', 0)) then
        begin
          CurProfilePath:= TbirdPath+IsAnsi2Utf8(spath+PathDelim);
        end else
        begin
           CurProfilePath:= spath+PathDelim;
        end;
         Paths[ProfilesCount-1]:= CurProfilePath;
        // Old profiles.ini version
        if (ProfileVersion=1) and (Boolean(ProfileIni.ReadInteger(ssect, 'Default', 0))) then CurrentProfile:= ProfilesCount-1 ;
      end;
    end;
  end;
  result:= Paths[CurrentProfile];
  if assigned(Sections) then Sections.Free;
  if Assigned(ProfileIni) then ProfileIni.free;
end;

// Click on file open button

procedure TFImpex.BtnAccFileClick(Sender: TObject);
var
  i: integer;
begin
   if ODImpex.Execute then
     if Fileexists(ODIMpex.FileName) then
     begin
       EXMLAcc.text:= ODIMpex.FileName;
       EXMLAcc.Hint:= ODIMpex.FileName;
       Case CBAccType.ItemIndex of
         0: begin    // Old mailattente accounts
              LBImpex.Items.Clear;
              ImpAccounts.Reset;
              ImportMailAttente(ODIMpex.FileName);
              for i:= 0 to ImpAccounts.count-1 do LBImpex.Items.Add(ImpAccounts.GetItem(i).Name);
            end;
         2: begin // external Thunderbird prefs.js file
              LBImpex.Items.Clear;
              ImpAccounts.Reset;
              ImportTBird(ODIMpex.FileName);
              for i:= 0 to ImpAccounts.count-1 do LBImpex.Items.Add(ImpAccounts.GetItem(i).Name);
            end;
       end {case};
     end;
end;

procedure TFImpex.FormShow(Sender: TObject);
begin

end;

procedure TFImpex.LBImpexSelectionChange(Sender: TObject; User: boolean);
var
  CurAcc: TAccount;
begin
  if LBImpex.ItemIndex >= 0 then
  Begin
   CurAcc:=  ImpAccounts.GetItem(LBImpex.ItemIndex);
    SGImpex.Cells[1,1]:= CurAcc.Name;
    SGImpex.Cells[1,2]:= CurAcc.Server;
    SGImpex.Cells[1,3]:= IntToStr(CurAcc.Port);
    SGImpex.Cells[1,4]:=ImpAccounts.ProtocolToString(CurAcc.Protocol);
    SGImpex.Cells[1,5]:= CurAcc.UserName;
    if length(Curacc.Password)=0
    then SGImpex.Cells[1,6]:= spassNotAvail
    else SGImpex.Cells[1,6]:= '**********' ;
    SGImpex.Cells[1,7]:= CurAcc.Email;
    SGImpex.Cells[1,8]:= CurAcc.ReplyEmail;
  end;
end;



procedure TFImpex.BtnOKClick(Sender: TObject);
begin

end;

procedure TFImpex.CBAccTypeChange(Sender: TObject);
var
  i: integer;
begin
  EXMLAcc.text:='';
  EXMLAcc.Hint:='';

  LBImpex.Items.Clear;
  ImpAccounts.Reset;
  Case CBAccType.ItemIndex of
   0: begin           // Old mailattente accounts
        ODImpex.Filter:= xmlFilter;
        ODImpex.FilterIndex:=1;
        BtnAccFile.Hint:= Format(sBtnAccFileHint,['Mailattente']);
        BtnAccFile.enabled:= true;
        EXMLAcc.text:= MailAttentePath;
        EXMLAcc.Hint:= MailAttentePath;
        ImportMailAttente(MailAttentePath);
      end;
   1: begin           // Outlook 2007-2013 accounts only windows and outlook found
        BtnAccFile.enabled:= false;
        If IsOutlook then
        begin
          ImportOutlook;
        end;
      end;
   2: begin          // Thunderbird accounts
        BtnAccFile.Hint:= Format(sBtnAccFileHint,['Thunderbird']);
        ODImpex.Filter:= jsFilter;
        ODImpex.FilterIndex:=1;
        BtnAccFile.enabled:= true;
        EXMLAcc.text:= TbirdProfilePath+'prefs.js';
        EXMLAcc.Hint:= TbirdProfilePath+'prefs.js';
        ImportTBird(TbirdProfilePath+'prefs.js');
      end;
  end;
  for i:= 0 to ImpAccounts.count-1 do LBImpex.Items.Add(ImpAccounts.GetItem(i).Name);
  LBImpex.ItemIndex:= 0;
  LBImpex.OnSelectionChange(Self, true);
end;



procedure TFImpex.FormDestroy(Sender: TObject);
begin
  if Assigned(ImpAccounts) then ImpAccounts.free;
  if Assigned(Reg) then Reg.free;
end;


function TFImpex.BinToWideStr(a: array of word): widestring;
var
  i: integer;
begin
  result:= '';
  for i:=0 to length(a)-1 do
  begin
    if a[i]>0 then result:= result+widechar(a[i]);
  end;
end;

// retrieve unicode string from registry when it is stored as binary

function TFImpex.RegGetBinaryString(regkey: Tregistry; attr: string): string;
var
  AK: Array of word;
  l: longint;
begin
  result:= '';
  if regkey.ValueExists(attr) then
  begin
    l:= regkey.GetDataSize(attr);
    setlength(AK, l div 2);
    regkey.ReadBinaryData(attr, Pointer(AK)^, l);
    result:= UTF8Encode(BinToWideStr(AK));
  end;
end;

procedure TFImpex.ImportMailAttente(filename: string);
begin
  ImpAccounts.ImportOldXML(FileName);
end;

procedure TFImpex.ImportOutlook;
var
  i, l: longint;
  B: Array of longint;
  Account: TAccount;
begin
  if IsOutlook then
  begin
    Account:= Default(TAccount);
    // open outlkook accounts key
    Reg.OpenKeyReadOnly(OutlRegKey);
    l:= Reg.GetDataSize(CLSID_OlkMail);
    Setlength(B, l div 4);
    // Get subkeys with accounts and populate array of subkeys
    reg.ReadBinaryData(CLSID_OlkMail, Pointer(B)^, l);
    reg.CloseKey;
    for i:=0 to length(B)-1 do
      begin
        Account.SSL:= 0;
        Account.SecureAuth:= false;
        Reg.OpenKeyReadOnly(OutlRegKey+'\'+IntToHex(B[i], 8));
        Account.Name:= RegGetBinaryString(reg, 'Display Name');
        Account.Email:= RegGetBinaryString(reg, 'Email');
        if reg.ValueExists('POP3 Server') then
        begin
          Account.Server:= RegGetBinaryString(reg, 'POP3 Server');
          Account.UserName:= RegGetBinaryString(reg, 'POP3 User');
          Account.Protocol:=ptcPOP3;
          if Reg.ValueExists('POP3 Port')
          then Account.Port:= reg.ReadInteger('POP3 Port')
          else Account.Port:= 110;       //Default value
          if Reg.ValueExists('POP3 Use SSL') then
                 Account.SSL:= reg.ReadInteger('POP3 Use SSL');
          if Reg.ValueExists('POP3 Use SPA') then
                 Account.SecureAuth:= Boolean(reg.ReadInteger('POP3 Use SPA'));
        end;
        if reg.ValueExists('IMAP Server') then
        begin
          Account.Server:= RegGetBinaryString(reg, 'IMAP Server');
          Account.UserName:= RegGetBinaryString(reg, 'IMAP User');
          Account.Protocol:=ptcIMAP;
          if Reg.ValueExists('IMAP Port')
          then Account.Port:= reg.ReadInteger('IMAP Port')
          else Account.Port:= 143;       //Default value
          if Reg.ValueExists('IMAP Use SSL') then
                 Account.SSL:= reg.ReadInteger('IMAP Use SSL');
          if Reg.ValueExists('IMAP Use SPA') then
                 Account.SecureAuth:= Boolean(reg.ReadInteger('IMAP Use SPA'));
        end;
        reg.CloseKey;
        ImpAccounts.AddAccount(Account);
      end;

  end;
end;

procedure TFImpex.ImportTBird(filename: string);
var
  CurProfilePath: string;
  sl: TStringList;
  slIds, slServ: TStringList;
  i: integer;
  s: string;
  A, B: TStringArray;
  AccNumber, PrevAccNumber, Acount: integer;
  CurAcc: TAccount;
  ndx: integer;
begin
  //Accounts are in prefs.js;
  sl:= TStringList.Create;
  sl.LoadFromFile(filename);
  sl.Sorted:= true;
  sl.sort;
  slIds:= TStringList.Create;
  slServ:=TStringList.Create ;
  // retrieve mail accounts
  PrevAccNumber:= 0;
  ACount:= 0;
  CurAcc:= Default(TAccount);
  for i:=0 to sl.Count-1 do
  begin
    s:= sl.Strings[i];
    // Search accounts identities
    if pos('user_pref("mail.account.ac', s)>0 then
    begin
      A:= s.split('"');
      B:= A[1].Split('.');
      AccNumber:= StringToInt(Copy(B[2], 8, 2));
      if B[3]= 'identities' then
      begin
        CurAcc.Name:=A[3];
        ImpAccounts.AddAccount(CurAcc);
        Inc(Acount);
      end;

      // Search server values
      if B[3]= 'server' then
      begin
        if AccNumber= PrevAccNumber then ImpAccounts.ModifyField(Acount-1, 'Server', A[3]);
      end;
      PrevAccNumber:= AccNumber;
    end;
    // Then create stringlisdt for identities and servers data
    if pos('user_pref("mail.id', s)>0 then slIds.Add(s);
    if pos('user_pref("mail.se', s)>0 then slServ.Add(s);
  end;
  // Now sort new stringlists
  slIds.Sorted:= true;
  slIds.Sort;
  slServ.Sorted:= true;
  slServ.Sort;
  LBImpex.Clear;
  For i:= 0 to ImpAccounts.count- 1 do
  begin
    // populate identity fields
    ndx:=-1;
    slIds.Find('user_pref("mail.identity.'+ImpAccounts.GetItem(i).Name+'.fullName"', ndx);
    if ndx >=0 then
    begin
      s:= slIds.Strings[ndx];
      A:= s.split('"');
      CurAcc.Name:= A[3];
    end;
    ndx:=-1;
    slIds.Find('user_pref("mail.identity.'+ImpAccounts.GetItem(i).Name+'.useremail"', ndx);
    if ndx>=0 then
    begin
      s:= slIds.Strings[ndx];
      A:= s.split('"');
      CurAcc.Email:= A[3];
    end;
    ndx:=-1;
    slIds.Find('user_pref("mail.identity.'+ImpAccounts.GetItem(i).Name+'.reply_to"', ndx);
    if ndx>=0 then
    begin
      s:= slIds.Strings[ndx];
      A:= s.split('"');
      CurAcc.ReplyEmail:= A[3];
    end;

    // Populate server fields
    ndx:=-1;
    slServ.Find('user_pref("mail.server.'+ImpAccounts.GetItem(i).Server+'.hostname', ndx);
    if ndx>=0 then
    begin
      s:= slServ.Strings[ndx];
      A:= s.split('"');
      CurAcc.Server := A[3];
    end;
    ndx:=-1;
    slServ.Find('user_pref("mail.server.'+ImpAccounts.GetItem(i).Server+'.type', ndx);
    if ndx>=0 then
    begin
      s:= slServ.Strings[ndx];
      A:= s.split('"');
      CurAcc.Protocol := ImpAccounts.StringToProtocol(A[3]);
    end;
    ndx:=-1;
    slServ.Find('user_pref("mail.server.'+ImpAccounts.GetItem(i).Server+'.userName', ndx);
    if ndx>=0 then
    begin
      s:= slServ.Strings[ndx];
      A:= s.split('"');
      CurAcc.UserName := A[3];
    end;
    ndx:=-1;
    slServ.Find('user_pref("mail.server.'+ImpAccounts.GetItem(i).Server+'.port', ndx);
    if ndx>=0 then
    begin
      s:= slServ.Strings[ndx];
      A:= s.split('"');
       B:= A[2].split(',)');
      CurAcc.Port := StringToInt(Trim(B[1]));
    end else Curacc.Port:= 0;
    if Curacc.Port= 0 then
    begin
      if CurAcc.Protocol=ptcPOP3 then CurAcc.Port:= 110;       //Default value
      if CurAcc.Protocol=ptcIMAP then CurAcc.Port:= 143;       //Default value
    end;
    ndx:=-1;
    slServ.Find('user_pref("mail.server.'+ImpAccounts.GetItem(i).Server+'.socketType', ndx);
    if ndx>=0 then
    begin
      s:= slServ.Strings[ndx];
      A:= s.split('"');
       B:= A[2].split(',)');
      CurAcc.SSL := StringToInt(Trim(B[1]));
    end else Curacc.SSL:= 0;
    if Curacc.SSL>0 then Curacc.SSL:= 1;  // implicit
    ImpAccounts.ModifyAccount(i, CurAcc);
  end;
  if Assigned(sl) then sl.free;
  if Assigned(slServ) then slServ.free;
  if Assigned(slIds) then slIds.free;

end;

end.

