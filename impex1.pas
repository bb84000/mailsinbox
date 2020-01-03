unit impex1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, accounts1, registry;

type

  { TFImpex }

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
    procedure BtnOKClick(Sender: TObject);
    procedure CBAccTypeChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BtnAccFileClick(Sender: TObject);
  private
    Reg: Tregistry;
    IsOutlook: boolean;
    procedure ImportOutlook ();
    function BinToWideStr(a: array of word): string;
    function RegGetBinaryString(regkey: Tregistry; attr: string): string;
  public
    ImpAccounts: TAccountsList;
    MailattAccName, OutlAccName: string;

  end;

var
  FImpex: TFImpex;

  const
    OutlRegKey='Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676';
    CLSID_OlkMail='{ed475418-b0d6-11d2-8c3b-00104b2a6676}';

implementation

{$R *.lfm}

{ TFImpex }

procedure TFImpex.BtnAccFileClick(Sender: TObject);
var
  i: integer;
begin
   if ODImpex.Execute then
     if Fileexists(ODIMpex.FileName) then
     begin
       EXMLAcc.text:= ODIMpex.FileName;
       Case CBAccType.ItemIndex of
         0: begin    // Old mailattente accounts
            if ImpAccounts.ImportOldXML(ODIMpex.FileName) then
            for i:= 0 to ImpAccounts.count-1 do
              LBImpex.Items.Add(ImpAccounts.GetItem(i).Name);
            end;
       end {case};
     end;
end;

procedure TFImpex.FormActivate(Sender: TObject);
begin
  CBAccType.Clear;
  CBAccType.Items.Add(MailattAccName);
  if IsOutlook then CBAccType.Items.add(OutlAccName);
  CBAccType.ItemIndex:= 0;
end;

procedure TFImpex.BtnOKClick(Sender: TObject);

begin

end;

procedure TFImpex.CBAccTypeChange(Sender: TObject);
begin
  EXMLAcc.text:='';
  LBImpex.Items.Clear;
  ImpAccounts.Reset;
  Case CBAccType.ItemIndex of
   0: begin           // Old mailattente accounts
        BtnAccFile.enabled:= true;
      end;
   1: begin           // Outlook 2007-2013 accounts only windows and outlook found
        BtnAccFile.enabled:= false;
        ImportOutlook();
      end;
 end;

end;

procedure TFImpex.FormCreate(Sender: TObject);
begin
  ImpAccounts:= TAccountsList.Create('impex');
  MailattAccName:= 'MailAttente accounts';
  OutlAccName:= 'Outlook 2007-2013 accounts';
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.KeyExists(OutlRegKey) then
  IsOutlook:= true;
end;

procedure TFImpex.FormDestroy(Sender: TObject);
begin
  if Assigned(ImpAccounts) then ImpAccounts.free;
  if Assigned(Reg) then Reg.free;
end;


function TFImpex.BinToWideStr(a: array of word): string;
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
    result:= BinToWideStr(AK);
  end;
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
    for i:= 0 to ImpAccounts.count-1 do
              LBImpex.Items.Add(ImpAccounts.GetItem(i).Name);
  end;
end;

end.

