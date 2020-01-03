unit mailsinbox1;

{$mode objfpc}{$H+}

interface

uses
  {$IFDEF WINDOWS}
  Win32Proc,
  {$ENDIF} Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Grids, ComCtrls, Buttons, Menus, IdPOP3, IdSSLOpenSSL,
  IdExplicitTLSClientServerBase, IdMessage, accounts1, lazbbutils,
  lazbbinifiles, lazbbosversion, LazUTF8, settings1, lazbbautostart, lazbbabout,
  Registry, Impex1, mailclients1, uxtheme, Types, IdComponent;

type
  TSaveMode = (None, Setting, All);
  TBtnSize=(Small, Large);
  { TFMailsInBox }

  TFMailsInBox = class(TForm)
    BtnAbout: TSpeedButton;
    BtnAddAcc: TSpeedButton;
    BtnClose: TSpeedButton;
    BtnDelete: TSpeedButton;
    BtnEditAcc: TSpeedButton;
    BtnGetAccMail: TSpeedButton;
    BtnGetAllMail: TSpeedButton;
    BtnLaunchClient: TSpeedButton;
    BtnLog: TSpeedButton;
    BtnQuit: TSpeedButton;
    BtnSettings: TSpeedButton;
    GBInfos: TGroupBox;
    IdPOP3_1: TIdPOP3;
    ILMail: TImageList;
    ImgAccounts: TImageList;
    LStatus: TLabel;
    ListBox1: TListBox;
    LVAccounts: TListView;
    MInfos: TMemo;
    MnuMoveDown: TMenuItem;
    MnuMoveUp: TMenuItem;
    PnlMails: TPanel;
    PnlAccounts: TPanel;
    PnlStatus: TPanel;
    PnlToolbar: TPanel;
    BtnFirst: TSpeedButton;
    BtnPrev: TSpeedButton;
    BtnNext: TSpeedButton;
    BtnLast: TSpeedButton;
    BtnImport: TSpeedButton;
    MnuAccount: TPopupMenu;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    SGMails: TStringGrid;
    TimerTray: TTimer;
    TrayMail: TTrayIcon;
    procedure BtnAboutClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure BtnGetAccMailClick(Sender: TObject);
    procedure BtnGetAllMailClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure BtnLaunchClientClick(Sender: TObject);
    procedure BtnLogClick(Sender: TObject);
    procedure BtnNavClick(Sender: TObject);
    procedure BtnQuitClick(Sender: TObject);
    procedure BtnSettingsClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure ChangeBounds(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnEditAccClick(Sender: TObject);
    procedure BtnCloseClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure IdPOP3_1Connected(Sender: TObject);
    procedure IdPOP3_1Disconnected(Sender: TObject);
    procedure IdPOP3_1Status(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure LVAccountsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure LVAccountsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure MnuAccountPopup(Sender: TObject);
    procedure MnuMoveDownClick(Sender: TObject);
    procedure MnuMoveUpClick(Sender: TObject);
    procedure SGMailsBeforeSelection(Sender: TObject; aCol, aRow: Integer);
    procedure SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
      aRect: TRect; aState: TGridDrawState);
    procedure SGMailsEnter(Sender: TObject);
    procedure SGMailsExit(Sender: TObject);
    procedure SGMailsPrepareCanvas(sender: TObject; aCol, aRow: Integer;
      aState: TGridDrawState);
    procedure SGMailsSelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
    procedure SGMailsSelection(Sender: TObject; aCol, aRow: Integer);
    procedure TimerTrayTimer(Sender: TObject);
  private
    First: boolean;
    OS, OSTarget, CRLF: string;
    CompileDateTime: TDateTime;
    UserPath, UserAppsDataPath: string;
    MailsInBoxAppsData: string;
    ProgName: string;
    LangStr: string;
    LangFile: TBbIniFile;
    LangNums: TStringList;
    LangFound: boolean;
    SettingsChanged: boolean;
    AccountsChanged:Boolean;
    ConfigFile: string;
    traycount: integer;
    canCloseMsg: string;
    CanClose: boolean;
    AccountCaption, EmailCaption, LastCheckCaption, NextCheckCaption : string;
    BaseUpdateUrl, ChkVerURL, version: string;
    NoLongerChkUpdates, LastUpdateSearch, UpdateAvailable, UpdateAlertBox: string;
    OKBtn, YesBtn, NoBtn, CancelBtn: string;
    BtnLogHint, BtnGetAccMailHint, BtnDeleteHint, BtnEditAccHint: string;
    BmpArray: array [0..15] of TBitmap;
    DefMailClient: string;
    AccImportd, AccImportds: string;
    mailcolsiz: string;
    CheckingMail: boolean;
    ButtonStates: array of Boolean;
    SGHasFocus: boolean;
    DefCursor: TCursor;
    MsgFound, MsgsFound: string;
    procedure LoadCfgFile(filename: string);
    procedure SettingsOnChange(Sender: TObject);
    procedure SettingsOnStateChange(Sender: TObject);
    procedure AccountsOnChange(Sender: TObject);
    function SaveConfig(Typ: TSaveMode): boolean;
    procedure PopulateAccountsList;
    procedure ModLangue;
    procedure SetSmallBtns(small: boolean);
    function GetPop3Mail(index: integer): boolean;
    procedure PopulateMailsList(index: integer);
    procedure EnableControls(Enable: boolean);
    procedure UpdateInfos;
    procedure DrawTheIcon(Bmp: TBitmap; NewCount: integer; CircleColor: TColor);
  public
    OsInfo: TOSInfo;

  end;

var
  FMailsInBox: TFMailsInBox;

implementation

{$R *.lfm}


// TFMailsInBox : This is the main form of the program

procedure TFMailsInBox.FormCreate(Sender: TObject);
var
  s: string;
begin
  First := True;
  CompileDateTime:= StringToDateTime({$I %DATE%}+' '+{$I %TIME%}, 'yyyy/mm/dd hh:nn:ss');
  OS := 'Unk';
  UserPath := GetUserDir;
  UserAppsDataPath := UserPath;
  {$IFDEF Linux}
    OS := 'Linux';
    CRLF := #10;
    LangStr := GetEnvironmentVariable('LANG');
    x := pos('.', LangStr);
    LangStr := Copy(LangStr, 0, 2);
    wxbitsrun := 0;
    OSTarget:= '';
    // Get mail client
  {$ENDIF}
  {$IFDEF WINDOWS}
    OS := 'Windows ';
    CRLF := #13#10;
    // get user data folder
    s := ExtractFilePath(ExcludeTrailingPathDelimiter(GetAppConfigDir(False)));
    if Ord(WindowsVersion) < 7 then
      UserAppsDataPath := s                     // NT to XP
    else
    UserAppsDataPath := ExtractFilePath(ExcludeTrailingPathDelimiter(s)) + 'Roaming'; // Vista to W10
    LazGetShortLanguageID(LangStr);
    // Get mail client
    {$IFDEF WIN32}
      OSTarget := '32 bits';
    {$ENDIF}
    {$IFDEF WIN64}
      OSTarget := '64 bits';
    {$ENDIF}
  {$ENDIF}
  GetSysInfo(OsInfo);
  ProgName := 'MailsInBox';
  // Chargement des chaînes de langue...
  LangFile := TBbIniFile.Create(ExtractFilePath(Application.ExeName) + LowerCase(ProgName)+'.lng');
  LangNums := TStringList.Create;
  MailsInBoxAppsData := UserAppsDataPath + PathDelim + ProgName + PathDelim;
  if not DirectoryExists(MailsInBoxAppsData) then
    CreateDir(MailsInBoxAppsData);
  CheckingMail:= false;
end;

procedure TFMailsInBox.FormActivate(Sender: TObject);
var
  i: integer;
  defmailcli: string;
begin
  if not first then exit;
  First:= false;
  FSettings.Settings.AppName:= LowerCase(ProgName);
  FAccounts.Accounts.AppName := LowerCase(ProgName);
  ConfigFile := MailsInBoxAppsData + ProgName + '.xml';
  if not FileExists(ConfigFile) then
  begin
    if FileExists(MailsInBoxAppsData + ProgName + '.bk0') then
    begin
      RenameFile(MailsInBoxAppsData + ProgName + '.bk0', ConfigFile);
      for i := 1 to 5 do
        if FileExists(MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i))
        // Renomme les précédentes si elles existent
        then
          RenameFile(MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i),
            MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i - 1));
    end else SaveConfig(All);
  end;
  BaseUpdateUrl :='https://www.sdtp.com/versions/version.php?program=mailsinbox&version=%s&language=%s';
  ChkVerURL := 'https://www.sdtp.com/versions/versions.csv';
  version := GetVersionInfo.ProductVersion;
  LoadCfgFile(ConfigFile);
  // AboutBox.UrlUpdate:= BaseUpdateURl+Version+'&language='+Settings.LangStr;    // In Modlangue
  // AboutBox.LUpdate.Caption:= 'Recherche de mise à jour';                       // in Modlangue
  // Aboutbox.Caption:= 'A propos du Gestionnaire de contacts';                   // in ModLangue
  AboutBox.Width:= 300; // to have more place for the long product name
  AboutBox.Image1.Picture.Icon.LoadFromResourceName(HInstance, 'MAINICON');
  AboutBox.LProductName.Caption := GetVersionInfo.FileDescription;
  AboutBox.LCopyright.Caption :=
  GetVersionInfo.CompanyName + ' - ' + DateTimeToStr(CompileDateTime);
  AboutBox.LVersion.Caption := 'Version: ' + Version + ' (' + OS + OSTarget + ')';
  AboutBox.LUpdate.Hint := LastUpdateSearch + ': ' + DateToStr(FSettings.Settings.LastUpdChk);
  AboutBox.UrlWebsite := GetVersionInfo.Comments;
  BtnLaunchClient.Enabled:= not (FSettings.Settings.MailClient='');
  FAccounts.Accounts.SortType:= cdcindex;
  FAccounts.Accounts.DoSort;
  PopulateAccountsList;
  if  FAccounts.Accounts.count > 0 then LVAccounts.ItemIndex:= 0
  else LVAccounts.ItemIndex:=  -1;
  FSettings.Settings.OnChange := @SettingsOnChange;
  FSettings.Settings.OnStateChange := @SettingsOnStateChange;
  FAccounts.Accounts.OnChange:= @AccountsOnChange;
  FSettings.LStatus.Caption := OsInfo.VerDetail;
  traycount:= 0;
  //Save large buttons glyphs
  for i:=0 to length(BmpArray)-1 do
  begin
    BmpArray[i]:= TBitmap.create;
    BmpArray[i].Assign(TSpeedButton(PnlToolbar.Controls[i]).Glyph);
  end;
  // set small buttons only if asked
  if FSettings.Settings.SmallBtns then SetSmallBtns(FSettings.Settings.SmallBtns);
  Constraints.MinWidth:= BtnQuit.left+BtnQuit.width+10;
  if width < Constraints.MinWidth then width := Constraints.MinWidth;
  // Get default mail client and stores its name
  defmailcli:= FSettings.GetDefaultMailCllient;
  if length(FSettings.Settings.MailClient)=0 then  FSettings.Settings.MailClient:= defmailcli;
  //DefMailClient:=TrimFileExt(ExtractFileName(FSettings.Settings.MailClient));
  //if length(DefMailClient)>0 then DefMailClient[1]:= UpCase(DefMailClient[1]);
  //ExecuteProcess(FSettings.Settings.MailClient, '', []);
end;

procedure TFMailsInBox.ChangeBounds(Sender: TObject);
begin
  SettingsChanged:= FSettings.Settings.SavSizePos;
end;


// change size of buttons in toolbar

procedure TFMailsInBox. SetSmallBtns(small:boolean);
var
  Pict:TPicture;
  i: integer;
  ImgName:string;
begin
  if Small then PnlToolbar.height:= 34
  else PnlToolbar.height:= 42;
  Pict:= TPicture.Create;
  for i:= 0 to PnlToolbar.ControlCount-1 do
  begin
    ImgName:= uppercase(TSpeedButton(PnlToolbar.Controls[i]).name);
    ImgName:=copy(ImgName,4, length(ImgName)-3);
    if Small then
      begin
        if PnlToolbar.Controls[i].ClassType=TSpeedButton then
      try
        TSpeedButton(PnlToolbar.Controls[i]).Height:= 24;
        TSpeedButton(PnlToolbar.Controls[i]).Width:=24;
        Pict.LoadFromResourceName(HInstance, ImgName+'16');
        TSpeedButton(PnlToolbar.Controls[i]).Left:= 10+i*(24+8);
        TSpeedButton(PnlToolbar.Controls[i]).Glyph.Assign(Pict.Bitmap);
      except
      end;

      end else
      begin
        TSpeedButton(PnlToolbar.Controls[i]).Height:= 32;
        TSpeedButton(PnlToolbar.Controls[i]).Width:=32;
        TSpeedButton(PnlToolbar.Controls[i]).Left:= 10+i*(32+8);
        TSpeedButton(PnlToolbar.Controls[i]).Glyph.Assign(BmpArray[i]);
      end;
  end;
  Pict.Free;
end;

// Load configuration and database from file

procedure TFMailsInBox.LoadCfgFile(filename: string);
var
  winstate: TWindowState;
  i: integer;
begin
  with FSettings do
  begin
    Settings.LoadXMLFile(filename);
    if Settings.SavSizePos then
    try
      WinState := TWindowState(StrToInt('$' + Copy(Settings.WState, 1, 4)));
      if Winstate = wsMinimized then
        Application.Minimize
      else
        self.WindowState := WinState;
        self.Top := StrToInt('$' + Copy(Settings.WState, 5, 4));
        self.Left := StrToInt('$' + Copy(Settings.WState, 9, 4));
        self.Height := StrToInt('$' + Copy(Settings.WState, 13, 4));
        self.Width := StrToInt('$' + Copy(Settings.WState, 17, 4));
        self.PnlAccounts.width:= StrToInt('$' + Copy(Settings.WState, 21, 4));
        mailcolsiz:=copy(Settings.WState,25,20);      // to detect change of cols width
        For i:= 0 to 4 do self.SGMails.Columns[i].Width:= StrToInt('$'+Copy(Settings.WState,25+(i*4),4)) ;
    except
    end;
    if Settings.Startup and settings.StartMini then Application.Minimize;
    // Détermination de la langue (si pas dans settings, langue par défaut)
    if Settings.LangStr = '' then Settings.LangStr := LangStr;
    LangFile.ReadSections(LangNums);
    if LangNums.Count > 1 then
      for i := 0 to LangNums.Count - 1 do
      begin
        FSettings.CBLangue.Items.Add(LangFile.ReadString(LangNums.Strings[i], 'Language',
          'Aucune'));
        if LangNums.Strings[i] = Settings.LangStr then
        begin
          LangFound := True;
          FSettings.CBLangue.ItemIndex:= i;
        end;
      end;
    // Si la langue n'est pas traduite, alors on passe en Anglais
    if not LangFound then
    begin
      Settings.LangStr := 'en';
    end;
  end;
  FAccounts.Accounts.Reset;
  FAccounts.Accounts.LoadXMLfile(filename);
  Modlangue;
  SettingsChanged := False;
end;

procedure TFMailsInBox.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  s: string;
  i: integer;
begin
  // Check if columns width have changed
  s:='';
  For i:= 0 to 4 do s:= s+IntToHex(self.SGMails.Columns [i].Width, 4);
  if s <> mailcolsiz then SettingsChanged:= true;
  if CanClose then
  begin
    if FSettings.Settings.Startup then SetAutostart(progname, Application.exename)
    else UnSetAutostart(progname);
    if AccountsChanged then SaveConfig(All)
    else if SettingsChanged then SaveConfig(Setting) ;
    CloseAction := caFree;
  end else CloseAction := caNone;
end;


// Save configuration and database to file

function TFMailsInBox.SaveConfig(typ: TSaveMode): boolean;
var
  FilNamWoExt: string;
  i: integer;
begin
  Result := False;
  if (Typ= Setting) or (Typ = All) then
  with FSettings do
  begin
    //FSettings.Settings.DataFolder:= MailsInBoxAppsData;
    Settings.WState:= '';
    if self.Top < 0 then self.Top:= 0;
    if self.Left < 0 then self.Left:= 0;
    Settings.WState:= IntToHex(ord(self.WindowState), 4)+IntToHex(self.Top, 4)+
                      IntToHex(self.Left, 4)+IntToHex(self.Height, 4)+IntToHex(self.width, 4)+
                      IntToHex(self.PnlAccounts.width, 4);
    For i:= 0 to 4 do Settings.WState:= Settings.WState+IntToHex(self.SGMails.Columns [i].Width, 4);
    Settings.Version:= version;
    if FileExists (ConfigFile) then
    begin
      if (Typ = All) then
      begin
        // On sauvegarde les versions précédentes parce que la base de données a changé
        FilNamWoExt:= TrimFileExt(ConfigFile);
        if FileExists (FilNamWoExt+'.bk5')                   // Efface la plus ancienne
        then  DeleteFile(FilNamWoExt+'.bk5');                // si elle existe
        For i:= 4 downto 0
        do if FileExists (FilNamWoExt+'.bk'+IntToStr(i))     // Renomme les précédentes si elles existent
           then  RenameFile(FilNamWoExt+'.bk'+IntToStr(i), FilNamWoExt+'.bk'+IntToStr(i+1));
        RenameFile(ConfigFile, FilNamWoExt+'.bk0');
        FAccounts.Accounts.SaveToXMLfile(ConfigFile);
      end;
      // la base n'a pas changé, on ne fait pas de backup
      FSettings.settings.SaveToXMLfile(ConfigFile);
    end else
    begin
      FAccounts.Accounts.SaveToXMLfile(ConfigFile);
      settings.SaveToXMLfile(ConfigFile); ;
    end;
    result:= true;
  end;
end;

procedure TFMailsInBox.PopulateAccountsList;
var
  Listitem: TlistItem;
  i: Integer;
  AccBmp:Tbitmap;
Begin
  if FAccounts.Accounts.Count = 0 then exit;
  AccBmp:= TBitmap.create;
  LVAccounts.Clear;
  if Assigned(LVAccounts.SmallImages) then LVAccounts.SmallImages.Clear;
  for i := 0 to FAccounts.Accounts.Count-1 do
  Try
    AccBmp.LoadFromResourceName(HInstance, 'ACCOUNT');
    ListItem := LVAccounts.items.add;  // prépare l'ajout
    if FAccounts.Accounts.GetItem(i).Mails.count > 0 then
      DrawTheIcon(AccBmp, FAccounts.Accounts.GetItem(i).Mails.count ,
                  FAccounts.Accounts.GetItem(i).Color  );
    LVAccounts.SmallImages.AddMasked(AccBmp,$FF00FF);
    ListItem.ImageIndex := i;
    Listitem.Caption :=  FAccounts.Accounts.GetItem(i).Name;    // ajoute le nom
  Except
    ShowMessage(inttostr(i));
  end;
  LVAccounts.ItemIndex:= 0;
  //LVAccounts.SetFocus;
  AccBmp.free;
end;


// Event fired by any change of settings values

procedure TFMailsInBox.SettingsOnChange(Sender: TObject);
begin
  SettingsChanged := True;
end;

// Event fired by any state change (window state and position)

procedure TFMailsInBox.SettingsOnStateChange(Sender: TObject);
begin
  SettingsChanged := True;
end;

// Event fired by account change

procedure TFMailsInBox.AccountsOnChange(Sender: TObject);
begin
  AccountsChanged := True;
end;

procedure TFMailsInBox.BtnEditAccClick(Sender: TObject);
var
  Account: TAccount;
begin
  with FAccounts do
  begin
    if (TSpeedButton(Sender).Name='BtnAddAcc') then
    begin
      Account:= Default(TAccount);
      Caption:= BtnAddAcc.Hint;
    end;
    if (TSpeedButton(Sender).Name='BtnEditAcc') and  (LVAccounts.ItemIndex>=0) then
    begin
      Account:= Accounts.GetItem(LVAccounts.ItemIndex);
      Caption:= BtnEditAcc.Hint;
    end;
    EName.Text:= Account.Name;
    EServer.Text:= Account.Server;
    CBProtocol.ItemIndex:=ord(Account.Protocol);
    EUserName.Text:= Account.UserName;
    EPassword.Text:= Account.Password;
    EEmail.Text:= Account.Email;
    CBSSL.ItemIndex:= Account.SSL;
    EPort.Text:= IntToStr(Account.Port);
    CBSecureAuth.Checked:= Account.SecureAuth;
    EInterval.Value:= Account.Interval;
    EReplyEmail.Text:= Account.ReplyEmail;
    CBEnabledAcc.Checked:= Account.Enabled;
    CBColorAcc.Selected:= Account.Color;
    ESoundFile.Text:= Account.SoundFile;
    BtnPlaySound.Enabled:=not (length(ESoundFile.Text)=0);
    //mail client

    if ShowModal=mrOK then
    begin
      Account.Name:= EName.Text;
      Account.Server:= EServer.Text;
      Account.protocol:= TProtocols(CBProtocol.ItemIndex);
      Account.UserName:= EUserName.Text;
      Account.Password:= EPassword.Text;
      Account.Email:= EEmail.Text;
      Account.SSL:= CBSSL.ItemIndex;
      Account.Port:= StringToInt(EPort.Text);
      Account.SecureAuth:= CBSecureAuth.Checked;
      Account.Interval:= EInterval.Value;
      Account.ReplyEmail:= EReplyEmail.Text;
      Account.Enabled:= CBEnabledAcc.Checked;
      Account.Color:= CBColorAcc.Selected;
      if (TSpeedButton(Sender).Name='BtnEditAcc') and  (LVAccounts.ItemIndex>=0) then
      Accounts.ModifyAccount(LVAccounts.ItemIndex, Account)
      else Accounts.AddAccount(Account);
      PopulateAccountsList ();

    end;

  end;
end;

procedure TFMailsInBox.BtnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFMailsInBox.FormDestroy(Sender: TObject);
var
  i: integer;
begin
  if Assigned(LangNums) then LangNums.free;
  if Assigned(LangFile) then LangFile.free;
  for i:=0 to length(BmpArray)-1 do
    if Assigned(BmpArray[i]) then BmpArray[i].free;

end;

procedure TFMailsInBox.FormShow(Sender: TObject);
begin

end;

procedure TFMailsInBox.IdPOP3_1Connected(Sender: TObject);
begin
  LStatus.Caption:= 'Connecté au serveur';
end;

procedure TFMailsInBox.IdPOP3_1Disconnected(Sender: TObject);
begin
  LStatus.Caption:= 'Déconnecté du serveur';
end;

procedure TFMailsInBox.IdPOP3_1Status(ASender: TObject;
  const AStatus: TIdStatus; const AStatusText: string);
begin
end;

procedure TFMailsInBox.LVAccountsChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
var
  index: integer;
  i: integer;
begin
  index:= LVAccounts.ItemIndex;


end;

procedure TFMailsInBox.LVAccountsSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  s: string;
  i: integer;
  j: integer;
  k: integer;
begin
  i:= LVAccounts.ItemIndex;
  j:= LVAccounts.Items.count;
  if i >= 0 then
  begin
    s:= FAccounts.Accounts.GetItem(i).Name;
    BtnLog.Hint:= Format(BtnLogHint, [s]);
    BtnGetAccMail.Hint:= Format(BtnGetAccMailHint, [s]);
    BtnDelete.Hint:= Format(BtnDeleteHint, [s]);
    BtnEditAcc.Hint:= Format(BtnEditAccHint, [s]);
    BtnFirst.Enabled:= not (i=0);
    BtnPrev.Enabled:= BtnFirst.Enabled;
    BtnLast.Enabled:= not(i=j-1);
    BtnNext.Enabled:= BtnLast.Enabled;
    UpdateInfos;

  end;
  PopulateMailsList(i);



  //PnlAccountName.Caption:= 'Compte '+FAccounts.Accounts.GetItem(LVAccounts.ItemIndex).Name;
end;

procedure TFMailsInBox.UpdateInfos;
var
  i: integer;
  msgs: integer;
begin
  i:= LVAccounts.ItemIndex;
  if i >= 0 then
  begin
    msgs:= FAccounts.Accounts.GetItem(i).Mails.Count;
  MInfos.Clear;
  MInfos.Lines.Add(Format(AccountCaption, [FAccounts.Accounts.GetItem(i).Name]));
  MInfos.Lines.Add(Format(EmailCaption, [FAccounts.Accounts.GetItem(i).Email]));
  MInfos.Lines.Add(Format(LastCheckCaption, [DateTimeToString(FAccounts.Accounts.GetItem(i).LastFire)]));
  if msgs>1 then MInfos.Lines.Add(Format(MsgsFound, [msgs]))
  else MInfos.Lines.Add(Format(MsgFound, [msgs]));
  MInfos.Lines.add(Format(NextCheckCaption, [DateTimeToString(FAccounts.Accounts.GetItem(i).NextFire)]));

  end;
end;

procedure TFMailsInBox.PopulateMailsList(index: integer);
var
  i: integer;
begin
  //SGMails.Clear;
  SGMails.RowCount:=1;
  if (index<0) or (FAccounts.Accounts.GetItem(index).Mails.Count=0) then exit;
  SGMails.RowCount:= FAccounts.Accounts.GetItem(index).Mails.Count+1;
  for i:= 0 to SGMails.RowCount-2 do
  begin
    SGMails.Cells[0,i+1]:=FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageFrom;
    SGMails.Cells[1,i+1]:=FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageTo;
    //SGMails.Cells[1,i+1]:=FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageFrom;
    SGMails.Cells[2,i+1]:=FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageSubject;
    SGMails.Cells[3,i+1]:= DateTimeToString(FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageDate);
    SGMails.Cells[4,i+1]:= InttoStr(FAccounts.Accounts.GetItem(index).Mails.GetItem(i).MessageSize);
  end;

end;

procedure TFMailsInBox.MnuAccountPopup(Sender: TObject);
begin
  MnuMoveUp.Enabled:= not (LVAccounts.ItemIndex=0);
  MnuMoveDown.Enabled:= not (LVAccounts.ItemIndex=LVAccounts.Items.count-1);
end;

procedure TFMailsInBox.MnuMoveDownClick(Sender: TObject);
var
  oldndx: integer;
begin
  oldndx:= LVAccounts.ItemIndex;
  if oldndx<LVAccounts.Items.count-1 then
  begin
    FAccounts.Accounts.sorttype:= cdcNone;
    FAccounts.Accounts.ModifyField(oldndx, 'index', oldndx+1);
    FAccounts.Accounts.ModifyField(oldndx+1, 'index', oldndx);
    FAccounts.Accounts.sorttype:= cdcIndex;
    PopulateAccountsList();
    LVAccounts.ItemIndex:= oldndx+1;
  end;
end;

procedure TFMailsInBox.MnuMoveUpClick(Sender: TObject);
var
  oldndx: integer;
begin
  oldndx:= LVAccounts.ItemIndex;
  if oldndx>0 then
  begin
    FAccounts.Accounts.sorttype:= cdcNone;
    FAccounts.Accounts.ModifyField(oldndx, 'index', oldndx-1);
    FAccounts.Accounts.ModifyField(oldndx-1, 'index', oldndx);
    FAccounts.Accounts.sorttype:= cdcIndex;
    PopulateAccountsList();
    LVAccounts.ItemIndex:= oldndx-1;
  end;

end;

procedure TFMailsInBox.SGMailsBeforeSelection(Sender: TObject; aCol,
  aRow: Integer);
begin
  SGMails.Invalidate;
end;

procedure TFMailsInBox.SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
var
  s: string;
  R: TRect;
  bmp: Tbitmap;
  bmppos: integer;
begin
  if arow=0 then exit;
  // remove selection highlight if the control has not the focus
  if not SGHasFocus then
  begin
    SGMails.Canvas.Brush.Color := clWindow;
    SGMails.Canvas.FillRect (ARect);
    SGMails.Canvas.font.Color:= clDefault;  ;
  end;
  // Add mail image
  if acol=0 then
  begin
    bmppos:= 0;
    R.Left:= ARect.Left+2;
    R.Top:= ARect.Top+2;
    R.Right:=R.Left+18;
    R.Bottom:=R.Top+16;
    Bmp:= Tbitmap.Create;
    if FAccounts.Accounts.GetItem(LVAccounts.ItemIndex).Mails.GetItem(aRow-1).MessageNew then
      bmppos:= 1;
    if Pos ('multipart', FAccounts.Accounts.GetItem(LVAccounts.ItemIndex).Mails.GetItem(aRow-1).MessageContentType) >0 then
      bmppos:= bmppos+2;
    IlMail.GetBitmap(bmppos, bmp);
    SGMails.Canvas.StretchDraw(R, bmp);
    bmp.free;
    SGMails.Canvas.TextOut(ARect.Left+22,ARect.Top+3, SGMails.Cells[aCol, aRow]);
  end else
  begin
    SGMails.Canvas.TextOut(ARect.Left+2,ARect.Top+3, SGMails.Cells[aCol, aRow]);
  end;
end;

// 3 next Procedures to remove selection highlight when the list has not the focus

procedure TFMailsInBox.SGMailsEnter(Sender: TObject);
begin
   SGHasFocus:= true;
end;

procedure TFMailsInBox.SGMailsExit(Sender: TObject);
begin
  SGHasFocus:= false;
  //SGMails.Invalidate;
end;

procedure TFMailsInBox.SGMailsPrepareCanvas(sender: TObject; aCol,
  aRow: Integer; aState: TGridDrawState);
begin

end;

procedure TFMailsInBox.SGMailsSelectCell(Sender: TObject; aCol, aRow: Integer;
  var CanSelect: Boolean);
begin

end;

procedure TFMailsInBox.SGMailsSelection(Sender: TObject; aCol, aRow: Integer);
begin

end;

procedure TFMailsInBox.BtnSettingsClick(Sender: TObject);
var
  i, oldlng, oldmailsel, old: integer;
begin
  with FSettings do
  begin
    oldmailsel:=-1;
    CBStartup.Checked:= Settings.Startup;
    CBStartMini.Checked:= Settings.StartMini;
    CBSavSizePos.Checked:= Settings.SavSizePos;
    CBMailClientMini.Checked:= Settings.MailClientMini;
    CBRestNewMsg.Checked:= Settings.RestNewMsg;
    CBSaveLogs.Checked:= Settings.SaveLogs;
    CBNoChkNewVer.Checked:= Settings.NoChkNewVer;
    CBStartupCheck.Checked:= Settings.StartupCheck;
    CBSmallBtns.Checked:= Settings.SmallBtns;
    GetDefaultMailCllient;    // and get mail clientrs list
    CBMailClient.Items.Clear; //Reset list of mail clients
    for i:=0 to length(MailClients)-1 do
    begin
      CBMailClient.Items.add(MailClients[i].Name);
      CBUrl.Checked:= MailClients[i].Url;
      if MailClients[i].tag=true  then
      begin
        oldmailsel:= i;   // selected mail client
        CBMailClient.ItemIndex:= oldmailsel;
      end;
      if (MailClients[i].Defaut=true) then
        if(oldmailsel<0) then CBMailClient.ItemIndex:= i; // default mail client
    end;
    ESoundFile.Text:= Settings.SoundFile;
    BtnPlaySound.Enabled:= not (length(ESoundFile.Text)=0);
    oldlng := CBLangue.ItemIndex;
    if ShowModal= mrOK then
    begin
      Settings.Startup := CBStartup.Checked;
      Settings.StartMini := CBStartMini.Checked;
      Settings.SavSizePos := CBSavSizePos.Checked;
      Settings.MailClientMini:= CBMailClientMini.Checked;
      Settings.RestNewMsg:= CBRestNewMsg.Checked;
      Settings.SaveLogs:= CBSaveLogs.Checked;
      Settings.NoChkNewVer := CBNoChkNewVer.Checked;
      Settings.StartupCheck:= CBStartupCheck.Checked;
      if Settings.SmallBtns <> CBSmallBtns.Checked then  // Buttons size has changed
      SetSmallBtns(CBSmallBtns.Checked);
      Settings.SmallBtns:= CBSmallBtns.Checked;
      Settings.MailClient:= MailClients[CBMailClient.ItemIndex].Command;
      Settings.MailClientName:= MailClients[CBMailClient.ItemIndex].Name;
      Settings.MailClientIsUrl:= CBUrl.Checked;
      //if CBMailClient.ItemIndex <> oldmailsel then       // remove old bold style
      //MailClients[CBMailClient.ItemIndex].Tag:= false;
      Settings.SoundFile:= ESoundFile.Text;
      Settings.LangStr := LangNums.Strings[CBLangue.ItemIndex];
      if FSettings.CBLangue.ItemIndex <> oldlng then ModLangue;
      if SettingsChanged then PopulateAccountsList();  // Needed to change language on hints
    end;
  end;
end;


procedure TFMailsInBox.BtnAboutClick(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TFMailsInBox.BtnDeleteClick(Sender: TObject);
begin
  ShowMessage('Todo: delete selected account');
end;

procedure TFMailsInBox.BtnGetAllMailClick(Sender: TObject);
begin
  EnableControls(false);
  Application.ProcessMessages;
  ShowMessage('Todo: Get all accounts mails');
  EnableControls(true);
end;


procedure TFMailsInBox.BtnGetAccMailClick(Sender: TObject);
var
  ndx: integer;
begin
  EnableControls(false);
  Application.ProcessMessages;
  ndx:= LVAccounts.ItemIndex;   // Current selected account
  if (ndx>=0) and not CheckingMail then
  begin
    if FAccounts.Accounts.GetItem(ndx).Protocol=ptcPOP3 then
    GetPop3Mail(ndx) ;
    PopulateMailsList(ndx);
    //UpdateInfos;
  end;
  EnableControls(true);
  PopulateAccountsList;
  LVAccounts.ItemIndex:=ndx;
end;

// retreive pop3 mail

function TFMailsInBox.GetPop3Mail(index: integer): boolean;
var
  msgs : Integer;
  idMsg: TIdMessage;
  i, siz: integer;
  sfrom: string;
  sto: string;
  mail: TMail;
  mails: TMailsList;
  min: TTime;
begin
  CheckingMail:= true;
  mails:= TMailsList.create;
  IdPOP3_1.Host:= FAccounts.Accounts.GetItem(index).Server;
  IdPOP3_1.Port:= FAccounts.Accounts.GetItem(index).Port;
  IdPOP3_1.Username:= FAccounts.Accounts.GetItem(index).UserName;
  IdPOP3_1.Password:= FAccounts.Accounts.GetItem(index).Password;
  sto:= FAccounts.Accounts.GetItem(index).Name;
  try
    LStatus.Caption:= 'Connexion au serveur '+IdPOP3_1.Host;
    idMsg:= TIdMessage.Create(self);
    IdPop3_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(idPop3_1);
    IdPop3_1.UseTLS := TIdUseTLS(FAccounts.Accounts.GetItem(index).SSL);
    IdPOP3_1.Connect;
    Application.ProcessMessages;
    msgs := IdPop3_1.CheckMessages;

    if msgs > 0 then
    begin
      //TAccount(FAccounts.Accounts.Items[index]^).Mails.Reset;
      //SGMails.RowCount:= msgs;
      for i:= 1 to msgs do
      begin
        if IdPop3_1.RetrieveHeader(i, idMsg) then
        begin
          siz:= IdPop3_1.RetrieveMsgSize(i);
          sfrom:= idMsg.From.Name;
          if length(sfrom)=0 then sfrom:= idMsg.From.Address;
          Mail.AccountName:= sto;
          Mail.MessageFrom:= sfrom;
          Mail.FromAddress:= idMsg.From.Address;
          Mail.MessageUIDL:= idMsg.UID;
          Mail.MessageSubject:= idMsg.Subject;
          Mail.MessageTo:= idMsg.Recipients[0].Address;
          Mail.MessageDate:= idMsg.Date;
          Mail.MessageSize:= siz;
          Mail.MessageContentType:=  IdMsg.ContentType ;
          if TAccount(FAccounts.Accounts.Items[index]^).Mails.FindUIDL(Mail.MessageUIDL)>=0 then
             Mail.MessageNew:= false else Mail.MessageNew:= true;
          Mails.AddMail(Mail);
        end;
      end;
    end;
    LStatus.Caption:= 'Déconnexion du serveur';
    idPop3_1.Disconnect;

    // log...
  except
    on E: Exception do
    begin
      LStatus.Caption:= E.Message;
    end;
  end;
  if msgs>1 then LStatus.Caption:= sto+' : '+Format(MsgsFound, [msgs])
  else LStatus.Caption:= sto+' : '+Format(MsgFound, [msgs]) ;
  // Update account checkmail dates
  FAccounts.Accounts.ModifyField(index, 'LASTFIRE', now);
  min:= EncodeTime(0,FAccounts.Accounts.GetItem(index).interval,0,0);
  FAccounts.Accounts.ModifyField(index, 'NEXTFIRE', now+min);
  TAccount(FAccounts.Accounts.Items[index]^).Mails.Reset;
  if Mails.count > 0 then
    for i:=0 to Mails.count-1 do
     TAccount(FAccounts.Accounts.Items[index]^).Mails.AddMail(Mails.GetItem(i));
  if assigned (Mails) then Mails.free;
  UpdateInfos;
  CheckingMail:= false;
end;




// Import external accounts
// Currently : complete mailattente accounts
//             Outlook 2007-2013 accounts (password is not retrieved)

procedure TFMailsInBox.BtnImportClick(Sender: TObject);
var
  i, j: integer;
  s: string;
begin
  with FImpex do
  begin
    if ShowModal=mrOK then
    begin
      j:=0;
      for i:=0 to ImpAccounts.count-1 do
      begin
        if LBImpex.Selected[i] then
        begin
          FAccounts.Accounts.AddAccount(ImpAccounts.GetItem(i));
          inc(j);
        end;
      end;
      PopulateAccountsList();
      if j>1 then s:= AccImportds else s:= AccImportd;
      MsgDlg(Caption, Format(s, [j, CBAccType.Items[CBAccType.ItemIndex]]),
        mtInformation, [mbOK], [OKBtn], 0);
    end;
  end;
end;

procedure TFMailsInBox.BtnLaunchClientClick(Sender: TObject);
begin
  //ShowMessage('Todo: Launch default mail client');
end;

procedure TFMailsInBox.BtnLogClick(Sender: TObject);
begin
  ShowMessage('Todo: display log file');
end;

procedure TFMailsInBox.BtnNavClick(Sender: TObject);
var
  i, j: integer;
begin
  i:= LVAccounts.ItemIndex;
  j:= LVAccounts.Items.count;
  if TSpeedButton(Sender).Name='BtnFirst' then LVAccounts.ItemIndex:= 0;
  if TSpeedButton(Sender).Name='BtnLast' then LVAccounts.ItemIndex:= j-1;
  if (TSpeedButton(Sender).Name='BtnNext') and (i<j-1) then LVAccounts.ItemIndex:= i+1;
  if (TSpeedButton(Sender).Name='BtnPrev') and (i>0) then LVAccounts.ItemIndex:= i-1;
end;

procedure TFMailsInBox.BtnQuitClick(Sender: TObject);
begin
  CanClose:= true;
  Close();
end;

// Disable controls during mail check to avoid conflicts
// Display hourglass cursor

procedure TFMailsInBox.EnableControls(Enable: boolean);
var
  i: integer;
  j: integer;
  curs: TCursor;
begin
  j:=0;
  if enable then
  begin
    for i := 0 to PnlToolbar.ControlCount - 1 do
      if (PnlToolbar.Controls[i] is TSpeedButton) then
      begin
        PnlToolbar.Controls[i].Enabled := ButtonStates[j];
        Inc(j);
      end;
      LVAccounts.Enabled:= true;
      SGMails.Enabled:= true;
      SCreen.Cursor:= DefCursor;
  end else
  begin
    Screen.Cursor:= crHourGlass;
    SetLength(ButtonStates, PnlToolbar.ControlCount);
    for i := 0 to PnlToolbar.ControlCount - 1 do
      if (PnlToolbar.Controls[i] is TSpeedButton) then
      begin
        ButtonStates[j] := PnlToolbar.Controls[i].Enabled;
        PnlToolbar.Controls[i].Enabled:= false;
        Inc(j);
      end;
    LVAccounts.Enabled:= false;
    SGMails.Enabled:= false;

  end;
end;

procedure TFMailsInBox.DrawTheIcon(Bmp: TBitmap; NewCount: integer; CircleColor: TColor);
var
  i : integer;
  s: string;

begin
  With Bmp.Canvas do
  begin
    Brush.Style := bsSolid;
    Brush.Color := circlecolor;
    // Meme couleur que le masque
    If Brush.Color = $FF00FF then Brush.Color:= $FF01FF;
    Pen.Color := Brush.Color;
    Pen.Width:= 2;
    if NewCount > 9 then Rectangle(3,1,15,12)
    else Ellipse(3,4,15,16);
    // font
    Font.Name := 'Arial';
    Font.Style := [fsBold];
    // Clair ou foncé ?
    if DarkColor(Brush.color) then Font.Color := clWhite
    else     Font.Color := clBlack;
    //Font.Color := clWhite;
    Font.Size := 7;
    Brush.Style := bsClear;
    // number
    Case NewCount of
      -2: begin
            s:= '?';
            Font.Size := 8;
          end;
      -1: begin
            s:= '!';
            Font.Size := 8;
          end;
      else s:= IntToStr(NewCount);
    end;
    i := TextWidth(s) div 2;
    TextOut(9-i,4,s);
  end;
end;

// Animate tray icon

procedure TFMailsInBox.TimerTrayTimer(Sender: TObject);
begin
  if not CheckingMail then
  begin;
    TrayMail.Icon.LoadFromResourceName(HINSTANCE, 'XATT'+(InttoStr(traycount)));
    inc (traycount);
    if traycount > 5 then traycount:=0;
  end else
  begin
    TrayMail.Icon.LoadFromResourceName(HINSTANCE, 'TRAYICON');
  end;
end;

// Load control captions and text variable translations
// from mailsinbox.lng

procedure TFMailsInBox.ModLangue;
begin
  with LangFile do
  begin
    LangStr:=FSettings.Settings.LangStr;
    //Main Form
    Caption:=ReadString(LangStr,'Caption','Courrier en attente');
    OKBtn:= ReadString(LangStr, 'OKBtn','OK');
    YesBtn:=ReadString(LangStr,'YesBtn','Oui');
    NoBtn:=ReadString(LangStr,'NoBtn','Non');
    CancelBtn:=ReadString(LangStr,'CancelBtn','Annuler');
    BtnFirst.Hint:=ReadString(LangStr,'BtnFirst.Hint',BtnFirst.Hint);
    BtnPrev.Hint:=ReadString(LangStr,'BtnPrev.Hint',BtnPrev.Hint);
    BtnNext.Hint:=ReadString(LangStr,'BtnNext.Hint',BtnNext.Hint);
    BtnLast.Hint:=ReadString(LangStr,'BtnLast.Hint',BtnLast.Hint);
    BtnLogHint:=ReadString(LangStr,'BtnLogHint','Journal du compte %s');
    BtnGetAllMail.Hint:=ReadString(LangStr,'BtnGetAllMail.Hint',BtnGetAllMail.Hint);
    BtnGetAccMailHint:=ReadString(LangStr,'BtnGetAccMailHint','Vérifier le compte %s');
    BtnLaunchClient.Hint:=ReadString(LangStr,'BtnLaunchClient.Hint', BtnLaunchClient.Hint);
    BtnDeleteHint:=ReadString(LangStr,'BtnDeleteHint','Supprimer le compte %s');
    BtnAddAcc.Hint:=ReadString(LangStr,'BtnAddAcc.Hint',BtnAddAcc.Hint);
    BtnEditAccHint:=ReadString(LangStr,'BtnEditAccHint','Modifier le compte %s');
    BtnSettings.Hint:=ReadString(LangStr,'BtnSettings.Hint',BtnSettings.Hint);
    BtnAbout.Hint:=ReadString(LangStr,'BtnAbout.Hint',BtnAbout.Hint);
    BtnClose.Hint:=ReadString(LangStr,'BtnClose.Hint',BtnClose.Hint);
    BtnQuit.Hint:=ReadString(LangStr,'BtnQuit.Hint',BtnQuit.Hint);
    AccountCaption:=ReadString(LangStr,'AccountCaption','Compte: %s');
    EmailCaption:=ReadString(LangStr,'EmailCaption','Courriel: %s');
    LastCheckCaption:=ReadString(LangStr,'LastCheckCaption','Dernière vérification: %s');
    NextCheckCaption:=ReadString(LangStr,'NextCheckCaption','Prochaine vérification: %s');
    AccImportd:=ReadString(LangStr,'AccImportd','%d compte %s importé');
    AccImportds:=ReadString(LangStr,'AccExportds','%d comptes %s importés');
    MsgFound:=ReadString(LangStr,'MsgFound','%d message trouvé');
    MsgsFound:=ReadString(LangStr,'MsgsFound','%d messages trouvés');

    // About
    NoLongerChkUpdates:=ReadString(LangStr,'NoLongerChkUpdates','Ne plus rechercher les mises à jour');
    LastUpdateSearch:=ReadString(LangStr,'LastUpdateSearch','Dernière recherche de mise à jour');
    UpdateAvailable:=ReadString(LangStr,'UpdateAvailable','Nouvelle version %s disponible');
    UpdateAlertBox:=ReadString(LangStr,'UpdateAlertBox','Version actuelle: %sUne nouvelle version %s est disponible');
    Aboutbox.Caption:=ReadString(LangStr,'Aboutbox.Caption','A propos du Gestionnaire de Contacts');
    AboutBox.LUpdate.Caption:=ReadString(LangStr,'AboutBox.LUpdate.Caption','Recherche de mise à jour');
    AboutBox.UrlUpdate:=Format(BaseUpdateURl,[Version,LangStr]);

    //Accounts
    FAccounts.BtnOk.Caption:=OKBtn;
    FAccounts.BtnCancel.Caption:=CancelBtn;
    FAccounts.LAccName.Caption:=ReadString(LangStr,'FAccounts.LAccName.Caption',FAccounts.LAccName.Caption);
    FAccounts.LHost.Caption:=ReadString(LangStr,'FAccounts.LHost.Caption',FAccounts.LHost.Caption);
    FAccounts.LProtocol.Caption:=ReadString(LangStr,'FAccounts.LProtocol.Caption', FAccounts.LProtocol.Caption);
    FAccounts.CBProtocol.Items[0]:=ReadString(LangStr,'FAccounts.CBProtocol.Items_0',FAccounts.CBProtocol.Items[0]);
    FAccounts.LUserName.Caption:=ReadString(LangStr,'FAccounts.LUserName.Caption',FAccounts.LUserName.Caption);
    FAccounts.LPassword.Caption:=ReadString(LangStr,'FAccounts.LPassword.Caption',FAccounts.LPassword.Caption);
    FAccounts.LEmail.Caption:=ReadString(LangStr,'FAccounts.LEmail.Caption',FAccounts.LEmail.Caption);
    FAccounts.LColor.Caption:=ReadString(LangStr,'FAccounts.LColor.Caption',FAccounts.LColor.Caption);
    FAccounts.LMailClient.Caption:=ReadString(LangStr,'FAccounts.LMailClient.Caption',FAccounts.LMailClient.Caption);
    FAccounts.LSoundFile.Caption:=ReadString(LangStr,'FAccounts.LSoundFile.Caption',FAccounts.LSoundFile.Caption);
    FAccounts.LSSL.Caption:=ReadString(LangStr,'FAccounts.LSSL.Caption',FAccounts.LSSL.Caption);
    FAccounts.CBSSL.Items[0]:=ReadString(LangStr,'FAccounts.CBSSL.Items_0',FAccounts.CBSSL.Items[0]);
    FAccounts.CBSSL.Items[1]:=ReadString(LangStr,'FAccounts.CBSSL.Items_1',FAccounts.CBSSL.Items[1]);
    FAccounts.CBSSL.Items[2]:=ReadString(LangStr,'FAccounts.CBSSL.Items_2',FAccounts.CBSSL.Items[2]);
    FAccounts.LPort.Caption:=ReadString(LangStr,'FAccounts.LPort.Caption',FAccounts.LPort.Caption);
    FAccounts.CBSecureAuth.Caption:=ReadString(LangStr,'FAccounts.CBSecureAuth.Caption',FAccounts.CBSecureAuth.Caption);
    FAccounts.LInterval.Caption:=ReadString(LangStr,'FAccounts.LInterval.Caption',FAccounts.LInterval.Caption);
    FAccounts.CBShowPass.Caption:=ReadString(LangStr,'FAccounts.CBShowPass.Caption',FAccounts.CBShowPass.Caption);
    FAccounts.LReply.Caption:=ReadString(LangStr,'FAccounts.LReply.Caption',FAccounts.LReply.Caption);
    FAccounts.CBEnabledAcc.Caption:=ReadString(LangStr,'FAccounts.CBEnabledAcc.Caption',FAccounts.CBEnabledAcc.Caption);
    FAccounts.BtnMailClient.Hint:=ReadString(LangStr,'FAccounts.BtnMailClient.Hint',FAccounts.BtnMailClient.Hint);
    FAccounts.BtnPlaySound.Hint:=ReadString(LangStr,'FAccounts.BtnPlaySound.Hint',FAccounts.BtnPlaySound.Hint);
    FAccounts.BtnSoundFile.Hint:=ReadString(LangStr,'FAccounts.BtnSoundFile.Hint',FAccounts.BtnSoundFile.Hint);

    //Settings
    FSettings.BtnOK.Caption:=OKBtn;
    FSettings.BtnCancel.Caption:=CancelBtn;
    FSettings.Caption:=ReadString(LangStr,'FSettings.Caption',FSettings.Caption);
    FSettings.GBSystem.Caption:=ReadString(LangStr,'FSettings.GBSystem.Caption',FSettings.GBSystem.Caption);
    FSettings.CBStartup.Caption:=ReadString(LangStr,'FSettings.CBStartup.Caption',FSettings.CBStartup.Caption);
    FSettings.CBStartMini.Caption:=ReadString(LangStr,'FSettings.CBStartMini.Caption',FSettings.CBStartMini.Caption);
    FSettings.CBSavSizePos.Caption:=ReadString(LangStr,'FSettings.CBSavSizePos.Caption',FSettings.CBSavSizePos.Caption);
    FSettings.CBMailClientMini.Caption:=ReadString(LangStr,'FSettings.CBMailClientMini.Caption',FSettings.CBMailClientMini.Caption);
    FSettings.CBRestNewMsg.Caption:=ReadString(LangStr,'FSettings.CBRestNewMsg.Caption',FSettings.CBRestNewMsg.Caption);
    FSettings.CBSaveLogs.Caption:=ReadString(LangStr,'FSettings.CBSaveLogs.Caption',FSettings.CBSaveLogs.Caption);
    FSettings.CBNoChkNewVer.Caption:=ReadString(LangStr,'FSettings.CBNoChkNewVer.Caption',FSettings.CBNoChkNewVer.Caption);
    FSettings.CBStartupCheck.Caption:=ReadString(LangStr,'FSettings.CBStartupCheck.Caption',FSettings.CBStartupCheck.Caption);
    FSettings.CBSmallBtns.Caption:=ReadString(LangStr,'FSettings.CBSmallBtns.Caption',FSettings.CBSmallBtns.Caption);
    FSettings.LMailClient.Caption:=FAccounts.LMailClient.Caption;
    FSettings.LSoundFile.Caption:=FAccounts.LSoundFile.Caption;
    FSettings.LLangue.Caption:=ReadString(LangStr,'FSettings.LLangue.Caption',FSettings.LLangue.Caption);
    FSettings.BtnMailClient.Hint:=FAccounts.BtnMailClient.Hint;
    FSettings.BtnPlaySound.Hint:=FAccounts.BtnPlaySound.Hint;
    FSettings.BtnSoundFile.Hint:=FAccounts.BtnSoundFile.Hint;
    FSettings.CBUrl.Hint:=ReadString(LangStr,'FSettings.CBUrl.Hint',FSettings.CBUrl.Hint);
    FSettings.GMailWeb:=ReadString(LangStr,'FSettings.GMailWeb','Site Web de GMail');
    FSettings.OutlookWeb:=ReadString(LangStr,'FSettings.OutlookWeb','Site Web d''Outlook.com');
    FSettings.Win10Mail:=ReadString(LangStr,'FSettings.Win10Mail','Application Courrier de Windows 10');

    // Choose mail client
    FMailClientChoose.BtnOK.Caption:=OKBtn;
    FMailClientChoose.BtnCancel.Caption:=CancelBtn;
    FMailClientChoose.Caption:=ReadString(LangStr,'FMailClientChoose.Caption',FMailClientChoose.Caption);
    FMailClientChoose.LName.Caption:=ReadString(LangStr,'FMailClientChoose.LName.Caption',FMailClientChoose.LName.Caption);
    FMailClientChoose.LCommand.Caption:=ReadString(LangStr,'FMailClientChoose.LCommand.Caption',FMailClientChoose.LCommand.Caption);
    FMailClientChoose.CBUrl.Hint:=FSettings.CBUrl.Hint;
    FMailClientChoose.BtnMailClient.Hint:=FAccounts.BtnMailClient.Hint;

    // Impex
    FImpex.BtnOK.Caption:=OKBtn;
    FImpex.BtnCancel.Caption:=CancelBtn;
    FImpex.LAccTyp.Caption:=ReadString(LangStr,'FImpex.LAccTyp.Caption',FImpex.LAccTyp.Caption);
    FImpex.Caption:=ReadString(LangStr,'FImpex.Caption',FImpex.Caption);
    FImpex.LFilename.Caption:=ReadString(LangStr,'FImpex.LFilename.Caption',FImpex.LFilename.Caption);
    Fimpex.MailattAccName:=ReadString(LangStr,'Fimpex.MailattAccName', 'Comptes MailAttente');
    Fimpex.OutlAccName:=ReadString(LangStr,'Fimpex.OutlAccName', 'Comptes Outlook 2007-2010');
  end;

end;

end.

