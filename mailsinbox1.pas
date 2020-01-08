unit mailsinbox1;

{$mode objfpc}{$H+}

interface

uses
  {$IFDEF WINDOWS}
  Win32Proc,
  {$ENDIF} Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Grids, ComCtrls, Buttons, Menus, IdPOP3, IdSSLOpenSSL,
  IdExplicitTLSClientServerBase, IdMessage, IdIMAP4, accounts1, lazbbutils,
  lazbbinifiles, lazbbosversion, LazUTF8, settings1, lazbbautostart, lazbbabout,
  Impex1, mailclients1, uxtheme, Types, IdComponent, fptimer,
  IdMessageCollection, RichMemo, csvdocument, log1;

type
  TSaveMode = (None, Setting, All);
  TBtnSize=(Small, Large);

  {Helper to detect stringgrid columns width change}
  {Can access to protected function GetGridState   }

  TCustomGridHelper = class helper for TCustomGrid
    function GetGridState : TGridState;
  end;

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
    BtnAccountLog: TSpeedButton;
    BtnLog: TSpeedButton;
    BtnQuit: TSpeedButton;
    BtnSettings: TSpeedButton;
    GBInfos: TGroupBox;
    IdIMAP4_1: TIdIMAP4;
    IdPOP3_1: TIdPOP3;
    ILMail: TImageList;
    ILTray: TImageList;
    ImgAccounts: TImageList;
    LStatus: TLabel;
    LVAccounts: TListView;
    MnuInfos: TMenuItem;
    MnuMoveDown: TMenuItem;
    MnuMoveUp: TMenuItem;
    PnlAccounts: TPanel;
    PnlInfos: TPanel;
    PnlMails: TPanel;
    PnlLeft: TPanel;
    PnlStatus: TPanel;
    PnlToolbar: TPanel;
    BtnImport: TSpeedButton;
    MnuAccount: TPopupMenu;
    MnuMails: TPopupMenu;
    RMInfos: TRichMemo;
    SplitterV: TSplitter;
    SplitterH: TSplitter;
    SGMails: TStringGrid;
    GetMailTimer: TTimer;
    TrayTimer: TTimer;
    TrayMail: TTrayIcon;
    procedure BtnAboutClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure BtnGetAccMailClick(Sender: TObject);
    procedure BtnGetAllMailClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure BtnLaunchClientClick(Sender: TObject);
    procedure BtnAccountLogClick(Sender: TObject);
    procedure BtnLogClick(Sender: TObject);
    procedure BtnQuitClick(Sender: TObject);
    procedure BtnSettingsClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure DoChangeBounds(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnEditAccClick(Sender: TObject);
    procedure BtnCloseClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GetMailTimerTimer(Sender: TObject);
    procedure IdPOP3_1Connected(Sender: TObject);
    procedure IdPOP3_1Disconnected(Sender: TObject);
    procedure IdPOP3_1Status(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure LVAccountsDrawItem(Sender: TCustomListView; AItem: TListItem;
      ARect: TRect; AState: TOwnerDrawState);
    procedure LVAccountsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure MnuAccountPopup(Sender: TObject);
    procedure MnuInfosClick(Sender: TObject);
    procedure MnuMoveDownClick(Sender: TObject);
    procedure MnuMoveUpClick(Sender: TObject);
    procedure SGMailsBeforeSelection(Sender: TObject; aCol, aRow: Integer);
    procedure SGMailsClick(Sender: TObject);
    procedure SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
      aRect: TRect; aState: TGridDrawState);
    procedure SGMailsEnter(Sender: TObject);
    procedure SGMailsExit(Sender: TObject);
    procedure SGMailsKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SGMailsMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SGMailsMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TrayTimerTimer(Sender: TObject);
    procedure OnChkMailTimer(Sender: TObject);
  private
    Initialized: boolean;
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
    ChkMailTimerTick: integer;
    canCloseMsg: string;
    CanClose: boolean;
    AccountCaption, EmailCaption, LastCheckCaption, NextCheckCaption : string;
    BaseUpdateUrl, ChkVerURL, version: string;
    NoLongerChkUpdates, LastUpdateSearch, UpdateAvailable, UpdateAlertBox: string;
    OKBtn, YesBtn, NoBtn, CancelBtn: string;
    BtnLogHint, BtnGetAccMailHint, BtnDeleteHint, BtnEditAccHint: string;
    BmpArray: array of TBitmap;
    AccImportd, AccImportds: string;
    mailcolsiz: string;
    CheckingMail: boolean;
    SGHasFocus: boolean;
    DefCursor: TCursor;
    MsgFound, MsgsFound: string;
    ChkMailTimer: TFPTimer;
    TrayTimerTick: integer;
    TrayTimerBmp: TBitmap;
    LStatusCaption: String;
    MainLog: String;
    SessionLog: TStringList;
    CurAccPend: integer;
    ConnectToServer: string;
    DisconnectServer: string;
    ConnectedToServer: string;
    DisconnectedServer: string;
    ConnectErrorMsg: string;
    AccountChanged: string;
    AccountAdded: string;
    AccountDisabled: string;
    AccountEnabled: string;
    AccountStatus: string;
    Logfile: string;
    TrayHint: string;
    procedure Initialize;
    procedure LoadCfgFile(filename: string);
    procedure SettingsOnChange(Sender: TObject);
    procedure SettingsOnStateChange(Sender: TObject);
    procedure AccountsOnChange(Sender: TObject);
    function SaveConfig(Typ: TSaveMode): boolean;
    procedure PopulateAccountsList;
    procedure ModLangue;
    procedure SetSmallBtns(small: boolean);
    function GetPendingMail(index: integer): integer;
    procedure PopulateMailsList(index: integer);
    procedure EnableControls(Enable: boolean);
    procedure UpdateInfos;
    procedure DrawTheIcon(Bmp: TBitmap; NewCount: integer; CircleColor: TColor);
    function MailChecking(status: boolean): boolean;
    procedure GetMailInfos(CurName: String; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
    function LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
    procedure GetAccMail(ndx: integer);
  public
    OsInfo: TOSInfo;

  end;

var
  FMailsInBox: TFMailsInBox;

implementation

{$R *.lfm}

// Unprotect TStringgrid function GetGridState
// Allow detection of columns width, row heigth change in mouseup event

function TCustomGridHelper.GetGridState: TGridState;
begin
  Result := FGridState;
end;

// TFMailsInBox : This is the main form of the program

procedure TFMailsInBox.FormCreate(Sender: TObject);
var
  s: string;
begin
  // Initialize timers stuff
  ChkMailTimer:= TFPTimer.Create(self);
  ChkMailTimer.Interval:= 100;
  ChkMailTimer.UseTimerThread:= true;
  ChkMailTimer.Enabled:= true;
  ChkMailTimer.OnTimer:= @OnChkMailTimer;
  ChkMailTimerTick:= 0;
  //ChkMailTimer.StartTimer;
  TrayTimerTick:=0;
  TrayTimerBmp:= TBitmap.Create;
  // Flag needed to execute once some processes in Form activation
  Initialized := False;
  MainLog:= '';
  SessionLog:= TStringList.Create;
  LogAddLine(-1, now, 'Opening MailsInBox');
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
  if not DirectoryExists(MailsInBoxAppsData) then CreateDir(MailsInBoxAppsData);
  LogFile:= MailsInBoxAppsData+ProgName+'.log';
  // Mail checking process flag
  CheckingMail:= false;
end;

function TFMailsInBox.LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
begin
 result:= SessionLog.Add('"'+Inttostr(acc)+'",'+
                      '"'+DateTimeToString(dat)+'",'+
                      '"'+evnt+'"');
end;

// Form activation only needed once

procedure TFMailsInBox.FormActivate(Sender: TObject);

begin
  Initialize;

end;

// Initializing stuff

procedure TFMailsInBox.Initialize;
var
  i: integer;
  defmailcli: string;
  IniFile: TBbIniFile;
  tmplog: TstringList;
  min: TDateTime;
  CurAcc: TAccount;
begin
  if initialized then exit;

  FSettings.Settings.AppName:= LowerCase(ProgName);
  FAccounts.Accounts.AppName := LowerCase(ProgName);
  ConfigFile := MailsInBoxAppsData + ProgName + '.xml';
  if not FileExists(ConfigFile) then
  begin
    if FileExists(MailsInBoxAppsData + ProgName + '.bk0') then
    begin
      LogAddLine(-1, now, 'Retreive configuration backup');
      RenameFile(MailsInBoxAppsData + ProgName + '.bk0', ConfigFile);
      for i := 1 to 5 do
        if FileExists(MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i))
        // Renomme les précédentes si elles existent
        then
          RenameFile(MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i),
            MailsInBoxAppsData + ProgName + '.bk' + IntToStr(i - 1));
    end else
    begin
      SaveConfig(All);
      LogAddLine(-1, now, 'Creating new configuration file');
    end;

  end;
  // Check inifile with URLs
  IniFile:= TBbInifile.Create('mailsinbox.ini');
  BaseUpdateUrl:= IniFile.ReadString('urls', 'BaseUpdateUrl',
    'https://www.sdtp.com/versions/version.php?program=mailsinbox&version=%s&language=%s');
  ChkVerURL := IniFile.ReadString('urls', 'ChkVerURL',
    'https://www.sdtp.com/versions/versions.csv');
  if Assigned(IniFile) then IniFile.free;
  version := GetVersionInfo.ProductVersion;
  // AboutBox.UrlUpdate:= BaseUpdateURl+Version+'&language='+Settings.LangStr;    // In Modlangue
  // AboutBox.LUpdate.Caption:= 'Recherche de mise à jour';                       // in Modlangue
  // Aboutbox.Caption:= 'A propos du Gestionnaire de contacts';                   // in ModLangue
  AboutBox.Width:= 300; // to have more place for the long product name
  AboutBox.Image1.Picture.Icon.LoadFromResourceName(HInstance, 'MAINICON');
  AboutBox.LProductName.Caption := GetVersionInfo.FileDescription;
  AboutBox.LCopyright.Caption := GetVersionInfo.CompanyName + ' - ' + DateTimeToStr(CompileDateTime);
  AboutBox.LVersion.Caption := 'Version: ' + Version + ' (' + OS + OSTarget + ')';
  AboutBox.UrlWebsite := GetVersionInfo.Comments;
  LogAddLine(-1, now, AboutBox.LVersion.Caption);
  LoadCfgFile(ConfigFile);
  AboutBox.LUpdate.Hint := LastUpdateSearch + ': ' + DateToStr(FSettings.Settings.LastUpdChk);
  tmplog:= TStringList.Create;
  if FileExists(logfile) then
  begin
    tmpLog.LoadFromFile(logfile);
    MainLog:= tmpLog.text;
  end;
  BtnLaunchClient.Enabled:= not (FSettings.Settings.MailClient='');
  FAccounts.Accounts.SortType:= cdcindex;
  FAccounts.Accounts.DoSort;
  PopulateAccountsList;
  if  FAccounts.Accounts.count > 0 then LVAccounts.ItemIndex:= 0
  else LVAccounts.ItemIndex:=  -1;
  FSettings.Settings.OnChange := @SettingsOnChange;
  FSettings.Settings.OnStateChange := @SettingsOnStateChange;
  FAccounts.Accounts.OnChange:= @AccountsOnChange;

  for i:= 0 to FAccounts.Accounts.count-1 do
  begin
    // print accounts in log
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled then
    begin
      AccountStatus:=Format('%s, Id : %u', [Format(AccountEnabled,[CurAcc.Name]),CurAcc.UID]);
      // Delay mail checking until next interval if no check at startup
      if not FSettings.Settings.StartupCheck then
      begin
        min:= EncodeTime(0,Curacc.interval,0,0);
        if CurAcc.NextFire < now then
        FAccounts.Accounts.ModifyField(i, 'NEXTFIRE', now+min);
      end;
    end
    else AccountStatus:=Format('%s, Id : %u',[Format(AccountDisabled, [CurAcc.Name]),CurAcc.UID]);
    LogAddLine(-1, now, AccountStatus);
  end;
  // Update infos as we may have changed nextfire time
  UpdateInfos;
  FSettings.LStatus.Caption := OsInfo.VerDetail;
  //Save large buttons glyphs in an array
  SetLength(BmpArray, PnlToolbar.ControlCount);
  for i:=0 to PnlToolbar.ControlCount-1 do
  begin
    BmpArray[i]:= TBitmap.create;
    BmpArray[i].Assign(TSpeedButton(PnlToolbar.Controls[i]).Glyph);
  end;
  // set small buttons only if asked
  if FSettings.Settings.SmallBtns then SetSmallBtns(FSettings.Settings.SmallBtns);
  Constraints.MinWidth:= BtnQuit.left+BtnQuit.width+10;
  if width < Constraints.MinWidth then width := Constraints.MinWidth;
  // Get default mail client and stores its name
  // initialize settings form to properly set default mail client
  defmailcli:= FSettings.GetDefaultMailCllient;
  // Launch config dailog to set mail client and other stuff
  if length(FSettings.Settings.MailClient)=0 then
  begin
    FSettings.Settings.MailClient:= defmailcli;
    BtnSettingsClick(self);
  end;

  LogAddLine(-1, now, 'Mail Client: '+FSettings.Settings.MailClientName);
  // TStringgrid MousetoCell give cell nearest the mouse click if
  // this property is true. To get -1 when mouse is outside a cell then
  SGMails.AllowOutboundEvents:=false;
  TrayMail.Hint:= 'Aucun message';
  GetMailTimer.enabled:= true;

  Initialized:= true;

end;

// Change of form size or splitters position

procedure TFMailsInBox.DoChangeBounds(Sender: TObject);
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
    LogAddLine(-1, now, 'Load settings');
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
        self.PnlLeft.width:= StrToInt('$' + Copy(Settings.WState, 21, 4));
        self.PnlAccounts.Height:= StrToInt('$' + Copy(Settings.WState, 25, 4));
        For i:= 0 to 4 do self.SGMails.Columns[i].Width:= StrToInt('$'+Copy(Settings.WState,29+(i*4),4)) ;
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
  LogAddLine(-1, now, 'Load accounts');
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
  if CanClose then
  begin
    if FSettings.Settings.Startup then SetAutostart(progname, Application.exename)
    else UnSetAutostart(progname);
    if AccountsChanged then SaveConfig(All)
    else if SettingsChanged then SaveConfig(Setting) ;
    CloseAction := caFree;
    LogAddLine(-1, now, 'Closing MailsInBox');
    LogAddLine(-1, now, '************************');
    if FSettings.Settings.SaveLogs then
    begin
      s:= Mainlog+SessionLog.text;
      SessionLog.text:=s;
    end;
    SessionLog.SaveToFile(Logfile);
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
                      IntToHex(self.PnlLeft.width, 4)+
                      IntToHex(self.PnlAccounts.height, 4);
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
  TrayBmp:TBitmap;
Begin
  if FAccounts.Accounts.Count = 0 then exit;
  TrayHint:='';
  AccBmp:= TBitmap.create;
  TrayBmp:= TBitmap.create;
  LVAccounts.Clear;
  if Assigned(LVAccounts.SmallImages) then LVAccounts.SmallImages.Clear;
  ILTray.Clear;
  TrayBmp.LoadFromResourceName(HInstance, 'TRAY');
  ILTray.AddMasked(TrayBmp, $FF00FF); // default icon
  for i := 0 to FAccounts.Accounts.Count-1 do
  Try
    if FAccounts.Accounts.GetItem(i).Enabled then AccBmp.LoadFromResourceName(HInstance, 'ACCOUNT')
    else AccBmp.LoadFromResourceName(HInstance, 'ACCOUNTD') ;
    TrayBmp.LoadFromResourceName(HInstance, 'TRAY');
    ListItem := LVAccounts.items.add;  // prépare l'ajout
    if FAccounts.Accounts.GetItem(i).Mails.count > 0 then
    begin
      DrawTheIcon(AccBmp, FAccounts.Accounts.GetItem(i).Mails.count ,
                  FAccounts.Accounts.GetItem(i).Color  );
      DrawTheIcon(TrayBmp, FAccounts.Accounts.GetItem(i).Mails.count ,
                  FAccounts.Accounts.GetItem(i).Color  );
      ILTray.AddMasked(TrayBmp, $FF00FF);  // modified icon
      TrayHint:= TrayHint+Format('%s : %u message(s)'#10,
                             [FAccounts.Accounts.GetItem(i).Name,
                              FAccounts.Accounts.GetItem(i).Mails.count]);
    end;
    LVAccounts.SmallImages.AddMasked(AccBmp,$FF00FF);
    ListItem.ImageIndex := i;
    Listitem.Caption :=  FAccounts.Accounts.GetItem(i).Name;    // ajoute le nom
  Except
    ShowMessage(inttostr(i));
  end;
  if trayHint='' then
  begin
    TrayHint:= 'Aucun Message'
  end else
  begin
    TrayMail.BalloonHint:= TrayHint ;
  end;
  TrayMail.Hint:= TrayHint;
  LVAccounts.ItemIndex:= 0;
  //LVAccounts.SetFocus;
  AccBmp.free;
  TrayBmp.free;
  if FSettings.Settings.Notifications then TrayMail.ShowBalloonHint;

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
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  with FAccounts do
  begin
    if (TSpeedButton(Sender).Name='BtnAddAcc') then
    begin
      Account:= Default(TAccount);
      Account.UID:=0;
      Caption:= BtnAddAcc.Hint;
    end;
    if (TSpeedButton(Sender).Name='BtnEditAcc') and  (ndx>=0) then
    begin
      Account:= Accounts.GetItem(ndx);
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
      if (TSpeedButton(Sender).Name='BtnEditAcc') and  (ndx>=0) then
      Accounts.ModifyAccount(ndx, Account)
      else Accounts.AddAccount(Account);
      PopulateAccountsList ();
      if (TSpeedButton(Sender).Name='BtnEditAcc') then
      begin
        LVAccounts.ItemIndex:= ndx;
        LogAddLine(Account.UID, now, Format(AccountChanged, [Account.Name]));
      end else
      begin
        LVAccounts.ItemIndex:= LVAccounts.Items.count-1 ;
        LogAddLine(Accounts.GetItem(LVAccounts.ItemIndex).UID , now, Format(AccountAdded, [Account.Name]));
      end;

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
  if Assigned(TrayTimerBmp) then TrayTimerBmp.free;
  if Assigned(LangNums) then LangNums.free;
  if Assigned(LangFile) then LangFile.free;
  for i:=0 to length(BmpArray)-1 do
  if Assigned(BmpArray[i]) then BmpArray[i].free;
  if Assigned(ChkMailTimer) then ChkMailTimer.Destroy;
  if Assigned(SessionLog) then SessionLog.free;
end;

procedure TFMailsInBox.FormShow(Sender: TObject);
begin

end;

procedure TFMailsInBox.GetMailTimerTimer(Sender: TObject);
var
  i: integer;
  min: TDateTime;
  CurAcc: TAccount;
begin
    for i:=0 to FAccounts.Accounts.count-1 do
    begin
      // current account is enabled and interval defined
      CurAcc:= FAccounts.Accounts.GetItem(i);
      if CurAcc.Enabled and (CurAcc.Interval>0) and (now>CurAcc.NextFire) then
      begin
        min:= EncodeTime(0,CurAcc.interval,0,0);
        GetAccMail(i);
        FAccounts.Accounts.ModifyField(i, 'NEXTFIRE', now+min);
      end;
    end;
end;



procedure TFMailsInBox.IdPOP3_1Connected(Sender: TObject);
var
  Curacc: TAccount;
begin
  CurAcc:= FAccounts.Accounts.GetItem(CurAccPend);
  LStatus.Caption:= Format(ConnectedToServer, [Curacc.Name, CurAcc.Server]);
  LogAddLine(CurAcc.UID, now, LStatus.Caption );
end;

procedure TFMailsInBox.IdPOP3_1Disconnected(Sender: TObject);
var
  Curacc: TAccount;
begin
  CurAcc:= FAccounts.Accounts.GetItem(CurAccPend);
  LStatus.Caption:= Format(DisconnectedServer, [Curacc.Name, CurAcc.Server]);
  LogAddLine(CurAcc.UID, now, LStatus.Caption );
end;

procedure TFMailsInBox.IdPOP3_1Status(ASender: TObject;
  const AStatus: TIdStatus; const AStatusText: string);
begin
end;

procedure TFMailsInBox.LVAccountsDrawItem(Sender: TCustomListView;
  AItem: TListItem; ARect: TRect; AState: TOwnerDrawState);
begin

end;



procedure TFMailsInBox.LVAccountsSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  AccName: string;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx >= 0 then
  begin
    BtnGetAccMail.Enabled:= FAccounts.Accounts.GetItem(ndx).enabled;
    LVAccounts.PopupMenu:= MnuAccount;
    AccName:= FAccounts.Accounts.GetItem(ndx).Name;
    BtnAccountLog.Hint:= Format(BtnLogHint, [AccName]);
    BtnGetAccMail.Hint:= Format(BtnGetAccMailHint, [AccName]);
    BtnDelete.Hint:= Format(BtnDeleteHint, [AccName]);
    BtnEditAcc.Hint:= Format(BtnEditAccHint, [AccName]);
    UpdateInfos;
  end else LVAccounts.PopupMenu:= nil;
  PopulateMailsList(ndx);
end;

procedure TFMailsInBox.UpdateInfos;
var
  ndx: integer;
  msgs: integer;
  msgsfnd: string;
  CurAcc: TAccount;
  slastfire: string;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx >= 0 then
  begin
    CurAcc:= FAccounts.Accounts.GetItem(ndx);
    msgs:= CurAcc.Mails.Count;
    slastfire:= DateTimeToString(CurAcc.LastFire);
    RMInfos.Clear;
    if CurAcc.Enabled then RMInfos.Lines.Add(Format(AccountEnabled, [CurAcc.Name]))
    else RMInfos.Lines.Add(Format(AccountDisabled, [CurAcc.Name]));
    RMInfos.Lines.Add(Format(EmailCaption, [CurAcc.Email]));
    RMInfos.Lines.Add(Format(LastCheckCaption, [slastfire]));
    if FAccounts.Accounts.GetItem(ndx).Enabled then
    begin
      if msgs>1 then msgsfnd:= Format(MsgsFound, [msgs])
      else msgsfnd:= Format(MsgFound, [msgs]);
      RMInfos.Lines.Add(msgsfnd);
      LStatus.Caption:= Format(LStatusCaption, [msgsfnd,
                   FAccounts.Accounts.GetItem(ndx).Name, slastfire]);
      RMInfos.Lines.add(Format(NextCheckCaption, [DateTimeToString(FAccounts.Accounts.GetItem(ndx).NextFire)]));
    end else
    begin
      LStatus.Caption:= Format(AccountDisabled, [CurAcc.Name]);
    end;



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

procedure TFMailsInBox.MnuInfosClick(Sender: TObject);
var
  andx: integer;
  mndx: integer;
  mail:TMail;
begin
  andx:= LVAccounts.ItemIndex;
  if andx <0 then exit;
  mndx:= SGMails.row-1;
  mail:= FAccounts.Accounts.GetItem(andx).Mails.GetItem(mndx);
  ShowMessage(Mail.MessageFrom+#10+
              Mail.MessageTo+#10+
              Mail.MessageSubject+#10);
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
  tmpndx: integer;
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

procedure TFMailsInBox.SGMailsClick(Sender: TObject);
begin
  ShowMessage('Test');
end;

procedure TFMailsInBox.SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
var
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

procedure TFMailsInBox.SGMailsKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

end;

procedure TFMailsInBox.SGMailsMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  pf: TPoint;
  col1, row1: integer;
begin
  if Button = TMouseButton.mbRight then
  begin
    Col1:=0;
    Row1:=0;
    if SGMails.RowCount<2 then exit;
    SGMails.SetFocus;
    SGMails.MouseToCell(X, Y, Col1, Row1);
    pf := SGMails.ClientToScreen(Point(X, Y));
    if row1>0 then
    begin
      SGMails.Row:= (row1);
      // Do not use the grids PopupMenu property, it
      // prevents this event handler comletely.
      // Instead, activate the menu manually here.
      MnuMails.Popup(pf.X, pf.Y);
    end;
  end;
end;

// Check if we have changed column width

procedure TFMailsInBox.SGMailsMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  grid: TStringGrid;
begin
  if Button = mbLeft then
  begin
    grid := TStringGrid(Sender);
    if grid.GetGridState = gsColSizing then  SettingsChanged:= true;
  end;
end;

procedure TFMailsInBox.TrayTimerTimer(Sender: TObject);

begin
  if not CheckingMail then
  begin
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    if TrayTimerTick<ILtray.count-1 then inc (TrayTimerTick) else TrayTimerTick:= 0;
    // Automatic check, upon interval

  end;
end;

procedure TFMailsInBox.BtnSettingsClick(Sender: TObject);
var
  i, oldlng, oldmailsel: integer;
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
    CBNotifications.Checked:= Settings.Notifications;
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
      Settings.Notifications:= CBNotifications.Checked;
      Settings.SmallBtns:= CBSmallBtns.Checked;
      Settings.MailClient:= MailClients[CBMailClient.ItemIndex].Command;
      Settings.MailClientName:= MailClients[CBMailClient.ItemIndex].Name;
      Settings.MailClientIsUrl:= CBUrl.Checked;
      //if CBMailClient.ItemIndex <> oldmailsel then       // remove old bold style
      //MailClients[CBMailClient.ItemIndex].Tag:= false;
      Settings.SoundFile:= ESoundFile.Text;
      Settings.LangStr := LangNums.Strings[CBLangue.ItemIndex];
      if FSettings.CBLangue.ItemIndex <> oldlng then ModLangue;
      if SettingsChanged then
      begin
        PopulateAccountsList();  // Needed to change language on hints
        LogAddLine(-1, now, 'Settings changed');
      end;
    end;
  end;
end;


procedure TFMailsInBox.BtnAboutClick(Sender: TObject);

begin
  AboutBox.LastUpdate:= FSettings.Settings.LastUpdChk;
  AboutBox.ShowModal;
  // Truncate date to avoid changes if there is the same day (hh:mm are in the decimal part of the date)
  if trunc(AboutBox.LastUpdate) > trunc(FSettings.Settings.LastUpdChk) then
  begin
    FSettings.Settings.LastUpdChk:= AboutBox.LastUpdate;
    LogAddLine(-1, FSettings.Settings.LastUpdChk, LastUpdateSearch);
  end;
end;

procedure TFMailsInBox.BtnDeleteClick(Sender: TObject);
begin
  ShowMessage('Todo: delete selected account');
end;

procedure TFMailsInBox.BtnGetAllMailClick(Sender: TObject);
var
  i: integer;
  ndx: Integer;
begin
  if FAccounts.Accounts.count = 0 then exit;
  LogAddLine(-1, now, BtnGetAllMail.Hint);
  ndx:= LVAccounts.ItemIndex;   // Current selected account
  MailChecking(true);
  Application.ProcessMessages;
  for i:= 0 to FAccounts.Accounts.count-1 do
  begin
    CurAccPend:= i;
    if FAccounts.Accounts.GetItem(i).Enabled then GetPendingMail(i);
    if i=ndx then PopulateMailsList(i);
  end;
  MailChecking(false);
  PopulateAccountsList;
  LVAccounts.ItemIndex:=ndx;
  LVAccounts.SetFocus;
end;


procedure TFMailsInBox.BtnGetAccMailClick(Sender: TObject);
var
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;   // Current selected account
  GetAccMail(ndx);
end;

procedure TFMailsInBox.GetAccMail(ndx: integer);
begin
  if (ndx>=0) and not CheckingMail then
  begin
    LogAddLine(-1, now, BtnGetAccMail.Hint);
    MailChecking(true);
    CurAccPend:= ndx;
    Application.ProcessMessages;
    GetPendingMail(ndx);
    PopulateMailsList(ndx);
    MailChecking(false);
    PopulateAccountsList;
    LVAccounts.ItemIndex:=ndx;
    LVAccounts.SetFocus;
  end;
end;

// During mail checking

function TFMailsInBox.MailChecking(status: boolean): boolean;
begin
  CheckingMail:= status;
  EnableControls(not status);
  if status then
  begin
    ChkMailTimerTick:= 0;
    ChkMailTimer.StartTimer;
    Screen.Cursor:= crHourGlass;

  end else
  begin
    ChkMailTimer.StopTimer;
    TrayTimerTick:= 0;
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    Screen.Cursor:= DefCursor;
  end;

  result:= status;
end;

// retreive pop3 mail



// retreive imap mail

function TFMailsInBox.GetPendingMail(index: integer): Integer;
var
  msgs : Integer;
  idMsg: TIdMessage;
  i, siz: integer;
  CurName: string;
  mail: TMail;
  mails: TMailsList;
  min: TTime;
  HeaderOK: boolean;
  CurAcc: TAccount;
  idMsgList: TIdMessageCollection;
  AMailBoxList: TStringList;
begin
  result:= 0;
  msgs:= 0;
  mails:= TMailsList.create;
  idMsgList:= TIdMessageCollection.create;
  AMailBoxList:= TStringList.Create;
  CurAcc:= FAccounts.Accounts.GetItem(index);
  CurName:= CurAcc.Name;
  Case Curacc.Protocol of
    ptcPOP3:
      begin
        IdPOP3_1.Host:= CurAcc.Server;
        IdPOP3_1.Port:= CurAcc.Port;
        IdPOP3_1.Username:= CurAcc.UserName;
        IdPOP3_1.Password:= CurAcc.Password;
        // Authentication method
        if Curacc.SecureAuth then IdPOP3_1.AuthType:= patSASL
        else IdPOP3_1.AuthType:= patUserPass;
      end;
    ptcIMAP:
      begin
        IdIMAP4_1.Host:= CurAcc.Server;
        IdIMAP4_1.Port:= CurAcc.Port;
        IdIMAP4_1.Username:= CurAcc.UserName;
        IdIMAP4_1.Password:= CurAcc.Password;
        // Authentication method
        if Curacc.SecureAuth then IdIMAP4_1.AuthType:= iatSASL
        else IdIMAP4_1.AuthType:= iatUserPass;
      end;
  end;
  try
    LStatus.Caption:= Format(ConnectToServer, [Curacc.Name, CurAcc.Server]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
    Application.ProcessMessages;
    idMsg:= TIdMessage.Create(self);
    idMsgList:= TIdMessageCollection.Create;
    Case Curacc.Protocol of
      ptcPOP3:
        begin
          IdPop3_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(idPop3_1);
          IdPop3_1.UseTLS := TIdUseTLS(CurAcc.SSL);
          try
            IdPOP3_1.Connect;
            Application.ProcessMessages;
            msgs := IdPop3_1.CheckMessages;
           except
            on E: Exception do
            begin
              LStatus.Caption:= Format(ConnectErrorMsg, [CurAcc.Name, E.Message]);
              LogAddLine(CurAcc.UID, now, LStatus.Caption );
            end;
          end;
       end;
      ptcIMAP:
        begin
          IdIMAP4_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(IdIMAP4_1);
          IdIMAP4_1.UseTLS := TIdUseTLS(CurAcc.SSL);
          try
            if IdIMAP4_1.Connect then
            IdIMAP4_1.ListSubscribedMailBoxes(AMailBoxList);
            Application.ProcessMessages;
            if not IdIMAP4_1.SelectMailBox('Inbox') then      // select proper mailbox
            begin
              // todo select proper mailbox
            end;
            msgs:= IdIMAP4_1.MailBox.TotalMsgs;
          except
            on E: Exception do
            begin
              LStatus.Caption:= Format(ConnectErrorMsg, [CurAcc.Name, E.Message]);
              LogAddLine(CurAcc.UID, now, LStatus.Caption );
            end;
          end;
        end;
    end;
    if msgs > 0 then
    begin
      for i:= 1 to msgs do
      begin
        Mail:= Default(TMail);
        Case Curacc.Protocol of
          ptcPOP3:
            begin
              try
                HeaderOK:= IdPop3_1.RetrieveHeader(i, idMsg);
                if HeaderOK then siz:= IdPop3_1.RetrieveMsgSize(i) else exit;
              except
                on E: Exception do
                begin
                  LStatus.Caption:= Format(ConnectErrorMsg, [CurAcc.Name, E.Message]);
                  LogAddLine(CurAcc.UID, now, LStatus.Caption );
                end;
              end;
            end;
          ptcIMAP:
            begin
              try
                siz:= IdIMAP4_1.RetrieveMsgSize(i);
                IdIMAP4_1.RetrieveHeader(i, idMsg);
                siz:= siz+length(idMsg.Headers.Text) ;
              except
                on E: Exception do
                begin
                  LStatus.Caption:= Format(ConnectErrorMsg, [CurAcc.Name, E.Message]);
                  LogAddLine(CurAcc.UID, now, LStatus.Caption );
                end;
              end;

            end;
        end;
        Application.ProcessMessages;
        GetMailInfos(CurName, Mail, IdMsg, siz);
        if CurAcc.Mails.FindUIDL(Mail.MessageUIDL)>=0 then
         Mail.MessageNew:= false else Mail.MessageNew:= true;
         Mails.AddMail(Mail);
      end;

    end;
    Case Curacc.Protocol of
      ptcPOP3:
        begin
          LStatus.Caption:= Format(DisconnectServer, [Curacc.Name, CurAcc.Server]);
          LogAddLine(CurAcc.UID, now, LStatus.Caption );
          IdPop3_1.Disconnect;
        end;
      ptcIMAP:
      begin
        LStatus.Caption:= Format(DisconnectServer, [Curacc.Name, CurAcc.Server]);
        LogAddLine(CurAcc.UID, now, LStatus.Caption );
        IdIMAP4_1.Disconnect(true);
      end;
    end;
  except
    on E: Exception do
    begin
      LStatus.Caption:= Format(ConnectErrorMsg, [CurAcc.Name, E.Message]);
      LogAddLine(CurAcc.UID, now, LStatus.Caption );
    end;
  end;
  if msgs>1 then begin
    LStatus.Caption:= CurName+' : '+Format(MsgsFound, [msgs]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
  end else
  begin
    LStatus.Caption:= CurName+' : '+Format(MsgFound, [msgs]) ;
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
  end;
  // Update account checkmail dates
  FAccounts.Accounts.ModifyField(index, 'LASTFIRE', now);
  Application.ProcessMessages;
  min:= EncodeTime(0,FAccounts.Accounts.GetItem(index).interval,0,0);
  FAccounts.Accounts.ModifyField(index, 'NEXTFIRE', now+min);
  TAccount(FAccounts.Accounts.Items[index]^).Mails.Reset;
  if Mails.count > 0 then
    for i:=0 to Mails.count-1 do
    begin
     TAccount(FAccounts.Accounts.Items[index]^).Mails.AddMail(Mails.GetItem(i));
     Application.ProcessMessages;
    end;
  if assigned (Mails) then Mails.free;
  if assigned(idMsgList) then idMsgList.free;
  result:= msgs;
end;

procedure TFMailsInBox.GetMailInfos(CurName: String; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
var
  sfrom: string;
begin
  try
  sfrom:= IdMsg.From.Name;
  if length(sfrom)=0 then sfrom:= idMsg.From.Address;
  Mail.AccountName:= CurName;
  Mail.MessageFrom:= sfrom;
  Mail.FromAddress:= idMsg.From.Address;
  Mail.MessageUIDL:= idMsg.UID;
  Mail.MessageSubject:= idMsg.Subject;
  Mail.MessageTo:= idMsg.Recipients[0].Address;
  Mail.MessageDate:= idMsg.Date;
  Mail.MessageSize:= siz;
  Mail.MessageContentType:=  IdMsg.ContentType ;
  except
    on E: Exception do LStatus.Caption:= E.Message;
  end;
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

procedure TFMailsInBox.BtnAccountLogClick(Sender: TObject);
var
  Curacc:TAccount;
  csvdoc: TCSVDocument;
  i: integer;
  s: string;
begin
  csvdoc:= TCSVDocument.Create;
  if FSettings.Settings.SaveLogs then csvdoc.CSVText:= MainLog+SessionLog.text
  else csvdoc.CSVText:= SessionLog.text;
   if TBitBtn(Sender).Name= 'BtnAccountLog' then
  begin
    if LVAccounts.ItemIndex>=0 then
    begin
      CurAcc:= FAccounts.Accounts.GetItem(LVAccounts.ItemIndex);
      for i:= 0 to csvdoc.RowCount-1 do
      begin
        if csvdoc.Cells[0,i]= IntToStr(Curacc.UID) then
        s:=s+csvdoc.Cells[1,i]+' - '+csvdoc.Cells[2,i]+#10;
      end;
    end;
  end;
  if TBitBtn(Sender).Name= 'BtnLog' then
  begin
    for i:= 0 to csvdoc.RowCount-1 do
    begin
      if pos('**', csvdoc.Cells[2,i]) >0 then s:= s+csvdoc.Cells[2,i]+#10
      else s:=s+csvdoc.Cells[1,i]+' - '+csvdoc.Cells[2,i]+#10;
    end;
  end;
  Flog.Caption:= TBitBtn(Sender).Hint;
  Flog.RMLog.rtf:='';
  Flog.RMLog.Text:=s;
  Flog.RMLog.SelStart:=0;
  Flog.RMLog.Sellength:=0;
  FLog.showmodal;

  csvdoc.free;
end;

procedure TFMailsInBox.BtnLogClick(Sender: TObject);
begin

end;



procedure TFMailsInBox.BtnQuitClick(Sender: TObject);
begin
  CanClose:= true;
  Close();
end;

// Disable controls during mail check to avoid conflicts
// Display hourglass cursor

procedure TFMailsInBox.EnableControls(Enable: boolean);
begin
  LVAccounts.Enabled:= enable;
  SGMails.Enabled:= enable;
  BtnGetAccMail.Enabled:= enable;
  BtnGetAllMail.Enabled:= enable;
  BtnDelete.Enabled:= enable;
  BtnAddAcc.Enabled:= enable;
  BtnEditAcc.Enabled:= enable;
  BtnQuit.Enabled:= enable;
  if enable then SCreen.Cursor:= DefCursor
  else Screen.Cursor:= crHourGlass;
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
    else Ellipse(0,1,12,13);
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
    TextOut(6-i,1,s);
  end;
end;

// Animate tray icon

procedure TFMailsInBox.OnChkMailTimer(Sender: TObject);
begin
  if CheckingMail then
  begin;
    TrayMail.Icon.LoadFromResourceName(HINSTANCE, 'XATT'+(InttoStr(ChkMailTimerTick)));
    inc (ChkMailTimerTick);
    if ChkMailTimerTick > 5 then ChkMailTimerTick:=0;
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
    //BtnFirst.Hint:=ReadString(LangStr,'BtnFirst.Hint',BtnFirst.Hint);
    //BtnPrev.Hint:=ReadString(LangStr,'BtnPrev.Hint',BtnPrev.Hint);
    //BtnNext.Hint:=ReadString(LangStr,'BtnNext.Hint',BtnNext.Hint);
    //BtnLast.Hint:=ReadString(LangStr,'BtnLast.Hint',BtnLast.Hint);
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
    LStatusCaption:=ReadString(LangStr,'LStatusCaption','%s sur le compte %s le %');
    ConnectToServer:=ReadString(LangStr,'ConnectToServer','%s : Connexion au serveur %s');
    ConnectedToServer:=ReadString(LangStr,'ConnectedToServer','%s : Connecté au serveur %s');
    ConnectErrorMsg:=ReadString(LangStr,'ConnectErrorMsg','%s : Erreur : %s');
    DisconnectServer:=ReadString(LangStr,'DisconnectServer','%s : Déonnexion du serveur %s');
    DisconnectedServer:=ReadString(LangStr,'DisconnectedServer','%s : Déconnecté du serveur %s');
    AccountChanged:=ReadString(LangStr,'AccountChanged', 'Le compte %s a été modifié');
    AccountAdded:=ReadString(LangStr,'AccountAdded', 'Le compte %s a été ajouté');
    AccountDisabled:=ReadString(LangStr,'AccountDisabled', 'Compte %s désactivé');
    AccountEnabled:=ReadString(LangStr,'AccountEnabled', 'Compte %s activé');

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

